#!/usr/bin/env python3
"""Same-input A/B profile for the S1 source structure reader projection."""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import os
import shutil
import statistics
import subprocess
import sys
import tempfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


WORKSPACE = Path(__file__).resolve().parents[3]
AUDIT_ROOT = WORKSPACE / "audit-core"
FIXTURE_DRIVER = AUDIT_ROOT / "tests" / "compare" / "perf_s1_conditional_profile_20261030.py"
DEFAULT_REPORT = AUDIT_ROOT / "probe" / "reports" / "s1_structure_source_projection_20261030.json"
TARGET_DIR = Path(r"F:\RustBuild\audit")
TEST_NAME = "s1_structure_projection_source_perf"
TEST_FILTER = "profile_source_structure_projection_same_input"


def load_module(name: str, path: Path):
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot load fixture helper {path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def git_revision() -> str:
    result = subprocess.run(
        ["git", "rev-parse", "--short", "HEAD"],
        cwd=WORKSPACE,
        check=True,
        capture_output=True,
        text=True,
    )
    return result.stdout.strip()


def median(rows: list[dict[str, Any]], key: str) -> float:
    return round(statistics.median(float(row[key]) / 1000 for row in rows), 3)


def run_profile(rows: int, rounds: int, warmup: int, report_path: Path) -> dict[str, Any]:
    fixture_module = load_module("s1_structure_projection_fixture", FIXTURE_DRIVER)
    with tempfile.TemporaryDirectory(prefix="audit-s1-structure-source-") as temp_name:
        temp_dir = Path(temp_name)
        source = temp_dir / "source.xlsx"
        template = temp_dir / "template.xlsx"
        fixture_shape = fixture_module.generate_fixture(source, rows)
        shutil.copy2(source, template)
        source_hash = sha256(source)
        template_hash = sha256(template)

        environment = os.environ.copy()
        environment["CARGO_TARGET_DIR"] = str(TARGET_DIR)
        environment["CARGO_BUILD_JOBS"] = "1"
        environment["AUDIT_S1_STRUCTURE_PROFILE_SOURCE"] = str(source)
        environment["AUDIT_S1_STRUCTURE_PROFILE_TEMPLATE"] = str(template)
        environment["AUDIT_S1_STRUCTURE_PROFILE_ROUNDS"] = str(rounds)
        environment["AUDIT_S1_STRUCTURE_PROFILE_WARMUP"] = str(warmup)
        command = [
            "cargo",
            "test",
            "--locked",
            "--release",
            "-p",
            "audit-core",
            "--test",
            TEST_NAME,
            TEST_FILTER,
            "--",
            "--ignored",
            "--nocapture",
            "--test-threads=1",
        ]
        result = subprocess.run(
            command,
            cwd=WORKSPACE,
            env=environment,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
        output = result.stdout + "\n" + result.stderr
        if result.returncode != 0:
            raise RuntimeError(
                "focused release profile failed\n"
                f"command: {' '.join(command)}\n"
                f"stdout:\n{result.stdout[-6000:]}\n"
                f"stderr:\n{result.stderr[-6000:]}"
            )

        marker = "S1_STRUCTURE_SOURCE_PROFILE="
        payloads = [line.split(marker, 1)[1] for line in output.splitlines() if marker in line]
        if len(payloads) != 1:
            raise RuntimeError(f"expected one profile marker, found {len(payloads)}")
        profile = json.loads(payloads[0])
        if profile["source_sha256_before"] != source_hash or profile["source_sha256_after"] != source_hash:
            raise RuntimeError("source file hash changed during the profile")
        if profile["template_sha256_before"] != template_hash or profile["template_sha256_after"] != template_hash:
            raise RuntimeError("template file hash changed during the profile")

    rounds_data = profile["rounds"]
    full_median_ms = median(rounds_data, "full_total_us")
    projection_median_ms = median(rounds_data, "projection_total_us")
    report = {
        "schema_version": 1,
        "generated_at_utc": datetime.now(timezone.utc).isoformat(),
        "git_revision": git_revision(),
        "platform": sys.platform,
        "python": sys.version.split()[0],
        "target_dir": str(TARGET_DIR),
        "test_target": TEST_NAME,
        "fixture": {
            **fixture_shape,
            "source_bytes": profile["source_bytes"],
            "template_bytes": profile["template_bytes"],
            "source_sha256": source_hash,
            "template_sha256": template_hash,
            "inputs_byte_identical": source_hash == template_hash,
        },
        "method": {
            "baseline": "excel-reader full WorkbookDump read + existing structure comparison",
            "candidate": "Calamine structure projection + existing structure comparison",
            "same_file_and_process": True,
            "alternating_order": True,
            "rounds": rounds,
            "warmup": warmup,
            "correctness_checked_each_round": True,
            "source_and_template_hashes_unchanged": True,
        },
        "summary": {
            "full_read_plus_compare_median_ms": full_median_ms,
            "projection_path_median_ms": projection_median_ms,
            "speedup": round(full_median_ms / projection_median_ms, 3) if projection_median_ms else None,
            "full_read_median_ms": median(rounds_data, "full_read_us"),
            "full_compare_median_ms": median(rounds_data, "full_compare_us"),
            "all_matches_identical": all(row["matched"] for row in rounds_data),
        },
        "profile": profile,
        "cargo_command": command,
    }
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return report


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--rows", type=int, default=20000)
    parser.add_argument("--rounds", type=int, default=5)
    parser.add_argument("--warmup", type=int, default=1)
    parser.add_argument("--report", type=Path, default=DEFAULT_REPORT)
    args = parser.parse_args()
    if args.rows < 1 or args.rounds < 1 or args.warmup < 0:
        parser.error("rows/rounds must be positive and warmup must be nonnegative")
    print(json.dumps(run_profile(args.rows, args.rounds, args.warmup, args.report), ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
