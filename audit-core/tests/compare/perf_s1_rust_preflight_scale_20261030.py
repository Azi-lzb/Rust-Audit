#!/usr/bin/env python3
"""Rust-only scale profile for S1 structure preflight + conditional audit.

This companion avoids the Python reference's expensive repeated prefix scan on
large synthetic inputs. Cross-engine parity remains covered by the 4k-row A/B
driver; this script measures only Rust stages at a larger scale.
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import platform
import statistics
import tempfile
from datetime import datetime, timezone
from pathlib import Path
import sys


ROOT = Path(__file__).resolve().parents[3]
AUDIT_ROOT = ROOT / "audit-core"
DEFAULT_BINARY = Path(r"F:\RustBuild\audit\release\open-data-audit.exe")
DEFAULT_REPORT = AUDIT_ROOT / "probe" / "reports" / "s1_rust_preflight_scale_20261030.json"
PROFILE_PATH = AUDIT_ROOT / "tests" / "compare" / "perf_s1_preflight_reads_20261030.py"
FIXTURE_PATH = AUDIT_ROOT / "tests" / "compare" / "perf_s1_conditional_profile_20261030.py"
HELPER_PATH = AUDIT_ROOT / "tests" / "compare" / "compare_s1_conditional_audit.py"


def load_module(name: str, path: Path):
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot load module: {path}")
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


def median(values: list[float | int | None]) -> float | None:
    present = [float(value) for value in values if value is not None]
    return round(statistics.median(present), 3) if present else None


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--binary", type=Path, default=DEFAULT_BINARY)
    parser.add_argument("--rows", type=int, default=20000)
    parser.add_argument("--rounds", type=int, default=3)
    parser.add_argument("--warmup", type=int, default=1)
    parser.add_argument("--report", type=Path, default=DEFAULT_REPORT)
    args = parser.parse_args()
    if args.rows < 1 or args.rounds < 1 or args.warmup < 0:
        parser.error("rows/rounds must be positive and warmup nonnegative")
    binary = args.binary.resolve(strict=True)
    profile = load_module("s1_rust_scale_profile", PROFILE_PATH)
    fixture_builder = load_module("s1_rust_scale_fixture", FIXTURE_PATH)
    normalize = load_module("s1_rust_scale_helper", HELPER_PATH).normalize_issue

    records = []
    with tempfile.TemporaryDirectory(prefix="audit_s1_rust_preflight_scale_") as temp:
        root = Path(temp)
        source = root / "S1-Source.xlsx"
        template = root / "S1-Template.xlsx"
        fixture = fixture_builder.generate_fixture(source, args.rows)
        template.write_bytes(source.read_bytes())
        input_profile = {
            **fixture,
            "source_bytes": source.stat().st_size,
            "template_bytes": template.stat().st_size,
        }
        before = {"source": sha256(source), "template": sha256(template)}
        all_results = []

        for index in range(args.warmup + args.rounds):
            run_dir = root / f"round-{index + 1}"
            run_dir.mkdir()
            result = profile.run_rust(
                binary,
                source,
                template,
                args.rows,
                run_dir,
                index + 1,
                normalize,
            )
            all_results.append(result)
            if index < args.warmup:
                continue
            digest = hashlib.sha256(
                "\n".join(result["issue_digests"]).encode("utf-8")
            ).hexdigest()
            records.append(
                {
                    "round": index - args.warmup + 1,
                    "preflight": result["preflight"],
                    "conditional": result["conditional"],
                    "cli_wall_ms": result["external_ms"],
                    "peak_working_set_mb": result["peak_rss_mb"],
                    "issue_count": result["issue_count"],
                    "issue_digest_sha256": digest,
                }
            )
            print(
                f"round {len(records)}: cli={result['external_ms']}ms "
                f"read={result['preflight']['source_read_ms']}ms "
                f"issues={result['issue_count']}",
                flush=True,
            )

        after = {"source": sha256(source), "template": sha256(template)}

    if before != after:
        raise RuntimeError("synthetic source or template changed during profiling")
    expected_hashes = {record["issue_digest_sha256"] for record in records}
    if len(expected_hashes) != 1:
        raise RuntimeError("Rust issue output changed between repeated runs")
    if any(not result["preflight"]["matched"] for result in all_results):
        raise RuntimeError("Rust structure preflight did not match the generated template")

    report = {
        "schema_version": 1,
        "date": datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds"),
        "feature": "S1-F1",
        "scenario": "synthetic-rust-only-preflight-conditional-scale",
        "scope": "Rust-only; synthetic input generated from the existing S1 structure/conditional profile fixture. Cross-engine parity is not claimed at this scale.",
        "platform": platform.platform(),
        "rows": args.rows,
        "input_profile": input_profile,
        "binary": str(binary),
        "binary_sha256": sha256(binary),
        "input_sha256_before": before,
        "input_sha256_after": after,
        "inputs_unchanged": before == after,
        "warmup": args.warmup,
        "rounds": records,
        "medians": {
            "source_read_ms": median([row["preflight"]["source_read_ms"] for row in records]),
            "template_preflight_read_ms": median([row["preflight"]["template_read_ms"] for row in records]),
            "preflight_elapsed_ms": median([row["preflight"]["elapsed_ms"] for row in records]),
            "conditional_evaluation_ms": median([row["conditional"]["evaluation_ms"] for row in records]),
            "conditional_issue_build_ms": median([row["conditional"]["issue_build_ms"] for row in records]),
            "conditional_command_ms": median([row["conditional"]["command_elapsed_ms"] for row in records]),
            "cli_wall_ms": median([row["cli_wall_ms"] for row in records]),
            "peak_working_set_mb": median([row["peak_working_set_mb"] for row in records]),
            "issue_count": median([row["issue_count"] for row in records]),
        },
        "repeatability": {
            "issue_digest_sha256_matches_all_rounds": len(expected_hashes) == 1,
            "structure_match_all_rounds": all(result["preflight"]["matched"] for result in all_results),
        },
    }
    args.report.parent.mkdir(parents=True, exist_ok=True)
    args.report.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"status": "passed", "report": str(args.report), "medians": report["medians"]}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
