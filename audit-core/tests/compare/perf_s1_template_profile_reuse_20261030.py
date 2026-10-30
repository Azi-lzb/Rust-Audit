#!/usr/bin/env python3
"""Run the release-profile S1 template-profile reuse benchmark and save JSON."""

from __future__ import annotations

import json
import hashlib
import os
import platform
import subprocess
from datetime import datetime, timezone
from pathlib import Path


ROOT = Path(__file__).resolve().parents[3]
DEFAULT_REPORT = ROOT / "audit-core" / "probe" / "reports" / "s1_template_profile_reuse_20261030.json"
BENCHMARK_MARKER = "S1_TEMPLATE_PROFILE_REUSE_BENCHMARK="


def main() -> int:
    import argparse

    parser = argparse.ArgumentParser()
    parser.add_argument("--report", type=Path, default=DEFAULT_REPORT)
    args = parser.parse_args()

    environment = os.environ.copy()
    environment.setdefault("CARGO_TARGET_DIR", r"F:\RustBuild\audit")
    command = [
        "cargo",
        "test",
        "--locked",
        "--release",
        "-p",
        "audit-core",
        "--test",
        "s1_template_profile_reuse",
        "--",
        "--nocapture",
    ]
    result = subprocess.run(command, cwd=ROOT, env=environment, text=True, capture_output=True)
    if result.returncode != 0:
        print(result.stdout)
        print(result.stderr)
        return result.returncode

    benchmark_lines = [
        line.split(BENCHMARK_MARKER, 1)[1]
        for line in (result.stdout + "\n" + result.stderr).splitlines()
        if BENCHMARK_MARKER in line
    ]
    if len(benchmark_lines) != 1:
        raise RuntimeError(f"expected one benchmark JSON line, found {len(benchmark_lines)}")
    measurement = json.loads(benchmark_lines[0])
    source_commit = subprocess.run(
        ["git", "rev-parse", "--short", "HEAD"],
        cwd=ROOT,
        check=True,
        text=True,
        capture_output=True,
    ).stdout.strip()
    source_worktree_dirty = bool(
        subprocess.run(
            ["git", "status", "--porcelain"],
            cwd=ROOT,
            check=True,
            text=True,
            capture_output=True,
        ).stdout.strip()
    )
    release_binary = Path(environment["CARGO_TARGET_DIR"]) / "release" / "open-data-audit.exe"
    release_sha256 = (
        hashlib.sha256(release_binary.read_bytes()).hexdigest()
        if release_binary.is_file()
        else None
    )
    report = {
        "schema_version": 1,
        "date": datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds"),
        "feature": "S1-F1/S1-F2",
        "scenario": "template-profile-reuse-substep",
        "platform": platform.platform(),
        "source_commit": source_commit,
        "source_worktree_dirty": source_worktree_dirty,
        "release_binary_sha256": release_sha256,
        "scope": "Release-profile function-level benchmark; measures sheet-directory recommendation only, not full GUI audit or S1-F2 export.",
        "timer_scope": "Five paired in-process medians after one warmup. Legacy path reloads template directory and all template sheet names once per source. Reuse path loads the same profiles once per batch; both paths read every source sheet directory once.",
        "measurement": measurement,
        "correctness": {
            "recommendations_identical_for_every_source": measurement[
                "all_21_recommendations_identical"
            ],
            "inputs_written_once_then_reused_for_both_paths": True,
            "output_workbooks_created": False,
            "peak_memory_measured": False,
        },
        "test_command": "cargo test --locked --release -p audit-core --test s1_template_profile_reuse -- --nocapture",
        "notes": [
            "The synthetic XLSX files contain sheet directories only and no cell values; this isolates repeated template metadata loading.",
            "Reported speedup is limited to this matching substep and workload. The absolute saving is small relative to Excel recalculation and does not represent a full-feature speedup.",
        ],
    }
    args.report.parent.mkdir(parents=True, exist_ok=True)
    args.report.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
