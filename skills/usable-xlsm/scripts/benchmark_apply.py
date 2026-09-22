"""Measure the real Excel apply path with reproducible disposable copies."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
import shutil
import statistics
import subprocess
import time


def _run_one(root: Path, source: Path, script: Path, target: Path, mode: str) -> dict[str, object]:
    shutil.copy2(source, target)
    command = [
        "uv",
        "run",
        "--project",
        "skills/usable-xlsm",
        "usable-xlsm",
        "apply",
        "--script",
        str(script),
        "--workbook",
        str(target),
        "--trust-workbook",
    ]
    if mode == "dev":
        command.append("--no-test")
    elif mode == "shared":
        command.append("--shared-test-copy")
    started = time.perf_counter()
    completed = subprocess.run(command, cwd=root, capture_output=True, text=True, encoding="utf-8")
    wall = time.perf_counter() - started
    if completed.returncode:
        raise RuntimeError(
            f"apply failed for {target} (exit {completed.returncode})\n"
            f"stdout={completed.stdout}\nstderr={completed.stderr}"
        )
    report = json.loads(completed.stdout)
    return {
        "target": str(target),
        "wall_seconds": round(wall, 3),
        "apply_duration": round(float(report["duration"]), 3),
        "excel_jobs": report["excel_jobs"],
        "test_mode": report["test_mode"],
        "ok": report["ok"],
    }


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--workbook", type=Path, required=True)
    parser.add_argument("--script", type=Path, required=True)
    parser.add_argument("--output-dir", type=Path, required=True)
    parser.add_argument("--mode", choices=("dev", "isolated", "shared"), default="dev")
    parser.add_argument("--runs", type=int, default=3)
    args = parser.parse_args()
    if args.runs < 1:
        parser.error("--runs must be positive")

    root = Path(__file__).resolve().parents[3]
    source = (root / args.workbook).resolve()
    script = (root / args.script).resolve()
    output_dir = (root / args.output_dir).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    runs = [
        _run_one(root, source, script, output_dir / f"apply-{args.mode}-{index}.xlsm", args.mode)
        for index in range(1, args.runs + 1)
    ]
    walls = [float(run["wall_seconds"]) for run in runs]
    durations = [float(run["apply_duration"]) for run in runs]
    print(
        json.dumps(
            {
                "mode": args.mode,
                "runs": runs,
                "p50_wall_seconds": round(statistics.median(walls), 3),
                "p50_apply_duration": round(statistics.median(durations), 3),
            },
            ensure_ascii=False,
            indent=2,
        )
    )


if __name__ == "__main__":
    main()
