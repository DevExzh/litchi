"""Run the frozen owner/workspace quality commands serially with source custody.

An attempt is exclusive and never resumes implicitly. Failed attempts remain
evidence; a later attempt must use a new label. Root chooses the final passing
attempt only after deciding whether the production candidate is retained.
"""

import argparse
import json
from pathlib import Path
import subprocess
import sys

import check_attempt as CHECK
import run as R


def read(path):
    return json.loads(path.read_text())


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source_stage", choices=("candidate", "final"))
    parser.add_argument("label")
    args = parser.parse_args()
    assert args.label and all(c.isalnum() or c in "-_" for c in args.label)
    stage_manifest_path = R.HERE / args.source_stage / "source-manifest.json"
    manifest = read(stage_manifest_path)
    binding = read(R.HERE / "workspace-lock.json")
    expected = dict(manifest)
    expected["Cargo.lock"] = binding["sha256"]
    assert CHECK.source_manifest() == expected, "quality source inventory differs"
    supplemental = read(R.HERE / "supplemental-inputs.json")
    for name, digest in supplemental["files"].items():
        assert R.sha(R.REPO / name) == digest, name
    plan_path = R.HERE / "quality-plan.json"
    commands = read(plan_path)["commands"]
    assert len(commands) == 11
    folder = R.HERE / "quality-attempts" / args.label
    folder.mkdir(parents=True, exist_ok=False)
    scripts = {path.name: R.sha(path) for path in (
        Path(__file__), R.HERE / "check_attempt.py", R.HERE / "run.py")}
    inputs = {
        "schema": "xlsx_0553_quality_inputs_v1",
        "created_utc": R.now(), "source_stage": args.source_stage,
        "source_manifest_sha256": R.sha(stage_manifest_path),
        "workspace_lock_sha256": binding["sha256"],
        "quality_plan_sha256": R.sha(plan_path),
        "supplemental_inputs_sha256": R.sha(R.HERE / "supplemental-inputs.json"),
        "scripts": scripts, "commands": commands,
    }
    R.write(folder / "inputs.json", inputs)
    rows = []
    for name, command in commands.items():
        assert CHECK.source_manifest() == expected, "source changed between checks"
        assert R.sha(plan_path) == inputs["quality_plan_sha256"]
        for script, digest in scripts.items():
            assert R.sha(R.HERE / script) == digest, script
        attempt_name = args.label + "-" + name
        result = subprocess.run(
            [sys.executable, "-B", str(R.HERE / "check_attempt.py"),
             attempt_name, "--", *command], cwd=R.REPO)
        attempt = R.HERE / "check-attempts" / attempt_name
        receipt_path = attempt / "receipt.json"
        receipt = read(receipt_path)
        assert receipt["command"] == command
        assert read(attempt / "source-manifest.json") == expected
        assert receipt["script_sha256"] == scripts["check_attempt.py"]
        assert receipt["run_sha256"] == scripts["run.py"]
        stable = receipt["source_stable"] and CHECK.source_manifest() == expected
        rows.append({
            "name": name, "command": command,
            "attempt": str(attempt.relative_to(R.REPO)),
            "receipt_sha256": R.sha(receipt_path),
            "exit_code": receipt["exit_code"],
            "runner_exit_code": result.returncode,
            "source_stable": stable,
        })
        if result.returncode or receipt["exit_code"] or not stable:
            break
    passed = len(rows) == len(commands) and all(
        row["exit_code"] == 0 and row["runner_exit_code"] == 0
        and row["source_stable"] for row in rows)
    for script, digest in scripts.items():
        assert R.sha(R.HERE / script) == digest, script
    R.write(folder / "result.json", {
        "schema": "xlsx_0553_quality_v1",
        "status": "pass" if passed else "failed",
        "source_stage": args.source_stage,
        "source_manifest_sha256": inputs["source_manifest_sha256"],
        "workspace_lock_sha256": inputs["workspace_lock_sha256"],
        "quality_plan_sha256": inputs["quality_plan_sha256"],
        "inputs_path": str((folder / "inputs.json").relative_to(R.REPO)),
        "inputs_sha256": R.sha(folder / "inputs.json"),
        "completed_utc": R.now(), "commands": rows,
    })
    raise SystemExit(0 if passed else 1)


if __name__ == "__main__":
    main()
