"""Summarize completed source-bound quality receipts without rerunning gates."""

from __future__ import annotations

import argparse
import json
import re

import checks
import run as R


def summarize(stage: str) -> dict:
    folder = R.HERE / stage
    expected = {"check-" + name + ".receipt.json" for name, _ in checks.COMMANDS}
    paths = sorted(folder.glob("check-*.receipt.json"))
    assert {path.name for path in paths} == expected
    source_sha = R.sha(folder / "source-manifest.json")
    rows = []
    for path in paths:
        value = json.loads(path.read_text(encoding="utf-8"))
        assert value["exit_code"] == 0
        assert value["execution_stage"] == stage
        assert value["source_manifest_sha256"] == source_sha
        stdout = path.with_name(path.name.replace(".receipt.json", ".stdout"))
        count = sum(int(number) for number in re.findall(
            r"test result: ok\. (\d+) passed;", stdout.read_text(encoding="utf-8")
        ))
        rows.append({"name": path.name, "receipt_sha256": R.sha(path),
                     "executed_tests": count})
    result = {"status": "pass", "stage": stage, "checks": rows,
              "executed_tests": sum(row["executed_tests"] for row in rows)}
    R.write(R.HERE / "quality-summary.json", result)
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--stage", choices=("candidate", "final"), required=True)
    args = parser.parse_args()
    result = summarize(args.stage)
    print(json.dumps({"status": result["status"], "stage": result["stage"],
                      "checks": len(result["checks"]),
                      "executed_tests": result["executed_tests"]}))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
