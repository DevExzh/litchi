"""Offline summary of the retained 0806 production-crate quality gate."""

from __future__ import annotations

import json
import re
import sys
from pathlib import Path


P = Path(__file__).resolve().parent
PATTERN = re.compile(
    r"test result: (\w+)\.\s+(\d+) passed;\s+(\d+) failed;\s+"
    r"(\d+) ignored;\s+(\d+) measured;\s+(\d+) filtered out"
)


def log_path(value: dict) -> Path:
    path = Path(value["path"])
    if not path.is_file():
        path = P / path.name
    assert path.is_file() and not path.is_symlink()
    assert path.stat().st_size == value["bytes"]
    return path


def summarize() -> dict:
    quality = json.loads((P / "quality.json").read_text(encoding="utf-8"))
    rows = []
    for row in quality["rows"]:
        command = row["command"]
        if "test" not in command:
            continue
        text = log_path(row["log"]).read_text(encoding="utf-8", errors="replace")
        matches = list(PATTERN.finditer(text))
        assert matches
        for match in matches:
            status, *counts = match.groups()
            rows.append(dict(zip(
                ("status", "passed", "failed", "ignored", "measured", "filtered"),
                (status, *map(int, counts)),
            )))
    assert rows and all(row["status"] == "ok" and row["failed"] == 0 for row in rows)
    return {
        "schema": "litchi.performance.0806.quality-summary.v1",
        "gates": 6,
        "suites": len(rows),
        "passed": sum(row["passed"] for row in rows),
        "failed": sum(row["failed"] for row in rows),
        "ignored": sum(row["ignored"] for row in rows),
        "results": rows,
    }


if __name__ == "__main__":
    result = summarize()
    output = P / "quality-summary.json"
    if "--check" in sys.argv:
        assert json.loads(output.read_text(encoding="utf-8")) == result
    else:
        assert not output.exists()
        output.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n",
                          encoding="utf-8")
    print(json.dumps({key: value for key, value in result.items() if key != "results"},
                     sort_keys=True))
