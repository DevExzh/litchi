"""Offline stack-depth and loss diagnostics for retained 0829 perf frames.

An empty stack or an absent truncation marker is evidence about the retained
text only. It cannot establish complete unwinding or assign causal cost.
"""

from __future__ import annotations

from collections import Counter
import gzip
import hashlib
import json
from pathlib import Path
import re
import sys


P = Path(__file__).resolve().parent
HEADER = re.compile(
    r"^\s*\S+\s+\d+\s+(?:\[\d+\]\s+)?\d+(?:\.\d+)?:\s+"
    r"(\d+)\s+cycles:u:\s*$"
)
FRAME = re.compile(r"^\s*[0-9a-fA-F]+\s+(.+?)\s+\(([^)]*)\)\s*$")


def read(path: Path):
    return json.loads(path.read_text(encoding="utf-8"))


def checked(value, label: str) -> bytes:
    assert isinstance(value, dict) and isinstance(value.get("path"), str), label
    data = Path(value["path"]).read_bytes()
    assert len(data) == value["bytes"] and hashlib.sha256(data).hexdigest() == value["sha256"], label
    return data


def parse(data: bytes):
    text = data.decode("utf-8", errors="replace")
    lines = text.splitlines()
    stacks = []
    current = None
    malformed = []
    for line in lines:
        match = HEADER.fullmatch(line)
        if match:
            if current is not None:
                stacks.append(current)
            current = {"period": int(match[1]), "frames": [], "lines": [line]}
            continue
        if current is None:
            continue
        if not line.strip():
            stacks.append(current)
            current = None
            continue
        current["lines"].append(line)
        if FRAME.fullmatch(line) is None:
            malformed.append(line)
        else:
            current["frames"].append(line)
    if current is not None:
        stacks.append(current)
    truncation = [line for line in lines if re.search(r"truncat|stack depth|callchain", line, re.I)]
    lost = [line for line in lines if "PERF_RECORD_LOST" in line
            or ("lost" in line.lower() and "sample" in line.lower())]
    sample_lines = {line for stack in stacks for line in stack["lines"]}
    status = [line for line in lines if line.strip() and HEADER.fullmatch(line) is None
              and line not in sample_lines]
    return stacks, {"malformed": malformed, "truncation": truncation,
                    "lost": lost, "status": status, "raw_lines": len(lines),
                    "raw_bytes": len(data)}


def main() -> None:
    assert sys.argv[1:] in (["--write"], ["--check"])
    complete = read(P / "perf/decode-complete.json")
    result = {"schema": "litchi.performance.0829.stack-diagnostics.v1",
              "status": complete.get("status"), "repeats": [],
              "complete_unwinding_proven": False,
              "limitation": "Stack depth, empty stacks, and explicit markers are diagnostics only; missing callers and unmarked truncation cannot be ruled out."}
    if complete.get("status") == "available":
        for member in read(P / "perf/compression.json"):
            if member.get("kind") != "frames":
                continue
            stored = checked(member["compressed"], f"compressed frames {member['repeat']}")
            raw = gzip.decompress(stored)
            assert len(raw) == member["original"]["bytes"]
            assert hashlib.sha256(raw).hexdigest() == member["original"]["sha256"]
            stacks, diagnostics = parse(raw)
            depths = Counter(len(stack["frames"]) for stack in stacks)
            result["repeats"].append({
                "repeat": member["repeat"],
                "frames_sha256": member["original"]["sha256"],
                "samples": len(stacks),
                "max_observed_stack_depth": max(depths) if depths else 0,
                "stack_depth_counts": dict(sorted(depths.items())),
                "empty_stack_samples": depths.get(0, 0),
                "explicit_truncation_lines": diagnostics["truncation"],
                "explicit_truncation_count": len(diagnostics["truncation"]),
                "lost_event_lines": diagnostics["lost"],
                "lost_event_count": len(diagnostics["lost"]),
                "status_lines": diagnostics["status"],
                "status_line_count": len(diagnostics["status"]),
                "malformed_frame_lines": diagnostics["malformed"],
                "malformed_frame_count": len(diagnostics["malformed"]),
                "stacks_ending_unknown": sum(
                    bool(stack["frames"] and
                         ("[unknown]" in stack["frames"][-1].lower()
                          or "??" in stack["frames"][-1])) for stack in stacks),
                "complete_unwinding_proven": False,
                "limitation": result["limitation"],
                "raw_text_lines": diagnostics["raw_lines"],
                "raw_text_bytes": diagnostics["raw_bytes"],
            })
        result["repeats"].sort(key=lambda row: row["repeat"])
    encoded = json.dumps(result, indent=2, sort_keys=True) + "\n"
    out = P / "stack-diagnostics.json"
    if sys.argv[1] == "--write":
        assert not out.exists(), "refusing to overwrite stack diagnostics"
        out.write_text(encoded, encoding="utf-8")
    else:
        assert out.is_file() and out.read_text(encoding="utf-8") == encoded
    print("0829 stack diagnostics PASS")


if __name__ == "__main__":
    main()
