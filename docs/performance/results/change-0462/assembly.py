#!/usr/bin/env python3
"""Retain release disassembly for the attribute-cache mechanism review.

The 0462 candidate introduces a shape-specific known-key path.  A symbol can
be absent because the relevant implementation is still in the baseline, or
because the compiler inlined it into ``Parser::shape_builder``.  Both cases
are recorded explicitly; raw ``.asm.txt`` output is retained for every
command, including empty/missing output.
The verifier authenticates the resulting whitelist and the derived
stack/call observations.
"""

from __future__ import annotations

import hashlib
import json
from pathlib import Path
import re
import subprocess
import sys


ROOT = Path(__file__).resolve().parent
VARIANTS = {"baseline", "candidate"}
TARGETS = (
    ("get", "<litchi_odp::codec::parser::codec::xml::validation::ElementAttrs>::get"),
    ("lookup", "<litchi_odp::codec::parser::codec::xml::validation::ElementAttrs>::lookup"),
    ("get-known", "<litchi_odp::codec::parser::codec::xml::validation::ShapeAttrs>::get_known"),
    ("shape-builder", "<litchi_odp::codec::parser::codec::Parser>::shape_builder"),
)
INSTRUCTION_RE = re.compile(r"^\s*[0-9a-f]+:\s+\S")
STACK_RE = re.compile(r"\bsub\s+\$0x([0-9a-f]+),%rsp\b")
CALL_RE = re.compile(r"\bcall(?:q)?\s+[^<\n]*<(.+)>")
HEADER_RE = re.compile(r"^\s*[0-9a-f]+\s+<.*>:\s*$")


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def disassemble(binary: Path, symbol: str) -> tuple[list[str], bytes]:
    argv = ["objdump", "-d", "--demangle", "--no-show-raw-insn", "--disassemble=" + symbol, str(binary)]
    return argv, subprocess.check_output(argv)


def body_present(raw: bytes, symbol: str) -> bool:
    return observations(raw, symbol)["status"] == "present"


def observations(raw: bytes, symbol: str) -> dict[str, object]:
    text = raw.decode("utf-8", errors="replace")
    lines = text.splitlines()
    marker = symbol + ">:"
    try:
        start = next(index for index, line in enumerate(lines) if marker in line)
    except StopIteration:
        return {"status": "missing_or_inlined", "instruction_count": 0, "stack_frame_bytes": None, "calls": []}
    body: list[str] = []
    for line in lines[start + 1 :]:
        if HEADER_RE.match(line):
            break
        body.append(line)
    instructions = [line for line in body if INSTRUCTION_RE.search(line)]
    frames = [int(match.group(1), 16) for line in body if (match := STACK_RE.search(line))]
    calls = sorted({match.group(1) for line in body if (match := CALL_RE.search(line))})
    return {
        "status": "present" if instructions else "missing_or_inlined",
        "instruction_count": len(instructions),
        "stack_frame_bytes": max(frames, default=0),
        "calls": calls,
    }


def generated_code_evidence(targets: list[dict[str, object]]) -> dict[str, object]:
    by_id = {str(target["id"]): target for target in targets}
    lookup = by_id["lookup"]
    known = by_id["get-known"]
    builder = by_id["shape-builder"]

    def calls_old_lookup(target: dict[str, object]) -> bool:
        calls = target.get("calls")
        return isinstance(calls, list) and any(
            isinstance(call, str) and "ElementAttrs" in call and "lookup" in call
            for call in calls
        )

    builder_calls_lookup = calls_old_lookup(builder)
    known_calls_lookup = calls_old_lookup(known)
    return {
        "known_cached_scan": {
            "old_lookup_symbol": lookup["symbol"],
            "shape_builder_calls_old_lookup": builder_calls_lookup,
            "known_path_calls_old_lookup": known_calls_lookup,
            "eliminated": builder["status"] == "present" and not builder_calls_lookup,
        },
        "stack_cost": {
            "shape_builder_bytes": builder["stack_frame_bytes"],
            "shape_attrs_get_known_bytes": known["stack_frame_bytes"],
        },
    }


def main() -> int:
    if len(sys.argv) != 2 or sys.argv[1] not in VARIANTS:
        raise SystemExit("usage: assembly.py baseline|candidate")
    variant = sys.argv[1]
    binding_path = ROOT / (variant + "-binding.json")
    binding = json.loads(binding_path.read_text(encoding="utf-8"))
    binary_row = binding["binaries"]["normal"]
    binary = Path(binary_row["path"])
    assert sha(binary) == binary_row["sha256"]

    targets: list[dict[str, object]] = []
    artifacts: list[dict[str, object]] = []
    expected_artifact_names = {
        f"{known_variant}-{label}.asm.txt"
        for known_variant in VARIANTS
        for label, _ in TARGETS
    }
    existing_asm = {
        path.relative_to(ROOT).as_posix()
        for path in ROOT.rglob("*.asm.txt")
        if path.is_file() and not path.is_symlink()
    }
    if existing_asm - expected_artifact_names:
        raise AssertionError(f"unexpected {variant} assembly artifacts: {sorted(existing_asm - expected_artifact_names)}")

    for label, symbol in TARGETS:
        argv, raw = disassemble(binary, symbol)
        observation = observations(raw, symbol)
        path = ROOT / f"{variant}-{label}.asm.txt"
        if path.is_symlink():
            raise AssertionError(f"{path.name}: symlinked assembly artifact")
        if path.exists():
            assert path.read_bytes() == raw
        else:
            path.write_bytes(raw)
        artifact = {"path": path.name, "bytes": len(raw), "sha256": sha(path)}
        artifacts.append(artifact)
        targets.append({
            "id": label,
            "symbol": symbol,
            "command": argv,
            "status": observation.pop("status"),
            "artifact": artifact,
            **observation,
        })

    row = {
        "change": 462,
        "variant": variant,
        "driver_sha256": sha(Path(__file__).resolve()),
        "binding_sha256": sha(binding_path),
        "binary_sha256": sha(binary),
        "targets": targets,
        "commands": [target["command"] for target in targets],
        "artifacts": artifacts,
        "generated_code_evidence": generated_code_evidence(targets),
    }
    output = ROOT / (variant + "-assembly.json")
    with output.open("x", encoding="utf-8") as stream:
        json.dump(row, stream, indent=2)
        stream.write("\n")
    print(variant, "assembly bound")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
