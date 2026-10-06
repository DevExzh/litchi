#!/usr/bin/env python3
"""Decode one Heaptrack 1.5 interpreted profile without filtering its scope.

The decoder reports whole-process allocation events, requested bytes, peak live
heap, and live bytes at EOF. It intentionally does not infer timed-operation
or native-latency metrics.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import subprocess
from pathlib import Path
from typing import Iterator


EXPECTED_VERSION = "10500 3"


class ParseError(RuntimeError):
    """The interpreted profile is malformed or unsupported."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise ParseError(message)


def hex_integer(value: str, *, line_number: int, field: str) -> int:
    """Parse a Heaptrack hexadecimal field without permitting Python indexing."""

    try:
        parsed = int(value, 16)
    except ValueError as error:
        raise ParseError(f"line {line_number}: invalid hexadecimal {field} {value!r}") from error
    require(parsed >= 0, f"line {line_number}: negative {field} is invalid")
    return parsed


def descriptor(path: Path) -> dict[str, object]:
    require(path.is_file(), f"missing profile: {path}")
    digest = hashlib.sha256()
    size = 0
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1 << 20), b""):
            size += len(chunk)
            digest.update(chunk)
    return {"path": str(path.resolve()), "bytes": size, "sha256": digest.hexdigest()}


def lines(path: Path) -> Iterator[str]:
    process = subprocess.Popen(
        ["zstdcat", str(path)],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding="utf-8",
    )
    try:
        require(process.stdout is not None, "zstdcat stdout is unavailable")
        assert process.stdout is not None
        for line in process.stdout:
            yield line.rstrip("\n")
        stderr = process.stderr.read() if process.stderr is not None else ""
        status = process.wait()
        require(status == 0, f"zstdcat failed ({status}): {stderr.strip()}")
    finally:
        # Parsing can fail before the compressed stream reaches EOF.  Close
        # the pipes and reap zstdcat on every path so a malformed profile
        # cannot leave a decompressor running in the capture process.
        if process.poll() is None:
            process.kill()
        process.wait()
        process.stdout.close()
        if process.stderr is not None:
            process.stderr.close()


def parse(path: Path) -> dict[str, object]:
    strings = [""]
    instructions: list[tuple[str, ...]] = [()]
    traces: list[frozenset[str]] = [frozenset()]
    allocations: list[tuple[int, int]] = []
    active_allocations: list[int] = []
    command = None
    version_seen = False
    summary_comments: set[str] = set()
    allocation_events = 0
    deallocation_events = 0
    requested_bytes = 0
    live_bytes = 0
    peak_live_bytes = 0

    stream = lines(path)
    try:
        for line_number, line in enumerate(stream, start=1):
            tag, separator, rest = line.partition(" ")
            require(separator or tag in {"", "R"}, f"line {line_number}: malformed record")
            if summary_comments == {"strings", "ips"} and tag not in {"", "#"}:
                raise ParseError(f"line {line_number}: record follows terminal summaries")
            if tag == "v":
                require(not version_seen and rest == EXPECTED_VERSION,
                        f"line {line_number}: unsupported or repeated format version {rest!r}")
                version_seen = True
            elif tag == "X":
                require(command is None, f"line {line_number}: repeated command record")
                command = rest
            elif tag == "s":
                fields = rest.split(" ", 1)
                require(len(fields) == 2, f"line {line_number}: malformed string record")
                length = hex_integer(fields[0], line_number=line_number, field="string length")
                value = fields[1]
                require(len(value.encode("utf-8")) == length,
                        f"line {line_number}: string length mismatch")
                strings.append(value)
            elif tag == "i":
                fields = rest.split()
                require(len(fields) >= 2,
                        f"line {line_number}: malformed instruction record")
                names = []
                for index in range(2, len(fields), 3):
                    string_index = hex_integer(fields[index], line_number=line_number,
                                               field="instruction string reference")
                    require(string_index < len(strings),
                            f"line {line_number}: instruction string reference out of range")
                    names.append(strings[string_index])
                instructions.append(tuple(names))
            elif tag == "t":
                fields = rest.split()
                require(len(fields) == 2, f"line {line_number}: malformed trace record")
                instruction_index = hex_integer(fields[0], line_number=line_number,
                                                field="trace instruction reference")
                parent_index = hex_integer(fields[1], line_number=line_number,
                                           field="trace parent reference")
                require(instruction_index < len(instructions) and parent_index < len(traces),
                        f"line {line_number}: trace reference out of range")
                traces.append(traces[parent_index] | frozenset(instructions[instruction_index]))
            elif tag == "a":
                fields = rest.split()
                require(len(fields) == 2, f"line {line_number}: malformed allocation record")
                size = hex_integer(fields[0], line_number=line_number, field="allocation size")
                trace_index = hex_integer(fields[1], line_number=line_number,
                                          field="allocation trace reference")
                require(trace_index < len(traces), f"line {line_number}: allocation trace out of range")
                allocations.append((size, trace_index))
                active_allocations.append(0)
            elif tag == "+":
                index = hex_integer(rest, line_number=line_number, field="allocation event reference")
                require(index < len(allocations), f"line {line_number}: allocation event out of range")
                size, _ = allocations[index]
                active_allocations[index] += 1
                allocation_events += 1
                requested_bytes += size
                live_bytes += size
                peak_live_bytes = max(peak_live_bytes, live_bytes)
            elif tag == "-":
                index = hex_integer(rest, line_number=line_number,
                                    field="deallocation event reference")
                require(index < len(allocations), f"line {line_number}: deallocation event out of range")
                require(active_allocations[index] > 0,
                        f"line {line_number}: deallocation has no live allocation at index {index}")
                size, _ = allocations[index]
                active_allocations[index] -= 1
                deallocation_events += 1
                live_bytes -= size
                require(live_bytes >= 0, f"line {line_number}: live heap underflow")
            elif tag == "R":
                # R is an RSS sample, not an end marker.  It can occur many
                # times throughout the profile, including before more +/−
                # events.  The final strings/ips comments below provide the
                # truncation check for interpreted profiles.
                hex_integer(rest.strip(), line_number=line_number, field="RSS sample")
            elif tag == "#":
                comment = rest.strip()
                for name in ("strings", "ips"):
                    prefix = f"{name}:"
                    if comment.startswith(prefix):
                        count = comment[len(prefix):].strip()
                        hex_integer(count, line_number=line_number,
                                    field=f"{name} summary count")
                        require(name not in summary_comments,
                                f"line {line_number}: repeated {name} summary")
                        summary_comments.add(name)
                        break
            elif tag in {"", "I", "c"}:
                continue
            else:
                raise ParseError(f"line {line_number}: unsupported record tag {tag!r}")
    finally:
        stream.close()

    require(version_seen, "profile has no supported format version")
    require(command is not None, "profile has no debuggee command")
    require(summary_comments == {"strings", "ips"},
            "profile is missing final strings/ips summaries (truncated profile)")
    require(allocation_events >= deallocation_events,
            "deallocation event count exceeds allocation event count")
    return {
        "schema": "litchi.performance.0923.heaptrack-profile.v2",
        "profile": descriptor(path),
        "format": EXPECTED_VERSION,
        "debuggee_command": command,
        "allocation_event_count": allocation_events,
        "deallocation_event_count": deallocation_events,
        "requested_bytes": requested_bytes,
        "peak_live_heap_bytes": peak_live_bytes,
        "live_bytes_at_end": live_bytes,
        "live_allocation_info_entries": sum(value > 0 for value in active_allocations),
        "scope": "whole_process",
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("profile", type=Path)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    try:
        value = parse(args.profile.resolve())
        data = json.dumps(value, indent=2, sort_keys=True) + "\n"
        if args.output:
            output = args.output.resolve()
            if output.exists():
                raise ParseError(f"refusing to replace retained output: {output}")
            output.write_text(data, encoding="utf-8")
        else:
            print(data, end="")
        return 0
    except (OSError, ValueError, ParseError, subprocess.SubprocessError) as error:
        print(f"0923 Heaptrack parse failed: {error}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
