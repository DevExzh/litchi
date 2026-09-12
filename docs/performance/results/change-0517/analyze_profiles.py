#!/usr/bin/env python3
"""Validate and summarize scoped publication Callgrind profiles.

The profile is collected only while the selected publication function is on
the stack.  This parser uses the raw positive-call edge to recover that
inclusive total; a process summary alone is not sufficient evidence of scope.
"""

from __future__ import annotations

import argparse
import collections
import glob
import hashlib
import json
import re
import sys
from pathlib import Path


FUNCTION_RE = re.compile(r"^(fn|cfn)=\((\d+)\)(?:\s+(.*))?$")
CALLS_RE = re.compile(r"^calls=([\d,]+)")
SUMMARY_RE = re.compile(r"^summary:\s*(.*)$")


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _integer_values(text: str) -> list[int]:
    values: list[int] = []
    for token in text.split():
        try:
            values.append(int(token))
        except ValueError:
            break
    return values


def _cost(line: str, event_index: int) -> int | None:
    """Return the selected event from a Callgrind cost line."""
    fields = line.split()
    if len(fields) <= event_index + 1:
        return None
    position = fields[0]
    if not position or position[0] not in "+-*" and not position[0].isdigit():
        return None
    try:
        return int(fields[event_index + 1])
    except ValueError:
        return None


def parse_profile(path: Path, selected_name: str) -> dict:
    events: list[str] = []
    summary_values: list[int] | None = None
    names: dict[int, str] = {}
    functions: dict[int, dict] = collections.defaultdict(
        lambda: {"name": "", "self": 0, "edges": []}
    )
    current_id: int | None = None
    pending_callee_id: int | None = None
    pending_calls: int | None = None
    pending_line: int | None = None

    with path.open("r", encoding="utf-8", errors="replace") as stream:
        for line_number, raw_line in enumerate(stream, 1):
            line = raw_line.strip()
            if not line:
                continue
            if line.startswith("events:"):
                events = line.split(":", 1)[1].split()
                continue
            summary_match = SUMMARY_RE.match(line)
            if summary_match:
                summary_values = _integer_values(summary_match.group(1))
                continue

            function_match = FUNCTION_RE.match(line)
            if function_match:
                kind, function_text_id, function_name = function_match.groups()
                function_id = int(function_text_id)
                if function_name:
                    names[function_id] = function_name
                if kind == "fn":
                    current_id = function_id
                    pending_callee_id = pending_calls = pending_line = None
                    if function_name:
                        functions[function_id]["name"] = function_name
                else:
                    if current_id is None:
                        raise ValueError(f"{path}:{line_number}: cfn outside fn")
                    pending_callee_id = function_id
                    pending_calls = None
                    pending_line = line_number
                continue

            calls_match = CALLS_RE.match(line)
            if calls_match:
                if current_id is None or pending_callee_id is None:
                    raise ValueError(f"{path}:{line_number}: calls outside cfn")
                if pending_calls is not None:
                    raise ValueError(f"{path}:{line_number}: multiple calls records for cfn")
                pending_calls = int(calls_match.group(1).replace(",", ""))
                continue

            # These records describe object/source context, not event costs.
            if line.startswith(
                (
                    "creator:",
                    "cmd:",
                    "pid:",
                    "version:",
                    "positions:",
                    "part:",
                    "ob=",
                    "cob=",
                    "fl=",
                    "cfi=",
                    "fi=",
                    "fe=",
                    "totals:",
                )
            ):
                continue

            if current_id is None or not events:
                continue
            value = _cost(line, events.index("Ir") if "Ir" in events else 0)
            if value is None:
                continue
            if pending_callee_id is not None and pending_calls is not None:
                functions[current_id]["edges"].append(
                    {
                        "callee_id": pending_callee_id,
                        "calls": pending_calls,
                        "cost": value,
                        "line": pending_line,
                    }
                )
            else:
                functions[current_id]["self"] += value
            # A calls record describes this one cost record.  Following costs
            # are self costs until the next cfn/calls pair.
            pending_callee_id = pending_calls = pending_line = None

    if "Ir" not in events:
        raise ValueError(f"{path}: no Ir event in events header")
    event_index = events.index("Ir")
    if summary_values is None or len(summary_values) <= event_index:
        raise ValueError(f"{path}: no usable Ir summary")
    summary_ir = summary_values[event_index]

    for function_id, function in functions.items():
        if not function["name"] and function_id in names:
            function["name"] = names[function_id]

    selected_ids = [
        function_id
        for function_id, function in functions.items()
        if selected_name in function["name"]
    ]
    if len(selected_ids) != 1:
        raise ValueError(
            f"{path}: selected function {selected_name!r} matched {selected_ids}"
        )
    selected_id = selected_ids[0]
    selected = functions[selected_id]

    incoming = []
    for parent_id, function in functions.items():
        for edge in function["edges"]:
            if (
                edge["callee_id"] == selected_id
                and edge["calls"] is not None
                and edge["calls"] > 0
            ):
                incoming.append(
                    {
                        "caller_id": parent_id,
                        "caller": functions[parent_id]["name"]
                        or names.get(parent_id, ""),
                        "calls": edge["calls"],
                        "inclusive_ir": edge["cost"],
                        "line": edge["line"],
                    }
                )

    direct_by_id: dict[int, dict] = {}
    for edge in selected["edges"]:
        if edge["cost"] <= 0 and not edge["calls"]:
            continue
        callee_id = edge["callee_id"]
        item = direct_by_id.setdefault(
            callee_id,
            {
                "callee_id": callee_id,
                "callee": names.get(callee_id, ""),
                "calls": 0,
                "edge_count": 0,
                "inclusive_ir": 0,
            },
        )
        item["calls"] += edge["calls"] or 0
        item["edge_count"] += 1
        item["inclusive_ir"] += edge["cost"]

    direct_ir = sum(edge["cost"] for edge in selected["edges"])
    selected_self_ir = selected["self"]
    exact_one_call = len(incoming) == 1 and incoming[0]["calls"] == 1
    selected_inclusive_ir = incoming[0]["inclusive_ir"] if len(incoming) == 1 else None
    equation_valid = (
        selected_inclusive_ir is not None
        and selected_inclusive_ir == selected_self_ir + direct_ir
    )
    scoped_total = selected_inclusive_ir == summary_ir

    receipt_path = path.with_suffix(".json")
    if not receipt_path.exists():
        raise ValueError(f"{path}: missing sibling receipt {receipt_path}")
    receipt = json.loads(receipt_path.read_text(encoding="utf-8"))
    receipt_checks = {
        "exit_code_zero": receipt.get("exit_code") == 0,
        "cleanup_verified": receipt.get("cleanup_verified") is True,
        "samples_one": receipt.get("samples") == 1,
        "warmups_zero": receipt.get("warmups") == 0,
        "repeats_one": receipt.get("repeats") == 1,
    }
    if not all(receipt_checks.values()):
        failed = [key for key, value in receipt_checks.items() if not value]
        raise ValueError(f"{path}: receipt checks failed: {', '.join(failed)}")

    direct = sorted(
        direct_by_id.values(), key=lambda item: (-item["inclusive_ir"], item["callee_id"])
    )
    result = {
        "path": str(path),
        "sha256": sha256(path),
        "receipt": str(receipt_path),
        "source_manifest_sha256": receipt.get("source_manifest_sha256"),
        "binary_sha256": receipt.get("binary_sha256"),
        "selected_function": selected["name"],
        "selected_function_id": selected_id,
        "summary_ir": summary_ir,
        "positive_incoming_edges": incoming,
        "selected_call_count": sum(edge["calls"] for edge in incoming),
        "selected_inclusive_ir": selected_inclusive_ir,
        "selected_self_ir": selected_self_ir,
        "selected_direct_ir": direct_ir,
        "selected_direct_callees": direct,
        "validation": {
            "exactly_one_positive_call": exact_one_call,
            "selected_self_plus_direct_equals_inclusive": equation_valid,
            "publication_total_scoped_to_summary": scoped_total,
            "receipt": receipt_checks,
        },
    }
    if not all(
        (
            exact_one_call,
            equation_valid,
            scoped_total,
        )
    ):
        raise ValueError(f"{path}: raw publication scope validation failed")
    return result


def expand_paths(arguments: list[str], default_dir: Path) -> list[Path]:
    patterns = arguments or [str(default_dir / "*.callgrind")]
    paths: list[Path] = []
    for pattern in patterns:
        matches = (
            [Path(match) for match in glob.glob(pattern)]
            if any(char in pattern for char in "*?[")
            else [Path(pattern)]
        )
        paths.extend(matches)
    unique = sorted({path for path in paths if path.is_file()})
    if not unique:
        raise ValueError("no Callgrind profiles matched")
    return unique


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("profiles", nargs="*", help="raw .callgrind files or glob patterns")
    parser.add_argument(
        "--function",
        default="litchi_docx::source_backed::Package::publish_document_commit_to_stream",
        help="selected publication function substring",
    )
    parser.add_argument("--output", type=Path, help="write JSON analysis to this path")
    args = parser.parse_args()

    try:
        root = Path(__file__).resolve().parent
        paths = expand_paths(args.profiles, root / "profile-preflight")
        profiles = [parse_profile(path, args.function) for path in paths]
    except (OSError, ValueError, json.JSONDecodeError) as error:
        print(f"analyze_profiles.py: error: {error}", file=sys.stderr)
        return 2

    document = {
        "schema": "docx_callgrind_publication_analysis_v1",
        "selected_function": args.function,
        "profile_count": len(profiles),
        "validation": {
            "all_profiles_exactly_one_positive_call": all(
                profile["validation"]["exactly_one_positive_call"] for profile in profiles
            ),
            "all_profiles_publication_total_scoped": all(
                profile["validation"]["publication_total_scoped_to_summary"]
                for profile in profiles
            ),
        },
        "profiles": profiles,
    }
    encoded = json.dumps(document, indent=2, sort_keys=False) + "\n"
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(encoded, encoding="utf-8")
    else:
        print(encoded, end="")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
