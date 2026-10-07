#!/usr/bin/env python3
"""Independent byte-span audit of retained 0904 XLSX XML fixtures.

The audit uses Expat's namespace-aware callbacks only to locate complete cell
spans. It binds every input through the retained 0947 identity receipt, matches
cells by worksheet member, row number, and cell address, and compares the byte
stream left after removing changed cell spans. It makes no Rust or workload
calls.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import re
from pathlib import Path
import xml.parsers.expat as expat
from typing import Any

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
SEP = "\x1f"


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1 << 20), b""):
            digest.update(block)
    return digest.hexdigest()


def unique_json(path: Path) -> Any:
    def reject_duplicate(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
        out: dict[str, Any] = {}
        for key, value in pairs:
            if key in out:
                raise ValueError(f"duplicate JSON key {key!r} in {path}")
            out[key] = value
        return out
    return json.loads(path.read_text(encoding="utf-8"), object_pairs_hook=reject_duplicate)


def require(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)


def local_name(expanded: str) -> tuple[str, str]:
    if SEP in expanded:
        uri, local = expanded.split(SEP, 1)
        return uri, local
    return "", expanded


def tag_end(data: bytes, start: int) -> tuple[int, bool]:
    quote: int | None = None
    index = start
    while index < len(data):
        byte = data[index]
        if quote is not None:
            if byte == quote:
                quote = None
        elif byte in (ord("'"), ord('"')):
            quote = byte
        elif byte == ord(">"):
            before = data[start:index].rstrip()
            return index + 1, before.endswith(b"/")
        index += 1
    raise ValueError(f"unterminated XML start tag at byte {start}")


def close_tag_end(data: bytes, start: int) -> int:
    quote: int | None = None
    index = start
    while index < len(data):
        byte = data[index]
        if quote is not None:
            if byte == quote:
                quote = None
        elif byte in (ord("'"), ord('"')):
            quote = byte
        elif byte == ord(">"):
            return index + 1
        index += 1
    raise ValueError(f"unterminated XML end tag at byte {start}")


def parse_xml(path: Path) -> dict[str, Any]:
    data = path.read_bytes()
    parser = expat.ParserCreate(namespace_separator=SEP)
    stack: list[dict[str, Any]] = []
    cells: dict[tuple[str | None, str], dict[str, Any]] = {}
    namespace_decls: list[dict[str, str | None]] = []
    mce_ignorables: list[str] = []

    def start_namespace(prefix: str | None, uri: str) -> None:
        namespace_decls.append({"prefix": prefix, "uri": uri})

    def start(name: str, attrs: dict[str, str]) -> None:
        uri, local = local_name(name)
        start_at = parser.CurrentByteIndex
        end_at, self_closing = tag_end(data, start_at)
        row = (attrs.get("r") if uri == MAIN_NS and local == "row"
               else next((frame["row"] for frame in reversed(stack)
                          if frame["local"] == "row" and frame["uri"] == MAIN_NS), None))
        frame = {"name": name, "uri": uri, "local": local, "start": start_at,
                 "start_end": end_at, "self_closing": self_closing, "row": row,
                 "attrs": dict(attrs)}
        stack.append(frame)
        if uri == MC_NS and local == "AlternateContent":
            frame["mce_alternate_content"] = True
        for key, value in attrs.items():
            key_uri, key_local = local_name(key)
            if key_uri == MC_NS and key_local == "Ignorable":
                mce_ignorables.append(value)
        if uri == MAIN_NS and local == "c":
            address = attrs.get("r")
            require(isinstance(address, str) and address, f"cell without address in {path}")
            match = re.search(r"([0-9]+)$", address)
            require(match is not None, f"cell address lacks row number {address!r} in {path}")
            address_row = match.group(1)
            require(row is not None and address_row == row,
                    f"cell row/address mismatch {row!r}/{address!r} in {path}")
            key = (row, address)
            require(key not in cells, f"duplicate cell {key!r} in {path}")
            cells[key] = {"row": row, "address": address, "namespace": uri,
                          "start": start_at, "start_end": end_at,
                          "self_closing": self_closing, "attrs": dict(attrs),
                          "end": None}

    def end(name: str) -> None:
        require(stack, f"unbalanced XML end {name} in {path}")
        frame = stack.pop()
        require(frame["name"] == name, f"XML stack mismatch in {path}")
        if frame["uri"] == MAIN_NS and frame["local"] == "c":
            key = (frame["row"], frame["attrs"].get("r")) if "attrs" in frame else None
            # The cell frame is found by its unique start offset, which avoids
            # relying on Expat's self-closing callback byte index.
            found = next((item for item in cells.values()
                          if item["start"] == frame["start"]), None)
            require(found is not None, f"cell end without cell start in {path}")
            end_at = frame["start_end"] if frame["self_closing"] else close_tag_end(
                data, parser.CurrentByteIndex)
            found["end"] = end_at
            found["full_span_bytes"] = end_at - frame["start"]
            found["balanced"] = True

    parser.StartNamespaceDeclHandler = start_namespace
    parser.StartElementHandler = start
    parser.EndElementHandler = end
    try:
        parser.Parse(data, True)
    except (expat.ExpatError, AssertionError, ValueError) as error:
        raise ValueError(f"XML parse failed for {path}: {error}") from error
    require(not stack, f"unclosed XML elements in {path}")
    for key, cell in cells.items():
        require(cell["end"] is not None and cell["end"] > cell["start"],
                f"cell span incomplete for {key!r} in {path}")
    return {"path": str(path), "bytes": data, "cells": cells,
            "namespaces": namespace_decls, "mce_ignorables": mce_ignorables,
            "sha256": sha256(path)}


def gap_parts(data: bytes, intervals: list[tuple[int, int]]) -> list[bytes]:
    cursor = 0
    parts: list[bytes] = []
    for start, end in sorted(intervals):
        require(0 <= start < end <= len(data), "invalid byte interval")
        require(start >= cursor, "overlapping byte intervals")
        parts.append(data[cursor:start])
        cursor = end
    parts.append(data[cursor:])
    return parts


def without_intervals(data: bytes, intervals: list[tuple[int, int]]) -> bytes:
    return b"".join(gap_parts(data, intervals))


def xml_diff_summary(original: bytes, replacement: bytes) -> dict[str, Any]:
    # This is only for a small non-cell metadata file (workbook.xml), so a
    # compact common-prefix/suffix description is sufficient and exact.
    prefix = 0
    while prefix < len(original) and prefix < len(replacement) and original[prefix] == replacement[prefix]:
        prefix += 1
    suffix = 0
    while (suffix < len(original) - prefix and suffix < len(replacement) - prefix
           and original[len(original) - 1 - suffix] == replacement[len(replacement) - 1 - suffix]):
        suffix += 1
    return {"common_prefix_bytes": prefix, "common_suffix_bytes": suffix,
            "original_gap_bytes": len(original) - prefix - suffix,
            "replacement_gap_bytes": len(replacement) - prefix - suffix,
            "original_gap_sha256": hashlib.sha256(original[prefix:len(original)-suffix or None]).hexdigest(),
            "replacement_gap_sha256": hashlib.sha256(replacement[prefix:len(replacement)-suffix or None]).hexdigest()}


def compare_member(original_path: Path, replacement_path: Path, expected: dict[str, Any]) -> dict[str, Any]:
    original = parse_xml(original_path)
    replacement = parse_xml(replacement_path)
    require(original["sha256"] == expected["original"]["sha256"], f"original hash mismatch {original_path}")
    require(replacement["sha256"] == expected["replacement"]["sha256"], f"replacement hash mismatch {replacement_path}")
    require(len(original["bytes"]) == expected["original"]["bytes"], f"original size mismatch {original_path}")
    require(len(replacement["bytes"]) == expected["replacement"]["bytes"], f"replacement size mismatch {replacement_path}")
    original_keys = set(original["cells"])
    replacement_keys = set(replacement["cells"])
    require(original_keys == replacement_keys, f"cell keys differ in {original_path}")
    changed = []
    unchanged = 0
    for key in sorted(original_keys, key=lambda item: ((item[0] or ""), item[1])):
        left = original["cells"][key]
        right = replacement["cells"][key]
        require(left["balanced"] and right["balanced"], f"unbalanced cell {key}")
        left_bytes = original["bytes"][left["start"]:left["end"]]
        right_bytes = replacement["bytes"][right["start"]:right["end"]]
        if left_bytes == right_bytes:
            unchanged += 1
            continue
        changed.append({
            "row": key[0], "address": key[1], "namespace": left["namespace"],
            "original_start": left["start"], "original_end": left["end"],
            "replacement_start": right["start"], "replacement_end": right["end"],
            "original_bytes": len(left_bytes), "replacement_bytes": len(right_bytes),
            "original_sha256": hashlib.sha256(left_bytes).hexdigest(),
            "replacement_sha256": hashlib.sha256(right_bytes).hexdigest(),
        })
    left_intervals = [
        (original["cells"][(item["row"], item["address"])] ["start"],
         original["cells"][(item["row"], item["address"])] ["end"])
        for item in changed
    ]
    right_intervals = [
        (replacement["cells"][(item["row"], item["address"])] ["start"],
         replacement["cells"][(item["row"], item["address"])] ["end"])
        for item in changed
    ]
    left_gaps = gap_parts(original["bytes"], left_intervals)
    right_gaps = gap_parts(replacement["bytes"], right_intervals)
    require(len(left_gaps) == len(right_gaps), "gap count differs")
    left_outside = b"".join(left_gaps)
    right_outside = b"".join(right_gaps)
    gap_pairs = [
        {
            "index": index,
            "original_bytes": len(left), "replacement_bytes": len(right),
            "original_sha256": hashlib.sha256(left).hexdigest(),
            "replacement_sha256": hashlib.sha256(right).hexdigest(),
            "byte_identical": left == right,
        }
        for index, (left, right) in enumerate(zip(left_gaps, right_gaps))
    ]
    return {
        "original_bytes": len(original["bytes"]), "replacement_bytes": len(replacement["bytes"]),
        "original_sha256": original["sha256"], "replacement_sha256": replacement["sha256"],
        "namespace_declarations_original": original["namespaces"],
        "namespace_declarations_replacement": replacement["namespaces"],
        "mce_ignorables_original": original["mce_ignorables"],
        "mce_ignorables_replacement": replacement["mce_ignorables"],
        "namespace_aware_main_cells": True,
        "original_cell_count": len(original_keys), "replacement_cell_count": len(replacement_keys),
        "original_full_cell_span_count": len(original_keys),
        "replacement_full_cell_span_count": len(replacement_keys),
        "matched_unique_cell_count": len(original_keys), "unchanged_cell_count": unchanged,
        "changed_cell_count": len(changed),
        "changed_cell_original_bytes": sum(item["original_bytes"] for item in changed),
        "changed_cell_replacement_bytes": sum(item["replacement_bytes"] for item in changed),
        "changed_cell_original_fraction": (sum(item["original_bytes"] for item in changed) / len(original["bytes"])
                                           if original["bytes"] else None),
        "changed_cell_replacement_fraction": (sum(item["replacement_bytes"] for item in changed) / len(replacement["bytes"])
                                               if replacement["bytes"] else None),
        "changed_cells": changed,
        "outside_changed_cell_spans": {
            "gap_count": len(left_gaps),
            "gap_pairs": gap_pairs,
            "original_gap_lengths": [len(part) for part in left_gaps],
            "replacement_gap_lengths": [len(part) for part in right_gaps],
            "gap_lengths_equal": [len(part) for part in left_gaps] == [len(part) for part in right_gaps],
            "all_gap_bytes_identical": all(a == b for a, b in zip(left_gaps, right_gaps)),
            "original_bytes": len(left_outside), "replacement_bytes": len(right_outside),
            "original_sha256": hashlib.sha256(left_outside).hexdigest(),
            "replacement_sha256": hashlib.sha256(right_outside).hexdigest(),
            "byte_identical": left_outside == right_outside,
        },
        "balanced_full_cell_spans": True,
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--repo", type=Path, required=True)
    parser.add_argument("--identity", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    repo = args.repo.resolve()
    identity_path = args.identity.resolve()
    identity = unique_json(identity_path)
    require(identity["status"] == "pass", "identity receipt is not passing")
    artifact_root = repo / "docs/performance/results/change-0904/artifacts"
    receipt: dict[str, Any] = {
        "schema": "litchi.xlsx-changed-cell-span-review-0947.v4",
        "status": "pass",
        "scope": "Untimed offline namespace-aware XML span audit of retained 0904 fixtures; no Rust/build/workload/native measurement.",
        "identity_receipt": {"path": str(identity_path), "bytes": identity_path.stat().st_size,
                             "sha256": sha256(identity_path)},
        "method": {
            "parser": "Python Expat namespace-aware callbacks with exact UTF-8 byte offsets",
            "cell_key": "worksheet member + row @r + main-namespace cell @r",
            "changed_span": "complete balanced main-namespace <c>...</c> or self-closing <c/> span whose raw bytes differ",
            "outside_check": "remove matched changed cell spans from each side and compare the remaining byte stream exactly",
            "namespace_and_mce": "record namespace declarations and mc:Ignorable; never normalize or discard XML bytes",
        },
        "cases": [],
        "limitations": [
            "Workbook.xml has no cell spans; its calcPr addition is reported as a separate non-cell metadata gap.",
            "A changed cell span proves byte locality, not semantic equivalence or safe publication reuse.",
            "The retained identity receipt binds fixture bytes and adopted-candidate identities; this review performs no new candidate execution.",
        ],
    }
    for case in identity["cases"]:
        case_out = {"id": case["id"], "changed_members": case["changed_members"], "parts": []}
        for part in case["parts"]:
            original = artifact_root / part["original"]["path"]
            replacement = artifact_root / part["replacement"]["path"]
            result = compare_member(original, replacement, part)
            result["member"] = part["member"]
            if result["changed_cell_count"] == 0 and result["original_sha256"] != result["replacement_sha256"]:
                result["non_cell_difference"] = xml_diff_summary(original.read_bytes(), replacement.read_bytes())
                result["non_cell_difference"]["kind"] = "metadata-only XML difference"
            case_out["parts"].append(result)
        case_out["totals"] = {
            "changed_member_count": len(case_out["parts"]),
            "members_with_changed_cells": sum(p["changed_cell_count"] > 0 for p in case_out["parts"]),
            "changed_cell_pair_count": sum(p["changed_cell_count"] for p in case_out["parts"]),
            "changed_cell_original_bytes": sum(p["changed_cell_original_bytes"] for p in case_out["parts"]),
            "changed_cell_replacement_bytes": sum(p["changed_cell_replacement_bytes"] for p in case_out["parts"]),
            "outside_cell_span_checks_passed": sum(p["outside_changed_cell_spans"]["byte_identical"] for p in case_out["parts"]),
            "outside_cell_span_checks_total": len(case_out["parts"]),
        }
        receipt["cases"].append(case_out)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    require(not args.output.exists(), f"refusing existing output {args.output}")
    args.output.write_text(json.dumps(receipt, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    print(json.dumps({"status": "pass", "output": str(args.output),
                      "output_sha256": sha256(args.output),
                      "output_bytes": args.output.stat().st_size,
                      "cases": [{"id": c["id"], "totals": c["totals"]} for c in receipt["cases"]]}, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
