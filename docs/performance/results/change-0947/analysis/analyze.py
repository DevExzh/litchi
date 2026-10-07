#!/usr/bin/env python3
"""Untimed diagnostics for retained XLSX source-replacement XML pairs.

The Rust probe is authoritative for the production ReplacementProof.  This
file only adds byte-diff spans and a deliberately conservative lexical
eligibility estimate for a future multi-window proof.  It never decides or
replaces a publication audit.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable


SCHEMA = "litchi.xlsx-replacement-proof-diagnostic.v1"
NAME_RE = re.compile(rb"[A-Za-z_][A-Za-z0-9_.:-]*")
MC_MARKERS = (
    b"mc:",
    b"AlternateContent",
    b"Fallback",
    b"Choice",
    b"Ignorable",
    b"ProcessContent",
    b"PreserveAttributes",
    b"PreserveElements",
    b"MustUnderstand",
)


@dataclass(frozen=True)
class Token:
    start: int
    end: int
    kind: str
    name: bytes | None = None
    depth: int = 0
    namespace_marker: bool = False
    mc_marker: bool = False


@dataclass(frozen=True)
class Element:
    start: int
    end: int
    name: bytes
    depth: int
    start_token: Token
    end_token: Token
    namespace_marker: bool
    mc_marker: bool


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def span_dict(start: int, end: int) -> dict[str, int]:
    return {"start": start, "end": end, "bytes": end - start}


def find_tag_end(data: bytes, start: int) -> int:
    quote = None
    index = start + 1
    while index < len(data):
        value = data[index]
        if quote is not None:
            if value == quote:
                quote = None
        elif value in (ord('"'), ord("'")):
            quote = value
        elif value == ord('>'):
            return index + 1
        index += 1
    raise ValueError(f"unterminated markup at {start}")


def token_name(raw: bytes, start: int) -> bytes | None:
    match = NAME_RE.match(raw, start)
    return match.group(0) if match else None


def marker(raw: bytes) -> tuple[bool, bool]:
    namespace = b"xmlns" in raw
    compatibility = any(part in raw for part in MC_MARKERS)
    return namespace, compatibility


def scan_lexically(data: bytes) -> tuple[list[Token], list[Element], list[str]]:
    """Collect token and balanced-element spans for already-valid XML.

    This is intentionally a diagnostic lexer.  It handles the XML constructs
    present in the retained OOXML artifacts, tracks quoted `>` characters and
    matching element nesting, and refuses to manufacture a balanced result
    when it encounters an unsupported shape.
    """

    tokens: list[Token] = []
    elements: list[Element] = []
    stack: list[tuple[bytes, int, int, Token, bool, bool]] = []
    reasons: list[str] = []
    index = 0
    while index < len(data):
        if data[index] != ord('<'):
            end = data.find(b"<", index)
            if end < 0:
                end = len(data)
            if end > index:
                tokens.append(Token(index, end, "text", depth=len(stack)))
            index = end
            continue

        if data.startswith(b"<!--", index):
            end = data.find(b"-->", index + 4)
            if end < 0:
                reasons.append("unterminated_comment")
                break
            end += 3
            namespace, compatibility = marker(data[index:end])
            tokens.append(Token(index, end, "markup", depth=len(stack),
                                namespace_marker=namespace, mc_marker=compatibility))
            index = end
            continue
        if data.startswith(b"<![CDATA[", index):
            end = data.find(b"]]>", index + 9)
            if end < 0:
                reasons.append("unterminated_cdata")
                break
            end += 3
            tokens.append(Token(index, end, "cdata", depth=len(stack)))
            index = end
            continue
        if data.startswith(b"<?", index):
            end = data.find(b"?>", index + 2)
            if end < 0:
                reasons.append("unterminated_pi")
                break
            end += 2
            namespace, compatibility = marker(data[index:end])
            tokens.append(Token(index, end, "markup", depth=len(stack),
                                namespace_marker=namespace, mc_marker=compatibility))
            index = end
            continue
        if data.startswith(b"<!", index):
            try:
                end = find_tag_end(data, index)
            except ValueError:
                reasons.append("unterminated_declaration")
                break
            namespace, compatibility = marker(data[index:end])
            tokens.append(Token(index, end, "markup", depth=len(stack),
                                namespace_marker=namespace, mc_marker=compatibility))
            index = end
            continue

        try:
            end = find_tag_end(data, index)
        except ValueError:
            reasons.append("unterminated_tag")
            break
        closing = data.startswith(b"</", index)
        body_start = index + (2 if closing else 1)
        name = token_name(data, body_start)
        if name is None:
            reasons.append("missing_element_name")
            index = end
            continue
        namespace, compatibility = marker(data[index:end])
        if closing:
            token = Token(index, end, "end", name=name, depth=len(stack),
                          namespace_marker=namespace, mc_marker=compatibility)
            tokens.append(token)
            if not stack:
                reasons.append("unmatched_end")
            else:
                opened_name, opened_start, opened_depth, opened_token, opened_ns, opened_mc = stack.pop()
                if opened_name != name:
                    reasons.append("mismatched_end")
                elements.append(Element(
                    opened_start,
                    end,
                    opened_name,
                    opened_depth,
                    opened_token,
                    token,
                    opened_ns or namespace,
                    opened_mc or compatibility,
                ))
            index = end
            continue

        stripped = data[index:end - 1].rstrip()
        empty = stripped.endswith(b"/")
        kind = "empty" if empty else "start"
        token = Token(index, end, kind, name=name, depth=len(stack),
                      namespace_marker=namespace, mc_marker=compatibility)
        tokens.append(token)
        if empty:
            elements.append(Element(index, end, name, len(stack), token, token,
                                    namespace, compatibility))
        else:
            stack.append((name, index, len(stack), token, namespace, compatibility))
        index = end
    if stack:
        reasons.append("unclosed_element")
    return tokens, elements, reasons


def local_difference(source: bytes, replacement: bytes, source_start: int = 0, replacement_start: int = 0) -> dict[str, Any] | None:
    """Return one linear common-prefix/suffix difference for paired regions."""

    prefix = 0
    common = min(len(source), len(replacement))
    while prefix < common and source[prefix] == replacement[prefix]:
        prefix += 1
    if prefix == len(source) and prefix == len(replacement):
        return None
    suffix = 0
    while (
        suffix < len(source) - prefix
        and suffix < len(replacement) - prefix
        and source[len(source) - 1 - suffix] == replacement[len(replacement) - 1 - suffix]
    ):
        suffix += 1
    source_end = len(source) - suffix
    replacement_end = len(replacement) - suffix
    source_span = span_dict(source_start + prefix, source_start + source_end)
    replacement_span = span_dict(replacement_start + prefix, replacement_start + replacement_end)
    if source_span["bytes"] == 0:
        kind = "insert"
    elif replacement_span["bytes"] == 0:
        kind = "delete"
    else:
        kind = "replace"
    return {
        "kind": kind,
        "source": source_span,
        "replacement": replacement_span,
        "source_sha256": sha256_bytes(source[prefix:source_end]),
        "replacement_sha256": sha256_bytes(replacement[prefix:replacement_end]),
    }


def common_prefix_suffix(original: bytes, replacement: bytes) -> tuple[int, int]:
    prefix = 0
    common = min(len(original), len(replacement))
    while prefix < common and original[prefix] == replacement[prefix]:
        prefix += 1
    suffix = 0
    while (
        suffix < len(original) - prefix
        and suffix < len(replacement) - prefix
        and original[len(original) - 1 - suffix] == replacement[len(replacement) - 1 - suffix]
    ):
        suffix += 1
    return prefix, suffix


def attribute_value(data: bytes, element: Element, name: bytes) -> bytes | None:
    raw = data[element.start_token.start:element.start_token.end]
    match = re.search(rb"(?:^|\s)" + re.escape(name) + rb"\s*=\s*([\"'])(.*?)\1", raw)
    return match.group(2) if match else None


def cell_elements(data: bytes, elements: list[Element]) -> dict[bytes, Element]:
    cells: dict[bytes, Element] = {}
    for element in elements:
        if element.name.rsplit(b":", 1)[-1] != b"c":
            continue
        key = attribute_value(data, element, b"r")
        if key is not None and key not in cells:
            cells[key] = element
    return cells


def payload_difference(
    original: bytes,
    replacement: bytes,
    source_elements: list[Element],
    replacement_elements: list[Element],
) -> tuple[list[dict[str, Any]], str]:
    """Find changed source/output spans without quadratic repeated-byte matching.

    XLSX value fixtures preserve cell addresses, so pairing those balanced
    elements gives exact local byte spans for the scattered edits. Workbook
    metadata and any unmatched shape use a single linear common-prefix/suffix
    span. The result is a diagnostic partition, not a semantic writer trace.
    """

    source_cells = cell_elements(original, source_elements)
    replacement_cells = cell_elements(replacement, replacement_elements)
    changes: list[dict[str, Any]] = []
    paired = 0
    for key in sorted(source_cells.keys() | replacement_cells.keys()):
        source_element = source_cells.get(key)
        replacement_element = replacement_cells.get(key)
        if source_element is None:
            changes.append({
                "kind": "insert-element",
                "source": span_dict(0, 0),
                "replacement": span_dict(replacement_element.start, replacement_element.end),
                "source_sha256": sha256_bytes(b""),
                "replacement_sha256": sha256_bytes(replacement[replacement_element.start:replacement_element.end]),
                "cell": key.decode("utf-8", "replace"),
            })
            continue
        if replacement_element is None:
            changes.append({
                "kind": "delete-element",
                "source": span_dict(source_element.start, source_element.end),
                "replacement": span_dict(0, 0),
                "source_sha256": sha256_bytes(original[source_element.start:source_element.end]),
                "replacement_sha256": sha256_bytes(b""),
                "cell": key.decode("utf-8", "replace"),
            })
            continue
        paired += 1
        change = local_difference(
            original[source_element.start:source_element.end],
            replacement[replacement_element.start:replacement_element.end],
            source_element.start,
            replacement_element.start,
        )
        if change is not None:
            change["cell"] = key.decode("utf-8", "replace")
            changes.append(change)
    if changes or original == replacement:
        return changes, "paired-cell-local-prefix-suffix"
    # A non-cell worksheet change or a workbook change is retained as one
    # bounded span. This path is linear and intentionally does not guess an
    # edit alignment from repeated XML bytes.
    change = local_difference(original, replacement)
    return ([change] if change is not None else []), "whole-payload-prefix-suffix"


def enclosing(
    elements: Iterable[Element],
    start: int,
    end: int,
    target_local_name: bytes | None = None,
) -> Element | None:
    candidates = [
        e for e in elements
        if e.depth > 0
        and (target_local_name is None or e.name.rsplit(b":", 1)[-1] == target_local_name)
        and e.start <= start
        and e.end >= end
    ]
    if not candidates:
        return None
    return min(candidates, key=lambda e: (e.end - e.start, e.start))


def contains(element: Element, start: int, end: int) -> bool:
    if start == end:
        return element.start <= start <= element.end
    return element.start <= start and end <= element.end


def intersects(token: Token, start: int, end: int) -> bool:
    if start == end:
        return token.start <= start <= token.end
    return token.start < end and start < token.end


def candidate_for_change(
    change: dict[str, Any],
    source_elements: list[Element],
    replacement_elements: list[Element],
    target_local_name: bytes,
) -> tuple[dict[str, Any] | None, str | None]:
    source = change["source"]
    replacement = change["replacement"]
    source_element = enclosing(source_elements, source["start"], source["end"], target_local_name)
    replacement_element = enclosing(
        replacement_elements,
        replacement["start"],
        replacement["end"],
        target_local_name,
    )
    if source_element is None or replacement_element is None:
        return None, "outside_nonroot_element"
    if source_element.name != replacement_element.name:
        return None, "element_name_context_differs"
    source_bytes = source_element.end - source_element.start
    replacement_bytes = replacement_element.end - replacement_element.start
    flags: list[str] = []
    if source_element.namespace_marker or replacement_element.namespace_marker:
        flags.append("namespace_marker_in_window")
    if source_element.mc_marker or replacement_element.mc_marker:
        flags.append("markup_compatibility_marker_in_window")
    return {
        "source": span_dict(source_element.start, source_element.end),
        "replacement": span_dict(replacement_element.start, replacement_element.end),
        "name": source_element.name.decode("utf-8", "replace"),
        "depth": source_element.depth,
        "changed_source": source,
        "changed_replacement": replacement,
        "source_coverage_of_artifact": source_bytes,
        "replacement_coverage_of_artifact": replacement_bytes,
        "boundary_flags": flags,
        "lexical_balanced": True,
    }, None


def unique_candidates(
    changes: list[dict[str, Any]],
    source_elements: list[Element],
    replacement_elements: list[Element],
    target_local_name: bytes,
) -> tuple[list[dict[str, Any]], dict[str, int]]:
    candidates: list[dict[str, Any]] = []
    reasons: dict[str, int] = {}
    seen: set[tuple[int, int, int, int]] = set()
    for change in changes:
        candidate, reason = candidate_for_change(
            change,
            source_elements,
            replacement_elements,
            target_local_name,
        )
        if candidate is None:
            assert reason is not None
            reasons[reason] = reasons.get(reason, 0) + 1
            continue
        key = (
            candidate["source"]["start"], candidate["source"]["end"],
            candidate["replacement"]["start"], candidate["replacement"]["end"],
        )
        if key not in seen:
            candidates.append(candidate)
            seen.add(key)
    candidates.sort(key=lambda item: (item["replacement"]["start"], item["replacement"]["end"]))
    return candidates, reasons


def eligibility_summary(
    changes: list[dict[str, Any]],
    source_elements: list[Element],
    replacement_elements: list[Element],
    original: bytes,
    replacement: bytes,
    target_local_name: bytes,
) -> dict[str, Any]:
    candidates, ineligible_reasons = unique_candidates(
        changes,
        source_elements,
        replacement_elements,
        target_local_name,
    )
    source_changed = sum(item["source"]["bytes"] for item in changes)
    replacement_changed = sum(item["replacement"]["bytes"] for item in changes)
    candidate_source = union_bytes(item["source"] for item in candidates)
    candidate_replacement = union_bytes(item["replacement"] for item in candidates)
    return {
        "target_local_name": target_local_name.decode("ascii"),
        "candidate_count": len(candidates),
        "eligible_changed_groups": len(candidates),
        "ineligible_changed_groups": sum(ineligible_reasons.values()),
        "ineligible_reasons": ineligible_reasons,
        "candidates": candidates,
        "changed_source_bytes": source_changed,
        "changed_replacement_bytes": replacement_changed,
        "candidate_source_bytes_union": candidate_source,
        "candidate_replacement_bytes_union": candidate_replacement,
        "changed_source_coverage": candidate_source / source_changed if source_changed else 1.0,
        "changed_replacement_coverage": candidate_replacement / replacement_changed if replacement_changed else 1.0,
        "replacement_artifact_coverage": candidate_replacement / len(replacement),
        "replacement_bytes_avoided_if_all_candidates_proved": None,
        "replacement_bytes_avoided_model": "unavailable_without_exact_per-gap_identity_and_multi-window_proof",
        "candidate_is_heuristic": True,
        "semantic_proof_status": "not_attempted",
    }


def union_bytes(spans: Iterable[dict[str, int]]) -> int:
    ordered = sorted((item["start"], item["end"]) for item in spans)
    total = 0
    current: tuple[int, int] | None = None
    for start, end in ordered:
        if current is None:
            current = (start, end)
        elif start <= current[1]:
            current = (current[0], max(current[1], end))
        else:
            total += current[1] - current[0]
            current = (start, end)
    if current is not None:
        total += current[1] - current[0]
    return total


def token_count(tokens: list[Token], start: int = 0, end: int | None = None) -> int:
    if end is None:
        end = max((token.end for token in tokens), default=0)
    return sum(1 for token in tokens if start <= token.start and token.end <= end)


def parse_report(value: str) -> dict[str, Any]:
    fields = value.split(",")
    if not fields:
        return {"status": "unknown"}
    if fields[0] != "ok":
        return {"status": "error", "detail": value}
    result: dict[str, Any] = {"status": "ok"}
    for field in fields[1:]:
        key, _, raw = field.partition("=")
        result[key] = int(raw)
    return result


def parse_proof(value: str) -> dict[str, Any]:
    fields = value.split(",")
    kind = fields[0]
    result: dict[str, Any] = {"kind": kind}
    if kind == "Window" and len(fields) >= 3:
        result["source"] = {"start": int(fields[1].split("..", 1)[0]), "end": int(fields[1].split("..", 1)[1])}
        result["replacement"] = {"start": int(fields[2].split("..", 1)[0]), "end": int(fields[2].split("..", 1)[1])}
        result["source"]["bytes"] = result["source"]["end"] - result["source"]["start"]
        result["replacement"]["bytes"] = result["replacement"]["end"] - result["replacement"]["start"]
    elif kind == "Error":
        result["detail"] = ",".join(fields[1:])
    return result


def read_probe(path: Path) -> dict[tuple[str, str], dict[str, Any]]:
    rows: dict[tuple[str, str], dict[str, Any]] = {}
    lines = path.read_text(encoding="utf-8").splitlines()
    if not lines or lines[0] != "case\tmember\toriginal_bytes\treplacement_bytes\toriginal_report\treplacement_report\tproof":
        raise ValueError("unexpected probe header")
    for line in lines[1:]:
        fields = line.split("\t")
        if len(fields) != 7:
            raise ValueError(f"unexpected probe row: {line!r}")
        case, member, original_bytes, replacement_bytes, original_report, replacement_report, proof = fields
        rows[(case, member)] = {
            "original_bytes": int(original_bytes),
            "replacement_bytes": int(replacement_bytes),
            "original_report": parse_report(original_report),
            "replacement_report": parse_report(replacement_report),
            "proof": parse_proof(proof),
        }
    return rows


def manifest_source_hash(manifest: dict[str, Any]) -> str:
    encoded = json.dumps(manifest, indent=2, sort_keys=False).encode() + b"\n"
    return sha256_bytes(encoded)


def selected_parts(manifest: dict[str, Any]) -> Iterable[tuple[str, dict[str, Any], str]]:
    for case in manifest["cases"]:
        for part in case["parts"]:
            yield case["id"], part, case["id"].replace("-one-percent", "/one-percent").replace("-one-edit", "/one-edit")


def run() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--artifacts", type=Path, required=True)
    parser.add_argument("--manifest", type=Path, required=True)
    parser.add_argument("--probe", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    if args.output.exists():
        raise SystemExit(f"refusing existing output: {args.output}")

    manifest_bytes = args.manifest.read_bytes()
    manifest = json.loads(manifest_bytes)
    probe = read_probe(args.probe)
    records: list[dict[str, Any]] = []
    aggregate = {
        "parts": 0,
        "proof_kinds": {"Identical": 0, "Window": 0, "Complete": 0, "Error": 0},
        "changed_groups": 0,
        "lexically_eligible_groups": 0,
        "lexically_ineligible_groups": 0,
        "source_changed_bytes": 0,
        "replacement_changed_bytes": 0,
        "lexical_candidate_source_bytes": 0,
        "lexical_candidate_replacement_bytes": 0,
        "cell_candidate_source_bytes": 0,
        "cell_candidate_replacement_bytes": 0,
        "full_replacement_bytes": 0,
        "actual_window_replacement_bytes": 0,
        "actual_window_events": 0,
        "actual_window_parts": 0,
    }

    for case_id, part, scenario in selected_parts(manifest):
        member = part["member"]
        original_path = args.artifacts / scenario / "original" / member
        replacement_path = args.artifacts / scenario / "replacement" / member
        original = original_path.read_bytes()
        replacement = replacement_path.read_bytes()
        expected_original = part["original"]
        expected_replacement = part["replacement"]
        if len(original) != expected_original["bytes"] or sha256_bytes(original) != expected_original["sha256"]:
            raise ValueError(f"original identity mismatch: {case_id} {member}")
        if len(replacement) != expected_replacement["bytes"] or sha256_bytes(replacement) != expected_replacement["sha256"]:
            raise ValueError(f"replacement identity mismatch: {case_id} {member}")
        row = probe[(case_id, member)]
        if row["original_bytes"] != len(original) or row["replacement_bytes"] != len(replacement):
            raise ValueError(f"probe length mismatch: {case_id} {member}")
        expected_proof = part["proof"]
        actual_proof = row["proof"]
        if expected_proof["kind"] != actual_proof["kind"]:
            raise ValueError(f"proof kind mismatch: {case_id} {member}: {expected_proof} vs {actual_proof}")
        if expected_proof["kind"] == "Window":
            expected_source = tuple(expected_proof["original"])
            expected_replacement = tuple(expected_proof["replacement"])
            actual_source = (actual_proof["source"]["start"], actual_proof["source"]["end"])
            actual_replacement = (actual_proof["replacement"]["start"], actual_proof["replacement"]["end"])
            if expected_source != actual_source or expected_replacement != actual_replacement:
                raise ValueError(f"proof range mismatch: {case_id} {member}")

        source_tokens, source_elements, source_lexical_reasons = scan_lexically(original)
        replacement_tokens, replacement_elements, replacement_lexical_reasons = scan_lexically(replacement)
        changes, change_method = payload_difference(
            original,
            replacement,
            source_elements,
            replacement_elements,
        )
        source_changed = sum(item["source"]["bytes"] for item in changes)
        replacement_changed = sum(item["replacement"]["bytes"] for item in changes)
        common_prefix, common_suffix = common_prefix_suffix(original, replacement)
        value_eligibility = eligibility_summary(
            changes,
            source_elements,
            replacement_elements,
            original,
            replacement,
            b"v",
        )
        cell_eligibility = eligibility_summary(
            changes,
            source_elements,
            replacement_elements,
            original,
            replacement,
            b"c",
        )
        actual_window = None
        actual_window_events = None
        actual_window_lexically_balanced = None
        if actual_proof["kind"] == "Window":
            actual_window = {
                "source": actual_proof["source"],
                "replacement": actual_proof["replacement"],
                "source_bytes_of_artifact": actual_proof["source"]["bytes"] / len(original),
                "replacement_bytes_of_artifact": actual_proof["replacement"]["bytes"] / len(replacement),
                "replacement_bytes_avoided": len(replacement) - actual_proof["replacement"]["bytes"],
                "changed_source_bytes_covered": sum(
                    item["source"]["bytes"] for item in changes
                    if actual_proof["source"]["start"] <= item["source"]["start"]
                    and item["source"]["end"] <= actual_proof["source"]["end"]
                ),
                "changed_replacement_bytes_covered": sum(
                    item["replacement"]["bytes"] for item in changes
                    if actual_proof["replacement"]["start"] <= item["replacement"]["start"]
                    and item["replacement"]["end"] <= actual_proof["replacement"]["end"]
                ),
            }
            actual_window_events = token_count(
                replacement_tokens,
                actual_proof["replacement"]["start"],
                actual_proof["replacement"]["end"],
            )
            actual_window["lexical_events"] = actual_window_events
            actual_window_lexically_balanced = any(
                e.start == actual_proof["source"]["start"]
                and e.end == actual_proof["source"]["end"]
                for e in source_elements
            )

        record = {
            "case": case_id,
            "member": member,
            "original": {
                "path": str(original_path),
                "bytes": len(original),
                "sha256": sha256_bytes(original),
                "report": row["original_report"],
                "lexical_tokens": len(source_tokens),
                "lexical_elements": len(source_elements),
                "lexical_reasons": source_lexical_reasons,
            },
            "replacement": {
                "path": str(replacement_path),
                "bytes": len(replacement),
                "sha256": sha256_bytes(replacement),
                "report": row["replacement_report"],
                "lexical_tokens": len(replacement_tokens),
                "lexical_elements": len(replacement_elements),
                "lexical_reasons": replacement_lexical_reasons,
            },
            "common_prefix_bytes": common_prefix,
            "common_suffix_bytes": common_suffix,
            "changed_spans": changes,
            "changed_span_method": change_method,
            "audit_proof": {"expected": expected_proof, "observed": actual_proof},
            "actual_window": actual_window,
            "actual_window_lexically_balanced": actual_window_lexically_balanced,
            "lexical_multi_window_eligibility": value_eligibility,
            "cell_window_eligibility": cell_eligibility,
        }
        records.append(record)
        aggregate["parts"] += 1
        aggregate["proof_kinds"][actual_proof["kind"]] += 1
        aggregate["changed_groups"] += len(changes)
        aggregate["lexically_eligible_groups"] += value_eligibility["candidate_count"]
        aggregate["lexically_ineligible_groups"] += value_eligibility["ineligible_changed_groups"]
        aggregate["source_changed_bytes"] += source_changed
        aggregate["replacement_changed_bytes"] += replacement_changed
        aggregate["lexical_candidate_source_bytes"] += value_eligibility["candidate_source_bytes_union"]
        aggregate["lexical_candidate_replacement_bytes"] += value_eligibility["candidate_replacement_bytes_union"]
        aggregate["cell_candidate_source_bytes"] += cell_eligibility["candidate_source_bytes_union"]
        aggregate["cell_candidate_replacement_bytes"] += cell_eligibility["candidate_replacement_bytes_union"]
        aggregate["full_replacement_bytes"] += len(replacement)
        if actual_window is not None:
            aggregate["actual_window_parts"] += 1
            aggregate["actual_window_replacement_bytes"] += actual_window["replacement"]["bytes"]
            aggregate["actual_window_events"] += actual_window_events or 0

    output = {
        "schema": SCHEMA,
        "scope": {
            "artifact_manifest": str(args.manifest),
            "artifact_manifest_sha256": sha256_bytes(manifest_bytes),
            "probe_tsv_sha256": sha256_bytes(args.probe.read_bytes()),
            "limits": "xml_minifier::audit::Limits::default()",
            "proof_authority": "Rust verify_source_replacement from the current xml-minifier source",
            "lexical_authority": "diagnostic only; does not establish a publication proof",
            "multi_window_status": "heuristic eligibility only; semantic proof not implemented",
        },
        "aggregate": aggregate,
        "records": records,
    }
    args.output.write_text(json.dumps(output, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(run())
    except (OSError, ValueError, KeyError, IndexError) as error:
        print(f"error: {error}", file=sys.stderr)
        raise SystemExit(1)
