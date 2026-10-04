#!/usr/bin/env python3
"""Census selected XLSX XML fixtures without invoking the Rust auditor.

The exporter writes a manifest whose paths are relative to the manifest's
directory.  This helper verifies every manifest hash, verifies that each
selected XML sidecar is the corresponding member of its source or output ZIP,
and then parses the complete sidecar bytes with CPython's Expat parser.

The parser deliberately leaves namespace processing disabled.  Consequently
``xmlns`` declarations remain ordinary raw attributes, as they do in the
lexical input seen by the XML auditor.  ``ordered_attributes`` is enabled so
the callback receives the lexical attribute list; this census only uses its
cardinality and the raw ``xml:space`` name.

The output is structural evidence about the selected full XML artifacts.  It
does not replay the Rust auditor, does not re-audit a Window proof span as an
independent XML document, and makes no native performance claim.  Window span
bytes are therefore recorded as ranges and hashes only.
"""

from __future__ import annotations

import argparse
import hashlib
import io
import json
import os
from pathlib import Path
import sys
from typing import Any, Iterable
import zipfile
from xml.parsers import expat


INPUT_SCHEMA = "litchi.xlsx-attribute-fixtures.v1"
OUTPUT_SCHEMA = "litchi.xlsx-attribute-census.v1"
PROOF_KINDS = frozenset({"Complete", "Identical", "Window"})
CARDINALITY_KEYS = ("0", "1", "2..32", ">32")


class CensusError(Exception):
    """A manifest, archive, or XML input failed a fail-closed check."""


def _where(value: Any) -> str:
    return repr(value)


def _require_object(value: Any, label: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise CensusError(f"{label} must be an object, got {_where(value)}")
    return value


def _require_string(value: Any, label: str) -> str:
    if not isinstance(value, str) or not value:
        raise CensusError(f"{label} must be a non-empty string")
    return value


def _require_nonnegative_int(value: Any, label: str) -> int:
    # bool is an int subclass, but it is not a byte count or an offset.
    if isinstance(value, bool) or not isinstance(value, int) or value < 0:
        raise CensusError(f"{label} must be a non-negative integer")
    return value


def _require_sha256(value: Any, label: str) -> str:
    value = _require_string(value, label)
    if len(value) != hashlib.sha256().digest_size * 2:
        raise CensusError(f"{label} must be a 64-character SHA-256 digest")
    try:
        int(value, 16)
    except ValueError as exc:
        raise CensusError(f"{label} must be hexadecimal") from exc
    return value.lower()


def _require_keys(value: dict[str, Any], required: Iterable[str], label: str) -> None:
    missing = sorted(set(required).difference(value))
    if missing:
        raise CensusError(f"{label} is missing required keys: {', '.join(missing)}")


def _hash_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _file_descriptor(value: Any, label: str) -> dict[str, Any]:
    value = _require_object(value, label)
    _require_keys(value, ("path", "bytes", "sha256"), label)
    path = _require_string(value["path"], f"{label}.path")
    if Path(path).is_absolute():
        raise CensusError(f"{label}.path must be relative to the manifest")
    return {
        "path": path,
        "bytes": _require_nonnegative_int(value["bytes"], f"{label}.bytes"),
        "sha256": _require_sha256(value["sha256"], f"{label}.sha256"),
    }


def _safe_relative_path(base: Path, relative: str, label: str) -> Path:
    candidate = (base / relative).resolve()
    try:
        candidate.relative_to(base.resolve())
    except ValueError as exc:
        raise CensusError(f"{label} escapes the manifest directory: {relative!r}") from exc
    if not candidate.is_file():
        raise CensusError(f"{label} does not name a regular file: {relative!r}")
    return candidate


def _read_verified_file(
    base: Path,
    descriptor: dict[str, Any],
    label: str,
) -> tuple[bytes, dict[str, Any]]:
    path = _safe_relative_path(base, descriptor["path"], f"{label}.path")
    data = path.read_bytes()
    actual_bytes = len(data)
    actual_sha256 = _hash_bytes(data)
    if actual_bytes != descriptor["bytes"]:
        raise CensusError(
            f"{label} byte count mismatch for {descriptor['path']!r}: "
            f"manifest {descriptor['bytes']}, actual {actual_bytes}"
        )
    if actual_sha256 != descriptor["sha256"]:
        raise CensusError(
            f"{label} SHA-256 mismatch for {descriptor['path']!r}: "
            f"manifest {descriptor['sha256']}, actual {actual_sha256}"
        )
    return data, dict(descriptor)


def _archive_member(archive: bytes, member: str, label: str) -> bytes:
    try:
        with zipfile.ZipFile(io.BytesIO(archive)) as zipped:
            names = [info.filename for info in zipped.infolist()]
            if len(names) != len(set(names)):
                raise CensusError(f"{label} contains duplicate ZIP member names")
            try:
                info = zipped.getinfo(member)
            except KeyError as exc:
                raise CensusError(f"{label} is missing selected member {member!r}") from exc
            if info.is_dir():
                raise CensusError(f"{label} selected member {member!r} is a directory")
            return zipped.read(info)
    except CensusError:
        raise
    except (OSError, zipfile.BadZipFile, zipfile.LargeZipFile) as exc:
        raise CensusError(f"{label} is not a readable ZIP archive: {exc}") from exc


def _span(value: Any, length: int, label: str) -> tuple[int, int]:
    if not isinstance(value, list) or len(value) != 2:
        raise CensusError(f"{label} must be a two-item array")
    start = _require_nonnegative_int(value[0], f"{label}[0]")
    end = _require_nonnegative_int(value[1], f"{label}[1]")
    if start > end or end > length:
        raise CensusError(
            f"{label} {value!r} is outside the corresponding artifact of {length} bytes"
        )
    return start, end


def _proof_descriptor(
    value: Any,
    original: bytes,
    replacement: bytes,
    label: str,
) -> dict[str, Any]:
    proof = _require_object(value, label)
    _require_keys(proof, ("kind",), label)
    kind = _require_string(proof["kind"], f"{label}.kind")
    if kind not in PROOF_KINDS:
        raise CensusError(f"{label}.kind must be Complete, Identical, or Window")

    if kind == "Identical":
        if "original" in proof or "replacement" in proof:
            raise CensusError(f"{label} Identical proof must not carry spans")
        if original != replacement:
            raise CensusError(f"{label} Identical proof has different original/replacement bytes")
        return {"kind": kind, "bytes_equal": True, "full_replacement_audit": False}

    if kind == "Complete":
        if "original" in proof or "replacement" in proof:
            raise CensusError(f"{label} Complete proof must not carry spans")
        return {"kind": kind, "bytes_equal": original == replacement, "full_replacement_audit": True}

    _require_keys(proof, ("original", "replacement"), label)
    original_span = _span(proof["original"], len(original), f"{label}.original")
    replacement_span = _span(proof["replacement"], len(replacement), f"{label}.replacement")

    def span_descriptor(data: bytes, bounds: tuple[int, int]) -> dict[str, Any]:
        start, end = bounds
        span_bytes = data[start:end]
        return {
            "range": [start, end],
            "bytes": len(span_bytes),
            "sha256": _hash_bytes(span_bytes),
        }

    return {
        "kind": kind,
        "bytes_equal": original == replacement,
        "full_replacement_audit": False,
        "original_span": span_descriptor(original, original_span),
        "replacement_span": span_descriptor(replacement, replacement_span),
        "span_parsed_as_xml": False,
    }


def _empty_metrics() -> dict[str, Any]:
    return {
        "start_element_events": 0,
        "raw_attribute_cardinality": {key: 0 for key in CARDINALITY_KEYS},
        "attributes": 0,
        "xml_space_present_events": 0,
        "xml_space_absent_events": 0,
        "xml_space_attribute_occurrences": 0,
        "eligible_zero_without_xml_space": 0,
        "eligible_one_without_xml_space": 0,
        "eligible_zero_or_one_without_xml_space": 0,
    }


def _validate_metrics(metrics: dict[str, Any], label: str) -> None:
    histogram = metrics["raw_attribute_cardinality"]
    histogram_total = sum(histogram.values())
    starts = metrics["start_element_events"]
    if histogram_total != starts:
        raise CensusError(f"{label} cardinality total {histogram_total} != starts {starts}")
    if metrics["attributes"] < histogram["1"]:
        raise CensusError(f"{label} attribute total is below one-attribute event count")
    xml_space_total = (
        metrics["xml_space_present_events"] + metrics["xml_space_absent_events"]
    )
    if xml_space_total != starts:
        raise CensusError(f"{label} xml:space total {xml_space_total} != starts {starts}")
    eligible = (
        metrics["eligible_zero_without_xml_space"]
        + metrics["eligible_one_without_xml_space"]
    )
    if eligible != metrics["eligible_zero_or_one_without_xml_space"]:
        raise CensusError(f"{label} zero/one eligibility does not reconcile")
    if eligible > metrics["xml_space_absent_events"]:
        raise CensusError(f"{label} eligible events exceed xml:space-absent events")
    if metrics["xml_space_attribute_occurrences"] < metrics["xml_space_present_events"]:
        raise CensusError(f"{label} xml:space attribute occurrences do not reconcile")


def _add_metrics(left: dict[str, Any], right: dict[str, Any], label: str) -> dict[str, Any]:
    result = _empty_metrics()
    result["start_element_events"] = left["start_element_events"] + right["start_element_events"]
    result["attributes"] = left["attributes"] + right["attributes"]
    for key in CARDINALITY_KEYS:
        result["raw_attribute_cardinality"][key] = (
            left["raw_attribute_cardinality"][key]
            + right["raw_attribute_cardinality"][key]
        )
    for key in (
        "xml_space_present_events",
        "xml_space_absent_events",
        "xml_space_attribute_occurrences",
        "eligible_zero_without_xml_space",
        "eligible_one_without_xml_space",
        "eligible_zero_or_one_without_xml_space",
    ):
        result[key] = left[key] + right[key]
    _validate_metrics(result, label)
    return result


def _census_xml(data: bytes, label: str) -> dict[str, Any]:
    metrics = _empty_metrics()
    parser = expat.ParserCreate(encoding=None, namespace_separator=None)
    try:
        parser.ordered_attributes = True
    except (AttributeError, TypeError) as exc:
        raise CensusError("this Python Expat binding cannot enable ordered_attributes") from exc

    def reject_construct(*_args: Any) -> None:
        raise CensusError(f"{label} contains a DTD or entity declaration")

    parser.StartDoctypeDeclHandler = reject_construct
    parser.EndDoctypeDeclHandler = reject_construct
    parser.EntityDeclHandler = reject_construct
    parser.ExternalEntityRefHandler = reject_construct
    parser.UnparsedEntityDeclHandler = reject_construct
    parser.NotationDeclHandler = reject_construct
    parser.ElementDeclHandler = reject_construct
    parser.AttlistDeclHandler = reject_construct
    try:
        parser.SetParamEntityParsing(expat.XML_PARAM_ENTITY_PARSING_NEVER)
    except (AttributeError, expat.error) as exc:
        raise CensusError("this Python Expat binding cannot disable parameter entities") from exc

    def on_start(_name: str, attributes: list[str]) -> None:
        if len(attributes) % 2:
            raise CensusError(f"{label} returned an odd ordered attribute list")
        count = len(attributes) // 2
        if count == 0:
            bucket = "0"
        elif count == 1:
            bucket = "1"
        elif count <= 32:
            bucket = "2..32"
        else:
            bucket = ">32"
        metrics["start_element_events"] += 1
        metrics["attributes"] += count
        metrics["raw_attribute_cardinality"][bucket] += 1

        xml_space_count = sum(
            1 for index in range(0, len(attributes), 2) if attributes[index] == "xml:space"
        )
        metrics["xml_space_attribute_occurrences"] += xml_space_count
        if xml_space_count:
            metrics["xml_space_present_events"] += 1
        else:
            metrics["xml_space_absent_events"] += 1
            if count == 0:
                metrics["eligible_zero_without_xml_space"] += 1
            elif count == 1:
                metrics["eligible_one_without_xml_space"] += 1
            if count <= 1:
                metrics["eligible_zero_or_one_without_xml_space"] += 1

    parser.StartElementHandler = on_start
    try:
        parser.Parse(data, True)
    except CensusError:
        raise
    except expat.ExpatError as exc:
        raise CensusError(
            f"{label} is not well-formed according to stdlib Expat at byte "
            f"{getattr(exc, 'offset', '?')}: {exc}"
        ) from exc
    _validate_metrics(metrics, label)
    return metrics


def _manifest_descriptor(path: Path, data: bytes) -> dict[str, Any]:
    return {"path": path.name, "bytes": len(data), "sha256": _hash_bytes(data)}


def _crosscheck_exporter_census(
    value: Any,
    metrics: dict[str, Any],
    label: str,
) -> dict[str, Any]:
    """Compare the exporter-side quick-xml census with independent Expat counts."""
    census = _require_object(value, label)
    required = (
        "start_or_empty_tags",
        "zero_attribute_tags",
        "one_attribute_tags",
        "multiple_attribute_tags",
        "attributes",
        "xml_space_attributes",
        "xml_space_elements",
    )
    _require_keys(census, required, label)
    expected = {
        "start_or_empty_tags": metrics["start_element_events"],
        "zero_attribute_tags": metrics["raw_attribute_cardinality"]["0"],
        "one_attribute_tags": metrics["raw_attribute_cardinality"]["1"],
        "multiple_attribute_tags": (
            metrics["raw_attribute_cardinality"]["2..32"]
            + metrics["raw_attribute_cardinality"][">32"]
        ),
        "attributes": metrics["attributes"],
        "xml_space_attributes": metrics["xml_space_attribute_occurrences"],
        "xml_space_elements": metrics["xml_space_present_events"],
    }
    normalized: dict[str, Any] = {}
    for key in required:
        normalized[key] = _require_nonnegative_int(census[key], f"{label}.{key}")
        if normalized[key] != expected[key]:
            raise CensusError(
                f"{label}.{key} mismatch: exporter {normalized[key]}, Expat {expected[key]}"
            )
    return normalized


def _case_result(base: Path, case: dict[str, Any], index: int) -> dict[str, Any]:
    label = f"cases[{index}]"
    _require_keys(case, ("id", "shape", "case", "source", "output", "parts"), label)
    case_id = _require_string(case["id"], f"{label}.id")
    shape = _require_string(case["shape"], f"{label}.shape")
    case_name = _require_string(case["case"], f"{label}.case")

    source_desc = _file_descriptor(case["source"], f"{label}.source")
    output_desc = _file_descriptor(case["output"], f"{label}.output")
    source_archive, source_bound = _read_verified_file(base, source_desc, f"{label}.source")
    output_archive, output_bound = _read_verified_file(base, output_desc, f"{label}.output")

    parts = case["parts"]
    if not isinstance(parts, list) or not parts:
        raise CensusError(f"{label}.parts must be a non-empty array")
    seen_members: set[str] = set()
    source_total = _empty_metrics()
    replacement_total = _empty_metrics()
    part_results: list[dict[str, Any]] = []

    for part_index, raw_part in enumerate(parts):
        part_label = f"{label}.parts[{part_index}]"
        part = _require_object(raw_part, part_label)
        _require_keys(
            part,
            ("member", "original", "replacement", "original_census", "replacement_census", "proof"),
            part_label,
        )
        member = _require_string(part["member"], f"{part_label}.member")
        if member in seen_members:
            raise CensusError(f"{label} lists selected member {member!r} more than once")
        seen_members.add(member)
        if not member.lower().endswith(".xml"):
            raise CensusError(f"{part_label}.member is not an XML member: {member!r}")

        original_desc = _file_descriptor(part["original"], f"{part_label}.original")
        replacement_desc = _file_descriptor(part["replacement"], f"{part_label}.replacement")
        original, original_bound = _read_verified_file(
            base, original_desc, f"{part_label}.original"
        )
        replacement, replacement_bound = _read_verified_file(
            base, replacement_desc, f"{part_label}.replacement"
        )

        archive_original = _archive_member(source_archive, member, f"{label}.source")
        archive_replacement = _archive_member(output_archive, member, f"{label}.output")
        if archive_original != original:
            raise CensusError(
                f"{part_label}.original does not match source archive member {member!r}"
            )
        if archive_replacement != replacement:
            raise CensusError(
                f"{part_label}.replacement does not match output archive member {member!r}"
            )

        original_metrics = _census_xml(original, f"{part_label}.original")
        replacement_metrics = _census_xml(replacement, f"{part_label}.replacement")
        original_exporter_census = _crosscheck_exporter_census(
            part["original_census"], original_metrics, f"{part_label}.original_census"
        )
        replacement_exporter_census = _crosscheck_exporter_census(
            part["replacement_census"],
            replacement_metrics,
            f"{part_label}.replacement_census",
        )
        proof = _proof_descriptor(
            part["proof"], original, replacement, f"{part_label}.proof"
        )
        source_total = _add_metrics(source_total, original_metrics, f"{label}.source_total")
        replacement_total = _add_metrics(
            replacement_total, replacement_metrics, f"{label}.replacement_total"
        )
        part_results.append(
            {
                "member": member,
                "original": {
                    "artifact": original_bound,
                    "exporter_census": original_exporter_census,
                    "full_xml_census": original_metrics,
                },
                "replacement": {
                    "artifact": replacement_bound,
                    "exporter_census": replacement_exporter_census,
                    "full_xml_census": replacement_metrics,
                },
                "proof": proof,
                "claims": {
                    "full_original_structural_census": True,
                    "full_replacement_structural_census": True,
                    "replacement_full_audit_coverage": proof["full_replacement_audit"],
                    "native_benefit": False,
                },
            }
        )

    _validate_metrics(source_total, f"{label}.source_total")
    _validate_metrics(replacement_total, f"{label}.replacement_total")
    return {
        "id": case_id,
        "shape": shape,
        "case": case_name,
        "source_archive": source_bound,
        "output_archive": output_bound,
        "parts": part_results,
        "totals": {
            "original": source_total,
            "replacement": replacement_total,
        },
        "claims": {
            "selected_parts_only": True,
            "full_original_structural_census": True,
            "full_replacement_structural_census": True,
            "native_benefit": False,
        },
    }


def _load_manifest(manifest_path: Path) -> tuple[bytes, dict[str, Any]]:
    data = manifest_path.read_bytes()
    try:
        manifest = json.loads(data)
    except json.JSONDecodeError as exc:
        raise CensusError(f"manifest is not valid JSON: {exc}") from exc
    manifest = _require_object(manifest, "manifest")
    if manifest.get("schema") != INPUT_SCHEMA:
        raise CensusError(
            f"manifest.schema must be {INPUT_SCHEMA!r}, got {manifest.get('schema')!r}"
        )
    _require_keys(manifest, ("schema", "cases"), "manifest")
    if not isinstance(manifest["cases"], list) or not manifest["cases"]:
        raise CensusError("manifest.cases must be a non-empty array")
    return data, manifest


def build_result(input_dir: Path, manifest_path: Path) -> dict[str, Any]:
    input_dir = input_dir.resolve()
    manifest_path = manifest_path.resolve()
    try:
        manifest_path.relative_to(input_dir)
    except ValueError as exc:
        raise CensusError("manifest must be inside input_dir") from exc
    manifest_bytes, manifest = _load_manifest(manifest_path)
    manifest_bound = _manifest_descriptor(manifest_path, manifest_bytes)
    helper_path = Path(__file__).resolve()
    helper_bytes = helper_path.read_bytes()

    seen_ids: set[str] = set()
    case_results: list[dict[str, Any]] = []
    original_total = _empty_metrics()
    replacement_total = _empty_metrics()
    for index, raw_case in enumerate(manifest["cases"]):
        case = _require_object(raw_case, f"manifest.cases[{index}]")
        case_result = _case_result(input_dir, case, index)
        if case_result["id"] in seen_ids:
            raise CensusError(f"duplicate case id {case_result['id']!r}")
        seen_ids.add(case_result["id"])
        case_results.append(case_result)
        original_total = _add_metrics(
            original_total, case_result["totals"]["original"], "totals.original"
        )
        replacement_total = _add_metrics(
            replacement_total, case_result["totals"]["replacement"], "totals.replacement"
        )

    return {
        "schema": OUTPUT_SCHEMA,
        "input_schema": INPUT_SCHEMA,
        "manifest": manifest_bound,
        "helper": {
            "path": helper_path.name,
            "bytes": len(helper_bytes),
            "sha256": _hash_bytes(helper_bytes),
        },
        "scope": {
            "selected_parts_only": True,
            "full_xml_artifacts_parsed": True,
            "namespace_processing": "disabled",
            "ordered_attributes": True,
            "raw_attribute_cardinality": ["0", "1", "2..32", ">32"],
            "proof_windows_parsed_as_xml": False,
            "rust_auditor_replayed": False,
            "native_benefit_claim": False,
        },
        "cases": case_results,
        "totals": {
            "case_count": len(case_results),
            "part_count": sum(len(case["parts"]) for case in case_results),
            "original": original_total,
            "replacement": replacement_total,
        },
    }


def _write_once(path: Path, result: dict[str, Any]) -> None:
    path = path.resolve()
    path.parent.mkdir(parents=True, exist_ok=True)
    encoded = (json.dumps(result, ensure_ascii=True, indent=2, sort_keys=True) + "\n").encode()
    try:
        fd = os.open(path, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o644)
    except FileExistsError as exc:
        raise CensusError(f"refusing to overwrite existing output: {path}") from exc
    try:
        with os.fdopen(fd, "wb") as stream:
            stream.write(encoded)
            stream.flush()
            os.fsync(stream.fileno())
    except BaseException:
        try:
            path.unlink()
        except OSError:
            pass
        raise


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "input_dir",
        type=Path,
        help="directory containing manifest.json and exporter artifacts",
    )
    parser.add_argument(
        "--manifest",
        type=Path,
        help="manifest path, relative to input_dir (default: manifest.json)",
    )
    parser.add_argument("--output", type=Path, required=True, help="new JSON output path")
    args = parser.parse_args(argv)
    input_dir = args.input_dir.resolve()
    if not input_dir.is_dir():
        parser.error(f"input_dir is not a directory: {input_dir}")
    manifest_path = args.manifest or (input_dir / "manifest.json")
    if not manifest_path.is_absolute():
        manifest_path = input_dir / manifest_path
    try:
        result = build_result(input_dir, manifest_path)
        _write_once(args.output, result)
    except (CensusError, OSError) as exc:
        print(f"census.py: error: {exc}", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
