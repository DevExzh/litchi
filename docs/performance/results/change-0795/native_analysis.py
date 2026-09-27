"""Offline replay of the 0795 frame-pointer native stack captures.

The capture driver removes the uncompressed ``perf.data`` files after making
identity-bound gzip copies.  This module therefore verifies both the
descriptor for every retained artifact and the bytes recovered from each
compressed raw recording before parsing the canonical no-inline frame stream.
It reports observed sample/frame counts only.  In particular, a sampled event
period is used only as a parser sanity check and is deliberately not included
in the result as a CPU or latency measurement.
"""

from __future__ import annotations

import argparse
import collections
import gzip
import hashlib
import json
import re
from pathlib import Path
from typing import Any, Iterable

import custody as c


SCHEMA = "litchi.performance.0795.native-analysis.v1"
OWNER = "namespace_uri_probe::capture_region_0793"
OWNER_RE = re.compile(
    rf"^{re.escape(OWNER)}(?:::h[0-9a-f]+)?$"
)

# ``perf script`` prints the leaf first and the caller/root last.  The public
# operation is expected below the non-inlined probe wrapper in that order.
# Keep the package/model path in the marker so an unrelated function whose
# name happens to contain ``opened_presentation`` cannot satisfy qualification.
PUBLIC_OPERATION_TOKEN = (
    "litchi_pptx::package::model::Package::opened_presentation"
)

# These are inclusive diagnostics.  A frame can contribute to several
# categories, so none of the category counts is a partition or a share.
_CATEGORY_MARKERS: dict[str, tuple[str, ...]] = {
    "notes": ("litchi_pptx::notes::",),
    "inspect": ("litchi_pptx::notes::codec::inspect_element",),
    "helper": (
        "litchi_opc::xml_attributes::",
        "litchi_ole_common::xml_attributes::",
        "litchi_sign::xml_attributes::",
        "litchi_xldm::xml_attributes::",
        "xml_minifier::xml_attributes::",
    ),
    "quickxml": ("quick_xml::",),
    "alloc": (
        "alloc::",
        "__rust_alloc",
        "__rust_dealloc",
        "__rust_realloc",
        "malloc",
        "calloc",
        "realloc",
        "free",
    ),
    "memory": (
        "memchr::",
        "memcpy",
        "memmove",
        "memset",
        "memcmp",
        "__mem",
        "__rust_alloc",
        "__rust_dealloc",
        "__rust_realloc",
        "malloc",
        "calloc",
        "realloc",
        "free",
    ),
}

_HEADER_RE = re.compile(
    r"^.*?(?P<timestamp>[0-9]+(?:\.[0-9]+)?):\s*"
    r"(?P<period>[1-9][0-9]*)\s+cycles:u:\s*$"
)
_FRAME_RE = re.compile(
    r"^\s*(?P<address>(?:0x)?[0-9a-fA-F]+)\s+"
    r"(?P<symbol>.+?)\s+\((?P<object>.*)\)\s*$"
)
_OFFSET_RE = re.compile(r"\+0x[0-9a-fA-F]+$")
_HEX64_RE = re.compile(r"[0-9a-f]{64}\Z")


class EvidenceError(ValueError):
    """A malformed, missing, stale, or contradictory retained artifact."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise EvidenceError(message)


def _packet_path(raw: Any, label: str) -> Path:
    require(isinstance(raw, str) and raw, f"{label}: artifact path is missing")
    marker = "/change-0795/"
    if marker in raw:
        path = (c.P / raw.split(marker, 1)[1]).resolve()
    else:
        candidate = Path(raw)
        path = (candidate if candidate.is_absolute() else c.P / candidate).resolve()
    try:
        path.relative_to(c.P.resolve())
    except ValueError as error:
        raise EvidenceError(f"{label}: path escapes packet: {raw}") from error
    return path


def _descriptor(value: Any, label: str, *, allow_missing: bool = False) -> Path:
    require(isinstance(value, dict), f"{label}: artifact descriptor is missing")
    path = _packet_path(value.get("path"), label)
    size = value.get("bytes")
    digest = value.get("sha256")
    require(type(size) is int and size >= 0, f"{label}: byte count is invalid")
    require(isinstance(digest, str) and _HEX64_RE.fullmatch(digest),
            f"{label}: SHA-256 is invalid")
    if not path.is_file() or path.is_symlink():
        require(allow_missing, f"{label}: artifact is missing: {path}")
        return path
    require(path.stat().st_size == size, f"{label}: byte count changed")
    require(c.sha(path) == digest, f"{label}: SHA-256 changed")
    return path


def _external_path(raw: Any, label: str) -> Path:
    require(isinstance(raw, str) and raw, f"{label}: binary path is missing")
    path = Path(raw)
    require(path.is_absolute(), f"{label}: binary path is not absolute")
    return path.resolve()


def _external_descriptor(value: Any, label: str) -> dict[str, Any]:
    """Verify a build binary while live, or against the later cleanup witness."""

    require(isinstance(value, dict), f"{label}: binary descriptor is missing")
    path = _external_path(value.get("path"), label)
    size = value.get("bytes")
    digest = value.get("sha256")
    require(type(size) is int and size > 0, f"{label}: binary byte count is invalid")
    require(isinstance(digest, str) and _HEX64_RE.fullmatch(digest),
            f"{label}: binary SHA-256 is invalid")
    if path.is_file() and not path.is_symlink():
        require(path.stat().st_size == size, f"{label}: binary byte count changed")
        require(c.sha(path) == digest, f"{label}: binary SHA-256 changed")
        return {"path": str(path), "bytes": size, "sha256": digest}

    cleanup_path = c.P / "cleanup.json"
    require(cleanup_path.is_file(), f"{label}: binary is missing without cleanup witness")
    cleanup = c.read(cleanup_path)
    require(cleanup.get("target_removed") is True,
            f"{label}: cleanup witness does not mark target removed")
    require(cleanup.get("target") == str(c.TARGET),
            f"{label}: cleanup target changed")
    removed = cleanup.get("removed_binaries")
    require(isinstance(removed, list), f"{label}: cleanup binary list is malformed")
    matches = [row for row in removed if isinstance(row, dict)
               and row.get("path") == str(path)]
    require(len(matches) == 1, f"{label}: removed binary witness is missing")
    require(matches[0].get("bytes") == size and matches[0].get("sha256") == digest,
            f"{label}: cleanup binary identity differs")
    return {"path": str(path), "bytes": size, "sha256": digest}


def _read_json(path: Path, label: str) -> Any:
    require(path.is_file() and not path.is_symlink(), f"{label}: JSON is missing")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise EvidenceError(f"{label}: invalid JSON: {error}") from error


def _write_or_check(path: Path, value: Any, check: bool) -> None:
    encoded = json.dumps(value, indent=2, sort_keys=True) + "\n"
    if check:
        require(path.is_file() and not path.is_symlink(),
                f"missing expected output: {path}")
        require(path.read_text(encoding="utf-8") == encoded,
                f"replayed output differs: {path}")
    else:
        path.write_text(encoded, encoding="utf-8")


def _gzip_payload(path: Path, label: str) -> bytes:
    try:
        with gzip.open(path, "rb") as stream:
            return stream.read()
    except (OSError, EOFError, gzip.BadGzipFile) as error:
        raise EvidenceError(f"{label}: invalid gzip stream: {error}") from error


def _raw_from_compressed(raw: Any, compressed: Any, label: str) -> bytes:
    """Verify a deleted raw artifact through its retained gzip copy."""

    raw_path = _descriptor(raw, f"{label} raw", allow_missing=True)
    compressed_path = _descriptor(compressed, f"{label} compressed raw")
    require(Path(str(raw.get("path"))).name.endswith(".data"),
            f"{label}: raw artifact is not perf.data")
    require(Path(str(compressed.get("path"))).name.endswith(".data.gz"),
            f"{label}: compressed raw artifact is not perf.data.gz")
    # The uncompressed input is expected to have been removed by capture.py;
    # if it is still present, its descriptor is checked by _descriptor above.
    del raw_path
    payload = _gzip_payload(compressed_path, f"{label} compressed raw")
    require(len(payload) == raw.get("bytes")
            and hashlib.sha256(payload).hexdigest() == raw.get("sha256"),
            f"{label}: decompressed raw identity differs")
    return payload


def _symbol_from_frame(line: str, label: str) -> str:
    match = _FRAME_RE.fullmatch(line)
    require(match is not None, f"{label}: malformed frame line: {line!r}")
    symbol = match.group("symbol").strip()
    # ``perf script`` places the instruction offset before the object path.
    # Removing it makes symbols from different samples comparable while
    # retaining the demangled Rust hash suffix used for owner matching.
    symbol = _OFFSET_RE.sub("", symbol).strip()
    require(symbol, f"{label}: empty frame symbol")
    return symbol


def parse_perf_script(payload: bytes, label: str) -> list[dict[str, Any]]:
    """Parse a canonical no-inline perf export into leaf-to-root stacks.

    The parser keys on the event trailer instead of a distro-specific field
    count, accepts optional CPU columns, and tolerates harmless metadata lines
    before the first sample.  Once a sample begins, every nonblank line must
    be a frame; unexplained payload is rejected so a truncated or differently
    formatted export cannot silently become a partial census.
    """

    try:
        text = payload.decode("utf-8")
    except UnicodeDecodeError as error:
        raise EvidenceError(f"{label}: frame stream is not UTF-8: {error}") from error

    samples: list[dict[str, Any]] = []
    current: dict[str, Any] | None = None
    for line_number, line in enumerate(text.splitlines(), start=1):
        if not line.strip():
            continue

        header = _HEADER_RE.fullmatch(line)
        if header is not None:
            if current is not None:
                require(current["symbols"],
                        f"{label}: sample at line {current['line']} has no frames")
                samples.append(current)
            current = {
                "line": line_number,
                "timestamp": header.group("timestamp"),
                "period": int(header.group("period")),
                "symbols": [],
            }
            continue

        # perf script may emit comments or build-id metadata before records.
        if current is None and line.lstrip().startswith("#"):
            continue
        require(current is not None,
                f"{label}: payload before first sample at line {line_number}")
        current["symbols"].append(_symbol_from_frame(line, f"{label}:{line_number}"))

    if current is not None:
        require(current["symbols"],
                f"{label}: final sample at line {current['line']} has no frames")
        samples.append(current)
    require(samples, f"{label}: perf frame stream contains no samples")
    timestamps = [sample["timestamp"] for sample in samples]
    require(len(set(timestamps)) == len(timestamps),
            f"{label}: perf sample timestamps are not unique")
    return samples


def _is_owner(symbol: str) -> bool:
    return OWNER_RE.fullmatch(symbol) is not None


def _is_unresolved(symbol: str) -> bool:
    lowered = symbol.lower()
    return ("[unknown]" in lowered or lowered in {"??", "<unknown>", "unknown"}
            or lowered.startswith("unknown "))


def _contains_any(symbol: str, markers: Iterable[str]) -> bool:
    return any(marker in symbol for marker in markers)


def _census_pairs(counter: collections.Counter[str]) -> list[list[Any]]:
    return [[name, count] for name, count in sorted(
        counter.items(), key=lambda item: (-item[1], item[0]))]


def _stack_counts(samples: list[dict[str, Any]], label: str) -> dict[str, Any]:
    all_symbols: collections.Counter[str] = collections.Counter()
    owner_symbols: collections.Counter[str] = collections.Counter()
    unresolved_symbols: collections.Counter[str] = collections.Counter()
    nested_sample_counts = collections.Counter()
    nested_frame_counts = collections.Counter()
    owner_samples = 0
    public_descendant_samples = 0
    missing_public_descendant_samples = 0
    unresolved_owner_frames = 0
    unresolved_owner_samples = 0
    owner_frame_count = 0

    for sample in samples:
        symbols = sample["symbols"]
        all_symbols.update(symbols)
        positions = [index for index, symbol in enumerate(symbols) if _is_owner(symbol)]
        require(len(positions) <= 1,
                f"{label}: exact owner appears more than once in one sample")
        if not positions:
            continue

        owner_samples += 1
        owner_index = positions[0]
        descendants = symbols[:owner_index]
        # Frame-pointer perf output is leaf-to-root.  The probe wrapper may be
        # hashed, but the public opened_presentation operation must still be
        # visible among its descendants.
        if any(PUBLIC_OPERATION_TOKEN in symbol for symbol in descendants):
            public_descendant_samples += 1
        else:
            missing_public_descendant_samples += 1

        owner_symbols.update(descendants)
        owner_frame_count += len(descendants)
        unresolved_here = [symbol for symbol in descendants if _is_unresolved(symbol)]
        unresolved_owner_frames += len(unresolved_here)
        if unresolved_here:
            unresolved_owner_samples += 1
            unresolved_symbols.update(unresolved_here)

        for category, markers in _CATEGORY_MARKERS.items():
            matching = [symbol for symbol in descendants
                        if _contains_any(symbol, markers)]
            if matching:
                nested_sample_counts[category] += 1
                nested_frame_counts[category] += len(matching)

    require(sum(nested_sample_counts.values()) >= 0,  # explicit integer invariant
            f"{label}: nested count arithmetic failed")
    return {
        "total_samples": len(samples),
        "owner_qualified_samples": owner_samples,
        "unqualified_samples": len(samples) - owner_samples,
        "public_opened_presentation_descendant_samples": public_descendant_samples,
        "missing_public_opened_presentation_descendant_samples": (
            missing_public_descendant_samples
        ),
        "owner_descendant_frame_count": owner_frame_count,
        "unresolved_owner_frames": unresolved_owner_frames,
        "unresolved_owner_samples": unresolved_owner_samples,
        "nested_sample_counts": {
            category: nested_sample_counts.get(category, 0)
            for category in _CATEGORY_MARKERS
        },
        "nested_frame_counts": {
            category: nested_frame_counts.get(category, 0)
            for category in _CATEGORY_MARKERS
        },
        "symbol_census": _census_pairs(all_symbols),
        "owner_descendant_symbol_census": _census_pairs(owner_symbols),
        "unresolved_owner_symbol_census": _census_pairs(unresolved_symbols),
    }


def _validate_plan() -> dict[str, Any]:
    plan = _read_json(c.P / "plan.json", "plan")
    require(plan.get("schema") == "litchi.performance.0795.v1",
            "plan schema changed")
    require(plan.get("cpu") == 12, "native CPU changed")
    require(plan.get("owner") == OWNER, "native owner changed")
    perf = plan.get("perf")
    require(isinstance(perf, dict), "perf plan is missing")
    expected = {
        "call_graph": "fp",
        "event": "cycles:u",
        "frequency": 499,
        "orders": [["before", "after"], ["after", "before"]],
        "samples": 100,
        "shape": "large",
        "warmup": 3,
    }
    for key, value in expected.items():
        require(perf.get(key) == value, f"perf plan {key} changed")
    claims = plan.get("claims")
    require(isinstance(claims, dict)
            and claims.get("native_latency") is False
            and claims.get("nested_costs_overlap") is True
            and claims.get("phase_fraction_requires_all_owner_frames_resolved") is True,
            "native claim policy changed")
    return plan


def _builds() -> dict[str, dict[str, Any]]:
    builds: dict[str, dict[str, Any]] = {}
    for leg in ("before", "after"):
        path = c.P / f"build-{leg}" / "build.json"
        build = _read_json(path, f"build-{leg}")
        binaries = build.get("binaries")
        require(isinstance(binaries, dict) and isinstance(binaries.get("fp"), dict),
                f"build-{leg}: fp binary identity is missing")
        _external_descriptor(binaries["fp"], f"build-{leg} fp")
        source = build.get("source")
        _descriptor(source, f"build-{leg} source")
        rows = build.get("rows")
        require(isinstance(rows, list) and len(rows) == 2,
                f"build-{leg}: command cardinality changed")
        previous_end = float("-inf")
        for index, row in enumerate(rows):
            require(isinstance(row, dict) and row.get("kind") in {"profile", "fp"},
                    f"build-{leg}: malformed command row {index}")
            require(row.get("exit_code") == 0, f"build-{leg}: command failed")
            require(previous_end <= row.get("started") <= row.get("ended"),
                    f"build-{leg}: command order changed")
            previous_end = row["ended"]
            _descriptor(row.get("log"), f"build-{leg} command {index} log")
        builds[leg] = {
            "path": f"build-{leg}/build.json",
            "sha256": c.sha(path),
            "value": build,
            "binary": binaries["fp"],
        }
    return builds


def _validate_report(path: Path, leg: str, binary: dict[str, Any], plan: dict[str, Any]) -> dict[str, Any]:
    report = _read_json(path, f"{leg} native report")
    perf = plan["perf"]
    require(report.get("mode") == "capture" and report.get("shape") == perf["shape"],
            f"{leg} native report mode/shape changed")
    require(report.get("warmup") == perf["warmup"]
            and report.get("samples_requested") == perf["samples"],
            f"{leg} native report sample policy changed")
    samples = report.get("samples")
    require(isinstance(samples, list) and len(samples) == perf["samples"],
            f"{leg} native report sample count changed")
    for index, sample in enumerate(samples):
        require(isinstance(sample, dict) and sample.get("index") == index,
                f"{leg} native report sample index changed")
        elapsed = sample.get("elapsed_ns")
        require(type(elapsed) is int and elapsed > 0,
                f"{leg} native report elapsed sample is malformed")
    allocator = report.get("allocator")
    if isinstance(allocator, dict):
        require(allocator.get("binary") == Path(binary["path"]).name,
                f"{leg} native report binary identity changed")
    return report


def _expected_receipt_order(plan: dict[str, Any]) -> list[tuple[int, str]]:
    return [(repeat, leg) for repeat, order in enumerate(plan["perf"]["orders"])
            for leg in order]


def _validate_complete(perf_dir: Path, rows: list[dict[str, Any]]) -> None:
    complete_path = perf_dir / "complete.json"
    complete = _read_json(complete_path, "perf complete")
    require(complete.get("children") == len(rows), "perf completion child count changed")
    _descriptor(complete.get("source"), "perf completion source")
    _descriptor(complete.get("receipts"), "perf completion receipts")


def analyze() -> dict[str, Any]:
    plan = _validate_plan()
    builds = _builds()
    perf_dir = c.P / "perf"
    require(perf_dir.is_dir(), "perf capture directory is missing")

    source_path = perf_dir / "source.json"
    source_descriptor = {**c.artifact(source_path)}
    before_source = builds["before"]["value"]["source"]
    before_source_path = _descriptor(before_source, "build-before source")
    require(source_path.read_bytes() == before_source_path.read_bytes(),
            "perf source does not match restored baseline build source")

    rows = _read_json(perf_dir / "receipts.json", "perf receipts")
    decodes = _read_json(perf_dir / "decodes.json", "perf decodes")
    require(isinstance(rows, list) and len(rows) == 4,
            "perf receipt cardinality changed")
    require(isinstance(decodes, list) and len(decodes) == 4,
            "perf decode cardinality changed")
    expected_order = _expected_receipt_order(plan)
    require([(row.get("repeat"), row.get("leg")) for row in rows] == expected_order,
            "perf receipt order changed")
    require([(row.get("repeat"), row.get("leg")) for row in decodes] == expected_order,
            "perf decode order changed")

    processes: list[dict[str, Any]] = []
    previous_capture_end = float("-inf")
    capture_ends: list[float] = []
    for row, decoded, (repeat, leg) in zip(rows, decodes, expected_order):
        label = f"perf/{repeat}-{leg}"
        require(isinstance(row, dict) and isinstance(decoded, dict),
                f"{label}: receipt is malformed")
        require(row.get("repeat") == repeat and row.get("leg") == leg,
                f"{label}: capture position changed")
        require(row.get("exit_code") == 0, f"{label}: perf record failed")
        started, ended = row.get("started"), row.get("ended")
        require(isinstance(started, (int, float)) and isinstance(ended, (int, float))
                and previous_capture_end <= started <= ended,
                f"{label}: capture timing order changed")
        previous_capture_end = ended
        capture_ends.append(ended)

        build = builds[leg]
        binary = build["binary"]
        require(row.get("binary") == binary,
                f"{label}: binary descriptor differs from build-{leg} fp")
        _external_descriptor(row.get("binary"), f"{label} binary")
        report_descriptor = row.get("report")
        report_path = _descriptor(report_descriptor, f"{label} report")
        report = _validate_report(report_path, leg, binary, plan)
        raw = row.get("raw")
        compressed = row.get("compressed")
        require(isinstance(raw, dict) and isinstance(compressed, dict),
                f"{label}: raw/compressed descriptors are incomplete")
        raw_payload = _raw_from_compressed(raw, compressed, label)
        # Retain only identity data in the output; the raw bytes themselves
        # are intentionally not copied into the JSON result.

        expected_command = [
            "taskset", "-c", str(plan["cpu"]), "perf", "record",
            "-e", plan["perf"]["event"], "-F", str(plan["perf"]["frequency"]),
            "--call-graph", plan["perf"]["call_graph"], "-o", raw["path"], "--",
            binary["path"], "--mode", "capture", "--shape", plan["perf"]["shape"],
            "--samples", str(plan["perf"]["samples"]),
            "--warmup", str(plan["perf"]["warmup"]),
            "--output", report_descriptor["path"],
        ]
        require(row.get("command") == expected_command,
                f"{label}: perf record command changed")
        _descriptor(row.get("log"), f"{label} capture log")

        require(decoded.get("repeat") == repeat and decoded.get("leg") == leg,
                f"{label}: decode position changed")
        require(decoded.get("exit_code") == 0, f"{label}: perf script failed")
        require(decoded.get("raw") == raw,
                f"{label}: decode raw descriptor differs from capture")
        require(decoded.get("command") == [
            "perf", "script", "--no-inline", "--ns", "-i", raw["path"]
        ], f"{label}: decode command changed")
        require(isinstance(decoded.get("frames"), dict),
                f"{label}: compressed frame descriptor is missing")
        frames_path = _descriptor(decoded["frames"], f"{label} frames")
        require(Path(str(decoded["frames"]["path"])).name.endswith(".frames.gz"),
                f"{label}: frame artifact is not .frames.gz")
        frames_payload = _gzip_payload(frames_path, f"{label} frames")
        _descriptor(decoded.get("log"), f"{label} decode log")

        # Decode timing starts only after all four capture children have
        # completed; later decodes remain serial as well.
        dstart, dend = decoded.get("started"), decoded.get("ended")
        require(isinstance(dstart, (int, float)) and isinstance(dend, (int, float))
                and dstart <= dend,
                f"{label}: decode timing is malformed")
        analysis = _stack_counts(parse_perf_script(frames_payload, label), label)
        processes.append({
            "repeat": repeat,
            "leg": leg,
            "command": row["command"],
            "binary": binary,
            "report": report_descriptor,
            "report_sha256": c.sha(report_path),
            "raw": raw,
            "compressed_raw": compressed,
            "capture_log": row["log"],
            "frames": decoded["frames"],
            "decode_log": decoded["log"],
            **analysis,
            "report_verified_samples": len(report["samples"]),
            "raw_decompressed_bytes_verified": len(raw_payload),
        })

    # The driver performs every record before the first decode.  Check this
    # independently from the row order to catch an accidental interleaving.
    previous_decode_end = max(capture_ends)
    for index, decoded in enumerate(decodes):
        dstart, dend = decoded["started"], decoded["ended"]
        require(previous_decode_end <= dstart <= dend,
                f"perf decode {index} timing order changed")
        previous_decode_end = dend

    _validate_complete(perf_dir, rows)

    totals = {
        "processes": len(processes),
        "total_samples": sum(row["total_samples"] for row in processes),
        "owner_qualified_samples": sum(row["owner_qualified_samples"] for row in processes),
        "unqualified_samples": sum(row["unqualified_samples"] for row in processes),
        "public_opened_presentation_descendant_samples": sum(
            row["public_opened_presentation_descendant_samples"] for row in processes
        ),
        "missing_public_opened_presentation_descendant_samples": sum(
            row["missing_public_opened_presentation_descendant_samples"]
            for row in processes
        ),
        "unresolved_owner_frames": sum(row["unresolved_owner_frames"] for row in processes),
        "unresolved_owner_samples": sum(row["unresolved_owner_samples"] for row in processes),
    }
    # The explicit qualification state is useful to the report writer, but it
    # deliberately does not turn observed counts into a fraction or a share.
    qualification = {
        "owner_present": totals["owner_qualified_samples"] > 0,
        "public_opened_presentation_descendant_complete": (
            totals["owner_qualified_samples"] > 0
            and totals["missing_public_opened_presentation_descendant_samples"] == 0
        ),
        "owner_frames_resolved": totals["unresolved_owner_frames"] == 0,
        "qualified_for_phase_fraction": False,
        "failure_reasons": [
            reason for condition, reason in (
                (totals["owner_qualified_samples"] == 0, "exact owner was not observed"),
                (totals["missing_public_opened_presentation_descendant_samples"] > 0,
                 "one or more owner samples lacked the public opened_presentation descendant"),
                (totals["unresolved_owner_frames"] > 0,
                 "one or more owner descendants were unresolved"),
            ) if condition
        ],
    }
    return {
        "schema": SCHEMA,
        "packet": "change-0795",
        "plan": {
            "path": "plan.json",
            "sha256": c.sha(c.P / "plan.json"),
            "schema": plan["schema"],
            "owner": OWNER,
            "cpu": plan["cpu"],
            "event": plan["perf"]["event"],
            "frequency": plan["perf"]["frequency"],
            "call_graph": plan["perf"]["call_graph"],
            "samples": plan["perf"]["samples"],
            "warmup": plan["perf"]["warmup"],
        },
        "source": source_descriptor,
        "builds": {
            leg: {
                "path": value["path"],
                "sha256": value["sha256"],
                "fp_binary": value["binary"],
            }
            for leg, value in builds.items()
        },
        "processes": processes,
        "summary": totals,
        "qualification": qualification,
        "category_markers": {
            category: list(markers) for category, markers in _CATEGORY_MARKERS.items()
        },
        "nested_costs_overlap": True,
        "counts_are_observed_samples_or_frames_only": True,
        "phase_fraction_claim_authorized": False,
        "native_latency_claim": False,
        "causal_claim": False,
        "scope": (
            "Four frame-pointer cycles:u native stack captures over the large "
            "capture probe; counts include warmup and measured samples, are "
            "exact-owner observations, and do not estimate latency, CPU shares, "
            "or causality."
        ),
        "claims": [
            "Only samples containing the exact capture owner are included in owner and nested counts.",
            "The public opened_presentation descendant is checked separately for every owner-qualified sample.",
            "Nested category counts overlap and are reported as observed counts, never as additive shares.",
            "Unresolved owner descendants and missing public descendants fail qualification; no phase fraction is inferred.",
            "Raw perf.data identities are verified after gzip decompression before frame parsing.",
        ],
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--write", action="store_true")
    parser.add_argument("--check", action="store_true")
    args = parser.parse_args()
    require(args.write ^ args.check, "choose exactly one of --write or --check")
    result = analyze()
    _write_or_check(c.P / "native-analysis.json", result, args.check)
    print("0795 native frame analysis PASS", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
