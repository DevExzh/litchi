"""Independent frame and public-workflow audit for the 0812 fresh samples.

The reader consumes only retained receipts, reports, gzip members, and frame
text.  It does not execute the rebuilt probe or inspect its assembly.  The
fresh binary identity is taken from ``fresh/complete.json`` and is attached to
the exact-owner DSO checks; instruction mapping remains a root-side join.
"""

from __future__ import annotations

import gzip
import hashlib
import json
import re
import sys
from collections import Counter
from pathlib import Path
from typing import Any


PACKET = Path(__file__).resolve().parent
ROOT = PACKET.parents[3]
FRESH = PACKET / "fresh"
OLD = PACKET.parent / "change-0811"
OUTPUT = PACKET / "fresh-offset-audit.json"

OWNER = "namespace_uri_probe::capture_region_0793"
SCANNER = "litchi_pptx::notes::codec::scan_processed_xml"
EVENT = "cycles:u"
BINARY_PATH = "/home/zhuhe/code/litchi-target-0811/fp"

HEADER_RE = re.compile(
    r"(?P<command>\S+)\s+(?P<pid>[0-9]+)\s+"
    r"(?P<timestamp>[0-9]+\.[0-9]+):\s+"
    r"(?P<period>[0-9]+)\s+cycles:u:\s*"
)
FRAME_RE = re.compile(
    r"\s*(?P<ip>[0-9a-f]+)\s+(?P<symbol>.+?)"
    r"(?:\+(?P<offset>0x[0-9a-f]+))?\s+\((?P<dso>[^()]+)\)\s*"
)


class AuditError(RuntimeError):
    """A retained fresh artifact is missing or inconsistent."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise AuditError(message)


def read_json(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing JSON: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise AuditError(f"invalid JSON {path}: {error}") from error


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for chunk in iter(lambda: stream.read(1 << 20), b""):
                digest.update(chunk)
    except OSError as error:
        raise AuditError(f"cannot hash {path}: {error}") from error
    return digest.hexdigest()


def artifact(path: Path) -> dict[str, Any]:
    require(path.is_file() and not path.is_symlink(), f"missing artifact: {path}")
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": sha256_file(path)}


def descriptor(value: Any, expected: dict[str, Any], label: str) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label}: malformed descriptor")
    require(value == expected, f"{label}: artifact identity changed")
    return expected


def require_string(value: Any, label: str) -> str:
    require(isinstance(value, str) and value, f"{label}: expected non-empty string")
    return value


def require_integer(value: Any, label: str, *, positive: bool = False) -> int:
    require(
        isinstance(value, int)
        and not isinstance(value, bool)
        and (value > 0 if positive else value >= 0),
        f"{label}: expected integer",
    )
    return value


def verify_gzip_header(data: bytes, label: str) -> None:
    require(len(data) >= 10, f"{label}: gzip member is too short")
    require(data[:2] == b"\x1f\x8b" and data[2] == 8, f"{label}: invalid gzip header")
    require(data[3] & 0x04 == 0, f"{label}: gzip extra field is not allowed")
    require(data[4:8] == b"\x00\x00\x00\x00", f"{label}: gzip mtime is not zero")


def verify_compression() -> tuple[dict[str, Any], dict[tuple[int, str], dict[str, Any]]]:
    path = FRESH / "compression.json"
    receipt = artifact(path)
    rows = read_json(path)
    require(isinstance(rows, list) and len(rows) == 4, "fresh compression row count changed")
    result: dict[tuple[int, str], dict[str, Any]] = {}
    for index, row in enumerate(rows):
        label = f"fresh compression[{index}]"
        require(isinstance(row, dict), f"{label}: malformed row")
        repeat = require_integer(row.get("repeat"), f"{label}.repeat")
        kind = require_string(row.get("kind"), f"{label}.kind")
        require(repeat in (0, 1) and kind in ("data", "frames"),
                f"{label}: unexpected repeat/kind")
        key = (repeat, kind)
        require(key not in result, f"{label}: duplicate repeat/kind")
        compressed_path = FRESH / f"{repeat}.{kind}.gz"
        original_path = FRESH / f"{repeat}.{kind}"
        compressed = row.get("compressed")
        original = row.get("original")
        require(isinstance(compressed, dict) and isinstance(original, dict),
                f"{label}: malformed artifact descriptors")
        require(compressed.get("path") == str(compressed_path),
                f"{label}: compressed path changed")
        require(original.get("path") == str(original_path),
                f"{label}: original path changed")
        compressed_identity = descriptor(compressed, artifact(compressed_path),
                                         f"{label}.compressed")
        compressed_bytes = compressed_path.read_bytes()
        verify_gzip_header(compressed_bytes, label)
        try:
            decompressed = gzip.decompress(compressed_bytes)
        except (OSError, EOFError) as error:
            raise AuditError(f"{label}: gzip decompression failed: {error}") from error
        original_bytes = require_integer(original.get("bytes"), f"{label}.original.bytes")
        original_sha = require_string(original.get("sha256"), f"{label}.original.sha256")
        decompressed_sha = sha256_bytes(decompressed)
        require(len(decompressed) == original_bytes, f"{label}: decompressed length changed")
        require(decompressed_sha == original_sha, f"{label}: decompressed hash changed")
        result[key] = {
            "repeat": repeat,
            "kind": kind,
            "compressed": compressed_identity,
            "original": {"path": str(original_path), "bytes": original_bytes,
                         "sha256": decompressed_sha},
            "gzip_mtime": 0,
        }
    require(set(result) == {(0, "data"), (0, "frames"), (1, "data"), (1, "frames")},
            "fresh compression repeat/kind matrix changed")
    return receipt, result


def report_against_oracle(report: dict[str, Any], oracle: dict[str, Any], label: str) -> None:
    require(set(report) == set(oracle), f"{label}: report schema keys changed")
    for key in report:
        if key != "samples":
            require(report[key] == oracle[key], f"{label}: oracle field changed: {key}")
    samples = report.get("samples")
    oracle_samples = oracle.get("samples")
    require(isinstance(samples, list) and len(samples) == 100,
            f"{label}: measured sample count changed")
    require(isinstance(oracle_samples, list) and len(oracle_samples) == 100,
            "0811 oracle sample count changed")
    for index, (sample, expected) in enumerate(zip(samples, oracle_samples, strict=True)):
        require(isinstance(sample, dict) and isinstance(expected, dict),
                f"{label} sample {index}: malformed")
        require(sample.get("index") == index and expected.get("index") == index,
                f"{label} sample {index}: index changed")
        elapsed = sample.get("elapsed_ns")
        require(isinstance(elapsed, int) and elapsed > 0
                and sample.get("metrics", {}).get("elapsed_ns") == elapsed,
                f"{label} sample {index}: elapsed timing is malformed")
        require(sample.get("source_sha256") == expected.get("source_sha256"),
                f"{label} sample {index}: source identity changed")
        require(sample.get("output") == expected.get("output"),
                f"{label} sample {index}: output identity changed")
        require(sample.get("verification") == expected.get("verification"),
                f"{label} sample {index}: verification changed")
        metrics = sample.get("metrics")
        expected_metrics = expected.get("metrics")
        require(isinstance(metrics, dict) and isinstance(expected_metrics, dict),
                f"{label} sample {index}: metrics malformed")
        require({k: v for k, v in metrics.items() if k != "elapsed_ns"}
                == {k: v for k, v in expected_metrics.items() if k != "elapsed_ns"},
                f"{label} sample {index}: capture metrics changed")


def verify_receipts(
    complete: dict[str, Any],
    compression: dict[tuple[int, str], dict[str, Any]],
    oracle: dict[str, Any],
) -> tuple[dict[str, Any], list[dict[str, Any]], dict[str, Any]]:
    binary = complete.get("binary")
    require(isinstance(binary, dict) and binary.get("path") == BINARY_PATH,
            "fresh binary descriptor/path changed")
    require(require_integer(binary.get("bytes"), "fresh binary bytes", positive=True) > 0,
            "fresh binary length is invalid")
    require(re.fullmatch(r"[0-9a-f]{64}", require_string(binary.get("sha256"),
                                                         "fresh binary sha256")),
            "fresh binary hash is invalid")
    receipts_path = FRESH / "receipts.json"
    receipts_descriptor = artifact(receipts_path)
    receipts = read_json(receipts_path)
    require(isinstance(receipts, list) and len(receipts) == 2, "fresh receipt count changed")
    reports: list[dict[str, Any]] = []
    seen: set[int] = set()
    for index, row in enumerate(receipts):
        label = f"fresh receipt {index}"
        require(isinstance(row, dict), f"{label}: malformed")
        repeat = require_integer(row.get("repeat"), f"{label}.repeat")
        require(repeat in (0, 1) and repeat not in seen, f"{label}: repeat changed")
        seen.add(repeat)
        require(row.get("binary") == binary and row.get("exit_code") == 0,
                f"{label}: binary/exit status changed")
        expected_raw = compression[(repeat, "data")]["original"]
        require(row.get("raw") == expected_raw, f"{label}: raw identity changed")
        report_path = FRESH / f"{repeat}.json"
        require(isinstance(row.get("report"), dict)
                and row["report"].get("path") == str(report_path),
                f"{label}: report path changed")
        report_identity = artifact(report_path)
        require(row["report"] == report_identity, f"{label}: report identity changed")
        require(row.get("output", {}).get("path") == str(FRESH / f"{repeat}.log"),
                f"{label}: log path changed")
        require(row.get("errors") is None, f"{label}: stderr artifact changed")
        command = row.get("command")
        expected_command = [
            "taskset", "-c", "12", "perf", "record", "--no-buildid-cache",
            "-e", EVENT, "-F", "499", "--call-graph", "fp", "-o",
            str(FRESH / f"{repeat}.data"), "--", BINARY_PATH,
            "--mode", "capture", "--shape", "large", "--samples", "100",
            "--warmup", "0", "--output", str(report_path),
        ]
        require(command == expected_command, f"{label}: command contract changed")
        report = read_json(report_path)
        report_against_oracle(report, oracle, f"fresh/{repeat}.json")
        reports.append({"repeat": repeat, "artifact": report_identity, "report": report})

        individual = FRESH / f"{repeat}.receipt.json"
        require(read_json(individual) == row, f"{label}: individual receipt differs")
        artifact(individual)
    require(seen == {0, 1}, "fresh receipt repeat set changed")
    return receipts_descriptor, reports, binary


def verify_decode(
    binary: dict[str, Any],
    compression: dict[tuple[int, str], dict[str, Any]],
) -> dict[str, Any]:
    decode_path = FRESH / "decode.json"
    decode_descriptor = artifact(decode_path)
    rows = read_json(decode_path)
    require(isinstance(rows, list) and len(rows) == 2, "fresh decode count changed")
    seen: set[int] = set()
    for index, row in enumerate(rows):
        label = f"fresh decode {index}"
        require(isinstance(row, dict), f"{label}: malformed")
        repeat = require_integer(row.get("repeat"), f"{label}.repeat")
        require(repeat in (0, 1) and repeat not in seen, f"{label}: repeat changed")
        seen.add(repeat)
        require(row.get("binary") == binary and row.get("exit_code") == 0,
                f"{label}: binary/exit status changed")
        require(row.get("raw") == compression[(repeat, "data")]["original"],
                f"{label}: raw identity changed")
        require(row.get("output") == compression[(repeat, "frames")]["original"],
                f"{label}: frame identity changed")
        require(row.get("command") == [
            "perf", "script", "--no-inline", "--ns", "-i",
            str(FRESH / f"{repeat}.data"),
        ], f"{label}: decode command changed")
        errors = row.get("errors")
        require(isinstance(errors, dict) and errors.get("path") == str(FRESH / f"{repeat}.decode.log"),
                f"{label}: decode stderr identity changed")
        require(artifact(FRESH / f"{repeat}.decode.log") == errors,
                f"{label}: decode stderr changed")
        individual = FRESH / f"{repeat}.decode.json"
        require(read_json(individual) == row, f"{label}: individual decode differs")
        artifact(individual)
    require(seen == {0, 1}, "fresh decode repeat set changed")
    return decode_descriptor


def parse_frames(
    path: Path,
    binary: dict[str, Any],
) -> dict[str, Any]:
    compressed = path.read_bytes()
    try:
        text = gzip.decompress(compressed).decode("utf-8")
    except (OSError, EOFError, UnicodeDecodeError) as error:
        raise AuditError(f"{path}: frame text is unreadable: {error}") from error
    lost_lines = [
        line for line in text.splitlines()
        if "PERF_RECORD_LOST" in line
        or ("lost" in line.lower() and "sample" in line.lower())
    ]
    blocks = text.strip().split("\n\n")
    require(blocks and blocks != [""], f"{path}: frame text is empty")
    whole_period = 0
    owner_period = 0
    owner_samples = 0
    unknown_interior_samples = 0
    unknown_interior_frames = 0
    unresolved_samples = 0
    unresolved_frames = 0
    scanner_offsets: Counter[int] = Counter()
    scanner_periods: Counter[int] = Counter()

    for block_index, block in enumerate(blocks):
        lines = block.splitlines()
        require(lines, f"{path}: empty block {block_index}")
        first = lines[0]
        if "PERF_RECORD_LOST" in first or ("lost" in first.lower() and "sample" in first.lower()):
            continue
        header = HEADER_RE.fullmatch(first)
        require(header is not None, f"{path}: malformed cycles:u header: {first!r}")
        period = int(header["period"])
        require(period > 0, f"{path}: non-positive period in block {block_index}")
        whole_period += period
        frames: list[dict[str, Any]] = []
        for line_index, line in enumerate(lines[1:], start=1):
            match = FRAME_RE.fullmatch(line)
            require(match is not None, f"{path}: malformed frame line: {line!r}")
            frame = {
                "ip": int(match["ip"], 16),
                "symbol": match["symbol"],
                "offset": int(match["offset"], 16) if match["offset"] else None,
                "dso": match["dso"],
            }
            frames.append(frame)
        require(frames, f"{path}: stack has no frames in block {block_index}")
        unknown = [
            frame for frame in frames
            if any(token in frame["symbol"].lower() for token in ("[unknown]", "??", "<unknown>"))
        ]
        if unknown:
            unresolved_samples += 1
            unresolved_frames += len(unknown)
        owner_indices = [index for index, frame in enumerate(frames)
                         if frame["symbol"] == OWNER]
        require(len(owner_indices) <= 1, f"{path}: owner repeated in block {block_index}")
        if not owner_indices:
            continue
        owner_index = owner_indices[0]
        owner_samples += 1
        owner_period += period
        require(frames[owner_index]["dso"] == binary["path"],
                f"{path}: owner DSO changed in block {block_index}")
        interior = frames[:owner_index]
        interior_unknown = [
            frame for frame in interior
            if any(token in frame["symbol"].lower() for token in ("[unknown]", "??", "<unknown>"))
        ]
        if interior_unknown:
            unknown_interior_samples += 1
            unknown_interior_frames += len(interior_unknown)
        if not interior:
            continue
        leaf = interior[0]
        if leaf["symbol"] != SCANNER:
            continue
        require(leaf["dso"] == binary["path"] and leaf["offset"] is not None,
                f"{path}: scanner leaf DSO/offset changed in block {block_index}")
        scanner_offsets[leaf["offset"]] += 1
        scanner_periods[leaf["offset"]] += period

    return {
        "path": path.name,
        "whole_process_samples": len(blocks) - len(lost_lines),
        "whole_process_period": whole_period,
        "owner_qualified_samples": owner_samples,
        "owner_qualified_period": owner_period,
        "unknown_interior_samples": unknown_interior_samples,
        "unknown_interior_frames": unknown_interior_frames,
        "unresolved_samples": unresolved_samples,
        "unresolved_frames": unresolved_frames,
        "lost_event_lines": len(lost_lines),
        "lost_event_period": 0,
        "scanner_leaf_symbol": SCANNER,
        "scanner_leaf_samples": sum(scanner_offsets.values()),
        "scanner_leaf_period": sum(scanner_periods.values()),
        "offsets": [
            {"offset_hex": hex(offset), "samples": scanner_offsets[offset],
             "period": scanner_periods[offset]}
            for offset in sorted(scanner_offsets)
        ],
    }


def main() -> None:
    require(len(sys.argv) == 2 and sys.argv[1] in ("--write", "--check"),
            "usage: fresh_offset_audit.py --write|--check")
    complete_path = FRESH / "complete.json"
    complete_descriptor = artifact(complete_path)
    complete = read_json(complete_path)
    require(complete.get("schema") == "litchi.performance.0812.fresh-complete.v1",
            "fresh completion schema changed")
    require(complete.get("reports") == 2 and complete.get("samples") == 200,
            "fresh completion sample contract changed")
    frozen_path = FRESH / "frozen.json"
    frozen_descriptor = artifact(frozen_path)
    frozen = read_json(frozen_path)
    require(frozen.get("schema") == "litchi.performance.0812.fresh-frozen.v1",
            "fresh frozen schema changed")
    require(frozen.get("binary") == complete.get("binary"),
            "fresh frozen/completion binary identity changed")
    require(frozen.get("inputs") == complete.get("inputs"),
            "fresh frozen/completion input custody changed")

    compression_descriptor, compression = verify_compression()
    oracle_path = OLD / "perf/0.json"
    seal = read_json(OLD / "seal.json")
    seal_files = seal.get("files")
    require(isinstance(seal_files, dict), "0811 seal files missing")
    oracle_descriptor = artifact(oracle_path)
    require(seal_files.get("docs/performance/results/change-0811/perf/0.json")
            == oracle_descriptor["sha256"], "0811 oracle seal hash changed")
    oracle = read_json(oracle_path)
    receipts_descriptor, reports, binary = verify_receipts(complete, compression, oracle)
    decode_descriptor = verify_decode(binary, compression)

    rows = []
    for repeat in (0, 1):
        frames_path = FRESH / f"{repeat}.frames.gz"
        frame_compression = compression[(repeat, "frames")]
        require(artifact(frames_path) == frame_compression["compressed"],
                f"fresh/{repeat}.frames.gz identity changed")
        row = parse_frames(frames_path, binary)
        row["repeat"] = repeat
        row["frames"] = frame_compression
        row["report"] = reports[repeat]["artifact"]
        rows.append(row)
    require(sum(row["whole_process_samples"] for row in rows) == 6161,
            "fresh whole-process sample total changed")
    require(sum(row["owner_qualified_samples"] for row in rows) == 1930,
            "fresh exact-owner sample total changed")
    require(sum(row["scanner_leaf_samples"] for row in rows) == 458,
            "fresh scanner leaf sample total changed")

    result = {
        "schema": "litchi.performance.0812.fresh-offset-audit.v1",
        "scope": (
            "Independent fresh 0812 cycles:u frame census for the exact rebuilt "
            "binary and owner. Public-workflow output and verification fields are "
            "checked against the sealed 0811 oracle. Counts and offsets are sampled "
            "diagnostics; no assembly mapping, instruction cost, phase fraction, "
            "speedup, or adoption claim is made."
        ),
        "event": EVENT,
        "owner": OWNER,
        "binary": binary,
        "inputs": {
            "complete": complete_descriptor,
            "frozen": frozen_descriptor,
            "compression": compression_descriptor,
            "receipts": receipts_descriptor,
            "decode": decode_descriptor,
            "oracle": oracle_descriptor,
            "reports": [row["report"] for row in rows],
        },
        "rows": rows,
        "verification": {
            "gzip_members_and_decompressed_identities": True,
            "strict_cycles_u_headers": True,
            "exact_owner_and_dso_verified": True,
            "public_workflow_reports": 2,
            "public_workflow_samples": 200,
            "oracle_output_and_verification_fields_verified": True,
            "assembly_dependency": False,
        },
    }
    encoded = json.dumps(result, indent=2, sort_keys=True) + "\n"
    if sys.argv[1] == "--write":
        require(not OUTPUT.exists(), f"immutable output already exists: {OUTPUT}")
        OUTPUT.write_text(encoded, encoding="utf-8")
        print("0812 fresh offset audit PASS")
    else:
        require(OUTPUT.is_file() and not OUTPUT.is_symlink(),
                f"missing immutable output: {OUTPUT}")
        require(OUTPUT.read_text(encoding="utf-8") == encoded,
                f"fresh offset audit output changed: {OUTPUT}")
        print("0812 fresh offset audit --check PASS")


if __name__ == "__main__":
    try:
        main()
    except AuditError as error:
        raise SystemExit(f"fresh offset audit FAILED: {error}") from error
