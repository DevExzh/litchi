"""Audit retained 0811 scanner leaf instruction offsets without a binary.

This reader is deliberately independent of the 0812 rebuild and of any
assembly listing.  It verifies the retained gzip members and their custody
receipts, parses the two sealed ``perf script`` outputs, and counts offsets
only for exact-owner samples whose first frame is the scanner symbol.  The
binary identity is carried from the sealed 0811 build receipt so a later
assembly join can require the exact DSO without making this reader depend on
that DSO being present.
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
SEALED = PACKET.parent / "change-0811"
PERF = SEALED / "perf"
OUTPUT = PACKET / "offset-audit.json"

OWNER = "namespace_uri_probe::capture_region_0793"
SCANNER = "litchi_pptx::notes::codec::scan_processed_xml"
BINARY_PATH = "/home/zhuhe/code/litchi-target-0811/fp"
EVENT = "cycles:u"

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
    """A retained evidence file is missing, stale, or internally inconsistent."""


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
    return {
        "path": str(path),
        "bytes": path.stat().st_size,
        "sha256": sha256_file(path),
    }


def relative_root(path: Path) -> str:
    try:
        return path.relative_to(ROOT).as_posix()
    except ValueError as error:
        raise AuditError(f"path is outside repository root: {path}") from error


def sealed_artifact(path: Path, seal_files: dict[str, Any]) -> dict[str, Any]:
    value = artifact(path)
    key = relative_root(path)
    require(key in seal_files, f"0811 seal has no entry for {key}")
    require(seal_files[key] == value["sha256"], f"0811 seal hash changed for {key}")
    return value


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


def path_descriptor(value: Any, label: str) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label}: malformed artifact descriptor")
    raw_path = require_string(value.get("path"), f"{label}.path")
    path = Path(raw_path)
    require(path.is_absolute(), f"{label}.path: expected absolute path")
    expected = artifact(path)
    require(value.get("bytes") == expected["bytes"], f"{label}.bytes changed")
    require(value.get("sha256") == expected["sha256"], f"{label}.sha256 changed")
    return expected


def sealed_descriptor(value: Any, label: str, seal_files: dict[str, Any]) -> dict[str, Any]:
    descriptor = path_descriptor(value, label)
    key = relative_root(Path(descriptor["path"]))
    require(key in seal_files, f"{label}: missing 0811 seal entry")
    require(seal_files[key] == descriptor["sha256"], f"{label}: 0811 seal hash changed")
    return descriptor


def pair_list(value: Any, label: str) -> dict[str, int]:
    require(isinstance(value, list), f"{label}: expected pair list")
    result: dict[str, int] = {}
    for index, row in enumerate(value):
        require(
            isinstance(row, list) and len(row) == 2,
            f"{label}[{index}]: malformed pair",
        )
        name = require_string(row[0], f"{label}[{index}].name")
        count = require_integer(row[1], f"{label}[{index}].count")
        require(name not in result, f"{label}: duplicate {name}")
        result[name] = count
    return result


def sorted_pairs(counter: Counter[str]) -> list[list[Any]]:
    return [[name, counter[name]] for name in sorted(counter)]


def verify_gzip_header(data: bytes, label: str) -> None:
    require(len(data) >= 10, f"{label}: gzip member is too short")
    require(data[:2] == b"\x1f\x8b", f"{label}: not a gzip member")
    require(data[2] == 8, f"{label}: gzip compression method changed")
    require(data[3] & 0x04 == 0, f"{label}: gzip extra field is not allowed")
    require(data[4:8] == b"\x00\x00\x00\x00", f"{label}: gzip mtime changed")


def verify_compression(
    compression_path: Path,
    seal_files: dict[str, Any],
) -> tuple[dict[str, Any], list[dict[str, Any]], dict[tuple[int, str], dict[str, Any]]]:
    compression_descriptor = sealed_artifact(compression_path, seal_files)
    rows = read_json(compression_path)
    require(isinstance(rows, list) and len(rows) == 4, "0811 compression row count changed")

    retained: list[dict[str, Any]] = []
    by_kind: dict[tuple[int, str], dict[str, Any]] = {}
    for index, row in enumerate(rows):
        label = f"compression[{index}]"
        require(isinstance(row, dict), f"{label}: malformed row")
        repeat = require_integer(row.get("repeat"), f"{label}.repeat")
        kind = require_string(row.get("kind"), f"{label}.kind")
        require(repeat in (0, 1), f"{label}.repeat: expected 0 or 1")
        require(kind in ("raw", "frames"), f"{label}.kind: unexpected kind")
        key = (repeat, kind)
        require(key not in by_kind, f"{label}: duplicate repeat/kind")

        suffix = "data" if kind == "raw" else "frames"
        expected_compressed = PERF / f"{repeat}.{suffix}.gz"
        expected_original = PERF / f"{repeat}.{suffix}"
        compressed = row.get("compressed")
        original = row.get("original")
        require(isinstance(compressed, dict), f"{label}.compressed: malformed")
        require(isinstance(original, dict), f"{label}.original: malformed")
        require(compressed.get("path") == str(expected_compressed), f"{label}: compressed path changed")
        require(original.get("path") == str(expected_original), f"{label}: original path changed")
        require(row.get("compression") == "gzip", f"{label}: compression changed")
        require(row.get("gzip_mtime") == 0, f"{label}: gzip mtime changed")

        compressed_descriptor = sealed_descriptor(
            compressed,
            f"{label}.compressed",
            seal_files,
        )
        compressed_bytes = expected_compressed.read_bytes()
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
        require(row.get("decompressed_sha256") == decompressed_sha,
                f"{label}: decompressed receipt hash changed")

        value = {
            "repeat": repeat,
            "kind": kind,
            "compressed": compressed_descriptor,
            "decompressed": {
                "path": str(expected_original),
                "bytes": original_bytes,
                "sha256": decompressed_sha,
            },
            "compression": "gzip",
            "gzip_mtime": 0,
        }
        by_kind[key] = value
        retained.append(value)

    require(
        sorted(by_kind) == [(0, "frames"), (0, "raw"), (1, "frames"), (1, "raw")],
        "0811 compression repeat/kind matrix changed",
    )
    retained.sort(key=lambda row: (row["repeat"], row["kind"]))
    return compression_descriptor, retained, by_kind


def parse_frames(
    path: Path,
    expected_binary: dict[str, Any],
) -> dict[str, Any]:
    compressed = path.read_bytes()
    try:
        text = gzip.decompress(compressed).decode("utf-8")
    except (OSError, EOFError, UnicodeDecodeError) as error:
        raise AuditError(f"{path}: retained frame text is unreadable: {error}") from error

    stripped = text.strip()
    require(stripped, f"{path}: frame text is empty")
    blocks = stripped.split("\n\n")
    whole_period = 0
    owner_period = 0
    owner_samples = 0
    unknown_interior = 0
    leaf_counts: Counter[str] = Counter()
    leaf_period: Counter[str] = Counter()
    nested_counts: Counter[str] = Counter()
    frame_occurrences: Counter[str] = Counter()
    scanner_offsets: Counter[int] = Counter()
    scanner_offset_period: Counter[int] = Counter()

    for block_index, block in enumerate(blocks):
        lines = block.splitlines()
        require(lines and lines[0], f"{path}: empty block {block_index}")
        header = HEADER_RE.fullmatch(lines[0])
        require(header is not None, f"{path}: malformed cycles:u header: {lines[0]!r}")
        period = int(header["period"])
        require(period > 0, f"{path}: non-positive period in block {block_index}")
        whole_period += period

        frames: list[dict[str, Any]] = []
        for line_index, line in enumerate(lines[1:], start=1):
            match = FRAME_RE.fullmatch(line)
            require(match is not None, f"{path}: malformed frame line {line!r}")
            frame = {
                "ip": int(match["ip"], 16),
                "symbol": match["symbol"],
                "offset": int(match["offset"], 16) if match["offset"] else None,
                "dso": match["dso"],
            }
            require(frame["symbol"], f"{path}: empty symbol at frame {line_index}")
            require(frame["dso"], f"{path}: empty DSO at frame {line_index}")
            frames.append(frame)

        owner_indices = [index for index, frame in enumerate(frames) if frame["symbol"] == OWNER]
        require(len(owner_indices) <= 1, f"{path}: owner appears more than once in block {block_index}")
        if not owner_indices:
            continue

        owner_index = owner_indices[0]
        require(frames[owner_index]["dso"] == expected_binary["path"],
                f"{path}: owner DSO changed in block {block_index}")
        owner_samples += 1
        owner_period += period
        inside = frames[:owner_index]
        names = [frame["symbol"] for frame in frames]
        frame_occurrences.update(names[: owner_index + 1])
        nested_counts.update(set(names[:owner_index]))
        if any(
            any(token in name.lower() for token in ("[unknown]", "??", "<unknown>"))
            for name in names[:owner_index]
        ):
            unknown_interior += 1

        leaf = inside[0] if inside else frames[owner_index]
        leaf_name = leaf["symbol"]
        leaf_counts[leaf_name] += 1
        leaf_period[leaf_name] += period

        if leaf_name == SCANNER:
            require(leaf["dso"] == expected_binary["path"],
                    f"{path}: scanner leaf DSO changed in block {block_index}")
            require(leaf["offset"] is not None,
                    f"{path}: scanner leaf has no symbol offset in block {block_index}")
            scanner_offsets[leaf["offset"]] += 1
            scanner_offset_period[leaf["offset"]] += period

    return {
        "path": path.name,
        "whole_process_samples": len(blocks),
        "whole_process_period": whole_period,
        "owner_qualified_samples": owner_samples,
        "owner_qualified_period": owner_period,
        "unknown_interior_samples": unknown_interior,
        "leaf_counts": dict(sorted(leaf_counts.items())),
        "leaf_period": dict(sorted(leaf_period.items())),
        "nested_counts": dict(sorted(nested_counts.items())),
        "frame_occurrences": dict(sorted(frame_occurrences.items())),
        "scanner_leaf_samples": leaf_counts[SCANNER],
        "scanner_leaf_period": leaf_period[SCANNER],
        "scanner_leaf_offsets": [
            {
                "offset_hex": hex(offset),
                "samples": scanner_offsets[offset],
                "period": scanner_offset_period[offset],
            }
            for offset in sorted(scanner_offsets)
        ],
    }


def verify_references(
    computed: list[dict[str, Any]],
    analysis: dict[str, Any],
    root_audit: dict[str, Any],
) -> None:
    require(root_audit.get("schema") == "litchi.performance.0811.root-audit.v1",
            "0811 root-audit schema changed")
    require(analysis.get("schema") == "litchi.performance.0811.current-production-analysis.v1",
            "0811 analysis schema changed")
    root_rows = root_audit.get("frames")
    require(isinstance(root_rows, list) and len(root_rows) == 2,
            "0811 root-audit frame row count changed")
    root_by_path: dict[str, dict[str, Any]] = {}
    for row in root_rows:
        require(isinstance(row, dict), "0811 root-audit frame row malformed")
        path = require_string(row.get("path"), "0811 root-audit frame path")
        require(path not in root_by_path, f"0811 root-audit duplicate frame row {path}")
        root_by_path[path] = row

    reports = analysis.get("perf", {}).get("reports")
    require(isinstance(reports, list) and len(reports) == 2,
            "0811 analysis perf report count changed")
    reports_by_repeat: dict[int, dict[str, Any]] = {}
    for report in reports:
        require(isinstance(report, dict), "0811 analysis perf report malformed")
        repeat = require_integer(report.get("repeat"), "0811 analysis repeat")
        require(repeat not in reports_by_repeat, f"0811 analysis duplicate repeat {repeat}")
        reports_by_repeat[repeat] = report

    require({row["path"] for row in computed} == {"0.frames.gz", "1.frames.gz"},
            "computed frame repeat set changed")
    for row in computed:
        path = row["path"]
        repeat = int(path.split(".", 1)[0])
        root = root_by_path.get(path)
        require(root is not None, f"0811 root-audit has no row for {path}")
        report = reports_by_repeat.get(repeat)
        require(report is not None, f"0811 analysis has no row for repeat {repeat}")
        stack = report.get("stack")
        require(isinstance(stack, dict), f"0811 analysis stack missing for repeat {repeat}")

        require(row["whole_process_samples"] == root.get("whole_process_samples"),
                f"{path}: whole-process count disagrees with root-audit")
        require(row["owner_qualified_samples"] == root.get("owner_samples"),
                f"{path}: owner count disagrees with root-audit")
        require(row["owner_qualified_period"] == root.get("owner_period"),
                f"{path}: owner period disagrees with root-audit")
        require(row["unknown_interior_samples"] == root.get("unknown_interior"),
                f"{path}: unknown-interior count disagrees with root-audit")
        require(row["leaf_counts"] == root.get("leaf_counts"),
                f"{path}: leaf counts disagree with root-audit")
        require(row["nested_counts"] == root.get("nested_counts"),
                f"{path}: nested counts disagree with root-audit")
        require(row["frame_occurrences"] == root.get("frame_occurrences"),
                f"{path}: frame occurrences disagree with root-audit")

        require(row["whole_process_samples"] == stack.get("whole_process_samples"),
                f"{path}: whole-process count disagrees with analysis")
        require(row["whole_process_period"] == stack.get("whole_process_period"),
                f"{path}: whole-process period disagrees with analysis")
        require(row["owner_qualified_samples"] == stack.get("owner_qualified_samples"),
                f"{path}: owner count disagrees with analysis")
        require(row["owner_qualified_period"] == stack.get("owner_qualified_period"),
                f"{path}: owner period disagrees with analysis")
        require(row["unknown_interior_samples"] == stack.get("unknown_interior_frames"),
                f"{path}: unknown-interior count disagrees with analysis")
        require(row["leaf_counts"] == dict(pair_list(stack.get("self_leaf_samples"),
                                                       f"{path}: analysis self leaves")),
                f"{path}: leaf counts disagree with analysis")
        require(row["leaf_period"] == dict(pair_list(stack.get("self_leaf_period"),
                                                       f"{path}: analysis self-leaf periods")),
                f"{path}: leaf periods disagree with analysis")
        require(row["scanner_leaf_samples"] == 217 + 37 * repeat,
                f"{path}: scanner leaf count is not the sealed expected count")


def verify_perf_inputs(
    build: dict[str, Any],
    build_descriptor: dict[str, Any],
    compression_by_kind: dict[tuple[int, str], dict[str, Any]],
    perf_receipts: list[Any],
    perf_receipts_descriptor: dict[str, Any],
    perf_complete: dict[str, Any],
    expected_binary: dict[str, Any],
) -> None:
    require(build.get("schema") == "litchi.performance.0811.build.v1", "0811 build schema changed")
    require(build.get("target") == str(Path(BINARY_PATH).parent), "0811 build target changed")
    require(perf_complete.get("schema") == "litchi.performance.0811.perf.complete.v1",
            "0811 perf-complete schema changed")
    require(perf_complete.get("event") == EVENT, "0811 perf event changed")
    require(perf_complete.get("call_graph") == "fp", "0811 perf call graph changed")
    require(perf_complete.get("frequency_hz") == 499, "0811 perf frequency changed")
    require(perf_complete.get("owner") == OWNER, "0811 perf owner changed")
    require(perf_complete.get("reports") == 2 and perf_complete.get("samples") == 200,
            "0811 perf sample contract changed")
    require(perf_complete.get("receipts", {}).get("sha256") == perf_receipts_descriptor["sha256"],
            "0811 perf receipt hash changed")
    require(perf_complete.get("build_sha256") == build_descriptor["sha256"],
            "0811 perf build hash changed")

    binaries = build.get("binaries")
    require(isinstance(binaries, dict), "0811 build binaries missing")
    fp = binaries.get("fp")
    require(isinstance(fp, dict), "0811 frame-pointer binary descriptor missing")
    require(fp == expected_binary, "0811 frame-pointer binary descriptor changed")

    require(isinstance(perf_receipts, list) and len(perf_receipts) == 2,
            "0811 perf receipt count changed")
    seen: set[int] = set()
    for index, receipt in enumerate(perf_receipts):
        require(isinstance(receipt, dict), f"perf receipt {index}: malformed")
        repeat = require_integer(receipt.get("repeat"), f"perf receipt {index}.repeat")
        require(repeat in (0, 1) and repeat not in seen, f"perf receipt {index}: repeat changed")
        seen.add(repeat)
        require(receipt.get("lane") == "perf", f"perf receipt {index}: lane changed")
        require(receipt.get("binary") == expected_binary, f"perf receipt {index}: binary changed")
        raw = receipt.get("raw")
        report = receipt.get("report")
        require(isinstance(raw, dict) and isinstance(report, dict),
                f"perf receipt {index}: artifact descriptors malformed")
        expected_raw = compression_by_kind[(repeat, "raw")]["decompressed"]
        require(raw == expected_raw, f"perf receipt {index}: raw descriptor changed")
        expected_report = PERF / f"{repeat}.json"
        require(report.get("path") == str(expected_report),
                f"perf receipt {index}: report path changed")
        sealed_report = artifact(expected_report)
        require(report.get("bytes") == sealed_report["bytes"]
                and report.get("sha256") == sealed_report["sha256"],
                f"perf receipt {index}: report identity changed")
        command = receipt.get("command")
        require(isinstance(command, list), f"perf receipt {index}: command missing")
        require(EVENT in command and "--call-graph" in command and "fp" in command,
                f"perf receipt {index}: cycles:u/fp command contract changed")
        require(BINARY_PATH in command, f"perf receipt {index}: DSO path changed")


def build_result() -> dict[str, Any]:
    seal_path = SEALED / "seal.json"
    seal = read_json(seal_path)
    seal_files = seal.get("files")
    require(isinstance(seal_files, dict), "0811 seal files are missing")

    build_path = SEALED / "build/build.json"
    build_descriptor = sealed_artifact(build_path, seal_files)
    build = read_json(build_path)
    expected_binary = build.get("binaries", {}).get("fp")
    require(isinstance(expected_binary, dict), "0811 frame-pointer binary descriptor missing")
    require(expected_binary.get("path") == BINARY_PATH, "0811 frame-pointer DSO path changed")
    require(require_integer(expected_binary.get("bytes"), "0811 frame-pointer DSO bytes", positive=True) > 0,
            "0811 frame-pointer DSO length is invalid")
    require(re.fullmatch(r"[0-9a-f]{64}", require_string(expected_binary.get("sha256"),
                                                         "0811 frame-pointer DSO sha256")),
            "0811 frame-pointer DSO hash is invalid")

    compression_path = PERF / "compression.json"
    compression_descriptor, retained_compression, compression_by_kind = verify_compression(
        compression_path,
        seal_files,
    )

    frame_receipts_path = PERF / "frame-receipts.json"
    frame_receipts_descriptor = sealed_artifact(frame_receipts_path, seal_files)
    frame_receipts = read_json(frame_receipts_path)
    expected_frame_receipts = [
        {
            "compressed": row["compressed"],
            "compression": row["compression"],
            "decompressed_sha256": row["decompressed"]["sha256"],
            "gzip_mtime": row["gzip_mtime"],
            "kind": row["kind"],
            "original": row["decompressed"],
            "repeat": row["repeat"],
        }
        for row in retained_compression
        if row["kind"] == "frames"
    ]
    require(isinstance(frame_receipts, list) and len(frame_receipts) == 2,
            "0811 frame receipt count changed")
    for index, row in enumerate(frame_receipts):
        expected = expected_frame_receipts[index]
        require(row.get("repeat") == expected["repeat"],
                f"frame receipt {index}: repeat changed")
        require(row.get("kind") == "frames", f"frame receipt {index}: kind changed")
        require(row.get("compression") == "gzip" and row.get("gzip_mtime") == 0,
                f"frame receipt {index}: compression changed")
        require(row.get("decompressed_sha256") == expected["decompressed_sha256"],
                f"frame receipt {index}: decompressed hash changed")
        require(row.get("compressed") == expected["compressed"],
                f"frame receipt {index}: compressed descriptor changed")

    perf_receipts_path = PERF / "receipts.json"
    perf_receipts_descriptor = sealed_artifact(perf_receipts_path, seal_files)
    perf_receipts = read_json(perf_receipts_path)
    perf_complete_path = PERF / "complete.json"
    perf_complete_descriptor = sealed_artifact(perf_complete_path, seal_files)
    perf_complete = read_json(perf_complete_path)
    verify_perf_inputs(
        build,
        build_descriptor,
        compression_by_kind,
        perf_receipts,
        perf_receipts_descriptor,
        perf_complete,
        expected_binary,
    )

    analysis_path = SEALED / "analysis.json"
    analysis_descriptor = sealed_artifact(analysis_path, seal_files)
    analysis = read_json(analysis_path)
    root_audit_path = SEALED / "root-audit.json"
    root_audit_descriptor = sealed_artifact(root_audit_path, seal_files)
    root_audit = read_json(root_audit_path)

    frame_paths = [PERF / "0.frames.gz", PERF / "1.frames.gz"]
    computed = [parse_frames(path, expected_binary) for path in frame_paths]
    verify_references(computed, analysis, root_audit)

    sealed_entries = {
        relative_root(path): seal_files[relative_root(path)]
        for path in (
            build_path,
            compression_path,
            frame_receipts_path,
            perf_receipts_path,
            perf_complete_path,
            analysis_path,
            root_audit_path,
            *frame_paths,
        )
    }
    rows = []
    for row in computed:
        frame_compression = compression_by_kind[(int(row["path"].split(".", 1)[0]), "frames")]
        rows.append(
            {
                "repeat": int(row["path"].split(".", 1)[0]),
                "frames": frame_compression,
                "whole_process_samples": row["whole_process_samples"],
                "whole_process_period": row["whole_process_period"],
                "owner_qualified_samples": row["owner_qualified_samples"],
                "owner_qualified_period": row["owner_qualified_period"],
                "unknown_interior_samples": row["unknown_interior_samples"],
                "scanner_leaf_symbol": SCANNER,
                "scanner_leaf_samples": row["scanner_leaf_samples"],
                "scanner_leaf_period": row["scanner_leaf_period"],
                "offsets": row["scanner_leaf_offsets"],
                "leaf_counts": row["leaf_counts"],
            }
        )

    return {
        "schema": "litchi.performance.0812.offset-audit.v1",
        "scope": (
            "Independent retained 0811 cycles:u scanner leaf offset census for the "
            "exact owner. Counts and periods are descriptive sampled observations; "
            "no assembly mapping, instruction cost, phase fraction, causal cost, "
            "speedup, or adoption claim is made."
        ),
        "event": EVENT,
        "owner": OWNER,
        "binary": expected_binary,
        "inputs": {
            "build": build_descriptor,
            "compression": compression_descriptor,
            "frame_receipts": frame_receipts_descriptor,
            "perf_receipts": perf_receipts_descriptor,
            "perf_complete": perf_complete_descriptor,
            "analysis": analysis_descriptor,
            "root_audit": root_audit_descriptor,
            "sealed_0811_entries": sealed_entries,
        },
        "rows": rows,
        "verification": {
            "compression_members_verified": True,
            "decompressed_lengths_and_hashes_verified": True,
            "sealed_0811_hashes_verified": True,
            "strict_cycles_u_headers": True,
            "exact_owner_verified": True,
            "scanner_leaf_counts_match_root_audit_and_analysis": True,
            "assembly_dependency": False,
        },
    }


def main() -> None:
    require(len(sys.argv) == 2 and sys.argv[1] in ("--write", "--check"),
            "usage: offset_audit.py --write|--check")
    value = build_result()
    encoded = json.dumps(value, indent=2, sort_keys=True) + "\n"
    if sys.argv[1] == "--write":
        require(not OUTPUT.exists(), f"immutable output already exists: {OUTPUT}")
        OUTPUT.write_text(encoded, encoding="utf-8")
        print("0812 retained scanner offset audit PASS")
    else:
        require(OUTPUT.is_file() and not OUTPUT.is_symlink(),
                f"missing immutable output: {OUTPUT}")
        require(OUTPUT.read_text(encoding="utf-8") == encoded,
                f"offset audit output changed: {OUTPUT}")
        print("0812 retained scanner offset audit --check PASS")


if __name__ == "__main__":
    try:
        main()
    except AuditError as error:
        raise SystemExit(f"offset audit FAILED: {error}") from error
