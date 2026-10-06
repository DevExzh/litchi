#!/usr/bin/env python3
"""Freeze, qualify, capture, reduce, and replay the prospective 0922 DOC comparison.

The comparison deliberately reuses the unchanged 0917 probe and report
validator.  This file owns only the version pairing, custody, and reduction
logic for the baseline and the isolated candidate.  It never builds Rust
code.  ``freeze`` is a one-way identity bind; ``run`` refuses to overwrite a
child; ``reduce`` refuses incomplete or failed schedules.
"""

from __future__ import annotations

import argparse
import datetime as dt
import gzip
import hashlib
import importlib.util
import json
import math
import os
from functools import lru_cache
from pathlib import Path
import random
import re
import statistics
import subprocess
import sys
import tempfile
from typing import Any, Iterable


HERE = Path(__file__).resolve().parent
# A prospective scratch freeze may copy this packet outside the checkout.
# Keep repository inputs separately bound to the actual checkout while allowing
# all generated packet outputs to live under an explicit external root.
REPO_ROOT = Path(os.environ.get("LITCHI_PERF_0922_REPO_ROOT", str(HERE.parents[3]))).resolve()
PROTOCOL_PATH = HERE / "comparison-protocol.json"
SCRIPT_PATH = HERE / "comparison.py"
REPLAY_PATH = HERE / "comparison-replay.py"
PACKET_ROOT = Path(os.environ.get("LITCHI_PERF_0922_PACKET_ROOT", str(HERE))).resolve()
MANIFEST_PATH = PACKET_ROOT / "comparison-manifest.json"
REDUCED_PATH = PACKET_ROOT / "comparison-reduced.json"
LOG_DIR = PACKET_ROOT / "comparison-logs"
RAW_DIR = PACKET_ROOT / "comparison-raw"
DEFAULT_EXTERNAL = Path(
    os.environ.get("LITCHI_PERF_0922_ROOT", "/tmp/litchi-performance-0922-tooling/comparison")
).resolve()
VALIDATOR_PATH = REPO_ROOT / "docs/performance/results/change-0917/capture.py"
PROBE_PATH = REPO_ROOT / "docs/performance/results/change-0917/probe"
MAX_PACKET_ARTIFACT = 1 << 20
HEX = set("0123456789abcdef")
VERSIONS = ("baseline", "candidate")
LANES = ("native", "allocator")
ROUTES = ("zero", "default", "release")
EDITS = (1, 2, 8, 32)
BASELINE_REVISION = "bb19ca74bd52044fbd2c1c7cc9875e975bb334c0"
# The 0922 source receipt contains only the real source/probe file map. The
# reviewed patch is a separate custody input and is never fabricated as a
# source-tree metadata row.
CANDIDATE_METADATA_FILES: set[str] = set()
# The candidate source is deliberately unresolved until the candidate author,
# reviewer, and final build receipt agree on one source tree.  ``freeze`` binds
# the concrete revision from that receipt; this runner must never invent one.
CURRENT_REVISION = None


class EvidenceError(RuntimeError):
    """A fail-closed evidence or custody error."""


def fail(message: str) -> None:
    raise EvidenceError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def now() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat()


def parse_utc(value: object, label: str) -> dt.datetime:
    require(isinstance(value, str), f"{label}: timestamp missing")
    try:
        parsed = dt.datetime.fromisoformat(value)
    except ValueError as error:
        fail(f"{label}: malformed timestamp: {error}")
    require(parsed.tzinfo is not None, f"{label}: timestamp has no timezone")
    return parsed.astimezone(dt.timezone.utc)


def sha256(path: Path) -> str:
    require(path.is_file() and not path.is_symlink(), f"missing regular file: {path}")
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def json_bytes(value: object) -> bytes:
    return (json.dumps(value, indent=2, sort_keys=True) + "\n").encode("utf-8")


def read_json(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing JSON: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(f"invalid JSON {path}: {error}")


def write_exclusive(path: Path, data: bytes) -> None:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    require(not path.exists(), f"refusing to replace retained file: {path}")
    temporary = path.with_name(path.name + f".tmp-{os.getpid()}")
    require(not temporary.exists(), f"temporary output already exists: {temporary}")
    try:
        with temporary.open("xb") as stream:
            stream.write(data)
        if path.exists():
            fail(f"retained output appeared during write: {path}")
        temporary.replace(path)
    except Exception:
        if temporary.exists():
            temporary.unlink()
        raise


def write_json(path: Path, value: object) -> None:
    data = json_bytes(value)
    require(len(data) < MAX_PACKET_ARTIFACT, f"packet JSON is at least 1 MiB: {path}")
    write_exclusive(path, data)


def digest(value: object, label: str) -> str:
    require(isinstance(value, str) and len(value) == 64 and set(value) <= HEX,
            f"{label}: invalid SHA-256")
    return value


def git_revision(value: object, label: str) -> str:
    """Validate the full Git object name used for a source binding.

    Git source identities in the 0922 protocol are full SHA-1 object names,
    not SHA-256 content digests.  Keeping this validator separate from
    :func:`digest` prevents a valid 40-hex revision from being rejected as a
    malformed artifact hash.
    """
    require(isinstance(value, str) and len(value) == 40 and set(value) <= HEX,
            f"{label}: invalid Git revision")
    return value


def integer(value: object, label: str, minimum: int = 0) -> int:
    require(type(value) is int and value >= minimum, f"{label}: invalid integer")
    return value


def descriptor(path: Path, *, allow_large: bool = True) -> dict[str, object]:
    path = Path(path).resolve()
    require(path.is_file() and not path.is_symlink(), f"missing regular file: {path}")
    size = path.stat().st_size
    if not allow_large:
        require(size < MAX_PACKET_ARTIFACT, f"packet artifact is at least 1 MiB: {path}")
    return {"path": str(path), "bytes": size, "sha256": sha256(path)}


def repository_relative(path: Path) -> str | None:
    try:
        return path.resolve().relative_to(REPO_ROOT).as_posix()
    except ValueError:
        return None


def packet_descriptor(path: Path, *, allow_large: bool = True) -> dict[str, object]:
    identity = descriptor(path, allow_large=allow_large)
    relative = repository_relative(path)
    if relative is not None:
        identity["path"] = relative
    return identity


def manifest_path(identity: object) -> Path:
    require(isinstance(identity, dict), "descriptor is not an object")
    raw = identity.get("path")
    require(isinstance(raw, str), "descriptor path missing")
    path = Path(raw)
    return path.resolve() if path.is_absolute() else (REPO_ROOT / path).resolve()


def verify_descriptor(identity: object, label: str, *, allow_missing: bool = False) -> Path:
    require(isinstance(identity, dict), f"{label}: descriptor missing")
    path = manifest_path(identity)
    if allow_missing and not path.exists():
        integer(identity.get("bytes"), f"{label}: bytes", 1)
        digest(identity.get("sha256"), f"{label}: hash")
        return path
    actual = descriptor(path)
    if isinstance(identity.get("path"), str) and not Path(identity["path"]).is_absolute():
        actual["path"] = repository_relative(path) or actual["path"]
    require(actual == identity, f"{label}: content changed")
    return path


def protocol() -> dict[str, Any]:
    value = read_json(PROTOCOL_PATH)
    require(isinstance(value, dict), "comparison protocol is not an object")
    require(value.get("schema") == "litchi.performance.0922.doc-comparison.protocol.v1",
            "unexpected 0922 protocol schema")
    require(value.get("status") == "prospective",
            "0922 protocol is only a prospective protocol until candidate review closes")
    require(value.get("candidate_source_head") is None,
            "unresolved 0922 protocol must not claim a candidate source head")
    require(value.get("baseline_source_head") == BASELINE_REVISION,
            "0922 baseline source identity changed")
    require(value.get("candidate_base_revision") == BASELINE_REVISION,
            "0922 candidate base revision must remain the retained baseline revision")
    fixtures = value.get("fixtures")
    require(isinstance(fixtures, list) and [row.get("case") for row in fixtures]
            == ["docnohf", "docfloat"], "fixture order changed")
    expected = {
        "docnohf": ("test-data/ole/doc/NoHeadFoot.doc", 26112,
                     "45e5df073f34314da6f39d2dad119fb2ef23470878fd2df67f632864cd92ea48"),
        "docfloat": ("test-data/ole/doc/FloatingPictures.doc", 335360,
                      "41cc0cd8f2d7390266f844f5bddbcd29d08ac5338abc33832d9c9a6ed5d0f85d"),
    }
    for row in fixtures:
        require(isinstance(row, dict), "malformed fixture row")
        case = row.get("case")
        require(case in expected, f"unknown fixture {case!r}")
        path, size, checksum = expected[case]
        require(row.get("path") == path and row.get("bytes") == size
                and row.get("sha256") == checksum, f"{case}: fixture identity changed")
    require(value.get("edits") == list(EDITS), "edit counts changed")
    require(value.get("retention_routes") == list(ROUTES), "route order changed")
    native = value.get("native")
    require(isinstance(native, dict), "native section missing")
    require(native.get("cpu") == 14 and native.get("rounds") == 12
            and native.get("retention") == "default"
            and native.get("warmups_per_child") == 5
            and native.get("samples_per_child") == 50
            and native.get("children") == 192
            and native.get("paired_cells") == 96, "native schedule changed")
    allocator = value.get("allocator")
    require(isinstance(allocator, dict), "allocator section missing")
    require(allocator.get("cpu") == 14 and allocator.get("repetitions") == 2
            and allocator.get("warmups_per_child") == 0
            and allocator.get("samples_per_child") == 1
            and allocator.get("children") == 96, "allocator schedule changed")
    bootstrap = value.get("bootstrap")
    require(isinstance(bootstrap, dict) and bootstrap.get("draws") == 10000
            and bootstrap.get("seed") == 922
            and bootstrap.get("minimum_lower_ratio") == 1 / 1.05,
            "bootstrap contract changed")
    flags = value.get("review_flags")
    require(isinstance(flags, dict) and flags.get("latency_increase_fraction") == 0.05
            and flags.get("peak_live_increase_fraction") == 0.05
            and flags.get("rss_increase_fraction") == 0.05
            and flags.get("allocation_increase_fraction") == 0.01,
            "review flag contract changed")
    gates = value.get("mandatory_gates")
    require(isinstance(gates, list) and len(gates) == 10, "mandatory gate list changed")
    parse_utc(value.get("created_utc"), "protocol created")
    return value


def candidate_scope_files(p: dict[str, Any]) -> set[str] | None:
    """Return the reviewed source scope, or None while candidate review is open."""
    value = p.get("candidate_source_files")
    if value is None:
        return None
    require(isinstance(value, list) and value,
            "candidate_source_files must be a nonempty reviewed list or null")
    result: set[str] = set()
    for name in value:
        require(isinstance(name, str) and not Path(name).is_absolute()
                and ".." not in Path(name).parts,
                "candidate_source_files contains an unsafe path")
        result.add(name)
    return result


def fixture_rows(p: dict[str, Any]) -> dict[str, dict[str, Any]]:
    result: dict[str, dict[str, Any]] = {}
    for row in p["fixtures"]:
        case = str(row["case"])
        path = (REPO_ROOT / str(row["path"])).resolve()
        require(path.is_file() and not path.is_symlink(), f"missing fixture: {path}")
        require(path.stat().st_size == row["bytes"] and sha256(path) == row["sha256"],
                f"fixture changed: {case}")
        result[case] = {**row, "path_abs": str(path)}
    return result


def tree_descriptors(root: Path) -> dict[str, dict[str, object]]:
    root = root.resolve()
    require(root.is_dir() and not root.is_symlink(), f"missing probe source: {root}")
    rows: dict[str, dict[str, object]] = {}
    for path in sorted(root.rglob("*")):
        if path.is_file() and not path.is_symlink():
            rows[path.relative_to(root).as_posix()] = {
                "bytes": path.stat().st_size,
                "sha256": sha256(path),
            }
    require(rows, f"probe source is empty: {root}")
    return rows


def verify_tree(root: Path, expected: object, label: str) -> None:
    require(isinstance(expected, dict) and expected, f"{label}: source map missing")
    require(tree_descriptors(root) == expected, f"{label}: source tree changed")


def fixture_custody(p: dict[str, Any]) -> dict[str, dict[str, object]]:
    return {
        case: {key: value for key, value in row.items() if key != "path_abs"}
        for case, row in fixture_rows(p).items()
    }


def git_head() -> str:
    try:
        result = subprocess.run(["git", "rev-parse", "HEAD"], cwd=REPO_ROOT,
                                check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                text=True)
    except (OSError, subprocess.SubprocessError) as error:
        fail(f"cannot read repository head: {error}")
    return result.stdout.strip()


def source_revision(receipt: Path) -> str | None:
    value = read_json(receipt)
    if not isinstance(value, dict):
        return None
    for key in ("source_revision", "revision", "base_revision"):
        if isinstance(value.get(key), str):
            return value[key]
    source = value.get("source")
    if isinstance(source, dict):
        for key in ("source_revision", "revision", "base_revision"):
            if isinstance(source.get(key), str):
                return source[key]
    return None


def build_binary_hash(receipt: Path, lane: str) -> str | None:
    """Find the named probe binary hash in a structured build receipt."""
    value = read_json(receipt)
    name = "doc_retention_probe_alloc" if lane == "allocator" else "doc_retention_probe"

    def walk(node: object) -> str | None:
        if isinstance(node, dict):
            for key in ("binaries", "artifacts"):
                binaries = node.get(key)
                if not isinstance(binaries, list):
                    continue
                for row in binaries:
                    if not isinstance(row, dict):
                        continue
                    candidate_name = row.get("name")
                    candidate_path = row.get("path")
                    if candidate_name == name or (isinstance(candidate_path, str)
                                                  and Path(candidate_path).name == name):
                        checksum = row.get("sha256")
                        if isinstance(checksum, str):
                            return digest(checksum, f"{receipt} {name} hash")
            for child in node.values():
                found = walk(child)
                if found is not None:
                    return found
        elif isinstance(node, list):
            for child in node:
                found = walk(child)
                if found is not None:
                    return found
        return None

    return walk(value)


def build_source_binding(receipt: Path, label: str) -> dict[str, str]:
    """Read the reviewed candidate build envelope's source identity.

    The fresh baseline build receipt intentionally contains only the build
    command and nested artifacts.  The candidate envelope must additionally
    bind the same base revision, reviewed patch hash, and source-tree manifest
    hash; matching executable bytes alone is not a source identity.
    """
    value = read_json(receipt)
    candidates: list[dict[str, object]] = []

    def walk(node: object) -> None:
        if isinstance(node, dict):
            source = node.get("source")
            if isinstance(source, dict):
                candidates.append(source)
            if any(key in node for key in
                   ("candidate_patch_sha256", "candidate_manifest_sha256")):
                candidates.append(node)
            for child in node.values():
                walk(child)
        elif isinstance(node, list):
            for child in node:
                walk(child)

    walk(value)
    for source in candidates:
        revision = source.get("revision", source.get("source_revision",
                                                    source.get("base_revision")))
        patch = source.get("candidate_patch_sha256")
        manifest = source.get("candidate_manifest_sha256")
        if isinstance(revision, str) and isinstance(patch, str) and isinstance(manifest, str):
            return {
                "revision": git_revision(revision, f"{label}: source revision"),
                "candidate_patch_sha256": digest(patch, f"{label}: candidate patch hash"),
                "candidate_manifest_sha256": digest(manifest,
                                                      f"{label}: candidate manifest hash"),
            }
    fail(f"{label}: build receipt lacks reviewed source binding")
    return {}  # unreachable; keeps static type checkers satisfied


def candidate_patch_identity(receipt: Path, patch: Path | None = None) -> dict[str, object] | None:
    value = read_json(receipt)
    if isinstance(value, dict):
        for key in ("candidate_patch", "patch", "candidate_diff", "patch_descriptor"):
            candidate = value.get(key)
            if isinstance(candidate, dict) and isinstance(candidate.get("sha256"), str):
                digest(candidate["sha256"], f"candidate receipt {key} hash")
                size = integer(candidate.get("bytes"), f"candidate receipt {key} bytes", 1)
                return {"candidate_patch": {"bytes": size, "sha256": candidate["sha256"]}}
        candidate_hash = value.get("candidate_patch_sha256")
        if isinstance(candidate_hash, str):
            candidate_hash = digest(candidate_hash, "candidate receipt candidate_patch_sha256")
            require(patch is not None, "candidate patch path is required for candidate_patch_sha256")
            patch_path = Path(patch).resolve()
            require(patch_path.is_file() and not patch_path.is_symlink()
                    and patch_path.stat().st_size > 0,
                    "candidate patch must be a nonempty regular file")
            require(sha256(patch_path) == candidate_hash,
                    "candidate patch file differs from candidate_patch_sha256")
            rows = value.get("files")
            if isinstance(rows, list):
                patch_rows = [row for row in rows
                              if isinstance(row, dict)
                              and row.get("path") == "docs/performance/results/change-0922/candidate.patch"]
                if patch_rows:
                    require(len(patch_rows) == 1,
                            "candidate source receipt has duplicate candidate.patch rows")
                    patch_row = patch_rows[0]
                    require(patch_row.get("bytes") == patch_path.stat().st_size
                            and patch_row.get("sha256") == candidate_hash,
                            "candidate source receipt candidate.patch row differs from supplied patch")
            return {"candidate_patch": {
                "bytes": patch_path.stat().st_size,
                "sha256": candidate_hash,
            }}
        source_hashes = value.get("source_hashes")
        if isinstance(source_hashes, dict) and isinstance(source_hashes.get("candidate_patch"), str):
            digest(source_hashes["candidate_patch"], "candidate source_hashes patch")
            if patch is not None:
                return {"candidate_patch": {
                    "bytes": patch.stat().st_size,
                    "sha256": source_hashes["candidate_patch"],
                }}
    return None


def source_file_map(receipt: Path, label: str) -> dict[str, str]:
    value = read_json(receipt)
    require(isinstance(value, dict), f"{label}: source receipt is not an object")
    files = value.get("source_files")
    if isinstance(files, dict):
        require(files, f"{label}: source receipt has an empty source_files map")
        result: dict[str, str] = {}
        for name, checksum in files.items():
            require(isinstance(name, str) and isinstance(checksum, str),
                    f"{label}: malformed source file identity")
            require(not Path(name).is_absolute() and ".." not in Path(name).parts,
                    f"{label}: unsafe source file path: {name}")
            result[name] = digest(checksum, f"{label}: {name}")
        return result

    # Candidate source receipts may record bytes and hashes as a list. Keep
    # every real source/probe row in the map; patch custody is handled by the
    # separate candidate-patch input.
    rows = value.get("files")
    require(isinstance(rows, list) and rows,
            f"{label}: source receipt lacks source_files or files map")
    result = {}
    for row in rows:
        require(isinstance(row, dict), f"{label}: malformed source file row")
        name = row.get("path")
        require(isinstance(name, str) and not Path(name).is_absolute()
                and ".." not in Path(name).parts,
                f"{label}: unsafe source file path: {name}")
        require(name not in result, f"{label}: duplicate source file row: {name}")
        integer(row.get("bytes"), f"{label}: {name} bytes", 0)
        result[name] = digest(row.get("sha256"), f"{label}: {name}")
    declared_count = value.get("candidate_file_count")
    if declared_count is not None:
        require(declared_count == len(result),
                f"{label}: candidate file count does not match files list")
    extras = value.get("extra_files", [])
    require(isinstance(extras, list) and all(isinstance(name, str) for name in extras),
            f"{label}: malformed candidate extra_files list")
    extra_set = set(extras)
    require(extra_set <= set(result), f"{label}: extra_files contains an unknown path")
    missing = value.get("missing_files")
    unexpected = value.get("unexpected_mismatches")
    require(missing in (None, []) and unexpected in (None, []),
            f"{label}: candidate source receipt reports missing or unexpected files")
    for name in CANDIDATE_METADATA_FILES:
        result.pop(name, None)
    require(result, f"{label}: candidate source map is empty after metadata removal")
    return result


def candidate_source_delta(receipt: Path, label: str) -> dict[str, object] | None:
    """Return the reviewed candidate receipt's explicit source-tree delta."""
    value = read_json(receipt)
    if not isinstance(value, dict) or not isinstance(value.get("files"), list):
        return None
    modified = value.get("modified_files")
    extras = value.get("extra_files", [])
    manifest_hash = value.get("candidate_manifest_sha256",
                              value.get("source_manifest_sha256"))
    require(isinstance(modified, list) and all(isinstance(name, str) for name in modified),
            f"{label}: modified_files list missing")
    require(isinstance(extras, list) and all(isinstance(name, str) for name in extras),
            f"{label}: extra_files list missing")
    require(isinstance(manifest_hash, str), f"{label}: candidate manifest hash missing")
    return {
        "modified_files": sorted(set(modified)),
        "extra_files": sorted(set(extras)),
        "metadata_files": sorted(set(extras) & CANDIDATE_METADATA_FILES),
        "candidate_manifest_sha256": digest(manifest_hash,
                                               f"{label}: candidate manifest hash"),
    }


def copy_receipt_to_packet(source: Path, target_name: str, label: str) -> dict[str, object]:
    original = descriptor(source, allow_large=False)
    target = PACKET_ROOT / "comparison-receipts" / target_name
    write_exclusive(target, source.read_bytes())
    copied = packet_descriptor(target, allow_large=False)
    require(copied["bytes"] == original["bytes"]
            and copied["sha256"] == original["sha256"],
            f"{label}: receipt changed while copying")
    return {"descriptor": copied, "external": original,
            "source_revision": source_revision(source)}


def copy_file_to_packet(source: Path, target_name: str, label: str) -> dict[str, object]:
    original = descriptor(source, allow_large=False)
    target = PACKET_ROOT / "comparison-receipts" / target_name
    write_exclusive(target, source.read_bytes())
    copied = packet_descriptor(target, allow_large=False)
    require(copied["bytes"] == original["bytes"]
            and copied["sha256"] == original["sha256"],
            f"{label}: file changed while copying")
    return {"descriptor": copied, "external": original}


def binary_descriptor(path: Path, label: str) -> dict[str, object]:
    identity = descriptor(path)
    try:
        path.resolve().relative_to(REPO_ROOT)
    except ValueError:
        pass
    else:
        fail(f"{label}: compiled binary must stay outside checkout")
    return identity


def protocol_descriptor() -> dict[str, object]:
    return packet_descriptor(PROTOCOL_PATH, allow_large=False)


def script_descriptor(path: Path) -> dict[str, object]:
    return packet_descriptor(path, allow_large=False)


def safe_label(value: str) -> str:
    require(re.fullmatch(r"[a-z0-9-]+", value) is not None, f"unsafe child label: {value}")
    return value


def schedule(p: dict[str, Any], lane: str) -> list[dict[str, object]]:
    require(lane in LANES, f"unknown lane: {lane}")
    fixtures = [str(row["case"]) for row in p["fixtures"]]
    edits = [int(value) for value in p["edits"]]
    result: list[dict[str, object]] = []
    if lane == "native":
        for round_no in range(1, int(p["native"]["rounds"]) + 1):
            cases = fixtures if round_no % 2 else list(reversed(fixtures))
            counts = edits if round_no % 2 else list(reversed(edits))
            versions = list(VERSIONS) if round_no % 2 else list(reversed(VERSIONS))
            for case in cases:
                for count in counts:
                    pair_id = f"native-r{round_no:02d}-{case}-e{count:02d}"
                    for version in versions:
                        label = f"{pair_id}-{version}-default"
                        result.append({
                            "lane": lane, "round": round_no, "case": case,
                            "edits": count, "retention": "default", "version": version,
                            "pair_id": pair_id, "label": label,
                        })
    else:
        cells = [{"case": case, "edits": count, "retention": route}
                 for case in fixtures for count in edits for route in ROUTES]
        for repetition in (1, 2):
            ordered = cells if repetition == 1 else list(reversed(cells))
            versions = list(VERSIONS) if repetition == 1 else list(reversed(VERSIONS))
            for cell in ordered:
                pair_id = (f"allocator-r{repetition:02d}-{cell['case']}-e{cell['edits']:02d}-"
                           f"{cell['retention']}")
                for version in versions:
                    label = f"{pair_id}-{version}"
                    result.append({
                        "lane": lane, "repetition": repetition, "case": cell["case"],
                        "edits": cell["edits"], "retention": cell["retention"],
                        "version": version, "pair_id": pair_id, "label": label,
                    })
    expected = 192 if lane == "native" else 96
    require(len(result) == expected, f"{lane}: schedule length is {len(result)}, expected {expected}")
    require(len({str(row["label"]) for row in result}) == expected,
            f"{lane}: schedule labels are not unique")
    return result


def all_schedule(p: dict[str, Any]) -> list[dict[str, object]]:
    return schedule(p, "native") + schedule(p, "allocator")


def command_for(manifest: dict[str, Any], spec: dict[str, object]) -> list[str]:
    p = protocol()
    version = str(spec["version"])
    lane = str(spec["lane"])
    binary = manifest["versions"][version]["binaries"][lane]
    binary_path = str(binary["path"])
    fixture = fixture_rows(p)[str(spec["case"])]
    warmups = int(p[lane]["warmups_per_child"])
    samples = int(p[lane]["samples_per_child"])
    return [
        "/usr/bin/taskset", "-c", str(p[lane]["cpu"]), binary_path,
        "--case", str(spec["case"]), "--input", str(fixture["path_abs"]),
        "--edits", str(spec["edits"]), "--retention", str(spec["retention"]),
        "--warmup", str(warmups), "--samples", str(samples),
    ]


def process_command(command: list[str], process_log: Path) -> list[str]:
    return ["/usr/bin/time", "-v", "-o", str(process_log), *command]


def parse_time_v_text(text: str, label: str) -> dict[str, object]:
    patterns = {
        "user_seconds": r"^\s*User time \(seconds\):\s*([0-9.]+)$",
        "system_seconds": r"^\s*System time \(seconds\):\s*([0-9.]+)$",
        "max_rss_kib": r"^\s*Maximum resident set size \(kbytes\):\s*(\d+)$",
        "elapsed": r"^\s*Elapsed \(wall clock\) time \(h:mm:ss or m:ss\):\s*(\S+)$",
        "exit_status": r"^\s*Exit status:\s*(\d+)$",
    }
    values: dict[str, object] = {}
    for key, pattern in patterns.items():
        match = re.search(pattern, text, re.MULTILINE)
        if match is not None:
            values[key] = float(match.group(1)) if key.endswith("seconds") else (
                int(match.group(1)) if key in {"max_rss_kib", "exit_status"} else match.group(1))
    require("max_rss_kib" in values and "exit_status" in values,
            f"time log lacks required metrics: {label}")
    values["path"] = label
    return values


def parse_time_v(path: Path) -> dict[str, object]:
    return parse_time_v_text(path.read_text(encoding="utf-8", errors="strict"), str(path))


def report_json_from_stdout(path: Path) -> tuple[dict[str, Any] | None, str | None]:
    try:
        text = path.read_text(encoding="utf-8")
    except (OSError, UnicodeError) as error:
        return None, f"stdout cannot be decoded: {error}"
    for line in reversed([line.strip() for line in text.splitlines() if line.strip()]):
        try:
            value = json.loads(line)
        except json.JSONDecodeError:
            continue
        if isinstance(value, dict):
            return value, None
    return None, "stdout contains no JSON object"


@lru_cache(maxsize=2)
def validator_module(path: Path = VALIDATOR_PATH):
    spec = importlib.util.spec_from_file_location("litchi_perf_0917_validator", path)
    require(spec is not None and spec.loader is not None, f"cannot load validator: {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def validate_probe_report(report: Any, child: dict[str, object], *, warmups: int,
                          samples: int, validator_path: Path = VALIDATOR_PATH) -> dict[str, object]:
    module = validator_module(validator_path)
    p = module.protocol()
    fixture = module.fixture_rows(p)[str(child["case"])]
    summary = module.validate_report(
        report, {"lane": child["lane"], "case": child["case"],
                 "edits": child["edits"], "retention": child["retention"],
                 "label": child["label"]}, fixture, str(child["lane"]),
        expected_warmups=warmups, expected_samples=samples,
    )
    require(isinstance(summary, dict), f"{child['label']}: validator returned no summary")
    return summary


def artifact_for_read(item: object, label: str) -> Path:
    require(isinstance(item, dict), f"{label}: artifact descriptor missing")
    storage = item.get("storage")
    identity = item.get("descriptor")
    require(storage in {"packet", "external"} and isinstance(identity, dict),
            f"{label}: malformed artifact descriptor")
    path = verify_descriptor(identity, label)
    compressed = item.get("compressed", identity)
    uncompressed = item.get("uncompressed", item.get("external"))
    require(isinstance(compressed, dict) and isinstance(uncompressed, dict),
            f"{label}: compressed/uncompressed descriptors missing")
    require(item.get("compression") == "gzip", f"{label}: unknown compression")
    require(compressed == identity, f"{label}: descriptor is not compressed identity")
    digest(uncompressed.get("sha256"), f"{label}: uncompressed hash")
    integer(uncompressed.get("bytes"), f"{label}: uncompressed bytes")
    verify_descriptor(compressed, f"{label}: compressed artifact")
    with gzip.open(manifest_path(compressed), "rb") as stream:
        data = stream.read()
    require(len(data) == uncompressed.get("bytes")
            and sha256_bytes(data) == uncompressed.get("sha256"),
            f"{label}: compressed artifact does not reproduce raw source")
    return path


def retain_artifact(source: Path, packet_target: Path) -> dict[str, object]:
    original = descriptor(source)
    data = source.read_bytes()
    compressed = gzip.compress(data, compresslevel=9, mtime=0)
    # Keep the uncompressed raw file in the external capture root and retain
    # only a deterministic gzip in the packet.  This applies to every child
    # JSON/log, including small ones, so the committed packet cannot grow
    # with the 9,600 native samples.
    packet_gzip = packet_target.with_name(packet_target.name + ".gz")
    if len(compressed) < MAX_PACKET_ARTIFACT:
        write_exclusive(packet_gzip, compressed)
        compressed_identity = packet_descriptor(packet_gzip, allow_large=False)
        storage = "packet"
    else:
        gzip_target = source.with_name(source.name + ".gz")
        require(not gzip_target.exists(), f"refusing to replace gzip artifact: {gzip_target}")
        write_exclusive(gzip_target, compressed)
        compressed_identity = descriptor(gzip_target)
        storage = "external"
    return {
        "storage": storage,
        "descriptor": compressed_identity,
        "compressed": compressed_identity,
        "uncompressed": original,
        "external": original,
        "compression": "gzip",
    }


def manifest_descriptor() -> dict[str, object]:
    return packet_descriptor(MANIFEST_PATH, allow_large=False)


def freeze(args: argparse.Namespace) -> None:
    require(not MANIFEST_PATH.exists(), f"comparison manifest already exists: {MANIFEST_PATH}")
    p = protocol()
    external = Path(args.external_root).resolve() if args.external_root else DEFAULT_EXTERNAL
    require(not external.exists() or not any(external.iterdir()),
            f"external comparison root is not empty: {external}")
    external.mkdir(parents=True, exist_ok=True)
    fixture_custody(p)
    probe_root = Path(args.probe_source).resolve() if args.probe_source else PROBE_PATH
    validator_path = Path(args.validator).resolve() if args.validator else VALIDATOR_PATH
    require(probe_root == PROBE_PATH.resolve(), "comparison must use the unchanged 0917 probe")
    require(validator_path == VALIDATOR_PATH.resolve(), "comparison must use the 0917 validator")
    probe = {"path": repository_relative(probe_root) or str(probe_root),
             "files": tree_descriptors(probe_root)}
    source_receipts: dict[str, dict[str, object]] = {}
    build_receipts: dict[str, dict[str, object]] = {}
    candidate_patch: dict[str, object] | None = None
    candidate_patch_path = (Path(args.candidate_patch).resolve() if args.candidate_patch
                            else PACKET_ROOT / "candidate.patch").resolve()
    candidate_patch_file = copy_file_to_packet(candidate_patch_path, "candidate.patch",
                                               "candidate patch")
    for version, source_arg, build_arg in (
        ("baseline", args.baseline_source_receipt, args.baseline_build_receipt),
        ("candidate", args.candidate_source_receipt, args.candidate_build_receipt),
    ):
        source_path = Path(source_arg).resolve()
        build_path = Path(build_arg).resolve()
        source_receipts[version] = copy_receipt_to_packet(
            source_path, f"{version}-source.json", f"{version} source receipt")
        build_receipts[version] = copy_receipt_to_packet(
            build_path, f"{version}-build.json", f"{version} build receipt")
    require(source_receipts["baseline"]["source_revision"] == BASELINE_REVISION,
            "baseline source receipt does not bind the retained base revision")
    candidate_revision = source_receipts["candidate"].get("source_revision")
    require(candidate_revision == BASELINE_REVISION,
            "candidate must bind the retained base revision plus an explicit reviewed patch")
    baseline_files = source_file_map(Path(args.baseline_source_receipt).resolve(),
                                     "baseline source")
    candidate_files = source_file_map(Path(args.candidate_source_receipt).resolve(),
                                      "candidate source")
    candidate_scope = candidate_scope_files(p)
    require(candidate_scope is not None,
            "freeze requires candidate_source_files from completed candidate review")
    missing_files = set(baseline_files) - set(candidate_files)
    added_files = set(candidate_files) - set(baseline_files)
    require(not missing_files,
            f"candidate source is missing baseline files: {sorted(missing_files)}")
    changed_files = {name for name in baseline_files
                     if baseline_files[name] != candidate_files[name]}
    require(changed_files or added_files,
            "candidate source receipt does not bind a source delta")
    if candidate_scope is not None:
        require(added_files <= candidate_scope,
                f"candidate source added files outside reviewed scope: {sorted(added_files - candidate_scope)}")
        require(changed_files <= candidate_scope,
                f"candidate source changes are outside reviewed scope: {sorted(changed_files - candidate_scope)}")
        require(candidate_scope & (changed_files | added_files),
                "candidate source receipt does not bind the reviewed implementation files")
    candidate_delta = candidate_source_delta(
        Path(args.candidate_source_receipt).resolve(), "candidate source")
    require(candidate_delta is not None,
            "candidate source receipt must use the reviewed source-tree envelope")
    if candidate_delta is not None:
        require(set(candidate_delta["modified_files"]) == changed_files,
                "candidate source modified_files does not match source hashes")
        require(set(candidate_delta["extra_files"]) == (added_files | set(candidate_delta["metadata_files"])),
                "candidate source extra_files does not match source hashes")
        require(set(candidate_delta["metadata_files"]) == CANDIDATE_METADATA_FILES,
                "candidate source metadata custody set changed")
    candidate_patch_receipt = candidate_patch_identity(
        Path(args.candidate_source_receipt).resolve(), candidate_patch_path)
    require(candidate_patch_receipt is not None,
            "candidate source receipt does not bind an explicit nonempty patch descriptor")
    build_receipts["baseline"]["source_revision"] = BASELINE_REVISION
    build_receipts["candidate"]["source_revision"] = BASELINE_REVISION
    require(candidate_patch_receipt["candidate_patch"] == {
        "bytes": candidate_patch_file["descriptor"]["bytes"],
        "sha256": candidate_patch_file["descriptor"]["sha256"],
    },
            "candidate patch receipt identity differs from supplied patch")
    candidate_build_binding = build_source_binding(
        Path(args.candidate_build_receipt).resolve(), "candidate build")
    require(candidate_build_binding["revision"] == BASELINE_REVISION
            and candidate_build_binding["candidate_patch_sha256"]
            == candidate_patch_receipt["candidate_patch"]["sha256"]
            and candidate_build_binding["candidate_manifest_sha256"]
            == candidate_delta["candidate_manifest_sha256"],
            "candidate build source binding differs from reviewed source receipt")
    candidate_patch = candidate_patch_file
    candidate_patch["receipt"] = candidate_patch_receipt
    binaries: dict[str, dict[str, dict[str, object]]] = {version: {} for version in VERSIONS}
    binary_args = {
        "baseline": {"native": args.baseline_native, "allocator": args.baseline_allocator},
        "candidate": {"native": args.candidate_native, "allocator": args.candidate_allocator},
    }
    build_receipt_paths = {
        "baseline": Path(args.baseline_build_receipt).resolve(),
        "candidate": Path(args.candidate_build_receipt).resolve(),
    }
    for version in VERSIONS:
        for lane in LANES:
            identity = binary_descriptor(Path(binary_args[version][lane]).resolve(),
                                         f"{version}/{lane} binary")
            receipt_hash = build_binary_hash(build_receipt_paths[version], lane)
            require(receipt_hash is not None and receipt_hash == identity["sha256"],
                    f"{version}/{lane}: build receipt binary hash does not match executable")
            binaries[version][lane] = identity
    versions = {}
    for version in VERSIONS:
        versions[version] = {
            "source_receipt": source_receipts[version],
            "build_receipt": build_receipts[version],
            "binaries": binaries[version],
        }
    manifest = {
        "schema": "litchi.performance.0922.doc-comparison.manifest.v1",
        "created_utc": now(),
        "repository_root": str(REPO_ROOT.resolve()),
        "candidate_source_head": None,
        "candidate_base_revision": candidate_revision,
        "candidate_source_delta": {
            "modified_files": sorted(changed_files),
            "added_source_files": sorted(added_files),
            "metadata_files": (candidate_delta["metadata_files"]
                                if candidate_delta is not None else []),
            "candidate_manifest_sha256": (candidate_delta["candidate_manifest_sha256"]
                                           if candidate_delta is not None else None),
        },
        "baseline_source_head": p["baseline_source_head"],
        "candidate_patch": candidate_patch,
        "protocol": protocol_descriptor(),
        "comparison_script": script_descriptor(SCRIPT_PATH),
        "replay_script": script_descriptor(REPLAY_PATH),
        "validator": script_descriptor(validator_path),
        "probe_source": probe,
        "fixtures": fixture_custody(p),
        "versions": versions,
        "external_root": str(external),
        "protocol_sha256": sha256(PROTOCOL_PATH),
        "schedule_counts": {"native": 192, "allocator": 96},
    }
    write_json(MANIFEST_PATH, manifest)
    print(json.dumps({"status": "frozen", "manifest": manifest_descriptor()}, sort_keys=True))


def load_manifest() -> dict[str, Any]:
    value = read_json(MANIFEST_PATH)
    require(isinstance(value, dict)
            and value.get("schema") == "litchi.performance.0922.doc-comparison.manifest.v1",
            "comparison manifest schema mismatch")
    p = protocol()
    candidate_base_revision = value.get("candidate_base_revision")
    require(candidate_base_revision == p["baseline_source_head"]
            and value.get("baseline_source_head") == p["baseline_source_head"],
            "manifest source lineage changed")
    require(value.get("protocol_sha256") == sha256(PROTOCOL_PATH),
            "protocol changed after freeze")
    verify_descriptor(value.get("protocol"), "manifest protocol")
    verify_descriptor(value.get("comparison_script"), "manifest comparison script")
    verify_descriptor(value.get("replay_script"), "manifest replay script")
    verify_descriptor(value.get("validator"), "manifest validator")
    parse_utc(value.get("created_utc"), "manifest created")
    probe_raw = value["probe_source"].get("path") if isinstance(value["probe_source"], dict) else None
    require(isinstance(probe_raw, str), "manifest probe path missing")
    probe_path = Path(probe_raw)
    if not probe_path.is_absolute():
        probe_path = (REPO_ROOT / probe_path).resolve()
    verify_tree(probe_path, value["probe_source"]["files"], "manifest probe")
    require(value.get("fixtures") == fixture_custody(p), "manifest fixture custody changed")
    versions = value.get("versions")
    require(isinstance(versions, dict) and set(versions) == set(VERSIONS),
            "manifest version set changed")
    candidate_patch_record = value.get("candidate_patch")
    require(isinstance(candidate_patch_record, dict)
            and isinstance(candidate_patch_record.get("descriptor"), dict)
            and isinstance(candidate_patch_record.get("receipt"), dict),
            "manifest candidate patch identity missing")
    verify_descriptor(candidate_patch_record["descriptor"], "manifest candidate patch")
    for version in VERSIONS:
        row = versions[version]
        require(isinstance(row, dict), f"manifest {version} row malformed")
        for key in ("source_receipt", "build_receipt"):
            receipt = row.get(key)
            require(isinstance(receipt, dict), f"manifest {version} {key} missing")
            receipt_path = verify_descriptor(receipt.get("descriptor"), f"manifest {version} {key}")
            if key == "source_receipt":
                require(source_revision(receipt_path) == row[key].get("source_revision"),
                        f"manifest {version} {key} source revision changed")
            else:
                # The raw build receipt may be a command envelope without a
                # revision field.  Its binding is the exact source receipt,
                # candidate patch, binary hashes, and this manifest.
                require(row[key].get("source_revision") == BASELINE_REVISION,
                        f"manifest {version} build binding changed")
            if version == "candidate":
                require(row[key].get("source_revision") == candidate_base_revision,
                        f"manifest candidate {key} is not bound to candidate_base_revision")
            if version == "candidate" and key == "source_receipt":
                require(candidate_patch_identity(
                    receipt_path, manifest_path(candidate_patch_record["descriptor"]))
                    == candidate_patch_record["receipt"],
                        "manifest candidate patch identity changed")
        binaries = row.get("binaries")
        require(isinstance(binaries, dict) and set(binaries) == set(LANES),
                f"manifest {version} binary set changed")
        build_path = manifest_path(row["build_receipt"]["descriptor"])
        for lane in LANES:
            verify_descriptor(binaries[lane], f"manifest {version}/{lane} binary",
                              allow_missing=True)
            require(build_binary_hash(build_path, lane) == binaries[lane].get("sha256"),
                    f"manifest {version}/{lane}: build receipt hash changed")
    baseline_source_path = manifest_path(versions["baseline"]["source_receipt"]["descriptor"])
    candidate_source_path = manifest_path(versions["candidate"]["source_receipt"]["descriptor"])
    baseline_files = source_file_map(baseline_source_path, "manifest baseline source")
    candidate_files = source_file_map(candidate_source_path, "manifest candidate source")
    candidate_scope = candidate_scope_files(p)
    require(candidate_scope is not None,
            "frozen manifest requires reviewed candidate_source_files")
    missing_files = set(baseline_files) - set(candidate_files)
    added_files = set(candidate_files) - set(baseline_files)
    require(not missing_files, "manifest candidate source inventory changed")
    changed_files = {name for name in baseline_files
                     if baseline_files[name] != candidate_files[name]}
    require(changed_files or added_files, "manifest candidate source delta changed")
    if candidate_scope is not None:
        require(added_files <= candidate_scope and changed_files <= candidate_scope,
                "manifest candidate source delta changed")
    candidate_delta = candidate_source_delta(candidate_source_path,
                                              "manifest candidate source")
    recorded_delta = value.get("candidate_source_delta")
    require(isinstance(recorded_delta, dict), "manifest candidate source delta missing")
    require(recorded_delta.get("modified_files") == sorted(changed_files)
            and recorded_delta.get("added_source_files") == sorted(added_files),
            "manifest candidate source delta does not match source receipts")
    if candidate_delta is not None:
        require(recorded_delta.get("metadata_files") == candidate_delta["metadata_files"]
                and recorded_delta.get("candidate_manifest_sha256")
                == candidate_delta["candidate_manifest_sha256"],
                "manifest candidate source metadata custody changed")
    candidate_build_path = manifest_path(versions["candidate"]["build_receipt"]["descriptor"])
    candidate_build_binding = build_source_binding(candidate_build_path,
                                                    "manifest candidate build")
    candidate_patch_hash = candidate_patch_record["receipt"]["candidate_patch"].get("sha256")
    require(candidate_build_binding["revision"] == BASELINE_REVISION
            and candidate_build_binding["candidate_patch_sha256"] == candidate_patch_hash
            and candidate_build_binding["candidate_manifest_sha256"]
            == recorded_delta.get("candidate_manifest_sha256"),
            "manifest candidate build source binding changed")
    external = value.get("external_root")
    require(isinstance(external, str), "manifest external root missing")
    return value


def child_paths(spec: dict[str, object], external: Path) -> dict[str, Path]:
    label = safe_label(str(spec["label"]))
    child = external / "children" / label
    return {
        "child": child,
        "stdout": child / "stdout.jsonl",
        "stderr": child / "stderr.log",
        "process": child / "process.time",
    }


def read_artifact_bytes(item: object, label: str) -> bytes:
    path = artifact_for_read(item, label)
    if isinstance(item, dict) and item.get("compressed") is not None:
        with gzip.open(manifest_path(item["compressed"]), "rb") as stream:
            return stream.read()
    return path.read_bytes()


def schedule_prior(schedule_rows: list[dict[str, object]], spec: dict[str, object]) -> Iterable[dict[str, object]]:
    index = next(i for i, row in enumerate(schedule_rows) if row["label"] == spec["label"])
    return schedule_rows[:index]


def receipt_path(spec: dict[str, object]) -> Path:
    return LOG_DIR / f"{safe_label(str(spec['label']))}.json"


def retained_receipt_artifact(source: Path, label: str) -> dict[str, object]:
    target = RAW_DIR / safe_label(label) / source.name
    return retain_artifact(source, target)


def environment_snapshot() -> dict[str, object]:
    try:
        affinity = sorted(os.sched_getaffinity(0))
    except (AttributeError, OSError):
        affinity = None
    try:
        load = list(os.getloadavg())
    except OSError:
        load = None
    return {"cwd": str(REPO_ROOT), "locale": os.environ.get("LC_ALL"),
            "affinity": affinity, "loadavg": load}


def run_one(manifest: dict[str, Any], spec: dict[str, object], schedule_rows: list[dict[str, object]]) -> None:
    p = protocol()
    external = Path(manifest["external_root"]).resolve()
    paths = child_paths(spec, external)
    receipt = receipt_path(spec)
    require(not receipt.exists(), f"refusing to replace receipt: {receipt}")
    require(not paths["child"].exists(), f"refusing to replace child directory: {paths['child']}")
    paths["child"].mkdir(parents=True, exist_ok=False)
    command = command_for(manifest, spec)
    wrapped = process_command(command, paths["process"])
    env = {"PATH": os.environ.get("PATH", ""), "LC_ALL": "C",
           "RUST_BACKTRACE": "0", "RUSTUP_TOOLCHAIN": os.environ.get("RUSTUP_TOOLCHAIN", "")}
    started = now()
    start_environment = environment_snapshot()
    with paths["stdout"].open("xb") as stdout, paths["stderr"].open("xb") as stderr:
        try:
            completed = subprocess.run(wrapped, cwd=REPO_ROOT, env=env, stdin=subprocess.DEVNULL,
                                       stdout=stdout, stderr=stderr, check=False)
        except OSError as error:
            completed = None
            error_text = f"runner could not start child: {error}\n".encode()
            stderr.write(error_text)
    finished = now()
    end_environment = environment_snapshot()
    status = "passed" if completed is not None and completed.returncode == 0 else "failed"
    receipt_value: dict[str, object] = {
        "schema": "litchi.performance.0922.child-receipt.v1",
        "status": status,
        "lane": spec["lane"], "version": spec["version"], "pair_id": spec["pair_id"],
        "spec": spec, "protocol": protocol_descriptor(),
        "manifest": manifest_descriptor(),
        "binary": manifest["versions"][str(spec["version"])] ["binaries"][str(spec["lane"])],
        "validator": manifest["validator"],
        "command": wrapped, "environment": env,
        "started_utc": started, "finished_utc": finished,
        "start_environment": start_environment, "end_environment": end_environment,
        "exit_code": completed.returncode if completed is not None else 127,
        "stdout": retained_receipt_artifact(paths["stdout"], str(spec["label"])),
        "stderr": retained_receipt_artifact(paths["stderr"], str(spec["label"])),
        "process": retained_receipt_artifact(paths["process"], str(spec["label"]))
        if paths["process"].exists() else None,
    }
    if paths["process"].exists():
        receipt_value["process_metrics"] = parse_time_v(paths["process"])
    if status == "passed":
        try:
            report, parse_error = report_json_from_stdout(paths["stdout"])
            require(parse_error is None and report is not None,
                    f"{spec['label']}: {parse_error}")
            warmups = int(p[str(spec["lane"])] ["warmups_per_child"])
            samples = int(p[str(spec["lane"])] ["samples_per_child"])
            summary = validate_probe_report(report, spec, warmups=warmups, samples=samples)
            receipt_value["report"] = summary
            receipt_value["report_canonical_sha256"] = summary["report_sha256"]
        except Exception as error:
            status = "failed"
            receipt_value["status"] = status
            receipt_value["validation_error"] = str(error)
    write_json(receipt, receipt_value)
    if status != "passed":
        fail(f"{spec['label']}: child failed; attempt retained at {receipt}")


def check_schedule_prefix(schedule_rows: list[dict[str, object]]) -> None:
    for index, spec in enumerate(schedule_rows):
        path = receipt_path(spec)
        if not path.exists():
            require(all(receipt_path(previous).exists() for previous in schedule_rows[:index]),
                    f"schedule has a gap before {spec['label']}")
            return
    fail("all scheduled children already have receipts; refusing a rerun")


def run(args: argparse.Namespace) -> None:
    manifest = load_manifest()
    p = protocol()
    rows = schedule(p, args.lane)
    if args.label is not None:
        safe_label(args.label)
        matching = [row for row in rows if row["label"] == args.label]
        require(len(matching) == 1, f"unknown {args.lane} child label: {args.label}")
        target_index = rows.index(matching[0])
        require(all(receipt_path(row).exists() for row in rows[:target_index]),
                f"cannot selectively run {args.label}; schedule prefix is incomplete")
        require(not receipt_path(matching[0]).exists(), f"child already retained: {args.label}")
        run_one(manifest, matching[0], rows)
        return
    check_schedule_prefix(rows)
    for spec in rows:
        run_one(manifest, spec, rows)


def qualify(args: argparse.Namespace) -> None:
    require(args.warmup == 0 and args.samples == 1,
            "qualification is fixed to --warmup 0 --samples 1")
    require(args.version in VERSIONS and args.lane in LANES, "invalid qualification identity")
    p = protocol()
    fixture = fixture_rows(p)[args.case]
    child = {"lane": args.lane, "version": args.version, "case": args.case,
             "edits": args.edits, "retention": args.retention,
             "pair_id": f"qualification-{args.version}-{args.case}-e{args.edits}-{args.retention}",
             "label": f"qualification-{args.version}-{args.case}-e{args.edits}-{args.retention}"}
    binary = Path(args.binary).resolve()
    binary_descriptor(binary, "qualification binary")
    command = ["/usr/bin/taskset", "-c", str(p[args.lane]["cpu"]), str(binary),
               "--case", args.case, "--input", str(fixture["path_abs"]),
               "--edits", str(args.edits), "--retention", args.retention,
               "--warmup", "0", "--samples", "1"]
    root = (Path(args.qualification_root).resolve() if args.qualification_root
            else (Path(os.environ.get("LITCHI_PERF_0922_QUALIFY_ROOT",
                                    "/tmp/litchi-performance-0922-tooling/qualification"))
                  .resolve()))
    label = safe_label(args.label or str(child["label"]))
    child_dir = root / label
    require(not child_dir.exists(), f"refusing to replace qualification child: {child_dir}")
    child_dir.mkdir(parents=True, exist_ok=False)
    stdout_path = child_dir / "stdout.jsonl"
    stderr_path = child_dir / "stderr.log"
    process_path = child_dir / "process.time"
    wrapped = process_command(command, process_path)
    env = {"PATH": os.environ.get("PATH", ""), "LC_ALL": "C", "RUST_BACKTRACE": "0"}
    started = now()
    with stdout_path.open("xb") as stdout, stderr_path.open("xb") as stderr:
        try:
            completed = subprocess.run(wrapped, cwd=REPO_ROOT, env=env,
                                       stdin=subprocess.DEVNULL, stdout=stdout,
                                       stderr=stderr, check=False)
        except OSError as error:
            completed = None
            stderr.write(f"qualification could not start child: {error}\n".encode())
    finished = now()
    status = "passed" if completed is not None and completed.returncode == 0 else "failed"
    receipt_path = child_dir / "qualification.json"
    receipt: dict[str, object] = {
        "schema": "litchi.performance.0922.qualification.v1",
        "status": status, "version": args.version, "lane": args.lane,
        "spec": child, "protocol": protocol_descriptor(),
        "validator": script_descriptor(VALIDATOR_PATH),
        "binary": binary_descriptor(binary, "qualification binary"),
        "fixture": {"path": str(fixture["path_abs"]), "bytes": fixture["bytes"],
                     "sha256": fixture["sha256"]},
        "command": wrapped, "environment": env,
        "started_utc": started, "finished_utc": finished,
        "exit_code": completed.returncode if completed is not None else 127,
        "stdout": descriptor(stdout_path), "stderr": descriptor(stderr_path),
        "process": descriptor(process_path) if process_path.exists() else None,
        "process_metrics": parse_time_v(process_path) if process_path.exists() else None,
    }
    if status == "passed":
        try:
            report, parse_error = report_json_from_stdout(stdout_path)
            require(parse_error is None and report is not None,
                    f"qualification report parse failed: {parse_error}")
            summary = validate_probe_report(report, child, warmups=0, samples=1)
            receipt["report"] = summary
        except Exception as error:
            status = "failed"
            receipt["status"] = status
            receipt["validation_error"] = str(error)
    write_json(receipt_path, receipt)
    if status != "passed":
        fail(f"qualification failed; retained raw custody at {child_dir}")
    print(json.dumps({"status": "qualified", "version": args.version, "lane": args.lane,
                      "case": args.case, "edits": args.edits, "retention": args.retention,
                      "receipt": descriptor(receipt_path), "summary": receipt["report"]},
                     sort_keys=True))


def read_receipt(manifest: dict[str, Any], spec: dict[str, object],
                 schedule_rows: list[dict[str, object]]) -> dict[str, Any]:
    path = receipt_path(spec)
    value = read_json(path)
    require(isinstance(value, dict), f"{spec['label']}: receipt is not an object")
    require(value.get("schema") == "litchi.performance.0922.child-receipt.v1",
            f"{spec['label']}: receipt schema changed")
    require(value.get("status") == "passed" and value.get("exit_code") == 0,
            f"{spec['label']}: child did not pass")
    require(value.get("spec") == spec and value.get("protocol") == manifest["protocol"]
            and value.get("manifest") == manifest_descriptor(),
            f"{spec['label']}: schedule or manifest binding changed")
    require(value.get("binary") == manifest["versions"][str(spec["version"])]
            ["binaries"][str(spec["lane"])], f"{spec['label']}: binary binding changed")
    started = parse_utc(value.get("started_utc"), f"{spec['label']} start")
    finished = parse_utc(value.get("finished_utc"), f"{spec['label']} finish")
    require(started <= finished, f"{spec['label']}: finish precedes start")
    for prior in schedule_prior(schedule_rows, spec):
        prior_value = read_json(receipt_path(prior))
        require(parse_utc(prior_value.get("finished_utc"), f"{prior['label']} finish") <= started,
                f"{spec['label']}: schedule overlap")
    for key in ("stdout", "stderr", "process"):
        require(value.get(key) is not None, f"{spec['label']}: {key} artifact missing")
        artifact_for_read(value[key], f"{spec['label']}: {key}")
    artifact_for_read(value["process"], f"{spec['label']}: process")
    metrics = value.get("process_metrics")
    require(isinstance(metrics, dict), f"{spec['label']}: process metrics missing")
    process_bytes = read_artifact_bytes(value["process"], f"{spec['label']}: process")
    actual_metrics = parse_time_v_text(
        process_bytes.decode("utf-8", "strict"), f"{spec['label']}: retained process log")
    require({key: item for key, item in actual_metrics.items() if key != "path"}
            == {key: item for key, item in metrics.items() if key != "path"},
            f"{spec['label']}: process metrics changed")
    stdout_bytes = read_artifact_bytes(value["stdout"], f"{spec['label']}: stdout")
    temporary_fd, temporary_name = tempfile.mkstemp(prefix="litchi-0922-report-")
    os.close(temporary_fd)
    temporary = Path(temporary_name)
    try:
        temporary.write_bytes(stdout_bytes)
        report, parse_error = report_json_from_stdout(temporary)
    finally:
        temporary.unlink(missing_ok=True)
    require(parse_error is None and report is not None,
            f"{spec['label']}: retained stdout invalid: {parse_error}")
    p = protocol()
    summary = validate_probe_report(
        report, spec, warmups=int(p[str(spec["lane"])] ["warmups_per_child"]),
        samples=int(p[str(spec["lane"])] ["samples_per_child"]),
    )
    require(value.get("report") == summary
            and value.get("report_canonical_sha256") == summary["report_sha256"],
            f"{spec['label']}: receipt report summary changed")
    return {"spec": spec, "receipt": value, "report": report, "summary": summary}


def nearest_rank(values: list[float], fraction: float) -> float:
    require(values, "empty metric vector")
    ordered = sorted(values)
    index = min(max(math.ceil(fraction * len(ordered)) - 1, 0), len(ordered) - 1)
    return ordered[index]


def stats(values: list[float]) -> dict[str, float]:
    require(values, "empty metric vector")
    return {"p50": nearest_rank(values, 0.50), "p95": nearest_rank(values, 0.95),
            "p99": nearest_rank(values, 0.99), "mean": statistics.fmean(values)}


def fraction_flag(before: float, after: float, limit: float) -> dict[str, object]:
    increase = float("inf") if before == 0 and after > 0 else (
        after / before - 1.0 if before else 0.0)
    return {"before": before, "after": after, "increase_fraction": increase,
            "flag": increase > limit}


def bootstrap(ratios: list[float], draws: int, seed: int) -> dict[str, object]:
    require(ratios, "empty paired ratio vector")
    rng = random.Random(seed)
    medians = [statistics.median(ratios[rng.randrange(len(ratios))] for _ in ratios)
               for _ in range(draws)]
    return {
        "draws": draws, "seed": seed, "ratios": ratios,
        "median_ratio": statistics.median(ratios),
        "confidence_interval_95": {
            "method": "nearest-rank 2.5th/97.5th percentiles",
            "lower": nearest_rank(medians, 0.025),
            "upper": nearest_rank(medians, 0.975),
        },
    }


def capacity_signature(report: dict[str, Any]) -> dict[str, Any]:
    """Return only the explicitly admitted capacity fields.

    This is intentionally path-based.  A field is capacity-exempt only when
    it is one of the report fields produced by the unchanged 0917 validator;
    a newly added field whose name happens to contain ``capacity`` remains in
    the semantic signature and therefore cannot be silently ignored.
    """
    routes: dict[str, object] = {}
    for route in report["route_outputs"]:
        name = str(route["retention"])
        require(name not in routes, f"duplicate route in capacity map: {name}")
        stages = []
        for stage in route["stages"]:
            stages.append({key: stage[key] for key in ("before", "after", "after_release")})
        routes[name] = {
            "output_capacity": route["output_capacity"],
            "final_retained_before_commit": route["final_retained_before_commit"],
            "stages": stages,
        }
    samples = []
    for sample in report["samples"]:
        stages = sample.get("stages", sample.get("stage_observations"))
        sample_capacity: dict[str, object] = {
            "output_capacity": sample["output_capacity"],
            "stages": [{key: stage[key] for key in ("before", "after", "after_release")}
                       for stage in stages],
        }
        allocation = sample.get("allocation")
        sample_capacity["retained_bytes"] = (
            allocation.get("retained_bytes") if isinstance(allocation, dict) else None
        )
        samples.append(sample_capacity)
    allocation = report.get("allocation")
    return {
        "output_capacity": report["output_capacity"],
        "routes": routes,
        "samples": samples,
        "retained_bytes": allocation.get("retained_bytes") if isinstance(allocation, dict) else None,
    }


def _remove_capacity_fields(value: dict[str, Any]) -> None:
    """Remove only the fields enumerated by :func:`capacity_signature`."""
    value.pop("output_capacity", None)
    for route in value.get("route_outputs", []):
        if isinstance(route, dict):
            route.pop("output_capacity", None)
            route.pop("final_retained_before_commit", None)
            for stage in route.get("stages", []):
                if isinstance(stage, dict):
                    for key in ("before", "after", "after_release"):
                        stage.pop(key, None)


def _remove_allocation_metrics(value: dict[str, Any]) -> None:
    """Remove measured allocator metrics, retaining unknown fields for equality."""
    metric_names = {"allocated_bytes", "deallocated_bytes", "allocation_calls",
                    "peak_live_bytes", "retained_bytes"}

    def normalize(node: object) -> None:
        if not isinstance(node, dict):
            return
        for key in metric_names:
            node.pop(key, None)

    normalize(value.get("allocation"))
    for sample in value.get("samples", []):
        if isinstance(sample, dict):
            normalize(sample.get("allocation"))
            allocations = sample.get("allocations")
            if isinstance(allocations, dict):
                for allocation in allocations.values():
                    normalize(allocation)
    for sample in value.get("samples", []):
        if isinstance(sample, dict):
            sample.pop("output_capacity", None)
            stages = sample.get("stages", sample.get("stage_observations", []))
            for stage in stages:
                if isinstance(stage, dict):
                    for key in ("before", "after", "after_release"):
                        stage.pop(key, None)


def semantic_signature(report: dict[str, Any]) -> dict[str, Any]:
    """Return controls and identities while excluding metrics and known capacities."""
    value = json.loads(json.dumps(report))
    value.pop("elapsed_ns", None)
    for sample in value.get("samples", []):
        if isinstance(sample, dict):
            sample.pop("elapsed_ns", None)
    _remove_allocation_metrics(value)
    _remove_capacity_fields(value)
    return value


def compare_capacity_fields(baseline: dict[str, Any], candidate: dict[str, Any],
                            label: str) -> dict[str, object]:
    """Require every explicit capacity and retained transition to match exactly."""
    before = capacity_signature(baseline)
    after = capacity_signature(candidate)
    require(before.keys() == after.keys(), f"{label}: capacity field set changed")
    changes: list[dict[str, object]] = []

    def check(path: str, base: object, cand: object) -> None:
        require((base is None) == (cand is None),
                f"{label}: retained transition changed at {path}")
        if base is None:
            changes.append({"path": path, "baseline": None, "candidate": None,
                            "equal": True, "nonincreasing": True})
            return
        require(type(base) is int and type(cand) is int and base >= 0 and cand >= 0,
                f"{label}: invalid capacity value at {path}")
        require(cand == base,
                f"{label}: candidate capacity changed at {path}: {base} != {cand}")
        changes.append({"path": path, "baseline": base, "candidate": cand,
                        "equal": True, "nonincreasing": True})

    check("output_capacity", before["output_capacity"], after["output_capacity"])
    require(before["routes"].keys() == after["routes"].keys(),
            f"{label}: route capacity identities changed")
    for route_name in sorted(before["routes"]):
        base_route = before["routes"][route_name]
        cand_route = after["routes"][route_name]
        check(f"route_outputs[{route_name}].output_capacity",
              base_route["output_capacity"], cand_route["output_capacity"])
        check(f"route_outputs[{route_name}].final_retained_before_commit",
              base_route["final_retained_before_commit"],
              cand_route["final_retained_before_commit"])
        require(len(base_route["stages"]) == len(cand_route["stages"]),
                f"{label}: stage capacity count changed for {route_name}")
        for index, (base_stage, cand_stage) in enumerate(
                zip(base_route["stages"], cand_route["stages"], strict=True), start=1):
            for key in ("before", "after", "after_release"):
                check(f"route_outputs[{route_name}].stages[{index}].{key}",
                      base_stage[key], cand_stage[key])
    require(len(before["samples"]) == len(after["samples"]),
            f"{label}: sample capacity count changed")
    for index, (base_sample, cand_sample) in enumerate(
            zip(before["samples"], after["samples"], strict=True)):
        check(f"samples[{index}].output_capacity",
              base_sample["output_capacity"], cand_sample["output_capacity"])
        check(f"samples[{index}].retained_bytes",
              base_sample["retained_bytes"], cand_sample["retained_bytes"])
        require(len(base_sample["stages"]) == len(cand_sample["stages"]),
                f"{label}: sample stage capacity count changed at {index}")
        for stage_index, (base_stage, cand_stage) in enumerate(
                zip(base_sample["stages"], cand_sample["stages"], strict=True), start=1):
            for key in ("before", "after", "after_release"):
                check(f"samples[{index}].stages[{stage_index}].{key}",
                      base_stage[key], cand_stage[key])
    check("allocation.retained_bytes", before["retained_bytes"], after["retained_bytes"])
    def canonical(value: object) -> bytes:
        return json.dumps(value, sort_keys=True, separators=(",", ":")).encode("utf-8")

    numeric = [row for row in changes if row["baseline"] is not None]
    decreases = [int(row["baseline"]) - int(row["candidate"]) for row in numeric]
    return {
        "field_count": len(changes),
        "numeric_field_count": len(numeric),
        "none_field_count": len(changes) - len(numeric),
        "changed_field_count": sum(row["baseline"] != row["candidate"] for row in changes),
        "all_equal": all(bool(row["equal"]) for row in changes),
        "all_nonincreasing": all(bool(row["nonincreasing"]) for row in changes),
        "baseline_capacity_sha256": sha256_bytes(canonical(before)),
        "candidate_capacity_sha256": sha256_bytes(canonical(after)),
        "baseline_max": max((int(row["baseline"]) for row in numeric), default=None),
        "candidate_max": max((int(row["candidate"]) for row in numeric), default=None),
        "max_decrease": max(decreases, default=0),
    }


def compare_reports(baseline: dict[str, Any], candidate: dict[str, Any], label: str) -> dict[str, object]:
    require(semantic_signature(baseline) == semantic_signature(candidate),
            f"{label}: non-capacity semantic identity differs")
    return compare_capacity_fields(baseline, candidate, label)


def report_elapsed(report: dict[str, Any]) -> list[float]:
    values = report.get("elapsed_ns")
    require(isinstance(values, list) and values, "native elapsed vector missing")
    return [float(integer(value, "elapsed sample", 1)) for value in values]


def allocations(report: dict[str, Any]) -> dict[str, int]:
    value = report.get("allocation")
    require(isinstance(value, dict), "allocator top allocation missing")
    return {key: integer(value.get(key), f"allocation.{key}")
            for key in ("allocated_bytes", "deallocated_bytes", "allocation_calls",
                        "peak_live_bytes", "retained_bytes")}


def pair_rows(items: list[dict[str, Any]]) -> dict[tuple[Any, ...], dict[str, dict[str, Any]]]:
    result: dict[tuple[Any, ...], dict[str, dict[str, Any]]] = {}
    for item in items:
        spec = item["spec"]
        key = ((spec["round"], spec["case"], spec["edits"]) if spec["lane"] == "native"
               else (spec["repetition"], spec["case"], spec["edits"], spec["retention"]))
        result.setdefault(key, {})[str(spec["version"])] = item
    return result


def reduce_native(manifest: dict[str, Any], p: dict[str, Any]) -> dict[str, object]:
    rows = schedule(p, "native")
    loaded = [read_receipt(manifest, spec, rows) for spec in rows]
    pairs = pair_rows(loaded)
    cells: list[dict[str, object]] = []
    for case in [str(row["case"]) for row in p["fixtures"]]:
        for edits in EDITS:
            round_pairs = [pairs[(round_no, case, edits)] for round_no in range(1, 13)]
            pair_details = []
            ratios = []
            for round_no, pair in enumerate(round_pairs, start=1):
                require(set(pair) == set(VERSIONS), f"native pair incomplete: {case}/e{edits}/r{round_no}")
                base = pair["baseline"]
                candidate = pair["candidate"]
                capacity = compare_reports(
                    base["report"], candidate["report"],
                    f"native {case}/e{edits}/r{round_no}")
                base_values = report_elapsed(base["report"])
                candidate_values = report_elapsed(candidate["report"])
                base_stats = stats(base_values)
                candidate_stats = stats(candidate_values)
                ratio = base_stats["p50"] / candidate_stats["p50"]
                ratios.append(ratio)
                base_rss = float(base["receipt"]["process_metrics"]["max_rss_kib"])
                candidate_rss = float(candidate["receipt"]["process_metrics"]["max_rss_kib"])
                pair_details.append({
                    "round": round_no,
                    "baseline": {"label": base["spec"]["label"], "stats": base_stats,
                                 "rss_kib": base_rss},
                    "candidate": {"label": candidate["spec"]["label"], "stats": candidate_stats,
                                   "rss_kib": candidate_rss},
                    "p50_ratio_baseline_over_candidate": ratio,
                    "capacity": capacity,
                    "rss": fraction_flag(base_rss, candidate_rss, 0.05),
                    "latency": {field: fraction_flag(base_stats[field], candidate_stats[field], 0.05)
                                for field in ("p50", "p95", "p99")},
                })
            base_all = [value for pair in round_pairs for value in report_elapsed(pair["baseline"]["report"])]
            candidate_all = [value for pair in round_pairs for value in report_elapsed(pair["candidate"]["report"])]
            boot = bootstrap(ratios, 10000, 922)
            lower = float(boot["confidence_interval_95"]["lower"])
            cells.append({
                "case": case, "edits": edits, "retention": "default",
                "baseline": {"pooled": stats(base_all)},
                "candidate": {"pooled": stats(candidate_all)},
                "rounds": pair_details,
                "paired_bootstrap": boot,
                "latency_gate": lower >= 1 / 1.05,
            })
    return {"schema": "litchi.performance.0922.native-reduction.v1", "children": len(loaded),
            "paired_cells": 96, "scenario_cells": len(cells), "paired_rounds": 12,
            "cells": cells,
            "claim": "paired baseline/candidate native comparison at default retention"}


def reduce_allocator(manifest: dict[str, Any], p: dict[str, Any]) -> dict[str, object]:
    rows = schedule(p, "allocator")
    loaded = [read_receipt(manifest, spec, rows) for spec in rows]
    pairs = pair_rows(loaded)
    cells: list[dict[str, object]] = []
    for key in sorted(pairs):
        pair = pairs[key]
        require(set(pair) == set(VERSIONS), f"allocator pair incomplete: {key}")
        base = pair["baseline"]
        candidate = pair["candidate"]
        capacity = compare_reports(base["report"], candidate["report"],
                                   f"allocator {key}")
        base_alloc = allocations(base["report"])
        candidate_alloc = allocations(candidate["report"])
        cells.append({
            "repetition": key[0], "case": key[1], "edits": key[2], "retention": key[3],
            "baseline": base_alloc, "candidate": candidate_alloc,
            "capacity": capacity,
            "allocation": fraction_flag(float(base_alloc["allocated_bytes"]),
                                         float(candidate_alloc["allocated_bytes"]), 0.01),
            "peak_live": fraction_flag(float(base_alloc["peak_live_bytes"]),
                                        float(candidate_alloc["peak_live_bytes"]), 0.05),
            "rss": fraction_flag(float(base["receipt"]["process_metrics"]["max_rss_kib"]),
                                  float(candidate["receipt"]["process_metrics"]["max_rss_kib"]), 0.05),
        })
    return {"schema": "litchi.performance.0922.allocator-reduction.v1", "children": len(loaded),
            "cells": cells, "claim": "paired System allocator region comparison; no latency claim"}


def reduce_data(manifest: dict[str, Any]) -> dict[str, object]:
    p = protocol()
    native = reduce_native(manifest, p)
    allocator = reduce_allocator(manifest, p)
    native_gates = [bool(cell["latency_gate"]) for cell in native["cells"]]
    allocation_cells = allocator["cells"]
    capacity_rows = [row["capacity"] for cell in native["cells"] for row in cell["rounds"]]
    capacity_rows.extend(cell["capacity"] for cell in allocation_cells)
    capacity_equal = all(
        bool(row["all_equal"])
        for row in capacity_rows
    )
    no_alloc_increase = all(not bool(cell["allocation"]["flag"]) for cell in allocation_cells)
    no_peak_increase = all(not bool(cell["peak_live"]["flag"]) for cell in allocation_cells)
    default32 = [cell for cell in allocation_cells if cell["edits"] == 32 and cell["retention"] == "default"]
    by_target: dict[str, list[dict[str, object]]] = {}
    for cell in default32:
        by_target.setdefault(str(cell["case"]), []).append(cell)
    target_rows = []
    for case, values in sorted(by_target.items()):
        base = statistics.median(float(row["baseline"]["allocated_bytes"]) for row in values)
        candidate = statistics.median(float(row["candidate"]["allocated_bytes"]) for row in values)
        target_rows.append({"case": case, "baseline_median_allocated_bytes": base,
                            "candidate_median_allocated_bytes": candidate,
                            "candidate_over_baseline": candidate / base if base else float("inf"),
                            "saves_at_least_5_percent": candidate <= base * 0.95})
    savings = any(bool(row["saves_at_least_5_percent"]) for row in target_rows)
    gates = {
        "all_native_default_latency_lower_bounds": all(native_gates),
        "candidate_saves_at_least_5_percent_allocated_bytes_in_one_32_edit_default_fixture": savings,
        "no_allocator_allocated_bytes_increase_over_1_percent": no_alloc_increase,
        "no_allocator_peak_live_increase_over_5_percent": no_peak_increase,
        "all_explicit_capacity_fields_equal": capacity_equal,
        "all_children_and_controls_validated": True,
        "adoption": all(native_gates) and savings and no_alloc_increase and no_peak_increase
        and capacity_equal,
    }
    return {
        "schema": "litchi.performance.0922.doc-comparison.reduction.v1",
        "protocol": protocol_descriptor(), "manifest": manifest_descriptor(),
        "native": native, "allocator": allocator,
        "admission_targets": target_rows, "gates": gates,
        "review_flags": {
            "latency": [row for cell in native["cells"] for row in cell["rounds"]],
            "allocation_peak_rss": allocation_cells,
        },
    }


def reduce(args: argparse.Namespace) -> None:
    manifest = load_manifest()
    result = reduce_data(manifest)
    output = args.output.resolve() if args.output else REDUCED_PATH
    write_json(output, {**result, "generated_utc": now()})
    print(json.dumps({"status": "reduced", "output": packet_descriptor(output, allow_large=False),
                      "adoption": result["gates"]["adoption"]}, sort_keys=True))


def replay(args: argparse.Namespace) -> dict[str, object]:
    manifest = load_manifest()
    retained = read_json(args.reduced.resolve() if args.reduced else REDUCED_PATH)
    require(isinstance(retained, dict)
            and retained.get("schema") == "litchi.performance.0922.doc-comparison.reduction.v1",
            "reduced comparison schema mismatch")
    require(retained.get("protocol") == manifest["protocol"]
            and retained.get("manifest") == manifest_descriptor(),
            "reduced provenance mismatch")
    deterministic = reduce_data(manifest)
    for key in ("protocol", "manifest", "native", "allocator", "admission_targets", "gates", "review_flags"):
        require(retained.get(key) == deterministic.get(key),
                f"reduction replay mismatch: {key}")
    expected = {str(row["label"]) for row in all_schedule(protocol())}
    actual = {path.stem for path in LOG_DIR.glob("*.json") if path.is_file()}
    require(actual == expected,
            f"receipt set differs: missing={sorted(expected - actual)}, "
            f"extra={sorted(actual - expected)}")
    for path in PACKET_ROOT.rglob("*"):
        if path.is_file() and path.stat().st_size >= MAX_PACKET_ARTIFACT \
                and path.name not in {"comparison.py", "comparison-replay.py"}:
            fail(f"packet contains an artifact at or above 1 MiB: {path}")
    return {"status": "verified", "native_children": 192, "allocator_children": 96,
            "adoption": deterministic["gates"]["adoption"]}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="action", required=True)
    freeze_parser = sub.add_parser("freeze")
    freeze_parser.add_argument("--baseline-native", required=True)
    freeze_parser.add_argument("--baseline-allocator", required=True)
    freeze_parser.add_argument("--candidate-native", required=True)
    freeze_parser.add_argument("--candidate-allocator", required=True)
    freeze_parser.add_argument("--baseline-source-receipt", required=True)
    freeze_parser.add_argument("--baseline-build-receipt", required=True)
    freeze_parser.add_argument("--candidate-source-receipt", required=True)
    freeze_parser.add_argument("--candidate-build-receipt", required=True)
    freeze_parser.add_argument("--candidate-patch")
    freeze_parser.add_argument("--probe-source")
    freeze_parser.add_argument("--validator")
    freeze_parser.add_argument("--external-root")
    run_parser = sub.add_parser("run")
    run_parser.add_argument("--lane", choices=LANES, required=True)
    run_parser.add_argument("--label")
    qualify_parser = sub.add_parser("qualify")
    qualify_parser.add_argument("--version", choices=VERSIONS, required=True)
    qualify_parser.add_argument("--lane", choices=LANES, required=True)
    qualify_parser.add_argument("--binary", required=True)
    qualify_parser.add_argument("--case", choices=("docnohf", "docfloat"), required=True)
    qualify_parser.add_argument("--edits", type=int, choices=EDITS, required=True)
    qualify_parser.add_argument("--retention", choices=ROUTES, required=True)
    qualify_parser.add_argument("--warmup", type=int, required=True)
    qualify_parser.add_argument("--samples", type=int, required=True)
    qualify_parser.add_argument("--qualification-root")
    qualify_parser.add_argument("--label")
    reduce_parser = sub.add_parser("reduce")
    reduce_parser.add_argument("--output", type=Path)
    replay_parser = sub.add_parser("replay")
    replay_parser.add_argument("--reduced", type=Path)
    args = parser.parse_args()
    try:
        if args.action == "freeze":
            freeze(args)
        elif args.action == "run":
            run(args)
        elif args.action == "qualify":
            qualify(args)
        elif args.action == "reduce":
            reduce(args)
        else:
            print(json.dumps(replay(args), sort_keys=True))
    except (EvidenceError, OSError, KeyError, TypeError, ValueError, subprocess.SubprocessError) as error:
        print(f"0922 comparison failed: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
