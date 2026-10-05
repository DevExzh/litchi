#!/usr/bin/env python3
"""Freeze, capture, reduce, and replay the 0917 real-DOC retention packet.

This runner is deliberately a same-build route comparison.  It does not build
production code and it has no candidate/adoption gate.  ``freeze`` binds the
protocol, this script, the probe source, build/source receipts, binaries and
fixtures.  ``run`` executes the already-frozen probe in the immutable schedule;
``reduce`` validates every retained child and writes compact per-cell summaries.
Large raw outputs stay in the external capture root.  A raw report or process
log smaller than one MiB is copied into the packet for offline replay.
"""

from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import importlib.util
import json
import math
import os
from pathlib import Path
import random
import re
import shutil
import statistics
import subprocess
import sys
import time
from typing import Any, Iterable


HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[3]
PROTOCOL_PATH = HERE / "native-protocol.json"
SCRIPT_PATH = HERE / "capture.py"
REPLAY_PATH = HERE / "replay.py"
MANIFEST_PATH = HERE / "manifest.json"
REDUCED_PATH = HERE / "reduced.json"
PACKET_RAW = HERE / "raw"
PACKET_LOGS = HERE / "logs"
EXTERNAL = Path(
    os.environ.get("LITCHI_PERF_0917_ROOT", "/home/zhuhe/litchi-performance-0917")
).resolve()
EXTERNAL_CHILDREN = EXTERNAL / "children"
MAX_PACKET_ARTIFACT = 1 << 20
HEX = set("0123456789abcdef")


class EvidenceError(RuntimeError):
    pass


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


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def sha256(path: Path) -> str:
    require(path.is_file() and not path.is_symlink(), f"missing regular file: {path}")
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def descriptor(path: Path, *, allow_large: bool = True) -> dict[str, object]:
    path = Path(path)
    require(path.is_file() and not path.is_symlink(), f"missing regular file: {path}")
    resolved = path.resolve()
    size = resolved.stat().st_size
    if not allow_large:
        require(size < MAX_PACKET_ARTIFACT, f"packet artifact is at least 1 MiB: {resolved}")
    return {"path": str(resolved), "bytes": size, "sha256": sha256(resolved)}


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
    require(len(data) < MAX_PACKET_ARTIFACT,
            f"JSON artifact is at least 1 MiB: {path}")
    write_exclusive(path, data)


def copy_exclusive(source: Path, target: Path) -> dict[str, object]:
    source_identity = descriptor(source)
    data = source.read_bytes()
    require(len(data) < MAX_PACKET_ARTIFACT, f"refusing large committed artifact: {source}")
    write_exclusive(target, data)
    committed = packet_descriptor(target, allow_large=False)
    require(
        committed["bytes"] == source_identity["bytes"]
        and committed["sha256"] == source_identity["sha256"],
        f"source changed while copying {source}",
    )
    return committed


def digest(value: object, label: str) -> str:
    require(isinstance(value, str) and len(value) == 64 and set(value) <= HEX,
            f"{label}: invalid SHA-256")
    return value


def integer(value: object, label: str, minimum: int = 0) -> int:
    require(type(value) is int and value >= minimum, f"{label}: invalid integer")
    return value


def protocol() -> dict[str, Any]:
    value = read_json(PROTOCOL_PATH)
    require(isinstance(value, dict), "protocol is not an object")
    require(value.get("schema") == "litchi.performance.0917.protocol.v1",
            "unexpected 0917 protocol schema")
    fixtures = value.get("fixtures")
    require(isinstance(fixtures, list) and [x.get("case") for x in fixtures]
            == ["docnohf", "docfloat"], "fixture order changed")
    expected_fixtures = {
        "docnohf": ("test-data/ole/doc/NoHeadFoot.doc", 26112,
                     "45e5df073f34314da6f39d2dad119fb2ef23470878fd2df67f632864cd92ea48"),
        "docfloat": ("test-data/ole/doc/FloatingPictures.doc", 335360,
                      "41cc0cd8f2d7390266f844f5bddbcd29d08ac5338abc33832d9c9a6ed5d0f85d"),
    }
    for row in fixtures:
        require(isinstance(row, dict), "fixture row is not an object")
        case = row.get("case")
        require(case in expected_fixtures, f"unknown fixture: {case!r}")
        relative, size, expected_hash = expected_fixtures[case]
        require(row.get("path") == relative and row.get("bytes") == size
                and row.get("sha256") == expected_hash, f"{case}: protocol fixture changed")
    require(value.get("edits") == [1, 2, 8, 32], "edit counts changed")
    require(value.get("retention_routes") == ["zero", "default", "release"],
            "retention route order changed")
    require(value.get("retention_ceiling_bytes") == 8 * 1024 * 1024,
            "retention ceiling changed")
    native = value.get("native")
    require(isinstance(native, dict), "native section missing")
    require(native.get("cpu") == 12 and native.get("rounds") == 6
            and native.get("warmups_per_child") == 3
            and native.get("samples_per_child") == 30
            and native.get("children") == 144, "native schedule changed")
    require(native.get("route_orders") == [
        ["zero", "default", "release"],
        ["zero", "release", "default"],
        ["default", "zero", "release"],
        ["default", "release", "zero"],
        ["release", "zero", "default"],
        ["release", "default", "zero"],
    ], "native route permutations changed")
    allocator = value.get("allocator")
    require(isinstance(allocator, dict), "allocator section missing")
    require(allocator.get("repetitions") == 2 and allocator.get("children") == 48
            and allocator.get("samples_per_child") == 1
            and allocator.get("warmups_per_child") == 0, "allocator schedule changed")
    bootstrap = value.get("bootstrap")
    require(isinstance(bootstrap, dict) and bootstrap.get("draws") == 10000
            and bootstrap.get("seed") == 917
            and bootstrap.get("unit") == "matched round within fixed fixture/edit count"
            and bootstrap.get("statistic") ==
            "median of per-round zero-p50/route-p50 ratios for default and release"
            and bootstrap.get("interval") == "nearest-rank2.5th/97.5th percentiles",
            "bootstrap changed")
    flags = value.get("review_flags")
    require(isinstance(flags, dict) and flags.get("latency_increase_fraction") == 0.05
            and flags.get("peak_live_increase_fraction") == 0.05
            and flags.get("rss_increase_fraction") == 0.05, "review flags changed")
    gates = value.get("mandatory_gates")
    require(isinstance(gates, list) and len(gates) >= 10, "mandatory gates missing")
    return value


def fixture_rows(p: dict[str, Any]) -> dict[str, dict[str, Any]]:
    result: dict[str, dict[str, Any]] = {}
    for row in p["fixtures"]:
        case = str(row["case"])
        path = ROOT / str(row["path"])
        require(path.is_file() and not path.is_symlink(), f"missing fixture: {path}")
        require(path.stat().st_size == row["bytes"] and sha256(path) == row["sha256"],
                f"fixture changed: {case}")
        result[case] = {**row, "path_abs": str(path.resolve())}
    return result


def fixture_custody(p: dict[str, Any]) -> dict[str, dict[str, Any]]:
    """Return relocatable fixture identities for the frozen manifest."""
    return {
        case: {key: value for key, value in row.items() if key != "path_abs"}
        for case, row in fixture_rows(p).items()
    }


def tree_descriptors(root: Path) -> dict[str, dict[str, object]]:
    require(root.is_dir() and not root.is_symlink(), f"missing probe source directory: {root}")
    rows: dict[str, dict[str, object]] = {}
    for path in sorted(root.rglob("*")):
        if path.is_file():
            relative = path.relative_to(root).as_posix()
            rows[relative] = {"bytes": path.stat().st_size, "sha256": sha256(path)}
    require(rows, f"probe source is empty: {root}")
    return rows


def verify_tree(root: Path, rows: object, label: str) -> None:
    require(isinstance(rows, dict) and rows, f"{label}: source map is empty")
    actual = tree_descriptors(root)
    require(set(rows) == set(actual), f"{label}: source file set changed")
    for relative, expected in rows.items():
        require(isinstance(expected, dict), f"{label}: malformed descriptor {relative}")
        require(actual[relative] == expected, f"{label}: changed source file {relative}")


def repository_relative(path: Path) -> str | None:
    try:
        return path.resolve().relative_to(ROOT).as_posix()
    except ValueError:
        return None


def packet_descriptor(path: Path, *, allow_large: bool = True) -> dict[str, object]:
    """Describe a packet file with a repository-relative path.

    External evidence keeps its absolute path, while packet-local identities
    must follow a copied checkout so replay cannot silently read the original
    packet after relocation.
    """
    identity = descriptor(path, allow_large=allow_large)
    relative = repository_relative(path)
    if relative is not None:
        identity["path"] = relative
    return identity


def manifest_path(identity: dict[str, object]) -> Path:
    raw = identity.get("path")
    require(isinstance(raw, str), "descriptor path missing")
    path = Path(raw)
    if path.is_absolute():
        return path
    return (ROOT / path).resolve()


def verify_descriptor(identity: object, label: str, *, allow_missing_binary: bool = False) -> None:
    require(isinstance(identity, dict), f"{label}: descriptor missing")
    path = manifest_path(identity)
    if allow_missing_binary and not path.exists():
        integer(identity.get("bytes"), f"{label}: bytes", 1)
        digest(identity.get("sha256"), f"{label}: hash")
        return
    actual = descriptor(path)
    if isinstance(identity.get("path"), str) and not Path(identity["path"]).is_absolute():
        actual["path"] = repository_relative(path) or actual["path"]
    require(actual == identity, f"{label}: content changed")


def protocol_descriptor() -> dict[str, object]:
    return packet_descriptor(PROTOCOL_PATH, allow_large=False)


def script_descriptor(path: Path) -> dict[str, object]:
    return packet_descriptor(path, allow_large=False)


def manifest_descriptor() -> dict[str, object]:
    return packet_descriptor(MANIFEST_PATH, allow_large=False)


def binary_descriptor(path: Path) -> dict[str, object]:
    identity = descriptor(path)
    # Binaries are allowed as external inputs but never as packet files.
    require(not path.resolve().is_relative_to(HERE), f"binary must stay outside packet: {path}")
    return identity


def read_receipt(path: Path | None, label: str) -> dict[str, object] | None:
    if path is None:
        return None
    identity = packet_descriptor(path, allow_large=False)
    value = read_json(path)
    require(isinstance(value, dict), f"{label}: receipt is not an object")
    return {"label": label, "descriptor": identity}


def live_repository_head() -> str:
    try:
        return subprocess.run(
            ["git", "rev-parse", "HEAD"], cwd=ROOT, check=True,
            stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True,
        ).stdout.strip()
    except (OSError, subprocess.SubprocessError) as error:
        fail(f"cannot verify repository source revision: {error}")


def freeze(args: argparse.Namespace) -> None:
    require(not MANIFEST_PATH.exists(), f"manifest already exists: {MANIFEST_PATH}")
    p = protocol()
    require(live_repository_head() == p["source_head"],
            "protocol source revision differs from current checkout at freeze")
    probe_root = (args.probe_source or HERE / "probe").resolve()
    probe_files = tree_descriptors(probe_root)
    binary_path = args.binary.resolve()
    allocator_path = args.allocator_binary.resolve()
    require(binary_path != allocator_path, "native and allocator binaries must differ")
    binaries = {"native": binary_descriptor(binary_path),
                "allocator": binary_descriptor(allocator_path)}
    receipts = []
    source_receipt = read_receipt(args.source_receipt, "source receipt")
    build_receipt = read_receipt(args.build_receipt, "build receipt")
    require(source_receipt is not None and build_receipt is not None,
            "freeze requires both --source-receipt and --build-receipt")
    receipts.extend((source_receipt, build_receipt))
    external = EXTERNAL
    require(not external.exists() or not any(external.iterdir()),
            f"external capture root is not empty: {external}")
    external.mkdir(parents=True, exist_ok=True)
    manifest = {
        "schema": "litchi.performance.0917.capture-manifest.v1",
        "created_utc": now(),
        "repository": str(ROOT.resolve()),
        "source_head": p["source_head"],
        "protocol": protocol_descriptor(),
        "capture_script": script_descriptor(SCRIPT_PATH),
        "replay_script": script_descriptor(REPLAY_PATH) if REPLAY_PATH.exists() else None,
        "probe_source": {"path": repository_relative(probe_root) or str(probe_root),
                          "files": probe_files},
        "binaries": binaries,
        "fixtures": fixture_custody(p),
        "receipts": receipts,
        "external_root": str(external),
        "external_children": str(EXTERNAL_CHILDREN),
        "protocol_sha256": sha256(PROTOCOL_PATH),
    }
    write_json(MANIFEST_PATH, manifest)
    print(json.dumps({"status": "frozen", "manifest": manifest_descriptor()},
                     sort_keys=True))


def manifest() -> dict[str, Any]:
    value = read_json(MANIFEST_PATH)
    require(isinstance(value, dict), "manifest is not an object")
    require(value.get("schema") == "litchi.performance.0917.capture-manifest.v1",
            "manifest schema mismatch")
    p = protocol()
    require(value.get("source_head") == p["source_head"],
            "manifest source revision changed")
    require(value.get("protocol_sha256") == sha256(PROTOCOL_PATH),
            "protocol changed since freeze")
    verify_descriptor(value.get("protocol"), "protocol")
    verify_descriptor(value.get("capture_script"), "capture script")
    if value.get("replay_script") is not None:
        verify_descriptor(value.get("replay_script"), "replay script")
    parse_utc(value.get("created_utc"), "manifest created")
    parse_utc(p.get("frozen_utc"), "protocol frozen")
    require(parse_utc(p["frozen_utc"], "protocol frozen") <=
            parse_utc(value["created_utc"], "manifest created"),
            "manifest predates frozen protocol")
    fixtures = fixture_custody(p)
    require(value.get("fixtures") == fixtures, "fixture custody changed")
    probe = value.get("probe_source")
    require(isinstance(probe, dict) and isinstance(probe.get("path"), str),
            "probe source descriptor missing")
    verify_tree(manifest_path({"path": probe["path"]}), probe.get("files"), "probe source")
    binaries = value.get("binaries")
    require(isinstance(binaries, dict) and set(binaries) == {"native", "allocator"},
            "binary descriptors changed")
    for key, identity in binaries.items():
        verify_descriptor(identity, f"{key} binary", allow_missing_binary=True)
    receipts = value.get("receipts")
    require(isinstance(receipts, list) and receipts, "receipt list missing")
    for row in receipts:
        require(isinstance(row, dict) and isinstance(row.get("label"), str),
                "receipt row malformed")
        verify_descriptor(row.get("descriptor"), row["label"] + " descriptor")
    require(isinstance(value.get("external_root"), str),
            "manifest external root missing")
    require(Path(value["external_root"]).resolve() == EXTERNAL,
            "external root differs from frozen root")
    return value


def cell_key(case: str, edits: int) -> tuple[str, int]:
    return case, edits


def native_schedule(p: dict[str, Any]) -> list[dict[str, object]]:
    fixtures = [str(x["case"]) for x in p["fixtures"]]
    edits = [int(x) for x in p["edits"]]
    routes = [str(x) for x in p["retention_routes"]]
    orders = p["native"]["route_orders"]
    result: list[dict[str, object]] = []
    for round_no in range(1, int(p["native"]["rounds"]) + 1):
        case_order = fixtures if round_no % 2 else list(reversed(fixtures))
        edit_order = edits if round_no % 2 else list(reversed(edits))
        route_order = orders[round_no - 1]
        require(route_order and set(route_order) == set(routes), f"round {round_no}: route order")
        for case in case_order:
            for count in edit_order:
                for leg, retention in enumerate(route_order, start=1):
                    label = f"native-r{round_no:02d}-{case}-e{count:02d}-{retention}"
                    result.append({"lane": "native", "round": round_no, "case": case,
                                   "edits": count, "leg": leg, "retention": retention,
                                   "label": label})
    require(len(result) == int(p["native"]["children"]), "native schedule count mismatch")
    return result


def allocation_schedule(p: dict[str, Any]) -> list[dict[str, object]]:
    cases = [str(x["case"]) for x in p["fixtures"]]
    edits = [int(x) for x in p["edits"]]
    routes = [str(x) for x in p["retention_routes"]]
    cells = [{"case": case, "edits": count, "retention": retention}
             for case in cases for count in edits for retention in routes]
    result: list[dict[str, object]] = []
    for repetition, ordered in enumerate((cells, list(reversed(cells))), start=1):
        for item in ordered:
            label = (f"allocator-r{repetition:02d}-{item['case']}-e{item['edits']:02d}-"
                     f"{item['retention']}")
            result.append({"lane": "allocator", "repetition": repetition, "leg": 1,
                           "label": label, **item})
    require(len(result) == int(p["allocator"]["children"]), "allocator schedule count mismatch")
    return result


def all_schedule(p: dict[str, Any], lane: str) -> list[dict[str, object]]:
    return native_schedule(p) if lane == "native" else allocation_schedule(p)


def safe_label(value: str) -> str:
    require(re.fullmatch(r"[a-z0-9-]+", value) is not None, f"unsafe child label: {value}")
    return value


def command_for(m: dict[str, Any], spec: dict[str, object], child: Path) -> list[str]:
    p = protocol()
    binary = m["binaries"]["native" if spec["lane"] == "native" else "allocator"]["path"]
    command = ["/usr/bin/taskset", "-c", str(p["native"]["cpu"]), binary,
               "--case", str(spec["case"]), "--input", fixture_rows(p)[str(spec["case"])]
               ["path"], "--edits", str(spec["edits"]), "--retention",
               str(spec["retention"]), "--warmup",
               str(p["native"]["warmups_per_child"] if spec["lane"] == "native"
                   else p["allocator"]["warmups_per_child"]), "--samples",
               str(p["native"]["samples_per_child"] if spec["lane"] == "native"
                   else p["allocator"]["samples_per_child"])]
    # The probe writes JSON to stdout.  Keeping an explicit child directory in
    # the command receipt makes provenance auditable even when raw output is
    # later moved outside the repository.
    return command


def run_process(command: list[str], cwd: Path, env: dict[str, str], stdout: Path,
                stderr: Path, process_log: Path) -> int:
    time_bin = "/usr/bin/time"
    wrapped = [time_bin, "-v", "-o", str(process_log), *command]
    with stdout.open("xb") as out, stderr.open("xb") as err:
        try:
            result = subprocess.run(wrapped, cwd=cwd, env=env, stdin=subprocess.DEVNULL,
                                    stdout=out, stderr=err, check=False)
        except OSError:
            # Keep a concrete process-log witness for a failed attempt.  The
            # empty log is intentionally retained and causes validation to
            # fail with a useful missing-metric error during reduction.
            if not process_log.exists():
                write_exclusive(process_log, b"")
            return 127
    return result.returncode


def parse_time_v(path: Path) -> dict[str, object]:
    require(path.is_file(), f"missing GNU time log: {path}")
    text = path.read_text(encoding="utf-8", errors="strict")
    values: dict[str, object] = {}
    patterns = {
        "user_seconds": r"^\s*User time \(seconds\):\s*([0-9.]+)$",
        "system_seconds": r"^\s*System time \(seconds\):\s*([0-9.]+)$",
        "max_rss_kib": r"^\s*Maximum resident set size \(kbytes\):\s*(\d+)$",
        "elapsed": r"^\s*Elapsed \(wall clock\) time \(h:mm:ss or m:ss\):\s*(\S+)$",
        "exit_status": r"^\s*Exit status:\s*(\d+)$",
    }
    for key, pattern in patterns.items():
        match = re.search(pattern, text, re.MULTILINE)
        if match is not None:
            values[key] = float(match.group(1)) if key.endswith("seconds") else (
                int(match.group(1)) if key in {"max_rss_kib", "exit_status"} else match.group(1))
    require("max_rss_kib" in values and "exit_status" in values,
            f"time log lacks required metrics: {path}")
    values["path"] = str(path)
    return values


def retain_external_or_packet(source: Path, packet_target: Path) -> dict[str, object]:
    identity = descriptor(source)
    if identity["bytes"] < MAX_PACKET_ARTIFACT:
        committed = copy_exclusive(source, packet_target)
        return {"storage": "packet", "descriptor": committed,
                "external": identity}
    return {"storage": "external", "descriptor": identity}


def report_json_from_stdout(path: Path) -> tuple[dict[str, Any] | None, str | None]:
    data = path.read_bytes()
    try:
        text = data.decode("utf-8")
    except UnicodeDecodeError as error:
        return None, f"stdout is not UTF-8: {error}"
    # Probe output is one JSON object, but preserve diagnostics before/after it
    # by selecting the last complete non-empty line.
    candidates = [line.strip() for line in text.splitlines() if line.strip()]
    if not candidates:
        return None, "probe produced empty stdout"
    for candidate in reversed(candidates):
        try:
            value = json.loads(candidate)
        except json.JSONDecodeError:
            continue
        if isinstance(value, dict):
            return value, None
    return None, "probe stdout contains no JSON object"


def integer_field(value: Any, path: str, minimum: int = 0) -> int:
    require(type(value) is int and value >= minimum, f"{path}: expected integer >= {minimum}")
    return value


def validate_stage(stage: Any, index: int, retention: str, edits: int,
                   ceiling: int) -> None:
    require(isinstance(stage, dict), f"stage {index}: not an object")
    require(type(stage.get("stage")) is int and stage.get("stage") == index,
            f"stage {index}: sequence mismatch")
    # The probe may call the two retained capacities before/after a mutation
    # ``before`` and ``after``.  No stage may be silently omitted.
    for key in ("before", "after", "after_release"):
        value = stage.get(key)
        require(value is None or (type(value) is int and 0 <= value <= ceiling),
                f"stage {index}: invalid {key}")
    before = stage.get("before")
    after = stage.get("after")
    after_release = stage.get("after_release")
    if retention == "zero":
        require(before is None and after is None and after_release is None,
                f"stage {index}: zero route retained render")
    elif retention == "release":
        require(after is not None and after_release is None,
                f"stage {index}: release route did not clear render")
    else:
        require(after is not None and after_release is None,
                f"stage {index}: default route stage retention mismatch")
    if index == 1:
        require(before is None, "first stage has a retained predecessor")
    else:
        if retention == "default":
            require(before is not None, f"stage {index}: default route lost predecessor")
        else:
            require(before is None, f"stage {index}: released/zero route retained predecessor")


def gate_value(gates: dict[str, Any], names: tuple[str, ...], label: str) -> None:
    """Accept the probe's stable spelling or one documented equivalent."""
    for name in names:
        if name in gates:
            value = gates[name]
            if isinstance(value, dict):
                value = value.get("ok", value.get("passed"))
            require(value is True, f"{label}: gate failed: {name}")
            return
    fail(f"{label}: missing gate ({'/'.join(names)})")


def sample_allocation(sample: dict[str, Any], report: dict[str, Any]) -> dict[str, Any] | None:
    value = sample.get("allocation")
    if value is None and isinstance(sample.get("allocations"), dict):
        value = sample["allocations"].get("whole")
    return value if isinstance(value, dict) else None


def sample_stages(sample: dict[str, Any], report: dict[str, Any]) -> list[Any] | None:
    value = sample.get("stages", sample.get("stage_observations"))
    return value if isinstance(value, list) else None


def validate_report(report: Any, spec: dict[str, object], fixture: dict[str, Any],
                    lane: str, *, expected_warmups: int | None = None,
                    expected_samples: int | None = None) -> dict[str, object]:
    require(isinstance(report, dict), f"{spec['label']}: report is not an object")
    p = protocol()
    if expected_warmups is None:
        expected_warmups = int(p["native"]["warmups_per_child"] if lane == "native"
                               else p["allocator"]["warmups_per_child"])
    if expected_samples is None:
        expected_samples = int(p["native"]["samples_per_child"] if lane == "native"
                               else p["allocator"]["samples_per_child"])
    require(type(expected_warmups) is int and expected_warmups >= 0,
            f"{spec['label']}: invalid expected warmup count")
    require(type(expected_samples) is int and expected_samples >= 1,
            f"{spec['label']}: invalid expected sample count")
    # The probe's stable identity fields are intentionally checked before any
    # metric is consumed.  This catches a wrong fixture/count/route quickly.
    require(type(report.get("schema_version")) is int
            and report.get("schema_version") == 1
            and report.get("schema") == "litchi.performance.0917.doc_retention.v1",
            f"{spec['label']}: report schema changed")
    require(report.get("change") == "0917", f"{spec['label']}: report change mismatch")
    require(report.get("case") == spec["case"]
            and type(report.get("edits")) is int
            and report.get("edits") == spec["edits"],
            f"{spec['label']}: case or edit count mismatch")
    require(type(report.get("selected_paragraph")) is int
            and report.get("selected_paragraph") == 0,
            f"{spec['label']}: selected paragraph changed")
    input_path = report.get("input")
    require(isinstance(input_path, str)
            and input_path in {fixture["path"], str(ROOT / fixture["path"]),
                               fixture.get("path_abs")},
            f"{spec['label']}: fixture path mismatch")
    retention_name = str(spec["retention"])
    expected_retention = {"zero": "zero", "default": "default_8mib",
                          "release": "release_8mib"}[retention_name]
    require(report.get("retention") == expected_retention,
            f"{spec['label']}: retention identity mismatch")
    source_hash = digest(report.get("source_sha256"),
                         f"{spec['label']}: source_sha256")
    source_bytes = integer_field(report.get("source_bytes"),
                                 f"{spec['label']}: source_bytes", 1)
    require(source_hash == fixture["sha256"] and source_bytes == fixture["bytes"],
            f"{spec['label']}: source identity mismatch")
    require(report.get("alternating_texts") == [
        "litchi copy-through baseline replacement text",
        "litchi retained render second replacement text",
    ], f"{spec['label']}: edit sequence changed")
    require(report.get("allocator_timing_claim") is False,
            f"{spec['label']}: allocator timing claim must be false")
    output_hash = digest(report.get("output_sha256"),
                         f"{spec['label']}: output_sha256")
    reference_hash = digest(report.get("reference_output_sha256"),
                            f"{spec['label']}: reference_output_sha256")
    output_bytes = integer_field(report.get("output_bytes"),
                                 f"{spec['label']}: output_bytes", 1)
    reference_bytes = integer_field(report.get("reference_output_bytes"),
                                    f"{spec['label']}: reference_output_bytes", 1)
    output_capacity = integer_field(report.get("output_capacity"),
                                    f"{spec['label']}: output_capacity", 1)
    require(output_hash == reference_hash and output_bytes == reference_bytes
            and output_capacity >= output_bytes
            and report.get("exact_output_equal_reference") is True,
            f"{spec['label']}: output equality gate failed")
    ceiling = 0 if retention_name == "zero" else 8 * 1024 * 1024
    require(type(report.get("retention_ceiling_bytes")) is int
            and report.get("retention_ceiling_bytes") == ceiling,
            f"{spec['label']}: retention ceiling mismatch")
    top_elapsed = report.get("elapsed_ns")
    require(isinstance(top_elapsed, list), f"{spec['label']}: elapsed_ns vector missing")
    if lane == "native":
        require(len(top_elapsed) == expected_samples
                and all(type(x) is int and x > 0 for x in top_elapsed),
                f"{spec['label']}: top-level elapsed_ns vector mismatch")
    else:
        require(top_elapsed == [], f"{spec['label']}: allocator emitted timing vector")
    route_outputs = report.get("route_outputs")
    require(isinstance(route_outputs, list) and len(route_outputs) == 3,
            f"{spec['label']}: route output set mismatch")
    expected_routes = ["zero", "default_8mib", "release_8mib"]
    route_by_name: dict[str, dict[str, Any]] = {}
    for route_output in route_outputs:
        require(isinstance(route_output, dict), f"{spec['label']}: route output malformed")
        route = route_output.get("retention")
        require(route in expected_routes and route not in route_by_name,
                f"{spec['label']}: duplicate route output")
        route_by_name[str(route)] = route_output
        route_ceiling = route_output.get("retention_ceiling_bytes")
        require(type(route_ceiling) is int and route_ceiling ==
                (0 if route == "zero" else 8 * 1024 * 1024),
                f"{spec['label']}: route ceiling mismatch")
        route_hash = digest(route_output.get("output_sha256"),
                            f"{spec['label']}: route output_sha256")
        route_bytes = integer_field(route_output.get("output_bytes"),
                                    f"{spec['label']}: route output_bytes", 1)
        route_capacity = integer_field(route_output.get("output_capacity"),
                                       f"{spec['label']}: route output_capacity", 1)
        require(route_hash == reference_hash and route_bytes == reference_bytes,
                f"{spec['label']}: cross-route output mismatch")
        require(route_capacity >= route_bytes,
                f"{spec['label']}: route output capacity invalid")
        route_stages = route_output.get("stages")
        require(isinstance(route_stages, list) and len(route_stages) == 32,
                f"{spec['label']}: route fixed32 stages missing")
        route_name = {"zero": "zero", "default_8mib": "default",
                      "release_8mib": "release"}[str(route)]
        require(type(route_output.get("semantic_change_count")) is int
                and route_output.get("semantic_change_count") == spec["edits"],
                f"{spec['label']}: route semantic change count mismatch")
        for stage_index, stage in enumerate(route_stages[:int(spec["edits"])], start=1):
            validate_stage(stage, stage_index, route_name, int(spec["edits"]),
                           0 if route_name == "zero" else 8 * 1024 * 1024)
            require(stage.get("requested_text") == (
                "litchi copy-through baseline replacement text" if stage_index % 2 else
                "litchi retained render second replacement text"),
                    f"{spec['label']}: stage text sequence mismatch")
        for stage in route_stages[int(spec["edits"]):]:
            require(type(stage.get("stage")) is int and stage.get("stage") == 0
                    and stage.get("requested_text") == ""
                    and stage.get("before") is None and stage.get("after") is None
                    and stage.get("after_release") is None,
                    f"{spec['label']}: inactive stage is populated")
        route_final = route_output.get("final_retained_before_commit")
        require(route_final is None or (type(route_final) is int and 0 <= route_final <=
                                        route_ceiling),
                f"{spec['label']}: route final capacity invalid")
        if route_name in {"zero", "release"}:
            require(route_final is None, f"{spec['label']}: cleared route retained final render")
        else:
            require(route_final is not None, f"{spec['label']}: default final capacity missing")
    require(set(route_by_name) == set(expected_routes), f"{spec['label']}: route set incomplete")
    selected_route_output = route_by_name[expected_retention]
    require(report.get("output_capacity") == selected_route_output.get("output_capacity"),
            f"{spec['label']}: selected output capacity mismatch")
    samples = report.get("samples")
    if lane == "native":
        require(isinstance(samples, list) and len(samples) == expected_samples,
                f"{spec['label']}: native sample count mismatch")
        require(type(report.get("warmups")) is int
                and type(report.get("samples_requested")) is int
                and report.get("warmups") == expected_warmups
                and report.get("samples_requested") == expected_samples,
                f"{spec['label']}: warmup/sample request mismatch")
        expected_timing = True
    else:
        require(type(report.get("warmups")) is int
                and type(report.get("samples_requested")) is int
                and report.get("warmups") == expected_warmups
                and report.get("samples_requested") == expected_samples,
                f"{spec['label']}: allocator request mismatch")
        require(isinstance(samples, list) and len(samples) == expected_samples
                and isinstance(samples[0], dict)
                and type(samples[0].get("index")) is int
                and samples[0].get("index") == 0,
                f"{spec['label']}: allocator sample count mismatch")
        expected_timing = False
    require(report.get("timing_claim") is expected_timing,
            f"{spec['label']}: timing claim mismatch")
    require(report.get("allocator_instrumented") is (lane == "allocator"),
            f"{spec['label']}: allocator lane mismatch")
    require(isinstance(report.get("region_definition"), str)
            or isinstance(report.get("phase_contract"), str),
            f"{spec['label']}: region definition missing")
    all_output_hashes: list[str] = []
    all_stage_lists: list[list[Any]] = []
    all_allocations: list[dict[str, Any] | None] = []
    for index, sample in enumerate(samples):
        require(isinstance(sample, dict) and type(sample.get("index")) is int
                and sample.get("index") == index,
                f"{spec['label']}: sample index mismatch")
        if lane == "native":
            integer_field(sample.get("elapsed_ns"), f"{spec['label']}: sample elapsed", 1)
            require(sample["elapsed_ns"] == top_elapsed[index],
                    f"{spec['label']}: sample/top elapsed mismatch")
        output_sha = digest(sample.get("output_sha256"),
                            f"{spec['label']}: sample output_sha256")
        reference_sha = reference_hash
        exact = sample.get("exact_output_equal_reference")
        require(output_sha == reference_sha and exact is True,
                f"{spec['label']}: sample output equality gate failed")
        all_output_hashes.append(output_sha)
        sample_bytes = integer_field(sample.get("output_bytes"),
                                     f"{spec['label']}: sample output_bytes", 1)
        sample_capacity = integer_field(sample.get("output_capacity"),
                                        f"{spec['label']}: sample output_capacity", 1)
        require(sample_bytes == output_bytes and sample_capacity == output_capacity,
                f"{spec['label']}: sample output size mismatch")
        require(type(sample.get("semantic_change_count")) is int
                and sample.get("semantic_change_count") == spec["edits"],
                f"{spec['label']}: sample semantic change count mismatch")
        stages = sample_stages(sample, report)
        require(isinstance(stages, list) and len(stages) == 32,
                f"{spec['label']}: fixed32 stage count mismatch")
        for stage_index, stage in enumerate(stages[:int(spec["edits"])], start=1):
            validate_stage(stage, stage_index, retention_name, int(spec["edits"]), ceiling)
            require(stage.get("requested_text") == (
                "litchi copy-through baseline replacement text" if stage_index % 2 else
                "litchi retained render second replacement text"),
                    f"{spec['label']}: sample stage text sequence mismatch")
        for stage_index, stage in enumerate(stages[int(spec["edits"]):], start=int(spec["edits"]) + 1):
            require(isinstance(stage, dict) and type(stage.get("stage")) is int
                    and stage.get("stage") == 0
                    and stage.get("requested_text") == ""
                    and stage.get("before") is None and stage.get("after") is None
                    and stage.get("after_release") is None,
                    f"{spec['label']}: inactive stage {stage_index} is populated")
        all_stage_lists.append(stages)
        allocation = sample_allocation(sample, report)
        all_allocations.append(allocation)
        if lane == "allocator":
            require(isinstance(allocation, dict), f"{spec['label']}: allocation report missing")
            for key in ("allocated_bytes", "deallocated_bytes", "allocation_calls",
                        "peak_live_bytes"):
                integer_field(allocation.get(key), f"{spec['label']}: allocation.{key}")
            require(type(allocation.get("retained_bytes")) is int
                    and allocation["retained_bytes"] >= 0,
                    f"{spec['label']}: allocation retained bytes invalid")
            require(allocation["retained_bytes"] == sample_capacity,
                        f"{spec['label']}: retained allocation is not output capacity")
        else:
            require(allocation is None, f"{spec['label']}: native report carries allocation metrics")
        sample_oracle = sample.get("oracle")
        require(sample.get("final_paragraph_semantic_readback") is True
                and sample.get("final_public_reader_reopen") is True
                and sample.get("retained_capacities_within_ceiling") is True
                and sample.get("allocator_retained_matches_output_capacity") is True
                and isinstance(sample_oracle, dict)
                and sample_oracle.get("oracle_ok") is True,
                f"{spec['label']}: sample correctness gate failed")
    gates = report.get("correctness")
    require(isinstance(gates, dict), f"{spec['label']}: correctness gates missing")
    for name in (
        "all_stages_change_text", "exact_cross_route_outputs",
        "final_paragraph_semantic_readback", "final_public_reader_reopen",
        "exact_noop_source",
        "same_value_preserves_retention", "operation_limit_refusal_atomic",
        "noop_shares_source_allocation", "forward_patch_exact_output",
        "inverse_patch_exact_source", "foreign_patch_refused",
        "stale_patch_refused", "retained_capacities_within_ceiling",
        "zero_route_clears_retention", "release_route_clears_every_stage",
        "untouched_streams_preserved", "directory_metadata_preserved",
        "semantic_oracles_passed", "semantic_change_count_matches_edit_count",
        "allocator_retained_matches_output_capacity",
    ):
        require(gates.get(name) is True, f"{spec['label']}: correctness gate failed: {name}")
    top_allocation = report.get("allocation")
    if lane == "allocator":
        require(isinstance(top_allocation, dict), f"{spec['label']}: top allocation missing")
        for key in ("allocated_bytes", "deallocated_bytes", "allocation_calls",
                    "peak_live_bytes"):
            integer_field(top_allocation.get(key), f"{spec['label']}: top allocation.{key}")
        require(type(top_allocation.get("retained_bytes")) is int
                and top_allocation["retained_bytes"] >= 0
                and top_allocation["retained_bytes"] == report.get("output_capacity"),
                f"{spec['label']}: top retained allocation mismatch")
    else:
        require(top_allocation is None, f"{spec['label']}: native top allocation present")
    require(len(set(all_output_hashes)) == 1, f"{spec['label']}: sample outputs differ")
    require(output_hash == all_output_hashes[0],
            f"{spec['label']}: top-level output differs")
    allocation = all_allocations[0] if all_allocations else None
    return {"case": spec["case"], "edits": spec["edits"], "retention": retention_name,
            "report_sha256": sha256_bytes(json_bytes(report)),
            "output_sha256": output_hash, "output_bytes": output_bytes,
            "output_capacity": output_capacity, "stages": all_stage_lists[0],
            "allocation": allocation}


def prior_schedule_items(schedule: list[dict[str, object]], spec: dict[str, object]) -> Iterable[dict[str, object]]:
    index = next(i for i, value in enumerate(schedule) if value["label"] == spec["label"])
    return schedule[:index]


def child_paths(m: dict[str, Any], spec: dict[str, object]) -> dict[str, Path]:
    label = safe_label(str(spec["label"]))
    child = EXTERNAL_CHILDREN / label
    return {"child": child, "stdout": child / "stdout.jsonl", "stderr": child / "stderr.txt",
            "process": child / "process.txt", "start": child / "start.json",
            "finish": child / "finish.json"}


def run_child(m: dict[str, Any], spec: dict[str, object], schedule: list[dict[str, object]]) -> None:
    paths = child_paths(m, spec)
    child = paths["child"]
    require(not child.exists(), f"refusing to replace external child: {child}")
    label = str(spec["label"])
    receipt_path = PACKET_LOGS / f"{label}.json"
    require(not receipt_path.exists(), f"refusing to replace child receipt: {receipt_path}")
    for previous in prior_schedule_items(schedule, spec):
        previous_receipt = PACKET_LOGS / f"{previous['label']}.json"
        require(previous_receipt.is_file(), f"prior child missing: {previous['label']}")
        prior = read_json(previous_receipt)
        require(prior.get("status") == "passed", f"prior child failed: {previous['label']}")
    child.mkdir(parents=True)
    command = command_for(m, spec, child)
    env = {"PATH": "/usr/bin:/bin", "LC_ALL": "C", "LANG": "C", "TZ": "UTC",
           "RUST_BACKTRACE": "0"}
    started = now()
    start_record = {"schema": "litchi.performance.0917.child-start.v1", "label": label,
                    "spec": spec, "command": command, "cwd": str(ROOT),
                    "environment": env, "manifest": manifest_descriptor(),
                    "protocol": protocol_descriptor(), "host_load_before": list(os.getloadavg()),
                    "started_utc": started}
    write_json(paths["start"], start_record)
    exit_code = run_process(command, ROOT, env, paths["stdout"], paths["stderr"], paths["process"])
    finished = now()
    report, parse_error = report_json_from_stdout(paths["stdout"])
    fixture = fixture_rows(protocol())[str(spec["case"])]
    validation: dict[str, object] | None = None
    validation_error = parse_error
    if report is not None:
        try:
            validation = validate_report(report, spec, fixture, str(spec["lane"]))
            validation_error = None
        except Exception as error:
            # Preserve every malformed attempt as a failed receipt rather
            # than allowing an unexpected report type to abort before custody
            # metadata is written.
            validation_error = f"{type(error).__name__}: {error}"
    if exit_code != 0:
        validation_error = validation_error or f"probe exit status {exit_code}"
    process_metrics: dict[str, object] | None = None
    try:
        process_metrics = parse_time_v(paths["process"])
    except EvidenceError as error:
        validation_error = validation_error or str(error)
    finish_record = {"schema": "litchi.performance.0917.child-finish.v1", "label": label,
                     "finished_utc": finished, "exit_code": exit_code,
                     "host_load_after": list(os.getloadavg()),
                     "stdout": descriptor(paths["stdout"]), "stderr": descriptor(paths["stderr"]),
                     "process": descriptor(paths["process"]), "process_metrics": process_metrics,
                     "report": validation, "validation_error": validation_error}
    write_json(paths["finish"], finish_record)
    retained_stdout = retain_external_or_packet(paths["stdout"], PACKET_RAW / f"{label}.stdout")
    retained_stderr = retain_external_or_packet(paths["stderr"], PACKET_RAW / f"{label}.stderr")
    retained_process = retain_external_or_packet(paths["process"], PACKET_RAW / f"{label}.time")
    receipt = {"schema": "litchi.performance.0917.child-receipt.v1", "label": label,
               "spec": spec, "started_utc": started, "finished_utc": finished,
               "command": command, "cwd": str(ROOT), "environment": env,
               "protocol": protocol_descriptor(), "manifest": manifest_descriptor(),
               "binary": m["binaries"]["native" if spec["lane"] == "native" else "allocator"],
               "host_load_before": start_record["host_load_before"],
               "host_load_after": finish_record["host_load_after"],
               "exit_code": exit_code, "process_metrics": process_metrics,
               "stdout": retained_stdout, "stderr": retained_stderr, "process": retained_process,
               "report": validation, "validation_error": validation_error,
               "status": "passed" if exit_code == 0 and validation_error is None else "failed"}
    write_json(receipt_path, receipt)
    if receipt["status"] != "passed":
        fail(f"{label} failed; all artifacts retained")
    print(json.dumps({"status": "child-complete", "label": label}, sort_keys=True), flush=True)


def run(args: argparse.Namespace) -> None:
    m = manifest()
    p = protocol()
    lane = args.lane
    schedule = all_schedule(p, lane)
    selected = schedule
    if args.label:
        selected = [x for x in schedule if x["label"] == args.label]
        require(selected, f"unknown label: {args.label}")
    if args.round is not None:
        selected = [x for x in selected if x.get("round", x.get("repetition")) == args.round]
        require(selected, f"unknown {lane} round/repetition: {args.round}")
    errors: list[str] = []
    for spec in selected:
        try:
            run_child(m, spec, schedule)
        except EvidenceError as error:
            errors.append(str(error))
            print(f"capture failed: {error}", file=sys.stderr)
            # Never continue past a failed child: schedule ordering remains
            # auditable and the failed row is retained for review.
            break
    if errors:
        fail("capture stopped after a failed child; retained all attempted artifacts")


def nearest_rank(values: list[float], fraction: float) -> float:
    require(values, "empty percentile vector")
    ordered = sorted(values)
    index = min(max(math.ceil(fraction * len(ordered)) - 1, 0), len(ordered) - 1)
    return ordered[index]


def stats(values: list[float]) -> dict[str, float]:
    require(values, "empty metric vector")
    return {"p50": nearest_rank(values, 0.50), "p95": nearest_rank(values, 0.95),
            "p99": nearest_rank(values, 0.99), "mean": statistics.fmean(values)}


def bootstrap(ratios: list[float], draws: int, seed: int) -> dict[str, object]:
    require(ratios, "empty bootstrap ratio vector")
    rng = random.Random(seed)
    medians = [statistics.median(ratios[rng.randrange(len(ratios))] for _ in ratios)
               for _ in range(draws)]
    return {"draws": draws, "seed": seed,
            "statistic": "median of per-round zero-p50/route-p50 ratios",
            "ratios": ratios, "median_ratio": statistics.median(ratios),
            "confidence_interval_95": {"method": "nearest-rank2.5th/97.5th",
                                        "lower": nearest_rank(medians, 0.025),
                                        "upper": nearest_rank(medians, 0.975)}}


def read_retained_file(item: dict[str, object], label: str) -> Path:
    require(isinstance(item, dict), f"{label}: artifact descriptor missing")
    storage = item.get("storage")
    identity = item.get("descriptor")
    require(storage in {"packet", "external"} and isinstance(identity, dict),
            f"{label}: artifact descriptor malformed")
    path = manifest_path(identity)
    verify_descriptor(identity, label)
    original = item.get("external")
    if original is not None:
        require(isinstance(original, dict), f"{label}: original descriptor malformed")
        require(original.get("bytes") == identity.get("bytes")
                and original.get("sha256") == identity.get("sha256"),
                f"{label}: original artifact hash differs")
        digest(original.get("sha256"), f"{label}: original sha256")
    return path


def load_child(m: dict[str, Any], spec: dict[str, object], schedule: list[dict[str, object]]) -> dict[str, Any]:
    label = str(spec["label"])
    path = PACKET_LOGS / f"{label}.json"
    value = read_json(path)
    require(isinstance(value, dict), f"{label}: receipt is not an object")
    require(value.get("schema") == "litchi.performance.0917.child-receipt.v1",
            f"{label}: receipt schema mismatch")
    require(value.get("status") == "passed" and value.get("exit_code") == 0,
            f"{label}: failed child retained")
    require(value.get("spec") == spec and value.get("protocol") == m["protocol"],
            f"{label}: schedule/protocol binding mismatch")
    started = parse_utc(value.get("started_utc"), f"{label} start")
    finished = parse_utc(value.get("finished_utc"), f"{label} finish")
    require(started <= finished, f"{label}: finish precedes start")
    for previous in prior_schedule_items(schedule, spec):
        previous_value = read_json(PACKET_LOGS / f"{previous['label']}.json")
        require(parse_utc(previous_value.get("finished_utc"), f"{previous['label']} finish") <= started,
                f"{label}: schedule overlap")
    expected_manifest = manifest_descriptor()
    require(value.get("manifest") == expected_manifest,
            f"{label}: manifest provenance mismatch")
    require(value.get("protocol") == m["protocol"],
            f"{label}: protocol provenance mismatch")
    expected_binary = m["binaries"]["native" if spec["lane"] == "native" else "allocator"]
    require(value.get("binary") == expected_binary,
            f"{label}: binary provenance mismatch")
    verify_descriptor(value.get("manifest"), f"{label}: manifest")
    verify_descriptor(value.get("protocol"), f"{label}: protocol")
    verify_descriptor(value.get("binary"), f"{label}: binary", allow_missing_binary=True)
    require(value.get("command") == command_for(m, spec, EXTERNAL_CHILDREN / label),
            f"{label}: command provenance mismatch")
    for key in ("stdout", "stderr", "process"):
        read_retained_file(value.get(key), f"{label}: {key}")
    metrics = value.get("process_metrics")
    require(isinstance(metrics, dict), f"{label}: process metrics missing")
    process_path = read_retained_file(value["process"], f"{label}: process")
    actual_metrics = parse_time_v(process_path)
    require({k: v for k, v in actual_metrics.items() if k != "path"} ==
            {k: v for k, v in metrics.items() if k != "path"},
            f"{label}: process metrics changed")
    report = value.get("report")
    require(isinstance(report, dict),
            f"{label}: report validation missing")
    # Re-open retained stdout and validate the report again; this detects a
    # receipt that was edited independently from the captured raw report.
    stdout = read_retained_file(value["stdout"], f"{label}: stdout")
    raw, parse_error = report_json_from_stdout(stdout)
    require(parse_error is None and raw is not None, f"{label}: retained stdout invalid")
    expected = validate_report(raw, spec, fixture_rows(protocol())[str(spec["case"])],
                               str(spec["lane"]))
    require(expected == report, f"{label}: receipt validation differs from raw report")
    return {"receipt": value, "report": raw, "validation": expected}


def fraction_flag(before: float, after: float, limit: float) -> dict[str, object]:
    increase = float("inf") if before == 0 and after > 0 else (after / before - 1.0 if before else 0.0)
    return {"before": before, "after": after, "increase_fraction": increase,
            "flag": increase > limit}


def metric_samples(report: dict[str, Any]) -> list[float]:
    samples = report.get("samples")
    require(isinstance(samples, list) and samples, "native timing samples missing")
    values = []
    for sample in samples:
        if isinstance(sample, dict):
            values.append(float(sample["elapsed_ns"]))
        else:
            values.append(float(sample))
    return values


def reduction_for_lane(m: dict[str, Any], lane: str) -> dict[str, object]:
    p = protocol()
    review_flags = p["review_flags"]
    schedule = all_schedule(p, lane)
    loaded = [load_child(m, spec, schedule) for spec in schedule]
    if lane == "allocator":
        by_key: dict[tuple[str, int, str], list[dict[str, Any]]] = {}
        for spec, item in zip(schedule, loaded, strict=True):
            key = (str(spec["case"]), int(spec["edits"]), str(spec["retention"]))
            by_key.setdefault(key, []).append(item["report"])
        cells = []
        comparisons = []
        for key, reports in sorted(by_key.items()):
            rows = []
            for report in reports:
                alloc = report["allocation"]
                rows.append({k: alloc[k] for k in ("allocated_bytes", "deallocated_bytes",
                                                    "allocation_calls", "peak_live_bytes",
                                                    "retained_bytes")})
            # Equality is checked only when observed; variability is retained.
            cells.append({"case": key[0], "edits": key[1], "retention": key[2],
                          "repetitions": rows,
                          "repeat_equal": len({json.dumps(x, sort_keys=True) for x in rows}) == 1,
                          "variability": {field: stats([float(x[field]) for x in rows])
                                          for field in rows[0]}})
        # Pair each retention route with the zero route within the same
        # repetition.  Allocation values are observations; variability is
        # retained even when the two repetitions disagree.
        for case in [str(x["case"]) for x in p["fixtures"]]:
            for edits in p["edits"]:
                for repetition in (1, 2):
                    def one_item(route: str) -> dict[str, Any]:
                        spec = next(x for x in allocation_schedule(p)
                                    if x["repetition"] == repetition and x["case"] == case
                                    and x["edits"] == edits and x["retention"] == route)
                        return loaded[schedule.index(spec)]
                    zero_item = one_item("zero")
                    zero = zero_item["report"]["allocation"]
                    for route in ("default", "release"):
                        route_item = one_item(route)
                        value = route_item["report"]["allocation"]
                        comparisons.append({
                            "case": case, "edits": edits, "repetition": repetition,
                            "route": route,
                            "peak_live": fraction_flag(float(zero["peak_live_bytes"]),
                                                         float(value["peak_live_bytes"]),
                                                         review_flags["peak_live_increase_fraction"]),
                            "retained": fraction_flag(float(zero["retained_bytes"]),
                                                       float(value["retained_bytes"]),
                                                       review_flags["peak_live_increase_fraction"]),
                            "allocated": fraction_flag(float(zero["allocated_bytes"]),
                                                        float(value["allocated_bytes"]),
                                                        review_flags["peak_live_increase_fraction"]),
                            "rss": fraction_flag(
                                float(zero_item["receipt"]["process_metrics"]["max_rss_kib"]),
                                float(route_item["receipt"]["process_metrics"]["max_rss_kib"]),
                                review_flags["rss_increase_fraction"]),
                        })
        return {"schema": "litchi.performance.0917.allocator-reduction.v1",
                "children": len(loaded), "cells": cells, "comparisons": comparisons,
                "claim": "allocator region metrics only; no latency claim"}
    # Native cells pair all six rounds by fixture/count/route, preserving each
    # route's own p50/p95/p99 and RSS values.
    by_cell: dict[tuple[str, int, str], list[tuple[dict[str, object], dict[str, Any]]]] = {}
    for spec, item in zip(schedule, loaded, strict=True):
        key = (str(spec["case"]), int(spec["edits"]), str(spec["retention"]))
        by_cell.setdefault(key, []).append((spec, item))
    cells = []
    flags = []
    for case in [str(x["case"]) for x in p["fixtures"]]:
        for edits in p["edits"]:
            baseline_key = (case, int(edits), "zero")
            baseline_rows = by_cell[baseline_key]
            require(len(baseline_rows) == 6, f"{case}/e{edits}: zero round count")
            route_summary: dict[str, object] = {}
            for route in ("zero", "default", "release"):
                rows = by_cell[(case, int(edits), route)]
                require(len(rows) == 6, f"{case}/e{edits}/{route}: round count")
                per_round = []
                ratios = []
                for round_index, ((spec, item), (_, base_item)) in enumerate(zip(rows, baseline_rows), start=1):
                    values = metric_samples(item["report"])
                    metric = stats(values)
                    base = stats(metric_samples(base_item["report"]))
                    ratios.append(base["p50"] / metric["p50"] if metric["p50"] else float("inf"))
                    rss = float(item["receipt"]["process_metrics"]["max_rss_kib"])
                    base_rss = float(base_item["receipt"]["process_metrics"]["max_rss_kib"])
                    per_round.append({"round": round_index, "label": spec["label"],
                                      "stats": metric, "rss_kib": rss,
                                      "zero_rss_kib": base_rss,
                                      "rss_flag": fraction_flag(base_rss, rss,
                                                                 review_flags["rss_increase_fraction"])})
                pooled = stats([x for _, item in rows for x in metric_samples(item["report"])])
                zero_pooled = stats([x for _, item in baseline_rows for x in metric_samples(item["report"])])
                rss_values = [float(item["receipt"]["process_metrics"]["max_rss_kib"]) for _, item in rows]
                zero_rss_values = [float(item["receipt"]["process_metrics"]["max_rss_kib"]) for _, item in baseline_rows]
                route_summary[route] = {"pooled": pooled, "rounds": per_round,
                                        "pooled_rss_kib": stats(rss_values),
                                        "rss_flag": fraction_flag(statistics.median(zero_rss_values),
                                                                    statistics.median(rss_values),
                                                                    review_flags["rss_increase_fraction"]),
                                        "ratios": bootstrap(ratios, 10000, 917)}
                if route != "zero":
                    latency_flags = {field: fraction_flag(
                        zero_pooled[field], pooled[field], review_flags["latency_increase_fraction"])
                                     for field in ("p50", "p95", "p99")}
                    peak_rounds = []
                    for base_item, (_, item) in zip(baseline_rows, rows, strict=True):
                        base_alloc = base_item["report"].get("allocation")
                        alloc = item["report"].get("allocation")
                        if base_alloc and alloc:
                            peak_rounds.append(fraction_flag(
                                float(base_alloc["peak_live_bytes"]),
                                float(alloc["peak_live_bytes"]),
                                review_flags["peak_live_increase_fraction"]))
                    peak_flags: dict[str, object] = {
                        "available": bool(peak_rounds),
                        "reason": "native lane has no allocator instrumentation",
                    }
                    if peak_rounds:
                        peak_flags["rounds"] = peak_rounds
                    flags.append({"case": case, "edits": edits, "route": route,
                                  "latency": latency_flags, "peak_live": peak_flags,
                                  "rss": route_summary[route]["rss_flag"]})
            cells.append({"case": case, "edits": edits, "routes": route_summary})
    return {"schema": "litchi.performance.0917.native-reduction.v1", "children": len(loaded),
            "cells": cells, "review_flags": flags,
            "bootstrap": p["bootstrap"],
            "claim": "same-build route comparison; no optimization adoption gate"}


def reduce(args: argparse.Namespace) -> None:
    m = manifest()
    native = reduction_for_lane(m, "native")
    allocator = reduction_for_lane(m, "allocator")
    result = {"schema": "litchi.performance.0917.reduction.v1", "protocol": m["protocol"],
              "manifest": manifest_descriptor(), "native": native,
              "allocator": allocator, "generated_utc": now(),
              "adoption": "No optimization or production adoption gate; review flags are descriptive."}
    output = args.output.resolve() if args.output else REDUCED_PATH
    write_json(output, result)
    output_identity = packet_descriptor(output, allow_large=False) if repository_relative(output) \
        is not None else descriptor(output, allow_large=False)
    print(json.dumps({"status": "reduced", "output": output_identity}, sort_keys=True))


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="action", required=True)
    freeze_parser = sub.add_parser("freeze")
    freeze_parser.add_argument("--binary", type=Path, required=True)
    freeze_parser.add_argument("--allocator-binary", type=Path, required=True)
    freeze_parser.add_argument("--probe-source", type=Path)
    freeze_parser.add_argument("--source-receipt", type=Path)
    freeze_parser.add_argument("--build-receipt", type=Path)
    run_parser = sub.add_parser("run")
    run_parser.add_argument("--lane", choices=("native", "allocator"), required=True)
    run_parser.add_argument("--round", type=int)
    run_parser.add_argument("--label")
    reduce_parser = sub.add_parser("reduce")
    reduce_parser.add_argument("--output", type=Path)
    validate_parser = sub.add_parser(
        "validate-report",
        help="validate one qualification report without changing frozen packet state",
    )
    validate_parser.add_argument("--report", type=Path, required=True)
    validate_parser.add_argument("--lane", choices=("native", "allocator"), required=True)
    validate_parser.add_argument("--case", choices=("docnohf", "docfloat"), required=True)
    validate_parser.add_argument("--edits", type=int, choices=(1, 2, 8, 32), required=True)
    validate_parser.add_argument("--retention", choices=("zero", "default", "release"), required=True)
    validate_parser.add_argument("--warmup", type=int, required=True)
    validate_parser.add_argument("--samples", type=int, required=True)
    sub.add_parser("replay", help="use replay.py; retained for command discoverability")
    args = parser.parse_args()
    try:
        if args.action == "freeze":
            freeze(args)
        elif args.action == "run":
            run(args)
        elif args.action == "reduce":
            reduce(args)
        elif args.action == "validate-report":
            p = protocol()
            fixture = fixture_rows(p)[args.case]
            spec = {"lane": args.lane, "case": args.case, "edits": args.edits,
                    "retention": args.retention,
                    "label": f"qualification-{args.case}-e{args.edits}-{args.retention}"}
            report = read_json(args.report)
            summary = validate_report(report, spec, fixture, args.lane,
                                      expected_warmups=args.warmup,
                                      expected_samples=args.samples)
            print(json.dumps({"status": "valid", "report": descriptor(args.report),
                              "case": args.case, "edits": args.edits,
                              "retention": args.retention,
                              "allocator": summary["allocation"] is not None}, sort_keys=True))
        else:
            require(REPLAY_PATH.is_file(), f"missing replay helper: {REPLAY_PATH}")
            replay_spec = importlib.util.spec_from_file_location("litchi_perf_0917_replay",
                                                                  REPLAY_PATH)
            require(replay_spec is not None and replay_spec.loader is not None,
                    "cannot load replay helper")
            replay_module = importlib.util.module_from_spec(replay_spec)
            replay_spec.loader.exec_module(replay_module)
            print(json.dumps(replay_module.replay(HERE), sort_keys=True))
    except (EvidenceError, OSError, ValueError, subprocess.SubprocessError) as error:
        print(f"0917 capture failed: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
