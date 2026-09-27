"""Pure offline audit for the 0781 observer lanes.

The capture driver is :mod:`observers`; this module never starts a command and
never writes an output file.  ``analyze`` only reads retained receipts and
reports.  It is intentionally separate from the primary 0781 native and
allocator analysis because the ``perf stat`` lane measures whole processes,
including probe setup and verification, while the heaptrack lane is an
allocation-site diagnostic.
"""

from __future__ import annotations

import hashlib
import json
import math
import re
import subprocess
import sys
from collections import Counter
from pathlib import Path
from typing import Any, Iterable


PACKET = Path(__file__).resolve().parent
SOURCE_FILE = "crates/litchi-ppt/src/writer/core/codec.rs"
LEGS = ("before", "after")
HEX = frozenset("0123456789abcdefABCDEF")
REPORT_SCHEMA = "litchi.ppt.borrowed-text-probe.v1"
REPORT_TOOL = "ppt-borrowed-text-probe-0781"
MARKER = "litchi-perf-0781-borrowed-text"
SHAPES = {
    "tiny": (2, 3),
    "many": (100, 10),
    "payload": (16, 4),
    "unicode": (16, 4),
    "rich": (2, 2),
}
CONVERSION_SYMBOL = "convert_shape_to_escher_with_sound_mapping"
TARGET_SYMBOLS = (CONVERSION_SYMBOL,)
MAX_TRACE_BYTES = 512 * 1024 * 1024
MAX_TRACE_RECORDS = 20_000_000
MAX_HISTOGRAM_ROWS = 10_000_000


class ObserverError(RuntimeError):
    """Raised when retained observer evidence is missing or contradictory."""


def fail(message: str) -> None:
    raise ObserverError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def read_json(path: Path, label: str) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing {label}: {path}")
    try:
        return json.loads(path.read_text())
    except (OSError, ValueError) as error:
        fail(f"invalid {label} {path}: {error}")


def digest(path: Path) -> str:
    value = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for chunk in iter(lambda: stream.read(1 << 20), b""):
                value.update(chunk)
    except OSError as error:
        fail(f"cannot hash retained artifact {path}: {error}")
    return value.hexdigest()


def is_sha(value: Any) -> bool:
    return isinstance(value, str) and len(value) == 64 and all(c in HEX for c in value)


def nonnegative_int(value: Any, label: str) -> None:
    require(isinstance(value, int) and not isinstance(value, bool) and value >= 0,
            f"{label} is not a non-negative integer")


def finite_number(value: Any, label: str) -> None:
    require(isinstance(value, (int, float)) and not isinstance(value, bool)
            and math.isfinite(float(value)), f"{label} is not finite")


def _path(value: Any, label: str) -> Path:
    require(isinstance(value, str) and value, f"{label}.path is missing")
    return Path(value)


class Context:
    """Packet paths plus the origin mapping needed after worktree cleanup."""

    def __init__(self, packet: Path):
        self.packet = packet.resolve()
        self.root = self.packet.parents[3]
        self.origin = read_json(self.packet / "origin.json", "origin.json")
        require(isinstance(self.origin, dict), "origin.json is not an object")
        owned = self.origin.get("owned_worktree")
        require(isinstance(owned, str) and owned, "origin owned worktree is missing")
        self.owned = Path(owned).resolve()

    def normalized(self, value: str) -> str:
        """Map only the archived owned-worktree prefix to this checkout."""

        old = str(self.owned)
        current = str(self.root.resolve())
        if value == old:
            return current
        if value.startswith(old + "/"):
            return current + value[len(old):]
        return value

    def candidates(self, value: str) -> list[Path]:
        raw = Path(value)
        candidates: list[Path] = []
        if raw.is_absolute():
            mapped = Path(self.normalized(value))
            if mapped != raw:
                candidates.append(mapped)
            candidates.append(raw)
            try:
                relative = raw.relative_to(self.packet)
            except ValueError:
                relative = None
            if relative is not None:
                candidates.append(self.packet / relative)
        else:
            text = value.replace("\\", "/")
            prefix = "docs/performance/results/change-0781/"
            if text.startswith(prefix):
                candidates.append(self.packet / text[len(prefix):])
            candidates.extend((self.packet / raw, self.root / raw))
        unique: list[Path] = []
        for candidate in candidates:
            candidate = candidate.resolve(strict=False)
            if candidate not in unique:
                unique.append(candidate)
        return unique

    def resolve(self, value: str, *, packet_bound: bool) -> Path:
        candidates = self.candidates(value)
        for candidate in candidates:
            if not candidate.is_file() or candidate.is_symlink():
                continue
            if packet_bound:
                try:
                    candidate.relative_to(self.packet)
                except ValueError:
                    continue
            return candidate.resolve()
        path = (candidates[0] if candidates else Path(value)).resolve(strict=False)
        if packet_bound:
            try:
                path.relative_to(self.packet)
            except ValueError:
                fail(f"artifact path escaped packet: {value}")
        return path

    def relative(self, path: Path) -> str:
        try:
            return str(path.relative_to(self.packet))
        except ValueError:
            return str(path)

    def artifact(self, value: Any, label: str, *, packet_bound: bool = True,
                 allow_missing: bool = False) -> tuple[Path | None, dict[str, Any]]:
        require(isinstance(value, dict), f"{label} is not an artifact receipt")
        raw = _path(value.get("path"), label)
        size = value.get("bytes", value.get("size"))
        nonnegative_int(size, f"{label}.bytes")
        sha = value.get("sha256", value.get("digest"))
        require(is_sha(sha), f"{label}.sha256 is invalid")
        path = self.resolve(str(raw), packet_bound=packet_bound)
        if not path.is_file():
            if allow_missing:
                return None, {"path": str(raw), "bytes": size, "sha256": sha}
            fail(f"missing {label}: {raw}")
        require(not path.is_symlink(), f"{label} is a symlink: {raw}")
        require(path.stat().st_size == size, f"{label}.bytes changed")
        require(digest(path) == sha, f"{label}.sha256 changed")
        return path, {"path": str(raw), "bytes": size, "sha256": sha}


def source_manifest(value: Any, label: str) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label} is not an object")
    revision = value.get("revision")
    require(isinstance(revision, str) and revision and all(c in HEX for c in revision),
            f"{label}.revision is invalid")
    files = value.get("files")
    require(isinstance(files, dict) and files, f"{label}.files is missing")
    for name, sha in files.items():
        require(isinstance(name, str) and name and is_sha(sha),
                f"{label} contains an invalid file digest")
    return {"revision": revision, "files": dict(files)}


def cleanup_is_verified(cleanup: Any) -> bool:
    return isinstance(cleanup, dict) and any(
        cleanup.get(key) is True
        for key in ("verified", "executables_verified_before_removal",
                    "binary_removal_verified")
    )


def cleanup_contains(value: Any, receipt: dict[str, Any]) -> bool:
    """Find an exact raw path/size/SHA tuple in a cleanup witness."""

    wanted = (receipt.get("path"), receipt.get("bytes", receipt.get("size")),
              receipt.get("sha256", receipt.get("digest")))
    if not (isinstance(wanted[0], str) and isinstance(wanted[1], int)
            and not isinstance(wanted[1], bool) and is_sha(wanted[2])):
        return False
    if isinstance(value, dict):
        actual = (value.get("path"), value.get("bytes", value.get("size")),
                  value.get("sha256", value.get("digest")))
        if actual == wanted:
            return True
        return any(cleanup_contains(child, receipt) for child in value.values())
    if isinstance(value, list):
        return any(cleanup_contains(child, receipt) for child in value)
    return False


def verify_binary(ctx: Context, receipt: Any, label: str, cleanup: Any,
                  cleanup_verified: bool) -> dict[str, Any]:
    require(isinstance(receipt, dict), f"{label} is missing")
    path, normalized_receipt = ctx.artifact(receipt, label, packet_bound=False,
                                             allow_missing=True)
    if path is not None:
        # Custody has already been checked against the retained file.  Keep
        # only the immutable identity in the returned analysis: retained
        # absolute paths differ after worktree relocation.
        return {"bytes": normalized_receipt["bytes"],
                "sha256": normalized_receipt["sha256"]}
    require(cleanup_verified, f"{label} is missing without cleanup verification")
    require(cleanup_contains(cleanup, receipt), f"{label} lacks an exact cleanup witness")
    # A cleanup witness proves the same receipt identity.  Do not expose the
    # witness/retention state in the analysis because it changes when the
    # packet is sealed and executables are removed.
    return {"bytes": normalized_receipt["bytes"],
            "sha256": normalized_receipt["sha256"]}


def load_cleanup(ctx: Context) -> tuple[Any, bool]:
    path = ctx.packet / "cleanup.json"
    if not path.is_file():
        return None, False
    value = read_json(path, "cleanup.json")
    require(isinstance(value, dict), "cleanup.json is malformed")
    return value, cleanup_is_verified(value)


def load_plan(ctx: Context) -> tuple[dict[str, Any], dict[str, Any]]:
    plan = read_json(ctx.packet / "plan.json", "plan.json")
    observer = read_json(ctx.packet / "observer-plan.json", "observer-plan.json")
    require(isinstance(plan, dict) and isinstance(observer, dict), "observer plans are malformed")
    require(plan.get("schema") == "litchi.performance.0781.v1", "main plan schema changed")
    require(observer.get("schema") == "litchi.performance.0781.observers.v1",
            "observer plan schema changed")
    require(plan.get("cpu") == 12 and observer.get("cpu") == 12, "capture CPU changed")
    perf = observer.get("perf_stat")
    heap = observer.get("heaptrack")
    require(isinstance(perf, dict) and isinstance(heap, dict), "observer lane configuration missing")
    require(perf.get("blocks") == 3 and perf.get("samples") == 30
            and perf.get("warmup") == 3, "perf observer configuration changed")
    require(perf.get("events") == "instructions:u,cycles:u", "perf event set changed")
    require(perf.get("cases") == [
        {"mode": "write", "shape": "payload"},
        {"mode": "lifecycle", "shape": "payload"},
    ], "perf observer cases changed")
    require(perf.get("scope") ==
            "Whole process including setup and verification; not operation-region counters.",
            "perf observer scope changed")
    require(heap == {
        "mode": "write", "samples": 3,
        "scope": "Whole-process attribution of plain-text conversion allocations; no timing/RSS/causal claim.",
        "shape": "payload", "warmup": 0,
    }, "heaptrack observer configuration changed")
    orders = plan.get("native", {}).get("orders")
    require(orders == [
        ["before", "after"], ["after", "before"], ["before", "after"],
        ["after", "before"], ["after", "before"], ["before", "after"],
    ], "native order used by observer changed")
    changed = ctx.packet / "candidate/changed-files.json"
    if changed.is_file():
        changed_value = read_json(changed, "candidate changed-files.json")
        if isinstance(changed_value, list):
            plan["source_allowlist"] = changed_value
        elif isinstance(changed_value, dict):
            files = changed_value.get("files", changed_value.get("changed_files"))
            if isinstance(files, dict):
                files = list(files)
            require(isinstance(files, list), "candidate source allowlist is malformed")
            plan["source_allowlist"] = files
    plan.setdefault("source_allowlist", [SOURCE_FILE])
    require(isinstance(plan["source_allowlist"], list)
            and plan["source_allowlist"]
            and all(isinstance(name, str) and name for name in plan["source_allowlist"]),
            "candidate source allowlist is missing")
    require(len(set(plan["source_allowlist"])) == len(plan["source_allowlist"]),
            "candidate source allowlist contains duplicates")
    return plan, observer


def probe_contract(plan: dict[str, Any]) -> tuple[str, str, str]:
    """Return the frozen report identity shared with the probe analyzer."""

    value = plan.get("probe", {})
    require(isinstance(value, dict), "probe contract is malformed")
    contract = (
        value.get("schema", REPORT_SCHEMA),
        value.get("tool", REPORT_TOOL),
        value.get("marker", MARKER),
    )
    require(all(isinstance(item, str) and item for item in contract),
            "probe contract identity is missing")
    return contract


def load_build(ctx: Context, leg: str, cleanup: Any, cleanup_verified: bool) -> dict[str, Any]:
    manifest_path = ctx.packet / f"build-{leg}/build.json"
    build = read_json(manifest_path, f"{leg} build.json")
    require(isinstance(build, dict), f"{leg} build manifest is malformed")
    source_path, _ = ctx.artifact(build.get("source"), f"{leg} build source")
    assert source_path is not None
    source = source_manifest(read_json(source_path, f"{leg} source manifest"),
                             f"{leg} source manifest")
    binaries = build.get("binaries")
    require(isinstance(binaries, dict) and set(binaries) == {"native", "allocation"},
            f"{leg} binary identities are incomplete")
    binary_state = {
        name: verify_binary(ctx, receipt, f"{leg} {name} binary", cleanup, cleanup_verified)
        for name, receipt in binaries.items()
    }
    rows = build.get("rows")
    require(isinstance(rows, list) and len(rows) == 2, f"{leg} build rows are incomplete")
    manifest = ctx.packet / "probe-src/Cargo.toml"
    expected_base = ["cargo", "build", "--offline", "--release", "--manifest-path",
                     str(manifest)]
    commands: dict[str, list[str]] = {}
    for row in rows:
        require(isinstance(row, dict) and row.get("exit_code") == 0,
                f"{leg} build command failed")
        log_path, _ = ctx.artifact(row.get("log"), f"{leg} build log")
        require(log_path is not None, f"{leg} build log is missing")
        command = row.get("command")
        require(isinstance(command, list) and all(isinstance(item, str) for item in command),
                f"{leg} build command is malformed")
        normalized = [ctx.normalized(item) for item in command]
        if "allocator-metrics" in normalized:
            kind = "allocation"
            expected = expected_base + ["--features", "allocator-metrics"]
        else:
            kind = "native"
            expected = expected_base
        require(normalized in (expected, expected + ["--locked"]),
                f"{leg} {kind} build command changed")
        require(kind not in commands, f"{leg} has duplicate {kind} build row")
        commands[kind] = normalized
    require(set(commands) == {"native", "allocation"}, f"{leg} build kinds are incomplete")
    environment = build.get("environment")
    require(isinstance(environment, dict) and environment.get("CARGO_BUILD_JOBS") == "2"
            and environment.get("CARGO_INCREMENTAL") == "0",
            f"{leg} build environment changed")
    lock = build.get("lock")
    lock_path, lock_receipt = ctx.artifact(lock, f"{leg} probe lock")
    require(lock_path is not None, f"{leg} probe lock is missing")
    require(lock_path == (ctx.packet / "probe-src/Cargo.lock").resolve(),
            f"{leg} lock receipt did not relocate to packet")
    probe = build.get("probe")
    require(isinstance(probe, dict) and probe, f"{leg} probe inventory is missing")
    for name, sha in probe.items():
        require(isinstance(name, str) and is_sha(sha), f"{leg} probe digest is invalid: {name}")
        path = ctx.packet / name
        require(path.is_file() and not path.is_symlink(), f"missing {leg} probe file: {name}")
        require(digest(path) == sha, f"{leg} probe file changed: {name}")
    return {"manifest": build, "source": source, "source_path": source_path,
            "binaries": binaries, "binary_state": binary_state,
            "lock": lock_receipt, "probe": probe}


def expected_jobs(plan: dict[str, Any], observer: dict[str, Any]) -> list[dict[str, Any]]:
    perf = observer["perf_stat"]
    jobs: list[dict[str, Any]] = []
    for block in range(perf["blocks"]):
        for case in perf["cases"]:
            for leg in plan["native"]["orders"][block]:
                jobs.append({"lane": "perf", "block": block, **case, "leg": leg,
                             "samples": perf["samples"], "warmup": perf["warmup"]})
    return jobs


def expected_perf_command(ctx: Context, plan: dict[str, Any], job: dict[str, Any],
                          binary: dict[str, Any], report: dict[str, Any],
                          stats: dict[str, Any]) -> list[str]:
    return ["taskset", "-c", str(plan["cpu"]), "perf", "stat", "-x", ";", "-e",
            "instructions:u,cycles:u", "-o", ctx.normalized(str(stats["path"])), "--",
            ctx.normalized(str(binary["path"])), "--mode", job["mode"], "--shape",
            job["shape"], "--samples", str(job["samples"]), "--warmup",
            str(job["warmup"]), "--output", ctx.normalized(str(report["path"]))]


def expected_heap_command(ctx: Context, plan: dict[str, Any], job: dict[str, Any],
                          binary: dict[str, Any], report: dict[str, Any], lane: str) -> list[str]:
    return ["taskset", "-c", str(plan["cpu"]), "heaptrack", "-o",
            str(ctx.packet / lane / "trace"), ctx.normalized(str(binary["path"])),
            "--mode", job["mode"], "--shape", job["shape"], "--samples",
            str(job["samples"]), "--warmup", str(job["warmup"]), "--output",
            ctx.normalized(str(report["path"]))]


def identity(value: Any, label: str) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label} identity is missing")
    size = value.get("bytes", value.get("size"))
    nonnegative_int(size, f"{label}.bytes")
    sha = value.get("sha256", value.get("digest"))
    require(is_sha(sha), f"{label}.sha256 is invalid")
    return {"bytes": size, "sha256": sha}


def validate_report(report: Any, job: dict[str, Any], label: str,
                    expected_binary: str, contract: tuple[str, str, str]) -> dict[str, Any]:
    report_schema, report_tool, _marker = contract
    require(isinstance(report, dict), f"{label} report is not an object")
    require(report.get("schema") == report_schema, f"{label} report schema changed")
    require(report.get("tool") == report_tool, f"{label} report tool changed")
    require(report.get("mode") == job["mode"] and report.get("shape") == job["shape"],
            f"{label} report mode/shape changed")
    expected_dimensions = SHAPES[job["shape"]]
    require((report.get("slides"), report.get("boxes_per_slide")) == expected_dimensions,
            f"{label} corpus dimensions changed")
    require(report.get("text_boxes") == expected_dimensions[0] * expected_dimensions[1],
            f"{label} text-box count changed")
    source = identity(report.get("source"), f"{label} source")
    require(source["bytes"] > 0, f"{label} source is empty")
    require(isinstance(report.get("timing_scope"), str) and report["timing_scope"],
            f"{label} timing scope is missing")
    expected_text_bytes = report.get("expected_semantic_text_bytes")
    nonnegative_int(expected_text_bytes, f"{label} expected semantic text bytes")
    expected_raw_bytes = report.get("expected_raw_text_bytes")
    nonnegative_int(expected_raw_bytes, f"{label} expected raw text bytes")
    require(report.get("samples_requested") == job["samples"]
            and report.get("warmup") == job["warmup"],
            f"{label} sample configuration changed")
    allocator = report.get("allocator")
    require(isinstance(allocator, dict), f"{label} allocator identity is missing")
    require(allocator.get("binary") == expected_binary,
            f"{label} native binary identity changed")
    samples = report.get("samples")
    require(isinstance(samples, list) and len(samples) == job["samples"],
            f"{label} report sample count changed")
    outputs: list[dict[str, Any]] = []
    for index, sample in enumerate(samples):
        require(isinstance(sample, dict) and sample.get("index") == index,
                f"{label} sample {index} identity changed")
        elapsed = sample.get("elapsed_ns")
        nonnegative_int(elapsed, f"{label} sample {index} elapsed_ns")
        metrics = sample.get("metrics")
        require(isinstance(metrics, dict)
                and metrics.get("elapsed_ns") == elapsed
                and metrics.get("slides") == report.get("slides")
                and metrics.get("boxes_per_slide") == report.get("boxes_per_slide"),
                f"{label} sample {index} elapsed metric changed")
        require(sample.get("source_sha256") == source["sha256"],
                f"{label} sample {index} source parity failed")
        verification = sample.get("verification")
        require(isinstance(verification, dict)
                and verification.get("reopened") is True
                and verification.get("semantic_check") is True
                and verification.get("exact_text_match") is True
                and verification.get("slide_count_match") is True
                and verification.get("slide_count") == report.get("slides")
                and verification.get("expected_slide_count") == report.get("slides")
                and verification.get("expected_semantic_text_bytes") == expected_text_bytes
                and is_sha(verification.get("expected_semantic_text_sha256"))
                and is_sha(verification.get("semantic_text_sha256"))
                and verification.get("expected_semantic_text_sha256")
                == verification.get("semantic_text_sha256")
                and verification.get("semantic_text_bytes") == expected_text_bytes
                and isinstance(verification.get("raw_text_box_count"), int)
                and not isinstance(verification.get("raw_text_box_count"), bool)
                and verification.get("raw_text_box_count") >= 0
                and isinstance(verification.get("expected_raw_text_box_count"), int)
                and not isinstance(verification.get("expected_raw_text_box_count"), bool)
                and verification.get("expected_raw_text_box_count") >= 0
                and isinstance(verification.get("raw_text_atom_count"), int)
                and not isinstance(verification.get("raw_text_atom_count"), bool)
                and verification.get("raw_text_atom_count") >= 0
                and verification.get("raw_text_check") is True
                and verification.get("raw_text_match") is True
                and verification.get("raw_text_boxes_match") is True
                and verification.get("raw_text_box_count")
                == verification.get("expected_raw_text_box_count")
                and verification.get("expected_raw_text_bytes") == expected_raw_bytes
                and verification.get("raw_text_bytes") == expected_raw_bytes
                and is_sha(verification.get("expected_raw_text_sha256"))
                and is_sha(verification.get("raw_text_sha256"))
                and verification.get("raw_text_sha256")
                == verification.get("expected_raw_text_sha256"),
                f"{label} sample {index} semantic verification failed")
        output = identity(sample.get("output"), f"{label} sample {index} output")
        require(sample.get("allocation") is None,
                f"{label} native observer report contains allocation metrics")
        outputs.append(output)
    if outputs:
        require(all(item == outputs[0] for item in outputs),
                f"{label} output identity is not stable across samples")
    return {"source": source, "output": outputs[0] if outputs else None,
            "samples": len(samples), "expected_semantic_text_bytes": expected_text_bytes,
            "expected_raw_text_bytes": expected_raw_bytes}


def parse_number(value: str) -> float | int | None:
    token = value.strip().replace(",", "")
    if not token or token.startswith("<"):
        return None
    token = token.rstrip("%")
    try:
        number = float(token)
    except ValueError:
        return None
    if not math.isfinite(number):
        return None
    return int(number) if number.is_integer() else number


def parse_perf_stats(path: Path, label: str) -> dict[str, dict[str, float | int]]:
    events: dict[str, dict[str, float | int]] = {}
    try:
        lines = path.read_text().splitlines()
    except OSError as error:
        fail(f"cannot read {label} perf stats: {error}")
    for line in lines:
        if not line.strip() or line.lstrip().startswith("#"):
            continue
        fields = [field.strip() for field in line.split(";")]
        event_index = next((
            index for index, field in enumerate(fields)
            if field.lower().startswith(("instructions", "cycles"))
        ), None)
        if event_index is None:
            continue
        event_field = fields[event_index].lower()
        event = "instructions" if event_field.startswith("instructions") else "cycles"
        value = parse_number(fields[0]) if fields else None
        require(value is not None, f"{label} {event} counter is unavailable")
        running: float | None = None
        for field in fields[event_index + 1:]:
            candidate = parse_number(field)
            if candidate is not None and 0 <= float(candidate) <= 100:
                running = float(candidate)
                break
        require(running is not None, f"{label} {event} running percentage is missing")
        require(event not in events, f"{label} duplicates {event} counter")
        events[event] = {"value": value, "running_percent": running}
    require(set(events) == {"instructions", "cycles"},
            f"{label} does not contain exactly instructions and cycles")
    return events


def validate_perf(ctx: Context, plan: dict[str, Any], observer: dict[str, Any],
                  builds: dict[str, dict[str, Any]], cleanup: Any,
                  cleanup_verified: bool, contract: tuple[str, str, str]) -> dict[str, Any]:
    lane = "perf"
    directory = ctx.packet / lane
    require(directory.is_dir(), "missing perf observer directory")
    source = source_manifest(read_json(directory / "source.json", "perf source manifest"),
                             "perf source manifest")
    rows_path, _ = ctx.artifact(
        read_json(directory / "complete.json", "perf complete.json").get("receipts"),
        "perf receipts")
    assert rows_path is not None
    rows = read_json(rows_path, "perf receipts")
    jobs = expected_jobs(plan, observer)
    require(isinstance(rows, list) and len(rows) == len(jobs) == 12,
            "perf observer cardinality is not 12 processes")
    complete = read_json(directory / "complete.json", "perf complete.json")
    require(complete.get("children") == 12, "perf complete child count changed")
    entries: list[dict[str, Any]] = []
    seen: set[tuple[Any, ...]] = set()
    for index, (row, job) in enumerate(zip(rows, jobs)):
        label = f"perf child {index}"
        require(isinstance(row, dict), f"{label} is not an object")
        require(all(row.get(key) == job[key] for key in ("leg", "block", "mode", "shape")),
                f"{label} identity changed")
        key = (job["block"], job["shape"], job["mode"], job["leg"])
        require(key not in seen, f"duplicate perf identity: {key}")
        seen.add(key)
        require(row.get("exit_code") == 0, f"{label} failed")
        binary = row.get("binary")
        expected_binary = builds[job["leg"]]["manifest"]["binaries"]["native"]
        require(isinstance(binary, dict)
                and binary.get("path") == expected_binary.get("path")
                and binary.get("bytes") == expected_binary.get("bytes")
                and binary.get("sha256") == expected_binary.get("sha256"),
                f"{label} native binary receipt differs from build identity")
        binary_state = verify_binary(ctx, binary, f"{label} native binary", cleanup,
                                     cleanup_verified)
        log_path, log_receipt = ctx.artifact(row.get("log"), f"{label} log")
        report_path, report_receipt = ctx.artifact(row.get("report"), f"{label} report")
        stats_path, stats_receipt = ctx.artifact(row.get("stats"), f"{label} perf stats")
        assert log_path is not None and report_path is not None and stats_path is not None
        command = row.get("command")
        require(isinstance(command, list) and all(isinstance(item, str) for item in command),
                f"{label} command is malformed")
        require([ctx.normalized(item) for item in command]
                == expected_perf_command(ctx, plan, job, expected_binary,
                                         report_receipt, stats_receipt),
                f"{label} capture command changed")
        report = validate_report(read_json(report_path, f"{label} report"), job, label,
                                 Path(str(expected_binary.get("path"))).name,
                                 contract)
        counters = parse_perf_stats(stats_path, label)
        entries.append({"identity": job, "binary": binary_state,
                        "report": {"path": ctx.relative(report_path),
                                    "bytes": report_receipt["bytes"],
                                    "sha256": report_receipt["sha256"]},
                        "source": report["source"], "output": report["output"],
                        "counters": counters,
                        "artifacts": {
                            "log": {"path": ctx.relative(log_path), "sha256": log_receipt["sha256"]},
                            "stats": {"path": ctx.relative(stats_path), "sha256": stats_receipt["sha256"]},
                            "report": {"path": ctx.relative(report_path), "sha256": report_receipt["sha256"]},
                        }})
    require(source["files"] == builds["after"]["source"]["files"],
            "perf source receipt differs from after build source files")
    parity = source_output_parity(entries)
    pairs = perf_pairs(entries)
    return {"processes": len(entries), "blocks": observer["perf_stat"]["blocks"],
            "cases": observer["perf_stat"]["cases"], "scope": observer["perf_stat"]["scope"],
            "whole_process": True, "source_output_parity": parity,
            "pairs": pairs, "receipts": entries}


def source_output_parity(entries: list[dict[str, Any]]) -> dict[str, Any]:
    groups: dict[tuple[str, str], list[dict[str, Any]]] = {}
    for entry in entries:
        identity = entry["identity"]
        groups.setdefault((identity["shape"], identity["mode"]), []).append(entry)
    result: list[dict[str, Any]] = []
    for (shape, mode), group in sorted(groups.items()):
        sources = {(item["source"]["bytes"], item["source"]["sha256"]) for item in group}
        outputs = {(item["output"]["bytes"], item["output"]["sha256"])
                   for item in group if item["output"] is not None}
        require(len(sources) == 1, f"{shape}/{mode} source identity differs across paired processes")
        require(len(outputs) == 1, f"{shape}/{mode} output identity differs across paired processes")
        result.append({"shape": shape, "mode": mode, "processes": len(group),
                       "source": {"bytes": next(iter(sources))[0],
                                  "sha256": next(iter(sources))[1]},
                       "output": (None if not outputs else
                                  {"bytes": next(iter(outputs))[0],
                                   "sha256": next(iter(outputs))[1]})})
    return {"groups": result, "checked": True}


def percent_change(before: float | int, after: float | int) -> float | None:
    if before == 0:
        return None
    return (float(after) - float(before)) * 100.0 / abs(float(before))


def fraction(part: int, whole: int) -> float | None:
    """Return a descriptive allocation share, or null for a zero total."""

    if whole == 0:
        return None
    return float(part) / float(whole)


def perf_pairs(entries: list[dict[str, Any]]) -> list[dict[str, Any]]:
    by_key = {(item["identity"]["block"], item["identity"]["shape"],
               item["identity"]["mode"], item["identity"]["leg"]): item
              for item in entries}
    result: list[dict[str, Any]] = []
    for block in range(3):
        for shape, mode in (("payload", "write"), ("payload", "lifecycle")):
            before = by_key[(block, shape, mode, "before")]
            after = by_key[(block, shape, mode, "after")]
            counters: dict[str, Any] = {}
            for event in ("instructions", "cycles"):
                left = before["counters"][event]
                right = after["counters"][event]
                counters[event] = {
                    "before": left["value"], "after": right["value"],
                    "before_running_percent": left["running_percent"],
                    "after_running_percent": right["running_percent"],
                    "change_percent": percent_change(left["value"], right["value"]),
                }
            result.append({"block": block, "shape": shape, "mode": mode,
                           "scope": "whole process including setup and verification",
                           "counters": counters})
    return result


def trace_lines(path: Path) -> Iterable[bytes]:
    """Stream an already retained Heaptrack v3 trace through zstd."""

    require(path.suffix == ".zst", f"heaptrack trace is not a .zst stream: {path}")
    try:
        process = subprocess.Popen(
            ["zstd", "--decompress", "--stdout", "--", str(path)],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
    except OSError as error:
        fail(f"cannot start offline zstd decode for {path}: {error}")
    assert process.stdout is not None and process.stderr is not None
    total = 0
    try:
        for raw in process.stdout:
            total += len(raw)
            require(total <= MAX_TRACE_BYTES,
                    f"decompressed trace exceeds {MAX_TRACE_BYTES} bytes: {path}")
            yield raw
    finally:
        process.stdout.close()
    error = process.stderr.read().decode("utf-8", "replace")
    process.stderr.close()
    require(process.wait() == 0, f"offline zstd decode failed for {path}: {error.strip()}")


def trace_hex(value: str, label: str, record: int) -> int:
    try:
        return int(value, 16)
    except ValueError as error:
        fail(f"invalid hexadecimal {label} at trace record {record}: {value!r}")
        raise AssertionError from error


def parse_heap_trace(path: Path) -> dict[str, Any]:
    """Parse allocation events and ancestry from an interpreted Heaptrack v3 trace."""

    strings: dict[int, str] = {}
    instructions: dict[int, tuple[int, ...]] = {}
    traces: dict[int, tuple[int, int]] = {}
    allocation_info: list[tuple[int, int]] = []
    plus: Counter[int] = Counter()
    minus: Counter[int] = Counter()
    records = 0
    for raw in trace_lines(path):
        records += 1
        require(records <= MAX_TRACE_RECORDS,
                f"trace has more than {MAX_TRACE_RECORDS} records: {path}")
        try:
            line = raw.rstrip(b"\n").decode("utf-8")
        except UnicodeDecodeError as error:
            fail(f"Heaptrack trace is not UTF-8 at record {records}: {error}")
        if not line:
            continue
        fields = line.split()
        mode = fields[0]
        if mode == "s":
            require(len(fields) >= 3, f"malformed Heaptrack string record {records}")
            declared = trace_hex(fields[1], "string length", records)
            text = line.split(None, 2)[2]
            require(len(text.encode("utf-8")) == declared,
                    f"Heaptrack string length mismatch at record {records}")
            strings[len(strings) + 1] = text
        elif mode == "i":
            require(len(fields) >= 3, f"malformed Heaptrack instruction record {records}")
            tokens = fields[1:]
            function_ids: list[int] = []
            if len(tokens) >= 3:
                function_id = trace_hex(tokens[2], "function index", records)
                if function_id:
                    function_ids.append(function_id)
                remaining = tokens[3:]
                if remaining:
                    require(len(remaining) >= 2 and (len(remaining) - 2) % 3 == 0,
                            f"malformed Heaptrack inline frames at record {records}")
                    for index in range(2, len(remaining), 3):
                        function_id = trace_hex(remaining[index], "inline function index", records)
                        if function_id:
                            function_ids.append(function_id)
            instructions[len(instructions) + 1] = tuple(function_ids)
        elif mode == "t":
            require(len(fields) == 3, f"malformed Heaptrack trace record {records}")
            trace_id = len(traces) + 1
            instruction = trace_hex(fields[1], "instruction index", records)
            parent = trace_hex(fields[2], "trace parent", records)
            require(not instruction or instruction in instructions,
                    f"unknown instruction at Heaptrack record {records}")
            require(not parent or parent < trace_id,
                    f"forward Heaptrack trace parent at record {records}")
            traces[trace_id] = (instruction, parent)
        elif mode == "a":
            require(len(fields) == 3, f"malformed Heaptrack allocation-info record {records}")
            size = trace_hex(fields[1], "allocation size", records)
            trace_id = trace_hex(fields[2], "allocation trace", records)
            require(not trace_id or trace_id in traces,
                    f"unknown allocation trace at Heaptrack record {records}")
            allocation_info.append((size, trace_id))
        elif mode in {"+", "-"}:
            require(len(fields) == 2, f"malformed Heaptrack allocation event {records}")
            index = trace_hex(fields[1], "allocation-info index", records)
            require(index < len(allocation_info),
                    f"unknown allocation-info index at Heaptrack record {records}")
            (plus if mode == "+" else minus)[index] += 1
        # Heaptrack metadata records (v, X, I, c, ...) are intentionally
        # retained by the trace but do not participate in this allocation audit.

    names_cache: dict[int, tuple[str, ...]] = {}

    def names_for_trace(trace_id: int) -> tuple[str, ...]:
        if trace_id in names_cache:
            return names_cache[trace_id]
        names: list[str] = []
        seen: set[int] = set()
        current = trace_id
        while current:
            require(current not in seen, "Heaptrack trace ancestry cycle detected")
            seen.add(current)
            instruction, parent = traces.get(current, (0, 0))
            # Heaptrack's instruction record stores the primary function and
            # its inline function triples together.  Include every function
            # id before walking the parent trace, so an inlined conversion is
            # attributed by its complete ancestry rather than only by the
            # out-of-line caller.
            for string_id in instructions.get(instruction, ()):
                require(string_id in strings, f"unknown Heaptrack string {string_id}")
                names.append(strings[string_id])
            current = parent
        names_cache[trace_id] = tuple(names)
        return names_cache[trace_id]

    totals = {
        "allocation_events": 0,
        "requested_bytes": 0,
        "deallocation_events": 0,
        "deallocated_bytes": 0,
    }
    conversion = {
        "symbol": CONVERSION_SYMBOL, "allocation_info_records": 0,
        "allocation_events": 0, "requested_bytes": 0,
        "deallocation_events": 0, "deallocated_bytes": 0,
    }
    target_rows: list[dict[str, Any]] = []
    for index, (size, trace_id) in enumerate(allocation_info):
        occurrences = plus[index]
        deallocations = minus[index]
        totals["allocation_events"] += occurrences
        totals["requested_bytes"] += size * occurrences
        totals["deallocation_events"] += deallocations
        totals["deallocated_bytes"] += size * deallocations
        if not occurrences:
            continue
        names = names_for_trace(trace_id)
        matches = {"conversion": any(CONVERSION_SYMBOL in name for name in names)}
        if not matches["conversion"]:
            continue
        row = {"allocation_info_index": index, "size": size,
               "allocation_events": occurrences, "deallocation_events": deallocations,
               "requested_bytes": size * occurrences,
               "deallocated_bytes": size * deallocations,
               "ancestry": list(names),
               "matches": matches, "trace_id": trace_id}
        target_rows.append(row)
        for target, matched in matches.items():
            if not matched:
                continue
            summary = conversion
            summary["allocation_info_records"] += 1
            summary["allocation_events"] += occurrences
            summary["requested_bytes"] += size * occurrences
            summary["deallocation_events"] += deallocations
            summary["deallocated_bytes"] += size * deallocations
    return {
        "format": "interpreted_heaptrack_v3",
        "trace_records_read": records,
        "trace_records": len(traces),
        "instruction_records": len(instructions),
        "allocation_info_records": len(allocation_info),
        "totals": totals,
        "targets": {"conversion": conversion},
        "target_rows": target_rows,
    }


def parse_histogram(path: Path) -> dict[str, int]:
    rows = 0
    allocation_events = 0
    requested_bytes = 0
    try:
        lines = path.read_text().splitlines()
    except OSError as error:
        fail(f"cannot read Heaptrack histogram {path}: {error}")
    for line in lines:
        if not line:
            continue
        rows += 1
        require(rows <= MAX_HISTOGRAM_ROWS, f"Heaptrack histogram exceeds {MAX_HISTOGRAM_ROWS} rows")
        fields = line.split("\t")
        require(len(fields) == 2 and all(field.isdecimal() for field in fields),
                f"malformed Heaptrack histogram row {rows}: {line!r}")
        size, occurrences = (int(field) for field in fields)
        allocation_events += occurrences
        requested_bytes += size * occurrences
    require(rows > 0, f"empty Heaptrack histogram: {path}")
    return {"rows": rows, "allocation_events": allocation_events,
            "requested_bytes": requested_bytes}


def parse_print_summary(path: Path) -> dict[str, int]:
    try:
        lines = path.read_text(errors="replace").splitlines()
    except OSError as error:
        fail(f"cannot read Heaptrack print log {path}: {error}")
    result: dict[str, int] = {}
    for line in lines:
        if line.startswith("calls to allocation functions:"):
            result["allocation_events"] = int(line.split(":", 1)[1].split("(", 1)[0].strip())
        elif line.startswith("temporary memory allocations:"):
            result["temporary_allocations"] = int(line.split(":", 1)[1].split("(", 1)[0].strip())
    require("allocation_events" in result, f"Heaptrack print log has no allocation total: {path}")
    return result


def decoded_receipts(ctx: Context, lane: str,
                     trace_artifacts: list[dict[str, Any]]) -> tuple[dict[str, Any] | None,
                                                                    list[tuple[str, Path]]]:
    metadata: dict[str, Any] | None = None
    for name in ("decode.json", "decoded.json", "heaptrack-decode.json"):
        path = ctx.packet / lane / name
        if path.is_file():
            value = read_json(path, f"{lane} decode metadata")
            require(isinstance(value, dict), f"{lane} decode metadata is malformed")
            metadata = value
            break
    if metadata is not None:
        require(metadata.get("exit_code") == 0, f"{lane} decode command failed")
        command = metadata.get("command")
        require(isinstance(command, list) and command
                and all(isinstance(item, str) for item in command),
                f"{lane} decode command is missing")
        require("trace" in metadata and "histogram" in metadata,
                f"{lane} decode metadata is missing trace or histogram")
        trace_path, trace_receipt = ctx.artifact(metadata["trace"], f"{lane} decoded trace")
        histogram_path, histogram_receipt = ctx.artifact(
            metadata["histogram"], f"{lane} decoded histogram")
        assert trace_path is not None and histogram_path is not None
        require(any(trace_receipt["sha256"] == item["sha256"]
                    and trace_receipt["bytes"] == item["bytes"]
                    for item in trace_artifacts),
                f"{lane} decode trace differs from captured trace")
        expected_command = ["heaptrack_print", "-f", ctx.normalized(str(trace_receipt["path"])),
                            "-H", ctx.normalized(str(histogram_receipt["path"])), "-n", "15",
                            "-s", "3"]
        require([ctx.normalized(item) for item in command] == expected_command,
                f"{lane} decode command changed")
        found: list[tuple[str, Path]] = []
        for key in ("print", "print_output", "log", "histogram", "histogram_output"):
            if key not in metadata:
                continue
            path, _ = ctx.artifact(metadata[key], f"{lane} {key}")
            assert path is not None
            found.append((key, path))
        require(found, f"{lane} decode metadata has no retained outputs")
        return metadata, found
    # A root-only decode may be present before its metadata receipt is sealed.
    # Parse it for a diagnostic, but label the result unbound so it cannot be
    # mistaken for a custody-complete heaptrack result.
    found = []
    for path in sorted((ctx.packet / lane).iterdir()):
        if not path.is_file() or path.name in {"source.json", "receipts.json",
                                                "complete.json"}:
            continue
        if path.suffix.lower() not in {".txt", ".log", ".json"}:
            continue
        if any(token in path.name.lower() for token in ("print", "hist", "decode")):
            found.append((path.name, path))
    return None, found


def _metric_from_lines(lines: list[str], index: int) -> dict[str, Any]:
    window = " ".join(lines[max(0, index - 120):index + 1])
    metric: dict[str, Any] = {}
    matches = list(re.finditer(
        r"([\d,]+)\s+peak memory consumed over\s+([\d,]+)\s+calls", window,
        re.IGNORECASE))
    match = matches[-1] if matches else None
    if match:
        metric["bytes_or_units"] = int(match.group(1).replace(",", ""))
        metric["calls"] = int(match.group(2).replace(",", ""))
    matches = list(re.finditer(
        r"([\d,]+)\s+calls(?:\s+with\s+([^\s]+)\s+peak)?", window,
        re.IGNORECASE))
    match = matches[-1] if matches else None
    if match:
        metric["calls"] = int(match.group(1).replace(",", ""))
        if match.group(2):
            metric["peak_text"] = match.group(2)
    return metric


def text_attribution(path: Path) -> list[dict[str, Any]]:
    try:
        lines = path.read_text(errors="replace").splitlines()
    except OSError as error:
        fail(f"cannot read heaptrack decode {path}: {error}")
    hits: list[dict[str, Any]] = []
    for index, line in enumerate(lines):
        if any(symbol.lower() in line.lower() for symbol in TARGET_SYMBOLS):
            hits.append({"line": index + 1, "symbol": line.strip(),
                         "metric": _metric_from_lines(lines, index)})
    return hits


def json_attribution(value: Any, path: str = "$") -> list[dict[str, Any]]:
    hits: list[dict[str, Any]] = []
    if isinstance(value, dict):
        searchable = " ".join(str(value[key]) for key in value
                               if key.lower() in {"symbol", "function", "name", "frame",
                                                  "location", "stack", "callstack"})
        if any(symbol.lower() in searchable.lower() for symbol in TARGET_SYMBOLS):
            numbers = {key: item for key, item in value.items()
                       if key.lower() in {"calls", "count", "bytes", "allocated_bytes",
                                          "peak_bytes", "memory", "consumed"}
                       and isinstance(item, (int, float)) and not isinstance(item, bool)}
            hits.append({"path": path, "symbol": searchable, "metric": numbers})
        for key, item in value.items():
            hits.extend(json_attribution(item, f"{path}.{key}"))
    elif isinstance(value, list):
        for index, item in enumerate(value):
            hits.extend(json_attribution(item, f"{path}[{index}]"))
    return hits


def heaptrack_diagnostic(ctx: Context, lanes: dict[str, dict[str, Any]]) -> dict[str, Any]:
    decoded: dict[str, Any] = {}
    for lane, data in lanes.items():
        metadata, files = decoded_receipts(ctx, lane, data["traces"])
        hits: list[dict[str, Any]] = []
        artifacts: list[dict[str, Any]] = []
        for name, path in files:
            sha = digest(path)
            artifacts.append({"name": name, "path": ctx.relative(path), "sha256": sha,
                              "bytes": path.stat().st_size})
            if path.suffix.lower() == ".json":
                hits.extend(json_attribution(read_json(path, f"{lane} decoded JSON"), name))
            else:
                hits.extend(text_attribution(path))
        exact = None
        if metadata is not None:
            require("trace" in metadata and "histogram" in metadata and "log" in metadata,
                    f"{lane} decode metadata is missing exact attribution artifacts")
            trace_path, _ = ctx.artifact(metadata["trace"], f"{lane} exact trace")
            histogram_path, _ = ctx.artifact(metadata["histogram"], f"{lane} exact histogram")
            print_path, _ = ctx.artifact(metadata["log"], f"{lane} exact print log")
            assert trace_path is not None and histogram_path is not None and print_path is not None
            trace = parse_heap_trace(trace_path)
            histogram = parse_histogram(histogram_path)
            print_summary = parse_print_summary(print_path)
            require(trace["totals"]["allocation_events"] == histogram["allocation_events"],
                    f"{lane} trace allocation count differs from histogram")
            require(trace["totals"]["requested_bytes"] == histogram["requested_bytes"],
                    f"{lane} trace requested bytes differ from histogram")
            require(print_summary["allocation_events"] == histogram["allocation_events"],
                    f"{lane} print allocation count differs from histogram")
            exact = {"trace": trace, "histogram": histogram,
                     "print_summary": print_summary,
                     "total_allocations": trace["totals"],
                     "conversion": trace["targets"]["conversion"],
                     "crosscheck": {"trace_histogram_print": True}}
        decoded[lane] = {"status": ("parsed" if metadata is not None else
                                     ("unbound" if files else "pending_decode")),
                         "metadata_bound": metadata is not None,
                         "artifacts": artifacts, "stack_witnesses": hits,
                         "exact_attribution": exact}
    statuses = {value["status"] for value in decoded.values()}
    complete = statuses == {"parsed"} and all(
        value["exact_attribution"] is not None for value in decoded.values()
    )
    comparison: dict[str, Any] = {}
    if complete:
        before = decoded["heaptrack-before"]["exact_attribution"]
        after = decoded["heaptrack-after"]["exact_attribution"]
        shares = {
            "before": {
                "allocation_events": fraction(
                    before["conversion"]["allocation_events"],
                    before["total_allocations"]["allocation_events"],
                ),
                "requested_bytes": fraction(
                    before["conversion"]["requested_bytes"],
                    before["total_allocations"]["requested_bytes"],
                ),
            },
            "after": {
                "allocation_events": fraction(
                    after["conversion"]["allocation_events"],
                    after["total_allocations"]["allocation_events"],
                ),
                "requested_bytes": fraction(
                    after["conversion"]["requested_bytes"],
                    after["total_allocations"]["requested_bytes"],
                ),
            },
        }
        comparison = {
            "before": {"total_allocations": before["total_allocations"],
                       "conversion": before["conversion"]},
            "after": {"total_allocations": after["total_allocations"],
                      "conversion": after["conversion"]},
            "delta": {
                target: {
                    field: after[target][field] - before[target][field]
                    for field in ("allocation_events", "requested_bytes",
                                  "deallocation_events", "deallocated_bytes")
                }
                for target in ("total_allocations", "conversion")
            },
            "conversion_share_of_total": shares,
        }
    return {
        "before": decoded["heaptrack-before"],
        "after": decoded["heaptrack-after"],
        "complete": complete,
        "comparison": comparison,
        "timing_claim": False,
        "scope": "whole-process Heaptrack trace with exact full ancestry (including inline frames) for convert_shape_to_escher_with_sound_mapping",
        "interpretation": "diagnostic only; conversion attribution is compared with total process allocations and carries no timing, RSS, physical-copy, or causal cost claim",
    }


def validate_heap_lane(ctx: Context, plan: dict[str, Any], observer: dict[str, Any],
                       builds: dict[str, dict[str, Any]], lane: str, cleanup: Any,
                       cleanup_verified: bool, contract: tuple[str, str, str]) -> dict[str, Any]:
    directory = ctx.packet / lane
    require(directory.is_dir(), f"missing {lane} observer directory")
    complete = read_json(directory / "complete.json", f"{lane} complete.json")
    source = source_manifest(read_json(directory / "source.json", f"{lane} source manifest"),
                             f"{lane} source manifest")
    receipts_path, _ = ctx.artifact(complete.get("receipts"), f"{lane} receipts")
    assert receipts_path is not None
    rows = read_json(receipts_path, f"{lane} receipts")
    require(complete.get("children") == 1 and isinstance(rows, list) and len(rows) == 1,
            f"{lane} must contain exactly one payload write process")
    row = rows[0]
    leg = "before" if lane == "heaptrack-before" else "after"
    job = {"lane": lane, "block": 0, "mode": observer["heaptrack"]["mode"],
           "shape": observer["heaptrack"]["shape"], "leg": leg,
           "samples": observer["heaptrack"]["samples"],
           "warmup": observer["heaptrack"]["warmup"]}
    require(all(row.get(key) == job[key] for key in ("leg", "block", "mode", "shape")),
            f"{lane} identity changed")
    require(row.get("exit_code") == 0, f"{lane} process failed")
    expected_binary = builds[leg]["manifest"]["binaries"]["native"]
    binary = row.get("binary")
    require(isinstance(binary, dict)
            and binary.get("path") == expected_binary.get("path")
            and binary.get("bytes") == expected_binary.get("bytes")
            and binary.get("sha256") == expected_binary.get("sha256"),
            f"{lane} native binary receipt differs from build identity")
    binary_state = verify_binary(ctx, binary, f"{lane} native binary", cleanup, cleanup_verified)
    log_path, log_receipt = ctx.artifact(row.get("log"), f"{lane} log")
    report_path, report_receipt = ctx.artifact(row.get("report"), f"{lane} report")
    assert log_path is not None and report_path is not None
    traces = row.get("traces")
    require(isinstance(traces, list) and traces, f"{lane} heaptrack trace receipt is missing")
    trace_artifacts = []
    for index, trace in enumerate(traces):
        trace_path, trace_receipt = ctx.artifact(trace, f"{lane} trace {index}")
        assert trace_path is not None
        trace_artifacts.append({"path": ctx.relative(trace_path), "bytes": trace_receipt["bytes"],
                                "sha256": trace_receipt["sha256"]})
    command = row.get("command")
    require(isinstance(command, list) and all(isinstance(item, str) for item in command),
            f"{lane} command is malformed")
    require([ctx.normalized(item) for item in command]
            == expected_heap_command(ctx, plan, job, expected_binary, report_receipt, lane),
            f"{lane} capture command changed")
    report = validate_report(read_json(report_path, f"{lane} report"), job, lane,
                             Path(str(expected_binary.get("path"))).name,
                             contract)
    expected_source = builds[leg]["source"]["files"]
    require(source["files"] == expected_source,
            f"{lane} source receipt differs from {leg} build source files")
    return {"processes": 1, "scope": observer["heaptrack"]["scope"],
            "binary": binary_state, "report": {"path": ctx.relative(report_path),
                                                 "sha256": report_receipt["sha256"]},
            "source": report["source"], "traces": trace_artifacts,
            "artifacts": {"log": {"path": ctx.relative(log_path),
                                    "sha256": log_receipt["sha256"]},
                          "report": {"path": ctx.relative(report_path),
                                      "sha256": report_receipt["sha256"]}}}


def analyze(packet: Path | str = PACKET) -> dict[str, Any]:
    """Validate the frozen observer evidence and return a deterministic summary.

    This function is pure with respect to the packet: it reads files, hashes
    retained artifacts, and returns data.  It does not create ``analysis.json``
    or hash this source file, so a final analyzer can call it before sealing
    the packet without receipt regeneration.
    """

    ctx = Context(Path(packet))
    plan, observer = load_plan(ctx)
    contract = probe_contract(plan)
    cleanup, cleanup_verified = load_cleanup(ctx)
    builds = {
        leg: load_build(ctx, leg, cleanup, cleanup_verified)
        for leg in LEGS
    }
    before = builds["before"]["source"]["files"]
    after = builds["after"]["source"]["files"]
    changed = sorted(name for name in set(before) | set(after) if before.get(name) != after.get(name))
    allowlist = sorted(plan["source_allowlist"])
    require(changed and set(changed).issubset(allowlist),
            f"build source census changed outside explicit allowlist {allowlist}: {changed}")
    perf = validate_perf(ctx, plan, observer, builds, cleanup, cleanup_verified,
                         contract)
    heap_before = validate_heap_lane(ctx, plan, observer, builds, "heaptrack-before",
                                     cleanup, cleanup_verified, contract)
    heap_after = validate_heap_lane(ctx, plan, observer, builds, "heaptrack-after",
                                    cleanup, cleanup_verified, contract)
    require(heap_before["source"] == heap_after["source"],
            "heaptrack before/after source identity differs")
    perf_payload_sources = [
        entry["source"] for entry in perf["receipts"]
        if entry["identity"]["shape"] == "payload"
    ]
    require(perf_payload_sources
            and all(source == perf_payload_sources[0] for source in perf_payload_sources)
            and heap_before["source"] == perf_payload_sources[0],
            "payload report source identity is not shared across observer lanes")
    diagnostic = heaptrack_diagnostic(ctx, {
        "heaptrack-before": heap_before,
        "heaptrack-after": heap_after,
    })
    return {
        "schema": "litchi-0781-observer-analysis-v1",
        "plan_schema": observer["schema"],
        # Path/custody details are validated above but deliberately omitted
        # from the returned mapping.  Retention and worktree relocation are
        # lifecycle changes and must not invalidate observer-analysis.json.
        "capture": {"cpu": plan["cpu"], "custody_checked": True},
        "builds": {leg: {"source_revision": builds[leg]["source"]["revision"],
                          "source_files": builds[leg]["source"]["files"],
                          "native_binary": builds[leg]["binary_state"]["native"]}
                   for leg in LEGS},
        "perf": perf,
        "heaptrack": {"before": heap_before, "after": heap_after,
                       "source_output_parity": {
                           "source": heap_before["source"], "output": None,
                           "checked": True,
                       },
                       "diagnostic": diagnostic},
        "limits": [
            "perf instructions and cycles are whole-process counters including setup and verification",
            "perf pairs are descriptive before/after evidence, not operation-region timing",
            "heaptrack conversion attribution is diagnostic only and carries no timing or causal claim",
        ],
    }


def main() -> None:
    packet = Path(sys.argv[1]) if len(sys.argv) > 1 else PACKET
    result = analyze(packet)
    print(json.dumps({"schema": result["schema"], "perf_processes": result["perf"]["processes"],
                      "heaptrack_processes": result["heaptrack"]["before"]["processes"]
                      + result["heaptrack"]["after"]["processes"],
                      "heaptrack_diagnostic": result["heaptrack"]["diagnostic"]["complete"]},
                     sort_keys=True))


if __name__ == "__main__":
    main()
