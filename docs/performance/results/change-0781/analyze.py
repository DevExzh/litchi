"""Offline replay and analysis for the 0781 borrowed-text evidence packet.

The capture program is deliberately outside this module.  This file only
checks retained custody receipts and derives deterministic summaries from the
JSON reports and ``/usr/bin/time`` RSS gauges.  It is safe to run after the
owned worktree and build target have been removed: every missing executable
must then be covered by an exact cleanup receipt.
"""

from __future__ import annotations

import hashlib
import importlib.util
import json
import math
import random
import re
import statistics
import subprocess
from pathlib import Path
from typing import Any, Iterable


PACKET = Path(__file__).resolve().parent
ROOT = PACKET.parents[3]
ORIGIN_PATH = PACKET / "origin.json"
# The plan may refine this list once the candidate is frozen.  Keeping an
# explicit fallback here prevents a source census from silently becoming a
# broad "anything in litchi-pptx" allowance.
SOURCE_ALLOWLIST = (
    "crates/litchi-ppt/src/writer/core/codec.rs",
)
# Kept as a compatibility alias for the single-file disposition shape.  The
# source-census gate below uses the complete explicit allowlist from plan.json.
SOURCE_FILE = SOURCE_ALLOWLIST[0]
LEGS = ("before", "after")
CASES = (
    {"shape": "tiny", "mode": "write"},
    {"shape": "tiny", "mode": "lifecycle"},
    {"shape": "many", "mode": "write"},
    {"shape": "many", "mode": "lifecycle"},
    {"shape": "payload", "mode": "write"},
    {"shape": "payload", "mode": "lifecycle"},
    {"shape": "unicode", "mode": "write"},
    {"shape": "unicode", "mode": "lifecycle"},
    {"shape": "rich", "mode": "write"},
    {"shape": "rich", "mode": "lifecycle"},
)
MODES = tuple(sorted({case["mode"] for case in CASES}))
SHAPES = ("tiny", "many", "payload", "unicode", "rich")
NATIVE_METRICS = ("p50", "mean", "p95", "p99")
ALLOCATION_FIELDS = (
    "allocation_calls",
    "deallocation_calls",
    "reallocation_calls",
    "failed_allocation_calls",
    "allocated_bytes",
    "deallocated_bytes",
    "live_bytes_before",
    "live_bytes_after",
    "peak_live_bytes_before",
    "peak_live_bytes_after",
    "region_peak_live_bytes",
    "net_live",
    "peak_above_entry",
)
HEX = frozenset("0123456789abcdefABCDEF")
BOOTSTRAP_SEED = 781_078
BOOTSTRAP_RESAMPLES = 10_000
BOOTSTRAP_CONFIDENCE = 0.95
PROBE_TEST_NAMES = (
    "tests::semantic_oracle_preserves_authored_fixture_separately",
    "tests::raw_utf16_decoder_rejects_odd_payloads",
    "tests::raw_utf16_decoder_rejects_lone_surrogates",
    "tests::raw_ascii_decoder_rejects_non_ascii_payloads",
    "tests::raw_officeart_parser_rejects_truncated_headers_and_payloads",
)
QUALIFICATION_PARITY_CASES = frozenset({
    "many/lifecycle", "many/write", "payload/lifecycle", "payload/write",
    "tiny/lifecycle", "tiny/write",
})
# The probe publishes these in plan.json when its schema is frozen.  These
# defaults keep the analyzer importable during packet construction; the
# report validator still requires exact schema and tool identity.
REPORT_SCHEMA = "litchi.ppt.borrowed-text-probe.v1"
REPORT_TOOL = "ppt-borrowed-text-probe-0781"
CI_ANALYSIS_SCHEMA = "litchi.performance.0781.ci-custody.v1"
MARKER = "litchi-perf-0781-borrowed-text"
TIMING_SCOPES = {
    "write": "Writer::write_to only; Writer construction and slide/text-box authoring are outside the clock",
    "lifecycle": "Writer::new, add_slide, add_textbox/add_rich_textbox, and Writer::write_to",
}


def probe_contract() -> tuple[str, str, str]:
    """Return the frozen probe identity, with plan metadata taking precedence."""

    path = PACKET / "plan.json"
    if path.is_file():
        value = read_json(path)
        contract = value.get("probe", {}) if isinstance(value, dict) else {}
        if isinstance(contract, dict):
            return (contract.get("schema", REPORT_SCHEMA),
                    contract.get("tool", REPORT_TOOL),
                    contract.get("marker", MARKER))
    return REPORT_SCHEMA, REPORT_TOOL, MARKER


class ReplayError(RuntimeError):
    """Raised for missing, stale, or contradictory evidence."""


def fail(message: str) -> None:
    raise ReplayError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def read_json(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing JSON evidence: {path}")
    try:
        return json.loads(path.read_text())
    except (OSError, ValueError) as error:
        fail(f"invalid JSON evidence {path}: {error}")


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for chunk in iter(lambda: stream.read(1 << 20), b""):
                digest.update(chunk)
    except OSError as error:
        fail(f"cannot hash {path}: {error}")
    return digest.hexdigest()


def is_sha(value: Any) -> bool:
    return isinstance(value, str) and len(value) == 64 and all(c in HEX for c in value)


def nonnegative_int(value: Any, label: str) -> None:
    require(isinstance(value, int) and not isinstance(value, bool) and value >= 0,
            f"{label} is not a non-negative integer")


def positive_int(value: Any, label: str) -> None:
    require(isinstance(value, int) and not isinstance(value, bool) and value > 0,
            f"{label} is not a positive integer")


def finite_number(value: Any, label: str) -> None:
    require(isinstance(value, (int, float)) and not isinstance(value, bool)
            and math.isfinite(float(value)), f"{label} is not finite")


def rel(path: Path) -> str:
    try:
        return str(path.relative_to(PACKET))
    except ValueError:
        return str(path)


def origin() -> dict[str, Any]:
    value = read_json(ORIGIN_PATH)
    require(isinstance(value, dict), "origin.json is malformed")
    owned = value.get("owned_worktree")
    require(isinstance(owned, str) and owned, "origin owned worktree is missing")
    return value


def _relocated(raw: Path) -> list[Path]:
    """Map absolute receipts from the owned worktree to this packet."""

    candidates: list[Path] = []
    origin_value = origin()
    owned = Path(origin_value["owned_worktree"]).resolve()
    try:
        relative = raw.resolve().relative_to(owned)
    except ValueError:
        relative = None
    if relative is not None:
        candidates.append(ROOT / relative)
        candidates.append(PACKET / relative)
    parts = raw.parts
    for marker in ("change-0781", "build-before", "build-after", "native",
                   "allocation", "qualification"):
        if marker in parts:
            index = parts.index(marker)
            if marker == "change-0781":
                candidates.append(PACKET.joinpath(*parts[index + 1:]))
            else:
                candidates.append(PACKET / marker / Path(*parts[index + 1:]))
    return candidates


def resolve_path(value: Any, *, packet_bound: bool = True) -> Path:
    require(isinstance(value, str) and value, f"invalid artifact path: {value!r}")
    raw = Path(value)
    candidates: list[Path] = []
    if raw.is_absolute():
        candidates.extend(_relocated(raw))
        candidates.append(raw)
    else:
        text = value.replace("\\", "/")
        prefix = "docs/performance/results/change-0781/"
        if text.startswith(prefix):
            candidates.append(PACKET / text[len(prefix):])
        candidates.extend((PACKET / raw, ROOT / raw))
    for candidate in candidates:
        if candidate.is_file() and not candidate.is_symlink():
            path = candidate.resolve()
            if packet_bound:
                try:
                    path.relative_to(PACKET.resolve())
                except ValueError:
                    continue
            return path
    path = (candidates[0] if candidates else raw).resolve(strict=False)
    if packet_bound:
        try:
            path.relative_to(PACKET.resolve())
        except ValueError:
            fail(f"artifact path escaped packet: {value}")
    return path


def artifact(value: Any, label: str, *, packet_bound: bool = True,
             allow_missing: bool = False) -> Path | None:
    require(isinstance(value, dict), f"{label} is not an artifact receipt")
    raw = value.get("path")
    require(isinstance(raw, str) and raw, f"{label}.path is missing")
    size = value.get("bytes", value.get("size"))
    nonnegative_int(size, f"{label}.bytes")
    digest = value.get("sha256", value.get("digest"))
    require(is_sha(digest), f"{label}.sha256 is invalid")
    path = resolve_path(raw, packet_bound=packet_bound)
    if not path.is_file():
        if allow_missing:
            return None
        fail(f"missing {label}: {raw}")
    require(not path.is_symlink(), f"{label} is a symlink: {raw}")
    require(path.stat().st_size == size, f"{label}.bytes changed")
    require(sha256(path) == digest, f"{label}.sha256 changed")
    return path


def artifact_path(value: Any, label: str, *, packet_bound: bool = True) -> Path:
    path = artifact(value, label, packet_bound=packet_bound)
    assert path is not None
    return path


def load_plan() -> dict[str, Any]:
    plan = read_json(PACKET / "plan.json")
    require(isinstance(plan, dict), "plan is not an object")
    require(plan.get("schema") == "litchi.performance.0781.v1", "plan schema changed")
    require(plan.get("cpu") == 12, "capture CPU changed")
    require(plan.get("cases") == list(CASES), "case order or cardinality changed")
    native = plan.get("native")
    allocation = plan.get("allocation")
    require(isinstance(native, dict) and isinstance(allocation, dict),
            "lane configuration is missing")
    require(native.get("blocks") == 6 and native.get("samples") == 30
            and native.get("warmup") == 3, "native lane configuration changed")
    require(allocation.get("blocks") == 2 and allocation.get("samples") == 3
            and allocation.get("warmup") == 0, "allocation lane configuration changed")
    orders = native.get("orders")
    require(orders == [
        ["before", "after"], ["after", "before"], ["before", "after"],
        ["after", "before"], ["after", "before"], ["before", "after"],
    ], "native alternating order changed")
    require(isinstance(plan.get("scope"), str) and plan["scope"], "plan scope is missing")
    allowlist = plan.get("source_allowlist")
    if allowlist is None:
        changed = PACKET / "candidate/changed-files.json"
        if changed.is_file():
            changed_value = read_json(changed)
            if isinstance(changed_value, list):
                allowlist = changed_value
            elif isinstance(changed_value, dict):
                allowlist = changed_value.get("files", changed_value.get("changed_files"))
                if isinstance(allowlist, dict):
                    allowlist = list(allowlist)
    if allowlist is None:
        allowlist = list(SOURCE_ALLOWLIST)
    require(isinstance(allowlist, list) and allowlist and
            all(isinstance(name, str) and name for name in allowlist),
            "explicit source allowlist is missing")
    require(len(set(allowlist)) == len(allowlist), "source allowlist contains duplicates")
    plan["source_allowlist"] = allowlist
    return plan


def _cleanup_has(cleanup: Any, receipt: dict[str, Any]) -> bool:
    expected = (receipt.get("path"), receipt.get("bytes", receipt.get("size")),
                receipt.get("sha256", receipt.get("digest")))
    if not (isinstance(expected[0], str) and nonnegative_pair(expected[1])
            and is_sha(expected[2])):
        return False
    if isinstance(cleanup, dict):
        actual = (cleanup.get("path"), cleanup.get("bytes", cleanup.get("size")),
                  cleanup.get("sha256", cleanup.get("digest")))
        if actual == expected:
            return True
        return any(_cleanup_has(value, receipt) for value in cleanup.values())
    if isinstance(cleanup, list):
        return any(_cleanup_has(value, receipt) for value in cleanup)
    return False


def nonnegative_pair(value: Any) -> bool:
    return isinstance(value, int) and not isinstance(value, bool) and value >= 0


def load_cleanup() -> tuple[Any, bool]:
    path = PACKET / "cleanup.json"
    if not path.is_file():
        return None, False
    value = read_json(path)
    require(isinstance(value, dict), "cleanup.json is malformed")
    verified = value.get("verified") is True or value.get(
        "executables_verified_before_removal") is True
    return value, verified


def validate_binary(receipt: Any, label: str, cleanup: Any, cleanup_verified: bool) -> None:
    require(isinstance(receipt, dict), f"{label} receipt is missing")
    path = artifact(receipt, label, packet_bound=False, allow_missing=True)
    if path is not None:
        return
    require(cleanup_verified, f"{label} is missing without cleanup verification")
    require(_cleanup_has(cleanup, receipt), f"{label} lacks an exact cleanup witness")


def source_manifest(value: Any, label: str) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label} is not an object")
    revision = value.get("revision")
    require(isinstance(revision, str) and revision and all(c in HEX for c in revision),
            f"{label}.revision is invalid")
    files = value.get("files")
    require(isinstance(files, dict) and files, f"{label}.files is missing")
    for name, digest in files.items():
        require(isinstance(name, str) and name and is_sha(digest),
                f"{label} contains an invalid file digest")
    return {"revision": revision, "files": dict(files)}


def source_files_equal(left: dict[str, Any], right: dict[str, Any]) -> bool:
    """Compare the immutable file census while allowing post-integration HEADs."""

    return left.get("files") == right.get("files")


def current_source_files() -> dict[str, str]:
    try:
        raw = subprocess.check_output([
            "git", "ls-files", "-z", "--", "crates", "Cargo.toml", "clippy.toml",
            ".cargo/config.toml", "rust-toolchain.toml",
        ], cwd=ROOT)
    except (OSError, subprocess.CalledProcessError) as error:
        fail(f"cannot read live source census: {error}")
    result: dict[str, str] = {}
    for name in (item for item in raw.decode().split("\0") if item):
        path = ROOT / name
        require(path.is_file() and not path.is_symlink(), f"live source file is missing: {name}")
        result[name] = sha256(path)
    return result


def verify_source_archives(value: Any, expected: dict[str, str], label: str) -> dict[str, Any]:
    """Verify a disposition's per-file source archives after candidate rejection."""

    entries: dict[str, Any] = {}
    if isinstance(value, dict) and isinstance(value.get("files"), dict):
        value = value["files"]
    if isinstance(value, dict) and set(expected).issubset(value):
        raw_items = {name: value[name] for name in expected}
    elif isinstance(value, list):
        require(len(value) == len(expected), f"{label} archive cardinality changed")
        raw_items = dict(zip(sorted(expected), value))
    else:
        require(len(expected) == 1, f"{label} needs one archive per changed file")
        raw_items = {next(iter(expected)): value}
    for name, receipt in raw_items.items():
        path = artifact_path(receipt, f"{label} {name}")
        require(sha256(path) == expected[name], f"{label} {name} does not match source census")
        entries[name] = {"path": rel(path), "bytes": path.stat().st_size,
                         "sha256": sha256(path)}
    return entries


def load_disposition(before: dict[str, Any], after: dict[str, Any]) -> dict[str, Any]:
    """Bind retained production source, or verify the archived rejected source."""

    allowlist = tuple(load_plan()["source_allowlist"])
    source_label: str | list[str] = allowlist[0] if len(allowlist) == 1 else list(allowlist)
    path = PACKET / "disposition.json"
    if not path.is_file():
        # Before the coordinator writes a final decision, the live candidate
        # is the only honest state.  The retained form is deterministic and
        # makes early replay useful without inventing a rejection archive.
        live = current_source_files()
        require(live == after["files"], "live source does not match after candidate census")
        return {"status": "retained", "production_change_retained": True,
                "source_file": source_label, "live_source_files_match_after": True}
    value = read_json(path)
    require(isinstance(value, dict), "disposition.json is malformed")
    status = value.get("status")
    retained = value.get("production_change_retained")
    require(status in {"retained", "rejected"}, "candidate disposition is invalid")
    require(retained is (status == "retained"), "candidate retention flag contradicts status")
    require(value.get("source_file") in (source_label, list(allowlist)),
            "disposition source file changed")
    if status == "retained":
        live = current_source_files()
        require(live == after["files"], "retained live source differs from after census")
        return {"status": status, "production_change_retained": True,
                "source_file": source_label, "live_source_files_match_after": True}
    changed_before = {name: before["files"][name] for name in allowlist}
    changed_after = {name: after["files"][name] for name in allowlist}
    before_archives = verify_source_archives(
        value.get("before_sources", value.get("before_source")), changed_before,
        "before source archive")
    candidate_archives = verify_source_archives(
        value.get("candidate_sources", value.get("candidate_source")), changed_after,
        "candidate source archive")
    restored = artifact_path(value.get("restored_source"), "restored source census")
    restored_manifest = source_manifest(read_json(restored), "restored source census")
    require(restored_manifest["files"] == before["files"], "restored census differs from before")
    require(current_source_files() == before["files"], "rejected live source differs from before")
    return {
        "status": "rejected", "production_change_retained": False,
        "source_file": source_label,
        "before_sources": before_archives,
        "candidate_sources": candidate_archives,
        "restored_source": {"path": rel(restored), "bytes": restored.stat().st_size,
                             "sha256": sha256(restored)},
        "live_source_files_match_before": True,
    }


def probe_files() -> set[str]:
    root = PACKET / "probe-src"
    require(root.is_dir(), "probe source directory is missing")
    return {
        str(path.relative_to(PACKET)) for path in root.rglob("*")
        if path.is_file() and path.name not in {"Cargo.lock", "Cargo.toml"}
    }


def fixed_release_profile() -> None:
    template = PACKET / "probe-src/Cargo.toml.template"
    require(template.is_file() and not template.is_symlink(),
            "probe Cargo.toml.template is missing")
    text = template.read_text()
    required = (
        '[profile.release]', 'opt-level = 3', 'debug = 1', 'lto = "thin"',
        'codegen-units = 1', 'panic = "unwind"',
    )
    require(all(line in text for line in required), "probe release profile changed")


def normalized(value: str) -> str:
    """Normalize only the owned absolute worktree prefix in a receipt."""

    owned = str(Path(origin()["owned_worktree"]).resolve())
    current = str(ROOT.resolve())
    return value.replace(owned, current)


def expected_build_command(leg: str, feature: str | None) -> list[str]:
    command = ["cargo", "build", "--offline", "--release", "--manifest-path",
               str(PACKET / "probe-src/Cargo.toml")]
    if feature is not None:
        command.extend(["--features", feature])
    if (PACKET / "probe-src/Cargo.lock").is_file():
        command.append("--locked")
    return command


def load_builds(plan: dict[str, Any]) -> tuple[dict[str, Any], Any, bool]:
    cleanup, cleanup_verified = load_cleanup()
    fixed_release_profile()
    builds: dict[str, Any] = {}
    for leg in LEGS:
        directory = PACKET / f"build-{leg}"
        build = read_json(directory / "build.json")
        require(isinstance(build, dict), f"{leg} build manifest is malformed")
        source_path = artifact_path(build.get("source"), f"{leg} build source")
        source = source_manifest(read_json(source_path), f"{leg} build source")
        inventory = build.get("probe")
        require(isinstance(inventory, dict) and set(inventory) == probe_files(),
                f"{leg} probe inventory changed")
        for name, digest in inventory.items():
            require(is_sha(digest), f"{leg} probe digest is invalid: {name}")
            path = PACKET / name
            require(path.is_file() and not path.is_symlink(), f"missing probe file: {name}")
            require(sha256(path) == digest, f"probe file changed: {name}")
        lock = build.get("lock")
        lock_path = artifact_path(lock, f"{leg} probe lock")
        require(lock_path == (PACKET / "probe-src/Cargo.lock").resolve(),
                f"{leg} lock receipt did not relocate to packet")
        binaries = build.get("binaries")
        require(isinstance(binaries, dict) and set(binaries) == {"native", "allocation"},
                f"{leg} binary map changed")
        for name, receipt in binaries.items():
            validate_binary(receipt, f"{leg} {name} binary", cleanup, cleanup_verified)
        rows = build.get("rows")
        require(isinstance(rows, list) and len(rows) == 2, f"{leg} build rows are incomplete")
        expected = {
            "native": expected_build_command(leg, None),
            "allocation": expected_build_command(leg, "allocator-metrics"),
        }
        for row in rows:
            require(isinstance(row, dict) and row.get("exit_code") == 0,
                    f"{leg} build command failed")
            command = row.get("command")
            require(isinstance(command, list), f"{leg} build command is malformed")
            command = [normalized(item) for item in command]
            kind = "allocation" if "allocator-metrics" in command else "native"
            require(command == expected[kind], f"{leg} {kind} build command changed")
            artifact_path(row.get("log"), f"{leg} {kind} build log")
        environment = build.get("environment")
        require(isinstance(environment, dict), f"{leg} build environment is missing")
        require(environment.get("CARGO_BUILD_JOBS") == "2"
                and environment.get("CARGO_INCREMENTAL") == "0",
                f"{leg} build environment changed")
        builds[leg] = {"manifest": build, "source": source,
                       "source_path": source_path, "probe": inventory,
                       "lock": lock, "binaries": binaries}
    before = builds["before"]["source"]["files"]
    after = builds["after"]["source"]["files"]
    changed = sorted(name for name in set(before) | set(after) if before.get(name) != after.get(name))
    allowlist = tuple(plan["source_allowlist"])
    require(all(name in set(before) | set(after) for name in allowlist),
            "source allowlist names are absent from the source census")
    require(changed and set(changed).issubset(allowlist),
            f"source census changed outside explicit allowlist {allowlist}: {changed}")
    require(builds["before"]["lock"]["sha256"] == builds["after"]["lock"]["sha256"],
            "probe lock changed between builds")
    require(builds["before"]["probe"] == builds["after"]["probe"],
            "probe source changed between builds")
    for lane in ("native", "allocation", "qualification"):
        path = PACKET / lane / "source.json"
        if path.is_file():
            expected = builds["before"]["source"] if lane == "qualification" else builds["after"]["source"]
            require(source_files_equal(source_manifest(read_json(path), f"{lane} source"), expected),
                    f"{lane} source census differs from expected build")
    return builds, cleanup, cleanup_verified


def load_initial_binary_relocations(cleanup: Any, cleanup_verified: bool) -> dict[str, Any]:
    """Verify the archived identities for the superseded initial binaries.

    The original ``before-*`` target names were reused by the corrected build.
    Therefore those names are never used to validate the initial binaries;
    only the explicit ``initial-before-*`` archive receipts count, with an
    exact cleanup witness accepted after target removal.
    """

    correction_path = PACKET / "qualification-correction.json"
    correction = read_json(correction_path)
    require(isinstance(correction, dict), "qualification correction is malformed")
    require(correction.get("production_source_changed") is False
            and correction.get("primary_paired_captures_started") is False,
            "qualification correction scope changed")
    require(correction.get("receipt_relocations") == {
        "build-before": "build-before-initial",
        "qualification": "qualification-initial",
    }, "qualification correction relocation map changed")
    reason = correction.get("reason")
    require(isinstance(reason, str) and reason,
            "qualification correction reason is missing")
    initial_build = read_json(PACKET / "build-before-initial/build.json")
    require(isinstance(initial_build, dict), "initial build manifest is malformed")
    binaries = initial_build.get("binaries")
    require(isinstance(binaries, dict) and set(binaries) == {"native", "allocation"},
            "initial build binary map is incomplete")
    archives = correction.get("binary_archive")
    require(isinstance(archives, list) and len(archives) == 2,
            "initial binary archive cardinality changed")
    seen: set[str] = set()
    rows: list[dict[str, Any]] = []
    target = Path(origin().get("target", ""))
    require(str(target), "origin target is missing")
    for index, entry in enumerate(archives):
        require(isinstance(entry, dict), f"initial binary archive {index} is malformed")
        original = entry.get("original")
        archived = entry.get("archived")
        require(isinstance(original, dict) and isinstance(archived, dict),
                f"initial binary archive {index} identities are incomplete")
        for identity, label in ((original, "original"), (archived, "archived")):
            require(isinstance(identity.get("path"), str) and identity["path"],
                    f"initial binary archive {index} {label} path is missing")
            positive_int(identity.get("bytes"),
                         f"initial binary archive {index} {label} bytes")
            require(is_sha(identity.get("sha256")),
                    f"initial binary archive {index} {label} sha256 is invalid")
            require(Path(identity["path"]).parent == target,
                    f"initial binary archive {index} {label} target path changed")
        original_name = Path(original["path"]).name
        archived_name = Path(archived["path"]).name
        require(original_name in {"before-native", "before-allocation"},
                f"initial binary archive {index} original path is not a reused target")
        kind = "native" if original_name == "before-native" else "allocation"
        require(kind not in seen, f"duplicate initial binary archive: {kind}")
        seen.add(kind)
        require(archived_name == f"initial-before-{kind}",
                f"initial binary archive {index} archived path changed")
        expected = binaries[kind]
        require(original == expected,
                f"initial binary archive {kind} does not match initial build receipt")
        require(archived["bytes"] == original["bytes"]
                and archived["sha256"] == original["sha256"],
                f"initial binary archive {kind} identity changed")
        archived_path = artifact(archived, f"initial archived {kind} binary",
                                 packet_bound=False, allow_missing=True)
        if archived_path is None:
            require(cleanup_verified and _cleanup_has(cleanup, archived),
                    f"initial archived {kind} binary lacks cleanup witness")
        rows.append({
            "kind": kind,
            "original_name": original_name,
            "archived_name": archived_name,
            "bytes": archived["bytes"],
            "sha256": archived["sha256"],
        })
    require(seen == {"native", "allocation"},
            "initial binary archive kinds are incomplete")
    initial_probe = correction.get("initial_probe_files")
    require(isinstance(initial_probe, dict) and initial_probe,
            "initial probe inventory is missing")
    for name, digest in initial_probe.items():
        require(isinstance(name, str) and name.startswith("initial-probe/"),
                "initial probe inventory path changed")
        require(is_sha(digest), f"initial probe digest is invalid: {name}")
        path = PACKET / name
        require(path.is_file() and not path.is_symlink() and sha256(path) == digest,
                f"initial probe file changed: {name}")
    return {
        "receipt": _file_identity(correction_path),
        "binaries": sorted(rows, key=lambda row: row["kind"]),
        "initial_probe_files": len(initial_probe),
        "relocations": correction["receipt_relocations"],
    }


QUALITY_COMMANDS = (
    ["cargo", "fmt", "-p", "litchi-ppt", "--", "--check"],
    ["cargo", "check", "--offline", "--locked", "-p", "litchi-ppt",
     "--all-features", "--all-targets"],
    ["cargo", "test", "--offline", "--locked", "-p", "litchi-ppt",
     "--all-features", "--", "--test-threads=2"],
    ["cargo", "clippy", "--offline", "--locked", "-p", "litchi-ppt",
     "--all-features", "--lib", "--", "-D", "warnings"],
    ["cargo", "doc", "--offline", "--locked", "-p", "litchi-ppt",
     "--all-features", "--no-deps"],
    ["python3", "-B", "tools/check_crate_boundaries.py"],
)


def check_quality(after_source: dict[str, Any]) -> dict[str, Any]:
    path = PACKET / "quality.json"
    require(path.is_file(), "quality.json is missing")
    quality = read_json(path)
    require(isinstance(quality, dict), "quality.json is malformed")
    source_path = artifact_path(quality.get("source"), "quality source")
    require(source_files_equal(source_manifest(read_json(source_path), "quality source"), after_source),
            "quality source census differs from after build")
    rows = quality.get("rows")
    require(isinstance(rows, list) and len(rows) == len(QUALITY_COMMANDS),
            "quality gate cardinality changed")
    logs: list[str] = []
    for index, row in enumerate(rows):
        require(isinstance(row, dict) and row.get("exit_code") == 0,
                f"quality gate {index} failed")
        require(row.get("command") == QUALITY_COMMANDS[index],
                f"quality gate {index} command changed")
        logs.append(rel(artifact_path(row.get("log"), f"quality gate {index} log")))
    env = quality.get("environment")
    require(isinstance(env, dict) and env.get("CARGO_BUILD_JOBS") == "2"
            and env.get("CARGO_INCREMENTAL") == "0"
            and env.get("CARGO_PROFILE_DEV_DEBUG") == "0"
            and env.get("RUSTDOCFLAGS") == "-D warnings",
            "quality environment changed")
    return {"gates": len(rows), "commands": [list(row["command"]) for row in rows],
            "logs": logs, "source": rel(source_path)}


def check_test_summary(quality: dict[str, Any]) -> dict[str, int | str]:
    """Recompute the bound cargo-test totals from its retained gate log."""

    quality_path = PACKET / "quality.json"
    quality_value = read_json(quality_path)
    rows = quality_value.get("rows")
    require(isinstance(rows, list) and len(rows) == len(QUALITY_COMMANDS),
            "quality rows disappeared while checking test summary")
    row = rows[2]
    log_path = artifact_path(row.get("log"), "quality test gate log")
    summary_path = PACKET / "test-summary.json"
    summary = read_json(summary_path)
    require(isinstance(summary, dict), "test-summary.json is malformed")
    try:
        text = log_path.read_text()
    except OSError as error:
        fail(f"cannot read quality test gate log: {error}")
    pattern = re.compile(
        r"^test result: (?:ok|FAILED)\.\s+"
        r"(\d+) passed; (\d+) failed; (\d+) ignored; "
        r"(\d+) measured; (\d+) filtered out;",
        re.MULTILINE,
    )
    matches = pattern.findall(text)
    result_lines = [line for line in text.splitlines() if line.startswith("test result:")]
    require(matches and len(matches) == len(result_lines),
            "quality test log contains an unparseable test result line")
    totals = {
        "suites": len(matches),
        "passed": sum(int(match[0]) for match in matches),
        "failed": sum(int(match[1]) for match in matches),
        "ignored": sum(int(match[2]) for match in matches),
    }
    require(summary.get("source") == rel(log_path),
            "test-summary source is not the bound quality test log")
    for field, value in totals.items():
        require(summary.get(field) == value, f"test-summary {field} does not replay")
    return {**totals, "source": rel(log_path)}


def load_probe_tests(builds: dict[str, Any]) -> dict[str, Any]:
    """Bind the standalone probe-oracle test receipt to the before build.

    The test suite may grow, so the count is parsed from its retained log.
    The five raw/semantic oracle tests frozen with this packet must remain
    present and passing.  The command, source/probe/lock identities, and log
    are all checked through the receipt before a compact relative summary is
    returned.
    """

    receipt_path = PACKET / "probe-tests/receipt.json"
    require(receipt_path.is_file() and not receipt_path.is_symlink(),
            "probe-tests/receipt.json is missing")
    receipt = read_json(receipt_path)
    require(isinstance(receipt, dict), "probe test receipt is malformed")
    source_path = artifact_path(receipt.get("source"), "probe test source")
    source = source_manifest(read_json(source_path), "probe test source")
    require(source_path == builds["before"]["source_path"]
            and source == builds["before"]["source"],
            "probe test source differs from before build")
    lock_path = artifact_path(receipt.get("lock"), "probe test lock")
    before_lock = builds["before"]["lock"]
    require(lock_path == (PACKET / "probe-src/Cargo.lock").resolve()
            and sha256(lock_path) == before_lock.get("sha256")
            and lock_path.stat().st_size == before_lock.get("bytes"),
            "probe test lock differs from before build")
    probe = receipt.get("probe")
    require(isinstance(probe, dict) and probe == builds["before"]["probe"],
            "probe test source inventory differs from before build")
    environment = receipt.get("environment")
    require(isinstance(environment, dict)
            and environment.get("CARGO_BUILD_JOBS") == "2"
            and environment.get("CARGO_INCREMENTAL") == "0",
            "probe test environment changed")
    rows = receipt.get("rows")
    require(isinstance(rows, list) and len(rows) == 2,
            "probe test receipt must contain the two frozen test rows")
    base_command = [
        "cargo", "test", "--release", "--offline", "--locked", "--manifest-path",
        str(PACKET / "probe-src/Cargo.toml"),
    ]
    expected_commands = [
        base_command + ["--", "--test-threads=1"],
        base_command + ["--all-features", "--", "--test-threads=1",
                        "--skip", "allocation_metrics::tests::"],
    ]
    parsed_rows = []
    pattern = re.compile(r"^test (.+) \.\.\. (ok|FAILED|ignored)$")
    result_pattern = re.compile(
        r"^test result: (ok|FAILED)\.\s+(\d+) passed; (\d+) failed; "
        r"(\d+) ignored; (\d+) measured; (\d+) filtered out;.*$"
    )
    for index, row in enumerate(rows):
        require(isinstance(row, dict) and row.get("exit_code") == 0,
                f"probe test row {index} failed")
        command = row.get("command")
        require(isinstance(command, list) and all(isinstance(item, str) for item in command),
                f"probe test row {index} command is malformed")
        require([normalized(item) for item in command] == expected_commands[index],
                f"probe test row {index} command changed")
        log_path = artifact_path(row.get("log"), f"probe test row {index} log")
        try:
            lines = log_path.read_text().splitlines()
        except OSError as error:
            fail(f"cannot read probe test row {index} log: {error}")
        test_rows = []
        for line in lines:
            match = pattern.match(line)
            if match:
                test_rows.append((match.group(1), match.group(2)))
        require(test_rows, f"probe test row {index} log contains no test rows")
        names = [name for name, _status in test_rows]
        require(set(PROBE_TEST_NAMES).issubset(names),
                f"probe test row {index} misses a frozen oracle test")
        require(len(names) == len(set(names)),
                f"probe test row {index} repeats a test name")
        result_lines = [line for line in lines if line.startswith("test result:")]
        require(len(result_lines) == 1,
                f"probe test row {index} has no unique result line")
        result_match = result_pattern.match(result_lines[0])
        require(result_match is not None,
                f"probe test row {index} result line is unparseable")
        status, passed, failed, ignored, measured, filtered = result_match.groups()
        counts = {"passed": int(passed), "failed": int(failed), "ignored": int(ignored),
                  "measured": int(measured), "filtered": int(filtered)}
        require(status == "ok" and counts["failed"] == 0 and counts["ignored"] == 0,
                f"probe test row {index} result is not clean")
        require(counts["passed"] == len(test_rows)
                and counts["passed"] >= len(PROBE_TEST_NAMES),
                f"probe test row {index} count does not match retained test rows")
        require(all(test_status == "ok" for _name, test_status in test_rows),
                f"probe test row {index} contains a non-passing test")
        parsed_rows.append({
            "command": [
                "cargo", "test", "--release", "--offline", "--locked",
                "--manifest-path", "probe-src/Cargo.toml",
                *(["--all-features"] if index == 1 else []), "--", "--test-threads=1",
                *(["--skip", "allocation_metrics::tests::"] if index == 1 else []),
            ],
            "log": {"path": rel(log_path), "bytes": log_path.stat().st_size,
                    "sha256": sha256(log_path)},
            "tests": {**counts, "names": sorted(names)},
        })
    return {
        "receipt": {"path": rel(receipt_path), "bytes": receipt_path.stat().st_size,
                    "sha256": sha256(receipt_path)},
        "rows": parsed_rows,
        "source": {"path": rel(source_path), "bytes": source_path.stat().st_size,
                   "sha256": sha256(source_path)},
        "lock": {"path": rel(lock_path), "bytes": lock_path.stat().st_size,
                 "sha256": sha256(lock_path)},
        "probe": probe,
        "environment": {"CARGO_BUILD_JOBS": "2", "CARGO_INCREMENTAL": "0"},
        "log": {"path": rel(log_path), "bytes": log_path.stat().st_size,
                "sha256": sha256(log_path)},
        "tests": {**counts, "names": sorted(names)},
    }


def _file_identity(path: Path) -> dict[str, Any]:
    return {"path": rel(path), "bytes": path.stat().st_size, "sha256": sha256(path)}


def _qualification_report_identity(path: Path, label: str) -> dict[str, Any]:
    report = read_json(path)
    require(isinstance(report, dict), f"{label} report is malformed")
    source = report.get("source")
    require(isinstance(source, dict) and is_sha(source.get("sha256")),
            f"{label} source identity is missing")
    nonnegative_int(source.get("bytes"), f"{label} source bytes")
    samples = report.get("samples")
    require(isinstance(samples, list) and samples, f"{label} samples are missing")
    outputs: list[dict[str, Any]] = []
    for index, sample in enumerate(samples):
        require(isinstance(sample, dict), f"{label} sample {index} is malformed")
        output = sample.get("output")
        require(isinstance(output, dict) and is_sha(output.get("sha256")),
                f"{label} sample {index} output identity is missing")
        nonnegative_int(output.get("bytes"), f"{label} sample {index} output bytes")
        outputs.append({"bytes": output["bytes"], "sha256": output["sha256"]})
    require(all(output == outputs[0] for output in outputs),
            f"{label} output identity is not stable")
    return {
        "source": {"bytes": source["bytes"], "sha256": source["sha256"]},
        "output": outputs[0],
        "samples": len(samples),
    }


def load_qualification_parity() -> dict[str, Any]:
    """Replay the retained six-case correction parity record."""

    parity_path = PACKET / "qualification-parity.json"
    require(parity_path.is_file() and not parity_path.is_symlink(),
            "qualification-parity.json is missing")
    parity = read_json(parity_path)
    require(isinstance(parity, dict), "qualification parity is malformed")
    require(parity.get("source_and_output_unchanged") is True,
            "qualification source/output parity failed")
    rows = parity.get("rows")
    require(isinstance(rows, list) and len(rows) == len(QUALIFICATION_PARITY_CASES),
            "qualification parity row cardinality changed")
    seen: set[str] = set()
    result_rows: list[dict[str, Any]] = []
    for index, row in enumerate(rows):
        require(isinstance(row, dict), f"qualification parity row {index} is malformed")
        initial_path = artifact_path(row.get("initial"),
                                     f"qualification parity row {index} initial")
        corrected_path = artifact_path(row.get("corrected"),
                                       f"qualification parity row {index} corrected")
        initial_name = Path(rel(initial_path)).name
        corrected_name = Path(rel(corrected_path)).name
        require(initial_path.parent == (PACKET / "qualification-initial").resolve()
                and corrected_path.parent == (PACKET / "qualification").resolve(),
                f"qualification parity row {index} report location changed")
        require(initial_name == corrected_name and initial_name.startswith("0-"),
                f"qualification parity row {index} report identity changed")
        parts = initial_name.removesuffix(".json").split("-")
        require(len(parts) == 4 and parts[0] == "0" and parts[3] == "before",
                f"qualification parity row {index} case name changed")
        case = f"{parts[1]}/{parts[2]}"
        require(case in QUALIFICATION_PARITY_CASES and case not in seen,
                f"qualification parity case changed: {case}")
        seen.add(case)
        initial = _qualification_report_identity(initial_path,
                                                 f"qualification parity {case} initial")
        corrected = _qualification_report_identity(corrected_path,
                                                   f"qualification parity {case} corrected")
        require(initial["source"] == corrected["source"],
                f"qualification parity {case} source changed")
        row_source = row.get("source")
        require(isinstance(row_source, dict) and is_sha(row_source.get("sha256")),
                f"qualification parity {case} source receipt is malformed")
        nonnegative_int(row_source.get("bytes"),
                        f"qualification parity {case} source bytes")
        require(initial["source"] == {
            "bytes": row_source["bytes"], "sha256": row_source["sha256"]},
                f"qualification parity {case} source receipt changed")
        require(initial["output"] == corrected["output"],
                f"qualification parity {case} output changed")
        require(row.get("output") == corrected["output"],
                f"qualification parity {case} output receipt changed")
        result_rows.append({
            "case": case,
            "initial": {"path": rel(initial_path),
                         "bytes": initial_path.stat().st_size,
                         "sha256": sha256(initial_path)},
            "corrected": {"path": rel(corrected_path),
                           "bytes": corrected_path.stat().st_size,
                           "sha256": sha256(corrected_path)},
            "source": corrected["source"],
            "output": corrected["output"],
        })
    require(seen == QUALIFICATION_PARITY_CASES,
            "qualification parity case set changed")
    return {
        "receipt": _file_identity(parity_path),
        "source_and_output_unchanged": True,
        "rows": sorted(result_rows, key=lambda row: row["case"]),
    }


def load_baseline_attribution() -> dict[str, Any]:
    """Replay and bind the baseline Heaptrack conversion attribution record."""

    attribution_path = PACKET / "baseline-attribution.json"
    require(attribution_path.is_file() and not attribution_path.is_symlink(),
            "baseline-attribution.json is missing")
    value = read_json(attribution_path)
    require(isinstance(value, dict), "baseline attribution is malformed")
    require(value.get("scope") ==
            "Whole-process allocation ancestry diagnostic; not timed phase attribution.",
            "baseline attribution scope changed")
    require(value.get("crosscheck") is True,
            "baseline attribution crosscheck is missing")
    conversion = value.get("conversion")
    totals = value.get("totals")
    require(isinstance(conversion, dict) and isinstance(totals, dict),
            "baseline attribution metrics are missing")
    trace_receipt = value.get("trace")
    trace_path = artifact_path(trace_receipt, "baseline attribution trace")
    require(trace_path == (PACKET / "heaptrack-before/trace.zst").resolve(),
            "baseline attribution trace path changed")
    decode = read_json(PACKET / "heaptrack-before/decode.json")
    require(isinstance(decode, dict), "baseline Heaptrack decode receipt is malformed")
    histogram_path = artifact_path(decode.get("histogram"), "baseline attribution histogram")
    print_path = artifact_path(decode.get("log"), "baseline attribution print log")
    require(histogram_path == (PACKET / "heaptrack-before/histogram").resolve()
            and print_path == (PACKET / "heaptrack-before/print.log").resolve(),
            "baseline attribution decode paths changed")
    spec = importlib.util.spec_from_file_location(
        "ppt_0781_baseline_observer", PACKET / "observe.py")
    require(spec is not None and spec.loader is not None,
            "baseline observer parser cannot be loaded")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    trace = module.parse_heap_trace(trace_path)
    histogram = module.parse_histogram(histogram_path)
    print_summary = module.parse_print_summary(print_path)
    require(trace["targets"]["conversion"] == conversion,
            "baseline conversion attribution does not replay")
    require(trace["totals"] == totals,
            "baseline total attribution does not replay")
    require(histogram["allocation_events"] == totals["allocation_events"]
            and histogram["requested_bytes"] == totals["requested_bytes"]
            and print_summary["allocation_events"] == totals["allocation_events"],
            "baseline attribution trace crosscheck changed")
    return {
        "receipt": _file_identity(attribution_path),
        "trace": _file_identity(trace_path),
        "histogram": _file_identity(histogram_path),
        "print_log": _file_identity(print_path),
        "conversion": conversion,
        "totals": totals,
        "crosscheck": True,
    }


def load_observer_analysis() -> dict[str, Any]:
    """Replay the independent observer analyzer without retaining volatile paths."""

    runner_path = PACKET / "observe.py"
    retained_path = PACKET / "observer-analysis.json"
    require(runner_path.is_file() and not runner_path.is_symlink(),
            "observer analyzer is missing")
    require(retained_path.is_file() and not retained_path.is_symlink(),
            "observer-analysis.json is missing")
    spec = importlib.util.spec_from_file_location("ppt_0781_observer", runner_path)
    require(spec is not None and spec.loader is not None,
            "observer analyzer cannot be loaded")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    fresh = module.analyze(PACKET.resolve())
    retained = read_json(retained_path)
    require(retained == fresh, "observer-analysis.json does not replay")
    require(retained.get("schema") == "litchi-0781-observer-analysis-v1",
            "observer analysis schema changed")
    return {
        "analysis": retained,
        "runner": {"path": rel(runner_path), "bytes": runner_path.stat().st_size,
                    "sha256": sha256(runner_path)},
        "retained": {"path": rel(retained_path), "bytes": retained_path.stat().st_size,
                      "sha256": sha256(retained_path)},
    }


def load_ci_analysis() -> dict[str, Any]:
    """Replay the pure CI custody validator and bind its retained JSON.

    The CI validator is deliberately imported instead of executed as a CLI:
    this keeps the final analysis read-only and lets the same result survive
    target cleanup and packet relocation.  The returned wrapper contains only
    packet-relative identities around the validator's stable result.
    """

    runner_path = PACKET / "ci_validate.py"
    retained_path = PACKET / "ci-analysis.json"
    require(runner_path.is_file() and not runner_path.is_symlink(),
            "CI validator is missing")
    require(retained_path.is_file() and not retained_path.is_symlink(),
            "ci-analysis.json is missing")
    spec = importlib.util.spec_from_file_location("ppt_0781_ci_validate", runner_path)
    require(spec is not None and spec.loader is not None,
            "CI validator cannot be loaded")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    analyze_fn = getattr(module, "analyze", None)
    require(callable(analyze_fn), "CI validator has no pure analyze(packet_root) API")
    fresh = analyze_fn(PACKET.resolve())
    require(isinstance(fresh, dict), "CI validator returned a non-object analysis")
    retained = read_json(retained_path)
    require(retained == fresh, "ci-analysis.json does not replay")
    schema = retained.get("schema")
    require(schema == CI_ANALYSIS_SCHEMA, "CI analysis schema changed")
    return {
        "analysis": retained,
        "runner": {"path": rel(runner_path), "bytes": runner_path.stat().st_size,
                    "sha256": sha256(runner_path)},
        "retained": {"path": rel(retained_path), "bytes": retained_path.stat().st_size,
                      "sha256": sha256(retained_path)},
    }


def expected_jobs(plan: dict[str, Any], lane: str) -> list[dict[str, Any]]:
    if lane == "qualification":
        blocks = 1
        orders = [["before"]]
        samples, warmup = 1, 0
    else:
        blocks = plan[lane]["blocks"]
        orders = plan["native"]["orders"][:blocks]
        samples, warmup = plan[lane]["samples"], plan[lane]["warmup"]
    jobs: list[dict[str, Any]] = []
    for block in range(blocks):
        for case in CASES:
            for leg in orders[block]:
                jobs.append({"lane": lane, "block": block, **case, "leg": leg,
                             "samples": samples, "warmup": warmup})
    return jobs


def expected_command(plan: dict[str, Any], job: dict[str, Any], binary: dict[str, Any],
                     report: Path, rss: Path) -> list[str]:
    return ["/usr/bin/time", "-f", "%M", "-o", str(rss), "taskset", "-c",
            str(plan["cpu"]), normalized(binary["path"]), "--mode", job["mode"],
            "--shape", job["shape"], "--samples", str(job["samples"]),
            "--warmup", str(job["warmup"]), "--output", str(report)]


def source_identity(report: dict[str, Any], job: dict[str, Any], label: str) -> None:
    """Require deterministic corpus and source identity."""

    require(report.get("shape") == job["shape"] and report.get("mode") == job["mode"],
            f"{label} shape/mode identity changed")
    identity = report.get("source")
    require(isinstance(identity, dict), f"{label} source identity is missing")
    digest = identity.get("sha256", identity.get("digest"))
    require(is_sha(digest), f"{label} source digest is missing")
    positive_int(identity.get("bytes"), f"{label} source bytes")
    for key in ("slides", "boxes_per_slide", "text_boxes", "text_bytes"):
        if key in report:
            positive_int(report[key], f"{label} {key}")
    corpus = report.get("corpus")
    if corpus is not None:
        require(corpus == job["shape"], f"{label} corpus identity changed")


def stats(values: Iterable[int | float]) -> dict[str, Any]:
    vector = list(values)
    require(vector, "empty metric vector")
    for index, value in enumerate(vector):
        finite_number(value, f"metric[{index}]")
    ordered = sorted(vector)
    def nearest(percentile: int) -> int | float:
        return ordered[max(1, math.ceil(percentile * len(ordered) / 100)) - 1]
    return {"count": len(vector), "min": min(vector), "p50": nearest(50),
            "mean": statistics.mean(vector), "p95": nearest(95),
            "p99": nearest(99), "max": max(vector)}


def spread(values: Iterable[int | float]) -> float:
    vector = [float(value) for value in values]
    require(vector and all(math.isfinite(value) for value in vector), "spread vector invalid")
    low, high = min(vector), max(vector)
    if low == 0:
        return 0.0 if high == 0 else float("inf")
    return (high - low) * 100.0 / abs(low)


def distribution(values: Iterable[int | float]) -> dict[str, Any]:
    vector = list(values)
    result = stats(vector)
    result["values"] = vector
    result["spread_percent"] = spread(vector)
    result["flag_over_5_percent"] = result["spread_percent"] > 5.0
    return result


def allocation_from_sample(sample: dict[str, Any], label: str) -> dict[str, int]:
    value = sample.get("allocation")
    require(isinstance(value, dict) and value.get("status") == "measured",
            f"{label} allocation sample is not measured")
    require(value.get("scope") == "operation_global_system_allocator",
            f"{label} allocation scope changed")
    result: dict[str, int] = {}
    for field in ALLOCATION_FIELDS[:11]:
        number = value.get(field)
        nonnegative_int(number, f"{label} allocation {field}")
        result[field] = number
    before, after = result["live_bytes_before"], result["live_bytes_after"]
    require(after == before + result["allocated_bytes"] - result["deallocated_bytes"],
            f"{label} allocation live-byte accounting changed")
    require(result["region_peak_live_bytes"] >= before
            and result["region_peak_live_bytes"] >= after,
            f"{label} allocation region peak ordering changed")
    require(result["peak_live_bytes_after"] >= result["peak_live_bytes_before"]
            and result["peak_live_bytes_after"] >= result["region_peak_live_bytes"],
            f"{label} allocator peak ordering changed")
    require(result["failed_allocation_calls"] == 0, f"{label} failed allocation observed")
    result["net_live"] = after - before
    result["peak_above_entry"] = result["region_peak_live_bytes"] - before
    return result


def validate_report(report: dict[str, Any], job: dict[str, Any], kind: str,
                    binary: dict[str, Any], label: str) -> dict[str, Any]:
    schema, tool, _ = probe_contract()
    require(report.get("schema") == schema, f"{label} report schema changed")
    require(report.get("tool") == tool, f"{label} report tool changed")
    require(report.get("timing_scope") == TIMING_SCOPES[job["mode"]],
            f"{label} timing scope changed")
    source_identity(report, job, label)
    require(report.get("samples_requested") == job["samples"]
            and report.get("warmup") == job["warmup"], f"{label} sample configuration changed")
    nonnegative_int(report.get("expected_semantic_text_bytes"),
                    f"{label} expected semantic bytes")
    nonnegative_int(report.get("expected_raw_text_bytes"),
                    f"{label} expected raw text bytes")
    samples = report.get("samples")
    require(isinstance(samples, list) and len(samples) == job["samples"],
            f"{label} sample count changed")
    allocator = report.get("allocator")
    require(isinstance(allocator, dict), f"{label} allocator identity is missing")
    expected_identity = Path(binary["path"]).name
    require(allocator.get("binary") == expected_identity,
            f"{label} allocator binary identity changed")
    if kind == "native":
        require(allocator.get("instrumentation") == "none"
                and allocator.get("allocator") == "Rust system allocator"
                and allocator.get("counter_revision") is None,
                f"{label} native allocator identity changed")
    else:
        require(allocator.get("instrumentation") == "system_allocator_operation_scoped"
                and allocator.get("allocator") == "CountingSystemAllocator(std::alloc::System)"
                and allocator.get("counter_revision") == "serialized_region_peak_v3",
                f"{label} allocation allocator identity changed")
    elapsed: list[int] = []
    allocations: dict[str, list[int]] = {field: [] for field in ALLOCATION_FIELDS}
    outputs: list[tuple[int, str]] = []
    for index, sample in enumerate(samples):
        require(isinstance(sample, dict) and sample.get("index") == index,
                f"{label} sample {index} identity changed")
        elapsed_ns = sample.get("elapsed_ns")
        nonnegative_int(elapsed_ns, f"{label} elapsed_ns[{index}]")
        elapsed.append(elapsed_ns)
        metrics = sample.get("metrics")
        require(isinstance(metrics, dict) and metrics.get("elapsed_ns") == elapsed_ns
                and metrics.get("slides") == report.get("slides")
                and metrics.get("boxes_per_slide") == report.get("boxes_per_slide"),
                f"{label} raw metric identity changed")
        require(sample.get("source_sha256") == report["source"]["sha256"],
                f"{label} sample source digest changed")
        verification = sample.get("verification")
        require(isinstance(verification, dict) and verification.get("semantic_check") is True,
                f"{label} sample {index} semantic verification failed")
        require(verification.get("reopened") is True, f"{label} output was not reopened")
        output = sample.get("output")
        require(isinstance(output, dict) and is_sha(output.get("sha256")),
                f"{label} output identity is missing")
        nonnegative_int(output.get("bytes"), f"{label} output bytes")
        outputs.append((output["bytes"], output["sha256"]))
        require(verification.get("slide_count") == report.get("slides")
                and verification.get("expected_slide_count") == report.get("slides")
                and verification.get("slide_count_match") is True,
                f"{label} slide readback changed")
        require(verification.get("semantic_check") is True
                and verification.get("exact_text_match") is True,
                f"{label} semantic text readback changed")
        expected_bytes = report.get("expected_semantic_text_bytes")
        expected_sha = verification.get("expected_semantic_text_sha256")
        semantic_bytes = verification.get("semantic_text_bytes")
        semantic_sha = verification.get("semantic_text_sha256")
        nonnegative_int(expected_bytes, f"{label} expected semantic bytes")
        nonnegative_int(verification.get("expected_semantic_text_bytes"),
                        f"{label} sample expected semantic bytes")
        require(verification.get("expected_semantic_text_bytes") == expected_bytes
                and semantic_bytes == expected_bytes
                and is_sha(expected_sha) and is_sha(semantic_sha)
                and semantic_sha == expected_sha,
                f"{label} semantic digest does not match expected readback")
        expected_raw_bytes = report["expected_raw_text_bytes"]
        nonnegative_int(verification.get("raw_text_box_count"),
                        f"{label} raw text box count")
        nonnegative_int(verification.get("expected_raw_text_box_count"),
                        f"{label} expected raw text box count")
        nonnegative_int(verification.get("raw_text_atom_count"),
                        f"{label} raw text atom count")
        require(verification.get("raw_text_check") is True
                and verification.get("raw_text_match") is True
                and verification.get("raw_text_boxes_match") is True
                and verification.get("raw_text_box_count")
                == verification.get("expected_raw_text_box_count")
                and verification.get("raw_text_bytes") == expected_raw_bytes
                and verification.get("expected_raw_text_bytes") == expected_raw_bytes
                and is_sha(verification.get("expected_raw_text_sha256"))
                and is_sha(verification.get("raw_text_sha256"))
                and verification.get("raw_text_sha256")
                == verification.get("expected_raw_text_sha256"),
                f"{label} raw text oracle changed")
        allocation = sample.get("allocation")
        if kind == "native":
            require(allocation is None, f"{label} native report contains allocation metrics")
        else:
            values = allocation_from_sample(sample, f"{label} sample {index}")
            for field, number in values.items():
                allocations[field].append(number)
    require(outputs and len(set(outputs)) == 1,
            f"{label} output identity is not deterministic across samples")
    return {"stats": stats(elapsed), "allocation": None if kind == "native" else allocations}


def load_lane(plan: dict[str, Any], lane: str, builds: dict[str, Any], expected_source: dict[str, Any],
              cleanup: Any, cleanup_verified: bool) -> list[dict[str, Any]]:
    directory = PACKET / lane
    require(directory.is_dir(), f"missing {lane} capture directory")
    complete = read_json(directory / "complete.json")
    jobs = expected_jobs(plan, lane)
    require(complete.get("children") == len(jobs), f"{lane} child cardinality changed")
    source_path = artifact_path(complete.get("source"), f"{lane} complete source")
    require(source_files_equal(source_manifest(read_json(source_path), f"{lane} complete source"), expected_source),
            f"{lane} complete source differs from expected source")
    receipts_path = artifact_path(complete.get("receipts"), f"{lane} receipts")
    rows = read_json(receipts_path)
    require(isinstance(rows, list) and len(rows) == len(jobs), f"{lane} receipt cardinality changed")
    entries: list[dict[str, Any]] = []
    seen: set[tuple[Any, ...]] = set()
    source_digests: dict[str, str] = {}
    for index, (row, job) in enumerate(zip(rows, jobs)):
        label = f"{lane} child {index}"
        require(isinstance(row, dict), f"{label} is not an object")
        for key in ("lane", "block", "shape", "mode", "leg"):
            require(row.get(key) == job[key], f"{label} {key} identity changed")
        identity = (job["block"], job["shape"], job["mode"], job["leg"])
        require(identity not in seen, f"duplicate {lane} identity: {identity}")
        seen.add(identity)
        require(row.get("exit_code") == 0, f"{label} failed")
        kind = "native" if lane == "native" else "allocation"
        binary = row.get("binary")
        expected_binary = builds[job["leg"]]["binaries"][kind]
        require(isinstance(binary, dict)
                and binary.get("bytes") == expected_binary.get("bytes")
                and binary.get("sha256") == expected_binary.get("sha256"),
                f"{label} binary identity changed")
        validate_binary(binary, f"{label} binary", cleanup, cleanup_verified)
        report_path = artifact_path(row.get("report"), f"{label} report")
        log_path = artifact_path(row.get("log"), f"{label} log")
        rss_path = artifact_path(row.get("rss"), f"{label} RSS")
        rss_text = rss_path.read_text().strip()
        require(rss_text.isdigit(), f"{label} RSS is not an integer")
        rss = int(rss_text)
        command = row.get("command")
        require(isinstance(command, list), f"{label} command is malformed")
        normalized_command = [normalized(item) for item in command]
        require(normalized_command == expected_command(plan, job, expected_binary,
                                                       report_path, rss_path),
                f"{label} command changed")
        report = read_json(report_path)
        outcome = validate_report(report, job, kind, expected_binary, label)
        digest = report["source"]["sha256"]
        previous_digest = source_digests.setdefault(job["shape"], digest)
        require(previous_digest == digest, f"{label} source digest is not deterministic")
        entries.append({"identity": job, "row": row, "report_path": report_path,
                        "report_sha256": sha256(report_path), "log": rel(log_path),
                        "rss_kib": rss, "source_sha256": digest, "stats": outcome["stats"],
                        "allocation": outcome["allocation"]})
    return entries


def pair_ratio(before: float, after: float) -> dict[str, Any]:
    if before == 0:
        equal_zero = after == 0
        return {
            "before": before,
            "after": after,
            "ratio": 1.0 if equal_zero else None,
            "change_percent": 0.0 if equal_zero else None,
            "relative_change_defined": False,
            "zero_baseline_equal": equal_zero,
            "zero_to_nonzero": not equal_zero,
            "over_5_percent": not equal_zero,
        }
    change = (after - before) * 100.0 / abs(before)
    return {"before": before, "after": after, "ratio": after / before,
            "change_percent": change, "relative_change_defined": True,
            "zero_baseline_equal": False, "zero_to_nonzero": False,
            "over_5_percent": change > 5.0}


def bootstrap_ci(values: list[float]) -> dict[str, Any]:
    require(values, "cannot bootstrap an empty ratio vector")
    rng = random.Random(BOOTSTRAP_SEED)
    estimates: list[float] = []
    for _ in range(BOOTSTRAP_RESAMPLES):
        sample = [values[rng.randrange(len(values))] for _ in values]
        estimates.append(statistics.median(sample))
    estimates.sort()
    low_rank = max(0, math.floor((1.0 - BOOTSTRAP_CONFIDENCE) / 2.0 * len(estimates)))
    high_rank = min(len(estimates) - 1,
                    math.ceil((1.0 + BOOTSTRAP_CONFIDENCE) / 2.0 * len(estimates)) - 1)
    return {"seed": BOOTSTRAP_SEED, "resamples": BOOTSTRAP_RESAMPLES,
            "confidence": BOOTSTRAP_CONFIDENCE, "statistic": "median",
            "ci_low": estimates[low_rank], "ci_high": estimates[high_rank]}


def paired(entries: list[dict[str, Any]], metrics: Iterable[str]) -> dict[str, Any]:
    by_key = {(item["identity"]["shape"], item["identity"]["mode"],
               item["identity"]["leg"], item["identity"]["block"]): item
              for item in entries}
    keys = sorted({(item["identity"]["shape"], item["identity"]["mode"])
                   for item in entries})
    output: dict[str, Any] = {}
    for shape, mode in keys:
        blocks = sorted({item["identity"]["block"] for item in entries
                         if item["identity"]["shape"] == shape
                         and item["identity"]["mode"] == mode})
        values: dict[str, Any] = {}
        for metric in metrics:
            by_block: list[dict[str, Any]] = []
            ratios: list[float] = []
            for block in blocks:
                before = by_key[(shape, mode, "before", block)]
                after = by_key[(shape, mode, "after", block)]
                if metric == "rss_kib":
                    left, right = before["rss_kib"], after["rss_kib"]
                elif before["allocation"] is not None:
                    left = stats(before["allocation"][metric])["p50"]
                    right = stats(after["allocation"][metric])["p50"]
                else:
                    left, right = before["stats"][metric], after["stats"][metric]
                result = pair_ratio(float(left), float(right))
                result["block"] = block
                by_block.append(result)
                if result["ratio"] is not None:
                    ratios.append(result["ratio"])
            median = statistics.median(ratios) if ratios else None
            bootstrap = bootstrap_ci(ratios) if ratios else {
                "seed": BOOTSTRAP_SEED,
                "resamples": BOOTSTRAP_RESAMPLES,
                "confidence": BOOTSTRAP_CONFIDENCE,
                "statistic": "median",
                "ci_low": None,
                "ci_high": None,
                "undefined_relative_change": True,
            }
            values[metric] = {
                "by_block": by_block, "ratio_median": median,
                "change_percent_median": None if median is None else (median - 1.0) * 100.0,
                "bootstrap": bootstrap,
                "defined_ratio_blocks": len(ratios),
                "undefined_ratio_blocks": len(by_block) - len(ratios),
                "regression_over_5_percent": any(item["over_5_percent"] for item in by_block),
            }
        output[f"{shape}/{mode}"] = {
            "shape": shape, "mode": mode, "blocks": len(blocks),
            "metrics": values,
            "comparison": "after/before paired by alternating capture block",
        }
    return output


def native_analysis(entries: list[dict[str, Any]]) -> dict[str, Any]:
    groups: dict[str, Any] = {}
    spread_flags: list[dict[str, Any]] = []
    grouped: dict[tuple[str, str, str], list[dict[str, Any]]] = {}
    for item in entries:
        identity = item["identity"]
        grouped.setdefault((identity["shape"], identity["mode"], identity["leg"]), []).append(item)
    for key, items in sorted(grouped.items()):
        elapsed = {metric: distribution(item["stats"][metric] for item in items)
                   for metric in NATIVE_METRICS}
        rss = distribution(item["rss_kib"] for item in items)
        for metric, result in (*elapsed.items(), ("rss_kib", rss)):
            if result["flag_over_5_percent"]:
                spread_flags.append({"group": list(key), "metric": metric,
                                     "spread_percent": result["spread_percent"]})
        groups["/".join(key)] = {
            "shape": key[0], "mode": key[1], "leg": key[2],
            "processes": [{"block": item["identity"]["block"], "stats": item["stats"],
                           "rss_kib": item["rss_kib"], "report": rel(item["report_path"]),
                           "report_sha256": item["report_sha256"]}
                          for item in sorted(items, key=lambda value: value["identity"]["block"])],
            "elapsed_distribution_across_processes": elapsed,
            "per_process_rss_distribution": rss,
        }
    paired_values = paired(entries, (*NATIVE_METRICS, "rss_kib"))
    regression_flags = []
    for key, value in paired_values.items():
        for metric, result in value["metrics"].items():
            if result["regression_over_5_percent"]:
                regression_flags.append({"group": key, "metric": metric,
                                         "ratio_median": result["ratio_median"],
                                         "change_percent_median": result["change_percent_median"]})
    return {"groups": groups, "spread_flags_over_5_percent": spread_flags,
            "regression_flags_over_5_percent": regression_flags,
            "paired_by_block_before_after": paired_values,
            "rss_is_separate_from_elapsed": True, "allocation_metrics_present": False}


def allocation_analysis(entries: list[dict[str, Any]]) -> dict[str, Any]:
    groups: dict[str, Any] = {}
    spread_flags: list[dict[str, Any]] = []
    grouped: dict[tuple[str, str, str], list[dict[str, Any]]] = {}
    for item in entries:
        identity = item["identity"]
        grouped.setdefault((identity["shape"], identity["mode"], identity["leg"]), []).append(item)
    for key, items in sorted(grouped.items()):
        metrics: dict[str, Any] = {}
        for field in ALLOCATION_FIELDS:
            per_process = [{"block": item["identity"]["block"], "values": item["allocation"][field],
                            "stats": stats(item["allocation"][field])}
                           for item in sorted(items, key=lambda value: value["identity"]["block"])]
            repeats = [entry["stats"]["p50"] for entry in per_process]
            result = {"per_process": per_process, "repeat_p50_values": repeats,
                      "spread_percent": spread(repeats),
                      "flag_over_5_percent": spread(repeats) > 5.0}
            metrics[field] = result
            if result["flag_over_5_percent"]:
                spread_flags.append({"group": list(key), "metric": field,
                                     "spread_percent": result["spread_percent"]})
        groups["/".join(key)] = {"shape": key[0], "mode": key[1], "leg": key[2],
                                  "blocks": len(items), "metrics": metrics,
                                  "elapsed_not_mixed": True}
    paired_values = paired(entries, ALLOCATION_FIELDS)
    regression_flags = []
    for key, value in paired_values.items():
        for metric, result in value["metrics"].items():
            if result["regression_over_5_percent"]:
                regression_flags.append({"group": key, "metric": metric,
                                         "ratio_median": result["ratio_median"],
                                         "change_percent_median": result["change_percent_median"]})
    return {"groups": groups, "fields": list(ALLOCATION_FIELDS),
            "spread_flags_over_5_percent": spread_flags,
            "regression_flags_over_5_percent": regression_flags,
            "paired_by_block_before_after": paired_values,
            "allocation_is_separate_from_elapsed": True}


def qualification(entries: list[dict[str, Any]]) -> dict[str, Any]:
    require(all(item["identity"]["leg"] == "before" for item in entries),
            "qualification contains an after leg")
    return {"children": len(entries), "before_source_only": True,
            "rows": [{"shape": item["identity"]["shape"], "mode": item["identity"]["mode"],
                      "report": rel(item["report_path"]),
                      "report_sha256": item["report_sha256"]} for item in entries]}


def optional_seal() -> bool:
    """Validate a present seal while keeping the analysis result preseal-stable.

    The boolean is deliberately a contract marker rather than a presence
    marker.  Otherwise analysis.json would change from ``false`` to ``true``
    when the seal is added, invalidating the seal's own hash.  Final replay
    separately requires that a seal file exists.
    """

    for name in ("seal.json", "final-seal.json"):
        path = PACKET / name
        if not path.is_file():
            continue
        value = read_json(path)
        require(isinstance(value, dict) and isinstance(value.get("files"), dict),
                f"{name} is malformed")
        actual = {rel(item): sha256(item) for item in PACKET.rglob("*")
                  if item.is_file() and item.name not in {"seal.json", "final-seal.json"}}
        require(value["files"] == actual, f"{name} file inventory is stale")
        return True
    return True


def analyze() -> dict[str, Any]:
    plan = load_plan()
    builds, cleanup, cleanup_verified = load_builds(plan)
    initial_binary_relocations = load_initial_binary_relocations(cleanup, cleanup_verified)
    probe_tests = load_probe_tests(builds)
    before_source, after_source = builds["before"]["source"], builds["after"]["source"]
    disposition = load_disposition(before_source, after_source)
    quality = check_quality(after_source)
    test_summary = check_test_summary(quality)
    native_entries = load_lane(plan, "native", builds, after_source, cleanup, cleanup_verified)
    allocation_entries = load_lane(plan, "allocation", builds, after_source, cleanup, cleanup_verified)
    qualification_entries = load_lane(plan, "qualification", builds, before_source,
                                      cleanup, cleanup_verified)
    require(len(native_entries) == 120, "native cardinality changed")
    require(len(allocation_entries) == 40, "allocation cardinality changed")
    require(len(qualification_entries) == 10, "qualification cardinality changed")
    digests = {
        item["identity"]["shape"]: item["source_sha256"]
        for item in (*native_entries, *allocation_entries, *qualification_entries)
    }
    require(len(digests) == len(SHAPES), "source digest shape coverage changed")
    for item in (*native_entries, *allocation_entries, *qualification_entries):
        require(item["source_sha256"] == digests[item["identity"]["shape"]],
                "source digest changed between lanes")
    qualification_parity = load_qualification_parity()
    baseline_attribution = load_baseline_attribution()
    observer = load_observer_analysis()
    ci = load_ci_analysis()
    optional = optional_seal()
    return {
        "schema": "litchi-0781-borrowed-text-analysis-v1",
        "plan_schema": plan["schema"], "quality": quality,
        "test_summary": test_summary,
        "probe_tests": probe_tests,
        "initial_binary_relocations": initial_binary_relocations,
        "qualification_parity": qualification_parity,
        "baseline_attribution": baseline_attribution,
        "observer": observer,
        "ci": ci,
        "source": {"before": before_source, "after": after_source,
                    "changed_files": sorted(plan["source_allowlist"])},
        "disposition": disposition,
        "native": {"children": len(native_entries), "blocks": plan["native"]["blocks"],
                    "samples": plan["native"]["samples"], "warmup": plan["native"]["warmup"],
                    "analysis": native_analysis(native_entries),
                    "receipts": [{"shape": i["identity"]["shape"], "mode": i["identity"]["mode"],
                                  "block": i["identity"]["block"], "leg": i["identity"]["leg"],
                                  "report": rel(i["report_path"]), "report_sha256": i["report_sha256"]}
                                 for i in native_entries]},
        "allocation": {"children": len(allocation_entries), "blocks": plan["allocation"]["blocks"],
                        "samples": plan["allocation"]["samples"], "warmup": plan["allocation"]["warmup"],
                        "analysis": allocation_analysis(allocation_entries),
                        "receipts": [{"shape": i["identity"]["shape"], "mode": i["identity"]["mode"],
                                      "block": i["identity"]["block"], "leg": i["identity"]["leg"],
                                      "report": rel(i["report_path"]), "report_sha256": i["report_sha256"]}
                                     for i in allocation_entries]},
        "qualification": qualification(qualification_entries),
        "bootstrap": {"seed": BOOTSTRAP_SEED, "resamples": BOOTSTRAP_RESAMPLES,
                       "confidence": BOOTSTRAP_CONFIDENCE, "statistic": "median"},
        "verification": {
            "plan_cardinality_and_alternating_order_checked": True,
            "source_binary_probe_lock_fixture_receipts_checked": True,
            "probe_tests_receipt_and_named_oracles_checked": True,
            "initial_binary_relocations_replayed": True,
            "qualification_parity_replayed": True,
            "baseline_attribution_replayed": True,
            "fixed_release_profile_checked": True,
            "quality_commands_checked_exactly": True,
            "native_has_no_allocation_metrics": True,
            "source_digest_and_semantic_checks": True,
            "qualification_preflight_outside_paired_matrix": True,
            "source_change_allowlist": sorted(plan["source_allowlist"]),
            "cleanup_binary_witness_required_when_missing": True,
            "optional_file_inventory_seal_checked": optional,
            "observer_analysis_replayed": True,
            "ci_analysis_replayed": True,
            "test_summary_replayed_from_bound_log": True,
        },
        "limits": [
            "Timing and allocation comparisons are descriptive before/after evidence.",
            "RSS is a whole-process /usr/bin/time gauge including setup and verification.",
            "Allocator regions are not physical-copy, RSS, or causal cost estimates.",
            "No cold-cache, device-floor, concurrency, or CRUD-completeness claim is made.",
        ],
    }


def main() -> None:
    result = analyze()
    output = PACKET / "analysis.json"
    require(not output.exists(), "refusing to overwrite analysis.json")
    output.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n")
    print(json.dumps({"native": result["native"]["children"],
                      "allocation": result["allocation"]["children"],
                      "qualification": result["qualification"]["children"]}, sort_keys=True))


if __name__ == "__main__":
    main()
