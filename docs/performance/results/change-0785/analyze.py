"""Offline replay and analysis for the 0785 known-URI evidence packet.

The capture program is deliberately outside this module.  This file only
checks retained custody receipts and derives deterministic summaries from the
JSON reports and ``/usr/bin/time`` RSS gauges.  It is safe to run after the
owned worktree and build target have been removed: every missing executable
must then be covered by an exact cleanup receipt.
"""

from __future__ import annotations

import hashlib
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
# The plan may refine this list once the candidate is frozen.  Keeping a
# narrow fallback here prevents an unfinished plan from silently becoming a
# broad "anything in litchi-pptx" allowance.
SOURCE_ALLOWLIST = (
    "crates/litchi-pptx/src/notes/mod.rs",
)
# Kept as a compatibility alias for the single-file disposition shape.  The
# source-census gate below uses the complete explicit allowlist from plan.json.
SOURCE_FILE = SOURCE_ALLOWLIST[0]
LEGS = ("before", "after")
CASES = (
    {"shape": "tiny", "mode": "capture"},
    {"shape": "tiny", "mode": "commit"},
    {"shape": "tiny", "mode": "lifecycle"},
    {"shape": "medium", "mode": "capture"},
    {"shape": "medium", "mode": "commit"},
    {"shape": "medium", "mode": "lifecycle"},
    {"shape": "large", "mode": "capture"},
    {"shape": "large", "mode": "commit"},
    {"shape": "large", "mode": "lifecycle"},
    {"shape": "vendor", "mode": "capture"},
    {"shape": "vendor", "mode": "commit"},
    {"shape": "vendor", "mode": "lifecycle"},
    {"shape": "unicode-vendor", "mode": "capture"},
    {"shape": "unicode-vendor", "mode": "commit"},
    {"shape": "unicode-vendor", "mode": "lifecycle"},
)
MODES = tuple(sorted({case["mode"] for case in CASES}))
SHAPES = ("tiny", "medium", "large", "vendor", "unicode-vendor")
SHAPE_DIMENSIONS = {
    "tiny": (3, 4),
    "medium": (12, 8),
    "large": (100, 100),
    "vendor": (12, 8),
    "unicode-vendor": (12, 8),
}
ORDINARY_CASES = tuple(
    {"shape": shape, "mode": mode}
    for shape in ("tiny", "medium", "large")
    for mode in ("capture", "commit", "lifecycle")
)
HISTORICAL_PACKET = "docs/performance/results/change-0780"
PRODUCTION_PATHSPEC = (
    "crates",
    "Cargo.toml",
    "clippy.toml",
    ".cargo/config.toml",
    "rust-toolchain.toml",
)
FROZEN_INPUT_FILES = (
    "adoption-policy.json",
    "architecture-inputs.json",
    "build.py",
    "capture.py",
    "plan.json",
)
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
BOOTSTRAP_SEED = 785_078
BOOTSTRAP_RESAMPLES = 10_000
BOOTSTRAP_CONFIDENCE = 0.95
REPORT_SCHEMA = "litchi.pptx.namespace-uri-probe.v1"
REPORT_TOOL = "namespace-uri-probe-0785"
MARKER = "litchi-perf-0780-static-mce-capabilities"
TIMING_SCOPES = {
    "capture": "Package::opened_presentation only",
    "commit": "Transaction::commit only; package capture and one set_shape_text staging are outside the clock",
    "lifecycle": "Package::opened_presentation, edit, set_shape_text, commit, apply_opened_presentation_commit, and Package::to_bytes",
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


def is_git_revision(value: Any) -> bool:
    return isinstance(value, str) and len(value) == 40 and all(c in HEX for c in value)


def digest_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


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
    for marker in ("change-0785", "build-before", "build-after", "native",
                   "allocation", "qualification"):
        if marker in parts:
            index = parts.index(marker)
            if marker == "change-0785":
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
        prefix = "docs/performance/results/change-0785/"
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
    require(plan.get("schema") == "litchi.performance.0785.v1", "plan schema changed")
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
    contract = plan.get("probe")
    if contract is None:
        contract = {"schema": REPORT_SCHEMA, "tool": REPORT_TOOL, "marker": MARKER}
        plan["probe"] = contract
    require(isinstance(contract, dict), "probe contract is malformed")
    require(contract.get("schema") == REPORT_SCHEMA
            and contract.get("tool") == REPORT_TOOL,
            "probe schema or tool changed")
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


def _git_output(arguments: list[str], label: str) -> str:
    try:
        return subprocess.check_output(["git", *arguments], cwd=ROOT, text=True)
    except (OSError, subprocess.CalledProcessError) as error:
        fail(f"cannot read {label}: {error}")


def _git_blob_bytes(revision: str, names: Iterable[str], label: str) -> dict[str, bytes]:
    """Read immutable Git blobs in one batch and return their exact bytes."""

    ordered = list(names)
    require(is_git_revision(revision), f"{label} revision is invalid")
    require(len(set(ordered)) == len(ordered), f"{label} contains duplicate paths")
    require(all(name and "\n" not in name and "\0" not in name for name in ordered),
            f"{label} contains an invalid path")
    request = "".join(f"{revision}:{name}\n" for name in ordered).encode()
    try:
        completed = subprocess.run(
            ["git", "cat-file", "--batch"], cwd=ROOT, input=request,
            stdout=subprocess.PIPE, check=True,
        )
    except (OSError, subprocess.CalledProcessError) as error:
        fail(f"cannot read {label} Git blobs: {error}")
    data = completed.stdout
    offset = 0
    result: dict[str, bytes] = {}
    for name in ordered:
        end = data.find(b"\n", offset)
        require(end >= 0, f"{label} Git blob header is truncated: {name}")
        header = data[offset:end].split()
        require(len(header) == 3 and is_git_revision(header[0].decode(errors="replace"))
                and header[1] == b"blob", f"{label} Git blob is not a blob: {name}")
        try:
            size = int(header[2])
        except ValueError:
            fail(f"{label} Git blob size is invalid: {name}")
        offset = end + 1
        require(size >= 0 and offset + size <= len(data),
                f"{label} Git blob is truncated: {name}")
        result[name] = data[offset:offset + size]
        offset += size
        require(data[offset:offset + 1] == b"\n",
                f"{label} Git blob terminator is missing: {name}")
        offset += 1
    require(offset == len(data), f"{label} Git blob batch has trailing data")
    return result


def _git_blob_digests(revision: str, names: Iterable[str], label: str) -> dict[str, str]:
    return {name: digest_bytes(value)
            for name, value in _git_blob_bytes(revision, names, label).items()}


def load_revision_transition(builds: dict[str, Any]) -> dict[str, Any]:
    """Prove that the successful baseline build was a docs-only descendant."""

    path = PACKET / "revision-transition.json"
    value = read_json(path)
    require(isinstance(value, dict), "revision transition is malformed")
    base = origin().get("base")
    require(value.get("base") == base and is_git_revision(base),
            "revision transition base changed")
    build_revision = value.get("build_revision")
    require(is_git_revision(build_revision), "revision transition build revision is invalid")
    require(builds["before"]["source"]["revision"] == build_revision,
            "before census revision differs from recorded build revision")
    commits = value.get("commits")
    require(isinstance(commits, list) and commits, "revision transition commit custody is incomplete")
    commit_ids: list[str] = []
    for commit in commits:
        require(isinstance(commit, str) and commit, "revision transition commit is malformed")
        commit_id = commit.split(maxsplit=1)[0]
        require(is_git_revision(commit_id), "revision transition commit ID is invalid")
        commit_ids.append(commit_id)
    require(build_revision in commit_ids, "revision transition commit custody is incomplete")
    require(value.get("production_diff_empty") is True,
            "revision transition production-diff marker changed")
    _git_output(["merge-base", "--is-ancestor", base, build_revision],
                "revision transition ancestry")
    changed = _git_output(
        ["diff", "--name-only", f"{base}..{build_revision}", "--", *PRODUCTION_PATHSPEC],
        "production revision diff",
    ).splitlines()
    require(changed == [], f"docs-only revision changed production files: {changed}")
    expected_files = builds["before"]["source"]["files"]
    base_files = _git_output(
        ["ls-tree", "-r", "--name-only", base, "--", *PRODUCTION_PATHSPEC],
        "base production census",
    ).splitlines()
    require(sorted(base_files) == sorted(expected_files),
            "base production file census differs from before census")
    base_digests = _git_blob_digests(base, base_files, "base production census")
    require(base_digests == expected_files, "base production blob hashes differ from before census")
    return {
        "receipt": _file_identity(path),
        "base": base,
        "build_revision": build_revision,
        "commits": list(commits),
        "commit_ids": commit_ids,
        "ancestor_checked": True,
        "production_diff_empty": True,
        "base_file_census_matches": True,
        "base_blob_hashes_checked": len(base_digests),
    }


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


def load_frozen_inputs(directory: Path, label: str) -> dict[str, str]:
    path = directory / "frozen-inputs.json"
    value = read_json(path)
    require(isinstance(value, dict) and set(value) == set(FROZEN_INPUT_FILES),
            f"{label} frozen input set changed")
    result: dict[str, str] = {}
    for name in FROZEN_INPUT_FILES:
        digest = value.get(name)
        require(is_sha(digest), f"{label} frozen input digest is invalid: {name}")
        input_path = PACKET / name
        require(input_path.is_file() and not input_path.is_symlink(),
                f"{label} frozen input is missing: {name}")
        require(sha256(input_path) == digest,
                f"{label} frozen input changed: {name}")
        result[name] = digest
    return result


def load_builds(plan: dict[str, Any]) -> tuple[dict[str, Any], Any, bool]:
    cleanup, cleanup_verified = load_cleanup()
    fixed_release_profile()
    builds: dict[str, Any] = {}
    for leg in LEGS:
        directory = PACKET / f"build-{leg}"
        build = read_json(directory / "build.json")
        require(isinstance(build, dict), f"{leg} build manifest is malformed")
        frozen_inputs = load_frozen_inputs(directory, leg)
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
                       "lock": lock, "binaries": binaries,
                       "frozen_inputs": frozen_inputs}
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
    require(builds["before"]["frozen_inputs"] == builds["after"]["frozen_inputs"],
            "frozen build inputs changed between legs")
    for lane in ("native", "allocation", "qualification"):
        path = PACKET / lane / "source.json"
        if path.is_file():
            expected = builds["before"]["source"] if lane == "qualification" else builds["after"]["source"]
            require(source_files_equal(source_manifest(read_json(path), f"{lane} source"), expected),
                    f"{lane} source census differs from expected build")
    return builds, cleanup, cleanup_verified


def load_architecture_inputs() -> dict[str, Any]:
    """Bind the 35 architecture references to both live files and origin Git blobs."""

    path = PACKET / "architecture-inputs.json"
    value = read_json(path)
    require(isinstance(value, dict) and len(value) == 35,
            "architecture input cardinality changed")
    files: dict[str, str] = {}
    for name, digest in value.items():
        require(isinstance(name, str) and name and not name.startswith("/"),
                "architecture input path is invalid")
        require(is_sha(digest), f"architecture input digest is invalid: {name}")
        live = ROOT / name
        require(live.is_file() and not live.is_symlink(),
                f"architecture input is missing: {name}")
        require(sha256(live) == digest, f"architecture input changed in live files: {name}")
        files[name] = digest
    base = origin().get("base")
    require(is_git_revision(base), "origin base revision is invalid")
    git_digests = _git_blob_digests(base, files, "architecture inputs")
    require(git_digests == files, "architecture input origin Git blobs changed")
    return {
        "receipt": _file_identity(path),
        "revision": base,
        "count": len(files),
        "files": files,
        "live_files_match": True,
        "origin_blob_hashes_match": True,
    }


QUALITY_COMMANDS = (
    ["cargo", "fmt", "-p", "litchi-pptx", "--", "--check"],
    ["cargo", "check", "--offline", "--locked", "-p", "litchi-pptx",
     "--all-features", "--all-targets"],
    ["cargo", "test", "--offline", "--locked", "-p", "litchi-pptx",
     "--all-features", "--", "--test-threads=2"],
    ["cargo", "clippy", "--offline", "--locked", "-p", "litchi-pptx",
     "--all-features", "--lib", "--", "-D", "warnings"],
    ["cargo", "doc", "--offline", "--locked", "-p", "litchi-pptx",
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


def load_baseline_fixture_parity() -> dict[str, Any]:
    """Bind ordinary qualification fixtures to the sealed 0780 oracle.

    This is fixture provenance only.  The previous paths are retained as
    identities in the packet and are never used as timing or capture input.
    """

    parity_path = PACKET / "baseline-fixture-parity.json"
    parity = read_json(parity_path)
    require(isinstance(parity, dict), "baseline fixture parity is malformed")
    scope = parity.get("scope")
    require(isinstance(scope, str) and "0780" in scope
            and "parity" in scope.lower() and "timing" in scope.lower(),
            "baseline fixture parity scope changed")
    rows = parity.get("rows")
    require(isinstance(rows, list) and len(rows) == len(ORDINARY_CASES),
            "baseline fixture parity row cardinality changed")
    expected_cases = {f"0-{case['shape']}-{case['mode']}-before" for case in ORDINARY_CASES}
    seen: set[str] = set()
    result_rows: list[dict[str, Any]] = []
    for index, row in enumerate(rows):
        require(isinstance(row, dict), f"baseline fixture parity row {index} is malformed")
        case = row.get("case")
        require(isinstance(case, str) and case in expected_cases and case not in seen,
                f"baseline fixture parity case changed: {case}")
        seen.add(case)
        current = row.get("current")
        previous = row.get("previous")
        source = row.get("source")
        output = row.get("output")
        for identity, label in ((current, "current"), (previous, "previous"),
                                (source, "source"), (output, "output")):
            require(isinstance(identity, dict),
                    f"baseline fixture parity {case} {label} identity is missing")
            require(is_sha(identity.get("sha256")),
                    f"baseline fixture parity {case} {label} digest is invalid")
            nonnegative_int(identity.get("bytes"),
                           f"baseline fixture parity {case} {label} bytes")
        current_path = resolve_path(current["path"])
        require(current_path == (PACKET / "qualification" / f"{case}.json").resolve(),
                f"baseline fixture parity {case} current path changed")
        require(current["bytes"] == current_path.stat().st_size
                and current["sha256"] == sha256(current_path),
                f"baseline fixture parity {case} current receipt changed")
        current_identity = _qualification_report_identity(
            current_path, f"baseline fixture parity {case} current")
        require(current_identity["source"] == {
            "bytes": source["bytes"], "sha256": source["sha256"]},
                f"baseline fixture parity {case} source changed")
        require(current_identity["output"] == {
            "bytes": output["bytes"], "sha256": output["sha256"]},
                f"baseline fixture parity {case} output changed")
        previous_path = previous.get("path")
        require(isinstance(previous_path, str)
                and "/change-0780/qualification/" in previous_path.replace("\\", "/")
                and Path(previous_path).name == case + ".json",
                f"baseline fixture parity {case} previous path changed")
        previous_file = ROOT / HISTORICAL_PACKET / "qualification" / f"{case}.json"
        require(previous_file.is_file()
                and previous_file.stat().st_size == previous["bytes"]
                and sha256(previous_file) == previous["sha256"],
                f"baseline fixture parity {case} historical receipt changed")
        historical_seal = read_json(ROOT / HISTORICAL_PACKET / "seal.json")
        require(historical_seal.get("files", {}).get(f"qualification/{case}.json")
                == previous["sha256"], f"historical qualification seal differs: {case}")
        require(_qualification_report_identity(previous_file, f"historical {case}")
                == current_identity, f"historical fixture/output differs: {case}")
        result_rows.append({
            "case": case,
            "current": {"path": rel(current_path), "bytes": current_path.stat().st_size,
                         "sha256": sha256(current_path)},
            "previous": {"path": previous_path, "bytes": previous["bytes"],
                          "sha256": previous["sha256"]},
            "source": {"bytes": source["bytes"], "sha256": source["sha256"]},
            "output": {"bytes": output["bytes"], "sha256": output["sha256"]},
        })
    require(seen == expected_cases, "baseline fixture parity case set changed")
    return {"receipt": _file_identity(parity_path), "scope": scope,
            "rows": sorted(result_rows, key=lambda row: row["case"])}


def load_historical_qualification() -> dict[str, Any]:
    """Check the sealed 0780 qualification census against the origin Git tree."""

    packet = ROOT / HISTORICAL_PACKET
    base = origin().get("base")
    require(is_git_revision(base), "historical qualification origin revision is invalid")
    prefix = f"{HISTORICAL_PACKET}/"
    metadata = [
        f"{prefix}plan.json",
        f"{prefix}seal.json",
        f"{prefix}qualification/complete.json",
        f"{prefix}qualification/receipts.json",
    ]
    git_metadata = _git_blob_bytes(base, metadata, "historical qualification metadata")
    for name, contents in git_metadata.items():
        current = ROOT / name
        require(current.is_file() and not current.is_symlink(),
                f"historical qualification file is missing: {name}")
        require(sha256(current) == digest_bytes(contents),
                f"historical qualification Git anchor changed: {name}")
    plan_path, seal_path, complete_path, receipts_path = metadata
    try:
        plan = json.loads(git_metadata[plan_path])
        seal = json.loads(git_metadata[seal_path])
        complete = json.loads(git_metadata[complete_path])
        receipts = json.loads(git_metadata[receipts_path])
    except ValueError as error:
        fail(f"historical qualification Git JSON is malformed: {error}")
    require(isinstance(plan, dict) and plan.get("schema") == "litchi.performance.0780.v1",
            "historical 0780 plan is malformed")
    cases = plan.get("cases")
    require(isinstance(cases, list) and len(cases) == 10,
            "historical 0780 case cardinality changed")
    ordinary = [case for case in cases
                if isinstance(case, dict)
                and case.get("mode") in {"capture", "commit", "lifecycle"}]
    diagnostic = [case for case in cases
                  if isinstance(case, dict) and case.get("mode") == "capabilities"]
    require(len(ordinary) == 9 and len(diagnostic) == 1
            and diagnostic[0].get("shape") == "tiny",
            "historical 0780 ordinary/diagnostic case split changed")
    require(all(isinstance(case, dict) and isinstance(case.get("shape"), str)
                and isinstance(case.get("mode"), str) for case in cases),
            "historical 0780 case identity is malformed")
    require(isinstance(complete, dict) and complete.get("children") == 10,
            "historical 0780 qualification child cardinality changed")
    require(isinstance(receipts, list) and len(receipts) == 10,
            "historical 0780 qualification receipt cardinality changed")
    require(isinstance(seal, dict) and seal.get("schema") == "litchi.performance.0780.seal.v1"
            and isinstance(seal.get("files"), dict),
            "historical 0780 seal is malformed")
    sealed = seal["files"]
    expected_cases = {
        f"0-{case['shape']}-{case['mode']}-before"
        for case in cases
    }
    require(len(expected_cases) == 10, "historical 0780 case identities changed")
    report_paths = [f"{prefix}qualification/{case}.json" for case in sorted(expected_cases)]
    git_reports = _git_blob_bytes(base, report_paths, "historical qualification reports")
    rows: list[dict[str, Any]] = []
    seen: set[str] = set()
    for row in receipts:
        require(isinstance(row, dict), "historical 0780 qualification receipt is malformed")
        require(row.get("lane") == "qualification" and row.get("block") == 0
                and row.get("leg") == "before",
                "historical 0780 qualification receipt identity changed")
        case = f"0-{row.get('shape')}-{row.get('mode')}-{row.get('leg')}"
        require(case in expected_cases and case not in seen,
                f"historical 0780 qualification case changed: {case}")
        seen.add(case)
        report = packet / "qualification" / f"{case}.json"
        relative = f"qualification/{case}.json"
        report_receipt = row.get("report")
        require(isinstance(report_receipt, dict)
                and Path(str(report_receipt.get("path", ""))).name == f"{case}.json"
                and report.is_file() and not report.is_symlink(),
                f"historical 0780 report receipt changed: {case}")
        require(relative in sealed and sealed[relative] == sha256(report),
                f"historical 0780 sealed report changed: {case}")
        git_name = f"{prefix}{relative}"
        require(digest_bytes(git_reports[git_name]) == sha256(report),
                f"historical 0780 report Git anchor changed: {case}")
        require(report_receipt.get("bytes") == report.stat().st_size
                and report_receipt.get("sha256") == sha256(report),
                f"historical 0780 report receipt digest changed: {case}")
        rows.append({
            "case": case,
            "report": relative,
            "sha256": sha256(report),
            "git_sha256": digest_bytes(git_reports[git_name]),
        })
    require(seen == expected_cases, "historical 0780 qualification case set changed")
    return {
        "packet": HISTORICAL_PACKET,
        "git_revision": base,
        "metadata_files": len(metadata),
        "plan_cases": len(cases),
        "ordinary_cases": len(ordinary),
        "diagnostic_cases": len(diagnostic),
        "qualification_reports": len(rows),
        "seal_schema": seal.get("schema"),
        "sealed_reports": sorted(rows, key=lambda row: row["case"]),
        "sealed_git_comparison": True,
    }


def load_adoption_policy() -> dict[str, Any]:
    """Bind the frozen decision guard without turning it into a result."""

    policy_path = PACKET / "adoption-policy.json"
    value = read_json(policy_path)
    require(isinstance(value, dict), "adoption policy is malformed")
    require(set(value) == {
        "allocation_count_alone_sufficient", "benefit", "frozen_before_build",
        "latency", "memory", "scope", "useful_public_workflow_benefit_required",
    } and "schema" not in value, "adoption policy fields changed")
    require(value.get("frozen_before_build") is True,
            "adoption policy freeze marker changed")
    latency = value.get("latency")
    require(isinstance(latency, dict)
            and set(latency) == {
                "any_case_violation_rejects", "bootstrap95_low_must_exceed",
                "maximum_ratio", "metric", "resamples", "seed",
            }
            and latency.get("metric") == "paired process p50"
            and latency.get("maximum_ratio") == 1.05
            and latency.get("bootstrap95_low_must_exceed") == 1.0
            and latency.get("resamples") == BOOTSTRAP_RESAMPLES
            and latency.get("seed") == BOOTSTRAP_SEED
            and latency.get("any_case_violation_rejects") is True,
            "adoption latency guard changed")
    memory = value.get("memory")
    require(isinstance(memory, dict)
            and set(memory) == {"net_live_increase_allowed", "peak_above_entry_increase_allowed"}
            and memory.get("net_live_increase_allowed") == 0
            and memory.get("peak_above_entry_increase_allowed") == 0,
            "adoption memory guard changed")
    benefit = value.get("benefit")
    require(isinstance(benefit, dict)
            and set(benefit) == {
                "at_least_one_case_required", "bootstrap95_high_below",
                "eligible_modes", "minimum_improvement_percent",
            }
            and benefit.get("minimum_improvement_percent") == 3.0
            and benefit.get("bootstrap95_high_below") == 1.0
            and benefit.get("eligible_modes") == ["capture", "lifecycle"]
            and benefit.get("at_least_one_case_required") is True,
            "adoption benefit requirement changed")
    require(value.get("allocation_count_alone_sufficient") is False
            and value.get("useful_public_workflow_benefit_required") is True
            and isinstance(value.get("scope"), str) and value["scope"],
            "adoption rationale changed")
    return {"receipt": _file_identity(policy_path), "policy": value}


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
    expected_marker = probe_contract()[2]
    if "marker" in report and expected_marker:
        require(report.get("marker") == expected_marker,
                f"{label} marker identity changed")


def fixture_identity(report: dict[str, Any], job: dict[str, Any], label: str) -> None:
    """Check the generated fixture, including the six per-text vendor attrs."""

    slides, shapes_per_slide = SHAPE_DIMENSIONS[job["shape"]]
    require(report.get("slides") == slides
            and report.get("shapes_per_slide") == shapes_per_slide,
            f"{label} fixture dimensions changed")
    fixture = report.get("fixture")
    require(isinstance(fixture, dict), f"{label} fixture metadata is missing")
    vendor = job["shape"] in {"vendor", "unicode-vendor"}
    expected_injection = {
        "vendor": "same-length-known-uri-near-misses",
        "unicode-vendor": "same-length-valid-utf8-unknown-uris",
    }.get(job["shape"], "none")
    require(fixture.get("injection") == expected_injection,
            f"{label} fixture injection changed")
    expected_tags = slides * shapes_per_slide if vendor else 0
    expected_parts = slides if vendor else 0
    require(fixture.get("slide_parts") == expected_parts
            and fixture.get("replaced_text_tags") == expected_tags,
            f"{label} fixture text coverage changed")
    for key in ("namespace_uris", "attribute_names"):
        values = fixture.get(key)
        require(isinstance(values, list), f"{label} fixture {key} is missing")
        if vendor:
            require(len(values) == 6 and all(isinstance(value, str) and value for value in values)
                    and len(set(values)) == 6,
                    f"{label} fixture {key} cardinality changed")
        else:
            require(values == [], f"{label} ordinary fixture {key} changed")
    require(fixture.get("namespace_declarations") == (6 if vendor else 0)
            and fixture.get("namespaced_attributes") == (6 if vendor else 0),
            f"{label} fixture namespace cardinality changed")


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
    fixture_identity(report, job, label)
    require(report.get("samples_requested") == job["samples"]
            and report.get("warmup") == job["warmup"], f"{label} sample configuration changed")
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
        require(isinstance(metrics, dict) and metrics.get("elapsed_ns") == elapsed_ns,
                f"{label} raw metric identity changed")
        for key in ("slides", "shapes_per_slide", "boxes_per_slide", "text_boxes",
                    "vendor_attributes_per_text"):
            if key in report:
                require(metrics.get(key) == report[key],
                        f"{label} raw metric {key} identity changed")
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
        require(isinstance(verification.get("expected_text"), str)
                and isinstance(verification.get("actual_text"), str)
                and verification.get("expected_text") == verification.get("actual_text"),
                f"{label} semantic text readback changed")
        nonnegative_int(verification.get("semantic_text_bytes"),
                        f"{label} semantic text bytes")
        require(is_sha(verification.get("semantic_text_sha256")),
                f"{label} semantic text digest is missing")
        require(verification.get("readback_bytes") == output["bytes"]
                and verification.get("readback_sha256") == output["sha256"],
                f"{label} readback digest does not match output")
        vendor = job["shape"] in {"vendor", "unicode-vendor"}
        if vendor:
            require(verification.get("unknown_namespace_check") is True
                    and verification.get("unknown_namespace_occurrences")
                    == SHAPE_DIMENSIONS[job["shape"]][0]
                    * SHAPE_DIMENSIONS[job["shape"]][1],
                    f"{label} unknown namespace oracle changed")
        else:
            require(verification.get("unknown_namespace_check") is None
                    and verification.get("unknown_namespace_occurrences") is None,
                    f"{label} ordinary namespace oracle changed")
        expected_marker_matches = True if job["mode"] in {"commit", "lifecycle"} else None
        require("marker_matches" in verification
                and verification.get("marker_matches") is expected_marker_matches,
                f"{label} marker verification failed")
        allocation = sample.get("allocation")
        if kind == "native":
            require(allocation is None, f"{label} native report contains allocation metrics")
        else:
            values = allocation_from_sample(sample, f"{label} sample {index}")
            for field, number in values.items():
                allocations[field].append(number)
    require(outputs and len(set(outputs)) == 1,
            f"{label} output identity is not deterministic across samples")
    return {"stats": stats(elapsed), "allocation": None if kind == "native" else allocations,
            "outputs": outputs}


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
                        "allocation": outcome["allocation"],
                        "outputs": outcome["outputs"]})
    return entries


def check_fixture_parity(entries: Iterable[dict[str, Any]]) -> dict[str, Any]:
    """Require stable source and serialized output identities across all legs.

    The ordinary and vendor fixtures are generated by the probe itself.  Every
    lane must therefore see the same source bytes for a shape, and a given
    shape/mode must publish the same bytes in before and after legs.  This is
    a correctness gate for the measurement corpus, not a historical timing
    comparison.
    """

    source_by_shape: dict[str, str] = {}
    output_by_case: dict[tuple[str, str], tuple[int, str]] = {}
    rows = 0
    for item in entries:
        identity = item["identity"]
        shape, mode = identity["shape"], identity["mode"]
        digest = item["source_sha256"]
        previous_source = source_by_shape.setdefault(shape, digest)
        require(previous_source == digest,
                f"source fixture changed across lanes for {shape}")
        outputs = item.get("outputs")
        require(isinstance(outputs, list) and outputs
                and len({tuple(value) for value in outputs}) == 1,
                f"output identity is not stable for {shape}/{mode}")
        output = tuple(outputs[0])
        key = (shape, mode)
        previous_output = output_by_case.setdefault(key, output)
        require(previous_output == output,
                f"output fixture changed across legs for {shape}/{mode}")
        rows += 1
    require(set(source_by_shape) == set(SHAPES), "source fixture shape coverage changed")
    require(set(output_by_case) == {(case["shape"], case["mode"]) for case in CASES},
            "output fixture case coverage changed")
    return {
        "source_by_shape": dict(sorted(source_by_shape.items())),
        "output_by_case": {
            f"{shape}/{mode}": {"bytes": value[0], "sha256": value[1]}
            for (shape, mode), value in sorted(output_by_case.items())
        },
        "rows": rows,
        "ordinary_shapes": ["tiny", "medium", "large"],
        "vendor_shapes": ["vendor", "unicode-vendor"],
    }


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


def decision_guards(native: dict[str, Any], allocation: dict[str, Any],
                    policy: dict[str, Any]) -> dict[str, Any]:
    """Derive the frozen adoption guards from the paired measurements."""

    value = policy["policy"]
    latency_policy = value["latency"]
    benefit_policy = value["benefit"]
    maximum = float(latency_policy["maximum_ratio"])
    low_limit = float(latency_policy["bootstrap95_low_must_exceed"])
    minimum_improvement = float(benefit_policy["minimum_improvement_percent"])
    latency_violations: list[dict[str, Any]] = []
    benefits: list[dict[str, Any]] = []
    p50_groups = native["paired_by_block_before_after"]
    for key, group in sorted(p50_groups.items()):
        metric = group["metrics"]["p50"]
        median = metric["ratio_median"]
        bootstrap = metric["bootstrap"]
        low = bootstrap["ci_low"]
        high = bootstrap["ci_high"]
        if median is not None and low is not None and median > maximum and low > low_limit:
            latency_violations.append({
                "case": key,
                "ratio_median": median,
                "bootstrap_ci_low": low,
                "bootstrap_ci_high": high,
                "change_percent_median": metric["change_percent_median"],
            })
        shape, mode = key.split("/", 1)
        if mode in set(benefit_policy["eligible_modes"]):
            if median is not None and high is not None:
                improvement = (1.0 - median) * 100.0
                if improvement >= minimum_improvement and high < 1.0:
                    benefits.append({
                        "case": key,
                        "improvement_percent": improvement,
                        "ratio_median": median,
                        "bootstrap_ci_low": low,
                        "bootstrap_ci_high": high,
                    })
    resource_violations: list[dict[str, Any]] = []
    for key, group in sorted(allocation["paired_by_block_before_after"].items()):
        for metric_name in ("net_live", "peak_above_entry"):
            metric = group["metrics"][metric_name]
            for row in metric["by_block"]:
                if row["after"] > row["before"]:
                    resource_violations.append({
                        "case": key,
                        "metric": metric_name,
                        "block": row["block"],
                        "before": row["before"],
                        "after": row["after"],
                    })
    return {
        "latency_violations": latency_violations,
        "resource_violations": resource_violations,
        "eligible_benefits": benefits,
        "benefit_satisfied": bool(benefits),
        "latency_guard_passed": not latency_violations,
        "resource_guard_passed": not resource_violations,
        "adoption_eligible": not latency_violations and not resource_violations and bool(benefits),
        "policy_thresholds": {
            "maximum_ratio": maximum,
            "bootstrap95_low_must_exceed": low_limit,
            "minimum_improvement_percent": minimum_improvement,
        },
    }


def analyze() -> dict[str, Any]:
    plan = load_plan()
    builds, cleanup, cleanup_verified = load_builds(plan)
    before_source, after_source = builds["before"]["source"], builds["after"]["source"]
    revision_transition = load_revision_transition(builds)
    architecture_inputs = load_architecture_inputs()
    historical_qualification = load_historical_qualification()
    disposition = load_disposition(before_source, after_source)
    quality = check_quality(after_source)
    test_summary = check_test_summary(quality)
    native_entries = load_lane(plan, "native", builds, after_source, cleanup, cleanup_verified)
    allocation_entries = load_lane(plan, "allocation", builds, after_source, cleanup, cleanup_verified)
    qualification_entries = load_lane(plan, "qualification", builds, before_source,
                                      cleanup, cleanup_verified)
    require(len(native_entries) == 180, "native cardinality changed")
    require(len(allocation_entries) == 60, "allocation cardinality changed")
    require(len(qualification_entries) == 15, "qualification cardinality changed")
    digests = {
        item["identity"]["shape"]: item["source_sha256"]
        for item in (*native_entries, *allocation_entries, *qualification_entries)
    }
    require(len(digests) == len(SHAPES), "source digest shape coverage changed")
    all_entries = (*native_entries, *allocation_entries, *qualification_entries)
    sample_count = sum(item["stats"]["count"] for item in all_entries)
    require(len(all_entries) == 255 and sample_count == 5595,
            "aggregate report or sample cardinality changed")
    for item in all_entries:
        require(item["source_sha256"] == digests[item["identity"]["shape"]],
                "source digest changed between lanes")
    fixture_parity = check_fixture_parity(all_entries)
    baseline_fixture_parity = load_baseline_fixture_parity()
    adoption_policy = load_adoption_policy()
    native_result = native_analysis(native_entries)
    allocation_result = allocation_analysis(allocation_entries)
    guards = decision_guards(native_result, allocation_result, adoption_policy)
    optional = optional_seal()
    return {
        "schema": "litchi-0785-known-uri-analysis-v1",
        "plan_schema": plan["schema"], "quality": quality,
        "test_summary": test_summary,
        "counts": {
            "reports": len(all_entries),
            "samples": sample_count,
            "native_reports": len(native_entries),
            "allocation_reports": len(allocation_entries),
            "qualification_reports": len(qualification_entries),
        },
        "fixture_parity": fixture_parity,
        "baseline_fixture_parity": baseline_fixture_parity,
        "revision_transition": revision_transition,
        "architecture_inputs": architecture_inputs,
        "historical_qualification": historical_qualification,
        "adoption_policy": adoption_policy,
        "decision_guards": guards,
        "source": {"before": before_source, "after": after_source,
                    "changed_files": sorted(plan["source_allowlist"])},
        "disposition": disposition,
        "native": {"children": len(native_entries), "blocks": plan["native"]["blocks"],
                    "samples": plan["native"]["samples"], "warmup": plan["native"]["warmup"],
                    "analysis": native_result,
                    "receipts": [{"shape": i["identity"]["shape"], "mode": i["identity"]["mode"],
                                  "block": i["identity"]["block"], "leg": i["identity"]["leg"],
                                  "report": rel(i["report_path"]), "report_sha256": i["report_sha256"]}
                                 for i in native_entries]},
        "allocation": {"children": len(allocation_entries), "blocks": plan["allocation"]["blocks"],
                        "samples": plan["allocation"]["samples"], "warmup": plan["allocation"]["warmup"],
                        "analysis": allocation_result,
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
            "frozen_inputs_checked": True,
            "revision_transition_checked": True,
            "architecture_inputs_checked": True,
            "historical_qualification_git_checked": True,
            "aggregate_counts_checked": True,
            "probe_contract_checked": True,
            "fixture_output_parity_checked": True,
            "baseline_fixture_parity_replayed": True,
            "adoption_policy_checked": True,
            "fixed_release_profile_checked": True,
            "quality_commands_checked_exactly": True,
            "native_has_no_allocation_metrics": True,
            "source_digest_and_semantic_checks": True,
            "qualification_preflight_outside_paired_matrix": True,
            "source_change_allowlist": sorted(plan["source_allowlist"]),
            "cleanup_binary_witness_required_when_missing": True,
            "optional_file_inventory_seal_checked": optional,
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
