#!/usr/bin/env python3
"""Prepare the source-matched W5 OLE owner qualification drivers.

The pre-owner source does not contain the later performance examples.  This
preparer therefore builds two *external* Cargo driver directories whose only
consumer files are the reviewed probe inputs.  It never copies a candidate
production file into the baseline worktree and never invokes Cargo, Rust, or a
probe.  The receipt binds both clean source worktrees, every driver input, and
the reviewed consumer adapter before the native runner is allowed to build.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
from typing import Any, Mapping


SCHEMA = "litchi.ole-owner-qualification.preparation.v1"
LANE = "w5-ole-owner"
BASELINE_COMMIT = "0ce4bb8b61f03a1e5497633f51b954e76f98aa8c"
REVIEWED_PROBE_SHA256 = "ad4c52c01333fbc13c7b3cca18e2c0a0d2ccf89d5c07972e2e453fa9bd8d713a"
CANDIDATE_ADAPTER_SHA256 = "dc11e6533c43434c314bd359ad68e458555291f4c134b44b13c1f3ce29fb4508"
CLONE_PROBE_SHA256 = "0a23c99ed8958e0df1aa51e5f37e7a949328149074b2096613c972d2b8cf6e98"
LOCK_SHA256 = "bc5a133f0668142d18d637b2b196097e855c8890697cb3bbf43aeb937145f04a"
ORACLE_SOURCE_SHA256 = "60b506c1bfc277c260933730ce2f754e59f4745d40c753ffa95900b49d276daf"
PROBE_PATH = Path("tools/perf-baseline/examples/ole_w5_probe.rs")
CLONE_PATH = Path("tools/perf-baseline/examples/ole_package_clone_probe.rs")
ORACLE_PATH = Path("tools/perf-baseline/examples/ole_owner_oracle.rs")
CFB_SOURCE_PATH = Path("crates/litchi-cfb/src/file.rs")
# These are provenance labels, not trust anchors.  The preparer reads the
# selected source and probe bytes below and binds their actual hashes in the
# receipt; a historical label alone can never admit a build.
REVIEWED_PROBE_SOURCE = (
    "external matched-harness artifact from 0359j; "
    "historical source a84650cd70390641f7734dc1f3362bcb9cd1e025 "
    "tools/perf-baseline/examples/ole_w5_probe.rs"
)
CLONE_PROBE_SOURCE = (
    "external matched-harness artifact from 0359j; "
    "historical source a84650cd70390641f7734dc1f3362bcb9cd1e025 "
    "tools/perf-baseline/examples/ole_package_clone_probe.rs"
)
LOCK_SOURCE = (
    "external matched-harness lock retained by 0359j W5 qualification; "
    "Cargo.lock from the reviewed current driver"
)
FULL_SHA = re.compile(r"^[0-9a-f]{40}$")
HASH = re.compile(r"^[0-9a-f]{64}$")
UPSTREAM_COMPONENT = "3rd" + "party"

# This is the small standalone package used by the historical W5 driver.  It
# deliberately has no workspace membership and no production source files.
# Path dependencies are filled with the selected source root at preparation
# time; the pinned lock is supplied separately because it is an external
# qualification input.
HARNESS_MANIFEST = """[package]
name = \"litchi-perf-baseline\"
version = \"0.1.0\"
edition = \"2024\"
publish = false

[features]
default = [\"current-api\"]
current-api = [\"current-owner-api\", \"current-property-api\"]
current-owner-api = []
current-property-api = []
legacy-owner-api = []
legacy-property-api = []
w5-counting-allocator = []

[[example]]
name = \"ole_w5_probe\"
path = \"examples/ole_w5_probe.rs\"

[[example]]
name = \"ole_package_clone_probe\"
path = \"examples/ole_package_clone_probe.rs\"

[[example]]
name = \"ole_owner_oracle\"
path = \"examples/ole_owner_oracle.rs\"

[dependencies]
litchi-cfb = { path = \"@SOURCE_ROOT@/crates/litchi-cfb\", features = [\"write\"] }
litchi-doc = { path = \"@SOURCE_ROOT@/crates/litchi-doc\", features = [\"performance-diagnostics\"] }
litchi-ole-common = { path = \"@SOURCE_ROOT@/crates/litchi-ole-common\" }
litchi-xls = { path = \"@SOURCE_ROOT@/crates/litchi-xls\" }
sha2 = \"0.11\"

[workspace]
"""

REQUIRED_PROBE_MARKERS = (
    "const SIBLINGS: usize = 64;",
    "const DOC_HISTORY_STEPS: usize = 128;",
    "const XLS_HISTORY_STEPS: usize = 128;",
    '"open"',
    '"begin_edit"',
    '"noop_commit"',
    '"replace_stream"',
    '"read"',
    '"edit"',
    '"doc_tracked_revision"',
    '"xls_cells"',
    '"xls_cells_stage"',
    '"doc_history"',
    '"xls_history"',
    "verify_doc_history",
    "verify_xls_history",
    "patches",
    "packages",
    "inverse()",
    "record_dealloc",
    "record_realloc",
    "#[global_allocator]",
)
REQUIRED_CLONE_MARKERS = (
    "const SMALL_SIBLINGS: usize = 1;",
    "const LARGE_SIBLINGS: usize = 1_024;",
    "clone_drop",
    "black_box(&clone)",
)
BASELINE_CFB_OPEN = "pub fn open(reader: R) -> Result<Self, OleError>"
BASELINE_CFB_LIMITED_OPEN = "pub fn open_with_limits(mut reader: R, limits: OleFileLimits)"
CANDIDATE_CFB_OPEN = "pub fn open(mut reader: R, limits: OleFileLimits) -> Result<Self, OleError>"
CANDIDATE_CFB_NO_DEFAULT_OPEN = "pub fn open(reader: R) -> Result<Self, OleError>"


class PreparationError(RuntimeError):
    """A source, adapter, or driver identity is not admissible."""


def _resolve(path: str | os.PathLike[str], *, base: Path | None = None) -> Path:
    candidate = Path(path)
    if not candidate.is_absolute():
        if base is None:
            raise PreparationError(f"path must be absolute: {path!r}")
        candidate = base / candidate
    probe = Path(candidate.anchor)
    for part in candidate.parts[1:]:
        probe /= part
        if probe.is_symlink():
            raise PreparationError(f"symlink path component is forbidden: {candidate}")
    resolved = candidate.expanduser().resolve(strict=False)
    if UPSTREAM_COMPONENT in resolved.parts:
        raise PreparationError(f"untracked upstream checkout path is forbidden: {resolved}")
    return resolved


def _under(path: Path, parent: Path) -> bool:
    try:
        path.relative_to(parent)
    except ValueError:
        return False
    return True


def _disjoint(paths: Mapping[str, Path]) -> None:
    entries = list(paths.items())
    for index, (left_name, left) in enumerate(entries):
        for right_name, right in entries[index + 1 :]:
            if left == right or _under(left, right) or _under(right, left):
                raise PreparationError(
                    f"paths must be pairwise disjoint: {left_name}={left}, {right_name}={right}"
                )


def _sha256(path: Path) -> tuple[str, int]:
    if path.is_symlink() or not path.is_file():
        raise PreparationError(f"expected a regular file: {path}")
    digest = hashlib.sha256()
    size = 0
    with path.open("rb") as stream:
        while block := stream.read(1024 * 1024):
            digest.update(block)
            size += len(block)
    return digest.hexdigest(), size


def _file_record(path: Path) -> dict[str, Any]:
    digest, size = _sha256(path)
    return {"path": str(path), "bytes": size, "sha256": digest}


def _git(root: Path, *args: str) -> str:
    result = subprocess.run(
        ["git", "-C", str(root), *args],
        check=False,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )
    if result.returncode:
        raise PreparationError(f"git {' '.join(args)} failed in {root}: {result.stderr.strip()}")
    return result.stdout


def _status_paths(status: str) -> list[str]:
    paths: list[str] = []
    for line in status.splitlines():
        if len(line) < 4:
            raise PreparationError(f"malformed git status line: {line!r}")
        path = line[3:]
        if " -> " in path:
            path = path.split(" -> ", 1)[1]
        paths.append(path)
    return paths


def _source_state(root: Path) -> dict[str, Any]:
    root = _resolve(root)
    if not root.is_dir():
        raise PreparationError(f"source root is not a directory: {root}")
    top = _resolve(_git(root, "rev-parse", "--show-toplevel").strip())
    if top != root:
        raise PreparationError(f"source is not a worktree top-level: {root} != {top}")
    status = _git(root, "status", "--porcelain=v1", "--untracked-files=all")
    diff = _git(root, "diff", "--no-ext-diff", "--binary")
    return {
        "root": str(root),
        "head": _git(root, "rev-parse", "HEAD").strip(),
        "tree": _git(root, "rev-parse", "HEAD^{tree}").strip(),
        "status": status,
        "status_paths": _status_paths(status),
        "clean": not bool(status),
        "status_sha256": hashlib.sha256(status.encode()).hexdigest(),
        "diff_sha256": hashlib.sha256(diff.encode()).hexdigest(),
        "production_tree_sha256": hashlib.sha256(
            _git(root, "ls-tree", "-r", "--full-tree", "HEAD").encode()
        ).hexdigest(),
    }


def _require_hash(value: str, label: str) -> str:
    if not HASH.fullmatch(value):
        raise PreparationError(f"{label} must be a full lowercase SHA-256")
    return value


def _require_commit(value: str, label: str) -> str:
    if not FULL_SHA.fullmatch(value):
        raise PreparationError(f"{label} must be a full lowercase commit SHA")
    return value


def _check_markers(path: Path, markers: tuple[str, ...], label: str) -> None:
    try:
        text = path.read_text(encoding="utf-8")
    except (OSError, UnicodeDecodeError) as error:
        raise PreparationError(f"cannot read {label}: {error}") from error
    missing = [marker for marker in markers if marker not in text]
    if missing:
        raise PreparationError(f"{label} is missing reviewed workload markers: {missing!r}")


def _cfb_api_binding(root: Path, *, baseline: bool) -> dict[str, Any]:
    """Bind the probe's CFB call arity to the actual selected source API.

    The baseline and candidate probes are intentionally different only at the
    CFB limits call site.  A probe hash and a Git revision are not sufficient
    evidence for that claim: a dirty generated source tree, stale adapter, or
    feature mismatch could otherwise compile a different API than the receipt
    describes.  Keep this check deliberately textual and narrow so preparation
    remains build-free while still checking the public function signatures.
    """

    source = _resolve(CFB_SOURCE_PATH, base=root)
    try:
        text = source.read_text(encoding="utf-8")
    except (OSError, UnicodeDecodeError) as error:
        raise PreparationError(f"cannot read selected CFB API source: {error}") from error
    if baseline:
        required = BASELINE_CFB_OPEN
        if required not in text or BASELINE_CFB_LIMITED_OPEN not in text:
            raise PreparationError(
                "baseline CFB source does not expose the reviewed one-argument open plus open_with_limits API"
            )
        if CANDIDATE_CFB_OPEN in text:
            raise PreparationError("baseline CFB source unexpectedly exposes the current two-argument open API")
        return {
            "source_path": str(source),
            "open_signature": required,
            "limited_open_signature": BASELINE_CFB_LIMITED_OPEN,
            "api_generation": "pre-owner explicit open_with_limits",
        }
    required = CANDIDATE_CFB_OPEN
    if required not in text or CANDIDATE_CFB_NO_DEFAULT_OPEN in text:
        raise PreparationError(
            "candidate CFB source does not expose only the reviewed two-argument open API"
        )
    return {
        "source_path": str(source),
        "open_signature": required,
        "api_generation": "current explicit-limits open",
    }


def _probe_api_binding(path: Path, *, baseline: bool) -> dict[str, Any]:
    try:
        text = path.read_text(encoding="utf-8")
    except (OSError, UnicodeDecodeError) as error:
        raise PreparationError(f"cannot read selected probe API adapter: {error}") from error
    if baseline:
        if "OleFile::open(Cursor::new(package.into_inner()))" not in text:
            raise PreparationError("reviewed baseline probe does not use the one-argument CFB open")
        if "OleFileLimits::default()" in text:
            raise PreparationError("reviewed baseline probe unexpectedly contains a CFB limits argument")
        return {"open_call": "OleFile::open(reader)", "adapter": "immutable historical baseline"}
    if text.count("OleFileLimits::default()") < 2 or "OleFile::open(\n        Cursor::new(package.into_inner())," not in text:
        raise PreparationError("candidate probe is not the reviewed explicit-CFB-limits adapter")
    return {"open_call": "OleFile::open(reader, OleFileLimits::default())", "adapter": "CFB limits arity only"}


def _render_manifest(source_root: Path) -> str:
    # The source roots are resolved absolute paths.  Only POSIX paths are
    # admitted by the surrounding worktree; escaping backslashes keeps this
    # safe if the protocol is inspected on another host.
    rendered = str(source_root).replace("\\", "\\\\").replace('"', '\\"')
    return HARNESS_MANIFEST.replace("@SOURCE_ROOT@", rendered)


def _fresh_directory(path: Path, label: str) -> None:
    if path.is_symlink():
        raise PreparationError(f"{label} may not be a symlink: {path}")
    if path.exists():
        if not path.is_dir():
            raise PreparationError(f"{label} is not a directory: {path}")
        try:
            next(path.iterdir())
        except StopIteration:
            return
        raise PreparationError(f"{label} must be fresh and empty: {path}")
    path.mkdir(parents=True)


def _write_driver(
    *,
    destination: Path,
    source_root: Path,
    probe: Path,
    clone_probe: Path,
    oracle_probe: Path,
    lock: Path,
) -> dict[str, Any]:
    _fresh_directory(destination, "driver root")
    examples = destination / "examples"
    examples.mkdir()
    manifest = destination / "Cargo.toml"
    lock_destination = destination / "Cargo.lock"
    probe_destination = examples / PROBE_PATH.name
    clone_destination = examples / CLONE_PATH.name
    oracle_destination = examples / ORACLE_PATH.name
    manifest.write_text(_render_manifest(source_root), encoding="utf-8")
    shutil.copyfile(lock, lock_destination)
    shutil.copyfile(probe, probe_destination)
    shutil.copyfile(clone_probe, clone_destination)
    shutil.copyfile(oracle_probe, oracle_destination)
    return {
        "root": str(destination),
        "source_root": str(source_root),
        "files": [
            _file_record(manifest),
            _file_record(lock_destination),
            _file_record(probe_destination),
            _file_record(clone_destination),
            _file_record(oracle_destination),
        ],
    }


def _driver_files(driver: Mapping[str, Any]) -> list[Path]:
    root = _resolve(str(driver["root"]))
    files = driver.get("files")
    if not isinstance(files, list) or len(files) != 5:
        raise PreparationError("driver must record exactly five input files")
    result: list[Path] = []
    for record in files:
        if not isinstance(record, Mapping):
            raise PreparationError("driver file record is malformed")
        path = _resolve(str(record.get("path", "")))
        if not _under(path, root):
            raise PreparationError(f"driver file escapes driver root: {path}")
        digest, size = _sha256(path)
        if digest != record.get("sha256") or size != record.get("bytes"):
            raise PreparationError(f"driver file identity changed: {path}")
        result.append(path)
    expected = {
        root / "Cargo.toml",
        root / "Cargo.lock",
        root / "examples" / PROBE_PATH.name,
        root / "examples" / CLONE_PATH.name,
        root / "examples" / ORACLE_PATH.name,
    }
    if set(result) != expected:
        raise PreparationError("driver file set differs from Cargo manifest, lock, probes, and oracle")
    actual_files = {
        path
        for path in root.rglob("*")
        if path.is_file() and not path.is_symlink()
    }
    if actual_files != expected:
        raise PreparationError("driver contains an unexpected or missing input file")
    return result


def _validate_source_record(record: Mapping[str, Any], *, expected_head: str | None = None) -> dict[str, Any]:
    root = _resolve(str(record.get("root", "")))
    current = _source_state(root)
    for key in (
        "root",
        "head",
        "tree",
        "status",
        "status_paths",
        "clean",
        "status_sha256",
        "diff_sha256",
        "production_tree_sha256",
    ):
        if current.get(key) != record.get(key):
            raise PreparationError(f"source identity changed for {root}: {key}")
    if expected_head is not None and current["head"] != expected_head:
        raise PreparationError(f"source HEAD is not the reviewed {expected_head}: {current['head']}")
    if not current["clean"]:
        raise PreparationError(f"source worktree is dirty: {root}")
    return current


def prepare(
    *,
    baseline_root: Path,
    candidate_root: Path,
    candidate_ref: str,
    reviewed_probe: Path,
    clone_probe: Path,
    oracle_probe: Path,
    lock: Path,
    drivers_root: Path,
    output: Path,
) -> dict[str, Any]:
    baseline_root = _resolve(baseline_root)
    candidate_root = _resolve(candidate_root)
    reviewed_probe = _resolve(reviewed_probe)
    clone_probe = _resolve(clone_probe)
    oracle_probe = _resolve(oracle_probe)
    lock = _resolve(lock)
    drivers_root = _resolve(drivers_root)
    output = _resolve(output)
    _require_commit(candidate_ref, "candidate_ref")
    baseline = _source_state(baseline_root)
    candidate = _source_state(candidate_root)
    if baseline["head"] != BASELINE_COMMIT:
        raise PreparationError(f"baseline HEAD must be {BASELINE_COMMIT}, got {baseline['head']}")
    if not baseline["clean"]:
        raise PreparationError("baseline must be clean; no production overlay is permitted")
    if candidate["head"] != candidate_ref or not candidate["clean"]:
        raise PreparationError("candidate must be clean and exactly at candidate_ref")
    _disjoint({"baseline": baseline_root, "candidate": candidate_root, "drivers": drivers_root})
    if _under(output, baseline_root) or _under(output, candidate_root):
        raise PreparationError("preparation receipt must be outside both source worktrees")
    if output.is_symlink() or output.exists():
        raise PreparationError(f"preparation receipt must be a new file: {output}")
    if _under(output, drivers_root) or output == drivers_root:
        raise PreparationError("preparation receipt must be outside the driver directory")
    reviewed_hash, reviewed_bytes = _sha256(reviewed_probe)
    clone_hash, clone_bytes = _sha256(clone_probe)
    oracle_hash, oracle_bytes = _sha256(oracle_probe)
    lock_hash, lock_bytes = _sha256(lock)
    if reviewed_hash != REVIEWED_PROBE_SHA256:
        raise PreparationError("reviewed immutable probe hash differs from ad4c input")
    if clone_hash != CLONE_PROBE_SHA256:
        raise PreparationError("clone probe hash differs from reviewed input")
    if oracle_hash != ORACLE_SOURCE_SHA256:
        raise PreparationError("oracle source hash differs from reviewed input")
    if lock_hash != LOCK_SHA256:
        raise PreparationError("Cargo.lock hash differs from the pinned W5 lock")
    candidate_probe = _resolve(PROBE_PATH, base=candidate_root)
    candidate_probe_hash, candidate_probe_bytes = _sha256(candidate_probe)
    if candidate_probe_hash != CANDIDATE_ADAPTER_SHA256:
        raise PreparationError("candidate probe is not the reviewed explicit-CFB-limits adapter")
    _check_markers(reviewed_probe, REQUIRED_PROBE_MARKERS, "reviewed immutable probe")
    _check_markers(candidate_probe, REQUIRED_PROBE_MARKERS, "candidate probe adapter")
    _check_markers(clone_probe, REQUIRED_CLONE_MARKERS, "clone probe")
    baseline_cfb = _cfb_api_binding(baseline_root, baseline=True)
    candidate_cfb = _cfb_api_binding(candidate_root, baseline=False)
    baseline_probe_binding = _probe_api_binding(reviewed_probe, baseline=True)
    candidate_probe_binding = _probe_api_binding(candidate_probe, baseline=False)
    _fresh_directory(drivers_root, "drivers root")
    baseline_driver = _write_driver(
        destination=drivers_root / "baseline",
        source_root=baseline_root,
        probe=reviewed_probe,
        clone_probe=clone_probe,
        oracle_probe=oracle_probe,
        lock=lock,
    )
    candidate_driver = _write_driver(
        destination=drivers_root / "candidate",
        source_root=candidate_root,
        probe=candidate_probe,
        clone_probe=clone_probe,
        oracle_probe=oracle_probe,
        lock=lock,
    )
    after_baseline = _source_state(baseline_root)
    if after_baseline["tree"] != baseline["tree"] or after_baseline["status_paths"]:
        raise PreparationError("baseline production tree changed while preparing external drivers")
    if after_baseline["production_tree_sha256"] != baseline["production_tree_sha256"]:
        raise PreparationError("baseline production tree digest changed while preparing drivers")
    output.parent.mkdir(parents=True, exist_ok=True)
    preparation: dict[str, Any] = {
        "schema": SCHEMA,
        "lane": LANE,
        "baseline": {
            **baseline,
            "after_preparation": after_baseline,
            "commit": BASELINE_COMMIT,
            "api_features": ["legacy-owner-api", "legacy-property-api"],
            "driver": baseline_driver,
        },
        "candidate": {
            **candidate,
            "commit": candidate_ref,
            "api_features": ["current-api"],
            "probe": {
                "path": str(candidate_probe),
                "sha256": candidate_probe_hash,
                "bytes": candidate_probe_bytes,
                "adapter_of": REVIEWED_PROBE_SHA256,
                "adapter_policy": "reviewed CFB explicit-limits consumer arity only",
            },
            "driver": candidate_driver,
        },
        "shared_inputs": {
            "reviewed_probe": {
                "path": str(reviewed_probe),
                "sha256": reviewed_hash,
                "bytes": reviewed_bytes,
                "provenance": REVIEWED_PROBE_SOURCE,
            },
            "fixture": {
                "kind": "probe-owned deterministic OLE/DOC/XLS corpus",
                "generator_sha256": reviewed_hash,
                "candidate_generator_sha256": candidate_probe_hash,
                "external_path": None,
            },
            "clone_probe": {
                "path": str(clone_probe),
                "sha256": clone_hash,
                "bytes": clone_bytes,
                "provenance": CLONE_PROBE_SOURCE,
            },
            "oracle_probe": {
                "path": str(oracle_probe),
                "sha256": oracle_hash,
                "bytes": oracle_bytes,
                "provenance": "tracked source-only correctness oracle; compiled in both endpoint drivers",
            },
            "lock": {
                "path": str(lock),
                "sha256": lock_hash,
                "bytes": lock_bytes,
                "provenance": LOCK_SOURCE,
            },
            "api_binding": {
                "baseline": {"cfb": baseline_cfb, "probe": baseline_probe_binding},
                "candidate": {"cfb": candidate_cfb, "probe": candidate_probe_binding},
                "policy": "the baseline one-argument open and candidate explicit-limits open are checked against source text; no implementation overlay",
            },
            "routes": [
                "open",
                "begin_edit",
                "noop_commit",
                "replace_stream",
                "read",
                "edit",
                "doc_tracked_revision",
                "xls_cells",
                "xls_cells_stage",
                "doc_history",
                "xls_history",
            ],
            "history_retention": {
                "doc_steps": 128,
                "xls_steps": 128,
                "packages": 129,
                "patches": 128,
            },
            "clone_sizes": {"small_siblings": 1, "large_siblings": 1024},
        },
        "baseline_overlay": {
            "allowed_paths": [],
            "production_tree": "unchanged; driver files are external",
            "baseline_status_paths": [],
        },
        "drivers_root": str(drivers_root),
        "claim_scope": "external source and consumer preparation only; no Cargo, timing, or acceptance claim",
    }
    output.write_text(json.dumps(preparation, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    return preparation


def validate_preparation_receipt(
    receipt_path: Path,
    *,
    candidate_root: Path | None = None,
    candidate_ref: str | None = None,
) -> dict[str, Any]:
    receipt_path = _resolve(receipt_path)
    if receipt_path.is_symlink() or not receipt_path.is_file():
        raise PreparationError(f"preparation receipt is missing: {receipt_path}")
    try:
        document = json.loads(receipt_path.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as error:
        raise PreparationError(f"cannot read preparation receipt: {error}") from error
    if not isinstance(document, dict) or document.get("schema") != SCHEMA or document.get("lane") != LANE:
        raise PreparationError("preparation schema or lane differs")
    baseline = document.get("baseline")
    candidate = document.get("candidate")
    shared = document.get("shared_inputs")
    overlay = document.get("baseline_overlay")
    if not isinstance(baseline, Mapping) or not isinstance(candidate, Mapping) or not isinstance(shared, Mapping):
        raise PreparationError("preparation endpoints or shared inputs are missing")
    if not isinstance(overlay, Mapping) or overlay.get("allowed_paths") != []:
        raise PreparationError("baseline preparation permits a production overlay")
    if overlay.get("production_tree") != "unchanged; driver files are external":
        raise PreparationError("baseline production-tree policy differs")
    _validate_source_record(baseline, expected_head=BASELINE_COMMIT)
    selected_candidate_root = _resolve(candidate_root) if candidate_root is not None else _resolve(str(candidate.get("root", "")))
    if selected_candidate_root != _resolve(str(candidate.get("root", ""))):
        raise PreparationError("candidate root differs from preparation")
    selected_ref = candidate_ref or str(candidate.get("commit", ""))
    _require_commit(selected_ref, "candidate commit")
    if selected_ref != candidate.get("commit"):
        raise PreparationError("candidate commit differs from preparation")
    _validate_source_record(candidate, expected_head=selected_ref)
    if not isinstance(shared.get("routes"), list) or shared["routes"] != [
        "open",
        "begin_edit",
        "noop_commit",
        "replace_stream",
        "read",
        "edit",
        "doc_tracked_revision",
        "xls_cells",
        "xls_cells_stage",
        "doc_history",
        "xls_history",
    ]:
        raise PreparationError("preparation route set is incomplete or reordered")
    retention = shared.get("history_retention")
    if retention != {"doc_steps": 128, "xls_steps": 128, "packages": 129, "patches": 128}:
        raise PreparationError("retained history workload differs")
    clone_sizes = shared.get("clone_sizes")
    if clone_sizes != {"small_siblings": 1, "large_siblings": 1024}:
        raise PreparationError("clone workload differs")
    reviewed = shared.get("reviewed_probe")
    fixture = shared.get("fixture")
    clone = shared.get("clone_probe")
    oracle = shared.get("oracle_probe")
    lock = shared.get("lock")
    if (
        not isinstance(reviewed, Mapping)
        or not isinstance(fixture, Mapping)
        or not isinstance(clone, Mapping)
        or not isinstance(oracle, Mapping)
        or not isinstance(lock, Mapping)
    ):
        raise PreparationError("shared probe/fixture/clone/lock identities are missing")
    if reviewed.get("provenance") != REVIEWED_PROBE_SOURCE:
        raise PreparationError("reviewed probe provenance is not the reviewed external artifact")
    if clone.get("provenance") != CLONE_PROBE_SOURCE:
        raise PreparationError("clone probe provenance is not the reviewed external artifact")
    if lock.get("provenance") != LOCK_SOURCE:
        raise PreparationError("Cargo.lock provenance is not the reviewed external artifact")
    if oracle.get("provenance") != "tracked source-only correctness oracle; compiled in both endpoint drivers":
        raise PreparationError("oracle provenance is missing")
    if fixture.get("kind") != "probe-owned deterministic OLE/DOC/XLS corpus" or fixture.get("external_path") is not None:
        raise PreparationError("fixture must remain probe-owned and external-path free")
    for record, expected, label in (
        (reviewed, REVIEWED_PROBE_SHA256, "reviewed probe"),
        (clone, CLONE_PROBE_SHA256, "clone probe"),
        (oracle, ORACLE_SOURCE_SHA256, "oracle probe"),
        (lock, LOCK_SHA256, "Cargo.lock"),
    ):
        if record.get("sha256") != expected:
            raise PreparationError(f"{label} pin differs")
        path = _resolve(str(record.get("path", "")))
        digest, size = _sha256(path)
        if digest != expected or size != record.get("bytes"):
            raise PreparationError(f"{label} identity changed")
    candidate_probe = candidate.get("probe")
    if not isinstance(candidate_probe, Mapping) or candidate_probe.get("sha256") != CANDIDATE_ADAPTER_SHA256:
        raise PreparationError("candidate reviewed adapter identity differs")
    candidate_probe_path = _resolve(str(candidate_probe.get("path", "")))
    candidate_digest, candidate_bytes = _sha256(candidate_probe_path)
    if candidate_digest != CANDIDATE_ADAPTER_SHA256 or candidate_bytes != candidate_probe.get("bytes"):
        raise PreparationError("candidate reviewed adapter changed")
    if fixture.get("generator_sha256") != REVIEWED_PROBE_SHA256 or fixture.get("candidate_generator_sha256") != CANDIDATE_ADAPTER_SHA256:
        raise PreparationError("fixture generator identity differs from reviewed probe pins")
    _check_markers(_resolve(str(reviewed["path"])), REQUIRED_PROBE_MARKERS, "reviewed immutable probe")
    _check_markers(candidate_probe_path, REQUIRED_PROBE_MARKERS, "candidate probe adapter")
    _check_markers(_resolve(str(clone["path"])), REQUIRED_CLONE_MARKERS, "clone probe")
    _check_markers(_resolve(str(oracle["path"])), ("fn check_ole_outputs", "allocator_evidence", "owner_target"), "oracle probe")
    api_binding = shared.get("api_binding")
    if not isinstance(api_binding, Mapping) or api_binding.get("policy") != (
        "the baseline one-argument open and candidate explicit-limits open are checked against source text; no implementation overlay"
    ):
        raise PreparationError("source API binding policy is missing")
    for endpoint, root, probe, is_baseline in (
        ("baseline", baseline, _resolve(str(reviewed["path"])), True),
        ("candidate", candidate, candidate_probe_path, False),
    ):
        recorded = api_binding.get(endpoint)
        if not isinstance(recorded, Mapping):
            raise PreparationError(f"source API binding is missing for {endpoint}")
        current_cfb = _cfb_api_binding(_resolve(str(root["root"])), baseline=is_baseline)
        if current_cfb != recorded.get("cfb"):
            raise PreparationError(f"{endpoint} CFB API binding changed")
        current_probe = _probe_api_binding(probe, baseline=is_baseline)
        if current_probe != recorded.get("probe"):
            raise PreparationError(f"{endpoint} probe API binding changed")
    for endpoint in (baseline, candidate):
        driver = endpoint.get("driver")
        if not isinstance(driver, Mapping):
            raise PreparationError("endpoint driver is missing")
        files = _driver_files(driver)
        probe_path = next(path for path in files if path.name == PROBE_PATH.name)
        clone_path = next(path for path in files if path.name == CLONE_PATH.name)
        oracle_path = next(path for path in files if path.name == ORACLE_PATH.name)
        lock_path = next(path for path in files if path.name == "Cargo.lock")
        manifest_path = next(path for path in files if path.name == "Cargo.toml")
        probe_digest, _ = _sha256(probe_path)
        clone_digest, _ = _sha256(clone_path)
        oracle_digest, _ = _sha256(oracle_path)
        lock_digest, _ = _sha256(lock_path)
        if clone_digest != CLONE_PROBE_SHA256 or oracle_digest != ORACLE_SOURCE_SHA256 or lock_digest != LOCK_SHA256:
            raise PreparationError("driver clone, oracle, or lock differs from fixed input")
        expected_probe = REVIEWED_PROBE_SHA256 if endpoint is baseline else CANDIDATE_ADAPTER_SHA256
        if probe_digest != expected_probe:
            raise PreparationError("driver probe does not match its endpoint adapter")
        manifest_text = manifest_path.read_text(encoding="utf-8")
        source_root = _resolve(str(endpoint["root"]))
        if _render_manifest(source_root) != manifest_text:
            raise PreparationError("driver Cargo.toml is not the reviewed source-bound template")
    return document


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--baseline-root", type=Path, required=True)
    parser.add_argument("--candidate-root", type=Path, required=True)
    parser.add_argument("--candidate-ref", required=True)
    parser.add_argument("--reviewed-probe", type=Path, required=True)
    parser.add_argument("--clone-probe", type=Path, required=True)
    parser.add_argument("--oracle-probe", type=Path, required=True)
    parser.add_argument("--lock", type=Path, required=True)
    parser.add_argument("--drivers-root", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args(argv)
    try:
        prepare(
            baseline_root=args.baseline_root,
            candidate_root=args.candidate_root,
            candidate_ref=args.candidate_ref,
            reviewed_probe=args.reviewed_probe,
            clone_probe=args.clone_probe,
            oracle_probe=args.oracle_probe,
            lock=args.lock,
            drivers_root=args.drivers_root,
            output=args.output,
        )
    except (OSError, PreparationError) as error:
        print(f"error: {error}", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
