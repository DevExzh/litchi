#!/usr/bin/env python3
"""Prepare the external PPT lazy-native endpoint overlays.

The historical parent predates the PPT probe and its default fixture.  This
command materializes the reviewed probe closure in a temporary baseline
worktree, including the fixture consumed by ``include_bytes!``.  It never
changes a tracked production tree and never invokes Cargo or a probe.  The
receipt is the binding between the reviewed candidate bytes and the temporary
baseline endpoint; the build receipt must retain it.
"""

from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import re
import shutil
import subprocess
from typing import Any


SCHEMA = "litchi.ppt.lazy-native-qualification.preparation.v1"
BASELINE_REF = "8779d059add6c90c74552801372c44f770db9f0f"
BASELINE_TREE = "27e0b56e2f8e3ab11b00c119604b5636f358dfe2"
BASELINE_LOCK_PATH = Path("tools/perf-baseline/Cargo.lock")
BASELINE_LOCK_BYTES = 57_158
BASELINE_LOCK_SHA256 = "4def018e9c1d5f409b8c60d6cc0b8f114c486ab3f8fed2c7a4e9c7baf21c5184"
FIXTURE_PATH = Path("test-data/unified-ops/ppt-lazy-native-fixture.ppt")
FIXTURE_BYTES = 14_848
FIXTURE_SHA256 = "02fa44ca76903943e80a3a2da1f9dbd326c7a4a047540c3aa9cf211e34ef617c"
# The historical standalone perf-baseline crate compiles its complete library
# even for the PPT probe bin.  These are the other tracked include_bytes!
# inputs in that library/probe closure.  They are copied only into the
# external baseline adapter; no production tree is changed.
COMPILE_FIXTURES = (
    (
        Path("test-data/poi/test-data/slideshow/ppt_with_embeded.ppt"),
        180_736,
        "afa1dd5adc36765e6e40725a133b0330e5ad1febd1506ca93ceae8c5c5400b53",
    ),
    (
        Path("test-data/unified-ops/0294/video.ppt"),
        43_520,
        "f32f0614958a177a5195cd7537cdfd2efef52c57de4decc13cf65eadc66cdb8a",
    ),
    (
        Path("test-data/poi/test-data/slideshow/bug61881.ppt"),
        74_752,
        "ba2bf551e87f5fd7667dec52ac19005396bdf9af00f0006f9a247a9fede2e8ee",
    ),
    (
        Path("test-data/poi/test-data/spreadsheet/54016.xls"),
        984_576,
        "2e050f1fbb31868b097aa6d4d0fe0a16af8e39c252af82d01cd8f93c4f9a911a",
    ),
    (
        Path("test-data/rtf/watermark.rtf"),
        69_471,
        "48d62dcd959e737b06ebb8255780bcaaf1e88056ff9c3d5a21d3ff5cd3ddf9cb",
    ),
)
PROBE_PATHS = (
    Path("tools/perf-baseline/src/ppt_lazy_native.rs"),
    Path("tools/perf-baseline/src/bin/ppt_lazy_native.rs"),
    Path("tools/perf-baseline/src/bin/ppt_lazy_native_alloc.rs"),
)
REGISTRATION_PATH = Path("tools/perf-baseline/src/lib.rs")
REGISTRATION_LINE = "pub mod ppt_lazy_native;"
REGISTRATION_ANCHOR = "pub mod pptx_retention;"
FULL_SHA = re.compile(r"^[0-9a-f]{40}$")
HEX64 = re.compile(r"^[0-9a-f]{64}$")
UPSTREAM_COMPONENT = "3rd" + "party"


class PreparationError(RuntimeError):
    """The requested endpoint pair is not the reviewed pair."""


def _resolve(path: Path, *, must_exist: bool = False) -> Path:
    path = path.absolute()
    if UPSTREAM_COMPONENT in path.parts:
        raise PreparationError(f"untracked upstream checkout is forbidden: {path}")
    current = Path(path.anchor)
    for part in path.parts[1:]:
        current /= part
        if current.is_symlink():
            raise PreparationError(f"symlink path component is forbidden: {current}")
    if must_exist and not path.exists():
        raise PreparationError(f"path does not exist: {path}")
    return path.resolve(strict=False)


def _under(path: Path, parent: Path) -> bool:
    try:
        path.relative_to(parent)
    except ValueError:
        return False
    return True


def _sha256(path: Path) -> tuple[str, int]:
    path = _resolve(path, must_exist=True)
    if path.is_symlink() or not path.is_file():
        raise PreparationError(f"expected a regular file: {path}")
    digest = hashlib.sha256()
    size = 0
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
            size += len(block)
    return digest.hexdigest(), size


def _identity(path: Path) -> dict[str, Any]:
    sha256, bytes_ = _sha256(path)
    return {"path": str(_resolve(path)), "bytes": bytes_, "sha256": sha256}


def _git(root: Path, *args: str) -> str:
    result = subprocess.run(
        ["git", "-C", str(root), *args],
        check=False,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        timeout=30,
    )
    if result.returncode:
        raise PreparationError(
            f"git {' '.join(args)} failed in {root}: {result.stderr.strip()}"
        )
    return result.stdout.strip()


def _git_blob_identity(root: Path, relative: Path) -> dict[str, Any]:
    result = subprocess.run(
        ["git", "-C", str(root), "show", f"HEAD:{relative}"],
        check=False,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
    )
    if result.returncode:
        raise PreparationError(
            f"baseline tracked fixture {relative} cannot be read from HEAD: "
            f"{result.stderr.decode(errors='replace').strip()}"
        )
    return {
        "path": str(root / relative),
        "bytes": len(result.stdout),
        "sha256": hashlib.sha256(result.stdout).hexdigest(),
    }


def _materialize_tracked_baseline_fixture(
    root: Path,
    relative: Path,
    expected_bytes: int,
    expected_sha256: str,
) -> bool:
    """Restore one reviewed tracked blob omitted by a sparse external checkout."""

    destination = _resolve(root / relative)
    records = _git(root, "ls-files", "-v", "-z", "--", str(relative))
    tracked = records in (f"H {relative}\0", f"S {relative}\0")
    if not tracked:
        return False
    baseline_identity = _git_blob_identity(root, relative)
    if baseline_identity["bytes"] != expected_bytes or baseline_identity["sha256"] != expected_sha256:
        raise PreparationError(f"baseline tracked fixture differs from the reviewed closure: {relative}")
    if destination.exists() or destination.is_symlink():
        if destination.is_symlink() or not destination.is_file():
            raise PreparationError(f"baseline tracked fixture is not a regular file: {destination}")
        actual = _identity(destination)
        if actual["bytes"] != expected_bytes or actual["sha256"] != expected_sha256:
            raise PreparationError(f"baseline tracked fixture bytes differ: {relative}")
        return False
    result = subprocess.run(
        [
            "git",
            "-C",
            str(root),
            "checkout-index",
            "--force",
            "--ignore-skip-worktree-bits",
            "--",
            str(relative),
        ],
        check=False,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )
    if result.returncode:
        raise PreparationError(
            f"could not materialize baseline tracked fixture {relative}: {result.stderr.strip()}"
        )
    actual = _identity(destination)
    if actual["bytes"] != expected_bytes or actual["sha256"] != expected_sha256:
        raise PreparationError(f"materialized baseline fixture differs: {relative}")
    return True


def _tracked_baseline_path(root: Path, relative: Path) -> bool:
    records = _git(root, "ls-files", "-v", "-z", "--", str(relative))
    return records in (f"H {relative}\0", f"S {relative}\0")


def _validate_initial_baseline_state(
    state: dict[str, Any],
    *,
    allowed_sparse_paths: set[str],
) -> None:
    """Accept only the tracked fixture omissions caused by sparse checkout."""

    for line in state["status"].splitlines():
        if len(line) < 3:
            raise PreparationError(f"malformed baseline status line: {line!r}")
        paths = _status_paths(line)
        if len(paths) != 1 or paths[0] not in allowed_sparse_paths or line[:2] not in {" D", "D "}:
            raise PreparationError(
                "baseline has unexpected changes before fixture materialization: "
                f"{line!r}"
            )


def _source_state(root: Path) -> dict[str, Any]:
    root = _resolve(root, must_exist=True)
    top = _resolve(Path(_git(root, "rev-parse", "--show-toplevel")))
    if top != root:
        raise PreparationError(f"source is not a worktree top-level: {root}")
    status = _git(root, "status", "--porcelain=v1", "--untracked-files=all")
    return {
        "root": str(root),
        "ref": _git(root, "rev-parse", "HEAD"),
        "tree": _git(root, "rev-parse", "HEAD^{tree}"),
        "status": status,
        "status_sha256": hashlib.sha256(status.encode()).hexdigest(),
        "clean": not bool(status),
    }


def _production_tree_sha256(root: Path) -> str:
    return hashlib.sha256(
        _git(root, "ls-tree", "-r", "--full-tree", "HEAD").encode()
    ).hexdigest()


def _status_paths(status: str) -> list[str]:
    paths: list[str] = []
    for line in status.splitlines():
        if len(line) < 4:
            raise PreparationError(f"malformed git status line: {line!r}")
        # ``_git`` strips the outer status stream.  If the first entry was a
        # worktree-only modification, that removes its leading status-column
        # space; recognise both the normal three-column form and the compact
        # ``M path`` form rather than silently dropping the first path byte.
        if len(line) >= 3 and line[1] == " ":
            # Normal two-column status such as ``D  path`` has a second
            # separator before the path; _git's outer strip turns `` M
            # path`` into the compact ``M path`` form.
            path = line[3:] if line[2] == " " else line[2:]
        else:
            path = line[3:]
        if " -> " in path:
            path = path.split(" -> ", 1)[1]
        paths.append(path)
    return paths


def _tracked_regular(root: Path, relative: Path, label: str) -> Path:
    path = _resolve(root / relative, must_exist=True)
    if not _under(path, root) or path.is_symlink() or not path.is_file():
        raise PreparationError(f"{label} is not a regular file below its root")
    records = _git(root, "ls-files", "-v", "-z", "--", str(relative))
    if records != f"H {relative}\0":
        raise PreparationError(f"{label} is not an ordinary tracked file")
    return path


def _copy_new(source: Path, destination: Path, label: str) -> dict[str, Any]:
    source = _resolve(source, must_exist=True)
    destination = _resolve(destination)
    if destination.exists() or destination.is_symlink():
        raise PreparationError(f"{label} destination must be absent: {destination}")
    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(source, destination)
    source_identity = _identity(source)
    destination_identity = _identity(destination)
    if source_identity["bytes"] != destination_identity["bytes"] or source_identity["sha256"] != destination_identity["sha256"]:
        raise PreparationError(f"{label} was not copied byte-for-byte")
    return {
        "source": source_identity,
        "destination": destination_identity,
    }


def _reviewed_fixture_specs() -> tuple[tuple[Path, int, str], ...]:
    return ((FIXTURE_PATH, FIXTURE_BYTES, FIXTURE_SHA256), *COMPILE_FIXTURES)


def _add_registration(root: Path, candidate: Path) -> dict[str, Any]:
    destination = _resolve(root / REGISTRATION_PATH, must_exist=True)
    candidate_registration = _tracked_regular(root=candidate, relative=REGISTRATION_PATH, label="candidate registration")
    candidate_text = candidate_registration.read_text(encoding="utf-8")
    if candidate_text.count(REGISTRATION_LINE) != 1:
        raise PreparationError("candidate registration must contain exactly one PPT module line")
    if destination.is_symlink() or not destination.is_file():
        raise PreparationError("baseline registration destination is not a regular file")
    before = destination.read_bytes()
    if REGISTRATION_LINE in before.decode("utf-8"):
        raise PreparationError("baseline registration already contains the PPT module")
    text = before.decode("utf-8")
    anchor_count = text.count(REGISTRATION_ANCHOR)
    if anchor_count != 1:
        raise PreparationError("baseline registration anchor is not unique")
    line_ending = "\r\n" if "\r\n" in text else "\n"
    insertion = REGISTRATION_ANCHOR + line_ending + REGISTRATION_LINE
    after_text = text.replace(REGISTRATION_ANCHOR, insertion, 1)
    destination.write_text(after_text, encoding="utf-8", newline="")
    before_sha = hashlib.sha256(before).hexdigest()
    after_sha, after_bytes = _sha256(destination)
    return {
        "path": str(REGISTRATION_PATH),
        "line": REGISTRATION_LINE,
        "before": {"bytes": len(before), "sha256": before_sha},
        "after": {"bytes": after_bytes, "sha256": after_sha},
        "candidate": _identity(candidate_registration),
    }


def prepare(
    *,
    baseline_root: Path,
    candidate_root: Path,
    candidate_ref: str,
    candidate_lock_sha256: str,
    output: Path,
) -> dict[str, Any]:
    """Materialize the exact external baseline adapter and write a receipt."""

    if FULL_SHA.fullmatch(candidate_ref) is None:
        raise PreparationError("candidate_ref must be a full lowercase commit SHA")
    if HEX64.fullmatch(candidate_lock_sha256) is None:
        raise PreparationError("candidate_lock_sha256 must be a lowercase SHA-256")
    baseline_root = _resolve(baseline_root, must_exist=True)
    candidate_root = _resolve(candidate_root, must_exist=True)
    output = _resolve(output)
    if baseline_root == candidate_root or _under(baseline_root, candidate_root) or _under(candidate_root, baseline_root):
        raise PreparationError("baseline and candidate worktrees must be disjoint")
    if output.exists() or _under(output, baseline_root) or _under(output, candidate_root):
        raise PreparationError("preparation receipt must be a new file outside both worktrees")

    # Validate the endpoint identity before touching a sparse worktree.  A
    # wrong endpoint must fail closed rather than being made to look correct
    # by restoring a reviewed file into it.
    baseline_initial = _source_state(baseline_root)
    if baseline_initial["ref"] != BASELINE_REF or baseline_initial["tree"] != BASELINE_TREE:
        raise PreparationError("baseline is not the pinned pre-lazy parent")
    sparse_paths: set[str] = set()
    for relative, _, _ in COMPILE_FIXTURES:
        tracked = _tracked_baseline_path(baseline_root, relative)
        destination = _resolve(baseline_root / relative)
        if tracked:
            sparse_paths.add(str(relative))
        elif destination.exists() or destination.is_symlink():
            raise PreparationError(
                f"baseline compile fixture exists outside the pinned tracked tree: {relative}"
            )
    _validate_initial_baseline_state(
        baseline_initial,
        allowed_sparse_paths=sparse_paths,
    )
    candidate = _source_state(candidate_root)
    if candidate["ref"] != candidate_ref or not candidate["clean"]:
        raise PreparationError("candidate must be clean and at candidate_ref")

    # A detached baseline checkout may have a sparse/omitted worktree file for
    # a blob that is nevertheless part of its pinned Git tree.  Materialize
    # only the reviewed compile fixtures after identity validation;
    # checkout-index restores the exact pinned blob and therefore leaves the
    # production tree unchanged.
    for relative, expected_bytes, expected_sha256 in COMPILE_FIXTURES:
        _materialize_tracked_baseline_fixture(
            baseline_root, relative, expected_bytes, expected_sha256
        )

    baseline_before = _source_state(baseline_root)
    if not baseline_before["clean"]:
        raise PreparationError("baseline must be clean before the external overlay")

    baseline_lock = _tracked_regular(baseline_root, BASELINE_LOCK_PATH, "baseline Cargo.lock")
    baseline_lock_identity = _identity(baseline_lock)
    if baseline_lock_identity["bytes"] != BASELINE_LOCK_BYTES or baseline_lock_identity["sha256"] != BASELINE_LOCK_SHA256:
        raise PreparationError("baseline Cargo.lock differs from the pinned standalone lock")
    candidate_lock = _tracked_regular(candidate_root, BASELINE_LOCK_PATH, "candidate Cargo.lock")
    candidate_lock_identity = _identity(candidate_lock)
    if candidate_lock_identity["sha256"] != candidate_lock_sha256:
        raise PreparationError("candidate Cargo.lock differs from its supplied identity")

    overlays = []
    for relative in PROBE_PATHS:
        source = _tracked_regular(candidate_root, relative, f"candidate probe {relative}")
        destination = baseline_root / relative
        overlays.append(_copy_new(source, destination, f"baseline probe {relative}"))
    registration = _add_registration(baseline_root, candidate_root)
    fixture_overlays: dict[str, dict[str, Any]] = {}
    fixture_identities: dict[str, dict[str, Any]] = {}
    for relative, expected_bytes, expected_sha256 in _reviewed_fixture_specs():
        source = _tracked_regular(candidate_root, relative, f"reviewed PPT fixture {relative}")
        identity = _identity(source)
        if identity["bytes"] != expected_bytes or identity["sha256"] != expected_sha256:
            raise PreparationError(f"candidate reviewed fixture differs: {relative}")
        destination = baseline_root / relative
        if destination.exists() or destination.is_symlink():
            destination_identity = _identity(destination)
            if destination_identity["bytes"] != expected_bytes or destination_identity["sha256"] != expected_sha256:
                raise PreparationError(f"baseline reviewed fixture differs: {relative}")
            overlay = {
                "source": identity,
                "destination": destination_identity,
            }
        else:
            overlay = _copy_new(source, destination, f"baseline reviewed fixture {relative}")
        key = str(relative)
        fixture_overlays[key] = overlay
        fixture_identities[key] = identity

    baseline_after = _source_state(baseline_root)
    expected_overlay_paths = sorted(
        [str(path) for path in PROBE_PATHS]
        + [str(REGISTRATION_PATH)]
        + [str(path) for path, _, _ in _reviewed_fixture_specs()]
    )
    normal_status = baseline_after["status"]
    observed_status = _status_paths(normal_status)
    normal_paths = set(observed_status)
    extra_status_lines: list[str] = []
    # Some reviewed fixtures live below repository ignore rules because the
    # historical baseline never tracked that test corpus.  Query only those
    # exact paths with ``--ignored``; a repository-wide ignored scan would
    # traverse the large historical fixture corpus and make preparation
    # needlessly unbounded.
    for relative, _, _ in COMPILE_FIXTURES:
        ignored_status = _git(
            baseline_root,
            "status",
            "--porcelain=v1",
            "--untracked-files=all",
            "--ignored",
            "--",
            str(relative),
        )
        for line in ignored_status.splitlines():
            paths = _status_paths(line)
            if paths and paths[0] not in normal_paths:
                extra_status_lines.append(line)
                observed_status.extend(paths)
    if extra_status_lines:
        baseline_after["status"] = "\n".join([normal_status, *extra_status_lines])
        baseline_after["status_sha256"] = hashlib.sha256(
            baseline_after["status"].encode()
        ).hexdigest()
    observed_status = sorted(set(observed_status))
    expected_status = sorted(
        {
            *[str(path) for path in PROBE_PATHS],
            str(REGISTRATION_PATH),
            str(FIXTURE_PATH),
            *[
                str(path)
                for path, _, _ in COMPILE_FIXTURES
                if str(path) in observed_status
            ],
        }
    )
    if observed_status != expected_status:
        raise PreparationError(f"baseline overlay changed unexpected paths: {observed_status!r}")
    if baseline_after["tree"] != baseline_before["tree"]:
        raise PreparationError("baseline tracked production tree changed during preparation")
    output.parent.mkdir(parents=True, exist_ok=True)
    receipt = {
        "schema": SCHEMA,
        "status": "prepared",
        "lane": "ppt-lazy-native",
        "baseline": {
            "sparse_before": baseline_initial,
            "before": baseline_before,
            "after_overlay": baseline_after,
            "ref": baseline_before["ref"],
            "tree": baseline_before["tree"],
            "production_tree_sha256": _production_tree_sha256(baseline_root),
            "lock": baseline_lock_identity,
            "overlay_paths": expected_overlay_paths,
            "probes": overlays,
            "registration": registration,
            "fixture": fixture_overlays[str(FIXTURE_PATH)],
            "compile_fixtures": [
                fixture_overlays[str(path)] for path, _, _ in COMPILE_FIXTURES
            ],
        },
        "candidate": {
            "state": candidate,
            "ref": candidate["ref"],
            "tree": candidate["tree"],
            "lock": candidate_lock_identity,
            "fixture": fixture_identities[str(FIXTURE_PATH)],
            "compile_fixtures": [
                fixture_identities[str(path)] for path, _, _ in COMPILE_FIXTURES
            ],
            "probes": [_identity(_tracked_regular(candidate_root, path, f"candidate probe {path}")) for path in PROBE_PATHS],
        },
        "claim_scope": "source preparation identity only; no build, correctness, timing, or acceptance claim",
    }
    output.write_text(json.dumps(receipt, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    return receipt


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--baseline-root", type=Path, required=True)
    parser.add_argument("--candidate-root", type=Path, required=True)
    parser.add_argument("--candidate-ref", required=True)
    parser.add_argument("--candidate-lock-sha256", required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args(argv)
    try:
        prepare(
            baseline_root=args.baseline_root,
            candidate_root=args.candidate_root,
            candidate_ref=args.candidate_ref,
            candidate_lock_sha256=args.candidate_lock_sha256,
            output=args.output,
        )
    except (OSError, PreparationError) as error:
        print(f"error: {error}")
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
