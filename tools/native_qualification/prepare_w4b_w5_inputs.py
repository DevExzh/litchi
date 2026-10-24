#!/usr/bin/env python3
"""Prepare the historical W4b/W5 baseline source overlays.

This is source preparation only.  It never invokes Cargo, Rust, a benchmark,
or a timing command.  The accepted W4b baseline is a synthetic merge tree
because the record measured the pre-W4b code after combining the W4-XLSX tree
with trunk.  The accepted W5 baseline predates the source-backed probe.  Both
baselines therefore need the exact probe consumers from the candidate tree,
with the recorded W4b/W5 compatibility overlays applied for the old APIs.
"""

from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import re
import shutil
import stat
import subprocess
import sys
from typing import Any


W4B_BASE_TREE = "684a028b2ae30a3a110b87bf6743fc598c6691a2"
W5_BASE_HEAD = "7afba5f146aa1a16eba73ebed2144dd041b33fb5"
W5_ROUTES_BEFORE_SHA256 = "237d7242d9fa4db2a3e1589f26b175eff564bb6de680bb217df24f021c698915"
W4B_PROBE = Path("crates/litchi-ops/examples/xlsb_w4b_probe.rs")
W4B_CRUD = Path("tools/perf-baseline/src/bin/xlsb_crud.rs")
W5_PROBE = Path("crates/litchi-ops/examples/legacy_w5_source_probe.rs")
W5_ROUTES = Path("docs/unified-ops/results/0359a/routes_before.rs")
_FULL_COMMIT = re.compile(r"[0-9a-f]{40}")


class PreparationError(RuntimeError):
    """The requested source pair is not the recorded qualification pair."""


def _sha256(path: Path) -> tuple[str, int]:
    if path.is_symlink() or not path.is_file():
        raise PreparationError(f"expected a regular file, not a symlink: {path}")
    digest = hashlib.sha256()
    size = 0
    with path.open("rb") as stream:
        while block := stream.read(1024 * 1024):
            digest.update(block)
            size += len(block)
    return digest.hexdigest(), size


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
    return result.stdout.strip()


def _source_state(root: Path) -> dict[str, Any]:
    root = root.resolve()
    top = Path(_git(root, "rev-parse", "--show-toplevel")).resolve()
    if top != root:
        raise PreparationError(f"source is not a worktree top-level: {root} != {top}")
    status = _git(root, "status", "--porcelain=v1", "--untracked-files=all")
    return {
        "path": str(root),
        "head": _git(root, "rev-parse", "HEAD"),
        "tree": _git(root, "rev-parse", "HEAD^{tree}"),
        "status_sha256": hashlib.sha256(status.encode()).hexdigest(),
        "clean": not bool(status),
    }


def _require_clean(state: dict[str, Any], label: str) -> None:
    if not state["clean"]:
        raise PreparationError(f"{label} source must be clean before overlays: {state['path']}")


def _resolve_no_symlink(path: Path) -> Path:
    """Resolve a path without allowing an existing component to be a link."""

    if not path.is_absolute():
        path = Path.cwd() / path
    probe = Path(path.anchor)
    for part in path.parts[1:]:
        probe /= part
        if probe.is_symlink():
            raise PreparationError(f"path component is a symlink: {path}")
    return path.resolve(strict=False)


def _require_disjoint_paths(paths: dict[str, Path]) -> None:
    items = list(paths.items())
    for index, (left_name, left) in enumerate(items):
        for right_name, right in items[index + 1 :]:
            if left == right or left in right.parents or right in left.parents:
                raise PreparationError(
                    f"paths must be pairwise disjoint: {left_name}={left}, {right_name}={right}"
                )


def _require_disjoint_sources(before: Path, after: Path, output: Path) -> tuple[Path, Path, Path]:
    roots = {"before": _resolve_no_symlink(before), "after": _resolve_no_symlink(after)}
    provenance = _resolve_no_symlink(output)
    _require_disjoint_paths({**roots, "provenance": provenance})
    return roots["before"], roots["after"], provenance


def _write_exact(source: Path, destination: Path, *, mode: int | None = None) -> dict[str, Any]:
    source = _resolve_no_symlink(source)
    destination = _resolve_no_symlink(destination)
    source_hash, source_bytes = _sha256(source)
    if destination.exists() or destination.is_symlink():
        destination_hash, destination_bytes = _sha256(destination)
        if (destination_hash, destination_bytes) != (source_hash, source_bytes):
            raise PreparationError(f"overlay destination differs from its source: {destination}")
    else:
        destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(source, destination)
        if mode is None:
            mode = stat.S_IMODE(source.stat().st_mode)
        destination.chmod(mode)
    return {
        "source": str(source.resolve()),
        "destination": str(destination.resolve()),
        "source_sha256": source_hash,
        "destination_sha256": _sha256(destination)[0],
        "bytes": source_bytes,
    }


def _write_text_exact(destination: Path, content: str, source: Path) -> dict[str, Any]:
    source = _resolve_no_symlink(source)
    destination = _resolve_no_symlink(destination)
    encoded = content.encode("utf-8")
    source_hash, source_bytes = _sha256(source)
    content_hash = hashlib.sha256(encoded).hexdigest()
    if destination.exists() or destination.is_symlink():
        destination_hash, destination_bytes = _sha256(destination)
        if (destination_hash, destination_bytes) != (content_hash, len(encoded)):
            raise PreparationError(f"generated overlay destination differs: {destination}")
    else:
        destination.parent.mkdir(parents=True, exist_ok=True)
        destination.write_bytes(encoded)
    return {
        "source": str(source.resolve()),
        "destination": str(destination.resolve()),
        "source_sha256": source_hash,
        "source_bytes": source_bytes,
        "content_sha256": content_hash,
        "destination_sha256": _sha256(destination)[0],
        "bytes": len(encoded),
    }


def _transform_text_exact(
    destination: Path, expected_before: str, content: str, source: Path
) -> dict[str, Any]:
    """Apply a guarded known transformation to an existing regular file."""

    source = _resolve_no_symlink(source)
    destination = _resolve_no_symlink(destination)
    expected_bytes = expected_before.encode("utf-8")
    encoded = content.encode("utf-8")
    source_hash, source_bytes = _sha256(source)
    before_hash = hashlib.sha256(expected_bytes).hexdigest()
    if destination.is_symlink() or not destination.is_file():
        raise PreparationError(f"expected existing regular transformation target: {destination}")
    actual_before_hash, actual_before_bytes = _sha256(destination)
    if (actual_before_hash, actual_before_bytes) != (before_hash, len(expected_bytes)):
        raise PreparationError(
            f"transformation target changed unexpectedly: {destination}"
        )
    after_hash = hashlib.sha256(encoded).hexdigest()
    destination.write_bytes(encoded)
    actual_after_hash, actual_after_bytes = _sha256(destination)
    if (actual_after_hash, actual_after_bytes) != (after_hash, len(encoded)):
        raise PreparationError(f"transformation target was not written exactly: {destination}")
    return {
        "source": str(source.resolve()),
        "destination": str(destination.resolve()),
        "source_sha256": source_hash,
        "source_bytes": source_bytes,
        "before_sha256": actual_before_hash,
        "before_bytes": actual_before_bytes,
        "after_sha256": actual_after_hash,
        "after_bytes": actual_after_bytes,
        "bytes": actual_after_bytes,
    }


def _adapt_w4b_crud(baseline_crud: str) -> str:
    """Apply only the recorded trait import to the historical CRUD source."""

    import_line = "use litchi_core::sheet::Cell as _;"
    marker = "use litchi::sheet::{WorkbookTrait, Worksheet};"
    if import_line in baseline_crud:
        raise PreparationError("W4b baseline CRUD already contains the Cell import")
    if marker not in baseline_crud:
        raise PreparationError("W4b CRUD consumer marker is missing")
    return baseline_crud.replace(marker, f"{marker}\n{import_line}", 1)


def _adapt_w4b_probe(candidate_probe: str) -> str:
    """Translate the candidate-only calls for the historical W4b API."""

    # The candidate's edit_cell_values returns a lazy Edit and therefore
    # exposes Edit::resolve. The fixed synthetic baseline resolves the
    # selector at the package entry point and returns Result<Option<Edit>>.
    # Keep this overlay exact and fail closed if the candidate probe changes.
    pattern = re.compile(
        r'^(?P<indent> *)\.edit_cell_values\(0_usize\)\n'
        r'(?P=indent)\.expect\("cell edit"\)\n'
        r'(?P=indent)\.resolve\(\)\n'
        r'(?P=indent)\.expect\("worksheet resolution"\)\n'
        r'(?P=indent)\.expect\("worksheet"\);',
        re.MULTILINE,
    )
    translated, replacements = pattern.subn(
        lambda match: (
            f'{match.group("indent")}.edit_cell_values(0_usize)\n'
            f'{match.group("indent")}.expect("cell edit")\n'
            f'{match.group("indent")}.expect("worksheet");'
        ),
        candidate_probe,
    )
    if candidate_probe.count(".resolve()") != 3 or replacements != 3:
        raise PreparationError(
            "W4b probe must contain exactly three historical selector-resolution calls"
        )
    if ".resolve()" in translated:
        raise PreparationError("W4b baseline probe retained a candidate-only resolve call")
    if candidate_probe.count("commit.into_file()") != 1:
        raise PreparationError(
            "W4b probe must contain exactly one candidate commit into_file call"
        )
    translated = translated.replace(
        "black_box(commit.into_file());",
        "black_box(commit.file);",
        1,
    )
    if "commit.into_file()" in translated:
        raise PreparationError("W4b baseline probe retained a candidate-only commit accessor")
    return translated


def _prepare_w4b(before: Path, after: Path) -> dict[str, Any]:
    before_state = _source_state(before)
    after_state = _source_state(after)
    if before_state["tree"] != W4B_BASE_TREE:
        raise PreparationError(
            f"W4b baseline tree is {before_state['tree']}, expected recorded merge tree {W4B_BASE_TREE}"
        )
    _require_clean(before_state, "W4b baseline")
    _require_clean(after_state, "W4b candidate")
    probe_source = (after / W4B_PROBE).read_text(encoding="utf-8")
    probe_text = _adapt_w4b_probe(probe_source)
    probe = _write_text_exact(before / W4B_PROBE, probe_text, after / W4B_PROBE)
    probe["candidate_sha256"] = _sha256(after / W4B_PROBE)[0]
    probe["candidate_bytes"] = len(probe_source.encode("utf-8"))
    probe["adaptation"] = (
        "recorded three-call historical selector-resolution and commit-field translations"
    )

    crud_source = after / W4B_CRUD
    crud_destination = before / W4B_CRUD
    baseline_crud = crud_destination.read_text(encoding="utf-8")
    # The CRUD harness is intentionally each tree's own source.  The
    # candidate contains the W4b API migration; only the historical base
    # needs the one-line trait import recorded in 0360b §6.
    candidate_hash, candidate_bytes = _sha256(crud_source)
    crud_text = _adapt_w4b_crud(baseline_crud)
    crud = _transform_text_exact(crud_destination, baseline_crud, crud_text, crud_source)
    crud["candidate_sha256"] = candidate_hash
    crud["candidate_bytes"] = candidate_bytes
    crud["adaptation"] = "recorded one-line litchi_core::sheet::Cell import"
    return {
        "baseline_before_overlay": before_state,
        "baseline": _source_state(before),
        "candidate": after_state,
        "overlays": [probe, crud],
    }


def _adapt_w5_probe(candidate_probe: str, routes_before: str) -> str:
    """Translate the recorded W5 probe onto the historical public APIs.

    The source-backed route module is intentionally replaced by the exact
    0359a ``routes_before.rs`` module.  The follow-up DOC publication controls
    were added after that module was recorded, though, so retain those two
    route bodies from the candidate and append them to the historical module.
    They use the historical DOC package API unchanged.  The only common
    consumer translation is the old ``File::edit_with_limits`` entry point and
    its public ``Commit::file`` field.
    """
    marker = '\n#[cfg(all(feature = "doc", feature = "xls", feature = "rtf"))]\nmod routes {'
    start = candidate_probe.rfind(marker)
    if start < 0:
        raise PreparationError("W5 candidate probe has no final routes module")
    main = candidate_probe.find("\nfn main() {", start)
    if main < 0:
        raise PreparationError("W5 candidate probe has no main after routes module")

    # These two functions are the DOC publication controls introduced after
    # the pinned 0359 tip.  Extract only the known route span and fail closed
    # if the candidate changes its shape or count; silently dropping a route
    # would invalidate the matched workload.
    leaf_start_marker = "\n    /// Run the 1k leaf body-toggle workload through complete DOC publication."
    leaf_end_marker = "\n    fn workbook("
    leaf_start = candidate_probe.find(leaf_start_marker, start, main)
    leaf_end = candidate_probe.find(leaf_end_marker, leaf_start, main)
    if leaf_start < 0 or leaf_end < 0 or leaf_end <= leaf_start:
        raise PreparationError("W5 candidate probe has no complete DOC publication controls")
    leaf_functions = candidate_probe[leaf_start + 1 : leaf_end].rstrip()
    expected_leaf_functions = (
        "pub(super) fn doc_batch_leaf_write",
        "pub(super) fn doc_leaf_write",
    )
    for function in expected_leaf_functions:
        if leaf_functions.count(function) != 1:
            raise PreparationError(f"W5 probe must contain exactly one {function} route")
    if "pub(super) fn doc_" not in leaf_functions or leaf_functions.count("pub(super) fn doc_") != 2:
        raise PreparationError("W5 probe has an unexpected DOC route in the compatibility span")
    if routes_before.count("pub(super) fn doc_batch_leaf_write") != 0:
        raise PreparationError("W5 historical routes unexpectedly contain the DOC batch route")
    if routes_before.count("pub(super) fn doc_leaf_write") != 0:
        raise PreparationError("W5 historical routes unexpectedly contain the DOC leaf route")
    historical_routes = routes_before.rstrip()
    if not historical_routes.endswith("}"):
        raise PreparationError("W5 historical routes module has no closing brace")
    historical_routes = (
        historical_routes[:-1].rstrip()
        + "\n\n"
        + leaf_functions
        + "\n}"
    )

    translated = candidate_probe[:start].rstrip()
    options_pattern = re.compile(
        r"^(?P<indent> +)\.edit_with_options\(\n"
        r" +litchi_ops::EditOptions::new\(\)\.semantic_limits\(limits\),$",
        re.MULTILINE,
    )
    translated, options_replacements = options_pattern.subn(
        lambda match: f'{match.group("indent")}.edit_with_limits(&limits,',
        translated,
    )
    if options_replacements != 1:
        raise PreparationError(
            "W5 probe must contain exactly one unified File options entry point"
        )
    if translated.count("commit.file()") != 1:
        raise PreparationError("W5 probe must contain exactly one unified Commit::file accessor")
    translated = translated.replace("black_box(commit.file());", "black_box(commit.file);", 1)
    if "litchi_ops::EditOptions" in translated or "commit.file()" in translated:
        raise PreparationError("W5 baseline probe retained a candidate-only ops API")
    if translated.count("`File::edit_with_options`") != 1:
        raise PreparationError("W5 probe API documentation marker changed unexpectedly")
    translated = translated.replace("`File::edit_with_options`", "`File::edit_with_limits`", 1)

    return (
        translated
        + "\n\n"
        + historical_routes
        + "\n\n"
        + candidate_probe[main:].lstrip()
        + "\n"
    )


def _prepare_w5(before: Path, after: Path) -> dict[str, Any]:
    before_state = _source_state(before)
    after_state = _source_state(after)
    if before_state["head"] != W5_BASE_HEAD:
        raise PreparationError(
            f"W5 baseline HEAD is {before_state['head']}, expected recorded 0359 tip {W5_BASE_HEAD}"
        )
    _require_clean(before_state, "W5 baseline")
    _require_clean(after_state, "W5 candidate")
    candidate_probe = (after / W5_PROBE).read_text(encoding="utf-8")
    routes_path = after / W5_ROUTES
    routes_hash, _ = _sha256(routes_path)
    if routes_hash != W5_ROUTES_BEFORE_SHA256:
        raise PreparationError(
            f"W5 routes_before hash is {routes_hash}, expected {W5_ROUTES_BEFORE_SHA256}"
        )
    routes = routes_path.read_text(encoding="utf-8")
    adapted = _adapt_w5_probe(candidate_probe, routes)
    adapted_path = before / W5_PROBE
    generated = _write_text_exact(adapted_path, adapted, after / W5_PROBE)
    generated["candidate_probe_sha256"] = _sha256(after / W5_PROBE)[0]
    generated["routes_before_sha256"] = routes_hash
    routes_overlay = _write_exact(routes_path, before / W5_ROUTES)
    return {
        "baseline_before_overlay": before_state,
        "baseline": _source_state(before),
        "candidate": after_state,
        "overlays": [generated, routes_overlay],
    }


def _analyzer(path: Path) -> dict[str, Any]:
    digest, size = _sha256(path)
    return {"path": str(path.resolve()), "sha256": digest, "bytes": size}


def _require_candidate_ref(root: Path, state: dict[str, Any], expected: str, label: str) -> str:
    if not _FULL_COMMIT.fullmatch(expected):
        raise PreparationError(f"{label} candidate ref must be a full 40-digit commit SHA")
    resolved = _git(root, "rev-parse", "--verify", f"{expected}^{{commit}}")
    if state["head"] != resolved:
        raise PreparationError(
            f"{label} candidate HEAD is {state['head']}, expected selected ref {resolved}"
        )
    return resolved


def _parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--w4b-before", type=Path, required=True)
    parser.add_argument("--w4b-after", type=Path, required=True)
    parser.add_argument("--w4b-after-ref", required=True, help="full selected W4b candidate commit SHA")
    parser.add_argument("--w5-before", type=Path, required=True)
    parser.add_argument("--w5-after", type=Path, required=True)
    parser.add_argument("--w5-after-ref", required=True, help="full selected W5 candidate commit SHA")
    parser.add_argument("--output", type=Path, required=True, help="external provenance JSON path")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = _parse_args(argv)
    try:
        w4b_before, w4b_after, output = _require_disjoint_sources(
            args.w4b_before, args.w4b_after, args.output
        )
        w5_before, w5_after, _ = _require_disjoint_sources(
            args.w5_before, args.w5_after, args.output
        )
        _require_disjoint_paths(
            {
                "w4b-before": w4b_before,
                "w4b-after": w4b_after,
                "w5-before": w5_before,
                "w5-after": w5_after,
                "provenance": output,
            }
        )
        if output.is_symlink() or output.exists():
            raise PreparationError(f"refusing to overwrite or follow provenance output: {output}")
        w4b_after_state = _source_state(w4b_after)
        w5_after_state = _source_state(w5_after)
        w4b_candidate_ref = _require_candidate_ref(
            w4b_after, w4b_after_state, args.w4b_after_ref, "W4b"
        )
        w5_candidate_ref = _require_candidate_ref(
            w5_after, w5_after_state, args.w5_after_ref, "W5"
        )
        # Validate all immutable inputs before touching either baseline.  This
        # keeps a typo in an analyzer path from leaving a half-prepared pair.
        w4b_analyzer = _analyzer(w4b_after / "docs/unified-ops/results/0360b/analyze.py")
        w5_analyzer = _analyzer(w5_after / "docs/unified-ops/results/0359a/analyze.py")
        w4b = _prepare_w4b(w4b_before, w4b_after)
        w5 = _prepare_w5(w5_before, w5_after)
        provenance = {
            "schema": "litchi.native-qualification.source-preparation.v1",
            "w4b": {
                **w4b,
                "analyzer": w4b_analyzer,
                "baseline_tree": W4B_BASE_TREE,
                "selected_candidate_ref": w4b_candidate_ref,
            },
            "w5": {
                **w5,
                "analyzer": w5_analyzer,
                "baseline_head": W5_BASE_HEAD,
                "routes_before_sha256": W5_ROUTES_BEFORE_SHA256,
                "selected_candidate_ref": w5_candidate_ref,
            },
        }
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(json.dumps(provenance, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    except (OSError, PreparationError) as error:
        print(f"source preparation refused: {error}", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
