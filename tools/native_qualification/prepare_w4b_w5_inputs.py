#!/usr/bin/env python3
"""Prepare the historical W4b/W5 baseline source overlays.

This is source preparation only.  It never invokes Cargo, Rust, a benchmark,
or a timing command.  The accepted W4b baseline is a synthetic merge tree
because the record measured the pre-W4b code after combining the W4-XLSX tree
with trunk.  The accepted W5 baseline predates the source-backed probe.  Both
baselines therefore need the exact probe consumers from the candidate tree,
with the recorded W5 route module substituted for the old API.
"""

from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
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


def _adapt_w4b_crud(baseline_crud: str) -> str:
    """Apply only the recorded trait import to the historical CRUD source."""

    import_line = "use litchi_core::sheet::Cell as _;"
    marker = "use litchi::sheet::{WorkbookTrait, Worksheet};"
    if import_line in baseline_crud:
        raise PreparationError("W4b baseline CRUD already contains the Cell import")
    if marker not in baseline_crud:
        raise PreparationError("W4b CRUD consumer marker is missing")
    return baseline_crud.replace(marker, f"{marker}\n{import_line}", 1)


def _prepare_w4b(before: Path, after: Path) -> dict[str, Any]:
    before_state = _source_state(before)
    after_state = _source_state(after)
    if before_state["tree"] != W4B_BASE_TREE:
        raise PreparationError(
            f"W4b baseline tree is {before_state['tree']}, expected recorded merge tree {W4B_BASE_TREE}"
        )
    _require_clean(before_state, "W4b baseline")
    _require_clean(after_state, "W4b candidate")
    probe = _write_exact(after / W4B_PROBE, before / W4B_PROBE)

    crud_source = after / W4B_CRUD
    crud_destination = before / W4B_CRUD
    baseline_crud = crud_destination.read_text(encoding="utf-8")
    # The CRUD harness is intentionally each tree's own source.  The
    # candidate contains the W4b API migration; only the historical base
    # needs the one-line trait import recorded in 0360b §6.
    candidate_hash, candidate_bytes = _sha256(crud_source)
    crud_text = _adapt_w4b_crud(baseline_crud)
    crud = _write_text_exact(crud_destination, crud_text, crud_source)
    crud["baseline_sha256"] = hashlib.sha256(baseline_crud.encode("utf-8")).hexdigest()
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
    marker = '\n#[cfg(all(feature = "doc", feature = "xls", feature = "rtf"))]\nmod routes {'
    start = candidate_probe.rfind(marker)
    if start < 0:
        raise PreparationError("W5 candidate probe has no final routes module")
    main = candidate_probe.find("\nfn main() {", start)
    if main < 0:
        raise PreparationError("W5 candidate probe has no main after routes module")
    return (
        candidate_probe[:start].rstrip()
        + "\n\n"
        + routes_before.strip()
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


def _parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--w4b-before", type=Path, required=True)
    parser.add_argument("--w4b-after", type=Path, required=True)
    parser.add_argument("--w5-before", type=Path, required=True)
    parser.add_argument("--w5-after", type=Path, required=True)
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
            },
            "w5": {
                **w5,
                "analyzer": w5_analyzer,
                "baseline_head": W5_BASE_HEAD,
                "routes_before_sha256": W5_ROUTES_BEFORE_SHA256,
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
