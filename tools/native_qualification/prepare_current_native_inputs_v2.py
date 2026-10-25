#!/usr/bin/env python3
"""Record source-only inputs for the versioned current-native policy.

This preparer intentionally has no ``--candidate-lock`` or ``--expected``
option.  It records observations for review while the candidate policy is
pending, but marks the receipt blocked; a build or qualification consumer must
use :mod:`current_native_policy_v2` and will reject that receipt.
"""

from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import sys
from typing import Any

from native_qualification.current_native_policy_v2 import (
    EXECUTION_RECEIPT_RELATIVE,
    POLICY_SCHEMA,
    PREPARATION_SCHEMA,
    PolicyError,
    TRUSTED_POLICY,
    observe_source,
    validate_routes,
)


class PreparationError(ValueError):
    """The source-only preparation contract was not satisfied."""


def _canonical(value: Any) -> bytes:
    return (json.dumps(value, sort_keys=True, separators=(",", ":"), ensure_ascii=True) + "\n").encode("utf-8")


def _load_canonical(path: Path, label: str) -> tuple[dict[str, Any], str, int]:
    if path.is_symlink() or not path.is_file():
        raise PreparationError(f"{label} is not a regular file: {path}")
    try:
        raw = path.read_bytes()
        value = json.loads(raw.decode("utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise PreparationError(f"cannot read {label}: {error}") from error
    if not isinstance(value, dict):
        raise PreparationError(f"{label} must contain a JSON object")
    if _canonical(value) != raw:
        raise PreparationError(f"{label} is not canonical JSON")
    return value, hashlib.sha256(raw).hexdigest(), len(raw)


def _external_output(output: Path, roots: tuple[Path, ...]) -> Path:
    lexical = output.absolute()
    cursor = Path(lexical.anchor)
    for component in lexical.parts[1:]:
        cursor /= component
        if cursor.is_symlink():
            raise PreparationError(f"preparation receipt path contains a symlink component: {cursor}")
    if lexical.exists() or lexical.is_symlink():
        raise PreparationError(f"refusing to overwrite preparation receipt: {output}")
    resolved = lexical.parent.resolve() / lexical.name
    for root in roots:
        root = root.resolve()
        if resolved == root or root in resolved.parents:
            raise PreparationError(f"preparation receipt must be outside source roots: {output}")
    resolved.parent.mkdir(parents=True, exist_ok=True)
    return resolved


def _tooling_identity(path: Path, label: str) -> dict[str, Any]:
    lexical = path.absolute()
    cursor = Path(lexical.anchor)
    for component in lexical.parts[1:]:
        cursor /= component
        if cursor.is_symlink():
            raise PreparationError(f"{label} contains a symlink component: {cursor}")
    if path.is_symlink() or not path.is_file():
        raise PreparationError(f"{label} is not a regular file: {path}")
    payload = path.read_bytes()
    return {"path": str(path.resolve()), "sha256": hashlib.sha256(payload).hexdigest(), "bytes": len(payload)}


def _disjoint(paths: list[tuple[str, Path]]) -> None:
    resolved = [(label, path.absolute().resolve()) for label, path in paths]
    for index, (left_label, left) in enumerate(resolved):
        for right_label, right in resolved[index + 1 :]:
            if left == right or left in right.parents or right in left.parents:
                raise PreparationError(f"paths must be pairwise disjoint: {left_label}={left}, {right_label}={right}")


def _create_fresh_output_root(path: Path | None) -> dict[str, Any] | None:
    if path is None:
        return None
    lexical = path.absolute()
    cursor = Path(lexical.anchor)
    for component in lexical.parts[1:]:
        cursor /= component
        if cursor.is_symlink():
            raise PreparationError(f"qualification output root contains a symlink component: {cursor}")
    if lexical.exists() or lexical.is_symlink():
        raise PreparationError(f"qualification output root must be fresh and absent: {path}")
    parent = lexical.parent
    if parent.is_symlink() or not parent.is_dir():
        raise PreparationError(f"qualification output root parent is not a regular directory: {parent}")
    lexical.mkdir()
    return {
        "path": str(lexical.resolve()),
        "entries": [],
        "sha256": hashlib.sha256(_canonical([])).hexdigest(),
    }


def prepare(
    *,
    lane: str,
    manifest: Path,
    before: Path,
    after: Path,
    output: Path,
    before_archive: Path | None = None,
    after_archive: Path | None = None,
    output_root: Path | None = None,
) -> dict[str, Any]:
    """Write a blocked or ready source receipt without invoking Cargo.

    The candidate observations are useful to the reviewer while the exact
    source fence is being selected.  A pending policy is never converted into
    a usable receipt: ``status`` is ``blocked`` and ``execution_admitted`` is
    false.  Once the trusted policy is complete this same path validates both
    endpoints against its fixed pins before writing a ready receipt.
    """

    if lane not in TRUSTED_POLICY.routes:
        raise PreparationError(f"lane is not admitted by current-source policy v2: {lane}")
    manifest_value, manifest_sha256, manifest_bytes = _load_canonical(Path(manifest), "qualification manifest")
    if manifest_value.get("schema") != PREPARATION_SCHEMA:
        raise PreparationError(f"manifest schema must be {PREPARATION_SCHEMA!r}")
    if manifest_value.get("policy_schema") != POLICY_SCHEMA:
        raise PreparationError("manifest is not bound to current-source policy v2")
    if manifest_value.get("lane") != lane:
        raise PreparationError("manifest lane does not match preparation lane")
    if manifest_value.get("trusted_inputs") != TRUSTED_POLICY.manifest_binding():
        raise PreparationError("manifest trusted source/lock/probe/fixture binding differs from policy")
    routes = manifest_value.get("routes")
    if not isinstance(routes, list):
        raise PreparationError("manifest routes must be an array")
    try:
        validate_routes(lane, routes)
    except PolicyError as error:
        raise PreparationError(str(error)) from error

    before_path = Path(before).absolute()
    after_path = Path(after).absolute()
    disjoint = [("baseline source", before_path), ("candidate source", after_path)]
    if before_archive is not None:
        disjoint.append(("baseline source archive", Path(before_archive)))
    if after_archive is not None:
        disjoint.append(("candidate source archive", Path(after_archive)))
    if output_root is not None:
        disjoint.append(("qualification output root", Path(output_root)))
    disjoint.append(("preparation receipt", Path(output)))
    _disjoint(disjoint)
    before_observed = observe_source(before_path, "baseline", before_archive)
    after_observed = observe_source(after_path, "candidate", after_archive)
    if before_observed["source_root"] == after_observed["source_root"]:
        raise PreparationError("baseline and candidate source roots must be distinct")

    policy_status = TRUSTED_POLICY.status
    ready = policy_status == "ready"
    if ready:
        # Import locally so the diagnostic/pending path never mistakes an
        # observation for an accepted identity.
        from native_qualification.current_native_policy_v2 import validate_source

        before_observed = validate_source(before_path, "baseline", before_archive)
        after_observed = validate_source(after_path, "candidate", after_archive)
        status = "ready"
        reason = None
    else:
        status = "blocked"
        reason = "exact candidate source/tree/lock/probe/fixture pins are not frozen"

    output_root_identity = _create_fresh_output_root(Path(output_root) if output_root is not None else None)

    preparer = Path(__file__).resolve()
    tooling = {
        "policy": _tooling_identity(preparer.with_name("current_native_policy_v2.py"), "policy module"),
        "preparer": _tooling_identity(preparer, "preparer"),
        "validator": _tooling_identity(preparer.with_name("validate_current_native_v2.py"), "validator"),
    }
    receipt: dict[str, Any] = {
        "schema": PREPARATION_SCHEMA,
        "policy_schema": POLICY_SCHEMA,
        "policy_sha256": TRUSTED_POLICY.digest(),
        "lane": lane,
        "status": status,
        "execution_admitted": ready,
        "blocked_reason": reason,
        "manifest": {"path": str(Path(manifest).absolute()), "sha256": manifest_sha256, "bytes": manifest_bytes},
        "output_root": str(Path(output_root).absolute().resolve()) if output_root is not None else None,
        "output_root_initial": output_root_identity,
        "execution_receipt_relative": EXECUTION_RECEIPT_RELATIVE if output_root_identity is not None else None,
        "preparer": tooling["preparer"],
        "tooling": tooling,
        "protocol": TRUSTED_POLICY.json()["protocol"],
        "source_transitions": list(TRUSTED_POLICY.json()["source_transitions"]),
        "candidate_hints": dict(TRUSTED_POLICY.json()["candidate_hints"]),
        "identities": {"before": before_observed, "after": after_observed},
        "routes": routes,
    }
    destination = _external_output(
        Path(output), tuple(path for label, path in disjoint if label != "preparation receipt")
    )
    destination.write_bytes(_canonical(receipt))
    return receipt


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--lane", choices=sorted(TRUSTED_POLICY.routes), required=True)
    parser.add_argument("--manifest", type=Path, required=True)
    parser.add_argument("--before", type=Path, required=True)
    parser.add_argument("--after", type=Path, required=True)
    parser.add_argument("--before-archive", type=Path)
    parser.add_argument("--after-archive", type=Path)
    parser.add_argument("--output-root", type=Path)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args(argv)
    try:
        receipt = prepare(
            lane=args.lane,
            manifest=args.manifest,
            before=args.before,
            after=args.after,
            output=args.output,
            before_archive=args.before_archive,
            after_archive=args.after_archive,
            output_root=args.output_root,
        )
    except (OSError, PolicyError, PreparationError) as error:
        print(f"current-native preparation refused: {error}", file=sys.stderr)
        return 1
    print(f"current-native preparation {receipt['status']}: {receipt['policy_sha256']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
