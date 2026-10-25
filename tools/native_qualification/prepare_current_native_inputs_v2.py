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
    if lexical.exists() or lexical.is_symlink():
        raise PreparationError(f"refusing to overwrite preparation receipt: {output}")
    resolved = lexical.parent.resolve() / lexical.name
    for root in roots:
        root = root.resolve()
        if resolved == root or root in resolved.parents:
            raise PreparationError(f"preparation receipt must be outside source roots: {output}")
    resolved.parent.mkdir(parents=True, exist_ok=True)
    return resolved


def prepare(
    *,
    lane: str,
    manifest: Path,
    before: Path,
    after: Path,
    output: Path,
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
    routes = manifest_value.get("routes")
    if not isinstance(routes, list):
        raise PreparationError("manifest routes must be an array")
    try:
        validate_routes(lane, routes)
    except PolicyError as error:
        raise PreparationError(str(error)) from error

    before_path = Path(before).absolute()
    after_path = Path(after).absolute()
    before_observed = observe_source(before_path, "baseline")
    after_observed = observe_source(after_path, "candidate")
    if before_observed["source_root"] == after_observed["source_root"]:
        raise PreparationError("baseline and candidate source roots must be distinct")

    policy_status = TRUSTED_POLICY.status
    ready = policy_status == "ready"
    if ready:
        # Import locally so the diagnostic/pending path never mistakes an
        # observation for an accepted identity.
        from native_qualification.current_native_policy_v2 import validate_source

        validate_source(before_path, "baseline")
        validate_source(after_path, "candidate")
        status = "ready"
        reason = None
    else:
        status = "blocked"
        reason = "exact candidate source/tree/lock/probe/fixture pins are not frozen"

    preparer = Path(__file__).resolve()
    preparer_bytes = preparer.read_bytes()
    receipt: dict[str, Any] = {
        "schema": PREPARATION_SCHEMA,
        "policy_schema": POLICY_SCHEMA,
        "policy_sha256": TRUSTED_POLICY.digest(),
        "lane": lane,
        "status": status,
        "execution_admitted": ready,
        "blocked_reason": reason,
        "manifest": {"path": str(Path(manifest).absolute()), "sha256": manifest_sha256, "bytes": manifest_bytes},
        "preparer": {"path": str(preparer), "sha256": hashlib.sha256(preparer_bytes).hexdigest(), "bytes": len(preparer_bytes)},
        "protocol": TRUSTED_POLICY.json()["protocol"],
        "source_transitions": list(TRUSTED_POLICY.json()["source_transitions"]),
        "candidate_hints": dict(TRUSTED_POLICY.json()["candidate_hints"]),
        "identities": {"before": before_observed, "after": after_observed},
        "routes": routes,
    }
    destination = _external_output(Path(output), (Path(before), Path(after)))
    destination.write_bytes(_canonical(receipt))
    return receipt


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--lane", choices=sorted(TRUSTED_POLICY.routes), required=True)
    parser.add_argument("--manifest", type=Path, required=True)
    parser.add_argument("--before", type=Path, required=True)
    parser.add_argument("--after", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args(argv)
    try:
        receipt = prepare(
            lane=args.lane,
            manifest=args.manifest,
            before=args.before,
            after=args.after,
            output=args.output,
        )
    except (OSError, PolicyError, PreparationError) as error:
        print(f"current-native preparation refused: {error}", file=sys.stderr)
        return 1
    print(f"current-native preparation {receipt['status']}: {receipt['policy_sha256']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
