#!/usr/bin/env python3
"""Run the frozen 0531 eager analysis with its recorded binary-path aliases.

The capture guard is immutable after the eager children were recorded.  The
retained ``binary-normal.json`` files use the frozen ``/tmp`` custody alias,
whereas the child embeds the canonical retained-binary path in its report.
This wrapper changes only the expected path passed to the frozen raw-report
validator.  Every other custody, schema, identity, timing, RSS, ABBA, drift,
and review check remains in ``eager_guard.py``.

The mapping is deliberately literal.  It does not call ``Path.resolve`` or
depend on the temporary alias being present, so the wrapper can be replayed
after cleanup while preserving the original path representation in the
captured receipts and metadata.
"""

from __future__ import annotations

import argparse
import importlib.util
import json
from pathlib import Path
import sys
from typing import Any


HERE = Path(__file__).parent.absolute()
FROZEN_PATH = HERE / "eager_guard.py"
CLEANUP_PATH = HERE / "cleanup.json"


class EvidenceError(ValueError):
    """A missing frozen guard or an unexpected retained binary alias."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise EvidenceError(message)


def load_frozen_guard() -> Any:
    require(FROZEN_PATH.is_file(), f"missing frozen eager guard: {FROZEN_PATH}")
    spec = importlib.util.spec_from_file_location("xlsx_0531_frozen_eager_guard", FROZEN_PATH)
    require(spec is not None and spec.loader is not None,
            f"cannot load frozen eager guard: {FROZEN_PATH}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


GUARD = load_frozen_guard()
ORIGINAL_VALIDATE_EAGER_REPORT = GUARD.validate_eager_report
ORIGINAL_BINARY_IDENTITY = GUARD.binary_identity

# These are the paths recorded by the retained binary metadata and by the
# child reports respectively.  The values are frozen campaign custody facts,
# not symlink resolution results.
FROZEN_ALIAS_TO_CANONICAL = {
    "/tmp/litchi-goal-0531/baseline-normal":
        "/home/zhuhe/litchi-goal-0531-target/retained-binaries/baseline-normal",
    "/tmp/litchi-goal-0531/candidate-normal":
        "/home/zhuhe/litchi-goal-0531-target/retained-binaries/candidate-normal",
}


def validate_eager_report(raw: Any, primary: dict[str, Any],
                          eager: dict[str, Any], job: dict[str, Any],
                          binary: dict[str, Any]) -> dict[str, Any]:
    """Delegate unchanged validation after replacing only the expected path."""

    alias = binary.get("path")
    canonical = FROZEN_ALIAS_TO_CANONICAL.get(alias)
    require(canonical is not None,
            f"no frozen canonical mapping for retained binary alias {alias!r}")
    expected = dict(binary)
    expected["path"] = canonical
    # The frozen validator still checks the raw report's path, binary digest,
    # byte count, and every other eager identity against this one mapped value.
    return ORIGINAL_VALIDATE_EAGER_REPORT(raw, primary, eager, job, expected)


# The retained binary aliases are intentionally removed by campaign cleanup.
# This replay path reconstructs only the frozen metadata identity when the
# alias is absent; it never claims that a deleted executable was re-hashed.
def replay_binary_identity(stage: str, eager: dict[str, Any],
                           identity: dict[str, Any], alias: str) -> dict[str, Any]:
    primary = GUARD.read_json(GUARD.PLAN_PATH)
    require(isinstance(primary, dict), "primary plan is not an object")
    owned = primary.get("owned_paths")
    require(isinstance(owned, list) and owned, "primary plan owned paths are missing")
    cleanup = GUARD.read_json(CLEANUP_PATH)
    require(isinstance(cleanup, dict), "cleanup receipt is not an object")
    require(cleanup.get("plan_sha256") == GUARD.sha256(GUARD.PLAN_PATH),
            "cleanup receipt plan binding differs")
    require(cleanup.get("removed") == owned,
            "cleanup receipt removed paths differ from the primary plan")
    require(cleanup.get("owned_paths_absent") is True,
            "cleanup receipt does not certify owned paths absent")
    require(cleanup.get("accessible_process_references") == [],
            "cleanup receipt retains accessible owned-path references")
    require(cleanup.get("python_cache_absent") is True,
            "cleanup receipt does not certify Python caches absent")
    for raw_path in owned:
        require(isinstance(raw_path, str) and raw_path,
                "cleanup receipt contains a malformed owned path")
        path = Path(raw_path)
        require(not path.exists() and not path.is_symlink(),
                f"owned cleanup path remains present: {raw_path}")

    expected_alias = str(Path(owned[0]) / f"{stage}-normal")
    require(alias == expected_alias,
            f"{stage} binary metadata alias differs from frozen custody path")
    require(alias in FROZEN_ALIAS_TO_CANONICAL,
            f"no frozen canonical mapping for retained binary alias {alias!r}")
    expected_sha = eager["binary_sha256"].get(stage)
    require(identity.get("sha256") == expected_sha,
            f"{stage} deleted binary metadata digest differs from eager plan")
    GUARD.digest(identity.get("sha256"), f"{stage} deleted binary SHA")
    GUARD.positive_integer(identity.get("bytes"),
                           f"{stage} deleted binary byte count")

    stage_dir = HERE / stage
    source_manifest = stage_dir / "source-manifest.json"
    source_manifest_sha = GUARD.sha256(source_manifest)
    require(identity.get("source_manifest_sha256") == source_manifest_sha,
            f"{stage} deleted binary source manifest binding differs")
    build_receipt = stage_dir / "build-normal.receipt.json"
    require(identity.get("build_receipt_sha256") == GUARD.sha256(build_receipt),
            f"{stage} deleted binary build receipt binding differs")
    receipt = GUARD.read_json(build_receipt)
    require(isinstance(receipt, dict) and receipt.get("exit_code") == 0,
            f"{stage} normal build receipt is not successful")
    require(receipt.get("plan_sha256") == GUARD.sha256(GUARD.PLAN_PATH),
            f"{stage} normal build receipt plan binding differs")
    require(receipt.get("script_sha256") == GUARD.sha256(GUARD.RUN_PATH),
            f"{stage} normal build receipt script binding differs")
    require(receipt.get("source_manifest_sha256") == source_manifest_sha
            and receipt.get("working_source_manifest_sha256") == source_manifest_sha,
            f"{stage} normal build receipt source binding differs")
    return {
        "path": alias,
        "sha256": identity["sha256"],
        "bytes": identity["bytes"],
    }


def binary_identity(stage: str, eager: dict[str, Any]) -> tuple[dict[str, Any], Path]:
    """Use the frozen live check, or the strict post-cleanup replay record."""

    metadata = GUARD.read_json(HERE / stage / "binary-normal.json")
    require(isinstance(metadata, dict), f"{stage} binary identity is not an object")
    alias = metadata.get("path")
    require(isinstance(alias, str) and alias,
            f"{stage} binary metadata path is missing")
    alias_path = Path(alias)
    if alias_path.exists():
        # Preserve the frozen binary hash/size/executable checks whenever the
        # retained file is still live.
        return ORIGINAL_BINARY_IDENTITY(stage, eager)
    identity = replay_binary_identity(stage, eager, metadata, alias)
    # The returned path is metadata custody, not a live executable path.  The
    # frozen receipt validator consumes only the recorded identity below.
    return identity, alias_path


# ``validate_stage`` resolves these names in the frozen module's globals.
# Rebinding only the two representation-compatibility hooks leaves all raw
# report, receipt, timing, RSS, ABBA, drift, and review checks frozen.
GUARD.validate_eager_report = validate_eager_report
GUARD.binary_identity = binary_identity


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path,
                        default=HERE / "eager-comparison.json",
                        help="comparison output path")
    args = parser.parse_args()
    try:
        output = args.output
        output.parent.mkdir(parents=True, exist_ok=True)
        value = GUARD.analyze(output)
        GUARD.write_json(output, value)
    except (EvidenceError, GUARD.EvidenceError, OSError,
            json.JSONDecodeError, AssertionError) as error:
        print(f"eager_analysis.py: error: {error}", file=sys.stderr)
        return 2
    print(f"0531 eager analysis verified: {output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
