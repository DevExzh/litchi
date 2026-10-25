#!/usr/bin/env python3
"""Validate a DrawingML raw v2 result and its source preparation receipt.

The probe report format is intentionally shared with the reviewed v1 probe:
changing the source policy must not make old reports appear to be new
evidence.  The *summary* and preparation receipts are v2, however, and this
adapter validates the v2 receipt before delegating the row/fidelity checks to
the unchanged v1 report validator.

This module is execution-ready when the checked-in v2 candidate binding is
filled in.  Until then the preparation module fails closed; no command-line
hash or receipt value can activate the protocol.
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys
from typing import Any

# The validator is launched by the generic runner as a file path.  Bootstrap
# the checked-in tools package before importing the v2 policy modules.
TOOLS_ROOT = Path(__file__).resolve().parents[1]
if str(TOOLS_ROOT) not in sys.path:
    sys.path.insert(0, str(TOOLS_ROOT))

from native_qualification import drawingml_raw_validator as legacy
from native_qualification import prepare_drawingml_raw_inputs_v2 as preparation


SCHEMA = legacy.SCHEMA
SUMMARY_SCHEMA = "litchi.drawingml.raw.validation.v2"
PROTOCOL_VERSION = 2
ABBA = legacy.ABBA
LANES = legacy.LANES
ValidationError = legacy.ValidationError


def _validate_preparation(arguments: argparse.Namespace) -> dict[str, Any]:
    receipt = getattr(arguments, "preparation_receipt", None)
    baseline_root = getattr(arguments, "baseline_root", None)
    candidate_root = getattr(arguments, "candidate_root", None)
    trusted_root = getattr(arguments, "trusted_root", None)
    if not all(isinstance(value, Path) for value in (receipt, baseline_root, candidate_root, trusted_root)):
        raise ValidationError("v2 validator requires its bound preparation receipt and source roots")
    try:
        return preparation.validate_preparation_receipt(
            receipt,
            baseline_root=baseline_root,
            candidate_root=candidate_root,
            baseline_revision=str(arguments.baseline_revision),
            candidate_revision=str(arguments.candidate_revision),
            trusted_root=trusted_root,
        )
    except (OSError, ValueError) as error:
        raise ValidationError(f"v2 preparation receipt validation failed: {error}") from error


def validate(arguments: argparse.Namespace) -> dict[str, Any]:
    """Validate v2 source identity first, then the established probe rows."""

    receipt = _validate_preparation(arguments)
    summary = dict(legacy.validate(arguments))
    # The child reports deliberately retain the v1 probe schema.  Only this
    # source-policy summary is v2, making an old v1 receipt impossible to
    # relabel by simply copying it into a new output directory.
    summary["schema"] = SUMMARY_SCHEMA
    summary["protocol_version"] = PROTOCOL_VERSION
    summary["preparation_receipt"] = {
        "path": str(Path(arguments.preparation_receipt).resolve()),
        "schema": preparation.SCHEMA,
        "baseline_revision": receipt["baseline_revision"],
        "candidate_revision": receipt["candidate_revision"],
    }
    return summary


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output-dir", type=Path, required=True)
    parser.add_argument("--lane", required=True)
    parser.add_argument("--endpoint", choices=("raw", "owner"), default="raw")
    parser.add_argument("--fixture", type=Path, required=True)
    parser.add_argument("--baseline-revision", required=True)
    parser.add_argument("--candidate-revision", required=True)
    parser.add_argument("--source-fence", required=True)
    parser.add_argument("--preparation-receipt", type=Path, required=True)
    parser.add_argument("--baseline-root", type=Path, required=True)
    parser.add_argument("--candidate-root", type=Path, required=True)
    parser.add_argument("--trusted-root", type=Path, required=True)
    arguments = parser.parse_args(argv)
    try:
        summary = validate(arguments)
        target = arguments.output_dir / "summary.json"
        if target.exists() or target.is_symlink():
            raise ValidationError(f"refusing to overwrite {target}")
        target.write_text(json.dumps(summary, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    except (OSError, ValidationError) as error:
        print(f"drawingml raw v2 validation refused: {error}", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
