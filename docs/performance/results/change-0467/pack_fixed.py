#!/usr/bin/env python3
"""Qualify and package the fixed-checkout 0467 ABBA reports.

This is the additive packaging entry point for the fixed-path experiment.  It
first runs :mod:`fixed_qualify`, derives one claim-scoped primary report per
ABBA leg, and then calls the repository's standard ``perf_abba_package`` API.
The full four-row reports remain in qualification custody; the package's
one-row projections contain the XLSX dense-wide primary scope while all three
DOC rows remain retained guards.

The helper never edits the claim registry.  Existing qualification, summaries,
and derived projection inputs are authenticated and reused; the standard
packager still requires a fresh or empty package directory for publication.
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys
from typing import Any, Iterable, Mapping


ROOT = Path(__file__).resolve().parent
REPO_ROOT = next(
    (candidate for candidate in ROOT.parents if (candidate / "tools").is_dir()),
    ROOT,
)
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

import fixed_qualify as fixed  # noqa: E402
from tools import perf_abba_package  # noqa: E402


PACKAGE_CHANGE_ID = fixed.PACKAGE_CHANGE_ID
PACKAGE_MANIFEST_NAME = f"{PACKAGE_CHANGE_ID}-manifest.json"
PACKAGE_SUMMARY_NAME = "summary.json"
DEFAULT_PACKAGE_DIR = "fixed-package"
DEFAULT_PACKAGE_INPUT_DIR = "fixed-package-inputs"
DEFAULT_QUALIFICATION_NAME = "fixed-qualification.json"
DEFAULT_FULL_SUMMARY_NAME = "fixed-summary.json"
DEFAULT_SUMMARY_NAME = "fixed-primary-summary.json"

FIXED_REPORTS = {
    "a1": "A1-fixed/report.json",
    "b1": "B1-fixed/report.json",
    "b2": "B2-fixed/report.json",
    "a2": "A2-fixed/report.json",
}
FIXED_ARTIFACT_NAMES = {
    "a1": "a1-fixed.json",
    "b1": "b1-fixed.json",
    "b2": "b2-fixed.json",
    "a2": "a2-fixed.json",
}


class PackagingPreparationError(fixed.QualificationError):
    """Raised when fixed evidence cannot be safely packaged."""


def _display(path: Path, root: Path) -> str:
    """Use stable bundle-relative paths whenever the path is inside root."""

    path = path.resolve()
    try:
        return str(path.relative_to(root))
    except ValueError:
        return str(path)


def _write_json_exclusive(path: Path, value: Mapping[str, Any]) -> None:
    """Write deterministic helper output without replacing existing evidence."""

    path = path.resolve()
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        with path.open("x", encoding="utf-8") as stream:
            json.dump(value, stream, indent=2, sort_keys=True, allow_nan=False)
            stream.write("\n")
    except FileExistsError as error:
        raise PackagingPreparationError(f"output already exists: {path}") from error
    except OSError as error:
        raise PackagingPreparationError(f"cannot write {path}: {error}") from error


def _reuse_or_write_json(path: Path, value: Mapping[str, Any], label: str) -> None:
    """Authenticate an existing deterministic output or create it once."""

    path = path.resolve()
    if path.is_file():
        try:
            existing = fixed.normal.load_json(path)
        except (OSError, ValueError, fixed.QualificationError) as error:
            raise PackagingPreparationError(f"cannot read existing {label} {path}: {error}") from error
        if existing != dict(value):
            raise PackagingPreparationError(
                f"existing {label} differs from the fixed qualification: {path}"
            )
        return
    _write_json_exclusive(path, value)


def _fixed_report_paths(root: Path) -> dict[str, Path]:
    """Return and check the four raw reports used by the standard package."""

    reports = {role: root / relative for role, relative in FIXED_REPORTS.items()}
    missing = [str(path) for path in reports.values() if not path.is_file()]
    if missing:
        raise PackagingPreparationError(
            f"fixed normal reports are incomplete: {missing!r}"
        )
    return reports


def _primary_package_inputs(
    root: Path,
    reports: Mapping[str, Path],
    input_dir: Path,
) -> tuple[dict[str, Path], list[dict[str, Any]]]:
    """Create one-row report projections while retaining source custody data."""

    projected_reports: dict[str, Path] = {}
    custody: list[dict[str, Any]] = []
    projected_values: dict[str, dict[str, Any]] = {}
    for role, report_path in reports.items():
        source = fixed.normal.load_json(report_path)
        projected_values[role] = fixed._primary_report_projection(
            source, f"{role} source report"
        )

    # Build all projections in memory before creating the derived-input
    # directory.  If one report cannot be projected, no partial input set is
    # left behind.
    for role, report_path in reports.items():
        output_path = input_dir / f"{role}-primary.json"
        _reuse_or_write_json(output_path, projected_values[role], f"{role} projected report")
        projected_reports[role] = output_path
        custody.append(
            {
                "role": role,
                "source_report": _display(report_path, root),
                "source_report_sha256": fixed.normal.raw_sha256(report_path),
                "projected_report": _display(output_path, root),
                "projected_report_sha256": fixed.normal.raw_sha256(output_path),
                "retained_source_rows": 4,
                "projected_claim_rows": 1,
                "removed_guard_case": fixed.DOC_CASE,
                "removed_guard_shapes": list(fixed.DOC_SHAPES),
            }
        )
    return projected_reports, custody


def package_fixed(
    root: Path = ROOT,
    *,
    protocol_path: Path | None = None,
    package_dir: Path | None = None,
    package_input_dir: Path | None = None,
    qualification_out: Path | None = None,
    summary_out: Path | None = None,
    zstd_executable: str = "zstd",
) -> dict[str, Any]:
    """Qualify fixed evidence and create the standard deterministic package.

    ``fixed_qualify.qualify`` is run before any package destination is
    created.  Its all-row ``abba_summary`` is retained as the full guard
    summary.  A one-row primary projection is written to ``summary_out`` and
    supplied to ``perf_abba_package.package_artifacts``.  The standard
    packager independently recomputes that scoped summary from the four
    projected reports and refuses a mismatch.
    """

    root = root.resolve()
    reports = _fixed_report_paths(root)
    package_root = (package_dir or root / DEFAULT_PACKAGE_DIR).resolve()
    input_root = (
        package_input_dir or root / DEFAULT_PACKAGE_INPUT_DIR
    ).resolve()
    qualification_path = (
        qualification_out or root / DEFAULT_QUALIFICATION_NAME
    ).resolve()
    summary_path = (summary_out or root / DEFAULT_SUMMARY_NAME).resolve()

    qualification = fixed.qualify(root, protocol_path=protocol_path)
    _reuse_or_write_json(qualification_path, qualification, "fixed qualification")
    full_summary_path = root / DEFAULT_FULL_SUMMARY_NAME
    _reuse_or_write_json(
        full_summary_path,
        qualification["abba_summary"],
        "full fixed ABBA summary",
    )
    package_reports, custody = _primary_package_inputs(root, reports, input_root)
    try:
        projected_values = [
            fixed.normal.load_json(package_reports[role])
            for role in ("a1", "b1", "b2", "a2")
        ]
        strict_summary = fixed.perf_abba_summary.summarize_reports(
            projected_values,
            drift_ceilings=fixed.DRIFT_CEILINGS,
            cases=(fixed.XLSX_CASE,),
            shapes=(fixed.XLSX_SHAPE,),
        )
    except Exception as error:
        raise PackagingPreparationError(
            f"primary projected ABBA summary validation failed: {error}"
        ) from error
    _reuse_or_write_json(summary_path, strict_summary, "primary package summary")

    manifest = perf_abba_package.package_artifacts(
        change_id=PACKAGE_CHANGE_ID,
        output_dir=package_root,
        summary=summary_path,
        artifacts=package_reports,
        summary_name=PACKAGE_SUMMARY_NAME,
        manifest_name=PACKAGE_MANIFEST_NAME,
        artifact_names=FIXED_ARTIFACT_NAMES,
        zstd_executable=zstd_executable,
    )

    claim_registration = qualification.get("claim_registration")
    if not isinstance(claim_registration, dict):
        raise PackagingPreparationError(
            "fixed qualification omitted claim_registration proposal"
        )
    return {
        "schema": "litchi-0467-fixed-package-v1",
        "qualification": {
            "path": _display(qualification_path, root),
            "sha256": fixed.normal.raw_sha256(qualification_path),
        },
        "summary": {
            "source_path": _display(summary_path, root),
            "package_path": _display(package_root / PACKAGE_SUMMARY_NAME, root),
            "sha256": fixed.normal.raw_sha256(summary_path),
            "canonical_sha256": fixed.normal.canonical_sha256(strict_summary),
        },
        "package_inputs": {
            "directory": _display(input_root, root),
            "scope": {
                "case": fixed.XLSX_CASE,
                "shape": fixed.XLSX_SHAPE,
                "rows_per_role": 1,
            },
            "full_source_reports_retained_in_qualification": True,
            "custody": custody,
        },
        "package": {
            "directory": _display(package_root, root),
            "manifest_path": _display(
                package_root / PACKAGE_MANIFEST_NAME, root
            ),
            "manifest": manifest,
        },
        "scope": {
            "primary_case": fixed.XLSX_CASE,
            "primary_shape": fixed.XLSX_SHAPE,
            "guard_case": fixed.DOC_CASE,
            "guard_shapes": list(fixed.DOC_SHAPES),
            "artifact_roles": list(FIXED_ARTIFACT_NAMES),
        },
        "claim_registration": claim_registration,
    }


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=ROOT)
    parser.add_argument("--protocol", type=Path)
    parser.add_argument("--package-dir", type=Path)
    parser.add_argument("--package-input-dir", type=Path)
    parser.add_argument("--qualification-out", type=Path)
    parser.add_argument("--summary-out", type=Path)
    parser.add_argument("--zstd", dest="zstd_executable", default="zstd")
    parser.add_argument("--json-out", type=Path)
    return parser


def main(argv: Iterable[str] | None = None) -> int:
    args = build_parser().parse_args(list(argv) if argv is not None else None)
    try:
        root = args.root.resolve()
        result = package_fixed(
            root,
            protocol_path=args.protocol.resolve() if args.protocol else None,
            package_dir=args.package_dir.resolve() if args.package_dir else None,
            package_input_dir=(
                args.package_input_dir.resolve()
                if args.package_input_dir
                else None
            ),
            qualification_out=(
                args.qualification_out.resolve()
                if args.qualification_out
                else None
            ),
            summary_out=(args.summary_out.resolve() if args.summary_out else None),
            zstd_executable=args.zstd_executable,
        )
        if args.json_out is not None:
            _write_json_exclusive(args.json_out, result)
        json.dump(result, sys.stdout, indent=2, sort_keys=True, allow_nan=False)
        sys.stdout.write("\n")
        return 0
    except (
        PackagingPreparationError,
        fixed.QualificationError,
        perf_abba_package.ArtifactPackagingError,
        OSError,
        ValueError,
    ) as error:
        print(f"litchi-0467-fixed-package: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
