"""Read the bounded 0516 Callgrind profile as diagnostic attribution.

The profile lane is deliberately kept separate from the native timing lanes.
This module reuses the 0515 raw Callgrind parser through a registered dynamic
module import, reads the 0516 ``--show-percs=no`` annotation format, and
reports only the selected direct edges.  It does not authorize a performance
or speedup claim.
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import math
import re
import sys
from pathlib import Path
from typing import Any


HERE = Path(__file__).resolve().parent
ANALYZER = HERE.parent / "change-0515" / "analyze.py"
STAGES = ("before", "after")
PROFILE_NAME = "profile"

RUNNER = "litchi_perf_baseline::run_xlsx_update_commit_save"
OPERATION = "litchi_perf_baseline::xlsx_commit_save_operation"
COMMIT = "litchi_xlsx::workbook::edit::semantic::transaction::Edit::commit"
WRITER = "litchi_opc::pkgwriter::PackageWriter::write_to_stream"
ROLE_SPECS = (
    ("runner_to_operation", RUNNER, OPERATION),
    ("operation_to_commit", OPERATION, COMMIT),
    ("operation_to_writer", OPERATION, WRITER),
)
EXPECTED_CALLS = 3
SHA256_RE = re.compile(r"^[0-9a-f]{64}$")
ANNOTATION_TOTAL_RE = re.compile(
    r"^\s*([\d,]+)\s+PROGRAM TOTALS\s*$", re.MULTILINE
)
BRK_RE = re.compile(r"\bbrk\s+segment\s+overflow\b", re.IGNORECASE)
# 0516 annotations were generated with --show-percs=no.  Keep the optional
# percentage and the ``.`` cost accepted so this parser also reads an older
# callgrind_annotate rendering when one is retained alongside the new lane.
ANNOTATION_FUNCTION_RE = re.compile(
    r"^\s*(?:[\d,]+|\.)\s+(?:\([^)]*\)\s+)?\*\s+(?P<function>.+?)\s*$"
)
ANNOTATION_EDGE_RE = re.compile(
    r"^\s*(?P<cost>[\d,]+)\s+"
    r"(?:(?:\(\s*(?P<percent>[-+]?\d+(?:\.\d+)?)\s*%\)\s+))?"
    r">\s+(?P<function>.+?)\s+\((?P<count>[\d,]+)x\)"
    r"(?:\s+\[[^]]*\])?\s*$"
)


class ProfileError(ValueError):
    """A present profile artifact is missing, malformed, or inconsistent."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise ProfileError(message)


def _no_duplicate_pairs(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        require(key not in result, f"duplicate JSON object key {key!r}")
        result[key] = value
    return result


def _reject_constant(value: str) -> None:
    raise ProfileError(f"non-finite JSON number {value!r}")


def load_json(path: Path) -> Any:
    try:
        return json.loads(
            path.read_text(encoding="utf-8"),
            object_pairs_hook=_no_duplicate_pairs,
            parse_constant=_reject_constant,
        )
    except ProfileError:
        raise
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise ProfileError(f"cannot load {path}: {error}") from error


def sha(path: Path) -> str:
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
    except OSError as error:
        raise ProfileError(f"cannot read {path}: {error}") from error
    return digest.hexdigest()


def read_text(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8", errors="strict")
    except (OSError, UnicodeError) as error:
        raise ProfileError(f"cannot read {path}: {error}") from error


def check_hash(value: Any, context: str) -> str:
    require(isinstance(value, str) and SHA256_RE.fullmatch(value) is not None,
            f"{context} is not a lowercase SHA-256 digest")
    return value


def regular_nonempty(path: Path, context: str) -> None:
    require(path.is_file() and not path.is_symlink(),
            f"{context} is missing or not a regular file")
    try:
        require(path.stat().st_size > 0, f"{context} is empty")
    except OSError as error:
        raise ProfileError(f"cannot stat {context}: {error}") from error


def _safe_artifact_path(directory: Path, name: Any, context: str) -> Path:
    require(isinstance(name, str) and name, f"{context} artifact name is invalid")
    relative = Path(name)
    require(not relative.is_absolute() and ".." not in relative.parts,
            f"{context} artifact path is unsafe")
    target = directory / relative
    require(target.parent.resolve() == directory.resolve() or
            directory.resolve() in target.parent.resolve().parents,
            f"{context} artifact path escapes its stage")
    return target


def _verify_artifact_map(directory: Path, artifacts: Any, context: str) -> dict[str, str]:
    require(isinstance(artifacts, dict) and artifacts,
            f"{context}.artifacts is missing")
    result: dict[str, str] = {}
    for name, digest in artifacts.items():
        target = _safe_artifact_path(directory, name, context)
        regular_nonempty(target, f"{context} artifact {name}")
        expected = check_hash(digest, f"{context}.artifacts.{name}")
        actual = sha(target)
        require(actual == expected, f"{context} artifact {name} hash differs")
        result[name] = actual
    return result


def _stage_paths(root: Path, stage: str) -> dict[str, Path]:
    directory = root / stage
    return {
        "raw": directory / "profile.out",
        "inclusive": directory / "profile-inclusive.txt",
        "exclusive": directory / "profile-exclusive.txt",
        "log": directory / "profile.log",
        "receipt": directory / "profile-receipt.json",
        "annotations": directory / "profile-annotations.json",
        "report": directory / "profile-report.json",
        "catalog": directory / "profile-catalog.json",
    }


def _load_analyzer() -> Any:
    """Load 0515's parser with dataclass module registration before execution."""

    regular_nonempty(ANALYZER, "0515 analyze.py")
    spec = importlib.util.spec_from_file_location("litchi_change0515_analyze", ANALYZER)
    require(spec is not None and spec.loader is not None,
            "cannot construct the 0515 analyzer module spec")
    module = importlib.util.module_from_spec(spec)
    # dataclasses resolves the module by __module__ while decorating the
    # parser's RawEdge and AnnotationEdge classes.  Register first so the
    # dynamic import remains valid on current Python versions.
    sys.modules[spec.name] = module
    try:
        spec.loader.exec_module(module)
    except Exception as error:
        sys.modules.pop(spec.name, None)
        raise ProfileError(f"cannot load 0515 analyze.py: {error}") from error
    return module


def _parse_annotation_without_percent(module: Any, path: Path) -> list[Any]:
    lines = read_text(path).splitlines()
    edges: list[Any] = []
    parent: str | None = None
    for line in lines:
        function = ANNOTATION_FUNCTION_RE.match(line)
        if function:
            parent = function.group("function").strip()
            continue
        edge = ANNOTATION_EDGE_RE.match(line)
        if edge is None or parent is None:
            continue
        percent = edge.group("percent")
        edges.append(module.AnnotationEdge(
            parent=parent,
            child=edge.group("function").strip(),
            calls=int(edge.group("count").replace(",", "")),
            cost=int(edge.group("cost").replace(",", "")),
            percent=float(percent) if percent is not None else 0.0,
        ))
    require(edges, f"{path.name} has no direct-callee annotation edges")
    return edges


def parse_annotation(module: Any, path: Path) -> tuple[list[Any], str]:
    """Use the 0515 parser where possible, then support 0516 no-percent text."""

    try:
        return module.parse_annotation(path), "0515.parse_annotation"
    except Exception as original:
        try:
            parsed = _parse_annotation_without_percent(module, path)
        except Exception as fallback:
            raise ProfileError(
                f"{path.name} annotation parsing failed: {original}; compatibility parser: {fallback}"
            ) from fallback
        return parsed, "0516.no-percent-compatible-parser"


def _annotation_total(path: Path) -> int:
    matches = ANNOTATION_TOTAL_RE.findall(read_text(path))
    require(len(matches) == 1 and int(matches[0].replace(",", "")) > 0,
            f"{path.name} must contain one nonzero PROGRAM TOTALS value")
    return int(matches[0].replace(",", ""))


def _positive_edges(module: Any, edges: list[Any], parent: str,
                    child: str, label: str) -> list[Any]:
    selected = [
        edge for edge in edges
        if module.base_name(edge.parent) == parent
        and module.base_name(edge.child) == child
    ]
    positive = [edge for edge in selected if edge.calls > 0]
    require(positive, f"{label} is absent from selected profile edges")
    require(all(edge.calls >= 0 and edge.cost >= 0 for edge in selected),
            f"{label} has a negative edge value")
    return positive


def _selected(module: Any, edges: list[Any], parent: str, child: str,
              source: str, label: str) -> dict[str, Any]:
    selected = _positive_edges(module, edges, parent, child, label)
    all_selected = [
        edge for edge in edges
        if module.base_name(edge.parent) == parent
        and module.base_name(edge.child) == child
    ]
    calls = sum(edge.calls for edge in selected)
    cost = sum(edge.cost for edge in selected)
    require(calls == EXPECTED_CALLS,
            f"{label} {source} calls {calls} != {EXPECTED_CALLS}")
    return {
        "calls": calls,
        "ir": cost,
        "record_count": len(all_selected),
        "positive_record_count": len(selected),
        "zero_call_ir": sum(edge.cost for edge in all_selected if edge.calls == 0),
    }


def _role(module: Any, raw: list[Any], inclusive: list[Any], exclusive: list[Any],
          name: str, parent: str, child: str) -> dict[str, Any]:
    label = f"{name} ({parent} -> {child})"
    selected = {
        "raw": _selected(module, raw, parent, child, "raw", label),
        "inclusive": _selected(module, inclusive, parent, child, "inclusive", label),
        "exclusive": _selected(module, exclusive, parent, child, "exclusive", label),
    }
    require(selected["raw"]["calls"] == selected["inclusive"]["calls"] ==
            selected["exclusive"]["calls"] == EXPECTED_CALLS,
            f"{label} raw/inclusive/exclusive call counts differ")
    # Inclusive annotation is the independent presentation of the raw direct
    # edge.  Exclusive costs are retained for attribution but are not forced
    # to equal inclusive costs because that would discard their meaning.
    require(selected["raw"]["ir"] == selected["inclusive"]["ir"],
            f"{label} raw and inclusive Ir costs differ")
    return {
        "parent": parent,
        "child": child,
        "expected_calls": EXPECTED_CALLS,
        "total_ir": selected["inclusive"]["ir"],
        "selected_costs_ir": {
            source: values["ir"] for source, values in selected.items()
        },
        "selected_calls": {
            source: values["calls"] for source, values in selected.items()
        },
        "records": {
            source: {
                "record_count": values["record_count"],
                "positive_record_count": values["positive_record_count"],
                "zero_call_ir": values["zero_call_ir"],
            }
            for source, values in selected.items()
        },
    }


def _verify_receipt(paths: dict[str, Path], stage: str) -> dict[str, str]:
    receipt = load_json(paths["receipt"])
    require(isinstance(receipt, dict), "profile receipt must be an object")
    require(receipt.get("stage") == stage, "profile receipt stage differs")
    require(receipt.get("exit_code") == 0, "profile capture did not exit successfully")
    artifacts = _verify_artifact_map(paths["receipt"].parent, receipt.get("artifacts"),
                                     "profile receipt")
    required = {"profile-report.json", "profile-catalog.json", "profile.log", "profile.out"}
    require(required <= set(artifacts), "profile receipt omits a required artifact")
    return artifacts


def _verify_annotations(paths: dict[str, Path], raw_sha: str) -> dict[str, str]:
    manifest = load_json(paths["annotations"])
    require(isinstance(manifest, dict), "profile annotation manifest must be an object")
    require(manifest.get("exit_code") == 0,
            "profile annotation generation did not exit successfully")
    require(check_hash(manifest.get("raw_sha256"), "profile annotations raw_sha256") == raw_sha,
            "profile annotation manifest is bound to a different raw profile")
    artifacts = _verify_artifact_map(paths["annotations"].parent, manifest.get("artifacts"),
                                     "profile annotations")
    required = {"profile-inclusive.txt", "profile-exclusive.txt"}
    require(required <= set(artifacts), "profile annotations omit a required rendering")
    return artifacts


def profile_result(root: Path, stage: str, module: Any) -> dict[str, Any] | None:
    paths = _stage_paths(root, stage)
    # Any profile-side artifact makes the stage present.  In particular, an
    # orphan report/catalog must not be mistaken for an absent profile and
    # bypass the receipt and annotation hash checks below.
    known = tuple(paths)
    if not any(path.exists() or path.is_symlink() for name, path in paths.items()
               if name in known):
        return None
    for name in known:
        regular_nonempty(paths[name], f"{stage}/{paths[name].name}")

    raw_sha = sha(paths["raw"])
    receipt_artifacts = _verify_receipt(paths, stage)
    annotation_artifacts = _verify_annotations(paths, raw_sha)
    module_hash = sha(ANALYZER)
    try:
        summary_ir, raw_edges = module.parse_raw(paths["raw"])
    except Exception as error:
        raise ProfileError(f"{stage}/profile.out raw parsing failed: {error}") from error
    try:
        inclusive_edges, inclusive_parser = parse_annotation(module, paths["inclusive"])
        exclusive_edges, exclusive_parser = parse_annotation(module, paths["exclusive"])
    except ProfileError:
        raise
    require(_annotation_total(paths["inclusive"]) == summary_ir,
            f"{stage}/profile-inclusive.txt total Ir differs from raw summary")
    require(_annotation_total(paths["exclusive"]) == summary_ir,
            f"{stage}/profile-exclusive.txt total Ir differs from raw summary")

    roles = {
        name: _role(module, raw_edges, inclusive_edges, exclusive_edges,
                    name, parent, child)
        for name, parent, child in ROLE_SPECS
    }
    log = read_text(paths["log"])
    brk_matches = BRK_RE.findall(log)
    hashes = {
        "raw": raw_sha,
        "inclusive": sha(paths["inclusive"]),
        "exclusive": sha(paths["exclusive"]),
        "log": sha(paths["log"]),
        "receipt": sha(paths["receipt"]),
        "annotations": sha(paths["annotations"]),
        "report": sha(paths["report"]),
        "catalog": sha(paths["catalog"]),
        "parser": module_hash,
    }
    require(receipt_artifacts["profile.out"] == hashes["raw"],
            f"{stage} receipt raw profile hash differs")
    require(annotation_artifacts["profile-inclusive.txt"] == hashes["inclusive"],
            f"{stage} inclusive annotation hash differs")
    require(annotation_artifacts["profile-exclusive.txt"] == hashes["exclusive"],
            f"{stage} exclusive annotation hash differs")
    return {
        "stage": stage,
        "lane": PROFILE_NAME,
        "profile_total_ir": summary_ir,
        "roles": roles,
        "hashes": hashes,
        "annotation_parsers": {
            "inclusive": inclusive_parser,
            "exclusive": exclusive_parser,
        },
        "brk_warning": {
            "present": bool(brk_matches),
            "count": len(brk_matches),
            "log": f"{stage}/profile.log",
        },
        "scope": "Selected direct Callgrind edges; diagnostic instruction attribution only",
    }


def _percent(after: int, before: int) -> float | None:
    if before <= 0:
        return None
    value = (float(after) / float(before) - 1.0) * 100.0
    require(math.isfinite(value), "profile percent change is not finite")
    return value


def compare(before: dict[str, Any], after: dict[str, Any]) -> dict[str, Any]:
    result: dict[str, Any] = {
        "status": "compared",
        "profile_total_ir_after_minus_before_percent": _percent(
            after["profile_total_ir"], before["profile_total_ir"]
        ),
        "roles": {},
        "scope": "Descriptive after/before Callgrind diagnostic; no speedup claim",
    }
    for name, _, _ in ROLE_SPECS:
        left = before["roles"][name]
        right = after["roles"][name]
        result["roles"][name] = {
            "after_minus_before_percent": {
                source: _percent(
                    right["selected_costs_ir"][source],
                    left["selected_costs_ir"][source],
                )
                for source in ("raw", "inclusive", "exclusive")
            },
            "before_selected_costs_ir": left["selected_costs_ir"],
            "after_selected_costs_ir": right["selected_costs_ir"],
        }
    return result


def collect(root: Path = HERE) -> dict[str, Any]:
    root = root.resolve()
    module = _load_analyzer()
    profiles: dict[str, dict[str, Any]] = {}
    issues: list[dict[str, str]] = []
    missing: list[str] = []
    for stage in STAGES:
        try:
            result = profile_result(root, stage, module)
        except ProfileError as error:
            issues.append({"stage": stage, "message": str(error)})
            continue
        if result is None:
            missing.append(stage)
        else:
            profiles[stage] = result

    comparison: dict[str, Any] = {
        "status": "after_missing" if "after" not in profiles else (
            "before_missing" if "before" not in profiles else "compared"
        )
    }
    if "before" in profiles and "after" in profiles:
        comparison = compare(profiles["before"], profiles["after"])
    return {
        "schema": "litchi-0516-profile-metrics-v1",
        "performance_claim": "none",
        "scope": (
            "Callgrind raw/inclusive/exclusive selected direct-edge costs and calls; instruction diagnostic only"
        ),
        "profiles": profiles,
        "comparison": comparison,
        "missing_stages": missing,
        "issues": issues,
        "status": "error" if issues else ("partial" if missing else "complete"),
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Extract descriptive metrics from the 0516 Callgrind profile lane."
    )
    parser.add_argument("--root", type=Path, default=HERE,
                        help="0516 evidence directory (default: this script's directory)")
    args = parser.parse_args(argv)
    try:
        result = collect(args.root)
    except (OSError, ProfileError) as error:
        print(json.dumps({
            "schema": "litchi-0516-profile-metrics-v1",
            "performance_claim": "none",
            "status": "error",
            "issues": [{"message": str(error)}],
        }, indent=2, sort_keys=True))
        return 1
    print(json.dumps(result, indent=2, sort_keys=True))
    return 1 if result["issues"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
