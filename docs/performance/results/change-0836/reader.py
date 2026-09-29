#!/usr/bin/env python3
"""Independent evidence reader for the 0836 frame-pointer profile.

The 0836 packet compares two builds of the same committed source: the normal
release build (``baseline``) and the release build with frame pointers.  This
module is deliberately offline.  It reads retained source freezes, build and
command receipts, and native JSON reports; it never builds, runs a workload,
or interprets a timing as an optimization claim.

Report-level evidence is delegated to the pinned 0835 reader, which in turn
loads the sealed 0834 strict ZIP/OPC/PPTX implementation.  This wrapper owns
the 0836 packet binding and the stage-aware custody checks.  In particular,
the two stages have independent environment and source-oracle bindings, while
fresh-child PIDs remain unique across a whole manifest.
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import math
import re
from pathlib import Path
from typing import Any, Iterable


HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[3]
BASE = "5bb91a4de42a403cd1dfcb6e10058e3fe62d2228"
STAGES = ("baseline", "fp")
STATES = ("warm", "cold-verified")
CASES = (
    "opc_file_eager_open",
    "opc_file_source_open",
    "opc_file_eager_one_part_atomic_save",
    "opc_file_source_one_part_atomic_save",
    "pptx_file_eager_open_selected_slide_lifecycle",
    "pptx_file_source_open_selected_slide_lifecycle",
)
QUALITY_GATES = ("fmt", "check", "test", "clippy", "doc", "boundaries")
ALLOWED_SOURCES = (
    "tools/perf-baseline/src/filesystem.rs",
    "tools/perf-baseline/src/filesystem/aligned_zip.rs",
    "tools/perf-baseline/README.md",
)
PINNED_0835_SHA256 = "331422a310c4fbacc93b923b89b9370b60c669abc46a08b53fc93cad503ccf3d"
TOOL_MANIFEST = "tools/perf-baseline/Cargo.toml"
SHA256_RE = re.compile(r"[0-9a-f]{64}\Z")
U64_MAX = (1 << 64) - 1


class QualificationError(AssertionError):
    """An immutable packet or report admission invariant failed."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise QualificationError(message)


def read(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing regular JSON: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise QualificationError(f"{path}: invalid JSON: {error}") from error


def sha(path: Path) -> str:
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1 << 20), b""):
                digest.update(block)
    except OSError as error:
        raise QualificationError(f"cannot hash {path}: {error}") from error
    return digest.hexdigest()


def _finite(value: Any, path: str = "json") -> None:
    if isinstance(value, float):
        require(math.isfinite(value), f"{path}: non-finite number")
    elif isinstance(value, dict):
        for key, child in value.items():
            require(isinstance(key, str), f"{path}: non-string key")
            _finite(child, f"{path}.{key}")
    elif isinstance(value, list):
        for index, child in enumerate(value):
            _finite(child, f"{path}[{index}]")


def _same(actual: Any, expected: Any, path: str) -> None:
    require(actual == expected, f"{path}: expected {expected!r}, got {actual!r}")


def _uint(value: Any, path: str, *, positive: bool = False) -> int:
    minimum = 1 if positive else 0
    require(type(value) is int and minimum <= value <= U64_MAX,
            f"{path}: expected {'positive' if positive else 'non-negative'} integer")
    return value


def _hash(value: Any, path: str) -> str:
    require(isinstance(value, str) and SHA256_RE.fullmatch(value) is not None,
            f"{path}: expected lowercase SHA-256")
    return value


def _descriptor(value: Any, base: Path, label: str, *, allow_missing: bool = False) -> Path:
    require(isinstance(value, dict), f"{label}: descriptor is missing")
    path_value = value.get("path")
    require(isinstance(path_value, str) and path_value, f"{label}.path: missing")
    _uint(value.get("bytes"), f"{label}.bytes")
    _hash(value.get("sha256"), f"{label}.sha256")
    candidate = Path(path_value)
    if not candidate.is_absolute():
        candidate = base / candidate
    require(not candidate.is_symlink(), f"{label}: descriptor is a symlink")
    candidate = candidate.resolve()
    if not candidate.exists():
        require(allow_missing, f"{label}: file is missing: {candidate}")
        return candidate
    require(candidate.is_file() and not candidate.is_symlink(),
            f"{label}: descriptor is not a regular file")
    _same(candidate.stat().st_size, value["bytes"], f"{label}.bytes")
    _same(sha(candidate), value["sha256"], f"{label}.sha256")
    return candidate


def _contains_descriptor(value: Any, expected: dict[str, Any]) -> bool:
    if isinstance(value, dict):
        if all(value.get(field) == expected.get(field)
               for field in ("path", "bytes", "sha256")):
            return True
        return any(_contains_descriptor(child, expected) for child in value.values())
    if isinstance(value, list):
        return any(_contains_descriptor(child, expected) for child in value)
    return False


def _load_pinned_reader() -> Any:
    pinned_path = ROOT / "docs/performance/results/change-0835/reader.py"
    require(pinned_path.is_file() and not pinned_path.is_symlink(),
            f"missing pinned 0835 reader: {pinned_path}")
    _same(sha(pinned_path), PINNED_0835_SHA256,
          f"{pinned_path}: pinned reader hash differs")
    spec = importlib.util.spec_from_file_location("litchi_reader_0835_pinned", pinned_path)
    require(spec is not None and spec.loader is not None,
            f"cannot load pinned 0835 reader: {pinned_path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


_PINNED = _load_pinned_reader()
_LEGACY = _PINNED._LEGACY
_strict_stats = _PINNED._strict_stats


def reconstruct_acquisition_samples(statistics: dict[str, Any]) -> list[int]:
    """Return elapsed values in acquisition order after strict validation."""

    values = _strict_stats(statistics, "statistics", len(statistics.get("samples", [])))
    result = [0] * len(values)
    for sorted_index, acquisition_index in enumerate(statistics["sample_order"]):
        result[acquisition_index] = values[sorted_index]
    return result


def _configure_stage(stage: str) -> None:
    """Bind both layers of the pinned reader to this packet and stage."""

    require(stage in STAGES, f"unknown build stage: {stage}")
    _PINNED.HERE = HERE
    _PINNED.ROOT = ROOT
    _PINNED.BASE = BASE
    _PINNED.STAGE = stage
    _LEGACY.HERE = HERE
    _LEGACY.ROOT = ROOT
    _LEGACY.BASE = BASE
    _LEGACY.ACTIVE_STAGE = stage
    _LEGACY.ACTIVE_SUFFIX = stage
    _LEGACY._SOURCE_STAGE = stage
    _LEGACY._SOURCE_ORACLE_CACHE = None


def _origin(packet: Path) -> dict[str, Any]:
    path = packet / "origin.json"
    origin = read(path)
    _finite(origin, str(path))
    require(isinstance(origin, dict), f"{path}: origin is not an object")
    _same(origin.get("base"), BASE, f"{path}.base")
    for group_name in ("normative", "unrelated"):
        group = origin.get(group_name)
        require(isinstance(group, dict) and group, f"{path}.{group_name}: missing")
        for relative, digest in group.items():
            relative_path = Path(relative) if isinstance(relative, str) else Path("/")
            require(isinstance(relative, str) and relative and
                    not relative_path.is_absolute() and ".." not in relative_path.parts,
                    f"{path}.{group_name}: invalid path {relative!r}")
            _hash(digest, f"{path}.{group_name}.{relative}")
            candidate = ROOT / relative
            require(candidate.is_file() and not candidate.is_symlink(),
                    f"{path}.{group_name}.{relative}: current file is missing")
            _same(sha(candidate), digest, f"{path}.{group_name}.{relative}")
    return origin


def _source_freeze(packet: Path, stage: str, origin: dict[str, Any]) -> dict[str, Any]:
    path = packet / f"freeze-{stage}.json"
    freeze = read(path)
    _finite(freeze, str(path))
    require(isinstance(freeze, dict), f"{path}: freeze is not an object")
    _same(freeze.get("stage"), stage, f"{path}.stage")
    source = freeze.get("source")
    require(isinstance(source, dict) and source, f"{path}.source: missing")
    driver_path = _descriptor(freeze.get("driver"), packet, f"{path}.driver")
    origin_path = _descriptor(freeze.get("origin"), packet, f"{path}.origin")
    _same(driver_path, packet / "driver.py", f"{path}.driver.path")
    _same(origin_path, packet / "origin.json", f"{path}.origin.path")

    for relative, digest in source.items():
        relative_path = Path(relative) if isinstance(relative, str) else Path("/")
        require(isinstance(relative, str) and relative and
                not relative_path.is_absolute() and ".." not in relative_path.parts,
                f"{path}.source: invalid path {relative!r}")
        _hash(digest, f"{path}.source.{relative}")
        frozen = packet / "sources" / stage / relative
        current = ROOT / relative
        candidate = frozen if frozen.exists() else current
        require(candidate.is_file() and not candidate.is_symlink(),
                f"{path}.source.{relative}: retained/current file is missing")
        _same(sha(candidate), digest, f"{path}.source.{relative}")

    # The stage-specific source archive is part of the packet's custody.  The
    # remainder of the inventory is checked against the committed checkout,
    # exactly as the root driver freezes it.
    for relative in ALLOWED_SOURCES:
        frozen = packet / "sources" / stage / relative
        require(frozen.is_file() and not frozen.is_symlink(),
                f"{path}: required retained source is missing: {relative}")
        require(relative in source, f"{path}.source: missing allowed source {relative}")
        _same(sha(frozen), source[relative], f"{path}.sources.{stage}.{relative}")
    return {"freeze": freeze, "freeze_path": path, "source": source,
            "origin": origin}


def _receipt_file(path: Path, label: str, expected_exit: int = 0) -> dict[str, Any]:
    require(path.is_file() and not path.is_symlink(), f"{label}: receipt is missing")
    value = read(path)
    require(isinstance(value, dict), f"{label}: receipt is not an object")
    _same(value.get("exit_code"), expected_exit, f"{label}.exit_code")
    _same(value.get("error"), None, f"{label}.error")
    _descriptor(value.get("log"), path.parent, f"{label}.log")
    return value


def _receipt_freeze(receipt: dict[str, Any], binding: dict[str, Any], label: str) -> None:
    descriptor = receipt.get("freeze")
    freeze_path = binding["freeze_path"]
    require(isinstance(descriptor, dict), f"{label}.freeze: missing")
    _same(descriptor.get("path"), str(freeze_path), f"{label}.freeze.path")
    _same(descriptor.get("bytes"), freeze_path.stat().st_size,
          f"{label}.freeze.bytes")
    _same(descriptor.get("sha256"), sha(freeze_path), f"{label}.freeze.sha256")


def _prepare_binding(stage: str) -> dict[str, Any]:
    _configure_stage(stage)
    origin = _origin(HERE)
    frozen = _source_freeze(HERE, stage, origin)
    build_path = HERE / f"build-{stage}.json"
    build = read(build_path)
    require(isinstance(build, dict), f"{build_path}: build receipt is malformed")
    if "stage" in build:
        _same(build.get("stage"), stage, f"{build_path}.stage")
    binary = build.get("binary")
    binary_path = _descriptor(binary, HERE, f"{build_path}.binary", allow_missing=True)
    _same(binary_path.name, "litchi-perf-baseline", f"{build_path}.binary.name")
    _same(binary_path.parent.name, stage, f"{build_path}.binary.parent")
    if not binary_path.exists():
        cleanup_path = HERE / "cleanup.json"
        cleanup = read(cleanup_path)
        require(isinstance(cleanup, dict), f"{cleanup_path}: cleanup is malformed")
        _same(cleanup.get("status"), "pass", f"{cleanup_path}.status")
        require(_contains_descriptor(cleanup, binary),
                f"{cleanup_path}: retained {stage} binary descriptor is missing")

    receipt_descriptor = build.get("receipt")
    receipt_path = _descriptor(receipt_descriptor, HERE, f"{build_path}.receipt")
    _same(receipt_path.name, "receipt.json", f"{build_path}.receipt.filename")
    _same(receipt_path.parent.name, f"build-{stage}", f"{build_path}.receipt.parent")
    receipt = _receipt_file(receipt_path, f"{build_path}.receipt")
    _receipt_freeze(receipt, frozen, f"{build_path}.receipt")
    argv = receipt.get("argv")
    require(isinstance(argv, list), f"{build_path}.receipt.argv: missing")
    require(argv and argv[0] == "cargo", f"{build_path}.receipt.argv: expected cargo")
    for option in ("--offline", "--locked", "--release", "--manifest-path", "--bin"):
        require(option in argv, f"{build_path}.receipt.argv: missing {option}")
    manifest_index = argv.index("--manifest-path")
    require(manifest_index + 1 < len(argv), f"{build_path}.receipt.argv: manifest value missing")
    _same(argv[manifest_index + 1], TOOL_MANIFEST,
          f"{build_path}.receipt.argv.--manifest-path")
    bin_index = argv.index("--bin")
    require(bin_index + 1 < len(argv), f"{build_path}.receipt.argv: binary value missing")
    _same(argv[bin_index + 1], "litchi-perf-baseline",
          f"{build_path}.receipt.argv.--bin")
    return {**frozen, "build": build, "build_path": build_path,
            "binary": binary, "binary_path": binary_path,
            "build_receipt": receipt}


def _parse_states(value: Any, label: str) -> tuple[str, ...]:
    if isinstance(value, str):
        values = tuple(part for part in value.split(",") if part)
    elif isinstance(value, (list, tuple)):
        values = tuple(value)
    else:
        raise QualificationError(f"{label}: expected cache-state string or list")
    require(values and all(state in STATES for state in values),
            f"{label}: unknown cache state")
    require(len(set(values)) == len(values), f"{label}: duplicate cache state")
    return values


def _case_from_report(value: dict[str, Any], label: str) -> str:
    configuration = value.get("configuration")
    require(isinstance(configuration, dict), f"{label}.configuration: missing")
    cases = configuration.get("cases")
    require(isinstance(cases, list) and len(cases) == 1,
            f"{label}.configuration.cases: expected one profiled case")
    case = cases[0]
    require(case in CASES, f"{label}.configuration.cases: unknown case {case!r}")
    return case


def _report_identity(checked: dict[str, Any], binding: dict[str, Any], label: str) -> None:
    identity = checked.get("binary")
    require(isinstance(identity, dict), f"{label}: checked binary identity missing")
    _same(identity.get("binary_sha256"), binding["binary"].get("sha256"),
          f"{label}.binary_identity.binary_sha256")
    _same(identity.get("binary_bytes"), binding["binary"].get("bytes"),
          f"{label}.binary_identity.binary_bytes")
    _same(identity.get("path"), binding["binary"].get("path"),
          f"{label}.binary_identity.path")


def _bind_sample_order(report: dict[str, Any], states: tuple[str, ...],
                       samples: int, label: str) -> None:
    """Tie sorted elapsed vectors back to acquisition-indexed evidence."""

    evidence = report.get("filesystem_evidence")
    results = report.get("results")
    require(isinstance(evidence, list) and isinstance(results, list),
            f"{label}: sample-order evidence is missing")
    for state in states:
        state_evidence = [sample for record in evidence
                          if isinstance(record, dict)
                          for sample in record.get("samples", [])
                          if isinstance(sample, dict) and sample.get("cache_state") == state]
        require(len(state_evidence) == samples,
                f"{label}.{state}: evidence count differs from {samples}")
        by_index: dict[int, dict[str, Any]] = {}
        for sample in state_evidence:
            sample_index = _uint(sample.get("sample_index"),
                                 f"{label}.{state}.sample_index")
            require(sample_index not in by_index,
                    f"{label}.{state}: duplicate evidence sample index {sample_index}")
            by_index[sample_index] = sample
        _same(sorted(by_index), list(range(samples)),
              f"{label}.{state}: evidence sample-index permutation")
        result = next((item for item in results
                       if isinstance(item, dict) and item.get("cache_state") == state), None)
        require(isinstance(result, dict), f"{label}.{state}: timed result is missing")
        statistics = result.get("elapsed_ns")
        sorted_values = _PINNED._strict_stats(
            statistics, f"{label}.{state}.elapsed_ns", samples
        )
        order = statistics["sample_order"]
        for sorted_index, acquisition_index in enumerate(order):
            _same(by_index[acquisition_index].get("elapsed_ns"), sorted_values[sorted_index],
                  f"{label}.{state}: sorted statistic/evidence mismatch at {sorted_index}")


def validate_report(path: Path | dict[str, Any], stage: str, state: Any,
                    samples: int, warmup: int, *,
                    expected_environment: dict[str, Any] | None = None,
                    seen_pids: set[int] | None = None) -> dict[str, Any]:
    """Validate one retained report against the selected build stage.

    The public positional contract is intentionally
    ``validate_report(path, stage, state, samples, warmup)``.  ``state`` may be
    one cache state or the comma-separated state string used by the native
    driver.  Optional keyword arguments let a manifest enforce environment
    consistency and fresh-child PID uniqueness across reports.
    """

    require(stage in STAGES, f"unknown build stage: {stage}")
    states = _parse_states(state, "report state")
    require(type(samples) is int and samples > 0, "report sample count must be positive")
    require(type(warmup) is int and warmup >= 0, "report warmup must be non-negative")
    binding = _prepare_binding(stage)
    if isinstance(path, Path):
        value = read(path)
        label = str(path)
    else:
        value = path
        label = "<report>"
    _finite(value, label)
    require(isinstance(value, dict), f"{label}: report is not an object")
    case = _case_from_report(value, label)
    checked = _LEGACY.validate_report(
        value, case, states, samples, warmup, binding["binary"],
        expected_environment=expected_environment, seen_pids=seen_pids,
    )
    _report_identity(checked, binding, label)
    _bind_sample_order(checked["report"], states, samples, label)
    return {**checked, "stage": stage, "case": case, "binding": binding}


def _field(row: dict[str, Any], *names: str) -> Any:
    for name in names:
        if name in row:
            return row[name]
    nested = row.get("result")
    if isinstance(nested, dict):
        for name in names:
            if name in nested:
                return nested[name]
    return None


def _row_label(row: dict[str, Any], index: int) -> str:
    value = _field(row, "label", "name", "command")
    require(isinstance(value, str) and value and "/" not in value and ".." not in value,
            f"manifest.rows[{index}].label: missing or unsafe")
    return value


def _row_descriptor(row: dict[str, Any], names: tuple[str, ...], label: str) -> Any:
    value = _field(row, *names)
    require(isinstance(value, dict), f"{label}: descriptor is missing")
    return value


def _receipt_contract(receipt: dict[str, Any], binding: dict[str, Any], label: str,
                      case: str, states: tuple[str, ...], samples: int,
                      warmup: int, report_path: Path,
                      receipt_path: Path | None = None) -> None:
    _receipt_freeze(receipt, binding, label)
    argv = receipt.get("argv")
    require(isinstance(argv, list), f"{label}.argv: missing")
    binary_path = binding["binary"].get("path")
    require(isinstance(binary_path, str) and binary_path, f"{label}: binary path missing")
    _same(argv.count(binary_path), 1, f"{label}.argv: binary binding")
    require(argv.count("--case") == 1, f"{label}.argv: missing/duplicate --case")
    case_index = argv.index("--case")
    require(case_index + 1 < len(argv), f"{label}.argv: --case value missing")
    _same(argv[case_index + 1], case, f"{label}.argv.--case")
    expected_state = ",".join(states)
    for option, expected in (("--samples", str(samples)),
                             ("--warmup", str(warmup)),
                             ("--filesystem-cache", expected_state)):
        require(argv.count(option) == 1, f"{label}.argv: missing/duplicate {option}")
        index = argv.index(option)
        require(index + 1 < len(argv), f"{label}.argv: {option} value missing")
        _same(argv[index + 1], expected, f"{label}.argv.{option}")
    require(argv.count("--filesystem-root") == 1 and argv.count("--json") == 1,
            f"{label}.argv: filesystem/report binding is missing")
    json_index = argv.index("--json")
    require(json_index + 1 < len(argv), f"{label}.argv: --json value missing")
    _same(Path(argv[json_index + 1]).resolve(), report_path.resolve(),
          f"{label}.argv.--json")
    root_index = argv.index("--filesystem-root")
    require(root_index + 1 < len(argv), f"{label}.argv: filesystem root value missing")
    require(Path(argv[root_index + 1]).is_absolute(),
            f"{label}.argv.--filesystem-root: expected absolute path")
    if "cwd" in receipt:
        _same(Path(receipt["cwd"]).resolve(), ROOT, f"{label}.cwd")
    started_path = (receipt_path.parent / "started.json") if receipt_path is not None else None
    if started_path is not None and started_path.is_file() and not started_path.is_symlink():
        started = read(started_path)
        require(isinstance(started, dict), f"{label}.started: malformed")
        _same(started.get("argv"), argv, f"{label}.started.argv")
        _same(Path(started.get("cwd", "")).resolve(), ROOT, f"{label}.started.cwd")


def _report_descriptor(row: dict[str, Any], packet: Path, label: str) -> tuple[Path, str]:
    descriptor = _row_descriptor(row, ("report", "report_descriptor", "reportdesc", "report_desc"),
                                 f"{label}.report")
    path = _descriptor(descriptor, packet, f"{label}.report")
    _same(path.parent, packet, f"{label}.report.parent")
    _same(sha(path), descriptor.get("sha256"), f"{label}.report.sha256")
    return path, descriptor["sha256"]


def _receipt_descriptor(row: dict[str, Any], packet: Path, label: str) -> tuple[Path, dict[str, Any]]:
    descriptor = _row_descriptor(
        row, ("receipt", "receipt_descriptor", "receiptdesc", "receipt_desc"),
        f"{label}.receipt",
    )
    path = _descriptor(descriptor, packet, f"{label}.receipt")
    _same(path.name, "receipt.json", f"{label}.receipt.filename")
    _same(path.parent.parent, packet / "commands", f"{label}.receipt.parent")
    value = _receipt_file(path, f"{label}.receipt")
    return path, value


def validate_manifest(path: Path) -> dict[str, Any]:
    """Validate a stage-aware manifest with complete command custody."""

    path = path.resolve()
    packet = path.parent
    value = read(path)
    _finite(value, str(path))
    require(isinstance(value, dict), f"{path}: manifest is not an object")
    _same(value.get("status"), "commands_pass", f"{path}.status")
    rows = value.get("rows")
    require(isinstance(rows, list) and rows, f"{path}.rows: missing")
    if "report_count" in value:
        _same(value.get("report_count"), len(rows), f"{path}.report_count")

    # Validate origin and both stage freezes before consuming any report.  This
    # ensures a manifest cannot silently omit a build stage it claims to use.
    origin = _origin(packet)
    bindings: dict[str, dict[str, Any]] = {}
    for stage in STAGES:
        bindings[stage] = _prepare_binding(stage)
        _same(bindings[stage]["origin"], origin, f"{path}: {stage} origin differs")

    expected_environment: dict[str, dict[str, Any]] = {}
    seen_pids: set[int] = set()
    checked: list[dict[str, Any]] = []
    labels: set[str] = set()
    expected_sample_count = 0
    for index, row in enumerate(rows):
        label = f"{path}.rows[{index}]"
        require(isinstance(row, dict), f"{label}: malformed")
        row_name = _row_label(row, index)
        require(row_name not in labels, f"{label}.label: duplicate {row_name!r}")
        labels.add(row_name)
        stage = _field(row, "stage")
        require(stage in STAGES, f"{label}.stage: expected baseline or fp")
        states = _parse_states(_field(row, "state", "cache_state", "cache_states"),
                               f"{label}.state")
        samples = _uint(_field(row, "samples"), f"{label}.samples", positive=True)
        warmup = _uint(_field(row, "warmup"), f"{label}.warmup")
        report_path, report_digest = _report_descriptor(row, packet, label)
        receipt_path, receipt = _receipt_descriptor(row, packet, label)
        _same(receipt_path.parent.name, row_name, f"{label}.receipt.parent")
        report_value = read(report_path)
        _finite(report_value, str(report_path))
        case = _case_from_report(report_value, str(report_path))
        row_case = _field(row, "case")
        if row_case is not None:
            _same(row_case, case, f"{label}.case")
        _receipt_contract(receipt, bindings[stage], f"{label}.receipt", case,
                          states, samples, warmup, report_path, receipt_path)
        checked_report = validate_report(
            report_value, stage, states, samples, warmup,
            expected_environment=expected_environment.get(stage),
            seen_pids=seen_pids,
        )
        expected_environment.setdefault(stage, checked_report["environment"])
        expected_sample_count += samples * len(states)
        checked.append({"index": index, "label": row_name, "stage": stage,
                        "case": case, "state": ",".join(states),
                        "samples": samples, "warmup": warmup,
                        "report": str(report_path), "report_sha256": report_digest,
                        "receipt": str(receipt_path),
                        "outputs": checked_report["outputs"]})
    if "sample_count" in value:
        _same(value.get("sample_count"), expected_sample_count, f"{path}.sample_count")
    _same(len(seen_pids), expected_sample_count,
          f"{path}: fresh-child PID count differs from declared samples")
    return {"status": "pass", "manifest_valid": True, "rows": checked,
            "validated_sample_count": expected_sample_count,
            "stage_environments": expected_environment,
            "claim_authorized": False, "performance_claim": "none"}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("manifest", nargs="?", type=Path)
    parser.add_argument("--write", type=Path,
                        help="also write the validation object to this path")
    parser.add_argument("--check", action="store_true",
                        help="compare an existing validation object instead of writing it")
    args = parser.parse_args()
    path = (args.manifest or (HERE / "qualification.json")).resolve()
    try:
        result = validate_manifest(path)
    except Exception as error:
        result = {"status": "invalid", "manifest_valid": False,
                  "claim_authorized": False, "performance_claim": "none",
                  "error": str(error)}
        exit_code = 1
    else:
        exit_code = 0
    encoded = json.dumps(result, sort_keys=True, indent=2) + "\n"
    destination = args.write.resolve() if args.write else path.parent / "validation.json"
    if args.check:
        require(destination.is_file() and not destination.is_symlink(),
                f"validation output is missing: {destination}")
        _same(read(destination), result, f"{destination}: validation differs")
    elif args.write:
        require(destination.parent.is_dir(),
                f"validation output parent is missing: {destination.parent}")
        require(not destination.exists(), f"validation output already exists: {destination}")
        destination.write_text(encoded, encoding="utf-8")
    print(encoded, end="")
    return exit_code


if __name__ == "__main__":
    raise SystemExit(main())
