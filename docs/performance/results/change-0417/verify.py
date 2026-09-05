#!/usr/bin/env python3
"""Verify the complete 0417 CRUD baseline evidence bundle.

This is deliberately a bundle wrapper, rather than another report parser.  The
report and corpus semantics remain owned by ``tools.summarize_crud_baseline``;
this file checks the stronger evidence contract around that parser: the
declared matrix, provenance, command line, preflight, and serialized capture
order.  Historical absolute paths in captured argv values are accepted when
their suffix is the manifest-relative path, so an exported bundle remains
verifiable from a different checkout.
"""

from __future__ import annotations

import argparse
import copy
import datetime as dt
import hashlib
import json
from pathlib import Path, PureWindowsPath
import re
import sys
from typing import Any


class VerificationError(ValueError):
    """A captured evidence contract is invalid."""


SHA256_RE = re.compile(r"^[0-9a-fA-F]{64}$")
EXPECTED_CHANGE = 417
EXPECTED_NORMAL_SAMPLES = 500
EXPECTED_NORMAL_WARMUPS = 20
EXPECTED_ALLOCATOR_SAMPLES = 30
EXPECTED_ALLOCATOR_WARMUPS = 3
EXPECTED_CPU = 2
EXPECTED_WORKERS = 1
EXPECTED_SELECTOR_COUNT = 30


def _repo_root() -> Path:
    """Find the checkout containing this published verifier and ``tools``."""

    here = Path(__file__).resolve()
    parents = list(here.parents)
    candidates = ([parents[4]] if len(parents) > 4 else []) + parents
    for candidate in candidates:
        if (candidate / "tools" / "summarize_crud_baseline.py").is_file():
            return candidate
    candidate = Path.cwd().resolve()
    if (candidate / "tools" / "summarize_crud_baseline.py").is_file():
        return candidate
    raise VerificationError("cannot locate tools/summarize_crud_baseline.py")


REPO_ROOT = _repo_root()
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from tools import summarize_crud_baseline  # noqa: E402  (path set above)


def _fail(path: str, message: str) -> None:
    raise VerificationError(f"{path}: {message}")


def _load_json(path: Path, label: str | None = None) -> Any:
    name = label or path.as_posix()

    def no_duplicate_pairs(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
        result: dict[str, Any] = {}
        for key, value in pairs:
            if key in result:
                _fail(name, f"duplicate JSON key {key!r} is not allowed")
            result[key] = value
        return result

    try:
        with path.open("r", encoding="utf-8") as stream:
            return json.load(
                stream,
                object_pairs_hook=no_duplicate_pairs,
                parse_constant=lambda value: _reject_constant(value, name),
            )
    except FileNotFoundError:
        _fail(name, "file is missing")
    except json.JSONDecodeError as error:
        _fail(name, f"invalid JSON: {error}")
    raise AssertionError("unreachable")


def _reject_constant(value: str, path: str) -> None:
    _fail(path, f"non-finite JSON constant {value!r} is not allowed")


def _object(value: Any, path: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        _fail(path, "must be an object")
    return value


def _list(value: Any, path: str) -> list[Any]:
    if not isinstance(value, list):
        _fail(path, "must be a list")
    return value


def _string(value: Any, path: str) -> str:
    if not isinstance(value, str) or not value:
        _fail(path, "must be a non-empty string")
    return value


def _integer(value: Any, path: str, *, minimum: int | None = None) -> int:
    if isinstance(value, bool) or not isinstance(value, int):
        _fail(path, "must be an integer")
    if minimum is not None and value < minimum:
        _fail(path, f"must be at least {minimum}")
    return value


def _sha256(value: Any, path: str) -> str:
    value = _string(value, path)
    if not SHA256_RE.fullmatch(value):
        _fail(path, "must be a 64-character hexadecimal SHA-256")
    return value.lower()


def _canonical(value: Any) -> str:
    return json.dumps(value, sort_keys=True, separators=(",", ":"), allow_nan=False)


def _same(left: Any, right: Any, path: str) -> None:
    if _canonical(left) != _canonical(right):
        _fail(path, "does not match the bound value")


def _sha256_file(path: Path, label: str) -> str:
    if not path.is_file():
        _fail(label, "file is missing")
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _required_file(root: Path, relative: str) -> Path:
    path = root / relative
    if not path.is_file():
        _fail(relative, "file is missing")
    return path


def _relative_file(root: Path, value: Any, path: str) -> tuple[str, Path]:
    value = _string(value, path)
    candidate = Path(value)
    if candidate.is_absolute() or PureWindowsPath(value).is_absolute():
        _fail(path, "manifest file paths must be relative")
    parts = PureWindowsPath(value).parts
    if not parts or any(part in ("", ".", "..") for part in parts):
        _fail(path, "must be a normalized relative path")
    relative = PureWindowsPath(*parts).as_posix()
    resolved = (root / Path(*parts)).resolve()
    root_resolved = root.resolve()
    try:
        resolved.relative_to(root_resolved)
    except ValueError:
        _fail(path, "escapes the capture root")
    if not resolved.is_file():
        _fail(path, "file is missing")
    return relative, resolved


def _portable_path(value: Any, relative: str, path: str) -> None:
    """Accept a relative path or an old absolute path ending in ``relative``."""

    value = _string(value, path)
    expected_parts = tuple(PureWindowsPath(relative).parts)
    candidate = PureWindowsPath(value)
    if Path(value).is_absolute() or candidate.is_absolute():
        actual_parts = tuple(candidate.parts)
        if len(actual_parts) < len(expected_parts) or actual_parts[-len(expected_parts) :] != expected_parts:
            _fail(path, f"absolute path does not end in {relative!r}")
        return
    if tuple(candidate.parts) != expected_parts:
        _fail(path, f"relative path must be {relative!r}")


def _argv_value(argv: list[str], flag: str, path: str) -> str:
    positions = [index for index, value in enumerate(argv) if value == flag]
    if len(positions) != 1 or positions[0] + 1 >= len(argv):
        _fail(path, f"must contain exactly one {flag} value")
    value = argv[positions[0] + 1]
    if not isinstance(value, str):
        _fail(path, f"{flag} value must be a string")
    return value


def _binary_name(identity: dict[str, Any], path: str) -> str:
    binary_path = _string(identity.get("path"), f"{path}.path")
    if not Path(binary_path).is_absolute() and not PureWindowsPath(binary_path).is_absolute():
        _fail(f"{path}.path", "must be absolute")
    name = Path(binary_path).name or PureWindowsPath(binary_path).name
    return _string(name, f"{path}.path basename")


def _parse_time(value: Any, path: str) -> dt.datetime:
    value = _string(value, path)
    try:
        parsed = dt.datetime.fromisoformat(value)
    except ValueError as error:
        _fail(path, f"invalid ISO-8601 timestamp: {error}")
    if parsed.tzinfo is None or parsed.utcoffset() is None:
        _fail(path, "timestamp must include a UTC offset")
    return parsed


def _validate_protocol(root: Path, matrix: dict[str, Any]) -> None:
    protocol = _object(_load_json(_required_file(root, "protocol.json"), "protocol.json"), "protocol.json")
    if protocol.get("change") != EXPECTED_CHANGE:
        _fail("protocol.change", f"must be {EXPECTED_CHANGE}")
    if protocol.get("matrix") != "matrix.json":
        _fail("protocol.matrix", "must bind matrix.json")
    normal = _object(protocol.get("normal"), "protocol.normal")
    allocator = _object(protocol.get("allocator"), "protocol.allocator")
    for section, values in (
        (
            "protocol.normal",
            {"repeats": 2, "samples": EXPECTED_NORMAL_SAMPLES, "warmups": EXPECTED_NORMAL_WARMUPS},
        ),
        (
            "protocol.allocator",
            {"repeats": 2, "samples": EXPECTED_ALLOCATOR_SAMPLES, "warmups": EXPECTED_ALLOCATOR_WARMUPS},
        ),
    ):
        current = normal if section.endswith("normal") else allocator
        for key, expected in values.items():
            if current.get(key) != expected:
                _fail(f"{section}.{key}", f"must be {expected}")
    if protocol.get("cpu") != EXPECTED_CPU:
        _fail("protocol.cpu", f"must be {EXPECTED_CPU}")
    if protocol.get("workers") != EXPECTED_WORKERS:
        _fail("protocol.workers", f"must be {EXPECTED_WORKERS}")
    if "no speedup" not in str(protocol.get("classification", "")).lower():
        _fail("protocol.classification", "must retain the no-speedup qualification")
    common_flags = _list(matrix.get("common_flags"), "matrix.common_flags")
    if not all(isinstance(value, str) for value in common_flags):
        _fail("matrix.common_flags", "must contain only strings")
    if _flag_value(common_flags, "--workers", "matrix.common_flags") != str(EXPECTED_WORKERS):
        _fail("matrix.common_flags", "must select one worker")
    if _flag_value(common_flags, "--filesystem-cache", "matrix.common_flags") != "warm":
        _fail("matrix.common_flags", "must select warm cache")


def _flag_value(tokens: list[str], flag: str, path: str) -> str:
    positions = [index for index, value in enumerate(tokens) if value == flag]
    if len(positions) != 1 or positions[0] + 1 >= len(tokens):
        _fail(path, f"must contain exactly one {flag} value")
    value = tokens[positions[0] + 1]
    if not isinstance(value, str):
        _fail(path, f"{flag} value must be a string")
    return value


def _validate_matrix(root: Path) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    matrix = _object(_load_json(_required_file(root, "matrix.json"), "matrix.json"), "matrix.json")
    jobs = _list(matrix.get("jobs"), "matrix.jobs")
    if len(jobs) != EXPECTED_SELECTOR_COUNT:
        _fail("matrix.jobs", f"must contain exactly {EXPECTED_SELECTOR_COUNT} selectors")
    normalized_jobs: list[dict[str, Any]] = []
    selectors: set[str] = set()
    for index, value in enumerate(jobs):
        job = _object(value, f"matrix.jobs[{index}]")
        selector = _string(job.get("selector"), f"matrix.jobs[{index}].selector")
        category = _string(job.get("category"), f"matrix.jobs[{index}].category")
        status = _string(job.get("index_status"), f"matrix.jobs[{index}].index_status")
        if selector in selectors:
            _fail(f"matrix.jobs[{index}].selector", "selector is duplicated")
        selectors.add(selector)
        if status not in {"measured", "correctness-only"}:
            _fail(f"matrix.jobs[{index}].index_status", "must be measured or correctness-only")
        normalized_jobs.append({"selector": selector, "category": category, "index_status": status})

    index_path = _required_file(root, "inputs/original-crud-index.json")
    taxonomy_path = _required_file(root, "inputs/CRUD_Scenario_Checklist.md")
    index_sha = _sha256_file(index_path, "inputs/original-crud-index.json")
    taxonomy_sha = _sha256_file(taxonomy_path, "inputs/CRUD_Scenario_Checklist.md")
    if index_sha != _sha256(matrix.get("index_sha256"), "matrix.index_sha256"):
        _fail("matrix.index_sha256", "does not match inputs/original-crud-index.json")
    if taxonomy_sha != _sha256(matrix.get("taxonomy_sha256"), "matrix.taxonomy_sha256"):
        _fail("matrix.taxonomy_sha256", "does not match inputs/CRUD_Scenario_Checklist.md")

    index = _object(_load_json(index_path, "inputs/original-crud-index.json"), "original-crud-index.json")
    categories = _list(index.get("categories"), "original-crud-index.json.categories")
    index_jobs: list[dict[str, Any]] = []
    for category_index, category_value in enumerate(categories):
        category = _object(category_value, f"original-crud-index.json.categories[{category_index}]")
        category_id = _string(category.get("id"), f"original-crud-index.json.categories[{category_index}].id")
        scenarios = _list(
            category.get("scenarios"),
            f"original-crud-index.json.categories[{category_index}].scenarios",
        )
        for scenario_index, scenario_value in enumerate(scenarios):
            scenario = _object(
                scenario_value,
                f"original-crud-index.json.categories[{category_index}].scenarios[{scenario_index}]",
            )
            # The index also retains explicit unsupported checklist rows.  A
            # row without a selector is outside this executable matrix and is
            # intentionally excluded from the 30 selector comparison.
            if "selector" not in scenario:
                continue
            selector = _string(scenario.get("selector"), "original CRUD selector")
            status = _string(scenario.get("status"), "original CRUD selector status")
            index_jobs.append({"selector": selector, "category": category_id, "index_status": status})
    if len(index_jobs) != EXPECTED_SELECTOR_COUNT:
        _fail("original-crud-index.json", "does not contain the declared 30 selector scenarios")
    if index_jobs != normalized_jobs:
        _fail("matrix.jobs", "ordered selectors/categories/statuses do not match the bound original index")
    _validate_protocol(root, matrix)
    return matrix, normalized_jobs


def _validate_build_identity(root: Path) -> tuple[dict[str, Any], dict[str, Any]]:
    build = _object(
        _load_json(_required_file(root, "build-identity.json"), "build-identity.json"),
        "build-identity.json",
    )
    revision = _string(build.get("revision"), "build-identity.revision")
    if build.get("source_status") != "":
        _fail("build-identity.source_status", "must be an empty clean-tree status")
    if build.get("exit_code") != 0:
        _fail("build-identity.exit_code", "must be zero")
    binaries = _object(build.get("binaries"), "build-identity.binaries")
    identities: dict[str, Any] = {}
    for phase in ("normal", "allocator"):
        identity = _object(binaries.get(phase), f"build-identity.binaries.{phase}")
        _binary_name(identity, f"build-identity.binaries.{phase}")
        _sha256(identity.get("sha256"), f"build-identity.binaries.{phase}.sha256")
        _integer(identity.get("bytes"), f"build-identity.binaries.{phase}.bytes", minimum=1)
        identities[phase] = identity
        binary_path = Path(identity["path"])
        if binary_path.is_file():
            actual = _sha256_file(binary_path, f"build-identity.binaries.{phase}.path")
            if actual != identity["sha256"].lower():
                _fail(f"build-identity.binaries.{phase}", "on-disk binary hash differs from identity")
    return build, identities


def _validate_capture_identity(
    capture: dict[str, Any],
    build: dict[str, Any],
    identities: dict[str, Any],
    path: str,
) -> None:
    if capture.get("revision") != build.get("revision"):
        _fail(f"{path}.revision", "does not match build-identity.revision")
    capture_binaries = _object(capture.get("binaries"), f"{path}.binaries")
    build_binaries = _object(build.get("binaries"), "build-identity.binaries")
    for phase in ("normal", "allocator"):
        captured = _object(capture_binaries.get(phase), f"{path}.binaries.{phase}")
        bound = _object(build_binaries.get(phase), f"build-identity.binaries.{phase}")
        for field in ("sha256", "bytes", "path"):
            if field not in captured:
                _fail(f"{path}.binaries.{phase}", f"missing {field!r}")
        captured_sha = _sha256(captured["sha256"], f"{path}.binaries.{phase}.sha256")
        captured_bytes = _integer(captured["bytes"], f"{path}.binaries.{phase}.bytes", minimum=1)
        captured_path = _string(captured["path"], f"{path}.binaries.{phase}.path")
        if captured_sha != bound["sha256"].lower():
            _fail(f"{path}.binaries.{phase}.sha256", "does not match build identity")
        if captured_bytes != bound["bytes"]:
            _fail(f"{path}.binaries.{phase}.bytes", "does not match build identity")
        if captured_path != bound["path"]:
            _fail(f"{path}.binaries.{phase}.path", "does not match build identity")
        if captured_sha != identities[phase]["sha256"].lower():
            _fail(f"{path}.binaries.{phase}.sha256", "does not match bound binary hash")


def _validate_environment(root: Path) -> None:
    environment = _object(
        _load_json(_required_file(root, "environment.json"), "environment.json"),
        "environment.json",
    )
    if environment.get("measurement_cpu") != EXPECTED_CPU:
        _fail("environment.measurement_cpu", f"must be {EXPECTED_CPU}")
    affinity = environment.get("affinity_cpus")
    if not isinstance(affinity, list) or EXPECTED_CPU not in affinity:
        _fail("environment.affinity_cpus", "must include measurement CPU 2")
    if not isinstance(environment.get("rustc"), str) or not environment["rustc"].startswith("rustc "):
        _fail("environment.rustc", "must retain rustc version output")
    if not isinstance(environment.get("cargo"), str) or not environment["cargo"].startswith("cargo "):
        _fail("environment.cargo", "must retain cargo version output")
    if not isinstance(environment.get("perf"), str) or not environment["perf"].startswith("perf "):
        _fail("environment.perf", "must retain perf version output")


def _validate_preflight_summary(root: Path) -> None:
    """Check the retained preflight index before validating its raw reports."""

    preflight = _object(
        _load_json(
            _required_file(root, "checks/preflight-verification.json"),
            "checks/preflight-verification.json",
        ),
        "checks/preflight-verification.json",
    )
    if preflight.get("result") != "pass":
        _fail("checks/preflight-verification.json.result", "must be pass")
    for field in ("selectors", "report_rows", "catalogs"):
        if preflight.get(field) != EXPECTED_SELECTOR_COUNT:
            _fail(
                f"checks/preflight-verification.json.{field}",
                f"must be {EXPECTED_SELECTOR_COUNT}",
            )
    validators = preflight.get("validators")
    required_validators = {
        "validate_perf_corpus_binding.validate_paths",
        "perf_compare.validate_parallel_metrics",
    }
    if not isinstance(validators, list) or not required_validators.issubset(set(validators)):
        _fail("checks/preflight-verification.json.validators", "does not retain both preflight validators")


def _validate_command(
    run: dict[str, Any],
    *,
    root: Path,
    matrix: dict[str, Any],
    phase: str,
    samples: int,
    warmups: int,
    identity: dict[str, Any],
) -> None:
    path = f"{run['_path']}.argv"
    argv = _list(run.get("argv"), path)
    if not argv or any(not isinstance(value, str) for value in argv):
        _fail(path, "must be a non-empty string list")
    if len(argv) < 9:
        _fail(path, "is too short for taskset/time/harness command")
    if Path(argv[0]).name != "taskset" or argv[1:3] != ["-c", str(EXPECTED_CPU)]:
        _fail(path, "must pin the child to CPU 2 with taskset -c 2")
    if Path(argv[3]).name != "time" or argv[4:6] != ["-v", "-o"]:
        _fail(path, "must collect verbose /usr/bin/time output")
    time_relative = run["time"]
    _portable_path(argv[6], time_relative, f"{path} time output")
    expected_binary_name = _binary_name(identity, f"build-identity.binaries.{phase}")
    if Path(argv[7]).name != expected_binary_name:
        _fail(path, f"must invoke the {phase} build binary")
    expected = ["--case", run["selector"], *_list(matrix.get("common_flags"), "matrix.common_flags")]
    expected.extend(("--samples", str(samples), "--warmup", str(warmups)))
    expected.extend(("--json", run["report"], "--corpus-manifest", run["catalog"]))
    tail = argv[8:]
    if len(tail) != len(expected):
        _fail(path, "harness flags do not match the predeclared command")
    for index, (actual, wanted) in enumerate(zip(tail, expected)):
        if wanted in {"--json", "--corpus-manifest"}:
            if actual != wanted:
                _fail(path, f"argument {index} must be {wanted}")
            continue
        if index > 0 and tail[index - 1] in {"--json", "--corpus-manifest"}:
            _portable_path(actual, wanted, f"{path} argument {wanted}")
            continue
        if actual != wanted:
            _fail(path, f"argument {index} {actual!r} does not match {wanted!r}")
    _portable_path(_argv_value(argv, "--json", path), run["report"], f"{path} --json")
    _portable_path(
        _argv_value(argv, "--corpus-manifest", path), run["catalog"], f"{path} --corpus-manifest"
    )


def _normalized_run(run: dict[str, Any], root: Path, identity_phase: str) -> dict[str, Any]:
    """Adapt a capture record to the summarizer's existing report contract."""

    normalized = copy.deepcopy(run)
    normalized["phase"] = identity_phase
    normalized["report_path"] = (root / run["report"]).resolve()
    normalized["catalog_path"] = (root / run["catalog"]).resolve()
    report_arg = _argv_value(run["argv"], "--json", f"{run['_path']}.argv")
    catalog_arg = _argv_value(run["argv"], "--corpus-manifest", f"{run['_path']}.argv")
    try:
        report_root = summarize_crud_baseline._historical_root(
            report_arg, run["report"], f"{run['_path']}.argv --json"
        )
        catalog_root = summarize_crud_baseline._historical_root(
            catalog_arg, run["catalog"], f"{run['_path']}.argv --corpus-manifest"
        )
    except Exception as error:
        _fail(run["_path"], f"historical argv paths are invalid: {error}")
    if report_root != catalog_root:
        _fail(run["_path"], "report and catalog argv paths use different historical roots")
    # The summarizer's private validator uses this already established root to
    # compare old absolute argv values with the bundle-relative manifest paths.
    normalized["historical_root"] = report_root
    return normalized


def _validate_report(
    run: dict[str, Any],
    *,
    root: Path,
    identity_phase: str,
    identity: dict[str, Any],
    revision: str,
    samples: int,
    warmups: int,
) -> dict[str, Any]:
    report = _object(_load_json(run["report_path"], run["report"]), run["report"])
    results = _list(report.get("results"), f"{run['report']}.results")
    if len(results) != 1:
        _fail(f"{run['report']}.results", "must contain exactly one result row")
    result = _object(results[0], f"{run['report']}.results[0]")
    if result.get("case") != run["selector"]:
        _fail(f"{run['report']}.results[0].case", "does not match selector")
    tool = _object(report.get("tool"), f"{run['report']}.tool")
    expected_instrumentation = "none" if identity_phase == "normal" else "system_allocator_operation_scoped"
    if tool.get("instrumentation") != expected_instrumentation:
        _fail(
            f"{run['report']}.tool.instrumentation",
            f"must be {expected_instrumentation!r} for {identity_phase}",
        )
    environment = _object(report.get("environment"), f"{run['report']}.environment")
    if environment.get("cpu_affinity") != str(EXPECTED_CPU):
        _fail(f"{run['report']}.environment.cpu_affinity", "must be '2'")
    if environment.get("git_revision") != revision:
        _fail(f"{run['report']}.environment.git_revision", "does not match build revision")
    if environment.get("git_worktree_dirty") is not False:
        _fail(f"{run['report']}.environment.git_worktree_dirty", "must be false")
    build_flags = _object(run["build_environment"], "build-identity.environment")
    if environment.get("rustflags") != build_flags.get("RUSTFLAGS"):
        _fail(f"{run['report']}.environment.rustflags", "does not match build identity")
    configuration = _object(report.get("configuration"), f"{run['report']}.configuration")
    expected_config = {
        "samples_per_case": samples,
        "warmup_iterations_per_case": warmups,
        "filesystem_cache_states": ["warm"],
        "execution_workers": [EXPECTED_WORKERS],
        "corpus_shapes": ["many-small"],
        "payload_kinds": ["compressible"],
        "writer_shapes": ["large"],
        "xlsx_shapes": ["medium"],
        "xlsx_cell_crud_shapes": ["medium"],
        "xlsx_row_visibility_shapes": ["medium"],
        "semantic_shapes": ["medium"],
    }
    for key, expected in expected_config.items():
        if configuration.get(key) != expected:
            _fail(f"{run['report']}.configuration.{key}", f"must equal {expected!r}")
    if configuration.get("cases") != [run["selector"]]:
        _fail(f"{run['report']}.configuration.cases", "must contain exactly the selected case")

    normalized = _normalized_run(run, root, identity_phase)
    try:
        record = summarize_crud_baseline._validate_report_and_select(
            normalized, revision, identity, samples, warmups
        )
    except Exception as error:
        _fail(run["report"], f"existing CRUD report/corpus validator rejected it: {error}")
    return record


def _expected_sequence(jobs: list[dict[str, Any]], phases: tuple[str, ...]) -> list[tuple[str, int, str]]:
    sequence: list[tuple[str, int, str]] = []
    for phase in phases:
        for repeat in (1, 2):
            selected = jobs if repeat == 1 else list(reversed(jobs))
            sequence.extend((phase, repeat, job["selector"]) for job in selected)
    return sequence


def _validate_runs(
    records: list[Any],
    *,
    root: Path,
    matrix: dict[str, Any],
    jobs: list[dict[str, Any]],
    revision: str,
    identities: dict[str, Any],
    build_environment: dict[str, Any],
    expected_sequence: list[tuple[str, int, str]],
    capture_label: str,
    preflight: bool = False,
) -> list[dict[str, Any]]:
    if len(records) != len(expected_sequence):
        _fail(f"{capture_label}.runs", f"must contain exactly {len(expected_sequence)} runs")
    validated: list[dict[str, Any]] = []
    seen_paths: set[str] = set()
    previous_finish: dt.datetime | None = None
    for index, value in enumerate(records):
        path = f"{capture_label}.runs[{index}]"
        run = _object(value, path)
        expected_phase, expected_repeat, expected_selector = expected_sequence[index]
        expected_record_phase = "preflight" if preflight else expected_phase
        if run.get("phase") != expected_record_phase:
            _fail(f"{path}.phase", f"must be {expected_record_phase!r} at capture position {index}")
        if run.get("repeat") != expected_repeat:
            _fail(f"{path}.repeat", f"must be {expected_repeat} at capture position {index}")
        if run.get("selector") != expected_selector:
            _fail(f"{path}.selector", f"must be {expected_selector!r} at capture position {index}")
        if run.get("exit_code") != 0:
            _fail(f"{path}.exit_code", "must be zero")
        started = _parse_time(run.get("started_utc"), f"{path}.started_utc")
        finished = _parse_time(run.get("finished_utc"), f"{path}.finished_utc")
        if finished < started:
            _fail(path, "finished_utc precedes started_utc")
        if previous_finish is not None and started < previous_finish:
            _fail(path, "run overlaps the preceding serialized run")
        previous_finish = finished
        report_rel, report_path = _relative_file(root, run.get("report"), f"{path}.report")
        catalog_rel, catalog_path = _relative_file(root, run.get("catalog"), f"{path}.catalog")
        time_rel, _ = _relative_file(root, run.get("time_v"), f"{path}.time_v")
        if report_rel in seen_paths or catalog_rel in seen_paths or time_rel in seen_paths:
            _fail(path, "report, catalog, or time-v path is reused")
        seen_paths.update((report_rel, catalog_rel, time_rel))
        expected_folder = "preflight" if preflight else expected_phase
        expected_stem = f"{expected_repeat}-{expected_selector}"
        expected_report = f"{expected_folder}/{expected_stem}.json"
        expected_catalog = f"{expected_folder}/{expected_stem}.catalog.json"
        expected_time = f"{expected_folder}/{expected_stem}.time.txt"
        if report_rel != expected_report or catalog_rel != expected_catalog or time_rel != expected_time:
            _fail(path, "report/catalog/time paths do not match the declared matrix position")
        run_data = {
            "_path": path,
            "phase": expected_phase,
            "repeat": expected_repeat,
            "selector": expected_selector,
            "report": report_rel,
            "report_path": report_path,
            "catalog": catalog_rel,
            "catalog_path": catalog_path,
            "time": time_rel,
            "argv": run.get("argv"),
            "build_environment": build_environment,
        }
        samples, warmups = (
            (1, 0)
            if preflight
            else (
                (EXPECTED_NORMAL_SAMPLES, EXPECTED_NORMAL_WARMUPS)
                if expected_phase == "normal"
                else (EXPECTED_ALLOCATOR_SAMPLES, EXPECTED_ALLOCATOR_WARMUPS)
            )
        )
        identity_phase = "normal" if preflight else expected_phase
        identity = identities[identity_phase]
        _validate_command(
            run_data,
            root=root,
            matrix=matrix,
            phase=identity_phase,
            samples=samples,
            warmups=warmups,
            identity=identity,
        )
        _validate_report(
            run_data,
            root=root,
            identity_phase=identity_phase,
            identity=identity,
            revision=revision,
            samples=samples,
            warmups=warmups,
        )
        validated.append(run_data)
    return validated


def _validate_capture_file(
    root: Path,
    *,
    filename: str,
    matrix: dict[str, Any],
    jobs: list[dict[str, Any]],
    build: dict[str, Any],
    identities: dict[str, Any],
    build_environment: dict[str, Any],
    preflight: bool,
) -> list[dict[str, Any]]:
    label = filename
    capture = _object(_load_json(_required_file(root, filename), label), label)
    _validate_capture_identity(capture, build, identities, label)
    records = _list(capture.get("runs"), f"{label}.runs")
    if preflight:
        expected = [("normal", 1, job["selector"]) for job in jobs]
    else:
        expected = _expected_sequence(jobs, ("normal", "allocator"))
    return _validate_runs(
        records,
        root=root,
        matrix=matrix,
        jobs=jobs,
        revision=build["revision"],
        identities=identities,
        build_environment=build_environment,
        expected_sequence=expected,
        capture_label=label,
        preflight=preflight,
    )


def _recompute_summary(root: Path) -> dict[str, Any]:
    try:
        return summarize_crud_baseline.summarize(
            root,
            samples=EXPECTED_NORMAL_SAMPLES,
            warmups=EXPECTED_NORMAL_WARMUPS,
            allocation_samples=EXPECTED_ALLOCATOR_SAMPLES,
            allocation_warmups=EXPECTED_ALLOCATOR_WARMUPS,
            drift_threshold_percent=5.0,
        )
    except Exception as error:
        _fail("summary", f"existing CRUD summarizer rejected the retained capture: {error}")
    raise AssertionError("unreachable")


def verify(root: Path, *, write_summary: bool = False) -> dict[str, Any]:
    root = root.resolve()
    if not root.is_dir():
        _fail("--root", f"not a directory: {root}")
    tooling = _object(_load_json(root / "tool-source-identity.json"), "tool-source-identity.json")
    files = _object(tooling.get("files"), "tool-source-identity.files")
    expected_tools = {"tools/summarize_crud_baseline.py", "tools/perf_compare.py",
                      "tools/validate_perf_corpus_binding.py"}
    if set(files) != expected_tools:
        _fail("tool-source-identity.files", "must bind exactly the three replay tools")
    for relative, digest in files.items():
        if _sha256_file(REPO_ROOT / relative, relative) != _sha256(digest, relative):
            _fail(relative, "replay tool source differs from the retained identity")
    matrix, jobs = _validate_matrix(root)
    build, identities = _validate_build_identity(root)
    _validate_environment(root)
    _validate_preflight_summary(root)
    build_environment = _object(build.get("environment"), "build-identity.environment")
    if build_environment.get("RUSTFLAGS") != "-C force-frame-pointers=yes -C force-unwind-tables=yes":
        _fail("build-identity.environment.RUSTFLAGS", "does not bind the frame-pointer/unwind build")
    if build_environment.get("RUSTUP_TOOLCHAIN") != "1.98.1":
        _fail("build-identity.environment.RUSTUP_TOOLCHAIN", "must be 1.98.1")
    capture = _validate_capture_file(
        root,
        filename="capture.json",
        matrix=matrix,
        jobs=jobs,
        build=build,
        identities=identities,
        build_environment=build_environment,
        preflight=False,
    )
    preflight = _validate_capture_file(
        root,
        filename="checks/preflight-capture.json",
        matrix=matrix,
        jobs=jobs,
        build=build,
        identities=identities,
        build_environment=build_environment,
        preflight=True,
    )

    summary = _recompute_summary(root)
    summary_path = root / "summary.json"
    retained = None
    if summary_path.is_file():
        retained = _load_json(summary_path, "summary.json")
        _same(retained, summary, "summary.json")
    elif not write_summary:
        _fail("summary.json", "is missing; rerun with --write-summary to retain the recomputed summary")
    if write_summary:
        summary_path.write_text(
            json.dumps(summary, indent=2, sort_keys=True, allow_nan=False) + "\n",
            encoding="utf-8",
        )
        retained = summary
    return {
        "status": "verified",
        "change": EXPECTED_CHANGE,
        "selectors": len(jobs),
        "capture_runs": len(capture),
        "preflight_runs": len(preflight),
        "summary": "written" if write_summary else "matched",
        "revision": build["revision"],
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--root",
        type=Path,
        default=Path(__file__).resolve().parent,
        help="change-0417 evidence directory (default: this script's directory)",
    )
    parser.add_argument(
        "--write-summary",
        action="store_true",
        help="write the exactly recomputed summary.json after validation",
    )
    args = parser.parse_args(argv)
    try:
        result = verify(args.root, write_summary=args.write_summary)
    except (VerificationError, OSError, TypeError, ValueError) as error:
        print(f"0417 verification failed: {error}", file=sys.stderr)
        return 1
    print(json.dumps(result, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
