#!/usr/bin/env python3
"""Analyze paired 0413 ``perf script --no-inline`` XLS profiles.

The input profiles are ``cycles:u`` samples from the one-cell owned-source
case.  Sample periods are retained as weighted event periods for CPU-stack
attribution.  They are not elapsed time and do not isolate the benchmark's
timed phase.  This postprocessor binds each profile to its report, command
journal, build identity, protocol, and captured source revision before
emitting the paired diagnostic summary.

The stack parser is reused from the committed 0412 attribution parser.  A
profile is rejected when that parser reports lost records, malformed or
non-cycle headers, unparsed frame lines, empty stacks, invalid periods, or an
unterminated block.  Unknown frames are retained and accounted for so their
presence is visible without silently discarding their sample weight.
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import re
import shlex
import subprocess
import sys
import tempfile
from collections import Counter
from pathlib import Path
from typing import Any, Callable, Iterable


SCHEMA = "0413-xls-paired-cpu-attribution-v1"
DEFAULT_CAPTURE = Path("/tmp/litchi-goal-0413-capture")
DEFAULT_OUTPUT = Path("/tmp/litchi-goal-0413-profile/profile-attribution.json")
DEFAULT_PARSER = Path(
    "/home/zhuhe/code/litchi/docs/performance/results/change-0412/attribution/attribute.py"
)
DEFAULT_CONTROL_BUILD = Path(
    "/tmp/litchi-goal-0413-control-binaries/identity.json"
)
DEFAULT_CANDIDATE_BUILD = Path(
    "/tmp/litchi-goal-0413-candidate-binaries/identity.json"
)
DEFAULT_REPO = Path("/home/zhuhe/code/litchi")
PROFILE_CASE = "xls_owned_source_open_one_cell"
EXPECTED_EVENT = "cycles:u"
EXPECTED_FREQUENCY = 999
EXPECTED_SAMPLES = 10000
EXPECTED_WARMUP = 20
EXPECTED_CPU = "2"
EXPECTED_FLAGS = "-C force-frame-pointers=yes -C force-unwind-tables=yes"
EXPECTED_BINARY = "litchi-perf-baseline"
EXPECTED_ORDER = ("control", "candidate")
SOURCE_OWNER = "litchi_xls::workbook::source::SourceBackedWorkbook"
EAGER_OWNER = "litchi_xls::workbook::model::Workbook"
KNOWN_ARCHIVE_SHA256 = "6a57231ba681bc7bdd38d447ebd5348ef3b1fefedeefb1e61c97f22faa074e53"
KNOWN_WORKBOOK_SHA256 = "c78e03ba3743132e04b08bf6f4579ceb1c112a22c441c1e036381d3e06c6d041"
KNOWN_ARCHIVE_BYTES = 16995840
KNOWN_WORKBOOK_BYTES = 80946
KNOWN_OUTPUT_SHA256 = "e726a50d216e6d71d7c53aabd23ab5e0d4677c3ef1f41fc35410143ebe6381c1"


class ProfileError(Exception):
    """The profile evidence failed a required identity or parser check."""


def fail(path: str, message: str) -> None:
    raise ProfileError(f"{path}: {message}")


def as_dict(value: Any, path: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        fail(path, "expected an object")
    return value


def as_list(value: Any, path: str) -> list[Any]:
    if not isinstance(value, list):
        fail(path, "expected an array")
    return value


def as_string(value: Any, path: str) -> str:
    if not isinstance(value, str) or not value:
        fail(path, "expected a non-empty string")
    return value


def as_int(value: Any, path: str, minimum: int = 0) -> int:
    if type(value) is not int or value < minimum:
        fail(path, f"expected an integer >= {minimum}")
    return value


def digest(value: Any, path: str) -> str:
    text = as_string(value, path).lower()
    if not re.fullmatch(r"[0-9a-f]{64}", text):
        fail(path, "expected a SHA-256 digest")
    return text


def sha256_file(path: Path) -> str:
    result = hashlib.sha256()
    with path.open("rb") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            result.update(chunk)
    return result.hexdigest()


def artifact_path(path: Path) -> Path:
    """Resolve a logical artifact name to its bundled plain or zstd file."""

    if path.is_file():
        return path
    if path.suffix != ".zst":
        compressed = Path(str(path) + ".zst")
        if compressed.is_file():
            return compressed
    return path


def file_binding(path: Path) -> dict[str, Any]:
    resolved = artifact_path(path).resolve()
    if not resolved.is_file():
        fail(str(path), "required artifact is missing")
    return {
        "path": str(resolved),
        "bytes": resolved.stat().st_size,
        "sha256": sha256_file(resolved),
    }


def read_json(path: Path) -> Any:
    chosen = artifact_path(path)
    if not chosen.is_file():
        fail(str(path), "required JSON artifact is missing")
    try:
        if chosen.suffix == ".zst":
            raw = subprocess.check_output(["zstd", "-q", "-dc", str(chosen)])
        else:
            raw = chosen.read_bytes()
        return json.loads(raw.decode("utf-8"))
    except (OSError, subprocess.CalledProcessError, UnicodeDecodeError, json.JSONDecodeError) as error:
        fail(str(chosen), f"cannot read JSON: {error}")


def load_parser(path: Path) -> Any:
    if not path.is_file():
        fail(str(path), "committed 0412 parser is missing")
    spec = importlib.util.spec_from_file_location("litchi_goal_0412_attribute", path)
    if spec is None or spec.loader is None:
        fail(str(path), "cannot load committed parser")
    module = importlib.util.module_from_spec(spec)
    try:
        sys.modules[spec.name] = module
        spec.loader.exec_module(module)
    except (OSError, ImportError, SyntaxError) as error:
        fail(str(path), f"cannot load committed parser: {error}")
    for name in ("parse_samples", "rank_symbols", "weighted_metric", "source_open"):
        if not callable(getattr(module, name, None)):
            fail(str(path), f"committed parser lacks {name}")
    return module


def protocol_identity(path: Path) -> dict[str, Any]:
    obj = as_dict(read_json(path), "protocol.json")
    if obj.get("change") != "0413":
        fail("protocol.change", "expected 0413")
    for field in ("control_revision", "candidate_revision"):
        revision = as_string(obj.get(field), f"protocol.{field}")
        if not re.fullmatch(r"[0-9a-fA-F]{40}", revision):
            fail(f"protocol.{field}", "expected a 40-digit hexadecimal revision")
    if obj["control_revision"].lower() == obj["candidate_revision"].lower():
        fail("protocol", "control and candidate revisions must differ")
    profile = as_dict(obj.get("profile"), "protocol.profile")
    if profile.get("case") != PROFILE_CASE:
        fail("protocol.profile.case", f"expected {PROFILE_CASE!r}")
    if profile.get("samples") != EXPECTED_SAMPLES:
        fail("protocol.profile.samples", "does not match 10000")
    if profile.get("warmup") != EXPECTED_WARMUP:
        fail("protocol.profile.warmup", "does not match 20")
    if profile.get("event") != EXPECTED_EVENT:
        fail("protocol.profile.event", "does not match cycles:u")
    if profile.get("frequency") != EXPECTED_FREQUENCY:
        fail("protocol.profile.frequency", "does not match 999")
    if profile.get("call_graph") != "fp,127":
        fail("protocol.profile.call_graph", "does not match fp,127")
    if profile.get("roles") != list(EXPECTED_ORDER):
        fail("protocol.profile.roles", "must be control,candidate")
    if obj.get("cpu") != 2:
        fail("protocol.cpu", "must be CPU 2")
    if obj.get("filesystem_cache") != "warm":
        fail("protocol.filesystem_cache", "must be warm")
    return obj


def build_identity(path: Path, role: str, protocol: dict[str, Any]) -> dict[str, Any]:
    obj = as_dict(read_json(path), f"{role}.build_identity")
    expected_revision = protocol[f"{role}_revision"].lower()
    revision = as_string(obj.get("revision"), f"{role}.build_identity.revision").lower()
    if revision != expected_revision:
        fail(f"{role}.build_identity.revision", "does not match protocol")
    if obj.get("source_status", "") != "":
        fail(f"{role}.build_identity.source_status", "build source was dirty")
    if obj.get("exit_code") != 0:
        fail(f"{role}.build_identity.exit_code", "build record is not successful")
    environment = obj.get("build_environment")
    if environment is not None:
        environment = as_dict(environment, f"{role}.build_identity.build_environment")
        if environment.get("RUSTUP_TOOLCHAIN") != "1.98.1":
            fail(f"{role}.build_identity.build_environment.RUSTUP_TOOLCHAIN", "must be 1.98.1")
        if environment.get("RUSTFLAGS") != EXPECTED_FLAGS:
            fail(f"{role}.build_identity.build_environment.RUSTFLAGS", "does not match profile flags")
    binaries = as_dict(obj.get("binaries"), f"{role}.build_identity.binaries")
    normal = as_dict(binaries.get(EXPECTED_BINARY), f"{role}.build_identity.binaries.{EXPECTED_BINARY}")
    binary_sha = digest(normal.get("sha256"), f"{role}.build_identity.binaries.{EXPECTED_BINARY}.sha256")
    binary_bytes = as_int(normal.get("bytes"), f"{role}.build_identity.binaries.{EXPECTED_BINARY}.bytes", 1)
    return {
        "revision": revision,
        "binary": {"sha256": binary_sha, "bytes": binary_bytes},
        "file": file_binding(path),
        "build_cwd": obj.get("build_cwd"),
    }


def option(argv: list[str], name: str, path: str) -> str:
    try:
        index = argv.index(name)
    except ValueError:
        fail(path, f"missing {name}")
    if index + 1 >= len(argv):
        fail(path, f"missing value after {name}")
    return str(argv[index + 1])


def command_identity(
    capture: Path,
    protocol: dict[str, Any],
    role: str,
    build: dict[str, Any],
) -> dict[str, Any]:
    records = read_json(capture / "commands.json")
    if isinstance(records, dict):
        records = records.get("commands", records.get("entries"))
    rows = as_list(records, "commands.json")
    label = f"{role}-profile"
    matches = [as_dict(row, f"commands.json[{index}]") for index, row in enumerate(rows) if isinstance(row, dict) and row.get("label") == label]
    if len(matches) != 1:
        fail("commands.json", f"expected exactly one {label} journal row")
    row = matches[0]
    path = f"commands.json.{label}"
    if row.get("variant") != role:
        fail(f"{path}.variant", f"expected {role!r}")
    if row.get("source_status", "") != "":
        fail(f"{path}.source_status", "command source was dirty")
    if row.get("exit_code") != 0:
        fail(f"{path}.exit_code", "profile command failed")
    protocol_sha = digest(row.get("protocol_sha256"), f"{path}.protocol_sha256")
    if protocol_sha != sha256_file(capture / "protocol.json"):
        fail(f"{path}.protocol_sha256", "does not match protocol.json")
    revision = as_string(row.get("revision"), f"{path}.revision").lower()
    if revision != protocol[f"{role}_revision"].lower() or revision != build["revision"]:
        fail(f"{path}.revision", "does not match protocol and build identity")
    binary_sha = digest(row.get("binary_sha256"), f"{path}.binary_sha256")
    if binary_sha != build["binary"]["sha256"]:
        fail(f"{path}.binary_sha256", "does not match build identity")
    argv = [str(value) for value in as_list(row.get("argv"), f"{path}.argv")]
    if not argv or argv[0] != "taskset":
        fail(f"{path}.argv", "must be pinned with taskset")
    if option(argv, "-c", path) != EXPECTED_CPU:
        fail(f"{path}.argv", "must be pinned to CPU 2")
    try:
        perf_index = next(index for index, value in enumerate(argv) if Path(value).name == "perf")
    except StopIteration:
        fail(f"{path}.argv", "does not invoke perf")
    if argv[perf_index + 1 : perf_index + 2] != ["record"]:
        fail(f"{path}.argv", "must use perf record")
    if "--no-buildid-cache" not in argv:
        fail(f"{path}.argv", "must disable build-id cache for this capture")
    if option(argv, "-e", path) != EXPECTED_EVENT:
        fail(f"{path}.argv", "event differs from protocol")
    if option(argv, "-F", path) != str(EXPECTED_FREQUENCY):
        fail(f"{path}.argv", "frequency differs from protocol")
    if option(argv, "--call-graph", path) != "fp,127":
        fail(f"{path}.argv", "call graph differs from protocol")
    data_path = Path(option(argv, "-o", path))
    if data_path.name != f"{label}.data":
        fail(f"{path}.argv", "perf data path does not identify the profile role")
    if "--" not in argv:
        fail(f"{path}.argv", "missing benchmark command separator")
    benchmark = argv[argv.index("--") + 1 :]
    if not any(Path(value).name == EXPECTED_BINARY for value in benchmark):
        fail(f"{path}.argv", "does not invoke the expected benchmark binary")
    if option(benchmark, "--filesystem-cache", path) != "warm":
        fail(f"{path}.argv", "filesystem cache differs from protocol")
    if option(benchmark, "--case", path) != PROFILE_CASE:
        fail(f"{path}.argv", "case differs from protocol")
    if option(benchmark, "--samples", path) != str(EXPECTED_SAMPLES):
        fail(f"{path}.argv", "sample count differs from protocol")
    if option(benchmark, "--warmup", path) != str(EXPECTED_WARMUP):
        fail(f"{path}.argv", "warmup differs from protocol")
    report_path = Path(option(benchmark, "--json", path))
    if report_path.name != f"{label}.json":
        fail(f"{path}.argv", "report path does not identify the profile role")
    catalog_path = Path(option(benchmark, "--corpus-manifest", path))
    if catalog_path.name != f"{label}.catalog.json":
        fail(f"{path}.argv", "catalog path does not identify the profile role")
    return {
        "label": label,
        "variant": role,
        "revision": revision,
        "binary_sha256": binary_sha,
        "argv": argv,
        "data_path": str(capture / data_path.name),
        "report_path": str(capture / report_path.name),
        "catalog_path": str(capture / catalog_path.name),
        "journal": row,
    }


def permutation(value: Any, count: int, path: str) -> list[int]:
    values = as_list(value, path)
    if len(values) != count or sorted(values) != list(range(count)):
        fail(path, f"must be an exact permutation of 0..{count - 1}")
    return [as_int(item, f"{path}[{index}]") for index, item in enumerate(values)]


def validate_report(path: Path, role: str, protocol: dict[str, Any], build: dict[str, Any]) -> dict[str, Any]:
    obj = as_dict(read_json(path), f"{role}.report")
    prefix = f"{role}.report"
    if obj.get("schema_version") != 1:
        fail(f"{prefix}.schema_version", "expected report schema 1")
    tool = as_dict(obj.get("tool"), f"{prefix}.tool")
    if tool.get("binary") != EXPECTED_BINARY or tool.get("profile") != "release":
        fail(f"{prefix}.tool", "unexpected benchmark binary/profile")
    binary = as_dict(obj.get("binary_identity"), f"{prefix}.binary_identity")
    if Path(as_string(binary.get("path"), f"{prefix}.binary_identity.path")).name != EXPECTED_BINARY:
        fail(f"{prefix}.binary_identity.path", "unexpected benchmark executable")
    report_sha = digest(binary.get("binary_sha256"), f"{prefix}.binary_identity.binary_sha256")
    report_bytes = as_int(binary.get("binary_bytes"), f"{prefix}.binary_identity.binary_bytes", 1)
    if report_sha != build["binary"]["sha256"] or report_bytes != build["binary"]["bytes"]:
        fail(f"{prefix}.binary_identity", "does not match build identity")
    environment = as_dict(obj.get("environment"), f"{prefix}.environment")
    revision = as_string(environment.get("git_revision"), f"{prefix}.environment.git_revision").lower()
    if revision != protocol[f"{role}_revision"].lower() or revision != build["revision"]:
        fail(f"{prefix}.environment.git_revision", "does not match protocol/build identity")
    if environment.get("git_worktree_dirty") is not False:
        fail(f"{prefix}.environment.git_worktree_dirty", "profile source was dirty")
    if not as_string(environment.get("rustc_version"), f"{prefix}.environment.rustc_version").startswith("rustc 1.98.1 "):
        fail(f"{prefix}.environment.rustc_version", "must use Rust 1.98.1")
    if environment.get("rustflags") != EXPECTED_FLAGS:
        fail(f"{prefix}.environment.rustflags", "does not match profile flags")
    if environment.get("cpu_affinity") != EXPECTED_CPU:
        fail(f"{prefix}.environment.cpu_affinity", "must be CPU 2")
    config = as_dict(obj.get("configuration"), f"{prefix}.configuration")
    if config.get("cases") != [PROFILE_CASE]:
        fail(f"{prefix}.configuration.cases", "must contain only the profile case")
    if config.get("samples_per_case") != EXPECTED_SAMPLES or config.get("warmup_iterations_per_case") != EXPECTED_WARMUP:
        fail(f"{prefix}.configuration", "sample/warmup count differs from protocol")
    if config.get("filesystem_cache_states") != ["warm"]:
        fail(f"{prefix}.configuration.filesystem_cache_states", "must be warm")
    if config.get("execution_workers") != [1]:
        fail(f"{prefix}.configuration.execution_workers", "must use one worker")
    results = as_list(obj.get("results"), f"{prefix}.results")
    if len(results) != 1:
        fail(f"{prefix}.results", "expected one profile result")
    result = as_dict(results[0], f"{prefix}.results[0]")
    if result.get("case") != PROFILE_CASE:
        fail(f"{prefix}.results[0].case", "does not match profile case")
    corpus = as_dict(result.get("corpus"), f"{prefix}.results[0].corpus")
    archive_sha = digest(corpus.get("archive_sha256"), f"{prefix}.results[0].corpus.archive_sha256")
    target_sha = digest(corpus.get("target_payload_sha256"), f"{prefix}.results[0].corpus.target_payload_sha256")
    archive_bytes = as_int(corpus.get("archive_bytes"), f"{prefix}.results[0].corpus.archive_bytes", 1)
    target_bytes = as_int(corpus.get("target_payload_bytes"), f"{prefix}.results[0].corpus.target_payload_bytes", 1)
    output_sha = digest(result.get("output_sha256"), f"{prefix}.results[0].output_sha256")
    if archive_sha != KNOWN_ARCHIVE_SHA256 or archive_bytes != KNOWN_ARCHIVE_BYTES:
        fail(f"{prefix}.results[0].corpus", "does not match the independent archive oracle")
    if target_sha != KNOWN_WORKBOOK_SHA256 or target_bytes != KNOWN_WORKBOOK_BYTES:
        fail(f"{prefix}.results[0].corpus", "does not match the independent Workbook oracle")
    if output_sha != KNOWN_OUTPUT_SHA256:
        fail(f"{prefix}.results[0].output_sha256", "does not match the independent one-cell oracle")
    elapsed = as_dict(result.get("elapsed_ns"), f"{prefix}.results[0].elapsed_ns")
    samples = as_list(elapsed.get("samples"), f"{prefix}.results[0].elapsed_ns.samples")
    if len(samples) != EXPECTED_SAMPLES:
        fail(f"{prefix}.results[0].elapsed_ns.samples", "does not match profile sample count")
    for index, sample in enumerate(samples):
        if isinstance(sample, bool) or not isinstance(sample, (int, float)):
            fail(f"{prefix}.results[0].elapsed_ns.samples[{index}]", "expected a numeric sample")
    order = permutation(elapsed.get("sample_order"), EXPECTED_SAMPLES, f"{prefix}.results[0].elapsed_ns.sample_order")
    operation = as_dict(result.get("operation_metrics"), f"{prefix}.results[0].operation_metrics")
    operation_indices = permutation(operation.get("sample_indices"), EXPECTED_SAMPLES, f"{prefix}.results[0].operation_metrics.sample_indices")
    if operation_indices != order:
        fail(f"{prefix}.results[0].operation_metrics.sample_indices", "does not equal elapsed sample_order")
    catalog = as_dict(read_json(path.with_name(f"{role}-profile.catalog.json")), f"{role}.catalog")
    if catalog.get("manifest_version") != 2:
        fail(f"{role}.catalog.manifest_version", "expected catalog manifest version 2")
    catalog_sha = digest(catalog.get("catalog_sha256"), f"{role}.catalog.catalog_sha256")
    reference = obj.get("corpus_catalog")
    if reference is not None:
        reference = as_dict(reference, f"{prefix}.corpus_catalog")
        if reference.get("catalog_sha256") != catalog_sha:
            fail(f"{prefix}.corpus_catalog.catalog_sha256", "does not match catalog sidecar")
    bindings = as_list(catalog.get("case_bindings"), f"{role}.catalog.case_bindings")
    if not any(isinstance(item, dict) and item.get("case") == PROFILE_CASE for item in bindings):
        fail(f"{role}.catalog.case_bindings", "missing profile case binding")
    # The sidecar is accepted only when it carries the same corpus hashes as
    # the report.  This binds the stack capture to the exact generated input.
    catalog_text = json.dumps(catalog, sort_keys=True)
    if archive_sha not in catalog_text or target_sha not in catalog_text:
        fail(f"{role}.catalog", "does not carry report corpus digests")
    return {
        "file": file_binding(path),
        "revision": revision,
        "binary": {"sha256": report_sha, "bytes": report_bytes},
        "corpus": {
            "archive_sha256": archive_sha,
            "target_payload_sha256": target_sha,
            "archive_bytes": archive_bytes,
            "target_payload_bytes": target_bytes,
        },
        "output_sha256": output_sha,
        "catalog": {"file": file_binding(path.with_name(f"{role}-profile.catalog.json")), "catalog_sha256": catalog_sha},
        "sample_order_sha256": hashlib.sha256(json.dumps(order, separators=(",", ":")).encode()).hexdigest(),
    }


def source_binding(repo: Path, revision: str, role: str, require_head: bool = True) -> dict[str, Any]:
    checks = {
        "tools/perf-baseline/src/lib.rs": {
            "run_xls_source_backed_case": r"fn\s+run_xls_source_backed_case\s*\(",
            "instrumented_source": r"struct\s+InstrumentedSource\s*\{",
        },
        "crates/litchi-xls/src/workbook/source.rs": {
            "source_backed_open": r"pub\s+fn\s+from_read_at\s*\(",
            "source_backed_cell": r"pub\s+fn\s+cell_value_by_index\s*\(",
        },
        "crates/litchi-xls/src/workbook/package.rs": {
            "eager_new": r"pub\s+fn\s+new\s*\(",
            "eager_from_ole_file": r"pub\s+fn\s+from_ole_file\s*\(",
        },
        "crates/litchi-cfb/src/file.rs": {
            "try_push": r"fn\s+try_push\s*<",
            "collect_exact": r"fn\s+collect_exact\s*\(",
            "validation": r"fn\s+validate_(?:stream_allocations|physical_sector_layout)\s*\(",
        },
    }
    result: dict[str, Any] = {"role": role, "repo": str(repo), "revision": revision, "status": "bound", "files": []}
    if not repo.is_dir():
        result["status"] = "unavailable"
        result["message"] = "captured worktree is unavailable; profile symbols remain bound to build identity only"
        return result
    try:
        actual = subprocess.check_output(["git", "-C", str(repo), "rev-parse", "HEAD"], text=True).strip().lower()
    except (OSError, subprocess.CalledProcessError) as error:
        result["status"] = "unavailable"
        result["message"] = f"cannot read captured git revision: {error}"
        return result
    result["repo_head"] = actual
    result["head_check"] = "required" if require_head else "skipped_for_explicit_repo"
    if require_head and actual != revision.lower():
        fail(f"source.{role}.revision", "worktree HEAD does not match build identity")
    for relative, file_checks in checks.items():
        try:
            blob = subprocess.check_output(["git", "-C", str(repo), "show", f"{revision}:{relative}"], stderr=subprocess.STDOUT)
        except (OSError, subprocess.CalledProcessError) as error:
            fail(f"source.{role}.{relative}", f"cannot read captured git blob: {error}")
        text_blob = blob.decode("utf-8", "replace")
        found = {
            name: bool(re.search(pattern, text_blob)) for name, pattern in file_checks.items()
        }
        if not all(found.values()):
            fail(f"source.{role}.{relative}", f"required source markers absent: {found}")
        result["files"].append({
            "path": relative,
            "bytes": len(blob),
            "sha256": hashlib.sha256(blob).hexdigest(),
            "checks": found,
        })
    return result


def metric(period: int, blocks: int, subset: int, whole: int) -> dict[str, Any]:
    return {
        "weighted_event_period": period,
        "raw_stack_blocks": blocks,
        "share_of_subset_percent": period / subset * 100.0 if subset else None,
        "share_of_whole_process_percent": period / whole * 100.0 if whole else None,
    }


def rank(
    samples: Iterable[Any],
    predicate: Callable[[Any], bool],
    subset: int,
    whole: int,
    top: int,
    inclusive: bool,
) -> list[dict[str, Any]]:
    periods: Counter[str] = Counter()
    blocks: Counter[str] = Counter()
    for sample in samples:
        if not predicate(sample):
            continue
        symbols = set(sample.symbols) if inclusive else {sample.symbols[0] if sample.symbols else "<empty-stack>"}
        for symbol in symbols:
            periods[symbol] += sample.period
            blocks[symbol] += 1
    return [
        {"symbol": symbol, **metric(period, blocks[symbol], subset, whole)}
        for symbol, period in periods.most_common(top)
    ]


def subset(
    samples: list[Any],
    name: str,
    predicate: Callable[[str], bool],
    whole: int,
    top: int,
) -> dict[str, Any]:
    matching = [sample for sample in samples if any(predicate(symbol) for symbol in sample.symbols)]
    period = sum(sample.period for sample in matching)
    sample_predicate = lambda sample: any(predicate(symbol) for symbol in sample.symbols)
    return {
        "name": name,
        "scope": metric(period, len(matching), period, whole),
        "marker_observed": bool(matching),
        "leaf_period_weighted_ranking": rank(samples, sample_predicate, period, whole, top, False),
        "inclusive_period_weighted_ranking": rank(samples, sample_predicate, period, whole, top, True),
        "observed_marker_symbols": sorted({symbol for sample in matching for symbol in sample.symbols if predicate(symbol)}),
    }


def source_open(symbol: str) -> bool:
    return re.search(
        rf"(?:<)?{re.escape(SOURCE_OWNER)}(?:>|)::from_read_at(?:$|::|\{{)",
        symbol,
    ) is not None


def source_cell(symbol: str) -> bool:
    return re.search(
        rf"(?:<)?{re.escape(SOURCE_OWNER)}(?:>|)::cell_value_by_index(?:$|::|\{{)",
        symbol,
    ) is not None


def eager_method(symbol: str, method: str) -> bool:
    return EAGER_OWNER in symbol and re.search(
        rf"::{re.escape(method)}(?:$|::|\{{)", symbol
    ) is not None


def symbol_contains(*needles: str) -> Callable[[str], bool]:
    lowered = tuple(needle.lower() for needle in needles)
    return lambda symbol: all(needle in symbol.lower() for needle in lowered)


def diagnostic_subsets() -> dict[str, Callable[[str], bool]]:
    return {
        "source_backed_from_read_at": source_open,
        "source_backed_cell_value_by_index": source_cell,
        "eager_workbook_new": lambda symbol: eager_method(symbol, "new"),
        "eager_workbook_from_ole_file": lambda symbol: eager_method(symbol, "from_ole_file"),
        "try_push": symbol_contains("::try_push"),
        "collect_exact": symbol_contains("::collect_exact"),
        "validation": symbol_contains("litchi_cfb", "validate"),
    }


def diagnostic_source_helpers() -> dict[str, Callable[[str], bool]]:
    """Selectors for source/CFB adapter helpers, independent of phase timing."""

    return {
        "instrumented_source_read_at": lambda symbol: "instrumentedsource" in symbol.lower()
        and "read_at" in symbol.lower(),
        "read_at_trait_helper": lambda symbol: "litchi_core::source::readat" in symbol.lower()
        and "read_at" in symbol.lower(),
        "shared_ole_open_with_limits": symbol_contains(
            "litchi_cfb::shared::sharedolefile", "open_with_limits"
        ),
        "shared_ole_read_stream_range": symbol_contains(
            "litchi_cfb::shared::sharedolefile", "read_stream_range"
        ),
        "shared_ole_stream_cursor_at": symbol_contains(
            "litchi_cfb::shared::sharedolefile", "stream_cursor_at"
        ),
        "shared_ole_stream_read_exact": symbol_contains(
            "litchi_cfb::shared::sharedolestreamcursor", "read_exact"
        ),
        "shared_ole_stream_move_forward": symbol_contains(
            "litchi_cfb::shared::sharedolestreamcursor", "move_forward"
        ),
        "source_query_cell": symbol_contains(
            "litchi_xls::workbook::source::query_cell"
        ),
    }


def parse_profile(path: Path, parser: Any) -> tuple[list[Any], dict[str, Any]]:
    if not path.is_file():
        fail(str(path), "perf script input is missing")
    try:
        samples, stats = parser.parse_samples(path)
    except (OSError, UnicodeError, ValueError) as error:
        fail(str(path), f"parser failed: {error}")
    reject = {
        "lost_sample_count": "lost perf samples",
        "lost_metadata_lines": "lost metadata",
        "unquantified_lost_lines": "unquantified lost metadata",
        "malformed_cycle_headers": "malformed cycle headers",
        "invalid_cycle_event_headers": "non-cycle event headers",
        "invalid_cycle_period_headers": "invalid cycle periods",
        "invalid_sample_blocks": "invalid sample blocks",
        "unparsed_frame_lines": "unparsed frame lines",
        "empty_stack_blocks": "empty stacks",
        "unterminated_blocks": "unterminated stacks",
        "zero_or_negative_period_samples": "zero or negative periods",
    }
    for key, label in reject.items():
        if stats.get(key, 0):
            fail(str(path), f"{label}: {stats[key]}")
    if stats.get("cycle_headers", 0) == 0 or not samples:
        fail(str(path), "no valid cycles:u sample blocks")
    if any(getattr(sample, "event", EXPECTED_EVENT) != EXPECTED_EVENT for sample in samples):
        fail(str(path), "parsed a non-cycles:u sample")
    return samples, stats


def parser_summary(samples: list[Any], stats: dict[str, Any], top: int) -> dict[str, Any]:
    whole = sum(sample.period for sample in samples)
    unknown_period = sum(
        sample.period for sample in samples if any(parser_unknown(symbol) for symbol in sample.symbols)
    )
    rows: dict[str, Any] = {}
    for name, predicate in diagnostic_subsets().items():
        rows[name] = subset(samples, name, predicate, whole, top)
    helpers: dict[str, Any] = {}
    for name, predicate in diagnostic_source_helpers().items():
        helpers[name] = subset(samples, name, predicate, whole, top)
    return {
        "sample_parser": {
            **stats,
            "event": EXPECTED_EVENT,
            "raw_stack_blocks": len(samples),
            "whole_process_weighted_event_period": whole,
            "unknown_frame_period": unknown_period,
            "unknown_frame_share_of_whole_process_percent": unknown_period / whole * 100.0 if whole else None,
        },
        "whole_process": {
            "scope": "all parsed cycles:u stack samples in this perf-script export",
            "weighted_event_period": whole,
            "raw_stack_blocks": len(samples),
            "leaf_period_weighted_ranking": rank(samples, lambda _sample: True, whole, whole, top, False),
            "inclusive_period_weighted_ranking": rank(samples, lambda _sample: True, whole, whole, top, True),
        },
        "diagnostic_subsets": rows,
        "diagnostic_source_helpers": helpers,
    }


def parser_unknown(symbol: str) -> bool:
    normalized = symbol.strip().lower()
    return normalized in {"??", "unknown", "<unknown>", "[unknown]"} or normalized.startswith("[unknown")


def source_repo(build: dict[str, Any], fallback: Path) -> Path:
    cwd = build.get("build_cwd")
    if isinstance(cwd, str) and Path(cwd).is_dir():
        return Path(cwd)
    return fallback


def make_script_from_data(data: Path, role: str, output_dir: Path) -> Path:
    if not data.is_file():
        fail(str(data), "perf data input is missing")
    output_dir.mkdir(parents=True, exist_ok=True)
    output = output_dir / f"{role}-profile-script.stdout"
    try:
        with output.open("wb") as sink:
            result = subprocess.run(
                [
                    "perf",
                    "script",
                    "--no-inline",
                    "-i",
                    str(data),
                ],
                stdout=sink,
                stderr=subprocess.PIPE,
                check=False,
            )
    except OSError as error:
        fail(str(data), f"cannot postprocess perf data: {error}")
    if result.returncode != 0:
        fail(str(data), f"perf script failed with exit {result.returncode}: {result.stderr.decode(errors='replace').strip()}")
    return output


def decompress_artifact(path: Path, role: str, suffix: str, output_dir: Path) -> Path:
    """Materialize a zstd bundle member for tools that require a file path."""

    if path.suffix != ".zst":
        return path
    output_dir.mkdir(parents=True, exist_ok=True)
    output = output_dir / f"{role}-profile.{suffix}"
    try:
        with output.open("wb") as sink:
            result = subprocess.run(
                ["zstd", "-q", "-dc", str(path)],
                stdout=sink,
                stderr=subprocess.PIPE,
                check=False,
            )
    except OSError as error:
        fail(str(path), f"cannot decompress zstd artifact: {error}")
    if result.returncode != 0:
        try:
            output.unlink()
        except OSError:
            pass
        fail(str(path), f"zstd decompression failed with exit {result.returncode}: {result.stderr.decode(errors='replace').strip()}")
    return output


def profile_input(path: Path, data: Path, role: str, output_dir: Path) -> tuple[Path, dict[str, Any]]:
    data = artifact_path(data)
    if not data.is_file():
        fail(str(data), "perf data input is missing")
    path = artifact_path(path)
    files: dict[str, Any] = {"data": file_binding(data)}
    if path.is_file():
        files["script"] = file_binding(path)
        if path.suffix == ".zst":
            materialized = decompress_artifact(path, role, "script.txt", output_dir)
            files["script_decompressed"] = file_binding(materialized)
            return materialized, {**files, "postprocessed": False}
        return path, {**files, "postprocessed": False}
    perf_data = data
    if data.suffix == ".zst":
        perf_data = decompress_artifact(data, role, "data", output_dir)
        files["data_decompressed"] = file_binding(perf_data)
    generated = make_script_from_data(perf_data, role, output_dir)
    return generated, {**files, "script": file_binding(generated), "postprocessed": True}


def default_script_path(capture: Path, role: str) -> Path:
    """Use the frozen postprocess name, with the earlier 0412 name as fallback."""

    for name in (
        f"{role}-profile.script.txt",
        f"{role}-profile.script.txt.zst",
        f"{role}-profile-script.stdout",
        f"{role}-profile-script.stdout.zst",
    ):
        candidate = capture / name
        if candidate.is_file():
            return candidate
    return capture / f"{role}-profile.script.txt"


def default_build_identity(capture: Path, role: str) -> Path:
    """Prefer the checks sidecar beside a published capture bundle."""

    bundled = capture.parent / "checks" / f"{role}-build-identity.json"
    if bundled.is_file():
        return bundled
    return DEFAULT_CONTROL_BUILD if role == "control" else DEFAULT_CANDIDATE_BUILD


def parser_input(path: Path, repo: Path | None) -> Path:
    """Accept the compact replay token used by the evidence instructions."""

    if path == Path("committed0412"):
        root = repo or DEFAULT_REPO
        return root / "docs/performance/results/change-0412/attribution/attribute.py"
    return path


def override_artifact(logical: Path, override: Path | None, role: str, kind: str) -> Path:
    if override is None:
        return logical
    override = override.resolve()
    logical_name = logical.name.removesuffix(".zst")
    accepted = {logical_name, f"{logical_name}.zst"}
    if override.name not in accepted:
        fail(f"{role}.{kind}", f"override must retain journal filename {logical_name!r}")
    return override


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--capture", type=Path, default=DEFAULT_CAPTURE)
    parser.add_argument("--control-script", type=Path)
    parser.add_argument("--candidate-script", type=Path)
    parser.add_argument("--control-data", type=Path)
    parser.add_argument("--candidate-data", type=Path)
    parser.add_argument("--control-report", type=Path)
    parser.add_argument("--candidate-report", type=Path)
    parser.add_argument("--control-build-identity", type=Path)
    parser.add_argument("--candidate-build-identity", type=Path)
    parser.add_argument("--control-repo", type=Path)
    parser.add_argument("--candidate-repo", type=Path)
    parser.add_argument("--parser", type=Path, default=DEFAULT_PARSER)
    parser.add_argument("--repo", type=Path)
    parser.add_argument("--top", type=int, default=40)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    args = parser.parse_args(argv)
    if args.top <= 0:
        raise SystemExit("--top must be positive")
    capture = args.capture.resolve()
    output = args.output.resolve()
    protocol = protocol_identity(capture / "protocol.json")
    parser_path = parser_input(args.parser, args.repo).resolve()
    parser_module = load_parser(parser_path)
    roles: dict[str, Any] = {}
    for role in EXPECTED_ORDER:
        build_path = args.control_build_identity if role == "control" else args.candidate_build_identity
        build_path = build_path or default_build_identity(capture, role)
        build = build_identity(build_path.resolve(), role, protocol)
        command = command_identity(capture, protocol, role, build)
        report_path = (args.control_report if role == "control" else args.candidate_report)
        report = validate_report((report_path or capture / f"{role}-profile.json").resolve(), role, protocol, build)
        journal_data_path = Path(command["data_path"])
        data_arg = args.control_data if role == "control" else args.candidate_data
        data_path = override_artifact(journal_data_path, data_arg, role, "data")
        script_arg = args.control_script if role == "control" else args.candidate_script
        journal_script_path = default_script_path(capture, role)
        script_path = override_artifact(journal_script_path, script_arg, role, "script").resolve()
        script_path, profile_files = profile_input(script_path, data_path, role, output.parent)
        samples, stats = parse_profile(script_path, parser_module)
        role_repo_arg = args.control_repo if role == "control" else args.candidate_repo
        if role_repo_arg is not None:
            source_root = role_repo_arg.resolve()
            source_requires_head = False
        elif args.repo is not None:
            source_root = args.repo.resolve()
            source_requires_head = False
        else:
            source_root = source_repo(build, DEFAULT_REPO)
            source_requires_head = True
        roles[role] = {
            "build_identity": build,
            "command": command,
            "report": report,
            "profile_files": profile_files,
            "source_binding": source_binding(
                source_root,
                build["revision"],
                role,
                source_requires_head,
            ),
            "attribution": parser_summary(samples, stats, args.top),
        }
    control_report = roles["control"]["report"]
    candidate_report = roles["candidate"]["report"]
    if control_report["corpus"] != candidate_report["corpus"]:
        fail("reports", "control and candidate corpus identities differ")
    if control_report["output_sha256"] != candidate_report["output_sha256"]:
        fail("reports", "control and candidate semantic output digests differ")
    output.parent.mkdir(parents=True, exist_ok=True)
    result = {
        "schema": SCHEMA,
        "status": "pass",
        "timing_semantics": {
            "phase_latency": False,
            "weight": "perf cycles:u sample period",
            "whole_process_scope": "all parsed process samples in each perf-script export",
            "production_scope": "inclusive presence of an exact demangled production ancestor",
        },
        "contract": {
            "case": PROFILE_CASE,
            "event": EXPECTED_EVENT,
            "frequency_hz": EXPECTED_FREQUENCY,
            "samples": EXPECTED_SAMPLES,
            "warmup": EXPECTED_WARMUP,
            "call_graph": "fp,127",
            "perf_script_format": "default perf script --no-inline output: comm,pid,time,period,event header followed by leaf-to-root ip/symbol/dso frames",
            "no_inline": True,
            "cpu": EXPECTED_CPU,
            "filesystem_cache": "warm",
        },
        "parser": file_binding(parser_path),
        "protocol": {
            "file": file_binding(capture / "protocol.json"),
            "change": protocol["change"],
            "control_revision": protocol["control_revision"],
            "candidate_revision": protocol["candidate_revision"],
        },
        "shared_report_identity": {
            "corpus": control_report["corpus"],
            "output_sha256": control_report["output_sha256"],
            "control_candidate_equal": True,
        },
        "roles": roles,
        "limitations": [
            "Weighted values are cycles:u sample periods, not elapsed time or phase latency.",
            "Whole-process stacks include startup, corpus/setup, source construction, warmups, timed calls, and post-operation work.",
            "Calls to the same constructors outside the timed bracket can share callchains; ancestor subsets cannot fully separate them.",
            "Unknown frames are retained and weighted in whole-process totals; missing production markers mean unobserved stacks, not zero work.",
            "This paired profile is descriptive CPU attribution and makes no production speedup, regression, or phase-cost claim.",
        ],
    }
    output.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    print(json.dumps({
        "output": str(output),
        "status": "pass",
        "control_period": roles["control"]["attribution"]["sample_parser"]["whole_process_weighted_event_period"],
        "candidate_period": roles["candidate"]["attribution"]["sample_parser"]["whole_process_weighted_event_period"],
        "control_blocks": roles["control"]["attribution"]["sample_parser"]["raw_stack_blocks"],
        "candidate_blocks": roles["candidate"]["attribution"]["sample_parser"]["raw_stack_blocks"],
    }, sort_keys=True))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except ProfileError as error:
        print(f"litchi-goal-0413-profile: FAIL: {error}", file=sys.stderr)
        raise SystemExit(2)
