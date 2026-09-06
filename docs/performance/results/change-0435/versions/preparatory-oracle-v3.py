#!/usr/bin/env python3
"""Fail-closed oracle for one 0435 fresh ODT paragraph report.

The Rust harness reports one selected ODT shape per invocation.  This checker
reconstructs the paragraph fixture and its semantic digest independently,
checks the package/catalog identity and the measured resource vectors, and
returns a compact identity record for the outer three-role verifier.  It does
not run Cargo, Git, a workload, or a filesystem probe.

Archive bytes are intentionally treated as role-local evidence.  The ODT
semantic projection, styles/meta defaults, member topology, and package gates
must agree across roles; lexical content.xml bytes remain role-local because
the buffered and streaming writers may choose different schema-optional
children or escaping.  A ZIP hash is never used as a cross-role semantic
equivalence claim.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import math
from pathlib import Path
import re
import shlex
import sys
from typing import Any, Iterable


CHANGE = 435
SCHEMA_VERSION = 1
CPU = 2
WORKERS = 1
FORMAL_SAMPLES = 30
FORMAL_WARMUPS = 3
# The formal matrix is 30/3.  Small pre-apply pilots are allowed to invoke
# this same oracle with explicit 3/1 values; the override is local to one
# report and never changes the frozen protocol dimensions.
SAMPLES = FORMAL_SAMPLES
WARMUPS = FORMAL_WARMUPS
SHAPES = {"tiny": 64, "medium": 8_192, "large": 32_768}
ROLES = ("before-buffered", "after-buffered", "after-streaming")
MODES = ("normal", "allocator")
ALLOCATOR_REVISION = "serialized_region_peak_v3"
ALLOCATOR_SCOPE = "operation_global_system_allocator"
ALIGNMENT = "elapsed_ns.samples_by_elapsed_then_sample_index"
LATENCY_CLAIM = "comparable_timed_operation"
U64_MAX = (1 << 64) - 1
MAX_JSON_BYTES = 512 * 1024 * 1024
HEX40 = re.compile(r"^[0-9a-fA-F]{40}$")
HEX64 = re.compile(r"^[0-9a-fA-F]{64}$")

ALLOCATOR_FIELDS = (
    "allocation_calls", "deallocation_calls", "reallocation_calls",
    "failed_allocation_calls", "allocated_bytes", "deallocated_bytes",
    "live_bytes_before", "live_bytes_after", "peak_live_bytes_before",
    "peak_live_bytes_after", "region_peak_live_bytes",
)
PROCESS_FIELDS = (
    "user_cpu_ticks", "system_cpu_ticks", "clock_ticks_per_second",
    "minor_faults", "major_faults", "voluntary_context_switches",
    "nonvoluntary_context_switches", "rss_delta_bytes", "peak_rss_bytes",
    "rchar", "wchar", "read_bytes", "write_bytes",
    "cancelled_write_bytes", "syscr", "syscw",
)
SINK_VECTOR_FIELDS = ("accepted_bytes", "write_calls", "largest_write")
SINK_BUCKET_FIELDS = (
    "bytes_0", "bytes_1_to_512", "bytes_513_to_4096",
    "bytes_4097_to_16384", "bytes_16385_to_65536", "bytes_over_65536",
)
ODT_MEMBERS = (
    "mimetype", "content.xml", "styles.xml", "meta.xml",
    "META-INF/manifest.xml",
)
ODT_MIMETYPE = "application/vnd.oasis.opendocument.text"
ODT_COMPRESSION = "mimetype=stored;xml=deflate"
ODT_PAYLOAD_KIND = "deterministic-mixed-unicode-entities-plain-paragraphs"
ODT_DEFAULT_STYLES_SHA256 = "4d1dfc46e4c369722193548ff59269167f3e838a4b61bf1ddca1ecafe5092277"
ODT_DEFAULT_META_SHA256 = "c7e55a3560c73aa42da85eec4751c3e78b5cc53ff964f50acba6c5cd105e6719"
ODT_MANIFEST_ENTRIES = ("/", "content.xml", "styles.xml", "meta.xml")
TEXT_VARIANTS = (
    "plain paragraph",
    "Unicode café Δ 中",
    'entities <&> "quoted"',
    "mixed façade Ω <&> value",
)
TEXT_CONTRACT = "four-cycle plain/Unicode/XML-significant UTF-8 text with single interior spaces; one logical text run per paragraph"
SEMANTIC_DOMAIN = b"litchi-odt-buffered-semantic-v1\0"
T_CRITICAL_95 = {2: 4.303, 29: 2.045}


class VerificationError(ValueError):
    """A retained report violates the 0435 report contract."""


def fail(path: str, message: str) -> None:
    raise VerificationError(f"{path}: {message}")


def obj(value: Any, path: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        fail(path, "expected an object")
    return value


def array(value: Any, path: str) -> list[Any]:
    if not isinstance(value, list):
        fail(path, "expected an array")
    return value


def text(value: Any, path: str, *, allow_empty: bool = False) -> str:
    if not isinstance(value, str) or (not allow_empty and not value):
        fail(path, "expected a non-empty string" if not allow_empty else "expected a string")
    return value


def boolean(value: Any, path: str) -> bool:
    if not isinstance(value, bool):
        fail(path, "expected a boolean")
    return value


def u64(value: Any, path: str, *, positive: bool = False) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < (1 if positive else 0) or value > U64_MAX:
        fail(path, "expected a positive u64" if positive else "expected a u64")
    return value


def finite(value: Any, path: str, *, nonnegative: bool = True) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        fail(path, "expected a finite number")
    number = float(value)
    if not math.isfinite(number) or (nonnegative and number < 0):
        fail(path, "expected a finite number in the permitted range")
    return number


def digest(value: Any, path: str) -> str:
    value = text(value, path)
    if HEX64.fullmatch(value) is None:
        fail(path, "expected a SHA-256 hexadecimal digest")
    return value.lower()


def exact_keys(value: dict[str, Any], required: Iterable[str], optional: Iterable[str], path: str) -> None:
    required_set = set(required)
    allowed = required_set | set(optional)
    missing = sorted(required_set - set(value))
    unknown = sorted(set(value) - allowed)
    if missing or unknown:
        details = []
        if missing:
            details.append(f"missing={missing}")
        if unknown:
            details.append(f"unknown={unknown}")
        fail(path, "keys mismatch: " + ", ".join(details))


def reject_duplicate_pairs(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise VerificationError(f"duplicate JSON object key: {key}")
        result[key] = value
    return result


def reject_constant(value: str) -> Any:
    raise VerificationError(f"non-finite JSON constant is not permitted: {value}")


def load_json(path: Path, label: str | None = None) -> Any:
    label = label or str(path)
    try:
        if not path.is_file():
            fail(label, "file is missing")
        if path.stat().st_size > MAX_JSON_BYTES:
            fail(label, f"file exceeds {MAX_JSON_BYTES} bytes")
        raw = path.read_bytes()
        return json.loads(raw.decode("utf-8"), object_pairs_hook=reject_duplicate_pairs, parse_constant=reject_constant)
    except (OSError, UnicodeDecodeError, json.JSONDecodeError, VerificationError) as error:
        fail(label, f"invalid JSON: {error}")
    raise AssertionError("unreachable")


def canonical(value: Any) -> bytes:
    try:
        return json.dumps(value, sort_keys=True, separators=(",", ":"), ensure_ascii=False, allow_nan=False).encode("utf-8")
    except (TypeError, ValueError, OverflowError) as error:
        raise VerificationError(f"cannot canonicalize JSON: {error}") from error


def sha256_json(value: Any) -> str:
    return hashlib.sha256(canonical(value)).hexdigest()


def protocol_path() -> Path:
    root = Path(__file__).resolve().parent
    frozen = root / "protocol.json"
    return frozen if frozen.is_file() else root / "protocol-draft.json"


def protocol() -> dict[str, Any]:
    value = obj(load_json(protocol_path(), "0435 protocol"), "0435 protocol")
    if u64(value.get("change"), "protocol.change") != CHANGE:
        fail("protocol.change", f"must be {CHANGE}")
    if value.get("samples") != FORMAL_SAMPLES or value.get("warmups") != FORMAL_WARMUPS or value.get("workers") != WORKERS:
        fail("protocol", "must retain samples=30, warmups=3, workers=1")
    if value.get("cpu") != CPU or value.get("repeats") != 2:
        fail("protocol", "must bind CPU 2 and two repeats")
    if value.get("modes") != list(MODES) or value.get("shapes") != SHAPES:
        fail("protocol", "mode or shape matrix differs from the ODT contract")
    roles = obj(value.get("roles"), "protocol.roles")
    if set(roles) != set(ROLES):
        fail("protocol.roles", "must contain exactly the three declared ODT roles")
    for role in ROLES:
        spec = obj(roles[role], f"protocol.roles.{role}")
        text(spec.get("selector"), f"protocol.roles.{role}.selector")
        text(spec.get("source_field"), f"protocol.roles.{role}.source_field")
        if spec["source_field"] != "odt_paragraphs":
            fail(f"protocol.roles.{role}.source_field", "must be odt_paragraphs")
        text(spec.get("source_role"), f"protocol.roles.{role}.source_role")
        text(spec.get("implementation"), f"protocol.roles.{role}.implementation")
        text(spec.get("retention"), f"protocol.roles.{role}.retention")
    corpus = obj(value.get("corpus"), "protocol.corpus")
    if corpus.get("package_format") != "ODT/ODF/ZIP" or corpus.get("compression") != ODT_COMPRESSION or corpus.get("payload_kind") != ODT_PAYLOAD_KIND or corpus.get("mimetype") != ODT_MIMETYPE:
        fail("protocol.corpus", "does not identify ODT")
    if corpus.get("default_styles_xml_sha256") != ODT_DEFAULT_STYLES_SHA256 or corpus.get("default_meta_xml_sha256") != ODT_DEFAULT_META_SHA256:
        fail("protocol.corpus", "does not pin the immutable ODT style/meta defaults")
    if corpus.get("member_names") != list(ODT_MEMBERS) or corpus.get("manifest_entry_count") != 4:
        fail("protocol.corpus", "member topology differs from the ODT contract")
    if corpus.get("target_entry") != "content.xml":
        fail("protocol.corpus.target_entry", "must be content.xml")
    fixture = obj(corpus.get("paragraph_fixture"), "protocol.corpus.paragraph_fixture")
    if fixture.get("prefix") != "litchi-perf-odt-buffered-" or fixture.get("variants") != list(TEXT_VARIANTS) or fixture.get("cycle") != 4:
        fail("protocol.corpus.paragraph_fixture", "does not pin the ODT paragraph fixture")
    if fixture.get("text_contract") != TEXT_CONTRACT:
        fail("protocol.corpus.paragraph_fixture.text_contract", "does not pin the text contract")
    return value


def role_spec(p: dict[str, Any], role: str) -> dict[str, Any]:
    if role not in ROLES:
        fail("role", f"must be one of {ROLES}")
    return obj(p["roles"][role], f"protocol.roles.{role}")


def paragraph_text(index: int) -> str:
    return f"litchi-perf-odt-buffered-{index:05} {TEXT_VARIANTS[index % 4]}"


def expected_paragraphs(shape: str) -> list[str]:
    if shape not in SHAPES:
        fail("shape", f"must be one of {sorted(SHAPES)}")
    return [paragraph_text(index) for index in range(SHAPES[shape])]


def semantic_digest(paragraphs: list[str]) -> str:
    hasher = hashlib.sha256(SEMANTIC_DOMAIN)
    hasher.update(len(paragraphs).to_bytes(8, "little"))
    for paragraph in paragraphs:
        encoded = paragraph.encode("utf-8")
        hasher.update(len(encoded).to_bytes(8, "little"))
        hasher.update(encoded)
    return hasher.hexdigest()


def semantic_input_bytes(paragraphs: list[str]) -> int:
    return sum(len(value.encode("utf-8")) for value in paragraphs) + max(0, len(paragraphs) - 1)


def check_tool(report: dict[str, Any], mode: str) -> None:
    tool = obj(report.get("tool"), "report.tool")
    required = ("name", "version", "binary", "profile", "target_os", "target_arch", "instrumentation")
    optional = ("allocator_counter_revision",)
    exact_keys(tool, required, optional, "report.tool")
    expected_binary = "litchi-perf-baseline-alloc" if mode == "allocator" else "litchi-perf-baseline"
    expected_instrumentation = "system_allocator_operation_scoped" if mode == "allocator" else "none"
    for key, expected in (("name", "litchi-perf-baseline"), ("version", "0.1.0"), ("binary", expected_binary), ("profile", "release"), ("target_os", "linux"), ("target_arch", "x86_64"), ("instrumentation", expected_instrumentation)):
        if tool.get(key) != expected:
            fail(f"report.tool.{key}", f"must be {expected!r}")
    if mode == "allocator" and tool.get("allocator_counter_revision") != ALLOCATOR_REVISION:
        fail("report.tool.allocator_counter_revision", f"must be {ALLOCATOR_REVISION!r}")
    if mode == "normal" and "allocator_counter_revision" in tool:
        fail("report.tool.allocator_counter_revision", "normal reports must not claim allocator instrumentation")


def check_binary(report: dict[str, Any]) -> None:
    identity = obj(report.get("binary_identity"), "report.binary_identity")
    exact_keys(identity, ("path", "binary_sha256", "binary_bytes", "mode_bits", "executable", "profile"), (), "report.binary_identity")
    if not Path(text(identity["path"], "report.binary_identity.path")).is_absolute():
        fail("report.binary_identity.path", "must be absolute")
    digest(identity["binary_sha256"], "report.binary_identity.binary_sha256")
    u64(identity["binary_bytes"], "report.binary_identity.binary_bytes", positive=True)
    if identity["profile"] != "release" or identity["executable"] is not True:
        fail("report.binary_identity", "must identify an executable release binary")
    bits = identity["mode_bits"]
    if isinstance(bits, bool) or not isinstance(bits, int) or not 0 <= bits <= 0o7777 or bits & 0o111 == 0:
        fail("report.binary_identity.mode_bits", "must contain executable Unix permission bits")


def normalize_rustflags(value: str) -> list[str]:
    try:
        tokens = shlex.split(value)
    except ValueError as error:
        raise VerificationError(f"invalid RUSTFLAGS: {error}") from error
    result: list[str] = []
    index = 0
    while index < len(tokens):
        if tokens[index] == "-C":
            if index + 1 >= len(tokens):
                fail("report.environment.rustflags", "-C has no option")
            result.append("-C" + tokens[index + 1])
            index += 2
        else:
            result.append(tokens[index])
            index += 1
    return result


def check_environment(report: dict[str, Any], mode: str) -> None:
    environment = obj(report.get("environment"), "report.environment")
    fields = ("rustc_version", "git_revision", "git_worktree_dirty", "logical_cpus_available", "allocator", "rustflags", "cargo_build_target", "perf_event_paranoid", "os", "kernel", "cpu_model", "total_memory_bytes", "page_size_bytes", "filesystem_type", "source_destination_same_device", "cpu_affinity", "storage_identifier")
    exact_keys(environment, fields, (), "report.environment")
    text(environment["rustc_version"], "report.environment.rustc_version")
    if HEX40.fullmatch(text(environment["git_revision"], "report.environment.git_revision")) is None:
        fail("report.environment.git_revision", "must be a 40-character revision")
    boolean(environment["git_worktree_dirty"], "report.environment.git_worktree_dirty")
    u64(environment["logical_cpus_available"], "report.environment.logical_cpus_available", positive=True)
    expected_allocator = "CountingSystemAllocator(std::alloc::System)" if mode == "allocator" else "Rust system allocator"
    if environment["allocator"] != expected_allocator:
        fail("report.environment.allocator", f"must be {expected_allocator!r}")
    flags = environment["rustflags"]
    if flags is None:
        fail("report.environment.rustflags", "must bind release frame-pointer flags")
    tokens = normalize_rustflags(text(flags, "report.environment.rustflags", allow_empty=True))
    if "-Cforce-frame-pointers=yes" not in tokens:
        fail("report.environment.rustflags", "must include -C force-frame-pointers=yes")
    for key in ("cargo_build_target", "perf_event_paranoid", "os", "kernel", "cpu_model", "filesystem_type", "cpu_affinity", "storage_identifier"):
        if environment[key] is not None:
            text(environment[key], f"report.environment.{key}", allow_empty=True)
    if environment["os"] != "linux" or environment["cpu_affinity"] != str(CPU):
        fail("report.environment", "must identify Linux CPU2 capture")
    for key in ("total_memory_bytes", "page_size_bytes"):
        if environment[key] is not None:
            u64(environment[key], f"report.environment.{key}", positive=True)
    if environment["source_destination_same_device"] is not None:
        boolean(environment["source_destination_same_device"], "report.environment.source_destination_same_device")


def check_configuration(report: dict[str, Any], shape: str, selector: str) -> None:
    config = obj(report.get("configuration"), "report.configuration")
    fields = ("samples_per_case", "warmup_iterations_per_case", "filesystem_cache_states", "filesystem_fresh_child_per_sample", "filesystem_process_isolated", "filesystem_root_selected", "cases", "corpus_shapes", "payload_kinds", "writer_shapes", "xlsx_shapes", "xlsb_shapes", "xlsx_cell_crud_shapes", "xlsx_row_visibility_shapes", "semantic_shapes", "rtf_variants", "range_simulation", "execution_workers", "opc_cache_lock_diagnostics")
    exact_keys(config, fields, (), "report.configuration")
    if config["samples_per_case"] != SAMPLES or config["warmup_iterations_per_case"] != WARMUPS:
        fail("report.configuration", "sample and warmup counts do not match protocol")
    expected = {"filesystem_cache_states": ["warm", "cold-requested"], "cases": [selector], "corpus_shapes": ["tiny", "many-small", "few-large", "wide-root"], "payload_kinds": ["compressible", "incompressible"], "writer_shapes": ["tiny", "large", "payload-heavy"], "xlsx_shapes": ["tiny", "medium", "dense-wide"], "xlsb_shapes": ["tiny", "medium", "large", "sparse"], "xlsx_cell_crud_shapes": ["medium", "dense-sparse"], "xlsx_row_visibility_shapes": ["medium", "large"], "semantic_shapes": [shape], "rtf_variants": ["plain"], "execution_workers": [WORKERS]}
    for key, value in expected.items():
        if config[key] != value:
            fail(f"report.configuration.{key}", f"must be exactly {value!r}")
    if config["filesystem_fresh_child_per_sample"] is not True or config["filesystem_process_isolated"] is not True or config["filesystem_root_selected"] is not False or config["opc_cache_lock_diagnostics"] is not False:
        fail("report.configuration", "fresh-process or lock-diagnostic flags differ from the protocol")
    simulation = obj(config["range_simulation"], "report.configuration.range_simulation")
    exact_keys(simulation, ("fixed_latency_us", "request_overhead_us", "bandwidth_bytes_per_second", "max_physical_range_bytes"), (), "report.configuration.range_simulation")
    if simulation != {"fixed_latency_us": 100, "request_overhead_us": 25, "bandwidth_bytes_per_second": 50 * 1024 * 1024, "max_physical_range_bytes": 4 * 1024}:
        fail("report.configuration.range_simulation", "does not match harness defaults")


def check_parallel(report: dict[str, Any], archive_hash: str, selector: str) -> None:
    metrics = obj(report.get("parallel_metrics"), "report.parallel_metrics")
    exact_keys(metrics, ("schema_version", "scope", "claim", "configured_worker_budget", "observed_process_thread_count", "cases"), (), "report.parallel_metrics")
    if metrics["schema_version"] != 1 or metrics["scope"] != "explicit_local_execution_only" or metrics["claim"] != "descriptive":
        fail("report.parallel_metrics", "worker evidence is not explicitly local/descriptive")
    budget = obj(metrics["configured_worker_budget"], "report.parallel_metrics.configured_worker_budget")
    if budget.get("status") != "measured" or budget.get("value") != [WORKERS] or budget.get("scope") != "configuration.execution_workers":
        fail("report.parallel_metrics.configured_worker_budget", "does not bind one worker")
    cases = array(metrics["cases"], "report.parallel_metrics.cases")
    if len(cases) != 1:
        fail("report.parallel_metrics.cases", "must contain the selected case")
    case = obj(cases[0], "report.parallel_metrics.cases[0]")
    exact_keys(case, ("case", "corpus_sha256", "configured_worker_count", "observed_local_worker_count", "deterministic_task_count", "deterministic_chunk_count", "lock_wait_ns"), (), "report.parallel_metrics.cases[0]")
    if case["case"] != selector or digest(case["corpus_sha256"], "report.parallel_metrics.cases[0].corpus_sha256") != archive_hash:
        fail("report.parallel_metrics.cases[0]", "case or corpus identity differs")
    for key in ("configured_worker_count", "observed_local_worker_count", "deterministic_task_count", "deterministic_chunk_count"):
        metric = obj(case[key], f"report.parallel_metrics.cases[0].{key}")
        if metric.get("status") not in {"measured", "not_applicable", "unavailable"}:
            fail(f"report.parallel_metrics.cases[0].{key}", "unknown metric status")
    lock = obj(case["lock_wait_ns"], "report.parallel_metrics.cases[0].lock_wait_ns")
    if lock.get("status") not in {"unavailable", "not_applicable"}:
        fail("report.parallel_metrics.cases[0].lock_wait_ns", "must not imply a lock contention measurement")


def check_corpus(value: Any, shape: str, role: str, path: str = "report.results[0].corpus") -> dict[str, Any]:
    corpus = obj(value, path)
    required = ("name", "generator", "package_format", "shape", "payload_kind", "compression", "entry_count", "archive_member_count", "entry_bytes", "uncompressed_payload_bytes", "archive_bytes", "archive_sha256", "target_entry", "target_payload_bytes", "target_payload_sha256", "xlsx")
    exact_keys(corpus, required, ("rtf_variant",), path)
    paragraphs = expected_paragraphs(shape)
    for key in ("name", "generator", "package_format", "shape", "payload_kind", "compression", "target_entry"):
        text(corpus[key], f"{path}.{key}")
    if corpus["shape"] != shape or corpus["package_format"] != "ODT/ODF/ZIP" or corpus["compression"] != ODT_COMPRESSION or corpus["payload_kind"] != ODT_PAYLOAD_KIND:
        fail(path, "does not identify the ODT corpus")
    source_role = "streaming" if role == "after-streaming" else "buffered"
    if corpus["name"] != f"odt-{source_role}-paragraphs-{shape}" or corpus["generator"] != f"litchi-odt-{source_role}-paragraphs-v1":
        fail(path, "does not identify the exact role-local ODT producer corpus")
    if corpus["target_entry"] != "content.xml" or corpus["archive_member_count"] != len(ODT_MEMBERS) or corpus["entry_count"] != len(paragraphs):
        fail(path, "member or paragraph counts differ from the ODT contract")
    expected_entry_bytes = len(paragraphs[0].encode("utf-8"))
    if corpus["entry_bytes"] != expected_entry_bytes:
        fail(f"{path}.entry_bytes", "does not match the first deterministic paragraph")
    for key in ("entry_count", "archive_member_count", "entry_bytes", "uncompressed_payload_bytes", "archive_bytes", "target_payload_bytes"):
        u64(corpus[key], f"{path}.{key}")
    if corpus["uncompressed_payload_bytes"] != semantic_input_bytes(paragraphs):
        fail(f"{path}.uncompressed_payload_bytes", "does not match the paragraph projection")
    if corpus["archive_bytes"] == 0 or corpus["target_payload_bytes"] == 0:
        fail(path, "archive and target payload must be non-empty")
    archive_hash = digest(corpus["archive_sha256"], f"{path}.archive_sha256")
    target_hash = digest(corpus["target_payload_sha256"], f"{path}.target_payload_sha256")
    if corpus.get("rtf_variant") is not None or corpus["xlsx"] is not None:
        fail(path, "ODT corpus must not carry RTF/XLSX metadata")
    return {"name": corpus["name"], "archive_hash": archive_hash, "archive_bytes": corpus["archive_bytes"], "target_payload_bytes": corpus["target_payload_bytes"], "target_payload_hash": target_hash, "input_bytes": corpus["uncompressed_payload_bytes"], "member_names": list(ODT_MEMBERS), "mimetype": ODT_MIMETYPE, "manifest_entry_count": 4, "paragraph_count": len(paragraphs)}


def check_source(value: Any, role: str, shape: str, p: dict[str, Any], path: str = "report.results[0].source") -> dict[str, Any]:
    source = obj(value, path)
    exact_keys(source, ("read_calls", "read_bytes", "ordinary_payload_read_calls", "ordinary_payload_read_bytes", "max_in_flight_reads", "odt_paragraphs"), (), path)
    for key in ("read_calls", "read_bytes", "ordinary_payload_read_calls", "ordinary_payload_read_bytes", "max_in_flight_reads"):
        if array(source[key], f"{path}.{key}"):
            fail(f"{path}.{key}", "fresh ODT authoring must not claim source reads")
    summary = obj(source["odt_paragraphs"], f"{path}.odt_paragraphs")
    required = tuple(obj(p["source_contract"], "protocol.source_contract")["required_fields"])
    exact_keys(summary, required, (), f"{path}.odt_paragraphs")
    spec = role_spec(p, role)
    expected_role = spec["source_role"]
    if summary["role"] != expected_role or summary["implementation"] != spec["implementation"]:
        fail(f"{path}.odt_paragraphs", "role or implementation differs from the protocol")
    timing = text(summary["timing_scope"], f"{path}.odt_paragraphs.timing_scope")
    claim = text(summary["performance_claim"], f"{path}.odt_paragraphs.performance_claim")
    if "corpus" not in timing.lower() or "outside" not in timing.lower() or "no " not in claim.lower():
        fail(f"{path}.odt_paragraphs", "timing/claim text does not disclose setup and descriptive scope")
    paragraphs = expected_paragraphs(shape)
    if digest(summary["semantic_sha256"], f"{path}.odt_paragraphs.semantic_sha256") != semantic_digest(paragraphs):
        fail(f"{path}.odt_paragraphs.semantic_sha256", "does not match the independent paragraph digest")
    for key in ("content_xml_sha256", "styles_xml_sha256", "meta_xml_sha256"):
        digest(summary[key], f"{path}.odt_paragraphs.{key}")
    if summary["styles_xml_sha256"].lower() != ODT_DEFAULT_STYLES_SHA256 or summary["meta_xml_sha256"].lower() != ODT_DEFAULT_META_SHA256:
        fail(f"{path}.odt_paragraphs", "styles.xml/meta.xml do not match the pinned Builder defaults")
    for key in ("archive_member_set_verified", "manifest_bindings_verified", "semantic_reopen_verified", "immutable_styles_meta_verified"):
        if summary[key] is not True:
            fail(f"{path}.odt_paragraphs.{key}", "semantic/package gate must be true")
    if u64(summary["paragraph_count"], f"{path}.odt_paragraphs.paragraph_count") != len(paragraphs) or u64(summary["run_count"], f"{path}.odt_paragraphs.run_count") != len(paragraphs) or summary["text_contract"] != TEXT_CONTRACT:
        fail(f"{path}.odt_paragraphs", "paragraph count or text contract differs")
    return {"role": summary["role"], "implementation": summary["implementation"], "semantic_sha256": summary["semantic_sha256"].lower(), "content_xml_sha256": summary["content_xml_sha256"].lower(), "styles_xml_sha256": summary["styles_xml_sha256"].lower(), "meta_xml_sha256": summary["meta_xml_sha256"].lower(), "member_names": list(ODT_MEMBERS), "mimetype": ODT_MIMETYPE, "manifest_entry_count": 4, "paragraph_count": len(paragraphs)}


def median_int(values: list[int]) -> int:
    low = values[(len(values) - 1) // 2]
    high = values[len(values) // 2]
    return (low + high) // 2


def nearest_rank(values: list[int], percentile: int) -> int:
    index = min(((percentile * len(values) + 99) // 100) - 1, len(values) - 1)
    return values[index]


def check_elapsed(value: Any, path: str = "report.results[0].elapsed_ns") -> tuple[list[int], list[int]]:
    elapsed = obj(value, path)
    exact_keys(elapsed, ("unit", "samples", "sample_order", "min", "p50", "p95", "p99", "max", "mean", "standard_deviation", "confidence_interval_95"), (), path)
    if elapsed["unit"] != "ns":
        fail(f"{path}.unit", "must be ns")
    raw = array(elapsed["samples"], f"{path}.samples")
    if len(raw) != SAMPLES:
        fail(f"{path}.samples", f"must contain {SAMPLES} samples")
    samples = [u64(item, f"{path}.samples[{i}]", positive=True) for i, item in enumerate(raw)]
    if samples != sorted(samples):
        fail(f"{path}.samples", "must be elapsed-sorted")
    order = [u64(item, f"{path}.sample_order[{i}]") for i, item in enumerate(array(elapsed["sample_order"], f"{path}.sample_order"))]
    if len(order) != SAMPLES or sorted(order) != list(range(SAMPLES)):
        fail(f"{path}.sample_order", "must be a complete original-sample permutation")
    for i in range(1, SAMPLES):
        if samples[i] == samples[i - 1] and order[i] <= order[i - 1]:
            fail(f"{path}.sample_order", "must increase across tied elapsed samples")
    expected_int = {"min": samples[0], "p50": median_int(samples), "p95": nearest_rank(samples, 95), "p99": nearest_rank(samples, 99), "max": samples[-1]}
    for key, expected in expected_int.items():
        if u64(elapsed[key], f"{path}.{key}") != expected:
            fail(f"{path}.{key}", "does not match retained samples")
    mean = sum(samples) / SAMPLES
    sd = math.sqrt(sum((sample - mean) ** 2 for sample in samples) / (SAMPLES - 1))
    if not math.isclose(finite(elapsed["mean"], f"{path}.mean"), mean, rel_tol=1e-12, abs_tol=1e-12) or not math.isclose(finite(elapsed["standard_deviation"], f"{path}.standard_deviation"), sd, rel_tol=1e-12, abs_tol=1e-12):
        fail(path, "mean or standard deviation does not match retained samples")
    interval = obj(elapsed["confidence_interval_95"], f"{path}.confidence_interval_95")
    exact_keys(interval, ("method", "lower", "upper"), (), f"{path}.confidence_interval_95")
    if interval["method"] != "two-sided Student's t interval for the mean":
        fail(f"{path}.confidence_interval_95.method", "does not identify the frozen interval")
    if SAMPLES - 1 not in T_CRITICAL_95:
        fail("report.elapsed_ns", "sample count has no pinned Student's t critical value")
    margin = T_CRITICAL_95[SAMPLES - 1] * sd / math.sqrt(SAMPLES)
    if not math.isclose(finite(interval["lower"], f"{path}.confidence_interval_95.lower"), max(mean - margin, 0), rel_tol=1e-12, abs_tol=1e-12) or not math.isclose(finite(interval["upper"], f"{path}.confidence_interval_95.upper"), mean + margin, rel_tol=1e-12, abs_tol=1e-12):
        fail(f"{path}.confidence_interval_95", "does not match the Student's t calculation")
    return samples, order


def check_metric_vector(value: Any, path: str, *, status: str | None = None, count: int | None = None, scope: str | None = None, values_required: bool | None = None) -> list[int] | None:
    if count is None:
        count = SAMPLES
    vector = obj(value, path)
    exact_keys(vector, ("status", "scope"), ("values",), path)
    actual = text(vector["status"], f"{path}.status")
    if actual not in {"measured", "not_applicable", "unavailable", "overflow"}:
        fail(f"{path}.status", "unknown metric status")
    if status is not None and actual != status:
        fail(f"{path}.status", f"must be {status!r}")
    if scope is not None and vector["scope"] != scope:
        fail(f"{path}.scope", f"must be {scope!r}")
    present = "values" in vector
    if values_required is None:
        values_required = actual == "measured"
    if present != values_required:
        fail(path, "values presence does not match metric status")
    if not present:
        return None
    values = [u64(item, f"{path}.values[{i}]") for i, item in enumerate(array(vector["values"], f"{path}.values"))]
    if len(values) != count:
        fail(f"{path}.values", f"must contain {count} values")
    return values


def check_allocation(value: Any, order: list[int], path: str) -> dict[str, Any]:
    allocation = obj(value, path)
    exact_keys(allocation, ("status", "scope", *ALLOCATOR_FIELDS), (), path)
    if allocation["status"] != "measured" or allocation["scope"] != ALLOCATOR_SCOPE:
        fail(path, "allocator metrics must be measured with the serialized v3 scope")
    vectors = {key: check_metric_vector(allocation[key], f"{path}.{key}") for key in ALLOCATOR_FIELDS}
    assert all(value is not None for value in vectors.values())
    values = {key: value for key, value in vectors.items() if value is not None}
    original = sorted(range(SAMPLES), key=lambda index: order[index])
    previous_after: int | None = None
    for index in original:
        before = values["peak_live_bytes_before"][index]
        after = values["peak_live_bytes_after"][index]
        if previous_after is not None and (before < previous_after or after < previous_after):
            fail(path, "absolute allocator high-water counters decrease in original order")
        if after < before:
            fail(path, "peak_live_bytes_after is below peak_live_bytes_before")
        previous_after = after
    for index in range(SAMPLES):
        before = values["live_bytes_before"][index]
        after = values["live_bytes_after"][index]
        if values["failed_allocation_calls"][index] != 0:
            fail(path, "a retained sample reports a failed allocation")
        if after != before + values["allocated_bytes"][index] - values["deallocated_bytes"][index]:
            fail(f"{path}[{index}]", "allocated/deallocated bytes do not balance live bytes")
        if values["allocation_calls"][index] < values["reallocation_calls"][index] or values["region_peak_live_bytes"][index] < max(before, after) or values["region_peak_live_bytes"][index] > values["peak_live_bytes_after"][index]:
            fail(f"{path}[{index}]", "allocator call/region peak invariant failed")
    return allocation


def check_operation(value: Any, mode: str, order: list[int], sink: dict[str, Any], path: str = "report.results[0].operation_metrics") -> dict[str, Any]:
    operation = obj(value, path)
    required = ("sample_count", "sample_indices", "alignment", "latency_claim", "source", "process", "sink", "publication", "materialization", "cfb_phases")
    exact_keys(operation, required, ("allocation", "opc_zip"), path)
    if operation["sample_count"] != SAMPLES or operation["sample_indices"] != order or operation["alignment"] != ALIGNMENT or operation["latency_claim"] != LATENCY_CLAIM:
        fail(path, "sample alignment or comparable-operation claim differs from protocol")
    source = obj(operation["source"], f"{path}.source")
    source_required = ("status", "counter_scope", "logical_read_calls", "logical_read_requested_bytes", "logical_read_returned_bytes", "logical_read_largest_requested_bytes", "logical_read_largest_returned_bytes", "logical_read_pattern", "compressed_bytes", "decompressed_bytes", "recompressed_bytes", "max_concurrent_reads")
    exact_keys(source, source_required, (), f"{path}.source")
    if source["status"] != "not_applicable" or source["counter_scope"] != "not_applicable_in_process_sink":
        fail(f"{path}.source", "fresh ODT authoring must identify source metrics as not applicable")
    for key in source_required[2:]:
        vector = source[key]
        if key == "logical_read_pattern":
            pattern = obj(vector, f"{path}.source.{key}")
            exact_keys(pattern, ("status", "scope"), (), f"{path}.source.{key}")
            if pattern["status"] != "not_applicable":
                fail(f"{path}.source.{key}", "must be not_applicable")
        else:
            check_metric_vector(vector, f"{path}.source.{key}", status="not_applicable", values_required=False)
    process = obj(operation["process"], f"{path}.process")
    exact_keys(process, ("status", *PROCESS_FIELDS), (), f"{path}.process")
    process_status = text(process["status"], f"{path}.process.status")
    if process_status not in {"measured", "unavailable"}:
        fail(f"{path}.process.status", "must be measured or unavailable")
    for key in PROCESS_FIELDS:
        check_metric_vector(process[key], f"{path}.process.{key}", status=process_status)
    sink_metrics = obj(operation["sink"], f"{path}.sink")
    exact_keys(sink_metrics, ("status", "output_bytes", "write_status", *SINK_VECTOR_FIELDS, "write_size_buckets"), (), f"{path}.sink")
    if sink_metrics["status"] != "not_applicable" or sink_metrics["write_status"] != "measured":
        fail(f"{path}.sink", "ODT discard sink status is inconsistent")
    check_metric_vector(sink_metrics["output_bytes"], f"{path}.sink.output_bytes", status="not_applicable", values_required=False)
    for key in SINK_VECTOR_FIELDS:
        values = check_metric_vector(sink_metrics[key], f"{path}.sink.{key}")
        if values is None or any(value != sink[key] for value in values):
            fail(f"{path}.sink.{key}", "does not repeat the deterministic top-level sink summary")
    buckets = obj(sink_metrics["write_size_buckets"], f"{path}.sink.write_size_buckets")
    exact_keys(buckets, ("status", *SINK_BUCKET_FIELDS), (), f"{path}.sink.write_size_buckets")
    if buckets["status"] != "measured":
        fail(f"{path}.sink.write_size_buckets.status", "must be measured")
    for key in SINK_BUCKET_FIELDS:
        values = check_metric_vector(buckets[key], f"{path}.sink.write_size_buckets.{key}")
        if values is None or any(value != sink["buckets"][key] for value in values):
            fail(f"{path}.sink.write_size_buckets.{key}", "does not repeat the top-level sink bucket")
    for name in ("publication", "materialization"):
        section = obj(operation[name], f"{path}.{name}")
        if section.get("status") != "not_applicable":
            fail(f"{path}.{name}.status", "must be not_applicable for fresh ODT authoring")
        for key, metric in section.items():
            if key != "status":
                if isinstance(metric, dict):
                    check_metric_vector(metric, f"{path}.{name}.{key}", status="not_applicable", values_required=False)
    phases = obj(operation["cfb_phases"], f"{path}.cfb_phases")
    if phases.get("status") != "not_applicable":
        fail(f"{path}.cfb_phases.status", "must be not_applicable for ODT")
    if mode == "allocator":
        allocation = check_allocation(operation.get("allocation"), order, f"{path}.allocation")
    else:
        allocation = operation.get("allocation")
        if allocation is not None:
            allocation = obj(allocation, f"{path}.allocation")
            exact_keys(allocation, ("status", "scope", *ALLOCATOR_FIELDS), (), f"{path}.allocation")
            if allocation["status"] not in {"unavailable", "not_applicable"}:
                fail(f"{path}.allocation.status", "normal report cannot claim measured allocation vectors")
            for key in ALLOCATOR_FIELDS:
                check_metric_vector(allocation[key], f"{path}.allocation.{key}", status=allocation["status"], values_required=False)
    return {"process": process, "allocation": allocation}


def check_sink(value: Any, corpus: dict[str, Any], role: str, p: dict[str, Any], path: str = "report.results[0].sink") -> dict[str, Any]:
    sink = obj(value, path)
    required = ("accepted_bytes", "write_calls", "largest_write", "write_size_buckets", "retained_output_bytes", "paragraphs", "runs", "input_bytes", "authored_part_bytes")
    exact_keys(sink, required, ("retained_authoring_window_bytes",), path)
    for key in required:
        if key != "write_size_buckets":
            u64(sink[key], f"{path}.{key}")
    if sink["accepted_bytes"] != corpus["archive_bytes"] or sink["paragraphs"] != corpus["paragraph_count"] or sink["runs"] != corpus["paragraph_count"] or sink["input_bytes"] != corpus["input_bytes"] or sink["authored_part_bytes"] != corpus["target_payload_bytes"]:
        fail(path, "sink summary does not match the deterministic corpus")
    if sink["retained_output_bytes"] != 0 or sink["write_calls"] == 0 or sink["largest_write"] > sink["accepted_bytes"]:
        fail(path, "sink retention/write counters are impossible")
    spec = role_spec(p, role)
    if spec["retention"] == "fixed_explicit_window":
        if sink.get("retained_authoring_window_bytes") != 4_096:
            fail(path, "streaming role must publish the fixed 4096-byte authoring window")
    elif "retained_authoring_window_bytes" in sink:
        fail(path, "buffered role must not fabricate a fixed authoring window")
    buckets = obj(sink["write_size_buckets"], f"{path}.write_size_buckets")
    exact_keys(buckets, SINK_BUCKET_FIELDS, (), f"{path}.write_size_buckets")
    counts = {key: u64(buckets[key], f"{path}.write_size_buckets.{key}") for key in SINK_BUCKET_FIELDS}
    if sum(counts.values()) != sink["write_calls"]:
        fail(f"{path}.write_size_buckets", "bucket counts do not sum to write_calls")
    sink["buckets"] = counts
    return sink


def check_catalog(report_path: Path, report: dict[str, Any], corpus: dict[str, Any], selector: str) -> None:
    reference = obj(report["corpus_catalog"], "report.corpus_catalog")
    exact_keys(reference, ("manifest_version", "catalog_id", "catalog_sha256", "content_set_sha256"), (), "report.corpus_catalog")
    if reference["manifest_version"] != 2 or reference["catalog_id"] != "litchi-perf-corpus-v2":
        fail("report.corpus_catalog", "unsupported catalog identity")
    reference_catalog = digest(reference["catalog_sha256"], "report.corpus_catalog.catalog_sha256")
    reference_content = digest(reference["content_set_sha256"], "report.corpus_catalog.content_set_sha256")
    sidecar_path = report_path.with_name(report_path.stem + "-catalog.json")
    catalog = obj(load_json(sidecar_path, "corpus catalog sidecar"), "corpus catalog sidecar")
    exact_keys(catalog, ("manifest_version", "manifest_kind", "catalog_id", "canonicalization", "catalog_sha256", "content_set_sha256", "build", "corpora", "case_bindings"), (), "corpus catalog sidecar")
    if catalog["manifest_version"] != 2 or catalog["manifest_kind"] != "corpus-catalog" or catalog["catalog_id"] != "litchi-perf-corpus-v2":
        fail("corpus catalog sidecar", "unsupported catalog identity")
    if catalog["catalog_sha256"] != reference_catalog or catalog["content_set_sha256"] != reference_content:
        fail("report.corpus_catalog", "report reference differs from catalog sidecar")
    canonicalization = obj(catalog["canonicalization"], "corpus catalog sidecar.canonicalization")
    if canonicalization != {"algorithm": "sorted-json-utf8-compact-v1", "hash": "sha256"}:
        fail("corpus catalog sidecar.canonicalization", "does not identify canonical hashing")
    without_hash = dict(catalog)
    del without_hash["catalog_sha256"]
    if sha256_json(without_hash) != reference_catalog:
        fail("corpus catalog sidecar.catalog_sha256", "canonical catalog hash is stale")
    corpora = array(catalog["corpora"], "corpus catalog sidecar.corpora")
    if len(corpora) != 1:
        fail("corpus catalog sidecar.corpora", "must contain exactly the selected ODT corpus")
    matched = False
    content_rows = []
    for index, value in enumerate(corpora):
        row = obj(value, f"corpus catalog sidecar.corpora[{index}]")
        bytes_row = obj(row.get("bytes"), f"corpus catalog sidecar.corpora[{index}].bytes")
        if digest(bytes_row.get("archive_sha256"), f"corpus catalog sidecar.corpora[{index}].bytes.archive_sha256") == corpus["archive_hash"]:
            matched = True
        expected_id = f"odt-odf-zip:sha256:{corpus['archive_hash']}"
        if row.get("id") != expected_id or row.get("legacy_v1") != report["results"][0]["corpus"]:
            fail("corpus catalog sidecar.corpora", "legacy manifest or content id differs from the report")
        if bytes_row.get("archive_bytes") != corpus["archive_bytes"] or bytes_row.get("logical_payload_bytes") != corpus["input_bytes"]:
            fail("corpus catalog sidecar.corpora.bytes", "byte counts differ from the report")
        if row.get("targets") != [{"entry": "content.xml", "logical_bytes": corpus["target_payload_bytes"], "sha256": corpus["target_payload_hash"]}]:
            fail("corpus catalog sidecar.corpora.targets", "target identity differs from content.xml")
        members = obj(row.get("members"), f"corpus catalog sidecar.corpora[{index}].members")
        items = array(members.get("items"), f"corpus catalog sidecar.corpora[{index}].members.items")
        names = [text(obj(item, "catalog member").get("name"), "catalog member.name") for item in items]
        if len(set(names)) != len(names):
            fail(f"corpus catalog sidecar.corpora[{index}].members", "contains duplicate member names")
        if members.get("status") != "unavailable" or items:
            fail("corpus catalog sidecar.corpora.members", "unexpected member catalog; package proof belongs to the source summary gates")
        content_rows.append({"id": row["id"], "archive_sha256": bytes_row["archive_sha256"], "members": []})
    if not matched:
        fail("corpus catalog sidecar.corpora", "does not contain the selected archive identity")
    bindings = array(catalog["case_bindings"], "corpus catalog sidecar.case_bindings")
    expected_binding = {"case": selector, "corpus_id": expected_id, "legacy_name": corpus["name"], "legacy_archive_sha256": corpus["archive_hash"], "role": "timed"}
    if bindings != [expected_binding]:
        fail("corpus catalog sidecar.case_bindings", "does not exactly bind the selected ODT case")
    content_identity = {"corpora": content_rows, "case_bindings": [{"case": selector, "corpus_id": expected_id, "role": "timed"}]}
    if sha256_json(content_identity) != reference_content:
        fail("corpus catalog sidecar.content_set_sha256", "canonical content identity hash is stale")


def validate_report(report_path: Path, mode: str, shape: str, role: str, *, samples: int = FORMAL_SAMPLES, warmups: int = FORMAL_WARMUPS) -> dict[str, Any]:
    global SAMPLES, WARMUPS
    if mode not in MODES:
        fail("mode", "must be normal or allocator")
    if isinstance(samples, bool) or not isinstance(samples, int) or not 1 <= samples <= FORMAL_SAMPLES:
        fail("samples", f"must be an integer between 1 and {FORMAL_SAMPLES}")
    if isinstance(warmups, bool) or not isinstance(warmups, int) or not 0 <= warmups <= FORMAL_WARMUPS:
        fail("warmups", f"must be an integer between 0 and {FORMAL_WARMUPS}")
    SAMPLES, WARMUPS = samples, warmups
    p = protocol()
    spec = role_spec(p, role)
    report = obj(load_json(report_path, "report"), "report")
    exact_keys(report, ("schema_version", "tool", "binary_identity", "environment", "configuration", "parallel_metrics", "results", "corpus_catalog"), (), "report")
    if report["schema_version"] != SCHEMA_VERSION:
        fail("report.schema_version", "must be one")
    check_tool(report, mode)
    check_binary(report)
    check_environment(report, mode)
    check_configuration(report, shape, spec["selector"])
    results = array(report["results"], "report.results")
    if len(results) != 1:
        fail("report.results", "must contain exactly one selected case")
    result = obj(results[0], "report.results[0]")
    exact_keys(result, ("case", "corpus", "elapsed_ns", "sink", "source", "output_sha256", "operation_metrics"), (), "report.results[0]")
    if result["case"] != spec["selector"]:
        fail("report.results[0].case", "does not match the protocol role selector")
    corpus = check_corpus(result["corpus"], shape, role)
    source = check_source(result["source"], role, shape, p)
    if source["content_xml_sha256"] != corpus["target_payload_hash"]:
        fail("report.results[0].source.odt_paragraphs.content_xml_sha256", "does not match the corpus content.xml target hash")
    elapsed, order = check_elapsed(result["elapsed_ns"])
    sink = check_sink(result["sink"], corpus, role, p)
    if digest(result["output_sha256"], "report.results[0].output_sha256") != corpus["archive_hash"]:
        fail("report.results[0].output_sha256", "must equal the corpus archive identity")
    metrics = check_operation(result["operation_metrics"], mode, order, sink)
    check_parallel(report, corpus["archive_hash"], spec["selector"])
    check_catalog(report_path, report, corpus, spec["selector"])
    identity = {
        "role": role, "selector": spec["selector"], "mode": mode, "shape": shape,
        "paragraph_count": corpus["paragraph_count"], "semantic_sha256": source["semantic_sha256"],
        "content_xml_sha256": source["content_xml_sha256"],
        "styles_xml_sha256": source["styles_xml_sha256"], "meta_xml_sha256": source["meta_xml_sha256"],
        "member_names": corpus["member_names"], "mimetype": corpus["mimetype"],
        "manifest_entry_count": corpus["manifest_entry_count"], "target_payload_bytes": corpus["target_payload_bytes"],
        "target_payload_sha256": corpus["target_payload_hash"], "archive_bytes": corpus["archive_bytes"],
        "archive_sha256": corpus["archive_hash"], "sink_accepted_bytes": sink["accepted_bytes"],
    }
    return {"report": report, "result": result, "corpus": corpus, "source": source, "elapsed": elapsed, "sample_order": order, "metrics": metrics, "identity": identity, "protocol": p}


def compare_identities(records: Iterable[dict[str, Any]]) -> dict[str, Any]:
    rows = [record["identity"] if "identity" in record else record for record in records]
    if not rows:
        fail("identities", "at least one identity is required")
    common = ("paragraph_count", "semantic_sha256", "styles_xml_sha256", "meta_xml_sha256", "member_names", "mimetype", "manifest_entry_count")
    for key in common:
        if any(row[key] != rows[0][key] for row in rows[1:]):
            fail(f"identity.{key}", "cross-role identity differs")
    return {"common_identity": {key: rows[0][key] for key in common}, "content_xml_comparable": False, "content_xml_comparison": "role-local-only", "archive_bytes_are_role_local": True}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--report", type=Path, required=True)
    parser.add_argument("--mode", choices=MODES, required=True)
    parser.add_argument("--shape", choices=tuple(SHAPES), required=True)
    parser.add_argument("--role", choices=ROLES, required=True)
    parser.add_argument("--samples", type=int, default=FORMAL_SAMPLES)
    parser.add_argument("--warmups", type=int, default=FORMAL_WARMUPS)
    args = parser.parse_args()
    try:
        validate_report(args.report, args.mode, args.shape, args.role, samples=args.samples, warmups=args.warmups)
    except (OSError, KeyError, TypeError, VerificationError, AssertionError) as error:
        print(f"INVALID: {error}", file=sys.stderr)
        return 1
    print("VALID")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
