"""Offline replay of the four 0815 owner-scoped Callgrind publications.

This reader consumes retained receipts and raw Callgrind text only.  Cargo,
the workload, and Valgrind are deliberately absent from this module.  The
guest ``Ir`` values are attribution diagnostics for the exact probe wrapper;
they are not latency, RSS, native-cycle, phase-fraction, or adoption gates.
"""

from __future__ import annotations

import argparse
import importlib.util
import json
import re
import sys
from pathlib import Path
from typing import Any

import custody as c


P = c.P
ROOT = c.ROOT
OWNER = "namespace_uri_probe::capture_region_0793"
SCANNER = "litchi_pptx::notes::codec::scan_processed_xml"
PROCESS_EVENT = "quick_xml::reader::ns_reader::NsReader<R>::process_event"
READ_EVENT = "quick_xml::reader::Reader<R>::read_event_impl"
RESOLVER_PUSH = "quick_xml::name::NamespaceResolver::push"
PARSER_SHA = "70b0bb3fc665860a9aa586725d72ced75e41de06f95e3b1c5408c70998e926b6"
# The profile binary is the final repaired 0806 public-workflow probe.  The
# Callgrind wrapper changes only collection around that executable; it does
# not create a separate report schema or tool identity.
PROFILE_SCHEMA = "litchi.pptx.public-workflow-probe-0806.v1"
PROFILE_TOOL = "public-pptx-probe-0806"
MARKER = "litchi-perf-0780-static-mce-capabilities"
FIXTURE_PACKET = ROOT / "docs" / "performance" / "results" / "change-0813"
FIXTURE_RELATIVE = "qualification/0-large-capture-before.json"
EXPECTED_ARTIFACTS = {
    "json",
    "log",
    "callgrind",
    "callgrind.1",
}
REPORT_KEYS = frozenset({
    "schema", "tool", "mode", "shape", "slides", "shapes_per_slide",
    "timing_scope", "marker", "source", "fixture", "warmup",
    "samples_requested", "samples", "allocator",
})
SOURCE_KEYS = frozenset({"bytes", "sha256"})
FIXTURE_KEYS = frozenset({
    "injection", "slide_parts", "replaced_text_tags", "namespace_declarations",
    "namespaced_attributes", "namespace_uris", "attribute_names",
})
COMMON_VERIFICATION_KEYS = frozenset({
    "semantic_check", "reopened", "expected_text", "actual_text",
    "semantic_text_bytes", "semantic_text_sha256", "readback_bytes",
    "readback_sha256", "marker_matches", "unknown_namespace_check",
    "unknown_namespace_occurrences",
})


class EvidenceError(ValueError):
    """A missing, stale, malformed, or contradictory retained artifact."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise EvidenceError(message)


def read(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing evidence: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise EvidenceError(f"invalid JSON {path}: {error}") from error


def write_or_check(path: Path, value: Any, check: bool) -> None:
    encoded = json.dumps(value, indent=2, sort_keys=True) + "\n"
    if path.exists():
        require(path.is_file() and not path.is_symlink(), f"output is not regular: {path}")
        require(path.read_text(encoding="utf-8") == encoded,
                f"replayed output differs: {path.name}")
    else:
        require(not check, f"missing expected output: {path.name}")
        path.write_text(encoded, encoding="utf-8")


def sha(path: Path) -> str:
    return c.sha(path)


def relative(path: Path) -> str:
    try:
        return str(path.relative_to(P))
    except ValueError as error:
        raise EvidenceError(f"packet path escapes change-0815: {path}") from error


def packet_artifact(value: Any, label: str) -> Path:
    require(isinstance(value, dict), f"{label}: artifact is malformed")
    raw = value.get("path")
    require(isinstance(raw, str) and raw, f"{label}: artifact path is missing")
    path = Path(raw)
    if not path.is_absolute():
        path = P / path
    path = path.resolve()
    relative(path)
    size = value.get("bytes")
    digest = value.get("sha256")
    require(type(size) is int and size >= 0, f"{label}: artifact byte count is invalid")
    require(isinstance(digest, str) and re.fullmatch(r"[0-9a-f]{64}", digest),
            f"{label}: artifact SHA-256 is invalid")
    require(path.is_file() and not path.is_symlink(), f"{label}: artifact is missing")
    require(path.stat().st_size == size, f"{label}: artifact byte count changed")
    require(sha(path) == digest, f"{label}: artifact SHA-256 changed")
    return path


def external_binary(value: Any, label: str, cleanup: dict[str, Any] | None) -> dict[str, Any]:
    """Verify a build executable live or against the exact cleanup witness."""

    require(isinstance(value, dict), f"{label}: binary identity is malformed")
    raw = value.get("path")
    require(isinstance(raw, str) and raw, f"{label}: binary path is missing")
    path = Path(raw).resolve()
    require(path.is_absolute() and path.parent == c.TARGET,
            f"{label}: binary escapes the owned target")
    size = value.get("bytes")
    digest = value.get("sha256")
    require(type(size) is int and size > 0, f"{label}: binary byte count is invalid")
    require(isinstance(digest, str) and re.fullmatch(r"[0-9a-f]{64}", digest),
            f"{label}: binary SHA-256 is invalid")
    if path.is_file() and not path.is_symlink():
        require(path.stat().st_size == size and sha(path) == digest,
                f"{label}: live binary identity changed")
        return {"path": str(path), "bytes": size, "sha256": digest}
    require(cleanup is not None and cleanup.get("target_removed") is True,
            f"{label}: missing binary has no cleanup witness")
    require(cleanup.get("target") == str(c.TARGET),
            f"{label}: cleanup target changed")
    require(not c.TARGET.exists(), f"{label}: owned target still exists after cleanup")
    removed = cleanup.get("removed_binaries")
    require(isinstance(removed, list), "cleanup removed_binaries is malformed")
    matches = [row for row in removed if isinstance(row, dict)
               and row.get("path") == str(path)]
    require(len(matches) == 1, f"{label}: exact removed binary witness is missing")
    witness = matches[0]
    require(witness.get("bytes") == size and witness.get("sha256") == digest,
            f"{label}: cleanup binary identity differs")
    return {"path": str(path), "bytes": size, "sha256": digest}


def sealed_fixture() -> tuple[dict[str, Any], dict[str, Any], dict[str, Any]]:
    seal_path = FIXTURE_PACKET / "seal.json"
    fixture_path = FIXTURE_PACKET / FIXTURE_RELATIVE
    seal = read(seal_path)
    files = seal.get("files")
    require(isinstance(files, dict), "0813 seal file map is missing")
    require(files.get(str(fixture_path.relative_to(ROOT))) == sha(fixture_path),
            "0813 large capture fixture is not sealed")
    fixture = read(fixture_path)
    samples = fixture.get("samples")
    require(isinstance(samples, list) and len(samples) == 1
            and isinstance(samples[0], dict), "0813 fixture sample is malformed")
    return fixture, samples[0], {
        "packet": "change-0813",
        "path": FIXTURE_RELATIVE,
        "sha256": sha(fixture_path),
        "seal_sha256": sha(seal_path),
    }


def probe_custody() -> dict[str, str]:
    """Bind this packet to the exact six-file probe retained by 0813."""

    sealed_path = ROOT / "docs" / "performance" / "results" / "change-0813" / "probe-src"
    expected = {
        str(path.relative_to(sealed_path)): sha(path)
        for path in sorted(sealed_path.rglob("*")) if path.is_file()
    }
    require(isinstance(expected, dict)
            and set(expected) == {"Cargo.lock", "Cargo.toml", "Cargo.toml.template",
                                  "src/allocation_metrics.rs", "src/counting_allocator.rs",
                                  "src/main.rs"}
            and all(isinstance(value, str) and re.fullmatch(r"[0-9a-f]{64}", value)
                    for value in expected.values()),
            "sealed probe file set changed")
    require(read(ROOT / "docs" / "performance" / "results" / "change-0813"
                 / "build-before" / "probe.json") == expected,
            "0813 probe reference differs from sealed 0813 copy")
    current = {
        str(path.relative_to(P / "probe-src")): sha(path)
        for path in sorted((P / "probe-src").rglob("*")) if path.is_file()
    }
    require(current == expected, "0815 probe is not the exact 0813 six-file copy")
    return dict(expected)


def load_hash_bound_parser() -> Any:
    """Load only the sealed 0784 raw parser; never invoke its old driver."""

    parser_path = ROOT / "docs" / "performance" / "results" / "change-0784" / "profile_analysis.py"
    seal = read(ROOT / "docs" / "performance" / "results" / "change-0784" / "seal.json")
    files = seal.get("files")
    require(isinstance(files, dict) and files.get("profile_analysis.py") == PARSER_SHA,
            "0784 parser is not bound by its seal")
    require(sha(parser_path) == PARSER_SHA, "0784 parser bytes changed")
    spec = importlib.util.spec_from_file_location("sealed_change0784_profile_parser_0815", parser_path)
    require(spec is not None and spec.loader is not None, "cannot load sealed 0784 parser")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    # The parser's path rendering is packet-local.  Its parser implementation
    # and source hash remain 0784; only the output-root used in diagnostics is
    # relocated to this packet.
    module.HERE = P
    module.ROOT = ROOT
    module.OLD_PACKET = ROOT / "docs" / "performance" / "results" / "change-0780"
    return module


def check_plan() -> dict[str, Any]:
    plan = read(P / "plan.json")
    require(plan.get("schema") == "litchi.performance.0815.v1", "plan schema changed")
    require(plan.get("cpu") == 12, "profile CPU changed")
    profile = plan.get("profile")
    require(isinstance(profile, dict), "profile plan is missing")
    expected = {
        "owner": OWNER,
        "binary": "profile",
        "mode": "capture",
        "shape": "large",
        "orders": [["before", "after"], ["after", "before"]],
        "repeats": 2,
        "samples": 1,
        "warmup": 0,
        "collect_at_start": False,
        "events": ["Ir"],
        "expected_numbered_parts": 1,
        "reports": 4,
        "samples_total": 4,
    }
    for key, value in expected.items():
        require(profile.get(key) == value, f"profile plan {key} changed")
    return plan


def check_frozen_inputs(path: Path, label: str) -> dict[str, Any]:
    value = read(path)
    names = {
        "plan.json", "adoption-policy.json", "analysis-plan.json", "custody.py",
        "build.py", "capture.py", "profile.py", "quality.py", "probe_quality.py",
        "apply_candidate.py", "restore_candidate.py", "origin.json", "host.json",
        "inheritance.json", "architecture-inputs.json", "inputs/root-Cargo.lock",
        "inputs/rustfmt.toml", "quality_reuse.py", "protocol-review.md", "codegen.py",
        "codegen_analysis.py", "source-review.md",
        "toolchain.json", "quality-reuse.json", "quality-reuse/reuse-inputs.json",
    }
    candidate = P / "candidate"
    require(candidate.is_dir(), f"{label}: candidate archive is missing")
    names.update(str(item.relative_to(P)) for item in candidate.rglob("*")
                 if item.is_file())
    require(isinstance(value, dict)
            and value.get("schema") == "litchi.performance.0815.frozen-inputs.v1"
            and set(value) == {"schema", "packet", "root_inputs", "architecture", "unrelated"}
            and isinstance(value.get("packet"), dict)
            and set(value["packet"]) == names,
            f"{label}: frozen input envelope changed")
    for name, digest in value["packet"].items():
        path = P / name
        require(isinstance(digest, str) and re.fullmatch(r"[0-9a-f]{64}", digest)
                and path.is_file() and not path.is_symlink() and sha(path) == digest,
                f"{label}: frozen packet input changed: {name}")
    require(value.get("root_inputs") == {
        "Cargo.lock": sha(P / "inputs/root-Cargo.lock"),
        "rustfmt.toml": sha(P / "inputs/rustfmt.toml"),
    }, f"{label}: frozen root inputs changed")
    architecture = read(P / "architecture-inputs.json")
    unrelated = read(P / "origin.json").get("unrelated")
    require(value.get("architecture") == architecture and value.get("unrelated") == unrelated,
            f"{label}: frozen auxiliary custody changed")
    require(isinstance(architecture, dict) and isinstance(unrelated, dict)
            and all(isinstance(digest, str) and re.fullmatch(r"[0-9a-f]{64}", digest)
                    for digest in architecture.values())
            and all(isinstance(digest, str) and re.fullmatch(r"[0-9a-f]{64}", digest)
                    for digest in unrelated.values()),
            f"{label}: frozen auxiliary custody malformed")
    for name, digest in architecture.items():
        require(sha(ROOT / name) == digest, f"{label}: architecture input changed: {name}")
    for name, digest in unrelated.items():
        require(sha(ROOT / name) == digest, f"{label}: unrelated input changed: {name}")
    return value


def check_builds(plan: dict[str, Any]) -> tuple[dict[str, Any], dict[str, Any], dict[str, Any], dict[str, Any] | None]:
    cleanup_path = P / "cleanup.json"
    cleanup = read(cleanup_path) if cleanup_path.is_file() else None
    builds: dict[str, Any] = {}
    for leg in ("before", "after"):
        directory = P / f"build-{leg}"
        manifest = read(directory / "build.json")
        source_path = directory / "source.json"
        source = read(source_path)
        frozen_ref = manifest.get("frozen_inputs")
        require(isinstance(frozen_ref, dict), f"{leg}: frozen inputs receipt is missing")
        frozen_path = packet_artifact(frozen_ref, f"{leg} frozen inputs")
        require(frozen_path == (directory / "frozen-inputs.json").resolve(),
                f"{leg}: frozen inputs path changed")
        check_frozen_inputs(frozen_path, f"{leg} build")
        require(manifest.get("root_inputs") == {
            "Cargo.lock": sha(P / "inputs/root-Cargo.lock"),
            "rustfmt.toml": sha(P / "inputs/rustfmt.toml"),
        }, f"{leg}: root input receipt changed")
        require(manifest.get("architecture") == read(P / "architecture-inputs.json")
                and manifest.get("unrelated") == read(P / "origin.json").get("unrelated"),
                f"{leg}: auxiliary custody changed")
        source_ref = manifest.get("source")
        require(isinstance(source_ref, dict), f"{leg}: build source descriptor is missing")
        source_artifact = packet_artifact(source_ref, f"{leg} build source")
        require(source_artifact == source_path.resolve(), f"{leg}: source path changed")
        require(read(source_artifact) == source, f"{leg}: source descriptor content changed")
        files = source.get("files")
        require(isinstance(files, dict) and len(files) == 9196,
                f"{leg}: source census changed")
        probe = manifest.get("probe")
        require(isinstance(probe, dict) and probe, f"{leg}: probe manifest is missing")
        normalized_probe = {
            str(Path(name).as_posix()).removeprefix("probe-src/"): digest
            for name, digest in probe.items()
        }
        require(normalized_probe == probe_custody(), f"{leg}: probe manifest changed")
        lock = manifest.get("lock")
        if lock is not None:
            packet_artifact(lock, f"{leg} probe lock")
        rows = manifest.get("rows", manifest.get("commands"))
        require(isinstance(rows, list) and len(rows) == 3, f"{leg}: build row count changed")
        for row in rows:
            require(isinstance(row, dict) and row.get("exit_code") == 0,
                    f"{leg}: build command failed")
            packet_artifact(row.get("log"), f"{leg} build log")
        binaries = manifest.get("binaries")
        require(isinstance(binaries, dict) and set(binaries) == {"native", "allocation", "profile"},
                f"{leg}: binary matrix changed")
        verified = {
            name: external_binary(value, f"{leg} {name}", cleanup)
            for name, value in binaries.items()
        }
        builds[leg] = {
            "manifest": manifest,
            "source": source,
            "source_sha256": sha(source_path),
            "binaries": verified,
        }

    before = builds["before"]["source"]
    after = builds["after"]["source"]
    require(before.get("revision") == read(P / "origin.json").get("base"),
            "before source revision differs from origin")
    changed = {name for name in before["files"].keys() | after["files"].keys()
               if before["files"].get(name) != after["files"].get(name)}
    require(changed == set(plan["source_allowlist"]),
            f"build source chain changed outside allowlist: {sorted(changed)}")
    return builds, before, after, cleanup


def check_report(path: Path, leg: str, binary_name: str,
                 fixture: dict[str, Any], fixture_sample: dict[str, Any]) -> dict[str, Any]:
    report = read(path)
    label = relative(path)
    require(set(report) == REPORT_KEYS, f"{label}: report fields changed")
    require(report.get("schema") == PROFILE_SCHEMA, f"{label}: report schema changed")
    require(report.get("tool") == PROFILE_TOOL, f"{label}: report tool changed")
    require(report.get("mode") == "capture" and report.get("shape") == "large",
            f"{label}: report mode or shape changed")
    require((report.get("slides"), report.get("shapes_per_slide")) == (100, 100),
            f"{label}: dimensions changed")
    require(report.get("timing_scope") == "Package::opened_presentation only",
            f"{label}: timing scope changed")
    require(report.get("marker") == MARKER, f"{label}: marker changed")
    source = report.get("source")
    require(isinstance(source, dict) and set(source) == SOURCE_KEYS
            and isinstance(source.get("sha256"), str)
            and re.fullmatch(r"[0-9a-f]{64}", source["sha256"])
            and type(source.get("bytes")) is int and source["bytes"] > 0,
            f"{label}: source fields changed")
    require(report.get("fixture") == fixture.get("fixture"), f"{label}: fixture changed")
    require(isinstance(report.get("fixture"), dict)
            and set(report["fixture"]) == FIXTURE_KEYS,
            f"{label}: fixture fields changed")
    require(report.get("source") == fixture.get("source"), f"{label}: source differs from 0813")
    require(report.get("warmup") == 0 and report.get("samples_requested") == 1,
            f"{label}: sample policy changed")
    require(isinstance(report.get("allocator"), dict)
            and set(report["allocator"]) == {
                "binary", "allocator", "instrumentation", "counter_revision",
            }
            and report["allocator"] == {
                "binary": binary_name,
                "allocator": "Rust system allocator",
                "instrumentation": "none",
                "counter_revision": None,
            }, f"{label}: allocator identity changed")
    samples = report.get("samples")
    require(isinstance(samples, list) and len(samples) == 1, f"{label}: sample count changed")
    sample = samples[0]
    require(set(sample) == {
        "index", "elapsed_ns", "metrics", "source_sha256", "output", "verification",
    }, f"{label}: sample fields changed")
    require(sample.get("index") == 0, f"{label}: sample index changed")
    require(type(sample.get("elapsed_ns")) is int and sample["elapsed_ns"] > 0,
            f"{label}: elapsed value is invalid")
    require(sample.get("source_sha256") == fixture_sample["source_sha256"],
            f"{label}: sample source differs from 0813")
    require(sample.get("output") == fixture_sample.get("output"),
            f"{label}: output identity differs from 0813")
    require(sample.get("verification") == fixture_sample.get("verification"),
            f"{label}: full semantic verification differs from 0813")
    require("allocation" not in sample, f"{label}: profile report contains allocation metrics")
    verification = sample.get("verification")
    require(isinstance(verification, dict) and set(verification) == COMMON_VERIFICATION_KEYS,
            f"{label}: verification fields changed")
    require(verification.get("semantic_check") is True
            and verification.get("reopened") is True
            and isinstance(verification.get("expected_text"), str)
            and verification.get("expected_text") == verification.get("actual_text")
            and isinstance(verification.get("semantic_text_sha256"), str)
            and re.fullmatch(r"[0-9a-f]{64}", verification["semantic_text_sha256"]),
            f"{label}: semantic verification failed")
    require(type(verification.get("semantic_text_bytes")) is int
            and verification["semantic_text_bytes"] >= 0,
            f"{label}: semantic text bytes changed")
    output = sample.get("output")
    require(isinstance(output, dict) and set(output) == {"bytes", "sha256"}
            and type(output.get("bytes")) is int and output["bytes"] >= 0
            and isinstance(output.get("sha256"), str)
            and re.fullmatch(r"[0-9a-f]{64}", output["sha256"]),
            f"{label}: output fields changed")
    require(verification.get("readback_bytes") == output["bytes"]
            and verification.get("readback_sha256") == output["sha256"],
            f"{label}: readback changed")
    require(verification.get("marker_matches") is None
            and verification.get("unknown_namespace_check") is None
            and verification.get("unknown_namespace_occurrences") is None,
            f"{label}: capture namespace or marker oracle changed")
    require(sample.get("metrics") == {
        "elapsed_ns": sample["elapsed_ns"],
        "slides": 100,
        "shapes_per_slide": 100,
        "captured_slides": 100,
        "captured_shapes_per_slide": 100,
    }, f"{label}: sample metrics changed")
    return {
        "path": relative(path),
        "sha256": sha(path),
        "bytes": path.stat().st_size,
        "source": report["source"],
        "output": sample["output"],
        "verification": sample["verification"],
        "fixture": report["fixture"],
    }


def function_record(parser: Any, parsed: dict[str, Any], name: str,
                    *, required: bool = False) -> dict[str, Any] | None:
    matches = [function for function in parsed["functions"].values()
               if function["name"] == name]
    require(len(matches) <= 1, f"{name}: duplicate function identities")
    if not matches:
        require(not required, f"{name}: function is missing")
        return None
    function = matches[0]
    names = parsed["names"]
    view = parser.function_view(function, names)
    edges = []
    for edge in function["edges"]:
        edges.append({
            **edge,
            "caller": names.get(edge["caller_id"], "<unnamed>"),
            "callee": names.get(edge["callee_id"], "<unnamed>"),
        })
    incoming = parser.incoming_edges(parsed, function["id"])
    call_counts = [
        {"callee": edge["callee"], "calls": edge["calls"]}
        for edge in view["direct_children"]
    ]
    return {
        "id": function["id"],
        "name": function["name"],
        "records": function["records"],
        "self_ir": function["self_ir"],
        "edges": edges,
        "direct_children": view["direct_children"],
        "direct_children_ir": view["direct_children_ir"],
        "inclusive_ir": function["self_ir"] + view["direct_children_ir"],
        "call_counts": call_counts,
        "outgoing_call_count": sum(item["calls"] for item in call_counts),
        "incoming_edges": incoming,
    }


def named_edges(record: dict[str, Any] | None, name: str) -> list[dict[str, Any]]:
    if record is None:
        return []
    return [edge for edge in record["direct_children"] if edge["callee"] == name]


def profile_raw(parser: Any, numbered: Path, terminal: Path, stem: str) -> dict[str, Any]:
    number = parser.parse_raw(numbered)
    final = parser.parse_raw(terminal, allow_empty=True)
    require(number["header"]["events"] == ["Ir"], f"{stem}: event set changed")
    require(number["header"]["part"] == 1
            and number["header"]["trigger"] == "--dump-after=" + OWNER,
            f"{stem}: positive dump header changed")
    require(number["header"]["summary_ir"] > 1000
            and number["header"]["totals_ir"] == number["header"]["summary_ir"],
            f"{stem}: positive Ir summary changed")
    require(final["header"]["events"] == ["Ir"]
            and final["header"]["part"] == 2
            and final["header"]["trigger"] == "Program termination"
            and final["header"]["summary_ir"] == 0
            and final["header"]["totals_ir"] == 0,
            f"{stem}: termination dump is not empty")
    require(number["header"].get("command") and final["header"].get("command"),
            f"{stem}: Callgrind command header is missing")
    require(number["header"]["command"] == final["header"]["command"],
            f"{stem}: Callgrind command headers differ")
    owner_ids = [fid for fid, value in number["functions"].items()
                 if value["name"] == OWNER]
    require(len(owner_ids) == 1, f"{stem}: exact wrapper owner is not unique")
    owner_id = owner_ids[0]
    incoming = parser.incoming_edges(number, owner_id)
    require(len(incoming) == 1 and incoming[0]["calls"] == 1
            and incoming[0]["inclusive_ir"] == number["header"]["summary_ir"],
            f"{stem}: owner incoming edge changed")
    owner_view = parser.function_view(number["functions"][owner_id], number["names"])
    require(owner_view["self_ir"] + owner_view["direct_children_ir"]
            == number["header"]["summary_ir"],
            f"{stem}: owner self plus immediate children does not reconstruct summary")
    function_views = [parser.function_view(value, number["names"])
                      for value in number["functions"].values()]
    all_self = sum(value["self_ir"] for value in function_views)
    require(all_self == number["header"]["summary_ir"],
            f"{stem}: whole-function self Ir does not reconstruct summary")

    selected = {
        "scanner": function_record(parser, number, SCANNER, required=True),
        # The candidate may leave this generic symbol in the binary through
        # unrelated users or remove it entirely.  The strict check is the
        # scanner's direct edge, not global symbol presence or call count.
        "process_event": function_record(parser, number, PROCESS_EVENT),
        "reader_read_event_impl": function_record(parser, number, READ_EVENT, required=True),
        "namespace_resolver_push": function_record(parser, number, RESOLVER_PUSH, required=True),
        "owner": function_record(parser, number, OWNER, required=True),
    }
    scanner = selected["scanner"]
    resolver_push = selected["namespace_resolver_push"]
    require(scanner is not None and resolver_push is not None,
            f"{stem}: selected mechanism function is missing")
    require(resolver_push["self_ir"] > 0, f"{stem}: namespace push has no self Ir")
    scanner_process = named_edges(scanner, PROCESS_EVENT)
    scanner_read = named_edges(scanner, READ_EVENT)
    process_event = selected["process_event"]
    return {
        "file": relative(numbered),
        "sha256": sha(numbered),
        "bytes": numbered.stat().st_size,
        "termination": {
            "file": relative(terminal),
            "sha256": sha(terminal),
            "bytes": terminal.stat().st_size,
            "summary_ir": 0,
        },
        "summary_ir": number["header"]["summary_ir"],
        "owner": {
            "id": owner_id,
            "name": OWNER,
            "incoming": incoming[0],
            "self_ir": owner_view["self_ir"],
            "direct_children_ir": owner_view["direct_children_ir"],
            "direct_children": owner_view["direct_children"],
            "partition_equation": "owner self Ir + immediate-child inclusive Ir = owner inclusive Ir",
            "partition_disjoint": True,
            "nested_inclusive_rows_excluded": True,
        },
        "parser": number["statistics"],
        "all_function_self_ir": all_self,
        "mechanism_functions": selected,
        "mechanism": {
            "scanner_metrics": {
                "records": scanner["records"],
                "self_ir": scanner["self_ir"],
                "inclusive_ir": scanner["inclusive_ir"],
                "direct_children_ir": scanner["direct_children_ir"],
                "call_counts": scanner["call_counts"],
                "outgoing_call_count": scanner["outgoing_call_count"],
            },
            "scanner_to_process_event_edges": scanner_process,
            "scanner_to_read_event_impl_edges": scanner_read,
            "namespace_resolver_push_self_ir": resolver_push["self_ir"],
            "namespace_resolver_push_direct_children_ir": resolver_push["direct_children_ir"],
            "global_process_event_incoming_edges": (
                [] if process_event is None else process_event["incoming_edges"]
            ),
            "global_process_event_zero_not_required": True,
        },
        "validation": {
            "exact_trigger": True,
            "summary_nonzero": True,
            "termination_zero_ir": True,
            "exactly_one_positive_owner_call": True,
            "owner_inclusive_matches_summary": True,
            "owner_self_plus_immediate_children_equals_summary": True,
            "whole_function_self_ir_equals_summary": True,
            "compressed_names_resolved": number["statistics"]["compressed_name_declarations"] > 0,
            "relative_or_wildcard_positions_parsed": any(
                number["statistics"]["position_kinds"].get(kind, 0) > 0
                for kind in ("relative", "wildcard")
            ),
        },
    }


def profile_receipt(row: dict[str, Any], repeat: int, leg: str,
                    plan: dict[str, Any], builds: dict[str, Any], parser: Any,
                    fixture: dict[str, Any], fixture_sample: dict[str, Any]) -> dict[str, Any]:
    stem = f"{repeat}-{leg}"
    require(row.get("schema") == "litchi.performance.0815.callgrind-receipt.v1",
            f"{stem}: receipt schema changed")
    require((row.get("repeat"), row.get("leg")) == (repeat, leg),
            f"{stem}: receipt order changed")
    require(row.get("exit_code") == 0, f"{stem}: profiler process failed")
    require(row.get("binary") == builds[leg]["manifest"]["binaries"]["profile"],
            f"{stem}: binary descriptor changed")
    require(row.get("driver_sha256") == sha(P / "profile.py"),
            f"{stem}: profile driver changed")
    started, ended = row.get("started"), row.get("ended")
    require(isinstance(started, (int, float)) and isinstance(ended, (int, float))
            and started <= ended, f"{stem}: timestamps are malformed")
    command = row.get("command")
    require(isinstance(command, list), f"{stem}: command is malformed")
    raw_path = P / "profiles" / f"{stem}.callgrind"
    report_path = P / "profiles" / f"{stem}.json"
    expected = [
        "taskset", "-c", str(plan["cpu"]), "valgrind", "--tool=callgrind",
        "--collect-atstart=no", "--toggle-collect=" + OWNER,
        "--zero-before=" + OWNER, "--dump-after=" + OWNER,
        "--callgrind-out-file=" + str(raw_path),
        builds[leg]["manifest"]["binaries"]["profile"]["path"],
        "--mode", "capture", "--shape", "large", "--samples", "1",
        "--warmup", "0", "--output", str(report_path),
    ]
    require(command == expected, f"{stem}: Callgrind command changed")
    artifacts = row.get("artifacts")
    require(isinstance(artifacts, dict)
            and set(artifacts) == {f"{stem}.{suffix}" for suffix in EXPECTED_ARTIFACTS},
            f"{stem}: artifact set changed")
    paths = {name: packet_artifact(value, f"{stem} {name}")
             for name, value in artifacts.items()}
    require(paths[f"{stem}.json"] == report_path.resolve(), f"{stem}: report path changed")
    require(paths[f"{stem}.callgrind"] == raw_path.resolve(), f"{stem}: raw path changed")
    numbered = P / "profiles" / f"{stem}.callgrind.1"
    require(paths[f"{stem}.callgrind.1"] == numbered.resolve(),
            f"{stem}: numbered raw path changed")
    unexpected = sorted((P / "profiles").glob(f"{stem}.callgrind.*"))
    require([path.name for path in unexpected] == [f"{stem}.callgrind.1"],
            f"{stem}: additional numbered publication exists")
    report_identity = check_report(paths[f"{stem}.json"], leg,
                                   Path(row["binary"]["path"]).name,
                                   fixture, fixture_sample)
    raw = profile_raw(parser, numbered, paths[f"{stem}.callgrind"], stem)
    return {
        "repeat": repeat,
        "leg": leg,
        "report": report_identity,
        "log": relative(paths[f"{stem}.log"]),
        "command": command,
        "timestamps": {"started": started, "ended": ended},
        "raw": raw,
    }


def analyze() -> dict[str, Any]:
    plan = check_plan()
    builds, before, after, cleanup = check_builds(plan)
    fixture, fixture_sample, fixture_meta = sealed_fixture()
    parser = load_hash_bound_parser()
    complete = read(P / "profiles" / "complete.json")
    expected_complete = {
        "schema": "litchi.performance.0815.callgrind.complete.v1",
        "processes": 4,
        "reports": 4,
        "samples": 4,
        "plan_sha256": sha(P / "plan.json"),
        "build_before_sha256": sha(P / "build-before/build.json"),
        "build_after_sha256": sha(P / "build-after/build.json"),
        "scope": "namespace_uri_probe::capture_region_0793 only; Ir attribution without latency or RSS claim",
    }
    for key, value in expected_complete.items():
        require(complete.get(key) == value, f"profile completion field changed: {key}")
    receipts = read(P / "profiles" / "receipts.json")
    require(isinstance(receipts, list) and len(receipts) == 4,
            "profile receipt cardinality changed")
    jobs = [(repeat, leg) for repeat, order in enumerate(plan["profile"]["orders"])
            for leg in order]
    profiles: list[dict[str, Any]] = []
    previous_end = float("-inf")
    for row, (repeat, leg) in zip(receipts, jobs):
        result = profile_receipt(row, repeat, leg, plan, builds, parser,
                                 fixture, fixture_sample)
        require(previous_end <= result["timestamps"]["started"],
                f"{repeat}-{leg}: profile process overlaps prior process")
        previous_end = result["timestamps"]["ended"]
        profiles.append(result)

    pairs = {(item["repeat"], item["leg"]): item for item in profiles}
    mechanism_pairs = []
    for repeat in range(2):
        before_profile = pairs[(repeat, "before")]["raw"]
        after_profile = pairs[(repeat, "after")]["raw"]
        before_edges = before_profile["mechanism"]["scanner_to_process_event_edges"]
        after_edges = after_profile["mechanism"]["scanner_to_process_event_edges"]
        require(not before_edges and not after_edges,
                f"profile pair {repeat}: scanner process_event edge unexpectedly present")
        require(before_profile["mechanism"]["scanner_to_read_event_impl_edges"],
                f"profile pair {repeat}: baseline reader edge missing")
        require(after_profile["mechanism"]["scanner_to_read_event_impl_edges"],
                f"profile pair {repeat}: candidate reader edge missing")
        mechanism_pairs.append({
            "repeat": repeat,
            "before_scanner_metrics": before_profile["mechanism"]["scanner_metrics"],
            "after_scanner_metrics": after_profile["mechanism"]["scanner_metrics"],
            "before_scanner_to_process_event": before_edges,
            "after_scanner_to_process_event": after_edges,
            "before_scanner_to_read_event_impl": before_profile["mechanism"]["scanner_to_read_event_impl_edges"],
            "after_scanner_to_read_event_impl": after_profile["mechanism"]["scanner_to_read_event_impl_edges"],
            "before_namespace_resolver_push_self_ir": before_profile["mechanism"]["namespace_resolver_push_self_ir"],
            "after_namespace_resolver_push_self_ir": after_profile["mechanism"]["namespace_resolver_push_self_ir"],
            "global_process_event_incoming_before": before_profile["mechanism"]["global_process_event_incoming_edges"],
            "global_process_event_incoming_after": after_profile["mechanism"]["global_process_event_incoming_edges"],
            "global_zero_call_check": False,
        })

    return {
        "schema": "litchi-0815-callgrind-profile-analysis-v1",
        "packet": "change-0815",
        "plan": {
            "path": "plan.json",
            "sha256": sha(P / "plan.json"),
            "schema": plan["schema"],
            "owner": OWNER,
            "cpu": plan["cpu"],
        },
        "build": {
            "before": {
                "path": "build-before/build.json",
                "sha256": sha(P / "build-before/build.json"),
                "source_sha256": builds["before"]["source_sha256"],
                "binaries": builds["before"]["binaries"],
            },
            "after": {
                "path": "build-after/build.json",
                "sha256": sha(P / "build-after/build.json"),
                "source_sha256": builds["after"]["source_sha256"],
                "binaries": builds["after"]["binaries"],
            },
            "source_chain": {
                "before_file_count": len(before["files"]),
                "after_file_count": len(after["files"]),
                "changed_files": sorted(set(plan["source_allowlist"])),
            },
            # This result must replay byte-for-byte before and after target
            # cleanup.  ``check_builds`` has already verified each of the six
            # binaries either live or through the exact cleanup witness; the
            # fact that the witness exists is intentionally not recorded.
            "binary_custody_verified": True,
            "cleanup_contract": {
                "path": "cleanup.json",
                "target": str(c.TARGET),
                "binary_count": 6,
                "binary_names": [
                    "before-native",
                    "before-allocation",
                    "before-profile",
                    "after-native",
                    "after-allocation",
                    "after-profile",
                ],
            },
        },
        "fixture_oracle": fixture_meta,
        "profiles": profiles,
        "mechanism": {
            "pairs": mechanism_pairs,
            "scanner_to_process_event_absent_in_both_profiles": True,
            "scanner_call_metrics_recorded_without_savings_claim": True,
            "other_nsreader_callers_are_not_a_global_zero_call_gate": True,
            "namespace_push_costs_retained_before_after": True,
        },
        "summary": {
            "profile_count": len(profiles),
            "positive_ir_profiles": sum(item["raw"]["summary_ir"] > 0 for item in profiles),
            "inclusive_ir_min": min(item["raw"]["summary_ir"] for item in profiles),
            "inclusive_ir_max": max(item["raw"]["summary_ir"] for item in profiles),
            "owner_self_ir_values": [item["raw"]["owner"]["self_ir"] for item in profiles],
            "all_exact_scope_checks_pass": True,
            "all_output_fixture_checks_pass": True,
            "all_four_profiles_complete": True,
        },
        "scope": "Callgrind Ir is guest-instruction attribution for the exact capture wrapper; it is not native latency, RSS, a CPU fraction, a phase fraction, or a production speedup.",
        "claims": [
            "The exact owner has one incoming call and its self Ir plus immediate-child inclusive Ir reconstructs the owner summary.",
            "Whole-function self Ir reconstructs each positive publication; nested inclusive rows are retained as diagnostics and are not added.",
            "The scanner-to-NsReader::process_event edge is absent in both retained profiles; this reader makes no before/after savings claim.",
            "Guest-Ir differences do not add an adoption threshold or imply a latency, RSS, cycle, or phase-fraction gain.",
            "Report source, output, and full semantic verification are bound to the sealed 0813 fixture.",
        ],
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--write", action="store_true")
    parser.add_argument("--check", action="store_true")
    args = parser.parse_args()
    require(args.write ^ args.check, "choose exactly one of --write or --check")
    write_or_check(P / "profile-analysis.json", analyze(), args.check)
    print("0815 Callgrind analysis PASS", flush=True)
    return 0


if __name__ == "__main__":
    sys.dont_write_bytecode = True
    try:
        raise SystemExit(main())
    except EvidenceError as error:
        print(f"0815 Callgrind analysis failed: {error}", file=sys.stderr)
        raise SystemExit(1)
