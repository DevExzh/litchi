#!/usr/bin/env python3
"""Validate and attribute the 0709 DOCX ordinary-save Callgrind packet.

The analyzer selects the measured part from a positive raw incoming edge from
``ordinary_save::run_case``.  Setup probes and reference publications remain
in the report, including their raw caller edges and any warnings.  The
publication owner is the exact demangled ``Package::write_plain`` name; the
four monomorphs share that name while their ``write_plain::{{closure}}``
helpers are intentionally outside the toggle pattern.

Callgrind ``Ir`` is guest-instruction attribution.  This packet does not infer
allocation counts, native latency, hardware counters, or production changes.
"""

from __future__ import annotations

import argparse
import copy
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import re
import subprocess
import sys
from typing import Any, Callable, NoReturn


HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
RETAINED_0705 = HERE.parent / "change-0705" / "analyze_profiles.py"
RETAINED_0521 = HERE.parent / "change-0521" / "analyze_profiles.py"

EDIT_OWNER = "litchi_perf_baseline::ordinary_save::Owner::edit"
WRITE_PLAIN_OWNER = (
    "litchi_docx::package::codec::<impl "
    "litchi_docx::package::model::Package>::write_plain"
)
MEASURED_PARENT = "litchi_perf_baseline::ordinary_save::run_case"
DOCUMENT_MUT = (
    "litchi_docx::package::package::document::<impl "
    "litchi_docx::package::model::Package>::document_mut"
)
PHASES = ("edit", "counting_publish")
CORPORA = ("generated", "numbered-list")
PHASE_LABELS = {
    "edit": "edit",
    "counting_publish": "serialize-to-counting-sink",
}
PHASE_TIMING_SCOPES = {
    "edit": "the semantic edit and its commit only; the documented open and every verification are outside the clock",
    "counting_publish": "the documented sequential serialization into a bounded counting sink only; the open, the edit and the byte accounting are outside the clock",
}
PERL_ENV = {"PERL_HASH_SEED": "0", "PERL_PERTURB_KEYS": "0"}
CALLGRIND_FUNCTION_RE = re.compile(r"^(fn|cfn)=\((\d+)\)(?:\s+(.*))?$")
CALLGRIND_CALLS_RE = re.compile(r"^calls=([,\d]+)")
HEX = set("0123456789abcdef")


class EvidenceError(ValueError):
    """A missing, malformed, or contradictory profile artifact."""


def fail(message: str) -> NoReturn:
    raise EvidenceError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def sha(path: Path) -> str:
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
    except OSError as error:
        fail(f"cannot hash {path}: {error}")
    return digest.hexdigest()


def canonical_json(value: Any) -> bytes:
    return (json.dumps(value, sort_keys=True, separators=(",", ":")) + "\n").encode()


def digest_json(value: Any) -> str:
    return hashlib.sha256(canonical_json(value)).hexdigest()


def read_json(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing JSON input: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(f"invalid JSON in {path}: {error}")


def read_text(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8", errors="replace")
    except OSError as error:
        fail(f"cannot read {path}: {error}")


def write_json(path: Path, value: Any) -> None:
    path.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n", encoding="utf-8")


def relative(path: Path) -> str:
    try:
        return str(path.relative_to(HERE))
    except ValueError as error:
        fail(f"path is outside packet: {path}")


def load_retained_helpers() -> Any:
    """Load the retained 0705 wrapper and its immutable 0521 parser helpers."""

    require(RETAINED_0705.is_file(), f"missing retained helper: {RETAINED_0705}")
    require(RETAINED_0521.is_file(), f"missing retained helper: {RETAINED_0521}")
    spec = importlib.util.spec_from_file_location("retained_0705_profile_0709", RETAINED_0705)
    require(spec is not None and spec.loader is not None,
            f"cannot load retained helper: {RETAINED_0705}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module.H


H = load_retained_helpers()


def retained_helper_identities() -> dict[str, str]:
    identity_path = HERE / "helper-identities.json"
    if identity_path.is_file():
        value = read_json(identity_path)
        require(isinstance(value, dict), "helper-identities.json is not an object")
        result: dict[str, str] = {}
        for name, digest in value.items():
            require(isinstance(name, str) and isinstance(digest, str),
                    "helper identity entry is malformed")
            path = REPO / name
            require(path.is_file() and not path.is_symlink(),
                    f"retained helper is missing: {name}")
            require(sha(path) == digest, f"retained helper hash differs: {name}")
            result[name] = digest
        return dict(sorted(result.items()))
    return {
        str(RETAINED_0705.relative_to(REPO)): sha(RETAINED_0705),
        str(RETAINED_0521.relative_to(REPO)): sha(RETAINED_0521),
    }


def source_census() -> dict[str, str]:
    paths: list[Path] = [REPO / "Cargo.toml", REPO / "Cargo.lock"]
    paths.extend(path for path in (REPO / ".cargo").rglob("*") if path.is_file())
    for folder in ("crates", "tools/perf-baseline"):
        paths.extend(
            path
            for path in (REPO / folder).rglob("*")
            if path.is_file()
            and "target" not in path.parts
            and (path.suffix == ".rs" or path.name in {"Cargo.toml", "Cargo.lock"})
        )
    return {str(path.relative_to(REPO)): sha(path) for path in sorted(set(paths))}


def load_plan() -> dict[str, Any]:
    plan = read_json(HERE / "plan.json")
    require(isinstance(plan, dict), "plan.json is not an object")
    require(isinstance(plan.get("revision"), str) and plan["revision"],
            "plan revision is missing")
    require(plan.get("native") == {"repeats": 3, "samples": 100, "warmup": 10},
            "native plan changed")
    require(plan.get("allocator") == {"repeats": 2, "samples": 3, "warmup": 0},
            "allocator plan changed")
    require(isinstance(plan.get("cpu"), int) and plan["cpu"] >= 0,
            "plan CPU is invalid")
    corpora = plan.get("corpora")
    require(isinstance(corpora, list) and len(corpora) == 3,
            "ordinary-save corpus matrix changed")
    ids: set[str] = set()
    for corpus in corpora:
        require(isinstance(corpus, dict), "ordinary-save corpus entry is malformed")
        identity = corpus.get("id")
        require(isinstance(identity, str) and identity not in ids,
                "ordinary-save corpus id is missing or duplicated")
        ids.add(identity)
    require(ids == {"generated", "numbered-list", "alt-chunk-header"},
            "ordinary-save corpus ids changed")
    return plan


def load_profile_plan(plan: dict[str, Any]) -> dict[str, Any]:
    profile = read_json(HERE / "profile-plan.json")
    require(isinstance(profile, dict), "profile-plan.json is not an object")
    require(profile.get("schema_version") == 1, "profile plan schema changed")
    require(profile.get("revision") == plan["revision"], "profile revision differs")
    require(profile.get("corpora") == list(CORPORA), "profile corpus matrix changed")
    require(profile.get("phases") == list(PHASES), "profile phase matrix changed")
    require(profile.get("repeats") == 2 and profile.get("samples") == 1
            and profile.get("warmup") == 0, "profile sample policy changed")
    require(profile.get("owners") == {"edit": EDIT_OWNER,
                                      "counting_publish": WRITE_PLAIN_OWNER},
            "profile owners changed")
    require(profile.get("owner_match") == {"edit": "exact", "counting_publish": "exact"},
            "profile owner match policy changed")
    require(profile.get("measured_parent") == MEASURED_PARENT,
            "profile measured caller changed")
    require(profile.get("expected_numbered_dumps") == {
        "edit": {"generated": 5, "numbered-list": 5},
        "counting_publish": {"generated": 6, "numbered-list": 5},
    },
            "profile dump-count policy changed")
    require(profile.get("setup_parents") == {
        "edit": {
            "litchi_perf_baseline::ordinary_save::build_corpus": 1,
            "litchi_perf_baseline::ordinary_save::publish_reference": 3,
        },
        "counting_publish": {
            "generated": {
                "litchi_perf_baseline::semantic_docx_bytes": 1,
                "litchi_opc::atomic::replace_with_impl": 4,
            },
            "numbered-list": {
                "litchi_opc::atomic::replace_with_impl": 4,
            },
        },
    }, "profile setup call policy changed")
    options = profile.get("options")
    require(isinstance(options, dict), "profile Callgrind options are missing")
    for phase in PHASES:
        value = options.get(phase)
        require(isinstance(value, list) and all(isinstance(item, str) for item in value),
                f"{phase}: profile Callgrind options are invalid")
        for token in ("--tool=callgrind", "--collect-atstart=no",
                      "--toggle-collect=" + profile["owners"][phase],
                      "--zero-before=" + profile["owners"][phase],
                      "--dump-after=" + profile["owners"][phase]):
            require(token in value, f"{phase}: profile options omit {token}")
    return profile


def build_info() -> dict[str, Any]:
    records_path = HERE / "build-baseline.json"
    records = read_json(records_path)
    require(isinstance(records, list), "build-baseline.json is not a list")
    matches = [item for item in records if isinstance(item, dict)
               and Path(str(item.get("binary", ""))).name == "baseline-native"]
    require(len(matches) == 1, "baseline-native build record is not unique")
    record = matches[0]
    require(record.get("exit_code") == 0, "baseline-native build failed")
    binary = Path(str(record.get("binary", ""))).resolve()
    binary_sha = record.get("binary_sha256")
    binary_bytes = record.get("binary_bytes")
    require(isinstance(binary_sha, str) and len(binary_sha) == 64,
            "baseline-native binary digest is missing")
    require(isinstance(binary_bytes, int) and binary_bytes > 0,
            "baseline-native binary size is invalid")
    require(binary.is_file() and not binary.is_symlink(),
            f"baseline-native binary is missing: {binary}")
    require(sha(binary) == binary_sha and binary.stat().st_size == binary_bytes,
            "baseline-native binary identity changed")
    source_path = HERE / "source-baseline.json"
    source = read_json(source_path)
    require(isinstance(source, dict) and source, "source-baseline.json is invalid")
    require(record.get("source_manifest_sha256") == sha(source_path),
            "build/source manifest binding changed")
    symbols_json_path = HERE / "symbols.json"
    symbols_text_path = HERE / "symbols.txt"
    symbols = read_json(symbols_json_path)
    require(isinstance(symbols, dict), "symbols.json is not an object")
    require(symbols.get("binary_sha256") == binary_sha,
            "nm symbols bind a different native binary")
    require(symbols.get("log_sha256") == sha(symbols_text_path),
            "nm symbols text digest changed")
    lines = read_text(symbols_text_path).splitlines()
    require(sum(line.endswith(" " + EDIT_OWNER) for line in lines) == 1,
            "nm symbols do not prove one exact Owner::edit symbol")
    require(sum(line.endswith(" " + MEASURED_PARENT) for line in lines) == 1,
            "nm symbols do not prove one exact run_case symbol")
    require(sum(line.endswith(" " + WRITE_PLAIN_OWNER) for line in lines) == 4,
            "nm symbols do not prove four exact write_plain symbols")
    require(any(line.endswith(" " + WRITE_PLAIN_OWNER + "::{{closure}}") for line in lines),
            "nm symbols omit write_plain closure evidence")
    return {
        "record": relative(records_path),
        "record_sha256": sha(records_path),
        "binary": str(binary),
        "binary_sha256": binary_sha,
        "binary_bytes": binary_bytes,
        "source_manifest": relative(source_path),
        "source_manifest_sha256": sha(source_path),
        "source": source,
        "source_census_sha256": digest_json(source),
        "symbols": {
            "json": relative(symbols_json_path),
            "json_sha256": sha(symbols_json_path),
            "text": relative(symbols_text_path),
            "text_sha256": sha(symbols_text_path),
            "edit_count": 1,
            "run_case_count": 1,
            "write_plain_count": 4,
        },
    }


def expected_jobs(plan: dict[str, Any], profile: dict[str, Any]) -> list[dict[str, Any]]:
    corpora = {item["id"]: item for item in plan["corpora"]}
    result: list[dict[str, Any]] = []
    order = 0
    for repeat in range(1, profile["repeats"] + 1):
        phases = list(PHASES)
        corpus_ids = list(CORPORA)
        if repeat % 2 == 0:
            phases.reverse()
            corpus_ids.reverse()
        for phase in phases:
            for corpus_id in corpus_ids:
                corpus = corpora[corpus_id]
                generated = corpus["origin"] == "generated-harness-corpus"
                result.append({
                    "name": f"profile-r{repeat}-{corpus_id}-{phase}",
                    "repeat": repeat,
                    "corpus_id": corpus_id,
                    "corpus": corpus,
                    "phase": phase,
                    "case": ("docx_ordinary_save_" if generated
                             else "docx_real_file_ordinary_save_") + phase,
                    "order_index": order,
                })
                order += 1
    return result


def validate_source_custody(receipt: dict[str, Any], build: dict[str, Any], label: str) -> dict[str, Any]:
    custody = receipt.get("current_checkout_source")
    require(isinstance(custody, dict) and custody.get("unchanged_during_child") is True,
            f"{label}: source custody is missing or changed")
    expected = build["source"]
    for side in ("before", "after"):
        filename = custody.get(side + "_artifact")
        require(isinstance(filename, str) and Path(filename).name == filename,
                f"{label}: source {side} artifact path is invalid")
        path = HERE / filename
        value = read_json(path)
        require(value == expected, f"{label}: source {side} census differs")
        require(custody.get(side + "_sha256") == digest_json(value),
                f"{label}: source {side} content digest differs")
        require(custody.get(side + "_file_sha256") == sha(path),
                f"{label}: source {side} file digest differs")
    require(custody.get("before_sha256") == custody.get("after_sha256"),
            f"{label}: source census changed during child")
    return {
        "before": custody["before_artifact"],
        "after": custody["after_artifact"],
        "before_sha256": custody["before_file_sha256"],
        "after_sha256": custody["after_file_sha256"],
    }


def fixture_binding(corpus: dict[str, Any]) -> dict[str, Any] | None:
    if corpus["origin"] == "generated-harness-corpus":
        return None
    path = (REPO / corpus["path"]).resolve()
    require(path.is_file() and not path.is_symlink(), f"fixture is missing: {path}")
    return {
        "plan_path": corpus["path"],
        "resolved_path": str(path),
        "bytes": path.stat().st_size,
        "sha256": sha(path),
    }


def validate_receipt(job: dict[str, Any], receipt: dict[str, Any], plan: dict[str, Any],
                     profile: dict[str, Any], build: dict[str, Any],
                     plan_sha: str, profile_sha: str, constraints_sha: str,
                     script_sha: str) -> dict[str, Any]:
    label = job["name"] + ".receipt.json"
    for key, expected in {
        "schema_version": 1,
        "lane": "profile",
        "name": job["name"],
        "repeat": job["repeat"],
        "corpus_id": job["corpus_id"],
        "phase": job["phase"],
        "case": job["case"],
        "samples": 1,
        "warmup": 0,
        "cpu": plan["cpu"],
        "owner": profile["owners"][job["phase"]],
        "owner_match": profile["owner_match"][job["phase"]],
        "measured_parent": MEASURED_PARENT,
        "exit_code": 0,
        "binary_sha256": build["binary_sha256"],
        "binary_bytes": build["binary_bytes"],
        "build_record_sha256": build["record_sha256"],
        "build_source_manifest": build["source_manifest"],
        "build_source_manifest_sha256": build["source_manifest_sha256"],
        "retained_source_census_sha256": build["source_census_sha256"],
        "retained_source_entry_count": len(build["source"]),
        "plan_sha256": plan_sha,
        "profile_plan_sha256": profile_sha,
        "constraints_sha256": constraints_sha,
        "script_sha256": script_sha,
    }.items():
        require(receipt.get(key) == expected, f"{label}: {key} differs")
    command = receipt.get("command")
    require(isinstance(command, list) and command[:4] ==
            ["taskset", "-c", str(plan["cpu"]), "valgrind"],
            f"{label}: command prefix differs")
    options = profile["options"][job["phase"]]
    for token in options:
        require(token in command, f"{label}: command omits {token}")
    require(f"--warmup" in command and command[command.index("--warmup") + 1] == "0",
            f"{label}: warmup command differs")
    require(f"--samples" in command and command[command.index("--samples") + 1] == "1",
            f"{label}: sample command differs")
    require("--case" in command and command[command.index("--case") + 1] == job["case"],
            f"{label}: case command differs")
    require(Path(command[command.index("--json") + 1]).name == job["name"] + ".json",
            f"{label}: JSON output command differs")
    callgrind = command[[item.startswith("--callgrind-out-file=") for item in command].index(True)]
    require(Path(callgrind.split("=", 1)[1]).name == job["name"] + ".callgrind",
            f"{label}: Callgrind output command differs")
    real = job["corpus"]["origin"] == "caller-named-real-file"
    require(("--ooxml-file" in command) == real,
            f"{label}: real-file command binding differs")
    if real:
        require(command[command.index("--ooxml-file") + 1] == job["corpus"]["path"],
                f"{label}: real-file path command differs")
    custody = validate_source_custody(receipt, build, label)
    planned_fixture = fixture_binding(job["corpus"])
    fixture = receipt.get("fixture")
    require(isinstance(fixture, dict), f"{label}: fixture custody is missing")
    require(fixture.get("before") == planned_fixture and fixture.get("after") == planned_fixture,
            f"{label}: fixture custody differs")
    artifacts = receipt.get("artifacts")
    require(isinstance(artifacts, dict) and not receipt.get("artifact_inventory_error"),
            f"{label}: artifact inventory is missing or incomplete")
    for filename, digest in artifacts.items():
        require(isinstance(filename, str) and Path(filename).name == filename,
                f"{label}: artifact escapes packet: {filename}")
        path = HERE / filename
        require(path.is_file() and not path.is_symlink() and sha(path) == digest,
                f"{label}: artifact digest differs: {filename}")
    required = {
        job["name"] + suffix
        for suffix in (".json", ".stdout", ".stderr", ".callgrind",
                       ".source-before.json", ".source-after.json",
                       ".fixture-before.json", ".fixture-after.json")
    }
    require(required.issubset(artifacts),
            f"{label}: required artifacts omitted: {sorted(required - set(artifacts))}")
    numbered = sorted(
        (name for name in artifacts if name.startswith(job["name"] + ".callgrind.")
         and name.rsplit(".", 1)[-1].isdigit()),
        key=lambda name: int(name.rsplit(".", 1)[-1]),
    )
    require(numbered == [job["name"] + f".callgrind.{part}"
                         for part in range(1, len(numbered) + 1)],
            f"{label}: numbered Callgrind inventory is not contiguous")
    return {"source_custody": custody, "fixture": fixture, "artifacts": artifacts}


def stable(value: Any, path: str = "result") -> Any:
    """Remove sample/timing envelopes and collapse deterministic vectors."""

    if isinstance(value, dict):
        return {
            key: stable(child, f"{path}.{key}")
            for key, child in sorted(value.items())
            if key not in {"elapsed_ns", "operation_metrics"}
        }
    if isinstance(value, list):
        if not value:
            return []
        values = [stable(child, f"{path}[{index}]") for index, child in enumerate(value)]
        if all(child == values[0] for child in values):
            return [values[0]]
        return values
    return value


def validate_report(path: Path, job: dict[str, Any], build: dict[str, Any],
                    samples: int = 1, warmup: int = 0) -> tuple[dict[str, Any], Any]:
    label = relative(path)
    report = read_json(path)
    require(report.get("schema_version") == 1, f"{label}: schema differs")
    tool = report.get("tool")
    require(isinstance(tool, dict) and tool.get("binary") == "litchi-perf-baseline"
            and tool.get("profile") == "release"
            and tool.get("instrumentation") == "none",
            f"{label}: tool identity differs")
    identity = report.get("binary_identity")
    require(isinstance(identity, dict)
            and identity.get("binary_sha256") == build["binary_sha256"]
            and identity.get("binary_bytes") == build["binary_bytes"],
            f"{label}: binary identity differs")
    environment = report.get("environment")
    require(isinstance(environment, dict), f"{label}: environment is missing")
    plan = read_json(HERE / "plan.json")
    require(environment.get("git_revision") == plan["revision"],
            f"{label}: revision differs")
    configuration = report.get("configuration")
    require(isinstance(configuration, dict)
            and configuration.get("cases") == [job["case"]]
            and configuration.get("samples_per_case") == samples
            and configuration.get("warmup_iterations_per_case") == warmup,
            f"{label}: profile configuration differs")
    results = report.get("results")
    require(isinstance(results, list) and len(results) == 1,
            f"{label}: expected one result")
    result = results[0]
    require(isinstance(result, dict) and result.get("case") == job["case"],
            f"{label}: result case differs")
    source = result.get("source")
    ordinary = source.get("ordinary_save") if isinstance(source, dict) else None
    require(isinstance(ordinary, dict), f"{label}: ordinary_save evidence is missing")
    require(ordinary.get("format") == "DOCX"
            and ordinary.get("origin") == job["corpus"]["origin"]
            and ordinary.get("phase") == PHASE_LABELS[job["phase"]]
            and ordinary.get("timing_scope") == PHASE_TIMING_SCOPES[job["phase"]],
            f"{label}: ordinary-save phase identity differs")
    evidence = ordinary.get("corpus")
    manifest = result.get("corpus")
    require(isinstance(evidence, dict) and isinstance(manifest, dict),
            f"{label}: corpus evidence is missing")
    require(manifest.get("package_format") == "DOCX/OPC/ZIP"
            and manifest.get("archive_sha256") == evidence.get("source_archive_sha256")
            and manifest.get("archive_bytes") == evidence.get("source_archive_bytes")
            and manifest.get("archive_member_count") == evidence.get("source_member_count"),
            f"{label}: corpus manifest differs from source evidence")
    require(evidence.get("edit_admitted") is True
            and evidence.get("edit_outcome") == "admitted",
            f"{label}: selected corpus edit was not admitted")
    require(evidence.get("edit_description") ==
            'Package::document_mut().add_paragraph_with_text("litchi-perf-0638-ordinary-save")',
            f"{label}: edit description differs")
    require(evidence.get("repeated_cycles_identical") is True
            and evidence.get("repeated_saves_identical") is True
            and evidence.get("repeated_cycle_sha256") == evidence.get("published_sha256")
            and evidence.get("repeated_save_sha256") == evidence.get("published_sha256"),
            f"{label}: determinism proof failed")
    if job["corpus"]["origin"] == "caller-named-real-file":
        real_file = evidence.get("real_file")
        fixture = fixture_binding(job["corpus"])
        require(isinstance(real_file, dict) and fixture is not None
                and Path(str(real_file.get("path"))).resolve() == Path(fixture["resolved_path"])
                and real_file.get("sha256") == fixture["sha256"]
                and real_file.get("bytes") == fixture["bytes"],
                f"{label}: real-file identity differs")
    else:
        require(evidence.get("real_file") is None,
                f"{label}: generated corpus carries real-file provenance")
    outcome_digest = hashlib.sha256(b"admitted").hexdigest()
    outcome_hashes = ordinary.get("edit_outcome_sha256")
    require(isinstance(outcome_hashes, list) and len(outcome_hashes) == samples
            and all(item == outcome_digest for item in outcome_hashes),
            f"{label}: edit outcome vector differs")
    published = ordinary.get("published_sha256")
    if job["phase"] == "edit":
        require(published == [] and result.get("output_sha256") is None,
                f"{label}: edit unexpectedly publishes output")
        require(result.get("sink") is None, f"{label}: edit carries sink evidence")
    else:
        require(isinstance(published, list) and len(published) == samples
                and all(item == evidence.get("published_sha256") for item in published)
                and result.get("output_sha256") == evidence.get("published_sha256"),
                f"{label}: publication digest vector differs")
        sink = result.get("sink")
        require(isinstance(sink, dict)
                and sink.get("accepted_bytes") == evidence.get("byte_split", {}).get("output_total_bytes"),
                f"{label}: counting sink evidence differs")
    identity = stable(result)
    return report, identity


def callgrind_summary(text: str, label: str) -> int:
    return H.summary_ir(text, label)


def raw_edges(path: Path) -> list[dict[str, Any]]:
    """Parse positive ``cfn`` edges from a retained raw Callgrind part."""

    lines = read_text(path).splitlines()
    events: list[str] = []
    names: dict[int, str] = {}
    for line in lines:
        stripped = line.strip()
        if stripped.startswith("events:"):
            events = stripped.split(":", 1)[1].split()
        match = CALLGRIND_FUNCTION_RE.match(stripped)
        if match and match.group(3):
            names[int(match.group(2))] = match.group(3)
    require("Ir" in events, f"{relative(path)}: Ir event is missing")
    event_index = events.index("Ir")
    current: int | None = None
    callee: int | None = None
    calls: int | None = None
    result: list[dict[str, Any]] = []
    for line in lines:
        stripped = line.strip()
        match = CALLGRIND_FUNCTION_RE.match(stripped)
        if match:
            if match.group(1) == "fn":
                current = int(match.group(2))
                callee = None
                calls = None
            else:
                require(current is not None, f"{relative(path)}: cfn outside fn")
                callee = int(match.group(2))
                calls = None
            continue
        call_match = CALLGRIND_CALLS_RE.match(stripped)
        if call_match:
            require(current is not None and callee is not None,
                    f"{relative(path)}: calls outside cfn")
            calls = int(call_match.group(1).replace(",", ""))
            continue
        if current is None:
            continue
        cost = H.RAW._cost(stripped, event_index)
        if cost is None:
            continue
        if callee is not None and calls is not None and cost > 0 and calls > 0:
            result.append({
                "caller": names.get(current, ""),
                "callee": names.get(callee, ""),
                "calls": calls,
                "inclusive_ir": cost,
            })
            callee = None
            calls = None
    return result


def edge_summary(path: Path, predicate: Callable[[str], bool], target: str) -> dict[str, Any]:
    matches = [edge for edge in raw_edges(path) if predicate(edge["callee"])]
    by_caller: dict[str, dict[str, Any]] = {}
    for edge in matches:
        caller = edge["caller"]
        value = by_caller.setdefault(caller, {"caller": caller, "calls": 0,
                                              "inclusive_ir": 0, "edge_count": 0})
        value["calls"] += edge["calls"]
        value["inclusive_ir"] += edge["inclusive_ir"]
        value["edge_count"] += 1
    return {
        "target": target,
        "positive_edge_count": len(matches),
        "calls": sum(edge["calls"] for edge in matches),
        "inclusive_ir": sum(edge["inclusive_ir"] for edge in matches),
        "edges": matches,
        "by_caller": [by_caller[key] for key in sorted(by_caller)],
    }


def raw_function_names(path: Path) -> set[str]:
    names: set[str] = set()
    for line in read_text(path).splitlines():
        match = CALLGRIND_FUNCTION_RE.match(line.strip())
        if match and match.group(3):
            names.add(match.group(3))
    return names


def numbered_paths(stem: Path) -> list[Path]:
    paths = [path for path in HERE.glob(stem.name + ".[0-9]*")
             if path.is_file() and path.suffix[1:].isdigit()]
    paths.sort(key=lambda path: int(path.suffix[1:]))
    parts = [int(path.suffix[1:]) for path in paths]
    require(parts == list(range(1, len(parts) + 1)),
            f"{relative(stem)}: numbered parts are not contiguous: {parts}")
    require(paths, f"{relative(stem)}: no numbered Callgrind parts")
    return paths


def setup_caller_allowed(phase: str, caller: str) -> bool:
    if phase == "edit":
        return caller in {
            "litchi_perf_baseline::ordinary_save::build_corpus",
            "litchi_perf_baseline::ordinary_save::publish_reference",
        }
    # The serializer is passed through the save/atomic generic.  Depending on
    # which of the four exact monomorphs is entered, Callgrind can expose the
    # source call site, Owner::save, save_plain_impl, or atomic closure.
    return any(prefix in caller for prefix in (
        "litchi_perf_baseline::semantic_docx_bytes",
        "litchi_perf_baseline::ordinary_save::publish_reference",
        "litchi_perf_baseline::ordinary_save::Owner::save",
        "litchi_docx::package::codec::<impl ",
        "litchi_opc::atomic::replace_with",
    ))


def classify_parts(paths: list[Path], job: dict[str, Any], profile: dict[str, Any]) -> tuple[list[dict[str, Any]], Path, list[str]]:
    phase = job["phase"]
    owner = profile["owners"][phase]
    predicate = lambda name: name == owner
    rows: list[dict[str, Any]] = []
    measured: list[Path] = []
    warnings: list[str] = []
    for path in paths:
        label = relative(path)
        text = read_text(path)
        part = int(path.suffix[1:])
        require(H.part_number(text, label) == part,
                f"{label}: part number differs from suffix")
        require(H.trigger(text, label) == f"--dump-after={owner}",
                f"{label}: Trigger does not identify exact owner")
        require("events: Ir" in text.splitlines(), f"{label}: Ir event is missing")
        summary = callgrind_summary(text, label)
        incoming = edge_summary(path, predicate, owner)
        require(incoming["positive_edge_count"] > 0 and incoming["calls"] > 0,
                f"{label}: exact owner has no positive incoming edge")
        if incoming["inclusive_ir"] != summary:
            warnings.append(
                f"{label}: raw owner incoming Ir {incoming['inclusive_ir']} != dump summary {summary}"
            )
        measured_edges = [edge for edge in incoming["edges"]
                          if edge["caller"] == MEASURED_PARENT]
        measured_calls = sum(edge["calls"] for edge in measured_edges)
        is_measured = measured_calls > 0
        if is_measured:
            measured.append(path)
            role = "measured"
            if len(measured_edges) != 1:
                warnings.append(f"{label}: measured owner edge is split across raw records")
        else:
            role = "setup"
            for caller in sorted({edge["caller"] for edge in incoming["edges"]}):
                if not setup_caller_allowed(phase, caller):
                    warnings.append(f"{label}: setup owner caller not in policy: {caller}")
        nested = None
        if phase == "edit":
            nested = edge_summary(path, lambda name: name == DOCUMENT_MUT, DOCUMENT_MUT)
            if nested["positive_edge_count"] == 0:
                warnings.append(f"{label}: document_mut nested diagnostic has no positive edge")
        rows.append({
            "file": label,
            "part": part,
            "sha256": sha(path),
            "summary_ir": summary,
            "owner_incoming": incoming,
            "nested_document_mut": nested,
            "role": role,
            "selection_edge": {
                "caller": MEASURED_PARENT if is_measured else "setup",
                "calls": measured_calls if is_measured else incoming["calls"],
            },
        })
    expected = profile["expected_numbered_dumps"][phase][job["corpus_id"]]
    if len(paths) != expected:
        warnings.append(f"{job['name']}: expected {expected} numbered dumps, retained {len(paths)}")
    require(len(measured) == 1,
            f"{job['name']}: expected one measured run_case dump, got {len(measured)}")
    setup_count = len(paths) - len(measured)
    require(setup_count >= expected - 1,
            f"{job['name']}: too few setup dumps: {setup_count}")
    if phase == "edit":
        caller_counts: dict[str, int] = {}
        for row in rows:
            if row["role"] == "setup":
                for edge in row["owner_incoming"]["edges"]:
                    caller_counts[edge["caller"]] = caller_counts.get(edge["caller"], 0) + 1
        expected_callers = profile["setup_parents"]["edit"]
        for caller, count in expected_callers.items():
            require(caller_counts.get(caller, 0) == count,
                    f"{job['name']}: setup caller {caller} count {caller_counts.get(caller, 0)} != {count}")
    else:
        expected_callers = profile["setup_parents"]["counting_publish"][job["corpus_id"]]
        caller_counts: dict[str, int] = {}
        for row in rows:
            if row["role"] != "setup":
                continue
            for edge in row["owner_incoming"]["edges"]:
                caller_counts[edge["caller"]] = caller_counts.get(edge["caller"], 0) + 1
        for caller, count in expected_callers.items():
            if caller_counts.get(caller, 0) != count:
                warnings.append(
                    f"{job['name']}: calibration setup caller {caller} count "
                    f"{caller_counts.get(caller, 0)} != expected {count}"
                )
    return rows, measured[0], warnings


def validate_terminal(stem: Path, numbered: list[Path]) -> dict[str, Any]:
    require(stem.is_file() and not stem.is_symlink(),
            f"{relative(stem)}: Program termination part is missing")
    text = read_text(stem)
    part = H.part_number(text, relative(stem))
    require(part == int(numbered[-1].suffix[1:]) + 1,
            f"{relative(stem)}: termination part does not follow numbered parts")
    require(H.trigger(text, relative(stem)) == "Program termination",
            f"{relative(stem)}: termination Trigger differs")
    summary = callgrind_summary(text, relative(stem))
    require(summary == 0, f"{relative(stem)}: termination Ir is {summary}, not zero")
    require("events: Ir" in text.splitlines(), f"{relative(stem)}: termination Ir is missing")
    return {"file": relative(stem), "part": part, "sha256": sha(stem),
            "summary_ir": summary, "trigger": "Program termination"}


def write_or_check(path: Path, text: str) -> None:
    if path.exists():
        require(read_text(path) == text, f"annotation replay differs: {relative(path)}")
    else:
        path.write_text(text, encoding="utf-8")


def annotate(selected: Path, name: str, owner: str, summary: int) -> dict[str, Any]:
    try:
        inclusive, inclusive_command = H.run_annotation(selected, True)
        exclusive, self_command = H.run_annotation(selected, False)
    except Exception as error:
        fail(f"{relative(selected)}: callgrind_annotate failed: {error}")
    inclusive_path = HERE / f"{name}.inclusive.txt"
    self_path = HERE / f"{name}.self.txt"
    write_or_check(inclusive_path, inclusive)
    write_or_check(self_path, exclusive)
    parsed_inclusive = H.parse_annotation(inclusive, owner, relative(inclusive_path))
    parsed_self = H.parse_annotation(exclusive, owner, relative(self_path))
    require(parsed_inclusive["selected_ir"] == summary,
            f"{name}: inclusive annotation Ir differs from raw summary")
    require(parsed_self["selected_ir"] <= parsed_inclusive["selected_ir"],
            f"{name}: self annotation exceeds inclusive owner")
    require(parsed_self["direct"] == parsed_inclusive["direct"],
            f"{name}: inclusive/self direct children differ")
    direct_ir = sum(item["inclusive_ir"] for item in parsed_inclusive["direct"])
    partition_ok = parsed_self["selected_ir"] + direct_ir == parsed_inclusive["selected_ir"]
    require(partition_ok,
            f"{name}: self plus immediate children does not reconstruct owner Ir")
    return {
        "environment": dict(PERL_ENV),
        "command": {"inclusive": inclusive_command, "self": self_command},
        "files": {
            "inclusive": relative(inclusive_path),
            "inclusive_sha256": sha(inclusive_path),
            "self": relative(self_path),
            "self_sha256": sha(self_path),
        },
        "owner": {
            "name": owner,
            "inclusive_ir": parsed_inclusive["selected_ir"],
            "self_ir": parsed_self["selected_ir"],
            "direct_callees": parsed_inclusive["direct"],
            "direct_callee_ir": H.direct_map(parsed_inclusive["direct"]),
        },
        "immediate_child_partition": {
            "self_ir": parsed_self["selected_ir"],
            "direct_children_ir": direct_ir,
            "owner_ir": parsed_inclusive["selected_ir"],
            "disjoint": True,
            "equation": "self_ir + sum(immediate direct-child inclusive Ir) = owner inclusive Ir",
            "nested_inclusive_costs_excluded": True,
        },
        "validation": {
            "inclusive_owner_matches_raw": True,
            "self_plus_immediate_children_equals_inclusive": True,
            "inclusive_and_self_direct_children_match": True,
        },
    }


def analyze_job(job: dict[str, Any], plan: dict[str, Any], profile: dict[str, Any],
                build: dict[str, Any], plan_sha: str, profile_sha: str,
                constraints_sha: str, script_sha: str) -> tuple[dict[str, Any], list[str]]:
    name = job["name"]
    receipt_path = HERE / f"{name}.receipt.json"
    report_path = HERE / f"{name}.json"
    require(receipt_path.is_file(), f"{relative(receipt_path)}: receipt is missing")
    require(report_path.is_file(), f"{relative(report_path)}: profile report is missing")
    receipt = read_json(receipt_path)
    custody = validate_receipt(job, receipt, plan, profile, build,
                               plan_sha, profile_sha, constraints_sha, script_sha)
    report, identity = validate_report(report_path, job, build)
    stem = HERE / f"{name}.callgrind"
    numbered = numbered_paths(stem)
    parts, selected, warnings = classify_parts(numbered, job, profile)
    terminal = validate_terminal(stem, numbered)
    selected_row = next(row for row in parts if row["file"] == relative(selected))
    annotations = annotate(selected, name, profile["owners"][job["phase"]],
                            selected_row["summary_ir"])
    profile_row = {
        "name": name,
        "repeat": job["repeat"],
        "corpus_id": job["corpus_id"],
        "phase": job["phase"],
        "order_index": job["order_index"],
        "receipt": relative(receipt_path),
        "receipt_sha256": sha(receipt_path),
        "result": relative(report_path),
        "result_sha256": sha(report_path),
        "result_identity": identity,
        "source_custody": custody["source_custody"],
        "fixture": custody["fixture"],
        "parts": parts,
        "termination": terminal,
        "selected": relative(selected),
        "selected_by": {
            "caller": MEASURED_PARENT,
            "positive_owner_call_count": selected_row["selection_edge"]["calls"],
            "selection_is_raw_incoming_edge_based": True,
        },
        "annotations": annotations,
        "total_ir": annotations["owner"]["inclusive_ir"],
        "self_ir": annotations["owner"]["self_ir"],
        "direct_ir": annotations["immediate_child_partition"]["direct_children_ir"],
        "direct_callee_ir": annotations["owner"]["direct_callee_ir"],
        "warnings": warnings,
        "limitations": [
            "The setup parts are retained as lifecycle evidence and excluded from the measured Ir row.",
            "Valgrind allocation edges, if present, are not converted into allocation counts.",
        ],
    }
    return profile_row, warnings


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, default=HERE / "profile-analysis.json")
    args = parser.parse_args()
    try:
        plan = load_plan()
        profile = load_profile_plan(plan)
        expected_source = read_json(HERE / "source-baseline.json")
        require(expected_source == source_census(),
                "current source differs from retained baseline")
        constraints_path = HERE / "constraints.json"
        constraints = read_json(constraints_path)
        require(isinstance(constraints, dict), "constraints.json is not an object")
        for name, digest in constraints.items():
            path = REPO / name
            require(path.is_file() and sha(path) == digest,
                    f"constraint changed: {name}")
        build = build_info()
        plan_sha = sha(HERE / "plan.json")
        profile_sha = sha(HERE / "profile-plan.json")
        constraints_sha = sha(constraints_path)
        profile_script = HERE / "profile.py"
        require(profile_script.is_file(), "profile.py is missing")
        script_sha = sha(profile_script)
        rows: list[dict[str, Any]] = []
        warnings: list[str] = []
        native_identity: dict[tuple[int, str, str], Any] = {}
        for job in expected_jobs(plan, profile):
            row, row_warnings = analyze_job(job, plan, profile, build,
                                            plan_sha, profile_sha,
                                            constraints_sha, script_sha)
            rows.append(row)
            warnings.extend(row_warnings)
            native_path = HERE / f"native-r{job['repeat']}-{job['corpus_id']}-{job['phase']}.json"
            if native_path.is_file():
                native_job = dict(job)
                native_job["case"] = job["case"]
                _, native_value = validate_report(native_path, native_job, build,
                                                  samples=100, warmup=10)
                native_identity[(job["repeat"], job["corpus_id"], job["phase"])] = native_value
                require(native_value == row["result_identity"],
                        f"{job['name']}: profile/native stable result identity differs")
            else:
                warnings.append(f"{job['name']}: native parity report is not present")
        require(len(rows) == 8, f"profile matrix is incomplete: {len(rows)}")
        aggregate: dict[str, Any] = {}
        for phase in PHASES:
            for corpus in CORPORA:
                selected = [row for row in rows
                            if row["phase"] == phase and row["corpus_id"] == corpus]
                total = sum(row["total_ir"] for row in selected)
                aggregate[f"{corpus}:{phase}"] = {
                    "profile_count": len(selected),
                    "owner_total_ir": total,
                    "owner_self_ir": sum(row["self_ir"] for row in selected),
                    "immediate_child_partition_ir": sum(row["direct_ir"] for row in selected),
                    "partition_equation": "aggregate self Ir + aggregate immediate-child Ir = aggregate owner Ir",
                    "nested_diagnostics_added": False,
                }
        output = {
            "schema": "docx_callgrind_ordinary_save_profile_analysis_v1",
            "status": "pass",
            "packet": "change-0709-docx-ordinary-save",
            "revision": plan["revision"],
            "owners": {
                "edit": EDIT_OWNER,
                "counting_publish": WRITE_PLAIN_OWNER,
                "measured_parent": MEASURED_PARENT,
                "nested_edit_diagnostic": DOCUMENT_MUT,
            },
            "plan": relative(HERE / "plan.json"),
            "plan_sha256": plan_sha,
            "profile_plan": relative(HERE / "profile-plan.json"),
            "profile_plan_sha256": profile_sha,
            "profile_script": relative(HERE / "profile.py"),
            "profile_script_sha256": script_sha,
            "analyzer_script": relative(HERE / "analyze_profiles.py"),
            "analyzer_script_sha256": sha(Path(__file__).resolve()),
            "build": {key: value for key, value in build.items() if key not in {"source"}},
            "helpers": retained_helper_identities(),
            "profiles": rows,
            "aggregate": aggregate,
            "warnings": sorted(set(warnings)),
            "validation": {
                "profile_matrix": len(rows) == 8,
                "binary_and_nm_bindings": True,
                "source_and_fixture_custody": True,
                "all_numbered_raw_parts_retained": True,
                "termination_parts_zero_ir": True,
                "measured_selected_by_raw_run_case_edge": True,
                "setup_parts_preserved": True,
                "immediate_children_disjoint_partition": True,
                "nested_document_mut_diagnostic_excluded_from_partition": True,
                "allocation_counts_inferred": False,
                "native_latency_claim": False,
            },
            "limitations": [
                "Callgrind Ir is guest-instruction attribution, not native latency, hardware instructions, cycles, allocation counts, RSS, or cache counters.",
                "The edit owner covers the direct Owner::edit call, including document_mut admission; setup probes and reference publications are separate retained parts.",
                "The publication owner covers exact write_plain symbols; write_plain closure helpers are outside the toggle pattern, and their nested inclusive costs are represented only inside the owner total.",
                "Immediate direct children are the only disjoint partition. Nested document_mut and serializer helper rows are overlapping diagnostics and are never added to owner totals.",
            ],
        }
        write_json(args.output, output)
        print(f"verified {len(rows)} DOCX Callgrind profiles; wrote {args.output}")
        return 0
    except (EvidenceError, OSError, KeyError, ValueError) as error:
        print(f"analyze_profiles.py: error: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
