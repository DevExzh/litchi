#!/usr/bin/env python3
"""Capture the 0778 ordinary-save durability policy matrix.

The coordinator runs this file in three separate lanes::

    python3 -B capture.py qualification
    python3 -B capture.py native
    python3 -B capture.py allocation

Qualification is intentionally a one-sample, zero-warmup run.  It freezes the
complete harness corpus identity (including the typed edit outcome) before the
timed matrix is allowed to start.  ``native`` and ``allocation`` refuse to run
until the independent export oracle has produced ``admission.json``.  Every
child is a fresh process; every output, log, RSS receipt, input hash and
binary hash is retained.

This runner owns no production code and never invokes Cargo.  The root
coordinator builds the binaries and runs these lanes serially.
"""

from __future__ import annotations

import argparse
import copy
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import subprocess
import sys
import time
from typing import Any


PACKET = Path(__file__).resolve().parent
ROOT = PACKET.parents[3]
# The packet itself is the capture root.  The lane directories are numbered
# so a failed lane is retained and a later invocation cannot silently replace
# it; the diagnostics runner consumes ``native-0/complete.json`` directly.
CAPTURE = PACKET
PLAN_PATH = PACKET / "plan.json"
SOURCE_PATH = PACKET / "source.json"
BUILD_PATH = PACKET / "build.json"
ADMISSION_PATH = PACKET / "admission.json"
RUNNER_PATH = Path(__file__).resolve()

_custody_spec = importlib.util.spec_from_file_location("custody0778", PACKET / "custody.py")
if _custody_spec is None or _custody_spec.loader is None:
    raise RuntimeError("cannot load packet custody.py")
C = importlib.util.module_from_spec(_custody_spec)
_custody_spec.loader.exec_module(C)


PHASE_NAMES = {
    "lifecycle": "open+edit+save",
    "atomic_publish": "save-to-path",
    "edit": "edit",
    "counting_publish": "serialize-to-counting-sink",
}
REAL_FILE_ORIGIN = "caller-named-real-file"
GENERATED_ORIGIN = "generated-harness-corpus"
TIMED_PHASES = {"lifecycle", "atomic_publish"}


def read(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def write(path: Path, value: Any, *, replace: bool = False) -> None:
    if path.exists() and not replace:
        raise RuntimeError(f"refusing to overwrite existing artifact: {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n", encoding="utf-8")


def sha(path: Path) -> str:
    return C.sha(path)


def canonical_sha(value: Any) -> str:
    encoded = json.dumps(value, sort_keys=True, separators=(",", ":")).encode("utf-8")
    return hashlib.sha256(encoded).hexdigest()


def packet_path(path: Path) -> str:
    try:
        return str(path.relative_to(PACKET))
    except ValueError:
        return str(path)


def artifact(path: Path) -> dict[str, Any]:
    if not path.is_file():
        raise RuntimeError(f"missing artifact: {path}")
    return {"path": packet_path(path), "bytes": path.stat().st_size, "sha256": sha(path)}


def git_head() -> str:
    return subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=ROOT, text=True).strip()


def git_is_ancestor(ancestor: str, descendant: str) -> bool:
    """Return whether the frozen integration base is in the source history."""
    if not isinstance(ancestor, str) or not isinstance(descendant, str):
        return False
    result = subprocess.run(
        ["git", "merge-base", "--is-ancestor", ancestor, descendant],
        cwd=ROOT,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )
    return result.returncode == 0


def load_inputs() -> tuple[dict[str, Any], dict[str, Any], dict[str, Any]]:
    for path in (PLAN_PATH, SOURCE_PATH, BUILD_PATH):
        if not path.is_file():
            raise RuntimeError(f"root must write {path.name} before capture")
    plan, source, build = read(PLAN_PATH), read(SOURCE_PATH), read(BUILD_PATH)
    if plan.get("schema") != "litchi-0778-durability-plan-v1":
        raise RuntimeError(f"unexpected plan schema: {plan.get('schema')!r}")
    if Path(plan["root"]).resolve() != ROOT:
        raise RuntimeError(f"plan root is not this worktree: {plan.get('root')!r}")
    if Path(plan["target"]).resolve() == ROOT:
        raise RuntimeError("release target must be outside the source worktree")
    head = git_head()
    base = plan.get("base")
    if not git_is_ancestor(base, head):
        raise RuntimeError("plan base is not an ancestor of the frozen source revision")
    if source.get("revision") != head:
        raise RuntimeError("source revision differs from current HEAD")
    if not isinstance(source.get("files"), dict) or not source["files"]:
        raise RuntimeError("source.json has no file census")
    if not isinstance(build.get("binaries"), dict):
        raise RuntimeError("build.json has no binary map")
    for name in ("native", "export", "allocation"):
        if name not in build["binaries"]:
            raise RuntimeError(f"build.json is missing {name} binary")
    validate_plan(plan)
    return plan, source, build


def validate_plan(plan: dict[str, Any]) -> None:
    corpora = plan.get("corpora")
    policies = plan.get("policies")
    phases = plan.get("phases")
    controls = plan.get("controls")
    if not isinstance(corpora, list) or len(corpora) != 7:
        raise RuntimeError("0778 requires exactly seven corpora")
    if policies != ["default", "full", "file-only", "no-sync"]:
        raise RuntimeError(f"unexpected policy order: {policies!r}")
    if phases != ["lifecycle", "atomic_publish"]:
        raise RuntimeError(f"unexpected timed phase order: {phases!r}")
    if controls != ["edit", "counting_publish"]:
        raise RuntimeError(f"unexpected control order: {controls!r}")
    ids: set[str] = set()
    formats = {"docx", "xlsx", "pptx"}
    for corpus in corpora:
        if not isinstance(corpus, dict):
            raise RuntimeError("malformed corpus row")
        ident = corpus.get("id")
        if not isinstance(ident, str) or ident in ids:
            raise RuntimeError(f"duplicate or missing corpus id: {ident!r}")
        ids.add(ident)
        if corpus.get("format") not in formats:
            raise RuntimeError(f"unsupported corpus format: {corpus.get('format')!r}")
        path = corpus.get("path")
        if path is None:
            if corpus.get("expected_edit_admitted") is not True:
                raise RuntimeError(f"generated corpus is not admitted: {ident}")
        else:
            if not isinstance(path, str) or not path or Path(path).is_absolute():
                raise RuntimeError(f"corpus path must be repository-relative: {ident}")
            if not isinstance(corpus.get("bytes"), int) or not isinstance(corpus.get("sha256"), str):
                raise RuntimeError(f"real corpus lacks a frozen byte identity: {ident}")
    orders = plan.get("policy_orders")
    if orders != [
        ["default", "full", "no-sync", "file-only"],
        ["full", "file-only", "default", "no-sync"],
        ["file-only", "no-sync", "full", "default"],
        ["no-sync", "default", "file-only", "full"],
    ]:
        raise RuntimeError("policy orders are not the frozen Williams matrix")
    native = plan.get("native", {})
    qualification = plan.get("qualification", {})
    allocation = plan.get("allocation", {})
    if native.get("blocks") != 4 or native.get("samples") != 100 or native.get("warmup") != 10:
        raise RuntimeError("unexpected native settings")
    if qualification.get("samples") != 1 or qualification.get("warmup") != 0:
        raise RuntimeError("unexpected qualification settings")
    if allocation.get("blocks") != 2 or allocation.get("samples") != 3 or allocation.get("warmup") != 0:
        raise RuntimeError("unexpected allocation settings")
    expected = plan.get("expected_children", {})
    if expected != {"allocation": 140, "native": 280, "qualification": 28}:
        raise RuntimeError(f"unexpected expected child counts: {expected!r}")
    if plan.get("cpu") != 12:
        raise RuntimeError("0778 CPU pin changed")
    if not Path("/usr/bin/time").is_file():
        raise RuntimeError("/usr/bin/time is required for process RSS receipts")


def source_guard(source: dict[str, Any]) -> dict[str, Any]:
    current = C.census()
    if current != source:
        raise RuntimeError("source census or HEAD changed during capture")
    return current


def binary_receipts(build: dict[str, Any]) -> dict[str, dict[str, Any]]:
    result: dict[str, dict[str, Any]] = {}
    for name, row in build["binaries"].items():
        if not isinstance(row, dict) or not isinstance(row.get("path"), str):
            raise RuntimeError(f"malformed {name} binary receipt")
        path = Path(row["path"])
        if not path.is_absolute():
            path = ROOT / path
        actual = C.artifact(path)
        expected = {"path": str(path), "bytes": row.get("bytes"), "sha256": row.get("sha256")}
        if actual != expected:
            raise RuntimeError(f"{name} binary does not match build.json: {actual!r} != {expected!r}")
        result[name] = {**actual, "path": str(path)}
    return result


def fixture_inventory(plan: dict[str, Any]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for corpus in plan["corpora"]:
        path_value = corpus.get("path")
        if path_value is None:
            result[corpus["id"]] = None
            continue
        path = ROOT / path_value
        if path.is_symlink() or not path.is_file():
            raise RuntimeError(f"fixture is not a regular file: {path}")
        actual = {"path": path_value, "bytes": path.stat().st_size, "sha256": sha(path)}
        expected = {"path": path_value, "bytes": corpus.get("bytes"), "sha256": corpus.get("sha256")}
        if actual != expected:
            raise RuntimeError(f"fixture identity differs for {corpus['id']}: {actual!r} != {expected!r}")
        result[corpus["id"]] = actual
    return result


def fixture_guard(plan: dict[str, Any], expected: dict[str, Any]) -> None:
    actual = fixture_inventory(plan)
    if actual != expected:
        raise RuntimeError("fixture bytes changed during capture")


def runner_receipt() -> dict[str, Any]:
    return {"path": str(RUNNER_PATH), "bytes": RUNNER_PATH.stat().st_size, "sha256": sha(RUNNER_PATH)}


def freeze_inputs(
    plan: dict[str, Any], source: dict[str, Any], build: dict[str, Any]
) -> dict[str, Any]:
    return {
        "schema": "litchi-0778-durability-freeze-v1",
        "plan": artifact(PLAN_PATH),
        "source": artifact(SOURCE_PATH),
        "build": artifact(BUILD_PATH),
        "runner": runner_receipt(),
        "source_revision": source["revision"],
        "build_source_sha256": build.get("source_sha256"),
        "binaries": binary_receipts(build),
        "fixtures": fixture_inventory(plan),
    }


def verify_frozen(
    plan: dict[str, Any], source: dict[str, Any], build: dict[str, Any], freeze: dict[str, Any]
) -> dict[str, Any]:
    for path, expected in ((PLAN_PATH, freeze["plan"]), (SOURCE_PATH, freeze["source"]), (BUILD_PATH, freeze["build"])):
        actual = artifact(path)
        if actual["bytes"] != expected["bytes"] or actual["sha256"] != expected["sha256"]:
            raise RuntimeError(f"frozen input changed: {path.name}")
    runner = runner_receipt()
    if runner != freeze["runner"]:
        raise RuntimeError("capture runner changed after freeze")
    if source_guard(source) != source:
        raise RuntimeError("source guard failed")
    if build.get("source_sha256") != sha(SOURCE_PATH):
        raise RuntimeError("build.json is not bound to source.json")
    binaries = binary_receipts(build)
    if binaries != freeze["binaries"]:
        raise RuntimeError("a frozen binary changed during capture")
    fixture_guard(plan, freeze["fixtures"])
    return binaries


def corpus_by_id(plan: dict[str, Any]) -> dict[str, dict[str, Any]]:
    return {row["id"]: row for row in plan["corpora"]}


def case_name(corpus: dict[str, Any], phase: str) -> str:
    prefix = f"{corpus['format']}_ordinary_save_" if corpus.get("path") is None else f"{corpus['format']}_real_file_ordinary_save_"
    return prefix + phase


def jobs_for_lane(plan: dict[str, Any], lane: str) -> list[dict[str, Any]]:
    corpora = plan["corpora"]
    policies = plan["policies"]
    jobs: list[dict[str, Any]] = []
    if lane == "qualification":
        for corpus in corpora:
            for policy in policies:
                jobs.append({
                    "lane": lane,
                    "corpus_id": corpus["id"],
                    "phase": "atomic_publish",
                    "policy": policy,
                    "block": None,
                    "repeat": None,
                    "case": case_name(corpus, "atomic_publish"),
                    "samples": plan["qualification"]["samples"],
                    "warmup": plan["qualification"]["warmup"],
                    "order_index": len(jobs),
                })
        return jobs
    if lane not in {"native", "allocation"}:
        raise RuntimeError(f"unknown lane: {lane}")
    blocks = plan[lane]["blocks"]
    for block in range(blocks):
        order = plan["policy_orders"][block]
        for corpus in corpora:
            for phase in plan["phases"]:
                for policy in order:
                    jobs.append({
                        "lane": lane,
                        "corpus_id": corpus["id"],
                        "phase": phase,
                        "policy": policy,
                        "block": block,
                        "repeat": None if lane == "native" else block,
                        "case": case_name(corpus, phase),
                        "samples": plan[lane]["samples"],
                        "warmup": plan[lane]["warmup"],
                        "order_index": len(jobs),
                    })
            for phase in plan["controls"]:
                jobs.append({
                    "lane": lane,
                    "corpus_id": corpus["id"],
                    "phase": phase,
                    "policy": "default",
                    "block": block,
                    "repeat": None if lane == "native" else block,
                    "case": case_name(corpus, phase),
                    "samples": plan[lane]["samples"],
                    "warmup": plan[lane]["warmup"],
                    "order_index": len(jobs),
                })
    expected = plan["expected_children"][lane]
    if len(jobs) != expected:
        raise RuntimeError(f"{lane} schedule has {len(jobs)} jobs, expected {expected}")
    return jobs


def job_stem(job: dict[str, Any]) -> str:
    corpus = job["corpus_id"]
    phase = job["phase"]
    policy = job["policy"]
    if job["lane"] == "qualification":
        return f"qualification-{corpus}-{policy}"
    if job["lane"] == "native":
        return f"native-b{job['block']}-{corpus}-{phase}-{policy}"
    return f"allocation-r{job['repeat']}-{corpus}-{phase}-{policy}"


def environment(cpu: int) -> dict[str, str]:
    allowed = sorted(os.sched_getaffinity(0)) if hasattr(os, "sched_getaffinity") else []
    if allowed and cpu not in allowed:
        raise RuntimeError(f"CPU {cpu} is not available in current affinity: {allowed}")
    env = os.environ.copy()
    env.update({"LC_ALL": "C", "LANG": "C", "TZ": "UTC", "PYTHONDONTWRITEBYTECODE": "1"})
    return env


def binary_path(binary: dict[str, Any], name: str) -> Path:
    path = Path(binary[name]["path"])
    if not path.is_absolute():
        path = ROOT / path
    return path


def command_for(
    plan: dict[str, Any], job: dict[str, Any], binary: Path, report: Path, rss: Path
) -> list[str]:
    corpus = corpus_by_id(plan)[job["corpus_id"]]
    command = [
        "/usr/bin/time",
        "-f",
        "%M",
        "-o",
        str(rss),
        "taskset",
        "-c",
        str(plan["cpu"]),
        str(binary),
        "--samples",
        str(job["samples"]),
        "--warmup",
        str(job["warmup"]),
        "--case",
        job["case"],
        "--json",
        str(report),
        "--filesystem-root",
        str(plan["filesystem_root"]),
    ]
    if corpus.get("path") is not None:
        command.extend(["--ooxml-file", corpus["path"]])
    if job["phase"] in TIMED_PHASES and job["policy"] != "default":
        command.extend(["--save-durability", job["policy"]])
    return command


def report_identity(value: dict[str, Any]) -> dict[str, Any]:
    results = value.get("results")
    if not isinstance(results, list) or len(results) != 1 or not isinstance(results[0], dict):
        raise RuntimeError("ordinary-save child did not return exactly one result")
    result = results[0]
    source = result.get("source")
    if not isinstance(source, dict) or not isinstance(source.get("ordinary_save"), dict):
        raise RuntimeError("ordinary-save child omitted source.ordinary_save")
    ordinary = source["ordinary_save"]
    corpus = ordinary.get("corpus")
    if not isinstance(corpus, dict) or not isinstance(result.get("corpus"), dict):
        raise RuntimeError("ordinary-save child omitted corpus identity")
    return {"result_corpus": result["corpus"], "ordinary_corpus": corpus}


def validate_report(
    plan: dict[str, Any],
    job: dict[str, Any],
    report: Path,
    frozen_identity: dict[str, Any] | None,
    binary_receipt: dict[str, Any],
) -> tuple[dict[str, Any], dict[str, Any]]:
    value = read(report)
    config = value.get("configuration")
    if not isinstance(config, dict):
        raise RuntimeError(f"missing configuration in {report}")
    if config.get("samples_per_case") != job["samples"] or config.get("warmup_iterations_per_case") != job["warmup"]:
        raise RuntimeError(f"wrong sample settings in {report}")
    if config.get("cases") != [job["case"]]:
        raise RuntimeError(f"wrong case configuration in {report}: {config.get('cases')!r}")
    if config.get("filesystem_root_selected") is not True:
        raise RuntimeError(f"filesystem root was not selected in {report}")
    results = value.get("results")
    if not isinstance(results, list) or len(results) != 1:
        raise RuntimeError(f"wrong result cardinality in {report}")
    result = results[0]
    if result.get("case") != job["case"]:
        raise RuntimeError(f"wrong result case in {report}")
    elapsed = result.get("elapsed_ns")
    if not isinstance(elapsed, dict) or elapsed.get("unit") != "ns":
        raise RuntimeError(f"missing elapsed_ns in {report}")
    samples = elapsed.get("samples")
    order = elapsed.get("sample_order")
    if not isinstance(samples, list) or len(samples) != job["samples"]:
        raise RuntimeError(f"wrong elapsed sample cardinality in {report}")
    if not isinstance(order, list) or sorted(order) != list(range(job["samples"])):
        raise RuntimeError(f"elapsed sample order is not a permutation in {report}")
    ordinary = result["source"]["ordinary_save"]
    expected_phase = PHASE_NAMES[job["phase"]]
    if ordinary.get("phase") != expected_phase:
        raise RuntimeError(f"wrong ordinary-save phase in {report}: {ordinary.get('phase')!r}")
    selected_policy = ordinary.get("save_durability")
    expected_policy = None if job["policy"] == "default" else job["policy"]
    if selected_policy != expected_policy:
        raise RuntimeError(f"wrong durability policy in {report}: {selected_policy!r} != {expected_policy!r}")
    identity = report_identity(value)
    ordinary_corpus = identity["ordinary_corpus"]
    corpus = corpus_by_id(plan)[job["corpus_id"]]
    if ordinary_corpus.get("format") != corpus["format"].upper():
        raise RuntimeError(f"wrong format identity in {report}")
    expected_admitted = corpus["expected_edit_admitted"]
    if ordinary_corpus.get("edit_admitted") is not expected_admitted:
        raise RuntimeError(f"edit admission differs from plan in {report}")
    outcome = ordinary_corpus.get("edit_outcome")
    if not isinstance(outcome, str) or (expected_admitted and outcome != "admitted") or (not expected_admitted and not outcome.startswith("refused:")):
        raise RuntimeError(f"unexpected edit outcome in {report}: {outcome!r}")
    if corpus.get("path") is None:
        if identity["result_corpus"].get("shape") != "medium":
            raise RuntimeError(f"generated corpus is not fixed medium in {report}")
        if ordinary_corpus.get("origin") != GENERATED_ORIGIN:
            raise RuntimeError(f"generated corpus origin drifted in {report}")
    else:
        real_file = ordinary_corpus.get("real_file")
        expected_file = {"path": corpus["path"], "bytes": corpus["bytes"], "sha256": corpus["sha256"]}
        if real_file != expected_file:
            raise RuntimeError(f"real-file identity differs in {report}: {real_file!r} != {expected_file!r}")
        if ordinary_corpus.get("origin") != REAL_FILE_ORIGIN:
            raise RuntimeError(f"real-file corpus origin drifted in {report}")
    source_archive_sha256 = ordinary_corpus.get("source_archive_sha256")
    if not isinstance(source_archive_sha256, str) or len(source_archive_sha256) != 64:
        raise RuntimeError(f"missing source archive digest in {report}")
    if corpus.get("path") is not None and source_archive_sha256 != corpus["sha256"]:
        raise RuntimeError(f"real-file source digest differs from the fixture in {report}")
    if corpus.get("path") is not None and ordinary_corpus.get("source_archive_bytes") != corpus["bytes"]:
        raise RuntimeError(f"real-file source byte length differs from the fixture in {report}")
    result_corpus = identity["result_corpus"]
    outer_archive_sha256 = result_corpus.get("archive_sha256")
    if isinstance(outer_archive_sha256, str) and outer_archive_sha256 != source_archive_sha256:
        raise RuntimeError(f"outer and ordinary corpus source digests differ in {report}")
    publication_identity = ordinary_corpus.get("published_sha256")
    if not isinstance(publication_identity, str) or len(publication_identity) != 64:
        raise RuntimeError(f"missing frozen publication digest in {report}")
    if ordinary_corpus.get("repeated_cycle_sha256") != publication_identity or ordinary_corpus.get("repeated_save_sha256") != publication_identity:
        raise RuntimeError(f"reference publication determinism is false in {report}")
    if ordinary_corpus.get("repeated_cycles_identical") is not True or ordinary_corpus.get("repeated_saves_identical") is not True:
        raise RuntimeError(f"reference publication identity is not deterministic in {report}")
    if ordinary.get("publications_identical") is not True or ordinary.get("edit_outcomes_identical") is not True:
        raise RuntimeError(f"sample identity gate failed in {report}")
    outcome_digest = sha_bytes(outcome.encode("utf-8"))
    outcome_rows = ordinary.get("edit_outcome_sha256")
    if not isinstance(outcome_rows, list) or len(outcome_rows) != job["samples"] or any(row != outcome_digest for row in outcome_rows):
        raise RuntimeError(f"edit outcome digest sequence differs in {report}")
    published_rows = ordinary.get("published_sha256")
    if job["phase"] == "edit":
        if published_rows != []:
            raise RuntimeError(f"edit phase unexpectedly published bytes in {report}")
    else:
        if not isinstance(published_rows, list) or len(published_rows) != job["samples"] or any(row != publication_identity for row in published_rows):
            raise RuntimeError(f"published digest sequence differs in {report}")
    binary_identity = value.get("binary_identity")
    if not isinstance(binary_identity, dict) or binary_identity.get("binary_sha256") != binary_receipt["sha256"] or binary_identity.get("binary_bytes") != binary_receipt["bytes"]:
        raise RuntimeError(f"report binary identity differs in {report}")
    if job["lane"] == "allocation":
        allocation = result.get("operation_metrics", {}).get("allocation")
        if not isinstance(allocation, dict) or allocation.get("status") != "measured":
            raise RuntimeError(f"allocation metrics are not measured in {report}")
    if frozen_identity is not None and identity != frozen_identity:
        raise RuntimeError(f"corpus identity changed after qualification in {report}")
    return identity, {
        "publication_sha256": published_rows,
        "edit_outcome_sha256": outcome_rows,
        "edit_admitted": expected_admitted,
        "edit_outcome": outcome,
        "source_archive_sha256": source_archive_sha256,
    }


def sha_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def report_path_for(lane_dir: Path, stem: str, suffix: str) -> Path:
    return lane_dir / f"{stem}.{suffix}"


def is_sha256(value: Any) -> bool:
    return isinstance(value, str) and len(value) == 64 and all(
        character in "0123456789abcdefABCDEF" for character in value
    )


def packet_file(value: Any, label: str) -> Path:
    if not isinstance(value, str) or not value:
        raise RuntimeError(f"{label} must name a file")
    path = Path(value)
    if not path.is_absolute():
        path = PACKET / path
    if not path.is_file() or path.is_symlink():
        raise RuntimeError(f"{label} is not a regular file: {path}")
    return path.resolve()


def expected_artifact_path(expected: dict[str, Any], label: str) -> Path:
    raw_path = expected.get("path")
    if not isinstance(raw_path, str) or not raw_path:
        raise RuntimeError(f"{label} has no artifact path")
    path = Path(raw_path)
    if not path.is_absolute():
        path = PACKET / path
    return path.resolve()


def require_artifact_binding(value: Any, expected: dict[str, Any], label: str) -> Path:
    expected_path = expected_artifact_path(expected, label)
    if isinstance(value, str):
        actual_path = packet_file(value, label)
        if (
            actual_path != expected_path
            or actual_path.stat().st_size != expected.get("bytes")
            or sha(actual_path) != expected.get("sha256")
        ):
            raise RuntimeError(f"{label} differs from its expected artifact")
        return actual_path
    if not isinstance(value, dict):
        raise RuntimeError(f"{label} must be an artifact receipt")
    actual_path = packet_file(value.get("path"), f"{label}.path")
    if actual_path != expected_path:
        raise RuntimeError(f"{label}.path does not bind the expected file")
    declared_bytes = value.get("bytes", value.get("size"))
    declared_sha = value.get("sha256")
    if (
        declared_bytes != expected.get("bytes")
        or declared_sha != expected.get("sha256")
        or actual_path.stat().st_size != expected.get("bytes")
        or sha(actual_path) != expected.get("sha256")
    ):
        raise RuntimeError(f"{label} does not match its expected artifact")
    return actual_path


def first_mapping(value: dict[str, Any], keys: tuple[str, ...]) -> dict[str, Any] | None:
    for key in keys:
        candidate = value.get(key)
        if isinstance(candidate, dict):
            return candidate
    return None


def record_for(container: Any, corpus_id: str, aliases: tuple[str, ...] = ()) -> dict[str, Any] | None:
    names = (corpus_id, *aliases)
    if isinstance(container, dict):
        for name in names:
            candidate = container.get(name)
            if isinstance(candidate, dict):
                return candidate
        own_id = container.get("id", container.get("case_id", container.get("corpus_id")))
        if own_id in names:
            return container
        for key in ("rows", "entries", "cases", "corpora", "exports", "identities"):
            candidate = record_for(container.get(key), corpus_id, aliases)
            if candidate is not None:
                return candidate
    elif isinstance(container, list):
        for candidate in container:
            found = record_for(candidate, corpus_id, aliases)
            if found is not None:
                return found
    return None


def nested_identity(record: dict[str, Any], keys: tuple[str, ...]) -> dict[str, Any] | None:
    for key in keys:
        value = record.get(key)
        if isinstance(value, dict):
            return value
    return None


def source_output_identity(record: dict[str, Any], label: str) -> dict[str, Any]:
    source = nested_identity(record, ("source_archive", "source", "input", "source_file"))
    source_sha = record.get("source_archive_sha256", record.get("source_sha256"))
    source_bytes = record.get("source_archive_bytes", record.get("source_bytes"))
    if source is not None:
        source_sha = source.get("sha256", source.get("source_sha256", source.get("archive_sha256", source_sha)))
        source_bytes = source.get("bytes", source.get("size", source.get("archive_bytes", source_bytes)))
    published = nested_identity(record, ("published", "reference", "output", "result"))
    published_sha = record.get("published_sha256", record.get("output_sha256"))
    published_bytes = record.get("published_bytes", record.get("output_bytes"))
    if published is not None:
        published_sha = published.get("sha256", published.get("output_sha256", published.get("archive_sha256", published_sha)))
        published_bytes = published.get("bytes", published.get("size", published.get("archive_bytes", published_bytes)))
    if not is_sha256(source_sha) or not isinstance(source_bytes, int):
        raise RuntimeError(f"{label} has no complete source archive identity")
    if not is_sha256(published_sha) or not isinstance(published_bytes, int):
        raise RuntimeError(f"{label} has no complete published output identity")
    outputs = record.get(
        "policy_outputs",
        record.get("outputs", record.get("artifacts", record.get("policies"))),
    )
    if isinstance(outputs, dict):
        output_rows = outputs
    elif isinstance(outputs, list):
        output_rows = {}
        for item in outputs:
            if not isinstance(item, dict):
                raise RuntimeError(f"{label} has a malformed policy output")
            policy = item.get("policy", item.get("durability", item.get("level")))
            if not isinstance(policy, str) or policy in output_rows:
                raise RuntimeError(f"{label} has an invalid or duplicate policy output")
            output_rows[policy] = item
    else:
        raise RuntimeError(f"{label} has no policy output identities")
    policy_digests: dict[str, str] = {}
    for policy in ("default", "full", "file-only", "no-sync"):
        item = output_rows.get(policy)
        if not isinstance(item, dict):
            raise RuntimeError(f"{label} is missing {policy} output identity")
        digest = item.get("sha256", item.get("output_sha256", item.get("archive_sha256")))
        bytes_value = item.get(
            "bytes", item.get("output_bytes", item.get("archive_bytes", published_bytes))
        )
        if not is_sha256(digest) or not isinstance(bytes_value, int):
            raise RuntimeError(f"{label} has an incomplete {policy} output identity")
        if digest != published_sha or bytes_value != published_bytes:
            raise RuntimeError(f"{label} {policy} output differs from its qualification publication")
        policy_digests[policy] = digest
    return {
        "source_sha256": source_sha,
        "source_bytes": source_bytes,
        "published_sha256": published_sha,
        "published_bytes": published_bytes,
        "policy_outputs": policy_digests,
    }


def check_generated_identity(
    corpus: dict[str, Any], qualification: dict[str, Any], record: dict[str, Any], label: str
) -> None:
    expected_ordinary = qualification.get("ordinary_corpus")
    if not isinstance(expected_ordinary, dict):
        raise RuntimeError(f"qualification identity for {corpus['id']} is malformed")
    expected = source_output_identity({
        "source_archive_sha256": expected_ordinary.get("source_archive_sha256"),
        "source_archive_bytes": expected_ordinary.get("source_archive_bytes"),
        "published_sha256": expected_ordinary.get("published_sha256"),
        "published_bytes": expected_ordinary.get("published_bytes"),
        "policy_outputs": {policy: {
            "sha256": expected_ordinary.get("published_sha256"),
            "bytes": expected_ordinary.get("published_bytes"),
        } for policy in ("default", "full", "file-only", "no-sync")},
    }, f"qualification {corpus['id']}")
    actual = source_output_identity(record, label)
    if actual != expected:
        raise RuntimeError(f"{label} generated source/output identity differs from qualification")
    expected_result = qualification.get("result_corpus")
    actual_result = record.get("corpus")
    if isinstance(expected_result, dict) and isinstance(actual_result, dict):
        expected_archive = (
            expected_result.get("archive_sha256"),
            expected_result.get("archive_bytes"),
        )
        actual_archive = (
            actual_result.get("archive_sha256"),
            actual_result.get("archive_bytes"),
        )
        if not is_sha256(expected_archive[0]) or not isinstance(expected_archive[1], int):
            raise RuntimeError(f"qualification {corpus['id']} has no complete corpus archive identity")
        if actual_archive != expected_archive:
            raise RuntimeError(f"{label} corpus archive identity differs from qualification")
    if "origin" in record and record["origin"] != GENERATED_ORIGIN:
        raise RuntimeError(f"{label} is not marked as a generated corpus")
    if "edit_admitted" in record and record["edit_admitted"] is not True:
        raise RuntimeError(f"{label} generated admission differs from the plan")


def check_real_fixture_rows(plan: dict[str, Any], container: Any, label: str) -> None:
    if container is None:
        raise RuntimeError(f"{label} has no real fixture identity map")
    for corpus in plan["corpora"]:
        path_value = corpus.get("path")
        if path_value is None:
            continue
        row = record_for(container, corpus["id"], (path_value,))
        if row is None:
            raise RuntimeError(f"{label} is missing fixture identity for {corpus['id']}")
        identity = nested_identity(row, ("source", "fixture", "file", "source_archive")) or row
        if identity.get("path", path_value) != path_value or identity.get("bytes", identity.get("size")) != corpus["bytes"] or identity.get("sha256") != corpus["sha256"]:
            raise RuntimeError(f"{label} fixture identity differs for {corpus['id']}")


def export_and_admission_guards(
    plan: dict[str, Any], source: dict[str, Any], build: dict[str, Any], admission: dict[str, Any]
) -> tuple[Path, Path]:
    export_path = PACKET / "export.json"
    if not export_path.is_file():
        raise RuntimeError("admission requires the retained export.json receipt")
    export_receipt = read(export_path)
    if not isinstance(export_receipt, dict) or export_receipt.get("exit_code") != 0:
        raise RuntimeError("export.json is missing or the exporter did not complete")
    export_binary = build["binaries"]["export"]
    if export_receipt.get("binary") != export_binary:
        raise RuntimeError("export receipt binary is not the frozen export binary")
    for key, expected in (("source_sha256", sha(SOURCE_PATH)), ("plan_sha256", sha(PLAN_PATH)), ("build_sha256", sha(BUILD_PATH))):
        if export_receipt.get(key) != expected:
            raise RuntimeError(f"export receipt {key} is not bound to the frozen input")
    if export_receipt.get("runner_sha256") != sha(PACKET / "export.py"):
        raise RuntimeError("export receipt runner binding changed")
    manifest_value = export_receipt.get("manifest")
    if not isinstance(manifest_value, dict):
        raise RuntimeError("export receipt has no manifest artifact")
    manifest_path = require_artifact_binding(manifest_value, manifest_value, "export receipt manifest")
    real_fixtures = {row["id"]: {
        "path": row["path"], "bytes": row["bytes"], "sha256": row["sha256"]
    } for row in plan["corpora"] if row.get("path") is not None}
    receipt_fixtures = export_receipt.get("fixtures")
    for corpus in plan["corpora"]:
        if corpus.get("path") is None:
            continue
        expected = real_fixtures[corpus["id"]]
        candidate = None
        if isinstance(receipt_fixtures, dict):
            candidate = receipt_fixtures.get(expected["path"], receipt_fixtures.get(corpus["id"]))
        if not isinstance(candidate, dict) or candidate.get("bytes") != expected["bytes"] or candidate.get("sha256") != expected["sha256"]:
            raise RuntimeError(f"export receipt fixture identity differs for {corpus['id']}")

    binding = first_mapping(admission, ("export", "export_receipt", "export_binding"))
    if binding is None:
        raise RuntimeError("admission.json has no export receipt binding")
    receipt_binding = binding.get("receipt", binding.get("export_receipt", binding.get("record")))
    if receipt_binding is None and binding.get("path"):
        receipt_binding = binding
    require_artifact_binding(receipt_binding, artifact(export_path), "admission export receipt")
    binary_binding = binding.get("binary", binding.get("binary_receipt"))
    if binary_binding is None:
        binary_binding = admission.get("export_binary")
    if binary_binding is None:
        raise RuntimeError("admission export binding has no binary receipt")
    require_artifact_binding(binary_binding, export_binary, "admission export binary")
    source_binding_sha = binding.get("source_sha256", admission.get("source_sha256"))
    source_binding_revision = binding.get("source_revision", admission.get("source_revision"))
    if source_binding_sha != sha(SOURCE_PATH) or source_binding_revision != source.get("revision"):
        raise RuntimeError("admission export binding is not source-bound")
    if binding.get("plan_sha256", admission.get("plan_sha256")) != sha(PLAN_PATH) or binding.get("build_sha256", admission.get("build_sha256")) != sha(BUILD_PATH):
        raise RuntimeError("admission export binding is not plan/build-bound")
    manifest_binding = binding.get("manifest", binding.get("manifest_receipt", admission.get("export_manifest")))
    if manifest_binding is None:
        raise RuntimeError("admission export binding has no manifest receipt")
    require_artifact_binding(manifest_binding, artifact(manifest_path), "admission export manifest")
    check_real_fixture_rows(plan, admission.get("real_fixtures", admission.get("fixtures")), "admission")

    qualification_path = CAPTURE / "qualification-identities.json"
    qualification = read(qualification_path)
    qualification_rows = qualification.get("corpora")
    if not isinstance(qualification_rows, dict):
        raise RuntimeError("qualification identities are malformed")
    generated_container = admission.get(
        "generated",
        admission.get(
            "generated_corpora",
            admission.get("generated_identities", admission.get("corpora")),
        ),
    )
    if generated_container is None and isinstance(admission.get("oracle"), dict):
        generated_container = admission["oracle"].get("corpora")
    if generated_container is None:
        raise RuntimeError("admission.json has no generated source/output identities")
    for corpus in plan["corpora"]:
        if corpus.get("path") is not None:
            continue
        aliases = (f"{corpus['id']}-medium", f"generated-{corpus['format']}-medium")
        row = record_for(generated_container, corpus["id"], aliases)
        if row is None:
            raise RuntimeError(f"admission is missing generated identity for {corpus['id']}")
        expected_qualification = qualification_rows.get(corpus["id"])
        if not isinstance(expected_qualification, dict):
            raise RuntimeError(f"qualification identity is missing for {corpus['id']}")
        check_generated_identity(corpus, expected_qualification, row, f"admission {corpus['id']}")

    manifest = read(manifest_path)
    manifest_cases = manifest.get("cases") if isinstance(manifest, dict) else None
    if not isinstance(manifest_cases, list):
        raise RuntimeError("export manifest has no case list")
    for corpus in plan["corpora"]:
        if corpus.get("path") is not None:
            continue
        aliases = {corpus["id"], f"{corpus['id']}-medium", f"generated-{corpus['format']}-medium"}
        matches = [row for row in manifest_cases if isinstance(row, dict) and row.get("case_id", row.get("id")) in aliases]
        if len(matches) != 1:
            raise RuntimeError(f"export manifest has no unique generated identity for {corpus['id']}")
        expected_qualification = qualification_rows[corpus["id"]]
        check_generated_identity(corpus, expected_qualification, matches[0], f"export manifest {corpus['id']}")
    return export_path, manifest_path


def admission_guard(
    plan: dict[str, Any], source: dict[str, Any], build: dict[str, Any]
) -> dict[str, Any]:
    if not ADMISSION_PATH.is_file():
        raise RuntimeError("native/allocation lanes require admission.json from the export oracle")
    admission = read(ADMISSION_PATH)
    if not isinstance(admission, dict) or admission.get("oracle_pass") is not True:
        raise RuntimeError("export oracle did not pass")
    expected_sha = admission.get("oracle_report_sha256")
    if not is_sha256(expected_sha):
        raise RuntimeError("admission.json has no oracle report digest")
    declared = []
    for key in ("oracle_report", "oracle_report_path", "report"):
        value = admission.get(key)
        if isinstance(value, str):
            path = Path(value)
            if not path.is_absolute():
                path = PACKET / path
            if path.is_file() and not path.is_symlink():
                declared.append(path.resolve())
        elif isinstance(value, dict) and isinstance(value.get("path"), str):
            path = Path(value["path"])
            if not path.is_absolute():
                path = PACKET / path
            if path.is_file() and not path.is_symlink():
                declared.append(path.resolve())
    if not declared:
        declared = [path for path in PACKET.rglob("*") if path.is_file() and not path.is_symlink() and path != ADMISSION_PATH and sha(path) == expected_sha]
    matching = [path for path in declared if sha(path) == expected_sha]
    if len(matching) != 1:
        raise RuntimeError(f"cannot bind oracle report digest {expected_sha}: {matching!r}")
    identities_sha = admission.get("qualification_identities_sha256")
    identities = CAPTURE / "qualification-identities.json"
    if not is_sha256(identities_sha) or not identities.is_file() or sha(identities) != identities_sha:
        raise RuntimeError("admission qualification identities do not match the frozen qualification")
    export_path, manifest_path = export_and_admission_guards(plan, source, build, admission)
    if admission.get("oracle_runner_sha256") != sha(PACKET / "oracle.py"):
        raise RuntimeError("admitted oracle implementation changed")
    oracle_report = read(matching[0])
    if oracle_report.get("status") != "pass" or oracle_report.get("manifest_sha256") != sha(manifest_path):
        raise RuntimeError("oracle report is not bound to the exported manifest")
    if packet_file(oracle_report.get("manifest"), "oracle manifest") != manifest_path:
        raise RuntimeError("oracle report names a different export manifest")
    return {
        "admission": artifact(ADMISSION_PATH),
        "oracle_report": artifact(matching[0]),
        "export_receipt": artifact(export_path),
        "export_manifest": artifact(manifest_path),
        "qualification_identities_sha256": identities_sha,
        "oracle_pass": True,
    }


def check_admission_receipt(
    expected: dict[str, Any], plan: dict[str, Any], source: dict[str, Any], build: dict[str, Any]
) -> None:
    actual = admission_guard(plan, source, build)
    if actual != expected:
        raise RuntimeError("admission/export/oracle receipt changed after timed lane start")


def run_job(
    plan: dict[str, Any],
    source: dict[str, Any],
    build: dict[str, Any],
    freeze: dict[str, Any],
    fixtures: dict[str, Any],
    identities: dict[str, dict[str, Any]],
    job: dict[str, Any],
    lane_dir: Path,
    lane_rows: list[dict[str, Any]],
    environment_value: dict[str, str],
) -> dict[str, Any]:
    binaries = verify_frozen(plan, source, build, freeze)
    corpus = corpus_by_id(plan)[job["corpus_id"]]
    stem = job_stem(job)
    report = report_path_for(lane_dir, stem, "json")
    stdout = report_path_for(lane_dir, stem, "stdout")
    stderr = report_path_for(lane_dir, stem, "stderr")
    rss = report_path_for(lane_dir, stem, "rss")
    for path in (report, stdout, stderr, rss):
        if path.exists():
            raise RuntimeError(f"refusing to overwrite job artifact: {path}")
    binary_name = "allocation" if job["lane"] == "allocation" else "native"
    binary = binary_path(binaries, binary_name)
    command = command_for(plan, job, binary, report, rss)
    fixture_before = copy.deepcopy(fixtures)
    started_ns = time.time_ns()
    monotonic_start = time.monotonic_ns()
    with stdout.open("wb") as out, stderr.open("wb") as err:
        process = subprocess.run(command, cwd=ROOT, env=environment_value, stdout=out, stderr=err)
    monotonic_end = time.monotonic_ns()
    fixture_after = fixture_inventory(plan)
    if fixture_after != fixture_before or fixture_after != fixtures:
        raise RuntimeError(f"fixture changed during {stem}")
    if process.returncode != 0:
        # Preserve the failed process evidence in the manifest before raising.
        row = {
            "job": job,
            "command": command,
            "exit_code": process.returncode,
            "started_ns": started_ns,
            "ended_ns": time.time_ns(),
            "monotonic_elapsed_ns": monotonic_end - monotonic_start,
            "binary": binaries[binary_name],
            "stdout": artifact(stdout),
            "stderr": artifact(stderr),
            "rss": artifact(rss) if rss.is_file() else None,
            "fixture_before": fixture_before,
            "fixture_after": fixture_after,
            "plan_sha256": sha(PLAN_PATH),
            "source_sha256": sha(SOURCE_PATH),
            "build_sha256": sha(BUILD_PATH),
            "runner_sha256": sha(RUNNER_PATH),
        }
        lane_rows.append(row)
        write(lane_dir.parent / f"{job['lane']}.json", {"schema": "litchi-0778-durability-runs-v1", "rows": lane_rows}, replace=True)
        raise RuntimeError(f"{stem} exited {process.returncode}; see {stderr}")
    if not report.is_file():
        raise RuntimeError(f"{stem} exited successfully without a JSON report")
    rss_text = rss.read_text(encoding="utf-8").strip()
    if not rss_text.isdigit() or int(rss_text) <= 0:
        raise RuntimeError(f"invalid RSS receipt for {stem}: {rss_text!r}")
    identity, sequences = validate_report(
        plan,
        job,
        report,
        identities.get(job["corpus_id"]),
        binaries[binary_name],
    )
    if job["corpus_id"] not in identities:
        identities[job["corpus_id"]] = identity
    row = {
        "job": job,
        "command": command,
        "exit_code": process.returncode,
        "started_ns": started_ns,
        "ended_ns": time.time_ns(),
        "monotonic_elapsed_ns": monotonic_end - monotonic_start,
        "environment": {key: environment_value.get(key) for key in ("LC_ALL", "LANG", "TZ")},
        "binary": binaries[binary_name],
        "stdout": artifact(stdout),
        "stderr": artifact(stderr),
        "rss": {**artifact(rss), "rss_kib": int(rss_text)},
        "report": artifact(report),
        "report_identity_sha256": canonical_sha(identity),
        "publication_sha256": sequences["publication_sha256"],
        "edit_outcome_sha256": sequences["edit_outcome_sha256"],
        "edit_admitted": sequences["edit_admitted"],
        "edit_outcome": sequences["edit_outcome"],
        "source_archive_sha256": sequences["source_archive_sha256"],
        "fixture_before": fixture_before,
        "fixture_after": fixture_after,
        "plan_sha256": sha(PLAN_PATH),
        "source_sha256": sha(SOURCE_PATH),
        "build_sha256": sha(BUILD_PATH),
        "runner_sha256": sha(RUNNER_PATH),
    }
    lane_rows.append(row)
    write(lane_dir.parent / f"{job['lane']}.json", {"schema": "litchi-0778-durability-runs-v1", "rows": lane_rows}, replace=True)
    print(f"{stem} PASS", flush=True)
    return row


def run_qualification(plan: dict[str, Any], source: dict[str, Any], build: dict[str, Any]) -> None:
    lane_dir = CAPTURE / "qualification-0"
    if lane_dir.exists() or (CAPTURE / "qualification.json").exists():
        raise RuntimeError(f"refusing to overwrite existing qualification lane: {lane_dir}")
    binaries = binary_receipts(build)
    fixtures = fixture_inventory(plan)
    lane_dir.mkdir()
    freeze = freeze_inputs(plan, source, build)
    write(CAPTURE / "runner.json", runner_receipt())
    write(CAPTURE / "freeze.json", freeze)
    write(CAPTURE / "fixtures-before.json", fixtures)
    rows: list[dict[str, Any]] = []
    identities: dict[str, dict[str, Any]] = {}
    env = environment(plan["cpu"])
    for job in jobs_for_lane(plan, "qualification"):
        run_job(plan, source, build, freeze, fixtures, identities, job, lane_dir, rows, env)
        # Qualification must establish a single identity per corpus across all four policies.
        if job["corpus_id"] not in identities:
            raise RuntimeError("qualification failed to freeze a corpus identity")
    verify_frozen(plan, source, build, freeze)
    if len(rows) != plan["expected_children"]["qualification"] or len(identities) != len(plan["corpora"]):
        raise RuntimeError("qualification cardinality or identity cardinality is incomplete")
    identity_record = {
        "schema": "litchi-0778-durability-qualification-v1",
        "plan_sha256": sha(PLAN_PATH),
        "source_sha256": sha(SOURCE_PATH),
        "build_sha256": sha(BUILD_PATH),
        "runner_sha256": sha(RUNNER_PATH),
        "corpora": identities,
    }
    write(CAPTURE / "qualification-identities.json", identity_record)
    write(CAPTURE / "qualification.json", {
        "schema": "litchi-0778-durability-qualification-runs-v1",
        "rows": rows,
        "expected_children": plan["expected_children"]["qualification"],
        "identities_sha256": sha(CAPTURE / "qualification-identities.json"),
        "binaries": binaries,
        "fixtures": fixtures,
    }, replace=True)
    write(lane_dir / "complete.json", {
        "schema": "litchi-0778-durability-lane-v1",
        "lane": "qualification",
        "rows": len(rows),
        "expected_rows": plan["expected_children"]["qualification"],
        "serial": True,
        "source_unchanged": True,
        "fixtures_unchanged": True,
        "binaries_unchanged": True,
        "runner_unchanged": True,
        "qualification_identities_sha256": sha(CAPTURE / "qualification-identities.json"),
    })
    print(json.dumps({"lane": "qualification", "rows": len(rows), "identities": len(identities)}, indent=2), flush=True)


def run_timed_lane(plan: dict[str, Any], source: dict[str, Any], build: dict[str, Any], lane: str) -> None:
    if not (CAPTURE / "qualification-0" / "complete.json").is_file():
        raise RuntimeError("qualification must complete before timed lanes")
    if lane == "allocation" and not (CAPTURE / "native-0" / "complete.json").is_file():
        raise RuntimeError("native lane must complete before allocation")
    freeze = read(CAPTURE / "freeze.json")
    qualification = read(CAPTURE / "qualification-identities.json")
    identities = qualification.get("corpora")
    if not isinstance(identities, dict) or len(identities) != len(plan["corpora"]):
        raise RuntimeError("qualification identities are missing or incomplete")
    qualification_runs = read(CAPTURE / "qualification.json")
    if qualification_runs.get("identities_sha256") != sha(CAPTURE / "qualification-identities.json"):
        raise RuntimeError("qualification identity digest is inconsistent")
    admission = admission_guard(plan, source, build)
    admission_receipt_path = CAPTURE / "admission-before.json"
    if admission_receipt_path.exists():
        expected_admission = read(admission_receipt_path)
        if expected_admission != admission:
            raise RuntimeError("admission changed between timed lanes")
    else:
        write(admission_receipt_path, admission)
    lane_dir = CAPTURE / f"{lane}-0"
    if lane_dir.exists() or (CAPTURE / f"{lane}.json").exists():
        raise RuntimeError(f"refusing to overwrite existing {lane} lane")
    lane_dir.mkdir()
    fixtures = read(CAPTURE / "fixtures-before.json")
    if not isinstance(fixtures, dict):
        raise RuntimeError("malformed fixture freeze")
    rows: list[dict[str, Any]] = []
    env = environment(plan["cpu"])
    for job in jobs_for_lane(plan, lane):
        check_admission_receipt(admission, plan, source, build)
        run_job(plan, source, build, freeze, fixtures, identities, job, lane_dir, rows, env)
    verify_frozen(plan, source, build, freeze)
    check_admission_receipt(admission, plan, source, build)
    expected_count = plan["expected_children"][lane]
    if len(rows) != expected_count:
        raise RuntimeError(f"{lane} cardinality is {len(rows)}, expected {expected_count}")
    write(lane_dir / "complete.json", {
        "schema": "litchi-0778-durability-lane-v1",
        "lane": lane,
        "rows": len(rows),
        "expected_rows": expected_count,
        "serial": True,
        "admission": admission,
        "qualification_identities_sha256": sha(CAPTURE / "qualification-identities.json"),
        "source_unchanged": True,
        "fixtures_unchanged": True,
        "binaries_unchanged": True,
        "runner_unchanged": True,
    })
    print(json.dumps({"lane": lane, "rows": len(rows)}, indent=2), flush=True)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("lane", choices=("qualification", "native", "allocation"))
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    plan, source, build = load_inputs()
    if args.lane == "qualification":
        run_qualification(plan, source, build)
    else:
        run_timed_lane(plan, source, build, args.lane)
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as error:
        print(f"capture failed: {error}", file=sys.stderr)
        raise
