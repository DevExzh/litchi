#!/usr/bin/env python3
"""Capture the gated 0711 DOCX edit-owner mechanism evidence.

This packet is deliberately downstream of the native paired pilot.  It never
builds Cargo artifacts and never edits or restores the checkout.  The
coordinator prepares ``build-baseline.json`` and ``build-candidate.json`` and
leaves the candidate source checked out; this driver then runs only the frozen
Callgrind and ``/usr/bin/time -v`` child matrix.  A baseline binary may
therefore execute while the candidate checkout is current, but that relation is
recorded and checked instead of being hidden.
"""

from __future__ import annotations

import argparse
import datetime as _datetime
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import re
import subprocess
import sys
import time
from typing import Any, NoReturn


HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
PLAN_PATH = HERE / "mechanism-plan.json"
PILOT_PLAN_PATH = HERE / "plan.json"
OWNER = "litchi_perf_baseline::ordinary_save::Owner::edit"
MEASURED_PARENT = "litchi_perf_baseline::ordinary_save::run_case"
TARGET_DELTA = "crates/litchi-docx/src/alt/codec.rs"
STAGE_BY_LABEL = {
    "baseline-A1": {"source": "baseline", "build": "baseline", "pair": "pair-1", "repeat": 1, "order": "forward"},
    "candidate-B1": {"source": "candidate", "build": "candidate", "pair": "pair-1", "repeat": 1, "order": "forward"},
    "candidate-B2": {"source": "candidate", "build": "candidate", "pair": "pair-2", "repeat": 2, "order": "reverse"},
    "baseline-A2": {"source": "baseline", "build": "baseline", "pair": "pair-2", "repeat": 2, "order": "reverse"},
}
PASS_WORDS = {"pass", "passed", "ok", "accepted", "green"}
FAIL_WORDS = {"fail", "failed", "reject", "rejected", "blocked", "red"}
PEAK_RE = re.compile(r"^Maximum resident set size \(kbytes\):\s*(\d+)\s*$", re.MULTILINE)


class EvidenceError(RuntimeError):
    """A missing, changed, or contradictory frozen input."""


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
    require(not path.exists() and not path.is_symlink(), f"refusing to replace {path}")
    path.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n", encoding="utf-8")


def utc_now() -> str:
    return _datetime.datetime.now(_datetime.timezone.utc).isoformat()


def slug(value: str) -> str:
    return "".join(char if char.isalnum() or char in "-_" else "_" for char in value)


def load_module(path: Path, name: str) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing helper: {path}")
    spec = importlib.util.spec_from_file_location(name, path)
    require(spec is not None and spec.loader is not None, f"cannot load helper: {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def load_custody() -> Any:
    custody_path = HERE / "custody.py"
    module = load_module(custody_path, "custody_0711_mechanism")
    require(callable(getattr(module, "census", None)), "custody.py has no census helper")
    return module


def source_census() -> dict[str, str]:
    custody = load_custody()
    value = custody.census()
    require(isinstance(value, dict) and value, "source census is empty")
    return dict(sorted(value.items()))


def load_plan() -> dict[str, Any]:
    plan = read_json(PLAN_PATH)
    require(isinstance(plan, dict), "mechanism-plan.json is not an object")
    require(plan.get("schema_version") == 1, "mechanism plan schema changed")
    pilot_plan = read_json(PILOT_PLAN_PATH)
    require(plan.get("revision") == pilot_plan.get("revision"),
            "mechanism revision differs from the native pilot revision")
    require(plan.get("cpu") == 12, "mechanism CPU changed")
    require(plan.get("filesystem_root") == "/home/zhuhe/code/litchi-0711-fs",
            "mechanism filesystem root changed")
    corpora = plan.get("corpora")
    require(isinstance(corpora, list) and len(corpora) == 2,
            "mechanism must contain exactly two corpora")
    ids: set[str] = set()
    for corpus in corpora:
        require(isinstance(corpus, dict), "malformed corpus entry")
        identity = corpus.get("id")
        require(isinstance(identity, str) and identity and identity not in ids,
                "corpus ids must be unique")
        ids.add(identity)
        require(corpus.get("origin") in {"generated-harness-corpus", "caller-named-real-file"},
                f"unsupported corpus origin for {identity}")
        if corpus["origin"] == "generated-harness-corpus":
            require(corpus.get("path") is None and corpus.get("sha256") is None,
                    f"generated corpus {identity} has a fixture binding")
        else:
            require(isinstance(corpus.get("path"), str) and corpus["path"],
                    f"real corpus {identity} has no path")
            digest = corpus.get("sha256")
            require(isinstance(digest, str) and len(digest) == 64 and
                    set(digest) <= set("0123456789abcdef"),
                    f"real corpus {identity} has an invalid digest")
    require(ids == {"generated", "numbered-list"}, "mechanism corpus matrix changed")
    stages = plan.get("stages")
    require(stages == [{"label": label, **STAGE_BY_LABEL[label]} for label in STAGE_BY_LABEL],
            "mechanism stage order or metadata changed")
    gate = plan.get("pilot_gate")
    require(isinstance(gate, dict) and gate.get("required_status") == "pass",
            "native pilot gate policy changed")
    callgrind = plan.get("callgrind")
    require(isinstance(callgrind, dict), "Callgrind policy is missing")
    require(callgrind.get("repeats") == 2 and callgrind.get("samples") == 1
            and callgrind.get("warmup") == 0 and callgrind.get("phase") == "edit",
            "Callgrind sample policy changed")
    require(callgrind.get("owner") == OWNER and callgrind.get("owner_match") == "exact",
            "Callgrind owner policy changed")
    require(callgrind.get("measured_parent") == MEASURED_PARENT,
            "Callgrind measured parent changed")
    require(callgrind.get("setup_parents") == {
        "litchi_perf_baseline::ordinary_save::build_corpus": 1,
        "litchi_perf_baseline::ordinary_save::publish_reference": 3,
    }, "Callgrind setup policy changed")
    require(callgrind.get("expected_numbered_parts") == 5,
            "Callgrind part policy changed")
    rss = plan.get("rss")
    require(isinstance(rss, dict) and rss.get("tool") == "/usr/bin/time"
            and rss.get("verbose") is True, "RSS tool policy changed")
    require(rss.get("repeats") == 2 and rss.get("samples") == 10
            and rss.get("warmup") == 2 and rss.get("phases") == ["edit", "lifecycle"],
            "RSS sample policy changed")
    require(rss.get("peak_flag_percent") == 5 and rss.get("repeat_flag_percent") == 5
            and rss.get("no_peak_summation") is True, "RSS threshold policy changed")
    custody = plan.get("source_custody")
    require(isinstance(custody, dict)
            and custody.get("allowed_delta_paths") == [TARGET_DELTA]
            and custody.get("require_target_changed") is True,
            "source-delta policy changed")
    helper_paths = plan.get("retained_helper_sha256")
    require(isinstance(helper_paths, dict) and set(helper_paths) == set(plan["retained_helpers"]),
            "retained helper identities are incomplete")
    for raw, digest in helper_paths.items():
        path = REPO / raw
        require(path.is_file() and not path.is_symlink() and sha(path) == digest,
                f"retained helper changed: {raw}")
    return plan


def constraints_check() -> dict[str, str]:
    path = HERE / "constraints.json"
    constraints = read_json(path)
    require(isinstance(constraints, dict), "constraints.json is not an object")
    for raw, digest in constraints.items():
        target = REPO / raw
        require(target.is_file() and not target.is_symlink() and sha(target) == digest,
                f"constraint changed: {raw}")
    return {str(raw): str(digest) for raw, digest in sorted(constraints.items())}


def corpus_map(plan: dict[str, Any]) -> dict[str, dict[str, Any]]:
    return {item["id"]: item for item in plan["corpora"]}


def fixture_binding(corpus: dict[str, Any]) -> dict[str, Any]:
    raw = corpus.get("path")
    if corpus["origin"] == "generated-harness-corpus":
        return {"plan_path": None, "resolved_path": None, "bytes": None, "sha256": None}
    require(isinstance(raw, str), f"{corpus['id']}: fixture path is missing")
    path = (REPO / raw).resolve() if not Path(raw).is_absolute() else Path(raw).resolve()
    require(path.is_file() and not path.is_symlink(), f"missing fixture: {path}")
    digest = sha(path)
    require(digest == corpus["sha256"], f"fixture digest changed: {corpus['id']}")
    require(path.stat().st_size == corpus["bytes"], f"fixture size changed: {corpus['id']}")
    return {"plan_path": raw, "resolved_path": str(path),
            "bytes": path.stat().st_size, "sha256": digest}


def source_diff(left: dict[str, str], right: dict[str, str]) -> list[str]:
    return sorted(name for name in set(left) | set(right) if left.get(name) != right.get(name))


def source_state(plan: dict[str, Any]) -> dict[str, Any]:
    custody = plan["source_custody"]
    baseline_path = HERE / custody["baseline_manifest"]
    candidate_path = HERE / custody["candidate_manifest"]
    baseline = read_json(baseline_path)
    candidate = read_json(candidate_path)
    require(isinstance(baseline, dict) and isinstance(candidate, dict),
            "source manifests are not objects")
    changed = source_diff(baseline, candidate)
    require(changed == [TARGET_DELTA], f"candidate source delta is {changed}, expected only {TARGET_DELTA}")
    current = source_census()
    require(current == candidate, "current checkout is not the candidate source manifest")
    return {
        "baseline": baseline,
        "candidate": candidate,
        "baseline_path": baseline_path.name,
        "candidate_path": candidate_path.name,
        "baseline_sha256": sha(baseline_path),
        "candidate_sha256": sha(candidate_path),
        "baseline_census_sha256": digest_json(baseline),
        "candidate_census_sha256": digest_json(candidate),
        "changed_paths": changed,
        "current": current,
    }


def build_info(plan: dict[str, Any], stage: str) -> dict[str, Any]:
    metadata = STAGE_BY_LABEL[stage]
    record_path = HERE / f"build-{metadata['build']}.json"
    rows = read_json(record_path)
    require(isinstance(rows, list), f"{record_path.name} is not a record list")
    binary_name = f"{metadata['build']}-native"
    matches = [item for item in rows if isinstance(item, dict)
               and Path(str(item.get("binary", ""))).name == binary_name]
    require(len(matches) == 1, f"no unique native build record for {binary_name}")
    record = matches[0]
    require(record.get("exit_code") == 0, f"{binary_name} build did not pass")
    binary_raw = record.get("binary")
    binary = Path(str(binary_raw)).resolve()
    digest = record.get("binary_sha256")
    size = record.get("binary_bytes")
    require(isinstance(digest, str) and len(digest) == 64,
            f"{binary_name} binary digest is missing")
    require(isinstance(size, int) and size > 0, f"{binary_name} binary size is invalid")
    require(binary.is_file() and not binary.is_symlink(), f"missing native binary: {binary}")
    require(sha(binary) == digest and binary.stat().st_size == size,
            f"{binary_name} binary identity changed")
    source_path = HERE / f"source-{metadata['source']}.json"
    source = read_json(source_path)
    require(isinstance(source, dict) and source, f"{source_path.name} is invalid")
    require(record.get("source_manifest_sha256") == sha(source_path),
            f"{binary_name} source manifest binding changed")
    return {
        "record": record,
        "record_path": record_path.name,
        "record_sha256": sha(record_path),
        "binary": str(binary),
        "binary_sha256": digest,
        "binary_bytes": size,
        "source_label": metadata["source"],
        "source_path": source_path.name,
        "source_manifest_sha256": sha(source_path),
        "source_census_sha256": digest_json(source),
        "source_entry_count": len(source),
    }


def stage_record(plan: dict[str, Any], stage: str) -> dict[str, Any]:
    return build_info(plan, stage)


def case_name(corpus: dict[str, Any], phase: str) -> str:
    prefix = ("docx_ordinary_save_" if corpus["origin"] == "generated-harness-corpus"
              else "docx_real_file_ordinary_save_")
    return prefix + phase


def stage_metadata(stage: str) -> dict[str, Any]:
    require(stage in STAGE_BY_LABEL, f"unknown stage: {stage}")
    return STAGE_BY_LABEL[stage]


def ordered_corpora(plan: dict[str, Any], stage: str) -> list[dict[str, Any]]:
    corpora = list(plan["corpora"])
    if stage_metadata(stage)["order"] == "reverse":
        corpora.reverse()
    return corpora


def profile_name(stage: str, corpus: dict[str, Any]) -> str:
    return f"mechanism-profile-{slug(stage)}-{slug(corpus['id'])}-edit"


def rss_name(stage: str, corpus: dict[str, Any], phase: str) -> str:
    return f"mechanism-rss-{slug(stage)}-{slug(corpus['id'])}-{slug(phase)}"


def pilot_name(stage: str, corpus: dict[str, Any], phase: str) -> str:
    # capture.py's canonical child_name puts the lane after the stage.
    return f"{slug(stage)}-native-{slug(corpus['id'])}-{slug(phase)}"


def artifact_paths(name: str, kind: str) -> list[Path]:
    common = [f"{name}.json", f"{name}.stdout", f"{name}.stderr",
              f"{name}.source-before.json", f"{name}.source-after.json",
              f"{name}.fixture-before.json", f"{name}.fixture-after.json",
              f"{name}.receipt.json"]
    if kind == "profile":
        common += [f"{name}.callgrind"]
        common += [f"{name}.callgrind.{number}" for number in range(1, 128)]
    else:
        common += [f"{name}.time-v"]
    return [HERE / path for path in common]


def refuse_existing(name: str, kind: str) -> None:
    existing = [path for path in artifact_paths(name, kind) if path.exists() or path.is_symlink()]
    require(not existing, f"refusing to replace existing {name} artifacts: {existing[:5]}")


def source_artifact(name: str, value: dict[str, str], suffix: str) -> tuple[str, str]:
    path = HERE / f"{name}.{suffix}.json"
    write_json(path, value)
    return path.name, sha(path)


def fixture_artifact(name: str, value: dict[str, Any], suffix: str) -> tuple[str, str]:
    path = HERE / f"{name}.{suffix}.json"
    write_json(path, value)
    return path.name, sha(path)


def stable_output(value: Any) -> Any:
    """Drop only per-run measurement envelopes from a benchmark report."""

    if isinstance(value, list):
        return [stable_output(item) for item in value]
    if not isinstance(value, dict):
        return value
    ignored = {
        "elapsed_ns", "operation_metrics", "allocation_metrics", "timing_metrics",
        "sample_indices", "samples", "confidence_interval_95", "standard_deviation",
        "min", "p50", "p95", "p99", "max", "mean", "elapsed", "wall_time_ns",
    }
    return {key: stable_output(item) for key, item in sorted(value.items()) if key not in ignored}


def report_identity(path: Path) -> dict[str, Any]:
    report = read_json(path)
    require(isinstance(report, dict), f"{path.name}: benchmark report is not an object")
    results = report.get("results")
    require(isinstance(results, list) and len(results) == 1,
            f"{path.name}: expected one benchmark result")
    result = results[0]
    require(isinstance(result, dict), f"{path.name}: result is not an object")
    require(isinstance(result.get("case"), str), f"{path.name}: case is missing")
    require(isinstance(result.get("corpus"), dict), f"{path.name}: corpus identity is missing")
    identity = {
        "case": result["case"],
        "corpus": stable_output(result["corpus"]),
        "sink": stable_output(result.get("sink")),
        "source": stable_output(result.get("source")),
        "published_sha256": stable_output(
            result.get("source", {}).get("ordinary_save", {}).get("published_sha256")
            if isinstance(result.get("source"), dict) else None),
        "output_sha256": stable_output(
            result.get("source", {}).get("ordinary_save", {}).get("repeated_save_sha256")
            if isinstance(result.get("source"), dict) else None),
    }
    return identity


def pilot_gate_status(value: Any) -> str | None:
    if not isinstance(value, dict):
        return None
    direct_keys = ("native_pilot_gate", "pilot_gate", "mechanism_gate", "gate_status",
                   "status", "verdict")
    for key in direct_keys:
        raw = value.get(key)
        if isinstance(raw, str):
            lowered = raw.lower()
            if lowered in PASS_WORDS:
                return "pass"
            if lowered in FAIL_WORDS:
                return "fail"
        if isinstance(raw, bool) and key != "status":
            return "pass" if raw else "fail"
        if isinstance(raw, dict):
            status = pilot_gate_status(raw)
            if status:
                return status
    for key, child in value.items():
        if isinstance(key, str) and "gate" in key.lower():
            if isinstance(child, str):
                lowered = child.lower()
                if lowered in PASS_WORDS:
                    return "pass"
                if lowered in FAIL_WORDS:
                    return "fail"
            if isinstance(child, bool):
                return "pass" if child else "fail"
            status = pilot_gate_status(child)
            if status:
                return status
    return None


def pilot_receipt_names(plan: dict[str, Any]) -> list[tuple[str, str, dict[str, Any], str]]:
    corpora = corpus_map(plan)
    rows: list[tuple[str, str, dict[str, Any], str]] = []
    for stage in STAGE_BY_LABEL:
        for corpus in ordered_corpora(plan, stage):
            phase = "edit"
            name = pilot_name(stage, corpus, phase)
            rows.append((name, stage, corpus, phase))
            phase = "lifecycle"
            name = pilot_name(stage, corpus, phase)
            rows.append((name, stage, corpus, phase))
    require(len(rows) == 16 and set(corpora) == {"generated", "numbered-list"},
            "native pilot child matrix changed")
    return rows


def validate_pilot_gate(plan: dict[str, Any]) -> dict[str, Any]:
    gate_path: Path | None = None
    gate_value: Any = None
    for raw in plan["pilot_gate"]["analysis_candidates"]:
        path = HERE / raw
        if path.is_file() and not path.is_symlink():
            gate_path = path
            gate_value = read_json(path)
            break
    require(gate_path is not None,
            "native pilot gate is absent; defer mechanism until pilot analysis explicitly passes")
    status = pilot_gate_status(gate_value)
    require(status == "pass",
            f"native pilot gate in {gate_path.name} is {status!r}; mechanism is deferred")
    identities: dict[tuple[str, str, str], dict[str, Any]] = {}
    receipts: list[dict[str, Any]] = []
    for name, stage, corpus, phase in pilot_receipt_names(plan):
        receipt_path = HERE / f"{name}.receipt.json"
        result_path = HERE / f"{name}.json"
        require(receipt_path.is_file() and result_path.is_file(),
                f"native pilot child is missing: {name}")
        receipt = read_json(receipt_path)
        require(receipt.get("exit_code") == 0, f"native pilot child failed: {name}")
        require(receipt.get("stage") == stage and receipt.get("lane") == "native"
                and receipt.get("phase") == phase and receipt.get("corpus_id") == corpus["id"],
                f"native pilot receipt identity differs: {name}")
        identity = report_identity(result_path)
        require(identity["case"] == case_name(corpus, phase),
                f"native pilot case differs: {name}")
        key = (stage, corpus["id"], phase)
        identities[key] = identity
        receipts.append({"name": name, "receipt": receipt_path.name,
                         "receipt_sha256": sha(receipt_path), "report": result_path.name,
                         "report_sha256": sha(result_path), "identity": identity})
    # Each stage must produce the same semantic result for both corpora across
    # the two ordered pairs.  This is the native-pilot gate's output parity.
    for corpus in corpus_map(plan):
        for phase in ("edit", "lifecycle"):
            values = [identities[(stage, corpus, phase)] for stage in STAGE_BY_LABEL]
            require(all(value == values[0] for value in values[1:]),
                    f"native pilot output parity failed for {corpus}/{phase}")
    return {"artifact": gate_path.name, "artifact_sha256": sha(gate_path),
            "status": "pass", "receipts": receipts,
            "child_count": len(receipts), "output_parity": True,
            "identities": {"/".join(key): value for key, value in sorted(identities.items())}}


def command_common(binary: str, plan: dict[str, Any], corpus: dict[str, Any], phase: str,
                   output: Path, samples: int, warmup: int) -> list[str]:
    command = ["taskset", "-c", str(plan["cpu"]), binary,
               "--warmup", str(warmup), "--samples", str(samples),
               "--case", case_name(corpus, phase), "--json", str(output),
               "--filesystem-root", str(plan["filesystem_root"])]
    if corpus["origin"] == "caller-named-real-file":
        command.extend(["--ooxml-file", corpus["path"]])
    return command


def run_process(command: list[str], stdout_path: Path, stderr_path: Path,
                env: dict[str, str]) -> tuple[int, int, str | None]:
    started = time.monotonic_ns()
    error: str | None = None
    try:
        with stdout_path.open("wb") as stdout, stderr_path.open("wb") as stderr:
            result = subprocess.run(command, cwd=REPO, stdout=stdout, stderr=stderr,
                                    env=env, check=False)
        code = result.returncode
    except OSError as exc:
        stderr_path.write_text(f"mechanism subprocess failed: {exc}\n", encoding="utf-8")
        code = 127
        error = str(exc)
    return code, time.monotonic_ns() - started, error


def base_receipt(name: str, kind: str, stage: str, corpus: dict[str, Any], phase: str,
                 plan: dict[str, Any], source: dict[str, Any], build: dict[str, Any],
                 command: list[str], before_name: str, after_name: str,
                 fixture_before_name: str, fixture_after_name: str,
                 before: dict[str, str], after: dict[str, str], fixture_before: dict[str, Any],
                 fixture_after: dict[str, Any], started: str, finished: str,
                 elapsed_ns: int, exit_code: int, error: str | None) -> dict[str, Any]:
    metadata = stage_metadata(stage)
    script_path = Path(__file__).resolve()
    result = {
        "schema_version": 1,
        "packet": plan["packet"],
        "kind": kind,
        "name": name,
        "stage": stage,
        "source_label": metadata["source"],
        "build_label": metadata["build"],
        "pair": metadata["pair"],
        "repeat": metadata["repeat"],
        "stage_order": metadata["order"],
        "corpus_id": corpus["id"],
        "corpus_label": corpus["label"],
        "corpus_origin": corpus["origin"],
        "phase": phase,
        "case": case_name(corpus, phase),
        "command": command,
        "started_utc": started,
        "finished_utc": finished,
        "wall_time_ns": elapsed_ns,
        "exit_code": exit_code,
        "subprocess_error": error,
        "cpu": plan["cpu"],
        "binary": build["binary"],
        "binary_sha256": build["binary_sha256"],
        "binary_bytes": build["binary_bytes"],
        "build_record": build["record_path"],
        "build_record_sha256": build["record_sha256"],
        "build_source_manifest": build["source_path"],
        "build_source_manifest_sha256": build["source_manifest_sha256"],
        "build_source_census_sha256": build["source_census_sha256"],
        "current_checkout_source": {
            "before_artifact": before_name,
            "after_artifact": after_name,
            "before_sha256": digest_json(before),
            "after_sha256": digest_json(after),
            "before_file_sha256": sha(HERE / before_name),
            "after_file_sha256": sha(HERE / after_name),
            "before_entry_count": len(before),
            "after_entry_count": len(after),
            "unchanged_during_child": before == after,
            "candidate_checkout": True,
            "binary_source_label": build["source_label"],
            "changed_paths_from_binary_source": source_diff(
                source[build["source_label"]], after),
        },
        "fixture": {
            "before_artifact": fixture_before_name,
            "after_artifact": fixture_after_name,
            "before": fixture_before,
            "after": fixture_after,
            "plan_path": corpus.get("path"),
            "plan_sha256": corpus.get("sha256"),
        },
        "source_delta": {
            "allowed_paths": [TARGET_DELTA],
            "baseline_manifest": source["baseline_sha256"],
            "candidate_manifest": source["candidate_sha256"],
            "binary_source_label": build["source_label"],
            "current_checkout_label": "candidate",
            "current_checkout_matches_candidate": after == source["candidate"],
        },
        "plan_sha256": sha(PLAN_PATH),
        "pilot_plan_sha256": sha(PILOT_PLAN_PATH),
        "script_sha256": sha(script_path),
        "constraints_sha256": sha(HERE / "constraints.json"),
        "environment": {key: os.environ.get(key)
                         for key in ("RUSTFLAGS", "LD_PRELOAD", "MALLOC_CONF", "GLIBC_TUNABLES",
                                     "LC_ALL", "LANG", "TZ")},
        "artifacts": {},
    }
    return result


def finish_artifacts(receipt: dict[str, Any], name: str, kind: str) -> None:
    paths = [path for path in artifact_paths(name, kind)
             if path.is_file() and not path.is_symlink()]
    # Wildcard candidates are not all emitted; only actual files are retained.
    receipt["artifacts"] = {path.name: sha(path) for path in sorted(paths)}


def run_child(*, plan: dict[str, Any], source: dict[str, Any], stage: str,
              corpus: dict[str, Any], phase: str, kind: str, order_index: int,
              gate: dict[str, Any]) -> None:
    build = stage_record(plan, stage)
    name = profile_name(stage, corpus) if kind == "profile" else rss_name(stage, corpus, phase)
    refuse_existing(name, kind)
    current = source_census()
    require(current == source["candidate"], f"{name}: current source is not candidate")
    before_name, _ = source_artifact(name, current, "source-before")
    fixture_before = fixture_binding(corpus)
    fixture_before_name, _ = fixture_artifact(name, fixture_before, "fixture-before")
    output_path = HERE / f"{name}.json"
    stdout_path = HERE / f"{name}.stdout"
    stderr_path = HERE / f"{name}.stderr"
    command = command_common(build["binary"], plan, corpus, phase, output_path,
                             plan["callgrind"]["samples"] if kind == "profile" else plan["rss"]["samples"],
                             plan["callgrind"]["warmup"] if kind == "profile" else plan["rss"]["warmup"])
    time_path: Path | None = None
    callgrind_path: Path | None = None
    if kind == "profile":
        callgrind_path = HERE / f"{name}.callgrind"
        command = ["taskset", "-c", str(plan["cpu"]), plan["callgrind"]["tool"],
                   *plan["callgrind"]["options"], f"--callgrind-out-file={callgrind_path}",
                   *command[3:]]
    else:
        time_path = HERE / f"{name}.time-v"
        # Keep the same affinity for time(1) and the measured process, while
        # storing verbose time output separately from benchmark stdout/stderr.
        command = ["taskset", "-c", str(plan["cpu"]), plan["rss"]["tool"], "-v",
                   "-o", str(time_path), *command[3:]]
    env = dict(os.environ)
    env.update({"LC_ALL": "C", "LANG": "C", "TZ": "UTC",
                "PERL_HASH_SEED": "0", "PERL_PERTURB_KEYS": "0"})
    started = utc_now()
    exit_code, elapsed_ns, error = run_process(command, stdout_path, stderr_path, env)
    after = source_census()
    after_name, _ = source_artifact(name, after, "source-after")
    fixture_after = fixture_binding(corpus)
    fixture_after_name, _ = fixture_artifact(name, fixture_after, "fixture-after")
    require(before == after, f"{name}: source changed during child")
    require(after == source["candidate"], f"{name}: checkout changed from candidate")
    require(fixture_before == fixture_after, f"{name}: fixture changed during child")
    require(sha(Path(build["binary"])) == build["binary_sha256"],
            f"{name}: binary changed during child")
    require((output_path.is_file() and not output_path.is_symlink()) or exit_code != 0,
            f"{name}: successful child emitted no report")
    if kind == "profile" and exit_code == 0:
        require(callgrind_path is not None and callgrind_path.is_file(),
                f"{name}: successful child emitted no Callgrind termination part")
    if kind == "rss" and exit_code == 0:
        require(time_path is not None and time_path.is_file(),
                f"{name}: successful child emitted no time -v report")
    receipt = base_receipt(name, kind, stage, corpus, phase, plan, source, build, command,
                           before_name, after_name, fixture_before_name, fixture_after_name,
                           before, after, fixture_before, fixture_after,
                           started, utc_now(), elapsed_ns, exit_code, error)
    receipt["order_index"] = order_index
    receipt["pilot_gate"] = {key: gate[key] for key in ("artifact", "artifact_sha256", "status")}
    if output_path.is_file() and not output_path.is_symlink():
        receipt["output_identity"] = report_identity(output_path)
    else:
        receipt["output_identity"] = None
    if kind == "profile":
        receipt["owner"] = OWNER
        receipt["owner_match"] = "exact"
        receipt["measured_parent"] = MEASURED_PARENT
        receipt["setup_parents"] = plan["callgrind"]["setup_parents"]
        receipt["expected_numbered_parts"] = plan["callgrind"]["expected_numbered_parts"]
        receipt["retains_zero_ir_program_termination"] = True
    else:
        time_text = read_text(time_path) if time_path and time_path.is_file() else ""
        match = PEAK_RE.search(time_text)
        receipt["time_verbose"] = {"path": time_path.name if time_path else None,
                                    "sha256": sha(time_path) if time_path and time_path.is_file() else None,
                                    "maximum_resident_set_size_kb": int(match.group(1)) if match else None}
        require(match is not None or exit_code != 0,
                f"{name}: time -v report lacks maximum RSS")
        receipt["rss_peak_policy"] = "per-child maximum RSS; never summed across phases or repeats"
    finish_artifacts(receipt, name, kind)
    receipt_path = HERE / f"{name}.receipt.json"
    write_json(receipt_path, receipt)
    require(exit_code == 0, f"{name}: child exited with {exit_code}; receipt retained")


def parity_key(receipt: dict[str, Any]) -> tuple[str, str, str]:
    return (str(receipt["corpus_id"]), str(receipt["phase"]), str(receipt["source_label"]))


def read_mechanism_receipts(plan: dict[str, Any], kind: str) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for stage in STAGE_BY_LABEL:
        for corpus in ordered_corpora(plan, stage):
            phases = ["edit"] if kind == "profile" else ["edit", "lifecycle"]
            for phase in phases:
                name = profile_name(stage, corpus) if kind == "profile" else rss_name(stage, corpus, phase)
                path = HERE / f"{name}.receipt.json"
                report = HERE / f"{name}.json"
                require(path.is_file() and report.is_file(), f"missing mechanism receipt/report: {name}")
                receipt = read_json(path)
                require(receipt.get("exit_code") == 0 and receipt.get("kind") == kind,
                        f"failed or mismatched mechanism receipt: {name}")
                require(receipt.get("output_identity") == report_identity(report),
                        f"output identity changed after receipt: {name}")
                rows.append(receipt)
    return rows


def write_parity(plan: dict[str, Any], gate: dict[str, Any], profile_rows: list[dict[str, Any]],
                 rss_rows: list[dict[str, Any]]) -> None:
    path = HERE / "mechanism-parity.json"
    require(not path.exists() and not path.is_symlink(), "refusing to replace mechanism-parity.json")
    all_rows = profile_rows + rss_rows
    comparisons: list[dict[str, Any]] = []
    for kind in ("profile", "rss"):
        rows = [row for row in all_rows if row["kind"] == kind]
        by_corpus_phase_source: dict[tuple[str, str, str], list[dict[str, Any]]] = {}
        for row in rows:
            by_corpus_phase_source.setdefault(parity_key(row), []).append(row)
        for key, values in sorted(by_corpus_phase_source.items()):
            require(len(values) == 2, f"{kind} parity matrix does not have two source stages: {key}")
            identities = [value["output_identity"] for value in values]
            require(identities[0] == identities[1], f"baseline/candidate output parity failed: {kind}/{key}")
            comparisons.append({"kind": kind, "key": key, "stages": [value["stage"] for value in values],
                                "identity": identities[0], "baseline_candidate_equal": True})
    pilot_matches: list[dict[str, Any]] = []
    for row in all_rows:
        pilot = HERE / f"{slug(row['stage'])}-native-{slug(row['corpus_id'])}-{slug(row['phase'])}.json"
        if pilot.is_file() and not pilot.is_symlink():
            identity = report_identity(pilot)
            require(identity == row["output_identity"],
                    f"mechanism/native pilot output parity failed: {row['name']}")
            pilot_matches.append({"mechanism": row["name"], "pilot": pilot.name,
                                  "pilot_sha256": sha(pilot), "equal": True})
    write_json(path, {
        "schema_version": 1,
        "packet": plan["packet"],
        "plan_sha256": sha(PLAN_PATH),
        "pilot_gate": gate,
        "callgrind_children": len(profile_rows),
        "rss_children": len(rss_rows),
        "comparisons": comparisons,
        "pilot_matches": pilot_matches,
        "all_mechanism_outputs_equal_between_stages": True,
        "native_pilot_matches_when_available": True,
        "rss_peaks_are_not_summed": True,
    })


def expected_jobs(plan: dict[str, Any], kind: str) -> list[tuple[str, dict[str, Any], str]]:
    jobs: list[tuple[str, dict[str, Any], str]] = []
    for stage in STAGE_BY_LABEL:
        for corpus in ordered_corpora(plan, stage):
            phases = ["edit"] if kind == "profile" else ["edit", "lifecycle"]
            for phase in phases:
                jobs.append((stage, corpus, phase))
    return jobs


def plan_only(plan: dict[str, Any]) -> None:
    for kind in ("profile", "rss"):
        for index, (stage, corpus, phase) in enumerate(expected_jobs(plan, kind)):
            print(json.dumps({"kind": kind, "order_index": index, "stage": stage,
                              "repeat": stage_metadata(stage)["repeat"],
                              "corpus_id": corpus["id"], "phase": phase,
                              "name": profile_name(stage, corpus) if kind == "profile"
                              else rss_name(stage, corpus, phase)}, sort_keys=True))


def run(plan: dict[str, Any], kinds: list[str]) -> None:
    constraints_check()
    gate = validate_pilot_gate(plan)
    source = source_state(plan)
    profile_rows: list[dict[str, Any]] = []
    rss_rows: list[dict[str, Any]] = []
    if "profile" in kinds:
        for index, (stage, corpus, phase) in enumerate(expected_jobs(plan, "profile")):
            run_child(plan=plan, source=source, stage=stage, corpus=corpus, phase=phase,
                      kind="profile", order_index=index, gate=gate)
        profile_rows = read_mechanism_receipts(plan, "profile")
    if "rss" in kinds:
        for index, (stage, corpus, phase) in enumerate(expected_jobs(plan, "rss")):
            run_child(plan=plan, source=source, stage=stage, corpus=corpus, phase=phase,
                      kind="rss", order_index=index, gate=gate)
        rss_rows = read_mechanism_receipts(plan, "rss")
    # If a caller runs only one lane, compare that lane.  A full run is the
    # normal handoff and leaves one stable parity witness for the analyzer.
    if profile_rows or rss_rows:
        write_parity(plan, gate, profile_rows, rss_rows)
    print(f"captured mechanism profile={len(profile_rows)} rss={len(rss_rows)}", flush=True)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--plan-only", action="store_true",
                        help="validate the frozen plan and print the child matrix")
    parser.add_argument("--kind", choices=("profile", "rss", "all"), default="all",
                        help="capture Callgrind profiles, RSS children, or both")
    args = parser.parse_args()
    try:
        plan = load_plan()
        if args.plan_only:
            plan_only(plan)
            return 0
        run(plan, ["profile", "rss"] if args.kind == "all" else [args.kind])
        return 0
    except (EvidenceError, OSError, KeyError, ValueError) as error:
        print(f"mechanism.py: error: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
