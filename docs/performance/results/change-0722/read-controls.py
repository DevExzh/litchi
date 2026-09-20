#!/usr/bin/env python3
"""Capture and adjudicate the two DOCX read-consumer controls.

The coordinator builds the frozen native binaries and runs ``pilot.py`` for a
stage first.  ``capture STAGE`` then runs exactly one fresh top-level child for
each control (the pinned filesystem control fans out to 500 inner children),
while the live checkout remains the candidate source.  This file does not
build, edit, restore, or clean a checkout.
"""

from __future__ import annotations

import argparse
import datetime as datetime_module
import hashlib
import importlib.util
import json
import math
import os
from pathlib import Path
import statistics
import subprocess
import sys
import time
from typing import Any

sys.dont_write_bytecode = True


HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
STAGES = (
    "baseline-A1", "candidate-B1", "candidate-B2", "baseline-A2",
    "baseline-A3", "candidate-B3", "candidate-B4", "baseline-A4",
)
PHASES = ("edit", "lifecycle")
ENVIRONMENT_KEYS = (
    "LC_ALL", "LANG", "TZ", "RUSTFLAGS", "LD_PRELOAD", "MALLOC_CONF",
    "GLIBC_TUNABLES", "PERL_HASH_SEED", "PERL_PERTURB_KEYS",
)
EXCLUDED_RESULT_FIELDS = {"elapsed_ns", "operation_metrics"}


_CUSTODY_SPEC = importlib.util.spec_from_file_location(
    "change0722_read_controls_custody", HERE / "custody.py"
)
if _CUSTODY_SPEC is None or _CUSTODY_SPEC.loader is None:
    raise RuntimeError(f"cannot load packet custody module: {HERE / 'custody.py'}")
_CUSTODY = importlib.util.module_from_spec(_CUSTODY_SPEC)
_CUSTODY_SPEC.loader.exec_module(_CUSTODY)


def fail(message: str) -> None:
    raise RuntimeError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def sha(path: Path) -> str:
    require(path.is_file() and not path.is_symlink(), f"missing file: {path}")
    return hashlib.sha256(path.read_bytes()).hexdigest()


def canonical(value: Any) -> bytes:
    return (json.dumps(value, sort_keys=True, separators=(",", ":")) + "\n").encode()


def digest(value: Any) -> str:
    return hashlib.sha256(canonical(value)).hexdigest()


def read(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing JSON: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(f"invalid JSON {path}: {error}")


def write_new(path: Path, value: Any) -> None:
    require(not path.exists() and not path.is_symlink(), f"refusing to replace {path}")
    path.write_text(json.dumps(value, indent=2) + "\n", encoding="utf-8")


def slug(value: str) -> str:
    return "".join(char if char.isalnum() or char in "-_" else "_" for char in value)


def utc_now() -> str:
    return datetime_module.datetime.now(datetime_module.timezone.utc).isoformat()


def resolve(raw: str) -> Path:
    path = Path(raw)
    return (REPO / path).resolve() if not path.is_absolute() else path.resolve()


def source_census() -> dict[str, str]:
    value = _CUSTODY.census()
    require(isinstance(value, dict) and value, "custody census is empty")
    return value


def load_plan() -> dict[str, Any]:
    plan = read(HERE / "read-controls-plan.json")
    require(plan.get("schema_version") == 1, "read-control plan schema changed")
    require(plan.get("revision") == read(HERE / "revision.json").get("revision"),
            "read-control revision is not the packet revision")
    require(plan.get("cpu") == 12, "read-control CPU changed")
    require(plan.get("native") == {"samples": 200, "warmup": 100},
            "native read-control sample plan changed")
    require(plan.get("binary_root") == "/home/zhuhe/code/litchi-0722-bin",
            "read-control binary root changed")
    require(plan.get("filesystem_root") == "/home/zhuhe/code/litchi-0722-fs",
            "read-control filesystem root changed")
    require(tuple(item["label"] for item in plan.get("stages", [])) == STAGES,
            "read-control stage order changed")
    interleave = plan.get("interleave")
    require(isinstance(interleave, dict)
            and interleave.get("primary_lane") == "native"
            and interleave.get("read_children_per_stage") == 2
            and interleave.get("top_level_read_invocations") == 16
            and interleave.get("pinned_filesystem_internal_children_per_invocation") == {
                "warmup": 100,
                "priming": 200,
                "measured": 200,
                "reported_total": 500,
            }
            and interleave.get("pinned_filesystem_internal_children_across_stages") == 4000,
            "read-control interleave or inner-child count changed")
    require(plan.get("environment") == {
        "LC_ALL": "C", "LANG": "C", "TZ": "UTC",
        "RUSTFLAGS": None, "LD_PRELOAD": None, "MALLOC_CONF": None,
        "GLIBC_TUNABLES": None, "PERL_HASH_SEED": "0",
        "PERL_PERTURB_KEYS": "0",
    }, "read-control child environment changed")
    require(plan.get("thresholds") == {
        "regression_percent": 3,
        "tail_flag_percent": 5,
        "repeat_drift_flag_percent": 5,
    }, "read-control thresholds changed")
    controls = plan.get("controls")
    require(isinstance(controls, list) and [c.get("id") for c in controls] == [
        "generated-medium-list-paragraphs", "pinned-media-eager-paragraph-count",
    ], "read-control list changed")
    for control in controls:
        require(control.get("case") in {
            "docx_semantic_list_paragraphs", "docx_file_eager_paragraph_count",
        }, f"unknown read-control case: {control.get('case')}")
        require(control.get("origin") in {
            "generated-harness-corpus", "pinned-filesystem-corpus",
        }, f"unknown read-control origin: {control.get('origin')}")
        if control.get("origin") == "generated-harness-corpus":
            require(control.get("semantic_shape") == "medium"
                    and control.get("filesystem_cache") is None
                    and control.get("expected_corpus") is None,
                    "generated medium control scope changed")
        else:
            expected = control.get("expected_corpus")
            require(control.get("semantic_shape") is None
                    and control.get("filesystem_cache") == "warm"
                    and isinstance(expected, dict)
                    and expected.get("generator") == "litchi-docx-source-edit-media-v1"
                    and expected.get("archive_sha256") ==
                    "a4a2e4921235a6da6b38e31d26ddcca1301909885e37330ab4f83ecc0c4e04f4"
                    and expected.get("archive_bytes") == 16793036,
                    "pinned filesystem control scope changed")
    return plan


def check_constraints(plan: dict[str, Any]) -> None:
    constraints_path = resolve(plan["constraints"])
    constraints = read(constraints_path)
    require(isinstance(constraints, dict) and constraints, "constraints are empty")
    for raw, expected in constraints.items():
        path = resolve(raw)
        require(path.is_file() and sha(path) == expected, f"constraint changed: {raw}")


def load_sources(plan: dict[str, Any]) -> tuple[dict[str, str], dict[str, str]]:
    baseline = read(resolve(plan["source_maps"]["baseline"]))
    candidate = read(resolve(plan["source_maps"]["candidate"]))
    require(isinstance(baseline, dict) and baseline, "baseline source map is invalid")
    require(isinstance(candidate, dict) and candidate, "candidate source map is invalid")
    changed = sorted(name for name in set(baseline) | set(candidate)
                     if baseline.get(name) != candidate.get(name))
    require(changed, "candidate source map is identical to baseline")
    main_plan = read(HERE / "plan.json")
    allowlist = set(main_plan.get("source_delta_allowlist", []))
    require(set(changed) <= allowlist,
            f"candidate source delta exceeds the reviewed allowlist: {changed}")
    return baseline, candidate


def final_source_state(baseline: dict[str, str], candidate: dict[str, str]) -> tuple[
    dict[str, str] | None, dict[str, Any] | None
]:
    """Accept a restored baseline only with the packet's explicit disposition."""
    final_path = HERE / "source-final.json"
    disposition_path = HERE / "disposition.json"
    require(final_path.exists() == disposition_path.exists(),
            "source-final.json and disposition.json must be created together")
    current = source_census()
    if not final_path.exists():
        require(current == candidate, "current checkout is not the frozen live candidate source")
        return None, None
    final_source = read(final_path)
    disposition = read(disposition_path)
    require(isinstance(final_source, dict) and final_source,
            "source-final.json is not a raw source map")
    require(isinstance(disposition, dict)
            and set(disposition) == {"retained", "final_source"},
            "disposition.json must contain retained and final_source")
    final_label = disposition.get("final_source")
    require(final_label in {"baseline", "candidate"},
            "disposition final_source must be baseline or candidate")
    expected = baseline if final_label == "baseline" else candidate
    require(final_source == expected and current == final_source,
            "current checkout does not match the explicit final source")
    require(isinstance(disposition.get("retained"), bool)
            and disposition["retained"] == (final_label == "candidate"),
            "disposition retained flag is inconsistent with final_source")
    return final_source, disposition


def cleanup_witnesses() -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for filename in ("cleanup.json", "cleanup-witness.json", "cleanup-receipt.json"):
        path = HERE / filename
        if not path.is_file() or path.is_symlink():
            continue
        value = read(path)

        def walk(item: Any) -> None:
            if isinstance(item, dict):
                raw = item.get("path")
                digest_value = item.get("sha256", item.get("binary_sha256"))
                size = item.get("bytes", item.get("binary_bytes"))
                if isinstance(raw, str) and isinstance(digest_value, str):
                    result.append({"path": raw, "sha256": digest_value, "bytes": size})
                for child in item.values():
                    walk(child)
            elif isinstance(item, list):
                for child in item:
                    walk(child)

        walk(value)
    return result


def binary_identity(binary: Path, expected_sha: str, expected_bytes: int) -> str:
    if binary.is_file() and not binary.is_symlink():
        require(binary.stat().st_size == expected_bytes and sha(binary) == expected_sha,
                f"live binary identity changed: {binary}")
        return "live-binary"
    for witness in cleanup_witnesses():
        if (resolve(str(witness["path"])) == binary.resolve()
                and witness.get("sha256") == expected_sha
                and witness.get("bytes") == expected_bytes):
            return "exact-cleanup-witness"
    fail(f"binary is absent without an exact cleanup witness: {binary}")


def build_info(plan: dict[str, Any], source_label: str, baseline: dict[str, str],
               candidate: dict[str, str], *, executable: bool) -> dict[str, Any]:
    path = resolve(plan["build_records"][source_label])
    records = read(path)
    require(isinstance(records, list), f"{path.name} is not a build record list")
    name = f"{source_label}-native"
    matches = [row for row in records if isinstance(row, dict)
               and Path(str(row.get("binary", ""))).name == name]
    require(len(matches) == 1, f"no unique native build record for {name}")
    record = matches[0]
    require(record.get("exit_code") == 0, f"{name} build failed")
    expected_command = [
        "cargo", "build", "--release", "--locked", "--manifest-path",
        "tools/perf-baseline/Cargo.toml", "--bin", "litchi-perf-baseline",
        "--target-dir", "/home/zhuhe/code/litchi-target-0722", "-j", "2",
    ]
    require(record.get("command") == expected_command, f"{name} build command changed")
    binary = Path(str(record.get("binary", ""))).resolve()
    expected_binary = Path(plan["binary_root"]) / name
    require(binary == expected_binary.resolve(), f"{name} binary path changed")
    expected_source = baseline if source_label == "baseline" else candidate
    manifest = resolve(plan["source_maps"][source_label])
    require(record.get("source_manifest_sha256") == sha(manifest),
            f"{name} source manifest binding changed")
    require(read(manifest) == expected_source, f"{name} source manifest differs from source pair")
    binary_sha = record.get("binary_sha256")
    binary_bytes = record.get("binary_bytes")
    require(isinstance(binary_sha, str) and len(binary_sha) == 64,
            f"{name} binary SHA is malformed")
    require(isinstance(binary_bytes, int) and binary_bytes > 0,
            f"{name} binary byte count is malformed")
    custody = binary_identity(binary, binary_sha, binary_bytes)
    if executable:
        require(binary.is_file() and not binary.is_symlink(), f"capture binary is unavailable: {binary}")
    return {
        "name": name, "binary": str(binary), "binary_sha256": binary_sha,
        "binary_bytes": binary_bytes, "binary_custody": custody,
        "record": record, "record_path": path.name, "record_sha256": sha(path),
        "source_manifest": manifest.name, "source_manifest_sha256": sha(manifest),
        "source": expected_source,
    }


def stage_meta(plan: dict[str, Any], stage: str) -> dict[str, Any]:
    require(stage in STAGES, f"unknown stage {stage!r}")
    return next(item for item in plan["stages"] if item["label"] == stage)


def control_name(stage: str, control: dict[str, Any]) -> str:
    return f"{slug(stage)}-read-{slug(control['id'])}"


def primary_name(stage: str, corpus_id: str, phase: str) -> str:
    return f"{stage}-native-{corpus_id}-{phase}"


def primary_case(corpus_id: str, phase: str) -> str:
    prefix = ("docx_ordinary_save_" if corpus_id == "generated"
              else "docx_real_file_ordinary_save_")
    return prefix + phase


def primary_command(plan: dict[str, Any], build: dict[str, Any],
                    case: str, name: str) -> list[str]:
    command = [
        "taskset", "-c", str(plan["cpu"]), build["binary"],
        "--warmup", str(plan["native"]["warmup"]),
        "--samples", str(plan["native"]["samples"]), "--case", case,
        "--json", str(HERE / f"{name}.json"),
        "--filesystem-root", plan["filesystem_root"],
    ]

    if case.startswith("docx_real_file_ordinary_save_"):
        primary_plan = read(HERE / "plan.json")
        fixture = next(c for c in primary_plan["corpora"] if c["id"] == "numbered-list")
        command += ["--ooxml-file", fixture["path"]]
    return command


def primary_interleave(plan: dict[str, Any], stage: str,
                       candidate_source: dict[str, str], build: dict[str, Any]) -> dict[str, Any]:
    """Validate the already completed native primary stage before read children."""
    latest: datetime_module.datetime | None = None
    receipts: list[str] = []
    expected_source_digest = digest(candidate_source)
    primary_script = resolve(plan["primary_script"])
    primary_plan = HERE / "plan.json"
    constraints = resolve(plan["constraints"])
    expected_metadata = stage_meta(plan, stage)
    primary_order = [(corpus_id, phase)
                     for phase in (list(PHASES)[::-1]
                                   if expected_metadata["order"] == "reverse" else PHASES)
                     for corpus_id in (("generated", "numbered-list")[::-1]
                                       if expected_metadata["order"] == "reverse"
                                       else ("generated", "numbered-list"))]
    for corpus_id, control_id in (("generated", "generated-medium-list-paragraphs"),
                                  ("numbered-list", "pinned-media-eager-paragraph-count")):
        for phase in PHASES:
            name = primary_name(stage, corpus_id, phase)
            path = HERE / f"{name}.receipt.json"
            receipt = read(path)
            require(receipt.get("packet") == plan["primary_packet"],
                    f"{name}: wrong primary packet")
            require(receipt.get("stage") == stage and receipt.get("lane") == "native",
                    f"{name}: primary stage/lane changed")
            require(receipt.get("corpus_id") == corpus_id and receipt.get("phase") == phase,
                    f"{name}: primary control identity changed")
            require(receipt.get("case") == primary_case(corpus_id, phase),
                    f"{name}: primary case changed")
            require(receipt.get("command") == primary_command(
                plan, build, primary_case(corpus_id, phase), name),
                    f"{name}: primary command changed")
            require(receipt.get("source_label") == expected_metadata["source"]
                    and receipt.get("build_label") == expected_metadata["build"]
                    and receipt.get("pair") == expected_metadata["pair"]
                    and receipt.get("cycle") == expected_metadata["cycle"]
                    and receipt.get("stage_order") == expected_metadata["order"]
                    and receipt.get("lane") == "native"
                    and receipt.get("samples") == plan["native"]["samples"]
                    and receipt.get("warmup") == plan["native"]["warmup"]
                    and receipt.get("cpu") == plan["cpu"]
                    and receipt.get("fresh_child_process") is True,
                    f"{name}: primary execution metadata changed")
            require(receipt.get("stage_order_index") == primary_order.index((corpus_id, phase)),
                    f"{name}: primary stage order index changed")
            require(receipt.get("exit_code") == 0, f"{name}: primary child failed")
            require(receipt.get("plan_sha256") == sha(primary_plan)
                    and receipt.get("script_sha256") == sha(primary_script)
                    and receipt.get("constraints_sha256") == sha(constraints),
                    f"{name}: primary packet binding changed")
            for key in ("binary_path", "binary_sha256", "binary_bytes"):
                require(receipt.get(key) == build[{
                    "binary_path": "binary", "binary_sha256": "binary_sha256",
                    "binary_bytes": "binary_bytes",
                }[key]], f"{name}: primary binary binding changed")
            live = receipt.get("live_checkout_source", receipt.get("current_checkout_source"))
            require(isinstance(live, dict) and live.get("unchanged_during_child") is True,
                    f"{name}: primary source custody is missing")
            require(live.get("source_census_sha256") == expected_source_digest
                    and live.get("before_sha256") == expected_source_digest
                    and live.get("after_sha256") == expected_source_digest
                    and live.get("recensus_before") is True
                    and live.get("recensus_after") is True
                    and isinstance(live.get("relation_before"), dict)
                    and live["relation_before"].get("changed_paths") == []
                    and isinstance(live.get("relation_after"), dict)
                    and live["relation_after"].get("changed_paths") == [],
                    f"{name}: primary did not run on the candidate checkout")
            artifacts = receipt.get("artifacts")
            require(isinstance(artifacts, dict)
                    and set(artifacts) == {f"{name}.json", f"{name}.stdout", f"{name}.stderr"},
                    f"{name}: primary artifact inventory changed")
            for filename, expected in artifacts.items():
                artifact = HERE / filename
                require(artifact.is_file() and sha(artifact) == expected,
                        f"{name}: primary artifact changed: {filename}")
            end = receipt.get("end_utc")
            require(isinstance(end, str), f"{name}: primary end time is missing")
            ended = datetime_module.datetime.fromisoformat(end)
            latest = ended if latest is None or ended > latest else latest
            receipts.append(name)
    require(latest is not None, f"{stage}: primary native stage is empty")
    return {"children": receipts, "latest_end_utc": latest.isoformat()}


def expected_command(plan: dict[str, Any], control: dict[str, Any], build: dict[str, Any],
                     output: Path) -> list[str]:
    command = [
        "taskset", "-c", str(plan["cpu"]), build["binary"],
        "--warmup", str(plan["native"]["warmup"]),
        "--samples", str(plan["native"]["samples"]),
        "--case", control["case"],
    ]
    if control["semantic_shape"] is not None:
        command += ["--semantic-shape", control["semantic_shape"]]
    if control["filesystem_cache"] is not None:
        command += ["--filesystem-cache", control["filesystem_cache"],
                    "--filesystem-root", plan["filesystem_root"]]
    command += ["--json", str(output)]
    return command


def child_environment(plan: dict[str, Any]) -> tuple[dict[str, str], dict[str, str | None]]:
    environment = os.environ.copy()
    for key, value in plan["environment"].items():
        if value is None:
            environment.pop(key, None)
        else:
            environment[key] = value
    observed = {key: environment.get(key) for key in ENVIRONMENT_KEYS}
    require(observed == plan["environment"], "read-control environment could not be frozen")
    return environment, observed


def run_control(plan: dict[str, Any], stage: str, control: dict[str, Any],
                build: dict[str, Any], candidate_source: dict[str, str],
                primary: dict[str, Any]) -> None:
    name = control_name(stage, control)
    output = HERE / f"{name}.json"
    stdout_path = HERE / f"{name}.stdout"
    stderr_path = HERE / f"{name}.stderr"
    receipt_path = HERE / f"{name}.receipt.json"
    for path in (output, stdout_path, stderr_path, receipt_path):
        require(not path.exists() and not path.is_symlink(), f"refusing to replace {path}")
    check_constraints(plan)
    before = source_census()
    expected_source_digest = digest(candidate_source)
    require(before == candidate_source, f"{name}: live checkout is not candidate source")
    command = expected_command(plan, control, build, output)
    environment, observed_environment = child_environment(plan)
    primary_end = datetime_module.datetime.fromisoformat(primary["latest_end_utc"])
    started = utc_now()
    require(datetime_module.datetime.fromisoformat(started) >= primary_end,
            f"{name}: read child started before the primary native stage ended")
    tick = time.monotonic()
    with stdout_path.open("xb") as stdout, stderr_path.open("xb") as stderr:
        result = subprocess.run(command, cwd=REPO, env=environment,
                                stdout=stdout, stderr=stderr)
    elapsed_seconds = time.monotonic() - tick
    after = source_census()
    require(before == after == candidate_source, f"{name}: live source changed during child")
    require(sha(Path(build["binary"])) == build["binary_sha256"],
            f"{name}: frozen binary changed during child")
    require(output.is_file() and not output.is_symlink(), f"{name}: report is missing")
    require(stdout_path.read_bytes() == b"", f"{name}: stdout was not empty")
    require(stderr_path.read_bytes() == b"", f"{name}: stderr was not empty")
    artifacts = {path.name: {"sha256": sha(path), "bytes": path.stat().st_size}
                 for path in (output, stdout_path, stderr_path)}
    metadata = stage_meta(plan, stage)
    live = {
        "manifest": "source-candidate.json",
        "manifest_sha256": sha(resolve(plan["source_maps"]["candidate"])),
        "source_census_sha256": expected_source_digest,
        "source_entry_count": len(candidate_source),
        "before_sha256": digest(before), "after_sha256": digest(after),
        "before_entry_count": len(before), "after_entry_count": len(after),
        "changed_paths": [], "unchanged_during_child": True,
    }
    receipt = {
        "schema_version": 1, "packet": plan["packet"], "name": name,
        "stage": stage, "source_label": metadata["source"],
        "build_label": metadata["build"], "pair": metadata["pair"],
        "cycle": metadata["cycle"], "order": metadata["order"],
        "control_id": control["id"], "control_label": control["label"],
        "origin": control["origin"], "case": control["case"],
        "samples": plan["native"]["samples"], "warmup": plan["native"]["warmup"],
        "fresh_child_process": True, "command": command,
        "start_utc": started, "end_utc": utc_now(),
        "seconds": elapsed_seconds, "exit_code": result.returncode,
        "cpu": plan["cpu"], "binary_path": build["binary"],
        "binary_sha256": build["binary_sha256"], "binary_bytes": build["binary_bytes"],
        "build_record": build["record_path"], "build_record_sha256": build["record_sha256"],
        "build_source_manifest": build["source_manifest"],
        "build_source_manifest_sha256": build["source_manifest_sha256"],
        "binary_source": {"manifest": build["source_manifest"],
                           "manifest_sha256": build["source_manifest_sha256"],
                           "source_census_sha256": digest(build["source"]),
                           "source_entry_count": len(build["source"])},
        "live_checkout_source": live, "current_checkout_source": live,
        "expected_corpus": control.get("expected_corpus"),
        "primary_native": primary,
        "plan_sha256": sha(HERE / "read-controls-plan.json"),
        "script_sha256": sha(Path(__file__).resolve()),
        "constraints_sha256": sha(resolve(plan["constraints"])),
        "environment": observed_environment, "artifacts": artifacts,
        "stdout_stderr": {
            "stdout": {"empty": True, **artifacts[f"{name}.stdout"]},
            "stderr": {"empty": True, **artifacts[f"{name}.stderr"]},
        },
    }
    write_new(receipt_path, receipt)
    require(result.returncode == 0, f"{name}: benchmark failed with exit {result.returncode}")
    print(f"{name} passed", flush=True)


def capture(plan: dict[str, Any], stage: str) -> None:
    baseline, candidate = load_sources(plan)
    metadata = stage_meta(plan, stage)
    require(source_census() == candidate, f"{stage}: capture requires live candidate source")
    build = build_info(plan, metadata["build"], baseline, candidate, executable=True)
    primary = primary_interleave(plan, stage, candidate, build)
    for control in plan["controls"]:
        run_control(plan, stage, control, build, candidate, primary)
    print(f"stage {stage} read controls complete ({len(plan['controls'])} children)", flush=True)


def finite_positive(value: Any, label: str) -> None:
    require(isinstance(value, (int, float)) and not isinstance(value, bool)
            and math.isfinite(float(value)) and float(value) > 0, f"{label} is not positive")


def stats(samples: list[int]) -> dict[str, Any]:
    require(len(samples) == 200, f"expected 200 retained samples, got {len(samples)}")
    require(all(isinstance(value, int) and not isinstance(value, bool) and value > 0
                for value in samples), "elapsed sample vector contains an invalid value")
    ordered = sorted(samples)
    midpoint = ordered[(len(ordered) - 1) // 2] // 2 + ordered[len(ordered) // 2] // 2
    midpoint += (ordered[(len(ordered) - 1) // 2] % 2 + ordered[len(ordered) // 2] % 2) // 2

    def nearest_rank(percentile: int) -> int:
        return ordered[min((percentile * len(ordered) + 99) // 100 - 1, len(ordered) - 1)]

    return {
        "count": len(samples), "min": ordered[0], "p50": midpoint,
        "p95": nearest_rank(95), "p99": nearest_rank(99), "max": ordered[-1],
        "mean": statistics.fmean(samples), "samples_sha256": digest(samples),
    }


def check_report(plan: dict[str, Any], control: dict[str, Any], build: dict[str, Any],
                 path: Path) -> tuple[dict[str, Any], dict[str, Any], dict[str, Any]]:
    report = read(path)
    require(report.get("schema_version") == 1, f"{path.name}: report schema changed")
    tool = report.get("tool")
    require(isinstance(tool, dict) and tool.get("name") == "litchi-perf-baseline"
            and tool.get("instrumentation") == "none", f"{path.name}: tool identity changed")
    binary = report.get("binary_identity")
    require(isinstance(binary, dict) and binary.get("path") == build["binary"]
            and binary.get("binary_sha256") == build["binary_sha256"]
            and binary.get("binary_bytes") == build["binary_bytes"],
            f"{path.name}: report binary identity changed")
    configuration = report.get("configuration")
    require(isinstance(configuration, dict)
            and configuration.get("samples_per_case") == 200
            and configuration.get("warmup_iterations_per_case") == 100
            and configuration.get("cases") == [control["case"]],
            f"{path.name}: selected configuration changed")
    if control["origin"] == "generated-harness-corpus":
        require(configuration.get("semantic_shapes") == ["medium"]
                and configuration.get("filesystem_root_selected") is False,
                f"{path.name}: generated semantic scope changed")
        require("filesystem_evidence" not in report, f"{path.name}: generated filesystem evidence appeared")
    else:
        require(configuration.get("filesystem_cache_states") == ["warm"]
                and configuration.get("filesystem_root_selected") is True
                and configuration.get("filesystem_process_isolated") is True,
                f"{path.name}: real filesystem scope changed")
    results = report.get("results")
    require(isinstance(results, list) and len(results) == 1, f"{path.name}: result count changed")
    result = results[0]
    require(result.get("case") == control["case"] and result.get("sink") is None,
            f"{path.name}: result identity changed")
    require(result.get("output_sha256") is None, f"{path.name}: read control produced an output hash")
    corpus = result.get("corpus")
    require(isinstance(corpus, dict), f"{path.name}: corpus identity is missing")
    if control["origin"] == "generated-harness-corpus":
        require(corpus.get("generator") == "litchi-docx-semantic-v1"
                and corpus.get("package_format") == "DOCX/OPC/ZIP"
                and corpus.get("shape") == "medium"
                and corpus.get("entry_count") == 200,
                f"{path.name}: generated medium corpus scope changed")
    expected_corpus = control.get("expected_corpus")
    if expected_corpus is not None:
        require(isinstance(expected_corpus, dict), f"{path.name}: expected corpus binding is malformed")
        require(all(corpus.get(key) == value for key, value in expected_corpus.items()),
                f"{path.name}: corpus binding changed")
    if control["origin"] == "pinned-filesystem-corpus":
        require(result.get("cache_state") == "warm", f"{path.name}: cache state changed")
        evidence = report.get("filesystem_evidence")
        require(isinstance(evidence, list) and len(evidence) == 1
                and evidence[0].get("case") == control["case"]
                and evidence[0].get("cache_states") == ["warm"]
                and evidence[0].get("sample_count") == 200
                and evidence[0].get("fresh_child_per_sample") is True
                and len(evidence[0].get("samples", [])) == 200,
                f"{path.name}: filesystem evidence scope changed")
        require(evidence[0].get("corpus") == result.get("corpus"),
                f"{path.name}: filesystem corpus identity changed")
    elapsed = result.get("elapsed_ns")
    require(isinstance(elapsed, dict) and elapsed.get("unit") == "ns", f"{path.name}: elapsed schema changed")
    samples = elapsed.get("samples")
    require(isinstance(samples, list), f"{path.name}: raw sample vector is missing")
    sample_order = elapsed.get("sample_order")
    require(isinstance(sample_order, list) and sorted(sample_order) == list(range(200)),
            f"{path.name}: sample order is not a permutation")
    recomputed = stats(samples)
    require(elapsed.get("p50") == recomputed["p50"] and elapsed.get("min") == recomputed["min"]
            and elapsed.get("max") == recomputed["max"],
            f"{path.name}: report summary does not match its raw vector")
    finite_positive(elapsed.get("mean"), f"{path.name}: report mean")
    require(math.isclose(float(elapsed["mean"]), recomputed["mean"],
                         rel_tol=1e-9, abs_tol=0.01),
            f"{path.name}: report mean does not match its raw vector")
    normalized = {key: value for key, value in result.items()
                  if key not in EXCLUDED_RESULT_FIELDS}
    return report, normalized, recomputed


def percent(candidate: float, baseline: float) -> float:
    require(baseline > 0, "baseline statistic must be positive")
    return (candidate / baseline - 1.0) * 100.0


def analyze(plan: dict[str, Any], output: Path) -> None:
    baseline_source, candidate_source = load_sources(plan)
    _final_source, disposition = final_source_state(baseline_source, candidate_source)
    check_constraints(plan)
    builds = {
        label: build_info(plan, label, baseline_source, candidate_source, executable=False)
        for label in ("baseline", "candidate")
    }
    jobs: dict[tuple[str, str], dict[str, Any]] = {}
    normalized: dict[tuple[str, str], dict[str, Any]] = {}
    for stage in STAGES:
        metadata = stage_meta(plan, stage)
        build = builds[metadata["build"]]
        primary = primary_interleave(plan, stage, candidate_source, build)
        for control in plan["controls"]:
            name = control_name(stage, control)
            report_path = HERE / f"{name}.json"
            receipt = read(HERE / f"{name}.receipt.json")
            require(receipt.get("packet") == plan["packet"]
                    and receipt.get("stage") == stage
                    and receipt.get("control_id") == control["id"]
                    and receipt.get("exit_code") == 0,
                    f"{name}: receipt identity changed")
            require(receipt.get("name") == name
                    and receipt.get("source_label") == metadata["source"]
                    and receipt.get("build_label") == metadata["build"]
                    and receipt.get("pair") == metadata["pair"]
                    and receipt.get("cycle") == metadata["cycle"]
                    and receipt.get("order") == metadata["order"]
                    and receipt.get("control_label") == control["label"]
                    and receipt.get("origin") == control["origin"]
                    and receipt.get("samples") == plan["native"]["samples"]
                    and receipt.get("warmup") == plan["native"]["warmup"]
                    and receipt.get("fresh_child_process") is True
                    and receipt.get("cpu") == plan["cpu"],
                    f"{name}: receipt execution metadata changed")
            require(receipt.get("command") == expected_command(plan, control, build, report_path),
                    f"{name}: command changed")
            require(receipt.get("binary_path") == build["binary"]
                    and receipt.get("binary_sha256") == build["binary_sha256"]
                    and receipt.get("binary_bytes") == build["binary_bytes"],
                    f"{name}: receipt binary binding changed")
            require(receipt.get("plan_sha256") == sha(HERE / "read-controls-plan.json")
                    and receipt.get("script_sha256") == sha(Path(__file__).resolve())
                    and receipt.get("constraints_sha256") == sha(resolve(plan["constraints"]))
                    and receipt.get("environment") == plan["environment"],
                    f"{name}: script or plan binding changed")
            expected_binary_source = {
                "manifest": build["source_manifest"],
                "manifest_sha256": build["source_manifest_sha256"],
                "source_census_sha256": digest(build["source"]),
                "source_entry_count": len(build["source"]),
            }
            require(receipt.get("binary_source") == expected_binary_source
                    and receipt.get("build_record") == build["record_path"]
                    and receipt.get("build_record_sha256") == build["record_sha256"]
                    and receipt.get("build_source_manifest") == build["source_manifest"]
                    and receipt.get("build_source_manifest_sha256") == build["source_manifest_sha256"],
                    f"{name}: build/source custody changed")
            require(receipt.get("expected_corpus") == control.get("expected_corpus"),
                    f"{name}: corpus plan binding changed")
            require(receipt.get("primary_native") == primary,
                    f"{name}: interleave witness changed")
            started = receipt.get("start_utc")
            ended = receipt.get("end_utc")
            require(isinstance(started, str) and isinstance(ended, str),
                    f"{name}: read timing custody is missing")
            started_at = datetime_module.datetime.fromisoformat(started)
            ended_at = datetime_module.datetime.fromisoformat(ended)
            require(ended_at >= started_at
                    and started_at >= datetime_module.datetime.fromisoformat(
                        primary["latest_end_utc"]),
                    f"{name}: read child did not follow the primary stage")
            live = receipt.get("live_checkout_source", receipt.get("current_checkout_source"))
            expected_source_digest = digest(candidate_source)
            require(isinstance(live, dict) and live.get("unchanged_during_child") is True
                    and live.get("source_census_sha256") == expected_source_digest
                    and live.get("before_sha256") == expected_source_digest
                    and live.get("after_sha256") == expected_source_digest
                    and live.get("changed_paths") == [],
                    f"{name}: candidate source custody changed")
            artifacts = receipt.get("artifacts")
            require(isinstance(artifacts, dict)
                    and set(artifacts) == {f"{name}.json", f"{name}.stdout", f"{name}.stderr"},
                    f"{name}: artifact inventory changed")
            for filename, details in artifacts.items():
                artifact = HERE / filename
                require(artifact.is_file() and details == {
                    "sha256": sha(artifact), "bytes": artifact.stat().st_size,
                }, f"{name}: artifact changed: {filename}")
            require(receipt.get("stdout_stderr") == {
                "stdout": {"empty": True, **artifacts[f"{name}.stdout"]},
                "stderr": {"empty": True, **artifacts[f"{name}.stderr"]},
            },
                    f"{name}: stdout/stderr emptiness was not witnessed")
            report, stable, recomputed = check_report(plan, control, build, report_path)
            jobs[(stage, control["id"])] = {
                "stage": stage, "control": control, "receipt": receipt,
                "report": report, "stats": recomputed,
            }
            normalized[(stage, control["id"])] = stable
    require(len(jobs) == len(STAGES) * len(plan["controls"]),
            f"expected {len(STAGES) * len(plan['controls'])} read children, got {len(jobs)}")

    parity: dict[str, Any] = {}
    parity_failures: list[str] = []
    for control in plan["controls"]:
        cid = control["id"]
        reference = normalized[("baseline-A1", cid)]
        key = cid
        parity[key] = {
            "reference": "baseline-A1",
            "normalized_sha256": digest(reference),
            "reports": {},
        }
        for stage in STAGES:
            value = normalized[(stage, cid)]
            parity[key]["reports"][stage] = digest(value)
            if value != reference:
                parity_failures.append(f"{stage}/{cid}")
    parity_pass = not parity_failures

    pair_stage = {
        "pair-1": ("baseline-A1", "candidate-B1"),
        "pair-2": ("baseline-A2", "candidate-B2"),
        "pair-3": ("baseline-A3", "candidate-B3"),
        "pair-4": ("baseline-A4", "candidate-B4"),
    }
    comparisons: dict[str, Any] = {}
    hard_gates: list[dict[str, Any]] = []
    tail_flags: list[dict[str, Any]] = []
    for pair, (baseline_stage, candidate_stage) in pair_stage.items():
        comparisons[pair] = {}
        for control in plan["controls"]:
            cid = control["id"]
            baseline_stats = jobs[(baseline_stage, cid)]["stats"]
            candidate_stats = jobs[(candidate_stage, cid)]["stats"]
            metrics: dict[str, Any] = {}
            for metric in ("p50", "mean", "p95", "p99", "max"):
                delta = percent(float(candidate_stats[metric]), float(baseline_stats[metric]))
                metrics[metric] = {
                    "baseline": baseline_stats[metric],
                    "candidate": candidate_stats[metric],
                    "delta_percent": delta,
                }
                if metric in {"p50", "mean"}:
                    hard_gates.append({
                        "pair": pair, "control_id": cid, "metric": metric,
                        "threshold_percent": plan["thresholds"]["regression_percent"],
                        "observed_delta_percent": delta,
                        "pass": delta <= plan["thresholds"]["regression_percent"],
                    })
                elif delta > plan["thresholds"]["tail_flag_percent"]:
                    tail_flags.append({
                        "pair": pair, "control_id": cid, "metric": metric,
                        "delta_percent": delta, "flag_over_5_percent": True,
                    })
            comparisons[pair][cid] = {
                "baseline_stage": baseline_stage,
                "candidate_stage": candidate_stage,
                "native": metrics,
            }

    repeat_flags: list[dict[str, Any]] = []
    repeat_groups = plan["statistics"]["repeat_groups"]
    for left_stage, right_stage in repeat_groups:
        for control in plan["controls"]:
            cid = control["id"]
            left = jobs[(left_stage, cid)]["stats"]
            right = jobs[(right_stage, cid)]["stats"]
            for metric in ("p50", "mean", "p95", "p99", "max"):
                denominator = min(float(left[metric]), float(right[metric]))
                drift = abs(float(right[metric]) - float(left[metric])) * 100.0 / denominator
                if drift > plan["thresholds"]["repeat_drift_flag_percent"]:
                    repeat_flags.append({
                        "stages": [left_stage, right_stage], "control_id": cid,
                        "metric": metric, "drift_percent": drift,
                        "flag_over_5_percent": True,
                    })

    raw_statistics = {
        f"{stage}/{cid}": jobs[(stage, cid)]["stats"]
        for stage in STAGES for cid in (control["id"] for control in plan["controls"])
    }
    all_gates_pass = bool(hard_gates) and all(item["pass"] for item in hard_gates)
    accepted = parity_pass and all_gates_pass
    if disposition is not None:
        if disposition["final_source"] == "baseline":
            require(disposition["retained"] is False and not accepted,
                    "baseline final disposition requires an explicit rejected decision")
        else:
            require(disposition["retained"] is True and accepted,
                    "candidate final disposition requires an accepted decision")
    result = {
        "schema_version": 1,
        "packet": plan["packet"],
        "revision": plan["revision"],
        "scope": {
            "controls": [control["id"] for control in plan["controls"]],
            "stages": list(STAGES), "children": len(jobs),
            "top_level_read_invocations": len(jobs),
            "native_samples": plan["native"]["samples"],
            "capture_live_source": "candidate",
            "final_source": (None if disposition is None else disposition["final_source"]),
            "pinned_filesystem_internal_children_per_invocation": {
                "warmup": plan["native"]["warmup"],
                "priming": plan["native"]["samples"],
                "measured": plan["native"]["samples"],
                "reported_total": plan["native"]["warmup"] + (2 * plan["native"]["samples"]),
            },
            "pinned_filesystem_internal_children_across_stages": (
                len(STAGES) * (plan["native"]["warmup"] + (2 * plan["native"]["samples"]))
            ),
        },
        "raw_sample_statistics": raw_statistics,
        "normalized_output_parity": {
            "verified": parity_pass,
            "keys": parity,
            "failures": parity_failures,
            "excluded_fields": plan["statistics"]["excluded_from_normalized_identity"],
            "note": "Only result.elapsed_ns and result.operation_metrics are excluded; case, cache state, corpus metadata, sink, output identity, and every other deterministic result field remain compared.",
        },
        "paired_comparisons": comparisons,
        "review_flags": {
            "native_tail_regressions_over_5_percent": tail_flags,
            "repeat_drift_over_5_percent": repeat_flags,
            "note": "p95, p99, and max tail deltas plus every within-variant repeat drift are retained; no samples or pairs are omitted.",
        },
        "decision": {
            "hard_gates": hard_gates,
            "all_hard_gates_pass": all_gates_pass,
            "deterministic_output_parity_pass": parity_pass,
            "accepted": accepted,
            "acceptance_rule": "accept only when exact normalized identity parity and every paired p50/mean non-regression gate pass",
        },
        "verification": {
            "raw_elapsed_statistics_recomputed": True,
            "all_retained_sample_count": 200,
            "allocation_fields_compared": False,
            "timing_fields_compared_for_identity": False,
            "primary_native_interleaving_verified": True,
            "candidate_source_unchanged_for_every_read_child": True,
            "final_source_matches_checkout": True,
            "final_source_label": (None if disposition is None else disposition["final_source"]),
            "explicit_disposition_required_for_restored_baseline": True,
            "stdout_and_stderr_empty_for_every_read_child": True,
        },
        "binary_custody": {
            label: {"path": builds[label]["binary"], "sha256": builds[label]["binary_sha256"],
                    "bytes": builds[label]["binary_bytes"],
                    "validation": "validated-live-or-exact-cleanup-witness"}
            for label in ("baseline", "candidate")
        },
        "source_manifests": {
            label: {"path": plan["source_maps"][label],
                    "sha256": sha(resolve(plan["source_maps"][label])),
                    "entry_count": len(value)}
            for label, value in (("baseline", baseline_source), ("candidate", candidate_source))
        },
        "plan_sha256": sha(HERE / "read-controls-plan.json"),
        "capture_script_sha256": sha(Path(__file__).resolve()),
        "constraints_sha256": sha(resolve(plan["constraints"])),
    }
    write_new(output, result)
    print(f"verified {len(jobs)} read children; wrote {output}")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    subparsers = parser.add_subparsers(dest="command", required=True)
    capture_parser = subparsers.add_parser("capture", help="capture one interleaved stage")
    capture_parser.add_argument("stage", choices=STAGES)
    analyze_parser = subparsers.add_parser("analyze", help="validate and compare all captures")
    analyze_parser.add_argument("--output", default=str(HERE / "read-controls-analysis.json"))
    args = parser.parse_args()
    try:
        plan = load_plan()
        if args.command == "capture":
            capture(plan, args.stage)
        else:
            analyze(plan, Path(args.output).resolve())
    except (OSError, RuntimeError, ValueError, KeyError, TypeError) as error:
        print(f"read-controls failed: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
