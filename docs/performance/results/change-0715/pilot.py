#!/usr/bin/env python3
"""Capture one frozen stage of the 0715 DOCX publication-collection pilot.

The coordinator prepares the two source manifests and the four frozen binary
identities.  This driver only runs one stage/lane at a time.  It keeps the
candidate checkout in place while a baseline binary is run, which is why the
receipt records both the retained binary source and the live candidate source.
It never builds Cargo artifacts, edits production files, changes a checkout,
or replaces an evidence artifact.

Public helpers used by ``analyze_pilot.py`` are ``load_plan``,
``load_source_pair``, ``build_info``, ``ordered_jobs``, ``expected_command``,
``fixture_info``, and ``source_census``.  The command-line API is::

    python3 pilot.py baseline-A1 native
    python3 pilot.py candidate-B1 allocator

The two reverse stages must be run after their forward counterparts.  Across
both lanes the four stages produce 32 children: two corpora, two phases, two
ABBA pairs, and native plus allocator instrumentation.
"""

from __future__ import annotations

import argparse
import datetime as _datetime
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import subprocess
import sys
import time
from typing import Any, NoReturn


HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
PHASES = ("counting_publish", "lifecycle")
LANES = ("native", "allocator")
LANE_ALIASES = {"alloc": "allocator", "allocation": "allocator"}
STAGE_LABELS = ("baseline-A1", "candidate-B1", "candidate-B2", "baseline-A2")
STAGE_METADATA = {
    "baseline-A1": {"source": "baseline", "build": "baseline", "pair": "pair-1",
                    "order": "forward", "repeat": 1},
    "candidate-B1": {"source": "candidate", "build": "candidate", "pair": "pair-1",
                     "order": "forward", "repeat": 1},
    "candidate-B2": {"source": "candidate", "build": "candidate", "pair": "pair-2",
                     "order": "reverse", "repeat": 2},
    "baseline-A2": {"source": "baseline", "build": "baseline", "pair": "pair-2",
                    "order": "reverse", "repeat": 2},
}
HEX = set("0123456789abcdef")


def fail(message: str) -> NoReturn:
    raise RuntimeError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def sha(path: Path) -> str:
    try:
        digest = hashlib.sha256()
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
        return digest.hexdigest()
    except OSError as error:
        fail(f"cannot hash {path}: {error}")


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


def write_json(path: Path, value: Any) -> None:
    require(not path.exists() and not path.is_symlink(), f"refusing to replace {path}")
    path.write_text(json.dumps(value, indent=2) + "\n", encoding="utf-8")


def utc_now() -> str:
    return _datetime.datetime.now(_datetime.timezone.utc).isoformat()


def check_sha(value: Any, label: str) -> None:
    require(isinstance(value, str) and len(value) == 64 and set(value) <= HEX,
            f"{label} is not a lowercase SHA-256 digest")


def resolve_repo_path(raw: str) -> Path:
    path = Path(raw)
    return path.resolve() if path.is_absolute() else (REPO / path).resolve()


def source_census() -> dict[str, str]:
    """Return the packet's single source census, shared with its build tool."""

    custody_path = HERE / "custody.py"
    if custody_path.is_file() and not custody_path.is_symlink():
        spec = importlib.util.spec_from_file_location("custody_0715_capture", custody_path)
        require(spec is not None and spec.loader is not None,
                f"cannot load source census helper: {custody_path}")
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        value = module.census()
        require(isinstance(value, dict), "custody census is not an object")
        return value

    # This fallback keeps the script usable while the coordinator is preparing
    # the packet.  The normal capture path always uses custody.py.
    paths: list[Path] = [REPO / "Cargo.toml", REPO / "Cargo.lock"]
    paths.extend(path for path in (REPO / ".cargo").rglob("*") if path.is_file())
    for folder in ("crates", "tools/perf-baseline"):
        paths.extend(
            path for path in (REPO / folder).rglob("*")
            if path.is_file() and "target" not in path.parts
            and (path.suffix == ".rs" or path.name in {"Cargo.toml", "Cargo.lock"})
        )
    return {str(path.relative_to(REPO)): sha(path) for path in sorted(set(paths))}


def _validate_source_map(value: Any, label: str) -> dict[str, str]:
    require(isinstance(value, dict) and value, f"{label} is not a non-empty source map")
    for path, digest in value.items():
        require(isinstance(path, str) and path and not Path(path).is_absolute(),
                f"{label} has an invalid path")
        check_sha(digest, f"{label}[{path}]")
    return value


def load_plan() -> dict[str, Any]:
    plan = read_json(HERE / "pilot-plan.json")
    require(isinstance(plan, dict), "pilot-plan.json is not an object")
    require(plan.get("schema_version") == 1, "pilot plan schema changed")
    require(plan.get("packet") == "change-0715-docx-publication-collection-pilot",
            "pilot packet identity changed")
    require(plan.get("freeze_status") in {"draft", "frozen"},
            "pilot plan freeze status is invalid")
    revision = plan.get("revision")
    require(revision is None or (isinstance(revision, str) and len(revision) == 40
                                 and set(revision) <= HEX),
            "pilot revision must be null or a lowercase commit digest")
    require(plan.get("cpu") == 12, "CPU binding changed")
    require(plan.get("phase_order") == list(PHASES), "phase order changed")
    for lane, expected in (("native", {"repeats": 2, "samples": 100, "warmup": 10}),
                           ("allocator", {"repeats": 2, "samples": 3, "warmup": 0})):
        require(plan.get(lane) == expected, f"{lane} sample plan changed")

    corpora = plan.get("corpora")
    require(isinstance(corpora, list) and len(corpora) == 2,
            "pilot must contain exactly two DOCX corpora")
    ids: set[str] = set()
    for corpus in corpora:
        require(isinstance(corpus, dict), "corpus entry is malformed")
        identity = corpus.get("id")
        require(identity in {"generated", "numbered-list"} and identity not in ids,
                "corpus ids/order changed")
        ids.add(identity)
        require(isinstance(corpus.get("expected_edit_admitted"), bool),
                f"{identity}: edit admission is missing")
        if corpus.get("origin") == "generated-harness-corpus":
            require(identity == "generated" and corpus.get("path") is None
                    and corpus.get("sha256") is None and corpus.get("bytes") is None,
                    f"{identity}: generated fixture binding changed")
        else:
            require(corpus.get("origin") == "caller-named-real-file"
                    and isinstance(corpus.get("path"), str)
                    and corpus["path"], f"{identity}: real fixture binding missing")
            check_sha(corpus.get("sha256"), f"{identity}.sha256")
            require(isinstance(corpus.get("bytes"), int) and corpus["bytes"] > 0,
                    f"{identity}.bytes is invalid")
    require([item["id"] for item in corpora] == ["generated", "numbered-list"],
            "corpus order changed")

    stages = plan.get("stages")
    require(isinstance(stages, list) and [item.get("label") for item in stages]
            == list(STAGE_LABELS), "stage order changed")
    for item in stages:
        label = item.get("label")
        require(item == {"label": label, **STAGE_METADATA[label]},
                f"stage metadata changed: {label}")

    thresholds = plan.get("thresholds")
    require(isinstance(thresholds, dict), "thresholds are missing")
    require(thresholds.get("generated_counting_improvement_percent") == 10
            and thresholds.get("nonregression_percent") == 3
            and thresholds.get("allocation_nonregression_percent") == 3
            and thresholds.get("repeat_flag_percent") == 5
            and thresholds.get("tail_metrics") == ["p95", "p99"],
            "pilot thresholds changed")
    allowlist = plan.get("source_delta_allowlist")
    require(isinstance(allowlist, list) and allowlist
            and allowlist == sorted(set(allowlist)),
            "source delta allowlist is not canonical")
    require(all(isinstance(item, str) and not Path(item).is_absolute() for item in allowlist),
            "source delta allowlist contains an invalid path")
    files = plan.get("files")
    require(isinstance(files, dict)
            and files.get("builds") == "pilot-builds.json"
            and files.get("source_baseline") == "source-baseline.json"
            and files.get("source_candidate") == "source-candidate.json"
            and files.get("constraints") == "constraints.json",
            "pilot input filenames changed")
    return plan


def load_source_pair(plan: dict[str, Any] | None = None) -> tuple[dict[str, str], dict[str, str] | None]:
    plan = load_plan() if plan is None else plan
    files = plan["files"]
    baseline = _validate_source_map(read_json(HERE / files["source_baseline"]),
                                    files["source_baseline"])
    candidate_path = HERE / files["source_candidate"]
    candidate = (None if not candidate_path.exists()
                 else _validate_source_map(read_json(candidate_path),
                                           files["source_candidate"]))
    if candidate is None:
        return baseline, None
    changed = sorted(path for path in set(baseline) | set(candidate)
                     if baseline.get(path) != candidate.get(path))
    allowlist = sorted(plan["source_delta_allowlist"])
    require(set(changed) <= set(allowlist),
            f"candidate source delta contains paths outside allowlist: {changed}")
    return baseline, candidate


def verify_freeze(plan: dict[str, Any]) -> None:
    """Verify the coordinator's immutable script/plan freeze when present."""

    freeze_path = HERE / "pilot-freeze.json"
    require(freeze_path.is_file() and not freeze_path.is_symlink(),
            "pilot-freeze.json is required before the first child")
    freeze = read_json(freeze_path)
    require(isinstance(freeze, dict), "pilot-freeze.json is not an object")
    entries = freeze.get("files", freeze)
    require(isinstance(entries, dict) and entries, "pilot-freeze.json has no file digests")
    for raw, digest in entries.items():
        require(isinstance(raw, str) and not Path(raw).is_absolute(),
                "pilot freeze path is invalid")
        target = HERE / raw
        require(target.is_file() and not target.is_symlink(),
                f"pilot freeze file is missing: {raw}")
        check_sha(digest, f"pilot-freeze.json[{raw}]")
        require(sha(target) == digest, f"pilot freeze digest changed: {raw}")


def constraints_digest(plan: dict[str, Any]) -> str | None:
    path = HERE / plan["files"]["constraints"]
    if not path.exists():
        return None
    require(path.is_file() and not path.is_symlink(), f"invalid constraints file: {path}")
    constraints = read_json(path)
    require(isinstance(constraints, dict), "constraints file is not an object")
    for raw, digest in constraints.items():
        require(isinstance(raw, str) and not Path(raw).is_absolute(),
                "constraints path is invalid")
        check_sha(digest, f"constraints[{raw}]")
        target = REPO / raw
        require(target.is_file() and sha(target) == digest,
                f"constraint changed during capture: {raw}")
    return sha(path)


def _build_file_for(stage: str) -> Path:
    return HERE / f"build-{STAGE_METADATA[stage]['build']}.json"


def _normalize_binary(raw: Any, label: str) -> dict[str, Any]:
    require(isinstance(raw, dict), f"{label} binary record is not an object")
    path_value = raw.get("path", raw.get("binary"))
    digest = raw.get("sha256", raw.get("binary_sha256"))
    size = raw.get("bytes", raw.get("binary_bytes"))
    require(isinstance(path_value, str) and path_value, f"{label} binary path missing")
    check_sha(digest, f"{label}.binary_sha256")
    require(isinstance(size, int) and size > 0, f"{label}.binary_bytes is invalid")
    binary = resolve_repo_path(path_value)
    if binary.exists():
        require(binary.is_file() and not binary.is_symlink(), f"invalid binary: {binary}")
        require(binary.stat().st_size == size and sha(binary) == digest,
                f"{label} binary identity changed")
    else:
        witness = {"path": str(binary), "sha256": digest, "bytes": size}
        require(witness in read_json(HERE / "cleanup.json")["binaries"],
                f"{label} missing exact cleanup witness")
    return {"binary": str(binary), "binary_sha256": digest, "binary_bytes": size,
            "path": str(binary)}


def _record_from_list(path: Path, stage: str, lane: str) -> tuple[dict[str, Any], str]:
    rows = read_json(path)
    require(isinstance(rows, list), f"{path.name} is not a build-record list")
    names = {"allocator": "alloc", "native": "native"}
    matches = [row for row in rows if isinstance(row, dict)
               and (row.get("lane") == names[lane]
                    or row.get("lane") == lane
                    or Path(str(row.get("binary", ""))).name
                    in {f"{stage.split('-')[0]}-{names[lane]}",
                        f"{STAGE_METADATA[stage]['build']}-{names[lane]}"})]
    require(len(matches) == 1, f"{path.name} has no unique {stage}/{lane} row")
    row = matches[0]
    require(row.get("exit_code") == 0, f"{stage}/{lane} build failed")
    manifest = row.get("source_manifest", f"source-{STAGE_METADATA[stage]['source']}.json")
    normalized = _normalize_binary(row, f"{stage}/{lane}")
    normalized.update({
        "stage": stage,
        "lane": lane,
        "build_record": path.name,
        "build_record_sha256": sha(path),
        "source_manifest": manifest,
        "source_manifest_sha256": row.get("source_manifest_sha256"),
    })
    return normalized, path.name


def build_info(stage: str, lane: str, plan: dict[str, Any] | None = None) -> dict[str, Any]:
    """Load and validate one frozen binary identity.

    The preferred manifest is ``pilot-builds.json``.  For coordinator
    compatibility, the existing 0713 list-shaped ``build-baseline.json`` and
    ``build-candidate.json`` are accepted as an equivalent input.
    """

    plan = load_plan() if plan is None else plan
    require(stage in STAGE_METADATA and lane in LANES, "invalid stage or lane")
    builds_path = HERE / plan["files"]["builds"]
    record: dict[str, Any]
    if builds_path.is_file():
        manifest = read_json(builds_path)
        require(isinstance(manifest, dict), f"{builds_path.name} is not an object")
        stage_entry = manifest.get(STAGE_METADATA[stage]["build"])
        require(isinstance(stage_entry, dict), f"{builds_path.name} lacks {stage} entry")
        raw = stage_entry.get(lane, stage_entry.get("alloc" if lane == "allocator" else "native"))
        record = _normalize_binary(raw, f"{stage}/{lane}")
        build_record_name = builds_path.name
        manifest_name = stage_entry.get("source_manifest",
                                       f"source-{STAGE_METADATA[stage]['source']}.json")
        record.update({"stage": stage, "lane": lane,
                       "build_record": build_record_name,
                       "build_record_sha256": sha(builds_path),
                       "source_manifest": manifest_name,
                       "source_manifest_sha256": stage_entry.get("source_manifest_sha256")})
    else:
        record, build_record_name = _record_from_list(_build_file_for(stage), stage, lane)

    expected_manifest = f"source-{STAGE_METADATA[stage]['source']}.json"
    require(record["source_manifest"] == expected_manifest,
            f"{stage}/{lane} source manifest binding changed")
    source_path = HERE / expected_manifest
    source = _validate_source_map(read_json(source_path), expected_manifest)
    require(record.get("source_manifest_sha256") == sha(source_path),
            f"{stage}/{lane} source manifest digest changed")
    record["source"] = source
    record["source_manifest_sha256"] = sha(source_path)
    record["build_record"] = build_record_name
    return record


def fixture_info(corpus: dict[str, Any]) -> tuple[Path | None, dict[str, Any] | None]:
    if corpus["origin"] == "generated-harness-corpus":
        return None, None
    raw = str(corpus["path"])
    path = resolve_repo_path(raw)
    require(path.is_file() and not path.is_symlink(), f"missing fixture: {path}")
    digest = sha(path)
    require(digest == corpus["sha256"] and path.stat().st_size == corpus["bytes"],
            f"fixture digest or size changed: {corpus['id']}")
    return path, {"path": raw, "resolved_path": str(path), "bytes": path.stat().st_size,
                  "sha256": digest}


def phase_case(corpus: dict[str, Any], phase: str) -> str:
    prefix = ("docx_ordinary_save_" if corpus["origin"] == "generated-harness-corpus"
              else "docx_real_file_ordinary_save_")
    return prefix + phase


def child_name(stage: str, lane: str, corpus: dict[str, Any], phase: str) -> str:
    return f"{stage.replace('-', '_')}-{lane}-{corpus['id']}-{phase}"


def artifacts_for(name: str) -> list[Path]:
    return [HERE / f"{name}{suffix}" for suffix in (
        ".json", ".stdout", ".stderr", ".source-before.json", ".source-after.json",
        ".receipt.json")]


def ordered_jobs(plan: dict[str, Any], stage: str, lane: str | None = None) -> list[dict[str, Any]]:
    require(stage in STAGE_METADATA, f"unknown stage {stage!r}")
    metadata = STAGE_METADATA[stage]
    phases = list(plan["phase_order"])
    corpora = list(plan["corpora"])
    if metadata["order"] == "reverse":
        phases.reverse()
        corpora.reverse()
    result: list[dict[str, Any]] = []
    for index, (phase, corpus) in enumerate((pair for phase in phases for pair in
                                               ((phase, corpus) for corpus in corpora))):
        if lane is None:
            current_lane = "native"
        else:
            current_lane = lane
        name = child_name(stage, current_lane, corpus, phase)
        result.append({
            "lane": current_lane,
            "stage": stage,
            "repeat": metadata["repeat"],
            "pair": metadata["pair"],
            "stage_order": metadata["order"],
            "stage_order_index": index,
            "corpus": corpus,
            "corpus_id": corpus["id"],
            "phase": phase,
            "name": name,
            "case": phase_case(corpus, phase),
            "samples": plan[current_lane]["samples"],
            "warmup": plan[current_lane]["warmup"],
        })
    return result


def expected_command(job: dict[str, Any], build: dict[str, Any], plan: dict[str, Any]) -> list[str]:
    command = [
        "taskset", "-c", str(plan["cpu"]), build["binary"],
        "--warmup", str(job["warmup"]), "--samples", str(job["samples"]),
        "--case", job["case"], "--json", str(HERE / f"{job['name']}.json"),
        "--filesystem-root", str(plan["filesystem_root"]),
    ]
    if job["corpus"]["origin"] == "caller-named-real-file":
        # Keep the fixture argument relative to the repository, matching the
        # frozen 0713/0714 command contract.
        command += ["--ooxml-file", str(job["corpus"]["path"])]
    return command


def _check_constraints(plan: dict[str, Any]) -> str | None:
    return constraints_digest(plan)


def run_child(*, plan: dict[str, Any], stage: str, lane: str, job: dict[str, Any],
              build: dict[str, Any], live_source: dict[str, str], live_source_label: str,
              candidate_source: dict[str, str] | None,
              baseline_source: dict[str, str], constraints_sha: str | None) -> None:
    name = job["name"]
    for path in artifacts_for(name):
        require(not path.exists() and not path.is_symlink(), f"refusing to replace {path}")
    _check_constraints(plan)
    before = source_census()
    require(before == live_source,
            f"{name}: checkout is not source-{live_source_label}.json")
    source_before = HERE / f"{name}.source-before.json"
    write_json(source_before, before)
    fixture_path, fixture_before = fixture_info(job["corpus"])
    command = expected_command(job, build, plan)
    started = utc_now()
    tick = time.monotonic()
    env = dict(os.environ)
    env.update({"LC_ALL": "C", "LANG": "C", "TZ": "UTC",
                "PERL_HASH_SEED": "0", "PERL_PERTURB_KEYS": "0"})
    stdout_path = HERE / f"{name}.stdout"
    stderr_path = HERE / f"{name}.stderr"
    with stdout_path.open("wb") as stdout, stderr_path.open("wb") as stderr:
        result = subprocess.run(command, cwd=REPO, env=env, stdout=stdout, stderr=stderr)
    seconds = time.monotonic() - tick
    after = source_census()
    require(after == live_source, f"{name}: checkout changed during child")
    source_after = HERE / f"{name}.source-after.json"
    write_json(source_after, after)
    require(before == after, f"{name}: source changed during child")
    require(sha(resolve_repo_path(build["binary"])) == build["binary_sha256"],
            f"{name}: binary changed during child")
    fixture_after = None if fixture_path is None else fixture_info(job["corpus"])[1]
    require(fixture_before == fixture_after, f"{name}: fixture changed during child")

    source_delta = ([] if candidate_source is None else
                    sorted(path for path in set(baseline_source) | set(candidate_source)
                           if baseline_source.get(path) != candidate_source.get(path)))
    artifacts: dict[str, str] = {}
    for path in (HERE / f"{name}.json", stdout_path, stderr_path, source_before, source_after):
        require(path.is_file() and not path.is_symlink(), f"{name}: missing {path.name}")
        artifacts[path.name] = sha(path)
    metadata = STAGE_METADATA[stage]
    build_record_path = HERE / build["build_record"]
    receipt = {
        "schema_version": 1,
        "packet": plan["packet"],
        "name": name,
        "stage": stage,
        "source_label": metadata["source"],
        "build_label": metadata["build"],
        "pair": metadata["pair"],
        "repeat": metadata["repeat"],
        "lane": lane,
        "stage_order": metadata["order"],
        "stage_order_index": job["stage_order_index"],
        "corpus_id": job["corpus_id"],
        "corpus_label": job["corpus"]["label"],
        "corpus_origin": job["corpus"]["origin"],
        "phase": job["phase"],
        "case": job["case"],
        "samples": job["samples"],
        "warmup": job["warmup"],
        "command": command,
        "start_utc": started,
        "end_utc": utc_now(),
        "seconds": seconds,
        "exit_code": result.returncode,
        "cpu": plan["cpu"],
        "binary_path": build["binary"],
        "binary_sha256": build["binary_sha256"],
        "binary_bytes": build["binary_bytes"],
        "build_record": build["build_record"],
        "build_record_sha256": build["build_record_sha256"],
        "retained_binary_source": {
            "manifest": build["source_manifest"],
            "manifest_sha256": build["source_manifest_sha256"],
            "source_census_sha256": digest_json(build["source"]),
            "source_entry_count": len(build["source"]),
        },
        "current_checkout_source": {
            "manifest": f"source-{live_source_label}.json",
            "manifest_sha256": sha(HERE / f"source-{live_source_label}.json"),
            "before_artifact": source_before.name,
            "after_artifact": source_after.name,
            "before_sha256": digest_json(before),
            "after_sha256": digest_json(after),
            "before_file_sha256": sha(source_before),
            "after_file_sha256": sha(source_after),
            "before_entry_count": len(before),
            "after_entry_count": len(after),
            "unchanged_during_child": before == after,
        },
        "source_delta": {
            "baseline_manifest": plan["files"]["source_baseline"],
            "candidate_manifest": (None if candidate_source is None
                                    else plan["files"]["source_candidate"]),
            "allowed_paths": list(plan["source_delta_allowlist"]),
            "changed_paths": source_delta,
            "candidate_available": candidate_source is not None,
            "within_allowlist": set(source_delta) <= set(plan["source_delta_allowlist"]),
        },
        "fixture": {"plan_path": job["corpus"].get("path"),
                    "plan_sha256": job["corpus"].get("sha256"),
                    "before": fixture_before, "after": fixture_after},
        "pilot_plan_sha256": sha(HERE / "pilot-plan.json"),
        "script_sha256": sha(Path(__file__).resolve()),
        "constraints_sha256": constraints_sha,
        "environment": {key: env.get(key)
                         for key in ("LC_ALL", "LANG", "TZ", "RUSTFLAGS", "LD_PRELOAD",
                                     "MALLOC_CONF", "GLIBC_TUNABLES", "PERL_HASH_SEED",
                                     "PERL_PERTURB_KEYS")},
        "artifacts": artifacts,
    }
    write_json(HERE / f"{name}.receipt.json", receipt)
    require(result.returncode == 0, f"{name}: benchmark failed with exit code {result.returncode}")
    print(f"{name} passed", flush=True)


def run_lane(stage: str, raw_lane: str) -> None:
    plan = load_plan()
    require(plan["freeze_status"] == "frozen" and plan.get("revision") is not None,
            "pilot-plan.json must be frozen before capture")
    verify_freeze(plan)
    require(stage in STAGE_METADATA, f"unknown stage {stage!r}")
    lane = LANE_ALIASES.get(raw_lane, raw_lane)
    require(lane in LANES, "lane must be native or allocator")
    baseline_source, candidate_source = load_source_pair(plan)
    if stage == "baseline-A1":
        live_source, live_source_label = baseline_source, "baseline"
    else:
        require(candidate_source is not None,
                f"{stage} requires source-candidate.json")
        live_source, live_source_label = candidate_source, "candidate"
    build = build_info(stage, lane, plan)
    constraints_sha = _check_constraints(plan)
    jobs = ordered_jobs(plan, stage, lane)
    require(len(jobs) == 4, f"fixed stage child count changed: {len(jobs)}")
    for job in jobs:
        run_child(plan=plan, stage=stage, lane=lane, job=job, build=build,
                  live_source=live_source, live_source_label=live_source_label,
                  candidate_source=candidate_source, baseline_source=baseline_source,
                  constraints_sha=constraints_sha)
    print(f"stage {stage} lane {lane} complete ({len(jobs)} children)", flush=True)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("stage", help=", ".join(STAGE_LABELS))
    parser.add_argument("lane", help="native, allocator, or alloc")
    args = parser.parse_args()
    stage, lane = args.stage, args.lane
    # Preserve the convenient lane-first spelling used by a few old packet
    # drivers while receipts always retain the canonical stage-first command.
    if stage in LANES or stage in LANE_ALIASES:
        stage, lane = lane, stage
    try:
        run_lane(stage, lane)
    except (AssertionError, KeyError, OSError, RuntimeError, TypeError, ValueError) as error:
        print(f"capture failed: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
