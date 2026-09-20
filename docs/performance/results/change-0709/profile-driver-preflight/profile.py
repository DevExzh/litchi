#!/usr/bin/env python3
"""Capture the packet-only 0709 DOCX Callgrind profiles.

The native build is owned by the coordinator.  This driver consumes the
already frozen release binary, starts one Valgrind process per
``(repeat, corpus, phase)`` job, and retains every Callgrind part produced by
that process.  It does not invoke Cargo, change production sources, or use
Valgrind as an allocation counter.
"""

from __future__ import annotations

import argparse
import datetime as _datetime
import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys
import time
from typing import Any, NoReturn


HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]


def fail(message: str) -> NoReturn:
    raise RuntimeError(message)


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


def write_json(path: Path, value: Any) -> None:
    path.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n", encoding="utf-8")


def utc_now() -> str:
    return _datetime.datetime.now(_datetime.timezone.utc).isoformat()


def source_census() -> dict[str, str]:
    """Match the 0709 build/capture source census exactly."""

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
    return {
        str(path.relative_to(REPO)): sha(path)
        for path in sorted(set(paths))
    }


def load_plan() -> dict[str, Any]:
    value = read_json(HERE / "plan.json")
    require(isinstance(value, dict), "plan.json is not an object")
    require(isinstance(value.get("revision"), str) and value["revision"],
            "plan revision is missing")
    require(isinstance(value.get("cpu"), int) and value["cpu"] >= 0,
            "plan cpu is invalid")
    require(value.get("native") == {"repeats": 3, "samples": 100, "warmup": 10},
            "native plan changed")
    require(value.get("allocator") == {"repeats": 2, "samples": 3, "warmup": 0},
            "allocator plan changed")
    require(value.get("phase_order") == ["lifecycle", "edit", "atomic_publish", "counting_publish"],
            "ordinary-save phase order changed")
    corpora = value.get("corpora")
    require(isinstance(corpora, list) and len(corpora) == 3,
            "ordinary-save corpus matrix changed")
    ids: set[str] = set()
    for corpus in corpora:
        require(isinstance(corpus, dict), "ordinary-save corpus entry is malformed")
        identity = corpus.get("id")
        require(isinstance(identity, str) and identity and identity not in ids,
                "ordinary-save corpus id is invalid or duplicated")
        ids.add(identity)
        origin = corpus.get("origin")
        require(origin in {"generated-harness-corpus", "caller-named-real-file"},
                f"unsupported corpus origin: {origin!r}")
        if origin == "generated-harness-corpus":
            require(corpus.get("path") is None and corpus.get("sha256") is None,
                    "generated corpus unexpectedly has a fixture binding")
        else:
            raw_path = corpus.get("path")
            digest = corpus.get("sha256")
            require(isinstance(raw_path, str) and isinstance(digest, str),
                    f"{identity}: real fixture binding is incomplete")
            path = (REPO / raw_path).resolve()
            require(path.is_file() and not path.is_symlink(),
                    f"{identity}: real fixture is missing: {path}")
            require(sha(path) == digest, f"{identity}: real fixture digest changed")
            require(corpus.get("bytes") == path.stat().st_size,
                    f"{identity}: real fixture size changed")
    require(ids == {"generated", "numbered-list", "alt-chunk-header"},
            f"ordinary-save corpus ids changed: {sorted(ids)}")
    return value


def load_profile_plan(plan: dict[str, Any]) -> dict[str, Any]:
    value = read_json(HERE / "profile-plan.json")
    require(isinstance(value, dict), "profile-plan.json is not an object")
    require(value.get("schema_version") == 1, "profile plan schema changed")
    require(value.get("revision") == plan["revision"], "profile plan revision differs")
    require(value.get("corpora") == ["generated", "numbered-list"],
            "profile corpus matrix changed")
    require(value.get("phases") == ["edit", "counting_publish"],
            "profile phase matrix changed")
    require(value.get("repeats") == 2 and value.get("samples") == 1
            and value.get("warmup") == 0, "profile sample policy changed")
    owners = value.get("owners")
    require(owners == {
        "edit": "litchi_perf_baseline::ordinary_save::Owner::edit",
        "counting_publish": "litchi_docx::package::codec::<impl litchi_docx::package::model::Package>::write_plain",
    }, "profile owner symbols changed")
    require(value.get("owner_match") == {"edit": "exact", "counting_publish": "exact"},
            "profile owner match policy changed")
    options = value.get("options")
    require(isinstance(options, dict), "profile Callgrind options are missing")
    for phase in ("edit", "counting_publish"):
        phase_options = options.get(phase)
        require(isinstance(phase_options, list) and all(isinstance(item, str) for item in phase_options),
                f"{phase}: profile Callgrind options are invalid")
        for token in ("--tool=callgrind", "--collect-atstart=no",
                      "--toggle-collect=" + owners[phase],
                      "--zero-before=" + owners[phase],
                      "--dump-after=" + owners[phase]):
            require(token in phase_options, f"{phase}: profile options omit {token}")
    expected = value.get("expected_numbered_dumps")
    require(expected == {
        "edit": {"generated": 5, "numbered-list": 5},
        "counting_publish": {"generated": 6, "numbered-list": 5},
    },
            "profile dump-count policy changed")
    require(value.get("measured_parent") ==
            "litchi_perf_baseline::ordinary_save::run_case",
            "profile measured caller changed")
    setup = value.get("setup_parents")
    require(setup == {
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
        }
    }, "profile setup caller policy changed")
    return value


def fixture_binding(corpus: dict[str, Any]) -> dict[str, Any] | None:
    if corpus["origin"] == "generated-harness-corpus":
        return None
    path = (REPO / corpus["path"]).resolve()
    return {
        "plan_path": corpus["path"],
        "resolved_path": str(path),
        "bytes": path.stat().st_size,
        "sha256": sha(path),
    }


def build_info() -> dict[str, Any]:
    records_path = HERE / "build-baseline.json"
    records = read_json(records_path)
    require(isinstance(records, list), "build-baseline.json is not a list")
    matches = [item for item in records
               if isinstance(item, dict)
               and Path(str(item.get("binary", ""))).name == "baseline-native"]
    require(len(matches) == 1, "baseline-native build record is not unique")
    record = matches[0]
    require(record.get("exit_code") == 0, "baseline-native build failed")
    binary = Path(str(record.get("binary", ""))).resolve()
    digest = record.get("binary_sha256")
    size = record.get("binary_bytes")
    require(isinstance(digest, str) and len(digest) == 64,
            "baseline-native binary digest is missing")
    require(isinstance(size, int) and size > 0, "baseline-native binary size is invalid")
    require(binary.is_file() and not binary.is_symlink(),
            f"baseline-native binary is missing: {binary}")
    require(sha(binary) == digest and binary.stat().st_size == size,
            "baseline-native binary identity changed")
    source_path = HERE / "source-baseline.json"
    source = read_json(source_path)
    require(isinstance(source, dict) and source, "source-baseline.json is invalid")
    require(record.get("source_manifest_sha256") == sha(source_path),
            "baseline-native source manifest binding changed")
    symbols_json_path = HERE / "symbols.json"
    symbols_text_path = HERE / "symbols.txt"
    symbols = read_json(symbols_json_path)
    require(isinstance(symbols, dict), "symbols.json is not an object")
    require(symbols.get("binary_sha256") == digest,
            "nm symbol artifact is bound to a different binary")
    require(symbols.get("log_sha256") == sha(symbols_text_path),
            "nm symbol text digest changed")
    symbol_text = symbols_text_path.read_text(encoding="utf-8", errors="replace").splitlines()
    edit_symbol = "litchi_perf_baseline::ordinary_save::Owner::edit"
    run_case_symbol = "litchi_perf_baseline::ordinary_save::run_case"
    write_plain_symbol = (
        "litchi_docx::package::codec::<impl "
        "litchi_docx::package::model::Package>::write_plain"
    )
    require(sum(line.endswith(" " + edit_symbol) for line in symbol_text) == 1,
            "nm symbol artifact does not contain one Owner::edit symbol")
    require(sum(line.endswith(" " + run_case_symbol) for line in symbol_text) == 1,
            "nm symbol artifact does not contain one run_case symbol")
    require(sum(line.endswith(" " + write_plain_symbol) for line in symbol_text) == 4,
            "nm symbol artifact does not contain four exact write_plain symbols")
    require(any(line.endswith(" " + write_plain_symbol + "::{{closure}}")
                for line in symbol_text),
            "nm symbol artifact unexpectedly omits write_plain closure evidence")
    return {
        "record": record,
        "record_path": records_path.name,
        "record_sha256": sha(records_path),
        "binary": str(binary),
        "binary_sha256": digest,
        "binary_bytes": size,
        "source_path": source_path.name,
        "source_manifest_sha256": sha(source_path),
        "source": source,
        "source_census_sha256": digest_json(source),
        "symbols": {
            "json": str(symbols_json_path),
            "json_sha256": sha(symbols_json_path),
            "text": str(symbols_text_path),
            "text_sha256": sha(symbols_text_path),
            "edit_symbol_count": 1,
            "run_case_symbol_count": 1,
            "write_plain_exact_symbol_count": 4,
        },
    }


def slug(value: str) -> str:
    return "".join(char if char.isalnum() or char in "-_" else "_" for char in value)


def profile_name(repeat: int, corpus_id: str, phase: str) -> str:
    return f"profile-r{repeat}-{slug(corpus_id)}-{slug(phase)}"


def expected_jobs(profile: dict[str, Any], plan: dict[str, Any]) -> list[dict[str, Any]]:
    corpora = {item["id"]: item for item in plan["corpora"]}
    result: list[dict[str, Any]] = []
    index = 0
    for repeat in range(1, profile["repeats"] + 1):
        phases = list(profile["phases"])
        corpus_ids = list(profile["corpora"])
        if repeat % 2 == 0:
            phases.reverse()
            corpus_ids.reverse()
        for phase in phases:
            for corpus_id in corpus_ids:
                corpus = corpora[corpus_id]
                generated = corpus["origin"] == "generated-harness-corpus"
                case = ("docx_ordinary_save_" if generated
                        else "docx_real_file_ordinary_save_") + phase
                result.append({
                    "name": profile_name(repeat, corpus_id, phase),
                    "repeat": repeat,
                    "corpus_id": corpus_id,
                    "corpus": corpus,
                    "phase": phase,
                    "case": case,
                    "samples": profile["samples"],
                    "warmup": profile["warmup"],
                    "order_index": index,
                })
                index += 1
    return result


def refuse_existing(name: str) -> None:
    patterns = [
        f"{name}.json", f"{name}.stdout", f"{name}.stderr",
        f"{name}.receipt.json", f"{name}.source-before.json",
        f"{name}.source-after.json", f"{name}.fixture-before.json",
        f"{name}.fixture-after.json", f"{name}.callgrind",
    ]
    patterns.extend(f"{name}.callgrind.{part}" for part in range(1, 64))
    existing = [HERE / filename for filename in patterns if (HERE / filename).exists()]
    require(not existing, f"refusing to replace existing profile artifacts: {existing[0]}")


def write_custody(name: str, expected_source: dict[str, str], corpus: dict[str, Any]) -> tuple[dict[str, Any], dict[str, Any] | None]:
    source = source_census()
    require(source == expected_source, f"{name}: current source differs before child")
    source_path = HERE / f"{name}.source-before.json"
    write_json(source_path, source)
    fixture = fixture_binding(corpus)
    fixture_path: Path | None = None
    if fixture is None:
        write_json(HERE / f"{name}.fixture-before.json", {
            "plan_path": None, "resolved_path": None, "bytes": None, "sha256": None,
        })
    else:
        fixture_path = HERE / f"{name}.fixture-before.json"
        write_json(fixture_path, fixture)
    return {
        "before_artifact": source_path.name,
        "before_sha256": digest_json(source),
        "before_file_sha256": sha(source_path),
    }, fixture


def finish_custody(name: str, expected_source: dict[str, str], corpus: dict[str, Any],
                   custody: dict[str, Any], fixture_before: dict[str, Any] | None) -> dict[str, Any]:
    source = source_census()
    after_path = HERE / f"{name}.source-after.json"
    write_json(after_path, source)
    require(source == expected_source, f"{name}: current source differs after child")
    require(source == json.loads((HERE / custody["before_artifact"]).read_text()),
            f"{name}: source changed during child")
    custody.update({
        "after_artifact": after_path.name,
        "after_sha256": digest_json(source),
        "after_file_sha256": sha(after_path),
        "unchanged_during_child": True,
    })
    fixture_after = fixture_binding(corpus)
    write_json(HERE / f"{name}.fixture-after.json", fixture_after or {
        "plan_path": None, "resolved_path": None, "bytes": None, "sha256": None,
    })
    require(fixture_after == fixture_before, f"{name}: fixture changed during child")
    return custody


def callgrind_parts(name: str) -> list[Path]:
    paths = [path for path in HERE.glob(f"{name}.callgrind.[0-9]*")
             if path.is_file() and path.suffix[1:].isdigit()]
    paths.sort(key=lambda path: int(path.suffix[1:]))
    parts = [int(path.suffix[1:]) for path in paths]
    require(parts == list(range(1, len(parts) + 1)),
            f"{name}: Callgrind numbered parts are not contiguous: {parts}")
    require(paths, f"{name}: Callgrind emitted no numbered parts")
    return paths


def artifact_inventory(name: str, custody: dict[str, Any], fixture: dict[str, Any] | None) -> dict[str, str]:
    paths = [
        HERE / f"{name}.json", HERE / f"{name}.stdout", HERE / f"{name}.stderr",
        HERE / f"{name}.callgrind", HERE / custody["before_artifact"],
        HERE / custody["after_artifact"], HERE / f"{name}.fixture-before.json",
        HERE / f"{name}.fixture-after.json",
    ]
    paths.extend(callgrind_parts(name))
    result: dict[str, str] = {}
    for path in paths:
        require(path.is_file() and not path.is_symlink(), f"{name}: missing artifact {path.name}")
        result[path.name] = sha(path)
    return dict(sorted(result.items()))


def run_job(job: dict[str, Any], plan: dict[str, Any], profile: dict[str, Any],
            build: dict[str, Any], expected_source: dict[str, str],
            plan_sha: str, profile_sha: str, constraints_sha: str) -> None:
    name = job["name"]
    refuse_existing(name)
    owner = profile["owners"][job["phase"]]
    options = list(profile["options"][job["phase"]])
    callgrind_path = HERE / f"{name}.callgrind"
    json_path = HERE / f"{name}.json"
    stdout_path = HERE / f"{name}.stdout"
    stderr_path = HERE / f"{name}.stderr"
    command = ["taskset", "-c", str(plan["cpu"]), "valgrind", *options,
               f"--callgrind-out-file={callgrind_path}", build["binary"],
               "--warmup", str(job["warmup"]), "--samples", str(job["samples"]),
               "--case", job["case"], "--json", str(json_path),
               "--filesystem-root", plan["filesystem_root"]]
    if job["corpus"]["origin"] == "caller-named-real-file":
        command.extend(["--ooxml-file", job["corpus"]["path"]])
    before, fixture_before = write_custody(name, expected_source, job["corpus"])
    start = time.monotonic_ns()
    started = utc_now()
    try:
        with stdout_path.open("w", encoding="utf-8") as stdout, stderr_path.open("w", encoding="utf-8") as stderr:
            process = subprocess.run(command, cwd=REPO, stdout=stdout, stderr=stderr,
                                     check=False, env=dict(os.environ))
        exit_code = process.returncode
    except OSError as error:
        stderr_path.write_text(str(error) + "\n", encoding="utf-8")
        exit_code = 127
    finished = utc_now()
    elapsed = time.monotonic_ns() - start
    after = finish_custody(name, expected_source, job["corpus"], before, fixture_before)
    # Keep a receipt even for a failed child, so a failed attempt is auditable
    # and the next invocation can refuse to overwrite it explicitly.
    receipt: dict[str, Any] = {
        "schema_version": 1,
        "lane": "profile",
        "name": name,
        "repeat": job["repeat"],
        "corpus_id": job["corpus_id"],
        "corpus_origin": job["corpus"]["origin"],
        "phase": job["phase"],
        "case": job["case"],
        "samples": job["samples"],
        "warmup": job["warmup"],
        "order_index": job["order_index"],
        "cpu": plan["cpu"],
        "owner": owner,
        "owner_match": profile["owner_match"][job["phase"]],
        "measured_parent": profile["measured_parent"],
        "setup_parents": profile["setup_parents"][job["phase"]],
        "expected_numbered_dumps": profile["expected_numbered_dumps"][job["phase"]][job["corpus_id"]],
        "command": command,
        "started_utc": started,
        "finished_utc": finished,
        "wall_time_ns": elapsed,
        "exit_code": exit_code,
        "binary": build["binary"],
        "binary_sha256": build["binary_sha256"],
        "binary_bytes": build["binary_bytes"],
        "build_record": build["record_path"],
        "build_record_sha256": build["record_sha256"],
        "build_source_manifest": build["source_path"],
        "build_source_manifest_sha256": build["source_manifest_sha256"],
        "retained_source_census_sha256": build["source_census_sha256"],
        "retained_source_entry_count": len(expected_source),
        "plan_sha256": plan_sha,
        "profile_plan_sha256": profile_sha,
        "constraints_sha256": constraints_sha,
        "script_sha256": sha(Path(__file__).resolve()),
        "current_checkout_source": after,
        "fixture": {
            "before_artifact": f"{name}.fixture-before.json",
            "after_artifact": f"{name}.fixture-after.json",
            "before": fixture_before,
            "after": fixture_binding(job["corpus"]),
        },
        "warnings": [
            "Callgrind Ir is guest-instruction attribution; no allocation count is inferred from Valgrind.",
            "The analyzer selects the measured dump from a raw incoming edge; dump numbering is retained as evidence.",
        ],
    }
    try:
        receipt["artifacts"] = artifact_inventory(name, after, fixture_before)
    except RuntimeError as error:
        receipt["artifact_inventory_error"] = str(error)
        receipt["artifacts"] = {}
    write_json(HERE / f"{name}.receipt.json", receipt)
    require(exit_code == 0, f"{name}: child exited with {exit_code}; receipt retained")
    require((HERE / f"{name}.json").is_file(), f"{name}: child did not emit JSON")
    require((HERE / f"{name}.callgrind").is_file(), f"{name}: child did not emit Callgrind output")


def plan_only(plan: dict[str, Any], profile: dict[str, Any]) -> None:
    for job in expected_jobs(profile, plan):
        print(json.dumps({
            "name": job["name"], "repeat": job["repeat"],
            "corpus_id": job["corpus_id"], "phase": job["phase"],
            "case": job["case"], "owner": profile["owners"][job["phase"]],
        }, sort_keys=True))


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--plan-only", action="store_true",
                        help="validate the frozen profile plan and print its jobs")
    parser.add_argument("--repeat", type=int,
                        help="run only one frozen repeat (for controlled recovery)")
    args = parser.parse_args()
    try:
        plan = load_plan()
        profile = load_profile_plan(plan)
        if args.repeat is not None:
            require(args.repeat in (1, 2), "--repeat must be 1 or 2")
        if args.plan_only:
            plan_only(plan, profile)
            return 0
        expected_source = read_json(HERE / "source-baseline.json")
        require(expected_source == source_census(),
                "current source differs from source-baseline.json before profiling")
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
        jobs = expected_jobs(profile, plan)
        if args.repeat is not None:
            jobs = [job for job in jobs if job["repeat"] == args.repeat]
        require(jobs, "profile job matrix is empty")
        for job in jobs:
            run_job(job, plan, profile, build, expected_source,
                    plan_sha, profile_sha, constraints_sha)
        print(f"captured {len(jobs)} DOCX Callgrind profile jobs")
        return 0
    except (RuntimeError, OSError, KeyError, ValueError) as error:
        print(f"profile.py: error: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
