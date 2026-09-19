#!/usr/bin/env python3
"""Verify the frozen 0697 MCE attribution packet without measuring."""

from __future__ import annotations

import ast
import hashlib
import json
import math
import os
import re
import subprocess
import sys
from pathlib import Path
from typing import Any


P = Path(__file__).resolve().parent
ROOT = P.parents[3]
EVENTS = (
    "cycles",
    "instructions",
    "branches",
    "branch-misses",
    "cache-misses",
    "page-faults",
    "task-clock",
)
SEQUENCE_NAMES = ["all", "presentation", "slides"] + [f"slide{i}" for i in range(1, 14)]
PROFILE_NAMES = ["all", "presentation", "slides"]
HEX = re.compile(r"^0x[0-9a-fA-F]+$")


def stop(message: str) -> None:
    raise AssertionError(message)


def need(path: Path) -> Path:
    if not path.is_file():
        stop(f"missing packet file: {path.relative_to(P)}")
    return path


def read(name: str) -> Any:
    return json.loads(need(P / name).read_text())


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def packet_path(value: str | os.PathLike[str]) -> Path:
    path = Path(value)
    if not path.is_absolute():
        path = ROOT / path
    return path


def same(a: Any, b: Any, label: str) -> None:
    if a != b:
        stop(f"{label}: {a!r} != {b!r}")


def parse_identity(line: str) -> dict[str, str]:
    if not line.startswith("IDENTITY\t"):
        stop(f"malformed identity line: {line!r}")
    result: dict[str, str] = {}
    for field in line.split("\t")[1:]:
        key, sep, value = field.partition("=")
        if not sep or key in result:
            stop(f"malformed identity field: {line!r}")
        result[key] = value
    required = {
        "index",
        "input_bytes",
        "input_sha256",
        "output_bytes",
        "output_sha256",
        "borrowed",
    }
    same(set(result), required, "identity fields")
    if not re.fullmatch(r"[0-9a-f]{64}", result["input_sha256"]):
        stop(f"invalid input digest: {line!r}")
    if not re.fullmatch(r"[0-9a-f]{64}", result["output_sha256"]):
        stop(f"invalid output digest: {line!r}")
    return result


def parse_probe(path: Path) -> tuple[list[dict[str, str]], dict[str, str], list[int]]:
    identities: list[dict[str, str]] = []
    meta: dict[str, str] | None = None
    samples: list[int] = []
    stage = "identity"
    for line in path.read_text().splitlines():
        if line.startswith("IDENTITY\t"):
            if stage != "identity":
                stop(f"identity after metadata/samples: {path.name}")
            identities.append(parse_identity(line))
        elif line.startswith("META\t"):
            if meta is not None or stage == "samples":
                stop(f"duplicate/out-of-order META: {path.name}")
            meta = {}
            for field in line.split("\t")[1:]:
                key, sep, value = field.partition("=")
                if not sep or key in meta:
                    stop(f"malformed META: {path.name}")
                meta[key] = value
            stage = "samples"
        elif line.startswith("SAMPLE\t"):
            if meta is None:
                stop(f"sample before META: {path.name}")
            fields = dict(field.split("=", 1) for field in line.split("\t")[1:])
            if set(fields) != {"index", "elapsed_ns"}:
                stop(f"malformed SAMPLE: {path.name}")
            if int(fields["index"]) != len(samples):
                stop(f"non-contiguous SAMPLE index: {path.name}")
            elapsed = int(fields["elapsed_ns"])
            if elapsed <= 0:
                stop(f"non-positive sample: {path.name}")
            samples.append(elapsed)
        else:
            stop(f"unknown probe output line in {path.name}: {line!r}")
    if meta is None:
        stop(f"missing META: {path.name}")
    return identities, meta, samples


def stat(values: list[float]) -> dict[str, float]:
    ordered = sorted(values)
    n = len(ordered)
    return {
        "p50_ns": (ordered[n // 2 - 1] + ordered[n // 2]) / 2
        if n % 2 == 0
        else ordered[n // 2],
        "mean_ns": sum(ordered) / n,
        "p95_ns": ordered[math.ceil(0.95 * n) - 1],
        "p99_ns": ordered[math.ceil(0.99 * n) - 1],
        "min_ns": ordered[0],
        "max_ns": ordered[-1],
    }


def expected_identities(manifest: dict[str, Any], sequence: list[str]) -> list[dict[str, str]]:
    members = {row["corpus_path"]: row for row in manifest["members"]}
    output_lengths = {}
    for row in manifest["calls"]:
        output_lengths.setdefault(row["relative_path"], row["trace_output_len"])
    result = []
    seen: set[str] = set()
    for relative in sequence:
        if relative in seen:
            continue
        seen.add(relative)
        member = members.get(relative)
        if member is None:
            stop(f"sequence references unlisted member: {relative}")
        result.append(
            {
                "index": str(len(result)),
                "input_bytes": str(member["bytes"]),
                "input_sha256": member["sha256"],
                "output_bytes": str(output_lengths[relative]),
                "borrowed": "false",
            }
        )
    return result


def validate_corpus() -> tuple[dict[str, Any], dict[str, list[str]]]:
    manifest = read("corpus/manifest.json")
    same(manifest.get("schema"), "litchi-0697-mce-attribution-v1", "corpus schema")
    same(manifest.get("status"), "prepared", "corpus status")
    source = ROOT / manifest["source_archive"]["path"]
    need(source)
    same(sha(source), manifest["source_archive"]["sha256"], "source archive hash")
    members = manifest.get("members", [])
    same(len(members), 14, "member count")
    by_path = {}
    seen_uris: set[str] = set()
    for row in members:
        if row["corpus_path"] in by_path or row["uri"] in seen_uris:
            stop("duplicate corpus member identity")
        seen_uris.add(row["uri"])
        path = ROOT / row["corpus_path"]
        need(path)
        same(len(path.read_bytes()), row["bytes"], f"member bytes {row['uri']}")
        same(sha(path), row["sha256"], f"member hash {row['uri']}")
        by_path[row["corpus_path"]] = row

    expected: dict[str, list[str]] = {}
    expected["all"] = ["../corpus/presentation.xml"] * 3 + [
        f"../corpus/slide{i}.xml" for i in range(1, 14)
    ] + ["../corpus/presentation.xml"] * 2
    expected["presentation"] = ["../corpus/presentation.xml"] * 5
    expected["slides"] = [f"../corpus/slide{i}.xml" for i in range(1, 14)]
    expected.update({f"slide{i}": [f"../corpus/slide{i}.xml"] for i in range(1, 14)})
    sequences = manifest.get("sequences", {})
    same(set(sequences), set(SEQUENCE_NAMES), "sequence census")
    for name in SEQUENCE_NAMES:
        row = sequences[name]
        path = ROOT / row["path"]
        need(path)
        same(path.read_text().splitlines(), expected[name], f"sequence rows {name}")
        same(sha(path), row["sha256"], f"sequence hash {name}")
        same(row["relative_paths"], expected[name], f"manifest sequence rows {name}")
        same(row["lines"], len(expected[name]), f"sequence length {name}")
        base = path.parent
        for reference in expected[name]:
            resolved = (base / reference).resolve()
            if not resolved.is_file() or not resolved.is_relative_to(P.resolve()):
                stop(f"sequence path escapes packet: {name}/{reference}")
    calls = manifest.get("calls", [])
    same(len(calls), 18, "call count")
    uris = ["/ppt/presentation.xml"] * 3 + [f"/ppt/slides/slide{i}.xml" for i in range(1, 14)] + [
        "/ppt/presentation.xml"
    ] * 2
    same([row["uri"] for row in calls], uris, "historical call order")
    same([row["call"] for row in calls], list(range(1, 19)), "historical call ordinals")
    same(manifest["groups"]["all"]["count"], 18, "all group count")
    same(manifest["groups"]["presentation"]["count"], 5, "presentation group count")
    same(manifest["groups"]["slides"]["count"], 13, "slides group count")
    sequence_copy = P / "corpus" / "sequence.txt"
    need(sequence_copy)
    zip_member_by_uri = {row["uri"]: row["zip_member"] for row in members}
    same(
        sequence_copy.read_text().splitlines(),
        [zip_member_by_uri[row["uri"]] for row in calls],
        "plain sequence copy",
    )
    return manifest, expected


def validate_build(manifest: dict[str, Any]) -> tuple[Path, str]:
    build = read("build.json")
    same(build.get("exit_code"), 0, "build exit")
    same(build.get("head"), manifest["production_revision"]["commit"], "build revision")
    current_head = subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=ROOT, text=True).strip()
    same(current_head, manifest["production_revision"]["commit"], "current production revision")
    binary = packet_path(build["binary"])
    binary_hash = build["binary_sha256"]
    cleanup_path = P / "cleanup.json"
    if binary.is_file():
        same(sha(binary), binary_hash, "probe binary hash")
    elif cleanup_path.is_file():
        cleanup = read("cleanup.json")
        same(cleanup.get("binary_sha256"), binary_hash, "cleanup binary hash")
    else:
        stop("probe binary missing before cleanup")
    same(sha(ROOT / "Cargo.lock"), build["root_lock_sha256"], "workspace lock hash")
    for name, digest in build["source_sha256"].items():
        path = ROOT / name
        need(path)
        same(sha(path), digest, f"production source binding {name}")
    tracked = subprocess.check_output(["git", "ls-files", "crates"], cwd=ROOT, text=True).splitlines()
    expected_sources = {
        name
        for name in tracked
        if name.endswith(".rs") or Path(name).name == "Cargo.toml"
    }
    same(set(build["source_sha256"]), expected_sources, "production source census")
    probe_map = {
        str(path.relative_to(P)): sha(path)
        for path in (P / "probe").rglob("*")
        if path.is_file()
    }
    same(probe_map, build["probe_sha256"], "probe source bindings")
    return binary, binary_hash


def validate_native(manifest: dict[str, Any], expected: dict[str, list[str]], binary: Path, binary_hash: str) -> None:
    rows = read("runs.json")
    same(len(rows), 64, "native run count")
    by_case: dict[str, list[dict[str, Any]]] = {}
    identities_by_case: dict[str, list[list[dict[str, str]]]] = {}
    member_by_path = {row["corpus_path"]: row for row in manifest["members"]}
    output_lengths = {}
    for row in manifest["calls"]:
        output_lengths.setdefault(row["relative_path"], row["trace_output_len"])
    for row in rows:
        case, leg = row["case"], row["leg"]
        if case not in expected or leg not in range(4):
            stop(f"unexpected native row {case}/{leg}")
        path = P / row["output"]
        error = path.with_suffix(".stderr")
        need(path)
        need(error)
        same(sha(path), row["output_sha256"], f"native output hash {case}/{leg}")
        same(sha(error), row["stderr_sha256"], f"native stderr hash {case}/{leg}")
        same(row["exit_code"], 0, f"native exit {case}/{leg}")
        same(row["binary_sha256"], binary_hash, f"native binary binding {case}/{leg}")
        sequence_path = P / row["sequence"]
        same(row["sequence_sha256"], sha(sequence_path), f"native sequence hash {case}/{leg}")
        same(row["batch"], 4, f"native batch {case}/{leg}")
        same(row["samples"], 200, f"native sample count {case}/{leg}")
        command = row["command"]
        same(
            command,
            ["taskset", "-c", "12", str(binary), str(sequence_path), "10", "200", "4"],
            f"native command binding {case}/{leg}",
        )
        identities, meta, samples = parse_probe(path)
        same(meta, {"probe": "0697", "warmups": "10", "samples": "200", "batch": "4", "calls": str(len(expected[case]))}, f"native META {case}/{leg}")
        same(len(samples), 200, f"native samples {case}/{leg}")
        refs = expected[case]
        unique = []
        for ref in refs:
            if ref not in unique:
                unique.append(ref)
        expected_identity = []
        for index, ref in enumerate(unique):
            member = member_by_path[str((P / "sequences" / ref).resolve().relative_to(ROOT))]
            expected_identity.append({
                "index": str(index),
                "input_bytes": str(member["bytes"]),
                "input_sha256": member["sha256"],
                "output_bytes": str(output_lengths[member["corpus_path"]]),
                "borrowed": "false",
            })
        same([{k: identity[k] for k in expected_identity[0]} for identity in identities], expected_identity, f"native identities {case}/{leg}")
        for identity in identities:
            if int(identity["output_bytes"]) <= 0 or identity["borrowed"] != "false":
                stop(f"unexpected output ownership/length {case}/{leg}")
        same(row["identity"], ["IDENTITY\t" + "\t".join(f"{key}={value}" for key, value in identity.items()) for identity in identities], f"native receipt identities {case}/{leg}")
        normalized = [value / 4 for value in samples]
        same(row["stats"], stat(normalized), f"native stats {case}/{leg}")
        by_case.setdefault(case, []).append(row)
        identities_by_case.setdefault(case, []).append(identities)
    for case in expected:
        group = by_case.get(case, [])
        same(len(group), 4, f"native legs {case}")
        identities = [tuple(row["identity"]) for row in group]
        same(len(set(identities)), 1, f"native repeated identity {case}")
        same(
            len({tuple(item.items()) for item in identities_by_case[case][0]}),
            len(identities_by_case[case][0]),
            f"native unique identities {case}",
        )
    all_identity = identities_by_case["all"][0]
    all_by_input = {identity["input_sha256"]: identity for identity in all_identity}
    for case in ("presentation", "slides") + tuple(f"slide{i}" for i in range(1, 14)):
        for identity in identities_by_case[case][0]:
            reference = all_by_input.get(identity["input_sha256"])
            if reference is None:
                stop(f"identity is absent from all-sequence group: {case}")
            comparable = {key: identity[key] for key in identity if key != "index"}
            expected_comparable = {key: reference[key] for key in reference if key != "index"}
            same(comparable, expected_comparable, f"cross-group output identity {case}")
    summaries = read("summary.json")
    same({row["case"] for row in summaries}, set(expected), "native summary census")
    for summary in summaries:
        group = by_case[summary["case"]]
        medians = [row["stats"]["p50_ns"] for row in group]
        same(summary["leg_median_min_ns"], min(medians), f"summary min {summary['case']}")
        same(summary["leg_median_max_ns"], max(medians), f"summary max {summary['case']}")
        same(summary["median_of_leg_medians_ns"], stat(medians)["p50_ns"], f"summary median {summary['case']}")
        same(summary["leg_median_spread_pct"], (max(medians) / min(medians) - 1) * 100, f"summary spread {summary['case']}")


def parse_perf(path: Path) -> dict[str, tuple[float, float, str]]:
    values = {}
    for line in path.read_text().splitlines():
        fields = line.split("\t")
        if len(fields) < 5 or fields[2] not in EVENTS:
            continue
        values[fields[2]] = (float(fields[0]), float(fields[4]), fields[3])
    same(set(values), set(EVENTS), f"perf event set {path.name}")
    for event, (_value, running, _unit) in values.items():
        if running < 99.0:
            stop(f"multiplexed perf event {event} in {path.name}")
    return values


def validate_profile(binary: Path, binary_hash: str) -> None:
    rows = read("profile-runs.json")
    expected = []
    for repeat in range(3):
        cases = PROFILE_NAMES if repeat % 2 == 0 else list(reversed(PROFILE_NAMES))
        counts = [10, 210] if repeat % 2 == 0 else [210, 10]
        expected.extend(f"{case}-{repeat}-{count}" for case in cases for count in counts)
    expected += ["record", "self", "inclusive"]
    same([row["name"] for row in rows], expected, "profile run order")
    values = {}
    profile_identity_by_input: dict[str, dict[str, str]] = {}
    expected_calls = {"all": 18, "presentation": 5, "slides": 13}
    for row in rows:
        same(row["exit_code"], 0, f"profile exit {row['name']}")
        same(row["binary_sha256"], binary_hash, f"profile binary {row['name']}")
        out = P / "profile" / f"{row['name']}.stdout"
        err = P / "profile" / f"{row['name']}.stderr"
        need(out); need(err)
        same(sha(out), row["stdout_sha256"], f"profile stdout {row['name']}")
        same(sha(err), row["stderr_sha256"], f"profile stderr {row['name']}")
        if row["name"] in {"self", "inclusive"}:
            if not out.read_text().strip():
                stop(f"empty perf report {row['name']}")
            continue
        identities, meta, samples = parse_probe(out)
        if row["name"] == "record":
            same(meta, {"probe": "0697", "warmups": "10", "samples": "1000", "batch": "1", "calls": "18"}, "record META")
            same(len(samples), 1000, "record sample count")
            expected_command = [
                "perf", "record", "-F", "997", "-g", "--call-graph", "dwarf,16384",
                "-o", str(packet_path(read("profile-data.json")["path"])), "--", "taskset", "-c", "12",
                str(binary), str(P / "sequences/all.txt"), "10", "1000", "1",
            ]
            same(row["command"], expected_command, "perf record command")
            for identity in identities:
                previous = profile_identity_by_input.setdefault(identity["input_sha256"], identity)
                same(
                    {key: identity[key] for key in identity if key != "index"},
                    {key: previous[key] for key in previous if key != "index"},
                    "record output identity",
                )
            continue
        case, repeat, count = row["name"].rsplit("-", 2)
        sequence = P / "sequences" / f"{case}.txt"
        expected_command = [
            "perf", "stat", "-x", "\t", "-e", ",".join(EVENTS), "--",
            "taskset", "-c", "12", str(binary), str(sequence), "0", count, "10",
        ]
        same(row["command"], expected_command, f"profile command {row['name']}")
        same(meta, {"probe": "0697", "warmups": "0", "samples": count, "batch": "10", "calls": str(expected_calls[case])}, f"profile META {row['name']}")
        same(len(samples), int(count), f"profile samples {row['name']}")
        for identity in identities:
            previous = profile_identity_by_input.setdefault(identity["input_sha256"], identity)
            same(
                {key: identity[key] for key in identity if key != "index"},
                {key: previous[key] for key in previous if key != "index"},
                f"profile output identity {row['name']}",
            )
        values.setdefault((case, int(repeat)), {})[int(count)] = parse_perf(err)
    data_path = str(packet_path(read("profile-data.json")["path"]))
    record = next(row for row in rows if row["name"] == "record")
    for name, children, limit in (("self", "--no-children", "0.1"), ("inclusive", "--children", "1")):
        report = next(row for row in rows if row["name"] == name)
        same(report["command"], ["perf", "report", "--stdio", children, "--percent-limit", limit, "-i", data_path], f"perf report command {name}")
    summary = read("counter-summary.json")
    same(len(summary), 9, "counter summary rows")
    same(
        {(row["case"], row["repeat"]) for row in summary},
        {(case, repeat) for case in PROFILE_NAMES for repeat in range(3)},
        "counter summary census",
    )
    for row in summary:
        key = (row["case"], row["repeat"])
        same(set(row["counts"]), {"10", "210"}, f"counter count keys {key}")
        parsed = values[key]
        for count in (10, 210):
            expected_values = {event: parsed[count][event][0] for event in EVENTS}
            same(row["counts"][str(count)], expected_values, f"counter raw values {key}/{count}")
        slope = {event: (parsed[210][event][0] - parsed[10][event][0]) / 2000 for event in EVENTS}
        same(row["per_sequence"], slope, f"counter slope {key}")
        same(row["ipc"], slope["instructions"] / slope["cycles"], f"counter IPC {key}")
        same(row["task_clock_unit"], "milliseconds", f"task-clock unit {key}")
    profile_data = read("profile-data.json")
    data = packet_path(profile_data["path"])
    cleanup = P / "cleanup.json"
    if data.is_file():
        same(sha(data), profile_data["sha256"], "perf data hash")
    elif cleanup.is_file():
        same(read("cleanup.json").get("profile_data_sha256"), profile_data["sha256"], "cleanup perf data hash")
    else:
        stop("perf data missing before cleanup")


def validate_cleanup() -> None:
    path = P / "cleanup.json"
    if not path.is_file():
        return
    cleanup = read("cleanup.json")
    required = {str(ROOT.parent / "litchi-target-0697"), str(ROOT.parent / "litchi-0697-profile")}
    rows = cleanup.get("removed", [])
    same({row["path"] for row in rows}, required, "cleanup paths")
    for row in rows:
        if row.get("removed") is not True or Path(row["path"]).exists():
            stop(f"cleanup removal not proven: {row.get('path')}")
    same(sha(ROOT / "Cargo.lock"), cleanup["root_lock_sha256"], "cleanup root lock")


def validate_topology() -> None:
    binding = read("topology-binding.json")
    same(binding.get("historical_trace", {}).get("change"), "0693", "historical topology change")
    same(binding.get("historical_trace", {}).get("commit"), "829bed696500c8ad2f461156985ff1bcf075346f", "historical topology commit")
    same(binding.get("current_candidate", {}).get("change"), "0696", "current topology change")
    same(binding.get("current_candidate", {}).get("commit"), "1bf58ace2c2d69db4ab88e21f62aff5f1ba7cf1e", "current topology commit")
    trees = binding.get("pptx_git_tree", {})
    same(set(trees), {
        "829bed696500c8ad2f461156985ff1bcf075346f",
        "1bf58ace2c2d69db4ab88e21f62aff5f1ba7cf1e",
    }, "PPTX topology commits")
    for commit, expected_tree in trees.items():
        actual_tree = subprocess.check_output(
            ["git", "rev-parse", f"{commit}:crates/litchi-pptx"], cwd=ROOT, text=True
        ).strip()
        same(actual_tree, expected_tree, f"PPTX tree {commit}")
    names = subprocess.check_output(
        ["git", "ls-tree", "-r", "--name-only", "1bf58ace2c2d69db4ab88e21f62aff5f1ba7cf1e", "crates/litchi-pptx"],
        cwd=ROOT,
        text=True,
    ).splitlines()
    for name in names:
        current = ROOT / name
        need(current)
        expected = subprocess.check_output(["git", "show", f"1bf58ace2c2d69db4ab88e21f62aff5f1ba7cf1e:{name}"], cwd=ROOT)
        same(current.read_bytes(), expected, f"current PPTX source {name}")


def validate_environment() -> None:
    environment = read("environment.json")
    records = environment.get("records")
    if not isinstance(records, list) or not records:
        stop("environment records")
    expected_commands = {
        ("uname", "-a"),
        ("lscpu",),
        ("rustc", "-Vv"),
        ("cargo", "-V"),
        ("perf", "--version"),
        ("python3", "--version"),
        ("ldd", "--version"),
        ("free", "-b"),
        ("df", "-T", str(P)),
        ("objdump", "--version"),
    }
    actual_commands = set()
    for row in records:
        command = tuple(row.get("command", []))
        actual_commands.add(command)
        same(row.get("exit_code"), 0, f"environment command {command}")
        if not isinstance(row.get("stdout"), str) or not isinstance(row.get("stderr"), str):
            stop(f"environment output {command}")
    same(actual_commands, expected_commands, "environment command census")
    allowed = environment.get("allowed_cpus")
    same(environment.get("measured_cpu"), 12, "measured CPU")
    if not isinstance(allowed, list) or 12 not in allowed or any(not isinstance(cpu, int) for cpu in allowed):
        stop("environment CPU affinity")
    same(environment.get("allocator"), "Rust default system allocator; no allocator instrumentation", "allocator binding")
    same(environment.get("host"), "shared; warm processing; no quiescence or cold-cache claim", "host scope")


def validate_gates_and_scripts() -> None:
    constraints = read("constraints.json")
    same(constraints.get("previously_read"), True, "constraint receipt")
    same(len(constraints.get("sha256", {})), 33, "constraint count")
    for name, digest in constraints["sha256"].items():
        same(sha(ROOT / name), digest, f"constraint hash {name}")
    validation = read("validation.json")
    names = [row["name"] for row in validation]
    expected_inherited = ["probe-fmt", "probe-clippy", "probe-doc", "crate-boundaries", "claims", "claims-structural", "report", "coverage", "non-iwork"]
    same(names, expected_inherited, "validation gate census")
    same(len(validation), 9, "validation gate count")
    for row in validation:
        same(row["exit_code"], 0, f"validation gate {row['name']}")
        log = P / "validation" / f"{row['name']}.log"
        need(log)
        same(sha(log), row["log_sha256"], f"validation log {row['name']}")
    scripts = P / "script-hashes.json"
    if scripts.is_file():
        manifest = read("script-hashes.json")
        expected = {str(path.relative_to(P)) for path in P.rglob("*.py")}
        same(set(manifest), expected, "script hash census")
        for name, digest in manifest.items():
            path = P / name
            need(path)
            same(sha(path), digest, f"script hash {name}")
            ast.parse(path.read_text(), filename=str(path))


def validate_instructions() -> None:
    """Run the independently owned instruction-evidence verifier when present."""
    verifier = P / "audit-instructions.py"
    if not verifier.is_file():
        stop("missing instruction-evidence verifier")
    result = subprocess.run(
        [sys.executable, str(verifier)],
        cwd=ROOT,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )
    same(result.returncode, 0, "instruction evidence verifier")


def rerun_deterministic_scripts() -> None:
    for name in ("report-metrics.py", "count-elements.py", "summarize-instructions.py"):
        output_names = {
            "report-metrics.py": ("tables.md", "attribution.json"),
            "count-elements.py": ("element-counts.json",),
            "summarize-instructions.py": ("instruction-summary.json",),
        }[name]
        before = {output: sha(P / output) for output in output_names}
        env = {**os.environ, "PYTHONDONTWRITEBYTECODE": "1"}
        result = subprocess.run(
            [os.environ.get("PYTHON", "python3"), str(P / name)],
            cwd=ROOT,
            env=env,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
        )
        same(result.returncode, 0, f"deterministic script {name}")
        for output, digest in before.items():
            same(sha(P / output), digest, f"deterministic output {output}")


def validate_report() -> None:
    report = need(ROOT / "docs/performance/0697-mce-context-ownership-attribution.md").read_text()
    if "performance_claim: none" not in report or "not additive capture fractions" not in report:
        stop("report does not state attribution scope")
    metrics = read("attribution.json")
    same(metrics.get("performance_claim"), "none", "packet performance claim")
    summary = read("summary.json")
    all_median = next(row["median_of_leg_medians_ns"] for row in summary if row["case"] == "all")
    expected_ratios = {
        name: next(row["median_of_leg_medians_ns"] for row in summary if row["case"] == name) / all_median
        for name in ("presentation", "slides", "slide11")
    }
    expected_ratios["sum_individual_over_full"] = expected_ratios["presentation"] + sum(
        next(row["median_of_leg_medians_ns"] for row in summary if row["case"] == f"slide{i}") / all_median
        for i in range(1, 14)
    )
    expected_ratios["sum_groups_over_full"] = expected_ratios["presentation"] + expected_ratios["slides"]
    same(metrics["denominator"], "median of four leg medians of isolated all-sequence time; not capture latency", "attribution denominator")
    same(metrics["ratios"], expected_ratios, "attribution ratios")
    same(metrics["max_leg_median_spread_pct"], max(row["leg_median_spread_pct"] for row in summary), "attribution spread")
    elements = read("element-counts.json")
    same(elements.get("method"), "Python Expat raw QName start events; source syntax, not MCE branch execution instrumentation", "element census method")
    same(len(elements.get("members", [])), len(SEQUENCE_NAMES) - 3 + 1, "element census member count")
    member_uris = {row["uri"] for row in elements["members"]}
    manifest_rows = manifest_members()
    same(member_uris, {row["uri"] for row in manifest_rows}, "element census member identities")
    manifest_by_uri = {row["uri"]: row for row in manifest_rows}
    for row in elements["members"]:
        source = ROOT / manifest_by_uri[row["uri"]]["corpus_path"]
        same(row["sha256"], sha(source), f"element census source hash {row['uri']}")
        for key in ("starts", "starts_with_namespace_declarations", "namespace_declarations"):
            if not isinstance(row.get(key), int) or row[key] < 0:
                stop(f"element census member value {row['uri']}/{key}")
    weighted = elements.get("weighted_capture_sequence")
    if not isinstance(weighted, dict) or set(weighted) != {"starts", "starts_with_namespace_declarations", "namespace_declarations"}:
        stop("element census weighted summary")
    if any(not isinstance(value, int) or value < 0 for value in weighted.values()):
        stop("element census weighted values")
    rows = ["| Case | Calls per sequence | Input bytes | Median range (µs) |", "| --- | ---: | ---: | ---: |"]
    for name in SEQUENCE_NAMES:
        sequence = P / "sequences" / f"{name}.txt"
        sources = [(sequence.parent / line).resolve() for line in sequence.read_text().splitlines()]
        item = next(row for row in summary if row["case"] == name)
        rows.append(f"| {name} | {len(sources)} | {sum(path.stat().st_size for path in sources):,} | {item['leg_median_min_ns']/1000:.3f}–{item['leg_median_max_ns']/1000:.3f} |")
    same((P / "tables.md").read_text(), "\n".join(rows) + "\n", "report table receipt")


def manifest_members() -> list[dict[str, Any]]:
    return read("corpus/manifest.json")["members"]


def main() -> None:
    manifest, expected = validate_corpus()
    binary, binary_hash = validate_build(manifest)
    validate_native(manifest, expected, binary, binary_hash)
    validate_profile(binary, binary_hash)
    validate_topology()
    validate_environment()
    validate_gates_and_scripts()
    validate_report()
    rerun_deterministic_scripts()
    validate_instructions()
    validate_cleanup()
    print("PASS: 0697 corpus, 18-call sequence, native samples, profile counters, raw summaries, instruction evidence, gates and cleanup bindings")


if __name__ == "__main__":
    main()
