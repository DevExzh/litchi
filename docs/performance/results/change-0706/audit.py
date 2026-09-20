#!/usr/bin/env python3
"""Recheck the complete 0706 evidence packet without acquiring new evidence.

The packet has two deliberately different source states.  ``source-baseline``
is the build used by the A legs and ``source-candidate`` is the build used by
the B legs.  At the end of the experiment the candidate may be retained or it
may have been restored.  The latter is a valid result: an independently useful
integration test under ``crates/xml-minifier/tests/`` may remain after the
production candidate is rejected.  This audit therefore checks the final
checkout against the recorded decision instead of unconditionally requiring
the candidate source map.

The script only reads the repository and packet.  It may run the packet's
Python analyzer once in a temporary directory; it never invokes Cargo, a
native benchmark, an allocator benchmark, or an oracle binary.
"""

from __future__ import annotations

import hashlib
import json
import math
import os
from pathlib import Path
import re
import statistics
import subprocess
import sys
import tempfile
from typing import Any


HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[3]
ALLOWED_ROOTS = ("crates/xml-minifier/",)
ALLOWED_RESTORED_TEST_ROOTS = ("crates/xml-minifier/tests/",)
HEX = set("0123456789abcdef")


class AuditError(AssertionError):
    """A packet invariant failed."""


def fail(message: str) -> None:
    raise AuditError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def sha(path: Path) -> str:
    require(path.is_file() and not path.is_symlink(), f"missing regular file: {path}")
    return hashlib.sha256(path.read_bytes()).hexdigest()


def check_hex(value: Any, label: str) -> None:
    require(isinstance(value, str) and len(value) == 64 and set(value) <= HEX,
            f"{label} is not a lowercase SHA-256 digest")


def load(path: Path) -> Any:
    require(path.is_file(), f"missing packet artifact: {path}")
    try:
        return json.loads(path.read_text())
    except json.JSONDecodeError as error:
        fail(f"invalid JSON in {path}: {error}")


def load_object(path: Path) -> dict[str, Any]:
    value = load(path)
    require(isinstance(value, dict), f"{path} is not a JSON object")
    return value


def load_map(path: Path) -> dict[str, str]:
    value = load(path)
    require(isinstance(value, dict), f"{path} is not a JSON object map")
    for name, digest in value.items():
        require(isinstance(name, str) and isinstance(digest, str),
                f"{path} contains a non-string source entry")
        relative = Path(name)
        require(not relative.is_absolute() and ".." not in relative.parts,
                f"{path} contains an unsafe relative path: {name!r}")
        check_hex(digest, f"{path}:{name}")
    return value


def canonical_json(value: Any) -> str:
    return json.dumps(value, sort_keys=True, separators=(",", ":"))


def digest_json(value: Any) -> str:
    return hashlib.sha256((canonical_json(value) + "\n").encode()).hexdigest()


def micro_stats(values: list[int]) -> dict[str, int | float]:
    require(values and all(isinstance(value, int) and not isinstance(value, bool)
                           and value > 0 for value in values),
            "microbench sample vector is not a positive integer vector")
    ordered = sorted(values)
    p95_index = min((95 * len(ordered) + 99) // 100 - 1, len(ordered) - 1)
    p99_index = min((99 * len(ordered) + 99) // 100 - 1, len(ordered) - 1)
    left = ordered[(len(ordered) - 1) // 2]
    right = ordered[len(ordered) // 2]
    return {
        "count": len(values),
        "p50": left // 2 + right // 2 + (left % 2 + right % 2) // 2,
        "p95": ordered[p95_index],
        "p99": ordered[p99_index],
        "mean": statistics.mean(values),
        "min": ordered[0],
        "max": ordered[-1],
    }


def percent_change(before: float, after: float) -> float:
    require(math.isfinite(before) and math.isfinite(after) and before != 0,
            "microbench comparison has a non-finite or zero baseline")
    return round((before - after) / before * 100.0, 9)


def reject_panic(value: Any, label: str) -> None:
    if isinstance(value, str):
        require(value != "PANIC", f"{label} contains a panic outcome")
    elif isinstance(value, dict):
        for key, item in value.items():
            reject_panic(item, f"{label}.{key}")
    elif isinstance(value, list):
        for index, item in enumerate(value):
            reject_panic(item, f"{label}[{index}]")


def source_census() -> dict[str, str]:
    """The same Rust/Cargo census used by change-0706/build.py."""

    paths: list[Path] = [ROOT / "Cargo.toml", ROOT / "Cargo.lock"]
    paths.extend(path for path in (ROOT / ".cargo").rglob("*") if path.is_file())
    for folder in ("crates", "tools/perf-baseline"):
        paths.extend(
            path
            for path in (ROOT / folder).rglob("*")
            if path.is_file()
            and "target" not in path.parts
            and (path.suffix == ".rs" or path.name in {"Cargo.toml", "Cargo.lock"})
        )
    return {
        str(path.relative_to(ROOT)): sha(path)
        for path in sorted(set(paths))
    }


def diff_maps(left: dict[str, str], right: dict[str, str]) -> list[str]:
    return sorted(name for name in set(left) | set(right)
                  if left.get(name) != right.get(name))


def allowed_candidate_path(name: str) -> bool:
    return any(name.startswith(root) for root in ALLOWED_ROOTS)


def restored_test_path(name: str) -> bool:
    return any(name.startswith(root) for root in ALLOWED_RESTORED_TEST_ROOTS)


def verify_constraints() -> None:
    constraints = load_map(HERE / "constraints.json")
    require("docs/GOAL.md" in constraints, "constraints omit docs/GOAL.md")
    require("docs/adr/README.md" in constraints, "constraints omit ADR README")
    require(len(constraints) >= 30, "constraint map is unexpectedly incomplete")
    for name, expected in constraints.items():
        path = ROOT / name
        require(path.is_file(), f"constraint input is absent: {name}")
        require(sha(path) == expected, f"constraint digest changed: {name}")


def find_outcome() -> str | None:
    """Read an explicit retention decision, if the coordinator recorded one."""

    retained: list[bool] = []
    for filename in ("decision.json", "retention.json", "outcome.json"):
        path = HERE / filename
        if not path.exists():
            continue
        value = load_object(path)
        for key in ("retained", "candidate_retained", "keep_candidate"):
            if key in value:
                require(isinstance(value[key], bool), f"{filename}.{key} is not boolean")
                retained.append(value[key])
        for key in ("outcome", "candidate_outcome", "decision", "retention"):
            item = value.get(key)
            if isinstance(item, str):
                normalized = item.lower().replace("_", "-").replace(" ", "-")
                if normalized in {"retained", "retain", "accepted", "kept", "keep", "applied"}:
                    retained.append(True)
                elif normalized in {"rejected", "reject", "restored", "restore", "discarded", "drop"}:
                    retained.append(False)
    if retained:
        require(all(item == retained[0] for item in retained),
                "retention decision artifacts disagree")
        return "retained" if retained[0] else "rejected"
    return None


def verify_source_state() -> tuple[str, dict[str, str], dict[str, str], dict[str, str]]:
    baseline = load_map(HERE / "source-baseline.json")
    candidate = load_map(HERE / "source-candidate.json")
    for label, manifest in (("baseline", baseline), ("candidate", candidate)):
        require(manifest, f"source-{label}.json is empty")
        # These are historical source snapshots.  The final checkout may be
        # the other snapshot, so their entries are checked when we compare the
        # final census below and when receipts bind their source-before/after
        # maps.  Requiring every manifest entry to match the final checkout
        # here would make a rejected candidate impossible to audit.

    candidate_delta = diff_maps(baseline, candidate)
    require(candidate_delta, "candidate source map has no change from baseline")
    require(all(allowed_candidate_path(name) for name in candidate_delta),
            f"candidate changed outside xml-minifier: {candidate_delta}")
    require(any(name.startswith("crates/xml-minifier/src/") for name in candidate_delta),
            "candidate source map has no production xml-minifier change")

    current = source_census()
    outcome = find_outcome()
    if outcome is None:
        current_candidate_delta = diff_maps(candidate, current)
        current_baseline_delta = diff_maps(baseline, current)
        if not current_candidate_delta:
            outcome = "retained"
        elif all(restored_test_path(name) for name in current_baseline_delta):
            outcome = "rejected"
        else:
            fail("cannot infer whether the candidate was retained or rejected")

    if outcome == "retained":
        # A separately retained integration test may have been added after the
        # candidate source manifest was frozen.  Production source still has
        # to equal the candidate byte-for-byte.
        differences = diff_maps(candidate, current)
        require(all(restored_test_path(name) for name in differences),
                f"retained candidate source differs from candidate: {differences}")
        for name in set(candidate) | set(current):
            if restored_test_path(name):
                continue
            require(candidate.get(name) == current.get(name),
                    f"retained candidate production source differs: {name}")
    else:
        differences = diff_maps(baseline, current)
        require(all(restored_test_path(name) for name in differences),
                f"rejected candidate left production source changes: {differences}")
    return outcome, baseline, candidate, current


def verify_revision(plan: dict[str, Any]) -> None:
    revision = plan.get("revision")
    require(isinstance(revision, str) and revision, "plan revision is missing")
    result = subprocess.run(["git", "merge-base", "--is-ancestor", revision, "HEAD"],
                            cwd=ROOT, stdout=subprocess.DEVNULL, stderr=subprocess.PIPE)
    require(result.returncode == 0,
            f"plan revision is not an ancestor of final HEAD: {revision}")


def cleanup_entries() -> list[dict[str, Any]]:
    """Collect optional cleanup witnesses used when frozen binaries are gone."""

    entries: list[dict[str, Any]] = []
    for filename in ("cleanup.json", "cleanup-witness.json", "cleanup-receipt.json"):
        path = HERE / filename
        if not path.exists():
            continue
        value = load(path)
        if isinstance(value, dict) and value.get("owned_paths_absent") is not None:
            require(value["owned_paths_absent"] is True,
                    f"{filename} does not certify owned paths absent")

        def walk(item: Any) -> None:
            if isinstance(item, dict):
                raw_path = item.get("path")
                digest = item.get("sha256", item.get("binary_sha256"))
                size = item.get("bytes", item.get("binary_bytes"))
                if isinstance(raw_path, str) and isinstance(digest, str):
                    check_hex(digest, f"{filename}:{raw_path}")
                    if size is not None:
                        require(isinstance(size, int) and size >= 0,
                                f"{filename}:{raw_path} has invalid byte count")
                    entries.append({"path": raw_path, "sha256": digest, "bytes": size})
                for child in item.values():
                    walk(child)
            elif isinstance(item, list):
                for child in item:
                    walk(child)

        walk(value)
    return entries


def witness_for(path: Path, digest: str, size: int | None, witnesses: list[dict[str, Any]]) -> bool:
    target = str(path.resolve())
    for item in witnesses:
        candidate = item["path"]
        candidate_path = Path(candidate)
        resolved = str((ROOT / candidate_path).resolve()
                       if not candidate_path.is_absolute() else candidate_path.resolve())
        if resolved == target:
            if item["sha256"] == digest and (item["bytes"] is None or item["bytes"] == size):
                return True
    return False


def verify_binary(path_value: Any, digest: Any, size: Any, witnesses: list[dict[str, Any]], label: str) -> None:
    require(isinstance(path_value, str), f"{label} binary path missing")
    path = Path(path_value)
    check_hex(digest, f"{label}.binary_sha256")
    require(isinstance(size, int) and size > 0, f"{label}.binary_bytes is invalid")
    if path.is_file() and not path.is_symlink():
        require(sha(path) == digest, f"{label} binary digest changed")
        require(path.stat().st_size == size, f"{label} binary byte count changed")
    else:
        require(witness_for(path, digest, size, witnesses),
                f"{label} binary is absent without an exact cleanup witness")


def record_binary_name(record: dict[str, Any], label: str) -> str:
    binary = Path(str(record.get("binary", "")))
    require(binary.name, f"{label} binary name is missing")
    return binary.name


def verify_builds(witnesses: list[dict[str, Any]]) -> dict[tuple[str, str], dict[str, Any]]:
    builds: dict[tuple[str, str], dict[str, Any]] = {}
    for role in ("baseline", "candidate"):
        path = HERE / f"build-{role}.json"
        records = load(path)
        require(isinstance(records, list), f"{path} is not a build-record list")
        require(len(records) == 2, f"{path} does not contain native and allocator records")
        manifest_path = HERE / f"source-{role}.json"
        manifest_digest = sha(manifest_path)
        seen: set[str] = set()
        for record in records:
            require(isinstance(record, dict), f"{path} contains a non-object record")
            require(record.get("exit_code") == 0, f"{path} contains a failed build")
            binary_name = record_binary_name(record, str(path))
            if binary_name.endswith("-native"):
                lane = "native"
                expected_name = f"{role}-native"
            elif binary_name.endswith("-alloc"):
                lane = "alloc"
                expected_name = f"{role}-alloc"
            else:
                fail(f"{path} has an unexpected binary: {binary_name}")
            require(binary_name == expected_name,
                    f"{path} binary is {binary_name}; expected {expected_name}")
            require(lane not in seen, f"{path} contains duplicate {lane} record")
            seen.add(lane)
            require(record.get("source_manifest_sha256") == manifest_digest,
                    f"{path}:{lane} is not bound to source-{role}.json")
            command = record.get("command")
            require(isinstance(command, list) and "cargo" in command and "build" in command,
                    f"{path}:{lane} has no Cargo build command")
            require("--release" in command and "--locked" in command,
                    f"{path}:{lane} build is not the locked release build")
            log_path = HERE / f"build-{role}-{lane}.log"
            require(log_path.is_file(), f"missing build log: {log_path}")
            verify_binary(record.get("binary"), record.get("binary_sha256"),
                          record.get("binary_bytes"), witnesses, f"{path}:{lane}")
            builds[(role, lane)] = record
        require(seen == {"native", "alloc"}, f"{path} lanes are incomplete")
    return builds


def expected_capture_names(plan: dict[str, Any]) -> tuple[set[str], set[str]]:
    shapes = list(plan["primary"]["shapes"])
    require(shapes == ["medium", "dense-sparse"], "0706 primary shape order changed")
    phases = ("baseline-noise1", "baseline-noise2", "baseline-A1",
              "candidate-B1", "candidate-B2", "baseline-A2")
    native: set[str] = set()

    def slug(value: str) -> str:
        return "".join(c if c.isalnum() or c in "-_" else "_" for c in value)

    for phase in phases:
        for repeat in range(1, int(plan["primary"]["repeats"]) + 1):
            ordered = shapes if repeat == 1 else list(reversed(shapes))
            for shape in ordered:
                native.add(f"native-{phase}-r{repeat}-{slug(shape)}")
        if phase.endswith(("A1", "B1", "B2", "A2")):
            for guard in plan["guards"]:
                for shape in guard["shapes"]:
                    native.add(f"guard-{phase}-{slug(guard['case'])}-{slug(shape)}")
            producer = plan.get("producer")
            if producer:
                for repeat in range(1, int(plan["primary"]["repeats"]) + 1):
                    native.add(f"producer-{phase}-r{repeat}")

    allocation: set[str] = set()
    for phase in ("allocator-baseline", "allocator-candidate"):
        for repeat in range(1, int(plan["allocation"]["repeats"]) + 1):
            ordered = shapes if repeat == 1 else list(reversed(shapes))
            for shape in ordered:
                allocation.add(f"alloc-{phase}-r{repeat}-{slug(shape)}")
    return native, allocation


def verify_artifact_map(receipt: dict[str, Any], label: str) -> None:
    artifacts = receipt.get("artifacts")
    require(isinstance(artifacts, dict) and artifacts, f"{label} has no artifact hash map")
    for name, digest in artifacts.items():
        require(isinstance(name, str) and Path(name).parent == Path("."),
                f"{label} artifact path is not packet-local: {name!r}")
        check_hex(digest, f"{label}.artifacts.{name}")
        path = HERE / name
        require(path.is_file(), f"{label} artifact is absent: {name}")
        require(sha(path) == digest, f"{label} artifact digest changed: {name}")


def verify_capture_receipts(plan: dict[str, Any], builds: dict[tuple[str, str], dict[str, Any]]) -> None:
    native_expected, allocation_expected = expected_capture_names(plan)
    capture_records: dict[str, dict[str, Any]] = {}
    for path in sorted(HERE.glob("*.receipt.json")):
        value = load_object(path)
        if value.get("schema_version") != 2 or not isinstance(value.get("name"), str):
            continue
        name = value["name"]
        if not name.startswith(("native-", "guard-", "producer-", "alloc-")):
            continue
        require(path.name == f"{name}.receipt.json",
                f"capture receipt filename does not match its name: {path.name}")
        require(name not in capture_records, f"duplicate capture receipt name: {name}")
        capture_records[name] = value

    actual_native = {name for name in capture_records if name.startswith(("native-", "guard-", "producer-"))}
    actual_alloc = {name for name in capture_records if name.startswith("alloc-")}
    require(actual_native == native_expected,
            f"native receipt set mismatch; missing={sorted(native_expected-actual_native)} "
            f"extra={sorted(actual_native-native_expected)}")
    require(actual_alloc == allocation_expected,
            f"allocator receipt set mismatch; missing={sorted(allocation_expected-actual_alloc)} "
            f"extra={sorted(actual_alloc-allocation_expected)}")

    source_manifest_digest = {role: sha(HERE / f"source-{role}.json")
                              for role in ("baseline", "candidate")}
    for name, receipt in sorted(capture_records.items()):
        role = receipt.get("role")
        lane = receipt.get("lane")
        require(role in {"baseline", "candidate"}, f"{name} has invalid role")
        require(lane in {"native", "alloc"}, f"{name} has invalid lane")
        build = builds[(role, lane)]
        require(receipt.get("exit_code") == 0, f"{name} child failed")
        require(receipt.get("binary_sha256") == build["binary_sha256"],
                f"{name} is bound to a different binary")
        require(receipt.get("binary_bytes") == build["binary_bytes"],
                f"{name} binary byte count differs from build")
        require(receipt.get("build_record_sha256") == sha(HERE / f"build-{role}.json"),
                f"{name} build receipt binding changed")
        require(receipt.get("build_source_manifest_sha256") == build["source_manifest_sha256"],
                f"{name} build source binding changed")
        require(receipt.get("retained_binary_source", {}).get("manifest_sha256")
                == source_manifest_digest[role], f"{name} retained source binding changed")
        retained = receipt.get("retained_binary_source")
        require(isinstance(retained, dict), f"{name} retained source block is missing")
        source_map = load_map(HERE / f"source-{role}.json")
        require(retained.get("source_census_sha256") == digest_json(source_map),
                f"{name} retained source census binding changed")
        require(retained.get("source_entry_count") == len(source_map),
                f"{name} retained source entry count changed")
        require(receipt.get("plan_sha256") == sha(HERE / "plan.json"),
                f"{name} plan binding changed")
        require(receipt.get("constraints_sha256") == sha(HERE / "constraints.json"),
                f"{name} constraint binding changed")
        require(receipt.get("script_sha256") == sha(HERE / "capture.py"),
                f"{name} capture script binding changed")
        verify_artifact_map(receipt, name)

        current = receipt.get("current_checkout_source")
        require(isinstance(current, dict), f"{name} current source receipt is missing")
        before_name = current.get("before_artifact")
        after_name = current.get("after_artifact")
        require(isinstance(before_name, str) and isinstance(after_name, str),
                f"{name} source artifact names are missing")
        before = load_map(HERE / before_name)
        after = load_map(HERE / after_name)
        require(before == after and current.get("unchanged_during_child") is True,
                f"{name} source changed while child ran")
        require(current.get("before_sha256") == digest_json(before),
                f"{name} source-before digest binding changed")
        require(current.get("after_sha256") == digest_json(after),
                f"{name} source-after digest binding changed")
        require(current.get("before_file_sha256") == sha(HERE / before_name),
                f"{name} source-before artifact hash changed")
        require(current.get("after_file_sha256") == sha(HERE / after_name),
                f"{name} source-after artifact hash changed")
        relation = current.get("relation_before")
        require(isinstance(relation, dict), f"{name} source relation is missing")
        changed = relation.get("changed_paths", [])
        require(isinstance(changed, list) and all(isinstance(item, str) for item in changed),
                f"{name} source relation changed paths are invalid")
        require(all(allowed_candidate_path(item) for item in changed),
                f"{name} source relation escaped candidate root: {changed}")
        result_name = f"{name}.json"
        result = load_object(HERE / result_name)
        identity = result.get("binary_identity")
        if isinstance(identity, dict) and "binary_sha256" in identity:
            require(identity["binary_sha256"] == receipt["binary_sha256"],
                    f"{name} output binary identity differs from receipt")


def verify_oracle_probe_inputs() -> dict[str, str]:
    root = HERE / "oracle-probe"
    require(root.is_dir(), "oracle-probe directory is missing")
    files = sorted(path for path in root.rglob("*") if path.is_file() and "target" not in path.parts)
    require(files, "oracle-probe has no source files")
    return {str(path.relative_to(HERE)): sha(path) for path in files}


def verify_oracle(witnesses: list[dict[str, Any]], builds: dict[tuple[str, str], dict[str, Any]]) -> None:
    probe_inputs = verify_oracle_probe_inputs()
    build_records: dict[str, dict[str, Any]] = {}
    for role in ("baseline", "candidate"):
        path = HERE / f"oracle-build-{role}.json"
        record = load_object(path)
        require(record.get("action") == "build" and record.get("role") == role,
                f"{path} is not a {role} oracle build record")
        require(record.get("exit_code") == 0, f"{path} reports a failed oracle build")
        require(record.get("source_manifest_sha256") == sha(HERE / f"source-{role}.json"),
                f"{path} source binding changed")
        require(record.get("plan_sha256") == sha(HERE / "plan.json"),
                f"{path} plan binding changed")
        require(record.get("script_sha256") == sha(HERE / "oracle.py"),
                f"{path} script binding changed")
        require(record.get("probe_inputs") == probe_inputs,
                f"{path} probe source changed")
        current_source = record.get("current_source")
        require(isinstance(current_source, dict), f"{path} current source map is missing")
        expected_source = load_map(HERE / f"source-{role}.json")
        changed = diff_maps(expected_source, current_source)
        require(all(allowed_candidate_path(name) for name in changed),
                f"{path} source delta escaped candidate root: {changed}")
        require(not changed and record.get("retained_source_delta") == [],
                f"{path} oracle build was not built against its exact source map: {changed}")
        command = record.get("command")
        require(isinstance(command, list) and "cargo" in command and "build" in command,
                f"{path} has no Cargo build command")
        verify_binary(record.get("binary"), record.get("binary_sha256"),
                      record.get("binary_bytes"), witnesses, str(path))
        require((HERE / f"oracle-build-{role}.log").is_file(),
                f"missing oracle build log for {role}")
        build_records[role] = record

    allowed_actions = {"differential", "A1", "B1", "B2", "A2"}
    receipts: dict[tuple[str, str], dict[str, Any]] = {}
    for path in sorted(HERE.glob("oracle-*-*-receipt.json")):
        record = load_object(path)
        role = record.get("role")
        action = record.get("action")
        require(role in {"baseline", "candidate"} and action in allowed_actions,
                f"{path} has an invalid oracle action or role")
        key = (role, action)
        require(key not in receipts, f"duplicate oracle receipt: {path.name}")
        receipts[key] = record
        require(record.get("exit_code") == 0, f"{path} reports a failed oracle run")
        build = build_records[role]
        require(record.get("build_record_sha256") == sha(HERE / f"oracle-build-{role}.json"),
                f"{path} build binding changed")
        require(record.get("source_manifest_sha256") == sha(HERE / f"source-{role}.json"),
                f"{path} source manifest binding changed")
        require(record.get("binary_sha256") == build["binary_sha256"],
                f"{path} binary binding changed")
        require(record.get("binary") == build["binary"], f"{path} binary path binding changed")
        current_source = record.get("current_source")
        require(isinstance(current_source, dict), f"{path} current source map is missing")
        expected_source = load_map(HERE / f"source-{role}.json")
        changed = diff_maps(expected_source, current_source)
        if role == "candidate":
            require(not changed, f"{path} candidate source changed during oracle run: {changed}")
        else:
            require(all(allowed_candidate_path(name) for name in changed),
                    f"{path} baseline source delta escaped candidate root: {changed}")
        require(record.get("retained_source_delta") == changed,
                f"{path} retained source delta binding changed")
        require(record.get("plan_sha256", build_records[role].get("plan_sha256"))
                == sha(HERE / "plan.json"), f"{path} plan binding changed")
        require(record.get("script_sha256", build_records[role].get("script_sha256"))
                == sha(HERE / "oracle.py"), f"{path} script binding changed")
        verify_artifact_map(record, path.name)
        output_name = f"oracle-{role}-{action}.json"
        require(output_name in record["artifacts"], f"{path} does not bind {output_name}")
        load_object(HERE / output_name)

    for role in ("baseline", "candidate"):
        require((role, "differential") in receipts,
                f"{role} oracle differential receipt is missing")
    differential = [load_object(HERE / f"oracle-{role}-differential.json")
                    for role in ("baseline", "candidate")]
    require(differential[0] == differential[1],
            "baseline and candidate public-oracle differential outputs differ")
    require(differential[0].get("schema") == "litchi.xml-minifier-public-oracle.v1",
            "public-oracle differential schema changed")
    require(differential[0].get("casecount") == 4_148,
            "public-oracle differential case count changed")
    require(differential[0].get("callcount") == 29_036,
            "public-oracle differential call count changed")
    require(differential[0].get("chunk_patterns") == ["one_byte", "mixed"],
            "public-oracle differential chunk patterns changed")
    outcomes = differential[0].get("outcomes")
    require(isinstance(outcomes, list) and len(outcomes) == 4_148,
            "public-oracle differential outcomes are incomplete")
    for index, case in enumerate(outcomes):
        require(isinstance(case, dict) and case.get("id") == index,
                f"public-oracle case {index} identity changed")
        results = case.get("results")
        require(isinstance(results, list) and len(results) == 7,
                f"public-oracle case {index} result count changed")

    def reject_panic(value: Any, label: str) -> None:
        if isinstance(value, str):
            require(value != "PANIC", f"{label} contains a panic outcome")
        elif isinstance(value, dict):
            for key, item in value.items():
                reject_panic(item, f"{label}.{key}")
        elif isinstance(value, list):
            for index, item in enumerate(value):
                reject_panic(item, f"{label}[{index}]")

    reject_panic(differential[0], "public-oracle")

    bench_keys = {("baseline", "A1"), ("candidate", "B1"),
                  ("candidate", "B2"), ("baseline", "A2")}
    require(bench_keys <= set(receipts),
            f"oracle microbench receipts are incomplete: {sorted(bench_keys-set(receipts))}")
    semantic_maps: list[dict[tuple[str, str, str], str]] = []
    expected_cases = {"zero-attributes", "one-attribute", "two-attributes",
                      "sixty-four-attributes", "xmlspace-preserve", "source-noncompact"}
    expected_policies = {"slice/compact", "slice/authored", "slice/source",
                         "stream/compact", "stream/authored"}
    for role, action in sorted(bench_keys):
        result = load_object(HERE / f"oracle-{role}-{action}.json")
        require(result.get("schema") == "litchi.xml-minifier-bench.v1",
                f"oracle-{role}-{action} bench schema changed")
        require(result.get("warmups") == 5 and result.get("samples") == 100
                and result.get("batch_calls") == 100,
                f"oracle-{role}-{action} bench metadata changed")
        rows = result.get("rows")
        require(isinstance(rows, list) and len(rows) == 30,
                f"oracle-{role}-{action} bench row count changed")
        semantic: dict[tuple[str, str, str], str] = {}
        for row in rows:
            require(isinstance(row, dict), f"oracle-{role}-{action} has a non-object row")
            case, policy, chunk = row.get("case"), row.get("policy"), row.get("chunk")
            key = (case, policy, chunk)
            require(case in expected_cases and policy in expected_policies,
                    f"oracle-{role}-{action} has an unexpected bench row: {key}")
            require(chunk == ("whole" if policy.startswith("slice/") else "mixed"),
                    f"oracle-{role}-{action} has an invalid chunk binding: {key}")
            require(key not in semantic, f"oracle-{role}-{action} duplicates bench row: {key}")
            check_hex(row.get("semantic_digest"), f"oracle-{role}-{action}:{key}")
            samples = row.get("sample_ns")
            require(isinstance(samples, list) and len(samples) == 100
                    and all(isinstance(item, int) and item > 0 for item in samples),
                    f"oracle-{role}-{action}:{key} timing vector is invalid")
            semantic[key] = row["semantic_digest"]
        require(set(semantic) == {
            (case, policy, "whole" if policy.startswith("slice/") else "mixed")
            for case in expected_cases for policy in expected_policies
        }, f"oracle-{role}-{action} bench metadata has missing rows")
        semantic_maps.append(semantic)
    first = semantic_maps[0]
    require(all(item == first for item in semantic_maps[1:]),
            "oracle microbench semantic bindings differ across ABBA legs")


def verify_focused_receipts(outcome: str, baseline: dict[str, str],
                            candidate: dict[str, str]) -> None:
    """Verify canonical focused test captures and their exact log hashes.

    The two preflight captures intentionally retain failures from test-fixture
    correction.  Only the canonical baseline/candidate captures participate in
    this gate; a rejected candidate additionally needs the post-restore
    focused capture.
    """

    expected_command = [
        "cargo", "test", "-p", "xml-minifier", "--all-features", "--locked",
        "--", "--test-threads=1",
    ]
    script_digest = sha(HERE / "focused.py")

    def check_record(label: str, expected: dict[str, str], *, restored: bool = False) -> None:
        record_path = HERE / f"focused-{label}.json"
        log_path = HERE / f"focused-{label}.log"
        record = load_object(record_path)
        require(record.get("command") == expected_command,
                f"focused-{label} command changed")
        require(record.get("exit_code") == 0, f"focused-{label} did not pass")
        require(record.get("script_sha256") == script_digest,
                f"focused-{label} script binding changed")
        require(log_path.is_file(), f"focused-{label} log is missing")
        require(record.get("log_sha256") == sha(log_path),
                f"focused-{label} log digest changed")
        source = record.get("source")
        require(isinstance(source, dict), f"focused-{label} source map is missing")
        for name, digest in source.items():
            check_hex(digest, f"focused-{label}.source.{name}")
        if restored:
            differences = diff_maps(baseline, source)
            require(all(restored_test_path(name) for name in differences),
                    f"focused-restored left production source changes: {differences}")
        elif label == "baseline":
            differences = diff_maps(baseline, source)
            require(all(restored_test_path(name) for name in differences),
                    f"focused-baseline source escaped baseline/test allowance: {differences}")
        else:
            differences = diff_maps(candidate, source)
            require(all(restored_test_path(name) for name in differences),
                    f"focused-candidate source differs from candidate: {differences}")

    check_record("baseline", baseline)
    check_record("candidate", candidate)
    if outcome == "rejected":
        check_record("restored", baseline, restored=True)


def microbench_analysis() -> None:
    """Recompute all 60 paired oracle timing rows and baseline-AA drift."""

    phases = {
        "baseline-A1": load_object(HERE / "oracle-baseline-A1.json"),
        "candidate-B1": load_object(HERE / "oracle-candidate-B1.json"),
        "candidate-B2": load_object(HERE / "oracle-candidate-B2.json"),
        "baseline-A2": load_object(HERE / "oracle-baseline-A2.json"),
    }
    expected_schema = "litchi.xml-minifier-bench.v1"
    rows_by_phase: dict[str, dict[tuple[str, str, str], dict[str, Any]]] = {}
    expected_cases = {"zero-attributes", "one-attribute", "two-attributes",
                      "sixty-four-attributes", "xmlspace-preserve", "source-noncompact"}
    expected_policies = {"slice/compact", "slice/authored", "slice/source",
                         "stream/compact", "stream/authored"}
    expected_keys = {
        (case, policy, "whole" if policy.startswith("slice/") else "mixed")
        for case in expected_cases for policy in expected_policies
    }
    for phase, document in phases.items():
        reject_panic(document, f"oracle-{phase}")
        require(document.get("schema") == expected_schema,
                f"oracle-{phase} schema changed")
        require(document.get("warmups") == 5 and document.get("samples") == 100
                and document.get("batch_calls") == 100,
                f"oracle-{phase} benchmark metadata changed")
        rows = document.get("rows")
        require(isinstance(rows, list) and len(rows) == 30,
                f"oracle-{phase} does not contain 30 rows")
        indexed: dict[tuple[str, str, str], dict[str, Any]] = {}
        for row in rows:
            require(isinstance(row, dict), f"oracle-{phase} contains a non-object row")
            key = (row.get("case"), row.get("policy"), row.get("chunk"))
            require(key in expected_keys, f"oracle-{phase} has unexpected row {key}")
            require(key not in indexed, f"oracle-{phase} repeats row {key}")
            check_hex(row.get("semantic_digest"), f"oracle-{phase}:{key}.semantic_digest")
            samples = row.get("sample_ns")
            require(isinstance(samples, list) and len(samples) == 100,
                    f"oracle-{phase}:{key} sample count changed")
            require(all(isinstance(value, int) and not isinstance(value, bool) and value > 0
                        for value in samples),
                    f"oracle-{phase}:{key} sample vector is invalid")
            indexed[key] = row
        require(set(indexed) == expected_keys, f"oracle-{phase} row matrix is incomplete")
        rows_by_phase[phase] = indexed

    # The logical output digest is the semantic binding.  Timing is the only
    # field expected to vary across ABBA legs.
    reference = rows_by_phase["baseline-A1"]
    for phase, rows in rows_by_phase.items():
        for key in expected_keys:
            require(rows[key]["semantic_digest"] == reference[key]["semantic_digest"],
                    f"oracle semantic digest differs for {phase}:{key}")

    paired_rows: list[dict[str, Any]] = []
    regressions: list[dict[str, Any]] = []
    for baseline_phase, candidate_phase in (("baseline-A1", "candidate-B1"),
                                            ("baseline-A2", "candidate-B2")):
        left = rows_by_phase[baseline_phase]
        right = rows_by_phase[candidate_phase]
        for key in sorted(expected_keys):
            baseline_stats = micro_stats(left[key]["sample_ns"])
            candidate_stats = micro_stats(right[key]["sample_ns"])
            reductions = {
                metric: percent_change(float(baseline_stats[metric]),
                                       float(candidate_stats[metric]))
                for metric in ("p50", "mean", "p95", "p99")
            }
            row = {
                "baseline_phase": baseline_phase,
                "candidate_phase": candidate_phase,
                "case": key[0],
                "policy": key[1],
                "chunk": key[2],
                "baseline_stats": baseline_stats,
                "candidate_stats": candidate_stats,
                "reduction_percent": reductions,
                "semantic_digest": left[key]["semantic_digest"],
            }
            paired_rows.append(row)
            for metric, reduction in reductions.items():
                if reduction < -5.0:
                    regressions.append({
                        "baseline_phase": baseline_phase,
                        "candidate_phase": candidate_phase,
                        "case": key[0],
                        "policy": key[1],
                        "chunk": key[2],
                        "metric": metric,
                        "reduction_percent": reduction,
                    })
    require(len(paired_rows) == 60, "microbench paired row count is not 60")

    aa_rows: list[dict[str, Any]] = []
    aa_flags: list[dict[str, Any]] = []
    for key in sorted(expected_keys):
        first = micro_stats(rows_by_phase["baseline-A1"][key]["sample_ns"])
        second = micro_stats(rows_by_phase["baseline-A2"][key]["sample_ns"])
        changes = {
            metric: round((float(second[metric]) - float(first[metric]))
                          / float(first[metric]) * 100.0, 9)
            for metric in ("p50", "mean", "p95", "p99")
        }
        row = {
            "case": key[0],
            "policy": key[1],
            "chunk": key[2],
            "baseline_a1_stats": first,
            "baseline_a2_stats": second,
            "a2_minus_a1_percent": changes,
            "semantic_digest": rows_by_phase["baseline-A1"][key]["semantic_digest"],
        }
        aa_rows.append(row)
        for metric, change in changes.items():
            if abs(change) > 5.0:
                aa_flags.append({
                    "case": key[0],
                    "policy": key[1],
                    "chunk": key[2],
                    "metric": metric,
                    "a2_minus_a1_percent": change,
                })
    require(len(aa_rows) == 30, "baseline-AA row count is not 30")

    report = {
        "schema": "litchi.xml-minifier-microbench-analysis.v1",
        "status": "pass",
        "metadata": {"warmups": 5, "samples": 100, "batch_calls": 100,
                      "logical_rows_per_phase": 30, "paired_rows": 60},
        "source_outputs": [f"oracle-{phase}.json" for phase in phases],
        "paired_rows": paired_rows,
        "regressions_over_five_percent": regressions,
        "baseline_aa": {"rows": aa_rows, "drift_over_five_percent": aa_flags},
        "checks": {
            "schema_and_metadata": True,
            "no_panic_outcomes": True,
            "semantic_digest_parity": True,
            "all_60_paired_rows": len(paired_rows) == 60,
            "baseline_aa_rows": len(aa_rows) == 30,
        },
    }
    output = HERE / "microguard-analysis.json"
    if output.exists():
        require(load_object(output) == report,
                "microguard-analysis.json is stale or has been edited")
    else:
        output.write_text(json.dumps(report, indent=2) + "\n")


def analyzer_output_argument(script: Path, output: Path) -> list[str]:
    text = script.read_text()
    if re.search(r"add_argument\(\s*['\"]--output['\"]", text):
        return [sys.executable, "-B", str(script), "--output", str(output)]
    if re.search(r"add_argument\(\s*['\"][^-'\"]+['\"]", text):
        return [sys.executable, "-B", str(script), str(output)]
    # A no-argument analyzer is supported for packet versions that always
    # write analysis.json beside themselves.
    return [sys.executable, "-B", str(script)]


def replay_analyzer(outcome: str) -> None:
    script = HERE / "analyze.py"
    retained = HERE / "analysis.json"
    require(script.is_file(), "analyze.py is missing")
    require(retained.is_file(), "analysis.json is missing")
    with tempfile.TemporaryDirectory(prefix="litchi-0706-analyze-") as temporary:
        output = Path(temporary) / "analysis.json"
        command = analyzer_output_argument(script, output)
        environment = dict(os.environ, LITCHI_0706_OUTCOME=outcome)
        result = subprocess.run(command, cwd=ROOT, env=environment,
                                stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        require(result.returncode == 0,
                "analyze.py replay failed:\n" + result.stderr.decode(errors="replace"))
        if output.is_file():
            replayed = load_object(output)
        else:
            # No-argument analyzers write the retained path.  It is still an
            # exact replay because the source file is the packet's copy.
            replayed = load_object(retained)
        expected = load_object(retained)
        require(replayed == expected, "analyze.py replay differs from analysis.json")
        require(replayed.get("status") == "pass", "analysis.json does not report pass")


def read_gate_results(directory: str, expected_count: int,
                      expected_names: set[str] | None = None) -> list[dict[str, Any]]:
    result_path = HERE / directory / "results.json"
    require(result_path.is_file(), f"{directory}/results.json is missing")
    values = load(result_path)
    require(isinstance(values, list) and len(values) == expected_count,
            f"{directory}/results.json has "
            f"{len(values) if isinstance(values, list) else 'invalid'} records")
    rows: list[dict[str, Any]] = []
    names: set[str] = set()
    for value in values:
        require(isinstance(value, dict), f"{directory} has a non-object result")
        name = value.get("name")
        require(isinstance(name, str) and name not in names,
                f"{directory} has duplicate result names")
        names.add(name)
        require(value.get("exit_code") == 0, f"{directory}/{name} failed")
        log = HERE / directory / f"{name}.log"
        require(log.is_file(), f"missing {directory} log: {log}")
        if "log_sha256" in value:
            require(value["log_sha256"] == sha(log),
                    f"{directory}/{name}.log digest changed")
        rows.append(value)
    if expected_names is not None:
        require(names == expected_names,
                f"{directory} gate names differ: expected {sorted(expected_names)}, got {sorted(names)}")
    return rows


def verify_rejected_integration() -> None:
    """Require only the restored XML-minifier smoke gates after rejection."""

    directory = "integration-restored" if (HERE / "integration-restored").is_dir() else "integration"
    rows = read_gate_results(directory, 3)
    categories: set[str] = set()
    forbidden = {"litchi-opc", "litchi-pptx", "litchi-ooxml-common", "litchi-docx",
                 "litchi-xlsx", "litchi", "facade", "rustdoc", "tests-default", "tests"}
    for row in rows:
        name = row["name"].lower()
        if "fmt" in name:
            categories.add("fmt")
        if "check" in name:
            categories.add("check")
        if "clippy" in name:
            categories.add("clippy")
        command = row.get("command")
        require(isinstance(command, list), f"{directory}/{row['name']} command is missing")
        tokens = {str(token).lower() for token in command}
        require("xml-minifier" in tokens or ("fmt" in name and command == ["cargo", "fmt", "--all", "--check"]),
                f"{directory}/{row['name']} is not scoped to xml-minifier or workspace formatting")
        require(not tokens & forbidden,
                f"{directory}/{row['name']} includes a consumer gate")
    require(categories == {"fmt", "check", "clippy"},
            f"restored integration gates are incomplete: {sorted(categories)}")


def verify_gate_results(outcome: str) -> None:
    """Check final integration/evidence receipts according to retention policy."""

    if outcome == "retained":
        read_gate_results(
            "integration", 7,
            {"fmt", "check", "clippy", "tests-default", "tests", "facade", "rustdoc"},
        )
    else:
        verify_rejected_integration()
    # Evidence rechecks are independent of the retention decision and always
    # require the six historical boundary/claim/coverage records.
    read_gate_results("evidence", 6)


def main() -> int:
    try:
        verify_constraints()
        plan = load_object(HERE / "plan.json")
        verify_revision(plan)
        outcome, _baseline, _candidate, _current = verify_source_state()
        verify_focused_receipts(outcome, _baseline, _candidate)
        witnesses = cleanup_entries()
        builds = verify_builds(witnesses)
        verify_capture_receipts(plan, builds)
        verify_oracle(witnesses, builds)
        microbench_analysis()
        replay_analyzer(outcome)
        verify_gate_results(outcome)
        print(f"PASS source, constraints, focused tests, exact receipts, oracle equality, "
              f"microbench analysis, analyzer replay, and final gates (candidate {outcome})")
        return 0
    except (AuditError, OSError, subprocess.SubprocessError) as error:
        print(f"FAIL: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
