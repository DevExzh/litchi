#!/usr/bin/env python3
"""Capture the conditional 0708 Callgrind and process-RSS mechanism lanes.

The native pilot and both builds are owned by the coordinator.  This driver
only consumes the already frozen 0708 native binaries and runs one explicitly
requested profile or RSS phase.  It refuses to replace any artifact and binds
each child to its build record, source manifest, plans, constraints, native
ABBA result, and this script's bytes.

Phases are deliberately serialized:

    profile-baseline profile-candidate rss-baseline rss-candidate

The Callgrind lane retains all four numbered owner dumps and the termination
dump.  The RSS lane uses GNU ``time -v`` around the whole native child; its
maximum resident set size includes setup and oracle work and is never treated
as an operation-level peak.
"""

from __future__ import annotations

import argparse
import datetime as _datetime
import hashlib
import json
import os
from pathlib import Path
import re
import subprocess
import time
from typing import Any


HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
BIN_ROOT = REPO.parent / "litchi-0708-bin"

PHASES = {
    "profile-baseline": ("baseline", "profile"),
    "profile-candidate": ("candidate", "profile"),
    "rss-baseline": ("baseline", "rss"),
    "rss-candidate": ("candidate", "rss"),
}
ALIASES = {
    "profile-A": "profile-baseline",
    "profile-B": "profile-candidate",
    "rss-A": "rss-baseline",
    "rss-B": "rss-candidate",
}


def require(condition: bool, message: str) -> None:
    if not condition:
        raise RuntimeError(message)


def sha(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def digest_json(value: Any) -> str:
    encoded = (json.dumps(value, sort_keys=True, separators=(",", ":")) + "\n").encode()
    return hashlib.sha256(encoded).hexdigest()


def read_json(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing JSON input: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as error:
        raise RuntimeError(f"invalid JSON input {path}: {error}") from error


def write_json(path: Path, value: Any) -> None:
    path.write_text(json.dumps(value, indent=2) + "\n", encoding="utf-8")


def utc_now() -> str:
    return _datetime.datetime.now(_datetime.timezone.utc).isoformat()


def source_census() -> dict[str, str]:
    paths: list[Path] = [REPO / "Cargo.toml", REPO / "Cargo.lock"]
    paths.extend(path for path in (REPO / ".cargo").rglob("*") if path.is_file())
    for folder in ("crates", "tools/perf-baseline"):
        paths.extend(
            path for path in (REPO / folder).rglob("*")
            if path.is_file()
            and "target" not in path.parts
            and (path.suffix == ".rs" or path.name in {"Cargo.toml", "Cargo.lock"})
        )
    return {
        str(path.relative_to(REPO)): sha(path)
        for path in sorted(set(paths))
    }


def plans() -> tuple[dict[str, Any], dict[str, Any]]:
    main = read_json(HERE / "plan.json")
    mechanism = read_json(HERE / "mechanism-plan.json")
    require(isinstance(main, dict) and isinstance(mechanism, dict), "plans are not objects")
    require(main.get("revision") == mechanism.get("revision"), "plan revisions differ")
    require(isinstance(main.get("primary"), dict), "primary native plan is missing")
    require(main["primary"].get("case") == mechanism.get("case"), "plan cases differ")
    require(main["primary"].get("shapes") == mechanism.get("shapes"), "plan shapes differ")
    primary = main.get("primary")
    require(
        isinstance(primary, dict)
        and primary.get("repeats") == 2
        and primary.get("samples") == 200
        and primary.get("warmup") == 20,
        "0708 primary native plan differs",
    )
    profile = mechanism.get("profile")
    rss = mechanism.get("rss")
    require(
        isinstance(profile, dict)
        and profile.get("repeats") == 2
        and profile.get("samples") == 1
        and profile.get("warmup") == 0,
        "mechanism profile matrix differs",
    )
    require(
        isinstance(rss, dict)
        and rss.get("repeats") == 2
        and rss.get("samples") == 3
        and rss.get("warmup") == 2,
        "mechanism RSS matrix differs",
    )
    require(mechanism.get("cpu") == 12, "mechanism CPU differs")
    return main, mechanism


def constraints_check() -> None:
    constraints = read_json(HERE / "constraints.json")
    require(isinstance(constraints, dict), "constraints are not an object")
    for name, expected in constraints.items():
        path = REPO / str(name)
        require(path.is_file() and sha(path) == expected,
                f"constraint changed during mechanism capture: {name}")


def build_for(stage: str) -> tuple[dict[str, Any], dict[str, str], Path]:
    records_path = HERE / f"build-{stage}.json"
    source_path = HERE / f"source-{stage}.json"
    records = read_json(records_path)
    expected = read_json(source_path)
    require(isinstance(records, list) and isinstance(expected, dict),
            f"invalid {stage} build/source records")
    matches = [
        item for item in records
        if isinstance(item, dict) and Path(str(item.get("binary", ""))).name == f"{stage}-native"
    ]
    require(len(matches) == 1, f"expected one {stage}-native build record")
    record = matches[0]
    require(record.get("exit_code") == 0, f"{stage}-native build failed")
    require(record.get("source_manifest_sha256") == sha(source_path),
            f"{stage}-native source manifest binding differs")
    binary = Path(str(record.get("binary", ""))).resolve()
    expected_sha = record.get("binary_sha256")
    require(isinstance(expected_sha, str) and len(expected_sha) == 64,
            f"{stage}-native binary hash is missing")
    require(binary == (BIN_ROOT / f"{stage}-native").resolve(),
            f"{stage}-native binary path differs from frozen binary root")
    require(binary.is_file() and not binary.is_symlink() and sha(binary) == expected_sha,
            f"{stage}-native frozen binary is unavailable or changed")
    return {
        **record,
        "record_path": records_path,
        "record_sha256": sha(records_path),
        "source_path": source_path,
        "source_sha256": sha(source_path),
        "expected_source": expected,
    }, expected, binary


def source_relation(stage: str, expected: dict[str, str], current: dict[str, str],
                    main: dict[str, Any], label: str) -> dict[str, Any]:
    changed = sorted(
        name for name in set(expected) | set(current)
        if expected.get(name) != current.get(name)
    )
    allowed_roots = tuple(main.get("candidate_roots", ()))
    if stage == "candidate":
        require(not changed, f"{label}: candidate source census differs: {changed}")
        mode = "exact"
    else:
        candidate_path = HERE / "source-candidate.json"
        if candidate_path.is_file():
            candidate = read_json(candidate_path)
            require(isinstance(candidate, dict), "source-candidate.json is not a source map")
            require(current == expected or current == candidate,
                    f"{label}: baseline binary ran under an unbound source state")
        require(all(any(name.startswith(root) for root in allowed_roots) for name in changed),
                f"{label}: baseline source changed outside candidate roots: {changed}")
        mode = "exact" if not changed else "baseline-retained-under-allowed-candidate-delta"
    return {
        "mode": mode,
        "changed_paths": changed,
        "allowed_roots": list(allowed_roots),
        "expected_entry_count": len(expected),
        "current_entry_count": len(current),
    }


def native_reference(stage: str, repeat: int, shape: str, build: dict[str, Any],
                     mechanism: dict[str, Any]) -> dict[str, Any]:
    phase = mechanism["native_reference"]["phase_by_stage"][stage]
    name = f"native-{phase}-r{repeat}-{shape}"
    result_path = HERE / f"{name}.json"
    receipt_path = HERE / f"{name}.receipt.json"
    require(result_path.is_file() and not result_path.is_symlink(),
            f"missing 0708 ABBA native result: {result_path.name}")
    require(receipt_path.is_file() and not receipt_path.is_symlink(),
            f"missing 0708 ABBA native receipt: {receipt_path.name}")
    receipt = read_json(receipt_path)
    require(isinstance(receipt, dict) and receipt.get("exit_code") == 0,
            f"native reference failed: {receipt_path.name}")
    require(receipt.get("binary_sha256") == build["binary_sha256"],
            f"native reference binary differs: {name}")
    require(receipt.get("shape") == shape and receipt.get("role") == stage
            and receipt.get("phase") == phase and receipt.get("kind") == "primary",
            f"native reference lane differs: {name}")
    return {
        "name": name,
        "result": result_path.name,
        "result_sha256": sha(result_path),
        "receipt": receipt_path.name,
        "receipt_sha256": sha(receipt_path),
    }


def child_name(kind: str, stage: str, repeat: int, shape: str) -> str:
    return f"mechanism-{kind}-{stage}-r{repeat}-{shape}"


def child_artifacts(name: str, kind: str) -> list[Path]:
    fixed = [
        HERE / f"{name}.json",
        HERE / f"{name}.stdout",
        HERE / f"{name}.stderr",
        HERE / f"{name}.source-before.json",
        HERE / f"{name}.source-after.json",
    ]
    if kind == "profile":
        fixed += [HERE / f"{name}.callgrind"]
        fixed += sorted(HERE.glob(f"{name}.callgrind.*"))
    else:
        fixed += [HERE / f"{name}.time.txt"]
    return fixed


def brk_segment_overflow(paths: list[Path]) -> dict[str, Any]:
    matches: list[str] = []
    pattern = re.compile(r"brk.{0,96}overflow|overflow.{0,96}brk", re.IGNORECASE)
    for path in paths:
        if not path.is_file():
            continue
        for line in path.read_text(encoding="utf-8", errors="replace").splitlines():
            if pattern.search(line):
                matches.append(f"{path.name}: {line.strip()}")
    return {
        "present": bool(matches),
        "evidence": matches,
        "scope": "explicit text in child stdout, stderr, GNU time output, or Callgrind output",
        "interpretation": "observation only; no numerical allocation or call-count claim",
    }


def profile_command(binary: Path, output: Path, callgrind: Path, shape: str,
                    profile: dict[str, Any], cpu: int) -> list[str]:
    command = ["taskset", "-c", str(cpu), "valgrind"]
    command.extend(profile["options"])
    command += [f"--callgrind-out-file={callgrind}", str(binary)]
    command += [
        "--warmup", str(profile["warmup"]),
        "--samples", str(profile["samples"]),
        "--case", "xlsx_source_backed_cell_values_one_percent_edit_save",
        "--xlsx-cell-crud-shape", shape,
        "--json", str(output),
    ]
    return command


def rss_command(binary: Path, output: Path, time_output: Path, shape: str,
                rss: dict[str, Any], cpu: int) -> list[str]:
    command = [
        "taskset", "-c", str(cpu), "/usr/bin/time", "-v",
        "-o", str(time_output), str(binary),
        "--warmup", str(rss["warmup"]),
        "--samples", str(rss["samples"]),
        "--case", "xlsx_source_backed_cell_values_one_percent_edit_save",
        "--xlsx-cell-crud-shape", shape,
        "--json", str(output),
    ]
    return command


def run_child(stage: str, kind: str, repeat: int, shape: str,
              main: dict[str, Any], mechanism: dict[str, Any]) -> None:
    build, expected, binary = build_for(stage)
    name = child_name(kind, stage, repeat, shape)
    receipt_path = HERE / f"{name}.receipt.json"
    require(not receipt_path.exists() and not receipt_path.is_symlink(),
            f"refusing to replace receipt: {receipt_path.name}")
    require(not any(path.exists() or path.is_symlink() for path in child_artifacts(name, kind)),
            f"refusing to replace child artifacts: {name}")
    constraints_check()
    native = native_reference(stage, repeat, shape, build, mechanism)
    before = source_census()
    relation_before = source_relation(stage, expected, before, main, name)
    before_path = HERE / f"{name}.source-before.json"
    after_path = HERE / f"{name}.source-after.json"
    write_json(before_path, before)
    output_path = HERE / f"{name}.json"
    stdout_path = HERE / f"{name}.stdout"
    stderr_path = HERE / f"{name}.stderr"
    callgrind_path = HERE / f"{name}.callgrind"
    time_path = HERE / f"{name}.time.txt"
    if kind == "profile":
        command = profile_command(binary, output_path, callgrind_path, shape,
                                   mechanism["profile"], mechanism["cpu"])
    else:
        command = rss_command(binary, output_path, time_path, shape,
                              mechanism["rss"], mechanism["cpu"])
    started = utc_now()
    tick = time.monotonic()
    with stdout_path.open("wb") as stdout, stderr_path.open("wb") as stderr:
        result = subprocess.run(command, cwd=REPO, stdout=stdout, stderr=stderr)
    elapsed = time.monotonic() - tick
    after = source_census()
    relation_after = source_relation(stage, expected, after, main, name)
    write_json(after_path, after)
    require(before == after, f"{name}: source census changed during child")
    require(sha(binary) == build["binary_sha256"], f"{name}: frozen binary changed")
    artifacts: dict[str, str] = {}
    for path in sorted(set(child_artifacts(name, kind))):
        require(path.is_file() and not path.is_symlink(), f"{name}: missing artifact {path.name}")
        artifacts[path.name] = sha(path)
    receipt = {
        "schema_version": 1,
        "name": name,
        "kind": kind,
        "role": stage,
        "lane": kind,
        "repeat": repeat,
        "shape": shape,
        "case": mechanism["case"],
        "command": command,
        "start_utc": started,
        "end_utc": utc_now(),
        "seconds": elapsed,
        "exit_code": result.returncode,
        "cpu": mechanism["cpu"],
        "binary_path": str(binary),
        "binary_sha256": sha(binary),
        "binary_bytes": binary.stat().st_size,
        "build_record": str(build["record_path"].relative_to(HERE)),
        "build_record_sha256": build["record_sha256"],
        "build_source_manifest": str(build["source_path"].relative_to(HERE)),
        "build_source_manifest_sha256": build["source_sha256"],
        "retained_source_census_sha256": digest_json(expected),
        "retained_source_entry_count": len(expected),
        "retained_binary_source": {
            "manifest": str(build["source_path"].relative_to(HERE)),
            "manifest_sha256": build["source_sha256"],
            "source_census_sha256": digest_json(expected),
            "source_entry_count": len(expected),
        },
        "native_reference": native,
        "current_checkout_source": {
            "before_artifact": before_path.name,
            "after_artifact": after_path.name,
            "before_sha256": digest_json(before),
            "after_sha256": digest_json(after),
            "before_file_sha256": sha(before_path),
            "after_file_sha256": sha(after_path),
            "before_entry_count": len(before),
            "after_entry_count": len(after),
            "relation_before": relation_before,
            "relation_after": relation_after,
            "unchanged_during_child": before == after,
        },
        "plan_sha256": sha(HERE / "plan.json"),
        "mechanism_plan_sha256": sha(HERE / "mechanism-plan.json"),
        "script_sha256": sha(Path(__file__)),
        "constraints_sha256": sha(HERE / "constraints.json"),
        "environment": {
            key: os.environ.get(key)
            for key in ("RUSTFLAGS", "LD_PRELOAD", "MALLOC_CONF", "GLIBC_TUNABLES")
        },
        "artifacts": artifacts,
        "rss_scope": (
            "whole_child_including_setup_and_oracles; no operation-level peak claim"
            if kind == "rss" else None
        ),
        "allocation_call_interpretation": "disabled; unresolved 0707 discrepancy",
    }
    receipt["brk_segment_overflow"] = brk_segment_overflow(
        [stdout_path, stderr_path, time_path, callgrind_path,
         *sorted(HERE.glob(f"{name}.callgrind.*"))]
    )
    write_json(receipt_path, receipt)
    require(result.returncode == 0, f"{name}: child failed with exit code {result.returncode}")
    if kind == "profile":
        parts = sorted(
            path for path in HERE.glob(f"{name}.callgrind.*")
            if path.is_file() and path.suffix[1:].isdigit()
        )
        expected_parts = [f"{name}.callgrind.{part}" for part in mechanism["profile"]["numbered_parts"]]
        require([path.name for path in parts] == expected_parts,
                f"{name}: Callgrind numbered dumps differ")
    else:
        require(time_path.is_file(), f"{name}: GNU time output is missing")
    print(f"{name} passed", flush=True)


def run_phase(raw_phase: str) -> None:
    phase = ALIASES.get(raw_phase, raw_phase)
    require(phase in PHASES, f"unknown mechanism phase: {raw_phase}")
    main, mechanism = plans()
    constraints_check()
    stage, kind = PHASES[phase]
    repeats = int(mechanism[kind]["repeats"])
    for repeat in range(1, repeats + 1):
        shapes = list(mechanism["shapes"])
        if repeat == 2:
            shapes.reverse()
        for shape in shapes:
            run_child(stage, kind, repeat, shape, main, mechanism)
    print(f"phase {phase} complete ({repeats * len(mechanism['shapes'])} children)", flush=True)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("phase", help="one serialized profile or RSS phase")
    args = parser.parse_args()
    try:
        run_phase(args.phase)
    except (OSError, RuntimeError, ValueError) as error:
        print(f"mechanism capture failed: {error}", file=os.sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
