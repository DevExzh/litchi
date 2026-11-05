#!/usr/bin/env python3
"""Run the unrun suffix of the Q472a final-gate diagnostic.

This packet is deliberately not a final gate.  It regenerates the reviewed
full53 sequence and executes only full53 indices 29..52 plus runtime
supplement indices 7..11.  Compiler rows and full53 indices 0..28 are
recorded as skipped; no receipt from the failed C33 run is imported.

Preparation and dry-check perform only source, tool, plan, and filesystem
identity checks.  ``--go`` is explicit and coordinator-controlled.  Ordinary
command failures are retained and the next selected row is attempted, while
resource, provenance, quiet-window, process-leak, or OfficeCLI-barrier
failures stop the diagnostic.  A completed diagnostic is never an acceptance
claim for the eventual all-65 candidate gate.
"""
from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import os
import re
import shutil
import stat
import sys
import time
from pathlib import Path
from typing import Any


COORD = Path("/home/zhuhe/code/litchi-worktrees/.coord")
ROOT_DEFAULT = Path("/home/zhuhe/code/litchi-worktrees/uo2-final-gate-drawingml-tests")
V4_SCRIPT = COORD / "final-gate-coverage-780cf4c7/rebinder-v4/rebind_final_gate.py"
V4_SHA256 = "16f68c596c3f0b68f56ac7818a0d683eb3ff8667079831ef234b28c6d840a2e4"
APPROVED_RUNNER = Path(
    "/home/zhuhe/code/litchi-worktrees/"
    "uo2-gate-runner-self-attribution/docs/unified-ops/results/0362ul29/tooling/"
    "run_safe_gate.py"
)
APPROVED_HELPER = APPROVED_RUNNER.with_name("safe_chunked_tests.py")
APPROVED_CARGO = APPROVED_RUNNER.parent / "bin/cargo"
APPROVED_FALLBACK = COORD / "bin/cargo"
TOOL_HASHES = {
    "runner": "1904177040fc7c523582433028d9bba56efc30fe0d9d96beb554a2fab343691e",
    "helper": "3e03aa2227ac0002c3052cf5080d032cd2e9150f584c926b9aa0d499175ec67f",
    "cargo": "495e24c6f60b238378386d86ac4158cba904992ce8870a9494f2c4a0eedd256d",
    "fallback": "1b22a69fd625d57565041bf6c634d8ebd1a82be8b3d0006390472c86324230d8",
}

BASE_COMMIT = "9a3e0b17e7da92614353de342e778f2ac4a81cb4"
Q_COMMIT = "472a89a3fae1cbe81d177b6c2f139e03d788bf43"
Q_TREE = "aeaeb94378f4bccd36f1092b607d66a33201aa1d"
Q_PARENT = "33a0a63d3553712c1276da592f7846a0c7dcabd6"
Q_BRANCH = "uo2/final-gate-drawingml-tests"
Q_CHANGED_RUST_COUNT = 590
Q_CHANGED_RUST_PATHLIST_SHA256 = "ffe3238a3d460597f4bb1f95a727011577c63bb7197f3b616461c80b5cd2e405"
EXPECTED_LOCKS = {
    "Cargo.lock": (139207, "02fda2543226f28c2ded9bbf7e18618f8b0aab1c4d02c7b2c4fdd73511e4ff7a"),
    "tools/perf-baseline/Cargo.lock": (56941, "de1f8ae53b3455a4246da0b9340896881cd72ca0295b8e52444e0d55e6c9c4ea"),
    "tools/native-resave/Cargo.lock": (56689, "989da70f10deab4460a77e4afa3a1b840b4697bbea2c07e134e22d9b4c5084ae"),
}

TARGET_START_BYTES = 18_000_000_000
TARGET_LIVE_BYTES = 19_000_000_000
TARGET_HARD_BYTES = 20_000_000_000
MINIMUM_FREE_BYTES = 12_000_000_000
FULL53_SUFFIX = tuple(range(29, 53))
RUNTIME_SUFFIX = tuple(range(7, 12))


class DiagnosticError(RuntimeError):
    """A fail-closed packet, source, or host-guard error."""


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def canonical(value: Any) -> bytes:
    return json.dumps(value, sort_keys=True, separators=(",", ":"), ensure_ascii=False).encode()


def lexical(path: Path) -> Path:
    return Path(os.path.abspath(os.fspath(path)))


def reject_symlink_components(path: Path) -> Path:
    path = lexical(path)
    current = Path(path.anchor)
    for component in path.parts[1:]:
        current /= component
        if current.is_symlink():
            raise DiagnosticError(f"symlink path component is forbidden: {current}")
    return path


def reject_overlap(label_a: str, path_a: Path, label_b: str, path_b: Path) -> None:
    a = lexical(path_a)
    b = lexical(path_b)
    if a == b or a.is_relative_to(b) or b.is_relative_to(a):
        raise DiagnosticError(f"{label_a} and {label_b} overlap: {a} / {b}")


def check_layout(root: Path, packet: Path, state: Path, target: Path) -> None:
    root = reject_symlink_components(root)
    packet = reject_symlink_components(packet)
    state = reject_symlink_components(state)
    target = reject_symlink_components(target)
    reject_overlap("source", root, "packet", packet)
    reject_overlap("source", root, "state", state)
    reject_overlap("source", root, "target", target)
    reject_overlap("packet", packet, "state", state)
    reject_overlap("packet", packet, "target", target)
    reject_overlap("state", state, "target", target)


def require_absent(path: Path) -> None:
    reject_symlink_components(path)
    if path.exists() or path.is_symlink():
        raise DiagnosticError(f"fresh path already exists: {path}")


def require_empty_directory(path: Path, label: str) -> None:
    reject_symlink_components(path)
    if not path.is_dir() or path.is_symlink() or any(path.iterdir()):
        raise DiagnosticError(f"{label} must be a fresh empty directory: {path}")


def load_v4() -> Any:
    if not V4_SCRIPT.is_file() or sha256_file(V4_SCRIPT) != V4_SHA256:
        raise DiagnosticError("reviewed v4 rebinder is missing or changed")
    spec = importlib.util.spec_from_file_location("q472a_reviewed_v4", V4_SCRIPT)
    if spec is None or spec.loader is None:
        raise DiagnosticError("unable to load reviewed v4 rebinder")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def expected_source() -> dict[str, Any]:
    return {
        "commit": Q_COMMIT,
        "tree": Q_TREE,
        "parents": [Q_PARENT],
        "branch": Q_BRANCH,
        "changed_rust_count": Q_CHANGED_RUST_COUNT,
        "changed_rust_pathlist_sha256": Q_CHANGED_RUST_PATHLIST_SHA256,
    }


def validate_source(v4: Any, root: Path) -> dict[str, Any]:
    source = v4.validate_source_guards(root, BASE_COMMIT)
    expected = expected_source()
    for key, value in expected.items():
        if source.get(key) != value:
            raise DiagnosticError(f"Q source {key} mismatch: {source.get(key)!r}")
    if source.get("locks") != {
        relative: {"bytes": size, "sha256": digest}
        for relative, (size, digest) in EXPECTED_LOCKS.items()
    }:
        raise DiagnosticError("Q Cargo.lock identity mismatch")
    return source


def runner_source(v4: Any, root: Path) -> dict[str, Any]:
    value = v4.load_runner().source_state(root)
    expected = {
        "commit": Q_COMMIT, "branch": Q_BRANCH, "status": "",
        "cargo_locks": {path: digest for path, (_, digest) in EXPECTED_LOCKS.items()},
    }
    if value != expected:
        raise DiagnosticError("runner source identity changed")
    return value


def validate_tools(v4: Any) -> dict[str, Any]:
    tools = v4.validate_tools()
    if {name: receipt["sha256"] for name, receipt in tools.items()} != TOOL_HASHES:
        raise DiagnosticError("approved runner/tool identity changed")
    return tools


def projection(step: dict[str, Any]) -> dict[str, Any]:
    return {key: step.get(key) for key in ("name", "kind", "argv", "command", "env", "isolated", "pass_fds")}


def projected_receipt(phase: str, index: int, step: dict[str, Any]) -> dict[str, Any]:
    item = {"phase": phase, "index": index, **projection(step)}
    item["sha256"] = hashlib.sha256(canonical(projection(step))).hexdigest()
    return item


def selected_steps(v4: Any, root: Path, state: Path, target: Path, source: dict[str, Any]) -> tuple[list[dict[str, Any]], list[dict[str, Any]], dict[str, Any]]:
    full = v4.full53_steps(root, state, target)
    supplements = v4.supplement_steps(state / "runtime", target, source["changed_rust_paths"])
    if len(full) != 53 or len(supplements) != 12:
        raise DiagnosticError(f"reviewed step counts changed: full53={len(full)} supplements={len(supplements)}")
    if tuple(range(len(full))) != tuple(range(53)):
        raise DiagnosticError("full53 index space changed")
    row37 = full[37]
    row37_argv = row37.get("argv", [])
    if row37.get("name") != "cli-tests" or "-p" not in row37_argv or "litchi-cli" not in row37_argv:
        raise DiagnosticError(f"row37 OfficeCLI barrier moved: {row37.get('name')!r} / {row37_argv!r}")
    v4.require_unique_steps(full + supplements[:7] + supplements[7:])
    selected_full = [{"phase": "full53", "index": index, "step": full[index]} for index in FULL53_SUFFIX]
    selected_runtime = [{"phase": "runtime-supplements", "index": index, "step": supplements[index]} for index in RUNTIME_SUFFIX]
    skipped = {
        "compiler_supplements": [
            {"phase": "compiler-supplements", "index": index, "name": supplements[index]["name"], "sha256": hashlib.sha256(canonical(projection(supplements[index]))).hexdigest()}
            for index in range(7)
        ],
        "full53": [
            {"phase": "full53", "index": index, "name": full[index]["name"], "sha256": hashlib.sha256(canonical(projection(full[index]))).hexdigest()}
            for index in range(29)
        ],
        "runtime_supplements": [
            {"phase": "runtime-supplements", "index": index, "name": supplements[index]["name"], "sha256": hashlib.sha256(canonical(projection(supplements[index]))).hexdigest()}
            for index in range(7)
        ],
    }
    return selected_full, selected_runtime, {"full53": full, "supplements": supplements, "skipped": skipped}


def officecli_barrier(state: Path, target: Path) -> dict[str, Any]:
    return {
        "full53_index": 37,
        "before_full53_index": 38,
        "source": str(target / "tmp/officecli-corpus-outcomes.tsv"),
        "destination": str(state / "officecli-row37/officecli-corpus-outcomes.tsv"),
        "required": True,
        "requires_row_status": "passed",
        "verification": ["bytes", "sha256", "mode"],
    }


def plan_digest(packet: Path) -> str:
    return sha256_file(packet / "plan.json")


def write_fresh(path: Path, value: Any) -> None:
    if path.exists():
        raise DiagnosticError(f"refusing to overwrite {path}")
    path.write_text(json.dumps(value, indent=2) + "\n", encoding="utf-8")


def build_plan(root: Path, packet: Path, state: Path, target: Path) -> dict[str, Any]:
    if os.environ.get("LITCHI_BLESS"):
        raise DiagnosticError("LITCHI_BLESS must be unset")
    check_layout(root, packet, state, target)
    require_absent(packet)
    require_absent(state)
    require_absent(target)
    v4 = load_v4()
    tools = validate_tools(v4)
    source = validate_source(v4, root)
    # All identity checks precede directory creation.
    packet.parent.mkdir(parents=True, exist_ok=True)
    packet.mkdir()
    state.mkdir()
    target.mkdir(parents=True)
    selected_full, selected_runtime, all_steps = selected_steps(v4, root, state, target, source)
    full = all_steps["full53"]
    supplements = all_steps["supplements"]
    plan = {
        "schema": "0362ul-q472a-rest29-diagnostic-v2",
        "status": "prelaunch-ready-inert",
        "scope": "diagnostic-only",
        "acceptance": "never; this packet cannot qualify the final all-65 candidate gate",
        "no_commands_run": True,
        "source": source,
        "runner_source": runner_source(v4, root),
        "source_binding": expected_source(),
        "base_commit": BASE_COMMIT,
        "tools": tools,
        "driver": {"path": str(Path(__file__).resolve()), "sha256": sha256_file(Path(__file__).resolve()), "v4_sha256": V4_SHA256},
        "approved_runner": {"path": str(APPROVED_RUNNER), "sha256": TOOL_HASHES["runner"]},
        "state": str(state),
        "target": str(target),
        "policy": {
            "offline": True,
            "locked": True,
            "jobs": 1,
            "test_threads": 1,
            "memory_max": "20G",
            "target_start_bytes": TARGET_START_BYTES,
            "target_live_bytes": TARGET_LIVE_BYTES,
            "target_hard_bytes": TARGET_HARD_BYTES,
            "minimum_free_bytes": MINIMUM_FREE_BYTES,
            "quiet_before_every_step": True,
            "continue_on": "ordinary command/test/compile failure only",
            "stop_on": ["resource", "provenance", "quiet-window", "leftover-process", "OfficeCLI-barrier"],
        },
        "phase_order": ["full53-suffix", "runtime-supplements-suffix"],
        "skipped_prefix": {
            "compiler_supplements": 7,
            "full53_indices": list(range(29)),
            "runtime_supplement_indices": list(range(7)),
            "reason": "already covered by neither this packet nor reusable acceptance evidence; regenerated only for identity and coverage receipts",
        },
        "coverage": {
            "original_full53_count": 53,
            "selected_full53_indices": list(FULL53_SUFFIX),
            "selected_full53_count": len(selected_full),
            "original_runtime_supplement_count": 12,
            "selected_runtime_supplement_indices": list(RUNTIME_SUFFIX),
            "selected_runtime_supplement_count": len(selected_runtime),
            "selected_top_level_count": len(selected_full) + len(selected_runtime),
            "all65_equivalent_count": 65,
            "ordinary_failures_do_not_skip_later_selected_rows": True,
            "step_receipts": [
                *[projected_receipt("full53", index, full[index]) for index in FULL53_SUFFIX],
                *[projected_receipt("runtime-supplements", index, supplements[index]) for index in RUNTIME_SUFFIX],
            ],
        },
        "full53_suffix": [{"index": item["index"], "name": item["step"]["name"], "step": item["step"]} for item in selected_full],
        "runtime_suffix": [{"index": item["index"], "name": item["step"]["name"], "step": item["step"]} for item in selected_runtime],
        "skipped_receipts": all_steps["skipped"],
        "officecli_barrier": officecli_barrier(state, target),
        "result_policy": {
            "code": 0,
            "ordinary_failure": 1,
            "guard_stop": 2,
            "note": "zero means all selected rows passed; one means all selected rows completed with ordinary failures; two means fail-closed guard stop",
        },
    }
    (packet / "changed-rust-paths.txt").write_text("\n".join(source["changed_rust_paths"]) + "\n", encoding="utf-8")
    write_fresh(packet / "plan.json", plan)
    return plan


def read_plan(packet: Path) -> dict[str, Any]:
    path = packet / "plan.json"
    if not path.is_file():
        raise DiagnosticError(f"plan missing: {path}")
    plan = json.loads(path.read_text(encoding="utf-8"))
    if plan.get("schema") != "0362ul-q472a-rest29-diagnostic-v2" or plan.get("status") != "prelaunch-ready-inert":
        raise DiagnosticError("invalid or mutable diagnostic plan")
    return plan


def verify_plan(root: Path, packet: Path, plan: dict[str, Any], *, go: bool, expected_plan_sha256: str | None) -> tuple[Any, Any, dict[str, Any]]:
    if expected_plan_sha256 is None or plan_digest(packet) != expected_plan_sha256:
        raise DiagnosticError("plan SHA-256 mismatch")
    if os.environ.get("LITCHI_BLESS"):
        raise DiagnosticError("LITCHI_BLESS must be unset")
    state = lexical(Path(plan["state"]))
    target = lexical(Path(plan["target"]))
    check_layout(root, packet, state, target)
    if lexical(Path(plan["driver"]["path"])) != lexical(Path(__file__)) or plan["driver"]["sha256"] != sha256_file(Path(__file__).resolve()):
        raise DiagnosticError("driver identity changed")
    v4 = load_v4()
    tools = validate_tools(v4)
    if tools != plan.get("tools"):
        raise DiagnosticError("tool identity changed")
    source = validate_source(v4, root)
    if source != plan.get("source"):
        raise DiagnosticError("source identity changed")
    current_runner_source = runner_source(v4, root)
    require_identity("runner source", plan.get("runner_source"), current_runner_source)
    selected_full, selected_runtime, all_steps = selected_steps(v4, root, state, target, source)
    expected = [
        projected_receipt("full53", index, all_steps["full53"][index]) for index in FULL53_SUFFIX
    ] + [
        projected_receipt("runtime-supplements", index, all_steps["supplements"][index]) for index in RUNTIME_SUFFIX
    ]
    if plan["coverage"]["step_receipts"] != expected:
        raise DiagnosticError("selected command sequence changed")
    if plan["officecli_barrier"] != officecli_barrier(state, target):
        raise DiagnosticError("OfficeCLI barrier changed")
    if go:
        require_empty_directory(state, "state")
        require_empty_directory(target, "target")
        runner = v4.load_runner()
        if runner.builders({os.getpid(), os.getppid()}, root, target):
            raise DiagnosticError("quiet-window builder guard failed before GO")
        if shutil.disk_usage(target.parent).free < MINIMUM_FREE_BYTES:
            raise DiagnosticError("free disk below 12 GiB guard before GO")
    return v4.load_runner(), v4, {"source": source, "runner_source": current_runner_source, "selected_full": selected_full, "selected_runtime": selected_runtime, "all_steps": all_steps}


def classify_row(row: dict[str, Any]) -> str:
    """Return the only failure classes that may affect continuation."""
    if row.get("resource_failure") or row.get("target_live_guard_errors"):
        return "resource"
    if row.get("provenance_failure"):
        return "provenance"
    if row.get("leftover_processes") or row.get("owned_processes"):
        return "leftover-process"
    if row.get("barrier_failure"):
        return "preservation"
    if any(sample.get("owned") or sample.get("builders") for sample in row.get("post_exit", {}).get("samples", [])):
        return "leftover-process"
    status, code = row.get("status"), row.get("effective_exit")
    if status == "passed" and code == 0:
        return "pass"
    if status != "failed" or not isinstance(code, int):
        return "invalid-row"
    if code < 0 or code >= 128:
        return "resource"
    if row.get("log"):
        pattern = re.compile(rb"signal: 9|SIGKILL|out of memory|memory allocation of [0-9]+ bytes failed|oom-kill", re.I)
        with Path(row["log"]).open("rb") as stream:
            tail = b""
            for chunk in iter(lambda: stream.read(65536), b""):
                combined = tail + chunk
                if pattern.search(combined):
                    return "resource"
                tail = combined[-128:]
    return "code" if code else "invalid-row"


def continuation_action(row: dict[str, Any]) -> str:
    return "continue" if classify_row(row) in {"pass", "code"} else "stop"


def classify_exception(error: BaseException) -> str:
    text = repr(error).lower()
    if any(token in text for token in ("quiet", "builder", "free disk", "target", "resource", "process")):
        return "resource"
    if any(token in text for token in ("source", "provenance", "lock", "tool identity", "plan sha", "driver identity")):
        return "provenance"
    return "blocked"


def require_identity(label: str, expected: Any, actual: Any) -> None:
    if expected != actual:
        raise DiagnosticError(f"{label} drift: expected {expected!r}, got {actual!r}")


def assert_identity(root: Path, packet: Path, plan: dict[str, Any], v4: Any, expected_plan_sha256: str) -> dict[str, Any]:
    require_identity("plan", expected_plan_sha256, plan_digest(packet))
    require_identity("driver", plan["driver"]["sha256"], sha256_file(Path(__file__)))
    require_identity("runner source", plan["runner_source"], runner_source(v4, root))
    current = validate_source(v4, root)
    require_identity("source", plan["source"], current)
    require_identity("tools", plan["tools"], validate_tools(v4))
    return current


def preserve_officecli(v4: Any, target: Path, state: Path, row: dict[str, Any]) -> dict[str, Any]:
    if row.get("status") != "passed" or row.get("effective_exit") != 0:
        raise DiagnosticError("OfficeCLI row37 did not pass; barrier cannot open")
    try:
        return v4.preserve_officecli(target, state)
    except Exception as error:
        raise DiagnosticError(f"OfficeCLI preservation failed: {error}") from error


def execute(root: Path, packet: Path, plan: dict[str, Any], expected_plan_sha256: str) -> int:
    runner, v4, details = verify_plan(root, packet, plan, go=True, expected_plan_sha256=expected_plan_sha256)
    state = Path(plan["state"])
    target = Path(plan["target"])
    runtime_state = state / "runtime"
    runtime_state.mkdir()
    manifest = {
        "schema": "0362ul-q472a-rest29-diagnostic-execution-v2",
        "status": "running",
        "scope": "diagnostic-only",
        "started": time.time(),
        "plan_sha256": expected_plan_sha256,
        "driver_sha256": sha256_file(Path(__file__).resolve()),
        "source": details["source"],
        "runner_source": details["runner_source"],
        "state": str(state),
        "target": str(target),
        "selected_full53_indices": list(FULL53_SUFFIX),
        "selected_runtime_indices": list(RUNTIME_SUFFIX),
        "old_c33_receipts_reused": False,
    }
    write_fresh(state / "execution-manifest.json", manifest)
    rows: list[dict[str, Any]] = []
    hard_stop: dict[str, Any] | None = None
    ordinary_failures = 0
    sequence = [("full53", index, details["all_steps"]["full53"][index], state, 53) for index in FULL53_SUFFIX]
    sequence += [("runtime-supplements", index, details["all_steps"]["supplements"][index], runtime_state, 12) for index in RUNTIME_SUFFIX]
    for phase, index, step, phase_state, total in sequence:
        try:
            assert_identity(root, packet, plan, v4, expected_plan_sha256)
            if runner.builders({os.getpid(), os.getppid()}, root, target):
                raise DiagnosticError(f"quiet-window builder guard failed before {phase}:{index}")
            row = runner.run_step(step, root, target, phase_state, details["runner_source"], index, total, attempt=f"q472a-rest29-{phase}-{index:02d}")
            row["phase"] = phase
            row["original_index"] = index
            kind = classify_row(row)
            row["continuation_class"] = kind
            rows.append(row)
            write_fresh_or_replace(state / "rows.json", rows)
            if phase == "full53" and index == 37:
                try:
                    row["officecli_preservation"] = preserve_officecli(v4, target, state, row)
                except DiagnosticError as error:
                    row["barrier_failure"] = str(error)
                    row["continuation_class"] = "preservation"
                    write_fresh_or_replace(state / "rows.json", rows)
                    hard_stop = {"phase": phase, "index": index, "reason": str(error)}
                    break
                write_fresh_or_replace(state / "rows.json", rows)
            assert_identity(root, packet, plan, v4, expected_plan_sha256)
            if runner.builders({os.getpid(), os.getppid()}, root, target):
                raise DiagnosticError(f"quiet-window builder guard failed after {phase}:{index}")
            if continuation_action(row) == "stop":
                hard_stop = {"phase": phase, "index": index, "reason": kind, "row": row}
                break
            if kind == "code":
                ordinary_failures += 1
        except Exception as error:
            kind = classify_exception(error)
            blocked = {
                "phase": phase,
                "original_index": index,
                "name": step["name"],
                "status": "blocked",
                "continuation_class": kind,
                "error": repr(error),
            }
            rows.append(blocked)
            write_fresh_or_replace(state / "rows.json", rows)
            hard_stop = blocked
            break
    status = "diagnostic-complete" if hard_stop is None else "diagnostic-stopped"
    result = {
        "schema": "0362ul-q472a-rest29-diagnostic-result-v2",
        "status": status,
        "scope": "diagnostic-only",
        "acceptance": False,
        "ordinary_failure_count": ordinary_failures,
        "attempted_count": len(rows),
        "planned_count": 29,
        "hard_stop": hard_stop,
        "rows": rows,
        "selected_full53_indices": list(FULL53_SUFFIX),
        "selected_runtime_indices": list(RUNTIME_SUFFIX),
        "skipped_prefix": plan["skipped_prefix"],
        "source": plan["source"],
        "plan_sha256": expected_plan_sha256,
    }
    write_fresh_or_replace(state / "result.json", result)
    if hard_stop is not None:
        return 2
    return 1 if ordinary_failures else 0


def write_fresh_or_replace(path: Path, value: Any) -> None:
    # rows/result are intentionally updated during one run; packet/plan files
    # remain immutable.  This helper makes the mutable state boundary explicit.
    path.write_text(json.dumps(value, indent=2) + "\n", encoding="utf-8")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--root", type=Path, default=ROOT_DEFAULT)
    parser.add_argument("--packet", type=Path, required=True)
    parser.add_argument("--state", type=Path)
    parser.add_argument("--target", type=Path)
    parser.add_argument("--prepare", action="store_true")
    parser.add_argument("--dry-check", action="store_true")
    parser.add_argument("--go", action="store_true")
    parser.add_argument("--expected-plan-sha256")
    args = parser.parse_args(argv)
    if sum(bool(flag) for flag in (args.prepare, args.dry_check, args.go)) != 1:
        parser.error("choose exactly one of --prepare, --dry-check, --go")
    try:
        root = lexical(args.root)
        packet = lexical(args.packet)
        if args.prepare:
            if not args.state or not args.target:
                parser.error("--prepare requires --state and --target")
            plan = build_plan(root, packet, lexical(args.state), lexical(args.target))
            print(json.dumps({"status": plan["status"], "scope": plan["scope"], "source": Q_COMMIT, "selected_steps": 29, "plan_sha256": plan_digest(packet)}, sort_keys=True))
            return 0
        plan = read_plan(packet)
        expected = args.expected_plan_sha256
        if not expected:
            parser.error("--dry-check/--go requires --expected-plan-sha256")
        if args.state and lexical(args.state) != lexical(Path(plan["state"])):
            raise DiagnosticError("--state does not match plan")
        if args.target and lexical(args.target) != lexical(Path(plan["target"])):
            raise DiagnosticError("--target does not match plan")
        if args.dry_check:
            verify_plan(root, packet, plan, go=False, expected_plan_sha256=expected)
            require_empty_directory(Path(plan["state"]), "state")
            require_empty_directory(Path(plan["target"]), "target")
            print(json.dumps({"status": "dry-check-pass", "scope": "diagnostic-only", "source": Q_COMMIT, "selected_steps": 29}, sort_keys=True))
            return 0
        return execute(root, packet, plan, expected)
    except DiagnosticError as error:
        print(f"q472a-rest29: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
