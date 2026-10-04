#!/usr/bin/env python3
"""Freeze/capture the isolated MCE-stream heaptrack diagnostic.

The allocation total covers one fresh process, including setup and teardown.
It is diagnostic evidence only; this helper makes no latency or adoption claim.
"""

from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
from pathlib import Path
import re
import stat
import subprocess
from typing import Any


P = Path(__file__).resolve().parent
SCHEMA = "litchi.xlsx-qname-direct-stream-heaptrack-0896.v1"
CAPTURE_SCHEMA = "litchi.xlsx-qname-direct-stream-heaptrack-capture-0896.v1"
CASES = ("mce_stream_count_worksheet", "mce_stream_count_document")
ROLES = ("baseline", "candidate", "candidate", "baseline")
CPU, WARMUP, SAMPLES, SUFFIX = "12", 0, 20, ".zst"
ENV = {"HOME": "/home/zhuhe", "PATH": "/usr/bin:/bin", "LANG": "C", "LC_ALL": "C", "TZ": "UTC"}
OVERLAY = {"crates/litchi-ooxml-common/src/mce/stream.rs", "crates/litchi-ooxml-common/src/mce/tests.rs"}
FIXTURES = {
    "xlsx": "test-data/ooxml/xlsx/StructuredRefs-lots-with-lookups.xlsx",
    "docx": "test-data/ooxml/docx/drawing.docx",
}
ALLOCATIONS = re.compile(r"^\s*allocations:\s*([0-9][0-9,]*)\s*$", re.MULTILINE)


def need(ok: bool, message: str) -> None:
    if not ok:
        raise ValueError(message)


def digest(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for block in iter(lambda: f.read(1024 * 1024), b""):
            h.update(block)
    return h.hexdigest()


def desc(path: Path) -> dict[str, Any]:
    path = Path(path)
    need(path.is_file() and not path.is_symlink(), f"regular file required: {path}")
    info = path.stat()
    return {"path": str(path.resolve()), "bytes": info.st_size, "sha256": digest(path), "mode": stat.S_IMODE(info.st_mode), "executable": bool(info.st_mode & stat.S_IXUSR)}


def read(path: Path) -> Any:
    with Path(path).open(encoding="utf-8") as f:
        return json.load(f)


def write(path: Path, value: Any) -> None:
    path = Path(path)
    need(not path.exists() and not path.is_symlink(), f"output already exists: {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("x", encoding="utf-8") as f:
        json.dump(value, f, indent=2, sort_keys=True)
        f.write("\n")


def same(expected: Any, actual: dict[str, Any], label: str) -> None:
    need(isinstance(expected, dict), f"{label} descriptor missing")
    for key in ("path", "bytes", "sha256", "mode", "executable"):
        need(expected.get(key) == actual.get(key), f"{label} {key} changed")


def file_path(raw: Any, base: Path, label: str) -> Path:
    need(isinstance(raw, (str, Path)), f"{label} path missing")
    path = Path(raw)
    if not path.is_absolute():
        path = base / path
    need(path.is_file() and not path.is_symlink(), f"{label} missing or symlinked: {path}")
    return path.resolve()


def versioned_tool(path: Path) -> dict[str, Any]:
    identity = desc(path)
    need(identity["executable"], f"tool is not executable: {path}")
    probe = subprocess.run([str(path), "--version"], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, check=False)
    need(probe.returncode == 0 and probe.stdout.strip(), f"cannot read tool version: {path}")
    return {"identity": identity, "path": identity["path"], "version": probe.stdout.strip()}


def check_build(path: Path, root: Path, label: str) -> dict[str, Any]:
    value = read(path)
    need(isinstance(value, dict) and value.get("status") == "pass", f"{label} did not pass")
    argv, env = value.get("argv"), value.get("environment")
    need(isinstance(argv, list) and isinstance(env, dict), f"{label} command receipt incomplete")
    need("build" in argv and "--locked" in argv and "--offline" in argv and "--bin" in argv and "--manifest-path" in argv, f"{label} is not locked standalone")
    need(argv[argv.index("--bin") + 1] == "xml_attribute_bounds", f"{label} built the wrong binary")
    manifest = Path(argv[argv.index("--manifest-path") + 1]).resolve()
    need(manifest == (root / "tools/perf-baseline/Cargo.toml").resolve(), f"{label} manifest changed")
    need(Path(value.get("cwd", "")).resolve() == root.resolve(), f"{label} cwd changed")
    for key, expected in {"CARGO_INCREMENTAL": "0", "CARGO_PROFILE_DEV_DEBUG": "0", "CARGO_PROFILE_TEST_DEBUG": "0", "CARGO_NET_OFFLINE": "true", "RUSTFLAGS": ""}.items():
        need(env.get(key) == expected, f"{label} {key} changed")
    return value


def freeze(a: argparse.Namespace) -> dict[str, Any]:
    output = Path(a.output).resolve()
    need(not output.exists() and not output.is_symlink(), f"plan already exists: {output}")
    root = Path(a.source_root).resolve()
    need(root.is_dir() and not root.is_symlink(), "source root missing or symlinked")
    source_path, candidate_source_path = Path(a.source_receipt).resolve(), Path(a.candidate_source_receipt).resolve()
    source, candidate_source = read(source_path), read(candidate_source_path)
    need(source.get("status") == "pass" and Path(source.get("source_root", "")).resolve() == root, "source receipt does not bind source root")
    need(candidate_source.get("status") == "pass" and candidate_source.get("production_adopted") is False, "candidate source receipt is not approved")
    need(candidate_source.get("guard_contract") == "post-endpoint-per-worksheet-tolerance-v1", "candidate guard contract changed")
    need(set(candidate_source.get("overlay", {})) == OVERLAY, "candidate overlay path set changed")
    need(Path(candidate_source.get("source", {}).get("path", "")).resolve() == source_path, "candidate source base changed")
    for name, item in candidate_source["overlay"].items():
        same(item, desc(Path(item["path"])), f"candidate overlay {name}")
    baseline_build_path, candidate_build_path = Path(a.baseline_build_receipt).resolve(), Path(a.candidate_build_receipt).resolve()
    check_build(baseline_build_path, root, "baseline build")
    check_build(candidate_build_path, root, "candidate build")
    binaries = {role: desc(Path(getattr(a, f"{role}_binary")).resolve()) for role in ("baseline", "candidate")}
    need(all(item["executable"] for item in binaries.values()), "diagnostic binaries must be executable")
    tools = {"heaptrack": versioned_tool(Path(a.heaptrack).resolve()), "taskset": versioned_tool(Path(a.taskset).resolve())}
    fixtures = {name: {"relative": relative, **desc(root / relative)} for name, relative in FIXTURES.items()}
    rows = [{"case": case, "ordinal": i, "pair": f"{case}-abba", "role": role, "samples": SAMPLES, "warmup": WARMUP} for i, (case, role) in enumerate((x for case in CASES for x in ((case, r) for r in ROLES)))]
    plan = {
        "schema": SCHEMA, "status": "frozen", "frozen_utc": dt.datetime.now(dt.timezone.utc).isoformat(),
        "source_root": str(root), "cpu_affinity": CPU, "cases": list(CASES), "roles": list(ROLES),
        "warmup": WARMUP, "samples": SAMPLES, "profile_suffix": SUFFIX, "environment": ENV,
        "record_argv_template": ["{taskset}", "-c", CPU, "{heaptrack}", "--record-only", "-o", "{profile_base}", "{binary}", "adversarial", "--case", "{case}", "--samples", str(SAMPLES), "--warmup", str(WARMUP), "--json", "{report}"],
        "tools": tools, "fixtures": fixtures,
        "source": {"receipt": desc(source_path), "candidate_receipt": desc(candidate_source_path), "candidate_overlay": candidate_source["overlay"], "baseline_archive": source.get("archive")},
        "build_receipts": {"baseline": desc(baseline_build_path), "candidate": desc(candidate_build_path)},
        "binaries": binaries, "helper": desc(Path(__file__)), "rows": rows, "fresh_child": True, "fail_fast": True,
        "measurement_scope": "whole-process heaptrack allocation count including setup and teardown; diagnostic only", "performance_claim": None,
    }
    write(output, plan)
    return plan


def load(plan_path: Path) -> dict[str, Any]:
    plan = read(plan_path)
    need(plan.get("schema") == SCHEMA and plan.get("status") == "frozen", "invalid direct profile plan")
    need(plan.get("cases") == list(CASES) and tuple(plan.get("roles", ())) == ROLES and plan.get("warmup") == WARMUP and plan.get("samples") == SAMPLES and plan.get("cpu_affinity") == CPU, "frozen protocol changed")
    need(plan.get("profile_suffix") == SUFFIX and plan.get("environment") == ENV and plan.get("fresh_child") is True and plan.get("fail_fast") is True, "frozen process contract changed")
    root = Path(plan["source_root"]).resolve()
    need(root.is_dir() and not root.is_symlink(), "frozen source root missing")
    same(plan["helper"], desc(Path(__file__)), "helper")
    for name, tool in plan["tools"].items():
        same(tool["identity"], desc(Path(tool["path"])), f"tool {name}")
    for role in ("baseline", "candidate"):
        same(plan["binaries"][role], desc(Path(plan["binaries"][role]["path"])), f"{role} binary")
        receipt = file_path(plan["build_receipts"][role]["path"], P, f"{role} build receipt")
        check_build(receipt, root, f"{role} build")
    for name in ("receipt", "candidate_receipt"):
        receipt = file_path(plan["source"][name]["path"], P, f"source {name}")
        same(plan["source"][name], desc(receipt), f"source {name}")
    for name, item in plan["fixtures"].items():
        same(item, desc(root / item["relative"]), f"fixture {name}")
    for name, item in plan["source"].get("candidate_overlay", {}).items():
        need(name in OVERLAY, f"unexpected candidate overlay path: {name}")
        same(item, desc(Path(item["path"])), f"overlay {name}")
    rows = plan.get("rows")
    need(isinstance(rows, list) and len(rows) == len(CASES) * len(ROLES), "row population changed")
    for i, row in enumerate(rows):
        need(row["ordinal"] == i and row["case"] == CASES[i // 4] and row["role"] == ROLES[i % 4] and row["samples"] == SAMPLES and row["warmup"] == WARMUP, "row schedule changed")
    return plan


def report(path: Path, case: str) -> dict[str, Any]:
    value = read(path)
    need(value.get("case") == case and value.get("samples") == SAMPLES and value.get("warmup") == WARMUP, "child protocol changed")
    elapsed, outcomes = value.get("elapsed_ns"), value.get("outcomes")
    need(isinstance(elapsed, list) and len(elapsed) == SAMPLES and all(isinstance(n, int) and n > 0 for n in elapsed), "child elapsed vector is not 20 positive values")
    need(isinstance(outcomes, list) and len(outcomes) == 1 and isinstance(outcomes[0], str) and outcomes[0].startswith("ok:") and outcomes[0], "child outcome oracle is not one successful value")
    need(isinstance(value.get("input_sha256"), str) and len(value["input_sha256"]) == 64 and isinstance(value.get("input_bytes"), int) and value["input_bytes"] > 0, "child input identity is missing")
    return value


def allocation_log(path: Path) -> int:
    matches = ALLOCATIONS.findall(path.read_text(encoding="utf-8"))
    need(len(matches) == 1, f"heaptrack allocation total is not unique: {path}")
    return int(matches[0].replace(",", ""))


def capture(a: argparse.Namespace) -> dict[str, Any]:
    plan_path = Path(a.plan).resolve()
    plan, output = load(plan_path), Path(a.output_root).resolve()
    need(not output.exists() and not output.is_symlink(), f"capture output already exists: {output}")
    output.mkdir(parents=True)
    plan_sha, root = digest(plan_path), Path(plan["source_root"])
    observations, rows, previous = {case: [] for case in CASES}, [], None
    for row in plan["rows"]:
        folder = output / "runs" / f"{row['ordinal']:03d}-{row['case']}-{row['role']}"
        folder.mkdir(parents=True)
        base, profile, report_path = folder / "heaptrack", folder / "heaptrack.zst", folder / "report.json"
        values = {"taskset": plan["tools"]["taskset"]["path"], "heaptrack": plan["tools"]["heaptrack"]["path"], "profile_base": str(base), "binary": plan["binaries"][row["role"]]["path"], "case": row["case"], "report": str(report_path)}
        argv = [token.format(**values) for token in plan["record_argv_template"]]
        started = dt.datetime.now(dt.timezone.utc).isoformat()
        command = {"schema": f"{CAPTURE_SCHEMA}-command", "row": row, "argv": argv, "cwd": str(root), "environment": ENV, "plan_sha256": plan_sha, "started_utc": started}
        stdout_path, stderr_path = folder / "stdout.txt", folder / "stderr.txt"
        with stdout_path.open("xb") as stdout, stderr_path.open("xb") as stderr:
            process = subprocess.run(argv, cwd=root, env=ENV.copy(), stdout=stdout, stderr=stderr, check=False)
        finished = dt.datetime.now(dt.timezone.utc).isoformat()
        command.update({"finished_utc": finished, "exit_code": process.returncode})
        command_path = folder / "command.json"
        write(command_path, command)
        need(process.returncode == 0 and profile.is_file() and not profile.is_symlink(), f"heaptrack child {row['ordinal']} failed")
        value, total = report(report_path, row["case"]), allocation_log(stderr_path)
        if previous is not None:
            need(dt.datetime.fromisoformat(started) >= previous, "children were not sequential")
        previous = dt.datetime.fromisoformat(finished)
        observation = {"case": row["case"], "role": row["role"], "ordinal": row["ordinal"], "input_bytes": value["input_bytes"], "input_sha256": value["input_sha256"], "outcomes": value["outcomes"], "heaptrack_allocations": total}
        observations[row["case"]].append(observation)
        artifacts = {name: desc(folder / name) for name in ("stdout.txt", "stderr.txt", "report.json", "command.json", "heaptrack.zst")}
        write(folder / "receipt.json", {"schema": CAPTURE_SCHEMA, "status": "pass", "row": row, "argv": argv, "cwd": str(root), "environment": ENV, "plan_sha256": plan_sha, "started_utc": started, "finished_utc": finished, "exit_code": 0, "command": desc(command_path), "artifacts": artifacts, "heaptrack_allocations": total, "workload_launched": True, "performance_claim": None})
        rows.append({"row": row, "receipt": str((folder / "receipt.json").relative_to(output))})
    for case, values in observations.items():
        expected = (values[0]["input_bytes"], values[0]["input_sha256"], values[0]["outcomes"])
        need(len(values) == 4 and all((v["input_bytes"], v["input_sha256"], v["outcomes"]) == expected for v in values), f"{case} input/outcome oracle differs across ABBA")
    result = {"schema": CAPTURE_SCHEMA, "status": "pass", "plan": desc(plan_path), "rows": rows, "children": len(rows), "cases": observations, "fresh_child": True, "fail_fast": True, "measurement_scope": plan["measurement_scope"], "performance_claim": None, "workload_launched": True}
    write(output / "capture.json", result)
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="action", required=True)
    f = sub.add_parser("freeze")
    for name in ("source_root", "source_receipt", "candidate_source_receipt", "baseline_build_receipt", "candidate_build_receipt", "baseline_binary", "candidate_binary"):
        f.add_argument(f"--{name.replace('_', '-')}", type=Path, required=True)
    f.add_argument("--heaptrack", type=Path, default=Path("/usr/bin/heaptrack"))
    f.add_argument("--taskset", type=Path, default=Path("/usr/bin/taskset"))
    f.add_argument("--output", type=Path, required=True)
    c = sub.add_parser("capture")
    c.add_argument("--plan", type=Path, required=True)
    c.add_argument("--output-root", type=Path, required=True)
    args = parser.parse_args()
    print(json.dumps(freeze(args) if args.action == "freeze" else capture(args), indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
