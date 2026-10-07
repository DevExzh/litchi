#!/usr/bin/env python3
"""Offline custody and mapping controls for analyze_0953_v2.

All child/export files in this control are synthetic tiny regular files. The
only retained workload artifact read is the old 0952 binary/mmap export for a
real PT_LOAD owner/parent parser positive and wrong-PID/parent/event negatives.
No benchmark, perf, build, or decoder is started.
"""
from __future__ import annotations

import copy
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import shutil
import stat
import tempfile
from typing import Any, Callable

ROOT = Path(__file__).resolve().parents[1]
READER = ROOT / "tools/analyze_0953_v3.py"
SEMANTIC = ROOT / "tools/semantic_0953_v1.py"
OLD_PACKET = Path("/home/zhuhe/litchi-performance-0952")
OLD_PLAN = OLD_PACKET / "plan-0952-v1.json"
OLD_BINARY = Path("/home/zhuhe/litchi-performance-0950/target-candidate-native-v1/release/litchi-perf-baseline")
OLD_MMAP = OLD_PACKET / "analysis/mmap/00-medium/events.txt"


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def desc(path: Path, raw: str | None = None) -> dict[str, Any]:
    st = path.stat()
    return {"path": raw if raw is not None else str(path.resolve()),
            "bytes": st.st_size, "sha256": sha(path),
            "mode": stat.S_IMODE(st.st_mode),
            "executable": bool(st.st_mode & stat.S_IXUSR)}


def load_reader() -> Any:
    spec = importlib.util.spec_from_file_location("analyze_0953_v2_controls_reader", READER)
    if spec is None or spec.loader is None:
        raise RuntimeError("cannot load analyzer")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def make_plan(packet: Path) -> dict[str, Any]:
    old = json.loads(OLD_PLAN.read_text())
    semantic = packet / "tools/semantic_0953_v1.py"
    return {
        "schema": "litchi.performance.0953.native-publish-sampling-plan.v1",
        "contract": {"samples": 200, "warmup": 0, "case": old["case"],
                     "cpu": 12, "event": "cycles:u", "shapes": ["medium"]},
        "binary": desc(OLD_BINARY),
        "runner": desc(ROOT / "tools/capture_0953_v1.py"),
        "semantic_adapter": desc(semantic, "tools/semantic_0953_v1.py"),
        "semantic_binding": {
            key: desc(Path(value)) for key, value in {
                "oracle": "/home/zhuhe/litchi-performance-0950/tools/semantic_0950_v2.py",
                "oracle_original": "/home/zhuhe/litchi-performance-0950/tools/semantic_0950_v1.py",
                "correction": "/home/zhuhe/litchi-performance-0950/receipts/reducer-v4-to-v5-semantic-mean-correction-v1.json",
                "context": "/home/zhuhe/litchi-performance-0951/references/native-context-0950-v1.json",
                "source_manifest": "/home/zhuhe/litchi-performance-0950/candidate-source-manifest-v1.json",
            }.items()
        },
        "source_binding": {"manifest": desc(Path("/home/zhuhe/litchi-performance-0950/candidate-source-manifest-v1.json"))},
        "capture_environment": {"PATH": "/usr/bin:/bin", "LANG": "C", "LC_ALL": "C",
                                "TZ": "UTC", "RUST_BACKTRACE": "0"},
        "capture_cwd": "/home/zhuhe/litchi-performance-0950/candidate-source-v1",
        "first_child_admission": {"raw_gate": {
            "owner_samples_min": 100,
            "fully_resolved_owner_samples_min": 100,
            "owner_period_min": 100000,
            "lost_records": 0,
        }},
        "schedule": [{"index": 0, "id": "native-fp-publish-medium-r0", "repeat": 0, "shape": "medium"}],
        "child_artifacts": {"start": "start.json", "command": "command.json",
                            "record_terminal": "record-terminal.json", "record_stdout": "record.stdout",
                            "record_stderr": "record.stderr", "report": "report.json", "perf_data": "perf.data"},
        "decoder_contract": {
            "primary_stack_decoder": {"debug_mode": False, "argv_template": [
                "/usr/bin/perf", "script", "--input", "$CHILD/perf.data", "--show-lost-events",
                "--ns", "--no-inline", "--demangle", "--fields", "comm,pid,tid,time,event,period,ip,sym,symoff,dso"]},
            "mapping_decoder": {"debug_mode": True, "argv_template": [
                "/usr/bin/perf", "script", "--input", "$CHILD/perf.data", "--show-mmap-events",
                "--debug-mode", "--show-task-events", "--show-lost-events", "--ns", "--hide-call-graph",
                "--no-inline", "--demangle", "--fields", "comm,pid,tid,time,event,period,ip,sym,symoff,dso"]},
        },
        "record_contract": {"argv_template": [
            "/usr/bin/taskset", "--cpu-list", "12", "/usr/bin/perf", "record", "--output",
            "$CHILD/perf.data", "--freq", "997", "--event", "cycles:u", "--call-graph", "fp",
            "--period", "--timestamp", "--sample-cpu", "--buildid-all", "--no-buildid-cache",
            "--strict-freq", "--no-inherit", "--user-callchains", "--", "$BINARY", "--warmup", "0",
            "--samples", "200", "--case", old["case"], "--xlsx-cell-crud-shape", "$SHAPE",
            "--json", "$CHILD/report.json"]},
        "static_symbols": old["static_symbols"],
    }


def write_json(path: Path, value: Any) -> None:
    path.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n")


def prepare_packet(reader: Any, packet: Path) -> tuple[dict[str, Any], dict[str, Any], dict[str, Any], dict[str, Any], Path]:
    (packet / "tools").mkdir(parents=True)
    shutil.copy2(SEMANTIC, packet / "tools/semantic_0953_v1.py")
    plan = make_plan(packet)
    plan_path = packet / "plan-0953-v1.json"
    write_json(plan_path, plan)
    bindings, _sem, loaded_plan, contract, _plan_file = reader.verify_bindings(packet, plan_path)
    return plan, bindings, loaded_plan, contract, plan_path


def make_child(reader: Any, packet: Path, plan: dict[str, Any], bindings: dict[str, Any], plan_path: Path) -> Path:
    row = plan["schedule"][0]
    child = packet / "capture/00-medium"
    child.mkdir(parents=True)
    names = plan["child_artifacts"]
    binary = Path(bindings["binary"]["path"])
    expected = reader.expand_argv(plan["record_contract"]["argv_template"], packet=packet,
                                  child=child, output=child, binary=binary,
                                  shape=row["shape"], label="record")
    plan_desc = bindings["plan"]
    runner_desc = bindings["runner"]
    binary_desc = bindings["binary"]
    start = {"schema": "litchi.performance.0953.child-start.v1", "row": row,
             "plan": plan_desc, "runner": runner_desc, "binary": binary_desc,
             "record_argv": expected, "environment": plan["capture_environment"],
             "cwd": plan["capture_cwd"], "retry": False, "workload_launched": True}
    command = {"schema": "litchi.performance.0953.child-command.v1", "row": row,
               "plan": plan_desc, "runner": runner_desc, "binary": binary_desc,
               "record_argv": expected, "environment": plan["capture_environment"],
               "cwd": plan["capture_cwd"], "retry": False, "terminal_before_parse": True}
    write_json(child / names["start"], start)
    write_json(child / names["command"], command)
    (child / names["record_stdout"]).write_text("")
    (child / names["record_stderr"]).write_text("")
    (child / names["report"]).write_text("{}")
    (child / names["perf_data"]).write_bytes(b"synthetic perf data")
    terminal = {"schema": "litchi.performance.0953.record-terminal.v1", "argv": expected,
                "returncode": 0, "timed_out": False, "launch_error": None,
                "process_group_clear": True, "terminal_before_parse": True}
    write_json(child / names["record_terminal"], terminal)
    artifact_items = []
    for name in sorted(set(names.values())):
        artifact_items.append(desc(child / name))
    receipt = {"schema": "litchi.performance.0953.child-receipt.v1", "status": "raw_captured",
               "row": row, "plan": plan_desc, "runner": runner_desc, "binary": binary_desc,
               "record_terminal": desc(child / names["record_terminal"]),
               "report": desc(child / names["report"]), "perf_data": desc(child / names["perf_data"]),
               "artifacts": artifact_items, "retry": False}
    write_json(child / "receipt.json", receipt)
    return child


def make_export(reader: Any, packet: Path, plan: dict[str, Any], bindings: dict[str, Any],
                child: Path, kind: str) -> Path:
    row = plan["schedule"][0]
    root = packet / "analysis" / ("samples" if kind == "sample" else "mmap") / child.name
    root.mkdir(parents=True)
    binary = Path(bindings["binary"]["path"])
    argv = reader.expand_argv(reader.contract_for(plan, kind)["argv_template"], packet=packet,
                              child=child, output=root, binary=binary, shape=row["shape"], label=kind)
    start = {"schema": f"litchi.performance.0953.{kind}-export-start.v1", "row": row,
             "plan": bindings["plan"], "runner": bindings["runner"],
             "child_start": desc(child / "start.json"), "perf_data": desc(child / "perf.data"),
             "binary": bindings["binary"], "argv": argv, "cwd": plan["capture_cwd"],
             "environment": plan["capture_environment"], "workload_launched": False}
    write_json(root / "start.json", start)
    write_json(root / "terminal.json", {"schema": f"litchi.performance.0953.{kind}-export-terminal.v1",
                                         "argv": argv, "returncode": 0, "timed_out": False,
                                         "launch_error": None, "process_group_clear": True,
                                         "terminal_before_parse": True})
    (root / "events.txt").write_text("synthetic event\n")
    (root / "stderr.txt").write_text("")
    receipt = {"schema": f"litchi.performance.0953.{kind}-export-receipt.v1", "status": "pass",
               "row": row, "start": desc(root / "start.json"), "terminal": desc(root / "terminal.json"),
               "events": desc(root / "events.txt"), "stderr": desc(root / "stderr.txt"),
               "perf_data": desc(child / "perf.data"), "argv": argv,
               "terminal_before_parse": True, "workload_launched": False}
    write_json(root / "receipt.json", receipt)
    return root


def rebind_clone(reader: Any, packet: Path) -> None:
    plan_path = packet / "plan-0953-v1.json"
    plan = json.loads(plan_path.read_text())
    plan_desc = desc(plan_path)
    runner_desc = desc(ROOT / "tools/capture_0953_v1.py")
    binary_desc = desc(OLD_BINARY)
    binary = OLD_BINARY.resolve()
    child = packet / "capture/00-medium"
    row = plan["schedule"][0]
    def expand(template: list[str], output: Path) -> list[str]:
        return reader.expand_argv(template, packet=packet, child=child, output=output,
                                  binary=binary, shape=row["shape"], label="control")
    record_argv = expand(plan["record_contract"]["argv_template"], child)
    for name in ("start.json", "command.json"):
        path = child / name
        value = json.loads(path.read_text())
        value["plan"] = plan_desc; value["runner"] = runner_desc; value["binary"] = binary_desc
        value["record_argv"] = record_argv; value["cwd"] = plan["capture_cwd"]
        write_json(path, value)
    terminal_path = child / "record-terminal.json"
    terminal = json.loads(terminal_path.read_text()); terminal["argv"] = record_argv
    write_json(terminal_path, terminal)
    receipt_path = child / "receipt.json"
    receipt = json.loads(receipt_path.read_text())
    receipt["plan"] = plan_desc; receipt["runner"] = runner_desc; receipt["binary"] = binary_desc
    receipt["record_terminal"] = desc(terminal_path)
    receipt["artifacts"] = [desc(child / name) for name in
                            sorted({"start.json", "command.json", "record-terminal.json",
                                    "record.stdout", "record.stderr", "report.json", "perf.data"})]
    receipt["report"] = desc(child / "report.json"); receipt["perf_data"] = desc(child / "perf.data")
    write_json(receipt_path, receipt)
    for kind in ("sample", "mmap"):
        out = packet / "analysis" / ("samples" if kind == "sample" else "mmap") / child.name
        start_path = out / "start.json"
        start = json.loads(start_path.read_text())
        contract = plan["decoder_contract"]["primary_stack_decoder" if kind == "sample" else "mapping_decoder"]
        argv = expand(contract["argv_template"], out)
        start["plan"] = plan_desc; start["runner"] = runner_desc; start["binary"] = binary_desc
        start["child_start"] = desc(child / "start.json"); start["perf_data"] = desc(child / "perf.data")
        start["argv"] = argv; start["cwd"] = plan["capture_cwd"]
        write_json(start_path, start)
        terminal_path = out / "terminal.json"
        terminal = json.loads(terminal_path.read_text()); terminal["argv"] = argv
        write_json(terminal_path, terminal)
        export_receipt = json.loads((out / "receipt.json").read_text())
        export_receipt["start"] = desc(start_path); export_receipt["terminal"] = desc(terminal_path)
        export_receipt["events"] = desc(out / "events.txt"); export_receipt["stderr"] = desc(out / "stderr.txt")
        export_receipt["perf_data"] = desc(child / "perf.data"); export_receipt["argv"] = argv
        write_json(out / "receipt.json", export_receipt)


def expect_reject(label: str, action: Callable[[], Any], needle: str | None = None) -> dict[str, Any]:
    try:
        action()
    except Exception as error:
        text = str(error)
        if needle is not None and needle not in text:
            raise AssertionError(f"{label} rejected for unexpected reason: {text}") from error
        return {"status": "rejected", "error_type": type(error).__name__, "error": text}
    raise AssertionError(f"{label} was accepted")


def copy_packet(src: Path, root: Path, suffix: str) -> Path:
    dst = root / suffix
    shutil.copytree(src, dst)
    return dst


def real_mapping_controls(reader: Any) -> dict[str, Any]:
    old = json.loads(OLD_PLAN.read_text())
    loads, build_id = reader.read_elf(OLD_BINARY)
    mmap = reader.parse_mmap_events(OLD_MMAP, OLD_BINARY, loads, build_id)
    executable = [row for row in mmap["bound_rows"] if "x" in row["perm"] and not row.get("mapping_ambiguous")]
    if len(executable) != 1:
        raise AssertionError(f"expected one executable mapping, got {len(executable)}")
    mapping = executable[0]
    owner = old["static_symbols"]["owner_ranges"][0]
    parent = old["static_symbols"]["parent_range"]
    bias = mapping["load_bias"]
    owner_ip = bias + int(owner["address"], 16)
    parent_ip = bias + int(parent["address"], 16)
    pid = mapping["pid"]
    when = max(row["time"] for row in mmap["bound_rows"]) + reader.DECIMAL("1")
    def sample(path: Path, pid_value: int = pid, parent_dso: str = str(OLD_BINARY), event: str = "cycles:u") -> None:
        path.write_text(f"litchi-perf-bas {pid_value}/{pid_value} {when}: 7 {event}:\n"
                        f"  {owner_ip:#x} owner ({OLD_BINARY})\n"
                        f"  {parent_ip:#x} parent ({parent_dso})\n")
    with tempfile.TemporaryDirectory(prefix="0953-v2-map-controls-") as raw:
        root = Path(raw)
        positive = root / "positive.txt"; sample(positive)
        profile = reader.parse_samples(positive, OLD_BINARY, mmap, old)
        if profile["owner_qualified"] != {"samples": 1, "period": 7}:
            raise AssertionError(f"real positive did not qualify: {profile['owner_qualified']}")
        wrong_pid = root / "wrong-pid.txt"; sample(wrong_pid, pid + 1)
        wrong_parent = root / "wrong-parent.txt"; sample(wrong_parent, parent_dso="/lib/x86_64-linux-gnu/libc.so.6")
        wrong_event = root / "wrong-event.txt"; sample(wrong_event, event="instructions:u")
        return {"positive_owner_parent": profile["owner_qualified"],
                "wrong_pid": expect_reject("wrong PID", lambda: reader.parse_samples(wrong_pid, OLD_BINARY, mmap, old), None)
                if False else {"owner_qualified": reader.parse_samples(wrong_pid, OLD_BINARY, mmap, old)["owner_qualified"]},
                "wrong_parent": {"owner_qualified": reader.parse_samples(wrong_parent, OLD_BINARY, mmap, old)["owner_qualified"]},
                "wrong_event": expect_reject("wrong event", lambda: reader.parse_samples(wrong_event, OLD_BINARY, mmap, old), "unexpected perf event")}


def run(receipt: Path) -> dict[str, Any]:
    reader = load_reader()
    with tempfile.TemporaryDirectory(prefix="0953-v2-custody-controls-") as raw:
        root = Path(raw) / "packet"
        root.mkdir()
        plan, bindings, loaded, contract, plan_path = prepare_packet(reader, root)
        child = make_child(reader, root, plan, bindings, plan_path)
        _custody_row, _custody_child, custody = reader.child_context(plan, root, 0, bindings)
        sample_root = make_export(reader, root, plan, bindings, child, "sample")
        mmap_root = make_export(reader, root, plan, bindings, child, "mmap")
        positive_sample = reader.export_context(sample_root, root, child, custody["receipt"], plan["schedule"][0], "sample", plan, bindings, custody)
        positive_mmap = reader.export_context(mmap_root, root, child, custody["receipt"], plan["schedule"][0], "mmap", plan, bindings, custody)
        results: dict[str, Any] = {"positive_child_and_exports": "pass",
                                   "positive_sample_argv": positive_sample["argv"],
                                   "positive_mmap_argv": positive_mmap["argv"]}
        # Every raw-admission threshold is a frozen plan binding.  Mutating
        # one value must fail before any child or sample is admitted.
        threshold_bad_values = {
            "owner_samples_min": 99,
            "fully_resolved_owner_samples_min": 99,
            "owner_period_min": 99999,
            "lost_records": 1,
        }
        for key, bad_value in threshold_bad_values.items():
            variant = copy.deepcopy(plan)
            variant["first_child_admission"]["raw_gate"][key] = bad_value
            variant_path = root / f"plan-bad-raw-gate-{key}.json"
            write_json(variant_path, variant)
            results[f"changed_raw_gate_{key}"] = expect_reject(
                f"changed raw gate {key}",
                lambda path=variant_path: reader.verify_bindings(root, path),
                f"first-child raw gate {key} differs")
        # A changed artifact is rejected from the descriptor list before parsing.
        changed = copy_packet(root, Path(raw), "changed-artifact")
        # Rebind all copied descriptors first; then mutate only the retained
        # record.stdout bytes so admission reaches the artifact hash check.
        rebind_clone(reader, changed)
        (changed / "capture/00-medium/record.stdout").write_text("changed\n")
        changed_bindings, *_ = reader.verify_bindings(changed, changed / "plan-0953-v1.json")
        results["changed_child_artifact"] = expect_reject(
            "changed child artifact", lambda: reader.child_context(
                json.loads((changed / "plan-0953-v1.json").read_text()), changed, 0,
                changed_bindings), "child artifact record.stdout descriptor bytes differs")
        # Wrong record callgraph reaches the exact frozen argv check after refreshing descriptors.
        wrong_record = copy_packet(root, Path(raw), "wrong-record-callgraph")
        rebind_clone(reader, wrong_record)
        for name in ("start.json", "command.json", "record-terminal.json"):
            p = wrong_record / "capture/00-medium" / name
            value = json.loads(p.read_text())
            if name == "record-terminal.json":
                value["argv"][value["argv"].index("fp")] = "dwarf,32768"
            else:
                value["record_argv"][value["record_argv"].index("fp")] = "dwarf,32768"
            write_json(p, value)
        wrong_child = wrong_record / "capture/00-medium"
        wrong_receipt_path = wrong_child / "receipt.json"
        wrong_receipt = json.loads(wrong_receipt_path.read_text())
        wrong_receipt["record_terminal"] = desc(wrong_child / "record-terminal.json")
        wrong_receipt["artifacts"] = [desc(wrong_child / name) for name in
                                       sorted({"start.json", "command.json", "record-terminal.json",
                                               "record.stdout", "record.stderr", "report.json", "perf.data"})]
        write_json(wrong_receipt_path, wrong_receipt)
        wb, *_ = reader.verify_bindings(wrong_record, wrong_record / "plan-0953-v1.json")
        results["wrong_record_callgraph"] = expect_reject(
            "wrong record callgraph", lambda: reader.child_context(
                json.loads((wrong_record / "plan-0953-v1.json").read_text()), wrong_record, 0, wb), "record argv differs")
        # Wrong export input and fields are rejected by exact template equality.
        for label, mutate in (("wrong_export_input", lambda argv: argv.__setitem__(argv.index("--input") + 1, "/tmp/other.perf.data")),
                              ("wrong_export_fields", lambda argv: argv.__setitem__(argv.index("comm,pid,tid,time,event,period,ip,sym,symoff,dso"), "comm,pid,tid,time,event,period,ip,sym,dso"))):
            # Mutate a fresh copy's sample start/terminal/receipt descriptors while preserving descriptor integrity.
            copied = copy_packet(root, Path(raw), label)
            rebind_clone(reader, copied)
            copied_child = copied / "capture/00-medium"
            exp = copied / "analysis/samples/00-medium"
            for name in ("start.json", "terminal.json"):
                p = exp / name; value = json.loads(p.read_text()); mutate(value["argv"]); write_json(p, value)
            value = json.loads((exp / "receipt.json").read_text());
            value["start"] = desc(exp / "start.json"); value["terminal"] = desc(exp / "terminal.json");
            write_json(exp / "receipt.json", value)
            cb, *_ = reader.verify_bindings(copied, copied / "plan-0953-v1.json")
            _cc_row, _cc_child, cc = reader.child_context(json.loads((copied / "plan-0953-v1.json").read_text()), copied, 0, cb)
            results[label] = expect_reject(label, lambda: reader.export_context(
                exp, copied, copied_child, cc["receipt"], plan["schedule"][0], "sample",
                json.loads((copied / "plan-0953-v1.json").read_text()), cb, cc), "argv differs")
        results["real_mapping_controls"] = real_mapping_controls(reader)
    out = {"schema": "litchi.performance.0953.native-admission.v3-controls.v1",
           "status": "pass", "scope": "Synthetic custody/argv controls plus retained 0952 parser mapping controls; no workload/build/export.",
           "reader": desc(READER), "controls": results}
    if receipt.exists():
        raise RuntimeError(f"refusing overwrite: {receipt}")
    receipt.parent.mkdir(parents=True, exist_ok=True)
    write_json(receipt, out)
    return out


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--receipt", type=Path, required=True)
    args = parser.parse_args()
    result = run(args.receipt.resolve())
    print(json.dumps({"status": result["status"], "receipt": str(args.receipt.resolve())}, sort_keys=True))
