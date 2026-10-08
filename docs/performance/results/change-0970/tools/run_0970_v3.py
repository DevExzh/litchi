#!/usr/bin/env python3
"""Serial 0970 historical-baseline affinity diagnostic runner."""
from __future__ import annotations

import argparse
import json
import os
import signal
import subprocess
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from protocol_0970_v3 import (
    BOOTSTRAP_RESAMPLES, BOOTSTRAP_SEED, CASES, CORRECTION_RECEIPT, DESIGN_SHA256, EVENTS,
    EXPECTED_ALLOWED_AFFINITY, EXPECTED_PIN_CPU, REQUIRED_LOCALE,
    GO_SCHEMA, HELPER_SEAL_SCHEMA, MANIFEST_SCHEMA, MAX_STDOUT_BYTES,
    PAIRS, PLAN_SCHEMA, RETENTION, ROOT, SCHEMA, SAMPLES, SOURCE_ROOT,
    STAGE_SCHEMA, TIMEOUT_SECONDS, VALIDATOR_PATH, VALIDATOR_SHA256,
    WARMUPS, ProtocolError, binding, cell_by_id, cells, check_binding,
    check_packet, expected_command, expected_order, parse_affinity,
    parse_launch_sidecar, sha256, strict_load_json, validate_affinity_controls,
    validate_source_custody,
    write_json_exclusive,
)

BASELINE_BINARY = Path("/home/zhuhe/litchi-performance-0967/builds/target-baseline-cfb/release/doc_retention_probe_0942_cfb_profile")
BASELINE_SHA256 = "0db34357a1e725dc397a4fd895532433db1214c5862b7bc96bf23f21c0396ecb"
FIXTURES = {
    "docnohf": Path("/home/zhuhe/litchi-performance-0969/qa-core/test-data/ole/doc/NoHeadFoot.doc"),
    "docfloat": Path("/home/zhuhe/litchi-performance-0969/qa-core/test-data/ole/doc/FloatingPictures.doc"),
}
VALIDATOR_SCHEMA = "litchi.performance.0969.protocol_common.v4"
TIME_FORMAT = '{"user_s":%U,"system_s":%S,"voluntary_context_switches":%w,"involuntary_context_switches":%c,"minor_faults":%R,"major_faults":%F,"max_rss_kb":%M}'
SUSPICIOUS_TOKENS = (
    "cargo", "rustc", "rustdoc", "pytest", "perf stat", "doc_retention_probe",
    "run_0969", "run_0970", "criterion", "bencher",
)


def utc() -> str:
    return datetime.now(timezone.utc).isoformat()


def json_object_from_bytes(data: bytes, label: str) -> dict[str, Any]:
    if len(data) > MAX_STDOUT_BYTES:
        raise ProtocolError(f"{label} exceeds {MAX_STDOUT_BYTES} bytes")
    lines = [line for line in data.decode("utf-8", errors="strict").splitlines() if line.strip()]
    if not lines:
        raise ProtocolError(f"{label} has no JSON report")

    def pairs(items: list[tuple[str, Any]]) -> dict[str, Any]:
        result: dict[str, Any] = {}
        for key, value in items:
            if key in result:
                raise ProtocolError(f"{label} has duplicate JSON key {key!r}")
            result[key] = value
        return result

    try:
        value = json.loads(lines[-1], object_pairs_hook=pairs)
    except (UnicodeError, json.JSONDecodeError) as exc:
        raise ProtocolError(f"{label} final line is not JSON: {exc}") from exc
    if not isinstance(value, dict):
        raise ProtocolError(f"{label} final JSON is not an object")
    return value


def path_row(path: Path) -> dict[str, Any]:
    if path.is_file():
        return binding(path)
    return {"path": str(path.resolve()), "sha256": None, "bytes": 0}


def process_ancestors() -> set[int]:
    result: set[int] = set()
    current = os.getpid()
    while current > 1 and current not in result:
        result.add(current)
        try:
            current = int(Path(f"/proc/{current}/stat").read_text().split(")", 1)[1].split()[1])
        except (OSError, ValueError, IndexError):
            break
    return result


def quiescence_scan() -> dict[str, Any]:
    ancestors = process_ancestors()
    offenders: list[dict[str, Any]] = []
    for entry in Path("/proc").iterdir():
        if not entry.name.isdigit():
            continue
        pid = int(entry.name)
        if pid in ancestors:
            continue
        try:
            command = (entry / "cmdline").read_bytes().replace(b"\x00", b" ").decode("utf-8", errors="replace").strip()
        except OSError:
            continue
        lowered = command.lower()
        if any(token in lowered for token in SUSPICIOUS_TOKENS):
            offenders.append({"pid": pid, "command": command[:512]})
    return {"status": "pass" if not offenders else "fail", "checked_utc": utc(), "offenders": offenders}


def make_preexec(mask: list[int]):
    def set_mask() -> None:
        os.sched_setaffinity(0, set(mask))
    return set_mask


def child_environment() -> dict[str, str]:
    environment = os.environ.copy()
    environment["LC_ALL"] = REQUIRED_LOCALE
    environment["LANG"] = REQUIRED_LOCALE
    environment["PYTHONDONTWRITEBYTECODE"] = "1"
    return environment


def run_affinity_control(python: Path, condition: str, target: list[int], output: Path) -> dict[str, Any]:
    script = (
        "import json,os,sys;"
        "target=json.loads(sys.argv[1]);"
        "before=sorted(os.sched_getaffinity(0));"
        "os.sched_setaffinity(0,set(target));"
        "after=sorted(os.sched_getaffinity(0));"
        "os.sched_setaffinity(0,set(before));"
        "restored=sorted(os.sched_getaffinity(0));"
        "print(json.dumps({'condition':sys.argv[2],'before':before,'target':target,'after':after,'restored':restored},sort_keys=True))"
    )
    try:
        completed = subprocess.run(
            [str(python), "-c", script, json.dumps(target, separators=(",", ":")), condition],
            stdin=subprocess.DEVNULL, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
            text=False, timeout=20, check=False, preexec_fn=make_preexec(target), env=child_environment(),
        )
    except (OSError, subprocess.TimeoutExpired) as exc:
        raise ProtocolError(f"affinity control {condition} failed: {exc}") from exc
    if completed.returncode != 0:
        raise ProtocolError(f"affinity control {condition} exit {completed.returncode}: {completed.stderr!r}")
    report = json_object_from_bytes(completed.stdout, f"affinity control {condition}")
    if set(report) != {"condition", "before", "target", "after", "restored"}:
        raise ProtocolError("affinity control fields drifted")
    if report["condition"] != condition:
        raise ProtocolError("affinity control condition drifted")
    for field in ("before", "target", "after", "restored"):
        report[field] = parse_affinity(report[field], f"affinity control {field}")
    if report["target"] != target or report["after"] != target or report["restored"] != report["before"]:
        raise ProtocolError(f"affinity control mask mismatch for {condition}")
    write_json_exclusive(output, report)
    return report


def check_plan(plan_path: Path, *, require_pre_capture: bool = False) -> dict[str, Any]:
    plan = strict_load_json(plan_path.resolve())
    if plan.get("schema") != PLAN_SCHEMA or plan.get("status") != "pre_capture":
        raise ProtocolError("0970 plan schema/status drifted")
    design = plan.get("design")
    check_binding(design, "design v3", expected_path=ROOT / "diagnostic-design-v3.md")
    if design["sha256"] != DESIGN_SHA256:
        raise ProtocolError("design v3 hash drifted")
    reviews = plan.get("reviews") or {}
    reuse = reviews.get("baseline_reuse") or {}
    reuse_path = check_binding(reuse, "baseline reuse review")
    if reuse.get("status") != "pass":
        raise ProtocolError("baseline reuse review status is not pass")
    if strict_load_json(reuse_path).get("status") != "pass_bounded_output_oracle_and_matched_baseline_reuse":
        raise ProtocolError("baseline reuse review content is not the bounded pass")
    check_binding(reviews.get("source_variability"), "source variability review")
    if (reviews.get("independent_tool_review") or {}).get("status") != "pending":
        raise ProtocolError("independent tool review must remain pending before admission")
    helper_seal_path = check_binding(plan.get("helper_seal"), "helper seal")
    seal = strict_load_json(helper_seal_path)
    if seal.get("schema") != HELPER_SEAL_SCHEMA or seal.get("status") != "pass" or (seal.get("design") or {}).get("sha256") != DESIGN_SHA256:
        raise ProtocolError("helper seal is not a passing design-v3 seal")
    correction = check_binding(seal.get("correction_receipt"), "helper seal correction receipt")
    correction_doc = strict_load_json(correction)
    if correction.resolve() != CORRECTION_RECEIPT.resolve() or correction_doc.get("schema") != "litchi.performance.0970.correction_receipt.v2" or correction_doc.get("status") != "pass":
        raise ProtocolError("helper seal correction receipt is invalid")
    prior_correction = check_binding(correction_doc.get("prior_correction"), "prior F1-F5 correction receipt", expected_path=ROOT / "receipts/seal-timestamp-correction-0970-v1.json")
    if strict_load_json(prior_correction).get("original_seal_time") != "unknown":
        raise ProtocolError("prior correction does not preserve unknown original seal time")
    plan_correction = check_binding(plan.get("correction_receipt"), "plan correction receipt", expected_path=CORRECTION_RECEIPT)
    if sha256(plan_correction) != sha256(correction):
        raise ProtocolError("plan/helper correction receipt bindings differ")
    helpers = plan.get("helpers") or {}
    required_helpers = {"launcher", "protocol", "runner", "reducer", "plan_generator"}
    if set(helpers) != required_helpers:
        raise ProtocolError("helper set drifted")
    for name in required_helpers:
        check_binding(helpers[name], f"helper {name}")
        row = (seal.get("helpers") or {}).get(name)
        if not isinstance(row, dict) or row.get("path") != helpers[name]["path"] or row.get("sha256") != helpers[name]["sha256"]:
            raise ProtocolError(f"helper seal mismatch for {name}")
    source = plan.get("source")
    source_summary = validate_source_custody(source)
    if source["manifest"]["sha256"] != "f1818af5c306fb6ea01b59ca8df29710e773386b246071202bd4aa896f4e8e2b" or source["file_manifest"]["sha256"] != "99e65d43e0feed89133eec60b095130d68058a071a02e8a15b05dea01209bb6c":
        raise ProtocolError("source manifest binding drifted")
    if source["root"] != str(SOURCE_ROOT.resolve()) or source_summary["entries"] != 924:
        raise ProtocolError("source root/entry binding drifted")
    validator = plan.get("validator") or {}
    check_binding(validator, "0969 strict validator", expected_path=VALIDATOR_PATH)
    if validator.get("sha256") != VALIDATOR_SHA256 or validator.get("schema") != VALIDATOR_SCHEMA:
        raise ProtocolError("0969 strict validator binding drifted")
    binary = plan.get("binary") or {}
    check_binding(binary, "historical baseline binary", expected_path=BASELINE_BINARY)
    if binary.get("sha256") != BASELINE_SHA256 or not (BASELINE_BINARY.stat().st_mode & 0o111):
        raise ProtocolError("historical baseline binary binding drifted")
    fixtures = plan.get("fixtures") or {}
    if set(fixtures) != set(CASES):
        raise ProtocolError("fixture set drifted")
    for case in CASES:
        check_binding(fixtures[case], f"fixture {case}", expected_path=FIXTURES[case])
    host = plan.get("host") or {}
    for name in ("audit", "tool_addendum", "telemetry_probe", "affinity_probe"):
        check_binding(host.get(name), f"host {name}")
    tools = host.get("tools") or {}
    if set(tools) != {"/usr/bin/perf", "/usr/bin/time", "/usr/bin/python3.14"}:
        raise ProtocolError("host tool set drifted")
    addendum = strict_load_json(Path(host["tool_addendum"]["path"]))
    addendum_rows = {row["resolved_path"]: row for row in addendum.get("tools", [])}
    for path in tools:
        check_binding(tools[path], f"tool {path}", expected_path=Path(path))
        if path not in addendum_rows or tools[path]["sha256"] != addendum_rows[path]["resolved_path_sha256"]:
            raise ProtocolError(f"tool addendum binding drifted for {path}")
    matrix = plan.get("matrix") or {}
    if matrix.get("cells") != cells() or (matrix.get("pairs_per_cell"), matrix.get("warmups"), matrix.get("samples")) != (PAIRS, WARMUPS, SAMPLES):
        raise ProtocolError("matrix cell/sample schedule drifted")
    if matrix.get("seed") != BOOTSTRAP_SEED or matrix.get("bootstrap_resamples") != BOOTSTRAP_RESAMPLES:
        raise ProtocolError("matrix bootstrap binding drifted")
    if matrix.get("order") != {"odd_pair": ["unpinned", "pinned"], "even_pair": ["pinned", "unpinned"]}:
        raise ProtocolError("matrix order binding drifted")
    affinity = plan.get("affinity") or {}
    if affinity.get("selector") != "sealed_live_mask_exact:[0..31]" or affinity.get("requested_actual_must_match") is not True or affinity.get("expected_allowed") != EXPECTED_ALLOWED_AFFINITY or affinity.get("expected_pin_cpu") != EXPECTED_PIN_CPU:
        raise ProtocolError("affinity selector/mask binding drifted")
    wrapper = plan.get("wrapper") or {}
    if wrapper.get("events") != list(EVENTS) or wrapper.get("time_format") != TIME_FORMAT:
        raise ProtocolError("wrapper event/format binding drifted")
    if Path(wrapper.get("launcher", "")).resolve() != Path(helpers["launcher"]["path"]).resolve():
        raise ProtocolError("wrapper launcher binding drifted")
    execution = plan.get("execution") or {}
    if execution.get("serial") is not True or execution.get("retry") is not False or execution.get("process_quiescence_required") is not True or execution.get("candidate") is not False or execution.get("timeout_seconds") != TIMEOUT_SECONDS or execution.get("lc_all") != REQUIRED_LOCALE or execution.get("lang") != REQUIRED_LOCALE or execution.get("python_dontwritebytecode") != "1":
        raise ProtocolError("execution policy/locale drifted")
    environment = plan.get("environment") or {}
    if environment != {"LC_ALL": REQUIRED_LOCALE, "LANG": REQUIRED_LOCALE, "PYTHONDONTWRITEBYTECODE": "1"}:
        raise ProtocolError("plan launch environment drifted")
    analysis = plan.get("analysis") or {}
    if analysis.get("diagnostic_only") is not True or analysis.get("no_qualification") is not True or analysis.get("statuses") != ["complete", "incomplete"]:
        raise ProtocolError("diagnostic interpretation policy drifted")
    outputs = plan.get("outputs") or {}
    if Path(outputs.get("capture_root", "")).resolve() != (ROOT / "capture/diagnostic-v3").resolve():
        raise ProtocolError("capture output root drifted")
    go = outputs.get("root_go") or {}
    if go.get("status") != "pending" or Path(go.get("path", "")).resolve() != (ROOT / "receipts/root-go-0970-v4.json").resolve():
        raise ProtocolError("root GO declaration drifted")
    if require_pre_capture and plan.get("no_build_or_workload") is not True:
        raise ProtocolError("pre-capture plan must remain workload-free")
    return plan


def validate_go(go_path: Path, plan_path: Path, plan: dict[str, Any]) -> dict[str, Any]:
    go = strict_load_json(go_path.resolve())
    if go.get("schema") != GO_SCHEMA or go.get("authorized") is not True or go.get("status") != "authorized" or go.get("no_retry") is not True:
        raise ProtocolError("root GO authorization is invalid")
    quiescence = go.get("quiescence") or {}
    if quiescence.get("status") != "pass":
        raise ProtocolError("root GO lacks a passing quiescence receipt")
    check_binding(quiescence, "root GO quiescence")
    environment = go.get("environment")
    if environment != {"LC_ALL": REQUIRED_LOCALE, "LANG": REQUIRED_LOCALE, "PYTHONDONTWRITEBYTECODE": "1"}:
        raise ProtocolError("root GO environment is not the sealed C environment")
    host = go.get("host") or {}
    if host.get("allowed_affinity") != EXPECTED_ALLOWED_AFFINITY or host.get("pin_cpu") != EXPECTED_PIN_CPU:
        raise ProtocolError("root GO audited affinity mask/pin is invalid")
    for name in ("audit", "tool_addendum", "telemetry_probe", "affinity_probe"):
        expected = plan["host"][name]
        observed = host.get(name)
        if not isinstance(observed, dict) or observed.get("path") != expected["path"] or observed.get("sha256") != expected["sha256"]:
            raise ProtocolError(f"root GO host {name} binding mismatch")
    review = go.get("independent_review") or {}
    review_path = check_binding(review, "independent tool review")
    review_doc = strict_load_json(review_path)
    if review.get("status") != "pass" or review_doc.get("status") != "pass":
        raise ProtocolError("root GO lacks a passing independent tool review")
    row = go.get("plan") or {}
    if row.get("path") != str(plan_path.resolve()) or row.get("sha256") != sha256(plan_path) or row.get("bytes") != plan_path.stat().st_size:
        raise ProtocolError("root GO plan binding mismatch")
    for name, expected in (("design", plan["design"]), ("helper_seal", plan["helper_seal"]), ("validator", plan["validator"]), ("binary", plan["binary"])):
        observed = go.get(name)
        if not isinstance(observed, dict) or observed.get("path") != expected["path"] or observed.get("sha256") != expected["sha256"]:
            raise ProtocolError(f"root GO {name} binding mismatch")
    for name, expected in plan["helpers"].items():
        observed = (go.get("helpers") or {}).get(name)
        if not isinstance(observed, dict) or observed.get("path") != expected["path"] or observed.get("sha256") != expected["sha256"]:
            raise ProtocolError(f"root GO helper {name} binding mismatch")
    for name, expected in plan["source"].items():
        if name not in {"manifest", "file_manifest"}:
            continue
        observed = (go.get("source") or {}).get(name)
        if not isinstance(observed, dict) or observed.get("path") != expected["path"] or observed.get("sha256") != expected["sha256"]:
            raise ProtocolError(f"root GO source {name} binding mismatch")
    for case, expected in plan["fixtures"].items():
        observed = (go.get("fixtures") or {}).get(case)
        if not isinstance(observed, dict) or observed.get("path") != expected["path"] or observed.get("sha256") != expected["sha256"]:
            raise ProtocolError(f"root GO fixture {case} binding mismatch")
    correction = go.get("correction_receipt")
    if not isinstance(correction, dict):
        raise ProtocolError("root GO correction receipt binding is missing")
    correction_path = check_binding(correction, "root GO correction receipt")
    correction_doc = strict_load_json(correction_path)
    if correction_path.resolve() != CORRECTION_RECEIPT.resolve() or correction_doc.get("schema") != "litchi.performance.0970.correction_receipt.v2" or correction_doc.get("status") != "pass":
        raise ProtocolError("root GO correction receipt is invalid")
    prior_correction = check_binding(correction_doc.get("prior_correction"), "root GO prior correction receipt", expected_path=ROOT / "receipts/seal-timestamp-correction-0970-v1.json")
    if strict_load_json(prior_correction).get("original_seal_time") != "unknown":
        raise ProtocolError("root GO prior correction does not preserve unknown original seal time")
    return {"path": str(go_path.resolve()), "sha256": sha256(go_path), "receipt": go}


def path_row(path: Path) -> dict[str, Any]:
    return binding(path) if path.is_file() else {"path": str(path.resolve()), "sha256": None, "bytes": 0}


def build_packet(plan: dict[str, Any], output_root: Path, context: dict[str, Any], expected_mask: list[int]) -> dict[str, Any]:
    packet_id = context["packet_id"]
    raw = output_root / "raw"
    side = output_root / "sidecars"
    raw.mkdir(parents=True, exist_ok=True)
    side.mkdir(parents=True, exist_ok=True)
    paths = {
        "stdout": raw / f"{packet_id}.stdout",
        "stderr": raw / f"{packet_id}.stderr",
        "launch": side / f"{packet_id}.launch.json",
        "time": side / f"{packet_id}.time.json",
        "perf": side / f"{packet_id}.perf.csv",
    }
    if any(path.exists() for path in paths.values()):
        raise ProtocolError(f"refusing to overwrite packet artifact {packet_id}")
    context = dict(context)
    context["expected_affinity"] = expected_mask
    packet = {
        "schema": SCHEMA, "packet_id": packet_id, "created_utc": utc(),
        "output_root": str(output_root.resolve()), "context": context,
        "binary": {"path": str(BASELINE_BINARY.resolve()), "sha256": plan["binary"]["sha256"]},
        "fixture": plan["fixtures"][context["cell"]["case"]],
        "sidecars": {name: path_row(paths[name]) for name in ("launch", "time", "perf")},
        "stdout": path_row(paths["stdout"]), "stderr": path_row(paths["stderr"]),
        "command": [], "terminal": {"returncode": None, "timed_out": False, "launch_error": None},
        "report": None, "launcher": None, "validation_error": None,
    }
    packet["command"] = expected_command(plan, packet)
    return packet


def kill_process_group(process: subprocess.Popen[bytes]) -> None:
    try:
        os.killpg(process.pid, signal.SIGKILL)
    except ProcessLookupError:
        pass


def launch_packet(plan: dict[str, Any], output_root: Path, context: dict[str, Any], expected_mask: list[int]) -> dict[str, Any]:
    packet = build_packet(plan, output_root, context, expected_mask)
    stdout_path = output_root / "raw" / f"{context['packet_id']}.stdout"
    stderr_path = output_root / "raw" / f"{context['packet_id']}.stderr"
    start = time.monotonic_ns()
    try:
        with stdout_path.open("xb") as stdout, stderr_path.open("xb") as stderr:
            process = subprocess.Popen(
                packet["command"], stdin=subprocess.DEVNULL, stdout=stdout, stderr=stderr,
                start_new_session=True, preexec_fn=make_preexec(expected_mask), env=child_environment(),
            )
            try:
                process.wait(timeout=TIMEOUT_SECONDS)
            except subprocess.TimeoutExpired:
                packet["terminal"]["timed_out"] = True
                kill_process_group(process)
                process.wait()
            packet["terminal"]["returncode"] = process.returncode
    except (OSError, ValueError) as exc:
        packet["terminal"]["launch_error"] = str(exc)
    packet["terminal"]["wall_ns_excluded_from_metric"] = time.monotonic_ns() - start
    packet["stdout"] = path_row(stdout_path)
    packet["stderr"] = path_row(stderr_path)
    for name in ("launch", "time", "perf"):
        packet["sidecars"][name] = path_row(Path(packet["sidecars"][name]["path"]))
    try:
        if packet["terminal"]["returncode"] != 0 or packet["terminal"]["timed_out"] or packet["terminal"]["launch_error"]:
            raise ProtocolError("child terminal status failed")
        packet["report"] = json_object_from_bytes(stdout_path.read_bytes(), f"{context['packet_id']} stdout")
        launch = parse_launch_sidecar(Path(packet["sidecars"]["launch"]["path"]), helper_sha256=plan["helpers"]["launcher"]["sha256"], packet_id=context["packet_id"], expected_mask=expected_mask, binary=BASELINE_BINARY, helper_path=Path(plan["helpers"]["launcher"]["path"]))
        packet["launcher"] = launch
        packet["context"]["observed_affinity"] = launch["actual_affinity"]
        check_packet(packet, plan)
    except (OSError, UnicodeError, ProtocolError, ValueError, KeyError, TypeError, json.JSONDecodeError) as exc:
        packet["validation_error"] = str(exc)
    return packet


def run_capture(plan_path: Path, plan: dict[str, Any], go_binding: dict[str, Any], output_root: Path) -> dict[str, Any]:
    if output_root.exists():
        raise ProtocolError(f"fresh output root required; refusing existing path {output_root}")
    allowed = parse_affinity(sorted(os.sched_getaffinity(0)), "live allowed affinity")
    if allowed != EXPECTED_ALLOWED_AFFINITY:
        raise ProtocolError(f"live allowed affinity drifted: {allowed}")
    pin_cpu = max(allowed)
    if pin_cpu != EXPECTED_PIN_CPU:
        raise ProtocolError(f"live pin CPU drifted: {pin_cpu}")
    output_root.mkdir(parents=True)
    quiescence = quiescence_scan()
    write_json_exclusive(output_root / "quiescence-v2.json", quiescence)
    if quiescence["status"] != "pass":
        stage = {"schema": STAGE_SCHEMA, "status": "incomplete", "reason": "process quiescence failed", "packet_count": 0, "no_retry": True}
        write_json_exclusive(output_root / "capture-stage-v2.json", stage)
        write_json_exclusive(output_root / "capture-manifest-v2.json", {
            "schema": MANIFEST_SCHEMA, "no_retry": True, "complete": False,
            "plan": {"path": str(plan_path.resolve()), "sha256": sha256(plan_path)}, "root_go": go_binding, "packets": [],
        })
        return stage
    controls = {}
    python = Path(plan["host"]["tools"]["/usr/bin/python3.14"]["path"])
    for condition, target in (("unpinned", allowed), ("pinned", [pin_cpu])):
        controls[condition] = run_affinity_control(python, condition, target, output_root / f"affinity-control-{condition}.json")
    controls_path = output_root / "affinity-controls-v2.json"
    controls_document = {
        "schema": "litchi.performance.0970.affinity_controls.v2", "allowed": allowed, "pin_cpu": pin_cpu, "controls": controls,
    }
    validate_affinity_controls(controls_document)
    write_json_exclusive(controls_path, controls_document)
    controls_binding = binding(controls_path)
    packets: list[Path] = []
    failure = None
    for cell in cells():
        for pair in range(1, PAIRS + 1):
            order = expected_order(pair)
            for index, condition in enumerate(order):
                packet_id = f"diag-{cell['case']}-e{cell['edits']}-default-p{pair:02d}-{condition}"
                context = {
                    "packet_id": packet_id, "cell_id": cell["cell_id"], "cell": cell,
                    "pair": pair, "order": list(order), "order_position": "first" if index == 0 else "second",
                    "condition": condition, "source_manifest_sha256": plan["source"]["manifest"]["sha256"],
                    "source_file_manifest_sha256": plan["source"]["file_manifest"]["sha256"],
                    "binary_sha256": plan["binary"]["sha256"], "fixture_sha256": plan["fixtures"][cell["case"]]["sha256"],
                    "launcher_sha256": plan["helpers"]["launcher"]["sha256"],
                    "lc_all": REQUIRED_LOCALE,
                    "lang": REQUIRED_LOCALE,
                    "python_dontwritebytecode": "1",
                }
                packet = launch_packet(plan, output_root, context, allowed if condition == "unpinned" else [pin_cpu])
                packet_path = output_root / "packets" / f"{packet_id}.json"
                write_json_exclusive(packet_path, packet)
                packets.append(packet_path)
                if packet.get("validation_error") is not None:
                    failure = f"{packet_id}: {packet['validation_error']}"
                    break
            if failure:
                break
        if failure:
            break
    complete = failure is None and len(packets) == len(cells()) * PAIRS * 2
    stage = {
        "schema": STAGE_SCHEMA, "status": "complete" if complete else "incomplete",
        "packet_count": len(packets), "expected_packet_count": len(cells()) * PAIRS * 2,
        "no_retry": True, "failure": failure, "pin_cpu": pin_cpu, "allowed_affinity": allowed,
        "lc_all": REQUIRED_LOCALE, "lang": REQUIRED_LOCALE, "python_dontwritebytecode": "1",
    }
    write_json_exclusive(output_root / "capture-stage-v2.json", stage)
    write_json_exclusive(output_root / "capture-manifest-v2.json", {
        "schema": MANIFEST_SCHEMA, "no_retry": True, "complete": complete,
        "plan": {"path": str(plan_path.resolve()), "sha256": sha256(plan_path)}, "root_go": go_binding,
        "affinity_controls": controls_binding,
        "packets": [str(path.resolve()) for path in packets],
        "stage": str((output_root / "capture-stage-v2.json").resolve()),
    })
    return stage


def self_test() -> int:
    from protocol_0970_v3 import parse_perf_sidecar, safe_ratio, strict_load_json
    import tempfile
    with tempfile.TemporaryDirectory(prefix="litchi-0970-runner-self-test-") as directory:
        root = Path(directory)
        perf = root / "valid.csv"
        perf.write_text("\n".join(f"{idx + 1},,{event},1,100.00,," for idx, event in enumerate(EVENTS)) + "\n", encoding="utf-8")
        parsed = parse_perf_sidecar(perf)
        if parsed["all_events_present"] is not True or parsed["multiplexed_events"]:
            raise SystemExit("valid perf parser self-test failed")
        wrapper_perf = root / "wrapper.perf.csv"
        wrapper_time = root / "wrapper.time.json"
        wrapper_command = [
            "/usr/bin/perf", "stat", "-x,", "-o", str(wrapper_perf),
            "-e", ",".join(EVENTS), "--", "/usr/bin/time", "-f", TIME_FORMAT,
            "-o", str(wrapper_time), "/usr/bin/true",
        ]
        wrapped = subprocess.run(
            wrapper_command, stdin=subprocess.DEVNULL, stdout=subprocess.PIPE,
            stderr=subprocess.PIPE, timeout=20, check=False, env=child_environment(),
        )
        if wrapped.returncode != 0 or not wrapper_perf.is_file():
            raise SystemExit(f"actual perf wrapper self-test failed: {wrapped.returncode}: {wrapped.stderr!r}")
        wrapped_parsed = parse_perf_sidecar(wrapper_perf)
        if wrapped_parsed["all_events_present"] is not True or set(wrapped_parsed["events"]) != set(EVENTS):
            raise SystemExit("actual perf wrapper header self-test failed")
        bad = root / "bad.csv"
        bad.write_text("1,,task-clock,1,100.00,,\n<not counted>,,cycles,1,0.00,,\n", encoding="utf-8")
        try:
            parse_perf_sidecar(bad)
        except ProtocolError:
            pass
        else:
            raise SystemExit("not-counted perf parser self-test failed")
        rows = [f"{idx + 1},,{event},1,100.00,," for idx, event in enumerate(EVENTS)]
        malformed = root / "malformed.csv"
        malformed.write_text("\n".join(rows + ["malformed-row"]) + "\n", encoding="utf-8")
        duplicate_event = root / "duplicate-event.csv"
        duplicate_rows = list(rows)
        duplicate_rows[-1] = "6,,task-clock,1,100.00,,"
        duplicate_event.write_text("\n".join(duplicate_rows) + "\n", encoding="utf-8")
        missing_event = root / "missing-event.csv"
        missing_event.write_text("\n".join(rows[:-1]) + "\n", encoding="utf-8")
        unexpected_comment = root / "unexpected-comment.csv"
        unexpected_comment.write_text("# not the perf header\n" + "\n".join(rows) + "\n", encoding="utf-8")
        for rejected, label in ((malformed, "malformed row"), (duplicate_event, "duplicate event"), (missing_event, "missing event"), (unexpected_comment, "unexpected comment")):
            try:
                parse_perf_sidecar(rejected)
            except ProtocolError:
                pass
            else:
                raise SystemExit(f"{label} was accepted")
        if safe_ratio(0.0, 0.0, "zero_iqr") != {"value": None, "reason": "zero_iqr"}:
            raise SystemExit("zero-dispersion ratio self-test failed")
        duplicate = root / "duplicate.json"
        duplicate.write_text('{"a":1,"a":2}', encoding="utf-8")
        try:
            strict_load_json(duplicate)
        except ProtocolError:
            pass
        else:
            raise SystemExit("duplicate JSON self-test failed")
    print(json.dumps({"schema": SCHEMA, "status": "self_test_pass", "workload_started": False, "build_started": False, "zero_dispersion_null": True, "not_counted_rejected": True, "actual_perf_header_accepted": True, "malformed_row_rejected": True, "duplicate_event_rejected": True, "missing_event_rejected": True, "unexpected_comment_rejected": True}))
    return 0


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--plan", type=Path)
    parser.add_argument("--check-plan", action="store_true")
    parser.add_argument("--run", action="store_true")
    parser.add_argument("--go", type=Path)
    parser.add_argument("--output-root", type=Path)
    parser.add_argument("--self-test", action="store_true")
    args = parser.parse_args()
    if args.self_test:
        return self_test()
    if args.plan is None:
        print("runner rejected input: --plan is required", file=sys.stderr)
        return 2
    try:
        plan_path = args.plan.resolve()
        plan = check_plan(plan_path, require_pre_capture=not args.run)
        if not args.run:
            print(json.dumps({"schema": SCHEMA, "status": "preflight_pass", "plan": str(plan_path), "no_workload_started": True}))
            return 0
        if args.go is None or args.output_root is None:
            raise ProtocolError("--run requires --go and --output-root")
        if args.output_root.resolve() != Path(plan["outputs"]["capture_root"]).resolve():
            raise ProtocolError("--output-root must equal the bound plan output root")
        go_binding = validate_go(args.go, plan_path, plan)
        stage = run_capture(plan_path, plan, go_binding, args.output_root.resolve())
        print(json.dumps(stage, sort_keys=True))
        return 0 if stage.get("status") == "complete" else 2
    except (OSError, ProtocolError, ValueError, KeyError, TypeError, UnicodeError, json.JSONDecodeError) as exc:
        print(f"runner rejected input: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
