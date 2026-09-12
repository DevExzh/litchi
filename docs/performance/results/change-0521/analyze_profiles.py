#!/usr/bin/env python3
"""Validate and compare the 0521 XLSX commit Callgrind profiles.

The capture driver uses one Callgrind process for each repeat/shape.  The
first three numbered dumps are lifecycle gates and the fourth is the single
timed ``MultiSourceEdit::commit`` call.  This analyzer keeps all four raw
dumps in the report, checks their positive incoming edges, and annotates the
fourth dump with a deterministic Perl environment.

The raw Callgrind edge parser is intentionally reused from the retained 0519
helpers.  The 0519 parser is not changed and is loaded by path so this report
does not acquire a dependency on the current source tree.  Callgrind Ir is a
mechanism diagnostic; it is not substituted for native latency or hardware
cycle measurements.
"""

from __future__ import annotations

import argparse
import importlib.util
import json
import os
from pathlib import Path
import re
import subprocess
import sys
from typing import Any


HERE = Path(__file__).resolve().parent
PLAN_PATH = HERE / "plan.json"
HELPER_DIR = HERE.parent / "change-0519"

OWNER = "litchi_xlsx::cell_values::source::MultiSourceEdit::commit"
LIFECYCLE_PARENT = "litchi_perf_baseline::run_xlsx_cell_value_lifecycle_gates"
DIRECT_PARENT = "litchi_perf_baseline::run_xlsx_cell_values_edit_save"
VALIDATOR = "litchi_xlsx::cell_values::validation::validate_xml"
EVENT_INTO_OWNED = "quick_xml::events::Event::into_owned"

NUMBERED_PARTS = (1, 2, 3, 4)
TIMING_FIELDS = frozenset(
    {"open_ns", "plan_ns", "commit_ns", "publication_ns", "reopen_ns"}
)
PERL_ENV = {"PERL_HASH_SEED": "0", "PERL_PERTURB_KEYS": "0"}

FUNCTION_RE = re.compile(r"^(fn|cfn)=\((\d+)\)(?:\s+(.*))?$")
SUMMARY_RE = re.compile(r"^summary:\s*(.*)$")
PART_RE = re.compile(r"^part:\s*(\d+)$")
TRIGGER_RE = re.compile(r"^desc:\s+Trigger:\s+(.*)$")
CALLS_RE = re.compile(r"\(([\d,]+)x\)")


class EvidenceError(ValueError):
    """A missing, malformed, or contradictory evidence artifact."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise EvidenceError(message)


def sha256(path: Path) -> str:
    try:
        import hashlib

        digest = hashlib.sha256()
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
        return digest.hexdigest()
    except OSError as error:
        raise EvidenceError(f"cannot hash {path}: {error}") from error


def read_json(path: Path) -> Any:
    try:
        with path.open(encoding="utf-8") as stream:
            return json.load(stream)
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise EvidenceError(f"cannot read JSON {path}: {error}") from error


def read_text(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8", errors="replace")
    except OSError as error:
        raise EvidenceError(f"cannot read {path}: {error}") from error


def load_0519_helpers() -> tuple[Any, Any]:
    """Load the immutable raw profile and annotation helpers from 0519."""

    raw_path = HELPER_DIR / "analyze_profiles.py"
    edges_path = HELPER_DIR / "compare_profile_lanes.py"
    require(raw_path.is_file(), f"missing retained helper {raw_path}")
    require(edges_path.is_file(), f"missing retained helper {edges_path}")

    raw_spec = importlib.util.spec_from_file_location(
        "litchi_0519_raw_analyze_profiles", raw_path
    )
    require(raw_spec is not None and raw_spec.loader is not None,
            f"cannot load retained helper {raw_path}")
    raw = importlib.util.module_from_spec(raw_spec)
    raw_spec.loader.exec_module(raw)

    # compare_profile_lanes.py imports the old parser by its historical module
    # name.  Install exactly that immutable module for the duration of its
    # load, just as the 0520 analyzer does.
    sys.modules["analyze_profiles"] = raw
    edges_spec = importlib.util.spec_from_file_location(
        "litchi_0519_profile_edges", edges_path
    )
    require(edges_spec is not None and edges_spec.loader is not None,
            f"cannot load retained helper {edges_path}")
    edges = importlib.util.module_from_spec(edges_spec)
    edges_spec.loader.exec_module(edges)
    return raw, edges


RAW, EDGES = load_0519_helpers()


def relative(path: Path) -> str:
    try:
        return str(path.relative_to(HERE))
    except ValueError as error:
        raise EvidenceError(f"path is outside evidence directory: {path}") from error


def plan_data() -> dict[str, Any]:
    plan = read_json(PLAN_PATH)
    require(isinstance(plan, dict), "plan is not an object")
    require(plan.get("profile", {}).get("owner") == OWNER,
            "plan profile owner differs from the selected owner")
    require(plan.get("profile", {}).get("shapes") == ["medium", "dense-sparse"],
            "plan profile shape matrix differs")
    require(plan.get("profile", {}).get("repeats") == 2,
            "plan profile repeat count differs")
    require(plan.get("profile", {}).get("warmup") == 0,
            "plan profile warmup differs")
    require(plan.get("profile", {}).get("samples") == 1,
            "plan profile sample count differs")
    return plan


def summary_ir(text: str, label: str) -> int:
    values: list[int] = []
    for line in text.splitlines():
        match = SUMMARY_RE.match(line.strip())
        if not match:
            continue
        for token in match.group(1).split():
            try:
                values.append(int(token.replace(",", "")))
            except ValueError:
                break
    require(values, f"{label}: missing usable Ir summary")
    return values[0]


def part_number(text: str, label: str) -> int:
    values = [int(match.group(1)) for match in
              (PART_RE.match(line.strip()) for line in text.splitlines()) if match]
    require(len(values) == 1, f"{label}: expected one part record, got {values}")
    return values[0]


def trigger(text: str, label: str) -> str:
    values = [match.group(1) for match in
              (TRIGGER_RE.match(line.strip()) for line in text.splitlines()) if match]
    require(len(values) == 1, f"{label}: expected one Trigger description, got {values}")
    return values[0]


def validate_numbered_dump(path: Path, part: int, expected_parent: str) -> dict[str, Any]:
    text = read_text(path)
    label = relative(path)
    require(part_number(text, label) == part,
            f"{label}: part number does not match .{part}")
    expected_trigger = f"--dump-after={OWNER}"
    require(trigger(text, label) == expected_trigger,
            f"{label}: Trigger does not identify the exact owner")
    require("events: Ir" in text.splitlines(), f"{label}: event set is not exactly Ir")
    summary = summary_ir(text, label)
    incoming = EDGES.raw_incoming_call_summary(path, OWNER)
    require(incoming["positive_edge_count"] == 1 and incoming["calls"] == 1,
            f"{label}: owner incoming edge is not exactly one positive call")
    edge = incoming["edges"][0]
    require(edge["caller"] == expected_parent,
            f"{label}: owner caller is {edge['caller']!r}, expected {expected_parent!r}")
    require(edge["inclusive_ir"] == summary,
            f"{label}: owner incoming Ir does not equal dump summary")
    return {
        "part": part,
        "file": label,
        "sha256": sha256(path),
        "summary_ir": summary,
        "incoming": edge,
        "expected_parent": expected_parent,
        "validation": {
            "part_matches_suffix": True,
            "exact_trigger": True,
            "one_positive_owner_call": True,
            "incoming_owner_ir_matches_summary": True,
        },
    }


def validate_final_dump(path: Path) -> dict[str, Any]:
    text = read_text(path)
    label = relative(path)
    require(part_number(text, label) == 5,
            f"{label}: final process dump must be part 5")
    require(trigger(text, label) == "Program termination",
            f"{label}: final process dump is not Program termination")
    summary = summary_ir(text, label)
    require(summary == 0, f"{label}: final process dump Ir is {summary}, expected zero")
    require("events: Ir" in text.splitlines(), f"{label}: final dump has no Ir event")
    return {
        "file": label,
        "sha256": sha256(path),
        "part": 5,
        "trigger": "Program termination",
        "summary_ir": summary,
        "validation": {"final_process_dump_zero_ir": True},
    }


def annotation_command(path: Path, inclusive: bool) -> list[str]:
    return [
        "callgrind_annotate",
        "--auto=no",
        "--threshold=100",
        "--show-percs=no",
        "--inclusive=" + ("yes" if inclusive else "no"),
        "--tree=both",
        str(path),
    ]


def run_annotation(path: Path, inclusive: bool) -> tuple[str, list[str]]:
    environment = dict(os.environ)
    environment.update(PERL_ENV)
    command = annotation_command(path, inclusive)
    try:
        process = subprocess.run(
            command,
            capture_output=True,
            text=True,
            check=True,
            env=environment,
        )
    except (OSError, subprocess.CalledProcessError) as error:
        raise EvidenceError(f"callgrind_annotate failed for {relative(path)}: {error}") from error
    require(process.stderr == "", f"callgrind_annotate emitted stderr for {relative(path)}")
    return process.stdout, command


def display_name(text: str) -> str:
    return EDGES.display_name(text)


def parse_annotation(text: str, selected: str, label: str) -> dict[str, Any]:
    """Read the selected function row and its immediate ``>`` children."""

    rows = []
    for index, line in enumerate(text.splitlines()):
        match = EDGES.STAR_RE.match(line)
        if match and display_name(match.group(2)) == selected:
            rows.append((index, int(match.group(1).replace(",", ""))))
    require(len(rows) == 1, f"{label}: expected one annotation row for {selected}, got {rows}")
    row_index, selected_ir = rows[0]
    direct: list[dict[str, Any]] = []
    lines = text.splitlines()
    for line in lines[row_index + 1:]:
        match = EDGES.EDGE_RE.match(line)
        if not match:
            break
        rendered_name = match.group(2)
        calls_match = CALLS_RE.search(rendered_name)
        direct.append({
            "name": display_name(rendered_name),
            "inclusive_ir": int(match.group(1).replace(",", "")),
            "calls": (int(calls_match.group(1).replace(",", ""))
                      if calls_match else None),
        })
    return {"selected_ir": selected_ir, "direct": direct}


def direct_map(edges: list[dict[str, Any]]) -> dict[str, int]:
    result: dict[str, int] = {}
    for edge in edges:
        name = edge["name"]
        result[name] = result.get(name, 0) + edge["inclusive_ir"]
    return result


def target_edge_summary(path: Path, target: str, caller: str | None = None) -> dict[str, Any]:
    incoming = EDGES.raw_incoming_call_summary(path, target)
    matching = incoming["edges"]
    if caller is not None:
        matching = [edge for edge in matching if edge["caller"] == caller]
    return {
        "target": target,
        "caller": caller,
        "positive_edge_count": len(matching),
        "calls": sum(edge["calls"] for edge in matching),
        "inclusive_ir": sum(edge["inclusive_ir"] for edge in matching),
        "edges": matching,
    }


def validate_event_edge(path: Path, stage: str) -> dict[str, Any]:
    edge = target_edge_summary(path, EVENT_INTO_OWNED, VALIDATOR)
    present = edge["positive_edge_count"] > 0
    if stage == "baseline":
        require(edge["positive_edge_count"] == 1,
                f"{relative(path)}: baseline validator -> Event::into_owned edge missing or duplicated")
    else:
        require(edge["positive_edge_count"] == 0,
                f"{relative(path)}: candidate retained a positive validator -> Event::into_owned edge")
    return {
        **edge,
        "present": present,
        "policy": "present" if stage == "baseline" else "absent",
    }


def option(command: list[str], name: str) -> str:
    try:
        index = command.index(name)
        return command[index + 1]
    except (ValueError, IndexError) as error:
        raise EvidenceError(f"profile command is missing {name}") from error


def validate_receipt_artifacts(stage_dir: Path, receipt: dict[str, Any],
                               expected: set[str], label: str) -> None:
    artifacts = receipt.get("artifacts")
    require(isinstance(artifacts, dict), f"{label}: artifacts is not an object")
    require(expected.issubset(artifacts),
            f"{label}: receipt omits {sorted(expected - set(artifacts))}")
    for filename, digest in artifacts.items():
        require(isinstance(filename, str) and Path(filename).name == filename,
                f"{label}: artifact path is not a stage-local basename: {filename!r}")
        path = stage_dir / filename
        require(path.is_file() and not path.is_symlink(),
                f"{label}: artifact is missing: {filename}")
        require(sha256(path) == digest, f"{label}: artifact hash differs: {filename}")


def validate_build(stage: str, plan: dict[str, Any]) -> dict[str, Any]:
    stage_dir = HERE / stage
    source_manifest = stage_dir / "source-manifest.json"
    build_receipt_path = stage_dir / "build-normal.receipt.json"
    binary_path = stage_dir / "binary-normal.json"
    require(source_manifest.is_file(), f"{stage}: source manifest is missing")
    require(build_receipt_path.is_file(), f"{stage}: build receipt is missing")
    require(binary_path.is_file(), f"{stage}: binary identity is missing")
    source_manifest_sha = sha256(source_manifest)
    plan_sha = sha256(PLAN_PATH)
    run_sha = sha256(HERE / "run.py")
    build = read_json(build_receipt_path)
    require(build.get("exit_code") == 0, f"{stage}: build did not exit successfully")
    require(build.get("source_manifest_sha256") == source_manifest_sha,
            f"{stage}: build/source manifest binding differs")
    require(build.get("plan_sha256") == plan_sha, f"{stage}: build/plan binding differs")
    require(build.get("script_sha256") == run_sha, f"{stage}: build/script binding differs")
    validate_receipt_artifacts(stage_dir, build, {"build-normal.stdout", "build-normal.stderr"},
                               f"{stage}/build-normal")
    binary = read_json(binary_path)
    require(isinstance(binary, dict), f"{stage}: binary identity is not an object")
    binary_sha = binary.get("sha256")
    require(isinstance(binary_sha, str) and len(binary_sha) == 64,
            f"{stage}: binary hash is missing")
    require(binary.get("source_manifest_sha256") == source_manifest_sha,
            f"{stage}: binary/source manifest binding differs")
    require(binary.get("build_receipt_sha256") == sha256(build_receipt_path),
            f"{stage}: binary/build receipt binding differs")
    binary_file = Path(binary.get("path", ""))
    if binary_file.is_file() and not binary_file.is_symlink():
        require(sha256(binary_file) == binary_sha,
                f"{stage}: available binary hash differs from identity")
    return {
        "source_manifest_sha256": source_manifest_sha,
        "source_manifest": relative(source_manifest),
        "plan_sha256": plan_sha,
        "run_script_sha256": run_sha,
        "build_receipt": relative(build_receipt_path),
        "build_receipt_sha256": sha256(build_receipt_path),
        "binary_identity": relative(binary_path),
        "binary_sha256": binary_sha,
        "binary_path": binary.get("path"),
    }


def expected_profile_names(plan: dict[str, Any]) -> list[dict[str, Any]]:
    profile = plan["profile"]
    jobs = []
    for repeat in range(1, profile["repeats"] + 1):
        for shape in profile["shapes"]:
            jobs.append({
                "name": f"profile-r{repeat}-{shape}",
                "repeat": repeat,
                "shape": shape,
            })
    return jobs


def expected_native_name(repeat: int, shape: str) -> str:
    return f"native-r{repeat}-primary-{shape}"


def canonical(value: Any, label: str) -> Any:
    """Collapse a constant sample vector into its value for parity checks."""

    if isinstance(value, list):
        require(bool(value), f"{label}: empty vector")
        values = [canonical(item, f"{label}[{index}]") for index, item in enumerate(value)]
        require(all(item == values[0] for item in values),
                f"{label}: vector is not constant")
        return values[0]
    if isinstance(value, dict):
        return {
            key: canonical(item, f"{label}.{key}")
            for key, item in sorted(value.items())
        }
    return value


def source_counter_identity(source: Any, label: str) -> Any:
    require(isinstance(source, dict), f"{label}: source is not an object")
    result: dict[str, Any] = {}
    for key, value in sorted(source.items()):
        if key == "xlsx_cell_values":
            require(isinstance(value, dict), f"{label}.{key}: not an object")
            nested = {}
            for nested_key, nested_value in sorted(value.items()):
                if nested_key in TIMING_FIELDS or nested_key == "commit_allocation_metrics":
                    continue
                nested[nested_key] = canonical(nested_value, f"{label}.{key}.{nested_key}")
            result[key] = nested
        else:
            result[key] = canonical(value, f"{label}.{key}")
    return result


def result_identity(result: dict[str, Any], label: str) -> dict[str, Any]:
    require(isinstance(result, dict), f"{label}: result is not an object")
    require("corpus" in result and "sink" in result and "source" in result,
            f"{label}: result is missing corpus/sink/source")
    return {
        "corpus": result["corpus"],
        "sink": result["sink"],
        "source_counters": source_counter_identity(result["source"], label + ".source"),
        "output_sha256": result.get("output_sha256"),
    }


def validate_result_report(path: Path, binary_sha: str, plan: dict[str, Any],
                           repeat: int, shape: str, profile_report: bool) -> tuple[dict[str, Any], dict[str, Any]]:
    report = read_json(path)
    label = relative(path)
    require(report.get("schema_version") == 1, f"{label}: schema version is not 1")
    tool = report.get("tool")
    require(isinstance(tool, dict), f"{label}: tool is not an object")
    require(tool.get("binary") == "litchi-perf-baseline", f"{label}: unexpected binary")
    require(tool.get("profile") == "release", f"{label}: report is not release profile")
    require(tool.get("instrumentation") == "none", f"{label}: report is instrumented")
    identity = report.get("binary_identity")
    require(isinstance(identity, dict) and identity.get("binary_sha256") == binary_sha,
            f"{label}: report is bound to a different binary")
    environment = report.get("environment")
    require(isinstance(environment, dict) and environment.get("git_revision") == plan["revision"],
            f"{label}: report revision differs from plan")
    configuration = report.get("configuration")
    require(isinstance(configuration, dict), f"{label}: configuration is not an object")
    require(configuration.get("cases") == [plan["primary"]["case"]],
            f"{label}: case configuration differs")
    require(configuration.get("xlsx_cell_crud_shapes") == [shape],
            f"{label}: shape configuration differs")
    expected_samples = 1 if profile_report else plan["primary"]["samples"]
    expected_warmup = 0 if profile_report else plan["primary"]["warmup"]
    require(configuration.get("samples_per_case") == expected_samples,
            f"{label}: sample configuration differs")
    require(configuration.get("warmup_iterations_per_case") == expected_warmup,
            f"{label}: warmup configuration differs")
    results = report.get("results")
    require(isinstance(results, list) and len(results) == 1,
            f"{label}: expected one result")
    result = results[0]
    return report, result_identity(result, label + ".results[0]")


def validate_native_receipt(stage_dir: Path, native_name: str, stage_meta: dict[str, Any],
                            plan: dict[str, Any]) -> dict[str, Any]:
    receipt_path = stage_dir / f"{native_name}.receipt.json"
    output_path = stage_dir / f"{native_name}.json"
    require(receipt_path.is_file(), f"{relative(output_path)}: native receipt is missing")
    receipt = read_json(receipt_path)
    require(receipt.get("exit_code") == 0, f"{relative(receipt_path)}: native child failed")
    require(receipt.get("binary_sha256") == stage_meta["binary_sha256"],
            f"{relative(receipt_path)}: native binary binding differs")
    require(receipt.get("source_manifest_sha256") == stage_meta["source_manifest_sha256"],
            f"{relative(receipt_path)}: native source binding differs")
    require(receipt.get("plan_sha256") == stage_meta["plan_sha256"],
            f"{relative(receipt_path)}: native plan binding differs")
    expected = {
        output_path.name,
        f"{native_name}.stdout",
        f"{native_name}.stderr",
        f"{native_name}.rss.json",
    }
    validate_receipt_artifacts(stage_dir, receipt, expected, relative(receipt_path))
    return {
        "file": relative(output_path),
        "receipt": relative(receipt_path),
        "output_sha256": sha256(output_path),
        "receipt_sha256": sha256(receipt_path),
    }


def validate_profile_receipt(stage_dir: Path, name: str, receipt: dict[str, Any],
                             stage_meta: dict[str, Any], plan: dict[str, Any], shape: str) -> None:
    label = f"{stage_dir.name}/{name}.receipt.json"
    require(receipt.get("exit_code") == 0, f"{label}: profile child failed")
    require(receipt.get("binary_sha256") == stage_meta["binary_sha256"],
            f"{label}: profile binary binding differs")
    require(receipt.get("source_manifest_sha256") == stage_meta["source_manifest_sha256"],
            f"{label}: profile source binding differs")
    require(receipt.get("plan_sha256") == stage_meta["plan_sha256"],
            f"{label}: profile plan binding differs")
    require(receipt.get("script_sha256") == sha256(HERE / "run.py"),
            f"{label}: profile script binding differs")
    command = receipt.get("command")
    require(isinstance(command, list), f"{label}: command is not a list")
    owner_options = {
        "--warmup": "0",
        "--samples": "1",
        "--case": plan["primary"]["case"],
        "--xlsx-cell-crud-shape": shape,
    }
    for option_name, expected_value in owner_options.items():
        require(option(command, option_name) == expected_value,
                f"{label}: {option_name} differs")
    required_tokens = [
        "taskset",
        "-c",
        str(plan["cpu"]),
        "valgrind",
        "--tool=callgrind",
        "--collect-atstart=no",
        f"--toggle-collect={OWNER}",
        f"--zero-before={OWNER}",
        f"--dump-after={OWNER}",
    ]
    for token in required_tokens:
        require(token in command, f"{label}: command omits {token}")
    try:
        output_path = Path(option(command, "--json"))
    except EvidenceError as error:
        raise EvidenceError(f"{label}: command output path is missing") from error
    require(output_path.name == f"{name}.json",
            f"{label}: command output path differs")
    expected_artifacts = {
        f"{name}.json",
        f"{name}.stdout",
        f"{name}.stderr",
        f"{name}.callgrind",
        *(f"{name}.callgrind.{part}" for part in NUMBERED_PARTS),
    }
    validate_receipt_artifacts(stage_dir, receipt, expected_artifacts, label)


def annotate_profile(stage: str, stage_dir: Path, name: str, dump: Path,
                     expected_total: int, expected_self: int | None = None) -> dict[str, Any]:
    inclusive_text, inclusive_command = run_annotation(dump, True)
    self_text, self_command = run_annotation(dump, False)
    inclusive_path = stage_dir / f"{name}.inclusive.txt"
    self_path = stage_dir / f"{name}.self.txt"
    inclusive_path.write_text(inclusive_text, encoding="utf-8")
    self_path.write_text(self_text, encoding="utf-8")
    inclusive = parse_annotation(inclusive_text, OWNER, relative(inclusive_path))
    exclusive = parse_annotation(self_text, OWNER, relative(self_path))
    require(inclusive["selected_ir"] == expected_total,
            f"{relative(inclusive_path)}: inclusive owner Ir differs from raw edge")
    require(exclusive["direct"] == inclusive["direct"],
            f"{relative(self_path)}: direct children differ from inclusive annotation")
    # The selected inclusive row is the raw owner total.  The self annotation
    # row plus its visible direct children must reconstruct that total.  The
    # threshold is retained from 0520; these profiles have the same complete
    # visible owner edge set under that threshold.
    require(exclusive["selected_ir"] + sum(
        edge["inclusive_ir"] for edge in exclusive["direct"]
    ) == inclusive["selected_ir"],
            f"{relative(self_path)}: self plus direct Ir does not equal inclusive Ir")
    if expected_self is not None:
        require(exclusive["selected_ir"] == expected_self,
                f"{relative(self_path)}: self owner Ir differs from raw owner accounting")
    validator_values = [
        int(match.group(1).replace(",", ""))
        for line in inclusive_text.splitlines()
        if (match := EDGES.STAR_RE.match(line))
        and display_name(match.group(2)) == VALIDATOR
    ]
    require(len(validator_values) == 1,
            f"{relative(inclusive_path)}: expected one validator inclusive row, got {validator_values}")
    validator_inclusive = validator_values[0]
    return {
        "environment": dict(PERL_ENV),
        "command": {
            "inclusive": [
                *inclusive_command[:-1],
                relative(dump),
            ],
            "self": [
                *self_command[:-1],
                relative(dump),
            ],
        },
        "files": {
            "inclusive": relative(inclusive_path),
            "inclusive_sha256": sha256(inclusive_path),
            "self": relative(self_path),
            "self_sha256": sha256(self_path),
        },
        "owner": {
            "inclusive_ir": inclusive["selected_ir"],
            "self_ir": exclusive["selected_ir"],
            "direct_callees": inclusive["direct"],
            "direct_callee_ir": direct_map(inclusive["direct"]),
        },
        "validator": {
            "inclusive_ir": validator_inclusive,
            "source": "callgrind_annotate --inclusive=yes --tree=both",
            "caveat": (
                "This is an inclusive Ir row. It includes child calls; child-call "
                "metadata in the annotation can include calls attributable outside "
                "the selected direct edge, so it is a mechanism diagnostic."
            ),
        },
        "validation": {
            "inclusive_owner_matches_raw": True,
            "self_plus_direct_equals_inclusive": True,
            "inclusive_and_self_direct_children_match": True,
            "validator_row_present": True,
        },
    }


def analyze_profile(stage: str, plan: dict[str, Any], stage_meta: dict[str, Any],
                    job: dict[str, Any]) -> dict[str, Any]:
    stage_dir = HERE / stage
    name = job["name"]
    shape = job["shape"]
    repeat = job["repeat"]
    receipt_path = stage_dir / f"{name}.receipt.json"
    profile_path = stage_dir / f"{name}.json"
    require(receipt_path.is_file(), f"{relative(receipt_path)}: receipt is missing")
    require(profile_path.is_file(), f"{relative(profile_path)}: profile result is missing")
    receipt = read_json(receipt_path)
    validate_profile_receipt(stage_dir, name, receipt, stage_meta, plan, shape)
    profile_report, profile_identity = validate_result_report(
        profile_path, stage_meta["binary_sha256"], plan, repeat, shape, True
    )
    native_name = expected_native_name(repeat, shape)
    native_path = stage_dir / f"{native_name}.json"
    require(native_path.is_file(), f"{relative(native_path)}: native counterpart is missing")
    native_receipt_meta = validate_native_receipt(stage_dir, native_name, stage_meta, plan)
    native_report, native_identity = validate_result_report(
        native_path, stage_meta["binary_sha256"], plan, repeat, shape, False
    )
    require(profile_identity == native_identity,
            f"{name}: profile/native corpus, sink, output, or counter identity differs")

    raw_dumps = []
    for part in NUMBERED_PARTS:
        dump = stage_dir / f"{name}.callgrind.{part}"
        require(dump.is_file(), f"{relative(dump)}: numbered raw dump is missing")
        parent = LIFECYCLE_PARENT if part < 4 else DIRECT_PARENT
        raw_dumps.append(validate_numbered_dump(dump, part, parent))
    final_dump = stage_dir / f"{name}.callgrind"
    require(final_dump.is_file(), f"{relative(final_dump)}: final process dump is missing")
    final = validate_final_dump(final_dump)
    event_edge = validate_event_edge(stage_dir / f"{name}.callgrind.4", stage)
    validator_incoming = target_edge_summary(
        stage_dir / f"{name}.callgrind.4", VALIDATOR, None
    )
    annotation = annotate_profile(
        stage,
        stage_dir,
        name,
        stage_dir / f"{name}.callgrind.4",
        raw_dumps[-1]["summary_ir"],
    )
    return {
        "name": name,
        "repeat": repeat,
        "shape": shape,
        "receipt": relative(receipt_path),
        "receipt_sha256": sha256(receipt_path),
        "profile_result": relative(profile_path),
        "profile_result_sha256": sha256(profile_path),
        "native": native_receipt_meta,
        "native_result_identity": native_identity,
        "raw_dumps": raw_dumps,
        "final_process_dump": final,
        "annotations": annotation,
        "validator_raw_incoming": validator_incoming,
        "validator_event_into_owned_edge": event_edge,
        "total_ir": annotation["owner"]["inclusive_ir"],
        "self_ir": annotation["owner"]["self_ir"],
        "direct_ir": sum(annotation["owner"]["direct_callee_ir"].values()),
        "direct_callee_ir": annotation["owner"]["direct_callee_ir"],
        "validator_inclusive_ir": annotation["validator"]["inclusive_ir"],
        "validation": {
            "receipt_and_command": True,
            "profile_native_identity_parity": True,
            "all_numbered_raw_dumps_retained_and_valid": True,
            "final_process_dump_zero_ir": True,
            "validator_event_edge_policy": True,
            **annotation["validation"],
        },
    }


def stage_has_profiles(stage: str, plan: dict[str, Any]) -> bool:
    stage_dir = HERE / stage
    jobs = expected_profile_names(plan)
    return stage_dir.is_dir() and any(
        (stage_dir / f"{job['name']}.callgrind.4").is_file() for job in jobs
    )


def analyze_stage(stage: str, plan: dict[str, Any]) -> dict[str, Any]:
    stage_dir = HERE / stage
    require(stage_dir.is_dir(), f"stage directory is missing: {stage}")
    meta = validate_build(stage, plan)
    profiles = [analyze_profile(stage, plan, meta, job)
                for job in expected_profile_names(plan)]
    require(len(profiles) == 4, f"{stage}: profile matrix is incomplete")
    return {
        "stage": stage,
        "metadata": meta,
        "profile_count": len(profiles),
        "profiles": profiles,
        "validation": {
            "expected_profile_matrix": True,
            "all_receipts_and_hashes_valid": True,
            "all_native_output_parity_valid": True,
            "all_raw_parts_valid": True,
            "all_final_dumps_zero_ir": True,
        },
    }


def metric(baseline: int | None, candidate: int | None) -> dict[str, Any]:
    result = {
        "baseline": baseline,
        "candidate": candidate,
        "delta_ir": (candidate - baseline
                     if baseline is not None and candidate is not None else None),
        "candidate_over_baseline": None,
        "delta_percent": None,
    }
    if baseline not in (None, 0) and candidate is not None:
        result["candidate_over_baseline"] = candidate / baseline
        result["delta_percent"] = (candidate / baseline - 1.0) * 100.0
    return result


def profile_by_key(stage_report: dict[str, Any]) -> dict[tuple[int, str], dict[str, Any]]:
    return {(item["repeat"], item["shape"]): item for item in stage_report["profiles"]}


def edge_presence(item: dict[str, Any]) -> dict[str, Any]:
    edge = item["validator_event_into_owned_edge"]
    return {
        "present": edge["present"],
        "positive_edge_count": edge["positive_edge_count"],
        "calls": edge["calls"],
        "inclusive_ir": (edge["inclusive_ir"] if edge["present"] else None),
        "edges": edge["edges"],
    }


def compare_stages(baseline: dict[str, Any], candidate: dict[str, Any]) -> dict[str, Any]:
    left = profile_by_key(baseline)
    right = profile_by_key(candidate)
    require(set(left) == set(right), "baseline/candidate profile matrices differ")
    rows = []
    aggregate_names: set[str] = set()
    aggregate: dict[str, dict[str, int]] = {
        "total_ir": {"baseline": 0, "candidate": 0},
        "self_ir": {"baseline": 0, "candidate": 0},
        "direct_ir": {"baseline": 0, "candidate": 0},
        "validator_inclusive_ir": {"baseline": 0, "candidate": 0},
    }
    for key in sorted(left):
        base = left[key]
        cand = right[key]
        aggregate_names.update(base["direct_callee_ir"])
        aggregate_names.update(cand["direct_callee_ir"])
        for field in aggregate:
            aggregate[field]["baseline"] += base[field]
            aggregate[field]["candidate"] += cand[field]
        base_identity = base["native_result_identity"]
        cand_identity = cand["native_result_identity"]
        identity_equal = base_identity == cand_identity
        require(identity_equal,
                f"{base['name']}: baseline/candidate native identity differs")
        names = set(base["direct_callee_ir"]) | set(cand["direct_callee_ir"])
        direct = [
            {
                "name": name,
                "metric": metric(
                    base["direct_callee_ir"].get(name, 0),
                    cand["direct_callee_ir"].get(name, 0),
                ),
            }
            for name in sorted(names,
                              key=lambda name: (
                                  -max(base["direct_callee_ir"].get(name, 0),
                                       cand["direct_callee_ir"].get(name, 0)),
                                  name,
                              ))
        ]
        rows.append({
            "repeat": key[0],
            "shape": key[1],
            "baseline_profile": base["name"],
            "candidate_profile": cand["name"],
            "native_identity_equal": identity_equal,
            "metrics": {
                "commit_inclusive_ir": metric(base["total_ir"], cand["total_ir"]),
                "commit_self_ir": metric(base["self_ir"], cand["self_ir"]),
                "commit_direct_ir": metric(base["direct_ir"], cand["direct_ir"]),
                "validator_inclusive_ir": metric(
                    base["validator_inclusive_ir"], cand["validator_inclusive_ir"]
                ),
            },
            "direct_callee_ir": direct,
            "validate_xml_to_event_into_owned": {
                "baseline": edge_presence(base),
                "candidate": edge_presence(cand),
                "interpretation": (
                    "Baseline has one positive validate_xml -> Event::into_owned "
                    "edge; candidate policy requires that positive edge to be absent."
                ),
            },
        })
    aggregate_direct = [
        {
            "name": name,
            "metric": metric(
                sum(item["direct_callee_ir"].get(name, 0) for item in left.values()),
                sum(item["direct_callee_ir"].get(name, 0) for item in right.values()),
            ),
        }
        for name in sorted(
            aggregate_names,
            key=lambda name: (
                -max(
                    sum(item["direct_callee_ir"].get(name, 0) for item in left.values()),
                    sum(item["direct_callee_ir"].get(name, 0) for item in right.values()),
                ),
                name,
            ),
        )
    ]
    return {
        "available": True,
        "profile_count": len(rows),
        "profiles": rows,
        "aggregate_totals": {
            field: metric(values["baseline"], values["candidate"])
            for field, values in aggregate.items()
        },
        "aggregate_direct_callee_ir": aggregate_direct,
        "validator_inclusive_cost_caveat": (
            "Validator inclusive Ir includes child calls. Callgrind child-call "
            "metadata can include calls attributable outside the selected direct "
            "validate_xml -> Event::into_owned edge; use this row as mechanism "
            "evidence alongside the raw edge policy and native parity."
        ),
        "validation": {
            "profile_matrix_matches": True,
            "native_identity_parity": True,
            "direct_callee_ir_compared": True,
            "validator_edge_baseline_present_candidate_absent": all(
                row["validate_xml_to_event_into_owned"]["baseline"]["present"]
                and not row["validate_xml_to_event_into_owned"]["candidate"]["present"]
                for row in rows
            ),
        },
    }


def analyze(stage_selection: str = "both") -> dict[str, Any]:
    plan = plan_data()
    selected = []
    if stage_selection == "baseline":
        selected = ["baseline"]
    elif stage_selection == "candidate":
        selected = ["candidate"]
    else:
        selected = [stage for stage in ("baseline", "candidate")
                    if stage_has_profiles(stage, plan)]
        require(selected, "no completed baseline or candidate profile stage is available")
    stages = {stage: analyze_stage(stage, plan) for stage in selected}
    comparison = None
    if "baseline" in stages and "candidate" in stages:
        comparison = compare_stages(stages["baseline"], stages["candidate"])
    return {
        "schema": "xlsx_callgrind_commit_profile_analysis_v1",
        "status": "pass",
        "selected_function": OWNER,
        "plan": relative(PLAN_PATH),
        "plan_sha256": sha256(PLAN_PATH),
        "stage_selection": selected,
        "stages": stages,
        "comparison": comparison,
        "helpers": {
            "analyze_profiles.py": sha256(HELPER_DIR / "analyze_profiles.py"),
            "compare_profile_lanes.py": sha256(HELPER_DIR / "compare_profile_lanes.py"),
        },
        "limitations": (
            "Callgrind Ir and inclusive rows are mechanism diagnostics, not native "
            "latencies, hardware instructions, cycles, allocation counts, cold-cache, "
            "range, scaling, or native Office-producer measurements."
        ),
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("output", nargs="?", type=Path,
                        help="JSON destination; defaults to profile-analysis.json")
    parser.add_argument("--output", dest="output_option", type=Path,
                        help="JSON destination (alternative to the positional path)")
    parser.add_argument("--stage", choices=("baseline", "candidate", "both"), default="both",
                        help="validate one stage, or both stages when available")
    args = parser.parse_args()
    if args.output is not None and args.output_option is not None:
        parser.error("provide the output path either positionally or with --output")
    output = args.output_option or args.output or (HERE / "profile-analysis.json")
    try:
        document = analyze(args.stage)
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(json.dumps(document, indent=2) + "\n", encoding="utf-8")
    except (EvidenceError, OSError, json.JSONDecodeError) as error:
        print(f"analyze_profiles.py: error: {error}", file=sys.stderr)
        return 2
    print(f"Profile raw scope, native parity, and annotations verified: {output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
