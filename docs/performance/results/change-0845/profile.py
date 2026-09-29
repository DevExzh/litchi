"""Capture and independently reduce the 0845 Callgrind side population.

This module is deliberately separate from the main native/Heaptrack lanes.  It
never builds the probe and never changes the main plan.  Capture uses the two
already admitted release binaries, while analysis replays every retained
report and raw Callgrind file before storing the small symbol/edge projection
used by the review.
"""

from __future__ import annotations

import json
import sys
import time
from pathlib import Path
from typing import Any, Mapping, Sequence

import driver as d
import reader as r


P = Path(__file__).resolve().parent
ROOT = P.parents[3]
PLAN_PATH = P / "profile-plan.json"
PROFILE_DIR = P / "profile-runs"
CAPTURE_PATH = P / "capture-profile.json"
ANALYSIS_PATH = P / "profile-analysis.json"
PROFILE_OPTIONS = tuple(r.PROFILE_OPTIONS)
PROFILE_OWNER = r.PROFILE_OWNER
LEGS = ("before", "after")
SHAPES = ("shared", "deep-unique")
PARTS = (64, 4096)


def fail(message: str) -> None:
    raise AssertionError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def read(path: Path) -> Any:
    return d.read(path)


def bind(value: Mapping[str, Any], label: str = "descriptor") -> dict[str, Any]:
    require(isinstance(value, Mapping), f"{label}: descriptor is not an object")
    path = Path(value.get("path", ""))
    actual = d.desc(path)
    require(actual == dict(value), f"{label}: descriptor changed: {path}")
    return actual


def descriptor(path: Path) -> dict[str, Any]:
    return d.desc(path)


def plan() -> dict[str, Any]:
    value = read(PLAN_PATH)
    require(value.get("schema") == "litchi.opc-preservation-callgrind-plan-0845.v1",
            "profile plan schema changed")
    require(value.get("base") == d.read(P / "origin.json").get("base"),
            "profile plan base differs from packet origin")
    require(value.get("cpu") == 12, "profile CPU changed")
    require(value.get("mode") == "publish", "profile mode changed")
    require(value.get("parts") == list(PARTS), "profile part set changed")
    require(value.get("shapes") == list(SHAPES), "profile shape set changed")
    require(value.get("legs") == list(LEGS), "profile leg set changed")
    require(value.get("repeats") == 2 and value.get("samples") == 1
            and value.get("warmup") == 0, "profile sample contract changed")
    require(value.get("reports") == 16 and value.get("formal_samples") == 16,
            "profile population size changed")
    require(value.get("profile_owner") == PROFILE_OWNER, "profile owner changed")
    require(tuple(value.get("profile_options", ())) == PROFILE_OPTIONS,
            "Callgrind options changed")
    provenance = value.get("reader_provenance")
    require(isinstance(provenance, Mapping)
            and provenance.get("previous_packet") == "../change-0844/reader.py"
            and provenance.get("sha256") == "bf8b4c65775ef49130815e90b20b29110c00535fff43d356a8344f6b1ac882cf",
            "reader provenance changed")
    previous_reader = P.parent / "change-0844" / "reader.py"
    require(d.sha(P / "reader.py") == provenance["sha256"]
            and d.sha(previous_reader) == provenance["sha256"],
            "0845 reader is not byte-identical to the 0844 reader")
    rows = value.get("rows")
    require(isinstance(rows, list) and len(rows) == 16, "profile row count changed")
    expected_keys = {"block", "leg", "mode", "parts", "repeat", "samples", "shape", "warmup"}
    require(all(set(row) == expected_keys for row in rows), "profile row fields changed")
    require([row["block"] for row in rows] == sorted(row["block"] for row in rows),
            "profile blocks are not ordered")
    require(sorted({row["block"] for row in rows}) == list(range(8)),
            "profile block set changed")
    for row in rows:
        require(row["leg"] in LEGS and row["shape"] in SHAPES
                and row["parts"] in PARTS and row["mode"] == "publish"
                and row["samples"] == 1 and row["warmup"] == 0
                and row["repeat"] in (0, 1), f"invalid profile row: {row}")
    for block in range(8):
        pair = [row for row in rows if row["block"] == block]
        require(len(pair) == 2 and {row["leg"] for row in pair} == set(LEGS),
                f"block {block} does not contain both legs")
        require(len({(row["shape"], row["parts"], row["repeat"]) for row in pair}) == 1,
                f"block {block} case differs between legs")
    for shape in SHAPES:
        for parts in PARTS:
            blocks = [row for row in rows if row["shape"] == shape and row["parts"] == parts]
            require(len(blocks) == 4 and {row["repeat"] for row in blocks} == {0, 1},
                    f"{shape}/{parts}: repeat set changed")
            first = {}
            for repeat in (0, 1):
                block = min(row["block"] for row in blocks if row["repeat"] == repeat)
                first[repeat] = next(row["leg"] for row in rows
                                     if row["block"] == block and row["shape"] == shape
                                     and row["parts"] == parts)
            require(set(first.values()) == set(LEGS),
                    f"{shape}/{parts}: before/after order is not balanced")
    bindings = value.get("legs_bindings")
    require(bindings == {
        "before": {"source": "source-before.json", "build": "build-before.json"},
        "after": {"source": "source-after.json", "build": "build-after.json"},
    }, "profile leg bindings changed")
    require(value.get("probe") == ["probe/Cargo.lock", "probe/Cargo.toml", "probe/src/main.rs"],
            "profile probe set changed")
    require(value.get("run_root") == "profile-runs", "profile run root changed")
    require(value.get("capture_manifest") == "capture-profile.json",
            "profile capture manifest changed")
    require(value.get("analysis_output") == "profile-analysis.json",
            "profile analysis output changed")
    symbols = value.get("selected_symbols")
    require(isinstance(symbols, Mapping), "selected symbol contract is missing")
    for category in ("topology_limits", "physical_lookup", "metadata", "entry_id"):
        prefixes = symbols.get(category)
        require(isinstance(prefixes, list) and prefixes and all(isinstance(x, str) and x for x in prefixes),
                f"selected {category} prefixes are missing")
    return value


def admission_index() -> tuple[dict[str, Any], dict[str, dict[str, Any]]]:
    path = P / "admission.json"
    require(path.is_file(), "admission.json is required before profile capture")
    value = read(path)
    require(value.get("status") == "pass", "admission is not passing")
    files = value.get("files")
    require(isinstance(files, list) and files, "admission file inventory is missing")
    index: dict[str, dict[str, Any]] = {}
    for item in files:
        actual = bind(item, "admission file")
        key = str(Path(actual["path"]).resolve())
        require(key not in index, f"duplicate admission path: {key}")
        index[key] = actual
    return value, index


def require_admitted(path: Path, index: Mapping[str, Mapping[str, Any]], label: str) -> dict[str, Any]:
    actual = descriptor(path)
    admitted = index.get(str(path.resolve()))
    require(admitted is not None, f"{label}: file is not in admission: {path}")
    require(dict(admitted) == actual, f"{label}: admitted descriptor changed: {path}")
    return actual


def validate_binary(binary: Mapping[str, Any], leg: str) -> dict[str, Any]:
    """Bind a binary while allowing the packet's post-cleanup witness path."""

    path = Path(binary["path"])
    if path.exists():
        actual = descriptor(path)
        require(actual == dict(binary), f"{leg}: build binary descriptor changed")
        return actual
    cleanup_path = P / "cleanup.json"
    require(cleanup_path.is_file(), f"{leg}: binary is absent without cleanup witness")
    cleanup = read(cleanup_path)
    require(cleanup.get("status") == "pass", f"{leg}: cleanup witness is not passing")
    witnesses = cleanup.get("verified_binaries")
    require(isinstance(witnesses, list) and dict(binary) in witnesses,
            f"{leg}: absent binary is not covered by cleanup witness")
    require(not d.TARGET.exists(), f"{leg}: absent binary requires complete target cleanup")
    return dict(binary)


def validate_bindings(value: Mapping[str, Any], admission: Mapping[str, Mapping[str, Any]],
                      require_current_after: bool = False) -> dict[str, Any]:
    """Bind both release legs to their existing main-matrix build freezes."""

    main = value["main_packet"]
    required_packet = [
        "profile-plan.json", "profile.py", "driver.py", "reader.py", "qualify.py",
        "plan.json", "origin.json",
    ]
    for name in required_packet:
        require_admitted(P / name, admission, f"packet binding {name}")
    for name in main.values():
        if name != "admission.json":
            require_admitted(P / name, admission, f"main packet binding {name}")
    require_admitted(P / "probe/Cargo.lock", admission, "probe lock binding")
    require_admitted(P / "probe/Cargo.toml", admission, "probe manifest binding")
    require_admitted(P / "probe/src/main.rs", admission, "probe source binding")

    result: dict[str, Any] = {}
    for leg in LEGS:
        names = value["legs_bindings"][leg]
        source_path = P / names["source"]
        build_path = P / names["build"]
        source_desc = require_admitted(source_path, admission, f"{leg} source descriptor")
        build_desc = require_admitted(build_path, admission, f"{leg} build descriptor")
        source_value = read(source_path)
        source_map = source_value.get("source")
        require(isinstance(source_map, Mapping) and source_map,
                f"{leg}: source map is missing")
        # During capture the after source is the current checkout.  The before
        # source is an earlier admitted source map and is intentionally allowed
        # to differ in the two candidate files.  After cleanup both source
        # inventories remain the immutable custody oracle and the current
        # checkout may represent the selected decision leg.
        if require_current_after and leg == "after":
            require(d.source() == dict(source_map),
                    f"{leg}: current source differs from source descriptor")
            for name, digest in source_map.items():
                require(d.sha(ROOT / name) == digest, f"{leg}: source changed: {name}")
        build_value = read(build_path)
        binary = build_value.get("binary")
        freeze_ref = build_value.get("freeze")
        require(isinstance(binary, Mapping) and isinstance(freeze_ref, Mapping),
                f"{leg}: build descriptor is incomplete")
        binary_actual = validate_binary(binary, leg)
        freeze_path = Path(freeze_ref["path"])
        require_admitted(freeze_path, admission, f"{leg} freeze")
        require(dict(descriptor(freeze_path)) == dict(freeze_ref),
                f"{leg}: freeze descriptor changed")
        freeze = read(freeze_path)
        require(freeze.get("source") == dict(source_map), f"{leg}: build/source maps differ")
        require(freeze.get("plan") == descriptor(P / "plan.json"),
                f"{leg}: build was not made from the frozen main plan")
        require(freeze.get("driver") == descriptor(P / "driver.py"),
                f"{leg}: build driver descriptor differs")
        probe = freeze.get("probe")
        require(isinstance(probe, Mapping), f"{leg}: probe descriptor is missing")
        require(set(probe) == set(value["probe"]), f"{leg}: probe file set differs")
        for name in value["probe"]:
            require(d.sha(P / name) == probe[name], f"{leg}: probe source changed: {name}")
            require_admitted(P / name, admission, f"{leg} probe {name}")
        result[leg] = {
            "source": source_desc,
            "source_value": source_value,
            "build": build_desc,
            "build_value": build_value,
            "freeze": descriptor(freeze_path),
            "freeze_value": freeze,
            "binary": binary_actual,
        }
    return result


def folder(row: Mapping[str, Any]) -> Path:
    return PROFILE_DIR / f"{row['block']:02}-{row['shape']}-{row['parts']:04}-{row['leg']}"


def profile_argv(row: Mapping[str, Any], binary: Mapping[str, Any], out: Path) -> list[str]:
    return [
        "taskset", "-c", "12", "/usr/bin/time", "-f", "%M %R %F",
        "-o", str(out / "time.txt"), "valgrind", *PROFILE_OPTIONS,
        "--callgrind-out-file=" + str(out / "profile.callgrind"),
        str(binary["path"]), "--shape", row["shape"], "--parts", str(row["parts"]),
        "--mode", "publish", "--samples", "1", "--warmup", "0",
        "--output", str(out / "report.json"),
    ]


def qualification_index(leg: str, admission: Mapping[str, Mapping[str, Any]]) -> dict[tuple[str, int], dict[str, Any]]:
    path = P / f"qualification-{leg}.json"
    require_admitted(path, admission, f"{leg} qualification summary")
    value = read(path)
    rows = value.get("rows")
    require(value.get("status") == "pass" and isinstance(rows, list),
            f"{leg} qualification summary is not passing")
    selected: dict[tuple[str, int], dict[str, Any]] = {}
    for item in rows:
        row = item.get("row")
        require(isinstance(row, Mapping) and row.get("leg") == leg,
                f"{leg} qualification row is malformed")
        if (row.get("mode"), row.get("shape"), row.get("parts")) \
                not in {("publish", shape, parts) for shape in SHAPES for parts in PARTS}:
            continue
        key = (row.get("shape"), row.get("parts"))
        require(row.get("samples") == 1 and row.get("warmup") == 0 and key not in selected,
                f"{leg} qualification rows are not unique")
        selected[key] = dict(item)
    require(set(selected) == {(shape, parts) for shape in SHAPES for parts in PARTS},
            f"{leg} qualification case set differs")
    return selected


def expected_qualification(row: Mapping[str, Any], qindex: Mapping[tuple[str, int], Mapping[str, Any]],
                           bindings: Mapping[str, Any], admission: Mapping[str, Mapping[str, Any]]) -> dict[str, Any]:
    item = qindex[(row["shape"], row["parts"])]
    qrow = item["row"]
    require(qrow["mode"] == row["mode"] and qrow["shape"] == row["shape"]
            and qrow["parts"] == row["parts"] and qrow["leg"] == row["leg"],
            f"qualification identity differs for {row}")
    qfolder = Path(item["folder"])
    require(qfolder.is_dir(), f"missing qualification folder: {qfolder}")
    qreceipt_path = qfolder / "receipt.json"
    qstarted_path = qfolder / "started.json"
    require_admitted(qreceipt_path, admission, f"qualification receipt {qfolder}")
    require_admitted(qstarted_path, admission, f"qualification start {qfolder}")
    qreceipt = read(qreceipt_path)
    qstarted = read(qstarted_path)
    require(qreceipt.get("lane") == "qualification" and qreceipt.get("exit_code") == 0,
            f"qualification command did not pass: {qfolder}")
    require(qreceipt.get("row") == dict(qrow) and qstarted.get("row") == dict(qrow),
            f"qualification row custody changed: {qfolder}")
    build_path = P / f"build-{row['leg']}.json"
    build_value = bindings[row["leg"]]["build_value"]
    require(qstarted.get("build") == descriptor(build_path),
            f"qualification build receipt changed: {qfolder}")
    require(qstarted.get("binary") == bindings[row["leg"]]["binary"],
            f"qualification binary receipt changed: {qfolder}")
    argv = [
        "taskset", "-c", "12", "/usr/bin/time", "-f", "%M %R %F",
        "-o", str(qfolder / "time.txt"), str(build_value["binary"]["path"]),
        "--shape", row["shape"], "--parts", str(row["parts"]), "--mode", "publish",
        "--samples", "1", "--warmup", "0", "--output", str(qfolder / "report.json"),
        "--artifact", str(qfolder / "artifact.zip"),
    ]
    command_ref = qreceipt.get("command")
    require(isinstance(command_ref, Mapping), f"qualification command descriptor missing: {qfolder}")
    bind(command_ref, f"qualification command {qfolder}")
    expected_command_path = P / "commands" / qfolder.name / "receipt.json"
    require(Path(command_ref["path"]).resolve() == expected_command_path.resolve(),
            f"qualification command label changed: {qfolder}")
    command = read(Path(command_ref["path"]))
    command_started = read(Path(command_ref["path"]).parent / "started.json")
    bind(command.get("log"), f"qualification command log {qfolder}")
    require({item.name for item in Path(command_ref["path"]).parent.iterdir()}
            == {"started.json", "receipt.json", "output.log"},
            f"qualification command artifact inventory changed: {qfolder}")
    require(command.get("exit_code") == 0 and command.get("argv") == argv
            and command_started.get("argv") == argv
            and command.get("started_unix") == command_started.get("started_unix"),
            f"qualification command receipt changed: {qfolder}")
    require(command.get("started_unix", 0) >= qstarted.get("started_unix", 0)
            and command.get("finished_unix", 0) <= qreceipt.get("finished_unix", 0),
            f"qualification receipt timing changed: {qfolder}")
    required = {"artifact.zip", "report.json", "source.zip", "started.json", "time.txt", "receipt.json"}
    require({p.name for p in qfolder.iterdir()} == required,
            f"qualification artifact inventory changed: {qfolder}")
    artifacts = qreceipt.get("artifacts")
    require(isinstance(artifacts, list) and [x["path"] for x in artifacts]
            == sorted(x["path"] for x in artifacts),
            f"qualification receipt artifact order changed: {qfolder}")
    require({Path(x["path"]).name for x in artifacts} == required - {"receipt.json"},
            f"qualification receipt artifact names changed: {qfolder}")
    for artifact in artifacts:
        bind(artifact, f"qualification artifact receipt {qfolder}")
    for name in required - {"receipt.json"}:
        require_admitted(qfolder / name, admission, f"qualification artifact {qfolder/name}")
    report_path = qfolder / "report.json"
    require(item.get("report") == descriptor(report_path),
            f"qualification report descriptor changed: {qfolder}")
    validated = r.validate_report(
        report_path, row["parts"], row["mode"], 1, 0,
        artifact_path=qfolder / "artifact.zip", source_path=qfolder / "source.zip",
        lane="qualification", shape=row["shape"],
    )
    return {
        "item": item,
        "folder": qfolder,
        "report_path": report_path,
        "report": read(report_path),
        "validated": validated,
    }


def validate_profile_receipt(row: Mapping[str, Any], bindings: Mapping[str, Any],
                             qualification: Mapping[str, Any], admission: Mapping[str, Mapping[str, Any]]) -> dict[str, Any]:
    out = folder(row)
    require(out.is_dir(), f"missing profile folder: {out}")
    started_path = out / "started.json"
    receipt_path = out / "receipt.json"
    started = read(started_path)
    receipt = read(receipt_path)
    require(started.get("row") == dict(row) and receipt.get("row") == dict(row),
            f"profile row custody changed: {out}")
    require(receipt.get("exit_code") == 0, f"profile command did not pass: {out}")
    binary = bindings[row["leg"]]["binary"]
    build = bindings[row["leg"]]["build"]
    require(started.get("binary") == binary and started.get("build") == build,
            f"profile binary/build binding changed: {out}")
    require(started.get("argv") == profile_argv(row, binary, out),
            f"profile start argv changed: {out}")
    require(started.get("plan") == descriptor(PLAN_PATH), f"profile plan binding changed: {out}")
    require(started.get("source") == bindings[row["leg"]]["source"],
            f"profile source binding changed: {out}")
    require(started.get("qualification") == descriptor(qualification["report_path"]),
            f"profile qualification binding changed: {out}")
    argv = profile_argv(row, binary, out)
    command_ref = receipt.get("command")
    require(isinstance(command_ref, Mapping), f"profile command descriptor missing: {out}")
    bind(command_ref, f"profile command {out}")
    expected_command_path = P / "commands" / ("profile-" + out.name) / "receipt.json"
    require(Path(command_ref["path"]).resolve() == expected_command_path.resolve(),
            f"profile command label changed: {out}")
    command_path = Path(command_ref["path"])
    command = read(command_path)
    command_started = read(command_path.parent / "started.json")
    bind(command.get("log"), f"profile command log {out}")
    require({item.name for item in command_path.parent.iterdir()}
            == {"started.json", "receipt.json", "output.log"},
            f"profile command artifact inventory changed: {out}")
    require(command.get("exit_code") == 0 and command.get("argv") == argv
            and command_started.get("argv") == argv
            and command.get("started_unix") == command_started.get("started_unix"),
            f"profile command argv changed: {out}")
    require(command.get("cwd") == str(ROOT) and command_started.get("cwd") == str(ROOT),
            f"profile command cwd changed: {out}")
    require(command.get("started_unix", 0) >= started.get("started_unix", 0)
            and command.get("finished_unix", 0) <= receipt.get("finished_unix", 0),
            f"profile command timing custody changed: {out}")
    expected_artifacts = {
        "profile.callgrind", "profile.callgrind.1", "report.json", "started.json", "time.txt", "receipt.json"
    }
    require({p.name for p in out.iterdir()} == expected_artifacts,
            f"profile artifact inventory changed: {out}")
    artifacts = receipt.get("artifacts")
    require(isinstance(artifacts, list) and [x["path"] for x in artifacts]
            == sorted(x["path"] for x in artifacts), f"profile artifact order changed: {out}")
    require({Path(x["path"]).name for x in artifacts} == expected_artifacts - {"receipt.json"},
            f"profile receipt artifact names changed: {out}")
    for item in artifacts:
        bind(item, f"profile artifact {out}")
    counts = (out / "time.txt").read_text(encoding="utf-8").split()
    require(len(counts) == 3 and all(token.isdigit() for token in counts)
            and int(counts[0]) > 0, f"profile time receipt malformed: {out}")
    report = r.validate_report(
        out / "report.json", row["parts"], row["mode"], 1, 0,
        expected=qualification["report"], lane="profiles", shape=row["shape"],
    )
    profile = r.validate_profile(
        out / "profile.callgrind.1", out / "profile.callgrind",
        owner=PROFILE_OWNER, command=argv,
    )
    return {
        "row": dict(row),
        "folder": str(out),
        "report": report,
        "profile": profile,
        "command": descriptor(command_path),
        "receipt": descriptor(receipt_path),
    }


def command_receipts() -> list[tuple[Path, dict[str, Any]]]:
    values = []
    for path in sorted((P / "commands").glob("*/receipt.json")):
        values.append((path, read(path)))
    return values


def overlaps(left: Mapping[str, Any], right: Mapping[str, Any]) -> bool:
    return left.get("started_unix", 0) < right.get("finished_unix", 0) \
        and right.get("started_unix", 0) < left.get("finished_unix", 0)


def check_no_overlap(rows: Sequence[Mapping[str, Any]]) -> None:
    profile_intervals = []
    for row in rows:
        command = read(Path(row["command"]["path"]))
        profile_intervals.append((row["row"], command))
    ordered = sorted(profile_intervals, key=lambda item: item[1]["started_unix"])
    require(all(not overlaps(a[1], b[1]) for a, b in zip(ordered, ordered[1:])),
            "profile commands overlap")
    profile_paths = {str(Path(row["command"]["path"]).resolve()) for row in rows}
    for path, command in command_receipts():
        if str(path.resolve()) in profile_paths:
            continue
        argv = [str(item) for item in command.get("argv", [])]
        is_cargo = any(Path(item).name == "cargo" for item in argv)
        is_heap = "heaptrack" in [Path(item).name for item in argv]
        is_probe = any("opc-preservation-probe" in Path(item).name for item in argv)
        if not (is_cargo or is_heap or is_probe):
            continue
        for _, profile_command in profile_intervals:
            require(not overlaps(profile_command, command),
                    f"profile overlaps main workload/tool command: {path}")


def compact_functions(parsed: Mapping[str, Any], prefixes: Sequence[str]) -> list[dict[str, Any]]:
    def matches(name: str) -> bool:
        return any(name == prefix or name.startswith(prefix + "<") or name.startswith(prefix + "::")
                   for prefix in prefixes)

    rows = []
    for function in parsed["functions"].values():
        name = function.get("name", "")
        if not matches(name):
            continue
        edges = [dict(edge) for edge in function.get("edges", [])]
        edges.sort(key=lambda edge: (
            edge.get("caller_id", 0), edge.get("callee_id", 0), edge.get("line", 0),
            edge.get("position", ""), edge.get("calls", 0), edge.get("inclusive_ir", 0),
        ))
        rows.append({
            "id": function["id"],
            "name": name,
            "file": function.get("file", ""),
            "records": function.get("records", 0),
            "self_ir": function.get("self_ir", 0),
            "edges": edges,
        })
    return sorted(rows, key=lambda row: (row["name"], row["id"], row["file"]))


def compact_profile(raw_path: Path, terminal_path: Path, profile: Mapping[str, Any],
                    value: Mapping[str, Any], selected_symbols: Mapping[str, Sequence[str]]) -> dict[str, Any]:
    # validate_profile has already checked the entire file, including source
    # attribution, all function self costs, the sole owner call and the empty
    # terminal dump.  Parsing again here is intentional: only after those
    # checks do we discard unrelated functions.
    parsed = r.parse_raw(raw_path)
    require(parsed["sha256"] == profile["sha256"], "raw Callgrind hash changed during reduction")
    selected = {}
    for category, prefixes in selected_symbols.items():
        functions = compact_functions(parsed, prefixes)
        selected[category] = {
            "prefixes": list(prefixes),
            "functions": functions,
            "function_count": len(functions),
            "self_ir": sum(item["self_ir"] for item in functions),
            "edge_count": sum(len(item["edges"]) for item in functions),
            "edge_ir": sum(edge["inclusive_ir"] for item in functions for edge in item["edges"]),
        }
    owner = profile["owner"]
    return {
        "raw_callgrind": {
            "numbered": descriptor(raw_path),
            "terminal": descriptor(terminal_path),
        },
        "summary_ir": profile["summary_ir"],
        "owner": {
            "id": owner["id"],
            "name": owner["name"],
            "file": owner["file"],
            "self_ir": owner["self_ir"],
            "direct_children_ir": owner["direct_children_ir"],
            "inclusive_ir": profile["summary_ir"],
            "incoming": [dict(owner["incoming"])],
        },
        "selected": selected,
        "validation": dict(profile["validation"]),
        "report": {
            "path": value["report"]["report"]["path"],
            "sha256": value["report"]["report"]["sha256"],
            "source": value["report"]["source"],
            "output": value["report"]["output"],
            "logical": value["report"]["logical"],
        },
    }


def selected_totals(compact: Mapping[str, Any], category: str) -> dict[str, int]:
    value = compact["selected"][category]
    return {
        "function_count": value["function_count"],
        "self_ir": value["self_ir"],
        "edge_count": value["edge_count"],
        "edge_ir": value["edge_ir"],
    }


def compare_profiles(rows: Sequence[Mapping[str, Any]], value: Mapping[str, Any]) -> list[dict[str, Any]]:
    groups: dict[tuple[str, int, int], dict[str, Mapping[str, Any]]] = {}
    for item in rows:
        row = item["row"]
        key = (row["shape"], row["parts"], row["repeat"])
        groups.setdefault(key, {})[row["leg"]] = item
    result = []
    categories = tuple(value["selected_symbols"])
    for shape in SHAPES:
        for parts in PARTS:
            for repeat in range(2):
                group = groups[(shape, parts, repeat)]
                require(set(group) == set(LEGS), f"missing profile leg for {shape}/{parts}/{repeat}")
                before = group["before"]["compact"]
                after = group["after"]["compact"]
                selected = {}
                for category in categories:
                    b = selected_totals(before, category)
                    a = selected_totals(after, category)
                    selected[category] = {
                        "before": b,
                        "after": a,
                        "delta_self_ir": a["self_ir"] - b["self_ir"],
                        "delta_edge_ir": a["edge_ir"] - b["edge_ir"],
                    }
                result.append({
                    "shape": shape,
                    "parts": parts,
                    "repeat": repeat,
                    "owner": {
                        "before_summary_ir": before["summary_ir"],
                        "after_summary_ir": after["summary_ir"],
                        "before_self_ir": before["owner"]["self_ir"],
                        "after_self_ir": after["owner"]["self_ir"],
                        "delta_summary_ir": after["summary_ir"] - before["summary_ir"],
                    },
                    "selected": selected,
                })
    return result


def validate_all(value: Mapping[str, Any], admission: Mapping[str, Mapping[str, Any]]) -> dict[str, Any]:
    bindings = validate_bindings(value, admission)
    qindices = {leg: qualification_index(leg, admission) for leg in LEGS}
    qcache: dict[tuple[str, str, int], dict[str, Any]] = {}
    rows = []
    for row in value["rows"]:
        key = (row["leg"], row["shape"], row["parts"])
        if key not in qcache:
            qcache[key] = expected_qualification(row, qindices[row["leg"]], bindings, admission)
        validated = validate_profile_receipt(row, bindings, qcache[key], admission)
        compact = compact_profile(
            P / validated["folder"] / "profile.callgrind.1",
            P / validated["folder"] / "profile.callgrind",
            validated["profile"],
            validated,
            value["selected_symbols"],
        )
        rows.append({
            "row": dict(row),
            "folder": validated["folder"],
            "receipt": validated["receipt"],
            "command": validated["command"],
            "compact": compact,
        })
    require({item["folder"] for item in rows}
            == {str(folder(row)) for row in value["rows"]},
            "profile folder inventory differs from frozen plan")
    require({p.name for p in PROFILE_DIR.iterdir() if p.is_dir()}
            == {folder(row).name for row in value["rows"]},
            "profile run directory inventory contains an unexpected directory")
    require({p.name for p in PROFILE_DIR.iterdir() if p.is_file()} == set(),
            "profile run root contains an unexpected file")
    expected_command_labels = {"profile-" + folder(row).name for row in value["rows"]}
    actual_command_labels = {
        path.parent.name for path in (P / "commands").glob("profile-*/receipt.json")
    }
    require(actual_command_labels == expected_command_labels,
            "profile command receipt inventory differs from frozen plan")
    ordered = sorted(rows, key=lambda item: read(Path(item["command"]["path"]))["started_unix"])
    require([item["row"] for item in ordered] == value["rows"],
            "profile receipt order differs from frozen plan")
    check_no_overlap(rows)
    capture = read(CAPTURE_PATH)
    require(capture.get("status") == "pass", "profile capture manifest is not passing")
    require(capture.get("plan") == descriptor(PLAN_PATH), "capture plan descriptor changed")
    require(capture.get("runner") == descriptor(P / "profile.py"), "capture runner descriptor changed")
    require(capture.get("bindings") == {
        leg: {
            "source": bindings[leg]["source"],
            "build": bindings[leg]["build"],
            "freeze": bindings[leg]["freeze"],
            "binary": bindings[leg]["binary"],
        }
        for leg in LEGS
    }, "capture leg bindings changed")
    receipts = capture.get("receipts")
    expected_receipts = [descriptor(P / "profile-runs" / folder(row).name / "receipt.json")
                         for row in value["rows"]]
    require(receipts == expected_receipts, "capture receipt inventory/order changed")
    return {
        "bindings": {
            leg: {
                "source": bindings[leg]["source"],
                "build": bindings[leg]["build"],
                "freeze": bindings[leg]["freeze"],
                "binary": bindings[leg]["binary"],
            }
            for leg in LEGS
        },
        "rows": rows,
        "comparisons": compare_profiles(rows, value),
    }


def capture() -> None:
    value = plan()
    require(not CAPTURE_PATH.exists(), "capture manifest already exists; refusing overwrite")
    require(not PROFILE_DIR.exists(), "profile-runs already exists; refusing overwrite")
    _, admission = admission_index()
    bindings = validate_bindings(value, admission, require_current_after=True)
    qindices = {leg: qualification_index(leg, admission) for leg in LEGS}
    qcache: dict[tuple[str, str, int], dict[str, Any]] = {}
    PROFILE_DIR.mkdir()
    receipts = []
    for row in value["rows"]:
        out = folder(row)
        out.mkdir()
        if (row["leg"], row["shape"], row["parts"]) not in qcache:
            key = (row["leg"], row["shape"], row["parts"])
            qcache[key] = expected_qualification(row, qindices[row["leg"]], bindings, admission)
        qualification = qcache[(row["leg"], row["shape"], row["parts"])]
        binary = bindings[row["leg"]]["binary"]
        argv = profile_argv(row, binary, out)
        d.write(out / "started.json", {
            "row": dict(row),
            "argv": argv,
            "binary": binary,
            "build": bindings[row["leg"]]["build"],
            "plan": descriptor(PLAN_PATH),
            "qualification": descriptor(qualification["report_path"]),
            "source": bindings[row["leg"]]["source"],
            "started_unix": time.time(),
        })
        label = "profile-" + out.name
        code = d.run(label, argv)
        require(code == 0, f"profile command failed: {label}")
        # Run the full independent readers before the receipt is sealed.  The
        # same checks are replayed by analyze() after all commands finish.
        r.validate_report(out / "report.json", row["parts"], row["mode"], 1, 0,
                          expected=qualification["report"], lane="profiles", shape=row["shape"])
        r.validate_profile(out / "profile.callgrind.1", out / "profile.callgrind",
                           owner=PROFILE_OWNER, command=argv)
        artifacts = [descriptor(path) for path in sorted(out.rglob("*"))
                     if path.is_file() and path.name != "receipt.json"]
        d.write(out / "receipt.json", {
            "row": dict(row),
            "exit_code": code,
            "command": descriptor(P / "commands" / label / "receipt.json"),
            "artifacts": artifacts,
            "finished_unix": time.time(),
        })
        receipts.append(descriptor(out / "receipt.json"))
        print(label, "captured", flush=True)
    d.write(CAPTURE_PATH, {
        "status": "pass",
        "scope": "diagnostic-only-callgrind; separate from native and Heaptrack populations",
        "plan": descriptor(PLAN_PATH),
        "runner": descriptor(P / "profile.py"),
        "receipts": receipts,
        "bindings": {
            leg: {
                "source": bindings[leg]["source"],
                "build": bindings[leg]["build"],
                "freeze": bindings[leg]["freeze"],
                "binary": bindings[leg]["binary"],
            }
            for leg in LEGS
        },
    })


def analyze() -> dict[str, Any]:
    """Replay and compact the complete retained profile population."""

    value = plan()
    admission_value, admission = admission_index()
    validated = validate_all(value, admission)
    result = {
        "status": "pass",
        "disposition": "diagnostic-only-no-native-pooling",
        "scope": "Callgrind Ir counts inside profile_publish; no cycle, elapsed-phase, or native timing claim",
        "reports": len(validated["rows"]),
        "samples": len(validated["rows"]),
        "plan": descriptor(PLAN_PATH),
        "runner": descriptor(P / "profile.py"),
        "reader": descriptor(P / "reader.py"),
        "admission": descriptor(P / "admission.json"),
        "admission_status": admission_value["status"],
        "rows": validated["rows"],
        "comparisons": validated["comparisons"],
        "selected_symbols": value["selected_symbols"],
        "limits": value["limits"],
    }
    if ANALYSIS_PATH.exists():
        require(read(ANALYSIS_PATH) == result, "profile-analysis.json already exists with different content")
    else:
        d.write(ANALYSIS_PATH, result)
    print("Profile analysis PASS", result["reports"], "reports;", result["samples"], "samples", flush=True)
    return result


if __name__ == "__main__":
    require(len(sys.argv) == 2, "usage: profile.py capture|analyze")
    if sys.argv[1] == "capture":
        capture()
    elif sys.argv[1] == "analyze":
        analyze()
    else:
        fail(f"unknown mode: {sys.argv[1]}")
