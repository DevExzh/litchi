#!/usr/bin/env python3
"""Offline replay of the 0789 RSS accounting calibration."""
from __future__ import annotations
import gzip
import hashlib
import json
import re
import sys
from pathlib import Path
from typing import Any
PACKET = Path(__file__).resolve().parent
PLAN = PACKET / "plan.json"
HOST = PACKET / "host.json"
MATRIX_RECEIPTS = PACKET / "matrix" / "receipts.json"
TRACE_RECEIPTS = PACKET / "traces" / "receipts.json"
BUILD = PACKET / "build.json"
CLEANUP = PACKET / "cleanup.json"
OUT_JSON = PACKET / "analysis.json"
OUT_MD = PACKET / "analysis.md"
PLAN_SCHEMA = "litchi.rss-accounting-calibration.0789.v1"
RESULT_SCHEMA = "litchi.rss-accounting-analysis.0789.v1"
CHILD_SCHEMA = "litchi.rss-accounting-child.0789.v1"
PHASES = ("startup", "mapped", "touched", "workers_joined", "unmapped", "final")
MIBS = (0, 1, 4, 16, 64)
LAUNCHERS = ("direct", "time")
AFFINITIES = ("single", "all")
OBSERVERS = ("full", "identity")
CASES = [{"mib": mib, "touch": touch, "workers": workers}
         for mib in MIBS for workers, touch in
         ((0, "main"), (4, "main"), (4, "workers"))]
CASES += [{"mib": 4, "touch": "main", "workers": 32},
          {"mib": 4, "touch": "workers", "workers": 32}]
FULL_RAW = ("stat", "smaps", "smaps_rollup", "maps", "status")
IDENTITY_RAW = ("stat",)
HEADER = re.compile(r"^([0-9a-f]+)-([0-9a-f]+)\s+([-rwxps]{4})\s+"
                    r"([0-9a-f]+)\s+([0-9a-f]+:[0-9a-f]+)\s+(\d+)"
                    r"(?:\s+(.*?))?\s*$")
KV = re.compile(r"^([A-Za-z][A-Za-z0-9_]*):\s+([0-9]+)(?:\s+(kB))?\s*$")
class ReplayError(RuntimeError):
    pass
def fail(message: str) -> None:
    raise ReplayError(message)
def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)
def read_json(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing JSON: {path}")
    try:
        return json.loads(path.read_text())
    except (OSError, ValueError) as error:
        fail(f"invalid JSON {path}: {error}")
def digest(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1 << 20), b""):
            h.update(block)
    return h.hexdigest()
def integer(value: Any, label: str, positive: bool = False) -> int:
    require(isinstance(value, int) and not isinstance(value, bool),
            f"{label} is not an integer")
    require(value > 0 if positive else value >= 0,
            f"{label} is outside its allowed range")
    return value
def artifact(value: Any, label: str, allow_witness: bool = False) -> tuple[Path | None, dict[str, Any]]:
    require(isinstance(value, dict), f"{label} is not an artifact")
    raw = value.get("path")
    size = value.get("bytes")
    expected = value.get("sha256")
    require(isinstance(raw, str) and raw, f"{label}.path is missing")
    integer(size, f"{label}.bytes")
    require(isinstance(expected, str) and re.fullmatch(r"[0-9a-f]{64}", expected),
            f"{label}.sha256 is invalid")
    path = Path(raw)
    if not path.is_absolute():
        path = PACKET / path
    path = path.resolve(strict=False)
    if not path.is_file():
        if not allow_witness:
            fail(f"missing {label}: {raw}")
        require(CLEANUP.is_file(), f"missing {label} and cleanup witness")
        witness = read_json(CLEANUP)
        require(isinstance(witness, dict) and witness.get("target_removed") is True and
                witness.get("target") == str(Path(raw).parent),
                "cleanup witness does not prove target removal")
        removed = witness.get("removed_binaries")
        require(isinstance(removed, list), "cleanup binary witness is missing")
        matches = [item for item in removed if isinstance(item, dict) and
                   item.get("path") == raw and item.get("bytes") == size and
                   item.get("sha256") == expected]
        require(len(matches) == 1, f"cleanup witness does not contain unique {label}")
        return None, {"path": raw, "bytes": size, "sha256": expected}
    require(not path.is_symlink(), f"{label} is a symlink: {raw}")
    require(path.stat().st_size == size, f"{label}.bytes changed")
    require(digest(path) == expected, f"{label}.sha256 changed")
    return path, {"path": raw, "bytes": size, "sha256": expected}
def artifact_text(value: Any, label: str) -> tuple[str, dict[str, Any]]:
    path, receipt = artifact(value, label)
    try:
        return path.read_text(), receipt
    except OSError as error:
        fail(f"cannot read {label}: {error}")
def artifact_json(value: Any, label: str) -> tuple[Any, dict[str, Any]]:
    path, receipt = artifact(value, label)
    try:
        raw = path.read_bytes()
        if raw[:2] == b"\x1f\x8b":
            raw = gzip.decompress(raw)
        return json.loads(raw), receipt
    except (OSError, ValueError) as error:
        fail(f"invalid JSON artifact {label}: {error}")
def load_plan() -> dict[str, Any]:
    plan = read_json(PLAN)
    require(isinstance(plan, dict) and plan.get("schema") == PLAN_SCHEMA,
            "plan schema changed")
    require(plan.get("phases") == list(PHASES), "phase order changed")
    require(plan.get("repeats") == 2, "repeat count changed")
    require(plan.get("cases") == CASES, "case matrix changed")
    require(plan.get("expected") == {"checkpoints": 1656, "identity_snapshots": 816,
        "matrix_children": 272, "matrix_full_snapshots": 816,
        "trace_children": 4, "trace_full_snapshots": 24}, "expected counts changed")
    require(plan.get("snapshot_full") == ["smaps", "smaps_rollup", "maps",
                                            "status", "stat", "task_ids", "exe"],
            "full snapshot contract changed")
    require(plan.get("snapshot_identity") == ["stat", "task_ids", "exe"],
            "identity snapshot contract changed")
    host = read_json(HOST)
    require(isinstance(host, dict) and host.get("page_size") == 4096,
            "host page size changed")
    require(plan.get("ack") == "+\n", "ACK bytes changed")
    return plan
def load_receipts() -> list[dict[str, Any]]:
    expanded = []
    for lane, path in (("matrix", MATRIX_RECEIPTS), ("traces", TRACE_RECEIPTS)):
        rows = read_json(path)
        require(isinstance(rows, list), f"{lane} receipt list is missing")
        for index, row in enumerate(rows):
            report_value, _ = artifact_json(row, f"{lane} receipt {index}")
            expanded.append(report_value)
    return expanded
def case_of(row: dict[str, Any], label: str) -> dict[str, Any]:
    value = row.get("case")
    require(isinstance(value, dict), f"{label}.case is missing")
    case = {key: value.get(key) for key in ("mib", "touch", "workers")}
    require(case in CASES,
            f"{label}.case is outside frozen matrix")
    return case
def row_key(row: dict[str, Any], trace: bool = False) -> tuple[Any, ...]:
    case = case_of(row, "receipt")
    return (case["mib"], case["touch"], case["workers"], row.get("repeat"),
            row.get("launcher"), row.get("affinity"), row.get("observer"), trace)
def parse_probe(text: str, label: str) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    rss: list[dict[str, Any]] = []
    ack: list[dict[str, Any]] = []
    for line_no, line in enumerate(text.splitlines(), 1):
        fields = line.split("\t")
        if not fields or not line:
            continue
        if fields[0] == "RSS0789":
            require(len(fields) == 8, f"{label}:{line_no} RSS field count changed")
            require(fields[1] in PHASES, f"{label}:{line_no} RSS phase invalid")
            values = [integer(int(item), f"{label}:{line_no}") for item in fields[2:7]]
            rss.append({"phase": fields[1], **dict(zip(("pid", "maxrss_kib", "minor_faults", "major_faults", "mapped_bytes"), values)),
                        "checksum": integer(int(fields[7]), f"{label}:{line_no}")})
        elif fields[0] == "ACK0789":
            require(len(fields) == 6, f"{label}:{line_no} ACK field count changed")
            require(fields[1] in PHASES, f"{label}:{line_no} ACK phase invalid")
            values = [integer(int(item), f"{label}:{line_no}") for item in fields[2:6]]
            ack.append({"phase": fields[1], **dict(zip(("pid", "maxrss_kib", "minor_faults", "major_faults"), values))})
        else:
            fail(f"{label}:{line_no} contains an unexpected line")
    require([item["phase"] for item in rss] == list(PHASES),
            f"{label} RSS phase sequence changed")
    require([item["phase"] for item in ack] == list(PHASES),
            f"{label} ACK phase sequence changed")
    return rss, ack
def parse_stat(raw: str, label: str) -> dict[str, int]:
    match = re.match(r"^(\d+)\s+\((.*)\)\s+(.*)$", raw.strip())
    require(match is not None, f"{label} stat is malformed")
    rest = match.group(3).split()
    require(len(rest) >= 20, f"{label} stat is truncated")
    try:
        return {"pid": int(match.group(1)), "ppid": int(rest[1]),
                "starttime": int(rest[19]), "threads": int(rest[17])}
    except ValueError as error:
        fail(f"{label} stat has invalid numeric fields: {error}")
def identity_from_point(point: dict[str, Any], label: str) -> dict[str, Any]:
    ident = point.get("identity")
    require(isinstance(ident, dict), f"{label}.identity is missing")
    result = {}
    for key in ("pid", "ppid", "starttime"):
        result[key] = integer(ident.get(key), f"{label}.identity.{key}", True)
    require(isinstance(ident.get("exe"), str) and ident["exe"],
            f"{label}.identity.exe is missing")
    result["exe"] = ident["exe"]
    tasks = ident.get("task_ids")
    require(isinstance(tasks, list) and tasks, f"{label}.identity.task_ids missing")
    result["task_ids"] = [integer(item, f"{label}.task_ids", True) for item in tasks]
    require(len(set(result["task_ids"])) == len(result["task_ids"]),
            f"{label}.task_ids contains duplicates")
    require(result["pid"] in result["task_ids"], f"{label}.task_ids lacks pid")
    return result
def parse_kv(raw: str, label: str) -> dict[str, int]:
    result: dict[str, int] = {}
    for line_no, line in enumerate(raw.splitlines(), 1):
        match = KV.match(line.strip())
        if match:
            result[match.group(1)] = int(match.group(2))
    require(result, f"{label} has no numeric fields")
    return result
def parse_smaps(raw: str, label: str) -> tuple[int, int, int]:
    rss = 0
    anon_virtual = 0
    anon_rss = 0
    current_anon = False
    current_size = 0
    current_anonymous = 0
    found = False
    for line_no, line in enumerate(raw.splitlines(), 1):
        header = HEADER.match(line)
        if header:
            if found and current_anon:
                anon_virtual += current_size
            if found:
                anon_rss += current_anonymous
            start, end = int(header.group(1), 16), int(header.group(2), 16)
            require(end > start, f"{label}:{line_no} empty mapping")
            path = (header.group(7) or "").strip()
            current_anon = not path or path.startswith("[anon")
            current_size = end - start
            current_anonymous = 0
            found = True
            continue
        match = KV.match(line.strip())
        if match and found:
            if match.group(1) == "Rss":
                rss += int(match.group(2))
            elif match.group(1) == "Anonymous":
                current_anonymous = int(match.group(2))
            elif match.group(1) == "Size":
                current_size = int(match.group(2)) * 1024
    require(found, f"{label} has no mappings")
    if current_anon:
        anon_virtual += current_size
    if found:
        anon_rss += current_anonymous
    return rss, anon_virtual, anon_rss
def expected_checksum(mib: int, page_size: int) -> int:
    bytes_count = mib * 1024 * 1024
    pages = (bytes_count + page_size - 1) // page_size
    full, remainder = divmod(pages, 251)
    return full * sum(range(1, 252)) + sum(range(1, remainder + 1))
def parse_snapshots(row: dict[str, Any], points: list[dict[str, Any]],
                    observer: str, page_size: int) -> list[dict[str, Any]]:
    require(len(points) == len(PHASES), f"receipt {row['index']} checkpoint count changed")
    output = []
    stable: tuple[int, int, int, str] | None = None
    for index, (point, phase) in enumerate(zip(points, PHASES)):
        label = f"receipt {row['index']} {phase}"
        require(point.get("phase") == phase, f"{label} phase changed")
        identity = identity_from_point(point, label)
        raw = point.get("raw")
        require(isinstance(raw, dict), f"{label}.raw is missing")
        require(set(raw) == set(FULL_RAW if observer == "full" else IDENTITY_RAW),
                f"{label}.raw file set changed")
        if "stat" in raw:
            stat = parse_stat(raw["stat"], f"{label}.stat")
            require(stat["pid"] == identity["pid"] and stat["ppid"] == identity["ppid"] and
                    stat["starttime"] == identity["starttime"],
                    f"{label} stat identity changed")
            require(stat["threads"] == 1 and identity["task_ids"] == [identity["pid"]],
                    f"{label} captured before workers joined")
        key = (identity["pid"], identity["ppid"], identity["starttime"], identity["exe"])
        if stable is None:
            stable = key
        require(key == stable, f"{label} process identity changed")
        rss = None
        anon_virtual = None
        anon_rss = None
        status_rss = None
        rollup_rss = None
        if observer == "full":
            smaps_rss, anon_virtual, anon_rss = parse_smaps(raw["smaps"], f"{label}.smaps")
            rollup = parse_kv(raw["smaps_rollup"], f"{label}.smaps_rollup")
            require("Rss" in rollup, f"{label}.smaps_rollup lacks Rss")
            rollup_rss = rollup["Rss"]
            status = parse_kv(raw["status"], f"{label}.status")
            require("Pid" in status and "PPid" in status and "Threads" in status and
                    "VmRSS" in status and "VmHWM" in status,
                    f"{label}.status lacks VmRSS/VmHWM")
            require(status["Pid"] == identity["pid"] and status["PPid"] == identity["ppid"] and
                    status["Threads"] == 1, f"{label}.status identity changed")
            status_rss = status["VmRSS"]
            maps = [line for line in raw["maps"].splitlines() if line.strip()]
            require(maps and all(HEADER.match(line) for line in maps),
                    f"{label}.maps is malformed")
            rss = {"smaps_sum_kib": smaps_rss, "rollup_kib": rollup_rss,
                   "status_vmrss_kib": status_rss, "status_vmhwm_kib": status["VmHWM"],
                   "smaps_minus_rollup_kib": smaps_rss - rollup_rss,
                   "status_minus_rollup_kib": status_rss - rollup_rss,
                   "anon_virtual_bytes": anon_virtual, "anon_rss_kib": anon_rss}
        output.append({"phase": phase, "identity": identity,
                       "before": point.get("before"), "after": point.get("after"),
                       "rss": rss})
    return output
def read_usage(value: Any, label: str) -> dict[str, Any]:
    if isinstance(value, dict) and "maxrss_kib" in value:
        result = {key: integer(value.get(key), f"{label}.{key}")
                  for key in ("maxrss_kib", "minor_faults", "major_faults")}
        return result
    text, _ = artifact_text(value, label)
    fields = text.strip().split()
    require(len(fields) == 3, f"{label} is not %M %R %F")
    return {"maxrss_kib": integer(int(fields[0]), label),
            "minor_faults": integer(int(fields[1]), label),
            "major_faults": integer(int(fields[2]), label)}
def verify_source_custody() -> dict[str, Any]:
    build = read_json(BUILD)
    require(isinstance(build, dict), "build.json is malformed")
    _, binary_receipt = artifact(build.get("binary"), "build.binary", allow_witness=True)
    inputs_value, _ = artifact_json(build.get("inputs"), "build.inputs")
    require(isinstance(inputs_value, dict), "build.inputs is malformed")
    for name, expected in inputs_value.items():
        path = PACKET / name
        require(path.is_file() and not path.is_symlink() and digest(path) == expected,
                f"frozen input hash changed: {name}")
    source = PACKET / "probe.c"
    require(source.is_file() and source.suffix == ".c" and
            digest(source) == inputs_value.get("probe.c"),
            "probe.c source hash changed")
    source_receipt = {"path": "probe.c", "bytes": source.stat().st_size,
                      "sha256": digest(source)}
    return {"binary": binary_receipt, "source": source_receipt}
def verify_row(row: dict[str, Any], source: dict[str, Any], plan: dict[str, Any]) -> dict[str, Any]:
    require(isinstance(row, dict), "child receipt is not an object")
    index = integer(row.get("index"), "child index")
    require(row.get("schema") == CHILD_SCHEMA, f"receipt {index} schema changed")
    case = case_of(row, f"receipt {index}")
    launcher, affinity, observer = row.get("launcher"), row.get("affinity"), row.get("observer")
    require(launcher in LAUNCHERS and affinity in AFFINITIES and observer in OBSERVERS,
            f"receipt {index} launcher/affinity/observer changed")
    repeat = integer(row.get("repeat"), f"receipt {index}.repeat")
    require(repeat < 2, f"receipt {index}.repeat outside range")
    _, binary_receipt = artifact(row.get("binary"), f"receipt {index}.binary", allow_witness=True)
    require(binary_receipt == source["binary"], f"receipt {index} binary hash changed")
    transcript, transcript_receipt = artifact_text(row.get("transcript"), f"receipt {index}.transcript")
    rss_lines, ack_lines = parse_probe(transcript, f"receipt {index}.transcript")
    wait4 = row.get("wait4")
    require(isinstance(wait4, dict) and wait4.get("exit_code") == 0,
            f"receipt {index} wait4 did not succeed")
    wait_ru = read_usage(wait4.get("rusage"), f"receipt {index}.wait4.rusage")
    if launcher == "time":
        require("gnu_time" in row, f"receipt {index} GNU time receipt missing")
        gnu = read_usage(row["gnu_time"], f"receipt {index}.gnu_time")
    else:
        require("gnu_time" not in row, f"receipt {index} direct child has GNU time")
        gnu_receipt, gnu = None, None
    points_value, points_receipt = artifact_json(row.get("points"), f"receipt {index}.points")
    require(isinstance(points_value, list), f"receipt {index}.points is not a list")
    page_size = int(plan.get("page_size", 4096))
    snapshots = parse_snapshots({"index": index, "case": case}, points_value,
                                observer, page_size)
    require(len(rss_lines) == len(ack_lines) == len(snapshots) == len(PHASES),
            f"receipt {index} protocol cardinality changed")
    checksum = expected_checksum(case["mib"], page_size)
    mapped_bytes = case["mib"] * 1024 * 1024
    for phase, before, after, snapshot in zip(PHASES, rss_lines, ack_lines, snapshots):
        require(before["pid"] == after["pid"] == snapshot["identity"]["pid"],
                f"receipt {index} {phase} pid changed")
        for key in ("maxrss_kib", "minor_faults", "major_faults"):
            require(after[key] >= before[key], f"receipt {index} {phase} {key} decreased")
        expected_bytes = mapped_bytes if phase in ("mapped", "touched", "workers_joined") else 0
        require(before["mapped_bytes"] == expected_bytes,
                f"receipt {index} {phase} mapped byte witness changed")
        expected_before = 0 if phase in ("startup", "mapped") or checksum == 0 else checksum
        require(before["checksum"] == expected_before,
                f"receipt {index} {phase} checksum phase changed")
        if phase not in ("startup", "mapped") and checksum != 0:
            require(before["checksum"] == checksum,
                    f"receipt {index} {phase} checksum changed")
        require(before["pid"] > 0, f"receipt {index} {phase} pid is invalid")
        point_before = snapshot.get("before")
        point_after = snapshot.get("after")
        require(isinstance(point_before, dict) and isinstance(point_after, dict),
                f"receipt {index} {phase} point rusage is missing")
        pre = ("maxrss_kib", "minor_faults", "major_faults", "mapped_bytes", "checksum")
        post = ("maxrss_kib", "minor_faults", "major_faults")
        require(all(point_before.get(key) == before[key] for key in pre),
                f"receipt {index} {phase} pre-ACK rusage is not aligned")
        require(all(point_after.get(key) == after[key] for key in post),
                f"receipt {index} {phase} post-ACK rusage is not aligned")
    identity = snapshots[0]["identity"]
    require(identity["pid"] == rss_lines[0]["pid"], f"receipt {index} PID mismatch")
    wait4_scope = ("strace-wrapper" if row.get("trace") else
                   "gnu-time-wrapper" if launcher == "time" else "child")
    return {"index": index, "case": case, "repeat": repeat, "launcher": launcher,
            "affinity": affinity, "observer": observer,
            "trace": bool(row.get("trace", False)), "pid": identity["pid"],
            "starttime": identity["starttime"], "wait4": wait_ru,
            "wait4_scope": wait4_scope,
            "gnu_time": gnu, "transcript": transcript_receipt,
            "points": points_receipt, "snapshots": snapshots,
            "rusage_deltas": [{"phase": phase,
                "maxrss_kib": after["maxrss_kib"] - before["maxrss_kib"],
                "minor_faults": after["minor_faults"] - before["minor_faults"],
                "major_faults": after["major_faults"] - before["major_faults"]}
                for phase, before, after in zip(PHASES, rss_lines, ack_lines)],
            "high_water": {"self_before_ack_maxrss_kib": max(x["maxrss_kib"] for x in rss_lines),
                           "self_after_ack_maxrss_kib": max(x["maxrss_kib"] for x in ack_lines),
                           "wait4_maxrss_kib": wait_ru["maxrss_kib"],
                           "gnu_time_maxrss_kib": None if gnu is None else gnu["maxrss_kib"],
                           "gnu_time_minus_wait4": None if gnu is None else {
                               key: gnu[key] - wait_ru[key]
                               for key in ("maxrss_kib", "minor_faults", "major_faults")}}}
def expected_keys(plan: dict[str, Any]) -> tuple[set[tuple[Any, ...]], set[tuple[Any, ...]]]:
    matrix = set()
    for case in plan["cases"]:
        for repeat in range(2):
            for launcher in LAUNCHERS:
                for affinity in AFFINITIES:
                    for observer in OBSERVERS:
                        matrix.add((case["mib"], case["touch"], case["workers"], repeat,
                                    launcher, affinity, observer, False))
    traces = {tuple(item[k] for k in ("mib", "touch", "workers")) +
              (0, "time", item["affinity"], "full", True)
              for item in plan["traces"]}
    return matrix, traces
def summarize(rows: list[dict[str, Any]]) -> dict[str, Any]:
    full = [row for row in rows if row["observer"] == "full"]
    identity = [row for row in rows if row["observer"] == "identity"]
    growth = []
    differences = []
    for row in full:
        points = row["snapshots"]
        mapped, touched = points[1], points[2]
        growth.append({"index": row["index"], "case": row["case"],
                       "affinity": row["affinity"],
                       "anon_virtual_growth_bytes": touched["rss"]["anon_virtual_bytes"] - mapped["rss"]["anon_virtual_bytes"],
                       "anon_resident_growth_kib": touched["rss"]["anon_rss_kib"] - mapped["rss"]["anon_rss_kib"]})
        for point in points:
            rss = point["rss"]
            differences.append({"index": row["index"], "phase": point["phase"],
                                "smaps_minus_rollup_kib": rss["smaps_minus_rollup_kib"],
                                "status_minus_rollup_kib": rss["status_minus_rollup_kib"]})
    return {"full_children": len(full), "identity_children": len(identity),
            "full_snapshots": len(full) * 6, "identity_snapshots": len(identity) * 6,
            "anon_mapped_to_touched": growth, "rss_counter_differences": differences,
            "gnu_time_vs_wait4": [{"index": row["index"],
                "wait4_maxrss_kib": row["wait4"]["maxrss_kib"],
                "gnu_time_maxrss_kib": row["gnu_time"]["maxrss_kib"]}
                for row in rows if row["gnu_time"] is not None]}
def build_analysis() -> dict[str, Any]:
    plan = load_plan()
    source = verify_source_custody()
    raw = load_receipts()
    matrix_keys, trace_keys = expected_keys(plan)
    rows = []
    seen: set[tuple[Any, ...]] = set()
    for raw_row in raw:
        trace = bool(raw_row.get("trace", False)) if isinstance(raw_row, dict) else False
        key = row_key(raw_row, trace)
        expected = trace_keys if trace else matrix_keys
        require(key in expected, f"unexpected child matrix key: {key}")
        require(key not in seen, f"duplicate child matrix key: {key}")
        seen.add(key)
        rows.append(verify_row(raw_row, source, plan))
    require(seen == matrix_keys | trace_keys, "aggregate child matrix is incomplete")
    rows.sort(key=lambda row: (row["trace"], row["index"]))
    counts = {"children": len(rows), "matrix_children": sum(not row["trace"] for row in rows),
              "trace_children": sum(row["trace"] for row in rows),
              "full_snapshots": sum(row["observer"] == "full" for row in rows) * 6,
              "identity_snapshots": sum(row["observer"] == "identity" for row in rows) * 6}
    require(counts == {"children": 276, "matrix_children": 272, "trace_children": 4,
                       "full_snapshots": 840, "identity_snapshots": 816},
            f"aggregate counts changed: {counts}")
    return {"schema": RESULT_SCHEMA, "plan_schema": plan["schema"],
            "purpose": plan["purpose"], "diagnostic_only": True,
            "adoption_decision": "not evaluated", "statistics": "two repeats; no bootstrap",
            "source": source, "counts": counts, "rows": rows,
            "summary": summarize(rows),
            "checks": {"matrix_checked": True, "protocol_alignment_checked": True,
                       "source_and_binary_hashes_checked": True,
                       "self_rusage_before_after_ack_retained": True,
                       "wait4_and_gnu_time_scopes_separate": True,
                       "smaps_rollup_status_differences_retained": True,
                       "known_anonymous_growth_retained": True,
                       "touched_page_checksum_checked": True,
                       "no_native_performance_claim": True}}
def main() -> int:
    import argparse
    parser = argparse.ArgumentParser()
    group = parser.add_mutually_exclusive_group()
    group.add_argument("--write", action="store_true")
    group.add_argument("--check", action="store_true")
    args = parser.parse_args()
    try:
        result = build_analysis()
        encoded = json.dumps(result, indent=2, sort_keys=True) + "\n"
        counts = result["counts"]
        rendered = ("# 0789 RSS accounting calibration\n\n"
                    "Diagnostic-only six-phase replay; two repeats are observations, "
                    "not confidence intervals. RUSAGE, wait4, GNU time, smaps and "
                    "status remain separate; no performance or adoption decision.\n\n"
                    f"- Children: {counts['children']} ({counts['matrix_children']} matrix, "
                    f"{counts['trace_children']} traces)\n"
                    f"- Full snapshots: {counts['full_snapshots']}; identity snapshots: "
                    f"{counts['identity_snapshots']}\n")
        if args.check:
            require(OUT_JSON.read_text() == encoded, "analysis.json is stale")
            require(OUT_MD.read_text() == rendered, "analysis.md is stale")
            print("0789 analysis replay PASS")
        else:
            require(not OUT_JSON.exists() and not OUT_MD.exists(),
                    "refusing to overwrite retained analysis outputs")
            OUT_JSON.write_text(encoded)
            OUT_MD.write_text(rendered)
            print("0789 analysis written")
        return 0
    except (ReplayError, OSError) as error:
        print(f"0789 analysis failed: {error}", file=sys.stderr)
        return 1
if __name__ == "__main__":
    raise SystemExit(main())
