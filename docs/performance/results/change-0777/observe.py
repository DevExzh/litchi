"""Validate the separate instruction and allocation observations for 0777.

The native timing table and the two instrumentation families have deliberately
separate evidence paths.  This module only reads their retained receipts.  It
does not run a probe, inspect a live target directory, or treat an unsupported
``perf`` qualification as a zero-result observation.

``analyze.py`` remains the source of truth for the 114 native identities.  The
reports below are checked against the corresponding entries in
``analysis.json`` before their counters or allocation totals are summarized.
"""

from __future__ import annotations

import hashlib
import json
import math
import re
import statistics
from pathlib import Path
from typing import Any, Iterable

import analyze


PACKET = analyze.PACKET
CAPTURE = analyze.CAPTURE
EVENTS = ("instructions:u", "cycles:u", "branches:u", "branch-misses:u")
INSTRUCTION_DIRS = (
    ("primary", PACKET / "instructions-0", 7, 56, 8),
    ("n256", PACKET / "instructions-256", 1, 8, 0),
)
ALLOCATION_DIRS = (
    ("primary", PACKET / "allocations-0", 5, 10),
    ("n256", PACKET / "allocations-256", 1, 2),
)
INSTRUCTION_CASES = (
    ("mce_benign_worksheet", None),
    ("mce_benign_document", None),
    ("mce_prefixed_32", None),
    ("opc_relationship_declarations", 8),
    ("opc_relationship_declarations", 30),
    ("opc_relationship_declarations", 4096),
    ("opc_relationship_declarations", 16384),
)
SECONDARY_INSTRUCTION_CASES = (("opc_relationship_declarations", 256),)
ALLOCATION_CASES = (
    ("mce_benign_worksheet", None),
    ("mce_benign_document", None),
    ("opc_relationship_declarations", 29),
    ("opc_relationship_declarations", 30),
    ("opc_relationship_declarations", 4096),
)
SECONDARY_ALLOCATION_CASES = (("opc_relationship_declarations", 256),)
PAIR_SCHEDULE = (("before", 3), ("after", 3), ("after", 23), ("before", 23))
_DIGEST = re.compile(r"^[0-9a-fA-F]{64}$")
_INTEGER = re.compile(r"^[0-9]+$")


def fail(message: str) -> None:
    raise RuntimeError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def read_json(path: Path) -> Any:
    require(path.is_file(), f"missing JSON receipt: {path}")
    try:
        return json.loads(path.read_text())
    except (OSError, ValueError) as error:
        fail(f"invalid JSON receipt {path}: {error}")


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for chunk in iter(lambda: source.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def digest(value: Any, label: str) -> str:
    require(isinstance(value, str) and _DIGEST.fullmatch(value) is not None, f"{label} has invalid SHA-256")
    return value


def relative(path: Path) -> str:
    try:
        return str(path.relative_to(PACKET))
    except ValueError:
        return str(path)


def check_local_file(root: Path, name: Any, expected_sha: Any, label: str) -> Path:
    require(isinstance(name, str) and name, f"{label} has no artifact name")
    relative_name = Path(name)
    require(
        not relative_name.is_absolute()
        and len(relative_name.parts) == 1
        and relative_name.name == name
        and name not in (".", ".."),
        f"{label} is not a local artifact name: {name!r}",
    )
    path = root / relative_name
    require(path.is_file(), f"missing {label}: {path}")
    expected = digest(expected_sha, f"{label} SHA-256")
    actual = sha256(path)
    require(actual == expected, f"{label} SHA-256 changed")
    return path


def check_optional_local_file(root: Path, name: Any, expected_sha: Any, label: str) -> Path | None:
    """Check a retained artifact when an unsupported command produced one.

    ``instructions.py`` writes ``None`` for a CSV digest if ``perf`` failed
    before creating the CSV.  That is a valid failed qualification receipt,
    provided the explicit no-claim marker is present; a file with a missing or
    wrong digest remains an error.
    """

    if name is None and expected_sha is None:
        return None
    require(isinstance(name, str) and name, f"{label} has an invalid artifact name")
    candidate = root / Path(name)
    if not candidate.is_file():
        require(expected_sha is None, f"missing {label} with a recorded SHA-256")
        return None
    require(expected_sha is not None, f"{label} exists without a SHA-256")
    return check_local_file(root, name, expected_sha, label)


def as_int(value: Any, label: str, *, positive: bool = False) -> int:
    require(isinstance(value, int) and not isinstance(value, bool), f"{label} is not an integer")
    require(not positive or value > 0, f"{label} is not positive")
    return value


def case_key(case: str, n: int | None) -> str:
    return analyze.case_id(case, n)


def load_capture_and_analysis() -> tuple[dict[str, Any], dict[str, Any], dict[str, Any]]:
    """Replay capture receipts and load the already-produced native analysis."""

    complete, _runs = analyze.load_complete()
    analysis_path = PACKET / "analysis.json"
    analysis = read_json(analysis_path)
    require(isinstance(analysis, dict), "analysis.json is not an object")
    require(
        analysis.get("schema") == "litchi-0777-xml-attribute-analysis-v1",
        "analysis.json schema changed",
    )
    require(analysis.get("candidate") == complete.get("candidate"), "analysis candidate differs from capture")
    require(analysis.get("base") == complete.get("base"), "analysis base differs from capture")
    cases = analysis.get("cases")
    require(isinstance(cases, dict), "analysis.json has no cases")
    require(len(cases) == len(analyze.EXPECTED_CASES), "analysis.json native case count is incomplete")
    require(analysis.get("differential", {}).get("changes") == {}, "native differential contains changes")
    require(analysis.get("equivalence", {}).get("mismatches") == 0, "native equivalence contains mismatches")
    binaries = read_json(CAPTURE / "binaries.json")
    check_binaries_receipt(binaries)
    return complete, analysis, binaries


def check_binaries_receipt(binaries: Any) -> None:
    require(isinstance(binaries, dict), "capture binaries.json is not an object")
    expected = {
        "before": ("attribute_checks", "mce-stream-probe"),
        "after": ("attribute_checks", "attribute_checks_equivalence", "mce-stream-probe"),
    }
    require(set(binaries) == set(expected), "capture binary legs changed")
    for leg, names in expected.items():
        values = binaries.get(leg)
        require(isinstance(values, dict) and set(values) == set(names), f"capture binary set changed for {leg}")
        for name in names:
            record = values[name]
            require(isinstance(record, dict), f"capture binary {leg}/{name} is invalid")
            as_int(record.get("bytes"), f"capture binary {leg}/{name} byte count", positive=True)
            digest(record.get("sha256"), f"capture binary {leg}/{name}")
            path = record.get("path")
            require(isinstance(path, str) and Path(path).name == name, f"capture binary path changed: {leg}/{name}")


def binary_sha(binaries: dict[str, Any], leg: str, case: str, n: int | None) -> str:
    name = "mce-stream-probe" if n is None else "attribute_checks"
    value = binaries[leg][name]
    return digest(value["sha256"], f"{leg}/{name} binary")


def binary_path(binaries: dict[str, Any], leg: str, n: int | None) -> str:
    name = "mce-stream-probe" if n is None else "attribute_checks"
    value = binaries[leg][name]
    path = value.get("path")
    require(isinstance(path, str) and Path(path).name == name, f"{leg}/{name} binary path is invalid")
    return path


def native_identity(analysis: dict[str, Any], case: str, n: int | None) -> dict[str, Any]:
    key = case_key(case, n)
    value = analysis["cases"].get(key)
    require(isinstance(value, dict), f"analysis has no native identity for {key}")
    require(value.get("case") == case and value.get("n") == n, f"analysis identity changed for {key}")
    as_int(value.get("input_bytes"), f"analysis {key} input bytes", positive=True)
    outcomes = value.get("outcomes")
    require(isinstance(outcomes, list) and outcomes and all(isinstance(item, str) for item in outcomes), f"analysis {key} outcomes invalid")
    require(isinstance(value.get("samples"), int) and value["samples"] > 0, f"analysis {key} samples invalid")
    require(isinstance(value.get("warmup"), int) and value["warmup"] >= 0, f"analysis {key} warmup invalid")
    if n is None:
        digest(value.get("input_sha256"), f"analysis {key} input")
    else:
        require(value.get("input_sha256") is None, f"analysis {key} unexpectedly has input SHA")
    return value


def report_timings(report: dict[str, Any], label: str) -> list[int]:
    elapsed = report.get("elapsed_ns")
    durations = report.get("durations_ns")
    require(elapsed is not None or durations is not None, f"{label} has no timings")
    if elapsed is not None and durations is not None:
        require(elapsed == durations, f"{label} timing fields disagree")
    values = elapsed if elapsed is not None else durations
    require(isinstance(values, list) and values, f"{label} timings are invalid")
    require(all(isinstance(item, int) and not isinstance(item, bool) and item > 0 for item in values), f"{label} timings are invalid")
    return values


def check_probe_report(
    path: Path,
    row: dict[str, Any],
    analysis: dict[str, Any],
    case: str,
    n: int | None,
    label: str,
) -> dict[str, Any]:
    value = read_json(path)
    require(isinstance(value, dict), f"{label} report is not an object")
    identity = native_identity(analysis, case, n)
    require(value.get("case") == case, f"{label} report case changed")
    if n is None:
        require(value.get("n") is None, f"{label} report unexpectedly has n")
        require(value.get("input_sha256") == identity.get("input_sha256"), f"{label} input identity changed")
    else:
        require(value.get("n") == n, f"{label} report n changed")
        require("input_sha256" not in value, f"{label} OPC report gained input SHA")
    require(value.get("input_bytes") == identity["input_bytes"], f"{label} input size changed")
    require(value.get("outcomes") == identity["outcomes"], f"{label} outcomes differ from native analysis")
    require(value.get("samples") == row["samples"], f"{label} sample count changed")
    require(value.get("warmup") == row["warmup"], f"{label} warmup changed")
    timings = report_timings(value, label)
    require(len(timings) == row["samples"], f"{label} timing count changed")
    if "p50_ns" in value:
        require(value["p50_ns"] == statistics.median(timings), f"{label} p50 changed")
    return {
        "input_bytes": identity["input_bytes"],
        "input_sha256": identity.get("input_sha256"),
        "outcomes": list(identity["outcomes"]),
        "timings_count": len(timings),
    }


def parse_perf_csv(
    path: Path, label: str, expected_events: Iterable[str] = EVENTS
) -> dict[str, dict[str, float | int]]:
    """Parse perf's ``-x ';'`` rows and retain the running fraction.

    In this output format field 3 is the enabled runtime in nanoseconds and
    field 4 is ``time_running / time_enabled * 100``.  The latter is a
    multiplexing fraction, not a second time counter.
    """

    expected = tuple(expected_events)
    require(expected and set(expected) <= set(EVENTS), f"{label} requested invalid perf events")
    values: dict[str, dict[str, float | int]] = {}
    for line_number, raw in enumerate(path.read_text().splitlines(), 1):
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        parts = line.split(";")
        require(len(parts) >= 5, f"{label} line {line_number} is not perf CSV")
        value_text = parts[0].strip()
        event = parts[2].strip()
        running_text = parts[3].strip()
        running_percent_text = parts[4].strip()
        require(event in EVENTS, f"{label} has unexpected event {event!r}")
        require(event not in values, f"{label} repeats event {event}")
        require(_INTEGER.fullmatch(value_text) is not None, f"{label} {event} count is not numeric")
        try:
            running = float(running_text)
            running_percent = float(running_percent_text)
        except ValueError:
            fail(f"{label} {event} timing fields are not numeric")
        require(math.isfinite(running) and running > 0, f"{label} {event} has no running time")
        require(
            math.isfinite(running_percent) and 0 < running_percent <= 100,
            f"{label} {event} has invalid running fraction",
        )
        running_value: float | int = (
            int(running_text) if _INTEGER.fullmatch(running_text) is not None else running
        )
        values[event] = {
            "count": int(value_text),
            "time_running_ns": running_value,
            "running_percent": running_percent,
        }
    require(set(values) == set(expected), f"{label} does not contain the expected perf events")
    return values


def check_qualification(root: Path, complete: dict[str, Any], label: str) -> dict[str, Any]:
    qualification = read_json(root / "qualification.json")
    require(isinstance(qualification, dict), f"{label} qualification is not an object")
    require_command(qualification.get("command"), qualification_command(root), f"{label} qualification")
    claimed_supported = complete.get("supported")
    require(isinstance(claimed_supported, bool), f"{label} complete.supported is not boolean")
    if claimed_supported:
        csv_path = check_local_file(root, qualification.get("csv"), qualification.get("csv_sha256"), f"{label} qualification CSV")
        log_path = check_local_file(root, qualification.get("log"), qualification.get("log_sha256"), f"{label} qualification log")
    else:
        csv_path = check_optional_local_file(root, qualification.get("csv"), qualification.get("csv_sha256"), f"{label} qualification CSV")
        log_path = check_optional_local_file(root, qualification.get("log"), qualification.get("log_sha256"), f"{label} qualification log")
    exit_code = as_int(qualification.get("exit"), f"{label} qualification exit")
    perf: dict[str, Any] | None = None
    if csv_path is not None:
        try:
            perf = parse_perf_csv(csv_path, f"{label} qualification", ("instructions:u",))
        except RuntimeError:
            perf = None
    instruction_supported = perf is not None and "instructions:u" in perf
    supported = claimed_supported
    require(supported == (exit_code == 0 and instruction_supported), f"{label} qualification support flag disagrees with receipts")
    if supported:
        require(complete.get("source_unchanged") is True, f"{label} source guard is false")
        require(complete.get("fixtures_unchanged") is True, f"{label} fixture guard is false")
        require(exit_code == 0, f"{label} supported qualification failed")
    else:
        reason = complete.get("reason")
        require(isinstance(reason, str) and "No instruction result claimed" in reason, f"{label} unsupported qualification lacks explicit no-claim reason")
    return {
        "supported": supported,
        "exit": exit_code,
        "csv": relative(csv_path) if csv_path is not None else None,
        "csv_sha256": digest(qualification["csv_sha256"], f"{label} qualification CSV") if qualification.get("csv_sha256") is not None else None,
        "log": relative(log_path) if log_path is not None else None,
        "log_sha256": digest(qualification["log_sha256"], f"{label} qualification log") if qualification.get("log_sha256") is not None else None,
        "perf": perf,
        "reason": complete.get("reason") if not supported else None,
    }


def row_artifact(root: Path, row: dict[str, Any], field: str, label: str) -> tuple[Path, str]:
    sha_field = f"{field}_sha256"
    path = check_local_file(root, row.get(field), row.get(sha_field), f"{label} {field}")
    return path, digest(row[sha_field], f"{label} {field}")


def expected_name(case: str, n: int | None, repeat: int, leg: str, samples: int) -> str:
    return f"{case}-{n}-{repeat}-{leg}-{samples}"


def require_command(actual: Any, expected: list[str], label: str) -> None:
    # Commands retain their capture-time output paths after packet relocation.
    # Bind those paths to the frozen source root, not the replay checkout.
    frozen_root = Path(read_json(CAPTURE / "source-after.json")["root"])
    frozen_packet = frozen_root / "docs/performance/results/change-0777"
    expected = [
        str(frozen_packet / Path(argument).relative_to(PACKET))
        if argument.startswith(str(PACKET) + "/") else argument
        for argument in expected
    ]
    require(isinstance(actual, list) and actual == expected, f"{label} command argv changed")


def qualification_command(root: Path) -> list[str]:
    return [
        "perf",
        "stat",
        "-x",
        ";",
        "--no-big-num",
        "-e",
        "instructions:u",
        "-o",
        str(root / "qualification.csv"),
        "--",
        "taskset",
        "-c",
        "12",
        "/usr/bin/true",
    ]


def instruction_command(
    root: Path,
    binaries: dict[str, Any],
    leg: str,
    case: str,
    n: int | None,
    samples: int,
    report_name: str,
) -> list[str]:
    binary = binary_path(binaries, leg, n)
    args = [
        binary,
        "adversarial" if n is None else "probe",
        "--case",
        case,
        "--samples",
        str(samples),
        "--warmup",
        "2",
        "--json",
        str(root / report_name),
    ]
    if n is not None:
        args.extend(("--n", str(n)))
    return [
        "perf",
        "stat",
        "-x",
        ";",
        "--no-big-num",
        "-e",
        ",".join(EVENTS),
        "-o",
        str(root / f"{report_name.removesuffix('.json')}.csv"),
        "--",
        "taskset",
        "-c",
        "12",
        *args,
    ]


def iterator_command(
    root: Path, binaries: dict[str, Any], mode: str, repeat: int, rounds: int
) -> list[str]:
    return [
        "perf",
        "stat",
        "-x",
        ";",
        "--no-big-num",
        "-e",
        ",".join(EVENTS),
        "-o",
        str(root / f"iterator-{mode}-{repeat}-{rounds}.csv"),
        "--",
        "taskset",
        "-c",
        "12",
        binaries["after"]["attribute_checks_equivalence"]["path"],
        "--bench",
        mode,
        "--rounds",
        str(rounds),
        "test-data/ooxml",
    ]


def allocation_command(
    root: Path,
    binaries: dict[str, Any],
    leg: str,
    case: str,
    n: int | None,
    report_name: str,
) -> list[str]:
    stem = report_name.removesuffix(".json")
    args = [
        binary_path(binaries, leg, n),
        "adversarial" if n is None else "probe",
        "--case",
        case,
        "--samples",
        "1",
        "--warmup",
        "0",
        "--json",
        str(root / report_name),
    ]
    if n is not None:
        args.extend(("--n", str(n)))
    return [
        "taskset",
        "-c",
        "12",
        "heaptrack",
        "--record-only",
        "-o",
        str(root / stem),
        *args,
    ]


def decode_command(root: Path, capture_name: str, histogram_name: str) -> list[str]:
    return [
        "heaptrack_print",
        "-f",
        str(root / capture_name),
        "-H",
        str(root / histogram_name),
        "-p",
        "0",
        "-a",
        "0",
        "-T",
        "0",
        "-l",
        "0",
    ]


def check_instruction_rows(
    root: Path,
    rows: Any,
    cases: tuple[tuple[str, int | None], ...],
    analysis: dict[str, Any],
    binaries: dict[str, Any],
    label: str,
) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    expected_schedule = [
        (case, n, repeat, leg, samples)
        for case, n in cases
        for repeat in range(2)
        for leg, samples in PAIR_SCHEDULE
    ]
    require(isinstance(rows, list) and len(rows) == len(expected_schedule), f"{label} instruction row count changed")
    output_rows: list[dict[str, Any]] = []
    seen_artifacts: set[Path] = set()
    grouped: dict[tuple[str, int | None, int], dict[tuple[str, int], dict[str, Any]]] = {}
    for index, (row, expected) in enumerate(zip(rows, expected_schedule)):
        case, n, repeat, leg, samples = expected
        require(isinstance(row, dict), f"{label} row {index} is not an object")
        actual = (row.get("case"), row.get("n"), row.get("repeat"), row.get("leg"), row.get("samples"))
        require(actual == expected, f"{label} row {index} schedule changed: {actual!r} != {expected!r}")
        require(row.get("exit") == 0, f"{label} row {index} failed")
        require(row.get("warmup") == 2 or "warmup" not in row, f"{label} row {index} has an unexpected warmup field")
        expected_binary = binary_sha(binaries, leg, case, n)
        require(row.get("binary_sha256") == expected_binary, f"{label} row {index} binary identity changed")
        stem = expected_name(case, n, repeat, leg, samples)
        require_command(
            row.get("command"),
            instruction_command(root, binaries, leg, case, n, samples, f"{stem}.json"),
            f"{label} row {index}",
        )
        paths: dict[str, Path] = {}
        hashes: dict[str, str] = {}
        for field in ("log", "csv", "report"):
            path, file_sha = row_artifact(root, row, field, f"{label} row {index}")
            require(path.name == f"{stem}.{field if field != 'report' else 'json'}" if field != "csv" else path.name == f"{stem}.csv", f"{label} row {index} {field} name changed")
            require(path not in seen_artifacts, f"{label} row {index} artifact collision: {path.name}")
            seen_artifacts.add(path)
            paths[field] = path
            hashes[field] = file_sha
        csv_values = parse_perf_csv(paths["csv"], f"{label} row {index} CSV")
        report_row = dict(row)
        report_row["warmup"] = 2
        identity = check_probe_report(paths["report"], report_row, analysis, case, n, f"{label} row {index}")
        log_text = paths["log"].read_text()
        require(log_text == "", f"{label} row {index} native output leaked into instruction log")
        compact = {
            "case": case,
            "n": n,
            "repeat": repeat,
            "leg": leg,
            "samples": samples,
            "warmup": row.get("warmup", 2),
            "binary_sha256": expected_binary,
            "log": relative(paths["log"]),
            "log_sha256": hashes["log"],
            "csv": relative(paths["csv"]),
            "csv_sha256": hashes["csv"],
            "report": relative(paths["report"]),
            "report_sha256": hashes["report"],
            "perf": csv_values,
            "identity": identity,
        }
        output_rows.append(compact)
        grouped.setdefault((case, n, repeat), {})[(leg, samples)] = compact
    summaries: dict[str, Any] = {}
    for case, n in cases:
        key_prefix = case_key(case, n)
        repeats: list[dict[str, Any]] = []
        for repeat in range(2):
            group = grouped[(case, n, repeat)]
            require(set(group) == set(PAIR_SCHEDULE), f"{label} {key_prefix} repeat {repeat} pair is incomplete")
            marginal: dict[str, Any] = {}
            percent: dict[str, float | None] = {}
            raw_counts: dict[str, Any] = {}
            for event in EVENTS:
                before_3 = int(group[("before", 3)]["perf"][event]["count"])
                before_23 = int(group[("before", 23)]["perf"][event]["count"])
                after_3 = int(group[("after", 3)]["perf"][event]["count"])
                after_23 = int(group[("after", 23)]["perf"][event]["count"])
                before_delta = (before_23 - before_3) / 20.0
                after_delta = (after_23 - after_3) / 20.0
                change = None if before_delta == 0 else 100.0 * (after_delta / before_delta - 1.0)
                raw_counts[event] = {
                    "before_3": before_3,
                    "before_23": before_23,
                    "after_3": after_3,
                    "after_23": after_23,
                }
                marginal[event] = {"before_per_operation": before_delta, "after_per_operation": after_delta}
                percent[event] = change
            repeats.append({"repeat": repeat, "raw_counts": raw_counts, "marginal": marginal, "after_vs_before_percent": percent})
        identity = native_identity(analysis, case, n)
        summaries[key_prefix] = {
            "case": case,
            "n": n,
            "input_bytes": identity["input_bytes"],
            "input_sha256": identity.get("input_sha256"),
            "outcomes": list(identity["outcomes"]),
            "repeats": repeats,
        }
    return output_rows, summaries


def iterator_log(path: Path, row: dict[str, Any], label: str) -> dict[str, Any]:
    try:
        value = json.loads(path.read_text())
    except (OSError, ValueError) as error:
        fail(f"invalid {label} JSON log: {error}")
    require(isinstance(value, dict), f"{label} log is not an object")
    require(value.get("iterator") == row["mode"], f"{label} iterator identity changed")
    require(value.get("rounds") == row["rounds"], f"{label} rounds identity changed")
    as_int(value.get("tags"), f"{label} tags", positive=True)
    as_int(value.get("checksum"), f"{label} checksum")
    as_int(value.get("elapsed_ns"), f"{label} elapsed", positive=True)
    return value


def check_iterator_rows(root: Path, rows: Any, binaries: dict[str, Any], label: str) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    expected = [
        (repeat, mode, rounds)
        for repeat in range(2)
        for mode, rounds in (("quick-xml", 3), ("checked", 3), ("checked", 23), ("quick-xml", 23))
    ]
    require(isinstance(rows, list) and len(rows) == len(expected), f"{label} iterator row count changed")
    output: list[dict[str, Any]] = []
    groups: dict[tuple[int, str, int], dict[str, Any]] = {}
    binary = digest(binaries["after"]["attribute_checks_equivalence"]["sha256"], f"{label} iterator binary")
    for index, (row, expected_key) in enumerate(zip(rows, expected)):
        repeat, mode, rounds = expected_key
        require(isinstance(row, dict), f"{label} iterator row {index} is not an object")
        require((row.get("repeat"), row.get("mode"), row.get("rounds")) == expected_key, f"{label} iterator schedule changed")
        require(row.get("exit") == 0, f"{label} iterator row {index} failed")
        require(row.get("binary_sha256") == binary, f"{label} iterator row {index} binary changed")
        stem = f"iterator-{mode}-{repeat}-{rounds}"
        require_command(
            row.get("command"),
            iterator_command(root, binaries, mode, repeat, rounds),
            f"{label} iterator row {index}",
        )
        log_path, log_sha = row_artifact(root, row, "log", f"{label} iterator row {index}")
        csv_path, csv_sha = row_artifact(root, row, "csv", f"{label} iterator row {index}")
        require(log_path.name == f"{stem}.log" and csv_path.name == f"{stem}.csv", f"{label} iterator row {index} artifact name changed")
        perf = parse_perf_csv(csv_path, f"{label} iterator row {index} CSV")
        log_value = iterator_log(log_path, row, f"{label} iterator row {index}")
        compact = {
            "repeat": repeat,
            "mode": mode,
            "rounds": rounds,
            "binary_sha256": binary,
            "log": relative(log_path),
            "log_sha256": log_sha,
            "csv": relative(csv_path),
            "csv_sha256": csv_sha,
            "perf": perf,
            "identity": {"checksum": log_value["checksum"], "tags": log_value["tags"], "rounds": log_value["rounds"]},
        }
        output.append(compact)
        groups[(repeat, mode, rounds)] = compact
    identity_by_round: dict[str, Any] = {}
    repeats: list[dict[str, Any]] = []
    for repeat in range(2):
        repeat_rows: dict[str, Any] = {}
        for rounds in (3, 23):
            quick = groups[(repeat, "quick-xml", rounds)]["identity"]
            checked = groups[(repeat, "checked", rounds)]["identity"]
            require(quick == checked, f"{label} iterator checksum/tags differ at repeat {repeat}, rounds {rounds}")
            identity_by_round[str(rounds)] = quick
            repeat_rows[str(rounds)] = quick
        marginal: dict[str, Any] = {}
        for mode in ("quick-xml", "checked"):
            events: dict[str, float] = {}
            for event in EVENTS:
                small = groups[(repeat, mode, 3)]["perf"][event]["count"]
                large = groups[(repeat, mode, 23)]["perf"][event]["count"]
                events[event] = (large - small) / 20.0
            marginal[mode] = events
        repeats.append({"repeat": repeat, "identity_by_round": repeat_rows, "marginal_per_operation": marginal})
    # Both repeats must describe the same corpus and checksum for each round.
    require(repeats[0]["identity_by_round"] == repeats[1]["identity_by_round"], f"{label} iterator identities differ across repeats")
    return output, {"repeats": repeats, "identity_by_round": identity_by_round}


def validate_instruction_dir(
    label: str,
    root: Path,
    case_count: int,
    row_count: int,
    iterator_count: int,
    cases: tuple[tuple[str, int | None], ...],
    analysis: dict[str, Any],
    binaries: dict[str, Any],
) -> dict[str, Any]:
    complete = read_json(root / "complete.json")
    require(isinstance(complete, dict), f"{label} instructions complete is not an object")
    qualification = check_qualification(root, complete, label)
    if not qualification["supported"]:
        runs_path = root / "runs.json"
        require(not runs_path.exists(), f"{label} unsupported run retained unexpected runs.json")
        require(iterator_count == 0, f"{label} unsupported iterator claim is nonzero")
        return {
            "supported": False,
            "qualification": qualification,
            "claims": None,
            "rows": 0,
            "iterator_rows": 0,
            "reason": qualification["reason"],
        }
    require(complete.get("runs") == row_count, f"{label} instruction complete run count changed")
    require(complete.get("iterator_runs") == iterator_count, f"{label} iterator complete run count changed")
    rows = read_json(root / "runs.json")
    instruction_rows, summaries = check_instruction_rows(root, rows, cases, analysis, binaries, label)
    iterator_summary: dict[str, Any] | None = None
    iterator_rows: list[dict[str, Any]] = []
    iterator_path = root / "iterator-runs.json"
    if iterator_count:
        iterator_value = read_json(iterator_path)
        iterator_rows, iterator_summary = check_iterator_rows(root, iterator_value, binaries, label)
    else:
        require(not iterator_path.exists(), f"{label} has an unexpected iterator-runs.json")
    return {
        "supported": True,
        "qualification": qualification,
        "claims": {
            "scope": complete.get("scope"),
            "event_order": list(EVENTS),
            "cases": summaries,
            "rows": instruction_rows,
            "iterator": iterator_summary,
        },
        "rows": len(instruction_rows),
        "iterator_rows": len(iterator_rows),
    }


def parse_histogram(path: Path, label: str) -> tuple[int, int, int]:
    calls = 0
    allocated = 0
    entries = 0
    seen_sizes: set[int] = set()
    for line_number, raw in enumerate(path.read_text().splitlines(), 1):
        line = raw.strip()
        if not line:
            continue
        fields = line.split()
        require(len(fields) == 2, f"{label} histogram line {line_number} is malformed")
        try:
            size, count = (int(item) for item in fields)
        except ValueError:
            fail(f"{label} histogram line {line_number} is not numeric")
        require(size > 0 and count > 0, f"{label} histogram line {line_number} is not positive")
        require(size not in seen_sizes, f"{label} histogram repeats allocation size {size}")
        seen_sizes.add(size)
        calls += count
        allocated += size * count
        entries += 1
    require(entries > 0, f"{label} histogram is empty")
    return entries, calls, allocated


def summary_call_count(path: Path, label: str) -> int:
    match = re.search(r"calls to allocation functions:\s*([0-9]+)", path.read_text())
    require(match is not None, f"{label} summary has no allocation call count")
    return int(match.group(1))


def check_allocation_rows(
    root: Path,
    rows: Any,
    cases: tuple[tuple[str, int | None], ...],
    analysis: dict[str, Any],
    binaries: dict[str, Any],
    label: str,
) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    expected = [(case, n, leg) for case, n in cases for leg in ("before", "after")]
    require(isinstance(rows, list) and len(rows) == len(expected), f"{label} allocation row count changed")
    compact_rows: list[dict[str, Any]] = []
    grouped: dict[tuple[str, int | None], dict[str, Any]] = {}
    seen: set[Path] = set()
    for index, (row, (case, n, leg)) in enumerate(zip(rows, expected)):
        require(isinstance(row, dict), f"{label} allocation row {index} is not an object")
        require((row.get("case"), row.get("n"), row.get("leg")) == (case, n, leg), f"{label} allocation row {index} schedule changed")
        require(row.get("exit") == 0 and row.get("decode_exit") == 0, f"{label} allocation row {index} failed")
        require(row.get("binary_sha256") == binary_sha(binaries, leg, case, n), f"{label} allocation row {index} binary changed")
        stem = f"{case}-{n}-{leg}"
        require_command(
            row.get("command"),
            allocation_command(root, binaries, leg, case, n, f"{stem}.json"),
            f"{label} allocation row {index}",
        )
        paths: dict[str, Path] = {}
        hashes: dict[str, str] = {}
        for field, suffix in (("stdout", ".stdout"), ("stderr", ".stderr"), ("report", ".json"), ("capture", None), ("summary", ".summary"), ("histogram", ".histogram")):
            path, file_sha = row_artifact(root, row, field, f"{label} row {index}")
            if suffix is not None:
                require(path.name == stem + suffix, f"{label} row {index} {field} name changed")
            else:
                require(path.name in (stem + ".zst", stem + ".gz"), f"{label} row {index} capture name changed")
            require(path not in seen, f"{label} row {index} artifact collision: {path.name}")
            seen.add(path)
            paths[field] = path
            hashes[field] = file_sha
        require_command(
            row.get("decode_command"),
            decode_command(root, paths["capture"].name, paths["histogram"].name),
            f"{label} allocation row {index} decode",
        )
        entries, histogram_calls, histogram_bytes = parse_histogram(paths["histogram"], f"{label} row {index}")
        calls = as_int(row.get("allocation_calls"), f"{label} row {index} allocation calls")
        allocated = as_int(row.get("allocated_bytes"), f"{label} row {index} allocated bytes")
        require(histogram_calls == calls, f"{label} row {index} histogram calls disagree")
        require(histogram_bytes == allocated, f"{label} row {index} histogram bytes disagree")
        summary_calls = summary_call_count(paths["summary"], f"{label} row {index}")
        require(summary_calls == calls, f"{label} row {index} summary calls disagree")
        probe_row = {"samples": 1, "warmup": 0}
        identity = check_probe_report(paths["report"], probe_row, analysis, case, n, f"{label} row {index}")
        compact = {
            "case": case,
            "n": n,
            "leg": leg,
            "binary_sha256": row["binary_sha256"],
            "stdout": relative(paths["stdout"]),
            "stdout_sha256": hashes["stdout"],
            "stderr": relative(paths["stderr"]),
            "stderr_sha256": hashes["stderr"],
            "report": relative(paths["report"]),
            "report_sha256": hashes["report"],
            "capture": relative(paths["capture"]),
            "capture_sha256": hashes["capture"],
            "summary": relative(paths["summary"]),
            "summary_sha256": hashes["summary"],
            "histogram": relative(paths["histogram"]),
            "histogram_sha256": hashes["histogram"],
            "histogram_entries": entries,
            "allocation_calls": calls,
            "allocated_bytes": allocated,
            "identity": identity,
        }
        compact_rows.append(compact)
        grouped.setdefault((case, n), {})[leg] = compact
    summaries: dict[str, Any] = {}
    for case, n in cases:
        key = case_key(case, n)
        pair = grouped[(case, n)]
        require(set(pair) == {"before", "after"}, f"{label} {key} allocation pair is incomplete")
        before = pair["before"]
        after = pair["after"]
        calls_before = before["allocation_calls"]
        bytes_before = before["allocated_bytes"]
        calls_after = after["allocation_calls"]
        bytes_after = after["allocated_bytes"]
        identity = native_identity(analysis, case, n)
        summaries[key] = {
            "case": case,
            "n": n,
            "input_bytes": identity["input_bytes"],
            "input_sha256": identity.get("input_sha256"),
            "outcomes": list(identity["outcomes"]),
            "before": {"allocation_calls": calls_before, "allocated_bytes": bytes_before},
            "after": {"allocation_calls": calls_after, "allocated_bytes": bytes_after},
            "change": {
                "allocation_calls": calls_after - calls_before,
                "allocated_bytes": bytes_after - bytes_before,
                "allocation_calls_percent": None if calls_before == 0 else 100.0 * (calls_after / calls_before - 1.0),
                "allocated_bytes_percent": None if bytes_before == 0 else 100.0 * (bytes_after / bytes_before - 1.0),
            },
        }
    return compact_rows, summaries


def validate_allocation_dir(
    label: str,
    root: Path,
    row_count: int,
    cases: tuple[tuple[str, int | None], ...],
    analysis: dict[str, Any],
    binaries: dict[str, Any],
) -> dict[str, Any]:
    complete = read_json(root / "complete.json")
    require(isinstance(complete, dict), f"{label} allocations complete is not an object")
    require(complete.get("source_unchanged") is True, f"{label} source guard is false")
    require(complete.get("fixtures_unchanged") is True, f"{label} fixture guard is false")
    require(complete.get("runs") == row_count, f"{label} allocation complete run count changed")
    rows = read_json(root / "runs.json")
    compact, summaries = check_allocation_rows(root, rows, cases, analysis, binaries, label)
    return {
        "supported": True,
        "scope": complete.get("scope"),
        "claims": {"cases": summaries, "rows": compact},
        "rows": len(compact),
    }


def observe() -> dict[str, Any]:
    complete, analysis, binaries = load_capture_and_analysis()
    instruction_results: dict[str, Any] = {}
    for label, root, case_count, row_count, iterator_count in INSTRUCTION_DIRS:
        cases = INSTRUCTION_CASES if label == "primary" else SECONDARY_INSTRUCTION_CASES
        require(len(cases) == case_count, f"{label} instruction case schema changed")
        instruction_results[label] = validate_instruction_dir(
            label, root, case_count, row_count, iterator_count, cases, analysis, binaries
        )
    allocation_results: dict[str, Any] = {}
    for label, root, case_count, row_count in ALLOCATION_DIRS:
        cases = ALLOCATION_CASES if label == "primary" else SECONDARY_ALLOCATION_CASES
        require(len(cases) == case_count, f"{label} allocation case schema changed")
        allocation_results[label] = validate_allocation_dir(label, root, row_count, cases, analysis, binaries)
    return {
        "schema": "litchi-0777-xml-attribute-observations-v1",
        "base": complete["base"],
        "candidate": complete["candidate"],
        "capture": {
            "schema": complete.get("schema"),
            "source_unchanged": complete.get("source_unchanged"),
            "fixtures_unchanged": complete.get("fixtures_unchanged"),
            "cpu": complete.get("cpu"),
        },
        "instructions": instruction_results,
        "allocations": allocation_results,
        "limits": (
            "Instruction values are whole-process user counters. Per-repeat (23-3)/20 values are marginal estimates and include process/report work; they are not exact timed-region or function attribution. Allocation totals are one whole-process heaptrack sample per case and leg and include startup, input construction, parsing, observation, and report serialization."
        ),
    }


if __name__ == "__main__":
    result = observe()
    (PACKET / "observations.json").write_text(json.dumps(result, indent=2, sort_keys=True) + "\n")
    print(
        json.dumps(
            {
                "schema": result["schema"],
                "instruction_rows": sum(item["rows"] for item in result["instructions"].values()),
                "iterator_rows": sum(item["iterator_rows"] for item in result["instructions"].values()),
                "allocation_rows": sum(item["rows"] for item in result["allocations"].values()),
            },
            sort_keys=True,
        )
    )
