"""Replay and summarize the complete 0777 paired capture.

This module intentionally treats the capture as an evidence receipt.  It does
not run a binary, rebuild a probe, or make changed outputs equivalent.  The
capture must have reached its explicit completion marker before any result is
accepted.
"""

from __future__ import annotations

import hashlib
import json
import statistics
from pathlib import Path
from typing import Any, Iterable


PACKET = Path(__file__).resolve().parent
CAPTURE = PACKET / "capture-0"
LEGS = ("before", "after", "after", "before", "before", "after")
SAMPLES = 9
WARMUP = 2
EXPECTED_MCE = (
    "mce_benign_worksheet",
    "mce_benign_document",
    "mce_stream_count_worksheet",
    "mce_stream_count_document",
    "mce_prefixed_1",
    "mce_prefixed_2",
    "mce_prefixed_8",
    "mce_prefixed_9",
    "mce_prefixed_32",
)
EXPECTED_OPC = "opc_relationship_declarations"
EXPECTED_N = (0, 8, 29, 30, 32, 33, 256, 1024, 4096, 16384)
EXPECTED_OPC_LIMIT_ERROR = (
    "err:Quick-XML error: start tag declares more than 256 namespace bindings; "
    "raise the limit with NamespaceResolver::set_max_declarations_per_element"
)
EXPECTED_CASES = tuple((case, None) for case in EXPECTED_MCE) + tuple(
    (EXPECTED_OPC, n) for n in EXPECTED_N
)
EXPECTED_NATIVE_PROCESSES = len(EXPECTED_CASES) * len(LEGS)
PROBE_INPUTS = (
    "attribute_checks.rs",
    "attribute_checks_equivalence.rs",
    "main.rs",
    "Cargo.toml.template",
)
_HEX = frozenset("0123456789abcdefABCDEF")


def fail(message: str) -> None:
    raise RuntimeError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for chunk in iter(lambda: source.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def read_json(path: Path) -> Any:
    require(path.is_file(), f"missing JSON receipt: {path}")
    try:
        return json.loads(path.read_text())
    except (OSError, ValueError) as error:
        fail(f"invalid JSON receipt {path}: {error}")


def is_hex_digest(value: Any) -> bool:
    return isinstance(value, str) and len(value) == 64 and all(
        character in _HEX for character in value
    )


def relative_packet(path: Path) -> str:
    try:
        return str(path.relative_to(PACKET))
    except ValueError:
        return str(path)


def _candidate_from_old_path(raw: Path) -> Path | None:
    """Map an absolute path from a previous packet checkout to this packet.

    Capture rows historically stored absolute paths.  The packet is often
    replayed after its worktree has been removed, so using the old path would
    make a valid receipt appear missing.  Mapping by the stable packet or
    capture directory component also avoids a basename-only collision between
    the 114 native reports.
    """

    parts = raw.parts
    for marker in (CAPTURE.name, PACKET.name):
        if marker in parts:
            index = parts.index(marker)
            suffix = parts[index + 1 :]
            candidate = (CAPTURE if marker == CAPTURE.name else PACKET).joinpath(
                *suffix
            )
            return candidate
    return None


def resolve_packet_path(value: Any, *, prefer_capture: bool = False) -> Path:
    require(isinstance(value, str) and value, f"invalid artifact path: {value!r}")
    raw = Path(value)
    candidates: list[Path] = []
    if raw.is_absolute():
        relocated = _candidate_from_old_path(raw)
        if relocated is not None:
            candidates.append(relocated)
        candidates.append(raw)
    else:
        text = value.replace("\\", "/")
        if text == CAPTURE.name or text.startswith(CAPTURE.name + "/"):
            candidates.append(CAPTURE / Path(text).relative_to(CAPTURE.name))
        if prefer_capture:
            candidates.append(CAPTURE / raw)
        candidates.append(PACKET / raw)
        candidates.append(CAPTURE / raw)
    for candidate in candidates:
        if candidate.is_file():
            return candidate
    # Return the packet-local candidate to make the error useful and to keep
    # callers from accidentally falling back to an unrelated basename.
    if candidates:
        return candidates[0]
    return raw


def check_receipt(
    receipt: Any,
    label: str,
    *,
    capture_bound: bool = True,
    allow_missing: bool = False,
) -> Path | None:
    """Check a ``path/bytes/sha256`` receipt and return its local path.

    Capture artifacts must resolve inside this packet.  External binary
    receipts are allowed to point at cleaned target directories; when such a
    path no longer exists the recorded size and digest remain the evidence.
    """

    require(isinstance(receipt, dict), f"{label} is not an artifact object")
    raw = receipt.get("path")
    require(isinstance(raw, str) and raw, f"{label} has no path")
    expected_bytes = receipt.get("bytes")
    require(
        isinstance(expected_bytes, int) and expected_bytes >= 0,
        f"{label} has invalid byte count",
    )
    expected_sha = receipt.get("sha256")
    require(is_hex_digest(expected_sha), f"{label} has invalid SHA-256")
    path = resolve_packet_path(raw, prefer_capture=capture_bound)
    if capture_bound:
        try:
            path.relative_to(CAPTURE)
        except ValueError:
            fail(f"{label} did not relocate into capture packet: {raw}")
    if not path.is_file():
        if allow_missing and not capture_bound:
            return None
        fail(f"missing {label} artifact after path relocation: {raw}")
    actual_bytes = path.stat().st_size
    actual_sha = sha256(path)
    require(
        actual_bytes == expected_bytes,
        f"{label} byte count changed: {actual_bytes} != {expected_bytes}",
    )
    require(actual_sha == expected_sha, f"{label} SHA-256 changed")
    return path


def check_text_receipt(
    path_value: Any, expected_sha: Any, label: str, *, allow_missing: bool = False
) -> Path | None:
    path = resolve_packet_path(path_value)
    if not path.is_file():
        if allow_missing:
            return None
        fail(f"missing {label}: {path_value}")
    require(is_hex_digest(expected_sha), f"{label} has invalid SHA-256")
    require(sha256(path) == expected_sha, f"{label} SHA-256 changed")
    return path


def artifact_path(receipt: Any, label: str) -> Path:
    path = check_receipt(receipt, label, capture_bound=True)
    assert path is not None
    return path


def load_complete() -> tuple[dict[str, Any], list[dict[str, Any]]]:
    require(CAPTURE.is_dir(), f"capture is missing: {CAPTURE}")
    complete = read_json(CAPTURE / "complete.json")
    require(isinstance(complete, dict), "complete.json is not an object")
    require(complete.get("complete") is True, "capture is not complete")
    require(complete.get("serial") is True, "capture was not serial")
    require(tuple(complete.get("order", ())) == LEGS, "capture order changed")
    require(complete.get("samples") == SAMPLES, "capture samples changed")
    require(complete.get("warmup") == WARMUP, "capture warmup changed")
    require(
        complete.get("native_cases") == len(EXPECTED_CASES),
        "capture native case count is incomplete",
    )
    require(
        complete.get("native_processes") == EXPECTED_NATIVE_PROCESSES,
        "capture native process count is incomplete",
    )
    runs_value = read_json(CAPTURE / "runs.json")
    require(isinstance(runs_value, list), "runs.json is not an array")
    runs: list[dict[str, Any]] = []
    for index, row in enumerate(runs_value):
        require(isinstance(row, dict), f"runs[{index}] is not an object")
        require(row.get("exit") == 0, f"run {index} did not succeed")
        require("log" in row, f"run {index} has no log receipt")
        artifact_path(row["log"], f"runs[{index}].log")
        for field in ("report", "rss"):
            if field in row and row[field] is not None:
                artifact_path(row[field], f"runs[{index}].{field}")
        binary = row.get("binary")
        if binary is not None:
            # This is an external receipt and may have been cleaned already.
            check_receipt(
                binary,
                f"runs[{index}].binary",
                capture_bound=False,
                allow_missing=True,
            )
        runs.append(row)
    require(len(runs) == 120, f"expected 120 total run rows, found {len(runs)}")
    kinds = {kind: sum(row.get("kind") == kind for row in runs) for kind in {
        "standalone-lock",
        "build",
        "equivalence",
        "differential",
        "native",
    }}
    require(
        kinds == {
            "standalone-lock": 1,
            "build": 2,
            "equivalence": 1,
            "differential": 2,
            "native": EXPECTED_NATIVE_PROCESSES,
        },
        f"unexpected run kinds: {kinds}",
    )
    return complete, runs


def timing_values(report: dict[str, Any], label: str) -> list[int]:
    elapsed = report.get("elapsed_ns")
    durations = report.get("durations_ns")
    require(elapsed is not None or durations is not None, f"{label} has no timings")
    if elapsed is not None and durations is not None:
        require(elapsed == durations, f"{label} has conflicting timing fields")
    values = elapsed if elapsed is not None else durations
    require(isinstance(values, list), f"{label} timings are not an array")
    require(len(values) == SAMPLES, f"{label} has wrong timing count")
    require(
        all(isinstance(value, int) and value > 0 for value in values),
        f"{label} has invalid timing values",
    )
    return values


def summary(values: Iterable[int]) -> dict[str, Any]:
    values = list(values)
    require(values, "cannot summarize an empty sample")
    median = statistics.median(values)
    require(median > 0, "cannot compute spread with a zero median")
    return {
        "median": median,
        "min": min(values),
        "max": max(values),
        "spread_percent": 100.0 * (max(values) - min(values)) / median,
    }


def case_id(case: str, n: int | None) -> str:
    return case if n is None else f"{case}::n{n}"


def check_probe_generator() -> None:
    """Bind the OPC input to the archived deterministic generator source."""

    source = PACKET / "probe-src/attribute_checks.rs"
    require(source.is_file(), "missing archived OPC probe source")
    text = source.read_text()
    for marker in (
        "fn relationship_declarations(n: usize)",
        "for index in 0..n",
        "StreamingArchiveWriter::new()",
        'write_stored("_rels/.rels"',
    ):
        require(marker in text, f"OPC generator marker missing: {marker}")


def native_rows(
    runs: list[dict[str, Any]],
) -> tuple[dict[tuple[str, int | None], dict[str, Any]], list[dict[str, Any]]]:
    check_probe_generator()
    rows = [row for row in runs if row.get("kind") == "native"]
    require(len(rows) == EXPECTED_NATIVE_PROCESSES, "native rows are incomplete")
    expected_set = set(EXPECTED_CASES)
    grouped: dict[tuple[str, int | None], list[dict[str, Any]]] = {}
    report_paths: set[Path] = set()
    log_paths: set[Path] = set()
    rss_paths: set[Path] = set()
    for row_index, row in enumerate(rows):
        case = row.get("case")
        n_value = row.get("n")
        n: int | None
        if n_value is None:
            n = None
        else:
            require(isinstance(n_value, int), f"native row {row_index} has invalid n")
            n = n_value
        key = (case, n)
        require(key in expected_set, f"unexpected native case: {key!r}")
        require(row.get("order_index") == len(grouped.get(key, [])), f"bad order index for {key}")
        expected_leg = LEGS[row.get("order_index", -1)] if isinstance(row.get("order_index"), int) and 0 <= row.get("order_index", -1) < len(LEGS) else None
        require(row.get("leg") == expected_leg, f"bad leg/order for {key}")
        report_path = artifact_path(row.get("report"), f"native {key} report")
        log_path = artifact_path(row.get("log"), f"native {key} log")
        rss_path = artifact_path(row.get("rss"), f"native {key} RSS")
        require(report_path not in report_paths, f"native report path collision: {report_path}")
        require(log_path not in log_paths, f"native log path collision: {log_path}")
        require(rss_path not in rss_paths, f"native RSS path collision: {rss_path}")
        report_paths.add(report_path)
        log_paths.add(log_path)
        rss_paths.add(rss_path)
        report = read_json(report_path)
        require(isinstance(report, dict), f"native {key} report is not an object")
        require(report.get("case") == case, f"native {key} report case changed")
        if n is None:
            require(report.get("n") in (None,), f"MCE report unexpectedly has n: {key}")
            input_sha = report.get("input_sha256")
            require(is_hex_digest(input_sha), f"MCE report lacks input SHA: {key}")
        else:
            require(report.get("n") == n, f"OPC report n changed: {key}")
            require(
                "input_sha256" not in report,
                f"OPC report must remain bound to source generator, not input SHA: {key}",
            )
        input_bytes = report.get("input_bytes")
        require(isinstance(input_bytes, int) and input_bytes > 0, f"bad input size: {key}")
        require(report.get("samples") == SAMPLES, f"bad samples: {key}")
        require(report.get("warmup") == WARMUP, f"bad warmup: {key}")
        values = timing_values(report, f"native {key}")
        p50 = statistics.median(sorted(values))
        if "p50_ns" in report:
            require(report["p50_ns"] == p50, f"bad report p50: {key}")
        outcomes = report.get("outcomes")
        require(isinstance(outcomes, list) and len(outcomes) == 1 and isinstance(outcomes[0], str), f"bad outcomes: {key}")
        if n is not None and n > 256:
            require(outcomes[0] == EXPECTED_OPC_LIMIT_ERROR, f"unexpected OPC boundary outcome: {key}")
        else:
            require(outcomes[0].startswith("ok:"), f"accepted control did not produce one ok outcome: {key}")
        rss_text = rss_path.read_text().strip()
        require(rss_text.isdigit(), f"bad RSS receipt: {key}")
        rss_kib = int(rss_text)
        require(rss_kib > 0, f"zero RSS receipt: {key}")
        binary = row.get("binary")
        require(isinstance(binary, dict), f"native {key} has no binary receipt")
        expected_binary = "mce-stream-probe" if n is None else "attribute_checks"
        require(Path(str(binary.get("path", ""))).name == expected_binary, f"wrong binary for {key}")
        item = {
            "row": row,
            "report": report,
            "report_path": report_path,
            "log_path": log_path,
            "rss_path": rss_path,
            "input_bytes": input_bytes,
            "input_sha256": report.get("input_sha256"),
            "outcomes": tuple(outcomes),
            "samples": report["samples"],
            "warmup": report["warmup"],
            "timings": values,
            "p50_ns": p50,
            "sample_max_ns": max(values),
            "rss_kib": rss_kib,
        }
        grouped.setdefault(key, []).append(item)
    require(set(grouped) == expected_set, "native case set is incomplete")
    for key, items in grouped.items():
        require(len(items) == len(LEGS), f"wrong process count for {key}")
        require(
            [item["row"].get("leg") for item in items] == list(LEGS),
            f"wrong process order for {key}",
        )
        signatures = {
            (
                item["input_bytes"],
                item["input_sha256"],
                item["outcomes"],
                item["samples"],
                item["warmup"],
            )
            for item in items
        }
        require(len(signatures) == 1, f"input or outcome changed across legs for {key}")
        if key[1] is not None:
            require(items[0]["input_sha256"] is None, f"OPC report gained input SHA: {key}")
    return grouped, rows


def differential_and_equivalence(
    runs: list[dict[str, Any]],
) -> tuple[dict[str, Any], dict[str, Any]]:
    differentials = [row for row in runs if row.get("kind") == "differential"]
    differential_legs = {
        row.get("leg", row.get("name")) for row in differentials
    }
    require(differential_legs == {"before", "after"}, "differential legs incomplete")
    reports: dict[str, dict[str, Any]] = {}
    for row in differentials:
        leg = row.get("leg", row.get("name"))
        path = artifact_path(row.get("report"), f"differential {leg} report")
        value = read_json(path)
        require(isinstance(value, dict) and isinstance(value.get("results"), dict), "invalid differential report")
        reports[leg] = value
        failures = value.get("alias_oracle_failures")
        if failures is not None:
            require(failures == [], f"alias oracle failures in {leg} differential")
    before = reports["before"]
    after = reports["after"]
    if before != after:
        before_results = before.get("results", {})
        after_results = after.get("results", {})
        changed = sorted(
            key
            for key in set(before_results) | set(after_results)
            if before_results.get(key) != after_results.get(key)
        )
        fail(f"differential changed ({len(changed)} result keys); expected exact equality")
    equivalence_rows = [row for row in runs if row.get("kind") == "equivalence"]
    require(len(equivalence_rows) == 1, "candidate equivalence row missing")
    equivalence_path = artifact_path(equivalence_rows[0].get("report"), "candidate equivalence report")
    equivalence = read_json(equivalence_path)
    require(isinstance(equivalence, dict), "candidate equivalence report is not an object")
    require(equivalence.get("mismatches") == 0, "candidate equivalence has mismatches")
    if "mismatch_examples" in equivalence:
        require(equivalence["mismatch_examples"] == [], "candidate equivalence has mismatch examples")
    result = {
        "comparisons": len(before["results"]),
        "unchanged": len(before["results"]),
        "changes": {},
        "summary": {key: value for key, value in before.items() if key != "results"},
    }
    return result, {
        "mismatches": equivalence["mismatches"],
        "report": relative_packet(equivalence_path),
    }


def analyze() -> dict[str, Any]:
    complete, runs = load_complete()
    grouped, _native = native_rows(runs)
    differential, equivalence = differential_and_equivalence(runs)
    cases: dict[str, Any] = {}
    for (case, n), items in grouped.items():
        identity = items[0]
        leg_values: dict[str, Any] = {}
        for leg in ("before", "after"):
            leg_items = [item for item in items if item["row"]["leg"] == leg]
            p50_values = [item["p50_ns"] for item in leg_items]
            rss_values = [item["rss_kib"] for item in leg_items]
            max_values = [item["sample_max_ns"] for item in leg_items]
            leg_values[leg] = {
                "processes": [
                    {
                        "order_index": item["row"]["order_index"],
                        "p50_ns": item["p50_ns"],
                        "sample_max_ns": item["sample_max_ns"],
                        "rss_kib": item["rss_kib"],
                        "report": relative_packet(item["report_path"]),
                        "log": relative_packet(item["log_path"]),
                        "rss": relative_packet(item["rss_path"]),
                    }
                    for item in leg_items
                ],
                "p50_ns": summary(p50_values),
                "rss_kib": summary(rss_values),
                "sample_max_ns": summary(max_values),
            }
        p50_ratio = leg_values["after"]["p50_ns"]["median"] / leg_values["before"]["p50_ns"]["median"]
        rss_ratio = leg_values["after"]["rss_kib"]["median"] / leg_values["before"]["rss_kib"]["median"]
        cases[case_id(case, n)] = {
            "case": case,
            "n": n,
            "input_bytes": identity["input_bytes"],
            "input_sha256": identity["input_sha256"],
            "outcomes": list(identity["outcomes"]),
            "samples": identity["samples"],
            "warmup": identity["warmup"],
            "before": leg_values["before"],
            "after": leg_values["after"],
            "p50_change_percent": 100.0 * (p50_ratio - 1.0),
            "p50_regression_flag": p50_ratio > 1.05,
            "rss_change_percent": 100.0 * (rss_ratio - 1.0),
            "rss_regression_flag": rss_ratio > 1.05,
        }
    return {
        "schema": "litchi-0777-xml-attribute-analysis-v1",
        "candidate": complete.get("candidate"),
        "base": complete.get("base"),
        "cases": cases,
        "differential": differential,
        "equivalence": equivalence,
        "limits": (
            "Three serial processes per leg and nine measured samples per process. "
            "Sample maxima are descriptive, not stable p95/p99 estimates. RSS includes "
            "whole-process startup, input generation, parsing, and report overhead."
        ),
    }


if __name__ == "__main__":
    result = analyze()
    (PACKET / "analysis.json").write_text(
        json.dumps(result, indent=2, sort_keys=True) + "\n"
    )
    print(
        json.dumps(
            {
                "cases": len(result["cases"]),
                "native_processes": EXPECTED_NATIVE_PROCESSES,
                "differential_changes": len(result["differential"]["changes"]),
                "equivalence_mismatches": result["equivalence"]["mismatches"],
            },
            sort_keys=True,
        )
    )
