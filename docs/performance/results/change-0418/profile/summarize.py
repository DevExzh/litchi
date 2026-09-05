#!/usr/bin/env python3
"""Summarize the retained 0418 CPU and PMU evidence.

The result is intentionally descriptive.  CPU percentages are whole-command
``perf`` period weights, with setup, corpus generation, warmups, children,
verification and report writing in scope.  They are not elapsed-time phase
percentages and this tool makes no speedup claim.  ``--write`` reads only
retained analysis/profile artifacts.  ``--replay`` recomputes the summary and
compares it byte-for-byte with ``summary.json``; it remains usable after raw
perf files, temporary ELFs, and worktrees have been removed, provided the
analysis text and command manifests were retained.
"""

from __future__ import annotations

import argparse
import collections
import csv
import json
import math
import re
import sys
from pathlib import Path
from typing import Any


PROFILE_DIR = Path(__file__).resolve().parent
DEFAULT_ROOT = PROFILE_DIR.parent
EXPECTED_CHANGE = 418
ROLES = ("control", "candidate")
EVENT_ORDER = (
    "cycles:u", "instructions:u", "branches:u", "branch-misses:u",
    "cache-references:u", "cache-misses:u",
)

if str(PROFILE_DIR) not in sys.path:
    sys.path.insert(0, str(PROFILE_DIR))
import analyze  # noqa: E402


class SummaryError(ValueError):
    """A retained profile artifact failed strict summary validation."""


def fail(message: str) -> None:
    raise SummaryError(message)


def _mapping(value: Any, label: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        fail(f"{label} must be an object")
    return value


def _integer(value: Any, label: str, minimum: int = 0) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < minimum:
        fail(f"{label} must be an integer >= {minimum}")
    return value


def _finite_number(value: Any, label: str) -> float | int:
    if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(float(value)):
        fail(f"{label} must be finite numeric")
    return value


def _read_json_bytes(data: bytes, label: str) -> Any:
    try:
        value = json.loads(
            data.decode("utf-8"), object_pairs_hook=analyze._pairs,
            parse_constant=analyze._constant,
        )
    except (UnicodeError, json.JSONDecodeError, analyze.AnalysisError) as error:
        fail(f"{label}: invalid strict JSON: {error}")
    analyze._finite(value, label)
    return value


def _canonical(value: Any) -> bytes:
    return analyze.canonical(value)


def _file_artifact(root: Path, relative: str, label: str) -> tuple[bytes, dict[str, Any]]:
    try:
        return analyze.read_artifact(root, relative, label)
    except analyze.AnalysisError as error:
        fail(str(error))
    raise AssertionError("unreachable")


def _command_records(root: Path, manifest: dict[str, Any]) -> list[dict[str, Any]]:
    commands = manifest.get("commands")
    if not isinstance(commands, list):
        fail("analysis-commands.json.commands must be a list")
    seen: set[tuple[str, str, str]] = set()
    result: list[dict[str, Any]] = []
    for index, command in enumerate(commands):
        command = _mapping(command, f"analysis.commands[{index}]")
        role, selector, kind = command.get("role"), command.get("selector"), command.get("kind")
        if not all(isinstance(value, str) and value for value in (role, selector, kind)):
            fail(f"analysis.commands[{index}] has invalid role/selector/kind")
        key = (role, selector, kind)
        if key in seen:
            fail(f"duplicate analysis command {key}")
        seen.add(key)
        if command.get("exit_code") != 0:
            fail(f"analysis command {key} did not complete successfully")
        argv = command.get("argv")
        if not isinstance(argv, list) or not argv or not all(isinstance(value, str) for value in argv):
            fail(f"analysis command {key} has invalid argv")
        output = command.get("output")
        stderr = command.get("stderr")
        if not isinstance(output, str) or not isinstance(stderr, str):
            fail(f"analysis command {key} has missing output paths")
        output_data, output_info = _file_artifact(root, output, f"analysis/{key}/output")
        stderr_data, stderr_info = _file_artifact(root, stderr, f"analysis/{key}/stderr")
        if command.get("output_sha256") != output_info["sha256"] or command.get("output_bytes") != output_info["bytes"]:
            fail(f"analysis command {key} output hash/size mismatch")
        if command.get("stderr_sha256") != stderr_info["sha256"] or command.get("stderr_bytes") != stderr_info["bytes"]:
            fail(f"analysis command {key} stderr hash/size mismatch")
        command = dict(command)
        command["_output_data"] = output_data
        command["_stderr_data"] = stderr_data
        result.append(command)
    return result


def _validate_analysis(context: dict[str, Any]) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    root = context["root"]
    path = root / "profile" / "analysis-commands.json"
    manifest = _mapping(analyze.load_json(path), "analysis-commands")
    if manifest.get("schema_version") != 1 or manifest.get("change") != EXPECTED_CHANGE or manifest.get("status") != "complete":
        fail("analysis-commands.json is not a complete 0418 manifest")
    expected_hashes = {
        "protocol": context["hashes"]["protocol"],
        "build_identity": context["hashes"]["build"],
        "capture": context["hashes"]["capture"],
        "profile": context["hashes"]["profile"],
    }
    for key, expected in expected_hashes.items():
        if _mapping(manifest.get(key), f"analysis.{key}").get("sha256") != expected:
            fail(f"analysis {key} binding is stale")
    commands = _command_records(root, manifest)
    expected_pairs = {(role, selector) for role in ROLES for selector in context["selectors"]}
    expected_kinds = {"perf-header", "perf-self", "perf-children", "perf-script"}
    for pair in expected_pairs:
        kinds = {item["kind"] for item in commands if (item["role"], item["selector"]) == pair}
        if kinds != expected_kinds:
            fail(f"analysis views incomplete for {pair}: {sorted(kinds)}")
    comments_path = root / "profile" / "elf-comments.json"
    comments_manifest = _mapping(analyze.load_json(comments_path), "elf-comments")
    entries = comments_manifest.get("entries")
    if not isinstance(entries, list) or len(entries) != 4:
        fail("elf-comments.json must retain all four owned binaries")
    comment_keys = set()
    for index, entry in enumerate(entries):
        entry = _mapping(entry, f"elf-comments.entries[{index}]")
        key = (entry.get("role"), entry.get("mode"))
        comment_keys.add(key)
        if entry.get("exit_code") != 0 or not isinstance(entry.get("stdout"), str):
            fail(f"elf-comments.entries[{index}] is incomplete")
        identity = _mapping(entry.get("binary"), f"elf-comments.entries[{index}].binary")
        role, mode = key
        if role not in ROLES or mode not in ("normal", "allocator"):
            fail(f"invalid ELF comment role/mode {key}")
        expected = context["binaries"][role][mode]["identity"]
        if identity.get("sha256") != expected.get("sha256") or identity.get("bytes") != expected.get("bytes"):
            fail(f"ELF comment identity differs for {role}/{mode}")
        if analyze.sha256_bytes(entry["stdout"].encode()) != entry.get("stdout_sha256"):
            fail(f"ELF comment stdout hash differs for {role}/{mode}")
        if analyze.sha256_bytes(entry.get("stderr", "").encode()) != entry.get("stderr_sha256"):
            fail(f"ELF comment stderr hash differs for {role}/{mode}")
    if comment_keys != {(role, mode) for role in ROLES for mode in ("normal", "allocator")}:
        fail("ELF comment manifest does not cover exactly four owned binaries")
    # The CPU binding is rechecked without touching raw perf data or binaries.
    bindings = []
    for role in ROLES:
        for selector in context["selectors"]:
            record = analyze._record_for(context, role, "cpu-record", selector)
            bindings.append(analyze._formal_binding(context, role, selector, record))
            data_descriptor = analyze._artifact_descriptor(record, "data", f"profile/{role}/{selector}")
            if data_descriptor.get("sha256") != next(
                item["input"]["sha256"] for item in commands
                if item["role"] == role and item["selector"] == selector and item["kind"] == "perf-script"
            ):
                fail(f"analysis perf input hash differs for {role}/{selector}")
    return manifest, commands


HEADER_RE = re.compile(
    r"^\s*(?P<comm>.+?)\s+(?P<pid>[0-9]+)"
    r"(?:(?:/(?P<tid_slash>[0-9]+))|(?:\s+(?P<tid_space>[0-9]+)(?=\s+[0-9]+(?:\.[0-9]+)?:)))?"
    r"(?:\s+\[[0-9]+\])?\s+(?P<time>[0-9]+(?:\.[0-9]+)?):\s+"
    r"(?P<period>[0-9]+)\s+(?P<event>cycles:u):\s*$"
)
FRAME_RE = re.compile(
    r"^\s*(?P<ip>(?:0x)?[0-9a-fA-F]+)\s+(?P<symbol>.*?)\s+"
    r"\((?P<dso>.*)\)\s*$"
)
FRAME_FALLBACK_RE = re.compile(r"^\s*(?P<ip>(?:0x)?[0-9a-fA-F]+)\s+(?P<symbol>\S.*?)\s*$")


def _parse_perf_script(data: bytes, label: str) -> list[dict[str, Any]]:
    try:
        text = data.decode("utf-8")
    except UnicodeDecodeError as error:
        fail(f"{label}: perf script is not UTF-8: {error}")
    blocks: list[dict[str, Any]] = []
    current: dict[str, Any] | None = None

    def finish() -> None:
        nonlocal current
        if current is None:
            return
        if not current["frames"]:
            fail(f"{label}: sample has no callchain")
        blocks.append(current)
        current = None

    for raw_line in text.splitlines() + [""]:
        line = raw_line.rstrip("\r")
        if not line.strip():
            finish()
            continue
        if line.lstrip().startswith("#"):
            continue
        header = HEADER_RE.fullmatch(line)
        if header is not None:
            if current is not None:
                finish()
            period = int(header.group("period"))
            if period <= 0:
                fail(f"{label}: non-positive sample period")
            current = {
                "comm": header.group("comm"), "pid": int(header.group("pid")),
                "tid": int(header.group("tid_slash") or header.group("tid_space"))
                if header.group("tid_slash") or header.group("tid_space") else None,
                "time": header.group("time"),
                "period": period, "event": header.group("event"), "frames": [],
            }
            continue
        frame = FRAME_RE.fullmatch(line)
        if frame is None:
            frame = FRAME_FALLBACK_RE.fullmatch(line)
            if frame is None:
                fail(f"{label}: unparsed perf script line {line!r}")
            dso = "[unknown]"
        else:
            dso = frame.group("dso")
        if current is None:
            fail(f"{label}: frame appeared outside a sample")
        symbol = frame.group("symbol").strip()
        if not symbol:
            symbol = "[unknown]"
        current["frames"].append({"symbol": symbol, "dso": dso})
    if not blocks:
        fail(f"{label}: no perf samples")
    return blocks


def _counter_rows(counter: collections.Counter[tuple[str, str]], total: int) -> list[dict[str, Any]]:
    rows = []
    for (symbol, dso), period in sorted(
        counter.items(), key=lambda item: (-item[1], item[0])
    ):
        rows.append({
            "symbol": symbol, "dso": dso, "period": period,
            "percent": round(period * 100.0 / total, 9),
        })
    return rows


def _match_class(frame: tuple[str, str], family: str) -> bool:
    symbol = frame[0].lower()
    if family == "deflate":
        # ``deflate`` is intentionally matched as a word fragment: Rust
        # symbols contain names such as ``DeflateEncoder``.  It cannot match
        # ``inflate``/``inflate2``.  Keep flate2 and zlib scoped to their
        # compression namespaces rather than matching arbitrary ``flate``.
        return (
            "deflate" in symbol
            or "zlib_rs::deflate" in symbol
            or "flate2::mem::compress" in symbol
        )
    if family == "sha":
        # Avoid false positives from symbols such as ``shared_strings`` and
        # ``shape``.  These are the actual digest namespaces in this build.
        return bool(re.search(r"(?:^|:)\b(?:sha1|sha2|sha256|digest)::", symbol))
    raise AssertionError(family)


def _weighted(blocks: list[dict[str, Any]]) -> dict[str, Any]:
    leaf: collections.Counter[tuple[str, str]] = collections.Counter()
    inclusive: collections.Counter[tuple[str, str]] = collections.Counter()
    events = collections.Counter()
    # Family-inclusive attribution is a sample property: if a callchain has
    # any frame from a family, charge that sample's period once.  Summing the
    # per-symbol inclusive counters would charge one sample once per matching
    # ancestor and can exceed 100 percent.
    family_sample_period = collections.Counter()
    for block in blocks:
        period = block["period"]
        events[block["event"]] += period
        frames = [(frame["symbol"], frame["dso"]) for frame in block["frames"]]
        leaf[frames[0]] += period
        for frame in set(frames):
            inclusive[frame] += period
        for family in ("deflate", "sha"):
            if any(_match_class(frame, family) for frame in set(frames)):
                family_sample_period[family] += period
    total = sum(leaf.values())
    if total <= 0:
        fail("perf script period total is zero")
    unresolved = sum(period for (symbol, dso), period in leaf.items() if "[unknown]" in symbol.lower() or "[unknown]" in dso.lower())
    result: dict[str, Any] = {
        "sample_blocks": len(blocks),
        "total_period": total,
        "events": dict(sorted(events.items())),
        "leaf_count": len(leaf),
        "inclusive_symbol_count": len(inclusive),
        "leaf": _counter_rows(leaf, total),
        "inclusive": _counter_rows(inclusive, total),
        "unresolved_leaf_period": unresolved,
        "unresolved_leaf_percent": round(unresolved * 100.0 / total, 9),
    }
    families: dict[str, Any] = {}
    for family in ("deflate", "sha"):
        leaf_period = sum(period for frame, period in leaf.items() if _match_class(frame, family))
        inclusive_period = family_sample_period[family]
        families[family] = {
            "leaf_period": leaf_period,
            "leaf_percent": round(leaf_period * 100.0 / total, 9),
            "inclusive_period": inclusive_period,
            "inclusive_percent": round(inclusive_period * 100.0 / total, 9),
            "inclusive_method": "one period per sample containing any family frame",
        }
    result["families"] = families
    result["family_matchers"] = {
        "deflate": "symbol contains deflate, zlib_rs::deflate, or flate2::mem::compress",
        "sha": "symbol contains namespace sha1::, sha2::, sha256::, or digest::",
    }
    return result


LOST_RE = re.compile(r"#\s*Total Lost Samples:\s*([0-9][0-9,]*)")
COUNT_RE = re.compile(r"Event count \(approx\.\):\s*([0-9][0-9,]*)")


def _validate_report_text(self_data: bytes, weighted: dict[str, Any], label: str) -> dict[str, Any]:
    text = self_data.decode("utf-8", errors="replace")
    lost_match = LOST_RE.search(text)
    count_match = COUNT_RE.search(text)
    if lost_match is None:
        fail(f"{label}: self report lacks Total Lost Samples")
    if count_match is None:
        fail(f"{label}: self report lacks Event count (approx.)")
    lost = int(lost_match.group(1).replace(",", ""))
    event_count = int(count_match.group(1).replace(",", ""))
    period = weighted["total_period"]
    if event_count != period:
        fail(f"{label}: perf report event count {event_count} differs from script period {period}")
    return {
        "lost_samples": lost,
        "event_count_approx": event_count,
        "event_count_matches_period": True,
        "report_sha256": analyze.sha256_bytes(self_data),
        "report_bytes": len(self_data),
    }


def _cpu_summary(context: dict[str, Any], commands: list[dict[str, Any]]) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for role in ROLES:
        for selector in context["selectors"]:
            chosen = {
                item["kind"]: item for item in commands
                if item["role"] == role and item["selector"] == selector
            }
            script = chosen["perf-script"]
            self_report = chosen["perf-self"]
            blocks = _parse_perf_script(script["_output_data"], f"{role}/{selector}/perf-script")
            weighted = _weighted(blocks)
            report_quality = _validate_report_text(self_report["_output_data"], weighted, f"{role}/{selector}/perf-self")
            result.append({
                "role": role, "selector": selector,
                "boundary": {"samples": 20, "warmups": 3, "cpu": 2, "workers": 1},
                "input": script["input"],
                "views": {
                    kind: {"path": item["output"], "bytes": item["output_bytes"], "sha256": item["output_sha256"]}
                    for kind, item in chosen.items()
                },
                "command_argv": script["argv"],
                "command_cwd": script["cwd"],
                "period_scope": "whole command and descendants",
                "report_quality": report_quality,
                "weighted": weighted,
            })
    return result


def _parse_counter(value: str) -> tuple[str, int | float | None]:
    value = value.strip()
    if value in ("<not supported>", "<not counted>", "<not enabled>", "<failure>"):
        return value[1:-1].replace(" ", "_"), None
    try:
        number: int | float
        if re.fullmatch(r"[+-]?[0-9]+", value):
            number = int(value)
        else:
            number = float(value)
        if not math.isfinite(float(number)):
            fail(f"PMU counter is non-finite: {value!r}")
        return "measured", number
    except ValueError:
        return "unparsed", None


def _parse_stat(data: bytes, label: str) -> dict[str, Any]:
    text = data.decode("utf-8", errors="replace")
    raw_lines = [line for line in text.splitlines() if line.strip() and not line.lstrip().startswith("#")]
    if not raw_lines:
        fail(f"{label}: PMU stat CSV is empty")
    rows: list[dict[str, Any]] = []
    for line in raw_lines:
        fields = next(csv.reader([line]))
        if len(fields) < 5:
            fail(f"{label}: PMU row has fewer than five fields: {line!r}")
        value, unit, event, runtime, running = (field.strip() for field in fields[:5])
        status, number = _parse_counter(value)
        runtime_status, runtime_value = _parse_counter(runtime)
        running_status, running_value = _parse_counter(running)
        rows.append({
            "raw": line, "fields": fields, "value_raw": value, "unit": unit,
            "event": event, "status": status, "value": number,
            "runtime_raw": runtime, "runtime_status": runtime_status,
            "runtime_ns": runtime_value, "running_raw": running,
            "running_status": running_status, "running_percent": running_value,
        })
    # On this host the cache-reference alias reports zero while the paired
    # cache-miss event reports positive values.  Preserve both raw counters,
    # but do not present the zero as a measured hardware reference count.
    by_event = {row["event"]: row for row in rows}
    references = by_event.get("cache-references:u")
    misses = by_event.get("cache-misses:u")
    if references is not None and references["status"] == "measured" and references["value"] == 0:
        if misses is not None and misses["status"] == "measured" and isinstance(misses["value"], (int, float)) and misses["value"] > 0:
            references["status"] = "unvalidated_alias_zero"
    statuses = {"measured": 0, "not_supported": 0, "not_counted": 0,
                "unvalidated_alias_zero": 0, "other": 0}
    for row in rows:
        status = row["status"]
        if status in statuses and status != "other":
            statuses[status] += 1
        else:
            statuses["other"] += 1
    return {
        "raw_sha256": analyze.sha256_bytes(data), "raw_bytes": len(data),
        "raw_lines": raw_lines, "events": rows,
        "availability": statuses,
        "cache_interpretation": (
            "withheld: cache-references:u is an unvalidated zero alias while "
            "cache-misses:u is positive"
            if references is not None and references["status"] == "unvalidated_alias_zero"
            else "descriptive only; no hardware cache claim"
        ),
    }


def _pmu_summary(context: dict[str, Any]) -> dict[str, Any]:
    protocol_pmu = _mapping(context["protocol"].get("pmu"), "protocol.pmu")
    expected_selector = protocol_pmu.get("selector")
    expected_events = protocol_pmu.get("events")
    if not isinstance(expected_selector, str) or not isinstance(expected_events, list):
        fail("protocol PMU selector/events are invalid")
    per_role: dict[str, Any] = {}
    for role in ROLES:
        record = analyze._record_for(context, role, "pmu-stat", expected_selector)
        analyze._validate_run_identity(context, record, role, "normal", f"PMU/{role}")
        if record.get("samples") != 20 or record.get("warmups") != 3:
            fail(f"PMU {role} boundary must be 20 samples/3 warmups")
        stat_data, stat_info = analyze._artifact_from_record(context["root"], record, "stat", f"PMU/{role}")
        report_data, _ = analyze._artifact_from_record(context["root"], record, "report", f"PMU/{role}")
        catalog_data, _ = analyze._artifact_from_record(context["root"], record, "catalog", f"PMU/{role}")
        report = _mapping(_read_json_bytes(report_data, f"PMU/{role}/report"), f"PMU/{role}/report")
        catalog = _mapping(_read_json_bytes(catalog_data, f"PMU/{role}/catalog"), f"PMU/{role}/catalog")
        formal = analyze._capture_record(context, role, expected_selector)
        analyze._validate_run_identity(context, formal, role, "normal", f"formal PMU/{role}")
        formal_report_data, _ = analyze._artifact_from_record(context["root"], formal, "report", f"formal PMU/{role}")
        formal_catalog_data, _ = analyze._artifact_from_record(context["root"], formal, "catalog", f"formal PMU/{role}")
        formal_report = _mapping(_read_json_bytes(formal_report_data, f"formal PMU/{role}/report"), "formal PMU report")
        formal_catalog = _mapping(_read_json_bytes(formal_catalog_data, f"formal PMU/{role}/catalog"), "formal PMU catalog")
        analyze._validate_report_pair(
            context, role, expected_selector, report, catalog, formal_report,
            formal_catalog, 20, 3, f"PMU binding/{role}",
        )
        parsed = _parse_stat(stat_data, f"PMU/{role}/stat")
        actual_events = [row["event"] for row in parsed["events"]]
        if actual_events != expected_events:
            fail(f"PMU/{role}: raw event order differs from protocol: {actual_events!r}")
        per_role[role] = {
            "record": {
                "argv": record.get("argv"), "cwd": record.get("cwd"),
                "binary_sha256": record.get("binary_sha256"),
                "source_revision": record.get("source", {}).get("revision"),
                "samples": record.get("samples"), "warmups": record.get("warmups"),
            },
            "stat": parsed,
            "report": {"sha256": analyze.sha256_bytes(report_data), "bytes": len(report_data)},
            "catalog": {"sha256": analyze.sha256_bytes(catalog_data), "bytes": len(catalog_data)},
            "formal_output_sha256": formal_report.get("results", [{}])[0].get("output_sha256"),
        }
    paired: dict[str, Any] = {}
    by_event = {
        role: {row["event"]: row for row in per_role[role]["stat"]["events"]}
        for role in ROLES
    }
    for event in expected_events:
        pair = {}
        for role in ROLES:
            row = by_event[role].get(event)
            if row is None:
                fail(f"PMU {role} has no {event} row")
            pair[role] = {
                "status": row["status"], "value": row["value"],
                "runtime_ns": row["runtime_ns"], "running_percent": row["running_percent"],
                "raw": row["raw"], "fields": row["fields"],
            }
        paired[event] = pair
    return {
        "selector": expected_selector, "events": expected_events,
        "scope": protocol_pmu.get("scope"), "roles": per_role, "paired": paired,
        "cache_metrics_claim": "withheld: cache-references:u zero is an unvalidated alias",
        "interpretation": "availability, runtime and multiplexing are retained descriptively; no IPC or cache-miss claim is made",
    }


def _summary(context: dict[str, Any], analysis_manifest: dict[str, Any], commands: list[dict[str, Any]]) -> dict[str, Any]:
    cpu = _cpu_summary(context, commands)
    pmu = _pmu_summary(context)
    root = context["root"]
    analysis_path = root / "profile" / "analysis-commands.json"
    comments_path = root / "profile" / "elf-comments.json"
    return {
        "schema_version": 1, "change": EXPECTED_CHANGE, "status": "complete",
        "identity": {
            key: {"path": analyze.relpath(context["paths"][key], root), "sha256": context["hashes"][key]}
            for key in ("protocol", "build", "capture", "profile", "gate")
        },
        "analysis": {
            "path": analyze.relpath(analysis_path, root), "sha256": analyze.sha256_file(analysis_path),
            "elf_comments_path": analyze.relpath(comments_path, root),
            "elf_comments_sha256": analyze.sha256_file(comments_path),
            "commands": [
                {"role": item["role"], "selector": item["selector"], "kind": item["kind"],
                 "argv": item["argv"], "cwd": item["cwd"], "output": item["output"],
                 "output_sha256": item["output_sha256"], "stderr": item["stderr"],
                 "stderr_sha256": item["stderr_sha256"]}
                for item in commands
            ],
        },
        "boundary": {"samples": 20, "warmups": 3, "cpu": 2, "workers": 1,
                     "profile_event": "cycles:u", "profile_frequency": 999,
                     "call_graph": "fp,127"},
        "source_binary_gates": {
            "source_revisions": {role: context["build"]["roles"][role]["source"]["revision"] for role in ROLES},
            "binary_hashes": {
                role: {mode: context["binaries"][role][mode]["identity"]["sha256"] for mode in ("normal", "allocator")}
                for role in ROLES
            },
            "preflight_validation": "pass",
            "capture_status": context["capture"].get("status"),
            "profile_status": context["profile"].get("status"),
        },
        "cpu": cpu,
        "pmu": pmu,
        "scope": "Whole profiled command and descendants. Period weights include setup, corpus generation, warmups, timed children, verification and report writing; they do not provide phase-local elapsed attribution.",
        "claims": {"performance_speedup": "none", "phase_latency": "none", "allocator_elapsed_comparison": "none"},
        "replay": {"deterministic_json": True, "compressed_sidecars": [".zst", ".zstd", ".gz"], "raw_profile_replay_requires_elfs": False},
    }


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=DEFAULT_ROOT, help="change-0418 result directory")
    action = parser.add_mutually_exclusive_group(required=True)
    action.add_argument("--write", action="store_true", help="write summary.json")
    action.add_argument("--replay", action="store_true", help="recompute and compare summary.json")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    root = args.root.resolve()
    try:
        # Summary/replay intentionally does not hash temporary binaries.  It
        # checks their recorded identities and the retained ELF comment proof,
        # allowing this step after build/worktree cleanup.
        context = analyze._load_context(root, require_binaries=False)
        analysis_manifest, commands = _validate_analysis(context)
        result = _summary(context, analysis_manifest, commands)
        output = root / "profile" / "summary.json"
        if args.replay:
            existing = analyze.load_json(output)
            if _canonical(existing) != _canonical(result):
                fail("summary replay differs from retained summary.json")
            print("0418 profile summary replay passed")
            return 0
        if output.exists():
            fail(f"refusing to overwrite existing summary: {output}")
        analyze.write_json(output, result)
        print(f"0418 profile summary complete: {output}")
        return 0
    except (SummaryError, analyze.AnalysisError, OSError) as error:
        print(f"0418 profile summary failed: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
