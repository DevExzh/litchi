#!/usr/bin/env python3
"""Small, fail-closed support module for the 0971 instruction profile.

The module only validates custody and parses retained profiler text.  It does
not build or launch the Office harness.  All writes in the runner use O_EXCL.
"""
from __future__ import annotations

import hashlib
import importlib.util
import json
import math
import re
from pathlib import Path
from typing import Any

SCHEMA = "litchi.performance.0971.instruction_profile.v2"
PLAN_SCHEMA = "litchi.performance.0971.plan.v2"
GO_SCHEMA = "litchi.performance.0971.root_go.v1"
PACKET_SCHEMA = "litchi.performance.0971.child_packet.v2"
HELPER_SCHEMA = "litchi.performance.0971.helper_seal.v2"
MARKER = "doc_retention_probe_0942_cfb_profile::profiled_measured_case"
BINARY_SHA = "0db34357a1e725dc397a4fd895532433db1214c5862b7bc96bf23f21c0396ecb"
BINARY_BYTES = 47909224
BINARY_PATH = "/home/zhuhe/litchi-performance-0967/builds/target-baseline-cfb/release/doc_retention_probe_0942_cfb_profile"
BINARY_BUILD_ID = "e5c457384eeebd057932c444efda77618f137399"
FIXTURE_SHA = "41cc0cd8f2d7390266f844f5bddbcd29d08ac5338abc33832d9c9a6ed5d0f85d"
FIXTURE_BYTES = 335360
FIXTURE_PATH = "/home/zhuhe/litchi-performance-0969/qa-core/test-data/ole/doc/FloatingPictures.doc"
CPU = 31
WARMUPS = 5
SAMPLES = 80
CALLS = WARMUPS + SAMPLES
PERIOD = 5_000_000
CHILDREN = 12
PERF_SHA = "174229fb48b8557aafa821dd828d883dce47c35c31067e1f31e4a01a91c6c98b"
LAUNCH_SCHEMA = "litchi.performance.0970.launch.v2"
LAUNCH_SHA = "3b5b79351883a9b4efb680f47498d8e10bdf77d131a60d11c1ecb54ad95430f5"
VALIDATOR_PATH = "/home/zhuhe/litchi-performance-0969/tools/protocol_common_v4.py"
VALIDATOR_SHA = "5ce0f2ac08fc7909df054a4cc2d5137b3279af4cd726acf5bfb4558d97b05c5c"
SOURCE_MANIFEST_SHA = "f1818af5c306fb6ea01b59ca8df29710e773386b246071202bd4aa896f4e8e2b"
SOURCE_FILES_SHA = "99e65d43e0feed89133eec60b095130d68058a071a02e8a15b05dea01209bb6c"
SOURCE_ROOT = "/home/zhuhe/litchi-performance-0969/source-baseline"
SOURCE_PROOF_SHA = "20e0f9fad68da357e2278157a8d1514c8b9b226993e16a91656fddc7ff38edd6"
PROPOSAL_SHA = "c099fb6dc0572bbb09e1d911445e6012988f64a3bb3110be69377861ef2ef437"
CAPABILITY_SHA = "50bc1b5c3253ed44cc3b1516e258a6028db40a3f049336814264365b539141f2"
SOURCE_VARIATION_SHA = "86487399a1ffa2dbd51ba7a2aa7b02adf11d45e4b556ee3b48a618e2c65d5217"
ADR_RECEIPT_SHA = "08b1a8d25be0ee9777221d6ce409c7da93a892981a6dadc7dd84779b4c7b471b"


class ProtocolError(ValueError):
    pass


def _duplicate_hook(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    out: dict[str, Any] = {}
    for key, value in pairs:
        if key in out:
            raise ProtocolError(f"duplicate JSON key: {key}")
        out[key] = value
    return out


def load_json(path: Path) -> dict[str, Any]:
    try:
        value = json.loads(path.read_text(encoding="utf-8"), object_pairs_hook=_duplicate_hook)
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        raise ProtocolError(f"invalid JSON {path}: {exc}") from exc
    if not isinstance(value, dict):
        raise ProtocolError(f"{path}: expected JSON object")
    return value


def write_json_exclusive(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        with path.open("x", encoding="utf-8") as handle:
            json.dump(value, handle, indent=2, sort_keys=True, allow_nan=False)
            handle.write("\n")
            handle.flush()
    except FileExistsError as exc:
        raise ProtocolError(f"refusing to overwrite {path}") from exc


def write_text_exclusive(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        with path.open("x", encoding="utf-8") as handle:
            handle.write(text)
            handle.flush()
    except FileExistsError as exc:
        raise ProtocolError(f"refusing to overwrite {path}") from exc


def sha256(path: Path) -> str:
    if not path.is_file():
        raise ProtocolError(f"missing file: {path}")
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def file_binding(path: str | Path) -> dict[str, Any]:
    # Preserve the explicitly supplied custody path.  ``resolve()`` would
    # collapse the retained 0969 fixture's hard link into the checkout path,
    # which would weaken the external source/fixture binding.
    p = Path(path).absolute()
    return {"path": str(p), "bytes": p.stat().st_size, "sha256": sha256(p)}


def check_binding(row: Any, label: str, *, expected_path: str | None = None, expected_sha: str | None = None) -> Path:
    if not isinstance(row, dict):
        raise ProtocolError(f"{label}: binding missing")
    raw = row.get("path")
    if not isinstance(raw, str) or not raw:
        raise ProtocolError(f"{label}: path missing")
    path = Path(raw).absolute()
    actual = sha256(path)
    wanted = expected_sha or row.get("sha256")
    if not isinstance(wanted, str) or actual != wanted:
        raise ProtocolError(f"{label}: SHA mismatch ({actual} != {wanted})")
    if expected_path is not None and str(path) != str(Path(expected_path).absolute()):
        raise ProtocolError(f"{label}: path mismatch")
    if row.get("bytes") is not None and row.get("bytes") != path.stat().st_size:
        raise ProtocolError(f"{label}: byte count mismatch")
    return path


def load_validator() -> Any:
    path = Path(VALIDATOR_PATH)
    if sha256(path) != VALIDATOR_SHA:
        raise ProtocolError("0969 report validator changed")
    spec = importlib.util.spec_from_file_location("litchi_protocol_common_v4", path)
    if spec is None or spec.loader is None:
        raise ProtocolError("cannot import retained report validator")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def validate_report(report: dict[str, Any]) -> None:
    """Validate the raw report against the sealed cell, before delegation.

    The retained 0969 validator is intentionally reused for its 20 correctness
    gates, but its caller-supplied cell is never derived from report fields.
    """
    if not isinstance(report, dict):
        raise ProtocolError("probe report must be an object")
    if report.get("case") != "docfloat" or report.get("edits") != 1 or report.get("retention") != "default_8mib":
        raise ProtocolError("probe report does not match fixed docfloat/e1/default cell")
    if report.get("source_sha256") != FIXTURE_SHA or report.get("source_bytes") != FIXTURE_BYTES:
        raise ProtocolError("probe input fixture identity mismatch")
    if report.get("warmups") != WARMUPS or report.get("samples_requested") != SAMPLES:
        raise ProtocolError("probe warmup/sample schedule mismatch")
    elapsed = report.get("elapsed_ns")
    rows = report.get("samples")
    if not isinstance(elapsed, list) or len(elapsed) != SAMPLES or not isinstance(rows, list) or len(rows) != SAMPLES:
        raise ProtocolError("probe elapsed/sample row count mismatch")
    if any(type(value) is not int or value <= 0 for value in elapsed):
        raise ProtocolError("probe elapsed_ns must be positive integers")
    for index, row in enumerate(rows):
        if not isinstance(row, dict) or row.get("index") != index or row.get("elapsed_ns") != elapsed[index]:
            raise ProtocolError(f"probe sample {index}: index/elapsed identity mismatch")
    fixture = {"sha256": FIXTURE_SHA, "bytes": FIXTURE_BYTES}
    validator = load_validator()
    validator.check_probe_report(report, cell={"case": "docfloat", "edits": 1, "retention": "default"}, fixture=fixture, allocator=False)


def _proposal_retention(value: Any) -> str:
    if value == "default_8mib":
        return "default"
    return value


_HEADER = re.compile(
    r"^\s*(?:(?P<comm>\S+)\s+)?(?P<pid>\d+)(?:/\d+)?"
    r"(?:\s+\[(?P<cpu>\d+)\])?\s+(?P<time>\d+(?:\.\d+)?):\s+"
    r"(?P<period>\d+)\s+(?P<event>[^\n]+):\s*$"
)
_FRAME = re.compile(r"^\s*(?P<ip>[0-9a-fA-F]+)\s+(?P<symbol>.*?)\s+\((?P<dso>.*)\)\s*$")


def _frame(line: str) -> dict[str, Any] | None:
    match = _FRAME.match(line)
    if match is None:
        return None
    symbol = match.group("symbol").strip()
    base, _, offset = symbol.partition("+")
    return {
        "ip": match.group("ip").lower(),
        "symbol": symbol,
        "symbol_base": base.strip(),
        "offset": offset or None,
        "dso": match.group("dso").strip(),
        "resolved": base.strip() not in {"[unknown]", "unknown", ""} and match.group("dso").strip() not in {"[unknown]", ""},
    }


def parse_perf_script(
    text: str,
    *,
    marker: str = MARKER,
    marker_dso: str | None = None,
    expected_pid: int | None = None,
    expected_cpu: int | None = CPU,
    require_event: str = "instructions:u",
) -> dict[str, Any]:
    """Parse perf's physical, no-inline stack form.

    Header/blank lines are accepted; any nonblank line that is neither a
    sample header nor an indented frame is rejected.  Lost-event records are
    retained as an explicit count and never converted to zero silently.
    """
    samples: list[dict[str, Any]] = []
    current: dict[str, Any] | None = None
    lost = 0
    trunc = 0
    for raw in text.splitlines():
        line = raw.rstrip("\n")
        stripped = line.strip()
        if not stripped:
            if current is not None:
                samples.append(current)
                current = None
            continue
        low = stripped.lower()
        lost_match = re.search(r"(?:lost|dropped)[^0-9]*(\d+)", low)
        if lost_match and ("lost" in low or "drop" in low):
            lost += int(lost_match.group(1))
            continue
        if stripped in {"...", "[stack truncated]", "[truncated]"} or "stack truncated" in low:
            trunc += 1
            continue
        header = _HEADER.match(line)
        if header:
            if current is not None:
                samples.append(current)
            event = header.group("event").strip()
            if event != require_event:
                raise ProtocolError(f"unexpected perf event {event!r}")
            current = {
                "pid": int(header.group("pid")),
                "cpu": int(header.group("cpu")) if header.group("cpu") is not None else None,
                "time": float(header.group("time")),
                "period": int(header.group("period")),
                "event": event,
                "frames": [],
            }
            if current["period"] <= 0:
                raise ProtocolError("perf sample period must be a positive integer")
            if expected_pid is not None and current["pid"] != expected_pid:
                raise ProtocolError(f"perf sample PID mismatch: {current['pid']} != {expected_pid}")
            if expected_cpu is not None and current["cpu"] != expected_cpu:
                raise ProtocolError(f"perf sample CPU mismatch: {current['cpu']} != {expected_cpu}")
            continue
        if current is not None and line[:1].isspace():
            parsed = _frame(line)
            if parsed is None:
                raise ProtocolError(f"malformed perf frame: {line!r}")
            current["frames"].append(parsed)
            continue
        raise ProtocolError(f"malformed perf script line: {line!r}")
    if current is not None:
        samples.append(current)
    if not samples and lost == 0:
        raise ProtocolError("perf script has no samples or loss record")
    for sample in samples:
        for frame in sample["frames"]:
            if frame["symbol_base"] == marker and marker_dso is not None and frame["dso"] != marker_dso:
                raise ProtocolError(f"marker symbol resolved in unexpected DSO: {frame['dso']!r}")
    marker_samples = [sample for sample in samples if any(frame["symbol_base"] == marker and (marker_dso is None or frame["dso"] == marker_dso) for frame in sample["frames"])]
    marker_period = sum(sample["period"] for sample in marker_samples)
    total_period = sum(sample["period"] for sample in samples)
    unresolved = sum(sum(not frame["resolved"] for frame in sample["frames"]) for sample in samples)
    for sample in samples:
        if not sample["frames"]:
            trunc += 1
    owners: dict[str, int] = {}
    inclusive: dict[str, int] = {}
    deepest: dict[str, int] = {}
    root_owner: dict[str, int] = {}
    for sample in marker_samples:
        resolved = [frame for frame in sample["frames"] if frame["resolved"]]
        if resolved:
            owner = resolved[0]["symbol_base"]
            owners[owner] = owners.get(owner, 0) + sample["period"]
            # perf's physical stack is leaf/deepest first; the first resolved
            # frame is therefore the exclusive deepest owner.
            deepest[resolved[0]["symbol_base"]] = deepest.get(resolved[0]["symbol_base"], 0) + sample["period"]
            root_owner[resolved[-1]["symbol_base"]] = root_owner.get(resolved[-1]["symbol_base"], 0) + sample["period"]
        # Inclusive coverage is per owner per sample.  Recursive/repeated
        # frames therefore do not double-count one sample's period; different
        # owners intentionally overlap because this is coverage, not a sum.
        for owner in {frame["symbol_base"] for frame in sample["frames"] if frame["resolved"]}:
            inclusive[owner] = inclusive.get(owner, 0) + sample["period"]
    return {
        "sample_count": len(samples),
        "total_period": total_period,
        "marker": marker,
        "marker_sample_count": len(marker_samples),
        "marker_period": marker_period,
        "marker_period_share": (marker_period / total_period) if total_period else None,
        "unresolved_frames": unresolved,
        "stack_truncated": trunc,
        "lost_events": lost,
        "exclusive_deepest_resolved_period": deepest,
        "exclusive_leaf_resolved_period": owners,
        "exclusive_root_resolved_period": root_owner,
        "inclusive_marker_ancestor_period": inclusive,
        "samples": samples,
    }


def percentile(values: list[float], q: float) -> float:
    if not values:
        raise ProtocolError("percentile of empty sequence")
    ordered = sorted(float(v) for v in values)
    pos = (len(ordered) - 1) * q
    lo, hi = math.floor(pos), math.ceil(pos)
    if lo == hi:
        return ordered[lo]
    return ordered[lo] + (ordered[hi] - ordered[lo]) * (pos - lo)


def pearson(xs: list[float], ys: list[float]) -> float | None:
    if len(xs) != len(ys) or len(xs) < 2:
        return None
    mx, my = sum(xs) / len(xs), sum(ys) / len(ys)
    denx = sum((x - mx) ** 2 for x in xs)
    deny = sum((y - my) ** 2 for y in ys)
    if denx <= 0 or deny <= 0:
        return None
    return sum((x - mx) * (y - my) for x, y in zip(xs, ys)) / math.sqrt(denx * deny)


def spearman(xs: list[float], ys: list[float]) -> float | None:
    def ranks(values: list[float]) -> list[float]:
        order = sorted(range(len(values)), key=lambda i: values[i])
        out = [0.0] * len(values)
        i = 0
        while i < len(order):
            j = i + 1
            while j < len(order) and values[order[j]] == values[order[i]]:
                j += 1
            rank = (i + j - 1) / 2 + 1
            for k in order[i:j]:
                out[k] = rank
            i = j
        return out
    return pearson(ranks(xs), ranks(ys))


def finite_positive(value: Any, label: str) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(float(value)) or float(value) <= 0:
        raise ProtocolError(f"{label}: expected finite positive number")
    return float(value)


def check_launch_metadata(path: Path, *, packet_id: str, binary: Path) -> dict[str, Any]:
    metadata = load_json(path)
    if metadata.get("schema") != LAUNCH_SCHEMA or metadata.get("packet_id") != packet_id:
        raise ProtocolError("launcher metadata schema/packet identity mismatch")
    check_binding({"path": metadata.get("helper_path"), "sha256": metadata.get("helper_sha256")}, "launcher helper", expected_path="/home/zhuhe/litchi-performance-0970/tools/launch_baseline_v2.py", expected_sha=LAUNCH_SHA)
    if metadata.get("execve") is not True or metadata.get("exec_path") != str(binary.resolve()):
        raise ProtocolError("launcher did not report exact execve target")
    if metadata.get("requested_affinity") != [CPU] or metadata.get("actual_affinity") != [CPU]:
        raise ProtocolError("launcher affinity is not exactly CPU 31")
    if metadata.get("pid") != metadata.get("exec_pid") or not isinstance(metadata.get("pid"), int):
        raise ProtocolError("launcher PID/exec PID binding invalid")
    return metadata


def check_artifact_bindings(packet: dict[str, Any], packet_path: Path | None = None) -> dict[str, Path]:
    artifacts = packet.get("artifacts")
    if not isinstance(artifacts, dict):
        raise ProtocolError("packet artifact binding section is missing")
    required = ("report", "stderr", "perf_data", "launch", "command", "terminal", "no_inline", "inline", "header", "buildid")
    if set(artifacts) != set(required):
        raise ProtocolError("packet raw artifact set drifted")
    result: dict[str, Path] = {}
    for name in required:
        row = artifacts[name]
        path = check_binding(row, f"packet artifact {name}")
        result[name] = path
    if packet_path is not None:
        child_dir = packet_path.parent
        for name, path in result.items():
            if path.parent != child_dir:
                raise ProtocolError(f"packet artifact {name} escaped child directory")
    return result


def check_buildid_export(text: str, *, binary_path: str = BINARY_PATH, build_id: str = BINARY_BUILD_ID) -> None:
    for line in text.splitlines():
        fields = line.split()
        if len(fields) >= 2 and fields[0].lower() == build_id.lower() and fields[-1] == binary_path:
            return
    raise ProtocolError("perf buildid-list lacks exact retained binary path/build ID")


def check_plan_bindings(plan: dict[str, Any], plan_path: Path | None = None) -> None:
    if plan.get("schema") != PLAN_SCHEMA or plan.get("status") != "pre_capture":
        raise ProtocolError("plan schema/status is not pre-capture")
    if plan.get("no_build_or_workload") is not True:
        raise ProtocolError("plan must prohibit build/workload before GO")
    if (plan.get("execution") or {}).get("serial") is not True or (plan.get("execution") or {}).get("retry") is not False:
        raise ProtocolError("plan must be serial and no-retry")
    matrix = plan.get("matrix") or {}
    if matrix.get("children") != CHILDREN or matrix.get("warmups") != WARMUPS or matrix.get("samples") != SAMPLES or matrix.get("cpu") != CPU:
        raise ProtocolError("matrix drift")
    cells = matrix.get("cells")
    if cells != [{"case": "docfloat", "edits": 1, "retention": "default"}]:
        raise ProtocolError("0971 matrix must contain exactly the Float/e1/default cell")
    if matrix.get("serial_order") != list(range(1, CHILDREN + 1)):
        raise ProtocolError("child order drift")
    perf = plan.get("perf") or {}
    if perf.get("event") != "instructions:u" or perf.get("period") != PERIOD or perf.get("callgraph") != "dwarf,16384":
        raise ProtocolError("perf event/period/callgraph drift")
    marker = plan.get("marker") or {}
    if marker.get("symbol") != MARKER or marker.get("calls") != CALLS or marker.get("p50_scope") != "80 measured samples only" or marker.get("marker_scope") != "5 warmups + 80 measured calls combined":
        raise ProtocolError("marker scope drift")
    if marker.get("dso_path") != BINARY_PATH or marker.get("build_id") != BINARY_BUILD_ID:
        raise ProtocolError("marker DSO/build-ID binding drift")
    pilot = plan.get("pilot_gate") or {}
    if pilot != {"child": 1, "min_marker_samples": 100, "lost_events": 0, "marker_period_share_positive": True, "on_failure": "stop_without_retry"}:
        raise ProtocolError("pilot gate drift")
    for section in ("binary", "fixture", "source_manifest", "source_file_manifest", "source_proof", "proposal", "capability", "source_variation", "adr_revalidation", "validator", "launcher", "python", "taskset", "perf", "helper_seal", "review_context", "correction_receipt"):
        row = plan.get(section)
        if not isinstance(row, dict):
            raise ProtocolError(f"missing binding {section}")
        check_binding(row, section)
    binary = plan["binary"]
    if binary.get("path") != BINARY_PATH or binary.get("sha256") != BINARY_SHA or binary.get("bytes") != BINARY_BYTES or binary.get("build_id") != BINARY_BUILD_ID:
        raise ProtocolError("historical binary binding drift")
    fixture = plan["fixture"]
    if fixture.get("path") != FIXTURE_PATH or fixture.get("sha256") != FIXTURE_SHA or fixture.get("bytes") != FIXTURE_BYTES:
        raise ProtocolError("fixture binding drift")
    if plan["perf"].get("sample_flags") != ["--period", "--timestamp", "--sample-cpu", "--buildid-all", "--no-buildid-cache", "--user-callchains"]:
        raise ProtocolError("perf sample flags drift")
    env = plan.get("environment") or {}
    if env != {"LC_ALL": "C", "LANG": "C", "TZ": "UTC"}:
        raise ProtocolError("environment binding drift")
    if plan["perf"].get("path") != "/usr/bin/perf" or plan["perf"].get("sha256") != PERF_SHA:
        raise ProtocolError("perf tool binding drift")
    if plan["perf"].get("buildid_export") != ["buildid-list", "-i", "<child>.perf.data"]:
        raise ProtocolError("perf build-ID export binding drift")
    if plan.get("report_contract") != {"case": "docfloat", "edits": 1, "retention": "default_8mib", "warmups": WARMUPS, "samples_requested": SAMPLES, "sample_indices": list(range(SAMPLES)), "elapsed_ns_identity": "samples[i].elapsed_ns == elapsed_ns[i]"}:
        raise ProtocolError("report contract drift")
    if plan.get("raw_artifacts") != ["report", "stderr", "perf_data", "launch", "command", "terminal", "no_inline", "inline", "header", "buildid"]:
        raise ProtocolError("raw artifact contract drift")
    if plan["python"].get("path") != "/usr/bin/python3.14" or plan["taskset"].get("path") != "/usr/bin/taskset":
        raise ProtocolError("python/taskset tool binding drift")
    if plan["source_proof"].get("sha256") != SOURCE_PROOF_SHA:
        raise ProtocolError("source content proof binding drift")
    proof = load_json(Path(plan["source_proof"]["path"]))
    if proof.get("schema") != "litchi.performance.0971.source_content_proof.v2" or proof.get("status") != "pass" or proof.get("entries_checked") != 924 or proof.get("mismatches") != [] or proof.get("source_file_manifest_sha256") != SOURCE_FILES_SHA:
        raise ProtocolError("source content proof is not a passing 924-entry baseline proof")
    if plan["helper_seal"].get("sha256") != plan.get("helpers_seal_sha256"):
        raise ProtocolError("helper seal binding checksum missing")
    helper_seal = load_json(Path(plan["helper_seal"]["path"]))
    if helper_seal.get("schema") != HELPER_SCHEMA or helper_seal.get("status") != "sealed":
        raise ProtocolError("helper seal is not sealed v2")
    helpers = plan.get("helpers") or {}
    if set(helpers) != {"protocol", "runner", "reducer", "plan_generator"} or set(helper_seal.get("helpers") or {}) != set(helpers):
        raise ProtocolError("helper binding set drifted")
    for name, row in (helper_seal.get("helpers") or {}).items():
        if name not in helpers or helpers[name].get("path") != row.get("path") or helpers[name].get("sha256") != row.get("sha256") or sha256(Path(row["path"])) != row.get("sha256"):
            raise ProtocolError(f"helper seal/helper binding mismatch: {name}")
    if plan["correction_receipt"].get("status") != "pass":
        raise ProtocolError("v2 correction receipt must be marked pass")
    if plan_path is not None and plan.get("plan_path") not in {None, str(plan_path.resolve())}:
        raise ProtocolError("plan self path mismatch")


def check_go(go: dict[str, Any], plan: dict[str, Any], plan_path: Path, go_path: Path) -> None:
    if go.get("schema") != GO_SCHEMA or go.get("authorized") is not True or go.get("status") != "authorized":
        raise ProtocolError("root GO is not authorized")
    binding = go.get("plan") or {}
    if binding.get("path") != str(plan_path.resolve()) or binding.get("sha256") != sha256(plan_path) or binding.get("bytes") != plan_path.stat().st_size:
        raise ProtocolError("GO plan binding mismatch")
    if (go.get("quiescence") or {}).get("status") != "pass":
        raise ProtocolError("GO lacks passing quiescence")
    helper = go.get("helper_seal") or {}
    if helper.get("path") != plan["helper_seal"]["path"] or helper.get("sha256") != plan["helper_seal"]["sha256"] or helper.get("bytes") != plan["helper_seal"]["bytes"]:
        raise ProtocolError("GO helper seal mismatch")
    seal = load_json(Path(helper["path"]))
    for name, row in (seal.get("helpers") or {}).items():
        actual = go.get("helper_bindings", {}).get(name) if isinstance(go.get("helper_bindings"), dict) else None
        if not isinstance(actual, dict) or actual.get("path") != row.get("path") or actual.get("bytes") != row.get("bytes") or actual.get("sha256") != row.get("sha256") or sha256(Path(row["path"])) != row.get("sha256"):
            raise ProtocolError(f"GO helper file binding mismatch: {name}")
    binary = go.get("binary") or {}
    if binary.get("path") != plan["binary"]["path"] or binary.get("bytes") != plan["binary"]["bytes"] or binary.get("sha256") != plan["binary"]["sha256"] or binary.get("build_id") != BINARY_BUILD_ID or sha256(Path(binary["path"])) != BINARY_SHA:
        raise ProtocolError("GO binary binding mismatch")
    fixture = go.get("fixture") or {}
    if fixture.get("path") != plan["fixture"]["path"] or fixture.get("bytes") != plan["fixture"]["bytes"] or fixture.get("sha256") != plan["fixture"]["sha256"] or sha256(Path(fixture["path"])) != FIXTURE_SHA:
        raise ProtocolError("GO fixture binding mismatch")
    tools = go.get("tools") or {}
    for name in ("perf", "python", "taskset"):
        expected = plan[name]
        actual = tools.get(name)
        if not isinstance(actual, dict) or actual.get("path") != expected.get("path") or actual.get("bytes") != expected.get("bytes") or actual.get("sha256") != expected.get("sha256") or sha256(Path(actual["path"])) != expected.get("sha256"):
            raise ProtocolError(f"GO tool binding mismatch: {name}")
    if (go.get("environment") or {}) != {"LC_ALL": "C", "LANG": "C", "TZ": "UTC"}:
        raise ProtocolError("GO environment binding mismatch")
    host = go.get("host") or {}
    allowed = host.get("allowed_affinity")
    if not isinstance(allowed, list) or CPU not in allowed or host.get("selected_cpu") != CPU:
        raise ProtocolError("GO host affinity binding does not admit CPU 31")
    for name in ("capability", "correction_receipt", "protocol_seal", "independent_review"):
        row = go.get(name)
        if not isinstance(row, dict):
            raise ProtocolError(f"GO {name} binding missing")
        check_binding(row, f"GO {name}")
    if go["capability"] != plan["capability"] or go["correction_receipt"]["path"] != plan["correction_receipt"]["path"] or go["correction_receipt"]["sha256"] != plan["correction_receipt"]["sha256"]:
        raise ProtocolError("GO capability/correction binding does not match plan")
    review = load_json(Path(go["independent_review"]["path"]))
    if review.get("status") not in {"pass", "passed"}:
        raise ProtocolError("GO independent review is not passing")
    correction = load_json(Path(go["correction_receipt"]["path"]))
    if correction.get("status") != "pass":
        raise ProtocolError("GO correction receipt is not passing")
    source = go.get("source") or {}
    for name in ("manifest", "file_manifest", "content_proof", "harness_source", "harness_cargo"):
        row = source.get(name)
        if not isinstance(row, dict):
            raise ProtocolError(f"GO source binding missing: {name}")
        check_binding(row, f"GO source.{name}")
        if name in plan.get("source", {}):
            expected_source = plan["source"][name]
            if row.get("path") != expected_source.get("path") or row.get("sha256") != expected_source.get("sha256") or row.get("bytes") != expected_source.get("bytes"):
                raise ProtocolError(f"GO source binding does not match plan: {name}")
    if source["manifest"]["sha256"] != SOURCE_MANIFEST_SHA or source["file_manifest"]["sha256"] != SOURCE_FILES_SHA or source["content_proof"]["sha256"] != SOURCE_PROOF_SHA:
        raise ProtocolError("GO source custody hash mismatch")
    outputs = go.get("outputs") or {}
    for name in ("capture_root", "analysis_root"):
        raw = outputs.get(name)
        if not isinstance(raw, str) or not raw:
            raise ProtocolError(f"GO output root missing: {name}")
        if raw != plan["outputs"][name]:
            raise ProtocolError(f"GO output root does not match plan: {name}")
        path = Path(raw)
        if path.exists() and any(path.iterdir()):
            raise ProtocolError(f"GO output root is not empty: {path}")
    if not go_path.is_file():
        raise ProtocolError("GO path is not a file")
