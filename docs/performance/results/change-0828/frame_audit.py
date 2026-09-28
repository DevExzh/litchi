"""Independent perf-frame census for the 0828 edit-phase packet.

This reader intentionally repeats the owner and nested-phase partitioning
logic instead of importing ``analysis.py``. It consumes only retained gzip
members and static plan/build descriptors; it never invokes perf, Cargo, or a
decoder. Phase rows are a mutually exclusive sample partition and are not
timing or causal fractions.
"""

from __future__ import annotations

from collections import Counter
import gzip
import hashlib
import json
from pathlib import Path
import re
import sys


P = Path(__file__).resolve().parent
OWNER = "pptx_edit_profile_0828::edit_region_0828"
PHASE_OWNERS = {
    "capture": "pptx_edit_profile_0828::phase_opened_presentation_transaction_0828",
    "set_text": "pptx_edit_profile_0828::phase_set_shape_text_0828",
    "publish": "pptx_edit_profile_0828::phase_commit_apply_opened_presentation_commit_0828",
}
PHASE_ORDER = tuple(PHASE_OWNERS)
HEADER = re.compile(
    r"^\s*\S+\s+\d+\s+(?:\[\d+\]\s+)?\d+(?:\.\d+)?:\s+"
    r"(\d+)\s+cycles:u:\s*$"
)
FRAME = re.compile(r"^\s*[0-9a-fA-F]+\s+(.+?)\s+\(([^)]*)\)\s*$")


def read(path: Path):
    return json.loads(path.read_text(encoding="utf-8"))


def descriptor(value, label: str) -> Path:
    assert isinstance(value, dict) and isinstance(value.get("path"), str), label
    path = Path(value["path"])
    data = path.read_bytes()
    assert len(data) == value["bytes"], label
    assert hashlib.sha256(data).hexdigest() == value["sha256"], label
    return path


def parse_frames(data: bytes):
    """Parse headers one at a time so a header-only (empty) sample survives."""
    text = data.decode("utf-8", errors="replace")
    lines = text.splitlines()
    samples = []
    current = None
    status_lines = []
    malformed = []

    def finish():
        nonlocal current
        if current is not None:
            samples.append(current)
            current = None

    for line in lines:
        match = HEADER.fullmatch(line)
        if match:
            finish()
            current = {"index": len(samples), "period": int(match[1]), "frames": [],
                       "raw_lines": [line]}
            continue
        if current is None:
            if line.strip():
                status_lines.append(line)
            continue
        if not line.strip():
            finish()
            continue
        frame = FRAME.fullmatch(line)
        if frame is None:
            malformed.append(line)
            current["raw_lines"].append(line)
            continue
        symbol = re.sub(r"\+0x[0-9a-fA-F]+$", "", frame[1].strip())
        current["raw_lines"].append(line)
        current["frames"].append({"symbol": symbol, "dso": frame[2].strip()})
    finish()
    assert samples, "decoded perf stream has no sample headers"
    assert all(sample["period"] > 0 for sample in samples)
    return {
        "samples": samples,
        "whole_samples": len(samples),
        "whole_period": sum(sample["period"] for sample in samples),
        "status_lines": status_lines,
        "malformed_lines": malformed,
        "lost_lines": [line for line in lines if "PERF_RECORD_LOST" in line
                        or ("lost" in line.lower() and "sample" in line.lower())],
        "truncation_lines": [line for line in lines
                             if re.search(r"truncat|stack depth|callchain", line, re.I)],
        "raw_lines": len(lines),
        "raw_bytes": len(data),
    }


def exact_dso(binary: dict) -> str:
    raw = binary.get("artifact") if isinstance(binary.get("artifact"), dict) else binary
    assert isinstance(raw, dict) and isinstance(raw.get("path"), str)
    return raw["path"]


def phase_partition(frames: list[dict], binary_path: str):
    hits = []
    for index, frame in enumerate(frames):
        for phase, symbol in PHASE_OWNERS.items():
            if frame["symbol"] == symbol:
                hits.append({"phase": phase, "index": index, "dso": frame["dso"]})
    exact = [hit for hit in hits if hit["dso"] == binary_path]
    other = [hit for hit in hits if hit["dso"] != binary_path]
    if len(hits) == 1 and len(exact) == 1:
        return "phase", exact[0]["phase"], hits, exact, other
    if len(hits) > 1 or len(exact) > 1:
        return "ambiguous", None, hits, exact, other
    return "unclassified", None, hits, exact, other


def census(parsed: dict, binary_path: str, repeat: int):
    phase_counts = Counter()
    phase_period = Counter()
    phase_leaves = {phase: Counter() for phase in PHASE_ORDER}
    phase_leaf_period = {phase: Counter() for phase in PHASE_ORDER}
    unclassified = Counter()
    unclassified_period = Counter()
    ambiguous = Counter()
    ambiguous_period = Counter()
    leaves = Counter()
    leaf_period = Counter()
    owner_samples = []
    owner_other_dso = 0
    owner_repeated = 0
    unknown_interior = 0
    phase_other_dso = 0
    phase_marker_hits = 0
    for sample in parsed["samples"]:
        owner_hits = [i for i, frame in enumerate(sample["frames"])
                      if frame["symbol"] == OWNER]
        qualified = [i for i in owner_hits
                     if sample["frames"][i]["dso"] == binary_path]
        if len(owner_hits) > 1:
            owner_repeated += 1
        if owner_hits and not qualified:
            owner_other_dso += 1
        if len(qualified) != 1:
            continue
        owner_index = qualified[0]
        descendants = sample["frames"][:owner_index]
        leaf = (descendants[0]["symbol"], descendants[0]["dso"]) \
            if descendants else (OWNER, binary_path)
        leaves[leaf] += 1
        leaf_period[leaf] += sample["period"]
        unknown = any("[unknown]" in frame["symbol"].lower()
                      or frame["symbol"].lower() in {"unknown", "??", "<unknown>"}
                      for frame in descendants)
        unknown_interior += int(unknown)
        status, phase, hits, exact, other = phase_partition(descendants, binary_path)
        phase_marker_hits += len(hits)
        phase_other_dso += len(other)
        if status == "phase":
            phase_counts[phase] += 1
            phase_period[phase] += sample["period"]
            phase_leaves[phase][leaf] += 1
            phase_leaf_period[phase][leaf] += sample["period"]
        elif status == "ambiguous":
            ambiguous[leaf] += 1
            ambiguous_period[leaf] += sample["period"]
        else:
            unclassified[leaf] += 1
            unclassified_period[leaf] += sample["period"]
        owner_samples.append({
            "sample_index": sample["index"], "period": sample["period"],
            "owner_frame_index": owner_index, "leaf": leaf[0], "leaf_dso": leaf[1],
            "phase_status": status, "phase": phase,
            "phase_hits": hits, "phase_exact_hits": exact,
            "phase_other_dso_hits": other, "unknown_interior": unknown,
        })

    assert owner_samples, f"repeat {repeat}: exact owner has zero samples"
    counts = {phase: phase_counts[phase] for phase in PHASE_ORDER}
    counts["unclassified"] = sum(unclassified.values())
    counts["ambiguous"] = sum(ambiguous.values())
    periods = {phase: phase_period[phase] for phase in PHASE_ORDER}
    periods["unclassified"] = sum(unclassified_period.values())
    periods["ambiguous"] = sum(ambiguous_period.values())
    assert sum(counts.values()) == len(owner_samples)
    assert sum(periods.values()) == sum(row["period"] for row in owner_samples)

    def render(values, periods_):
        return [{"symbol": symbol, "dso": dso, "samples": count,
                 "period": periods_[(symbol, dso)]}
                for (symbol, dso), count in sorted(values.items(),
                                                   key=lambda item: (-item[1], item[0]))]

    return {
        "repeat": repeat,
        "whole_process_samples": parsed["whole_samples"],
        "whole_process_period": parsed["whole_period"],
        "empty_stack_samples": sum(not sample["frames"] for sample in parsed["samples"]),
        "owner_samples": len(owner_samples),
        "outside_owner_samples": parsed["whole_samples"] - len(owner_samples),
        "owner_period": sum(row["period"] for row in owner_samples),
        "outside_owner_period": parsed["whole_period"] - sum(row["period"] for row in owner_samples),
        "owner_symbol_other_dso_samples": owner_other_dso,
        "owner_repeated_samples": owner_repeated,
        "unknown_interior_samples": unknown_interior,
        "phase_partition": {
            "phase_order": list(PHASE_ORDER), "phase_owners": dict(PHASE_OWNERS),
            "sample_counts": counts, "periods": periods,
            "owner_samples": len(owner_samples),
            "classified_samples": sum(phase_counts.values()),
            "unclassified_samples": counts["unclassified"],
            "ambiguous_samples": counts["ambiguous"],
            "phase_marker_hits": phase_marker_hits,
            "phase_marker_other_dso_hits": phase_other_dso,
            "leaf_census": {phase: render(phase_leaves[phase], phase_leaf_period[phase])
                            for phase in PHASE_ORDER},
            "unclassified_leaf_census": render(unclassified, unclassified_period),
            "ambiguous_leaf_census": render(ambiguous, ambiguous_period),
            "sample_rows": owner_samples,
            "additive_timing_claim": False,
            "causal_fraction_claim": False,
        },
        "leaves": render(leaves, leaf_period),
        "decode_diagnostics": {
            "status_line_count": len(parsed["status_lines"]),
            "status_lines": parsed["status_lines"],
            "malformed_frame_count": len(parsed["malformed_lines"]),
            "malformed_frame_lines": parsed["malformed_lines"],
            "lost_event_lines": len(parsed["lost_lines"]),
            "lost_event_text": parsed["lost_lines"],
            "explicit_truncation_count": len(parsed["truncation_lines"]),
            "explicit_truncation_lines": parsed["truncation_lines"],
            "raw_text_lines": parsed["raw_lines"],
            "raw_text_bytes": parsed["raw_bytes"],
        },
    }


def main() -> None:
    assert sys.argv[1:] in (["--write"], ["--check"])
    plan = read(P / "plan.json")
    assert plan.get("schema") == "litchi.performance.0828.pptx-edit-profile.v1"
    assert plan.get("probe", {}).get("owner") == OWNER
    assert plan["probe"].get("phase_owners") == PHASE_OWNERS
    complete = read(P / "perf/decode-complete.json")
    result = {"schema": "litchi.performance.0828.frame-audit.v1",
              "status": complete.get("status"), "repeats": [],
              "phase_owners": dict(PHASE_OWNERS),
              "partition_contract": {
                  "mutually_exclusive": True, "exhaustive": True,
                  "additive_timing_claim": False, "causal_fraction_claim": False,
              }}
    if complete.get("status") == "available":
        build = read(P / "build.json")
        binary = build["binaries"][plan["perf"]["binary"]]
        members = read(P / "perf/compression.json")
        seen = set()
        for member in members:
            if member.get("kind") != "frames":
                continue
            repeat = member["repeat"]
            assert repeat not in seen
            seen.add(repeat)
            stored = descriptor(member["compressed"], f"compressed frames {repeat}")
            raw = gzip.decompress(stored.read_bytes())
            assert len(raw) == member["original"]["bytes"]
            assert hashlib.sha256(raw).hexdigest() == member["original"]["sha256"]
            result["repeats"].append(census(parse_frames(raw), exact_dso(binary), repeat))
        assert sorted(seen) == list(range(plan["perf"]["repeats"]))
        result["repeats"].sort(key=lambda row: row["repeat"])
    encoded = json.dumps(result, sort_keys=True, indent=2) + "\n"
    out = P / "frame-audit.json"
    if sys.argv[1] == "--write":
        assert not out.exists(), "refusing to overwrite frame audit"
        out.write_text(encoded, encoding="utf-8")
    else:
        assert out.is_file() and out.read_text(encoding="utf-8") == encoded
    print("0828 independent frame audit PASS")


if __name__ == "__main__":
    main()
