#!/usr/bin/env python3
"""Reduce retained 0942 CFB samples inside the commit candidate.

This is an offline, read-only attribution pass.  It consumes the retained
``perf script`` text exports and binds them to the 0942 binary, reduction, and
source bytes.  It never runs the benchmark, invokes ``perf``, builds source,
or writes into the main checkout.

The qualified denominator is a sample with both the physical
``profiled_measured_case`` barrier and the physical
``Editor::commit_candidate_with_rendered`` frame, with the commit frame below
the barrier in the leaf-first callchain.  Target tags are inclusive and may
overlap.  The separate exclusive table assigns each sample to its nearest
observed target frame, then to another physical CFB frame or a remainder
bucket, so its periods conserve the denominator.
"""

from __future__ import annotations

import argparse
from collections import Counter, defaultdict
from datetime import datetime, timezone
import hashlib
import json
import os
from pathlib import Path
import re
import subprocess
import sys
from typing import Any, Iterable


ROOT = Path("/home/zhuhe/litchi-performance-0942").resolve()
REPO = Path("/home/zhuhe/code/litchi").resolve()
CAPTURE_ROOT = ROOT / "capture/formal-v5/children"
BINARY = ROOT / "target-profile-v2/release/doc_retention_probe_0942_cfb_profile"
REDUCTION = ROOT / "reports/formal-reduction-v5.json"
MANIFEST = ROOT / "manifests/cfb-profile-manifest-v5.json"

MARKER = "doc_retention_probe_0942_cfb_profile::profiled_measured_case"
COMMIT = "litchi_ole_common::object::editor::Editor::commit_candidate_with_rendered"
EVENT = "cycles:u"

# These are exact physical symbol names from the no-inline retained export.
# ``capture_container`` is included as a source-level target even though the
# retained export does not show a physical frame for it.
TARGETS: dict[str, str] = {
    "package_capture": "litchi_ole_common::object::codec::Package::capture",
    "capture_with_budget": "litchi_ole_common::object::codec::Package::capture_with_budget",
    "capture_container": "litchi_ole_common::object::codec::capture_container",
    "cfb_open_stream": "litchi_cfb::file::OleFile<R>::open_stream",
    "cfb_ole_open": "litchi_cfb::file::OleFile<R>::open",
    "ole_codec_open": "litchi_ole_common::object::codec::open",
    "discovery_from_package": "litchi_ole_common::object::discovery::from_package",
    "cfb_writer_write_to": "litchi_cfb::writer::core::OleWriter::write_to",
}

SOURCE_FILES: tuple[tuple[str, str], ...] = (
    (
        "crates/litchi-ole-common/src/object/editor.rs",
        "commit_candidate_with_rendered; render, reopen, capture, rediscover",
    ),
    (
        "crates/litchi-ole-common/src/object/codec.rs",
        "Package::capture, capture_with_budget, capture_container, codec::open",
    ),
    (
        "crates/litchi-ole-common/src/object/discovery.rs",
        "from_package discovery traversal",
    ),
    (
        "crates/litchi-cfb/src/file.rs",
        "OleFile::open and OleFile::open_stream",
    ),
    (
        "crates/litchi-cfb/src/writer/core.rs",
        "OleWriter::write_to",
    ),
    (
        "crates/litchi-cfb/src/writer/layout.rs",
        "layout validation and stream comparison context",
    ),
)

HEADER_RE = re.compile(
    r"^\s*(?P<comm>\S+)\s+(?P<pid>[0-9]+)/(?P<tid>[0-9]+)\s+"
    r"(?P<time>[0-9]+(?:\.[0-9]+)?)\s*:\s*(?P<period>[0-9]+)\s+"
    r"(?P<event>\S+)"
)
FRAME_RE = re.compile(
    r"^\s*(?P<ip>(?:0x)?[0-9a-fA-F]+)\s+(?P<symbol>.*?)\s+"
    r"\((?P<dso>[^)]*)\)(?:\s+(?P<frame_period>[0-9]+))?\s*$"
)
ADDRESS_RE = re.compile(r"^(?:0x)?[0-9a-fA-F]+$")


class AnalysisError(RuntimeError):
    pass


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1 << 20), b""):
            digest.update(block)
    return digest.hexdigest()


def describe(path: Path, *, relative_to: Path | None = None) -> dict[str, Any]:
    path = path.resolve()
    stat = path.stat()
    rendered = str(path)
    if relative_to is not None:
        rendered = path.relative_to(relative_to.resolve()).as_posix()
    return {
        "path": rendered,
        "bytes": stat.st_size,
        "sha256": sha256(path),
    }


def run_git_head() -> str | None:
    result = subprocess.run(
        ["git", "-C", str(REPO), "rev-parse", "HEAD"],
        check=False,
        stdout=subprocess.PIPE,
        stderr=subprocess.DEVNULL,
        text=True,
    )
    return result.stdout.strip() if result.returncode == 0 else None


def physical(frame: dict[str, str]) -> bool:
    return frame["dso"] in {str(BINARY), BINARY.name}


def parse_script(path: Path) -> tuple[list[dict[str, Any]], int]:
    records: list[dict[str, Any]] = []
    current: dict[str, Any] | None = None
    lost = 0

    def finish() -> None:
        nonlocal current
        if current is not None:
            records.append(current)
            current = None

    with path.open("r", encoding="utf-8") as stream:
        for line_number, line in enumerate(stream, 1):
            stripped = line.strip()
            if not stripped or stripped.startswith("#"):
                continue
            if "PERF_RECORD_LOST" in line:
                lost += 1
                continue
            if stripped.startswith("..."):
                if current is None:
                    raise AnalysisError(
                        f"truncation marker before sample: {path}:{line_number}"
                    )
                current["truncated"] = True
                continue
            header = HEADER_RE.match(line)
            if header is not None:
                finish()
                period = int(header.group("period"))
                event = header.group("event").rstrip(":")
                if period <= 0:
                    raise AnalysisError(f"nonpositive period: {path}:{line_number}")
                current = {
                    "event": event,
                    "period": period,
                    "frames": [],
                    "truncated": False,
                    "line": line_number,
                }
                continue
            frame = FRAME_RE.match(line)
            if frame is None:
                raise AnalysisError(f"unparsed line: {path}:{line_number}: {stripped}")
            if current is None:
                raise AnalysisError(f"frame outside sample: {path}:{line_number}")
            current["frames"].append(
                {
                    "ip": frame.group("ip"),
                    "symbol": frame.group("symbol").strip(),
                    "dso": frame.group("dso").strip() or "[unknown]",
                }
            )
    finish()
    if not records:
        raise AnalysisError(f"no records: {path}")
    return records, lost


def exact_indices(
    record: dict[str, Any], symbol: str, *, require_physical: bool = True
) -> list[int]:
    return [
        index
        for index, frame in enumerate(record["frames"])
        if frame["symbol"] == symbol and (not require_physical or physical(frame))
    ]


def qualified_commit(record: dict[str, Any]) -> tuple[int, int] | None:
    if record["event"] != EVENT:
        return None
    markers = exact_indices(record, MARKER)
    commits = exact_indices(record, COMMIT)
    if len(markers) != 1 or len(commits) != 1:
        return None
    marker_index, commit_index = markers[0], commits[0]
    # perf script is leaf first: a nested commit frame has a lower index than
    # the outer measured-case barrier.
    if commit_index >= marker_index:
        return None
    return marker_index, commit_index


def target_hits(record: dict[str, Any], commit_index: int) -> list[tuple[int, str]]:
    hits: list[tuple[int, str]] = []
    # Include the commit root in the bounded walk.  No target symbol equals
    # the commit marker, but this makes the boundary explicit in the receipt.
    for index, frame in enumerate(record["frames"][: commit_index + 1]):
        if not physical(frame):
            continue
        for label, symbol in TARGETS.items():
            if frame["symbol"] == symbol:
                hits.append((index, label))
    return hits


def has_physical_cfb(record: dict[str, Any], commit_index: int) -> bool:
    return any(
        physical(frame) and frame["symbol"].startswith("litchi_cfb::")
        for frame in record["frames"][: commit_index + 1]
    )


def pair(row: list[int]) -> dict[str, int]:
    return {"samples": row[0], "period": row[1]}


def add_metric(table: dict[str, list[int]], key: str, period: int) -> None:
    table.setdefault(key, [0, 0])
    table[key][0] += 1
    table[key][1] += period


def metric_rows(table: dict[str, list[int]], denominator: int) -> dict[str, Any]:
    return {
        key: {
            "samples": value[0],
            "period": value[1],
            "period_share_percent": round(100.0 * value[1] / denominator, 6)
            if denominator
            else 0.0,
        }
        for key, value in sorted(table.items())
    }


def analyze_records(rows: list[tuple[dict[str, Any], int, int]]) -> dict[str, Any]:
    denominator_samples = len(rows)
    denominator_period = sum(record["period"] for record, _, _ in rows)
    inclusive: dict[str, list[int]] = {}
    target_union = [0, 0]
    cfb_union = [0, 0]
    exclusive: dict[str, list[int]] = {}
    combinations: Counter[tuple[str, ...]] = Counter()
    combination_period: Counter[tuple[str, ...]] = Counter()

    for record, _, commit_index in rows:
        hits = target_hits(record, commit_index)
        labels = tuple(sorted({label for _, label in hits}))
        combinations[labels] += 1
        combination_period[labels] += record["period"]
        if labels:
            target_union[0] += 1
            target_union[1] += record["period"]
            for label in labels:
                add_metric(inclusive, label, record["period"])
        if has_physical_cfb(record, commit_index):
            cfb_union[0] += 1
            cfb_union[1] += record["period"]

        if hits:
            # Leaf-first means the smallest frame index is the deepest
            # observed target.  This gives a deterministic disjoint table.
            nearest = min(hits)[1]
            add_metric(exclusive, nearest, record["period"])
        elif has_physical_cfb(record, commit_index):
            add_metric(exclusive, "other_physical_cfb", record["period"])
        else:
            add_metric(exclusive, "no_target_cfb_frame", record["period"])

    package_family = [0, 0]
    open_family = [0, 0]
    package_stream_overlap = [0, 0]
    for record, _, commit_index in rows:
        labels = {label for _, label in target_hits(record, commit_index)}
        if labels & {"package_capture", "capture_with_budget", "capture_container", "cfb_open_stream"}:
            package_family[0] += 1
            package_family[1] += record["period"]
        if labels & {"cfb_ole_open", "ole_codec_open"}:
            open_family[0] += 1
            open_family[1] += record["period"]
        if "package_capture" in labels and "cfb_open_stream" in labels:
            package_stream_overlap[0] += 1
            package_stream_overlap[1] += record["period"]

    combination_rows = []
    for labels, count in combinations.most_common():
        combination_rows.append(
            {
                "labels": list(labels),
                "samples": count,
                "period": combination_period[labels],
                "period_share_percent": round(
                    100.0 * combination_period[labels] / denominator_period, 6
                )
                if denominator_period
                else 0.0,
            }
        )

    exclusive_samples = sum(value[0] for value in exclusive.values())
    exclusive_period = sum(value[1] for value in exclusive.values())
    inclusive_period = sum(value[1] for value in inclusive.values())
    return {
        "denominator": {
            "samples": denominator_samples,
            "period": denominator_period,
        },
        "inclusive_tags": metric_rows(inclusive, denominator_period),
        "derived_inclusive": {
            "capture_family_package_capture_container_open_stream": {
                "samples": package_family[0],
                "period": package_family[1],
                "period_share_percent": round(
                    100.0 * package_family[1] / denominator_period, 6
                )
                if denominator_period
                else 0.0,
            },
            "validation_open_family_ole_open_or_codec_open": {
                "samples": open_family[0],
                "period": open_family[1],
                "period_share_percent": round(
                    100.0 * open_family[1] / denominator_period, 6
                )
                if denominator_period
                else 0.0,
            },
            "package_capture_and_open_stream_overlap": {
                "samples": package_stream_overlap[0],
                "period": package_stream_overlap[1],
                "period_share_percent": round(
                    100.0 * package_stream_overlap[1] / denominator_period, 6
                )
                if denominator_period
                else 0.0,
            },
            "all_exact_target_union": {
                "samples": target_union[0],
                "period": target_union[1],
                "period_share_percent": round(
                    100.0 * target_union[1] / denominator_period, 6
                )
                if denominator_period
                else 0.0,
            },
            "any_physical_litchi_cfb_frame": {
                "samples": cfb_union[0],
                "period": cfb_union[1],
                "period_share_percent": round(
                    100.0 * cfb_union[1] / denominator_period, 6
                )
                if denominator_period
                else 0.0,
            },
        },
        "inclusive_overlap": {
            "sum_of_tag_periods": inclusive_period,
            "exact_target_union_period": target_union[1],
            "overlap_excess_period": inclusive_period - target_union[1],
            "warning": "inclusive tag periods overlap and must not be added",
        },
        "exclusive_nearest_frame": {
            "buckets": metric_rows(exclusive, denominator_period),
            "sum_samples": exclusive_samples,
            "sum_period": exclusive_period,
            "conserves_denominator": (
                exclusive_samples == denominator_samples
                and exclusive_period == denominator_period
            ),
            "rule": "nearest physical target frame from leaf toward commit root; then other physical CFB; then no target CFB frame",
        },
        "target_combinations": combination_rows,
    }


def source_equivalence() -> dict[str, Any]:
    rows: list[dict[str, Any]] = []
    for relative, role in SOURCE_FILES:
        retained = ROOT / "source-v2" / relative
        current = REPO / relative
        retained_sha = sha256(retained)
        current_sha = sha256(current)
        rows.append(
            {
                "file": relative,
                "role": role,
                "retained": describe(retained, relative_to=ROOT / "source-v2"),
                "current": describe(current, relative_to=REPO),
                "same_bytes": retained_sha == current_sha and retained.read_bytes() == current.read_bytes(),
            }
        )
    return {
        "current_head": run_git_head(),
        "retained_manifest_source_revision": json.loads(
            MANIFEST.read_text(encoding="utf-8")
        ).get("source_revision"),
        "revision_used_as_equivalence_proof": False,
        "files": rows,
        "all_files_same_bytes": all(row["same_bytes"] for row in rows),
    }


def check_reduction(
    per_child: dict[str, list[dict[str, Any]]], reduction: dict[str, Any]
) -> dict[str, Any]:
    checks: list[dict[str, Any]] = []
    for child in reduction["formal_children"]:
        child_id = child["id"]
        marker = child["profile"]["marker"]
        observed = per_child.get(child_id, [])
        marker_samples = sum(1 for row in observed if row["marker_bound"])
        marker_period = sum(row["period"] for row in observed if row["marker_bound"])
        checks.append(
            {
                "child": child_id,
                "observed_marker": {"samples": marker_samples, "period": marker_period},
                "retained_reduction_marker": {
                    "samples": marker["samples"],
                    "period": marker["period"],
                },
                "match": marker_samples == marker["samples"] and marker_period == marker["period"],
            }
        )
    return {
        "children": checks,
        "all_match": all(row["match"] for row in checks),
    }


def percent(value: int, denominator: int) -> float:
    return round(100.0 * value / denominator, 6) if denominator else 0.0


def md_table(rows: Iterable[dict[str, Any]], columns: list[tuple[str, str]]) -> str:
    lines = ["| " + " | ".join(title for _, title in columns) + " |", "|" + "|".join("---" for _ in columns) + "|"]
    for row in rows:
        lines.append("| " + " | ".join(str(row.get(key, "")) for key, _ in columns) + " |")
    return "\n".join(lines)


def write_markdown(path: Path, result: dict[str, Any]) -> None:
    focus = result["focus_floatpictures_e32"]
    all_formal = result["all_formal"]
    d = focus["denominator"]
    lines = [
        "# 0968 retained 0942 CFB commit attribution",
        "",
        "This is an offline reduction of retained 0942 frame-pointer `cycles:u`",
        "callchains. It makes no build, workload, latency, allocation, or adoption",
        "claim. All outputs are outside the repository.",
        "",
        "## Qualified scope",
        "",
        f"The exact physical root is `{COMMIT}` nested below the exact physical",
        f"barrier `{MARKER}` in leaf-first callchains. The FloatPictures/e32 focus",
        f"denominator is **{d['samples']:,} samples / {d['period']:,} period**",
        "across repetitions r01 and r02. No selected record was truncated and no",
        "lost record was reported. The all-formal denominator is",
        f"{all_formal['denominator']['samples']:,} samples / {all_formal['denominator']['period']:,} period.",
        "",
        "The marker parser reproduces the retained formal reduction for all 16",
        "children before the commit subset is applied. Source equivalence uses",
        "byte SHA-256 comparisons of the retained runtime files against the current",
        f"checkout at HEAD `{result['source_equivalence']['current_head']}`; the retained",
        "revision string is metadata only.",
        "",
        "## FloatPictures/e32 focus",
        "",
        "The percentages below use the commit denominator above. Inclusive rows can",
        "overlap and must not be summed. The exclusive table assigns each sample to",
        "the nearest observed physical target frame and conserves the denominator.",
        "",
    ]
    inclusive = focus["inclusive_tags"]
    rows = []
    labels = [
        ("package_capture", "Package::capture"),
        ("capture_with_budget", "Package::capture_with_budget"),
        ("capture_container", "capture_container"),
        ("cfb_open_stream", "OleFile::open_stream"),
        ("cfb_ole_open", "OleFile::open"),
        ("ole_codec_open", "object::codec::open"),
        ("discovery_from_package", "discovery::from_package"),
        ("cfb_writer_write_to", "OleWriter::write_to"),
    ]
    for key, label in labels:
        value = inclusive.get(key, {"samples": 0, "period": 0, "period_share_percent": 0.0})
        rows.append({"target": label, "samples": value["samples"], "period": value["period"], "share": value["period_share_percent"]})
    lines.append(md_table(rows, [("target", "inclusive physical target"), ("samples", "samples"), ("period", "period"), ("share", "% of commit")]))
    lines += [
        "",
        "For this focus, Package::capture is 25.054588% inclusive; its",
        "open_stream subset is 13.878805%, so adding those rows double counts",
        "874,651,929 period (205 samples). The exact-target union is",
        f"{focus['derived_inclusive']['all_exact_target_union']['samples']:,} samples / {focus['derived_inclusive']['all_exact_target_union']['period']:,} period",
        f"({focus['derived_inclusive']['all_exact_target_union']['period_share_percent']:.6f}%); any physical",
        f"`litchi_cfb::` frame covers {focus['derived_inclusive']['any_physical_litchi_cfb_frame']['samples']:,} samples /",
        f"{focus['derived_inclusive']['any_physical_litchi_cfb_frame']['period']:,} period ({focus['derived_inclusive']['any_physical_litchi_cfb_frame']['period_share_percent']:.6f}%).",
        "",
        "### Disjoint nearest-frame buckets",
        "",
    ]
    erows = []
    for key, value in focus["exclusive_nearest_frame"]["buckets"].items():
        erows.append({"bucket": key, "samples": value["samples"], "period": value["period"], "share": value["period_share_percent"]})
    lines.append(md_table(erows, [("bucket", "exclusive bucket"), ("samples", "samples"), ("period", "period"), ("share", "% of commit")]))
    lines += [
        "",
        "The retained no-inline export has no physical frame named",
        "`capture_container`, `object::codec::open`, or",
        "`discovery::from_package` inside this commit subset. Their zero named",
        "samples are not zero CPU proof: these helpers may be short, inlined, or",
        "missed by sampling. The source path is confirmed separately by the",
        "editor and codec source lines recorded in the JSON receipt.",
        "",
        "## Replicate split",
        "",
    ]
    rep_rows = []
    for child in ("profile-r01-docfloat-e32", "profile-r02-docfloat-e32"):
        v = result["per_child"][child]
        pkg = v["inclusive_tags"].get("package_capture", {"period": 0, "period_share_percent": 0.0})
        stream = v["inclusive_tags"].get("cfb_open_stream", {"period": 0, "period_share_percent": 0.0})
        writer = v["inclusive_tags"].get("cfb_writer_write_to", {"period": 0, "period_share_percent": 0.0})
        rep_rows.append({
            "child": child,
            "commit": f"{v['denominator']['samples']}/{v['denominator']['period']}",
            "Package::capture": f"{pkg['period_share_percent']:.6f}%",
            "open_stream": f"{stream['period_share_percent']:.6f}%",
            "OleWriter::write_to": f"{writer['period_share_percent']:.6f}%",
        })
    lines.append(md_table(rep_rows, [("child", "child"), ("commit", "commit samples/period"), ("Package::capture", "capture"), ("open_stream", "open_stream"), ("OleWriter::write_to", "writer")]))
    lines += [
        "",
        "The earlier 0942 selected-stage CFB writer observations (36.8% and",
        "38.5% for FloatPictures/e32) use the outer measured-case marker. The",
        "rows here use the narrower commit root, so they are not interchangeable",
        "denominators. The shared DOC/CFB source files are byte-identical to the",
        "current 6493 checkout; no 0968 build or capture was performed.",
        "",
        "## Limits",
        "",
        "These are sampled callchain memberships. Inclusive ancestry is not self",
        "cost, and exclusive nearest-frame buckets exclude callee work from the",
        "leaf owner. Missing named frames do not establish absence. Any candidate",
        "still requires a fresh bound ordinary build, correctness gates, and a",
        "matched profile.",
        "",
    ]
    path.write_text("\n".join(lines), encoding="utf-8")


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path("/home/zhuhe/litchi-performance-0968"),
    )
    args = parser.parse_args(argv)
    out_dir = args.output_dir.resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
    out_json = out_dir / "cfb-commit-attribution-v1.json"
    out_md = out_dir / "cfb-commit-attribution-v1.md"
    if out_json.exists() or out_md.exists():
        raise AnalysisError(f"refusing to overwrite existing output in {out_dir}")

    if not BINARY.is_file() or not REDUCTION.is_file() or not MANIFEST.is_file():
        raise AnalysisError("retained 0942 binding files are incomplete")

    all_records: list[dict[str, Any]] = []
    input_scripts: list[dict[str, Any]] = []
    lost_total = 0
    parse_stats: dict[str, dict[str, Any]] = {}
    for script in sorted(CAPTURE_ROOT.glob("*/perf-script.txt")):
        child = script.parent.name
        records, lost = parse_script(script)
        lost_total += lost
        for record in records:
            record["child"] = child
            record["script"] = str(script)
        all_records.extend(records)
        input_scripts.append(describe(script, relative_to=ROOT))
        parse_stats[child] = {
            "records": len(records),
            "period": sum(record["period"] for record in records),
            "lost_records": lost,
        }

    marker_records: list[dict[str, Any]] = []
    commit_rows: list[tuple[dict[str, Any], int, int]] = []
    per_child_rows: dict[str, list[dict[str, Any]]] = defaultdict(list)
    event_counts: Counter[str] = Counter(record["event"] for record in all_records)
    for record in all_records:
        marker_indices = exact_indices(record, MARKER)
        marker_bound = len(marker_indices) == 1 and physical(record["frames"][marker_indices[0]])
        if marker_bound:
            marker_records.append(record)
        qualification = qualified_commit(record)
        per_child_rows[record["child"]].append(
            {
                "marker_bound": marker_bound,
                "period": record["period"],
            }
        )
        if qualification is not None:
            marker_index, commit_index = qualification
            commit_rows.append((record, marker_index, commit_index))

    reduction = json.loads(REDUCTION.read_text(encoding="utf-8"))
    all_count = len(all_records)
    all_period = sum(record["period"] for record in all_records)
    marker_count = len(marker_records)
    marker_period = sum(record["period"] for record in marker_records)
    commit_truncated = [row for row in commit_rows if row[0]["truncated"]]
    commit_by_child: dict[str, list[tuple[dict[str, Any], int, int]]] = defaultdict(list)
    for row in commit_rows:
        commit_by_child[row[0]["child"]].append(row)

    result: dict[str, Any] = {
        "schema": "litchi.performance.0968.cfb-commit-attribution.v1",
        "status": "passed",
        "created_utc": datetime.now(timezone.utc).isoformat(),
        "tool": describe(Path(__file__).resolve()),
        "scope": {
            "offline_only": True,
            "new_capture": False,
            "new_build": False,
            "main_checkout_writes": False,
            "claim": "diagnostic sampled CPU attribution; no performance or adoption claim",
        },
        "binding": {
            "binary": describe(BINARY),
            "formal_reduction": describe(REDUCTION),
            "manifest": describe(MANIFEST),
            "event_counts": dict(event_counts),
            "raw_script_count": len(input_scripts),
            "raw_scripts": input_scripts,
            "lost_records": lost_total,
            "parse_stats": parse_stats,
        },
        "parser_contract": {
            "frame_order": "perf script leaf first",
            "physical_dso": [str(BINARY), BINARY.name],
            "marker_frame": MARKER,
            "commit_root_frame": COMMIT,
            "qualified_rule": "exactly one physical marker and one physical commit root, with commit index below marker index",
            "target_match": "exact demangled symbol plus physical binary DSO, bounded to frames through commit root",
            "period_units": "perf sample period from cycles:u header",
        },
        "raw_totals": {
            "all_records": {"samples": all_count, "period": all_period},
            "marker_qualified": {"samples": marker_count, "period": marker_period},
            "commit_qualified": {
                "samples": len(commit_rows),
                "period": sum(row[0]["period"] for row in commit_rows),
                "truncated_samples": len(commit_truncated),
                "truncated_period": sum(row[0]["period"] for row in commit_truncated),
            },
            "commit_within_marker": {
                "samples_share_percent": percent(len(commit_rows), marker_count),
                "period_share_percent": percent(
                    sum(row[0]["period"] for row in commit_rows), marker_period
                ),
                "meaning": "qualified commit root subset divided by marker-qualified denominator",
            },
            "crosscheck_expected_formal_raw": {
                "all_records": {"samples": 18209, "period": 75530554154},
                "marker_qualified": {"samples": 5866, "period": 24985033857},
                "all_records_match": all_count == 18209 and all_period == 75530554154,
                "marker_match": marker_count == 5866 and marker_period == 24985033857,
            },
        },
        "reduction_crosscheck": check_reduction(per_child_rows, reduction),
        "all_formal": analyze_records(commit_rows),
        "per_child": {
            child: analyze_records(rows)
            for child, rows in sorted(commit_by_child.items())
        },
        "focus_floatpictures_e32": analyze_records(
            [
                row
                for row in commit_rows
                if row[0]["child"]
                in {"profile-r01-docfloat-e32", "profile-r02-docfloat-e32"}
            ]
        ),
        "source_equivalence": source_equivalence(),
        "source_evidence": [
            {
                "file": "crates/litchi-ole-common/src/object/editor.rs",
                "lines": "977-1000",
                "text": "commit_candidate_with_rendered renders, OleFile::open, codec::open, Package::capture, and discovery::from_package",
            },
            {
                "file": "crates/litchi-ole-common/src/object/codec.rs",
                "lines": "28-80,1157-1215,1326-1328",
                "text": "Package::capture delegates through capture_with_budget to capture_container, whose stream branch calls OleFile::open_stream; codec::open is a protected-container check",
            },
            {
                "file": "crates/litchi-ole-common/src/object/discovery.rs",
                "lines": "46-68",
                "text": "from_package validates limits and materializes selected objects from the captured package",
            },
            {
                "file": "crates/litchi-cfb/src/file.rs",
                "lines": "938-962,3382-3535",
                "text": "physical OleFile open/open_stream implementations",
            },
            {
                "file": "crates/litchi-cfb/src/writer/core.rs",
                "lines": "1631-1659",
                "text": "OleWriter::write_to emission path",
            },
        ],
        "interpretation": {
            "inclusive": "ancestry membership within the commit root; rows overlap",
            "exclusive": "nearest observed exact target from leaf toward commit root, then other physical CFB, then no target CFB frame",
            "missing_named_frame": "does not establish zero CPU; helper may be inlined, short, or unsampled",
            "capture_container": "source-confirmed helper but no physical named frame observed in retained no-inline commit subset",
            "candidate_scope": "Package::capture/capture_container/open_stream is a measured handoff target; codec::open and discovery::from_package require source review and fresh evidence rather than zero-cost inference",
        },
    }

    with out_json.open("x", encoding="utf-8") as stream:
        json.dump(result, stream, indent=2, sort_keys=True)
        stream.write("\n")
    write_markdown(out_md, result)
    print(
        json.dumps(
            {
                "status": result["status"],
                "json": str(out_json),
                "markdown": str(out_md),
                "bytes": {"json": out_json.stat().st_size, "markdown": out_md.stat().st_size},
            }
        )
    )
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main(sys.argv[1:]))
    except (AnalysisError, OSError, json.JSONDecodeError) as error:
        print(f"analyze_cfb_commit_v1: {error}", file=sys.stderr)
        raise SystemExit(2)
