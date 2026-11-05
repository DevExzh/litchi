#!/usr/bin/env python3
"""M0a S1 / S10 capture driver for the plan 0003 section 5.3 routes.

Runs every route group as a fresh child process of a fixed perf-baseline
binary, pinned to one CPU, wrapped in GNU time for whole-child peak RSS, and
records load average before/after each child. Groups run round-robin over
repeats (R1: all groups, R2: all groups, ...) so slow drift on the shared
host spreads over every group instead of landing on one.

S10 reuses this driver unchanged: pass the pre-M0a binary and the post-S9
binary alternately (ABBA) with --label and a separate --out directory.

Usage:
  capture.py --binary PATH --out DIR --repeats N --cpu C [--groups g1,g2]
             [--samples N] [--warmup N] [--slow-samples N] [--slow-warmup N]
             [--first-repeat N] [--label TEXT]
"""

import argparse
import hashlib
import json
import os
import subprocess
import sys
import time
from datetime import datetime, timezone

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", "..", "..", ".."))
OLE2 = [
    "test-data/ole/doc/documentProperties.doc",
    "test-data/ole/doc/FloatingPictures.doc",
    "test-data/ole/doc/duplicate-style-names.doc",
    "test-data/ole/ppt/ppt_with_png.ppt",
    "test-data/ole/ppt/SampleShow.ppt",
]

FACADE_CASES = [
    "doc_facade_file_open", "doc_facade_file_full_text",
    "doc_facade_file_one_paragraph", "ppt_facade_file_open",
    "ppt_facade_file_full_text", "ppt_facade_file_one_slide_text",
]
# The detection runner's results collide in the schema-2 corpus catalog
# ("content id ... maps to conflicting V1 corpus objects"), so it runs without
# --corpus-manifest; its per-row corpus identity stays in the report.
NO_MANIFEST = {"detection"}
# The *_file_* routes spawn a fresh child per sample. Only the warm state is
# captured: "cold-requested" is a request, not verified eviction.
FILESYSTEM_GROUPS = {"pptx_file", "docx_file", "odt_file", "xlsx", "odp_file", "ods_file"}

# Route group -> (cases, extra args). Cases are the real selector names found
# in tools/perf-baseline/src/lib.rs at the pre-M0a commit; see README.md for
# the mapping from the plan's wildcard names.
GROUPS = {
    "pptx_file": ([
        "pptx_file_eager_open", "pptx_file_source_open",
        "pptx_file_eager_list_slides", "pptx_file_source_list_slides",
        "pptx_file_eager_slide_count", "pptx_file_source_slide_count",
        "pptx_file_eager_selected_slide", "pptx_file_source_selected_slide",
        "pptx_file_eager_open_slide_count_lifecycle",
        "pptx_file_source_open_slide_count_lifecycle",
        "pptx_file_eager_open_selected_slide_lifecycle",
        "pptx_file_source_open_selected_slide_lifecycle",
    ], []),
    "docx_file": ([
        "docx_file_eager_open", "docx_file_source_open",
        "docx_file_eager_paragraph_count", "docx_file_source_paragraph_count",
        "docx_file_eager_list_paragraphs", "docx_file_source_list_paragraphs",
        "docx_file_eager_full_text", "docx_file_source_full_text",
        "docx_file_eager_open_paragraph_count_lifecycle",
        "docx_file_source_open_paragraph_count_lifecycle",
        "docx_file_eager_open_full_text_lifecycle",
        "docx_file_source_open_full_text_lifecycle",
    ], []),
    "docx_read_one": ([
        "docx_semantic_one_paragraph", "docx_semantic_one_paragraph_text",
        "docx_producer_source_selected_paragraph",
    ], []),
    "odt_file": ([
        "odt_file_eager_open", "odt_file_source_open",
        "odt_file_eager_open_full_text_lifecycle",
        "odt_file_source_open_full_text_lifecycle",
    ], []),
    "xlsx": ([
        "xlsx_file_open", "xlsx_file_open_lifecycle", "xlsx_file_selected_cell",
        "xlsx_bytes_open", "xlsx_bytes_open_lifecycle",
    ], []),
    "rtf": (["rtf_file_open", "rtf_file_open_lifecycle"], []),
    # --ole2-file admits one file per format per child, so the five 0638
    # fixtures are split across three children.
    "ole2_facade_small": (FACADE_CASES, ["--ole2-file", OLE2[0], "--ole2-file", OLE2[3]]),
    "ole2_facade_large": (FACADE_CASES, ["--ole2-file", OLE2[1], "--ole2-file", OLE2[4]]),
    "ole2_facade_refusal": (FACADE_CASES[:3], ["--ole2-file", OLE2[2]]),
    "odp_file": ([
        "odp_file_eager_open", "odp_file_source_open",
        "odp_file_eager_selected_slide", "odp_file_source_selected_slide",
    ], []),
    "ods_file": ([
        "ods_file_eager_open", "ods_file_source_open",
        "ods_file_eager_selected_cell", "ods_file_source_selected_cell",
        "ods_file_eager_selected_media", "ods_file_source_selected_media",
    ], []),
    "detection": ([
        f"detect_{fmt}_{kind}"
        for fmt in ("odt", "ods", "odp")
        for kind in ("bytes", "reader", "polyglot", "catalog_alias")
    ], []),
}


def sha256_file(path):
    digest = hashlib.sha256()
    with open(path, "rb") as handle:
        for block in iter(lambda: handle.read(1 << 20), b""):
            digest.update(block)
    return digest.hexdigest()


def now():
    return datetime.now(timezone.utc).isoformat()


def loadavg():
    with open("/proc/loadavg") as handle:
        return handle.read().split()[:3]


def parse_time_file(path):
    out = {}
    with open(path) as handle:
        for line in handle:
            line = line.strip()
            if line.startswith("Maximum resident set size (kbytes):"):
                out["peak_rss_kib"] = int(line.rsplit(":", 1)[1])
            elif line.startswith("Elapsed (wall clock) time"):
                out["wall_clock"] = line.rsplit(": ", 1)[1]
            elif line.startswith("Voluntary context switches:"):
                out["voluntary_ctx_switches"] = int(line.rsplit(":", 1)[1])
            elif line.startswith("Involuntary context switches:"):
                out["involuntary_ctx_switches"] = int(line.rsplit(":", 1)[1])
    return out


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--binary", required=True)
    parser.add_argument("--out", required=True)
    parser.add_argument("--repeats", type=int, default=3)
    parser.add_argument("--first-repeat", type=int, default=1)
    parser.add_argument("--cpu", required=True)
    parser.add_argument("--groups", default=",".join(GROUPS))
    parser.add_argument("--samples", type=int, default=100)
    parser.add_argument("--warmup", type=int, default=10)
    # pptx_file/docx_file spawn a fresh child per sample (~1 s each).
    parser.add_argument("--slow-groups", default="pptx_file,docx_file")
    parser.add_argument("--slow-samples", type=int, default=30)
    parser.add_argument("--slow-warmup", type=int, default=5)
    parser.add_argument("--label", default="before")
    args = parser.parse_args()

    groups = args.groups.split(",")
    slow = set(args.slow_groups.split(","))
    binary_sha = sha256_file(args.binary)
    os.makedirs(os.path.join(args.out, "raw"), exist_ok=True)
    log_path = os.path.join(args.out, "children.jsonl")
    for repeat in range(args.first_repeat, args.first_repeat + args.repeats):
        for group in groups:
            cases, extra = GROUPS[group]
            stem = f"R{repeat}-{group}"
            json_out = os.path.join(args.out, "raw", stem + ".json")
            manifest_out = os.path.join(args.out, "raw", stem + ".corpus-manifest.json")
            time_out = os.path.join(args.out, "raw", stem + ".time")
            stderr_out = os.path.join(args.out, "raw", stem + ".stderr")
            for path in (json_out, manifest_out):
                if os.path.exists(path):
                    sys.exit(f"refusing to overwrite {path}")
            argv = [
                "taskset", "-c", args.cpu,
                "/usr/bin/time", "-v", "-o", time_out,
                args.binary,
                "--case", ",".join(cases),
                "--warmup", str(args.slow_warmup if group in slow else args.warmup),
                "--samples", str(args.slow_samples if group in slow else args.samples),
                "--workers", "1",
                *extra,
                "--json", json_out,
            ]
            if group in FILESYSTEM_GROUPS:
                argv += ["--filesystem-cache", "warm"]
            if group not in NO_MANIFEST:
                argv += ["--corpus-manifest", manifest_out]
            record = {
                "label": args.label,
                "repeat": repeat,
                "group": group,
                "argv": argv,
                "cwd": REPO,
                "binary_sha256": binary_sha,
                "started_utc": now(),
                "loadavg_before": loadavg(),
            }
            started = time.monotonic()
            with open(stderr_out, "wb") as stderr:
                proc = subprocess.run(argv, cwd=REPO, stdout=subprocess.DEVNULL, stderr=stderr,
                                      env={**os.environ, "LC_ALL": "C", "TZ": "UTC"})
            record.update({
                "finished_utc": now(),
                "seconds": round(time.monotonic() - started, 3),
                "loadavg_after": loadavg(),
                "exit_code": proc.returncode,
            })
            if os.path.exists(time_out):
                record["gnu_time"] = parse_time_file(time_out)
            with open(log_path, "a") as handle:
                handle.write(json.dumps(record) + "\n")
            print(f"{stem}: exit={proc.returncode} {record['seconds']}s "
                  f"load={record['loadavg_before'][0]}->{record['loadavg_after'][0]}",
                  flush=True)


if __name__ == "__main__":
    main()
