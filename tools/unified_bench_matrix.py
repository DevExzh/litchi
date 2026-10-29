#!/usr/bin/env python3
"""Generate the unified-vs-leaf benchmark matrix (records 0289, 0294, 0306, 0310, 0315, 0326, 0337, 0340, 0344, 0391, 0399, 0403).

The matrix collects every paired measurement recorded by the unified
operations records from 0249 on (`docs/unified-ops/02[4-9]*.md`, `03*.md`)
into one table: format x operation x scenario/size -> leaf, unified, ratio, budget,
status (meets / miss / no-budget / unqualified, or "miss (superseded by NNNN)"), record,
source and, for every miss, the explanation or follow-up the record gives.

Two sources are merged:

* parsed rows, read from the machine-readable evidence under
  `docs/unified-ops/results/` (JSON, JSONL, TSV and the harnesses' text
  outputs; each record's evidence has its own shape, so there is one
  parser per harness family), and
* curated rows, from `docs/unified-ops/results/0289/bench-matrix-curated.json`,
  for measurements that exist only in record prose or tables. Each curated
  row cites its record file and section, with the record's numbers verbatim.

Rows are keyed by (record, series, stage). A parsed row wins over a curated
row with the same key; when their ratios disagree the conflict is listed in
the Markdown output. The curated file also holds `annotations` that attach
the record's explanation/follow-up text to parsed rows.

Conventions:

* ratio = unified / leaf (or the comparison named in the `comparison`
  column); with several trials or runs the worst (largest) value is the
  ratio and `ratio_range` holds min-max. Status uses the worst value.
* Times are in microseconds; `leaf`/`unified` are medians over trials.
* Budgets follow docs/UNIFIED_OPS_GOAL.md "PERFORMANCE BAR": open 1.05 (or
  +20 us, whichever is larger), single-op write 1.10, 1k-op batch 2.0; other
  batch sizes are "scaling (no ratio budget)" unless the record applies 2.0;
  reads/queries use what the record states (the 1.10 review trigger of
  record 0262 for full queries), else none.
* A row whose series is measured again by a later record (or a later stage
  of the same record) is superseded; a superseded miss is reported as
  "miss (superseded by NNNN)" and the open misses are the latest rows.
* Before/after-build records (0250, 0263, 0268, 0289, 0290, 0293) compare
  builds, not unified with leaf: comparison "after/before", budget 1.05.

Records 0285-0326 (record 0294 extended the tool to 0302; record 0306
added the 0298, 0301, 0302, 0303, 0304 and 0306 parsers; record 0310 the
0305, 0309, 0310, 0311, 0312 and 0313 parsers; records 0315, 0326, 0337 and 0340 their own;
record 0344 those of 0307, 0308, 0314, 0316-0318, 0323-0325, 0328-0330,
0333, 0335, 0336, 0338 and 0342; record 0348 and 0356 their own) are collected only when their record file
is on this tree; the range ends at the highest record on the tree, and the
Markdown lists every record of it, absent ones as "not on this tree", so
the tool can be rerun unchanged after later merges.

Outputs: docs/unified-ops/results/0289/bench-matrix.md and, per format under
bench-matrix/, the row page <format>.md (record 0344) and the row table
<format>.tsv (record 0348; one bench-matrix.tsv grew towards the limit):
each file stays below the repository's 1 MiB evidence limit, which `--check`
enforces.
`--check` regenerates in memory and exits non-zero when the committed
outputs differ, when a miss lacks an explanation, when an open miss names
neither a follow-up nor an explained residual (record 0337), or when a
record's evidence under `results/` (a `<record>/` directory or a
`<record>-*` entry, records 0249 on) contributes no rows and the curated
`evidence_without_rows` gives no reason (record 0344).
"""

from __future__ import annotations

import argparse
import csv
import io
import json
import re
import statistics
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
UO = ROOT / "docs" / "unified-ops"
RES = UO / "results"
OUT_DIR = RES / "0289"
CURATED = OUT_DIR / "bench-matrix-curated.json"
# Record 0348: the single row table this tool wrote before the per-format
# tables; `--check` reports it while it is still on the tree.
OLD_TSV = OUT_DIR / "bench-matrix.tsv"
OUT_MD = OUT_DIR / "bench-matrix.md"
# Record 0344: the per-format row tables, one page each, so that every
# generated file stays below the repository's 1 MiB evidence limit.
OUT_FORMAT_DIR = OUT_DIR / "bench-matrix"
SIZE_LIMIT = 1 << 20

# Budgets: (numeric value or None, text, kind)
OPEN = (1.05, "1.05 or +20 µs (open)", "open")
SINGLE = (1.10, "1.10 (single-op write)", "single")
BATCH = (2.0, "2.0 (batch vs one leaf edit)", "batch")
SCALING = (None, "scaling (no ratio budget)", "scaling")
NONE = (None, "none", "none")
QUERY_TRIGGER = (1.10, "1.10 (full-query review trigger, record 0262 §3)", "read")
XLS_READ = (1.10, "1.10 (review trigger; record 0259: target near 1×)", "read")


def regression(record: str):
    return (1.05, f"1.05 (no regression above 5%, record {record})", "regression")


TSV_COLUMNS = [
    "format", "operation", "scenario", "size", "comparison", "stage",
    "leaf", "unified", "unit", "ratio", "ratio_range", "budget", "status",
    "record", "source", "explanation",
]


# ---------------------------------------------------------------- helpers

def rel(path: Path) -> str:
    return path.relative_to(RES).as_posix()


def load_json(path: Path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def load_jsonl(path: Path):
    rows = []
    with open(path, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line.startswith("{"):
                rows.append(json.loads(line))
    return rows


def med(values):
    vals = [v for v in values if v is not None]
    return statistics.median(vals) if vals else None


def us(ns):
    return None if ns is None else ns / 1000.0


def row(record, fmt, op, scenario, size, leaf, unified, ratios, budget, series,
        source, stage="", stage_rank=1, comparison="unified/leaf", unit="µs",
        explanation=""):
    ratios = [r for r in ratios if r is not None]
    return {
        "record": record, "format": fmt, "operation": op,
        "scenario": scenario, "size": "" if size is None else str(size),
        "leaf": leaf, "unified": unified, "unit": unit,
        "ratio": max(ratios) if ratios else None,
        "ratio_min": min(ratios) if ratios else None,
        "budget_value": budget[0], "budget": budget[1], "budget_kind": budget[2],
        "series": series, "stage": stage, "stage_rank": stage_rank,
        "comparison": comparison, "source": source,
        "explanation": explanation, "verdict": None,
    }


def trials(paths):
    return [p for p in sorted(paths) if p.is_file()]


def src(paths):
    paths = sorted(paths)
    if not paths:
        return "parsed"
    if len(paths) == 1:
        return "parsed:" + rel(paths[0])
    first = rel(paths[0])
    return f"parsed:{first} (+{len(paths) - 1} more)"


def group_cases(docs, key_fn):
    """Group per-trial records by key, preserving first-seen order."""
    groups: dict = {}
    for doc in docs:
        for item in doc:
            k = key_fn(item)
            if k is None:
                continue
            groups.setdefault(k, []).append(item)
    return groups


# ---------------------------------------------------------------- parsers

def p0250():
    path = RES / "0250-xlsx-w2-owner" / "abba.json"
    d = load_json(path)
    out = []
    for r in d["results"]:
        case = r["case"]
        op = ("open" if "open" in case else "save" if "save" in case else "commit")
        out.append(row("0250", "xlsx", op, case, r["fixture"],
                       us(r["base_p50_median_ns"]), us(r["new_p50_median_ns"]),
                       [r["new_p50_median_ns"] / r["base_p50_median_ns"]],
                       regression("0250"), f"abba0250:{case}:{r['fixture']}",
                       src([path]), comparison="after/before (leaf)"))
    return out


DOCX_ELEMENT_OPS = {
    "cell_paragraph_align": "set", "cell_text": "set",
    "hyperlink_run_bold": "set", "table_add_3x3": "add",
    "hyperlink_add_url": "add", "comment_add_paragraph": "add",
    "picture_add_png": "add", "list_style_bullet": "set",
}


def docx_element(record, paths):
    docs = [load_json(p)["cases"] for p in paths]
    out = []
    groups = group_cases(docs, lambda c: None if c.get("skipped") else c["case"])
    for case, items in groups.items():
        out.append(row(record, "docx", DOCX_ELEMENT_OPS.get(case, "set"), case, "",
                       us(med(i["leaf_median_ns"] for i in items)),
                       us(med(i["unified_median_ns"] for i in items)),
                       [i["ratio"] for i in items], SINGLE,
                       f"docx_element_set:{case}", src(paths)))
    return out


def p0251():
    out = []
    for record, d in [("0251a", "0251a-docx-element-writes"),
                      ("0251b", "0251b-docx-hyperlink-writes"),
                      ("0251c", "0251c-docx-comments"),
                      ("0251d", "0251d-docx-pictures"),
                      ("0251e", "0251e-docx-lists")]:
        out += docx_element(record, trials((RES / d).glob("trial-*.json")))
    return out


def docx_fanout(record, paths, stage="", stage_rank=1):
    docs = [load_json(p)["cases"] for p in paths]
    out = []
    for n, items in group_cases(docs, lambda c: c["paragraphs"]).items():
        scen = "set /body/p[1] text" if n == 1 else "set /body > p text (fanout)"
        out.append(row(record, "docx", "set" if n == 1 else "set (fanout)", scen,
                       f"{n} paragraphs",
                       us(med(i["leaf_median_ns"] for i in items)),
                       us(med(i["unified_median_ns"] for i in items)),
                       [i["ratio"] for i in items], SINGLE,
                       f"docx_fanout:{n}", src(paths), stage, stage_rank))
    return out


def p0251f():
    return docx_fanout("0251f", trials((RES / "0251f-docx-fanout").glob("trial-*.json")))


def pptx_run_format(record, paths, stage="", stage_rank=1):
    docs = [load_json(p)["results"] for p in paths]
    out = []
    for (owner, n), items in group_cases(docs, lambda r: (r["owner"], r["shape_count"])).items():
        out.append(row(record, "pptx", "set", f"run format ({owner} owner)",
                       f"{n} shapes",
                       us(med(i["native_median_ns"] for i in items)),
                       us(med(i["unified_median_ns"] for i in items)),
                       [i["unified_over_native"] for i in items], SINGLE,
                       f"pptx_run_format:{owner}:{n}", src(paths), stage, stage_rank,
                       comparison="unified/native"))
    return out


def p0252():
    return pptx_run_format("0252", trials((RES / "0252-pptx-p0-writes").glob("run-format-trial-*.json")))


META_LINE = re.compile(r"^(\w+) set /title \((\w+) open\): native (\d+) ns, unified (\d+) ns, ratio ([\d.]+)")
FIRST_LINE = re.compile(r"^Pptx open \+ first set /title .*: native (\d+) ns, unified (\d+) ns, ratio ([\d.]+)")
APP_LINE = re.compile(r"^xlsx open: (\d+) ns; app\.xml read (\d+) ns")


def ooxml_metadata(record, paths):
    acc: dict = {}
    for p in paths:
        for line in Path(p).read_text(encoding="utf-8").splitlines():
            m = META_LINE.match(line)
            if m:
                fmt, owner = m.group(1).lower(), m.group(2)
                acc.setdefault((fmt, "metadata set", f"set /metadata title ({owner} open)",
                                f"ooxml_metadata:{fmt}:{owner}"), []).append(
                    (int(m.group(3)), int(m.group(4)), float(m.group(5))))
                continue
            m = FIRST_LINE.match(line)
            if m:
                acc.setdefault(("pptx", "open + metadata set",
                                "source open + first /metadata set vs leaf eager open + edit",
                                "ooxml_metadata:pptx:open_first_set"), []).append(
                    (int(m.group(1)), int(m.group(2)), float(m.group(3))))
                continue
            m = APP_LINE.match(line)
            if m:
                o, a = int(m.group(1)), int(m.group(2))
                acc.setdefault(("xlsx", "open", "source open + app.xml read (vs open alone)",
                                "ooxml_metadata:xlsx:open_app_read"), []).append(
                    (o, o + a, (o + a) / o))
    out = []
    for (fmt, op, scen, series), items in acc.items():
        budget = OPEN if op == "open" else SINGLE
        out.append(row(record, fmt, op, scen, "", us(med(i[0] for i in items)),
                       us(med(i[1] for i in items)), [i[2] for i in items], budget,
                       series, src(paths)))
    return out


def p0253():
    return ooxml_metadata("0253", [RES / "0253" / "ooxml-metadata-bench.txt"])


def p0254():
    path = RES / "0254" / "ole-metadata-bench.jsonl"
    rows = load_jsonl(path)
    out = []
    for case, items in group_cases([rows], lambda r: r["case"]).items():
        out.append(row("0254", case, "metadata set", "set /metadata title (+to_bytes)",
                       f"{items[0]['bytes']} bytes",
                       med(i["native_us"] for i in items), med(i["unified_us"] for i in items),
                       [i["unified_us"] / i["native_us"] for i in items], SINGLE,
                       f"ole_metadata:{case}", src([path]), comparison="unified/native"))
    return out


PY_SCEN = {
    "get_docx_p10_d0": ("get", "DOCX get(/body/p[10], depth 0)"),
    "get_docx_p10_d1": ("get", "DOCX get(/body/p[10], depth 1)"),
    "get_xlsx_cell_d0": ("get", "XLSX get(/ExcelTable/C2, 0)"),
    "query_docx_paragraph": ("query", "DOCX query(paragraph), 227 matches"),
    "query_xlsx_cell_100": ("query", "XLSX query(cell, 100), 75k-cell sheet"),
    "set_docx_run_bold": ("set", "DOCX set run bold"),
    "set_xlsx_cell": ("set", "XLSX set one cell"),
    "batch_xlsx_1k": ("batch", "XLSX batch of 1k cell sets (dicts)"),
    "batch_xlsx_1k_json": ("batch", "XLSX batch of 1k cell sets (JSON text vs dict Rust leg)"),
    "edit_xlsx_1k": ("batch", "XLSX edit(fn) with 1k tx.set"),
    "batch_docx_1k": ("batch", "DOCX batch of 1k run-bold sets"),
}


def p0255():
    path = RES / "0255-python-overhead.json"
    d = load_json(path)["scenarios"]
    out = []
    for key, s in d.items():
        op, scen = PY_SCEN.get(key, ("call", key))
        rust = s.get("rust_median")
        if rust is None and key == "batch_xlsx_1k_json":
            rust = d["batch_xlsx_1k"]["rust_median"]
        py = s["python_median"]
        out.append(row("0255", "python", op, scen, "", us(rust), us(py), [py / rust],
                       (None, "none (record 0255 §5: no Python budget in the goal)", "none"),
                       f"python:{key}", src([path]), comparison="python/rust"))
    return out


def p0256():
    path = RES / "0256-cli" / "get-cost.json"
    d = load_json(path)["results"]
    out = []
    for fixture, legs in d.items():
        fmt = fixture.split("/")[1]
        leaf = legs["in-process"]["median_us"]
        cli = legs["cli-get"]["median_us"]
        out.append(row("0256", "cli", "get", f"litchi get --json ({fmt}) vs in-process open+get+JSON",
                       fixture.split("/", 2)[2], leaf, cli, [cli / leaf],
                       (None, "none (record 0256 §6: no budget applies to process start)", "none"),
                       f"cli:{fixture}", src([path]), comparison="cli process/in-process"))
    return out


DOC_OPS = {"paragraph_text_set": "set", "run_format_set": "set",
           "paragraph_insert": "add", "paragraph_text_get": "get",
           "run_bold_get": "get"}


def p0258():
    base = RES / "0258-doc-p0"
    out = []

    def p0(paths, stage, rank):
        docs = [load_json(p) for p in paths]
        res = []
        for (scen, n), items in group_cases(docs, lambda r: (r["scenario"], r["paragraphs"])).items():
            op = DOC_OPS.get(scen, "set")
            budget = SINGLE if op in ("set", "add") else NONE
            res.append(row("0258", "doc", op, scen, f"{n} paragraphs",
                           us(med(i["native_median_ns"] for i in items)),
                           us(med(i["unified_median_ns"] for i in items)),
                           [i["ratio"] for i in items], budget,
                           f"doc_p0:{scen}:{n}", src(paths), stage, rank,
                           comparison="unified/native"))
        return res

    out += p0(trials(base.glob("p0-trial-*.json")), "", 1)
    out += p0([base / "p0-before-lookup.json"], "before (pre-fix diagnostic)", 0)
    for tag, stage, rank in (("T", "trunk e467bcf4e8", 0), ("M", "", 1)):
        paths = trials(base.glob(f"frozen-paragraph-{tag}-*.json"))
        docs = [load_json(p)["results"] for p in paths]
        for n, items in group_cases(docs, lambda r: r["paragraphs"]).items():
            out.append(row("0258", "doc", "set",
                           "frozen 0188/0190 route: fresh File per sample, paragraph text",
                           f"{n} paragraphs",
                           us(med(i["native_median_ns"] for i in items)),
                           us(med(i["unified_median_ns"] for i in items)),
                           [i["unified_over_native"] for i in items], SINGLE,
                           f"doc_frozen_paragraph:{n}", src(paths), stage, rank,
                           comparison="unified/native"))
    return out


def xls_read(record, paths, stage="", stage_rank=1):
    docs = [load_json(p)["cases"] for p in paths]
    out = []
    for (cells, owner), items in group_cases(docs, lambda c: (c["cells"], c["owner"])).items():
        for kind, label in (("get", "get each cell"), ("scan", "query all cells")):
            out.append(row(record, "xls", "get" if kind == "get" else "query",
                           f"{label} ({owner} owner)", f"{cells} cells",
                           us(med(i[f"median_native_{kind}_ns"] for i in items)),
                           us(med(i[f"median_unified_{kind}_ns"] for i in items)),
                           [i[f"{kind}_ratio"] for i in items], XLS_READ,
                           f"xls_read:{kind}:{owner}:{cells}", src(paths), stage, stage_rank,
                           comparison="unified/native"))
    return out


def p0259():
    base = RES / "0259-xls-unified-p0"
    out = xls_read("0259", [base / "read-before.json"], "before (adapter changes)", 0)
    out += xls_read("0259", [base / "read-after.json"])
    path = base / "write-p0.json"
    for c in load_json(path)["cases"]:
        scen = c["scenario"]
        out.append(row("0259", "xls", "set" if scen not in ("sheet_add", "sheet_move") else
                       scen.split("_")[1], f"{scen} ({c['owner']} owner)", f"{c['cells']} cells",
                       us(c["median_native_ns"]), us(c["median_unified_ns"]), [c["ratio"]],
                       SINGLE, f"xls_p0_write:{scen}:{c['owner']}:{c['cells']}", src([path]),
                       comparison="unified/native"))
    return out


PPT_PAIRED_OPS = {
    "single_set_eager": ("set", "shape text set, warm owner"),
    "single_set_eager_cold": ("set", "shape text set, cold owner"),
    "text_batch_eager": ("batch", "text batch (one atomic edit)"),
    "outline_title_set_sampleshow": ("set", "outline title set (SampleShow)"),
    "outline_body_set_41246": ("set", "outline body set (41246-1)"),
    "paragraph_set_basic": ("set", "paragraph set (basic_test)"),
    "notes_set_basic": ("set", "notes set (basic_test)"),
    "slide_add_sampleshow": ("add", "slide add (SampleShow)"),
    "slide_add_41246": ("add", "slide add (41246-1)"),
}


def ppt_paired_row(record, scen, size, leaf, unified, ratios, source, stage, rank):
    op, label = PPT_PAIRED_OPS.get(scen, ("set", scen))
    if op == "batch":
        budget, sz = BATCH, f"{size} ops"
    elif scen.startswith("single"):
        budget, sz = SINGLE, f"{size} shapes"
    else:
        budget, sz = SINGLE, f"{size} bytes"
    return row(record, "ppt", op, label, sz, leaf, unified, ratios, budget,
               f"ppt_paired:{scen}:{size}", source, stage, rank, comparison="unified/native")


def p0260():
    path = RES / "0260-ppt-paired.json"
    res = load_json(path)["results"]
    out = []
    for stage, rank in (("before", 0), ("after", 2)):
        for mode in ("single", "single-cold", "batch", "paths"):
            key = f"{mode}-{stage}"
            for r in res.get(key, []):
                ratios = [r["pair_ratio_median"]]
                if mode == "paths":
                    for k, v in res.items():
                        if re.fullmatch(rf"paths-{stage}-r\d+", k):
                            ratios += [x["pair_ratio_median"] for x in v
                                       if x["scenario"] == r["scenario"]]
                out.append(ppt_paired_row("0260", r["scenario"], r["size"],
                                          r["native_median_us"], r["unified_median_us"],
                                          ratios, src([path]), stage, rank))
    return out


def xlsb_meta(record, paths):
    docs = [load_jsonl(p) for p in paths]
    out = []
    for (corpus, pair), items in group_cases(docs, lambda r: (r["corpus"], r["pair"])).items():
        op = "metadata set + save" if pair.startswith("save:") else "metadata set"
        out.append(row(record, "xlsb", op, f"/metadata title, {pair}", corpus,
                       us(med(i["leaf_ns"] for i in items)), us(med(i["unified_ns"] for i in items)),
                       [i["ratio"] for i in items], SINGLE,
                       f"xlsb_metadata_source:{corpus}:{pair}", src(paths),
                       comparison=pair.replace("save: ", "")))
    return out


def p0261():
    base = RES / "0261"
    out = xlsb_meta("0261", trials(base.glob("meta-trial-*.jsonl")))
    paths = trials(base.glob("p0-trial-*.jsonl"))
    docs = [load_jsonl(p) for p in paths]
    for (cells, owner, op), items in group_cases(docs, lambda r: (r["cells"], r["owner"], r["op"])).items():
        out.append(row("0261", "xlsb", "set" if op != "name" else "add", f"{op} ({owner} owner)",
                       f"{cells} cells", us(med(i["leaf_ns"] for i in items)),
                       us(med(i["unified_ns"] for i in items)), [i["ratio"] for i in items],
                       SINGLE, f"xlsb_p0_write:{op}:{owner}:{cells}", src(paths)))
    return out


def query_routes(record, paths, stage="", stage_rank=1, build=""):
    docs = [load_jsonl(p) for p in paths]
    out = []
    for route, items in group_cases(docs, lambda r: r["route"]).items():
        name = re.sub(r"_v\d+$", "", route)
        fmt = name.split("_")[0]
        parts = name.split("_")
        if "get" in parts:
            op, budget = "get", NONE
        elif "_all_" in name:
            op, budget = "query (full)", QUERY_TRIGGER
        else:
            op, budget = "query (early stop)", NONE
        tag = f"[{build}]" if build else ""
        out.append(row(record, fmt, op, name + (f" ({build} build)" if build else ""),
                       f"{items[0]['matches']} matches",
                       us(min(i["leaf_median_ns"] for i in items)),
                       us(min(i["unified_median_ns"] for i in items)),
                       [i["ratio_median"] for i in items], budget, f"query{tag}:{name}",
                       src(paths), stage, stage_rank))
    return out


def p0262():
    base = RES / "0262"
    out = query_routes("0262", trials(base.glob("query-*-query-after.jsonl")))
    out += xls_read("0262", trials(base.glob("xls-after-*.json")))
    out += pptx_run_format("0262", trials(base.glob("runfmt-*-runfmt-after.json")))
    out += ooxml_metadata("0262", trials(base.glob("meta-*-meta-after.txt")))
    return out


def p0264():
    path = RES / "0264-ppt-scale.json"
    summary = load_json(path)["summary"]
    out = []
    for key, s in summary.items():
        if key.count("/") != 2:
            continue  # experiment and native-phase entries carry no unified pair
        _, scen, size = key.split("/")
        out.append(ppt_paired_row("0264", scen, int(size), s["native_after_us"],
                                  s["unified_after_us"], [s["pair_ratio_after"]],
                                  src([path]), "after (4096 bound)", 1))
    path = RES / "0264-ppt-bound.json"
    raw = load_json(path)["raw"]
    for prefix, stage, rank in (("batch-after-", "after (4096 bound, bound run)", 2),
                                ("batch-bound16k-", "16384 bound (owner-approved)", 3)):
        groups: dict = {}
        for k, v in raw.items():
            if k.startswith(prefix):
                for r in v:
                    groups.setdefault(r["size"], []).append(r)
        for size, items in sorted(groups.items()):
            out.append(ppt_paired_row("0264", "text_batch_eager", size,
                                      med(i["native_median_us"] for i in items),
                                      med(i["unified_median_us"] for i in items),
                                      [i["pair_ratio_median"] for i in items],
                                      src([path]), stage, rank))
    return out


def scenarios_us(record, paths, mapping):
    paths = [paths] if isinstance(paths, Path) else paths
    docs = [list(load_json(p)["scenarios"].items()) for p in paths]
    out = []
    for key, items in group_cases(docs, lambda kv: kv[0]).items():
        items = [s for _, s in items]
        fmt, op, scen, size, budget, series = mapping(key, items[0])
        out.append(row(record, fmt, op, scen, size, med(s["leaf_us"]["median"] for s in items),
                       med(s["unified_us"]["median"] for s in items),
                       [s["ratio_of_medians"] for s in items], budget, series, src(paths)))
    return out


def p0266():
    def m(key, s):
        kind, fmt = key.split("/")
        if kind == "import":
            rows = s.get("rows")
            return (fmt, "import", "CSV import (one atomic commit)", f"{rows} rows × 5",
                    (None, "none (record 0266 states no budget)", "none"), f"m4:{key}")
        return (fmt, "recalculate", "evaluate uncached formulas", f"{s.get('formulas')} formulas",
                (None, "none (record 0266 states no budget)", "none"), f"m4:{key}")
    return scenarios_us("0266", RES / "0266-csv-eval" / "m4-bench.json", m)


def p0267():
    base = RES / "0267"
    out = query_routes("0267", trials(base.glob("query-*-after.jsonl")))
    out += xls_read("0267", trials(base.glob("xls-*-after.json")))
    return out


def p0268():
    out = []
    for sub, dest in (("runs-tmpfs", "tmpfs"), ("runs-disk", "ext4")):
        acc: dict = {}
        files = trials((RES / "0268" / sub).glob("*.tsv"))
        for f in files:
            tag = "before" if "before" in f.name else "after"
            for line in Path(f).read_text(encoding="utf-8").splitlines():
                if not line.strip():
                    continue
                k, v = line.split("\t")
                acc.setdefault(k, {}).setdefault(tag, []).append(int(v))
        for key, legs in acc.items():
            fmt, op = key.split(".")
            if dest == "ext4" and not op.startswith("save"):
                continue
            if "before" not in legs:
                continue
            b, a = med(legs["before"]), med(legs["after"])
            scen = f"leaf {op}" + (f" ({dest})" if op.startswith("save") else "")
            out.append(row("0268", fmt, op.replace("_", " "), scen, "",
                           us(b), us(a), [a / b],
                           (None, "none (record 0268: leaf before/after check)", "none"),
                           f"odf_w3:{key}:{dest}", src(files), comparison="after/before (leaf)"))
    return out


CORPUS_PARITY = {
    "xlsx_apply_refuse_vs_warn": ("xlsx", "set", "apply_with one cell, nothing refused",
                                  "warn/refuse", (None, "none", "none")),
    "xlsx_apply_subset_vs_warn_one_refused": ("xlsx", "set", "apply_with, one refused property (subset under Refuse vs Warn)",
                                              "warn/refuse", (None, "none (record 0281 §5: no budget covers the warn retry)", "none")),
    "xlsx_batch_100_refuse_vs_warn": ("xlsx", "batch", "100-cell batch, Warn vs coalesced Refuse",
                                      "warn/refuse", BATCH),
    "pptx_add_formatted_text_box": ("pptx", "add", "add text box with size, bold, colour, fill, geometry",
                                    "unified/leaf", SINGLE),
    "pptx_add_plain_text_box": ("pptx", "add", "add plain text box", "unified/leaf", SINGLE),
    "pptx_set_shape_fill_geometry": ("pptx", "set", "set shape fill + geometry", "unified/leaf", SINGLE),
    "pptx_set_slide_transition": ("pptx", "set", "set slide transition", "unified/leaf", SINGLE),
}


def corpus_parity(record, paths, stage="", stage_rank=1):
    docs = [load_jsonl(p) for p in paths]
    out = []
    for (scen, size), items in group_cases(docs, lambda r: (r["scenario"], r["size"])).items():
        fmt, op, label, comp, budget = CORPUS_PARITY[scen]
        sz = f"{size} shapes" if fmt == "pptx" else (f"{size} cells")
        out.append(row(record, fmt, op, label, sz,
                       us(med(i["leaf_or_a_median_ns"] for i in items)),
                       us(med(i["unified_or_b_median_ns"] for i in items)),
                       [i["ratio"] for i in items], budget,
                       f"corpus_parity:{scen}:{size}", src(paths), stage, stage_rank,
                       comparison=comp))
    return out


def p0269():
    return corpus_parity("0269", trials((RES / "0269-corpus-parity").glob("trial-*.jsonl")))


def chart(record, paths, kind):
    docs = [load_json(p)["results"] for p in paths]
    out = []
    nkey = "existing_charts"
    for (route, n), items in group_cases(docs, lambda r: (r["route"], r[nkey])).items():
        fmt = route.split("_")[0]
        name = re.sub(r"_v\d+$", "", route)
        if kind == "dotted":
            leaf = us(med(i["data_median_ns"] for i in items))
            uni = us(med(i["dotted_median_ns"] for i in items))
            ratios = [i["dotted_over_data"] for i in items]
            comp, op = "dotted/data=", "chart add"
            series = f"chart_add_dotted:{name}:{n}"
        else:
            leaf = us(med(i["native_median_ns"] for i in items))
            uni = us(med(i["unified_median_ns"] for i in items))
            ratios = [i["unified_over_native"] for i in items]
            comp = "unified/native"
            op = "chart set" if "_set" in name else "chart add"
            series = f"chart:{name}:{n}"
        out.append(row(record, fmt, op, name, f"{n} charts", leaf, uni, ratios, SINGLE,
                       series, src(paths), comparison=comp))
    return out


def p0270():
    base = RES / "0270"
    return chart("0270", [base / "chart-add.json", base / "chart-add-repeat.json"], "native")


def p0271():
    base = RES / "0271"
    return (chart("0271", trials(base.glob("dotted-*.json")), "dotted")
            + chart("0271", [base / "native.json"], "native"))


def find_replace(record, paths):
    docs = [load_json(p)["cases"] for p in paths]
    out = []
    for (fmt, cont, matches), items in group_cases(
            docs, lambda c: (c["format"], c["containers"], c["matches"])).items():
        unit = "paragraphs" if fmt.startswith("docx") else "slides"
        scen = f"find/replace, {matches} matches"
        if fmt == "docx-paragraph":
            scen = f"find/replace scoped to one paragraph (/body/p[5001] of 10,000), {matches} matches"
        out.append(row(record, fmt.split("-")[0], "set (find/replace)", scen,
                       f"{cont} {unit}", us(med(i["leaf_median_ns"] for i in items)),
                       us(med(i["unified_median_ns"] for i in items)),
                       [i["ratio"] for i in items], SINGLE,
                       f"find_replace:{fmt}:{cont}:{matches}", src(paths)))
    return out


def p0273():
    return find_replace("0273", [RES / "0273-find-replace-bench.json"])


MERGE_LABELS = {
    "records/auto": ("1,000 records into an 8-placeholder template, source-backed (auto) open", "1000 records"),
    "records/eager": ("1,000 records into an 8-placeholder template, eager open", "1000 records"),
    "placeholders-10k": ("one record into a 2,000-paragraph template with 10,000 placeholders", "10000 placeholders"),
}


def merge_label(key, s):
    scen, size = MERGE_LABELS[key]
    return ("docx", "merge", scen, size, SINGLE, f"merge:{key}")


def p0274():
    return scenarios_us("0274", RES / "0274-merge" / "merge-bench.json", merge_label)


RTF_ROUTES = (("batch", "edit closure (coalesced)"),
              ("fanout", "query fan-out"),
              ("fanout_plan", "query fan-out, selector compiled once into a Plan"),
              ("direct", "direct File::set/remove"),
              # Record 0305 route G: the fan-out passing a borrowed &Query.
              ("fanout_ref", "query fan-out, borrowed &Query"))


def rtf_destination(record, paths):
    docs = [load_json(p)["cases"] for p in paths]
    out = []
    for (fam, count), items in group_cases(docs, lambda c: (c["family"], c["count"])).items():
        for route, label in RTF_ROUTES:
            ratios = [i.get(f"{route}_ratio") for i in items]
            if all(r is None for r in ratios):
                continue
            native_key = "native_shared_median_ns" if route.startswith("fanout") else "native_median_ns"
            budget = SINGLE if count == 1 else BATCH
            op = "remove" if fam == "picture_removal" else "set"
            if count > 1:
                op = "batch " + op
            out.append(row(record, "rtf", op, f"{fam}, {label}", f"{count} targets",
                           us(med(i[native_key] for i in items)),
                           us(med(i.get(f"unified_{route}_median_ns") for i in items)),
                           ratios, budget, f"rtf_destination:{fam}:{count}:{route}",
                           src(paths), comparison="unified/native"))
    return out


def p0275():
    base = RES / "0275-rtf-parity"
    out = rtf_destination("0275", trials(base.glob("destination-trial-*.json")))
    path = base / "tiny-input-abba.json"
    d = load_json(path)["harnesses"]
    for harness, h in d.items():
        for case, c in h["cases"].items():
            kv = dict(part.split("=") for part in case.split(","))
            native = us(med(c["new_native_median_ns"]))
            if harness == "unified-rtf-add":
                op, scen, size, budget = "add", "add paragraph", f"{kv['paragraphs']} paragraphs", SINGLE
            elif harness == "unified-rtf-run-add":
                op, scen = "add", f"add run ({kv['shape']}, {kv['route']})"
                size, budget = f"{kv['groups']} groups", SINGLE
            elif harness == "unified-rtf-layout":
                n = int(kv["paragraphs"])
                if kv["route"] == "last_path":
                    op, scen, budget = "set", "paragraph layout set, path", SINGLE
                else:
                    op, scen = "set (query fan-out)", "paragraph layout set, query over all paragraphs"
                    budget = SINGLE if n == 1 else BATCH
                size = f"{n} paragraphs"
            else:
                n = int(kv["count"])
                op = "set (edit closure)" if n == 1 else "batch set (edit closure)"
                scen = f"{kv['family']} ({kv['shape']})"
                size, budget = f"{n} targets", SINGLE if n == 1 else BATCH
            out.append(row("0275", "rtf", op, scen, size, native, None, c["new_ratio"], budget,
                           f"rtf_tiny:{harness}:{case}", src([path]),
                           comparison="unified/native"))
    return out


def odt_budget(record, scen, n_ops):
    if scen == "open_first_get":
        return OPEN, "open + first get"
    if scen.startswith("query_all"):
        if record >= "0280":
            return (1.10, "1.10 (review trigger; record 0280 follows up the open PERFORMANCE BAR misses)", "read"), scen
        return (None, "none (record 0276 §6: no numeric query budget)", "none"), scen
    if scen == "single_set":
        return SINGLE, "single set (paragraph text)"
    if n_ops in (1000, 10000):
        return BATCH, f"batch bold ×{n_ops}"
    return SCALING, f"batch bold ×{n_ops}"


def odt_rows(record, items_by_key, source):
    out = []
    for (scen, n), items in items_by_key.items():
        n_ops = int(scen.rsplit("_", 1)[1]) if scen.startswith("batch_bold_") else None
        budget, label = odt_budget(record, scen, n_ops)
        op = ("open" if scen == "open_first_get" else "query (full)" if scen.startswith("query")
              else "set" if scen == "single_set" else "batch set")
        out.append(row(record, "odt", op, label if not scen.startswith("query") else scen,
                       f"{n} paragraphs", med(i[0] for i in items), med(i[1] for i in items),
                       [i[2] for i in items], budget, f"odt_paired:{scen}:{n}", source))
    return out


ODT_LINE = re.compile(r"^\s*(\S+)\s+(\d+)\s+leaf\s+([\d.]+) µs\s+unified\s+([\d.]+) µs\s+ratio\s+([\d.]+)")


def odt_text(record, paths):
    acc: dict = {}
    for p in paths:
        for line in Path(p).read_text(encoding="utf-8").splitlines():
            m = ODT_LINE.match(line)
            if m:
                acc.setdefault((m.group(1), int(m.group(2))), []).append(
                    (float(m.group(3)), float(m.group(4)), float(m.group(5))))
    return odt_rows(record, acc, src(paths))


def p0276():
    paths = trials((RES / "0276").glob("paired-cpu9-run*.json"))
    acc: dict = {}
    for p in paths:
        for r in load_json(p)["rows"]:
            acc.setdefault((r["scenario"], r["paragraphs"]), []).append(
                (r["leaf_us"], r["unified_us"], r["ratio"]))
    return odt_rows("0276", acc, src(paths))


def ods_rows(record, paths):
    docs = [load_json(p)["scenarios"] for p in paths]
    out = []
    for name, items in group_cases(docs, lambda s: s["name"]).items():
        if name == "open_get":
            op, budget = "open", OPEN
        elif name.startswith("query"):
            op = "query (full)"
            budget = ((1.10, "1.10 (review trigger; record 0280 follows up the open PERFORMANCE BAR misses)", "read")
                      if record >= "0280" else (None, "none (record 0277 §6: no budget applies)", "none"))
        elif name == "set":
            op, budget = "set", SINGLE
        else:
            op, budget = "batch set", BATCH
        out.append(row(record, "ods", op, name, "1000 × 10 sheet",
                       us(med(i["leaf_median_ns"] for i in items)),
                       us(med(i["unified_median_ns"] for i in items)),
                       [i["ratio"] for i in items], budget, f"ods_bench:{name}", src(paths)))
    return out


def p0277():
    return ods_rows("0277", trials((RES / "0277").glob("ods-unified-bench-run-*.json")))


def p0278():
    path = RES / "0278" / "odp_unified_bench.tsv"
    lines = open(path, encoding="utf-8").read().split("\n\n")
    summary = list(csv.DictReader(io.StringIO(lines[1].strip()), delimiter="\t"))
    out = []
    for s in summary:
        scen = s["scenario"]
        if scen.startswith("open"):
            op, budget = "open", OPEN
        elif scen.startswith("query"):
            op, budget = "query (full)", (None, "none (record 0278 §6: no budget applies to query cost)", "none")
        elif scen == "single set":
            op, budget = "set", SINGLE
        else:
            n = int(scen.split()[1])
            op = "batch set"
            budget = BATCH if n == 1000 else SCALING
        leaf, uni = float(s["leaf_median_us"]), float(s["unified_median_us"])
        out.append(row("0278", "odp", op, scen, "50 slides × 20 paragraphs" if scen != "batch 10000"
                       else "100 slides × 100 paragraphs", leaf, uni, [uni / leaf], budget,
                       f"odp_bench:{scen}", src([path])))
    return out


def p0279():
    labels = {
        "toc-1000": ("docx", "refresh", "TOC over 1,000 headings, built (1,000 bookmarks added)", "1000 headings"),
        "toc-1000-current": ("docx", "refresh", "TOC already current (exact no-op)", "1000 headings"),
        "xlsx-cache-6k": ("xlsx", "refresh", "6,000 uncached formulas written back", "6000 formulas"),
    }

    def m(key, s):
        fmt, op, scen, size = labels[key]
        return (fmt, op, scen, size, SINGLE, f"refresh:{key}")
    return scenarios_us("0279", RES / "0279-refresh-bench.json", m)


def p0280():
    base = RES / "0280"
    out = query_routes("0280", trials(base.glob("query-*-after.jsonl")))
    # Record 0280 applies the review trigger to every full query, including the
    # fair (owned/fenced) baselines: they are the follow-up of 0262's misses.
    out += odt_text("0280", trials(base.glob("odt-*-after.txt")))
    out += ods_rows("0280", trials(base.glob("ods-*-after.json")))
    out += rtf_destination("0280", trials(base.glob("rtf-*-after.json")))
    return out


def p0281():
    base = RES / "0281"
    out = corpus_parity("0281", [base / "corpus_parity_bench.txt"], "after", 2)
    out += [dict(r, stage="after", stage_rank=2) for r in
            docx_element("0281", [base / "docx_element_set_bench.txt"])]
    out += find_replace("0281", [base / "find_replace_bench.txt"])
    out += [dict(r, stage="after", stage_rank=2) for r in
            ooxml_metadata("0281", [base / "ooxml_metadata_bench.txt"])]
    out += xlsb_meta("0281", [base / "xlsb_metadata_source_bench.txt"])
    out += docx_fanout("0281", [base / "docx_fanout_bench.txt"], "after", 2)
    out += xls_read("0281", [base / "xls_read_bench.txt"])
    path = base / "xls_write_bench.txt"
    for c in load_json(path)["cases"]:
        leaf, uni = med(c["native_ns"]), med(c["unified_set_ns"])
        out.append(row("0281", "xls", "set", "set one cell value (eager)", f"{c['cells']} cells",
                       us(leaf), us(uni), [uni / leaf], SINGLE, f"xls_write:{c['cells']}",
                       src([path]), comparison="unified/native"))
    for r in out:
        if r["stage"] == "":
            r["stage"], r["stage_rank"] = "after", 2
    return out


def p0282():
    path = RES / "0282" / "odf_leaves_bench.json"
    d = load_json(path)
    per_run: dict = {}
    for run in d["runs"]:
        for r in run:
            per_run.setdefault((r["format"], r["scenario"]), []).append(r)
    out = []
    for m in d["medians"]:
        fmt, scen = m["format"], m["scenario"]
        runs = per_run.get((fmt, scen), [])
        budget = OPEN if scen == "open_first_get" else SINGLE
        op = "open" if scen == "open_first_get" else "set"
        if m.get("leaf_us") is not None:
            out.append(row("0282", fmt, op, scen, "", m["leaf_us"], m["unified_us"],
                           [r["ratio_unified_leaf"] for r in runs if r.get("ratio_unified_leaf")]
                           or [m["unified_us"] / m["leaf_us"]], budget,
                           f"odf_leaves:{fmt}:{scen}:leaf", src([path])))
        if m.get("engine_us") is not None:
            out.append(row("0282", fmt, op, scen, "", m["engine_us"], m["unified_us"],
                           [r["ratio_unified_engine"] for r in runs], budget,
                           f"odf_leaves:{fmt}:{scen}:engine", src([path]),
                           comparison="unified/engine (litchi_odt)"))
    return out


def p0283():
    base = RES / "0283"
    return (chart("0283", trials(base.glob("add-*.json")), "native")
            + chart("0283", trials(base.glob("set-*.json")), "native"))


def p0284():
    base = RES / "0284"
    paths = [base / "smartart-writes.json", base / "smartart-writes-repeat.json"]
    docs = [load_json(p)["results"] for p in paths]
    out = []
    for (route, n), items in group_cases(docs, lambda r: (r["route"], r["existing_smartart"])).items():
        name = re.sub(r"_v\d+$", "", route)
        op = "smartart add" if name.endswith("smartart_add") else (
            "smartart node set" if "node_text" in name else "smartart node add")
        out.append(row("0284", "pptx", op, name, f"{n} SmartArt",
                       us(med(i["native_median_ns"] for i in items)),
                       us(med(i["unified_median_ns"] for i in items)),
                       [i["unified_over_native"] for i in items], SINGLE,
                       f"smartart:{name}:{n}", src(paths), comparison="unified/native"))
    return out


# ------------------------------------------------- records 0285-0302
#
# Records in FOLLOW_RANGE are collected only when their record file is on
# this tree (`needs`), so the tool can be rerun unchanged after later merges;
# absent records are listed in the Markdown output as "not on this tree".

def record_file(record: str):
    hits = sorted(UO.glob(f"{record}-*.md"))
    return hits[0] if hits else None


def last_record() -> int:
    """The highest record number on this tree (record 0344: the range
    follows the tree, so a newly merged record is listed, and checked,
    without editing the tool)."""
    numbers = [int(m.group(1)) for p in UO.glob("[0-9][0-9][0-9][0-9]*-*.md")
               if (m := re.match(r"(\d{4})", p.name))]
    return max([340, *numbers])


FOLLOW_RANGE = (285, last_record())


def needs(record: str):
    def wrap(fn):
        def inner():
            return fn() if record_file(record) else []
        inner.__name__ = fn.__name__
        return inner
    return wrap


@needs("0285")
def p0285():
    # The "before" legs (trunk 9fe45629ab) re-measure 0273/0274/0281 states;
    # only the after build is listed (convention: earlier states are not repeated).
    base = RES / "0285"
    return (find_replace("0285", trials(base.glob("find.after*.json")))
            + scenarios_us("0285", trials(base.glob("merge.after*.json")), merge_label))


@needs("0286")
def p0286():
    base = RES / "0286"
    return (ooxml_metadata("0286", trials(base.glob("meta-after-*.txt")))
            + chart("0286", trials(base.glob("chartset-after-*.json")), "native")
            + docx_fanout("0286", trials(base.glob("fanout-after-*.json")))
            + docx_element("0286", trials(base.glob("elem-after-*.json")))
            + query_routes("0286", trials(base.glob("query-after-*.jsonl"))))


@needs("0287")
def p0287():
    base = RES / "0287"
    out = query_routes("0287", trials(base.glob("query-[0-9]-after.jsonl")))
    out += query_routes("0287", trials(base.glob("query-all-*-after.jsonl")), build="all-formats")
    out += rtf_destination("0287", trials(base.glob("rtf-*-after.json")))
    out += ods_rows("0287", trials(base.glob("ods-*-after.json")))
    out += odt_text("0287", trials(base.glob("odt-*-after.txt")))
    out += xls_read("0287", trials(base.glob("xls-*-after.json")))
    return out


ODF_LINE = re.compile(r"^\s*([a-z]+)\s+([a-z_]+\d*)\s+leaf\s+([\d.]+) µs\s+unified\s+([\d.]+) µs\s+ratio\s+([\d.]+)")

ODF_GAPS_BATCH = {"odg": ("shape text", "1000 shapes"),
                  "odb": ("query command", "1000 queries"),
                  "odf": ("token text", "1000 tokens")}
ODF_SINGLE = {
    ("odg", "metadata_set"): ("metadata set", "/metadata set", "3page.odg"),
    ("odg", "insert_page"): ("add", "insert page first", "3page.odg"),
    ("odg", "move_page"): ("move", "move page last", "3page.odg"),
    ("ods", "metadata_set"): ("metadata set", "/metadata set", "calc-two-sheets"),
    ("odp", "metadata_set"): ("metadata set", "/metadata set", "impress-basic"),
    ("odc", "title_set"): ("set", "title set (typed chart)", ""),
    ("odi", "geometry_set"): ("set", "geometry set (fresh image)", ""),
    ("odg", "cross_page_move"): ("move", "cross-page shape move", "2 pages, 100 shapes"),
    ("odb", "query_move"): ("move", "query move", "100 queries"),
    ("odc", "series_swap"): ("swap", "series swap (typed chart)", "10 series"),
    ("odf", "element_move"): ("move", "formula element move", "50 tokens"),
    ("fodt", "open"): ("open", "flat ODT open", "54 KB"),
    ("fodt", "paragraph_set"): ("set", "flat ODT paragraph text set", "54 KB"),
}


def odf_lines(record, paths, family):
    """odf_gaps_bench / odf_verbs_bench text output (records 0288, 0291, 0292)."""
    acc: dict = {}
    for p in paths:
        for line in Path(p).read_text(encoding="utf-8").splitlines():
            m = ODF_LINE.match(line)
            if m:
                acc.setdefault((m.group(1), m.group(2)), []).append(
                    (float(m.group(3)), float(m.group(4)), float(m.group(5))))
    out = []
    for (fmt, scen), items in acc.items():
        if scen.startswith("batch_"):
            n = int(scen.split("_")[1])
            what, size = ODF_GAPS_BATCH[fmt]
            op = "set (batch of 1)" if n == 1 else "batch set"
            label = f"batch {n} ({what}), one leaf Edit vs one File::edit"
            budget = SINGLE if n == 1 else BATCH if n == 1000 else SCALING
        else:
            op, label, size = ODF_SINGLE[(fmt, scen)]
            budget = OPEN if op == "open" else SINGLE
        out.append(row(record, fmt, op, label, size, med(i[0] for i in items),
                       med(i[1] for i in items), [i[2] for i in items], budget,
                       f"{family}:{fmt}:{scen}", src(paths)))
    return out


@needs("0288")
def p0288():
    return odf_lines("0288", [RES / "0288" / "odf_gaps_bench.txt"], "odf_gaps")


def abba_builds(record, paths, labels, comparison):
    """Before/after build ABBA JSONL (records 0289, 0290, 0293): per run the
    median of each build's per-leg medians; ratio = after/before per run."""
    acc: dict = {}
    for n, p in enumerate(paths, 1):
        for r in load_jsonl(p):
            run = r.get("run", n)
            build = re.sub(r"\d+$", "", r.get("build") or r.get("leg"))  # 0289 run 2: "before2"
            for field in ("median_us", "native_median_us", "unified_median_us"):
                if field in r:
                    acc.setdefault((r["scenario"], field), {}).setdefault(run, {}).setdefault(
                        build, []).append(r[field])
    out = []
    for (scen, field), runs in acc.items():
        # A scenario one build refuses (record 0303's base) has no pair.
        if not all("before" in b and "after" in b for b in runs.values()):
            continue
        fmt, op, label = labels[scen]
        leg = {"median_us": "", "native_median_us": "native", "unified_median_us": "unified"}[field]
        before = [med(b["before"]) for _, b in sorted(runs.items())]
        after = [med(b["after"]) for _, b in sorted(runs.items())]
        out.append(row(record, fmt, op, label + (f" ({leg} leg)" if leg else ""), "",
                       med(before), med(after), [a / b for a, b in zip(after, before)],
                       regression(record), f"abba{record}:{scen}" + (f":{leg}" if leg else ""),
                       src(paths), comparison=f"after/before ({leg or comparison})"))
    return out


@needs("0289")
def p0289():
    labels = {
        "docx_edit_bold_1k": ("docx", "batch (edit)", "docx_edit_bold_1k: 1,000 run sets in one edit"),
        "doc_edit_text_200": ("doc", "batch (edit)", "doc_edit_text_200"),
        "odt_metadata_set_publish": ("odt", "metadata set", "odt_metadata_set_publish (LibreOffice ODT)"),
    }
    return abba_builds("0289", trials((RES / "0289").glob("phase-g-bench-abba-*.jsonl")),
                       labels, "unified")


@needs("0290")
def p0290():
    labels = {
        "pptx_query_shape": ("pptx", "query", "pptx_query_shape (shapes.pptx)"),
        "pptx_chart_view": ("pptx", "get", "pptx_chart_view (tdf114821.pptx)"),
        "odt_metadata_set_publish": ("odt", "metadata set", "odt_metadata_set_publish (writer-definition-lists.odt)"),
        "odp_metadata_set_publish": ("odp", "metadata set", "odp_metadata_set_publish (impress-master-layouts.odp)"),
        "odp_metadata_read": ("odp", "get", "odp_metadata_read (impress-master-layouts.odp)"),
    }
    return abba_builds("0290", [RES / "0290" / "checklist-fixes-bench-abba.jsonl"], labels, "unified")


RTF_RUN_LINE = re.compile(r"^\('(\w+)', (\d+)\) (.*)$")


def rtf_eight_runs(record, text_path, json_paths):
    """rtf-ratios-8-runs.txt (record 0291): per-run ratios of the after binary;
    the µs columns come from the committed after JSON runs."""
    times = {}
    for doc in (load_json(p)["cases"] for p in json_paths):
        for c in doc:
            times.setdefault((c["family"], c["count"]), []).append(c)
    section, out = None, []
    for line in Path(text_path).read_text(encoding="utf-8").splitlines():
        if line.startswith("=="):
            section = line.split()[1]
            continue
        m = RTF_RUN_LINE.match(line.strip())
        if section != "after" or not m:
            continue
        fam, count = m.group(1), int(m.group(2))
        parts = m.group(3).split()
        legs = dict(zip(parts[0::2], parts[1::2]))
        items = times.get((fam, count), [])
        for route, label in RTF_ROUTES:
            vals = legs.get("plan" if route == "fanout_plan" else route, "-")
            if vals == "-":
                continue
            native_key = "native_shared_median_ns" if route.startswith("fanout") else "native_median_ns"
            budget = SINGLE if count == 1 else BATCH
            op = "remove" if fam == "picture_removal" else "set"
            if count > 1:
                op = "batch " + op
            out.append(row(record, "rtf", op, f"{fam}, {label}", f"{count} targets",
                           us(med(i.get(native_key) for i in items)),
                           us(med(i.get(f"unified_{route}_median_ns") for i in items)),
                           [float(v) for v in vals.split("/")], budget,
                           f"rtf_destination:{fam}:{count}:{route}",
                           src([text_path] + list(json_paths)), comparison="unified/native"))
    return out


@needs("0291")
def p0291():
    # Profiling-profile builds (record 0291 §3), not release.
    base = RES / "0291"
    out = docx_fanout("0291", trials(base.glob("docx_fanout_bench-after-*.json")))
    out += rtf_eight_runs("0291", base / "rtf-ratios-8-runs.txt",
                          trials(base.glob("rtf_destination_bench-after-*.json")))
    out += odf_lines("0291", trials(base.glob("odf_gaps_bench-after-*.txt")), "odf_gaps")
    out += ods_rows("0291", trials(base.glob("ods_unified_bench-after-*.json")))
    out += odt_text("0291", trials(base.glob("odt_paired_bench-after-*.txt")))
    out += xls_read("0291", trials(base.glob("xls_read_bench-after-*.json")))
    return out


@needs("0292")
def p0292():
    base = RES / "0292"
    return (odf_lines("0292", trials(base.glob("odf_verbs_bench-run*.txt")), "odf_verbs")
            + odf_lines("0292", [base / "odf_gaps_bench-rerun.txt"], "odf_gaps"))


@needs("0293")
def p0293():
    path = RES / "0293" / "producer-structure-bench-abba.jsonl"
    labels = {
        "xls_complete_open_20k": ("xls", "open", "complete XLS reader open, 20k-cell workbook"),
        "xls_complete_open_formula_eval": ("xls", "open", "complete XLS reader open, FormulaEvalTestData.xls"),
        "xls_row_insert_10k": ("xls", "add (row insert)", "row insert, 10k-cell Litchi workbook"),
        "rtf_move_1k": ("rtf", "move", "paragraph move, 1k paragraphs"),
    }
    out = abba_builds("0293", [path], labels, "leaf reader")
    rows = [r for r in load_jsonl(path) if r.get("build") == "after" and "ratio" in r]
    for scen, items in group_cases([rows], lambda r: r["scenario"]).items():
        fmt, op, label = labels[scen]
        out.append(row("0293", fmt, op, label, "",
                       med(i["native_median_us"] for i in items),
                       med(i["unified_median_us"] for i in items),
                       [i["ratio"] for i in items], SINGLE, f"producer_structure:{scen}",
                       src([path]), comparison="unified/native"))
    return out


M4_FORMATS = {
    "odt_find_1": ("set (find/replace)", "find/replace, 1 match"),
    "odt_find_1000": ("set (find/replace)", "find/replace, every paragraph matches"),
    "odt_find_10000": ("set (find/replace)", "find/replace, every paragraph matches"),
    "odt_merge": ("merge", "template merge"),
    "doc_find_1": ("set (find/replace)", "find/replace, 1 match"),
    "doc_find_100": ("set (find/replace)", "find/replace, 100 matches"),
    "ods_import": ("import", "CSV import (one atomic commit)"),
}


@needs("0295")
def p0295():
    paths = trials((RES / "0295").glob("m4-formats-run*.json"))
    docs = [load_json(p)["scenarios"] for p in paths]
    out = []
    for (scen, size), items in group_cases(docs, lambda s: (s["scenario"], s["size"])).items():
        fmt = scen.split("_")[0]
        op, label = M4_FORMATS[scen]
        budget = (None, "none (record 0266 states no budget)", "none") if op == "import" else SINGLE
        out.append(row("0295", fmt, op, label,
                       f"{size} rows × 4" if fmt == "ods" else f"{size} paragraphs",
                       med(i["leaf_median_us"] for i in items),
                       med(i["unified_median_us"] for i in items),
                       [i["ratio"] for i in items], budget, f"m4_formats:{scen}:{size}", src(paths)))
    return out


def iwork_summary(record, path, size_fn, labels=None):
    """iwork_bench / pages_bench / keynote_bench summary JSONL (records 0296,
    0298, 0299, 0302): one row per run with the per-run median ratio. `path`
    is one file or a list of files (one run each); `labels` maps a scenario
    to (operation, label, budget) where the generic naming does not fit."""
    paths = path if isinstance(path, list) else [path]
    rows = [r for p in paths for r in load_jsonl(p) if r.get("summary")]
    out = []
    for (fmt, scen), items in group_cases([rows], lambda r: (r["format"], r["scenario"])).items():
        if labels and scen in labels:
            op, label, budget = labels[scen]
        elif scen.startswith("batch_"):
            n = int(scen.split("_")[1])
            op, budget = "batch set", BATCH if n in (1000, 10000) else SCALING
            label = f"batch {n} (one File::edit vs one leaf edit)"
        elif scen in ("open", "open_get"):
            op, budget = ("open" if scen == "open" else "open + get"), OPEN
            label = "open (in memory)" if scen == "open" else "open + one element's text"
        else:
            op, budget, label = scen, SINGLE, f"{scen} (one op and commit)"
        out.append(row(record, fmt, op, label, size_fn(fmt, scen),
                       us(med(i["leaf_median_ns"] for i in items)),
                       us(med(i["unified_median_ns"] for i in items)),
                       [i["ratio_median"] for i in items], budget, f"iwork:{fmt}:{scen}",
                       src(paths)))
    return out


@needs("0296")
def p0296():
    return iwork_summary("0296", RES / "0296" / "iwork-bench-summary.jsonl", lambda f, s: "")


NUMBERS_SCEN = {
    "open_get": ("open + get", "open + one cell's text", OPEN),
    "set": ("set", "set one value", SINGLE),
    "batch_1000": ("batch set", "batch 1000 (one File::edit vs one leaf edit)", BATCH),
    "formula": ("formula", "formula", SINGLE),
    "rename_sheet": ("set (sheet rename)", "rename sheet", SINGLE),
    "merge_add": ("add (merge)", "merge add", SINGLE),
    "sheet_move": ("move", "sheet move", SINGLE),
    "sheet_add": ("add (sheet)", "sheet add", SINGLE),
    "sheet_remove": ("remove (sheet)", "sheet remove", SINGLE),
    "table_add": ("add (table)", "table add", SINGLE),
    "table_remove": ("remove (table)", "table remove", SINGLE),
    "table_move": ("move (table)", "table move within its sheet", SINGLE),
    # Record 0314's routes, first parsed by record 0326.
    "batch_1000_owned": ("batch set", "batch 1000, items copied outside the clock", BATCH),
    "batch_1000_parsed": ("batch set", "batch 1000, paths parsed outside the clock", BATCH),
    # Record 0347's swaps: the leaf's moves (four steps across sheets).
    "sheet_swap": ("swap (sheet)", "sheet swap (three sheets; two leaf moves)", SINGLE),
    "table_swap": ("swap (table)", "table swap within a sheet (two leaf moves)", SINGLE),
    "table_swap_sheets": ("swap (table)",
                          "table swap across sheets (two relocations, two moves)", SINGLE),
}


def numbers_docs(record, docs, paths):
    """numbers_unified_bench scenarios (records 0297, 0301, 0302)."""
    out = []
    for name, items in group_cases(docs, lambda s: s["name"]).items():
        if name in NUMBERS_SCEN:
            op, label, budget = NUMBERS_SCEN[name]
        else:
            # Record 0309's element writes, against its bar.
            (op, label), budget = NUMBERS_0309[name], ELEMENT_0309
        fixture = Path(items[0]["fixture"]).name
        out.append(row(record, "numbers", op, label, fixture,
                       us(med(i["leaf_median_ns"] for i in items)),
                       us(med(i["unified_median_ns"] for i in items)),
                       [i["ratio"] for i in items], budget, f"numbers:{name}", src(paths)))
    return out


@needs("0297")
def p0297():
    path = RES / "0297" / "numbers-bench.jsonl"
    return numbers_docs("0297", [d["scenarios"] for d in load_jsonl(path)], [path])


NO_BUDGET_0298 = (None, "none (record 0298 §5: —)", "none")
KEYNOTE_SCEN = {
    "open": ("open", "open (in memory)", OPEN),
    "open_get": ("open + get", "open + one element's text", OPEN),
    "set": ("set", "set slide title", SINGLE),
    "notes": ("set (notes)", "set speaker notes", SINGLE),
    "frame": ("set (frame)", "set drawable frame", SINGLE),
    "batch_1000": ("batch set", "batch 1000 (one File::edit vs one leaf edit)", BATCH),
    "duplicate": ("add (duplicate slide)", "duplicate slide", NO_BUDGET_0298),
    "add_slide": ("add (slide)", "add slide", NO_BUDGET_0298),
    "add_textbox": ("add (text box)", "add text box", NO_BUDGET_0298),
    "remove_slide": ("remove (slide)", "remove slide", NO_BUDGET_0298),
}


@needs("0298")
def p0298():
    return iwork_summary("0298", RES / "0298" / "keynote-bench-summary.jsonl",
                         lambda f, s: "multi-slide.key", KEYNOTE_SCEN)


@needs("0301")
def p0301():
    path = RES / "0301" / "numbers-structure-bench.jsonl"
    return numbers_docs("0301", [d["scenarios"] for d in load_jsonl(path)], [path])


@needs("0302")
def p0302():
    """Record 0302 measures base and after builds of each harness; the base
    legs re-measure states earlier records carry, so only the after legs
    are rows (record 0294's convention)."""
    base = RES / "0302"
    out = docx_fanout("0302", trials(base.glob("docx-after-*.json")))
    odt = trials(base.glob("odt-after-*.txt"))
    docs = [d["scenarios"] for p in odt for d in load_jsonl(p)]
    for (scen, size), items in group_cases(docs, lambda s: (s["scenario"], s["size"])).items():
        op, label = M4_FORMATS[scen]
        out.append(row("0302", "odt", op, label, f"{size} paragraphs",
                       med(i["leaf_median_us"] for i in items),
                       med(i["unified_median_us"] for i in items),
                       [i["ratio"] for i in items], SINGLE, f"m4_formats:{scen}:{size}",
                       src(odt)))
    out += iwork_summary("0302", trials(base.glob("pages-after-*.jsonl")),
                         lambda f, s: "10,000-paragraph document")
    numbers = trials(base.glob("numbers-after-*.json"))
    out += numbers_docs("0302", [load_json(p)["scenarios"] for p in numbers], numbers)
    out += rtf_destination("0302", trials((base / "rtf").glob("rtf-after-*.json")))
    return out


@needs("0299")
def p0299():
    return iwork_summary("0299", RES / "0299" / "pages-bench-summary.jsonl",
                         lambda f, s: "10,000-paragraph document" if s.startswith("batch_")
                         else "basic.pages")


PPTX_GAPS = {
    "pptx_slide_remove_owned": ("pptx", "remove (slide)",
                                "remove slide 1 of 45545_Comment.pptx (notes, comments, VML, OLE)"),
    "pptx_ole_remove": ("pptx", "remove (OLE frame)", "remove the OLE frame of ole.pptx"),
    "pptx_section_move": ("pptx", "move (slide)", "move a slide of slide-sections.pptx"),
    "pptx_slide_remove_plain": ("pptx", "remove (slide)",
                                "remove dependency-free slide 6 of shapes.pptx"),
    "pptx_slide_move_plain": ("pptx", "move (slide)", "move slide 1 after slide 2 of shapes.pptx"),
    "pptx_shape_text_set": ("pptx", "set", "set one shape's text of shapes.pptx"),
}


@needs("0303")
def p0303():
    """pptx_gaps_bench ABBA JSONL: unified/native in the after build, and the
    before/after builds on the plain paths (the base refuses the others)."""
    path = RES / "0303" / "pptx-gaps-bench-abba.jsonl"
    rows = [r for r in load_jsonl(path) if r.get("build") == "after" and "ratio" in r]
    out = []
    for scen, items in group_cases([rows], lambda r: r["scenario"]).items():
        fmt, op, label = PPTX_GAPS[scen]
        out.append(row("0303", fmt, op, label, "",
                       med(i["native_median_us"] for i in items),
                       med(i["unified_median_us"] for i in items),
                       [i["ratio"] for i in items], SINGLE, f"pptx_gaps:{scen}",
                       src([path]), comparison="unified/native"))
    return out + abba_builds("0303", [path], PPTX_GAPS, "unified/native")


TEXT_SET = {
    "odt_plain": ("odt", "set (paragraph text)", "plain paragraph, 1,000 paragraphs"),
    "odt_frames": ("odt", "set (paragraph text)", "paragraph with a frame"),
    "pptx_shape": ("pptx", "set (shape text)", "shape text"),
    "pptx_paragraph": ("pptx", "set (paragraph text)", "paragraph text"),
    "pptx_cell": ("pptx", "set (table cell text)", "table cell text"),
    "ppt_shape": ("ppt", "set (shape text)", "shape text"),
    "docx_cell": ("docx", "set (table cell text)", "table cell text"),
    "odp_shape": ("odp", "set (shape text)", "shape text"),
}


@needs("0304")
def p0304():
    """text_set_bench runs (record 0304): after/before per after run against
    the median of the before runs, and unified/leaf where the run has a leaf
    leg (ODT)."""
    paths = trials((RES / "0304").glob("run*-*.json"))
    runs = {"before": [], "after": []}
    for p in paths:
        runs["after" if p.stem.endswith("after") else "before"].append(load_json(p))
    out = []
    for scen, (fmt, op, label) in TEXT_SET.items():
        before = [i["unified_us"] for doc in runs["before"] for i in doc if i["scenario"] == scen]
        after = [i for doc in runs["after"] for i in doc if i["scenario"] == scen]
        if not before or not after:
            continue
        base = med(before)
        out.append(row("0304", fmt, op, label, "", base, med(i["unified_us"] for i in after),
                       [i["unified_us"] / base for i in after], regression("0304"),
                       f"text_set:{scen}", src(paths),
                       comparison="after/before (unified leg)"))
        paired = [i for i in after if i.get("leaf_us")]
        if paired:
            out.append(row("0304", fmt, op, label, "", med(i["leaf_us"] for i in paired),
                           med(i["unified_us"] for i in paired),
                           [i["unified_us"] / i["leaf_us"] for i in paired], SINGLE,
                           f"text_set_unified:{scen}", src(paths)))
    return out


IWORK_0306 = {
    "dir_open": ("open (package directory)", "open the package directory fixture", OPEN),
    "file_open": ("open (path)", "open the single-file fixture from its path", OPEN),
    "bytes_open": ("open (bytes)", "open in-memory bytes (leaf from shared bytes)", OPEN),
    "bytes_open_options": ("open (bytes, leaf options)",
                           "open in-memory bytes with the leaf options set", OPEN),
}


@needs("0306")
def p0306():
    return iwork_summary("0306", RES / "0306" / "iwork-directory-bench.jsonl",
                         lambda f, s: {"pages": "basic.pages", "keynote": "basic.key",
                                       "numbers": "basic.numbers"}[f], IWORK_0306)


DOCX_LARGE = (("set", "auto", "text set, default (source-backed) owner"),
              ("set", "eager", "text set, eager owner"),
              ("add", "auto", "paragraph add, default (source-backed) owner"),
              ("add", "eager", "paragraph add, eager owner"),
              # Record 0311: legs on a handle already committed from the
              # default owner.
              ("set", "committed", "text set, committed handle"),
              ("add", "committed", "paragraph add, committed handle"))


def docx_large(record, paths):
    """docx_large_bench text output (record 0305 §3.6): `<leg>_<op> <µs> us`
    lines; unified legs over the leaf leg of the same operation."""
    runs = []
    for p in paths:
        values = {}
        for line in Path(p).read_text(encoding="utf-8").splitlines():
            parts = line.split()
            if len(parts) == 3 and parts[2] == "us":
                values[parts[0]] = float(parts[1])
        runs.append(values)
    out = []
    for op, owner, label in DOCX_LARGE:
        leaf, unified = f"leaf_{op}", f"{owner}_{op}"
        items = [r for r in runs if leaf in r and unified in r]
        if not items:
            continue
        out.append(row(record, "docx", op, label, "1,000 paragraphs",
                       med(r[leaf] for r in items), med(r[unified] for r in items),
                       [r[unified] / r[leaf] for r in items], SINGLE,
                       f"docx_large:{op}:{owner}", src(paths)))
    return out


@needs("0305")
def p0305():
    """Record 0305: after legs of its keynote, numbers, docx, rtf and pages
    harnesses and of docx_large_bench (the base legs re-measure states
    earlier records carry)."""
    base = RES / "0305"
    out = iwork_summary("0305", trials(base.glob("keynote-after-*.jsonl")),
                        lambda f, s: "multi-slide.key", KEYNOTE_SCEN)
    numbers = trials(base.glob("numbers-after-*.json"))
    out += numbers_docs("0305", [load_json(p)["scenarios"] for p in numbers], numbers)
    out += docx_fanout("0305", trials(base.glob("docx-after-*.json")))
    out += rtf_destination("0305", trials(base.glob("rtf-after-*.json")))
    out += iwork_summary("0305", trials(base.glob("pages-after-*.jsonl")),
                         lambda f, s: "10,000-paragraph document")
    out += docx_large("0305", trials(base.glob("docx-large-after-*.txt")))
    return out


ELEMENT_0309 = (1.05, "1.05 or +20 µs (record 0309 §6)", "open")
IWORK_0309 = {
    "run_style": ("set (run style)", "run character style", ELEMENT_0309),
    "placeholder_hide": ("set (placeholder)", "hide a placeholder", ELEMENT_0309),
    "media_remove": ("remove (media)", "remove a movie/audio", ELEMENT_0309),
    "table_cell": ("set (table cell)", "table cell value", ELEMENT_0309),
    "paragraph_font": ("set (paragraph font)", "paragraph font", ELEMENT_0309),
}
NUMBERS_0309 = {
    "row_insert": ("add (row)", "insert a row"),
    "row_remove": ("remove (row)", "remove a row"),
    "col_insert": ("add (column)", "insert a column"),
    "col_remove": ("remove (column)", "remove a column"),
    "cell_font": ("set (cell font)", "cell font"),
}


@needs("0309")
def p0309():
    """Record 0309 §6: the element-write scenarios of numbers_unified_bench,
    keynote_bench and pages_bench, against its 5 % or 20 µs bar."""
    base = RES / "0309"
    out = iwork_summary("0309", trials(base.glob("keynote-*.jsonl")),
                        lambda f, s: "multi-slide.key", IWORK_0309)
    out += iwork_summary("0309", trials(base.glob("pages-*.jsonl")),
                         lambda f, s: "basic.pages", IWORK_0309)
    paths = trials(base.glob("numbers-*.json"))
    docs = [load_json(p)["scenarios"] for p in paths]
    for name, items in group_cases(docs, lambda s: s["name"]).items():
        op, label = NUMBERS_0309[name]
        out.append(row("0309", "numbers", op, label, Path(items[0]["fixture"]).name,
                       us(med(i["leaf_median_ns"] for i in items)),
                       us(med(i["unified_median_ns"] for i in items)),
                       [i["ratio"] for i in items], ELEMENT_0309, f"numbers:{name}",
                       src(paths)))
    return out


XLS_P0_OPS = {"sheet_add": "add", "sheet_move": "move"}


def xls_p0_rows(record, paths):
    """xls_p0_write_bench runs (records 0259, 0310): unified/native per case,
    the median of each run's per-case medians."""
    docs = [load_json(p)["cases"] for p in paths]
    out = []
    for (scen, owner, cells), items in group_cases(
            docs, lambda c: (c["scenario"], c["owner"], c["cells"])).items():
        out.append(row(record, "xls", XLS_P0_OPS.get(scen, "set"), f"{scen} ({owner} owner)",
                       f"{cells} cells", us(med(i["median_native_ns"] for i in items)),
                       us(med(i["median_unified_ns"] for i in items)),
                       [i["ratio"] for i in items], SINGLE,
                       f"xls_p0_write:{scen}:{owner}:{cells}", src(paths),
                       comparison="unified/native"))
    return out


RTF_TINY_KEYS = {
    "unified-rtf-add": ("paragraphs",),
    "unified-rtf-layout": ("paragraphs", "route"),
    "unified-rtf-run-add": ("shape", "groups", "route"),
    "unified-rtf-batch-set": ("family", "shape", "count"),
}


def rtf_tiny_runs(record, harness, paths):
    """Per-run JSON of the record 0275 tiny-input harnesses (record 0310 re-ran
    them): the same rows and series as record 0275's summary, one ratio per
    run."""
    docs = [load_json(p)["cases"] for p in paths]
    keys = RTF_TINY_KEYS[harness]
    out = []
    for case, items in group_cases(
            docs, lambda c: ",".join(f"{k}={c[k]}" for k in keys)).items():
        kv = dict(part.split("=") for part in case.split(","))
        if harness == "unified-rtf-add":
            op, scen, size, budget = "add", "add paragraph", f"{kv['paragraphs']} paragraphs", SINGLE
        elif harness == "unified-rtf-run-add":
            op, scen = "add", f"add run ({kv['shape']}, {kv['route']})"
            size, budget = f"{kv['groups']} groups", SINGLE
        elif harness == "unified-rtf-layout":
            n = int(kv["paragraphs"])
            if kv["route"] == "last_path":
                op, scen, budget = "set", "paragraph layout set, path", SINGLE
            elif kv["route"] == "last_path_typed":
                # Record 0315's typed-props route: the lengths as typed values.
                op, scen, budget = "set", "paragraph layout set, path, typed lengths", SINGLE
            else:
                op, scen = "set (query fan-out)", "paragraph layout set, query over all paragraphs"
                budget = SINGLE if n == 1 else BATCH
            size = f"{n} paragraphs"
        else:
            n = int(kv["count"])
            op = "set (edit closure)" if n == 1 else "batch set (edit closure)"
            scen = f"{kv['family']} ({kv['shape']})"
            size, budget = f"{n} targets", SINGLE if n == 1 else BATCH
        out.append(row(record, "rtf", op, scen, size,
                       us(med(i["native_median_ns"] for i in items)),
                       us(med(i["unified_median_ns"] for i in items)),
                       [i["ratio"] for i in items], budget, f"rtf_tiny:{harness}:{case}",
                       src(paths), comparison="unified/native"))
    return out


def text_set_against(record, paths, before_paths, scenarios):
    """text_set_bench runs of a later build against the before runs of record
    0304 (record 0310): the 0304 series, so the later reading supersedes it."""
    before_docs = [load_json(p) for p in before_paths]
    docs = [load_json(p) for p in paths]
    out = []
    for scen in scenarios:
        fmt, op, label = TEXT_SET[scen]
        base = med(i["unified_us"] for d in before_docs for i in d if i["scenario"] == scen)
        after = [i for d in docs for i in d if i["scenario"] == scen]
        out.append(row(record, fmt, op, label, "", base, med(i["unified_us"] for i in after),
                       [i["unified_us"] / base for i in after], regression("0304"),
                       f"text_set:{scen}", src(paths),
                       comparison="trunk/0304 before (unified leg)"))
        paired = [i for i in after if i.get("leaf_us")]
        if paired:
            out.append(row(record, fmt, op, label, "", med(i["leaf_us"] for i in paired),
                           med(i["unified_us"] for i in paired),
                           [i["unified_us"] / i["leaf_us"] for i in paired], SINGLE,
                           f"text_set_unified:{scen}", src(paths)))
    return out


OFFICE_GAPS_0311 = {
    "doc_text_set": ("doc", "set", "paragraph text set (against the body-text snapshot leaf)"),
    "doc_package_text_set": ("doc", "set", "paragraph text set (against the Package::edit leaf)"),
}


@needs("0311")
def p0311():
    """Record 0311: after legs of docx_large_bench (with its committed-handle
    legs) and of the office_gaps_bench DOC scenarios."""
    base = RES / "0311"
    out = docx_large("0311", trials(base.glob("docx-large-after-*.txt")))
    paths = trials(base.glob("doc-after-*.jsonl"))
    for scen, items in group_cases([load_jsonl(p) for p in paths], lambda r: r["scenario"]).items():
        fmt, op, label = OFFICE_GAPS_0311[scen]
        out.append(row("0311", fmt, op, label, "",
                       med(i["native_median_us"] for i in items),
                       med(i["unified_median_us"] for i in items),
                       [i["ratio"] for i in items], SINGLE, f"office_gaps:{scen}",
                       src(paths), comparison="unified/native"))
    return out


@needs("0312")
def p0312():
    """Record 0312: its keynote_bench, pages_bench and numbers_unified_bench
    runs (the 0305/0309 scenarios after the lazy iWork tables)."""
    base = RES / "0312"
    out = iwork_summary("0312", trials(base.glob("keynote-*.jsonl")),
                        lambda f, s: "multi-slide.key", {**KEYNOTE_SCEN, **IWORK_0309})
    out += iwork_summary("0312", trials(base.glob("pages-*.jsonl")),
                         lambda f, s: "10,000-paragraph document" if s.startswith("batch_")
                         else "basic.pages", IWORK_0309)
    numbers = trials(base.glob("numbers-*.json"))
    out += numbers_docs("0312", [load_json(p)["scenarios"] for p in numbers], numbers)
    return out


@needs("0313")
def p0313():
    """Record 0313: after legs of docx_large_bench (DOCX insertions with
    successor proofs)."""
    return docx_large("0313", trials((RES / "0313").glob("docx-large-after-*.txt")))


@needs("0310")
def p0310():
    """Record 0310: `merged/` re-measures every harness of an open miss on
    the merged tree (trunk 8f5569d0d8 with this record's changes). The
    record's `trunk/` runs (its base 0f3d8158cb) and `abba/` runs (base
    against its pre-merge changes, B A A B) are evidence for its prose and
    are not rows: `merged/` measures the same series later."""
    base = RES / "0310" / "merged"
    out = query_routes("0310", trials(base.glob("query-[0-9].jsonl")))
    out += query_routes("0310", trials(base.glob("query-all-[0-9].jsonl")), build="all-formats")
    out += rtf_destination("0310", trials(base.glob("rtf_destination-*.json")))
    out += xls_read("0310", trials(base.glob("xls_read-*.json")))
    out += ods_rows("0310", trials(base.glob("ods-*.json")))
    out += odt_text("0310", trials(base.glob("odt-*.txt")))
    out += docx_fanout("0310", trials(base.glob("docx_fanout-*.json")))
    out += docx_element("0310", trials(base.glob("docx_element-*.json")))
    out += xls_p0_rows("0310", trials(base.glob("xls_p0-*.json")))
    for harness in RTF_TINY_KEYS:
        out += rtf_tiny_runs("0310", harness, trials(base.glob(f"{harness}-[0-9].json")))
    out += text_set_against("0310", trials(base.glob("text_set-*.json")),
                            trials((RES / "0304").glob("run*-before.json")), ["odt_plain"])
    out += iwork_summary("0310", trials(base.glob("iwork-directory-*.jsonl")),
                         lambda f, s: {"pages": "basic.pages", "keynote": "basic.key",
                                       "numbers": "basic.numbers"}[f], IWORK_0306)
    out += iwork_summary("0310", trials(base.glob("keynote-*.jsonl")),
                         lambda f, s: "multi-slide.key", {**KEYNOTE_SCEN, **IWORK_0309})
    numbers = trials(base.glob("numbers-*.json"))
    out += numbers_docs("0310", [load_json(p)["scenarios"] for p in numbers], numbers)
    return out


@needs("0315")
def p0315():
    """Record 0315: the after legs of its B A A B run (`abba/`, base 30eed270b5
    with the record's new routes against the record's tree): the query
    harness in both builds, xls_read_bench, the ODS and ODT queries, and the
    RTF single-target harnesses (one target or one paragraph; no RTF code
    changed, so only the single-target rows of G9 are taken). The base legs
    are evidence for the record's before/after tables, not rows."""
    base = RES / "0315" / "abba"
    out = query_routes("0315", trials(base.glob("query-after-[0-9].jsonl")))
    out += query_routes("0315", trials(base.glob("query-all-after-[0-9].jsonl")),
                        build="all-formats")
    out += xls_read("0315", trials(base.glob("xls_read-after-*.json")))
    out += ods_rows("0315", trials(base.glob("ods-after-*.json")))
    out += [r for r in odt_text("0315", trials(base.glob("odt-after-*.txt")))
            if r["series"].startswith("odt_paired:query_all")]
    out += [r for r in rtf_destination("0315", trials(base.glob("rtf_destination-after-*.json")))
            if r["size"] == "1 targets"]
    for harness in ("unified-rtf-layout", "unified-rtf-batch-set"):
        out += [r for r in rtf_tiny_runs("0315", harness,
                                         trials(base.glob(f"{harness}-after-[0-9].json")))
                if r["size"] in ("1 paragraphs", "1 targets")]
    return out


READ_0326 = (1.10, "1.10 (point read; records 0319-0322 counted a read above 10 % as a miss)", "read")
IWORK_0326 = {
    "batch_1000_parsed": ("batch set", "batch 1000, paths parsed outside the clock", BATCH),
    "batch_10000_parsed": ("batch set", "batch 10000, paths parsed outside the clock", SCALING),
}


def paired_reads_0326(record, fmt, harness, paths, cases):
    """Record 0326's element-read legs: `cases` maps a case name to its
    scenario label; each file holds one JSON line (xlsx/xlsb feature
    benches) or document (docx_parts_bench, unified-pptx-elements) per run."""
    runs = {}
    for p in paths:
        text = p.read_text(encoding="utf-8").strip()
        if not text:
            continue
        docs = []
        if text.startswith("{\n") or text.startswith("{\r"):
            docs = [json.loads(text)]
        else:
            docs = [json.loads(line) for line in text.splitlines() if line.strip()]
        for d in docs:
            for c in d.get("cases") or d.get("results") or [d]:
                name = c.get("case") or c.get("scenario") or c.get("route")
                if name not in cases or c.get("skipped"):
                    continue
                if "leaf_median_ns" in c:
                    leaf, unified = c["leaf_median_ns"] / 1000, c["unified_median_ns"] / 1000
                elif "native_median_ns" in c:
                    leaf, unified = c["native_median_ns"] / 1000, c["unified_median_ns"] / 1000
                else:
                    leaf, unified = c["leaf_median_us"], c["unified_median_us"]
                runs.setdefault(name, []).append((leaf, unified))
    out = []
    for name, label in cases.items():
        items = runs.get(name, [])
        if not items:
            continue
        out.append(row(record, fmt, "get", label, "",
                       med(i[0] for i in items), med(i[1] for i in items),
                       [i[1] / i[0] for i in items], READ_0326, f"{harness}:{name}",
                       src(paths)))
    return out


@needs("0326")
def p0326():
    """Record 0326: the after legs of its B A A B run (`abba/`, base = trunk
    16b800350b with the record's new path_parse_bench example, after = the
    merged tree): the PPTX query routes, the iWork batch legs of record
    0314 (text paths and `_parsed`), the RTF single-target typed and text
    routes, and the element reads of records 0319-0322 it re-measures."""
    base = RES / "0326" / "abba"
    out = query_routes("0326", trials(base.glob("query-after-[0-9].jsonl")))
    out += iwork_summary("0326", trials(base.glob("pages-after-*.jsonl")),
                         lambda f, s: "10,000-paragraph document", IWORK_0326)
    out += iwork_summary("0326", trials(base.glob("keynote-after-*.jsonl")),
                         lambda f, s: "multi-slide.key", {**KEYNOTE_SCEN, **IWORK_0326})
    numbers = trials(base.glob("numbers-after-*.json"))
    out += numbers_docs("0326", [load_json(p)["scenarios"] for p in numbers], numbers)
    for harness in ("unified-rtf-layout", "unified-rtf-batch-set"):
        out += [r for r in rtf_tiny_runs("0326", harness,
                                         trials(base.glob(f"{harness}-after-[0-9].json")))
                if r["size"] in ("1 paragraphs", "1 targets")]
    out += paired_reads_0326("0326", "docx", "docx_parts",
                             trials(base.glob("docx-*-after-*.json")),
                             {"style_get_format": "style get + format (/styles/Heading1)",
                              "paragraph_list_format": "list paragraph get + format (/body/p[3])"})
    out += paired_reads_0326("0326", "xlsx", "xlsx_features",
                             trials(base.glob("xlsx-after-*.jsonl")),
                             {"comment_get": "note get + first format entry (comment[20] of 40)"})
    out += paired_reads_0326("0326", "xlsb", "xlsb_features",
                             trials(base.glob("xlsb-after-*.jsonl")),
                             {"comment_get": "note get + first format entry (comment[20] of 40)"})
    out += paired_reads_0326("0326", "pptx", "pptx_elements",
                             trials(base.glob("pptx-elements-after-*.json")),
                             {"pptx_equation_get_v1": "equation get + format (/slide[1]/shape[1]/equation[1])"})
    return out


BATCH_0337_SCENARIOS = {
    "run_bold_set": "run bold sets on distinct paragraphs",
    "cell_value_set": "cell value sets on distinct cells",
    "cell_remove": "cell removals on distinct cells",
    "shape_text_set": "shape text sets on distinct shapes",
    "paragraph_text_set": "paragraph text sets on distinct paragraphs",
    "series_label_set": "series label sets on distinct series",
    "query_command_set": "query command sets on distinct queries",
    "token_text_set": "token text sets on distinct tokens",
    "frame_property_set": "every independent frame property",
}
STAGES_0337 = {"trunk": ("", 1), "after": ("after record 0337", 2)}


def batch_closure(record, paths, stages=STAGES_0337):
    """Record 0337's `unified-batch-closure` harness: one JSON document per
    process (trial); a case is one (format, owner, scenario, operations). The
    unified leg is one atomic `File::batch` of N items, the leaf leg one
    hand-written leaf edit (split only where the leaf's own bound forces
    it); every case was checked equal (exact bytes or decoded members)
    before timing. Ratios are per trial; leaf/unified are medians of the
    trials' medians."""
    groups = {}
    for p in paths:
        m = re.search(r"batch-[A-Z]-([a-z]+)-\d+(?:-n\d+)?\.json$", p.name)
        stage, rank = stages["trunk" if m.group(1) in ("trunk", "before") else m.group(1)]
        for c in load_json(p)["cases"]:
            k = (c["format"], c["owner"], c["scenario"], c["operations"], stage, rank,
                 c["fixture_targets"])
            groups.setdefault(k, []).append((c, p))
    out = []
    for (fmt, owner, scen, n, stage, rank, fixture), items in sorted(groups.items()):
        budget = BATCH if n >= 1000 or fmt == "odi" else SCALING
        cases = [c for c, _ in items]
        held = "one frame" if fmt == "odi" else f"fixture of {fixture:,}"
        out.append(row(record, fmt, "batch", f"{BATCH_0337_SCENARIOS[scen]} ({owner} owner; "
                       f"one File::batch vs one leaf edit, {held})", f"{n} ops",
                       us(med(c["leaf_median_ns"] for c in cases)),
                       us(med(c["unified_median_ns"] for c in cases)),
                       [c["ratio"] for c in cases], budget,
                       f"batch_closure:{fmt}:{owner}:{scen}:{n}:fixture{fixture}", src(sorted({p for _, p in items})),
                       stage=stage, stage_rank=rank))
    return out


@needs("0337")
def p0337():
    """Record 0337: 1k-batch evidence for every format (`abba/batch-*.json`)."""
    return batch_closure("0337", sorted((RES / "0337" / "abba").glob("batch-*.json")))


@needs("0340")
def p0340():
    """Record 0340: the iWork and ODC batches re-measured after the Keynote
    paragraph-run and object-navigator fixes (`abba/batch-A-after-*.json`;
    the `before` runs of the same session repeat record 0337's trunk and are
    not listed)."""
    return batch_closure("0340", sorted((RES / "0340" / "abba").glob("batch-*-after-*.json")),
                         stages={"after": ("", 1)})


# ------------------------------------------------- record 0344's parsers
#
# Records 0307-0342 whose evidence the matrix did not read (second
# definition-of-done audit §5). Each reads the after legs of its record's
# own harness; "before" legs of a before/after-build comparison give the
# comparison's denominator only.

def slug(text):
    return re.sub(r"[^a-z0-9]+", "_", text.lower()).strip("_")


@needs("0307")
def p0307():
    """Record 0307 §6 (`docx_mc_bench`, twelve runs alternated before/after):
    unified/leaf in the after build, and after/before of each leg (medians
    of the six runs of each binary, as the record compares them). The
    1,000-paragraph text set and add on the default owner are the series of
    `docx_large_bench` (records 0305, 0311, 0313): the same scenario on the
    same owner, measured again after the 0305/0311 fixes."""
    base = RES / "0307"
    paths = trials(base.glob("run*-*.json"))
    runs = {"before": [], "after": []}
    for p in paths:
        runs["after" if p.stem.endswith("after") else "before"].append(load_json(p))
    out = []
    large = {"text_set": ("set", "auto", "text set, default (source-backed) owner"),
             "add_paragraph": ("add", "auto", "paragraph add, default (source-backed) owner")}
    for scen, (op, owner, label) in large.items():
        after = [i for d in runs["after"] for i in d if i["scenario"] == scen]
        out.append(row("0307", "docx", op, label, "1,000 paragraphs",
                       med(i["leaf_us"] for i in after), med(i["unified_us"] for i in after),
                       [i["unified_us"] / i["leaf_us"] for i in after], SINGLE,
                       f"docx_large:{op}:{owner}", src(paths)))
    labels = {"text_set": ("set", "text set, middle paragraph"),
              "add_paragraph": ("add", "paragraph add, middle"),
              "run_bold": ("set", "run bold set (compacts the paragraph)"),
              "find": ("find/replace", "find/replace over the body")}
    for scen, (op, label) in labels.items():
        for leg in ("leaf", "unified"):
            before = [i[f"{leg}_us"] for d in runs["before"] for i in d
                      if i["scenario"] == scen and i.get(f"{leg}_us") is not None]
            after = [i[f"{leg}_us"] for d in runs["after"] for i in d
                     if i["scenario"] == scen and i.get(f"{leg}_us") is not None]
            if not before or not after:
                continue
            b, a = med(before), med(after)
            out.append(row("0307", "docx", op, f"{label} ({leg} leg; docx_mc_bench)",
                           "1,000 paragraphs", b, a, [a / b],
                           regression("0307"), f"docx_mc:{scen}:{leg}", src(paths),
                           comparison=f"after/before ({leg} leg)"))
    return out


# office_gaps_bench (record 0308) and xls_odp_followups_bench (record 0316):
# scenario -> (format, operation, label, budget of the unified/native pair).
GAPS_SCEN = {
    "xls_complete_open_20k": ("xls", "open", "complete open, 20k cells"),
    "xls_complete_open_formula_eval": ("xls", "open", "complete open, FormulaEvalTestData"),
    "xls_source_visit_20k": ("xls", "query", "source-backed walk, 20k cells"),
    "xls_row_insert_10k": ("xls", "add (row)", "row insert, 10k cells"),
    "xls_row_insert_images": ("xls", "add (row)", "row insert, SimpleWithImages.xls"),
    "xls_row_insert_unindexed": ("xls", "add (row)", "row insert, sheet without INDEX"),
    "xls_row_insert_owners_10k": ("xls", "add (row)", "row insert, 10k cells + array formula + data table"),
    "xls_row_insert_conditional_formats": ("xls", "add (row)", "row insert, WithConditionalFormatting.xls (CF, CFEx)"),
    "xls_row_insert_validations": ("xls", "add (row)", "row insert, validation.xls (DV PtgRefN)"),
    "xls_row_insert_cf12": ("xls", "add (row)", "row insert, NewStyleConditionalFormattings.xls (CF12)"),
    "xls_row_insert_locked_anchors": ("xls", "add (row)", "row insert, SimpleWithComments.xls (position-locked)"),
    "rtf_move_plain_1k": ("rtf", "move", "move, 1k plain paragraphs"),
    "rtf_move_producer_1k": ("rtf", "move", "move, 1k producer paragraphs (exact splice)"),
    "rtf_producer_insert_1k": ("rtf", "add", "producer insert, 1k paragraphs"),
    "rtf_producer_replay_1k": ("rtf", "move", "value-exact move (replayed state), 1k paragraphs"),
    "odp_scene_parse_200": ("odp", "open (leaf Scene::parse)", "Scene::parse, 200 slides, 400 styles"),
    "odp_run_read_200": ("odp", "get", "bold of 200 runs (unified, resolved)"),
    "odp_run_format_200": ("odp", "set", "bold + size on one run, 200 slides"),
    "odp_shape_format_200": ("odp", "set", "bold on a whole shape, 200 slides"),
    "odp_shape_read_200": ("odp", "query", "open + shared bold of 200 shapes"),
    "doc_text_set": ("doc", "set", "paragraph text set (against the body-text snapshot leaf)"),
}


def gaps_runs(record, harness, before, after, new=(), stage="", stage_rank=1,
              only=None, pair_budget=None):
    """office_gaps-style JSONL runs: one line per scenario, either one leg
    (`median_us`) or a native/unified pair. `before`/`after` are run files of
    the base and the record's build: after/before rows per leg, the median
    of the after runs over the median of the before runs (the records'
    own comparison; one noisy run of a shared host does not decide it). Pairs give unified/native rows from the after
    runs, and `new` files (after-only scenarios) unified/native rows."""
    def lines(paths):
        return [r for p in paths for r in load_jsonl(p) if only is None or r["scenario"] in only]
    b_rows, a_rows = lines(before), lines(after)
    out = []
    srcs = src(list(before) + list(after))
    for scen, a_items in group_cases([a_rows], lambda r: r["scenario"]).items():
        fmt, op, label = GAPS_SCEN[scen]
        b_items = [r for r in b_rows if r["scenario"] == scen]
        legs = (("", "median_us"),) if "median_us" in a_items[0] else \
            (("leaf", "native_median_us"), ("unified", "unified_median_us"))
        for leg, field in legs:
            base = med(r[field] for r in b_items)
            if base is None:
                continue
            after = med(r[field] for r in a_items)
            out.append(row(record, fmt, op, label + (f" ({leg} leg)" if leg else ""), "",
                           base, after, [after / base], regression(record),
                           f"{harness}:{scen}" + (f":{leg}" if leg else ""), srcs,
                           stage, stage_rank,
                           comparison="after/before" + (f" ({leg} leg)" if leg else "")))
        if "ratio" in a_items[0]:
            out.append(row(record, fmt, op, label, "",
                           med(r["native_median_us"] for r in a_items),
                           med(r["unified_median_us"] for r in a_items),
                           [r["ratio"] for r in a_items], (pair_budget or {}).get(scen, SINGLE),
                           f"office_gaps:{scen}" if scen == "doc_text_set" else f"{harness}:{scen}:pair",
                           srcs, stage, stage_rank, comparison="unified/native"))
    n_rows = lines(new)
    for scen, items in group_cases([n_rows], lambda r: r["scenario"]).items():
        fmt, op, label = GAPS_SCEN[scen]
        out.append(row(record, fmt, op, label, "",
                       med(r["native_median_us"] for r in items),
                       med(r["unified_median_us"] for r in items),
                       [r["ratio"] for r in items], (pair_budget or {}).get(scen, SINGLE),
                       f"{harness}:{scen}:pair", src(new), comparison="unified/native"))
    return out


@needs("0308")
def p0308():
    """Record 0308 §7: the final `r2-*` runs (base trunk 6c8d517eff against
    the record's second build), the new-path harness `r2-new*`, and the ODP
    scene re-measured on the third build (`odp-*`, a later stage)."""
    base = RES / "0308"
    out = gaps_runs("0308", "office_gaps", trials(base.glob("r2-run*-before.jsonl")),
                    trials(base.glob("r2-run*-after2.jsonl")), trials(base.glob("r2-new*.jsonl")))
    out += gaps_runs("0308", "office_gaps", trials(base.glob("odp-run*-before.jsonl")),
                     trials(base.glob("odp-run*-after3.jsonl")),
                     stage="after3 (scan check limited to depth ≤ 2)", stage_rank=2,
                     only={"odp_scene_parse_200"})
    return out


IWORK_0344 = {
    "table_add": ("add (table)", "table add", SINGLE),
    "notes_run_style": ("set (notes run style)", "speaker-notes run character style", SINGLE),
}
NUMBERS_SCEN.update({
    "row_insert_formulas": ("add (row)", "insert a row above formulas", SINGLE),
    "comment_set": ("set (comment)", "set a cell comment", SINGLE),
    "row_remove_ref": ("remove (row)", "remove a row that formulas reference (#REF!)", SINGLE),
    "comment_empty": ("set (comment)", "comment edge case: empty text", SINGLE),
})


def iwork_after(record, base, glob_fmt):
    """keynote_bench, pages_bench and numbers_unified_bench after legs
    (records 0314, 0318, 0329); `glob_fmt` names a format's after files."""
    out = []
    keynote = trials(base.glob(glob_fmt.format(fmt="keynote", ext="jsonl")))
    if keynote:
        out += iwork_summary(record, keynote, lambda f, s: "multi-slide.key",
                             {**KEYNOTE_SCEN, **IWORK_0309, **IWORK_0326, **IWORK_0344})
    pages = trials(base.glob(glob_fmt.format(fmt="pages", ext="jsonl")))
    if pages:
        out += iwork_summary(record, pages,
                             lambda f, s: "10,000-paragraph document" if s.startswith("batch_")
                             else "basic.pages", {**IWORK_0309, **IWORK_0326, **IWORK_0344})
    numbers = trials(base.glob(glob_fmt.format(fmt="numbers", ext="json")))
    if numbers:
        out += numbers_docs(record, [load_json(p)["scenarios"] for p in numbers], numbers)
    return out


@needs("0314")
def p0314():
    """Record 0314 §3: the after legs (`out/*-after-*`) of its B A A B run."""
    return iwork_after("0314", RES / "0314" / "out", "{fmt}-after-*.{ext}")


@needs("0316")
def p0316():
    """Record 0316 §6: the final `r3-*` runs (base 30eed270b5 against the
    merged tree) and the new-path harness `r3-new*`."""
    base = RES / "0316"
    return gaps_runs("0316", "xls_odp_followups", trials(base.glob("r3-run*-before.jsonl")),
                     trials(base.glob("r3-run*-after.jsonl")), trials(base.glob("r3-new*.jsonl")),
                     pair_budget={"odp_shape_read_200": QUERY_TRIGGER})


PPTX_LARGE = (("set", "unified_set_write", "shape text set + write"),
              ("add", "unified_add_write", "shape add + write"),
              ("set", "committed_set_write", "shape text set + write, committed handle"),
              ("add", "committed_add_write", "shape add + write, committed handle"))


@needs("0317")
def p0317():
    """Record 0317 §4: the after legs of docx_large_bench and of the new
    pptx_large_bench (1,000-shape slide; µs lines like docx_large_bench)."""
    base = RES / "0317"
    out = docx_large("0317", trials(base.glob("docx-large-after-*.txt")))
    paths = trials(base.glob("pptx-large-after-*.txt"))
    runs = []
    for p in paths:
        values = {}
        for line in Path(p).read_text(encoding="utf-8").splitlines():
            parts = line.split()
            if len(parts) == 3 and parts[2] == "us":
                values[parts[0]] = float(parts[1])
        runs.append(values)
    for op, unified, label in PPTX_LARGE:
        leaf = f"leaf_{op}_write"
        items = [r for r in runs if leaf in r and unified in r]
        out.append(row("0317", "pptx", op, label, "1,000 shapes",
                       med(r[leaf] for r in items), med(r[unified] for r in items),
                       [r[unified] / r[leaf] for r in items], SINGLE,
                       f"pptx_large:{unified}", src(paths)))
    return out


@needs("0318")
def p0318():
    """Record 0318 §4: numbers_unified_bench after legs."""
    return iwork_after("0318", RES / "0318", "{fmt}-*-after.{ext}")


READ_POINT = (1.10, "1.10 (point read; records 0319-0328 count a read above 10 % as a miss)", "read")
PPT_ELEMENTS = {
    "ppt_comment_set_v1": ("set (comment)", "comment text set"),
    "ppt_comment_add_v1": ("add (comment)", "comment add"),
    "ppt_cell_text_set_v1": ("set (table cell text)", "table cell text set"),
    "ppt_hyperlink_add_v1": ("add (hyperlink)", "hyperlink add"),
    "ppt_run_format_v1": ("set (run)", "run format set"),
    "ppt_comment_get_v1": ("get", "comment get + format"),
    "ppt_run_get_v1": ("get", "run get + format"),
}


@needs("0323")
def p0323():
    """Record 0323 §5: ppt_elements_probe, two runs."""
    paths = trials((RES / "0323").glob("ppt-elements*.json"))
    out = []
    for (scen, n), items in group_cases([load_json(p) for p in paths],
                                        lambda r: (r["scenario"], r["slides"])).items():
        op, label = PPT_ELEMENTS[scen]
        out.append(row("0323", "ppt", op, label, f"{n} slides",
                       med(i["native_median_us"] for i in items),
                       med(i["unified_median_us"] for i in items),
                       [i["pair_ratio_median"] for i in items],
                       READ_POINT if op == "get" else SINGLE, f"ppt_elements:{scen}:{n}",
                       src(paths), comparison="unified/native"))
    return out


def parts_trials(record, harness, paths, fmt_of, ops):
    """`*_parts_bench` trials (records 0324, 0328, 0335): `cases` with leaf
    and unified medians and the ratio; reads carry the point-read budget."""
    docs = [load_json(p)["cases"] for p in paths]
    out = []
    for case, items in group_cases(docs, lambda c: c["case"]).items():
        op, budget = ops(case)
        out.append(row(record, fmt_of(case), op, case, "",
                       us(med(i["leaf_median_ns"] for i in items)),
                       us(med(i["unified_median_ns"] for i in items)),
                       [i["ratio"] for i in items], budget, f"{harness}:{case}", src(paths)))
    return out


def part_op(case):
    if case.endswith(("_get_format", "_list_format")):
        return "get", READ_POINT
    if "_add" in case:
        return "add", SINGLE
    if "_remove" in case:
        return "remove", SINGLE
    return "set", SINGLE


def view_sections(path, header, after_label):
    """`== <name> <label>` sections of `<key>\\t<ns>` lines (records 0324,
    0325, 0335): per key, the after and before values."""
    acc: dict = {}
    current = None
    for line in Path(path).read_text(encoding="utf-8").splitlines():
        line = line.rstrip("\n")
        m = header.match(line)
        if m:
            current = "after" if m.group("label") == after_label else "before"
            continue
        if current and "\t" in line:
            k, v = line.rsplit("\t", 1)
            acc.setdefault(k, {"before": [], "after": []})[current].append(float(v))
    return acc


def view_rows(record, fmt_of, harness, acc, source, labels=None):
    out = []
    for k, legs in acc.items():
        b, a = med(legs["before"]), med(legs["after"])
        out.append(row(record, fmt_of(k), "get (view)", (labels or {}).get(k, k), "", b, a,
                       [a / b], regression(record), f"{harness}:{k}", source,
                       comparison="after/before", unit="ns"))
    return out


@needs("0324")
def p0324():
    """Record 0324 §6: doc_parts_bench trials and doc_view_regression (two
    runs, B = before)."""
    base = RES / "0324"
    out = parts_trials("0324", "doc_parts", trials(base.glob("parts-trial-*.json")),
                       lambda c: "doc", part_op)
    path = base / "view-regression.txt"
    header = re.compile(r"^== round(?P<round>\d+) view-\d+-(?P<label>[AB])\.tsv$")
    rounds: dict = {}
    current = None
    for line in Path(path).read_text(encoding="utf-8").splitlines():
        line = line.rstrip("\n")
        m = header.match(line)
        if m:
            current = (m.group("round"), "after" if m.group("label") == "A" else "before")
            continue
        if current and "\t" in line:
            k, v = line.split("\t")
            rounds.setdefault(k, {}).setdefault(current, []).append(float(v))
    for k, legs in rounds.items():
        ratios = [med(legs[(r, "after")]) / med(legs[(r, "before")]) for r in ("1", "2")]
        b = med(legs[("1", "before")] + legs[("2", "before")])
        a = med(legs[("1", "after")] + legs[("2", "after")])
        out.append(row("0324", "doc", "set (view)" if k.endswith("_set") else "get (view)",
                       k, "", b, a, ratios, regression("0324"), f"doc_view:{k}",
                       src([path]), comparison="after/before", unit="ns"))
    return out


OOXML_0325 = {
    "story_paragraph_set": ("docx", "set"), "story_text_set": ("docx", "set"),
    "bookmark_rename": ("docx", "set"), "field_instruction_set": ("docx", "set"),
    "field_add": ("docx", "add"), "levels_add": ("docx", "add"),
    "shape_link_set": ("pptx", "set"), "note_runs_add": ("xlsx", "add"),
    "note_font_set": ("xlsx", "set"), "note_text_set": ("xlsx", "set"),
    "rich_comment_get": ("xlsx", "get"),
}


@needs("0325")
def p0325():
    """Record 0325 §6: ooxml_followups_bench (three runs) and the public-API
    read probe (B A A B B A; `A` sections are the after build)."""
    base = RES / "0325"
    path = base / "followups-bench.jsonl"
    out = []
    for scen, items in group_cases([load_jsonl(path)], lambda r: r["scenario"]).items():
        fmt, op = OOXML_0325[scen]
        out.append(row("0325", fmt, op, scen, "", med(i["leaf_median_us"] for i in items),
                       med(i["unified_median_us"] for i in items), [i["ratio"] for i in items],
                       READ_POINT if op == "get" else SINGLE, f"ooxml_followups:{scen}",
                       src([path])))
    probe = base / "read-regression.txt"
    acc = view_sections(probe, re.compile(r"^## (?P<label>[AB])\.\d+\.tsv$"), "A")
    out += view_rows("0325", lambda k: k.split()[0], "ooxml_read_probe", acc, src([probe]))
    return out


@needs("0328")
def p0328():
    """Record 0328 §6: odf_rtf_parts_bench trials and the summary of
    odf_rtf_view_regression (two runs; before/after min-max ranges)."""
    base = RES / "0328"
    out = parts_trials("0328", "odf_rtf_parts", trials(base.glob("parts-trial-*.json")),
                       lambda c: c.split("_", 1)[0], part_op)
    path = base / "view-regression.txt"
    for line in Path(path).read_text(encoding="utf-8").splitlines():
        parts = line.rstrip("\n").split("\t")
        if len(parts) != 5 or parts[0] == "scenario":
            continue
        k, before, after, r1, r2 = parts

        def mid(span):
            lo, hi = span.split("-")
            return (float(lo) + float(hi)) / 2

        out.append(row("0328", k.split("_", 1)[0], "set (view)" if k.endswith("_set") else "get (view)",
                       k, "", mid(before), mid(after), [float(r1), float(r2)],
                       regression("0328"), f"odf_rtf_view:{k}", src([path]),
                       comparison="after/before", unit="ns (mid-range)"))
    return out


@needs("0329")
def p0329():
    """Record 0329 §7: the after legs of keynote_bench, pages_bench and
    numbers_unified_bench (processes 2 and 3 of B A A B)."""
    return iwork_after("0329", RES / "0329", "{fmt}-*-after.{ext}")


@needs("0330")
def p0330():
    """Record 0330 §5: one 1,000-item coalesced XLSX cell batch, before and
    after the claim-list check (ABBA × 4, `abba.txt`)."""
    path = RES / "0330" / "abba.txt"
    legs = {"before": [], "after": []}
    for line in Path(path).read_text(encoding="utf-8").splitlines():
        parts = line.split()
        if len(parts) == 2 and parts[0] in legs:
            legs[parts[0]].append(float(parts[1]) / 1000)
    b, a = statistics.mean(legs["before"]), statistics.mean(legs["after"])
    return [row("0330", "xlsx", "batch", "1,000 cell value sets on create(Xlsx), coalesced",
                "1000 ops", b, a, [a / b], regression("0330"), "registry_check:xlsx_batch_1000",
                src([path]), comparison="after/before (mean of 8 medians)")]


@needs("0333")
def p0333():
    """Record 0333 §6: pptx_p0_bench, two runs (batches against one leaf
    edit staging the N text edits; single operations against the leaf
    verb)."""
    paths = trials((RES / "0333").glob("pptx-p0-bench-run*.jsonl"))
    out = []
    for scen, items in group_cases([load_jsonl(p) for p in paths], lambda r: r["scenario"]).items():
        if "/" in scen:
            family, owner, n = scen.split("/")
            n = int(n)
            op, label = "batch", f"{family.replace('_', ' ')} ({owner} owner; one File::batch vs one leaf edit)"
            size, budget = f"{n} ops", BATCH if n >= 1000 else SCALING
        else:
            op = "add" if "_add" in scen else "remove" if "_remove" in scen else "set"
            label, size, budget = scen, "", SINGLE
        out.append(row("0333", "pptx", op, label, size, med(i["leaf_median_us"] for i in items),
                       med(i["unified_median_us"] for i in items), [i["ratio"] for i in items],
                       budget, f"pptx_p0:{scen}", src(paths)))
    path = RES / "0333" / "batch-before-after.txt"
    vals: dict = {}
    for m in re.finditer(r"^(base|new)\s+eager n=(\d+) median_ms=([\d.]+)\s+source n=\d+ median_ms=([\d.]+)",
                         path.read_text(encoding="utf-8"), re.M):
        vals.setdefault(int(m.group(2)), {})[m.group(1)] = (float(m.group(3)), float(m.group(4)))
    for n, legs in sorted(vals.items()):
        if "base" not in legs:
            continue
        for i, owner in enumerate(("eager", "source")):
            b, a = legs["base"][i] * 1000, legs["new"][i] * 1000
            out.append(row("0333", "pptx", "batch",
                           f"run format batch ({owner} owner; per-item commits before, coalesced after)",
                           f"{n} ops", b, a, [a / b], regression("0333"),
                           f"pptx_p0_batch_build:{owner}:{n}", src([path]),
                           comparison="after/before (unified)"))
    return out


@needs("0335")
def p0335():
    """Record 0335 §6: docx_p0_residue_bench trials (warm handle; `_cold`
    cases open the bytes in both legs; `open_control` times the open alone)
    and docx_view_regression on two fixtures (B = before)."""
    base = RES / "0335"

    def ops(case):
        if case == "open_control":
            return "open", OPEN
        if case.endswith("_cold"):
            return "set (first write, cold handle)", SINGLE
        return part_op(case)

    out = parts_trials("0335", "docx_p0_residue", trials(base.glob("trial-*.json")),
                       lambda c: "docx", ops)
    for name in ("strict", "heading123"):
        path = base / f"view-{name}.txt"
        acc = view_sections(path, re.compile(rf"^== {name} (?P<label>[AB])$"), "A")
        out += view_rows("0335", lambda k: "docx", f"docx_view_0335:{name}",
                         {f"{name}.docx {k}": v for k, v in acc.items()}, src([path]))
    return out


@needs("0336")
def p0336():
    """Record 0336 §6: `File::dump` against a traversal reading what dump
    reads (three processes; the goal has no dump budget)."""
    path = RES / "0336" / "dump_bench.jsonl"
    labels = {"docx": "DOCX, 400 paragraphs (drawing.docx)",
              "xlsx": "XLSX, 75,770 cells (no_drawing_patriarch.xlsx)",
              "pptx": "PPTX, 200 generated slides"}
    out = []
    for scen, items in group_cases([load_jsonl(path)], lambda r: r["scenario"]).items():
        out.append(row("0336", scen, "dump", labels[scen], f"{items[0]['items']} items",
                       med(i["traverse_median_us"] for i in items),
                       med(i["dump_median_us"] for i in items), [i["ratio"] for i in items],
                       (None, "none (record 0336 §6: no dump budget in the goal)", "none"),
                       f"dump:{scen}", src([path]), comparison="dump/traversal"))
    return out


NATIVE_LINE = re.compile(r"^(\w+) (.+?): native (\d+) ns, unified (\d+) ns, ratio ([\d.]+)$")


def native_lines(record, harness, path, ops):
    """`<format> <scenario>: native N ns, unified N ns, ratio R` lines, one
    block per run (records 0338, 0342); `ops` maps a scenario to (operation,
    budget, comparison)."""
    acc: dict = {}
    for line in Path(path).read_text(encoding="utf-8").splitlines():
        m = NATIVE_LINE.match(line.strip())
        if m:
            acc.setdefault((m.group(1), m.group(2)), []).append(
                (int(m.group(3)), int(m.group(4)), float(m.group(5))))
    out = []
    for (fmt, scen), items in acc.items():
        op, budget, comparison = ops.get(scen, ("set", SINGLE, "unified/native"))
        out.append(row(record, fmt, op, scen, "", us(med(i[0] for i in items)),
                       us(med(i[1] for i in items)), [i[2] for i in items], budget,
                       f"{harness}:{fmt}:{slug(scen)}", src([path]), comparison=comparison))
    return out


@needs("0338")
def p0338():
    """Record 0338 §6 (p0_residue_bench, two runs) and §3 (json_number_bench,
    two builds, 6 ABBA blocks: after/before of the medians of 12 runs)."""
    base = RES / "0338"
    out = native_lines("0338", "p0_residue", base / "p0-residue-bench.txt", {
        "set / title": ("metadata set", SINGLE, "unified/native"),
        "set /metadata title vs set / title": ("metadata set", SINGLE,
                                               "set / title vs set /metadata title (one writer)"),
        "add cell (ref, value)": ("add", SINGLE, "unified/native"),
    })
    path = base / "json-number-bench.txt"
    legs: dict = {}
    for m in re.finditer(r"^(before|after) ([\w-]+): (\d+) ns median, (\d+) items, (\d+) bytes",
                         path.read_text(encoding="utf-8"), re.M):
        legs.setdefault((m.group(2), m.group(4), m.group(5)), {"before": [], "after": []})[
            m.group(1)].append(int(m.group(3)))
    labels = {"seeds": "OfficeCLI seeds (numbers as strings)",
              "cells-10k-short": "10k cell sets, short decimals",
              "cells-10k-long": "10k cell sets, 17-digit decimals with exponents",
              "cells-10k-int": "10k cell sets, integers"}
    for (payload, items, size), v in legs.items():
        b, a = us(med(v["before"])), us(med(v["after"]))
        out.append(row("0338", "batch", "decode", f"batch::from_bytes, {labels[payload]}",
                       f"{items} items, {int(size):,} bytes", b, a, [a / b], regression("0338"),
                       f"json_number:{payload}", src([path]),
                       comparison="after/before (float_roundtrip; medians of 12 runs)"))
    return out


@needs("0342")
def p0342():
    """Record 0342 §6: v1_closure_bench, two runs. The table-vs-row path
    resolution leg is an attribution of the row set's cost, not a pair."""
    return native_lines("0342", "v1_closure", RES / "0342" / "v1-closure-bench.txt", {
        "set / extended.company": ("metadata set", SINGLE, "unified/native"),
        "set / title": ("metadata set", SINGLE, "unified/native"),
        "resolve table vs row path (attribution)": (
            "resolve", (None, "none (record 0342 §6: attribution leg)", "none"),
            "row path/table path resolution"),
        "add paragraph listStyle=bullet": ("add", SINGLE, "unified/native"),
    })


@needs("0345")
def p0345():
    """Record 0345 §6: docx_p0_final_bench, three runs. The resolution,
    leaf-read and view legs are attributions, not unified/leaf pairs."""
    attribution = (None, "none (record 0345 §6: attribution leg)", "none")
    return native_lines("0345", "p0_final", RES / "0345" / "docx-p0-final-bench.txt", {
        "add cell": ("add", SINGLE, "unified/native"),
        "add column": ("add", SINGLE, "unified/native"),
        "add comment reply": ("add", SINGLE, "unified/native"),
        "add plain comment": ("add", SINGLE, "unified/native"),
        "resolve /body/tbl[1] vs /body/tbl[1]/col[1]": (
            "resolve", attribution, "column path/table path resolution"),
        "leaf document() vs document().comments()": (
            "read", attribution, "comments read/document handle (leaf)"),
        "view format: /body/p[2] vs /comments/comment[1]": (
            "view", attribution, "comment view/paragraph view"),
    })


@needs("0346")
def p0346():
    """Record 0346 §6: pptx_p0_final_bench, two runs (batches against one
    leaf edit staging the same text edits, `cold` on a fresh handle per
    batch and `warm` on one; single operations against the leaf verb)."""
    paths = trials((RES / "0346").glob("pptx-p0-final-bench-run*.jsonl"))
    out = []
    for scen, items in group_cases([load_jsonl(p) for p in paths], lambda r: r["scenario"]).items():
        if "/" in scen:
            family, owner, n, *handle = scen.split("/")
            n = int(n)
            state = f", {handle[0]} handle" if handle else ""
            op = "batch"
            label = f"{family.replace('_', ' ')} ({owner} owner{state}; one File::batch vs one leaf edit)"
            size, budget = f"{n} ops", BATCH if n >= 1000 else SCALING
        else:
            op = "add" if "_add" in scen else "set"
            label, size, budget = scen, "", SINGLE
        out.append(row("0346", "pptx", op, label, size, med(i["leaf_median_us"] for i in items),
                       med(i["unified_median_us"] for i in items), [i["ratio"] for i in items],
                       budget, f"pptx_p0_final:{scen}", src(paths)))
    return out


@needs("0348")
def p0348():
    """Record 0348 §6: every PPTX batch route re-measured, three runs each:
    pptx_p0_final_bench (record 0346's routes, cold and warm handles) and
    the batch routes of pptx_p0_bench (record 0333's), under the series of
    those records, which these rows supersede."""
    base = RES / "0348"
    out = []
    for prefix, series, glob in (("pptx_p0_final", "pptx_p0_final", "pptx-p0-final-bench-run*.jsonl"),
                                 ("pptx_p0", "pptx_p0", "pptx-p0-bench-run*.jsonl")):
        paths = trials(base.glob(glob))
        for scen, items in group_cases([load_jsonl(p) for p in paths], lambda r: r["scenario"]).items():
            if "/" not in scen:
                continue
            family, owner, n, *handle = scen.split("/")
            n = int(n)
            state = f", {handle[0]} handle" if handle else ""
            label = f"{family.replace('_', ' ')} ({owner} owner{state}; one File::batch vs one leaf edit)"
            out.append(row("0348", "pptx", "batch", label, f"{n} ops",
                           med(i["leaf_median_us"] for i in items),
                           med(i["unified_median_us"] for i in items), [i["ratio"] for i in items],
                           BATCH if n >= 1000 else SCALING, f"{series}:{scen}", src(paths)))
    return out


@needs("0347")
def p0347():
    """Record 0347 §6: numbers_unified_bench swaps (two runs, with the
    sheet and table moves as controls) and the p0_residue_bench picture
    `set src` after the 0338-P1 fix (the bench's other rows are unchanged
    paths, recorded under 0338)."""
    base = RES / "0347"
    paths = trials(base.glob("numbers-swap-run*.json"))
    out = numbers_docs("0347", [load_json(p)["scenarios"] for p in paths], paths)
    out += [r for r in native_lines("0347", "p0_residue", base / "p0-residue-bench.txt", {})
            if "picture" in r["scenario"]]
    return out


W1_DOCX_OPS = {
    "open (1,000 paragraphs)": ("open", OPEN, "unified/native"),
    "set paragraph text": ("set", SINGLE, "unified/native"),
    "add paragraph": ("add", SINGLE, "unified/native"),
    "set / title": ("metadata set", SINGLE, "unified/native"),
    "File::edit 1k paragraph text sets": ("batch", BATCH, "unified/native (one leaf edit)"),
    "File::batch 1k paragraph text sets": ("batch", BATCH, "unified/native (one leaf edit)"),
    "File::batch 1k paragraph text sets (eager owner)": ("batch", BATCH,
                                                          "unified/native (one leaf edit)"),
    "save (to_bytes) after one set": ("save", SINGLE, "unified/native"),
    "repeated save of one package (second to_bytes)": ("save", SINGLE, "unified/native"),
}


@needs("0355")
def p0355():
    """Record 0355 §5: docx_w1_contract_bench, release, pinned, ABBA (two
    base and two head runs). Head rows pair the unified call with the leaf
    call; after/before rows compare each leg at the wave's head with the same
    leg at its base (e652864de8)."""
    base, head = RES / "0355" / "docx-w1-contract-base.txt", RES / "0355" / "docx-w1-contract-head.txt"
    out = native_lines("0355", "docx_w1_contract", head, W1_DOCX_OPS)
    legs: dict = {}
    for leg, path in (("before", base), ("after", head)):
        for line in path.read_text(encoding="utf-8").splitlines():
            m = NATIVE_LINE.match(line.strip())
            if m:
                slot = legs.setdefault(m.group(2), {"before": ([], []), "after": ([], [])})[leg]
                slot[0].append(int(m.group(3)))
                slot[1].append(int(m.group(4)))
    for scen, v in legs.items():
        op = W1_DOCX_OPS.get(scen, ("set",))[0]
        for index, name in ((0, "leaf"), (1, "unified")):
            b, a = us(med(v["before"][index])), us(med(v["after"][index]))
            out.append(row("0355", "docx", op, f"{scen} ({name} leg, W1 head vs base)", "",
                           b, a, [a / b], regression("0355"),
                           f"docx_w1_contract:{name}:{slug(scen)}:after-before",
                           src([base, head]), comparison="after/before"))
    return out


FAMILY_LINE = re.compile(r"^docx-family (.+?): median (\d+) ns, min (\d+) ns, output (\d+) bytes, sha256 ([0-9a-f]{64})$")


@needs("0355a")
def p0355a():
    """Record 0355a §5: docx_w1_contract_bench at the 0355a head (the P1
    batch coalescing; unified/leaf rows) and against the 0355 tree
    (3e686506f3, after/before per leg), and source_family_bench (each
    source-backed family, open → edit → commit → write, against the same
    edit through the record-0355 stream publisher). Release, pinned, ABBA
    (base, head, head, base)."""
    res = RES / "0355a"
    base, head = res / "docx-w1-contract-base.txt", res / "docx-w1-contract-head.txt"
    out = native_lines("0355a", "docx_w1_contract", head, W1_DOCX_OPS)
    legs: dict = {}
    for leg, path in (("before", base), ("after", head)):
        for line in path.read_text(encoding="utf-8").splitlines():
            m = NATIVE_LINE.match(line.strip())
            if m:
                slot = legs.setdefault(m.group(2), {"before": ([], []), "after": ([], [])})[leg]
                slot[0].append(int(m.group(3)))
                slot[1].append(int(m.group(4)))
    for scen, v in legs.items():
        if not v["before"][0] or not v["after"][0]:
            continue
        op = W1_DOCX_OPS.get(scen, ("set",))[0]
        for index, name in ((0, "leaf"), (1, "unified")):
            b, a = us(med(v["before"][index])), us(med(v["after"][index]))
            out.append(row("0355a", "docx", op, f"{scen} ({name} leg, 0355a head vs 0355)", "",
                           b, a, [a / b], regression("0355a"),
                           f"docx_w1_contract:{name}:{slug(scen)}:0355a-after-before",
                           src([base, head]), comparison="after/before"))
    fbase, fhead = res / "source-family-base.txt", res / "source-family-head.txt"
    fam: dict = {}
    for leg, path in (("before", fbase), ("after", fhead)):
        for line in path.read_text(encoding="utf-8").splitlines():
            m = FAMILY_LINE.match(line.strip())
            if m:
                fam.setdefault(m.group(1), {"before": [], "after": []})[leg].append(int(m.group(2)))
    for scen, v in fam.items():
        b, a = us(med(v["before"])), us(med(v["after"]))
        ratios = [x / y for x, y in zip(v["after"], v["before"])]
        out.append(row("0355a", "docx", "save", f"{scen} (source-backed family vs publisher)",
                       "1,000 paragraphs + 2 MiB media", b, a, ratios, regression("0355a"),
                       f"source_family:{slug(scen)}:after-before", src([fbase, fhead]),
                       comparison="after/before (family commit + write_to vs publish_*_to_stream)"))
    return out


@needs("0357")
def p0357():
    """Record 0357 (W3 PPT contract): ppt_paired_probe single/batch and
    ppt_elements_probe, binary-level ABBA; the after legs are listed."""
    out = []
    paired = sorted((RES / "0357").glob("single-after-*.json")) + sorted(
        (RES / "0357").glob("batch-after-*.json"))
    groups: dict = {}
    for path in paired:
        for r in load_json(path):
            groups.setdefault((r["scenario"], r["size"]), []).append(r)
    for (scen, size), items in sorted(groups.items()):
        out.append(ppt_paired_row("0357", scen, size,
                                  med(i["native_median_us"] for i in items),
                                  med(i["unified_median_us"] for i in items),
                                  [i["pair_ratio_median"] for i in items],
                                  src(paired), "after record 0357", 2))
    paths = sorted((RES / "0357").glob("elements-after-*.json"))
    for (scen, n), items in group_cases([load_json(p) for p in paths],
                                        lambda r: (r["scenario"], r["slides"])).items():
        op, label = PPT_ELEMENTS[scen]
        out.append(row("0357", "ppt", op, label, f"{n} slides",
                       med(i["native_median_us"] for i in items),
                       med(i["unified_median_us"] for i in items),
                       [i["pair_ratio_median"] for i in items],
                       READ_POINT if op == "get" else SINGLE, f"ppt_elements:{scen}:{n}",
                       src(paths), "after record 0357", 2, comparison="unified/native"))
    return out


CONTRACT_0356 = {
    "open": ("open", OPEN), "save": ("save", SINGLE), "save_twice": ("save", SINGLE),
    "set": ("set", SINGLE), "set_save": ("set", SINGLE), "batch1k": ("batch", BATCH),
}


def pptx_contract_rows(record):
    """The PPTX batch routes of record 0348 re-measured (three runs each,
    superseding the earlier rows) and pptx_contract_bench, one source built
    on the base tree and on the branch, run A-B-B-A three times: after/before
    rows per scenario (median of the after runs over the median of the
    before runs). Records 0356 and 0356a."""
    base = RES / record
    out = []
    for series, glob in (("pptx_p0_final", "pptx-p0-final-bench-run*.jsonl"),
                         ("pptx_p0", "pptx-p0-bench-run*.jsonl")):
        paths = trials(base.glob(glob))
        for scen, items in group_cases([load_jsonl(p) for p in paths], lambda r: r["scenario"]).items():
            if "/" not in scen:
                continue
            family, owner, n, *handle = scen.split("/")
            n = int(n)
            state = f", {handle[0]} handle" if handle else ""
            label = f"{family.replace('_', ' ')} ({owner} owner{state}; one File::batch vs one leaf edit)"
            out.append(row(record, "pptx", "batch", label, f"{n} ops",
                           med(i["leaf_median_us"] for i in items),
                           med(i["unified_median_us"] for i in items), [i["ratio"] for i in items],
                           BATCH if n >= 1000 else SCALING, f"{series}:{scen}", src(paths)))
    before = trials(base.glob("contract-before-run*.jsonl"))
    after = trials(base.glob("contract-after-run*.jsonl"))
    b_rows = [r for p in before for r in load_jsonl(p)]
    srcs = src(list(before) + list(after))
    for scen, a_items in group_cases([load_jsonl(p) for p in after], lambda r: r["scenario"]).items():
        route, leg, fixture = scen.split("/")
        op, _ = CONTRACT_0356[route]
        b = med(r["median_us"] for r in b_rows if r["scenario"] == scen)
        a = med(r["median_us"] for r in a_items)
        if b is None:
            continue
        out.append(row(record, "pptx", op, f"{route.replace('_', ' ')} ({fixture})", "",
                       b, a, [a / b], regression(record), f"pptx_contract:{scen}", srcs,
                       comparison=f"after/before ({leg} leg)"))
    return out


W6_ODF_OPS = {"open_get": "open", "set": "set", "batch1k": "batch"}


@needs("0358")
def p0358():
    """Record 0358 §8 (W6 ODF contract): odf_w6_probe, one source built on
    the wave's base (trunk) and head, run A-B-B-A pinned in release; one
    after/before row per unified route (median of the after runs over the
    median of the before runs)."""
    base = RES / "0358"
    legs: dict = {"before": {}, "after": {}}
    paths = sorted(base.glob("probe-*-*.json"))
    for path in paths:
        leg = path.name.split("-")[1]
        for r in load_json(path):
            legs[leg].setdefault(r["scenario"], []).append(r["median_us"])
    out = []
    for scen, after in legs["after"].items():
        before = legs["before"].get(scen)
        if not before:
            continue
        fmt, op = scen.split("/")
        b, a = med(before), med(after)
        size = "1k ops" if op == "batch1k" else ""
        out.append(row("0358", fmt, W6_ODF_OPS.get(op, op), f"File {op.replace('_', ' ')} ({fmt}, W6 head vs base)",
                       size, b, a, [a / b], regression("0358"), f"odf_w6_probe:{scen}:after-before",
                       src(paths), comparison="after/before"))
    return out


W6_SOURCE_BACKED = {
    "ods/cells1": ("ods", "set", "SourceBackedSpreadsheet open + edit_cells (1 cell) + commit + write_to", "12k×8 cells"),
    "ods/cells100": ("ods", "set", "SourceBackedSpreadsheet open + edit_cells (100 cells) + commit + write_to", "12k×8 cells"),
    "ods/dde": ("ods", "set", "SourceBackedSpreadsheet open + edit_dde + commit + write_to", "12k×8 cells"),
    "ods/sheet_metadata": ("ods", "set", "SourceBackedSpreadsheet open + edit_sheet_metadata + commit + write_to", "12k×8 cells"),
    "odp/tail_append": ("odp", "add", "SourceBackedPresentation open + edit_tail_append + commit + write_to", "2,000 slides"),
}
W6_CODEGEN = ("explained residual: the same work (callgrind instructions within 0.3%); the "
              "fat-LTO codegen-unit partition and its inlining moved with the source-backed "
              "code added to the binary; with codegen-units=1 the ratio is {cgu1} "
              "(results/0358a/codegen.json)")


@needs("0361")
def p0361():
    """Record 0361 §8 (W7 iWork): the after legs of keynote_bench,
    pages_bench and numbers_unified_bench (processes 2 and 3 of B A A B) and
    the batch-closure 1k/10k iWork batches (`abba/batch-A-after-*.json`)."""
    base = RES / "0361"
    return (iwork_after("0361", base, "{fmt}-*-after.{ext}")
            + batch_closure("0361", sorted((base / "abba").glob("batch-A-after-*.json")),
                            stages={"after": ("", 1)}))


@needs("0361a")
def p0361a():
    """Record 0361a §4: the after legs of keynote_bench, pages_bench and
    numbers_unified_bench (processes 2, 3, 6 and 7 of B A A B B A A B) after
    the wire-codec limits change."""
    return iwork_after("0361a", RES / "0361a", "{fmt}-*-after.{ext}")


@needs("0358a")
def p0358a():
    """Record 0358a §6: the source-backed ODF families (after) against the
    former stream routes (before), byte-identical output, and the record
    0358 routes (0358a head vs 0358 head), both A-B-B-A pinned in release."""
    base = RES / "0358a"
    out = []
    for probe, label in (("sb", "after/before"), ("w6", "after/before")):
        legs: dict = {"before": {}, "after": {}}
        paths = sorted(base.glob(f"{probe}-*-*.json"))
        for path in paths:
            leg = path.name.split("-")[1]
            for r in load_json(path):
                legs[leg].setdefault(r["scenario"], []).append(r["median_us"])
        for scen, after in legs["after"].items():
            before = legs["before"].get(scen)
            if not before:
                continue
            b, a = med(before), med(after)
            if probe == "sb":
                fmt, op, scenario, size = W6_SOURCE_BACKED[scen]
                scenario += ", family route vs former stream route"
                explanation = ""
            else:
                fmt, op = scen.split("/")
                scenario = f"File {op.replace('_', ' ')} ({fmt}, 0358a head vs 0358 head)"
                size = "1k ops" if op == "batch1k" else ""
                op = W6_ODF_OPS.get(op, op)
                cgu1 = {"odt/set": "1.014", "fodt/open_get": "1.032"}.get(scen)
                explanation = W6_CODEGEN.format(cgu1=cgu1) if cgu1 and a / b > 1.05 else ""
            out.append(row("0358a", fmt, op, scenario, size, b, a, [a / b], regression("0358a"),
                           f"{probe}:{scen}:after-before", src(paths), comparison=label,
                           explanation=explanation))
    return out


W5_LEGACY_OPS = {"open_get": "open", "get": "get", "set": "set", "add": "add", "batch1k": "batch",
                 "batch256": "batch"}


@needs("0359")
def p0359():
    """Record 0359 §7 (W5 DOC/XLS/RTF contract): legacy_w5_probe, one source
    built on the wave's base (trunk) and head, run A-B-B-A pinned in
    release; one after/before row per unified route (median of the after
    runs over the median of the before runs)."""
    base = RES / "0359"
    legs: dict = {"before": {}, "after": {}}
    paths = sorted(base.glob("probe-*-*.json"))
    for path in paths:
        leg = path.name.split("-")[1]
        for r in load_json(path):
            legs[leg].setdefault(r["scenario"], []).append(r["median_us"])
    out = []
    for scen, after in legs["after"].items():
        before = legs["before"].get(scen)
        if not before:
            continue
        fmt, op = scen.split("/")
        b, a = med(before), med(after)
        size = {"batch1k": "1k ops", "batch256": "256 ops"}.get(op, "")
        out.append(row("0359", fmt, W5_LEGACY_OPS.get(op, op),
                       f"File {op.replace('_', ' ')} ({fmt}, W5 head vs base)",
                       size, b, a, [a / b], regression("0359"), f"legacy_w5_probe:{scen}:after-before",
                       src(paths), comparison="after/before"))
    return out


W8_OPS = {"open_get": "open", "set": "set", "batch1k": "batch", "import": "import", "merge": "merge"}
W8_UNQUALIFIED = ("follow-up 0362ha: historical measurements retained; the blanket 1.05 "
                  "threshold and code-placement explanation were returned on review. "
                  "Matched current-source performance qualification remains required.")


def w8_legs(directory):
    legs: dict = {"before": {}, "after": {}}
    paths = sorted(directory.glob("probe-*-*.json"))
    for path in paths:
        leg = path.stem.split("-")[2]
        for r in load_json(path):
            legs[leg].setdefault(r["scenario"], []).append(r["median_us"])
    return legs, paths


@needs("0362")
def p0362():
    """Record 0362 §5 (W8, the unified `File`): file_w8_probe, one source
    built on the wave's base (trunk) and head, run A-B-B-A twice pinned in
    release; one after/before row per format and route (run 1's medians, the
    ratio range over both runs)."""
    base = RES / "0362"
    (run1, paths1), (run2, paths2) = w8_legs(base / "run1"), w8_legs(base / "run2")
    out = []
    for scen, after in run1["after"].items():
        before = run1["before"].get(scen)
        if not before:
            continue
        fmt, op = scen.split("/")
        b, a = med(before), med(after)
        ratios = [a / b]
        if run2["after"].get(scen) and run2["before"].get(scen):
            ratios.append(med(run2["after"][scen]) / med(run2["before"][scen]))
        size = {"batch1k": "256 ops" if fmt == "doc" else "1k ops",
                "import": "100 rows" if fmt == "xls" else "1k rows"}.get(op, "")
        out.append(row("0362", fmt, W8_OPS.get(op, op),
                       f"File {op.replace('_', ' ')} ({fmt}, W8 head vs base)",
                       size, b, a, ratios, NONE, f"file_w8_probe:{scen}:after-before",
                       src(paths1 + paths2), comparison="after/before", explanation=W8_UNQUALIFIED))
    return out


W5_SOURCE_BACKED = {
    "doc/source_set": ("doc", "set", "source_backed::Package open + edit_paragraph + set_text + commit + write_to", "1k paragraphs"),
    "xls/source_cells1": ("xls", "set", "SourceBackedWorkbook open + edit_cells (1 cell) + commit + write_to", "100×10 cells"),
    "xls/source_cells100": ("xls", "set", "SourceBackedWorkbook open + edit_cells (100 cells) + commit + write_to", "100×10 cells"),
    "xls/source_comments1": ("xls", "set", "SourceBackedWorkbook open + edit_comments (1 comment) + commit + write_to", "256 comments"),
    "xls/source_comments256": ("xls", "set", "SourceBackedWorkbook open + edit_comments (256 comments) + commit + write_to", "256 comments"),
    "xls/source_hide7": ("xls", "set", "SourceBackedWorkbook open + edit_sheet_visibility (7 sheets) + commit + write_to", "8 sheets"),
    "rtf/source_append": ("rtf", "add", "source_backed::Package open + edit_tail_append + commit + write_to", "1k paragraphs"),
}
W5A_SAME_WORK = ("explained residual: the same work (callgrind instructions within 0.03% on both builds, "
                 "results/0359a/callgrind-totals.txt); the other A-B-B-A run is {other}")
W5A_LEAF = ("explained residual: process state, not the leaf edit. The leaf batch runs after the ops "
            "batch in one process; on the base that batch faults in about 15 GB first, and the leaf "
            "reuses the warmed heap. Alone in its process the leaf edit takes the same time on both "
            "builds (1.684-1.688 s vs 1.669-1.690 s, results/0359a/doc-batch-isolated.txt)")


def w5a_legs(paths):
    legs: dict = {"before": {}, "after": {}}
    for path in paths:
        leg = path.name.split("-")[1]
        for r in load_json(path):
            legs[leg].setdefault(r["scenario"], []).append(r["median_us"])
    return legs


@needs("0359a")
def p0359a():
    """Record 0359a §7: legacy_w5_source_probe (the source-backed routes and
    the DOC 1k batch; only its `routes` module differs between the base,
    record 0359's tip, and the head) and legacy_w5_probe (the 0359 routes),
    A-B-B-A twice pinned in release. Ratios span both runs (run 2 is
    `results/0359a/*.json`, run 1 `results/0359a/run1/`); the DOC batch row
    is the head's unified/leaf ratio per run."""
    base = RES / "0359a"
    runs = [base, base / "run1"]
    out = []
    for probe in ("source", "ops"):
        per_run = [w5a_legs(sorted(r.glob(f"{probe}-*-*.json"))) for r in runs]
        paths = sorted(base.glob(f"{probe}-*-*.json"))
        for scen, after in per_run[0]["after"].items():
            before = per_run[0]["before"].get(scen)
            if not before:
                continue
            b, a = med(before), med(after)
            ratios = [med(legs["after"][scen]) / med(legs["before"][scen]) for legs in per_run
                      if scen in legs["after"] and scen in legs["before"]]
            if probe == "source" and scen in W5_SOURCE_BACKED:
                fmt, op, scenario, size = W5_SOURCE_BACKED[scen]
                scenario += ", family route vs former publisher"
                explanation = ""
                series = f"legacy_w5_source_probe:{scen}:after-before"
            elif probe == "source":
                fmt, op = scen.split("/")
                label = {"doc/batch1k_ops": "File::edit_with_limits, 1,000 paragraph bold sets",
                         "doc/batch1k_leaf": "Package::edit_with_options, 1,000 set_paragraph_bold"}[scen]
                scenario, size, op = f"{label}, 0359a head vs 0359 head", "1k ops", "batch"
                explanation = W5A_LEAF if scen == "doc/batch1k_leaf" else ""
                series = f"legacy_w5_source_probe:{scen}:after-before"
            else:
                fmt, op = scen.split("/")
                scenario = f"File {op.replace('_', ' ')} ({fmt}, 0359a head vs 0359 head)"
                size = {"batch1k": "1k ops", "batch256": "256 ops"}.get(op, "")
                op = W5_LEGACY_OPS.get(op, op)
                explanation = ""
                if max(ratios) > 1.05:
                    other = ", ".join(f"{r:.3f}" for r in ratios)
                    explanation = W5A_SAME_WORK.format(other=other)
                series = f"legacy_w5_probe:{scen}:after-before"
            out.append(row("0359a", fmt, op, scenario, size, b, a, ratios, regression("0359a"),
                           series, src(paths), comparison="after/before", explanation=explanation))
        if probe == "source":
            for scen in ("xls/source_cells1", "xls/source_cells100"):
                plan = per_run[0]["before"].get(scen + "_plan")
                if not plan:
                    continue
                b, a = med(plan), med(per_run[0]["after"][scen])
                ratios = [med(legs["after"][scen]) / med(legs["before"][scen + "_plan"])
                          for legs in per_run if scen + "_plan" in legs["before"]]
                fmt, op, scenario, size = W5_SOURCE_BACKED[scen]
                out.append(row("0359a", fmt, op, scenario + ", family route vs former plan-only publisher",
                               size, b, a, ratios, regression("0359a"),
                               f"legacy_w5_source_probe:{scen}:after-before_plan", src(paths),
                               comparison="after/before",
                               explanation=("explained residual: the family commit reopens the "
                                            "committed workbook over the plan's composed source, "
                                            "which the forward-only plan publisher did not; the "
                                            "former materializing publisher it replaces is the "
                                            "after/before row") if max(ratios) > 1.05 else ""))
    per_run = [w5a_legs(sorted(r.glob("source-after-*.json")) + sorted(r.glob("source-before-*.json")))
               for r in runs]
    ratios = [med(legs["after"]["doc/batch1k_ops"]) / med(legs["after"]["doc/batch1k_leaf"])
              for legs in per_run]
    legs = per_run[0]
    out.append(row("0359a", "doc", "batch",
                   "File::edit_with_limits, 1,000 paragraph bold sets vs the leaf edit with its operation bound raised through EditOptions",
                   "1k ops", med(legs["after"]["doc/batch1k_leaf"]), med(legs["after"]["doc/batch1k_ops"]),
                   ratios, BATCH, "legacy_w5_source_probe:doc/batch1k:unified-leaf",
                   src(sorted(base.glob("source-after-*.json")))))
    return out


W4B_XLSB_OPS = {"open": "open", "open_get": "open", "set": "set", "set_save": "save", "format": "set",
                "metadata": "metadata", "batch1k": "batch", "comment": "add", "hyperlink": "add",
                "name": "add", "merge": "set", "chain": "set"}


def w4b_legs(base: Path):
    """The (before, after) medians per scenario of one record 0360b
    A-B-B-A run: the probe, the 0353a XLSB sheet scenarios and xlsb_crud."""
    legs: dict = {"before": {}, "after": {}}
    for path in sorted(base.glob("probe-*-*.json")):
        for r in load_json(path):
            legs[path.name.split("-")[1]].setdefault(r["scenario"], []).append(r["median_us"])
    for path in sorted(base.glob("sheet-*-*.txt")):
        for name, v in load_json(path)["scenarios"].items():
            if name.endswith("/xlsb"):
                legs[path.name.split("-")[1]].setdefault("sheet/" + name, []).append(v["median_us"])
    for path in sorted(base.glob("crud-*-*.json")):
        for case in load_json(path)["cases"]:
            legs[path.name.split("-")[1]].setdefault("crud/" + case["case"], []).append(
                case["statistics"]["p50_ns"] / 1000.0)
    return legs


@needs("0360b")
def p0360b():
    """Record 0360b §6 (W4 XLSB contract): xlsb_w4b_probe (leaf and unified
    routes with calls stable across the wave), sheet_dispatch_bench (the
    0353a XLSB scenarios) and perf-baseline xlsb_crud, base and head built
    in release and run A-B-B-A twice pinned (run 1 `results/0360b/`, run 2
    `results/0360b/run2/`); plus the head's unified 1k batch against the
    one-commit leaf edit."""
    base = RES / "0360b"
    runs = [w4b_legs(base), w4b_legs(base / "run2")]
    paths = sorted(base.glob("*-*-*.json")) + sorted(base.glob("sheet-*-*.txt"))
    out = []
    for scen, after in runs[0]["after"].items():
        before = runs[0]["before"].get(scen)
        if not before:
            continue
        b, a = med(before), med(after)
        ratios = [med(legs["after"][scen]) / med(legs["before"][scen]) for legs in runs
                  if scen in legs["after"] and scen in legs["before"]]
        if scen.startswith("sheet/"):
            op, fmt = scen[len("sheet/"):].split("/")
            scenario, size, series = f"sheet trait {op} (0353a), W4b head vs base", "2k rows", "sheet_dispatch_bench"
            op = "read"
        elif scen.startswith("crud/"):
            fmt, op = "xlsb", "edit" if "save" in scen else "read"
            scenario, size, series = f"xlsb_crud {scen[5:]}, W4b head vs base", "testVarious.xlsb", "xlsb_crud"
        else:
            owner, name = scen.split("/")
            fmt, op = "xlsb", W4B_XLSB_OPS.get(name, name)
            what = "leaf Package" if owner == "leaf" else f"File ({owner} owner)"
            scenario = f"{what} {name.replace('_', ' ')}, W4b head vs base"
            size, series = ("1k ops" if name == "batch1k" else "100×10 cells"), "xlsb_w4b_probe"
        out.append(row("0360b", fmt, op, scenario, size, b, a, ratios, regression("0360b"),
                       f"{series}:{scen}:after-before", src(paths), comparison="after/before"))
    for owner in ("eager", "source"):
        ratios = [med(legs["after"][f"{owner}/batch1k"]) / med(legs["after"]["leaf/batch1k"])
                  for legs in runs]
        legs = runs[0]
        out.append(row("0360b", "xlsb", "batch",
                       f"File::edit, 1,000 cell sets ({owner} owner) vs one leaf edit_cell_values commit",
                       "1k ops", med(legs["after"]["leaf/batch1k"]), med(legs["after"][f"{owner}/batch1k"]),
                       ratios, BATCH, f"xlsb_w4b_probe:{owner}/batch1k:unified-leaf",
                       src(sorted(base.glob("probe-after-*.json")))))
    return out


@needs("0357a")
def p0357a():
    """Record 0357a: record-sized scoped-family patches (1,000 retained
    single-op external-media commits, after vs before) and the restored
    source-backed text splice (after vs before on the primary edit; the
    splice itself has no before leg)."""
    out = []
    path = RES / "0357a" / "summary.json"
    s = load_json(path)
    for key, op, label in (("commit_median_us", "set", "scoped commit (external media path)"),
                           ("apply_median_us", "set", "scoped patch apply"),
                           ("inverse_apply_median_us", "set", "scoped patch inverse apply")):
        b, a = s["before"][key], s["after"][key]
        out.append(row("0357a", "ppt", op, f"{label}, after vs before", "1,000 patches",
                       b, a, [a / b], regression("0357a"),
                       f"scoped_patch:{slug(label)}:after-before", src([path]),
                       comparison="after/before"))
    path = RES / "0357a" / "text" / "summary.json"
    p = load_json(path)["probe"]
    b, a = p["before/primary"]["total_us"], p["after/primary"]["total_us"]
    out.append(row("0357a", "ppt", "set", "open + set_shape_text + commit + write_to, after vs before",
                   "1,000 shapes", b, a, [a / b], regression("0357a"),
                   "text_splice:primary:after-before", src([path]), comparison="after/before"))
    return out



@needs("0357b")
def p0357b():
    """Record 0357b: incremental block-root digests for the scoped-family
    patch (0357a-F1). The same probe as 0357a, ABC-CBA over the pre-0357a
    base, the 0357a tip and this record; the rows compare this record with
    the pre-0357a base, continuing the 0357a series, plus the first commit
    on a never-hashed package (a once-per-package cost, no budget)."""
    out = []
    path = RES / "0357b" / "summary.json"
    s = load_json(path)
    for key, op, label in (("commit_median_us", "set", "scoped commit (external media path)"),
                           ("apply_median_us", "set", "scoped patch apply"),
                           ("inverse_apply_median_us", "set", "scoped patch inverse apply")):
        b, a = s["base"][key], s["b"][key]
        out.append(row("0357b", "ppt", op, f"{label}, after vs before", "1,000 patches",
                       b, a, [a / b], regression("0357b"),
                       f"scoped_patch:{slug(label)}:after-before", src([path]),
                       comparison="after/before"))
    b, a = s["base"]["first_commit_median_us"], s["b"]["first_commit_median_us"]
    out.append(row("0357b", "ppt", "set",
                   "first scoped commit on a never-hashed package, after vs before",
                   "200 packages", b, a, [a / b], NONE,
                   "scoped_patch:first_commit_never_hashed:after-before", src([path]),
                   comparison="after/before",
                   explanation="Record 0357b §3: once per package, the first scoped commit "
                               "builds the source's block list (one SHA-256 pass over the "
                               "43,520-byte source); every later commit or apply of that "
                               "package or of a package published from it reuses it. "
                               "The 0357a tip paid two whole-image hashes here."))
    return out


@needs("0356")
def p0356():
    """Record 0356 §5 (see pptx_contract_rows)."""
    return pptx_contract_rows("0356")


@needs("0356a")
def p0356a():
    """Record 0356a §4: the 0356 routes on the review follow-ups, against
    trunk 184dd4d908 (record 0356 merged)."""
    return pptx_contract_rows("0356a")


@needs("0356b")
def p0356b():
    """Record 0356b §4: the source-backed cross-slide-copy routes after the
    copy moved from `SourceBackedPresentationEditor` onto
    `source::Package::edit_cross_slide_copy`, perf-baseline built on
    3e392d750f (before) and on the branch (after), run A-B-B-A three times:
    after/before per case (median of the per-run p50s)."""
    path = RES / "0356b" / "summary.json"
    out = []
    for case, v in load_json(path)["cases"].items():
        b, a = v["before_median_us"], v["after_median_us"]
        out.append(row("0356b", "pptx", "save", f"{case} (cross-slide copy, after vs before)",
                       "", b, a, [a / b], regression("0356b"),
                       f"perf_baseline:{case}:after-before", src([path]),
                       comparison="after/before"))
    return out

@needs("0356c")
def p0356c():
    """PPTX manifest fidelity: identical raw-edit probe, three isolated ABBAs."""
    path = RES / "0356c" / "summary.json"
    out = []
    for case, values in load_json(path)["cases"].items():
        before, after = values["before_median_us"], values["after_median_us"]
        out.append(row("0356c", "pptx", "edit", case, "", before, after,
                       [after / before], regression("0356c"),
                       f"opc_manifest:{case}:after-before", src([path]),
                       comparison="after/before",
                       explanation="Three isolated ABBA trials; identical probe and separate targets."))
    return out


@needs("0359b")
def p0359b():
    """RTF shared refusals: identical native probe, three isolated ABBAs."""
    path = RES / "0359b" / "summary.json"
    out = []
    for case, values in load_json(path)["cases"].items():
        before, after = values["before_median_us"], values["after_median_us"]
        out.append(row("0359b", "rtf", "edit", case, "", before, after,
                       [after / before], regression("0359b"),
                       f"refusal_contract:{case}:after-before", src([path]),
                       comparison="after/before",
                       explanation="Three isolated ABBA trials; identical probe and separate targets."))
    return out


def metric_row(record, fmt, op, scenario, size, before, after, unit, series,
               path, stage, comparison, explanation, verdict=None):
    """Add a measured diagnostic pair without treating it as a budget claim.

    The later W5 and W8 receipts contain useful allocator and instruction
    comparisons, but most deliberately do not contain a quiet native
    qualification.  Keeping these pairs in the matrix makes the measurement
    phase and metric explicit while ``NONE`` keeps the matrix from turning a
    diagnostic ratio into an acceptance decision.
    """
    ratios = [] if before in (None, 0) or after is None else [after / before]
    r = row(record, fmt, op, scenario, size, before, after, ratios, NONE,
            series, src([path]), stage=stage, comparison=comparison, unit=unit,
            explanation=explanation)
    if verdict:
        r["verdict"] = verdict
    return r


DIAGNOSTIC_NOTE = (
    "Deterministic diagnostic evidence only; it does not replace a quiet-host "
    "native A-B-B-A qualification or make a performance-acceptance claim."
)


def _metric_name(value):
    return re.sub(r"[^a-z0-9]+", "_", value.lower()).strip("_")


def _doc_toggle_rows(record, path):
    d = load_json(path)
    out = []
    callgrind = d.get("callgrind", {})
    for item in callgrind.get("rows", []):
        name = item["operation"]
        phases = []
        if "before_ir" in item:
            phases.append(("original", item["before_ir"], item["after_ir"],
                           "after/original"))
        else:
            phases.append(("original", item.get("original_ir"), item.get("candidate_ir"),
                           "candidate/original"))
            if item.get("previous_after_ir") is not None:
                phases.append(("previous after", item["previous_after_ir"],
                               item["candidate_ir"], "candidate/previous-after"))
        for phase, before, after, comparison in phases:
            if before is None or after is None:
                continue
            out.append(metric_row(
                record, "doc", "set", f"body toggle: {name}", "one toggle",
                before, after, "Ir", f"doc_body_toggle:{_metric_name(name)}:{phase}:ir",
                path, f"Callgrind Ir ({phase})", comparison,
                f"{DIAGNOSTIC_NOTE} Native status is {d.get('native', {}).get('status', 'not reported')}."))
    massif = d.get("massif", {})
    if "before" in massif:
        phases = [("original", massif["before"].get("peak_total_bytes"),
                   massif["after"].get("peak_total_bytes"), "after/original")]
    else:
        ratios = massif.get("ratios", {})
        candidate = massif.get("candidate", {}).get("peak_total_bytes")
        phases = [("original", None, candidate, "candidate/original"),
                  ("previous after", None, candidate, "candidate/previous-after")]
        # The compact candidate receipt gives the ratios but not the original
        # byte value.  Keep those ratios as ratios rather than manufacturing a
        # byte pair below.
        for phase, ratio_key in (("original", "candidate_over_original_total"),
                                 ("previous after", "candidate_over_previous_after_total")):
            if ratio_key in ratios:
                out.append(metric_row(
                    record, "doc", "set", "body toggle: ordinary route peak", "one toggle",
                    1.0, ratios[ratio_key], "ratio",
                    f"doc_body_toggle:ordinary:{phase}:massif", path,
                    f"Massif peak ratio ({phase})", "candidate/original" if phase == "original" else "candidate/previous-after",
                    f"{DIAGNOSTIC_NOTE} The receipt reports the byte values and this ratio; native qualification remains open."))
        phases = []
    for phase, before, after, comparison in phases:
        if before is None or after is None:
            continue
        out.append(metric_row(
            record, "doc", "set", "body toggle: ordinary route peak", "one toggle",
            before, after, "bytes", f"doc_body_toggle:ordinary:{phase}:massif",
            path, f"Massif peak ({phase})", comparison,
            f"{DIAGNOSTIC_NOTE} The measured heap phase is retained; no native timing claim is made."))
    if record == "0359qb":
        native = d.get("native", {})
        for block in native.get("blocks", []):
            a = block.get("a_average_us")
            b = block.get("b_average_us")
            if a is None or b is None:
                continue
            out.append(metric_row(
                record, "doc", "set", "body toggle: native diagnostic route", "one toggle",
                a, b, "µs", f"doc_body_toggle:native:{block['id']}", path,
                f"Native B/A ({block['id']}, rejected)", "B/A (native diagnostic)",
                "The native block is explicitly diagnostic_rejected because the host was busy; its ratio is retained without acceptance.",
                verdict="unqualified"))
    return out


@needs("0359qb")
def p0359qb():
    return _doc_toggle_rows("0359qb", RES / "0359qb" / "qualification.json")


@needs("0359qc")
def p0359qc():
    return _doc_toggle_rows("0359qc", RES / "0359qc" / "qualification.json")


@needs("0359qd")
def p0359qd():
    return _doc_toggle_rows("0359qd", RES / "0359qd" / "qualification.json")


@needs("0359r")
def p0359r():
    path = RES / "0359r" / "validation.json"
    d = load_json(path)
    out = []
    for name, values in d.get("layout", {}).items():
        for metric, index, unit in (("size", 0, "bytes"), ("alignment", 1, "bytes")):
            out.append(metric_row(
                "0359r", "rtf", "edit", f"packed refusal layout: {name}", "",
                values["normal"][index], values["packed"][index], unit,
                f"refusal_layout:{_metric_name(name)}:{metric}", path,
                f"Packed experiment ({metric})", "packed/normal",
                "The packed representation was a rejected temporary experiment; no implementation was landed and no native timing was qualified."))
    assembly = d["assembly"]
    for metric, unit in (("bytes", "bytes"), ("instructions", "instructions")):
        out.append(metric_row(
            "0359r", "rtf", "edit", "set_text_underline packed body", "",
            assembly["normal"][metric], assembly["packed"][metric], unit,
            f"refusal_assembly:set_text_underline:{metric}", path,
            f"Packed experiment ({metric})", "packed/normal",
            "Static assembly rejected the packed experiment: it grew the refusal producer and changed its alignment."))
    cold = d["validation"]["cold_constructor_experiment"]
    for metric, unit in (("bytes", "bytes"), ("instructions", "instructions")):
        out.append(metric_row(
            "0359r", "rtf", "edit", "set_text_underline cold refusal constructor", "",
            cold[f"normal_set_text_underline"][metric], cold[f"cold_set_text_underline"][metric], unit,
            f"refusal_assembly:cold_constructor:{metric}", path,
            f"Cold constructor experiment ({metric})", "cold/normal",
            "The cold-constructor experiment was reverted; static evidence is retained as a rejection, not as a native result."))
    return out


@needs("0359ra")
def p0359ra():
    path = RES / "0359ra" / "callgrind-summary.json"
    d = load_json(path)
    out = []
    for name, before, after in [("whole profile", d["profile"]["before_ir"], d["profile"]["after_ir"])] + [
            (item["name"], item["before_ir"], item["after_ir"])
            for item in d["profile"].get("function_costs", [])]:
        out.append(metric_row(
            "0359ra", "rtf", "edit", f"refusal contract: {name}", "34-sample probe",
            before, after, "Ir", f"refusal_contract:{_metric_name(name)}:callgrind",
            path, "Callgrind before/after", "after/before (deterministic diagnostic)",
            "Callgrind measurement only; record 0359ra does not qualify the native timing result or accept an implementation."))
    return out


@needs("0359rb")
def p0359rb():
    path = RES / "0359rb" / "callgrind-summary.json"
    d = load_json(path)
    out = []
    def add(scope, name, before, after, stage):
        out.append(metric_row(
            "0359rb", "rtf", "edit", f"{scope}: {name}", "34-sample probe",
            before, after, "Ir", f"rtf_fast_path:{scope}:{_metric_name(name)}",
            path, stage, "fast/current (deterministic diagnostic)",
            "The candidate remains native-unqualified; direct instruction controls are retained without a timing or acceptance claim."))
    for name in ("whole_ir", "set_text_underline_ir"):
        add("refusal route", name, d["refusal_route"]["current"][name],
            d["refusal_route"]["fast"][name], "Refusal route Callgrind")
    for name, values in d["normal_route_direct_ir"].items():
        add("normal route", name, values["current"], values["fast"], "Normal-route direct Ir")
    for phase in ("cold", "warm"):
        values = d["successful_route_controls"][phase]
        add(f"successful {phase} control", "whole_ir", values[f"whole_ir_current"],
            values["whole_ir_fast"], f"Successful {phase} control Callgrind")
    return out


@needs("0359rc")
def p0359rc():
    path = RES / "0359rc" / "callgrind-summary.json"
    d = load_json(path)
    out = []
    def add(scope, name, before, after, stage):
        if not isinstance(before, (int, float)) or not isinstance(after, (int, float)):
            return
        out.append(metric_row(
            "0359rc", "rtf", "edit", f"{scope}: {name}", "34-sample probe",
            before, after, "Ir", f"rtf_attribution:{scope}:{_metric_name(name)}",
            path, stage, "candidate/current (deterministic diagnostic)",
            "Call-inclusive attribution is diagnostic; the candidate remains unaccepted pending isolated native timing and heaptrack."))
    for scope, block_name, stage, keys in (("cold successful route", "cold_success_inclusive", "Cold Callgrind", ("whole_ir", "main_ir", "main_to_parse_internal", "main_to_success", "success_to_commit_staged", "tree_parse_internal", "parser_parse", "raw_vec_finish_grow")),
                                            ("warm successful route", "warm_success_inclusive", "Warm Callgrind", ("whole_ir", "main_ir", "main_to_success", "success_to_commit_staged")),
                                            ("refusal route", "refusal_inclusive", "Refusal Callgrind", ("whole_ir", "main_ir", "set_text_underline", "commit_staged"))):
        block = d[block_name]
        for name in keys:
            add(scope, name, block["current"][name], block["candidate"][name], stage)
    for name, values in d["normal_direct_ir"].items():
        add("normal direct control", name, values["current"], values["candidate"], "Normal direct control")
    return out


def _median_value(value):
    return med(value) if isinstance(value, list) else value


def _odp_reference_rows(record, path, candidate, reference, comparison):
    out = []
    routes = [route for route in candidate if route in reference]
    op_for = {"open": "open", "open_first_slide": "open", "open_first_text": "open",
              "first_slide": "get", "repeat_slide": "get", "first_text": "get",
              "repeat_text": "get", "write_text": "save", "repeat_write_text": "save",
              "unchanged_write": "save", "append_commit_write": "add", "replay": "replay",
              "replay_write": "replay"}
    for route in routes:
        for metric in ("peak", "allocated"):
            before = _median_value(reference[route].get(metric))
            after = _median_value(candidate[route].get(metric))
            out.append(metric_row(
                record, "odp", op_for.get(route, "read"), f"source owner {route}", "2,000 slides",
                before, after, "bytes", f"odp_source_owner:{route}:{metric}:{comparison}", path,
                f"Allocator {metric}", comparison,
                f"{DIAGNOSTIC_NOTE} The source-owner comparison is retained for attribution; native timing remains open."))
    return out


@needs("0362ya")
def p0362ya():
    path = RES / "0362ya" / "validation.json"
    d = load_json(path)
    return _odp_reference_rows("0362ya", path, d["allocation_probe"]["candidate"],
                               d["allocation_probe"]["eager_arc_string_reference"],
                               "candidate/eager reference")


@needs("0362yb")
def p0362yb():
    path = RES / "0362yb" / "validation.json"
    d = load_json(path)
    return _odp_reference_rows("0362yb", path, d["allocator_probe"],
                               d["paired_eager_arc_string_reference"],
                               "streamed/eager reference")


@needs("0362yc")
def p0362yc():
    path = RES / "0362yc" / "validation.json"
    d = load_json(path)
    return _odp_reference_rows("0362yc", path, d["allocator_probe"],
                               d["paired_eager_arc_string_reference"],
                               "corrected/eager reference")


@needs("0362yd")
def p0362yd():
    path = RES / "0362yd" / "validation.json"
    d = load_json(path)
    out = []
    op_for = {"open": "open", "first_slide": "get", "repeat_slide": "get",
              "first_text": "get", "repeat_text": "get", "write_text": "save",
              "repeat_write_text": "save", "unchanged_write": "save",
              "append_commit_write": "add", "replay": "replay"}
    for item in d.get("rows", []):
        route = item["route"]
        for metric in ("peak", "allocated"):
            out.append(metric_row(
                "0362yd", "odp", op_for.get(route, "read"), f"matched source owner {route}", "2,000 slides",
                item["before"][metric], item["after"][metric], "bytes",
                f"odp_matched_owner:{route}:{metric}", path, f"Allocator {metric}",
                "after/before", f"{DIAGNOSTIC_NOTE} Output equivalence passed; native timing was not run."))
        ir = item.get("instruction_references", {})
        out.append(metric_row(
            "0362yd", "odp", op_for.get(route, "read"), f"matched source owner {route}", "2,000 slides",
            ir.get("before"), ir.get("after"), "Ir", f"odp_matched_owner:{route}:ir", path,
            "Callgrind Ir", "after/before", f"{DIAGNOSTIC_NOTE} Instruction evidence is diagnostic, not native acceptance."))
    return out


def _markdown_table(path, header_prefix):
    lines = path.read_text(encoding="utf-8").splitlines()
    for i, line in enumerate(lines):
        if line.startswith(header_prefix):
            j = i + 2
            rows = []
            while j < len(lines) and lines[j].startswith("|"):
                cells = [cell.strip() for cell in lines[j].strip("|").split("|")]
                if cells and cells[0] != "":
                    rows.append(cells)
                j += 1
            return rows
    return []


@needs("0362ye")
def p0362ye():
    path = RES / "0362ye" / "allocator-fixed-probe.md"
    out = []
    op_for = {"open": "open", "first_slide": "get", "repeat_slide": "get",
              "first_text": "get", "repeat_text": "get", "write_text": "save",
              "repeat_write_text": "save", "unchanged_write": "save",
              "append_commit_write": "add", "replay": "replay"}
    for cells in _markdown_table(path, "| route | e4a baseline peak"):
        route = cells[0].strip("`")
        peak_before, allocated_before = (int(value) for value in cells[1].split("/"))
        peak_after, allocated_after = (int(value) for value in cells[2].split("/"))
        for metric, before, after in (("peak", peak_before, peak_after),
                                      ("allocated", allocated_before, allocated_after)):
            out.append(metric_row(
                "0362ye", "odp", op_for.get(route, "read"), f"fixed admission {route}", "2,000 slides",
                before, after, "bytes", f"odp_fixed_admission:{route}:{metric}", path,
                f"Allocator {metric}", "candidate/baseline", f"{DIAGNOSTIC_NOTE} This is allocator evidence; native timing remains pending."))
    for cells in _markdown_table(path, "| route | baseline Ir"):
        route = cells[0].strip("`")
        before, after = (int(cells[i].replace(",", "")) for i in (1, 2))
        out.append(metric_row(
            "0362ye", "odp", op_for.get(route, "read"), f"fixed admission {route}", "2,000 slides",
            before, after, "Ir", f"odp_fixed_admission:{route}:ir", path,
            "Callgrind Ir", "candidate/baseline", f"{DIAGNOSTIC_NOTE} Route-local instructions are diagnostic only."))
    for cells in _markdown_table(path, "| route | candidate / baseline"):
        route = cells[0].strip("`")
        ratio = float(cells[1].rstrip("×"))
        out.append(metric_row(
            "0362ye", "odp", op_for.get(route, "read"), f"fixed admission {route}", "2,000 slides",
            1.0, ratio, "ratio", f"odp_fixed_admission:{route}:native", path,
            "Native ABBA (diagnostic, rejected)", "candidate/baseline (native diagnostic)",
            "The pinned ABBA was explicitly diagnostic only under host load; it is not a quiet-host native qualification.",
            verdict="unqualified"))
    return out


@needs("0359rd")
def p0359rd():
    path = RES / "0359rd" / "qualification-summary.json"
    d = load_json(path)
    out = []
    heap = d["heaptrack"]
    for left, right, scope, comparison in (("original", "candidate", "original/candidate heap", "candidate/original"),
                                            ("current", "candidate", "current/candidate heap", "candidate/current")):
        for metric, unit in (("allocations", "count"), ("temporary", "count"),
                             ("peak_heap", "bytes"), ("leaked", "bytes")):
            out.append(metric_row(
                "0359rd", "rtf", "replay", f"{scope}: {metric}", "34-sample probe",
                _median_value(heap[left][metric]), _median_value(heap[right][metric]), unit,
                f"rtf_replay_heap:{left}:{right}:{metric}", path, "Heaptrack diagnostic",
                comparison, "Heap attribution was collected under host contamination and is explicitly rejected for native qualification."))
    return out


@needs("0359re")
def p0359re():
    path = RES / "0359re" / "qualification-summary.json"
    d = load_json(path)
    out = []
    heap = d["heaptrack"]
    for metric, unit in (("allocations", "count"), ("temporary", "count"),
                         ("peak_heap", "bytes"), ("leaked", "bytes")):
        if metric == "allocations":
            before, after = (_median_value(heap["original_allocations"]),
                             _median_value(heap["final_allocations"]))
            comparison, stage = "final/original", "Heaptrack paired"
            explanation = "Final-source allocation attribution is retained; compiler activity rejected it for native qualification."
        else:
            # The compact receipt does not label the temporary/peak/leaked
            # vectors by leg.  Keep the reported final absolute value, rather
            # than manufacturing a ratio against the same aggregate.
            before, after = None, _median_value(heap[metric])
            comparison, stage = "final absolute", "Heaptrack final-only"
            explanation = "Final-source heaptrack value only; the receipt has no labelled original counterpart, so no ratio or native qualification is inferred."
        out.append(metric_row(
            "0359re", "rtf", "replay", f"final shared replay storage: {metric}", "34-sample probe",
            before, after, unit, f"rtf_replay_shared:{metric}", path, stage,
            comparison, explanation))
    cg = d["callgrind"]
    for name, before, after in (("whole profile", med(cg["original_ir"]), med(cg["final_ir"])),
                                ("Package::apply", cg["old_api_package_apply_inclusive_ir"], cg["final_package_apply_inclusive_ir"]),
                                ("Commit::new", cg["old_api_commit_new_ir"], cg["final_commit_new_ir"]),
                                ("Commit drop", cg["old_api_commit_drop_ir"], cg["final_commit_drop_ir"]),
                                ("Arc drop", cg["old_api_arc_drop_ir"], cg["final_arc_drop_ir"])):
        out.append(metric_row(
            "0359re", "rtf", "replay", f"final shared replay storage: {name}", "34-sample probe",
            before, after, "Ir", f"rtf_replay_shared:{_metric_name(name)}:ir", path,
            "Callgrind diagnostic", "final/original", "Final-source instruction attribution is diagnostic and does not replace native timing."))
    return out


@needs("0359qe")
def p0359qe():
    """DOC body-toggle evidence with its explicit overlay baseline.

    0359qe is a controlled source-overlay comparison, not a leaf/facade
    operation pair.  Keep its instruction and heap observations as
    diagnostic, no-budget rows, and retain the busy-host wall observations as
    explicitly unqualified native diagnostics.  In particular, do not let a
    favorable deterministic or host-contaminated value turn into a native
    acceptance row.
    """
    path = RES / "0359qe" / "qualification.json"
    d = load_json(path)
    out = []
    note = (
        f"{DIAGNOSTIC_NOTE} The controlled baseline is the 3cad source with "
        "only the recorded body-publication overlay reversed; this is not a "
        "leaf/facade comparison."
    )
    # Every Ir value in this receipt is a labelled candidate/baseline pair.
    # Read only those explicit pairs; call counts or unrelated metadata do not
    # become synthetic measurements.
    expected_callgrind = {
        f"{route}_{metric}_ir"
        for route in ("ordinary", "batch")
        for metric in ("route", "edit_commit", "set_character_property_override", "write_to")
    }
    found_callgrind = set()
    for name, pair in d["callgrind"].items():
        if not name.endswith("_ir") or not isinstance(pair, dict):
            continue
        candidate = pair.get("candidate")
        baseline = pair.get("baseline")
        if not isinstance(candidate, (int, float)) or not isinstance(baseline, (int, float)):
            continue
        route = "batch" if name.startswith("batch_") else "ordinary"
        scenario = "body toggle batch route" if route == "batch" else "body toggle ordinary route"
        size = "1,000 toggles" if route == "batch" else "one toggle"
        metric = name.removeprefix(f"{route}_").removesuffix("_ir")
        found_callgrind.add(name)
        out.append(metric_row(
            "0359qe", "doc", "batch" if route == "batch" else "set",
            scenario, size, baseline, candidate, "Ir",
            f"doc_body_toggle:{route}:{_metric_name(metric)}:callgrind",
            path, f"Callgrind Ir ({route}, candidate/baseline)",
            "candidate/baseline", note))
    if found_callgrind != expected_callgrind:
        missing = sorted(expected_callgrind - found_callgrind)
        extra = sorted(found_callgrind - expected_callgrind)
        raise SystemExit(f"0359qe callgrind pairs changed; missing={missing}, extra={extra}")

    for route, scenario, size in (
        ("ordinary", "body toggle ordinary route", "one toggle"),
        ("batch", "body toggle batch route", "1,000 toggles"),
    ):
        metric = d["massif"][f"{route}_peak_total_bytes"]
        candidate = metric["candidate"]
        baseline = metric["baseline"]
        out.append(metric_row(
            "0359qe", "doc", "batch" if route == "batch" else "set",
            scenario, size, baseline, candidate, "bytes",
            f"doc_body_toggle:{route}:massif_peak", path,
            f"Massif peak ({route}, candidate/baseline)",
            "candidate/baseline", note))

    for route, scenario, size in (
        ("ordinary", "body toggle ordinary route", "one toggle"),
        ("batch", "body toggle batch route", "1,000 toggles"),
    ):
        metric = d["wall_diagnostic"][f"{route}1k_leaf_write_us_median" if route == "batch" else "ordinary_leaf_write_us_median"]
        candidate = metric["candidate"]
        baseline = metric["baseline"]
        out.append(metric_row(
            "0359qe", "doc", "batch" if route == "batch" else "set",
            scenario, size, baseline, candidate, "µs",
            f"doc_body_toggle:{route}:wall_diagnostic", path,
            f"Native wall ({route}, diagnostic rejected)",
            "candidate/baseline", (
                f"{note} The receipt marks this wall result diagnostic_only "
                "under host load; quiet-host native timing remains pending."
            ), verdict="unqualified"))
    return out


# Record 0377: performance qualification of the API-debt program against the
# pre-program trunk 839212fa9d (baseline) on the merged candidate.
P0377_DEFERRED = ("deferred by owner (owner decision 2026-10-26: ODT/ODP performance improvement "
                  "is deferred; measured and reported, not fixed; record 0377 §4). The ODT "
                  "engine's raw-XML admission (ResolvedReader, record 0362ul44) resolves every "
                  "event where the baseline's gated BindingTracker resolved only shallow ones.")
P0377_METADATA = ("explained residual (record 0377 §5.2): the W8 metadata commit (record 0362o) "
                  "inspects the strict graph of the core, app and custom families and builds an "
                  "exact reversible patch with an authenticated scope (contract 0351 §2.3) — a "
                  "fixed 60–190k Ir (≈1–3 µs) per commit; follow-up 0377-F2: read only the "
                  "changed family's payload.")
P0377_XLSX_BOUND = ("explained residual (record 0377 §5.3): contract 0351 §2.5 snapshot bound — "
                    "the unified edit is bounded, the 0337 leaf leg is not; after 0377 removed the "
                    "commit-time ledger rebuilds and the duplicate cell reservation, one ledger "
                    "admission per staged cell remains; follow-up 0377-F3: a flat per-sheet "
                    "ledger.")
P0377_XLSX_LEAF = ("explained residual (record 0377 §5.3): the leaf XLSX cell edit also carries "
                   "the W4 typed-cell and capture plumbing (1.062 in instructions, the "
                   "remainder of the wall ratio is the codegen residual of §5.5); follow-up "
                   "0377-F3.")
P0377_CODEGEN = ("explained residual of code generation (record 0377 §5.5): instructions are "
                 "within 5 % (callgrind); the wall difference is instruction-cache placement of "
                 "the fat-LTO 16-codegen-unit release build — the same perf-baseline cases built "
                 "with codegen-units=1 measure 0.995–1.011 (results/0377/perf-baseline-rerun.tsv).")
P0377_ODC_XML = ("explained residual (record 0377 §5.4): raw-XML admission of record 0362ul44 "
                 "(duplicate-attribute and UTF-8 checks, per-event namespace resolution) on the "
                 "odfdom fixture's many namespace declarations; 0377 added the namespace "
                 "normalization fast path (1.162 → 1.080); follow-up 0377-F4.")


def p0377_budget(group, scenario):
    if group == "batch":
        return BATCH if "batch1000" in scenario else SCALING
    if scenario.endswith("_open") or "_open_" in scenario or scenario.endswith("open_validated"):
        return OPEN
    if scenario.endswith("_get"):
        return READ_POINT
    return SINGLE


def p0377_op(route):
    for key, op in (("batch", "batch"), ("_open", "open"), ("save_", "save"), ("_get", "get"),
                    ("_add", "add"), ("remove", "remove"), ("move", "move")):
        if key in route:
            return op
    return "set"


def p0377_explain(route, kind):
    fmt = route.split("_", 1)[0]
    if fmt in ("odt", "oth", "odm"):
        return P0377_DEFERRED
    if "metadata_title_set" in route:
        return P0377_METADATA
    if route.startswith("xlsx_eager_cell_value_set"):
        return P0377_XLSX_LEAF if kind == "leaf" else P0377_XLSX_BOUND
    if fmt == "odc":
        return P0377_ODC_XML
    return P0377_CODEGEN if kind.startswith("wall") else ""


@needs("0375b")
def p0375b():
    """Historical full-facade chart-text instruction comparisons, not File timing."""
    path = RES / "0375b" / "integrated-callgrind.json"
    data = load_json(path)
    out = []
    seen = set()
    for route in data["routes"]:
        test = route["test"]
        fmt = test.split("_", 1)[0]
        operation = test.removesuffix("_cost_abba").rsplit("_", 1)[-1]
        host, leaf = route["host_ir"], route["leaf_ir"]
        if (test in seen or fmt not in ("numbers", "pages")
                or operation not in ("read", "set", "clear")
                or type(host) is not int or type(leaf) is not int
                or host <= 0 or leaf <= 0):
            raise SystemExit(f"0375b invalid or duplicate instruction route: {test}")
        seen.add(test)
        out.append(row(
            "0375b", fmt, "get" if operation == "read" else "set",
            test, "", host, leaf, [leaf / host], regression("0363"),
            f"chart_text:{test}:ir", src([path]),
            stage=f"integrated {data['source_commit'][:10]}",
            comparison="leaf/host (Callgrind Ir)", unit="Ir",
            explanation=(
                "Historical source-bound full-facade instruction comparison; "
                "baseline column is the legacy host, candidate column is the leaf. "
                "No wall-time, unified File, native acceptance, or current-trunk "
                "qualification is inferred.")))
    return out


@needs("0391")
def p0391():
    """Accepted scoped W1-G whole-cell Numbers instruction routes.

    The 595 direct semantic reads in the record are compatibility proof, not
    cost rows.  Only the three dedicated whole-cell routes below contribute
    instruction rows, and they remain source-bound until the coordinator
    requalifies the integrated trunk.
    """
    path = RES / "0391" / "whole-cell-text-style.json"
    data = load_json(path)
    source = data.get("source_commit")
    if (not isinstance(source, str) or not re.fullmatch(r"[0-9a-f]{40}", source)
            or data.get("source_clean") is not True):
        raise SystemExit("0391 invalid source identity")
    proof = data.get("compatibility_proof", {})
    if proof.get("direct_semantic_reads") != 595 or proof.get("nature") != (
            "Compatibility/read correctness coverage, not an instruction-cost route"):
        raise SystemExit("0391 compatibility proof is not the recorded 595-read scope")
    policy = data.get("policy", {})
    if policy.get("ratio") != "leaf/host inclusive Callgrind Ir" or policy.get("max_ratio") != 1.05:
        raise SystemExit("0391 instruction policy changed")
    routes = data.get("routes")
    if not isinstance(routes, list) or len(routes) != 3:
        raise SystemExit("0391 requires exactly three whole-cell routes")
    expected = {
        "w1g_cell_text_style_single": ("set", "1 cell"),
        "w1g_cell_text_style_exclusive_repeat": ("set", "1 cell repeated"),
        "w1g_cell_text_style_batch_1000": ("batch", "50x20 / 1,000 cells"),
    }
    out = []
    seen = set()
    for route in routes:
        route_id = route.get("id")
        if route_id in seen or route_id not in expected:
            raise SystemExit(f"0391 invalid or duplicate route: {route_id}")
        seen.add(route_id)
        operation, size = expected[route_id]
        if (route.get("operation"), route.get("size")) != (operation, size):
            raise SystemExit(f"0391 route shape changed: {route_id}")
        host = route.get("host_ir")
        leaf = route.get("leaf_ir")
        calls = route.get("calls_per_side")
        ratio = route.get("ratio")
        if (type(host) is not int or type(leaf) is not int or host <= 0 or leaf <= 0
                or calls != 2 or type(ratio) not in (int, float) or ratio <= 0
                or abs(ratio - leaf / host) > 1e-12
                or route.get("status") != "meets"):
            raise SystemExit(f"0391 invalid instruction result: {route_id}")
        out.append(row(
            "0391", "numbers", operation, route["scenario"], size,
            leaf, host, [leaf / host], regression("0391"),
            f"w1g_cell_text_style:{route_id}:ir", src([path]),
            stage=f"source-bound {source[:10]}",
            comparison="leaf/host (Callgrind Ir)", unit="Ir",
            explanation=(
                "Accepted scoped Numbers whole-cell route on clean source "
                f"{source[:10]}; final integrated-trunk requalification and "
                "native acceptance remain pending.")))
    if seen != set(expected):
        raise SystemExit("0391 route set is incomplete")
    return out


@needs("0399")
def p0399():
    """Current-source D7 full-file chart-font-size qualification.

    The reduced receipt retains one row per File/direct-leaf route.  Raw
    captures and the immutable binary remain external; the receipt is only
    the compact registry projection of the independently validated matrix.
    """
    path = RES / "0399" / "integrated-callgrind.json"
    data = load_json(path)
    source = data.get("source_commit")
    if (not isinstance(source, str) or not re.fullmatch(r"[0-9a-f]{40}", source)
            or data.get("source_clean") is not True):
        raise SystemExit("0399 invalid source identity")
    policy = data.get("policy", {})
    if (policy.get("ratio") != "unified/leaf inclusive Callgrind Ir"
            or policy.get("max_ratio") != 1.05
            or policy.get("cpu") != 5
            or policy.get("calls_per_side") != 2
            or policy.get("wall_time_claim") is not False):
        raise SystemExit("0399 instruction policy changed")
    routes = data.get("routes")
    if not isinstance(routes, list) or len(routes) != 9:
        raise SystemExit("0399 requires exactly nine chart-font-size routes")
    expected = {
        "pages": {"read", "set", "reset"},
        "keynote": {"read", "set", "reset"},
        "numbers": {"read", "set", "reset"},
    }
    out = []
    seen = set()
    for route in routes:
        route_id = route.get("id")
        fmt = route.get("format")
        operation = route.get("operation")
        if (route_id in seen or not isinstance(route_id, str)
                or fmt not in expected or operation not in expected[fmt]):
            raise SystemExit(f"0399 invalid or duplicate route: {route_id}")
        seen.add(route_id)
        unified = route.get("unified_ir")
        leaf = route.get("leaf_ir")
        ratio = route.get("ratio")
        if (type(unified) is not int or type(leaf) is not int
                or unified <= 0 or leaf <= 0
                or type(ratio) not in (int, float)
                or abs(ratio - unified / leaf) > 1e-12
                or ratio > policy["max_ratio"]
                or route.get("calls_per_side") != policy["calls_per_side"]
                or route.get("status") != "meets"):
            raise SystemExit(f"0399 invalid instruction result: {route_id}")
        op = "get" if operation == "read" else "set"
        out.append(row(
            "0399", fmt, op, route["scenario"], route.get("size", ""),
            leaf, unified, [ratio], regression("0399"),
            f"chart_font_size_full_file:{route_id}:ir", src([path]),
            stage=f"source-bound {source[:10]}",
            comparison="unified/leaf (inclusive Callgrind Ir)", unit="Ir",
            explanation=(
                "Accepted current-source full-file/direct-leaf instruction route; "
                "raw selected calls, full-function Ir, PROGRAM TOTALS and the "
                "threshold-100 aggregate annotation agree. No wall-time or native "
                "acceptance claim.")))
    if seen != {route["id"] for route in routes} or len(seen) != 9:
        raise SystemExit("0399 route set is incomplete")
    return out


@needs("0403")
def p0403():
    """Current-source Numbers ``Property::Symbols`` facade instruction routes.

    The four rows are black-box host-vs-leaf facade routes.  The raw
    Callgrind profiles and annotations remain outside the repository; this
    reduced receipt is the benchmark-matrix projection of the independently
    validated two-call-per-side capture.
    """
    path = RES / "0403" / "integrated-callgrind.json"
    data = load_json(path)
    source = data.get("source_commit")
    if (not isinstance(source, str) or not re.fullmatch(r"[0-9a-f]{40}", source)
            or data.get("source_clean") is not True):
        raise SystemExit("0403 invalid source identity")
    policy = data.get("policy", {})
    if (policy.get("ratio") != "leaf/host inclusive Callgrind Ir"
            or policy.get("max_ratio") != 1.05
            or policy.get("cpu") != 7
            or policy.get("calls_per_side") != 2
            or policy.get("wall_time_claim") is not False
            or policy.get("annotation") != (
                "callgrind_annotate --auto=no --inclusive=yes --threshold=100 "
                "--show-percs=yes")):
        raise SystemExit("0403 instruction policy changed")
    routes = data.get("routes")
    expected = {
        "numbers_symbols_all_read": ("read", "Property::Symbols all chart symbols", "all"),
        "numbers_symbols_all_set": ("set", "Property::Symbols all chart symbols", "all"),
        "numbers_symbols_single_read": ("read", "Property::Symbols single chart symbol", "1"),
        "numbers_symbols_single_set": ("set", "Property::Symbols single chart symbol", "1"),
    }
    if not isinstance(routes, list) or len(routes) != len(expected):
        raise SystemExit("0403 requires exactly four Numbers Symbols routes")
    out = []
    seen = set()
    for route in routes:
        route_id = route.get("id")
        if route_id in seen or route_id not in expected:
            raise SystemExit(f"0403 invalid or duplicate route: {route_id}")
        seen.add(route_id)
        operation, scenario, size = expected[route_id]
        if (route.get("operation"), route.get("scenario"), route.get("size")) != (
                operation, scenario, size):
            raise SystemExit(f"0403 route shape changed: {route_id}")
        host = route.get("host_ir")
        leaf = route.get("leaf_ir")
        ratio = route.get("ratio")
        if (type(host) is not int or type(leaf) is not int
                or host <= 0 or leaf <= 0
                or type(ratio) not in (int, float)
                or abs(ratio - leaf / host) > 1e-12
                or ratio > policy["max_ratio"]
                or route.get("calls_per_side") != policy["calls_per_side"]
                or route.get("status") != "meets"):
            raise SystemExit(f"0403 invalid instruction result: {route_id}")
        symbols = route.get("symbols")
        if (not isinstance(symbols, dict)
                or not isinstance(symbols.get("host"), str)
                or not isinstance(symbols.get("leaf"), str)):
            raise SystemExit(f"0403 missing selected symbols: {route_id}")
        out.append(row(
            "0403", "numbers", "get" if operation == "read" else "set",
            scenario, size, leaf, host, [ratio], regression("0403"),
            f"numbers_symbols_facade:{route_id}:ir", src([path]),
            stage=f"source-bound {source[:10]}",
            comparison="leaf/host (inclusive Callgrind Ir)", unit="Ir",
            explanation=(
                "Accepted current-source Numbers Property::Symbols facade route; "
                "raw selected calls, exact two-call policy, full-function Ir, "
                "PROGRAM TOTALS and threshold-100 aggregate annotation agree. "
                "No wall-time or native acceptance claim.")))
    if seen != set(expected):
        raise SystemExit("0403 route set is incomplete")
    return out


@needs("0377")
def p0377():
    """Record 0377: unified-fin-perf (one source built on the baseline with
    `--features base` and on the candidate), callgrind instructions of the
    timed region per call (primary evidence for the 5 % regression rule) and
    wall-clock unified/leaf ABBA blocks (the leaf-vs-unified budgets);
    perf-baseline default matrix (instructions and wall); odf_w6_probe;
    odf_source_backed_probe (ODP memory); the Python binding."""
    res = RES / "0377"
    out = []
    path = res / "callgrind.tsv"
    with open(path, encoding="utf-8") as f:
        for r in csv.DictReader(f, delimiter="\t"):
            route = r["route"]
            if route == "docx_eager_open":
                continue  # unpaired: see docx_eager_open_validated
            fmt = route.split("_", 1)[0]
            op = p0377_op(route)
            vals = {k: (float(r[k]) if r[k] else None) for k in
                    ("leaf_base_ir", "leaf_cand_ir", "unified_base_ir", "unified_cand_ir")}
            for leg in ("leaf", "unified"):
                b, c = vals[f"{leg}_base_ir"], vals[f"{leg}_cand_ir"]
                if b and c:
                    out.append(row("0377", fmt, op, f"{route} ({leg} leg, candidate vs pre-program baseline)",
                                   "", b, c, [c / b], regression("0377"),
                                   f"fin_perf:{route}:{leg}:after-before", src([path]),
                                   comparison="after/before (callgrind Ir)", unit="Ir",
                                   explanation=p0377_explain(route, leg) if c / b > 1.05 else ""))
            if vals["leaf_cand_ir"] and vals["unified_cand_ir"]:
                out.append(row("0377", fmt, op, f"{route} (instructions)", "",
                               vals["leaf_cand_ir"], vals["unified_cand_ir"],
                               [vals["unified_cand_ir"] / vals["leaf_cand_ir"]], NONE,
                               f"fin_perf:{route}:ir", src([path]), unit="Ir",
                               comparison="unified/leaf (callgrind Ir; budgets judged on wall)"))
    path = res / "abba.json"
    data: dict = {}
    for run in load_json(path)["runs"]:
        for c in run["cases"]:
            data.setdefault(c["route"], {"base": [], "cand": []})[run["build"]].append(c)
    for route, v in data.items():
        if not v["cand"]:
            continue
        group = v["cand"][0]["group"]
        fmt = route.split("_", 1)[0]
        leaf = med(x for c in v["cand"] for x in c["leaf_block_median_ns"])
        uni = med(x for c in v["cand"] for x in c["unified_block_median_ns"])
        ratios = [u / l for c in v["cand"] for u, l in
                  zip(c["unified_block_median_ns"], c["leaf_block_median_ns"])]
        budget = p0377_budget(group, route)
        out.append(row("0377", fmt, p0377_op(route), f"{route} (wall)", "",
                       us(leaf), us(uni), [med(ratios)], budget, f"fin_perf:{route}:wall",
                       src([path]), comparison="unified/leaf (wall, ABBA blocks)",
                       explanation=p0377_explain(route, "wall") if (
                           budget[0] and med(ratios) > budget[0]) else ""))
    path = res / "odf-w6-callgrind.tsv"
    with open(path, encoding="utf-8") as f:
        for r in csv.DictReader(f, delimiter="\t"):
            scen = r["scenario"]
            fmt, op = scen.split("/")
            b, c = float(r["base_ir_per_round"]), float(r["candidate_ir_per_round"])
            expl = ""
            if c / b > 1.05:
                expl = P0377_DEFERRED if fmt in ("odt", "fodt") else (
                    P0377_ODC_XML if fmt == "odc" else "")
            out.append(row("0377", fmt, W6_ODF_OPS.get(op, op),
                           f"odf_w6_probe {scen} (unified, candidate vs pre-program baseline)",
                           "", b, c, [c / b], regression("0377"), f"odf_w6:{scen}:after-before",
                           src([path]), comparison="after/before (callgrind Ir per round)",
                           unit="Ir", explanation=expl))
    path = res / "odf-sb-wall.json"
    for r in load_json(path)["rows"]:
        if "base_peak_heap_bytes" not in r:
            continue
        fmt, scen = r["scenario"].split("/")
        b, c = med(r["base_peak_heap_bytes"]), med(r["candidate_peak_heap_bytes"])
        out.append(row("0377", fmt, "set" if fmt == "ods" else "add",
                       f"odf_source_backed_probe {r['scenario']} peak heap", "", b, c, [c / b],
                       regression("0377"), f"odf_sb:{r['scenario']}:peak-heap", src([path]),
                       comparison="after/before (peak live heap)", unit="bytes"))
    path = res / "perf-baseline-callgrind.tsv"
    with open(path, encoding="utf-8") as f:
        for r in csv.DictReader(f, delimiter="\t"):
            case = r["case"]
            b, c = float(r["base_ir_per_sample"]), float(r["candidate_ir_per_sample"])
            fmt = case.split("_", 1)[0]
            expl = ""
            if c / b > 1.05:
                expl = P0377_DEFERRED if case.startswith("odp_") else (
                    "explained residual (record 0377 §5.5): the per-sample instruction count "
                    "includes untimed per-sample setup; the timed wall interval is 0.93–0.97 "
                    "of the baseline (perf-baseline-wall.tsv)")
            out.append(row("0377", fmt, "default-matrix",
                           f"perf-baseline {case} (all corpora, per sample)", "", b, c, [c / b],
                           regression("0377"), f"perf_baseline:{case}:ir", src([path]),
                           comparison="after/before (callgrind Ir per sample)", unit="Ir",
                           explanation=expl))
    path = res / "python.json"
    for r in load_json(path)["rows"]:
        b, c = med(r["base_median_us_per_run"]), med(r["candidate_median_us_per_run"])
        out.append(row("0377", "python", r["scenario"].split()[1] if " " in r["scenario"] else "",
                       f"Python {r['scenario']}", "", b, c, [c / b], NONE,
                       f"python:{slug(r['scenario'])}:after-before", src([path]),
                       comparison="after/before (wall, A-B-B-A processes)"))
    return out


@needs("0377a")
def p0377a():
    """Record 0377a: bounded follow-up remeasurement of F1--F5.

    The callgrind file contains corrected warm-up counts for the timed
    ``fin_perf_region``.  The ABBA receipt contains the fresh leaf/unified
    wall legs for F1, F3 and F5, and the CLI receipt compares the four
    process routes against the same baseline build.
    """
    res = RES / "0377a"
    out = []

    path = res / "callgrind.tsv"
    with open(path, encoding="utf-8") as f:
        for r in csv.DictReader(f, delimiter="\t"):
            route = r["route"]
            fmt = route.split("_", 1)[0]
            op = p0377_op(route)
            vals = {k: float(r[k]) for k in (
                "leaf_base_ir", "leaf_cand_ir", "unified_base_ir", "unified_cand_ir")}
            for leg in ("leaf", "unified"):
                before, after = vals[f"{leg}_base_ir"], vals[f"{leg}_cand_ir"]
                out.append(row(
                    "0377a", fmt, op,
                    f"{route} ({leg} leg, candidate vs pre-program baseline)", "",
                    before, after, [after / before], regression("0377a"),
                    f"fin_perf_0377a:{route}:{leg}:after-before", src([path]),
                    comparison="after/before (callgrind Ir)", unit="Ir",
                    explanation="Corrected-warmup callgrind pair; timed region only."))
            out.append(row(
                "0377a", fmt, op, f"{route} (instructions)", "",
                vals["leaf_cand_ir"], vals["unified_cand_ir"],
                [vals["unified_cand_ir"] / vals["leaf_cand_ir"]], NONE,
                f"fin_perf_0377a:{route}:ir", src([path]), unit="Ir",
                comparison="unified/leaf (callgrind Ir; budgets judged on wall)",
                explanation="Corrected-warmup callgrind diagnostic; unified/leaf is not a wall budget."))

    path = res / "abba.json"
    data: dict = {}
    for run in load_json(path)["runs"]:
        for c in run["cases"]:
            data.setdefault(c["route"], {"base": [], "cand": []})[run["build"]].append(c)
    for route, builds in data.items():
        if not builds["cand"]:
            continue
        case = builds["cand"][0]
        group = case["group"]
        fmt = route.split("_", 1)[0]
        leaf = med(v for c in builds["cand"] for v in c["leaf_block_median_ns"])
        unified = med(v for c in builds["cand"] for v in c["unified_block_median_ns"])
        ratios = [u / l for c in builds["cand"] for u, l in zip(
            c["unified_block_median_ns"], c["leaf_block_median_ns"])]
        budget = p0377_budget(group, route)
        out.append(row(
            "0377a", fmt, p0377_op(route), f"{route} (wall)", "",
            us(leaf), us(unified), [med(ratios)], budget,
            f"fin_perf_0377a:{route}:wall", src([path]),
            comparison="unified/leaf (wall, ABBA blocks)",
            explanation=(
                "Fresh corrected-warmup ABBA diagnostic under shared-host load; "
                "provisional pending a quiet coordinator rerun; every case "
                "reports exact byte equivalence." if case.get("equivalence") == "exact_bytes" else
                "Fresh corrected-warmup ABBA diagnostic under shared-host load; "
                "provisional pending a quiet coordinator rerun; ODC admission "
                "has no byte-output comparison.")))

    path = res / "cli.json"
    d = load_json(path)
    for fixture, baseline in d["baseline"]["cases"].items():
        candidate = d["candidate"]["cases"].get(fixture)
        if candidate is None:
            raise SystemExit(f"0377a CLI fixture missing candidate: {fixture}")
        fmt = fixture.split("/", 2)[1]
        for leg in ("in-process", "cli-get", "cli-version"):
            before = baseline[leg]["median_us"]
            after = candidate[leg]["median_us"]
            out.append(row(
                "0377a", "cli", "get",
                f"{fixture} {leg} (candidate vs baseline)", "",
                before, after, [after / before], NONE,
                f"cli_0377a:{fixture}:{leg}:after-before", src([path]),
                comparison="after/before (wall, ABBA processes)",
                explanation="Four-fixture CLI ABBA receipt; process-start diagnostic only."))
    return out


PARSERS = [p0250, p0251, p0251f, p0252, p0253, p0254, p0255, p0256, p0258, p0259,
           p0260, p0261, p0262, p0264, p0266, p0267, p0268, p0269, p0270, p0271,
           p0273, p0274, p0275, p0276, p0277, p0278, p0279, p0280, p0281, p0282,
           p0283, p0284, p0285, p0286, p0287, p0288, p0289, p0290, p0291, p0292,
           p0293, p0295, p0296, p0297, p0298, p0299, p0301, p0302, p0303, p0304, p0305,
           p0306, p0307, p0308, p0309, p0310, p0311, p0312,
           p0313, p0314, p0315, p0316, p0317, p0318, p0323, p0324, p0325, p0326,
           p0328, p0329, p0330, p0333, p0335, p0336, p0337, p0338, p0340, p0342, p0345, p0346,
           p0347, p0348, p0355, p0355a, p0356, p0356a, p0356b, p0356c, p0357, p0357a, p0357b, p0358, p0358a,
           p0359, p0359a, p0359b, p0359qb, p0359qc, p0359qd, p0359r, p0359ra, p0359rb, p0359rc,
           p0359rd, p0359re, p0359qe, p0360b, p0361, p0361a, p0362, p0362ya, p0362yb, p0362yc, p0362yd, p0362ye, p0375b, p0377, p0377a, p0391, p0399, p0403]




# ---------------------------------------------------------------- merge

def curated_rows(cur):
    out = []
    for c in cur.get("rows", []):
        budget = (c.get("budget_value"), c.get("budget", "none"),
                  c.get("budget_kind", "none"))
        ratios = [c["ratio"]] + ([c["ratio_min"]] if c.get("ratio_min") is not None else [])
        r = row(c["record"], c["format"], c["operation"], c["scenario"], c.get("size", ""),
                c.get("leaf"), c.get("unified"), ratios, budget, c["series"],
                "curated:" + c["cite"], c.get("stage", ""), c.get("stage_rank", 1),
                c.get("comparison", "unified/leaf"), c.get("unit", "µs"),
                c.get("explanation", ""))
        r["verdict"] = c.get("verdict")
        r["ratio_text"] = c.get("ratio_text")
        if c.get("disposition"):
            r["disposition"] = c["disposition"]
        out.append(r)
    return out


def key(r):
    return (r["record"], r["series"], r["stage"])


def merge(parsed, curated):
    rows = {}
    conflicts, agreements = [], []
    for r in parsed:
        k = key(r)
        if k in rows:
            raise SystemExit(f"duplicate parsed key {k}")
        rows[k] = r
    for c in curated:
        k = key(c)
        if k in rows:
            p = rows[k]
            ok = True
            for cv, pv in ((c["ratio"], p["ratio"]), (c["ratio_min"], p["ratio_min"])):
                if cv is None or pv is None:
                    continue
                if abs(cv - pv) > max(0.006, 0.006 * pv):
                    ok = False
            (agreements if ok else conflicts).append((c, p))
            continue
        rows[k] = c
    return list(rows.values()), conflicts, agreements


def apply_annotations(rows, cur):
    for a in cur.get("annotations", []):
        pat = re.compile(a["series"])
        hit = 0
        for r in rows:
            if r["record"] != a["record"] or not pat.fullmatch(r["series"]):
                continue
            if "stage" in a and r["stage"] != a["stage"]:
                continue
            hit += 1
            if "explanation" in a:
                r["explanation"] = a["explanation"]
            if "verdict" in a:
                r["verdict"] = a["verdict"]
            if "superseded_by" in a:
                r["superseded_by"] = a["superseded_by"]
            if "disposition" in a:
                r["disposition"] = a["disposition"]
            if "budget" in a:
                r["budget_value"], r["budget"] = a.get("budget_value"), a["budget"]
                r["budget_kind"] = a.get("budget_kind", r["budget_kind"])
        if hit == 0:
            raise SystemExit(f"annotation matched no row: {a['record']} {a['series']}")


def base_status(r):
    if r["verdict"]:
        return r["verdict"], ""
    if r["ratio"] is None or r["budget_value"] is None:
        return "no-budget", ""
    if r["ratio"] <= r["budget_value"] + 1e-12:
        return "meets", ""
    if (r["budget_kind"] == "open" and r["leaf"] is not None and r["unified"] is not None
            and r["unit"] == "µs" and r["unified"] - r["leaf"] <= 20.0):
        return "meets", f"meets the open budget through the ≤20 µs alternative (+{r['unified'] - r['leaf']:.1f} µs)"
    return "miss", ""


def finalize(rows):
    for r in rows:
        r["status"], note = base_status(r)
        if note:
            r["explanation"] = (note + ("; " + r["explanation"] if r["explanation"] else ""))
    by_series: dict = {}
    for r in rows:
        by_series.setdefault(r["series"], []).append(r)
    for series, items in by_series.items():
        items.sort(key=lambda r: (r["record"], r["stage_rank"]))
        latest = items[-1]
        for r in items[:-1]:
            later = [x for x in items if (x["record"], x["stage_rank"]) > (r["record"], r["stage_rank"])]
            nxt = later[0]
            # The superseding record is the first later measurement that meets
            # its budget (the fix), else the latest one (the miss stays open there).
            fixed = [x for x in later if x["status"] == "meets"]
            by = r.get("superseded_by") or (fixed[0] if fixed else latest)["record"]
            if r["status"] == "miss":
                r["status"] = f"miss (superseded by {by})"
                chain = "; ".join(
                    f"{x['record']}{' ' + x['stage'] if x['stage'] else ''}: {fmt_ratio(x)} ({x['status']})"
                    for x in later)
                text = f"Re-measured later — {chain}."
                if nxt["explanation"] and latest["status"] != "meets":
                    # Record 0344: the latest row carries its own explanation;
                    # repeating it here pushed the outputs past 1 MiB.
                    text += f" Latest explanation: the record {latest['record']} row."
                r["explanation"] = (text + (" Original: " + r["explanation"] if r["explanation"] else ""))
    return rows


# Record 0337: every open miss names a follow-up or an explained residual.
# A curated row or annotation may say so explicitly (`disposition`:
# "follow-up" or "residual"); otherwise the explanation's wording decides.
FOLLOW_UP_WORDS = re.compile(r"follow-?up|\bF\d+\b|-F\d+\b", re.I)
RESIDUAL_WORDS = re.compile(r"inherent|owner decision|owner-accepted|explained residual", re.I)
DISPOSITIONS = ("follow-up", "residual")


def disposition(r):
    d = r.get("disposition")
    if d:
        if d not in DISPOSITIONS:
            raise SystemExit(f"unknown disposition {d!r}: {r['record']} {r['series']}")
        return d
    if FOLLOW_UP_WORDS.search(r["explanation"]):
        return "follow-up"
    if RESIDUAL_WORDS.search(r["explanation"]):
        return "residual"
    return ""


# ---------------------------------------------------------------- output

def fmt_num(v):
    if v is None:
        return ""
    if abs(v) >= 1000:
        return f"{v:.1f}"
    if abs(v) >= 10:
        return f"{v:.2f}"
    return f"{v:.3f}"


def fmt_ratio(r):
    if r.get("ratio_text"):
        return r["ratio_text"]
    v = r["ratio"]
    if v is None:
        return ""
    return f"{v:.1f}" if v >= 100 else f"{v:.3f}"


def fmt_range(r):
    lo, hi = r["ratio_min"], r["ratio"]
    if lo is None or hi is None or abs(hi - lo) < 5e-4:
        return ""
    f = (lambda v: f"{v:.1f}") if hi >= 100 else (lambda v: f"{v:.3f}")
    return f"{f(lo)}–{f(hi)}"


def size_key(s):
    m = re.match(r"([\d,]+)", s or "")
    return (0, int(m.group(1).replace(",", ""))) if m else (1, s or "")


def sort_rows(rows):
    return sorted(rows, key=lambda r: (r["format"], r["operation"], r["scenario"],
                                       size_key(r["size"]), r["comparison"], r["record"],
                                       r["stage_rank"], r["stage"]))


def to_tsv(rows, f=None):
    """The row table, of format `f` only when given (record 0348)."""
    buf = io.StringIO()
    w = csv.writer(buf, delimiter="\t", lineterminator="\n", quoting=csv.QUOTE_MINIMAL)
    w.writerow(TSV_COLUMNS)
    for r in rows:
        if f is not None and r["format"] != f:
            continue
        w.writerow([r["format"], r["operation"], r["scenario"], r["size"], r["comparison"],
                    r["stage"], fmt_num(r["leaf"]), fmt_num(r["unified"]), r["unit"],
                    fmt_ratio(r), fmt_range(r), r["budget"], r["status"], r["record"],
                    r["source"], r["explanation"]])
    return buf.getvalue()


def md_cell(s):
    return str(s).replace("|", "\\|").replace("\n", " ")


def to_md(rows, conflicts, agreements, cur):
    L = []
    a = L.append
    a(f"# Unified-vs-leaf benchmark matrix (records 0249–{FOLLOW_RANGE[1]:04d})")
    a("")
    a("Generated by `tools/unified_bench_matrix.py` from the evidence under "
      "`docs/unified-ops/results/` and the curated rows in "
      "`results/0289/bench-matrix-curated.json`. Do not edit by hand; "
      "regenerate, and check with `python3 tools/unified_bench_matrix.py --check`. "
      "Records 0249–0284 were collected by record 0289, records 0285–0302 by record 0294; "
      "record 0306 added the parsers of 0298, 0301, 0302, 0303, 0304 and 0306; "
      "record 0310 those of 0305, 0309, 0310, 0311, 0312 and 0313; records 0315, 0326, 0337 and 0340 their own; "
      "record 0344 those of 0307, 0308, 0314, 0316–0318, 0323–0325, 0328–0330, 0333, 0335, 0336, 0338 and 0342, "
      "and the check that every record's evidence contributes rows.")
    a("")
    a("- Ratio = unified / leaf, or the comparison the `comparison` column names. "
      "With several trials or runs the ratio is the worst (largest) value and the "
      "range column shows min–max; status uses the worst value. Times are µs, "
      "medians over trials.")
    a("- Budgets (docs/UNIFIED_OPS_GOAL.md, PERFORMANCE BAR): open 1.05 or +20 µs; "
      "single-op write 1.10; 1k-op batch 2.0 against one hand-written leaf edit; "
      "other batch sizes are scaling only unless the record applies 2.0; "
      "reads and queries carry a budget only where the record states one "
      "(record 0262's 1.10 review trigger for full queries); before/after-build "
      "records use 1.05 (no regression above 5%).")
    a("- A series measured again later is superseded: its miss shows as "
      "`miss (superseded by NNNN)` and the open misses are the latest measurements.")
    a("- Source `parsed:<file>` rows come from evidence files; `curated:<record §>` "
      "rows quote record prose verbatim.")
    a("")
    for note in cur.get("notes", []):
        a(f"- {note}")
    a("")
    a(f"## Records {FOLLOW_RANGE[0]:04d}–{FOLLOW_RANGE[1]:04d}")
    a("")
    evidence = evidence_records()
    exempt = {e["record"]: e["reason"] for e in cur.get("evidence_without_rows", [])}
    a("| record | file | evidence | rows | parsed | curated | open misses |")
    a("|---|---|---|---:|---:|---:|---:|")
    for n in range(FOLLOW_RANGE[0], FOLLOW_RANGE[1] + 1):
        rec = f"{n:04d}"
        f = record_file(rec)
        if f is None:
            a(f"| {rec} | not on this tree | | | | | |")
            continue
        rr = [r for r in rows if r["record"] == rec]
        ev = ", ".join(f"`{e}`" for e in evidence.get(rec, [])) or "—"
        if rec in exempt:
            ev += f" (no rows: {exempt[rec]})"
        a(f"| {rec} | `{f.name}` | {md_cell(ev)} | {len(rr)} | "
          f"{sum(r['source'].startswith('parsed') for r in rr)} | "
          f"{sum(r['source'].startswith('curated') for r in rr)} | "
          f"{sum(r['status'] == 'miss' for r in rr)} |")
    a("")
    fmts = sorted({r["format"] for r in rows})
    a("## Summary")
    a("")
    a("| format | rows | parsed | curated | meets | miss (open) | miss (superseded) | unqualified | no-budget |")
    a("|---|---:|---:|---:|---:|---:|---:|---:|---:|")
    tot = [0] * 8
    for f in fmts:
        fr = [r for r in rows if r["format"] == f]
        c = [len(fr),
             sum(r["source"].startswith("parsed") for r in fr),
             sum(r["source"].startswith("curated") for r in fr),
             sum(r["status"] == "meets" for r in fr),
             sum(r["status"] == "miss" for r in fr),
             sum(r["status"].startswith("miss (") for r in fr),
             sum(r["status"] == "unqualified" for r in fr),
             sum(r["status"] == "no-budget" for r in fr)]
        tot = [x + y for x, y in zip(tot, c)]
        a(f"| {f} | " + " | ".join(str(x) for x in c) + " |")
    a("| **all** | " + " | ".join(f"**{x}**" for x in tot) + " |")
    a("")

    def miss_table(items):
        a("| format | operation | scenario | size | comparison | stage | ratio | budget | record | explanation / follow-up |")
        a("|---|---|---|---|---|---|---:|---|---|---|")
        for r in items:
            a("| " + " | ".join(md_cell(x) for x in (
                r["format"], r["operation"], r["scenario"], r["size"], r["comparison"],
                r["stage"], fmt_ratio(r) + (f" ({fmt_range(r)})" if fmt_range(r) else ""),
                r["budget"], r["record"], r["explanation"])) + " |")
        a("")

    a("## Misses")
    a("")
    open_m = [r for r in rows if r["status"] == "miss"]
    sup_m = [r for r in rows if r["status"].startswith("miss (")]
    a(f"### Open misses ({len(open_m)})")
    a("")
    n_fu = sum(disposition(r) == "follow-up" for r in open_m)
    n_res = sum(disposition(r) == "residual" for r in open_m)
    a(f"Dispositions (record 0337): {n_fu} name a follow-up, {n_res} are an explained "
      f"residual (inherent to the comparison, or kept by an owner decision), "
      f"{len(open_m) - n_fu - n_res} name neither. `--check` fails on the last kind.")
    a("")
    miss_table(open_m)
    a(f"### Superseded misses ({len(sup_m)})")
    a("")
    miss_table(sup_m)

    a("## Cross-checks of record prose against parsed evidence")
    a("")
    a(f"Curated rows that share a key with a parsed row are not added; their "
      f"ratios are compared (tolerance 0.006 or 0.6%). {len(agreements)} agree, "
      f"{len(conflicts)} conflict.")
    a("")
    if conflicts:
        a("| record | series | stage | prose ratio (min–max) | parsed ratio (min–max) | cite |")
        a("|---|---|---|---|---|---|")
        for c, p in conflicts:
            a("| " + " | ".join(md_cell(x) for x in (
                c["record"], c["series"], c["stage"],
                f"{fmt_num(c['ratio_min'])}–{fmt_num(c['ratio'])}",
                f"{fmt_num(p['ratio_min'])}–{fmt_num(p['ratio'])}", c["source"])) + " |")
        a("")
    a("## Rows by format")
    a("")
    a("Each format's rows are on their own page (record 0344) and in their own "
      "tab-separated table (record 0348), under `bench-matrix/`:")
    a("")
    for f in fmts:
        a(f"- [{f}](bench-matrix/{f}.md), [TSV](bench-matrix/{f}.tsv) "
          f"({sum(r['format'] == f for r in rows)} rows)")
    return "\n".join(L).rstrip() + "\n"


def format_page(f, rows):
    """One format's rows (record 0344 moved them out of bench-matrix.md)."""
    L = []
    a = L.append
    a(f"# Unified-vs-leaf benchmark matrix: {f}")
    a("")
    a("Generated by `tools/unified_bench_matrix.py`; do not edit by hand. Conventions, "
      "budgets and the misses are in [bench-matrix.md](../bench-matrix.md).")
    a("")
    a("| operation | scenario | size | comparison | stage | leaf µs | unified µs | ratio | range | budget | status | record | source |")
    a("|---|---|---|---|---|---:|---:|---:|---|---|---|---|---|")
    for r in rows:
        if r["format"] != f:
            continue
        unit_note = "" if r["unit"] == "µs" else f" {r['unit']}"
        a("| " + " | ".join(md_cell(x) for x in (
            r["operation"], r["scenario"], r["size"], r["comparison"], r["stage"],
            fmt_num(r["leaf"]) + (unit_note if r["leaf"] is not None else ""),
            fmt_num(r["unified"]) + (unit_note if r["unified"] is not None else ""),
            fmt_ratio(r), fmt_range(r), r["budget"], r["status"], r["record"],
            r["source"])) + " |")
    return "\n".join(L).rstrip() + "\n"


def outputs(rows, conflicts, agreements, cur):
    """Every generated file and its text."""
    files = {OUT_MD: to_md(rows, conflicts, agreements, cur)}
    for f in sorted({r["format"] for r in rows}):
        files[OUT_FORMAT_DIR / f"{f}.md"] = format_page(f, rows)
        files[OUT_FORMAT_DIR / f"{f}.tsv"] = to_tsv(rows, f)
    return files


# Record 0344: evidence under `results/` belongs to the record its name
# starts with (`0328/`, `0255-python-overhead.json`, `0251a-docx-…/`).
# Follow-up IDs may contain a numeric segment after their alphabetic suffix
# (`0362ul3`, `0362ul10f`); retain the full ID before the filename boundary.
EVIDENCE_NAME = re.compile(r"^(\d{4}(?:[a-z]+(?:\d+[a-z]*)?)?)(?=[-.]|$)")
# The matrix collects records from 0249 on (record 0289).
FIRST_RECORD = 249


def evidence_records(res=None):
    """Record id -> its evidence entries under `results/` (directories with
    at least one file, or files), for records the matrix covers."""
    res = res or RES
    found: dict = {}
    for entry in sorted(res.iterdir()):
        m = EVIDENCE_NAME.match(entry.name)
        if not m or int(m.group(1)[:4]) < FIRST_RECORD:
            continue
        if entry.is_dir() and not any(p.is_file() for p in entry.rglob("*")):
            continue
        found.setdefault(m.group(1), []).append(entry.name)
    return found


def coverage_problems(rows, cur, evidence):
    """Records whose evidence contributes no row (record 0344), unless the
    curated `evidence_without_rows` says why; and stale exemptions."""
    with_rows = {r["record"] for r in rows}
    exempt = {e["record"]: e for e in cur.get("evidence_without_rows", [])}
    problems = []
    for rec, entries in sorted(evidence.items()):
        if rec in with_rows or rec in exempt:
            continue
        problems.append(f"record {rec}: results/{entries[0]} holds evidence but the record "
                        f"contributes no rows; parse it, curate it, or add it to "
                        f"`evidence_without_rows` with the reason")
    for rec, e in sorted(exempt.items()):
        if not e.get("reason", "").strip():
            problems.append(f"record {rec}: `evidence_without_rows` entry has no reason")
        if rec in with_rows:
            problems.append(f"record {rec}: listed in `evidence_without_rows` but contributes rows")
        elif rec not in evidence:
            problems.append(f"record {rec}: listed in `evidence_without_rows` but has no evidence")
    return problems


def build():
    cur = load_json(CURATED)
    parsed = []
    for p in PARSERS:
        parsed += p()
    rows, conflicts, agreements = merge(parsed, curated_rows(cur))
    apply_annotations(rows, cur)
    rows = sort_rows(finalize(rows))
    problems = [r for r in rows if r["status"].startswith("miss") and not r["explanation"].strip()]
    problems += [r for r in rows if r["status"] == "miss" and r["explanation"].strip()
                 and not disposition(r)]
    problems += coverage_problems(rows, cur, evidence_records())
    return rows, conflicts, agreements, cur, problems


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0],
                                 formatter_class=argparse.RawDescriptionHelpFormatter,
                                 epilog="Outputs: " + rel(OUT_MD) + " and " +
                                 rel(OUT_FORMAT_DIR) + "/<format>.{md,tsv}" +
                                 " (paths relative to docs/unified-ops/results/).")
    ap.add_argument("--check", action="store_true",
                    help="regenerate in memory; exit 1 if the committed outputs differ, "
                         "a miss lacks an explanation or a disposition, or a record's "
                         "evidence contributes no rows")
    ap.add_argument("--stats", action="store_true", help="print row counts per format and source")
    args = ap.parse_args(argv)
    rows, conflicts, agreements, cur, problems = build()
    files = outputs(rows, conflicts, agreements, cur)
    rc = 0
    for path, text in files.items():
        if len(text.encode("utf-8")) >= SIZE_LIMIT:
            print(f"{path.relative_to(ROOT)} would be {len(text.encode('utf-8'))} bytes, "
                  f"at or above the 1 MiB limit", file=sys.stderr)
            rc = 1
    for r in problems:
        if isinstance(r, str):
            print(r, file=sys.stderr)
        else:
            what = "explanation" if not r["explanation"].strip() else "follow-up or residual"
            print(f"miss without {what}: {r['record']} {r['series']} {r['stage']}", file=sys.stderr)
        rc = 1
    stale_pages = sorted((set(OUT_FORMAT_DIR.glob("*.md")) | set(OUT_FORMAT_DIR.glob("*.tsv")))
                         - set(files)) if OUT_FORMAT_DIR.is_dir() else []
    if OLD_TSV.is_file():
        stale_pages.append(OLD_TSV)
    if args.check:
        for path, text in files.items():
            if not path.is_file() or path.read_text(encoding="utf-8") != text:
                print(f"{path.relative_to(ROOT)} is stale; run tools/unified_bench_matrix.py",
                      file=sys.stderr)
                rc = 1
        for path in stale_pages:
            print(f"{path.relative_to(ROOT)} is no longer generated; run tools/unified_bench_matrix.py",
                  file=sys.stderr)
            rc = 1
        return rc
    OUT_FORMAT_DIR.mkdir(parents=True, exist_ok=True)
    for path, text in files.items():
        path.write_text(text, encoding="utf-8")
    for path in stale_pages:
        path.unlink()
    n_p = sum(r["source"].startswith("parsed") for r in rows)
    print(f"{len(rows)} rows ({n_p} parsed, {len(rows) - n_p} curated); "
          f"{len(conflicts)} conflicts, {len(agreements)} agreements; "
          f"{sum(r['status'] == 'miss' for r in rows)} open misses")
    if args.stats:
        for f in sorted({r["format"] for r in rows}):
            fr = [r for r in rows if r["format"] == f]
            p = sum(r["source"].startswith("parsed") for r in fr)
            print(f"  {f}: {len(fr)} rows ({p} parsed, {len(fr) - p} curated)")
    return rc


if __name__ == "__main__":
    sys.exit(main())
