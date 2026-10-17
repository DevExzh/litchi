#!/usr/bin/env python3
"""Generate the unified-vs-leaf benchmark matrix (records 0289, 0294, 0306, 0310, 0315, 0326, 0337, 0340).

The matrix collects every paired measurement recorded by the unified
operations records 0249-0340 (`docs/unified-ops/02[4-9]*.md`, `03[0-4]*.md`)
into one table: format x operation x scenario/size -> leaf, unified, ratio, budget,
status (meets / miss / no-budget, or "miss (superseded by NNNN)"), record,
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
0305, 0309, 0310, 0311, 0312 and 0313 parsers; records 0315, 0326, 0337 and 0340 their own) are collected
only when their record file is on this tree; the Markdown lists every
record of that range, absent ones as "not on this tree", so the tool can
be rerun unchanged after later merges.

Outputs: docs/unified-ops/results/0289/bench-matrix.tsv and bench-matrix.md.
`--check` regenerates in memory and exits non-zero when the committed
outputs differ or when a miss lacks an explanation.
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
OUT_TSV = OUT_DIR / "bench-matrix.tsv"
OUT_MD = OUT_DIR / "bench-matrix.md"

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
        for line in open(p, encoding="utf-8"):
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
            for line in open(f, encoding="utf-8"):
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
        for line in open(p, encoding="utf-8"):
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

FOLLOW_RANGE = (285, 340)


def record_file(record: str):
    hits = sorted(UO.glob(f"{record}-*.md"))
    return hits[0] if hits else None


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
        for line in open(p, encoding="utf-8"):
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
    for line in open(text_path, encoding="utf-8"):
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
        for line in open(p, encoding="utf-8"):
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


PARSERS = [p0250, p0251, p0251f, p0252, p0253, p0254, p0255, p0256, p0258, p0259,
           p0260, p0261, p0262, p0264, p0266, p0267, p0268, p0269, p0270, p0271,
           p0273, p0274, p0275, p0276, p0277, p0278, p0279, p0280, p0281, p0282,
           p0283, p0284, p0285, p0286, p0287, p0288, p0289, p0290, p0291, p0292,
           p0293, p0295, p0296, p0297, p0298, p0299, p0301, p0302, p0303, p0304, p0305,
           p0306, p0309, p0310, p0311, p0312,
           p0313, p0315, p0326, p0337, p0340]


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
                    text += f" Latest explanation ({latest['record']}): {latest['explanation']}"
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


def to_tsv(rows):
    buf = io.StringIO()
    w = csv.writer(buf, delimiter="\t", lineterminator="\n", quoting=csv.QUOTE_MINIMAL)
    w.writerow(TSV_COLUMNS)
    for r in rows:
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
      "record 0310 those of 0305, 0309, 0310, 0311, 0312 and 0313; records 0315, 0326, 0337 and 0340 their own.")
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
    a("| record | file | rows | parsed | curated | open misses |")
    a("|---|---|---:|---:|---:|---:|")
    for n in range(FOLLOW_RANGE[0], FOLLOW_RANGE[1] + 1):
        rec = f"{n:04d}"
        f = record_file(rec)
        if f is None:
            a(f"| {rec} | not on this tree | | | | |")
            continue
        rr = [r for r in rows if r["record"] == rec]
        a(f"| {rec} | `{f.name}` | {len(rr)} | "
          f"{sum(r['source'].startswith('parsed') for r in rr)} | "
          f"{sum(r['source'].startswith('curated') for r in rr)} | "
          f"{sum(r['status'] == 'miss' for r in rr)} |")
    a("")
    fmts = sorted({r["format"] for r in rows})
    a("## Summary")
    a("")
    a("| format | rows | parsed | curated | meets | miss (open) | miss (superseded) | no-budget |")
    a("|---|---:|---:|---:|---:|---:|---:|---:|")
    tot = [0] * 7
    for f in fmts:
        fr = [r for r in rows if r["format"] == f]
        c = [len(fr),
             sum(r["source"].startswith("parsed") for r in fr),
             sum(r["source"].startswith("curated") for r in fr),
             sum(r["status"] == "meets" for r in fr),
             sum(r["status"] == "miss" for r in fr),
             sum(r["status"].startswith("miss (") for r in fr),
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
    for f in fmts:
        fr = [r for r in rows if r["format"] == f]
        a(f"## {f}")
        a("")
        a("| operation | scenario | size | comparison | stage | leaf µs | unified µs | ratio | range | budget | status | record | source |")
        a("|---|---|---|---|---|---:|---:|---:|---|---|---|---|---|")
        for r in fr:
            unit_note = "" if r["unit"] == "µs" else f" {r['unit']}"
            a("| " + " | ".join(md_cell(x) for x in (
                r["operation"], r["scenario"], r["size"], r["comparison"], r["stage"],
                fmt_num(r["leaf"]) + (unit_note if r["leaf"] is not None else ""),
                fmt_num(r["unified"]) + (unit_note if r["unified"] is not None else ""),
                fmt_ratio(r), fmt_range(r), r["budget"], r["status"], r["record"],
                r["source"])) + " |")
        a("")
    return "\n".join(L).rstrip() + "\n"


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
    return rows, conflicts, agreements, cur, problems


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0],
                                 formatter_class=argparse.RawDescriptionHelpFormatter,
                                 epilog="Outputs: " + rel(OUT_TSV) + ", " + rel(OUT_MD) +
                                 " (paths relative to docs/unified-ops/results/).")
    ap.add_argument("--check", action="store_true",
                    help="regenerate in memory; exit 1 if the committed outputs differ "
                         "or a miss lacks an explanation")
    ap.add_argument("--stats", action="store_true", help="print row counts per format and source")
    args = ap.parse_args(argv)
    rows, conflicts, agreements, cur, problems = build()
    tsv, md = to_tsv(rows), to_md(rows, conflicts, agreements, cur)
    rc = 0
    for r in problems:
        what = "explanation" if not r["explanation"].strip() else "follow-up or residual"
        print(f"miss without {what}: {r['record']} {r['series']} {r['stage']}", file=sys.stderr)
        rc = 1
    if args.check:
        for path, text in ((OUT_TSV, tsv), (OUT_MD, md)):
            if not path.is_file() or path.read_text(encoding="utf-8") != text:
                print(f"{path.relative_to(ROOT)} is stale; run tools/unified_bench_matrix.py",
                      file=sys.stderr)
                rc = 1
        return rc
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    OUT_TSV.write_text(tsv, encoding="utf-8")
    OUT_MD.write_text(md, encoding="utf-8")
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
