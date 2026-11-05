#!/usr/bin/env python3
"""Summarize a 0003 capture directory (before/ or after/).

Reads raw/R<n>-<group>.json reports and children.jsonl written by capture.py
and writes analysis.json and summary.csv next to them.

Per row (case, corpus name, corpus archive SHA-256, cache state):
- pooled p50/p95 over every retained sample of every repeat;
- a 95% bootstrap interval for the pooled p50 (resampling within each repeat,
  2,000 draws, fixed seed);
- per-repeat p50s and their spread, (max - min) / median, in percent. The
  spread is the run-to-run noise S10 has to beat: a before/after difference
  smaller than this spread is not distinguishable on this host.
Whole-child peak RSS (GNU time) is per group child, not per route.

Repeats R1-R3 cover every group; R4-R5 re-run only the groups whose R1-R3
spread exceeded 5% (see README). Both spreads are reported.

Corpus identity: every row carries its archive SHA-256; corpus_identity_sha256
hashes the sorted (case, corpus, archive SHA-256) set, and each group's schema-2
catalog/content-set hashes must agree across its repeats.

Usage: analyze.py DIR
"""

import gzip
import hashlib
import json
import os
import random
import statistics
import sys
from collections import defaultdict


def category(case):
    if case.startswith("detect_"):
        return "detection"
    if "full_text" in case:
        return "full-text"
    if any(k in case for k in ("list_slides", "list_paragraphs", "slide_count", "paragraph_count")):
        return "list"
    if any(k in case for k in ("selected_", "one_paragraph", "one_slide")):
        return "read-one"
    if case.endswith("_open") or case.endswith("_open_lifecycle"):
        return "open"
    return "other"


def p50(values):
    ordered = sorted(values)
    n = len(ordered)
    return (ordered[(n - 1) // 2] + ordered[n // 2]) / 2


def nearest_rank(values, pct):
    ordered = sorted(values)
    rank = max(1, -(-pct * len(ordered) // 100))
    return ordered[rank - 1]


def main():
    root = sys.argv[1]
    raw = os.path.join(root, "raw")
    rows = defaultdict(dict)  # key -> repeat -> samples
    meta = {}
    for name in sorted(os.listdir(raw)):
        if name.endswith(".corpus-manifest.json"):
            continue
        if name.endswith(".json"):
            stem, opener = name[:-5], open
        elif name.endswith(".json.gz"):
            stem, opener = name[:-8], gzip.open
        else:
            continue
        repeat = int(stem.split("-", 1)[0][1:])
        group = stem.split("-", 1)[1]
        with opener(os.path.join(raw, name), "rt") as handle:
            report = json.load(handle)
        for result in report["results"]:
            corpus = result["corpus"]
            key = (result["case"], corpus.get("name"), corpus.get("archive_sha256"),
                   result.get("cache_state"))
            if repeat in rows[key]:
                sys.exit(f"duplicate row {key} in repeat {repeat}")
            rows[key][repeat] = result["elapsed_ns"]["samples"]
            meta[key] = {"group": group, "archive_bytes": corpus.get("archive_bytes")}

    rss = defaultdict(list)
    children = []
    with open(os.path.join(root, "children.jsonl")) as handle:
        for line in handle:
            child = json.loads(line)
            children.append(child)
            if "peak_rss_kib" in child.get("gnu_time", {}):
                rss[child["group"]].append(child["gnu_time"]["peak_rss_kib"])

    fixtures = {}
    env_path = os.path.join(root, "environment.json")
    if os.path.exists(env_path):
        for path, digest in json.load(open(env_path)).get("ole2_fixtures", {}).items():
            fixtures[digest] = os.path.basename(path)

    catalogs = defaultdict(set)
    for name in os.listdir(raw):
        if name.endswith(".corpus-manifest.json"):
            manifest = json.load(open(os.path.join(raw, name)))
            group = name.split("-", 1)[1][: -len(".corpus-manifest.json")]
            catalogs[group].add((manifest["catalog_sha256"], manifest["content_set_sha256"]))
    unstable = sorted(g for g, v in catalogs.items() if len(v) != 1)
    identity = hashlib.sha256("".join(
        f"{k[0]}\t{k[1]}\t{k[2]}\n" for k in sorted(rows, key=lambda k: (k[0], k[1] or "", k[2] or ""))
    ).encode()).hexdigest()

    rng = random.Random(3)
    out = []
    for key in sorted(rows, key=lambda k: (category(k[0]), k[0], k[1] or "")):
        case, corpus_name, corpus_sha, cache_state = key
        per_repeat = rows[key]
        pooled = [v for r in sorted(per_repeat) for v in per_repeat[r]]
        repeat_p50 = {r: p50(v) for r, v in sorted(per_repeat.items())}
        draws = []
        for _ in range(2000):
            sample = []
            for values in per_repeat.values():
                sample.extend(rng.choice(values) for _ in values)
            draws.append(p50(sample))
        draws.sort()
        def spread_of(mids):
            return (max(mids) - min(mids)) / statistics.median(mids) * 100 if len(mids) > 1 else None
        spread = spread_of(list(repeat_p50.values()))
        spread_r1_r3 = spread_of([v for r, v in repeat_p50.items() if r <= 3])
        out.append({
            "category": category(case),
            "case": case,
            "corpus": corpus_name,
            "fixture": fixtures.get(corpus_sha),
            "corpus_archive_sha256": corpus_sha,
            "cache_state": cache_state,
            "group": meta[key]["group"],
            "archive_bytes": meta[key]["archive_bytes"],
            "repeats": len(per_repeat),
            "samples": len(pooled),
            "p50_ns": p50(pooled),
            "p50_ci95_bootstrap_ns": [draws[49], draws[1949]],
            "p95_ns": nearest_rank(pooled, 95),
            "repeat_p50_ns": {f"R{r}": v for r, v in repeat_p50.items()},
            "repeat_p50_spread_pct": spread,
            "repeat_p50_spread_r1_r3_pct": spread_r1_r3,
            "whole_child_peak_rss_kib_max": max(rss[meta[key]["group"]]) if rss[meta[key]["group"]] else None,
        })

    loads = [float(c["loadavg_before"][0]) for c in children] + \
            [float(c["loadavg_after"][0]) for c in children]
    analysis = {
        "schema": "litchi.unified-ops.0003.capture-analysis.v1",
        "directory": os.path.basename(os.path.abspath(root)),
        "children": len(children),
        "children_failed": [f"R{c['repeat']}-{c['group']}" for c in children if c["exit_code"] != 0],
        "loadavg_1min_min": min(loads) if loads else None,
        "loadavg_1min_max": max(loads) if loads else None,
        "corpus_identity_sha256": identity,
        "corpus_catalogs": {g: [{"catalog_sha256": c, "content_set_sha256": t} for c, t in sorted(v)]
                            for g, v in sorted(catalogs.items())},
        "corpus_catalogs_unstable_across_repeats": unstable,
        "rows": out,
    }
    with open(os.path.join(root, "analysis.json"), "w") as handle:
        json.dump(analysis, handle, indent=1)
        handle.write("\n")
    with open(os.path.join(root, "summary.csv"), "w") as handle:
        handle.write("category,case,corpus,cache_state,repeats,samples,p50_ns,p50_ci_lo_ns,p50_ci_hi_ns,"
                     "p95_ns,repeat_p50_spread_r1_r3_pct,repeat_p50_spread_pct,whole_child_peak_rss_kib\n")
        for row in out:
            def fmt(v):
                return "" if v is None else f"{v:.2f}"
            corpus = row["corpus"] + (f":{row['fixture']}" if row["fixture"] else "")
            handle.write(",".join(str(x) for x in (
                row["category"], row["case"], corpus, row["cache_state"] or "",
                row["repeats"], row["samples"], row["p50_ns"], row["p50_ci95_bootstrap_ns"][0],
                row["p50_ci95_bootstrap_ns"][1], row["p95_ns"],
                fmt(row["repeat_p50_spread_r1_r3_pct"]), fmt(row["repeat_p50_spread_pct"]),
                row["whole_child_peak_rss_kib_max"] or "")) + "\n")
    print(f"{len(out)} rows, {len(children)} children, failed={analysis['children_failed']}, "
          f"unstable catalogs={unstable}, corpus identity={identity}")


if __name__ == "__main__":
    main()
