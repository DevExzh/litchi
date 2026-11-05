#!/usr/bin/env python3
"""Analyze an M0a S10 ABBA capture (capture_s10.py) against the S1 rule.

Inputs:
- CAPTURE: the capture directory (raw/B<block>-<pos><leg>-<group>.json[.gz]
  and children.jsonl from capture_s10.py);
- ../before/analysis.json: S1's per-route noise (repeat_p50_spread_pct over
  R1-R5, "spread_all");
- ../s8b/route-map.json: predecessor -> unified successor -> leaf twin;
- explanations.json (next to this script): reviewed explanations for rows
  over their trigger, keyed by "<pair>|<after case>|<corpus>".

Per row and leg, the statistic is the median of the per-child p50s (each
child's p50 is the median of its retained samples, as in S1's analyze.py).

Pairs:
- facade: S1 selector on the pre-M0a binary (leg A) vs its unified successor
  on the post-S9 binary (leg U), only where the route map says the timed
  scopes match;
- kept: a predecessor S9 kept, same selector on both binaries (A vs U);
- overhead: unified successor (U) vs its leaf twin (L), both post-S9.
Successors whose timed scope differs from the predecessor's are listed as
budget-only rows (no facade verdict).

S1 rule (facade and kept pairs): trigger = max(5%, the row's S1 spread_all);
rows whose before p50 is under 5 us also need an absolute increase above
0.5 us; rows whose S1 spread is above 10% need full separation (every after
child's p50 above every before child's p50). doc_facade_file_open and
doc_facade_file_full_text on FloatingPictures.doc are bimodal per process in
S1, so they are compared mode by mode.

Overhead pair: "open" rows use the goal budget, unified - leaf <= max(5% of
leaf, 20 us); the other rows use the S1 rule as a review trigger, with the
predecessor's S1 spread as the noise.

Usage: analyze_s10.py CAPTURE OUTDIR
"""

import csv
import fnmatch
import gzip
import json
import os
import statistics
import sys
from collections import defaultdict

HERE = os.path.dirname(os.path.abspath(__file__))
S1_ANALYSIS = os.path.join(HERE, "..", "before", "analysis.json")
ROUTE_MAP = os.path.join(HERE, "..", "s8b", "route-map.json")
EXPLANATIONS = os.path.join(HERE, "explanations.json")

FLOOR_NS = 500
FLOOR_BELOW_NS = 5000
OPEN_BUDGET_NS = 20000
BIMODAL = {
    ("doc_facade_file_open", "FloatingPictures.doc"),
    ("doc_facade_file_full_text", "FloatingPictures.doc"),
}
KEPT = {
    "docx_semantic_one_paragraph", "docx_semantic_one_paragraph_text",
    "docx_producer_source_selected_paragraph",
    "ods_file_eager_selected_cell", "ods_file_source_selected_cell",
    "ods_file_eager_selected_media", "ods_file_source_selected_media",
} | {
    f"detect_{fmt}_{kind}"
    for fmt in ("odt", "ods", "odp")
    for kind in ("bytes", "reader", "polyglot", "catalog_alias")
}


def category(case):
    """S1's category, applied to the predecessor name."""
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


def spread(values):
    if len(values) < 2:
        return None
    return (max(values) - min(values)) / statistics.median(values) * 100


def open_report(path):
    opener = gzip.open if path.endswith(".gz") else open
    with opener(path, "rt") as handle:
        return json.load(handle)


def fixture_names():
    names = {}
    env = json.load(open(os.path.join(HERE, "..", "before", "environment.json")))
    for path, digest in env.get("ole2_fixtures", {}).items():
        names[digest] = os.path.basename(path)
    return names


def corpus_label(corpus, fixtures):
    sha = corpus.get("archive_sha256")
    fixture = fixtures.get(sha)
    return corpus.get("name") + (f":{fixture}" if fixture else "")


def report_evidence(report):
    """Output evidence per (case, corpus name): {field: sorted distinct values}.

    Filesystem routes: the per-sample digest fields of filesystem_evidence
    (XLSX semantic/source digests, DOCX timed full-text digest and length).
    In-process routes: the unified_ops projection digest and observation, the
    OLE2 legs_agree flag, and the pre-M0a facade_ole2 observation (both
    observations are "text:<bytes>:<sha256>" or "paragraphs:<n>"-style).
    """
    out = defaultdict(lambda: defaultdict(set))
    for entry in report.get("filesystem_evidence") or []:
        key = (entry["case"], entry["corpus"].get("name"))
        for sample in entry["samples"]:
            for field, value in sample.items():
                if value is None:
                    continue
                if field.endswith("_sha256") or field.endswith("digest") or field == "docx_timed_full_text_bytes":
                    out[key][field].add(str(value))
    for result in report["results"]:
        key = (result["case"], result["corpus"].get("name"))
        source = result.get("source") or {}
        unified = source.get("unified_ops") or {}
        for field in ("projection_sha256", "observation", "legs_agree"):
            if unified.get(field) is not None:
                out[key][field if field != "observation" else "observation"].add(str(unified[field]))
        facade = source.get("facade_ole2") or {}
        if facade.get("observation") is not None:
            out[key]["observation"].add(str(facade["observation"]))
    return out


def load_capture(root, fixtures):
    """Return rows[(leg, case, corpus label)] -> list of child records."""
    rows = defaultdict(list)
    children = [json.loads(line) for line in open(os.path.join(root, "children.jsonl"))]
    by_stem = {}
    for child in children:
        stem = f"B{child['block']}-{child['position']}{child['leg']}-{child['group']}"
        by_stem[stem] = child
    reports = {}
    raw = os.path.join(root, "raw")
    if os.path.isdir(raw):
        # The full harness reports, as capture_s10.py wrote them.
        for name in sorted(os.listdir(raw)):
            # Attempts capture_s10.py set aside after a tmpfs quota failure.
            if ".corpus-manifest.json" in name or ".quota-failed" in name:
                continue
            if name.endswith(".json.gz"):
                reports[name[:-8]] = os.path.join(raw, name)
            elif name.endswith(".json"):
                reports[name[:-5]] = os.path.join(raw, name)
    for name in sorted(os.listdir(root)):
        # The committed form: prune_reports.py bundles, {stem: pruned report}.
        if name.startswith("pruned-") and name.endswith(".json.gz"):
            for stem, report in open_report(os.path.join(root, name)).items():
                if stem in reports:
                    sys.exit(f"{stem} is both a raw report and in {name}")
                reports[stem] = report
    for stem in sorted(reports):
        child = by_stem.get(stem)
        if child is None:
            sys.exit(f"{stem} has no children.jsonl record")
        report = reports[stem]
        if isinstance(report, str):
            report = open_report(report)
        evidence = report_evidence(report)
        for result in report["results"]:
            corpus = result["corpus"]
            key = (child["leg"], result["case"], corpus_label(corpus, fixtures))
            samples = result["elapsed_ns"]["samples"]
            metrics = (result.get("operation_metrics") or {}).get("process") or {}
            peak_rss = (metrics.get("peak_rss_bytes") or {}).get("values")
            rows[key].append({
                "stem": stem,
                "block": child["block"],
                "position": child["position"],
                "group": child["group"],
                "load_1min_before": float(child["loadavg_before"][0]),
                "p50_ns": p50(samples),
                "samples": len(samples),
                "archive_sha256": corpus.get("archive_sha256"),
                "cache_state": result.get("cache_state"),
                "fresh_child_peak_rss_median_bytes": p50(peak_rss) if peak_rss else None,
                "evidence": {field: sorted(values) for field, values in
                             evidence.get((result["case"], corpus.get("name")), {}).items()},
            })
    return rows, children


def s1_noise():
    analysis = json.load(open(S1_ANALYSIS))
    fixtures = fixture_names()
    noise = {}
    p50s = {}
    for row in analysis["rows"]:
        label = row["corpus"] + (f":{row['fixture']}" if row.get("fixture") else "")
        noise[(row["case"], label)] = row["repeat_p50_spread_pct"]
        p50s[(row["case"], label)] = row["p50_ns"]
    return noise, p50s


def modes(a_values, b_values):
    """Split the pooled per-child p50s at their largest relative gap."""
    pooled = sorted(a_values + b_values)
    best = None
    for low, high in zip(pooled, pooled[1:]):
        gap = (high - low) / low
        if best is None or gap > best[0]:
            best = (gap, (low + high) / 2)
    if best is None or best[0] < 0.15:
        return None
    threshold = best[1]
    out = {}
    for name, lo, hi in (("low", 0, threshold), ("high", threshold, float("inf"))):
        a = [v for v in a_values if lo <= v < hi]
        b = [v for v in b_values if lo <= v < hi]
        entry = {"threshold_ns": threshold, "before_children": len(a), "after_children": len(b)}
        if a and b:
            ma, mb = statistics.median(a), statistics.median(b)
            entry.update({
                "before_median_p50_ns": ma,
                "after_median_p50_ns": mb,
                "ratio": mb / ma,
                "separated": min(b) > max(a),
            })
        out[name] = entry
    return out


def merged_evidence(children):
    merged = defaultdict(set)
    for child in children:
        for field, values in child["evidence"].items():
            merged[field].update(values)
    return merged


def compare_evidence(a_children, b_children):
    """Fields both legs report, and whether every child agreed on one value."""
    a, b = merged_evidence(a_children), merged_evidence(b_children)
    out = {}
    for field in sorted(set(a) & set(b)):
        if field in ("legs_agree",):
            continue
        out[field] = {
            "before": sorted(a[field]),
            "after": sorted(b[field]),
            "equal": a[field] == b[field] and len(a[field]) == 1,
        }
    for field in sorted(set(a) ^ set(b)):
        out[field] = {"before": sorted(a.get(field, ())), "after": sorted(b.get(field, ())),
                      "equal": None}
    agree = sorted(merged_evidence(b_children).get("legs_agree", ()))
    if agree:
        out["legs_agree"] = {"after": agree}
    return out


def compare(pair, before_case, after_case, label, a_children, b_children, noise_pct,
            noise_source, scenario, open_budget=False):
    a = [c["p50_ns"] for c in a_children]
    b = [c["p50_ns"] for c in b_children]
    ma, mb = statistics.median(a), statistics.median(b)
    ratio = mb / ma
    delta = mb - ma
    trigger = max(5.0, noise_pct if noise_pct is not None else 0.0)
    separated = min(b) > max(a)
    over_pct = (ratio - 1) * 100 > trigger
    floor_ok = ma >= FLOOR_BELOW_NS or delta > FLOOR_NS
    separation_needed = noise_pct is not None and noise_pct > 10
    if open_budget:
        budget_ns = max(0.05 * ma, OPEN_BUDGET_NS)
        over = delta > budget_ns
        rule = f"goal open budget: delta <= max(5% of leaf, 20 us) = {budget_ns:.0f} ns"
    else:
        budget_ns = None
        over = over_pct and floor_ok and (separated or not separation_needed)
        rule = (f"S1 rule: trigger {trigger:.1f}%"
                + (", 0.5 us floor" if ma < FLOOR_BELOW_NS else "")
                + (", full separation" if separation_needed else ""))
    row = {
        "pair": pair,
        "scenario": scenario,
        "before_case": before_case,
        "after_case": after_case,
        "corpus": label,
        "before_children": len(a),
        "after_children": len(b),
        "before_child_p50_ns": a,
        "after_child_p50_ns": b,
        "before_p50_ns": ma,
        "after_p50_ns": mb,
        "ratio": ratio,
        "delta_ns": delta,
        "noise_pct": noise_pct,
        "noise_source": noise_source,
        "trigger_pct": trigger,
        "s10_before_spread_pct": spread(a),
        "s10_after_spread_pct": spread(b),
        "separated": separated,
        "over_trigger": over,
        "rule": rule,
        "budget_ns": budget_ns,
        "archive_sha256_match": {c["archive_sha256"] for c in a_children}
        == {c["archive_sha256"] for c in b_children},
    }
    row["outputs"] = compare_evidence(a_children, b_children)
    if (before_case, label.split(":")[-1]) in BIMODAL and pair == "facade":
        row["modes"] = modes(a, b)
        mode_over = False
        if row["modes"]:
            for entry in row["modes"].values():
                if "ratio" in entry and (entry["ratio"] - 1) * 100 > 5.0 and entry["separated"]:
                    mode_over = True
            row["over_trigger"] = mode_over
            row["rule"] = "bimodal: per-mode median ratio above 5% with per-mode separation"
    return row


def main():
    capture, outdir = sys.argv[1], sys.argv[2]
    fixtures = fixture_names()
    rows, children = load_capture(capture, fixtures)
    noise, s1_p50 = s1_noise()
    explanations = json.load(open(EXPLANATIONS)) if os.path.exists(EXPLANATIONS) else {}
    routes = json.load(open(ROUTE_MAP))["routes"]

    def children_of(leg, case, label):
        return rows.get((leg, case, label), [])

    labels = defaultdict(set)
    for (leg, case, label) in rows:
        labels[case].add(label)

    out = []
    missing = []
    for route in routes:
        pred, uni, leaf = route["predecessor"], route["unified"], route["leaf"]
        scenario = route["scenario"]
        for label in sorted(labels.get(pred, set()) | labels.get(uni, set())):
            a = children_of("A", pred, label)
            u = children_of("U", uni, label)
            if a and u:
                pair = "facade" if route["timed_scope_matches_predecessor"] else "budget-only"
                out.append(compare(pair, pred, uni, label, a, u, noise.get((pred, label)),
                                   "S1 spread_all of the predecessor", scenario))
            else:
                missing.append({"pair": "facade", "before_case": pred, "after_case": uni,
                                "corpus": label, "before_children": len(a), "after_children": len(u)})
            if leaf:
                l = children_of("L", leaf, label)
                if l and u:
                    out.append(compare("overhead", leaf, uni, label, l, u, noise.get((pred, label)),
                                       "S1 spread_all of the predecessor", scenario,
                                       open_budget=category(pred) == "open"))
                else:
                    missing.append({"pair": "overhead", "before_case": leaf, "after_case": uni,
                                    "corpus": label, "before_children": len(l), "after_children": len(u)})
    for case in sorted(KEPT):
        for label in sorted(labels.get(case, set())):
            a = children_of("A", case, label)
            u = children_of("U", case, label)
            if a and u:
                out.append(compare("kept", case, case, label, a, u, noise.get((case, label)),
                                   "S1 spread_all", category(case)))
            else:
                missing.append({"pair": "kept", "before_case": case, "after_case": case,
                                "corpus": label, "before_children": len(a), "after_children": len(u)})

    for row in out:
        key = f"{row['pair']}|{row['after_case']}|{row['corpus']}"
        row["category"] = category(row["before_case"] if row["pair"] != "overhead" else
                                   next(r["predecessor"] for r in routes if r["unified"] == row["after_case"]))
        row["s1_p50_ns"] = s1_p50.get((row["before_case"], row["corpus"])) if row["pair"] != "overhead" else None
        # Keys are "<pair>|<after case>|<corpus>" with fnmatch wildcards.
        explanation = next((text for pattern, text in explanations.items()
                            if fnmatch.fnmatchcase(key, pattern)), None)
        if row["pair"] == "budget-only":
            row["verdict"] = "budget-only"
        elif row["pair"] == "overhead":
            row["verdict"] = ("over budget" if row["over_trigger"] else "within budget") \
                if row["budget_ns"] is not None else ("review" if row["over_trigger"] else "within trigger")
        else:
            row["verdict"] = "regression" if row["over_trigger"] else "pass"
        if row["over_trigger"] and explanation:
            row["explanation"] = explanation
            if row["pair"] in ("facade", "kept"):
                row["verdict"] = "explained"
            else:
                row["verdict"] += ", explained"

    # Whole-child peak RSS (GNU time) per group and leg.
    rss = defaultdict(list)
    for child in children:
        kib = child.get("gnu_time", {}).get("peak_rss_kib")
        if kib is not None and child["exit_code"] == 0:
            rss[(child["group"], child["leg"])].append(kib)
    whole_child = [
        {"group": group, "leg": leg, "children": len(values),
         "median_peak_rss_kib": statistics.median(values), "min": min(values), "max": max(values)}
        for (group, leg), values in sorted(rss.items())
    ]
    # Fresh-child peak RSS per filesystem sample (median over a child's samples).
    fresh = []
    for (leg, case, label), recs in sorted(rows.items()):
        values = [r["fresh_child_peak_rss_median_bytes"] for r in recs
                  if r["fresh_child_peak_rss_median_bytes"] is not None]
        if values:
            fresh.append({"leg": leg, "case": case, "corpus": label, "children": len(values),
                          "median_bytes": statistics.median(values)})

    loads = [float(c["loadavg_before"][0]) for c in children] + \
            [float(c["loadavg_after"][0]) for c in children]
    analysis = {
        "schema": "litchi.unified-ops.0003.s10-analysis.v1",
        "rule": {
            "statistic": "median of per-child p50s (each child's p50 = median of its retained samples)",
            "facade_trigger": "max(5%, S1 spread_all); +0.5 us absolute floor below 5 us; "
                              "full separation when S1 spread_all > 10%; bimodal rows per mode",
            "overhead_open_budget": "unified - leaf <= max(5% of leaf, 20 us) for S1-category open rows",
            "overhead_review_trigger": "the facade trigger with the predecessor's S1 noise",
        },
        "children": len(children),
        "children_failed": [
            {"child": f"B{c['block']}-{c['position']}{c['leg']}-{c['group']}",
             "attempt": c.get("attempt", 0), "exit_code": c["exit_code"],
             "quota_failure": c.get("quota_failure")}
            for c in children if c["exit_code"] != 0],
        "binaries": sorted({(c["leg"], c["binary_sha256"]) for c in children}),
        "loadavg_1min_min": min(loads) if loads else None,
        "loadavg_1min_max": max(loads) if loads else None,
        "load_gate": {
            "children_waited": sum(1 for c in children if c.get("load_gate") == "waited"),
            "children_cap_exhausted": sum(1 for c in children if c.get("load_gate") == "cap-exhausted"),
            "waited_total_s": max((c.get("waited_total_s", 0) for c in children), default=0),
        },
        "rows": out,
        "missing": missing,
        "memory": {"whole_child_peak_rss": whole_child, "fresh_child_peak_rss": fresh},
    }
    # Strip the bulky per-child source summaries from the committed analysis.
    with open(os.path.join(outdir, "analysis.json"), "w") as handle:
        json.dump(analysis, handle, indent=1, default=list)
        handle.write("\n")
    with open(os.path.join(outdir, "summary.csv"), "w", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(["pair", "scenario", "category", "before_case", "after_case", "corpus",
                         "before_children", "after_children", "before_p50_ns", "after_p50_ns",
                         "ratio", "delta_ns", "noise_pct", "trigger_pct", "s10_before_spread_pct",
                         "s10_after_spread_pct", "separated", "over_trigger", "verdict"])
        for row in out:
            def fmt(v, digits=2):
                return "" if v is None else f"{v:.{digits}f}"
            writer.writerow([row["pair"], row["scenario"], row["category"], row["before_case"],
                             row["after_case"], row["corpus"], row["before_children"],
                             row["after_children"], fmt(row["before_p50_ns"], 1),
                             fmt(row["after_p50_ns"], 1), fmt(row["ratio"], 4),
                             fmt(row["delta_ns"], 1), fmt(row["noise_pct"]), fmt(row["trigger_pct"]),
                             fmt(row["s10_before_spread_pct"]), fmt(row["s10_after_spread_pct"]),
                             row["separated"], row["over_trigger"], row["verdict"]])
    over = [r for r in out if r["over_trigger"]]
    print(f"{len(out)} rows, {len(children)} children, failed={analysis['children_failed']}, "
          f"missing={len(missing)}, over trigger={len(over)}")
    for r in over:
        print(f"  {r['pair']:9} {r['after_case']:55} {r['corpus']:45} "
              f"{r['before_p50_ns']:12.1f} -> {r['after_p50_ns']:12.1f} x{r['ratio']:.3f} "
              f"trigger {r['trigger_pct']:.1f}% {r['verdict']}")


if __name__ == "__main__":
    main()
