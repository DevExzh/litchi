#!/usr/bin/env python3
"""Render the README tables from analysis.json, memory-probe.json and children.jsonl.

usage: render_tables.py DIR CHILDREN_JSONL > tables.md
"""

import json
import os
import statistics
import sys
from collections import defaultdict


def us(ns):
    if ns is None:
        return ""
    value = ns / 1000
    if value >= 1000:
        return f"{value:,.0f}"
    if value >= 100:
        return f"{value:.0f}"
    if value >= 10:
        return f"{value:.1f}"
    return f"{value:.2f}"


def pct(value):
    return "" if value is None else f"{value:.1f}%"


def verdict(row):
    text = row["verdict"]
    if row.get("explanation"):
        text += " (see record)"
    return text


def facade_table(rows, pair):
    out = ["| Scenario | Before selector → after selector | Corpus | n (A/B) | before p50 µs | after p50 µs | ratio | noise (S1) | trigger | S10 spread A / B | verdict |",
           "|---|---|---|---|---|---|---|---|---|---|---|"]
    for row in rows:
        if row["pair"] != pair:
            continue
        names = (f"`{row['before_case']}`" if row["before_case"] == row["after_case"]
                 else f"`{row['before_case']}` → `{row['after_case']}`")
        out.append("| {} | {} | {} | {}/{} | {} | {} | {:.3f} | {} | {} | {} / {} | {} |".format(
            row["category"], names, row["corpus"], row["before_children"], row["after_children"],
            us(row["before_p50_ns"]), us(row["after_p50_ns"]), row["ratio"], pct(row["noise_pct"]),
            pct(row["trigger_pct"]), pct(row["s10_before_spread_pct"]), pct(row["s10_after_spread_pct"]),
            verdict(row)))
    return "\n".join(out)


def overhead_table(rows):
    out = ["| Scenario | Unified selector (vs `_leaf_` twin) | Corpus | n (L/U) | leaf p50 µs | unified p50 µs | ratio | Δ µs | rule | verdict |",
           "|---|---|---|---|---|---|---|---|---|---|"]
    for row in rows:
        if row["pair"] != "overhead":
            continue
        rule = ("open budget max(5%, 20 µs)" if row["budget_ns"] is not None
                else f"review trigger {row['trigger_pct']:.1f}%")
        out.append("| {} | `{}` | {} | {}/{} | {} | {} | {:.3f} | {:+.2f} | {} | {} |".format(
            row["category"], row["after_case"], row["corpus"], row["before_children"],
            row["after_children"], us(row["before_p50_ns"]), us(row["after_p50_ns"]), row["ratio"],
            row["delta_ns"] / 1000, rule, verdict(row)))
    return "\n".join(out)


def memory_probe_table(probe):
    out = ["| Format | Fixture | Route | Input bytes | Held after open: before | after | after, `retain_source(false)` | Δ after − before | Δ / input | Peak during open: before | after |",
           "|---|---|---|---:|---:|---:|---:|---:|---:|---:|---:|"]
    for e in probe["rows"]:
        b, a = e["before"], e["after"]
        r = e.get("after_retain_source_false") or {}
        name = e["fixture"].rsplit("/", 1)[-1]
        if len(name) > 32:
            name = name[:29] + "..."
        if b["open_error"] or a["open_error"]:
            out.append(f"| {e['format']} | {name} | {e['route']} | {b['input_bytes']:,} | refused | refused | refused | | | {b['peak_open']:,.0f} | {a['peak_open']:,.0f} |")
            continue
        released = f"{r['held_after_open']:,.0f}" if r.get("held_after_open") is not None else ""
        out.append("| {} | {} | {} | {:,} | {:,.0f} | {:,.0f} | {} | {:+,.0f} | {:.2f} | {:,.0f} | {:,.0f} |".format(
            e["format"], name, e["route"], b["input_bytes"], b["held_after_open"], a["held_after_open"],
            released, e["delta_held_after_open"], e["delta_held_after_open_over_input"], b["peak_open"],
            a["peak_open"]))
    return "\n".join(out)


def fresh_rss_table(analysis):
    by = defaultdict(dict)
    for entry in analysis["memory"]["fresh_child_peak_rss"]:
        by[(entry["case"], entry["corpus"])][entry["leg"]] = entry["median_bytes"]
    routes = {}
    for row in analysis["rows"]:
        if row["pair"] == "facade" or row["pair"] == "budget-only":
            routes[row["after_case"]] = row["before_case"]
    out = ["| After selector | Corpus | before route (A) MiB | unified (U) MiB | leaf (L) MiB | U − A MiB | U − L MiB |",
           "|---|---|---:|---:|---:|---:|---:|"]
    for (case, corpus), legs in sorted(by.items()):
        if "U" not in legs or case not in routes:
            continue
        leaf_case = case.replace("_unified_", "_leaf_")
        a = by.get((routes[case], corpus), {}).get("A")
        u = legs["U"]
        l = by.get((leaf_case, corpus), {}).get("L")
        mib = 1024 * 1024
        out.append("| `{}` | {} | {} | {:.2f} | {} | {} | {} |".format(
            case, corpus, f"{a / mib:.2f}" if a else "", u / mib, f"{l / mib:.2f}" if l else "",
            f"{(u - a) / mib:+.2f}" if a else "", f"{(u - l) / mib:+.2f}" if l else ""))
    return "\n".join(out)


def whole_child_table(analysis):
    by = defaultdict(dict)
    for entry in analysis["memory"]["whole_child_peak_rss"]:
        by[entry["group"]][entry["leg"]] = entry
    out = ["| Group | A median (min–max) KiB | U median (min–max) KiB | L median (min–max) KiB |",
           "|---|---|---|---|"]
    for group, legs in sorted(by.items()):
        cells = []
        for leg in "AUL":
            e = legs.get(leg)
            cells.append(f"{e['median_peak_rss_kib']:,.0f} ({e['min']:,}–{e['max']:,})" if e else "")
        out.append(f"| {group} | " + " | ".join(cells) + " |")
    return "\n".join(out)


def load_table(children_path):
    by = defaultdict(list)
    for line in open(children_path):
        child = json.loads(line)
        by[(child["group"], child["block"])].append(child)
    out = ["| Group | Block | Children | Start (UTC) | Minutes | 1-min load before each child (min–max) | Gate waits |",
           "|---|---|---|---|---|---|---|"]
    for (group, block), children in sorted(by.items(), key=lambda kv: kv[1][0]["started_utc"]):
        loads = [float(c["loadavg_before"][0]) for c in children]
        minutes = sum(c["seconds"] for c in children) / 60
        waits = sum(c.get("waited_s", 0) for c in children)
        out.append(f"| {group} | {block} | {len(children)} | {children[0]['started_utc'][:19]} | {minutes:.1f} | "
                   f"{min(loads):.2f}–{max(loads):.2f} | {waits} s |")
    return "\n".join(out)


def main():
    root, children = sys.argv[1], sys.argv[2]
    analysis = json.load(open(os.path.join(root, "analysis.json")))
    rows = analysis["rows"]
    print("## Facade comparison\n")
    print(facade_table(rows, "facade"))
    print("\n## Kept predecessors\n")
    print(facade_table(rows, "kept"))
    print("\n## Budget-only successors\n")
    print(facade_table(rows, "budget-only"))
    print("\n## Unified vs leaf\n")
    print(overhead_table(rows))
    probe_path = os.path.join(root, "memory-probe.json")
    if os.path.exists(probe_path):
        print("\n## Allocation probe\n")
        print(memory_probe_table(json.load(open(probe_path))))
    print("\n## Fresh-child peak RSS\n")
    print(fresh_rss_table(analysis))
    print("\n## Whole-child peak RSS\n")
    print(whole_child_table(analysis))
    print("\n## Load per run\n")
    print(load_table(children))


if __name__ == "__main__":
    main()
