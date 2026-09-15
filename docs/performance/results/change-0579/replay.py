#!/usr/bin/env python3
"""Recompute every number change 0579 cites, from this directory alone.

    python3 -B docs/performance/results/change-0579/replay.py

Reads `summary.json`, `chain-steps.json`, `counters/summary.json`,
`callgrind/chain-edges.txt`, `perf/*.csv` and the two corpus logs. Runs no
repository code and needs no build.
"""
from __future__ import annotations

import hashlib
import json
import pathlib
import re
import sys

HERE = pathlib.Path(__file__).resolve().parent
SUMMARY = json.loads((HERE / "summary.json").read_text())
FIXTURES = [
    ("flagship", "ConditionalFormattingSamples.xls"),
    ("cv", "WithCustomViews.xls"),
    ("54016", "54016.xls"),
    ("mid-images", "SimpleWithImages-mac.xls"),
    ("mini-colours", "SimpleWithColours.xls  (mini)"),
    ("mini-checkboxes", "WithCheckBoxes.xls  (mini)"),
    ("mini-extstyles", "WithExtendedStyles.xls  (mini)"),
    ("small-simple", "SimpleMultiCell.xls"),
]
failures: list[str] = []


def check(label: str, condition: bool) -> None:
    print(f"  {'ok  ' if condition else 'FAIL'}  {label}")
    if not condition:
        failures.append(label)


def section(title: str) -> None:
    print(f"\n{title}\n{'-' * len(title)}")


section("1. The control: not one byte of I/O moved")
counters = SUMMARY["counters"]
identical = True
for mode in ("owned-readat", "file-source", "tracked-file"):
    for op in ("open", "list", "one-cell"):
        legs = {}
        for leg in ("before", "after"):
            cell = counters[f"{leg}/{mode}/{op}"]
            tuples = cell["distinct_counter_tuples"]
            identical &= len(tuples) == 1 and tuples[0]["samples"] == cell["samples"]
            legs[leg] = {k: v for k, v in tuples[0].items() if k != "samples"}
        same = legs["before"] == legs["after"]
        identical &= same
        if op == "open" and mode == "owned-readat":
            print(f"  flagship {mode}/{op}: reads={legs['before']['read_calls']} "
                  f"bytes={legs['before']['read_bytes']} "
                  f"version={legs['before']['version_calls']}")
check("all 18 cells identical between legs, and constant across all 100 samples", identical)
check("the flagship open reproduces change 0574's retained 53 reads / 565,201 bytes / 29 version calls",
      counters["before/owned-readat/open"]["distinct_counter_tuples"][0]["read_calls"] == 53
      and counters["before/owned-readat/open"]["distinct_counter_tuples"][0]["read_bytes"] == 565_201
      and counters["before/owned-readat/open"]["distinct_counter_tuples"][0]["version_calls"] == 29)

section("2. Chain steps removed, and instructions, per open (callgrind isolation pairs)")
chain = SUMMARY["chain_steps_and_callgrind"]
print(f"  {'fixture':17} {'chain before':>12} {'after':>7} {'removed':>8} | "
      f"{'Ir before':>11} {'Ir after':>11} {'pct':>8}")
regressions = []
for stem, label in FIXTURES:
    before, after = chain[f"before/{stem}"], chain[f"after/{stem}"]
    pct = 100 * (after["whole_open_ir"] - before["whole_open_ir"]) / before["whole_open_ir"]
    if pct > 0.25:
        regressions.append((stem, pct))
    print(f"  {stem:17} {before['chain_calls_total']:12,.0f} {after['chain_calls_total']:7,.0f} "
          f"{before['chain_calls_total'] - after['chain_calls_total']:8,.0f} | "
          f"{before['whole_open_ir']:11,.0f} {after['whole_open_ir']:11,.0f} {pct:+7.2f}%")
check("the flagship walks 5,796 chain links before and 2,099 after",
      chain["before/flagship"]["chain_calls_total"] == 5796
      and chain["after/flagship"]["chain_calls_total"] == 2099)
check("change 0574's independently captured 5,796 is reproduced exactly",
      chain["before/flagship"]["chain_calls_total"] == 5796)
check("no fixture's instruction count rises by more than 0.25%", not regressions)

section("3. Hardware counters per open (perf stat isolation pairs)")
perf = SUMMARY["perf_stat_per_open"]
print(f"  {'fixture':10} {'mode':13} {'instructions':>22} {'cycles':>20} {'IPC':>14}")
for label in ("flagship", "cv", "54016"):
    for mode in ("owned-readat", "file-source"):
        before, after = perf[f"before/{label}/{mode}"], perf[f"after/{label}/{mode}"]
        di = 100 * (after["instructions"] - before["instructions"]) / before["instructions"]
        dc = 100 * (after["cycles"] - before["cycles"]) / before["cycles"]
        print(f"  {label:10} {mode:13} {before['instructions']:10,.0f}->{after['instructions']:9,.0f} "
              f"{di:+6.2f}% {before['cycles']:8,.0f}->{after['cycles']:8,.0f} {dc:+6.2f}% "
              f"{before['instructions'] / before['cycles']:5.3f}->{after['instructions'] / after['cycles']:5.3f}")
check("instructions and cycles both fall on every measured cell",
      all(perf[f"after/{label}/{mode}"][metric] < perf[f"before/{label}/{mode}"][metric]
          for label in ("flagship", "cv", "54016")
          for mode in ("owned-readat", "file-source")
          for metric in ("instructions", "cycles")))
check("the before leg reproduces change 0576's after leg within 0.2% on the flagship",
      abs(perf["before/flagship/owned-readat"]["instructions"] - 1_391_409) / 1_391_409 < 0.002)

section("4. Wall clock: paired medians, A/B/B/A, two windows")
floors = []
improvements = []
for window in ("window_1", "window_2"):
    print(f"  {window}")
    print(f"    {'fixture':10} {'mode':13} {'before p50':>11} {'after p50':>10} "
          f"{'dir1':>8} {'dir2':>8} {'A/A':>7} {'B/B':>7}")
    for row in SUMMARY["abba"][window]:
        floors += [abs(row["AA"]), abs(row["BB"])]
        improvements += [row["dir1"], row["dir2"]]
        print(f"    {row['fixture']:10} {row['mode']:13} {row['before_p50']:11,.0f} "
              f"{row['after_p50']:10,.0f} {row['dir1']:+7.2f}% {row['dir2']:+7.2f}% "
              f"{row['AA']:+6.2f}% {row['BB']:+6.2f}%")
floor = max(floors)
print(f"  widest same-binary excursion across both windows: {floor:.2f}%")
check("every paired direction is an improvement", all(value < 0 for value in improvements))
check("every improvement is larger than the widest same-binary excursion",
      all(abs(value) > floor for value in improvements))
check("the two directions agree to within 1.7 percentage points in every cell",
      all(abs(row["dir1"] - row["dir2"]) < 1.7
          for window in ("window_1", "window_2") for row in SUMMARY["abba"][window]))
check("logical counters are identical across all four legs of every A/B/B/A cell",
      all(row["counters_identical"]
          for window in ("window_1", "window_2") for row in SUMMARY["abba"][window]))

section("5. The corpus-wide differential")
before_log = (HERE / "corpus.before.txt").read_bytes()
after_log = (HERE / "corpus.after.txt").read_bytes()
rows = [line for line in after_log.decode().splitlines() if line.startswith("CORPUS ")]
total = [line for line in after_log.decode().splitlines() if line.startswith("CORPUSTOTAL")][0]
print(f"  {total}")
print(f"  rows: {len(rows)}  opened: {sum(1 for r in rows if ' ok reads=' in r)}  "
      f"refused: {sum(1 for r in rows if ' refused ' in r)}")
check("every fixture's read ranges, version count and refusal text are byte-identical",
      hashlib.sha256(before_log).hexdigest() == hashlib.sha256(after_log).hexdigest())
check("the retained logs match the digests recorded in summary.json",
      hashlib.sha256(before_log).hexdigest() == SUMMARY["corpus_differential"]["before_sha256"]
      and hashlib.sha256(after_log).hexdigest() == SUMMARY["corpus_differential"]["after_sha256"])

section("6. chain-steps.json agrees with the retained callgrind extract")
edges = (HERE / "callgrind" / "chain-edges.txt").read_text()
found: dict[str, int] = {}
current = None
pending: list[int] = []
for line in edges.splitlines():
    header = re.match(r"^===== tree100-(\S+)\.txt =====$", line)
    if header:
        current = header.group(1)
        found[current] = 0
        pending = []
        continue
    measured = re.match(r"^\s*([\d,]+) \(\s*[\d.]+%\)\s*", line)
    if current is None or not measured:
        continue
    rest = line[measured.end():]
    if rest.startswith("<"):
        calls = re.search(r"\((\d[\d,]*)x\)", rest)
        pending.append(int(calls.group(1).replace(",", "")) if calls else 0)
    elif rest.startswith("*"):
        if "next_chain_sector" in rest:
            found[current] += sum(pending)
        pending = []
    else:
        pending = []

ok = True
for stem, _ in FIXTURES:
    small, large = (20, 220)
    if stem == "cv":
        small, large = 20, 120
    if stem == "54016":
        small, large = 10, 60
    for leg in ("before", "after"):
        key_small, key_large = f"{leg}-{stem}-s{small}", f"{leg}-{stem}-s{large}"
        if key_small not in found or key_large not in found:
            continue
        derived = (found[key_large] - found[key_small]) / (large - small)
        recorded = SUMMARY["chain_steps_and_callgrind"][f"{leg}/{stem}"]["chain_calls_total"]
        ok &= abs(derived - recorded) < 0.5
check("every per-open chain-step figure recomputes from the retained edge extract", ok)

section("7. The rejected first implementation: a retained SharedOleStreamCursor")
rejected = SUMMARY["rejected_cursor_form"]
print(f"  {'leg':22} {'flagship chain':>14} {'instructions':>13} {'cycles':>10} {'IPC':>6}")
legs = [("before", chain["before/flagship"], perf["before/flagship/owned-readat"]),
        ("cursor (rejected)", rejected["chain_steps_and_callgrind"]["cursor/flagship"],
         rejected["perf_stat_per_open"]["cursor/flagship/owned-readat"]),
        ("resumable hint (kept)", chain["after/flagship"], perf["after/flagship/owned-readat"])]
for name, walk, counters in legs:
    print(f"  {name:22} {walk['chain_calls_total']:14,.0f} {counters['instructions']:13,.0f} "
          f"{counters['cycles']:10,.0f} {counters['instructions'] / counters['cycles']:6.3f}")
check("the rejected form removes more chain links than the one that landed",
      rejected["chain_steps_and_callgrind"]["cursor/flagship"]["chain_calls_total"]
      < chain["after/flagship"]["chain_calls_total"])
check("the rejected form removes instructions but adds cycles",
      rejected["perf_stat_per_open"]["cursor/flagship/owned-readat"]["instructions"]
      < perf["before/flagship/owned-readat"]["instructions"]
      and rejected["perf_stat_per_open"]["cursor/flagship/owned-readat"]["cycles"]
      > perf["before/flagship/owned-readat"]["cycles"])
print(f"  {'fixture':10} {'mode':13} {'before p50':>11} {'cursor p50':>11} {'dir1':>8} {'dir2':>8}")
for row in rejected["abba"]:
    print(f"  {row['fixture']:10} {row['mode']:13} {row['before_p50']:11,.0f} "
          f"{row['after_p50']:11,.0f} {row['dir1']:+7.2f}% {row['dir2']:+7.2f}%")
check("the rejected form is slower than the before leg in both directions on every cell",
      all(row["dir1"] > 0 and row["dir2"] > 0 for row in rejected["abba"]))
check("the rejected form's flagship regression exceeds its own same-binary floor",
      min(abs(row["dir1"]) for row in rejected["abba"] if row["fixture"] == "flagship")
      > max(max(abs(row["AA"]), abs(row["BB"])) for row in rejected["abba"]))

print()
if failures:
    print(f"{len(failures)} CHECK(S) FAILED")
    for failure in failures:
        print(f"  - {failure}")
    sys.exit(1)
print("all checks passed")
