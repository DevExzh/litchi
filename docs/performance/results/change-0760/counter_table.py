#!/usr/bin/env python3
"""Render the per-region probe counters of change 0760 as a Markdown table.

Cycle modes have the `setup` mode of their deck subtracted, so each row is
the per-operation count of the timed region alone.
"""

import json
import sys

summary = {}
for path in sys.argv[1:]:
    with open(path, encoding="utf-8") as handle:
        summary.update(json.load(handle)["summary"])

ROWS = [
    ("commit-one-large", "`Transaction::commit`, one edit", "large", None),
    ("commit-one-medium", "`Transaction::commit`, one edit", "medium", None),
    ("commit-pct-large", "`Transaction::commit`, one percent (100 slides rewritten)", "large", None),
    ("capture-large", "`Package::opened_presentation` (capture)", "large", None),
    ("capture-medium", "`Package::opened_presentation` (capture)", "medium", None),
    ("noop-large", "no-op edit/save cycle", "large", "setup-large"),
    ("noop-medium", "no-op edit/save cycle", "medium", "setup-medium"),
    ("one-large", "one-edit edit/save cycle", "large", "setup-large"),
    ("one-medium", "one-edit edit/save cycle", "medium", "setup-medium"),
    ("pct-large", "one-percent edit/save cycle", "large", "setup-large"),
    ("fulltext-large", "`presentation().text()` (control)", "large", None),
    ("fulltext-medium", "`presentation().text()` (control)", "medium", None),
]


def value(mode, leg, event, setup):
    total = summary[mode][leg][event]
    if setup:
        total -= summary[setup][leg][event]
    return total


def fmt(count):
    return f"{count / 1e6:,.2f} M"


print("| region | deck | before instructions | after instructions | change | before cycles | after cycles | change |")
print("| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |")
for mode, label, deck, setup in ROWS:
    if mode not in summary:
        continue
    cells = []
    for event in ("instructions:u", "cycles:u"):
        before = value(mode, "a", event, setup)
        after = value(mode, "b", event, setup)
        cells += [fmt(before), fmt(after), f"{(after / before - 1) * 100:+.2f}%"]
    print(f"| {label} | {deck} | " + " | ".join(cells) + " |")
