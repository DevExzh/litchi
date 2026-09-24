#!/usr/bin/env python3
"""Render the retained JSON-lines capture as a compact Markdown table."""

from __future__ import annotations

import json
from pathlib import Path


HERE = Path(__file__).resolve().parent
rows = [
    json.loads(line)
    for line in (HERE / "results" / "measurements.jsonl").read_text().splitlines()
    if line.strip()
]

print("| case | phase | p50 / repeat (ns) | p95 batch (ns) | p99 batch (ns) | work / repeat | allocator calls / op | requested bytes / op | retained bytes | RSS (KiB) |")
print("| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |")
for row in rows:
    repeat = row["repeat"]
    print(
        f"| `{row['case']}` | `{row['phase']}` | {row['elapsed_ns_per_repeat']} "
        f"| {row['elapsed_ns_p95']} | {row['elapsed_ns_p99']} | {row['work_per_repeat']} "
        f"| {row['allocator_calls_p50'] / repeat:.1f} "
        f"| {row['requested_bytes_p50'] / repeat:.1f} "
        f"| {row['memory_retained_p50']} | {row['rss_kib']} |"
    )
