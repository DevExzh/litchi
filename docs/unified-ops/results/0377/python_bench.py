#!/usr/bin/env python3
"""Record 0377 (W8): Python-binding routes of the unified API, timed on one
build of ``litchi_py``. The same script runs against the baseline and the
candidate extension (binary-level A-B-B-A by the caller); it uses only API
present in both. Fixtures are created through the API itself, outside the
clock.

Usage: PYTHONPATH=<dir with litchi_py.abi3.so> python3 -I python_bench.py [rounds]
Prints one JSON document: per scenario, the median of ``rounds`` batch means.
"""
import json
import statistics
import sys
import time

import litchi_py as litchi

ROUNDS = int(sys.argv[1]) if len(sys.argv) > 1 else 21


def measure(name, call, calls):
    for _ in range(max(1, calls // 4)):
        call()
    means = []
    for _ in range(ROUNDS):
        start = time.perf_counter_ns()
        for _ in range(calls):
            call()
        means.append((time.perf_counter_ns() - start) / calls / 1000.0)
    return {"scenario": name, "calls": calls, "median_us": statistics.median(means),
            "min_us": min(means), "max_us": max(means)}


def main():
    out = []
    # XLSX: a created workbook with 1,000 numbers in column A.
    book = litchi.create("xlsx")
    book = book.batch([{"command": "set", "path": f"/Sheet1/A{r}", "props": {"value": r}}
                       for r in range(1, 1001)]).file
    xlsx = book.to_bytes()
    out.append(measure("xlsx open_bytes", lambda: litchi.open_bytes(xlsx), 200))
    opened = litchi.open_bytes(xlsx)
    out.append(measure("xlsx set one cell", lambda: opened.set("/Sheet1/B2", {"value": 7}), 200))
    items = [{"command": "set", "path": f"/Sheet1/C{r}", "props": {"value": r}}
             for r in range(1, 1001)]
    out.append(measure("xlsx batch 1k cell sets", lambda: opened.batch(items), 10))
    out.append(measure("xlsx get one cell", lambda: opened.get("/Sheet1/A500"), 500))
    # DOCX: a created document of 200 paragraphs.
    doc = litchi.create("docx")
    doc = doc.batch([{"command": "add", "parent": "/body", "type": "paragraph",
                      "props": {"text": f"Paragraph {i}"}} for i in range(200)]).file
    docx = doc.to_bytes()
    out.append(measure("docx open_bytes", lambda: litchi.open_bytes(docx), 200))
    opened = litchi.open_bytes(docx)
    out.append(measure("docx set paragraph text",
                       lambda: opened.set("/body/p[100]", {"text": "changed"}), 100))
    out.append(measure("docx to_bytes", lambda: opened.to_bytes(), 200))
    print(json.dumps({"schema": "litchi.unified-fin-perf.python.v1", "record": "0377",
                      "rounds": ROUNDS, "rows": out}))


if __name__ == "__main__":
    main()
