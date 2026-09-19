#!/usr/bin/env python3
"""Check retained bindings/differentials and derive the scoped comparison table."""
import hashlib
import json
from pathlib import Path

here = Path(__file__).resolve().parent
root = here.parents[3]


def read(path):
    return json.loads((here / path).read_text())


def rows(path):
    return [json.loads(line) for line in (here / path).read_text().splitlines()]


def key(row):
    return row["route"], row["mode"], row["input_sha256"], row["repetitions"]


manifest = read("measurements/after/manifest.json")
assembly = manifest["assembly"]
assert hashlib.sha256((here / assembly["path"]).read_bytes()).hexdigest() == assembly["sha256"]
for relative, expected in manifest["source_sha256"].items():
    assert hashlib.sha256((root / relative).read_bytes()).hexdigest() == expected, relative
for relative, expected in manifest["probe_sha256"].items():
    assert hashlib.sha256((here / "probe" / relative).read_bytes()).hexdigest() == expected, relative
for relative, expected in manifest["raw_sha256"].items():
    assert hashlib.sha256((here / relative).read_bytes()).hexdigest() == expected, relative

before = {key(row): row for row in rows("measurements/before/matrix.jsonl")}
after = {key(row): row for row in rows("measurements/after/matrix.jsonl")}
assert before.keys() == after.keys()
for case, left in before.items():
    right = after[case]
    for field in ("result_digest", "expected_paragraphs", "measured_iterations", "archive_bytes", "xml_bytes"):
        assert left[field] == right[field], (case, field)
    if left["source_diagnostics"]:
        for field in ("cold_loads", "hits", "successful_loads", "retained_entries", "retained_bytes"):
            assert left["source_diagnostics"][field] == right["source_diagnostics"][field], (case, field)

for stream in ("abba/abba", "control-aa/control", "control-abba/control"):
    digests = {}
    for row in rows("measurements/" + stream + ".jsonl"):
        digests.setdefault(key(row), set()).add(row["result_digest"])
    assert all(len(values) == 1 for values in digests.values()), stream

aa = read("measurements/before/aa-summary.json")["paired"]
paired = read("measurements/abba/summary.json")["paired"]
fresh = []
for case, values in paired.items():
    if case.split("/")[1] != "fresh-count":
        continue
    a, b = values["A1"], values["B1"]
    fresh.append(dict(
        route=case.split("/")[0], corpus=Path(a["corpus_labels"][0]).name,
        repetitions=5,
        a1_p50_ns=a["elapsed_ns"]["p50"], b1_p50_ns=b["elapsed_ns"]["p50"],
        b1_vs_a1_percent=values["before_after_p50_percent"],
        b2_vs_a2_percent=values["after_before_p50_percent"],
        aa_p50_percent=aa[case]["p50_delta_percent"],
        allocation_calls_before=a["allocations"]["p50"],
        allocation_calls_after=b["allocations"]["p50"],
        allocated_bytes_before=a["allocated_bytes"]["p50"],
        allocated_bytes_after=b["allocated_bytes"]["p50"],
    ))

result = dict(
    performance_claim="none; scoped observations, not registered claims",
    source_files_verified=len(manifest["source_sha256"]),
    raw_streams_verified=len(manifest["raw_sha256"]),
    matched_matrix_cases=len(before),
    primary_abba_samples=len(rows("measurements/abba/abba.jsonl")),
    fresh_count=fresh,
    control_aa=read("measurements/control-aa/summary.json")["paired"],
    control_abba=read("measurements/control-abba/summary.json")["paired"],
)
(here / "comparison.json").write_text(json.dumps(result, indent=2) + "\n")
print("Verified source/probe/raw bindings and", len(before), "matrix cases; wrote comparison.json")
