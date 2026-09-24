#!/usr/bin/env python3
"""Regenerate the final radix performance report from retained measurements."""
import csv
import json
from pathlib import Path

P = Path(__file__).resolve().parent

def rows(path): return list(csv.DictReader(path.open()))
def load(path): return json.loads(path.read_text())
def fmt(value): return f"{value:,.2f}"

def main():
    old = rows(P / 'baseline/comparable/raw.csv')
    new = rows(P / 'candidate/comparable/raw.csv')
    radix = rows(P / 'candidate/radix/raw.csv')
    base = {(r['phase'], r['case']): r for r in old}
    flags = load(P / 'initial-flags.json')
    repeat = load(P / 'abab/summary.json')
    assert len(old) == len(new) == 141 and len(radix) == 159
    assert all(r['status'] == '0' for r in old + new + radix)
    for r in new:
        a = base[r['phase'], r['case']]
        for k in ['checksum_p50', 'failure', 'expected_success', 'alloc_calls_p50',
                  'requested_bytes_p50', 'peak_live_delta_p50', 'output_reserved_bytes_p50']:
            assert a[k] == r[k], (r['case'], k)
    lines = ["# Radix evaluator performance", "",
        "The final fourteen-function implementation passes the captured semantic preflights. "
        "Existing scalar, logical and bitwise controls retain identical allocation counts, requested "
        "bytes, peak tracked heap, result reservations and checksums. This is a measured feature "
        "addition; no overall speedup is claimed.", "",
        f"The initial screen flagged {len(flags)} of 141 comparable lanes. Four interleaved A/B "
        f"pairs produced {len(flags) * 8} repeat rows. No repeated median-latency delta exceeds 5%. "
        "Seven lanes retain tail-latency flags (5.09–55.10%); one evaluation lane retains a 5.48% "
        "process peak-RSS increase. These remain review limitations, not a blanket no-regression claim.", "",
        "## Setup and provenance", "",
        "The release harness runs serial processes pinned to CPU 2 on the shared AMD EPYC 9R45 "
        "x86_64 host. Each lane has three warmup batches and fifteen measured batches; each batch "
        "contains the recorded repeat count. No Cargo build overlaps the retained measurement runs. "
        "The host is not isolated from unrelated workloads. Environment and toolchain versions are "
        "in [environment.json](environment.json); exact build commands and before/after source hashes "
        "are retained under each revision.", "",
        "| Artifact | SHA-256 |", "| --- | --- |"]
    for variant in ['baseline', 'candidate']:
        b = load(P / variant / 'binary-provenance.json')
        lines.append(f"| {variant} release executable | `{b['sha256']}` |")
    source = load(P / 'candidate/source-sha256.json')['source_after']
    for name in ['crates/litchi-ods/src/codec/formula/evaluation.rs',
                 'crates/litchi-ods/src/codec/formula/evaluation/radix.rs',
                 'crates/litchi-ods/tests/ods_formula_radix_evaluation.rs']:
        lines.append(f"| `{name}` | `{source[name]}` |")
    for line in (P / 'candidate/harness-sha256.txt').read_text().splitlines():
        digest, name = line.split(maxsplit=1)
        lines.append(f"| harness `{name}` | `{digest}` |")
    lines += ["", "Both revisions use exactly the same four harness files. The baseline is commit "
        "`e1976c59d5c4785e0c73a5d27e9349ba082cca2f`. The workspace lock is retained as an explicit "
        "pre-existing input. The final candidate includes the reviewed DECIMAL reservation-order fix: "
        "the temporary TextValue is dropped as a whole before rounding or pushing its numeric result. "
        "Earlier draft expectations and pre-fix captures were superseded before this final comparison.", "",
        "## Corpus and measurement interpretation", "",
        "The comparable corpus has 47 cases across parse, evaluate and parse-evaluate phases "
        "(141 rows per revision). The radix corpus has 53 cases across the same phases (159 rows): "
        "all fourteen functions, accepted lexical spellings, fractional and range errors, resource "
        "and cancellation refusals, large finite values, call-count scales and padding scales. "
        "Radix evaluation is candidate-only because the baseline does not implement this family; "
        "it is not used to claim before/after speedup.", "",
        "The BASE/DECIMAL input-N lanes vary the number of calls, and DEC2HEX-padding-N repeats "
        "a seven-character result; these are not single-literal byte-length scales. BASE-padding-N "
        "varies the actual output length from 64 to 4096. Finite-maximum performance vectors use "
        "hexadecimal; independent Rust tests cover bases 2, 10 and 36 as well.", "",
        "Latency below is the recorded batch p50 divided by repeat (ns/op). Requested heap bytes "
        "and allocation calls are measured by a counting allocator. Peak heap and RSS are batch/process "
        "peaks and are never divided by repeat. RSS includes executable and runtime pages, beyond "
        "tracked live heap. Every retained row returns live heap to its pre-sample level. Formula "
        "error values count as successful evaluations; typed refusals are reported separately.", "",
        "With only fifteen measured batches, p95 and p99 both select the sample maximum. Their "
        "large deltas are finite-sample tail flags rather than reliable population quantiles. "
        "Four paired repeats improve the comparison, but do not prove the remaining flags are noise. "
        "No geometric mean hides individual results.", "",
        "## Repeated review flags", "",
        "[Initial screening deltas](initial-flags.json), [all repeat rows](abab/raw.csv), "
        "and [paired summaries](abab/summary.json) retain every selected lane.", "",
        "| Phase | Case | Median delta % | Tail delta % | RSS delta % |",
        "| --- | --- | ---: | ---: | ---: |"]
    for r in repeat:
        if any((r[m]['delta_pct'] or 0) > 5 for m in ['p50_ns', 'p95_ns', 'p99_ns', 'max_rss_kib']):
            lines.append(f"| {r['phase']} | {r['case']} | {fmt(r['p50_ns']['delta_pct'])} | "
                         f"{fmt(r['p95_ns']['delta_pct'])} | {fmt(r['max_rss_kib']['delta_pct'])} |")
    lines += ["", "## Existing workloads: initial comparable capture", "",
        "The table includes every initial comparable lane; flagged lanes must be interpreted with "
        "the repeats above. Exact allocations, requested bytes, output reservations, checksums, "
        "p95/p99 and commands are in [baseline raw data](baseline/comparable/raw.csv) and "
        "[candidate raw data](candidate/comparable/raw.csv).", "",
        "| Phase | Case | Baseline ns/op | Candidate ns/op | Delta % | Baseline RSS KiB | Candidate RSS KiB |",
        "| --- | --- | ---: | ---: | ---: | ---: | ---: |"]
    for r in new:
        a = base[r['phase'], r['case']]; n = int(r['repeat'])
        delta = (float(r['p50_ns']) / float(a['p50_ns']) - 1) * 100
        lines.append(f"| {r['phase']} | {r['case']} | {fmt(float(a['p50_ns'])/n)} | "
                     f"{fmt(float(r['p50_ns'])/n)} | {fmt(delta)} | {a['max_rss_kib']} | {r['max_rss_kib']} |")
    lines += ["", "## New radix evaluation workloads", "",
        "The following are the evaluate-phase rows; all three phases and all counters are retained "
        "in [radix raw data](candidate/radix/raw.csv). Refusal lanes measure rejection, not successful "
        "calculation. Input bytes count the complete generated formula.", "",
        "| Case | Input bytes | ns/op | Peak tracked heap bytes | RSS KiB | Typed failure |",
        "| --- | ---: | ---: | ---: | ---: | --- |"]
    for r in radix:
        if r['phase'] == 'evaluate':
            lines.append(f"| {r['case']} | {r['input_bytes']} | {fmt(float(r['p50_ns'])/int(r['repeat']))} | "
                         f"{r['peak_live_delta_p50']} | {r['max_rss_kib']} | {r['failure']} |")
    lines += ["", "## Hardware counters and mechanism", "",
        "Five serial perf-stat captures cover the common text-concatenation control in both "
        "revisions and candidate finite-maximum BASE, finite-maximum DECIMAL and 4096-byte padding. "
        "Counters include whole-process setup, preflight and warmups; these totals are not isolated "
        "function costs. Exact commands and binary hashes are in the metadata sidecars.", "",
        "| Capture | Cycles | Instructions | Branches | Branch misses |", "| --- | ---: | ---: | ---: | ---: |"]
    for meta in sorted((P / 'perf-stat').glob('*.metadata.json')):
        m = load(meta); assert m['status'] == 0
        stem = meta.name.removesuffix('.metadata.json')
        c = {r[2]: int(float(r[0])) for r in csv.reader((meta.parent / (stem + '.stderr')).read_text().splitlines()) if len(r) >= 3}
        lines.append(f"| {stem} | {c['cycles']:,} | {c['instructions']:,} | {c['branches']:,} | {c['branch-misses']:,} |")
    lines += ["", "The new arithmetic uses a fixed 32-limb magnitude and fixed reverse-digit storage, "
        "with no heap big-integer dependency. Fallible text allocation retains its explicit reservation "
        "until the result is dropped. The outlined handler adds no evaluator frame variants or hidden "
        "parallel work. These are implementation mechanisms; the retained measurements support only "
        "the listed corpus, not an extrapolated throughput or cross-platform claim.", "",
        "No Amdahl speedup estimate or scaling curve is applicable to this synchronous scalar family: "
        "the batch adds functionality and introduces no worker scheduling. Workbook/reference/array "
        "evaluation and Roman numerals remain outside this batch. Further profiling is needed before "
        "optimizing larger workbook workloads or treating shared-host tail/RSS thresholds as hard gates.", ""]
    (P / 'report.md').write_text('\n'.join(lines))

if __name__ == '__main__': main()
