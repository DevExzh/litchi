#!/usr/bin/env python3
"""Regenerate the final report from retained, independently verified captures."""
import csv
import json
from pathlib import Path

HERE = Path(__file__).resolve().parent


def load(path):
    return json.loads(path.read_text())


def rows(path):
    with path.open() as stream:
        return list(csv.DictReader(stream))


def main():
    baseline = rows(HERE / 'baseline/comparable/raw.csv')
    candidate = rows(HERE / 'candidate/comparable/raw.csv')
    roman = rows(HERE / 'candidate/roman/raw.csv')
    flags = load(HERE / 'initial-flags.json')
    repeated = load(HERE / 'abab/summary.json')
    assert all(r['status'] == '0' for r in baseline + candidate + roman)
    assert len(baseline) == len(candidate) == 159 and len(roman) == 123
    regressions = [r for r in repeated if r['p50_ns']['delta_pct'] > 5]
    tails = [r for r in repeated if r['p95_ns']['delta_pct'] > 5 or r['p99_ns']['delta_pct'] > 5]
    spec = load(HERE.parent / 'specification.json')
    text = [
        '# ROMAN/ARABIC: final measured function-family extension',
        '',
        'The candidate implements both remaining OpenFormula §6.19 functions and removes a per-byte cancellation-threshold comparison from string scanning. Checks remain bounded by 4096-byte windows, including a doubled quote crossing a window by one byte. This is a necessary measured feature extension, with explicit remaining performance qualifications; it is not an across-the-board speedup or no-regression claim.',
        '',
        f'Baseline: `{spec["base_commit"]}`. Candidate source is bound by [the exact patch](../candidate.patch), [gate hashes](../gates/results.json), and both build receipts. The scoped isolated gates pass 926 tests in 56 targets, two doctests, Clippy, rustdoc, formatting and crate boundaries. A separate main-worktree gate run also passed; unrelated root edits were preserved.',
        '',
        '## Method and reproducibility',
        '',
        'The final corpus contains 53 existing-function controls and 41 Roman/Arabic cases, each in parse, evaluate, and combined phases. All 441 initial rows pass exact result/refusal checks. Every phase uses three warmup batches and 15 timed batches. Repetition counts are fixed per case in [run.py](roman-harness/run.py); scaled inputs span 64, 256, 1024 and 4096 bytes/calls.',
        '',
        'Final children are pinned to CPU 6 on the shared Linux host described in [environment.json](environment.json). No hard scheduling isolation is claimed. An unrelated benchmark occupied CPU 2 during exploration; the final complete comparison uses CPU 6. Baseline and candidate use the same four harness files, locked offline release builds and identical Cargo inputs. The runner affinity is the only change from the independently reviewed original harness, as verified by [final-harness-affinity.json](final-harness-affinity.json).',
        '',
        'Timing excludes process launch, setup, preflight and evaluation-context construction. Evaluate mode reuses an immutable parsed expression; combined mode parses and evaluates per operation. Result destruction and allocator observation follow the harness contract. RSS is whole-process maximum RSS, including startup and preflight. Allocator counters include only the measured tracking interval; peak heap and RSS are not divided by repeat count. Hardware counters below cover the whole subprocess, including setup and warmups.',
        '',
        'Displayed ns/op divides a recorded batch median by its repeat count. The four-repeat comparison is a ratio of per-revision medians across interleaved A/B pairs. With only 15 batches, p95/p99 both select the maximum batch; these are diagnostic tail flags, not estimates of production service percentiles. There is no confidence-interval, broad throughput, package-memory, cold-I/O or multicore-scaling claim.',
        '',
        'To reproduce, materialize the base commit in the scratch path configured by [capture.py](capture.py), copy the retained [workspace lock](../gates/workspace-Cargo.lock), gate collector and four harness files, then run `python3 -B performance/capture.py baseline` from this evidence directory. Apply `candidate.patch` from the checkout root, run `gates/run.py` in the isolated candidate and retain its receipt here, then run `performance/capture.py candidate`. The exact release build commands, environment overrides and per-case commands are retained beside each capture. Run `performance/select_flags.py` to generate the >5% flag set from the four initial latency/RSS metrics, then run `performance/repeat_flags.py`, `performance/counters.py` and `performance/investigate_counters.py` serially. Recreate these disposable roots only for a new measurement run; final cleanup intentionally removes them.',
        '',
        '## Comparable results and regression disposition',
        '',
        'Every comparable result, checksum, refusal kind, allocation count, requested-byte count, peak tracked heap and output reservation matches the baseline. All individual results remain in [baseline CSV](baseline/comparable/raw.csv), [candidate CSV](candidate/comparable/raw.csv), and [four-pair CSV](abab/raw.csv). This exact allocation parity is scoped to these controls and does not establish package-wide RSS behavior.',
        '',
        f'All {len(flags)} initial scenarios exceeding 5% in p50/p95/p99 or RSS received four A/B pairs ({len(flags) * 8} rows). {len(regressions)} combined-phase p50 regressions remain. They are accepted as explicit qualifications of this necessary function-family extension under the review-trigger rule in `docs/GOAL.md`; follow-up should target the coercion/error paths and code-layout sensitivity using representative formula workloads. No remaining regression is dismissed as noise. No repeated RSS median exceeds 5%.',
        '',
        '| Phase/case | Baseline ns/op | Candidate ns/op | Repeated p50 change |',
        '| --- | ---: | ---: | ---: |',
    ]
    for item in regressions:
        metric = item['p50_ns']; repeat = item['repeat']
        text.append(f'| {item["phase"]} / {item["case"]} | {metric["baseline"] / repeat:.2f} | {metric["candidate"] / repeat:.2f} | {metric["delta_pct"]:+.2f}% |')
    text += ['', f'All {len(tails)} remaining tail-flag lanes are disclosed below. Full repeated results, including improvements and flags that cleared, are in [summary.json](abab/summary.json).', '',
             '| Phase/case | p95 change | p99 change |', '| --- | ---: | ---: |']
    for item in tails:
        text.append(f'| {item["phase"]} / {item["case"]} | {item["p95_ns"]["delta_pct"]:+.2f}% | {item["p99_ns"]["delta_pct"]:+.2f}% |')
    text += ['', '## New-function costs and scaling', '',
             'These candidate-only timings have no supported baseline implementation. All five formats, error/refusal cases, text scanning and concatenated calls are retained in [Roman CSV](candidate/roman/raw.csv). The following evaluate-only rows show individual calls and every scaling point; repeat batches amortize timing overhead without dividing peak heap.', '',
             '| Case | ns/op | Peak tracked heap bytes | Output reservation bytes |', '| --- | ---: | ---: | ---: |']
    for row in roman:
        if row['phase'] == 'evaluate' and (row['case'].startswith(('roman-3888-format-', 'arabic-input-', 'roman-concat-')) or row['case'] in ['arabic-empty', 'arabic-indirect']):
            text.append(f'| {row["case"]} | {float(row["p50_ns"]) / int(row["repeat"]):.2f} | {row["peak_live_delta_p50"]} | {row["output_reserved_bytes_p50"]} |')
    text += ['', 'ARABIC scans input linearly with a fixed i128 accumulator. ROMAN format 4 examines 64 fixed candidates and uses stack arrays rather than a 4000-entry heap table. Concatenation scaling includes the existing repeated-copy behavior of the scalar concatenation operator; it is not a claim of a linear rope or streaming implementation. The functions perform no workbook access, decompression, recompression or I/O and introduce no threads or caches.', '',
             '## Hardware counters and scanner investigation', '',
             'Five primary captures cover an existing text control before/after, classic ROMAN, shortest ROMAN, and long ARABIC input. Four additional captures investigate the prior UTF-8/coercion regressions with larger repeat batches. Counter totals are process totals and are not interchangeable with ns/op or production tail latency.', '',
             '| Capture | Evaluate ns/op | Cycles | Instructions | Branches | Branch misses |', '| --- | ---: | ---: | ---: | ---: | ---: |']
    for folder in ['perf-stat', 'investigation-counters']:
        for receipt_path in sorted((HERE / folder).glob('*.metadata.json')):
            stem = receipt_path.with_name(receipt_path.name.removesuffix('.metadata.json'))
            config, result = [dict(token.split('=', 1) for token in line.split()[1:]) for line in stem.with_suffix('.stdout').read_text().splitlines()]
            counters = {r[2]: int(float(r[0])) for r in csv.reader(stem.with_suffix('.stderr').read_text().splitlines()) if len(r) >= 3}
            text.append(f'| {folder}/{stem.name} | {int(result["p50_ns"]) / int(config["repeat"]):.2f} | ' + ' | '.join(str(counters[key]) for key in ['cycles', 'instructions', 'branches', 'branch-misses']) + ' |')
    text += ['', 'The larger-batch UTF-8 lane measures about 10.87→10.09 µs/op and bitwise text coercion about 584→594 ns/op. These evaluate-only measurements do not erase the two combined-phase A/B regressions. The finite scan-window change removes real branch work; blanket dispatch/scanner outlining did not robustly improve the controls and was discarded. No cold-path annotation or padding was adopted.', '',
             'Original captures, counter investigation, the CPU 6 confirmation probe and rejected experiments remain under [exploration](exploration/). Source and harness bytes for the original candidate were reconstructed and checked against the original receipts. Exploratory results are not substituted for the final source-bound CPU 6 captures. Shared-host interference and code placement limit causal attribution of small deltas.', '',
             '## Verification and cleanup', '',
             'The root verifier checks all raw stdout/status/time sidecars, command affinity, exact output/heap parity, all flag selection and repeated medians, source and harness custody, five primary and four investigative counter captures, and binary hashes. Both final ELFs and all three scratch roots were removed after review and measurement checks. [Cleanup](../gates/cleanup.json) records 1,555,263,488 allocated bytes reclaimed and hash-verified recovery of 22 unique files outside tmpfs. No broader formula or recalculation gap is declared closed.', '']
    (HERE / 'report.md').write_text('\n'.join(text))


if __name__ == '__main__':
    main()
