# Independent audit review: 0727 XLS warm-tail replication

This review owns `audit.py` and this document. The run is diagnostic only. It
replicates the four 0726 native failures with two controls: the matched
`45365-first` cell and the owned `54016-missing-1048576` benefit cell. It does
not retain the 0726 candidate, change production sources, run Cargo or probes,
or turn a rejected result into an acceptance.

## Scope and custody

The freeze binds the complete six-cell plan, the full baseline revision
`959daa11e5c5e9d0efdb13fff814bbcc66284740`, both serialized release binaries,
the exact 0726 candidate source archive, the two source censuses, the probe
source and lockfile, fixtures, and constraints. The independent audit and its
post-capture negative-control receipt are checked directly and included in the
terminal packet seal. The audit independently requires the live checkout to
be restored to the baseline source census. The candidate source is accepted
only when its one changed path is byte-for-byte
equal to `change-0726/candidate-source/crates/litchi-xls/src/workbook/source.rs`.

The terminal audit accepts a removed owned binary only with the exact path,
SHA-256 and byte-size witness in `cleanup.json`. Every raw JSON output and
stderr file remains in the packet and is checked against the capture manifest.
The manifest must contain exactly 324 unique `(cycle, leg, cell, replicate)`
runs and the independently reconstructed `taskset` command for each run.

## Raw evidence and recomputation

Each process must report the expected probe/schema header, fixture hash and
size, mode, coordinate, budget, eight queries, three warmups, 100 fresh owners,
and 100 records with the exact sample ordinals. Every open and query elapsed
time is a non-negative integer. All query outcomes, `all_queries_agree`, and
`agrees_with_first` flags are checked for every record and every process; the
outcome object must remain equal across all runs for a cell.

The audit computes all seven metrics independently for every process:
`open`, `q1`, `q2`, `q3`, `q8`, `q3-to-q8-mean`, and `open-plus-eight`. It keeps
the 100-owner process distribution intact and computes p50, mean, p95, p99,
and maximum without pooling processes or dropping observations. Each pair
compares the same cycle and replicate ordinal. The central B1/A1 and B2/A2
pairs use the frozen five-percent rule and the ten-nanosecond exception only
for q3, q8, and q3-to-q8-mean. The A/A and within-phase pairs are retained as
diagnostic noise controls.

The independent result must equal `analysis.json` at the process-row,
comparison-statistic, outcome, summary, and failed-check levels. This catches
raw-file omission, pooled statistics, replicate mispairing, altered formulas,
and changed pass flags. The ten negative controls also mutate raw hashes,
semantic outcomes, process identity, command headers, phase labels, and end
bindings, and exercise the warm absolute exception and strict workflow rule.

## Captured result

The terminal independent run used `--allow-rejected` so that all rows were
retained for review. It verified 324 processes, 32,400 fresh owners, 259,200
queries, and 1,890 corresponding-replicate comparisons. All custody, headers,
semantic outcomes, per-process statistics, primary-analysis equality, and ten
negative controls passed. The central timing matrix has 60 failed p50/mean
checks, including 14 checks in the four originally failed focus cells; it also
reports 313 paired tail flags. The audit therefore remains `REJECTED` for the
diagnostic timing criterion and does not authorize retention or retroactive
acceptance.

The short terminal log is sealed with SHA-256
`dcc164d76395c58e96b521143b34a5368a99c935d89356886530a421c7972a8f`, and the
stable independent record `audit.json` is sealed with SHA-256
`0c6c635043f469a47be6f09cf922051deafaa15ff5cb526467cc982dd676067a` before
owned-binary cleanup.

## Terminal interpretation

The ordinary command is:

```text
python3 docs/performance/results/change-0727/audit.py
```

It is fail-closed and returns one on any custody, raw-header, outcome,
statistics, negative-control, or central timing failure. If the diagnostic
timing matrix fails, the complete internally verified evidence can be emitted
with:

```text
python3 docs/performance/results/change-0727/audit.py --allow-rejected
```

That mode still returns one and records the failed rows. It has no retention
authority. The plan's explicit disposition and the 0726 rejection remain
binding even if every replication cell passes. No claim about universal XLS
latency, cold devices, remote sources, RSS, concurrency, or non-XLS CRUD is
derived from this bounded run.
