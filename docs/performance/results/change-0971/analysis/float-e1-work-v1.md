# 0971 Float/e1 retained-work analysis

- Status: complete; analysis only; workload/build started: False/False.
- Input: sealed 0970 docfloat/e1/default; 24 children, 12 pinned and 12 unpinned, 80 timed samples each.
- Capture/reduction custody: manifest b0943c457230, stage 244e033250f3, v6 reduction 560e4cae2cd3.

## Timing and whole-process instruction association

Each association below is descriptive across all 12 children in that arm. Pearson and Spearman are child-level correlations; no fast/slow threshold was selected.

| arm | timing target | instruction range | timing range (ns) | Pearson | Spearman |
|---|---|---:|---:|---:|---:|
| pinned | p50 | 6987088875–7494203611 | 978289–1270130 | 0.9552515503939408 | 0.8811188811188811 |
| pinned | p95 | 6987088875–7494203611 | 991918–1317121 | 0.9349908392118896 | 0.8671328671328671 |
| unpinned | p50 | 6988814257–7498294364 | 979605–1267456 | 0.8918102135472826 | 0.8671328671328671 |
| unpinned | p95 | 6988814257–7498294364 | 999579–1323006 | 0.8445932275781732 | 0.7482517482517482 |

The strongest whole-process associations are ranked in the JSON under each arm. The retained report stage/work values are invariant: 80 samples, 32 stage slots per sample, one active stage per sample, three output routes, and stable output bytes/capacity and semantic-change counts. Stage records have capacity fields (before, after, after_release) but no duration or counter.

## Scope and interpretation

- Timed region: public Snapshot::open_bounded source clone, one Edit with all paragraph-0 replacements and retention releases, commit, output Vec copy, and local snapshot/edit/commit destruction; source file read, references, correctness checks, hashing and JSON excluded
- Perf scope: perf measures the launched GNU time command and inherited launcher/baseline process tree; perf bookkeeping is excluded
- GNU time scope: launcher startup plus same-PID execed historical baseline
- Harness scope: baseline Instant samples only; excludes wrappers and sidecar I/O
- These associations do not attribute process-wide instructions to a stage. The report has no subphase timer or work counter from which such a decomposition could be computed.

The machine-readable output is float-e1-work-v1.json; the reproducible helper is tools/analyze_float_e1_work_v1.py.
