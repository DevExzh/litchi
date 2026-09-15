# Evidence: change 0601, a real-producer shape in the harness

Change record:
[`0601-perf-harness-real-producer-shape.md`](../../0601-perf-harness-real-producer-shape.md).

Disposition: retained, harness and corpus work. `performance_claim: none`.
**No file under `crates/` was modified.** Every number here is descriptive: it
is the first measurement of the XLSX open, selected-cell, planning and one-cell
edit/save scenarios, and of the DOCX one-paragraph and PPTX one-slide
scenarios, on input that carries what real Office producers write.

## Contents

| Path | What it is |
| --- | --- |
| `runs/provenance.txt` | Host, CPU, UTC timestamp, base commit, branch, `rustc`, the measured binary's SHA-256 and size, and the SHA-256 of the real fixture. |
| `runs/determinism-{1,2}.json` | Two independent release-mode harness runs at 1 sample, used only to establish that generation is deterministic. |
| `runs/evidence-{1,2}.json` | The producer-shape marker and refusal census emitted by those two runs (`--producer-evidence`), schema `litchi.perf-baseline.producer-shape-evidence.v1`. |
| `runs/determinism.diff` | `diff -u` of the two census sidecars. **Empty.** |
| `runs/corpus-identity-{1,2}.json` | Archive bytes, member count, part bytes, target entry and SHA-256 per corpus, extracted from the two runs. |
| `runs/corpus-identity.diff` | `diff -u` of those two. **Empty.** |
| `runs/baseline-{a,b}.json` | The two identical timing legs: 20 warm-ups, 100 samples, fourteen generated selectors, pinned to CPU 23, taken back to back in one window. `b` is the A/A control for `a`. |
| `runs/real-file.json` | The same open and selected-cell scenarios over `test-data/poi/test-data/spreadsheet/Excel_file_with_trash_item.xlsx` through `--real-file`, 20 warm-ups and 100 samples. |
| `runs/real-file-evidence.json` | The census of that real file, produced by the same code path as the generated census. |
| `runs/summary.json` | Per-case p50/p95/p99/mean for both legs, the per-case A/A p50 drift, and the run's floor. |
| `scripts/run-baseline.sh` | The whole capture, reproducible: provenance, the two determinism legs, the A/A pair, the real-file leg. |
| `scripts/corpus_identity.py` | Extracts the corpus identities from one report for the determinism diff. |
| `scripts/summarize.py` | Builds `summary.json` and prints the table the record quotes. |
| `gates.txt` | The tail of every gate run in the worktree. |
| `log-sections.md` | The four log paragraphs for `HOTSPOTS.md`, `GOAL_AUDIT.md`, `REPORT.md` and `ADR_COMPLIANCE.md`, for the coordinator to merge. |
| `decision.json` | The `litchi-perf-change-decision` record. |
| `cleanup.json` | What was removed from the session scratchpad after this packet was assembled. |

## How to reproduce

```
git -C <repo> worktree add -b <branch> <path> perf/0601-perf-harness-real-producer-shape
cargo build --release --locked --manifest-path <path>/tools/perf-baseline/Cargo.toml
<path>/docs/performance/results/change-0601/scripts/run-baseline.sh <path> <out> 23
python3 <path>/docs/performance/results/change-0601/scripts/summarize.py <out>
```

The determinism legs and the census are reproducible byte for byte on any host:
generation is a pure function of the shape. The timing legs are not, and are
not meant to be; what reproduces is the *ratio* against the marker-free control
captured in the same run.

## Provenance

Base commit `f8cf7d2a1d155615ab97eda6310c3d5ced0b82ae`, branch
`perf/0601-perf-harness-real-producer-shape`. Host: AMD EPYC 9R45, 32 cores,
Linux 7.0.0-1012-aws, rustc 1.95.0. Every measured process was pinned to CPU 23
with `taskset`. Binary SHA-256 and size are in `runs/provenance.txt`.

Seven other measurement agents were building and measuring on the same host
during the capture. That is why the A/A leg exists and why it is reported per
case rather than as one number: it is the only honest statement of the floor
under those conditions.

## What this packet does not establish

- No speedup, regression, allocation, peak-RSS, cold-cache, physical-I/O,
  range-source, concurrency or cross-platform result. Nothing was optimized.
- No instruction-level attribution. No callgrind pair and no hardware counters
  were taken for this batch; the record says where they belong next.
- No claim that the generated producer shape is representative of the
  real-producer *population*. It is shown to trip the same library gates as one
  real fixture, part for part, by the same code path. The breadth evidence
  remains change 0587's fixture census.
- No claim about the DOCX or PPTX numbers as ratios: no marker-free control
  variant was built for either format, so they are baselines only.
