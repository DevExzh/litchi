# Evidence packet, change 0613

Record: [`docs/performance/0613-opc-original-audit-memo.md`](../../0613-opc-original-audit-memo.md).

Outcome: **rejected, not implemented.** The memo of the original-bytes audit
verdict was built, tested, gated and measured; it fires zero times on every
reachable scenario, because every publication entry point consumes its package
and duplicate part names are refused inside one publication. The complete
implementation is retained here as a patch. What the measurement did establish
is the price of the audit it would have reused: **27.59% of source-backed XLSX
publication instructions for the original bytes**, 27.60% for the replacement.

## Contents

| Path | What it is |
| --- | --- |
| `decision.json` | The decision, its reason codes, accepted evidence and costs, known gaps and provenance. |
| `gates.txt` | The tail of every gate, run on the candidate tree and again on the committed tree. |
| `log-sections.md` | The four log paragraphs for HOTSPOTS, GOAL_AUDIT, REPORT and ADR_COMPLIANCE, for the coordinator to merge. |
| `binaries.sha256` | SHA-256 of both measured harness binaries. |
| `patch/0613-original-audit-memo.patch` | The complete implementation: the memo, its five internal tests, its three public-contract tests, and the one-line `#[derive(Clone)]` it needs on `xml_minifier::audit::Error`. Applies to `2d6fbeaed` with `git apply`. |
| `counts/audit-call-counts.txt` | `verify_authored` and `is_xml_part` call counts per caller, for all four callgrind runs, with each run's instruction total. |
| `counts/inclusive-ir.txt` | Inclusive instruction costs of publication, the audits and each audit half, for all four runs. |
| `counts/derived.txt` | The per-iteration arithmetic: totals, deltas, and the audit shares of publication. |
| `counts/*.callgrind.zst` | The four raw callgrind profiles (zstd −19). `before`/`after` × `--samples 1`/`--samples 3`. |
| `timing/timing-one-edit.json` | The six timing legs: per-leg n, p50, mean, p95, p99, min, max per corpus shape; the pooled before/after medians; the deltas in both directions; the A/A floor; and the output digest each leg produced. |
| `timing/timing-one-edit.txt` | The same report as printed. |
| `scripts/audit-call-counts.py` | Sums callgrind `calls=` lines per callee, attributed to each caller. |
| `scripts/paired-timing.py` | Runs the A1 B1 B2 A2 A3 A4 sequence and summarizes it. |
| `scripts/replay.sh` | Rebuilds both legs from the base commit and the patch, then reruns every measurement and gate. |

## Provenance

* Base commit: `2d6fbeaed2083de104bbb0b52b2990ce69ac7274` (branch
  `feat/office-format-completeness`).
* Branch: `perf/0613-opc-original-audit-memo`.
* Before leg: the shared read-only checkout at
  `/home/zhuhe/code/litchi-worktrees/before-2d6fbeaed`, built with its own
  `CARGO_TARGET_DIR`. Binary SHA-256
  `1e94ec13b829ef44d5bfc7027cb3fb3ba3b1c14955bd1df8b622995950ba671e`.
* After leg: the candidate worktree with the patch applied. Binary SHA-256
  `f1e42adeb2b20ac6ae92726e55b12e8fb8815981f4ad63e1ae03327c284ca04f`.
* Both legs: `cargo build --release --locked`, same workspace, same toolchain.
* Host: AMD EPYC 9R45, 32 cores, `Linux 7.0.0-1012-aws x86_64`,
  `rustc 1.98.1 (48a229cea 2026-09-01)`, `valgrind-3.26.0`. Seven other agents
  were active on the machine; every measured process was pinned to CPU 9 with
  `taskset`.
* Selector: `xlsx_source_backed_cell_values_one_edit_save`, corpus shapes
  `medium` and `dense-sparse`. One harness iteration publishes both shapes, so
  one iteration is two source-backed publications.

## How to read the counts

Callgrind writes one `calls=` line per call site, so summing them per callee
gives the exact number of times a function ran, attributed to its caller.
Profiling the same case at `--samples 1` and `--samples 3` and differencing the
totals isolates one iteration: `(n3 − n1) / 2`. The eager control path
(`PackageWriter::validate_authored_xml`, 62 calls) is built once before the
sample loop and is constant across both runs, which is how the differencing
separates it.

The headline counts:

| per iteration | before | after |
| --- | ---: | ---: |
| `verify_authored` calls in source-backed publication | 8 | 8 |
| — of which audits of the **original** | 8 | 4 |
| — of which audits of the replacement | (same 8) | 4 |
| memo hits | — | **0** |
| publication instructions | 188,320,847.5 | 187,896,532.0 |
| both audits, instructions | 103,695,957.0 | 103,696,710.5 |

In the before leg both halves go through one symbol, `validate_overlay_xml`, so
they cannot be told apart; the candidate routes the original half through
`validate_original_part_xml`, which is what makes the 27.59% / 27.60% split
measurable at all.

## What is not here

No allocation profile, no `perf stat` cycle counts, no cold-cache or
filesystem-source measurement, no managed-selector timing, and no real-producer
package: publication refuses 94 of 95 of those at the first original audit
(change 0602, D0), so they cannot exercise this path at all.
