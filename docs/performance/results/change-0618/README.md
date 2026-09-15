# Evidence packet: change 0618

Drive the Deflate compressor from the ZIP writer at both Office write sites, so
one authored save constructs one compressor instead of one per member, and one
regenerated preservation member no longer pays flate2's per-call output-buffer
zeroing.

Record: [`../../0618-zip-writer-deflate-state-reuse.md`](../../0618-zip-writer-deflate-state-reuse.md).
Implements SAVE-4 of [0587](../../0587-remaining-opportunity-survey.md), for the
two sites change 0607 identified.

## Provenance

| field | value |
| --- | --- |
| base commit (before leg) | `8fe9efa55728cf8b9592934f9e9de6f5a66fdc1b` |
| branch | `perf/0618-zip-preservation-writer-deflate-reuse` |
| host | AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws |
| toolchain | rustc 1.95.0 (59807616e 2026-04-14), cargo 1.95.0, valgrind 3.26.0 |
| build | `cargo build --release --locked` (workspace `[profile.release]`: `lto = true`, `panic = "abort"`); the scratch probes add `debug = 1` |
| pinning | every measured process `taskset -c 12` |
| host state | seven other agents were building and measuring concurrently; run-window load average 20-25 (`timing/window.txt`) |

Binary SHA-256:

| binary | sha256 |
| --- | --- |
| before `probe0618-opc` | `894381068256d8f54a13ff2edd7d69177d00137b13bfc5bc2c2d337e91993e98` |
| after `probe0618-opc` | `48c6e782434f248f813e2b595fbdb055ff20a0deebb14904f2ca3a818fae77c3` |
| before `probe0618-pptx` | `02c592b4699b0f750f3e3d5ab3d89c0ff129766ed6883006520e9e6dcfcdb6fb` |
| after `probe0618-pptx` | `abc26d4ca921f1fd8388afefd550471226bdce609fdf0208b90db48440a421c7` |
| before `examples/tabs` | `7459fdc5e00529a2f6f67cf3d9cc478ec966d01c4bee8d2e549a7d55665b702a` |
| after `examples/tabs` | `348575995243f9579cdd5a8ef4212e5ede8b75f195c6872c96a69e9f4ebdde29` |
| before `examples/edit_cells` | `0936050390914a220ef9ab30eca53b5b07e2f8af9feca70e7cad3b795b57a398` |
| after `examples/edit_cells` | `eb6f69965d701b7e0d046fae68b02a910defc681df7b93790edbf10b3c1480c5` |
| before `examples/append_plain_paragraph` | `751d0a0d4d202125d4dd09e9701f5f1772c56f6fda3e0df9a8f2e162eb9f64a4` |
| after `examples/append_plain_paragraph` | `c4da788b81311b5ca157b814bb4451469d5a5283e54051c660b11a3cf0a6382f` |

## Contents

| path | what it is |
| --- | --- |
| `probe-opc/src/main.rs` | the scratch publication probe; extends change 0593's probe with the `addrelN:<k>` and `addrelall` multi-member regeneration scenarios `tools/perf-baseline` has no selector for |
| `probe-opc/Cargo.toml.example` | its manifest template; point the `litchi-opc` path dependency at the leg being measured |
| `probe-pptx/src/main.rs` | change 0607's PPTX probe, reused unchanged: `create`/`resave` drive the authored save through the streaming archive writer |
| `probe-pptx/Cargo.toml.example` | its manifest template |
| `counts/callgrind-isolation-summary.txt` | per-operation Ir for both legs and every scenario, with the caveat that callgrind prices the compressor's state zeroing per byte |
| `counts/callgrind-raw-totals.txt` | the raw `I refs` totals behind those pairs |
| `counts/deflate-construction-counts.txt` | Deflate-compressor constructions and resets per save, before and after |
| `counts/perf-isolation-summary.txt` | the native price: `perf stat` cycles, instructions and minor page faults per operation |
| `counts/perf-raw.txt` | the raw `perf stat -x, -r 5` rows behind it |
| `counts/run-counts.sh`, `counts/run-perf.sh` | the two count runners |
| `counts/deflate-construction-counts.py`, `counts/callgrind-callers.py` | the callgrind call-count and caller-attribution extractors |
| `timing/opc-addrelall/`, `timing/opc-addrelall-repeat/` | per-sample publish nanoseconds for a 40-member regeneration, A1 B1 B2 A2, 2,000 samples per leg-run, two windows |
| `timing/opc-addrel/`, `timing/opc-addrel-repeat/` | the same for the single-member control |
| `timing/pptx-create50/` | mean nanoseconds per authored 50-slide save, A1 B1 B2 A2, 40 samples per leg-run |
| `timing/timing.sh`, `timing/stats.py` | the timing runner and the percentile/A-A-floor calculator |
| `timing/window.txt` | wall-clock start and end of each timing window with the host load average |
| `oracle/opc-corpus-*.txt` | 336 OOXML fixtures x 4 scenarios = 1,344 rows, each the published length and SHA-256 or the typed error |
| `oracle/opc-corpus-multi-*.txt` | the same 336 fixtures x the 3 multi-member scenarios = 1,008 rows |
| `oracle/editor-*.txt` | 422 rows from the real `tabs`, `edit_cells` and `append_plain_paragraph` routes over the whole corpus |
| `oracle/pptx-authored-*.txt` | authored-save digests at 1, 2, 5, 10, 50 and 200 slides, twice each (stability and one-edit) |
| `oracle/pptx-opened-*.txt` | 78 opened decks saved back, exact-source passthrough |
| `oracle/*-diff.txt` | empty: every oracle pair is byte-identical |
| `oracle/editor-oracle.sh` | the editor-level oracle runner |
| `gates.txt` | the tail of every gate |
| `run-gates.sh` | the gate runner that produced it |
| `decision.json` | the machine-readable decision, accepted evidence, accepted costs and known gaps |
| `log-sections.md` | the four log paragraphs for HOTSPOTS, GOAL_AUDIT, REPORT and ADR_COMPLIANCE |

## Replaying

The before leg is a detached checkout of the base commit; the after leg is the
branch worktree. Build each probe against its own leg with its own
`CARGO_TARGET_DIR`, then run `counts/run-counts.sh`, `counts/run-perf.sh` and
`timing/timing.sh` with the scratch directory as their only argument. The
fixture in every measured scenario is
`test-data/ooxml/xlsx/ConditionalFormattingSamples.xlsx` (132 members, 654,688
bytes, 40 parts carrying relationships) for the preservation writer, and an
authored deck for the streaming writer.
