# Evidence packet for change 0746

Record: [`docs/performance/0746-xls-validation-only-parse.md`](../../0746-xls-validation-only-parse.md).

A validation-only mode of the complete XLS reader for the `cell_values`,
comments and sheet-visibility edit owners, the validated-render handoff in their
three generic commits, and a determinism fix for two hash-ordered refusals;
then the fixes an independent review asked for (`review-fix/`).

## Provenance

| | |
| --- | --- |
| base | `009d515bef` (`feat/office-format-completeness`) |
| branch | `perf/0746-xls-validation-only-parse` |
| commits | `3a2f233cdc` fix, `ec747fbd64` validation-only mode, `fa88d7cc9b` `cell_values` handoff, `3eea4a8bec` comments/visibility adoption, `b85d3e534c` comments/visibility handoff; after review: `8fa1d4b54a` occupancy by occupied rows, `7f3c6c4b78` frozen multi-defect matrix and cross-tab `kept_cell` test, `3d49e03044` inline current-row path, `0cd989bb15` documentation |
| reviewed (first "after") source | `b85d3e534c` — everything outside `review-fix/` measures it |
| final source | `3d49e03044` code (`0cd989bb15` changes only a doc comment) — measured in `review-fix/` |
| before source | the read-only base checkout `base-009d515bef` |
| fix-only correctness leg | a detached worktree at `3a2f233cdc` |
| host | AMD EPYC 9R45, 32 cores (no SMT), 123 GiB, Linux 7.0.0-1012-aws, shared with other implementers |
| pinning | `taskset -c 20` for every measured process |
| probe toolchain | rustc 1.98.1 (host default), both legs; release, `debug = 1`, no LTO |
| harness toolchain | rustc 1.95.0 (pinned by `rust-toolchain.toml`), both legs |

`binaries.sha256` lists every measured binary. The probe legs differ only in the
path dependency (`probe/Cargo-{before,after,fix}.toml`). The harness after leg is
`tools/perf-baseline` built from `b85d3e534c` with
`cargo build --release --locked --offline --manifest-path tools/perf-baseline/Cargo.toml --bin litchi-perf-baseline`.
Its before leg was measured twice: against the shared prebuilt base binary
(`latency-harness/`) and, after the coordinator's note that the prebuilt binary
shifts untouched paths by 2.7–3.4%, against a base binary built with the
identical command (`latency-harness-selfbuilt/`, which the record quotes).

## Contents

| path | what it is |
| --- | --- |
| `probe/main.rs` | the timing and allocation probe (change 0620/0633's, plus `alloc-count`, `--generate-xls-large`, `--reuse-source`, the `comments-open`, `visibility-open`, `reader-open` operations and, for the review fixes, `--generate-sparse-bands` and `--generate-sparse-bands-minimal`) |
| `probe/corpus.rs` | the corpus census: every `.xls` through the five `cell_values` paths, the comments and visibility owners (open, generic, source-backed) and a public-reader digest |
| `probe/matrix.rs` | change 0633's 29-case first-error matrix |
| `probe/mutate.rs` | the mutation differential (200 deterministic record-level mutations per fixture, four owners per package) |
| `probe/Cargo-*.toml` | the three legs' manifests |
| `scripts/abba_probe.py`, `scripts/abba_harness.py` | ABBA drivers and summarizers (A B B A A B B A; their "bootstrap 95% CI" over four paired ratios equals the ratios' min–max, and the record reports it as that) |
| `scripts/isolate_instructions.py`, `scripts/isolate_harness.py` | isolation pairs under `perf stat` (4 vs 24 iterations) |
| `scripts/perftree.py`, `scripts/perfcallers.py` | frame-pointer call-tree and caller summarizers used for `profiles/` |
| `scripts/gates.sh`, `scripts/final_run.sh` | the gate script and the final measurement sequence |
| `latency-probe/` | shipped build: `summary.json` and the 120 raw probe reports |
| `latency-harness-selfbuilt/` | shipped build against the identically built base harness: `summary.json` and 152 raw reports (quoted) |
| `latency-harness/` | shipped build against the prebuilt base harness: `summary.json` and 152 raw reports |
| `instructions/isolation.json` | per-operation instructions and cycles, both legs, shipped build |
| `instructions-harness*/` | selector isolation pairs (per sample, including untimed per-iteration work) |
| `alloc/` | per-operation allocation counts, two runs per leg |
| `attribution/` | isolation pairs of the generic commit at `ec747fbd64` (validation-only mode, no handoff) |
| `layout/` | the build-sensitivity evidence: four builds in one window, symbol shares and the hottest instructions |
| `profiles/` | summarized frame-pointer call trees of the open and the three commits on `54016.xls`, before and after (the after trees are the `3eea4a8bec` build; its `cell_values` code is the shipped code) |
| `correctness/` | the matrix output of all three legs, one corpus-census output, `output-sha256.txt` for every correctness output of every leg, `mutation-summary.md`, `fixtures.txt` |
| `superseded-run-3eea4a8bec/` | the summaries of the first measurement run, taken on a build of `3eea4a8bec` before commit `b85d3e534c` existed |
| `binaries.sha256`, `environment.txt`, `gates.txt`, `cleanup.json`, `log-sections.md` | identity, host, gates (rerun at the final head), cleanup and the coordinator's log paragraphs |
| `review-fix/run.log` | the final review-fix window: both ABBA summaries as printed, the isolation lines, load averages |
| `review-fix/base-vs-final/`, `review-fix/banded-vs-final/` | ABBA of the base and of the reviewed build against the final source on the sparse-band case, `54016.xls` and `xls-large`: `summary.json` (with the four pair ratios and their min–max) and 104 raw reports each |
| `review-fix/isolation.json` | isolation-pair instructions and cycles of all three legs |
| `review-fix/inline-split-isolation.json` | the reviewed build, the first fix alone (`8fa1d4b54a`) and the final source, isolation pairs on five cases |
| `review-fix/alloc/` | allocation counts of all three legs, two runs each, including the commit operations |
| `review-fix/first-window/` | the first fix alone in a loaded window: ABBA summaries, isolation pairs, allocation counts and transcribed hardware counters (retained, not quoted) |
| `review-fix/correctness/base-multi-defect-matrix.txt` | the worksheet-level multi-defect matrix as generated on the base (three identical runs) |
| `scripts/abba_legs.py`, `scripts/isolate_legs.py`, `scripts/fix_final_run.sh` | the review-fix ABBA driver (min–max of the four pair ratios, no bootstrap), isolation driver and measurement sequence |
| `scripts/gates-final.sh` | the gate script as rerun at the final head, with a fresh target directory and test totals |
| `scripts/base-multi-defect-matrix-generator.diff` | the temporary test, applied to the base checkout together with a copy of `validation_only_tests/multi_defect_cases.rs`, that generated the frozen matrix's expectations; never committed |
| `probe/Cargo-banded.toml` | the review-fix manifest for the reviewed build (a detached worktree at `b85d3e534c`) |

## Reproduction

1. Build the probe for each leg (`probe/Cargo-*.toml` beside `probe/*.rs`, with
   the workspace `Cargo.lock` copied in), once plain and once with
   `--features alloc-count`.
2. `xls_edit_probe --generate-xls-large xls-large.xls` and check SHA-256
   `228c6585a4d26141aebfaf7b08844a2ee445b269d406006a1fdb0484619120fb`;
   `54016.xls` is `test-data/poi/test-data/spreadsheet/54016.xls`
   (`2e050f1fbb31868b097aa6d4d0fe0a16af8e39c252af82d01cd8f93c4f9a911a`).
3. `scripts/final_run.sh` (probe ABBA, isolation, allocation, harness ABBA and
   selector isolation), then `HARNESS_BEFORE=<identically built base harness>
   python3 scripts/abba_harness.py OUT 20`.
4. Correctness: run `xls_error_matrix`, `xls_edit_corpus` over
   `correctness/fixtures.txt` and `MUTATION_ROUNDS=200 xls_mutation_corpus` over
   the same list for each leg and compare SHA-256s with `output-sha256.txt`.

The 9.4 MB mutation outputs and the second copies of identical census outputs
are not retained; their SHA-256s are, and the generators are deterministic.

Review fixes: build the probe for the base, the reviewed build and the final
source with `probe/Cargo-{before,banded,after}.toml` (plain and `alloc-count`),
generate `sparse-bands-minimal.xls` with `xls_edit_probe
--generate-sparse-bands-minimal` (SHA-256
`87b0d72437a279183ed1a9cbb47ee62d787e274756a4323f6da87f55916e54cc`, identical
from every leg) and run `scripts/fix_final_run.sh` (its last lines are the commit-operation
allocation counts, run after the timing). In the review-fix summaries,
`bin/probe-before` is `probe-review-before`, `bin/probe-prefix` is
`probe-review-banded` and `bin/probe-after` is `probe-review-final` in
`binaries.sha256`, except in `review-fix/first-window/`, where
`bin/probe-after` was `probe-review-first-fix`. In
`review-fix/inline-split-isolation.json` the legs `banded`, `fixed` and `split`
are `probe-review-banded`, `probe-review-first-fix` and `probe-review-final`.
