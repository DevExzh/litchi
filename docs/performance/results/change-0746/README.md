# Evidence packet for change 0746

Record: [`docs/performance/0746-xls-validation-only-parse.md`](../../0746-xls-validation-only-parse.md).

A validation-only mode of the complete XLS reader for the `cell_values`,
comments and sheet-visibility edit owners, the validated-render handoff in their
three generic commits, and a determinism fix for two hash-ordered refusals.

## Provenance

| | |
| --- | --- |
| base | `009d515bef` (`feat/office-format-completeness`) |
| branch | `perf/0746-xls-validation-only-parse` |
| commits | `3a2f233cdc` fix, `ec747fbd64` validation-only mode, `fa88d7cc9b` `cell_values` handoff, `3eea4a8bec` comments/visibility adoption, `b85d3e534c` comments/visibility handoff |
| shipped (measured "after") source | `b85d3e534c` |
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
| `probe/main.rs` | the timing and allocation probe (change 0620/0633's, plus `alloc-count`, `--generate-xls-large`, `--reuse-source`, and the `comments-open`, `visibility-open`, `reader-open` operations) |
| `probe/corpus.rs` | the corpus census: every `.xls` through the five `cell_values` paths, the comments and visibility owners (open, generic, source-backed) and a public-reader digest |
| `probe/matrix.rs` | change 0633's 29-case first-error matrix |
| `probe/mutate.rs` | the mutation differential (200 deterministic record-level mutations per fixture, four owners per package) |
| `probe/Cargo-*.toml` | the three legs' manifests |
| `scripts/abba_probe.py`, `scripts/abba_harness.py` | ABBA drivers and summarizers (A B B A A B B A, bootstrap over paired ratios) |
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
| `binaries.sha256`, `environment.txt`, `gates.txt`, `cleanup.json`, `log-sections.md` | identity, host, gates, cleanup and the coordinator's log paragraphs |

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
