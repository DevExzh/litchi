# 0806 cross-format regression-control plan

The 0805 combined candidate is shared by the DOCX and XLSX reader graphs
through `litchi-opc`/`litchi-ooxml-common`, so the 0805 micro-input preflight
does not cover its consumer surface. A small fresh control can reuse the
existing `tools/perf-baseline` selectors without adding a benchmark or
changing a production crate. The 0806 before leg is exact production at
`e3ff267ee3454e71d66f177f54f3cd05e0d9cce5`; the after leg is the exact
archived 0805 candidate applied temporarily by the root coordinator. Source
and lane-adoption custody are recorded separately; this lane alone cannot
adopt the candidate or advance the global workflow decision. The 0806 main
packet retains the inherited benefit/resource policy and owns that decision.

## Selected matrix

Use the ordinary release `litchi-perf-baseline` binary with this exact case and
shape selection:

```text
--case docx_semantic_open,docx_semantic_full_text,xlsx_open_owned,xlsx_full_cell_scan
--semantic-shape tiny,large
--xlsx-shape tiny,dense-wide
--warmup 3
--samples 30
```

The selection produces eight result rows in each report:

| format | selectors | shapes | deterministic input sizes |
| --- | --- | --- | --- |
| DOCX | `docx_semantic_open`, `docx_semantic_full_text` | `tiny`, `large` | 24 and 10,000 paragraphs |
| XLSX | `xlsx_open_owned`, `xlsx_full_cell_scan` | `tiny`, `dense-wide` | 3 sheets × 8 × 8 cells and 2 sheets × 256 × 256 cells |

Run six paired blocks, with one fresh process for each leg of each block. Keep
the leg order exactly:

```text
block 1: A B
block 2: B A
block 3: A B
block 4: B A
block 5: B A
block 6: A B
```

Here A is the before binary and B is the candidate binary. Each process report
has eight rows × 30 measured samples, so the paired campaign has 12 reports and
2,880 measured samples. One before-only qualification invocation verifies all
eight rows with one sample each before the candidate is applied; its rows are
qualification evidence and are not pooled with the six blocks.

The dense-wide XLSX archive is intentionally retained: it exercises the same
existing full-sheet path used by the earlier XLSX baseline while the tiny rows
provide a short ordinary-workbook control. `xlsx_full_cell_scan` counts only
`Sheet1`; `xlsx_open_owned` still checks the complete workbook sheet count.

## Build and run commands

Build the two source states sequentially from the same checkout. The before
build is source-bound to the frozen baseline manifest, then the one-sample
qualification runs; after the candidate is applied and verified, build the
after binary and run the alternating native blocks. The build driver records
the source manifest, harness tool inventory, frozen packet inputs, lockfile,
compiler environment, and executable identity for each leg. Both builds may
reuse the serialized Cargo target cache at
`/home/zhuhe/code/litchi-target-0806/cross`; `CARGO_INCREMENTAL=0` is fixed and
each release executable is copied to its own hashed before/after path before
capture. This gives each report a source-bound executable without requiring a
second checkout or concurrent target directories. The exact root-only driver
commands are:

```sh
python3 -B docs/performance/results/change-0806/cross_build.py before
python3 -B docs/performance/results/change-0806/cross_capture.py qualification
# Root applies and verifies the exact 0805 candidate here, then builds after.
python3 -B docs/performance/results/change-0806/cross_build.py after
python3 -B docs/performance/results/change-0806/cross_capture.py native
```

The build driver uses the exact command
`cargo build --offline --locked --release --manifest-path
tools/perf-baseline/Cargo.toml --bin litchi-perf-baseline`, with two build jobs
and no feature override. The qualification command is intentionally
before-only; the native command consumes the copied before and after binaries
in the frozen six-block order.

No `--features` flag is needed. The harness manifest already depends on
`litchi-docx`, `litchi-xlsx`, `litchi-ooxml-common`, and `litchi-opc`, and its
umbrella `litchi` dependency enables the document/spreadsheet OOXML features.
Do not use `litchi-perf-baseline-alloc` for this timing control: its elapsed
values are allocator-instrumented and are not latency evidence. No profiler,
heaptrack, Callgrind, or custom probe is required here.

The exact root-run commands below use the copied before/after binaries and a
new output path for every process. Every output path must be absent before its
command starts:

```sh
COMMON=(
  --case docx_semantic_open,docx_semantic_full_text,xlsx_open_owned,xlsx_full_cell_scan
  --semantic-shape tiny,large
  --xlsx-shape tiny,dense-wide
  --warmup 3
  --samples 30
)

taskset -c 12 /home/zhuhe/code/litchi-target-0806/cross-before "${COMMON[@]}" \
  --json docs/performance/results/change-0806/cross-native/0-before.json
taskset -c 12 /home/zhuhe/code/litchi-target-0806/cross-after "${COMMON[@]}" \
  --json docs/performance/results/change-0806/cross-native/0-after.json

taskset -c 12 /home/zhuhe/code/litchi-target-0806/cross-after "${COMMON[@]}" \
  --json docs/performance/results/change-0806/cross-native/1-after.json
taskset -c 12 /home/zhuhe/code/litchi-target-0806/cross-before "${COMMON[@]}" \
  --json docs/performance/results/change-0806/cross-native/1-before.json

taskset -c 12 /home/zhuhe/code/litchi-target-0806/cross-before "${COMMON[@]}" \
  --json docs/performance/results/change-0806/cross-native/2-before.json
taskset -c 12 /home/zhuhe/code/litchi-target-0806/cross-after "${COMMON[@]}" \
  --json docs/performance/results/change-0806/cross-native/2-after.json

taskset -c 12 /home/zhuhe/code/litchi-target-0806/cross-after "${COMMON[@]}" \
  --json docs/performance/results/change-0806/cross-native/3-after.json
taskset -c 12 /home/zhuhe/code/litchi-target-0806/cross-before "${COMMON[@]}" \
  --json docs/performance/results/change-0806/cross-native/3-before.json

taskset -c 12 /home/zhuhe/code/litchi-target-0806/cross-after "${COMMON[@]}" \
  --json docs/performance/results/change-0806/cross-native/4-after.json
taskset -c 12 /home/zhuhe/code/litchi-target-0806/cross-before "${COMMON[@]}" \
  --json docs/performance/results/change-0806/cross-native/4-before.json

taskset -c 12 /home/zhuhe/code/litchi-target-0806/cross-before "${COMMON[@]}" \
  --json docs/performance/results/change-0806/cross-native/5-before.json
taskset -c 12 /home/zhuhe/code/litchi-target-0806/cross-after "${COMMON[@]}" \
  --json docs/performance/results/change-0806/cross-native/5-after.json
```

The capture driver uses the same selector command for the one-sample
qualification, with `--warmup 0 --samples 1` and a separate output tree. The
qualification must compare all eight result keys and corpus identities before
any timing summary is accepted. After all captures, the two copied cross
binaries are removed under a separate `cross-cleanup.json` witness containing
the exact two recorded artifact identities; the main six-binary target cleanup
has its own witness and is not substituted for this lane's cleanup.

## Report and semantic-oracle contract

Each harness report is `schema_version: 1`. The fields used for replay are:

* `configuration.samples_per_case`, `warmup_iterations_per_case`, `cases`,
  `semantic_shapes`, and `xlsx_shapes` must equal the frozen selection.
* Every `results[]` row is identified by `case` plus its complete `corpus`
  manifest. Require equal `name`, `generator`, `package_format`, `shape`,
  payload/compression fields, counts, and `archive_sha256` between A and B
  for each row. The expected generators are
  `litchi-docx-semantic-v1` and `litchi-xlsx-synthetic-v1`.
* `elapsed_ns.samples` contains the 30 in-process measured values; `p50`,
  `p95`, and `p99` are derived by the harness. Normal reports identify
  `tool.instrumentation` as `none`. These rows make a scoped timed-operation
  comparison; they provide no allocator, RSS, physical-I/O, or profiler claim.
* These four selectors do not emit `output_sha256`. Correctness is a fail-closed
  runtime oracle: a nonzero process exit invalidates the report. Do not replace
  that gate with an inferred digest from timing JSON.

The relevant existing oracle and timer paths are:

* `build_semantic_docx_corpus` calls `verify_semantic_docx` before any sample
  (`tools/perf-baseline/src/lib.rs:16553-16557`). Each open sample times only
  `Package::from_reader`, then verifies the complete semantic document
  (`:34187-34194`). Each full-text sample prepares the package/document before
  the clock, times `Document::text()`, and then runs the same full verifier
  (`:34252-34261`). Thus the full-text row is a useful semantic traversal
  control but does not price DOCX package open/setup.
* `build_xlsx_corpus` reopens its generated archive and calls
  `verify_xlsx_cells` before samples (`:22268-22280`). `xlsx_open_owned` times
  `Workbook::from_bytes` and checks the complete sheet count
  (`:42545-42576`). `xlsx_full_cell_scan` opens the workbook outside the clock,
  times `Sheet1.cells("A1:XFD1048576").count()`, and checks the expected stored
  cell count (`:42764-42799`).

The line references are source anchors for replay review; the capture receipt
must bind the exact source hash and executable SHA rather than relying on line
numbers alone.

## Pairing and interpretation

For each of the eight `(case, corpus)` keys, pair the A and B report `p50`
values within the same block and compute `B/A`. Bootstrap the six paired ratios
with replacement using `random.Random(806081)`, 10,000 draws, the median as the
statistic, and sorted zero-based endpoints 250 and 9,749. Treat a p50
regression as significant only when the bootstrap lower bound is above 1.0 and
the point estimate exceeds 1.05; retain every row and every pair range in the
report. This cross-format lane is a regression veto and supplies no additional
benefit requirement or adoption claim.

Before summarizing, require zero process failures, identical corpus hashes,
identical semantic case/shape configuration, and all runtime semantic gates.
An unchanged `docx_semantic_full_text` result cannot be used to infer that
DOCX open was unaffected because its timer starts after package/document setup;
the `docx_semantic_open` row carries that parser-path evidence. Likewise,
`xlsx_full_cell_scan` is a selected worksheet traversal control, while
`xlsx_open_owned` covers workbook parse/open.

The dedicated harness binary is sufficient for this control. The existing
Python resource-profile helper has a fixed large DOCX pair and a different
four-case tiny/medium XLSX borrowed-parser tuple, so using it would either
drop the requested tiny/dense-wide rows or add unrelated selectors. Reusing
the ordinary harness directly keeps the eight-row corpus and timer semantics
exactly as implemented and avoids a new driver or benchmark definition.

After the capture binaries have been removed and `cross-cleanup.json` has been
recorded, replay the retained evidence with the read-only analysis driver. The
cleanup witness uses schema `litchi.performance.0806.cross-cleanup.v1`, exact
copied paths `/home/zhuhe/code/litchi-target-0806/cross-before` and
`/home/zhuhe/code/litchi-target-0806/cross-after`, and exact byte/hash
identities. It is separate from the main probe target cleanup.

```sh
python3 -B docs/performance/results/change-0806/cross_analysis.py --write
python3 -B docs/performance/results/change-0806/cross_analysis.py --check
```

`--check` recomputes the same bootstrap and source/report custody checks and
compares the existing `cross-analysis.json` without writing it. It requires the
before qualification source to match `build-before/source.json` and the native
source to match `build-after/source.json`; qualification and native source
manifests therefore are expected to differ across the candidate transition.

## Frozen 0806 decision boundary

This lane has no benefit requirement and cannot authorize production adoption.
It rejects only when a row's median paired process-p50 ratio is greater than
1.05 and its 95% bootstrap lower endpoint is greater than 1.0. Every row,
including improvements and non-veto regressions, remains in the retained
analysis. The six blocks use `random.Random(806081)` with 10,000 resamples and
zero-based endpoints 250 and 9749. Qualification has one before-only sample
per row and is never pooled with the 2,880 native samples.
