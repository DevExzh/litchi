# Native OLE2 semantic baseline

Date: 2026-08-11

Production base: `a57506d2339ed9384629bbe8accc958279cba0b3`

This is a baseline and coverage tranche, not an optimization claim. It adds 18
opt-in native DOC/XLS/PPT semantic cases over the same deterministic `tiny` and
`large` artifacts used by the public writer benchmarks. No production crate,
dependency edge, public API, or format behavior changed.

## Coverage and boundaries

Each native format now measures public open, list, one-object lookup, complete
semantic extraction, exact no-op edit/publication, and one edit/publication:

- DOC: paragraphs and full text through `litchi_doc::Package`; edits through
  `body_text::Snapshot`.
- XLS: workbook tabs and cells through `litchi_xls::Workbook`; edits through
  `cell_values::Snapshot`.
- PPT: slides and shape text through `litchi_ppt::Package`; edits through the
  root `slide_order::Snapshot` transaction.

Read cases open before timing except the explicit open case. Edit cases start
from an already-open exact-source snapshot, then time transaction creation,
publication, and owned output materialization. After timing, every iteration
checks the complete deterministic semantics. Edit cases additionally check
exact no-op bytes or deterministic changed bytes, exact-source forward patch
application, inverse restoration, and a full public reopen. The ordinary DOC
one-paragraph and PPT one-shape paths necessarily materialize their public
paragraph/slide collections; the case names do not imply a hidden indexed API.

`payload-heavy` remains a writer-throughput shape. It is excluded here because
its long-text/string corpus does not share the numeric XLS edit contract and is
not a comparable cross-format semantic shape.

The default matrix remains 36 cases / 198 records. The harness now exposes 106
selectable cases. CI runs an 18-record tiny smoke and a scheduled/manual
36-record release matrix.

## Reproducible measurement

Command, abbreviated only by naming the complete 18-case list documented in
the tool README:

```text
taskset -c 2 litchi-perf-baseline \
  --warmup 3 --samples 15 --writer-shape tiny,large \
  --case <18 native DOC/XLS/PPT semantic cases> \
  --json docs/performance/results/ole2-semantic-baseline-a57506d23-2026-08-11.json
```

Environment: release profile, Rust 1.95.0, Linux 6.8.0-101-generic, x86_64 AMD
EPYC 9575F VM, Rust system allocator, one visible/pinned CPU. The raw 36-record
report SHA-256 is
`ef906818aa3da3514d1a3f0ea2f32a89487949f950123edb4cd6cdd90ef13e87`.
The worktree-dirty marker reflects this harness/documentation tranche; the
embedded production revision is the exact base above.

## Results

Large-corpus warm-memory release latency, microseconds:

| Case | p50 | p95 / p99 | Mean |
|---|---:|---:|---:|
| DOC open | 756.752 | 882.065 | 768.536 |
| DOC list paragraphs | 462.086 | 618.131 | 485.604 |
| DOC one paragraph | 500.719 | 722.724 | 523.466 |
| DOC full text | 1.041 | 1.782 | 1.059 |
| DOC no-op edit/save | 226.893 | 238.198 | 229.408 |
| DOC one edit/save | 1,416.060 | 1,607.954 | 1,425.133 |
| XLS open | 1,382.875 | 1,488.650 | 1,395.231 |
| XLS list worksheets | 0.220 | 0.411 | 0.228 |
| XLS one cell | 0.331 | 0.510 | 0.363 |
| XLS full cell scan | 89.684 | 100.608 | 89.275 |
| XLS no-op edit/save | 5.027 | 6.419 | 4.753 |
| XLS one edit/save | **1,722.062** | 1,969.482 | 1,732.117 |
| PPT open | 18.345 | 30.151 | 19.233 |
| PPT list slides | 8.892 | 17.544 | 10.107 |
| PPT one shape text | 28.239 | 68.375 | 32.154 |
| PPT full text | 32.636 | 47.335 | 35.763 |
| PPT no-op edit/save | 64.358 | 100.188 | 68.998 |
| PPT one edit/save | 357.002 | 422.482 | 363.354 |

The main tiny edit/publication p50s are DOC 30.752/83.475 us, XLS
0.131/31.092 us, and PPT 25.866/152.880 us for no-op/one-edit respectively.
All individual distributions and corpus hashes are in the raw report.

## Allocation, RSS, and counters

The six large edit cases, one warmup plus ten samples, produced 1,443,883
allocation calls, 631,947 temporary allocations, 8.12 MiB peak heap, and
22.71 MiB profiler RSS under Heaptrack. Those process totals include the
deliberately exhaustive post-timing verification, so they are not per-case
timed allocation counts. The source attribution remains useful:

- DOC process allocation calls are dominated by repeated `parse_sprms` and
  complete revision/document readback.
- XLS commit attribution includes 16,320 `BTreeMap::insert` calls beneath
  complete workbook reopen, with the cell-value parse/store growth site
  retaining about 590 KiB at peak in the combined profile.
- PPT commit attribution reaches document-structure encode/synchronize and
  complete validation; its measured latency is materially below DOC/XLS.

Uninstrumented family runs reported 30,848 KiB maximum RSS for DOC and 30,976
KiB for both XLS and PPT; the complete 36-record run was also 30,976 KiB. These
one-shot process peaks include binary/runtime baseline and are recorded as
guardrails, not precise retained-object sizes.

Hardware counters were available (`perf_event_paranoid=1`). Fifty large
one-edit samples per format plus three warmups consumed 5,926,951,516 cycles,
18,849,815,671 instructions (3.18 IPC), 4,753,521,189 branches, 17,885,872
branch misses (0.38%), 37,723 page faults, 30 context switches, and one CPU
migration. This aggregate identifies a CPU/allocation workload, not a
per-format comparison.

## Ranked next work

Native XLS one-cell publication is the next measured candidate: it has the
largest large-shape p50 in this matrix, its 1.722 ms commit exceeds the already
opened no-op by roughly 342x, and the profile reaches complete workbook reopen
and per-cell map construction. Any optimization must preserve the current
exact-source patch, one-stream diagnostics, complete BIFF/CFB validation,
semantic readback, and exact inverse behavior. A before/after change needs an
independent open guard because native XLS open is already the second-largest
large result at 1.383 ms.

DOC one-edit publication is second. Its source owner currently performs
complete revision/document validation, and any proposed reuse must retain the
tracked-revision, style, property, and exact-byte contracts. PPT is not the
first target from this matrix.

## Verification

- `cargo fmt --manifest-path tools/perf-baseline/Cargo.toml -- --check`
- `cargo test --locked --manifest-path tools/perf-baseline/Cargo.toml` — 23
  passed
- `cargo clippy --locked --manifest-path tools/perf-baseline/Cargo.toml --all-targets -- -D warnings`
- debug all-18 tiny semantic smoke — 18 records
- release pinned baseline — 36 records, 15 samples each
- CI JSON assertions — 18-record smoke and 36-record scheduled matrix

Two existing `litchi-odf-common` generic-array deprecation warnings remain
visible in dependency builds. They are outside this native OLE2 tranche and do
not affect the warning-denied harness clippy result.
