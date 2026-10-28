# 0377a — bounded API-debt performance follow-up

- Status: bounded source-bound qualification is recorded on `uo2/fin-perf` at
  repair commit `1abe5a5cf794a4ec78de7f51f4c346e77cf4e28b`, a descendant of the
  coordinator frozen candidate `bb2475b8b0c11c8cef304178ee3fca1f590e9633`.
  Its immediate parent is the bounded predecessor repair
  `27a22f7ccaf6ab3889040bb8d10300c802e325fe`. Final integrated qualification
  remains the coordinator's responsibility.
- Date: 2026-10-28.
- Baseline: `839212fa9da2e4eff14dcae9c3aa242e49f25729`.
- Historical candidate source for the F1/F3/F4/F5 and CLI measurements:
  `187863518e` plus the recorded `extra.rs` working-tree change; the source
  hashes and release binary hashes are in `results/0377a/cli.json`.
- Candidate source for the source-bound F2 remeasurement:
  `123409faecd8f6d5451e4d368ff21617e89f8289`; its compact artifact manifest
  is `results/0377a/manifest.json`.
- Scope: finish the bounded 0377a checks for F1 DOC first-paragraph parsing,
  F2 OOXML metadata publication, F3 XLSX snapshot/leaf-vs-unified cost, and
  F4/F5 ODC admission and lazy object-tree behavior. ODT/ODP remain explicitly
  owner-deferred and are not widened here.

## Contract and correctness

This follow-up adds no supported public API. The opaque package identity and
revision accessors are `doc(hidden)` implementation proof, and the duplicate
`new_with_own_snapshot` constructor was removed. The existing selector-first
transaction contracts remain in force. Snapshot fast paths now require the
captured package identity, an unchanged package revision, and sealed built-in
parts. Caller-defined `Part` implementations always use exact
`Patch::apply` authentication, including `Edit::new_with_options`,
`Edit::new_with_snapshot`, and core-only publication. The added tests close
the evidence gaps called out by the review:

- `doc_eager_first_paragraph_text_set` compares the leaf and unified output
  byte-for-byte in the perf harness. The unified route still opens and commits
  the paragraph edit through the same immutable-source path.
- `core_only_publication_replays_exactly_and_preserves_unread_families`
  publishes only core OOXML metadata, proves that app/custom XML remains
  byte-identical, reads the changed core and untouched families back, applies
  the exact patch to the original source, and applies `inverse()` to restore
  all three original members. This is the exact-source/readback proof for the
  core-only publication path; it is not a synthetic `Patch::check` shortcut.
- The snapshot constructor regressions cover a snapshot from another package,
  a same-handle payload mutation, and a content-type mutation. A custom
  interior-mutable `Part` can change its observable payload through shared
  interior state without a package seam; both edit constructors and the
  core-only publisher must refuse the stale snapshot through typed patch
  conflict authentication. The built-in fast path is guarded independently
  of this fixture, so arbitrary custom `Part` projections remain on exact
  `Patch::apply` authentication.
- The ODF tree test retains lazy construction for infallible ODC/ODG/ODI/formula
  trees, but creates a malformed ODB catalog that the leaf admits and the
  unified route must reject during open. This preserves the prior admission and
  error timing instead of deferring every object-tree refusal to first read.

No native-authored receipt or Apple application acceptance is claimed by this
record. The routes have no new native evidence; any later native requirement
remains in the coordinator's queue.

## Measurement method

The baseline and candidate use the corrected `fin_perf_region` harness source
(SHA-256 `551e5440ae82bfe983fb74c8a5e4e36e727176469efda0998e9b3c29d79b1a4d`)
with the warm-up call outside the counted region on both builds. Callgrind was
run in release mode with three counted calls per route (one for the 1,000-cell
batch), and the results are per-call Ir. F1/F3/F4/F5 retain the earlier
corrected-warmup run from the candidate source above. F2 was rebuilt from
source commit `123409faecd8` after the provenance repair and rerun for all
three OOXML metadata routes. The F2 raw captures are outside the repository
at
`/home/zhuhe/code/litchi-worktrees/.coord/fin-perf/cga-f2-final2-20261028/`;
the complete compact artifact manifest is
`results/0377a/manifest.json`. The committed receipt is
`results/0377a/callgrind.tsv`.

Leaf/unified wall checks used pinned CPU 47, ABBA binary order, five blocks,
and ten rounds per block for F3/F5 (five rounds for F1). The run was observed
under load averages of about 12.4–13.7 while other worktrees were compiling;
therefore these wall values are a measured diagnostic and must be rerun in a
quiet coordinator interval for final native qualification. All wall cases
reported exact bytes where the route has an output comparison. CLI was measured
with 30 iterations and 10 ABBA rounds per fixture (600 samples per leg).

## Results

The primary callgrind ratios are candidate/baseline; all are within the 1.05
regression bar. Wall ratios are candidate unified/leaf and use the normal
single or batch budget.

| family | route | leaf candidate Ir | unified candidate Ir | leaf ratio | unified ratio | wall unified/leaf |
|---|---|---:|---:|---:|---:|---:|
| F1 | DOC first paragraph set | 42,086,073 | 42,646,284 | 0.763 | 0.764 | 1.019 |
| F2 | DOCX metadata title set | 66,417 | 70,696 | 0.939 | 0.966 | — |
| F2 | XLSX metadata title set | 135,060 | 142,310 | 0.999 | 1.025 | — |
| F2 | PPTX metadata title set | 1,338,478 | 1,349,360 | 0.947 | 0.951 | — |
| F3 | XLSX single cell set | 332,510 | 352,752 | 1.000 | 1.018 | 1.081 |
| F3 | XLSX 1,000-cell batch | 20,232,379 | 30,951,182 | 0.988 | 1.032 | 1.775 |
| F4/F5 | ODC open | 26,440,028 | 26,419,395 | 0.991 | 0.924 | 0.999 |

The F3 single-cell wall route is below the 1.10 leaf/unified budget, and the
batch route is below 2.0. ODC is below the 1.05 open budget. The F2 and F3
instruction rows meet the requested <=1.05 snapshot/metadata bounds.

The CLI remeasurement found no unexplained regression in any of the four
fixtures. Candidate/baseline median ratios for `cli-get` were DOCX 0.997,
XLSX 0.966, PPTX 0.976 and XLS 0.959; `cli-version` was 0.925–0.954. The
in-process medians were 1.000–1.031. These process-start rows are diagnostic
and carry no unified leaf budget.

## Final XLSX batch repair remeasurement

The first integrated candidate, `bb2475b8b0c11c8cef304178ee3fca1f590e9633`,
failed the source-bound XLSX 1,000-cell unified instruction route. Its
preserved receipt is outside the repository at
`/home/zhuhe/code/litchi-worktrees/.coord/resume-20261028/perf/fin-final-bb2475b8b0/`:
base totals were 30,045,503 and 29,998,170 Ir, candidate totals were
31,582,848 and 31,580,006 Ir, with median ratio 1.051948538 and worst ratio
1.052825822. This failed receipt remains evidence; it was not overwritten.

The first bounded predecessor, `27a22f7ccaf6ab3889040bb8d10300c802e325fe`,
also failed under the matching release-debug configuration. Its preserved
ABBA capture directory is
`/home/zhuhe/code/litchi-worktrees/.coord/fin-perf/repair-27a22-batch1000-debug1-preopt/`:
base totals were 30,007,770 and 30,011,846 Ir, candidate totals were
31,611,640 and 31,690,204 Ir, with paired ratios 1.053448490 and
1.055923184, median 1.054685837 and worst 1.055923184. The earlier diagnostic
debug-info-off run is retained separately at
`/home/zhuhe/code/litchi-worktrees/.coord/fin-perf/repair-27a22-batch1000/`:
30,049,152 / 31,683,801 / 31,598,838 / 30,056,013 Ir, median 1.052865416
and worst 1.054399172. Neither failed receipt was replaced by the repaired
run below.

The bounded repair is intentionally narrower than a general setter change:
the native worksheet setter returns a hidden coalescer result, and the unified
route consumes it only for unsigned runs with distinct bounded targets. The
result is `false` only for a bounded exact source-equivalent no-op and `true`
after staging a payload. Signed/repeated runs retain the projected-state path,
as do rich/retype/clear operations. Rich and shared-string source checks still
distinguish same-visible-text representation changes, so a plain set over rich
text is not collapsed as a no-op.

The repaired source-bound candidate is commit `1abe5a5cf794a4ec78de7f51f4c346e77cf4e28b`.
It was built from the clean worktree with release debug information enabled by
the harness profile (matching the baseline), LTO and panic-abort:

```text
cwd=/home/zhuhe/code/litchi-worktrees/uo2-fin-perf
env PATH=/home/zhuhe/code/litchi-worktrees/.coord/bin:$HOME/.cargo/bin:$PATH CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/uo2-fin-perf/target CARGO_INCREMENTAL=0 CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0 CARGO_BUILD_JOBS=4 RUST_TEST_THREADS=4 cargo build --release --offline --locked --manifest-path tools/perf-baseline/unified-fin-perf/Cargo.toml
```

The candidate copy is
`/home/zhuhe/code/litchi-worktrees/.coord/fin-perf/bin/fin-cand-1abe5a5cf7-debug1`,
385,525,480 bytes, SHA-256
`0c55e2ab3f8142003c16b0431049924c923397774a85a0946d0e38cc64074c22`.
The baseline is `fin-base`, source `839212fa9da2e4eff14dcae9c3aa242e49f25729`,
362,495,936 bytes, SHA-256
`1c8d4b7abdf94f7d3ad0a7b6665cf512a4ebbc9ae9d76786d76779e1f96e883b`.

Using the corrected `fin_perf_region` harness (main source SHA-256
`551e5440ae82bfe983fb74c8a5e4e36e727176469efda0998e9b3c29d79b1a4d`),
Callgrind, `taskset -c 47`, and ABBA order `base, candidate, candidate, base`,
the retained captures are in
`/home/zhuhe/code/litchi-worktrees/.coord/fin-perf/repair-1abe5a5cf7-batch1000-debug1/`.
The route printed a successful JSON result in every log and all four counted
totals were positive:

| capture | build | Ir |
|---|---|---:|
| 01 | baseline | 30,048,818 |
| 02 | candidate | 30,290,668 |
| 03 | candidate | 30,209,692 |
| 04 | baseline | 30,055,506 |

The paired ratios are 1.008048569 and 1.005130042; median is 1.006589306 and
worst is 1.008048569. Both are below the 1.05 instruction-cost bar. This is
instruction-count evidence for the unified batch route; no quiet wall or CLI
qualification is claimed by this repair receipt.

The repaired captures can be regenerated from the source-bound binaries with
the following per-capture command (substitute the listed binary and output
prefix for each ABBA leg, without reusing an output file):

```sh
systemd-run --user --scope --quiet --collect \
  -p OOMPolicy=continue -p MemoryMax=20G -p MemorySwapMax=0 -- \
  taskset --cpu-list 47 valgrind --tool=callgrind --collect-atstart=no \
  --toggle-collect='*fin_perf_region*' \
  --callgrind-out-file=/home/zhuhe/code/litchi-worktrees/.coord/fin-perf/repair-1abe5a5cf7-batch1000-debug1/<capture>.cg \
  /home/zhuhe/code/litchi-worktrees/.coord/fin-perf/bin/<fin-base-or-fin-cand-1abe5a5cf7-debug1> \
  --groups batch --formats xlsx \
  --routes xlsx_eager_cell_value_set_batch1000 --callgrind unified --iterations 1 \
  > /home/zhuhe/code/litchi-worktrees/.coord/fin-perf/repair-1abe5a5cf7-batch1000-debug1/<capture>.log 2>&1
```

For symbol attribution, run the same `callgrind_annotate` command used by the
reviewed capture set on each resulting `.cg` file, writing to a new external
text file (the raw `.cg` and annotation files remain outside the repository):

```sh
callgrind_annotate --auto=yes --inclusive=no --threshold=1 \
  /home/zhuhe/code/litchi-worktrees/.coord/fin-perf/repair-1abe5a5cf7-batch1000-debug1/<capture>.cg \
  > /home/zhuhe/code/litchi-worktrees/.coord/fin-perf/repair-1abe5a5cf7-batch1000-debug1/<capture>.annotate.txt
```

The focused verification covered the unsigned coalescer's scalar/no-op and
formula transaction cases, plus native exact-source cases for same-text rich
inline/shared strings and formula state. `xlsx_coalescing` passed 11 tests and
the native cell edit suite passed 80 tests, including
`plain_set_over_same_visible_rich_text_erases_runs`,
`same_text_rich_run_property_change_is_not_collapsed_as_a_noop`, and
`unbounded_plain_set_preserves_shared_string_identity`.

## Verification and gaps

Scoped correctness checks passed before this record was written:

```text
cargo test -p litchi-ooxml-common --lib properties::transaction_tests -- --nocapture  # 40 passed
cargo test -p litchi-ops --all-features --lib format::odf_tree:: -- --nocapture  # passed
cargo test -p litchi-ops --no-default-features --features odc --all-targets  # passed
cargo test -p litchi-docx --lib metadata::tests -- --nocapture  # 13 passed
cargo test -p litchi-xlsx --lib metadata::tests -- --nocapture  # 11 passed
cargo test -p litchi-ops --all-features --test xlsx_coalescing -- --nocapture  # 11 passed
cargo test -p litchi-xlsx --all-features --lib workbook::edit::tests::cells:: -- --nocapture  # 80 passed
cargo test -p litchi-xlsx --all-features --lib workbook::edit::tests::worksheet_edit_state:: -- --nocapture  # 3 passed
```

The candidate release binaries were rebuilt after the harness change; the
current candidate target is 28 GiB (below the 30 GiB work-package limit), and no raw
captures or binaries are committed. Before handoff, this branch completed the
scoped format check, release harness check, transaction/ODF/DOCX/XLSX tests,
ODC all-target checks, strict scoped Clippy, rustdoc (including `litchi-sheet`), the
73-profile API glossary, ledger/matrix checks, and crate-boundary audit. The
source-bound callgrind and compact manifest identify their exact source and
binary hashes; F1/F3/F4/F5 wall values remain provisional under shared-host
load. The coordinator must rerun quiet wall checks and full integrated gates
on the final trunk candidate; only the coordinator may integrate it into
trunk. This handoff adds the bounded XLSX coalescer repair at
`1abe5a5cf794a4ec78de7f51f4c346e77cf4e28b` and the compact receipt
`results/0377a/xlsx-batch-repair.json`; it does not claim integrated trunk
qualification.
