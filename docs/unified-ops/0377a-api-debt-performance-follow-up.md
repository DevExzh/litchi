# 0377a — bounded API-debt performance follow-up

- Status: measured on `uo2/fin-perf`; coordinator review and final integrated
  qualification remain pending.
- Date: 2026-10-28.
- Baseline: `839212fa9da2e4eff14dcae9c3aa242e49f25729`.
- Candidate source: `187863518e` plus the recorded `extra.rs` working-tree
  change; the source hashes and release binary hashes are in
  `results/0377a/cli.json`.
- Scope: finish the bounded 0377a checks for F1 DOC first-paragraph parsing,
  F2 OOXML metadata publication, F3 XLSX snapshot/leaf-vs-unified cost, and
  F4/F5 ODC admission and lazy object-tree behavior. ODT/ODP remain explicitly
  owner-deferred and are not widened here.

## Contract and correctness

This follow-up changes no public API. The existing selector-first transaction
contracts remain in force. The two added tests close the evidence gaps called
out by the review:

- `doc_eager_first_paragraph_text_set` compares the leaf and unified output
  byte-for-byte in the perf harness. The unified route still opens and commits
  the paragraph edit through the same immutable-source path.
- `core_only_publication_replays_exactly_and_preserves_unread_families`
  publishes only core OOXML metadata, proves that app/custom XML remains
  byte-identical, reads the changed core and untouched families back, applies
  the exact patch to the original source, and applies `inverse()` to restore
  all three original members. This is the exact-source/readback proof for the
  core-only publication path; it is not a synthetic `Patch::check` shortcut.
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
batch), and the results are per-call Ir. Raw callgrind captures remain outside
the repository at
`/home/zhuhe/code/litchi-worktrees/.coord/fin-perf/cga-0377a-20261028/` and
`f1-20261028/`; the compact committed receipt is
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

| family | route | leaf Ir | unified Ir | leaf ratio | unified ratio | wall unified/leaf |
|---|---|---:|---:|---:|---:|---:|
| F1 | DOC first paragraph set | 0.763 | 0.764 | 0.763 | 0.764 | 1.019 |
| F2 | DOCX metadata title set | 0.938 | 0.951 | 0.938 | 0.951 | — |
| F2 | XLSX metadata title set | 0.995 | 1.023 | 0.995 | 1.023 | — |
| F2 | PPTX metadata title set | 0.948 | 0.945 | 0.948 | 0.945 | — |
| F3 | XLSX single cell set | 1.000 | 1.018 | 1.000 | 1.018 | 1.081 |
| F3 | XLSX 1,000-cell batch | 0.988 | 1.032 | 0.988 | 1.032 | 1.775 |
| F4/F5 | ODC open | 0.991 | 0.924 | 0.991 | 0.924 | 0.999 |

The F3 single-cell wall route is below the 1.10 leaf/unified budget, and the
batch route is below 2.0. ODC is below the 1.05 open budget. The F2 and F3
instruction rows meet the requested <=1.05 snapshot/metadata bounds.

The CLI remeasurement found no unexplained regression in any of the four
fixtures. Candidate/baseline median ratios for `cli-get` were DOCX 0.997,
XLSX 0.966, PPTX 0.976 and XLS 0.959; `cli-version` was 0.925–0.954. The
in-process medians were 1.000–1.031. These process-start rows are diagnostic
and carry no unified leaf budget.

## Verification and gaps

Scoped correctness checks passed before this record was written:

```text
cargo test -p litchi-ooxml-common --lib properties::transaction_tests -- --nocapture  # 35 passed
cargo test -p litchi-ops --all-features --lib format::odf_tree::tests::opened_trees_that_cannot_fail_are_built_on_first_read -- --nocapture  # passed
cargo test -p litchi-docx --lib metadata::tests -- --nocapture  # 13 passed
cargo test -p litchi-xlsx --lib metadata::tests -- --nocapture  # 11 passed
```

The candidate release binaries were rebuilt after the harness change; the
candidate target was 5.0 GiB (below the 30 GiB work-package limit), and no raw
captures or binaries are committed. Remaining gates are `cargo fmt`, scoped
Clippy/checks after the current trunk merge, the matrix `--check`, and a quiet
final wall qualification on the coordinator's final integrated commit. The
branch must be merged with current `feat/unified-ops-api` before handoff; only
the coordinator may integrate it into trunk.
