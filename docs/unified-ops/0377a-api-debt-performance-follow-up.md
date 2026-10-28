# 0377a — bounded API-debt performance follow-up

- Status: bounded source-bound qualification is recorded on `uo2/fin-perf`;
  the branch is synchronized with coordinator trunk at merge commit
  `b9f5c8bd15f3b670793e3ed2f9d987e0725b8f52`. Final integrated qualification
  remains the coordinator's responsibility.
- Date: 2026-10-28.
- Baseline: `839212fa9da2e4eff14dcae9c3aa242e49f25729`.
- Candidate source for the F1/F3/F4/F5 and CLI measurements:
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

## Verification and gaps

Scoped correctness checks passed before this record was written:

```text
cargo test -p litchi-ooxml-common --lib properties::transaction_tests -- --nocapture  # 40 passed
cargo test -p litchi-ops --all-features --lib format::odf_tree:: -- --nocapture  # passed
cargo test -p litchi-ops --no-default-features --features odc --all-targets  # passed
cargo test -p litchi-docx --lib metadata::tests -- --nocapture  # 13 passed
cargo test -p litchi-xlsx --lib metadata::tests -- --nocapture  # 11 passed
```

The candidate release binaries were rebuilt after the harness change; the
candidate target was 5.0 GiB (below the 30 GiB work-package limit), and no raw
captures or binaries are committed. Before handoff, this branch completed the
scoped format check, release harness check, transaction/ODF/DOCX/XLSX tests,
ODC all-target checks, Clippy with the documented pre-existing ODF
`manual_range_contains` allowance, rustdoc (including `litchi-sheet`), the
73-profile API glossary, ledger/matrix checks, and crate-boundary audit. The
source-bound callgrind and compact manifest identify their exact source and
binary hashes; F1/F3/F4/F5 wall values remain provisional under shared-host
load. The coordinator must rerun quiet wall checks and full integrated gates
on the final trunk candidate; only the coordinator may integrate it into
trunk. This branch is clean after merge commit
`b9f5c8bd15f3b670793e3ed2f9d987e0725b8f52`.
