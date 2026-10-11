# 0185 — XLSB sheet ordering

## Contract

The primary native `Edit::reorder_sheets` accepts a full permutation of semantic
sheet selectors resolved against the immutable source catalog. The catalog
includes worksheets, chart tabs, and other tab kinds. Repeated calls replace the
staged permutation; returning to source order produces an exact no-op. Changed
orders use durable operation tag 8, leaving historical operation tags unchanged.

Native lowering moves whole BundleSh records, retaining stable tab identifiers,
relationships and physical sheet payloads. It remaps workbook-view first/active
positions, sheet-scoped defined names, internal ExternSheet endpoints without
changing XTI entry order, and modeled volatile dependencies. Extended-property
catalog titles follow the new order; trailing named-range titles and unchanged
catalog title slots remain byte-exact. External workbook ranges and same-sheet /
invalid-reference sentinels retain their original encoding.

An internal 3-D range that would become noncontiguous receives typed refusal.
Unknown records interspersed with the catalog cannot be detached silently.
Other workbook records outside the proven dependency vocabulary are guarded,
including embedded data-model and slicer/timeline integration records.
Shared-workbook revision headers describe historical display order and are
explicitly guarded, rather than rewritten as if they described current order.
Other opaque dependency guards remain listed in capabilities. All replacements
are staged on a disposable candidate and published through the existing native
package commit; a failed edit leaves the caller's source snapshot unchanged.

The facade implements `move_to` and `swap` through one native permutation over
the complete catalog. It supports start/end, checked positions, relative anchors
and the workbook parent placement. Ordering flushes pending transaction families,
so mixed cell/property/order batches preserve sequential path semantics and
atomic rollback. Source freshness, edit admission, semantic quotas and retained
native policy remain in force.

## Scope

This adds tab ordering. Worksheet insertion/removal, broader structural CRUD,
homogeneous order-command coalescing, tier 2/3 writes, Python/CLI and full-goal
acceptance remain open. No batch-scaling or full-goal completion is claimed.

## Verification and review

The combined shared-OOXML/native-XLSB suites pass 1,645 tests/doctests. The
combined DOCX/XLSX/PPTX/XLSB/encryption facade suite passes 1,243. Final focused
facade ordering and capability checks pass 14 tests. The ordering fixtures cover
mixed chart/worksheet catalogs, stable cell identity, multiple staged orders,
exact cancellation, durable replay/inverse, merge barriers, truncated and
oversized wire counts, and duplicate durable order operations. Dependency tests
cover active/first workbook views, scoped/global names, real external links,
noncontiguous 3-D refusal, volatile references, opaque records and orphan parts.
A final shared-helper regression verifies absent optional title metadata requires
no rewrite.

Review tightened the implementation before publication: historical revision
headers are guarded, unresolved supporting-link indexes are rejected, workbook
records have an explicit admitted vocabulary, and volatile edits require the
unique workbook-owned graph plus a fully modeled source stream. Slicer/timeline
and revision-log content-type guards also cover orphan parts. Exact no-ops skip
these changed-order dependency rewrites. Workbook record staging borrows source
spans rather than retaining a separate byte vector for every record.

The boundary audit passes for 66 packages and 268 internal declarations with
11 existing debt items. Both XLSB glossary profiles pass. The generated
OfficeCLI delta and capability golden expose the remaining ordering restrictions.
Repository lockfiles are unchanged; workspace checks use an external copied
lockfile with cached cc 1.6.0. All five delegated worktrees were removed after
integration. No base-only commits were pending at turn start.

A preexisting conservative state-patch history charge (17 bytes versus its
10-byte wire operation) remains unchanged; this order operation uses its exact
9 + 8N wire size. This record does not claim Python, CLI or parser-fuzz acceptance.


Warnings-denied full/minimal all-target Clippy, affected-crate rustdoc, scoped
formatting and generated parity checks pass. Temporary build and measurement
artifacts are removed after retaining the bounded receipts below.

## Paired measurements

A new standalone harness compares native complete-catalog permutations with
unified `move_to("/sheet[1]", Placement::End)`. Each worksheet has one distinct
cell marker. Inputs are prepared outside timing; source promotion is timed on
both source legs. The matrix uses 2, 10, 100 and 1,000 tabs, eager/source owners,
nine ABBA rounds, CPU 0, and no concurrent builds, tests or heavy audits. Both
runs preserve tab order, cell values and active-tab identity, produce byte-exact
native/facade ZIPs in every case, and have empty failure matrices:
[after](results/0185-xlsb-sheet-order-after.json) and
[repeat](results/0185-xlsb-sheet-order-repeat.json).

| Owner / tabs | After native / facade µs | Repeat native / facade µs |
|---|---:|---:|
| eager / 2 | 393.622 / 411.802 | 399.172 / 418.182 |
| source / 2 | 569.532 / 579.033 | 574.552 / 591.073 |
| eager / 10 | 958.394 / 992.025 | 961.984 / 995.465 |
| source / 10 | 1375.446 / 1399.856 | 1385.126 / 1404.817 |
| eager / 100 | 6541.731 / 6748.671 | 6501.399 / 6678.871 |
| source / 100 | 9205.442 / 9349.813 | 9127.602 / 9259.692 |
| eager / 1000 | 64093.855 / 65990.124 | 64482.377 / 67215.958 |
| source / 1000 | 89298.811 / 90440.597 | 88698.057 / 89949.273 |

Measured facade overhead is 1.3–4.8%, within the scoped 10% single-operation
bar. The two runs use the same binary and source-file digest, recorded alongside
all individual samples and corpus hashes. There was no equivalent native
ordering API before this change, so these are paired after/repeat measurements,
not a before/after speedup claim. This does not establish command-batch scaling
or full acceptance.
