# 0399 — D7 unified chart point-size dispatch

## Status and scope

This is a bounded D7 dispatch slice on `uo2/iwa-d7-chart-font-size`, cut from
`e0e297a82b0a534dfd5611ced3a9811851edbf76`.  The implementation and generated
capability snapshots are in `4e88658a06`; the worktree was then synchronized
with coordinator trunk `bc9a37fd75977b01ab3f24cabfdc4b2fe977c4bc` at
`c1fb64d102f3794e815a632e208bf6ab33a6bf24` (the merge contains no D7 runtime
changes).  The subsequent refusal/query isolation repair is
`6790af8a96f3e0b5e9fcb08141f78bf1a11897a2`, with the final narrow query seam
in `c0c9b0d4e4d7fd56398cc8471627c2b6bdc1e15`.  The final write-boundary
representability check and exact-source extreme-size control are in
`cda454ba33103d8583e1e2cd342dbe8cec13829d`; the measured wrappers retain
every ABBA input and output with `black_box` in `cf9f3c1a8777b1798a728b772505415475634fed`.
This exposes only the migrated
chart point-size family through `litchi-ops`:

* Pages: `/body/chart[N]` → `font.size`;
* Keynote: `/slide[N]/chart[N]` → `font.size`;
* Numbers: `/sheet[N]/chart[N]` → `font.size`.

The property is a readable, writable, clearable `Length(FONT_SIZE)`.  A
positive value maps to the leaf's selector-first `edit_*_chart_font_size` set
transaction.  `null` maps only to that leaf family's reset transaction; it
does not mean an empty string or a new default object.  Selectors remain
semantic and kind-local.  Native chart object identifiers, UUIDs, component
names and protobuf values do not cross the unified boundary.

Chart identity/face traits (`font.name`, `font.bold`, `font.italic`), chart
titles, legends, axis style and every table family are explicitly outside this
slice.  The existing 18 leaf chart-font rows remain their own leaf evidence;
this record does not claim that leaf migration alone is unified exposure.

The implementation keeps chart nodes structural during tree construction.
`font.size` is resolved lazily by the node view only when that chart's format
is requested, and successful values are cached on the item.  A focused leaf
error remains a typed `Result` failure; it is not converted to an absent
property.  The iWork `Node::value` accessor is fallible in this slice and uses
the same deferred path as `Node::view().format()`, so the two supported value
views cannot disagree or hide a focused refusal.  The unified writer stages chart-size operations through each
format's existing exact-source leaf patch and retains terminal errors.

The accepted optimization adds a fixed per-container chart cache for the
three canonical chart paths.  Pages, Keynote and Numbers keep their ordinary
tree and query/write routes unchanged; the fast route constructs only the
selected body's, slide's or sheet's chart projection needed to resolve
`/body/chart[N]`, `/slide[N]/chart[N]` or `/sheet[N]/chart[N]`.  A recognized
canonical path that is out of range returns `path_not_found` directly.
Unsupported path syntax still uses the generic navigator.  Inventory/build
failures remain typed leaf errors, including the Pages arrangement and
Numbers drawable inventory seams, rather than being turned into an empty
projection.  Stable synthetic `At` namespaces keep earlier chart nodes valid
after another container is selected.  Active tests compare the fast
projection with the generic `chart[last()]` projection, target a duplicated
second Keynote slide, preserve prior nodes, and retain the Keynote refused-font
control.

## Owned paths

Runtime/schema/dispatch ownership is limited to:

* `crates/litchi-ops/src/format/iwork/{pages,keynote,numbers}.rs`;
* `crates/litchi-ops/src/format/iwork/{keynote/write,write,node,tree,error}.rs`;
* `crates/litchi-ops/src/format/iwork/{nav}.rs` and
  `crates/litchi-ops/src/query/{inspect,walk}.rs` (a default-off private
  typed predicate seam; only iWork opts in);
* `crates/litchi-ops/src/path/vocabulary.rs` (Numbers chart path spelling);
* `crates/litchi-ops/tests/iwork_chart_font_size.rs`;
* `crates/litchi-ops/tests/golden/capabilities/{pages,keynote,numbers}.json`.
* `tools/unified_bench_matrix.py` and the reduced current-source receipt at
  `docs/unified-ops/results/0399/integrated-callgrind.json`.

Leaf/shared chart-font crates are read-only dependencies for this slice.  No
ledger row is flipped by this record.  The coordinator owns `PROGRESS.md`,
integration and the final gate.

## Proof and evidence boundary

The source tests cover populated common fixtures, direct unified-vs-leaf reads,
set/reset, exact source preservation, invalid zero values and reopened
readback.  They also cover the original mixed Pages fixture: the focused leaf
and unified full projection both retain a typed `InvalidSource` refusal.  The
positive mutation fixture is `chart-grid-authored-native.*` because it has the
uniform chart font graph required by this family; the original native fixture
is retained as the refusal control, not silently discarded.  Keynote controls
prove frame/name/z-order reads remain structural when its font graph refuses,
and single, compound and frame-only query predicates use the same typed
property path.

The final all-feature run passed 10 active tests and intentionally left the 9
full-file ABBA wrappers ignored.  The generated capability reports were
regenerated and the all-feature capability suite passed (11 tests).

The optimized source passed `cargo fmt --all -- --check`, strict all-feature
Clippy for `litchi-ops`, the focused chart-font-size tests, rustdoc with
warnings denied, the ledger check, the crate-boundary check, no-`3rdparty`
references and the affected fuzz-lock inventory tests.  The affected glossary
profiles (`ops-iwork` and `ops-all`) are run against the final committed tip;
their receipts are retained outside the checkout under
`(deleted-coord-artifact)`.

The release binaries built from `9078ab5837d34d8c9baa18d54fce0903f5721017`
and `a52c96dd7346a4d88793cde5c07d66150a8e3539` are archived externally, but
are explicitly non-qualifying because subsequent query-scope repairs landed;
they must not be used for capture.  The earlier `c762b00374` matrix failed the
read-route bar (Pages 1.08733 and Keynote 1.09562), and the `91b5c65f69`
matrix is superseded historical evidence from before the per-container
correctness repair.  Neither matrix is current qualification.  That historical
`91b5c65f69` run passed all nine full-file unified/direct routes (18 legs):
each leg exited 0 with exactly one selected test pass, exactly two selected
wrapper calls, and raw full-function Ir equal to summary/totals and the
threshold-100 annotation.  The per-route unified/direct ratios were Pages
read/set/reset `1.032334883/1.015490353/1.021376734`, Keynote
read/set/reset `1.044168189/1.006126258/1.012796304`, and Numbers
read/set/reset `1.008898928/1.000106772/1.006676784`; all are at most 1.05.
The immutable binary is external at
`(deleted-coord-artifact)`
(357475520 bytes, SHA-256
`ae6d9ffb102a73cf4d06f25c6d6a2c1ae9c029a30b5873fcdbdd7f3cc88d0955`).  The
normalized manifest, raw captures, validator receipt and performance review
are external under
`(deleted-coord-artifact)`; the review SHA-256
is `343941dd6e2f80b28b840a5d0ac8f4b1a1c77a85e7a84f877f609d18fa4f71a8`.
This is Callgrind instruction evidence only: it makes no wall-time or native
acceptance claim.  The final integrated coordinator candidate must still be
requalified after later merges; the earlier c762 matrix remains profiling
history only.

The accepted current-source qualification is
`be2b784791d4070a000ae8f6b7c4ab863ea3783a`.  Its immutable release/debug-info
test binary is external at
`(deleted-coord-artifact)`
(359,772,600 bytes, mode `0555`, SHA-256
`2beaa8f51fc22b53a9fb4fe2ebd3f95efc24fd69a0cc752f83902b7e6a422f8b`).  The
fresh CPU5 matrix has 18 valid legs and a maximum unified/direct inclusive-Ir
ratio of 1.0485150956: Pages read/set/reset `1.031370/1.013274/1.020354`,
Keynote read/set/reset `1.048515/1.007203/1.011550`, and Numbers
read/set/reset `1.011123/1.001909/1.006218`.  The raw selected-call validator
(`c0637e4304f5fdbfae6fca6b0e9a9c17f7d10cb9a56c7df6846951fa4bdf6b8e`) confirmed
two calls per wrapper, raw Ir = summary = PROGRAM TOTALS = the threshold-100
aggregate annotation for every leg.  The current report, derived receipt and
raw evidence remain external under
`(deleted-coord-artifact)`; no wall-time or native
acceptance claim is made.  The final integrated coordinator candidate still
requires requalification after later merges.

Native mutation evidence is absent.  The nine D7 queue rows appended for this
slice remain `pending`; no E3/E4 claim is made.  Existing W2-A/W2-B/W2-C leaf
receipts and engine-only probes do not qualify these unified routes.

No raw captures or binaries were added.

## Per-container repair checkpoint

The follow-up cache repair keeps one lazily initialized chart projection per
body, slide or sheet.  The cache entry vector uses fallible reservation before
initialization; a reservation failure is an internal allocation refusal rather
than a partially initialized cache.  Keynote still calls the existing
`Package::slides()` semantic catalog to establish the bounded entry count (the
generic route already depends on that catalog); the optimization limits the
additional chart inventory/tree work to the selected slide.  Numbers applies
the same rule to the selected sheet.

The second-container regression uses the tracked synthetic host-package
fixture `test-data/iwork/numbers/chart-font-size-two-sheet.numbers` and its
receipt.  It contains a real standalone chart on each sheet and compares the
canonical and generic routes for path, complete format values and `font.size`,
then rechecks the same retained first-sheet node after selecting the second
sheet.  The tracked generator recipe first seeds the charts through the public
`litchi-iwa` editor and then uses direct `litchi-numbers` leaf transactions to
make sheet 1 13pt and sheet 2 27pt.  Its manifest, source and exact command
are recorded in the existing host crate and fixture receipt.  The recipe is
available only through its recorded pre-W3 host commit; W3 intentionally
removes it with `litchi-iwa`, while the fixture remains tracked.  The generated
package is a focused identity/path control only and carries no native
application acceptance claim.

The post-sync focused suite is 10 passing tests and 9 intentionally ignored
CPU-slot probes.  The current source-bound release/Callgrind qualification was
accepted on the be2 current source; the earlier `91b5c65f69` matrix remains
superseded profiling history only.
