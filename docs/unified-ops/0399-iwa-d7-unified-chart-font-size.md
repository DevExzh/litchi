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

The final all-feature run passed 5 active tests and intentionally left the 9
full-file ABBA wrappers ignored.  The generated capability reports were
regenerated and the all-feature capability suite passed (11 tests).

The final synchronized source passed `cargo fmt --all -- --check`, strict
all-feature Clippy for `litchi-ops`, the focused chart-font-size and capability
tests, the Keynote node refusal unit, the ledger check, the 70-package
crate-boundary check, and the no-3rdparty reference check.  Feature-specific
Clippy (no-format, Pages, Keynote and Numbers), rustdoc (Pages, Keynote,
Numbers and all-features), and the affected glossary profiles (`ops-iwork` and
`ops-all`) also passed.  Logs are retained outside the checkout under
`.coord/d7-chart-font-size-validation/logs/`.

The release binaries built from `9078ab5837d34d8c9baa18d54fce0903f5721017`
and `a52c96dd7346a4d88793cde5c07d66150a8e3539` are archived externally, but
are explicitly non-qualifying because subsequent query-scope repairs landed;
they must not be used for capture.  Qualification remains pending a fresh
coordinator release build from `cf9f3c1a8777b1798a728b772505415475634fed` and pinned-CPU
ABBA run for every route, with setup outside the selected symbols, exact
single-test success, positive inclusive attribution, exact call counts and the
prescribed <=1.05 route ratio.

Native mutation evidence is absent.  The nine D7 queue rows appended for this
slice remain `pending`; no E3/E4 claim is made.  Existing W2-A/W2-B/W2-C leaf
receipts and engine-only probes do not qualify these unified routes.

No raw captures or binaries were added.
