# 0399 — D7 unified chart point-size dispatch

## Status and scope

This is a bounded D7 dispatch slice on `uo2/iwa-d7-chart-font-size`, cut from
`e0e297a82b0a534dfd5611ced3a9811851edbf76`.  It exposes only the migrated
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
* `crates/litchi-ops/tests/iwork_chart_font_size.rs`.

Leaf/shared chart-font crates are read-only dependencies for this slice.  No
ledger row is flipped by this record.  The coordinator owns `PROGRESS.md`,
integration and the final gate.

## Proof and evidence boundary

The source tests cover populated common fixtures, direct unified-vs-leaf reads,
set/reset, exact source preservation, invalid zero values and reopened
readback.  Ignored wrappers are registered for nine full-file routes:
Pages/Keynote/Numbers × read/set/reset.  Qualification remains pending the
coordinator's release build and pinned-CPU ABBA run for every route, with
setup outside the selected symbols, exact single-test success, positive
inclusive attribution, exact call counts and the prescribed <=1.05 route
ratio.  No build, test or capture was run on this held source.

Native mutation evidence is absent.  The nine D7 queue rows appended for this
slice remain `pending`; no E3/E4 claim is made.  Existing W2-A/W2-B/W2-C leaf
receipts and engine-only probes do not qualify these unified routes.

No raw captures or binaries were added.
