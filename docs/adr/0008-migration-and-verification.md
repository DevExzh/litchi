# ADR 0008: Migration and verification

- Status: Accepted
- Date: 2026-07-31

## Migration policy

The refactor is breaking but proceeds in dependency-ordered, continuously
buildable phases. There are no released compatibility shims.

1. Foundation: reduce `litchi-core`, add source/budget/selector/scalar/error
   contracts, neutral vocabulary crates, dependency checks, and API tests.
2. OOXML substrate: split `litchi-ooxml-common` and `litchi-drawingml` from the
   monolith while preserving raw unknown content.
3. OOXML vertical slices: migrate XLSX first to stress large-file and concurrent
   CRUD, PPTX next to prove DrawingML/shape semantics, and DOCX to prove
   story/style semantics. The three may overlap where dependencies allow.
4. Migrate XLSB onto `litchi-sheet` and shared DrawingML without depending on
   XLSX.
5. Split `litchi-ole-common` and `litchi-odraw`, then migrate XLS, PPT, and DOC
   onto the same neutral vocabularies while retaining binary-native models.
6. Replace the umbrella facade, delete the old monolith crates, and run the full
   support certification matrix.

Every phase supplies concise compile-tested golden paths and negative
compile-fail cases. Public API snapshots reject redundant prefixes and facade
type noise. No format is advertised as supported until required checklist rows
pass.

The first implementation slice establishes bounded scalars, selectors,
positional sources, hierarchical budgets, neutral word/sheet/slide vocabulary
crates, and executable dependency fences. Markup compatibility, shared document
properties, and external-workbook relationship vocabulary have moved into
`litchi-ooxml-common`; the format-independent DrawingML XML primitives have
moved into `litchi-drawingml`. The monolith remains only as a migration host and
`litchi-core` still carries explicitly fenced extraction debt.

The second implementation slice moves shared OOXML namespace, entity,
attribute, and bounded OMML scanning helpers into `litchi-ooxml-common`. It also
introduces the independent `litchi-xlsx` crate: the canonical workbook-catalog
parser now lives there, while the migration host contains only conversions to
its legacy internal records. `litchi-xlsx::Workbook` is an immutable `Send +
Sync` snapshot with lifetime-free sheet handles, content-derived flavor, a
deterministic one-visible-sheet baseline, and selector-first `Result<Option<_>>`
lookup by name or checked zero-based position. This was the first XLSX vertical
slice; lossless edit/patch commits remain migration work.

The third implementation slice adds the first semantic worksheet read path.
`litchi-sheet` rows, columns, and addresses now have checked constructors and
cannot represent coordinates outside the modern spreadsheet grid. A half-open
`Rect` stores its exclusive boundary separately, so it can cover the final row
and column without admitting a sentinel as a valid cell. Borrowed A1 lookup such
as `sheet.cell("A1")` is the semantic main entry; raw zero-based `(row, column)`
tuples remain convenient, and reusable checked `Address` values avoid repeated
validation in hot loops. The initial lookup returned `Result<Option<&Cell>>`;
the twenty-second slice replaces that provisional shape with
`Result<cell::View<'_>>`, distinguishing missing, explicitly stored, and
merge-covered coordinates without an indexing panic or native identifier.

Worksheet payloads load on first use into hidden thread-safe snapshot caches.
The parser streams into a row-major compact sparse slice, resolves shared
strings through cheaply cloned immutable text, expands shared-formula storage
records, retains exact numeric lexical forms, and keeps formula cache origin and
freshness separate. Sparse `cells(range)` traversal and stored extent are
implemented; the eighth slice later separates declared, content, and direct-
style extents. Fully resolved formatted extents including row/column defaults,
rich-text formatting, dynamic-array spill states, shared-style
definition editing, dense budgeted grids, cache eviction, and operation budgets
remain open, as does replacing remaining parser `Invalid` messages with the
full structured context taxonomy. The current non-evicting cache is therefore a
safe migration step, not the weighted-cache design promised by ADR 0005.

The parser's Office-profile limits follow the checked-in `[MS-OE376]`
conformance notes: row and column grid bounds, non-decreasing row order, cell
style indexes through 65,490, one-based cell/value metadata indexes below 2³¹,
32,767-character cell and shared-string limits, bounded shared-string run/count
hints, required shared-formula indexes, and a false formula `bx` flag. Advisory
shared-string counts size allocations but never override the parsed items.

The fourth implementation slice adds transactional ordinary-cell writes to the
new `litchi-xlsx` crate. `Workbook::edit()` creates isolated state and
`Edit::sheet(selector)` retains the checked name/position lookup contract.
`SheetEdit` exposes the short verbs `set`, `clear`, and `remove`; dropping an
edit rolls it back, while `commit` validates every rewritten worksheet before
publishing a new immutable workbook. The source snapshot is never changed.
Plain strings remain inert text, formulas require an explicit checked
`Formula`, floating-point numbers reject non-finite values, and `t="d"` values
use a checked `Date` type that retains the original ISO 8601 lexical form.
`clear` retains the cell record and local metadata; `remove` deletes that
record without shifting surrounding cells. Clearing a cell created earlier in
the same transaction therefore produces an explicit empty record, while
clearing a missing source cell is a no-op.

The transaction structurally clones the OPC graph while sharing clean part
payloads through immutable `Arc` storage. It rewrites only affected worksheet
rows and cells, preserving untouched bytes plus touched-cell styles, metadata,
unknown attributes, and unrelated children. It deliberately refuses edits
whose semantics are not yet proven: signed packages, protected sheets, cells
covered by data validation or a merged-cell follower, multi-cell formula
groups, markup-compatibility-controlled sheet data, and unknown cell payloads
except explicit removal. Literal SpreadsheetML escape sequences and XML control
characters are encoded without changing their text meaning.

Successful commits return semantic before/after changes and an in-memory
source-checked reversible patch. Private part deltas share their byte owners;
application checks expected source bytes and relationship state before
mutation. It retains exact changed-part bytes and relationship fields, but does
not claim byte-identical ZIP containers. Cell edits remove an existing
calculation-chain relationship and part, set
calculation properties for a full refresh, and retain the removed graph in the
inverse patch. The boolean spellings used for calculation properties follow
the Office compatibility notes in `[MS-OE376]` §2.1.599. This is intentionally
not yet the format-independent deterministic patch wire representation required
by ADR 0003.

The fifth implementation slice makes ordinary-cell edit planning composable
across threads without exposing locks. Multiple `Edit` values may be prepared
independently from cheap clones of the same immutable workbook snapshot and
then joined. The initial `join` accepted only exact snapshot lineage and
cell-disjoint write sets; the seventh slice refines this to disjoint effect
facets on a cell. It checks ordered maps without materializing a combined copy,
then moves the incoming edit's action maps into the accepted edit. It does not
use last-writer-wins behavior.

An overlap returns a deterministic `ConflictSet` grouped by developer-facing
sheet name and checked position, with ordered cell addresses. A lineage mismatch
is distinct from an overlap. In either case the accepted edit remains unchanged
and `JoinError` returns ownership of the rejected edit, so error handling cannot
silently discard prepared work. `JoinError` also converts into the crate's
boxed error variant for concise `?` use. Workbook opening now rejects two
logical sheets that alias the same physical sheet part; otherwise apparently
disjoint logical edits could overwrite one another during commit.

The fifth-slice example artifact was produced by joining two independently
prepared edits. Microsoft Excel for macOS opened the generated package without
a repair or compatibility dialog, reported a used range of `A1:B1`, and exposed
`Revenue` at `A1` and `42` at `B1`. This verifies native open and cell values for
that artifact only; it is not evidence of native resave fidelity, contention,
or performance.

The sixth implementation slice replaces archive-sized buffering in the shared
OPC output path. `PackageWriter` now emits ZIP records directly to any
sequential `Write` sink; seeking is not required, and the memory-returning
`to_bytes` path remains an explicit choice. A caller-owned sink failure after
output begins returns typed `IncompleteOutput` context with the accepted byte
count. `litchi-xlsx::Workbook` exposes the concise `write_to` and `save` entry
points over this substrate.

Ordinary filesystem save now writes a sibling temporary artifact, finalizes and
flushes the ZIP, synchronizes the file, and only then replaces the destination.
On Unix it preserves an existing regular file's permissions. On all supported
platforms it refuses symbolic-link and non-file destinations, cleans up a
failed temporary artifact, and synchronizes the parent directory where the
platform supports it. Tests inject a failure
after a partial temporary write and prove the original bytes remain unchanged.
A separate 2 MiB incompressible-payload test uses a non-seekable sink that
rejects any write over 64 KiB; direct streaming succeeds and an injected sink
failure reports the exact 128 accepted bytes. These are functional output-shape
checks, not allocation, latency, or throughput measurements.

Desktop Excel for macOS opened the joined-edit artifact written through the new
atomic streaming path without a repair or compatibility dialog and reported
`Revenue` at `A1`, `42` at `B1`, and the expected `A1:B1` used range. This adds
native-open evidence for the physical save path only; it does not certify crash
durability on every filesystem or native resave fidelity.

The seventh implementation slice makes SpreadsheetML shared cell formats
first-class without exposing their numeric `s` identifiers. The workbook
validates that there is at most one internal styles relationship with the
required content type, then lazily preprocesses `styles.xml` through the common
markup-compatibility engine and validates the direct `cellXfs` collection.
Counts and cell references are bounded against the checked-in `[MS-OE376]`
profile: §2.1.599 defines `c/@s` as a `cellXfs` index through 65,490, while
§2.1.728 limits the `cellXfs` collection itself to 65,430 records. Missing,
empty, duplicate, count-mismatched, and out-of-range tables/references fail
before the affected style is observed or a style-dependent edit is published.
`Workbook::new()` now includes a minimal valid base style resource.

`Workbook::styles()` returns a compact immutable `Styles` view;
`Styles::base`, checked `get`, and iteration return cheap `Style` handles.
The ordinary semantic entry is `Sheet::style(cell)`, while numeric position is
the secondary import/diagnostic path. `Sheet::local_style` distinguishes a
missing cell, an existing cell with no local style, and an explicit shared
style reference; `Sheet::style` separately resolves the omitted local
reference to the base cell format. Handles carry exact snapshot ownership,
which pins the snapshot used by queries such as fan-out, plus a separate
pointer-identity lineage for the immutable shared-style table. That resource
lineage is retained across descendant cell-only commits, so handles and opaque
patch keys remain safely reusable while the table is unchanged; unrelated
tables are rejected by a typed error and cannot compare equal accidentally. No
public constructor or getter exposes the physical SpreadsheetML index. When a
source-checked patch is replayed onto a separately opened byte-identical
workbook, its reported style states are rebound to that target table's lineage
so every returned key resolves against the resulting snapshot. A patch whose
before/after state contains an explicit shared style also retains a shared byte
guard for the styles resource; replay against a different relationship target
or table is rejected rather than reinterpreting the opaque key.

`Style::fan_out` reports the number of stored worksheet cells whose effective
cell format is that resource, including implicit base-format use. It
deliberately does not guess at unstored grid positions or row/column defaults.
`SheetEdit::style` retargets a cell to an existing shared resource and creates a
styled empty record when necessary; `reset_style` removes only the explicit
local reference. Shared format definitions remain immutable in this slice, so
editing/forking a resource and reporting its prospective selection-wide
fan-out remain separate future operations.

Cell edit plans now represent payload and local-style effects as orthogonal
facets. Independent edits may join on the same cell when one changes content
and the other changes style; two payload effects, two style effects, or a
whole-cell removal still conflict deterministically. Commit moves accepted
actions into the rewrite plan instead of cloning their content. The XML surgery
rewrites only the touched cell tag for style-only changes, preserving its
payload, metadata, extensions, unknown attributes, and untouched bytes. Patch
states are now a data-bearing enum that makes invalid combinations
unrepresentable: a cell state always contains both content and exact local
style state, while a missing state contains neither. Tests cover local versus
resolved semantics, stored-cell fan-out, styled empty cells, reset, inverse
byte restoration, foreign-lineage rejection, malformed style graphs/tables,
same-cell threaded joins, and composed payload/style rewrites.

For native evidence, Excel for macOS applied bold formatting to `A1` and saved
the source workbook itself. The public `copy_style` example opened that native
file, obtained `A1` through `Sheet::style`, set `C1` to `42`, retargeted `C1`
with the borrowed handle, and atomically saved the result. Excel opened the
result without a repair or compatibility dialog, reported the used range as
`A1:C1`, exposed `A1` and `C1` as bold while `B1` remained plain, and showed
`42` at `C1`. The retained worksheet XML independently showed the same opaque
shared-style reference on `A1` and `C1`. This verifies native interpretation of
one existing shared cell format and the public retarget path; it does not yet
certify shared-format definition editing, every formatting component, or a
native resave of the Litchi-produced output.

The eighth implementation slice adds a semantic rectangular selector to
`litchi-sheet`. `Area` accepts an A1 cell/range as the primary entry, raw
zero-based half-open bounds as a convenience, or a reusable checked `Rect`.
`Rect::from_a1`, `FromStr`, compact display, single-cell construction, and
union all preserve the grid bounds in the type. Parsing follows the checked-in
`[MS-OE376]` §2.1.1119 Office grammar: one cell or two ordered cell references
separated by a colon. `Sheet::cells` now accepts these selector forms directly,
so sparse traversal can use `sheet.cells("B2:F20")` without constructing an
intermediate range or risking an indexing panic.

Each lazily loaded worksheet now computes a borrowed `Extents` view in the same
pass that builds its sparse cell store. It keeps producer-declared `dimension`,
all stored cell records, known/unknown content cells, and cells with explicit
local styles distinct; `used` is the union of direct content and direct cell
formatting and explicitly excludes row/column defaults. The parser validates a
single bounded direct `dimension` before `sheetData`; missing `ref`, reversed or
out-of-grid references, duplicates, and invalid order are rejected. This keeps
a producer hint available for diagnostics without treating it as authoritative
cell content.

New workbooks carry the conventional `A1` dimension. When an ordinary cell
commit creates a stored record outside an existing declared range, the XML
surgery expands that range and preserves its prefix, unknown attributes, and
surrounding bytes. It never narrows a producer range on clear/removal and does
not invent a missing dimension, because row/column formatting and other used-
range contributors are not modeled yet. Inverse patches retain the original
part bytes exactly. The rewrite action and payload types no longer implement
`Clone`; commit moves the accepted ordered maps through row partitioning, so a
future regression cannot silently reintroduce whole cell-content copies at
that boundary. Semantic before/after patch states still clone their owned
values intentionally. The public edit example produced `A1:D4`, serialized the
matching dimension, and the public open example recovered the same four sparse
cell records and extent categories. These type and functional checks do not
establish an allocation or latency result.

Desktop Excel for macOS opened that eighth-slice artifact without a repair or
compatibility prompt, reported its used range as A1:D4, displayed the A1 text
and B2 number, and exposed C3 as a formula evaluating to 84. This is native
evidence that the tested Excel build accepts the conservatively expanded
dimension and stored records in this artifact. It does not certify dimension
shrinking, an Office resave of this exact artifact, other producers, or other
worksheet feature families.

The ninth implementation slice adds checked row-visibility reads and
transactional updates without conflating them with destructive row insertion
or deletion. `RowAt` accepts a
checked coordinate or a raw zero-based index, while `Sheet::row` returns a
borrowed `row::Row` for every logical grid row and preserves the distinction
between an implicit default row and an explicit stored `<row>` record.
`Sheet::rows` lazily visits only explicit records. The coordinate type is
re-exported as `RowIndex`, leaving the short `Row` name for the semantic view.
The parser accepts the SpreadsheetML boolean forms for `hidden` and rejects
malformed values. This follows the checked-in `[MS-OE376]` §2.1.1788 note that
row hiding is represented by the `row` element's `hidden` attribute rather than
the unrelated comment-row marker.

Transactions select the same checked row through `sheet.row(index)?` and expose
the short verbs `hide` and `show`. Hiding an implicit row creates one sparse
empty row record; showing an implicit row is a no-op. Touched row tags retain
their namespace prefix, unknown attributes, children, and cell payloads, while
untouched bytes remain exact. Protected sheets return a typed row-edit block.
`Change` and `Conflict` are now non-exhaustive data-bearing enums separating
cell and row effects: visibility edits on the same row conflict, but a cell
payload and row visibility on that row are independent and may join. Reversible
patches record `Missing` versus `Stored { hidden }` row states and restore the
original package bytes. Recalculation invalidation and calculation-chain
removal now occur only when primary cell content actually changes; row
visibility and style-only transactions preserve those unrelated parts.

The public row example produced an A1:A3 workbook with row index 1 hidden.
Desktop Excel for macOS opened it without repair, displayed row headers 1 and 3
as adjacent to hidden cells while omitting row 2, and reported A1:A3 as the used
range. Excel then changed A3 and resaved the workbook. ZIP validation passed,
and the public Litchi reader recovered the Excel-authored A3 text, all three
cell values, and the hidden row state. This certifies this hide/read/native-
resave path on the tested Excel build; it does not certify other row properties,
column visibility, protected-sheet permission exceptions, or structural shifts.

The tenth implementation slice adds checked column-property reads and
transactional visibility updates. `ColumnAt` accepts a reusable checked
coordinate or a raw zero-based index; `Sheet::column` returns a borrowed
semantic `Column` for every grid position, while `Sheet::columns` lazily visits
only logical columns covered by explicit effective property records. The short
`Column` name is reserved for that view and the coordinate is re-exported as
`ColumnIndex`. Width is a finite checked `Width` in the Office range `0..=255`;
style references, outline levels, and boolean flags are validated before the
view is published. `Sheet::column_style` exposes a shared resource handle or
the explicit default state without leaking the physical `style` index. Shared
style fan-out continues to count stored cells only and does not silently fold
column defaults into its existing contract.

Excel's interpretation of overlapping producer records was established with a
desktop probe before implementing the resolver. For a hidden B:D record
followed by a C-only record that omitted `hidden`, Excel displayed B and D as
hidden and C as visible, then normalized the file on save into three disjoint
records. The parser therefore treats the last matching `<col>` as the complete
effective record; omitted fields in that later record do not inherit from an
earlier overlap. A bounded fixed-grid interval assignment tree applies each
record without work proportional to its covered width and compacts the result
into disjoint borrowed ranges. This is an algorithmic bound, not a measured
latency or allocation claim.

Transactions use `sheet.column(index)?.hide()` and `show()`. Hiding an implicit
column creates a narrow sparse record; showing an implicit column is a no-op.
For an existing range, byte surgery identifies the last physical owner of each
edited coordinate, splits only that owner, changes only `min`, `max`, and
`hidden`, and retains widths, styles, outline state, unknown attributes,
namespace prefixes, untouched records, and `sheetData` bytes. Adjacent new
hidden columns coalesce. A split that would duplicate an unmodeled child
payload, or insertion into an unmodeled `cols` payload, returns a typed markup-
compatibility block. Protected sheets return a distinct typed column block.
The rewritten worksheet is semantically reparsed before the new snapshot is
published.

Column changes and conflicts have their own checked state and coordinate
variants. Two visibility effects on the same column conflict; cell, row, and
column facets remain orthogonal and can be prepared on separate `Edit` values
and joined without exposing locks. Patch inversion restores exact source part
bytes. Unit tests cover complete overlap replacement, malformed bounds and
properties, shared-style reference validation, prefix/attribute preservation,
lossless interval splitting, implicit insertion, protected sheets, extension
blocks, no-op filtering, inverse restoration, checked bounds, and disjoint
joins.

The public column example produced an A1:C1 workbook with column index 1
hidden. Desktop Excel for macOS opened it without a repair or compatibility
dialog, displayed A and C as adjacent headers, omitted B, and reported A1:C1 as
the used range. Excel changed C1, saved the workbook, and normalized the hidden
record to width zero. ZIP validation passed; the public Litchi reader recovered
all three cell values, the Excel-authored C1 text, and the hidden column state.
This certifies this hide/read/native-resave path on the tested Excel build; it
does not certify column width editing, outline/style authoring, structural
column shifts, protected-sheet permission exceptions, or other Office builds.

The eleventh implementation slice completes the first workbook/row/column
visibility family with selector-first sheet-tab updates. Workbook-level
`edit.tab(name_or_position)?` deliberately accepts every sheet kind, while
worksheet-only `edit.sheet(...)` remains the cell and grid-property entry
point. A borrowed `TabEdit` exposes the short verbs `show`, `hide`, and
`very_hide`; the last maps to SpreadsheetML `veryHidden`, which Excel omits
from its ordinary Unhide dialog. Read-side `Visibility` retains unknown
producer values and adds `is_visible`, `is_hidden`, and `is_very_hidden` for
the recognized cases. This keeps native sheet IDs and relationship IDs out of
the facade without forcing routine callers to match an enum for a boolean
query.

The transaction computes the final visibility set before touching package
bytes. It refuses any result without a recognized visible tab, treating an
unknown producer state as insufficient proof of visibility. If the active tab
would become hidden, the transaction selects the next final visible tab in
workbook order with wraparound; if the source active position is absent or
invalid internally, it falls back to the first visible tab. That arithmetic is
total even for an empty catalog. The checked-in `[MS-XLSB]` §2.5.143 state
table corroborates visible, hidden, and very-hidden semantics, while
`[MS-OE376]` §2.1.622 records Office's 0-through-32766 `activeTab` bound. An
explicit `show` can replace an unknown source state, but ordinary commits never
guess that the unknown value already means visible.

Workbook XML surgery matches the selected semantic sheet through its private
relationship ID, regenerates only touched direct `sheet` start tags, and
removes `state` entirely for the default visible state. Unknown attributes,
namespace prefixes, child payloads, interstitial XML, subsequent workbook
views, and every untouched byte remain preserved. Active relocation updates
the first direct `workbookView`; an absent view is inserted into an existing
safe `bookViews` container or a prefixed `bookViews` is inserted in schema
order before `sheets`. The same atomic effect synchronizes
`sheetView[@workbookViewId=0]/@tabSelected` in the old and new active sheet
parts. Existing view payloads remain intact. When a new selected view is
needed, the editor inserts the minimal prefixed structure after the last
schema predecessor (`sheetPr`/`dimension`) for worksheet/dialog/macro sheets
or after `sheetPr` for chart sheets. This
implements the SpreadsheetML rule documented by Microsoft's
[`SheetView` reference](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.sheetview):
a singly selected sheet's `tabSelected` state should agree with
`activeTab`; multiple tabs may be selected, but only one is active.

Visibility mutation under workbook structure protection returns a typed block.
Effective sheets without a direct editable slot, or selection/view insertion
beside an unmodeled compatibility alternative, also return a typed block rather
than allowing a lossy rewrite. Unrelated compatibility content remains exact; a
structure-protection element nested in such content is still detected
conservatively. The final workbook catalog and every rewritten per-sheet view
are reparsed to verify requested visibility, active position, and selection
bits before publishing the immutable snapshot.

Tab visibility is an orthogonal transaction facet. Two prepared visibility
effects on the same tab conflict, while a tab effect can join cell, row, or
column work on that sheet without exposing a lock. A single composed workbook
part delta carries both tab changes and recalculation invalidation, while a
single part composer folds selection synchronization into any simultaneous
worksheet edit. This avoids duplicate source expectations for either part.
`Change::Visibility` records the semantic before/after states and checked
position; inverse patches restore the exact source bytes. Focused tests cover
name and numeric selectors,
visible/hidden/VeryHidden transitions, last-visible refusal, active relocation,
unknown-state repair, no-op filtering, structure protection, compatibility
blocks, prefix and unknown-attribute retention, non-worksheet tabs, per-sheet
selection synchronization and schema-ordered insertion, reversible bytes, tab
conflicts, tab/cell joins, and workbook/recalculation composition. The
rewriter reserves the exact computed output length and borrows tab names and
relationship IDs for the duration of the synchronous rewrite rather than
copying them into another owned catalog. These are functional and algorithmic
properties, not allocation or latency measurements.

The first native Excel probe prevented an incomplete implementation from being
accepted. Excel opened the initial hidden-tab artifact without repair and hid
Sheet2, but displayed `[Group]` in the title because the workbook's
`activeTab` had moved while Sheet2's `tabSelected="1"` remained stale. The
dependency was added to the transaction and the artifact regenerated from an
Excel-authored two-sheet baseline. Desktop Excel for macOS 16.110.2 (build
16.110.26062818) then opened the corrected ordinary-hidden artifact without
repair, displayed only Sheet1 with no grouped-sheet marker, edited C1, and
resaved. ZIP validation passed; the
public reader recovered the Excel-authored C1 text, Sheet1 as visible, and
Sheet2 as hidden. Excel also opened the VeryHidden artifact without repair;
its ordinary **Unhide Sheet** command was disabled, as expected when no
ordinary-hidden sheet exists. After another C1 edit and resave, ZIP validation
passed and the reader recovered both the Excel-authored text and Sheet2 as
VeryHidden. This certifies these two active-tab hide/edit/resave/reverse-read
paths on the tested Excel build. It does not certify grouped-selection
preservation under every visibility transition, protected-workbook unlocking,
other sheet kinds in native Office, or other Office builds.

The twelfth slice makes active-sheet selection an explicit semantic operation
instead of leaving it only as a visibility side effect. The same selector-first
proxy now supports `edit.tab(name_or_position)?.activate()`, and `Sheet::is_active`
answers the common read query without exposing `activeTab` or requiring handle
comparison. Activation applies to every recognized sheet kind. It is a
workbook-global transaction facet: a later call in one mutable transaction
replaces the earlier intent, while two independently prepared activation
intents conflict and return the rejected edit. Activation remains orthogonal
to cell, row, column, and visibility effects, so independently prepared work
can join when those facets do not overlap.

The transaction validates the requested target against the complete final
visibility plan before changing bytes. Hidden, very-hidden, and unknown-state
targets return `TabEditBlock::NotVisible`; callers can intentionally repair and
activate one in the same concise operation with `tab.show().activate()`. An
explicit target takes precedence over the automatic relocation used when an
active tab is hidden, but cannot override the invariant that an active tab is
visible. Positions outside Office's checked `0..=32_766` `activeTab` range in
`[MS-OE376]` section 2.1.622 return the typed `ActiveTabLimit` block. Structure
protection still blocks hide/show
because those mutate workbook structure, but it does not block selection of an
already visible sheet; this matches Microsoft's description of protected
structure in the [`Workbook.Protect`](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.protect)
and [`Workbook.ProtectStructure`](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.protectstructure)
references.

An effective activation reuses the same lossless workbook-view and sheet-view
rewriters, including prefix preservation, schema-ordered insertion,
compatibility blocking, reparsing, and exact inverse bytes. It clears the old
active sheet's view-zero selection and selects the new active sheet while
leaving unrelated selected tabs untouched; `activate` therefore does not
silently destroy a producer-authored grouped selection. `Change::Active`
records compact `ActiveTab` values with semantic names and checked positions,
including visibility-driven relocation, so patches no longer hide that
secondary effect. A workbook-part delta composes activation with visibility
and recalculation, and a sheet-part delta composes selection with simultaneous
cell edits. Regression tests cover name/numeric lookup, worksheet and chart
tabs, no-op filtering, last-call replacement, hidden-target refusal,
show-and-activate repair, contradictory plans, protected workbooks, global
activation conflicts, orthogonal joins, composed part counts, typed Office
bounds, semantic patch inspection, and byte-exact inversion.

Computer Use verification exercised both public operations in Microsoft Excel
for Mac 16.110.2 (build 16.110.26062818). An active-only output derived from an
Excel-authored two-worksheet workbook opened with Sheet1 selected, both tabs
visible, no repair or compatibility prompt, and no grouped-workbook marker.
After entering `Active Sheet1 survived Excel` in C1 and saving in Excel, the
public reader still reported Sheet1 active and returned that value. A
`show().activate()` output derived from an Excel-resaved workbook whose Sheet2
was very hidden likewise opened without a prompt, showed both tabs, and
selected only Sheet2. After entering `Shown and active survived Excel` in A1
and saving, the public reader reported Sheet2 visible and active and returned
the value. Both Excel-resaved packages passed ZIP integrity checks. This native
evidence covers ordinary worksheets and the first workbook view on this one
macOS build; it does not certify chart sheets, multi-view windows, or grouped
selection editing across Office versions.

The thirteenth slice adds lossless workbook-tab reordering. Semantic anchors
are the ordinary entry points: `edit.move_before(sheet, anchor)?` and
`edit.move_after(sheet, anchor)?` accept the same name-or-position selectors as
reads, return `Option` for a missing source or anchor, and apply to every
recognized sheet kind. `edit.move_to(sheet, position)?` retains checked raw
zero-based positioning for import and algorithmic workflows. Multiple moves in
one transaction operate on the pending order, retain their semantic sequence,
and can cancel to an empty patch. No native `sheetId` or relationship ID enters
the public API.

Order is one workbook-global conflict facet because independently prepared
permutations can change each other's anchor meaning. It remains orthogonal to
activation, visibility, and worksheet payload/property facets, so those edits
can be prepared concurrently and joined. Each effective operation yields a
`Change::Move` with its semantic sheet name and checked before/after positions.
Inverse patches reverse dependent operation, part, and graph-delta order before
swapping each effect, making multi-move undo exact rather than relying on
commutativity. Planning allocates one compact identity permutation lazily,
moves identities in place, and borrows names and physical relationship IDs only
inside the low-level rewrite boundary.

The physical rewrite moves each complete `sheet` element without rebuilding or
copying any sheet part. In the same single workbook-part delta it remaps
`activeTab` and `firstSheet` for every direct `workbookView`, including omitted
zero defaults and Office's `4294967286` `firstSheet` sentinel, and remaps every
direct sheet-local `definedName@localSheetId`. These positional fields follow
the checked-in `[MS-OE376]` workbook-view and defined-name notes and the
Open XML SDK references for
[`WorkbookView.FirstSheet`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.workbookview.firstsheet?view=openxml-3.0.1)
and
[`DefinedName.LocalSheetId`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.definedname.localsheetid?view=openxml-3.0.1).
In contrast,
`customWorkbookView@activeSheetId` remains exact because Microsoft defines it
as the native `sheetId`, not a workbook-order position; see the
[`CustomWorkbookView` reference](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.customworkbookview?view=openxml-3.0.1).
The active semantic identity is preserved across an ordinary reorder while its
reported position changes. An explicit activation still wins, including the
case where old and new active sheets occupy the same numeric position after a
swap, and only an identity change rewrites per-sheet selection bits.

Structure protection, revision-header tracking, unknown direct catalog/view/
defined-name payloads, compatibility alternatives that can own an order
dependency, invalid secondary-window positions, and active positions outside
Office's range produce typed blocks before bytes change. Unrelated direct
workbook `AlternateContent`, including Excel's `x15ac:absPath`, stays byte-exact
and does not prevent an otherwise modeled reorder. The rewriter validates a
full permutation with O(n) hash indexes, reserves mappings and output capacity
explicitly, reparses the effective catalog, and verifies relationship order,
visibility, active position, defined-name count/content, and every remapped
scope before publishing. Tests cover semantic and numeric selectors, all
source/destination pairs in a three-tab workbook, multi-move sequences,
cancellation and conflict-free composition after cancellation, worksheets and
chart sheets, local/global names, multiple views and default/sentinel fields,
same-position active-identity replacement, composed cell/visibility/activation
changes, global conflicts, protection, revision tracking, affected and
unrelated compatibility payloads, and byte-exact inversion.

Computer Use verification opened a Litchi-reordered, Excel-authored workbook
in Microsoft Excel for Mac 16.110.2 (build 16.110.26062818). Excel displayed
Sheet2 as the selected first tab and Sheet1 as the second tab without a repair,
compatibility, or grouped-workbook warning. After entering
`Reordered first survived Excel` in Sheet2 A1 and saving, the package passed a
ZIP integrity check and the public Litchi reader reported Sheet2 visible and
active at position zero, returned the new value, and preserved Sheet1's three
text cells and hidden middle column at position one. This native evidence
covers two ordinary worksheets, one workbook view, one reorder, and one macOS
Excel build; it does not certify chart sheets, local defined names, multiple
views, every markup-compatibility alternative, or other Office builds.

The fourteenth slice adds dependency-aware worksheet rename. The concise entry
is `edit.tab(selector)?.rename(name)?`; it accepts the same semantic name or
checked zero-based selectors as other tab operations. Borrowed strings are
checked and copied once, while an owned `String` or prevalidated
`sheet::Name` moves into the transaction. Source selectors remain stable for
the transaction, and multiple pending renames are interpreted simultaneously,
so swaps do not cascade through each other's formulas. `Change::Rename`
records the checked source position and case-preserving before/after names;
name is its own per-sheet join facet and remains orthogonal to cell, property,
visibility, activation, and order effects.

`sheet::Name` enforces the checked-in `[MS-OE376]` Office profile: a nonempty
maximum of 31 characters, no leading/trailing apostrophe, no NUL, ETX,
`* / : ? [ \\ ]`, and only XML 1.0-representable characters. Opening also
enforces the 32,767-sheet limit, native sheet IDs in `1..=65_534`, relationship
IDs through 255 characters, and case-insensitive uniqueness. Identity keys use
canonical normalization plus locale-independent full Unicode case folding
instead of `eq_ignore_ascii_case`; equality covers length-changing folds such
as `Straße`/`STRASSE` and canonically equivalent spellings. The selected
`caseless` 0.2.2 table is Unicode 16.0 and the selected
`unicode-normalization` 0.1.25 table is Unicode 17.0; both compile below the
workspace MSRV. This is a deterministic format-identity rule, not an ambient
locale decision.

Rename preserves native `sheetId`, relationship, part URI, sheet kind, order,
visibility, selection, and every sheet-part byte unless another dependency in
that part changes. A single-pass formula scanner rewrites local unquoted,
apostrophe-quoted, doubled-apostrophe, `[0]` current-workbook, and 3-D sheet
prefixes. It evaluates all source names before emitting any target name,
quotes target spellings only when the formula grammar requires it, and leaves
string constants, structured references, nonzero external-workbook indexes,
external paths/books, and VBA bytes inert.

The package transaction recognizes workbook defined names; worksheet cell,
conditional-format, and data-validation formulas; table calculated/totals
formulas; DrawingML chart formulas; Excel sparkline formulas; internal
hyperlink locations; pivot-cache `worksheetSource@sheet`; and extended
property `TitlesOfParts` sheet/named-range values. SpreadsheetML data-
consolidation `dataRef@sheet` is another checked direct-name carrier. Formula-
like fields not in the modeled set, including legacy VML `Fmla*` payloads,
produce `RenameBlocked::UnmodeledReference` when they contain the source
identity. Expanded names distinguish recognized formulas from namespace
spoofs, while the ordered extended-property sheet slots are checked against
the workbook catalog so an equal named-range title is not mistaken for a
sheet. A matching reference inside `mc:AlternateContent`
produces a distinct markup-compatibility block. External-link parts are
excluded deliberately. Workbook structure protection, revision-header
tracking, and signatures keep their existing typed blocks.

The rewriter plans byte spans and allocates a replacement part only after a
recognized match. Clean OPC payloads remain shared `Arc` owners; changed
worksheet formulas compose over simultaneous cell edits, and workbook formula
changes compose with catalog state, activation, order, and recalculation in
one part delta. XML nesting and simultaneous decoded-reference capture have
explicit edit budgets. Final catalog relationship/name pairs are reparsed and
checked before publication. Tests cover typed name failures, Unicode/canonical
collisions, case-only lookup, simultaneous swaps, owned/prevalidated input,
join facets, formulas followed by both A1 references and defined names, 3-D
references, workbook index zero, excluded external references, hyperlinks,
pivot sources, tables, charts, extended properties, compatibility and unknown
blocks, preservation of native identities, source immutability, and exact
inverse part bytes.

For native evidence, Microsoft Excel for Mac 16.110.2 authored and saved a
two-sheet baseline whose `Sheet2!A2` formula referred to `Sheet1!A1`. The
public `tabs` example renamed `Sheet1` to `Input Data`; Excel opened the result
without a repair or compatibility dialog, displayed the cached value, and
reported `='Input Data'!A1` in the formula bar. Excel then added a marker and
resaved that Litchi-produced workbook. ZIP validation passed, and the public
reader recovered the renamed tab, rewritten formula, Excel-authored marker,
and an existing hidden column. This certifies one local scalar-reference
rename/open/resave path on that Excel build. It does not certify every modeled
reference carrier, name grammar edge, external-link behavior, signed or
protected packages, macros, other Office applications, or other Office builds.

The fifteenth slice adds transactional worksheet creation. The concise entry
is `let mut sheet = edit.add(name)?`; the returned `NewSheet<'edit>` is a
borrowed transaction-local capability, so it can immediately use the same
short `set`, `clear`, `remove`, `style`, `row`, `column`, visibility, rename,
and activation verbs as an existing worksheet without exposing a native
`sheetId`, relationship ID, part URI, lock, or generic identity parameter.
Names accept borrowed strings, owned strings, or prevalidated `sheet::Name`
values. Creation is tail-only in this slice: existing reorder is applied first
and new worksheets retain call order. Semantic before/after insertion remains
future work rather than being approximated with a raw physical ID.

Commit validates the final Unicode-caseless name set and the 32,767-sheet
limit before graph mutation. The sheet-count, native-ID, relationship-ID, and
active-tab bounds come from `[MS-OE376]` sections 2.1.612, 2.1.613, and
2.1.622(c). Commit then allocates the lowest unused native sheet ID, workbook
relationship ID, and non-conflicting worksheet part name at the low-level
boundary. Allocation is deterministic, recognizes nonstandard IDs and gaps,
and derives strict versus transitional worksheet namespaces and relationship
types from the workbook root. `Change::Create`/`Change::Remove`
carry only the developer name, checked position, and visibility; the latter is
produced only by inverse patches in this slice. The private graph delta owns
the new part and relationship. Forward replay checks that both are absent;
inverse replay checks their complete expected identity and bytes before
removal.

The new worksheet is built and edited before publication, then reparsed with
the ordinary worksheet and shared-style validators. Formula-bearing creates
invalidate calculation properties and remove one exclusively owned calculation
chain through the same reversible graph transaction. Activating the new sheet
updates the first workbook view, clears the prior sheet's selection, and emits
the new sheet's selected view; creating a hidden active sheet is rejected by
the final-state invariant. Simultaneous existing-sheet renames rewrite local
references inside the newly created worksheet as well as existing parts.
When existing tab order is unchanged, recognized extended properties insert
new titles after the existing sheet-title prefix, update vector size and a
standard Worksheet heading count, and retain named-range titles. A simultaneous
reorder skips creation-metadata synchronization; absent another edit to that
part, it remains byte-exact until an order-aware splice is implemented. Missing,
stale, or nonstandard optional layouts are likewise preserved instead of
guessed.

Independent edits may append different names and populate their own new
worksheets concurrently, then join in explicit join order. Unicode-equivalent
new names, active-tab effects, and conflicting rename targets return the same
structured conflict families as existing sheets. Tests cover one-transaction
cell/formula/row/column creation, activation and hidden-state refusal, owned
names, Unicode canonical collisions, disjoint joins, strict dialects, gapped
native identity allocation, prefixed and empty sheet catalogs, protected and
compatibility-owned catalogs, rename/formula composition, extended-property
ordering, source immutability, source-checked forward replay, and byte-exact
inverse restoration.

Computer Use then exercised this exact slice in Microsoft Excel for Mac
16.110.2 (build 16.110.26062818). Excel opened the Litchi-generated workbook
without a repair or compatibility dialog, exposed all three expected tabs,
made `Active Data` active, displayed `42` at A1, retained `Summary!A1`,
calculated `Summary!B2` as `2`, and showed column C as hidden. Excel accepted
`Excel resave marker` at `Summary!D2` and saved the workbook without warning.
The resaved archive passed ZIP integrity validation, and the public `open`
example recovered all three tabs, the active Summary tab, the A1 text, the
`1+1` formula with cached value `2`, the D2 marker, hidden column C, and the
numeric `42`. This certifies one local append/populate/column-hide/activate/open/
edit/resave/reverse-read path on that Excel build. It does not certify insertion
at arbitrary positions, worksheet deletion, every cell or style family, other
Office builds, or performance.

The sixteenth slice adds conservative transactional worksheet deletion. The
selector-first entry is `edit.remove("Scratch")?`, returning
`Result<Option<&mut Edit>>`; case-insensitive developer names remain the main
path and checked zero-based positions remain available. It never exposes a
native `sheetId`, relationship ID, or part URI. Multiple distinct removals can
be collected directly or joined from independent edits. Removing the same
sheet conflicts deterministically. This slice is deliberately worksheet-only,
and it refuses deletion combined with creation, rename, reorder, activation,
visibility, cell, row, or column mutation as `RemoveBlock::MixedEdit`. That
typed boundary remains until name reuse, reference disposition, and mixed
final-state semantics are represented rather than inferred.

Commit proves that at least one sheet and one visible tab survive, preserves a
retained active tab, or selects the nearest visible successor and then the
nearest visible predecessor. It removes exact catalog slots; remaps
`activeTab` and `firstSheet` in every modeled workbook view, including the
special `firstSheet` sentinel; drops defined names scoped to removed sheets;
shifts surviving local scopes; and removes an empty `definedNames` container.
Recognized extended properties lose only the corresponding leading
`TitlesOfParts` entries and receive corrected vector and standard Worksheet
counts. Missing, stale, or producer-specific optional metadata remains
byte-exact. Workbook protection, revision tracking, unknown catalog payload,
and markup-compatibility-owned catalog state retain their existing typed
refusals. A custom workbook view whose `activeSheetId` equals the removed
native identity is a modeled incoming dependency, consistent with
`[MS-OE376]` section 2.1.600(q), and is refused rather than left dangling.

Dependency validation runs over the final planned bytes of every XML part
reachable after detaching the selected worksheet relationships and the
calculation chain. It includes reachable Custom XML rather than assuming that
only `/xl` parts matter, while deliberately excluding external-link parts.
All removal targets share one XML pass per retained part instead of reparsing
the package graph once per target.
Recognized direct references include local formula prefixes, implicit members
of 3-D sheet spans, hyperlink locations, pivot and consolidation sources,
embedded-object extents, and custom-workbook-view identities. Nonzero external
workbook prefixes are not local dependencies. Runtime reference construction
through `INDIRECT` or `EVALUATE`, unknown formula-like producer fields, and
matching names under markup-compatibility choices produce distinct typed
refusals. A VBA project blocks deletion because its dynamic sheet access cannot
be proven safe, and any additional OPC relationship targeting the worksheet
part blocks removal.

The forward patch now records `Change::Remove`. Its private graph delta owns
the exact worksheet part, its outgoing relationships, and the removed workbook
relationship, so source-checked replay and inverse restoration do not copy the
payload and can reproduce the original package bytes. OPC part-name checks use
the required ASCII case-insensitive equivalence. Child resources reachable
only from the removed worksheet are intentionally left as orphans; recursive
resource disposal remains the separate explicit `gc` operation from ADR 0003.
Formula-bearing state invalidates workbook calculation and removes one
exclusively owned calculation chain through the same reversible transaction.
The source snapshot remains immutable throughout planning and publication.

Tests cover name and numeric selectors, missing lookup, active-tab relocation,
multiple joined removals, exact forward replay and inverse bytes, static and
3-D formulas, runtime indirection, Custom XML producer fields, custom workbook
views, macro projects, extra incoming relationships, case-equivalent OPC
targets, last-sheet and last-visible invariants, mixed-edit refusal, local-name
scope rewrites, secondary workbook views, optional property synchronization,
catalog compatibility blocks, and scanner dependency classes.

Computer Use exercised the public `remove_sheet` example in Microsoft Excel
for Mac 16.110.2 (build 16.110.26062818), Office LTSC Standard for Mac 2024.
Excel opened the generated workbook without a repair or compatibility dialog,
showed exactly `Sheet1` and `Results`, made `Results` active, and displayed the
retained `A1` text and `B2` value. Excel accepted `Excel removal resave marker`
at `Results!C3` and saved without warning. The resaved archive passed ZIP
integrity validation, and the public `open` example recovered both tabs, the
active Results identity, the original values, and the marker. This certifies
one local remove/active-relocation/open/edit/resave/reverse-read path on that
Excel build. It does not certify mixed deletion plans, non-worksheet tabs,
recursive garbage collection, every dependency carrier, other Office builds,
or performance.

The seventeenth slice completes semantic worksheet insertion. The concise
entries are `edit.add_before(name, anchor)?` and
`edit.add_after(name, anchor)?`; both return `Result<Option<NewSheet<'edit>>>`
and use the ordinary case-insensitive developer-name or checked zero-based
selector. `None` means that the anchor did not resolve in the immutable source
snapshot. Native sheet IDs, relationship IDs, part names, and lock wrappers
remain private. `add(name)?` continues to mean tail insertion. Repeated
before/after additions at one anchor retain call order, and joined independent
edits retain explicit left-then-right join order.

Anchors are stable base-sheet identities rather than transient positions. A
pending base reorder is applied first; anchored additions then surround the
same identities in that effective order, and tail additions follow them. This
defines composition without exposing a more complex public ordering type.
`NewSheet::position()` is an intentionally current projection while its
borrowed capability is live: later structural intents may shift it after that
borrow ends. `Change::Create` records the authoritative final position.
`Change::Move` continues to describe the base-order phase, making patch event
ordering deterministic rather than pretending inserted sheets participated in
an earlier base move.

One private checked `FinalOrder` is the source of truth for name-collision
diagnostics, active-tab placement, sheet-view selection, creation changes,
catalog verification, defined-name scopes, and extended properties. It stores
compact `Base(index)`/`Added(index)` identities plus direct position maps. New
worksheet parts and relationships are still allocated deterministically and
physically appended at the low-level boundary. The catalog is then losslessly
reordered by private relationship identity only when semantic placement
requires it. The catalog rewriter remaps every modeled `activeTab`,
`firstSheet`, and sheet-local `localSheetId`; a requested active sheet is
applied only after the final order exists. The result is reparsed and every
final relationship slot, created sheet identity/state, active position, and
defined-name scope is checked before publication.

Recognized extended properties now synchronize their complete worksheet-title
prefix for reorder and insertion together. Existing title elements move as
complete byte spans, new escaped titles are synthesized, vector and standard
Worksheet counts are updated, and following named-range titles remain
byte-exact. Missing, stale, or producer-specific layouts remain untouched
rather than being guessed. Scope verification uses the already
reference-rewritten workbook as its baseline, so a simultaneous rename,
insertion, and reorder verifies both the formula rewrite and structural scope
mapping independently.

Tests cover name and numeric anchors, missing lookup, repeated before/after
order, tail composition, base reorder composition, deterministic joins,
population and activation, final create positions, local defined-name scope
shifts, named-range-title preservation, strict graph allocation inherited from
creation, source immutability, source-checked forward replay, and byte-exact
inverse restoration. These functional tests do not establish allocation,
latency, contention, cache, or scaling claims; those require the measurement
program in ADR 0005.

Computer Use then exercised the exact public `insert_sheet` artifact in
Microsoft Excel for Mac 16.110.2 (build 26062818), Office LTSC Standard for Mac
2024. Excel opened it without a repair or compatibility prompt and showed the
ordered tabs `Inputs`, `Sheet1`, `Results`, and `Archive`, with `Results`
active. It displayed `Revenue` and `120` on `Inputs`, and calculated
`Results!B1` to `132` from `=Inputs!B1*1.1`. Excel accepted
`Excel insertion resave marker` at `Results!C2` and saved without warning. The
resaved archive passed ZIP integrity validation, and the public `open` example
recovered the same four-tab order, active identity, source values, formula and
cached `132`, marker, and tail-sheet text. This certifies one local
before/after/tail insertion, population, formula calculation, activation,
open/edit/resave/reverse-read path on that build. It does not certify every
multi-anchor/reorder composition, optional metadata layout, other Office
applications or builds, or performance.

The eighteenth slice expands column CRUD from visibility into a typed,
orthogonal property surface. The primary selector is now an A1 column label,
so `sheet.column("B")?` and `edit.sheet("Sheet1")?.column("B")?` are the normal
lookup and mutation paths. Checked `ColumnIndex` values and raw zero-based
indexes remain available without exposing one-based SpreadsheetML `min`/`max`
fields. `ColumnIndex::from_a1` accepts case-insensitive labels with an optional
absolute marker and rejects malformed or out-of-grid values; `a1` produces the
compact canonical label. Every selector remains fallible rather than using the
panicking `Index` trait.

`column::{Width, Outline, Props, State}` keeps the public names short and the
wire invariants in types. Widths must be finite and inside `0..=255`; outlines
must be inside `0..=7`. `Props` separates width, shared style, outline, hidden,
best-fit, custom-width, phonetic, and collapsed state without publishing a
physical style ID. `State` preserves the difference between an implicit column
and a stored property record. Transactions expose `width`/`reset_width`,
`hide`/`show`, `best_fit`/`fixed`, `outline`, `collapse`/`expand`, and
`show_phonetic`/`hide_phonetic`. Inputs are checked before an action enters the
plan. Independent facet writes to the same column can join, while two writes to
one facet report a deterministic conflict. The accepted transaction moves the
incoming action map; no public `Arc<RwLock<...>>` or last-writer-wins rule is
introduced.

Worksheet surgery finds the last effective physical owner of an edited column,
splits only the necessary compact ranges, and changes only the selected
attributes. It preserves unrelated known and unknown attributes, prefixes,
children, and untouched bytes; default-only operations on implicit columns are
no-ops, materializing operations create sparse records, and adjacent identical
actions coalesce. Setting a width also establishes `customWidth`; resetting it
removes both attributes. Complete before/after column states make patches
reversible. A column carrying an explicit shared style also carries that
resource's lineage and byte guard, so replay against a changed style table is
rejected instead of reinterpreting an opaque key. Tests cover A1 and numeric
selectors, checked bounds, every property facet, lossless splits, reset and
inverse behavior, independent-facet joins, same-facet conflicts, shared-style
replay/rebinding, and malformed input. These are functional and type checks,
not allocation, latency, CPU, cache, or contention measurements.

Computer Use exercised the exact public `columns` artifact in Microsoft Excel
for Mac 16.110.2 (build 26062818), Office LTSC Standard for Mac 2024. Excel
opened it without a repair or compatibility prompt, displayed B as a wide
column, hid C, and interpreted D's level-one collapsed outline by omitting it
and showing the outline expansion control before E. Excel's Column Width dialog
reported `23.17` for the stored OOXML width `24`; this is the application's
display-unit normalization, and the saved package retained `24`. Excel accepted
`Excel column layout resave marker` at A2 and saved without warning. The resaved
archive passed ZIP integrity validation. The public reader recovered the
marker, B with width 24, C hidden with Excel-normalized width zero, D hidden
with width zero/outline one/collapsed, and Excel's adjacent E collapse marker
record. This certifies one local column-width/visibility/outline/open/edit/
resave/reverse-read path on that build. It does not certify every producer
normalization, column shared-style authoring, structural column shifts, other
Office applications or builds, or performance.

The nineteenth slice expands row CRUD from visibility into the corresponding
typed, orthogonal layout surface. `row::{Height, Outline, Props, State}` keeps
the public vocabulary compact; `Outline` is the same checked type used by
columns rather than a duplicate wire wrapper. `Height` stores the exact finite
point value and admits only Excel's `0..=409` range, while `Outline` admits only
`0..=7`. The parser validates the checked-in `[MS-OE376]` row profile, including
one-based row order, style references through 65,490, outline depth, and every
SpreadsheetML boolean spelling. `Sheet::row_style` exposes default or shared
resource identity without a physical style index. `row::Props` and the borrowed
row view separately report height, custom-height, shared style, outline,
hidden, collapsed, thick-top, thick-bottom, phonetic, and custom-format facets;
an implicit logical row remains distinct from a stored `<row>` record.

Transactions retain the concise numeric row selector and add
`height`/`reset_height`, `outline`, `hide`/`show`, `collapse`/`expand`,
`thick_top`/`normal_top`, `thick_bottom`/`normal_bottom`, and
`show_phonetic`/`hide_phonetic`. Inputs are fully checked before an action is
inserted. Independent edits to different facets of one row join; two writes to
the same facet conflict deterministically. The compact private action stores
orthogonal optional effects rather than a lock-bearing public object. Complete
before/after states retain shared-style lineage and byte guards, so patch replay
against another style table is rejected and replay against a byte-identical
snapshot safely rebinds the opaque handle.

Worksheet surgery changes only selected direct row attributes and preserves
cells, child payloads, namespace prefixes, unknown attributes, and untouched
bytes. Setting a height also sets `customHeight`; resetting removes both.
Default-only operations on an implicit row are no-ops, while materializing
operations create a sparse empty row record. Every rewritten worksheet is
reparsed before publication, and inverse patches restore exact source part
bytes. Tests cover checked and constant-evaluated bounds, malformed row
attributes, all writable facets, unrelated-attribute preservation, sparse
materialization, reset and inversion, independent-facet joins, same-facet
conflicts, and shared-style replay/rebinding. These are type and functional
checks, not allocation, latency, CPU, cache, or contention measurements.

Computer Use exercised the exact public `rows` artifact in Microsoft Excel for
Mac 16.110.2 (build 26062818), Office LTSC Standard for Mac 2024. Excel opened
it without a repair or compatibility prompt, rendered row 2 taller, omitted
hidden row 3, and showed the level-one outline control for collapsed row 4.
On this build, the Row Height dialog reported `40` for the source OOXML
`ht="30"`; Excel normalized the saved XML to `ht="40"`, reopened it at the
same displayed value, and retained the visible size. This observed producer
normalization is recorded rather than treated as a cross-producer numeric
stability guarantee. Excel accepted `Excel row-layout resave marker` at B2 and
saved without warning. The resaved archive passed ZIP integrity validation,
and the public reader recovered A1:B4, the marker, row 2's normalized height and
custom-height state, row 3's hidden state, and row 4's outline/collapsed state.
This certifies one local height/hide/outline/open/edit/resave/reverse-read path
on that build. It does not certify every height normalization, thick-edge or
phonetic rendering, row shared-style authoring, structural row shifts, other
Office applications or builds, or performance.

The twentieth slice makes row and column shared-style defaults writable without
publishing a physical `cellXfs` index. `RowEdit::style` and
`ColumnEdit::style` accept an existing lineage-checked `Style`; handles from an
unrelated style table are rejected before they mutate the plan. The paired
`reset_style` verbs remove only the selected grid-default reference. Row and
column style effects remain independent from height, width, visibility,
outline, and other property facets, so disjoint transaction plans join without
locks while two style writes to one logical record conflict deterministically.
The same borrowed editors work for existing and transaction-local new sheets.

Worksheet surgery writes row `s` together with its derived `customFormat`
marker, writes column `style`, and removes those exact attributes on reset. It
continues to preserve unrelated known and unknown attributes, child payloads,
namespace spelling, compact effective column ranges, and untouched bytes.
Style set materializes a sparse row or column property; reset on an implicit
record remains a no-op. A style-only implicit column is the deliberate
exception: the native Excel probe interpreted a style-only record with omitted
`col/@width` as zero-width. Commit therefore returns
`ColumnEditBlock::StyleNeedsWidth` unless the column already owns a width or the
same transaction stages one. This prevents a seemingly harmless format edit
from collapsing a visible column. Complete row/column patch states carry the
opaque style identity, source style-byte guard, and target-lineage rebinding;
inverse application restores the exact source package bytes.

Tests cover set/reset, row custom-format derivation, compact column splitting,
sparse materialization/no-op behavior, the implicit-column width block,
foreign-lineage rejection without plan mutation, existing and new sheets,
exact inverse restoration, byte-identical replay/rebinding, independent-facet
joins, and same-style-facet conflicts. The public `grid_styles` example obtains
the shared style semantically from `A1`, applies it to row 4, and applies it to
column D while setting a checked width of 12. It saves through the ordinary
transaction facade and verifies the committed resource identities without
using a numeric style ID.

Computer Use exercised that exact example output in Microsoft Excel for Mac
16.110.2 (build 26062818), Office LTSC Standard for Mac 2024. The first native
probe was intentionally treated as a failed verification: an implicit
style-only column became zero-width, and already-stored cells retained their
local formatting instead of being retroactively changed by a grid default.
That evidence produced the typed width block and clarified the documented
layer semantics. The corrected artifact opened without a repair or
compatibility prompt, kept width-12 column D visible, and retained row 4's
default style. Creating `row default verified` at previously missing A4 and
`column default verified` at previously missing D1 made Excel report and render
both values in the bold shared style originally authored on A1. Excel saved the
workbook without warning; the resaved archive passed ZIP integrity validation.
The public reader recovered declared/used/stored bounds A1:D4, both marker
cells, row 4 with `customFormat` and its shared style, and column D with width
12 and its shared style. The resaved XML retained `row/@s` plus
`customFormat`, `col/@style` plus width, and explicit shared-style references
on the newly created cells. This certifies one local row-default/column-default/
new-cell-inheritance/open/edit/resave/reverse-read path on that build. It does
not claim that grid defaults rewrite existing local cell styles, certify every
style family or producer normalization, cover structural row/column shifts,
other Office applications or builds, or establish performance.

The twenty-first slice makes worksheet-wide grid defaults and typographic
descent first-class. The focused `layout` module contains checked `Height`,
`Width`, and `Descent` values plus the complete modeled `Defaults` record. The
types remain separate from `row::Height` and `column::Width` because their wire
domains differ: `[MS-OE376]` §2.1.678 permits `baseColWidth` through 255 and
`defaultColWidth` in the finite `0..65536` interval, while the required default
row height and Microsoft `x14ac:dyDescent` are finite and non-negative. The
numeric wrappers use a nonzero bit encoding, so `Option<layout::Descent>` stays
eight bytes rather than adding another tag word to every stored row. This is a
layout assertion, not a workload memory measurement.

`Sheet::defaults()` returns `Result<Option<&layout::Defaults>>` and never
invents font-dependent state when `sheetFormatPr` is absent. The read model
preserves optional base/default widths, required height, custom/hidden/thick
flags, producer outline summaries, and sheet-level descent. Row views and patch
states now also expose row descent. Microsoft's `[MS-XLSX]` rule is represented
semantically: a stored descent makes `custom_height()` true even when the core
`customHeight` spelling is absent or false. Physical namespace prefixes and
native IDs remain below the facade.

Existing and new-sheet transactions both expose the short
`defaults()` editor. It supports base width and default width set/reset,
required height, hide/show, thick/normal edges, descent set/reset, and explicit
whole-record removal. Creating a record without supplying its required height
returns `DefaultsEditBlock::NeedsHeight` before rewrite; protected sheets and
unmodeled compatibility ownership also produce typed blocks. Compact
`layout::Fields` bitflags make each default facet independently joinable, while
same-facet writes and whole-record deletion conflict deterministically. Row
descent remains independent from row height and the other row-property facets.
No public lock wrapper or last-writer-wins rule is introduced.

The parser performs a narrow, depth-bounded capture of only direct
`x14ac:dyDescent` attributes before common MCE preprocessing. It does not claim
to understand the complete extension namespace or bypass `MustUnderstand`.
Worksheet surgery preserves untouched bytes, unknown attributes, child
payloads, existing qualified names, and unedited default facets. When a new
descent requires namespace declarations, it chooses collision-free prefixes,
adds the exact x14ac and MCE namespace bindings, and extends `mc:Ignorable`.
Rewritten parts are reparsed before publication; complete before/after default
and row states make patches reversible to exact source part bytes. Unit and
semantic tests cover checked domains, arbitrary prefixes, extension-depth and
malformed inputs, sparse insertion, set/reset/remove, protected and
compatibility blocks, new sheets, source-exact inversion, and disjoint/same-
facet joins. The public `sheet_defaults` example applies height 24, width 14,
sheet descent 0.2, and row-2 height/descent 32/0.3 through the ordinary facade.

Computer Use exercised that example against an Excel-authored baseline in
Microsoft Excel for Mac 16.110.2 (build 26062818), Office LTSC Standard for Mac
2024. Excel opened the Litchi output without a repair or compatibility prompt.
Its Row Height dialog reported exactly 32 for row 2 and exactly 24 for an
implicit row; its Column Width dialog reported 13.17 for stored OOXML width 14,
the application's display-unit normalization. Excel added
`Excel default-layout resave marker` at E4 and saved as a separate XLSX without
warning. The resaved archive passed ZIP integrity validation, and the public
reader recovered the marker, default height 24, default width 14, sheet descent
0.2, and row-2 height 32. Excel normalized the row-specific descent from 0.3 to
the sheet value 0.2 and materialized some row heights during resave; that
producer behavior is recorded rather than presented as lexical stability. A
second artifact created from Litchi's minimal workbook forced collision-safe
x14ac/MCE namespace injection and also opened in Excel without repair. This
certifies those local open/dialog/edit/resave/reverse-read and namespace-open
paths on that build. It does not establish stable row-specific descent across
Excel resave, every default flag or producer normalization, other Office
applications or builds, or performance.

The twenty-second slice implements sparse merged-cell CRUD in
`litchi-xlsx`. The worksheet parser validates `mergeCells` counts, child
context, non-single ranges, and two-dimensional non-overlap, then stores only
checked `Rect` values in a compact static interval tree. Each range contributes
one 32-bit span index; lookup descends balanced row partitions and binary-
searches column-disjoint spans instead of scanning or expanding rows.
The compact index reserves one `u32` sentinel and therefore admits at most
4,294,967,294 ranges, matching Excel's documented `mergeCell` occurrence limit
in the local `[MS-OE376]` section 2.1.661 conformance notes.
`Sheet::merges()` borrows the ranges directly. `Sheet::cell()` now returns
`cell::View`: anchors retain their stored payload, followers return
`Covered(Rect)`, and missing coordinates stay distinct. A producer follower
record is not discarded, but the structural view wins so callers cannot
accidentally treat hidden payload as an ordinary cell. No range is expanded
into per-coordinate state.

Existing and transaction-local sheets expose short `merge(area)` and
`unmerge(at)` verbs. Commit projects intents sequentially, returns typed errors
for single-cell or overlapping ranges, and refuses a new range whose final
follower cells contain data. Explicit clears/removals in the same transaction
are honored. Serialization runs in three phases—remove old merge structure,
apply ordinary sparse edits, add new merge structure—so editing a just-unmerged
follower and clearing before merge are both safe. Protected sheets, grouped
formulas, markup-compatibility ownership, and unmodeled `mergeCells` children
remain typed blocks rather than guessed rewrites.

Low-level surgery retains untouched `mergeCell` bytes, qualified names,
unknown attributes, and container attributes, updates the advisory count,
removes an empty container, inserts a new container at its schema position, and
only expands a producer `dimension` for newly added ranges. It never shrinks a
dimension or normalizes it merely because an existing merge lies outside the
hint. Every result is reparsed before publication. Complete semantic merge
changes join the source-checked reversible patch; disjoint merge rectangles can
join without locks, while intersecting intents and independently written
follower content produce deterministic structural conflicts. Independent
follower clears/removals remain joinable and run before merge creation.

Tests cover sparse lookup without follower materialization, malformed counts
and contexts, overlap rejection, lossless add/remove/container deletion,
schema ordering, typed dependency blocks, unmerge-and-edit, clear-and-merge,
new sheets, exact inverse restoration, concurrent snapshot reads, and
disjoint/overlapping joins. The public `merged_cells` example performs create,
unmerge, follower update, and a second merge through the ordinary facade. These
are functional and preservation gates; they make no latency, allocation, cache,
or contention claim.

Computer Use exercised that example in Microsoft Excel for Mac 16.110.2
(build 26062818), Office LTSC Standard for Mac 2024. Excel opened the generated
workbook without a repair or compatibility prompt and reported used range
A1:F4. Navigating to covered B2 selected anchor A1 and retained `Merged title`;
navigating to covered C4 selected A4 and retained `Merged footer`; F2 remained
an independent unmerged cell with `Unmerged follower`. Excel wrote
`Excel merge resave marker` to F3 and saved a separate XLSX without warning.
The resaved archive passed ZIP integrity validation, and the public reader
recovered exactly `A1:C2` and `A4:C4`, the unmerged F2 content, and the F3
marker. Excel added its normal theme, shared-string, style, and document-
property parts and materialized empty physical records for merge followers;
`cell::View` still reports those coordinates as covered, which is the intended
structural precedence. This certifies one local create/unmerge/remerge/open/
edit/resave/reverse-read path on that build. It does not certify merge styling,
structural row/column shifts, grouped formulas, every producer normalization,
other Office applications or builds, or performance.

The twenty-third implementation slice starts the binary Office dependency
split by introducing `litchi-ole-common`. Shared Custom XML Data Storage and
MS-OSHARED smart-tag property-bag records now live below the DOC, PPT, and XLS
semantic models and are consumed directly by the legacy migration host. The
old host modules are removed rather than retained as public compatibility
re-exports. The extracted crate has a deliberately narrow internal dependency
ceiling of `litchi-cfb` and `litchi-core`; it contains no concrete binary format
model, async runtime, or peer-format dependency. Custom XML schema references
remain inert metadata, XML reads are resource bounded, smart-tag indexes and
code pages are validated, and malformed GUID input returns a typed error
instead of relying on an unreachable branch.

The dependency fence is now a complete checked-in policy rather than a partial
set of hard-coded exceptions. It inventories every direct workspace crate and
every normal, optional, development, renamed, and target-specific internal
edge. Canonical downward edges and ordered migration debt are distinct; every
debt item has a reason and exit condition. CI rejects unclassified edges,
cycles, concrete peer-format coupling, upward common-crate dependencies,
runtime dependencies in neutral crates, unaudited manifests, newly introduced
core format debt, and stale debt whose underlying edge or feature has already
been removed. This gate makes the remaining monolith dependencies visible and
removable one at a time; it does not make either monolith part of the target
architecture. The same workflow denies compiler, Clippy, and rustdoc warnings
across every workspace target with the complete feature set so test-, example-,
and documentation-only regressions cannot hide behind default library builds.

The twenty-fourth implementation slice extracts the host-neutral binary
OfficeArt substrate into `litchi-odraw`. The crate owns bounded record headers,
record kinds, zero-copy record and container traversal, topology validation,
ordered property tables, typed color and complex-array values, shape flags,
and validated record writing. It has no internal workspace dependency and does
not depend on DOC, PPT, XLS, CFB, an async runtime, or a concrete host model. Host
records remain typed at their boundaries: Word anchors, PowerPoint client data,
and Excel client anchors are interpreted by their respective crates rather
than hidden behind an untyped common payload. Borrowed parsing keeps payloads
in their source allocation, while writer builders accept borrowed or owned
payloads through `Cow` and validate lengths and container invariants before
emission.

The DOC, PPT, and XLS readers and writers now consume `litchi-odraw` directly.
The legacy public Escher model and duplicate PowerPoint parser were removed
instead of retained as compatibility aliases. The migration host keeps only
format-specific bridges and semantic models. OfficeArt properties preserve
wire order, exact property identifiers, and flag bits for lossless inspection;
duplicate semantic identifiers are rejected at the parse boundary instead of
making ergonomic lookup ambiguous. Native record kinds and flags remain
representable without weakening the typed known cases. This is the neutral
ODraw extraction and first host migration slice, not completion of the planned
`litchi-doc`, `litchi-ppt`, and `litchi-xls` crate split or the complete binary
CRUD checklist.

The twenty-fifth implementation slice extracts shared `[MS-OFFCRYPTO]`
structures and transformations into the runtime-neutral `litchi-crypto`
crate. Bounded DataSpaces, property-integrity, protected-content, sensitivity-
label, and RC4 CryptoAPI code no longer lives in the binary-format monolith.
The crate depends downward on `litchi-cfb` and `litchi-ole-common`, has no
concrete DOC/PPT/XLS or OOXML dependency, and neither activates protected
content nor anchors asynchronous work to a runtime. Secret-derived RC4 state
is private, zeroizing, and move-only; cipher operations borrow it, and block-key
derivation does not allocate. Contextual modules avoid global prefixes:
`spaces::{Graph, Map, Transform}`, `integrity::Info`,
`protected::{Envelope, Kind}`, and `labels::{List, Label, Content}`. `Content`
and the CryptoAPI flag set use compact typed bitflags; sensitivity-label reads
retain unknown bits while writes reject them. The public CryptoAPI surface uses
module-scoped typed names and short verbs: `rc4::{Flags, Header, Context, Error}` plus
`build_header`, `parse_header`, `context`, `verify`, `apply`, and `apply_at`.
Unsupported algorithms and invalid flags are rejected as typed errors; no raw
flag integer, public secret field, assertion-based parser branch, or runtime
lock wrapper is required for ordinary use.

DOC, PPT, and XLS encryption paths now consume this capability directly
through the legacy host, while encrypted OOXML package inspection depends on
`litchi-crypto` and `litchi-cfb` rather than reaching through `litchi-ole`.
The dependency fence removes that cross-family monolith edge instead of
retaining a compatibility re-export. This completes only the shared crypto
prerequisite. The read-only extraction audits found that signatures, VBA,
generic embedded-object handling, the remaining host image bridge, and the
PPT/XLS chart-workbook bridge must also move to focused lower layers before
the internally dense DOC, PPT, and XLS trees can move atomically. Those concrete
crate splits and broader encryption interoperability certification remain open.

Computer Use exercised permanent generated smoke artifacts in the installed
macOS desktop Microsoft Word, PowerPoint, and Excel applications. PowerPoint
opened the generated PPT without repair, rendered the title, rectangle,
ellipse, text boxes, and all four cells of a table at their requested bounds.
Selecting a table cell exposed PowerPoint's native Table Design and Table
Layout tabs. Excel opened the final XLS without repair, rendered a
primitive rectangle, a Unicode text box, and a grouped ellipse/text-box pair;
selecting the text box exposed native resize/edit affordances and Shape Format.
Word opened the final DOC without repair, exposed `Rectangle 1` and `Text Box
2` as native floating objects, rendered their corrected blue and pale-green
fills, retained the text-box story, and displayed native resize handles when
the rectangle was selected.

These native checks found two specification defects that unit tests alone had
missed. An XLS text shape originally placed `OfficeArtClientTextbox` in the
same `MsoDrawing` fragment as `OfficeArtClientData`; `[MS-XLS]` requires the
OBJ record between those fragments, followed by TXO and CONTINUE text records.
The writer now splits the final client-textbox record at that BIFF boundary,
and the reader rejoins it only when the preceding top-level OfficeArt record is
structurally incomplete. The DOC writer also encoded `OfficeArtFDGG.cidcl` as
the number of IDCL entries. `[MS-ODRAW]` section 2.2.47 requires that count plus
one; Word rejected every floating drawing until the invariant was corrected.
Regression tests pin both exact wire orders and the cluster count. These checks
certify native open/render/select behavior for the generated artifacts on this
machine. They do not yet certify Office resave/reverse-read for these binary
artifacts, every Office build, the full binary CRUD matrix, or performance.

The complete regression sweep exposed four further interoperability defects.
The DOC writer had encoded direct OfficeArt RGB colors in blue-green-red byte
order; typed `ColorRef` parsing made the low-byte red, then green, then blue
wire contract explicit and the writer and integration expectations now agree.
A genuine Microsoft Word fixture also contains a non-visible direct background
sentinel with `fHaveAnchor` set but no `OfficeArtClientAnchor`. The parser keeps
strict flag/topology agreement for user shapes, accepts only that narrow
background omission, validates the rest of the sentinel, and excludes it from
the user-facing shape list. Regression tests pin the asymmetric color bytes,
the compatibility exception, and rejection of the same missing anchor on an
ordinary shape.

The PPT table writer had also placed an `OfficeArtChildAnchor` on a slide-level
table group whose shape was not marked as a child. It now uses the
PowerPoint-defined eight-byte `OfficeArtClientAnchor`; table/chart coexistence
therefore survives typed shape parsing and chart-frame attribution. A native
PowerPoint visual probe rejected the initial full-rectangle writer choice by
rendering transposed and compressed coordinates, so the writer now pins the
producer-compatible `SmallRectStruct` and rejects coordinates outside its
range. Table dimension setters and aggregate dimensions likewise return typed
errors when conversion or addition exceeds the representable EMU range instead
of saturating or overflowing. The permanent native smoke artifact includes a
table so this topology remains part of the desktop PowerPoint gate.

The PPT master writer also had a private numeric placeholder mini-enum whose
`TITLE = 0` encoded `PT_None`, while `BODY = 1` encoded `MasterTitle`. It now
takes the shared typed `PowerPointPlaceholderKind` directly, so neither call
site can silently manufacture those invalid semantic values. The
header/footer integration regression pins the corrected master-title and
master-body placeholder records. The regenerated final artifact was reopened
in desktop PowerPoint after this correction and rendered without a repair
dialog.

The twenty-sixth implementation slice moves generic legacy embedded-object
discovery and CFB rewrite ownership from `litchi-ole` into
`litchi-ole-common::object`. The migration host no longer declares the old
module or re-exports a compatibility facade. DOC embedded-object and
tracked-revision editors plus XLS object and chart editors consume the common
crate directly. Public names are contextual and short (`Format`, `Kind`,
`Object`, `Objects`, `Editor`, and `Limits`); `Objects::get` performs semantic
identifier lookup and `Objects::at` offers a checked raw ordinal without using
Rust's panic-defined indexing operator.

The common object model keeps format-owned metadata as opaque immutable bytes.
DOC alone exposes the typed `doc::embedded_object::Info` interpretation of its
`ObjInfo` flags, so the common crate does not own a concrete Word semantic
type. Large compound objects, native payloads, previews, and captured streams
use shared immutable slices. Atomic collection edits clone only the object
being edited, and editor snapshots share large buffers instead of copying the
whole container. `Editor::open` consumes the caller's `Vec<u8>`; consuming a
clean, uniquely owned editor returns that input allocation and exact byte
sequence instead of rerendering it. Changed output still uses a validated full
CFB render. This is
a structural reduction in avoidable copies, not a throughput or allocation
claim; those require the ADR 0005 measurement gates.

Checked parsing replaced the remaining assertion-based conversions in the
touched common, DOC object/revision, XLS object, and production XLS chart
paths. Offset arithmetic, array conversion, axis/series lookup, object records,
and empty-range invariants now return format errors rather than calling
`unwrap`, `expect`, or indexing an assumed element. Focused tests cover inert
DOC/XLS discovery, atomic mutation, byte-exact clean finish, DOC embedded
objects and revisions, XLS OLE objects and controls, and PPT/XLS chart
integration. This dependency/API move does not change an authored Office
artifact scenario, so it does not add a new desktop Office claim; the earlier
Word, PowerPoint, and Excel Computer Use evidence remains the applicable native
baseline.

Parallel read-only closure audits establish the next compilation-safe order.
The five shared OVBA modules move atomically from `litchi-cfb` into
`litchi-vba`, after which OLE and OOXML consumers migrate directly and both old
re-export facades are deleted. The minimum acyclic signature cut is
`litchi-sign::cfb` over CFB plus the current OPC detached-signature engine; the
final state must then invert the neutral XMLDSig engine out of OPC before OPC
can depend on signing, otherwise the two crates would cycle. OfficeArt BLIP and
FBSE grammar moves from image conversion into `litchi-odraw`, leaving codecs
and rendering in `litchi-imgconv`. The PPT/XLS chart peer edge is removed by
extracting the `[MS-OGRAPH]` model and compound chart codec into
`litchi-ograph`; XLS retains workbook mutation and PPT retains frame/object
integration. With signature, VBA, BLIP/image, and OGraph prerequisites below
the hosts, the internally dense DOC, PPT, and XLS trees can move atomically into
their concrete crates without compatibility monoliths.

The twenty-seventh implementation slice performs the first of those cuts.
All shared MS-OVBA compression, directory, project, and authoring code moves
from `litchi-cfb` into the runtime-neutral `litchi-vba` crate. CFB is again only
a compound-container layer; VBA depends downward on it and `litchi-core`, and
owns no DOC, PPT, XLS, OPC, OOXML, executor, or async-runtime concern. Its public
vocabulary is contextual rather than prefix-heavy: `codec`, `dir`, `project`,
and `build` provide short `Dir`, `Module`, `Kind`, `Project`, `Text`, `Id`, and
`Platform` names beneath their defining modules.

`Payload` is the checked ownership boundary for a standalone project. Reading
one consumes the caller's CFB bytes, validates the bounded project topology,
and retains those bytes for move-based package attachment; arbitrary bytes
cannot masquerade as a project through an infallible constructor. A detached
`build::Project` is consumed by `finish`, so successful authoring yields the
same validated capability without cloning a potentially large source tree.
The emitted hierarchy includes the required root `PROJECT`, `VBA/dir`,
`VBA/_VBA_PROJECT`, and one stream per declared module. The seven-byte
`_VBA_PROJECT` header uses the specified reserved marker and cache-free write
version; the optional `PROJECTwm` name map remains supported. Lower-level
borrowing APIs remain available for legacy CFB hosts that must copy project
streams into a larger compound file.

Standalone authoring rejects a CFB ceiling smaller than the 512-byte compound
header before encoding any project streams. Its final `Write + Seek` sink also
checks every attempted growth against `max_cfb_bytes`, including sparse seeks,
and reports a typed limit error without growing the output vector past the
ceiling. Encoded intermediate streams are released before that final output
allocation. This bounds the returned standalone buffer; it does not yet remove
the internal stream copies and staging performed by `litchi-cfb::OleWriter`.
That writer remains a measured follow-up rather than a zero-copy claim.

DOC, PPT, XLS, DOCX, PPTX, XLSB, and XLSX integrations now depend directly on
`litchi-vba`. The old `litchi-cfb::ovba`, `litchi-ole::ovba`, and
`litchi-ooxml::vba` facades are deleted rather than retained as aliases.
High-level mutation accepts typed builders or validated payloads through short,
consuming verbs, validates before changing host state, and exposes checked read
and clear operations without requiring raw project identifiers or lock-wrapper
types. New and changed production parser/writer branches return typed errors
rather than relying on `unwrap`, `expect`, or assertion-defined input
invariants.

OOXML project replacement and removal are transactions over a structurally
cloned OPC graph whose large immutable part bodies remain `Arc`-shared. Before
mutation, the implementation rejects canonical target names that already have
dangling inbound relationships and validates that every existing project,
supplemental-data, or signature part has only its declared owner. Any later
part-name, content-type, relationship, or signature-cleanup error drops the
staged graph and leaves the source package unchanged. A no-project clear
returns before taking that snapshot. Regression tests cover dangling canonical
targets, post-removal name conflicts, and a deliberately late content-type
failure.

The PowerPoint binary read path resolves `VbaProjectStg` once as a borrowed
strict record view. It checks the persisted record length, caller stored-byte
ceiling, and declared decompressed length before allocating; an uncompressed
CFB remains borrowed through `OleFile::open`, while zlib output is bounded and
must exactly match its declared length with no trailing input. This removes the
previous repeated payload copies from the VBA path without claiming a measured
latency or allocation result.

This ownership and API extraction deliberately preserves the existing wire
semantics and generated macro graphs. It therefore does not create a new native
Microsoft Office artifact claim: the earlier desktop Word, PowerPoint, and
Excel smoke evidence remains the applicable baseline. Native Office open,
edit/resave, and reverse-read must be rerun when a later slice changes emitted
project bytes or package relationships.

The slice passes warning-denied workspace check, Clippy, and rustdoc gates with
all targets and features enabled; the complete all-target workspace test run;
and workspace doctests other than the `cdylib`-only Python package. Focused
coverage includes 24 `litchi-vba` tests, the 2,504-test OLE library suite plus
its DOC/PPT/XLS VBA integrations, and the 2,068-test OOXML library suite plus
PPTX integration coverage. The dependency checker accepts 27 workspace crates
and 65 internal dependency declarations, including the new downward VBA edges.

The twenty-eighth implementation slice completes the acyclic shared-signature
cut. `litchi-sign` owns bounded, trust-neutral XMLDSig authoring and verification
plus the CFB storage adapter; `litchi-opc` depends on that engine and owns only
package-part selection, RelationshipTransform resolution, certificate-part
relationships, and staged graph mutation. The former duplicate OPC and OLE
signature engines and their long compatibility facades are deleted. DOC, PPT,
XLS, DOCX, PPTX, XLSB, and XLSX use the canonical `Signer`, `Policy`, `Limits`,
`Coverage`, `Report`, `Status`, and `Trust` vocabulary directly or through a
short host method. The safe ordinary `signatures`, `sign`, and additive-editor
paths use strict policy; permissive SHA-1 and partial-package handling requires
an explicit compatibility policy and returns `Coverage::Partial` rather than an
unqualified complete result.

The XML engine requires one uniquely bound package object whose sole direct
Manifest owns the external package references. Producer-specific Office
metadata and XAdES signed-properties references are accepted only as known,
uniquely resolved SignedInfo objects whose canonicalized digests also verify;
unknown, duplicate, ambiguous, or unverified fragment references fail. OPC
resolution borrows ordinary part bodies, transforms only relationship parts,
and merges related certificate DER as borrowed evidence. Signature addition and
replacement author and validate a structurally shared staged package before
commit. CFB editing consumes the source allocation, materializes streams lazily,
authors before clearing on replacement, and returns the exact input allocation
when clean. A changed CFB preserves sector size, root and exposed storage
CLSIDs, stream bytes, and hierarchy; state bits and directory timestamps remain
an explicit `litchi-cfb` metadata limitation rather than a false preservation
claim.

The twenty-ninth implementation slice moves OfficeArt image grammar and writing
from optional image conversion into `litchi-odraw::image`. Borrowed `Blip`,
`Bitmap`, `Meta`, `Entry`, `Store`, `Block`, and delayed-storage views retain
unknown records, JPEG record flavor, platform fields, direct BLIPs, dual UIDs,
and offset-zero semantics under explicit bounds. Consuming or borrowing writer
builders emit the required MD4 UID, cache it instead of hashing a large image
twice, use the `0xFE` uncompressed metafile marker, and stream payload bytes to
the caller's sink. `litchi-imgconv` now owns only bounded decode, decompression,
render, and codec conversion over those canonical views; its duplicate BLIP and
FBSE models are deleted.

The PowerPoint host follows the actual binary topology: the Pictures stream is
a headerless BStoreDelay sequence, while the FBSE table comes from the drawing
group. Picture lookup resolves `pib` to an FBSE and then `foDelay` to a BLIP,
including direct embedded BLIPs and the valid zero offset, without constructing
a self-referential package model. DOC and XLS retain their host-specific anchor,
object, and stream rules while consuming the same borrowed image grammar. This
is a parser/writer ownership correction, not a claim that every image codec or
every binary picture CRUD path is complete.

The thirtieth implementation slice establishes `litchi-ograph` as the neutral
chart foundation required to remove PPT/XLS peer-format coupling. Its raw layer
iterates borrowed BIFF frames under record and output budgets, its contextual
record modules expose short names such as `chart3d::BarShape`, `frame::Frame`,
`line::Line`, `pie::Format`, and `series::Parent`, and its owned package boundary
consumes validated standalone compound bytes without cloning them. Reserved
bits, unknown records, record order, and opaque streams remain lossless. XLS
workbook/tab/OBJ mutation and PPT frame/embedded-object integration deliberately
remain in their hosts; migrating those bridges and completing the full OGraph
grammar are later slices, so this foundation alone makes no chart authoring,
rendering, activation, or native-Office compatibility claim.

These three slices preserve the previously verified native Office artifacts and
do not add a desktop-application certification claim. Their verification is
therefore focused on the changed dependency boundaries, hostile parser cases,
real producer signature fixtures, exact round trips, transactional failures,
warnings, and documentation. Per the explicit review decision, the earlier
fully green workspace baseline is relied upon and the full workspace gate is
not repeated for this cut; any later change to emitted document relationships
or visible image/chart bytes still requires the applicable Microsoft Office
open, edit/resave, and reverse-read evidence.

Focused evidence for this cut is green. Warning-denied test and Clippy gates
cover 71 `litchi-cfb`, 17 `litchi-sign`, 50 `litchi-odraw`, 48
`litchi-imgconv`, and 18 `litchi-ograph` tests. `litchi-opc` passes 66 unit, 11
integration, and 5 documentation tests, including seven real Microsoft/POI
signed-package fixtures. The all-feature OLE gate passes 2,501 library tests
with zero failures and three pre-existing ignored cases plus every integration
and example target; the narrowed image and final hostile-extractor suites pass
28 and 17 tests respectively. OOXML's focused gates pass two real signed
DOCX/XLSX/PPTX fixture tests, 36 VBA unit tests, and two PPTX VBA integration
tests, with all-target check and Clippy warnings denied.

The umbrella's default, full-feature, and `ole,imgconv` library boundaries pass
warning-denied check and Clippy; its image facade documentation compiles and the
newly boxed format-neutral table/row facade passes fourteen full-feature focused
tests. The dependency checker accepts 29 workspace packages, 74 direct internal
edges, and 32 explicit
migration-debt entries; all seven checker regression tests pass. Formatting,
rustdoc on the extracted foundations, legacy-name searches, primary-path panic
searches, and diff validation are clean. These are focused crate and boundary
gates, not a repeated full-workspace run or a new native-Office certification.

The thirty-first implementation slice moves the binary chart host bridge onto
that neutral foundation. `litchi-ograph::chart` now distinguishes the strictly
validated package and BIFF BOF/EOF framing of a standalone Graph object from
allocation-free chart discovery inside an arbitrary Excel `Workbook` stream;
it does not claim the unavailable full Graph chart-sheet grammar. Borrowed
`Ref`, move-owned `Stream`, and move-owned `Book` preserve exact bytes and
allocations at the appropriate boundary. The semantic `Chart` uses
producer-specific, type-checked Graph and Excel link and cache grammars,
bounded identifiers and numeric settings, and short `axis`, `format`, and
`group` modules. A pristine parsed chart replays its original allocation
exactly; edits to a parsed chart are refused until opaque record placement can
be proved. Fresh semantic authoring is also refused until the complete
mandatory chart-format, series, axis-parent, and cache scaffolds required by
`[MS-XLS]` section 2.1.7.20.1 and the corresponding `[MS-OGRAPH]` chart-sheet
grammar are modeled; the safe facade does not expose a self-consistent but
Office-nonconforming abbreviated stream. Chart-group line and up/down-bar
values nevertheless own their mandatory format records, so the type system
cannot construct those abbreviated invalid collections when authoring resumes.

XLS now exposes the concise `xls::chart::{Chart, Editor, Selector, Location}`
facade, with semantic sheet-name and embedded-chart selectors as the primary
entry points and checked raw positions as an explicit secondary capability.
Seven duplicate record modules and their long compatibility names are deleted;
BIFF framing, discovery, bounded encoding, and shared chart records come from
`litchi-ograph`. Unsupported records remain inspectable, but an edit that would
guess their placement returns a typed `UnsafeEdit`; unchanged chart substreams
are copied exactly. PowerPoint chart discovery likewise returns one neutral,
bounded inventory across standalone Graph and embedded Excel chart objects,
supports semantic frame and slide selectors, and reports per-object failures
for payload decode and validation without discarding successful siblings.

This slice deliberately leaves the PowerPoint writer's temporary XLS chart
authoring peer seam and the remaining duplicate XLS semantic projection visible
as migration debt. Every public XLS add, insert, chart-sheet creation,
replacement, and standalone chart-workbook path, plus `PptWriter::add_chart`,
returns the typed `litchi_ograph::Error::UnsupportedAuthoring` without mutation
instead of emitting that incomplete legacy grammar. Exact replay and inventory
plus chart-sheet removal and whole-sheet reordering remain enabled. Structural
mutation of embedded charts is refused: the abbreviated fixture seam does not
model the enclosing OfficeArt drawing graph and is not an Office-conformance
basis for removal or z-order changes. The abbreviated generator survives only
under `cfg(test)` as parser-fixture scaffolding. Transaction snapshots that
clone whole XLS workbooks also remain a measured redesign target. Read-side PPT
code no longer depends on XLS, but restoring writer support requires a neutral
host-authoring contract, the complete mandatory scaffold, and new native Office
evidence rather than silently changing embedded-object topology. Because this
slice adds no new emitted chart artifact, Computer Use is not repeated and no
new native Microsoft Office claim is made. Per the explicit review decision,
verification for this slice uses only
the focused OGraph, OLE chart integration, dependency-boundary, warning,
Clippy, rustdoc, formatting, and diff gates; the earlier fully green workspace
baseline is relied upon instead of rerunning the full workspace gate.

Focused evidence is green. OGraph passes 31 warning-denied tests, all-target
check and Clippy, and warning-denied rustdoc. OLE passes 33 XLS chart-focused
unit tests, four PPT chart-focused unit tests, and nine XLS/PPT integration
tests, including typed atomic refusal, exact clean replay, real POI and
LibreOffice fixture gates, both chart host topologies, and per-object failure
isolation. Its all-feature/all-target warning-denied check and Clippy and its
all-feature warning-denied rustdoc are clean. The dependency checker accepts 29
workspace packages, 75 direct internal edges, and 33 explicit migration-debt
entries; all seven checker regression tests pass. Formatting, manifest order,
source-boundary searches, and diff validation are also clean.

The thirty-second implementation slice resolves the next chart dependency and
mutation decisions against the normative local specifications. `[MS-XLS]`
section 2.1.7.20.1 makes the minimal embedded-chart sequence explicit:
`BOF`, the mandatory page-setup records, `PrintSize`, `Units`, then `Chart /
Begin / Scl / PlotGrowth / ShtProps / AxesUsed / AxisParent / Begin / Pos /
ChartFormat / Begin / family / CrtLink / End / End / End`, followed by
`Dimensions`, the three ordered `SIIndex(1..=3)` sections, and `EOF`. Every
regular series owns exactly four `AI = BRAI [SeriesText]` bindings in
name/value/category/bubble order. Embedded charts forbid `WINDOW` and
`CUSTOMVIEW`, whereas a chart-sheet tab requires a window; a category axis
also requires `AxcExt`. These are now recorded as producer and placement
requirements rather than inferred defaults.

The neutral model advances without weakening its refusal boundary. Checked
zoom, fixed-point plot growth, axis-parent position, Excel-mandatory group
`CrtLink`, four-slot AI bindings, producer-owned dimensions, and the three
Excel cache sections are represented by short `layout`, `axis`, and `cache`
types. Only locally proven Excel ownership and order constraints are enforced;
they are not silently projected onto Graph. The encoder remains unavailable
through the public facade: page setup, host placement, complete axis ownership,
attached-label and frame collections are not yet fully represented. In
addition, the checked-in and official rendered `[MS-OGRAPH]` documents
reference but do not contain the normative `GraphWorkBookGrammar.abnf` or
`GraphChartSheetGrammar.abnf`; standalone Graph topology and authoring therefore
cannot be certified from the available authority and stay conservatively
bounded and typed-refused.

The type split closes several invalid-state paths found during self-review.
`Cache::{Excel, Graph}` owns producer-specific coordinates and XF/IFmt values,
and Excel alone owns typed Boolean/error cache values. Cache CRUD derives or
validates its `Dimensions`. Excel cached `Label.st` is parsed and encoded as
the required two-byte-count `XLUnicodeString`, distinct from Graph and
`SeriesText` short strings. `Owner::{Group, Trend, ErrorBar}` represents both
legal series-owner branches; auxiliary parent references and regular group
references are dependency-checked during removal. `GroupId` now means the
zero-based `SerToCrt` target, while `Order` means `ChartFormat.icrt` drawing
order. Axis and group values retain a checked `ParentId`. Mutable slice access
uses a short edit guard that marks parsed input dirty only on actual mutable
dereference, so inspection does not accidentally disable exact replay.

The PowerPoint writer no longer imports or fabricates an XLS chart semantic
model under `cfg(test)`. Its concise request validation remains, and public
authoring still fails atomically before presentation mutation. Self-review of
the initial XLS mutation plan caught an unsafe ownership assumption before
commit. `[MS-XLS]` defines worksheet objects as `MSODRAWING` followed by
`TEXTOBJECT` or `OBJ`, with `MSODRAWING = MsoDrawing *Continue` and
`OBJ = Obj *Continue *CHART`; an OfficeArt `ClientData` also requires the next
record to be its `Obj`. Moving or deleting only a bare `Obj` plus chart
substream would orphan anchors, continuation state, and drawing-group
bookkeeping. That rewrite has therefore been removed from the public path.
Embedded removal and non-identity reordering now return the distinct typed
`UnsupportedMutation` error without modifying the package. Identity reorder is
a byte-exact no-op; chart-sheet removal and whole-sheet reorder retain their
separate, exact substream path.

The low-level compound editor also exposes a shared immutable stream
capability, allowing the package editor and typed XLS inventory to retain the
same `Arc<[u8]>` Workbook allocation across open and validated replacement.
This is structural copy elimination, not a throughput claim; rendering a
changed CFB package still allocates and must be benchmarked under ADR 0005.

This slice emits no fresh chart artifact and enables no new embedded mutation:
fresh authoring and structural embedded edits are typed-refused. Existing
chart-sheet tab removal and whole-sheet reorder retain their prior evidence and
are not newly certified here. No Computer Use certification is therefore
claimed for this slice. Per the explicit user direction, verification is
limited to the affected crates and chart paths; the previous fully green
workspace gate is not rerun.

Focused evidence for this slice is green with warnings denied. `litchi-ograph`
passes 44 tests plus all-target check, Clippy, and rustdoc. `litchi-ole-common`
passes nine library and seven object-editor tests plus all-target check, Clippy,
and rustdoc. `litchi-ole` passes 34 XLS chart units, four PPT chart units, and
ten XLS/PPT chart integrations; the bundled `WithThreeCharts.xls` regression
proves byte-exact atomic refusal for embedded removal and non-identity reorder
and byte-exact success for identity reorder. Its all-feature/all-target check,
Clippy, and rustdoc are clean. Formatting and manifest order are clean. The
dependency checker accepts 29 workspace packages, 75 direct internal edges,
and 33 explicit migration-debt entries, and all seven checker regressions pass.
These are focused gates, not a repeated full-workspace or native-Office run.

The thirty-third implementation slice extracts legacy text conversion from the
format-neutral core into `litchi-codepage`. The foundation has no runtime,
container, or peer-format dependency and forbids unsafe code. Its exhaustive
private discriminant makes `Page`, `Mbcs`, and `Ansi` one-byte capabilities:
`Page` represents every exactly supported page, `Mbcs` excludes UTF-16 from
byte-terminated paths, and `Ansi` admits only the ANSI pages enumerated by the
local `[MS-OSHARED]` glossary. Numeric construction is checked; strict decode
and encode are the defaults; lossy decoding is named explicitly; and the
foundation never guesses record terminators. Approximate substitutions are
deleted, including CP437/CP850 as IBM866, locale identifiers as code pages,
ISO-8859-1 as Windows-1252, Macintosh variants as unrelated codecs, and UTF-7
as UTF-8. RTF retains its exact local CP437 and CP850 tables.

Self-review and an independent audit turned the capability split into format
invariants rather than convention. The first UTF-16 round-trip regression
caught that `encoding_rs` intentionally uses UTF-8 as the output encoding for
its UTF-16 decoder labels; `Page` now writes UTF-16LE/BE explicitly and rejects
partial input units. XLS carries either `Mbcs` or UTF-16LE, so a public variant
cannot route UTF-16BE through byte-NUL scanning. Shared smart-tag stores carry
`Ansi` directly, reject embedded NULs and unsupported pages, and expose typed
PowerPoint and Word override entry points. Word's LCID inference is
conservative and fallible instead of treating every unknown language as
Windows-1252. RTF rejects unsupported `ansicpg` declarations and its writer
stores a typed `Charset`. Font-table and embedded-font `cpg` values carry
`FontPage`, whose variants admit only exact `Mbcs`, CP437, or CP850 codecs;
UTF-16 and unsupported numeric identifiers therefore cannot reach the writer.
RTF `fcharset` is a separate one-byte `FontCharset` enum, with an absent
declaration preserved separately from ANSI charset zero. In accordance with
`[MSFT-RTF]`, explicit `cpg` supersedes `fcharset`, the parser flushes text at
font switches and `plain`, and each byte run is decoded with the currently
selected font's exact page. Font-table primary, `falt`, and `fname` text is
deferred until that same precedence is known, while a nested `fontfile` name
uses its own local `cpg`. Only an absent/default charset inherits the header;
an explicit unavailable charset accepts invariant ASCII transport and Unicode
escapes but rejects non-ASCII transport instead of guessing. Valid but
unavailable Mac Japanese, Mac Russian, symbol, and Johab codecs likewise
produce errors when selected for body text instead of being substituted with
Shift-JIS, KOI8-R, or another superficially related codec.
VBA builders, parsed directories, and projects carry `Mbcs` from validation
onward. Raw numeric entry points remain explicit, fallible secondary
conveniences.

CFB Property Set PID 1 is modeled as `CodePage::{Utf16Le, Mbcs}`. The backing
maps and order vectors are private, generic property CRUD reserves PID 1, and
`set_page`, `set_page_id`, and `clear_page` update the typed state, property,
and order atomically. Serialization validates both sides of that invariant and
returns an error for a missing ordered property instead of relying on an
`expect`. This supplies create/update/delete coverage for the shared page while
keeping low-level raw identifiers out of the primary mutation path.

Generic hexadecimal decoding is now `litchi-core::hex::decode`; OOXML font GUID
handling uses that focused function and returns a typed length error instead of
an unchecked conversion. `litchi-core` no longer depends on `encoding_rs` or
declares binary-Office and RTF feature flags. The now-unused `litchi-core` edges
from `litchi-ole-common` and `litchi-vba` are also removed. The boundary checker
therefore accepts 30 workspace packages and 78 direct internal dependencies
with 31 explicit debt entries: three core debt entries close, while the honest
temporary `litchi-ole -> litchi-codepage` host edge adds one. All seven checker
regressions pass.

Supported artifact encoding and package topology are unchanged; this slice
turns previously guessed, malformed, or unrepresentable inputs into typed
errors. It therefore adds no new native-Office artifact claim, and the prior
fully green workspace and Microsoft Office baselines remain the applicable
evidence. Per the explicit review direction, verification is limited to the
affected package test matrices, strictness regressions, warning-denied checks,
Clippy, rustdoc, formatting, manifest order, dependency boundaries, and diff
validation rather than repeating the full workspace or Computer Use gates. The
final central-parser correction is covered by the complete warning-denied
`litchi-rtf --all-targets` matrix, including real LibreOffice font metadata and
the RTF corpus.

The thirty-fourth implementation slice completes the atomic XLS ownership
extraction. The complete BIFF source tree, its integration tests, and its
examples move from the temporary `litchi-ole` migration host into the concrete
`litchi-xls` crate. A source-boundary audit found no production XLS reference
to the `litchi-ole` facade or to DOC/PPT implementation modules. The only two
root-host references were test-fixture uses of the CFB reader and writer; they
now name `litchi-cfb` directly. The reverse audit found no DOC or PPT source
that imports XLS. Moving the whole format tree is therefore the smallest
coherent ownership cut: splitting individual BIFF facilities first would
create duplicate format owners or an upward compatibility tunnel through the
migration host.

This is an intentionally breaking topology change. `litchi-ole` deletes its
`xls` module and XLS root re-exports and neither depends on nor re-exports
`litchi-xls`. The canonical direct entry is `litchi_xls`; the concise umbrella
entry is `litchi::xls`. There is deliberately no `litchi::ole::xls` or
`litchi_ole::XlsWorkbook` compatibility path. `litchi-ole` now owns only the
remaining DOC/PPT migration host while those formats await their concrete
crate cuts.

The canonical internal dependency ceiling for `litchi-xls` is `litchi-cfb`,
`litchi-codepage`, `litchi-core`, `litchi-crypto`, `litchi-odraw`,
`litchi-ograph`, `litchi-ole-common`, `litchi-sign`, and `litchi-vba`. The crate
has no feature-selected peer format or runtime dependency, and it has no edge
back to `litchi-ole`. Legacy text conversion remains a direct
`litchi-codepage` capability; `litchi-core` does not regain a binary-Office
feature switch. This keeps the format crate above focused, reusable
foundations and prevents DOC/PPT host state from leaking into BIFF APIs.

A production panic-path audit found and closed two malformed-input holes while
the owner was isolated. Shared-string parsing now walks any chain of empty
`CONTINUE` records and returns `UnexpectedEndOfStream` instead of indexing an
empty segment. Pivot rewrites now validate every `BoundSheet` target as a
unique in-range BIFF record boundary, require BOF/EOF-bounded substreams, and
use checked source and destination slices before rewriting offsets. Regression
tests exercise chained empty continuations plus out-of-range, duplicate, and
mid-record sheet offsets.

The ownership move and targeted safety hardening do not change valid BIFF
writer output, package topology, or emitted wire semantics. They therefore
create no new native Microsoft Office claim. Per the explicit review direction,
the prior fully green workspace and Office baselines remain applicable; this
slice uses focused package, dependency-boundary, warning, formatting, and diff
gates instead of repeating the full workspace or Computer Use verification.
The warning-denied `litchi-xls --all-targets` matrix passes 776 unit tests and
196 tests across 58 integration targets, and all 15 examples build. Clippy and
rustdoc pass with warnings denied. The boundary checker accepts 31 workspace
packages and 89 direct internal dependencies with 31 explicit migration-debt
entries; all seven checker regressions and edited-file diff validation pass.

The thirty-fifth implementation slice completes the atomic PPT ownership
extraction. The complete legacy PowerPoint source tree, integration tests, and
examples move from the temporary `litchi-ole` migration host into the concrete
`litchi-ppt` crate. Its reader, writer, shape model, persist graph, presentation
metadata, OfficeArt integration, embedded OGraph charts, security adapters, and
inert VBA support therefore have one format owner. The residual `litchi-ole`
crate is now a DOC-only migration host; its PPT-only `litchi-ograph` and
`litchi-opc` edges and its unused image-codec edge close rather than becoming
permanent host tunnels.

This is an intentionally breaking ownership and feature-gating change. The
canonical direct entry is `litchi_ppt`, and the concise umbrella entry is the
independent `ppt` feature and `litchi::ppt` module. The umbrella default and
`full` sets enable `ppt`, while `ole` gates only the remaining DOC host. There
is deliberately no `litchi_ole::ppt` or `litchi::ole::ppt` compatibility alias.
Presentation detection and high-level PPT opening follow `ppt`, so disabling
DOC does not disable PPT and enabling PPT does not pull in the DOC host. The
optional umbrella `formula` feature forwards to `litchi-ppt?/formula` without
coupling either legacy format to an async runtime.

The canonical internal dependency ceiling for `litchi-ppt` is `litchi-cfb`,
`litchi-codepage`, `litchi-core`, `litchi-crypto`, optional `litchi-formula`,
`litchi-odraw`, `litchi-ograph`, `litchi-ole-common`, `litchi-opc`,
`litchi-sign`, and `litchi-vba`. The crate has no dependency on a peer concrete
format, an async runtime, or the former `litchi-ole` host. This keeps the PPT
implementation above focused storage, drawing, graph, package, code-page, and
security capabilities while preventing DOC state from leaking into its API.

The extraction also closes the shared-image ownership seam. Host-neutral,
move-first OfficeArt discovery now returns `litchi_odraw::image::File` values
that either borrow validated input bytes or consume themselves into owned
storage. Optional codec verbs are the separate `litchi_imgconv::Convert`
extension trait. Thus `litchi-odraw` stays independent of codecs, format crates
stay independent of `litchi-imgconv`, and callers opt into decoding without an
OfficeArt-to-codec dependency cycle. The former host-private Escher facade is
removed; the few PPT writer constants and record constructors that remain
format-specific live in the crate-private `officeart_wire` module on top of
`litchi-odraw` rather than reintroducing a second public OfficeArt model.

The new ownership seams are strict at malformed-input boundaries. PPT record
headers and payload extents use checked offset arithmetic and exact slice
validation, including regressions for `usize::MAX`, near-overflow offsets, and
maximal declared payloads. DOC picture lengths, name extents, and OfficeArt
record extents are checked before advancing. The DOC-only PLCF helper moves
under DOC ownership as a borrowed view, removing its property-buffer copy, and
OfficeArt owned image files cache a checked data range rather than reparsing on
every native-byte access. Producer filenames are reduced to bounded portable
basenames before any facade suggests a filesystem path.

This ownership move and image seam do not intentionally change supported PPT
wire output or package topology, so they create no new native Microsoft Office
artifact claim. Focused warning-denied gates pass for all features and targets
of `litchi-ppt`, the residual `litchi-ole`, and the umbrella crate. The complete
PPT package test surface passes (including its moved integration targets and
examples); the DOC-only host reports 947 passing tests and two ignored tests.
The isolated umbrella combinations `ppt`, `ppt,imgconv`, `ole,imgconv`,
`ppt,ooxml`, and `ooxml_encryption` compile independently, the `ppt` facade's
unit tests pass, and `ooxml_encryption` no longer activates DOC. Clippy passes
with warnings denied across all affected crates and targets; rustdoc passes for
`litchi-ppt` and the isolated `ppt,imgconv` umbrella facade. Formatting,
manifest order, diff validation, and the boundary checker are green. The
boundary checker accepts 32 workspace packages and 98 direct internal
dependency declarations with 28 explicit migration-debt entries, and all seven
checker regressions pass. Per the explicit review direction, the full workspace
gate and Computer Use/native Microsoft Office reruns are skipped: the
previously green workspace and native Office baselines remain the applicable
evidence for unchanged wire semantics.

The thirty-sixth implementation slice completes the atomic DOC ownership
extraction and retires the final legacy-binary migration host. The complete Word
binary reader, writer, record model, integration tests, examples, and fuzz
target move from `litchi-ole` into `litchi-doc`. DOC package, encryption,
OfficeArt, embedded-object, signature, equation, and inert VBA integration now
have one concrete owner. With XLS and PPT already extracted, the empty
`litchi-ole` monolith is deleted instead of being preserved as a dependency
tunnel.

This is intentionally breaking at both the package and feature boundaries. The
canonical direct entry is `litchi_doc`, and the concise umbrella entry is the
independent `doc` feature and `litchi::doc` module. The umbrella default and
`full` sets enable `doc`; DOC detection and the high-level document facade
follow that feature directly. There is deliberately no `litchi_ole::doc`,
`litchi::ole::doc`, `litchi::ole`, or `ole` feature compatibility alias. DOC,
PPT, and XLS may therefore be enabled independently without compiling a peer
format or a compatibility monolith.

The canonical internal dependency ceiling for `litchi-doc` is `litchi-cfb`,
`litchi-codepage`, `litchi-core`, `litchi-crypto`, optional `litchi-formula`,
`litchi-odraw`, `litchi-ole-common`, `litchi-sign`, and `litchi-vba`. The crate
has no dependency on a peer concrete format, the removed host, or an async
runtime. The topology ledger converts those relationships from temporary host
debt into canonical concrete-format edges, removes every `litchi-ole` package
and facade debt entry, and permanently rejects reintroducing the retired
monolith. The boundary checker accepts 32 workspace packages and 98 direct
internal dependency declarations with 18 explicit debt items, and all nine
checker regressions pass.

The extraction also closes three unsafe-by-default edges found during the
ownership audit. `sprm::parse_sprms` is now an exact fallible parser: its short
`sprm::Error` reports the malformed opcode, length, extent, or operand offset,
and a valid-looking prefix is never returned as success. Every DOC parser
caller propagates that typed failure as document corruption. Equation Native
streams use checked header-plus-payload extents, reject truncated declared
payloads, propagate those failures through aggregate extraction, and truncate
the already-owned stream buffer instead of copying it. Public header, footer,
footnote, and endnote CP builders return `Result`, validate cumulative UTF-16
positions and byte extents, and use fallible reservation instead of panicking
on user-controlled sizes.

The ownership move does not intentionally change valid DOC wire output or CFB
topology, so it creates no new native Microsoft Office artifact claim. The
extracted package passes 956 warning-denied unit and integration tests with two
ignored fixture tests; its rustdoc surface adds 14 passing compile tests with
12 intentionally ignored examples. All-feature/all-target checks and Clippy
pass for `litchi-doc`; all-feature/all-target Clippy passes for the umbrella;
warning-denied rustdoc passes for both the direct crate and isolated
`doc,imgconv` facade. The isolated `doc`, `doc,imgconv`, and `doc,ooxml`
combinations compile, `ooxml_encryption` does not activate `litchi-doc`, and
the Python binding resolves the renamed feature. Formatting, diff validation,
the 32-package boundary inventory, and all nine boundary regressions are green.
Per the explicit review direction, the full-workspace and Computer Use/native
Office reruns are skipped; the previously green baselines remain the applicable
evidence for unchanged wire semantics.

The thirty-seventh implementation slice moves the shared classic-chart and
SmartArt grammar out of the OOXML migration host. Fourteen source files and
17,401 lines now have one canonical owner under `litchi-drawingml`: the
singular `chart` and `diagram` modules. Context replaces repeated prefixes, so
the chart internals use short `model` and `data` modules rather than nested
`chart::chart` and `chart::models` paths. The public codec entry points are
`chart::reader::read`, `chart::writer::write`, and the focused low-level
`chart::writer::write_with_rels` relationship seam. No plural-module,
old-function-name, or retired root `charts`/`diagrams` compatibility aliases
remain.

DOCX and PPTX now consume `litchi_drawingml::diagram` directly, while XLSX and
XLSB consume `litchi_drawingml::chart` directly. The shared owner depends only
on neutral core and OOXML-common vocabulary; it has no dependency on a concrete
document format or on `litchi-ooxml`. Concrete packages continue to own package
relationships, anchors, and resource graphs. The umbrella exposes the same
ownership through the concise `litchi::drawing::{chart, diagram}` facade under
the existing `ooxml` feature. This is a breaking ownership move, not a
compatibility tunnel through the migration host.

The move also hardens the newly independent parser boundary. Invalid chart
states that previously relied on `unreachable!` return the short crate-local
`Error`; diagram node, depth, text, and XML-depth accounting uses checked
arithmetic; and the shared production modules contain no `unwrap`, `expect`,
`panic!`, `unreachable!`, `todo!`, or `unimplemented!` path. Host errors wrap
the typed DrawingML error without reducing it to an unstructured string. The
touched move-style pivot-chart constructor likewise returns `Result<Self>`
instead of asserting that its built-in extension fragment remains valid.

Focused verification passes with warnings denied for all targets of
`litchi-drawingml` and all features and targets of `litchi-ooxml`.
`litchi-drawingml` passes 47 unit tests, one compiled rustdoc example, and
Clippy; the host passes four shared-OOXML adapter tests, 140 chart-filtered unit
tests, 20 SmartArt-filtered unit tests, and the 17 directly affected pivot-chart
tests. Clippy also passes for every host target and the isolated umbrella
facade. The isolated `ooxml` umbrella library and comprehensive XLSX chart
example compile with warnings denied.
Formatting, diff validation, the 32-package boundary inventory, and all nine
boundary regressions are green; the inventory now contains 101 direct internal
dependencies and 18 explicit migration-debt entries. Because this slice moves
ownership and hardens failures without intentionally changing emitted Office
artifacts, the full-workspace and Computer Use/native Microsoft Office reruns
are skipped per the explicit review direction.

The thirty-eighth implementation slice removes custom document properties and
Custom XML Data Storage from the OOXML migration host. Their sole canonical
owners are now `litchi-ooxml-common::{custom, custom_xml}`. The host's former
`custom_properties` and `custom_xml_data` files, root modules, long type names,
and function aliases are deleted rather than retained as a compatibility
tunnel. The umbrella consumes the common crate directly and exposes these
owners as `litchi::ooxml::{custom, custom_xml}`.

Custom properties now use the compact `custom::{Props, Value}` facade. Insert
is fallible and move-first; lookup, containment, and removal use canonical
Unicode caseless identity while preserving the original name spelling. The
parser applies explicit byte, depth, node, attribute, count, name, and text
budgets; rejects DTDs, malformed namespaces/cardinality, duplicate names or
PIDs, forbidden format IDs, illegal XML characters, and non-finite floats; and
uses checked PID allocation. Per local `[MS-OE376]` evidence, the canonical
Office format ID requires PIDs of at least two and case-insensitively unique
names, while `vt:filetime` uses the RFC3339/XML-date-time lexical form rather
than a numeric Windows FILETIME counter. Parsed `lpstr` versus `lpwstr` remains
lossless, and deterministic output follows PID order.

`Props::read` follows the package relationship instead of guessing a path. A
genuinely absent part is empty, while duplicate, external, orphaned,
wrong-content-type, colliding, or malformed graphs fail construction. DOCX
`open`, `from_reader`, and `from_opc_package` now propagate that typed failure;
they no longer silently replace corruption with an empty set. `Props::write`
updates a valid alternate target, creates the canonical target when absent, and
removes both part and relationship when cleared. Byte-identical writes preserve
signatures; actual graph or byte changes unsign the package.

Custom XML uses `Conformance`, `Props`, immutable loaded `Item`, and consuming
`NewProps`/`NewItem` capabilities. Grouping the properties part, relationship,
and value makes partial properties creation unrepresentable. Strict and
transitional vocabularies, MCE, declarations, namespaces, QNames, character
data, content types, item GUIDs, depth, elements, strings, parts, and package
relationships are bounded and validated without resolving a schema or running
XPath. Creation performs every fallible preparation before mutation, rolls back
defensive failures, and invalidates signatures only after commit. Loaded
payloads share the OPC part's immutable allocation, and `Item::xml()` lends a
slice; repeated relationships no longer multiply large payload allocations or
enable hidden aggregate clones.

The concrete DOCX facade also adopts contextual names and semantic verbs:
`docx::custom_xml::{NewStore, Binding, Part}` and
`Package::{custom_xml, custom_xml_by_id, add_custom_xml, set_custom_xml,
replace_custom_xml, remove_custom_xml, order_custom_xml,
custom_xml_bindings, validate_custom_xml_bindings}`. Custom properties are
`custom_props` and `custom_props_mut`. All former long methods and aliases are
removed. Binding-aware deletion still refuses to remove a referenced item;
shared-target deletion preserves parts with unrelated remaining references.

A read-only audit resolves the next independent common-OOXML dependency chain:
extract `embedded` first, `ribbon` second, and `web` last. Embedded objects are
already borrowed but need a complete host MIME policy and memoized validation.
Ribbon storage must eliminate payload copying during name allocation, validate
the Ribbon part's image-only relationships, and add direct removal. Web
extensions require the largest redesign: private valid states without raw
relationship IDs in the semantic facade, shared snapshot storage, indexed
graph lookup, checked arithmetic, explicit rollback, and removal of production
`expect`/`unreachable!` paths. Those are subsequent atomic slices, not hidden
inside this ownership cut.

Focused verification for this slice passes with warnings denied: all 49
`litchi-ooxml-common` library tests; two DOCX custom-property graph tests; five
DOCX Custom XML CRUD, binding, rollback, sharing, and malformed-graph tests;
and 16 bibliography-filtered host tests. All-feature/all-target checks pass for
the host. Common and host warning-denied all-target Clippy and rustdoc pass; the
isolated `ooxml` umbrella library and rustdoc pass. Formatting, manifest
ordering, diff validation, the 32-package boundary inventory, and all nine
regressions are green; the inventory contains 102 internal dependency edges and
18 explicit migration-debt entries.
Per explicit direction, the full-workspace rerun is skipped. This slice changes
custom-property and Custom XML writing behavior, so it makes no new native
Microsoft Office compatibility claim without a future artifact/open/edit/
resave/reverse-read run.

The thirty-ninth implementation slice completes the first dependency in that
audit by moving embedded-object and embedded-package inventory into the sole
canonical owner `litchi-ooxml-common::embedded`. The migration host's former
`embedded_object` module and root `EmbeddedPart*` aliases are deleted. The
compact shared vocabulary is `Kind::{Object, Package}`, borrowed `Payload`,
`Target`, `Entry`, and configurable `Limits`; `scan` applies safe defaults and
`scan_with` is the explicit lower-layer resource seam. DOCX, PPTX, XLSX, and
XLSB expose only the concise safe-default `embedded` facade. The umbrella
exports the common module directly as `litchi::ooxml::embedded`.

Every internal payload remains owned by its OPC part. `Payload::bytes` lends
that allocation, so inventory construction performs no payload copy and
entries cannot outlive their package. Duplicate relationship occurrences stay
visible, but canonical target names memoize outbound-relationship validation
and charge the aggregate payload budget only once. Entry and aggregate payload
relationship counters use checked arithmetic, result ordering is deterministic
by source and relationship ID, and production paths contain no panic macro.
These are allocation and validation-topology properties, not measured latency,
CPU, cache, or contention claims; those still require ADR 0005 profiling.

The source policy now covers every ISO OOXML Word, SpreadsheetML,
PresentationML, and chart source used by the two relationship families,
including Word document/template macro main-part content types. It also applies
the kind-specific additions in local `[MS-XLSB]` File Structure sections
2.1.7.36 and 2.1.7.37: binary worksheet, macro-sheet, and dialog-sheet parts may
source Object or Package relationships, while a binary external-link part may
source Object only. A real POI workbook proves that an Object target can
legitimately declare the server-specific `application/vnd.ms-excel` content
type; the common layer therefore preserves declared MIME opaquely instead of
guessing from bytes or enforcing a single container type. Strict and
transitional relationships are accepted, internal query/fragment targets and
missing or forbidden graphs fail with typed errors, and payload parts may own
only hyperlink relationships. External targets are returned inertly and never
contacted.

Focused verification passes with warnings denied: all 62
`litchi-ooxml-common` library tests, including 13 embedded inventory tests over
seven real producer fixtures and hostile synthetic graphs; eight DOCX OLE
writer tests; seven DOCX SmartArt writer tests; and two XLSX/XLSB facade tests.
Common and host all-target checks, Clippy, and rustdoc pass, including the
all-feature/all-target host build. Formatting and diff validation are green.
The isolated `ooxml` umbrella library and rustdoc, manifest ordering, the
32-package boundary inventory, and all nine boundary regressions are green; the
inventory remains at 102 internal dependency edges and 18 explicit
migration-debt entries. This slice changes only ownership, validation, and read
facades; it emits no new Office artifact, so the previously green full-workspace
and native Office baselines are relied upon instead of repeating those gates.
The next dependency-safe extraction is `ribbon`, followed by `web`.

The fortieth implementation slice completes that second dependency by moving
Ribbon customization ownership into the sole canonical
`litchi-ooxml-common::ribbon` module. The migration host's former `ribbonx`
module and its owned `RibbonCustomization*` vocabulary are deleted without
compatibility aliases. The contextual public surface is
`Version::{V2007, V2010, Ui2}`, `Family::{Legacy, Modern}`, borrowed `Ui`, a
fixed two-slot `Set`, configurable `Limits`, and the semantic verbs `load`,
`load_with`, `put`, `put_with`, and `remove`. `Set::effective` selects the
modern slot without allocating a temporary vector, while explicit family
accessors keep coexistence visible. DOCX, PPTX, XLSX, and XLSB expose only
`ribbon`, move-owning `put_ribbon`, and `remove_ribbon`; PowerPoint's immutable
presentation view exposes `ribbon`. The common module is re-exported for direct
host and umbrella users, so naming `Version` or `Family` does not require a
second dependency.

Loaded XML remains owned by its OPC part. `Ui::xml` lends that allocation and
cannot outlive the package; the read path no longer clones each payload.
Authoring accepts a `Vec<u8>` by value. Part-name probing validates candidate
URIs against one bounded, case-folded name snapshot before constructing one
XML part, so collisions do not repeatedly copy the payload or rescan the
package for every suffix. Byte-identical `put` is a true no-op that preserves
package signatures. Replacing a part that has another inbound relationship
forks it together with its validated image relationships before the package
relationship is retargeted, preserving unrelated consumers; otherwise
replacement updates in place. Successful updates, creates, and removals
invalidate signatures inside the common transaction. These are ownership and
algorithm-topology properties, not measured latency, CPU, cache, or concurrency
claims; ADR 0005 profiling remains required before making those claims.

The validator follows local `[MS-OE376]` section 3.4.1.3: Ribbon parts use
`application/xml`, are reached by an internal package relationship, have the
matching `customUI` root namespace, and may relate only to image parts. It also
applies the newer relationship and namespace family documented by the local
Office specifications. Both family cardinality and XML bytes, depth, nodes,
aggregate image relationships, package graph scans, allocation-name snapshots,
and deletion traversal are bounded. XML declarations, qualified names,
namespace bindings, expanded attribute identities, numeric references, and XML
1.0 characters are validated. Declarations may identify only the UTF-8
encoding that the zero-copy parser actually consumes; inert processing
instructions remain accepted.
Package and image targets with queries or fragments, external targets, missing
parts, mismatched namespaces or content types, duplicate families, Ribbon
relationships sourced by parts, and non-image outbound relationships fail with
typed errors. Image payloads stay inert and are never fetched or decoded.

`remove(Family)` validates and stages the graph before mutation. It removes the
selected package relationship, collects the selected Ribbon part only when no
other internal relationship still targets it, and similarly collects its image
parts only when they are no longer shared. Absence is `Ok(false)` and does not
drop signatures. This gives direct safe deletion without exposing raw
relationship IDs or leaving the caller to hand-edit a partially valid OPC
graph.

Focused verification passes with warnings denied: all 76
`litchi-ooxml-common` library tests, including 14 Ribbon tests, and ten host
integration tests across PPTX, DOCX/general OOXML, XLSX, and XLSB are green.
Common and host all-target checks, Clippy, and rustdoc pass, including the
all-feature/all-target host build. Formatting and diff validation are green.
The isolated `ooxml` umbrella library and rustdoc, manifest ordering, the
32-package boundary inventory, and all nine boundary regressions are green;
the inventory remains at 102 internal dependency edges and 18 explicit
migration-debt entries. Per explicit direction, the already-green
full-workspace gate and native Microsoft Office baseline are not rerun because
this ownership and validation slice emits no new Office artifact. At this
point, the remaining dependency-safe extraction was `web`.

The forty-first implementation slice completes that extraction by making
`litchi-ooxml-common::web` the sole canonical owner of persisted task panes and
Office Add-in metadata. The migration host's former `web_extensions` module is
deleted without compatibility aliases. The compact shared vocabulary is
`Panes`, `Pane`, `AddIn`, `Reference`, `Property`, `Binding`, `BindingKind`,
`Store`, `Dock`, `Snapshot`, `Image`, typed `Link::{Internal, External}`,
`Compression`, `Effect`, `ExtList`, configurable `Limits`, `Conformance`, and
`Selector`. The semantic entry points are `load`/`load_with`, consuming
`put`/`put_with`, and `remove`/`remove_with`. DOCX, PPTX, XLSX, and XLSB expose
`task_panes`, move-owning `put_task_panes`, and `remove_task_panes` on their
host facades; PowerPoint's immutable `Presentation` view also exposes
`task_panes`. XLSX worksheet-range validation and XLSB binary binding
validation consume the canonical `web::Binding` model. The umbrella exports
the common module directly as `litchi::ooxml::web`. The current XLSX facade and
`x15:webExtensions` range-binding implementation still live in the migration
host; extracting their canonical ownership into `litchi-xlsx` remains future
work.

Semantic record fields remain private behind checked constructors and short
mutators; `Limits` remains the explicit configurable resource-policy record.
Add-in IDs are the primary pane selector, while checked numeric positions stay
available for ordered workflows; relationship IDs do not become application
keys. `Panes::edit` replaces unrestricted collection-wide mutable access: it
clones one selected pane, runs a fallible edit closure, rechecks collection
identity and resource invariants, and swaps the candidate only on success.
Failure leaves the original pane untouched. `push` and `edit` canonicalize
internal snapshot resources both within one pane and across panes:
case-equivalent names with identical content type and bytes reuse one canonical
part name, while disagreeing resources fail before publication. Bindings,
properties, alternate references, pane state, every retained `extLst`, and
embedded or external snapshot resources support semantic CRUD. Embedded
snapshot bytes use shared `Arc<Vec<u8>>` ownership through
`BlobPart::new_shared`; transactional pane clones and borrowed `Image` views
share or lend that payload rather than copying it. External targets remain
typed, inert data and are never contacted or activated.

One bounded, ASCII-case-folded package graph index records canonical part
names and inbound and outbound internal edges. `Limits` covers per-part and
operation-wide XML, retained string, indexed/authored package-metadata, and
image bytes, plus XML depth and nodes, collection sizes, package parts and
relationships, name-allocation probes, and deletions. One operation budget is
threaded through the complete load/put/remove call, and one allocation-probe
counter covers task-pane and add-in name searches together. XML parsing uses
persistent, parent-linked `Arc` namespace scopes, including for retained
fragments, so depth does not require cloning the complete in-scope namespace
map. These are bounded ownership and algorithm-topology properties, not
workload performance measurements. Internal task-pane, add-in, and image
targets reject external-mode misuse and query or fragment suffixes; expected
roots, namespaces, content types, and outbound relationship families are
validated. Case-equivalent authored part replacement is rejected before
mutation. All fallible relationship construction is staged before infallible
part-map publication.

Sharing is handled conservatively across the whole owned graph. An owned part
with ingress from outside the task-pane graph protects itself and every owned
descendant reachable from it. `put` refuses to change a protected part, while
`remove` retains the protected transitive closure and deletes only unprotected
old parts. This is safe refusal, not yet copy-on-write graph forking. A
byte-identical `put` returns before any package mutation or signature
invalidation; absent removal is likewise a no-op. Successful create, change,
or removal invalidates signatures only after the staged commit, and errors
leave the original package and signature state intact. XLSX worksheet binding
replacement follows the same byte-change rule.

Spreadsheet host facades enforce package/binding integrity in both mutation
directions. XLSX validates every effective worksheet `appRef`, including queued
worksheet mutations, against exactly one binding in candidate task panes
before `put_task_panes`; `remove_task_panes` proves that no worksheet binding
would dangle, while worksheet binding replacement validates against the
current package graph. The new XLSB `xlsb::Workbook::{task_panes,
put_task_panes, remove_task_panes}` facade applies the same rule to every
binary worksheet `BrtWebExtension.appRef`; candidate task panes are rejected
when a binary reference is missing or ambiguous, and removal is refused while
any binary binding remains. These invariants are checked by the host operation
rather than left as caller convention. This follows the checked-in
`[MS-XLSB]` sections 2.4.303, 2.4.655, and 2.4.868 record definitions; in
particular, section 2.4.868 requires `BrtWebExtension.appRef` to equal its
`CT_OsfWebExtensionBinding` identifier.

The checked-in `3rdparty/` specification mirror does not currently contain an
`[MS-OWEXML]` source, so this slice does not claim a source-complete local
conformance audit. The authoritative references are Microsoft's
[`[MS-OWEXML]` publication page](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-owexml/a2cd741a-4cca-4b1a-ade4-b2c443972afa),
its [format overview](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-owexml/29f59f30-b835-461a-bd8a-ca400a7bc717),
and the [content Web Extension binding example](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-owexml/5b150f17-59a1-4bec-874e-83d25ef6eec9).
Vendoring the applicable published source and completing a section-by-section
review remain certification work.

Focused verification is green: all 113 `litchi-ooxml-common` library tests;
ten XLSX/XLSB web-binding tests; three host task-pane tests; the XLSX
byte-change/signature regression; the shared `BlobPart` allocation regression;
and three tests in each changed package and PPTX integration suite pass.
Warning-denied Clippy and rustdoc pass for the common crate, OOXML library, and
isolated `ooxml` umbrella; their corresponding focused checks pass as well.
Formatting, diff validation, manifest ordering, the 32-package boundary
inventory, and all nine boundary regressions are green. The inventory remains
at 102 internal dependency edges and 18 explicit migration-debt entries. Per
explicit direction, the previously green full-workspace gate was not repeated.
No native Microsoft Office artifact was opened, edited, resaved, and
reverse-read for this slice, so it adds no Office compatibility or measured
performance claim.

The forty-second implementation slice moves the remaining worksheet Web
Extensions ownership into `litchi-xlsx`. The canonical public vocabulary is
the short `web::{Binding, Bindings, Selector, Refs}` surface. Application
references are semantic primary keys; checked zero-based positions remain
available for ordered workflows. `add`, `put`, `edit`, `remove`, and `clear`
provide bounded CRUD without exposing relationship IDs. Inserts and
replacements consume values, removals return ownership, and transactional
`edit` clones only the selected small binding before validating and swapping
it. Whole binding collections can be moved directly into an edit. The old
migration-host codec and its long names are deleted without compatibility
aliases; `litchi-ooxml` now delegates its four worksheet-binding operations to
the canonical crate.

`Workbook::task_panes` and `Sheet::web_bindings` return borrowed models owned by
their immutable snapshot. Independent per-resource
`once_cell::sync::OnceCell` caches provide fallible single-flight first reads:
successful concurrent readers of one resource share one parse, while a failed
initialization remains retryable and unrelated sheets do not contend on a
workbook-global lock. Snapshot clones share the cached model and later
snapshots receive fresh caches, so no `Arc<RwLock<_>>` or mutable global facade
leaks into the API. The cache regression exercises concurrent first access,
pointer identity, absent package state, and snapshot scoping.

The worksheet codec `raw::web::{read, write, replace}` is bounded by XML size,
depth, item, individual-string, and aggregate-string ceilings. It rejects DTDs,
unknown prefixes, duplicate selected extensions and `appRef` values, malformed
roots, unexpected binding children or attributes, and any binding without
exactly one `xm:f`. Replacement changes only the selected extension byte span,
retaining unrelated worksheet XML, and derives the output SpreadsheetML
namespace from the source so Strict and Transitional worksheets stay in their
original vocabulary. Empty replacement removes only the selected extension;
an absent extension is a semantic no-op.

The implementation follows the published `[MS-XLSX]`
[worksheet extension rule](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/07d607af-5618-4ca2-b683-6a78dc0d9627),
including URI `{F7C9EE02-42E1-4005-9D12-6889AFFD525C}`, and the
[`CT_WebExtension` rule](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/386851b6-b7b6-42b8-8cf1-d94bab7b0731)
that requires one formula and an `appRef` matching the package-level
MS-OWEXML binding. The high-level constructor currently accepts a valid
`external-cell-reference` subset: a single sheet-qualified A1 cell or
rectangular area, including quoted worksheet names and absolute coordinates.
This is distinct from the forbidden `bang-reference` form `!A1`; an unqualified
local `A1` and a multi-sheet `Sheet1:Sheet2!A1` are also rejected. Optional
external-book prefixes and the wider permitted `ref-nospace-expression`
surface are not yet implemented, so this is not a claim of complete
formula-grammar conformance. The checked-in `3rdparty/`
mirror contains neither `[MS-XLSX]` nor `[MS-OWEXML]`; vendoring the applicable
published sources and completing the broader grammar audit remain
certification work.

`Workbook::Edit` now owns package task panes and worksheet bindings in one
transaction. Independently prepared pane and range edits can join when their
effects are disjoint, then commit validates every effective existing and new
worksheet reference against the final task-pane graph. Same-sheet whole-set
binding writes conflict deterministically; sheet removals conflict with pane
edits in either join direction. Removing task panes while a range remains
returns `DanglingWebBinding { sheet, app_ref }`. Clearing those ranges and
removing the graph in the same edit succeeds. A tab rename rewrites the
qualifying worksheet name in `xm:f` and records the final semantic before/after
state rather than leaving the patch metadata stale.

The resulting workbook patch combines exact `Arc<Vec<u8>>` worksheet-part
deltas with the common crate's opaque task-pane graph patch. All source parts,
relationships, destination names, and newly shared source or destination
parts are checked before the caller-visible snapshot can change. Applying to a
stale or different package fails; inverse patches swap shared allocations
rather than copy payload bytes. No-op plans preserve the original snapshot and
signature state, while a real common graph change invalidates signatures only
after its staged checks. Commit and patch application reparse the candidate
package and revalidate the complete cross-graph invariant before publication.

Computer Use verification against desktop Excel on macOS found one additional
safe-authoring invariant that schema-only tests missed. A generated workbook
whose primary reference used `storeType="FileSystem"` without `store` caused
Excel's repair dialog and could not be opened. Differential artifacts then
changed one factor at a time: arbitrary synchronized extension-instance and
`appRef` strings opened cleanly when `store="C:\Example"` was present, while
removing only that attribute reproduced the repair. ZIP integrity and package
topology were otherwise identical. The generic MS-OWEXML schema makes `store`
optional, but Microsoft's
[provider tuple guidance](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/automatically-open-a-task-pane-with-a-document#use-open-xml-to-tag-the-document)
defines it as the file-system share for this provider, matching the observed
desktop behavior. The evidence does not justify imposing a GUID grammar on
the extension instance or `appRef`; their normative types remain strings and
cross-part equality remains the XLSX invariant.

The safe common model therefore cannot represent a store-less file reference.
`Reference::file(id, version, location)` requires all three fields in one call
and rejects an empty location; generic
`Reference::new(..., Store::FileSystem)` and parsing of that location-less form
fail with typed errors. It does not yet validate manifest identity/version or
whether the location resolves to a deployed share. The misleading `catalog`
builder/accessor was replaced by `location`/`location_name` without an alias.
This is deliberately an Office-safe presence invariant over the permissive raw
schema; a future lower-level repair API may retain malformed source markup
without admitting it into the authoring model.

After that correction, Excel opened the generated example with the same
arbitrary instance and `appRef` strings without repair, changed B2 from 42 to
84, saved, and reopened it without a compatibility or repair dialog. Excel
relocated the web-extension parts under `xl/webextensions`, added its ordinary
producer metadata, and retained the file-system location, package binding,
worksheet `appRef`, and `xm:f`. `Workbook::task_panes`,
`Sheet::web_bindings`, and the ordinary cell facade reverse-read the
Office-saved artifact as `sales-range`, `Sheet1!$A$1:$B$2`, and numeric 84.
The `web_bindings` example now demonstrates atomic authoring and graph
readback; the existing `open` example supplied the ordinary cell readback.
This certifies only the tested inert binding graph and open/edit/save/reopen
path; it does not certify manifest discovery or add-in activation, other
providers, Windows or older Office builds, the wider formula grammar, or
performance.

Focused verification is green: 44 common Web Extensions tests, ten canonical
XLSX web/raw/cache/transaction tests, the symmetric pane/removal join
regression, three migration-host worksheet-binding tests, one host
task-pane-integrity test, and three package integration tests pass.
Warning-denied all-target Clippy and warning-denied rustdoc pass for
`litchi-ooxml-common`, `litchi-xlsx`, and `litchi-ooxml`. Formatting, diff
validation, and the boundary checker are green; the checker still accepts 32
workspace packages and 102 direct internal dependency declarations with 18
explicit migration-debt items. Per explicit direction, the previously green
full-workspace gate was not repeated. Single-flight behavior and shared
payload ownership are topology guarantees, not measured memory, latency, CPU,
cache, affinity, or contention results; ADR 0005 benchmarks and flame graphs
remain required before making those performance claims.

These slices do not yet shrink or synthesize absent worksheet `dimension`
hints or implement mixed deletion disposition, non-worksheet tab deletion,
recursive garbage collection, grouped-tab selection CRUD, workbook-protection
unlocking, row and column insertion/deletion, shifting references,
group-formula edits, dynamic-reference resolution, validation evaluation,
shared-style definition editing or forking, named-style and row/column/theme
resolution, rich text,
dynamic arrays, patch serialization, full structured diagnostics,
eviction/resource budgets, general range/structural effect joins, three-way
merge,
raw-copy preservation of clean compressed entries, cancellation-aware save
contexts, scratch planning, or output budgets. Those remain certification work;
no allocation, latency, contention, or scaling conclusion follows from the
functional tests.

The next dependency-decoupling slice establishes the first canonical owners
for all three remaining concrete OOXML families in parallel:
`litchi-docx`, `litchi-pptx`, and `litchi-xlsb`. The new crates are
runtime-neutral and have no concrete peer-format edges. The OOXML monolith
depends on them only as an explicitly ordered migration host and converts their
typed errors without string erasure. The old host-owned modules and flattened
type aliases are removed; callers use contextual canonical modules and short
names. The facade still depends on the migration host until each concrete
package/object graph moves, so this slice does not misrepresent a partial crate
as the completed top-level DOCX, PPTX, or XLSB facade.

`litchi-docx::font` now owns WordprocessingML font-table XML and its OPC
relationship graph. `Table` provides safe semantic name lookup as the ordinary
entry and checked discovery-order lookup where a numeric position is useful;
its add, replace, remove, and reorder operations return typed errors rather
than indexing panics. Lookup and all mutations use the same NFD plus Unicode
default-case-fold identity, including composed/decomposed and non-ASCII names;
malformed producer tables with ambiguous identities remain readable but cannot
be selected silently. Names and extension attributes are validated at
construction. Package `read` shares embedded-font payload allocations,
`Resource::bytes` lends a slice, and consuming `put` moves newly owned payloads
into the graph instead of cloning the complete table and every font program.
The service bounds XML, nodes, depth, font count, individual/aggregate payload
bytes, validates Strict/Transitional relationships and content types, checks
licensing and usage invariants, and removes only truly orphaned resources. It
never discovers, loads, interprets, renders, or executes font programs. The
migration host delegates the symmetric `fonts`, `put_fonts`, and
`remove_fonts` facade directly to this owner.

`litchi-pptx::transition` now owns the bounded transition reader/writer and a
semantic model whose effect variants contain only their legal option domain.
`Side`, `Axis`, `Corner`, `Origin`, `InOut`, `Shape`, `Ripple`, and `Spokes`
replace an effect plus independently invalid direction fields. Checked `Ms`
values cover custom duration and timed advance. Unknown direct effect and
extension children are retained as size-bounded shared inert XML with no safe
public raw constructor. Constructor-only effects, arbitrary writer-rejected
strings, and the former sound field are not kept as false typed capabilities;
they can return only after parse and write both preserve them. Slide, layout,
master, inheritance, and mutable-writer callers consume the canonical model
directly.

`litchi-xlsb::raw` now owns BIFF12 record framing independently of XLSX and
OPC. `Kind` admits the complete 14-bit future-record domain, constants live in
`raw::kind`, `Record` lends its payload from the input, and `Records`
distinguishes a clean record boundary at end-of-stream from a truncated kind,
length, or payload. `Header` and `Writer` implement the exact one/two-byte kind
and one-to-four-byte length grammar in checked-in `[MS-XLSB]` section 2.1.4;
the fourth length continuation bit is ignored without consuming a fifth byte.
Explicit finite payload and UTF-16 budgets apply before allocation or output,
`Cursor` has no guarded panic paths, strict UTF-16 rejects unpaired
surrogates, blob reads lend slices, and wide-string writing counts then streams
code units without collecting a temporary vector. Validated `Header` and
borrowed `Record` fields are private, with short `kind`, `len`, and `payload`
accessors preventing callers from manufacturing mismatched records. Checked-in
`[MS-XLSB]` section 2.5.123 also pins RK flag positions, signed 30-bit decoding,
and floating reconstruction: `Cursor::read_rk` owns decoding, and
`Writer::write_rk` emits only bit-exact direct, divide-by-100, or floating
representations and returns a typed error instead of rounding other values.
Both byte- and 32-bit Boolean reads reject values outside zero and one.
Existing semantic records remain temporarily in the host but delegate framing
and these scalar wire rules to this owner.

This ownership slice changes library boundaries, validation, and previously
false or permissive API states. It does not establish new native Office file
support or make a performance claim. Per explicit direction, the previously
green full-workspace and desktop Office gates are not repeated; focused crate,
host-regression, warning-denied, documentation, and dependency-direction gates
are the applicable evidence. Memory sharing, borrowed payload identity, and
bounded wire traversal are structural properties; latency, CPU, cache,
affinity, contention, and scaling still require ADR 0005 measurement and flame
graphs.

Focused evidence for this slice is green. Canonical owner tests pass with 9
DOCX unit tests plus 1 downstream-API test, 13 PPTX unit tests plus one passing
and one compile-fail documentation test, and 17 XLSB raw-wire integration
tests. Migration-host regressions pass with 1 DOCX font test, 20 PPTX
transition tests, 388 XLSB library tests, and 15 focused XLSB package tests.
All-target Clippy with warnings denied passes for the three owner crates and
`litchi-ooxml`; warning-denied checks also pass for the three migrated umbrella
examples. Rustdoc passes with warnings denied for all three owners and the
migration host. Scoped formatting and diff checks are clean. The executable
dependency fence accepts 35 workspace packages and 110 internal dependency
declarations with 21 explicitly ordered debt items, and all 10 boundary-policy
regression tests pass. These are the intentionally focused gates described
above, not a repeated full-workspace or native-Office certification run.

The subsequent dependency-decoupling slice gives DOCX alternative-format
imports, PPTX programmable tags, and XLSB workbook calculation properties their
canonical concrete owners, and removes the formula evaluator's production
runtime anchor. The migration host consumes these owners directly; the former
host alternative-format, tag, and calculation modules and their flattened long
names are deleted without compatibility aliases.

`litchi-docx::alt` owns typed alternative-format anchors and opaque package
payloads. `Chunk`, `Conformance`, `Data`, `Import`, `Kind`, `Part`, and `Target`
form the ordinary API; checked `Rel` and `Uri` values remain the low-level
metadata layer. `Data` and `Import` cannot be cloned, and internal insertion
moves the original byte allocation directly into the OPC part. Word's ten
document, template, MIME, HTML/XHTML, RTF, text, and XML media families come
from checked-in `[MS-OI29500]` section 2.1.527 and `[MS-OE376]` section 2.1.558,
including Word's case-sensitive Transitional `aFChunk` spelling. Reads lend
unknown or recognized payload bytes without opening nested packages; external
targets are returned as inert text and never contacted. The host exposes
ordered `add_alt`, `insert_alt`, `replace_alt`, `remove_alt`, and `move_alt`
package operations while keeping raw-ID writer mutation private. Parser and
authoring defaults cap payloads at 128 MiB, source XML at 32 MiB, nesting at 256
levels, and anchors at 4,096; unknown nested extension XML remains preserved.
Bounded markup-compatibility processing acts only as a visibility oracle: read
and mutable selectors retain original source coordinates, agree on the active
Choice/Fallback branch, and preserve inherited Strict or Transitional namespace
aliases plus inactive branch bytes.

`litchi-pptx::tag` owns bounded Strict/Transitional `tagLst` parsing, writing,
and slide relationship discovery. Its short `List`, `Tag`, `Key`, `Source`,
`Conformance`, and `raw::Attr` vocabulary keeps semantic name selection as the
ordinary path and checked zero-based positions as the repair path. `List`
supports detached add, insert, replace, value set, remove, and complete reorder
operations; inserted and replacement values move into place, and removal moves
the old value back to the caller. Lookup and every mutation use one
NFD/default-case-fold/NFD identity chosen by Litchi as a deterministic
implementation of the case-insensitive uniqueness requirement in checked-in
`[MS-OE376]` section 2.1.1170(c); the specification does not prescribe that
normalization algorithm. A malformed producer list with equivalent names is
retained for numeric inspection, while semantic lookup returns typed ambiguity
and new mutations cannot create another equivalent name. Tag values and bounded
extension attributes remain inert. Cached exact escaped-wire sizes preflight
the aggregate 8 MiB writer ceiling before every successful mutation and keep
replacement/value-set failure atomic without rescanning existing strings.
Discovery refuses external tag targets, wrong content types, duplicate targets,
and unexpected relationships on a tag part. Direct presentation and
common-slide-data anchors now have atomic package `load`, `put`, and `remove`;
the migration facade exposes these through semantic slide selectors while raw
relationship inventory remains a low-level diagnostic path. The separate
revision/change package stores remain add-only.

`litchi-xlsb::calc` owns `BrtCalcProp` independently of the OOXML migration
host. `Props` keeps all fields private and combines short checked setters with
consuming `with_*` builders. `Mode` closes the wire enumeration, `Opts` packs
the nine switches into a `u16` and rejects reserved bits, `Threads` enforces the
`1..=1024` domain, and `Delta` excludes NaN, infinity, subnormal values, and
negative zero. These are the checked-in `[MS-XLSB]` section 2.4.318 record rules
and section 2.5.172 `Xnum` rules, not guessed ergonomic restrictions. `read`
accepts the canonical 26-byte payload and the exact 25-byte one-byte-option-tail
form found in the checked-in Microsoft Excel 12 `Simple.xlsb` artifact. It
zero-extends that historical tail directly from the borrowed cursor, rejects
every other length, and does not weaken any semantic validation. `write`
streams the canonical 26-byte form without allocating an intermediate record
buffer. The parsed workbook exposes `calc`; the authoring host uses `calc`,
`calc_mut`, and move-accepting `put_calc`.

The `litchi-eval` production graph no longer contains Tokio or Reqwest.
Feature-gated web functions accept an explicit borrowed `Fetch` capability via
`FormulaEvaluator::with_fetch`; the trait returns a runtime-neutral future, so
the caller chooses the executor, transport, and network policy. Without that
capability, `WEBSERVICE` is network-inert and returns a connection cell error.
With one, URL shape, response bytes, UTF-8, and the 32,767-UTF-16-unit cell limit
are checked before a value is accepted. The evaluator also replaces shared
position and circular-visit state with a method-scoped `At` view over a borrowed
evaluation session. Concurrent top-level calls therefore cannot report one
another as a cycle, and a scoped visit guard removes its marker on success or
error. Private synchronous cache locks are held only for short cache operations
and never cross an await; no lock type enters the public facade.

The dependency policy now distinguishes production runtime coupling from test
execution: runtime-neutral crates are checked against normal Cargo edges,
including optional normal dependencies, while development-only runtimes remain
available to tests. `litchi-eval` is in that enforced set. This does not relax
the separate inventory of internal normal, optional, development, renamed, and
target-specific edges.

Focused evidence is green. The DOCX owner passes 20 unit tests, two downstream
API tests, and one doctest; 11 migration-host alternative-format tests pass.
Its canonical all-target and focused-host warning-denied Clippy, warning-denied
rustdoc, formatting, and diff gates pass. The PPTX owner passes 24 unit tests
and three documentation cases; four migration-host integration tests pass,
including discovery of seven real LibreOffice tag parts. Its all-target
warning-denied Clippy and warning-denied rustdoc gates pass. The XLSB owner
passes two unit tests, eight calculation integration tests, 17 existing
raw-wire tests, one passing doctest, and two compile-fail doctests, plus
all-target warning-denied Clippy. The evaluator passes 1,147 default-feature
and 1,161 all-feature unit tests, including all 17 focused web-function tests,
plus all-feature all-target warning-denied Clippy, warning-denied rustdoc, and
the umbrella's isolated
`eval_engine_web_functions` check. Normal dependency-tree checks find neither
Tokio nor Reqwest. The executable boundary checker accepts 35 workspace
packages and 111 internal dependency declarations with 21 ordered debt items;
all 13 policy regressions, Python bytecode
compilation, and diff validation pass.

This slice changes one authored Office artifact path: the new focused example
writes a valid HTML alternative-format import between ordinary Word paragraphs.
Computer Use verification on macOS Word opened that generated DOCX without a
repair dialog and displayed the imported heading and paragraph in the expected
body order. Word accepted a new trailing paragraph and saved a second DOCX;
the resaved ZIP passed integrity checks, and Litchi reverse-read five expected
paragraphs including both imported HTML strings and the Office edit. Word
removed the alternative-format part and anchor during resave, which is the
expected producer normalization after importing the foreign content. This gate
does not certify the other nine media families, Windows or web Word, older
Office versions, or pixel-identical layout.

Per explicit user direction, the previously green full-workspace gate was not
repeated. The compact options layout, streamed record write, borrowed contexts,
move-owned payload identity, and absence of production runtime edges are
structural facts, not latency, allocation, CPU, cache, affinity, contention, or
scaling measurements. No performance claim follows without the ADR 0005
benchmark and flame-graph work.

### Canonical OOXML encrypted-container ownership

The next dependency slice moves the application-neutral `[MS-OFFCRYPTO]`
envelope out of the OOXML migration host and into the feature-gated
`litchi-crypto::ooxml` owner. Its concise vocabulary is `Kind`, `Mode`,
`Limits`, `Password`, `Opened`, and `Error`; the host exposes the same owner
contextually as `litchi_ooxml::encryption`. Password-free `inspect`,
move-consuming `open`, `encrypt`, and `rekey`, and runtime-neutral bounded
`load` are the low-level operations. The `_with` variants accept one explicit
outer-encryption resource policy. Concrete hosts still apply the OPC archive's
independently bounded defaults after decryption; exposing a composite host open
policy is subsequent migration work. Ordinary OPC input moves through `open`
without reallocating, and `Password` is move-owned, non-cloneable, redacted in
`Debug`, and zeroized on drop. Public errors distinguish invalid policy,
resource ceiling, allocation, unsupported profile, malformed container,
incorrect password, missing authoring password, integrity, randomness, XML,
and I/O failures.

This owner implements the strictly validated AES-128/SHA-1 Standard and Agile
compatibility profiles. It does not claim AES-192/256 or SHA-2 Agile support;
those profiles fail as typed `Unsupported`. Standard encryption has no package
integrity mechanism in its specification. Agile authenticates the complete
encrypted package before trusting its declared clear size. Parsing bounds the
outer input, `EncryptionInfo`, Agile XML bytes, nesting, events, attributes,
spin count, password characters, clear package, encrypted stream, and emitted
CFB container. StrongEncryptionDataSpace validation is strict by default, with
the LibreOffice missing-graph exception available only through an explicit
policy flag. Length arithmetic is checked, fallible growth reports typed
allocation or limit failures, internally encoded password bytes and derived-key
material are zeroized, and verifier comparisons are constant-time.

The CFB handoff consumes the outer buffer and moves the encrypted-package
stream into `OleWriter`; it does not clone either large payload at that seam.
Writer plans use checked sector arithmetic and fallible FAT, MiniFAT, and DIFAT
construction, and reject the version-3 2 GiB stream boundary before emitting
output. These are structural copy and bound properties, not benchmark results.
The inert VBA bridge likewise borrows its source bytes and applies an explicit
CFB limit instead of copying the whole project before parsing.

DOCX, PPTX, and XLSX retain the opened package's `Mode`. Their ordinary `save`
paths reject an encrypted source before touching the destination; callers must
choose `save_reencrypted` or the conspicuous `save_plain`. New encrypted output
requires a non-empty password, borrows that password only for the operation,
and atomically replaces a path destination after serialization and encryption
succeed. PPTX presentation modify-password metadata remains a separate concern
from outer-package encryption. This is the fail-closed migration seam; it does
not yet claim the final consuming `Locked<T>`/`Sensitive<T>` type-state facade
required when each concrete format leaves the migration host.

Presentation modification protection is also made structurally separate and
safe. `Protection` owns an optional private `ModifyVerifier`
aggregate instead of public independent algorithm, spin-count, salt, hash, and
enabled fields. Callers can inspect it through short read-only accessors but
can construct it only through strict parsing or atomic password setting. The
SHA-512 initial input follows PowerPoint's checked-in conformance rule,
`UTF-16LE(password) || salt`, followed by the little-endian iteration counter.
Unknown algorithms/SIDs, malformed attributes, missing or duplicate fields,
invalid Base64, wrong digest lengths, and out-of-policy spin counts remain
typed errors rather than silently becoming SHA-1 or default values.

The dependency fence removes the migration host's direct CFB, cipher, HMAC,
hash, and random-source ownership for outer encryption. `litchi-crypto` keeps
those dependencies optional behind `ooxml`, and the umbrella re-exports the
canonical crypto owner only when the encryption capability is selected. This
does not assert that all transitive CFB use has disappeared: VBA and other
focused Office capabilities still legitimately depend on the CFB owner.

Focused evidence for this slice is green. CFB passes 87 unit tests, four
property-set integration tests, and six executed doctests (one documentation
case is intentionally ignored). The crypto owner passes 49 unit tests, two
compile doctests, and one compile-fail doctest. The OOXML host passes all five
encrypted-package integration scenarios and six focused presentation
protection tests. The focused OPC filter passes seven tests and VBA passes all
26 unit tests. Warning-denied Clippy passes for CFB, crypto, OPC, VBA, the
OOXML host/test/example targets, and both umbrella PowerPoint examples;
warning-denied rustdoc passes for the touched owners. The isolated umbrella
examples compile with and without encryption as their declared features
require. The boundary checker accepts 35 workspace packages and 110 internal
dependency declarations with 19 explicit debt items; all 13 policy regressions,
Python bytecode compilation, formatting, and diff validation pass. Per explicit
user direction, the previously green full-workspace gate was not repeated.

Computer Use verification on current macOS Microsoft Word opened both a
generated Agile and a generated Standard encrypted DOCX. Word presented its
native password prompt for each file, accepted the synthetic test password,
showed no repair dialog, and displayed the exact two-line marker headed
`Litchi encrypted DOCX verification`. Neither encrypted document nor the
previously open unrelated Word document was edited or saved. This certifies
password open and visible content for the implemented AES-128/SHA-1 Standard
and Agile profiles on that Word build. It does not certify Office edit/resave,
PowerPoint or Excel UI behavior, AES-192/256 or SHA-2 profiles, Windows/web or
older Office releases, or performance.

## PPTX programmable-tag graph CRUD and typed allocation failures

The canonical `litchi-pptx::tag` owner now carries direct presentation-root and
common-slide-data programmable tags through the complete package graph instead
of stopping at detached XML. Bounded singleton `load`, `put`, and `remove`
operations move an owned `List`, derive or preserve Strict/Transitional
dialect, validate candidate XML before commit, allocate relationship and part
names with finite collision scans, and retain existing noncanonical part
names. The direct `p:custDataLst/p:tags` anchor, its internal relationship, and
its tag-list part change as one transaction while existing `p:custData`
children remain intact and schema order is preserved.

Owner preflight fails closed on duplicate or out-of-order common-slide-data
children, on customer data placed after `p:tags`, and on any mismatch among
the owner's Strict/Transitional namespace, the anchor's relationship namespace,
the relationship type, and the tag-list part namespace. These failures occur
before package mutation and preserve signatures and shared blobs.

Replacement is byte-aware: an identical result preserves the original part
allocation and package signatures. A shared target is forked and only the
selected anchor is retargeted. Removal deletes a target only after a bounded
package-wide inbound scan proves it orphaned, and repeated removal safely
returns `None`. The high-level PPTX package hides raw identities behind short
`tags`, `put_tags`, and `remove_tags` methods, with exact slide name as the
common selector and checked slide position as the secondary path. It rejects a
dirty legacy presentation writer because later materialization could replace
the edited slide markup and relationships. Shape-level `nvPr` anchors remain
distinct and are deliberately not flattened into slide results; typed shape
authoring is pending, while the explicitly low-level relationship inventory
continues to expose existing producer content.

A native PowerPoint experiment caught the reason for this boundary. A
relationship-only Litchi deck opened without a repair dialog, but PowerPoint
silently removed the dangling relationship and tag part when the deck was
edited and saved. In contrast, PowerPoint edited and resaved LibreOffice's
`tdf103477.pptx` without removing any of its seven shape-anchored slide tag
lists, and Litchi read all seven afterward.

The corrected direct common-slide-data authoring then passed the same native
macOS PowerPoint gate. PowerPoint opened
`target/office-verification/pptx-tag-crud-anchored-generated.pptx` without a
repair prompt, accepted a visible title edit, saved, closed, and reopened the
deck cleanly. The PowerPoint-saved copy is
`target/office-verification/pptx-tag-crud-anchored-powerpoint.pptx`; Litchi
reverse-read both the edited title and
`LITCHI_VERIFY=pptx-tag-crud-v1`. This verifies direct `p:cSld` tag persistence
for that desktop build and workflow. It does not certify typed shape-tag
authoring, other Office platforms or versions, or performance.

The same safety pass replaces stringified capacity failures throughout the
canonical XLSX owner with a typed `Allocation { resource, source }` error. All
95 production fallible reservation sites now retain `TryReserveError`; the
migration OOXML and XLSB seams preserve that source as well. The shared
`litchi-core::Error` now carries the same typed allocation failure, so OOXML
and CFB conversions no longer erase the allocator source at the umbrella API.
This improves diagnostics without changing mutation ordering or making an
unmeasured performance claim.

Focused evidence is recorded after the corrected anchor implementation passes
its canonical and facade tests, warning-denied Clippy, rustdoc, formatting, and
diff validation. The previously green full-workspace gate is not repeated by
explicit user direction.

## Borrowed PPTX shape scenes and shape-owned tag CRUD

This PPTX slice replaces the migration host's copied
`BaseShape + ShapeType + Vec<u8>` read model with the canonical
`litchi-pptx::shape::{Scene, Shape}` owner. `Shape` is a non-exhaustive,
data-bearing enum; common semantic accessors are implemented once on the enum
and each contextual variant lends the same indexed record. After bounded MCE
preprocessing, one namespace-aware indexing scan records shapes in depth-first
pre-order, including nested group children, while `Group::shapes` retains the
direct-child hierarchy. The shape-tree root and a graphic-frame OLE preview
picture are not exposed as user shapes.

Exact decoded `cNvPr` names are the primary selector. Checked numeric pre-order
positions remain available for duplicate-name repair and source-order tooling.
Ordinary `get` lookup represents a missing name as `None`; strict `shape`
lookup, duplicate exact names, and out-of-range positions use explicit
`LookupError` variants. The facade does not implement panicking indexing or
require relationship IDs. Strict and Transitional namespaces are resolved in
the complete owner context, so inherited aliases do not disappear when a
shape is viewed.

MCE processing is bounded by owner bytes, output bytes, nesting, nodes, shape
count, and retained decoded text. An owner without an MCE rewrite remains
borrowed. When Choice/Fallback processing produces replacement markup, the
scene owns one processed owner buffer; every shape XML value is still a checked
span into that shared buffer. This removes per-shape subtree copies without
claiming allocation-free parsing or unmeasured speedups.

Shape-owned tag CRUD is publicly wired as
`litchi-pptx::tag::shape::{load, put, remove}` and reuses the canonical
`shape::Key`. It maps the processed-scene selection back to the corresponding
active raw-source shape span, then patches only that shape's `nvPr` anchor.
Inactive MCE branches, customer-data siblings, comments, unrelated attributes,
and schema order remain intact. The scanner covers `sp`, `pic`, `cxnSp`,
`graphicFrame`, and `grpSp`, including children of bounded nested groups, while
excluding the synthetic shape-tree root and nested OLE preview pictures.

The OOXML migration facade exposes concise package `shape_tags`,
`put_shape_tags`, and `remove_shape_tags` operations. They compose the existing
name-first/checked-position slide selector with the canonical name-first or
checked-depth-first shape selector, so ordinary callers never handle native
shape IDs, relationship IDs, or part names. An already-resolved
`Slide::shape_tags` performs the read directly against its package-backed owner.

The graph mutation is staged before commit. Byte-identical puts preserve the
part allocation and signatures; a relationship ID shared by active anchors or
a target reached by another package edge is forked for replacement; removal
retains shared edges and collects a target only when it is orphaned. Typed
failures leave the owner XML, relationships, parts, signatures, and shared
blobs unchanged, and owned `List` values move into successful writes.

Direct-owner `load`, `put`, and `remove` now apply one bounded MCE branch
selection policy. Mutation maps the processed owner root, insertion point,
direct `p:custDataLst`, and direct `p:tags` back to the corresponding active
raw-source elements before constructing a patch. It can therefore create an
anchor in an active missing or empty container, update an active anchor, and
remove it without editing a processed offset or touching an inactive
Choice/Fallback branch. Every preserved raw anchor participates in shared-ID
use counts, so an inactive branch that reuses the selected relationship forces
replacement to fork and removal to retain the old relationship and target.
Owner and tag-list preprocessing share the PowerPoint capability profile:
baseline OOXML plus the checked-in `p14` and `p15` extension namespaces, with
first-supported-Choice semantics. The processed and raw element sequences,
namespace profile, anchor ID, and staged post-edit semantic layout must agree
before any package graph mutation; divergence is a typed validation failure.
The temporary `MceOwnerMutation` boundary is removed rather than retained as a
compatibility alias.

Focused canonical and facade tests are green. They cover all seven direct
owner families, the five shape families and nested groups, both namespace
profiles, `p14` Choice and Fallback selection, missing and empty active
containers, presentation/common-slide-data schema order, raw-source MCE
mapping, byte preservation, atomic/no-op behavior, and shared-anchor/target
fork, retention, and collection. Warning-denied Clippy also passes for
`litchi-pptx` and the focused migration-host integration target. For native
direct-owner evidence, PowerPoint for macOS 16.110.2 opened the generated
`target/native-office/pptx-mce-choice-updated.pptx` p14-Choice deck without a
repair prompt, rendered its slide, saved a normalized copy, closed, and
reopened that copy without repair. The public facade reverse-read
`LITCHI_MCE=updated`. PowerPoint flattened the selected Choice into a direct
`p:custDataLst` and removed the inactive Fallback relationship and part during
its own resave; that normalization is recorded as native application behavior,
not as a Litchi source-preservation claim. The independently generated
removed-tag artifact also opened without repair and reverse-read with no active
tag list.

For shape-owned native evidence,
PowerPoint for macOS 16.110.2 opened the Litchi-mutated LibreOffice fixture
without a repair prompt, saved a normalized `.pptx` copy, closed it, and
reopened that copy without repair. A reverse read through the public facade
found and replaced the pre-existing `LitchiNativeCheck` tag on shape `Objekt 2`.
This certifies that specific shape-tag create/read/update and native-resave
path on the tested desktop build; it does not certify every owner family,
Office build, or unsupported extension. Per explicit user direction, this
slice uses focused gates and does not repeat the previously green
full-workspace run.

## OOXML physical-package ownership

The OOXML migration host no longer depends directly on `soapberry-zip`.
Embedded chart workbooks now write canonical `PackURI` members through
`litchi-opc::phys_pkg::PhysPkgWriter`; the one malformed-fixture filter uses
the matching OPC reader/writer boundary without naming or exposing the archive
implementation. The unused string-only XLSB ZIP error variant and its concrete
error conversion are deleted. Generated embedded workbooks are reopened as
logical `OpcPackage` values in both example and property tests, which proves
required part discovery rather than only a ZIP magic prefix.

Focused embedded-workbook and malformed-fixture regressions pass, as do
warning-denied all-target Clippy for the migration host, warning-denied rustdoc,
format and manifest checks, and the executable dependency policy. The boundary
checker now accepts 35 workspace packages and 107 direct internal dependencies
with 14 explicit debt items. Per explicit user direction, the previously green
full-workspace test run is not repeated for this ownership-only slice.

## Checked BIFF8 formula references

The legacy Excel writer no longer represents public cell and area tokens as
raw `u16` coordinate tuples. `writer::formula::{Ref, Area}` keep their fields
private, use `u16` for the exact zero-based BIFF8 row domain and `u8` for the
exact `A..=IV` column domain, and expose concise checked construction and
accessors. Every `Ptg` variant now uses its short contextual name without a
compatibility alias; reference-bearing variants carry the checked values.

The byte-oriented A1 scanner avoids the former temporary column string, uses
checked arithmetic, accepts the last cell `IV65536`, and returns
`InvalidCellReference` for `IW`, oversized columns, malformed rows, and
arithmetic overflow. Production `unwrap` and `expect` calls were removed from
the tokenizer path. Defined names preserve
their established reversed-corner normalization while formula cells,
conditional formats, data validation, and names share the same token boundary.

All 42 focused formula tests pass, including exact relative-flag bytes and
adversarial no-unwind cases. Public writer serialization returns the same typed
error without unwinding, and the defined-name regression passes. Warning-denied
Clippy and rustdoc, formatting, diff validation, and workspace lint are green.
No parser-throughput or allocation improvement is inferred from these safety
tests, and the previously green full-workspace suite is not repeated.

## Atomic PowerPoint speaker-note deletion

The opened-package facade now provides exact-name-first `remove_notes`, a
checked numeric-position fallback, and all-slide `clear_notes`; the mutable
authoring slide has an idempotent `clear_notes`. Relationship IDs and part
names stay below the common path. Dirty legacy-writer state and invalid or
ambiguous selectors fail before package mutation.

Deletion indexes and validates the complete Strict or Transitional notes graph
without copying notes, master, or theme payloads. It records actual stored OPC
keys, rejects orphan, duplicate, malformed, or unexpectedly shared notes
slides, and scans all inbound package edges. Slide owners are staged with
shared built-in payload allocations and edited relationship collections before
commit; only then are they replaced, exact notes parts removed, and signatures
invalidated. Notes masters and themes remain available.

Seven graph tests, one mutable-slide test, and four saved/reopened package tests
pass. They cover semantic selection, idempotence, malformed and shared-edge
atomicity, case-folded storage, retained shared infrastructure, and
byte-identical slide XML. Warning-denied Clippy and rustdoc, formatting, diff
validation, and workspace lint are green. Native PowerPoint open/resave and
representative memory profiles remain separate evidence; no such compatibility
or performance claim is made by this slice.

## Core-properties reader ownership

At the ADR 0014 boundary, the host-neutral OPC core-properties reader moved
from `litchi-ooxml` to the existing `litchi_ooxml_common::properties` owner.
The only public read entry was the contextual
`properties::read(&OpcPackage)`; the migration-host module was deleted without
a forwarding alias, and the document, presentation, XLSX, and XLSB facade
adapters called the common owner directly. ADR 0015 later adds the common
write and clear operations and routes DOCX, PPTX, and XLSX through retained
host caches.

That reader kept relationship-selected lookup, Transitional and Strict
namespaces, OPC M4 restrictions, content-type checks, datetime normalization,
entity decoding, and bounded retained text behind structured common errors.
ADR 0015 subsequently replaces normalization with lossless schema-typed
lexical values.
Both production `expect` paths were replaced by typed invalid-data results.
This changes no package bytes and does not claim that the remaining
umbrella-to-host dependency debt is resolved.

All 13 focused property tests pass, including Apache POI conformance fixtures
and a typed text-budget regression. Warning-denied Clippy and rustdoc pass for
the common owner, host, and isolated facade; workspace lint and the executable
boundary checks are green at 35 packages, 107 direct internal dependencies,
and 14 explicit debt items. Per explicit user direction, the full-workspace
test suite is not repeated.

## Lossless OOXML core-properties CRUD

The common owner now exposes the concise `Props` value and explicit `read`,
consuming `write`, and idempotent `clear` package operations. DOCX, PPTX, and
XLSX retain absence and the already validated value in a hidden dirty-tracked
slot; their public core-properties surface is limited to `props`, `props_mut`,
`put_props`, and `clear_props`. Untouched saves preserve exact core-property
bytes and signatures. Non-destructive updates retain noncanonical targets,
relationship IDs, Strict or Transitional dialects, and legal outbound
extension relationships. Destructive clear rejects shared inbound ownership,
then removes the core owner edge and part while leaving extension parts intact.

The schema-faithful semantic model follows the normative OPC schema rather
than historical facade assumptions: revision remains an arbitrary string,
`created` and `modified` retain every W3CDTF precision and optional timezone,
`lastPrinted` retains `xsd:dateTime`, and keyword text and language-bearing
`cp:value` children remain ordered mixed content. The non-schema
`cp:contentType` element is not part of the schema model. Text, lexical values,
package cardinality, and byte budgets fail through typed errors without
unwinding.

Focused common and host CRUD, malformed-input, no-op, graph-preservation, and
failure-atomicity tests pass. The `core_props_office` example generated the six
artifacts under `target/office-core-props`; its reproducible command is
`cargo +1.89 run -p litchi-ooxml --example core_props_office --all-features`.
The exact historical invocation and application versions were not recorded.
Through Computer Use in Microsoft Word, PowerPoint, and Excel on macOS, all
six opened without repair prompts, displayed the expected metadata before
clear and blank values afterward, and rendered the document, slide, and
worksheet content. This supports open-and-inspect compatibility for those
artifacts and tested desktop applications only; Office-side edit/resave and
reverse-read were not performed for this slice. At this historical slice, raw
mutable OPC access could still make the host slot stale, and a later host-save
failure could leave the in-memory package changed after a successful slot
flush. The later failure-atomic PPTX publication section supersedes this state
for PPTX; DOCX and the separately described XLSX seams remain open. Per user
direction, the previously green full-workspace suite is not repeated.

## Checked BIFF8 writer locations beyond ordinary cells

The legacy Excel writer now checks merges, validations, data-table inputs,
filters, sort keys, pivots, Web publications, RTD cells, and page breaks before
retaining or mutating location state. Dedicated values migrated in this slice
use the narrowest representation for the exact 65,536-by-256 BIFF8 cell grid,
while horizontal page-break column spans retain their distinct `[MS-XLS]`
range through 16,383. Inclusive public ranges have private fields, fallible
constructors, and short accessors. WebPub insertion validates its range,
source, and strings; RTD insertion validates subscriber coordinates and the
workbook-relative sheet in context.

The touched operations validate all fallible location inputs before changing
worksheet collections or defined names. Page-break ordering, overlap, and
count limits follow the relevant record sections, and validation serialization
consumes a prechecked range instead of inserting provisionally and recovering
with `unwrap` or `pop`. Three adversarial unit tests and 38 selected integration
tests pass, including exact maxima, overflow, reversal, overlap, count limits,
serialization/reopen, typed rejection, and no-unwind failure atomicity.
Warning-denied Clippy and rustdoc, formatting, and diff validation are green
for `litchi-xls`; no unmeasured performance or native Office claim is made.
AutoFilter ranges, sort keys, pivot locations, and data-table anchors retain
some legacy wide private storage after the checked boundary. `XlsSortData`
Rw12/Col12 policy and save-time RTD topic encoding also remain open.

## Format-owned deterministic OOXML producer templates

DOCX, PPTX, XLSX, and XLSB now own the exact minified assets they embed under
format-local `resources/generated` directories. Readable XML sources remain
beside them, and an `xml-minifier` integration test regenerates every mapped
asset and requires byte equality. Production `litchi-ooxml` no longer invokes
or depends on the development minifier, removing that executable dependency
and its boundary-debt entry.

Fresh core templates are present but semantically empty; they do not fabricate
an author, revision, or clock history. XLSB creation no longer reads ambient
time or assumes timestamp formatting is infallible. All 39 assets pass exact
parity, focused minifier and XLSB regressions pass, and warning-denied Clippy,
rustdoc, formatting, manifest, dependency, and executable-boundary checks are
green. The boundary policy records 35 packages, 106 direct internal
dependencies, and 13 explicit migration debts. Determinism and ownership do
not imply unmeasured speedups or universal native Office compatibility.

## Typed XLSX calculation-chain ownership

The calculation-chain implementation now lives exclusively in
`litchi-xlsx::chain`. Its short `Sheet`, `Step`, `Flags`, `Cell`, and `Chain`
model makes the native sheet-ID range, `l`/`s` exclusion, packed orthogonal
markers, checked grid addresses, and nonempty ordering explicit. Semantic
sheet/address lookup and CRUD are primary; checked calculation-order operations
remain available, and duplicate malformed producer keys are retained for
numeric repair while semantic selection reports ambiguity.

The bounded Strict/Transitional reader applies MCE, preserves extension XML and
qualified attributes, and does not evaluate formulas. Package `load`, `put`,
and `remove` reject incoherent, external, duplicate, orphaned, or shared graph
states before mutation. Exact stores preserve signatures; changed stores keep
the part and relationship conformance aligned, and removal retains a part that
another relationship still references. The migration host now caches the
canonical owner and exposes only `chain`, `chain_conformance`, consuming
`put_chain`, and `remove_chain`; the former host module and aliases are gone.

Nine owner tests and two host integration tests pass, together with warning-
denied Clippy and rustdoc. The executable boundary checker remains green at 35
packages, 106 direct internal dependencies, and 13 explicit debts. This is a
typed ownership and byte-preservation slice. The legacy host's larger save
transaction, representative performance measurements, and new native Office
evidence remain separate work.

## PPTX speaker-notes owner extraction

`litchi-pptx::notes` now owns the bounded notes graph, XML codec, package
service, text producer, and both source and generated notes-master assets. Its
public model uses `Conformance`, `Theme`, `Master`, `Slide`, and `Graph`; OPC
identities remain private. `load` produces an independently editable graph with
one bounded copy of each validated payload, focused `slide` copies only the
selected notes resource, and metadata-only delete operations copy none.
Presentation, slideshow, and template main parts are accepted in both macro-
free and macro-enabled families.

Consuming `put` stages the complete edit before commit and moves the caller's
buffers into OPC parts. Exact no-ops retain signatures. The text producer has
an explicit Strict/Transitional path, while fresh-package master generation is
the Transitional profile already used by the legacy authoring flow.

The OOXML host now keeps only semantic slide selection and dirty-writer guards
around graph reads and mutation. The former notes module, forwarding names,
template accessor, and duplicate slide XML writer are deleted. Focused owner,
host CRUD/graph, and minified-asset parity tests pass with warning-denied Clippy
and rustdoc. The `pptx_with_fonts` example produced a six-slide Transitional
artifact. Through Computer Use, desktop PowerPoint for macOS opened it without
repair, marked slide 1 as having notes, and displayed the expected speaker-note
text in the Notes pane. This is open-and-inspect evidence only: no Office edit,
resave, Strict master/theme synthesis, application-version matrix, or measured-
performance claim follows. The lifetime-free graph's one-copy read cost and
the remaining package-host adapter are explicit migration debt.

## Failure-atomic BIFF8 worksheet views

The XLS writer replaces public field bags with checked
`view::{Scale, Mode, Pane, Selection, View}` values. Frozen and split panes
encode their distinct cell-count and twip limits, selections validate pane
existence, grouping, range order, active index, and active-cell containment,
and a view rejects invalid origins, palette indices, or zooms before
publication. Nine display switches occupy a private BIFF-aligned `u16`; pane
group validation uses a non-allocating `u8` mask.

`put_scale`, `put_view`, and `put_pane` move new state in and return the old
owned state only after whole-state preflight. The old option names and raw
setters are removed, and `XlsSelectionRange::new` is now fallible. Six view
unit tests, two typed writer round trips, five workbook-view tests, and two
lint regressions pass. Warning-denied all-target/all-feature Clippy and rustdoc
are green for `litchi-xls`. The `xls_styles_example` artifact opened in desktop
Excel for macOS without repair, in expected BIFF8 Compatibility Mode. After
jumping to `M30`, Excel kept row 1 and column A visible, confirming its
interpretation of the frozen-pane records. This does not cover Office resave,
all view combinations, or an application-version matrix. The compact layout is
a structural result; cache, allocation, and latency claims still require
measurement.

## DOCX web-settings owner extraction

`litchi-docx::web` now owns bounded Strict/Transitional web-settings parsing,
semantic division and frameset CRUD, deterministic default bytes, and the
optional package graph. Producer-visible division IDs are the ordinary
selector; checked numeric source positions remain available for repair.
Missing lookup is `None`, while ambiguity, invalid positions, malformed XML,
mixed dialects, external/shared edges, and resource exhaustion are typed
errors. Package writes consume the completed `Settings`; unchanged and
semantic no-op stores retain exact producer bytes and signatures.

The migration host now keeps only the wider DOCX adapter and exposes `web`,
`put_web`, and `remove_web`. Its former parser, cache, writer, template
accessor, and duplicate assets are deleted. The complete owner gate passes 43
unit tests, two public API tests, and one doctest; focused host gates pass two
owner integrations, seven web-settings package regressions, and four shared-
color/underline regressions. Asset parity, warning-denied Clippy and rustdoc,
formatting, stale-name, panic-name, and boundary checks pass. Semantic edits
canonicalize the modeled XML and can drop unmodeled extensions; high-level
frame-target CRUD and exhaustive border-style semantics remain follow-up work.

The native Word gate caught two interoperability defects before this slice was
accepted. Empty true `bodyDiv` and `blockQuote` elements are schema-valid but
desktop Word rejected them; the owner now writes explicit numeric values while
retaining permissive lexical reads. A diagnostic matrix also proved that a
borderless plain division opens cleanly, so the API does not turn the producer
corpus's all-bordered convention into a false schema requirement. Separately,
the host heading helper confused `Heading 1`'s display name with style ID
`Heading1` and omitted default Heading4 through Heading9 definitions. Exact-ID
plus typed `Outline::{H1, ..., H9}` wire levels and save/reopen catalog
regressions cover the correction. The final
`owner_native_smoke` document contains scalar settings plus body and block-quote
divisions; Word for macOS opened it without repair, rendered its Heading 1 and
body text, and identified the heading with the native Heading 1 style. No
Office edit/resave, Strict native check, version matrix, extension-preserving
modeled edit, or measured-performance claim follows.

## PPTX table-style owner extraction

`litchi-pptx::table::style` now owns the bounded catalog model, exact source
retention, deterministic default bytes, and optional presentation graph. The
allocation-free GUID `Id` is the stable selector, checked `at` supports source-
order repair, and `named` returns all duplicate display-name matches. `Parts`
encodes the schema regions compactly and in normative sequence; detailed
formatting remains bounded opaque XML. Unchanged stores move the original
list-owned allocation back to OPC, while rename preserves a definition's
opaque body and `reset_parts` is explicitly destructive.

All six presentation/slideshow/template profiles are accepted in both macro
families and both XML dialects. Graph mutation rejects mixed, external, orphan,
shared, relationship-bearing, or wrong-content-type states before commit. The
host exposes only `styles`, `put_styles`, and `remove_styles`; legacy
Transitional slide materialization preserves an optional catalog edge's exact
ID/type/target and propagates the generated master relationship ID. Strict
legacy materialization is refused before mutation until that writer becomes
dialect-aware. Semantic no-op comparison preserves inherited `xml:space` and
all text in deeper opaque payloads, preventing whitespace-bearing extension
edits from being silently discarded. The owner passes 71 unit tests, three
doctests, and one compile-fail test; the focused host passes all eight
integration tests and producer-asset parity passes. Warning-denied Clippy and
rustdoc plus targeted formatting, diff, stale-name, and boundary gates pass.
The `owner_native_smoke` example
creates a typed definition used by a two-row table and verifies it after reopen.
Desktop PowerPoint for macOS opened that Transitional artifact without repair,
rendered the table and text, and exposed native Table Design and Table Layout
tabs on selection. No Office edit/resave, Strict native check, version matrix,
detailed-style rendering, or measured-performance claim follows.

## Checked BIFF8 shape anchors and SortData

The XLS writer now exposes `writer::shape::{Point, Anchor, Behavior, Rect}` and
`writer::sort::{Row, Col, Range, Axis, Method, Parent, Dxf, IconSet, Icon, On,
Key, Config}` instead of public raw field bags. The types prove grid, offset,
ordering, group-rectangle, anchor-flag, packed Rw12/Col12, and axis/key
invariants before retained state changes. Shape/group/comment insertion
reserves storage before mutation and object-ID assignment; SortData uses
move-returning `put_sort`, idempotent `remove_sort`, and borrowed `sort`.

Thirty-nine focused tests cover exact bounds, wire round trips, failure
atomicity, allocation ordering, explicit/automatic ID collisions, and related
list-object writers. Strict Clippy, warning-denied rustdoc, formatting, diff,
and no-run target compilation pass. The `odraw_native_smoke` example generated
an XLS with a rectangle, text box, and grouped ellipse/text pair. Through the
Computer Use skill, desktop Microsoft Excel for macOS opened it without repair
in Compatibility Mode and rendered the expected objects, fills, text, and
placement. No Excel edit/resave, SortData UI exercise, version matrix, or
performance measurement was performed.

## Failure-atomic XLSX late publication

The common core-properties `Slot` now stages a candidate mutation through a
lifetime-bound guard. Only consuming that guard after successful publication
can clear its exact originating slot; dropping it retains dirty intent for a
retry. The XLSX host avoids a package snapshot on unchanged opened saves. When
late metadata or worksheet overlays require one, it structurally snapshots the
graph while sharing built-in immutable part payloads, restores the original
worksheet owners after the sink, and commits the property guard only after all
late work succeeds. Metadata-only saves preserve producer application-
properties bytes instead of regenerating them.

Three focused host regressions prove the unchanged fast path, injected sink-
failure restoration with a successful retry, and invalid-property rejection
before the sink. The common guard regression proves drop-retains/commit-clears
semantics. At this slice, the earlier writer-model materialization phase,
custom `Part` clone policies, and DOCX/PPTX save transactions remained explicit
follow-up work. The later PPTX publication section closes its staged save
transaction; DOCX and XLSX's earlier materialization seam remain open. This is
correctness evidence, not an allocation or latency result.

Per explicit direction, no redundant manual full-workspace gate was scheduled.
The repository's mandatory pre-commit hook nevertheless ran
`cargo test --workspace --all-features --lib --tests` and the corresponding
workspace doctests, and both passed. That integration gate also closed stale
contracts exposed by the stricter owners: the umbrella table test now expects
a typed orphan-core rejection for its malformed relationship fixture; the
XLSB calculation owner accepts the exact one-byte option tail emitted by a
checked-in Microsoft Excel 12 artifact while writing only canonical 26-byte
records; a section-mutation fixture now supplies a conforming MCE Choice before
Fallback; and the placeholder regression expects its decoding error through
the canonical PPTX owner.

## DOCX glossary owner extraction

The glossary-document grammar, semantic building-block catalog, and complete
owned OPC graph now reside in `litchi-docx::glossary`. The OOXML migration host
exposes short package/document adapters plus the canonical module as a
contextual re-export, and deletes its duplicate implementation and legacy
aliases.
Name-first Unicode-caseless CRUD has a private normalized-name index, checked
numeric replacement/rename/removal/reorder repair selectors, and typed
ambiguity/out-of-bounds failures. Checked per-entry and catalog output sizes
update by selected-entry deltas instead of replanning unrelated bodies. Fresh
authoring requires a checked name while the reader accepts base-standard
producer states where `docPartPr` and its name may be absent and `<guid>` may be
present without `guid/@w:val`. It also retains Word's native
present-but-empty `<w:types/>` state without making it authorable through fresh
typed values. Untouched direct producer entries retain bounded serialized
inactive or ignorable MCE subtrees and all of their relationship references
across unrelated CRUD. XML 1.0 values,
content events, producer snapshots, projected opaque XML, and owned DOM
allocations have aggregate budgets; the streaming namespace resolver avoids a
depth-times-binding scan. Carriage returns remain distinct from line feeds.
Strict reads reject mixed relationship namespaces, Transitional-only on/off
lexicals, and active VML, while Transitional VML remains readable and writable
only in its own dialect.

The raw layer borrows its graph for failure-safe publication and shares
auxiliary payload owners. Repeated publication of an edited loaded graph is a
true signature-preserving no-op after canonicalization. Real changes unsign the
candidate before staging package mutations; signature infrastructure, OPC
manifests, and relationship-part names are forbidden graph payloads. Owner edges
rebase to a different destination main-part base and allocate a fresh ID on
collision.
Destination-bound semantic catalogs return before reserialization when their
catalog is unchanged. Relationship-bearing entries and backgrounds also carry
per-value lineage, so transplanting one entry cannot silently reuse a colliding
destination `r:id`. Shared auxiliary payloads use pointer identity before byte
comparison, and reference validation borrows cached IDs without cloning them.
Per-part and graph-wide payload, relationship, and metadata budgets; validated
MIME types and XML IDs; fallible reservations; and keyed linear graph comparison
address the reviewed raw-graph algorithmic-scaling hazards. Persistent namespace
frames and node/attribute/content/depth/namespace/owned-allocation budgets bound
the glossary XML codec and opaque semantic subtrees; extracted roots share an
aggregate output ceiling and inert auxiliary bytes are not structurally decoded.
Raw remove returns a graph that raw put can restore within the same published
limits.

The standards matrix is driven by the role assigned by an incoming edge rather
than by a payload's self-selected content type. It covers the normative root,
rich-story, settings, font-table, numbering, web-settings,
chart/chart-user-shapes, Custom XML data/properties, ActiveX descriptor/binary,
SmartArt, embedded object/package, and Microsoft customization edges in both
conformance families. This includes generic inert controls, settings recipient
data, chart theme overrides, 2011 and 2012 chart styles, the chart/chart-drawing
cycle, chart-drawing Custom XML, and diagram hyperlinks. Only exact ActiveX
descriptors may own ActiveX binaries. Target modes and content profiles are
checked independently. Internal hyperlinks are validated references and are
neither moved into nor deleted with the glossary graph. Conventional
`/word/glossary/` paths do not establish ownership without a typed relationship
edge. The Transitional relationship spelling is the normative case-sensitive
`aFChunk`; direct themes are rejected because the Word glossary part is not a
permitted theme owner.

Fresh semantic publication seeds canonical-first, collision-free glossary-local
styles, settings, font-table, and web-settings parts, sharing an existing
self-contained main-document blob where possible. `Package::new_template()`
selects a DOTX main content type.
Focused verification covers the owner unit suite, the host integration suite,
warning-denied Clippy, rustdoc, formatting, manifest sorting, and executable
crate-boundary checks; this slice intentionally does not schedule another
manual full-workspace run after the previously green workspace gate.

Computer Use produced both negative and positive native evidence in Microsoft
Word for macOS. The first DOCX, which lacked the later four-resource seed,
opened without repair in Compatibility Mode but did not expose the custom
AutoText entry; Word removed the glossary graph on resave. This is consistent
with templates being the native AutoText authoring container, but the subsequent
seeded DOTX also changed the resource graph and is not a controlled causal
comparison. The seeded DOTX opened without repair. Insert → AutoText showed the
exact `Litchi AutoText` row and enabled insertion; inserting it placed `Litchi
reusable native building block` into the document. Word saved the template to
the Mac.
The saved archive passed ZIP integrity, and the public example
reverse-read the entry and payload from the Word-saved copy through the
canonical owner, including Word's native `<w:types/>` rewrite. This is one
Transitional text-only AutoText
open/discover/insert/resave result in the observed Word Compatibility Mode on
that build, not evidence for images, fields, arbitrary dependency graphs,
Strict, a version matrix, or performance.

## PPTX embedded-font owner extraction

The PresentationML embedded-font grammar, semantic model, and package graph now
reside in `litchi-pptx::font`. The migration host deletes its duplicate module
and long raw-field exports, and exposes short `fonts`, `put_fonts`, and
`remove_fonts` adapters. `Font`, `Face`, typed styles, closed pitch/family
values, typed signed-byte charset, fixed-size PANOSE, compact licensing
metadata, and consuming collection CRUD replace public relationship IDs, part
paths, MIME strings, and optional resource bags. A documented library
Unicode-caseless key powers semantic selection; checked positions remain usable
for producer-duplicate repair.

Loaded payloads use shared immutable allocations, including repeated rIds and
targets, and unique-resource limits no longer multiply one physical program by
its number of face references. Fresh PowerPoint authoring validates or creates
an Embedded OpenType container under `application/x-fontdata`, sets
`embedTrueTypeFonts`, and
allocates canonical collision-free names. Standards-only `x-font-ttf` is an
explicit preservation profile. PPTX no longer exposes or emits Word-only font
obfuscation. The optional discovery/subsetting path is split into a pure owned
preparation phase plus DOCX/PPTX publishers, surfaces `OS/2.fsType`, applies
licensing and no-subsetting policy, and performs one bulk typed PPTX put. The
collector exposes scalar-only `Glyphs` keyed by typed family/style `Request`,
so automatic Bold/Italic faces are no longer mislabeled Regular. The OPC save
policy is `FontEmbedding::{None, Full, Subset}` instead of two booleans. Word's
obfuscation GUID is `FontKey([u8; 16])`, and raw strings exist only at the XML
codec boundary. Incomplete opened-document scans and unsupported full TTC face
extraction fail explicitly rather than silently publishing incomplete data.

Exact no-ops return before serialization and retain signatures. Real changes
validate and round-trip a candidate, then mutate and unsign only an Arc-sharing
OPC snapshot before final assignment. Unit and downstream API tests cover
semantic CRUD, shared allocations, collisions, unknown XML, failure atomicity,
Strict/Transitional reference packages, graph rejection, and the host adapter.
Feature-on and feature-off compilation cover the split authoring path.

Computer Use verified the generated Transitional artifact
`target/office-verification/pptx-font-crud-generated.pptx` in desktop
PowerPoint for macOS. It opened without a repair, recovery, compatibility, or
font-license warning; rendered the visible Boldonse `Test`; and reported
`Boldonse` in the font control when the text box was selected. The text was
changed to `Test Test`, saved as
`pptx-font-crud-powerpoint.pptx`, closed, and reopened without repair. The saved
ZIP passed integrity checks. The canonical reverse reader recovered the edited
text, the one-face Boldonse catalog, its presentation relationship, and its
inert EOT resource. The observed Office copy retained the exact 36,187-byte
payload and its original SHA-256, but the reusable verifier accepts a
structurally valid Office normalization rather than requiring byte identity.
This is evidence for that producer EOT artifact and desktop build only; it is
not a native gate for the automatic uncompressed-EOT wrapper, Strict,
`x-font-ttf`, other Office versions, or measured performance.

## Typed chart domains and native Excel compatibility

The DrawingML bubble-chart facade now uses the closed `bubble::Size` enum and
the checked `bubble::Scale` scalar instead of an open string and an unbounded
integer. SpreadsheetML borders likewise use the focused
`xlsx::styles::border::{Line, Rgb, Tint, Color, Side, Dir, Diagonal, Border}`
model throughout the parser, worksheet facade, cell format, and writer.
Unknown line-style and malformed ARGB tokens fail parsing, tint is bounded,
absence is represented by `Option<Side>`, and authored diagonals require both
a side and direction. On visible sides, theme, indexed, RGB, automatic, and
tint-only color states survive read/write; style-less or `none` sides
canonicalize to absence. Inside-edge, outline, and Strict logical-edge states
also survive read/write. Border resources compare complete values, and cell
formats key their resolved resource identities rather than trusting a hash
alone.

Computer Use first opened
`target/office-verification/typed-chart-domains.xlsx` in desktop Microsoft
Excel for macOS and observed a repair report that removed
`xl/drawings/drawing3.xml`. Quoting the worksheet name in the chart formulas
was necessary but did not remove the repair. Inspection of the checked-in
`[MS-OE376]` conformance notes identified section 2.1.1458(b): ISO/IEC 29500
permits `bubble3D` directly under `bubbleChart`, but Microsoft Office does not.
The reader now accepts that standards form and projects its value onto each
typed series; the writer emits only the Office-compatible series-level form.

After regeneration, the same artifact opened with no repair, recovery, or
compatibility warning. Excel's accessibility tree exposed both the scatter
chart and the bubble chart on the `Scatter & Bubble` sheet, and visual
inspection confirmed both rendered. This is native-open evidence for that
generated workbook and the tested desktop Excel build only.

A separate `xlsx_comprehensive_features.xlsx` artifact exercises the typed
border facade by authoring a black `Line::Thick` bottom `Side` on each header
cell. The same desktop Excel session opened it without a repair, recovery, or
compatibility warning. With `A1` selected, Excel's native **Format Cells >
Border** inspector reported `Thick` as the selected line style and rendered a
black bottom edge with the other edges absent. This verifies native
interpretation of that one RGB bottom-side combination. Neither artifact is an
Office-resave fidelity test, a version matrix, coverage of the remaining line,
color, logical-edge, diagonal, or chart variants, or measured performance.

## Typed fixed-domain continuation and native Office evidence

The next fixed-domain slice removes four more stringly or weakly bounded
facades. SpreadsheetML alignment now uses
`alignment::{Horizontal, Vertical, Rotation, Reading, Indent, Alignment}` from
parser through shared `CellFormat`, resolved worksheet reads, exact XF
deduplication, and writing. PresentationML modern-comment `complete` is the
checked, niche-encoded `Progress` value rather than `Option<String>`.
WordprocessingML section numbering, note placement, page-border color, and art
styles use `ChapterSep`, separate `FootnotePos`/`EndnotePos` domains,
`BorderColor`, and `PageBorderArt`. DrawingML diagram identifiers and relation
kinds use `Id`, `PointType`, and `ConnectionType`, with concise semantic CRUD,
Office's shared one-parent rule for `parOf` and `presParOf`, cascading removal,
and a preflighted canonical writer. The diagram codec explicitly does not claim
lossless round-tripping of rich XML outside its modeled subset.

Focused Rust evidence covers all DrawingML targets, DOCX section parsing and
writing, XLSX style parsing/writing and exact deduplication, PPTX modern-comment
unit and package CRUD tests, and compilation of the affected public examples.
Malformed fixed tokens, out-of-range percentages and alignment scalars,
cross-domain note positions, invalid identifiers, graph conflicts, XML control
characters, and aggregate diagram-output limits are rejected before
publication. These are correctness and API-shape gates, not allocation,
latency, throughput, or cache measurements.

Computer Use opened the generated
`target/office-verification/typed-alignment/xlsx_comprehensive_features.xlsx`
(SHA-256
`9e3ff40cf45d5a539832595acb2e87831cfd4f6afa3aa0b2aff0edc8ad08485e`)
in desktop Microsoft Excel for macOS without a repair, recovery, or
compatibility dialog. The retained `styles.xml` contains one header alignment
with `horizontal="center"`, `vertical="center"`, and `wrapText="1"`; with A1
selected, Excel's accessibility tree reported Center, Middle Align, and Wrap
Text as active. This verifies native interpretation of that one authored
alignment combination only.

A second generated artifact,
`target/office-verification/typed-docx-section/comprehensive_test.docx`
(SHA-256
`b1f71cfdd4f83ddf20dc7262a1aab38470652a90b4539871b3992462e2f17200`),
passed ZIP integrity and contains canonical `pageBottom`, `docEnd`, and blue
double-edge `1F4E78` section values. Desktop Microsoft Word for macOS opened
the exact file without a repair or recovery dialog, visibly rendered the page
border, placed both footnotes at the bottom of page 4, and placed both endnotes
at the document end on page 5. Word did label the document **Compatibility
Mode**, so this is deliberately not recorded as a no-compatibility-mode gate.
Neither native probe includes an Office edit/resave, reverse-read of an Office
copy, other domain variants or Office builds, or performance evidence.

## Geometry, page setup, and presentation-time fixed domains

The next fixed-domain slice centralizes DrawingML preset geometry in
`litchi-drawingml::geom::{Preset, TextPreset}`. The 187 shape tokens and 41
text-warp tokens were compared against `dml-main.xsd` in both checked-in schema
sources: the nested `OfficeOpenXML-XMLSchema-Strict.zip` inside
`ECMA-376-1_5th_edition_december_2016.zip`, and the nested
`OfficeOpenXML-XMLSchema-Transitional.zip` inside
`ECMA-376-4_5th_edition_december_2016.zip`. The ordered sets are identical
between those dialects. The old DOCX and XLSX partial enums, invalid text-warp
spellings, and open-ended preset strings are removed. XLSB authoring reuses the
same XLSX/DrawingML vocabulary during the migration instead of maintaining a
fourth token table.

XLSX parsing and authoring now use one
`Geometry::{Preset, Custom(Box<XlsxCustomGeometry>)}` state for ordinary and
connection shapes. Competing or duplicate geometry elements fail parsing;
preset/custom conflicts cannot be constructed in the semantic model; and
custom geometry can be borrowed or moved out without cloning its vectors. The
box is a deliberate cold-payload indirection that keeps `Geometry` within two
machine words rather than making every preset shape as large as the custom
model. Parsing preserves every bounded XML-safe `ST_GeomGuideName` token that
the schema admits, while authoring rejects empty, whitespace-bearing, and
numeric guide names because formula tokenization or the adjustable-value union
would make those references ambiguous. Focused tests cover all preset token
round trips, rejection of former non-schema tokens, preset/custom exclusivity,
custom-geometry parse/write round trips, DOCX text-box geometry and text-warp
parsing, and XLSB shape authoring.

PresentationML's ten modeled `ST_UniversalTimeOffset` fields now use
`litchi-pptx::time::Offset`. Parsing follows `[MS-PPTX]` section 2.3.4.6,
normalizes exact values to decimal milliseconds, and bounds both producer
spellings and canonical output. Semantic equality, hashing, and ordering make
equivalent unit spellings identical, including bookmark-uniqueness checks.
Laser traces, slide-show events, media trim/fade, and media bookmarks serialize
canonical values. Malformed values fail before entering a draft; exact
`Duration` conversion fails when precision or range would be lost. This is a
typed representation and correctness result, not a claim that arbitrary-
precision decimal arithmetic is allocation-free.

SpreadsheetML page authoring now uses
`page_setup::{Orientation, Paper, Scale, Fit, FirstPage, Copies, Dpi, Measure,
Order, Comments, ErrorMode, Setup}`. The numeric wrappers encode the Office
ranges in `[MS-OI29500]` section 2.1.638: reserved paper codes fail, scale
admits only automatic zero or 10 through 400, fit counts stop at 32,767,
copies are 1 through 32,767, DPI is positive, and the first-page domain is the
disjoint `-32,767..=-1 | 1..=32,767` range using SpreadsheetML's unsigned wire
encoding for negative values. The writer preserves absence versus explicit
defaults, and `remove_page` returns the owned setup while leaving independently
authored `pageSetUpPr` flags intact. `set_fit` updates both fit dimensions and
enables the independent policy as one semantic operation. The duplicate raw
`u32`/`bool` worksheet read model and its second parser are deleted; immutable
`Worksheet::page` is backed by the one complete typed parser.

Public `Setup` no longer carries a printer relationship ID. Relationship
projection and mutation live in `xlsx::printer_settings`, validate NCNames and
dialect-matched relationship namespaces, reuse an in-scope prefix or allocate
a non-conflicting one, and reject duplicate or dangling states. Page setup is
anchored to the worksheet-root dialect, so a Strict child under a Transitional
worksheet (or the inverse) is not accepted by selecting a dialect from the
child. The relationship-only projection intentionally tolerates unrelated
producer page-setting quirks needed to load real printer blobs, while the
public semantic setup parser remains strict. Focused parser, writer, package
round-trip, strict-Clippy, example, and formatting gates cover this slice.
These functional gates do not establish printer-specific output, pagination,
allocation, latency, or throughput.

The same migration closes additional token domains instead of adding
compatibility shims:

- SpreadsheetML font underline, scheme, and vertical script are compact exact
  enums. Explicit `none` and `baseline` survive read/write round trips.
- Conditional-format rule kind, operator, value kind, time period, data-bar
  direction/axis, color role, core icon set, Office 2010 icon set, sheet
  visibility, and sort icon set are typed. Core and x14 icon sets cannot be
  mixed accidentally. Conditional-format and tab colors reuse the checked
  four-byte `styles::Rgb` value.
- The legacy worksheet validation/conditional-format structs and the second
  streaming parse/copy path are removed. Dedicated strict parsers populate the
  typed models before any worksheet state is assigned, so a failure cannot
  partially mutate the worksheet.
- Word numbering format/multilevel type, note position, and all 65
  Transitional compatibility flags are fixed enums. Strict documents admit
  only the schema's seven-flag subset; unknown names fail.
- Shared DrawingML body/run properties use typed anchor, direction, wrap,
  autofit, underline, coordinate, column-count, and text-size domains across
  DOCX and XLSX. Content controls expose `content_control::Kind` rather than a
  tag-name string.

Focused exhaustive token tests compare these domains with the checked-in
Strict/Transitional schemas where both dialects apply. Unknown-token,
cross-dialect, explicit-default, and read/write tests are correctness evidence;
they are not performance measurements.

Presentation media now preserves absent trim/fade values separately from
explicit zero, makes seek time part of the `Seek { at }` event variant, and
uses checked shared DrawingML coordinates. The XSD audit is significant:
`ST_PositiveCoordinate` is integer-only and permits zero, so `Extent` is a
transparent checked `0..=27,273,042,316,900` value and rejects unit-bearing
spellings. Bounded inert media extension lists round-trip canonically; p14
relationship attributes use the required Transitional relationship namespace
even in a Strict package. Shared `MediaData` clones retain one immutable byte
allocation and expose a move-first `try_into_vec` path. Focused tests cover
media, slide-show events, coordinates, namespace behavior, extension bounds,
sharing, and integration round trips. These are functional/storage-invariant
tests, not throughput or resident-memory benchmarks.

The native Excel probe deliberately used a shape-only worksheet, rather than
one whose chart or image happened to force drawing attachment. The first probe
opened without a repair prompt but displayed no drawing objects. Package
inspection traced that result to a worksheet drawing relationship and part
without the required worksheet `<drawing>` reference: worksheet serialization
had incorrectly gated the element on chart/image collections and guessed
`rId1`. The writer now emits the element only from its assigned drawing
relationship ID, and the package regression resolves that exact ID to the
drawing part before reopening the workbook through the public API.

After that correction, Microsoft Excel for Mac opened
`target/office-verification/typed-page-shapes/xlsx_typed_page_and_shapes.xlsx`
without repair, recovery, or compatibility UI. Excel's accessibility model
reported 16 populated cells and two objects on sheet `Typed API`, exposed A1
as bold and single-underlined, identified the rose-to-green color scale on
A2:A6, identified red/yellow/green traffic-light icons on B2:B6, and reported
conditional formatting on C2:C6. It also reported the selected tab as blue.
Page Layout reported A4 selected, Landscape selected, width `1 page`, and
height `Automatic`, matching the typed `Setup`. The exact probed artifact
SHA-256 is
`16002ab60cbf775d244d03fb41b2e0f61e4319435f083aa278eea6265c6ac11c`;
ZIP integrity and package inspection also confirmed the named `Summary` and
`Status` objects, `roundRect`, `ellipse`, and `<drawing r:id="rId1"/>`. This is
evidence for opening and application interpretation in the installed Excel
build, not for edit/resave fidelity, other Office versions, printer drivers,
every fixed token, or performance.

## Failure-atomic PPTX publication and owned font buffers

PPTX save preparation now treats legacy-writer materialization, optional font
publication, core-property staging, and the final sink as one publication
transaction. An unchanged presentation with no property edit or requested font
work writes the original OPC graph directly. A dirty path snapshots bounded
package metadata while built-in part payloads retain their shared `Arc`
allocations, materializes the presentation, applies optional font work, and
stages core properties without clearing their edit intent. Only a successful
sink marks the presentation clean and commits the exact property guard. Any
error restores the prior package graph and leaves both edits retryable.

Encrypted file save includes serialization, encryption, and atomic destination
replacement inside that same transaction. Producing an encrypted byte vector
is itself a successful publication boundary, but a pre-publication filesystem
failure no longer inherits an earlier in-memory commit. The atomic layer does
not yet distinguish a directory-sync failure after rename. In that case the
new destination may already be visible while `write_with` restores dirty
in-memory state and encrypted-save profile state; this committed-but-not-known-
durable divergence remains follow-up work. A focused injected-sink regression
observes the fully staged slide and property in the candidate, then proves
pointer-identical presentation-payload restoration, retained dirty intent, a
successful retry, and semantic reopen.

PPTX raw OPC access is now a fallible, explicitly low-level boundary. `opc`
rejects a graph whose presentation writer, core-property slot, or font policy
still has managed state pending; encrypted sources also reject raw plaintext
exposure. Mutable access no longer returns an escaping `&mut OpcPackage`.
`edit_opc` applies a closure to a structural candidate whose built-in payloads
share immutable `Arc` storage. Error or unwind leaves that candidate
unpublished, but custom `Part` implementations retain their own clone and
interior-mutability policy. The boundary rejects automatic font-policy
changes, validates that the PowerPoint main relationship is singular and
internal with an allowed PPTX content type, validates the core-property graph,
reloads the property slot, and only then commits while disabling the legacy
writer. It does not parse the complete PresentationML graph. Covered notes,
tag, slide, master, layout, and theme graph mutators use the same
candidate-and-mode-transition rule, so a later writer materialization cannot
erase their successful edits. Managed `to_bytes` preserves the encrypted-
source policy, while `to_plain_bytes` makes plaintext extraction explicit.
DOCX, XLSX, and the binary-format hosts still require equivalent boundary work.

The system-font loader also consumes a uniquely owned font-kit memory `Arc` via
`Arc::unwrap_or_clone`: the unique case keeps the original `Vec`
allocation, while a genuinely shared handle copies and leaves its other owner
intact. Pointer-identity tests cover both ownership branches. These are
structural copy and transaction guarantees, not measured latency, allocation,
or throughput claims; ADR 0005 performance evidence remains outstanding.

## DOCX, PPTX, and XLSB semantic owner follow-up

The next concrete-crate migration batch closes three more pure-value seams while
leaving package traversal in the format adapter:

- DOCX formatting values and document statistics now live in `litchi-docx`.
  For these seams, `litchi-ooxml::docx` retains compatibility re-exports and
  owns only document traversal, including the checked aggregation of image and
  drawing counts.
- PPTX image/text formatting values now live in `litchi-pptx`; the OOXML host
  retains its historical module path as a re-export.
- XLSB date serial utilities now live in `litchi-xlsb`. The 1900/1904 date
  system selection remains compatible with the workbook `f1904` flag described
  by `[MS-XLSB]` section 3.7.

The owner crates retain focused unit coverage, and the DOCX adapter has a
reopened-package regression covering the public statistics snapshot. This is
functional and boundary evidence only; native Office and performance evidence
remain governed by the evidence levels below.

## Spreadsheet views, legacy comments, and hyperlink owner follow-up

The next concrete-crate migration batch closes four format-specific semantic
and codec seams while retaining compatibility paths in the OOXML host:

- SpreadsheetML worksheet-view parsing now lives in `litchi-xlsx::sheet_view`.
  The owner validates A1 cell and range references locally, models panes,
  selections, pivot selections, and retained extensions, and owns the bounded
  MCE-aware codec. `litchi-ooxml::xlsx::sheet_view` remains a thin adapter.
- PresentationML legacy comment values, XML codecs, and package-graph CRUD now
  live in `litchi-pptx::comments`. Presentation traversal and historical host
  paths remain adapters, with typed PPTX errors crossing the boundary.
- WordprocessingML hyperlink values and paragraph/document extraction now live
  in `litchi-docx::hyperlink`. Relationship resolution remains an explicit
  host-provided input, and the host document and paragraph APIs preserve their
  historical return types and paths.
- The XLSB `BrtHLink` range codec now lives in `litchi-xlsb::hyperlinks`, using
  the validated BIFF12 cursor and writer. Worksheet relationship creation and
  package orchestration remain in `litchi-ooxml`, while the host writer uses the
  fallible owner serialization path.

These migrations were checked against the corresponding checked-in
`[MS-XLSX]` worksheet-view, `[MS-PPTX]` legacy-comment, `[MS-OE376]`/
`[MS-DOCX]` hyperlink, and `[MS-XLSB]` section 2.4.693 specifications. The
owner crates pass 17 XLSB, 285 XLSX, 115 PPTX, and 136 DOCX unit tests, strict
all-features Clippy, and workspace formatting and boundary checks. The full
`litchi-ooxml` package suite passes, including 1,844 host unit tests and its
integration and doctest surfaces. This is functional and boundary evidence;
native Office and performance evidence remain governed by the evidence levels
below.

## DOCX fields, PPTX timing, and XLSX validation/protection owner follow-up

The next concrete-crate migration batch closes four format-owned semantic and
codec seams while retaining the historical host paths as compatibility
adapters:

- WordprocessingML field instruction/result semantics and bounded field XML
  extraction now live in `litchi-docx::field`. Document traversal remains in
  `litchi-ooxml`, with the host's legacy error type bridged to the owner error.
- PresentationML timing trees, animation behavior values, bounded timing XML,
  and animation relationship validation now live in `litchi-pptx::animations`.
  Slide/package traversal and writer ordering remain host adapters.
- SpreadsheetML data-validation collections, formula values, core/x14
  extension handling, and atomic worksheet replacement now live in
  `litchi-xlsx::data_validation`.
- SpreadsheetML sheet-protection metadata, protected-range references,
  legacy/strong verifier metadata, core/x14 extension handling, and ordered
  worksheet replacement now live in `litchi-xlsx::sheet_protection`.
  Workbook and worksheet transaction orchestration remains in `litchi-ooxml`.

The codecs were checked against the corresponding checked-in `[MS-DOCX]` and
`[MS-OE376]` field sections, `[MS-PPTX]` timing/animation sections, and
`[MS-XLSX]` data-validations sections 2.4.7 and 2.6.3--2.6.5 plus protected
ranges sections 2.4.10 and 2.6.55--2.6.56. The owner crates pass 227 DOCX,
160 PPTX, and 294 XLSX unit tests; DOCX also passes its nine doctests. Strict
all-features Clippy passes for all three owner crates, default all-target
Clippy passes for `litchi-ooxml`, formatting and boundary checks pass, and the
full host package surface passes 1,699 unit tests plus its integration and
doctest targets. Host all-features Clippy remains environment-blocked before
compilation because `pkg-config`/fontconfig is unavailable. This is functional
and boundary evidence; native Office and performance evidence remain governed
by the evidence levels below.

## PPTX properties and XLSX worksheet-settings owner follow-up

The next concrete-crate migration batch closes four additional codec seams
while retaining the historical `litchi-ooxml` module paths as thin adapters:

- PresentationML presentation properties, typed print/web/show values, and
  inert extension payloads now live in `litchi-pptx::presentation_properties`.
- PresentationML view, pane, guide, splitter, and slide-list values now live
  in `litchi-pptx::view_properties`. Presentation/package relationship
  traversal remains in the OOXML host and crosses the typed PPTX error
  boundary.
- SpreadsheetML `dataConsolidate`, `dataRefs`, and `dataRef` values now live in
  `litchi-xlsx::data_consolidation`, including checked A1 references, source
  relationship identifiers, bounded counts, and deterministic serialization.
- SpreadsheetML worksheet `pageSetup` values now live in
  `litchi-xlsx::page_setup`; printer-settings relationship projection remains
  in the host printer-settings adapter while the owner validates the typed
  relationship identifier.

The codec choices were checked against the checked-in `[MS-OE376]` 2.1.24
(Part 1 §13.3.7 presentation properties), 2.1.1148 (Part 4 §4.3.2.6 normal
view properties), and 2.1.666--2.1.667 (Part 4 §3.3.1.60--61 page margins and
page setup) references, the `[MS-OI29500]` 2.1.612 (Part 1 §18.3.1.29
`dataConsolidate`) and 2.1.637--2.1.638 page-settings references, and the
checked-in `[MS-PPTX]` presentation-properties extension structures. Owner
tests pass: 171 PPTX and 311 XLSX unit tests. The full host package suite also
passes 1,671 unit tests plus its integration and doctest targets. Owner
all-features Clippy, host default all-target Clippy, boundary, and formatting
checks pass. Host all-features Clippy remains environment-blocked before
compilation because `pkg-config`/fontconfig is unavailable. This is functional
and boundary evidence; native Office and performance evidence remain governed
by the evidence levels below.

## PPTX comments/media and XLSX query-table owner follow-up

The next owner batch moves three larger format-specific seams out of the
migration host while retaining package traversal and historical error/API
surfaces in explicit adapters:

- PresentationML modern comments and their author list, relationship graph,
  bounded MCE-aware codecs, and CRUD operations now live in
  `litchi-pptx::modern_comments` and `litchi-pptx::modern_comment_authors`.
  The host retains presentation/package traversal and maps the typed owner
  errors back to its historical boundary.
- PresentationML audio/video picture values, extension metadata, inert media
  resources, and the slide media OPC graph now live in
  `litchi-pptx::media_parts`. The host adapter preserves the existing
  `Slide::media` entry point and its historical invalid-format behavior while
  payload bytes remain shared and never decoded, fetched, or executed.
- SpreadsheetML query-table values, refresh/sort/field metadata, bounded XML,
  relationship validation, and worksheet query-table graph operations now live
  in `litchi-xlsx::query_table`. Workbook and worksheet facade traversal stays
  in `litchi-ooxml`, including the real LibreOffice query-table regression.

The checked-in specification evidence covers `[MS-PPTX]` §§2.1.1 (Media
Part), 2.1.5 (Comment Part), 2.1.6 (Author Part), 2.2.4 (Media Extensions),
and 2.16.1.1--2.16.3.7 for the Office 2018 modern-comment structures. Query
tables are pinned to `[MS-XLSX]` §§2.4.41, 2.6.88, and 2.2.4.7, with the
corresponding checked-in `[MS-OE376]` table-of-contents entries 2.1.854--
2.1.856 and `[MS-OI29500]` entries 2.1.826--2.1.828.

The owner suites pass 196 PPTX and 314 XLSX tests with all features enabled;
the complete host package suite passes 1,645 unit tests plus its integration
and doctest targets. Owner all-target strict Clippy, host default all-target
strict Clippy, formatting, diff, and crate-boundary checks pass. Host
all-features Clippy remains environment-blocked before compilation because
`pkg-config`/fontconfig is unavailable. This is functional and boundary
evidence; native Office and performance evidence remain governed by the
evidence levels below.

## XLSX filters, views, timelines, and DOCX modern-comment owner follow-up

The next owner batch moves five format-specific seams out of the migration host
while retaining package traversal, relationship orchestration, and historical
error/API surfaces in explicit adapters:

- WordprocessingML modern comment metadata (`commentsExtended`, `people`,
  `commentsIds`, and `commentsExtensible`) now lives in
  `litchi-docx::modern_comments`. The host retains document-part and
  relationship traversal and maps the typed owner errors to its historical
  `OoxmlError` boundary.
- SpreadsheetML conditional-formatting values and RGB parsing now live in
  `litchi-xlsx::{conditional_formatting,color}`. The host styles adapter
  re-exports the owner RGB type, while the conditional-formatting shim keeps
  worksheet traversal and legacy error conversion at the host boundary.
- SpreadsheetML sort states, auto-filters, and timeline cache/worksheet
  graphs now live in `litchi-xlsx::{sort,auto_filter,timelines}`. Host
  worksheet/workbook graph code remains responsible for relationship and part
  lifecycle operations.
- SpreadsheetML named sheet views now live in `litchi-xlsx::named_sheet_view`.
  The owner composes the owner sort/filter models and checked
  `litchi-sheet` addresses; the host retains worksheet discovery and package
  graph integration as an adapter.

The checked-in specification anchors are `[MS-DOCX]` §§2.1.2--2.1.5,
2.2.13, 2.5.1.5, 2.5.1.9, 2.5.3.2, 2.5.3.4, 2.8.1.1, 2.8.3.2, 2.10.1.1,
and 2.10.3.2; and `[MS-XLSX]` §§2.1.7--2.1.8, 2.2.2.2, 2.3.5--2.3.8,
2.4.6, 2.4.49--2.4.58, 2.4.88, 2.6.1--2.6.2, 2.6.98--2.6.118, and
2.6.210--2.6.211. These anchors cover the conditional-formatting,
auto-filter/sort, timeline, named-sheet-view, and modern-comment structures
implemented by this batch.

The all-feature owner suites pass 349 XLSX and 235 DOCX unit tests; DOCX's
doctest targets also pass. The complete `litchi-ooxml` package surface passes
1,609 host unit tests plus its integration and doctest targets. Owner
all-target strict Clippy, host default all-target strict Clippy, formatting,
diff, and crate-boundary checks pass. Host all-features Clippy remains
environment-blocked before compilation because `pkg-config`/fontconfig is
unavailable. This is functional and boundary evidence; native Office and
performance evidence remain governed by the evidence levels below.

## DOCX settings, XLSX chartsheets, and XLSB formula owner follow-up

The next concrete-crate migration batch moves three format-owned codec seams
out of the migration host while retaining package traversal, relationship
orchestration, and historical error/API paths as explicit adapters:

- WordprocessingML document-settings vocabulary and its bounded XML codec now
  live in `litchi-docx::settings`. Typed compatibility flags and options, note
  numbering and placement, protection, view, proofing, theme-font, and
  color-scheme values are owned by the standalone crate. The OOXML host keeps
  settings-part orchestration and host-only smart-tag, mail-merge, and
  attached-template state, and maps owner errors to its historical boundary.
- SpreadsheetML chartsheet values, conformance-aware XML parsing/writing, and
  bounded validation now live in `litchi-xlsx::chart_sheet`. The host retains
  the OPC graph for drawings, charts, images, VML, printer settings, and
  relationships, with the typed chartsheet codec crossing a narrow adapter.
- BIFF12 cell-formula buffers, RPN `Ptg` parsing and compilation, checked
  ranges, and array/shared formula records now live in `litchi-xlsb::formula`.
  Workbook link/name/table/pivot resolution and worksheet record orchestration
  remain in `litchi-ooxml`; the host wrapper preserves its `xlsb::Error` API while
  delegating binary validation and serialization.

The checked-in specification anchors for settings are `[MS-DOCX]` §§2.2.2 and
2.3, together with `[MS-OE376]` §§2.1.310--2.1.313, 2.1.403, 2.1.410,
2.1.435, 2.1.437--2.1.439, 2.1.471, 2.1.572, and 2.1.596. Chartsheet
parts and their Office relationship/profile variations are pinned to
`[MS-OE376]` §§2.1.10, 2.1.597, 2.1.639, 2.1.668, 2.1.680, 2.1.682,
2.1.684, 2.1.690, and 2.1.1126, with the corresponding `[MS-OI29500]`
chartsheet variations in §§2.1.7 and 2.1.597. Formula wire shapes and token
families are pinned to `[MS-XLSB]` §§2.2.2, 2.4.6, 2.4.796, 2.5.98.4,
2.5.98.12, 2.5.98.16, 2.5.98.88--2.5.98.92, and 2.5.98.98.

The owner suites pass 239 DOCX, 349 XLSX, and 21 XLSB unit tests; the XLSB
integration targets and DOCX doctests also pass. The host package's 1,609
unit tests pass, along with its integration and doctest targets. Owner
all-feature all-target strict Clippy, host default all-target strict Clippy,
formatting, diff, and the crate-boundary audit pass: 35 workspace packages,
107 internal dependency declarations, and the same 13 explicitly scheduled
debt edges remain. Host all-features Clippy remains environment-blocked before
compilation because `pkg-config`/fontconfig is unavailable. This is functional
and boundary evidence; native Office and performance evidence remain governed
by the evidence levels below.

## XLSX ActiveX, external links, and DOCX mail-merge owner follow-up

The next owner batch moves three format-owned codec seams out of the
migration host while retaining package traversal, relationship orchestration,
and historical error/API paths as explicit adapters:

- SpreadsheetML worksheet-control metadata, ActiveX descriptor/property XML,
  and opaque persistence/preview resources now live in
  `litchi-xlsx::active_x`. The owner enforces bounded XML and package
  relationships but never resolves a CLSID, instantiates a control, decodes
  persistence data, follows an external target, or executes code. The host
  retains the worksheet graph and maps owner errors to `OoxmlError`.
- SpreadsheetML external-link values, cached DDE/OLE/workbook data, and
  bounded external-link XML now live in `litchi-xlsx::external_links`. The
  host retains workbook collection and OPC part lifecycle operations, while
  external targets remain inert relationship metadata.
- WordprocessingML mail-merge settings, ODSO field maps, and recipient-data
  XML now live in `litchi-docx::mail_merge`. The host retains settings-part
  relationship/resource orchestration, maps owner errors at the host boundary,
  and exposes no duplicate semantic model or compatibility aliases; sources
  and recipient parts are never fetched, opened, or executed.

The checked-in specification anchors are `[MS-XLSX]` §§2.1.1, 2.2.4.3,
2.4.25, 2.4.89, 2.6.46, 2.6.215, and 2.6.227--2.6.228; `[MS-OE376]`
§§2.1.18, 2.1.32, 3.4.1.1, and 3.6.1.1--3.6.2.1 for the spreadsheet
control, external-workbook, and ActiveX relationship/profile variations; and
`[MS-OE376]` §§2.1.367, 2.1.381, 2.1.384, 2.1.386, 3.1.1.3, and
3.1.2.2.1.2--3.1.2.2.1.3 for mail-merge settings and recipient data. These
anchors cover the stored models, strict/transitional namespaces, relationship
validation, bounded cached values, and inert binary/base64 preservation
implemented by this batch.

The owner suites pass 244 DOCX and 360 XLSX unit tests; DOCX's doctest targets
also pass. The complete `litchi-ooxml` package surface passes 1,593 host unit
tests plus its integration and doctest targets. Owner all-feature all-target
strict Clippy, host default all-target strict Clippy, formatting, diff, and
the crate-boundary audit pass: 35 workspace packages, 107 internal dependency
declarations, and the same 13 explicitly scheduled debt edges remain. Host
all-features Clippy remains environment-blocked before compilation because
`pkg-config`/fontconfig is unavailable. This is functional and boundary
evidence; native Office and performance evidence remain governed by the
evidence levels below.

## XLSB conditional formatting and external links, XLSX slicer-cache follow-up

The next disjoint owner batch moves three format-owned codec seams out of the
migration host while retaining package traversal, relationship orchestration,
and historical host paths as explicit adapters:

- BIFF12 classic and Office 2013 conditional-formatting models, formula-bearing
  thresholds, visualizations, extension GUIDs, validation, and bounded record
  parsing/writing now live in `litchi-xlsb::conditional_formatting`. The host
  retains worksheet/contextual formula resolution and worksheet/package record
  orchestration.
- BIFF12 External Link models, restricted external-name Ptgs, workbook/DDE/OLE
  cached values, and bounded `BrtSupBook` stream parsing/writing now live in
  `litchi-xlsb::external_link`. The host retains external-link OPC part and
  relationship lifecycle operations. Links remain inert metadata: no external
  workbook, DDE server, OLE object, refresh, or code path is opened or invoked.
- SpreadsheetML slicer-cache definition XML now lives in
  `litchi-xlsx::slicer_cache`. The owner retains bounded inert
  `x14:slicerCacheDefinition` data, while the host retains workbook-extension
  edits, internal/no-outbound-relationship checks, slicer cross-validation, and
  atomic OPC graph mutations.

The checked-in specification anchors are `[MS-XLSB]` §§2.2.6.2.1,
2.4.23--2.4.24, 2.4.43--2.4.44, 2.4.91--2.4.92, 2.4.332--2.4.335,
2.4.380--2.4.381, 2.4.399--2.4.400, 2.4.445--2.4.446, 2.5.19--2.5.20,
2.5.98.7 for conditional formatting; and §§2.1.7.25, 2.2.7.4,
2.2.7.4.2--2.2.7.4.3, 2.4.235, 2.4.588, 2.4.720--2.4.721, 2.4.811,
2.5.34, 2.5.97, and 3.5 for External Links. Slicer-cache parts and their
relationship/profile rules are pinned to `[MS-XLSX]` §§2.1.4, 2.2.4.8,
2.3.2.1, 2.4.38, 2.4.60, 2.6.70--2.6.85, 2.6.97, and 2.6.103--2.6.104.

The owner suites pass 101 listed XLSB tests and 364 XLSX unit tests; the
complete `litchi-ooxml` package surface passes 1,543 host unit tests plus its
integration and doctest targets. Owner all-feature all-target strict Clippy,
host default all-target strict Clippy, formatting, diff, and the crate-boundary
audit pass: 35 workspace packages, 107 internal dependency declarations, and
the same 13 explicitly scheduled debt edges remain. Host all-features Clippy
remains environment-blocked before compilation because `pkg-config`/fontconfig
is unavailable. This is functional and boundary evidence; native Office and
performance evidence remain governed by the evidence levels below.

## XLSB data validation, XLSX workbook metadata, and DOCX bibliography follow-up

The next disjoint owner batch moves three bounded format-owned XML/record
seams out of the migration host while retaining package traversal, relationship
provenance, formula-context binding, and historical error/API paths as explicit
adapters:

- BIFF12 classic and Office 2013 data-validation records, collection settings,
  inline-list payload validation, formula-token retention, range limits, and
  semantic validation models now live in `litchi-xlsb::data_validation`. The
  host keeps worksheet record traversal and text-formula compilation/writing,
  implements the owner formula-resolution and binary-formula bridges, and maps
  owner failures to `xlsb::Error`.
- SpreadsheetML workbook `metadataTypes`, `futureMetadata`, `cellMetadata`,
  `valueMetadata`, and inert extension XML now live in
  `litchi-xlsx::workbook_metadata`. The host keeps the workbook relationship,
  content-type, no-outbound-relationship, and OPC discovery checks.
- Word bibliography `Sources`/`Source` namespace-aware parsing, scalar paths,
  style metadata, strict/transitional/legacy namespace recognition, and
  bounded Custom XML payload handling now live in `litchi-docx::bibliography`.
  The host keeps Custom XML item discovery, relationship/part provenance, and
  the existing source CRUD writer compatibility surface. Bibliography styles,
  citation matching, XSLT, and external resources remain inert.

The checked-in specification anchors are `[MS-XLSB]` §§2.4.55--2.4.56,
2.4.356--2.4.358, 2.5.36--2.5.37, 2.5.58--2.5.66, 2.5.98.8, and 2.5.156;
`[MS-XLSX]` §2.2.4.4 plus the referenced SpreadsheetML metadata structures;
and `[MS-OE376]` Part 4 §2.16.5.11 with normative variation §2.1.494 for
the BIBLIOGRAPHY vocabulary and field compatibility. These anchors cover the
record payloads, metadata ordering/counts, future extensions, stored scalar
paths, and strict/transitional/legacy namespace behavior implemented here.

The owner suites pass 114 XLSB, 367 XLSX, and 248 DOCX tests. The default
`litchi-ooxml` package test target passes 1,530 host library tests plus its
integration and doctest targets; the focused owner and host commands pass as
well. Owner all-feature all-target strict Clippy, host default all-target
strict Clippy, formatting, diff, and the crate-boundary audit pass: 35
workspace packages, 107 internal dependency declarations, and the same 13
explicitly scheduled debt edges remain. The all-features host test and full
workspace test commands remain environment-blocked before compilation because
`pkg-config`/fontconfig is unavailable. No native Office or performance claim
is made for this batch.

## XLSX tables, XLSB PivotTable views, and DOCX smart-tag settings follow-up

The next disjoint owner batch moves three format-owned model/codec seams out
of the migration host while retaining package traversal, worksheet/workbook
orchestration, relationship validation, and historical host paths as explicit
adapters:

- SpreadsheetML table models, range and column validation, table-column
  formulas, auto-filter/sort state, style information, bounded parsing, and
  deterministic XML writing now live in `litchi-xlsx::table`. The host retains
  worksheet table collections, table-part discovery, and package/relationship
  traversal. The legacy writer path delegates to the owner serializer.
- BIFF12 PivotTable-view framing, `BrtBeginSXView` identity extraction,
  `BrtEndSXView` boundary validation, bounded record scanning, and lossless
  stream retention now live in `litchi-xlsb::pivot_view`. The host retains
  workbook/sheet/package orchestration and maps owner failures to the
  historical `xlsb::Error` API.
- WordprocessingML smart-tag vocabulary values now live in
  `litchi-docx::settings`. The owner validates the checked client length
  domains while retaining empty-but-present attribute values; the host keeps
  settings-part parsing, required-attribute/cardinality checks, relationship
  validation, and attached-template/lossless settings orchestration.

The checked-in specification anchors are `[MS-XLSX]` §§2.4.22 and 2.6.35 for
the `table` global element and `CT_Table`; `[MS-XLSB]` §§2.1.7.40, 2.4.278,
2.4.631, and 2.5.169 for PivotTable parts, `BrtBeginSXView`,
`BrtEndSXView`, and `XLWideString`; and `[MS-OE376]` §§2.1.615--2.1.616 for
the `smartTagType` and `smartTagTypes` settings vocabulary. These anchors
cover the models, record boundaries, strict/transitional settings namespace
handling, bounded input/output, and inert lossless preservation implemented
by this batch.

The owner unit suites pass 387 XLSX, 86 XLSB, and 249 DOCX tests, with their
available integration and doctest targets passing as well. The no-default-
features `litchi-ooxml` all-target suite passes 1,504 host library tests plus
its integration and doctest targets. Owner all-feature all-target strict
Clippy, host no-default-feature all-target strict Clippy, formatting, diff,
and the crate-boundary audit pass: 35 workspace packages, 107 internal
dependency declarations, and the same 13 explicitly scheduled debt edges
remain. Host all-features Clippy, host all-features tests, and the full
workspace tests remain environment-blocked before compilation because
`pkg-config`/fontconfig is unavailable. No native Office or performance claim
is made for this batch.

## DOCX variables, XLSB comments, and XLSX custom-data owner follow-up

This batch applies the layered-module and concise-name rule across all three
owner crates. Each new owner surface has a small `mod.rs` facade plus semantic
`model.rs` and `codec.rs` layers; format-owned structs do not repeat the
module prefix. Existing host paths remain compatibility adapters while their
legacy flat modules are migrated in later disjoint batches.

- `litchi-docx::variables` owns `DocumentVariables` and the bounded
  WordprocessingML codec. The host retains OPC-part access, MCE preprocessing,
  and the historical `litchi_ooxml` error/API boundary.
- `litchi-xlsb::comments` owns `Comment`, `CommentRun`, BIFF12 record framing,
  and rich-string validation. The host retains `SharedStringRun` conversion
  and the historical comment model.
- `litchi-xlsx::custom_data` owns concise `Properties` and `ExtensionList`
  models plus their bounded XML codec. The host retains opaque payloads,
  package relationships, content types, and atomic load/store orchestration.

The checked-in specification anchors are `[MS-OE376]` §2.1.411 for `docVar`;
`[MS-XLSB]` §§2.1.7.8, 2.4.30--2.4.33, 2.4.340--2.4.341, and
2.4.387--2.4.390 for the comments grammar and records; and `[MS-XLSX]`
§§2.1.2--2.1.3, 2.4.35, 2.6.34, and 2.6.66 for custom-data parts,
`datastoreItem`, `embeddedDataId`, and `CT_DatastoreItem`.

The owner unit suites pass 253 DOCX, 89 XLSB, and 390 XLSX tests, with their
available integration and doctest targets passing. The no-default-features
`litchi-ooxml` all-target suite passes 1,504 host library tests plus its
integration and doctest targets. Owner all-feature all-target strict Clippy,
host no-default-feature all-target strict Clippy, formatting, diff, and the
crate-boundary audit pass: 35 workspace packages, 107 internal dependency
declarations, and the same 13 explicitly scheduled debt edges remain. Host
all-features tests/Clippy and the full workspace test remain blocked before
compilation because `pkg-config`/fontconfig is unavailable. No native Office
or performance claim is made for this batch.

## DOCX numbering, XLSB styles, and XLSX threaded-comments owner follow-up

This batch applies the concise-name and layered-module rule across all three
format crates. Each owner surface is split into `mod.rs`, `model.rs`, and
`codec.rs` (with `tests.rs` where the owner-level suite is substantial); host
files retain only MCE/OPC preprocessing, relationship/content-type traversal,
and legacy representation conversion:

- `litchi-docx::numbering` owns the package-neutral numbering collection,
  definitions, instances, levels, overrides, closed numbering enums, picture
  bullets, and bounded WordprocessingML state machine. The host exposes this
  owner vocabulary under its contextual `numbering` module and retains only
  MCE/OPC extraction and error mapping; `Numbering`, `AbstractNum`, `Num`, and
  the other prefix-expanded compatibility spellings are not retained.
- `litchi-xlsb::styles` owns neutral alignment, border, font, fill, number
  format, cell-format, styles-table models and the strict Brt* codec. The host
  preserves `StylesTable` and converts owner alignment/border values to the
  existing host types.
- `litchi-xlsx::threaded_comments` owns concise `Comment`, `Comments`,
  `Person`, `People`, `Mention`, and graph models plus bounded XML parsing,
  writing, and cross-reference validation. The host retains only package graph
  CRUD and relationship lifecycle operations; semantic values are available
  from the owner module and no prefix-expanded aliases are retained.

The checked-in specification anchors are `[MS-OE376]` §§2.1.277--2.1.291 and
2.1.580 for numbering domains and limits; `[MS-XLSB]` §§2.3.7, 2.4.12,
2.4.20, 2.4.22, 2.4.87, 2.4.89, 2.4.232, 2.4.314, 2.4.369, 2.4.377,
2.4.441, 2.4.585, 2.4.688, 2.4.690, and 2.4.876 for styles records; and
`[MS-XLSX]` §§2.1.17--2.1.18, 2.3.7--2.3.7.2, 2.4.85--2.4.86, and
2.6.202--2.6.207 for threaded comments, people, mentions, and part roots.

The owner suites pass 254 DOCX, 91 XLSB, and 395 XLSX tests, with their
available integration and doctest targets passing. The no-default-features
`litchi-ooxml` all-target suite passes 1,499 host library tests plus its
integration and doctest targets. Owner all-feature all-target strict Clippy,
host no-default-feature all-target strict Clippy, formatting, diff, and the
crate-boundary audit pass: 35 workspace packages, 107 internal dependency
declarations, and the same 13 explicitly scheduled debt edges remain. Host
all-features tests/Clippy and the full workspace test remain blocked before
compilation because `pkg-config`/fontconfig is unavailable. No native Office
or performance claim is made for this batch.

## PPTX modern comments, DOCX modern comments, and XLSB conditional-formatting owner follow-up

This batch applies the layered owner and concise-name rule to the three
format-specific seams highlighted by the public API audit:

- `litchi-pptx::modern_comments` is now the single owner folder for comments
  and authors, split into `model.rs`, `codec.rs`, and `package.rs`. The former
  `modern_comment_authors` module and the `ModernComment*` aliases were
  removed in the owner-only API convergence pass below; `Comment`, `Author`,
  `Part`, `Graph`, and related models are canonical.
- `litchi-docx::modern_comments` is split into `model.rs`, `codec.rs`, and
  `package.rs`. `Conformance`, `Comment`, `Reaction`, `Metadata`, `Person`,
  and related concise models are canonical; historical expanded spellings
  were removed in the owner-only API convergence pass below.
- `litchi-xlsb::conditional_formatting` is split into `model.rs`, `codec.rs`,
  and `tests.rs`. `Formatting`, `Rule`, `RuleType`, `Value`, `Color`, `Bar`, and
  related names are canonical. The owner now also contains focused facade
  tests, while the host publishes no conditional-formatting aliases or
  duplicate codec.

The shared OOXML `ST_Guid` lexical validator already owned by
`litchi-ooxml-common::custom_xml` is now reused by PPTX modern comments, XLSX
threaded comments, XLSX data validation, and OOXML revision parts. No
format-specific XML DOM or binary-record codec was forced into a common crate;
the OLE/IWA grammars remain separate until a format-neutral seam is proven.

Checked-in specification anchors are `[MS-PPTX]` §§2.1.5--2.1.6, 2.2.10,
2.4.3.2, 2.4.3.6, 2.16.1.1--2.16.1.3, 2.16.3.3--2.16.3.8, and
2.16.4.3--2.16.4.4; `[MS-DOCX]` §§2.1.2--2.1.5, 2.5, 2.8, and 2.10 plus
`[MS-OREACTXML]` §2.1; and `[MS-XLSB]` §§2.3 and 2.4.23--2.4.36,
2.4.332--2.4.335, and 2.4.380--2.4.393.

Focused owner all-target tests pass: 254 DOCX, 196 PPTX, 91 XLSB, and 395
XLSX library tests, with their integration/doctest targets passing. The
no-default-features `litchi-ooxml` suite passes 1,499 host library tests and
all targets; strict owner and host Clippy, formatting, diff, and the crate
boundary audit pass. The audit remains at 35 workspace packages, 107 internal
dependency declarations, and 13 scheduled debt edges. The full workspace
test remains blocked before project compilation because `pkg-config` is not
available for `yeslogic-fontconfig-sys`/fontconfig. No native Office or
performance claim is made for this batch.

## PPTX legacy comments, XLSX query tables, and XLSB formula owner layering

This batch applies the same semantic folder rule to three remaining flat
owners while preserving the existing host adapters and public compatibility
paths:

- `litchi-pptx::comments` is layered as `model.rs`, `codec.rs`, and
  `package.rs`. `Conformance`, `Author`, `Comment`, `List`, and `Comments`
  are the canonical contextual names; historical `PresentationComment*` and
  `SlideCommentList` names are aliases.
- `litchi-xlsx::query_table` is layered as `model.rs`, `codec.rs`, and
  `package.rs`. `Table`, `WorksheetTable`, and the unprefixed value/enums are
  canonical; historical `QueryTable*` names remain aliases.
- `litchi-xlsb::formula` is layered as `model.rs`, `codec.rs`, and
  `function_table.rs`. `Range`, `ParsedFormula`, `Token`, `Parser`,
  `Compiler`, and the table/reference models are canonical; historical
  `Formula*` and `CellParsedFormula` names remain aliases.

The shared OOXML `ST_Guid` validator from
`litchi-ooxml-common::custom_xml` is now reused by XLSX chartsheet, slicer
cache, and timeline owners in addition to the earlier common-validation
surfaces. No XML-tree or BIFF12 formula abstraction was promoted across
formats: the retained query-table extension tree and Ptg grammar have
format-specific limits and semantics.

Checked-in specification anchors are `[MS-PPTX]` §§2.1.5--2.1.6, 2.2.10,
and 2.4.3.2--2.4.3.6 for legacy comments; `[MS-XLSX]` §§2.2.4.7, 2.4.41,
and 2.6.88 for query tables; and `[MS-XLSB]` §§2.2.2, 2.5.98,
2.5.98.4, and 2.5.98.16 for formulas and Ptgs.

Focused owner all-target tests pass: 196 PPTX, 395 XLSX, and 91 XLSB
library tests, with integration/doctest targets passing. The no-default-
features `litchi-ooxml` suite passes 1,499 host library tests and all
targets; strict owner and host Clippy, formatting, diff, and the crate
boundary audit pass. The audit remains at 35 workspace packages, 107
internal dependency declarations, and 13 scheduled debt edges. The full
workspace test remains blocked before project compilation because `pkg-config`
is not available for `yeslogic-fontconfig-sys`/fontconfig. No native Office or
performance claim is made for this batch.

## XLSB formula facade canonicalization

The subsequent XLSB formula migration removes the compatibility layer
described in the earlier historical batch. `litchi-xlsb::formula` remains the
owner of the BIFF12 RPN/Ptg codec, while the host adapter is physically layered
as `host/formula/{mod,model,pivot,resolution,table,text}.rs`. It now consumes
the neutral `Parser`, `ParsedFormula`, `Group`, `Range`, and `Compiler` types
directly. Workbook-specific pivot, table, resolution, and formula-text
semantics remain in their contextual submodules; no XLSB/XLSX formula codec
edge was introduced.

The former `FormulaParser`, `FormulaConverter`, `FormulaResolution`,
`CellParsedFormula`, `FormulaRange`, `FormulaGroup`, `FormulaPivot*`, and
`FormulaTableDefinition` facade names were removed without aliases. The
canonical owner paths preserve the `[MS-XLSB]` RPN limits and conversions
while eliminating wrapper allocations around `rgce`/`rgcb` payloads.

Focused verification passes 408 XLSB library tests and two formula integration
tests, plus all-target compilation, formatting, and the crate-boundary audit.
The existing XLSX textual-formula owner remains independent as required by
its different `[MS-XLSX]` grammar.

## DOCX mail-merge, PPTX presentation-properties, XLSX ActiveX, and shared web-owner layering

This batch continues the layered-module and concise-name rule across three
format crates and one genuinely shared OOXML owner:

- `litchi-docx::mail_merge` is layered as `model.rs`, `codec.rs`, and
  `package.rs`. `Conformance`, `Settings`, `FieldMap`, `Recipient`, and
  related contextual values are canonical and exposed only from the owner; no
  `MailMerge*` compatibility aliases or prefix-expanded facade types remain.
  Mail-merge sources and recipient parts remain inert.
- `litchi-pptx::presentation_properties` is layered as `model.rs`, `codec.rs`,
  and `package.rs`. `Properties`, `HtmlPublish`, `Web`, `Print`, `Show`, and
  related values are canonical; historical `Presentation*` and `*Properties`
  names remain aliases.
- `litchi-xlsx::active_x` is layered as `model.rs`, `codec.rs`, and
  `package.rs`. `Control`, `Descriptor`, `Binary`, `PreviewImage`, and
  `ControlSet` use contextual names; historical `ActiveX*` names remain
  aliases. Control binaries stay opaque and external targets are not followed.
- `litchi-ooxml-common::web` is layered as `model.rs`, `codec.rs`, and
  `package.rs` while retaining its `raw` constants and public API. This
  common owner remains the single implementation used by DOCX, PPTX, XLSX,
  and XLSB task-pane/web-extension paths; no format-local copies were added.

Checked-in specification anchors are `[MS-OE376]` §§2.1.18, 2.1.32,
2.1.367, 2.1.381, 2.1.384, 2.1.386, 3.1.1.3, 3.1.2.2.1.2--3.1.2.2.1.3,
3.4.1.1, and 3.6.1.1--3.6.2.1; `[MS-XLSX]` §§2.1.1, 2.2.4.3, 2.4.25,
2.4.89, 2.6.46, 2.6.215, and 2.6.227--2.6.228; `[MS-PPTX]` presentation
properties and web-extension references; `[MS-DOCX]` mail-merge settings
and `[MS-OWEXML]` §§1.3, 2.1--2.2.10. These anchors cover the bounded
models, strict/transitional namespaces, relationship validation, and inert
payload preservation.

Focused owner tests pass: 157 common, 254 DOCX, 196 PPTX, and 395 XLSX
all-target tests. The no-default-features `litchi-ooxml` suite passes 1,499
host library tests and all targets. Strict owner and host Clippy, formatting,
diff, and the crate-boundary audit pass; the audit remains at 35 workspace
packages, 107 internal dependency declarations, and 13 scheduled debt edges.
The full workspace test remains blocked before project compilation because
`pkg-config` is unavailable for `yeslogic-fontconfig-sys`/fontconfig. No
native Office or performance claim is made for this batch.

## PPTX view-properties semantic layering

The existing `litchi-pptx::view_properties` owner is now physically layered
as `model.rs`, `codec.rs`, and `package.rs`. The model owns contextual
PresentationML values; the codec owns bounded MCE-aware XML parsing,
validation, strict/transitional serialization, and fixture tests; and the
package layer owns the presentation relationship, content-type, and outline
slide-target checks. The historical `view_properties` module path and
`load_view_properties` root re-export remain unchanged.

The checked-in `[MS-OE376]` normal-view and view-properties references and
`[MS-PPTX]` view-properties structures continue to govern the model and codec.
The focused PPTX all-target suite passes 196 library tests plus its
integration/doctest targets; strict all-features Clippy, formatting, diff, and
crate-boundary checks pass. No native Office or performance claim is made for
this structural-only migration.

## XLSX named-sheet-view semantic layering

The `litchi-xlsx::named_sheet_view` owner is now layered as `model.rs`,
`codec.rs`, and `package.rs`. The model exposes contextual names such as
`Views`, `View`, `Filter`, `ColumnFilter`, `SortRules`, and `Guid`; historical
`NamedSheetView*` names remain type aliases. The codec owns bounded
SpreadsheetML/MCE parsing, inert differential-format and extension retention,
and canonical serialization. The package layer owns worksheet relationships,
content-type checks, orphan detection, and failure-atomic add/update/remove
operations.

Checked-in `[MS-XLSX]` §§2.1.19, 2.3.8, 2.4.88, and 2.6.210--2.6.211 govern
the part, worksheet association, named-sheet-view collection, and filter/sort
model. The focused all-target XLSX suite passes 395 tests; package Clippy,
formatting, diff, and crate-boundary checks pass. Full workspace verification
remains subject to the existing native `pkg-config`/fontconfig environment
requirement.

## DOCX settings semantic layering

`litchi-docx::settings` is now physically layered by responsibility rather
than kept in one flat owner: `compatibility.rs`, `notes.rs`, `editing.rs`,
`colors.rs`, and `smart_tags.rs` hold focused vocabulary models;
`model.rs` holds the aggregate settings value; `codec.rs` owns bounded
`settings.xml` parsing and serialization; and `support.rs` contains shared
local XML/error helpers. The existing `litchi_docx::settings` module path,
public names, generic note-format mapping, strict/transitional behavior, and
malformed-input limits are unchanged.

The settings vocabulary is WordprocessingML-specific, so no logic was moved
to `litchi-ooxml-common` or `litchi-ole-common`; the common-crate boundary
remains reserved for behavior shared by multiple format owners. The focused
DOCX all-target suite passes 254 tests plus integration/doctest targets, and
strict DOCX Clippy and formatting checks pass. No native Office or performance
claim is made for this structural migration.

## XLSB external-link owner layering

The `litchi-xlsb::external_link` owner is now physically layered under one
semantic module: `model.rs` contains inert external-link values and invariant
validation, `codec.rs` contains bounded `BrtSupBook` parsing, and `package.rs`
contains bounded external-link stream authoring. The canonical owner names are
contextual (`Link`, `Kind`, `DefinedName`, `DdeItem`, `OleItem`, `ValueMatrix`,
and `Parsed`); the historical `XlsbExternal*` spellings remain compatibility
aliases. OPC relationship resolution and part placement remain in
`litchi-ooxml`, while relationship-type inventories remain shared in
`litchi-ooxml-common`; no XLSB-specific logic was duplicated into a common
crate.

The checked-in `[MS-XLSB]` §§2.1.7.25, 2.2.7.4, 2.4.235, 2.4.588,
2.4.811--2.4.822, 2.5.44, and 2.5.98.2 govern the part boundary, workbook/
DDE/OLE variants, `BrtSupBook` framing, name/item records, external-reference
kind, and cached error values. The focused `litchi-xlsb` all-target suite passes
91 + 8 + 17 tests; the host `litchi-ooxml` no-default all-target suite passes
1499 tests, strict Clippy and formatting checks pass, and crate-boundary checks
remain valid. Full all-features host verification remains subject to the
existing environment's missing `pkg-config`/fontconfig dependency.

## DOCX glossary, PPTX media, XLSX validation, and ODF data-style layering

Four remaining large owners are now semantically layered under their existing
public module paths:

- `litchi-docx::glossary/{model,codec,package,graph,tests}`;
- `litchi-pptx::media_parts/{model,codec,package,tests}`;
- `litchi-xlsx::data_validation/{model,codec,package}`; and
- `litchi-odf::data_styles/{model,tokens,codec,package,tests}`.

Models use contextual names (`Entry`, `Picture`, `Validation`, `Style`, and
their related values); former repeated format/module prefixes remain only as
compatibility aliases where the old public API requires them. Parsing,
validation, serialization, package ownership, and graph/opaque-payload
handling are now separate layers.

The cross-format audit found no new semantic implementation that is both
correctly shared and format-neutral. Existing common helpers remain the
single owners for bounded XML escaping and scanning, MCE preprocessing,
relationship inventories, GUID validation, and opaque payload handling in
`litchi-ooxml-common`, `litchi-ole-common`, and the neutral core. Format
specific media, glossary, validation, and number-style grammars stay in their
own crates; no speculative common abstraction was introduced.

Checked-in Microsoft anchors are `[MS-OE376]` §§2.1.314--2.1.316,
`[MS-PPTX]` §§2.1.1, 2.2.4, 2.3.1.18, and 2.3.3.11--2.3.3.18, and
`[MS-XLSX]` §§2.4.5, 2.4.7, 2.6.3--2.6.5, and 2.7.2. The repository does
not currently contain a checked-in ODF specification snapshot, so the ODF
layering change makes no external conformance claim beyond its existing
parser and regression fixtures.

Focused verification passes with the workspace compiler lints enabled:
254 DOCX unit tests plus three API targets, 196 PPTX unit tests plus one API
target, 395 XLSX unit tests and all examples, and 1,141 ODF unit tests plus
all integration/example targets. The no-default-features `litchi-ooxml`
host suite passes 1,499 unit tests and its integration targets; the crate
boundary audit remains green at 35 packages, 107 internal edges, and 13
scheduled debts. This structural batch makes no native Office, performance,
or full-workspace Clippy claim.

## DOCX fields, PPTX animations, XLSX conditional formatting, and ODF drawing-page layering

The next four large semantic owners are now layered under their existing
public paths:

- `litchi-docx::field/{model,codec,tests}`;
- `litchi-pptx::animations/{model,codec,package,tests}`;
- `litchi-xlsx::conditional_formatting/{model,codec,tests}`; and
- `litchi-odf::drawing_page_properties/{model,codec,package,tests}`.

Models own contextual values, codecs own bounded format XML parsing and
serialization, package layers own relationship/content-type context where it
exists, and tests stay beside the owner. Historical module paths and public
names remain available through focused compatibility aliases; new names do not
repeat the enclosing format or module prefix.

The cross-format audit reused existing neutral helpers for MCE processing,
bounded XML/name handling, escaping, and relationship inventories. DOCX field
codes, PresentationML timing, SpreadsheetML conditional formatting, and ODF
drawing-page vocabularies are different grammars, so no speculative logic was
moved into `litchi-ooxml-common`, `litchi-ole-common`, or another common crate.

The checked-in specification anchors are `[MS-OE376]` §§2.1.501, 2.1.516,
2.1.538, 2.1.543--2.1.551, 2.1.1729, and 2.1.1736; `[MS-PPTX]` §2.2.2 and
its timing-node/behavior structures in §2.3; and `[MS-XLSX]` §§2.2.2.2,
2.4.6, 2.4.24, 2.6.1--2.6.2, 2.6.27--2.6.30, and 2.6.49--2.6.50. The
repository has no checked-in ODF specification snapshot, so the ODF change
makes no external conformance claim beyond its existing parser and fixtures.

Integrated verification passes with workspace compiler lints enabled:
`cargo check` and `cargo test` cover all four crates with all features and
targets; the owner suites report 254 DOCX unit tests plus three API targets,
196 PPTX unit tests plus one API target, 395 XLSX unit tests plus examples,
and 1,554 ODF tests across 85 harnesses. The no-default-features
`litchi-ooxml` host suite also passes 1,499 unit tests and its integration and
example targets. Formatting, staged diff checks, and the workspace crate
boundary audit remain green.

## DOCX web settings, PPTX tags, XLSX timelines, and ODF graphic-property layering

This slice layers four semantic owners beneath their established
public paths:

- `litchi-docx::web/{model,codec,package,tests}`;
- `litchi-pptx::tag/{model,codec,package,tests}` while retaining `tag::raw` and
  `tag::shape` as focused submodules;
- `litchi-xlsx::timelines/{model,codec,package,tests}`; and
- `litchi-odf::graphic_properties/{model,codec,package,tests}`.

Models own contextual values, codecs own bounded XML conversion, package
layers own OPC/flat-document graph context, and tests remain beside the
owner. Historical module paths and public declarations remain available via
compatibility aliases; canonical names do not repeat their enclosing owner
prefix.

The shared-logic audit reuses existing neutral MCE, XML escaping, GUID, and
relationship helpers from `litchi-ooxml-common`, `litchi-opc`, and
`litchi-core`. Web-settings, PresentationML tags, SpreadsheetML timelines,
and ODF graphic-property vocabularies are not one grammar, so no speculative
document-model implementation was moved into a common crate. The workspace
now also contains `litchi-iwa-common`, which owns dependency-neutral bounded
IWA varint and wire primitives plus format-independent table-cell vocabulary;
it owns no strokes, appearance, archive/protobuf codecs, package identifiers,
or concrete object-model state. Concrete Pages, Numbers, and Keynote
object-model logic remains in the format crates while the migration proceeds.
The physical IWA substrate is also now owned by its leaf crates:
`litchi-iwa-protos` owns generated raw schemas and `litchi-iwa-core` owns
bounded archive framing plus checksum-free Snappy compression/decompression.
The facade's former duplicate Snappy implementation was deleted; concrete
readers pass compressed slices to the core and borrow decoded bytes through
`as_bytes()`. Application message decoding and package topology remain in
`litchi-iwa`.
The migration exit for this family is deletion of the duplicate facade-local
`wire.rs` and `varint.rs` kernels, with no public compatibility shim. Until
all callers have moved, focused owner adapters may remain private, but they
must preserve bounded-input policy and must not make the common structured
wire errors untyped by accident.
The varint portion of this exit is now complete: every facade caller imports
`litchi-iwa-common::varint`, the duplicate `litchi-iwa::varint` module is gone,
and the common bounded decoder/error type is the sole owner. The remaining
facade-local file is `wire.rs`, now only a private callback/error adapter. Its
parser returns `litchi_iwa_common::wire::WireField` directly, so the facade no
longer copies parsed fields into a second vector; all consumers use the common
typed accessors. Scalar, nested, repeated, and append mutation paths delegate
to the bounded common kernel, including its typed limit/allocation errors.
The callback traversal, fallible reservations, nesting/output limits, and
removal seam now live in `litchi-iwa-common::wire` and are generic over the
caller's error type; shared wire failures convert through `From<common::Error>`.
The facade `wire.rs` retains only thin adapters that infer `litchi-iwa::Error`
for existing callers. Deleting that file is now an import-migration task, not
a second wire implementation.
Canonical chart readers use the common `encoded_len` check instead of
allocating temporary varint vectors, and the reference-line reader decodes
directly from its borrowed payload while mapping malformed values to the typed
format error.

Checked-in anchors are `[MS-OE376]` §§2.1.444--2.1.462 for Word frameset/web
settings behavior and §2.1.1170 for PowerPoint programmable tags; `[MS-XLSX]`
§§2.1.7--2.1.8, 2.3.5, 2.4.49--2.4.58, and 2.6.98--2.6.118 for timeline
parts, relationships, and complex types. The repository has no checked-in
ODF specification snapshot, so the ODF change makes no external conformance
claim beyond its existing parser and fixture coverage.

Verification for this slice passes with `cargo check` and `cargo test` for
the four affected crates, all features and targets, followed by the
no-default-features `litchi-ooxml` host suite. Formatting, diff checks, and
the workspace crate-boundary audit also pass.

The Keynote soundtrack semantic migration moves playback values into the
archive-free `litchi_keynote::soundtrack::{Mode, Settings}` module. The IWA
adapter alone retains `KN.Soundtrack` decoding, object-graph and package-ID
selection, optional protobuf presence, unknown-field and media-reference
preservation, and failure-atomic edits. The focused gates are the semantic
crate's discriminant/volume validation tests and the IWA fixture tests for
unknown-field preservation, native media-reference stability, malformed graph
rejection, no-op byte stability, and transactional rollback. The old
`KeynoteSoundtrackMode` and `KeynoteSoundtrackSettings` owners and aliases are
deleted rather than retained as compatibility shims.

## OOXML common relationships, ODF common vocabulary, and owner migration

This slice makes the common/format-specific boundary explicit:

- `litchi-odf-common` owns ODF constants, coordinates, datatype vocabulary,
  and detection; `litchi-odf` remains a thin contextual detection facade that
  re-exports the established detector paths. Format semantics stay in the
  concrete ODT, ODS, ODP, and smaller ODF-family crates.
- `litchi-ooxml-common::relationships` owns Transitional/Strict relationship
  attribute decoding, including unresolved `r:` fragments. OOXML hosts keep
  their format-specific error and facade layers.
- `litchi-pptx::actions`, `litchi-xlsx::header_footer`,
  `litchi-xlsb::named_ranges`, and `litchi-odf::font_face` now use layered
  `{model,codec,package,tests}` owner folders. Canonical types are contextual
  and prefix-free; the owner facades expose only those canonical types.

The checked-in specification anchors for this batch are `[MS-PPTX]` §3.4 for
slide-show action references and `[MS-XLSB]` §§2.4.718 and 2.5.73 for defined
names and header/footer strings. The repository has no checked-in ODF
specification snapshot, so the ODF extraction makes no external conformance
claim. ADR 0009 continues to keep ODF detection in `litchi-odf-common`, and
ADR 0010 continues to keep archive grammar below the public facade.

Verification passes for the affected common and owner crates with all features
and targets, the no-default-features OOXML host suite, formatting, diff
checks, and the 36-package crate-boundary audit. Full workspace all-features
verification remains environment-limited by the existing native fontconfig
dependency, and broad strict-Clippy status is not claimed.

## ODF common package and manifest seam

The next ODF extraction is intentionally bounded under
`litchi-odf-common::package`:

- `Archive` owns borrowed ZIP access, manifest-location lookup, and the
  neutral archive operations used by every ODF document family;
- `Manifest` and `Entry` own only manifest file paths, media types, sizes, and
  neutral XML validation; and
- media-path classification is shared without importing any document-family
  model.

The `litchi-odf` manifest layer converts the common neutral model into its
encrypted-entry overlay and continues to own encryption metadata validation,
password decryption/authoring, digital signatures, `OwnedPackage`, and
document-family package orchestration. Encryption child elements are therefore
recognized by the common XML traversal but interpreted only by the format
crate. The remaining package seam is deliberate: moving signatures, password
state, or family orchestration would create an upward dependency or make the
common crate format-aware.

The layered common layout is
`litchi-odf-common/src/package/{mod,model,codec,tests}.rs`; canonical common
names are `Archive`, `Manifest`, and `Entry`, with manifest paths stored as
map keys so each path is allocated once.
No checked-in ODF specification snapshot is available, so this extraction
makes no new external conformance claim beyond the existing parser fixtures.

## Owner-only API convergence

The migration now removes compatibility-only aliases and duplicate host model
wrappers for the touched seams. DOCX numbering, PPTX and DOCX modern comments,
PPTX actions, XLSX header/footer, XLSB PivotTable views, ODF font faces, and
ODF datatypes are consumed through their canonical owner types. OOXML host
package methods may still map owner errors or resolve OPC relationships, but
they no longer invent a second semantic type or a prefix-expanded alias.
Callers that need a format-neutral model use the owning common crate directly.

For XLSB conditional formatting, the owner-only boundary is
`litchi-xlsb::conditional_formatting/{mod,model,codec,tests}.rs`. Its
canonical facade exposes `Formatting`, `Rule`, `RuleType`, `Value`, `Scale`,
`Bar`, `IconSet`, and the related record values without compatibility aliases
or format-expanded spellings. The OOXML migration host no longer publishes a
conditional-formatting forwarding module or writer codec; it retains worksheet
record orchestration and maps owner failures into `xlsb::Error` at the host
error boundary.

## ODF namespace vocabulary extraction

This bounded ODF common-vocabulary slice moves the generic XML namespace
model from `litchi-odf::elements::namespace` to the canonical
`litchi-odf-common::namespace` owner. `QualifiedName` resolves expanded names
and `NamespaceContext` carries prefix/default-namespace bindings; neither type
depends on package, manifest, encryption, or a document-family model.

`litchi-odf::elements::element` consumes those common types directly, while
the old nested host module is removed rather than retained as a compatibility
alias. The ODF facade exposes the common module at its short `namespace` path,
and the canonical names remain prefix-free. Namespace constants, mappings, and
focused model tests now live with the common owner. This extraction makes no
new ODF conformance claim and does not change package or manifest ownership.

## ODF media, metadata, and common package-path layering

The ODF owner now has two additional semantic folders. Image discovery and
safe source classification live under
`litchi-odf::media/{mod,model,codec}.rs` with canonical `Image`, `ImageFrame`,
`ImagePart`, and `ImageSource` values. The old `OdfImage*` spellings were
removed; packaged, flat, embedded-object, and spreadsheet consumers use the
same owner model without duplicate wrappers.

The ODF-neutral archive path rules live under
`litchi-odf-common::package::path`, alongside the archive and manifest owner.
`is_linked_href` and `resolve_package_path` are shared by image, embedded
object, RDF, and script package owners, so path normalization and traversal
protection are implemented once. The format crate retains family-specific
media classification and package orchestration.

Core metadata is layered as
`litchi-odf::core::metadata/{mod,model,codec,tests}.rs`. `Metadata` is the
canonical contextual value; the prior `OdfMetadata` spelling is not retained.
The model owns conversion to `litchi_core::Metadata`, while the codec owns
bounded `meta.xml` parsing, deterministic serialization, and source patching.
Package and document-family consumers now use that owner directly.

Focused verification for this slice passes 1,097 ODF unit tests, the image and
embedded-package integration targets, and the ODF common-package suite. No
new ODF specification or native Office claim is made by these structural
refactors.

## ODF chart read and semantic layering

The ODF chart read boundary is now shared by `litchi-odc` and ODT embedded
charts through `litchi-odf-common::chart`. Its layered modules own the bounded
namespace-aware `reader`, scalar chart vocabulary under `axis`, `grid`,
`legend`, and `plot_area`, and zero-copy semantic `view` types. `Attribute`,
`Element`, and `Kind` retain expanded names, unknown namespaces, extension
subtrees, and inert text while rejecting duplicate expanded attributes,
unsupported entities, malformed roots, excessive depth, element counts, and
text/attribute sizes. The common crate has no package-family dependency.

ODC now validates and exposes this retained chart model from its concise
`Chart` facade; its incomplete duplicate axis/legend/series models were
removed. ODT’s standalone chart wrapper, authoring, and mutation code consume
the same common reader and views, while ODT-specific `Object_N/` package
paths, manifest topology, inline roots, and host placement remain in the ODT
owner. Chart authoring and mutation are intentionally separate follow-up
layers, so this extraction does not claim complete ODF chart conformance.

The common chart reader and semantic tests, ODC all-target tests, and the full
ODT all-target suite (530 library tests plus integration targets) pass, along
with affected-crate checks and formatting. This slice also removes the former
ODT-local chart reader/semantic duplicate without compatibility aliases.

## ODF chart authoring ownership

Typed ODF chart authoring now belongs to the standalone `litchi-odc` owner.
Its layered `authoring::{model,data,extensions,writer,builder}` modules expose
the contextual `Definition`, axis/series/plot specifications, cached-table
values, inert calculation settings, retained extension trees, and bounded
deterministic `serialize_content`/fragment writers. `Builder` consumes a typed
definition, while `Chart::from_definition` provides the concise package
facade. The implementation reuses `litchi-odf-common::chart` vocabulary and
reader validation without introducing a dependency on any host family.

ODT now imports that authoring owner for standalone and embedded charts. ODT
retains only its chart-document mutation and package mechanics: `Object_N/`
paths, manifest entries, inline-root conversion, content replacement, and
text/sheet/page host placement remain local. No compatibility aliases or
duplicate authoring model are retained, and mutation extraction remains a
separate follow-up so this slice does not overclaim chart completeness.

Post-migration checks pass for ODC all targets (6 library tests plus its
integration targets), ODT library coverage (527 tests), focused ODT chart
mutation tests, formatting, and affected-crate compilation. A concurrent
attempt to link every ODT integration target exhausted the generated build
volume; this is an environment resource limit, not a source or test failure.

## ODF inert drawing-resource inventory

The ODF common boundary now also owns read-only inventory for resources that
can be embedded by multiple document families. `litchi-odf-common::drawing`
retains frame geometry and host placement metadata, while `embedded` and
`media` provide bounded package/flat-document scanners with contextual
`Object` and `Image` models. Their `Source` values distinguish inline data,
package parts, links, and missing resources without interpreting host-specific
mutation semantics. The package abstraction is the small `PackageLookup`
trait, so common scanners do not depend on an ODT package implementation.

ODT now exposes these common inventories through its ergonomic document
facades. ODT-specific package placement, cleanup overlays, object paths,
manifest mutation, and byte replacement remain in the ODT owner. ODS now
activates the same inventory through layered `drawing`, `media`, and
`embedded` facades, including safe package-local image extraction; the other
ODF families can adopt the scanners without inheriting ODT mechanics. No
compatibility aliases or duplicated scanner model are kept, and
authoring/mutation remains separate from this read-only inventory layer.

The format-neutral ODT facade also removes redundant `OpenDocument` prefixes:
`generic::{Family,Package,FlatDocument}` are the contextual owners, with all
internal and exercised deferred consumers migrated to those names. This
keeps the public path ergonomic while preserving the layered module context.

The common ODF all-target suite (178 unit tests plus integration targets),
ODT library suite (525 tests), boundary policy checks, formatting, and
affected-crate compilation pass; the ODS all-target suite passes 73 tests.
This is a structural inventory extraction; it makes no new ODF conformance or
native Office interoperability claim.

## ODF common lexical and family validation

Neutral bounded lexical contracts now live in
`litchi-odf-common::datatype::lexical`: finite numbers, `#RRGGBB` colors, and
caller-owned byte limits are shared by ODS conditional formats and sparklines.
Feature-specific limits and diagnostic contexts remain in ODS. The common
`core::validate_content_part` contract likewise centralizes content-size and
family-body-marker validation for detached builders; ODG consumes it while
retaining its own MIME and package facade.

No duplicate helper bodies or compatibility aliases remain. The affected
common, ODS, and ODG suites, formatting, and boundary checks pass.

## ODF annotation vocabulary layering

The shared `litchi-odf-common::annotation` module now exposes contextual
vocabulary directly as `Annotation`, `Element`, `Node`, and `Builder`. The
common annotation owns lossless mixed-content and typed metadata; ODT keeps
position, host scanning, and package mutation in its annotation facade, while
ODS consumers use the same neutral tree for rich cell content. The former
`Annotation*` and `CellAnnotation` names were removed without compatibility
aliases.

The common all-target suite (182 unit tests plus integration targets), ODT
library suite (525 tests), formatting, and boundary checks pass.

## ODF family package layering

The shared packaged-family owner is now `litchi-odf-common::core::family::Package`.
Each ODF family keeps its own contextual `package` facade and validation
policy; ODS refers to the shared owner through `core::family::Package` so its
public `package::Package` remains unambiguous. The former
`FamilyPackage` prefix was removed without a compatibility alias across ODB,
ODC, ODG, ODI, ODM, ODP, ODS, and OTH.

All affected family crates and the common crate pass formatting and all-target
compilation; the focused common, ODS, and ODG tests remain green.

## Shared BIFF framing and legacy binary owner migration

The legacy binary migration now has a neutral physical-record owner:
`litchi-biff` implements the four-byte BIFF frame from `[MS-XLS]` §2.1.4 and
`[MS-OGRAPH]` §2.1.4 (`u16` kind, `u16` payload length, and bounded payload
bytes). Its `RecordRef`, `Records`, `Record`, `Encoder`, and `Limits` APIs are
borrowed-first, lossless, allocation-bounded, and intentionally do not
interpret continuation records, chart grammar, workbook topology, or host
metadata. BIFF12 remains separately owned by `litchi-xlsb`.

`litchi-ograph` and `litchi-xls` now depend on this substrate directly. The
former local OGraph frame module and the XLS chart/workbook frame duplicates
are removed; chart, package, record, worksheet, and writer layers retain their
format-specific semantics and map shared framing failures into their own typed
errors. Unknown record kinds and exact encoded bytes remain available at the
neutral boundary for lossless higher-level handling.

The same binary migration batch also layers the remaining touched semantic
owners: DOC page borders are under
`litchi-doc::doc::section::borders::{model,codec}`, PPT ExOle objects and
references are under `litchi-ppt::embedded::{object,reference}`, and XLS
worksheet views remain under `view::{model,codec}`. Canonical names are
contextual and prefix-free; the removed `SectionPageBorder*`,
`PowerPointOle*`, `PowerPointExternalObject*`, and `XlsView*` spellings are not
retained as compatibility aliases.

Focused verification covers 15 `litchi-biff`, 40 `litchi-ograph`, 837
`litchi-doc`, 879 `litchi-ppt`, and 837 `litchi-xls` library tests, plus the
all-target suites for the new BIFF and legacy owners. `cargo check` covers all
five crates with all features and targets; formatting, metadata, and the
boundary policy pass for 46 workspace packages and 150 internal dependency
declarations with no explicit migration debts. DOC fixture targets that read
the checked-in corrupted POI OLE sample remain an external fixture failure,
not a regression in this slice.

The target-driven OLE object boundary is recorded below: neutral CFB and OLEDS
capture retains host metadata opaquely while DOC and XLS interpret their own
`ObjectPool`/`MBD`/`LNK` references. Additional legacy semantic owners,
including XLS layout rows and columns, remain format-specific follow-up work and
must not be folded into `litchi-biff`.

## Target-driven OLE object ownership

The OLE boundary is now layered as
`litchi-ole-common::object::{target,model,codec,discovery,editor}`. The common
crate accepts explicit `Target`/`Targets` values and owns only bounded CFB
capture, opaque `Storage`/`Stream` views, shared stream allocations, and
transactional rendering. It no longer infers DOC/XLS storage names or models
`Format`, `ObjInfo`, `CompObj`, native payloads, previews, or host object kinds.
This follows `[MS-CFB]` directory/storage rules and keeps the inert
`[MS-OLEDS]` object streams unactivated and unresolved.

DOC derives `ObjectPool/_<decimal-id>` targets from its own field/storage
semantics in `[MS-DOC]` sections 2.1.4 and 2.6.1, and interprets the opaque
`\u{003}ObjInfo` stream as the typed `doc::embedded_object::Info` described by
`[MS-DOC]` section 2.9.165. It also handles creating the first `ObjectPool`
storage without leaking that topology into the common owner. XLS reads its
`Workbook`/`Book` stream before common capture, derives deduplicated `MBD` and
`LNK` targets from `Obj`/`FtPictFmla` records, and retains `MBD`/`LNK` semantics
in the XLS owner as required by `[MS-XLS]` sections 2.1.7.5, 2.1.7.7, and
2.5.150–2.5.151. Chart-only XLS editing passes an intentionally empty target
catalog.

The common object suite passes 18 tests, DOC all-feature library coverage
passes 840 tests with two ignored, and XLS all-target coverage passes. The
focused DOC and XLS OLE suites cover absent `ObjectPool`, `ObjInfo`, MBD/LNK,
deduplication, shared-storage removal, and bounded workbook reads. Formatting,
metadata, and the boundary policy continue to pass for 46 workspace packages,
150 internal dependency declarations, and zero explicit migration debts. The
full DOC target suite still reports the pre-existing corrupted Apache POI OLE
fixture (`FAT entry 52`), which is external fixture data rather than a failure
in the target migration.

The next high-value binary seam is the remaining host-specific CFB signature
coverage and legacy layout ownership; new shared logic should be extracted only
after its format-neutral vocabulary and specification boundary are identified.

## DOCX section-border layering

The DOCX writer now moves its section implementation into
`writer::section::{mod,borders}.rs`. The page-border owner exposes contextual
`section::borders::{Border, Borders, Art, Color, Display, OffsetFrom, Style,
ZOrder}` values, while `section::SectionProperties` retains the ergonomic
`page_borders` field. The former `SectionPageBorder*` names and the flat
prefix-expanded facade were removed without aliases. The XML codec remains
responsible for `[ECMA-376]` `CT_Border`/`CT_PageBorders` (§17.6.16) parsing,
validation, deterministic writing, and lossless round-trip behavior.

DOCX default all-target coverage passes 652 library tests plus its integration
and example targets; the focused section suite passes 14 tests. Formatting
passes. Enabling the optional fonts feature is currently environment-blocked by
the missing system `pkg-config`/Fontconfig tool, unrelated to this module move;
the standalone default feature set remains verified.

## DOCX contextual enum ownership

The compiled DOCX enum facade now follows the same contextual rule. Section
layout owns `section::{Orientation, Start}`, header/footer semantics own
`header_footer::Kind`, and style semantics own `styles::Type`. The former
`WdOrientation`, `WdSectionStart`, `WdHeaderFooter`, and `WdStyleType` names,
the flat `enums` module, and root compatibility reexports were removed. XML
lexemes, numeric representations, defaults, and display behavior are unchanged;
only the semantic owner and ergonomic path changed.

DOCX library and integration coverage passes after this migration, and the
umbrella DOCX example now consumes the canonical `ooxml::common::Props` and
contextual enum paths. The next compiled prefix boundary is the nested
`Chart*` vocabulary in `litchi-drawingml::chart`; its host consumers should
continue to depend only on the neutral chart owner.

## DrawingML chart vocabulary layering

The shared DrawingML chart owner now follows its semantic module context. The
`litchi-drawingml::chart` facade retains the root `Chart` value and exposes
`HeaderFooter`, `PageMargins`, `PageOrientation`, `PageSetup`, `PrintSettings`,
`Protection`, `ExternalData`, `UserShapes`, `ShapeProperties`,
`TextProperties`, `ExtensionList`, `Lines`, and `types::Type` without repeated
`Chart` prefixes. These values model the chart schema structures described by
the checked-in `[MS-ODRAWXML]` chart schemas and the `[MS-OI29500]` chart
conformance material; host-specific package resources such as
`ChartExternalDataPart` remain in their XLSX/XLSB owners because their
relationship and storage semantics are not neutral.

The reader, writer, model, axis, legend, plot-area, series, and type modules
were migrated together, so XML names, defaults, validation, and lossless
extension handling are unchanged. XLSX and XLSB consumers now import the
neutral contextual vocabulary directly; DOCX and PPTX required no shared-model
adaptation. No compatibility aliases were retained.

Focused verification passes 38 DrawingML chart tests, 48 XLSX tests, 21 XLSB
tests, and the affected default-feature DOCX/PPTX all-target suites (704 and
441 tests respectively). Formatting and affected-crate checks pass. The
all-feature aggregate remains environment-blocked by the missing system
`pkg-config`/Fontconfig dependency, as recorded above.

## DrawingML text-body vocabulary layering

The first SpreadsheetDrawing text extraction moves the neutral `a:CT_TextBody`
vocabulary into `litchi-drawingml::text::body::{Body,Properties,Insets,
Paragraph,Run}`. The model owns DrawingML defaults, text insets, paragraph
joining, run formatting values, and body properties; the duplicated XLSX and
XLSB model structs and repeated `XlsxText*` spellings are removed. This follows
the shared text-body structures in the checked-in `[MS-ODRAWXML]` and
`[MS-PPTX]` specifications.

XLSX and XLSB retain their SpreadsheetDrawing anchors, bounded XML state
machines, worksheet/BrtDrawing package wiring, and authoring emission around
the common model. The shared `text::body::writer` now owns neutral body,
property, paragraph, and run XML emission; host writers only supply their
package-specific shape context. XML parser extraction and PPTX `p:txBody`
wrappers remain separate follow-up layers, so this slice does not claim
complete DrawingML text conformance. No compatibility aliases were retained.

Focused verification passes 89 DrawingML tests, 621 XLSX library/integration
tests, 408 XLSB library/integration tests, formatting, and boundary policy
checks.

## ODF master-page and authoring layering

The ODF style and authoring seam now has a neutral owner in
`litchi-odf-common`. `style::master::{Master,Child,ChildKind,Region,Kind}`
provides a bounded, namespace-aware, lossless model for master-page children
and header/footer regions, including ordered inert XML and typed text fields.
Its reader and writer validate the ODF master-page structure and apply
region/master edits without moving package orchestration into the common
crate. ODT retains the contextual XML element and package mutation facade.

The same rule applies to `drawing::authoring::{Anchor,Frame,Length}` and
`media::authoring::{Format,Part}`. Geometry, anchor semantics, image sniffing,
safe part paths, and bounded payload ownership are shared; ODT retains only
ODT element construction and package insertion. The former flat frame model
and prefix-expanded names were removed without compatibility aliases, and
the ODT facade is now layered as `frame::{mod,xml}`.

The common all-target suite passes 206 unit tests plus its integration targets;
ODT passes 501 library tests plus its integration targets. Formatting,
metadata, diff, and boundary checks pass for the affected crates. This slice
does not claim complete ODF master-page, drawing, or media conformance; it
establishes the shared semantic and authoring foundations for ODP, ODS, and
the remaining ODF owners.

## OLE Property Set version layering

`litchi-cfb::metadata` now models the versioned type constraints of
`[MS-OLEPS]` rather than accepting only version-zero property sets. Version
one permits the documented `VT_ARRAY|VT_I1` and `Behavior` special property
forms; `Behavior` remains a typed `VT_UI4` value with only the specified
values, and version-zero input is rejected when a version-one-only type is
used. The property-set editor keeps validation failure-atomic and preserves
unknown properties and exact source bytes at the existing snapshot boundary.

This is a focused typed-object-model increment, not a claim that every
property type or OLE metadata profile is complete. The CFB library suite (109
tests plus four integration targets) and all-target compilation pass.

## Legacy DOC facade and OLE object preservation

The outer `litchi-doc/src/doc` wrapper has been removed. Root declarations now
expose the semantic owners directly, while `parts`, `section`, and `writer`
remain nested where they represent real document layers. Internal imports and
tests use the canonical paths; no wrapper or compatibility alias remains.

The DOC, PPT, and XLS object-list paths also retain unsupported binary records
at their owning boundaries. DOC keeps its OLE package topology and typed
`ObjectPool` interpretation in the DOC owner; common CFB capture stays
format-neutral. The DOC suite passes 831 library tests with two ignored tests;
one checked-in malformed Apache POI fixture still fails with an invalid FAT
entry and is recorded as external fixture debt rather than hidden.

## Legacy XLS, PPT, and OfficeArt typed OLE layering

The legacy host facades now use contextual, prefix-free owners. XLS exposes
`ole_object::{Editor,OleObjectRecord,ObjectType,Ft*,Lbs*}` with strict
`[MS-XLS]` validation for OLE flags, form-control records, dropdown bounds,
and list-item sizes while retaining malformed/unknown records losslessly.
PPT's `embedded::object::UnknownRecord` preserves unmodeled `ExObjList`
children, source ordering slots, and borrowed payloads through collection
edits. OfficeArt exposes typed `shape::Bounds` for the exact 16-byte `FSPGR`
group-coordinate record, preserving unknown records around it. These seams
are host-specific and therefore remain outside `litchi-biff` and
`litchi-ole-common` until a neutral vocabulary is proven from the relevant
`[MS-ODRAW]`, `[MS-OGRAPH]`, `[MS-OSHARED]`, `[MS-PPT]`, and `[MS-XLS]`
grammars.

Focused verification passes the XLS OLE/form-control tests (18 and 4), PPT
object tests (15) plus its 881-test library suite, and 59 OfficeArt shape
tests; affected all-target checks and formatting pass. This is incremental
read/preserve/edit coverage, not a claim of complete legacy Office
conformance.

## Legacy shape, worksheet-layout, and external-media owner layering

The next legacy slice continues the same contextual-owner rule. DOC now
exposes `shape::{Shape,Kind,Bounds,UnknownRecord}` in a layered
`shape/{mod,model,codec}` owner. The codec retains Word's FIB/table-stream
story placement, while the model owns the owned OfficeArt shape tree, typed
`FSPGR` bounds, and exact unknown record bytes. This follows the Word drawing
and shape rules in `[MS-DOC]` and `[MS-ODRAW]`; the DOC facade does not expose
the former `DocShape*` or `DocDrawingShape` names.

XLS now exposes `layout::{Row,Column}` and
`worksheet::layout::Layout`. `layout/{mod,row,column}` owns BIFF8 `ROW` and
`COLINFO` semantics, while worksheet layout owns `GUTS`, `WSBOOL`,
`DefRowHeight`, and `DefColWidth` state. Record ordering, reserved bits,
coordinate sentinels, and format references are checked against `[MS-XLS]`;
the former `XlsRowLayout`, `XlsColumnLayout`, and `XlsWorksheetLayout` names
were removed without aliases.

PPT external media is layered as
`external_media::{model,codec,tests}` with contextual `Media`, `Video`,
`Movie`, `LinkedAudio`, `EmbeddedWav`, `CdAudio`, `Collection`, and `Object`
values. Strict `[MS-PPT]` codecs validate ExMedia, ExVideo, AVI/MCI, linked
audio, embedded WAV, and CD-audio records; paths remain inert, reserved bytes
round-trip, and unmodeled ExObjList children remain as `UnknownRecord` data.
The former `PowerPointExternal*` and `PowerPointLinkedAudio*` names were
removed without compatibility aliases.

Verification passes DOC's 831 library tests (829 passed, 2 ignored), XLS's
841-test all-target suite, PPT's 882 library tests (881 passed, 1 ignored)
plus integration targets, and the affected ODF/CFB/OfficeArt suites. Combined
all-target compilation, workspace formatting, metadata, diff, and boundary
policy checks pass. These additions improve typed extraction and lossless
preservation but do not claim complete legacy Office conformance.

## Legacy OLE and OOXML owner migration

The following batch continues the same breaking, contextual-owner migration
through the OLE2 and OOXML verticals. Word OLE-control metadata is now owned by
`parts::ole::controls::{Control,Controls}` in the layered
`parts/ole/controls/{model,codec,tests}` tree. Its codec retains the inert
`RgxOcxInfo` cookies, padded `OcxInfo` strides, uniqueness checks, and FIB
range validation from `[MS-DOC]` 2.9.161 and 2.9.229; the old
`DocumentOleControls` and `OleControlInfo` names were removed without aliases.

Legacy chart FRT records now belong to `litchi-xls::chart::frt`, with
`info`, `label`, `blocks`, `wrapper`, and `continuation` ownership. The
contextual `Info`, `Version`, `RecordRange`, `CatLab`, `StartBlock`,
`EndBlock`, `StartObject`, `EndObject`, `Wrapper`, and `CrtMlFrt` values retain
reserved bytes, FRT-header checks, version-dependent ranges, and continuation
chains specified by `[MS-XLS]`. PPT animation now exposes contextual
`animation::{Editor,Scope,Timeline,Hash10,LinkedSlide,SlideTime,Flags}` values
and removes the redundant `PowerPointAnimation*`/`PowerPoint10Slide*` public
prefixes while retaining the typed parser/writer and resource limits.

The shared OLE Custom XML datastore moved from the flat
`custom_xml_data` module to `litchi-ole-common::custom_xml`, exposing
`Store`, `Item`, `Properties`, `ItemId`, `RootName`, `Promotion`, and bounded
`inspect`/`write` operations. The OLE2 model remains inert: GUID and UTF-16
validation, promotion-marker rules, schema-reference retention, XML payload
validation, and allocation limits are shared by DOC, PPT, XLS, and the
crypto/DataSpaces graph. Downstream callers were migrated directly; no old
module facade remains. OOXML custom properties likewise moved to the layered
`litchi-ooxml-common::custom::{model,codec,package,schema}` owner.

DOCX drawing inventory is layered under `drawing::{model,codec}` with
contextual `Object`, `Kind`, and `Anchor` values. XLSX SpreadsheetDrawing,
ordinary chart integration, and pivot-chart integration now have separate
layered owners. Their namespace-aware codecs reuse `litchi-drawingml`, retain
unknown/inert relationship payloads, enforce worksheet/chart resource limits,
and distinguish the host anchor/package graph from shared chart and text
models. Pivot-chart `Source`, `Binding`, `Options`, `Series`, and sheet
inventory values no longer carry redundant `PivotChart` prefixes.

The affected owners pass combined all-target compilation. Focused suites pass
24 OLE-common unit tests plus its integration targets, 163 OOXML-common unit
tests plus integration targets, 654 DOCX tests, 881 PPT library tests plus
integration targets, 842 XLS tests plus integration targets, and 635 XLSX
tests plus integration targets. DOC's 829 library tests pass (two ignored);
one existing Apache POI integration fixture still fails before parsing because
its FAT entry 52 is beyond the physical file. This batch establishes typed,
bounded, lossless owner seams; it does not claim complete `[MS-DOC]`,
`[MS-ODRAW]`, `[MS-OGRAPH]`, `[MS-OSHARED]`, `[MS-PPT]`, or `[MS-XLS]`
conformance.

## Layered legacy and DrawingML owner continuation

The next migration batch completes another set of flat-owner removals while
keeping format context at the facade. DOCX section/page-layout semantics now
live under `section::{model,codec,tests}` with typed orientation, break,
measurement, margin, column, and header/footer-reference values. XLSX
what-if scenarios now live under `scenarios::{model,codec,tests}` with
contextual `CellReference`, `RangeReference`, and `Collection` values.
Both codecs keep bounded parsing, checked values, unknown markup, and source
ordering at the owner boundary; no redundant prefix aliases remain.

The same layered rule now covers PPT comments, XLS chart core records,
DOC subdocument stories, PPTX slides, DrawingML text primitives, OLE smart
tags, XLSB calculation properties, and OOXML embedded relationship
inventories. Each owner separates its semantic model from XML/binary codec
and regression tests, while shared DrawingML, OOXML, and OLE vocabulary stays
in the corresponding common crate. Host package orchestration remains in the
format crate, so peer formats do not become dependencies of one another.

Focused all-target checks passed for the affected common, legacy, DOCX, PPTX,
PPT, XLS, XLSB, and XLSX crates. The DOCX suite passed 655 unit tests plus
integration targets; the XLSX suite passed 636 tests plus integration targets;
the DOC library suite passed 829 tests with two ignored. Formatting,
metadata, diff, and boundary-policy checks also pass. This is a structural
and typed-object-model increment, not a claim of complete OLE2, OOXML,
DrawingML, or legacy Office conformance; the malformed Apache POI DOC fixture
remains the previously recorded external integration debt.

## Layered OLE2 and OOXML owner continuation

This turn extends the same breaking, prefix-free owner pattern across the OLE2
and OOXML migration hosts. Shared OLE2 property-set grammar now lives in
`litchi-ole-common::property_set::{model,codec,tests}` with the contextual
`Metadata`, `Stream`, `Section`, `Value`, `Standard`, and `Editor` facade;
CFB remains the container owner. DOCX header/footer stories,
XLS BIFF8 comments, XLSX auto-filters, and XLSX external links now each have
separate semantic model, bounded codec, package/orchestration, and regression
layers. Their public values no longer repeat the owner name, and direct callers
were migrated without compatibility aliases.

The migrated XLSX filter owner keeps unknown attributes/elements and source
ordering when it is meaningful, but normalizes canonical known-child ordering
out of the semantic snapshot. It accepts Office-produced sort-condition ranges
whose scope differs from `sortState@ref` while retaining checked range geometry.
The shared DrawingML chart reader now handles the normative formula-plus-cache
shape of series titles (`strRef` with `strCache`) and typed rich titles; the
DOCX POI/LibreOffice chart round-trip fixture passes again. The OOXML MCE owner
is layered under `litchi-ooxml-common::mce::{model,codec,tests}` and all its
callers use the shared common facade.

The affected crates pass combined all-target compilation. Focused tests pass
for CFB, DOCX, OOXML-common, PPT, XLS, XLSB, and XLSX; XLSX passes 640 library
tests plus all integration targets, DOC passes 830 library tests with two
ignored, and the DOCX chart fixture plus its 643-test library suite pass. The
full DOC all-target command still reports only the previously recorded malformed
Apache POI fixture (`FAT entry 52 beyond the physical file is not FREESECT`).
Formatting, metadata, diff, and 46-package boundary-policy checks pass. This
batch is a typed, bounded, lossless structural increment, not a claim of full
`[MS-DOC]`, `[MS-ODRAW]`, `[MS-OGRAPH]`, `[MS-OSHARED]`, `[MS-PPT]`, `[MS-XLS]`,
or OOXML conformance.

## Layered ODF and legacy owner continuation

ODF text owners now use semantic folders for bibliography configuration, chart
properties, content validation, document scripts, dynamic text, header/footer,
notes configuration, tracked changes, and variable declarations. Shared ODF
annotation, coordinate, and datatype owners are layered in `litchi-odf-common`.
The same prefix-free facade rule now covers DOC encryption and SPRM operations,
PPT view-set information, XLS data tables, and the existing PPT/XLSX owner
migrations. Each owner separates `model`, `codec`, package integration, and
tests; shared DrawingML wrappers remain in the common DrawingML/XML layers.

The XLSX shape facade additionally retains bounded unknown anchor objects and
DrawingML markup, parses nested groups and OLE graphic frames, validates OLE
aspects, and keeps picture/chart frames out of the typed inventory. Combined
all-target compilation passed. Focused all-target suites passed for ODF-common,
ODT, OOXML-common, PPT, XLS, and XLSX; DOC passed 832 library tests with two
ignored. The full DOC all-target run remains limited only by the previously
recorded malformed Apache POI FAT fixture. This is a migration and typed-model
increment, not a claim of complete Microsoft Office specification conformance.

## Layered ODF/OLE2 facade continuation

The ODF family slice continues the breaking split through `litchi-odt`,
`litchi-ods`, and `litchi-odp`, while ODT’s smaller owners use semantic
`model`/`codec`/`tests` folders instead of accumulating new flat modules. Ruby
annotations, forms, settings, DDE connections, outline styles, list-label
alignment, and related metadata are exposed through contextual owner facades;
redundant public prefixes and compatibility aliases were removed.

The shared OLE2 property-set reader/editor is now owned by
`litchi-ole-common`, and DOC/PPT/XLS call it directly. The umbrella `litchi`
crate no longer wraps the standalone `litchi-doc` package in an outer `doc`
module; internal unified-document code refers to the package directly and
format-specific users depend on `litchi-doc` itself. DOC header/footer and note
PLC writers are wired through checked FIB pointers, and the CFB container keeps
property sets outside its format-neutral core.

Focused checks passed for ODT, OLE-common, CFB, DOC, PPT, XLS, and XLSX,
including the ODT all-target suite, OLE property-set tests, CFB’s 96-test
suite, DOC writer/property-set integration tests, and PPT/XLS/XLSX all-target
compilation. These checks establish the current typed/layered migration
boundary; they do not claim complete `[MS-DOC]`, `[MS-ODRAW]`, `[MS-OGRAPH]`,
`[MS-OSHARED]`, `[MS-PPT]`, or `[MS-XLS]` conformance.

## Layered OLE2 host-owner continuation

The next host-owner cut removes six more high-value flat modules while keeping
their existing crate facades. DOC now layers section semantics under
`section::{model,borders,columns}`, paragraph content under
`paragraph::{model,tests}`, and OLE package orchestration under
`package::{model,codec,tests}`. Section value types use contextual names such
as `BreakKind`, `PageLayout`, and `TextFlow`; redundant `Section` prefixes and
compatibility aliases are not retained.

PPT now layers document-tail validation, host-specific OfficeArt projections,
and OLE package orchestration under `document_structure`, `odraw`, and
`package` facades. The shared `[MS-ODRAW]` grammar remains owned by
`litchi-odraw`; this cut only separates PPT models, codecs, package seams, and
tests. XLS now layers the BIFF8 OLE-object owner under
`ole_object::{model,codec,package,tests}`, separating typed object records from
checked binary parsing and CFB rewrite orchestration.

Public `crate::{section,paragraph,package,odraw,ole_object}` paths remain
ergonomic and the moves do not add peer-format dependencies. DOC, PPT, and XLS
all-target checks pass; focused suites pass 832 DOC tests with two ignored,
882 PPT tests with one ignored plus integration targets, and 844 XLS tests
plus integration targets. OLE-common passes its 22 unit, 9 integration, and 4
additional-target tests; CFB passes 96 tests. The umbrella facade passes its
all-target tests, and formatting, diff, and 46-package boundary checks pass.
This remains a typed/layered ownership increment, not a claim of complete
`[MS-DOC]`, `[MS-ODRAW]`, `[MS-OGRAPH]`, `[MS-OSHARED]`, `[MS-PPT]`, or
`[MS-XLS]` conformance.

## Layered OLE2 and OOXML core-owner continuation

The next breaking structural slice layers the remaining high-value core owners
without adding redundant public prefixes. DOC, PPT, and XLS now expose
`document`, `presentation`, and `workbook` facades backed by separate
`model`, `codec`, `package`, and test modules. DOCX and PPTX package/layout
owners follow the same semantic split. XLSX keeps its existing nested
`worksheet`, `edit`, `data_model`, and `comments` owners while layering the
workbook facade and calculation properties. The common OOXML external-workbook
relationship vocabulary now lives under
`litchi-ooxml-common::external_link::{model,codec}` with a small public facade.

These moves preserve ergonomic crate paths while keeping typed object models,
binary/XML codecs, package graph integration, and verification seams distinct.
They do not add compatibility aliases, expose native identifiers, or claim
complete `[MS-DOC]`, `[MS-ODRAW]`, `[MS-OGRAPH]`, `[MS-OSHARED]`, `[MS-PPT]`,
`[MS-XLS]`, or OOXML conformance.

The common owner, DOC, PPT, XLS, DOCX, and XLSX all-target checks passed,
including 163 common tests, 832 DOC library tests with two ignored, 882 PPT
tests with one ignored, 844 XLS tests, 643 DOCX library tests, and 642 XLSX
library tests plus their integration targets. The umbrella `litchi` crate
passed 162 tests and its integration targets. Focused PPTX tests passed for
the migrated `master_layout` and `package` owners; the full PPTX suite still
has the two previously observed `notes` tests failing on malformed synthetic
XML (`attribute value not closed`), outside this slice. The full DOC all-target
run remains limited by the previously recorded malformed Apache POI fixture
(`FAT entry 52 beyond the physical file is not FREESECT`). Formatting, diff,
and 46-package boundary checks pass.

## Layered OLE2, OOXML, and ODF owner continuation

This slice continues the same breaking migration through the largest remaining
flat owners. DOC field parts, PPT animation parser/types/writer, XLS list
objects, and XLSB workbooks now have contextual facades over separated model,
codec, package, and test seams. DOCX document and paragraph owners, PPTX
presentation, and XLSX transaction/edit owners follow the same structure. The
XLSX worksheet-view values are also isolated behind a small facade and model
test owner. ODT field elements, builder, and mutable authoring owners; ODS XML
content; and DrawingML's OLE property owner use the same nested organization.
The parser-only ODS content owner intentionally has no package layer because it
does not own package assembly.

Existing public owner paths remain intact, while the new folders keep typed
semantic values separate from binary/XML conversion, package graph work, and
tests. The changes do not add compatibility aliases or redundant format
prefixes, and unknown/lossless payload handling remains in the codec seams.

The combined all-target compile passed for DOC, DOCX, ODT, ODS, ODraw, PPT,
PPTX, XLS, XLSB, and XLSX. The reduced post-clean test matrix passed DOC's 832
library tests with two ignored, DOCX's 643, ODT's 512, ODS's 67, ODraw's 59,
PPT's 882 with one ignored, XLS's 844, XLSB's 408, and XLSX's 642. PPTX had
303 passing library tests and the two previously observed `notes` failures;
its four notes CRUD integration failures have the same malformed synthetic XML
(`attribute value not closed`) and do not touch these migrated owners. DOC
integration targets that consume the malformed Apache POI Word fixture likewise
retain the known CFB error (`FAT entry 52 beyond the physical file is not
FREESECT`); other DOC integration targets passed. Formatting, diff, and
46-package boundary checks pass. This remains a structural migration slice,
not a claim of complete Office-specification conformance.

## Layered RTF, DrawingML, OLE2, and ODF continuation

This slice continues the contextual owner migration through the remaining
large codec and writer files. RTF now separates its lexer, parser, writer, and
retained document model into nested `codec`/`model` facades with dedicated
model, codec, and test seams. DrawingML chart reading and writing and OGraph
chart records follow the same hierarchy. DOC writer core, DOCX field models,
ODT parsing, ODS tracked changes, and ODP parsing now separate semantic state,
format codecs, package boundaries, and focused tests. XLS writer core, XLSB
workbook writing, and the XLSX chart-sheet package received the corresponding
layering.

The public owner paths remain ergonomic while implementation files are
organized by responsibility. No compatibility aliases or redundant format
prefixes were introduced. During verification, the PPTX notes serializer also
fixed its namespace-attribute quoting defect; the notes unit suite (12 tests)
and CRUD integration suite (4 tests) now pass.

The affected-crate structural `cargo check --all-targets` matrix passes with
lint capping, and the same matrix passes under the workspace lint policy for
every affected crate. The
lint-capped affected library-test matrix passes, the umbrella `litchi` crate
passes 162 unit tests plus three integration targets, formatting and diff
checks pass, and the crate-boundary policy remains clean. The known malformed
Apache POI DOC integration fixture remains outside this structural slice. This
is migration evidence, not a claim of complete `[MS-DOC]`, `[MS-ODRAW]`,
`[MS-OGRAPH]`, `[MS-OSHARED]`, `[MS-PPT]`, `[MS-XLS]`, RTF, OOXML, or ODF
conformance.

## Layered OLE2 and OOXML facade continuation

This slice continues the breaking, prefix-free migration through the next
large legacy and OOXML owners. PPT writer core and Escher, PPTX ChartEx,
XLSB conditional-formatting and workbook codecs, ODT fields, ODS content,
DOCX document writing, XLSX worksheet snapshot editing, DOC PAP/TAP, and
XLS pivot tables now expose contextual facades over separated semantic model,
binary/XML codec, package, validation, and test seams. The DOC package facade
also uses the unprefixed `OpenOptions`, `EncryptionKind`, and `Error` names;
the section-border error remains explicitly `BorderError` to keep the two
semantic error domains unambiguous. No compatibility aliases were added.

Shared OLE Property Set parsing/editing is layered in `litchi-ole-common`,
including full CFB stream-path staging so nested streams with equal leaf names
are not lost during a metadata edit. Shared OOXML web-extension handling is
similarly split into semantic, XML, relationship, and package owners under
`litchi-ooxml-common`. These changes preserve typed snapshots, atomic edits,
lossless unknown content, and package-graph validation at their existing
facades; they do not claim complete Office-specification conformance.

The affected all-target `cargo check` matrix passes under the workspace lint
policy and with lint capping. The affected library-test matrix passes, as do
the DOC border and writer/encryption/leniency/VBA/glossary integrations, the
PPTX notes CRUD regression suite, the umbrella `litchi` suite (162 tests plus
three integration targets), formatting, diff, and the 46-package boundary
check. The known malformed Apache POI DOC fixture remains a separate
integration limitation recorded above.

## Layered OLE2, OOXML, ODF, and shared-codec continuation

This slice completes another breaking structural pass over the remaining
dense owners without retaining compatibility aliases or format-name prefixes
inside contextual modules. DOC now layers character properties, fields, and
writer package assembly; the public facade uses the concise `Leniency`,
`ToleranceReport`, `StylesheetDefect`, `EncryptionProfile`, `Element`, and
`Section` names. DOCX settings, DrawingML chart reading, ODP parsing, ODT
index writing, OGraph chart records, PPTX animations, XLS list objects and
writer core, XLSB worksheet writing, and XLSX catalog editing now expose the
same facade/model/codec/package/validation/test seams where applicable.

RTF's lexer, parser, and writer are now organized beneath semantic codec
facades. IWA media and Numbers editing received corresponding model/codec/
package seams. The format-neutral OfficeArt wire vocabulary and helpers moved
from PPT into `litchi-odraw`; PPT's Escher owner now consumes the shared typed
wire model. These moves preserve typed snapshots, deterministic output,
lossless unknown content, and existing ergonomic owner paths while reducing
format-local duplication.

The lint-capped affected all-target check passes for the complete migration
matrix, and strict checks pass for every affected crate. Library tests pass for
DOC (832 with two ignored), DOCX (643), DrawingML (92), IWA (1,529), ODraw (59), ODP (103),
ODT (512), OGraph (40), PPT (882 with one ignored), PPTX (305), RTF (287), XLS
(844), XLSB (413), and XLSX (645). The DOC facade integration targets pass,
the umbrella library passes 162 tests, formatting/diff checks pass, and the
crate-boundary policy remains clean. Root example linking was not used as a
gate because the environment's parallel `rust-lld` link crashed with SIGBUS;
that is an infrastructure limit, not a source diagnostic. This remains
migration evidence, not a claim of complete `[MS-DOC]`, `[MS-ODRAW]`,
`[MS-OGRAPH]`, `[MS-OSHARED]`, `[MS-PPT]`, `[MS-XLS]`, OOXML, or ODF
conformance.

## Prefix-free DOC facade and RouteSlip continuation

The next breaking DOC pass removes the remaining redundant `Doc` prefix from
the writer facade and adjacent owners. `Writer`, `WriteError`, `HeaderKind`,
`Picture`, `SmartTagEntry`, `StyleRevision`, `StyleDefinition` (under the
writer facade), `MtefEquationWriteOptions`, `TextBox`, `Revision`,
`RevisionKind`, `RevisionMetadata`, `RevisionEditor`, and the small writer
`IoError` seams now use contextual names without compatibility aliases. The
root facade keeps the reader-side `StyleDefinition` distinct from the writer
facade's same contextual name, avoiding an ambiguous root export.

The `[MS-DOC]` `RouteSlip`, `RouteSlipInfo`, and protection metadata now have
a typed, lossless `parts::route_slip` owner and contextual
`litchi_doc::route_slip` facade. Its byte-oriented ANSI strings avoid an
incorrect UTF-8 assumption; the codec validates Bool16 values, reserved
fields, enum domains, signed lengths, recipient counts, stage relationships,
truncation, overflow, and trailing bytes. It reads and writes the optional
FIB `fcRouteSlip`/`lcbRouteSlip` range through the table-stream seam.
`Document::route_slip()` exposes deferred optional metadata, while the package
editor owns immutable package snapshots, typed recipient selectors,
transactional lifecycle edits, reversible semantic/OLE patches, and a
round-trip reparse check. Route protection rejects lifecycle mutations unless
the policy is `Off`; the protection value is not conflated with DOP/range
protection, and authentication, mail transport, and host routing remain
inert. The focused route-slip suite covers selector errors, rollback,
stage/recipient lifecycle, clearing the FIB range, protected atomic rejection,
and `Document` round-trip visibility.

Strict DOC all-target and umbrella-library checks pass, as do lint-capped
checks, 837 DOC library tests (two ignored), 162 umbrella library tests, the
route-slip and renamed writer integration targets, formatting/diff checks, and
the crate-boundary policy. This remains bounded `[MS-DOC]` implementation and
API evidence, not a claim of complete Word routing workflow or Office
conformance.

## OLE2, OOXML, and ODF owner-wave continuation

The next owner wave extends the same layered topology across the remaining
large format-specific files. DOC now exposes inert, typed `[MS-DOC]`
`OcxInfo`/`RgxOcxInfo` records under `parts::ole_controls`, preserving
undefined handles and reserved bits while validating story domains, field
indices, cookie uniqueness, fixed record sizes, counts, and FIB table bounds.
It deliberately does not activate controls or promise an ActiveX lifecycle.

DOCX document-package, web-extension, and section-writing owners now separate
facades from semantic models, XML/package codecs, validation, relationships,
and focused tests. PPT writer-core model data and PPTX ChartEx semantic data
received the same treatment. XLS revision records, XLSB host cell reading,
XLSX pivot reading, ODS content traversal, and ODT mutable editing now use
nested owners with preserved typed APIs, snapshot/lossless behavior, and
package seams where applicable.

The affected all-target check passes both under the workspace lint policy and
with lint capping. Library tests pass for DOC (841 with two ignored), DOCX
(644), ODS (67), ODT (512), PPT (882 with one ignored), PPTX (305), XLS (844),
XLSB (413), and XLSX (645). Focused OcxInfo, web, section, pivot, revision,
and mutation tests also pass; formatting, diff, and the 46-package boundary
policy remain clean. ODS content tests outside the library target retain the
pre-existing orphaned-owner/path limitation reported by that owner. This is
bounded implementation and topology evidence, not a claim of complete
`[MS-DOC]`, `[MS-ODRAW]`, `[MS-OGRAPH]`, `[MS-OSHARED]`, `[MS-PPT]`,
`[MS-XLS]`, OOXML, or ODF conformance.

## Layered OLE2, OOXML, ODF, and shared-codec owner continuation

This owner wave continues the breaking, prefix-free topology migration through
the next dense semantic facades. DOC mail-merge, stylesheet, and TAP writer
owners; DOCX package ownership; ODraw shapes; OGraph chart models; PPT slide
types; PPTX animation XML and table styles; XLSB formula codecs; and XLSX raw
worksheets and workbook edits now separate contextual facades from typed
models, wire/XML codecs, validation, and focused tests. ODS data-pilot and ODT
graphic-property models receive the same nested organization. Existing public
owner paths remain ergonomic, with no compatibility aliases or repeated
format prefixes.

The shared `litchi-ole-common::toolbar` owner adds bounded, inert
`[MS-OSHARED]` `WString`, toolbar-header, control-header, flag, type, and
dimension codecs. It preserves borrowed UTF-16 payloads, reserved bits, and
deterministic serialization while deliberately not executing commands, macros,
icons, or UI behavior. DOC/PPT/XLS format-specific CTB/Customization lifecycle
integration remains a separate follow-up.

Strict and lint-capped all-target checks pass for the affected DOC, DOCX, ODS,
ODT, OLE common, ODraw, OGraph, PPT, PPTX, XLSB, and XLSX crates. The
lint-capped library matrix passes DOC (841 with two ignored), DOCX (644), ODraw
(59), ODS (67), ODT (512), OGraph (40), OLE common (32), PPT (882 with one
ignored), PPTX (305), XLSB (413), and XLSX (645). Formatting, diff, and the
46-package boundary check pass. This remains bounded migration and
specification evidence, not a claim of complete `[MS-DOC]`, `[MS-ODRAW]`,
`[MS-OGRAPH]`, `[MS-OSHARED]`, `[MS-PPT]`, `[MS-XLS]`, OOXML, or ODF
conformance.

## Layered OLE2, OOXML, ODF, RTF, and IWA continuation

This follow-up wave continues the same breaking, contextual migration through
behavior-heavy owners. DOC now has nested document state, writer-core state,
field tests, and a public bounded `CommandBars` owner. DOC command bars parse
and serialize the FIB `Tcg` seam with inert `PlfMcd`, `PlfAcd`, and `PlfKme`
records plus a lossless CTBWRAPPER shell; variable control-data and unknown
records are rejected rather than guessed, and no macro or UI behavior runs.
DOCX field tests, DrawingML chart-reader semantics, IWA Numbers semantics, ODS
traversal, ODT parser codec, PPTX ChartEx validation, RTF field content, XLSB
pivot writing, and XLSX workbook-edit tests now use nested contextual owners.

The public facades retain typed snapshots, borrowed data, deterministic
serialization, and format-local ergonomics. No compatibility aliases or
redundant format prefixes were introduced. One moved XLSX layout test required
an explicit `Value` import after the test tree split; no production behavior
changed.

The strict all-target matrix passes for every affected crate. The lint-capped
all-target matrix passes for DOC, DOCX, DrawingML, IWA, ODS, ODT, PPTX, RTF,
XLSB, and XLSX. The
lint-capped library matrix passes DOC (847 with two ignored), DOCX (644),
DrawingML (92), IWA (1,529), ODS (67), ODT (516), PPTX (305), RTF (287), XLSB
(413), and XLSX (645). Formatting, diff, and the 46-package boundary check
pass. This remains bounded implementation evidence, not a claim of complete
`[MS-DOC]`, `[MS-ODRAW]`, `[MS-OGRAPH]`, `[MS-OSHARED]`, `[MS-PPT]`, `[MS-XLS]`,
RTF, OOXML, or ODF conformance.

## Layered owner and XLS XCB continuation

This continuation completes another disjoint set of semantic owner moves while
adding the next bounded OLE2 feature. DOC numbering, DrawingML diagram data,
IWA editor tables, ODS style protection, ODT mutable semantics, PPT animation
parser tests, PPTX shape tags, XLS workbook codecs, XLSB conditional-formatting
binary codecs, and XLSX data-validation codecs now use contextual nested
facades over model, wire/codec, validation, and focused-test seams. The public
names remain prefix-free within their format contexts; no compatibility aliases
were added.

The XLS owner now exposes an inert `toolbar` facade for the `[MS-XLS]` XCB
stream. It reuses `litchi-ole-common::toolbar` for shared TB/TBC headers and
flags, preserves reserved fields and fixed visual bytes, and round-trips the
bounded CTBWRAPPER/CTBS/CTB structure. Controls requiring variable `TBCData`,
unknown control types, or uninterpreted command payloads are rejected rather
than guessed; no macro, UI, ActiveX, or external behavior executes.

Strict all-target checks pass for DOC, DrawingML, IWA, ODS, ODT, PPT, PPTX,
XLS, XLSB, and XLSX. With lint caps for the known workspace lint debt, the
library matrix passes DOC (847 with two ignored), DrawingML (92), IWA (1,529),
ODS (67), ODT (516), PPT (882 with one ignored), PPTX (305), XLS (855), XLSB
(413), and XLSX (645). Focused additions include 45 DOC numbering tests, 13
DrawingML diagram-data tests plus two API tests, 39 IWA table tests, 8 ODS
style-protection tests, 15 ODT mutable tests, 35 PPT parser tests, 7 XLS XCB
tests, 44 XLSB binary tests, and the XLSX data-validation codec/mutation tests.
Formatting, diff, and the 46-package boundary check pass. This is bounded
feature and topology evidence, not a claim of complete OLE2, OOXML, or ODF
conformance.

## Layered OLE2 and OOXML owner continuation with XCB package integration

The next migration wave keeps the semantic contextualization rule moving
through shared and format-local owners. `litchi-ole-common` now separates the
property-set binary codec into wire, VARIANT-semantic, validation, and test
owners. DOC image writing, PPT document-comparison and embedded-object owners,
ODraw image data, OGraph package assembly, XLS query tables, DOCX field tests,
PPTX tag packages and animation XML parsing, XLSB workbook-writer tests, XLSX
raw worksheet editing and workbook snapshots, ODS sheet traversal, and ODT
field codecs now use nested facades instead of dense single-file owners.

The bounded XLS toolbar codec is now connected to the format facade: the
`Workbook::toolbar` reader opens the optional root `XCB` stream into owned
metadata, while `Writer::set_toolbar` and `clear_toolbar` deterministically
create or remove that stream. The shared toolbar model owns the lifetime
conversion; all command, macro, UI, and ActiveX behavior remains inert.

Strict all-target checks pass for DOC, DOCX, DrawingML, ODraw, ODS, ODT,
OGraph, OLE common, PPT, PPTX, XLS, XLSB, and XLSX. With lint caps, the
library matrix passes DOC (847 with two ignored), DOCX (644), DrawingML (92),
ODraw (59), ODS (67), ODT (520), OGraph (40), OLE common (33), PPT (882 with
one ignored), PPTX (309), XLS (856), XLSB (413), and XLSX (648). The package
XCB integration tests pass 2/2. Formatting, diff, and the 46-package boundary
check pass. This remains bounded owner and feature evidence, not a claim of
complete OLE2, OOXML, or ODF conformance.

## Typed OLE2 metadata and deeper owner continuation

This continuation adds bounded typed behavior alongside the ongoing topology
migration. DOC `embedded_object::Info` now preserves the optional
`ODTPersist2` presence bit, defined and undefined persistence fields, and
deterministically serializes passive `ObjInfo` metadata under the `[MS-DOC]`
ODT grammar. Required-zero bits are validated and malformed lengths fail; no
OLE payload is opened, instantiated, or executed.

PPT now exposes inert `DiagramBuildContainer`/`DiagramBuildAtom` metadata from
`[MS-PPT]` §§2.8.13–2.8.14 and 2.13.7. Fixed-width unknown enum values and
reserved bytes round-trip, while diagram rendering, SmartArt authoring, and
animation playback remain outside the API. The same wave layers DOC document
semantics, sections, form fields, PPT writer records and animation editing,
ODraw properties, OGraph chart aggregates, XLS List12 and PivotTable writing,
XLSB pivot writing, XLSX package metadata, DOCX package tests, and PPTX shape
anchor mutation under contextual facades.

Strict all-target checks pass for DOC, DOCX, ODraw, OGraph, PPT, PPTX, XLS,
XLSB, and XLSX. With lint caps, the library matrix passes DOC (849 with two
ignored), DOCX (644), ODraw (59), OGraph (43), PPT (888 with one ignored),
PPTX (309), XLS (858), XLSB (413), and XLSX (648). Focused additions include
the DOC ODT persistence tests, six PPT diagram-build tests, 17 DOC section
tests, 25 PPT writer-record tests, 10 XLS PivotTable-writer tests, 15 ODraw
property tests, three OGraph aggregate tests, nine XLSB pivot-writer tests,
four XLSX metadata tests, and the preserved DOCX package-test families.
Formatting, diff, and the 46-package boundary check pass. This remains
bounded implementation evidence, not a claim of complete `[MS-DOC]`,
`[MS-ODRAW]`, `[MS-OGRAPH]`, `[MS-PPT]`, `[MS-XLS]`, OOXML, or ODF
conformance.

## Layered OLE2 and OOXML continuation with typed control metadata

This wave applies the same contextual hierarchy to another disjoint set of
dense owners. DOC embedded-object transactions, field parsing, writer package
semantics, and writer tests; ODS data-pilot parsing; PPT embedded objects,
animation behavior parsing, text-format and text-style writing; XLS OLE-object
and writer-stream owners; XLSB formula text and worksheet writing; and XLSX
ActiveX and XLDM package owners now use nested facades over semantic models,
wire/XML or BIFF codecs, validation, package/transaction seams, and focused
tests. The format facades remain prefix-free and no compatibility aliases were
added; one PPT editor finalization path also avoids an unnecessary output
buffer clone.

DOC `parts::ole::controls` now decodes the specified 20-byte `[MS-DOC]`
`OcxInfo` body (`dwCookie`, field index, accelerator metadata, flags, and
document selector) when present, retains shorter producer entries and
undefined tails, and provides deterministic `RgxOcxInfo` serialization. The
owner remains passive: it does not expose a control runtime, ObjectPool
activation, event dispatch, rendering, or macro execution.

Strict all-target checks pass for DOC, ODS, PPT, XLS, XLSB, and XLSX. With lint
caps, the library matrix passes DOC (851 with two ignored), ODS (67), PPT (888
with one ignored), XLS (862), XLSB (413), and XLSX (653). Formatting, diff,
and the 46-package boundary check pass. This is layered-owner and bounded
metadata evidence, not a claim of complete OLE2, OOXML, or ODF conformance.

## Shared TBCData and continued OLE2 owner migration

The next OLE2 wave moves the flag-controlled `[MS-OSHARED]` `TBCGeneralInfo`
and `TBCExtraInfo` structures into `litchi-ole-common::toolbar`. Borrowed
`WString` values, typed OLE host/server and menu-merge modes, disabled/UI
flags, and exact format-specific tails are retained without activating
commands, macros, ActiveX, or UI behavior. DOC command bars and XLS XCB
owners consume this common model at their format-specific record boundaries;
ambiguous or malformed variable payloads remain rejected rather than guessed.

The same continuation deepens the legacy owners with typed DOC embedded-object
inventory metadata, OfficeArt shape/group projections, FIB and TAP seams, PPT
chart transactions, Escher records, animation timing properties, and
failure-atomic embedded-storage snapshots. XLS form-control/OLE-object,
list-object, pivot/OLAP, and toolbar owners gain bounded typed metadata and
authoring operations. `litchi-ooxml-common` web extensions and
`litchi-drawingml` chart writing are split into contextual semantic, wire/XML,
package, validation, and test layers. Unsupported records remain lossless or
explicitly inert, and all facades retain prefix-free names.

These changes are specification-backed migration evidence for `[MS-DOC]`,
`[MS-ODRAW]`, `[MS-OGRAPH]`, `[MS-OSHARED]`, `[MS-PPT]`, `[MS-XLS]`, and
DrawingML/OOXML. They do not claim complete format conformance or runtime
control execution.

## Layered ODF/OOXML owner continuation and common object snapshots

The next migration wave keeps the same contextual hierarchy while moving
another set of dense owners behind thin facades. ODT document and text-element
owners, ODS database ranges and table-template styles, and ODP parser XML and
authoring builders now separate semantic state, XML/package codecs, validation,
and focused tests. DOC OLE metadata, PPT animation timing, and XLS differential
formats follow the same shape for the legacy binary owners.

The OOXML owners receive the corresponding glossary, animation, modern-comment,
XLSB formula/resolution, XLSX worksheet-snapshot, and XLSX workbook-transaction
seams. Their public types remain contextual and prefix-free; moving a file does
not create a second compatibility facade or duplicate a format-specific wire
grammar.

`litchi-ole-common::object` now also exposes an immutable `Snapshot`. Snapshot
clones share captured CFB stream allocations and create independent editors,
so DOC/PPT/XLS hosts can retain read state without copying large embedded
payloads. Format-owned metadata remains opaque in the common layer, while
semantic interpretation stays in the owning format crate.

The affected all-target check and lint-capped library matrix pass after the
migration. Focused evidence includes the common OLE object/property/toolbar
suites, ODP authoring/parser tests, ODT and OOXML owner suites, and DOC/PPT/XLS
integration tests. This is bounded topology and regression evidence, not a
claim of complete ODF, OOXML, or `[MS-DOC]`/`[MS-PPT]`/`[MS-XLS]` conformance.

## Typed OLE arrays and the next owner migration wave

`litchi-ole-common::property_set` now models `[MS-OLEPS]` `VT_ARRAY` values as
typed, bounded multidimensional arrays with row-major dimensions and checked
element cardinality. The same owner replaces untyped vectors with a scalar
typed `Vector`, while retaining the per-element headers required by
`VT_ARRAY|VT_VARIANT` and `VT_VECTOR|VT_VARIANT`. Unsupported array/vector
base types, malformed bounds, reserved fields, and unsafe nesting are rejected
or preserved as inert unknown values; no format crate owns a duplicate
Property Set grammar.

The shared `[MS-OSHARED]` toolbar model is split into contextual controls,
flags, headers, merge, restriction, and text/icon owners. DOC field navigation
and OfficeArt group snapshots, PPT chart/ODraw owners, XLS chart and inert OLE
control metadata, and XLSB pivot-cache definition/record validation now sit
behind semantic, wire/BIFF, validation, package, and test layers. ODT mutable
editing, ODS authoring/formula evaluation, and ODP authoring receive the same
ODF treatment. Standalone DOCX paragraph codecs, PPTX presentation properties,
and XLSX chart-sheet owners continue the OOXML migration without restoring the
former `litchi-ooxml` package.

The affected library matrix passes with DOC 869 (two ignored), DOCX 648, ODP
104, ODS 75, ODT 525, PPT 907 (one ignored), PPTX 312, XLS 863, XLSB 413,
XLSX 655, and 40 shared OLE tests; all affected targets also pass `cargo
check`. This is specification-backed ownership and round-trip evidence for
the touched `[MS-OLEPS]`, `[MS-OSHARED]`, `[MS-DOC]`, `[MS-ODRAW]`,
`[MS-OGRAPH]`, `[MS-PPT]`, `[MS-XLS]`, `[MS-XLSB]`, ODF, and DrawingML paths,
not a claim of complete format conformance.

The subsequent continuation adds focused, bounded owners for DOC route slips,
DOCX run effects, ODP handout masters, ODS metadata/settings, ODraw geometry,
OGraph chart patches, PPT masters and hosted charts, PPTX model3d resources,
and XLSB/XLSX slicer and timeline parts. The affected-crate check passes with
zero boundary debt; focused suites cover rollback, exact round trips,
relationship/dependency guards, malformed input, and source-checked OLE patch
application. This remains incremental specification coverage, not a claim of
complete Office or ODF conformance.

This continuation also removes the public outer `litchi_doc::document` module,
leaving the typed `Document` facade at the crate root, and adds the bounded
`litchi-ole-common::property_set::VersionedStream` model and binary codec for
`[MS-OLEPS]` 2.13. Its indirect stream selector is validated against the
owning property identifier; the referenced non-simple CFB stream remains
inert package data. Focused tests cover UTF-16/code-page round trips,
identifier mismatch rejection, and opaque unsupported stream variants.

This continuation also covers the shared `[MS-OSHARED]` `HeadingPairs` and
`DocParts` vector composites with code-page-aware binary codecs and cross-PID
validation; typed `[MS-XLS]` XML-map metadata and root-stream loading with
list-column dependency checks; bounded `[MS-DOC]` auto-summary authoring;
typed `[MS-ODRAW]` solver rules; and `[MS-PPT]` master `SlideNameAtom`
authoring. All external binding, schema resolution, macro execution, layout,
and rendering remain inert.

The DOCX vertical slice additionally types and validates `[MS-DOCX]` 2.6.2.1
and `[MS-ODRAWXML]` 2.18.2.1 `anchorId` values in the drawing inventory. The
parser accepts only eight-digit hexadecimal identifiers in the specified
nonzero, below-`0x80000000` range and preserves them without activating or
rendering the drawing.

## Evidence levels

For each applicable object/scenario, track:

1. read and extraction;
2. create, update, clear/remove, and structural dependencies;
3. lossless preservation of unsupported content;
4. schema/specification validation;
5. native Microsoft Office open without repair;
6. native Office edit and resave;
7. reverse-read with expected semantic and package diffs;
8. single-thread and multi-thread performance evidence where relevant.

Office verification uses generated and curated fixtures on current Windows and
macOS desktop Office plus pinned compatibility baselines. Web Office is
supplemental. Resave comparison permits only a reviewed metadata allowlist;
repair logs, inserted repair parts, missing content, or unexplained semantic
diffs fail.

## Baseline regression disposition

Computer Use verification exposed three destructive or repair-inducing
behaviors. Phase one applies safe containment while the lossless transaction
model is built:

- Opened XLSX mutation formerly entered a fresh writer workbook and could
  discard existing sheets and rows. `worksheet_mut` now returns a typed
  `UnsafeEdit`, and `save` rejects other legacy rebuild paths before creating
  the destination.
- Opened PPTX mutation formerly entered an empty presentation writer and could
  discard existing slides. `presentation_mut` now returns a typed `UnsafeEdit`.
- Newly created PPTX files formerly shared the slide-master theme with the
  notes master, causing desktop PowerPoint to repair the package. Notes now own
  `theme2.xml`, handouts reserve `theme3.xml`, and package-graph regression tests
  pin the relationships.

These guards prevent silent loss; they are not the final lossless edit API.
The umbrella detection/open path and concrete readers also still disagree on
some enabled spreadsheet formats and remain migration work.

Post-fix desktop PowerPoint verification on macOS opened a generated one-slide
artifact without a repair dialog. PowerPoint then added a second slide and
saved the file; Litchi reverse-read both slides, and the resaved package retained
only the intended `theme1.xml` and distinct notes `theme2.xml` theme parts.

Desktop Excel verification on macOS opened the deterministic workbook emitted
by the second slice without a repair or compatibility dialog. Excel accepted an
edit to cell A1 and saved the workbook; the Office-resaved archive passed ZIP
integrity checks. After the third slice, the public sparse cell facade
reverse-read A1 from that same Office-resaved artifact as the exact text
`"Office round trip"`. This certifies minimal-package creation, one native Excel
edit/resave, catalog recovery, shared-string resolution, and semantic readback
for that cell. It does not certify Litchi-authored cell create/update/clear/remove
or general Excel compatibility.

For the fourth slice, desktop Excel on macOS opened a Litchi-authored workbook
containing text at A1, the number 42 at B2, the formula `=B2*2` at C3, and an
explicit empty D4 without a repair or compatibility dialog. Excel reported the
used range as A1:D4. The current Excel session was configured for manual
calculation, so verification explicitly pressed F9; C3 then displayed 84 while
the formula bar retained `=B2*2`. Excel resave normalized the unstyled empty D4
away, materialized the formula cache as 84, and added producer-owned theme,
style, shared-string, and calculation-chain parts. This is reviewed producer
normalization, not evidence that an explicit empty record survives an Office
round trip.

Litchi reverse-read that Office-resaved artifact with the expected A1, B2, and
C3 semantics, edited it again, removed the stale calculation chain, and
preserved the unrelated Office-added parts. Excel then opened this
second-generation artifact without repair, again reported A1:D4, retained the
exact C3 formula, and produced 84 after F9. Together these checks certify this
ordinary-cell set/clear/remove slice on the tested macOS Excel build, including
one Office-resave/re-edit cycle. They do not certify other spreadsheet CRUD
families, automatic calculation under a manual application setting, Windows
Office, older Office versions, or performance.

For the Numbers formula ownership slice, a fresh Litchi-authored archive opened
in native Numbers with the generated formula result `323`. Numbers then changed
the input cell from `120` to `43` through its real cell editor, recalculated the
formula to `246`, saved, closed, and reopened the archive. The reopened document
retained both `43` and `246`; the archive passed ZIP integrity checks, and the
Numbers application was closed after verification.

For the shared font ownership slice, the existing native table-layout generator
created a fresh `/tmp/litchi-iwa-font-layouts.5zsTwA/table-layouts.numbers`
artifact using `CourierNewPSMT` through the extracted public font value. Native
Numbers opened it without a repair prompt and rendered the multi-line styled
cell. A real UI edit changed that cell to `Native font round trip`; the file was
saved, closed, reopened, and the edited text was retained. The generator's
reader-side verification also recovered the authored font and layout values;
Numbers was quit after the native check.

The worksheet parser also matches checked-in Apache POI and LibreOffice shared-
formula fixtures, including translated follower expressions and stored cached
results. Synthetic tests cover missing versus explicit empty cells, grid-bound
rejection, malformed shared-formula groups, exact numeric lexemes, sparse range
order, read-only serialization stability, and concurrent first access. These
are read-path regression gates, not performance evidence; allocation, latency,
contention, and scaling claims still require the measurement work in ADR 0005.

## PPT bookmark-summary owner layering

The flat legacy PPT `bookmark_summary` owner is now physically layered into
`model`, `codec`, `validation`, and focused `tests` modules. The public
`BookmarkSummary` type is replaced by the contextual `bookmark_summary::Summary`
name, with no prefix-expanded alias. Record framing and UTF-16 conversion stay
in the codec; entry-count, ID-seed, text-bookmark identity, and bounded string
invariants stay in validation; the owner remains inert and does not resolve or
activate bookmark targets.

The focused verification command was attempted with a single Cargo build job,
but the current workspace stops earlier in the unrelated `litchi-odraw` crate:
its `prop::gradient::Stops` `const fn`s call non-const `Array` accessors. This
slice therefore has formatting and diff evidence, but no new passing Cargo
test claim until that pre-existing dependency error is repaired in its own
bounded owner slice.

## Quality gates

The IWA migration has now established a value-model seam. `litchi-iwa-text`
owns `storage::{Storage, Run, Fragment}`; `litchi-pages`
owns `Section` and `SectionType`; and `litchi-keynote` owns `Slide`, `Show`,
build-animation, and transition values. Its `transition::Effect` now owns the
lossless native transition-effect identifiers and canonical-known-value check;
the monolith retains only archive-boundary transition settings and wire
mutation. The ordinary Keynote reader maps modern and legacy archive effect
identifiers into the same leaf-owned `SlideTransition::effect` value, so the
reader no longer collapses known or future effects to an untyped `Other` case.
The old implementations were removed from the extracted value owners,
while the monolith still has migration
adapters where existing reader/editor surfaces need archive-boundary context.
These adapters are staged ownership work, not a compatibility API. No archive,
protobuf, or application decoder was moved into these value crates.
The rich-text storage handoff is now complete for the bounded semantic seam:
`litchi-iwa-text::storage::{Storage, Run, Fragment}` contains no native object
or style identifiers and validates every published run against UTF-8 text.
Keynote, Pages, and `litchi-iwa-structured` consume the leaf directly, while
the IWA adapter performs decoded text-line joining and retains native lookup,
UTF-16 boundary conversion, and unsupported wire content. The focused leaf,
Keynote, Pages, and structured tests pass; the full IWA test target remains
blocked by the pre-existing Numbers conditional-highlighting test that still
passes the removed `TableCellCheckboxFormat` type to the new `Checkbox` API.
The adapters and the four corresponding `litchi-iwa` dependency edges are
staged ownership work, not compatibility API; their exit is to move the owning
readers before deleting the adapters. The Numbers value slice now owns
`litchi-numbers::cell::{Value, Type, Update}` and leaves only a monolith-local
private adapter for the remaining reader/editor migration. The BNC wire slice
now also owns `litchi-numbers::cell::wire::{BncCell, StoredValue,
CachedScalar, CellDataFormatKind}` and the dependency-free decimal128 codec;
the monolith retains only a private module alias plus its archive/protobuf
callers. The combined Numbers cell/semantic leaf suite has 26 tests, the new
formula vocabulary suite has 4 tests, and the IWA suite has 1,504 tests. The
boundary check plus native Numbers smoke cover the extraction. The dependency-
free `litchi-numbers::formula` module now owns formula caches, references,
operators, and expression construction; `litchi-iwa` retains only the
archive-boundary compiler, protobuf AST, and calculation-engine mutation.
The former IWA formula module is crate-private, and its root-level re-exports
are documented ergonomic aliases rather than compatibility shims. Formula
compilation performs an iterative preflight for bounded depth, AST nodes,
function arguments, and aggregate precedents before the recursive wire walk;
known fixed-arity functions and unary constructors are covered by focused
tests, while functions without validated arity metadata fail closed.
The rich-text font ownership slice is now complete: dependency-free
`litchi-iwa-text::font::{Font, Name}` owns the validated, one-allocation font
identity and typed `NameError`; IWA retains only a contextual alias and the
native archive adapter. The leaf's eight tests cover the existing storage
models plus bounded/strict font construction, owned-input consumption, default
semantics, and named identity. The scoped IWA example check passes with a
Numbers fixture that writes `CourierNewPSMT` through the public table-cell
font operation; native Numbers verification is recorded below.
The neutral color ownership slice is now complete: dependency-free
`litchi-iwa-common::color::{RgbColorSpace, Rgba}` owns the fixed-size validated
RGBA value and typed `color::Error`; IWA retains only protobuf conversion and a
facade error adapter. The common color tests cover compact representation,
opaque/transparent defaults, valid Display P3 values, and strict channel
validation. The existing Pages shape authoring example compiles through the
new common value and is the native artifact used for this slice's Computer Use
check.
For the table-appearance ownership slice, the fresh generated
`/tmp/litchi-iwa-appearance.G82gPP/table-appearance.numbers` opened in native
Numbers without a repair prompt. Selecting its `Appearance` table exposed
the authored alternating-row and row-fit controls in the table formatter.
The matching `table-appearance.key` opened in Keynote without repair and
displayed the six-by-three table; selecting it exposed the same alternating-row
and resize-to-fit controls. The Pages artifact reported that it was damaged;
the warning was dismissed without repair or save, matching the existing Pages
limitation for this generated table family. ZIP integrity passed for all three
artifacts, and Numbers, Pages, and Keynote were quit after the check.
The table-appearance ownership slice is now complete: dependency-free
`litchi-iwa-common::table::appearance::{Appearance, Banding, RowSizing,
GridlineVisibility, Gridlines}` owns the compact semantic value, including its
native-default representation. IWA retains only the archive adapter: strict
wire override decoding, bounded style inheritance, native bool conversion,
and transactional copy-on-write style mutation. Existing Numbers, Pages, and
Keynote CRUD surfaces continue to use contextual facade aliases while the
duplicate value implementation is removed from the monolith.
The table-cell layout ownership slice is now complete at
`litchi-iwa-common::table::cell::layout::{TextWrap, VerticalAlignment, Inset,
Insets, Layout}`. Its two-value enums and 4/16/20-byte layout components are
archive-free and heap-free; `Inset` rejects negative, NaN, and infinite input
with a typed allocation-free error. IWA retains only native alignment and
padding conversion, style inheritance, and package transactions. The Numbers,
Pages, Keynote, and layout-generator consumers now import the common owner
directly; the old public `litchi-iwa::table_cell_layout` module and its
migration aliases are removed rather than retained for compatibility.

Native Computer Use verification of a fresh
`/tmp/litchi-iwa-layout.ONYCvR` fixture opened `table-layouts.numbers` and
`table-layouts.key` without repair prompts. Selecting B2 in each application's
real text inspector reported `Wrap text in cell = 1` and `Vertical alignment =
middle`; the rendered multi-line value retained the authored 8-point inset
layout. `table-layouts.pages` reproduced the known generated-table damaged-file
limitation; the warning was dismissed without repair or save. All three ZIP
archives passed integrity checks, and Numbers, Pages, and Keynote were quit
after verification.
The shape text-layout ownership slice now follows that table-cell precedent at
`litchi-iwa-common::text::layout::{VerticalAlignment, AutoSize, Inset, Insets,
Layout}`. The common suite pins 4/16/20-byte value sizes and typed rejection of
negative, NaN, and infinite insets; the IWA suite keeps native enum/padding
conversion, bounded style inheritance, and transactional archive mutation. The
six shape and text-box creation examples import the common owner directly, and
the focused native gate opened fresh Numbers, Pages, and Keynote shape
artifacts without repair prompts. Numbers exposed `Bottom`, `Fixed`, and a
9-point inset; Pages exposed `Middle`, `Fixed`, and a 12-point inset; Keynote
exposed `Middle`, `ShrinkToFit`, a 14-point inset, and a checked `Autosize Text`
control. Each app accepted a real text edit and Save, and reverse-read through
the public layout APIs preserved those values and the edited text. All three
ZIP archives passed integrity checks, and Numbers, Pages, and Keynote were
quit after verification. The old `ShapeText*` model is deleted rather than
retained as a compatibility alias.
The media-kind ownership slice now follows the same boundary at
`litchi-iwa-common::media::Type`. Its common tests pin one-byte storage,
case-insensitive extension classification, representative signature families,
the explicit `Unknown` case, and conservative unknown-`ftyp` handling. IWA
retains media discovery, filesystem/package access, resource limits, metadata
rewrites, and transactional replacement; its image, audio, movie, chart, and
shape consumers import the common owner directly. Native media fixtures were
verified through the existing real iWork image/media paths: fresh Numbers,
Pages, and Keynote packages built from `test-data/images/png/lena.png` opened
in the installed applications without repair prompts, and each exposed the
embedded image object in its native canvas/inspector. The CLI reported
`Type::Image` and the reachable data identifier for all three packages; all
three ZIPs passed integrity checks, and the applications were quit after
verification. The semantic classifier itself is exercised by the
dependency-free suite, including complete and truncated ISO-BMFF headers.
The media playback ownership slice now follows the same archive boundary at
`litchi-iwa-common::media::playback`. Common tests cover compact volume
validation, builder-preserved optional fields, strict trim-range validation,
canonical known loop modes, and lossless genuinely unknown native values. The
IWA adapter retains protobuf decoding, legacy/modern loop reconciliation,
unknown-field-preserving wire patches, and transactional replacement; its
wire-focused tests continue to cover those behaviors while Pages, Numbers,
Keynote, and all six source-building media examples import the common types
directly. The old IWA semantic owners and root compatibility exports are
removed rather than duplicated.
The shape-path ownership slice now follows the same boundary at
`litchi-iwa-common::shape::path`. `Preset` owns the source-buildable geometry
vocabulary while `CornerRadius`, `PolygonSides`, `StarPoints`, and
`InnerRadiusRatio` are compact checked controls; their common tests pin scalar
size/alignment, `Copy`, finite/domain validation, and preservation of native
values. IWA retains structural path-family detection, native archive decoding,
natural-size corner-radius validation, protobuf conversion, and
wire-preserving mutation. Pages, Numbers, Keynote, and their shape examples
import the common owner directly, and the former redundant `Shape*` value
types are deleted rather than kept as compatibility aliases. Path-preset
mutation preserves the path-source envelope's known metadata, unknown fields,
and family-field position while replacing the owned family payload. Fresh generated
shape packages for all three applications opened without repair prompts during
the focused native gate; the matching semantic inspectors recovered the
expected preset and path kind, and Numbers, Pages, and Keynote were quit after
verification.
The chart-axis vocabulary slice follows the same ownership boundary at
`litchi-iwa-common::chart::axis::{Axis, TickMarkLocation}`. The common crate
owns the one-byte category/value selector and copyable tick-mark value,
including explicit preservation of unknown native integers; IWA retains the
axis archive slots, protobuf field mapping, shared-object checks, and
wire-preserving mutation. All Pages, Numbers, Keynote, and chart-example
consumers now use the short canonical names, with no `ChartAxis*` compatibility
aliases. Focused common and IWA library checks passed. The focused
`axis-label-angles-crate` fixtures opened in Numbers, Pages, and Keynote
without repair prompts: Numbers exposed Value (Y) `Right Diagonal`, Category
(X) `Left Diagonal`, and `Centered` tick marks; Pages and Keynote exposed the
same 2D chart with numeric Y and categorical X axes. The broader all-feature
Numbers chart generator still produces a ZIP-valid artifact that native
Numbers rejects as damaged, so it is treated as a fixture limitation rather
than chart-axis evidence. No native-resave claim is made; all three
applications were quit after verification.
The follow-on chart-axis value slice now places the archive-free vocabulary in
focused `litchi-iwa-common::chart::axis` children:
`bounds::{Bound, Bounds}`, `label_angle::LabelAngle`,
`label_position_3d::LabelPosition3d`, `scale::Scale`, and
`steps::{MajorStepCount, MinorStepCount, Steps}`. The facade retains archive,
protobuf, capability, and lossless wire adapters and removes the former long
`Chart*` value names. The common suite passed 43 tests, the IWA library suite
passed 1,502 tests, selected chart examples compiled, both scoped clippy gates
were clean, and the crate-boundary checker remained valid. The focused
`axis-values-crate` fixtures passed typed Rust readback and ZIP validation, and
fresh Numbers, Pages, and Keynote opens produced no repair prompts. Numbers
and Keynote exposed Value (Y) `Logarithmic`, `Max 30`, `Min 1`, and `Right
Diagonal` in the native Axis formatter; Pages exposed the same 2D chart with
numeric Y and categorical X axes. The native gate makes no resave claim, and
all three applications were quit after inspection.
The chart number-format ownership slice is complete as well:
`litchi-iwa-common::chart::number_format` now owns the compact
`FixedDecimalPlaces`, `DecimalPlaces`, `NegativeStyle`, `NumberFormat`, and
`LabelAffixes` values. The common suite passed 45 tests and the IWA library
suite passed all 1,502 tests; focused number-format and affix examples
compiled, their generated Numbers, Pages, and Keynote packages passed ZIP
validation, and the changed Rust files passed formatting and strict Clippy.
Native Computer Use opened both fixture families in all three iWork apps
without repair prompts. Numbers and Keynote exposed Number, parentheses (or
minus sign), thousands separator, two decimals, and the authored prefix and
suffix fields. Pages visibly rendered `14,000.00`, `(2,800.00)`, and `USD …
net` axis labels. No native-resave claim is made; Numbers, Pages, and Keynote
were quit after inspection.
The chart-series direction ownership slice is complete: the archive-free
`litchi-iwa-common::chart::Direction` value now owns row/column semantics and
lossless unknown-native preservation. IWA retains only protobuf field mapping,
archive lookup, and mutation validation; all three chart owners use the short
canonical value and the former `ChartSeriesDirection` implementation is gone.
The common suite passed 46 tests, the IWA library suite passed all 1,503
tests, and the focused axis-value fixture compiled, wrote three valid ZIP
packages, and set `Direction::Columns` in all three in-memory editors before
native inspection. Existing CRUD tests round-tripped the serialized direction
through each editor. Fresh Numbers, Pages, and Keynote opens produced no repair
prompts. Numbers' Add Chart Data popover reported `Plot Columns as Series`;
Keynote's Edit Chart Data dialog showed `Plot columns as series` selected with
the `Revenue`/`Cost` rows and `Q1`/`Q2`/`Q3` reference columns; Pages exposed the
same three chart series and visibly rendered `Revenue`/`Cost` categories. All
three applications were quit after inspection.
The BorderSide ownership slice
is complete: the dependency-neutral table-cell edge selector now lives at
`litchi-iwa-common::table::cell::BorderSide`; `Borders` and `ShapeStroke` remain
concrete IWA types, and the former Numbers-owned enum and compatibility path
were removed. This ownership change has no intentional wire-format change.
The physical IWA substrate slice is complete as well: raw schemas remain in
`litchi-iwa-protos`, bounded archive and Snappy framing remain in
`litchi-iwa-core`, the facade owns no duplicate Snappy codec, and its callers
use the allocation-conscious slice API. The core framing suite has 17 passing
tests. The facade varint exit is now complete too: `varint.rs` was deleted,
all callers use the common bounded implementation, and the IWA suite still
passes 1,504 tests. The facade `WireField` representation and direct wire
mutation exit is complete without a compatibility shim; the generic callback
error boundary is now common-owned, while `wire.rs` retains only thin
crate-error adapters pending the final import migration. The Numbers
table/sheet semantic slice is now extracted into dependency-free
`litchi-numbers::table` and `litchi-numbers::sheet`: finished tables use compact
coordinates and immutable boxed sparse storage, while builders provide
fallible append/replace operations and checked ownership handoff. The IWA
reader now exposes `NumbersDocument::semantic_sheets`, which moves those
finished tables into one lazily cached immutable `Arc<[Sheet]>` snapshot
without rebuilding cell maps; the opaque archive adapter remains available for
comments and native sidecars during the staged reader migration. Dense views
remain explicitly budgeted and reject ranges outside the declared extent. The
leaf suite has 26 cell/semantic tests plus 4 formula-vocabulary tests, the IWA
suite has 1,504 tests, and the generated Numbers round trip passes. The generic
structured-facade handoff
remains the next ownership slice. Pages and Keynote table readers now borrow
canonical sparse leaf tables
directly while retaining their format-owned comment and merge sidecars;
read-only comment snapshots use sorted boxed pairs, and Numbers ingress moves
table names into the adapter without a redundant clone.

The immutable Numbers table-data-list sidecars are now represented as sorted
boxed `(u32, value)` pairs rather than hash tables. The loader already rejects
duplicate keys, so binary search preserves strict missing-key behavior while
removing hash-bucket and hasher state from each read-only table. Sidecar
construction uses fallible reservations and checks the invariant again at the
compact representation boundary; this is an allocation/layout improvement,
not a measured throughput claim under ADR 0005.

The IWA Numbers ingress adapter now protects the leaf ownership seam. It
validates bounded dimensions before loading referenced data, rejects
duplicate or out-of-range tile keys and coordinates, requires one typed tile
payload, and maps allocation failures and finite table/cell budgets through
the shared structured error type. Offset decoding is sparse and single-pass
after a no-allocation shape/count scan; it never reserves from the archive's
untrusted `cell_count` alone. Focused malformed-input tests cover dimension,
coordinate, duplicate, odd-buffer, sparse-sentinel, descending-offset, and
allocation-amplification cases. The extractor distinguishes native type `6000`
TableInfoArchive metadata from type `6001` TableModelArchive cell payloads, so
metadata records cannot be mis-decoded as table models.

A fresh ZIP-valid generated artifact opened and rendered its 3×3 bordered table
in Numbers and Keynote, including the expected `Numbers` and `Keynote` cell
values. Pages reported the generated package as damaged; the warning was
dismissed without repair, so Pages native opening remains a tracked limitation
of this fixture rather than a claimed success.

The Keynote transition-scalar ownership slice is complete. The archive-free
`litchi-keynote::transition` module now owns compact `Direction`, `MosaicType`,
`Acceleration`, and `TextDelivery` values; `litchi-iwa` retains only aggregate
transition settings, archive decoding, protobuf field mapping, and
wire-preserving transactions. The leaf suite passed 6 tests and the IWA
library suite passed all 1,503 tests. The focused
`create_keynote_transition` example wrote
`/tmp/litchi-keynote-transition.n7Msjg/transition-vocabulary.key`, which passed
ZIP integrity validation and typed in-memory reopen checks. Native Keynote
opened it without a repair or recovery prompt; its Animate inspector exposed
`Magic Move`, `By Word`, and `Ease In & Out`, matching the authored effect,
text-delivery, and acceleration values. No native-resave claim is made, and
Keynote was quit after inspection.

The pie-visibility ownership slice is complete. The archive-free
`litchi-iwa-common::chart::pie` module now owns compact `LabelVisibility` and
lossless `LeaderLineVisibility`; all three concrete chart owners consume the
short values, and the old `ChartPie*Visibility` names are gone. The common
suite passed 48 tests, the focused IWA pie suite passed 16 tests, and the
Numbers CRUD test passed after adding stylesheet/component-registration
invariants for styled label overrides. The focused fixture saved and typed-
reopened valid Numbers, Pages, and Keynote ZIP packages. Native Pages and
Keynote opened the authored chart without repair prompts and exposed North
22%, South 33%, and West 44%; the screenshot showed the authored South-only
leader line. Numbers opened a minimal unmodified pie and accepted a native
resave, but rejected source-generated per-series pie mutation packages as
damaged, including both styled and geometry-only probe variants. This remains
a tracked Numbers series-non-style fixture limitation rather than claimed
native resave support for that path; Numbers, Pages, and Keynote were quit
after inspection.

The chart reference-line ownership slice is complete. The archive-free
`litchi-iwa-common::chart::reference_line` module now owns finite `Value`,
checked `Kind`, and compact `Line`; the IWA facade exposes the focused
`charts::reference_line` module and retains only archive/protobuf/graph
adapters. Nested custom-value patching retains unknown fields, recognized
fields reject duplicate, wrong-wire-type, and noncanonical framing, and the
reference-line graph is bounded before generated protobuf allocation. The
common suite passed 52 tests and the focused IWA reference-line suite passed 9
tests, including CRUD in Numbers, Pages, and Keynote. ZIP validation passed.
Native Pages and Keynote opened the generated line chart without repair
prompts and exposed its `Revenue thresholds` title and Revenue/Cost series;
Numbers rejected this source-generated chart fixture as damaged, which remains
a tracked Numbers chart-fixture limitation rather than native support evidence.
No native-resave claim is made, and Pages, Numbers, and Keynote were quit after
inspection.

The follow-on wire and payload-discovery hardening keeps the same ownership and
wire contract. Recognized reference-line fields now use the common checked
`WireField` payload and canonical key/length views, while unknown fields remain
permissive and wire-preserving. Pages, Numbers, and Keynote chart editors share
one allocation-free exact-one message-index scan, so malformed zero- or
duplicate-payload containers fail before chart decoding or mutation callbacks.
The common suite passed 53 tests, the focused reference-line suite passed 11
tests, and the exact-one scanner regression covers zero, one, and two matching
payloads. Formatting, diff, and scoped lint checks are the applicable gates for
this structural/performance slice; no additional native application run was
needed because serialized chart semantics were unchanged.

The reference-line graph preservation follow-up is now implemented in the IWA
archive codec. `set_reference_lines` performs a bounded, occurrence-aware
raw-wire merge for graph, axis, item, style, sparse-reference, `Reference`,
UUID, and axis-ID messages; it preserves unknown fields in their existing
positions and recursively validates every modeled known field before read,
update, or removal. The staged opaque-field candidate is validated before
assignment, and malformed deep-node updates remain atomic. The focused archive
module passed 10 tests, the reference-line semantic filter passed 15 tests,
the full IWA library passed 1,513 tests, and both affected crates passed
`-D warnings` clippy. Native Pages and Keynote opened the generated chart
fixture without repair prompts; Numbers rejected the source-generated chart as
damaged, which remains the tracked Numbers chart-fixture limitation above.
Pages, Numbers, and Keynote were quit after inspection.

The next bounded migration slice introduces the common source-bound wire
views. `litchi-iwa-common::wire::WireView<'a>` retains one borrowed source and
compact private spans, while `WireFieldView<'a>` provides canonical key,
length, framing, and payload checks without per-field byte ownership. Strict
reference-line readers now use this view through the thin IWA adapter; the
permissive mutation representation remains only in callers not yet migrated.
Singular wire overlays now index base and overlay fields once and emit one
exact-capacity output, removing the former quadratic reparse path while
preserving duplicate, wire-type, field-count, and output-size checks.

The Numbers editor layout follows the same staged ownership direction:
`numbers::editor::text_box_api` now contains the ordinary sheet text-box API
as a private child module, reducing the editor root without introducing a
compatibility layer. The common suite passed 58 tests, the full IWA library
passed 1,513 tests, strict `-D warnings` Clippy passed for both affected
libraries, and the crate-boundary audit remained valid. A fresh fixture built
from the changed branch opened in native Pages and Keynote without repair
prompts and exposed the authored Revenue/Cost 2D line chart with numeric Y and
categorical X axes. Native Numbers rejected the source-generated chart as
damaged; this remains the tracked Numbers chart-fixture limitation above.
All three applications were quit through their application menus after
inspection.

The table hidden-axis ownership slice now follows the focused table seam:
`litchi-iwa-common::table::axis::{AxisIndex, HiddenAxes}` owns the archive-free
row/column positions and the canonical duplicate-free hidden set. The set is a
single boxed slice sorted deterministically by row, then column, and duplicate
construction returns the typed axis-module error. The IWA implementation is
now a private archive adapter retaining native hidden-state UUIDs, protobuf
field mapping, package traversal, axis bounds, and transactional mutation.
Numbers, Pages, Keynote, and all five hidden-axis examples import the common
types directly; the former flat semantic definitions and contextual facade
aliases are gone. The focused common suite passed 3 axis tests, the IWA hidden-
axis suite passed 7 adapter/API tests, all five migrated examples compiled,
both changed crates passed no-dependency `-D warnings` Clippy, and the crate
boundary checker remained valid. No native application claim is made for this
structural ownership slice.

The Keynote slide-audio options slice now follows the archive boundary at
`litchi_keynote::slide::audio::Options`. The semantic value is a compact,
12-byte archive-free placement/duration pair: finite coordinates and positive
native `f32` duration representation are checked before package work begins.
The IWA adapter alone retains drawable and data identifiers, `TSD.MovieArchive`
decoding, graph lookup, zero-size geometry, raw wire updates, automatic build
objects, media records, and transactional publication. `KeynoteSlideAudioInfo`
and the removal result remain IWA-owned because they expose native IDs,
drawable properties, optional playback state, and package-GC disposition. The
shared playback settings continue to preserve explicit absence versus native
defaults, unknown loop values, and strict volume validation in their existing
IWA boundary. The old `KeynoteSlideAudioOptions` definition and re-exports are
removed rather than retained as compatibility aliases.

The next structural handoff isolates archive-free structured aggregation in
`litchi-iwa-structured`. `StructuredData` owns only the semantic table, slide,
and section vectors supplied by `litchi-numbers`, `litchi-keynote`, and
`litchi-pages`; native archive traversal remains in private `litchi-iwa`
extractors. The new leaf has no protobuf, package, ZIP, or facade dependency,
and the executable boundary policy records its three downward format edges.
The new leaf's focused tests, the IWA structured extractor tests, the boundary
policy tests, and strict no-dependency Clippy all pass. This is a crate-topology
handoff, not native application evidence; the existing native iWork fixture
matrix remains the authoritative verification for serialized packages.

The Numbers cell display-format seam is also complete for this migration
slice. `litchi-numbers::cell::data_format` now owns the archive-free checked
format values, while the IWA adapter retains native registry IDs, protobuf
codec details, BNC/control coordination, custom UUID registry handling, and
transactional package mutation. The focused Numbers leaf filter passed 15
tests, the IWA data-format filter passed 20 tests, strict no-dependency
Clippy and the IWA library check passed, and the regenerated verification
example compiled and ran. Native Numbers and Keynote opened the resulting
`table-number-formats.numbers` and `.key` files without repair prompts; their
accessibility trees exposed the expected formatted values, controls, and
custom formats.

The shared text-column ownership slice is complete. The archive-free
`litchi-iwa-text::columns` module now owns the focused `Columns`, `Count`,
`Gap`, `Width`, `Equal`, `Following`, and `Variable` values, with typed finite
validation, a 256-column budget, canonical negative-zero rejection, and one
boxed allocation for variable following columns. The IWA adapter retains only
`ColumnsArchive` decoding/encoding and native malformed-state mapping; Pages,
Numbers, Keynote, shape-style tests, and all three text-box creation examples
consume the leaf directly. The former flat `TextColumn*` owners and facade
reexports are gone. This structural slice is verified by the focused leaf and
library/example checks, strict no-dependency Clippy, and the crate-boundary
checks; the migrated conditional-highlighting test and table-number-format
example now use the canonical Numbers leaf markers. No additional native
iWork claim is needed for this ownership-only column change because serialized
column semantics are unchanged.

The Pages body-footnote selector slice is complete. The archive-free
`litchi-pages::footnote::body` module now owns bounded `Footnote`, UTF-16
`Position`, and source-order/position `Selector` values. `PagesEditor` no
longer returns or accepts native footnote IDs: its body-footnote CRUD resolves
selectors inside a staged transaction, while the private IWA adapter retains
reference/storage/marker graph identifiers and cleanup. The focused Pages leaf
tests and IWA body-footnote CRUD tests pass, and the migrated Pages examples
compile against selectors. This is an ownership/API slice; native Pages
verification remains part of the next fixture matrix run.

The archive-free IWA index foundation slice is complete. `litchi-iwa-index`
contains typed fragment/byte-span/object/reference values, deterministic
immutable indexing, typed duplicate/null/reference failures, and graph
queries without archive dependencies. Five leaf tests cover deterministic
multi-edge ordering plus fragment/source boundary rejection; strict Clippy and the
46-package/128-edge topology audit pass; private adapter integration remains
explicit follow-up work. The Keynote build leaf likewise passes its focused
semantic suite, strict Clippy, and formatting checks; its native adapter and
CRUD migration remain open, so this slice makes no new native Keynote claim.

The 2026-08-06 target-branch migration turn extends the same ownership direction.
`litchi-iwa::shapes::ShapePathKind` remains the structural path-family owner,
while `litchi-iwa-common::shape::path` owns only source-buildable path controls;
`litchi-keynote::ChartSelector` resolves chart edits by checked position or exact
visible name; `litchi-iwa-structured::StructuredData`
owns shared immutable semantic snapshots with borrowed text iteration; and
`litchi-pages::image::Options` owns validated image placement values. The archive
ingress/egress handoff is also recorded in `litchi-iwa-archive`, including ordered
logical opaque-entry storage, bounded legacy-package validation, and facade removal
of the direct ZIP dependency; byte-level ZIP-record preservation remains an ADR
0005 follow-up. The focused suites passed 113 common, 45 Keynote,
30 Pages, and 7 structured tests; the integrated facade library check and the
47-package boundary checker passed. Computer Use created and saved fresh native
fixtures at `/tmp/litchi-native-next-turn-20260806.pages`,
`/tmp/litchi-native-next-turn-20260806.numbers`, and
`/tmp/litchi-native-next-turn-20260806.key`; each reopened in its native iWork
application without a repair prompt, and Pages, Numbers, and Keynote were
closed after verification. This is structural/API and native-open evidence, not
a claim that the remaining monolithic semantic adapters or lazy physical catalog
have been completed.

The next 2026-08-06 migration slice adds three bounded ADR handoffs. The
archive leaf now retains immutable physical ZIP provenance (local and central
headers, timestamps, extras, comments, CRC, compressed ranges, and opaque
compression state) and exposes an exact byte-for-byte no-op write path; edited
entry reassembly remains an explicit follow-up. Format detection moved into the
archive-free `litchi-iwa-detect` leaf with typed limits/errors and a thin facade
adapter. `litchi-iwa-text` now owns checked UTF-16 positions/ranges and language
values, while native storage and protobuf traversal remain in the IWA adapter.
The bundle ingress path also preserves the catalog's lexical component order
without the former HashMap-to-Vec-to-sort allocation. The archive, detector, and
text leaves passed 14, 10, and 21 tests respectively; the full IWA library passed
1,489 tests; the bundle filter passed 20 tests; no-dependency leaf Clippy and the
boundary checker passed. Computer Use created and saved native Pages, Numbers,
and Keynote fixtures containing the migration markers, each opened natively
without a repair prompt, and all three applications were closed afterward.
The physical files are `/Users/ryker/CodeProjects/litchi/test-data/images/png/:tmp:litchi-native-archive-detect-20260806.pages`,
`/private/tmp/:tmp:litchi-native-archive-detect-20260806.numbers`, and
`/private/tmp/:tmp:litchi-native-archive-detect-20260806.key`. This remains a
bounded provenance/detection/value extraction slice; source-backed lazy object
access, weighted caches, edited-entry reassembly, paragraph-list extraction,
and full monolith removal remain open.

The following 2026-08-06 slice advances the crate split and typed text boundary.
`litchi-iwa-cache` is now a dependency-free leaf with weighted deterministic LRU
eviction, explicit invalidation, and generation-safe single-flight parsing; its
six concurrency, retry, weight, and eviction tests pass. It is deliberately a
standalone seam until package-state cache wiring can replace the current
single-entry adapter without mixing cache policy with archive ownership.
`litchi-iwa-text::paragraph::list` now owns checked list presets, bullets,
number formats, geometry, indentation, UTF-16 paragraph placement, and typed
errors; the IWA, Pages, Keynote, and Numbers adapters use `TextPosition` and
the leaf's semantic values. The former paragraph-level facade reexports were
removed so the module hierarchy is explicit and incremental-build friendly.
The text leaf passed 27 tests and production Clippy, while the migrated IWA
paragraph-list suite passed 51 tests and the IWA test-target check passed.
Stale resolved edges were removed from the boundary policy, a regression test
now rejects reintroducing them, and the checker reports 49 packages, 136
internal declarations, and 13 explicitly ordered debt items; its Python suite
passes 15 tests.

Computer Use created and saved fresh native fixtures through the real iWork
applications. Pages applied text bullets with a 110% bullet scale and saved
`/private/tmp/litchi-native-paragraph-list-20260806.pages`; Numbers saved a
table-cell marker to
`/private/tmp/litchi-native-paragraph-list-20260806.numbers`; and Keynote saved
the title/subtitle marker pair to
`/private/tmp/litchi-native-paragraph-list-20260806.key`. Each native window
displayed its saved file URL and expected semantic content without a repair
prompt, and all three applications were closed after verification. This is a
bounded leaf/topology and native-open slice; the remaining ADR debt includes
raw `IWorkTextEditor` object IDs and storage message exposure, eager
`Bundle`/`Document` materialization and missing `ReadAt` source identity,
weighted-cache integration into package state, edited ZIP-entry reassembly,
and the remaining monolithic semantic adapters.

The next 2026-08-06 slice wires immutable source access and bounded semantic
caching through the split. `litchi-iwa-archive::Catalog::from_read_at` and the
facade's `IWorkPackage::from_read_at` now read through positional `ReadAt`,
check the configured size before allocation, and reject a changed source with
typed `SourceChanged` errors. Flat catalogs retain their exact shared source
for validated no-op output; legacy nested catalogs deliberately normalize on
write. `PackageState` now uses the dependency-free `litchi-iwa-cache` weighted
LRU with decompressed-stream byte weights, per-key invalidation, copy-on-write
forks, and a bounded active-flight budget. The `litchi-iwa-text` leaf owns
`paragraph::border::{Sides, Offset}` and bookmark text semantics; the old flat
paragraph-border owners are gone, and list-bullet validation checks before
allocation. The archive, cache, and text leaves passed their focused suites
(20, 8, and 33 tests); the facade bookmark slice passed 10 tests, and the
workspace boundary checker remained green. Computer Use created and saved
fresh native fixtures containing the migration marker at
`/private/tmp/litchi-native-border-cache-20260806.pages`,
`/private/tmp/litchi-native-border-cache-20260806.numbers`, and
`/private/tmp/litchi-native-border-cache-20260806.key`; each was shown by its
native application and Pages, Numbers, and Keynote were closed afterward.
This is structural/API and native-open evidence, not a claim that edited ZIP
entry reassembly, raw `IWorkTextEditor` IDs/storage, full source-path
retention for byte-ingress `open`, or the remaining monolithic adapters are
complete.

The concrete-crate audit for ADRs 0001–0004 found no direct `litchi_iwa`
imports in `litchi-pages`, `litchi-numbers`, or `litchi-keynote`. It retained
three follow-up seams: Pages' flat document re-exports and opaque background
payload, Keynote's flat background re-exports and opaque payload, and the
Numbers `table::{Position, Range}` migration aliases. The first implementation
slice moved the archive-free Keynote slide-image insertion value from the
monolith into `litchi_keynote::slide::image::Options`. The leaf now validates
finite placement and strictly positive displayed/natural dimensions, stores
only three common geometry values, and reports typed image-specific errors;
the old `KeynoteSlideImageOptions` owner and facade re-exports were removed.
The focused leaf tests, four Keynote image CRUD tests, five image examples,
and strict Keynote Clippy passed. Table-lock ownership was intentionally
outside this audit scope.

This turn completes the next ADR 0001–0004 migration slice on the isolated
`feat/office-format-completeness` branch. `litchi-iwa-common::table::lock`
now owns the compact table-lock state consumed by the Pages, Numbers, and
Keynote facades. The neutral object index now stores each record once and
lets its format adapter borrow the record while retaining only source-local
position metadata. `litchi-iwa-text` owns the date-time, storage, and
paragraph drop-cap semantic values; `litchi-iwa-common::chart::error_bar`
owns checked error-bar values and bounded custom arrays. The Keynote image
options leaf, archive source retention, and `NumbersSheetInfo::id` complete
the corresponding public API handoffs without compatibility aliases.

`litchi-iwa-archive` now also has bounded physical ZIP reassembly for the
supported Store/Deflate edited entries, preserving untouched records and
metadata; unsupported legacy, ZIP64, and compression cases remain explicit
errors. Its 24 tests, detector's 10 tests, common's 118 tests, text's 40
tests, and the integrated IWA library's 1,489 tests passed. The crate
boundary checker reports 49 packages and 138 internal declarations, its
Python suite passes 15 tests, and the scoped leaf checks remain green. The
library-target Clippy checks for detector and Keynote also pass; all-target
`-D warnings` remains blocked only by pre-existing test-module unwrap, float,
and shadowing lints in those packages. The exhaustive detector mapping for
the new reassembly error and the date-time example's typed sheet-ID call are
included in the integration commit.

Computer Use verified native files through the real iWork applications.
Numbers opened `/private/tmp/litchi-native-table-lock-20260806d/table-lock.numbers`
and exposed `Locked Table` plus the disabled locked-table cells; the
date-time Numbers file exposed `Created: Friday, July 17, 2026`. Pages opened
the date-time fixture and exposed the same marker, while Keynote opened its
fixture and exposed the marker in a text box. These files were saved and
Numbers, Pages, and Keynote were closed afterward. The Pages table-lock
fixture at `/private/tmp/litchi-native-table-lock-20260806d/table-lock.pages`
was rejected by Pages as damaged, so it is recorded as a native-open gap,
not a pass. Remaining ADR debt includes raw `IWorkTextEditor` IDs/storage,
eager package materialization, fully source-backed lazy object access,
remaining monolithic adapters, and the intentionally unsupported archive
member classes above.

This 2026-08-06 slice completes the next text-leaf boundary migration. The
archive-free `litchi-iwa-text` crate now owns hyperlink, number-attachment,
ranged-comment, and paragraph-style values; the former IWA owners
`hyperlink_types.rs`, `number_attachment_types.rs`,
`paragraph_following_style.rs`, and `text_comment_types.rs` were deleted.
Each native identity is a compact opaque non-zero handle. Conversion to and
from a native archive identifier is available only through the explicit
`litchi_iwa_text::<leaf>::raw` adapter module, while normal semantic code has
no object-ID constructor or accessor. Ranged comments additionally keep their
records private, expose only checked accessors, and carry typed `Instant`,
`AuthorId`, and `Uuid` metadata; hyperlink and comment owned text validates
before borrowed input is allocated, and boxed native text is adopted without
another allocation. The IWA facade retains structured leaf error variants
instead of flattening these failures into `InvalidFormat(String)`.

The focused leaf suite passed 57 tests, the integrated IWA library suite
passed 1,485 tests, strict no-dependency Clippy passed for `litchi-iwa-text`,
and the IWA library Clippy target passed with the three pre-existing dead-code
groups explicitly allowed. The edited examples compile in their focused
targets; an all-example check still reports unrelated baseline API drift in
older table/image examples. An independent ADR audit found the initial
identity/metadata boundary issues; the changes above resolve those findings
without compatibility aliases. The workspace formatter remains noisy outside
the scoped files, so only changed-file formatting and `git diff --check` are
used for this migration turn.

Computer Use verified the generated fixtures in the real iWork applications.
Pages exposed `Page [attachment: 1] of [attachment: 1]` from
`/private/tmp/litchi-adr-pages-number-attachments.pages`; Numbers exposed the
`ADR leaf Numbers hyperlink comment` text box and its markup marker from
`/private/tmp/litchi-adr-numbers-text.numbers`; and Keynote exposed the
`ADR leaf Keynote hyperlink comment` text box and marker from
`/private/tmp/litchi-adr-keynote-text.key`. The Pages text fixture was also
opened during the same run. Pages, Numbers, and Keynote were saved and closed
after verification. This is a bounded semantic-leaf and native-open slice;
remaining debt includes the broader monolithic adapter split, raw storage
APIs outside these leaves, and the other ADR 0005/0006 resource work.

This follow-up completes the TextHighlight ownership handoff required by ADRs
0001–0004. `litchi_iwa_text::highlight` now owns the compact non-zero semantic
identity and UTF-16 range, including typed zero-ID rejection and a compact
`Option` representation. Native numeric conversion is confined to the focused
`highlight::raw` boundary; `litchi-iwa` retains annotation graph discovery,
protobuf validation, range-table mutation, package ownership checks, and
transactional CRUD. The former `litchi-iwa/src/text/highlight_types.rs` owner
was deleted, and the example's raw-ID operations now use the explicit adapter
boundary. No compatibility alias or archive dependency was added to the leaf.

The focused text suite passed 58 tests, the integrated IWA library suite passed
1,484 tests, strict no-dependency Clippy passed for both text and IWA library
targets (with the three pre-existing IWA dead-code groups explicitly allowed),
and the migrated highlight example compiles. CLI round trips exposed
non-zero highlight IDs and ranges in Pages (`0..8`), Numbers (`0..7`), and
Keynote (`0..9`). Computer Use opened the Pages and Keynote fixtures in the
real applications; their accessibility trees exposed the generated text and
their screenshots showed the turquoise native highlight. The newly generated
Numbers fixture was rejected by Numbers as damaged (an existing fixture
generation gap), so the known-good `/private/tmp/litchi-adr-numbers-text.numbers`
fixture was opened instead; its accessibility tree exposed the text box and
the CLI confirmed `TextHighlight { id: 131, range: 0..3 }`. Pages, Numbers, and
Keynote were closed after verification. Remaining text debt includes raw
`IWorkTextEditor` storage selectors and the broader monolithic adapter split.

This follow-up completes the text-appearance ownership handoff required by ADRs
0001–0004. The archive-free `litchi-iwa-text::appearance` module now owns the
focused `Outline`, `Shadow`, `Background`, and `ParagraphBackground` values.
They compose only the neutral `litchi-iwa-common` color, stroke, and shadow
primitives; no protobuf, archive, graph, package, or facade error state enters
the leaf. Text shadows retain the native text-inspector restriction to drop
shadows through a typed `UnsupportedShadowFamily` error. The IWA adapter keeps
all native color/stroke/shadow conversion, inheritance, null-marker validation,
style lookup, and transactional publication. The former flat `TextOutline`,
`TextShadow`, and `TextBackground` owners were removed rather than aliased, and
the internal property discriminants now distinguish character appearance from
paragraph background explicitly.

The appearance leaf passed 60 focused tests, the integrated IWA library suite
passed 1,484 tests, strict no-dependency Clippy passed for both
`litchi-iwa-text` and the IWA library target (with the three existing IWA
dead-code groups allowed), and the six appearance-related text examples
compiled. Independent ADR, common-API, migration-risk, and native-verification
audits agreed on the downward-only `litchi-iwa -> litchi-iwa-text ->
litchi-iwa-common` boundary and found no archive dependency in the new leaf.

The CLI text-style inspector read the migrated values from the known-good
native fixtures: Pages storage 147, Numbers storage 130, and Keynote storage
155 each reported a one-point `Stroke`, the standard one-point/five-point
opaque `Drop` shadow at 45 degrees, and the fixture's non-default solid
background. Computer Use opened `/private/tmp/adr-highlight-pages.pages`,
`/private/tmp/litchi-adr-numbers-text.numbers`, and
`/private/tmp/adr-highlight-keynote.key` in the real iWork applications. Pages
and Keynote accessibility trees exposed their marker text and screenshots
showed the native turquoise highlighted ranges; Numbers exposed the expected
text-box marker without a repair prompt. Pages, Numbers, and Keynote were
closed after verification. This remains a bounded semantic/API and native-open
slice; raw `IWorkTextEditor` storage selectors and the wider monolithic adapter
split remain open.

This 2026-08-06 slice extracts the structured text wire adapter into the new
`litchi-iwa-text-wire` crate. The leaf depends only on common wire errors,
generated IWA protobufs, and `litchi-iwa-text`; it converts one decoded
`TSWP.StorageArchive` into the canonical `Storage` value with one owned UTF-8
buffer, boxed semantic runs, checked fragment and allocation limits, and typed
errors. `litchi-iwa` now retains only the application-specific context/error
mapping and structured traversal, while the old 55-line facade-local
converter is deleted. No package, archive, graph, or application crate enters
the new leaf, and the boundary checker reports 50 packages and 142 internal
declarations with the existing 13 explicitly ordered OOXML debt edges.

The new leaf's three unit tests and doc-test target pass, strict no-dependency
Clippy passes for the leaf, the integrated IWA library passes 1,484 tests, and
the IWA library target passes strict Clippy with the three existing dead-code
groups allowed. ZIP integrity and the text-style inspector remain unchanged
for the known-good native fixtures: Pages storage 147, Numbers storage 130,
and Keynote storage 155 each retain their expected outline, drop shadow,
background, and highlight values. Computer Use opened those fixtures in the
real Pages, Numbers, and Keynote applications without repair prompts; their
accessibility trees exposed the expected Overview, Numbers hyperlink/comment,
and Quarterly result markers, and screenshots showed the native highlighted
content. All three applications were quit through their application UI and
confirmed absent from `sky.list_apps()` afterward.

This is a focused wire-layer split toward deleting the `litchi-iwa` monolith;
the dedicated IWA-owned `TextStorageId` migration, removal of public storage
message metadata, and the remaining monolithic adapters are intentionally the
next API/topology slices.

This follow-up completes the dedicated IWA text-storage identity seam. The
facade now owns `text::TextStorageId`, a transparent non-zero handle with
compact `Option` representation, checked parsing, and explicit native
conversion confined to the adapter. All 153 public `IWorkTextEditor` storage
selectors accept the typed identity; `TextStorageInfo` exposes only its typed
`id`, while message type and storage kind remain adapter-private. Pages,
Numbers, and Keynote text graph storage fields and their public storage
selectors were migrated accordingly. Raw `u64` remains only in private
archive/protobuf helpers and application graph IDs that have not crossed this
text boundary; no compatibility alias was added.

The scoped IWA library check, formatter, diff check, boundary checker, and
migrated text-inspection examples pass. The full unit-test build still has
older internal fixtures and older Numbers examples using raw native IDs; that
follow-on test/example migration is intentionally deferred rather than
weakening the typed public seam. Computer Use opened the existing known-good
Pages, Numbers, and Keynote fixtures in their native applications, confirmed
the expected text markers and highlighted content without repair prompts, and
closed all three applications afterward.

This 2026-08-06 follow-up completes the archive-free section-relative table
topology handoff. `litchi-numbers::table::topology` now owns the focused
`RowInsertion`, `ColumnInsertion`, `RowDeletion`, and `ColumnDeletion` values;
the Numbers, Pages, and Keynote IWA adapters consume those canonical types
directly, while native section counts, object identifiers, graph updates,
formula/merge maintenance, and wire publication remain adapter-owned. The
former `Table*`, `PagesTable*`, and `KeynoteTable*` topology facades and their
module files were removed rather than retained as compatibility aliases.

The semantic leaf and IWA library checks, strict no-dependency Clippy targets,
focused topology examples, topology unit test, diff check, and crate-boundary
checker pass. Native iWork verification for this slice is performed separately
with the real Pages, Numbers, and Keynote applications. Computer Use created
`/private/tmp/litchi-topology-pages.pages`,
`/private/tmp/litchi-topology-numbers.numbers`, and
`/private/tmp/litchi-topology-keynote.key`, inserted the marker
`Litchi topology verification`, reopened each saved document state, and found
no repair prompts. Pages, Numbers, and Keynote were then quit through their
application menus and confirmed absent from the final application list. The
broader stale internal fixture build remains deferred under the typed-selector
migration.

This follow-up completes the archive-free table-header semantic handoff.
`litchi_numbers::table::headers` now owns the focused `Count` and `Settings`
values consumed by Numbers, Pages, and Keynote. `Count` is a compact
`NonZeroU8` value with checked 1..=5 construction; `Settings` carries the
optional header-row, header-column, footer-row, freeze, and repeat semantics
without any archive or generated-protobuf dependency. IWA remains responsible
for native `u32` conversion, wire-field validation, bounds checks, transactions,
and publication. The former application-prefixed header facades and their
compatibility aliases were removed.

The semantic and IWA library checks, strict no-dependency Clippy targets,
focused header examples, Numbers unit suite, formatter/diff check, and
crate-boundary checker pass. The full IWA example target still exposes older,
unrelated stale native-ID fixtures and remains deferred under the typed
selector migration. Computer Use created and saved
`/private/tmp/litchi-headers-pages-native.pages`,
`/private/tmp/litchi-headers-numbers-native.numbers`, and
`/private/tmp/litchi-headers-keynote-native.key` in the real iWork
applications. Each native accessibility tree showed a 1/1/1
header-column/header-row/footer configuration and the `HEADER`, `BODY`, and
`FOOTER` markers. Pages, Numbers, and Keynote were then quit through their
application menus and the final app list contained only Finder, Terminal, and
ChatGPT.

This 2026-08-07 slice establishes the first constrained Buffa migration gate
without weakening the current production reader. Prost continues to generate
the complete compatibility surface while exact-version Buffa 0.9.1 eager and
lazy views are generated in a separate output directory for
`TSPMessages.proto` and `TSPArchiveMessages.proto`. The private, test-only
sidecar proves eager wire parity and lazy `ArchiveInfo` round trips. It is not
exposed to format crates or untrusted ingress: Buffa 0.9.1 does not charge all
deferred lazy-message range metadata to the configured element-memory budget,
and deferred children are not fully validated until access. The source-backed
raw wire layer therefore remains the preservation boundary, and Buffa coverage
will expand only with bounded format adapters. Limiting the committed seam to
the archive headers also avoids compiling the measured 73 MiB full-corpus
generated sidecar on every build.

Physical iWork ingress now maps every ZIP-indexing ceiling, including member
names, central-directory metadata, compressed entry size, uncompressed entry
size, aggregate size, and file count. Backend limit failures remain typed.
Keynote rejects oversized paths before allocating the source buffer, reads from
the same opened handle under a bounded loop, rejects growth past the selected
ceiling, and rejects an oversized borrowed slice before copying. The immutable
IWA index reserves its derived fragment tables fallibly, while the weighted
cache counts detached parsers until completion so invalidation cannot bypass
`max_flights`.

The Numbers document facade now accepts exact-name or checked-position sheet
selectors without exposing native identifiers. A real Numbers fixture also
exposed producer-padded row offset tables: missing-sentinel padding is accepted,
but any populated slot beyond the declared table width remains a typed format
error. The focused format suite passed 159 tests, the archive suite passed 26,
the cache/index suites passed 16, the Buffa/Prost parity suite and strict
generated-boundary Clippy passed, and the dependency-direction checker reports
62 packages, 213 internal declarations, and zero debt items. Format and archive
tests required `--cap-lints warn` only because of the existing five unused
scalar SIMD fallbacks in `litchi-core`; the changed focused leaf boundaries are
otherwise green.

Computer Use created checked-in native Pages, Numbers, and Keynote fixtures,
saved and closed them, then reopened each file in the corresponding real iWork
application without a repair prompt. Accessibility state confirmed the Pages
three-line marker, Numbers cells `B2` and `B3`, and the Keynote title, subtitle,
and date. Black-box format-crate tests open every fixture both by path and from
borrowed archive bytes, validate its semantic projection, and retain ZIP
integrity as a repeatable native compatibility gate.

## Production Buffa header seam and native preservation gate

The next 2026-08-07 migration slice promotes the constrained archive-header
sidecar into the first production Buffa path. Exact-version Buffa 0.9.1 lazy
views decode and encode `TSP.ArchiveInfo`, `MessageInfo`, `FieldInfo`, and
`FieldPath`; production `litchi-iwa-core` no longer imports `prost::Message` or
calls Prost encode/decode. A schema-directed common wire-tree preflight charges
aggregate scanned work, fields, nesting, message count, decoded memory, and
packed or unpacked metadata before the lazy adapter allocates. The adapter then
projects deferred children directly and fallibly, forces required
`MessageInfo.type`, `MessageInfo.length`, and `FieldInfo.path` presence, and
does not call `to_owned_message`. Buffa-generated values and errors remain
private behind neutral core types. Original hostile or noncanonical header
bytes remain the no-op preservation authority, including closed-enum negative
and noncanonical wire forms.

Physical package ingress now applies the constrained ZIP index before copying
raw local or central metadata, counts directory records against the physical
member limit, checks both physical name spellings and combined variable
metadata, rejects compressed size before materialization, and rejects a legacy
`Index.zip` declared size before decompression. Borrowed package input is
checked before its one owned snapshot allocation, and nested limit failures
remain typed. A corpus-locked preservation test rebuilds every decompressed IWA
component exactly: Numbers covers 37 components, 622 objects, 631 messages, and
373,043 bytes; Pages covers 7, 570, 576, and 360,855; Keynote covers 25, 959,
965, and 443,469 respectively.

At the semantic boundary, Numbers now decodes BNC type 9 as the native
discriminated union: rich-text or string identifiers produce text, otherwise
decimal or number fields produce a number. One allocation-free borrowed cell
view shares validation and scalar classification with the owned editor, while
metadata-only format application cannot silently convert the stored value.
Pages and Keynote expose exact-name or typed-position section/slide selectors
without raw IDs; names are distinct from headings and visible slide titles,
and duplicate names are typed ambiguities. Keynote maps the native navigator
name. The current Pages native projection intentionally exposes position
selection only until a real section-name field is identified; it never invents
identity from body text.

Computer Use opened the exact outputs produced by the migrated no-op example
in the real Numbers, Pages, and Keynote applications without repair prompts.
Numbers exposed the expected 22-by-7 table, marker, and value `42`; Pages
exposed the three expected body lines; Keynote exposed the title, body marker,
and date. Each application then duplicated and saved its own
`buffa-native-save` package under
`/private/tmp/litchi-buffa-noop.u5gndX/`. Those three app-authored packages
were fed back through the Buffa/Snappy/ZIP no-op path successfully, producing
three round-trip artifacts while preserving every decompressed IWA component
exactly.

Verification for this slice includes 130 common tests, 10 Buffa/Prost codec
tests, 1 core unit plus 21 hostile-framing tests, 35 archive units plus the
native preservation gate, 100 core tests, 75 Numbers units plus its native
fixture, 39 Pages units plus its native fixture, 51 Keynote units plus five
integration tests, and 18 focused monolith compatibility tests. Strict Clippy
passes for the production Buffa/common/core seam, the archive crate, and Pages;
the dependency checker reports 62 packages, 214 internal declarations, and
zero debt. Remaining migration debt is explicit: the generated Buffa sidecar
is about 2.7 MiB because Buffa selects file-level roots, format payload decoders
still use Prost, the monolithic `litchi-iwa` crate still exists, and native
Pages section-name extraction is not yet implemented. Buffa 0.9.1 also retains
internal infallible `Vec` growth, so these finite preflight budgets are not an
exact process-RSS or global out-of-memory guarantee.

## 2026-08-08 split-owner and adversarial-ingress slice

The next slice makes the monolith exit measurable. `litchi-iwa` is the only
declared migration host, and all 17 of its internal edges are ordered debt with
reasons and exit conditions. Three exact-parity physical examples—package
preservation, full IWA round trip, and decompressed package comparison—moved to
`litchi-iwa-archive`; structured extraction stayed in the host because direct
format packages still differ observably. The boundary checker now inventories
63 packages and 217 internal declarations.

Numbers BNC storage moved from `litchi-numbers::cell::wire` into the versioned
low-level `litchi-numbers-wire` crate. Numbers consumes it privately and the
legacy host depends on it directly; neither `litchi-numbers` nor the root
facade re-exports the codec. A compile-fail gate proves the old module is
private. The Numbers package error no longer exposes `prost::DecodeError`.
Format package and smart-detection `Debug` implementations are redacted so
ordinary logging cannot dump complete document sources or native catalogs.

Pages now resolves the native section graph: root field 5 supplies the initial
section, storage field 17 supplies later UTF-16 boundaries, and each reference
must resolve to one type-10011 section whose field 26 supplies the exact name.
The adapter validates unique references and boundaries, scalar positions, and
U+0004 break markers; it omits those markers while remapping text runs and
preserves absent versus empty names. The native fixture proves the authored
name `Blank`, and the root facade now traverses section text storages rather
than reporting a nonempty Pages body as zero paragraphs.

Archive headers still use the stock registry Buffa 0.9.1 dependency; no local
or git patch is required for downstream correctness. Core metadata is now
represented by core-owned `FieldPath`, `FieldInfo`, `FieldType`,
`UnknownFieldRule`, and `KnownFieldRule`. Optional presence and unknown signed
enum values round-trip exactly, generated public `From` implementations are
gone, and preflight charges both Buffa and neutral projection allocations.
Buffa differential tests cover duplicate merges, packed and unpacked scalars,
all unknown wire kinds, noncanonical encodings, every deferred child route,
and inclusive limit boundaries.

A separate five-file, at-most-32-KiB Buffa projection reads only repeated
field 3 of `TSWP.StorageArchive`. Its generated types remain private, unknown
retention is disabled, source bytes remain authoritative, and common-wire
preflight bounds bytes, fields, wire type, fragment count, UTF-8, aggregate
text, and repeated-view memory. The existing Prost-shaped conversion remains a
temporary compatibility oracle; focused format payload call sites have not yet
moved because unrelated known submessages are intentionally opaque in the
narrow projection.

Physical ZIP adversaries now cover independent local and central names,
nested central names, cumulative metadata, exact observed/maximum diagnostics,
and corrupt Deflate data whose declared nested `Index.zip` size must fail
before decompression. Native Numbers, Pages, and Keynote packages preserve
every decompressed IWA component through `to_bytes`, `write_to`, and empty
reassembly. Index snapshot construction reserves each derived table fallibly
and reports its precise allocation kind.

Computer Use opened the fresh Pages no-op and Numbers full round-trip artifacts
in real Pages and Numbers without repair prompts. Pages exposed all three body
lines; Numbers exposed the 22-by-7 table, fixture marker, and value `42`. Both
applications saved new copies under
`/private/tmp/litchi-iwork-split.tpXqvQ/`; those app-authored copies then passed
the bounded no-op path with zero changed decompressed components. A prior
Keynote-authored copy passed the same gate, and the original Keynote fixture
again exposed its title, body, and date in the real application.

The native gate rejected a proposed Keynote field-10 navigator-name
transaction even though internal parse/readback and unknown-byte tests passed:
both the focused prototype and the legacy editor caused Keynote to render
layout placeholder text and ignore the requested label. The concrete API,
example, and tests were removed. This is retained as negative evidence that a
semantic self-read is not a substitute for native application verification.

Focused verification includes 14 protobuf/Buffa tests, 1 core unit plus 28
framing tests, 12 text-wire tests, 38 archive units plus adversarial and native
preservation integrations, 8 index tests, 10 detector tests, 17 Numbers-wire
tests, 58 Numbers units plus its native fixture, 45 Pages units plus its native
fixture, 53 Keynote units plus five non-edit integration tests, and the root
Pages facade regression. The legacy 1,471-test library suite remains an
explicit compatibility gate. Remaining debt includes the monolithic editors and fuzz
targets, focused-format Prost payload decoders, format-owned error/limit
vocabularies, the strict Keynote storage-type audit, and a native-correct
Keynote name mutation design.

## 2026-08-08 Keynote skip-state transaction and native gate

The concrete Keynote owner now provides the first supported package mutation
outside the monolithic editor. `Slide::is_skipped` exposes playback omission as
a semantic Boolean. `Package::edit()` stages one bounded operation through
`set_slide_skipped`, `skip_slide`, or `include_slide`; callers select by exact
navigator name or checked zero-based `Position`, never by an IWA object ID.
`commit()` returns a named immutable package, compact diagnostics, and a
reversible patch. `Package::apply()` requires the exact retained source bytes,
checks the semantic precondition, fully reopens the retained target under the
source limits, and validates semantic readback. Equal-state commits share the
original source allocation and are byte-exact no-ops.

The private adapter requires one type-4 slide-node payload and a singular,
canonical varint field 4 whose value is exactly zero or one. It rejects missing
and duplicate occurrences, wrong wire types, noncanonical keys and values,
out-of-domain Booleans, ambiguous payload ownership, and candidate readback
mismatches. Canonical `false` and `true` have equal length, so mutation does not
change `ArchiveInfo.MessageInfo.length`: preserve-mode serialization retains
the original object header and changes exactly one decompressed IWA byte. ZIP
reassembly rewrites only the owning component and then the complete Keynote
package is reopened before publication. Generated protobuf values and native
component/object identities stay private; raw source bytes remain the
preservation authority.

The adversarial integration suite proves selector misses and duplicate-name
ambiguity are typed, missing-name display is redacted without cloning the
input, empty and second operations fail without publication, exact no-ops share
their allocation, the source snapshot never changes, unknown slide-node field
99, an unknown target `ArchiveInfo` field, and all non-target wire bytes remain
exact, and the complete decompressed target component has exactly one `0 -> 1`
byte change. In a target-last synthetic ZIP, unrelated local and central
records remain exact. The inverse patch restores the complete source artifact
byte-for-byte, unrelated exact sources conflict, patch debug output omits
fingerprints and byte lengths, and the public transaction values are
`Send + Sync`. The focused crate's all-target run contains 63 passing tests: 53
units, five skip-state adversarial tests, and five existing integration tests.
The former
`litchi-iwa` skip-slide example was removed and replaced by the selector-first
`litchi-keynote` example; the legacy editor method remains temporarily for
compatibility with the larger unmigrated editor surface.

Strict Clippy passes for the Keynote production library, the migrated example,
and the new adversarial integration target. Rustdoc with denied warnings and
the new crate-level doctest pass. A `keynote`-only root-facade integration test
proves the transaction is reachable without the migration host, and two legacy
skip/preservation regressions still pass. The dependency gate reports 63
workspace packages, 217 internal declarations, and the same 17 explicit host
debts; all 40 checker tests pass. The broader Keynote all-target Clippy command
still reports 33 pre-existing warnings in unrelated test modules, while none
comes from this slice.

Computer Use created a two-slide source and verified the generated output in
Keynote version 14.4 (7043.0.93). The source
`/private/tmp/litchi-keynote-skip-source-20260808.key` is 513,913 bytes with
SHA-256
`a5bd6289eaf1a82043585621d606f53f03719c4488eade6edf254055162fe05f`.
The Litchi-authored output
`/private/tmp/litchi-keynote-skip-output-20260808.key` is 513,916 bytes with
SHA-256
`cc7bb4fc4397efbb561047bca5548fba253d8c74468c87feb474e4f109e4c687`.
It opened without a repair, recovery, or conversion sheet; the navigator
identified slide two as `Skipped slide, Litchi SKIP TARGET 20260808`, and the
native Slide menu exposed `Unskip Slide`. Of 58 ZIP entries, only
`Index/Document.iwa` changed payload, no compared entry metadata changed, and
the decompressed comparison reported exactly one changed component with
unchanged archive metadata.

Keynote then saved the opened Litchi result as
`/private/tmp/litchi-keynote-skip-native-save-20260808.key` (513,870 bytes,
SHA-256
`503c5edcea5a42455b776326674ad56016d770ab99de090ac798973337b9267b`).
After close and reopen, the navigator still marked the slide skipped and the
menu still exposed `Unskip Slide`. Feeding that app-authored package back to the
focused example read `skipped true`, produced `false`, touched one component,
and emitted a valid ZIP. This distinguishes Litchi's source-preserving write
from Keynote's own package normalization while proving native semantic
round-trip compatibility.

Remaining debt is explicit. This in-memory patch is not yet the ADR 0003
durable JSON envelope, the focused crate does not yet provide an atomic
filesystem save API, most Keynote editor operations and compatibility tests
remain in `litchi-iwa`, and format payload projection still uses generated
Prost values. Promoting a bounded Buffa slide-node projection must preserve the
same raw-wire authority and cannot weaken the canonical/preflight or native
application gates.

- Stable Rust with workspace MSRV 1.89. The initial 1.85 placeholder was
  corrected because the workspace deliberately uses Rust 2024 `let` chains
  (stable in 1.88) and its measured x86 acceleration path uses stable AVX-512
  target features and intrinsics (stable in 1.89). Later bumps require a
  concrete safety, ergonomic, or measured-performance reason.
- Windows, macOS, and Linux CI; WASM where the selected I/O/crypto stack permits.
  `no_std` is not an initial support promise.
- Unit, integration, property, fuzz, Miri, sanitizer, malformed-corpus, and
  dependency-direction checks appropriate to each layer.
- Representative performance and resource budgets as defined by ADR 0005.
- Generated low-level schemas/records are deterministic, checked in, reviewed,
  and cite the source specification. Ergonomic facades remain handwritten.

## 2026-08-08 Keynote reachable-storage and semantic-budget slice

`litchi-keynote::Package` is now a production consumer of the focused
`TSWP.StorageArchive.text` Buffa projection. The package no longer scans every
message and speculatively Prost-decodes it as text. It requires the exact
document, show, slide-node, slide, placeholder/shape, note, and storage message
types while walking only reachable native references. A valid storage encoding
under an unrelated type is ignored, and duplicate storage payloads, ambiguous
placeholder/shape ownership, missing typed payloads, wrong text wire kinds,
and malformed known payloads fail closed. Only schema-proven type 2001 is
classified as `StorageArchive`; an incompatible native type-2022 sibling is
left opaque. Body and independent drawable text remain
`litchi-iwa-text::Storage` values with fragment ranges; title and notes move
into their existing plain-string semantic slots without a second text copy.
Plain-text order is title, visible body/drawable content, then speaker notes.

`ReadOptions` composes the retained physical archive limits with checked
format-owned limits for at most 1,000,000 native objects, 65,536 slides,
1,000,000 traversed references, 1,000,000 decoded storages, 1,000,000 retained
fragment ranges, and 64 MiB of aggregate semantic text. Callers may select
smaller non-zero limits but cannot exceed the hard ceilings. The package counts
objects before allocating a locator table, reserves the exact bounded capacity
fallibly, sorts compact locators once, rejects duplicate global identities,
and uses binary search for later reference resolution. Streaming show/slide/
build preflights apply slide, used-reference, name, and effect-identifier limits
before their corresponding generated vectors or semantic ownership conversions.
Required envelope presence is validated for every known graph payload consumed
by the adapter. Semantic and common-wire limit failures preserve resource kind,
observed/maximum counts, and a content-free semantic path. Slide records,
slide/build/storage vectors, and final plain-text output also reserve fallibly
after checked sizing. Skip-state commits and patch application retain both
physical and semantic profiles when reopening candidates.

The new adversarial integration target has thirteen passing tests. It proves
exact-limit acceptance and typed over-limit rejection for objects, slides,
references, storage count, fragment ranges, and aggregate UTF-8;
pre-materialization build/transition identifier rejection; package-ingress
rejection of duplicate identities; strict storage type/wire and owner
cardinality; exclusion of both a valid unreachable false-positive storage and
an incompatible type-2022 sibling; required proto2 envelope rejection;
deterministic concurrent first access;
preservation of synthetic fragment ranges; and native type-2001 Buffa/Prost
differential text on the checked-in Keynote fixture. The complete Keynote
all-target suite passes 84 tests. Strict Clippy passes for the Keynote
library, rich-storage and skip-state targets, the selector-free inspection
example, and the archive library; rustdoc and the workspace boundary gates
remain separate required checks before publication.

Computer Use created a new presentation in Keynote 14.4 (7043.0.93) containing
a title, a two-line body with a formatting boundary, an independent text box,
a footer, and presenter notes. Keynote saved, closed, and reopened
`/private/tmp/litchi-keynote-buffa-native-20260808.key` without a repair,
recovery, or conversion sheet and exposed every authored value after reopen.
The artifact is 519,374 bytes with SHA-256
`b40162d851b29de328f8ee04f32ee2e090852169c2028b29d96da7dd3cd2063b`.
The focused `inspect_text` example reported one slide, 964 objects, and all six
semantic text values. The archive-owner preserve-mode example emitted an exact
519,374-byte no-op with the same hash. Because this slice is read-only, no
Litchi-authored semantic output was presented to Keynote; source bytes remain
the preservation authority.

Remaining Keynote debt is explicit: the larger graph still uses generated
Prost messages, formatting/style tables are not yet a semantic rich-text
model, the slide-node skip path remains raw-wire rather than Buffa-projected,
and complete allocation-envelope preflight for ignored nested generated fields
still depends on the physical message ceiling. Most legacy Keynote mutations,
examples, fuzz targets, and parity tests remain in `litchi-iwa`. The monolith
deletion gate therefore remains open.

## 2026-08-08 Numbers rooted/global table parity gate

`litchi-numbers::Package` now distinguishes the ordinary rooted workbook from
the historical global structured-table projection. Rooted decoding requires
one canonical type-1 document, type-2 or type-3 sheets, and typed table-info
owners; it preserves sheet/drawable order, ignores typed false positives under
unrelated messages, and rejects duplicate sheet, drawable, or table-model
ownership. The native type-6000 and fixture-backed legacy type-6003 table-info
forms are explicit and mutually exclusive. Referenced table models prefer type
6001 and use the legacy type-6000 fallback only when canonical ownership is
absent.

The allocating `extract_structured_tables` method replaces the host algorithm
without changing ordinary semantics. It indexes only each object's first
message, runs the type-6001 group before type 6000, sorts by identity inside
each group, deduplicates candidates, includes valid physically retained
orphans, errors on malformed preferred models, and skips type-6000 candidates
that fail complete model extraction. The compact index counts objects before
fallible exact reservation, stores one sorted locator plus at most one primary
type entry per object, rejects cross-component duplicate identities, and uses
binary search for reference lookup. New read options preserve archive limits
and add hard-bounded object, sheet, table, and rooted-reference profiles with
content-free semantic paths.

Focused tests cover canonical root typing, duplicate payloads and identities,
secondary-type exclusion, unrelated false positives, malformed preferred and
legacy candidates, exact/exceeded object, sheet, and rooted-reference budgets,
and strict rooted ownership.
Four migration-host differential tests cover detached models, reversed
drawable/global order, canonical-before-legacy order with dual-kind
deduplication, object-vector stability, and exact/exceeded structured table
budgets. The public `read_numbers` example demonstrates sheet/table traversal
and the explicitly separate compatibility view without accepting a raw ID.

Computer Use authored and reopened the real Numbers workbook
`/private/tmp/litchi-numbers-order-oracle-20260808.numbers`. It contains two
reordered sheets, three named tables and marker cells, and a non-table text box.
After the native Arrange operation, focused readback reported rooted order
`B-only-table`, `A-new-table`, `A-old-table` and compatibility order
`B-only-table`, `A-old-table`, `A-new-table`; the final file is 133,740 bytes with
SHA-256
`781181e89c655da5c92b677b9ba5c939c85379e7b33ccf10e3846fe8588f9c5b`.
Numbers closed and reopened it without a repair or conversion prompt and
retained both table markers and the reordered sheet/table UI.

Remaining Numbers debt is explicit: the legacy aggregate structured API has
not yet been switched from its `litchi-iwa` adapter, the focused table decoder
still materializes the wider generated Prost graph, aggregate sidecar
allocations need deeper schema preflight, and nearly all mutation paths remain
in the migration host.
The focused order/orphan parity gate is therefore evidence for the next host
edge removal, not a claim that the monolith may already be deleted.

## 2026-08-08 Numbers compatibility-ingress prerequisite gate

The global compatibility projection now has a byte-owned entry point separate
from strict rooted `Package` construction:
`compatibility_tables_from_bytes[_with_options]`. It parses and validates one
immutable package snapshot, proves unambiguous Numbers ownership with the
focused manual-wire detector on the unique canonical type-1 payload, builds
the compact index, and runs the global table projection without constructing
the strict rooted workbook. Formula-name enrichment may resolve root sheets
and drawables lazily only after a non-empty formula sidecar is selected. A
bounded schema-directed preflight covers the canonical detector shape;
wire-type guards preserve ambiguous field-number collisions such as the
Numbers calculation-engine reference. Numbers-shaped noncanonical siblings
cannot establish or mask ownership or multiply detector work. A regression
with a valid Numbers root referencing a missing sheet proves that strict
`Package::from_bytes` fails while the rooted-independent empty compatibility
projection succeeds.
Pages and Keynote roots return the typed `NotNumbers` category; unknown and
canonically mixed application payloads fail closed.

Index and candidate admission are stricter at the focused boundary. Object
identifier zero is rejected as the native null sentinel. A primary type-6000
object no longer consults a secondary type-6001 payload, and duplicate
canonical or legacy model payloads are rejected. Formula-label lookup likewise
selects one canonical model payload, with a unique legacy fallback, instead of
scanning arbitrary siblings. The table ceiling wins before
decoding an additional canonical candidate, an intentional allocation-safety
precedence. Package protobuf errors retain the decoder as an error source
behind a Numbers-owned wrapper whose display text contains no native content.

Formula-reference maps moved from eager extractor construction to lazy
initialization after a non-empty formula sidecar is selected. The builder uses
the compact binary-search index, fallible map growth, `Arc<str>` sheet/table
names, and a caller ceiling for unique source-derived formula-enrichment
entries, including the table-name cache, categories, and owners. The fixed
`Grand Total` fallback is not input-derived and is exempt. Topology discovery,
candidate visits, cumulative source text, and encoded category bytes use
separate package-wide hard ceilings.

Type-6383 category bytes are now schema-preflighted before Buffa access. The
recursive topology scan and per-node UUID/CellValue scans share one aggregate
field/message work counter; every routed scalar field, wire kind, and string is
validated before label retention. A private generated projection contains only
an empty node envelope, UUID, and four scalar wrappers. Recursive children are
streamed directly from source bytes, avoiding both an input-width generated
fragment vector and per-node heap ownership. Ignored row, aggregate,
formatting, and error payloads remain opaque. An O(depth) iterator stack walks
children, while duplicate UUID labels retain the former last-wins behavior.
Raw-wire depth, aggregate wire bytes, nested-field work, empty fanout, wide
children, malformed projected siblings, duplicate labels, and Buffa/Prost
differential behavior cover this boundary.

Path ingress also opens once (nonblocking on Unix), rejects non-regular
descriptors, fills bounded `Vec` spare capacity directly through the standard
reader path, caps its initial allocation independently from an advisory logical
length, and stops after a one-byte over-limit lookahead. Descriptor metadata is
compared after the read so truncation, growth, or an observable in-place
modification fails atomically. Tests cover exact length, growth after metadata,
an already-oversized input, a changed descriptor version, and a non-regular
source.

The focused suite now covers direct application smuggling, mixed roots, null
identities, rooted-independent compatibility behavior, secondary candidate
promotion, duplicate canonical/legacy payloads, lazy formula enrichment,
pre-decode formula limits, application-field wire collisions, descriptor
growth, and `Package: Send + Sync`. The checked-in native fixture exercises
both the package method and the new byte entry point. The derived Buffa
projection generates five files and 146,678 bytes under a 160 KiB build gate;
its generated types remain private.
Computer Use reopened the unchanged native order oracle in Numbers without a
repair prompt, verified navigator order `SecondCreated`, `FirstCreated`, table
`B-only-table` with marker `B-only`, and the `A-old-table`/`A-new-table` layout,
then closed without saving.

The legacy aggregate API is deliberately not rerouted by this slice. Its
parsed `Bundle` does not retain a shareable source/catalog or the selected
physical profile; reopening would fail byte and directory inputs, permit
time-of-check/time-of-use drift, duplicate decoding, and change rooted/global
error timing. Aggregate compatibility budgets for cells, sidecars, text,
formula work/rendering, and a typed host-error bridge remain open gates.
Formula metadata is currently initialized for a non-empty formula sidecar;
deferring it until a rendered cell contains a cross-table or category node is
also still required before the aggregate host edge moves. Other table models,
formula owners, sidecars, and formula ASTs still use generated Prost values and
need their own complete pre-decode envelopes.

## 2026-08-08 Keynote Show/SlideTree lazy projection and slide-order migration

The concrete Keynote owner now projects `KN.ShowArchive` through a private
Buffa lazy view after a Keynote-owned schema-directed wire preflight. Required
theme, slide-tree, size, and stylesheet envelopes, optional show settings, all
known references, wire kinds, canonical scalar framing, required-envelope
uniqueness, and required nested reference identifiers are validated before
semantic publication. Ordered slide-node references are streamed from the
preflighted embedded `KN.SlideTreeArchive`; they are not represented as a
nested generated repeated-message view. This avoids an input-width Buffa
fragment index while preserving source order exactly. Optional setting
presence and unknown presentation-mode discriminants retain their existing
semantic behavior. Unknown content is not retained by the generated view, and
the accepted raw source remains authoritative for preservation.

`Package::edit_slide_order()` adds a bounded selector-first structural
transaction without changing the existing skip-state patch API. The source is
an exact navigator name or checked base position. The destination is the final
zero-based `Position` in the base list and must be below the original slide
count. `SlideOrderEdit` stages one move; `SlideOrderCommit` publishes a fully
reopened package; `SlideOrderDiagnostics` reports changed state, touched
components, and full-reparse publication; and `SlideOrderPatch` provides
exact-source-checked forward application and a reversible inverse.
`SlideOrderError` and `SlideOrderLimitKind` keep the operation and its resource
failures format-owned and content-free. Equal source and destination positions
reuse the original allocation and bytes. A changed order moves complete raw
slide-reference field records, including each encoded key, encoded length, and
nested reference payload. It preserves their unknown and deprecated fields and
does not rewrite slide components.

The focused implementation replaces the migration host's
`KeynoteEditor::move_slide`. Its move-specific compatibility assertions move to
the Keynote owner, and the raw-index-only host example is replaced by the
selector-first `litchi-keynote` `move_slide` example. Add, duplicate, and
remove-slide workflows are not folded into this transaction because they also
own component registration, identifier allocation, dependency disposition,
and reclamation.

The root integration and native acceptance gates completed on 2026-08-08:

- `cargo test --locked --offline -p litchi-iwa-protos` passed 38 unit tests;
  its 86 generated doctest snippets remain intentionally ignored. The focused
  codec portion passed 13 tests, including known-field canonicality, permissive
  opaque-unknown framing, direct message/recursion caps, native Prost parity,
  and a 4,096-reference bounded traversal.
- `cargo test --locked --offline -p litchi-keynote --all-features` passed 67
  unit tests, 37 integration tests, and 2 doctests. The slide-order target
  contributes 12 tests covering all 16 four-slide move pairs, selectors,
  exact/no-op/inverse patches, full-record preservation, flat-topology
  refusal, exact/one-under slide limits, fail-early source validation, and
  concurrent immutable commits. `cargo test --locked --offline -p litchi-iwa
  --lib` passed all 1,478 migration-host compatibility tests. The direct root
  Keynote facade passed 2 tests, the aggregate iWork facade passed 8, and
  `litchi-iwa-structured` passed 12.
- Warning-denied Clippy passed for `litchi-iwa-protos --all-targets`, for all
  Keynote production/library/example targets, and for the complete slide-order
  integration target. `cargo fmt --all -- --check` and `git diff --check`
  passed. `cargo check --locked --offline -p litchi-iwa --lib` passed. The
  broader host `--examples` check remains blocked by the pre-existing
  `list_iwork_chart_radar_grid_shapes` use of private Numbers sheet IDs; that
  same failure is present at the starting commit and is not represented as a
  Keynote regression or as a passing gate here.
- The generated Show projection is 1,682 source bytes and exactly five Buffa
  0.9.1 output files totaling 138,661 bytes, with zero `LazyRepeatedView`
  mentions. Build-time provenance matches the canonical `TSP.Reference`,
  `TSP.Size`, `KN.SlideTreeArchive`, and `KN.ShowArchive` declarations and the
  handwritten route constants. `tools/check_crate_boundaries.py` reports 63
  workspace packages, 224 internal declarations, and exactly 17 ordered debt
  items. `tools/check_iwork_public_api.py` reports no implementation type or
  raw-ID leak.
- The retired host implementation was run from detached commit `df1b76b5`
  against the same disposable source. Both host and focused outputs read back
  as `B/C/A`, and their extracted `Index/Document.iwa` bytes are identical
  (SHA-256
  `9ecd2426425491053898658f5b7584d0633b30d3a3b020bf226d397f7693d310`).
  The host artifact hash is
  `0172045ef824a5061564e013568f1343cbcb95a149a179bb8953a3bcac8842ff`;
  the focused artifact differs because it retains source ZIP metadata that the
  host normalized to the 1980 epoch.
- Apple Keynote 14.4 (7043.0.93) created the three-slide `A/B/C` source at
  `/private/tmp/litchi-keynote-order-oracle-20260808.B6vCko/source-abc.key`
  (SHA-256
  `49c7ee349cddb9fcd4671b7cd36c90008a76e457311cd3bb70d4b765f217b3df`).
  Moving position zero to final position two produced
  `litchi-moved-bca.key` (SHA-256
  `62960a755535fd719bffa53f6f9e9f6126fa22d2ae50c3b543e24f926da07779`).
  Keynote opened it directly with visible `B/C/A` navigator order and no
  repair, recovery, or conversion prompt. Save As produced
  `keynote-resaved-bca.key` (SHA-256
  `81f2e6010f68504fc58b2c948604f05f3651e3252ddba10c98b7eee29aed16e9`);
  close/reopen again showed `B/C/A`, and the focused reader recovered the same
  titles and bodies. Applying the public inverse patch restored a byte-exact
  artifact with the original source hash. ZIP-member comparison found only
  `Index/Document.iwa` changed; its source member hash is
  `505d0666be4a7711f952b8b21fea97bc9f54c67ade145f4095aad2843d08d7de`.
  Decompressed comparison likewise found exactly that component and only Show
  object 2652385 changed, with its archive metadata equal.

This is not a latency, RSS, allocation-performance, fuzz, or sanitizer claim.
All 17 migration-host dependency debts remain. Larger slide-node, slide,
build, drawable, note, table, chart, media, and mutation graphs still use host
code and/or generated Prost values; protobuf groups remain fail-closed at the
shared package preflight; durable patch serialization, an aggregate edit
transient-memory ceiling, atomic filesystem publication, remaining
example/test/fuzz ownership, and the root sanitizer campaign are open deletion
blockers.

## 2026-08-08 focused Keynote show-settings and graph-boundary evidence

The concrete Keynote package now exposes `Package::show_settings()` as a
direct bounded reader. It validates the complete known Show and SlideTree
envelope, including slide-reference limits, but forces only the private Buffa
size and scalar projection. The focused-reader regression proves equality with
`Package::show()?.settings()` while leaving the full semantic slide cache
uninitialized and retaining no slide-node identifier collection. A null root
show maps to `Settings::default()`; because no physical Show owner exists, only
its exact no-op edit is accepted.

`Package::edit_show_settings()` covers size plus all eight optional scalar
settings: slide-number visibility, looping, mode, autoplay transition delay,
autoplay build delay, idle-timer activation, idle-timer delay, and automatic
play-on-open. Tests exercise absent-to-present and present-to-absent changes,
canonical negative `int32` mode encoding, checked invalid sizes/delays/modes,
duplicate and wrong-wire rejection, exact no-op allocation identity, exact
patch conflict checks, inverse replay, retained non-default `ReadOptions`,
content-redacted `Debug`, and `Send + Sync` public values. Candidate
publication performs a full retained-options reopen and direct semantic
readback.

Preservation tests keep the immutable source artifact, all untouched ZIP
member records, non-setting Show fields, nested unknown Size fields, and
unrelated messages exact, including unchanged encoded field keys and length
headers. Only the owning `Index/Document.iwa` component is rewritten. A
changed payload's effective message type and length,
`MessageInfo`, and necessary enclosing framing belong to the mutation closure;
the test does not incorrectly require those dependent bytes to remain fixed.
Unknown preservation comes from raw source-record rewriting, not from Buffa.
Legacy nested-`Index.zip` input supports direct reading and an exact no-op but
returns `UnsupportedSource` for a change. Its former normalizing writer remains
in the migration host, so the host method, example, and compatibility tests are
not retired by this slice.

The object-index continuation removes only ordered dependency debt 007. The
host now sends authoritative `MessageInfo` references and schema-directed
fallback references straight to `litchi-iwa-index::IndexBuilder`. Focused
tests retain null filtering, the rule that a non-empty authoritative list
suppresses fallback even when its values are null, idempotent duplicate
handling, deterministic order, dangling-target observability, and strict
duplicate rejection for the ordinary builder method. Graph identities and the
immutable snapshot are consumed through the index owner's reexports; the
canonical `litchi-iwa-index -> litchi-iwa-graph` edge remains.

Executed Rust evidence is deliberately reported by stable target rather than
as one frozen moving-suite total:

- `litchi-iwa-core` passed 31 tests, `litchi-iwa-protos` passed 38, and
  `litchi-iwa-index` passed 9.
- The focused Keynote `show_settings` integration target passed all 11 tests;
  the migration-host library passed all 1,479 compatibility tests; the direct
  root Keynote facade passed 3; and Keynote passed 3 doctests.
- The final Keynote all-features/all-targets run passed 68 library tests and
  48 integration tests across its eight integration binaries.
- Warning-denied Clippy passed for the scoped changed Keynote, protobuf, index,
  focused test, and example targets. A full Keynote dependency traversal stops
  on 88 pre-existing `litchi-core` ARM SIMD lint failures; it is recorded as a
  baseline blocker, not a passing gate or a regression from this slice.
- Formatting, whitespace/diff, and the supported rustdoc public-API checks
  passed.
  `tools/check_crate_boundaries.py` reports 63 workspace packages, 223
  internal dependency declarations, and exactly 16 explicit debts. Identity
  007 is absent and later debt identities remain unchanged.

Computer Use exercised the exact-source transaction in Apple Keynote 14.4
(7043.0.93). The source
`/private/tmp/litchi-keynote-show-settings-20260808.g4cipH/source.key` has
SHA-256
`f3adcde9315b6df580805bcb63c995cc1e1ef569a4befa06a102485e13c883b2`.
The pristine Rust candidate was reproduced after the final code gate as
`/private/tmp/litchi-keynote-show-settings-20260808.g4cipH/final-rust-reproduced.key`
with SHA-256
`c8364bb21713892f6c3c5dfb37207f8d293f48010ad16c1ff3da0547ea9f0644`;
its public inverse reproduced the exact source hash. These are the same
candidate bytes originally presented to Keynote. The opened working path after
Keynote's in-place autosave is
`/private/tmp/litchi-keynote-show-settings-20260808.g4cipH/final-self-playing.key`
with SHA-256
`a106977db366e794be087a87ddfd874e7af3c26fa84d9fb5d573ca74efec739a`.
Keynote opened and automatically played the artifact without a repair,
recovery, or conversion prompt. Its inspector showed Self-Playing mode, loop
enabled, automatic play on open enabled, 1920-by-1080 Widescreen, a five-second
transition delay, and a two-second build delay.

Keynote Save As, close, and reopen succeeded and produced
`/private/tmp/litchi-keynote-show-settings-20260808.g4cipH/final-keynote-resaved.key`
with SHA-256
`a9109add346eb26c8a9cb6f7db7e6bd6f1a6366a6ba1c9d073ac1c7c64bc6857`.
Focused reverse-read recovered the expected settings. Focused no-op and inverse
outputs over this final native artifact were byte-identical to
`a9109add...`. Before native normalization, the Rust patch inverse restored the
exact `f3adcde9...` source. The Rust-authored package retained the ZIP
entry-name set and changed content only in `Index/Document.iwa`; that comparison
uses the pristine `c8364bb2...` reproduction rather than the app-autosaved file.

This is not an O(1), single-pass, latency, RSS, allocation-performance, fuzz,
sanitizer, complete Buffa-laziness, or full host-retirement claim. The legacy
normalizing settings path, most Keynote editors and Prost graph projections,
durable patch serialization, aggregate transaction peak-memory policy, atomic
filesystem publication, remaining examples/tests/fuzz work, the sanitizer
campaign, and the 16 remaining host debts remain explicit deletion blockers.

## 2026-08-08 amendment: Pages section-name transaction evidence

The concrete Pages owner now exposes selector-first exact-source section-name
edits, exact patch application and inverse replay, while the root `pages`
facade reexports the same canonical types. Focused tests cover exact-name and
position selection, longer and shorter names, clearing, explicit empty
presence, missing and ambiguous selectors, NUL rejection, duplicate
destinations, exact no-op allocation identity, changed-legacy refusal,
header/unknown-field preservation, one-member mutation, exact conflict checks,
inverse byte equality, tight output limits, redacted `Debug`, and `Send + Sync`.
The checked native fixture, synthetic adversarial packages, and root facade are
all exercised.

Executed Rust evidence is reported by stable target: the Pages
all-features/all-targets run passed 52 library tests, 1 native-fixture test,
and all 5 section-name integration tests, and built the focused example target.
The direct root Pages facade passed both tests. Warning-denied Clippy passed
for every Pages target and the focused root facade; warning-denied rustdoc,
the migration-host library check, formatting, diff checks, the iWork public-API
gate, and the boundary checker also passed. The boundary checker reports 63
workspace packages, 223 internal dependency declarations, and exactly 16
ordered debts. This scoped evidence does not represent the known baseline-red
full workspace and migration-host example inventory as green.

The preservation oracle keeps the immutable source, untouched member data,
raw local records, raw central records apart from required local-header
offsets, non-target messages, unknown section fields, and unknown IWA object
header fields exact. A changed section message's length and the enclosing IWA
and ZIP lengths/offsets form the intended mutation closure. Candidate
publication performs a full retained-limit reopen and field-by-field semantic
section verification. Unknown preservation comes from bounded raw-record
rewriting; no Buffa re-encode is involved.

Computer Use exercised the Rust-authored candidate in Apple Pages 14.4. The
pre-application Rust candidate had SHA-256
`9269594edd2ac2c13e1ed04780cf6bc5b3734b1cd4f42067d280de658aca1696`;
its public inverse restored the exact source hash
`21107bc9323fba6f1589152454c0b0b0cc8e239313c6a369bc4a891116601b42`.
Pages opened the candidate without repair, recovery, or conversion and showed
the unchanged fixture markers `Litchi native Pages fixture`, `Buffa lazy-view
migration verification`, and `2026-08-07`. Native Save As, close, and reopen
produced
`/private/tmp/litchi-pages-section-name-20260808.RM2Cz3/pages-resaved.pages`
with SHA-256
`fe879ecc03e0a3673a911c5b7d335f4d0e54c0766590f51d17e7670e8ce1b194`.
Restaging `Litchi Renamed Section` against that native artifact was an exact
no-op and emitted the same hash, proving focused semantic reverse-read after
native normalization. Pages does not expose the section's producer name in
its accessibility tree, so the claim is intentionally limited to native
acceptance plus focused reverse-read rather than visual name inspection.

This is not full host retirement, a durable patch, atomic filesystem
publication, an aggregate transaction peak-memory policy, fuzz/sanitizer
completion, or a performance result. The host legacy-normalizing compatibility
writer and the 16 ordered dependency debts remain deletion blockers.

## 2026-08-08 amendment: Pages section-pagination transaction evidence

The concrete Pages package now reads and edits native section-start,
continue/restart, and starting-page-number fields through `SectionSelector` and
the presence-preserving `Pagination` value. A strict bounded `WireView`
preflight rejects duplicate selected fields, wrong wire types, noncanonical
keys or values, values outside `u32`, and page zero. A private generated Buffa
lazy view then independently projects only fields 20--22 with zero unknown-field
and repeated-element retention. The projection is provenance-checked against
`TPArchives.proto`, has a 1 KiB source ceiling, generates five files under a
64 KiB build ceiling, and has no production encode path.

The mutation rewrites one selected section payload in one component. Unchanged
recognized records and every unrelated or unknown field record remain exact;
changed recognized records keep their source position, and newly present fields
are appended in numeric order. The complete IWA object header is preserved with
the bounded core replacement helper, the ZIP is reassembled with one entry
edit, and the candidate is fully reopened under the source limits before
semantic publication. Exact no-ops retain the original `Arc` and bytes, public
inverse replay restores the exact source, and changed legacy nested-`Index.zip`
sources fail as `UnsupportedSource`. The legacy settings and background writers
were also moved to the same header-preserving bounded replacement helper.

Focused adversarial coverage exercises all 27 absent/default/nondefault
combinations, exact-name and position selectors, missing and ambiguous
selection, duplicate fields, wrong wire types, noncanonical keys and varints,
page zero, changed-legacy refusal, unknown payload/header/ZIP-record
preservation, one-component mutation, patch conflict and inverse behavior,
retained output limits, exact no-op allocation identity, redacted diagnostics,
and `Send + Sync`. The executed Pages all-features/all-targets gate passed 53
library tests, one native-fixture test, five section-name tests, and all six
new pagination tests. The root Pages facade passed three tests;
`litchi-iwa-protos` passed 41 tests; and both focused migration-host regressions
passed. Warning-denied Clippy passed for all Pages targets, the protobuf crate,
and the focused facade. Warning-denied rustdoc, formatting, diff checks, the
iWork public-API gate, and the boundary checker passed. The boundary checker
reports 63 workspace packages, 223 internal dependency declarations, and 16
ordered debts. The unrelated root `litchi` manifest remains rejected by the
repository-wide sort check and is not represented as newly green.

Computer Use exercised the Rust-authored artifact in Apple Pages 14.4. The
candidate
`/private/tmp/litchi-pages-pagination-20260808.3xZUeE/pagination-right-restart-7.pages`
has SHA-256
`41f7bfc700a4d6342f5fc5f1324574775c5c1fa77d7edc2ab4d520d7da7b3737`;
its public inverse restored the exact fixture hash
`21107bc9323fba6f1589152454c0b0b0cc8e239313c6a369bc4a891116601b42`.
Pages opened it without repair, recovery, or conversion, retained all three
fixture body markers, labeled the canvas page 7, and showed restarted numbering
with `Start at: 7` in the Section inspector. Save As, close, and reopen produced
`/private/tmp/litchi-pages-pagination-20260808.3xZUeE/pages-resaved.pages`
with SHA-256
`0a9f7e1238f295fee38da2745a77c9bd92101341a783d8ce75389c72a2be3abe`.
Focused reverse-read recovered right-page start, restarted numbering, and page
7; restaging those settings emitted a byte-identical no-op with the same hash.

This is not full host retirement, durable patch serialization, atomic
filesystem publication, an aggregate transaction peak-memory policy, a
latency/RSS/allocation result, or fuzz/sanitizer completion. The remaining
Pages editors, host compatibility surface, examples/tests/fuzz inventory, and
all 16 ordered dependency debts remain deletion blockers.

## 2026-08-08 amendment: focused Keynote slide-transition evidence

The concrete Keynote package now owns selector-first read, set, native-none
clear, exact patch application, and inverse replay for an existing modern slide
transition. `Package::slide_transition`, `edit_slide_transition`, and
`apply_slide_transition` exchange only `transition::Settings`, semantic slide
selectors, and dedicated transition transaction types. All optional scalar
presence, future effect identifiers and enum values, opaque color and timing
curve payloads, and the nine effect-specific custom fields round-trip through
that archive-free value. A clear retains delay, automatic start, random seed,
and writing direction; writes `Transition`/`none`/one second; and removes
effect-specific values. A legacy-only envelope is readable as no modern
settings but is not synthesized or normalized by a changed edit.

A 2,347-byte derived schema supplies five private Buffa lazy-view messages for
`SlideArchive.transition`, its required attributes, all modern animation and
custom fields, and `SlideNodeArchive.hasTransition`. The build checks every
projected declaration against `KNArchives.proto`, prohibits repeated generated
storage and production encoding, and caps the Buffa 0.9.1 output at five files
and 224 KiB; the measured generated closure is 207,203 bytes. Strict handwritten
preflight precedes Buffa and rejects missing required envelopes, duplicate or
wrong-wire selected fields, noncanonical keys/lengths/varints/bools, excessive
message bytes, and excessive recursion. Buffa borrows strings and opaque
payloads and does not retain unknown fields. Validated caller-owned raw records
remain the sole preservation and rewrite authority.

The transaction patches all 25 modern leaves in their original nested records
without whole-message Prost or Buffa encoding. A caller-limit preflight bounds
every intermediate output and the aggregate repeated parse/copy work before
those compatibility patch helpers allocate. The complete IWA message header is
replaced with the bounded header-preserving helper. The slide-node cache marker
is validated even for a no-op and updated atomically with the slide payload; a
co-located slide and node rewrite one physical component, while a clear in the
native split topology rewrites the two actual owners. Candidate publication
then reopens the complete package under retained `ReadOptions`, validates it,
and performs focused semantic readback. Exact no-ops share the source `Arc` and
touch zero components. Changed legacy nested-`Index.zip` sources fail as
`UnsupportedSource`; exact source bytes, not the diagnostic fingerprint,
authorize patch application.

Adversarial tests cover every modeled field and presence state, future values,
opaque payload validation, native-none normalization, stale and malformed node
markers, missing/duplicate/wrong-wire/noncanonical transition records,
selector miss and ambiguity, exact no-op identity, split and co-located
components, retained limits, changed-legacy refusal, unknown nested fields,
hostile IWA headers, untouched ZIP records, patch conflict, inverse byte
equality, redacted `Debug`, root-facade availability, and `Send + Sync`. The
legacy host writer remains for compatibility, but now validates and
synchronizes `hasTransition`, preserves both message headers, and verifies its
round trip; no host deletion is claimed by this slice. Path ingress in the
migration host was also hardened to stream from one opened file handle under
the configured byte ceiling, detecting growth past metadata before
publication, while save-target errors no longer disclose destination paths.

Executed Rust evidence is reported by stable target. `litchi-iwa-protos`
passed all 47 unit tests. Keynote passed 68 library tests and 54 integration
tests across its nine binaries, including all six transition transaction
tests, and its four doctests passed. The direct root Keynote facade passed all
four tests, and the migration-host library passed all 1,481 compatibility
tests. Warning-denied Clippy passed for the changed protobuf, Keynote,
migration-host, focused example/test, and root-facade targets with dependency
lints excluded. Warning-denied Keynote rustdoc, formatting, diff checks, the
iWork rustdoc public-API gate, and the crate-boundary checker passed. The
boundary checker remains at 63 packages, 223 internal dependency declarations,
and exactly 16 ordered debts. The broader Keynote all-target Clippy invocation
remains baseline-red on 35 pre-existing test-only unwrap, float-comparison, and
shadowing findings and is not represented as a passing gate.

Computer Use exercised both Rust-authored paths in Apple Keynote 14.4
(7043.0.93). The native one-slide Dissolve source is
`/private/tmp/litchi-keynote-transition-20260808.udONyg/native-source-dissolve.key`
with SHA-256
`ab186d8d59c858e1b3c2596fd45463cec75ddd92e9fda9032da656a940e68dca`.
The final reproducible Magic Move candidate is
`final-rust-magic-move.key` (`d5d24386cb544374f4c26da4349f7be961be34180a4536578616886a56af8c1a`)
and the native-none candidate is `final-rust-cleared.key`
(`5235a3d03dbabced6d06a03b4873826da8602d97f478c61f6467b35d732a08e5`).
Each public inverse restored the exact source hash. The pristine Magic Move
candidate changes only `Index/Slide-2652150.iwa`; the clear additionally changes
`Index/Document.iwa` because its slide-node marker moves from true to false.
All other ZIP member payloads and the complete entry-name order remain exact.

Keynote opened the Rust candidate without a repair, recovery, conversion, or
warning dialog. Its transition inspector showed Magic Move, duration 2 s,
automatic start, and delay 2.25 s. The clear candidate showed `No Transition
Effect` while retaining automatic start and the 2.25-second delay. Native Save
As, close, and reopen reproduced both inspector states. The saved Magic Move
artifact is `keynote-resaved-magic-move.key` with SHA-256
`8443a71e58199df4506ebb0896323721e6debdd32c8a055f98f56d98d48cf7ac`;
the saved clear artifact is `keynote-resaved-cleared.key` with SHA-256
`e65429be7aa0bfd69c20ee1b3b17b86c9bd9c46b44af1f6db127fc017981c444`.
Focused reverse-read and restaging over each native artifact produced an exact
zero-component no-op with the same respective hash.

This is not full host retirement, durable patch serialization, atomic
filesystem publication, a complete aggregate transaction peak-memory model, a
latency/RSS/allocation-performance result, or fuzz/sanitizer completion. Prost
is still used privately to validate bounded opaque Color and PathSource
payloads, the legacy transition writer and broader Keynote graph editors remain
available, and all 16 ordered dependency debts remain deletion blockers.

## 2026-08-08 amendment: focused Pages section-text transaction evidence

`litchi-pages::Package` now exposes selector-first `section_text`,
`edit_section_text`, single-section `edit_body_text`, exact patch application,
and inverse replay. Its public root and the root `litchi::pages` facade export
the format-owned edit/commit/patch/diagnostic/error/limit family plus the
shared `TextPosition` and insertion-capable `TextSpan`. The production example
`litchi-pages/examples/edit_section_text.rs` parses paths with `args_os`,
accepts exact-name or semantic-index selection and set/clear/range modes,
constructs ranges as checked UTF-16 spans, publishes with create-new,
`write_all`, and `sync_all`, and can publish an exact inverse. It imports no
migration-host, raw-ID, generated-message, or Prost API.

The shared `litchi-iwa-text-wire` rewrite kernel performs one bounded splice
over raw `TSWP.StorageArchive` records. Its resource profile independently
caps input and output bytes, fields, nesting, text fragments and text bytes,
all positional-table entries, reference occurrences, and aggregate rewrite
work. It preserves wholly untouched text fragments and unknown wire records,
uses UTF-16 scalar boundaries, updates all recognized position-bearing tables
according to their native policy, and reports both field-specific and
aggregate reference deltas. A semantic no-op validates the source but retains
its exact bytes instead of normalizing it. Pages adds selector resolution,
section-relative to body-relative mapping, reserved-structure and dependent
content refusal, exact-provenance authorization, one-component reassembly,
complete retained-limit reopening, and semantic/topology readback.

The new 1,235-byte derived Pages body schema contains three singular messages
for document body references, section-boundary entries, and references. Buffa
0.9.1 emits the expected five-file 93,867-byte generated closure; the build
caps it at 96 KiB, prohibits generated repeated storage and production
encoding, and checks every routed declaration against the canonical
TP/TSWP/TSP schemas. Strict preflight rejects missing required reference
identifiers, duplicate selected singulars, wrong wire types, noncanonical
selected keys/lengths/varints, zero identifiers, excessive fields/work/bytes,
and insufficient recursion before forcing Buffa's lazy values. Buffa is a
borrowed semantic cross-check and never the preservation representation.

Eight focused integration tests cover astral UTF-16 replacement and boundary
shifts, exact and legacy no-op source sharing, the single-section body
convenience on the native fixture, raw unknown field and IWA-header
preservation, untouched ZIP members, exact inverse restoration, selector and
span failures, surrogate splits, reserved controls, hidden dependent
references, set/clear/delete neighbor preservation, retained output limits,
and malformed known-table refusal. The shared kernel adds deterministic
coverage for cross-fragment edits, insertion affinity, all positional-table
policies against a Prost oracle, conservative reference provenance,
noncanonical untouched framing, and typed malformed input. The focused
transaction test, production example, and root facade pass warning-denied
Clippy with dependency lints excluded. The new crate-root doctest compiles and
passes.

The direct root Pages facade passes all four tests, including compile-time
availability of every `SectionText*` type and both shared UTF-16 values,
exact-source no-op sharing, a name-selected insertion, and exact inverse
restoration.

The production example also executed all three CLI modes against
`test-data/iwork/pages/basic.pages`
(`21107bc9323fba6f1589152454c0b0b0cc8e239313c6a369bc4a891116601b42`).
The set, clear, and empty-span range outputs are retained under
`/private/tmp/litchi-pages-example.KdlErn/` with respective SHA-256 values
`0170017eec66373a428ee4a2599b50e4cfe5008d54cbc0daef66cb67e4307ffb`,
`63c2aa20f6064b9a8c5a536475d1a71b34175f4c6924a4d384f24c39fd5155e6`,
and `dd0405249a56e3e2b535e6a9541f02feda6299ce1a0959f4d68f7e44a0ae307a`;
each separately published inverse is byte-identical to the source.

The native acceptance gate used Pages 14.4 to open the Rust-authored set output
without a repair warning and visibly rendered the complete three-line value
`Litchi Pages text migration 🚀`, `Buffa lazy view verified`, and `2026-08-08`.
Pages then saved that document as
`/private/tmp/litchi-pages-text-20260808.Z5SKz0/pages-resaved.pages`
(`e2e7a9e67e499e2f8f003f091c8ade7fabbe2c59a577fcfe90fd3bd69022965a`),
closed it, and reopened that exact path without repair or loss of the requested
text. The Rust migration-host reader recovered the exact three-line value from the
Pages-resaved artifact. Running the focused example's semantic no-op and
inverse paths against that native-resaved package produced byte-identical
artifacts with the same SHA-256. The original changed transaction's inverse
also remained byte-identical to the source fixture.

This native run covers the single-section `set` path on `basic.pages`. The
multi-section boundary-shift path has deterministic synthetic coverage, but an
app-authored multi-section artifact, native clear/range operations, and rich
footnote/inline-object refusal cases remain explicit native gates.

This work does not claim durable patch serialization, atomic filesystem
publication, aggregate peak-memory measurement, fuzz/sanitizer completion, or
full host retirement. The raw-ID migration-host methods remain compatibility
surfaces, changed legacy nested packages remain unsupported by the focused
transaction, changed no-root/fallback bodies remain unsupported, and all 16
ordered dependency debts remain open.

## 2026-08-08 amendment: cache boundary, focused Numbers projection, and Pages clear/range gate

The preceding 16-debt status is historical and is superseded here. Cache-backed
`PackageState` moved from `litchi-iwa` to the physical
`litchi-iwa-archive` owner. The archive owns bounded physical parsed-component
state, while `litchi-iwa-cache` remains a dependency-free leaf and the host
retains format and error policy. Direct host-to-cache debt identity 003 is
retired without renumbering later identities. The checked boundary inventory is
now 63 workspace packages, 223 internal dependency declarations, and 15
ordered migration debts.

The Numbers change is intentionally limited to focused
`TableInfo.tableModel` reads. A strict small private Buffa projection replaces
those eager Prost reads only after bounded raw-wire preflight and requires a
nonzero table-model reference. It performs no encoding, retains no unknown
content, and stores no repeated fields; accepted raw source remains the
preservation authority. This is not a broader table-model or Numbers graph
migration.

The previously outstanding native Pages clear/range gates passed in Pages 14.4.
Pages opened
`/private/tmp/litchi-pages-example.KdlErn/clear.pages`
(`63c2aa20f6064b9a8c5a536475d1a71b34175f4c6924a4d384f24c39fd5155e6`)
and
`/private/tmp/litchi-pages-example.KdlErn/range.pages`
(`dd0405249a56e3e2b535e6a9541f02feda6299ce1a0959f4d68f7e44a0ae307a`)
without repair. The clear result was visibly empty. The range result displayed
exactly these three lines: `Range prefix: Litchi native Pages fixture`, `Buffa
lazy-view migration verification`, and `2026-08-07`.

Native Save As, close, and reopen produced
`clear-native-resaved-20260808.pages`
(`3ba278e1934688c653ab73f1ee2a194f670545dd160aa5d8e33c2054463a9676`)
and `range-native-resaved-20260808.pages`
(`74072d9d813282618db8e47f7ebc26cc59f7c17b1abf9d22c5bbf5473b942a9f`).
Focused semantic reread matched each expected result, and both the focused
no-op and inverse over each native-resaved artifact were byte-identical to its
respective hash. These observations close those two native gates only; they do
not claim the remaining dependent-content, multi-section, durability,
performance, fuzz/sanitizer, or host-retirement gates.

## 2026-08-08 amendment: Keynote speaker-notes transaction evidence

The supported Keynote package now reads and edits text in an existing
speaker-notes graph through semantic slide selectors and checked UTF-16
positions. Set, clear, insertion, deletion, replacement, exact patch
application, and inverse replay share one transaction boundary. Publication
requires unique package-wide ownership, exact known note/reference shapes,
canonical selected wire framing, complete retained-limit reopening, semantic
readback, and native topology readback. Changed publication rewrites one
component; exact no-ops report zero touched components and retain the original
immutable source.

The private notes projection contains five generated Buffa files totaling
151,735 bytes under a 160 KiB cap and no repeated view. Eight focused codec
tests cover Prost parity, opaque metadata, required and duplicate fields,
wrong wire types, noncanonical framing, zero or malformed references, and
resource ceilings. Twelve transaction tests additionally cover astral UTF-16
edits, exact no-op and inverse identity, unknown/header/ZIP preservation,
duplicate and aliased metadata ownership, dependent and unknown note shapes,
malformed-metadata replay, retained limits, and unrelated noncanonical outer
object prefixes. The production example supports semantic set, clear, and
range replacement with optional inverse output; the migration-host raw-ID
example is removed.

The app-authored Keynote 14.4 source is
`/private/tmp/litchi-keynote-buffa-native-20260808.key` with SHA-256
`b40162d851b29de328f8ee04f32ee2e090852169c2028b29d96da7dd3cd2063b`.
The public example produced set, range, and clear candidates under
`/private/tmp/litchi-keynote-notes-final.4JHtRJ/` with respective SHA-256
values
`8a1468b3f5706df983770d9ab6cded55321b9e1b9e57edd638ea7335d7be122f`,
`227a24d6843468115a5219a9f3a25f565e5d256698d3563591a6501bf9e8d7e1`,
and
`79187ac2da5c42ce13e58bb5b94b1d1a2aa543807a4aa5dde63c5f92dbc73342`.
Each changed only `Index/Slide-2652150.iwa`, retained the complete ZIP entry
list and one-slide/964-object topology, passed ZIP integrity, and produced an
inverse with the exact source hash.

Computer Use exercised all three candidates in Apple Keynote 14.4
(7043.0.93). Keynote opened each without a repair, recovery, conversion, or
warning prompt, retained the surrounding slide markers, and displayed the
requested three-line Unicode set value, the requested two-line range value,
or an existing empty notes pane. Native Save As, close, and reopen produced
`set-native-resaved-20260808.key`, `range-native-resaved-20260808.key`, and
`clear-native-resaved-20260808.key` with respective SHA-256 values
`8b8187a6f4e27b15461e0b0dafe90b3fe62b020ad9597bb4c5e2023d1ac76d9b`,
`86f57c855a5b300001c2ecbcab4278ef693c29184f211aed3e7c3936e1169051`,
and
`d359ebbbec7f98a9b30aac809f1ff467814a2cff986cd5934af3022d86cc5c2a`.
Focused semantic reread recovered the exact values. Over every native-resaved
artifact, an identical edit remained a zero-component byte-identical no-op;
a real temporary edit touched one component, and its inverse restored the
corresponding native-resaved hash exactly.

The checked boundary inventory is now 63 packages, 221 internal declarations,
and 15 ordered debts. This evidence is deliberately bounded: only the selected
notes owner references use Buffa lazy views, and only text mutation in an
existing notes graph has moved. Notes graph creation/deletion, remaining host
APIs and examples, legacy normalization, durable patch serialization, atomic
publication, aggregate peak-memory policy, fuzz/sanitizer completion, and
complete `litchi-iwa` deletion remain open.

## 2026-08-08 amendment: aggregate-budget and native-read gate

Focused regressions now lock Pages empty-root behavior, rooted text-fragment
preservation, section-name byte charging, exact retained-versus-rendered text
limits, and exact observed-byte propagation to the root error. Keynote
regressions lock empty valid shows and null shows as empty semantic results.
The neutral aggregate additionally tests document-backed Pages section text,
Keynote show titles, owned unknown identifiers, static effect names, and public
text-order diagnostics.

One bounded root `parse_iwork` AddressSanitizer/libFuzzer campaign completed
152,219 executions in 61 seconds with no crash, timeout, or OOM (coverage
7,454; feature count 12,062; 566 MiB RSS). This closes only that bounded
root-ingress run, not focused deep-message or complete fuzz verification.

Computer Use opened read-only disposable copies of the documented Pages,
Numbers, and Keynote fixtures in the matching Apple applications. Pages
rendered its three documented lines; Numbers exposed a 22-by-7 table with the
documented text and numeric `42`; Keynote rendered its documented title, body,
and date. No repair, recovery, conversion, or warning prompt appeared, each
document was reported locked, and post-close SHA-256 values remained exactly
`21107bc9...1b42`, `f225d5b1...b693`, and `3a3d0747...b9f42`. Writable
disposable copies were observed to normalize silently on open, so future
read-only verification must keep this permission and hash discipline. The
workspace originals were never opened and retained the same hashes.

The migration ledger intentionally remains at 63 packages, 221 internal
declarations, and 15 ordered debts. The host-to-structured seam cannot be
retired until its five Numbers compatibility oracles have focused/root owners:
detached models, type-9 numeric values, global ordering, canonical-6001 before
legacy-6000 deduplication, and inclusive/exceeded table limits. No host
deletion, complete Buffa conversion, edit/save compatibility, performance, or
exhaustive-fuzz claim follows from this amendment.

## 2026-08-08 amendment: Numbers oracle transfer and structured cutover

The five prerequisite Numbers oracles now have focused and root ownership.
The focused suite checks a deterministic 535-byte fixture byte-for-byte against
`test-data/synthetic-iwork/numbers/compatibility-oracles.hex` (SHA-256
`352ca6ad6891c7222f76cdb5fe48178f1efb340dc82ab5bc6755b71a2d2595bc`).
It proves detached-versus-rooted behavior, an independently encoded finite
decimal128 type-9 value with exact `f64` bits, rooted and package-global order
under physical reordering, canonical type-6001 precedence over legacy
type-6000 with single emission, and inclusive three-table versus exceeded
two-table limits. Root-only tests repeat the public semantic contract and add
exact content-free error tuples for table and late detached-text limits.

Legacy type-6000 admission now uses a bounded wire fingerprint before decode.
Unrelated payloads remain skippable, while a model-shaped payload propagates
malformed, limit, and allocation failures. Rooted model admission preserves
duplicate-error precedence, charges the table budget before decode, and uses a
fallible reserve. Numbers publishes a format-owned resource-error taxonomy;
the root maps it exhaustively without depending on `litchi-iwa-common`.

After those transfers, the host structured module, method, re-export, tests,
support hooks, and dependency were deleted. Verification recorded 100 focused
Numbers unit tests plus five compatibility oracles and its native/prepared
integrations; 77 root library tests plus three cutover tests; 1,470 host unit
tests, the generated-roundtrip integration, and 23 host doctests. Boundary
policy reported 63 packages, 220 internal declarations, and 14 ordered debts;
47 boundary-checker and 10 public-API-checker tests passed. Production host,
root, and neutral structured strict Clippy gates passed, as did a focused
Numbers gate with the repository's enumerated existing lint allowances.
Root `parse_iwork` and host `parse_iwa` sanitizer fuzz targets built
successfully; the earlier bounded root campaign remains the execution
evidence rather than a claim of exhaustive fuzzing.

Computer Use opened a locked disposable copy of the app-authored Numbers 14.4
order oracle (`781181e89c655da5c92b677b9ba5c939c85379e7b33ccf10e3846fe8588f9c5b`).
It showed `SecondCreated` before `FirstCreated`, with the expected
`B-only-table`, `A-new-table`, and `A-old-table` markers and an unrelated
non-table drawable. No repair, recovery, conversion, warning, or save prompt
appeared; close preserved the exact hash. Focused reread produced the same
rooted order and the documented package-global compatibility order.

This evidence closes debt 011 only. It does not complete `litchi-iwa`
deletion, whole-graph Buffa lazy decoding, durable patches, atomic native
publication, deep focused fuzzing, edit/save compatibility, or aggregate
peak-memory work. In particular, root source preparation currently
materializes unrelated ZIP members transiently and remains a measured memory
debt.

## 2026-08-09 amendment: Keynote title/body transaction gate

The focused Keynote package now distinguishes title and body roles while also
distinguishing an absent native placeholder from an existing empty storage.
Selector-first reads and one-operation checked UTF-16 edits share one
role-aware transaction. A changed edit commit proves exclusive graph
ownership, rewrites the selected storage, invalidates the selected slide-node
preview state, deletes root preview entries, reassembles the package, reopens
it under the retained limits, and performs semantic and topology verification.
The storage and slide node may share one IWA component or occupy two. A changed
`apply_slide_text` does not reassemble: it exact-source checks the patch,
reopens the patch's stored target bytes under the retained limits, runs the
same verification, and reports the one- or two-component count retained from
the originating edit. Exact edit no-ops rely on the immutable selected snapshot
established when editing began; exact patch no-ops check artifact identity.
Both share the source allocation, retain all preview/cache bytes, report zero
components, and deliberately do no whole-source validation, reassembly, or
candidate reopen. Inverse replay uses
the same exact-source checked patch-apply path and restores the complete
original artifact.

The exact Buffa ownership seam consists of two projections. The existing
speaker-notes codec now also exposes optional `KN.SlideArchive` field-5 title
and field-6 body references while retaining its bounded style, transition,
name, in-document, and note snapshot. The new placeholder codec follows
`KN.PlaceholderArchive` through the required `TSWP.ShapeInfoArchive`,
`TSD.ShapeArchive`, and empty `TSD.DrawableArchive` envelopes and returns only
the optional `ShapeInfoArchive.owned_storage` field 4 and placeholder kind.
The selected read forces the slide view. Package-wide ownership proof instead
raw-scans fields 5 and 6 of every slide candidate and forces the slide Buffa
view only for a candidate that references the selected placeholder. It
raw-scans every placeholder candidate and forces the placeholder Buffa view
only when a modern or deprecated edge can reference the selected storage. The
shared bounded scanner also audits `ShapeInfoArchive.deprecated_storage` field
2, `ShapeInfoArchive.text_flow` field 3, standalone shape-info references,
embedded `TSP.Reference` metadata, and `NoteArchive.containedStorage` field 1.
The alias scan does not force the Buffa `NoteArchive` view. The existing
`litchi-iwa-text-wire` storage codec and raw splice remain the text-value seam.

Schema-directed preflight rejects missing required envelopes, duplicate
selected fields, wrong wire types, noncanonical protobuf framing, zero or
malformed references, and resource overruns before Buffa values authorize the
operation. Format-side checks separately reject contradictory role hints,
shared ownership, aliases, dependent content, and noncanonical outer IWA object
lengths. Buffa neither encodes the changed payload nor retains unknown content;
accepted raw records remain authoritative.

Changed publication has one deliberate preservation exception for rendered
caches. A bounded raw `KN.SlideNodeArchive` rewrite removes fields 3 and 9
`database_thumbnail`/`database_thumbnails`, field 10 `thumbnailSizes`, field 16
`thumbnails`, and field 25 thumbnail digests; sets field 14
`thumbnailsAreDirty` to true; removes the corresponding preview object
references; and prunes only preview-owned aggregate and field data references.
Proven unrelated data references are retained, while ambiguous aggregate-only
ownership is rejected. Other slide-node fields remain raw-preserved. The
selected `KN.SlideArchive` must not contain field 37
`thumbnailTextForTitlePlaceholder` or field 38
`thumbnailTextForBodyPlaceholder`; a changed commit fails closed until those
separate string caches have a proven invalidation rule.

`litchi_iwa_archive::package::Catalog::reassemble_with_deletions_to_bytes`
provides the physical publication seam. It resolves edits and deletions by
exact normalized name, rejects missing, duplicate, overlapping, legacy, or
ZIP64 mutation shapes, bounds the complete output, permits deletion of an
opaque member without decoding it, and preserves retained members' raw names,
metadata, ordering, and compressed bytes. The Keynote transaction deletes each
existing root `preview.jpg`, `preview-micro.jpg`, and `preview-web.jpg` through
that API. Preview entries are ZIP members rather than IWA components and are
not included in `touched_components`. Native Keynote and package preview
consumers can otherwise continue presenting pixels rendered before the text
change. Storage, slide-node, and preview mutations become visible only together
as one candidate artifact.

Diagnostics are intentionally component-scoped. A changed edit commit and a
changed forward or inverse patch application report `changed = true`,
`touched_components = 1` or `2` from the originating edit, and
`full_reparse_performed = true`. Root preview deletions do not increment the
component count. Exact edit and patch no-ops report `false`, `0`, and `false`
respectively.

Verification admits only the selected storage rewrite, the selected slide-node
cache invalidation, and deletion of those three root preview names. It requires
the slide node to be observably dirty and free of preview references, all three root preview
names to be absent, all other IWA objects to compare exactly, and unselected
semantic slide state to remain unchanged. A changed inverse restores the exact
source bytes, including any former root previews and slide-node cache records.

The new placeholder projection generates exactly five files totaling 141,766
bytes under a 144 KiB cap and contains no repeated view. Adding the two
singular slide references leaves the existing slide-owner projection at five
files and 162,241 bytes under its 168 KiB cap, also with no repeated view. The
current focused inventory contains ten placeholder-codec tests, ten
speaker-codec tests, 25 slide-text integration tests, and six Keynote
doctests. Together they cover Prost parity; optional and unknown fields;
required envelopes; nonzero, duplicate, noncanonical, and malformed
references; exact resource ceilings; title/body role separation;
absent-versus-empty state; UTF-16 operations and boundaries; exact no-op,
conflict, and inverse behavior; raw/header/sibling preservation; ownership
metadata and aliases; cache invalidation and root-preview deletion; content-free
errors; and retained semantic/output limits. The archive package inventory also
covers deletion-only and combined edit/deletion publication, opaque-member
deletion, exact/disjoint selection, retained physical records, output limits,
and the exact empty-mutation path. The final integrated run passed all 89 proto
unit tests, all 72 Keynote library tests, all 25 slide-text integration tests,
all six Keynote doctests, 78 archive tests, 42 core tests, seven root-facade
tests, three focused root library tests, and 258 filtered migration-host
Keynote regressions.

The migration-host builder regression now focused-reads and sequentially edits
title, body, and notes, asserts that the other roles remain untouched, and
reopens the final values through the host. The add-slide regression covers a
focused edit and reopen of the new title; the visible slide-number example is
also in the verification inventory. The duplication fixture was hardened with
a valid style and complete archive metadata before focused title and notes
edits assert that the original slide is unchanged. Production ownership proof
remains strict: zero required slide styles and missing ownership metadata are
rejected rather than admitted as fixture compatibility.

The host parity cut removes exactly `set_slide_title`,
`replace_slide_title`, `clear_slide_title`, `set_slide_body`,
`replace_slide_body`, `clear_slide_body`, `set_slide_notes`,
`replace_slide_notes`, and `clear_slide_notes`, together with their two
private storage-resolution helpers. Builder, add-slide, and duplication
coverage use the focused semantic transactions for their replacement behavior,
and the raw-index title/body example is replaced by the focused example. The
wider host editor remains because its creation, placeholder
visibility/layout, arbitrary text-box, generic text-storage, and other graph
mutations have not moved.

The root presentation facade also makes a separate intentional semantic
correction. Public `Slide::Keynote` now retains navigator `name` and visible
`title` as distinct fields, and its `text` uses Keynote's complete plain-text
projection rather than omitting body storage. `Slide::name()` therefore returns
the navigator name; callers that previously treated it as the canvas title use
the new `Slide::title()` accessor. Downstream constructors or exhaustive
destructures of the public struct variant must add the `name` field.

Migration is explicit rather than aliased:

- `set_slide_title(index, text)` and `set_slide_body(index, text)` become a
  mutable local `SlideTextEdit` from `edit_slide_title(selector)` or
  `edit_slide_body(selector)`, followed by `set(text)` and the consuming
  `commit()` call.
- `replace_slide_{title,body}(index, start..end, text)` becomes the matching
  focused edit with `TextSpan::from_utf16_indexes(start, end)` and `replace`;
  the indices remain UTF-16 code-unit positions but are now checked types.
- `clear_slide_{title,body}` becomes `clear` on the matching focused edit.
  The three notes operations use the corresponding `SlideNotesEdit` methods.
- A numeric caller can use `Position::new(index)`; an exact navigator name is
  preferred when it is the durable semantic selector. Missing, ambiguous,
  contradictory, or shared ownership returns a focused error instead of
  falling through to the generic storage editor.

Each commit returns a new immutable `Package`; it does not mutate the package
that began the edit. Sequential changes must therefore start the next edit
from `commit.package()` or from `commit.into_package()`. This chaining rule is
also why the removed mutable host methods are not source-compatible aliases.
Changed-output compatibility is semantic rather than byte-local to the text
storage: callers and differential tests must allow the declared selected
slide-node cache invalidation and root-preview deletion. Exact no-ops remain
byte-identical and do not invalidate caches.

The native source was the 500,058-byte checked-in `basic.key` fixture with
SHA-256
`3a3d07476b45b6e543bcfba75fe38a245434176dcb3565e34570b817708b9f42`.
Before the cache-invalidating release blocker was integrated, the focused
example produced `title-rust.key` with SHA-256
`9a093d1d99f533549038d14744af77262a51781d685558442e2526a7a66b502a`;
its title inverse restored the exact source hash. Sequential title and body
edits produced `title-body-rust.key` with SHA-256
`02d95162b6bede695093f4e7cb7d7aff3f7a9217b70b13b691b231d2d0626318`;
its body inverse restored the exact title-only hash. These hashes remain
semantic text-splice and inverse evidence, but they are not byte-golden outputs
for the current writer because they predate slide-node invalidation and root
preview deletion.

Computer Use opened the sequential Rust candidate in Apple Keynote without a
repair, recovery, conversion, or warning prompt. Keynote displayed the exact
title `Litchi title mutation 2026-08-09`, exact body `Litchi body mutation
東京😀`, and untouched date `2026-08-07`. Native Save As, close, and exact-path
reopen produced `title-body-native-resaved.key` with SHA-256
`e34749fae1e112caf6a6b960da26433f6d92a8366e12f83b59cbef6b03b0b563`.
Focused reread recovered both requested values, and same-value title and body
commits each reported unchanged, touched zero components, skipped full
reparse, and retained the exact native-resaved hash. This native run likewise
predates the cache-hardening change: it proves canvas text and native
open/save/reopen behavior for the splice, not the final cache-invalidating byte
shape. A changed title/body writer must invalidate rendered caches even when
semantic readback already shows the new string, because navigator and package
preview surfaces are independently persisted native artifacts.
The cache-hardened rerun used the same source and produced a title-only
candidate with SHA-256
`8cd82df9d83a6beea473efc2a7a5251f50fb0e9ad4198766a3d2b6eb6cf4bc32`
and the sequential title/body candidate with SHA-256
`f3b13cd5bd614d93493cc6780ff177e6a203d990d15b9d5c592687ef40a48263`.
Both inverses restored their exact respective inputs. The forward candidate
contained no root preview member and invalidated the selected slide-node cache.
Computer Use opened it in Apple Keynote without repair, recovery, conversion,
or warning; the navigator and canvas showed the exact title `Litchi native
Keynote fixture — real-app 🚀`, exact body `Buffa native UI verification —
東京😀`, and untouched date `2026-08-07`. Native Save As regenerated all three
root preview members and produced SHA-256
`cb3f9b05613505bb422942ca43e237a731454f58753ee65f26ae639187b96a6c`.
Close and recent-path reopen were warning-free and displayed the same values.
Focused same-value title and body commits on the native copy each reported
unchanged, touched zero components, and reproduced that native hash exactly.

The boundary-checker inventory contains 51 regressions, including the
exact-declaration negative policy for the nine retired public methods and two
retired private helpers; all 51 pass. The repository-wide policy command still
reports 14 unrelated, pre-existing `soapberry-zip`/`xml-minifier` edge
classification errors. The vertical changes no workspace membership or manifest edge, and
the current metadata/policy inventory is 64 workspace packages, 235 internal
dependency declarations, and 14 ordered migration debts.

This gate is intentionally narrower than complete Keynote editing. It neither
creates nor deletes title/body placeholders, edits arbitrary text boxes,
serializes patches durably, publishes files atomically, converts the complete
Keynote graph to Buffa, nor deletes `litchi-iwa`. `SlideTextError` retains
typed limit kinds and observed/maximum counts, but does not yet retain the
content-free semantic object path required by ADR 0005; adding a
format-owned `SlideTextPath` and threading it through semantic, wire, archive,
cache, and rewrite failures remains explicit follow-up debt.

## 2026-08-10 amendment: Numbers table-lock migration gate

The concrete Numbers package now has a semantic, selector-first transaction
for the effective lock state of an attached table. `Package::table_lock`
selects a sheet and table by checked position or exact name.
`edit_table_lock` returns one mutable semantic staging value with `set_state`,
`lock`, and `unlock`; commit returns an immutable package, exact-source
reversible `TableLockPatch`, and `TableLockDiagnostics`. `apply_table_lock`
accepts only a patch whose complete retained source artifact, source
fingerprint, selected semantic position, and before-state match. Native object
identifiers, component locations, protobuf messages, and wire values remain
private.

Selection first uses the rooted archive-free document, then resolves the same
sheet position through the native document and sheet drawable sequence. Each
candidate drawable must have at most one canonical type-6000 or legacy
type-6003 `TableInfo` message and never both. The selected payload's required
table-model reference remains the ownership identity; missing, duplicate,
ambiguous, zero, malformed, or noncanonical selected shapes fail before
publication.

The exact lazy boundary is intentionally smaller than the full table graph.
`litchi-iwa-protos::table_info_codec::decode_table_info` performs bounded
schema-directed raw preflight across the selected `TableInfo`, its required
drawable `super`, optional field-5 `locked` scalar, and required nonzero model
reference. It preserves `None` versus `Some(false)` and accepts only canonical
Boolean zero or one. The private Buffa `TableInfo` view forces both deferred
lazy branches: the required drawable `super` with its optional `locked` value
and the required table-model reference. The complete presence-preserving
Buffa snapshot must equal the handwritten preflight snapshot. Neither path
encodes the message or retains unknown fields. The accepted raw message, raw
object-header metadata, and physical package remain authoritative for
rewriting and preservation.

This three-message `TableInfoArchive`/`DrawableArchive`/model-reference
projection explicitly supersedes the 2026-08-08 two-message, opaque-super,
64 KiB seam. Buffa now forces both `super.locked` and `table_model`; the five
generated files measure 83,529 bytes under the current 84 KiB build cap.

An exact no-op returns a snapshot sharing the source allocation, reports
`changed = false`, zero touched components, and no full reparse, and leaves an
absent field absent or an explicit false explicit. A changed commit patches
only nested drawable field 5, replaces the selected raw message under the
retained core limits, rewrites exactly one Snappy/IWA component, and asks the
archive owner's bounded exact-name reassembler to produce a flat package. It
then reopens the complete candidate under the original `ReadOptions` and
requires the selected effective state to match. Diagnostics report
`changed = true`, one touched component, and a full reparse. Unselected fields
within the drawable and `TableInfo`, sibling messages, unselected object-header
metadata, other components, and other ZIP entries remain preservation-owned.
Competing rooted sheet ownership, contradictory selected-sheet or selected-
TableInfo reference metadata, noncanonical object-length prefixes in the
selected component, and merge/diff metadata on the selected owner fail closed
instead of being normalized. Detached or unrooted pseudo-sheet and view-state
dependent references are opaque preservation data, not competing owners.
Changed publication from a normalized legacy nested
`Index.zip` source is refused; reads and exact no-ops retain that original
artifact.

`TableLockPatch` is process-local and non-compact: it retains both complete
package artifacts rather than providing durable serialization. Applying a
changed patch does not rewrite or reassemble; after exact-source checks it
reopens the stored target under the applying package's retained options and
rechecks the after-state. Applying its inverse follows the same path and
restores the exact original package, including whether the lock field was
absent or explicitly encoded. Replayed, stale, tampered, inverse-on-source,
and cross-selector patches fail with `PatchConflict`. Selector absence,
duplicate-name package-ingress failure, unsupported/invalid source, bounded resource and allocation
failures, verification failure, and conflicts remain distinct format-owned
errors; a failed commit never changes the immutable source snapshot.

The Numbers-specific host read and mutation surface is removed. This includes
`NumbersEditor::table_lock_state`, `set_table_lock_state`, private
`table_lock_context`, `NumbersTableInfo.lock_state` and its field-population
branch inside `tables()`, both model-specific shared helpers
`table_lock_state_for_model`/`set_table_lock_state_for_model`, and their
Numbers-only model-ID matching branch. This is a breaking replacement with
immutable semantic selection, not an alias: Numbers readback now also uses
`litchi_numbers::Package::table_lock`. The boundary checker ratchets five
exact functions in their collision-safe scopes—three names under the host
Numbers tree and both model-specific helpers in the shared codec—and
separately rejects a public `NumbersTableInfo.lock_state` field. The
field-population and matching-branch removals complete the broader compiled retirement
inventory. Pages and Keynote still use the generic shared
table-lock getter/setter and raw codec. Durable patch serialization,
library-owned atomic filesystem publication, and complete host deletion remain
outside this vertical.

The focused `edit_table_lock` example accepts semantic name/index selectors,
publishes to a sibling temporary with no-clobber persistence, and can emit an
optional inverse artifact only after exact byte restoration succeeds. The
mixed iWork table-lock example now uses the host only to construct its scratch
Numbers table, then uses the focused package for both mutation and readback;
its Pages and Keynote branches remain on their existing host APIs. Those
examples define the intended publication and migration workflow but are not,
by themselves, native-application evidence.

The verification inventory is two `litchi-numbers` lock-state unit tests, nine
strict TableInfo codec tests, and 15 focused package transaction tests. The
cases cover semantic selector resolution; absent, explicit-false, and true
wire states; exact no-op allocation sharing; unknown-field, sibling-message,
unknown object-header metadata, component, and ZIP preservation; one-component
changed publication; exact inverse; stale/tampered/replayed/cross-selector conflicts;
legacy read/no-op and changed refusal; typed output limits and failure
atomicity; contradictory selected-owner metadata or competing rooted sheet
ownership; detached/unrooted reference preservation; noncanonical object framing;
merge/diff refusal; redacted diagnostics; and deterministic concurrent reads.
The focused suite passed 15/15, including the rooted `FormBasedSheet` nested
drawable field path `[1, 2]`, checked-in native fixture semantics, and exact
inverse behavior. It also covers changed flat legacy type-6003 TableInfo
publication and exact partial-sink write accounting.
The bounded `numbers_table_lock` fuzz target compiles, and all 57 boundary
policy regressions pass. The full policy command still reports the 14
pre-existing soapberry-zip/xml-minifier annotations. A Numbers-only fuzz
package and a sustained sanitizer campaign remain open.

The current focused-writer native gate is complete. The source artifact has
SHA-256
`f225d5b1cd59e9da454f91a96fe8f81154bc31037c10029230e75d49b45fb693`;
the Rust locked artifact has SHA-256
`eb2e29c97c415c1b61ed1f8fe766e7211ed386c825c32dec056b72c9398d3e09`;
and applying the inverse restored the source bytes and source hash exactly.
Apple Numbers 14.4 (7043.0.93) opened the Rust artifact without a warning,
showed `Table 1` as locked with disabled cells, and retained
`B2 = Litchi native Numbers fixture`
and `B3 = 42`. Native Save As, close, and exact-path reopen completed without a
warning. The native-resaved artifact has SHA-256
`8aa87a3afcb145b66c5c6f4e10645cd1cf658f4b65f0976612ac6d62d4652995`;
focused reread returned `Locked`, and an equal-state lock transaction was an
exact no-op with that same hash.

This does not close the complete performance and publication contracts.
Physical, semantic, and wire stages have finite limits, typed allocation
errors, and deliberate early drops of large rewrite temporaries, but the
transaction has no measured or caller-selected aggregate peak-memory model
covering its retained source/target patch artifacts, selected-component
buffers, reassembly output, and fully reopened candidate. Complete repeated
traversal, fingerprint, copy, and reopen work is not charged to one aggregate
transaction-work ceiling, and a transitive proof that every allocation on the
complete path is fallible remains open. `Package::write_to` emits the exact
artifact and reports the accepted byte count on sink failure; the example
demonstrates sibling-temporary no-clobber
publication, but the library does not yet own an atomic, durable filesystem
save/replacement contract. Durable patch serialization, deeper fuzzing, and
final host deletion also remain outside this gate. The process-local patch has
no versioned semantic operation envelope, read/write sets, composition,
three-way merge, or bounded history.
Resource and allocation errors also lack the selected semantic table path,
and `Package::source_bytes` remains ordinary public surface rather than an
explicit advanced/raw boundary.
The flattened `TableLock*` transaction names likewise remain migration debt
against the focused-module short-name rule.
The archive-free `Table` snapshot does not yet carry lock state, and remaining
host table/cell mutations do not enforce that state by default; read-model
convergence and protection enforcement remain host-migration debt. The private
Numbers locator also remains specialized instead of converging on the neutral
IWA index owner.

## 2026-08-10 amendment: Pages page-layout cutover and verification contract

Pages page geometry now follows the focused immutable package flow. The public
semantic value remains the presence-preserving `page_layout::Layout`, while
`Package::{page_layout, edit_page_layout, apply_page_layout}` and the
format-owned edit, commit, reversible patch, diagnostics, error, and limit
types own the artifact transaction. Native object IDs, component names,
message types, generated messages, and layout/cache wire fields remain private.

The exact read call graph is `Package::page_layout` to the unique object 1 and
unique type-10000 `TP.DocumentArchive` payload, then a bounded canonical raw
preflight of required opaque field 15 plus layout fields 30 through 39 and 42,
then forced access to every projected scalar on the private Buffa lazy view,
and finally equality with the archive-free `Layout`. Raw preflight rejects
duplicates, wrong wire types, noncanonical keys or varints, invalid Booleans,
and invalid semantic geometry before the generated view can authorize the
result. The existing document-body projection, rather than a second competing
schema, owns the layout scalars. Its five generated files measure 122,114
bytes under a 124 KiB cap, contain no repeated view or production encoding
path, and deliberately leave the required `TP.DocumentArchive.super` payload
opaque. Buffa is a checked borrowed read seam, never the preservation or
encoding representation.

A changed commit raw-splices the eleven presence-preserving scalar fields in
the selected document payload. Cache ownership is proved through the rooted
raw call graph, not by claiming every package-wide view-state candidate. The
adapter rejects deprecated `TP.DocumentArchive` fields 11 and 12, parses its
required `super` field 15, follows `TSA.DocumentArchive.view_state` field 5 to
a unique referenced type-210 shared view-state object, follows that payload's
field 1 to a unique referenced type-10147 `TP.ViewStateRootArchive`, and then
strictly decodes the root's optional layout-state field 1 and UI-state field 2.
Every followed reference must be nonzero and local. The document-to-bridge
edge must occur exactly once in aggregate metadata and, when field metadata is
present, exactly once at path `[15, 5]`; the bridge-to-root edge has the same
contract at path `[1]`.
If layout state exists, the transaction removes only field 1 and its exactly
once-declared aggregate reference metadata plus the optional unique field
declaration at path `[1]`. It does not delete or decode the referenced
layout-state object, rewrite the type-210 bridge, or claim detached/unrooted
type-10147 objects; those remain opaque and exact. It also preserves field 2,
unknown view-state content, and unrelated reference metadata. Missing,
duplicate, or contradictory objects on the rooted chain, a shared layout/UI
identifier, selected merge/diff metadata, and noncanonical object-length
prefixes fail closed instead of being normalized.

Document and rooted view-state edits publish in one candidate. If both objects
share an IWA component, diagnostics report one touched component; if a separate
component carries the layout-state edge, diagnostics report two. The same
bounded deletion-aware archive operation removes root `preview.jpg`,
`preview-micro.jpg`, and `preview-web.jpg`, with deletions counted separately
from IWA components. This is the explicit preservation exception required to
prevent Pages from retaining layout-derived view and rendered preview state;
all other retained ZIP records and unselected IWA content remain exact.
Publication completes only after reopening the whole candidate under the
retained limits and verifying the new layout, absent layout-state edge, absent
root previews, stable package statistics, and unchanged section names, types,
headings, paragraphs, text storages, and page counts.
Bounded canonical unknown protobuf groups remain readable and exact no-ops
retain them, but changed page-layout publication currently refuses a
group-bearing document payload because the scalar splicer does not yet own a
group-aware rewrite rule.

An equal-layout commit is a byte-exact no-op: it retains optional-field
presence, view-state and preview bytes, shares the source allocation, reports
zero touched components and deleted previews, and performs neither reassembly
nor candidate reopen. A no-op patch additionally requires exact artifact
identity. A changed patch is bound to complete source and target artifacts;
application checks exact source bytes, semantic layout, layout-state identity,
and preview count, then reopens and verifies the stored target rather than
reassembling it. Its inverse swaps those artifacts and restores the original
document, cache edge, previews, and bytes exactly. Legacy nested `Index.zip`
packages remain readable and support exact no-ops, but changed publication is
rejected as `UnsupportedSource`.

Migration is intentionally breaking rather than aliased:

- `PagesEditor::page_layout()` becomes `Package::page_layout()`.
- `PagesEditor::set_page_layout(layout)` becomes
  `let mut edit = package.edit_page_layout()?; edit.set_layout(layout)?; let
  commit = edit.commit()?;`.
- The next immutable edit must begin from `commit.package()` or
  `commit.into_package()`; the package that began the edit is never mutated.
- Changed-output comparisons must allow the declared layout-state-edge and
  root-preview removal. Exact no-ops remain byte-identical.

The host methods, private `editor::page_layout` module/source, their duplicate
host tests, and `litchi-iwa/examples/edit_pages_layout.rs` are removed. The
focused `litchi-pages` example validates width, height, and orientation, writes
through a synced sibling temporary file with no clobber, and can emit an exact
inverse. Boundary checks forbid the retired host method declarations and
module/source from returning and reject physical IWA/protobuf vocabulary in
the focused public facade.
No workspace membership, internal dependency declaration, or ordered debt
changes: the checked inventory remains 64 packages, 235 internal declarations,
and 14 ordered migration debts.

The final deterministic gate passed all 92 `litchi-pages` tests and doctests,
including the focused 10/10 transaction suite, and the private page-layout
codec passed 6/6. Coverage includes one- and two-component invalidation,
rooted metadata proof at every hop, detached-cache preservation, preview
deletion, raw unknown/header locality, bounded canonical unknown groups,
malformed scalar and reference framing, deprecated cache paths, selected
merge/diff refusal, exact no-op and inverse behavior, stale/tampered patch
conflicts, legacy changed refusal, typed limits, failure atomicity, concurrent
reads, and public `Send + Sync` assertions. The package check and focused
no-dependency library Clippy with warnings denied pass. Full all-target Clippy
remains blocked upstream by 88 existing `litchi-core` SIMD lint errors. All 63
boundary-policy unit tests pass and the live Pages retirement/facade audits are
clean.

The `pages_page_layout` fuzz binary compiles. Its bounded harness sends
arbitrary bytes through checked ingress and interprets the same input as layout
commands over the checked-in native fixture, covering reads, no-op/change/clear,
apply, conflicts, inversion, exact restoration, limits, and redacted failures.
Thirty-two generated smoke inputs and a fixed changed-layout corpus completed.
The attempted sanitizer-backed 1,000-run `cargo fuzz` campaign did not start:
the active stable toolchain rejects `-Zsanitizer=address` and no nightly
toolchain is installed. This is compile and non-sanitized smoke evidence, not a
sustained sanitizer campaign.

The checked-in `basic.pages` transaction changed the native fixture to 792 by
612 point landscape, reported `changed=true`, touched two IWA components,
deleted all three root previews, retained the semantic body text, and left no
root preview entry in the Rust candidate. The source and current Rust-candidate
SHA-256 values are
`21107bc9323fba6f1589152454c0b0b0cc8e239313c6a369bc4a891116601b42`
and `79e00545ef6e2e30e366e3160b7d9126bf06cffac5fbbd5551e3d3789cc298e4`.
Applying the inverse restored the exact source hash. Reapplying the same layout
to the Rust candidate reported unchanged, zero touched components, zero
deleted previews, and preserved the candidate hash exactly.

Apple Pages 14.4 (7043.0.93) opened the current Rust candidate without warning,
repair, recovery, or conversion. The Document inspector showed Any Printer,
US Letter, Landscape selected, 11.00 by 8.50 inches, and Document Body checked.
The accessibility text contained exactly `Litchi native Pages fixture`,
`Buffa lazy-view migration verification`, and `2026-08-07`. Native Save As,
close, and exact-file reopen were warning-free and reconfirmed the same layout,
document kind, and body. The native-resaved SHA-256 is
`8228e7518bb080bd8e5ec134d0abc7484c8825ad3cde3d16cabf76c5dbd8ef82`,
and native Save As regenerated root `preview.jpg`, `preview-micro.jpg`, and
`preview-web.jpg`. A focused equal-layout transaction over that native artifact
reported unchanged, zero touched components, and zero deleted previews and
produced a byte-identical output with the same hash.

This cutover does not claim ownership of the detached opaque layout-state
object, other document settings, durable patch serialization, semantic patch
composition or merge, library-owned atomic durable filesystem replacement,
or a whole-Pages Buffa conversion. The example's publication workflow is
evidence, not the missing library contract. Aggregate transaction peak memory
and total work are not yet bounded across retained before/after artifacts,
component recompression, hashing, ZIP reassembly, and full candidate reopen;
nor is every transitive allocation proven fallible. In particular, the shared
archive encoder still deep-clones retained `ArchiveInfo` metadata through an
infallible path. Exact source bytes remain an ordinary `Package` API, and the
flattened `PageLayout*` names remain focused-module naming debt.

## 2026-08-10 amendment: combined Pages document-settings cutover

Pages document visibility/layout options and footnote formatting now migrate as
one immutable transaction because the native format stores both in a single
settings owner. The archive-free `document_settings::Settings` combines
`document_options::Options` with `footnote::Settings`; its public focused module
uses the canonical short `Edit`, `Commit`, `Patch`, `Diagnostics`, `Error`, and
`LimitKind` names. `Package::document_settings`, `edit_document_settings`, and
`apply_document_settings` are the only new package entry points. Their focused
method/type signatures and errors expose no raw ID, member name, message type,
generated type, field number, source byte slice, or retained patch artifact.

The exact rooted read graph is:

1. resolve the unique object 1 and unique type-10000 `TP.DocumentArchive`;
2. strictly decode required local `TP.DocumentArchive.settings` field 7, prove
   one aggregate reference and optional unique path-`[7]` field metadata;
3. resolve that nonzero identifier to exactly one object and exactly one
   type-10012 `TP.SettingsArchive`; and
4. decode fields 1, 2, 3, 9, 10, and 30 through 34 under one aggregate byte,
   field, work, and nesting budget.

The root preflight also checks required document `super` and singular body and
section edges without claiming their payloads. The settings field mapping is
body 1, headers 2, footers 3, hyphenation 9, `use_ligatures` 10, footnote kind
30, format 31, numbering 32, gap 33, and facing pages 34. Raw preflight rejects
duplicates, wrong wire types, noncanonical varints, invalid Booleans,
non-sign-extended `int32`, zero or external references, ambiguous owners,
contradictory reference metadata, selected merge/diff records, and
noncanonical selected component framing. It then forces both the document
settings-reference lazy view and all ten `PagesSettingsArchive` scalar fields;
the complete Buffa snapshot must equal preflight and the archive-free semantic
value. Valid newer enum integers map to canonical `Unknown` variants, while a
caller cannot construct an `Unknown` wrapper for a known value.

This supersedes the shared projection's page-layout-only 122,114-byte/124-KiB
record. The five generated body/layout/settings files total 174,682 bytes under
a 176-KiB cap and have deterministic aggregate SHA-256
`7618a60db84b87e28eea67a8acd85ce8eb19513cf4cee7654c1c4e78f405f824`.
The build rejects any repeated-view or production-encoding surface. Document
`super`, repeated section tables, unknown fields, raw headers, and all
preservation state remain outside Buffa and caller-owned.

A changed commit raw-splices only the ten selected settings scalars with exact
presence. It rewrites the settings-owner component and reuses the page-layout
vertical's rooted document-super/view-state/type-210/type-10147 cache proof.
The transaction removes only the rooted layout-state edge and its proven
metadata, leaves the opaque cache object and detached/unrooted candidates
untouched, and deletes root `preview.jpg`, `preview-micro.jpg`, and
`preview-web.jpg`. Settings and cache roots can share one component or occupy
two, while ZIP deletions are counted separately. Every accepted unselected
object/message/header, unrelated metadata reference, component, and retained
ZIP record remains exact.

Canonical unknown scalar fields remain exact through changed publication.
Bounded canonical protobuf groups are readable and retained by exact no-ops,
but a group-bearing settings payload is rejected for a changed splice because
no group-aware rewriter is owned. Noncanonical encodings are rejected even for
reads. Changed publication fully reopens under retained limits and verifies the
combined settings, absent cache edge and previews, stable statistics, and
unchanged section names, types, headings, paragraphs, storages, and page counts.

No-op ordering is deliberate: once the rooted settings value was read to begin
the edit, an equal `Settings` commit shares the exact source allocation before
cache traversal, preserves field presence/caches/previews, and reports zero
components, zero deletions, and no full reparse. A no-op patch needs only exact
artifact identity. Changed patches retain exact source and target artifacts;
application checks exact bytes and semantic/cache/preview preconditions, then
reopens the stored target rather than reassembling it. Replay, source tamper,
an inverse on the source, and competing patches conflict. A valid inverse
restores the exact original artifact.

Migration is intentionally combined and immutable:

- `PagesEditor::document_options()` becomes
  `Package::document_settings()?.options()`.
- `PagesEditor::footnote_settings()` becomes
  `Package::document_settings()?.footnotes()`.
- Either old setter becomes a read of the composite `Settings`, replacement of
  only `options` or `footnotes`, then
  `package.edit_document_settings()?.set(settings).commit()?`.
- A later transaction begins from `commit.package()` or
  `commit.into_package()`; the source package is never mutated.
- Changed-output comparisons permit the declared cache-edge and root-preview
  removal. Exact no-ops remain byte-identical.

The four host getter/setter methods, `document_options.rs`, its nested
`wire.rs`, `footnote_settings.rs`, two host examples, and duplicate host tests
are removed rather than shimmed. One focused example validates all ten semantic
choices, publishes via a synced sibling temporary with no clobber, and can emit
an exact inverse. Legacy nested `Index.zip` sources retain reads and exact
no-ops, but a changed edit returns `UnsupportedSource`; the old host's changed
normalization is deliberately not compatibility behavior. Boundary policy
ratchets all four methods, both module declarations, all three retired files,
and native/IWA/protobuf/source-byte leakage from the semantic and transaction
facades.

The final deterministic gate passes 108/108 Pages tests and doctests, including
14/14 combined transaction tests, 4/4 strict codec tests, and 6/6 root-facade
tests. Package check, strict no-dependency Clippy with warnings denied, strict
documentation, and 70/70 boundary-policy regressions pass. The live boundary
command remains red only for 14 unrelated pre-existing soapberry-zip and
xml-minifier declarations. The focused fuzz target compiles and its explicit
no-op and changed smoke inputs pass. The sanitizer-backed campaign cannot start
under the installed stable toolchain because cargo-fuzz requires
`-Zsanitizer=address` and nightly is unavailable.

The native gate used a fresh Apple Pages-authored file containing a real
footnote. Source and Rust-candidate SHA-256 values are
`9da01e2805459e05450551827140069eefe8049aeeacc7625d3c62d7e00ffeab` and
`3d052e7f1ec86e57ea0553e46f628de1d9fa5bdda615ded9410fca29c93f0995`.
The focused edit reported `changed=true`, two touched components, and three
deleted root previews; its inverse restored the source exactly. Apple Pages
14.4 (7043.0.93) opened the Rust candidate with no warning, repair, recovery,
or conversion. Its UI and focused readback showed body/header/footer enabled,
facing pages enabled, hyphenation and ligatures disabled, and Footnotes with
Roman markers, restart each page, and an 18-point gap. All three body markers
and the note text remained exact.

Native Save As, close, and exact-file reopen were warning-free and reconfirmed
the same settings and text. Save As regenerated all three root previews and
produced SHA-256
`803167e2479c459f9a33c8ecfc4d713f596fdc5d5d337090ab3c90e467a0cba6`.
A focused same-settings commit on that artifact reported unchanged, zero
components, and zero deletions and produced the same hash exactly; applying its
no-op inverse did likewise.

This gate does not close aggregate transaction peak-memory or total-work
accounting, the shared encoder's infallible retained-`ArchiveInfo` clone, the
complete fallible-allocation audit, group-aware changed splicing, exact
streaming and partial-output accounting, or a library-owned atomic durable
filesystem replacement. The process-local patch still lacks a stable versioned
serialization, semantic operation/read-write sets, composition, three-way
merge, and bounded history. Exact source bytes remain ordinary `Package`
surface. The opaque cache object and remaining Pages settings/render state are
not claimed. No manifest edge or ordered debt changes; the current inventory
remains 64 packages, 235 internal declarations, and 14 ordered debts.

## 2026-08-10 amendment: hardened Keynote show-settings cutover

This amendment supersedes the 2026-08-08 show-settings compatibility and
verification record. Callers use the archive-free
`show::{Settings, Edit, Patch, Commit, Diagnostics, Error, LimitKind}` family
and `Package::{show_settings, edit_show_settings, apply_show_settings}`. The
edit replacement is consuming:

```rust,ignore
let before = package.show_settings()?;
let mut after = before;
after.set_loop_presentation(Some(true));
let commit = package.edit_show_settings()?.set(after).commit()?;
commit.package().write_to(&mut output)?;
let restored = commit
    .package()
    .apply_show_settings(&commit.patch().inverse())?;
```

The new focused method/type signatures contain no native ID, ZIP/IWA member,
generated type, wire field, raw byte slice, or retained artifact accessor.
`Package::write_to` is the supported exact-output seam and reports the sink
offset reached on failure without allocating another package-sized buffer; it
does not flush or publish a filesystem path.

The exact rooted admission graph is:

1. select the unique component whose basename is `Document.iwa`;
2. select object 1 and exactly one type-1 `KN.DocumentArchive` message;
3. strictly decode required local `KN.DocumentArchive.show` field 2 and force
   the complete Buffa reference view, including legacy type/external presence;
4. require a nonzero show identifier exactly once in aggregate reference
   metadata, with optional unique matching field metadata only at path `[2]`;
5. resolve that identifier in exactly one component to one object with exactly
   one type-2 `KN.ShowArchive` message; and
6. strictly validate the complete known Show/SlideTree envelope, then force and
   cross-check the Buffa size, references, and eight optional scalar settings.

An explicitly external root reference is invalid. A zero reference maps to
`Settings::default()` for reads and exact no-ops, but changed publication is
`UnsupportedSource` because this focused transaction does not allocate an
object or register a component. Detached and unrelated native objects are not
claimed as owners.

The root projection's five generated files measure 58,630 bytes under 60 KiB
and have deterministic aggregate SHA-256
`7918aad2578cf3bd07eb0be36f2e31d11f93391584308c1e4adc1fd86ed065fd`.
The show projection's five files measure 138,661 bytes under 140 KiB and have
aggregate SHA-256
`747fe9f99dc5bb1855aae1bfcb16065a5fe6305bdbf8730a21ef24bb75e915ee`.
The repeated SlideTree is strictly routed by hand rather than retained in a
generated repeated view. Build ratchets prohibit repeated views and production
encoding in both projections. Buffa is a forced semantic cross-check, never
the preservation or mutation owner; canonical raw field records remain that
authority.

Changed publication adds guards that the read path deliberately does not use:
canonical IWA object-length framing for selected components and no
`should_merge`, base-message index, diff/merge version, diff field path,
fields-to-remove, or diff read version on either selected message. A valid
change raw-splices Show size field 4 and only the changed optional scalars at
fields 6, 8, 9, 10, 11, 15, 16, and 18, preserving optional presence and every
accepted unknown record. Exactly one Show-owner IWA component is reassembled.
The complete candidate is reopened with retained limits and checked for exact
ownership, requested settings, package/member structure, metadata closure,
cache policy, and unchanged semantic content.

Rendering invalidation is narrow:

- A size or slide-number-visibility change deletes every existing root
  `preview.jpg`, `preview-micro.jpg`, and `preview-web.jpg`; diagnostics report
  the actual zero-to-three deletion count separately from the one rewritten
  component.
- A playback-only change preserves all root previews exactly.
- Both cases preserve every slide component and slide-node thumbnail/playback
  cache exactly; show-level rendering does not authorize slide-cache pruning.
- An exact semantic no-op preserves every byte and cache, shares the source,
  reports zero components/deletions, and skips cache traversal, reassembly, and
  reopen.

Changed patch application checks the exact complete source artifact, settings,
ownership, and preview preconditions, then reopens the exact stored target
instead of reconstructing it. Replay, source tamper, wrong package, and inverse
on the source conflict. Applying the inverse to the target restores the exact
complete source. Output comparisons therefore permit only the Show component
rewrite and, for rendering changes, root-preview removal.

The compatibility table is now:

| Source | Read | Exact no-op | Changed edit |
| --- | --- | --- | --- |
| ordinary exact package | supported | byte-exact | supported after strict changed guards |
| null rooted show | default settings | byte-exact | `show::Error::UnsupportedSource` |
| legacy nested `Index.zip` | supported | byte-exact | `show::Error::UnsupportedSource` |

The last row is an intentional behavior break under Preserve policy. The
deleted host normalized a changed legacy package; the focused owner refuses to
silently change its physical provenance.

Migration removes `KeynoteEditor::show_settings`,
`KeynoteEditor::set_show_settings`, the `show_settings` editor module and
`keynote/editor/show_settings.rs`, `examples/edit_keynote_show.rs`, and the
direct editor mutation/compatibility tests. The focused
`litchi-keynote/examples/edit_show_settings.rs` now demonstrates semantic
staging, immutable chaining, exact inverse verification, distinct paths,
no-clobber temporary output, and `Package::write_to`.

The boundary is intentionally precise: `KeynoteDocument::show` remains a
read-only host path that decodes a Prost `KN.ShowArchive`. Therefore this gate
retires direct `KeynoteEditor` show-settings mutation, not every host read or
all native Show ownership. Other creation, slide, media, soundtrack,
transition, and graph code remains outside the focused vertical.

Current deterministic evidence passes 19/19 focused transaction tests,
106/106 full codec tests, 49/49 focused Keynote codec tests, Keynote all-target
checking, the `litchi-iwa` library check, umbrella Keynote facade compilation,
strict documentation, and 80/80 boundary regressions. The two focused live
audits for retired host surface and focused public leakage are empty. The
general repository boundary run still reports 14 unrelated pre-existing
diagnostics: 12 for six `soapberry-zip` dev-only edges and two for
`xml-minifier`. The fuzz target passes `cargo check`; the stable-built target
completed 32 bounded cases with expected missing-sanitizer-symbol warnings.
The cargo-fuzz sanitizer run could not start because it requires unavailable
nightly rather than the installed stable toolchain.

Apple Keynote 14.4 (7043.0.93) exercised two rendering-invalidating edits from
the same source, SHA-256
`f3adcde9315b6df580805bcb63c995cc1e1ef569a4befa06a102485e13c883b2`.
For the slide-number case, pristine Rust output SHA-256 was
`6d28d461c1203f00384fe6a758df1f903c7555b90ff02d2dc32d856aa9056c13`;
native Save As, close, and exact-path reopen yielded
`031a701040ed1ea9a5111fe3e298bcddcf33d498891f827b703d01328ba17224`.
For the size case, pristine 1280-by-720 Rust output was
`67e9ff0557683af105dfe57f999acabcde23f121f7aebb06102c93e03121c027`
and native resave was
`a3a2f6e072db4bd952f2c02e528f25c3656dba5810fbff75e93b5a699aac0eda`.
The public inverse of each pristine Rust candidate restored the exact source.

Both Rust candidates opened without warning, repair, recovery, or conversion
and auto-played. Inspectors showed Self-Playing, Loop and Play on Open enabled,
five-second transition and two-second build delays, with Widescreen
1920-by-1080 for the slide-number case and Custom 1280-by-720 for the size
case. Save As, close, and exact-path reopen preserved those settings and
auto-played. Rust removed all three root previews in both candidates; Keynote
regenerated them on native resave. Every one of the four `Index/Slide*.iwa`
hashes was unchanged between each pristine Rust candidate and its native
resave, providing native evidence that the conservative root-preview policy
does not disturb slide components or their caches.

One native normalization is explicit: Keynote resaved
`slide_numbers_visible = Some(true)` as absent/`None`. Restaging absence on the
native artifact is an exact no-op at the `031a7010...` hash, whereas restaging
true is a changed transaction. The native-resaved size artifact accepts a
same-settings no-op and its inverse byte-exactly at `a3a2f6e0...`. Therefore
the gate proves safe admission, slide-cache preservation, and conservative
preview invalidation; it does not claim native persistence of the
slide-number scalar.

Remaining exit work includes the host Prost `KeynoteDocument::show` reader and
other generated Show consumers; aggregate transaction peak-memory/total-work
accounting; a complete transitive fallible-allocation audit; canonical
group-aware changed splicing; versioned deterministic patch serialization,
semantic operations and read/write sets, composition, merge, and bounded
history; and a library-owned atomic durable filesystem replacement.
`write_to` closes raw-source exposure for output but does not flush, sync,
rename, or make a destination durable. A full sanitizer-backed fuzz campaign
remains explicit verification work.

## 2026-08-10 amendment: Numbers sheet/table names cutover

Numbers names now migrate as one immutable, final-state batch. The public
surface is `names::{Edit, Patch, Commit, Diagnostics, Error, LimitKind}` plus
semantic `Path` and `InvalidReason`, reached through infallible
`Package::edit_names` and exact `Package::apply_names`. `edit_names` is `O(1)`
and allocates nothing; each consuming `rename_sheet` or `rename_table` resolves
its selectors against the same immutable base, so later stages never depend on
an earlier staged spelling. The final batch permits swaps and collision-away
renames but rejects a repeated target, duplicate sheet name, or duplicate table
name within one sheet atomically.

No public method/type signature contains a native ID, component/member name,
archive/IWA/generated/Prost/Buffa/wire type, or raw source slice. The root
facade is `litchi::numbers::names`; flat aliases and globs are ratcheted out.
`Package::source_bytes` is crate-private. Exact output uses
`Package::write_to`, including its partial-sink byte-offset diagnostics; the
caller owns flush, sync, and durable filesystem publication.

Changed publication proves this rooted graph:

1. `Index/Document.iwa`, object 1, contains one selected TN document whose
   repeated field-1 references exactly match the rooted sheet sequence;
2. each reference is nonzero/local, declared exactly once in aggregate
   metadata and optionally once at matching path `[1]`, and resolves to exactly
   one `TN.SheetArchive` or `TN.FormBasedSheetArchive` message;
3. ordinary field 1 or forced form `super` field 1 owns the sheet name;
4. the selected sheet's drawable reference at `[2]`, or form path `[1, 2]`,
   resolves to one canonical type-6000 or legacy type-6003 TableInfo;
5. TableInfo's required local field-2 reference, with matching exact metadata,
   resolves to one canonical type-6001 or accepted legacy type-6000
   TableModel; and
6. TableModel required field 1 supplies identity while required field 8 owns
   the selected display name.

The changed plan also proves every selected model has exactly one rooted
TableInfo owner among all rooted sheet drawables. Competing rooted owners,
ambiguous canonical/legacy messages, external/zero/dangling references,
metadata contradictions, selected merge/diff state, noncanonical framing, or
semantic/native name disagreement fail closed. Detached/unselected native
objects remain opaque and are preserved.

Strict raw preflight precedes and is cross-checked with private Buffa lazy
projection for sheet, nested FormBasedSheet `super`, and TableModel identity
plus display name. The values are borrowed rather than allocated. Generated
provenance is exactly five files/82,641 bytes with aggregate SHA-256
`944b7637fd6bf0eb895174b1e9229aa9eb9c393e05c666a86dd2843792eefe3e`;
raw records, not Buffa, remain the mutation and unknown-field authority.

Changed-only safety refuses:

- a table rename whose selected TableInfo lock state is `Locked`;
- any table rename while any rooted table model has a pivot owner, because the
  vertical cannot update pivot naming dependencies; and
- any changed name when a rooted calculation-engine formula owner has nonempty
  volatile sheet/table-name cells.

A sheet-only rename remains supported when a table is locked. The conservative
pivot traversal is native Θ(T²), but the transaction overcharges and checks a
`WireWork` ceiling before any changed-only native scan. All touched operations
are sorted by component; each component is parsed and rewritten once, so one
batch publishes all names or none. Full reopen checks final semantic names and
exact locality.

Changed publication deletes each existing root `preview.jpg`,
`preview-micro.jpg`, and `preview-web.jpg` entry, reporting zero to three
deletions separately from touched components. It deliberately preserves
`Index/ViewState.iwa` and every unrelated package record, object, message,
metadata field, and unknown byte. A semantic no-op shares the original source,
keeps previews and ViewState, reports zero work, and bypasses changed-only
framing/cache/protection/dependency checks, reassembly, and reopen.

The patch privately retains exact source/target artifacts and the resolved
operation plan. Apply rejects stale/replayed/tampered/cross-artifact state,
then reopens the stored target and repeats semantic/locality checks; inverse is
`O(1)` to construct and restores the complete exact source, including previews.
It remains a process-local patch, not a durable serialized operation log.

Compatibility is explicit:

| Source/model | Read/no-op | Changed batch |
| --- | --- | --- |
| canonical Sheet/FormBasedSheet and canonical rooted TableInfo/TableModel | supported, exact | supported |
| accepted legacy type-6003 TableInfo/type-6000 TableModel | supported, exact | supported when unambiguous |
| nested legacy physical package | supported, exact | `names::Error::UnsupportedSource` |

The host migration map is:

- `NumbersEditor::rename_sheet(native_id, name)` becomes
  `package.edit_names().rename_sheet(semantic_selector, name)?.commit()?`;
- `NumbersEditor::rename_table(native_id, name)` becomes the corresponding
  semantic sheet/table selector stage; and
- multiple names should be staged on one edit and committed once, then later
  work must begin from `commit.package()` or `commit.into_package()`.

The two host methods, their direct mutation/compatibility tests, and
`litchi-iwa/examples/rename_numbers_items.rs` are deleted rather than shimmed.
The focused `edit_names` example provides combined selection, inverse checking,
`write_to`, and synced sibling-temporary/no-clobber output. The private
`rename_attached_table_in_package` helper remains for Numbers sheet
duplication, while its `rename_table_in_package` wrapper remains because Pages
and Keynote attached-table workflows still consume it; no public Numbers
editor mutation survives.

Final deterministic evidence passes 10/10 focused integration tests, 105/105
Numbers library tests, the 1/1 umbrella facade test with `--features numbers`,
89/89 boundary regressions, both live Numbers names/host audits,
`litchi-numbers --all-targets` checking, `litchi-iwa --lib` checking, and
strict rustdoc. Host `litchi-iwa --all-targets` is not claimed because
unrelated examples remain red. The stable fuzz target builds and its
control-flow smoke ran eight bounded cases with expected
missing-sanitizer-symbol warnings. This is not sanitizer execution; an ASan
campaign remains open.

The native writer gate used Apple Numbers 14.4 (7043.0.93). The ordinary,
unlocked source SHA-256 was
`f225d5b1cd59e9da454f91a96fe8f81154bc31037c10029230e75d49b45fb693`;
the Rust Unicode candidate was
`22f8bc21223317318ec23ec764b8998af77a2c7800c68cbe88351abdb26b6e56`,
and public inverse application restored the exact source. Numbers opened it
without warning/repair/recovery/conversion, showed sheet `Líneas 你好 🧪`, table
`表 Café №42`, the exact B2 text marker, and B3 numeric value 42. The table was
selectable/editable and the rename succeeded. Save As, close, and exact-path
reopen preserved the names/data and produced SHA-256
`e1803b0568454a345f7962c5b4c72e8cb3d78adb2c87d5db1e6c58288a9413c4`;
Numbers regenerated all three root previews. Focused equal restaging, no-op
apply, and inverse over the native artifact were all byte-exact at that hash.

A separate native protection oracle at SHA-256
`eb2e29c97c415c1b61ed1f8fe766e7211ed386c825c32dec056b72c9398d3e09`
reported `Locked` and `Locked items cannot be edited` in Numbers accessibility
state. Its cells were disabled, Unlock was enabled, and invoking Edit table
title produced no name change. That independently confirms the protection
state for the focused rule: table rename refuses, sheet-only rename remains
admissible.

This vertical changes no manifest edge. Ordered debt 015
(`litchi-iwa -> litchi-numbers`) therefore remains, and the inventory stays 64
packages, 235 internal dependency declarations, and 14 ordered debts.
Remaining debt includes the preflight-bounded native Θ(T²) pivot traversal, aggregate
transaction peak-memory/total-work accounting, complete fallible-allocation
proof, stable versioned patch serialization with semantic operations,
read/write sets, composition, merge, and bounded history, and library-owned
atomic durable filesystem replacement. Patches retain both complete artifacts
in memory; `write_to` is streaming, not durable publication.

## 2026-08-10 amendment: Keynote transition editor cutover

Keynote transition callers now use the archive-free semantic values and
canonical `transition::{Edit, Patch, Commit, Diagnostics, Error, LimitKind}`
types through `Package::slide_transition`, `edit_slide_transition`, and
`apply_slide_transition`. Exact name or checked position selectors replace
host slide indices at the API boundary. The focused signatures contain no
native ID, component/member name, generated/Prost/Buffa/wire type, raw slice,
or retained artifact accessor; returned packages are emitted with
`Package::write_to`.

Changed admission proves the complete selected owner chain:

1. the unique rooted Show contains the selected SlideNode reference in its
   SlideTree slides path `[3, 2]`;
2. that node has one expected SlideNode message and one required local field-2
   reference to the selected SlideArchive;
3. both edges resolve uniquely, occur once in aggregate metadata, and have at
   most one matching path-local metadata record with no competing path; and
4. the selected SlideArchive has one expected message whose strict transition
   projection equals the archive-free value, while the selected node's strict
   field-7 marker agrees with transition effect presence.

The ownership scan walks the rooted Show slide-node list once and uses the
package's sorted, globally unique object index for each node lookup. The audit
is `O(slides log objects)` and charges aggregate node-message and local
reference-payload bytes to `LimitKind::WireWork`; it does not reset a work
allowance for each candidate node.

Strict raw preflight runs before the five-message private Buffa lazy-view
projection. The 2,347-byte derived schema is checked against the canonical KN
declarations, contains neither repeated generated storage nor a production
encoder, and produces five files/208,052 bytes under the 224 KiB limit. Buffa
is a borrowed semantic cross-check; the validated raw messages remain the
preservation and splice authority. The strict decoder shares one aggregate
field budget and one strict-plus-Buffa work budget across the selected
SlideArchive, transition, attributes, and animation envelopes instead of
granting each nested envelope a new ceiling.

Changed-only guards reject noncanonical object framing and selected
`should_merge`, base-message, diff/merge-version, diff-field-path,
fields-to-remove, or diff-read-version state. The rewrite changes only the
SlideArchive field-4 transition subtree and, if effect presence changes, the
SlideNode field-7 marker. Co-located owners require one rewritten component;
split owners require at most two, and each component is parsed/reassembled
once. Candidate reopen under retained limits re-proves ownership, semantic
state, marker consistency, object/message metadata, and byte locality.

All unselected ZIP members, IWA objects/messages, raw unknown fields, reference
metadata, the three root previews, `Index/ViewState.iwa`, and slide/node
playback caches remain exact. This playback-only transaction deliberately does
not invoke root-preview deletion. An equal edit shares the source and does no
reassembly/reopen. `Edit::clear` uses Keynote's modern no-effect
settings when an editable envelope exists; if the selected transition is
already absent, clear is an idempotent exact no-op and does not invent a native
owner. Changed apply checks the exact complete source and semantic/ownership
preconditions, then reopens the patch's stored target. Replay, tamper, wrong
source, and inverse-on-source conflict; valid inverse-on-target restores the
complete source exactly.

Legacy nested physical packages remain readable and exact on no-op paths. A
changed set/clear returns `transition::Error::UnsupportedSource` instead of
normalizing their physical provenance.

Migration removes exactly:

- `KeynoteEditor::slide_transition`, `set_slide_transition`, and
  `clear_slide_transition`;
- the `transition_lifecycle` module declaration and
  `keynote/editor/transition_lifecycle.rs`;
- `clear_keynote_transition.rs`, `edit_keynote_transition.rs`, and
  `set_keynote_transition_effect.rs`; and
- five whole direct host mutation tests covering lifecycle, custom/animation
  parameters, marker synchronization, and transition/locality CRUD.

The focused `edit_slide_transition` example owns selector-first immutable
mutation and exact inverse output. This is not deletion of every host read or
creation seam: `KeynoteSlideInfo.transition` and host slide readers remain;
`transition_wire.rs` is retained for `KeynoteEditor::slides()` aggregate
decoding and no-op validation, while creation uses the separate
`creation.rs::transition()` helper and retained `create_keynote_transition`
workflow. Across the exact three methods, module/source, three examples, and
five tests, host transition scope changes by +120/-998 lines, net -878.

No manifest edge or debt item closes. Debt 014
(`litchi-iwa -> litchi-keynote`) remains. Topology stays at 64 workspace
packages, 235 internal dependency declarations, 14 `litchi-iwa` dependency
declarations, and 14 ordered debts.

Verification passes 8/8 focused transition integration tests, 79/79 Keynote
library tests, warning-denied 6/6 doctests, 7/7 root-facade tests with
`--features keynote`, 6/6 transition-codec tests, and the retained host
transition conversion/reader suites at 3/3 and 7/7. The shared common batch
gate passes 10/10 focused and 140/140 full tests plus strict library Clippy;
the archive exact-artifact gate reports 79 unit and 2 integration tests.
`cargo check -p litchi-keynote --all-targets` and
`cargo check -p litchi-iwa --lib`, the host no-run gate, formatting, and
focused diff checks pass. The boundary regression suite passes 101/101. Every
fuzz bin checks, and stable
control-flow smokes for generated no-op, fixed clear, and fixed set completed
six bounded runs each. Expected missing-sanitizer-symbol warnings mean these
were not sanitizer-backed runs.

The final Computer Use gate ran Apple Keynote 14.4 (7043.0.93) on disposable
copies of source SHA-256
`ab186d8d59c858e1b3c2596fd45463cec75ddd92e9fda9032da656a940e68dca`.
Pristine Rust Magic Move and clear candidates reproduced SHA-256
`d5d24386cb544374f4c26da4349f7be961be34180a4536578616886a56af8c1a`
and `5235a3d03dbabced6d06a03b4873826da8602d97f478c61f6467b35d732a08e5`;
their public inverses each restored the exact source. Both candidates opened
without repair, recovery, conversion, or warning. The inspector showed Magic
Move, 2 seconds, Automatic, and a 2.25-second delay before Save As and after
close/exact-path reopen. Clear showed No Transition Effect while preserving
Automatic and the 2.25-second delay through the same lifecycle.

Native resaves were respectively
`dda5049cf431b5c88ea0a9fb209c67edc0d7f0764c23a17eb4e9fdf947d786f6`
and `784069ca8bd2729829bcf204cccdced93f7fbea2b5f8c6b3e4965b47ef423e94`.
Equal restaging over both native artifacts reported `changed=false`,
`touched_components=0`, and byte-exact output/comparison; each no-op inverse
was exact at the same native hash. Remaining shared debt is aggregate
transaction peak-memory/total-work accounting, complete fallible-allocation
proof, process-local complete-source/target patches without stable semantic
serialization/read-write sets/composition/merge/history, library-owned atomic
durable save, and sanitizer-backed fuzzing. `write_to` is exact output, not a
flush/sync/rename durability primitive.

## 2026-08-10 amendment: Numbers table-header transaction cutover

The semantic value does not move in this cutover: the pre-existing
`litchi_numbers::table::headers::{Count, Settings}` continues to own compact
checked counts, all seven optional presence states, and effective-value
helpers. The transaction family is nested beside it as
`table::headers::transaction::{Edit, Patch, Commit, Diagnostics, Error,
LimitKind, Path, InvalidReason}` rather than introducing flat aliases. The
selector-first migration map is:

- `NumbersEditor::table_header_settings(selector)` becomes
  `package.table_header_settings(sheet_selector, table_selector)`;
- `NumbersEditor::set_table_header_settings(selector, settings)` becomes
  `package.edit_table_headers(sheet_selector, table_selector)?.set(settings).commit()?`;
  and
- later work begins from `commit.package()`/`commit.into_package()`, while an
  exact patch is replayed with `Package::apply_table_headers`.

The focused method/type signatures expose no native object ID, package path,
generated message, raw field, or new source artifact accessor. `write_to`
remains the exact output seam. `Edit::settings` borrows the staged value and
infallible consuming `Edit::set(self, Settings) -> Self` replaces it without a
second selector or package lookup.

This is a selector compatibility break: the host selected from a
workbook-wide table catalog, while the focused table selector is explicitly
scoped by its sheet selector.

Changed selection proves document field 1 to the rooted Sheet/FormBasedSheet,
the sheet drawable path `[2]` or form path `[1, 2]` to TableInfo, and TableInfo
field 2 to the selected TableModel. Every followed reference needs unique
resolution, exact aggregate metadata, and optional unique matching field
metadata. A competing rooted TableInfo owner or contradiction in selected
owner metadata is rejected; detached/unrooted pseudo-sheet references are not
promoted to owners and remain exact opaque state.

Only a changed edit checks the selected table's interactive lock. Header and
footer counts remain `1..=5` when present; effective header rows plus footer
rows must fit the table row count, and effective header columns must fit its
column count. Strict admission preserves absence versus explicit false/count
for TableModel fields 9/10/11/12/13/29/32 and rejects duplicate,
wrong-wire, or noncanonical selected encodings. Finite field, nesting,
allocation, traversal, rewrite, output, and aggregate-work ceilings are charged
before publication; retained-target application includes conservative source
plus target work before target reopen.

Strict raw preflight precedes and is cross-checked against the private Buffa
lazy view; raw selected records retain preservation authority. The deterministic
generated closure is five files/51,480 bytes, no repeated views, SHA-256
`5a94caa4620c56bb464792084c01325cef01744bebac97ef948466b9dea105dd`.

Changed dependency admission is exact:

- any selected TableModel field-85 pivot reference blocks every change;
- field 81, field 84, field 86, or nonempty field 83 blocks header-row/column
  counts, while active category/group state decoded through fields 81/83/86
  also blocks all section-count changes;
- selected TableInfo role/alias fields 4/5/7/8/15/16/17 are strictly decoded;
  active aliases block header counts, and section aliases 5/15/17 plus true
  field 16 also block footer/section counts;
- a rooted HeaderNameMgr blocks header-row/column counts; and
- changed repetition refuses deprecated sheet-level repeating-header field 4.

Each case returns `Error::UnsupportedDependency`. Footer/freeze/repetition and
dependency-free count changes remain admissible; the focused transaction never
attempts a partial dependency rewrite.

For an admitted change, only the selected TableModel header fields are
authorized to differ and its component is rewritten once. That locality rule
does not assert that all native count edits are TableModel-only. Candidate
reopen under retained limits verifies semantic readback, presence, ownership,
and byte locality. It deletes the existing zero-to-three
root previews as an explicit rendering-cache exception while preserving
`Index/ViewState.iwa` and every unrelated member/object/message/unknown byte.
An equal edit shares the source, keeps previews, reports zero touched
components/deletions, and does no changed-only lock, reassembly, or reopen
work.

Exact source/target artifacts authorize apply and inverse: stale, replayed,
tampered, or cross-package application conflicts; changed apply reopens the
retained target only after matching the exact retained selected source payload
and conservatively preflighting aggregate source-plus-target transaction work;
valid inverse-on-target restores the complete source and its previews. The
patch is immutable and reversible but remains process-local, unserialized, and
non-durable.

Cutover deletes exactly two public Numbers editor methods, two whole dedicated
mutation tests, one duplicated `Count` unit test, and
`edit_numbers_table_headers.rs`. Ten mixed structural/sort tests survive and
are migrated to private package helpers; seven surviving creation/topology
examples use focused `Package` handoffs. The `table_headers` module/source,
wire codec, attached read/set helpers, package bridge, row/column/sort callers,
and Pages/Keynote owners deliberately remain.

The focused implementation is physically divided into private `api`,
`dependencies`, `error`, `ownership`, `resolve`, and `rewrite` modules, all
under 600 lines, without changing public names or package visibility.
Category-owner group-reference declarations are collected and checked in one
fallible linear pass before group resolution. The required exactly-once
aggregate declaration and optional unique `[1]` field-path proof therefore do
not regress to quadratic per-identifier rescans.

Rooted canonical and accepted legacy TableInfo/TableModel roles remain
supported when unambiguous. Nested legacy physical packages keep exact reads
and no-ops, but changed publication returns `UnsupportedSource`. Locked reads
and no-ops remain admissible; changed edits refuse. Changed admitted edits also
delete root previews rather than retaining the old setter's stale rendering.
No manifest edge closes, so debt 015 and the
64-package/235-internal-declaration/14-debt topology remain.

The native dependency oracle used Apple Numbers 14.4 (7043.0.93) to change
source SHA-256
`f225d5b1cd59e9da454f91a96fe8f81154bc31037c10029230e75d49b45fb693`
to two header rows and two header columns. It opened/saved without warning,
preserved B2/B3, and produced a 136,213-byte artifact with SHA-256
`5c2323b509e5ea9a975b5f254bbd46cf42657aa1c3858d2c7e98f30f07e4b40c`.
Apple changed TableModel 904538 fields 9/10; expanded HeaderNameMgr 904995 from
105 to 157 bytes with a second row UID
`15231182135482363025,1922104131677953016` and column UID
`6719848427115008738,16566391804491244060`; added manager tile reference and
object 905526/type 6365; and changed CalcEngine 904977 formula count from 5 to
30 together with dependency references, locale, and timestamp. This is the
reason for typed manager-backed count refusal, not a Rust writer/parity gate.

The admitted freeze oracle began from the same pristine source. Numbers 14.4
toggled Freeze Header Rows off, retained 1/1 header counts and B2/B3, and
autosaved 136,199 bytes with SHA-256
`015568e6b922e80fbfb760491dc49994ccc2218356ed197131beb46c1bd75850`.
TableModel 904538 differed exactly at field 12, `Some(true)` to absent, and
HeaderNameMgr 904995 was unchanged. The native off-to-on control produced
SHA-256
`df44ed7d0b12c1d372dad7ad7361ed1140d41967921ee42b71a4072b78615721`.
Both native saves regenerated semantically equivalent ViewState topology and
payload while assigning different IDs. This is compatible with the focused
writer preserving raw `Index/ViewState.iwa`; it is not evidence that native
Save churn is byte-exact.

Executed evidence passes 8/8 focused `table_headers` tests and the same 8/8
with `--no-default-features`, 4/4 filtered codec tests, 2/2 root-facade tests
with `--features numbers` (one headers and one names), and 114/114 boundary
regressions. `cargo check -p litchi-numbers --all-targets`, formatting, and
diff checks pass. `cargo test -p litchi-numbers --doc` has one passing
compile-fail test and one ignored example; warning-denied
`cargo doc -p litchi-numbers --no-deps` passes. Strict Clippy finds no new
header-file issue, while full-crate Clippy remains baseline-red on unrelated
pre-existing codec/extractor/table warnings and is not represented as green.

The `numbers_table_headers` target passes
`cargo check --manifest-path crates/litchi/fuzz/Cargo.toml --bin numbers_table_headers`.
The stable fixed-input control-flow executable
completed eight `basic.numbers` runs after its Archive `InputBytes`
versus streaming `InputTooLarge` expectation was corrected. Its expected
missing-sanitizer-symbol warnings make this neither fuzzing nor sanitizer
evidence; no nightly cargo-fuzz sanitizer execution is claimed.

The focused CLI source and exact inverse hash are
`f225d5b1cd59e9da454f91a96fe8f81154bc31037c10029230e75d49b45fb693`;
the Rust changed artifact is
`a8b88d21806b547a5265c60662610f68f524173cac1ca4252d368596c8ef8d2a`.
Diagnostics reported changed=true, one touched component, and three deleted
root previews. This establishes exact artifact/locality behavior but does not
claim a native UI open of that Rust artifact.

After the private module split, a separate freeze-row-only Rust candidate with
SHA-256
`c938d74bcf04be692097488af838f5105a8470e337eafa06fdc8b94b36231d6a`
opened in Numbers 14.4 through Computer Use without repair or warning. The app
reported Table 1 as 22 rows by 7 columns, header/footer counts 1/1/0, an
unselected Freeze Header Rows menu item, B2's fixture text, and B3 value 42.
The exact inverse matched the pristine source hash.

Remaining debt is aggregate transaction peak-memory/total-work and complete
fallible-allocation accounting, process-local complete-artifact patches without
stable semantic serialization/read-write sets/composition/merge/history,
library-owned atomic durable output, baseline Clippy cleanup, and a
sanitizer-backed fuzz campaign. `write_to` remains exact streaming output, not
flush/sync/rename durability.

## 2026-08-10 amendment: Keynote placeholder-visibility cutover

The focused cutover surface is
`slide::placeholder::{Kind, State, Edit, Patch, Commit, Diagnostics, Error,
LimitKind}` and `Package::{slide_placeholder_visibility,
edit_slide_placeholder_visibility, apply_slide_placeholder_visibility}`.
These new method/type signatures carry semantic values rather than generated
messages or source bytes. An edit exposes its current state and consuming,
infallible `set`, `show`, and `hide`; a missing role reads as `None` and cannot
be created through this API.

This intentionally canonicalizes `SlideTextRole` to the shared
`slide::placeholder::Kind::{Title, Body}` discriminator used by both slide-text
and visibility operations. It is a source-breaking migration; sharing the
discriminator does not merge the operations' distinct ownership and mutation
contracts.

The compatibility map is exact:

| Role | Stable SlideArchive reference | Hidden | Visible |
| --- | --- | --- | --- |
| title | field 5 | selected ref absent from fields 7 and 42 | selected ref occurs exactly once in fields 7 and 42 |
| body | field 6 | selected ref absent from fields 7 and 42 | selected ref occurs exactly once in fields 7 and 42 |

Showing appends the selected reference to both lists; hiding removes only that
reference. This preserves other roles, date objects, ordering among remaining
drawables, placeholder content, and unknown fields. It does not transfer
slide-number, layout, placeholder creation, text-box, or style mutation.

Admission starts at the unique rooted Document show reference `[2]`, follows
Show/SlideTree `[3,2]` to SlideNode field 2 and the selected SlideArchive, and
requires exact aggregate reference metadata, unique ownership, and
placeholder/slide co-location. The raw scan is cross-checked against the
placeholder Buffa lazy view: title/body native roles 2/3, parent-slide path
`[1,1,1,2]`, unlocked path `[1,1,1,5]`, and agreeing modern/deprecated storage
references. Changed edits reject role aliases, selected object or slide-number
aliases, conflicting list-local metadata, merge/base/diff state, noncanonical
framing, nonzero slide layering field 41, selected cache fields 37/38,
layout-level visibility overrides, and builds targeting the selected
placeholder.

No-op commit and no-op patch application neither reassemble nor reopen and
retain exact bytes. A changed commit rewrites one slide component, or that
component and a separately stored SlideNode component, invalidates the selected
node rendering cache, deletes all three root previews, and reopens the result.
Changed patch application first validates the retained selected source payload,
then rewrites and reopens the supplied target. Exact artifacts make inverse
application byte-restoring. `Index/ViewState.iwa` and unrelated entries remain
preserved. Changed legacy nested sources are rejected rather than normalized.
Ownership preflight builds linear indexes for payload occurrence/kind and
metadata declarations. Increasing the bounded indexed fixture from 4,096 to
8,192 objects consumed no more than 2.3x recorded work. The budget-aware
SlideNode path conditionally invalidates and direction-aware exact-verifies in
one pass and merges exact reference/wire-work charges. Verification uses only
the bounded, fallibly allocated occurrence/declaration indexes; it makes no
full node/payload clone or verification rewrite. Zero remaining allowance
fails atomically before publication. Work accounting charges every
`MessageInfo` and `FieldInfo`, including empty records; 4,096 empty
`FieldInfo` records fail atomically under both zero and payload-only allowances.
The slide router's pre-allocation work budget is exactly
`source + output + 2 * fields`.
Full precharge includes selected and nonselected payload bytes; metadata
vectors, paths, features, and bases; every aggregate/`FieldInfo` reference in
both `Work` and `References`; and `header_length`. The low-allowance regression
atomically rejects a 256-KiB sibling plus 2,048 references/vectors.

Native Keynote 14.4 establishes the membership convention. The pristine
500,058-byte fixture is
`3a3d07476b45b6e543bcfba75fe38a245434176dcb3565e34570b817708b9f42`.
Title hidden
`d61a92b212d8a0f001bdfc24490d846e065b96885f0d0d0b86ef0be9f10e7580`
became reshown
`9d914ea25a42aaced4459a429e776b09b2024e2858133369f159dad7bce67325`
with title appended after body. Body hidden
`05ca9617ea5a23c57252c28c3029af96d4ec54345de331571d89b612566b8416`
became reshown
`8ee6ac8230273def64450b4cee86c9678849d77b5a7fbd11eb88e0c786279eee`
with body appended. Computer Use confirmed the checkboxes and canvas, retained
date/other role, and close/reopen behavior. Apple's regenerated caches support
semantic compatibility, not raw cache equality.

A Rust-authored title-hidden gate used that pristine artifact and its inverse
restored the exact pristine hash. The candidate SHA-256 was
`df119410433b97b9993d46619764a8ffb75f257b16c0680cd54faabd9a453cdd`;
diagnostics reported changed=true, two touched components, and three deleted
root previews. Keynote 14.4 opened it without warning with Title off, Body on,
and the body/date retained. Save As, close, and reopen kept the same state. The
475,102-byte native resave was
`c5c996415191758b9fc638a8fdf024a912a6fe2ac4c3989970f0cb611e0670e3`.

Rust also passed both Apple-hidden-to-shown directions. Title source
`d61a92b212d8a0f001bdfc24490d846e065b96885f0d0d0b86ef0be9f10e7580`
became
`3d36d31c6222b7622cab180f6dd9559ccf43f4b481e6b245c9d2c56fe8852b2c`,
then its inverse restored the exact source. Body source
`05ca9617ea5a23c57252c28c3029af96d4ec54345de331571d89b612566b8416`
became
`3e8855e954c16bd32350e057665b5ee4758a02e85ad23c3c6543f1caef177b13`,
then its inverse restored that exact source. Each show reported changed=true,
two touched components, and three deleted root previews.

The cut removes the three direct
`KeynoteEditor::{set_slide_text_placeholder_visible, set_slide_title_visible,
set_slide_body_visible}` mutators, public `KeynoteSlideTextPlaceholder`, the
entire 150-line `keynote/editor/placeholder_visibility.rs` source and its module
declaration, two whole direct tests and their exclusive constant, and the
30-line `set_keynote_placeholder_visibility` example. Five assertions in mixed
layout tests now read through the focused package. Shared placeholder ownership
and the layout and slide-number paths remain.

The completed gate is 94/94 Keynote library tests, 18/18 filtered slide-preview
tests, 5/5 focused visibility integration tests, 25/25 slide-text integration
tests, 8/8 root facade tests with `--features keynote`, 7/7 doctests, and
129/129 boundary regressions. Keynote all-target and host-library checks,
warning-denied library Clippy and rustdoc, formatting, and diff checks pass.
The expanded `keynote_slide_text` fuzz target compiles and its bounded stable
control-flow smoke completes, but expected missing sanitizer symbols mean it is
not sanitizer-backed fuzz evidence. The two-way native/exact-artifact results
above complete compatibility verification. No dependency edge or debt item is
removed.

## 2026-08-11 amendment: per-slide Keynote slide-number cutover

The title/body placeholder transaction now admits `Kind::SlideNumber` for
visibility only. This supersedes only the preceding amendment's statement that
per-slide slide-number mutation remains in the host. It does not absorb
presentation-wide `KN.ShowArchive.slideNumbersVisible` field 6, which remains
part of `show::Settings`, nor layout, creation, text, or style mutation. The
canonical surface remains
`slide::placeholder::{Kind, State, Edit, Patch, Commit, Diagnostics, Error,
LimitKind}` plus the three Package read/edit/apply methods. `Edit::set`,
`show`, and `hide` are consuming and infallible staging; preflight, commit, and
apply carry typed failures. Slide text deliberately rejects
`Kind::SlideNumber`.

Migration is one-for-one:

- `KeynoteEditor::set_slide_number_visible(index, value)` becomes
  `Package::edit_slide_placeholder_visibility(SlideSelector::index(index),
  Kind::SlideNumber)?`, then consuming `.show()` or `.hide()`, `.commit()`, and
  `write_to`.
- `KeynoteSlideInfo::is_slide_number_visible` remains a host read for existing
  broad callers; focused code uses
  `Package::slide_placeholder_visibility(..., Kind::SlideNumber)`.
- The surviving slide-number creation example builds the graph in the host,
  reopens it as a focused Package, commits the semantic visibility edit, and
  writes the result. The focused owner does not synthesize a missing
  placeholder.

Preflight proves the rooted Document field-2 -> Show/SlideTree `[3,2]` ->
SlideNode field-2 -> SlideArchive owner. SlideArchive field 20 must identify
one native-kind-1 slide-number placeholder. Visible is exactly Node field 18
true plus one selected reference in each of Slide fields 7 and 42; hidden is
field 18 false/absent plus no selected reference in either field. Showing
appends the canonical reference after the existing occurrences of each field;
hiding removes only the selected occurrences. Global scanning rejects a
competing rooted slide owner, aliases into title/body/object/template/build or
the reserved storage/dependency closure, and contradictory membership. Strict
field-18 parsing rejects duplicates, noncanonical keys, and values other than
zero or one. Exact hidden no-ops preserve absent versus explicit false; a
changed hide uses canonical false, with exact inverse bytes retained by the
patch.

Storage id zero is a supported native form and introduces no metadata reference
to zero. A nonzero storage must be one same-component type-2001
`TSWP.StorageArchive`: kind absent/3, `in_document=true`, text one U+FFFC, and
one attachment-table entry at character zero resolving to one same-component
type-2043 slide-number attachment. The storage's aggregate metadata and
optional dependency declarations/paths must exactly account for the attachment
and its style-sheet/attribute-table closure. The attachment has empty/absent
textual super, absent/zero kind, and no object references. Style visibility
overrides and unsupported closures return `UnsupportedSource`. Legacy nested
packages remain readable and exact no-ops remain exact, but changed edits are
the intentional `UnsupportedSource` compatibility break.

The format seam combines strict raw framing with forced Buffa lazy views. The
new `KNSlideNumberArchive.proto` projection covers Node field 18, storage
fields 1 and 10 plus the borrowed field-9 attachment table, and the attachment
textual super. Handwritten code routes the repeated table without a generated
repeated view or encoder, then cross-checks Buffa against raw parsing. Rooted
ownership/storage work lives in
`package/slide_placeholder_visibility/slide_number.rs`; the scalar splice and
direction-aware exact-delta check live in
`package/slide_preview/slide_number.rs`; shared transaction, drawable
membership, and reopen verification remain in their focused modules. The
generated build evidence is five files/112,101 bytes, zero repeated views,
cap 116 KiB, and aggregate SHA-256
`eacce4103b5c9f9f32fd98639b81249ae1d15fcd63da6fe636569e0a2a324c30`.
It measures deterministic build output, not preservation provenance.

Resource enforcement includes message bytes, fields, aggregate work, nesting,
rooted objects, payload occurrence and metadata declaration indexes, every
aggregate/`FieldInfo` reference, selected and nonselected payload bytes,
header length, output allocation, forward/inverse exact-delta verification,
and archive reassembly. The codec report is merged into the transaction
budget. Index allocation is bounded and fallible; there is no full node or
payload clone and no verification rewrite. Low allowances reject atomically
with content-redacted diagnostics.

No-op commit/apply shares the exact source and performs no reassembly, reopen,
component touch, or preview deletion. A changed commit rewrites the Node and
Slide archives, deletes each existing root preview, reassembles, and reopens.
A changed apply checks the exact retained selected source and exact stored
target, charges source-plus-target transaction work, then reopens the target.
The patch can restore absent/false framing and exact list positions for an
exact inverse, but remains process-local and unversioned.

Only Node field 18 and the selected Slide field-7/field-42 membership are
mutable. One component is touched when Node and Slide co-reside and two when
split. Unlike title/body visibility, the selected node thumbnail/cache is not
invalidated. `Index/ViewState.iwa`, other slide roles and slides, storage,
attachment, content, geometry, style, dependency bytes, unknowns, and global
Show field 6 remain exact. Changed output deliberately removes root
`preview.jpg`, `preview-micro.jpg`, and `preview-web.jpg`; those are the sole
preservation exception beyond the selected semantic delta.

The host retirement removes one public mutator,
`KeynoteEditor::set_slide_number_visible`; the complete 172-line
`keynote/editor/slide_number.rs` source and module; the 23-line direct mutation
example; and two whole direct tests plus four exclusive constants and their
fixture helper. The 53-line creation example remains, with its post-build edit
migrated to the focused package. Creation helpers/tests,
`KeynoteSlideInfo::is_slide_number_visible`, shared ownership, layout, title/body
visibility, and global Show settings remain. No Cargo edge or recorded debt is
removed.

The exact Rust/native gate starts from 500,058-byte `basic.key`, SHA-256
`3a3d07476b45b6e543bcfba75fe38a245434176dcb3565e34570b817708b9f42`.
The 455,859-byte shown candidate was
`a2dafcd4ffc57bafc3bbf7d7cd4ee8131bab2c06dd52adc292632d4208c126be`,
with changed=true, two touched components, three deleted previews, and an
inverse exactly restoring the source. Keynote 14.4 (7043.0.93) opened it
warning-free, displayed slide attachment `1`, checked Format > Slide Number,
and preserved title/body/date. Save As, close, and exact-path reopen preserved
state and content at 500,192 bytes and SHA-256
`b1edd073d309157d27508baf4aedbe93d6dee0687f727dd71f1e8232f6171882`.
Native Save As regenerated root preview hashes
`a7e0fafd160545583d3613b25211925991d9c70a47102cce49b7ddb53d3baab9`,
`fd4d7f8601683d404104d9535fda1fa957f05913a1a40b29c35c51d0e2c1e8db`,
and `c76e1b0b4b2c833232bc23d9c0bbb70255f557233ceba15f24d93d91baab43b0`,
while cached Data9074 remained exact at
`575645e2455199d7cc0c65fab8002b9e025765ba19b8b03c6e51c000f4915e89`.
Independent Apple-only hidden/visible/rehidden controls confirmed that the
selected native delta is field 18 false-to-true, retained field 20, one
field-7 and one field-42 append, metadata lengths only, exact cache data, and
an exact global Show field-6 closure.

The frozen-tree gate passes 8/8 focused slide-number codec tests, 98/98
Keynote library tests, 7/7 focused placeholder-visibility integration tests,
22/22 filtered slide-preview tests, 9/9 facade tests with `--features keynote`,
and 7/7 doctests. Keynote all-target checking, strict Keynote library Clippy
and rustdoc, host library check/no-run and examples, formatting, and diff
checks pass. The fuzz target checks and completes a bounded 16-run stable
control-flow smoke; expected missing sanitizer symbols make this control-flow
evidence, not sanitizer-backed fuzzing. Boundary regressions pass 138/138;
the live slide-number host, placeholder host, and focused audits are clean.
The full checker retains only the unchanged 14 dependency-policy baselines.
Exact-artifact and native compatibility verification is complete.

## 2026-08-11 amendment: Keynote soundtrack-settings cutover

The earlier semantic-only migration is superseded for playback-setting reads
and writes. The canonical direct namespace is now
`soundtrack::{Mode, Settings, Edit, Patch, Commit, Diagnostics, Error,
LimitKind}`, reached through
`Package::{soundtrack_settings, edit_soundtrack_settings,
apply_soundtrack_settings}`. Media items remain deliberately absent from
`Settings`. `None` from the read means no rooted soundtrack object, whereas
`Some(Settings::default())` means an object exists with both scalar settings
absent. Editing the first case fails with `SoundtrackNotFound`; the focused
transaction never synthesizes a soundtrack or media.

Migration replaces a direct `KeynoteEditor::soundtrack_settings` read with
`Package::soundtrack_settings`, and a direct settings mutation with
`Package::edit_soundtrack_settings()?.set(settings).commit()` plus `write_to`.
`Settings::new`/setters preserve presence and validate finite inclusive volume
and canonical known modes; future mode discriminants remain lossless.
`Edit::set` consumes the edit and cannot fail because `Settings` is already
validated. Patch application is against an exact immutable package snapshot,
not a long-lived host editor.

Changed preflight proves Document object 1/type 1 field 2 -> unique Show/type 2
-> Show field 17 -> unique type-21 `KN.Soundtrack`. Both object edges require
exact aggregate and field-path reference metadata, nonzero/nonexternal
identifiers, role disjointness, selected-message uniqueness, no merge/diff
state, and bounded component framing. Only Soundtrack fixed64 field 1 and
varint field 2 may change. Their absent/present spelling is semantic and is
rewritten with canonical framing.

Soundtrack field 3 is a retained resource boundary, not an opaque unchecked
tail. The codec streams every canonical nonzero movie-media data reference and
cross-checks its order with the selected message's aggregate data references
and optional exact field-3 metadata. Mutation additionally proves the matching
PackageMetadata component/data records, selected soundtrack owner/count, safe
relative data filename, and unique `Data/` entry. The media field records,
metadata, data file bytes, ordering, and all unknown soundtrack fields remain
exact.

The focused `KNSoundtrackSettingsArchive.proto` projection covers only the
two Soundtrack scalar settings; strict traversal validates the Show reference
separately. Handwritten raw preflight runs before forced/cross-checked Buffa
lazy views and streams media without a generated repeated collection. No
generated encoder participates in the rewrite. The deterministic build
produces five files/27,753 bytes, no repeated views, under 32 KiB, with
aggregate SHA-256
`458206e0b57d8ec5ae4c3fc706bf793ccd385ab867b7e92ac30d66ab1858b4d3`.
This is code-generation provenance; accepted raw records remain authoritative
for preservation.

The shared budget incorporates codec bytes, fields, work, nesting, media
reference count, and media payload bytes, then charges rooted/reference
metadata, PackageMetadata/data-member closure, component framing, output and
compression, archive reassembly, reopen, and exact candidate comparisons.
Allocations are fallible and errors remain typed and content-redacted. No
separate performance closure is inferred from these enforcement paths; shared
allocation, peak-memory, work-bound, output, durable-save, and process-local
patch debts remain.

No-op commit/application returns the exact source without component rewrite or
reopen and reports unchanged diagnostics. A changed commit rewrites the one
soundtrack component, reassembles, reopens once, and verifies every package
member plus the selected object/message delta and only its necessary ZIP
CRC/size/offset bookkeeping. Changed apply exact-authorizes the source and
retained target before reopening the target. Inverse artifacts
restore source bytes exactly. Legacy/non-exact provenance remains readable and
supports exact no-op behavior, but changed mutation intentionally returns
`UnsupportedSource` rather than normalizing it.

This is a playback-only settings cut. It preserves root previews, slide and
node caches, ViewState, every slide, field-3 item order, media bytes and
metadata, and unknown records. It does not retire the separate host soundtrack
item APIs, `KeynoteSoundtrackItemInfo`, soundtrack creation, data allocation,
replacement/reclamation, the item example/tests, or the shared wire/media
helper required by those paths.

The native gate used 506,640-byte Apple-resaved populated source SHA-256
`69795554212651b261f5ffd71dd5cf511544f285cab680d724a9de7d3f04b14d`.
The same-size Rust Loop/0.35 candidate was
`6367e38a2edeebe6e65b148d0fd2aae555ee219dc1a65c339954047eb533ce1a`;
only `Index/Document.iwa` differed and inverse application exactly restored the
source. Keynote opened warning-free, showed Loop and volume
0.3499999940395355, retained `ringin` at 00:00:01, and played it. Native Save
As produced 506,651-byte
`e264f4e714b0c44fca420b2c7b43e18f2ed1be99a766d25fe901f68d5f8bc299`.
The `ringin-9075.m4a` payload stayed exact at
`5a08f48c4f86074e14a763d4f19f49ca31196a7a5f52fb48960e76b6f3d3d96b`;
the slide and all three previews were exact, and a focused restage of the
normalized native setting was a byte-exact no-op.

The completed host migration removes
`KeynoteEditor::{soundtrack_settings, set_soundtrack_settings}`, the complete
68-line `keynote/editor/soundtrack.rs` module/source, settings-only
`patch_soundtrack_wire`, and the retained record's now-dead decoded-native
field. Production changes by +2/-91 lines. Two whole settings tests plus their
exclusive import/constants are removed in the 157-line test cut, as is the
29-line direct mutation example. The mixed inspector and README migrate to
focused Package reads/transactions. The item CRUD API and module, shared
wire/media reader and mutation helpers, creation, media lifecycle, item
example, and item tests remain.

No Cargo edge closes: debt 014 (`litchi-iwa -> litchi-keynote`) remains. The
inventory is unchanged at 64 workspace packages, 235 internal declarations,
14 `litchi-iwa` dependency declarations, and 14 ordered debts.

The frozen gates pass 5/5 codec tests, 1/1 focused scaling unit, 4/4 focused
soundtrack-settings tests, 99/99 Keynote library tests, 10/10 facade tests with
`--features keynote`, and 8/8 doctests. All-target Keynote checking, strict
Keynote Clippy/rustdoc,
focused and retained examples, host checks, formatting, and diff checks pass.
The performance review is clean at P0/P1. A test-only `media.rs` regression
routes realistic 4,096- and 8,192-entry metadata/media states through the real
streaming path; references double exactly and fields/work/references stay at
or below 2.3x. That is deterministic scaling evidence, not a wall-clock or
general performance-completion claim. Boundary tests pass 152/152; host and
focused audits each report zero diagnostics. The full
checker retains only the unchanged 14 baselines: six dev-only annotation
findings and eight edge classifications. Native, inverse, no-op, and
preservation verification is complete.

## 2026-08-11 amendment: Numbers sheet-order cutover

The direct sheet-order workflow moves to
`sheet::order::{Edit, Patch, Commit, Diagnostics, Error, LimitKind}` and
`Package::{edit_sheet_order, apply_sheet_order}`. Existing semantic sheet
iteration supplies the read order. Migration replaces
`NumbersEditor::move_sheet(selector, destination)` with
`package.edit_sheet_order().move_sheet(selector, destination)?.commit()` and
`write_to`. The destination is the final zero-based position after removal;
one Edit accepts exactly one move. Empty commits, missing selectors,
out-of-range destinations, second operations, unsupported sources, limits,
allocations, verification, and conflicts remain distinct typed failures.

A changed edit must prove both native order owners. Root object 1/type 1
`TN.DocumentArchive` has ordered sheet references at field 1 and a required
sidebar-tree-root reference at field 5. That unique type-205 root has the same
number of ordered children at field 2. Each child is a unique type-205 node
whose field 3 names the sheet at the corresponding Document position; child
field-2 descendants remain attached and exact. Every role is nonzero,
nonexternal, disjoint, uniquely resolved through the sorted object index, and
located in `Index/Document.iwa`; selected messages have canonical non-merge,
non-diff metadata. Ordinary type-2 sheets are supported. FormBasedSheet or
split-component changed sources fail with `UnsupportedSource`.

Document field-1 sheet references and sidebar-root field-2 child references
must each appear exactly once and in order as a selected subsequence of their
owner's aggregate object references. The transaction reorders those two
subsequences only. Any selected order reference in any `FieldInfo` is refused,
rather than guessing how to synchronize a field-attributed order. The sidebar
edge and child association/descendant references may have optional field
metadata only at exact field-5, field-3, and field-2 paths. Child identifiers,
nodes, field-3 associations, descendants, sheet objects, tables, and data are
preserved.

The only generated projection is
`TNNumbersSheetReferenceArchive.proto`, a scalar `TSP.Reference`. Strict
handwritten routing owns Document fields 1/5 and TreeNode fields 2/3 in two
passes, validates canonical framing and signed deprecated fields, and forces a
Buffa scalar parity check for every selected reference. Ordered repeated data
is never a generated view and publication never uses a generated encoder.
Raw source field records own preservation. The deterministic closure is five
files/32,579 bytes, zero repeated/lazy-repeated views, below 33 KiB, digest
`2a0850fd82cfbf337ed48e582d4a998bd27e5046eb63c61f6939fa5ff1a09854`.

Codec fields/work/depth/reference/reference-byte reports feed one shared
budget. The transaction also accounts for object-index lookup, complete
message and metadata structure, aggregate/field references, raw splice and
metadata reorder, decoded archive extent/allocation, compression, package
reassembly/deletions, source and target reopen, and exact package locality.
Every reservation is fallible and publication is atomic and content-redacted.

A positional no-op reads no native component, shares exact source bytes,
touches/deletes nothing, and skips candidate reopen. A changed source must
contain exactly the three canonical root previews, once each; missing or
repeated preview members refuse publication. Changed commit rewrites one
component, deletes all three, reassembles/reopens once, and verifies both
semantic and physical locality. Forward apply validates preview state 3 -> 0,
and inverse validates 0 -> 3. Changed apply first
authorizes its exact source, precharges source/target/reopen work, reopens the
retained target, and checks the stored moved-sheet identity at the destination.
Applying to another snapshot conflicts; inverse restores the exact source.
Legacy/non-exact provenance remains readable and allows a positional no-op but
changed mutation returns `UnsupportedSource`. Patch is process-local and
unversioned.

The only preservation exceptions are the two raw order sequences, their
aggregate selected subsequences, necessary owner/message/ZIP length fields,
and deletion of root `preview.jpg`, `preview-micro.jpg`, and
`preview-web.jpg`. ViewState, CalcEngine, sheet/table/drawable graphs, all child
IDs/nodes and descendants, global table order, data sidecars, and unknowns
remain exact. Sheet add/duplicate/remove, FormBasedSheet and general Document
reference handling, table/drawable CRUD, ID/component allocation/reclamation,
and the shared host substrate remain separate.

The P0/P1 performance review is clean: no release blocker and no O(S²) path was
found. The scaling regression runs strict codec decode, raw record reorder, and
core aggregate-header reorder for 4,096 and 8,192 references. The production
bound is 2.3x plus a fixed 32-unit allowance for work/references/payload; the
codec-only bound is strict 2.3x. No wall-clock claim is made. P2 tradeoffs are
explicit: the patch retains roughly `4 * sheet_count` snapshots, capped at
4,096 sheets, for no source reselection and O(1) inverse; Vec-to-Arc publication
can briefly duplicate target bytes; and authorization of a separately
allocated byte-equal source may perform one bounded O(package-bytes) comparison
before transaction charging, while allocation identity is O(1).

Matched native controls prove the dual-order contract. The control resave is
133,594 bytes/SHA-256
`f9c5cbec4f422484c63d1d39bd8d09da122d011596561a5feb2ad1e812574990`;
the native reorder is 153,498 bytes/
`7b3bcbc853346a433e84ee815d28671d01fc3da857e43b8b7d29b310f94e7e1a`.
Apple reverses Document field 1 and its aggregate order together with the
sidebar-root children and their aggregate order; child field-3 associations
remain exact. The control no-op retains its three previews, while reorder
regenerates all three. Ninety-three of 103 decompressed members, including all
table data sidecars, are exact. Apple-only cache culling and physical subgraph,
ViewState/tree-ID, package-ID/revision, metadata-order, property, and timestamp
churn are native normalization, not requirements for focused output.

Rust output
`97c76894503a2628c1828babd93d9a9a891794d86c86177cab60f09333997a68`
opened in Numbers 14.4 without repair, recovery, or conversion. Tabs
`FirstCreated`, `SecondCreated` and content markers `A-new`, `A-old`, `B-only`
were correct and CalcEngine preservation was benign. Native Save As, close,
and exact-path reopen retained those semantics at 103 members and SHA-256
`4aa257e4db61a3c03950360b29267c9495985d460ae22b6f679bee31f2693217`.
The regenerated preview hashes exactly match the Apple reorder:
`db372ed754b8702fb964760f5087cedb2b2cfac09ff2d898947458822446c1f6`,
`582e37b9fddd5e669e1929d64f54e31da3c2c22f13cbd0df1e74dfad34543f5e`,
and `6c7a226b0a64d5946cabbc517c5b416a677e23871a7ffd040fb4f225b1ac339d`.
A same-position restage on the native resave was byte-exact with diagnostics
0/0/0/false, and its inverse was exact at the same hash.

The focused implementation is exactly five sources: public
`sheet/order.rs`, transaction `package/sheet_order.rs`, and frozen private
`package/sheet_order/{error,resolve,rewrite}.rs`. The host cut removes
`NumbersEditor::move_sheet` and exclusive `selectors::sheet_index` (-58
production lines), changes tests by +2/-43, and deletes the 23-line legacy move
example. The retained remove example migrates +2/-6 to a semantic selector;
sheet add/duplicate/remove and shared helpers remain.

Focused codec/protobuf verification passes 7/7 and 132/132. Numbers passes
109/109 library tests, 4/4 private sheet-order tests, and 1/1 public integration
test. Boundary tests pass 165/165; Python compilation and diff checks pass; and
live host and focused audits are both empty. The full checker reports only the
unchanged 14 baselines: six missing dev-only `soapberry-zip` annotations and
eight unclassified edges (those six plus the `litchi-odf-common` and
`litchi-opc` edges to `xml-minifier`). Debt 014
(`litchi-iwa -> litchi-keynote`) remains. Topology is unchanged at 64 workspace
packages, 235 internal declarations, 14 `litchi-iwa` dependency declarations,
and 14 ordered debts.

## 2026-08-11 amendment: Numbers table-title transaction cutover

The Numbers-only table-title read/write path moves from direct editor methods
to `table::title::{Settings, Edit, Patch, Commit, Diagnostics, Error,
LimitKind, Path}` and
`Package::{table_title_settings, edit_table_title, apply_table_title}`. A
semantic sheet selector and a sheet-scoped table selector replace raw model
identifiers. The new method/type signatures expose no raw source, component,
object identifier, or generated value; source bytes remain crate-private and
publication uses `write_to`. The compact `Settings` value retains optional
presence for TableModel field 22 visibility and field 37 outline, so `None`
and `Some(false)` remain distinct. `Edit::set` consumes the edit and cannot
fail; validation occurs at commit.

Changed admission reuses the strict rooted Document -> Sheet/FormBasedSheet ->
TableInfo -> TableModel ownership and exact metadata proof from the focused
table-header owner. It rejects effective table locks. When the requested title
is visible, TableModel must also contain a finite nonnegative field-33 height,
exact local field-30 paragraph-style and field-36 shape-style references,
distinct nonaliased style identities, and unique canonical type-2022/type-2025
messages with valid required super framing. Missing or malformed rendering
dependencies return a typed fail-closed error. Changed admission also scans
`Index/ViewState.iwa`: any native type-6284 table-name-selection message is
transient title-selection state and returns `UnsupportedSource` rather than
being inferred or normalized. Reads and exact no-ops retain broad compatibility
because this check is changed-only. On an accepted changed source, all other
ViewState bytes remain outside the write set and exact.

Strict raw preflight owns fields 22/30/33/36/37, then forces and cross-checks a
three-scalar table-title Buffa view plus the reused scalar Reference lazy view.
Generated code never owns unknown preservation, repeated references, or
encoding. The five-file generated closure is 32,332 bytes under 33 KiB, digest
`56cfd70666ffa6079175bdab0a63a4ddd055099edf3c771ed3ad8b3051596ee1`;
the codec passes 9/9 and the complete protobuf suite passes 141/141.

No-op commit and application preserve exact bytes and avoid reassembly/reopen.
A change raw-splices only fields 22 and 37 in the selected TableModel message,
rewrites its `Index/CalculationEngine.iwa` component, deletes each existing
canonical root preview (zero to three), reassembles once, reopens once, and
verifies semantic readback and exact locality. Apply requires the exact source
and stored target; tamper or replay conflicts, and inverse restores exact
source bytes and preview state. Legacy/non-exact changed sources return
`UnsupportedSource`. The patch is an immutable process-local exact-artifact
capability, not a durable or serialized semantic log.

Numbers 14.4 matched controls record a 136,204-byte resave at
`25c9fc858ca4fb4f1fedeafb944e96afb81af03a082a41be297ecf6f2542dbdb`
and a 136,273-byte title-hidden artifact at
`ac8a7117ad6256b0da2e6d191b9e64f721b689d71696a89ac0f78bc6aa513a28`.
The native hidden form removes the raw field-22 occurrence instead of writing
false, while field-37 remains a separately presence-preserving setting. This
does not establish a right to rewrite type-6284 ViewState state.
Changed admission rejects that transient state instead.

The final exact-artifact gate starts from the 136,357-byte source
`f225d5b1cd59e9da454f91a96fe8f81154bc31037c10029230e75d49b45fb693`.
Rust hid the title in a 136,351-byte candidate,
`4c7f6340b6f2675240577c5b59d5c154de24c8a7e763a31257c56a9899a8e40c`,
whose inverse restores the source byte-for-byte. Numbers 14.4 opened it without
warning, reported Table Title off, retained the 22-by-7 table, B2 marker
`Litchi native Numbers fixture`, and B3 value 42, and preserved that state
through warning-free Save As, close, and exact-URL reopen. The 136,353-byte native resave is
`5b162f8431f45333f0ae9a8654dfa724794f2ec2b391ea11f6a5eee7822cbb10`.

The completed performance review has no P0/P1 finding. Real rooted Package
counters for 4,096 -> 8,192 objects are fields 53,307 -> 108,363 (2.0326x),
`WireWork` 315,936 -> 636,752 (2.0155x), references 16,386 -> 32,770
(exactly `2 + 4N`, 1.9999x), and `TransactionWork` 9,084,384 -> 18,298,157
(2.0142x). All remain within 2.3x, and maximum-minus-one work rejects before
output. Remaining P2 costs are linear selector temporary vectors and redundant
changed-edit decode passes; this is resource accounting, not wall-clock proof.

The cut deletes the two direct NumbersEditor methods, 32 production lines,
245 direct-test lines, and the old `edit_numbers_table_title` example. Private
package helpers and wire support remain for Pages and Keynote table-title
read/write paths, together with their format-specific CRUD. Boundary tests
pass 173/173. Final Numbers suites pass 111/111 library, 2/2 private
table-title, and 5/5 public table-title tests; codec/protobuf suites pass 9/9
and 141/141. The full checker retains only 14 unchanged dependency-policy
baselines. The topology snapshot is 64 packages/237 internal declarations/14
`litchi-iwa` dependency declarations/14 ordered debts; debt 015 remains and no
edge closes.

## 2026-08-11 amendment: aggregate Pages section-settings cutover

This cutover is complete only when one focused Pages transaction owns the
physical read/rewrite/publication path for `TP.SectionArchive` fields 17--22,
26, and 28, and the retained name and pagination APIs demonstrably delegate to
it. The public audit must find no raw object identifier, IWA/component/member
name, protobuf/Buffa type, field record, exact artifact, or source byte in the
focused signatures, errors, diagnostics, or `Debug` output.

The exact wire matrix is four optional canonical Booleans, three optional
canonical `uint32` values with a nonzero field-22 page, and one optional
canonical UTF-8 name. Strict raw routing rejects duplicates, wrong wire types,
noncanonical keys/values/lengths, Boolean values above one, invalid UTF-8,
page zero, truncation, malformed or mismatched groups, excess bytes/fields/
nesting/work/name storage, and disagreement with the forced aggregate Buffa
lazy view. Tests must cover all 81 Boolean presence/value combinations while
holding name and pagination values constant, all name-presence distinctions,
known and future pagination discriminants, and every exact/one-over resource
boundary. The projection must remain generated-private, borrow the optional
name, retain no repeated or unknown storage, provide no production encoder,
and have a build-ratcheted generated size and deterministic digest. Those
five generated files total 80,202 bytes under an 80 KiB ceiling, contain zero
`RepeatedView`/`LazyRepeatedView` declarations, and have deterministic
aggregate SHA-256
`2202f4b1d394346450cb9f88a41c2784ab476cff23b181fffbab6f37b4a42b62`.
The complete focused protobuf suite passes 149/149 tests.

Changed-source tests must prove the rooted section position and exact owner,
template prerequisites at fields 23--25 when activated by the target value,
previous-section inheritance closure, and exact rooted layout/cache state.
They must distinguish an absent otherwise-valid prerequisite as
`UnsupportedDependency` from malformed/ambiguous ownership as `InvalidSource`.
The changed byte oracle permits only the selected section payload and exact
object length framing plus required component/ZIP framing. The rooted cache
edge/reference metadata, opaque and detached cache objects, every canonical
root preview, templates, background field 30, fields 29/31, unknown section
records and order, sibling sections, unrelated messages/components, retained
member metadata, and package statistics remain exact. Preview subsets remain
unchanged; duplicate preview names fail ingress rather than becoming a changed
transaction concern.

Transaction tests must cover exact-name and checked-position selection,
missing/ambiguous selectors, duplicate destination names, exact no-op source
allocation identity, no-op legacy admission, changed-legacy refusal, exactly
one-component publication, zero preview deletion, complete
retained-limit reopen, stale/replayed/tampered/competing patch conflicts, exact
apply, exact inverse, failure atomicity, `Send + Sync`, and content-free
formatting. Dedicated equivalence cases must apply every name-only and
pagination-only edit through both its retained facade and the aggregate API and
compare the complete candidate bytes and diagnostics.

The host retirement gate removes `PagesEditor::section_settings`,
`set_section_settings`, and `set_section_name`; the direct raw-ID
`set_pages_section_settings` example; duplicate settings/name tests; and README
usage. It retains the separate section-background behavior by relocating or
renaming its private payload helpers, and retains mixed header/footer or
section-graph tests after migrating only their name-update step. Boundary
ratchets must forbid the three methods, example, stale README calls, and a
second physical writer in the compatibility facades.

The executed gate record must include the full Pages library/integration/
doctest suite, focused codec and full protobuf suites, Pages-feature root
facade, migration-host library and all examples, no-dependency warnings-denied
Clippy, strict rustdoc, formatting and diff checks, boundary regressions, live
boundary disposition, focused fuzz compile/smokes, and the available sanitizer
status. Linear-work evidence must exercise the maximum supported section
population rather than claim timing. The production test uses rooted real
packages with 4,096 and 8,192 total objects. Selected fields are 77 and 77
(1.0x), strict `WireWork` is 564 and 564, and selected references are 4 and 4.
`TransactionWork` is 292,154 and 587,222 (2.0100x). Both changed runs perform
exactly one bounded output allocation and one full reopen. With the configured
transaction-work ceiling set to maximum minus one, the typed limit error occurs
before output and reports zero output allocations and zero reopens. This is
deterministic resource-scaling evidence, not a wall-clock, RSS, or allocation-
latency claim.

The executed focused gate is 7/7 section-settings integration tests plus four
private production/security tests for exact budget observations,
4,096-to-8,192 object scaling, alias-metadata refusal, and repeated-reference
scaling/max-minus-one refusal. The strict/projection suite is 149/149, and the
final locality review has no finding. The complete Pages library/integration
total is 118/118: 67 library, 14 document-settings, 1 native-fixture, 10
page-layout, 5 section-name,
6 section-pagination, 7 section-settings, and 8 section-text tests. Boundary
regressions pass 181/181; the focused facade and host audits each report zero;
and the full checker reports only the unchanged 14 dependency-policy baselines.
Final topology is unchanged at 64 packages, 237 internal declarations, 14
`litchi-iwa` declarations, and 14 ordered debts.

Apple Pages 14.4 supplied a two-section native seed of 101,399 bytes with
SHA-256
`19b8a24c7bc0d57d87614a0f08215072c9c61519b15629827f5a448b29218422`.
Three matched control/change pairs isolate independent second-section Boolean
settings:

- field 17 control is 101,328 bytes,
  `67c1d16ef682814c720bff7f189b539bad486a998f48e79f9fd2282975abe40b`;
  match-previous is 101,323 bytes,
  `af4119b13cf4ff5d4db1fc172a55404b85f6a41833d755d4c1a5d22d40aacda9`;
- field 19 control is 101,376 bytes,
  `7724862901685f14f0c1262391df8464332390d681b5af6328a4de6483d70a7f`;
  left/right is 101,334 bytes,
  `a0956e21dff5b89fba0a2314224fce93bc19a3b32be91228d18b1bfa3032da2a`;
  and
- field 28 control is 101,368 bytes,
  `184753935f12a9d16ed6787e82b43fc9420d63cb47dffec35c217ab03342d438`;
  hide-first is 101,333 bytes,
  `f19639a8c93966b5d1a4b87d07c7908f5aaa3bc6e19de9e7eaee96650c3dbc18`.

Every pair changes one scalar on section object 1732889/type 10011 and nothing
else in the logical package: field 17 `88 01 00` to `88 01 01`, field 19
`98 01 00` to `98 01 01`, or field 28 `e0 01 00` to `e0 01 01`. The other
three selected flags remain explicitly false, including field 18; the 57-byte
message header and references, templates and header/footer storages, all 13
entry names, layout/cache state, and previews are exact. Warning-free Save As,
close, and exact-path reopen showed H1/F1 on pages 3 and 4 for enabled field 17,
H2/F2 on page 3 with blank page 4 for enabled field 19, and blank page 3 with
H2/F2 on page 4 for enabled field 28.

This matched native evidence freezes the scalar meaning and exact-preservation
closure. The focused Rust integration gate independently covers complete
eight-field read/edit/apply, same-settings byte equality, patch conflict, and
exact inverse; no separate claim is made that a Rust-authored aggregate
candidate was the artifact opened for the UI oracle. Producer-name proof
therefore remains focused reverse-read rather than a visual assertion.

This cut changes no manifest edge. Current topology before the cut is 64
packages, 237 internal declarations, 14 `litchi-iwa` dependency declarations,
and 14 ordered debts; the final verification must confirm that unchanged
inventory rather than copy historical Pages counts.

## 2026-08-11 amendment: Numbers table-cell read cutover

The first table-cell cutover transfers semantic reads, not mutation.
`litchi-numbers::table::cells::{State, Storage, Error, LimitKind, Path}` and
`Package::{table_cell, table_cells}` are now the canonical selector-first API.
The single-cell path checks one coordinate. The range path checks a half-open
range, applies retained-element and owned-string-byte limits, then publishes a
fallibly allocated dense row-major result. Missing coordinates are explicit;
a materialized `Value::Empty` remains `Storage::Stored`.

Migration verification must distinguish the current eager semantic read path
from the preparatory physical seam. `Package` reads the immutable semantic
table already produced by `litchi-numbers::package::extractor` and its existing
BNC/protobuf decode. The committed strict table-cell storage and dependency
Buffa codecs pass their own gates but are not invoked by these methods and do
not encode or preserve the source. This supersedes monolith-only BNC/read
wording in this ADR only to that extent. All existing `litchi-iwa` cell
mutators, formula/compiler and AST wire work, calculation-engine mutation,
cache refresh, previews, and publication remain unmigrated; there is no host
cut in this slice.

The performance gate uses analytical counts. With range area `A`, selected
row-span materialized cells `K`, selected owned-string bytes `B`, selected
owned strings `T`, and total materialized cells `M`, a non-empty range costs
`A + 2K + 2*O(log M)` in size-sensitive work, allocates one `A`-element result
vector, and performs `T` fallible string allocations totaling `B` bytes. Empty
ranges scan and allocate nothing. For paired shapes that double from 4,096 to
8,192, every size-sensitive term remains at or below 2.0x and the result
allocation remains one. Maximum-minus-one element and text limits reject
before result allocation. No wall-clock or RSS conclusion is drawn.

Native read evidence is deliberately non-mutating. The 136,357-byte
`basic.numbers` fixture has SHA-256
`f225d5b1cd59e9da454f91a96fe8f81154bc31037c10029230e75d49b45fb693`.
Numbers 14.4 and the focused reader agree that Sheet 1/Table 1 is 22x7, B2 is
stored text `Litchi native Numbers fixture`, B3 is stored number 42, A1 is
missing, A1:C3 is dense row-major, and A23 is out of bounds. The 140,498-byte
formula/rich-text oracle, SHA-256
`80deb7b87df27f58b26e6f247acee9d1fc6dcd3d268e85046c3efc16070b2edf`,
covers formula and rich-text-backed semantic values. Neither file contains an
explicit stored empty cell, so `Stored(Value::Empty)` versus `Missing` is
synthetic unit evidence only. These checks authorize no write or Save claim.

Final verification passes 114/114 Numbers library tests, 4/4 public read
integration tests, 13/13 strict preparatory codec tests, 163/163 full protobuf
tests, and 187/187 boundary regressions. Strict library/test Clippy and
warning-denied rustdoc are green; the full checker reports only 14 unchanged
dependency-policy baselines. No manifest edge or ordered debt changes, and
the inventory remains 64 packages, 237 internal declarations, 14
`litchi-iwa` dependency declarations, and 14 ordered debts.

## 2026-08-12 amendment: Numbers table-cell mutation verification

Mutation now builds on the read cutover through
`table::cells::{Input, Change, Edit, Patch, Commit, Diagnostics,
DependencyKind}` and `Package::{edit_table_cells, apply_table_cells}`. The
selector resolves before staging; the consuming edit bounds update count and
owned input bytes, rejects duplicate and out-of-bounds coordinates, and commits
one sorted final overlay. Exact semantic no-ops share the source and skip all
changed-only ownership, dependency, preview, reassembly, and reopen work.

The current positive matrix is verified separately: finite scalar and clear
batches; direct/unsegmented string-list key assignment/release with exact
refcounts; missing sparse tile growth for finite non-text scalars at the
synthetic 513-row boundary;
in-place authored-text replacement in uniquely owned rich backing while
retaining key/storage identity and releasing exact style references; and
supported formula-cache chains computed once from the final batch overlay.
Formula AST construction remains in the host. Negative gates cover locked
tables, HeaderNameMgr-backed header cells, segmented string lists, shared or
rich text requiring a FieldInfo reference transition,
noncanonical/ambiguous FieldInfo rich ownership, existing formula or error
cells, and modeled unsupported/cyclic/range/deletion/sparse formula dependencies as
`UnsupportedDependency`; a modeled missing storage prerequisite is
`UnsupportedDependency { CellStorage }`, while malformed storage routes are
`InvalidSource` and an unmodeled stored BNC value/source kind is
`UnsupportedSource`. Sparse text-to-missing-tile changes refuse as
`UnsupportedDependency { SharedString }`. Impacted active merge, pivot,
category, spill, hidden, and conditional-style states refuse by their matching
dependency kinds, while unrelated/inert state remains exact. Each failure is
typed and atomic.

Canonical payload field-1-to-storage and storage field-2-to-style FieldInfo
metadata may be present on the unique rich path and remains exact when no
field-specific reference transition is required.

The strict codec foundation is independently ratcheted. Storage generation is
five files/465,932 bytes with SHA-256
`1a894fd5d22b004db664bc7c348d9591a4608ab9263a8122c726c8a1ecb0c3b3`;
dependency generation is five files/544,538 bytes with SHA-256
`2fba7c22aef58ed3cfe6eba1f77e5eaf79d2597dd79966e05d20e50c0e2b33b3`.
Neither projection generates a repeated view or encodes. The separate strict
dependency-only formula projection remains five files/201,539 bytes with
SHA-256
`ccd972b3dcd76b6142342d36435f2f76a305c029265853ced04d64c1e2bf1752`.
Its focused gate passes 7/7 and the expanded full protobuf gate passes 178/178.
The PackageMetadata projection remains five files/145,681 bytes, generates no
repeated view, and has SHA-256
`ee49927f75c6b632c83055f9b7e647920b389be41bec10e25871a6ef7b56ab31`;
its focused gate passes 7/7.
Raw message bytes retain preservation authority. CalculationEngine field 14 is
projected and its rooted HeaderNameMgr reference validated; only the referenced
manager payload and update semantics are unprojected, so a manager-backed
header change refuses as `HeaderNameIndex`.

Changed verification covers grouped exact message replacement/append/delete,
complete aggregate and FieldInfo reference transitions, exact deletion of the
canonical source preview subset, one reassembly, one reopen, semantic readback,
and inverse restoration. `Patch` retains exact artifacts and private verified
source/target packages. Apply borrows the patch, authorizes the exact
directional source/read profile, and reruns locality against the retained
target without another reopen; inverse swaps the same evidence. The patch is
process-local and has no
serialized form, operation log, composition, merge, or history.

Reads and exact no-ops remain broad. Changed publication requires an exact
physical `SourceCatalog`; a semantic-only or nested legacy source fails as
`UnsupportedSource` rather than being normalized.

The completed performance review has no P0/P1 findings. Rooted
4,096-to-8,192 counters are:
numeric retained-elements/bytes/wire/output/transaction ratios are
1.9855/1.9820/1.2694/1.1904/1.1899; unique-text input/string/retained/output/
transaction ratios are 2.0/1.9998/1.9861/1.3213/1.2245; same-tile updates are
2.0x with 1.1396x transaction work, one tile read/write, and two components;
formula nodes, edges, hosts, and work are 2.0x with 1.8021x transaction work.
Required-minus-one formula and sparse cases report zero component,
reassembly, output, reopen, and locality work. These are counters, not elapsed
time or RSS.

Native compatibility is final for the scoped claims. The 136,357-byte source,
SHA-256
`f225d5b1cd59e9da454f91a96fe8f81154bc31037c10029230e75d49b45fb693`,
produced the 79,384-byte Rust B3=43 candidate, SHA-256
`7540c94f61d18fb4a8fe188544eef5854cdb6c06ffa6f1b8b0be1e06264f6b82`,
and inverse restored the source exactly. Numbers 14.4 opened, saved, closed,
and reopened warning-free with 22x7, B2 text, B3=43, and blanks preserved; the
Rust delta was the selected Tile payload plus deletion of three previews.

The formula-rich source is 140,498 bytes, SHA-256
`80deb7b87df27f58b26e6f247acee9d1fc6dcd3d268e85046c3efc16070b2edf`;
the Rust C2 candidate is 80,519 bytes, SHA-256
`5f13dbf7f1f78d3b4f6f313e3d6ca38bfac6985f1039e042a778a031d54c7826`,
and inverse restores the source exactly. Its Numbers-resaved/reopened artifact
is 140,838 bytes, SHA-256
`0ba1f436f3b44b8d1f30084d95dae8f53b5463c7ddc71ee157d5347a5a025060`;
the matched no-edit control is 140,484 bytes, SHA-256
`f8c17f6b69e4d996b6088c0ef8ebe156fe2ef7ed64adeb21e51a9cd8fbfff955`.
The UI retained the 22x7 table, A2=120, B2 formula result 323, headers/blanks,
and rich text box while C2 became `Litchi formula-rich edited`. Only C2's
DataList object is the semantic matched IWA delta; formula list, Tile, AST
`=SUM(A2,203)`, cache 323, and dependencies are exact. This is no-impact
preservation, not native impacted-formula-refresh evidence; unsupported
impacted dependencies still reject as `FormulaCache`.

The host cut removes `NumbersEditor::{set_cell, set_cells, clear_cell}`,
Numbers-only `model::{set_cell_in_package, set_cells_in_package}` and
`TableCellBatch::apply_numbers`, 15 obsolete direct tests, and
`edit_numbers_cell`. It retains shared `TableCellBatch::{collect, is_empty,
len, apply_attached}`, attached/package writers, storage/wire/cache/formula,
Pages/Keynote/builders, and narrow test-only fixture adapters. Gates pass 237
Numbers library tests with 4 ignored, 15/15 public cell tests, 178/178 full
protobuf tests, and 196/196 boundary tests. Host library gates pass 1,422/1,422
overall, 382/382 Numbers subset, and 15/15 focused cells; all host examples
compile in the green all-target example gate after their selector-first
migration without restoring the retired API.
The neutral `litchi-numbers -> litchi-iwa-text-wire` edge leaves topology at
64 packages, 238 internal declarations, 14 `litchi-iwa` declarations, and 14
ordered debts.

## 2026-08-12 amendment: focused Keynote existing-slide deletion cutover

The secure cutover moves existing-slide deletion from the `litchi-iwa`
migration host to the selector-first `litchi-keynote` package transaction.
`slide::delete::{Edit, Patch, Commit, Diagnostics, Error, LimitKind, Path}` and
`Package::{edit_slide_deletion, apply_slide_deletion}` are the only supported
surface. Exact navigator names and checked semantic positions replace the
host's raw index-to-native-ID operation, and the final slide cannot be
deleted.

Changed admission proves the unique supported rooted
Document -> Show/SlideTree `[3, 2]` -> SlideNode `[2]` -> Slide path and exact
aggregate reference occurrence. `FieldInfo` attribution is optional, but if
present it must occur exactly once at the canonical path with object-reference
typing. The selected Node and Slide each have one supported message. A
package-wide scan rejects another inbound owner, duplicate reference
occurrence, selected IDs appearing as data references, noncanonical field
typing, or any merge/base/diff state. Deprecated root-node, secondary
slide-list, hierarchy, and other non-flat topologies are `UnsupportedTopology`;
malformed or contradictory evidence is `InvalidSource`; a surviving owner is
`AmbiguousOwnership`.

The PackageMetadata preflight is an exact transition proof, not a best-effort
registry cleanup. It streams the single type-11006 payload, requires unique
current component identifier/effective-locator matches, requires the exact
Node and Slide UUID bindings, and requires exactly one supported unversioned
Node-component-to-Slide-component edge: either an object-specific external
reference to the Slide object, with matching optional weakness, or a
component-level reference. Aggregate selected data references are
authoritative; optional field attribution must match them exactly, and every
reference must have its exact object owner/count record. Publication removes
precisely those two UUID bindings, the object-specific external reference when
present, and selected data-owner records. A component-level edge is retained.
It retains the last-object identifier, current component registrations, global
data-catalog records and payload bytes, unrelated current/versioned metadata,
and raw unknown fields. A component data-reference record is retained with
surviving owners or removed when none survive. Candidate scanning verifies
every requested record is absent and every selector still denotes the same
retained component.

The physical transaction removes the Show's selected Node reference and the
two selected objects, not their IWA components. Other objects co-located in a
selected component remain byte-exact. The exact root preview names are deleted
as stale derived rendering state, while case-distinct and nested preview-like
members remain. All unrelated ZIP members retain data and physical records,
and the rewritten members retain their package metadata. The candidate is
reassembled once, reopened once under the source limits, fully validated, and
must expose the exact navigator-name sequence with the selected position
removed. Forward apply reproduces the committed exact target; inverse apply
restores the exact source.

This is deliberately not media GC. PackageMetadata owner/count records for
the deleted objects are removed, and an ownerless component data-reference
record can disappear with them, but no component registration, global
data-catalog record, `Data/` member, or media payload is reclaimed. Shared,
uncertain, and newly unreachable media remains preserved. The earlier host
implementation's conditional orphan-media removal is not part of the focused
parity contract; reclamation requires a separate reachability owner and native
gate.

The PackageMetadata Buffa projection remains private and non-encoding. It is
five generated files / 145,681 bytes, has zero repeated generated views, and
has SHA-256
`ee49927f75c6b632c83055f9b7e647920b389be41bec10e25871a6ef7b56ab31`.
The removal API extends codec/provenance behavior without changing the schema
file. Raw records remain the rewrite and preservation authority. The focused
codec gate passes 11/11 and focused slide-deletion integration tests pass
10/10. The security matrix covers
missing/duplicate/wrong/versioned UUID and external-reference ownership,
duplicate/mismatched/versioned data-owner records, ambiguous identifiers and
component locators, aggregate/field disagreement, extra selected messages,
surviving inbound references, stale patches, exact preview-name matching, and
content-free public values.

Deterministic production instrumentation covers two independent scaling axes.
For 4,096 -> 8,192 package objects, objects scanned, transaction work,
allocation events, and peak scratch bytes must each remain at or below 2.20x.
For 4,096 -> 8,192 reference occurrences, object scans remain constant while
references, transaction work, allocation events, and peak scratch bytes must
remain at or below 2.20x. Each successful case performs one output allocation
and one candidate reopen and deletes zero components. A required-minus-one
work ceiling refuses before component deletion, output allocation, or reopen.
The object axis reports objects 4,096 -> 8,192 (2.0000x), references 54 -> 54,
work 20,460 -> 36,844 (1.8008x), allocations 13 -> 13, and peak scratch 240 ->
240 bytes. With eight objects fixed, the reference axis reports references
8,246 -> 16,438 (1.9934x), work 12,300 -> 20,492 (1.6660x), allocations 13 ->
13, and peak scratch 240 -> 240 bytes. Every successful case reports one
output allocation, one reopen, and zero component deletions. The maximum-work
success consumes 36,846; a ceiling of 36,845 refuses with observed 36,846 and
zero component deletion, output allocation, or reopen. These are deterministic
work/allocation counters, not a latency, throughput, or RSS claim; the
temporary instrumentation is not retained in the production surface.
The focused performance gate passes 4/4.

The final native oracle begins with a 511,554-byte source, SHA-256
`49c7ee349cddb9fcd4671b7cd36c90008a76e457311cd3bb70d4b765f217b3df`.
The 471,837-byte Rust candidate is
`aae1026b91f454abae3a35aac395f2f8e433e070d89a7825be649a29036e1cf5`;
focused reread reports two slides, 989 objects, and exact navigator order A/C,
while inverse application restores the exact source bytes. Members change
60 -> 57 by removing only the exact root previews. Of the retained members,
Document changes 8,710 -> 8,629 bytes, Metadata 32,690 -> 32,624, and the
selected Slide component 975 -> 810; the other 54 retained members are exact.

Keynote 14.4 build 7043.0.93 opened the Rust candidate without warning,
completed Save As and close, and reopened the exact output path without
warning. The 503,213-byte native resave is
`0c1a7750dba9c1bda43a287251096fe55a5ae05a5e7adddd07b69b79f2e64ea4`;
focused reread reports two slides, 975 objects, exact A/C navigator order, and
a clean ZIP. The 82,966-byte `ui-after-saveas.jpeg` screenshot is
`99c6d1b80ef93ddeec31c25ec2317f7dcf0bb0c9e2a597ca8207ffcb6adac702`,
and the 82,608-byte `ui-after-reopen.jpeg` screenshot is
`e807af871bef5addbd28ecd5f24765c619ae33c3707db5029a6dc6de9e4048cb`;
both show the exact two-slide A/C navigator and selected C body/title.

The host retirement deletes `KeynoteEditor::remove_slide`, its complete
`keynote/editor/slide_delete.rs` module/source, the direct
`remove_keynote_slide` example, and obsolete direct host deletion tests. The
retained generated-presentation regression is creation-only. Its generated
slide has a child-object-to-parent-slide backlink, so focused deletion refuses
that topology as `AmbiguousOwnership`; it is not evidence that all
host-generated slides are deletable. No compatibility method or public bridge
alias replaces the retired host method. The boundary suite passes 204/204.
The `litchi-keynote` all-features gate passes 235/235 tests: 104 library tests
plus 131 tests across the integration targets. Its doctests pass 9/9. The
focused and retired-surface audits each report zero findings, while the full
boundary checker reports exactly the 14 established unrelated findings. The
full retained `litchi-iwa` library gate passes 1,418/1,418. Its permanent
`focused_slide_deletion_rejects_generated_child_backlinks_atomically`
regression proves the generated object-111-to-slide-110 backlink, the typed
`AmbiguousOwnership` refusal, and byte-exact source preservation.

These frozen results close the focused existing-slide deletion cut. They do
not close debt 014 or claim that the full checker is free of its established
unrelated findings.

Debt 014 (`litchi-iwa -> litchi-keynote`) and the manifest edge remain because
broader Keynote creation, drawable, chart, table, media, soundtrack-item,
example/test/fuzz, durable-patch, and atomic-save ownership still lives in or
depends on the migration host. Final topology is
64 workspace packages, 238 internal dependency declarations, 14 `litchi-iwa`
dependency declarations, and 14 explicit ordered debt items.

## 2026-08-12 amendment: Numbers formula-cache foundation verification

Internal cache-planner regressions prove byte-exact preservation of unrelated
cycle markers, refusal when an impacted marked formula survives the final
same-batch overlay, and success when that overlay removes the marked formula.
Graph work has exact max-minus-one refusal coverage before publication;
scratch and allocation remain bounded by the planner limits.

These are foundation gates only. No focused public formula authoring is
implemented or verified; the production host setters remain. No native
Numbers formula-authoring run, candidate artifact, performance measurement,
host retirement, dependency-edge removal, or debt closure is claimed.

## 2026-08-13 amendment: Pages section-background verification

The focused Pages background codec takes the strict raw-preserving route
TP.SectionArchive field 30 -> Fill fields 1/2/3 -> Color model field 1, RGB
fields 3--6, and color-space field 12. A private read-only Buffa lazy view is
an independent parity oracle after strict preflight; handwritten routing owns
preservation and rewrite. It retains no unknown/repeated generated state, uses
no generated production encoder, and refuses duplicate or wrong-wire field 30,
malformed nesting, noncanonical/invalid semantic color state, and strict/Buffa
disagreement.

The projection source SHA-256 is
`0a6f03a7046c285e431953b8752096a1f0117206724b561da294c64092aa9cfc`.
Deterministic generation produces five files totaling 99,593 bytes, with zero
`RepeatedView` or `LazyRepeatedView` mentions, aggregate SHA-256
`9abd261dfe79866b0718411e0da75e1001a1eeeda50770037400c9e309cbb9ca`.
Codec tests pass 8/8 and focused integration passes 8/8. The integration gate
covers absent/solid/unsupported read classification, selector failures,
malformed field refusal, exact no-op identity, set/clear, apply/inverse and
conflict, nested unknown preservation, selected-component locality, changed
legacy refusal, output-limit atomicity, and unsupported/ambiguous ownership
refusal.

The native gate starts from the Apple-authored 91,681-byte
`/private/tmp/litchi-pages-section-settings-native.aEp44s/section-background-solid.pages`,
SHA-256
`5d5795c9de521e54eb5e5986241ca752ec4e87d076bc9171c67ef4a281bedc8c`.
Focused CLI replacement produced a 91,695-byte dark-red sRGB candidate,
`d5b7605b2f24de197b9e29fc79f25e44f1a5d34a15c82528d72baea74d5d6118`,
and clear produced a 91,671-byte `None` candidate,
`c278fcb2a7504824385fb29fd33f74d5fe26a3030b47eab62ed742f13a6037f4`.
Each inverse was byte-exact to its source. Pages 14.4.1
(`M14.4-7043.0.93-4`) opened both without repair or conversion; Accessibility
identified page 1 replacement as Section Background `Color Fill`, color
`dark red 34`, and clear as `No Fill`. After Save, close, and exact-path
reopen, the same states persisted. The native-resaved ZIPs are valid:
replacement is 92,160 bytes / SHA-256
`8db386ebcba32086afa3b37ce1d2617c6360677814462ad49085558c949492eb`,
and clear is 89,121 bytes / SHA-256
`b097877d83ca6832285956be4cee15f9680fc90d046eb48acf4aebafbeba8d1b`.

Focused reread of each Pages-resaved artifact reports the corresponding
solid-to-solid or none-to-none exact no-op, zero touched components, and zero
deleted previews; its no-op and inverse remain byte-exact. Pages can otherwise
rewrite its own package on resave, so this gate makes no claim about native
resave member locality or byte preservation.

## 2026-08-13 amendment: legacy Keynote reader retirement gate

The retirement deletes the 933-line
`crates/litchi-iwa/src/keynote/document.rs`, its module declaration, and the
`KeynoteDocument` re-export. Source audit finds no retained public
`KeynoteDocument` or `KeynoteDocumentStats`. The duplicate reader's private
`Bundle`, `ObjectIndex`, semantic `OnceLock`, eager-Prost show/slide graph, and
wide text extractor disappear with the file; editor and builder code remain.

API audit maps every supported semantic capability to
`litchi_keynote::Document`: `open` and `open_with_options` accept complete ZIPs
or frozen app-authored package directories and eagerly return an archive-free
full show, rooted text, source-derived metadata, and source statistics. The
metadata begins with semantic Show values and incorporates narrowly decoded
canonical-properties scalars when that diagnostic exists; `Some` is not a
sidecar-presence signal. Cheap snapshots, slides, validation, and those
accessors remain on the detached value. Exact regular-file artifacts map
separately to `litchi_keynote::Package`, including `from_bytes`, semantic
projection, cheap shared `semantic_snapshot`, exact `write_to`, and edit
provenance. A package-derived semantic snapshot is intentionally
diagnostic-free: `metadata()` and `stats()` return `None`. The redundant
`from_archive_bytes` alias maps to `Package::from_bytes`. The constant
`KeynoteDocumentStats.application` field is intentionally omitted. Checked
physical and semantic limits and prepared-source ingress are additional
focused capabilities, not compatibility obligations.

Path parity is intentionally split by provenance. `Document` captures either
path shape through `PreparedSource` and completes semantic decoding before it
publishes the snapshot. The cross-format coordinator can delegate through that
same frozen-source boundary. `Package::open` requires the complete regular-file
ZIP because exact `write_to` and mutation patches bind to the full physical
artifact; it refuses directories rather than treating directory `Index.zip` or
loose `Index/` components as a complete writable `.key`. An archive-free
directory snapshot makes no promise to preserve other sidecars, `Data/`,
previews, or complete package bytes.

Parity is verified at the public behavior boundary, not by comparing legacy
and focused object layouts. The focused reader deliberately corrects three
legacy behaviors: text is limited to storages reachable from the rooted show
graph; rich storage fragments are preserved instead of flattening body/date
content; and plist metadata plus package validation are richer and stricter.
Accordingly no claim is made that legacy and focused Show values, text vectors,
or error classifications are structurally identical.

Metadata lookup authorizes the exact canonical logical
`Metadata/Properties.plist` path. The hostile near-name regression proves that
an unrelated entry sharing `Properties.plist` as its basename is ignored.
Catalog-normalized legacy nested-ZIP wrapper prefixes retain their logical
path, while arbitrary flat wrapper prefixes are not accepted as canonical. A
centralized 64 KiB hard ceiling admits the canonical diagnostic independently
of the broader source-entry limit, and only the scalar fields projected into
public metadata are decoded.

This cut removes the duplicate host eager-Prost path, not every Prost use in
the focused crate. Six generated-message decodes remain in semantic traversal
after bounded wire preflight; they do not form a second public reader.

The permanent `generated_roundtrip` gate opens a builder-generated flat `.key`
with `litchi_keynote::Package` and proves `Send + Sync`, snapshot/stats
agreement, validation, show and semantic-snapshot slide counts, and nonempty
rooted text.

The permanent path regressions prove packaged/directory Keynote semantic parity
both through focused `Document` and through the cross-format coordinator. They
match directory snapshots against the focused ZIP package for slide count,
rooted text, position, skip state, name, title, builds, and transition presence.
Boundary ratchets audit removal of the public reader and its duplicate
implementation, while the full checker continues to distinguish the existing
dependency-policy baseline.

The frozen ingress and semantic gates pass: archive-directory 16/16,
detection 18/18, focused Keynote native 7/7, coordinator `iwork_path` 7/7, and
metadata scalar/64 KiB-cap unit coverage 1/1.

This is a read-only retirement, so it requires no new native mutation or Save
oracle. The fixed Apple-authored read fixture is
`test-data/iwork/keynote/basic.key`, 500,058 bytes, SHA-256
`3a3d07476b45b6e543bcfba75fe38a245434176dcb3565e34570b817708b9f42`.
Its existing focused native reader coverage remains the read-only oracle; this
cut does not rewrite that fixture or claim a new mutation artifact. Keynote
14.4 build 7043.0.93 opened an isolated copy without repair, recovery, or
conversion warning. The navigator showed its single
`Litchi native Keynote fixture` slide; the canvas showed the expected title,
body marker, and `2026-08-07`. Separately, the automated focused native-fixture
gate covers deterministic package behavior, typed `Package` directory refusal,
focused `Document` ZIP/directory semantic reads, canonical metadata lookup with
a hostile near-name, and focused reread of one slide and 959 objects. Metadata
isolation also preserves the source exactly; these tests do not automate the
Keynote UI. Command-W produced no save prompt, but Keynote silently
autosave-normalized only the disposable copy to 500,011 bytes,
SHA-256
`5a6c5b260a3e3b6d77e848d3198a5d41b74fe9f1e9f9fd7d1a8050f9b4092427`;
focused semantics remain identical and ZIP integrity passes. The checked-in
source retained its size, hash, mtime, and inode, so this is read compatibility
evidence rather than byte-preservation evidence for native open.

`KeynoteEditor` and `KeynoteDocumentBuilder` remain. Debt 014 and the
`litchi-iwa -> litchi-keynote` manifest edge therefore remain open.

## 2026-08-13 amendment: legacy Pages reader retirement gate

The retirement deletes the 478-line
`crates/litchi-iwa/src/pages/document.rs`, its module declaration and local
re-export, and the `PagesDocument`, private `PagesDocumentState`, and
`PagesDocumentStats` types. The sole real host caller migrates to the focused
owner; stale host documentation points semantic readers to `litchi_pages`.
The permanent boundary audit rejects resurrection of the exact source, module,
re-export, types, callers, examples, and documentation while allowing
`PagesDocumentBuilder`, `PagesEditor`, and focused `Document`/`Package` use.
The live retirement audit reports zero findings.

Public capability mapping is provenance-specific. Semantic reads use
`litchi_pages::Document` for a ZIP path or checked app-authored directory on
supported path-ingress platforms, or borrowed ZIP bytes or shared ZIP bytes on
every supported platform, and obtain eager archive-free sections, text,
source-derived metadata/statistics, snapshots, and semantic validation. Exact
regular-file or byte reads use `litchi_pages::Package`; that owner retains
bytes, package metadata/statistics, physical validation, and editing,
and refuses directories. Package-derived and constructed semantic documents
return no source diagnostics. The retired archive-byte name maps to
`Document::from_bytes` for semantic reads or `Package::from_archive_bytes` when
exact artifact provenance is required.

Parity is asserted at supported behavior, with the focused Pages contract
authoritative for structure. Empty roots have zero sections. Rootless fallback
uses the retired 14-type object trigger, fully validates and projects every
registry storage message in that object, and newline-joins fragments in source
order, including empty fragments, before enforcing the aggregate retained-text
ceiling. Rooted bodies preserve rich storage runs and project the native section table, exact name
presence, and UTF-16 boundaries. Metadata parity uses only the exact canonical
Properties, BuildVersionHistory, and DocumentIdentifier paths, retaining at
most 64 KiB from each in the semantic handoff. Near-name sidecars do not
participate. ZIP and directory semantic reads agree, while only a complete ZIP
package carries exact-artifact provenance.

This cut removes the duplicate host's eager Prost root/body reader, not all
Prost use in focused Pages. Strict raw-wire preflight and private forced Buffa
lazy views select the root references and section boundaries. The rooted body
is then materialized by one bounded `TSWP.StorageArchive` Prost decode;
rootless fallback candidates instead pass full known-field raw storage
validation before the bounded Buffa text projection and object-level
coalescing. Raw records remain the preservation authority for mutations. The focused
reader is eager rather than lazy or single-flight, and no latency, RSS,
allocation, read-scaling, or complete Buffa-laziness claim follows.

The default read envelope is 1 GiB input, 100,000 entries, 512 MiB for one
expanded entry, 2 GiB aggregate expanded bytes, 512 MiB for one decoded IWA
component, 4,096 semantic sections or fallback storages, and 64 MiB retained
section-name plus text bytes. Caller options may tighten these ceilings. The
archive-free snapshot retains at most 64 KiB from each selected authority.
Directory capture enforces that ceiling before allocating a sidecar and charges
it to the source budgets. Packaged ZIP path, borrowed-byte, and shared-byte
capture first checks each exact raw logical authority's declared uncompressed
size and compression method from the physical ZIP headers, rejecting more than
64 KiB or unsupported compression before any package entry payload is
materialized. Only the selected legacy outer-package prefix is stripped before
that raw-byte comparison; near-names remain unrelated. Local and central ZIP
names and methods must also agree.

The permanent focused reader integration gate passes 15/15. It covers
ZIP/directory/borrowed-byte/shared-`Arc<[u8]>` parity, release of a caller's
shared source allocation after projection, `Send + Sync`, source-diagnostic
presence versus diagnostic-free package semantic documents, shared-section
snapshot identity, typed package-directory refusal, and typed Keynote-as-Pages
refusal for ZIP and directory inputs. It also covers all three canonical
metadata authorities with hostile near-names; survival after the captured
directory is deleted; exact and max-minus-one source and semantic budgets; the
64 KiB properties boundary; multi-fragment, co-located-message, empty-fragment,
and bogus-trigger fallback cases; atomic malformed-root refusal; content-free
path, member, and control-character error redaction; and checked semantic-limit
construction. The full `litchi-pages` all-feature gate passes 153/153: 77
library tests, 15 document-reader tests, 59 other integration tests, and two
doctests. Supporting archive tests pass 93/93, detector tests pass 32/32, and
the host generated-roundtrip gate passes 1/1. The boundary suite passes
227/227; its live retirement and focused-public-API audits each report zero
findings. All-target check, strict all-target Clippy, strict no-dependency
rustdoc, global formatting, and diff checks pass.

One materialization limitation and one platform capability caveat remain
explicit. Semantic ZIP ingress still builds the bounded package catalog and
may expand unrelated supported entries under the generic source limits before
discarding them; the selected metadata preflight is not a general ZIP member
filter. Windows Pages file and directory path ingress deliberately fails
closed because stable, reparse-safe source identity is not yet available;
borrowed and shared byte ingress remains supported there. Public archive-free
opens now return `ReadError`, whose display, debug, and source chain expose only
closed categories and numeric bounds rather than paths, member/component names,
native identifiers, content, or lower-layer diagnostic strings.

The read-only native oracle keeps the tracked
`test-data/iwork/pages/basic.pages` unopened and exact at 96,417 bytes and
SHA-256
`21107bc9323fba6f1589152454c0b0b0cc8e239313c6a369bc4a891116601b42`.
Apple Pages 14.4 build 7043.0.93 opened only an isolated copy without repair,
recovery, or conversion UI and showed exactly one page with
`Litchi native Pages fixture`, `Buffa lazy-view migration verification`, and
`2026-08-07`. Focused reread reports 570 objects, one Body section named
`Blank`, exact text, and Pages 14.4.1 metadata. This is compatibility evidence,
not byte-inert-open evidence: Pages silently normalized the disposable copy to
96,432 bytes and SHA-256
`665e4f6f26713d14a2346b129b0e19ea6cc83ffefe1d8244866b51fe6a79e127`,
while retaining a valid 13-member ZIP and the same focused semantics.

`PagesEditor`, `PagesDocumentBuilder`, creation, and the broader host examples
and tests remain. Debt 017 and the `litchi-iwa -> litchi-pages` manifest edge
therefore remain open.

## 2026-08-13 amendment: Numbers reader retirement verification

The migration host's duplicate Numbers reader is deleted together with its
module/export/state/statistics types and reader-only sheet adapter. Host
callers and examples now use `litchi_numbers::Document` for archive-free
semantic reads or `litchi_numbers::Package` when exact artifact, package text,
write, or edit provenance is required. Permanent boundary audits reject the
retired names, module shapes, aliases, public host facades over focused reader
types, raw/native vocabulary in focused reader signatures, and reintroduction
of a second reader source.

The final permanent document-reader gate must cover ZIP and app-authored
directory semantic-workbook parity, borrowed and shared bytes, shared-source
release after projection, snapshot and shared-sheet identity, `Send + Sync`,
canonical three-sidecar metadata projection, source diagnostics versus
diagnostic-free semantic and package-derived documents, rooted plain-text
ordering,
exact/max-minus-one source and semantic budgets, foreign iWork refusal,
malformed graph refusal, directory lifetime after capture, and content-free
error redaction. It must prove that generic/foreign source preparation remains
index-only until Numbers ownership is established, so hostile metadata cannot
change another format's classification. Directory coverage must also prove
that only the exact root `Index.zip` authority participates and root or nested
decoys are inert. Windows-specific regressions must freeze file and directory
path ingress as a typed fail-closed result while retaining borrowed/shared byte
ingress. Unix path tests must cover pinned, no-follow capture; other
non-Windows targets are qualified as version-checked path capture rather than
descriptor pinning.
Semantic-workbook parity is not metadata identity: the checked ZIP and
directory fixtures carry different canonical document identifiers and
revisions, which remain source-provenance diagnostics rather than values to
normalize across representations.
The plain-text oracle fixes the order as non-empty sheet name, then each
non-empty table name and non-empty materialized cell display in row-major
order, with exactly one newline between emitted values. Empty rendered
Text/Formula values are excluded rather than producing blank lines.
The checked-in native fixture also freezes a deliberate correction: legacy
`NumbersDocument` construction rejects that valid workbook before its public
`text()` can run. Independently recovered private legacy storage output equals
`Package::text` on this fixture and the formula-rich fixture, but that narrow
observation is not promoted into general parity.
The focused native semantic oracle is exactly
`Sheet 1\nTable 1\nLitchi native Numbers fixture\n42`; `text_len()` equals
that output's UTF-8 length. The source-backed statistics are exactly 622 source
records, one sheet, and one table.
The final metadata-hostile gate must fix 64 KiB per-authority physical
admission and the `plist::stream` event projector's exact/max-minus-one event,
depth, history-entry, selected-scalar, and retained-property budgets. The
projector must remain narrow rather than deserializing a general scalar DTO or
plist value tree.
The admission claim is publication-scoped: selected projections preflight
where implemented and all remaining schema decodes stay within physical,
component, and semantic ceilings, but not every ceiling is claimed to precede
every decode or intermediate allocation. These gates qualify focused
`Document` construction, not the preserve/edit `Package` path.
Caller-selected semantic zeroes remain exact zero ceilings rather than being
widened to one, and over-hard requests fail through `DocumentLimitsError`.
Owned attacker-scaled payload and collection growth is fallibly reserved where
the focused adapters control it. Standard-library `Arc` control blocks and
final immutable publication allocations remain an explicit allocator-abort
caveat; this is not an all-allocations-fallible claim.

The frozen verification record is explicit rather than inferred from one
monolithic run. Focused reader coverage passes 16/16, with a seventeenth
Windows-configured case. The Numbers library passes 240 cases with four
ignored; compatibility passes 5/5 and names pass 10/10. Archive coverage passes
127 cases (125 unit plus two integration), and detector coverage passes 40/40.
The host library passes 1,397/1,397, generated-roundtrip passes 1/1, and host
doctests pass nine with three ignored. Host all-target check and all-target
no-run pass, as do strict scoped
host Clippy (`--lib --test generated_roundtrip -D warnings`), focused all-target
Clippy, strict focused rustdoc, formatting, and diff checks. Boundary units pass
237/237; the live retirement and focused-public-API audits each report zero
findings. The host-scoped cut touches 15 files with 329 insertions and 888
deletions, net -559, including the 602 reader-owned source lines. Broad host
all-target Clippy remains blocked by unrelated existing lints, and the global
boundary policy continues to report 14 unrelated
`soapberry-zip`/`xml-minifier` debt findings.

Native evidence is read-only and isolated. The tracked
`test-data/iwork/numbers/basic.numbers` source was never opened and remains
136,357 bytes, a valid 43-entry ZIP, with SHA-256
`f225d5b1cd59e9da454f91a96fe8f81154bc31037c10029230e75d49b45fb693`.
Apple Numbers 14.4 build 7043.0.93 opened only the disposable copy in
`/tmp/litchi-numbers-native.DfdNPz` without repair, recovery, or conversion UI.
The UI showed exactly `Sheet 1` / `Table 1`, 22 rows by 7 columns, one header
row and column, no footer rows, title enabled, caption disabled, the visible
`Litchi native Numbers fixture` marker, and numeric `42`. Focused reread agreed:
one rooted sheet, one 22-by-7 table, two materialized cells, and one
compatibility table.

Escape followed by Command-W produced no save prompt, but Numbers silently
normalized the disposable copy to 136,374 bytes and SHA-256
`b2388ce97cc30dbb1fadb02eece6f92fbeeeecb3e1a258aa79ece7511dfb31d6`.
It remained a valid 43-entry ZIP. This is application acceptance and semantic
agreement, not evidence that native open is byte-inert, and the normalized
copy is not a package-locality or preservation oracle.

The cut makes no performance or complete Buffa claim. Focused document
construction is still eager and substantially Prost-backed. Verified gains
are the deletion of duplicate parsing, constant-time shared snapshots, bounded
sparse semantic retention, and release of source/package state after semantic
projection. The Numbers editor, builder, host table adapters, examples, tests,
manifest edge, and ordered debt 015 remain.
