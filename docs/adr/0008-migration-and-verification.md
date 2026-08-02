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
current package graph. The new XLSB `XlsbWorkbook::{task_panes,
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
requires one exact 26-byte payload, while `write` streams those fields through
the canonical raw writer without allocating an intermediate record buffer.
The parsed workbook exposes `calc`; the authoring host uses `calc`, `calc_mut`,
and move-accepting `put_calc`.

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
passes two unit tests, seven calculation integration tests, 17 existing
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
safe. `PresentationProtection` owns an optional private `ModifyVerifier`
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

The worksheet parser also matches checked-in Apache POI and LibreOffice shared-
formula fixtures, including translated follower expressions and stored cached
results. Synthetic tests cover missing versus explicit empty cells, grid-bound
rejection, malformed shared-formula groups, exact numeric lexemes, sparse range
order, read-only serialization stability, and concurrent first access. These
are read-path regression gates, not performance evidence; allocation, latency,
contention, and scaling claims still require the measurement work in ADR 0005.

## Quality gates

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
