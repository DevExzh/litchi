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

- `litchi-odf-common` owns ODF constants, coordinates, and datatype
  vocabulary; `litchi-odf` retains detection, package orchestration, and
  family-specific semantics while re-exporting the established paths.
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
claim. ADR 0009 continues to keep ODF detection in `litchi-odf`, and ADR 0010
continues to keep archive grammar below the public facade.

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
every affected crate except the pre-existing RTF strict-lint backlog. The
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
matrix, and strict checks pass for every affected crate except the known
pre-existing RTF workspace-lint backlog. Library tests pass for DOC (832 with
two ignored), DOCX (643), DrawingML (92), IWA (1,529), ODraw (59), ODP (103),
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

The strict all-target matrix passes for every affected crate except the known
pre-existing RTF workspace-lint backlog. The lint-capped all-target matrix
passes for DOC, DOCX, DrawingML, IWA, ODS, ODT, PPTX, RTF, XLSB, and XLSX. The
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
