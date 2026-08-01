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
It preserves an existing regular file's permissions, refuses symbolic-link and
non-file destinations, cleans up a failed temporary artifact, and synchronizes
the parent directory where the platform supports it. Tests inject a failure
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
