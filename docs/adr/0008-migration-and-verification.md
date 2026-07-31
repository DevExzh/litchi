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
validation in hot loops. Lookup returns `Result<Option<&Cell>>`: absence is
distinct from an explicitly stored `Cell::Empty`, and no `Index` implementation
converts a miss into a panic.

Worksheet payloads load on first use into hidden thread-safe snapshot caches.
The parser streams into a row-major compact sparse slice, resolves shared
strings through cheaply cloned immutable text, expands shared-formula storage
records, retains exact numeric lexical forms, and keeps formula cache origin and
freshness separate. Sparse `cells(range)` traversal and stored extent are now
implemented; declared/content/formatted extents, rich-text formatting, merge
coverage, dynamic-array spill states, shared-style handles, dense budgeted
grids, editable transactions, cache eviction, and operation budgets remain
open, as does replacing remaining parser `Invalid` messages with the full
structured context taxonomy. The current non-evicting cache is therefore a safe
migration step, not the weighted-cache design promised by ADR 0005.

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
not claim byte-identical ZIP containers. Cell
edits remove an existing calculation-chain relationship and part, set
calculation properties for a full refresh, and retain the removed graph in the
inverse patch. The boolean spellings used for calculation properties follow
the Office compatibility notes in `[MS-OE376]` §2.1.599. This is intentionally
not yet the format-independent deterministic patch wire representation required
by ADR 0003.

This slice does not yet update the worksheet `dimension` hint or implement row
and column insertion/deletion, shifting references, merge/group-formula edits,
validation evaluation, style/resource editing, rich text, dynamic arrays,
patch serialization, full structured diagnostics, eviction/resource budgets,
concurrent subedit merging, or the container-level raw-copy/atomic-replacement
save pipeline. Those remain certification work; no allocation, latency,
contention, or scaling conclusion follows from the functional tests.

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

- Stable Rust with workspace MSRV 1.85 initially; bumps require a concrete
  safety, ergonomic, or measured-performance reason.
- Windows, macOS, and Linux CI; WASM where the selected I/O/crypto stack permits.
  `no_std` is not an initial support promise.
- Unit, integration, property, fuzz, Miri, sanitizer, malformed-corpus, and
  dependency-direction checks appropriate to each layer.
- Representative performance and resource budgets as defined by ADR 0005.
- Generated low-level schemas/records are deterministic, checked in, reviewed,
  and cite the source specification. Ergonomic facades remain handwritten.
