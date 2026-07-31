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
