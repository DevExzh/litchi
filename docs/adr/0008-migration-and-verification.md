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
