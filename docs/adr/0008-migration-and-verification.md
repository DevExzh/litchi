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
diffs fail. iWork formats apply the same levels with Apple Keynote, Numbers,
and Pages as the native applications; a native resave is a compatibility
check, not an exact-locality oracle.

## Certification gate

[`docs/CRUD_Scenario_Checklist.md`](../CRUD_Scenario_Checklist.md) is the
support-certification checklist. A format's feature matrix marks a direction
supported only when every applicable checklist row is backed by tests or other
reproducible checks; an `N/A` row needs a format- or feature-specific reason.

The unified operations layer adds three rows per format, asserted by the
`litchi-ops` CRUD matrix test and replayed by the Python mirror:

- **atomic failure**: a batch whose item fails is refused naming that item and
  leaves the snapshot, its file, and refused save destinations byte-identical;
- **patch inverse**: a commit's patch replays onto its source, its inverse
  restores the exact source bytes, and the patch is refused on any other
  artifact; and
- **composition**: two edits made concurrently from one snapshot compose in
  either order to one result holding both changes, and each edit's patch is
  refused on the other's result.

## Where results live

ADRs record decisions, contracts, and limits; feature matrices record support
status. Neither records test counts, run results, artifact hashes, native-app
session notes, or other verification evidence. Such evidence is reproduced by
the tests and examples in the repository and otherwise kept outside it. Another
ADR states which checks a capability needs, never that they passed.
