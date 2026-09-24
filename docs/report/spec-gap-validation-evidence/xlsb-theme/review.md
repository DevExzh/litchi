# XLSB Theme review

Status: **CLEAR for the documented bounded Theme projection, with one
non-blocking caller-limit caveat recorded below**. Package graph,
conformance, and owner inverse paths are source-exact in the reviewed tests.
This is a read-only review; no production source was changed by the reviewer.

## Normative basis

MS-XLSB §2.1.7.52 (`3rdparty/specs/[MS-XLSB]/2 Structures/2.1 File
Structure.md:1291-1293`) delegates the Theme part to ISO/IEC 29500-1 §14.2.7
and its content to §20.1.4. The traceable vendored source is
`3rdparty/specs/ECMA-376/ECMA-376-1_5th_edition_december_2016.zip`, member
`Ecma Office Open XML Part 1 - Fundamentals And Markup Language Reference.pdf`:
printed pp. 135-136 (§14.2.7 Theme), p. 157 (§15.2.14 Image), and p. 4034
(Annex A `CT_OfficeStyleSheet`). Section 14.2.7 requires zero or one Theme
part for a SpreadsheetML/WordprocessingML package, owned by the Workbook/Main
Document implicit relationship, with an internal target. Theme may have only
Image outgoing relationships. Section 15.2.14 permits an Image target to be
internal or external; an internal image ZIP item needs an appropriate image
content type and an Image part has no ECMA-defined outgoing relationships.

Relevant project contracts are ADR 0003 (immutable source snapshots,
source-checked atomic patches and exact no-op preservation), ADR 0005
(bounded preflight/fallible allocation), and ADR 0024 (nested owner
responsibilities and clone-staged publication).

## Current candidate

The tree now has `crates/litchi-xlsb/src/theme.rs`, eager
`Workbook::theme/edit_theme/apply_theme` methods, an owner-level
`create/replace/remove` transaction, a source-backed read view, and shared
DrawingML scheme-range validation. The latest candidate also added
package-wide Theme ownership checks, image target checks, relationship/root
conformance, arbitrary-prefix Strict splicing, optional root-name insertion,
and checked output-size preflight.

The current tree resolves the earlier graph, allocation, and owner inverse
blockers:

* eager and source-backed readers count Theme content-type parts, reject
  orphan/multiple/non-Workbook inbound ownership, and require one internal
  Workbook edge;
* internal Theme Image targets require `image/*` content type and no outgoing
  relationships, while external Image targets remain inert and are not read;
* the Workbook Theme relationship and root DrawingML namespace must use the
  same Strict/Transitional profile, and Theme Image relationship URIs are
  checked against that profile;
* scheme edits use the shared namespace-aware range finder, preserve arbitrary
  prefixes and user values, preflight all replacement ranges/output size, and
  retain an omitted root `name` until explicitly set.
* owner create/remove/inverse retains image payload parts, workbook and Theme
  relationship tokens (including explicit empty `.rels` members), and the
  content-types token. The focused owner test compares serialized package bytes
  after remove followed by the inverse patch.

## Findings and disposition

1. **Resolved: full-document DrawingML profile check.** The eager and
   source-backed package readers now parse under the caller byte ceiling before
   conformance, then scan every bound Strict/Transitional DrawingML namespace
   against the Workbook relationship profile. A mixed descendant is refused;
   the package-level Image relationship URI check is profile-specific as well.

2. **The typed model is a bounded projection, not full Theme metadata.** It
   does not expose `fmtScheme`, object/default lists, extra/custom color
   schemes, extensions, color choices/transforms (the valid system-color
   token set is now complete, including `3dDkShadow`/`3dLight`), or all font
   attributes. Source bytes are retained and replacement refuses unsupported
   selected-scheme markup. The feature matrix explicitly scopes this to a
   source-preserving projection; it must not be advertised as complete Theme
   schema validation. In particular, reader projection may accept a source
   with an omitted `fmtScheme` or reordered required schemes because those
   fields are outside the typed projection; this is a bounded preservation
   policy, not an assertion that the source is XSD-valid. Generated/new Theme
   XML and edited supported fragments are separately checked against the
   vendored schema fixtures. The shared parser also refuses valid color choices
   outside its model.

3. **Resolved: selected-scheme namespace handling.** The shared scheme-range
   validator now requires namespace declarations in a selected scheme to bind
   the source DrawingML profile and refuses foreign/cross-family declarations
   before replacement. Unknown content outside the selected schemes remains
   source-preserved.

4. **Resolved: namespace/byte preflight.** `source_namespace` now applies the
   8 MiB hard byte guard and shared 256-declaration resolver ceiling; eager and
   deferred readers perform the caller `max_xml_bytes` check before invoking
   the conformance probe.

5. **Documented caller-limit caveat: MCE preprocessing.** `parse_theme` checks the
   raw source against `Limits.max_xml_bytes`, then `codec::read` runs
   `process_ooxml` using its own fixed hard output limit. A markup-compatibility
   source can therefore have a processed representation larger than the
   caller's configured limit. The shared codec still applies its finite 8 MiB
   processed Theme cap (and the MCE processor has its own finite global policy),
   so this does not create an unbounded path; it means a caller-selected limit
   below the hard cap is a raw-source ceiling rather than a complete MCE
   intermediate ceiling. A future API can thread the caller limit into MCE if
   strict per-request accounting is required.

The package graph tests now cover duplicate Workbook Theme edges, wrong Theme
content type, external Theme edge, unsupported outgoing edge, orphan/second
Theme/non-Workbook inbound edges, invalid/non-leaf image targets, writer
creation, strict/transitional relationship conformance, external and internal
Image targets, occupied candidate names, and exact owner remove/inverse
serialization including relationship/content-type bytes and image resources.
The latest integration run adds a mixed-descendant Strict refusal regression,
namespace/byte-boundary checks, and selected-scheme declaration rejection. An
external-image preservation fixture is covered by the owner inverse test; the
external payload remains inert.

## Evidence and checks

`python3 docs/report/spec-gap-validation-evidence/xlsb-theme/verify-theme-schema.py`
validates both native Transitional XLSB Theme parts:

* `test-data/poi/test-data/spreadsheet/testVarious.xlsb`, 8,390 bytes,
  SHA-256 `50b662d8ff0e562157ff80aa19c85ff214cd21357d66457661d070f50c6fdc59`;
* `test-data/ooxml/xlsb/62815.xlsb`, 7,646 bytes,
  SHA-256 `af2fb29aa44e18757cbc03afd5c6ddf60deef78419009b6d07367fff46136fb1`.

The durable generated Transitional fixture
`docs/report/spec-gap-validation-evidence/xlsb-theme/authored-theme.xml` is
2,129 bytes, SHA-256
`d51f6f8b5ddc9fca6173bebcac93399770bb41a6edd586d8090db65bcd143d0c`; its
`authored-schema-validation.json` records validation against the vendored
schema archive SHA-256 `d34187520749998af306faf1b730e568b0ca6d88ad24638a407c0a9bb4ca04fc`
and main schema SHA-256
`6978ba7e889070b0c3cb5b546b23e5a6c3516134afc53b87a21f482ca33f3858`.

The current focused evidence supplied by the integration run reports:

* `cargo test -p litchi-xlsb --test theme --offline --locked --no-fail-fast`
  (16/16, including lexical relationship staleness, new-owner inverse, and
  absent-`.rels`/no-op cases);
* `cargo test -p litchi-xlsb --test theme_source --offline --locked
  --no-fail-fast` (7/7);
* shared DrawingML theme tests (8), shared namespace/declaration regressions
  (2), PPTX theme tests (4), and the latest splice regressions passed in the
  pinned temporary target.

Earlier local rebuilds did encounter `No space left on device` while compiling
the large workspace target; the focused receipts above were run in the pinned
temporary target. No reviewer production source edits were made.
