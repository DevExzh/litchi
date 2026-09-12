# ODT flat text-template (`.fott`) design evidence

**Status:** bounded design only; no source or test implementation is included in
this report. The audit was checked on 2026-09-12 at `d000d977b`.

**Scope:** the single-XML OpenDocument Text template MIME
`application/vnd.oasis.opendocument.text-template`, exposed through the existing
`litchi-odt` flat APIs. The `.fott` suffix is a project filename convention for
that MIME value. It is not a filename requirement in ODF 1.4.

## Finding

The audit row is current, not stale. `crates/litchi-odt/src/generic/flat.rs:60-78`
already obtains and validates the flat root MIME, and
`crates/litchi-odt/src/generic/codec.rs:118-140` already classifies the text
template as `(Family::Text, true)`. The refusal is the explicit
`template || ...` guard in `generic/flat.rs:65-68`. `FlatDocument` has no
template bit, whereas the packaged `Package` model already stores one and
exposes `is_template()`.

The common detector needs no new MIME rule. `litchi-odf-common::detect::mime`
maps the document and template text MIME values to `Format::Odt`, and
`flat_mime` reads a namespace-resolved `office:mimetype` from an
`office:document` root. The neutral detector intentionally does not expose the
template distinction; the ODT owner must retain it after admission.

This is therefore an API and regression-fixture gap, not a format impossibility.
The package `.ott` path is already supported and must remain unchanged.

## Normative evidence

The repository's primary ODF archive is
`3rdparty/specs/OpenDocument-v1.4-os.zip`, SHA-256
`9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4`.
The relevant archive members and hashes are:

| Local primary source | SHA-256 | Rule used by this design |
| --- | --- | --- |
| `part3-schema/OpenDocument-v1.4-os-part3-schema.html` | `43fb603f9f54f030db7082518aff6f136297d182abb20124a859d271a7969a15` | §2.2.1 C.1–C.3 requires a well-formed single XML file rooted at `office:document` and valid against Appendix A. §2.2.3 A and C allow the text-template MIME and require `office:body/office:text`. §3.1.2 defines the single-file `office:document` element and `office:mimetype`. |
| `schemas/OpenDocument-v1.4-schema.rng` | `4034ec6be29205d5fc1ee5f42468ac6ef824287b3aba6d9289032af4fafbda7f` | Appendix-A grammar source for the XML/root validation boundary. |
| `part2-packages/OpenDocument-v1.4-os-part2-packages.html` | `fe131d18f2f46c586827bc4c6b918896b232f9c1a9c161d0364de3c55060daef` | §3.3 applies to packaged `.ott` output: a package `mimetype` entry contains the ASCII template MIME, is first, uncompressed, and has no extra field. It does not define a `.fott` filename or alter the flat contract. |

The normative distinction is the root MIME, not the suffix. A valid flat text
template therefore has the same `office:document`/`office:body`/`office:text`
shape as a flat text document, with only the MIME classification changing.
The implementation must not infer template state from a path ending in `.fott`.

The implementation's admission boundary remains the existing bounded
well-formed-root/body/family-child check in `validate_flat_document`. It is not
a new full Appendix-A RNG validator, an ODF 1.4 version validator, or a claim
of complete ODF conformance; those normative rules explain why this MIME is
admissible, while the current codec's bounded checks remain the executable
contract for this gap.

The accepted constraints are also direct:

- ADR 0001 (`docs/adr/0001-priorities-and-api-layers.md:18-52`) makes
  correctness and typed refusal win, while requiring unsupported content to be
  readable and preservable.
- ADR 0003 (`docs/adr/0003-snapshots-edits-and-patches.md:8-25`) requires cheap
  immutable snapshots, atomic commit validation, source immutability, and
  reversible patches.
- ADR 0005 (`docs/adr/0005-io-memory-and-performance.md:19-46`) requires finite
  resource budgets, resource-rich limit errors, fresh output, atomic filesystem
  replacement, and a distinction between random-access edits and streaming
  writers.
- ADR 0006 (`docs/adr/0006-validation-security-and-compatibility.md:6-23`)
  makes preservation the default, keeps lexical details when possible, and
  forbids normal save from silently repairing or normalizing.
- ADR 0009 (`docs/adr/0009-odf-detection-ownership.md:20-44`) leaves flat ODF
  detection in `litchi-odf-common` and family semantics in the concrete owner.
- ADR 0023 (`docs/adr/0023-odf-family-crate-split.md:51-77`) places flat formats
  and their snapshot/edit/patch boundary in the owning family crate.
- `docs/GOAL.md:83-106` repeats the non-negotiable preservation, atomic edit,
  exact no-op, validation, and bounded-resource rules.

## Minimal support contract

### Admission and classification

1. Keep `litchi-odf-common::detect::{flat_mime, flat}` unchanged as the common
   detector. It should continue to return the canonical, decoded MIME for a
   valid root and neutral `Format::Odt` for both text MIME values.
2. Add `template: bool` to `generic::FlatDocument`, mirroring `Package`. In
   `FlatDocument::from_bytes_with_limit`, admit exactly
   `(Family::Text, true)` in addition to the currently admitted normal flat
   families. Continue refusing every other template MIME (spreadsheet,
   presentation, drawing, chart, formula, image, master, web, and database) at
   this boundary; their family owners can make separate decisions. This avoids
   accidentally claiming flat ODS/ODP template support while closing the ODT
   gap.
3. Validate a template with the existing namespace-aware root/body validator,
   UTF-8 conversion, and finite encoded-byte limit. The template flag must come
   from the validated root MIME. A filename, including `.fott`, has no authority
   over classification.
4. Add `FlatDocument::is_template()`. Keep `family()` as `Family::Text` and
   `mimetype()` as the detector's decoded/trimmed canonical MIME. The owned XML
   must still retain the original lexical attribute spelling and whitespace.
5. Return `fodt` from `extension()` for normal text and `fott` for text
   templates. Include the public common mappings
   `fott -> ODF_TEXT_TEMPLATE` in `ODF_EXTENSIONS` and
   `ODF_TEXT_TEMPLATE -> fott` in `ODF_FLAT_MIMETYPES`; cover
   `get_mime_type_from_extension("fott")`,
   `get_flat_extension_from_mime_type(ODF_TEXT_TEMPLATE)`, and
   `is_odf_extension("fott")` in the constants tests. These mappings and the
   `FlatDocument::extension()` result are project conventions for a
   nonnormative extension, not a claim that ODF standardizes the suffix.
6. Add the same `is_template()`, `mimetype()`, and `extension()` delegates to
   `litchi_odt::flat::Document`. Its existing `Family::Text` check remains the
   family boundary. No new package, umbrella, or detector layer is needed.

### Source, provenance, and transactional behavior

- Opening a flat template stores the source XML in the existing owned `String`.
  `as_bytes()`, `to_bytes()`, and `into_bytes()` must round-trip the exact input
  bytes after valid UTF-8 admission. `mimetype()` may be canonicalized by the
  detector; exact lexical source remains in `xml()` and the byte accessors.
- Every existing generic XML mutator must retain the template flag and validate
  the unchanged text-template root. It must re-detect the candidate's
  `office:mimetype` and require the exact original canonical MIME string and
  the exact original `(family, template)` classification, in addition to the
  existing body validation. Family-only validation is insufficient. No
  operation may silently rewrite the root MIME to normal text or infer a new
  MIME from a destination suffix.
- `flat::Document` keeps its current immutable `Arc<FlatDocument>` snapshot,
  detached `Edit`, atomic `Commit`, and source-bound `Patch`. An empty edit is
  an exact no-op: it shares the source snapshot and its patch's before/after
  bytes and template classification are identical. A changed paragraph edit
  replaces only the selected plain text span, then reopens the candidate under
  the same limits and performs the existing semantic readback.
- `Patch::apply` continues to compare the complete source bytes before returning
  its retained target. `Patch::inverse()` swaps the retained snapshots, so a
  changed template edit followed by its inverse restores the original bytes,
  MIME, and `is_template()` value exactly. A different source, including a
  normal `.fodt` with otherwise equal body content, is a typed refusal.
- No-op saves preserve the original XML. Both `flat::Document::save` and
  generic `FlatDocument::save` must use the ADR 0005 sibling temporary,
  flush/sync, length validation, and atomic replacement path. The generic
  direct `std::fs::write` path is part of this support slice's source change;
  it is not an inherited limitation to scope out. Adding template support
  must not turn either API into a package writer. A `.fott` path is accepted
  as a caller convention, while bytes remain the source of truth.
- Preserve the current compactness gate for changed specialized edits: a
  formatted source can be opened and no-op saved exactly, while a changed edit
  that would violate the existing compact-XML validation still fails without
  mutating the source. Do not normalize a template merely to make an edit pass.

### Edits and output boundary

The existing text-family operations are the complete initial write surface:
paragraph selection, replacement text escaping, structural/opaque paragraph
refusal, candidate validation, semantic readback, and exact patch/inverse. The
same operations apply to a text template because ODF 1.4 gives the flat text
template the same `office:text` body grammar. Template instantiation, applying
a template to a document, field refresh, macro/script execution, external-link
refresh, and rendering remain unsupported and inert.

The generic family-neutral APIs may continue to expose modeled metadata and
opaque XML preservation. They must not gain a template-application verb or an
ODS/ODP template claim as a side effect of this change.

### Limits and memory

Keep the existing finite limits and error shape:

| Boundary | Default | Hard ceiling |
| --- | ---: | ---: |
| Generic flat input | 256 MiB | 512 MiB |
| `flat::Limits::max_document_bytes` | 16 MiB | 64 MiB |
| One paragraph replacement | 256 KiB | 1 MiB |
| Staged edits | 256 | 1,024 |
| XML depth inspected by an edit | 256 | 4,096 |

The future implementation should pass the specialized document limit to
`FlatDocument::from_bytes_with_limit` rather than falling back to the generic
default. Reader paths continue to read at most `limit + 1`, charge the input
before creating the owned XML string, and return `Error::ResourceLimit` with
resource, observed value, limit, and scope. The exact cap succeeds and cap plus
one fails for both a normal document and a template.

The template bit adds no content-sized allocation. MIME detection's bounded
canonical string, the one owned XML string, and the existing `Arc` snapshot are
the retained state. Changed edits retain the current checked output-length
calculation and `try_reserve_exact` before rendering; an output over the
configured document limit fails before publication. Failed admission or commit
must leave the source and any destination untouched.

## Regression evidence

There is no checked-in native `.fott` fixture in this workspace or in the
LibreOffice source tree reachable through `3rdparty/libreoffice-core` (a search
of `/home/zhuhe/code` found none). Existing flat ODT fixtures are normal text;
for example, `test-data/odf/odt/font-face-declarations-flat.fodt` has SHA-256
`b7badc6860eb63a587c762a1887b6e140946b1e932d0790830646937bcc829e9` and a
normal text MIME. This absence is evidence that a new flat-template fixture is
needed, not evidence that the format is invalid.

The native LibreOffice corpus does provide template package provenance. Each of
the following has the package `mimetype` value
`application/vnd.oasis.opendocument.text-template` (read from the uncompressed
`mimetype` member):

| Native package fixture | SHA-256 |
| --- | --- |
| `3rdparty/libreoffice-core/sw/qa/filter/md/data/template.ott` | `844e73a34df4d048c9f527a8c787bba5ffb0c6f747dc57f506b624ad7bcf33a7` |
| `3rdparty/libreoffice-core/filter/qa/unit/data/writer.ott` | `6c6db856e772bd541b3f644609e93b482786c929b4e772a17fd0391affa9503b` |
| `3rdparty/libreoffice-core/sw/qa/extras/odfexport/data/tdf131621.ott` | `d7d77d8fe15c16d238599299817a6a92ad9351e79b9c97a4741c672fccd571bf` |

These `.ott` files are useful native template evidence but are packages and
must not be passed directly to a flat reader. The first flat regression should
be a small checked-in XML fixture with a text-template root MIME, namespace
prefix variation, a preserved comment/attribute order, and at least one plain
and one structured paragraph. A native-derived flat fixture can be added only
after a reproducible LibreOffice export is available and its producer version
and hash are recorded; currently there is no such artifact to cite.

## Owned implementation and test paths

The eventual code change belongs to the following existing owners; this report
does not edit them:

- `crates/litchi-odt/src/generic/model.rs`: template state on
  `FlatDocument` and its accessors' storage.
- `crates/litchi-odt/src/generic/flat.rs`: text-template admission, bounded
  limit handoff, `fodt`/`fott` extension behavior, and atomic generic save.
- `crates/litchi-odt/src/generic/codec.rs`: existing MIME-to-family/template
  classification remains the single codec source of truth; its current text
  template arm should be covered by the acceptance test rather than duplicated.
- `crates/litchi-odt/src/generic/tests.rs`: generic template acceptance,
  classification, exact byte round-trip, bounded body/malformed rejection,
  and atomic save/reopen behavior for a template destination.
- `crates/litchi-odt/src/flat/mod.rs`: specialized metadata delegates and
  preservation of template state through edit, commit, save, and patch paths.
- `crates/litchi-odt/tests/flat_family_readers.rs`: fixture-backed template
  reads, metadata, exact bytes, save/reopen, and packaged-input rejection.
- `crates/litchi-odt/tests/flat_family_authoring.rs`: template no-op, paragraph
  edit, source conflict, inverse, atomic failure, and output-limit behavior.
- `crates/litchi-odf-common/src/constants.rs` and its tests: include the
  `.fott` convenience mapping required above. The detector logic
  in `crates/litchi-odf-common/src/detect.rs` should receive a regression test
  for flat text-template MIME if coverage is not already added, but it should
  not change classification semantics.

The feature matrix (`crates/litchi-odt/docs/FEATURE_MATRIX.md`) and
`docs/report/spec-gap-audit.md` are later status-document owners. They should
change only after the fixture-backed read, edit, save, limit, and inverse gates
pass. No ODS/ODP source or tests are part of this slice.

## Acceptance and non-goals

The gap is closed when a valid flat text-template is admitted with
`family == Family::Text`, `is_template() == true`, canonical template MIME,
and `extension() == "fott"`; normal text still reports `false`/`fodt`; both
round-trip exactly; edits preserve classification; no-op and inverse are exact;
source conflicts and resource caps fail atomically; and `.ott` package behavior
is unchanged.

This slice does not add flat spreadsheet or presentation templates, master/web
flat forms, package conversion, template instantiation, field or macro
execution, external-resource access, rendering, or a native `.fott` round-trip
claim. No Cargo, profile, benchmark, or implementation job was run for this
read-only design.
