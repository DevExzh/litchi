# XLSX SVG attachment lifecycle design

This is a design artifact only. It does not claim an XLSX SVG host
implementation, native Office acceptance, or a completed validation gate.
The bounded batch is ordinary worksheet `SpreadsheetDrawing` pictures with an
existing raster fallback. It covers reading, attaching an embedded SVG, and
detaching that SVG from an existing picture. Chartsheets, chart user-shape
drawings, DOCX, linked-resource fetching, SVG rendering, and picture creation
are separate work.

## Local normative basis

The host and shared codec should use these vendored references as the normative
input:

- `3rdparty/specs/[MS-ODRAWXML]/1 Introduction/1.3 Overview.md`, §1.3.3,
  defines `svgBlip`, the image-part relationship, and the rasterized PNG
  compatibility fallback.
- `3rdparty/specs/[MS-ODRAWXML]/2 Structures/2.26 http---schemas.microsoft.com-office-drawing-2016-SVG-main.md`,
  §§2.26.1.1 and 2.26.3.1, defines the `asvg:svgBlip` element and its
  `AG_Blob` `r:embed`/`r:link` attributes.
- `3rdparty/specs/[MS-ODRAWXML]/5 Appendix A - Full XML Schemas/5.24 http---schemas.microsoft.com-office-drawing-2016-SVG-main Schema.md`
  is the local SVG extension schema.
- `3rdparty/specs/[MS-ODRAWXML]/2 Structures/2.2 Extensions.md`, §§2.2.2
  and 2.2.6, defines the extension/MCE integration boundary and lists
  SpreadsheetDrawing `twoCellAnchor`, `oneCellAnchor`, `absoluteAnchor`,
  `pic`, and `grpSp` extension parents. The first batch admits direct picture
  owners only and refuses ambiguous group/MCE ancestry.
- The local ECMA schema archives
  `3rdparty/specs/ECMA-376/ECMA-376-1_5th_edition_december_2016.zip` and
  `ECMA-376-4_5th_edition_december_2016.zip`, nested
  `OfficeOpenXML-XMLSchema-Strict.zip` and
  `OfficeOpenXML-XMLSchema-Transitional.zip`, provide
  `dml-spreadsheetDrawing.xsd`, package relationship schemas, and content-type
  validation for the complete changed package.
- `docs/adr/0003-snapshots-edits-and-patches.md` requires immutable
  source-bound snapshots, atomic changed-dependency-closure validation, exact
  in-memory inverse patches, and stale-source refusal.
- `docs/adr/0007-office-object-models.md` requires borrowed semantic
  traversal and ownership-consuming insertion failures to return the rejected
  value, so a failed SVG insertion cannot lose the caller's payload.

The native extension URI used by current producer evidence is
`{96DAC541-7B7A-43D3-8B79-37D633B846F1}`. The local MS-ODRAWXML material
normatively defines the `svgBlip` child and relationship semantics; the GUID is
an admitted native host profile corroborated by the fixture below. An unknown
`a:ext@uri` remains opaque and is never inferred to be an SVG owner.

## Native evidence

The smallest useful producer fixture is:

`3rdparty/libreoffice-core/sc/qa/unit/data/xlsx/tdf169496_hidden_graphic.xlsx`

Full archive SHA-256:

`0b647da300a085f39914fdfae961463ae9e54ffe772b2e0eb9860a841ab93f72`

The archive is 12,470 bytes. Its ordinary worksheet drawing members are:

| Member | Size | SHA-256 |
|---|---:|---|
| `xl/drawings/drawing1.xml` | 2718 | `c9d4149c14d847d4979239fc9771de0f34383c5e01ee505e270aea1bce05ffe8` |
| `xl/drawings/_rels/drawing1.xml.rels` | 427 | `308917dc6427cacfe0ef7a7cc4b74fc058b7b31e21289de1055af09e59e675ce` |
| `xl/media/image2.svg` | 313 | `05769bd518f6ce504ee4f465a878eb8bd848c906e23c1adee6491e1653167cbf` |

`drawing1.xml` contains two `xdr:twoCellAnchor` pictures. Each picture has a
main raster `a:blip r:embed="rId1"`, followed by an `a:extLst/a:ext` with the
native URI and `asvg:svgBlip r:embed="rId2"`. Both pictures share the same
raster and SVG relationships. `[Content_Types].xml` uses a default `svg`
mapping to `image/svg+xml`.

This gives a real shared-owner test: detaching the first picture must leave the
SVG relationship and part because the second picture still owns them; after
detaching the second picture, the SVG resource can be removed only when a
complete package reachability check proves that no incoming relationship
remains. The fixture demonstrates producer shape and graph topology, not
acceptance of newly generated files by Excel or LibreOffice.

## Current implementation and the required base work

The current ordinary worksheet reader is intentionally small:

- [`crates/litchi-xlsx/src/drawing/mod.rs`](../../../crates/litchi-xlsx/src/drawing/mod.rs)
  exports `drawing::parse` and the inventory model.
- [`crates/litchi-xlsx/src/drawing/model.rs`](../../../crates/litchi-xlsx/src/drawing/model.rs)
  gives `Picture` complete `DrawingAnchor` geometry alongside the legacy
  cell-anchor projection, main image relationship ID, and description.
  SVG descriptors and source ranges remain lifecycle work.
- [`crates/litchi-xlsx/src/drawing/codec.rs`](../../../crates/litchi-xlsx/src/drawing/codec.rs)
  now recognizes all three anchor forms with actual geometry and `editAs`,
  validates direct ownership and geometry order, and reads the main
  `a:blip@r:embed`. MCE input/output and parser allocations are bounded.
  The numeric coordinate profile excludes universal-measure lexical forms.
- [`crates/litchi-xlsx/src/workbook/edit/drawing_transfer.rs`](../../../crates/litchi-xlsx/src/workbook/edit/drawing_transfer.rs)
  already scans source ranges and recognizes all three anchor element names
  for transfer layout purposes. That transfer path is not an SVG owner: it
  projects selected anchors and copies closed image/chart dependencies.
  `validate_image_leaf` already admits an internal `/xl/media/*.svg` part with
  `image/svg+xml`, but does not discover or mutate `svgBlip` extensions.
- [`crates/litchi-xlsx/src/workbook/edit/semantic/transaction.rs`](../../../crates/litchi-xlsx/src/workbook/edit/semantic/transaction.rs)
  stages worksheet bytes and drawing graph changes for cell transfers.
  [`workbook/edit/model.rs`](../../../crates/litchi-xlsx/src/workbook/edit/model.rs)
  has source-checked `PartChange`, relationship deltas, and `GraphChange`.
  The existing `GraphChange` with `GraphAction::Remove` removes both a relationship and its part;
  it cannot be used blindly for a shared SVG target.
- [`crates/litchi-drawingml/src/svg_blip.rs`](../../../crates/litchi-drawingml/src/svg_blip.rs)
  already supplies the bounded, namespace-aware, source-preserving
  `svgBlip::{read,write,write_to,SvgBlip}` codec. XLSX should reuse it and keep
  package/resource ownership in `litchi-xlsx`.

The first implementation sequence must therefore be explicit:

1. **Complete the base anchor projection.** Add `oneCellAnchor` to the XLSX
   reader or route the inventory through the existing
   `litchi-spreadsheet-drawing::shape::Anchor` model, which already has
   `TwoCell`, `OneCell`, and `Absolute` variants. Preserve `editAs` for a
   two-cell anchor, `from` plus `ext` for a one-cell anchor, and `pos` plus
   `ext` for an absolute anchor. Do not map absolute geometry to the current
   zero placeholder in a source-backed SVG selector.
2. **Add one source-backed worksheet drawing scan.** It must record anchor and
   direct `xdr:pic` spans, resolved namespaces, the main raster blip, the
   direct recognized SVG extension, and the source bytes needed for a narrow
   splice. The lossy `Drawing` inventory remains useful for read projection but
   is not the preservation authority for edits.
3. **Add the SVG owner projection.** Resolve one direct `a:blip/a:extLst/a:ext`
   owner by expanded names and the admitted URI profile. Parse its
   `asvg:svgBlip` fragment with the shared codec. Read resource metadata only
   after validating the drawing relationship closure.
4. **Add attach/detach staging for every admitted anchor form.** The operation
   must use the same source-backed owner and graph planner for two-cell,
   one-cell, and absolute pictures. A missing base one-cell reader is a
   prerequisite, not a reason to exclude one-cell pictures from the batch.

The existing `litchi-spreadsheet-drawing::shape::Anchor` reader/writer is a
possible base implementation, but importing it does not by itself provide
source spans, opaque extension preservation, or package relationship ownership.
Those remain XLSX host responsibilities.

## Source helper contract

The source-backed scanner supplies direct-picture records selected by semantic
drawing and picture ordinals. Each record carries:

- validated two-cell, one-cell, or absolute anchor geometry;
- raw source byte ranges for `xdr:pic`, its direct `a:blip`, and the admitted
  extension, together with the inherited namespace context;
- the raster `r:embed`; and
- at most one admitted direct SVG extension owner parsed by the shared
  `SvgBlip` codec, or a typed ambiguity/MCE refusal.

The scanner borrows source bytes and preserves opaque siblings. Source ranges
and namespace context remain the preservation authority; a typed projection
must not cause serialization of the whole drawing. The helper discovers
ownership and supplies source context. Mutation, byte splicing, and OPC graph
planning remain in the worksheet transaction, which accepts borrowed SVG
input and validates the complete staged dependency closure.

## Host ownership and ordinary API

The ordinary facade should select a worksheet and picture semantically. It
must not require a caller to supply a relationship ID, package part URI, XML
offset, or generated name. Selection uses checked drawing and direct-picture
ordinals in source order. Duplicate nonvisual IDs or an ambiguous selector
return a typed refusal.

The public shape can follow this outline; names are provisional:

```rust,ignore
pub struct SvgImageView<'a> {
    pub reference: &'a litchi_drawingml::svg_blip::Reference,
    pub payload: Option<SvgPayloadView<'a>>,
    pub profile: SvgOwnerProfile,
}

pub enum PictureSelector {
    Position { drawing: usize, picture: usize },
    // A semantic cNvPr selector may be added only with duplicate checks.
}

pub fn attach_svg(
    &mut self,
    worksheet: impl Into<Selector<'_>>,
    picture: PictureSelector,
    svg: SvgInput<'_>,
) -> Result<...>;

pub fn detach_svg(
    &mut self,
    worksheet: impl Into<Selector<'_>>,
    picture: PictureSelector,
) -> Result<...>;
```

`SvgInput` may borrow caller bytes for preflight or own an `Arc<[u8]>` for a
long-lived prepared request. The operation must validate the borrowed input
before copying it; if an ownership-consuming form fails, it returns the
rejected input as required by ADR 0007. Callers do not choose `rId`, media
part names, relationship targets, extension prefixes, or URI lexical spelling.

Attach is intentionally limited to an existing direct raster picture with no
recognized SVG owner. It adds one embedded SVG owner and retains the existing
raster fallback. Detach removes only the selected SVG owner. Replacing an
existing SVG payload, adding a new picture anchor, changing raster fallback,
and converting linked SVG to embedded SVG are separate verbs.

The read projection should expose the three anchor geometries without making
the resource bytes mandatory. A source-backed view may borrow the selected
drawing XML and relationship records; the package snapshot remains the
ownership authority. The view must distinguish:

- `TwoCell`: validated `from`, `to`, and `editAs`;
- `OneCell`: validated `from` and EMU `ext`; and
- `Absolute`: validated EMU `pos` and `ext`.

An SVG view with `r:link` may be reported as linked/inert, but the first edit
profile refuses it. An embedded owner must resolve to an internal image
relationship whose target is under `/xl/media/`, whose content type is exactly
`image/svg+xml`, and whose part has no outbound relationships. The raster
fallback required for newly authored output must resolve to an internal PNG
image relationship (under `/xl/media/`, with `image/png`) in accordance with the
local SVG overview's compatibility profile. Existing nonconforming pictures
are preserved as unsupported/inert rather than silently rewritten.

## Recognized owner and XML policy

The scanner recognizes only a direct `a:blip` child of a direct picture's
`xdr:blipFill`, its direct `a:extLst`, and an `a:ext` whose XML-token-normalized
URI is the admitted native GUID. It then requires exactly one direct
`asvg:svgBlip` child in the SVG namespace. Prefix spelling, inherited
namespace declarations, comments, and whitespace are resolved by expanded
name; untouched source bytes retain their original lexical form.

Unknown extension URIs, a foreign `svgBlip`-looking child, duplicate admitted
owners, nested group ownership, and any `mc:AlternateContent` ancestry are
opaque or refused according to operation:

- ordinary read retains them as inert data and does not infer an SVG resource;
- attach refuses an ambiguous or MCE-wrapped picture rather than appending a
  second effective owner; and
- detach refuses unless the selected direct owner is unambiguous.

Keep the core drawing and package relationship dialect separate from the
Microsoft extension vocabulary. Core DrawingML and raster relationship
attributes follow the source's Transitional or Strict dialect; physical image
relationship types follow that same host dialect. Newly authored
`asvg:svgBlip@r:embed` uses the Transitional relationship attribute namespace
required by the unmodified MS-ODRAWXML §5.24 schema in either host dialect.
The schema imports Transitional `a:AG_Blob`, so changing this extension
attribute to the Strict namespace fails that schema. A Strict core drawing
with the Microsoft extension retaining Transitional attributes validates
against both the Strict drawing schema and the unmodified extension schema.

The shared codec's acceptance of both attribute namespaces is a read
compatibility capability, not permission to author either vocabulary while
claiming the MS schema. Preserve existing source-only/no-op bytes; do not
silently normalize a legacy extension merely while reading it. Use local
namespace declarations or collision-free prefixes when the host's `r` prefix
is already bound to the Strict namespace.

When attaching, splice the selected `a:blip` without normalizing unrelated
attributes, namespace declarations, comments, processing instructions, or
extension siblings. Expand a self-closing `a:blip` only after a checked output
precharge. When detaching, remove the selected `a:ext`; remove its `a:extLst`
wrapper only if it has no opening-tag attributes (including namespace
declarations), unrelated extension, comment, or other preserved payload. A wrapper containing unrelated content remains byte-preserved.

## Package graph and dependency closure

An attach commit changes one or more of these members atomically:

1. the owning `xl/drawings/drawingN.xml` source bytes;
2. the drawing relationship part `xl/drawings/_rels/drawingN.xml.rels`;
3. a new `/xl/media/<generated>.svg` leaf; and
4. `[Content_Types].xml` only when the existing default/override mapping does
   not already cover the new SVG part.

The worksheet XML and its worksheet relationship remain unchanged for an
existing drawing owner, but the transaction must verify that closure rather
than assume it. Detach changes the drawing XML and drawing relationships. It
removes the SVG media part and an owned content-type entry only after a full
package incoming-edge scan proves that no other relationship targets the part.
If another drawing or another picture still references the same SVG target,
the target and its content type remain. A shared relationship ID in one drawing
is also handled as a reference-counted graph edge, not as an unconditional
part deletion.

The existing `GraphChange` with `GraphAction::Remove` couples edge removal to part removal and is
therefore insufficient for this shared-target case. The implementation needs
either a relationship-only graph delta plus a separately guarded part-removal
delta, or an owner-specific graph planner that validates reachability before
publishing. It must not delete a part merely because the selected edge was
removed.

The candidate package is reopened under the source limits after staged XML,
relationship, media, and content-type changes. Readback must verify:

- the selected anchor and picture still resolve;
- the main raster fallback is unchanged;
- the requested SVG owner state and payload hash are present or absent;
- every changed relationship has the expected type, target mode, and target;
- all changed media parts have matching content types and no forbidden
  outbound relationships; and
- no dangling relationship, duplicate content-type override, or orphaned
  media owner exists in the changed closure.

## Source, patch, and inverse requirements

The source drawing XML, relationship XML, and content-type XML are the
preservation authority. A semantic model must not be serialized back as a
canonical whole drawing because the current model does not retain unknown
extensions or producer formatting.

The staged edit records source fingerprints and exact immutable bytes for every
changed member, plus the selected owner fingerprint, relationship topology,
anchor geometry, and SVG payload hash. Applying it to a different drawing,
relationship part, content-type part, or package lineage returns a stale or
patch conflict before output. A failed operation leaves the transaction and
the borrowed/owned SVG input untouched.

An exact semantic no-op shares the source snapshot and performs no replacement
or candidate reopen. A changed commit publishes only after the complete
dependency closure has been validated and read back. Its in-memory inverse
restores the exact accepted source bytes and graph topology. An independently
saved and reopened package requires fresh owner/source authorization. This batch does not promise durable replay of an in-memory
inverse or recovery of original lexical choices after a fresh detach.

## Limits and allocation policy

The operation should reuse the existing finite XLSX and shared SVG limits and
add an explicit host budget for aggregate mutation work:

- drawing XML: the existing 32 MiB scanner bound;
- SVG fragment metadata: `litchi-drawingml::svg_blip`'s bounded XML,
  namespace, attribute, child, depth, and relationship-ID limits;
- SVG media payload: the existing bounded image-leaf limit, with a caller
  limit checked before copying or staging;
- anchors, relationships, package parts, and changed-member count: bounded
  before index/vector allocation;
- content-types and relationship XML: retained OPC limits plus a checked
  prospective output size; and
- aggregate source/output bytes: a caller-visible edit budget covering the
  drawing, relationship, content-type, and media staging buffers.

The scanner should use one raw source pass and share immutable ancestor/source
layers. It must not clone the whole drawing for every candidate picture. URI,
relationship ID, namespace, and media path bounds are checked before storing
decoded values. Every insertion path precharges the prospective output length
before constructing a replacement `Vec`; a final splice check alone is not
enough for a small caller limit. SVG payload bytes remain opaque and are not
parsed as an SVG document.

## Implementation and test sequence

The implementation should land in bounded steps so the base reader and the
host graph owner are independently reviewable.

1. **Anchor reader completion (implemented).** The reader reuses shared
   anchor geometry types for two-cell, one-cell, and absolute pictures.
   Tests cover Strict namespaces, missing/duplicate geometry, bounds,
   source order, XML numeric whitespace, and MCE expansion limits. The
   added public geometry field requires downstream struct literals to
   initialize it; legacy field access remains available.
2. **Source-backed owner scan.** Add a source range model for direct pictures,
   resolved `a:blip`/`extLst` ownership, MCE ancestry, unknown extension spans,
   and all three anchor geometries. Add tests for inherited/default prefixes,
   comments, CDATA, self-closing blips, duplicate recognized owners, and
   unknown URI preservation.
3. **Shared SVG projection.** Reuse `litchi_drawingml::svg_blip::codec::read`
   and expose deferred package payload metadata. Test embedded/linked reads,
   strict/transitional relationship namespaces, malformed fragments, and
   relationship closure failures.
4. **Attach/detach transaction.** Add source-bound worksheet selectors,
   borrowed/owned input handling, XML/rels/content-type/media staging, shared
   target reachability, exact inverse, stale checks, and candidate readback.
5. **Native and package gates.** Run the focused XLSX suite, complete package
   reopen tests, local ECMA schema checks for all three anchor forms, strict
   clippy/rustdoc/fmt, and fresh-process allocation/timing samples. Record
   source manifests and fixture hashes only after the implementation is frozen.

The focused lifecycle matrix should include:

| Case | Required result |
|---|---|
| Native `tdf169496_hidden_graphic.xlsx` read | Two pictures expose PNG fallback and the admitted embedded SVG owner. |
| Synthetic two-cell raster picture attach | Adds one valid `svgBlip`, SVG relation, media leaf, and required content type. |
| Synthetic one-cell attach | Same graph and source guarantees after the base one-cell reader is present. |
| Synthetic absolute attach | Same graph and source guarantees while retaining `pos`/`ext` geometry. |
| Detach one of the native shared owners | Removes only that owner; shared SVG relation/part remains. |
| Detach the final shared owner | Removes the SVG edge/part only after reachability and content-type checks. |
| Exact no-op, inverse, stale, replay | No-op shares bytes; inverse restores exact source; stale/replay refuse atomically. |
| Unknown/duplicate/MCE/linked/malformed owner | Read inert where safe; attach/detach refuses ambiguity without duplication. |
| Namespace/dialect variants | Recognizes resolved names, matches source dialect, and preserves lexical source. |
| Limits and failure sink | Rejects before oversized temporary buffers or partial package publication. |

## Profiling and evidence boundaries

Fresh-process profiling should vary drawing size, number of pictures, three
anchor forms, shared versus distinct SVG resources, large unknown extension
payloads, inherited namespace declarations, and caller output limits. Report
allocation and wall-time bounds for the scanner, source splice, graph planner,
and candidate reopen separately. No linear-runtime claim follows from the
existing parser without those measurements.

The native fixture establishes producer structure and relationship topology.
Offline ECMA/MS-ODRAWXML schema checks establish supported XML shape. Neither
establishes Excel rendering, Office acceptance of newly generated files, SVG
security/rendering semantics, external-link fetching, or image conversion.
Chart user-shape fixtures such as `test-data/libreoffice-core/chart2/qa/extras/data/xlsx/tdf143127.xlsx`
are deliberately excluded from this ordinary worksheet batch. DOCX and
chartsheet ownership are later host batches.

No production source or test file is changed by this design artifact.
