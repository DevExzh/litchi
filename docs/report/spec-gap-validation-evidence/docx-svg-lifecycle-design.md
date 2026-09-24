# DOCX SVG attachment lifecycle design

This is a design artifact only. It does not change DOCX production code or
tests, and it does not claim native Office acceptance or rendering. The
bounded batch is source-backed SVG read, attach, and detach for direct
`pic:pic` pictures in the main Word story (`/word/document.xml`). It admits
both `wp:inline` and `wp:anchor` placement, preserves the existing PNG
fallback, and edits only an existing picture. Headers, footers, notes,
comments, glossary stories, MCE branches, grouped drawings, picture creation,
and linked-resource fetching are separate work.

The story abstraction should be designed so that the same host scanner can be
enabled for subsidiary stories in subsequent increments. The retained native
positive evidence currently covers the main story. Completing this first
increment does not close DOCX SVG support across all story owners; those
owners still require implementation and validation under the broader goal.

## Local normative basis

The host and shared codec should use these vendored references as the
normative input:

- `3rdparty/specs/[MS-ODRAWXML]/1 Introduction/1.3 Overview.md`, §1.3.3,
  defines the SVG illustration profile: `svgBlip` carries the relationship
  identifier for embedded or linked SVG data, embedded data is an image part,
  and Office keeps a rasterized PNG fallback in the main blip for
  compatibility;
- `3rdparty/specs/[MS-ODRAWXML]/2 Structures/2.26
  http---schemas.microsoft.com-office-drawing-2016-SVG-main.md`, §§2.26.1.1
  and 2.26.3.1, defines the `asvg:svgBlip` element and its `r:embed`/
  `r:link` `AG_Blob` attributes;
- `3rdparty/specs/[MS-ODRAWXML]/5 Appendix A - Full XML Schemas/5.24
  http---schemas.microsoft.com-office-drawing-2016-SVG-main Schema.md` is the
  unmodified local SVG extension schema. It imports the Transitional
  DrawingML base types and Transitional Office relationship schema;
- `3rdparty/specs/[MS-ODRAWXML]/2 Structures/2.18
  http---schemas.microsoft.com-office-word-2010-wordprocessingDrawing.md`,
  and its §5.11 schema, define the Wordprocessing Drawing `inline`/`anchor`
  extension vocabulary, including `anchorId`, `editId`, and relative-size
  children. The base `CT_Inline` and `CT_Anchor` grammar is in the local
  ECMA schema archives;
- `3rdparty/specs/[MS-ODRAWXML]/3 Structure Examples/3.7 WordprocessingML
  Drawing.md` provides the local Word drawing nesting examples;
- `3rdparty/specs/ECMA-376/ECMA-376-1_5th_edition_december_2016.zip`, nested
  `OfficeOpenXML-XMLSchema-Strict.zip`, and
  `ECMA-376-4_5th_edition_december_2016.zip`, nested
  `OfficeOpenXML-XMLSchema-Transitional.zip`, provide `wml.xsd`,
  `dml-wordprocessingDrawing.xsd`, `dml-picture.xsd`, package relationship
  schemas, and content-type validation. The complete changed package must be
  checked against the selected core dialect and the unmodified SVG extension
  schema;
- `docs/adr/0003-snapshots-edits-and-patches.md` requires immutable,
  source-bound snapshots, atomic changed-dependency closure validation, exact
  in-memory inverse patches, and stale-source refusal;
- `docs/adr/0006-validation-security-and-compatibility.md` requires
  preservation and fail-closed handling of unsupported XML, relationships,
  and MCE ownership; and
- `docs/adr/0007-office-object-models.md` requires borrowed semantic
  traversal and ownership-consuming insertion failures to return the rejected
  value, so a failed SVG insertion cannot lose the caller's payload.

The native extension URI
`{96DAC541-7B7A-43D3-8B79-37D633B846F1}` is an admitted producer profile
only. The local normative material defines `svgBlip` and its relationship
semantics; it does not make this GUID a universal schema discriminator. An
unknown `a:ext@uri` remains opaque and is never inferred to be an SVG owner.

## Native evidence

The smallest retained inline producer fixture is:

`3rdparty/Open-XML-SDK/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestFiles/svg.docx`

Its full archive SHA-256 is
`1ce06ef88f89c5e14d1d22cd8e4b3a5f70951e7bd53e19053d4942873333d21d`
(74,104 bytes). Relevant members are:

| Member | Size | SHA-256 |
|---|---:|---|
| `word/document.xml` | 3,850 | `96598ce1dfeb683a681dc3ac4c75cc2f027aae6dc553efc5bc2997af2b2e5488` |
| `word/_rels/document.xml.rels` | 1,081 | `a7b137a8d52d66ab4acd4777d49efe56165e73ce5e11f02c9096c5da54206ead` |
| `word/media/image1.png` | 61,192 | `7b717c293ec7b9022c5600d34586770a529d04db49ccdd7349be2254e58485bb` |
| `word/media/image2.svg` | 261 | `69e5bc39f90811ab154fab3cc3312b692901605af3be636c181a16fff6bd100a` |
| `[Content_Types].xml` | 1,416 | `d68fd294a2a4733d96486a7e9ed856c3947803c2329747bcf868ada733af8bf0` |

Its document contains a direct `wp:inline` picture. The `pic:blipFill`
contains a raster `a:blip r:embed="rId4"` and one direct extension with the
observed GUID and `asvg:svgBlip r:embed="rId5"`. Both image relationships are
owned by `word/document.xml`; the SVG target is `word/media/image2.svg` and
the raster target is `word/media/image1.png`. The content-types part uses a
default `svg` mapping to `image/svg+xml`.

The retained floating-anchor producer fixture is:

`3rdparty/libreoffice-core/sw/qa/extras/ooxmlexport/data/tdf164835_nonDummyLineHeight.docx`

Its full archive SHA-256 is
`569572dc2ec5dfee9334504ae331482ebb8188b80a0d3c93d44b8884cda29283`
(302,210 bytes). Relevant members are:

| Member | Size | SHA-256 |
|---|---:|---|
| `word/document.xml` | 5,599 | `172c4d6ca0a3313e6e3a1f90e1c49730c44e97a7dc452c574c6fd02470f2421f` |
| `word/_rels/document.xml.rels` | 1,081 | `a7b137a8d52d66ab4acd4777d49efe56165e73ce5e11f02c9096c5da54206ead` |
| `word/media/image1.png` | 281,425 | `ad2edd992ada1d78cb6b24b404394383ca3b4c8b6ea69899fa60c8da93294c09` |
| `word/media/image2.svg` | 6,439 | `6d0c973af450465a7bb3e0a52443af244d0958374635d178e8e43bbc9a899491` |
| `[Content_Types].xml` | 1,416 | `d68fd294a2a4733d96486a7e9ed856c3947803c2329747bcf868ada733af8bf0` |

Its document contains direct `wp:anchor` pictures with the same raster/SVG
relationship pattern. The anchor geometry, wrapping, and `wp14` children are
producer data and must remain byte-preserved by an SVG-only edit. These two
fixtures establish inline and floating host shape, relationship ownership,
and the native GUID profile. They do not establish acceptance of newly
authored output by Word or LibreOffice.

## Current implementation and required base work

The existing ordinary image and drawing APIs are metadata readers and cannot
serve as the preservation authority:

- [`crates/litchi-docx/src/image.rs`](../../../crates/litchi-docx/src/image.rs)
  exposes `InlineImage`, resolves one `r:embed`, and parses only
  `wp:inline`. It ignores `wp:anchor`, direct `a:extLst/asvg:svgBlip`, and
  source ranges. Its parser also matches local names without the expanded-name
  ownership checks required for a safe edit.
- [`crates/litchi-docx/src/format.rs`](../../../crates/litchi-docx/src/format.rs)
  has no SVG image format. This is a read/authoring classification gap, not a
  reason to parse the SVG payload; the host should retain SVG bytes as opaque
  media and validate only package graph metadata.
- [`crates/litchi-docx/src/drawing/model.rs`](../../../crates/litchi-docx/src/drawing/model.rs)
  and `drawing/codec.rs` already inventory `Anchor::Inline` and
  `Anchor::Floating`, but their `Object` projection is shape-oriented and does
  not capture `pic:pic`, image relationships, extension owners, or source
  spans.
- [`crates/litchi-docx/src/writer/image.rs`](../../../crates/litchi-docx/src/writer/image.rs)
  emits a canonical raster inline picture. It does not preserve an existing
  drawing and must not be used for this lifecycle; it also cannot create a
  complete floating anchor or SVG extension.
- [`crates/litchi-docx/src/package/story.rs`](../../../crates/litchi-docx/src/package/story.rs)
  already derives story ownership from OPC relationships and validates main,
  header, footer, footnote, endnote, comment, and glossary parts. Its
  `StoryPart::source()` is an immutable eager snapshot. The crate-private
  `package/story/source.rs` retains `PartData` for deferred source-backed
  reads, which is the correct ownership model when a future source-backed
  facade is exposed.
- [`crates/litchi-docx/src/package/package/story_edit.rs`](../../../crates/litchi-docx/src/package/package/story_edit.rs)
  demonstrates source-checked story replacement and candidate readback.
  [`crates/litchi-docx/src/ink/transaction.rs`](../../../crates/litchi-docx/src/ink/transaction.rs)
  demonstrates graph closure, staged-byte limits, `OwnedXmlPart` replacement,
  stale checks, and atomic semantic publication.
- [`crates/litchi-docx/src/package/package/access.rs`](../../../crates/litchi-docx/src/package/package/access.rs)
  supplies the crate-private `edit_semantic_opc` transaction boundary. The
  implementation should use `litchi-opc`'s `OwnedXmlPart`,
  `OwnedRelationships`, `OwnedContentTypes`, and source-checked replacement
  APIs rather than the legacy mutable document writer.

The required base work is therefore one source-backed Word drawing scanner,
one graph-aware SVG owner projection, and one host transaction. It is not a
change to the shared `svg_blip` codec and it does not require a second SVG
grammar implementation.

## Host grammar and owner policy

The scanner must resolve expanded names and inherited namespace bindings. The
direct admitted path is:

```text
w:drawing
  / wp:inline | wp:anchor
    / a:graphic
      / a:graphicData[@uri="http://schemas.openxmlformats.org/drawingml/2006/picture"]
        / pic:pic
          / pic:blipFill
            / a:blip[@r:embed="...PNG..."]
              / a:extLst
                / a:ext[@uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"]
                  / asvg:svgBlip[@r:embed="...SVG..."]
```

The exact URI profile is XML-token normalized for comparison while the raw
lexical attribute remains in the source bytes. A direct recognized extension
must have exactly one direct `asvg:svgBlip` child in the SVG namespace. Its
`r:embed` reference is the embedded profile; an `r:link` reference is reported
as linked/inert and is refused by attach and detach. Duplicate admitted
extensions, duplicate direct SVG children, missing raster fallback, malformed
recognized content, or a relationship that does not satisfy the package
profile are typed refusals. An unknown extension URI or foreign SVG-looking
child is preserved as opaque data and does not acquire a relationship meaning.

The first batch admits only a direct `pic:pic` in a direct `graphicData` under
one `wp:inline` or `wp:anchor`. It does not descend into a group, canvas,
legacy `w:pict`, `w:object`, or another nested drawing owner. An
`mc:AlternateContent` ancestor, or any candidate in a Choice/Fallback branch,
is recorded as MCE ancestry and makes the owner inert for read and refused for
attach/detach. The transaction must never append a direct owner beside a
hidden branch. This is an explicit refusal policy, not support for selecting
or evaluating MCE branches.

The raster relationship must be internal, owned by the selected story part,
resolve to `/word/media/`, and have a PNG content type for newly authored
attachment output. Existing non-PNG or otherwise malformed pictures are
reported as unsupported and left unchanged. The SVG relationship must be
internal, resolve to `/word/media/` with `image/svg+xml`, and target a media
part with no outbound relationships. The SVG payload itself remains opaque;
the host does not sanitize, render, or fetch it.

Placement is part of the read projection and is not rewritten by attach or
detach:

```rust,ignore
pub enum DrawingPlacement {
    Inline,
    Floating,
}
```

For a floating anchor, all geometry, wrapping, relative positioning,
`anchorId`, `editId`, and extension children are source-owned bytes. For an
inline anchor, extent, effect extent, nonvisual properties, and extension
children are likewise preserved. The SVG operation changes only the selected
`a:blip` extension region and package dependency closure.

## Strict and Transitional authoring policy

The host dialect and the SVG extension dialect are separate. Core Word,
Wordprocessing Drawing, DrawingML, and the physical image relationship type
follow the source package's Transitional or Strict profile. Newly authored
`asvg:svgBlip@r:embed` must use the Transitional Office relationship namespace
even in a Strict core host. The retained
[`svg-strict-namespace`](svg-strict-namespace/README.md) evidence proves this
against the unmodified MS-ODRAWXML §5.24 child schema using Strict PPTX
slides and a native-derived XLSX drawing: a Strict SVG-child relationship
attribute is rejected, while a Transitional relationship attribute passes.
The same extension rule applies here; generated DOCX core XML still requires
its own schema gate. The physical image relationship may use the host's
Strict relationship type.

If the source's `r` prefix is already bound to the Strict relationship URI,
author a local collision-free prefix/declaration for the SVG child rather
than rebinding the host's existing prefix. Existing source-only/no-op bytes
retain their original namespace spelling and relationship attribute form.
The shared codec's acceptance of both relationship namespaces is a read
compatibility capability, not permission to normalize a legacy extension.

## Source-backed view and ordinary API

The ordinary facade should select a picture by semantic story and source-order
ordinal, never by `rId`, package URI, XML offset, or prefix spelling. A
provisional API shape is:

```rust,ignore
pub enum StorySelector {
    Main,
    // Header/Footer/Footnotes/Endnotes/Comments/Glossary are later scope.
}

pub struct PictureSelector {
    pub story: StorySelector,
    pub drawing: usize,
    pub picture: usize,
}

pub struct SvgPictureView<'a> {
    pub placement: DrawingPlacement,
    pub svg: Option<SvgResourceView<'a>>,
    pub raster: RasterResourceView<'a>,
    // source-bound owner token remains private to the package transaction
}

pub fn svg_pictures(&self, story: StorySelector)
    -> Result<Vec<SvgPictureView<'_>>>;

pub fn attach_svg(
    &mut self,
    selector: PictureSelector,
    svg: SvgInput<'_>,
) -> Result<SvgPatch>;

pub fn detach_svg(
    &mut self,
    selector: PictureSelector,
) -> Result<SvgPatch>;
```

Names are provisional. The important contract is that the view borrows the
source-backed story and relationship graph, while the package owns the source
tokens and private ranges. `SvgInput` may borrow bytes for preflight or own an
`Arc<[u8]>` for a long-lived request. If an ownership-consuming attach form
fails, it returns the rejected payload as required by ADR 0007. Attach is
limited to an existing direct PNG picture with no admitted SVG owner. Detach
removes only the selected admitted SVG owner. SVG replacement, raster fallback
replacement, picture creation/reordering, and linked-to-embedded conversion
are separate verbs.

## Source scanner and splice contract

The new scanner should make one bounded raw pass over the selected story and
record, for each direct picture:

- the `wp:inline` or `wp:anchor` placement and complete source span;
- direct `pic:pic`, `pic:blipFill`, raster `a:blip`, `a:extLst`, admitted
  `a:ext`, and `asvg:svgBlip` spans;
- expanded names and the inherited namespace context needed to write a
  standalone child without changing unrelated lexical bytes;
- the raster and SVG relationship IDs and resolved package targets;
- MCE ancestry/branch state, direct-owner count, unknown extension spans, and
  the source fingerprint used by stale checks.

The scanner borrows source bytes and shares immutable namespace/scope layers;
it must not clone the whole document for every picture. It should reject
invalid XML names, duplicate expanded attributes, malformed recognized
containers, and resource-limit violations before storing decoded values.

Attach should precharge the complete `a:extLst/a:ext/asvg:svgBlip` fragment and
prospective story output before constructing replacement buffers. It should
preserve comments, whitespace, processing instructions, unknown attributes,
namespace declarations, and unrelated extension siblings. Detach should
remove the selected `a:ext`; it may remove the direct `a:extLst` wrapper only
when the wrapper has no attributes or namespace declarations and no unrelated
extension, comment, or payload. A wrapper with preserved metadata remains
byte-preserved. Source replacement must use the exact expected story bytes and
`OwnedXmlPart` provenance; a canonical rewrite of the complete document is
outside this batch.

## Package graph and dependency closure

An attach commit changes this closure atomically:

1. the owning story XML (`/word/document.xml` in the first batch);
2. that story's `_rels` part, if present, with one image relationship for the
   SVG media;
3. a new collision-free `/word/media/<name>.svg` leaf; and
4. `[Content_Types].xml` only when the existing `svg` default or exact SVG
   override does not already cover the new part.

The existing raster relationship and PNG part remain unchanged. The story is
the relationship owner; the package root and the main-document relationship
are not substitutes for the owner. Transitional and Strict image relationship
types are selected from the host dialect for the physical edge, while the
new SVG child attribute follows the Transitional extension rule above.

Detach first changes the selected story XML. Count remaining references to
the selected relationship ID throughout that complete story, including opaque
and non-picture content. If any remain, retain the story relationship and SVG
part. Otherwise remove the now-unused story edge while retaining the media
when a package-wide incoming-edge scan finds another relationship to it.
Remove the media only when no incoming edge remains. The SVG content-type
default/override is removed only when no retained part needs it. Source-token
relationship and content-type deltas must restore exact lexical bytes under
inverse; semantic equality alone is insufficient. External SVG links remain
inert and are never fetched or removed by this batch.

The candidate is reopened after staged XML, relationship, media, and
content-type changes. Readback must verify the selected inline/anchor and
picture still resolve, the raster fallback is unchanged, the requested SVG
owner and payload hash are present or absent, all changed relationships have
the expected type/target/mode, and no dangling edge, duplicate content type,
or orphaned media owner was introduced.

For source-backed publication, complete semantic readback may use the
OPC-owned effective candidate described in
[the prepared-topology contract](opc-prepared-topology-design.md). That view
must include the actual generated XML, effective relationships and content
types, complete Part membership, and lazy unchanged payloads that publication
will use. A projection of only the selected story and SVG is insufficient.
Physical ZIP publication is still tested by reopening its output and comparing
that graph with the effective candidate. This separates semantic validation
from ZIP framing checks without introducing an implicit whole-archive memory
buffer or weakening changed-closure validation.

## Snapshots, patches, and inverse

The edit records source fingerprints and immutable before-bytes for every
changed member, the story topology token, selected owner span/digest,
relationship topology, content-type binding, and SVG payload hash. Applying
it to a package with a different story source, relationship source,
content-types source, or graph lineage returns a stale/patch conflict before
publication.

An exact semantic no-op shares the source snapshot and performs no replacement
or candidate reopen. A changed commit publishes only after complete closure
validation and readback. Its in-memory inverse restores the exact accepted
story, relationship, content-type, and media state. A separately saved and
reopened package has no patch provenance; a fresh detach after reopen is a new
semantic edit and need not recover the original paired/self-closing syntax.

## Limits and allocation policy

The implementation should reuse `StoryLimits` and the shared SVG codec limits,
then add a host edit budget for:

- story XML bytes, element nodes, depth, attributes, namespace declarations,
  and direct-picture count;
- decoded relationship IDs, part names, content-type tokens, and package
  relationships;
- SVG media payload bytes and staged source/output bytes;
- changed member count and prospective story, relationship, and content-type
  output sizes; and
- aggregate graph traversal work for incoming-edge checks.

Every prospective replacement length must be checked before allocating its
temporary `Vec`; a final package-size check is insufficient for a small caller
limit. The scanner should use one source pass and immutable ancestor scope
layers so many pictures do not multiply namespace or source allocations.
SVG bytes remain opaque and are not parsed as an SVG document.

## Implementation and test sequence

1. **Direct picture scanner.** Add a source-backed main-story scanner owned by
   `crates/litchi-docx/src/drawing` (or a focused sibling module). Resolve
   expanded names, direct inline/anchor placement, direct picture grammar,
   source spans, MCE ancestry, and the native extension profile. Keep the
   existing lossy image/drawing readers unchanged for their current APIs.
2. **Deferred resource view.** Add a DOCX-owned SVG picture view that reuses
   `litchi-drawingml::svg_blip` and validates the story relationship closure,
   `/word/media` target, `image/svg+xml` content type, internal mode, and no
   outbound SVG-part relationships. Add PNG fallback checks for authoring.
3. **Source transaction.** Add `Package::svg_pictures`,
   `attach_svg`, and `detach_svg` (names provisional) in a package-owned
   transaction module. Reuse `story_inventory_with_limits`,
   `source_xml_part`, `source_relationships`, `OwnedXmlPart`,
   `OwnedRelationships`, `OwnedContentTypes`, and `edit_semantic_opc`.
   Implement collision-safe relationship/media names, source checks, atomic
   closure, inverse, and candidate readback.
4. **Schema and native gates.** Run the focused lifecycle suite on both
   retained fixtures, offline ECMA core and MS-ODRAWXML §5.24 schema checks,
   strict Clippy/rustdoc/formatting, stale/no-op/inverse checks, and bounded
   fresh-process allocation probes. Add subsidiary story support only after
   this main-story gate is frozen.

The focused test matrix should include:

| Case | Required result |
|---|---|
| Native `svg.docx` read | Direct inline picture exposes PNG fallback and embedded SVG metadata/payload. |
| Native `tdf164835_nonDummyLineHeight.docx` read | Direct floating anchor exposes the same graph while retaining anchor geometry. |
| Synthetic inline attach | Adds one admitted SVG owner, SVG edge/media leaf, and needed content type; source siblings remain exact. |
| Synthetic floating attach | Same closure while preserving wrapping, positioning, anchor IDs, and extension bytes. |
| Detach and shared-target case | Removes only the selected edge; shared SVG remains until the final incoming edge is gone. |
| Strict core host | Core and physical relationship dialect remain Strict while new SVG `r:embed` remains Transitional and both schema checks pass. |
| Unknown/duplicate/linked/MCE/malformed owner | Read inert where safe; attach/detach refuse without duplicate effective owners or data loss. |
| Opaque extension preservation | Unknown siblings, comments, namespace declarations, whitespace, and attributes survive attach/detach. |
| No-op, inverse, stale, and rejected input | No-op shares bytes; inverse restores exact in-memory state; stale and over-limit edits are atomic; rejected owned input is returned. |

## Deliberate exclusions and evidence limits

This batch deliberately excludes subsidiary story publication, even though
`StoryInventory` already validates their ownership. Headers, footers,
footnotes, endnotes, comments, and glossary drawings need separate selectors,
fixture coverage, and graph tests. It also excludes `mc:AlternateContent`
selection/evaluation, nested group/canvas pictures, VML/legacy pictures,
picture creation, raster replacement, SVG replacement, linked SVG fetching,
SVG sanitization/rendering, and physical Office acceptance.

The two native archives establish producer structure for direct inline and
floating-anchor owners, relationship direction, media content types, and the
observed GUID. Local ECMA/MS-ODRAWXML schema checks establish XML shape only.
Neither is evidence that newly authored files are accepted or rendered by
Word or LibreOffice. No performance or linear-runtime claim follows until the
source-backed scanner, graph planner, and candidate reopen are measured under
many-picture, inherited-namespace, shared-target, and small-output-limit
workloads.

No production source or test file is changed by this design artifact.

## Design review evidence

Root independently read both native archives and recomputed all archive and
member hashes listed above. The inline PNG hash was corrected to the actual
fixture bytes. The detach plan was also clarified to distinguish same-story
XML references, relationship-only removal, and final package-wide media
cleanup. This review is design and fixture evidence, not implemented DOCX
lifecycle behavior or a runtime-performance claim.
