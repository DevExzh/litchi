# ODP creation data path audit (0437)

This is a source audit of the current ODP authoring and read paths.  No
production source was changed and no runtime, build, test, or profile command
was run for this audit.  The main sources are `crates/litchi-odp/src/authoring/builder.rs`,
`authoring/builder/package.rs`, `authoring/builder/xml.rs`,
`authoring/builder/validation.rs`, `authoring/builder/transition.rs`,
`model/{slide,page_metadata,page_layout,declaration,settings,media}.rs`,
`package/{presentation,source,catalog}.rs`, and the common writer/XML sources.

## Existing authoring path

| Stage | Owner and materialization | Validation and publication | Retained peak relevant to a large plain deck |
| --- | --- | --- | --- |
| Input | `Builder` owns `Vec<Slide>`, metadata, optional settings/declarations/page metadata, a page-layout collection, and a `BTreeMap<String, EmbeddedMedia>`. `add_slide_with_title` and `add_slide` copy every input `&str` into `String`s. A supplied `Slide` is moved into the vector, but its strings/vectors remain owned by the builder. | Most structural validation is deferred to `build`; some setters validate their collection immediately. | All slide title/body/notes strings, all shape trees and transitions, and all media payloads are retained before serialization. |
| Preflight | `build` calls `validation::validate(builder.snapshot())`. `Snapshot` is borrowed and does not clone. | Counts slides/media, validates layouts, settings, declarations, and page metadata. There is no aggregate content XML, total text, or output-size preflight. | The full model remains live for every subsequent phase. |
| `content.xml` | `generate_content_body` allocates one growing body `String`; the final `generate_content_xml` creates a second complete `String` with `format!`, copying the body into the final envelope. Per-slide page attributes, title/body paragraph strings, shape strings, notes, declaration fragments, and animation strings are temporary allocations. `effective_page_names` also allocates one owned name per slide. | Text is escaped and translated by `generate_text_paragraphs`: LF splits paragraphs, CR becomes `text:line-break`, tabs become `text:tab`, and repeated/edge spaces become `text:s`. Shape generation performs recursive checks. The common authored XML audit runs when the member is added. | Builder model + body + final content coexist. `content_xml` remains in scope while styles, meta, media, and finish are produced, so it is not released immediately after its ZIP member is written. |
| `styles.xml` | Starts with a complete `Structure::default_styles_xml()` `String`. Every custom layout calls `page_layout::set_xml`, which validates/scans the whole current XML and returns a new whole-document `String`; assignment temporarily overlaps old and new strings. | Layout limits and XML checks apply per collection/layout. The complete styles slice is then audited and deflated. | With layouts, repeated full-document copying is an independent cost. With no layouts, the fixed styles string is small. |
| `meta.xml` | `generate_meta_xml` builds one complete `String`, including only the metadata fields currently emitted by the builder (dates, title, author). | XML escaping and authored XML audit, then deflate. | At the call site content, styles, and meta strings are simultaneously in scope until the package function returns. |
| ZIP/package | `PackageWriter::new()` uses an in-memory `StreamingArchiveWriter<Cursor<Vec<u8>>>`. `add_file` borrows each XML slice while the ZIP compressor writes it, but the caller's strings remain alive. The writer retains manifest entries, manifest/member path sets, counters, and the archive output. Media bytes are borrowed from the builder map while compressed; the builder map remains live. | MIME is written first. Each member gets path/media-type/collision/manifest checks, bounded authored XML validation for XML parts, and ZIP entry limits. Manifest XML is generated and audited at finish. | Final archive bytes are retained in the writer's output `Vec`; this is additional to the still-live builder and XML strings. Finish adds a manifest `String` and central-directory/finalization staging. |

The current builder therefore has a full model-first path even for the simple
`titled slide + body` case.  The final archive is also returned as a `Vec<u8>`;
there is no caller-owned sequential sink in `Builder::build`.  The common
`PackageWriter<W>` can use a caller sink, but the existing builder does not use
that generic form.

`PackageWriter::set_mimetype_streaming` publishes the stored `mimetype` member
before content.  `add_generated_xml` admits the XML ZIP entry before its first
producer callback is consumed; `add_file` similarly starts publication while
the borrowed source slice is still live.  A future source producer must treat
any producer, context, audit, or sink failure after admission as a partial
package and report accepted output bytes.  `finish_to_writer` writes the
manifest only after all payload entries, then finalizes the ZIP central
records.

## Validation and repeated work

The top-level validation limits the slide and media maps to 65,536 and invokes
layout, settings, declaration, and page-metadata validators.  Serialization
repeats some of that work:

* `write_declaration_elements` validates the complete declaration collection
  again. `write_binding_attributes` then performs per-slide lookups for slide
  and notes bindings.
* `effective_page_names` validates page metadata and allocates a complete
  `Vec<String>` of names. Inside the slide loop, `metadata.validate_for_slides`
  is called again for every slide when page metadata is present. That is an
  O(slides × metadata-pages) validation path in the worst case.
* `validate_page_references` validates settings again after page names are
  built. The top-level settings validation has already run.
* Animation roots and extension namespaces are scanned over every slide before
  body generation. Transition styles are scanned over every slide again.
* The common XML audit scans each complete XML member after it has already been
  fully generated. ZIP compression then reads the same member again.

These are correctness checks, not evidence of a performance regression by
themselves. They are material to a bounded design because all of them currently
happen after the model and large XML strings exist.

## Current hard limits

The following limits are explicit in the inspected sources:

* Builder validation allows at most 65,536 slides and 65,536 embedded media
  files. Page metadata allows at most 65,536 pages. Declarations allow 65,536
  headers, 65,536 footers, and 65,536 date-time declarations, with 131,072
  bindings. Settings allow 65,536 custom shows and 65,536 pages per show.
* Page layouts allow 65,536 layouts, 4,096 placeholders per layout, 65,536
  bytes per layout value, 16 MiB aggregate layout name/display bytes, 256
  nesting levels, and an 64 MiB XML input. Page metadata/settings/declaration
  parsed XML inputs each have an 8 MiB cap; their individual text fields have
  a 1 MiB cap.
* Recursive shape emission rejects depth greater than 64 and more than 65,536
  nodes in one `generate_shape_xml` call. `generate_shape_xml` resets the node
  counter for each top-level shape, so this is not an aggregate document shape
  ceiling. Enhanced geometry has a 65,536-child/equation-style bound. There is
  no aggregate shape, body-text, or generated-content ceiling in the builder
  validation pass.
* Common family content validation caps materialized `content.xml` at
  256 MiB. The XML minifier hard ceilings used by common generated XML are
  256 MiB bytes, depth 4,096, 4,000,000 events, 1,000,000 attributes,
  64 MiB for one token, and 256 MiB aggregate character data. Provider limits
  must be narrower where the source contract needs a smaller window.
* The default streaming ZIP transport limits are 65,534 entries, 512 MiB
  compressed bytes per entry, 512 MiB uncompressed bytes per entry, 2 GiB
  aggregate uncompressed bytes, and 512 MiB complete output bytes. ODF
  manifest accounting and member-name/metadata limits are checked separately
  by `PackageWriter`.
* `GeneratedXmlReader` reserves one reusable `max_fragment_bytes` buffer,
  rejects zero or over-document fragment capacities, and keeps that capacity
  for the operation. Its current envelope constructor owns copies of prefix
  and suffix and audits a concatenated shell.
* The source-backed read facade has optional semantic projection thresholds of
  4 MiB `content.xml` / 256 slides for the slide cache and 16 MiB for text
  cache. These are cache-retention policies, not authoring limits. The
  catalog-first facade retains only bounded slide descriptors; selected slide
  parsing rereads content and styles.

The common source defaults also cap physical source bytes at 2 GiB, manifest
bytes at 16 MiB, and mimetype bytes at 4 KiB. Those are input/read limits and
do not prevent the existing authoring builder from creating a large in-memory
archive before reopening it.

## Read/reopen ownership

The ordinary `Presentation::from_bytes(Vec<u8>)` path retains the complete ZIP
input in the common `OwnedPackage` (the archive index points into that `Arc`/`Vec`).
Opening `content.xml` calls `get_file`, which materializes a decompressed `Vec`,
then `Content::from_bytes` copies its UTF-8 bytes into an owned `String`; the
same two-stage copy occurs for `styles.xml`. Metadata parsing has its own
materialized member and retained projection. The temporary decompressed member
vectors drop after construction, while the original archive bytes and decoded
XML strings remain. Every ordinary `slides()` query parses into a fresh
`Vec<Slide>` and fresh model strings; repeated queries do not turn the ordinary
facade into a bounded streaming projection.

`SourceBackedPresentation` retains a positional `SourceBackedPackage` index and
the source handle, then materializes and retains `content_xml` and optional
`styles_xml` strings. It does not retain media payloads until `media_data` is
called. Its optional slide/text caches add cloned semantic models only under
the explicit thresholds above. `SourceBackedPresentationCatalog` retains the
source/index and an `Arc<[SlideCatalogEntry]>` of positions and optional
`draw:name` values, not content/styles/metadata/slide models/media. This is the
read-side ownership model a future authoring source should emulate: one source
item and one bounded XML fragment at a time, with source/context checks around
all pull and publication boundaries.

## Immediate large-deck benchmark risk

The following is a source-derived risk assessment, not a measured result.  The
current path scales with all of the following at once: retained slide model,
per-slide page-name vector, growing body XML, copied final content XML, ZIP
archive output, and repeated audit/compression passes.

| Deck | Expected risk before measurement | Main mechanism |
| --- | --- | --- |
| 64 slides | A useful smoke baseline; fixed envelope, page names, and compressor overhead dominate. | Confirms semantic output and establishes the plain-lane control. |
| 8,192 slides | Multi-megabyte transient strings and repeated O(N) validation/name work are expected even with tiny text. Allocation growth and XML audit/compression become visible. | Model + body + final content + archive coexist; page metadata, if enabled, adds repeated validation. |
| 32,768 slides | High memory/latency risk. Depending on text and ZIP compression, the generated content or final archive can approach common 256 MiB/512 MiB ceilings; failure would occur after substantial staging because the builder has no aggregate preflight. | Full model, `Vec<String>` names, body/final copies, archive output, and audit/compressor staging all overlap. |

The current body capacity estimate (`slides * 128` plus input text and shallow
shape text) is only a growth hint. Escaping, paragraph/control elements,
nested shapes, declarations, page attributes, and transitions can exceed it;
there is no hard cap tied to the estimate.  A benchmark must therefore record
both output size and peak retained/temporary memory. It should not infer a
bounded-streaming claim from a successful 64-slide run or from the 65,536-slide
structural ceiling.

## Blocking compatibility observation for a future generated content member

A plain ODP content document cannot use `GeneratedXmlEnvelope::try_new` as-is
while preserving the current Builder grammar. The current content shell is:

```xml
<office:document-content>
  <office:scripts/>
  <office:font-face-decls/>
  <office:automatic-styles>
    <style:style style:name="dp1" style:family="drawing-page">
      <style:drawing-page-properties/>
    </style:style>
  </office:automatic-styles>
  <office:body><office:presentation>
    <!-- one or more draw:page elements -->
  </office:presentation></office:body>
</office:document-content>
```

`GeneratedXmlEnvelope::try_new(prefix, suffix)` permits only a declaration and
open elements in `prefix`, and only matching end tags in `suffix`. The fixed
`dp1` child must close `office:automatic-styles` before the fixed body and
presentation starts, while dynamic page fragments must be inserted after that
point. Moving the body opening into `suffix` would put start tags in the suffix;
putting the balanced `dp1` child in the current prefix would put end tags in
the prefix. A first fragment containing `automatic-styles` cannot work either:
subsequent page fragments would be siblings after that fragment's root, which
violates the one-root fragment contract.

The existing narrow constructor should keep rejecting such shells. An ODP
provider needs an opt-in common envelope seam that explicitly validates a
balanced fixed-child sequence before the final open insertion path, emits
`prefix + fixed children/body opening + fragments + suffix`, and includes the
fixed bytes/events/attributes/depth in the composed audit counters. The seam
must preserve the current prefix/suffix contract for other providers; relaxing
`try_new` in place would weaken a useful XML-shape invariant.
