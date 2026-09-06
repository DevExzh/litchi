# Bounded ODP plain-slide source design (0437)

This is a future API design only.  It is deliberately not applied to the
production tree and has not been compiled or run.  The goal is a bounded
sequential publisher for the exact simple shape already emitted by
`Builder::add_slide_with_title` / `Builder::add_slide`, while refusing rich
model fields instead of flattening or silently dropping them.

## Proposed public contract

Use an ODP-owned API whose public error/report types do not expose ZIP or common
archive implementation types:

```text
try_stream_plain_slides_to(sink, source, context, limits) ->
    Result<PlainSlideWriteReport, PlainSlideWriteError>

PlainSlide = {
    title: Option<Cow<str>>,
    body: Cow<str>,
}
```

The source is a finite, one-pass, fallible iterator/producer.  The concrete
Rust signature can use a provider-owned source trait if a generic iterator
cannot express a borrowed `Cow` lifetime cleanly; the required behavior is
`next -> Result<Option<PlainSlide>>`, with source errors retained as typed
causes.  `&str`, `String`, and `Cow<str>` inputs should be accepted without
forcing an owning model.  `title: Some("")` remains distinct from `None`, as
it is in the current `Slide` model.  An empty body is allowed and emits no
body frame, matching the builder.

`PlainSlide` is intentionally narrower than `Slide`.  It has no notes,
transition, animation, custom shape, media, page-layout, page metadata,
declaration, settings, hyperlink, or arbitrary geometry/style field.  A caller
with any of those fields must receive a typed unsupported-input/refusal result
from an adapter rather than silently losing them.  A future richer provider can
add a separate contract after its own bounded XML grammar is proven.

The operation accepts an `ExecutionContext` and a caller-owned sequential
`Write` sink.  The provider checks cancellation before pulling a source item,
before each fragment/XML write, and after each sink operation.  It reserves
memory for the fixed shell, one reusable slide-fragment window, text escaping
scratch, manifest bookkeeping, and the ZIP staging bound; the exact formula
must be tied to the common seam and transport limits rather than advertised as
an allocator-wide peak.  Every input/source charge uses checked byte/object
arithmetic.  Every sink write is wrapped by the provider's output budget and
reports the sink's accepted count, including partial writes.

The limits object should make all finite ceilings explicit:

* `max_slides` (at most the existing 65,536 structural ceiling);
* `max_title_bytes`, `max_body_bytes`, and `max_total_text_bytes` for decoded
  UTF-8 input;
* `max_slide_xml_bytes` / `max_fragment_bytes` for one reusable page fragment;
* `max_content_xml_bytes` for the composed `content.xml` (never above the
  common 256 MiB family limit);
* finite output bytes and ZIP entry/aggregate limits;
* finite memory/window bytes; and
* the provider's `GeneratedXmlLimits` profile (bytes, depth, events,
  attributes, token bytes, and text bytes).

The constructor should reject inconsistent limits before MIME publication where
possible.  A source that emits one item after `max_slides` must be rejected as
an extra-item/limit failure; it must not be silently truncated.  If a failure
occurs after `mimetype` or a content local header has been accepted, the output
is partial and must be discarded by the caller.  The typed error/report carries
`accepted_output_bytes` and the source/context/XML/ZIP cause.  A successful
report carries emitted slides and checked input/content/output counts.

## Exact simple-slide XML semantics

For each source item, emit one `draw:page` fragment.  The provider must retain
the current builder's fixed page and frame semantics:

* With no page metadata, the page has `draw:name="page{index + 1}"`,
  `draw:style-name="dp1"`, and `draw:master-page-name="Default"`.  The plain
  API has no custom page metadata field, so it uses these deterministic names;
  an adapter that has explicit `draw:name`/ID/layout/master metadata refuses
  rather than discards it.
* A `Some(title)` emits the exact title frame attributes currently hard-coded
  by `generate_content_body`: `draw:style-name="gr1"`,
  `draw:text-style-name="P1"`, `draw:layer="layout"`,
  `presentation:class="title"`, width `25.199cm`, height `3.506cm`, x
  `1.4cm`, y `0.962cm`, followed by a `draw:text-box` with P1 paragraphs.
* A nonempty body emits the exact body frame attributes:
  `draw:style-name="gr2"`, `draw:text-style-name="P2"`,
  `draw:layer="layout"`, `presentation:class="object"`, width `25.199cm`,
  height `10cm`, x `1.4cm`, and y `5.0cm` when a title exists or `2.0cm`
  otherwise.  It uses a P2 text box.  An empty body emits no body frame.
* Text encoding follows `authoring/builder/xml.rs` exactly for the supported
  valid-text subset: LF splits into separate `text:p` elements; CR becomes
  `text:line-break`; tab becomes `text:tab`; one interior space remains text;
  repeated and leading/trailing spaces become `text:s` with the same `text:c`
  behavior; XML special characters are escaped.  The provider must validate
  XML 1.0 characters before publication and reject prohibited controls with a
  typed source/content error rather than relying on common lexical audit to
  make the input safe.
* A slide always emits its `draw:page`, including a slide with no title and an
  empty body.  It emits no notes, transitions, animations, declarations,
  settings, custom shapes, or media because those values are outside this
  contract.

The output is semantically compatible with the current simple builder lane;
byte-for-byte archive identity is a separate acceptance criterion.  The
provider should compare reopened slides/text and XML fragment shape in tests,
then explicitly record any permitted archive-level differences (streaming ZIP
headers, member order only if the package contract permits it, or static
metadata construction).  It must not claim lexical equality without a fixture
check.

## Package assembly and ownership

The provider should use a caller-owned `PackageWriter<W>` with bounded ZIP
limits, not `Builder::build`:

1. Validate the source/limits/memory reservation and fixed static package
   grammar as far as possible before publication.
2. Call `set_mimetype_streaming`.  This writes `mimetype` first and establishes
   the manifest root entry.
3. Publish `content.xml` with a bounded generated-member callback, one page
   fragment at a time, keeping the existing member order.  Pull the next source
   only after the previous fragment has passed its provider checks and common
   lexical audit.  Keep a reusable page buffer, but never retain the complete
   body or complete content member.
4. Publish `styles.xml` and `meta.xml` through the existing typed XML writer
   methods using fixed, borrowed slices where possible.  A future implementation
   should avoid turning `Structure::default_styles_xml()` into an avoidable
   large owned intermediate; these fixed parts are small, but the ownership
   should be explicit.  Current default `styles.xml` has empty automatic styles;
   the `dp1` page style belongs to the current content.xml automatic-styles
   section, so moving it silently to styles.xml is not an exact-semantics
   substitution.
5. Finish the package, which creates the manifest and closes the ZIP.  Release
   the memory reservation only after finish succeeds or the poisoned writer is
   dropped.

The common `add_generated_xml` seam is useful because it owns path/media type,
manifest collision, XML classification, fragment capacity, per-fragment audit,
composed byte/event/depth counters, and ZIP reader staging.  It does not
validate ODP namespaces/schema or perform context/output accounting, so the
provider remains responsible for fixed grammar, `ExecutionContext`, source
lifetime, and accepted output bytes.  The callback runs after content entry
admission; that boundary must remain visible in the provider's typed partial
progress error.

## Required common seam extension

Current `GeneratedXmlEnvelope::try_new(prefix, suffix)` cannot represent the
ODP content shell while preserving the builder's `dp1` style.  It only permits
start/declaration events before its insertion boundary and end-tag events after
it.  ODP needs this event sequence before dynamic page fragments:

```text
open document-content
  fixed empty scripts and font-face-decls
  fixed automatic-styles containing balanced dp1 style
close automatic-styles
open body/presentation
<dynamic draw:page fragments>
close presentation/body/document-content
```

A narrow opt-in common constructor should accept a validated fixed prefix
sequence containing balanced child elements followed by the final open
insertion path, and a suffix that closes that path.  It should emit
`fixed-prefix + fragments + suffix`, retain bounded copies of the fixed shell,
and include all fixed bytes/events/attributes/text and the deepest fixed or
fragment nesting in the single composed report.  The old `try_new` contract
should remain unchanged and continue to reject end tags in its prefix; the new
constructor should not become a general unvalidated XML concatenator.  The
provider must still bind all namespaces and ensure that every page fragment is
an ODP `draw:page` root with no unsupported children.

If common chooses a more general “balanced fixed children around insertion”
representation, the same invariants are required: event-boundary checks,
matching namespace-aware lexical tags, no inherited `xml:space`, checked
shell+fragment counters, and no full composed-member allocation.  A design that
only moves `automatic-styles` into the first callback fragment is invalid,
because later page fragments would become sibling roots outside the one-root
fragment contract.

## Failure and budget boundaries

The source producer must be wrapped so that it cannot bypass the context or
output contract:

* Check context and reserve/charge input bytes and one source object before
  pulling/encoding a slide.  Validate title/body UTF-8 scalar policy and all
  per-slide limits before writing its first page byte.
* Charge generated XML work for fixed page markup, escaped text/control
  elements, and fragment bytes with checked arithmetic.  If a common audit,
  output, context, or source check fails after sink acceptance, preserve the
  accepted count and typed cause; do not replace it with an untyped ZIP error.
* Enforce per-fragment, aggregate content, and output limits before the next
  write where possible, and again after accepted writes.  A failed operation
  may have emitted a prefix of one page; the report must describe that partial
  progress, and the package must not be treated as finalized.
* Do not expose a public “rich slide” adapter that serializes only the title and
  body from a `Slide` while dropping notes/shapes/transitions/media.  Such an
  adapter should inspect the source model and return an unsupported-field
  refusal before any output if exact conversion is promised.

## 64 / 8,192 / 32,768 slide benchmark implications

The baseline should measure the existing builder before applying this design.
For tiny identical title/body strings, the existing path still retains one
`Slide` per item, one generated page name, a growing body string, a copied final
content string, and the archive output. At 64 slides it is mainly a semantic
smoke/reference case. At 8,192 slides, O(N) page-name/validation work and the
body/final/archive overlap should expose allocation and audit costs. At 32,768
slides, the content member and output may approach the 256 MiB family and
512 MiB transport ceilings depending on text; the existing builder has no
aggregate preflight, so it can fail only after large allocations and partial
ZIP staging. The future source API is intended to hold O(window + fixed
metadata + ZIP staging) generated XML while retaining only source state, but
that claim must be made only after peak-memory/output measurements and a
reopen semantic oracle.
