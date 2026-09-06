# ODP plain-slide streaming provider draft

This directory contains a source-only provider draft for the fresh plain-slide
ODP lane. It is intentionally outside the production tree. `streaming.rs` is
the proposed `crates/litchi-odp/src/streaming.rs`; `lib.patch` and
`facade.patch` show the only public exports.

The public contract is also recorded in `api.md`:

```rust
pub struct PlainSlide<S> {
    pub title: Option<S>,
    pub body: S,
}

pub fn stream_plain_slides_to<W, I, S>(
    output: &mut W,
    slides: I,
    context: &ExecutionContext,
    limits: StreamingLimits,
) -> Result<SlideStreamReport, StreamingError>
where
    W: Write + ?Sized,
    I: IntoIterator<Item = PlainSlide<S>>,
    S: AsRef<str>,
{
    /* implementation omitted */
}

pub fn try_stream_plain_slides_to<W, I, S>(
    output: &mut W,
    slides: I,
    context: &ExecutionContext,
    limits: StreamingLimits,
) -> Result<SlideStreamReport, StreamingError>
where
    W: Write + ?Sized,
    I: IntoIterator<Item = litchi_core::Result<PlainSlide<S>>>,
    S: AsRef<str>,
{
    /* implementation omitted */
}
```

The source is consumed once. It keeps a single reusable bounded `draw:page`
fragment and sends all ZIP output to the caller sink. It does not construct a
`Presentation`, `Builder`, slide model, whole `content.xml`, or complete ZIP
buffer. The first source/validation failure after the mimetype header poisons
the partial package and the error reports the accepted sink prefix.

The fixed content prelude is the Builder's plain no-transition grammar: the
20 namespace bindings, scripts and font-face declarations, one `dp1`
drawing-page style, and open body/presentation ancestors. The common
`GeneratedXmlEnvelope::try_new_with_prelude` seam is required for that
balanced fixed child before the callback insertion point. The old strict
two-part envelope constructor remains unchanged. The suffix closes
presentation, body, and document-content.

Each page uses the Builder defaults `pageN`, `dp1`, and `Default`. A present
title, including `Some("")`, emits the exact title frame and one P1 paragraph.
An empty body emits no body frame; a nonempty body emits the exact default body
frame and P2 paragraphs. LF splits paragraphs, CR emits `text:line-break`, tabs
emit `text:tab`, and leading/trailing/repeated spaces use `text:s` while XML
special characters use the Builder's entity spelling. Invalid XML 1.0 scalar
values are refused before that slide is admitted. No rich slide field is
flattened into this API.

The static auxiliary parts retain the Builder profile and member order:

```
mimetype, content.xml, styles.xml, meta.xml, META-INF/manifest.xml
```

`styles.xml` is the common `Structure::default_styles_xml()` byte grammar,
pinned by the existing ODP evidence at 1,960 bytes and SHA-256
`d9881e91085516246a19c30d9e5cde39a8b10d7e42120b135f48f5ca8afef8d2`. The
fixed ODP metadata is 387 bytes with SHA-256
`c7e55a3560c73aa42da85eec4751c3e78b5cc53ff964f50acba6c5cd105e6719`.
Independent harnesses should keep those byte/hash checks when this draft is
applied.

`StreamingLimits` bounds slide count, title/body/aggregate UTF-8 bytes, one
slide fragment, composed content XML, ZIP output, and the common XML lexical
audit. Defaults are finite and reconstructible through `StreamingLimits::new`:
the default aggregate text and XML audit text ceilings are 16 MiB, and the
default content XML ceiling is 32 MiB to fit the common 32 MiB audit-byte
profile. `required_memory_bytes()` covers the reusable slide window, two fixed
shell copies, and the bounded metadata reservation; it is a modeled provider
reservation and does not claim the complete ZIP/auditor allocator peak.

`StreamingError` distinguishes invalid XML/source data, execution/cancellation,
fallible producer errors, finite limits/counter overflow, and publication
failures. The `written` field on each applicable variant (and
`PublicationError::written()`) is the accepted caller-sink count maintained by
`BudgetedOutput`; no error's progress is inferred from an underlying attempted
write length.

Before applying, verify the common prelude method name/signature, the authored
XML method, and the exact default auxiliary-part hashes against the frozen
common/ODP baseline. The draft has not been compiled or executed in this
workspace per the source-only review boundary.
