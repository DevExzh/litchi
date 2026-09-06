# ODP streaming provider API contract (0437 draft)

Production is unchanged. This contract is the source for the pending
`streaming.rs`, `lib.patch`, and `facade.patch` drafts.

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

`S` may be `&str`, `String`, or `Cow<'_, str>`. The source is consumed once;
there is no retained `Vec<Slide>`. `title: Some("")` emits the title frame,
while `title: None` omits it. An empty body omits the body frame. CR, LF, tabs,
spaces, XML escapes, and the fixed Builder geometry/style/name contract follow
`Builder::add_slide_with_title` exactly; XML 1.0-invalid scalars are refused.
Rich shapes, notes, transitions, animations, media, page metadata, layouts,
declarations, and settings are outside this API and must not be silently
flattened.

`StreamingLimits` is provider-owned and finite. Its checked constructor fields
are, in order: `max_slides`, `max_title_text_bytes`, `max_body_text_bytes`,
`max_total_text_bytes`, `max_slide_xml_bytes`, `max_content_xml_bytes`,
`max_output_bytes`, and `XmlAuditLimits`. It exposes accessors for each field,
`required_memory_bytes()`, and `with_xml_audit_limits`. Defaults are within
ODP/common hard ceilings. The modeled memory reservation covers two fixed
content-shell copies, one reusable slide fragment, and fixed common metadata;
it does not claim all ZIP/auditor allocator peak.

`SlideStreamReport` exposes `slides`, `title_count`, `body_count`,
`title_text_bytes`, `body_text_bytes`, and authored `content_xml_bytes`.
`StreamingError` preserves invalid/producer/context/limit/counter failures and
typed publication failures with accepted sink bytes. The output sink is
caller-owned and may contain a partial poisoned ZIP after publication begins.
Discard it on any error.

The generated content envelope is the exact no-transition Builder root:
20 fixed namespace bindings, `office:scripts`, `office:font-face-decls`, the
balanced `dp1` drawing-page style inside `office:automatic-styles`, then
`office:body/office:presentation`; dynamic fragments are `draw:page` roots.
The provider calls the new opt-in
`GeneratedXmlEnvelope::try_new_with_prelude(CONTENT_PREFIX, CONTENT_SUFFIX)`;
the existing `try_new` contract remains unchanged.
