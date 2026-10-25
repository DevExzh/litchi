//! Borrowed semantic shape views.

use std::borrow::Cow;
use std::ops::Range;

use litchi_drawingml::text::body::Body;
use litchi_drawingml::text::body::reader::{self as text_reader};
use litchi_drawingml::transform::Transform;
use litchi_ooxml_common::mce::InScopeNamespaces;
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use quick_xml::XmlVersion;
use quick_xml::events::attributes::Attribute;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, PrefixDeclaration, QName};

use crate::{Error, Result};

/// A checked byte span in the markup-compatibility-processed owner XML.
///
/// Spans are compact so a large scene can retain its index without copying
/// individual shape subtrees. Use [`Scene::xml`](super::Scene::xml) to obtain
/// the owner against which this span is defined: `&scene.xml()[start..end]`
/// is the element. Offsets count every byte of that owner, including a
/// leading UTF-8 byte-order mark, which the owner keeps whenever markup
/// compatibility processing did not rewrite it (change 0765; before it,
/// spans of a marked owner were three bytes early).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Span {
    pub(super) start: u32,
    pub(super) len: u32,
}

impl Span {
    /// Byte offset from the beginning of the processed owner XML, counting a
    /// leading byte-order mark.
    #[inline]
    #[must_use]
    pub const fn start(self) -> u32 {
        self.start
    }

    /// Length of this element, including its start and end tags.
    #[inline]
    #[must_use]
    pub const fn len(self) -> u32 {
        self.len
    }

    /// Whether the element occupies no bytes.
    #[inline]
    #[must_use]
    pub const fn is_empty(self) -> bool {
        self.len == 0
    }

    /// Exclusive byte end, checked for arithmetic overflow.
    #[inline]
    #[must_use]
    pub const fn end(self) -> Option<u32> {
        self.start.checked_add(self.len)
    }

    pub(crate) fn range(self, owner_len: usize) -> Result<Range<usize>> {
        let start = usize::try_from(self.start)
            .map_err(|_err| Error::Invalid("shape offset does not fit usize".into()))?;
        let len = usize::try_from(self.len)
            .map_err(|_err| Error::Invalid("shape length does not fit usize".into()))?;
        let end = start
            .checked_add(len)
            .ok_or_else(|| Error::Invalid("shape byte span overflows usize".into()))?;
        if end > owner_len {
            return Err(Error::Invalid(
                "shape byte span is outside its processed owner XML".into(),
            ));
        }
        Ok(start..end)
    }
}

/// Local shape bounds in English Metric Units (EMUs).
///
/// A value is exposed only when one complete offset/extent pair is present;
/// inherited layout transforms are deliberately not guessed.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Bounds {
    x: i64,
    y: i64,
    width: i64,
    height: i64,
}

impl Bounds {
    #[inline]
    pub(super) const fn new(x: i64, y: i64, width: i64, height: i64) -> Self {
        Self {
            x,
            y,
            width,
            height,
        }
    }

    /// Left edge in EMUs.
    #[inline]
    #[must_use]
    pub const fn x(self) -> i64 {
        self.x
    }

    /// Top edge in EMUs.
    #[inline]
    #[must_use]
    pub const fn y(self) -> i64 {
        self.y
    }

    /// Width in EMUs.
    #[inline]
    #[must_use]
    pub const fn width(self) -> i64 {
        self.width
    }

    /// Height in EMUs.
    #[inline]
    #[must_use]
    pub const fn height(self) -> i64 {
        self.height
    }
}

#[derive(Debug, Clone, Copy)]
pub(super) struct TextSpan {
    pub(super) start: u32,
    pub(super) len: u32,
}

impl TextSpan {
    pub(super) fn get(self, arena: &str) -> Option<&str> {
        let start = self.start as usize;
        let len = self.len as usize;
        let end = start.checked_add(len)?;
        arena.get(start..end)
    }
}

#[derive(Debug, Clone, Copy)]
pub(super) struct PlaceholderRecord {
    pub(super) kind: Option<TextSpan>,
    pub(super) index: u32,
    pub(super) type_extension: Option<PlaceholderTypeExtension>,
}

/// PowerPoint 2023 placeholder type extension (`p232:phTypeExt`).
///
/// This is the stable owner URI used by this crate for the generic
/// `p:CT_Extension` entry that contains the p232 payload. The local p232
/// schema defines the child namespace and grammar but does not assign a
/// separate extension GUID. The extension is metadata only; it does not imply
/// that the library can render or otherwise interpret the placeholder
/// payload.
pub const PLACEHOLDER_TYPE_EXTENSION_URI: &str = "urn:litchi:pptx:p232:phTypeExt";

/// The schema-defined alternatives of `p232:phTypeExt/p232:type`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum PlaceholderTypeExtension {
    /// A cameo placeholder (`p232:type/p232:cameo`).
    Cameo,
    /// An explicitly unknown placeholder type (`p232:type/p232:unknown`).
    Unknown,
}

/// Placeholder metadata declared by one shape.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Placeholder<'a> {
    kind: Option<&'a str>,
    index: u32,
    type_extension: Option<PlaceholderTypeExtension>,
}

impl<'a> Placeholder<'a> {
    /// Producer token such as `title`, `body`, or `ctrTitle`.
    ///
    /// `None` preserves the schema-defaulted spelling instead of inventing a
    /// token that was not present in the owner XML.
    #[inline]
    #[must_use]
    pub const fn kind(self) -> Option<&'a str> {
        self.kind
    }

    /// Placeholder index. OOXML defaults an omitted index to zero.
    #[inline]
    #[must_use]
    pub const fn index(self) -> u32 {
        self.index
    }

    /// PowerPoint 2023 placeholder type extension, when present and valid.
    #[inline]
    #[must_use]
    pub const fn type_extension(self) -> Option<PlaceholderTypeExtension> {
        self.type_extension
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum Kind {
    Auto,
    Picture,
    Table,
    Chart,
    Diagram,
    Ole,
    Frame,
    Group,
    Connector,
    Content,
    Unknown,
}

/// The relationship attribute that names a shape's payload part.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum PayloadAttribute {
    /// An embedded picture: `a:blip/@r:embed`.
    Embed,
    /// A linked picture: `a:blip/@r:link`, when the blip embeds nothing.
    Link,
    /// A chart (`c:chart/@r:id`), an OLE object (`p:oleObj/@r:id`) or a
    /// content part (`p:contentPart/@r:id`).
    Id,
    /// A diagram's data model: `dgm:relIds/@r:dm`.
    DataModel,
}

#[derive(Debug, Clone, Copy)]
pub(super) struct PayloadRecord {
    /// The attribute value exactly as written, between its quotes.
    pub(super) value: Span,
    pub(super) attribute: PayloadAttribute,
    pub(super) error: Option<super::raw::Error>,
}

#[derive(Debug, Clone, Copy)]
pub(super) struct Record {
    pub(super) span: Span,
    pub(super) subtree_end: u32,
    pub(super) parent: Option<u32>,
    pub(super) kind: Kind,
    /// Whether a local root/shape namespace fast path is sufficient. Groups
    /// and mid-tree declarations disable it and use the bounded ancestor walk
    /// instead.
    pub(super) inherits_root_namespaces: bool,
    pub(super) name: Option<TextSpan>,
    pub(super) id: Option<u32>,
    pub(super) bounds: Option<Bounds>,
    pub(super) placeholder: Option<PlaceholderRecord>,
    pub(super) text: Option<TextSpan>,
    pub(super) source_name: Option<TextSpan>,
    /// The start tag of the shape's `cNvPr`.
    pub(super) non_visual: Option<Span>,
    /// The shape's own `a:xfrm` or `p:xfrm` element.
    pub(super) transform: Option<Span>,
    pub(super) transform_invalid: bool,
    /// The relationship naming the shape's payload part.
    pub(super) payload: Option<PayloadRecord>,
    /// The shape's direct `p:txBody` element.
    pub(super) text_body: Option<Span>,
    /// More than one direct `p:txBody` was observed.
    pub(super) text_body_invalid: bool,
}

/// Non-visual drawing properties a shape declares on its `cNvPr`:
/// alternative text, title and the hidden flag, as written.
///
/// Values are decoded from the shape's start tag when asked for; an absent
/// attribute is `None`.
#[derive(Clone, Default, PartialEq, Eq)]
pub struct NonVisual<'a> {
    description: Option<Cow<'a, str>>,
    title: Option<Cow<'a, str>>,
    hidden: Option<bool>,
}

impl<'a> NonVisual<'a> {
    /// Alternative text (`descr`).
    #[inline]
    #[must_use]
    pub fn description(&self) -> Option<&str> {
        self.description.as_deref()
    }

    /// Title (`title`).
    #[inline]
    #[must_use]
    pub fn title(&self) -> Option<&str> {
        self.title.as_deref()
    }

    /// The hidden flag (`hidden`) as written: `Some(false)` for an explicit
    /// `hidden="0"`, `None` when the attribute is absent.
    #[inline]
    #[must_use]
    pub const fn hidden(&self) -> Option<bool> {
        self.hidden
    }

    /// Move out the alternative text, title and hidden flag; the text
    /// borrows the owner XML where no decoding was needed.
    #[must_use]
    pub fn into_parts(self) -> (Option<Cow<'a, str>>, Option<Cow<'a, str>>, Option<bool>) {
        (self.description, self.title, self.hidden)
    }
}

/// The relationship through which a shape reaches its payload part: a
/// picture's image, a chart, a diagram's data model, an OLE object or a
/// content part.
#[derive(Clone, Copy, PartialEq, Eq)]
pub struct Payload<'a> {
    raw: &'a [u8],
    span: Span,
    attribute: PayloadAttribute,
}

impl<'a> Payload<'a> {
    /// The attribute that names the relationship.
    #[inline]
    #[must_use]
    pub const fn attribute(self) -> PayloadAttribute {
        self.attribute
    }

    /// The attribute value's span in the processed owner XML, between its
    /// quotes.
    #[inline]
    #[must_use]
    pub const fn span(self) -> Span {
        self.span
    }

    /// The relationship identifier, decoded.
    ///
    /// # Errors
    ///
    /// Returns [`super::raw::Error::InvalidAttribute`] when the value is not UTF-8 or holds a
    /// malformed character reference.
    pub fn relationship_id(self) -> std::result::Result<Cow<'a, str>, super::raw::Error> {
        decode_attribute_value(self.raw).map_err(|_| super::raw::Error::InvalidAttribute)
    }
}

/// Borrowed common properties shared by every semantic shape variant.
#[derive(Clone, Copy)]
pub struct Common<'a> {
    pub(super) xml: &'a [u8],
    pub(super) records: &'a [Record],
    pub(super) strings: &'a str,
    pub(super) limits: super::reader::Limits,
    pub(super) record: &'a Record,
    pub(super) index: usize,
}

impl<'a> Common<'a> {
    /// Zero-based pre-order position in the owning scene.
    #[inline]
    #[must_use]
    pub const fn index(self) -> usize {
        self.index
    }

    /// Checked span in the processed owner XML.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn span(self) -> Result<Span> {
        Ok(self.record.span)
    }

    /// Exact borrowed element XML without a per-shape copy.
    ///
    /// The returned bytes are an inner fragment of
    /// [`Scene::xml`](super::Scene::xml), reproduced exactly. The namespace
    /// declarations that govern it are the ones in scope at its position in
    /// that processed owner XML, and the markup-compatibility codec declares
    /// each namespace once, where XML requires it, so a binding an ancestor
    /// makes is not necessarily repeated on the fragment's own root element.
    /// Parsing the fragment on its own can therefore leave its prefixes
    /// unresolved; a caller that does so should use
    /// [`self_contained_xml`](Self::self_contained_xml) instead. Use this
    /// accessor to read or copy the shape's exact source bytes, and whenever
    /// the surrounding owner XML is parsed along with them.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn xml(self) -> Result<&'a [u8]> {
        let range = self.record.span.range(self.xml.len())?;
        self.xml
            .get(range)
            .ok_or_else(|| Error::Invalid("shape XML span became invalid".into()))
    }

    /// The same element XML with every inherited namespace declared on its
    /// root, so that it parses standalone.
    ///
    /// [`xml`](Self::xml) borrows the span exactly as it appears in the
    /// processed owner XML, where a prefix the shape uses may be bound by an
    /// ancestor outside the span. This accessor re-declares on the fragment's
    /// own root element every binding in scope at the span that the root does
    /// not already bind, so each qualified name resolves to the namespace it
    /// resolves to in the owner. The fragment is returned borrowed when
    /// nothing has to be added, which is the case whenever the root already
    /// carries every in-scope binding.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Invalid`] when the recorded span is no longer inside
    /// its processed owner XML, and
    /// [`Error::MarkupCompatibility`](crate::Error::MarkupCompatibility) when
    /// the owner XML is not well formed before the span, the span does not
    /// begin with an element, a bound of the shared fragment helper is
    /// reached, or the re-declared copy cannot be reserved.
    pub fn self_contained_xml(self) -> Result<Cow<'a, [u8]>> {
        let range = self.record.span.range(self.xml.len())?;
        let length = range.len();
        Ok(litchi_ooxml_common::mce::self_contained_fragment(
            self.xml,
            range.start,
            length,
            &shape_mce_limits(self.limits),
        )?)
    }

    /// Decoded producer name from `cNvPr`.
    #[must_use]
    pub fn name(self) -> Option<&'a str> {
        self.record.name?.get(self.strings)
    }

    /// Numeric non-visual ID used by low-level animation references.
    ///
    /// Semantic selectors should normally use [`name`](Self::name) or a
    /// checked pre-order position instead.
    #[must_use]
    pub fn id(self) -> Option<u32> {
        self.record.id
    }

    /// Complete local transform bounds when explicitly present.
    #[must_use]
    pub fn bounds(self) -> Option<Bounds> {
        self.record.bounds
    }

    /// Placeholder metadata when declared by the shape.
    #[must_use]
    pub fn placeholder(self) -> Option<Placeholder<'a>> {
        let value = self.record.placeholder?;
        Some(Placeholder {
            kind: value.kind.and_then(|span| span.get(self.strings)),
            index: value.index,
            type_extension: value.type_extension,
        })
    }

    /// Decoded `DrawingML` text captured in full-owner namespace context.
    ///
    /// `None` distinguishes shapes with no `DrawingML` text from an explicitly
    /// empty text body.
    #[must_use]
    pub fn text(self) -> Option<&'a str> {
        self.record.text?.get(self.strings)
    }

    /// Qualified source element name, useful for an unknown extension shape.
    #[must_use]
    pub fn source_name(self) -> Option<&'a str> {
        self.record.source_name?.get(self.strings)
    }

    /// Span of the start tag of the shape's `cNvPr`, when it has one.
    #[inline]
    #[must_use]
    pub(crate) const fn non_visual_span(self) -> Option<Span> {
        self.record.non_visual
    }

    /// The non-visual drawing properties the shape's `cNvPr` declares:
    /// alternative text, title and the hidden flag.
    ///
    /// They are decoded from the `cNvPr` start tag on each call; a shape
    /// without `cNvPr` has none.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Invalid`] when the recorded span is no longer inside
    /// its owner XML or `hidden` is not an XML Schema boolean, or for a value that is not UTF-8 or holds a malformed
    /// character reference.
    pub fn non_visual(self) -> Result<NonVisual<'a>> {
        let Some(span) = self.record.non_visual else {
            return Ok(NonVisual::default());
        };
        let range = span.range(self.xml.len())?;
        let tag = self
            .xml
            .get(range)
            .ok_or_else(|| Error::Invalid("shape cNvPr span became invalid".into()))?;
        read_non_visual(tag)
    }

    /// Alternative text (`cNvPr/@descr`); see [`Self::non_visual`].
    ///
    /// # Errors
    ///
    /// The errors of [`Self::non_visual`].
    pub fn description(self) -> Result<Option<Cow<'a, str>>> {
        Ok(self.non_visual()?.description)
    }

    /// Title (`cNvPr/@title`); see [`Self::non_visual`].
    ///
    /// # Errors
    ///
    /// The errors of [`Self::non_visual`].
    pub fn title(self) -> Result<Option<Cow<'a, str>>> {
        Ok(self.non_visual()?.title)
    }

    /// The hidden flag (`cNvPr/@hidden`) as written; see
    /// [`Self::non_visual`].
    ///
    /// # Errors
    ///
    /// The errors of [`Self::non_visual`].
    pub fn hidden(self) -> Result<Option<bool>> {
        Ok(self.non_visual()?.hidden)
    }

    /// The relationship that names the shape's payload part: an embedded or
    /// linked picture's image, a chart frame's chart, a diagram frame's data
    /// model, an OLE frame's object or a content part. `None` for shapes
    /// whose content is inline, and when the payload element carries no
    /// relationship. Competing relationships return a typed raw error.
    pub(crate) fn payload(self) -> std::result::Result<Option<Payload<'a>>, super::raw::Error> {
        let Some(payload) = self.record.payload else {
            return Ok(None);
        };
        if let Some(error) = payload.error {
            return Err(error);
        }
        let range = payload
            .value
            .range(self.xml.len())
            .map_err(|_| super::raw::Error::InvalidAttribute)?;
        Ok(Some(Payload {
            raw: self
                .xml
                .get(range)
                .ok_or(super::raw::Error::InvalidAttribute)?,
            span: payload.value,
            attribute: payload.attribute,
        }))
    }

    /// Span of the shape's own transform element (`p:spPr/a:xfrm`,
    /// `p:grpSpPr/a:xfrm` or a graphic frame's `p:xfrm`), when it has one.
    ///
    /// [`Self::bounds`] reads its offset and extent; this span also covers
    /// rotation, flips and a group's child coordinate space.
    #[inline]
    #[must_use]
    pub(crate) const fn transform_span(self) -> Option<Span> {
        self.record.transform
    }

    /// The shape's own transform, decoded from its recorded owner span.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Invalid`] when the recorded span is no longer inside
    /// its owner XML, and [`Error::Drawing`] when the transform does not
    /// follow `a:CT_Transform2D`.
    pub fn transform(self) -> Result<Option<Transform>> {
        if self.record.transform_invalid {
            return Err(Error::Invalid(
                "shape transform is ambiguous or has invalid namespaces".into(),
            ));
        }
        let Some(span) = self.record.transform else {
            return Ok(None);
        };
        let range = span.range(self.xml.len())?;
        let xml = self
            .xml
            .get(range)
            .ok_or_else(|| Error::Invalid("shape transform span became invalid".into()))?;
        let transform_limits = transform_limits_for(self.limits);
        Ok(Some(litchi_drawingml::transform::codec::read(
            xml,
            &transform_limits,
        )?))
    }

    /// Span of the shape's direct `p:txBody` element in processed owner XML.
    ///
    /// This is retained for the low-level [`super::raw`] view. Ordinary shape
    /// callers should use [`Self::text_body`] so duplicate bodies and the
    /// inherited namespace scope are checked before decoding.
    #[inline]
    #[must_use]
    pub(crate) const fn text_body_span(self) -> Option<Span> {
        self.record.text_body
    }

    /// The shape's local `DrawingML` text body.
    ///
    /// The body is decoded lazily in the namespace scope of its position in
    /// the processed owner. Placeholder, master, list-style and theme
    /// inheritance is intentionally outside this local view.
    ///
    /// # Errors
    ///
    /// Returns a content-free [`Error::Invalid`] when the shape declares
    /// competing direct bodies or the retained span is no longer valid, and
    /// [`Error::Drawing`] when the shared strict text-body reader rejects a
    /// value or resource bound.
    pub fn text_body(self) -> Result<Option<Body>> {
        if self.record.text_body_invalid {
            return Err(Error::Invalid(
                "shape contains duplicate direct text bodies".into(),
            ));
        }
        let Some(span) = self.record.text_body else {
            return Ok(None);
        };
        let range = span.range(self.xml.len())?;
        let scope = self.namespace_scope(range.start)?;
        let xml = self
            .xml
            .get(range)
            .ok_or_else(|| Error::Invalid("shape text body span became invalid".into()))?;
        let limits = text_body_limits_for(self.limits, self.xml.len());
        Ok(Some(text_reader::read(xml, scope, limits)?))
    }

    /// The paragraphs and text runs of [`Self::text_body`] (record 0348);
    /// `None` without a body.
    ///
    /// # Errors
    ///
    /// As [`Self::text_body`].
    pub fn text_index(self) -> Result<Option<BodyTextIndex>> {
        self.text_body()?
            .map(|body| crate::table::CellTextIndex::from_body(&body).map(BodyTextIndex))
            .transpose()
    }

    /// [`Self::text_index`] of a body [`Self::text_body`] has read (record
    /// 0348): from one pass over the body when it can, without decoding it
    /// again, so values the reader would refuse are not checked here.
    ///
    /// # Errors
    ///
    /// As [`Self::text_body`] (for competing bodies, and for a body the pass
    /// leaves to the reader).
    pub fn read_text_index(self) -> Result<Option<BodyTextIndex>> {
        if self.record.text_body_invalid {
            return Err(Error::Invalid(
                "shape contains duplicate direct text bodies".into(),
            ));
        }
        let Some(span) = self.record.text_body else {
            return Ok(None);
        };
        let range = span.range(self.xml.len())?;
        let scope = self.namespace_scope(range.start)?;
        let xml = self
            .xml
            .get(range)
            .ok_or_else(|| Error::Invalid("shape text body span became invalid".into()))?;
        if let Some(index) = crate::table::shape::scan_read_body_text_index(xml, scope)? {
            return Ok(Some(BodyTextIndex(index)));
        }
        self.text_index()
    }

    /// Resolve the namespaces inherited by a direct child body.
    fn namespace_scope(self, offset: usize) -> Result<NamespaceResolver> {
        if self.record.inherits_root_namespaces {
            let shape = self.record.span.range(self.xml.len())?;
            let mut resolver = NamespaceResolver::default();
            for tag in [self.xml, self.xml.get(shape.start..).unwrap_or_default()] {
                if let Some(element) = first_start_tag(tag)? {
                    resolver.push(&element).map_err(|_| {
                        Error::Invalid("shape text body namespace scope is invalid".into())
                    })?;
                }
            }
            return Ok(resolver);
        }
        let scope = InScopeNamespaces::at_offset(self.xml, offset, &shape_mce_limits(self.limits))
            .map_err(|_| Error::Invalid("shape text body namespace scope is invalid".into()))?;
        let mut resolver = NamespaceResolver::default();
        for (prefix, namespace) in scope.iter() {
            let declaration = if prefix.is_empty() {
                PrefixDeclaration::Default
            } else {
                PrefixDeclaration::Named(prefix)
            };
            resolver
                .add(declaration, Namespace(namespace))
                .map_err(|_| Error::Invalid("shape text body namespace scope is invalid".into()))?;
        }
        Ok(resolver)
    }
}

fn shape_mce_limits(limits: super::reader::Limits) -> litchi_ooxml_common::mce::Limits {
    litchi_ooxml_common::mce::Limits {
        max_input_bytes: limits.input_bytes(),
        max_output_bytes: limits.output_bytes(),
        max_depth: limits.depth(),
        ..litchi_ooxml_common::mce::Limits::default()
    }
}

fn transform_limits_for(limits: super::reader::Limits) -> litchi_drawingml::transform::Limits {
    litchi_drawingml::transform::Limits::new(
        limits.input_bytes().clamp(
            1,
            litchi_drawingml::transform::Limits::DEFAULT.max_xml_bytes,
        ),
        limits
            .nodes()
            .clamp(1, litchi_drawingml::transform::Limits::DEFAULT.max_nodes),
    )
}

fn text_body_limits_for(limits: super::reader::Limits, owner_bytes: usize) -> text_reader::Limits {
    text_reader::Limits::new(
        limits.input_bytes().min(owner_bytes).max(1),
        limits.depth(),
        limits.nodes(),
        limits.retained_text_bytes(),
    )
    .unwrap_or(text_reader::Limits::DEFAULT)
}

/// The first start tag (or empty element) of `xml`.
fn first_start_tag(xml: &[u8]) -> Result<Option<BytesStart<'_>>> {
    let mut reader = quick_xml::Reader::from_reader(xml);
    loop {
        match reader
            .read_event()
            .map_err(|_| Error::Invalid("malformed shape attribute XML".into()))?
        {
            Event::Start(element) | Event::Empty(element) => return Ok(Some(element)),
            Event::Eof => return Ok(None),
            _ => {},
        }
    }
}

/// Decode `descr`, `title` and `hidden` from one `cNvPr` start tag.
fn read_non_visual(tag: &[u8]) -> Result<NonVisual<'_>> {
    let element = first_start_tag(tag)?
        .ok_or_else(|| Error::Invalid("shape cNvPr span holds no start tag".into()))?;
    let mut values = NonVisual::default();
    for attribute in element.checked_attributes() {
        let attribute =
            attribute.map_err(|_| Error::Invalid("malformed shape attribute XML".into()))?;
        if attribute.key.prefix().is_some() {
            continue;
        }
        match attribute.key.local_name().as_ref() {
            b"descr" => {
                values.description = Some(decode_attribute_value(raw_in(tag, &attribute)?)?);
            },
            b"title" => values.title = Some(decode_attribute_value(raw_in(tag, &attribute)?)?),
            b"hidden" => {
                let value = decode_attribute_value(raw_in(tag, &attribute)?)?;
                let hidden = litchi_drawingml::text::parse_bool(&value)
                    .map_err(|_| Error::Invalid("invalid shape hidden flag".into()))?;
                values.hidden = Some(hidden);
            },
            _ => {},
        }
    }
    Ok(values)
}

/// The bytes of an attribute's value as a slice of `tag`, which the
/// attribute was read from; quick-xml lends values from its input.
fn raw_in<'a>(tag: &'a [u8], attribute: &Attribute<'_>) -> Result<&'a [u8]> {
    let value: &[u8] = attribute.value.as_ref();
    let start = (value.as_ptr() as usize).wrapping_sub(tag.as_ptr() as usize);
    start
        .checked_add(value.len())
        .and_then(|end| tag.get(start..end))
        .filter(|slice| slice.as_ptr() == value.as_ptr())
        .ok_or_else(|| Error::Invalid("shape attribute value is not part of its tag".into()))
}

/// Decode an attribute value as XML 1.0 attribute-value normalization does:
/// UTF-8, character and predefined entity references, and whitespace.
fn decode_attribute_value(raw: &[u8]) -> Result<Cow<'_, str>> {
    Attribute {
        key: QName(b"value"),
        value: Cow::Borrowed(raw),
    }
    .normalized_value(XmlVersion::Explicit1_0)
    .map_err(|_| Error::Invalid("malformed shape attribute XML".into()))
}

macro_rules! leaf {
    ($(#[$meta:meta])* $name:ident) => {
        $(#[$meta])*
        #[derive(Debug, Clone, Copy)]
        pub struct $name<'a>(pub(super) Common<'a>);

        impl<'a> $name<'a> {
            /// Borrow the properties common to all shape kinds.
            #[inline]
            pub const fn common(self) -> Common<'a> {
                self.0
            }
        }
    };
}

leaf!(/// An ordinary `PresentationML` auto shape (`p:sp`).
    Auto);
leaf!(/// A picture (`p:pic`).
    Picture);
leaf!(/// A graphic frame whose payload is a `DrawingML` table.
    Table);
leaf!(/// A graphic frame whose payload is a `DrawingML` chart.
    Chart);
leaf!(/// A graphic frame whose payload is a `DrawingML` diagram.
    Diagram);
leaf!(/// A graphic frame whose payload is an embedded or linked OLE object.
    Ole);
leaf!(/// A graphic frame with an unclassified or extension payload.
    Frame);
leaf!(/// A connector (`p:cxnSp`).
    Connector);
leaf!(/// A Microsoft content-part shape.
    Content);
leaf!(/// A retained shape-like extension not understood by this release.
    Unknown);

/// A grouped shape with direct hierarchical child access.
#[derive(Debug, Clone, Copy)]
pub struct Group<'a>(pub(super) Common<'a>);

impl<'a> Group<'a> {
    /// Borrow the properties common to all shape kinds.
    #[inline]
    #[must_use]
    pub const fn common(self) -> Common<'a> {
        self.0
    }

    /// Iterate direct children in source order.
    #[must_use]
    pub fn shapes(self) -> Shapes<'a> {
        let first = self.0.index.saturating_add(1);
        let end = self.0.record.subtree_end as usize;
        Shapes {
            xml: self.0.xml,
            records: self.0.records,
            strings: self.0.strings,
            limits: self.0.limits,
            cursor: first,
            end,
            parent: Some(self.0.index as u32),
            preorder: false,
        }
    }
}

/// One semantic `PresentationML` shape.
///
/// The enum carries a typed view rather than requiring users to compare a
/// separate numeric or discriminator value.
#[derive(Debug, Clone, Copy)]
#[non_exhaustive]
pub enum Shape<'a> {
    Auto(Auto<'a>),
    Picture(Picture<'a>),
    Table(Table<'a>),
    Chart(Chart<'a>),
    Diagram(Diagram<'a>),
    Ole(Ole<'a>),
    Frame(Frame<'a>),
    Group(Group<'a>),
    Connector(Connector<'a>),
    Content(Content<'a>),
    Unknown(Unknown<'a>),
}

impl<'a> Shape<'a> {
    pub(super) const fn from_common(common: Common<'a>) -> Self {
        match common.record.kind {
            Kind::Auto => Self::Auto(Auto(common)),
            Kind::Picture => Self::Picture(Picture(common)),
            Kind::Table => Self::Table(Table(common)),
            Kind::Chart => Self::Chart(Chart(common)),
            Kind::Diagram => Self::Diagram(Diagram(common)),
            Kind::Ole => Self::Ole(Ole(common)),
            Kind::Frame => Self::Frame(Frame(common)),
            Kind::Group => Self::Group(Group(common)),
            Kind::Connector => Self::Connector(Connector(common)),
            Kind::Content => Self::Content(Content(common)),
            Kind::Unknown => Self::Unknown(Unknown(common)),
        }
    }

    /// Borrow the properties common to every variant.
    #[must_use]
    pub const fn common(self) -> Common<'a> {
        match self {
            Self::Auto(value) => value.0,
            Self::Picture(value) => value.0,
            Self::Table(value) => value.0,
            Self::Chart(value) => value.0,
            Self::Diagram(value) => value.0,
            Self::Ole(value) => value.0,
            Self::Frame(value) => value.0,
            Self::Group(value) => value.0,
            Self::Connector(value) => value.0,
            Self::Content(value) => value.0,
            Self::Unknown(value) => value.0,
        }
    }

    /// # Errors
    ///
    /// Returns an error if the operation fails.
    #[inline]
    pub fn span(self) -> Result<Span> {
        self.common().span()
    }

    /// Exact borrowed element XML without a per-shape copy.
    ///
    /// The returned bytes are an inner fragment of
    /// [`Scene::xml`](super::Scene::xml) whose namespace declarations are the
    /// ones in scope at its position in that processed owner XML, not
    /// necessarily repeated on the fragment's own root element. A caller that
    /// parses the fragment standalone should use
    /// [`self_contained_xml`](Self::self_contained_xml) instead. See
    /// [`Common::xml`].
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    #[inline]
    pub fn xml(self) -> Result<&'a [u8]> {
        self.common().xml()
    }

    /// The same element XML with every inherited namespace declared on its
    /// root, so that it parses standalone.
    ///
    /// See [`Common::self_contained_xml`].
    ///
    /// # Errors
    ///
    /// Returns [`Error::Invalid`] when the recorded span is no longer inside
    /// its processed owner XML, and
    /// [`Error::MarkupCompatibility`](crate::Error::MarkupCompatibility) when
    /// the owner XML is not well formed before the span, the span does not
    /// begin with an element, a bound of the shared fragment helper is
    /// reached, or the re-declared copy cannot be reserved.
    #[inline]
    pub fn self_contained_xml(self) -> Result<Cow<'a, [u8]>> {
        self.common().self_contained_xml()
    }

    #[inline]
    #[must_use]
    pub fn name(self) -> Option<&'a str> {
        self.common().name()
    }

    #[inline]
    #[must_use]
    pub fn id(self) -> Option<u32> {
        self.common().id()
    }

    #[inline]
    #[must_use]
    pub fn bounds(self) -> Option<Bounds> {
        self.common().bounds()
    }

    #[inline]
    #[must_use]
    pub fn placeholder(self) -> Option<Placeholder<'a>> {
        self.common().placeholder()
    }

    #[inline]
    #[must_use]
    pub fn text(self) -> Option<&'a str> {
        self.common().text()
    }

    /// See [`Common::non_visual`].
    ///
    /// # Errors
    ///
    /// The errors of [`Common::non_visual`].
    #[inline]
    pub fn non_visual(self) -> Result<NonVisual<'a>> {
        self.common().non_visual()
    }

    /// See [`Common::description`].
    ///
    /// # Errors
    ///
    /// The errors of [`Common::non_visual`].
    #[inline]
    pub fn description(self) -> Result<Option<Cow<'a, str>>> {
        self.common().description()
    }

    /// See [`Common::title`].
    ///
    /// # Errors
    ///
    /// The errors of [`Common::non_visual`].
    #[inline]
    pub fn title(self) -> Result<Option<Cow<'a, str>>> {
        self.common().title()
    }

    /// See [`Common::hidden`].
    ///
    /// # Errors
    ///
    /// The errors of [`Common::non_visual`].
    #[inline]
    pub fn hidden(self) -> Result<Option<bool>> {
        self.common().hidden()
    }

    /// See [`Common::transform`].
    ///
    /// # Errors
    ///
    /// The errors of [`Common::transform`].
    #[inline]
    pub fn transform(self) -> Result<Option<Transform>> {
        self.common().transform()
    }

    /// The shape's local `DrawingML` text body.
    ///
    /// # Errors
    ///
    /// The errors of [`Common::text_body`].
    #[inline]
    pub fn text_body(self) -> Result<Option<Body>> {
        self.common().text_body()
    }
}

/// A borrowed shape iterator.
#[derive(Clone)]
pub struct Shapes<'a> {
    pub(super) xml: &'a [u8],
    pub(super) records: &'a [Record],
    pub(super) strings: &'a str,
    pub(super) limits: super::reader::Limits,
    pub(super) cursor: usize,
    pub(super) end: usize,
    pub(super) parent: Option<u32>,
    pub(super) preorder: bool,
}

impl std::fmt::Debug for Shapes<'_> {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("Shapes")
            .field("cursor", &self.cursor)
            .field("end", &self.end)
            .finish_non_exhaustive()
    }
}

impl<'a> Iterator for Shapes<'a> {
    type Item = Shape<'a>;

    fn next(&mut self) -> Option<Self::Item> {
        while self.cursor < self.end {
            let index = self.cursor;
            let record = self.records.get(index)?;
            self.cursor = if self.preorder {
                index.checked_add(1)?
            } else {
                (record.subtree_end as usize).max(index.checked_add(1)?)
            };
            if self.preorder || record.parent == self.parent {
                let common = Common {
                    xml: self.xml,
                    records: self.records,
                    strings: self.strings,
                    limits: self.limits,
                    record,
                    index,
                };
                return Some(Shape::from_common(common));
            }
        }
        None
    }

    fn size_hint(&self) -> (usize, Option<usize>) {
        let remaining = self.end.saturating_sub(self.cursor);
        if self.preorder {
            (remaining, Some(remaining))
        } else {
            (0, Some(remaining))
        }
    }
}

impl std::iter::FusedIterator for Shapes<'_> {}

impl std::fmt::Debug for NonVisual<'_> {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("NonVisual")
            .field("has_description", &self.description.is_some())
            .field("has_title", &self.title.is_some())
            .field("hidden", &self.hidden)
            .finish()
    }
}
impl std::fmt::Debug for Payload<'_> {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("Payload")
            .field("attribute", &self.attribute)
            .finish_non_exhaustive()
    }
}

impl std::fmt::Debug for Common<'_> {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("Common")
            .field("position", &self.index)
            .field("kind", &self.record.kind)
            .finish_non_exhaustive()
    }
}

/// The paragraphs and text runs of a shape's text body (record 0348); see
/// [`Common::text_index`].
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct BodyTextIndex(crate::table::CellTextIndex);

impl BodyTextIndex {
    /// The number of paragraphs.
    #[must_use]
    pub fn paragraph_count(&self) -> usize {
        self.0.paragraph_count(0, 0).unwrap_or(0)
    }

    /// The position among its paragraph's run children (`a:r`, `a:br`,
    /// `a:fld`) of the `ordinal`-th text run (zero-based) of paragraph
    /// `paragraph`; `None` when there is no such run.
    #[must_use]
    pub fn text_run(&self, paragraph: usize, ordinal: usize) -> Option<usize> {
        self.0.text_run(0, 0, paragraph, ordinal)
    }
}

#[cfg(test)]
mod tests {
    use super::{text_body_limits_for, transform_limits_for};
    use crate::shape::reader;

    #[test]
    fn subordinate_readers_use_the_scene_input_budget() {
        let limits = reader::Limits::new(173, 4096, 64, 128, 1024, 4096)
            .expect("finite asymmetric scene limits");
        assert_eq!(transform_limits_for(limits).max_xml_bytes, 173);
        assert_eq!(text_body_limits_for(limits, 2048).input_bytes(), 173);

        let reverse = reader::Limits::new(4096, 173, 64, 128, 1024, 4096)
            .expect("finite reverse asymmetric scene limits");
        assert_eq!(
            transform_limits_for(reverse).max_xml_bytes,
            4096.min(litchi_drawingml::transform::Limits::DEFAULT.max_xml_bytes)
        );
        assert_eq!(text_body_limits_for(reverse, 2048).input_bytes(), 2048);
    }
}
