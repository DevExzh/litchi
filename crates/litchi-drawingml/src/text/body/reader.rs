//! Bounded, namespace-aware reader for `DrawingML` text bodies.
//!
//! [`Builder`] is the one implementation of text-body reading in Litchi. It
//! consumes the events of one `a:CT_TextBody` element (`p:txBody`,
//! `xdr:txBody`, a table cell's `a:txBody`, ...) and builds a [`Body`]:
//! body properties, paragraphs with their local paragraph properties, and
//! runs, line breaks and fields with their local character properties. A host
//! that streams a larger part forwards the element's descendants to a
//! builder; [`read`] drives one over a single element span.
//!
//! Property values are *local*: an attribute that is absent stays `None`.
//! Markup the model does not cover (list styles, effects, outlines,
//! extension lists, foreign elements) is skipped without being interpreted.
//! An `a:tab` of an `a:tabLst` is a tab stop
//! ([`ParagraphProperties::tab_stops`](super::ParagraphProperties::tab_stops)),
//! never text; an `a:tab` or `a:br` inside a run
//! is a tab or line-break character of that run, as Litchi's shape text
//! extractors read it.
//!
//! ```
//! use litchi_drawingml::text::body::reader;
//! use litchi_drawingml::text::{Alignment, TextSize};
//! # fn main() -> Result<(), Box<dyn std::error::Error>> {
//!
//! let xml = br#"<p:txBody xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
//!     xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
//!     <a:bodyPr/>
//!     <a:p><a:pPr algn="ctr" lvl="1"/><a:r><a:rPr sz="2400" b="1"/><a:t>Hello</a:t></a:r></a:p>
//! </p:txBody>"#;
//! let body = reader::read(xml, quick_xml::name::NamespaceResolver::default(), reader::Limits::DEFAULT)?;
//! let paragraph = &body.paragraphs[0];
//! assert_eq!(paragraph.properties.alignment, Some(Alignment::Center));
//! assert_eq!(paragraph.properties.level, Some(1));
//! assert_eq!(paragraph.runs[0].text, "Hello");
//! assert_eq!(paragraph.runs[0].properties.size, TextSize::new(2400).ok());
//! assert_eq!(paragraph.runs[0].properties.bold, Some(true));
//! # Ok(())
//! # }
//! ```

use std::borrow::Cow;
use std::ops::Range;
use std::str::FromStr;

use litchi_core::xml::ReaderOrigin;
use litchi_ooxml_common::XmlError;

use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use litchi_ooxml_common::xml::{
    DRAWINGML_NAMESPACE, STRICT_DRAWINGML_NAMESPACE, decode_xml_reference,
    unqualified_attribute_value, xsd_token_atom,
};
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::attributes::Attribute;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

use super::{
    Body, Bullet, BulletColor, BulletFont, BulletSize, Field, Font, Hyperlink, Paragraph,
    Properties, Run, RunKind, RunProperties, Spacing, TabStop, TextFill,
};
use crate::color;
use crate::text::{
    Alignment, AutoNumber, Autofit, Caps, Coordinate32, FontAlignment, Strike, TabAlignment,
    TextPoint, TextSize, Underline, parse_bool,
};
use crate::{Error, Result};

/// Finite resources one text body may use.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Limits {
    input_bytes: usize,
    depth: usize,
    nodes: usize,
    text_bytes: usize,
}

impl Limits {
    /// Conservative defaults, matching the bounds of a `PresentationML` shape
    /// owner.
    pub const DEFAULT: Self = Self {
        input_bytes: 64 * 1024 * 1024,
        depth: 256,
        nodes: 1_000_000,
        text_bytes: 16 * 1024 * 1024,
    };

    /// Construct a finite, nonzero limit set.
    ///
    /// `input_bytes` bounds the element span [`read`] accepts, `depth`
    /// the nesting below the text-body element, `nodes` its descendant
    /// elements and `text_bytes` the run and field text the body keeps.
    #[must_use]
    pub const fn new(
        input_bytes: usize,
        depth: usize,
        nodes: usize,
        text_bytes: usize,
    ) -> Option<Self> {
        if input_bytes == 0 || depth == 0 || nodes == 0 || text_bytes == 0 {
            None
        } else {
            Some(Self {
                input_bytes,
                depth,
                nodes,
                text_bytes,
            })
        }
    }

    /// Largest element span [`read`] accepts.
    #[inline]
    #[must_use]
    pub const fn input_bytes(self) -> usize {
        self.input_bytes
    }

    /// Deepest nesting below the text-body element.
    #[inline]
    #[must_use]
    pub const fn depth(self) -> usize {
        self.depth
    }

    /// Most descendant elements.
    #[inline]
    #[must_use]
    pub const fn nodes(self) -> usize {
        self.nodes
    }

    /// Most run and field text, in bytes.
    #[inline]
    #[must_use]
    pub const fn text_bytes(self) -> usize {
        self.text_bytes
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// Read one text-body element span under an explicit namespace scope and
/// resource policy.
///
/// `xml` is the complete element, for example `<p:txBody>...</p:txBody>`;
/// its own name is not interpreted. Markup-compatibility processing is the
/// caller's: an `mc:` element left in the body is refused.
///
/// # Errors
///
/// Returns [`Error::Xml`] or [`Error::Common`] for malformed XML,
/// [`Error::Invalid`] for an invalid modeled value or a fragment that is not
/// one element, and [`Error::Limit`] when a bound is exceeded.
/// Read one text-body element span.
///
/// `scope` holds the namespace bindings the element inherits from the
/// document it was cut from (bindings the element declares itself need not be
/// in it); a host builds it by pushing the ancestors' start tags or by adding
/// the declarations in scope. Byte offsets the reader keeps are relative to
/// `xml`.
///
/// # Errors
///
/// The errors of [`read`], under `limits`.
pub fn read(xml: &[u8], scope: NamespaceResolver, limits: Limits) -> Result<Body> {
    if xml.len() > limits.input_bytes {
        return Err(Error::limit(
            crate::LimitKind::InputBytes,
            xml.len(),
            limits.input_bytes,
        ));
    }
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    *reader.resolver_mut() = scope;
    let origin = ReaderOrigin::of(xml);
    let mut builder = Builder::new(xml, limits)?;
    let mut depth = 0usize;
    let mut root_closed = false;
    loop {
        let start = position(&reader, origin)?;
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|_| Error::Xml("malformed text body XML".into()))?;
        let end = position(&reader, origin)?;
        let (namespace, event) = reader.resolver().resolve_event(event);
        match event {
            Event::Start(element) => {
                if root_closed {
                    return Err(invalid("text body XML contains more than one element"));
                }
                if depth > 0 {
                    builder.start(
                        &element,
                        &namespace,
                        reader.resolver(),
                        decoder,
                        start..end,
                        false,
                    )?;
                }
                depth = depth.saturating_add(1);
            },
            Event::Empty(element) => {
                if root_closed {
                    return Err(invalid("text body XML contains more than one element"));
                }
                if depth > 0 {
                    builder.start(
                        &element,
                        &namespace,
                        reader.resolver(),
                        decoder,
                        start..end,
                        true,
                    )?;
                } else {
                    root_closed = true;
                }
            },
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("text body XML closes an element it did not open"))?;
                if depth > 0 {
                    builder.end(end)?;
                } else {
                    root_closed = true;
                }
            },
            Event::Text(text) if depth > 0 => {
                let encoded = text
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|_error| Error::Xml("malformed text body XML".into()))?;
                builder.text_xml(&encoded)?;
            },
            Event::CData(text) if depth > 0 => {
                let decoded = text
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|_error| Error::Xml("malformed text body XML".into()))?;
                builder.text(&decoded)?;
            },
            Event::GeneralRef(reference) if depth > 0 => {
                builder.text(
                    &decode_xml_reference(&reference).map_err(|error| redact(error.into()))?,
                )?;
            },
            Event::Text(text)
                if text
                    .as_ref()
                    .iter()
                    .any(|byte| !matches!(byte, b' ' | b'\t' | b'\r' | b'\n')) =>
            {
                return Err(invalid("text body contains text outside its root"));
            },
            Event::CData(_) | Event::GeneralRef(_) => {
                return Err(invalid("text body contains data outside its root"));
            },
            Event::DocType(_) => {
                return Err(invalid("DOCTYPE is forbidden in a DrawingML text body"));
            },
            Event::Eof => break,
            // Character data outside the element, comments, processing
            // instructions and declarations carry no text-body content.
            _ => {},
        }
    }
    if !root_closed || depth != 0 {
        return Err(invalid("text body XML is not one complete element"));
    }
    builder.finish()
}

/// Apply the attributes of one `a:bodyPr` element to `properties`.
///
/// Each attribute present replaces the current value; absent ones keep it.
/// This is the body-property reading every text-body host shares.
///
/// # Errors
///
/// Returns [`Error::Common`] for malformed attribute syntax and
/// [`Error::Invalid`] for a value outside its schema domain.
pub(super) fn apply_body_properties(
    properties: &mut Properties,
    element: &BytesStart<'_>,
    decoder: Decoder,
) -> Result<()> {
    if let Some(value) = unqualified_attribute_value(element, b"lIns", decoder)? {
        properties.insets.left = parse_value(&value, "left text inset")?;
    }
    if let Some(value) = unqualified_attribute_value(element, b"tIns", decoder)? {
        properties.insets.top = parse_value(&value, "top text inset")?;
    }
    if let Some(value) = unqualified_attribute_value(element, b"rIns", decoder)? {
        properties.insets.right = parse_value(&value, "right text inset")?;
    }
    if let Some(value) = unqualified_attribute_value(element, b"bIns", decoder)? {
        properties.insets.bottom = parse_value(&value, "bottom text inset")?;
    }
    if let Some(value) = unqualified_attribute_value(element, b"anchor", decoder)? {
        properties.vertical_anchor = parse_value(&value, "text anchor")?;
    }
    if let Some(value) = unqualified_attribute_value(element, b"anchorCtr", decoder)? {
        properties.anchor_center = parse_dml_bool(&value, "anchorCtr")?;
    }
    if let Some(value) = unqualified_attribute_value(element, b"vert", decoder)? {
        properties.direction = parse_value(&value, "text direction")?;
    }
    if let Some(value) = unqualified_attribute_value(element, b"wrap", decoder)? {
        properties.wrap = parse_value(&value, "text wrap")?;
    }
    if let Some(value) = unqualified_attribute_value(element, b"numCol", decoder)? {
        properties.column_count = parse_value(&value, "text column count")?;
    }
    if let Some(value) = unqualified_attribute_value(element, b"spcFirstLastPara", decoder)? {
        properties.space_first_last_paragraph = parse_dml_bool(&value, "spcFirstLastPara")?;
    }
    Ok(())
}

/// The autofit an `a:bodyPr` child selects, if it is one of the three
/// autofit elements.
#[must_use]
pub fn autofit(local_name: &[u8]) -> Option<Autofit> {
    match local_name {
        b"noAutofit" => Some(Autofit::None),
        b"spAutoFit" => Some(Autofit::Shape),
        b"normAutofit" => Some(Autofit::Normal),
        _ => None,
    }
}

/// Which character properties an open `a:rPr`-like element fills.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum PropertiesTarget {
    /// The open run's, line break's or field's `a:rPr`.
    Run,
    /// The open paragraph's `a:pPr/a:defRPr`.
    ParagraphDefault,
    /// The open paragraph's `a:endParaRPr`.
    ParagraphEnd,
}

/// Where a captured color goes.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum ColorTarget {
    /// The pending character properties' solid fill.
    Fill,
    /// The pending character properties' highlight.
    Highlight,
    /// The open paragraph's bullet color.
    Bullet,
}

/// One paragraph spacing slot.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum SpacingSlot {
    Line,
    Before,
    After,
}

/// The modeled element a builder frame stands for.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Frame {
    /// `a:bodyPr`: its autofit child is read.
    BodyProperties,
    /// `a:p`.
    Paragraph,
    /// `a:pPr` of the open paragraph.
    ParagraphProperties,
    /// `a:lnSpc`, `a:spcBef` or `a:spcAft`.
    Spacing(SpacingSlot),
    /// `a:buBlip`: its `a:blip` names the picture.
    BulletPicture,
    /// `a:tabLst`.
    TabList,
    /// `a:r`, `a:br` or `a:fld`: the open run.
    Run,
    /// Character properties for one target.
    RunProperties(PropertiesTarget),
    /// `a:solidFill`, `a:highlight` or `a:buClr`: its first color choice is
    /// captured.
    ColorContainer(ColorTarget),
    /// A color choice whose bytes start at `start`.
    Color { target: ColorTarget, start: usize },
    /// `a:t` of the open run.
    Text,
    /// A modeled element whose children are not read.
    Leaf,
}

/// An event-driven text-body reader; see the [module documentation](self).
///
/// A host creates one when a text-body element starts and forwards every
/// event of that element's descendants, in document order and with byte
/// positions in one `document`: [`start`](Self::start) for each start tag or
/// empty element, [`end`](Self::end) for each end tag, and
/// [`text`](Self::text) for decoded character data. The text-body element's
/// own start and end tags are not forwarded. [`finish`](Self::finish) then
/// yields the body.
pub struct Builder<'x> {
    document: &'x [u8],
    limits: Limits,
    body: Body,
    paragraph: Option<Paragraph>,
    run: Option<Run>,
    /// Character properties being read for the innermost
    /// [`Frame::RunProperties`].
    pending: Option<RunProperties>,
    frames: Vec<Frame>,
    /// Open elements of a subtree being skipped; zero when none is.
    skipped: usize,
    nodes: usize,
    text_bytes: usize,
    seen: Vec<u64>,
}

impl<'x> Builder<'x> {
    /// Start reading the descendants of one text-body element of
    /// `document`.
    /// # Errors
    /// Refuses input larger than the configured physical byte bound.
    pub fn new(document: &'x [u8], limits: Limits) -> Result<Self> {
        if document.len() > limits.input_bytes {
            return Err(Error::limit(
                crate::LimitKind::InputBytes,
                document.len(),
                limits.input_bytes,
            ));
        }
        Ok(Self {
            document,
            limits,
            body: Body::default(),
            paragraph: None,
            run: None,
            pending: None,
            frames: Vec::new(),
            skipped: 0,
            nodes: 0,
            text_bytes: 0,
            seen: vec![0],
        })
    }

    /// Read the start tag of a descendant, or a whole empty element when
    /// `empty` is set.
    ///
    /// `namespace` is the element's resolved namespace, `resolver` resolves
    /// the prefixes of its attributes, and `span` is the tag's byte range in
    /// the builder's document.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid modeled value,
    /// malformed attribute syntax on an element whose
    /// attributes are read, or an exceeded [`Limits`] bound.
    pub fn start(
        &mut self,
        element: &BytesStart<'_>,
        namespace: &ResolveResult<'_>,
        resolver: &NamespaceResolver,
        decoder: Decoder,
        span: Range<usize>,
        empty: bool,
    ) -> Result<()> {
        self.count_node()?;
        if span.start > span.end || span.end > self.document.len() {
            return Err(invalid("text body event span is outside its input"));
        }
        if self.skipped > 0 {
            self.body.properties.has_unmodeled_content = true;
            if !empty {
                self.enter_skipped()?;
            }
            return Ok(());
        }
        if is_markup_compatibility(namespace) {
            return Err(invalid(
                "a DrawingML text body contains unprocessed markup compatibility",
            ));
        }
        let local = element.local_name();
        let local = local.as_ref();
        let frame = if is_drawingml(namespace) {
            self.note_unmodeled_attributes(element, local)?;
            self.check_singleton(local)?;
            self.open(local, element, resolver, decoder, &span)
                .map_err(redact)?
        } else {
            None
        };
        match frame {
            Some(frame) if empty => self.close(frame, span.end).map_err(redact),
            Some(frame) => {
                self.check_depth(self.frames.len().saturating_add(1))?;
                self.frames.push(frame);
                self.seen.push(0);
                Ok(())
            },
            None if empty => {
                if !(self.frames.is_empty() && local == b"lstStyle" && is_drawingml(namespace)) {
                    self.body.properties.has_unmodeled_content = true;
                }
                Ok(())
            },
            None => {
                if !(self.frames.is_empty() && local == b"lstStyle" && is_drawingml(namespace)) {
                    self.body.properties.has_unmodeled_content = true;
                }
                self.enter_skipped()
            },
        }
    }

    /// Read the end tag of the most recently opened descendant; `end` is the
    /// byte offset just after it.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Invalid`] for an end tag without an open descendant,
    /// and the errors of reading a captured color.
    pub fn end(&mut self, end: usize) -> Result<()> {
        if self.skipped > 0 {
            self.skipped -= 1;
            return Ok(());
        }
        self.seen.pop();
        let frame = self
            .frames
            .pop()
            .ok_or_else(|| invalid("text body end tag has no open element"))?;
        self.close(frame, end).map_err(redact)
    }

    /// Read decoded character data of the descendants.
    ///
    /// Only text inside a run's or field's `a:t` is kept.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Limit`] when the kept text exceeds
    /// [`Limits::text_bytes`].
    pub fn text(&mut self, text: &str) -> Result<()> {
        if !self.keeps_text() {
            self.note_text_content(text);
            return Ok(());
        }
        self.charge_text(text.len())?;
        self.append_text(text);
        Ok(())
    }

    /// Read XML character data after checking its decoded size.
    ///
    /// XML entity expansion is intentionally accounted for in decoded UTF-8
    /// bytes.  The budget check runs before `unescape` can allocate an
    /// expanded `String`, while valid escaped text keeps its historical
    /// semantic value.
    fn text_xml(&mut self, encoded: &str) -> Result<()> {
        let decoded_bytes = unescaped_text_len(encoded)?;
        let keeps_text = self.keeps_text();
        if keeps_text {
            self.charge_text(decoded_bytes)?;
        }
        let decoded = quick_xml::escape::unescape(encoded)
            .map_err(|_error| Error::Xml("malformed text body XML".into()))?;
        if keeps_text {
            self.append_text(&decoded);
        } else {
            self.note_text_content(&decoded);
        }
        Ok(())
    }

    fn keeps_text(&self) -> bool {
        self.skipped == 0 && self.frames.last() == Some(&Frame::Text)
    }

    fn append_text(&mut self, text: &str) {
        if let Some(run) = self.run.as_mut() {
            run.text.push_str(text);
        }
    }

    fn note_text_content(&mut self, text: &str) {
        if !text
            .trim_matches(|c| matches!(c, ' ' | '\n' | '\r' | '\t'))
            .is_empty()
        {
            self.body.properties.has_unmodeled_content = true;
        }
    }

    /// Finish the body.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Invalid`] when a descendant is still open.
    pub fn finish(self) -> Result<Body> {
        if self.skipped != 0 || !self.frames.is_empty() || self.paragraph.is_some() {
            return Err(invalid("text body ended with unclosed elements"));
        }
        Ok(self.body)
    }

    fn charge_text(&mut self, bytes: usize) -> Result<()> {
        self.text_bytes = self
            .text_bytes
            .checked_add(bytes)
            .filter(|next| *next <= self.limits.text_bytes)
            .ok_or_else(|| {
                Error::limit(
                    crate::LimitKind::TextBytes,
                    self.text_bytes.saturating_add(bytes),
                    self.limits.text_bytes,
                )
            })?;
        Ok(())
    }

    fn note_unmodeled_attributes(&mut self, element: &BytesStart<'_>, local: &[u8]) -> Result<()> {
        let allowed: &[&[u8]] = match local {
            b"bodyPr" => &[
                b"lIns",
                b"tIns",
                b"rIns",
                b"bIns",
                b"anchor",
                b"anchorCtr",
                b"vert",
                b"wrap",
                b"numCol",
                b"spcFirstLastPara",
            ],
            b"pPr" => &[
                b"lvl",
                b"algn",
                b"marL",
                b"marR",
                b"indent",
                b"defTabSz",
                b"rtl",
                b"eaLnBrk",
                b"latinLnBrk",
                b"hangingPunct",
                b"fontAlgn",
            ],
            b"rPr" | b"defRPr" | b"endParaRPr" => &[
                b"b",
                b"i",
                b"u",
                b"strike",
                b"sz",
                b"baseline",
                b"spc",
                b"kern",
                b"cap",
                b"lang",
                b"altLang",
            ],
            b"r" | b"br" | b"p" | b"lstStyle" | b"noAutofit" | b"spAutoFit" | b"normAutofit" => &[],
            b"t" => &[b"xml:space"],
            _ => return Ok(()),
        };
        for attribute in element.checked_attributes() {
            let attribute = attribute.map_err(|_| invalid("malformed text body attribute"))?;
            let key = attribute.key.as_ref();
            if key != b"xmlns" && !key.starts_with(b"xmlns:") && !allowed.contains(&key) {
                self.body.properties.has_unmodeled_content = true;
            }
        }
        Ok(())
    }

    fn check_singleton(&mut self, local: &[u8]) -> Result<()> {
        let slot = match self.frames.last() {
            None => match local {
                b"bodyPr" => 1,
                b"lstStyle" => 2,
                _ => 0,
            },
            Some(Frame::BodyProperties) => match local {
                b"noAutofit" | b"spAutoFit" | b"normAutofit" => 1,
                _ => 0,
            },
            Some(Frame::Paragraph) => match local {
                b"pPr" => 1,
                b"endParaRPr" => 2,
                _ => 0,
            },
            Some(Frame::Run) => match local {
                b"rPr" => 1,
                b"pPr" => 2,
                b"t" => 3,
                _ => 0,
            },
            Some(Frame::ParagraphProperties) => match local {
                b"lnSpc" => 1,
                b"spcBef" => 2,
                b"spcAft" => 3,
                b"buClrTx" | b"buClr" => 4,
                b"buSzTx" | b"buSzPct" | b"buSzPts" => 5,
                b"buFontTx" | b"buFont" => 6,
                b"buNone" | b"buAutoNum" | b"buChar" | b"buBlip" => 7,
                b"tabLst" => 8,
                b"defRPr" => 9,
                _ => 0,
            },
            Some(Frame::RunProperties(_)) => match local {
                b"latin" => 1,
                b"ea" => 2,
                b"cs" => 3,
                b"sym" => 4,
                b"noFill" | b"solidFill" | b"gradFill" | b"pattFill" | b"blipFill" | b"grpFill" => {
                    5
                },
                b"highlight" => 6,
                b"hlinkClick" => 7,
                _ => 0,
            },
            Some(Frame::Spacing(_)) => match local {
                b"spcPct" | b"spcPts" => 1,
                _ => 0,
            },
            Some(Frame::BulletPicture) => match local {
                b"blip" => 1,
                _ => 0,
            },
            Some(Frame::ColorContainer(_))
                if matches!(
                    local,
                    b"srgbClr" | b"scrgbClr" | b"hslClr" | b"sysClr" | b"schemeClr" | b"prstClr"
                ) =>
            {
                1
            },
            _ => 0,
        };
        if slot != 0 {
            let seen = self
                .seen
                .last_mut()
                .ok_or_else(|| invalid("text body state is inconsistent"))?;
            let bit = 1 << slot;
            if *seen & bit != 0 {
                return Err(invalid("text body has duplicate or competing properties"));
            }
            *seen |= bit;
        }
        Ok(())
    }

    fn count_node(&mut self) -> Result<()> {
        self.nodes = self
            .nodes
            .checked_add(1)
            .filter(|nodes| *nodes <= self.limits.nodes)
            .ok_or_else(|| {
                Error::limit(
                    crate::LimitKind::TextBodyElements,
                    self.nodes.saturating_add(1),
                    self.limits.nodes,
                )
            })?;
        Ok(())
    }

    fn check_depth(&self, frames: usize) -> Result<()> {
        if frames.saturating_add(self.skipped) > self.limits.depth {
            return Err(Error::limit(
                crate::LimitKind::TextBodyNestingDepth,
                frames.saturating_add(self.skipped),
                self.limits.depth,
            ));
        }
        Ok(())
    }

    fn enter_skipped(&mut self) -> Result<()> {
        self.skipped = self.skipped.saturating_add(1);
        self.check_depth(self.frames.len())
    }

    fn checked<T>(&self, result: Result<T>) -> Result<Option<T>> {
        result.map(Some).map_err(redact)
    }

    /// The frame a modeled `DrawingML` element opens, or `None` to skip it.
    fn open(
        &mut self,
        local: &[u8],
        element: &BytesStart<'_>,
        resolver: &NamespaceResolver,
        decoder: Decoder,
        span: &Range<usize>,
    ) -> Result<Option<Frame>> {
        let Some(parent) = self.frames.last().copied() else {
            return self.open_body_child(local, element, decoder);
        };
        match parent {
            Frame::BodyProperties => Ok(autofit(local).map(|autofit| {
                self.body.properties.autofit = autofit;
                Frame::Leaf
            })),
            Frame::Paragraph => self.open_paragraph_child(local, element, decoder),
            Frame::ParagraphProperties => {
                self.open_paragraph_properties_child(local, element, decoder)
            },
            Frame::Spacing(slot) => self.open_spacing_child(slot, local, element, decoder),
            Frame::BulletPicture => {
                if local != b"blip" {
                    return Ok(None);
                }
                let id = self.checked(
                    relationship_attribute(element, b"embed", decoder, resolver)
                        .map_err(Error::from),
                )?;
                if let Some(paragraph) = self.paragraph.as_mut() {
                    paragraph.properties.bullet = Some(Bullet::Picture(super::Picture {
                        relationship_id: id.flatten(),
                    }));
                }
                Ok(Some(Frame::Leaf))
            },
            Frame::TabList => {
                if local != b"tab" {
                    return Ok(None);
                }
                let stop = self.tab_stop(element, decoder)?;
                if let Some(stops) = self
                    .paragraph
                    .as_mut()
                    .and_then(|paragraph| paragraph.properties.tab_stops.as_mut())
                {
                    stops.push(stop);
                }
                Ok(Some(Frame::Leaf))
            },
            Frame::Run => self.open_run_child(local, element, decoder),
            Frame::RunProperties(_) => {
                self.open_run_properties_child(local, element, resolver, decoder)
            },
            Frame::ColorContainer(target) => {
                if !is_color_choice(local) {
                    return Ok(None);
                }
                // The first color choice is the color; a later one is
                // skipped as markup the schema does not allow.
                if self.color_is_set(target) {
                    return Ok(None);
                }
                Ok(Some(Frame::Color {
                    target,
                    start: span.start,
                }))
            },
            Frame::Color { .. } | Frame::Text | Frame::Leaf => Ok(None),
        }
    }

    fn open_body_child(
        &mut self,
        local: &[u8],
        element: &BytesStart<'_>,
        decoder: Decoder,
    ) -> Result<Option<Frame>> {
        match local {
            b"bodyPr" => {
                apply_body_properties(&mut self.body.properties, element, decoder)?;
                Ok(Some(Frame::BodyProperties))
            },
            b"p" => {
                self.paragraph = Some(Paragraph::default());
                Ok(Some(Frame::Paragraph))
            },
            _ => Ok(None),
        }
    }

    fn open_paragraph_child(
        &mut self,
        local: &[u8],
        element: &BytesStart<'_>,
        decoder: Decoder,
    ) -> Result<Option<Frame>> {
        match local {
            b"pPr" => {
                self.paragraph_attributes(element, decoder)?;
                Ok(Some(Frame::ParagraphProperties))
            },
            b"r" => {
                self.run = Some(Run::default());
                Ok(Some(Frame::Run))
            },
            b"br" => {
                self.charge_text(1)?;
                self.run = Some(Run {
                    kind: RunKind::Break,
                    text: "\n".to_owned(),
                    properties: RunProperties::default(),
                });
                Ok(Some(Frame::Run))
            },
            b"fld" => {
                let field = self.field(element, decoder)?;
                self.run = Some(Run {
                    kind: RunKind::Field(field),
                    ..Run::default()
                });
                Ok(Some(Frame::Run))
            },
            b"endParaRPr" => self.open_properties(PropertiesTarget::ParagraphEnd, element, decoder),
            _ => Ok(None),
        }
    }

    fn open_paragraph_properties_child(
        &mut self,
        local: &[u8],
        element: &BytesStart<'_>,
        decoder: Decoder,
    ) -> Result<Option<Frame>> {
        let frame = match local {
            b"lnSpc" => Frame::Spacing(SpacingSlot::Line),
            b"spcBef" => Frame::Spacing(SpacingSlot::Before),
            b"spcAft" => Frame::Spacing(SpacingSlot::After),
            b"buClrTx" => {
                self.set_paragraph(|properties| {
                    properties.bullet_color = Some(BulletColor::FollowText);
                });
                Frame::Leaf
            },
            b"buClr" => {
                self.set_paragraph(|properties| properties.bullet_color = None);
                Frame::ColorContainer(ColorTarget::Bullet)
            },
            b"buSzTx" => {
                self.set_paragraph(|properties| {
                    properties.bullet_size = Some(BulletSize::FollowText);
                });
                Frame::Leaf
            },
            b"buSzPct" => {
                let size = self.required_attribute(element, decoder, b"val", |value| {
                    bounded_percent(value, "bullet size percentage", 25_000, 400_000)
                })?;
                if let Some(size) = size {
                    self.set_paragraph(|properties| {
                        properties.bullet_size = Some(BulletSize::Percent(size));
                    });
                }
                Frame::Leaf
            },
            b"buSzPts" => {
                let size = self.required_attribute(element, decoder, b"val", |value| {
                    parse_value::<TextSize>(value, "bullet size points")
                })?;
                if let Some(size) = size {
                    self.set_paragraph(|properties| {
                        properties.bullet_size = Some(BulletSize::Points(size));
                    });
                }
                Frame::Leaf
            },
            b"buFontTx" => {
                self.set_paragraph(|properties| {
                    properties.bullet_font = Some(BulletFont::FollowText);
                });
                Frame::Leaf
            },
            b"buFont" => {
                if let Some(font) = self.font(element, decoder)? {
                    self.set_paragraph(|properties| {
                        properties.bullet_font = Some(BulletFont::Font(font));
                    });
                }
                Frame::Leaf
            },
            b"buNone" => {
                self.set_paragraph(|properties| properties.bullet = Some(Bullet::None));
                Frame::Leaf
            },
            b"buAutoNum" => {
                if let Some(bullet) = self.auto_number(element, decoder)? {
                    self.set_paragraph(|properties| properties.bullet = Some(bullet));
                }
                Frame::Leaf
            },
            b"buChar" => {
                let character = self
                    .required_attribute(element, decoder, b"char", |value| Ok(value.to_owned()))?;
                if let Some(character) = character {
                    self.set_paragraph(|properties| {
                        properties.bullet = Some(Bullet::Character(character));
                    });
                }
                Frame::Leaf
            },
            b"buBlip" => {
                self.set_paragraph(|properties| {
                    properties.bullet = Some(Bullet::Picture(super::Picture {
                        relationship_id: None,
                    }));
                });
                Frame::BulletPicture
            },
            b"tabLst" => {
                self.set_paragraph(|properties| properties.tab_stops = Some(Vec::new()));
                Frame::TabList
            },
            b"defRPr" => {
                return self.open_properties(PropertiesTarget::ParagraphDefault, element, decoder);
            },
            _ => return Ok(None),
        };
        Ok(Some(frame))
    }

    fn open_spacing_child(
        &mut self,
        slot: SpacingSlot,
        local: &[u8],
        element: &BytesStart<'_>,
        decoder: Decoder,
    ) -> Result<Option<Frame>> {
        let spacing = match local {
            b"spcPct" => self
                .required_attribute(element, decoder, b"val", |value| {
                    bounded_percent(value, "spacing percentage", 0, 13_200_000)
                })?
                .map(Spacing::Percent),
            b"spcPts" => self
                .required_attribute(element, decoder, b"val", |value| {
                    bounded_integer::<u32>(value, "spacing points", 0, 158_400)
                })?
                .map(Spacing::Points),
            _ => return Ok(None),
        };
        if let Some(spacing) = spacing {
            self.set_paragraph(|properties| {
                let target = match slot {
                    SpacingSlot::Line => &mut properties.line_spacing,
                    SpacingSlot::Before => &mut properties.space_before,
                    SpacingSlot::After => &mut properties.space_after,
                };
                *target = Some(spacing);
            });
        }
        Ok(Some(Frame::Leaf))
    }

    fn open_run_child(
        &mut self,
        local: &[u8],
        element: &BytesStart<'_>,
        decoder: Decoder,
    ) -> Result<Option<Frame>> {
        let breaks = self
            .run
            .as_ref()
            .is_some_and(|run| matches!(run.kind, RunKind::Break));
        match local {
            b"rPr" => self.open_properties(PropertiesTarget::Run, element, decoder),
            b"t" if !breaks => Ok(Some(Frame::Text)),
            // Run content the schema does not define but producers and the
            // shape text extractors read as a tab or line-break character.
            // (An `a:tab` of an `a:tabLst` is a tab stop, never text.)
            b"tab" | b"br" if !breaks => {
                self.charge_text(1)?;
                self.body.properties.has_unmodeled_content = true;
                if let Some(run) = self.run.as_mut() {
                    run.text.push(if local == b"tab" { '\t' } else { '\n' });
                }
                Ok(Some(Frame::Leaf))
            },
            _ => Ok(None),
        }
    }

    fn open_run_properties_child(
        &mut self,
        local: &[u8],
        element: &BytesStart<'_>,
        resolver: &NamespaceResolver,
        decoder: Decoder,
    ) -> Result<Option<Frame>> {
        let fill = match local {
            b"noFill" => Some(TextFill::None),
            b"solidFill" => Some(TextFill::Solid(None)),
            b"gradFill" => Some(TextFill::Gradient),
            b"blipFill" => Some(TextFill::Picture),
            b"pattFill" => Some(TextFill::Pattern),
            b"grpFill" => Some(TextFill::Group),
            _ => None,
        };
        if let Some(fill) = fill {
            let frame = if matches!(fill, TextFill::Solid(_)) {
                Frame::ColorContainer(ColorTarget::Fill)
            } else {
                Frame::Leaf
            };
            if let Some(pending) = self.pending.as_mut() {
                pending.fill = Some(fill);
            }
            return Ok(Some(frame));
        }
        match local {
            b"highlight" => {
                if let Some(pending) = self.pending.as_mut() {
                    pending.highlight = None;
                }
                Ok(Some(Frame::ColorContainer(ColorTarget::Highlight)))
            },
            b"latin" | b"ea" | b"cs" | b"sym" => {
                let font = self.font(element, decoder)?;
                if let (Some(font), Some(pending)) = (font, self.pending.as_mut()) {
                    let slot = match local {
                        b"latin" => &mut pending.latin,
                        b"ea" => &mut pending.east_asian,
                        b"cs" => &mut pending.complex_script,
                        _ => &mut pending.symbol,
                    };
                    *slot = Some(font);
                }
                Ok(Some(Frame::Leaf))
            },
            b"hlinkClick" => {
                if let Some(hyperlink) = self.hyperlink(element, resolver, decoder)?
                    && let Some(pending) = self.pending.as_mut()
                {
                    pending.hyperlink = Some(hyperlink);
                }
                Ok(Some(Frame::Leaf))
            },
            _ => Ok(None),
        }
    }

    /// Open an `a:rPr`-like element for `target`, starting from the
    /// target's current properties. Duplicate property elements are refused.
    fn open_properties(
        &mut self,
        target: PropertiesTarget,
        element: &BytesStart<'_>,
        decoder: Decoder,
    ) -> Result<Option<Frame>> {
        let mut properties = match target {
            PropertiesTarget::Run => self.run.as_ref().map(|run| run.properties.clone()),
            PropertiesTarget::ParagraphDefault => self
                .paragraph
                .as_ref()
                .and_then(|paragraph| paragraph.properties.default_run.as_deref().cloned()),
            PropertiesTarget::ParagraphEnd => self
                .paragraph
                .as_ref()
                .and_then(|paragraph| paragraph.end_properties.as_deref().cloned()),
        }
        .unwrap_or_default();
        self.run_attributes(&mut properties, element, decoder)?;
        self.pending = Some(properties);
        Ok(Some(Frame::RunProperties(target)))
    }

    fn close(&mut self, frame: Frame, end: usize) -> Result<()> {
        match frame {
            Frame::Paragraph => {
                if let Some(paragraph) = self.paragraph.take() {
                    self.body.paragraphs.push(paragraph);
                }
            },
            Frame::Run => {
                if let (Some(run), Some(paragraph)) = (self.run.take(), self.paragraph.as_mut()) {
                    paragraph.runs.push(run);
                }
            },
            Frame::RunProperties(target) => {
                let properties = self.pending.take().unwrap_or_default();
                match target {
                    PropertiesTarget::Run => {
                        if let Some(run) = self.run.as_mut() {
                            run.properties = properties;
                        }
                    },
                    PropertiesTarget::ParagraphDefault => self.set_paragraph(|paragraph| {
                        paragraph.default_run = Some(Box::new(properties));
                    }),
                    PropertiesTarget::ParagraphEnd => {
                        if let Some(paragraph) = self.paragraph.as_mut() {
                            paragraph.end_properties = Some(Box::new(properties));
                        }
                    },
                }
            },
            Frame::Color { target, start } => {
                let fragment = self
                    .document
                    .get(start..end)
                    .ok_or_else(|| invalid("text body color lies outside its document"))?;
                // The fragment reader types the color and keeps an unknown
                // choice or transform losslessly. Its scalar domains are
                // narrower than the schema's in places (a `satMod` above
                // 100%, which producers write, is refused), so a fragment
                // it refuses is kept as written rather than failing the
                // body.
                let color_limits = color::Limits::new(
                    self.limits.input_bytes().clamp(1, color::MAX_XML_BYTES),
                    self.limits.depth().clamp(1, color::MAX_DEPTH),
                    self.limits.nodes().clamp(1, color::MAX_NODES),
                    self.limits.nodes().clamp(1, color::MAX_TRANSFORMS),
                );
                let value = match color::read(fragment, &color_limits) {
                    Ok(value) => Some(value),
                    Err(error @ (Error::Limit { .. } | Error::InvalidLimit { .. })) => {
                        return Err(error);
                    },
                    Err(_) => self.checked(color::Value::unknown(fragment, &color_limits))?,
                };
                if let Some(value) = value {
                    self.set_color(target, value);
                }
            },
            Frame::ColorContainer(target) => {
                if !self.color_is_set(target) {
                    self.body.properties.has_unmodeled_content = true;
                }
            },
            Frame::BodyProperties
            | Frame::ParagraphProperties
            | Frame::Spacing(_)
            | Frame::BulletPicture
            | Frame::TabList
            | Frame::Text
            | Frame::Leaf => {},
        }
        Ok(())
    }

    fn color_is_set(&self, target: ColorTarget) -> bool {
        match target {
            ColorTarget::Fill => self
                .pending
                .as_ref()
                .is_some_and(|pending| pending.color().is_some()),
            ColorTarget::Highlight => self
                .pending
                .as_ref()
                .is_some_and(|pending| pending.highlight.is_some()),
            ColorTarget::Bullet => self.paragraph.as_ref().is_some_and(|paragraph| {
                matches!(
                    paragraph.properties.bullet_color,
                    Some(BulletColor::Color(_))
                )
            }),
        }
    }

    fn set_color(&mut self, target: ColorTarget, value: color::Value) {
        match target {
            ColorTarget::Fill => {
                if let Some(pending) = self.pending.as_mut() {
                    pending.fill = Some(TextFill::Solid(Some(value)));
                }
            },
            ColorTarget::Highlight => {
                if let Some(pending) = self.pending.as_mut() {
                    pending.highlight = Some(value);
                }
            },
            ColorTarget::Bullet => self.set_paragraph(|properties| {
                properties.bullet_color = Some(BulletColor::Color(value));
            }),
        }
    }

    fn set_paragraph(&mut self, update: impl FnOnce(&mut super::ParagraphProperties)) {
        if let Some(paragraph) = self.paragraph.as_mut() {
            update(&mut paragraph.properties);
        }
    }

    /// Read one required attribute and parse it; a missing attribute is an
    /// invalid value.
    fn required_attribute<T>(
        &self,
        element: &BytesStart<'_>,
        decoder: Decoder,
        name: &[u8],
        parse: impl FnOnce(&str) -> Result<T>,
    ) -> Result<Option<T>> {
        let value = unqualified_attribute_value(element, name, decoder)
            .map_err(Error::from)
            .and_then(|value| {
                value.ok_or_else(|| {
                    invalid(format!(
                        "DrawingML {} is missing its {} attribute",
                        String::from_utf8_lossy(element.local_name().as_ref()),
                        String::from_utf8_lossy(name)
                    ))
                })
            })
            .and_then(|value| parse(&value));
        self.checked(value)
    }

    /// Read the attributes of an `a:pPr` into the open paragraph.
    fn paragraph_attributes(&mut self, element: &BytesStart<'_>, decoder: Decoder) -> Result<()> {
        let mut properties = self
            .paragraph
            .as_ref()
            .map(|paragraph| paragraph.properties.clone())
            .unwrap_or_default();
        let read = for_each_attribute(element, |local, attribute| {
            let decoded = match self.checked(decode(&attribute, decoder))? {
                Some(decoded) => decoded,
                None => return Ok(()),
            };
            let decoded = decoded.as_ref();
            match local {
                b"lvl" => {
                    properties.level =
                        self.checked(bounded_integer::<u8>(decoded, "paragraph level", 0, 8))?;
                },
                b"algn" => {
                    properties.alignment =
                        self.checked(token::<Alignment>(decoded, "paragraph alignment"))?;
                },
                b"marL" => {
                    properties.margin_left = self.checked(bounded_integer::<i32>(
                        decoded,
                        "left margin",
                        0,
                        51_206_400,
                    ))?;
                },
                b"marR" => {
                    properties.margin_right = self.checked(bounded_integer::<i32>(
                        decoded,
                        "right margin",
                        0,
                        51_206_400,
                    ))?;
                },
                b"indent" => {
                    properties.indent = self.checked(bounded_integer::<i32>(
                        decoded,
                        "indent",
                        -51_206_400,
                        51_206_400,
                    ))?;
                },
                b"defTabSz" => {
                    properties.default_tab_size =
                        self.checked(parse_value::<Coordinate32>(decoded, "default tab size"))?;
                },
                b"rtl" => {
                    properties.right_to_left =
                        self.checked(parse_dml_bool(decoded, "paragraph rtl"))?;
                },
                b"eaLnBrk" => {
                    properties.east_asian_line_break =
                        self.checked(parse_dml_bool(decoded, "eaLnBrk"))?;
                },
                b"latinLnBrk" => {
                    properties.latin_line_break =
                        self.checked(parse_dml_bool(decoded, "latinLnBrk"))?;
                },
                b"hangingPunct" => {
                    properties.hanging_punctuation =
                        self.checked(parse_dml_bool(decoded, "hangingPunct"))?;
                },
                b"fontAlgn" => {
                    properties.font_alignment =
                        self.checked(token::<FontAlignment>(decoded, "font alignment"))?;
                },
                _ => {},
            }
            Ok(())
        });
        self.checked(read)?;
        if let Some(paragraph) = self.paragraph.as_mut() {
            paragraph.properties = properties;
        }
        Ok(())
    }

    /// Read the attributes of an `a:rPr`-like element into `properties`.
    ///
    fn run_attributes(
        &self,
        properties: &mut RunProperties,
        element: &BytesStart<'_>,
        decoder: Decoder,
    ) -> Result<()> {
        let read = for_each_attribute(element, |local, attribute| {
            let decoded = match self.checked(decode(&attribute, decoder))? {
                Some(decoded) => decoded,
                None => return Ok(()),
            };
            let decoded = decoded.as_ref();
            match local {
                b"b" => {
                    properties.bold = self.checked(parse_dml_bool(decoded, "run bold"))?;
                },
                b"i" => {
                    properties.italic = self.checked(parse_dml_bool(decoded, "run italic"))?;
                },
                b"u" => {
                    properties.underline = self.checked(parse_underline(decoded))?;
                },
                b"sz" => {
                    properties.size =
                        self.checked(parse_value::<TextSize>(decoded, "text size"))?;
                },
                b"strike" => {
                    properties.strike = self.checked(token::<Strike>(decoded, "strike"))?;
                },
                b"baseline" => {
                    properties.baseline = self.checked(color::parse_scaled(decoded, "baseline"))?;
                },
                b"spc" => {
                    properties.spacing =
                        self.checked(parse_value::<TextPoint>(decoded, "character spacing"))?;
                },
                b"kern" => {
                    properties.kerning = self.checked(
                        parse_value::<TextPoint>(decoded, "kerning")
                            .and_then(|kerning| non_negative(kerning, decoded, "kerning")),
                    )?;
                },
                b"cap" => {
                    properties.caps = self.checked(token::<Caps>(decoded, "caps"))?;
                },
                b"lang" => properties.language = Some(decoded.to_owned()),
                b"altLang" => properties.alternate_language = Some(decoded.to_owned()),
                _ => {},
            }
            Ok(())
        });
        self.checked(read).map(|_| ())
    }

    fn font(&self, element: &BytesStart<'_>, decoder: Decoder) -> Result<Option<Font>> {
        let mut font = Font::default();
        let mut typeface = false;
        let read = for_each_attribute(element, |local, attribute| {
            let decoded = match self.checked(decode(&attribute, decoder))? {
                Some(decoded) => decoded,
                None => return Ok(()),
            };
            let decoded = decoded.as_ref();
            match local {
                b"typeface" => {
                    font.typeface = decoded.to_owned();
                    typeface = true;
                },
                b"panose" => font.panose = self.checked(parse_panose(decoded))?,
                b"pitchFamily" => {
                    font.pitch_family =
                        self.checked(parse_value::<i8>(decoded, "font pitch family"))?;
                },
                b"charset" => {
                    font.charset =
                        self.checked(parse_value::<i8>(decoded, "font character set"))?;
                },
                _ => {},
            }
            Ok(())
        });
        if self.checked(read)?.is_none() {
            return Ok(None);
        }
        if !typeface {
            self.checked(Err::<(), _>(invalid(format!(
                "DrawingML {} font is missing its typeface",
                String::from_utf8_lossy(element.local_name().as_ref())
            ))))?;
            return Ok(None);
        }
        Ok(Some(font))
    }

    fn hyperlink(
        &self,
        element: &BytesStart<'_>,
        resolver: &NamespaceResolver,
        decoder: Decoder,
    ) -> Result<Option<Hyperlink>> {
        let relationship_id = relationship_attribute(element, b"id", decoder, resolver)
            .map_err(|error| redact(error.into()))?;
        let action = unqualified_attribute_value(element, b"action", decoder)
            .map_err(|error| redact(error.into()))?;
        let tooltip = unqualified_attribute_value(element, b"tooltip", decoder)
            .map_err(|error| redact(error.into()))?;
        Ok(Some(Hyperlink {
            relationship_id,
            action,
            tooltip,
        }))
    }

    fn field(&self, element: &BytesStart<'_>, decoder: Decoder) -> Result<Field> {
        let id = self.required_attribute(element, decoder, b"id", |value| {
            if litchi_ooxml_common::custom_xml::valid_guid(value) {
                Ok(value.to_owned())
            } else {
                Err(invalid("invalid text field identifier"))
            }
        })?;
        let kind = self
            .checked(unqualified_attribute_value(element, b"type", decoder).map_err(Error::from))?;
        Ok(Field {
            id: id.unwrap_or_default(),
            kind: kind.flatten(),
        })
    }

    fn auto_number(&self, element: &BytesStart<'_>, decoder: Decoder) -> Result<Option<Bullet>> {
        let scheme = self.required_attribute(element, decoder, b"type", |value| {
            token::<AutoNumber>(value, "auto-number scheme")
        })?;
        let Some(scheme) = scheme else {
            return Ok(None);
        };
        let start_at = self
            .checked(
                unqualified_attribute_value(element, b"startAt", decoder)
                    .map_err(Error::from)
                    .and_then(|value| {
                        value
                            .map(|value| {
                                bounded_integer::<u16>(&value, "auto-number start", 1, 32_767)
                            })
                            .transpose()
                    }),
            )?
            .flatten();
        Ok(Some(Bullet::AutoNumber { scheme, start_at }))
    }

    fn tab_stop(&self, element: &BytesStart<'_>, decoder: Decoder) -> Result<TabStop> {
        let position = self
            .checked(
                unqualified_attribute_value(element, b"pos", decoder)
                    .map_err(Error::from)
                    .and_then(|value| {
                        value
                            .map(|value| parse_value::<Coordinate32>(&value, "tab position"))
                            .transpose()
                    }),
            )?
            .flatten();
        let alignment = self
            .checked(
                unqualified_attribute_value(element, b"algn", decoder)
                    .map_err(Error::from)
                    .and_then(|value| {
                        value
                            .map(|value| token::<TabAlignment>(&value, "tab alignment"))
                            .transpose()
                    }),
            )?
            .flatten();
        Ok(TabStop {
            position,
            alignment,
        })
    }
}

/// Visit the unprefixed attributes of `element` in document order.
fn for_each_attribute<'e>(
    element: &'e BytesStart<'_>,
    mut visit: impl FnMut(&[u8], Attribute<'e>) -> Result<()>,
) -> Result<()> {
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|_error| {
            Error::Common(XmlError::Malformed("malformed text body attribute".into()))
        })?;
        if attribute.key.prefix().is_some() {
            continue;
        }
        let local = attribute.key.local_name().into_inner();
        visit(local, attribute)?;
    }
    Ok(())
}

fn decode<'v>(attribute: &Attribute<'v>, decoder: Decoder) -> Result<Cow<'v, str>> {
    attribute
        .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
        .map_err(|_error| {
            Error::Common(XmlError::Malformed("malformed text body attribute".into()))
        })
}

fn unescaped_text_len(encoded: &str) -> Result<usize> {
    let mut total = 0usize;
    let mut offset = 0usize;
    while let Some(relative) = encoded[offset..].find('&') {
        let ampersand = offset + relative;
        total = total
            .checked_add(ampersand.saturating_sub(offset))
            .ok_or_else(|| Error::Xml("text body decoded text length overflows".into()))?;
        let end = encoded[ampersand + 1..]
            .find(';')
            .map(|relative| ampersand + 1 + relative + 1)
            .ok_or_else(|| Error::Xml("malformed text body XML".into()))?;
        let entity = quick_xml::escape::unescape(&encoded[ampersand..end])
            .map_err(|_error| Error::Xml("malformed text body XML".into()))?;
        total = total
            .checked_add(entity.len())
            .ok_or_else(|| Error::Xml("text body decoded text length overflows".into()))?;
        offset = end;
    }
    total = total
        .checked_add(encoded.len().saturating_sub(offset))
        .ok_or_else(|| Error::Xml("text body decoded text length overflows".into()))?;
    Ok(total)
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn parse_value<T: FromStr>(value: &str, description: &str) -> Result<T> {
    value
        .parse()
        .map_err(|_error| invalid(format!("invalid {description}")))
}

fn parse_dml_bool(value: &str, attribute: &str) -> Result<bool> {
    parse_bool(value)
        .map_err(|error| invalid(format!("invalid DrawingML {attribute} boolean: {error}")))
}

fn parse_underline(value: &str) -> Result<Underline> {
    Underline::from_dml(value)
        .map_err(|error| invalid(format!("invalid DrawingML underline: {error}")))
}

/// Parse a closed token after the XML Schema `token` whitespace facet.
fn token<T: FromStr>(value: &str, description: &str) -> Result<T> {
    xsd_token_atom(value)
        .and_then(|atom| atom.parse().ok())
        .ok_or_else(|| invalid(format!("invalid DrawingML {description}")))
}

/// Parse an integer, after the XML Schema whitespace facet, within
/// `min..=max`.
fn bounded_integer<T>(value: &str, description: &str, min: T, max: T) -> Result<T>
where
    T: FromStr + PartialOrd + Copy,
{
    value
        .trim_matches(|character| matches!(character, ' ' | '\t' | '\n' | '\r'))
        .parse::<T>()
        .ok()
        .filter(|parsed| *parsed >= min && *parsed <= max)
        .ok_or_else(|| invalid(format!("invalid DrawingML {description}")))
}

/// Parse a thousandth-of-a-percent value (or Office's percent-sign form)
/// within `min..=max`.
fn bounded_percent(value: &str, description: &'static str, min: u32, max: u32) -> Result<u32> {
    let parsed = color::parse_scaled(value, description)?;
    u32::try_from(parsed)
        .ok()
        .filter(|parsed| *parsed >= min && *parsed <= max)
        .ok_or_else(|| invalid(format!("invalid DrawingML {description}")))
}

fn non_negative(value: TextPoint, _raw: &str, description: &str) -> Result<TextPoint> {
    if value.get() < 0 {
        return Err(invalid(format!("invalid DrawingML {description}")));
    }
    Ok(value)
}

/// Check an `ST_Panose` value: ten bytes as twenty hexadecimal digits.
fn parse_panose(value: &str) -> Result<String> {
    if value.len() == 20 && value.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        Ok(value.to_owned())
    } else {
        Err(invalid("invalid DrawingML font panose"))
    }
}

fn is_drawingml(namespace: &ResolveResult<'_>) -> bool {
    matches!(
        namespace,
        ResolveResult::Bound(Namespace(value))
            if *value == DRAWINGML_NAMESPACE || *value == STRICT_DRAWINGML_NAMESPACE
    )
}

fn is_markup_compatibility(namespace: &ResolveResult<'_>) -> bool {
    matches!(
        namespace,
        ResolveResult::Bound(Namespace(value))
            if *value == litchi_ooxml_common::mce::NAMESPACE.as_bytes()
    )
}

fn is_color_choice(local: &[u8]) -> bool {
    matches!(
        local,
        b"srgbClr" | b"schemeClr" | b"scrgbClr" | b"hslClr" | b"sysClr" | b"prstClr"
    )
}

fn position(reader: &NsReader<&[u8]>, origin: ReaderOrigin) -> Result<usize> {
    origin
        .offset(reader.buffer_position())
        .ok_or_else(|| invalid("text body position does not fit usize"))
}

fn redact(error: Error) -> Error {
    match error {
        Error::Invalid(_) => Error::Invalid("invalid text body value".into()),
        Error::Xml(_) => Error::Xml("malformed text body XML".into()),
        Error::Common(XmlError::Invalid(_)) => {
            Error::Common(XmlError::Invalid("invalid text body XML".into()))
        },
        Error::Common(XmlError::Malformed(_)) => {
            Error::Common(XmlError::Malformed("malformed text body XML".into()))
        },
        error => error,
    }
}

impl std::fmt::Debug for Builder<'_> {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("Builder")
            .field("nodes", &self.nodes)
            .finish_non_exhaustive()
    }
}

fn relationship_attribute(
    element: &BytesStart<'_>,
    name: &[u8],
    _decoder: Decoder,
    resolver: &NamespaceResolver,
) -> std::result::Result<Option<String>, XmlError> {
    const TRANS: &[u8] = b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const STRICT: &[u8] = b"http://purl.oclc.org/ooxml/officeDocument/relationships";
    let (owner, _) = resolver.resolve_element(element.name());
    let expected = if matches!(owner, ResolveResult::Bound(Namespace(uri)) if uri == STRICT_DRAWINGML_NAMESPACE)
    {
        STRICT
    } else {
        TRANS
    };
    let mut found = None;
    for attribute in element.checked_attributes() {
        let attribute = attribute
            .map_err(|_| XmlError::Malformed("malformed relationship attribute".into()))?;
        if attribute.key.local_name().as_ref() != name || attribute.key.prefix().is_none() {
            continue;
        }
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        match namespace {
            ResolveResult::Bound(Namespace(uri)) if uri == expected => {
                if found.is_some() {
                    return Err(XmlError::Invalid("ambiguous text relationship".into()));
                }
                found = Some(
                    attribute
                        .normalized_value(XmlVersion::Explicit1_0)
                        .map_err(|_| XmlError::Malformed("malformed relationship value".into()))?
                        .into_owned(),
                );
            },
            ResolveResult::Bound(Namespace(uri)) if uri == TRANS || uri == STRICT => {
                return Err(XmlError::Invalid(
                    "text relationship uses the wrong namespace".into(),
                ));
            },
            ResolveResult::Unknown(_) => {
                return Err(XmlError::Invalid(
                    "text relationship has an undeclared prefix".into(),
                ));
            },
            _ => {},
        }
    }
    Ok(found)
}
