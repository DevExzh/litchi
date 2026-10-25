#![allow(
    clippy::arbitrary_source_item_ordering,
    clippy::map_err_ignore,
    clippy::match_same_arms,
    clippy::shadow_reuse,
    clippy::similar_names,
    clippy::too_many_lines,
    clippy::wildcard_enum_match_arm,
    clippy::wildcard_imports,
    reason = "The streaming parser keeps its grammar and stable diagnostics together."
)]

//! Bounded `SpreadsheetDrawingML` XML codec.
//!
//! Parsing is inert and namespace-aware. Every source is subject to byte,
//! depth, text, object, group, and anchor budgets; unmodeled fragments are
//! retained as bounded opaque values.

use std::ops::Range;
use std::str::FromStr;

use litchi_core::xml::ReaderOrigin;
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, QName, ResolveResult};
use quick_xml::reader::NsReader;
use quick_xml::writer::Writer;

use crate::{Error, LimitKind, Result};
use litchi_drawingml::geometry::reader::{CustomGeometryBuilder, GeometryElement};
use litchi_drawingml::text::body::reader::{self as text_reader, Builder as TextBuilder};
use litchi_drawingml::text::parse_bool;
use litchi_ooxml_common::XmlError;
use litchi_ooxml_common::relationships::attribute_value;
use litchi_ooxml_common::xml::{
    decode_xml_reference, is_drawingml_name, unqualified_attribute_value, xsd_token_atom,
};

use super::model::*;
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;

const SPREADSHEET_DRAWING_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const STRICT_SPREADSHEET_DRAWING_NAMESPACE: &[u8] =
    b"http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing";

/// Parse one `SpreadsheetDrawing` part into its anchored shape inventory.
///
/// Returns `Ok(None)` when the part has no `xdr:wsDr` root. Markup-
/// compatibility processing is applied so `mc:AlternateContent` fallbacks
/// resolve before parsing. Pictures and chart graphic frames are skipped;
/// structurally invalid anchors are errors.
///
/// # Errors
///
/// Returns an error for malformed XML, invalid drawing markup, or configured limits.
pub fn read(xml: &str, limits: &Limits) -> Result<Option<Vec<AnchoredObject>>> {
    let limits = limits.validate()?;
    if xml.len() > limits.max_drawing_part_bytes {
        return Err(limit(
            LimitKind::DrawingBytes,
            xml.len(),
            limits.max_drawing_part_bytes,
        ));
    }
    let xml = litchi_ooxml_common::mce::process_str(xml)?;
    if xml.len() > limits.max_drawing_part_bytes {
        return Err(limit(
            LimitKind::DrawingBytes,
            xml.len(),
            limits.max_drawing_part_bytes,
        ));
    }
    Parser::parse(xml.as_ref(), limits)
}

/// Decode a relationship-namespace attribute while rejecting duplicates.
fn relationship_attribute_value(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<Option<String>> {
    attribute_value(element, name, decoder, resolver)
        .map_err(|_| malformed_xml("malformed relationship attribute"))
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Context {
    Root,
    Anchor,
    From,
    To,
    Marker(MarkerTarget, MarkerField),
    Object,
    CustomGeometry(GeometryElement),
    Body,
    Properties,
    Paragraph,
    Run,
    RunProperties,
    Text,
    Unknown,
    Other,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum UnknownTarget {
    AnchorObject,
    GroupChild,
    ObjectMarkup,
}

struct UnknownCapture {
    depth: usize,
    target: UnknownTarget,
    writer: Writer<Vec<u8>>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum MarkerTarget {
    From,
    To,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum MarkerField {
    Column,
    ColumnOffset,
    Row,
    RowOffset,
}

#[derive(Default)]
struct Marker {
    column: Option<u32>,
    column_offset: Option<i64>,
    row: Option<u32>,
    row_offset: Option<i64>,
}

impl Marker {
    fn finish(self, description: &str) -> Result<CellMarker> {
        Ok(CellMarker {
            column: self
                .column
                .ok_or_else(|| invalid(format!("{description} is missing its column")))?,
            column_offset: Emu(self
                .column_offset
                .ok_or_else(|| invalid(format!("{description} is missing its column offset")))?),
            row: self
                .row
                .ok_or_else(|| invalid(format!("{description} is missing its row")))?,
            row_offset: Emu(self
                .row_offset
                .ok_or_else(|| invalid(format!("{description} is missing its row offset")))?),
        })
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum AnchorKind {
    TwoCell,
    OneCell,
    Absolute,
}

#[derive(Default)]
struct PendingAnchor {
    kind: Option<AnchorKind>,
    edit_as: EditAs,
    from: Option<Marker>,
    to: Option<Marker>,
    position: Option<EmuOffset>,
    extent: Option<EmuExtent>,
    object: Option<Object>,
    client_data: ClientData,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum BuilderKind {
    Shape,
    Connection,
    Group,
    GraphicFrame,
}

#[derive(Default)]
struct BodyBuilder {
    properties: Properties,
    /// Paragraphs the shared text-body reader produced for `xdr:txBody`.
    paragraphs: Vec<Paragraph>,
    in_text_body: bool,
}

/// The shared `DrawingML` text-body reader for the open `xdr:txBody`.
///
/// Every event inside the body reaches it, including the ones this parser
/// also retains as opaque markup, so the paragraphs, runs and their local
/// properties come from the one shared implementation. It uses the strict
/// shared contract: modeled values are validated, fields are semantic runs,
/// and the host still retains the original field markup as opaque data.
struct TextBody<'x> {
    builder: TextBuilder<'x>,
    /// Parser stack length while `xdr:txBody` is open.
    depth: usize,
}

struct ObjectBuilder {
    kind: BuilderKind,
    non_visual: NonVisual,
    is_text_box: bool,
    geometry: Option<Geometry>,
    geometry_builder: Option<CustomGeometryBuilder>,
    start: Option<ConnectionEnd>,
    end: Option<ConnectionEnd>,

    transform: GroupTransform,
    saw_transform: bool,
    text_body: BodyBuilder,
    children: Vec<Object>,
    ole_object: Option<OleObject>,
}

impl ObjectBuilder {
    fn new(kind: BuilderKind) -> Self {
        Self {
            kind,
            non_visual: NonVisual::default(),
            is_text_box: false,
            geometry: None,
            geometry_builder: None,
            start: None,
            end: None,
            transform: GroupTransform::default(),
            saw_transform: false,
            text_body: BodyBuilder::default(),
            children: Vec::new(),
            ole_object: None,
        }
    }

    fn finish(self) -> Option<Object> {
        let text_body = if self.text_body.in_text_body {
            Some(Body {
                properties: self.text_body.properties,
                paragraphs: self.text_body.paragraphs,
            })
        } else {
            None
        };
        match self.kind {
            BuilderKind::Shape => Some(Object::Shape(Shape {
                non_visual: self.non_visual,
                is_text_box: self.is_text_box,
                geometry: self.geometry,
                text_body,
            })),
            BuilderKind::Connection => Some(Object::ConnectionShape(ConnectionShape {
                non_visual: self.non_visual,
                geometry: self.geometry,
                start: self.start,
                end: self.end,
                text_body,
            })),
            BuilderKind::Group => Some(Object::Group(Group {
                non_visual: self.non_visual,
                transform: self.saw_transform.then_some(self.transform),
                children: self.children,
            })),
            // Graphic frames only surface when they host a legacy OLE object;
            // chart frames are covered by the chart support.
            BuilderKind::GraphicFrame => self.ole_object.map(|mut ole_object| {
                ole_object.non_visual = self.non_visual;
                Object::OleObject(ole_object)
            }),
        }
    }
}

struct Parser<'x> {
    document: &'x [u8],
    limits: Limits,
    objects: Vec<AnchoredObject>,
    object_count: usize,
    text_bytes: usize,
    anchor: Option<PendingAnchor>,
    builders: Vec<ObjectBuilder>,
    marker_text: String,
    unknown: Option<UnknownCapture>,
    text: Option<TextBody<'x>>,
}

impl<'x> Parser<'x> {
    fn parse(xml: &'x str, limits: Limits) -> Result<Option<Vec<AnchoredObject>>> {
        let document = xml.as_bytes();
        let mut reader = NsReader::from_reader(document);
        reader.config_mut().trim_text(false);
        let origin = ReaderOrigin::of(document);
        let mut parser = Self {
            document,
            limits,
            objects: Vec::new(),
            object_count: 0,
            text_bytes: 0,
            anchor: None,
            builders: Vec::new(),
            marker_text: String::new(),
            unknown: None,
            text: None,
        };
        let mut stack = Vec::new();
        let mut closed_root = false;
        loop {
            let decoder = reader.decoder();
            let start = position(&reader, origin)?;
            let event = reader
                .read_event()
                .map_err(|_| malformed_xml("malformed SpreadsheetML drawing XML"))?
                .into_owned();
            let span = start..position(&reader, origin)?;
            let resolver = reader.resolver().clone();
            let (namespace, event) = resolver.resolve_event(event);
            if parser.unknown.is_some() {
                let complete = parser.capture_event(&event)?;
                if complete {
                    let context = stack.pop().ok_or_else(|| {
                        invalid("unknown drawing element closes outside its owner")
                    })?;
                    if context != Context::Unknown {
                        return Err(invalid(
                            "unknown drawing capture has an invalid stack owner",
                        ));
                    }
                    parser.finish(context)?;
                }
                // Captured markup inside a text body still reaches the
                // shared reader, which models paragraph properties, fields,
                // line breaks and end-of-paragraph properties. Character
                // data is forwarded as well so a captured field retains its
                // semantic cached result while its original XML stays raw.
                parser.forward_markup(&event, &namespace, &resolver, decoder, span)?;
                continue;
            }
            match event {
                Event::Start(element) if stack.is_empty() => {
                    if closed_root {
                        return Err(invalid("drawing XML contains multiple root elements"));
                    }
                    if !is_xdr_name(&namespace, element.name(), b"wsDr") {
                        return Ok(None);
                    }
                    stack.push(Context::Root);
                },
                Event::Empty(element) if stack.is_empty() => {
                    if !is_xdr_name(&namespace, element.name(), b"wsDr") {
                        return Ok(None);
                    }
                    return Ok(Some(parser.objects));
                },
                Event::Start(element) => {
                    let parent = *stack
                        .last()
                        .ok_or_else(|| invalid("missing drawing root"))?;
                    let context =
                        parser.start(parent, &namespace, &element, decoder, &resolver, false)?;
                    stack.push(context);
                    if stack.len() > parser.limits.max_xml_depth {
                        return Err(limit(
                            LimitKind::XmlDepth,
                            stack.len(),
                            parser.limits.max_xml_depth,
                        ));
                    }
                    if context == Context::Body {
                        parser.open_text(stack.len())?;
                    } else {
                        parser
                            .forward_start(&element, &namespace, &resolver, decoder, span, false)?;
                    }
                },
                Event::Empty(element) => {
                    let parent = *stack
                        .last()
                        .ok_or_else(|| invalid("missing drawing root"))?;
                    let context =
                        parser.start(parent, &namespace, &element, decoder, &resolver, true)?;
                    parser.finish(context)?;
                    if context != Context::Body {
                        parser
                            .forward_start(&element, &namespace, &resolver, decoder, span, true)?;
                    }
                },
                Event::Text(text) => {
                    let decoded = text
                        .xml_content(XmlVersion::Explicit1_0)
                        .map_err(|_| encoding_error("invalid SpreadsheetML drawing text"))?;
                    match stack.last() {
                        Some(Context::Marker(_, _)) => parser.marker_text.push_str(&decoded),
                        Some(Context::Text) => parser.push_run_text(&decoded)?,
                        _ => {},
                    }
                },
                Event::CData(text) => {
                    let decoded = text
                        .xml_content(XmlVersion::Explicit1_0)
                        .map_err(|_| encoding_error("invalid SpreadsheetML drawing text"))?;
                    match stack.last() {
                        Some(Context::Marker(_, _)) => parser.marker_text.push_str(&decoded),
                        Some(Context::Text) => parser.push_run_text(&decoded)?,
                        _ => {},
                    }
                },
                Event::GeneralRef(reference) => {
                    let decoded = decode_xml_reference(&reference)
                        .map_err(|_| malformed_xml("invalid SpreadsheetML drawing reference"))?;
                    match stack.last() {
                        Some(Context::Marker(_, _)) => parser.marker_text.push_str(&decoded),
                        Some(Context::Text) => parser.push_run_text(&decoded)?,
                        _ => {},
                    }
                },
                Event::End(element) => {
                    let closes_text = parser
                        .text
                        .as_ref()
                        .is_some_and(|text| text.depth == stack.len());
                    let context = stack
                        .pop()
                        .ok_or_else(|| invalid("drawing XML closes outside its root"))?;
                    parser.finish(context)?;
                    if !closes_text {
                        parser.forward_end(span.end)?;
                    }
                    if context == Context::Root {
                        if !is_xdr_name(&namespace, element.name(), b"wsDr") {
                            return Err(invalid("drawing XML has an invalid root closing element"));
                        }
                        closed_root = true;
                    }
                },
                Event::DocType(_) | Event::PI(_) => {
                    return Err(invalid("DTDs and processing instructions are rejected"));
                },
                Event::Eof if !closed_root || !stack.is_empty() => {
                    return Err(invalid("drawing XML has an unterminated root"));
                },
                Event::Eof => break,
                _ => {},
            }
        }
        Ok(Some(parser.objects))
    }

    /// Charge run text against the drawing's text budget and hand it to the
    /// shared text-body reader.
    fn push_run_text(&mut self, text: &str) -> Result<()> {
        self.text_bytes = self
            .text_bytes
            .checked_add(text.len())
            .ok_or_else(|| limit(LimitKind::TextBytes, usize::MAX, self.limits.max_text_bytes))?;
        if self.text_bytes > self.limits.max_text_bytes {
            return Err(limit(
                LimitKind::TextBytes,
                self.text_bytes,
                self.limits.max_text_bytes,
            ));
        }
        if let Some(body) = self.text.as_mut() {
            body.builder.text(text).map_err(text_error)?;
        }
        Ok(())
    }

    /// Start the shared text-body reader for the `xdr:txBody` just opened at
    /// parser stack length `depth`.
    fn open_text(&mut self, depth: usize) -> Result<()> {
        self.text = Some(TextBody {
            builder: TextBuilder::new(self.document, text_limits(self.limits)?)?,
            depth,
        });
        Ok(())
    }

    /// Hand one start tag or empty element inside the open text body to the
    /// shared reader.
    fn forward_start(
        &mut self,
        element: &BytesStart<'_>,
        namespace: &ResolveResult<'_>,
        resolver: &NamespaceResolver,
        decoder: Decoder,
        span: Range<usize>,
        empty: bool,
    ) -> Result<()> {
        let Some(body) = self.text.as_mut() else {
            return Ok(());
        };
        body.builder
            .start(element, namespace, resolver, decoder, span, empty)
            .map_err(text_error)
    }

    /// Hand one end tag inside the open text body to the shared reader.
    fn forward_end(&mut self, end: usize) -> Result<()> {
        let Some(body) = self.text.as_mut() else {
            return Ok(());
        };
        body.builder.end(end).map_err(text_error)
    }

    /// Hand one event of markup this parser captures as opaque to the shared
    /// reader when it lies inside the open text body. This preserves semantic
    /// field text while retaining the original event in the opaque payload.
    fn forward_markup(
        &mut self,
        event: &Event<'_>,
        namespace: &ResolveResult<'_>,
        resolver: &NamespaceResolver,
        decoder: Decoder,
        span: Range<usize>,
    ) -> Result<()> {
        match event {
            Event::Start(element) => {
                self.forward_start(element, namespace, resolver, decoder, span, false)
            },
            Event::Empty(element) => {
                self.forward_start(element, namespace, resolver, decoder, span, true)
            },
            Event::End(_) => self.forward_end(span.end),
            Event::Text(text) => {
                let decoded = text
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|_| encoding_error("invalid SpreadsheetML drawing text"))?;
                self.forward_text(&decoded)
            },
            Event::CData(text) => {
                let decoded = text
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|_| encoding_error("invalid SpreadsheetML drawing text"))?;
                self.forward_text(&decoded)
            },
            Event::GeneralRef(reference) => {
                let decoded = decode_xml_reference(reference)
                    .map_err(|_| malformed_xml("invalid SpreadsheetML drawing reference"))?;
                self.forward_text(&decoded)
            },
            _ => Ok(()),
        }
    }

    fn forward_text(&mut self, text: &str) -> Result<()> {
        let Some(body) = self.text.as_mut() else {
            return Ok(());
        };
        body.builder.text(text).map_err(text_error)
    }

    /// Finish the shared text-body reader when `xdr:txBody` closes.
    fn close_text(&mut self) -> Result<()> {
        let Some(body) = self.text.take() else {
            return Ok(());
        };
        let body = body.builder.finish().map_err(text_error)?;
        let text_body = &mut self.builder_mut()?.text_body;
        text_body.properties = body.properties;
        text_body.paragraphs = body.paragraphs;
        Ok(())
    }

    fn capture_event(&mut self, event: &Event<'_>) -> Result<bool> {
        let capture = self
            .unknown
            .as_mut()
            .ok_or_else(|| invalid("missing unknown drawing capture"))?;
        match event {
            Event::Start(_) => {
                capture.depth = capture.depth.checked_add(1).ok_or_else(|| {
                    limit(LimitKind::XmlDepth, usize::MAX, self.limits.max_xml_depth)
                })?;
                if capture.depth > self.limits.max_xml_depth {
                    return Err(limit(
                        LimitKind::XmlDepth,
                        capture.depth,
                        self.limits.max_xml_depth,
                    ));
                }
            },
            Event::End(_) => {
                capture.depth = capture
                    .depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("unknown drawing element depth underflow"))?;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            _ => {},
        }
        capture
            .writer
            .write_event(event.clone())
            .map_err(|_| malformed_xml("could not retain SpreadsheetML drawing markup"))?;
        if capture.writer.get_ref().len() > self.limits.max_drawing_part_bytes {
            return Err(limit(
                LimitKind::DrawingBytes,
                capture.writer.get_ref().len(),
                self.limits.max_drawing_part_bytes,
            ));
        }
        Ok(capture.depth == 0)
    }

    fn anchor_mut(&mut self) -> Result<&mut PendingAnchor> {
        self.anchor
            .as_mut()
            .ok_or_else(|| invalid("drawing object outside an anchor"))
    }

    fn builder_mut(&mut self) -> Result<&mut ObjectBuilder> {
        self.builders
            .last_mut()
            .ok_or_else(|| invalid("drawing shape content outside a shape"))
    }

    fn open_object(&mut self, kind: BuilderKind) -> Result<Context> {
        if self.anchor.is_none() {
            return Err(invalid("drawing object outside an anchor"));
        }
        if self.builders.len() >= self.limits.max_group_depth {
            return Err(limit(
                LimitKind::GroupDepth,
                self.builders.len().saturating_add(1),
                self.limits.max_group_depth,
            ));
        }
        self.builders.push(ObjectBuilder::new(kind));
        Ok(Context::Object)
    }

    fn close_object(&mut self) -> Result<()> {
        let builder = self
            .builders
            .pop()
            .ok_or_else(|| invalid("mismatched drawing object close"))?;
        let Some(object) = builder.finish() else {
            return Ok(());
        };
        self.attach_object(object)
    }

    #[allow(
        clippy::too_many_lines,
        reason = "The parser grammar is intentionally kept in one exhaustive dispatch function."
    )]
    fn start(
        &mut self,
        parent: Context,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
        decoder: Decoder,
        resolver: &NamespaceResolver,
        empty: bool,
    ) -> Result<Context> {
        let name = element.name();
        let local = name.local_name();
        let local = local.as_ref();
        let xdr = is_xdr(namespace);
        match parent {
            Context::Root if xdr => match local {
                b"twoCellAnchor" => {
                    self.open_anchor(AnchorKind::TwoCell)?;
                    let edit_as = unqualified_attribute_value(element, b"editAs", decoder)?;
                    self.anchor_mut()?.edit_as =
                        edit_as.as_deref().map_or(Ok(EditAs::TwoCell), |value| {
                            parse_value(value, "drawing editAs")
                        })?;
                    return Ok(Context::Anchor);
                },
                b"oneCellAnchor" => {
                    self.open_anchor(AnchorKind::OneCell)?;
                    return Ok(Context::Anchor);
                },
                b"absoluteAnchor" => {
                    self.open_anchor(AnchorKind::Absolute)?;
                    return Ok(Context::Anchor);
                },
                _ => {},
            },
            Context::Anchor if xdr => match local {
                b"from" => {
                    let anchor = self.anchor_mut()?;
                    if anchor.from.replace(Marker::default()).is_some() {
                        return Err(invalid("drawing anchor has duplicate from markers"));
                    }
                    return Ok(Context::From);
                },
                b"to" => {
                    let anchor = self.anchor_mut()?;
                    if anchor.to.replace(Marker::default()).is_some() {
                        return Err(invalid("drawing anchor has duplicate to markers"));
                    }
                    return Ok(Context::To);
                },
                b"pos" => {
                    let position = EmuOffset {
                        x: Emu(emu_attribute(element, b"x", decoder)?),
                        y: Emu(emu_attribute(element, b"y", decoder)?),
                    };
                    if self.anchor_mut()?.position.replace(position).is_some() {
                        return Err(invalid("drawing anchor has duplicate positions"));
                    }
                    return Ok(Context::Other);
                },
                b"ext" => {
                    let extent = EmuExtent {
                        width: Emu(emu_attribute(element, b"cx", decoder)?),
                        height: Emu(emu_attribute(element, b"cy", decoder)?),
                    };
                    if self.anchor_mut()?.extent.replace(extent).is_some() {
                        return Err(invalid("drawing anchor has duplicate extents"));
                    }
                    return Ok(Context::Other);
                },
                b"sp" => return self.open_object(BuilderKind::Shape),
                b"grpSp" => return self.open_object(BuilderKind::Group),
                b"cxnSp" => return self.open_object(BuilderKind::Connection),
                b"graphicFrame" => return self.open_object(BuilderKind::GraphicFrame),
                // Pictures have a dedicated image owner and remain outside
                // this shape inventory.
                b"pic" => return Ok(Context::Other),
                b"clientData" => {
                    let client_data = ClientData {
                        locks_with_sheet: bool_attribute(element, b"fLocksWithSheet", decoder)?,
                        prints_with_sheet: bool_attribute(element, b"fPrintsWithSheet", decoder)?,
                    };
                    self.anchor_mut()?.client_data = client_data;
                    return Ok(Context::Other);
                },
                _ => {},
            },
            Context::From | Context::To if xdr => {
                let target = if parent == Context::From {
                    MarkerTarget::From
                } else {
                    MarkerTarget::To
                };
                for (field_name, field) in [
                    (b"col".as_slice(), MarkerField::Column),
                    (b"colOff".as_slice(), MarkerField::ColumnOffset),
                    (b"row".as_slice(), MarkerField::Row),
                    (b"rowOff".as_slice(), MarkerField::RowOffset),
                ] {
                    if local == field_name {
                        self.marker_text.clear();
                        return Ok(Context::Marker(target, field));
                    }
                }
                return Ok(Context::Other);
            },
            Context::CustomGeometry(geometry_parent)
                if is_drawingml_name(namespace, name, local) =>
            {
                return self.open_geometry_child(geometry_parent, local, element, decoder);
            },
            // Objects nest only inside groups.
            Context::Object
                if xdr
                    && self
                        .builders
                        .last()
                        .is_some_and(|b| b.kind == BuilderKind::Group) =>
            {
                match local {
                    b"sp" => return self.open_object(BuilderKind::Shape),
                    b"grpSp" => return self.open_object(BuilderKind::Group),
                    b"cxnSp" => return self.open_object(BuilderKind::Connection),
                    b"graphicFrame" => return self.open_object(BuilderKind::GraphicFrame),
                    _ => {},
                }
            },
            _ => {},
        }
        if xdr {
            match local {
                b"cNvPr" if !self.builders.is_empty() => {
                    let builder = self.builder_mut()?;
                    builder.non_visual.id = unqualified_attribute_value(element, b"id", decoder)?
                        .and_then(|value| value.parse().ok());
                    builder.non_visual.name =
                        unqualified_attribute_value(element, b"name", decoder)?;
                    builder.non_visual.description =
                        unqualified_attribute_value(element, b"descr", decoder)?;
                    builder.non_visual.hidden =
                        bool_attribute(element, b"hidden", decoder)?.unwrap_or(false);
                    self.retain_unknown_attributes(
                        element,
                        decoder,
                        &[b"id", b"name", b"descr", b"hidden"],
                    )?;
                },
                b"cNvSpPr" if !self.builders.is_empty() => {
                    self.builder_mut()?.is_text_box =
                        bool_attribute(element, b"txBox", decoder)?.unwrap_or(false);
                    self.retain_unknown_attributes(element, decoder, &[b"txBox"])?;
                },
                b"txBody" if !self.builders.is_empty() => {
                    let builder = self.builder_mut()?;
                    if builder.text_body.in_text_body {
                        return Err(invalid("drawing shape contains duplicate text bodies"));
                    }
                    builder.text_body.in_text_body = true;
                    return Ok(Context::Body);
                },
                b"oleObject"
                    if self
                        .builders
                        .last()
                        .is_some_and(|b| b.kind == BuilderKind::GraphicFrame) =>
                {
                    let builder = self.builder_mut()?;
                    if builder.ole_object.is_some() {
                        return Err(invalid("graphic frame contains duplicate OLE objects"));
                    }
                    builder.ole_object = Some(OleObject {
                        program_id: unqualified_attribute_value(element, b"progId", decoder)?,
                        shape_id: unqualified_attribute_value(element, b"shapeId", decoder)?
                            .and_then(|value| value.parse().ok()),
                        data_or_view_aspect: unqualified_attribute_value(
                            element,
                            b"dvAspect",
                            decoder,
                        )?
                        .as_deref()
                        .map(Aspect::try_from)
                        .transpose()?,
                        auto_load: bool_attribute(element, b"autoLoad", decoder)?,
                        relationship_id: relationship_attribute_value(
                            element, b"id", decoder, resolver,
                        )?,
                        link_relationship_id: relationship_attribute_value(
                            element, b"link", decoder, resolver,
                        )?,
                        ..OleObject::default()
                    });
                },
                _ => {},
            }
        }
        if is_drawingml_name(namespace, name, local) && !self.builders.is_empty() {
            let builder_kind = self.builders.last().map(|b| b.kind);
            match local {
                b"prstGeom" => {
                    let preset = unqualified_attribute_value(element, b"prst", decoder)?
                        .ok_or_else(|| invalid("DrawingML prstGeom is missing required prst"))?;
                    let token = xsd_token_atom(&preset)
                        .ok_or_else(|| invalid("invalid DrawingML shape preset"))?;
                    let parsed = token
                        .parse()
                        .map_err(|_| invalid("invalid DrawingML shape preset"))?;
                    let builder = self.builder_mut()?;
                    if builder.geometry.is_some() || builder.geometry_builder.is_some() {
                        return Err(invalid(
                            "drawing shape contains competing or duplicate geometries",
                        ));
                    }
                    builder.geometry = Some(Geometry::Preset(parsed));
                },
                b"custGeom"
                    if matches!(
                        builder_kind,
                        Some(BuilderKind::Shape | BuilderKind::Connection)
                    ) =>
                {
                    return self.open_custom_geometry();
                },
                b"spLocks" | b"cxnSpLocks" | b"grpSpLocks"
                    if any_truthy_attribute(element, decoder)? =>
                {
                    self.builder_mut()?.non_visual.locked = true;
                },
                b"stCxn" | b"endCxn" if builder_kind == Some(BuilderKind::Connection) => {
                    let end = ConnectionEnd {
                        shape_id: required_u32_attribute(element, b"id", decoder, "connection ID")?,
                        site: required_u32_attribute(element, b"idx", decoder, "connection site")?,
                    };
                    let builder = self.builder_mut()?;
                    let slot = if local == b"stCxn" {
                        &mut builder.start
                    } else {
                        &mut builder.end
                    };
                    if slot.replace(end).is_some() {
                        return Err(invalid("connection shape has duplicate connection ends"));
                    }
                },
                b"off" | b"ext" | b"chOff" | b"chExt"
                    if builder_kind == Some(BuilderKind::Group) =>
                {
                    self.apply_group_transform(local, element, decoder)?;
                },
                b"bodyPr" => {
                    return Ok(Context::Properties);
                },
                _local if parent == Context::Properties => {
                    // The shared reader owns all body-property decoding.
                    // Keep these elements in the host stack without parsing
                    // them a second time here.
                },
                // The shared text-body reader builds paragraphs and runs
                // from the forwarded events; these contexts only steer which
                // markup this parser also retains as opaque.
                b"p" if parent == Context::Body => return Ok(Context::Paragraph),
                b"r" if parent == Context::Paragraph => return Ok(Context::Run),
                b"rPr" if parent == Context::Run => return Ok(Context::RunProperties),
                b"t" if parent == Context::Run => return Ok(Context::Text),
                _ => {},
            }
        }
        if let Some(target) = self.unknown_target(parent, namespace, local) {
            return self.open_unknown(element, target, empty);
        }
        Ok(Context::Other)
    }

    fn unknown_target(
        &self,
        parent: Context,
        namespace: &ResolveResult<'_>,
        local: &[u8],
    ) -> Option<UnknownTarget> {
        if parent == Context::Anchor && is_xdr(namespace) {
            return Some(UnknownTarget::AnchorObject);
        }
        if parent == Context::Object && !is_known_object_element(namespace, local) {
            return Some(
                if is_xdr(namespace)
                    && self
                        .builders
                        .last()
                        .is_some_and(|builder| builder.kind == BuilderKind::Group)
                {
                    UnknownTarget::GroupChild
                } else {
                    UnknownTarget::ObjectMarkup
                },
            );
        }
        if matches!(
            parent,
            Context::Body | Context::Properties | Context::Paragraph | Context::Run
        ) && is_drawingml_name(namespace, QName(local), local)
        {
            return Some(UnknownTarget::ObjectMarkup);
        }
        if parent == Context::Other
            && !self.builders.is_empty()
            && !is_known_object_element(namespace, local)
            && is_drawingml_name(namespace, QName(local), local)
        {
            return Some(UnknownTarget::ObjectMarkup);
        }
        None
    }

    fn open_unknown(
        &mut self,
        element: &BytesStart<'_>,
        target: UnknownTarget,
        empty: bool,
    ) -> Result<Context> {
        if self.unknown.is_some() {
            return Err(invalid("nested unknown drawing captures"));
        }
        let mut writer = Writer::new(Vec::new());
        let event = if empty {
            Event::Empty(element.to_owned())
        } else {
            Event::Start(element.to_owned())
        };
        writer
            .write_event(event)
            .map_err(|_| malformed_xml("could not retain SpreadsheetML drawing markup"))?;
        if writer.get_ref().len() > self.limits.max_drawing_part_bytes {
            return Err(limit(
                LimitKind::DrawingBytes,
                writer.get_ref().len(),
                self.limits.max_drawing_part_bytes,
            ));
        }
        self.unknown = Some(UnknownCapture {
            depth: usize::from(!empty),
            target,
            writer,
        });
        Ok(Context::Unknown)
    }

    fn finish_unknown(&mut self) -> Result<()> {
        let capture = self
            .unknown
            .take()
            .ok_or_else(|| invalid("missing unknown drawing capture"))?;
        if capture.depth != 0 {
            return Err(invalid("unknown drawing element is not closed"));
        }
        let element = UnknownElement::from_xml(capture.writer.into_inner())?;
        match capture.target {
            UnknownTarget::AnchorObject => {
                self.attach_object(Object::Unknown(Unknown::from_element(element)))?;
            },
            UnknownTarget::GroupChild => {
                self.attach_object(Object::Unknown(Unknown::from_element(element)))?;
            },
            UnknownTarget::ObjectMarkup => {
                self.builder_mut()?
                    .non_visual
                    .opaque
                    .push_element(element)?;
            },
        }
        Ok(())
    }

    fn retain_unknown_attributes(
        &mut self,
        element: &BytesStart<'_>,
        decoder: Decoder,
        known: &[&[u8]],
    ) -> Result<()> {
        let mut unknown = Vec::new();
        for attribute in element.checked_attributes() {
            let attribute = attribute
                .map_err(|_| malformed_xml("malformed SpreadsheetML drawing attribute"))?;
            let key = attribute.key.as_ref();
            if key == b"xmlns" || key.starts_with(b"xmlns:") {
                continue;
            }
            if known
                .iter()
                .any(|name| attribute.key.local_name().as_ref() == *name)
            {
                continue;
            }
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(|_| malformed_xml("malformed SpreadsheetML drawing attribute"))?;
            unknown.push(UnknownAttribute::new(
                String::from_utf8_lossy(key).into_owned(),
                value.into_owned(),
            )?);
        }
        let builder = self.builder_mut()?;
        for attribute in unknown {
            builder.non_visual.opaque.push_attribute(attribute)?;
        }
        Ok(())
    }

    fn attach_object(&mut self, object: Object) -> Result<()> {
        self.object_count = self
            .object_count
            .checked_add(1)
            .ok_or_else(|| limit(LimitKind::Objects, usize::MAX, self.limits.max_objects))?;
        if self.object_count > self.limits.max_objects {
            return Err(limit(
                LimitKind::Objects,
                self.object_count,
                self.limits.max_objects,
            ));
        }
        if let Some(parent) = self.builders.last_mut() {
            parent.children.push(object);
            return Ok(());
        }
        let anchor = self.anchor_mut()?;
        if anchor.object.replace(object).is_some() {
            return Err(invalid("drawing anchor contains multiple objects"));
        }
        Ok(())
    }

    /// Open the `a:custGeom` element of the current shape.
    fn open_custom_geometry(&mut self) -> Result<Context> {
        let geometry_limits = self.limits.geometry;
        let builder = self.builder_mut()?;
        if builder.geometry.is_some() || builder.geometry_builder.is_some() {
            return Err(invalid(
                "drawing shape contains competing or duplicate geometries",
            ));
        }
        builder.geometry_builder = Some(CustomGeometryBuilder::new(geometry_limits)?);
        Ok(Context::CustomGeometry(GeometryElement::CustomGeometry))
    }

    /// Route one `DrawingML` child of the custom geometry subtree into the
    /// geometry builder; unknown children are skipped inertly.
    fn open_geometry_child(
        &mut self,
        parent: GeometryElement,
        local: &[u8],
        element: &BytesStart<'_>,
        decoder: Decoder,
    ) -> Result<Context> {
        let builder = self.builder_mut()?;
        let Some(geometry) = builder.geometry_builder.as_mut() else {
            return Ok(Context::Other);
        };
        Ok(
            match geometry
                .open(parent, local, element, decoder)
                .map_err(geometry_error)?
            {
                Some(child) => Context::CustomGeometry(child),
                None => Context::Other,
            },
        )
    }

    /// Close one custom geometry element; the `a:custGeom` close finalizes
    /// the builder into the shape.
    fn finish_geometry(&mut self, element: GeometryElement) -> Result<()> {
        let builder = self.builder_mut()?;
        if element == GeometryElement::CustomGeometry {
            if let Some(geometry) = builder.geometry_builder.take() {
                builder.geometry = Some(geometry.finish().map_err(geometry_error)?.into());
            }
            return Ok(());
        }
        if let Some(geometry) = builder.geometry_builder.as_mut() {
            geometry.close(element).map_err(geometry_error)?;
        }
        Ok(())
    }

    fn apply_group_transform(
        &mut self,
        local: &[u8],
        element: &BytesStart<'_>,
        decoder: Decoder,
    ) -> Result<()> {
        let builder = self.builder_mut()?;
        builder.saw_transform = true;
        match local {
            b"off" => {
                builder.transform.offset = Some(EmuOffset {
                    x: Emu(emu_attribute(element, b"x", decoder)?),
                    y: Emu(emu_attribute(element, b"y", decoder)?),
                });
            },
            b"ext" => {
                builder.transform.extent = Some(EmuExtent {
                    width: Emu(emu_attribute(element, b"cx", decoder)?),
                    height: Emu(emu_attribute(element, b"cy", decoder)?),
                });
            },
            b"chOff" => {
                builder.transform.child_offset = Some(EmuOffset {
                    x: Emu(emu_attribute(element, b"x", decoder)?),
                    y: Emu(emu_attribute(element, b"y", decoder)?),
                });
            },
            b"chExt" => {
                builder.transform.child_extent = Some(EmuExtent {
                    width: Emu(emu_attribute(element, b"cx", decoder)?),
                    height: Emu(emu_attribute(element, b"cy", decoder)?),
                });
            },
            _ => {},
        }
        Ok(())
    }

    fn finish(&mut self, context: Context) -> Result<()> {
        match context {
            Context::Marker(target, field) => self.finish_marker(target, field),
            Context::Anchor => self.finish_anchor(),
            Context::Object => self.close_object(),
            Context::Unknown => self.finish_unknown(),
            Context::CustomGeometry(element) => self.finish_geometry(element),
            Context::Body => self.close_text(),
            _ => Ok(()),
        }
    }

    fn finish_marker(&mut self, target: MarkerTarget, field: MarkerField) -> Result<()> {
        let value = self.marker_text.trim();
        let anchor = self
            .anchor
            .as_mut()
            .ok_or_else(|| invalid("drawing marker outside an anchor"))?;
        let marker = match target {
            MarkerTarget::From => anchor.from.as_mut(),
            MarkerTarget::To => anchor.to.as_mut(),
        }
        .ok_or_else(|| invalid("drawing marker value outside from/to"))?;
        match field {
            MarkerField::Column => set_once(
                &mut marker.column,
                parse_value(value, "drawing column")?,
                "drawing column",
            ),
            MarkerField::ColumnOffset => set_once(
                &mut marker.column_offset,
                parse_value(value, "drawing column offset")?,
                "drawing column offset",
            ),
            MarkerField::Row => set_once(
                &mut marker.row,
                parse_value(value, "drawing row")?,
                "drawing row",
            ),
            MarkerField::RowOffset => set_once(
                &mut marker.row_offset,
                parse_value(value, "drawing row offset")?,
                "drawing row offset",
            ),
        }
    }

    fn finish_anchor(&mut self) -> Result<()> {
        let pending = self
            .anchor
            .take()
            .ok_or_else(|| invalid("missing pending drawing anchor"))?;
        // Anchors carrying only pictures, charts, or unknown objects are not
        // part of this inventory.
        let Some(object) = pending.object else {
            return Ok(());
        };
        let anchor = match pending.kind.ok_or_else(|| invalid("missing anchor kind"))? {
            AnchorKind::TwoCell => {
                let from = pending
                    .from
                    .ok_or_else(|| invalid("drawing anchor is missing its from marker"))?
                    .finish("drawing from marker")?;
                let to = pending
                    .to
                    .ok_or_else(|| invalid("drawing anchor is missing its to marker"))?
                    .finish("drawing to marker")?;
                check_marker_bounds(from)?;
                check_marker_bounds(to)?;
                Anchor::TwoCell {
                    from,
                    to,
                    edit_as: pending.edit_as,
                }
            },
            AnchorKind::OneCell => {
                let from = pending
                    .from
                    .ok_or_else(|| invalid("drawing anchor is missing its from marker"))?
                    .finish("drawing from marker")?;
                check_marker_bounds(from)?;
                let extent = pending
                    .extent
                    .ok_or_else(|| invalid("one-cell anchor is missing its extent"))?;
                Anchor::OneCell { from, extent }
            },
            AnchorKind::Absolute => {
                let position = pending
                    .position
                    .ok_or_else(|| invalid("absolute anchor is missing its position"))?;
                let extent = pending
                    .extent
                    .ok_or_else(|| invalid("absolute anchor is missing its extent"))?;
                Anchor::Absolute { position, extent }
            },
        };
        let next = self.objects.len().saturating_add(1);
        if next > self.limits.max_anchors {
            return Err(limit(LimitKind::Anchors, next, self.limits.max_anchors));
        }
        self.objects.push(AnchoredObject {
            anchor,
            object,
            client_data: pending.client_data,
        });
        Ok(())
    }

    fn open_anchor(&mut self, kind: AnchorKind) -> Result<()> {
        if self.anchor.is_some() {
            return Err(invalid("nested drawing anchor"));
        }
        self.anchor = Some(PendingAnchor {
            kind: Some(kind),
            ..PendingAnchor::default()
        });
        Ok(())
    }
}

fn check_marker_bounds(marker: CellMarker) -> Result<()> {
    if marker.column >= 16_384 || marker.row >= 1_048_576 {
        return Err(invalid("drawing anchor exceeds worksheet bounds"));
    }
    Ok(())
}

fn is_xdr(namespace: &ResolveResult<'_>) -> bool {
    matches!(
        namespace,
        ResolveResult::Bound(Namespace(value))
            if *value == SPREADSHEET_DRAWING_NAMESPACE
                || *value == STRICT_SPREADSHEET_DRAWING_NAMESPACE
    )
}

fn is_xdr_name(namespace: &ResolveResult<'_>, name: QName<'_>, local_name: &[u8]) -> bool {
    name.local_name().as_ref() == local_name && is_xdr(namespace)
}

fn is_known_object_element(namespace: &ResolveResult<'_>, local: &[u8]) -> bool {
    if is_xdr(namespace) {
        return matches!(
            local,
            b"nvSpPr"
                | b"nvCxnSpPr"
                | b"nvGrpSpPr"
                | b"nvGraphicFramePr"
                | b"cNvPr"
                | b"cNvSpPr"
                | b"cNvCxnSpPr"
                | b"cNvGrpSpPr"
                | b"cNvGraphicFramePr"
                | b"spPr"
                | b"grpSpPr"
                | b"txBody"
                | b"oleObject"
                | b"sp"
                | b"grpSp"
                | b"cxnSp"
                | b"graphicFrame"
                | b"graphic"
                | b"graphicData"
                | b"chart"
                | b"clientData"
        );
    }
    if matches!(
        local,
        b"spPr"
            | b"prstGeom"
            | b"custGeom"
            | b"avLst"
            | b"pathLst"
            | b"bodyPr"
            | b"lstStyle"
            | b"p"
            | b"pPr"
            | b"defRPr"
            | b"r"
            | b"rPr"
            | b"t"
            | b"br"
            | b"noAutofit"
            | b"spAutoFit"
            | b"normAutofit"
            | b"spLocks"
            | b"cxnSpLocks"
            | b"grpSpLocks"
            | b"stCxn"
            | b"endCxn"
            | b"graphic"
            | b"graphicData"
            | b"xfrm"
            | b"off"
            | b"ext"
            | b"chOff"
            | b"chExt"
    ) {
        return true;
    }
    false
}

fn parse_dml_bool(value: &str, attribute: &str) -> Result<bool> {
    parse_bool(value).map_err(|_| invalid(format!("invalid DrawingML {attribute} boolean")))
}

fn emu_attribute(element: &BytesStart<'_>, name: &[u8], decoder: Decoder) -> Result<i64> {
    let value = unqualified_attribute_value(element, name, decoder)?.ok_or_else(|| {
        invalid(format!(
            "drawing coordinate is missing '{}'",
            String::from_utf8_lossy(name)
        ))
    })?;
    value
        .parse()
        .map_err(|_| invalid("invalid drawing coordinate value"))
}

fn required_u32_attribute(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> Result<u32> {
    let value = unqualified_attribute_value(element, name, decoder)?
        .ok_or_else(|| invalid(format!("{description} attribute is missing")))?;
    value
        .parse()
        .map_err(|_| invalid(format!("invalid {description}")))
}

fn bool_attribute(element: &BytesStart<'_>, name: &[u8], decoder: Decoder) -> Result<Option<bool>> {
    unqualified_attribute_value(element, name, decoder)?
        .map(|value| parse_dml_bool(&value, &String::from_utf8_lossy(name)))
        .transpose()
}

fn any_truthy_attribute(element: &BytesStart<'_>, decoder: Decoder) -> Result<bool> {
    let mut any = false;
    for attribute in element.checked_attributes() {
        let attribute =
            attribute.map_err(|_| malformed_xml("malformed SpreadsheetML drawing attribute"))?;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|_| malformed_xml("malformed SpreadsheetML drawing attribute"))?;
        any |= parse_dml_bool(&value, "lock")?;
    }
    Ok(any)
}

/// The shared text-body reader's limits inside a drawing. The host's drawing
/// limits and the shared semantic text limits are deliberately finite and
/// use the same source budgets.
fn text_limits(limits: Limits) -> Result<text_reader::Limits> {
    text_reader::Limits::new(
        limits.max_drawing_part_bytes,
        limits.max_xml_depth.saturating_mul(2),
        limits.max_text_elements,
        limits.max_text_bytes,
    )
    .ok_or_else(|| invalid("invalid SpreadsheetDrawing text limits"))
}

/// Preserve the shared reader's content-free typed error source.
fn text_error(error: litchi_drawingml::Error) -> Error {
    Error::Drawing(error)
}

fn geometry_error(error: litchi_drawingml::Error) -> Error {
    use litchi_drawingml::Error as DrawingError;
    Error::Drawing(match error {
        DrawingError::Invalid(_) => DrawingError::Invalid("invalid custom geometry value".into()),
        DrawingError::Xml(_) => DrawingError::Xml("malformed custom geometry XML".into()),
        DrawingError::Common(XmlError::Invalid(_)) => {
            DrawingError::Common(XmlError::Invalid("invalid custom geometry XML".into()))
        },
        DrawingError::Common(XmlError::Malformed(_)) => {
            DrawingError::Common(XmlError::Malformed("malformed custom geometry XML".into()))
        },
        other => other,
    })
}

fn position(reader: &NsReader<&[u8]>, origin: ReaderOrigin) -> Result<usize> {
    origin
        .offset(reader.buffer_position())
        .ok_or_else(|| invalid("drawing XML position does not fit usize"))
}

fn set_once<T>(target: &mut Option<T>, value: T, description: &str) -> Result<()> {
    if target.replace(value).is_some() {
        return Err(invalid(format!("duplicate {description}")));
    }
    Ok(())
}

fn parse_value<T: FromStr>(value: &str, description: &str) -> Result<T> {
    value
        .parse()
        .map_err(|_| invalid(format!("invalid {description}")))
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn malformed_xml(message: &'static str) -> Error {
    Error::Xml(XmlError::Malformed(message.to_owned()))
}

fn encoding_error(message: &'static str) -> Error {
    Error::Encoding(message.to_owned())
}

fn limit(kind: LimitKind, observed: usize, maximum: usize) -> Error {
    Error::limit(kind, observed, maximum)
}
