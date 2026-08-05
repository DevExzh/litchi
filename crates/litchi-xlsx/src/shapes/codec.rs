//! Bounded SpreadsheetDrawingML XML codec.
//!
//! Parsing is inert and namespace-aware. Every source is subject to byte,
//! depth, text, object, group, and anchor budgets; unmodeled fragments are
//! retained as bounded opaque values.

use std::str::FromStr;

use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, QName, ResolveResult};
use quick_xml::reader::NsReader;
use quick_xml::writer::Writer;

use crate::error::{Error, Result};
use crate::raw::namespace::relationship_attribute_value;
use litchi_drawingml::geometry::reader::{CustomGeometryBuilder, GeometryElement};
use litchi_drawingml::text::parse_bool;
use litchi_ooxml_common::xml::{
    decode_xml_reference, is_drawingml_name, unqualified_attribute_value, xsd_token_atom,
};

use super::model::*;

const SPREADSHEET_DRAWING_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const STRICT_SPREADSHEET_DRAWING_NAMESPACE: &[u8] =
    b"http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing";

/// Parse one SpreadsheetDrawing part into its anchored shape inventory.
///
/// Returns `Ok(None)` when the part has no `xdr:wsDr` root. Markup-
/// compatibility processing is applied so `mc:AlternateContent` fallbacks
/// resolve before parsing. Pictures and chart graphic frames are skipped;
/// structurally invalid anchors are errors.
pub fn parse_drawing_shapes(xml: &str) -> Result<Option<Vec<AnchoredObject>>> {
    if xml.len() > MAX_DRAWING_PART_BYTES {
        return Err(limit("drawing part bytes"));
    }
    let xml = litchi_ooxml_common::mce::process_str(xml)?;
    Parser::parse(xml.as_ref())
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
    paragraphs: Vec<Paragraph>,
    paragraph: Option<Paragraph>,
    run: Option<Run>,
    in_text_body: bool,
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

struct Parser {
    objects: Vec<AnchoredObject>,
    object_count: usize,
    text_bytes: usize,
    anchor: Option<PendingAnchor>,
    builders: Vec<ObjectBuilder>,
    marker_text: String,
    unknown: Option<UnknownCapture>,
}

impl Parser {
    fn parse(xml: &str) -> Result<Option<Vec<AnchoredObject>>> {
        let mut reader = NsReader::from_reader(xml.as_bytes());
        reader.config_mut().trim_text(false);
        let mut parser = Self {
            objects: Vec::new(),
            object_count: 0,
            text_bytes: 0,
            anchor: None,
            builders: Vec::new(),
            marker_text: String::new(),
            unknown: None,
        };
        let mut stack = Vec::new();
        let mut closed_root = false;
        loop {
            let decoder = reader.decoder();
            let event = reader
                .read_event()
                .map_err(|error| Error::Invalid(error.to_string()))?
                .into_owned();
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
                    if stack.len() > MAX_XML_DEPTH {
                        return Err(limit("drawing XML depth"));
                    }
                },
                Event::Empty(element) => {
                    let parent = *stack
                        .last()
                        .ok_or_else(|| invalid("missing drawing root"))?;
                    let context =
                        parser.start(parent, &namespace, &element, decoder, &resolver, true)?;
                    parser.finish(context)?;
                },
                Event::Text(text) => {
                    let decoded = text
                        .xml_content(XmlVersion::Explicit1_0)
                        .map_err(|error| Error::Invalid(error.to_string()))?;
                    match stack.last() {
                        Some(Context::Marker(_, _)) => parser.marker_text.push_str(&decoded),
                        Some(Context::Text) => parser.push_run_text(&decoded)?,
                        _ => {},
                    }
                },
                Event::CData(text) => {
                    let decoded = text
                        .xml_content(XmlVersion::Explicit1_0)
                        .map_err(|error| Error::Invalid(error.to_string()))?;
                    match stack.last() {
                        Some(Context::Marker(_, _)) => parser.marker_text.push_str(&decoded),
                        Some(Context::Text) => parser.push_run_text(&decoded)?,
                        _ => {},
                    }
                },
                Event::GeneralRef(reference) => {
                    let decoded = decode_xml_reference(&reference)?;
                    match stack.last() {
                        Some(Context::Marker(_, _)) => parser.marker_text.push_str(&decoded),
                        Some(Context::Text) => parser.push_run_text(&decoded)?,
                        _ => {},
                    }
                },
                Event::End(element) => {
                    let context = stack
                        .pop()
                        .ok_or_else(|| invalid("drawing XML closes outside its root"))?;
                    parser.finish(context)?;
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

    fn push_run_text(&mut self, text: &str) -> Result<()> {
        self.text_bytes = self
            .text_bytes
            .checked_add(text.len())
            .ok_or_else(|| limit("shape text bytes"))?;
        if self.text_bytes > MAX_TEXT_BYTES {
            return Err(limit("shape text bytes"));
        }
        if let Some(builder) = self.builders.last_mut()
            && let Some(run) = builder.text_body.run.as_mut()
        {
            run.text.push_str(text);
        }
        Ok(())
    }

    fn capture_event(&mut self, event: &Event<'_>) -> Result<bool> {
        let capture = self
            .unknown
            .as_mut()
            .ok_or_else(|| invalid("missing unknown drawing capture"))?;
        match event {
            Event::Start(_) => {
                capture.depth = capture
                    .depth
                    .checked_add(1)
                    .ok_or_else(|| limit("unknown drawing element depth"))?;
                if capture.depth > MAX_XML_DEPTH {
                    return Err(limit("unknown drawing element depth"));
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
            .map_err(|error| Error::Invalid(error.to_string()))?;
        if capture.writer.get_ref().len() > MAX_DRAWING_PART_BYTES {
            return Err(limit("unknown drawing element bytes"));
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
        if self.builders.len() >= MAX_GROUP_DEPTH {
            return Err(limit("shape group depth"));
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

    #[allow(clippy::too_many_lines)]
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
                    let token = xsd_token_atom(&preset).ok_or_else(|| {
                        invalid(format!("invalid DrawingML shape preset '{preset}'"))
                    })?;
                    let parsed = token.parse().map_err(|error| {
                        invalid(format!(
                            "invalid DrawingML shape preset '{preset}': {error}"
                        ))
                    })?;
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
                    self.parse_properties(element, decoder)?;
                    return Ok(Context::Properties);
                },
                b"noAutofit" if parent == Context::Properties => {
                    self.builder_mut()?.text_body.properties.autofit = Autofit::None;
                },
                b"spAutoFit" if parent == Context::Properties => {
                    self.builder_mut()?.text_body.properties.autofit = Autofit::Shape;
                },
                b"normAutofit" if parent == Context::Properties => {
                    self.builder_mut()?.text_body.properties.autofit = Autofit::Normal;
                },
                b"p" if parent == Context::Body => {
                    let builder = self.builder_mut()?;
                    if builder.text_body.paragraph.is_some() {
                        return Err(invalid("nested drawing text paragraphs"));
                    }
                    builder.text_body.paragraph = Some(Paragraph::default());
                    return Ok(Context::Paragraph);
                },
                b"r" if parent == Context::Paragraph => {
                    let builder = self.builder_mut()?;
                    if builder.text_body.run.is_some() {
                        return Err(invalid("nested drawing text runs"));
                    }
                    builder.text_body.run = Some(Run::default());
                    return Ok(Context::Run);
                },
                b"rPr" if parent == Context::Run => {
                    self.parse_run_properties(element, decoder)?;
                    return Ok(Context::RunProperties);
                },
                b"t" if parent == Context::Run => {
                    return Ok(Context::Text);
                },
                b"br" if parent == Context::Paragraph => {
                    // A DrawingML break contributes a newline to the paragraph.
                    let builder = self.builder_mut()?;
                    if let Some(paragraph) = builder.text_body.paragraph.as_mut() {
                        paragraph.runs.push(Run {
                            text: "\n".to_string(),
                            ..Run::default()
                        });
                    }
                },
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
            .map_err(|error| Error::Invalid(error.to_string()))?;
        if writer.get_ref().len() > MAX_DRAWING_PART_BYTES {
            return Err(limit("unknown drawing element bytes"));
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
        for attribute in element.attributes() {
            let attribute = attribute.map_err(|error| Error::Invalid(error.to_string()))?;
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
                .map_err(|error| Error::Invalid(error.to_string()))?;
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
            .ok_or_else(|| limit("objects per drawing"))?;
        if self.object_count > MAX_OBJECTS_PER_DRAWING {
            return Err(limit("objects per drawing"));
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
        let builder = self.builder_mut()?;
        if builder.geometry.is_some() || builder.geometry_builder.is_some() {
            return Err(invalid(
                "drawing shape contains competing or duplicate geometries",
            ));
        }
        builder.geometry_builder = Some(CustomGeometryBuilder::new());
        Ok(Context::CustomGeometry(GeometryElement::CustomGeometry))
    }

    /// Route one DrawingML child of the custom geometry subtree into the
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
        Ok(match geometry.open(parent, local, element, decoder)? {
            Some(child) => Context::CustomGeometry(child),
            None => Context::Other,
        })
    }

    /// Close one custom geometry element; the `a:custGeom` close finalizes
    /// the builder into the shape.
    fn finish_geometry(&mut self, element: GeometryElement) -> Result<()> {
        let builder = self.builder_mut()?;
        if element == GeometryElement::CustomGeometry {
            if let Some(geometry) = builder.geometry_builder.take() {
                builder.geometry = Some(geometry.finish()?.into());
            }
            return Ok(());
        }
        if let Some(geometry) = builder.geometry_builder.as_mut() {
            geometry.close(element)?;
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

    fn parse_properties(&mut self, element: &BytesStart<'_>, decoder: Decoder) -> Result<()> {
        let body = &mut self.builder_mut()?.text_body.properties;
        if let Some(value) = unqualified_attribute_value(element, b"lIns", decoder)? {
            body.insets.left = parse_value(&value, "left text inset")?;
        }
        if let Some(value) = unqualified_attribute_value(element, b"tIns", decoder)? {
            body.insets.top = parse_value(&value, "top text inset")?;
        }
        if let Some(value) = unqualified_attribute_value(element, b"rIns", decoder)? {
            body.insets.right = parse_value(&value, "right text inset")?;
        }
        if let Some(value) = unqualified_attribute_value(element, b"bIns", decoder)? {
            body.insets.bottom = parse_value(&value, "bottom text inset")?;
        }
        if let Some(value) = unqualified_attribute_value(element, b"anchor", decoder)? {
            body.vertical_anchor = parse_value(&value, "text anchor")?;
        }
        if let Some(value) = unqualified_attribute_value(element, b"anchorCtr", decoder)? {
            body.anchor_center = parse_dml_bool(&value, "anchorCtr")?;
        }
        if let Some(value) = unqualified_attribute_value(element, b"vert", decoder)? {
            body.direction = parse_value(&value, "text direction")?;
        }
        if let Some(value) = unqualified_attribute_value(element, b"wrap", decoder)? {
            body.wrap = parse_value(&value, "text wrap")?;
        }
        if let Some(value) = unqualified_attribute_value(element, b"numCol", decoder)? {
            body.column_count = parse_value(&value, "text column count")?;
        }
        if let Some(value) = unqualified_attribute_value(element, b"spcFirstLastPara", decoder)? {
            body.space_first_last_paragraph = parse_dml_bool(&value, "spcFirstLastPara")?;
        }
        Ok(())
    }

    fn parse_run_properties(&mut self, element: &BytesStart<'_>, decoder: Decoder) -> Result<()> {
        let builder = self.builder_mut()?;
        let Some(run) = builder.text_body.run.as_mut() else {
            return Ok(());
        };
        if let Some(value) = unqualified_attribute_value(element, b"b", decoder)? {
            run.bold = Some(parse_dml_bool(&value, "run bold")?);
        }
        if let Some(value) = unqualified_attribute_value(element, b"i", decoder)? {
            run.italic = Some(parse_dml_bool(&value, "run italic")?);
        }
        if let Some(value) = unqualified_attribute_value(element, b"u", decoder)? {
            run.underline = Some(Underline::from_dml(&value).map_err(|error| {
                invalid(format!("invalid DrawingML underline '{value}': {error}"))
            })?);
        }
        if let Some(value) = unqualified_attribute_value(element, b"sz", decoder)? {
            run.font_size = Some(parse_value(&value, "text size")?);
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
            Context::Run => {
                if let Some(builder) = self.builders.last_mut()
                    && let (Some(run), Some(paragraph)) = (
                        builder.text_body.run.take(),
                        builder.text_body.paragraph.as_mut(),
                    )
                {
                    paragraph.runs.push(run);
                }
                Ok(())
            },
            Context::Paragraph => {
                if let Some(builder) = self.builders.last_mut()
                    && let Some(paragraph) = builder.text_body.paragraph.take()
                {
                    builder.text_body.paragraphs.push(paragraph);
                }
                Ok(())
            },
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
        if self.objects.len() >= MAX_ANCHORS_PER_DRAWING {
            return Err(limit("anchors per drawing"));
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
    parse_bool(value).map_err(|error| {
        invalid(format!(
            "invalid DrawingML {attribute} boolean '{value}': {error}"
        ))
    })
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
        .map_err(|_| invalid(format!("invalid drawing coordinate value '{value}'")))
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
        .map_err(|_| invalid(format!("invalid {description} '{value}'")))
}

fn bool_attribute(element: &BytesStart<'_>, name: &[u8], decoder: Decoder) -> Result<Option<bool>> {
    unqualified_attribute_value(element, name, decoder)?
        .map(|value| parse_dml_bool(&value, &String::from_utf8_lossy(name)))
        .transpose()
}

fn any_truthy_attribute(element: &BytesStart<'_>, decoder: Decoder) -> Result<bool> {
    let mut any = false;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Invalid(error.to_string()))?;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| Error::Invalid(error.to_string()))?;
        any |= parse_dml_bool(&value, "lock")?;
    }
    Ok(any)
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
        .map_err(|_| invalid(format!("invalid {description} '{value}'")))
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn limit(name: &str) -> Error {
    invalid(format!("XLSX drawing shape {name} limit exceeded"))
}
