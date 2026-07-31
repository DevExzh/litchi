//! SpreadsheetDrawing XML inventory for XLSB Drawings parts.
//!
//! Per MS-XLSB 2.1.7.23 the Drawings part of an XLSB package is the standard
//! SpreadsheetML drawing XML part (ISO/IEC 29500-1:2016 sections 12.3.8 and
//! 20.5), identical to XLSX — only sheet streams are binary. This module
//! parses that XML into an inert inventory of anchored objects (shapes,
//! pictures, graphic frames, connection shapes, and group shapes) with their
//! cell anchors. The standalone parser stores relationship identifiers and
//! content URIs verbatim. Workbook loading resolves internal image and Chart
//! parts into bounded inert payloads and the typed chart model shared with
//! the other formats ([`crate::charts`]), including chart external data,
//! user-shapes, and extension resources; external targets are never fetched.

use crate::common::xml::unqualified_attribute_value;
use crate::error::{OoxmlError, Result};
use crate::xlsb::error::XlsbResult;
use quick_xml::Decoder;
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

/// SpreadsheetDrawing namespace (transitional and strict).
const XDR: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const STRICT_XDR: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing";

/// OfficeDocument relationships namespace (transitional and strict), used
/// for `r:id` / `r:embed` / `r:link` attributes.
const REL: &[u8] = b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_REL: &[u8] = b"http://purl.oclc.org/ooxml/officeDocument/relationships";

/// `graphicData` URI of a DrawingML chart (`c:chart`), ISO/IEC 29500-1:2016
/// section 21.2.
pub const CHART_GRAPHIC_DATA_URI: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";

/// Maximum accepted size of one Drawings part.
const MAX_DRAWING_XML_BYTES: usize = 16 * 1024 * 1024;
/// Maximum XML element nesting depth.
const MAX_XML_DEPTH: usize = 256;
/// Maximum number of anchored objects in one Drawings part.
const MAX_ANCHORED_OBJECTS: usize = 100_000;

/// Cell anchor marker (`xdr:from` / `xdr:to`): a zero-based row/column plus
/// an EMU offset into the cell (ISO/IEC 29500-1:2016 section 20.5.2.14).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsbDrawingCellMarker {
    /// Zero-based column index (`xdr:col`).
    pub column: u32,
    /// EMU offset into the column (`xdr:colOff`).
    pub column_offset: i64,
    /// Zero-based row index (`xdr:row`).
    pub row: u32,
    /// EMU offset into the row (`xdr:rowOff`).
    pub row_offset: i64,
}

/// EMU position (`xdr:pos`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsbDrawingEmuPoint {
    /// X coordinate in EMUs.
    pub x: i64,
    /// Y coordinate in EMUs.
    pub y: i64,
}

/// EMU extent (`xdr:ext`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsbDrawingEmuSize {
    /// Width in EMUs (`cx`).
    pub width: i64,
    /// Height in EMUs (`cy`).
    pub height: i64,
}

/// How an object is anchored to the sheet grid.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum XlsbDrawingAnchorKind {
    /// Anchored between two cells, moving and sizing per `edit_as`
    /// (`xdr:twoCellAnchor`).
    TwoCell {
        /// Start marker.
        from: XlsbDrawingCellMarker,
        /// End marker.
        to: XlsbDrawingCellMarker,
        /// Resize behavior token (`editAs`: `twoCell`, `oneCell`, or
        /// `absolute`); `twoCell` when the attribute is absent.
        edit_as: Option<String>,
    },
    /// Anchored at one cell with an explicit extent (`xdr:oneCellAnchor`).
    OneCell {
        /// Start marker.
        from: XlsbDrawingCellMarker,
        /// Object extent.
        extent: XlsbDrawingEmuSize,
    },
    /// Anchored at an absolute position (`xdr:absoluteAnchor`).
    Absolute {
        /// Object position.
        position: XlsbDrawingEmuPoint,
        /// Object extent.
        extent: XlsbDrawingEmuSize,
    },
}

/// Non-visual identification shared by all drawing objects (`xdr:cNvPr`).
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct XlsbDrawingNonVisual {
    /// Drawing object identifier (`id`); absent IDs surface as 0.
    pub id: u32,
    /// Drawing object name (`name`).
    pub name: String,
    /// Optional alternative text (`descr`).
    pub description: Option<String>,
}

/// A graphic frame (`xdr:graphicFrame`) and the content it hosts.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsbDrawingGraphicFrame {
    /// Non-visual identification.
    pub non_visual: XlsbDrawingNonVisual,
    /// `graphicData` content URI, for example [`CHART_GRAPHIC_DATA_URI`].
    pub content_uri: String,
    /// Relationship identifier carried by the hosted content (for example
    /// `c:chart r:id`); stored verbatim and never dereferenced.
    pub rel_id: Option<String>,
}

impl XlsbDrawingGraphicFrame {
    /// Whether the hosted content is a DrawingML chart.
    pub fn is_chart(&self) -> bool {
        self.content_uri == CHART_GRAPHIC_DATA_URI
    }
}

/// One anchored drawing object.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum XlsbDrawingObject {
    /// A shape (`xdr:sp`).
    Shape(XlsbDrawingNonVisual),
    /// A picture (`xdr:pic`).
    Picture {
        /// Non-visual identification.
        non_visual: XlsbDrawingNonVisual,
        /// Relationship identifier of the image (`a:blip r:embed`), when
        /// declared; stored verbatim and never dereferenced.
        embed_rel_id: Option<String>,
    },
    /// A graphic frame hosting foreign content such as a chart
    /// (`xdr:graphicFrame`).
    GraphicFrame(XlsbDrawingGraphicFrame),
    /// A connection shape (`xdr:cxnSp`).
    ConnectionShape(XlsbDrawingNonVisual),
    /// A group shape (`xdr:grpSp`); nested objects are not inventoried.
    GroupShape(XlsbDrawingNonVisual),
}

/// One anchor and the object it anchors, in drawing order.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsbDrawingAnchor {
    /// How the object is anchored to the grid.
    pub anchor: XlsbDrawingAnchorKind,
    /// The anchored object.
    pub object: XlsbDrawingObject,
}

/// Inert inventory of one Drawings part.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct XlsbDrawing {
    /// Anchored objects in drawing order.
    pub anchors: Vec<XlsbDrawingAnchor>,
}

/// The drawing of one sheet plus the images and charts its objects reference.
#[derive(Debug, Clone)]
pub struct XlsbSheetDrawing {
    /// Zero-based sheet index in workbook sheet order.
    pub sheet_index: usize,
    /// Anchored-object inventory of the Drawings part.
    pub drawing: XlsbDrawing,
    /// Typed charts resolved from chart graphic frames.
    pub charts: Vec<XlsbEmbeddedChart>,
    /// Embedded image parts resolved from picture objects.
    pub images: Vec<XlsbEmbeddedImage>,
    /// Detailed standard DrawingML shapes, groups, and connectors.
    pub shapes: Vec<crate::xlsx::XlsxAnchoredObject>,
}

/// One embedded image part resolved through a drawing picture.
#[derive(Debug, Clone)]
pub struct XlsbEmbeddedImage {
    /// Name of the hosting picture object (`xdr:cNvPr name`).
    pub picture_name: String,
    /// Optional picture alternative text (`xdr:cNvPr descr`).
    pub description: Option<String>,
    /// Relationship identifier from the drawing part to the image part.
    pub rel_id: String,
    /// Typed encoded image format.
    pub format: crate::xlsb::drawing_image::XlsbWorksheetImageFormat,
    /// Exact encoded image bytes, shared when pictures reuse one Image part.
    pub data: std::sync::Arc<[u8]>,
}

/// One embedded chart part resolved through a drawing graphic frame.
#[derive(Debug, Clone)]
pub struct XlsbEmbeddedChart {
    /// Name of the hosting graphic frame (`xdr:cNvPr name`).
    pub frame_name: String,
    /// Relationship identifier from the drawing part to the chart part;
    /// stored verbatim and never dereferenced.
    pub rel_id: String,
    /// Typed chart parsed from the standard DrawingML `c:chartSpace` part
    /// (MS-XLSB 2.1.7.5 defers to ISO/IEC 29500-1:2016 section 21.2).
    pub chart: crate::charts::Chart,
    /// Embedded or linked external-data payload declared by the chart.
    pub external_data_part: Option<crate::xlsx::ChartExternalDataPart>,
    /// Chart user-shapes XML and its directly related inert resources.
    pub user_shapes_part: Option<crate::xlsx::ChartUserShapesPart>,
    /// Other relationships owned by the Chart part, including resources
    /// referenced by preserved extension fragments.
    pub additional_relationships: Vec<crate::xlsx::ChartRelationship>,
}

fn invalid(message: impl Into<String>) -> OoxmlError {
    OoxmlError::InvalidFormat(message.into())
}

fn limit(what: &str) -> OoxmlError {
    OoxmlError::InvalidFormat(format!("drawing XML exceeds {what} limit"))
}

/// Parse one SpreadsheetDrawing part into its anchored-object inventory.
///
/// Markup-compatibility processing is applied so `mc:AlternateContent`
/// fallbacks resolve before parsing. The root element must be `xdr:wsDr`.
pub fn parse_drawing_part(xml_bytes: &[u8]) -> XlsbResult<XlsbDrawing> {
    Ok(parse_drawing_part_xml(xml_bytes)?)
}

fn parse_drawing_part_xml(xml_bytes: &[u8]) -> Result<XlsbDrawing> {
    if xml_bytes.len() > MAX_DRAWING_XML_BYTES {
        return Err(limit("drawing part bytes"));
    }
    let xml = std::str::from_utf8(xml_bytes)
        .map_err(|error| invalid(format!("drawing part is not UTF-8: {error}")))?;
    let xml = litchi_ooxml_common::mce::process_str(xml)?;
    Parser::parse(xml.as_ref())
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Context {
    Root,
    Anchor,
    From,
    To,
    Marker(MarkerField),
    Object,
    Other,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum MarkerField {
    Column,
    ColumnOffset,
    Row,
    RowOffset,
}

impl MarkerField {
    fn description(self) -> &'static str {
        match self {
            MarkerField::Column => "marker column",
            MarkerField::ColumnOffset => "marker column offset",
            MarkerField::Row => "marker row",
            MarkerField::RowOffset => "marker row offset",
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum AnchorKind {
    TwoCell,
    OneCell,
    Absolute,
}

#[derive(Default)]
struct PendingMarker {
    column: Option<u32>,
    column_offset: Option<i64>,
    row: Option<u32>,
    row_offset: Option<i64>,
}

impl PendingMarker {
    fn assign(&mut self, field: MarkerField, text: &str) -> Result<()> {
        let description = field.description();
        let trimmed = text.trim();
        match field {
            MarkerField::Column | MarkerField::Row => {
                let value: u32 = trimmed
                    .parse()
                    .map_err(|_| invalid(format!("invalid {description} value '{trimmed}'")))?;
                let target = match field {
                    MarkerField::Column => &mut self.column,
                    _ => &mut self.row,
                };
                if target.replace(value).is_some() {
                    return Err(invalid(format!("{description} is duplicated")));
                }
            },
            MarkerField::ColumnOffset | MarkerField::RowOffset => {
                let value: i64 = trimmed
                    .parse()
                    .map_err(|_| invalid(format!("invalid {description} value '{trimmed}'")))?;
                let target = match field {
                    MarkerField::ColumnOffset => &mut self.column_offset,
                    _ => &mut self.row_offset,
                };
                if target.replace(value).is_some() {
                    return Err(invalid(format!("{description} is duplicated")));
                }
            },
        }
        Ok(())
    }

    fn finish(self, description: &str) -> Result<XlsbDrawingCellMarker> {
        Ok(XlsbDrawingCellMarker {
            column: self
                .column
                .ok_or_else(|| invalid(format!("{description} is missing its column")))?,
            column_offset: self
                .column_offset
                .ok_or_else(|| invalid(format!("{description} is missing its column offset")))?,
            row: self
                .row
                .ok_or_else(|| invalid(format!("{description} is missing its row")))?,
            row_offset: self
                .row_offset
                .ok_or_else(|| invalid(format!("{description} is missing its row offset")))?,
        })
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum ObjectKind {
    Shape,
    Picture,
    GraphicFrame,
    ConnectionShape,
    GroupShape,
}

struct PendingObject {
    kind: ObjectKind,
    non_visual: XlsbDrawingNonVisual,
    saw_non_visual: bool,
    embed_rel_id: Option<String>,
    content_uri: Option<String>,
    rel_id: Option<String>,
}

impl PendingObject {
    fn new(kind: ObjectKind) -> Self {
        PendingObject {
            kind,
            non_visual: XlsbDrawingNonVisual::default(),
            saw_non_visual: false,
            embed_rel_id: None,
            content_uri: None,
            rel_id: None,
        }
    }

    /// Inspect one element inside the object subtree.
    fn observe(
        &mut self,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
        decoder: Decoder,
        resolver: &NamespaceResolver,
    ) -> Result<()> {
        let local = element.name().local_name();
        let local = local.as_ref();
        if is_xdr(namespace) && local == b"cNvPr" && !self.saw_non_visual {
            self.saw_non_visual = true;
            self.non_visual.id = match unqualified_attribute_value(element, b"id", decoder)? {
                Some(value) => value
                    .parse()
                    .map_err(|_| invalid(format!("invalid drawing object id '{value}'")))?,
                None => 0,
            };
            self.non_visual.name =
                unqualified_attribute_value(element, b"name", decoder)?.unwrap_or_default();
            self.non_visual.description = unqualified_attribute_value(element, b"descr", decoder)?;
            return Ok(());
        }
        if self.kind == ObjectKind::Picture
            && self.embed_rel_id.is_none()
            && crate::common::xml::is_drawingml_name(namespace, element.name(), b"blip")
        {
            self.embed_rel_id = relationship_attribute(element, decoder, resolver)?;
            return Ok(());
        }
        if self.kind == ObjectKind::GraphicFrame
            && crate::common::xml::is_drawingml_name(namespace, element.name(), b"graphicData")
        {
            if self.content_uri.is_some() {
                return Err(invalid("graphic frame has multiple graphicData elements"));
            }
            self.content_uri =
                Some(unqualified_attribute_value(element, b"uri", decoder)?.unwrap_or_default());
            return Ok(());
        }
        if self.kind == ObjectKind::GraphicFrame && self.rel_id.is_none() {
            self.rel_id = relationship_attribute(element, decoder, resolver)?;
        }
        Ok(())
    }

    fn finish(self) -> XlsbDrawingObject {
        match self.kind {
            ObjectKind::Shape => XlsbDrawingObject::Shape(self.non_visual),
            ObjectKind::Picture => XlsbDrawingObject::Picture {
                non_visual: self.non_visual,
                embed_rel_id: self.embed_rel_id,
            },
            ObjectKind::GraphicFrame => XlsbDrawingObject::GraphicFrame(XlsbDrawingGraphicFrame {
                non_visual: self.non_visual,
                content_uri: self.content_uri.unwrap_or_default(),
                rel_id: self.rel_id,
            }),
            ObjectKind::ConnectionShape => XlsbDrawingObject::ConnectionShape(self.non_visual),
            ObjectKind::GroupShape => XlsbDrawingObject::GroupShape(self.non_visual),
        }
    }
}

struct PendingAnchor {
    kind: AnchorKind,
    edit_as: Option<String>,
    from: Option<PendingMarker>,
    to: Option<PendingMarker>,
    position: Option<XlsbDrawingEmuPoint>,
    extent: Option<XlsbDrawingEmuSize>,
    object: Option<XlsbDrawingObject>,
}

impl PendingAnchor {
    fn finish(self) -> Result<XlsbDrawingAnchor> {
        let anchor = match self.kind {
            AnchorKind::TwoCell => XlsbDrawingAnchorKind::TwoCell {
                from: self
                    .from
                    .ok_or_else(|| invalid("twoCellAnchor is missing its from marker"))?
                    .finish("twoCellAnchor from marker")?,
                to: self
                    .to
                    .ok_or_else(|| invalid("twoCellAnchor is missing its to marker"))?
                    .finish("twoCellAnchor to marker")?,
                edit_as: self.edit_as,
            },
            AnchorKind::OneCell => XlsbDrawingAnchorKind::OneCell {
                from: self
                    .from
                    .ok_or_else(|| invalid("oneCellAnchor is missing its from marker"))?
                    .finish("oneCellAnchor from marker")?,
                extent: self
                    .extent
                    .ok_or_else(|| invalid("oneCellAnchor is missing its extent"))?,
            },
            AnchorKind::Absolute => XlsbDrawingAnchorKind::Absolute {
                position: self
                    .position
                    .ok_or_else(|| invalid("absoluteAnchor is missing its position"))?,
                extent: self
                    .extent
                    .ok_or_else(|| invalid("absoluteAnchor is missing its extent"))?,
            },
        };
        let object = self
            .object
            .ok_or_else(|| invalid("drawing anchor has no object"))?;
        Ok(XlsbDrawingAnchor { anchor, object })
    }
}

struct Parser {
    anchors: Vec<XlsbDrawingAnchor>,
    anchor: Option<PendingAnchor>,
    object: Option<PendingObject>,
    marker_text: String,
}

impl Parser {
    fn parse(xml: &str) -> Result<XlsbDrawing> {
        let mut reader = NsReader::from_reader(xml.as_bytes());
        reader.config_mut().trim_text(false);
        let mut parser = Parser {
            anchors: Vec::new(),
            anchor: None,
            object: None,
            marker_text: String::new(),
        };
        let mut stack: Vec<Context> = Vec::new();
        let mut closed_root = false;
        loop {
            let decoder = reader.decoder();
            let event = reader
                .read_event()
                .map_err(|error| OoxmlError::Xml(error.to_string()))?
                .into_owned();
            let resolver = reader.resolver().clone();
            let (namespace, event) = resolver.resolve_event(event);
            match event {
                Event::Start(element) if stack.is_empty() => {
                    if closed_root {
                        return Err(invalid("drawing XML contains multiple root elements"));
                    }
                    if !is_xdr_name(&namespace, element.name().local_name().as_ref(), b"wsDr") {
                        return Err(invalid("drawing part root is not xdr:wsDr"));
                    }
                    stack.push(Context::Root);
                },
                Event::Empty(element) if stack.is_empty() => {
                    if !is_xdr_name(&namespace, element.name().local_name().as_ref(), b"wsDr") {
                        return Err(invalid("drawing part root is not xdr:wsDr"));
                    }
                    return Ok(XlsbDrawing {
                        anchors: parser.anchors,
                    });
                },
                Event::Start(element) => {
                    let parent = *stack
                        .last()
                        .ok_or_else(|| invalid("missing drawing root"))?;
                    let context = parser.start(parent, &namespace, &element, decoder, &resolver)?;
                    stack.push(context);
                    if stack.len() > MAX_XML_DEPTH {
                        return Err(limit("drawing XML depth"));
                    }
                },
                Event::Empty(element) => {
                    let parent = *stack
                        .last()
                        .ok_or_else(|| invalid("missing drawing root"))?;
                    let context = parser.start(parent, &namespace, &element, decoder, &resolver)?;
                    parser.finish(context, parent)?;
                },
                Event::Text(text) => {
                    if matches!(stack.last(), Some(Context::Marker(_))) {
                        let decoded = text
                            .xml_content(XmlVersion::Explicit1_0)
                            .map_err(|error| OoxmlError::Xml(error.to_string()))?;
                        parser.marker_text.push_str(&decoded);
                    }
                },
                Event::End(_) => {
                    let context = stack
                        .pop()
                        .ok_or_else(|| invalid("drawing XML depth underflow"))?;
                    if context == Context::Root {
                        closed_root = true;
                    } else {
                        let parent = stack.last().copied().unwrap_or(Context::Root);
                        parser.finish(context, parent)?;
                    }
                },
                Event::Eof => {
                    if !stack.is_empty() {
                        return Err(invalid("drawing XML ends inside the root element"));
                    }
                    if !closed_root {
                        return Err(invalid("drawing part is missing xdr:wsDr"));
                    }
                    return Ok(XlsbDrawing {
                        anchors: parser.anchors,
                    });
                },
                _ => {},
            }
        }
    }

    fn anchor_mut(&mut self) -> Result<&mut PendingAnchor> {
        self.anchor
            .as_mut()
            .ok_or_else(|| invalid("drawing anchor state is missing"))
    }

    fn start(
        &mut self,
        parent: Context,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
        decoder: Decoder,
        resolver: &NamespaceResolver,
    ) -> Result<Context> {
        // While an object is open, every descendant belongs to its subtree:
        // observe it (non-visual properties, blip embeds, graphicData
        // content) and swallow the rest of the subtree as `Other`.
        if let Some(object) = self.object.as_mut() {
            object.observe(namespace, element, decoder, resolver)?;
            return Ok(Context::Other);
        }
        let local = element.name().local_name();
        let local = local.as_ref();
        let xdr = is_xdr(namespace);
        match parent {
            Context::Root if xdr => {
                let kind = match local {
                    b"twoCellAnchor" => AnchorKind::TwoCell,
                    b"oneCellAnchor" => AnchorKind::OneCell,
                    b"absoluteAnchor" => AnchorKind::Absolute,
                    _ => return Ok(Context::Other),
                };
                if self.anchor.is_some() {
                    return Err(invalid("nested drawing anchors"));
                }
                if self.anchors.len() >= MAX_ANCHORED_OBJECTS {
                    return Err(limit("anchored object count"));
                }
                let edit_as = if kind == AnchorKind::TwoCell {
                    unqualified_attribute_value(element, b"editAs", decoder)?
                } else {
                    None
                };
                self.anchor = Some(PendingAnchor {
                    kind,
                    edit_as,
                    from: None,
                    to: None,
                    position: None,
                    extent: None,
                    object: None,
                });
                Ok(Context::Anchor)
            },
            Context::Anchor if xdr => match local {
                b"from" => {
                    let anchor = self.anchor_mut()?;
                    if anchor.from.is_some() {
                        return Err(invalid("drawing anchor has duplicate from markers"));
                    }
                    anchor.from = Some(PendingMarker::default());
                    Ok(Context::From)
                },
                b"to" => {
                    let anchor = self.anchor_mut()?;
                    if anchor.to.is_some() {
                        return Err(invalid("drawing anchor has duplicate to markers"));
                    }
                    anchor.to = Some(PendingMarker::default());
                    Ok(Context::To)
                },
                b"pos" => {
                    let anchor = self.anchor_mut()?;
                    if anchor.position.is_some() {
                        return Err(invalid("drawing anchor has duplicate positions"));
                    }
                    anchor.position = Some(XlsbDrawingEmuPoint {
                        x: emu_attribute(element, b"x", decoder)?,
                        y: emu_attribute(element, b"y", decoder)?,
                    });
                    Ok(Context::Other)
                },
                b"ext" => {
                    let anchor = self.anchor_mut()?;
                    if anchor.extent.is_some() {
                        return Err(invalid("drawing anchor has duplicate extents"));
                    }
                    anchor.extent = Some(XlsbDrawingEmuSize {
                        width: emu_attribute(element, b"cx", decoder)?,
                        height: emu_attribute(element, b"cy", decoder)?,
                    });
                    Ok(Context::Other)
                },
                b"sp" | b"pic" | b"graphicFrame" | b"cxnSp" | b"grpSp" => {
                    let kind = match local {
                        b"sp" => ObjectKind::Shape,
                        b"pic" => ObjectKind::Picture,
                        b"graphicFrame" => ObjectKind::GraphicFrame,
                        b"cxnSp" => ObjectKind::ConnectionShape,
                        _ => ObjectKind::GroupShape,
                    };
                    if self.object.is_some() {
                        return Err(invalid("nested drawing objects"));
                    }
                    if self.anchor_mut()?.object.is_some() {
                        return Err(invalid("drawing anchor has multiple objects"));
                    }
                    self.object = Some(PendingObject::new(kind));
                    Ok(Context::Object)
                },
                _ => Ok(Context::Other),
            },
            Context::From | Context::To if xdr => {
                let field = match local {
                    b"col" => MarkerField::Column,
                    b"colOff" => MarkerField::ColumnOffset,
                    b"row" => MarkerField::Row,
                    b"rowOff" => MarkerField::RowOffset,
                    _ => return Ok(Context::Other),
                };
                self.marker_text.clear();
                Ok(Context::Marker(field))
            },
            _ => Ok(Context::Other),
        }
    }

    fn finish(&mut self, context: Context, parent: Context) -> Result<()> {
        match context {
            Context::Anchor => {
                let anchor = self
                    .anchor
                    .take()
                    .ok_or_else(|| invalid("drawing anchor state is missing"))?;
                self.anchors.push(anchor.finish()?);
            },
            Context::Object => {
                let object = self
                    .object
                    .take()
                    .ok_or_else(|| invalid("drawing object state is missing"))?;
                self.anchor_mut()?.object = Some(object.finish());
            },
            Context::Marker(field) => {
                let text = std::mem::take(&mut self.marker_text);
                let anchor = self.anchor_mut()?;
                let marker = match parent {
                    Context::From => anchor.from.as_mut(),
                    Context::To => anchor.to.as_mut(),
                    _ => None,
                }
                .ok_or_else(|| invalid("drawing marker state is missing"))?;
                marker.assign(field, &text)?;
            },
            _ => {},
        }
        Ok(())
    }
}

fn is_xdr(namespace: &ResolveResult<'_>) -> bool {
    matches!(
        namespace,
        ResolveResult::Bound(Namespace(value)) if *value == XDR || *value == STRICT_XDR
    )
}

fn is_xdr_name(namespace: &ResolveResult<'_>, local: &[u8], expected: &[u8]) -> bool {
    local == expected && is_xdr(namespace)
}

/// Read the first relationship-namespaced attribute (`r:id`, `r:embed`,
/// `r:link`) of an element.
fn relationship_attribute(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<Option<String>> {
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| OoxmlError::Xml(error.to_string()))?;
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        let is_rel = matches!(
            namespace,
            ResolveResult::Bound(Namespace(value)) if value == REL || value == STRICT_REL
        );
        if is_rel
            && matches!(
                attribute.key.local_name().as_ref(),
                b"id" | b"embed" | b"link"
            )
        {
            return Ok(Some(
                attribute
                    .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                    .map_err(|error| OoxmlError::Xml(error.to_string()))?
                    .into_owned(),
            ));
        }
    }
    Ok(None)
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

#[cfg(test)]
mod tests;
