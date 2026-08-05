//! Typed, inert SpreadsheetDrawing shape and text-box inventory for XLSX
//! worksheets.
//!
//! A worksheet drawing part (`/xl/drawings/drawingN.xml`) anchors objects to
//! the sheet grid with `xdr:twoCellAnchor`, `xdr:oneCellAnchor`, or
//! `xdr:absoluteAnchor` elements (ECMA-376 part 1, SpreadsheetDrawingML).
//! Beyond pictures and charts — already covered by the image and chart
//! support — anchors may hold arbitrary DrawingML shapes (`xdr:sp` with an
//! `a:prstGeom` preset geometry and an `xdr:txBody` rich-text story),
//! connection shapes (`xdr:cxnSp`), shape groups (`xdr:grpSp`), and legacy
//! OLE objects (`xdr:oleObject` inside a `xdr:graphicFrame`).
//!
//! [`parse_drawing_shapes`] parses one drawing part into the typed model;
//! [`load_shapes`] and [`load_sheet_shapes`] resolve drawing parts
//! through the package relationship graph. Everything here is read-only and
//! inert: unknown elements are skipped, OLE payloads and external targets are
//! never followed, and all inputs are bounded by named limits.
//!
//! Specification anchors: [MS-XLSX] SpreadsheetDrawing complex types and
//! schema (structure 2.6, schema appendix 5.8), with shared shape content from
//! [MS-ODRAWXML]. The package and relationship traversal follows
//! [MS-OI29500] OPC boundaries and remains format-owned here.

use std::fmt;
use std::str::FromStr;

use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, QName, ResolveResult};
use quick_xml::reader::NsReader;

use crate::error::{Error, Result};
use crate::raw::Sheet;
use crate::raw::namespace::relationship_attribute_value;
use crate::raw::parse_catalog;
use litchi_drawingml::geom::Preset;
use litchi_drawingml::geometry::CustomGeometry;
use litchi_drawingml::geometry::reader::{CustomGeometryBuilder, GeometryElement};
pub use litchi_drawingml::text::body::{Body, Insets, Paragraph, Properties, Run};
use litchi_drawingml::text::parse_bool;
pub use litchi_drawingml::text::{
    Anchor as VerticalAnchor, Autofit, Columns, Coordinate32, Direction, TextSize, Underline, Wrap,
};

/// How a legacy OLE object is rendered inside its graphic frame.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Aspect {
    /// Render the embedded object's content.
    Content,
    /// Render the embedded object's icon.
    Icon,
}

impl Aspect {
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Content => "DVASPECT_CONTENT",
            Self::Icon => "DVASPECT_ICON",
        }
    }
}

impl FromStr for Aspect {
    type Err = Error;

    fn from_str(value: &str) -> Result<Self> {
        match value {
            "DVASPECT_CONTENT" => Ok(Self::Content),
            "DVASPECT_ICON" => Ok(Self::Icon),
            _ => Err(invalid(format!("invalid OLE data/view aspect '{value}'"))),
        }
    }
}

impl TryFrom<&str> for Aspect {
    type Error = Error;

    fn try_from(value: &str) -> Result<Self> {
        value.parse()
    }
}

impl fmt::Display for Aspect {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.as_str())
    }
}
use litchi_ooxml_common::xml::{
    decode_xml_reference, is_drawingml_name, unqualified_attribute_value, xsd_token_atom,
};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{OpcPackage, Part};

const SPREADSHEET_DRAWING_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const STRICT_SPREADSHEET_DRAWING_NAMESPACE: &[u8] =
    b"http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing";

const MAX_DRAWING_PART_BYTES: usize = 32 * 1024 * 1024;
const MAX_WORKBOOK_BYTES: usize = 32 * 1024 * 1024;
const MAX_DRAWINGS_PER_WORKSHEET: usize = 64;
const MAX_ANCHORS_PER_DRAWING: usize = 100_000;
const MAX_OBJECTS_PER_DRAWING: usize = 100_000;
const MAX_GROUP_DEPTH: usize = 32;
const MAX_XML_DEPTH: usize = 256;
const MAX_TEXT_BYTES: usize = 1024 * 1024;

/// An offset or extent in English Metric Units (EMU), the DrawingML length unit.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct Emu(pub i64);

impl Emu {
    /// The raw EMU value.
    pub fn emu(self) -> i64 {
        self.0
    }
}

impl From<i64> for Emu {
    fn from(value: i64) -> Self {
        Self(value)
    }
}

/// How a two-cell anchored object reacts to cell edits (`xdr:twoCellAnchor@editAs`).
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash)]
pub enum EditAs {
    /// Move and resize with both cells (`twoCell`, the ECMA-376 default).
    #[default]
    TwoCell,
    /// Move with the anchor cell but keep the size (`oneCell`).
    OneCell,
    /// Do not move or resize with cells (`absolute`).
    Absolute,
}

impl EditAs {
    /// Return the exact SpreadsheetDrawingML token.
    pub const fn token(self) -> &'static str {
        match self {
            Self::TwoCell => "twoCell",
            Self::OneCell => "oneCell",
            Self::Absolute => "absolute",
        }
    }
}

/// An invalid `xdr:twoCellAnchor@editAs` token.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct EditAsError;

impl fmt::Display for EditAsError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("invalid SpreadsheetDrawingML editAs token")
    }
}

impl std::error::Error for EditAsError {}

impl FromStr for EditAs {
    type Err = EditAsError;

    fn from_str(value: &str) -> std::result::Result<Self, Self::Err> {
        match value {
            "twoCell" => Ok(Self::TwoCell),
            "oneCell" => Ok(Self::OneCell),
            "absolute" => Ok(Self::Absolute),
            _ => Err(EditAsError),
        }
    }
}

impl fmt::Display for EditAs {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.token())
    }
}

/// One cell anchor point: a zero-based column/row plus an EMU offset into the cell.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct CellMarker {
    /// Zero-based column index.
    pub column: u32,
    /// Offset from the column edge, in EMUs.
    pub column_offset: Emu,
    /// Zero-based row index.
    pub row: u32,
    /// Offset from the row edge, in EMUs.
    pub row_offset: Emu,
}

/// An absolute position (`xdr:pos`), in EMUs.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct EmuOffset {
    /// Horizontal offset, in EMUs.
    pub x: Emu,
    /// Vertical offset, in EMUs.
    pub y: Emu,
}

/// An object extent (`xdr:ext` or `a:ext`), in EMUs.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct EmuExtent {
    /// Width, in EMUs.
    pub width: Emu,
    /// Height, in EMUs.
    pub height: Emu,
}

/// How an object is anchored on the worksheet grid.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Anchor {
    /// `xdr:twoCellAnchor` — bounded by a from and a to cell marker.
    TwoCell {
        /// Top-left anchor point.
        from: CellMarker,
        /// Bottom-right anchor point.
        to: CellMarker,
        /// Edit behavior recorded by `editAs`.
        edit_as: EditAs,
    },
    /// `xdr:oneCellAnchor` — anchored at one cell with an explicit extent.
    OneCell {
        /// Top-left anchor point.
        from: CellMarker,
        /// Object size.
        extent: EmuExtent,
    },
    /// `xdr:absoluteAnchor` — fixed position and size, independent of cells.
    Absolute {
        /// Top-left position.
        position: EmuOffset,
        /// Object size.
        extent: EmuExtent,
    },
}

/// Mutually exclusive DrawingML geometry of a worksheet shape.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Geometry {
    /// A schema-defined preset (`a:prstGeom`).
    Preset(Preset),
    /// A custom geometry (`a:custGeom`).
    Custom(Box<CustomGeometry>),
}

impl Geometry {
    /// Return the preset when this is a preset geometry.
    pub const fn preset(&self) -> Option<Preset> {
        match self {
            Self::Preset(preset) => Some(*preset),
            Self::Custom(_) => None,
        }
    }

    /// Borrow the custom geometry when this is custom.
    pub fn custom(&self) -> Option<&CustomGeometry> {
        match self {
            Self::Preset(_) => None,
            Self::Custom(geometry) => Some(geometry.as_ref()),
        }
    }

    /// Move out the custom geometry when this is custom.
    pub fn into_custom(self) -> Option<CustomGeometry> {
        match self {
            Self::Preset(_) => None,
            Self::Custom(geometry) => Some(*geometry),
        }
    }
}

impl From<Preset> for Geometry {
    fn from(preset: Preset) -> Self {
        Self::Preset(preset)
    }
}

impl From<CustomGeometry> for Geometry {
    fn from(geometry: CustomGeometry) -> Self {
        Self::Custom(Box::new(geometry))
    }
}

/// Non-visual identity shared by all drawing objects (`xdr:cNvPr` and lock flags).
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct NonVisual {
    /// Drawing object ID (`xdr:cNvPr@id`), when declared and well-formed.
    pub id: Option<u32>,
    /// Object name (`xdr:cNvPr@name`).
    pub name: Option<String>,
    /// Alternative text (`xdr:cNvPr@descr`).
    pub description: Option<String>,
    /// Whether the object is hidden (`xdr:cNvPr@hidden`).
    pub hidden: bool,
    /// Whether any lock flag (`a:spLocks`/`a:cxnSpLocks`/`a:grpSpLocks`) is set.
    pub locked: bool,
}

/// A DrawingML shape (`xdr:sp`), typically a text box.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Shape {
    /// Non-visual identity and flags.
    pub non_visual: NonVisual,
    /// Whether the shape is a text box (`xdr:cNvSpPr@txBox`).
    pub is_text_box: bool,
    /// Mutually exclusive preset or custom geometry, when declared.
    pub geometry: Option<Geometry>,
    /// Rich-text story (`xdr:txBody`), when present.
    pub text_body: Option<Body>,
}

impl Shape {
    /// Return the declared preset geometry, if any.
    pub fn preset(&self) -> Option<Preset> {
        self.geometry.as_ref().and_then(Geometry::preset)
    }

    /// Borrow the declared custom geometry, if any.
    pub fn custom_geometry(&self) -> Option<&CustomGeometry> {
        self.geometry.as_ref().and_then(Geometry::custom)
    }
}

/// One end of a connection shape (`a:stCxn`/`a:endCxn`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ConnectionEnd {
    /// Drawing object ID of the connected shape (`@id`).
    pub shape_id: u32,
    /// Connection site index on the connected shape (`@idx`).
    pub site: u32,
}

/// A connection shape (`xdr:cxnSp`) linking two shapes.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ConnectionShape {
    /// Non-visual identity and flags.
    pub non_visual: NonVisual,
    /// Mutually exclusive preset or custom geometry, when declared.
    pub geometry: Option<Geometry>,
    /// Start connection, when declared.
    pub start: Option<ConnectionEnd>,
    /// End connection, when declared.
    pub end: Option<ConnectionEnd>,
    /// Rich-text story (`xdr:txBody`), when present.
    pub text_body: Option<Body>,
}

impl ConnectionShape {
    /// Return the declared preset geometry, if any.
    pub fn preset(&self) -> Option<Preset> {
        self.geometry.as_ref().and_then(Geometry::preset)
    }

    /// Borrow the declared custom geometry, if any.
    pub fn custom_geometry(&self) -> Option<&CustomGeometry> {
        self.geometry.as_ref().and_then(Geometry::custom)
    }
}

/// Group coordinate transform (`xdr:grpSpPr/a:xfrm`).
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct GroupTransform {
    /// Group offset (`a:off`), when declared.
    pub offset: Option<EmuOffset>,
    /// Group extent (`a:ext`), when declared.
    pub extent: Option<EmuExtent>,
    /// Child coordinate-space offset (`a:chOff`), when declared.
    pub child_offset: Option<EmuOffset>,
    /// Child coordinate-space extent (`a:chExt`), when declared.
    pub child_extent: Option<EmuExtent>,
}

/// A shape group (`xdr:grpSp`) with its nested objects.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Group {
    /// Non-visual identity and flags.
    pub non_visual: NonVisual,
    /// Group coordinate transform, when declared.
    pub transform: Option<GroupTransform>,
    /// Nested objects in document order; groups may nest.
    pub children: Vec<Object>,
}

/// Inert metadata of a legacy OLE object anchored through a
/// `xdr:graphicFrame` (`xdr:oleObject`).
///
/// The referenced payload and link targets are recorded as relationship IDs
/// only; they are never resolved, fetched, or activated.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct OleObject {
    /// Non-visual identity of the hosting graphic frame.
    pub non_visual: NonVisual,
    /// OLE program ID (`@progId`), when declared.
    pub program_id: Option<String>,
    /// Shape ID linked to the worksheet OLE object record (`@shapeId`).
    pub shape_id: Option<u32>,
    /// Typed data-or-view aspect (`@dvAspect`).
    pub data_or_view_aspect: Option<Aspect>,
    /// Whether the object loads automatically (`@autoLoad`), when declared.
    pub auto_load: Option<bool>,
    /// Relationship ID of the embedded object (`r:id`), when declared.
    pub relationship_id: Option<String>,
    /// Relationship ID of the linked object (`r:link`), when declared.
    pub link_relationship_id: Option<String>,
}

/// One drawing object carried by an anchor or nested in a group.
///
/// Pictures and chart graphic frames are deliberately not represented: they
/// are covered by the image and chart support.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Object {
    /// A DrawingML shape or text box.
    Shape(Shape),
    /// A connection shape.
    ConnectionShape(ConnectionShape),
    /// A shape group.
    Group(Group),
    /// A legacy OLE object hosted by a graphic frame.
    OleObject(OleObject),
}

/// Sheet-interaction flags of one anchor (`xdr:clientData`).
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct ClientData {
    /// Whether the object locks with the sheet (`fLocksWithSheet`), when declared.
    pub locks_with_sheet: Option<bool>,
    /// Whether the object prints with the sheet (`fPrintsWithSheet`), when declared.
    pub prints_with_sheet: Option<bool>,
}

/// One anchored drawing object on a worksheet.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AnchoredObject {
    /// How the object is anchored to the grid.
    pub anchor: Anchor,
    /// The anchored object.
    pub object: Object,
    /// Sheet-interaction flags from the anchor's `xdr:clientData`.
    pub client_data: ClientData,
}

/// Shape inventory of one worksheet drawing part.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Shapes {
    /// Worksheet name from the workbook.
    pub worksheet_name: String,
    /// Worksheet part name (for example `/xl/worksheets/sheet1.xml`).
    pub worksheet_part_name: String,
    /// Anchored objects in drawing order.
    pub objects: Vec<AnchoredObject>,
}

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

/// Load the shape inventory of every worksheet in a workbook package.
///
/// One entry is returned per worksheet that anchors at least one shape-like
/// object; worksheets without shapes are omitted.
pub fn load_shapes(package: &OpcPackage) -> Result<Vec<Shapes>> {
    let workbook_part = package.main_document_part()?;
    let sheets = parse_workbook_sheets(workbook_part.blob())?;
    let mut output = Vec::new();
    for sheet in &sheets {
        let shapes = load_shapes_for_sheet(package, workbook_part, sheet)?;
        if !shapes.objects.is_empty() {
            output.push(shapes);
        }
    }
    Ok(output)
}

/// Load the shape inventory of one worksheet, addressed by sheet name.
pub fn load_sheet_shapes(package: &OpcPackage, sheet_name: &str) -> Result<Shapes> {
    let workbook_part = package.main_document_part()?;
    let sheets = parse_workbook_sheets(workbook_part.blob())?;
    let sheet = sheets
        .iter()
        .find(|sheet| sheet.name == sheet_name)
        .ok_or_else(|| invalid(format!("worksheet '{sheet_name}' not found")))?;
    load_shapes_for_sheet(package, workbook_part, sheet)
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
    Other,
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

    #[allow(clippy::too_many_lines)]
    fn start(
        &mut self,
        parent: Context,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
        decoder: Decoder,
        resolver: &NamespaceResolver,
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
                _ => return Ok(Context::Other),
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
                b"clientData" => {
                    let client_data = ClientData {
                        locks_with_sheet: bool_attribute(element, b"fLocksWithSheet", decoder)?,
                        prints_with_sheet: bool_attribute(element, b"fPrintsWithSheet", decoder)?,
                    };
                    self.anchor_mut()?.client_data = client_data;
                    return Ok(Context::Other);
                },
                _ => return Ok(Context::Other),
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
                },
                b"cNvSpPr" if !self.builders.is_empty() => {
                    self.builder_mut()?.is_text_box =
                        bool_attribute(element, b"txBox", decoder)?.unwrap_or(false);
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
        Ok(Context::Other)
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

fn parse_workbook_sheets(xml: &[u8]) -> Result<Vec<Sheet>> {
    if xml.len() > MAX_WORKBOOK_BYTES {
        return Err(limit("workbook XML bytes"));
    }
    Ok(parse_catalog(xml)?.sheets)
}

fn load_shapes_for_sheet(
    package: &OpcPackage,
    workbook_part: &dyn Part,
    sheet: &Sheet,
) -> Result<Shapes> {
    let relationship = workbook_part
        .rels()
        .get(&sheet.relationship_id)
        .ok_or_else(|| {
            invalid(format!(
                "worksheet '{}' references missing relationship '{}'",
                sheet.name, sheet.relationship_id
            ))
        })?;
    if !matches!(relationship.reltype(), rt::WORKSHEET | rt::STRICT_WORKSHEET) {
        return Err(invalid(format!(
            "worksheet '{}' relationship has invalid type '{}'",
            sheet.name,
            relationship.reltype()
        )));
    }
    if relationship.is_external() {
        return Err(invalid(format!(
            "worksheet '{}' relationship cannot be external",
            sheet.name
        )));
    }
    let sheet_uri = relationship.target_partname()?;
    let sheet_part = package.get_part(&sheet_uri)?;
    if sheet_part.content_type() != ct::SML_WORKSHEET {
        return Err(invalid(format!(
            "worksheet part has content type '{}', expected '{}'",
            sheet_part.content_type(),
            ct::SML_WORKSHEET
        )));
    }
    let mut objects = Vec::new();
    let drawings: Vec<_> = sheet_part
        .rels()
        .iter()
        .filter(|relationship| matches!(relationship.reltype(), rt::DRAWING | rt::STRICT_DRAWING))
        .collect();
    if drawings.len() > MAX_DRAWINGS_PER_WORKSHEET {
        return Err(limit("drawings per worksheet"));
    }
    for drawing_relationship in drawings {
        if drawing_relationship.is_external() {
            return Err(invalid("worksheet drawing relationship cannot be external"));
        }
        let drawing_uri = drawing_relationship.target_partname()?;
        let drawing_part = package.get_part(&drawing_uri)?;
        if drawing_part.content_type() != ct::OFC_DRAWING {
            return Err(invalid(format!(
                "drawing part has content type '{}', expected '{}'",
                drawing_part.content_type(),
                ct::OFC_DRAWING
            )));
        }
        if drawing_part.blob().len() > MAX_DRAWING_PART_BYTES {
            return Err(limit("drawing part bytes"));
        }
        let drawing_xml = std::str::from_utf8(drawing_part.blob())
            .map_err(|error| Error::Invalid(error.to_string()))?;
        let Some(anchored) = parse_drawing_shapes(drawing_xml)? else {
            continue;
        };
        for object in anchored {
            if objects.len() >= MAX_ANCHORS_PER_DRAWING {
                return Err(limit("shapes per worksheet"));
            }
            objects.push(object);
        }
    }
    Ok(Shapes {
        worksheet_name: sheet.name.clone(),
        worksheet_part_name: sheet_uri.to_string(),
        objects,
    })
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

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::BlobPart;
    use litchi_opc::PackURI;
    use std::mem::size_of;

    const XDR: &str = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    const STRICT_XDR: &str = "http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing";
    const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
    const STRICT_A: &str = "http://purl.oclc.org/ooxml/drawingml/main";
    const C: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    const POI_TEXT_BOXES: &[u8] =
        include_bytes!("../../../test-data/poi/test-data/spreadsheet/45540_form_Header.xlsx");

    fn marker(col: u32, col_off: i64, row: u32, row_off: i64) -> String {
        format!(
            "<xdr:col>{col}</xdr:col><xdr:colOff>{col_off}</xdr:colOff>\
             <xdr:row>{row}</xdr:row><xdr:rowOff>{row_off}</xdr:rowOff>"
        )
    }

    fn drawing(body: &str) -> String {
        format!("<xdr:wsDr xmlns:xdr=\"{XDR}\" xmlns:a=\"{A}\" xmlns:r=\"{R}\">{body}</xdr:wsDr>")
    }

    fn two_cell_anchor(object: &str) -> String {
        format!(
            "<xdr:twoCellAnchor editAs=\"oneCell\"><xdr:from>{}</xdr:from><xdr:to>{}</xdr:to>\
             {object}<xdr:clientData fLocksWithSheet=\"0\" fPrintsWithSheet=\"1\"/></xdr:twoCellAnchor>",
            marker(1, 100, 2, 200),
            marker(5, 300, 9, 400)
        )
    }

    fn text_box_shape() -> &'static str {
        "<xdr:sp macro=\"\" textlink=\"\">\
         <xdr:nvSpPr><xdr:cNvPr id=\"7\" name=\"Text Box 7\" descr=\"alt\" hidden=\"1\"/>\
         <xdr:cNvSpPr txBox=\"1\"><a:spLocks noChangeArrowheads=\"1\"/></xdr:cNvSpPr></xdr:nvSpPr>\
         <xdr:spPr><a:prstGeom prst=\"roundRect\"><a:avLst/></a:prstGeom></xdr:spPr>\
         <xdr:txBody><a:bodyPr lIns=\"182880\" tIns=\"91440\" rIns=\"182880\" bIns=\"91440\" \
         anchor=\"ctr\" anchorCtr=\"1\" vert=\"vert270\" wrap=\"none\" numCol=\"2\" \
         spcFirstLastPara=\"1\"><a:spAutoFit/></a:bodyPr><a:lstStyle/>\
         <a:p><a:pPr algn=\"l\"><a:defRPr sz=\"1000\"/></a:pPr>\
         <a:r><a:rPr lang=\"en-US\" sz=\"1200\" b=\"1\" i=\"true\" u=\"sng\"/><a:t>Bold</a:t></a:r>\
         <a:r><a:t xml:space=\"preserve\"> plain</a:t></a:r><a:br/></a:p>\
         <a:p><a:r><a:t>Second</a:t></a:r></a:p>\
         </xdr:txBody></xdr:sp>"
    }

    #[test]
    fn parses_two_cell_text_box() {
        let xml = drawing(&two_cell_anchor(text_box_shape()));
        let objects = parse_drawing_shapes(&xml).unwrap().unwrap();
        assert_eq!(objects.len(), 1);
        let anchored = &objects[0];
        assert_eq!(
            anchored.anchor,
            Anchor::TwoCell {
                from: CellMarker {
                    column: 1,
                    column_offset: Emu(100),
                    row: 2,
                    row_offset: Emu(200),
                },
                to: CellMarker {
                    column: 5,
                    column_offset: Emu(300),
                    row: 9,
                    row_offset: Emu(400),
                },
                edit_as: EditAs::OneCell,
            }
        );
        assert_eq!(anchored.client_data.locks_with_sheet, Some(false));
        assert_eq!(anchored.client_data.prints_with_sheet, Some(true));
        let Object::Shape(shape) = &anchored.object else {
            panic!("expected a shape");
        };
        assert_eq!(shape.non_visual.id, Some(7));
        assert_eq!(shape.non_visual.name.as_deref(), Some("Text Box 7"));
        assert_eq!(shape.non_visual.description.as_deref(), Some("alt"));
        assert!(shape.non_visual.hidden);
        assert!(shape.non_visual.locked);
        assert!(shape.is_text_box);
        assert_eq!(shape.preset(), Some(Preset::RoundRect));
        let body = shape.text_body.as_ref().unwrap();
        let properties = &body.properties;
        assert_eq!(properties.insets.left.as_emu(), Some(182880));
        assert_eq!(properties.insets.bottom.as_emu(), Some(91440));
        assert_eq!(properties.vertical_anchor, VerticalAnchor::Center);
        assert!(properties.anchor_center);
        assert_eq!(properties.direction, Direction::Vertical270);
        assert_eq!(properties.wrap, Wrap::None);
        assert_eq!(properties.autofit, Autofit::Shape);
        assert_eq!(properties.column_count.get(), 2);
        assert!(properties.space_first_last_paragraph);
        assert_eq!(body.paragraphs.len(), 2);
        let bold = &body.paragraphs[0].runs[0];
        assert_eq!(bold.text, "Bold");
        assert_eq!(bold.bold, Some(true));
        assert_eq!(bold.italic, Some(true));
        assert_eq!(bold.underline, Some(Underline::Single));
        assert_eq!(bold.font_size.map(TextSize::get), Some(1200));
        assert_eq!(body.paragraphs[0].runs[1].text, " plain");
        assert_eq!(body.paragraphs[0].runs[1].bold, None);
        // The break contributes a newline run.
        assert_eq!(body.paragraphs[0].runs[2].text, "\n");
        assert_eq!(body.text(), "Bold plain\n\nSecond");
    }

    #[test]
    fn preset_attributes_apply_xml_schema_token_whitespace() {
        let shape =
            text_box_shape().replace("prst=\"roundRect\"", "prst=\" &#x9;roundRect&#xA;&#xD; \"");
        let objects = parse_drawing_shapes(&drawing(&two_cell_anchor(&shape)))
            .unwrap()
            .unwrap();
        let Object::Shape(shape) = &objects[0].object else {
            panic!("expected a shape");
        };
        assert_eq!(shape.preset(), Some(Preset::RoundRect));
    }

    #[test]
    fn parses_one_cell_connection_shape() {
        let object = "<xdr:cxnSp><xdr:nvCxnSpPr><xdr:cNvPr id=\"9\" name=\"Connector 9\"/>\
            <xdr:cNvCxnSpPr><a:stCxn id=\"7\" idx=\"3\"/><a:endCxn id=\"11\" idx=\"1\"/></xdr:cNvCxnSpPr>\
            </xdr:nvCxnSpPr><xdr:spPr><a:prstGeom prst=\"bentConnector3\"/></xdr:spPr></xdr:cxnSp>";
        let anchor = format!(
            "<xdr:oneCellAnchor><xdr:from>{}</xdr:from>\
             <xdr:ext cx=\"914400\" cy=\"457200\"/>{object}<xdr:clientData/></xdr:oneCellAnchor>",
            marker(0, 0, 0, 0)
        );
        let objects = parse_drawing_shapes(&drawing(&anchor)).unwrap().unwrap();
        assert_eq!(objects.len(), 1);
        assert_eq!(
            objects[0].anchor,
            Anchor::OneCell {
                from: CellMarker::default(),
                extent: EmuExtent {
                    width: Emu(914400),
                    height: Emu(457200),
                },
            }
        );
        assert_eq!(objects[0].client_data, ClientData::default());
        let Object::ConnectionShape(connection) = &objects[0].object else {
            panic!("expected a connection shape");
        };
        assert_eq!(connection.non_visual.id, Some(9));
        assert!(!connection.non_visual.locked);
        assert_eq!(connection.preset(), Some(Preset::BentConnector3));
        assert_eq!(
            connection.start,
            Some(ConnectionEnd {
                shape_id: 7,
                site: 3,
            })
        );
        assert_eq!(
            connection.end,
            Some(ConnectionEnd {
                shape_id: 11,
                site: 1,
            })
        );
        assert!(connection.text_body.is_none());
    }

    #[test]
    fn connection_custom_geometry_is_typed_and_exclusive() {
        let object = "<xdr:cxnSp><xdr:nvCxnSpPr><xdr:cNvPr id=\"9\" name=\"Custom\"/>\
            <xdr:cNvCxnSpPr/></xdr:nvCxnSpPr><xdr:spPr>\
            <a:custGeom><a:pathLst/></a:custGeom></xdr:spPr></xdr:cxnSp>";
        let anchor = format!(
            "<xdr:oneCellAnchor><xdr:from>{}</xdr:from>\
             <xdr:ext cx=\"914400\" cy=\"457200\"/>{object}<xdr:clientData/></xdr:oneCellAnchor>",
            marker(0, 0, 0, 0)
        );
        let objects = parse_drawing_shapes(&drawing(&anchor)).unwrap().unwrap();
        let Object::ConnectionShape(connection) = &objects[0].object else {
            panic!("expected a connection shape");
        };
        assert_eq!(connection.preset(), None);
        assert!(connection.custom_geometry().is_some());
    }

    #[test]
    fn rejects_unknown_preset() {
        let object = "<xdr:sp><xdr:nvSpPr><xdr:cNvPr id=\"3\" name=\"Odd\"/><xdr:cNvSpPr/>\
            </xdr:nvSpPr><xdr:spPr><a:prstGeom prst=\"vendorWeird\"/></xdr:spPr></xdr:sp>";
        let anchor = format!(
            "<xdr:absoluteAnchor><xdr:pos x=\"123\" y=\"456\"/>\
             <xdr:ext cx=\"789\" cy=\"101\"/>{object}<xdr:clientData/></xdr:absoluteAnchor>"
        );
        let error = parse_drawing_shapes(&drawing(&anchor)).unwrap_err();
        assert!(error.to_string().contains("vendorWeird"));
    }

    #[test]
    fn rejects_missing_presets_and_invalid_text_domains() {
        let missing = "<xdr:sp><xdr:nvSpPr><xdr:cNvPr id=\"3\" name=\"Odd\"/><xdr:cNvSpPr/>\
            </xdr:nvSpPr><xdr:spPr><a:prstGeom/></xdr:spPr></xdr:sp>";
        let anchor = format!(
            "<xdr:absoluteAnchor><xdr:pos x=\"123\" y=\"456\"/>\
             <xdr:ext cx=\"789\" cy=\"101\"/>{missing}<xdr:clientData/></xdr:absoluteAnchor>"
        );
        let error = parse_drawing_shapes(&drawing(&anchor)).unwrap_err();
        assert!(error.to_string().contains("missing required prst"));

        assert!(parse_drawing_shapes(&drawing("<xdr:twoCellAnchor editAs=\"vendor\"/>")).is_err());

        for attribute in [
            "anchor=\"middle\"",
            "vert=\"diagonal\"",
            "wrap=\"tight\"",
            "anchorCtr=\"on\"",
            "numCol=\"17\"",
            "lIns=\"2147483648\"",
        ] {
            let object = format!(
                "<xdr:sp><xdr:nvSpPr><xdr:cNvPr id=\"3\" name=\"Text\"/><xdr:cNvSpPr/>\
                 </xdr:nvSpPr><xdr:txBody><a:bodyPr {attribute}/><a:p/></xdr:txBody></xdr:sp>"
            );
            assert!(
                parse_drawing_shapes(&drawing(&two_cell_anchor(&object))).is_err(),
                "accepted {attribute}"
            );
        }

        for run_properties in ["u=\"vendor\"", "sz=\"99\"", "b=\"on\""] {
            let object = format!(
                "<xdr:sp><xdr:nvSpPr><xdr:cNvPr id=\"3\" name=\"Text\"/><xdr:cNvSpPr/>\
                 </xdr:nvSpPr><xdr:txBody><a:bodyPr/><a:p><a:r><a:rPr {run_properties}/>\
                 <a:t>x</a:t></a:r></a:p></xdr:txBody></xdr:sp>"
            );
            assert!(
                parse_drawing_shapes(&drawing(&two_cell_anchor(&object))).is_err(),
                "accepted {run_properties}"
            );
        }
    }

    #[test]
    fn rejects_competing_shape_geometries() {
        let object = "<xdr:sp><xdr:nvSpPr><xdr:cNvPr id=\"3\" name=\"Both\"/><xdr:cNvSpPr/>\
            </xdr:nvSpPr><xdr:spPr><a:prstGeom prst=\"rect\"/>\
            <a:custGeom><a:pathLst/></a:custGeom></xdr:spPr></xdr:sp>";
        let anchor = format!(
            "<xdr:absoluteAnchor><xdr:pos x=\"123\" y=\"456\"/>\
             <xdr:ext cx=\"789\" cy=\"101\"/>{object}<xdr:clientData/></xdr:absoluteAnchor>"
        );
        let error = parse_drawing_shapes(&drawing(&anchor)).unwrap_err();
        assert!(error.to_string().contains("competing"));
    }

    #[test]
    fn geometry_keeps_the_custom_payload_off_the_hot_path() {
        assert!(size_of::<Geometry>() <= 2 * size_of::<usize>());
    }

    #[test]
    fn parses_nested_groups() {
        let inner = "<xdr:sp><xdr:nvSpPr><xdr:cNvPr id=\"21\" name=\"Inner\"/><xdr:cNvSpPr/>\
            </xdr:nvSpPr><xdr:spPr><a:prstGeom prst=\"ellipse\"/></xdr:spPr></xdr:sp>";
        let nested_group = "<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id=\"22\" name=\"Nested\"/>\
             <xdr:cNvGrpSpPr><a:grpSpLocks noChangeAspect=\"1\"/></xdr:cNvGrpSpPr></xdr:nvGrpSpPr>\
             <xdr:grpSpPr/><xdr:sp><xdr:nvSpPr><xdr:cNvPr id=\"23\" name=\"Deep\"/><xdr:cNvSpPr/>\
             </xdr:nvSpPr><xdr:spPr/></xdr:sp></xdr:grpSp>"
            .to_string();
        let group = format!(
            "<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id=\"20\" name=\"Group\"/><xdr:cNvGrpSpPr/>\
             </xdr:nvGrpSpPr><xdr:grpSpPr><a:xfrm><a:off x=\"1\" y=\"2\"/><a:ext cx=\"3\" cy=\"4\"/>\
             <a:chOff x=\"5\" y=\"6\"/><a:chExt cx=\"7\" cy=\"8\"/></a:xfrm></xdr:grpSpPr>\
             {inner}{nested_group}</xdr:grpSp>"
        );
        let objects = parse_drawing_shapes(&drawing(&two_cell_anchor(&group)))
            .unwrap()
            .unwrap();
        let Object::Group(group) = &objects[0].object else {
            panic!("expected a group");
        };
        assert_eq!(group.non_visual.id, Some(20));
        let transform = group.transform.unwrap();
        assert_eq!(transform.offset.unwrap().x, Emu(1));
        assert_eq!(transform.extent.unwrap().height, Emu(4));
        assert_eq!(transform.child_offset.unwrap().y, Emu(6));
        assert_eq!(transform.child_extent.unwrap().width, Emu(7));
        assert_eq!(group.children.len(), 2);
        let Object::Shape(inner) = &group.children[0] else {
            panic!("expected an inner shape");
        };
        assert_eq!(inner.preset(), Some(Preset::Ellipse));
        let Object::Group(nested) = &group.children[1] else {
            panic!("expected a nested group");
        };
        assert!(nested.non_visual.locked);
        assert!(nested.transform.is_none());
        assert_eq!(nested.children.len(), 1);
        assert_eq!(nested.non_visual.name.as_deref(), Some("Nested"));
    }

    #[test]
    fn skips_pictures_and_charts_but_keeps_ole_objects() {
        let picture = "<xdr:pic><xdr:nvPicPr><xdr:cNvPr id=\"50\" name=\"Logo\"/></xdr:nvPicPr>\
            <xdr:blipFill><a:blip r:embed=\"rId9\"/></xdr:blipFill></xdr:pic>";
        let chart = &format!(
            "<xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id=\"51\" name=\"Chart\"/>\
            </xdr:nvGraphicFramePr><a:graphic><a:graphicData>\
            <c:chart xmlns:c=\"{C}\" r:id=\"rId8\"/></a:graphicData></a:graphic></xdr:graphicFrame>"
        );
        let ole = "<xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id=\"52\" name=\"Object\"/>\
            </xdr:nvGraphicFramePr><a:graphic><a:graphicData>\
            <xdr:oleObject progId=\"Excel.Sheet.12\" shapeId=\"1027\" dvAspect=\"DVASPECT_ICON\" \
            autoLoad=\"1\" r:id=\"rId7\" r:link=\"rId6\"/></a:graphicData></a:graphic></xdr:graphicFrame>";
        let body = format!(
            "{}{}{}",
            two_cell_anchor(picture),
            two_cell_anchor(chart),
            two_cell_anchor(ole)
        );
        let objects = parse_drawing_shapes(&drawing(&body)).unwrap().unwrap();
        assert_eq!(objects.len(), 1);
        let Object::OleObject(ole_object) = &objects[0].object else {
            panic!("expected an OLE object");
        };
        assert_eq!(ole_object.non_visual.id, Some(52));
        assert_eq!(ole_object.non_visual.name.as_deref(), Some("Object"));
        assert_eq!(ole_object.program_id.as_deref(), Some("Excel.Sheet.12"));
        assert_eq!(ole_object.shape_id, Some(1027));
        assert_eq!(ole_object.data_or_view_aspect, Some(Aspect::Icon));
        assert_eq!(ole_object.auto_load, Some(true));
        assert_eq!(ole_object.relationship_id.as_deref(), Some("rId7"));
        assert_eq!(ole_object.link_relationship_id.as_deref(), Some("rId6"));
    }

    #[test]
    fn rejects_unknown_ole_object_aspect() {
        let ole = "<xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id=\"52\" name=\"Object\"/>\
            </xdr:nvGraphicFramePr><a:graphic><a:graphicData>\
            <xdr:oleObject shapeId=\"1027\" dvAspect=\"DVASPECT_THUMBNAIL\" \
            r:id=\"rId7\"/></a:graphicData></a:graphic></xdr:graphicFrame>";
        let xml = drawing(&two_cell_anchor(ole));

        let error = parse_drawing_shapes(&xml).unwrap_err();

        assert!(error.to_string().contains("DVASPECT_THUMBNAIL"));
    }

    #[test]
    fn parses_strict_namespace_dialect() {
        let strict_marker = |col: u32, row: u32| {
            format!(
                "<s:col>{col}</s:col><s:colOff>0</s:colOff>\
                 <s:row>{row}</s:row><s:rowOff>0</s:rowOff>"
            )
        };
        let xml = format!(
            "<s:wsDr xmlns:s=\"{STRICT_XDR}\" xmlns:a=\"{STRICT_A}\"><s:twoCellAnchor>\
             <s:from>{}</s:from><s:to>{}</s:to>\
             <s:sp><s:nvSpPr><s:cNvPr id=\"1\" name=\"Strict\"/><s:cNvSpPr txBox=\"1\"/></s:nvSpPr>\
             <s:spPr><a:prstGeom prst=\"rect\"/></s:spPr>\
             <s:txBody><a:bodyPr/><a:p><a:r><a:t>S</a:t></a:r></a:p></s:txBody></s:sp>\
             <s:clientData/></s:twoCellAnchor></s:wsDr>",
            strict_marker(0, 0),
            strict_marker(1, 1)
        );
        let objects = parse_drawing_shapes(&xml).unwrap().unwrap();
        let Object::Shape(shape) = &objects[0].object else {
            panic!("expected a shape");
        };
        assert!(shape.is_text_box);
        assert_eq!(shape.preset(), Some(Preset::Rect));
        assert_eq!(shape.text_body.as_ref().unwrap().text(), "S");
        // ECMA-376 default body properties apply when a:bodyPr is empty.
        let properties = &shape.text_body.as_ref().unwrap().properties;
        assert_eq!(
            properties.insets.left.as_emu(),
            Insets::default().left.as_emu()
        );
        assert_eq!(properties.column_count, Columns::ONE);
    }

    #[test]
    fn tolerates_empty_drawing_and_non_drawing_root() {
        let empty = drawing("");
        assert_eq!(
            parse_drawing_shapes(&empty).unwrap().unwrap(),
            Vec::<AnchoredObject>::new()
        );
        let empty_root = format!("<xdr:wsDr xmlns:xdr=\"{XDR}\"/>");
        assert!(
            parse_drawing_shapes(&empty_root)
                .unwrap()
                .unwrap()
                .is_empty()
        );
        assert!(parse_drawing_shapes("<other/>").unwrap().is_none());
    }

    #[test]
    fn rejects_malformed_drawings() {
        let anchored_shape = two_cell_anchor(text_box_shape());
        let cases = [
            // DTD is rejected.
            format!("<!DOCTYPE xdr:wsDr>{}", drawing(&anchored_shape)),
            // Processing instructions are rejected.
            drawing(&format!("<?xml-stylesheet href=\"x\"?>{anchored_shape}")),
            // Multiple roots.
            format!("{}{}", drawing(""), drawing("")),
            // Unterminated root.
            "<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\">".to_string(),
            // Shape anchor without markers.
            drawing(
                "<xdr:twoCellAnchor><xdr:sp><xdr:nvSpPr><xdr:cNvPr id=\"1\"/>\
                 <xdr:cNvSpPr/></xdr:nvSpPr></xdr:sp><xdr:clientData/></xdr:twoCellAnchor>",
            ),
            // Invalid marker value.
            drawing(
                "<xdr:twoCellAnchor><xdr:from><xdr:col>x</xdr:col></xdr:from>\
                 <xdr:to/><xdr:sp><xdr:nvSpPr><xdr:cNvPr id=\"1\"/><xdr:cNvSpPr/></xdr:nvSpPr>\
                 </xdr:sp><xdr:clientData/></xdr:twoCellAnchor>",
            ),
            // Marker outside worksheet bounds.
            drawing(
                "<xdr:twoCellAnchor><xdr:from><xdr:col>16384</xdr:col><xdr:colOff>0</xdr:colOff>\
                 <xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>\
                 <xdr:to><xdr:col>16385</xdr:col><xdr:colOff>0</xdr:colOff>\
                 <xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>\
                 <xdr:sp><xdr:nvSpPr><xdr:cNvPr id=\"1\"/><xdr:cNvSpPr/></xdr:nvSpPr></xdr:sp>\
                 <xdr:clientData/></xdr:twoCellAnchor>",
            ),
            // Two objects in one anchor.
            drawing(&format!(
                "<xdr:twoCellAnchor><xdr:from>{}</xdr:from><xdr:to>{}</xdr:to>\
                 <xdr:sp><xdr:nvSpPr><xdr:cNvPr id=\"1\"/><xdr:cNvSpPr/></xdr:nvSpPr></xdr:sp>\
                 <xdr:sp><xdr:nvSpPr><xdr:cNvPr id=\"2\"/><xdr:cNvSpPr/></xdr:nvSpPr></xdr:sp>\
                 <xdr:clientData/></xdr:twoCellAnchor>",
                marker(0, 0, 0, 0),
                marker(1, 0, 1, 0)
            )),
            // One-cell anchor without extent.
            drawing(&format!(
                "<xdr:oneCellAnchor><xdr:from>{}</xdr:from>\
                 <xdr:sp><xdr:nvSpPr><xdr:cNvPr id=\"1\"/><xdr:cNvSpPr/></xdr:nvSpPr></xdr:sp>\
                 <xdr:clientData/></xdr:oneCellAnchor>",
                marker(0, 0, 0, 0)
            )),
        ];
        for xml in cases {
            assert!(parse_drawing_shapes(&xml).is_err(), "accepted {xml}");
        }
    }

    fn package_with_shapes(drawing_xml: &str) -> OpcPackage {
        let mut package = OpcPackage::new();
        let mut workbook_part = BlobPart::new(
            PackURI::new("/xl/workbook.xml").unwrap(),
            ct::SML_SHEET_MAIN.to_string(),
            format!(
                "<workbook xmlns=\"{SML}\" xmlns:r=\"{R}\">\
                 <sheets><sheet name=\"Data\" sheetId=\"1\" r:id=\"rId1\"/>\
                 <sheet name=\"Empty\" sheetId=\"2\" r:id=\"rId2\"/></sheets></workbook>"
            )
            .into_bytes(),
        );
        workbook_part.relate_to("worksheets/sheet1.xml", rt::WORKSHEET);
        workbook_part.relate_to("worksheets/sheet2.xml", rt::WORKSHEET);
        let mut sheet_part = BlobPart::new(
            PackURI::new("/xl/worksheets/sheet1.xml").unwrap(),
            ct::SML_WORKSHEET.to_string(),
            format!("<worksheet xmlns=\"{SML}\"><sheetData/></worksheet>").into_bytes(),
        );
        sheet_part.relate_to("../drawings/drawing1.xml", rt::DRAWING);
        package.relate_to(
            "xl/workbook.xml",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
        );
        package.add_part(Box::new(workbook_part));
        package.add_part(Box::new(sheet_part));
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/xl/worksheets/sheet2.xml").unwrap(),
            ct::SML_WORKSHEET.to_string(),
            format!("<worksheet xmlns=\"{SML}\"><sheetData/></worksheet>").into_bytes(),
        )));
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/xl/drawings/drawing1.xml").unwrap(),
            ct::OFC_DRAWING.to_string(),
            drawing_xml.as_bytes().to_vec(),
        )));
        package
    }

    #[test]
    fn loads_shapes_through_the_package_graph() {
        let package = package_with_shapes(&drawing(&two_cell_anchor(text_box_shape())));
        let shapes = load_sheet_shapes(&package, "Data").unwrap();
        assert_eq!(shapes.worksheet_name, "Data");
        assert_eq!(shapes.worksheet_part_name, "/xl/worksheets/sheet1.xml");
        assert_eq!(shapes.objects.len(), 1);
        let Object::Shape(shape) = &shapes.objects[0].object else {
            panic!("expected a shape");
        };
        assert_eq!(shape.non_visual.name.as_deref(), Some("Text Box 7"));

        // Worksheets without shapes are omitted from the workbook inventory.
        let all = load_shapes(&package).unwrap();
        assert_eq!(all.len(), 1);
        assert_eq!(all[0].worksheet_name, "Data");
        assert!(
            load_sheet_shapes(&package, "Empty")
                .unwrap()
                .objects
                .is_empty()
        );
        assert!(load_sheet_shapes(&package, "Missing").is_err());
    }

    #[cfg(any())]
    #[test]
    fn workbook_accessor_loads_shapes() {
        let package = package_with_shapes(&drawing(&two_cell_anchor(text_box_shape())));
        let workbook = crate::workbook::Workbook::from_package(package).unwrap();
        let shapes = workbook.shapes_on_sheet("Data").unwrap();
        assert_eq!(shapes.objects.len(), 1);
        assert!(workbook.shapes_on_sheet("Missing").is_err());
    }

    #[test]
    fn poi_fixture_text_boxes_parse() {
        let package = OpcPackage::from_bytes(POI_TEXT_BOXES).unwrap();
        let all = load_shapes(&package).unwrap();
        let names: Vec<_> = all
            .iter()
            .flat_map(|sheet| sheet.objects.iter())
            .filter_map(|anchored| match &anchored.object {
                Object::Shape(shape) => Some(shape),
                _ => None,
            })
            .collect();
        assert_eq!(names.len(), 4);
        let text_box = names
            .iter()
            .find(|shape| shape.non_visual.name.as_deref() == Some("Text Box 35"))
            .unwrap();
        assert!(text_box.is_text_box);
        assert!(text_box.non_visual.locked);
        assert!(!text_box.non_visual.hidden);
        assert_eq!(text_box.preset(), Some(Preset::Rect));
        let body = text_box.text_body.as_ref().unwrap();
        assert_eq!(body.properties.insets.left.as_emu(), Some(27432));
        assert_eq!(body.text(), "State-Owned Enterprise");
        let run = &body.paragraphs[0].runs[0];
        assert_eq!(run.bold, Some(false));
        assert_eq!(run.font_size.map(TextSize::get), Some(900));
        // All four fixture text boxes anchor with two-cell anchors.
        assert!(
            all.iter()
                .flat_map(|sheet| sheet.objects.iter())
                .all(|anchored| matches!(anchored.anchor, Anchor::TwoCell { .. }))
        );
    }
}
