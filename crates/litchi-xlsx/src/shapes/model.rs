//! Typed semantic SpreadsheetDrawing shape models.
//!
//! The model is context-first: anchors own objects, groups own direct children,
//! and unknown markup remains inert and bounded rather than being discarded.

use std::fmt;
use std::str::FromStr;

use crate::error::{Error, Result, invalid};
use litchi_drawingml::geom::Preset;
use litchi_drawingml::geometry::CustomGeometry;
pub use litchi_drawingml::text::body::{Body, Insets, Paragraph, Properties, Run};
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

pub(crate) const MAX_DRAWING_PART_BYTES: usize = 32 * 1024 * 1024;
pub(crate) const MAX_WORKBOOK_BYTES: usize = 32 * 1024 * 1024;
pub(crate) const MAX_DRAWINGS_PER_WORKSHEET: usize = 64;
pub(crate) const MAX_ANCHORS_PER_DRAWING: usize = 100_000;
pub(crate) const MAX_OBJECTS_PER_DRAWING: usize = 100_000;
pub(crate) const MAX_GROUP_DEPTH: usize = 32;
pub(crate) const MAX_XML_DEPTH: usize = 256;
pub(crate) const MAX_TEXT_BYTES: usize = 1024 * 1024;

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
    /// Unknown attributes and child elements retained from the source object.
    pub opaque: Opaque,
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

/// One XML attribute retained because the shape owner does not interpret it.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct UnknownAttribute {
    name: Box<str>,
    value: Box<str>,
}

impl UnknownAttribute {
    pub fn name(&self) -> &str {
        &self.name
    }

    pub fn value(&self) -> &str {
        &self.value
    }

    pub(crate) fn new(name: String, value: String) -> Result<Self> {
        if name.is_empty() || name.len() > MAX_DRAWING_PART_BYTES {
            return Err(invalid("unknown drawing attribute name is out of bounds"));
        }
        if value.len() > MAX_DRAWING_PART_BYTES {
            return Err(invalid("unknown drawing attribute value is out of bounds"));
        }
        Ok(Self {
            name: name.into_boxed_str(),
            value: value.into_boxed_str(),
        })
    }
}

/// A bounded XML element retained without interpreting its semantics.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct UnknownElement {
    xml: Box<[u8]>,
}

impl UnknownElement {
    pub fn new(xml: impl Into<Vec<u8>>) -> Result<Self> {
        Self::from_xml(xml.into())
    }

    pub fn as_xml(&self) -> &[u8] {
        &self.xml
    }

    pub(crate) fn from_xml(xml: Vec<u8>) -> Result<Self> {
        if xml.is_empty() || xml.len() > MAX_DRAWING_PART_BYTES {
            return Err(invalid(
                "unknown drawing element bytes exceed the drawing limit",
            ));
        }
        Ok(Self {
            xml: xml.into_boxed_slice(),
        })
    }
}

/// Unknown shape-owned markup retained as inert, bounded data.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Opaque {
    attributes: Vec<UnknownAttribute>,
    elements: Vec<UnknownElement>,
}

impl Opaque {
    pub fn attributes(&self) -> &[UnknownAttribute] {
        &self.attributes
    }

    pub fn elements(&self) -> &[UnknownElement] {
        &self.elements
    }

    pub(crate) fn push_attribute(&mut self, value: UnknownAttribute) -> Result<()> {
        if self.attributes.len() >= MAX_OBJECTS_PER_DRAWING {
            return Err(invalid(
                "unknown drawing attributes exceed the drawing limit",
            ));
        }
        self.attributes.push(value);
        Ok(())
    }

    pub(crate) fn push_element(&mut self, value: UnknownElement) -> Result<()> {
        if self.elements.len() >= MAX_OBJECTS_PER_DRAWING {
            return Err(invalid("unknown drawing elements exceed the drawing limit"));
        }
        self.elements.push(value);
        Ok(())
    }
}

/// An unknown drawing object preserved as its bounded XML fragment.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Unknown {
    element: UnknownElement,
}

impl Unknown {
    pub fn as_xml(&self) -> &[u8] {
        self.element.as_xml()
    }

    pub fn element(&self) -> &UnknownElement {
        &self.element
    }

    pub(crate) fn from_element(element: UnknownElement) -> Self {
        Self { element }
    }
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
    /// An object not modeled by this owner, retained as bounded XML.
    Unknown(Unknown),
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
