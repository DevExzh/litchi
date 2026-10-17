#![allow(
    clippy::wildcard_enum_match_arm,
    reason = "the drawing scanner reads a handful of XML events and skips every other one"
)]

//! Worksheet pictures and shapes of an XLSB package: read, add, move and
//! remove (unified-ops record 0322).
//!
//! An XLSB worksheet's Drawings part is the same SpreadsheetML drawing XML
//! as XLSX's ([MS-XLSB] section 2.1.7.23); the worksheet names it with a
//! `BrtDrawing` record. [`crate::Workbook::drawing_elements`] and
//! [`crate::SourceBackedWorksheet::drawing_elements`] list the top-level
//! objects of a worksheet's drawing in drawing order, each an [`Element`]
//! with its [`Placement`], names and the properties a picture or a preset
//! shape carries. [`Package::edit_drawings`] starts an [`Edit`]:
//! [`Edit::add_picture`] stages a media part, its relationship and an
//! `xdr:pic`; [`Edit::add_shape`] stages an `xdr:sp` with a preset geometry,
//! fill, outline and text; [`Edit::set`] changes the placement, names,
//! rotation and flips of an element in place; and [`Edit::remove`] deletes
//! an element's anchor with the parts only it refers to. The drawing part
//! (and the worksheet's `BrtDrawing`) is created with the first object.
//! [`Edit::commit`] returns a [`Commit`] with the new package, the
//! reversible [`Patch`] and the [`Diagnostics`];
//! [`Package::apply_drawings`] replays a patch. Every other anchor and
//! worksheet record keeps its bytes; an exact no-op keeps every byte, and a
//! changed edit of a signed workbook is refused with [`Error::Signed`].
//! Colors are `0xRRGGBB`.
//!
//! ```
//! use litchi_xlsb::Package;
//! use litchi_xlsb::drawing::edit::{ElementKind, Marker, Placement, ShapeSpec};
//!
//! # fn main() -> Result<(), litchi_xlsb::package::PackageError> {
//! let package = Package::new()?;
//! let mut edit = package.edit_drawings()?;
//! let placement = Placement::TwoCell {
//!     from: Marker::cell(0, 4),
//!     to: Marker::cell(2, 7),
//! };
//! let mut shape = ShapeSpec::new(placement, "Rectangle 1");
//! shape.text = Some("rect".into());
//! edit.add_shape(0, &shape)?;
//! let after = edit.commit()?.into_package();
//! let elements = after.workbook()?.drawing_elements(0)?;
//! assert_eq!(elements[0].kind, ElementKind::Shape);
//! assert_eq!(elements[0].text.as_deref(), Some("rect"));
//! # Ok(())
//! # }
//! ```

use std::fmt::Write as _;
use std::ops::Range;

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI, Relationships};
use quick_xml::events::{BytesStart, Event};

use crate::Workbook;
use crate::package::Package;
use crate::package::error::Result;
use crate::part_edit::{Rank, SheetLayout, Staged, family, invalid, record, splice};
use crate::raw::{Encoder, kind};

/// A `0xRRGGBB` color.
pub type Rgb = u32;

fn rgb_from_hex(value: &str) -> Option<Rgb> {
    (value.len() == 6)
        .then(|| u32::from_str_radix(value, 16).ok())
        .flatten()
}

family!("drawing");

/// The most top-level objects one worksheet drawing may carry.
pub const MAX_ELEMENTS: usize = 16_384;
/// The largest drawing part this family reads, in bytes.
pub const MAX_DRAWING_BYTES: usize = 64 * 1024 * 1024;
/// The largest image this family stages, in bytes.
pub const MAX_IMAGE_BYTES: usize = 256 * 1024 * 1024;

/// A bounded dimension of this family.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Top-level objects of one drawing ([`MAX_ELEMENTS`]).
    Elements,
    /// Bytes of one drawing part ([`MAX_DRAWING_BYTES`]).
    DrawingBytes,
    /// Bytes of one staged image ([`MAX_IMAGE_BYTES`]).
    ImageBytes,
}

impl LimitKind {
    const fn maximum(self) -> usize {
        match self {
            Self::Elements => MAX_ELEMENTS,
            Self::DrawingBytes => MAX_DRAWING_BYTES,
            Self::ImageBytes => MAX_IMAGE_BYTES,
        }
    }

    fn check(self, observed: usize) -> Result<()> {
        if observed > self.maximum() {
            return Err(invalid(format!(
                "worksheet drawing {} limit exceeded: {observed} exceeds {}",
                match self {
                    Self::Elements => "object",
                    Self::DrawingBytes => "drawing byte",
                    Self::ImageBytes => "image byte",
                },
                self.maximum()
            )));
        }
        Ok(())
    }
}

/// The width of a default column and the height of a default row, in EMU:
/// the geometry a cell anchor's `xfrm` is estimated from.
pub const COLUMN_EMU: i64 = 609_600;
/// See [`COLUMN_EMU`].
pub const ROW_EMU: i64 = 190_500;

/// A cell marker: a zero-based column and row with EMU offsets.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct Marker {
    /// Zero-based column.
    pub col: u32,
    /// Offset into the column, in EMU.
    pub col_offset: i64,
    /// Zero-based row.
    pub row: u32,
    /// Offset into the row, in EMU.
    pub row_offset: i64,
}

impl Marker {
    /// The top-left corner of a cell.
    #[must_use]
    pub const fn cell(col: u32, row: u32) -> Self {
        Self {
            col,
            col_offset: 0,
            row,
            row_offset: 0,
        }
    }

    fn emu(self) -> (i64, i64) {
        (
            i64::from(self.col) * COLUMN_EMU + self.col_offset,
            i64::from(self.row) * ROW_EMU + self.row_offset,
        )
    }
}

/// Where an object sits on the worksheet.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Placement {
    /// `twoCellAnchor`: from one cell marker to another.
    TwoCell {
        /// The top-left marker.
        from: Marker,
        /// The bottom-right marker.
        to: Marker,
    },
    /// `oneCellAnchor`: a marker and an extent in EMU.
    OneCell {
        /// The top-left marker.
        from: Marker,
        /// Width in EMU.
        cx: i64,
        /// Height in EMU.
        cy: i64,
    },
    /// `absoluteAnchor`: a position and an extent in EMU.
    Absolute {
        /// Left edge in EMU.
        x: i64,
        /// Top edge in EMU.
        y: i64,
        /// Width in EMU.
        cx: i64,
        /// Height in EMU.
        cy: i64,
    },
}

impl Placement {
    fn validate(self) -> Result<()> {
        let marker = |marker: Marker| {
            if marker.col >= 16_384 || marker.row >= 1_048_576 {
                return Err(invalid("a drawing marker exceeds the worksheet bounds"));
            }
            if marker.col_offset < 0 || marker.row_offset < 0 {
                return Err(invalid("a drawing marker offset cannot be negative"));
            }
            Ok(())
        };
        match self {
            Self::TwoCell { from, to } => {
                marker(from)?;
                marker(to)?;
                if (to.row, to.row_offset) < (from.row, from.row_offset)
                    || (to.col, to.col_offset) < (from.col, from.col_offset)
                {
                    return Err(invalid("a drawing anchor cannot be descending"));
                }
            },
            Self::OneCell { from, cx, cy } => {
                marker(from)?;
                if cx <= 0 || cy <= 0 {
                    return Err(invalid("a drawing extent must be positive"));
                }
            },
            Self::Absolute { x, y, cx, cy } => {
                if x < 0 || y < 0 || cx <= 0 || cy <= 0 {
                    return Err(invalid(
                        "an absolute drawing position must be non-negative and its extent positive",
                    ));
                }
            },
        }
        Ok(())
    }

    /// The estimated `xfrm` offset and extent in EMU.
    fn frame(self) -> (i64, i64, i64, i64) {
        match self {
            Self::TwoCell { from, to } => {
                let (x, y) = from.emu();
                let (right, bottom) = to.emu();
                (x, y, (right - x).max(0), (bottom - y).max(0))
            },
            Self::OneCell { from, cx, cy } => {
                let (x, y) = from.emu();
                (x, y, cx, cy)
            },
            Self::Absolute { x, y, cx, cy } => (x, y, cx, cy),
        }
    }

    fn xml(self, out: &mut String) {
        let marker = |out: &mut String, tag: &str, marker: Marker| {
            let _ = write!(
                out,
                "<xdr:{tag}><xdr:col>{}</xdr:col><xdr:colOff>{}</xdr:colOff><xdr:row>{}</xdr:row><xdr:rowOff>{}</xdr:rowOff></xdr:{tag}>",
                marker.col, marker.col_offset, marker.row, marker.row_offset
            );
        };
        match self {
            Self::TwoCell { from, to } => {
                marker(out, "from", from);
                marker(out, "to", to);
            },
            Self::OneCell { from, cx, cy } => {
                marker(out, "from", from);
                let _ = write!(out, "<xdr:ext cx=\"{cx}\" cy=\"{cy}\"/>");
            },
            Self::Absolute { x, y, cx, cy } => {
                let _ = write!(
                    out,
                    "<xdr:pos x=\"{x}\" y=\"{y}\"/><xdr:ext cx=\"{cx}\" cy=\"{cy}\"/>"
                );
            },
        }
    }

    const fn element(self) -> &'static str {
        match self {
            Self::TwoCell { .. } => "twoCellAnchor",
            Self::OneCell { .. } => "oneCellAnchor",
            Self::Absolute { .. } => "absoluteAnchor",
        }
    }
}

/// What a top-level drawing object is.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum ElementKind {
    /// `xdr:pic`.
    Picture,
    /// `xdr:sp`.
    Shape,
    /// `xdr:graphicFrame` (a chart or another graphic).
    Frame,
    /// `xdr:grpSp`.
    Group,
    /// `xdr:cxnSp`.
    Connector,
    /// Anything else (content parts, markup-compatibility branches).
    Other,
}

/// The image a picture shows.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Image {
    /// The media part.
    pub part: String,
    /// Its content type.
    pub content_type: String,
    /// Its size in bytes.
    pub size: u64,
}

/// A paragraph alignment of shape text.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Align {
    /// `l`.
    Left,
    /// `ctr`.
    Center,
    /// `r`.
    Right,
    /// `just`.
    Justify,
}

impl Align {
    const fn token(self) -> &'static str {
        match self {
            Self::Left => "l",
            Self::Center => "ctr",
            Self::Right => "r",
            Self::Justify => "just",
        }
    }

    fn parse(value: &str) -> Option<Self> {
        Some(match value {
            "l" => Self::Left,
            "ctr" => Self::Center,
            "r" => Self::Right,
            "just" => Self::Justify,
            _ => return None,
        })
    }
}

/// A vertical text anchor of a shape.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum VerticalAlign {
    /// `t`.
    Top,
    /// `ctr`.
    Middle,
    /// `b`.
    Bottom,
}

impl VerticalAlign {
    const fn token(self) -> &'static str {
        match self {
            Self::Top => "t",
            Self::Middle => "ctr",
            Self::Bottom => "b",
        }
    }

    fn parse(value: &str) -> Option<Self> {
        Some(match value {
            "t" => Self::Top,
            "ctr" => Self::Middle,
            "b" => Self::Bottom,
            _ => return None,
        })
    }
}

/// A shape fill.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Fill {
    /// `a:noFill`.
    None,
    /// A solid RGB fill.
    Solid(Rgb),
    /// A solid theme (scheme) color fill: one of [`SCHEME_COLORS`].
    Scheme(&'static str),
}

/// The `DrawingML` scheme color names (`ST_SchemeColorVal`).
pub const SCHEME_COLORS: &[&str] = &[
    "bg1", "tx1", "bg2", "tx2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
    "hlink", "folHlink", "phClr", "dk1", "lt1", "dk2", "lt2",
];

/// The `DrawingML` preset dash names (`ST_PresetLineDashVal`).
pub const DASHES: &[&str] = &[
    "solid",
    "dot",
    "dash",
    "lgDash",
    "dashDot",
    "lgDashDot",
    "lgDashDotDot",
    "sysDash",
    "sysDot",
    "sysDashDot",
    "sysDashDotDot",
];

/// The canonical spelling of a name in `names`, matched case-insensitively.
#[must_use]
pub fn canonical(names: &'static [&'static str], value: &str) -> Option<&'static str> {
    names
        .iter()
        .copied()
        .find(|name| name.eq_ignore_ascii_case(value))
}

/// A shape outline.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Line {
    /// `a:ln` with `a:noFill`.
    None,
    /// A solid line of a color and an optional width in EMU.
    Solid {
        /// The color.
        color: Rgb,
        /// The width in EMU.
        width: Option<i64>,
        /// A preset dash, one of [`DASHES`].
        dash: Option<&'static str>,
    },
}

/// The run formatting of a shape's text, read from its first run.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct Font {
    /// Bold.
    pub bold: Option<bool>,
    /// Italic.
    pub italic: Option<bool>,
    /// An underline token (`sng`, `dbl`).
    pub underline: Option<String>,
    /// The size in points.
    pub size: Option<f64>,
    /// The color.
    pub color: Option<Rgb>,
    /// The Latin typeface.
    pub name: Option<String>,
}

impl Font {
    fn is_empty(&self) -> bool {
        *self == Self::default()
    }
}

/// One top-level object of a worksheet drawing.
#[derive(Debug, Clone, PartialEq)]
#[non_exhaustive]
pub struct Element {
    /// What the object is.
    pub kind: ElementKind,
    /// Where it sits.
    pub placement: Option<Placement>,
    /// `cNvPr/@id`.
    pub id: Option<u32>,
    /// `cNvPr/@name`.
    pub name: Option<String>,
    /// `cNvPr/@descr`, the alternative text.
    pub description: Option<String>,
    /// `cNvPr/@title`.
    pub title: Option<String>,
    /// Whether the object is marked decorative.
    pub decorative: bool,
    /// The click hyperlink target.
    pub hyperlink: Option<String>,
    /// Clockwise rotation in degrees.
    pub rotation: Option<f64>,
    /// Horizontal flip.
    pub flip_h: bool,
    /// Vertical flip.
    pub flip_v: bool,
    /// The preset geometry of a shape.
    pub geometry: Option<String>,
    /// The fill of a shape.
    pub fill: Option<Fill>,
    /// The outline of a shape.
    pub line: Option<Line>,
    /// The text of a shape, paragraphs joined by `\n`.
    pub text: Option<String>,
    /// The text's run formatting.
    pub font: Font,
    /// The first paragraph's alignment.
    pub align: Option<Align>,
    /// The vertical text anchor.
    pub valign: Option<VerticalAlign>,
    /// The text insets (left, top, right, bottom) in EMU, when the shape
    /// sets any.
    pub margin: Option<[i64; 4]>,
    /// The image of a picture.
    pub image: Option<Image>,
    span: Range<usize>,
    embed: Option<String>,
    link: Option<String>,
}

impl Element {
    fn new(kind: ElementKind, span: Range<usize>) -> Self {
        Self {
            kind,
            placement: None,
            id: None,
            name: None,
            description: None,
            title: None,
            decorative: false,
            hyperlink: None,
            rotation: None,
            flip_h: false,
            flip_v: false,
            geometry: None,
            fill: None,
            line: None,
            text: None,
            font: Font::default(),
            align: None,
            valign: None,
            margin: None,
            image: None,
            span,
            embed: None,
            link: None,
        }
    }
}

/// A picture to add.
#[derive(Debug, Clone, PartialEq)]
pub struct PictureSpec {
    /// Where it sits.
    pub placement: Placement,
    /// `cNvPr/@name`.
    pub name: String,
    /// `cNvPr/@descr`.
    pub description: Option<String>,
    /// `cNvPr/@title`.
    pub title: Option<String>,
    /// The image bytes.
    pub image: Vec<u8>,
    /// The image content type (`image/png`).
    pub content_type: String,
    /// The media part extension (`png`).
    pub extension: String,
    /// Clockwise rotation in degrees.
    pub rotation: Option<f64>,
    /// Horizontal flip.
    pub flip_h: bool,
    /// Vertical flip.
    pub flip_v: bool,
}

/// A preset shape to add.
#[derive(Debug, Clone, PartialEq)]
pub struct ShapeSpec {
    /// Where it sits.
    pub placement: Placement,
    /// `cNvPr/@name`.
    pub name: String,
    /// `cNvPr/@descr`.
    pub description: Option<String>,
    /// The preset geometry token (`rect`, `roundRect`, `ellipse`).
    pub geometry: String,
    /// The fill.
    pub fill: Option<Fill>,
    /// The outline.
    pub line: Option<Line>,
    /// The text; `\n` separates paragraphs.
    pub text: Option<String>,
    /// The text's run formatting.
    pub font: Font,
    /// The paragraph alignment.
    pub align: Option<Align>,
    /// The vertical text anchor.
    pub valign: Option<VerticalAlign>,
    /// The text insets (left, top, right, bottom) in EMU.
    pub margin: Option<[i64; 4]>,
    /// Clockwise rotation in degrees.
    pub rotation: Option<f64>,
    /// Horizontal flip.
    pub flip_h: bool,
    /// Vertical flip.
    pub flip_v: bool,
}

impl ShapeSpec {
    /// A rectangle named `name` at `placement`, with nothing else set.
    #[must_use]
    pub fn new(placement: Placement, name: impl Into<String>) -> Self {
        Self {
            placement,
            name: name.into(),
            description: None,
            geometry: "rect".to_owned(),
            fill: None,
            line: None,
            text: None,
            font: Font::default(),
            align: None,
            valign: None,
            margin: None,
            rotation: None,
            flip_h: false,
            flip_v: false,
        }
    }
}

/// The in-place changes [`Edit::set`] makes; `None` keeps a value.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct Change {
    /// A new placement; the anchor kind may change.
    pub placement: Option<Placement>,
    /// A new `cNvPr/@name`.
    pub name: Option<String>,
    /// A new `cNvPr/@descr`; `Some(None)` removes it.
    pub description: Option<Option<String>>,
    /// A new rotation in degrees.
    pub rotation: Option<f64>,
    /// New flips.
    pub flip_h: Option<bool>,
    /// See [`Change::flip_h`].
    pub flip_v: Option<bool>,
}

const DRAWING_ROOT: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
    "\n",
    r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" "#,
    r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"></xdr:wsDr>"#
);
const RELATIONSHIP_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

fn drawing_in(relationships: &Relationships) -> Result<Option<PackURI>> {
    let mut found = None;
    for relationship in relationships.iter() {
        if relationship.reltype() != rt::DRAWING && relationship.reltype() != rt::STRICT_DRAWING {
            continue;
        }
        if relationship.is_external() || found.is_some() {
            return Err(invalid(
                "a worksheet needs one internal drawing relationship",
            ));
        }
        found = Some(relationship.target_partname()?);
    }
    Ok(found)
}

fn attribute(element: &BytesStart<'_>, key: &[u8]) -> Option<String> {
    litchi_ooxml_common::xml::attributes::first_wins(element)
        .flatten()
        .find_map(|attribute| {
            (attribute.key.local_name().as_ref() == key).then(|| {
                let raw = String::from_utf8_lossy(&attribute.value).into_owned();
                quick_xml::escape::unescape(&raw).map_or(raw.clone(), std::borrow::Cow::into_owned)
            })
        })
}

fn relationship_attribute(element: &BytesStart<'_>, key: &[u8]) -> Option<String> {
    litchi_ooxml_common::xml::attributes::first_wins(element)
        .flatten()
        .find_map(|attribute| {
            (attribute.key.local_name().as_ref() == key && attribute.key.prefix().is_some())
                .then(|| String::from_utf8_lossy(&attribute.value).into_owned())
        })
}

fn flag(value: Option<String>) -> bool {
    value.is_some_and(|value| value == "1" || value == "true")
}

#[derive(Default)]
struct MarkerScan {
    values: [i64; 4],
}

/// Scan a drawing part's top-level anchors. Names are matched by local name
/// under the `SpreadsheetDrawing` and `DrawingML` vocabularies.
fn scan(xml: &[u8]) -> Result<(Vec<Element>, usize, u32)> {
    LimitKind::DrawingBytes.check(xml.len())?;
    let mut reader = quick_xml::Reader::from_reader(xml);
    let mut stack: Vec<Vec<u8>> = Vec::new();
    let mut elements = Vec::new();
    let mut current: Option<Element> = None;
    let mut anchor_start = 0;
    let mut anchor_kind = Vec::new();
    let mut markers: [MarkerScan; 2] = Default::default();
    let mut marker_slot = None::<usize>;
    let mut marker_field = None::<usize>;
    let mut extent = None::<(i64, i64)>;
    let mut position = None::<(i64, i64)>;
    let mut max_id = 0_u32;
    let mut root_close = None;
    let mut paragraphs: Vec<String> = Vec::new();
    let mut in_text = false;
    let mut first_run = true;
    loop {
        let start = usize::try_from(reader.buffer_position()).unwrap_or(usize::MAX);
        let event = reader
            .read_event()
            .map_err(|error| invalid(format!("drawing markup: {error}")))?;
        let end = usize::try_from(reader.buffer_position()).unwrap_or(usize::MAX);
        let (element, empty) = match &event {
            Event::Start(element) => (Some(element.clone()), false),
            Event::Empty(element) => (Some(element.clone()), true),
            _ => (None, false),
        };
        if let Some(element) = element {
            let local = element.local_name().as_ref().to_vec();
            let depth = stack.len();
            if depth == 1 {
                anchor_start = start;
                anchor_kind = local.clone();
                markers = Default::default();
                extent = None;
                position = None;
                if empty {
                    elements.push(Element::new(ElementKind::Other, start..end));
                }
            }
            if depth == 2 {
                match local.as_slice() {
                    b"from" => marker_slot = Some(0),
                    b"to" => marker_slot = Some(1),
                    b"ext" => {
                        extent = attribute(&element, b"cx")
                            .zip(attribute(&element, b"cy"))
                            .and_then(|(cx, cy)| Some((cx.parse().ok()?, cy.parse().ok()?)));
                    },
                    b"pos" => {
                        position = attribute(&element, b"x")
                            .zip(attribute(&element, b"y"))
                            .and_then(|(x, y)| Some((x.parse().ok()?, y.parse().ok()?)));
                    },
                    b"clientData" => {},
                    object => {
                        let kind = match object {
                            b"pic" => ElementKind::Picture,
                            b"sp" => ElementKind::Shape,
                            b"graphicFrame" => ElementKind::Frame,
                            b"grpSp" => ElementKind::Group,
                            b"cxnSp" => ElementKind::Connector,
                            _ => ElementKind::Other,
                        };
                        if current.is_none() {
                            current = Some(Element::new(kind, anchor_start..anchor_start));
                            paragraphs.clear();
                            first_run = true;
                        }
                    },
                }
            }
            if depth == 3 && marker_slot.is_some() {
                marker_field = match local.as_slice() {
                    b"col" => Some(0),
                    b"colOff" => Some(1),
                    b"row" => Some(2),
                    b"rowOff" => Some(3),
                    _ => None,
                };
            }
            if let Some(object) = current.as_mut()
                && depth >= 3
                // Only a group's own properties, not its children's.
                && (object.kind != ElementKind::Group || depth <= 4)
            {
                let own = true;
                match local.as_slice() {
                    b"cNvPr" if depth == 4 => {
                        object.id = attribute(&element, b"id").and_then(|id| id.parse().ok());
                        object.name = attribute(&element, b"name");
                        object.description = attribute(&element, b"descr");
                        object.title = attribute(&element, b"title");
                        if let Some(id) = object.id {
                            max_id = max_id.max(id);
                        }
                    },
                    b"hlinkClick" if depth == 5 => {
                        object.link = relationship_attribute(&element, b"id");
                    },
                    b"decorative" => {
                        object.decorative = attribute(&element, b"val").is_none_or(|v| v != "0");
                    },
                    b"xfrm" if depth == 4 && own => {
                        object.rotation = attribute(&element, b"rot")
                            .and_then(|rot| rot.parse::<f64>().ok())
                            .map(|rot| rot / 60_000.0);
                        object.flip_h = flag(attribute(&element, b"flipH"));
                        object.flip_v = flag(attribute(&element, b"flipV"));
                    },
                    b"prstGeom" if depth == 4 => {
                        object.geometry = attribute(&element, b"prst");
                    },
                    b"noFill" if depth == 4 && stack.get(3).is_some_and(|n| n == b"spPr") => {
                        object.fill = Some(Fill::None);
                    },
                    b"srgbClr" => {
                        let color = attribute(&element, b"val").and_then(|v| rgb_from_hex(&v));
                        let parent = stack.last().map(Vec::as_slice);
                        let grand = stack.len().checked_sub(2).and_then(|at| stack.get(at));
                        if let Some(color) = color {
                            match (grand.map(Vec::as_slice), parent, depth) {
                                (Some(b"spPr"), Some(b"solidFill"), 5) => {
                                    object.fill = Some(Fill::Solid(color));
                                },
                                (Some(b"ln"), Some(b"solidFill"), 6) => {
                                    if let Some(Line::Solid { color: slot, .. }) =
                                        object.line.as_mut()
                                    {
                                        *slot = color;
                                    }
                                },
                                (Some(b"rPr"), Some(b"solidFill"), _) if first_run => {
                                    object.font.color = Some(color);
                                },
                                _ => {},
                            }
                        }
                    },
                    b"schemeClr"
                        if depth == 5
                            && stack.get(3).is_some_and(|n| n == b"spPr")
                            && stack.last().is_some_and(|n| n == b"solidFill") =>
                    {
                        if let Some(name) = attribute(&element, b"val")
                            .and_then(|value| canonical(SCHEME_COLORS, &value))
                        {
                            object.fill = Some(Fill::Scheme(name));
                        }
                    },
                    b"prstDash" if depth == 5 && stack.last().is_some_and(|n| n == b"ln") => {
                        if let Some(Line::Solid { dash, .. }) = object.line.as_mut() {
                            *dash = attribute(&element, b"val")
                                .and_then(|value| canonical(DASHES, &value));
                        }
                    },
                    b"ln" if depth == 4 => {
                        let width = attribute(&element, b"w").and_then(|w| w.parse().ok());
                        object.line = Some(Line::Solid {
                            color: 0,
                            width,
                            dash: None,
                        });
                    },
                    b"noFill" if depth == 5 && stack.last().is_some_and(|n| n == b"ln") => {
                        object.line = Some(Line::None);
                    },
                    b"blip" => object.embed = relationship_attribute(&element, b"embed"),
                    b"bodyPr" => {
                        object.valign = attribute(&element, b"anchor")
                            .and_then(|value| VerticalAlign::parse(&value));
                        let insets = [b"lIns".as_slice(), b"tIns", b"rIns", b"bIns"].map(|key| {
                            attribute(&element, key).and_then(|v| v.parse::<i64>().ok())
                        });
                        if insets.iter().any(Option::is_some) {
                            // The schema defaults: 0.1in left and right, 0.05in
                            // top and bottom.
                            let defaults = [91_440, 45_720, 91_440, 45_720];
                            object.margin = Some(
                                [0, 1, 2, 3].map(|side| insets[side].unwrap_or(defaults[side])),
                            );
                        }
                    },
                    b"p" if stack.last().is_some_and(|n| n == b"txBody") => {
                        paragraphs.push(String::new());
                    },
                    b"pPr" if paragraphs.len() == 1 => {
                        object.align = attribute(&element, b"algn").and_then(|v| Align::parse(&v));
                    },
                    b"rPr" if first_run => {
                        object.font.bold =
                            attribute(&element, b"b").map(|v| v == "1" || v == "true");
                        object.font.italic =
                            attribute(&element, b"i").map(|v| v == "1" || v == "true");
                        object.font.underline = attribute(&element, b"u").filter(|v| v != "none");
                        object.font.size = attribute(&element, b"sz")
                            .and_then(|v| v.parse::<f64>().ok())
                            .map(|v| v / 100.0);
                    },
                    b"latin" if first_run && stack.last().is_some_and(|n| n == b"rPr") => {
                        object.font.name = attribute(&element, b"typeface");
                    },
                    b"t" if !empty => in_text = true,
                    _ => {},
                }
            }
            if !empty {
                stack.push(local);
            }
            continue;
        }
        match event {
            Event::Text(text) if in_text => {
                let raw = String::from_utf8_lossy(text.as_ref()).into_owned();
                let value = quick_xml::escape::unescape(&raw)
                    .map_or(raw.clone(), std::borrow::Cow::into_owned);
                if let Some(paragraph) = paragraphs.last_mut() {
                    paragraph.push_str(&value);
                }
            },
            Event::Text(text) if marker_field.is_some() => {
                if let (Some(slot), Some(field)) = (marker_slot, marker_field) {
                    let value = String::from_utf8_lossy(text.as_ref())
                        .trim()
                        .parse()
                        .unwrap_or(0);
                    markers[slot].values[field] = value;
                }
            },
            Event::GeneralRef(reference) if in_text => {
                if let (Some(paragraph), Some(value)) =
                    (paragraphs.last_mut(), entity(reference.as_ref()))
                {
                    paragraph.push_str(&value);
                }
            },
            Event::End(_) => {
                let local = stack.pop().unwrap_or_default();
                let depth = stack.len();
                if local == b"t" {
                    in_text = false;
                }
                if local == b"r" && current.is_some() {
                    first_run = false;
                }
                if depth == 3 {
                    marker_field = None;
                }
                if depth == 2 && (local == b"from" || local == b"to") {
                    marker_slot = None;
                }
                if depth == 1 {
                    let mut object = current
                        .take()
                        .unwrap_or_else(|| Element::new(ElementKind::Other, 0..0));
                    object.span = anchor_start..end;
                    let marker = |scan: &MarkerScan| Marker {
                        col: u32::try_from(scan.values[0]).unwrap_or(0),
                        col_offset: scan.values[1],
                        row: u32::try_from(scan.values[2]).unwrap_or(0),
                        row_offset: scan.values[3],
                    };
                    object.placement = match anchor_kind.as_slice() {
                        b"twoCellAnchor" => Some(Placement::TwoCell {
                            from: marker(&markers[0]),
                            to: marker(&markers[1]),
                        }),
                        b"oneCellAnchor" => extent.map(|(cx, cy)| Placement::OneCell {
                            from: marker(&markers[0]),
                            cx,
                            cy,
                        }),
                        b"absoluteAnchor" => position
                            .zip(extent)
                            .map(|((x, y), (cx, cy))| Placement::Absolute { x, y, cx, cy }),
                        _ => None,
                    };
                    if object.kind == ElementKind::Shape && !paragraphs.is_empty() {
                        let text = paragraphs.join("\n");
                        object.text = Some(text);
                    }
                    elements.push(object);
                    LimitKind::Elements.check(elements.len())?;
                }
                if depth == 0 {
                    root_close = Some(start);
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    let root_close = root_close.ok_or_else(|| invalid("the drawing part has no closed root"))?;
    Ok((elements, root_close, max_id))
}

/// Resolve the image and hyperlink relationships of scanned elements.
fn resolve(
    elements: &mut [Element],
    relationships: &Relationships,
    image: impl Fn(&PackURI) -> Option<(String, u64)>,
) {
    for element in elements {
        if let Some(id) = element.embed.as_deref()
            && let Some(relationship) = relationships.get(id)
            && !relationship.is_external()
            && let Ok(target) = relationship.target_partname()
            && let Some((content_type, size)) = image(&target)
        {
            element.image = Some(Image {
                part: target.as_str().to_owned(),
                content_type,
                size,
            });
        }
        if let Some(id) = element.link.as_deref()
            && let Some(relationship) = relationships.get(id)
        {
            element.hyperlink = Some(relationship.target_ref().to_owned());
        }
    }
}

/// The elements of a worksheet part's drawing in an eager package.
pub(crate) fn eager_elements(package: &OpcPackage, worksheet: &PackURI) -> Result<Vec<Element>> {
    let Some(drawing) = drawing_in(package.get_part(worksheet)?.rels())? else {
        return Ok(Vec::new());
    };
    let part = package.get_part(&drawing)?;
    let (mut elements, _, _) = scan(part.blob())?;
    resolve(&mut elements, part.rels(), |target| {
        let image = package.get_part(target).ok()?;
        Some((image.content_type().to_owned(), image.blob().len() as u64))
    });
    Ok(elements)
}

/// The elements of a worksheet part's drawing in a source-backed package:
/// only the worksheet's relationships, its drawing part and the image parts'
/// entries are read.
pub(crate) fn source_elements(
    package: &litchi_opc::SourceBackedPackage,
    worksheet: &PackURI,
) -> Result<Vec<Element>> {
    let Some(drawing) = drawing_in(package.part(worksheet)?.rels())? else {
        return Ok(Vec::new());
    };
    let part = package.part(&drawing)?;
    let declared = part.declared_uncompressed_size()?;
    LimitKind::DrawingBytes.check(usize::try_from(declared).unwrap_or(usize::MAX))?;
    let data = part.data()?;
    let (mut elements, _, _) = scan(data.as_bytes())?;
    resolve(&mut elements, part.rels(), |target| {
        let image = package.part(target).ok()?;
        Some((
            image.content_type().to_owned(),
            image.declared_uncompressed_size().ok()?,
        ))
    });
    Ok(elements)
}

/// A chart a worksheet drawing hosts: its frame's placement and name, and
/// the chart part's DrawingML (`c:chartSpace`) markup.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub struct ChartPart {
    /// Where the chart's frame sits.
    pub placement: Option<Placement>,
    /// The frame's `cNvPr/@name`.
    pub name: Option<String>,
    /// The chart part's markup.
    pub xml: Vec<u8>,
}

/// The chart relationship a frame names, among `relationships`.
fn chart_target(fragment: &[u8], relationships: &Relationships) -> Option<PackURI> {
    relationship_ids(fragment).into_iter().find_map(|id| {
        let relationship = relationships.get(&id)?;
        (!relationship.is_external()
            && matches!(relationship.reltype(), rt::CHART | rt::STRICT_CHART))
        .then(|| relationship.target_partname().ok())
        .flatten()
    })
}

/// The charts of a worksheet part's drawing in an eager package, in drawing
/// order.
pub(crate) fn eager_charts(package: &OpcPackage, worksheet: &PackURI) -> Result<Vec<ChartPart>> {
    let Some(drawing) = drawing_in(package.get_part(worksheet)?.rels())? else {
        return Ok(Vec::new());
    };
    let part = package.get_part(&drawing)?;
    let (elements, _, _) = scan(part.blob())?;
    let mut charts = Vec::new();
    for element in elements
        .iter()
        .filter(|element| element.kind == ElementKind::Frame)
    {
        let Some(target) = chart_target(&part.blob()[element.span.clone()], part.rels()) else {
            continue;
        };
        charts.push(ChartPart {
            placement: element.placement,
            name: element.name.clone(),
            xml: package.get_part(&target)?.blob().to_vec(),
        });
    }
    Ok(charts)
}

/// The charts of a worksheet part's drawing in a source-backed package.
pub(crate) fn source_charts(
    package: &litchi_opc::SourceBackedPackage,
    worksheet: &PackURI,
) -> Result<Vec<ChartPart>> {
    let Some(drawing) = drawing_in(package.part(worksheet)?.rels())? else {
        return Ok(Vec::new());
    };
    let part = package.part(&drawing)?;
    let declared = part.declared_uncompressed_size()?;
    LimitKind::DrawingBytes.check(usize::try_from(declared).unwrap_or(usize::MAX))?;
    let data = part.data()?;
    let (elements, _, _) = scan(data.as_bytes())?;
    let mut charts = Vec::new();
    for element in elements
        .iter()
        .filter(|element| element.kind == ElementKind::Frame)
    {
        let Some(target) = chart_target(&data.as_bytes()[element.span.clone()], part.rels()) else {
            continue;
        };
        let chart = package.part(&target)?;
        let declared = chart.declared_uncompressed_size()?;
        LimitKind::DrawingBytes.check(usize::try_from(declared).unwrap_or(usize::MAX))?;
        charts.push(ChartPart {
            placement: element.placement,
            name: element.name.clone(),
            xml: chart.data()?.as_bytes().to_vec(),
        });
    }
    Ok(charts)
}

impl Package {
    /// Start a worksheet drawing edit.
    ///
    /// # Errors
    ///
    /// Returns the workbook materialization errors of this package.
    pub fn edit_drawings(&self) -> Result<Edit<'_>> {
        Ok(Edit {
            staged: Staged::of_package(self)?,
        })
    }

    /// Apply a drawing patch to a new package; this one is unchanged. An
    /// empty patch returns this package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state and [`Error::Signed`] for a changed patch of a signed
    /// workbook.
    pub fn apply_drawings(&self, patch: &Patch) -> Result<Self> {
        patch.0.apply(self)
    }
}

impl Workbook {
    /// Start a worksheet drawing edit of this workbook without first copying
    /// it into a package.
    pub fn edit_drawings(&self) -> Edit<'_> {
        Edit {
            staged: Staged::of_workbook(self),
        }
    }
}

/// A staged worksheet drawing edit. Dropping it discards it.
#[derive(Debug)]
pub struct Edit<'a> {
    staged: Staged<'a>,
}

fn escape(value: &str) -> String {
    quick_xml::escape::escape(value).into_owned()
}

fn check_text(value: &str, what: &str) -> Result<()> {
    if value.contains('\0') || value.len() > 1024 * 1024 {
        return Err(invalid(format!("a drawing {what} is not valid text")));
    }
    Ok(())
}

fn rotation(degrees: Option<f64>) -> Result<Option<i64>> {
    degrees
        .map(|degrees| {
            if !degrees.is_finite() {
                return Err(invalid("a rotation must be a finite number of degrees"));
            }
            let normalized = degrees.rem_euclid(360.0);
            #[allow(
                clippy::cast_possible_truncation,
                reason = "a normalized rotation is below 21,600,000"
            )]
            Ok((normalized * 60_000.0).round() as i64)
        })
        .transpose()
        .map(|value| value.filter(|value| *value != 0))
}

fn xfrm(out: &mut String, placement: Placement, rotation: Option<i64>, flip_h: bool, flip_v: bool) {
    let (x, y, cx, cy) = placement.frame();
    out.push_str("<a:xfrm");
    if let Some(rotation) = rotation {
        let _ = write!(out, " rot=\"{rotation}\"");
    }
    if flip_h {
        out.push_str(" flipH=\"1\"");
    }
    if flip_v {
        out.push_str(" flipV=\"1\"");
    }
    let _ = write!(
        out,
        "><a:off x=\"{x}\" y=\"{y}\"/><a:ext cx=\"{cx}\" cy=\"{cy}\"/></a:xfrm>"
    );
}

fn name_attributes(
    out: &mut String,
    id: u32,
    name: &str,
    description: Option<&str>,
    title: Option<&str>,
) {
    let _ = write!(out, "<xdr:cNvPr id=\"{id}\" name=\"{}\"", escape(name));
    if let Some(description) = description {
        let _ = write!(out, " descr=\"{}\"", escape(description));
    }
    if let Some(title) = title {
        let _ = write!(out, " title=\"{}\"", escape(title));
    }
    out.push_str("/>");
}

fn anchor_open(out: &mut String, placement: Placement, edit_as: Option<&str>) {
    let _ = write!(
        out,
        "<xdr:{} xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"",
        placement.element()
    );
    if let (Placement::TwoCell { .. }, Some(edit_as)) = (placement, edit_as) {
        let _ = write!(out, " editAs=\"{edit_as}\"");
    }
    out.push('>');
    placement.xml(out);
}

fn anchor_close(out: &mut String, placement: Placement) {
    let _ = write!(out, "<xdr:clientData/></xdr:{}>", placement.element());
}

fn shape_xml(spec: &ShapeSpec, id: u32) -> Result<String> {
    spec.placement.validate()?;
    check_text(&spec.name, "name")?;
    if spec.geometry.is_empty() || !spec.geometry.bytes().all(|b| b.is_ascii_alphanumeric()) {
        return Err(invalid("a preset geometry is an alphanumeric token"));
    }
    let rotation = rotation(spec.rotation)?;
    let mut out = String::new();
    anchor_open(&mut out, spec.placement, None);
    out.push_str("<xdr:sp macro=\"\" textlink=\"\"><xdr:nvSpPr>");
    name_attributes(&mut out, id, &spec.name, spec.description.as_deref(), None);
    out.push_str("<xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr>");
    xfrm(&mut out, spec.placement, rotation, spec.flip_h, spec.flip_v);
    let _ = write!(
        out,
        "<a:prstGeom prst=\"{}\"><a:avLst/></a:prstGeom>",
        spec.geometry
    );
    match spec.fill {
        Some(Fill::None) => out.push_str("<a:noFill/>"),
        Some(Fill::Solid(color)) => {
            let _ = write!(
                out,
                "<a:solidFill><a:srgbClr val=\"{}\"/></a:solidFill>",
                rgb(color)
            );
        },
        Some(Fill::Scheme(name)) => {
            if canonical(SCHEME_COLORS, name).is_none() {
                return Err(invalid("a scheme fill names a DrawingML scheme color"));
            }
            let _ = write!(
                out,
                "<a:solidFill><a:schemeClr val=\"{name}\"/></a:solidFill>"
            );
        },
        None => {},
    }
    match spec.line {
        Some(Line::None) => out.push_str("<a:ln><a:noFill/></a:ln>"),
        Some(Line::Solid { color, width, dash }) => {
            out.push_str("<a:ln");
            if let Some(width) = width {
                if !(0..=20_116_800).contains(&width) {
                    return Err(invalid("a line width must be 0 to 1584 points"));
                }
                let _ = write!(out, " w=\"{width}\"");
            }
            let _ = write!(
                out,
                "><a:solidFill><a:srgbClr val=\"{}\"/></a:solidFill>",
                rgb(color)
            );
            if let Some(dash) = dash {
                if canonical(DASHES, dash).is_none() {
                    return Err(invalid("a line dash names a DrawingML preset dash"));
                }
                let _ = write!(out, "<a:prstDash val=\"{dash}\"/>");
            }
            out.push_str("</a:ln>");
        },
        None => {},
    }
    out.push_str("</xdr:spPr>");
    if spec.text.is_some()
        || !spec.font.is_empty()
        || spec.align.is_some()
        || spec.valign.is_some()
        || spec.margin.is_some()
    {
        out.push_str(
            "<xdr:txBody><a:bodyPr vertOverflow=\"clip\" horzOverflow=\"clip\" rtlCol=\"0\"",
        );
        if let Some(valign) = spec.valign {
            let _ = write!(out, " anchor=\"{}\"", valign.token());
        }
        if let Some(margin) = spec.margin {
            if margin.iter().any(|inset| !(0..=51_206_400).contains(inset)) {
                return Err(invalid("a text margin must be 0 to 4032 points"));
            }
            let [left, top, right, bottom] = margin;
            let _ = write!(
                out,
                " lIns=\"{left}\" tIns=\"{top}\" rIns=\"{right}\" bIns=\"{bottom}\""
            );
        }
        out.push_str("/><a:lstStyle/>");
        let text = spec.text.as_deref().unwrap_or("");
        check_text(text, "text")?;
        for paragraph in text.split('\n') {
            out.push_str("<a:p>");
            if let Some(align) = spec.align {
                let _ = write!(out, "<a:pPr algn=\"{}\"/>", align.token());
            }
            let mut properties = String::from("<a:rPr lang=\"en-US\"");
            if let Some(size) = spec.font.size {
                if !size.is_finite() || !(1.0..=4000.0).contains(&size) {
                    return Err(invalid("a font size must be 1 to 4000 points"));
                }
                #[allow(clippy::cast_possible_truncation, reason = "bounded above")]
                let _ = write!(properties, " sz=\"{}\"", (size * 100.0).round() as i64);
            }
            if let Some(bold) = spec.font.bold {
                let _ = write!(properties, " b=\"{}\"", u8::from(bold));
            }
            if let Some(italic) = spec.font.italic {
                let _ = write!(properties, " i=\"{}\"", u8::from(italic));
            }
            if let Some(underline) = &spec.font.underline {
                if !underline.bytes().all(|b| b.is_ascii_alphanumeric()) {
                    return Err(invalid("an underline is an alphanumeric token"));
                }
                let _ = write!(properties, " u=\"{underline}\"");
            }
            properties.push('>');
            if let Some(color) = spec.font.color {
                let _ = write!(
                    properties,
                    "<a:solidFill><a:srgbClr val=\"{}\"/></a:solidFill>",
                    rgb(color)
                );
            }
            if let Some(name) = &spec.font.name {
                check_text(name, "font name")?;
                let _ = write!(properties, "<a:latin typeface=\"{}\"/>", escape(name));
            }
            properties.push_str("</a:rPr>");
            if paragraph.is_empty() {
                let _ = write!(
                    out,
                    "<a:endParaRPr{}",
                    &properties["<a:rPr".len()..].replace("</a:rPr>", "</a:endParaRPr>")
                );
            } else {
                let _ = write!(
                    out,
                    "<a:r>{properties}<a:t>{}</a:t></a:r>",
                    escape(paragraph)
                );
            }
            out.push_str("</a:p>");
        }
        out.push_str("</xdr:txBody>");
    }
    out.push_str("</xdr:sp>");
    anchor_close(&mut out, spec.placement);
    Ok(out)
}

fn rgb(color: Rgb) -> String {
    format!("{:06X}", color & 0x00FF_FFFF)
}

fn picture_xml(spec: &PictureSpec, id: u32, relationship_id: &str) -> Result<String> {
    spec.placement.validate()?;
    check_text(&spec.name, "name")?;
    let rotation = rotation(spec.rotation)?;
    let mut out = String::new();
    anchor_open(&mut out, spec.placement, Some("oneCell"));
    out.push_str("<xdr:pic><xdr:nvPicPr>");
    name_attributes(
        &mut out,
        id,
        &spec.name,
        spec.description.as_deref(),
        spec.title.as_deref(),
    );
    let _ = write!(
        out,
        "<xdr:cNvPicPr><a:picLocks noChangeAspect=\"1\"/></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip xmlns:r=\"{RELATIONSHIP_NS}\" r:embed=\"{}\"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr>",
        escape(relationship_id)
    );
    xfrm(&mut out, spec.placement, rotation, spec.flip_h, spec.flip_v);
    out.push_str("<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>");
    anchor_close(&mut out, spec.placement);
    Ok(out)
}

impl Edit<'_> {
    /// The staged elements of the worksheet at `sheet`.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing sheet, a chartsheet, or a malformed
    /// drawing graph.
    pub fn elements(&self, sheet: usize) -> Result<Vec<Element>> {
        let worksheet = self.staged.worksheet(sheet)?;
        eager_elements(self.staged.opc(), &worksheet)
    }

    /// The worksheet's drawing part, created (with its relationship and the
    /// worksheet's `BrtDrawing` record) when it has none.
    fn drawing(&mut self, sheet: usize) -> Result<PackURI> {
        let worksheet = self.staged.worksheet(sheet)?;
        if let Some(drawing) = drawing_in(self.staged.opc().get_part(&worksheet)?.rels())? {
            return Ok(drawing);
        }
        let drawing = self.staged.free_name("/xl/drawings/drawing", ".xml")?;
        self.staged.add_part(BlobPart::new(
            drawing.clone(),
            ct::OFC_DRAWING.to_owned(),
            DRAWING_ROOT.as_bytes().to_vec(),
        ))?;
        let id = self
            .staged
            .opc_mut()?
            .get_part_mut(&worksheet)?
            .relate_to(&drawing.relative_ref(worksheet.base_uri()), rt::DRAWING);
        let bytes = self.staged.blob(&worksheet)?;
        let layout = SheetLayout::scan(bytes)?;
        if layout.units_of(kind::DRAWING).next().is_some() {
            return Err(invalid(
                "the worksheet names a drawing it has no relationship to",
            ));
        }
        let at = layout.insertion(Rank::Drawing);
        let mut payload = Vec::new();
        Encoder::new(&mut payload).write_wide_string(&id)?;
        let updated = splice(bytes, at..at, &record(kind::DRAWING, &payload)?);
        self.staged.replace_worksheet(&worksheet, updated)?;
        self.staged.touch(&worksheet);
        Ok(drawing)
    }

    fn append(&mut self, drawing: &PackURI, fragment: &str) -> Result<()> {
        let (_, root_close, _) = scan(self.staged.blob(drawing)?)?;
        let mut xml = self.staged.blob(drawing)?.to_vec();
        xml.splice(root_close..root_close, fragment.bytes());
        LimitKind::DrawingBytes.check(xml.len())?;
        self.staged.replace_blob(drawing, xml)?;
        Ok(())
    }

    /// Add a picture; returns its index among the worksheet's pictures.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing sheet or a chartsheet, an invalid
    /// placement, name or content type, a full drawing, or a signed workbook.
    pub fn add_picture(&mut self, sheet: usize, spec: &PictureSpec) -> Result<usize> {
        LimitKind::ImageBytes.check(spec.image.len())?;
        if !spec.content_type.starts_with("image/")
            || spec.extension.is_empty()
            || !spec.extension.bytes().all(|b| b.is_ascii_alphanumeric())
        {
            return Err(invalid(
                "a picture needs an image content type and extension",
            ));
        }
        spec.placement.validate()?;
        let drawing = self.drawing(sheet)?;
        let (_, _, max_id) = scan(self.staged.blob(&drawing)?)?;
        let media = self
            .staged
            .free_name("/xl/media/image", &format!(".{}", spec.extension))?;
        self.staged.add_part(BlobPart::new(
            media.clone(),
            spec.content_type.clone(),
            spec.image.clone(),
        ))?;
        let relationship_id = self
            .staged
            .opc_mut()?
            .get_part_mut(&drawing)?
            .relate_to(&media.relative_ref(drawing.base_uri()), rt::IMAGE);
        self.staged.touch(&drawing);
        let fragment = picture_xml(spec, max_id.saturating_add(1).max(2), &relationship_id)?;
        self.append(&drawing, &fragment)?;
        let count = self
            .elements(sheet)?
            .iter()
            .filter(|element| element.kind == ElementKind::Picture)
            .count();
        Ok(count - 1)
    }

    /// Add a preset shape; returns its index among the worksheet's shapes.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing sheet or a chartsheet, an invalid
    /// placement, geometry, text or font, a full drawing, or a signed
    /// workbook.
    pub fn add_shape(&mut self, sheet: usize, spec: &ShapeSpec) -> Result<usize> {
        spec.placement.validate()?;
        let drawing = self.drawing(sheet)?;
        let (_, _, max_id) = scan(self.staged.blob(&drawing)?)?;
        let fragment = shape_xml(spec, max_id.saturating_add(1).max(2))?;
        self.append(&drawing, &fragment)?;
        let count = self
            .elements(sheet)?
            .iter()
            .filter(|element| element.kind == ElementKind::Shape)
            .count();
        Ok(count - 1)
    }

    fn select(&self, sheet: usize, kind: ElementKind, index: usize) -> Result<(PackURI, Element)> {
        let worksheet = self.staged.worksheet(sheet)?;
        let drawing = drawing_in(self.staged.opc().get_part(&worksheet)?.rels())?
            .ok_or_else(|| invalid("the worksheet has no drawing"))?;
        let (elements, _, _) = scan(self.staged.blob(&drawing)?)?;
        let element = elements
            .into_iter()
            .filter(|element| element.kind == kind)
            .nth(index)
            .ok_or_else(|| invalid(format!("drawing element {index} does not exist")))?;
        Ok((drawing, element))
    }

    /// Change the placement, names, rotation or flips of the element at
    /// `index` among the worksheet's elements of `kind`. Returns whether
    /// anything changed; everything else of the element keeps its bytes.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing element, an invalid change, markup this
    /// editor cannot rewrite, or a changed edit of a signed workbook.
    pub fn set(
        &mut self,
        sheet: usize,
        kind: ElementKind,
        index: usize,
        change: &Change,
    ) -> Result<bool> {
        let (drawing, element) = self.select(sheet, kind, index)?;
        let xml = self.staged.blob(&drawing)?;
        let fragment = String::from_utf8(xml[element.span.clone()].to_vec())
            .map_err(|_overflow| invalid("drawing markup is not UTF-8"))?;
        let mut updated = fragment.clone();
        if let Some(placement) = change.placement {
            placement.validate()?;
            updated = replace_placement(&updated, placement)?;
        }
        if let Some(name) = &change.name {
            check_text(name, "name")?;
            updated = set_tag_attribute(&updated, "cNvPr", "name", Some(&escape(name)))?;
        }
        if let Some(description) = &change.description {
            if let Some(value) = description {
                check_text(value, "description")?;
            }
            updated = set_tag_attribute(
                &updated,
                "cNvPr",
                "descr",
                description.as_deref().map(escape).as_deref(),
            )?;
        }
        if change.rotation.is_some() || change.flip_h.is_some() || change.flip_v.is_some() {
            if !updated.contains(":xfrm") && !updated.contains("<xfrm") {
                return Err(invalid("the element has no transform to change"));
            }
            if let Some(degrees) = change.rotation {
                let value = rotation(Some(degrees))?.map(|value| value.to_string());
                updated = set_tag_attribute(&updated, "xfrm", "rot", value.as_deref())?;
            }
            for (key, value) in [("flipH", change.flip_h), ("flipV", change.flip_v)] {
                if let Some(value) = value {
                    updated = set_tag_attribute(&updated, "xfrm", key, value.then_some("1"))?;
                }
            }
        }
        if updated == fragment {
            return Ok(false);
        }
        let mut out = xml.to_vec();
        out.splice(element.span.clone(), updated.into_bytes());
        if scan(&out)?.0.len() != scan(xml)?.0.len() {
            return Err(invalid("the changed drawing element did not read back"));
        }
        self.staged.replace_blob(&drawing, out)
    }

    /// Remove the element at `index` among the worksheet's elements of
    /// `kind`, with its relationships and the parts only it refers to.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing element or a signed workbook.
    pub fn remove(&mut self, sheet: usize, kind: ElementKind, index: usize) -> Result<()> {
        let (drawing, element) = self.select(sheet, kind, index)?;
        let mut xml = self.staged.blob(&drawing)?.to_vec();
        let fragment = xml[element.span.clone()].to_vec();
        xml.drain(element.span.clone());
        let ids = relationship_ids(&fragment);
        let rest = String::from_utf8_lossy(&xml).into_owned();
        let opc = self.staged.opc_mut()?;
        let part = opc.get_part_mut(&drawing)?;
        part.set_blob(xml);
        let mut targets = Vec::new();
        for id in ids {
            if rest.contains(&format!("\"{id}\"")) {
                continue;
            }
            if let Some(relationship) = part.rels().get(&id)
                && !relationship.is_external()
                && let Ok(target) = relationship.target_partname()
            {
                targets.push(target);
            }
            part.rels_mut().remove(&id);
        }
        self.staged.touch(&drawing);
        for target in targets {
            let opc = self.staged.opc_mut()?;
            let referenced = opc.iter_parts().any(|part| {
                part.rels().iter().any(|relationship| {
                    !relationship.is_external()
                        && relationship.target_partname().ok().as_ref() == Some(&target)
                })
            });
            if !referenced {
                self.staged.remove_part(&target)?;
            }
        }
        Ok(())
    }

    /// Validate the staged package and publish it.
    ///
    /// # Errors
    ///
    /// Returns the workbook validation errors of the staged package.
    pub fn commit(self) -> Result<Commit> {
        Ok(Commit::new(self.staged.commit()?))
    }
}

/// The relationship ids an anchor fragment names (`r:embed`, `r:id`,
/// `r:link`, ...): every prefixed attribute whose value looks like one.
fn relationship_ids(fragment: &[u8]) -> Vec<String> {
    let mut ids = Vec::new();
    let mut reader = quick_xml::Reader::from_reader(fragment);
    while let Ok(event) = reader.read_event() {
        match event {
            Event::Start(element) | Event::Empty(element) => {
                for attribute in
                    litchi_ooxml_common::xml::attributes::first_wins(&element).flatten()
                {
                    if attribute
                        .key
                        .prefix()
                        .is_some_and(|prefix| prefix.as_ref() == b"r")
                    {
                        ids.push(String::from_utf8_lossy(&attribute.value).into_owned());
                    }
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    ids
}

/// Replace an anchor fragment's element name and placement children.
fn replace_placement(fragment: &str, placement: Placement) -> Result<String> {
    let mut reader = quick_xml::Reader::from_reader(fragment.as_bytes());
    let mut depth = 0usize;
    let mut prefix = String::new();
    let mut open_end = None;
    let mut object_start = None;
    let mut root_close = None;
    loop {
        let start = usize::try_from(reader.buffer_position()).unwrap_or(usize::MAX);
        let event = reader
            .read_event()
            .map_err(|error| invalid(format!("drawing markup: {error}")))?;
        let end = usize::try_from(reader.buffer_position()).unwrap_or(usize::MAX);
        match event {
            Event::Start(element) => {
                if depth == 0 {
                    let name = String::from_utf8_lossy(element.name().as_ref()).into_owned();
                    prefix = name
                        .split_once(':')
                        .map_or(String::new(), |(p, _)| format!("{p}:"));
                    open_end = Some(end);
                } else if depth == 1 {
                    let local = element.local_name();
                    if !matches!(local.as_ref(), b"from" | b"to" | b"ext" | b"pos")
                        && object_start.is_none()
                    {
                        object_start = Some(start);
                    }
                }
                depth += 1;
            },
            Event::Empty(element) => {
                if depth == 1 {
                    let local = element.local_name();
                    if !matches!(local.as_ref(), b"from" | b"to" | b"ext" | b"pos")
                        && object_start.is_none()
                    {
                        object_start = Some(start);
                    }
                }
            },
            Event::End(_) => {
                depth = depth.saturating_sub(1);
                if depth == 0 {
                    root_close = Some(start);
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    let (Some(open_end), Some(object_start), Some(root_close)) =
        (open_end, object_start, root_close)
    else {
        return Err(invalid("the drawing anchor cannot be rewritten"));
    };
    // Keep the root's attributes except a two-cell `editAs`.
    let open = &fragment[..open_end];
    let attributes = open
        .find(char::is_whitespace)
        .map_or("", |at| &open[at..open.len() - 1]);
    let attributes = if matches!(placement, Placement::TwoCell { .. }) {
        attributes.to_owned()
    } else {
        remove_attribute(attributes, "editAs")
    };
    let mut markers = String::new();
    placement.xml(&mut markers);
    if !prefix.is_empty() && prefix != "xdr:" {
        markers = markers.replace("xdr:", &prefix);
    } else if prefix.is_empty() {
        markers = markers.replace("xdr:", "");
    }
    let element = format!("{prefix}{}", placement.element());
    Ok(format!(
        "<{element}{attributes}>{markers}{}</{element}>",
        &fragment[object_start..root_close]
    ))
}

fn remove_attribute(attributes: &str, name: &str) -> String {
    let needle = format!(" {name}=\"");
    match attributes.find(&needle) {
        Some(at) => {
            let value_start = at + needle.len();
            let value_end = attributes[value_start..]
                .find('"')
                .map_or(attributes.len(), |offset| value_start + offset + 1);
            format!("{}{}", &attributes[..at], &attributes[value_end..])
        },
        None => attributes.to_owned(),
    }
}

/// Set (`Some`) or remove (`None`) an attribute of the first element with
/// local name `local` in a fragment.
fn set_tag_attribute(
    fragment: &str,
    local: &str,
    name: &str,
    value: Option<&str>,
) -> Result<String> {
    let mut at = None;
    for pattern in [
        format!(":{local} "),
        format!(":{local}>"),
        format!(":{local}/"),
        format!("<{local} "),
    ] {
        if let Some(found) = fragment.find(&pattern) {
            let start = fragment[..=found].rfind('<').unwrap_or(found);
            at = Some(at.map_or(start, |current: usize| current.min(start)));
        }
    }
    let start = at.ok_or_else(|| invalid(format!("the element has no {local}")))?;
    let end = fragment[start..]
        .find('>')
        .map(|offset| start + offset)
        .ok_or_else(|| invalid("unterminated start tag"))?;
    let tag = &fragment[start..end];
    let (body, self_closing) = match tag.strip_suffix('/') {
        Some(body) => (body, "/"),
        None => (tag, ""),
    };
    let mut body = remove_attribute(body, name);
    if let Some(value) = value {
        let _ = write!(body, " {name}=\"{value}\"");
    }
    Ok(format!(
        "{}{body}{self_closing}{}",
        &fragment[..start],
        &fragment[end..]
    ))
}

/// The text of a predefined or character entity reference (`amp`, `#x41`).
fn entity(name: &[u8]) -> Option<String> {
    let name = std::str::from_utf8(name).ok()?;
    Some(match name {
        "amp" => "&".to_owned(),
        "lt" => "<".to_owned(),
        "gt" => ">".to_owned(),
        "apos" => "'".to_owned(),
        "quot" => "\"".to_owned(),
        _ => {
            let code = name.strip_prefix('#')?;
            let value = match code.strip_prefix('x') {
                Some(hex) => u32::from_str_radix(hex, 16).ok()?,
                None => code.parse().ok()?,
            };
            char::from_u32(value)?.to_string()
        },
    })
}

#[cfg(test)]
mod tests {
    use super::{Change, ElementKind, Fill, Font, Marker, PictureSpec, Placement, ShapeSpec};
    use crate::Package;

    const PNG: &[u8] = b"\x89PNG\r\n\x1a\n\0\0\0\rIHDR\0\0\0\x01\0\0\0\x01\x08\x06\0\0\0\x1f\x15\xc4\x89\0\0\0\rIDATx\x9cc\xf8\x0f\0\0\x01\x01\0\x05\x18\xd8\xa8\x00\0\0\0IEND\xaeB`\x82";

    fn placement() -> Placement {
        Placement::TwoCell {
            from: Marker::cell(1, 1),
            to: Marker::cell(4, 6),
        }
    }

    #[test]
    fn shapes_and_pictures_round_trip() {
        let package = Package::new().unwrap();
        let mut edit = package.edit_drawings().unwrap();
        let mut shape = ShapeSpec::new(placement(), "Shape 1");
        shape.geometry = "ellipse".into();
        shape.text = Some("one\ntwo & three".into());
        shape.fill = Some(Fill::Solid(0x4472C4));
        shape.font = Font {
            bold: Some(true),
            color: Some(0xFFFFFF),
            size: Some(14.0),
            ..Font::default()
        };
        shape.rotation = Some(30.0);
        shape.flip_h = true;
        assert_eq!(edit.add_shape(0, &shape).unwrap(), 0);
        let picture = PictureSpec {
            placement: Placement::OneCell {
                from: Marker::cell(6, 2),
                cx: 914_400,
                cy: 914_400,
            },
            name: "Picture 1".into(),
            description: Some("logo".into()),
            title: None,
            image: PNG.to_vec(),
            content_type: "image/png".into(),
            extension: "png".into(),
            rotation: None,
            flip_h: false,
            flip_v: false,
        };
        assert_eq!(edit.add_picture(0, &picture).unwrap(), 0);
        let (after, patch) = edit.commit().unwrap().into_parts();
        let reopened = Package::from_bytes(&after.to_bytes().unwrap()).unwrap();
        let elements = reopened.workbook().unwrap().drawing_elements(0).unwrap();
        assert_eq!(elements.len(), 2);
        let shape = &elements[0];
        assert_eq!(shape.kind, ElementKind::Shape);
        assert_eq!(shape.geometry.as_deref(), Some("ellipse"));
        assert_eq!(shape.text.as_deref(), Some("one\ntwo & three"));
        assert_eq!(shape.rotation, Some(30.0));
        assert!(shape.flip_h);
        assert_eq!(shape.font.bold, Some(true));
        assert_eq!(shape.font.size, Some(14.0));
        assert_eq!(shape.font.color, Some(0xFFFFFF));
        assert_eq!(shape.placement, Some(placement()));
        let picture = &elements[1];
        assert_eq!(picture.kind, ElementKind::Picture);
        assert_eq!(picture.description.as_deref(), Some("logo"));
        assert_eq!(picture.image.as_ref().unwrap().content_type, "image/png");
        assert_eq!(
            package.apply_drawings(&patch).unwrap().to_bytes().unwrap(),
            after.to_bytes().unwrap()
        );

        let mut edit = after.edit_drawings().unwrap();
        let moved = Placement::TwoCell {
            from: Marker::cell(0, 0),
            to: Marker::cell(2, 2),
        };
        assert!(
            edit.set(
                0,
                ElementKind::Picture,
                0,
                &Change {
                    placement: Some(moved),
                    description: Some(None),
                    rotation: Some(90.0),
                    ..Change::default()
                }
            )
            .unwrap()
        );
        let elements = edit.elements(0).unwrap();
        assert_eq!(elements[1].placement, Some(moved));
        assert_eq!(elements[1].description, None);
        assert_eq!(elements[1].rotation, Some(90.0));
        edit.remove(0, ElementKind::Picture, 0).unwrap();
        edit.remove(0, ElementKind::Shape, 0).unwrap();
        let empty = edit.commit().unwrap().into_package();
        assert!(
            empty
                .workbook()
                .unwrap()
                .drawing_elements(0)
                .unwrap()
                .is_empty()
        );
    }
}
