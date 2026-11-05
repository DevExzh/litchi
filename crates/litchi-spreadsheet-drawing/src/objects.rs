//! Worksheet drawing objects: the format-neutral half of the XLSX and XLSB
//! drawing editors (unified-ops records 0320, 0322 and 0325).
//!
//! An XLSB worksheet's Drawings part is the same `SpreadsheetDrawing` XML as
//! an XLSX one, so both crates' `drawing::edit` families read and write it
//! through this module: [`read`] lists a drawing part's top-level anchors as
//! [`Element`]s, [`shape_fragment`] and [`picture_fragment`] author new
//! anchors, [`append`] adds one to a part, and [`change`] and [`detach`]
//! rewrite or drop one in place, keeping every other byte. Staging the
//! parts and relationships stays with each format's `Edit`.
//!
//! The model is generic over the [`Color`] each format uses for an RGB value
//! (`litchi_xlsx::Rgb`, or `0xRRGGBB` as a `u32` in XLSB).

use std::fmt::{self, Write as _};
use std::ops::Range;

use litchi_opc::constants::relationship_type as rt;
use litchi_opc::{PackURI, Relationships};
use quick_xml::events::{BytesStart, Event};

use crate::{Error, LimitKind, Result};

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

/// An RGB color of a drawing's fills, lines and text.
pub trait Color: Copy + PartialEq + fmt::Debug {
    /// Black, the color a line without a solid fill reads as.
    const BLACK: Self;

    /// Parse a `DrawingML` `srgbClr/@val`; `None` when it is not a color.
    fn from_hex(value: &str) -> Option<Self>;

    /// The six uppercase hexadecimal digits `srgbClr/@val` takes.
    fn hex(self) -> String;
}

/// `0xRRGGBB`.
impl Color for u32 {
    const BLACK: Self = 0;

    fn from_hex(value: &str) -> Option<Self> {
        (value.len() == 6)
            .then(|| Self::from_str_radix(value, 16).ok())
            .flatten()
    }

    fn hex(self) -> String {
        format!("{:06X}", self & 0x00FF_FFFF)
    }
}

/// The most top-level objects one worksheet drawing may carry.
pub const MAX_ELEMENTS: usize = 16_384;
/// The largest drawing part read or written, in bytes.
pub const MAX_DRAWING_BYTES: usize = 64 * 1024 * 1024;
/// The largest image a picture may carry, in bytes.
pub const MAX_IMAGE_BYTES: usize = 256 * 1024 * 1024;

/// The caller-selected bounds for one variable SpreadsheetDrawing XML
/// operation. Every raw read and rewrite takes this policy explicitly.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[must_use]
pub struct Limits {
    /// Maximum top-level objects retained by a scan.
    pub max_elements: usize,
    /// Maximum input or output drawing bytes.
    pub max_drawing_bytes: usize,
    /// Maximum image bytes admitted by picture validation.
    pub max_image_bytes: usize,
}

impl Limits {
    /// The established finite worksheet-drawing policy.
    pub const DEFAULT: Self = Self {
        max_elements: MAX_ELEMENTS,
        max_drawing_bytes: MAX_DRAWING_BYTES,
        max_image_bytes: MAX_IMAGE_BYTES,
    };

    /// Construct a caller-selected policy. [`Self::validate`] is run at each
    /// raw boundary so malformed policies cannot bypass the hard ceilings.
    #[must_use = "use the selected worksheet-drawing limits"]
    pub const fn new(
        max_elements: usize,
        max_drawing_bytes: usize,
        max_image_bytes: usize,
    ) -> Self {
        Self {
            max_elements,
            max_drawing_bytes,
            max_image_bytes,
        }
    }

    /// Derive a worksheet-drawing policy from the retained OPC source policy.
    /// The shared hard ceilings are never widened by an owner policy.
    #[must_use = "use the limits derived from the retained source policy"]
    pub fn from_source_limits(source: litchi_opc::SourceLimits) -> Self {
        Self {
            max_elements: source.max_events().clamp(1, MAX_ELEMENTS),
            max_drawing_bytes: usize::try_from(source.max_part_bytes())
                .unwrap_or(usize::MAX)
                .clamp(1, MAX_DRAWING_BYTES),
            max_image_bytes: usize::try_from(source.max_part_bytes())
                .unwrap_or(usize::MAX)
                .clamp(1, MAX_IMAGE_BYTES),
        }
    }

    /// Validate the policy against this module's hard ceilings.
    pub fn validate(self) -> Result<Self> {
        for (kind, value) in [
            (LimitKind::Elements, self.max_elements),
            (LimitKind::DrawingBytes, self.max_drawing_bytes),
            (LimitKind::ImageBytes, self.max_image_bytes),
        ] {
            let maximum = kind.maximum();
            if value == 0 || value > maximum {
                return Err(Error::invalid_limit(kind, value, maximum));
            }
        }
        Ok(self)
    }

    fn check(self, kind: LimitKind, observed: usize, maximum: usize) -> Result<()> {
        if observed > maximum {
            return Err(Error::limit(kind, observed, maximum));
        }
        Ok(())
    }

    fn drawing_bytes(self, observed: usize) -> Result<()> {
        self.check(LimitKind::DrawingBytes, observed, self.max_drawing_bytes)
    }

    fn elements(self, observed: usize) -> Result<()> {
        self.check(LimitKind::Elements, observed, self.max_elements)
    }

    fn image_bytes(self, observed: usize) -> Result<()> {
        self.check(LimitKind::ImageBytes, observed, self.max_image_bytes)
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self::DEFAULT
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
    /// Check that the markers lie in the worksheet, offsets are not negative,
    /// a two-cell anchor is not descending and extents are positive.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Invalid`] for the first violated bound.
    pub fn validate(self) -> Result<()> {
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

    fn xml<W: fmt::Write>(self, out: &mut W) {
        let marker = |out: &mut W, tag: &str, marker: Marker| {
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
pub enum Fill<C> {
    /// `a:noFill`.
    None,
    /// A solid RGB fill.
    Solid(C),
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
pub enum Line<C> {
    /// `a:ln` with `a:noFill`.
    None,
    /// A solid line of a color and an optional width in EMU.
    Solid {
        /// The color.
        color: C,
        /// The width in EMU.
        width: Option<i64>,
        /// A preset dash, one of [`DASHES`].
        dash: Option<&'static str>,
    },
}

/// The run formatting of a shape's text, read from its first run.
#[derive(Debug, Clone, PartialEq)]
pub struct Font<C> {
    /// Bold.
    pub bold: Option<bool>,
    /// Italic.
    pub italic: Option<bool>,
    /// An underline token (`sng`, `dbl`).
    pub underline: Option<String>,
    /// The size in points.
    pub size: Option<f64>,
    /// The color.
    pub color: Option<C>,
    /// The Latin typeface.
    pub name: Option<String>,
}

impl<C> Default for Font<C> {
    fn default() -> Self {
        Self {
            bold: None,
            italic: None,
            underline: None,
            size: None,
            color: None,
            name: None,
        }
    }
}

impl<C> Font<C> {
    fn is_empty(&self) -> bool {
        self.bold.is_none()
            && self.italic.is_none()
            && self.underline.is_none()
            && self.size.is_none()
            && self.color.is_none()
            && self.name.is_none()
    }
}

/// One top-level object of a worksheet drawing.
#[derive(Debug, Clone, PartialEq)]
#[non_exhaustive]
pub struct Element<C> {
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
    pub fill: Option<Fill<C>>,
    /// The outline of a shape.
    pub line: Option<Line<C>>,
    /// The text of a shape, paragraphs joined by `\n`.
    pub text: Option<String>,
    /// The text's run formatting.
    pub font: Font<C>,
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

impl<C> Element<C> {
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
pub struct ShapeSpec<C> {
    /// Where it sits.
    pub placement: Placement,
    /// `cNvPr/@name`.
    pub name: String,
    /// `cNvPr/@descr`.
    pub description: Option<String>,
    /// The preset geometry token (`rect`, `roundRect`, `ellipse`).
    pub geometry: String,
    /// The fill.
    pub fill: Option<Fill<C>>,
    /// The outline.
    pub line: Option<Line<C>>,
    /// The text; `\n` separates paragraphs.
    pub text: Option<String>,
    /// The text's run formatting.
    pub font: Font<C>,
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

impl<C> ShapeSpec<C> {
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

/// The in-place changes [`change`] makes; `None` keeps a value.
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

/// The markup of a new, empty drawing part.
pub const EMPTY_DRAWING: &str = concat!(
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
fn scan<C: Color>(xml: &[u8], limits: Limits) -> Result<(Vec<Element<C>>, usize, u32)> {
    limits.drawing_bytes(xml.len())?;
    let mut reader = quick_xml::Reader::from_reader(xml);
    let mut stack: Vec<Vec<u8>> = Vec::new();
    let mut elements = Vec::new();
    let mut current: Option<Element<C>> = None;
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
                        let color = attribute(&element, b"val").and_then(|v| C::from_hex(&v));
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
                            color: C::BLACK,
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
                    let next = elements.len().checked_add(1).ok_or_else(|| {
                        Error::limit(LimitKind::Elements, usize::MAX, limits.max_elements)
                    })?;
                    limits.elements(next)?;
                    elements.push(object);
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

/// Count the top-level XML elements a fragment would add to a drawing root.
///
/// The fragment is intentionally scanned without retaining events or source
/// text. This keeps the element policy enforceable before append allocates its
/// rewritten drawing.
fn fragment_elements(fragment: &str) -> Result<usize> {
    let mut reader = quick_xml::Reader::from_reader(fragment.as_bytes());
    let mut depth = 0usize;
    let mut count = 0usize;
    loop {
        match reader
            .read_event()
            .map_err(|error| invalid(format!("drawing fragment: {error}")))?
        {
            Event::Start(_) => {
                if depth == 0 {
                    count = count
                        .checked_add(1)
                        .ok_or_else(|| invalid("drawing fragment element count overflow"))?;
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("drawing fragment nesting overflow"))?;
            },
            Event::Empty(_) if depth == 0 => {
                count = count
                    .checked_add(1)
                    .ok_or_else(|| invalid("drawing fragment element count overflow"))?;
            },
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("drawing fragment closes outside an element"))?;
            },
            Event::Eof => {
                if depth != 0 {
                    return Err(invalid("drawing fragment has an unclosed element"));
                }
                return Ok(count);
            },
            _ => {},
        }
    }
}

/// Resolve the image and hyperlink relationships of scanned elements.
fn resolve<C>(
    elements: &mut [Element<C>],
    relationships: &Relationships,
    image: impl Fn(&PackURI) -> Option<(String, u64)>,
    limits: Limits,
) -> Result<()> {
    for element in elements {
        if let Some(id) = element.embed.as_deref()
            && let Some(relationship) = relationships.get(id)
            && !relationship.is_external()
            && let Ok(target) = relationship.target_partname()
            && let Some((content_type, size)) = image(&target)
        {
            let image_size = usize::try_from(size).unwrap_or(usize::MAX);
            limits.image_bytes(image_size)?;
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
    Ok(())
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

/// A formatting sink used to measure a fragment before its output buffer is
/// allocated. It deliberately retains no source text.
#[derive(Default)]
struct LengthWriter {
    length: usize,
}

impl fmt::Write for LengthWriter {
    fn write_str(&mut self, value: &str) -> fmt::Result {
        self.length = self.length.saturating_add(value.len());
        Ok(())
    }
}

fn put<W: fmt::Write>(out: &mut W, value: &str) -> Result<()> {
    out.write_str(value)
        .map_err(|_| invalid("drawing output could not be formatted"))
}

fn escaped<W: fmt::Write>(out: &mut W, value: &str) -> Result<()> {
    let mut start = 0;
    for (index, byte) in value.bytes().enumerate() {
        let replacement = match byte {
            b'<' => Some("&lt;"),
            b'>' => Some("&gt;"),
            b'&' => Some("&amp;"),
            b'\'' => Some("&apos;"),
            b'\"' => Some("&quot;"),
            _ => None,
        };
        let Some(replacement) = replacement else {
            continue;
        };
        put(out, &value[start..index])?;
        put(out, replacement)?;
        start = index + 1;
    }
    put(out, &value[start..])
}

fn xfrm<W: fmt::Write>(
    out: &mut W,
    placement: Placement,
    rotation: Option<i64>,
    flip_h: bool,
    flip_v: bool,
) {
    let (x, y, cx, cy) = placement.frame();
    let _ = out.write_str("<a:xfrm");
    if let Some(rotation) = rotation {
        let _ = write!(out, " rot=\"{rotation}\"");
    }
    if flip_h {
        let _ = out.write_str(" flipH=\"1\"");
    }
    if flip_v {
        let _ = out.write_str(" flipV=\"1\"");
    }
    let _ = write!(
        out,
        "><a:off x=\"{x}\" y=\"{y}\"/><a:ext cx=\"{cx}\" cy=\"{cy}\"/></a:xfrm>"
    );
}

fn name_attributes<W: fmt::Write>(
    out: &mut W,
    id: u32,
    name: &str,
    description: Option<&str>,
    title: Option<&str>,
) -> Result<()> {
    let _ = write!(out, "<xdr:cNvPr id=\"{id}\" name=\"");
    escaped(out, name)?;
    put(out, "\"")?;
    if let Some(description) = description {
        put(out, " descr=\"")?;
        escaped(out, description)?;
        put(out, "\"")?;
    }
    if let Some(title) = title {
        put(out, " title=\"")?;
        escaped(out, title)?;
        put(out, "\"")?;
    }
    put(out, "/>")?;
    Ok(())
}

fn anchor_open<W: fmt::Write>(out: &mut W, placement: Placement, edit_as: Option<&str>) {
    let _ = write!(
        out,
        "<xdr:{} xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"",
        placement.element()
    );
    if let (Placement::TwoCell { .. }, Some(edit_as)) = (placement, edit_as) {
        let _ = write!(out, " editAs=\"{edit_as}\"");
    }
    let _ = out.write_char('>');
    placement.xml(out);
}

fn anchor_close<W: fmt::Write>(out: &mut W, placement: Placement) {
    let _ = write!(out, "<xdr:clientData/></xdr:{}>", placement.element());
}

fn write_font_properties<C: Color, W: fmt::Write>(
    out: &mut W,
    font: &Font<C>,
    end_paragraph: bool,
) -> Result<()> {
    let tag = if end_paragraph {
        "a:endParaRPr"
    } else {
        "a:rPr"
    };
    let _ = write!(out, "<{tag} lang=\"en-US\"");
    if let Some(size) = font.size {
        if !size.is_finite() || !(1.0..=4000.0).contains(&size) {
            return Err(invalid("a font size must be 1 to 4000 points"));
        }
        #[allow(clippy::cast_possible_truncation, reason = "bounded above")]
        let _ = write!(out, " sz=\"{}\"", (size * 100.0).round() as i64);
    }
    if let Some(bold) = font.bold {
        let _ = write!(out, " b=\"{}\"", u8::from(bold));
    }
    if let Some(italic) = font.italic {
        let _ = write!(out, " i=\"{}\"", u8::from(italic));
    }
    if let Some(underline) = &font.underline {
        if !underline.bytes().all(|b| b.is_ascii_alphanumeric()) {
            return Err(invalid("an underline is an alphanumeric token"));
        }
        let _ = write!(out, " u=\"{underline}\"");
    }
    put(out, ">")?;
    if let Some(color) = font.color {
        let _ = write!(
            out,
            "<a:solidFill><a:srgbClr val=\"{}\"/></a:solidFill>",
            color.hex()
        );
    }
    if let Some(name) = &font.name {
        check_text(name, "font name")?;
        put(out, "<a:latin typeface=\"")?;
        escaped(out, name)?;
        put(out, "\"/>")?;
    }
    let _ = write!(out, "</{tag}>");
    Ok(())
}

fn write_shape_xml<C: Color, W: fmt::Write>(
    out: &mut W,
    spec: &ShapeSpec<C>,
    id: u32,
) -> Result<()> {
    spec.placement.validate()?;
    check_text(&spec.name, "name")?;
    if spec.geometry.is_empty() || !spec.geometry.bytes().all(|b| b.is_ascii_alphanumeric()) {
        return Err(invalid("a preset geometry is an alphanumeric token"));
    }
    let rotation = rotation(spec.rotation)?;
    anchor_open(out, spec.placement, None);
    put(out, r#"<xdr:sp macro="" textlink=""><xdr:nvSpPr>"#)?;
    name_attributes(out, id, &spec.name, spec.description.as_deref(), None)?;
    put(out, r#"<xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr>"#)?;
    xfrm(out, spec.placement, rotation, spec.flip_h, spec.flip_v);
    let _ = write!(
        out,
        "<a:prstGeom prst=\"{}\"><a:avLst/></a:prstGeom>",
        spec.geometry
    );
    match spec.fill {
        Some(Fill::None) => put(out, "<a:noFill/>")?,
        Some(Fill::Solid(color)) => {
            let _ = write!(
                out,
                "<a:solidFill><a:srgbClr val=\"{}\"/></a:solidFill>",
                color.hex()
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
        Some(Line::None) => put(out, "<a:ln><a:noFill/></a:ln>")?,
        Some(Line::Solid { color, width, dash }) => {
            put(out, "<a:ln")?;
            if let Some(width) = width {
                if !(0..=20_116_800).contains(&width) {
                    return Err(invalid("a line width must be 0 to 1584 points"));
                }
                let _ = write!(out, " w=\"{width}\"");
            }
            let _ = write!(
                out,
                "><a:solidFill><a:srgbClr val=\"{}\"/></a:solidFill>",
                color.hex()
            );
            if let Some(dash) = dash {
                if canonical(DASHES, dash).is_none() {
                    return Err(invalid("a line dash names a DrawingML preset dash"));
                }
                let _ = write!(out, "<a:prstDash val=\"{dash}\"/>");
            }
            put(out, "</a:ln>")?;
        },
        None => {},
    }
    put(out, "</xdr:spPr>")?;
    if spec.text.is_some()
        || !spec.font.is_empty()
        || spec.align.is_some()
        || spec.valign.is_some()
        || spec.margin.is_some()
    {
        put(
            out,
            r#"<xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip" rtlCol="0""#,
        )?;
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
        put(out, "/><a:lstStyle/>")?;
        let text = spec.text.as_deref().unwrap_or("");
        check_text(text, "text")?;
        for paragraph in text.split('\n') {
            put(out, "<a:p>")?;
            if let Some(align) = spec.align {
                let _ = write!(out, "<a:pPr algn=\"{}\"/>", align.token());
            }
            if paragraph.is_empty() {
                write_font_properties(out, &spec.font, true)?;
            } else {
                put(out, "<a:r>")?;
                write_font_properties(out, &spec.font, false)?;
                put(out, "<a:t>")?;
                escaped(out, paragraph)?;
                put(out, "</a:t></a:r>")?;
            }
            put(out, "</a:p>")?;
        }
        put(out, "</xdr:txBody>")?;
    }
    put(out, "</xdr:sp>")?;
    anchor_close(out, spec.placement);
    Ok(())
}

fn write_picture_xml<W: fmt::Write>(
    out: &mut W,
    spec: &PictureSpec,
    id: u32,
    relationship_id: &str,
) -> Result<()> {
    spec.placement.validate()?;
    check_text(&spec.name, "name")?;
    let rotation = rotation(spec.rotation)?;
    anchor_open(out, spec.placement, Some("oneCell"));
    put(out, "<xdr:pic><xdr:nvPicPr>")?;
    name_attributes(
        out,
        id,
        &spec.name,
        spec.description.as_deref(),
        spec.title.as_deref(),
    )?;
    let _ = write!(
        out,
        "<xdr:cNvPicPr><a:picLocks noChangeAspect=\"1\"/></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip xmlns:r=\"{RELATIONSHIP_NS}\" r:embed=\""
    );
    escaped(out, relationship_id)?;
    put(
        out,
        "\"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr>",
    )?;
    xfrm(out, spec.placement, rotation, spec.flip_h, spec.flip_v);
    put(
        out,
        "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>",
    )?;
    anchor_close(out, spec.placement);
    Ok(())
}

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
fn replace_placement(fragment: &str, placement: Placement, max_output: usize) -> Result<String> {
    let mut reader = quick_xml::Reader::from_reader(fragment.as_bytes());
    let mut depth = 0usize;
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
    let name_end = open[1..]
        .find(|character: char| {
            character.is_ascii_whitespace() || character == '>' || character == '/'
        })
        .map_or(open.len() - 1, |offset| offset + 1);
    let root_name = &open[1..name_end];
    let prefix = root_name.find(':').map_or("", |at| &root_name[..=at]);
    let attributes = open
        .find(char::is_whitespace)
        .map_or("", |at| &open[at..open.len() - 1]);
    let (attributes_before, attributes_after) = if matches!(placement, Placement::TwoCell { .. }) {
        (attributes, "")
    } else {
        attribute_parts(attributes, "editAs")
    };

    let mut measured = LengthWriter::default();
    write_replaced_placement(
        &mut measured,
        fragment,
        placement,
        prefix,
        attributes_before,
        attributes_after,
        object_start,
        root_close,
    )?;
    if measured.length > max_output {
        return Err(Error::limit(
            LimitKind::DrawingBytes,
            measured.length,
            max_output,
        ));
    }
    let mut output = String::with_capacity(measured.length);
    write_replaced_placement(
        &mut output,
        fragment,
        placement,
        prefix,
        attributes_before,
        attributes_after,
        object_start,
        root_close,
    )?;
    debug_assert_eq!(output.len(), measured.length);
    Ok(output)
}

fn write_replaced_placement<W: fmt::Write>(
    out: &mut W,
    fragment: &str,
    placement: Placement,
    prefix: &str,
    attributes_before: &str,
    attributes_after: &str,
    object_start: usize,
    root_close: usize,
) -> Result<()> {
    let element = placement.element();
    let _ = write!(out, "<{prefix}{element}");
    put(out, attributes_before)?;
    put(out, attributes_after)?;
    put(out, ">")?;
    write_placement(out, placement, prefix)?;
    put(out, &fragment[object_start..root_close])?;
    let _ = write!(out, "</{prefix}{element}>");
    Ok(())
}

fn write_placement<W: fmt::Write>(out: &mut W, placement: Placement, prefix: &str) -> Result<()> {
    let marker = |out: &mut W, tag: &str, marker: Marker| {
        let _ = write!(
            out,
            "<{prefix}{tag}><{prefix}col>{}</{prefix}col><{prefix}colOff>{}</{prefix}colOff><{prefix}row>{}</{prefix}row><{prefix}rowOff>{}</{prefix}rowOff></{prefix}{tag}>",
            marker.col, marker.col_offset, marker.row, marker.row_offset
        );
    };
    match placement {
        Placement::TwoCell { from, to } => {
            marker(out, "from", from);
            marker(out, "to", to);
        },
        Placement::OneCell { from, cx, cy } => {
            marker(out, "from", from);
            let _ = write!(out, "<{prefix}ext cx=\"{cx}\" cy=\"{cy}\"/>");
        },
        Placement::Absolute { x, y, cx, cy } => {
            let _ = write!(
                out,
                "<{prefix}pos x=\"{x}\" y=\"{y}\"/><{prefix}ext cx=\"{cx}\" cy=\"{cy}\"/>"
            );
        },
    }
    Ok(())
}

fn attribute_parts<'a>(attributes: &'a str, name: &str) -> (&'a str, &'a str) {
    let needle = format!(" {name}=\"");
    match attributes.find(&needle) {
        Some(at) => {
            let value_start = at + needle.len();
            let value_end = attributes[value_start..]
                .find('"')
                .map_or(attributes.len(), |offset| value_start + offset + 1);
            (&attributes[..at], &attributes[value_end..])
        },
        None => (attributes, ""),
    }
}

/// Set (`Some`) or remove (`None`) an attribute of the first element with
/// local name `local` in a fragment.
fn set_tag_attribute(
    fragment: &str,
    local: &str,
    name: &str,
    value: Option<&str>,
    escape_value: bool,
    max_output: usize,
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
    let (body_before, body_after) = attribute_parts(body, name);
    let value_length = match value {
        Some(value) if escape_value => {
            let mut measured = LengthWriter::default();
            escaped(&mut measured, value)?;
            measured.length
        },
        Some(value) => value.len(),
        None => 0,
    };
    let attribute_length = value.map_or(0, |_| {
        name.len().saturating_add(value_length).saturating_add(4)
    });
    let output_length = fragment[..start]
        .len()
        .checked_add(body_before.len())
        .and_then(|length| length.checked_add(body_after.len()))
        .and_then(|length| length.checked_add(attribute_length))
        .and_then(|length| length.checked_add(self_closing.len()))
        .and_then(|length| length.checked_add(fragment[end..].len()))
        .ok_or_else(|| Error::limit(LimitKind::DrawingBytes, usize::MAX, max_output))?;
    if output_length > max_output {
        return Err(Error::limit(
            LimitKind::DrawingBytes,
            output_length,
            max_output,
        ));
    }
    let mut output = String::with_capacity(output_length);
    output.push_str(&fragment[..start]);
    output.push_str(body_before);
    output.push_str(body_after);
    if let Some(value) = value {
        let _ = write!(output, " {name}=\"");
        if escape_value {
            escaped(&mut output, value)?;
        } else {
            output.push_str(value);
        }
        output.push('"');
    }
    output.push_str(self_closing);
    output.push_str(&fragment[end..]);
    debug_assert_eq!(output.len(), output_length);
    Ok(output)
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

/// The internal drawing part a worksheet's relationships name, if any.
///
/// # Errors
///
/// Returns [`Error::Invalid`] when the worksheet has an external or a second
/// drawing relationship, and [`Error::Opc`] for an unresolvable target.
pub fn drawing_of(relationships: &Relationships) -> Result<Option<PackURI>> {
    drawing_in(relationships)
}

/// The top-level objects of a drawing part, in drawing order, with their
/// images and click hyperlinks resolved through the part's relationships.
/// `image` answers an image part's content type and size.
///
/// # Errors
///
/// Returns [`Error::Invalid`] for malformed markup or an exceeded limit.
pub fn read<C: Color>(
    xml: &[u8],
    relationships: &Relationships,
    image: impl Fn(&PackURI) -> Option<(String, u64)>,
    limits: &Limits,
) -> Result<Vec<Element<C>>> {
    let limits = limits.validate()?;
    let (mut elements, _, _) = scan(xml, limits)?;
    resolve(&mut elements, relationships, image, limits)?;
    Ok(elements)
}

/// A chart frame of a drawing part: where it sits, its name and the chart
/// part its relationship targets.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub struct ChartFrame {
    /// Where the frame sits.
    pub placement: Option<Placement>,
    /// The frame's `cNvPr/@name`.
    pub name: Option<String>,
    /// The chart part.
    pub target: PackURI,
}

/// The chart frames of a drawing part, in drawing order.
///
/// # Errors
///
/// Returns [`Error::Invalid`] for malformed markup or an exceeded limit.
pub fn chart_frames(
    xml: &[u8],
    relationships: &Relationships,
    limits: &Limits,
) -> Result<Vec<ChartFrame>> {
    let limits = limits.validate()?;
    let (elements, _, _) = scan::<u32>(xml, limits)?;
    Ok(elements
        .into_iter()
        .filter(|element| element.kind == ElementKind::Frame)
        .filter_map(|element| {
            let target = relationship_ids(&xml[element.span.clone()])
                .into_iter()
                .find_map(|id| {
                    let relationship = relationships.get(&id)?;
                    (!relationship.is_external()
                        && matches!(relationship.reltype(), rt::CHART | rt::STRICT_CHART))
                    .then(|| relationship.target_partname().ok())
                    .flatten()
                })?;
            Some(ChartFrame {
                placement: element.placement,
                name: element.name,
                target,
            })
        })
        .collect())
}

/// How many top-level objects of `kind` a drawing part holds.
///
/// # Errors
///
/// Returns [`Error::Invalid`] for malformed markup or an exceeded limit.
pub fn count(xml: &[u8], kind: ElementKind, limits: &Limits) -> Result<usize> {
    let limits = limits.validate()?;
    Ok(scan::<u32>(xml, limits)?
        .0
        .iter()
        .filter(|element| element.kind == kind)
        .count())
}

/// The `cNvPr/@id` a new object of a drawing part takes: one past the
/// largest in use, and at least 2.
///
/// # Errors
///
/// Returns [`Error::Invalid`] for malformed markup or an exceeded limit.
pub fn next_id(xml: &[u8], limits: &Limits) -> Result<u32> {
    let limits = limits.validate()?;
    let (_, _, max_id) = scan::<u32>(xml, limits)?;
    Ok(max_id.saturating_add(1).max(2))
}

/// A drawing part with `fragment` appended as its last anchor.
///
/// # Errors
///
/// Returns [`Error::Invalid`] for malformed markup or a part that would
/// exceed [`MAX_DRAWING_BYTES`].
pub fn append(xml: &[u8], fragment: &str, limits: &Limits) -> Result<Vec<u8>> {
    let limits = limits.validate()?;
    let (elements, root_close, _) = scan::<u32>(xml, limits)?;
    let added = fragment_elements(fragment)?;
    let total = elements
        .len()
        .checked_add(added)
        .ok_or_else(|| Error::limit(LimitKind::Elements, usize::MAX, limits.max_elements))?;
    limits.elements(total)?;
    let output_len = xml.len().checked_add(fragment.len()).ok_or_else(|| {
        Error::limit(
            LimitKind::DrawingBytes,
            usize::MAX,
            limits.max_drawing_bytes,
        )
    })?;
    limits.drawing_bytes(output_len)?;
    let mut out = Vec::with_capacity(output_len);
    out.extend_from_slice(xml);
    out.splice(root_close..root_close, fragment.bytes());
    Ok(out)
}

/// Check a picture before anything is staged for it: its image size,
/// content type, extension and placement.
///
/// # Errors
///
/// Returns [`Error::Invalid`] for the first value that is out of bounds.
pub fn check_picture(spec: &PictureSpec, limits: &Limits) -> Result<()> {
    let limits = limits.validate()?;
    limits.image_bytes(spec.image.len())?;
    if !spec.content_type.starts_with("image/")
        || spec.extension.is_empty()
        || !spec.extension.bytes().all(|b| b.is_ascii_alphanumeric())
    {
        return Err(invalid(
            "a picture needs an image content type and extension",
        ));
    }
    spec.placement.validate()
}

/// The anchor markup of a new picture with `cNvPr/@id` `id` showing the
/// image relationship `relationship_id`.
///
/// # Errors
///
/// Returns [`Error::Invalid`] for an invalid placement, name or rotation.
pub fn picture_fragment(
    spec: &PictureSpec,
    id: u32,
    relationship_id: &str,
    limits: &Limits,
) -> Result<String> {
    let limits = limits.validate()?;
    let mut measured = LengthWriter::default();
    write_picture_xml(&mut measured, spec, id, relationship_id)?;
    limits.drawing_bytes(measured.length)?;
    let mut fragment = String::with_capacity(measured.length);
    write_picture_xml(&mut fragment, spec, id, relationship_id)?;
    debug_assert_eq!(fragment.len(), measured.length);
    Ok(fragment)
}

/// The anchor markup of a new preset shape with `cNvPr/@id` `id`.
///
/// # Errors
///
/// Returns [`Error::Invalid`] for an invalid placement, geometry, fill,
/// line, text or font.
pub fn shape_fragment<C: Color>(spec: &ShapeSpec<C>, id: u32, limits: &Limits) -> Result<String> {
    let limits = limits.validate()?;
    let mut measured = LengthWriter::default();
    write_shape_xml(&mut measured, spec, id)?;
    limits.drawing_bytes(measured.length)?;
    let mut fragment = String::with_capacity(measured.length);
    write_shape_xml(&mut fragment, spec, id)?;
    debug_assert_eq!(fragment.len(), measured.length);
    Ok(fragment)
}

/// The byte range of the object at `index` among a drawing part's objects of
/// `kind`.
fn select(xml: &[u8], kind: ElementKind, index: usize, limits: Limits) -> Result<Range<usize>> {
    let (elements, _, _) = scan::<u32>(xml, limits)?;
    elements
        .into_iter()
        .filter(|element| element.kind == kind)
        .nth(index)
        .map(|element| element.span)
        .ok_or_else(|| invalid(format!("drawing element {index} does not exist")))
}

/// A drawing part with the object at `index` among its objects of `kind`
/// changed in place: its placement, names, rotation or flips. `None` when
/// the change leaves the object as it is; everything else of the part keeps
/// its bytes.
///
/// # Errors
///
/// Returns [`Error::Invalid`] for a missing object, an invalid change, or
/// markup this module cannot rewrite.
pub fn change(
    xml: &[u8],
    kind: ElementKind,
    index: usize,
    change: &Change,
    limits: &Limits,
) -> Result<Option<Vec<u8>>> {
    let limits = limits.validate()?;
    let span = select(xml, kind, index, limits)?;
    let fragment = std::str::from_utf8(&xml[span.clone()])
        .map_err(|_utf8| invalid("drawing markup is not UTF-8"))?;
    let base_len = xml.len().checked_sub(span.len()).ok_or_else(|| {
        Error::limit(
            LimitKind::DrawingBytes,
            usize::MAX,
            limits.max_drawing_bytes,
        )
    })?;
    let max_updated = limits
        .max_drawing_bytes
        .checked_sub(base_len)
        .ok_or_else(|| {
            Error::limit(LimitKind::DrawingBytes, xml.len(), limits.max_drawing_bytes)
        })?;
    let mut updated = None;
    if let Some(placement) = change.placement {
        placement.validate()?;
        updated = Some(replace_placement(fragment, placement, max_updated)?);
    }
    if let Some(name) = &change.name {
        check_text(name, "name")?;
        let current = updated.as_deref().unwrap_or(fragment);
        updated = Some(set_tag_attribute(
            current,
            "cNvPr",
            "name",
            Some(name),
            true,
            max_updated,
        )?);
    }
    if let Some(description) = &change.description {
        if let Some(value) = description {
            check_text(value, "description")?;
        }
        let current = updated.as_deref().unwrap_or(fragment);
        updated = set_tag_attribute(
            current,
            "cNvPr",
            "descr",
            description.as_deref(),
            true,
            max_updated,
        )
        .map(Some)?;
    }
    if change.rotation.is_some() || change.flip_h.is_some() || change.flip_v.is_some() {
        let current = updated.as_deref().unwrap_or(fragment);
        if !current.contains(":xfrm") && !current.contains("<xfrm") {
            return Err(invalid("the element has no transform to change"));
        }
        if let Some(degrees) = change.rotation {
            let value = rotation(Some(degrees))?.map(|value| value.to_string());
            let current = updated.as_deref().unwrap_or(fragment);
            updated = Some(set_tag_attribute(
                current,
                "xfrm",
                "rot",
                value.as_deref(),
                false,
                max_updated,
            )?);
        }
        for (key, value) in [("flipH", change.flip_h), ("flipV", change.flip_v)] {
            if let Some(value) = value {
                let current = updated.as_deref().unwrap_or(fragment);
                updated = Some(set_tag_attribute(
                    current,
                    "xfrm",
                    key,
                    value.then_some("1"),
                    false,
                    max_updated,
                )?);
            }
        }
    }
    let updated_ref = updated.as_deref().unwrap_or(fragment);
    if updated_ref == fragment {
        return Ok(None);
    }
    let output_len = xml
        .len()
        .checked_sub(span.len())
        .and_then(|length| length.checked_add(updated_ref.len()))
        .ok_or_else(|| {
            Error::limit(
                LimitKind::DrawingBytes,
                usize::MAX,
                limits.max_drawing_bytes,
            )
        })?;
    limits.drawing_bytes(output_len)?;
    let mut out = Vec::with_capacity(output_len);
    out.extend_from_slice(xml);
    let replacement = updated
        .map(String::into_bytes)
        .ok_or_else(|| invalid("drawing change did not produce a replacement"))?;
    out.splice(span, replacement);
    if scan::<u32>(&out, limits)?.0.len() != scan::<u32>(xml, limits)?.0.len() {
        return Err(invalid("the changed drawing element did not read back"));
    }
    Ok(Some(out))
}

/// A drawing part with one object removed, and the relationships it named
/// that nothing left in the part names.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub struct Detached {
    /// The part without the object.
    pub xml: Vec<u8>,
    /// The relationship ids only the removed object named, in its order.
    pub released: Vec<String>,
}

/// Remove the object at `index` among a drawing part's objects of `kind`.
///
/// # Errors
///
/// Returns [`Error::Invalid`] for malformed markup or a missing object.
pub fn detach(xml: &[u8], kind: ElementKind, index: usize, limits: &Limits) -> Result<Detached> {
    let limits = limits.validate()?;
    let span = select(xml, kind, index, limits)?;
    let mut out = Vec::with_capacity(xml.len());
    out.extend_from_slice(xml);
    let fragment: Vec<u8> = out.drain(span).collect();
    let rest = String::from_utf8_lossy(&out).into_owned();
    let released = relationship_ids(&fragment)
        .into_iter()
        .filter(|id| !rest.contains(&format!("\"{id}\"")))
        .collect();
    Ok(Detached { xml: out, released })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn output_limits_are_admitted_before_rewrite_allocation() {
        let source = EMPTY_DRAWING.as_bytes();
        let exact = Limits::new(1, source.len(), MAX_IMAGE_BYTES);
        assert_eq!(count(source, ElementKind::Other, &exact).unwrap(), 0);
        assert!(matches!(
            append(source, "<xdr:twoCellAnchor/>", &exact),
            Err(Error::Limit {
                kind: LimitKind::DrawingBytes,
                ..
            })
        ));
    }

    #[test]
    fn an_invalid_caller_policy_is_typed() {
        assert!(matches!(
            count(
                EMPTY_DRAWING.as_bytes(),
                ElementKind::Other,
                &Limits::new(0, 1, 1)
            ),
            Err(Error::InvalidLimit {
                kind: LimitKind::Elements,
                value: 0,
                ..
            })
        ));
    }

    #[test]
    fn fragment_output_is_admitted_before_buffer_allocation() {
        let spec = ShapeSpec::<u32>::new(
            Placement::Absolute {
                x: 0,
                y: 0,
                cx: 1,
                cy: 1,
            },
            "shape",
        );
        let limits = Limits::new(1, 1, 1);
        assert!(matches!(
            shape_fragment(&spec, 2, &limits),
            Err(Error::Limit {
                kind: LimitKind::DrawingBytes,
                ..
            })
        ));
    }

    #[test]
    fn append_admits_all_fragment_elements_before_buffer_allocation() {
        let limits = Limits::new(1, MAX_DRAWING_BYTES, MAX_IMAGE_BYTES);
        let fragment = "<xdr:twoCellAnchor/><xdr:twoCellAnchor/>";
        assert!(matches!(
            append(EMPTY_DRAWING.as_bytes(), fragment, &limits),
            Err(Error::Limit {
                kind: LimitKind::Elements,
                observed: 2,
                maximum: 1,
            })
        ));
    }

    #[test]
    fn change_admits_replacement_bytes_before_rewrite_growth() {
        let source = concat!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
            "\n",
            r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"><xdr:twoCellAnchor><xdr:from/><xdr:to/><xdr:sp><xdr:nvSpPr><xdr:cNvPr id="2" name="a"/></xdr:nvSpPr></xdr:sp></xdr:twoCellAnchor></xdr:wsDr>"#
        );
        let update = Change {
            name: Some("x".repeat(1024)),
            ..Default::default()
        };
        let limits = Limits::new(1, source.len(), MAX_IMAGE_BYTES);
        assert!(matches!(
            change(source.as_bytes(), ElementKind::Shape, 0, &update, &limits),
            Err(Error::Limit {
                kind: LimitKind::DrawingBytes,
                ..
            })
        ));
    }

    #[test]
    fn change_preserves_noop_and_allows_a_bounded_shrink() {
        let noop_source = concat!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
            "\n",
            r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"><xdr:twoCellAnchor><xdr:from/><xdr:to/><xdr:sp><xdr:nvSpPr><xdr:cNvPr id="2" name="a"/></xdr:nvSpPr></xdr:sp></xdr:twoCellAnchor></xdr:wsDr>"#
        );
        let noop = Change {
            name: Some("a".to_owned()),
            ..Default::default()
        };
        let limits = Limits::new(1, noop_source.len(), MAX_IMAGE_BYTES);
        assert_eq!(
            change(
                noop_source.as_bytes(),
                ElementKind::Shape,
                0,
                &noop,
                &limits
            )
            .unwrap(),
            None
        );

        let shrink_source = concat!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
            "\n",
            r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"><xdr:twoCellAnchor><xdr:from/><xdr:to/><xdr:sp><xdr:nvSpPr><xdr:cNvPr id="2" name="a" descr="long description"/></xdr:nvSpPr></xdr:sp></xdr:twoCellAnchor></xdr:wsDr>"#
        );
        let shrink = Change {
            description: Some(None),
            ..Default::default()
        };
        let changed = change(
            shrink_source.as_bytes(),
            ElementKind::Shape,
            0,
            &shrink,
            &Limits::new(1, shrink_source.len(), MAX_IMAGE_BYTES),
        )
        .unwrap()
        .unwrap();
        assert!(changed.len() < shrink_source.len());
        assert!(!String::from_utf8_lossy(&changed).contains("descr="));
    }
}
