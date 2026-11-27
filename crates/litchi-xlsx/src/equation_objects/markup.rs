//! The markup of a worksheet OLE equation object, written as Excel writes an
//! embedded object, and the bounded scanners that find one object's markup
//! in the worksheet, its legacy VML drawing and its `DrawingML` drawing.
//!
//! An object has three pieces of markup that share its shape id `N`:
//!
//! - in the worksheet's `<oleObjects>` (ECMA-376 Part 1 §18.3.1.59 and
//!   §18.3.1.60), an `mc:AlternateContent` whose `x14` choice holds the
//!   `oleObject` with its `objectPr` (preview relationship and two-cell
//!   anchor, the Office 2010 extension of [MS-XLSX]) and whose fallback
//!   holds the bare `oleObject`, as Excel 2010 and later write it;
//! - in the legacy VML drawing (`legacyDrawing`), the picture shape
//!   `_x0000_sN` (`x:ClientData ObjectType="Pict"`) whose `v:imagedata`
//!   names the preview through the VML part's relationships: Excel reads an
//!   object's shape from the VML drawing ([MS-OI29500]);
//! - in the `DrawingML` drawing, a hidden `xdr:sp` with id `N` linked to the
//!   VML shape by `a14:compatExt spid="_x0000_sN"` inside an
//!   `mc:AlternateContent` that requires `a14` ([MS-ODRAWXML]).
//!
//! The Excel-written objects in `test-data` (`excel-links-objects.xlsx`,
//! `bug64512_embed.xlsx`) have exactly this shape.

use std::collections::{BTreeMap, HashSet};
use std::fmt::Write as _;
use std::ops::Range;

use litchi_core::xml::escape_xml;
use litchi_ooxml_common::xml::attributes::first_wins;
use quick_xml::events::{BytesStart, Event};

use crate::error::{Result, invalid};
use crate::ole_objects::OleObjectMarker;

/// EMUs per point.
pub(super) const EMU_PER_POINT: f64 = 12_700.0;
/// EMUs per pixel at 96 pixels per inch.
pub(super) const EMU_PER_PIXEL: i64 = 9_525;
/// The most objects one worksheet's `<oleObjects>` may hold.
pub(super) const MAX_OBJECTS: usize = 1_024;
/// The last zero-based column of a worksheet.
const LAST_COLUMN: u32 = 16_383;
/// The last zero-based row of a worksheet.
const LAST_ROW: u32 = 1_048_575;
/// The most rows past an object's top row whose heights are read.
const ROW_SPAN: u32 = 65_536;

const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const X14: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const XDR: &str = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const A14: &str = "http://schemas.microsoft.com/office/drawing/2010/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
/// The `a:ext` URI of `a14:compatExt` ([MS-ODRAWXML]).
const COMPAT_EXT_URI: &str = "{63B3BB69-23CF-44E3-9099-C40C66FF867C}";

/// The `_x0000_t75` picture-frame shape type Excel writes before the first
/// picture or object shape of a VML drawing.
pub(super) const PICTURE_SHAPE_TYPE: &str = concat!(
    "<v:shapetype id=\"_x0000_t75\" coordsize=\"21600,21600\" o:spt=\"75\" ",
    "o:preferrelative=\"t\" path=\"m@4@5l@4@11@9@11@9@5xe\" filled=\"f\" stroked=\"f\">",
    "<v:stroke joinstyle=\"miter\"/><v:formulas>",
    "<v:f eqn=\"if lineDrawn pixelLineWidth 0\"/><v:f eqn=\"sum @0 1 0\"/>",
    "<v:f eqn=\"sum 0 0 @1\"/><v:f eqn=\"prod @2 1 2\"/>",
    "<v:f eqn=\"prod @3 21600 pixelWidth\"/><v:f eqn=\"prod @3 21600 pixelHeight\"/>",
    "<v:f eqn=\"sum @0 0 1\"/><v:f eqn=\"prod @6 1 2\"/>",
    "<v:f eqn=\"prod @7 21600 pixelWidth\"/><v:f eqn=\"sum @8 21600 0\"/>",
    "<v:f eqn=\"prod @7 21600 pixelHeight\"/><v:f eqn=\"sum @10 21600 0\"/>",
    "</v:formulas><v:path o:extrusionok=\"f\" gradientshapeok=\"t\" o:connecttype=\"rect\"/>",
    "<o:lock v:ext=\"edit\" aspectratio=\"t\"/></v:shapetype>"
);

/// The two-cell anchor of an object: its top-left and bottom-right markers,
/// offsets in EMUs.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct Anchor {
    pub(super) from: OleObjectMarker,
    pub(super) to: OleObjectMarker,
}

/// Where an object sits in a worksheet: its anchor and its rectangle in
/// points, as the VML shape's style records it.
#[derive(Debug, Clone, Copy, PartialEq)]
pub(super) struct Frame {
    pub(super) anchor: Anchor,
    pub(super) left_pt: f64,
    pub(super) top_pt: f64,
    pub(super) width_pt: f64,
    pub(super) height_pt: f64,
}

impl Frame {
    /// The VML `x:Anchor` text: columns, rows and pixel offsets of both
    /// corners.
    fn vml_anchor(&self) -> String {
        let px = |emu: i64| emu / EMU_PER_PIXEL;
        let Anchor { from, to } = self.anchor;
        format!(
            "{}, {}, {}, {}, {}, {}, {}, {}",
            from.column,
            px(from.column_offset),
            from.row,
            px(from.row_offset),
            to.column,
            px(to.column_offset),
            to.row,
            px(to.row_offset)
        )
    }
}

/// A point length as VML writes it (`48pt`, `12.75pt`).
fn pt(value: f64) -> String {
    let rounded = (value * 100.0).round() / 100.0;
    format!("{rounded}pt")
}

// ---------------------------------------------------------------------------
// Worksheet grid

/// The column widths and row heights of a worksheet, read for placing an
/// object: `sheetFormatPr`, `cols` and the heights of the rows the object
/// can reach.
#[derive(Debug, Clone)]
pub(super) struct Grid {
    default_column_px: f64,
    /// Column spans `(first, last, pixels)`, zero-based.
    columns: Vec<(u32, u32, f64)>,
    default_row_pt: f64,
    rows: BTreeMap<u32, f64>,
}

/// The pixel width of a column of `width` characters in the default font
/// (maximum digit width 7 pixels, ECMA-376 Part 1 §18.3.1.13).
fn column_pixels(width: f64) -> f64 {
    if !(width.is_finite() && width > 0.0) {
        return 0.0;
    }
    (((256.0 * width + (128.0_f64 / 7.0).trunc()) / 256.0) * 7.0).trunc()
}

impl Grid {
    /// Read the grid of a worksheet part for an object whose top row is
    /// `top_row`.
    pub(super) fn read(xml: &[u8], top_row: u32) -> Result<Self> {
        let mut grid = Self {
            default_column_px: 64.0,
            columns: Vec::new(),
            default_row_pt: 15.0,
            rows: BTreeMap::new(),
        };
        let last_row = top_row.saturating_add(ROW_SPAN);
        let mut reader = quick_xml::Reader::from_reader(xml);
        let mut depth = 0usize;
        let mut in_cols = false;
        loop {
            let event = reader
                .read_event()
                .map_err(|error| invalid(format!("worksheet markup: {error}")))?;
            let (element, empty) = match &event {
                Event::Start(element) => (element, false),
                Event::Empty(element) => (element, true),
                Event::End(element) => {
                    depth = depth.saturating_sub(1);
                    if element.local_name().as_ref() == b"cols" {
                        in_cols = false;
                    }
                    continue;
                },
                Event::Eof => break,
                _ => continue,
            };
            match (depth, element.local_name().as_ref()) {
                (1, b"sheetFormatPr") => {
                    let base = number(element, b"baseColWidth").unwrap_or(8.0);
                    grid.default_column_px = match number(element, b"defaultColWidth") {
                        Some(width) => column_pixels(width),
                        // `baseColWidth` digits and 5 pixels of padding,
                        // rounded up to a multiple of 8 pixels.
                        None => ((base * 7.0 + 5.0) / 8.0).ceil() * 8.0,
                    };
                    if let Some(height) = number(element, b"defaultRowHeight") {
                        grid.default_row_pt = height.max(0.0);
                    }
                },
                (1, b"cols") => in_cols = !empty,
                (2, b"col") if in_cols => {
                    let first = number(element, b"min").unwrap_or(1.0);
                    let last = number(element, b"max").unwrap_or(first);
                    let width = if flag(element, b"hidden") {
                        0.0
                    } else {
                        number(element, b"width").map_or(grid.default_column_px, column_pixels)
                    };
                    if first >= 1.0 && last >= first && grid.columns.len() < 16_384 {
                        grid.columns.push((
                            index(first - 1.0, LAST_COLUMN),
                            index(last - 1.0, LAST_COLUMN),
                            width,
                        ));
                    }
                },
                (2, b"row") => {
                    if let Some(row) = number(element, b"r") {
                        let row = index(row - 1.0, LAST_ROW);
                        if (top_row..=last_row).contains(&row) {
                            if flag(element, b"hidden") {
                                grid.rows.insert(row, 0.0);
                            } else if let Some(height) = number(element, b"ht") {
                                grid.rows.insert(row, height.max(0.0));
                            }
                        }
                    }
                },
                _ => {},
            }
            if !empty {
                depth += 1;
            }
        }
        Ok(grid)
    }

    fn column_emu(&self, column: u32) -> i64 {
        let px = self
            .columns
            .iter()
            .find(|(first, last, _)| (*first..=*last).contains(&column))
            .map_or(self.default_column_px, |(_, _, px)| *px);
        (px * EMU_PER_PIXEL as f64).round() as i64
    }

    fn row_emu(&self, row: u32) -> i64 {
        let height = self.rows.get(&row).copied().unwrap_or(self.default_row_pt);
        (height * EMU_PER_POINT).round() as i64
    }

    /// The cell and offset `length` EMUs after the start of `start`.
    fn reach(start: u32, length: i64, last: u32, size: impl Fn(u32) -> i64) -> (u32, i64) {
        let mut index = start;
        let mut remaining = length.max(0);
        loop {
            let extent = size(index);
            if remaining < extent || index == last {
                return (index, remaining.min(extent.max(0)));
            }
            remaining -= extent;
            index += 1;
        }
    }

    /// The position in points of the start of `column` (`row`), summed over
    /// at most the columns (rows) before it.
    fn offset_pt(&self, column: u32, row: u32) -> (f64, f64) {
        let left = (0..column).map(|index| self.column_emu(index)).sum::<i64>();
        let top = if row > ROW_SPAN {
            // Far rows: the default height, without walking every row.
            (f64::from(row) * self.default_row_pt * EMU_PER_POINT) as i64
        } else {
            (0..row).map(|index| self.row_emu(index)).sum::<i64>()
        };
        (left as f64 / EMU_PER_POINT, top as f64 / EMU_PER_POINT)
    }

    /// The frame of an object of `size_pt` whose top-left corner is the
    /// top-left corner of the cell (`row`, `column`).
    pub(super) fn frame(&self, row: u32, column: u32, size_pt: (f64, f64)) -> Frame {
        self.frame_from(
            OleObjectMarker {
                column,
                column_offset: 0,
                row,
                row_offset: 0,
            },
            size_pt,
        )
    }

    /// [`Self::frame`] from an existing top-left marker, keeping its offsets.
    pub(super) fn frame_from(&self, from: OleObjectMarker, size_pt: (f64, f64)) -> Frame {
        let start_x = from.column_offset.max(0);
        let start_y = from.row_offset.max(0);
        let width = (size_pt.0 * EMU_PER_POINT).round() as i64 + start_x;
        let height = (size_pt.1 * EMU_PER_POINT).round() as i64 + start_y;
        let column = from.column.min(LAST_COLUMN);
        let row = from.row.min(LAST_ROW);
        let (to_column, column_offset) =
            Self::reach(column, width, LAST_COLUMN, |index| self.column_emu(index));
        let (to_row, row_offset) = Self::reach(row, height, LAST_ROW, |index| self.row_emu(index));
        let (left_pt, top_pt) = self.offset_pt(column, row);
        Frame {
            anchor: Anchor {
                from: OleObjectMarker {
                    column,
                    column_offset: start_x,
                    row,
                    row_offset: start_y,
                },
                to: OleObjectMarker {
                    column: to_column,
                    column_offset,
                    row: to_row,
                    row_offset,
                },
            },
            left_pt: left_pt + start_x as f64 / EMU_PER_POINT,
            top_pt: top_pt + start_y as f64 / EMU_PER_POINT,
            width_pt: size_pt.0,
            height_pt: size_pt.1,
        }
    }
}

fn attribute(element: &BytesStart<'_>, name: &[u8]) -> Option<String> {
    first_wins(element).flatten().find_map(|attribute| {
        (attribute.key.local_name().as_ref() == name)
            .then(|| String::from_utf8_lossy(&attribute.value).into_owned())
    })
}

fn number(element: &BytesStart<'_>, name: &[u8]) -> Option<f64> {
    attribute(element, name)?
        .trim()
        .parse::<f64>()
        .ok()
        .filter(|value| value.is_finite())
}

fn flag(element: &BytesStart<'_>, name: &[u8]) -> bool {
    attribute(element, name).is_some_and(|value| matches!(value.trim(), "1" | "true"))
}

/// A zero-based index from a parsed number, clamped to `last`.
fn index(value: f64, last: u32) -> u32 {
    if value <= 0.0 {
        0
    } else if value >= f64::from(last) {
        last
    } else {
        value as u32
    }
}

// ---------------------------------------------------------------------------
// Worksheet `<oleObjects>`

/// One direct child of `<oleObjects>`: an `mc:AlternateContent` or a bare
/// `oleObject`, with the shape id of the first `oleObject` it holds.
#[derive(Debug, Clone)]
pub(super) struct Entry {
    pub(super) span: Range<usize>,
    pub(super) shape_id: Option<u32>,
}

/// The children of a worksheet's `<oleObjects>` element, and where a new
/// child is inserted.
#[derive(Debug, Clone)]
pub(super) struct Collection {
    pub(super) span: Range<usize>,
    /// The offset of the closing tag, `None` for an empty element.
    pub(super) close: Option<usize>,
    pub(super) entries: Vec<Entry>,
}

impl Collection {
    /// Scan the `<oleObjects>` element at `span` of `xml`.
    pub(super) fn scan(xml: &[u8], span: Range<usize>) -> Result<Self> {
        let fragment = xml
            .get(span.clone())
            .ok_or_else(|| invalid("the oleObjects span is outside the worksheet"))?;
        let mut reader = quick_xml::Reader::from_reader(fragment);
        let mut depth = 0usize;
        let mut close = None;
        let mut entries = Vec::new();
        let mut open: Option<Entry> = None;
        loop {
            let start = offset(&reader)?;
            let event = reader
                .read_event()
                .map_err(|error| invalid(format!("oleObjects markup: {error}")))?;
            let end = offset(&reader)?;
            match &event {
                Event::Start(element) | Event::Empty(element) => {
                    let empty = matches!(event, Event::Empty(_));
                    if depth == 1 {
                        if entries.len() >= MAX_OBJECTS {
                            return Err(invalid("the worksheet holds too many OLE objects"));
                        }
                        open = Some(Entry {
                            span: span.start + start..span.start + end,
                            shape_id: None,
                        });
                    }
                    if depth >= 1
                        && element.local_name().as_ref() == b"oleObject"
                        && let Some(entry) = open.as_mut()
                        && entry.shape_id.is_none()
                    {
                        entry.shape_id = attribute(element, b"shapeId")
                            .and_then(|value| value.trim().parse().ok());
                    }
                    if depth == 1 && empty {
                        entries.extend(open.take());
                    }
                    if !empty {
                        depth += 1;
                    }
                },
                Event::End(_) => {
                    depth = depth.saturating_sub(1);
                    if depth == 1
                        && let Some(mut entry) = open.take()
                    {
                        entry.span.end = span.start + end;
                        entries.push(entry);
                    } else if depth == 0 {
                        close = Some(span.start + start);
                        break;
                    }
                },
                Event::Eof => break,
                _ => {},
            }
            if depth == 0 && matches!(event, Event::Empty(_)) {
                break;
            }
        }
        Ok(Self {
            span,
            close,
            entries,
        })
    }

    /// The entry of the object with shape id `shape_id`.
    pub(super) fn entry(&self, shape_id: u32) -> Option<&Entry> {
        self.entries
            .iter()
            .find(|entry| entry.shape_id == Some(shape_id))
    }
}

fn offset(reader: &quick_xml::Reader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position()).map_err(|_| invalid("XML offset overflow"))
}

/// The worksheet markup of a new object: Excel 2010's `x14` choice with the
/// preview and anchor, and the bare `oleObject` fallback. `prefix` is the
/// worksheet root's prefix with its colon, or empty.
pub(super) fn worksheet_entry(
    prefix: &str,
    prog_id: &str,
    shape_id: u32,
    object_id: &str,
    preview_id: &str,
    anchor: &Anchor,
) -> String {
    let mut out = String::with_capacity(768);
    let object = format!(
        "<{prefix}oleObject progId=\"{}\" shapeId=\"{shape_id}\" r:id=\"{}\"",
        escape_xml(prog_id),
        escape_xml(object_id)
    );
    let _ = write!(
        out,
        "<mc:AlternateContent xmlns:mc=\"{MC}\" xmlns:x14=\"{X14}\" xmlns:xdr=\"{XDR}\" xmlns:r=\"{REL}\"><mc:Choice Requires=\"x14\">{object}><{prefix}objectPr defaultSize=\"0\" r:id=\"{}\"><{prefix}anchor moveWithCells=\"1\">",
        escape_xml(preview_id)
    );
    for (name, marker) in [("from", &anchor.from), ("to", &anchor.to)] {
        let _ = write!(
            out,
            "<{prefix}{name}><xdr:col>{}</xdr:col><xdr:colOff>{}</xdr:colOff><xdr:row>{}</xdr:row><xdr:rowOff>{}</xdr:rowOff></{prefix}{name}>",
            marker.column, marker.column_offset, marker.row, marker.row_offset
        );
    }
    let _ = write!(
        out,
        "</{prefix}anchor></{prefix}objectPr></{prefix}oleObject></mc:Choice><mc:Fallback>{object}/></mc:Fallback></mc:AlternateContent>"
    );
    out
}

/// A new `<oleObjects>` element holding `entry`.
pub(super) fn worksheet_collection(prefix: &str, entry: &str) -> String {
    format!("<{prefix}oleObjects>{entry}</{prefix}oleObjects>")
}

/// The offset a new `<oleObjects>` element takes among the worksheet root's
/// children (after `picture`, before `controls`, ECMA-376 Part 1
/// §18.3.1.99), given the children's local names and spans.
pub(super) fn collection_position<'a>(
    children: impl Iterator<Item = (&'a [u8], usize)>,
    root_close: usize,
) -> usize {
    const LATER: [&[u8]; 4] = [b"controls", b"webPublishItems", b"tableParts", b"extLst"];
    children
        .filter(|(name, _)| LATER.contains(name))
        .map(|(_, start)| start)
        .min()
        .unwrap_or(root_close)
}

// ---------------------------------------------------------------------------
// `DrawingML` drawing

/// The hidden shape Excel keeps in the drawing part for an object, linked to
/// its VML shape (`a14:compatExt`), inside an `a14` choice with an empty
/// fallback.
pub(super) fn drawing_anchor(shape_id: u32, anchor: &Anchor) -> String {
    let mut out = String::with_capacity(1_024);
    let _ = write!(
        out,
        "<mc:AlternateContent xmlns:mc=\"{MC}\" xmlns:xdr=\"{XDR}\" xmlns:a=\"{A}\"><mc:Choice xmlns:a14=\"{A14}\" Requires=\"a14\"><xdr:twoCellAnchor editAs=\"oneCell\">"
    );
    for (name, marker) in [("from", &anchor.from), ("to", &anchor.to)] {
        let _ = write!(
            out,
            "<xdr:{name}><xdr:col>{}</xdr:col><xdr:colOff>{}</xdr:colOff><xdr:row>{}</xdr:row><xdr:rowOff>{}</xdr:rowOff></xdr:{name}>",
            marker.column, marker.column_offset, marker.row, marker.row_offset
        );
    }
    let _ = write!(
        out,
        "<xdr:sp macro=\"\" textlink=\"\"><xdr:nvSpPr><xdr:cNvPr id=\"{shape_id}\" name=\"Object {}\" hidden=\"1\"><a:extLst><a:ext uri=\"{COMPAT_EXT_URI}\"><a14:compatExt spid=\"_x0000_s{shape_id}\"/></a:ext></a:extLst></xdr:cNvPr><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr bwMode=\"auto\"><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></xdr:spPr></xdr:sp><xdr:clientData/></xdr:twoCellAnchor></mc:Choice><mc:Fallback/></mc:AlternateContent>",
        object_number(shape_id)
    );
    out
}

/// The number Excel gives an object in its drawing name (`Object 3` for
/// `_x0000_s1027`): its place in the sheet's block of 1024 shape ids.
fn object_number(shape_id: u32) -> u32 {
    (shape_id.saturating_sub(1) % 1024) + 1
}

/// The top-level children of a drawing part: their spans, the span holding
/// the `a14:compatExt` of `spid`, and every `cNvPr` id in use.
#[derive(Debug, Clone, Default)]
pub(super) struct DrawingScan {
    pub(super) children: usize,
    pub(super) object: Option<Range<usize>>,
    pub(super) ids: HashSet<u32>,
}

/// Scan a drawing part for the anchor of the VML shape `spid`.
pub(super) fn scan_drawing(xml: &[u8], spid: Option<&str>) -> Result<DrawingScan> {
    let mut reader = quick_xml::Reader::from_reader(xml);
    let mut depth = 0usize;
    let mut scan = DrawingScan::default();
    let mut child_start = 0usize;
    let mut matched = false;
    loop {
        let start = offset(&reader)?;
        let event = reader
            .read_event()
            .map_err(|error| invalid(format!("drawing markup: {error}")))?;
        let end = offset(&reader)?;
        match &event {
            Event::Start(element) | Event::Empty(element) => {
                let empty = matches!(event, Event::Empty(_));
                if depth == 1 {
                    child_start = start;
                    matched = false;
                    scan.children += 1;
                }
                match element.local_name().as_ref() {
                    b"cNvPr" => {
                        if let Some(id) =
                            attribute(element, b"id").and_then(|id| id.trim().parse().ok())
                        {
                            scan.ids.insert(id);
                        }
                    },
                    b"compatExt" if depth >= 2 => {
                        if spid.is_some_and(|spid| {
                            attribute(element, b"spid").is_some_and(|value| value == spid)
                        }) {
                            matched = true;
                        }
                    },
                    _ => {},
                }
                if depth == 1 && empty && matched && scan.object.is_none() {
                    scan.object = Some(child_start..end);
                }
                if !empty {
                    depth += 1;
                }
            },
            Event::End(_) => {
                depth = depth.saturating_sub(1);
                if depth == 1 && matched && scan.object.is_none() {
                    scan.object = Some(child_start..end);
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(scan)
}

// ---------------------------------------------------------------------------
// Legacy VML drawing

/// The VML shape of an object whose preview is the VML relationship
/// `relationship_id`.
pub(super) fn vml_shape(shape_id: u32, relationship_id: &str, frame: &Frame, z: usize) -> String {
    format!(
        "<v:shape id=\"_x0000_s{shape_id}\" type=\"#_x0000_t75\" style=\"position:absolute;margin-left:{};margin-top:{};width:{};height:{};z-index:{z}\"><v:imagedata o:relid=\"{}\" o:title=\"\"/><x:ClientData ObjectType=\"Pict\"><x:SizeWithCells/><x:Anchor>{}</x:Anchor><x:CF>Pict</x:CF><x:AutoPict/></x:ClientData></v:shape>",
        pt(frame.left_pt),
        pt(frame.top_pt),
        pt(frame.width_pt),
        pt(frame.height_pt),
        escape_xml(relationship_id),
        frame.vml_anchor()
    )
}

/// A new VML drawing for the worksheet at catalog position `sheet`, holding
/// `shape`: its shape-id block is `sheet + 1`, as Excel numbers them.
pub(super) fn new_vml_drawing(block: u32, shape: &str) -> String {
    format!(
        "<xml xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\"><o:shapelayout v:ext=\"edit\"><o:idmap v:ext=\"edit\" data=\"{block}\"/></o:shapelayout>{PICTURE_SHAPE_TYPE}{shape}</xml>",
    )
}

/// The value of attribute `name` in the start tag at the beginning of
/// `shape` (a VML shape's bytes), with its span; VML parts quote with `"`
/// or `'`.
fn tag_attribute(shape: &[u8], name: &[u8]) -> Option<(Range<usize>, String)> {
    let tag_end = shape.iter().position(|byte| *byte == b'>')?;
    let tag = &shape[..tag_end];
    let mut at = 0;
    while at + name.len() < tag.len() {
        let found = tag[at..]
            .windows(name.len())
            .position(|window| window == name)?
            + at;
        let before = found.checked_sub(1).and_then(|index| tag.get(index));
        let after = &tag[found + name.len()..];
        let rest = after
            .iter()
            .position(|byte| !byte.is_ascii_whitespace())
            .map(|skip| &after[skip..]);
        if before.is_some_and(|byte| byte.is_ascii_whitespace())
            && let Some(rest) = rest
            && rest.first() == Some(&b'=')
        {
            let value_at = found + name.len() + (after.len() - rest.len()) + 1;
            let quoted = &tag[value_at..];
            let skip = quoted.iter().position(|byte| !byte.is_ascii_whitespace())?;
            let quote = *quoted.get(skip)?;
            if quote != b'"' && quote != b'\'' {
                return None;
            }
            let start = value_at + skip + 1;
            let length = tag[start..].iter().position(|byte| *byte == quote)?;
            let value = String::from_utf8_lossy(&tag[start..start + length]).into_owned();
            return Some((start..start + length, value));
        }
        at = found + name.len();
    }
    None
}

/// The shape id `N` of a VML shape's `id="_x0000_sN"` (or `o:spid`).
pub(super) fn vml_shape_id(shape: &[u8]) -> Option<u32> {
    [b"o:spid".as_slice(), b"id"].into_iter().find_map(|name| {
        let (_, value) = tag_attribute(shape, name)?;
        value.trim().strip_prefix("_x0000_s")?.parse().ok()
    })
}

/// The `o:relid` of a VML shape's `v:imagedata`, with its span in `shape`.
pub(super) fn vml_image_relationship(shape: &[u8]) -> Option<(Range<usize>, String)> {
    let at = find(shape, b"<v:imagedata", 0)?;
    let (span, value) = tag_attribute(&shape[at..], b"o:relid")?;
    Some((at + span.start..at + span.end, value))
}

pub(super) fn find(haystack: &[u8], needle: &[u8], from: usize) -> Option<usize> {
    haystack
        .get(from..)?
        .windows(needle.len())
        .position(|window| window == needle)
        .map(|at| at + from)
}

/// `shape` (a VML object shape) resized to `frame`: the `width` and
/// `height` of its style and its `x:Anchor`. Other bytes are kept.
pub(super) fn resize_vml_shape(shape: &[u8], frame: &Frame) -> Vec<u8> {
    let mut edits: Vec<(Range<usize>, String)> = Vec::new();
    if let Some((span, style)) = tag_attribute(shape, b"style") {
        let mut seen = (false, false);
        let mut parts = style
            .split(';')
            .map(|part| {
                let key = part.split(':').next().unwrap_or("").trim();
                if key.eq_ignore_ascii_case("width") {
                    seen.0 = true;
                    format!("width:{}", pt(frame.width_pt))
                } else if key.eq_ignore_ascii_case("height") {
                    seen.1 = true;
                    format!("height:{}", pt(frame.height_pt))
                } else {
                    part.to_owned()
                }
            })
            .collect::<Vec<_>>();
        if !seen.0 {
            parts.push(format!("width:{}", pt(frame.width_pt)));
        }
        if !seen.1 {
            parts.push(format!("height:{}", pt(frame.height_pt)));
        }
        edits.push((span, parts.join(";")));
    }
    if let Some(open) = find(shape, b"<x:Anchor>", 0)
        && let Some(close) = find(shape, b"</x:Anchor>", open)
    {
        edits.push((open + b"<x:Anchor>".len()..close, frame.vml_anchor()));
    }
    splice(shape, edits)
}

/// `bytes` with each span replaced; spans must not overlap.
pub(super) fn splice(bytes: &[u8], mut edits: Vec<(Range<usize>, String)>) -> Vec<u8> {
    edits.sort_by_key(|(span, _)| span.start);
    let mut out = Vec::with_capacity(bytes.len() + 64);
    let mut at = 0;
    for (span, text) in edits {
        if span.start < at || span.end > bytes.len() {
            continue;
        }
        out.extend_from_slice(&bytes[at..span.start]);
        out.extend_from_slice(text.as_bytes());
        at = span.end;
    }
    out.extend_from_slice(&bytes[at..]);
    out
}

/// The text spans of the `col`, `colOff`, `row` and `rowOff` children of
/// every `to` element in `fragment`, with the value each takes from `to`.
pub(super) fn to_marker_edits(
    fragment: &[u8],
    to: &OleObjectMarker,
) -> Result<Vec<(Range<usize>, String)>> {
    let mut reader = quick_xml::Reader::from_reader(fragment);
    let mut in_to = 0usize;
    let mut field: Option<String> = None;
    let mut edits = Vec::new();
    let mut depth = 0usize;
    loop {
        let start = offset(&reader)?;
        let event = reader
            .read_event()
            .map_err(|error| invalid(format!("anchor markup: {error}")))?;
        let end = offset(&reader)?;
        match &event {
            Event::Start(element) => {
                depth += 1;
                let local = element.local_name();
                if local.as_ref() == b"to" {
                    in_to = depth;
                } else if in_to != 0 && depth == in_to + 1 {
                    field = match local.as_ref() {
                        b"col" => Some(to.column.to_string()),
                        b"colOff" => Some(to.column_offset.to_string()),
                        b"row" => Some(to.row.to_string()),
                        b"rowOff" => Some(to.row_offset.to_string()),
                        _ => None,
                    };
                }
            },
            Event::Text(_) => {
                if let Some(value) = field.take() {
                    edits.push((start..end, value));
                }
            },
            Event::End(_) => {
                if depth == in_to {
                    in_to = 0;
                }
                field = None;
                depth = depth.saturating_sub(1);
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(edits)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn default_grid_matches_excel_cells() {
        let grid = Grid::read(br#"<worksheet><sheetData/></worksheet>"#, 0).unwrap();
        // 64-pixel columns and 15-point rows.
        assert_eq!(grid.column_emu(0), 609_600);
        assert_eq!(grid.row_emu(0), 190_500);
        let frame = grid.frame(1, 1, (60.0, 20.0));
        assert_eq!(frame.anchor.to.column, 2);
        assert_eq!(frame.anchor.to.column_offset, 762_000 - 609_600);
        assert_eq!(frame.anchor.to.row, 2);
        assert_eq!(frame.anchor.to.row_offset, 254_000 - 190_500);
        assert_eq!((frame.left_pt, frame.top_pt), (48.0, 15.0));
    }

    #[test]
    fn grid_reads_column_widths_and_row_heights() {
        let xml = br#"<worksheet><sheetFormatPr defaultRowHeight="20"/><cols><col min="1" max="2" width="9.140625" customWidth="1"/><col min="3" max="3" hidden="1" width="20"/></cols><sheetData><row r="2" ht="30" customHeight="1"/></sheetData></worksheet>"#;
        let grid = Grid::read(xml, 0).unwrap();
        assert_eq!(grid.column_emu(0), 64 * EMU_PER_PIXEL);
        assert_eq!(grid.column_emu(2), 0);
        assert_eq!(grid.row_emu(0), 254_000);
        assert_eq!(grid.row_emu(1), 381_000);
    }

    #[test]
    fn collection_scan_finds_alternate_content_and_bare_entries() {
        let xml = br#"<w><oleObjects><mc:AlternateContent><mc:Choice Requires="x14"><oleObject shapeId="1025" r:id="rId1"><objectPr/></oleObject></mc:Choice><mc:Fallback><oleObject shapeId="1025" r:id="rId1"/></mc:Fallback></mc:AlternateContent><oleObject shapeId="1026" r:id="rId2"/></oleObjects></w>"#;
        let start = find(xml, b"<oleObjects>", 0).unwrap();
        let end = find(xml, b"</w>", 0).unwrap();
        let collection = Collection::scan(xml, start..end).unwrap();
        assert_eq!(collection.entries.len(), 2);
        assert_eq!(collection.entries[0].shape_id, Some(1025));
        assert_eq!(collection.entries[1].shape_id, Some(1026));
        let second = &xml[collection.entries[1].span.clone()];
        assert_eq!(second, br#"<oleObject shapeId="1026" r:id="rId2"/>"#);
        assert_eq!(
            &xml[collection.close.unwrap()..end],
            b"</oleObjects>".as_slice()
        );
    }

    #[test]
    fn vml_attributes_read_either_quote() {
        let shape = b"<v:shape id=\"_x0000_s1027\" type=\"#_x0000_t75\" style='position:absolute;\n  width:468pt;height:153pt;z-index:2'><v:imagedata o:relid=\"rId1\" o:title=\"\"/><x:ClientData ObjectType=\"Pict\"><x:Anchor>\n 0, 0, 11, 0, 2, 50, 20, 9</x:Anchor></x:ClientData></v:shape>";
        assert_eq!(vml_shape_id(shape), Some(1027));
        assert_eq!(vml_image_relationship(shape).unwrap().1, "rId1");
        let frame = Grid::read(b"<w/>", 0).unwrap().frame(0, 0, (10.0, 10.0));
        let resized = String::from_utf8(resize_vml_shape(shape, &frame)).unwrap();
        assert!(resized.contains("width:10pt;height:10pt;z-index:2'"));
        assert!(resized.contains("<x:Anchor>0, 0, 0, 0, 0, 13, 0, 13</x:Anchor>"));
    }

    #[test]
    fn to_marker_edits_touch_only_the_to_marker() {
        let fragment = b"<anchor><from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></from><to><xdr:col>3</xdr:col><xdr:colOff>5</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>6</xdr:rowOff></to></anchor>";
        let to = OleObjectMarker {
            column: 7,
            column_offset: 8,
            row: 9,
            row_offset: 10,
        };
        let edits = to_marker_edits(fragment, &to).unwrap();
        let out = String::from_utf8(splice(fragment, edits)).unwrap();
        assert!(out.contains("<from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff>"));
        assert!(out.contains("<to><xdr:col>7</xdr:col><xdr:colOff>8</xdr:colOff><xdr:row>9</xdr:row><xdr:rowOff>10</xdr:rowOff></to>"));
    }
}
