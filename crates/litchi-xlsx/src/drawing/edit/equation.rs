//! Office Math equations in worksheet drawings: add, replace and remove.
//!
//! A new equation is a two-cell anchor whose object is one
//! `mc:AlternateContent` block (ECMA-376 Part 3 for markup compatibility;
//! [MS-ODRAWXML] §2.3.3.41 for `a14:m`): the `mc:Choice Requires="a14"`
//! branch holds a text box whose paragraph holds the equation's OMML in
//! `a14:m`, and the `mc:Fallback` branch the same text box with the
//! equation's linear form (UnicodeMath, else LaTeX) as upright Cambria Math
//! text, for readers without Office Math. This follows Excel's known output
//! for an inserted equation; it was not checked against an Excel-authored
//! file in this repository, which has none with an equation.
//! [`Edit::add_equation`] writes that form (creating the drawing part and
//! its worksheet relationship when the worksheet has none),
//! [`Edit::set_equation`] replaces the OMML and the fallback of an existing
//! equation, keeping the fallback's kind (text, or the picture another
//! producer wrote), and [`Edit::remove_equation`] removes it with the
//! pictures and relationships nothing else uses. Every other byte of the
//! drawing part stays as it was.
//!
//! An equation is edited when its text box is the whole Choice branch of a
//! block around an anchor (`xdr:wsDr/mc:AlternateContent`) or around an
//! anchor's object (`xdr:*Anchor/mc:AlternateContent`), and the equation
//! sits directly in a paragraph of the text box. Other placements (a block
//! inside a paragraph, a Choice with several objects) are read but refused
//! for edits.

use std::collections::BTreeSet;
use std::ops::Range;

use litchi_core::edit::RefusalKind;
use litchi_drawingml::math::{
    A14, Bindings, DEFAULT_INSETS, EMU_PER_POINT, FALLBACK_FILL, Form, HostRun, MCE, Prepared,
    TRANSITIONAL_DML, carry_declarations, fallback_picture, fitted, horizontal, linear_text,
    max_raw_shape_id, render, require_declared_prefixes, root_bindings, tag_name, vertical,
};
use litchi_equation::Justification;
use litchi_opc::constants::relationship_type as rt;
use litchi_opc::{BlobPart, PackURI};
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use super::{Edit, Limits, Marker, Placement, objects};
use crate::drawing::math::{self, Block, Equation, EquationSource, EquationWrite, Level, Site};
use crate::error::{Error, Result, invalid};

/// The text size of a new equation when the caller names none, in points:
/// Excel's default for text boxes.
const DEFAULT_FONT_SIZE_PT: f64 = 11.0;
const XDR: &str = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const STRICT_XDR: &str = "http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing";
const STRICT_DML: &str = "http://purl.oclc.org/ooxml/drawingml/main";
const RELATIONSHIPS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_RELATIONSHIPS: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";

fn refused(operation: &'static str, reason: &'static str) -> Error {
    Error::refused(RefusalKind::UnsupportedContent, operation, reason)
}

/// The namespaces of a drawing part's conformance class: `SpreadsheetML`
/// drawing and `DrawingML`, and the image relationship type.
struct Namespaces {
    xdr: &'static str,
    dml: &'static str,
    image: &'static str,
}

impl Namespaces {
    /// Strict when the part's root binds the strict drawing namespace.
    fn of(bound: &Bindings) -> Self {
        if bound
            .iter()
            .any(|(_, uri)| uri.as_slice() == STRICT_XDR.as_bytes())
        {
            Self {
                xdr: STRICT_XDR,
                dml: STRICT_DML,
                image: rt::STRICT_IMAGE,
            }
        } else {
            Self {
                xdr: XDR,
                dml: TRANSITIONAL_DML,
                image: rt::IMAGE,
            }
        }
    }
}

/// The shape style Excel gives a text box (`xdr:style`).
const TEXT_BOX_STYLE: &str = concat!(
    r#"<xdr:style><a:lnRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:lnRef>"#,
    r#"<a:fillRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:fillRef>"#,
    r#"<a:effectRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:effectRef>"#,
    r#"<a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef></xdr:style>"#
);

/// The run of a text fallback that stands for an equation: its linear form
/// (UnicodeMath, else LaTeX) in upright Cambria Math, with the equation's
/// first host run properties (text size, language), else Excel's with
/// `size` (hundredths of a point).
fn linear_fallback_run(omml: &str, size: Option<u32>) -> String {
    let overrides = [
        ("b".to_owned(), "0".to_owned()),
        ("i".to_owned(), "0".to_owned()),
    ];
    let properties = HostRun::of(omml)
        .ok()
        .flatten()
        .unwrap_or_else(|| HostRun::native(size))
        .element(&overrides, "", false);
    format!(
        "<a:r>{properties}<a:t>{}</a:t></a:r>",
        quick_xml::escape::escape(linear_text(omml).as_str())
    )
}

/// `omml` without the namespace declaration `xmlns:{prefix}="{uri}"` on
/// its root element, when it makes one: the declaration is already in scope
/// where the OMML is written.
fn without_declaration(omml: &str, prefix: &str, uri: &str) -> String {
    let declaration = format!(r#" xmlns:{prefix}="{uri}""#);
    let root_end = omml.find('>').unwrap_or(0);
    match omml[..root_end].find(&declaration) {
        Some(at) => format!("{}{}", &omml[..at], &omml[at + declaration.len()..]),
        None => omml.to_owned(),
    }
}

/// The width and height, in EMU, of a text box that shows `omml` at
/// `points` inside `insets`: wide enough for the two-dimensional equation
/// and for its linear text in the Fallback, which readers without Office
/// Math show in the same frame (an estimate of 0.6 em per character).
fn text_frame(omml: &str, points: f64, insets: [i64; 4]) -> Result<(i64, i64)> {
    let options = litchi_equation::PngOptions::new(points)?;
    let (width_pt, height_pt) =
        litchi_equation::omml_size_pt(omml, &options, &litchi_equation::Limits::DEFAULT)?;
    #[allow(
        clippy::cast_precision_loss,
        reason = "an equation's linear text is far shorter than 2^52 characters"
    )]
    let linear_pt = linear_text(omml).chars().count() as f64 * points * 0.6;
    let (width_pt, height_pt) = (width_pt.max(linear_pt), height_pt.max(points * 1.2));
    let emu = |points: f64| -> i64 {
        #[allow(
            clippy::cast_possible_truncation,
            reason = "an equation is at most a few thousand points wide"
        )]
        let value = (points * EMU_PER_POINT).ceil() as i64;
        value.max(1)
    };
    let [left, top, right, bottom] = insets.map(|inset| inset.max(0));
    Ok((
        emu(width_pt).saturating_add(left).saturating_add(right),
        emu(height_pt).saturating_add(top).saturating_add(bottom),
    ))
}

/// EMU per pixel at 96 pixels per inch.
const EMU_PER_PIXEL: i64 = 9_525;
/// The widest column marker (`XFD`) and the last row, zero-based.
const LAST_COLUMN: u32 = 16_383;
const LAST_ROW: u32 = 1_048_575;

/// The sizes, in EMU, of a worksheet's columns and rows: the defaults of
/// `sheetFormatPr` and the explicit sizes of `col` and `row` (ECMA-376
/// Part 1, §18.3.1.13 and §18.3.1.73), at the maximum digit width of the
/// workbook's default font.
struct Grid {
    /// The maximum digit width, in pixels.
    digit: f64,
    default_column: i64,
    /// Column ranges and their widths, sorted by their first column.
    columns: Vec<(u32, u32, i64)>,
    default_row: i64,
    rows: std::collections::HashMap<u32, i64>,
}

impl Grid {
    /// The pixels of a column `width` characters wide (ECMA-376 Part 1,
    /// §18.3.1.13: `Truncate(((256 * width + Truncate(128 / digit)) / 256) *
    /// digit)`).
    fn column_pixels(&self, width: f64) -> i64 {
        let width = width.clamp(0.0, 255.0);
        #[allow(
            clippy::cast_possible_truncation,
            reason = "column widths are at most 255 characters"
        )]
        let pixels =
            (((256.0 * width + (128.0 / self.digit).trunc()) / 256.0) * self.digit).trunc() as i64;
        pixels
    }

    /// The maximum digit width, in pixels at 96 pixels per inch, of the
    /// default font of a styles part (its first `font`): Excel's value for
    /// the common default fonts, otherwise an estimate from the text size
    /// (about 0.48 em). 7 pixels without a styles part.
    fn digit_width(styles: Option<&[u8]>) -> f64 {
        const DEFAULT: f64 = 7.0;
        /// Maximum digit widths of default fonts at their sizes, in pixels.
        const KNOWN: [(&str, f64, f64); 6] = [
            ("Calibri", 11.0, 7.0),
            ("Calibri", 12.0, 8.0),
            ("Arial", 10.0, 7.0),
            ("Verdana", 10.0, 8.0),
            ("Aptos Narrow", 11.0, 7.0),
            ("Aptos Narrow", 12.0, 8.0),
        ];
        let Some((name, size)) = styles.and_then(Self::default_font) else {
            return DEFAULT;
        };
        let Some(size) = size else {
            return DEFAULT;
        };
        if let Some(name) = &name
            && let Some((_, _, digit)) = KNOWN.iter().find(|(known, known_size, _)| {
                known.eq_ignore_ascii_case(name.trim()) && (*known_size - size).abs() < 0.01
            })
        {
            return *digit;
        }
        (size * 96.0 / 72.0 * 0.48).round().max(1.0)
    }

    /// The name and size of the first `font` of a styles part.
    fn default_font(styles: &[u8]) -> Option<(Option<String>, Option<f64>)> {
        use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
        let value = |element: &quick_xml::events::BytesStart<'_>| -> Option<String> {
            element
                .checked_attributes()
                .flatten()
                .find(|attribute| attribute.key.local_name().as_ref() == b"val")
                .and_then(|attribute| String::from_utf8(attribute.value.into_owned()).ok())
        };
        let mut reader = quick_xml::Reader::from_reader(styles);
        let mut font: Option<(Option<String>, Option<f64>)> = None;
        loop {
            match reader.read_event() {
                Ok(Event::Start(ref element)) if element.local_name().as_ref() == b"font" => {
                    font = Some((None, None));
                },
                Ok(Event::End(ref element)) if element.local_name().as_ref() == b"font" => {
                    return font;
                },
                Ok(Event::Empty(ref element) | Event::Start(ref element)) => {
                    if let Some((name, size)) = font.as_mut() {
                        match element.local_name().as_ref() {
                            b"sz" => {
                                *size = value(element)
                                    .and_then(|text| text.trim().parse::<f64>().ok())
                                    .filter(|size| {
                                        size.is_finite() && (1.0..=409.0).contains(size)
                                    });
                            },
                            b"name" => *name = value(element),
                            _ => {},
                        }
                    }
                },
                Ok(Event::Eof) | Err(_) => return font,
                Ok(_) => {},
            }
        }
    }

    fn points(value: f64) -> i64 {
        #[allow(
            clippy::cast_possible_truncation,
            reason = "row heights are at most 409 points"
        )]
        let emu = (value.clamp(0.0, 409.5) * EMU_PER_POINT).round() as i64;
        emu
    }

    fn of(worksheet: &[u8], styles: Option<&[u8]>) -> Result<Self> {
        use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
        let number = |element: &quick_xml::events::BytesStart<'_>, name: &[u8]| -> Option<f64> {
            element
                .checked_attributes()
                .flatten()
                .find(|attribute| attribute.key.local_name().as_ref() == name)
                .and_then(|attribute| {
                    std::str::from_utf8(&attribute.value)
                        .ok()?
                        .trim()
                        .parse()
                        .ok()
                })
                .filter(|value: &f64| value.is_finite())
        };
        let flag = |element: &quick_xml::events::BytesStart<'_>, name: &[u8]| {
            number(element, name).is_some_and(|value| value != 0.0)
                || element.checked_attributes().flatten().any(|attribute| {
                    attribute.key.local_name().as_ref() == name
                        && attribute.value.as_ref() == b"true"
                })
        };
        // Excel's default column: 8 characters plus 5 pixels of padding,
        // rounded up to a multiple of 8 pixels; its default row: 15 points.
        let digit = Self::digit_width(styles);
        #[allow(
            clippy::cast_possible_truncation,
            reason = "the digit width is at most a few hundred pixels"
        )]
        let base = (8.0 * digit) as i64;
        let mut grid = Self {
            digit,
            default_column: (base + 5 + 7) / 8 * 8 * EMU_PER_PIXEL,
            columns: Vec::new(),
            default_row: Self::points(15.0),
            rows: std::collections::HashMap::new(),
        };
        let mut reader = quick_xml::Reader::from_reader(worksheet);
        let mut next_row = 0u32;
        loop {
            match reader
                .read_event()
                .map_err(|error| invalid(format!("worksheet markup: {error}")))?
            {
                Event::Start(ref element) | Event::Empty(ref element) => {
                    match element.local_name().as_ref() {
                        b"sheetFormatPr" => {
                            if let Some(width) = number(element, b"defaultColWidth") {
                                grid.default_column = grid.column_pixels(width) * EMU_PER_PIXEL;
                            } else if let Some(base) = number(element, b"baseColWidth") {
                                let pixels = (grid.column_pixels(base) + 5 + 7) / 8 * 8;
                                grid.default_column = pixels * EMU_PER_PIXEL;
                            }
                            if let Some(height) = number(element, b"defaultRowHeight") {
                                grid.default_row = Self::points(height);
                            }
                        },
                        b"col" => {
                            let (Some(min), Some(max)) =
                                (number(element, b"min"), number(element, b"max"))
                            else {
                                continue;
                            };
                            #[allow(
                                clippy::cast_possible_truncation,
                                clippy::cast_sign_loss,
                                reason = "column numbers are clamped to the sheet"
                            )]
                            let (first, last) = (
                                (min.clamp(1.0, 16_384.0) as u32) - 1,
                                (max.clamp(1.0, 16_384.0) as u32) - 1,
                            );
                            let width = if flag(element, b"hidden") {
                                0
                            } else {
                                number(element, b"width").map_or(grid.default_column, |width| {
                                    grid.column_pixels(width) * EMU_PER_PIXEL
                                })
                            };
                            grid.columns.push((first, last, width));
                        },
                        b"row" => {
                            #[allow(
                                clippy::cast_possible_truncation,
                                clippy::cast_sign_loss,
                                reason = "row numbers are clamped to the sheet"
                            )]
                            let row = number(element, b"r")
                                .map_or(next_row, |r| (r.clamp(1.0, 1_048_576.0) as u32) - 1);
                            next_row = row.saturating_add(1);
                            if flag(element, b"hidden") {
                                grid.rows.insert(row, 0);
                            } else if let Some(height) = number(element, b"ht") {
                                grid.rows.insert(row, Self::points(height));
                            }
                        },
                        _ => {},
                    }
                },
                Event::Eof => break,
                _ => {},
            }
        }
        grid.columns.sort_by_key(|(first, _, _)| *first);
        Ok(grid)
    }

    fn column(&self, column: u32) -> i64 {
        // The last range that starts at or before the column.
        let at = self
            .columns
            .partition_point(|(first, _, _)| *first <= column);
        at.checked_sub(1)
            .map(|at| self.columns[at])
            .filter(|(_, last, _)| column <= *last)
            .map_or(self.default_column, |(_, _, width)| width)
    }

    fn row(&self, row: u32) -> i64 {
        self.rows.get(&row).copied().unwrap_or(self.default_row)
    }

    /// The cell and offset `length` EMU after `offset` EMU into `start`.
    fn reach(
        start: u32,
        offset: i64,
        length: i64,
        last: u32,
        size: impl Fn(u32) -> i64,
    ) -> (u32, i64) {
        let mut index = start;
        let mut remaining = offset.max(0).saturating_add(length);
        loop {
            let extent = size(index);
            if remaining < extent || index == last {
                return (index, remaining.clamp(0, extent.max(0)));
            }
            remaining -= extent;
            index += 1;
        }
    }

    /// The frame's offset on the sheet (an estimate readers ignore in favour
    /// of the anchor) and the `to` marker of a frame of `size` EMU from
    /// `from`.
    fn span(&self, from: Marker, size: (i64, i64)) -> ((i64, i64), Marker) {
        let x = (0..from.col).map(|column| self.column(column)).sum::<i64>() + from.col_offset;
        let y = if self.rows.is_empty() {
            i64::from(from.row) * self.default_row
        } else {
            (0..from.row).map(|row| self.row(row)).sum::<i64>()
        } + from.row_offset;
        let (col, col_offset) =
            Self::reach(from.col, from.col_offset, size.0, LAST_COLUMN, |column| {
                self.column(column)
            });
        let (row, row_offset) = Self::reach(from.row, from.row_offset, size.1, LAST_ROW, |row| {
            self.row(row)
        });
        (
            (x, y),
            Marker {
                col,
                col_offset,
                row,
                row_offset,
            },
        )
    }
}

/// The relationship ID a staged edit is checked with before its picture
/// part exists; the real ID replaces it once the checks pass.
const PENDING_RELATIONSHIP: &str = "rIdPendingEquationPicture";

/// Replacements of byte ranges of one part, applied in one pass.
#[derive(Default, Clone)]
struct Splice {
    edits: Vec<(Range<usize>, String)>,
}

impl Splice {
    fn replace(&mut self, range: Range<usize>, with: impl Into<String>) {
        self.edits.push((range, with.into()));
    }

    fn apply(mut self, xml: &[u8], max_bytes: usize) -> Result<Vec<u8>> {
        self.edits.sort_by_key(|(range, _)| range.start);
        let size = self.edits.iter().fold(xml.len(), |size, (range, with)| {
            size.saturating_sub(range.len()).saturating_add(with.len())
        });
        if size > max_bytes {
            return Err(Error::limit(
                super::LimitKind::DrawingBytes,
                size,
                max_bytes,
            ));
        }
        let mut out = Vec::new();
        out.try_reserve_exact(size)
            .map_err(|source| Error::Allocation {
                resource: "worksheet drawing equation XML",
                source,
            })?;
        let mut at = 0usize;
        for (range, with) in &self.edits {
            if range.start < at || range.end > xml.len() {
                return Err(invalid("equation edits overlap"));
            }
            out.extend_from_slice(&xml[at..range.start]);
            out.extend_from_slice(with.as_bytes());
            at = range.end;
        }
        out.extend_from_slice(&xml[at..]);
        Ok(out)
    }
}

/// The value span of attribute `name` in the start tag `tag`.
fn attribute_span(tag: &[u8], name: &[u8]) -> Option<Range<usize>> {
    attribute_span_by(tag, |key| key == name)
}

/// The value span of the first attribute of the start tag `tag` whose
/// qualified name `matches`.
fn attribute_span_by(tag: &[u8], matches: impl Fn(&[u8]) -> bool) -> Option<Range<usize>> {
    let mut at = tag
        .iter()
        .position(|byte| byte.is_ascii_whitespace() || *byte == b'/' || *byte == b'>')?;
    loop {
        while tag.get(at).is_some_and(u8::is_ascii_whitespace) {
            at += 1;
        }
        let key_start = at;
        while tag.get(at).is_some_and(|byte| {
            *byte != b'=' && !byte.is_ascii_whitespace() && *byte != b'>' && *byte != b'/'
        }) {
            at += 1;
        }
        if at == key_start {
            return None;
        }
        let key = &tag[key_start..at];
        while tag.get(at).is_some_and(u8::is_ascii_whitespace) {
            at += 1;
        }
        if tag.get(at) != Some(&b'=') {
            return None;
        }
        at += 1;
        while tag.get(at).is_some_and(u8::is_ascii_whitespace) {
            at += 1;
        }
        let quote = *tag.get(at)?;
        if quote != b'"' && quote != b'\'' {
            return None;
        }
        let value_start = at + 1;
        let value_end = value_start + tag[value_start..].iter().position(|byte| *byte == quote)?;
        if matches(key) {
            return Some(value_start..value_end);
        }
        at = value_end + 1;
    }
}

/// The `cx` and `cy` of an extent start tag.
fn extent(tag: &[u8]) -> Option<(i64, i64)> {
    let value = |name: &[u8]| -> Option<i64> {
        let span = attribute_span(tag, name)?;
        std::str::from_utf8(&tag[span]).ok()?.trim().parse().ok()
    };
    Some((value(b"cx")?, value(b"cy")?)).filter(|(width, height)| *width > 0 && *height > 0)
}

/// Replace the `cx` and `cy` values of the extent start tag at `tag`.
fn resize(splice: &mut Splice, xml: &[u8], tag: &Range<usize>, size: (i64, i64)) -> Result<()> {
    let bytes = &xml[tag.clone()];
    for (name, value) in [(b"cx".as_slice(), size.0), (b"cy", size.1)] {
        let span = attribute_span(bytes, name)
            .ok_or_else(|| invalid("the equation frame has no extent"))?;
        splice.replace(
            tag.start + span.start..tag.start + span.end,
            value.to_string(),
        );
    }
    Ok(())
}

/// The relationship IDs `xml` refers to (attributes in the relationship
/// namespace, transitional or strict).
fn relationship_ids(xml: &[u8]) -> Result<BTreeSet<String>> {
    use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
    let mut reader = NsReader::from_reader(xml);
    let mut out = BTreeSet::new();
    loop {
        let (_, event) = reader
            .read_resolved_event()
            .map_err(|error| invalid(format!("drawing equations: {error}")))?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                for attribute in element.checked_attributes().flatten() {
                    let (namespace, _) = reader.resolver().resolve_attribute(attribute.key);
                    if matches!(namespace, ResolveResult::Bound(Namespace(uri))
                        if uri == RELATIONSHIPS.as_bytes() || uri == STRICT_RELATIONSHIPS.as_bytes())
                    {
                        out.insert(String::from_utf8_lossy(&attribute.value).into_owned());
                    }
                }
            },
            Event::Eof => return Ok(out),
            _ => {},
        }
    }
}

/// The equation at `index` and the block holding it, refused when the
/// equation is not in a text box that is the whole Choice branch of a
/// root- or anchor-level block.
fn located(
    sites: &[Site],
    blocks: &[Block],
    index: usize,
    operation: &'static str,
) -> Result<(Site, usize)> {
    let site = sites.get(index).ok_or_else(|| {
        Error::refused(
            RefusalKind::NotFound,
            operation,
            "the worksheet has no such equation",
        )
    })?;
    let block = site.block.ok_or_else(|| {
        refused(
            operation,
            "only an equation in a text box of an anchor-level markup-compatibility block is edited",
        )
    })?;
    let record = &blocks[block];
    if record.nested
        || record.choices != 1
        || record.choice_shape.is_none()
        || record.choice_children != 1
        || record.level == Level::Inner
    {
        return Err(refused(
            operation,
            "the equation's block holds more than its text box",
        ));
    }
    Ok((site.clone(), block))
}

/// Whether the block's text box holds this one equation and no other text.
const fn equation_only(block: &Block) -> bool {
    block.equations == 1 && !block.other_content
}

/// The edit that gives block `index` of `xml` a text Fallback: a copy of
/// its Choice branch in which every equation is a run of its linear form,
/// with no picture fill. Office ignores the Fallback; other readers show
/// the text.
fn text_fallback(xml: &[u8], index: usize) -> Result<(Range<usize>, String)> {
    let (sites, blocks) = math::scan(xml)?;
    let block = blocks
        .get(index)
        .ok_or_else(|| invalid("the equation block disappeared"))?;
    let content = block
        .choice_content
        .clone()
        .ok_or_else(|| invalid("the equation block has no Choice branch"))?;
    let mut copy = Splice::default();
    for site in sites
        .iter()
        .filter(|site| content.start <= site.element.start && site.element.end <= content.end)
    {
        let omml = std::str::from_utf8(&xml[site.content.clone()])
            .map_err(|error| invalid(error.to_string()))?;
        copy.replace(
            site.element.start - content.start..site.element.end - content.start,
            linear_fallback_run(omml, None),
        );
    }
    let copy = copy.apply(&xml[content.clone()], usize::MAX)?;
    // The branch's object keeps the declarations of the block and Choice
    // start tags, which do not enclose the Fallback.
    let copy = carry_declarations(&xml[block.range.start..content.start], trimmed(&copy))?;
    let block_name = tag_name(&xml[block.range.clone()])?;
    match (&block.fallback, &block.fallback_tag) {
        (Some((range, empty)), Some(tag)) => {
            let name = tag_name(&xml[tag.clone()])?;
            let open = std::str::from_utf8(&xml[tag.clone()])
                .map_err(|error| invalid(error.to_string()))?;
            let open = if *empty {
                let trimmed = open.trim_end_matches('>').trim_end().trim_end_matches('/');
                format!("{trimmed}>")
            } else {
                open.to_owned()
            };
            Ok((range.clone(), format!("{open}{copy}</{name}>")))
        },
        _ => {
            // No Fallback yet: one goes before the block's end tag, under
            // the block's own prefix.
            let prefix = block_name
                .split_once(':')
                .map_or(String::new(), |(prefix, _)| format!("{prefix}:"));
            let close = format!("</{block_name}>");
            let at = block
                .range
                .end
                .checked_sub(close.len())
                .filter(|at| xml[*at..block.range.end] == *close.as_bytes())
                .ok_or_else(|| invalid("the equation block is malformed"))?;
            Ok((
                at..at,
                format!("<{prefix}Fallback>{copy}</{prefix}Fallback>"),
            ))
        },
    }
}

/// `bytes` without surrounding XML whitespace.
fn trimmed(bytes: &[u8]) -> &[u8] {
    let start = bytes
        .iter()
        .position(|byte| !byte.is_ascii_whitespace())
        .unwrap_or(bytes.len());
    let end = bytes
        .iter()
        .rposition(|byte| !byte.is_ascii_whitespace())
        .map_or(start, |at| at + 1);
    &bytes[start..end]
}

/// The anchors a block's frame is written in: the anchors inside a
/// root-level block, or the anchor around an anchor-level one. `None` when
/// one of them has no extent or `to` marker to rewrite.
fn frame_anchors<'a>(block: &Block, anchors: &'a [math::Anchor]) -> Option<Vec<&'a math::Anchor>> {
    let involved: Vec<&math::Anchor> = match block.level {
        Level::Root => anchors
            .iter()
            .filter(|anchor| {
                block.range.start <= anchor.range.start && anchor.range.end <= block.range.end
            })
            .collect(),
        Level::Anchor => block
            .outer_anchor
            .and_then(|index| anchors.get(index))
            .into_iter()
            .collect(),
        Level::Inner => Vec::new(),
    };
    let rewritable = |anchor: &&math::Anchor| {
        if anchor.two_cell() {
            anchor.to_marker.is_some()
        } else {
            anchor.extent_tag.is_some()
        }
    };
    (!involved.is_empty() && involved.iter().all(rewritable)).then_some(involved)
}

/// Give a block's frame `size` (EMU): both branches' text box extents, and
/// its anchors' extents or, for a two-cell anchor, its `to` marker across
/// the worksheet's grid.
fn resize_frame(
    splice: &mut Splice,
    xml: &[u8],
    block: &Block,
    anchors: &[&math::Anchor],
    grid: &Grid,
    size: (i64, i64),
) -> Result<()> {
    for tag in [&block.choice_extent, &block.fallback_extent]
        .into_iter()
        .flatten()
    {
        resize(splice, xml, tag, size)?;
    }
    for anchor in anchors {
        if anchor.two_cell() {
            let range = anchor
                .to_marker
                .clone()
                .ok_or_else(|| invalid("the equation anchor has no end marker"))?;
            let name = tag_name(&xml[range.clone()])?;
            let prefix = name
                .split_once(':')
                .map_or(String::new(), |(prefix, _)| format!("{prefix}:"));
            let (_, to) = grid.span(anchor.from(), size);
            splice.replace(
                range,
                format!(
                    "<{name}><{prefix}col>{}</{prefix}col><{prefix}colOff>{}</{prefix}colOff><{prefix}row>{}</{prefix}row><{prefix}rowOff>{}</{prefix}rowOff></{name}>",
                    to.col, to.col_offset, to.row, to.row_offset
                ),
            );
        } else if let Some(tag) = &anchor.extent_tag {
            resize(splice, xml, tag, size)?;
        }
    }
    Ok(())
}

impl Edit {
    /// The staged equations of the worksheet at `sheet`, in drawing order.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing sheet, a chartsheet, or a malformed
    /// drawing.
    pub fn equations(&self, sheet: usize) -> Result<Vec<Equation>> {
        let worksheet = self.staged.worksheet(sheet)?;
        match objects::drawing_of(self.staged.opc().get_part(&worksheet)?.rels())? {
            Some(drawing) => math::equations(self.staged.blob(&drawing)?),
            None => Ok(Vec::new()),
        }
    }

    /// The column and row sizes of `worksheet`, at the digit width of the
    /// workbook's default font.
    fn grid(&self, worksheet: &PackURI) -> Result<Grid> {
        let styles = match self.staged.styles() {
            Some(styles) => self.staged.blob(&styles).ok(),
            None => None,
        };
        Grid::of(self.staged.blob(worksheet)?, styles)
    }

    /// The OMML of `source` for the drawing part `xml`, refused when the
    /// caller's text exceeds the drawing byte bound.
    fn prepare_equation(
        source: EquationSource<'_>,
        form: Form,
        run: &HostRun,
        dml: &str,
        bound: &Bindings,
        limits: &Limits,
    ) -> Result<Prepared> {
        let input = source.text();
        if input.len() > limits.max_drawing_bytes {
            return Err(Error::limit(
                super::LimitKind::DrawingBytes,
                input.len(),
                limits.max_drawing_bytes,
            ));
        }
        Ok(litchi_drawingml::math::prepare(
            source, form, run, dml, bound,
        )?)
    }

    /// Stage `next` as the drawing part `drawing`, after checking that it
    /// declares every prefix it uses, that its objects still read and that
    /// it holds `equations` equations; the relationships `before` referred to
    /// and `next` no longer does are removed, with the parts nothing else
    /// uses.
    fn publish_equation_edit(
        &mut self,
        drawing: &PackURI,
        before: &[u8],
        next: Vec<u8>,
        equations: usize,
        limits: &Limits,
    ) -> Result<()> {
        self.check_equation_xml(drawing, &next, equations, limits)?;
        let released: Vec<String> = relationship_ids(before)?
            .difference(&relationship_ids(&next)?)
            .cloned()
            .collect();
        self.staged.replace_blob(drawing, next)?;
        let mut targets = Vec::new();
        {
            let opc = self.staged.opc_mut()?;
            let part = opc.get_part_mut(drawing)?;
            for id in released {
                if let Some(relationship) = part.rels().get(&id)
                    && !relationship.is_external()
                    && let Ok(target) = relationship.target_partname()
                {
                    targets.push(target);
                }
                part.rels_mut().remove(&id);
            }
        }
        self.staged.touch(drawing);
        for target in targets {
            let opc = self.staged.opc_mut()?;
            let referenced = opc.iter_parts().any(|part| {
                part.rels().iter().any(|relationship| {
                    !relationship.is_external()
                        && relationship.target_partname().ok().as_ref() == Some(&target)
                })
            });
            if !referenced {
                opc.remove_part(&target);
                self.staged.touch(&target);
            }
        }
        Ok(())
    }

    /// Check the drawing XML an equation edit stages: it declares every
    /// prefix it uses, its objects still read, and it holds `equations`
    /// equations.
    fn check_equation_xml(
        &self,
        drawing: &PackURI,
        next: &[u8],
        equations: usize,
        limits: &Limits,
    ) -> Result<()> {
        require_declared_prefixes(next)?;
        objects::read::<crate::color::Rgb>(
            next,
            self.staged.opc().get_part(drawing)?.rels(),
            |_| None,
            limits,
        )?;
        if math::scan(next)?.0.len() != equations {
            return Err(invalid("the equation edit did not round-trip"));
        }
        Ok(())
    }

    /// Stage `png` as a new media part related to `drawing`; returns the
    /// relationship ID.
    fn stage_picture(&mut self, drawing: &PackURI, png: Vec<u8>, image: &str) -> Result<String> {
        let media = self.staged.free_name("/xl/media/image", ".png")?;
        let opc = self.staged.opc_mut()?;
        opc.add_part(Box::new(BlobPart::new(
            media.clone(),
            "image/png".to_owned(),
            png,
        )));
        let relationship = opc
            .get_part_mut(drawing)?
            .relate_to(&media.relative_ref(drawing.base_uri()), image);
        self.staged.touch(&media);
        self.staged.touch(drawing);
        Ok(relationship)
    }

    /// Add an equation as a new text box from `from`: a two-cell anchor
    /// around a markup-compatibility block whose Choice holds the OMML in a
    /// display paragraph and whose Fallback holds its linear text (see the
    /// module documentation for the form and its provenance). The worksheet's
    /// drawing part is created when it has none; no picture part is written.
    ///
    /// The anchor spans the equation's size at `font_size_pt` (11 when
    /// `None`, which is also written on the equation's runs) across the
    /// worksheet's column widths and row heights.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Equation`] for LaTeX or OMML that does not convert
    /// or render, an error for malformed OMML, a missing sheet or a
    /// chartsheet, an invalid anchor, a malformed drawing or exceeded
    /// bounds, and [`Error::Signed`] for a signed workbook.
    pub fn add_equation(
        &mut self,
        sheet: usize,
        source: EquationSource<'_>,
        from: Marker,
        font_size_pt: Option<f64>,
    ) -> Result<EquationWrite> {
        let points = font_size_pt.unwrap_or(DEFAULT_FONT_SIZE_PT);
        // ST_TextFontSize: 1 to 4000 points in hundredths.
        if !(points.is_finite() && (1.0..=4000.0).contains(&points)) {
            return Err(invalid("the equation text size must be 1 to 4000 points"));
        }
        #[allow(
            clippy::cast_possible_truncation,
            clippy::cast_sign_loss,
            reason = "the size is checked to lie in 1..=4000 points"
        )]
        let size = (points * 100.0).round() as u32;
        Placement::OneCell { from, cx: 1, cy: 1 }.validate()?;
        let limits = Limits::from_source_limits(self.staged.opc().read_limits());
        let worksheet = self.staged.worksheet(sheet)?;
        let existing = objects::drawing_of(self.staged.opc().get_part(&worksheet)?.rels())?;
        let current: Vec<u8> = match &existing {
            Some(drawing) => self.staged.blob(drawing)?.to_vec(),
            None => objects::EMPTY_DRAWING.as_bytes().to_vec(),
        };
        let bound = root_bindings(&current)?;
        let spaces = Namespaces::of(&bound);
        let run = HostRun::native(Some(size));
        let prepared = Self::prepare_equation(
            source,
            Form::Display(Some(Justification::CenterGroup)),
            &run,
            spaces.dml,
            &bound,
            &limits,
        )?;
        // The text box's frame: the equation's size inside the text insets.
        let size_emu = text_frame(&prepared.omml, points, DEFAULT_INSETS)?;
        let grid = self.grid(&worksheet)?;
        let (origin, to) = grid.span(from, size_emu);
        let placement = Placement::TwoCell { from, to };
        placement.validate()?;
        let shape_id = max_raw_shape_id(&current)?
            .checked_add(1)
            .ok_or_else(|| invalid("worksheet drawing shape ID overflow"))?;
        let before = math::scan(&current)?.0.len();
        let name = format!("TextBox {}", shape_id - 1);
        let shape = |paragraph: &str| {
            format!(
                concat!(
                    r#"<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="{id}" name="{name}"/><xdr:cNvSpPr txBox="1"/></xdr:nvSpPr>"#,
                    r#"<xdr:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{width}" cy="{height}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></xdr:spPr>"#,
                    "{style}",
                    r#"<xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip" wrap="none" rtlCol="0" anchor="t"><a:spAutoFit/></a:bodyPr><a:lstStyle/>"#,
                    r#"<a:p>{paragraph}<a:endParaRPr lang="en-US" sz="{size}"/></a:p></xdr:txBody></xdr:sp>"#,
                ),
                id = shape_id,
                name = name,
                x = origin.0,
                y = origin.1,
                width = size_emu.0,
                height = size_emu.1,
                style = TEXT_BOX_STYLE,
                paragraph = paragraph,
                size = size,
            )
        };
        // The anchor declares the `a` prefix the converter declares again.
        let omml = without_declaration(&prepared.omml, "a", spaces.dml);
        let choice = shape(&format!("<a14:m>{omml}</a14:m>"));
        let fallback = shape(&linear_fallback_run(&omml, Some(size)));
        let fragment = format!(
            concat!(
                r#"<xdr:twoCellAnchor xmlns:xdr="{xdr}" xmlns:a="{dml}">"#,
                "<xdr:from><xdr:col>{from_col}</xdr:col><xdr:colOff>{from_col_offset}</xdr:colOff><xdr:row>{from_row}</xdr:row><xdr:rowOff>{from_row_offset}</xdr:rowOff></xdr:from>",
                "<xdr:to><xdr:col>{to_col}</xdr:col><xdr:colOff>{to_col_offset}</xdr:colOff><xdr:row>{to_row}</xdr:row><xdr:rowOff>{to_row_offset}</xdr:rowOff></xdr:to>",
                r#"<mc:AlternateContent xmlns:mc="{mce}"><mc:Choice xmlns:a14="{a14}" Requires="a14">{choice}</mc:Choice><mc:Fallback>{fallback}</mc:Fallback></mc:AlternateContent>"#,
                "<xdr:clientData/></xdr:twoCellAnchor>",
            ),
            xdr = spaces.xdr,
            dml = spaces.dml,
            from_col = from.col,
            from_col_offset = from.col_offset,
            from_row = from.row,
            from_row_offset = from.row_offset,
            to_col = to.col,
            to_col_offset = to.col_offset,
            to_row = to.row,
            to_row_offset = to.row_offset,
            mce = MCE,
            a14 = A14,
            choice = choice,
            fallback = fallback,
        );
        // Check the new part before anything is staged.
        let probe = objects::append(&current, &fragment, &limits)?;
        require_declared_prefixes(&probe)?;
        let after = math::equations(&probe)?;
        if after.len() != before + 1 || after[before].shape_id() != Some(shape_id) {
            return Err(invalid("the equation shape did not round-trip"));
        }
        let drawing = self.drawing(sheet)?;
        let next = objects::append(self.staged.blob(&drawing)?, &fragment, &limits)?;
        self.staged.replace_blob(&drawing, next)?;
        Ok(EquationWrite {
            shape_id,
            index: before,
            losses: prepared.losses,
        })
    }

    /// Replace the content of the worksheet's equation at `equation` (its
    /// position among [`crate::Package::equations`]). The Choice branch gets
    /// the new OMML. When the text box holds only this equation and its
    /// Fallback shows a picture, the picture is redrawn; otherwise the
    /// Fallback becomes a copy of the Choice branch with every equation as a
    /// run of its linear form. A text box that shows only this equation and
    /// resizes to its text (`a:spAutoFit`) takes the new size in both
    /// branches, with its anchors' extents or, for a two-cell anchor, its
    /// end marker across the worksheet's column widths and row heights. The
    /// shape keeps its ID, start cell and other properties; pictures and
    /// relationships nothing else uses are removed. LaTeX keeps the
    /// equation's form (inline or display, with its justification) and its
    /// first run's host properties (text size, fonts). Returns `None` when
    /// the OMML is unchanged.
    ///
    /// # Errors
    ///
    /// Returns a not-found refusal for a missing equation, an
    /// unsupported-content refusal for an equation outside an editable
    /// block, [`Error::Equation`] for input that does not convert or render,
    /// and an error for malformed OMML or drawing XML.
    pub fn set_equation(
        &mut self,
        sheet: usize,
        equation: usize,
        source: EquationSource<'_>,
    ) -> Result<Option<EquationWrite>> {
        const OPERATION: &str = "set_equation";
        let drawing = self.equation_drawing(sheet, OPERATION)?;
        let limits = Limits::from_source_limits(self.staged.opc().read_limits());
        let xml = self.staged.blob(&drawing)?.to_vec();
        let (sites, blocks, anchors) = math::scan_all(&xml)?;
        let (site, index) = located(&sites, &blocks, equation, OPERATION)?;
        let block = blocks[index].clone();
        let bound = root_bindings(&xml)?;
        let spaces = Namespaces::of(&bound);
        let old = std::str::from_utf8(&xml[site.content.clone()])
            .map_err(|error| invalid(error.to_string()))?;
        let run = HostRun::of(old)?.unwrap_or_else(|| HostRun::native(None));
        let prepared =
            Self::prepare_equation(source, Form::of(old)?, &run, spaces.dml, &bound, &limits)?;
        let mut prepared = prepared;
        if bound.iter().any(|(prefix, uri)| {
            prefix.as_slice() == b"a" && uri.as_slice() == spaces.dml.as_bytes()
        }) {
            prepared.omml = without_declaration(&prepared.omml, "a", spaces.dml);
        }
        if prepared.omml == old.trim() {
            return Ok(None);
        }
        let shape_id = litchi_drawingml::math::cnvpr_id(
            &xml[block
                .choice_shape
                .clone()
                .ok_or_else(|| invalid("the equation block has no text box"))?],
        )?
        .ok_or_else(|| invalid("the equation shape has no ID"))?;
        let mut splice = Splice::default();
        let worksheet = self.staged.worksheet(sheet)?;
        let framed = frame_anchors(&block, &anchors);
        let picture_fallback = equation_only(&block) && block.blip.is_some();
        if !picture_fallback
            && equation_only(&block)
            && block.autofit
            && block.choice_extent.is_some()
            && let Some(framed) = &framed
        {
            // A self-sizing text box that shows only this equation takes
            // its new size, as on add.
            let points = run.font_size_pt().unwrap_or(DEFAULT_FONT_SIZE_PT);
            let size = text_frame(&prepared.omml, points, block.insets)?;
            resize_frame(
                &mut splice,
                &xml,
                &block,
                framed,
                &self.grid(&worksheet)?,
                size,
            )?;
        }
        // The fallback picture and the span of its relationship ID.
        let mut picture_edit = None;
        if picture_fallback {
            let picture = render(
                &prepared.omml,
                run.font_size_pt().unwrap_or(DEFAULT_FONT_SIZE_PT),
            )?;
            let blip = block
                .blip
                .clone()
                .ok_or_else(|| invalid("the equation fallback has no picture"))?;
            // The relationship prefix is the part's own: match the local
            // name of the prefixed `embed` attribute.
            let embed = attribute_span_by(&xml[blip.clone()], |key| key.ends_with(b":embed"))
                .ok_or_else(|| invalid("the equation picture is not embedded"))?;
            let embed = blip.start + embed.start..blip.start + embed.end;
            // The frame the picture covers: resized to the new picture when
            // the text box fits its text and its anchors have an extent of
            // their own, otherwise the text box's own.
            let frame = match (&block.choice_extent, &block.fallback_extent, &framed) {
                (Some(_), Some(_), Some(framed)) if block.autofit => {
                    let size = fitted(&picture, block.insets)?;
                    resize_frame(
                        &mut splice,
                        &xml,
                        &block,
                        framed,
                        &self.grid(&worksheet)?,
                        size,
                    )?;
                    Some(size)
                },
                (Some(choice), _, _) => extent(&xml[choice.clone()]),
                _ => None,
            };
            let frame = match frame {
                Some(frame) => frame,
                None => fitted(&picture, block.insets)?,
            };
            if let Some(stretch) = block.stretch.clone() {
                splice.replace(stretch, FALLBACK_FILL);
            }
            let place = (
                horizontal(&prepared.omml, site.align.as_deref())?,
                vertical(block.text_anchor.as_deref()),
            );
            let png = fallback_picture(&picture, frame, block.insets, place)?;
            picture_edit = Some((png, embed));
        }
        splice.replace(site.content.clone(), prepared.omml);
        let next = match picture_edit {
            Some((png, embed)) => {
                // Check the new part before the picture is staged.
                let mut probe = splice.clone();
                probe.replace(embed.clone(), PENDING_RELATIONSHIP);
                let probe = probe.apply(&xml, limits.max_drawing_bytes)?;
                self.check_equation_xml(&drawing, &probe, sites.len(), &limits)?;
                let relationship = self.stage_picture(&drawing, png, spaces.image)?;
                splice.replace(embed, relationship);
                splice.apply(&xml, limits.max_drawing_bytes)?
            },
            None => {
                let next = splice.apply(&xml, limits.max_drawing_bytes)?;
                let (range, fallback) = text_fallback(&next, index)?;
                let mut splice = Splice::default();
                splice.replace(range, fallback);
                splice.apply(&next, limits.max_drawing_bytes)?
            },
        };
        self.publish_equation_edit(&drawing, &xml, next, sites.len(), &limits)?;
        Ok(Some(EquationWrite {
            shape_id,
            index: equation,
            losses: prepared.losses,
        }))
    }

    /// Remove the worksheet's equation at `equation` (its position among
    /// [`crate::Package::equations`]). A text box whose text is only this
    /// equation leaves with it: the whole block, or for a block inside an
    /// anchor the whole anchor. In a text box whose text holds more, only the
    /// equation leaves: the Fallback becomes a copy of the Choice branch
    /// with every remaining equation as a run of its linear form, and a
    /// block left without equations is unwrapped to its Choice branch.
    /// Pictures and relationships nothing else uses are removed; the drawing
    /// part stays. Returns the ID of the shape that held the equation.
    ///
    /// # Errors
    ///
    /// Returns a not-found refusal for a missing equation, an
    /// unsupported-content refusal for an equation outside an editable
    /// block, and an error for malformed drawing XML.
    pub fn remove_equation(&mut self, sheet: usize, equation: usize) -> Result<u32> {
        const OPERATION: &str = "remove_equation";
        let drawing = self.equation_drawing(sheet, OPERATION)?;
        let limits = Limits::from_source_limits(self.staged.opc().read_limits());
        let xml = self.staged.blob(&drawing)?.to_vec();
        let (sites, blocks, anchors) = math::scan_all(&xml)?;
        let (site, index) = located(&sites, &blocks, equation, OPERATION)?;
        let block = &blocks[index];
        let shape_id = litchi_drawingml::math::cnvpr_id(
            &xml[block
                .choice_shape
                .clone()
                .ok_or_else(|| invalid("the equation block has no text box"))?],
        )?
        .ok_or_else(|| invalid("the equation shape has no ID"))?;
        let mut splice = Splice::default();
        if equation_only(block) {
            let range = match block.level {
                Level::Anchor => block
                    .outer_anchor
                    .and_then(|anchor| anchors.get(anchor))
                    .map(|anchor| anchor.range.clone())
                    .ok_or_else(|| invalid("the equation block has no anchor"))?,
                _ => block.range.clone(),
            };
            splice.replace(range, String::new());
            let next = splice.apply(&xml, limits.max_drawing_bytes)?;
            self.publish_equation_edit(&drawing, &xml, next, sites.len() - 1, &limits)?;
            return Ok(shape_id);
        }
        splice.replace(site.element.clone(), String::new());
        let staged = splice.apply(&xml, limits.max_drawing_bytes)?;
        let (_, staged_blocks) = math::scan(&staged)?;
        let staged_block = staged_blocks
            .get(index)
            .ok_or_else(|| invalid("the equation block disappeared"))?;
        let mut splice = Splice::default();
        if staged_block.equations == 0 {
            // No equation is left: the Choice branch stands without the
            // block, with the block's declarations it relies on.
            let content = staged_block
                .choice_content
                .clone()
                .ok_or_else(|| invalid("the equation block has no Choice branch"))?;
            let object = carry_declarations(
                &staged[staged_block.range.start..content.start],
                trimmed(&staged[content]),
            )?;
            splice.replace(staged_block.range.clone(), object);
        } else {
            let (range, fallback) = text_fallback(&staged, index)?;
            splice.replace(range, fallback);
        }
        let next = splice.apply(&staged, limits.max_drawing_bytes)?;
        self.publish_equation_edit(&drawing, &xml, next, sites.len() - 1, &limits)?;
        Ok(shape_id)
    }

    /// The drawing part of the worksheet at `sheet`, which an equation edit
    /// needs; a worksheet without one has no equations.
    fn equation_drawing(&self, sheet: usize, operation: &'static str) -> Result<PackURI> {
        let worksheet = self.staged.worksheet(sheet)?;
        objects::drawing_of(self.staged.opc().get_part(&worksheet)?.rels())?.ok_or_else(|| {
            Error::refused(
                RefusalKind::NotFound,
                operation,
                "the worksheet has no such equation",
            )
        })
    }
}

#[cfg(test)]
mod fixture_tests;

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn attributes_and_extents_are_found_in_start_tags() {
        let tag = br#"<a:ext cx="10" cy='20'/>"#;
        assert_eq!(extent(tag), Some((10, 20)));
        assert_eq!(extent(br#"<a:ext cx="0" cy="20"/>"#), None);
        assert_eq!(extent(br#"<a:ext cy="20"/>"#), None);
        let blip = br#"<a:blip xmlns:r="u" r:embed="rId7"/>"#;
        let span = attribute_span(blip, b"r:embed").unwrap();
        assert_eq!(&blip[span], b"rId7");
        assert!(attribute_span(b"<a:blip>", b"r:embed").is_none());
    }

    #[test]
    fn anchors_span_the_sheets_column_widths_and_row_heights() {
        // Defaults: 64-pixel columns, 15-point rows.
        let blank = Grid::of(br#"<worksheet><sheetData/></worksheet>"#, None).unwrap();
        let column = 64 * EMU_PER_PIXEL;
        let row = 190_500;
        let (origin, to) = blank.span(Marker::cell(1, 2), (column + 10, row / 2));
        assert_eq!(origin, (column, 2 * row));
        assert_eq!(
            (to.col, to.col_offset, to.row, to.row_offset),
            (2, 10, 2, row / 2)
        );
        // A wide column C (width 30: 210 pixels), a hidden column D, a
        // 30-point row 4 and the sheet's own defaults.
        let sheet = br#"<worksheet><sheetFormatPr defaultRowHeight="20"/><cols><col min="3" max="3" width="30" customWidth="1"/><col min="4" max="4" width="9" hidden="1"/></cols><sheetData><row r="4" ht="30" customHeight="1"/></sheetData></worksheet>"#;
        let grid = Grid::of(sheet, None).unwrap();
        assert_eq!(grid.column(2), 210 * EMU_PER_PIXEL);
        assert_eq!(grid.column(3), 0);
        assert_eq!(grid.row(3), 381_000);
        assert_eq!(grid.row(0), 254_000);
        let (_, to) = grid.span(Marker::cell(2, 3), (211 * EMU_PER_PIXEL, 381_000 + 1));
        // Past column C, through the hidden D, one pixel into E.
        assert_eq!((to.col, to.col_offset), (4, EMU_PER_PIXEL));
        assert_eq!((to.row, to.row_offset), (4, 1));
        // ECMA-376's width formula: 10.5 characters are 73 pixels at a
        // 7-pixel digit.
        let half = Grid::of(
            br#"<worksheet><cols><col min="1" max="1" width="10.5"/></cols></worksheet>"#,
            None,
        )
        .unwrap();
        assert_eq!(half.column(0), 73 * EMU_PER_PIXEL);
        // A 12-point default font has an 8-pixel digit: 8 characters plus
        // padding make a 72-pixel default column.
        let styles = br#"<styleSheet><fonts count="1"><font><sz val="12"/><name val="Calibri"/></font></fonts></styleSheet>"#;
        assert_eq!(Grid::digit_width(Some(styles)), 8.0);
        // Excel's default font, and a font outside the table (an estimate
        // from its size).
        let calibri = br#"<styleSheet><fonts count="2"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/></font><font><sz val="20"/><name val="Arial"/></font></fonts></styleSheet>"#;
        assert_eq!(Grid::digit_width(Some(calibri)), 7.0);
        let aptos = br#"<styleSheet><fonts><font><name val="aptos narrow"/><sz val="11"/></font></fonts></styleSheet>"#;
        assert_eq!(Grid::digit_width(Some(aptos)), 7.0);
        let other = br#"<styleSheet><fonts><font><sz val="14"/><name val="Georgia"/></font></fonts></styleSheet>"#;
        assert_eq!(Grid::digit_width(Some(other)), 9.0);
        let no_size =
            br#"<styleSheet><fonts><font><name val="Calibri"/></font></fonts></styleSheet>"#;
        assert_eq!(Grid::digit_width(Some(no_size)), 7.0);
        let twelve = Grid::of(br#"<worksheet/>"#, Some(styles)).unwrap();
        assert_eq!(twelve.column(5), 72 * EMU_PER_PIXEL);
        assert_eq!(Grid::digit_width(None), 7.0);
        // A frame wider than the sheet ends in its last column.
        let (_, to) = blank.span(Marker::cell(16_383, 0), (i64::MAX / 4, 1));
        assert_eq!(to.col, 16_383);
    }

    #[test]
    fn relationship_references_are_read_by_namespace() {
        let xml = format!(
            r#"<x xmlns:r="{RELATIONSHIPS}" xmlns:q="{STRICT_RELATIONSHIPS}"><a r:embed="rId1"/><b q:id="rId2" id="rId3"/></x>"#
        );
        assert_eq!(
            relationship_ids(xml.as_bytes()).unwrap(),
            BTreeSet::from(["rId1".to_owned(), "rId2".to_owned()])
        );
    }
}
