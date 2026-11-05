//! Worksheet features of an ODS sheet (record 0327): drawing frames and
//! custom shapes, `calcext` conditional formats, content validations, the
//! auto-filter database range and manual page breaks.
//!
//! [`read`] projects them from `content.xml` without decoding cell values.
//! The `Edit` verbs in `document.rs` change them as byte-exact provenance
//! splices of `content.xml` (and of `styles.xml` for the named cell style a
//! conditional format applies): every element an edit does not touch keeps
//! its bytes. A splice that needs one cell, row or column on its own (a
//! validation binding, a page break) splits the repeated run that holds it
//! and copies the rest of the run verbatim.
//!
//! Conditional-format formulas use the OpenFormula reference syntax
//! (`[.A1]`); [`to_odff`] and [`from_odff`] convert the A1 syntax of the
//! unified vocabulary.

use std::collections::BTreeMap;
use std::ops::Range;

use litchi_core::{Error, Result, xml::escape_xml};

use litchi_odf_common::core::{
    AuthoredXmlFragment, XmlSourcePart, XmlSplicePublication, rebuild_package_with_xml_splices,
};
use quick_xml::events::Event;

use super::{
    CALCEXT, DRAW, FO, OFFICE, STYLE, Span, TABLE, TEXT, XLINK, attribute, children, invalid,
    invalid_error, is_element, is_self_closing_span, scan, select_sheet,
};
use crate::model::conditional_format::{self as cf, Format};
use crate::package::Archive;

const SVG: &str = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
const LITCHI_CF_STYLE: &str = "Litchi_CF_";
const LITCHI_BREAK_STYLE: &str = "Litchi_Break_";
const LITCHI_SHAPE_STYLE: &str = "Litchi_Shape_";
/// The widest grid an ODS sheet addresses (`XFD1048576`).
const MAX_ROWS: u32 = 1_048_576;
const MAX_COLUMNS: u32 = 16_384;
/// The most cells one validation binding or materialization touches.
const MAX_MATERIALIZED_CELLS: u64 = 1_048_576;

// ---------------------------------------------------------------------------
// Public model

/// A zero-based inclusive `(first_row, first_column, last_row, last_column)`
/// cell rectangle.
pub type Rect = (u32, u32, u32, u32);

/// The features of one sheet.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct Features {
    /// Pictures and shapes, in document order.
    pub drawings: Vec<Drawing>,
    /// `calcext:conditional-format` elements, in document order.
    pub formats: Vec<Format>,
    /// Content validations bound to cells of the sheet, in catalog order.
    pub validations: Vec<Validation>,
    /// The sheet's auto-filter database range.
    pub auto_filter: Option<AutoFilter>,
    /// Rows (zero-based) whose style breaks the page before them.
    pub row_breaks: Vec<u32>,
    /// Columns (zero-based) whose style breaks the page before them.
    pub column_breaks: Vec<u32>,
}

/// Whether a drawing is a picture or another shape.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DrawingKind {
    /// A `draw:frame` holding a `draw:image`.
    Picture,
    /// A `draw:custom-shape` (or a legacy `draw:rect`/`draw:ellipse`).
    Shape,
}

/// A picture or shape of a sheet.
#[derive(Debug, Clone, PartialEq)]
pub struct Drawing {
    /// Picture or shape.
    pub kind: DrawingKind,
    /// `draw:name`.
    pub name: Option<String>,
    /// `svg:title`.
    pub title: Option<String>,
    /// `svg:desc` (alternative text).
    pub description: Option<String>,
    /// The cell a cell-anchored drawing sits in (zero-based row, column).
    pub cell: Option<(u32, u32)>,
    /// `table:end-cell-address` as a zero-based cell.
    pub end_cell: Option<(u32, u32)>,
    /// `svg:x`, `svg:y`, `svg:width`, `svg:height` in centimetres.
    pub frame: [Option<f64>; 4],
    /// `xlink:href` of the picture's `draw:image`.
    pub image: Option<String>,
    /// `draw:enhanced-geometry`'s `draw:type` (or the legacy element name).
    pub geometry: Option<String>,
    /// The shape's paragraph text, paragraphs joined by newlines.
    pub text: Option<String>,
    /// The fill of the shape's graphic style: `None` for no fill, the color
    /// for a solid one.
    pub fill: Option<Option<String>>,
    /// The stroke of the shape's graphic style: `None` for none, the color
    /// and width otherwise.
    pub line: Option<Option<(String, Option<String>)>>,
}

/// A content validation as it applies to one sheet.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Validation {
    /// `table:name`.
    pub name: String,
    /// `table:condition`.
    pub condition: String,
    /// The cell rectangles of the sheet bound to it, zero-based inclusive
    /// `(first_row, first_column, last_row, last_column)`.
    pub ranges: Vec<(u32, u32, u32, u32)>,
    /// `table:allow-empty-cell`.
    pub allow_empty: bool,
    /// `table:display-list` is not `none`.
    pub display_list: bool,
    /// The help message: title, text, displayed.
    pub help: Option<(Option<String>, Option<String>, bool)>,
    /// The error message: type, title, text, displayed.
    pub error: Option<(String, Option<String>, Option<String>, bool)>,
}

/// The auto-filter of a sheet: its database range.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AutoFilter {
    /// `table:name` of the database range.
    pub name: String,
    /// The zero-based inclusive rectangle it covers.
    pub range: (u32, u32, u32, u32),
}

/// A validation to write.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ValidationSpec {
    /// The rectangles it binds, zero-based inclusive.
    pub ranges: Vec<(u32, u32, u32, u32)>,
    /// `table:condition` (with its `of:` namespace prefix).
    pub condition: String,
    /// `table:allow-empty-cell`.
    pub allow_empty: bool,
    /// Whether a list shows its drop-down.
    pub display_list: bool,
    /// Help title, text and whether it shows.
    pub help: (Option<String>, Option<String>, bool),
    /// Error type (`stop`, `warning`, `information`), title, text and whether
    /// it shows.
    pub error: (String, Option<String>, Option<String>, bool),
}

/// The formatting a conditional format applies, written as a named cell
/// style in `styles.xml`.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct FormatStyle {
    /// `#RRGGBB` background.
    pub background: Option<String>,
    /// `#RRGGBB` text color.
    pub color: Option<String>,
    /// Bold.
    pub bold: Option<bool>,
    /// Italic.
    pub italic: Option<bool>,
    /// Underlined.
    pub underline: Option<bool>,
    /// Struck through.
    pub strike: Option<bool>,
    /// Font size in points.
    pub size_pt: Option<f64>,
    /// Font family.
    pub font: Option<String>,
}

impl FormatStyle {
    /// Whether the style sets nothing.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        *self == Self::default()
    }
}

/// A shape or picture frame to write.
#[derive(Debug, Clone, PartialEq)]
pub struct DrawingSpec {
    /// `draw:name`.
    pub name: String,
    /// `svg:desc`.
    pub description: Option<String>,
    /// `svg:title`.
    pub title: Option<String>,
    /// `svg:x`, `svg:y`, `svg:width`, `svg:height` in centimetres.
    pub frame: [f64; 4],
    /// The cell the drawing ends in (`table:end-cell-address`).
    pub end_cell: Option<(u32, u32)>,
    /// A shape's preset (`draw:type`); `None` for a picture.
    pub geometry: Option<String>,
    /// A shape's text.
    pub text: Option<String>,
    /// A shape's fill: `None` keeps the default, `Some(None)` no fill.
    pub fill: Option<Option<String>>,
    /// A shape's stroke color and width (`0.035cm`): `Some(None)` none.
    pub line: Option<Option<(String, String)>>,
    /// A picture's package path (`Pictures/…`).
    pub image: Option<String>,
}

/// Changes to a drawing; `None` keeps a value.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct DrawingChange {
    /// A new frame (centimetres).
    pub frame: Option<[f64; 4]>,
    /// A new end cell.
    pub end_cell: Option<Option<(u32, u32)>>,
    /// A new name.
    pub name: Option<String>,
    /// A new alternative text (`Some(None)` removes it).
    pub description: Option<Option<String>>,
}

// ---------------------------------------------------------------------------
// Addresses

/// The zero-based column of letters (`A` = 0).
fn column_of(letters: &str) -> Option<u32> {
    let mut column = 0u32;
    for byte in letters.bytes() {
        if !byte.is_ascii_alphabetic() {
            return None;
        }
        column = column
            .checked_mul(26)?
            .checked_add(u32::from(byte.to_ascii_uppercase() - b'A') + 1)?;
    }
    column.checked_sub(1).filter(|column| *column < MAX_COLUMNS)
}

/// The letters of a zero-based column.
#[must_use]
pub fn column_letters(column: u32) -> String {
    let mut letters = Vec::new();
    let mut value = column + 1;
    while value > 0 {
        let rem = (value - 1) % 26;
        letters.push(b'A' + rem as u8);
        value = (value - 1) / 26;
    }
    letters.reverse();
    String::from_utf8(letters).unwrap_or_default()
}

/// A cell `A1` (dollar signs ignored) as zero-based `(row, column)`.
fn cell_of(text: &str) -> Option<(u32, u32)> {
    let text = text.replace('$', "");
    let split = text.find(|char: char| char.is_ascii_digit())?;
    let (letters, digits) = text.split_at(split);
    let column = column_of(letters)?;
    let row: u32 = digits.parse().ok()?;
    let row = row.checked_sub(1).filter(|row| *row < MAX_ROWS)?;
    Some((row, column))
}

/// The cell or rectangle of an ODF address (`Sheet1.A1`, `$Sheet1.$A$1:.$B$2`,
/// `'My Sheet'.A1:'My Sheet'.C9`): its sheet and zero-based rectangle.
#[must_use]
pub fn parse_address(address: &str) -> Option<(Option<String>, (u32, u32, u32, u32))> {
    let parts = split_unquoted(address.trim(), ':');
    let first = parts.first()?;
    let (sheet, start) = address_cell(first)?;
    let end = match parts.get(1) {
        Some(part) => address_cell(part)?.1,
        None => start,
    };
    Some((
        sheet,
        (
            start.0.min(end.0),
            start.1.min(end.1),
            start.0.max(end.0),
            start.1.max(end.1),
        ),
    ))
}

fn split_unquoted(text: &str, separator: char) -> Vec<String> {
    let mut out = Vec::new();
    let mut current = String::new();
    let mut quoted = false;
    for char in text.chars() {
        if char == '\'' {
            quoted = !quoted;
        }
        if char == separator && !quoted {
            out.push(std::mem::take(&mut current));
        } else {
            current.push(char);
        }
    }
    out.push(current);
    out
}

fn address_cell(part: &str) -> Option<(Option<String>, (u32, u32))> {
    let part = part.trim();
    let (sheet, cell) = match part.rfind('.') {
        Some(dot) => (Some(&part[..dot]), &part[dot + 1..]),
        None => (None, part),
    };
    let sheet = sheet
        .map(|sheet| sheet.trim_start_matches('$'))
        .filter(|sheet| !sheet.is_empty())
        .map(|sheet| {
            sheet
                .strip_prefix('\'')
                .and_then(|sheet| sheet.strip_suffix('\''))
                .map_or_else(|| sheet.to_owned(), |sheet| sheet.replace("''", "'"))
        });
    Some((sheet, cell_of(cell)?))
}

/// The ODF name of a sheet in an address (quoted when needed).
fn quoted_sheet(sheet: &str) -> String {
    if !sheet.is_empty()
        && sheet
            .chars()
            .all(|char| char.is_alphanumeric() || char == '_')
    {
        sheet.to_owned()
    } else {
        format!("'{}'", sheet.replace('\'', "''"))
    }
}

/// The ODF address of a rectangle on a sheet (`Sheet1.A1:Sheet1.C9`).
#[must_use]
pub fn address(
    sheet: &str,
    (first_row, first_column, last_row, last_column): (u32, u32, u32, u32),
) -> String {
    let sheet = quoted_sheet(sheet);
    let start = format!("{sheet}.{}{}", column_letters(first_column), first_row + 1);
    if (first_row, first_column) == (last_row, last_column) {
        return start;
    }
    format!(
        "{start}:{sheet}.{}{}",
        column_letters(last_column),
        last_row + 1
    )
}

/// Convert the A1 references of a formula (`A1`, `$B$2:C3`, `Sheet2!A1`)
/// to OpenFormula references (`[.A1]`, `[.$B$2:.C3]`, `[Sheet2.A1]`);
/// strings are kept.
#[must_use]
pub fn to_odff(formula: &str) -> String {
    let chars: Vec<char> = formula.chars().collect();
    let mut out = String::with_capacity(formula.len() + 8);
    let mut index = 0;
    while index < chars.len() {
        let char = chars[index];
        if char == '"' {
            out.push(char);
            index += 1;
            while index < chars.len() {
                out.push(chars[index]);
                if chars[index] == '"' {
                    if chars.get(index + 1) == Some(&'"') {
                        out.push('"');
                        index += 2;
                        continue;
                    }
                    index += 1;
                    break;
                }
                index += 1;
            }
            continue;
        }
        if char.is_ascii_alphabetic() || char == '$' || char == '\'' {
            let previous = index.checked_sub(1).map(|at| chars[at]);
            let boundary = previous.is_none_or(|previous| {
                !(previous.is_alphanumeric()
                    || previous == '_'
                    || previous == '.'
                    || previous == '[')
            });
            if boundary {
                if let Some((length, text)) = reference_at(&chars[index..]) {
                    out.push_str(&text);
                    index += length;
                    continue;
                }
            }
            // Copy the whole identifier so its tail is not read as a
            // reference.
            while index < chars.len()
                && (chars[index].is_alphanumeric() || chars[index] == '_' || chars[index] == '.')
            {
                out.push(chars[index]);
                index += 1;
            }
            continue;
        }
        out.push(char);
        index += 1;
    }
    out
}

/// A reference at the start of `chars`: its length and OpenFormula text.
fn reference_at(chars: &[char]) -> Option<(usize, String)> {
    let mut at = 0;
    let mut sheet = None;
    // An optional sheet: `Name!` or `'Quoted Name'!`.
    if chars.first() == Some(&'\'') {
        let mut end = 1;
        while end < chars.len() && chars[end] != '\'' {
            end += 1;
        }
        if chars.get(end + 1) != Some(&'!') {
            return None;
        }
        sheet = Some(chars[1..end].iter().collect::<String>());
        at = end + 2;
    } else if let Some(bang) = chars.iter().take(64).position(|char| *char == '!') {
        let name: String = chars[..bang].iter().collect();
        if !name.is_empty()
            && name
                .chars()
                .all(|char| char.is_alphanumeric() || char == '_')
        {
            sheet = Some(name);
            at = bang + 1;
        }
    }
    let cell = |from: usize| -> Option<(usize, String)> {
        let mut end = from;
        if chars.get(end) == Some(&'$') {
            end += 1;
        }
        let letters_start = end;
        while end < chars.len() && chars[end].is_ascii_alphabetic() {
            end += 1;
        }
        if end == letters_start || end - letters_start > 3 {
            return None;
        }
        if chars.get(end) == Some(&'$') {
            end += 1;
        }
        let digits_start = end;
        while end < chars.len() && chars[end].is_ascii_digit() {
            end += 1;
        }
        if end == digits_start {
            return None;
        }
        Some((end, chars[from..end].iter().collect()))
    };
    let (end, first) = cell(at)?;
    // Not a function name (`LOG10(`) or a longer identifier.
    if chars
        .get(end)
        .is_some_and(|char| char.is_alphanumeric() || *char == '(' || *char == '_')
    {
        return None;
    }
    let (end, last) = if chars.get(end) == Some(&':') {
        match cell(end + 1) {
            Some((stop, last))
                if !chars
                    .get(stop)
                    .is_some_and(|char| char.is_alphanumeric() || *char == '(') =>
            {
                (stop, Some(last))
            },
            _ => (end, None),
        }
    } else {
        (end, None)
    };
    let prefix = sheet.map_or_else(
        || ".".to_owned(),
        |sheet| format!("{}.", quoted_sheet(&sheet)),
    );
    let text = match last {
        Some(last) => format!("[{prefix}{first}:.{last}]"),
        None => format!("[{prefix}{first}]"),
    };
    Some((end, text))
}

/// Convert OpenFormula references back to the A1 syntax.
#[must_use]
pub fn from_odff(formula: &str) -> String {
    let formula = formula.strip_prefix("of:").unwrap_or(formula);
    let formula = formula.strip_prefix('=').unwrap_or(formula);
    let mut out = String::with_capacity(formula.len());
    let mut rest = formula;
    let mut quoted = false;
    while let Some(char) = rest.chars().next() {
        if char == '"' {
            quoted = !quoted;
        }
        if char == '[' && !quoted {
            if let Some(close) = rest.find(']') {
                let inner = &rest[1..close];
                out.push_str(&odff_reference(inner));
                rest = &rest[close + 1..];
                continue;
            }
        }
        out.push(char);
        rest = &rest[char.len_utf8()..];
    }
    out
}

fn odff_reference(inner: &str) -> String {
    let parts = split_unquoted(inner, ':');
    let mut out = Vec::new();
    for (index, part) in parts.iter().enumerate() {
        let (sheet, cell) = match part.rfind('.') {
            Some(dot) => (&part[..dot], &part[dot + 1..]),
            None => ("", part.as_str()),
        };
        let sheet = sheet.trim_start_matches('$');
        if index == 0 && !sheet.is_empty() {
            out.push(format!("{sheet}!{cell}"));
        } else {
            out.push(cell.to_owned());
        }
    }
    out.join(":")
}

// ---------------------------------------------------------------------------
// The sheet grid

/// One row run: its span, first logical row and repeat count.
#[derive(Debug, Clone, Copy)]
struct RowRun {
    span: usize,
    first: u32,
    repeat: u32,
}

/// One cell run of a row.
#[derive(Debug, Clone, Copy)]
struct CellRun {
    span: usize,
    first: u32,
    repeat: u32,
}

fn repeat_of(xml: &str, span: &Span, name: &[u8]) -> Result<u32> {
    Ok(match attribute(xml, span, name)? {
        Some(value) => value
            .trim()
            .parse::<u32>()
            .ok()
            .filter(|value| *value > 0)
            .ok_or_else(|| invalid_error("ODS repeat count is invalid"))?,
        None => 1,
    })
}

/// The row runs of a table, through row groups and header rows.
fn row_runs(xml: &str, spans: &[Span], table: usize) -> Result<Vec<RowRun>> {
    let mut out = Vec::new();
    let mut next = 0u32;
    collect_rows(xml, spans, table, &mut next, &mut out)?;
    Ok(out)
}

fn collect_rows(
    xml: &str,
    spans: &[Span],
    parent: usize,
    next: &mut u32,
    out: &mut Vec<RowRun>,
) -> Result<()> {
    for (index, span) in spans.iter().enumerate().skip(parent + 1) {
        if span.start >= spans[parent].end {
            break;
        }
        if span.parent != Some(parent) || span.namespace.as_deref() != Some(TABLE) {
            continue;
        }
        match span.local.as_str() {
            "table-row" => {
                let repeat = repeat_of(xml, span, b"table:number-rows-repeated")?;
                out.push(RowRun {
                    span: index,
                    first: *next,
                    repeat,
                });
                *next = next.saturating_add(repeat);
            },
            "table-row-group" | "table-header-rows" | "table-rows" => {
                collect_rows(xml, spans, index, next, out)?;
            },
            _ => {},
        }
    }
    Ok(())
}

/// The cell runs of a row (covered cells included).
fn cell_runs(xml: &str, spans: &[Span], row: usize) -> Result<Vec<CellRun>> {
    let mut out = Vec::new();
    let mut next = 0u32;
    for (index, span) in spans.iter().enumerate().skip(row + 1) {
        if span.start >= spans[row].end {
            break;
        }
        if span.parent != Some(row)
            || span.namespace.as_deref() != Some(TABLE)
            || !matches!(span.local.as_str(), "table-cell" | "covered-table-cell")
        {
            continue;
        }
        let repeat = repeat_of(xml, span, b"table:number-columns-repeated")?;
        out.push(CellRun {
            span: index,
            first: next,
            repeat,
        });
        next = next.saturating_add(repeat);
    }
    Ok(out)
}

/// The column runs of a table, through column groups and header columns.
fn column_runs(xml: &str, spans: &[Span], table: usize) -> Result<Vec<RowRun>> {
    let mut out = Vec::new();
    let mut next = 0u32;
    collect_columns(xml, spans, table, &mut next, &mut out)?;
    Ok(out)
}

fn collect_columns(
    xml: &str,
    spans: &[Span],
    parent: usize,
    next: &mut u32,
    out: &mut Vec<RowRun>,
) -> Result<()> {
    for (index, span) in spans.iter().enumerate().skip(parent + 1) {
        if span.start >= spans[parent].end {
            break;
        }
        if span.parent != Some(parent) || span.namespace.as_deref() != Some(TABLE) {
            continue;
        }
        match span.local.as_str() {
            "table-column" => {
                let repeat = repeat_of(xml, span, b"table:number-columns-repeated")?;
                out.push(RowRun {
                    span: index,
                    first: *next,
                    repeat,
                });
                *next = next.saturating_add(repeat);
            },
            "table-column-group" | "table-header-columns" | "table-columns" => {
                collect_columns(xml, spans, index, next, out)?;
            },
            _ => {},
        }
    }
    Ok(())
}

/// The logical cell of every table-cell span of a table.
fn cell_positions(xml: &str, spans: &[Span], table: usize) -> Result<BTreeMap<usize, (u32, u32)>> {
    let mut out = BTreeMap::new();
    for run in row_runs(xml, spans, table)? {
        for cell in cell_runs(xml, spans, run.span)? {
            out.insert(cell.span, (run.first, cell.first));
        }
    }
    Ok(out)
}

// ---------------------------------------------------------------------------
// Reading

fn text_of(xml: &str, span: &Span) -> String {
    let inner = xml.get(span.tag_end..span.close_start).unwrap_or("");
    let mut out = String::new();
    let mut depth = 0usize;
    let mut chars = inner.char_indices().peekable();
    let mut tag = String::new();
    while let Some((_, char)) = chars.next() {
        if char == '<' {
            tag.clear();
            for (_, next) in chars.by_ref() {
                if next == '>' {
                    break;
                }
                tag.push(next);
            }
            if tag.starts_with("text:s")
                && (tag.len() == 6 || tag.as_bytes()[6] == b' ' || tag.as_bytes()[6] == b'/')
            {
                out.push(' ');
            } else if tag.starts_with("text:tab") {
                out.push('\t');
            } else if tag.starts_with("text:line-break") {
                out.push('\n');
            }
            if tag.starts_with('/') {
                depth = depth.saturating_sub(1);
            } else if !tag.ends_with('/') {
                depth += 1;
            }
            continue;
        }
        out.push(char);
    }
    unescape(&out)
}

fn unescape(text: &str) -> String {
    text.replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&quot;", "\"")
        .replace("&apos;", "'")
        .replace("&amp;", "&")
}

fn length_cm(value: &str) -> Option<f64> {
    let value = value.trim();
    let split = value.find(|char: char| char.is_ascii_alphabetic())?;
    let (number, unit) = value.split_at(split);
    let number: f64 = number.trim().parse().ok()?;
    Some(match unit {
        "cm" => number,
        "mm" => number / 10.0,
        "in" | "inch" => number * 2.54,
        "pt" => number * 2.54 / 72.0,
        "pc" => number * 2.54 / 6.0,
        "px" => number * 2.54 / 96.0,
        _ => return None,
    })
}

/// The automatic or common styles of `family` by name: their property
/// element attributes.
fn style_properties(
    xml: &str,
    spans: &[Span],
    family: &str,
    properties: &str,
) -> BTreeMap<String, String> {
    let mut out = BTreeMap::new();
    for (index, span) in spans.iter().enumerate() {
        if !is_element(span, STYLE, "style") {
            continue;
        }
        if attribute(xml, span, b"style:family")
            .ok()
            .flatten()
            .as_deref()
            != Some(family)
        {
            continue;
        }
        let Some(name) = attribute(xml, span, b"style:name").ok().flatten() else {
            continue;
        };
        let tags: String = children(spans, index, STYLE, properties)
            .into_iter()
            .filter_map(|child| xml.get(spans[child].start..spans[child].tag_end))
            .collect();
        out.insert(name, tags);
    }
    out
}

fn attr_in(tag: &str, name: &str) -> Option<String> {
    let needle = format!(" {name}=\"");
    let start = tag.find(&needle)? + needle.len();
    let end = tag[start..].find('"')? + start;
    Some(unescape(&tag[start..end]))
}

/// Every feature of the sheet `sheet` of a `content.xml`.
///
/// # Errors
///
/// Returns an error when the sheet does not exist or the XML is malformed.
pub fn read(content_xml: &str, sheet: &str) -> Result<Features> {
    let spans = scan(content_xml)?;
    read_spans(content_xml, &spans, sheet)
}

fn read_spans(xml: &str, spans: &[Span], sheet: &str) -> Result<Features> {
    let table = select_sheet(xml, spans, sheet)?;
    let mut out = Features::default();
    let positions = cell_positions(xml, spans, table)?;
    let graphics = style_properties(xml, spans, "graphic", "graphic-properties");
    // Drawings: children of the sheet's table:shapes and of its cells.
    for (index, span) in spans.iter().enumerate() {
        if span.start < spans[table].start || span.end > spans[table].end {
            continue;
        }
        let Some(parent) = span.parent else { continue };
        let parent_span = &spans[parent];
        let cell = if is_element(parent_span, TABLE, "shapes") && parent_span.parent == Some(table)
        {
            None
        } else if is_element(parent_span, TABLE, "table-cell") {
            match positions.get(&parent) {
                Some(position) => Some(*position),
                None => continue,
            }
        } else {
            continue;
        };
        if span.namespace.as_deref() != Some(DRAW) {
            continue;
        }
        let kind = match span.local.as_str() {
            "frame" => {
                if children(spans, index, DRAW, "image").is_empty() {
                    continue;
                }
                DrawingKind::Picture
            },
            "custom-shape" | "rect" | "ellipse" => DrawingKind::Shape,
            _ => continue,
        };
        out.drawings
            .push(drawing(xml, spans, index, kind, cell, &graphics)?);
    }
    // Conditional formats.
    for container in children(spans, table, CALCEXT, "conditional-formats") {
        for format in children(spans, container, CALCEXT, "conditional-format") {
            out.formats.push(parse_format(xml, spans, format)?);
        }
    }
    // Validations bound in this sheet.
    let mut bound: BTreeMap<String, Vec<(u32, u32, u32)>> = BTreeMap::new();
    for run in row_runs(xml, spans, table)? {
        for cell in cell_runs(xml, spans, run.span)? {
            if let Some(name) = attribute(xml, &spans[cell.span], b"table:content-validation-name")?
            {
                let last_column = cell.first + cell.repeat - 1;
                for row in run.first..run.first.saturating_add(run.repeat).min(MAX_ROWS) {
                    bound
                        .entry(name.clone())
                        .or_default()
                        .push((row, cell.first, last_column));
                }
            }
        }
    }
    if !bound.is_empty() {
        for container in descendants_of(spans, OFFICE, "spreadsheet", TABLE, "content-validations")
        {
            for definition in children(spans, container, TABLE, "content-validation") {
                let Some(name) = attribute(xml, &spans[definition], b"table:name")? else {
                    continue;
                };
                if let Some(cells) = bound.get(&name) {
                    out.validations.push(validation(
                        xml,
                        spans,
                        definition,
                        name.clone(),
                        rectangles(cells),
                    )?);
                }
            }
        }
    }
    // The auto-filter: a database range with filter buttons on this sheet.
    for container in descendants_of(spans, OFFICE, "spreadsheet", TABLE, "database-ranges") {
        for range in children(spans, container, TABLE, "database-range") {
            let span = &spans[range];
            if attribute(xml, span, b"table:display-filter-buttons")?.as_deref() != Some("true") {
                continue;
            }
            let Some(target) = attribute(xml, span, b"table:target-range-address")? else {
                continue;
            };
            let Some((Some(owner), rect)) = parse_address(&target) else {
                continue;
            };
            if owner == sheet {
                out.auto_filter = Some(AutoFilter {
                    name: attribute(xml, span, b"table:name")?.unwrap_or_default(),
                    range: rect,
                });
            }
        }
    }
    // Page breaks from row and column styles.
    let row_styles = style_properties(xml, spans, "table-row", "table-row-properties");
    let column_styles = style_properties(xml, spans, "table-column", "table-column-properties");
    let breaks = |styles: &BTreeMap<String, String>, name: Option<String>| {
        name.and_then(|name| styles.get(&name).cloned())
            .is_some_and(|tag| attr_in(&tag, "fo:break-before").as_deref() == Some("page"))
    };
    for run in row_runs(xml, spans, table)? {
        if breaks(
            &row_styles,
            attribute(xml, &spans[run.span], b"table:style-name")?,
        ) {
            out.row_breaks
                .extend(run.first..run.first.saturating_add(run.repeat.min(1_024)));
        }
    }
    for run in column_runs(xml, spans, table)? {
        if breaks(
            &column_styles,
            attribute(xml, &spans[run.span], b"table:style-name")?,
        ) {
            out.column_breaks
                .extend(run.first..run.first.saturating_add(run.repeat.min(1_024)));
        }
    }
    Ok(out)
}

fn descendants_of(
    spans: &[Span],
    namespace: &str,
    local: &str,
    child_namespace: &str,
    child_local: &str,
) -> Vec<usize> {
    spans
        .iter()
        .enumerate()
        .filter(|(_, span)| is_element(span, namespace, local))
        .flat_map(|(index, _)| children(spans, index, child_namespace, child_local))
        .collect()
}

/// Rectangles covering `cells` (row, first column, last column): equal
/// column spans on consecutive rows merge.
fn rectangles(cells: &[(u32, u32, u32)]) -> Vec<(u32, u32, u32, u32)> {
    let mut sorted = cells.to_vec();
    sorted.sort_unstable();
    // Merge horizontally adjacent runs of one row.
    let mut rows: Vec<(u32, u32, u32)> = Vec::new();
    for (row, first, last) in sorted {
        match rows.last_mut() {
            Some(held) if held.0 == row && held.2 + 1 == first => held.2 = last,
            _ => rows.push((row, first, last)),
        }
    }
    let mut out: Vec<(u32, u32, u32, u32)> = Vec::new();
    for (row, first, last) in rows {
        if let Some(open) = out
            .iter_mut()
            .find(|rect| rect.1 == first && rect.3 == last && rect.2 + 1 == row)
        {
            open.2 = row;
        } else {
            out.push((row, first, row, last));
        }
    }
    out.sort_unstable();
    out
}

fn child_text(
    xml: &str,
    spans: &[Span],
    parent: usize,
    namespace: &str,
    local: &str,
) -> Option<String> {
    children(spans, parent, namespace, local)
        .first()
        .map(|child| text_of(xml, &spans[*child]))
        .filter(|text| !text.is_empty())
}

fn drawing(
    xml: &str,
    spans: &[Span],
    index: usize,
    kind: DrawingKind,
    cell: Option<(u32, u32)>,
    graphics: &BTreeMap<String, String>,
) -> Result<Drawing> {
    let span = &spans[index];
    let length = |name: &[u8]| -> Result<Option<f64>> {
        Ok(attribute(xml, span, name)?.as_deref().and_then(length_cm))
    };
    let end_cell = attribute(xml, span, b"table:end-cell-address")?
        .and_then(|address| parse_address(&address))
        .map(|(_, rect)| (rect.0, rect.1));
    let image = children(spans, index, DRAW, "image")
        .first()
        .and_then(|image| attribute(xml, &spans[*image], b"xlink:href").ok().flatten());
    let geometry = match span.local.as_str() {
        "custom-shape" => children(spans, index, DRAW, "enhanced-geometry")
            .first()
            .and_then(|geometry| {
                attribute(xml, &spans[*geometry], b"draw:type")
                    .ok()
                    .flatten()
            }),
        "rect" => Some("rectangle".to_owned()),
        "ellipse" => Some("ellipse".to_owned()),
        _ => None,
    };
    let text = (kind == DrawingKind::Shape)
        .then(|| {
            let paragraphs: Vec<String> = children(spans, index, TEXT, "p")
                .into_iter()
                .map(|paragraph| text_of(xml, &spans[paragraph]))
                .collect();
            (!paragraphs.is_empty()).then(|| paragraphs.join("\n"))
        })
        .flatten();
    let style =
        attribute(xml, span, b"draw:style-name")?.and_then(|name| graphics.get(&name).cloned());
    let (fill, line) = match (&style, kind) {
        (Some(tag), DrawingKind::Shape) => {
            let fill = match attr_in(tag, "draw:fill").as_deref() {
                Some("none") => Some(None),
                Some("solid") => attr_in(tag, "draw:fill-color").map(Some),
                _ => None,
            };
            let line = match attr_in(tag, "draw:stroke").as_deref() {
                Some("none") => Some(None),
                Some(_) => attr_in(tag, "svg:stroke-color")
                    .map(|color| Some((color, attr_in(tag, "svg:stroke-width")))),
                None => None,
            };
            (fill, line)
        },
        _ => (None, None),
    };
    Ok(Drawing {
        kind,
        name: attribute(xml, span, b"draw:name")?,
        title: child_text(xml, spans, index, SVG, "title"),
        description: child_text(xml, spans, index, SVG, "desc"),
        cell,
        end_cell,
        frame: [
            length(b"svg:x")?,
            length(b"svg:y")?,
            length(b"svg:width")?,
            length(b"svg:height")?,
        ],
        image,
        geometry,
        text,
        fill,
        line,
    })
}

fn bool_attr(xml: &str, span: &Span, name: &[u8]) -> Result<Option<bool>> {
    Ok(attribute(xml, span, name)?.map(|value| value == "true"))
}

/// Parse one `calcext:conditional-format`.
fn parse_format(xml: &str, spans: &[Span], index: usize) -> Result<Format> {
    let span = &spans[index];
    let ranges = attribute(xml, span, b"calcext:target-range-address")?.unwrap_or_default();
    let mut format = Format {
        target_range_addresses: split_unquoted(&ranges, ' ')
            .into_iter()
            .filter(|range| !range.is_empty())
            .collect(),
        rules: Vec::new(),
    };
    for (child, rule) in spans.iter().enumerate().skip(index + 1) {
        if rule.start >= span.end {
            break;
        }
        if rule.parent != Some(index) || rule.namespace.as_deref() != Some(CALCEXT) {
            continue;
        }
        let entry_type = |entry: &Span| -> Result<cf::EntryType> {
            cf::EntryType::parse(&attribute(xml, entry, b"calcext:type")?.unwrap_or_default())
        };
        let value = |entry: &Span| -> Result<String> {
            Ok(attribute(xml, entry, b"calcext:value")?.unwrap_or_default())
        };
        let parsed = match rule.local.as_str() {
            "condition" => cf::Rule::Condition(cf::Condition {
                condition: value(rule)?,
                apply_style_name: attribute(xml, rule, b"calcext:apply-style-name")?
                    .unwrap_or_default(),
                base_cell_address: attribute(xml, rule, b"calcext:base-cell-address")?,
            }),
            "color-scale" => {
                let mut entries = Vec::new();
                for entry in children(spans, child, CALCEXT, "color-scale-entry") {
                    let entry = &spans[entry];
                    entries.push(cf::ColorScaleEntry {
                        entry_type: entry_type(entry)?,
                        value: value(entry)?,
                        color: attribute(xml, entry, b"calcext:color")?.unwrap_or_default(),
                    });
                }
                cf::Rule::ColorScale(cf::ColorScale { entries })
            },
            "data-bar" => {
                let mut entries = Vec::new();
                for entry in children(spans, child, CALCEXT, "formatting-entry") {
                    let entry = &spans[entry];
                    entries.push(cf::DataBarEntry {
                        entry_type: entry_type(entry)?,
                        value: value(entry)?,
                    });
                }
                cf::Rule::DataBar(cf::DataBar {
                    positive_color: attribute(xml, rule, b"calcext:positive-color")?,
                    negative_color: attribute(xml, rule, b"calcext:negative-color")?,
                    gradient: bool_attr(xml, rule, b"calcext:gradient")?,
                    axis_position: attribute(xml, rule, b"calcext:axis-position")?
                        .map(|position| cf::DataBarAxisPosition::parse(&position))
                        .transpose()?,
                    show_value: bool_attr(xml, rule, b"calcext:show-value")?,
                    axis_color: attribute(xml, rule, b"calcext:axis-color")?,
                    min_length: attribute(xml, rule, b"calcext:min-length")?,
                    max_length: attribute(xml, rule, b"calcext:max-length")?,
                    entries,
                })
            },
            "icon-set" => {
                let mut entries = Vec::new();
                for entry in children(spans, child, CALCEXT, "formatting-entry") {
                    let entry = &spans[entry];
                    entries.push(cf::IconSetEntry {
                        entry_type: entry_type(entry)?,
                        value: value(entry)?,
                        greater_equal: bool_attr(xml, entry, b"calcext:greater-equal")?,
                    });
                }
                cf::Rule::IconSet(cf::IconSet {
                    icon_set_type: cf::IconSetType::parse(
                        &attribute(xml, rule, b"calcext:icon-set-type")?.unwrap_or_default(),
                    )?,
                    show_value: bool_attr(xml, rule, b"calcext:show-value")?,
                    custom: bool_attr(xml, rule, b"calcext:custom")?,
                    custom_icons: Vec::new(),
                    entries,
                })
            },
            "date-is" => cf::Rule::DateIs(cf::DateIs {
                date: cf::DateType::parse(
                    &attribute(xml, rule, b"calcext:date")?.unwrap_or_default(),
                )?,
                style: attribute(xml, rule, b"calcext:style")?.unwrap_or_default(),
            }),
            _ => continue,
        };
        format.rules.push(parsed);
    }
    Ok(format)
}

fn validation(
    xml: &str,
    spans: &[Span],
    index: usize,
    name: String,
    ranges: Vec<(u32, u32, u32, u32)>,
) -> Result<Validation> {
    let span = &spans[index];
    let message = |local: &str| children(spans, index, TABLE, local).first().copied();
    let help = message("help-message")
        .map(|help| -> Result<_> {
            let help_span = &spans[help];
            Ok((
                attribute(xml, help_span, b"table:title")?,
                child_text(xml, spans, help, TEXT, "p"),
                attribute(xml, help_span, b"table:display")?.as_deref() != Some("false"),
            ))
        })
        .transpose()?;
    let error = message("error-message")
        .map(|error| -> Result<_> {
            let error_span = &spans[error];
            Ok((
                attribute(xml, error_span, b"table:message-type")?
                    .unwrap_or_else(|| "stop".to_owned()),
                attribute(xml, error_span, b"table:title")?,
                child_text(xml, spans, error, TEXT, "p"),
                attribute(xml, error_span, b"table:display")?.as_deref() != Some("false"),
            ))
        })
        .transpose()?;
    Ok(Validation {
        name,
        condition: attribute(xml, span, b"table:condition")?.unwrap_or_default(),
        ranges,
        allow_empty: attribute(xml, span, b"table:allow-empty-cell")?.as_deref() != Some("false"),
        display_list: attribute(xml, span, b"table:display-list")?.as_deref() != Some("none"),
        help,
        error,
    })
}

// ---------------------------------------------------------------------------
// Editing helpers

/// One staged replacement of `content.xml` (or `styles.xml`) bytes.
type Splice = (Range<usize>, Vec<u8>);

/// Whether a replacement is exactly one non-empty start tag (a lexical
/// attribute edit of an element whose body stays in place).
fn is_start_tag(bytes: &[u8]) -> bool {
    bytes.first() == Some(&b'<')
        && bytes.get(1) != Some(&b'/')
        && bytes.ends_with(b">")
        && !bytes.ends_with(b"/>")
        && bytes.iter().filter(|byte| **byte == b'<').count() == 1
}

/// Apply splices to one XML part of a package in one publication.
fn splice_part(
    source: &[u8],
    path: &str,
    mut edits: Vec<Splice>,
    max_output: usize,
) -> Result<Vec<u8>> {
    if edits.is_empty() {
        return Ok(source.to_vec());
    }
    edits.sort_by_key(|(range, _)| (range.start, range.end));
    for pair in edits.windows(2) {
        if pair[0].0.end > pair[1].0.start {
            return invalid("ODS feature splices overlap");
        }
    }
    let package = Archive::from_bytes(source.to_vec())?;
    let part = XmlSourcePart::load(package.package(), path)?;
    let mut publication = XmlSplicePublication::new(part.clone());
    for (range, replacement) in edits {
        let expected = part
            .bytes()
            .get(range.clone())
            .ok_or_else(|| invalid_error("ODS feature splice range is invalid"))?;
        if expected == replacement.as_slice() {
            continue;
        }
        let tag_edit = is_start_tag(&replacement) && is_start_tag(expected);
        let proof = part.checked_range(range, expected)?;
        let fragment = if replacement.is_empty() {
            AuthoredXmlFragment::deletion()
        } else if tag_edit {
            AuthoredXmlFragment::start_tag(replacement)?
        } else {
            AuthoredXmlFragment::markup(replacement)?
        };
        publication.replace(proof, fragment)?;
    }
    rebuild_package_with_xml_splices(package.package(), vec![publication], max_output)
}

fn content(source: &[u8]) -> Result<(String, Vec<Span>)> {
    let package = Archive::from_bytes(source.to_vec())?;
    let xml = package.content_xml().to_string();
    let spans = scan(&xml)?;
    Ok((xml, spans))
}

fn part_xml(source: &[u8], path: &str) -> Result<Option<String>> {
    let package = Archive::from_bytes(source.to_vec())?;
    Ok(XmlSourcePart::load(package.package(), path)
        .ok()
        .map(|part| String::from_utf8_lossy(part.bytes()).into_owned()))
}

/// Whether the content root binds `prefix`.
fn root_namespace_declared(xml: &str, prefix: &str) -> bool {
    let root_end = xml
        .find("<office:document-content")
        .and_then(|start| xml[start..].find('>').map(|end| start + end));
    root_end.is_some_and(|end| xml[..end].contains(&format!("xmlns:{prefix}=")))
}

/// Bind `prefix` on the first element of `markup` when the content root
/// does not bind it.
fn bind(markup: String, xml: &str, prefix: &str, namespace: &str) -> String {
    if root_namespace_declared(xml, prefix) {
        return markup;
    }
    let Some(space) = markup.find([' ', '>', '/']) else {
        return markup;
    };
    format!(
        "{} xmlns:{prefix}=\"{namespace}\"{}",
        &markup[..space],
        &markup[space..]
    )
}

/// A start tag's qualified name, attributes (decoded) and whether it is
/// self-closing.
fn tag_parts(tag: &str) -> Result<(String, Vec<(String, String)>, bool)> {
    let mut reader = quick_xml::Reader::from_str(tag);
    let event = reader
        .read_event()
        .map_err(|error| invalid_error(format!("invalid ODS start tag: {error}")))?;
    let (element, empty) = match event {
        Event::Start(element) => (element, false),
        Event::Empty(element) => (element, true),
        _ => return invalid("ODS element start tag is missing"),
    };
    let name = String::from_utf8_lossy(element.name().as_ref()).into_owned();
    let mut attributes = Vec::new();
    for candidate in element.attributes().with_checks(true) {
        let candidate =
            candidate.map_err(|error| invalid_error(format!("invalid ODS attribute: {error}")))?;
        let value = candidate
            .decoded_and_normalized_value(quick_xml::XmlVersion::Explicit1_0, reader.decoder())
            .map_err(|error| invalid_error(format!("invalid ODS attribute value: {error}")))?;
        attributes.push((
            String::from_utf8_lossy(candidate.key.as_ref()).into_owned(),
            value.into_owned(),
        ));
    }
    Ok((name, attributes, empty))
}

fn render_tag(name: &str, attributes: &[(String, String)], empty: bool) -> String {
    let mut out = format!("<{name}");
    for (key, value) in attributes {
        out.push_str(&format!(" {key}=\"{}\"", escape_xml(value)));
    }
    out.push_str(if empty { "/>" } else { ">" });
    out
}

/// A start tag with `name` set to `value` (or removed), rendered compactly.
fn set_attr(tag: &str, name: &str, value: Option<&str>) -> Result<String> {
    let (qname, mut attributes, empty) = tag_parts(tag)?;
    let position = attributes.iter().position(|(key, _)| key == name);
    match (position, value) {
        (Some(at), Some(value)) => attributes[at].1 = value.to_owned(),
        (Some(at), None) => {
            attributes.remove(at);
        },
        (None, Some(value)) => attributes.push((name.to_owned(), value.to_owned())),
        (None, None) => {},
    }
    Ok(render_tag(&qname, &attributes, empty))
}

// ---------------------------------------------------------------------------
// Conditional formats

fn format_markup(format: &Format) -> Result<String> {
    let mut markup = String::new();
    cf::write_conditional_formats(&mut markup, std::slice::from_ref(format))?;
    let inner = markup
        .strip_prefix("<calcext:conditional-formats>")
        .and_then(|rest| rest.strip_suffix("</calcext:conditional-formats>"))
        .ok_or_else(|| invalid_error("ODS conditional format markup is malformed"))?;
    Ok(inner.to_owned())
}

/// The named cell style a conditional format applies, created in
/// `styles.xml` when it differs from every existing one this editor wrote.
fn put_format_style(
    source: &[u8],
    style: &FormatStyle,
    max_output: usize,
) -> Result<(Vec<u8>, String)> {
    if style.is_empty() {
        return Ok((source.to_vec(), "Default".to_owned()));
    }
    let styles = part_xml(source, "styles.xml")?.ok_or_else(|| {
        Error::Refused(litchi_core::edit::Refusal::new(litchi_core::edit::RefusalKind::UnsupportedOperation, "ods.advanced.sheet_features", "an ODS package without styles.xml takes no conditional-format style (record 0327 follow-up: author styles.xml)"))
    })?;
    let spans = scan(&styles)?;
    let office_styles = spans
        .iter()
        .position(|span| is_element(span, OFFICE, "styles"))
        .ok_or_else(|| invalid_error("ODS styles.xml has no office:styles"))?;
    let mut properties = String::new();
    if let Some(color) = &style.background {
        properties.push_str(&format!(
            "<style:table-cell-properties fo:background-color=\"{}\"/>",
            escape_xml(color)
        ));
    }
    let mut text = String::new();
    if let Some(color) = &style.color {
        text.push_str(&format!(" fo:color=\"{}\"", escape_xml(color)));
    }
    if let Some(bold) = style.bold {
        text.push_str(if bold {
            " fo:font-weight=\"bold\""
        } else {
            " fo:font-weight=\"normal\""
        });
    }
    if let Some(italic) = style.italic {
        text.push_str(if italic {
            " fo:font-style=\"italic\""
        } else {
            " fo:font-style=\"normal\""
        });
    }
    if let Some(underline) = style.underline {
        text.push_str(if underline {
            " style:text-underline-style=\"solid\""
        } else {
            " style:text-underline-style=\"none\""
        });
    }
    if let Some(strike) = style.strike {
        text.push_str(if strike {
            " style:text-line-through-style=\"solid\""
        } else {
            " style:text-line-through-style=\"none\""
        });
    }
    if let Some(size) = style.size_pt {
        if !size.is_finite() || size <= 0.0 || size > 1_638.0 {
            return invalid("ODS conditional font size is out of range");
        }
        text.push_str(&format!(" fo:font-size=\"{size}pt\""));
    }
    if let Some(font) = &style.font {
        text.push_str(&format!(
            " style:font-name=\"{}\" fo:font-family=\"{}\"",
            escape_xml(font),
            escape_xml(font)
        ));
    }
    if !text.is_empty() {
        properties.push_str(&format!("<style:text-properties{text}/>"));
    }
    // Reuse an equal style this editor wrote.
    let mut highest = 0u32;
    for index in children(&spans, office_styles, STYLE, "style") {
        let Some(name) = attribute(&styles, &spans[index], b"style:name")? else {
            continue;
        };
        if let Some(number) = name
            .strip_prefix(LITCHI_CF_STYLE)
            .and_then(|number| number.parse::<u32>().ok())
        {
            highest = highest.max(number);
            let body = styles
                .get(spans[index].tag_end..spans[index].close_start)
                .unwrap_or("");
            if body == properties {
                return Ok((source.to_vec(), name));
            }
        }
    }
    let name = format!("{LITCHI_CF_STYLE}{}", highest + 1);
    let markup = format!(
        "<style:style xmlns:style=\"{STYLE}\" xmlns:fo=\"{FO}\" style:name=\"{name}\" style:family=\"table-cell\" style:parent-style-name=\"Default\">{properties}</style:style>"
    );
    let at = spans[office_styles].close_start;
    let bytes = splice_part(
        source,
        "styles.xml",
        vec![(at..at, markup.into_bytes())],
        max_output,
    )?;
    Ok((bytes, name))
}

/// The named cell styles of a package's `styles.xml`, as the formatting a
/// conditional format applies.
///
/// # Errors
///
/// Returns an error when the package or its `styles.xml` is malformed.
pub fn format_styles(package: &[u8]) -> Result<BTreeMap<String, FormatStyle>> {
    let mut out = BTreeMap::new();
    let Some(styles) = part_xml(package, "styles.xml")? else {
        return Ok(out);
    };
    let spans = scan(&styles)?;
    for (index, span) in spans.iter().enumerate() {
        if !is_element(span, STYLE, "style")
            || span
                .parent
                .is_none_or(|parent| !is_element(&spans[parent], OFFICE, "styles"))
            || attribute(&styles, span, b"style:family")?.as_deref() != Some("table-cell")
        {
            continue;
        }
        let Some(name) = attribute(&styles, span, b"style:name")? else {
            continue;
        };
        let mut style = FormatStyle::default();
        for child in children(&spans, index, STYLE, "table-cell-properties") {
            style.background = attribute(&styles, &spans[child], b"fo:background-color")?
                .filter(|color| color != "transparent");
        }
        for child in children(&spans, index, STYLE, "text-properties") {
            let child = &spans[child];
            let value = |name: &[u8]| attribute(&styles, child, name);
            style.color = value(b"fo:color")?;
            style.bold = value(b"fo:font-weight")?.map(|weight| {
                weight == "bold" || weight.parse::<u32>().is_ok_and(|weight| weight >= 600)
            });
            style.italic =
                value(b"fo:font-style")?.map(|font| font == "italic" || font == "oblique");
            style.underline = value(b"style:text-underline-style")?.map(|line| line != "none");
            style.strike = value(b"style:text-line-through-style")?.map(|line| line != "none");
            style.size_pt = value(b"fo:font-size")?.and_then(|size| {
                size.strip_suffix("pt")
                    .and_then(|size| size.trim().parse().ok())
            });
            style.font = value(b"style:font-name")?.or(value(b"fo:font-family")?);
        }
        out.insert(name, style);
    }
    Ok(out)
}

/// The conditional-format elements of a sheet.
fn format_spans(
    xml: &str,
    spans: &[Span],
    sheet: &str,
) -> Result<(usize, Option<usize>, Vec<usize>)> {
    let table = select_sheet(xml, spans, sheet)?;
    let containers = children(spans, table, CALCEXT, "conditional-formats");
    if containers.len() > 1 {
        return invalid("ODS sheet conditional-format container is duplicated");
    }
    let container = containers.first().copied();
    let formats = container.map_or_else(Vec::new, |container| {
        children(spans, container, CALCEXT, "conditional-format")
    });
    Ok((table, container, formats))
}

/// Add a conditional format; returns the package and its position.
pub(crate) fn add_format(
    source: &[u8],
    sheet: &str,
    format: &Format,
    style: &FormatStyle,
    max_output: usize,
) -> Result<(Vec<u8>, usize)> {
    let (source, style_name) = put_format_style(source, style, max_output)?;
    let format = styled(format, &style_name);
    let (xml, spans) = content(&source)?;
    let (table, container, formats) = format_spans(&xml, &spans, sheet)?;
    let markup = format_markup(&format)?;
    let edit = match container {
        Some(container) => {
            let at = spans[container].close_start;
            (at..at, bind(markup, &xml, "calcext", CALCEXT).into_bytes())
        },
        None => {
            let at = after_rows(&spans, table);
            let markup = bind(
                format!("<calcext:conditional-formats>{markup}</calcext:conditional-formats>"),
                &xml,
                "calcext",
                CALCEXT,
            );
            (at..at, markup.into_bytes())
        },
    };
    Ok((
        splice_part(&source, "content.xml", vec![edit], max_output)?,
        formats.len(),
    ))
}

/// Where a sheet extension element goes: after the last column, row or
/// shapes child (LibreOffice writes `calcext` elements after the rows).
fn after_rows(spans: &[Span], table: usize) -> usize {
    spans
        .iter()
        .filter(|span| span.parent == Some(table))
        .filter(|span| {
            span.namespace.as_deref() == Some(TABLE)
                && matches!(
                    span.local.as_str(),
                    "table-column"
                        | "table-columns"
                        | "table-column-group"
                        | "table-header-columns"
                        | "table-row"
                        | "table-row-group"
                        | "table-header-rows"
                        | "table-rows"
                        | "shapes"
                        | "table-source"
                        | "scenario"
                        | "title"
                        | "desc"
                )
                || is_element(span, OFFICE, "forms")
                || is_element(span, OFFICE, "dde-source")
        })
        .map(|span| span.end)
        .max()
        .unwrap_or(spans[table].tag_end)
}

/// The format with every condition's style set (the editor's style replaces
/// the placeholder a validated format carries) to `style`.
fn styled(format: &Format, style: &str) -> Format {
    let mut format = format.clone();
    for rule in &mut format.rules {
        match rule {
            cf::Rule::Condition(condition) => condition.apply_style_name = style.to_owned(),
            cf::Rule::DateIs(date) => date.style = style.to_owned(),
            _ => {},
        }
    }
    format
}

/// Replace the conditional format at `index`.
pub(crate) fn set_format(
    source: &[u8],
    sheet: &str,
    index: usize,
    format: &Format,
    style: &FormatStyle,
    max_output: usize,
) -> Result<Vec<u8>> {
    let (source, style_name) = put_format_style(source, style, max_output)?;
    let format = styled(format, &style_name);
    let (xml, spans) = content(&source)?;
    let (_, _, formats) = format_spans(&xml, &spans, sheet)?;
    let at = *formats
        .get(index)
        .ok_or_else(|| invalid_error("ODS conditional format is missing"))?;
    let span = &spans[at];
    let previous: Vec<String> = editor_styles(&xml, &spans, at)
        .into_iter()
        .filter(|name| *name != style_name)
        .collect();
    let markup = bind(format_markup(&format)?, &xml, "calcext", CALCEXT);
    let bytes = splice_part(
        &source,
        "content.xml",
        vec![(span.start..span.end, markup.into_bytes())],
        max_output,
    )?;
    drop_unused_styles(&bytes, &previous, max_output)
}

/// Remove the conditional format at `index` (its container with the last
/// one).
pub(crate) fn remove_format(
    source: &[u8],
    sheet: &str,
    index: usize,
    max_output: usize,
) -> Result<Vec<u8>> {
    let (xml, spans) = content(source)?;
    let (_, container, formats) = format_spans(&xml, &spans, sheet)?;
    let span = &spans[*formats
        .get(index)
        .ok_or_else(|| invalid_error("ODS conditional format is missing"))?];
    let range = match container {
        Some(container) if formats.len() == 1 => spans[container].start..spans[container].end,
        _ => span.start..span.end,
    };
    let styles_used: Vec<String> = editor_styles(&xml, &spans, formats[index]);
    let bytes = splice_part(source, "content.xml", vec![(range, Vec::new())], max_output)?;
    drop_unused_styles(&bytes, &styles_used, max_output)
}

fn editor_styles(xml: &str, spans: &[Span], format: usize) -> Vec<String> {
    spans
        .iter()
        .enumerate()
        .filter(|(index, span)| {
            super::is_descendant(spans, *index, format)
                && span.namespace.as_deref() == Some(CALCEXT)
        })
        .filter_map(|(_, span)| {
            attribute(xml, span, b"calcext:apply-style-name")
                .ok()
                .flatten()
                .or_else(|| attribute(xml, span, b"calcext:style").ok().flatten())
        })
        .filter(|name| name.starts_with(LITCHI_CF_STYLE))
        .collect()
}

/// Remove this editor's conditional-format styles nothing references any
/// more.
fn drop_unused_styles(source: &[u8], names: &[String], max_output: usize) -> Result<Vec<u8>> {
    if names.is_empty() {
        return Ok(source.to_vec());
    }
    let (xml, _) = content(source)?;
    let Some(styles) = part_xml(source, "styles.xml")? else {
        return Ok(source.to_vec());
    };
    let spans = scan(&styles)?;
    let mut edits = Vec::new();
    for name in names {
        if xml.contains(&format!("=\"{}\"", escape_xml(name))) {
            continue;
        }
        for (index, span) in spans.iter().enumerate() {
            if is_element(span, STYLE, "style")
                && attribute(&styles, span, b"style:name")?.as_deref() == Some(name.as_str())
            {
                let _ = index;
                edits.push((span.start..span.end, Vec::new()));
            }
        }
    }
    edits.dedup_by_key(|(range, _)| range.start);
    splice_part(source, "styles.xml", edits, max_output)
}

// ---------------------------------------------------------------------------
// Grid materialization

/// One group of consecutive logical indexes of a run that gets its own
/// element: its first index, length, start tag (without the repeat
/// attribute set) and body (everything after the start tag).
struct Group {
    first: u32,
    length: u32,
    tag: String,
    body: String,
}

/// Group consecutive indexes with equal keys.
fn groups<K: PartialEq + Clone>(hits: &BTreeMap<u32, K>) -> Vec<(u32, u32, K)> {
    let mut out: Vec<(u32, u32, K)> = Vec::new();
    for (index, key) in hits {
        match out.last_mut() {
            Some((first, length, held)) if *first + *length == *index && held == key => {
                *length += 1;
            },
            _ => out.push((*index, 1, key.clone())),
        }
    }
    out
}

/// The markup of a run split around `groups`: the parts of the run no group
/// covers keep the run's tag and body with their own repeat count.
fn run_markup(
    xml: &str,
    span: &Span,
    first: u32,
    repeat: u32,
    repeat_name: &str,
    groups: &[Group],
) -> Result<String> {
    let tag = xml
        .get(span.start..span.tag_end)
        .ok_or_else(|| invalid_error("ODS run tag is invalid"))?;
    let body = xml.get(span.tag_end..span.end).unwrap_or("");
    let repeated = |tag: &str, count: u32| {
        set_attr(
            tag,
            repeat_name,
            (count > 1).then(|| count.to_string()).as_deref(),
        )
    };
    let mut out = String::new();
    let mut cursor = first;
    for group in groups {
        if group.first > cursor {
            out.push_str(&repeated(tag, group.first - cursor)?);
            out.push_str(body);
        }
        out.push_str(&repeated(&group.tag, group.length)?);
        out.push_str(&group.body);
        cursor = group.first + group.length;
    }
    let last = first.saturating_add(repeat);
    if last > cursor {
        out.push_str(&repeated(tag, last - cursor)?);
        out.push_str(body);
    }
    Ok(out)
}

/// Edits that give the indexes of `hits` (within `runs`) the start tags
/// `retag` returns: a run one group covers whole keeps its body and only
/// its start tag changes; any other touched run is rewritten split.
fn retag_runs<K: PartialEq + Clone>(
    xml: &str,
    spans: &[Span],
    runs: &[RowRun],
    repeat_name: &str,
    hits: &BTreeMap<u32, K>,
    retag: &dyn Fn(&str, &K) -> Result<String>,
) -> Result<Vec<Splice>> {
    let mut edits = Vec::new();
    for run in runs {
        let last = run.first.saturating_add(run.repeat);
        let local: BTreeMap<u32, K> = hits
            .range(run.first..last)
            .map(|(index, key)| (*index, key.clone()))
            .collect();
        if local.is_empty() {
            continue;
        }
        let span = &spans[run.span];
        let tag = xml
            .get(span.start..span.tag_end)
            .ok_or_else(|| invalid_error("ODS run tag is invalid"))?;
        let grouped = groups(&local);
        if let [(first, length, key)] = grouped.as_slice()
            && *first == run.first
            && *length == run.repeat
        {
            edits.push((span.start..span.tag_end, retag(tag, key)?.into_bytes()));
            continue;
        }
        let body = xml.get(span.tag_end..span.end).unwrap_or("");
        let mut parts = Vec::new();
        for (first, length, key) in grouped {
            parts.push(Group {
                first,
                length,
                tag: retag(tag, &key)?,
                body: body.to_owned(),
            });
        }
        let markup = run_markup(xml, span, run.first, run.repeat, repeat_name, &parts)?;
        edits.push((span.start..span.end, markup.into_bytes()));
    }
    Ok(edits)
}

/// Normalized column spans of one row.
type Columns = Vec<(u32, u32)>;

/// Set (or clear) attribute `name` on every cell of `rects`, splitting the
/// row and cell runs that hold them.
fn set_cell_attribute(
    xml: &str,
    spans: &[Span],
    table: usize,
    rects: &[(u32, u32, u32, u32)],
    name: &str,
    value: Option<&str>,
) -> Result<Vec<Splice>> {
    let mut cells: u64 = 0;
    let mut rows: BTreeMap<u32, Columns> = BTreeMap::new();
    for rect in rects {
        cells =
            cells.saturating_add(u64::from(rect.2 - rect.0 + 1) * u64::from(rect.3 - rect.1 + 1));
        if cells > MAX_MATERIALIZED_CELLS {
            return invalid("ODS validation range exceeds the cell bound");
        }
        for row in rect.0..=rect.2 {
            rows.entry(row).or_default().push((rect.1, rect.3));
        }
    }
    for columns in rows.values_mut() {
        columns.sort_unstable();
        columns.dedup();
    }
    let runs = row_runs(xml, spans, table)?;
    let end = runs
        .last()
        .map_or(0, |run| run.first.saturating_add(run.repeat));
    let mut edits = Vec::new();
    // Rows past the sheet's last row are appended (only to bind cells).
    let beyond: BTreeMap<u32, Columns> = rows
        .range(end..)
        .map(|(row, columns)| (*row, columns.clone()))
        .collect();
    if value.is_some() && !beyond.is_empty() {
        let Some(last) = runs.last() else {
            return Err(Error::Refused(litchi_core::edit::Refusal::new(
                litchi_core::edit::RefusalKind::UnsupportedOperation,
                "ods.advanced.sheet_features",
                "an ODS sheet without rows takes no validation binding",
            )));
        };
        let mut markup = String::new();
        let mut cursor = end;
        for (first, length, columns) in groups(&beyond) {
            if first > cursor {
                markup.push_str(&row_markup(first - cursor, &[], name, value));
            }
            markup.push_str(&row_markup(length, &columns, name, value));
            cursor = first + length;
        }
        let at = spans[last.span].end;
        edits.push((at..at, markup.into_bytes()));
    }
    for run in &runs {
        let last = run.first.saturating_add(run.repeat);
        let local: BTreeMap<u32, Columns> = rows
            .range(run.first..last)
            .map(|(row, columns)| (*row, columns.clone()))
            .collect();
        if local.is_empty() {
            continue;
        }
        let span = &spans[run.span];
        let grouped = groups(&local);
        if let [(first, length, columns)] = grouped.as_slice()
            && *first == run.first
            && *length == run.repeat
        {
            // The whole run takes the same cells: edit its cells in place.
            edits.extend(cell_edits(xml, spans, run.span, columns, name, value)?);
            continue;
        }
        let tag = xml
            .get(span.start..span.tag_end)
            .ok_or_else(|| invalid_error("ODS row tag is invalid"))?;
        let mut parts = Vec::new();
        for (first, length, columns) in grouped {
            // The row's body with the cell edits applied.
            let mut body = xml.get(span.tag_end..span.end).unwrap_or("").to_owned();
            let mut local_edits = cell_edits(xml, spans, run.span, &columns, name, value)?;
            local_edits.sort_by_key(|(range, _)| std::cmp::Reverse(range.start));
            for (range, bytes) in local_edits {
                let start = range.start - span.tag_end;
                let stop = range.end - span.tag_end;
                body.replace_range(start..stop, &String::from_utf8_lossy(&bytes));
            }
            parts.push(Group {
                first,
                length,
                tag: tag.to_owned(),
                body,
            });
        }
        let markup = run_markup(
            xml,
            span,
            run.first,
            run.repeat,
            "table:number-rows-repeated",
            &parts,
        )?;
        edits.push((span.start..span.end, markup.into_bytes()));
    }
    Ok(edits)
}

/// A new row element repeated `repeat` times whose cells of `columns` carry
/// attribute `name`.
fn row_markup(repeat: u32, columns: &[(u32, u32)], name: &str, value: Option<&str>) -> String {
    let mut out = String::from("<table:table-row");
    if repeat > 1 {
        out.push_str(&format!(" table:number-rows-repeated=\"{repeat}\""));
    }
    out.push('>');
    out.push_str(&cells_markup(0, columns, name, value));
    if columns.is_empty() {
        out.push_str("<table:table-cell/>");
    }
    out.push_str("</table:table-row>");
    out
}

/// New cell elements from column `from`: padding, then the cells of
/// `columns` with attribute `name`.
fn cells_markup(from: u32, columns: &[(u32, u32)], name: &str, value: Option<&str>) -> String {
    let mut out = String::new();
    let mut cursor = from;
    let attribute = value.map_or_else(String::new, |value| {
        format!(" {name}=\"{}\"", escape_xml(value))
    });
    for (first, last) in columns {
        let first = (*first).max(cursor);
        if *last < first {
            continue;
        }
        if first > cursor {
            let gap = first - cursor;
            if gap > 1 {
                out.push_str(&format!(
                    "<table:table-cell table:number-columns-repeated=\"{gap}\"/>"
                ));
            } else {
                out.push_str("<table:table-cell/>");
            }
        }
        let count = last - first + 1;
        if count > 1 {
            out.push_str(&format!(
                "<table:table-cell{attribute} table:number-columns-repeated=\"{count}\"/>"
            ));
        } else {
            out.push_str(&format!("<table:table-cell{attribute}/>"));
        }
        cursor = last + 1;
    }
    out
}

/// The edits that set attribute `name` on the cells of `columns` of one row
/// (ranges are absolute; none lies in the row's start tag).
fn cell_edits(
    xml: &str,
    spans: &[Span],
    row: usize,
    columns: &[(u32, u32)],
    name: &str,
    value: Option<&str>,
) -> Result<Vec<Splice>> {
    let runs: Vec<RowRun> = cell_runs(xml, spans, row)?
        .into_iter()
        .map(|run| RowRun {
            span: run.span,
            first: run.first,
            repeat: run.repeat,
        })
        .collect();
    let width = runs
        .last()
        .map_or(0, |run| run.first.saturating_add(run.repeat));
    let mut hits: BTreeMap<u32, ()> = BTreeMap::new();
    let mut beyond: Vec<(u32, u32)> = Vec::new();
    for (first, last) in columns {
        for column in *first..=(*last).min(width.saturating_sub(1)) {
            if column < width {
                hits.insert(column, ());
            }
        }
        if *last >= width {
            beyond.push(((*first).max(width), *last));
        }
    }
    let retag = |tag: &str, (): &()| set_attr(tag, name, value);
    let mut edits = retag_runs(
        xml,
        spans,
        &runs,
        "table:number-columns-repeated",
        &hits,
        &retag,
    )?;
    if value.is_some() && !beyond.is_empty() {
        if is_self_closing_span(xml, &spans[row])? {
            return Err(Error::Refused(litchi_core::edit::Refusal::new(
                litchi_core::edit::RefusalKind::UnsupportedOperation,
                "ods.advanced.sheet_features",
                "an empty ODS row element takes no cells",
            )));
        }
        let at = spans[row].close_start;
        edits.push((
            at..at,
            cells_markup(width, &beyond, name, value).into_bytes(),
        ));
    }
    Ok(edits)
}
// ---------------------------------------------------------------------------
// Validations

fn validation_markup(name: &str, sheet: &str, spec: &ValidationSpec) -> String {
    let base = spec.ranges.first().map_or_else(String::new, |rect| {
        address(sheet, (rect.0, rect.1, rect.0, rect.1))
    });
    let condition = if spec.condition.is_empty() {
        String::new()
    } else {
        format!(" table:condition=\"{}\"", escape_xml(&spec.condition))
    };
    let mut markup = format!(
        "<table:content-validation table:name=\"{}\"{condition} table:allow-empty-cell=\"{}\" table:display-list=\"{}\" table:base-cell-address=\"{}\">",
        escape_xml(name),
        spec.allow_empty,
        if spec.display_list {
            "unsorted"
        } else {
            "none"
        },
        escape_xml(&base)
    );
    let paragraph = |text: &Option<String>| {
        text.as_ref().map_or_else(String::new, |text| {
            format!("<text:p>{}</text:p>", escape_xml(text))
        })
    };
    let (title, text, display) = &spec.help;
    if title.is_some() || text.is_some() || *display {
        markup.push_str("<table:help-message");
        if let Some(title) = title {
            markup.push_str(&format!(" table:title=\"{}\"", escape_xml(title)));
        }
        markup.push_str(&format!(
            " table:display=\"{display}\">{}</table:help-message>",
            paragraph(text)
        ));
    }
    let (kind, title, text, display) = &spec.error;
    markup.push_str(&format!(
        "<table:error-message table:message-type=\"{}\"",
        escape_xml(kind)
    ));
    if let Some(title) = title {
        markup.push_str(&format!(" table:title=\"{}\"", escape_xml(title)));
    }
    markup.push_str(&format!(
        " table:display=\"{display}\">{}</table:error-message>",
        paragraph(text)
    ));
    markup.push_str("</table:content-validation>");
    markup
}

/// The content-validation definitions of a document, by name.
fn definitions(xml: &str, spans: &[Span]) -> (Option<usize>, Vec<(usize, String)>) {
    let container = descendants_of(spans, OFFICE, "spreadsheet", TABLE, "content-validations")
        .first()
        .copied();
    let entries = container.map_or_else(Vec::new, |container| {
        children(spans, container, TABLE, "content-validation")
            .into_iter()
            .filter_map(|index| {
                attribute(xml, &spans[index], b"table:name")
                    .ok()
                    .flatten()
                    .map(|name| (index, name))
            })
            .collect()
    });
    (container, entries)
}

/// The definitions bound in a sheet, in catalog order.
fn sheet_validations(xml: &str, spans: &[Span], sheet: &str) -> Result<Vec<Validation>> {
    Ok(read_spans(xml, spans, sheet)?.validations)
}

/// Add a validation bound to `spec.ranges`; returns the package and its
/// position among the sheet's validations.
pub(crate) fn add_validation(
    source: &[u8],
    sheet: &str,
    spec: &ValidationSpec,
    max_output: usize,
) -> Result<(Vec<u8>, usize)> {
    let (xml, spans) = content(source)?;
    let table = select_sheet(&xml, &spans, sheet)?;
    let (container, entries) = definitions(&xml, &spans);
    let mut number = entries.len() + 1;
    while entries
        .iter()
        .any(|(_, name)| *name == format!("val{number}"))
    {
        number += 1;
    }
    let name = format!("val{number}");
    let markup = bind(validation_markup(&name, sheet, spec), &xml, "text", TEXT);
    let mut edits = set_cell_attribute(
        &xml,
        &spans,
        table,
        &spec.ranges,
        "table:content-validation-name",
        Some(&name),
    )?;
    match container {
        Some(container) if !is_self_closing_span(&xml, &spans[container])? => {
            let at = spans[container].close_start;
            edits.push((at..at, markup.into_bytes()));
        },
        Some(_) => return invalid("ODS content-validations container is empty"),
        None => {
            let spreadsheet = spans
                .iter()
                .position(|span| is_element(span, OFFICE, "spreadsheet"))
                .ok_or_else(|| invalid_error("ODS content has no office:spreadsheet"))?;
            // Before the label ranges and the first table (ODF 1.3 9.1.2).
            let at = spans
                .iter()
                .filter(|span| span.parent == Some(spreadsheet))
                .filter(|span| {
                    is_element(span, TABLE, "table") || is_element(span, TABLE, "label-ranges")
                })
                .map(|span| span.start)
                .min()
                .unwrap_or(spans[spreadsheet].close_start);
            let markup = bind(
                bind(
                    format!("<table:content-validations>{markup}</table:content-validations>"),
                    &xml,
                    "table",
                    TABLE,
                ),
                &xml,
                "text",
                TEXT,
            );
            edits.push((at..at, markup.into_bytes()));
        },
    }
    let position = sheet_validations(&xml, &spans, sheet)?.len();
    Ok((
        splice_part(source, "content.xml", edits, max_output)?,
        position,
    ))
}

/// Replace the validation at `index` of a sheet (its definition and its
/// bindings in this sheet).
pub(crate) fn set_validation(
    source: &[u8],
    sheet: &str,
    index: usize,
    spec: &ValidationSpec,
    max_output: usize,
) -> Result<Vec<u8>> {
    let (xml, spans) = content(source)?;
    let table = select_sheet(&xml, &spans, sheet)?;
    let current = sheet_validations(&xml, &spans, sheet)?;
    let held = current
        .get(index)
        .ok_or_else(|| invalid_error("ODS validation is missing"))?;
    let (_, entries) = definitions(&xml, &spans);
    let definition = entries
        .iter()
        .find(|(_, name)| *name == held.name)
        .map(|(span, _)| *span)
        .ok_or_else(|| invalid_error("ODS validation definition is missing"))?;
    if used_elsewhere(&xml, &spans, table, &held.name) {
        return Err(Error::Refused(litchi_core::edit::Refusal::new(
            litchi_core::edit::RefusalKind::UnsupportedOperation,
            "ods.advanced.sheet_features",
            "an ODS validation another sheet binds is not rewritten",
        )));
    }
    let markup = bind(
        validation_markup(&held.name, sheet, spec),
        &xml,
        "text",
        TEXT,
    );
    let mut edits = vec![(
        spans[definition].start..spans[definition].end,
        markup.into_bytes(),
    )];
    if spec.ranges != held.ranges {
        // Rebind: clear the old cells, then bind the new ones.
        let cleared = splice_part(
            source,
            "content.xml",
            set_cell_attribute(
                &xml,
                &spans,
                table,
                &held.ranges,
                "table:content-validation-name",
                None,
            )?,
            max_output,
        )?;
        let (xml, spans) = content(&cleared)?;
        let table = select_sheet(&xml, &spans, sheet)?;
        let (_, entries) = definitions(&xml, &spans);
        let definition = entries
            .iter()
            .find(|(_, name)| *name == held.name)
            .map(|(span, _)| *span)
            .ok_or_else(|| invalid_error("ODS validation definition is missing"))?;
        let mut edits = set_cell_attribute(
            &xml,
            &spans,
            table,
            &spec.ranges,
            "table:content-validation-name",
            Some(&held.name),
        )?;
        edits.push((
            spans[definition].start..spans[definition].end,
            bind(
                validation_markup(&held.name, sheet, spec),
                &xml,
                "text",
                TEXT,
            )
            .into_bytes(),
        ));
        return splice_part(&cleared, "content.xml", edits, max_output);
    }
    edits.truncate(1);
    splice_part(source, "content.xml", edits, max_output)
}

fn used_elsewhere(xml: &str, spans: &[Span], table: usize, name: &str) -> bool {
    let needle = format!("table:content-validation-name=\"{}\"", escape_xml(name));
    let table_span = &spans[table];
    xml[..table_span.start].contains(&needle) || xml[table_span.end..].contains(&needle)
}

/// Remove the validation at `index` of a sheet: its bindings, and its
/// definition when no other sheet binds it (the container with the last).
pub(crate) fn remove_validation(
    source: &[u8],
    sheet: &str,
    index: usize,
    max_output: usize,
) -> Result<Vec<u8>> {
    let (xml, spans) = content(source)?;
    let table = select_sheet(&xml, &spans, sheet)?;
    let current = sheet_validations(&xml, &spans, sheet)?;
    let held = current
        .get(index)
        .ok_or_else(|| invalid_error("ODS validation is missing"))?;
    let mut edits = set_cell_attribute(
        &xml,
        &spans,
        table,
        &held.ranges,
        "table:content-validation-name",
        None,
    )?;
    if !used_elsewhere(&xml, &spans, table, &held.name) {
        let (container, entries) = definitions(&xml, &spans);
        if let Some((definition, _)) = entries.iter().find(|(_, name)| *name == held.name) {
            let range = match container {
                Some(container) if entries.len() == 1 => {
                    spans[container].start..spans[container].end
                },
                _ => spans[*definition].start..spans[*definition].end,
            };
            edits.push((range, Vec::new()));
        }
    }
    splice_part(source, "content.xml", edits, max_output)
}

// ---------------------------------------------------------------------------
// Auto-filter

fn spreadsheet(spans: &[Span]) -> Result<usize> {
    spans
        .iter()
        .position(|span| is_element(span, OFFICE, "spreadsheet"))
        .ok_or_else(|| invalid_error("ODS content has no office:spreadsheet"))
}

/// The sheet's auto-filter database range: a range with filter buttons on
/// the sheet, or the sheet's anonymous database range.
fn auto_filter_span(
    xml: &str,
    spans: &[Span],
    sheet: &str,
    anonymous: &str,
) -> Result<(Option<usize>, Option<usize>, usize)> {
    let container = descendants_of(spans, OFFICE, "spreadsheet", TABLE, "database-ranges")
        .first()
        .copied();
    let mut found = None;
    let mut count = 0;
    if let Some(container) = container {
        for index in children(spans, container, TABLE, "database-range") {
            count += 1;
            let span = &spans[index];
            let owner = attribute(xml, span, b"table:target-range-address")?
                .and_then(|address| parse_address(&address))
                .and_then(|(owner, _)| owner);
            if owner.as_deref() != Some(sheet) {
                continue;
            }
            let buttons =
                attribute(xml, span, b"table:display-filter-buttons")?.as_deref() == Some("true");
            let name = attribute(xml, span, b"table:name")?;
            if buttons || name.as_deref() == Some(anonymous) {
                if found.is_some() {
                    return Err(Error::Refused(litchi_core::edit::Refusal::new(
                        litchi_core::edit::RefusalKind::UnsupportedOperation,
                        "ods.advanced.sheet_features",
                        "an ODS sheet with two filtered database ranges has no single auto-filter",
                    )));
                }
                found = Some(index);
            }
        }
    }
    Ok((container, found, count))
}

/// Set the sheet's auto-filter database range to `range`, or remove it.
pub(crate) fn set_auto_filter(
    source: &[u8],
    sheet: &str,
    range: Option<(u32, u32, u32, u32)>,
    max_output: usize,
) -> Result<Vec<u8>> {
    let (xml, spans) = content(source)?;
    let root = spreadsheet(&spans)?;
    let table = select_sheet(&xml, &spans, sheet)?;
    let sheet_index = children(&spans, root, TABLE, "table")
        .iter()
        .position(|index| *index == table)
        .unwrap_or(0);
    let anonymous = format!("__Anonymous_Sheet_DB__{sheet_index}");
    let (container, existing, count) = auto_filter_span(&xml, &spans, sheet, &anonymous)?;
    let edit = match (range, existing) {
        (None, None) => return Ok(source.to_vec()),
        (None, Some(index)) => {
            let container =
                container.ok_or_else(|| invalid_error("ODS database range has no container"))?;
            let removed = if count == 1 {
                &spans[container]
            } else {
                &spans[index]
            };
            (removed.start..removed.end, Vec::new())
        },
        (Some(rect), Some(index)) => {
            let span = &spans[index];
            let tag = xml
                .get(span.start..span.tag_end)
                .ok_or_else(|| invalid_error("ODS database range tag is invalid"))?;
            let tag = set_attr(
                tag,
                "table:target-range-address",
                Some(&address(sheet, rect)),
            )?;
            let tag = set_attr(&tag, "table:display-filter-buttons", Some("true"))?;
            (span.start..span.tag_end, tag.into_bytes())
        },
        (Some(rect), None) => {
            let markup = format!(
                "<table:database-range table:name=\"{anonymous}\" table:target-range-address=\"{}\" table:display-filter-buttons=\"true\"/>",
                escape_xml(&address(sheet, rect))
            );
            if let Some(container) = container {
                let at = spans[container].close_start;
                if is_self_closing_span(&xml, &spans[container])? {
                    let markup = format!("<table:database-ranges>{markup}</table:database-ranges>");
                    (
                        spans[container].start..spans[container].end,
                        markup.into_bytes(),
                    )
                } else {
                    (at..at, markup.into_bytes())
                }
            } else {
                // After the tables and named expressions (ODF 1.3 9.1.2 order).
                let at = spans
                    .iter()
                    .filter(|span| span.parent == Some(root))
                    .filter(|span| {
                        span.namespace.as_deref() == Some(TABLE)
                            && matches!(
                                span.local.as_str(),
                                "table"
                                    | "named-expressions"
                                    | "label-ranges"
                                    | "content-validations"
                                    | "calculation-settings"
                            )
                    })
                    .map(|span| span.end)
                    .max()
                    .unwrap_or(spans[root].close_start);
                let markup = bind(
                    format!("<table:database-ranges>{markup}</table:database-ranges>"),
                    &xml,
                    "table",
                    TABLE,
                );
                (at..at, markup.into_bytes())
            }
        },
    };
    splice_part(source, "content.xml", vec![edit], max_output)
}

// ---------------------------------------------------------------------------
// Page breaks

/// The automatic styles element of `content.xml` and the names it holds.
fn automatic_styles(xml: &str, spans: &[Span]) -> (Option<usize>, Vec<String>) {
    let container = spans.iter().position(|span| {
        is_element(span, OFFICE, "automatic-styles")
            && span
                .parent
                .is_some_and(|parent| spans[parent].parent.is_none())
    });
    let names = container.map_or_else(Vec::new, |container| {
        children(spans, container, STYLE, "style")
            .into_iter()
            .filter_map(|index| attribute(xml, &spans[index], b"style:name").ok().flatten())
            .collect()
    });
    (container, names)
}

/// The edit that appends `markup` to the content's automatic styles.
fn append_automatic_styles(xml: &str, spans: &[Span], markup: String) -> Result<Splice> {
    let (container, _) = automatic_styles(xml, spans);
    match container {
        Some(container) if !is_self_closing_span(xml, &spans[container])? => {
            let at = spans[container].close_start;
            Ok((at..at, markup.into_bytes()))
        },
        Some(container) => Ok((
            spans[container].start..spans[container].end,
            format!("<office:automatic-styles>{markup}</office:automatic-styles>").into_bytes(),
        )),
        None => {
            let body = spans
                .iter()
                .position(|span| is_element(span, OFFICE, "body"))
                .ok_or_else(|| invalid_error("ODS content has no office:body"))?;
            let at = spans[body].start;
            Ok((
                at..at,
                format!("<office:automatic-styles>{markup}</office:automatic-styles>").into_bytes(),
            ))
        },
    }
}

/// The most indexes one break edit restyles.
const MAX_BREAK_INDEXES: usize = 4_096;

/// Replace the manual breaks of one axis: the rows (or columns) of
/// `breaks` get a break style derived from their current one, and every
/// other row (column) whose style breaks gets the twin without the break.
pub(crate) fn set_breaks(
    source: &[u8],
    sheet: &str,
    rows: bool,
    breaks: &[u32],
    max_output: usize,
) -> Result<Vec<u8>> {
    let (xml, spans) = content(source)?;
    let table = select_sheet(&xml, &spans, sheet)?;
    let (family, properties, repeat_name, runs) = if rows {
        (
            "table-row",
            "table-row-properties",
            "table:number-rows-repeated",
            row_runs(&xml, &spans, table)?,
        )
    } else {
        (
            "table-column",
            "table-column-properties",
            "table:number-columns-repeated",
            column_runs(&xml, &spans, table)?,
        )
    };
    let limit = if rows { MAX_ROWS } else { MAX_COLUMNS };
    if breaks.iter().any(|index| *index == 0 || *index >= limit) {
        return invalid("ODS page break is outside the sheet");
    }
    if breaks.len() > MAX_BREAK_INDEXES {
        return invalid("ODS page break count exceeds the edit bound");
    }
    let styles = style_properties(&xml, &spans, family, properties);
    let style_of = |run: &RowRun| -> Result<Option<String>> {
        attribute(&xml, &spans[run.span], b"table:style-name")
    };
    let breaking = |name: &Option<String>| {
        name.as_ref()
            .and_then(|name| styles.get(name))
            .is_some_and(|tag| attr_in(tag, "fo:break-before").as_deref() == Some("page"))
    };
    let wanted: std::collections::BTreeSet<u32> = breaks.iter().copied().collect();
    let end = runs
        .last()
        .map_or(0, |run| run.first.saturating_add(run.repeat));
    // The indexes whose state changes, with their run's base style; a
    // break past the last row or column appends elements.
    let mut touched: BTreeMap<u32, (Option<String>, bool)> = BTreeMap::new();
    let beyond: Vec<u32> = wanted.range(end..).copied().collect();
    if !beyond.is_empty() && runs.is_empty() && rows {
        return Err(Error::Refused(litchi_core::edit::Refusal::new(
            litchi_core::edit::RefusalKind::UnsupportedOperation,
            "ods.advanced.sheet_features",
            "an ODS sheet without rows takes no row page break",
        )));
    }
    for index in &beyond {
        touched.insert(*index, (None, true));
    }
    for run in &runs {
        let base = style_of(run)?;
        let last = run.first.saturating_add(run.repeat);
        if breaking(&base) {
            if run.repeat as usize > MAX_BREAK_INDEXES {
                return Err(Error::Refused(litchi_core::edit::Refusal::new(
                    litchi_core::edit::RefusalKind::UnsupportedOperation,
                    "ods.advanced.sheet_features",
                    "an ODS run of more than 4096 breaking rows or columns is not restyled",
                )));
            }
            for index in run.first..last {
                if !wanted.contains(&index) {
                    touched.insert(index, (base.clone(), false));
                }
            }
        } else {
            for index in wanted.range(run.first..last) {
                touched.insert(*index, (base.clone(), true));
            }
        }
        if touched.len() > MAX_BREAK_INDEXES * 2 {
            return invalid("ODS page break edit exceeds its bound");
        }
    }
    if touched.is_empty() {
        return Ok(source.to_vec());
    }
    // A derived style per (base, break) pair, reusing an equal one.
    let (_, names) = automatic_styles(&xml, &spans);
    let mut created: BTreeMap<(Option<String>, bool), String> = BTreeMap::new();
    let mut new_styles = String::new();
    let mut next = 1u32;
    for (base, page) in touched.values() {
        if created.contains_key(&(base.clone(), *page)) {
            continue;
        }
        let base_tag = base
            .as_ref()
            .and_then(|name| styles.get(name))
            .cloned()
            .unwrap_or_default();
        let mut attributes: Vec<(String, String)> = if base_tag.is_empty() {
            Vec::new()
        } else {
            tag_parts(&base_tag)?.1
        };
        attributes.retain(|(key, _)| key != "fo:break-before");
        if *page {
            attributes.push(("fo:break-before".to_owned(), "page".to_owned()));
        }
        let wanted_tag = render_tag(&format!("style:{properties}"), &attributes, true);
        let reuse = styles.iter().find(|(name, tag)| {
            name.starts_with(LITCHI_BREAK_STYLE)
                && tag_parts(tag).is_ok_and(|(_, mut held, _)| {
                    let mut expected = attributes.clone();
                    held.sort();
                    expected.sort();
                    held == expected
                })
        });
        let name = if let Some((name, _)) = reuse {
            name.clone()
        } else if !*page && attributes.is_empty() {
            // No properties left: the default style.
            String::new()
        } else {
            let name = loop {
                let candidate = format!("{LITCHI_BREAK_STYLE}{next}");
                next += 1;
                if !names.contains(&candidate) && !styles.contains_key(&candidate) {
                    break candidate;
                }
            };
            new_styles.push_str(&bind(
                bind(
                    format!(
                        "<style:style style:name=\"{name}\" style:family=\"{family}\">{wanted_tag}</style:style>"
                    ),
                    &xml,
                    "style",
                    STYLE,
                ),
                &xml,
                "fo",
                FO,
            ));
            name
        };
        created.insert((base.clone(), *page), name);
    }
    let plan: BTreeMap<u32, String> = touched
        .iter()
        .map(|(index, key)| (*index, created[key].clone()))
        .collect();
    let retag = |tag: &str, name: &String| {
        set_attr(
            tag,
            "table:style-name",
            (!name.is_empty()).then_some(name.as_str()),
        )
    };
    let mut edits = retag_runs(&xml, &spans, &runs, repeat_name, &plan, &retag)?;
    if !beyond.is_empty() {
        let element = |count: u32, style: Option<&str>| -> String {
            let style =
                style.map_or_else(String::new, |name| format!(" table:style-name=\"{name}\""));
            let repeat = if count > 1 {
                format!(" {repeat_name}=\"{count}\"")
            } else {
                String::new()
            };
            if rows {
                format!("<table:table-row{style}{repeat}><table:table-cell/></table:table-row>")
            } else {
                format!("<table:table-column{style}{repeat}/>")
            }
        };
        let mut markup = String::new();
        let mut cursor = end;
        for index in &beyond {
            if *index > cursor {
                markup.push_str(&element(index - cursor, None));
            }
            markup.push_str(&element(
                1,
                plan.get(index)
                    .map(String::as_str)
                    .filter(|name| !name.is_empty()),
            ));
            cursor = index + 1;
        }
        let at = match runs.last() {
            Some(last) => spans[last.span].end,
            // Columns precede the rows.
            None => spans
                .iter()
                .filter(|span| span.parent == Some(table))
                .find(|span| {
                    span.namespace.as_deref() == Some(TABLE)
                        && matches!(
                            span.local.as_str(),
                            "table-row" | "table-row-group" | "table-header-rows" | "table-rows"
                        )
                })
                .map_or(spans[table].close_start, |span| span.start),
        };
        edits.push((at..at, markup.into_bytes()));
    }
    if !new_styles.is_empty() {
        edits.push(append_automatic_styles(&xml, &spans, new_styles)?);
    }
    let bytes = splice_part(source, "content.xml", edits, max_output)?;
    drop_unreferenced_break_styles(&bytes, max_output)
}

/// Remove this editor's break styles no row or column references any more.
fn drop_unreferenced_break_styles(source: &[u8], max_output: usize) -> Result<Vec<u8>> {
    let (xml, spans) = content(source)?;
    let (container, names) = automatic_styles(&xml, &spans);
    let Some(container) = container else {
        return Ok(source.to_vec());
    };
    let mut edits = Vec::new();
    for name in names
        .iter()
        .filter(|name| name.starts_with(LITCHI_BREAK_STYLE))
    {
        if xml.contains(&format!("table:style-name=\"{}\"", escape_xml(name))) {
            continue;
        }
        for child in children(&spans, container, STYLE, "style") {
            if attribute(&xml, &spans[child], b"style:name")?.as_deref() == Some(name.as_str()) {
                edits.push((spans[child].start..spans[child].end, Vec::new()));
            }
        }
    }
    let all = spans
        .iter()
        .filter(|span| span.parent == Some(container))
        .count();
    if !edits.is_empty() && edits.len() == all {
        edits = vec![(spans[container].start..spans[container].end, Vec::new())];
    }
    splice_part(source, "content.xml", edits, max_output)
}

// ---------------------------------------------------------------------------
// Drawings

fn cm(value: f64) -> String {
    let rounded = (value * 1_000.0).round() / 1_000.0;
    format!("{rounded}cm")
}

/// The sheet's drawings of `kind`, in document order.
fn drawing_spans(xml: &str, spans: &[Span], sheet: &str, kind: DrawingKind) -> Result<Vec<usize>> {
    let table = select_sheet(xml, spans, sheet)?;
    let positions = cell_positions(xml, spans, table)?;
    let mut out = Vec::new();
    for (index, span) in spans.iter().enumerate() {
        if span.start < spans[table].start
            || span.end > spans[table].end
            || span.namespace.as_deref() != Some(DRAW)
        {
            continue;
        }
        let Some(parent) = span.parent else { continue };
        let anchored = (is_element(&spans[parent], TABLE, "shapes")
            && spans[parent].parent == Some(table))
            || (is_element(&spans[parent], TABLE, "table-cell") && positions.contains_key(&parent));
        if !anchored {
            continue;
        }
        let this = match span.local.as_str() {
            "frame" if !children(spans, index, DRAW, "image").is_empty() => DrawingKind::Picture,
            "custom-shape" | "rect" | "ellipse" => DrawingKind::Shape,
            _ => continue,
        };
        if this == kind {
            out.push(index);
        }
    }
    Ok(out)
}

fn validate_geometry(frame: &[f64; 4]) -> Result<()> {
    if frame
        .iter()
        .any(|value| !value.is_finite() || value.abs() > 10_000.0)
        || frame[2] <= 0.0
        || frame[3] <= 0.0
    {
        return invalid("ODS drawing geometry is invalid");
    }
    Ok(())
}

/// Every `draw:name` of the sheet's drawings.
fn drawing_names(xml: &str, spans: &[Span], sheet: &str) -> Result<Vec<String>> {
    let table = select_sheet(xml, spans, sheet)?;
    Ok(spans
        .iter()
        .enumerate()
        .filter(|(index, span)| {
            span.namespace.as_deref() == Some(DRAW) && super::is_descendant(spans, *index, table)
        })
        .filter_map(|(_, span)| attribute(xml, span, b"draw:name").ok().flatten())
        .collect())
}

/// Add a picture frame or custom shape to the sheet's `table:shapes`;
/// returns the package and its position among the sheet's drawings of its
/// kind.
pub(crate) fn add_drawing(
    source: &[u8],
    sheet: &str,
    spec: &DrawingSpec,
    max_output: usize,
) -> Result<(Vec<u8>, usize)> {
    let (xml, spans) = content(source)?;
    let table = select_sheet(&xml, &spans, sheet)?;
    validate_geometry(&spec.frame)?;
    if spec.name.is_empty() || drawing_names(&xml, &spans, sheet)?.contains(&spec.name) {
        return invalid("ODS drawing name is empty or already used on the sheet");
    }
    let picture = spec.image.is_some();
    let kind = if picture {
        DrawingKind::Picture
    } else {
        DrawingKind::Shape
    };
    let position = drawing_spans(&xml, &spans, sheet, kind)?.len();
    let mut attributes = format!(
        " draw:name=\"{}\" svg:x=\"{}\" svg:y=\"{}\" svg:width=\"{}\" svg:height=\"{}\"",
        escape_xml(&spec.name),
        cm(spec.frame[0]),
        cm(spec.frame[1]),
        cm(spec.frame[2]),
        cm(spec.frame[3])
    );
    if let Some((row, column)) = spec.end_cell {
        attributes.push_str(&format!(
            " table:end-cell-address=\"{}\"",
            escape_xml(&address(sheet, (row, column, row, column)))
        ));
    }
    let mut edits = Vec::new();
    if !picture && (spec.fill.is_some() || spec.line.is_some()) {
        let (_, names) = automatic_styles(&xml, &spans);
        let mut number = 1;
        let name = loop {
            let candidate = format!("{LITCHI_SHAPE_STYLE}{number}");
            number += 1;
            if !names.contains(&candidate) {
                break candidate;
            }
        };
        let mut graphic = String::new();
        match &spec.fill {
            Some(Some(color)) => graphic.push_str(&format!(
                " draw:fill=\"solid\" draw:fill-color=\"{}\"",
                escape_xml(color)
            )),
            Some(None) => graphic.push_str(" draw:fill=\"none\""),
            None => {},
        }
        match &spec.line {
            Some(Some((color, width))) => graphic.push_str(&format!(
                " draw:stroke=\"solid\" svg:stroke-color=\"{}\" svg:stroke-width=\"{}\"",
                escape_xml(color),
                escape_xml(width)
            )),
            Some(None) => graphic.push_str(" draw:stroke=\"none\""),
            None => {},
        }
        let markup = bind(
            bind(
                format!(
                    "<style:style style:name=\"{name}\" style:family=\"graphic\"><style:graphic-properties{graphic}/></style:style>"
                ),
                &xml,
                "draw",
                DRAW,
            ),
            &xml,
            "svg",
            SVG,
        );
        let markup = bind(markup, &xml, "style", STYLE);
        edits.push(append_automatic_styles(&xml, &spans, markup)?);
        attributes.push_str(&format!(" draw:style-name=\"{name}\""));
    }
    let mut body = String::new();
    if let Some(title) = &spec.title {
        body.push_str(&format!("<svg:title>{}</svg:title>", escape_xml(title)));
    }
    if let Some(description) = &spec.description {
        body.push_str(&format!("<svg:desc>{}</svg:desc>", escape_xml(description)));
    }
    let element = if let Some(image) = &spec.image {
        crate::document::validate_detached_resource_path(image)?;
        format!(
            "<draw:frame{attributes}><draw:image xlink:href=\"{}\" xlink:type=\"simple\" xlink:show=\"embed\" xlink:actuate=\"onLoad\"/>{body}</draw:frame>",
            escape_xml(image)
        )
    } else {
        let mut text = String::new();
        if let Some(value) = &spec.text {
            for line in value.split('\n') {
                if line.is_empty() {
                    text.push_str("<text:p/>");
                } else {
                    text.push_str(&format!("<text:p>{}</text:p>", escape_text(line)));
                }
            }
        }
        let geometry = spec.geometry.as_deref().unwrap_or("rectangle");
        format!(
            "<draw:custom-shape{attributes}>{body}{text}<draw:enhanced-geometry svg:viewBox=\"0 0 21600 21600\" draw:type=\"{}\"/></draw:custom-shape>",
            escape_xml(geometry)
        )
    };
    let mut element = element;
    for (prefix, namespace) in [
        ("draw", DRAW),
        ("svg", SVG),
        ("xlink", XLINK),
        ("text", TEXT),
        ("table", TABLE),
    ] {
        element = bind(element, &xml, prefix, namespace);
    }
    let shapes = children(&spans, table, TABLE, "shapes");
    if shapes.len() > 1 {
        return invalid("ODS sheet has more than one shapes container");
    }
    match shapes.first() {
        Some(shapes) if !is_self_closing_span(&xml, &spans[*shapes])? => {
            let at = spans[*shapes].close_start;
            edits.push((at..at, element.into_bytes()));
        },
        Some(_) => return invalid("ODS sheet shapes container is empty"),
        None => {
            // `table:shapes` precedes the columns and rows (ODF 1.3 9.1.2).
            let at = spans
                .iter()
                .filter(|span| span.parent == Some(table))
                .find(|span| {
                    span.namespace.as_deref() == Some(TABLE)
                        && matches!(
                            span.local.as_str(),
                            "table-column"
                                | "table-columns"
                                | "table-column-group"
                                | "table-header-columns"
                                | "table-row"
                                | "table-row-group"
                                | "table-header-rows"
                                | "table-rows"
                        )
                })
                .map_or(spans[table].close_start, |span| span.start);
            let markup = bind(
                format!("<table:shapes>{element}</table:shapes>"),
                &xml,
                "table",
                TABLE,
            );
            edits.push((at..at, markup.into_bytes()));
        },
    }
    Ok((
        splice_part(source, "content.xml", edits, max_output)?,
        position,
    ))
}

/// Text escaped for character data, spaces kept as `text:s` where a run
/// would be collapsed.
fn escape_text(text: &str) -> String {
    let mut out = String::new();
    let mut previous_space = true;
    for char in text.chars() {
        match char {
            ' ' if previous_space => out.push_str("<text:s/>"),
            ' ' => out.push(' '),
            '\t' => out.push_str("<text:tab/>"),
            other => out.push_str(&escape_xml(&other.to_string())),
        }
        previous_space = char == ' ';
    }
    if out.ends_with(' ') {
        out.pop();
        out.push_str("<text:s/>");
    }
    out
}

/// Change a drawing's frame, end cell, name or alternative text.
pub(crate) fn set_drawing(
    source: &[u8],
    sheet: &str,
    kind: DrawingKind,
    index: usize,
    change: &DrawingChange,
    max_output: usize,
) -> Result<Vec<u8>> {
    let (xml, spans) = content(source)?;
    let drawings = drawing_spans(&xml, &spans, sheet, kind)?;
    let at = *drawings
        .get(index)
        .ok_or_else(|| invalid_error("ODS drawing is missing"))?;
    let span = &spans[at];
    let mut tag = xml
        .get(span.start..span.tag_end)
        .ok_or_else(|| invalid_error("ODS drawing tag is invalid"))?
        .to_owned();
    let original = tag.clone();
    if let Some(frame) = change.frame {
        validate_geometry(&frame)?;
        for (name, value) in ["svg:x", "svg:y", "svg:width", "svg:height"]
            .iter()
            .zip(frame)
        {
            tag = set_attr(&tag, name, Some(&cm(value)))?;
        }
    }
    if let Some(end) = change.end_cell {
        let value = end.map(|(row, column)| address(sheet, (row, column, row, column)));
        tag = set_attr(&tag, "table:end-cell-address", value.as_deref())?;
    }
    if let Some(name) = &change.name {
        let current = attribute(&xml, span, b"draw:name")?;
        if name.is_empty()
            || (current.as_deref() != Some(name.as_str())
                && drawing_names(&xml, &spans, sheet)?.contains(name))
        {
            return invalid("ODS drawing name is empty or already used on the sheet");
        }
        tag = set_attr(&tag, "draw:name", Some(name))?;
    }
    let mut edits = Vec::new();
    if tag != original {
        edits.push((span.start..span.tag_end, tag.into_bytes()));
    }
    if let Some(description) = &change.description {
        let existing = children(&spans, at, SVG, "desc").first().copied();
        let markup = description.as_ref().map(|text| {
            bind(
                format!("<svg:desc>{}</svg:desc>", escape_xml(text)),
                &xml,
                "svg",
                SVG,
            )
        });
        match (existing, markup) {
            (Some(existing), markup) => edits.push((
                spans[existing].start..spans[existing].end,
                markup.unwrap_or_default().into_bytes(),
            )),
            (None, Some(markup)) => {
                if is_self_closing_span(&xml, span)? {
                    return invalid("ODS drawing element is empty");
                }
                // After a title, else first.
                let after = children(&spans, at, SVG, "title")
                    .first()
                    .map_or(span.tag_end, |title| spans[*title].end);
                edits.push((after..after, markup.into_bytes()));
            },
            (None, None) => {},
        }
    }
    splice_part(source, "content.xml", edits, max_output)
}

/// Remove a drawing (and the shapes container with its last child) and the
/// graphic style this editor wrote for it; a picture returns its image path
/// when no other drawing references it.
pub(crate) fn remove_drawing(
    source: &[u8],
    sheet: &str,
    kind: DrawingKind,
    index: usize,
    max_output: usize,
) -> Result<(Vec<u8>, Option<String>)> {
    let (xml, spans) = content(source)?;
    let drawings = drawing_spans(&xml, &spans, sheet, kind)?;
    let at = *drawings
        .get(index)
        .ok_or_else(|| invalid_error("ODS drawing is missing"))?;
    let span = &spans[at];
    let image = children(&spans, at, DRAW, "image")
        .first()
        .and_then(|image| {
            attribute(&xml, &spans[*image], b"xlink:href")
                .ok()
                .flatten()
        });
    let style = attribute(&xml, span, b"draw:style-name")?
        .filter(|name| name.starts_with(LITCHI_SHAPE_STYLE));
    let parent = span
        .parent
        .ok_or_else(|| invalid_error("ODS drawing has no parent"))?;
    let siblings = spans
        .iter()
        .filter(|candidate| candidate.parent == Some(parent))
        .count();
    let range = if is_element(&spans[parent], TABLE, "shapes") && siblings == 1 {
        spans[parent].start..spans[parent].end
    } else {
        span.start..span.end
    };
    let mut edits = vec![(range, Vec::new())];
    if let Some(style) = style {
        let uses = xml
            .matches(&format!("draw:style-name=\"{}\"", escape_xml(&style)))
            .count();
        let (container, _) = automatic_styles(&xml, &spans);
        if let (1, Some(container)) = (uses, container) {
            let all = spans
                .iter()
                .filter(|candidate| candidate.parent == Some(container))
                .count();
            for child in children(&spans, container, STYLE, "style") {
                if attribute(&xml, &spans[child], b"style:name")?.as_deref() == Some(style.as_str())
                {
                    if all == 1 {
                        // The container held only this editor's style.
                        edits.push((spans[container].start..spans[container].end, Vec::new()));
                    } else {
                        edits.push((spans[child].start..spans[child].end, Vec::new()));
                    }
                }
            }
        }
    }
    let shared = image.as_ref().is_some_and(|image| {
        xml.matches(&format!("xlink:href=\"{}\"", escape_xml(image)))
            .count()
            > 1
    });
    Ok((
        splice_part(source, "content.xml", edits, max_output)?,
        image.filter(|_| !shared),
    ))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn formulas_convert_between_a1_and_openformula() {
        assert_eq!(to_odff("A1>5"), "[.A1]>5");
        assert_eq!(
            to_odff("SUM($B$2:C3)>LOG10(4)"),
            "SUM([.$B$2:.C3])>LOG10(4)"
        );
        assert_eq!(to_odff("Sheet2!A1+\"A1\""), "[Sheet2.A1]+\"A1\"");
        assert_eq!(from_odff("SUM([.E10:.E13])"), "SUM(E10:E13)");
        assert_eq!(from_odff("[Sheet2.A1]+1"), "Sheet2!A1+1");
        assert_eq!(from_odff(&to_odff("ISODD(A2)")), "ISODD(A2)");
    }

    #[test]
    fn addresses_round_trip() {
        assert_eq!(
            parse_address("Sheet1.A1:Sheet1.C9"),
            Some((Some("Sheet1".to_owned()), (0, 0, 8, 2)))
        );
        assert_eq!(
            parse_address("'My Sheet'.B2"),
            Some((Some("My Sheet".to_owned()), (1, 1, 1, 1)))
        );
        assert_eq!(
            address("My Sheet", (0, 0, 8, 2)),
            "'My Sheet'.A1:'My Sheet'.C9"
        );
        assert_eq!(
            rectangles(&[(0, 0, 1), (1, 0, 1), (3, 2, 2)]),
            vec![(0, 0, 1, 1), (3, 2, 3, 2)]
        );
    }
}
