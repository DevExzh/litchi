//! Lexical table-column and table-row edits of a slide graphic frame
//! (records 0321 and 0333).
//!
//! A column is one `a:gridCol` of the table's `a:tblGrid` and the cell at
//! its position in every `a:tr`; a row is one `a:tr`. The editor rewrites
//! only the bytes of the column or row it changes (the `w` attribute, an
//! inserted or removed `a:gridCol` and its row cells, an inserted or removed
//! `a:tr`), the merge attributes of the cells a merge moves to, and the
//! frame's `p:xfrm/a:ext` `cx` or `cy`, which it keeps equal to the sum of
//! the column widths or row heights (OfficeCLI's frame sync).
//!
//! Every row holds one `a:tc` per grid column, merged ones included
//! (`gridSpan` on a merge's first cell, `hMerge` on the cells it covers to
//! the right, `rowSpan` and `vMerge` likewise downwards), so a cell's grid
//! column is its position in its row. Merges stay consistent where that is
//! safe (record 0333):
//!
//! - a column is inserted at a boundary no horizontal merge crosses;
//! - removing a column shrinks the horizontal merges it is part of (the
//!   merge's next cell becomes its first when the first goes); cells that
//!   take part in a merge spanning both rows and columns are refused;
//! - a row is inserted at a boundary no vertical merge crosses;
//! - removing a row shrinks the vertical merges it is part of the same way.
//!
//! Tables whose rows do not hold one cell per grid column are refused for
//! every structural edit.

use std::ops::Range;

use litchi_core::xml::ReaderOrigin;
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use crate::{Error, Result};

const DML: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/main";
const STRICT_DML: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/main";
const PML: &[u8] = b"http://schemas.openxmlformats.org/presentationml/2006/main";
const STRICT_PML: &[u8] = b"http://purl.oclc.org/ooxml/presentationml/main";
const MAX_DEPTH: usize = 256;

/// The largest column width written, the `ST_PositiveCoordinate` bound.
pub(crate) const MAX_COLUMN_WIDTH: i64 = 27_273_042_316_900;

/// One column change.
#[derive(Debug, Clone, Copy)]
pub(crate) enum Change<'a> {
    /// Set the width of the column at a zero-based position.
    Width { column: usize, width: i64 },
    /// Insert a column before a zero-based position (the column count
    /// appends), with a width (`None`: the average of the existing columns,
    /// else one inch) and seed text for every new cell.
    Insert {
        position: usize,
        width: Option<i64>,
        text: &'a str,
    },
    /// Remove the column at a zero-based position.
    Remove { column: usize },
    /// Insert a row before a zero-based position (the row count appends),
    /// with a height (`None`: the first row's, else 370,840 EMU) and seed
    /// text for its cells, by column (missing columns are empty).
    InsertRow {
        position: usize,
        height: Option<i64>,
        texts: &'a [&'a str],
    },
    /// Remove the row at a zero-based position.
    RemoveRow { row: usize },
    /// Set the height of the row at a zero-based position.
    RowHeight { row: usize, height: i64 },
}

fn invalid(message: &str) -> Error {
    Error::Invalid(format!("table column edit: {message}"))
}

/// One `a:tc` of a row: its element span, start-tag span and merge
/// attributes.
#[derive(Clone)]
struct CellInfo {
    span: Range<usize>,
    tag: Range<usize>,
    grid_span: u32,
    row_span: u32,
    h_merge: bool,
    v_merge: bool,
}

impl CellInfo {
    /// Whether the cell is a horizontal merge's first or covered cell.
    const fn horizontal(&self) -> bool {
        self.grid_span > 1 || self.h_merge
    }

    /// Whether the cell is a vertical merge's first or covered cell.
    const fn vertical(&self) -> bool {
        self.row_span > 1 || self.v_merge
    }
}

/// One `a:tr`: its element span, the value span of its `h`, its cells and
/// the start of its end tag.
struct RowInfo {
    span: Range<usize>,
    height: Option<Range<usize>>,
    cells: Vec<CellInfo>,
    end: usize,
}

#[derive(Default)]
struct Grid {
    /// The `a:tblGrid` element: its start-tag span and, when it is not
    /// empty, the start of its end tag.
    grid: Option<(Range<usize>, Option<usize>)>,
    /// `(element span, w value span)` of each `a:gridCol`.
    columns: Vec<(Range<usize>, Option<Range<usize>>)>,
    /// The table's direct rows.
    rows: Vec<RowInfo>,
    /// The start of the `a:tbl` end tag, where an appended row goes.
    table_end: Option<usize>,
    /// The frame's `p:xfrm/a:ext` `cx` value span.
    extent: Option<Range<usize>>,
    /// The frame's `p:xfrm/a:ext` `cy` value span.
    extent_height: Option<Range<usize>>,
    /// The DrawingML prefix the table spells (`a`), empty for a default
    /// namespace.
    prefix: Option<Vec<u8>>,
}

impl Grid {
    /// Whether any cell spans or continues a horizontal merge.
    fn merged(&self) -> bool {
        self.rows
            .iter()
            .any(|row| row.cells.iter().any(CellInfo::horizontal))
    }

    /// Whether any cell spans or continues a vertical merge.
    fn vertically_merged(&self) -> bool {
        self.rows
            .iter()
            .any(|row| row.cells.iter().any(CellInfo::vertical))
    }

    /// Refuse a merged table whose rows do not hold one cell per column.
    fn require_regular(&self, operation: &'static str) -> Result<()> {
        if (self.merged() || self.vertically_merged())
            && self
                .rows
                .iter()
                .any(|row| row.cells.len() != self.columns.len())
        {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                operation,
                "a merged table's rows do not hold one cell per grid column",
            ));
        }
        Ok(())
    }
}

/// A non-negative integer attribute of a raw start tag; `default` when
/// absent.
fn number_attribute(tag: &[u8], name: &[u8], default: u32) -> Result<u32> {
    match attribute_value(tag, name) {
        None => Ok(default),
        Some(value) => std::str::from_utf8(&tag[value])
            .ok()
            .and_then(|value| value.trim().parse::<u32>().ok())
            .ok_or_else(|| invalid("a cell span is not a non-negative integer")),
    }
}

/// A boolean attribute of a raw start tag (`1`/`true`).
fn flag_attribute(tag: &[u8], name: &[u8]) -> bool {
    attribute_value(tag, name).is_some_and(|value| matches!(&tag[value], b"1" | b"true"))
}

/// Rewrite a raw start tag, setting (`Some`) or removing (`None`) unprefixed
/// attributes; every other attribute keeps its bytes, and set attributes
/// that were absent are appended.
fn set_attributes(tag: &[u8], changes: &[(&[u8], Option<String>)]) -> Vec<u8> {
    let mut out = tag.to_vec();
    // Apply from the last existing attribute to the first so earlier spans
    // stay valid.
    let mut present: Vec<(Range<usize>, Option<&String>)> = Vec::new();
    let mut appended: Vec<(&[u8], &String)> = Vec::new();
    for (name, value) in changes {
        match attribute_span(tag, name) {
            Some(span) => present.push((span, value.as_ref())),
            None => {
                if let Some(value) = value {
                    appended.push((name, value));
                }
            },
        }
    }
    present.sort_by_key(|(span, _)| std::cmp::Reverse(span.start));
    for (span, value) in present {
        match value {
            Some(value) => {
                let name_end = tag[span.clone()]
                    .iter()
                    .position(|byte| *byte == b'=')
                    .unwrap_or(0);
                let name = tag[span.start..span.start + name_end].to_vec();
                let mut replacement = name;
                replacement.extend_from_slice(b"=\"");
                replacement.extend_from_slice(value.as_bytes());
                replacement.push(b'"');
                out.splice(span, replacement);
            },
            None => {
                // Drop the attribute with the whitespace before it.
                let mut start = span.start;
                while start > 0 && tag[start - 1].is_ascii_whitespace() {
                    start -= 1;
                }
                out.splice(start..span.end, std::iter::empty());
            },
        }
    }
    if !appended.is_empty() {
        let close = if out.ends_with(b"/>") {
            out.len() - 2
        } else {
            out.len() - 1
        };
        let mut extra = Vec::new();
        for (name, value) in appended {
            extra.push(b' ');
            extra.extend_from_slice(name);
            extra.extend_from_slice(b"=\"");
            extra.extend_from_slice(value.as_bytes());
            extra.push(b'"');
        }
        out.splice(close..close, extra);
    }
    out
}

/// The span of a whole attribute (`name="value"`) of a raw start tag.
fn attribute_span(tag: &[u8], name: &[u8]) -> Option<Range<usize>> {
    let mut at = 1;
    while at < tag.len() && !tag[at].is_ascii_whitespace() && tag[at] != b'>' && tag[at] != b'/' {
        at += 1;
    }
    loop {
        while at < tag.len() && tag[at].is_ascii_whitespace() {
            at += 1;
        }
        let start = at;
        while at < tag.len() && tag[at] != b'=' && !tag[at].is_ascii_whitespace() {
            if tag[at] == b'>' || tag[at] == b'/' {
                return None;
            }
            at += 1;
        }
        let key = &tag[start..at];
        while at < tag.len() && tag[at].is_ascii_whitespace() {
            at += 1;
        }
        if tag.get(at) != Some(&b'=') {
            return None;
        }
        at += 1;
        while at < tag.len() && tag[at].is_ascii_whitespace() {
            at += 1;
        }
        let quote = *tag.get(at)?;
        if quote != b'"' && quote != b'\'' {
            return None;
        }
        at += 1;
        while at < tag.len() && tag[at] != quote {
            at += 1;
        }
        if at >= tag.len() {
            return None;
        }
        at += 1;
        if key == name {
            return Some(start..at);
        }
    }
}

fn is_ns(namespace: &ResolveResult<'_>, wanted: [&[u8]; 2]) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(value)) if wanted.contains(value))
}

/// The value span of attribute `name` in one raw start tag, relative to the
/// tag.
pub(crate) fn attribute_value(tag: &[u8], name: &[u8]) -> Option<Range<usize>> {
    let mut at = 1;
    while at < tag.len() && !tag[at].is_ascii_whitespace() && tag[at] != b'>' && tag[at] != b'/' {
        at += 1;
    }
    loop {
        while at < tag.len() && tag[at].is_ascii_whitespace() {
            at += 1;
        }
        let start = at;
        while at < tag.len() && tag[at] != b'=' && !tag[at].is_ascii_whitespace() {
            if tag[at] == b'>' || tag[at] == b'/' {
                return None;
            }
            at += 1;
        }
        let key = &tag[start..at];
        while at < tag.len() && tag[at].is_ascii_whitespace() {
            at += 1;
        }
        if tag.get(at) != Some(&b'=') {
            return None;
        }
        at += 1;
        while at < tag.len() && tag[at].is_ascii_whitespace() {
            at += 1;
        }
        let quote = *tag.get(at)?;
        if quote != b'"' && quote != b'\'' {
            return None;
        }
        at += 1;
        let value = at;
        while at < tag.len() && tag[at] != quote {
            at += 1;
        }
        if at >= tag.len() {
            return None;
        }
        if key == name {
            return Some(value..at);
        }
        at += 1;
    }
}

fn read_grid(xml: &[u8], frame: Range<usize>) -> Result<Grid> {
    let origin = ReaderOrigin::of(xml);
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().check_end_names = true;
    // The open elements inside the frame: (local name, start of element,
    // is DrawingML, is PresentationML).
    let mut stack: Vec<(Vec<u8>, usize, bool, bool)> = Vec::new();
    let mut grid = Grid::default();
    loop {
        let before = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("offset overflow"))?;
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let dml = is_ns(&namespace, [DML, STRICT_DML]);
        let pml = is_ns(&namespace, [PML, STRICT_PML]);
        let after = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("offset overflow"))?;
        if matches!(event, Event::Eof) {
            break;
        }
        if before < frame.start {
            continue;
        }
        if before >= frame.end {
            break;
        }
        match &event {
            Event::Start(element) | Event::Empty(element) => {
                let local = element.local_name().as_ref().to_vec();
                let tag = before..after;
                let parent = stack
                    .last()
                    .map(|(name, _, dml, pml)| (name.as_slice(), *dml, *pml));
                let depth = stack.len();
                if dml && grid.prefix.is_none() && (local == b"tbl" || local == b"tblGrid") {
                    grid.prefix = Some(
                        element
                            .name()
                            .prefix()
                            .map(|prefix| prefix.into_inner().to_vec())
                            .unwrap_or_default(),
                    );
                }
                match (local.as_slice(), parent) {
                    (b"ext", Some((b"xfrm", _, true))) if dml && depth == 2 => {
                        grid.extent = attribute_value(&xml[tag.clone()], b"cx")
                            .map(|value| value.start + tag.start..value.end + tag.start);
                        grid.extent_height = attribute_value(&xml[tag.clone()], b"cy")
                            .map(|value| value.start + tag.start..value.end + tag.start);
                    },
                    (b"tblGrid", Some((b"tbl", true, _))) if dml => {
                        grid.grid = Some((tag.clone(), None));
                    },
                    (b"gridCol", Some((b"tblGrid", true, _))) if dml => {
                        let width = attribute_value(&xml[tag.clone()], b"w")
                            .map(|value| value.start + tag.start..value.end + tag.start);
                        grid.columns.push((tag.start..tag.end, width));
                    },
                    (b"tr", Some((b"tbl", true, _))) if dml => {
                        let height = attribute_value(&xml[tag.clone()], b"h")
                            .map(|value| value.start + tag.start..value.end + tag.start);
                        let empty = matches!(event, Event::Empty(_));
                        grid.rows.push(RowInfo {
                            span: tag.clone(),
                            height,
                            cells: Vec::new(),
                            // An empty row has no end tag to insert before.
                            end: if empty { usize::MAX } else { 0 },
                        });
                    },
                    (b"tc", Some((b"tr", true, _))) if dml => {
                        let raw = &xml[tag.clone()];
                        let cell = CellInfo {
                            span: tag.clone(),
                            tag: tag.clone(),
                            grid_span: number_attribute(raw, b"gridSpan", 1)?,
                            row_span: number_attribute(raw, b"rowSpan", 1)?,
                            h_merge: flag_attribute(raw, b"hMerge"),
                            v_merge: flag_attribute(raw, b"vMerge"),
                        };
                        if let Some(row) = grid.rows.last_mut() {
                            row.cells.push(cell);
                        }
                    },
                    _ => {},
                }
                if matches!(event, Event::Start(_)) {
                    if stack.len() >= MAX_DEPTH {
                        return Err(invalid("the table is nested too deeply"));
                    }
                    stack.push((local, tag.start, dml, pml));
                }
            },
            Event::End(_) => {
                let Some((local, start, dml, _)) = stack.pop() else {
                    break;
                };
                let parent_dml = stack.last().is_some_and(|(name, _, dml, _)| {
                    *dml && (name.as_slice() == b"tblGrid" || name.as_slice() == b"tr")
                });
                if dml && local == b"gridCol" && parent_dml {
                    if let Some(column) = grid.columns.last_mut()
                        && column.0.start == start
                    {
                        column.0.end = after;
                    }
                } else if dml && local == b"tc" && parent_dml {
                    if let Some(cell) = grid.rows.last_mut().and_then(|row| row.cells.last_mut())
                        && cell.span.start == start
                    {
                        cell.span.end = after;
                    }
                } else if dml && local == b"tr" {
                    if let Some(row) = grid.rows.last_mut()
                        && row.span.start == start
                    {
                        row.end = before;
                        row.span.end = after;
                    }
                } else if dml && local == b"tbl" {
                    grid.table_end = Some(before);
                } else if dml && local == b"tblGrid" {
                    if let Some((_, end)) = grid.grid.as_mut() {
                        *end = Some(before);
                    }
                }
                if stack.is_empty() {
                    break;
                }
            },
            _ => {},
        }
    }
    Ok(grid)
}

/// A span attribute's value: absent at one.
fn span_value(span: u32) -> Option<String> {
    (span > 1).then(|| span.to_string())
}

/// Every row's height (`a:tr@h`, zero when absent).
fn row_heights(xml: &[u8], grid: &Grid) -> Result<Vec<i64>> {
    grid.rows
        .iter()
        .map(|row| match &row.height {
            Some(range) => std::str::from_utf8(&xml[range.clone()])
                .ok()
                .and_then(|value| value.trim().parse::<i64>().ok())
                .filter(|value| *value >= 0)
                .ok_or_else(|| invalid("a row height is not a non-negative integer")),
            None => Ok(0),
        })
        .collect()
}

fn parse_width(xml: &[u8], range: &Option<Range<usize>>) -> Result<i64> {
    let range = range
        .clone()
        .ok_or_else(|| invalid("a grid column has no width"))?;
    std::str::from_utf8(&xml[range])
        .ok()
        .and_then(|value| value.trim().parse::<i64>().ok())
        .filter(|value| *value >= 0)
        .ok_or_else(|| invalid("a grid column width is not a non-negative integer"))
}

fn escape(text: &str) -> String {
    let mut out = String::with_capacity(text.len());
    for ch in text.chars() {
        match ch {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            _ => out.push(ch),
        }
    }
    out
}

fn new_cell(prefix: &str, text: &str) -> String {
    let p = if prefix.is_empty() {
        String::new()
    } else {
        format!("{prefix}:")
    };
    let paragraph = if text.is_empty() {
        format!("<{p}endParaRPr lang=\"en-US\"/>")
    } else {
        format!(
            "<{p}r><{p}rPr lang=\"en-US\"/><{p}t>{}</{p}t></{p}r>",
            escape(text)
        )
    };
    format!(
        "<{p}tc><{p}txBody><{p}bodyPr/><{p}lstStyle/><{p}p>{paragraph}</{p}p></{p}txBody><{p}tcPr/></{p}tc>"
    )
}

/// Apply `change` to the table of the graphic frame at `frame` in `xml`.
/// `Ok(None)` is an exact no-op (an equal width).
pub(crate) fn change(
    xml: &[u8],
    frame: Range<usize>,
    change: Change<'_>,
) -> Result<Option<Vec<u8>>> {
    let grid = read_grid(xml, frame)?;
    if grid.grid.is_none() {
        return Err(invalid("the graphic frame has no table grid"));
    }
    let widths = grid
        .columns
        .iter()
        .map(|(_, width)| parse_width(xml, width))
        .collect::<Result<Vec<_>>>()?;
    let prefix = String::from_utf8(grid.prefix.clone().unwrap_or_default())
        .map_err(|_| invalid("the DrawingML prefix is not UTF-8"))?;
    let qualified = |local: &str| {
        if prefix.is_empty() {
            local.to_owned()
        } else {
            format!("{prefix}:{local}")
        }
    };
    // (range, replacement) edits, applied in ascending order.
    let mut edits: Vec<(Range<usize>, Vec<u8>)> = Vec::new();
    let mut next_widths = widths.clone();
    // The row heights after the edit, when rows change.
    let mut next_heights: Option<Vec<i64>> = None;
    match change {
        Change::Width { column, width } => {
            if !(0..=MAX_COLUMN_WIDTH).contains(&width) {
                return Err(invalid("the column width is out of range"));
            }
            let Some((_, Some(range))) = grid.columns.get(column) else {
                return Err(invalid("the column position is past the last column"));
            };
            if widths[column] == width {
                return Ok(None);
            }
            edits.push((range.clone(), width.to_string().into_bytes()));
            next_widths[column] = width;
        },
        Change::Insert {
            position,
            width,
            text,
        } => {
            grid.require_regular("insert_table_column")?;
            if position > widths.len() {
                return Err(invalid("the column position is past the column count"));
            }
            // A merge covering the column at the position started left of
            // it: the new column would cut it.
            if grid
                .rows
                .iter()
                .any(|row| row.cells.get(position).is_some_and(|cell| cell.h_merge))
            {
                return Err(Error::refused(
                    litchi_core::edit::RefusalKind::UnsupportedTarget,
                    "insert_table_column",
                    "the position cuts through cells merged across columns",
                ));
            }
            let width = match width {
                Some(width) => width,
                None if widths.is_empty() => 914_400,
                None => {
                    let total: i128 = widths.iter().map(|width| i128::from(*width)).sum();
                    i64::try_from(total / widths.len() as i128).unwrap_or(914_400)
                },
            };
            if !(0..=MAX_COLUMN_WIDTH).contains(&width) {
                return Err(invalid("the column width is out of range"));
            }
            let column = format!("<{} w=\"{width}\"/>", qualified("gridCol"));
            let (grid_tag, grid_end) = grid.grid.clone().unwrap_or_default();
            match (grid.columns.get(position), grid_end) {
                (Some((span, _)), _) => edits.push((span.start..span.start, column.into_bytes())),
                (None, Some(end)) => edits.push((end..end, column.into_bytes())),
                (None, None) => {
                    // `<a:tblGrid/>`: open it.
                    let name = qualified("tblGrid");
                    edits.push((grid_tag, format!("<{name}>{column}</{name}>").into_bytes()));
                },
            }
            let cell = new_cell(&prefix, text);
            for row in &grid.rows {
                if row.end == usize::MAX {
                    return Err(invalid("an empty row cannot take a cell"));
                }
                let at = row
                    .cells
                    .get(position)
                    .map_or(row.end, |cell| cell.span.start);
                if row.cells.len() < position {
                    return Err(invalid("a row has fewer cells than the column position"));
                }
                edits.push((at..at, cell.clone().into_bytes()));
            }
            next_widths.insert(position, width);
        },
        Change::Remove { column } => {
            grid.require_regular("remove_table_column")?;
            let Some((span, _)) = grid.columns.get(column) else {
                return Err(invalid("the column position is past the last column"));
            };
            if grid.columns.len() == 1 {
                return Err(Error::refused(
                    litchi_core::edit::RefusalKind::UnsupportedOperation,
                    "remove_table_column",
                    "a table keeps at least one column",
                ));
            }
            edits.push((span.clone(), Vec::new()));
            for row in &grid.rows {
                let cell = row
                    .cells
                    .get(column)
                    .ok_or_else(|| invalid("a row has no cell at the column position"))?;
                if cell.horizontal() && cell.vertical() {
                    return Err(Error::refused(
                        litchi_core::edit::RefusalKind::UnsupportedTarget,
                        "remove_table_column",
                        "the column takes part in a merge spanning rows and columns",
                    ));
                }
                if cell.h_merge {
                    // A covered cell: its merge's first cell spans one
                    // column less.
                    let (anchor, first) = row.cells[..column]
                        .iter()
                        .enumerate()
                        .rev()
                        .find(|(_, cell)| !cell.h_merge)
                        .ok_or_else(|| invalid("a covered cell has no merge to its left"))?;
                    if first.vertical() {
                        return Err(Error::refused(
                            litchi_core::edit::RefusalKind::UnsupportedTarget,
                            "remove_table_column",
                            "the column takes part in a merge spanning rows and columns",
                        ));
                    }
                    if first.grid_span as usize + anchor > column {
                        edits.push((
                            first.tag.clone(),
                            set_attributes(
                                &xml[first.tag.clone()],
                                &[(b"gridSpan", span_value(first.grid_span - 1))],
                            ),
                        ));
                    }
                } else if cell.grid_span > 1 {
                    // A merge's first cell: the next cell becomes it.
                    let next = row
                        .cells
                        .get(column + 1)
                        .filter(|next| next.h_merge && !next.vertical())
                        .ok_or(Error::refused(
                            litchi_core::edit::RefusalKind::UnsupportedTarget,
                            "remove_table_column",
                            "a merged cell's covered cells are not in its row",
                        ))?;
                    edits.push((
                        next.tag.clone(),
                        set_attributes(
                            &xml[next.tag.clone()],
                            &[
                                (b"hMerge", None),
                                (b"gridSpan", span_value(cell.grid_span - 1)),
                            ],
                        ),
                    ));
                }
                edits.push((cell.span.clone(), Vec::new()));
            }
            next_widths.remove(column);
        },
        Change::InsertRow {
            position,
            height,
            texts,
        } => {
            grid.require_regular("insert_table_row")?;
            if position > grid.rows.len() {
                return Err(invalid("the row position is past the row count"));
            }
            if widths.is_empty() {
                return Err(invalid("a table without grid columns cannot take a row"));
            }
            if texts.len() > widths.len() {
                return Err(invalid("the row has more seed texts than grid columns"));
            }
            // A vertical merge covering the row at the position started
            // above it: the new row would cut it.
            if grid
                .rows
                .get(position)
                .is_some_and(|row| row.cells.iter().any(|cell| cell.v_merge))
            {
                return Err(Error::refused(
                    litchi_core::edit::RefusalKind::UnsupportedTarget,
                    "insert_table_row",
                    "the position cuts through cells merged across rows",
                ));
            }
            let mut heights = row_heights(xml, &grid)?;
            let height = match height {
                Some(height) => height,
                None => heights.first().copied().unwrap_or(370_840),
            };
            if !(0..=MAX_COLUMN_WIDTH).contains(&height) {
                return Err(invalid("the row height is out of range"));
            }
            let mut row = format!("<{} h=\"{height}\">", qualified("tr"));
            for column in 0..widths.len() {
                row.push_str(&new_cell(&prefix, texts.get(column).copied().unwrap_or("")));
            }
            row.push_str(&format!("</{}>", qualified("tr")));
            let at = match (grid.rows.get(position), grid.rows.last(), grid.table_end) {
                (Some(next), _, _) => next.span.start,
                (None, Some(last), _) => last.span.end,
                (None, None, Some(end)) => end,
                (None, None, None) => return Err(invalid("the table has no end tag")),
            };
            edits.push((at..at, row.into_bytes()));
            heights.insert(position, height);
            next_heights = Some(heights);
        },
        Change::RowHeight { row, height } => {
            if !(0..=MAX_COLUMN_WIDTH).contains(&height) {
                return Err(invalid("the row height is out of range"));
            }
            let Some(target) = grid.rows.get(row) else {
                return Err(invalid("the row position is past the last row"));
            };
            let mut heights = row_heights(xml, &grid)?;
            if heights[row] == height && target.height.is_some() {
                return Ok(None);
            }
            match &target.height {
                Some(range) => edits.push((range.clone(), height.to_string().into_bytes())),
                None => {
                    let tag_end = xml[target.span.start..]
                        .iter()
                        .position(|byte| *byte == b'>')
                        .map(|offset| target.span.start + offset)
                        .ok_or_else(|| invalid("a row start tag is unterminated"))?;
                    let at = if xml[..tag_end].ends_with(b"/") {
                        tag_end - 1
                    } else {
                        tag_end
                    };
                    edits.push((at..at, format!(" h=\"{height}\"").into_bytes()));
                },
            }
            heights[row] = height;
            next_heights = Some(heights);
        },
        Change::RemoveRow { row } => {
            grid.require_regular("remove_table_row")?;
            let Some(removed) = grid.rows.get(row) else {
                return Err(invalid("the row position is past the last row"));
            };
            if grid.rows.len() == 1 {
                return Err(Error::refused(
                    litchi_core::edit::RefusalKind::UnsupportedOperation,
                    "remove_table_row",
                    "a table keeps at least one row",
                ));
            }
            let mut heights = row_heights(xml, &grid)?;
            for (column, cell) in removed.cells.iter().enumerate() {
                if cell.v_merge && !cell.h_merge {
                    // A covered cell: its merge's first cell spans one row
                    // less.
                    let first = grid.rows[..row]
                        .iter()
                        .rev()
                        .filter_map(|above| above.cells.get(column))
                        .find(|above| !above.v_merge)
                        .ok_or_else(|| invalid("a covered cell has no merge above it"))?;
                    if first.row_span > 1 {
                        edits.push((
                            first.tag.clone(),
                            set_attributes(
                                &xml[first.tag.clone()],
                                &[(b"rowSpan", span_value(first.row_span - 1))],
                            ),
                        ));
                    }
                } else if cell.row_span > 1 && !cell.v_merge {
                    // A merge's first cell: the cells below become it.
                    let below = grid
                        .rows
                        .get(row + 1)
                        .ok_or_else(|| invalid("a merged cell's covered cells are missing"))?;
                    let columns = (cell.grid_span.max(1)) as usize;
                    for offset in 0..columns {
                        let next = below
                            .cells
                            .get(column + offset)
                            .filter(|next| next.v_merge)
                            .ok_or(Error::refused(
                                litchi_core::edit::RefusalKind::UnsupportedTarget,
                                "remove_table_row",
                                "a merged cell's covered cells are not below it",
                            ))?;
                        let mut changes: Vec<(&[u8], Option<String>)> = vec![(b"vMerge", None)];
                        if offset == 0 {
                            changes.push((b"rowSpan", span_value(cell.row_span - 1)));
                            changes.push((b"gridSpan", span_value(cell.grid_span)));
                        }
                        edits.push((
                            next.tag.clone(),
                            set_attributes(&xml[next.tag.clone()], &changes),
                        ));
                    }
                }
            }
            edits.push((removed.span.clone(), Vec::new()));
            heights.remove(row);
            next_heights = Some(heights);
        },
    }
    if let Some(heights) = &next_heights {
        if let Some(extent) = &grid.extent_height {
            let total: i128 = heights.iter().map(|height| i128::from(*height)).sum();
            let total = i64::try_from(total)
                .ok()
                .filter(|total| *total <= MAX_COLUMN_WIDTH)
                .ok_or_else(|| invalid("the table height is out of range"))?;
            let current = std::str::from_utf8(&xml[extent.clone()])
                .ok()
                .and_then(|value| value.trim().parse::<i64>().ok());
            if current != Some(total) {
                edits.push((extent.clone(), total.to_string().into_bytes()));
            }
        }
    } else if let Some(extent) = &grid.extent {
        let total: i128 = next_widths.iter().map(|width| i128::from(*width)).sum();
        let total = i64::try_from(total)
            .ok()
            .filter(|total| *total <= MAX_COLUMN_WIDTH)
            .ok_or_else(|| invalid("the table width is out of range"))?;
        let current = std::str::from_utf8(&xml[extent.clone()])
            .ok()
            .and_then(|value| value.trim().parse::<i64>().ok());
        if current != Some(total) {
            edits.push((extent.clone(), total.to_string().into_bytes()));
        }
    }
    edits.sort_by_key(|(range, _)| (range.start, range.end));
    if edits.windows(2).any(|pair| pair[0].0.end > pair[1].0.start) {
        return Err(invalid("table edits overlap"));
    }
    let mut out = Vec::with_capacity(xml.len() + 256);
    let mut cursor = 0;
    for (range, replacement) in edits {
        out.extend_from_slice(&xml[cursor..range.start]);
        out.extend_from_slice(&replacement);
        cursor = range.end;
    }
    out.extend_from_slice(&xml[cursor..]);
    Ok(Some(out))
}

/// The widths of the grid columns of the table element at `table` in `xml`,
/// in EMU.
pub(crate) fn widths(xml: &[u8], table: Range<usize>) -> Result<Vec<i64>> {
    let grid = read_grid(xml, table)?;
    grid.columns
        .iter()
        .map(|(_, width)| parse_width(xml, width))
        .collect()
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::*;

    const FRAME: &str = concat!(
        r#"<p:graphicFrame xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" "#,
        r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">"#,
        r#"<p:xfrm><a:off x="0" y="0"/><a:ext cx="200" cy="50"/></p:xfrm>"#,
        r#"<a:graphic><a:graphicData><a:tbl><a:tblPr/><a:tblGrid>"#,
        r#"<a:gridCol w="100"/><a:gridCol w="100"/></a:tblGrid>"#,
        r#"<a:tr h="50"><a:tc><a:txBody><a:bodyPr/><a:p/></a:txBody></a:tc>"#,
        r#"<a:tc><a:txBody><a:bodyPr/><a:p/></a:txBody></a:tc></a:tr>"#,
        r#"</a:tbl></a:graphicData></a:graphic></p:graphicFrame>"#
    );

    #[test]
    fn width_insert_and_remove_keep_the_frame_in_step() {
        let xml = FRAME.as_bytes();
        let all = 0..xml.len();
        assert!(
            change(
                xml,
                all.clone(),
                Change::Width {
                    column: 0,
                    width: 100
                }
            )
            .unwrap()
            .is_none()
        );
        let wide = change(
            xml,
            all.clone(),
            Change::Width {
                column: 1,
                width: 300,
            },
        )
        .unwrap()
        .unwrap();
        let wide = String::from_utf8(wide).unwrap();
        assert!(wide.contains(r#"<a:gridCol w="300"/>"#));
        assert!(wide.contains(r#"<a:ext cx="400" cy="50"/>"#));
        let inserted = change(
            xml,
            all.clone(),
            Change::Insert {
                position: 1,
                width: None,
                text: "x & y",
            },
        )
        .unwrap()
        .unwrap();
        let inserted_text = String::from_utf8(inserted.clone()).unwrap();
        assert_eq!(inserted_text.matches("<a:tc>").count(), 3);
        assert!(inserted_text.contains("x &amp; y"));
        assert!(inserted_text.contains(r#"cx="300""#));
        assert_eq!(
            widths(&inserted, 0..inserted.len()).unwrap(),
            vec![100, 100, 100]
        );
        let removed = change(&inserted, 0..inserted.len(), Change::Remove { column: 1 })
            .unwrap()
            .unwrap();
        assert_eq!(removed, xml);
    }

    #[test]
    fn merged_tables_and_last_columns_are_refused() {
        let merged = FRAME.replacen("<a:tc>", r#"<a:tc gridSpan="2">"#, 1);
        let bytes = merged.as_bytes();
        assert!(change(bytes, 0..bytes.len(), Change::Remove { column: 0 }).is_err());
        let xml = FRAME.as_bytes();
        let once = change(xml, 0..xml.len(), Change::Remove { column: 0 })
            .unwrap()
            .unwrap();
        assert!(change(&once, 0..once.len(), Change::Remove { column: 0 }).is_err());
        assert!(change(xml, 0..xml.len(), Change::Remove { column: 2 }).is_err());
    }

    /// A three-column table whose first row merges its first two cells and
    /// whose first column merges its first two rows.
    fn merged() -> String {
        concat!(
            r#"<p:graphicFrame xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" "#,
            r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">"#,
            r#"<p:xfrm><a:off x="0" y="0"/><a:ext cx="300" cy="150"/></p:xfrm>"#,
            r#"<a:graphic><a:graphicData><a:tbl><a:tblPr/><a:tblGrid>"#,
            r#"<a:gridCol w="100"/><a:gridCol w="100"/><a:gridCol w="100"/></a:tblGrid>"#,
            r#"<a:tr h="50"><a:tc rowSpan="2"><a:txBody><a:bodyPr/><a:p/></a:txBody></a:tc>"#,
            r#"<a:tc gridSpan="2"><a:txBody><a:bodyPr/><a:p/></a:txBody></a:tc>"#,
            r#"<a:tc hMerge="1"><a:txBody><a:bodyPr/><a:p/></a:txBody></a:tc></a:tr>"#,
            r#"<a:tr h="50"><a:tc vMerge="1"><a:txBody><a:bodyPr/><a:p/></a:txBody></a:tc>"#,
            r#"<a:tc><a:txBody><a:bodyPr/><a:p/></a:txBody></a:tc>"#,
            r#"<a:tc><a:txBody><a:bodyPr/><a:p/></a:txBody></a:tc></a:tr>"#,
            r#"<a:tr h="50"><a:tc><a:txBody><a:bodyPr/><a:p/></a:txBody></a:tc>"#,
            r#"<a:tc><a:txBody><a:bodyPr/><a:p/></a:txBody></a:tc>"#,
            r#"<a:tc><a:txBody><a:bodyPr/><a:p/></a:txBody></a:tc></a:tr>"#,
            r#"</a:tbl></a:graphicData></a:graphic></p:graphicFrame>"#
        )
        .to_owned()
    }

    fn apply(xml: &str, edit: Change<'_>) -> Result<String> {
        let bytes = xml.as_bytes();
        Ok(String::from_utf8(change(bytes, 0..bytes.len(), edit)?.expect("changed")).unwrap())
    }

    #[test]
    fn column_removal_shrinks_horizontal_merges() {
        let xml = merged();
        // The covered cell goes: the merge spans one column.
        let covered = apply(&xml, Change::Remove { column: 2 }).unwrap();
        assert!(covered.contains(r#"<a:tc><a:txBody>"#));
        assert!(!covered.contains("gridSpan") && !covered.contains("hMerge"));
        assert!(covered.contains(r#"cx="200""#));
        // The first cell goes: the covered cell becomes a plain cell.
        let first = apply(&xml, Change::Remove { column: 1 }).unwrap();
        assert!(!first.contains("gridSpan") && !first.contains("hMerge"));
        assert_eq!(first.matches("</a:tc>").count(), 6);
        // A column inside a vertical merge goes with the whole merge.
        let vertical = apply(&xml, Change::Remove { column: 0 }).unwrap();
        assert!(!vertical.contains("rowSpan") && !vertical.contains("vMerge"));
    }

    #[test]
    fn column_insertion_keeps_out_of_merges() {
        let xml = merged();
        assert!(matches!(
            apply(
                &xml,
                Change::Insert {
                    position: 2,
                    width: None,
                    text: ""
                }
            ),
            Err(Error::Refused(_))
        ));
        let inserted = apply(
            &xml,
            Change::Insert {
                position: 1,
                width: Some(50),
                text: "",
            },
        )
        .unwrap();
        assert_eq!(inserted.matches("</a:tc>").count(), 12);
        assert!(inserted.contains(r#"gridSpan="2""#) && inserted.contains(r#"cx="350""#));
    }

    #[test]
    fn rows_are_inserted_and_removed_with_vertical_merges_kept() {
        let xml = merged();
        let appended = apply(
            &xml,
            Change::InsertRow {
                position: 3,
                height: None,
                texts: &["x & y"],
            },
        )
        .unwrap();
        assert_eq!(appended.matches("<a:tr ").count(), 4);
        assert!(appended.contains("x &amp; y"));
        assert!(appended.contains(r#"cy="200""#));
        let tall = apply(
            &xml,
            Change::InsertRow {
                position: 0,
                height: Some(80),
                texts: &[],
            },
        )
        .unwrap();
        assert!(tall.find(r#"<a:tr h="80">"#) < tall.find(r#"<a:tr h="50">"#));
        assert!(tall.contains(r#"<a:tr h="80">"#) && tall.contains(r#"cy="230""#));
        // A new row cannot cut the vertical merge of the first column.
        assert!(matches!(
            apply(
                &xml,
                Change::InsertRow {
                    position: 1,
                    height: None,
                    texts: &[]
                }
            ),
            Err(Error::Refused(_))
        ));
        // The covered row goes: the merge spans one row.
        let covered = apply(&xml, Change::RemoveRow { row: 1 }).unwrap();
        assert!(!covered.contains("rowSpan") && !covered.contains("vMerge"));
        assert!(covered.contains(r#"cy="100""#));
        // The first row goes: the covered cell becomes a plain cell.
        let first = apply(&xml, Change::RemoveRow { row: 0 }).unwrap();
        assert!(!first.contains("rowSpan") && !first.contains("vMerge"));
        assert_eq!(first.matches("<a:tr ").count(), 2);
        // The last row is kept.
        let one = FRAME.as_bytes();
        assert!(matches!(
            change(one, 0..one.len(), Change::RemoveRow { row: 0 }),
            Err(Error::Refused(_))
        ));
    }

    #[test]
    fn longer_vertical_merges_shrink_by_one() {
        let xml = merged()
            .replacen(r#"rowSpan="2""#, r#"rowSpan="3""#, 1)
            .replacen(
                r#"<a:tr h="50"><a:tc><a:txBody>"#,
                r#"<a:tr h="50"><a:tc vMerge="1"><a:txBody>"#,
                1,
            );
        let first = apply(&xml, Change::RemoveRow { row: 0 }).unwrap();
        assert!(first.contains(r#"<a:tc rowSpan="2">"#), "{first}");
        assert_eq!(first.matches("vMerge").count(), 1);
        let middle = apply(&xml, Change::RemoveRow { row: 2 }).unwrap();
        assert!(middle.contains(r#"<a:tc rowSpan="2">"#));
    }

    #[test]
    fn row_heights_are_set_and_the_frame_follows() {
        let xml = FRAME.as_bytes();
        let all = 0..xml.len();
        assert!(
            change(xml, all.clone(), Change::RowHeight { row: 0, height: 50 })
                .unwrap()
                .is_none()
        );
        let tall = apply(FRAME, Change::RowHeight { row: 0, height: 80 }).unwrap();
        assert!(tall.contains(r#"<a:tr h="80">"#) && tall.contains(r#"cy="80""#));
        let bare = FRAME.replacen(r#"<a:tr h="50">"#, "<a:tr>", 1);
        let set = apply(&bare, Change::RowHeight { row: 0, height: 60 }).unwrap();
        assert!(set.contains(r#"<a:tr h="60">"#), "{set}");
        assert!(apply(FRAME, Change::RowHeight { row: 3, height: 1 }).is_err());
    }
}
