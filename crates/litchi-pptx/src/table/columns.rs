//! Lexical table-column edits of a slide graphic frame (record 0321).
//!
//! A column is one `a:gridCol` of the table's `a:tblGrid` and the cell at
//! its position in every `a:tr`. The editor rewrites only the bytes of the
//! column it changes (the `w` attribute, an inserted or removed `a:gridCol`
//! and its row cells) and the frame's `p:xfrm/a:ext/@cx`, which it keeps
//! equal to the sum of the column widths (OfficeCLI's frame sync). Tables
//! with merged cells across columns (`gridSpan`, `hMerge`) are refused for
//! insertion and removal: moving a span boundary is not a column edit.

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
}

fn invalid(message: &str) -> Error {
    Error::Invalid(format!("table column edit: {message}"))
}

#[derive(Default)]
struct Grid {
    /// The `a:tblGrid` element: its start-tag span and, when it is not
    /// empty, the start of its end tag.
    grid: Option<(Range<usize>, Option<usize>)>,
    /// `(element span, w value span)` of each `a:gridCol`.
    columns: Vec<(Range<usize>, Option<Range<usize>>)>,
    /// Each row's direct cells: the element spans, and the start of the
    /// row's end tag.
    rows: Vec<(Vec<Range<usize>>, usize)>,
    /// Whether any cell spans or continues a horizontal merge.
    merged: bool,
    /// The frame's `p:xfrm/a:ext` start-tag span and its `cx` value span.
    extent: Option<Range<usize>>,
    /// The DrawingML prefix the table spells (`a`), empty for a default
    /// namespace.
    prefix: Option<Vec<u8>>,
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
                        grid.rows.push((Vec::new(), 0));
                    },
                    (b"tc", Some((b"tr", true, _))) if dml => {
                        let raw = &xml[tag.clone()];
                        if attribute_value(raw, b"hMerge").is_some()
                            || attribute_value(raw, b"gridSpan")
                                .is_some_and(|value| &raw[value] != b"1")
                        {
                            grid.merged = true;
                        }
                        if let Some((cells, _)) = grid.rows.last_mut() {
                            cells.push(tag.start..tag.end);
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
                    if let Some(cell) = grid.rows.last_mut().and_then(|(cells, _)| cells.last_mut())
                        && cell.start == start
                    {
                        cell.end = after;
                    }
                } else if dml && local == b"tr" {
                    if let Some(row) = grid.rows.last_mut() {
                        row.1 = before;
                    }
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
            if grid.merged {
                return Err(Error::UnsafeEdit {
                    operation: "insert_table_column",
                    reason: "the table has cells merged across columns",
                });
            }
            if position > widths.len() {
                return Err(invalid("the column position is past the column count"));
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
            for (cells, end) in &grid.rows {
                let at = cells.get(position).map_or(*end, |span| span.start);
                if cells.len() < position {
                    return Err(invalid("a row has fewer cells than the column position"));
                }
                edits.push((at..at, cell.clone().into_bytes()));
            }
            next_widths.insert(position, width);
        },
        Change::Remove { column } => {
            if grid.merged {
                return Err(Error::UnsafeEdit {
                    operation: "remove_table_column",
                    reason: "the table has cells merged across columns",
                });
            }
            let Some((span, _)) = grid.columns.get(column) else {
                return Err(invalid("the column position is past the last column"));
            };
            if grid.columns.len() == 1 {
                return Err(Error::UnsafeEdit {
                    operation: "remove_table_column",
                    reason: "a table keeps at least one column",
                });
            }
            edits.push((span.clone(), Vec::new()));
            for (cells, _) in &grid.rows {
                let span = cells
                    .get(column)
                    .ok_or_else(|| invalid("a row has no cell at the column position"))?;
                edits.push((span.clone(), Vec::new()));
            }
            next_widths.remove(column);
        },
    }
    if let Some(extent) = &grid.extent {
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
        return Err(invalid("column edits overlap"));
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
}
