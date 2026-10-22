//! Table grid edits (record 0345): a single cell added to a row, a column
//! added to or removed from a table, and a column's width.
//!
//! Each edit rewrites the table as one element replacement, so the grid
//! (`w:tblGrid`), the rows' cells and their widths change together, replay
//! applies them under the table's exact preconditions, and the inverse
//! restores the table's bytes. Bytes outside the spliced cells and grid
//! columns keep their spelling.

use std::sync::Arc;

use litchi_core::Position;
use litchi_core::xml::ReaderOrigin;
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use quick_xml::events::{BytesStart, Event};

use super::super::super::{Conformance, Edit, Error, Refusal, validate_authored_text};
use super::super::{
    NewParagraph, NewRun, Removal, Result, buffer, invalid, paragraph_dependencies,
    paragraph_markup, spliced, staging_add,
};
use super::{
    ElementPath, ElementStep, attribute, bindings_at, close, locate, not_found, open, refusal,
};
use crate::properties::{Twips, Width};

/// The width OfficeCLI gives a grid column when the table has none to
/// average: 2400 twips.
const DEFAULT_COLUMN_WIDTH: u32 = 2400;

/// A cell to add to a row (record 0345): its text, one paragraph, and its
/// preferred width (`w:tcW`).
///
/// ```
/// use litchi_docx::document::NewCell;
/// use litchi_docx::properties::{Twips, Width};
///
/// let cell = NewCell::new().with_text("Total").with_width(Width::Twips(Twips(2400)));
/// assert_eq!(cell.text(), "Total");
/// ```
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct NewCell {
    text: String,
    width: Option<Width>,
}

impl NewCell {
    /// An empty cell: one empty paragraph and no width.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Set the text of the cell's paragraph.
    #[must_use]
    pub fn with_text(mut self, text: impl Into<String>) -> Self {
        self.text = text.into();
        self
    }

    /// Set the cell's preferred width.
    #[must_use]
    pub const fn with_width(mut self, width: Width) -> Self {
        self.width = Some(width);
        self
    }

    /// The text.
    #[must_use]
    pub fn text(&self) -> &str {
        &self.text
    }

    /// The preferred width.
    #[must_use]
    pub const fn width(&self) -> Option<Width> {
        self.width
    }
}

/// One cell of a row, relative to the table fragment.
struct CellLayout {
    start: usize,
    end: usize,
    /// `w:gridSpan`, 1 when absent.
    span: usize,
    /// Whether the cell is vertically or horizontally merged.
    merged: bool,
    /// The `w:tcW` element, when the cell has one.
    width: Option<(usize, usize)>,
}

struct RowLayout {
    /// Offset of the row's end tag; `None` for an empty `<w:tr/>`.
    close: Option<usize>,
    cells: Vec<CellLayout>,
}

/// One grid column, relative to the table fragment.
struct GridColumn {
    start: usize,
    end: usize,
    width: Option<u32>,
}

struct GridLayout {
    start: usize,
    end: usize,
    /// Offset of the grid's end tag; `None` for an empty `<w:tblGrid/>`.
    close: Option<usize>,
    columns: Vec<GridColumn>,
}

/// The parts of a table fragment the grid edits touch.
struct TableLayout {
    /// The prefix of the table's element name.
    prefix: Vec<u8>,
    /// The end of `w:tblPr`, where a missing grid goes.
    properties_end: Option<usize>,
    /// The start of the first row, where a missing grid goes when the table
    /// has no properties.
    first_child: Option<usize>,
    grid: Option<GridLayout>,
    rows: Vec<RowLayout>,
}

impl TableLayout {
    /// Read the table fragment `xml`. Row-level and cell-level content
    /// controls and custom-XML wrappers are refused: readers count the
    /// rows and cells inside them, so a position would not name a direct
    /// child.
    fn read(xml: &[u8]) -> Result<Self> {
        let mut reader = quick_xml::Reader::from_reader(xml);
        let origin = ReaderOrigin::of(xml);
        let mut layout = Self {
            prefix: Vec::new(),
            properties_end: None,
            first_child: None,
            grid: None,
            rows: Vec::new(),
        };
        // (local name, start) of every open element, outermost first.
        let mut stack: Vec<(Vec<u8>, usize)> = Vec::new();
        loop {
            let start = origin
                .offset(reader.buffer_position())
                .ok_or_else(|| invalid("table offset overflow"))?;
            let event = reader
                .read_event()
                .map_err(|_| Error::Document(crate::Error::Xml("invalid Word table XML".into())))?;
            let end = origin
                .offset(reader.buffer_position())
                .ok_or_else(|| invalid("table offset overflow"))?;
            match event {
                Event::Start(element) => {
                    layout.open(&element, &stack, start, end, false)?;
                    stack.push((element.local_name().as_ref().to_vec(), start));
                },
                Event::Empty(element) => layout.open(&element, &stack, start, end, true)?,
                Event::End(_) => {
                    let (local, opened) = stack
                        .pop()
                        .ok_or_else(|| invalid("unbalanced Word table XML"))?;
                    layout.close(&local, &stack, opened, start, end);
                    if stack.is_empty() {
                        return Ok(layout);
                    }
                },
                Event::Eof => return Err(invalid("unterminated Word table XML")),
                _ => {},
            }
        }
    }

    fn open(
        &mut self,
        element: &BytesStart<'_>,
        stack: &[(Vec<u8>, usize)],
        start: usize,
        end: usize,
        empty: bool,
    ) -> Result<()> {
        let local = element.local_name();
        let local = local.as_ref();
        let path: Vec<&[u8]> = stack.iter().map(|(name, _)| name.as_slice()).collect();
        match path.as_slice() {
            [] => {
                if local != b"tbl" {
                    return Err(not_found());
                }
                if empty {
                    return Err(refusal(Refusal::ComplexContent));
                }
                self.prefix = element
                    .name()
                    .prefix()
                    .map_or_else(Vec::new, |prefix| prefix.into_inner().to_vec());
            },
            [b"tbl"] => {
                if self.first_child.is_none() && local != b"tblPr" {
                    self.first_child = Some(start);
                }
                match local {
                    b"sdt" | b"customXml" => return Err(refusal(Refusal::ComplexContent)),
                    b"tblPr" if empty => self.properties_end = Some(end),
                    b"tblGrid" => {
                        self.grid = Some(GridLayout {
                            start,
                            end,
                            close: None,
                            columns: Vec::new(),
                        });
                    },
                    b"tr" => self.rows.push(RowLayout {
                        close: None,
                        cells: Vec::new(),
                    }),
                    _ => {},
                }
            },
            [b"tbl", b"tblGrid"] if local == b"gridCol" => {
                let width = element
                    .unchecked_attributes()
                    .filter_map(std::result::Result::ok)
                    .find(|attribute| attribute.key.local_name().as_ref() == b"w")
                    .and_then(|attribute| std::str::from_utf8(&attribute.value).ok()?.parse().ok());
                if let Some(grid) = self.grid.as_mut() {
                    grid.columns.push(GridColumn {
                        start,
                        end: if empty { end } else { start },
                        width,
                    });
                }
            },
            [b"tbl", b"tr"] => match local {
                b"sdt" | b"customXml" => return Err(refusal(Refusal::ComplexContent)),
                b"tc" => {
                    if let Some(row) = self.rows.last_mut() {
                        row.cells.push(CellLayout {
                            start,
                            end,
                            span: 1,
                            merged: false,
                            width: None,
                        });
                    }
                },
                _ => {},
            },
            [b"tbl", b"tr", b"tc", b"tcPr"] => {
                let cell = self
                    .rows
                    .last_mut()
                    .and_then(|row| row.cells.last_mut())
                    .ok_or_else(|| invalid("table cell properties outside a cell"))?;
                match local {
                    b"gridSpan" => {
                        cell.span = element
                            .unchecked_attributes()
                            .filter_map(std::result::Result::ok)
                            .find(|attribute| attribute.key.local_name().as_ref() == b"val")
                            .and_then(|attribute| {
                                std::str::from_utf8(&attribute.value)
                                    .ok()?
                                    .trim()
                                    .parse()
                                    .ok()
                            })
                            .filter(|span: &usize| *span > 0)
                            .unwrap_or(1);
                    },
                    b"vMerge" | b"hMerge" => cell.merged = true,
                    b"tcW" => cell.width = Some((start, if empty { end } else { start })),
                    _ => {},
                }
            },
            _ => {},
        }
        Ok(())
    }

    fn close(
        &mut self,
        local: &[u8],
        stack: &[(Vec<u8>, usize)],
        opened: usize,
        start: usize,
        end: usize,
    ) {
        let path: Vec<&[u8]> = stack.iter().map(|(name, _)| name.as_slice()).collect();
        match (path.as_slice(), local) {
            ([b"tbl"], b"tblPr") => self.properties_end = Some(end),
            ([b"tbl"], b"tblGrid") => {
                if let Some(grid) = self.grid.as_mut() {
                    grid.end = end;
                    grid.close = Some(start);
                }
            },
            ([b"tbl"], b"tr") => {
                if let Some(row) = self.rows.last_mut() {
                    row.close = Some(start);
                }
            },
            ([b"tbl", b"tblGrid"], b"gridCol") => {
                if let Some(column) = self.grid.as_mut().and_then(|grid| grid.columns.last_mut())
                    && column.start == opened
                {
                    column.end = end;
                }
            },
            ([b"tbl", b"tr"], b"tc") => {
                if let Some(cell) = self.rows.last_mut().and_then(|row| row.cells.last_mut()) {
                    cell.end = end;
                }
            },
            ([b"tbl", b"tr", b"tc", b"tcPr"], b"tcW") => {
                if let Some(cell) = self.rows.last_mut().and_then(|row| row.cells.last_mut())
                    && let Some(width) = cell.width.as_mut()
                    && width.0 == opened
                {
                    width.1 = end;
                }
            },
            _ => {},
        }
    }

    fn widths(&self) -> Vec<Option<u32>> {
        self.grid
            .as_ref()
            .map(|grid| grid.columns.iter().map(|column| column.width).collect())
            .unwrap_or_default()
    }

    /// The cell of `row` that starts at grid column `slot` and spans it
    /// alone, or `None` when a horizontal merge covers the slot or the row
    /// is shorter.
    fn cell_at_slot(row: &RowLayout, slot: usize) -> Option<(usize, &CellLayout)> {
        let mut at = 0usize;
        for (index, cell) in row.cells.iter().enumerate() {
            if at == slot {
                return Some((index, cell));
            }
            at += cell.span;
            if at > slot {
                return None;
            }
        }
        None
    }
}

/// OfficeCLI's average grid width: the mean of the columns' widths (a
/// column without one counts 2400), truncated; 2400 for an empty grid.
fn average_width(widths: &[Option<u32>]) -> u32 {
    if widths.is_empty() {
        return DEFAULT_COLUMN_WIDTH;
    }
    let total: u64 = widths
        .iter()
        .map(|width| u64::from(width.unwrap_or(DEFAULT_COLUMN_WIDTH)))
        .sum();
    u32::try_from(total / widths.len() as u64).unwrap_or(u32::MAX)
}

/// The byte splices of one table rewrite, applied in offset order.
#[derive(Default)]
struct Splices(Vec<(usize, usize, Vec<u8>)>);

impl Splices {
    fn insert(&mut self, at: usize, bytes: Vec<u8>) {
        self.0.push((at, at, bytes));
    }

    fn replace(&mut self, start: usize, end: usize, bytes: Vec<u8>) {
        self.0.push((start, end, bytes));
    }

    fn apply(mut self, xml: &[u8]) -> Result<Vec<u8>> {
        // Stable: insertions at one offset keep their staging order.
        self.0.sort_by_key(|(start, end, _)| (*start, *end));
        let extra = self.0.iter().map(|(_, _, bytes)| bytes.len()).sum();
        let mut output = buffer(staging_add(xml.len(), extra)?)?;
        let mut cursor = 0usize;
        for (start, end, bytes) in self.0 {
            if start < cursor {
                return Err(invalid("overlapping table splices"));
            }
            output.extend_from_slice(&xml[cursor..start]);
            output.extend_from_slice(&bytes);
            cursor = end;
        }
        output.extend_from_slice(&xml[cursor..]);
        Ok(output)
    }
}

/// The prefix new table children are written with, and the namespace a
/// default-namespace table's children declare.
struct Markup<'a> {
    prefix: &'a [u8],
    declare: Option<&'static str>,
    conformance: Conformance,
}

impl Markup<'_> {
    fn open(&self, xml: &mut Vec<u8>, local: &str) {
        open(xml, self.prefix, local);
        if let Some(namespace) = self.declare {
            xml.extend_from_slice(b" xmlns:");
            xml.extend_from_slice(self.prefix);
            xml.extend_from_slice(b"=\"");
            xml.extend_from_slice(namespace.as_bytes());
            xml.push(b'"');
        }
    }

    fn grid_column(&self, width: u32) -> Vec<u8> {
        let mut xml = Vec::new();
        self.open(&mut xml, "gridCol");
        attribute(&mut xml, self.prefix, "w", &width.to_string());
        xml.extend_from_slice(b"/>");
        xml
    }

    fn cell(&self, text: &str, width: Option<Width>) -> Result<Vec<u8>> {
        validate_authored_text(text).map_err(refusal)?;
        let mut xml = Vec::new();
        self.open(&mut xml, "tc");
        xml.push(b'>');
        if let Some(width) = width {
            let (value, kind) = width_attributes(width)?;
            open(&mut xml, self.prefix, "tcPr");
            xml.push(b'>');
            open(&mut xml, self.prefix, "tcW");
            attribute(&mut xml, self.prefix, "w", &value);
            attribute(&mut xml, self.prefix, "type", kind);
            xml.extend_from_slice(b"/>");
            close(&mut xml, self.prefix, "tcPr");
        }
        let paragraph = if text.is_empty() {
            NewParagraph::new()
        } else {
            NewParagraph::new().with_run(NewRun::new(text))
        };
        let strict = matches!(self.conformance, Conformance::Strict);
        let markup = paragraph_markup(Some(self.prefix), self.conformance, &paragraph, strict)
            .map_err(spliced(Position::new(0)))?;
        xml.extend_from_slice(&markup);
        close(&mut xml, self.prefix, "tc");
        Ok(xml)
    }
}

/// `w:tcW`'s `w:w` and `w:type` for a width, as OfficeCLI writes them.
fn width_attributes(width: Width) -> Result<(String, &'static str)> {
    Ok(match width {
        Width::Nil => ("0".to_owned(), "nil"),
        Width::Auto => ("0".to_owned(), "auto"),
        Width::Twips(Twips(twips)) => {
            if twips < 0 {
                return Err(refusal(Refusal::InvalidPropertyValue));
            }
            (twips.to_string(), "dxa")
        },
        Width::Percent(fiftieths) => {
            if fiftieths < 0 {
                return Err(refusal(Refusal::InvalidPropertyValue));
            }
            (fiftieths.to_string(), "pct")
        },
    })
}

impl Edit {
    /// The located table at `table`, its bytes and its layout.
    fn table_layout(&self, table: &ElementPath) -> Result<(Arc<Vec<u8>>, TableLayout)> {
        if !matches!(table.steps().last(), Some(ElementStep::Table(_))) {
            return Err(not_found());
        }
        let located = locate(&self.projected, table.steps())?;
        let before = self.element_bytes(located)?;
        let layout = TableLayout::read(&before)?;
        Ok((before, layout))
    }

    fn markup<'a>(&self, layout: &'a TableLayout) -> Markup<'a> {
        let conformance = self.projected.conformance;
        if layout.prefix.is_empty() {
            Markup {
                prefix: b"w",
                declare: Some(conformance.namespace()),
                conformance,
            }
        } else {
            Markup {
                prefix: &layout.prefix,
                declare: None,
                conformance,
            }
        }
    }

    /// Add `cell` to the row at `row`, before its cell `at` (or after its
    /// last cell), and return the new cell's position (record 0345).
    ///
    /// As OfficeCLI's `AddCell`: when the row then spans more grid columns
    /// than the table's grid has, the grid gains columns of the grid's
    /// average width, and every other row gains as many empty cells, so the
    /// rows stay aligned. A table without a grid gets none.
    ///
    /// ```
    /// use litchi_core::Position;
    /// use litchi_docx::document::{raw, ElementAnchor, ElementPath, ElementStep, NewCell, NewTable};
    ///
    /// let mut edit = raw::Tree::from_xml(
    ///     br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:sectPr/></w:body></w:document>"#
    ///         .to_vec(),
    ///     raw::Limits::default(),
    /// )?;
    /// edit.insert_table_at(&ElementPath::body(), ElementAnchor::End, &NewTable::new(2, 2)?)?;
    /// let row = ElementPath::new([
    ///     ElementStep::Table(Position::new(0)),
    ///     ElementStep::Row(Position::new(0)),
    /// ]);
    /// assert_eq!(edit.insert_cell_at(&row, None, &NewCell::new().with_text("x"))?, Position::new(2));
    /// let bytes = edit.to_bytes(raw::Limits::default())?;
    /// let xml = String::from_utf8_lossy(&bytes).into_owned();
    /// assert_eq!(xml.matches("<w:gridCol ").count(), 3);
    /// assert_eq!(xml.matches("<w:tc>").count(), 6);
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::BlockNotFound`] for a missing table, row or cell,
    /// [`Refusal::ComplexContent`] for wrapped rows or cells,
    /// [`Refusal::StructuralText`] for text a run cannot hold, or a
    /// resource-limit error.
    pub fn insert_cell_at(
        &mut self,
        row: &ElementPath,
        at: Option<Position>,
        cell: &NewCell,
    ) -> Result<Position> {
        self.ensure_unmanaged("insert_cell_at")?;
        let (table, ElementStep::Row(row_position)) = row.split_last().ok_or_else(not_found)?
        else {
            return Err(not_found());
        };
        let (before, layout) = self.table_layout(&table)?;
        let markup = self.markup(&layout);
        let target = layout.rows.get(row_position.get()).ok_or_else(not_found)?;
        let close_at = target
            .close
            .ok_or_else(|| refusal(Refusal::ComplexContent))?;
        let index = match at {
            None => target.cells.len(),
            Some(position) if position.get() == target.cells.len() => target.cells.len(),
            Some(position) if position.get() < target.cells.len() => position.get(),
            Some(_) => return Err(not_found()),
        };
        let offset = target.cells.get(index).map_or_else(
            || target.cells.last().map_or(close_at, |last| last.end),
            |cell| cell.start,
        );
        let mut splices = Splices::default();
        splices.insert(offset, markup.cell(&cell.text, cell.width)?);
        // Keep the grid in step (OfficeCLI BUG-R1-P0-2).
        if let Some(grid) = &layout.grid {
            let span = target.cells.iter().map(|cell| cell.span).sum::<usize>() + 1;
            let missing = span.saturating_sub(grid.columns.len());
            if missing > 0 {
                let width = average_width(&layout.widths());
                let mut columns = Vec::new();
                for _ in 0..missing {
                    columns.extend_from_slice(&markup.grid_column(width));
                }
                append_grid_columns(&mut splices, &before, grid, &markup, columns)?;
                for (other, sibling) in layout.rows.iter().enumerate() {
                    if other == row_position.get() {
                        continue;
                    }
                    let close = sibling
                        .close
                        .ok_or_else(|| refusal(Refusal::ComplexContent))?;
                    let at = sibling.cells.last().map_or(close, |last| last.end);
                    let mut pads = Vec::new();
                    for _ in 0..missing {
                        pads.extend_from_slice(&markup.cell("", None)?);
                    }
                    splices.insert(at, pads);
                }
            }
        }
        let after = splices.apply(&before)?;
        self.replace_element(&table, &before, &Arc::new(after))?;
        Ok(Position::new(index))
    }

    /// Add a column to the table at `table`, before its grid column `at`
    /// (or after the last one), and return the new column's position
    /// (record 0345).
    ///
    /// As OfficeCLI's `AddTableColumn`: the grid gains a `w:gridCol` of
    /// `width` twips (the grid's average width when `None`, 2400 for an
    /// empty grid), and every row gains a cell holding one paragraph of
    /// `text`, at the same position. A column inserted where a row's cell
    /// spans several grid columns or is merged is refused.
    ///
    /// ```
    /// use litchi_core::Position;
    /// use litchi_docx::document::{raw, ElementAnchor, ElementPath, ElementStep, NewTable};
    ///
    /// let mut edit = raw::Tree::from_xml(
    ///     br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:sectPr/></w:body></w:document>"#
    ///         .to_vec(),
    ///     raw::Limits::default(),
    /// )?;
    /// edit.insert_table_at(&ElementPath::body(), ElementAnchor::End, &NewTable::new(2, 2)?)?;
    /// let table = ElementPath::new([ElementStep::Table(Position::new(0))]);
    /// assert_eq!(edit.insert_column_at(&table, Some(Position::new(0)), Some(1200), "Q")?, Position::new(0));
    /// let bytes = edit.to_bytes(raw::Limits::default())?;
    /// let xml = String::from_utf8_lossy(&bytes).into_owned();
    /// assert!(xml.contains(r#"<w:tblGrid><w:gridCol w:w="1200"/>"#));
    /// assert_eq!(xml.matches("<w:t>Q</w:t>").count(), 2);
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::BlockNotFound`] for a missing table,
    /// [`Refusal::ComplexContent`] for merged cells at the slot or wrapped
    /// rows, [`Refusal::InvalidPropertyValue`] for a zero width,
    /// [`Refusal::StructuralText`] for text a run cannot hold, or a
    /// resource-limit error.
    pub fn insert_column_at(
        &mut self,
        table: &ElementPath,
        at: Option<Position>,
        width: Option<u32>,
        text: &str,
    ) -> Result<Position> {
        self.ensure_unmanaged("insert_column_at")?;
        if width == Some(0) {
            return Err(refusal(Refusal::InvalidPropertyValue));
        }
        let (before, layout) = self.table_layout(table)?;
        let markup = self.markup(&layout);
        let widths = layout.widths();
        let slot = at.map_or(widths.len(), |at| at.get().min(widths.len()));
        let width = width.unwrap_or_else(|| average_width(&widths));
        let mut splices = Splices::default();
        let column = markup.grid_column(width);
        match &layout.grid {
            Some(grid) if slot < grid.columns.len() => {
                splices.insert(grid.columns[slot].start, column)
            },
            Some(grid) => append_grid_columns(&mut splices, &before, grid, &markup, column)?,
            None => {
                let mut grid = Vec::new();
                markup.open(&mut grid, "tblGrid");
                grid.push(b'>');
                grid.extend_from_slice(&column);
                close(&mut grid, markup.prefix, "tblGrid");
                let at = layout
                    .properties_end
                    .or(layout.first_child)
                    .ok_or_else(|| refusal(Refusal::ComplexContent))?;
                splices.insert(at, grid);
            },
        }
        for row in &layout.rows {
            let close = row.close.ok_or_else(|| refusal(Refusal::ComplexContent))?;
            let at = if slot < widths.len() {
                match TableLayout::cell_at_slot(row, slot) {
                    Some((_, cell)) if !cell.merged => cell.start,
                    Some(_) => return Err(refusal(Refusal::ComplexContent)),
                    // A row shorter than the grid takes the cell at its end.
                    None if row.cells.iter().map(|cell| cell.span).sum::<usize>() <= slot => {
                        row.cells.last().map_or(close, |last| last.end)
                    },
                    None => return Err(refusal(Refusal::ComplexContent)),
                }
            } else {
                row.cells.last().map_or(close, |last| last.end)
            };
            splices.insert(at, markup.cell(text, None)?);
        }
        let after = splices.apply(&before)?;
        self.replace_element(table, &before, &Arc::new(after))?;
        Ok(Position::new(slot))
    }

    /// Remove grid column `column` of the table at `table`: its
    /// `w:gridCol` and each row's cell in that column (record 0345).
    ///
    /// As OfficeCLI's `RemoveTableColumn`, a column holding a merged or
    /// spanning cell is refused; so are the table's last column and a cell
    /// holding one end of a range whose other end is outside it.
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::BlockNotFound`] for a missing table or column,
    /// [`Refusal::ComplexContent`] for merged cells, wrapped rows or the
    /// last column, a dependency refusal, or a resource-limit error.
    pub fn remove_column_at(&mut self, table: &ElementPath, column: Position) -> Result<&mut Self> {
        self.ensure_unmanaged("remove_column_at")?;
        let (before, layout) = self.table_layout(table)?;
        let grid = layout.grid.as_ref().ok_or_else(not_found)?;
        let slot = column.get();
        let grid_column = grid.columns.get(slot).ok_or_else(not_found)?;
        if grid.columns.len() == 1 {
            return Err(refusal(Refusal::ComplexContent));
        }
        let located = locate(&self.projected, table.steps())?;
        let inherited = bindings_at(self.projected.xml_bytes(), located.start)?;
        let mut splices = Splices::default();
        splices.replace(grid_column.start, grid_column.end, Vec::new());
        for row in &layout.rows {
            match TableLayout::cell_at_slot(row, slot) {
                Some((_, cell)) if cell.merged || cell.span != 1 => {
                    return Err(refusal(Refusal::ComplexContent));
                },
                Some((_, cell)) => {
                    if row.cells.len() == 1 {
                        return Err(refusal(Refusal::ComplexContent));
                    }
                    paragraph_dependencies(&before[cell.start..cell.end], &inherited)
                        .refuse(Removal::Delete)
                        .map_err(refusal)?;
                    splices.replace(cell.start, cell.end, Vec::new());
                },
                None if row.cells.iter().map(|cell| cell.span).sum::<usize>() <= slot => {},
                None => return Err(refusal(Refusal::ComplexContent)),
            }
        }
        let after = splices.apply(&before)?;
        self.replace_element(table, &before, &Arc::new(after))
    }

    /// Set the width of grid column `column` of the table at `table` to
    /// `width` twips (record 0345): its `w:gridCol`, and the `w:tcW` of
    /// each row's cell that occupies that column alone and states a width
    /// in twips, so Word's fixed layout follows. An unchanged width is an
    /// exact no-op.
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::BlockNotFound`] for a missing table or column,
    /// [`Refusal::InvalidPropertyValue`] for a zero width, or a
    /// resource-limit error.
    pub fn set_column_width_at(
        &mut self,
        table: &ElementPath,
        column: Position,
        width: u32,
    ) -> Result<&mut Self> {
        self.ensure_unmanaged("set_column_width_at")?;
        if width == 0 {
            return Err(refusal(Refusal::InvalidPropertyValue));
        }
        let (before, layout) = self.table_layout(table)?;
        let grid = layout.grid.as_ref().ok_or_else(not_found)?;
        let slot = column.get();
        let grid_column = grid.columns.get(slot).ok_or_else(not_found)?;
        let value = width.to_string();
        let mut splices = Splices::default();
        let tag = &before[grid_column.start..grid_column.end];
        if grid_column.width != Some(width) {
            let key = crate::picture_spans::attribute_named(tag, "w")
                .unwrap_or_else(|| qualified(&layout.prefix, "w"));
            let rewritten = crate::picture_spans::with_attribute(tag, &key, &value)
                .ok_or_else(|| invalid("invalid grid column tag"))?;
            splices.replace(grid_column.start, grid_column.end, rewritten);
        }
        for row in &layout.rows {
            let Some((_, cell)) = TableLayout::cell_at_slot(row, slot) else {
                continue;
            };
            let Some((start, end)) = cell.width.filter(|_| cell.span == 1) else {
                continue;
            };
            let tag = &before[start..end];
            let kind = crate::picture_spans::attribute_named(tag, "type")
                .and_then(|key| attribute_value(tag, &key));
            if !matches!(kind.as_deref(), None | Some("dxa")) {
                continue;
            }
            let key = crate::picture_spans::attribute_named(tag, "w")
                .unwrap_or_else(|| qualified(&layout.prefix, "w"));
            if attribute_value(tag, &key).as_deref() == Some(value.as_str()) {
                continue;
            }
            let rewritten = crate::picture_spans::with_attribute(tag, &key, &value)
                .ok_or_else(|| invalid("invalid cell width tag"))?;
            splices.replace(start, end, rewritten);
        }
        if splices.0.is_empty() {
            return Ok(self);
        }
        let after = splices.apply(&before)?;
        self.replace_element(table, &before, &Arc::new(after))
    }
}

/// Stage `columns` at the end of `grid`, opening an empty `<w:tblGrid/>`.
fn append_grid_columns(
    splices: &mut Splices,
    xml: &[u8],
    grid: &GridLayout,
    markup: &Markup<'_>,
    columns: Vec<u8>,
) -> Result<()> {
    match grid.close {
        Some(close) => splices.insert(close, columns),
        None => {
            let tag = &xml[grid.start..grid.end];
            let opened = super::super::opened(tag)?;
            let mut replacement = opened;
            replacement.extend_from_slice(&columns);
            close(&mut replacement, markup.prefix, "tblGrid");
            splices.replace(grid.start, grid.end, replacement);
        },
    }
    Ok(())
}

fn qualified(prefix: &[u8], local: &str) -> String {
    if prefix.is_empty() {
        local.to_owned()
    } else {
        format!("{}:{local}", String::from_utf8_lossy(prefix))
    }
}

/// The value of the attribute named exactly `key` in a start tag.
fn attribute_value(tag: &[u8], key: &str) -> Option<String> {
    let mut reader = quick_xml::Reader::from_reader(tag);
    let (Ok(Event::Start(element)) | Ok(Event::Empty(element))) = reader.read_event() else {
        return None;
    };
    element
        .unchecked_attributes()
        .filter_map(std::result::Result::ok)
        .find(|attribute| attribute.key.as_ref() == key.as_bytes())
        .and_then(|attribute| String::from_utf8(attribute.value.into_owned()).ok())
}

#[cfg(test)]
mod tests {
    use super::*;

    const TABLE: &[u8] = br#"<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p/></w:tc><w:tc><w:p/></w:tc></w:tr><w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>"#;

    #[test]
    fn the_layout_reads_grid_rows_and_spans() {
        let layout = TableLayout::read(TABLE).unwrap();
        assert_eq!(layout.prefix, b"w");
        assert_eq!(layout.widths(), [Some(2000), Some(3000)]);
        assert_eq!(layout.rows.len(), 2);
        assert_eq!(layout.rows[0].cells.len(), 2);
        assert_eq!(layout.rows[1].cells[0].span, 2);
        let width = layout.rows[0].cells[0].width.unwrap();
        assert_eq!(
            &TABLE[width.0..width.1],
            br#"<w:tcW w:w="2000" w:type="dxa"/>"#
        );
        assert!(TableLayout::cell_at_slot(&layout.rows[1], 1).is_none());
        assert_eq!(average_width(&layout.widths()), 2500);
        assert_eq!(average_width(&[]), 2400);
    }

    #[test]
    fn merged_slots_refuse_column_edits() {
        let xml = format!(
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{}<w:sectPr/></w:body></w:document>"#,
            std::str::from_utf8(TABLE).unwrap()
        );
        let snapshot = crate::document::Snapshot::from_xml(xml).unwrap();
        let table = ElementPath::new([ElementStep::Table(Position::new(0))]);
        let mut edit = snapshot.edit();
        // Row 2's first cell spans both grid columns.
        assert!(
            edit.insert_column_at(&table, Some(Position::new(1)), None, "")
                .is_err()
        );
        assert!(edit.remove_column_at(&table, Position::new(0)).is_err());
        // Appending needs no free slot; the width is the grid's average.
        assert_eq!(
            edit.insert_column_at(&table, None, None, "x").unwrap(),
            Position::new(2)
        );
        // A width set rewrites the grid column and the dxa cell width.
        edit.set_column_width_at(&table, Position::new(0), 1800)
            .unwrap();
        let committed = edit.commit().unwrap();
        let text = String::from_utf8_lossy(committed.snapshot().xml_bytes()).into_owned();
        assert!(
            text.contains(
                r#"<w:gridCol w:w="1800"/><w:gridCol w:w="3000"/><w:gridCol w:w="2500"/>"#
            ),
            "{text}"
        );
        assert!(
            text.contains(r#"<w:tcW w:w="1800" w:type="dxa"/>"#),
            "{text}"
        );
    }

    #[test]
    fn wrapped_rows_are_refused() {
        let wrapped = br#"<w:tbl><w:sdt><w:sdtContent><w:tr><w:tc><w:p/></w:tc></w:tr></w:sdtContent></w:sdt></w:tbl>"#;
        assert!(TableLayout::read(wrapped).is_err());
    }
}
