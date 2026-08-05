//! Table implementation for Word documents.

#[cfg(any(feature = "doc", feature = "odf"))]
use litchi_core::Error;
use litchi_core::Result;

#[cfg(feature = "doc")]
use crate::doc;

#[cfg(feature = "ooxml")]
use crate::ooxml;

use super::CellMerge;

/// The column span of a cell that participates in no horizontal merge.
const UNMERGED_SPAN: usize = 1;

/// A table in a Word document.
#[derive(Debug, Clone)]
pub enum Table {
    #[cfg(feature = "doc")]
    Doc(Box<doc::Table>),
    #[cfg(feature = "ooxml")]
    Docx(Box<ooxml::docx::Table>),
    #[cfg(feature = "rtf")]
    Rtf(Box<litchi_rtf::Table<'static>>),
    #[cfg(feature = "odf")]
    Odt(Box<litchi_odt::elements::table::Table>),
}

impl Table {
    /// Get the number of rows in the table.
    pub fn row_count(&self) -> Result<usize> {
        match self {
            #[cfg(feature = "doc")]
            Table::Doc(t) => t.row_count().map_err(Error::from),
            #[cfg(feature = "ooxml")]
            Table::Docx(t) => t.row_count().map_err(crate::ooxml::map_ooxml_error),
            #[cfg(feature = "rtf")]
            Table::Rtf(t) => Ok(t.row_count()),
            #[cfg(feature = "odf")]
            Table::Odt(t) => t
                .row_count()
                .map_err(|e| Error::ParseError(format!("Failed to get row count: {}", e))),
        }
    }

    /// Get the rows in this table.
    ///
    /// **Performance Note**: This method allocates and clones the entire row collection.
    /// For better performance when iterating, consider using `row_count()` and `row_at(index)`.
    pub fn rows(&self) -> Result<Vec<Row>> {
        match self {
            #[cfg(feature = "doc")]
            Table::Doc(t) => {
                let rows = t.rows().map_err(Error::from)?;
                Ok(rows
                    .into_iter()
                    .map(|row| Row::Doc(Box::new(row)))
                    .collect())
            },
            #[cfg(feature = "ooxml")]
            Table::Docx(t) => {
                let rows = t.rows().map_err(crate::ooxml::map_ooxml_error)?;
                Ok(rows.into_iter().map(|r| Row::Docx(Box::new(r))).collect())
            },
            #[cfg(feature = "rtf")]
            Table::Rtf(t) => {
                let rows = t.rows();
                Ok(rows
                    .iter()
                    .cloned()
                    .map(|row| Row::Rtf(Box::new(row)))
                    .collect())
            },
            #[cfg(feature = "odf")]
            Table::Odt(t) => {
                let rows = t
                    .rows()
                    .map_err(|e| Error::ParseError(format!("Failed to get rows: {}", e)))?;
                Ok(rows
                    .into_iter()
                    .map(|row| Row::Odt(Box::new(row)))
                    .collect())
            },
        }
    }

    /// Get a specific row by index without allocating a collection.
    ///
    /// This is more efficient than calling `rows()` and then indexing,
    /// as it avoids cloning the entire row collection.
    ///
    /// Returns `None` if the index is out of bounds.
    pub fn row_at(&self, index: usize) -> Result<Option<Row>> {
        match self {
            #[cfg(feature = "doc")]
            Table::Doc(t) => {
                let rows = t.rows().map_err(Error::from)?;
                Ok(rows.get(index).cloned().map(|row| Row::Doc(Box::new(row))))
            },
            #[cfg(feature = "ooxml")]
            Table::Docx(t) => {
                let rows = t.rows().map_err(crate::ooxml::map_ooxml_error)?;
                Ok(rows.get(index).cloned().map(|r| Row::Docx(Box::new(r))))
            },
            #[cfg(feature = "rtf")]
            Table::Rtf(t) => {
                let rows = t.rows();
                Ok(rows.get(index).cloned().map(|row| Row::Rtf(Box::new(row))))
            },
            #[cfg(feature = "odf")]
            Table::Odt(t) => {
                let rows = t
                    .rows()
                    .map_err(|e| Error::ParseError(format!("Failed to get rows: {}", e)))?;
                Ok(rows.get(index).cloned().map(|row| Row::Odt(Box::new(row))))
            },
        }
    }
}

/// A table row in a Word document.
#[derive(Debug, Clone)]
pub enum Row {
    #[cfg(feature = "doc")]
    Doc(Box<doc::Row>),
    #[cfg(feature = "ooxml")]
    Docx(Box<ooxml::docx::Row>),
    #[cfg(feature = "rtf")]
    Rtf(Box<litchi_rtf::Row<'static>>),
    #[cfg(feature = "odf")]
    Odt(Box<litchi_odt::elements::table::TableRow>),
}

impl Row {
    /// Get the number of cells in this row.
    pub fn cell_count(&self) -> Result<usize> {
        match self {
            #[cfg(feature = "doc")]
            Row::Doc(r) => r.cell_count().map_err(Error::from),
            #[cfg(feature = "ooxml")]
            Row::Docx(r) => r.cell_count().map_err(crate::ooxml::map_ooxml_error),
            #[cfg(feature = "rtf")]
            Row::Rtf(r) => Ok(r.cell_count()),
            #[cfg(feature = "odf")]
            Row::Odt(r) => r
                .cell_count()
                .map_err(|e| Error::ParseError(format!("Failed to get cell count: {}", e))),
        }
    }

    /// Get the cells in this row.
    ///
    /// **Performance Note**: This method allocates and clones the entire cell collection.
    /// For better performance when iterating, consider using `cell_count()` and `cell_at(index)`.
    pub fn cells(&self) -> Result<Vec<Cell>> {
        match self {
            #[cfg(feature = "doc")]
            Row::Doc(r) => {
                let cells = r.cells().map_err(Error::from)?;
                Ok(cells.into_iter().map(Cell::Doc).collect())
            },
            #[cfg(feature = "ooxml")]
            Row::Docx(r) => {
                let cells = r.cells().map_err(crate::ooxml::map_ooxml_error)?;
                Ok(cells.into_iter().map(Cell::Docx).collect())
            },
            #[cfg(feature = "rtf")]
            Row::Rtf(r) => {
                let cells = r.cells();
                Ok(cells.iter().cloned().map(Box::new).map(Cell::Rtf).collect())
            },
            #[cfg(feature = "odf")]
            Row::Odt(r) => {
                let cells = r
                    .cells()
                    .map_err(|e| Error::ParseError(format!("Failed to get cells: {}", e)))?;
                Ok(cells.into_iter().map(Box::new).map(Cell::Odt).collect())
            },
        }
    }

    /// Get a specific cell by index without allocating a collection.
    ///
    /// This is more efficient than calling `cells()` and then indexing,
    /// as it avoids cloning the entire cell collection.
    ///
    /// Returns `None` if the index is out of bounds.
    pub fn cell_at(&self, index: usize) -> Result<Option<Cell>> {
        match self {
            #[cfg(feature = "doc")]
            Row::Doc(r) => {
                let cells = r.cells().map_err(Error::from)?;
                Ok(cells.get(index).cloned().map(Cell::Doc))
            },
            #[cfg(feature = "ooxml")]
            Row::Docx(r) => {
                let cells = r.cells().map_err(crate::ooxml::map_ooxml_error)?;
                Ok(cells.get(index).cloned().map(Cell::Docx))
            },
            #[cfg(feature = "rtf")]
            Row::Rtf(r) => {
                let cells = r.cells();
                Ok(cells.get(index).cloned().map(Box::new).map(Cell::Rtf))
            },
            #[cfg(feature = "odf")]
            Row::Odt(r) => {
                let cells = r
                    .cells()
                    .map_err(|e| Error::ParseError(format!("Failed to get cells: {}", e)))?;
                Ok(cells.get(index).cloned().map(Box::new).map(Cell::Odt))
            },
        }
    }

    /// Resolve the true column span of the cell at `index` using row context.
    ///
    /// Returns `Ok(None)` when the index is out of bounds, or when the cell is
    /// covered by a merge that an earlier cell in the row began — a covered cell
    /// has no span of its own.
    ///
    /// This succeeds for every supported format. DOCX and ODF read their stored
    /// span counts directly; DOC and RTF, which tag each covered cell with a
    /// role instead of storing a count, are resolved by counting the run of
    /// continuation cells that follows the range's first cell.
    pub fn grid_span_at(&self, index: usize) -> Result<Option<usize>> {
        let cells = self.cells()?;
        let Some(cell) = cells.get(index) else {
            return Ok(None);
        };

        match cell.horizontal_merge()? {
            CellMerge::Continuation => Ok(None),
            CellMerge::None => Ok(Some(cell.grid_span()?.max(UNMERGED_SPAN))),
            CellMerge::Start => {
                // Count-based formats already know the width; role-based ones
                // must measure the continuation run that follows.
                let stored = cell.grid_span()?;
                if stored > UNMERGED_SPAN {
                    return Ok(Some(stored));
                }
                let mut span = UNMERGED_SPAN;
                for following in &cells[index + 1..] {
                    if following.horizontal_merge()? != CellMerge::Continuation {
                        break;
                    }
                    span += 1;
                }
                Ok(Some(span))
            },
        }
    }
}

/// A table cell in a Word document.
#[derive(Debug, Clone)]
pub enum Cell {
    #[cfg(feature = "doc")]
    Doc(doc::Cell),
    #[cfg(feature = "ooxml")]
    Docx(ooxml::docx::Cell),
    #[cfg(feature = "rtf")]
    Rtf(Box<litchi_rtf::Cell<'static>>),
    #[cfg(feature = "odf")]
    Odt(Box<litchi_odt::elements::table::TableCell>),
}

impl Cell {
    /// Get the text content of the cell.
    pub fn text(&self) -> Result<String> {
        match self {
            #[cfg(feature = "doc")]
            Cell::Doc(c) => c.text().map(|s| s.to_string()).map_err(Error::from),
            #[cfg(feature = "ooxml")]
            Cell::Docx(c) => c
                .text()
                .map(|s| s.to_string())
                .map_err(crate::ooxml::map_ooxml_error),
            #[cfg(feature = "rtf")]
            Cell::Rtf(c) => Ok(c.text().to_string()),
            #[cfg(feature = "odf")]
            Cell::Odt(c) => c
                .text()
                .map_err(|e| Error::ParseError(format!("Failed to get cell text: {}", e))),
        }
    }

    /// Get the grid span (colspan) of this cell.
    ///
    /// Returns the number of columns this cell spans, where 1 means "no
    /// horizontal merge".
    ///
    /// Only DOCX (`w:gridSpan`) and ODF (`table:number-columns-spanned`) store a
    /// span count on the cell itself. Binary DOC and RTF instead tag every
    /// covered cell with a merge role, so a cell in isolation cannot know how
    /// wide its range is and this method reports 1 for them. Use
    /// [`Row::grid_span_at`] to resolve a true column count for every format, or
    /// [`Cell::horizontal_merge`] for the portable merge signal.
    pub fn grid_span(&self) -> Result<usize> {
        match self {
            // DOC records merge roles per cell, not a span count; the width of
            // the range is only recoverable from the surrounding row.
            #[cfg(feature = "doc")]
            Cell::Doc(_) => Ok(1),
            #[cfg(feature = "ooxml")]
            Cell::Docx(c) => c.grid_span().map_err(crate::ooxml::map_ooxml_error),
            // RTF `\clmgf`/`\clmrg` are roles, not counts; see `Row::grid_span_at`.
            #[cfg(feature = "rtf")]
            Cell::Rtf(_) => Ok(1),
            #[cfg(feature = "odf")]
            Cell::Odt(c) => Ok(c.colspan().max(1)),
        }
    }

    /// Get the row span (rowspan) of this cell where the format stores one.
    ///
    /// Returns `None` for formats that express vertical merges as per-cell roles
    /// (DOC, DOCX, and RTF) rather than as an explicit count; for those, use
    /// [`Cell::vertical_merge`] and walk the following rows. ODF stores
    /// `table:number-rows-spanned` directly and always yields `Some`.
    pub fn row_span(&self) -> Result<Option<usize>> {
        match self {
            #[cfg(feature = "doc")]
            Cell::Doc(_) => Ok(None),
            // `w:vMerge` marks participation only; the count is not stored.
            #[cfg(feature = "ooxml")]
            Cell::Docx(_) => Ok(None),
            #[cfg(feature = "rtf")]
            Cell::Rtf(_) => Ok(None),
            #[cfg(feature = "odf")]
            Cell::Odt(c) => Ok(Some(c.rowspan().max(1))),
        }
    }

    /// Get the horizontal merge state of this cell.
    ///
    /// This is the portable form of [`Cell::grid_span`]: it is resolved for
    /// every supported format, including the role-based DOC and RTF encodings.
    pub fn horizontal_merge(&self) -> Result<CellMerge> {
        match self {
            #[cfg(feature = "doc")]
            Cell::Doc(c) => Ok(c
                .properties()
                .map(|p| doc_merge(p.merge_status))
                .unwrap_or_default()),
            // A `w:gridSpan` above 1 absorbs the covered columns outright, so a
            // DOCX cell is either the owner of a range or unmerged.
            #[cfg(feature = "ooxml")]
            Cell::Docx(c) => Ok(span_merge(
                c.grid_span().map_err(crate::ooxml::map_ooxml_error)?,
            )),
            #[cfg(feature = "rtf")]
            Cell::Rtf(c) => Ok(rtf_merge(c.merge().horizontal)),
            // ODF covered cells are `table:covered-table-cell` elements, which
            // the reader does not surface, so only owners remain.
            #[cfg(feature = "odf")]
            Cell::Odt(c) => Ok(span_merge(c.colspan())),
        }
    }

    /// Get the vertical merge state of this cell.
    ///
    /// Unlike the format-specific `Cell::v_merge` accessor, this is available regardless of which format
    /// features are enabled and is resolved for every supported format.
    pub fn vertical_merge(&self) -> Result<CellMerge> {
        match self {
            #[cfg(feature = "doc")]
            Cell::Doc(c) => Ok(c
                .vertical_merge_status()
                .map(vertical_doc_merge)
                .unwrap_or_default()),
            #[cfg(feature = "ooxml")]
            Cell::Docx(c) => Ok(match c.v_merge().map_err(crate::ooxml::map_ooxml_error)? {
                None => CellMerge::None,
                Some(crate::ooxml::docx::VMergeState::Restart) => CellMerge::Start,
                Some(crate::ooxml::docx::VMergeState::Continue) => CellMerge::Continuation,
            }),
            #[cfg(feature = "rtf")]
            Cell::Rtf(c) => Ok(rtf_merge(c.merge().vertical)),
            #[cfg(feature = "odf")]
            Cell::Odt(c) => Ok(span_merge(c.rowspan())),
        }
    }

    /// Get the vertical merge state of this cell as a DOCX-specific value.
    ///
    /// Prefer [`Cell::vertical_merge`], which is format-neutral and does not
    /// require the `ooxml` feature. Formats other than DOCX always return `None`
    /// here because they have no `w:vMerge` equivalent to report.
    #[cfg(feature = "ooxml")]
    pub fn v_merge(&self) -> Result<Option<crate::ooxml::docx::VMergeState>> {
        match self {
            #[cfg(feature = "doc")]
            Cell::Doc(_) => Ok(None),
            Cell::Docx(c) => c.v_merge().map_err(crate::ooxml::map_ooxml_error),
            #[cfg(feature = "rtf")]
            Cell::Rtf(_) => Ok(None),
            #[cfg(feature = "odf")]
            Cell::Odt(_) => Ok(None),
        }
    }
}

/// Interpret an explicit span count as a merge state.
///
/// Formats that store counts never surface covered cells, so a count above the
/// unmerged width means this cell owns a range.
#[cfg(any(feature = "ooxml", feature = "odf"))]
#[inline]
fn span_merge(span: usize) -> CellMerge {
    if span > UNMERGED_SPAN {
        CellMerge::Start
    } else {
        CellMerge::None
    }
}

/// Map a binary-DOC horizontal `TC80` merge role onto the facade enum.
#[cfg(feature = "doc")]
#[inline]
fn doc_merge(status: litchi_doc::parts::tap::CellMergeStatus) -> CellMerge {
    use litchi_doc::parts::tap::CellMergeStatus;
    match status {
        CellMergeStatus::None => CellMerge::None,
        CellMergeStatus::First => CellMerge::Start,
        CellMergeStatus::Merged => CellMerge::Continuation,
    }
}

/// Map a binary-DOC vertical `TC80` merge role onto the facade enum.
#[cfg(feature = "doc")]
#[inline]
fn vertical_doc_merge(status: litchi_doc::parts::tap::VerticalMergeStatus) -> CellMerge {
    use litchi_doc::parts::tap::VerticalMergeStatus;
    match status {
        VerticalMergeStatus::None => CellMerge::None,
        VerticalMergeStatus::First => CellMerge::Start,
        VerticalMergeStatus::Merged => CellMerge::Continuation,
    }
}

/// Map an RTF `\clmgf`/`\clmrg`-style merge role onto the facade enum.
#[cfg(feature = "rtf")]
#[inline]
fn rtf_merge(role: Option<litchi_rtf::TableCellMergeRole>) -> CellMerge {
    match role {
        None => CellMerge::None,
        Some(litchi_rtf::TableCellMergeRole::First) => CellMerge::Start,
        Some(litchi_rtf::TableCellMergeRole::Continuation) => CellMerge::Continuation,
    }
}

#[cfg(test)]
mod tests {
    use super::super::{CellMerge, Document};
    use std::path::PathBuf;

    fn test_data_path() -> PathBuf {
        PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../test-data")
    }

    #[test]
    #[cfg(all(feature = "ooxml", feature = "doc"))]
    fn test_table_row_count_docx() {
        let path = test_data_path().join("ooxml/docx/table_footnotes.docx");
        let doc = Document::open(&path).expect("Failed to open DOCX");
        let tables = doc.tables().expect("Failed to get tables");

        for table in &tables {
            let row_count = table.row_count().expect("Failed to get row count");
            assert!(row_count > 0, "Table should have at least one row");
        }
    }

    #[test]
    #[cfg(all(feature = "ooxml", feature = "doc"))]
    fn test_table_rows_docx() {
        let path = test_data_path().join("ooxml/docx/table_footnotes.docx");
        let doc = Document::open(&path).expect("Failed to open DOCX");
        let tables = doc.tables().expect("Failed to get tables");

        for table in &tables {
            let rows = table.rows().expect("Failed to get rows");
            assert!(!rows.is_empty(), "Expected at least one row");

            for row in &rows {
                let cell_count = row.cell_count().expect("Failed to get cell count");
                assert!(cell_count > 0, "Row should have at least one cell");
            }
        }
    }

    #[test]
    #[cfg(all(feature = "ooxml", feature = "doc"))]
    fn test_table_row_at_docx() {
        let path = test_data_path().join("ooxml/docx/table_footnotes.docx");
        let doc = Document::open(&path).expect("Failed to open DOCX");
        let tables = doc.tables().expect("Failed to get tables");

        for table in &tables {
            let first_row = table.row_at(0).expect("Failed to get row at index 0");
            assert!(first_row.is_some(), "Expected to find first row");

            let nonexistent_row = table
                .row_at(9999)
                .expect("Failed to check row at index 9999");
            assert!(nonexistent_row.is_none(), "Expected no row at index 9999");
        }
    }

    #[test]
    #[cfg(all(feature = "ooxml", feature = "doc"))]
    fn test_table_cells_docx() {
        let path = test_data_path().join("ooxml/docx/table_footnotes.docx");
        let doc = Document::open(&path).expect("Failed to open DOCX");
        let tables = doc.tables().expect("Failed to get tables");

        for table in &tables {
            let rows = table.rows().expect("Failed to get rows");

            for row in &rows {
                let cells = row.cells().expect("Failed to get cells");
                assert!(!cells.is_empty(), "Expected at least one cell");

                for cell in &cells {
                    let _text = cell.text().expect("Failed to get cell text");
                }
            }
        }
    }

    #[test]
    #[cfg(all(feature = "ooxml", feature = "doc"))]
    fn test_table_cell_at_docx() {
        let path = test_data_path().join("ooxml/docx/table_footnotes.docx");
        let doc = Document::open(&path).expect("Failed to open DOCX");
        let tables = doc.tables().expect("Failed to get tables");

        for table in &tables {
            let rows = table.rows().expect("Failed to get rows");

            for row in &rows {
                let first_cell = row.cell_at(0).expect("Failed to get cell at index 0");
                assert!(first_cell.is_some(), "Expected to find first cell");

                let nonexistent_cell = row
                    .cell_at(9999)
                    .expect("Failed to check cell at index 9999");
                assert!(nonexistent_cell.is_none(), "Expected no cell at index 9999");
            }
        }
    }

    #[test]
    #[cfg(all(feature = "ooxml", feature = "doc"))]
    fn test_table_cell_grid_span_docx() {
        let path = test_data_path().join("ooxml/docx/table_footnotes.docx");
        let doc = Document::open(&path).expect("Failed to open DOCX");
        let tables = doc.tables().expect("Failed to get tables");

        for table in &tables {
            let rows = table.rows().expect("Failed to get rows");

            for row in &rows {
                let cells = row.cells().expect("Failed to get cells");

                for cell in &cells {
                    let grid_span = cell.grid_span().expect("Failed to get grid span");
                    assert!(grid_span >= 1, "Grid span should be at least 1");
                }
            }
        }
    }

    #[test]
    #[cfg(all(feature = "ooxml", feature = "doc"))]
    fn test_table_document_with_tables() {
        let test_files = [
            "ooxml/docx/table_footnotes.docx",
            "ooxml/docx/table-alignment.docx",
        ];

        for file in &test_files {
            let path = test_data_path().join(file);
            if path.exists() {
                let doc = Document::open(&path);
                assert!(doc.is_ok(), "Failed to open {}", file);

                if let Ok(d) = doc {
                    let tables = d.tables().expect("Failed to get tables");
                    for table in &tables {
                        let row_count = table.row_count().expect("Failed to get row count");
                        assert!(row_count > 0, "Expected at least one row in {}", file);
                    }
                }
            }
        }

        // This fixture uses the non-OPC
        // `officedocument/.../metadata/core-properties` relationship type.
        // Keep it as a regression for rejecting a present core-properties part
        // that has no valid package-level owner rather than silently reading an
        // incoherent package graph.
        let invalid_path = test_data_path().join("ooxml/docx/table-indent.docx");
        if invalid_path.exists() {
            match Document::open(invalid_path) {
                Err(Error::InvalidFormat(message)) => {
                    assert!(message.contains("core-properties part"));
                    assert!(message.contains("is orphaned"));
                },
                Err(error) => panic!("unexpected table-indent error: {error}"),
                Ok(_) => panic!("table-indent unexpectedly accepted an orphaned core part"),
            }
        }
    }

    /// The RTF fixture is a single row of two cells joined by `\clmgf`
    /// (range owner) and `\clmrg` (covered), so the row spans two columns.
    #[test]
    #[cfg(feature = "rtf")]
    fn rtf_horizontal_merge_roles_resolve_to_a_two_column_span() {
        let path = test_data_path().join("rtf/table-cell-horizontal-merge.rtf");
        let doc = Document::open(&path).expect("Failed to open RTF");
        let tables = doc.tables().expect("Failed to get tables");
        let table = tables.first().expect("fixture has one table");
        let row = table
            .row_at(0)
            .expect("Failed to read row")
            .expect("fixture has one row");

        let cells = row.cells().expect("Failed to get cells");
        assert_eq!(cells.len(), 2, "fixture row has two cells");
        assert_eq!(
            cells[0].horizontal_merge().expect("merge state"),
            CellMerge::Start,
            "\\clmgf marks the first cell as the range owner"
        );
        assert_eq!(
            cells[1].horizontal_merge().expect("merge state"),
            CellMerge::Continuation,
            "\\clmrg marks the second cell as covered"
        );

        assert_eq!(
            row.grid_span_at(0).expect("span"),
            Some(2),
            "the owner spans itself plus one continuation cell"
        );
        assert_eq!(
            row.grid_span_at(1).expect("span"),
            None,
            "a covered cell has no span of its own"
        );
        assert_eq!(row.grid_span_at(99).expect("span"), None);
    }

    /// `\clvmgf` opens a vertical merge; the facade must report it without the
    /// DOCX-specific `v_merge` accessor.
    #[test]
    #[cfg(feature = "rtf")]
    fn rtf_vertical_merge_roles_are_reported() {
        let path = test_data_path().join("rtf/table-cell-vertical-merge.rtf");
        let doc = Document::open(&path).expect("Failed to open RTF");
        let tables = doc.tables().expect("Failed to get tables");

        let mut starts = 0usize;
        for table in &tables {
            for row in &table.rows().expect("Failed to get rows") {
                for cell in &row.cells().expect("Failed to get cells") {
                    if cell.vertical_merge().expect("merge state") == CellMerge::Start {
                        starts += 1;
                    }
                }
            }
        }
        assert!(
            starts > 0,
            "fixture uses \\clvmgf, so at least one vertical merge must start"
        );
    }

    /// ODF stores an explicit `table:number-columns-spanned`, so the span is
    /// readable from the cell alone. This previously returned a hardcoded 1.
    #[test]
    #[cfg(feature = "odf")]
    fn odt_column_span_is_read_from_the_stored_count() {
        let path = test_data_path().join("odf/odt/table-cell-column-span.odt");
        let doc = Document::open(&path).expect("Failed to open ODT");
        let tables = doc.tables().expect("Failed to get tables");

        let mut spanned = Vec::new();
        for (table_index, table) in tables.iter().enumerate() {
            for row in &table.rows().expect("Failed to get rows") {
                for (index, cell) in row.cells().expect("Failed to get cells").iter().enumerate() {
                    let span = cell.grid_span().expect("grid span");
                    if span > 1 {
                        spanned.push((table_index, span));
                        assert_eq!(
                            cell.horizontal_merge().expect("merge state"),
                            CellMerge::Start,
                            "a cell spanning columns owns its range"
                        );
                        assert_eq!(
                            row.grid_span_at(index).expect("span"),
                            Some(span),
                            "row context must agree with the stored count"
                        );
                    }
                    assert!(
                        cell.row_span().expect("row span").is_some(),
                        "ODF always stores a row span"
                    );
                }
            }
        }

        assert!(
            spanned.iter().any(|(_, span)| *span == 3),
            "fixture declares table:number-columns-spanned=\"3\", got {spanned:?}"
        );
    }

    /// Binary DOC records `TC80` merge roles per cell; the span is only
    /// recoverable with row context, which `grid_span_at` supplies.
    #[test]
    #[cfg(feature = "doc")]
    fn doc_merge_roles_resolve_through_row_context() {
        let path = test_data_path().join("ole/doc/table-merged-cells.doc");
        let doc = Document::open(&path).expect("Failed to open DOC");
        let tables = doc.tables().expect("Failed to get tables");
        assert!(!tables.is_empty(), "fixture contains merged tables");

        let mut horizontal_starts = 0usize;
        let mut vertical_participants = 0usize;
        for table in &tables {
            for row in &table.rows().expect("Failed to get rows") {
                let cells = row.cells().expect("Failed to get cells");
                for (index, cell) in cells.iter().enumerate() {
                    match cell.horizontal_merge().expect("merge state") {
                        CellMerge::Start => {
                            horizontal_starts += 1;
                            let span = row.grid_span_at(index).expect("span").expect("owner span");
                            assert!(
                                span >= 1 && span <= cells.len(),
                                "span {span} must stay within the row's {} cells",
                                cells.len()
                            );
                        },
                        CellMerge::Continuation => {
                            assert_eq!(
                                row.grid_span_at(index).expect("span"),
                                None,
                                "covered cells report no span"
                            );
                        },
                        CellMerge::None => {
                            assert_eq!(row.grid_span_at(index).expect("span"), Some(1));
                        },
                    }
                    if cell.vertical_merge().expect("merge state").is_merged() {
                        vertical_participants += 1;
                    }
                }
            }
        }

        // The fixture merges cells vertically. Horizontal `fFirstMerged`
        // merges are rare in Word output (it usually drops cell boundaries
        // instead), so the horizontal mapping is covered by the TAP parser's
        // own unit tests; here it is exercised through the `CellMerge::None`
        // and `Start` arms above.
        assert_eq!(
            horizontal_starts, 0,
            "fixture is not expected to use horizontal TC80 merges"
        );
        assert!(
            vertical_participants > 0,
            "table-merged-cells.doc must expose vertically merged cells"
        );
    }

    /// DOCX absorbs covered columns into `w:gridSpan`, so merged cells are
    /// always range owners and never continuations.
    #[test]
    #[cfg(feature = "ooxml")]
    fn docx_grid_span_owners_are_never_continuations() {
        let path = test_data_path().join("ooxml/docx/drawing.docx");
        let doc = Document::open(&path).expect("Failed to open DOCX");
        let tables = doc.tables().expect("Failed to get tables");

        let mut widest = 1usize;
        for table in &tables {
            for row in &table.rows().expect("Failed to get rows") {
                for (index, cell) in row.cells().expect("Failed to get cells").iter().enumerate() {
                    let span = cell.grid_span().expect("grid span");
                    widest = widest.max(span);
                    assert_ne!(
                        cell.horizontal_merge().expect("merge state"),
                        CellMerge::Continuation,
                        "DOCX never surfaces covered cells"
                    );
                    assert_eq!(row.grid_span_at(index).expect("span"), Some(span));
                    assert!(
                        cell.row_span().expect("row span").is_none(),
                        "w:vMerge carries no span count"
                    );
                }
            }
        }

        assert!(
            widest >= 3,
            "fixture declares w:gridSpan=\"3\", got a widest span of {widest}"
        );
    }

    /// The format-neutral accessor must agree with the DOCX-specific one.
    #[test]
    #[cfg(feature = "ooxml")]
    fn docx_vertical_merge_matches_the_format_specific_accessor() {
        use crate::ooxml::docx::VMergeState;

        let path = test_data_path().join("ooxml/docx/drawing.docx");
        let doc = Document::open(&path).expect("Failed to open DOCX");
        let tables = doc.tables().expect("Failed to get tables");

        for table in &tables {
            for row in &table.rows().expect("Failed to get rows") {
                for cell in &row.cells().expect("Failed to get cells") {
                    let expected = match cell.v_merge().expect("v_merge") {
                        None => CellMerge::None,
                        Some(VMergeState::Restart) => CellMerge::Start,
                        Some(VMergeState::Continue) => CellMerge::Continuation,
                    };
                    assert_eq!(cell.vertical_merge().expect("merge state"), expected);
                }
            }
        }
    }

    #[test]
    #[cfg(feature = "rtf")]
    fn test_table_rtf() {
        let path = test_data_path().join("rtf/chtoutline.rtf");
        let doc = Document::open(&path).expect("Failed to open RTF");
        let tables = doc.tables().expect("Failed to get tables");

        for table in &tables {
            let row_count = table.row_count().expect("Failed to get row count");
            assert!(row_count > 0, "RTF table should have at least one row");

            let rows = table.rows().expect("Failed to get rows");
            for row in &rows {
                let cells = row.cells().expect("Failed to get cells");
                assert!(!cells.is_empty(), "RTF row should have cells");

                for cell in &cells {
                    let _text = cell.text().expect("Failed to get cell text");
                }
            }
        }
    }
}
