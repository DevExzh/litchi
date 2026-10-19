//! Table, row and cell wrappers, moved from the facade's `document::table`.

use super::CellMerge;
use crate::Result;

/// The column span of a cell that participates in no horizontal merge.
const UNMERGED_SPAN: usize = 1;

/// A table of a word-processing document.
///
/// **Scheduled for deletion in M1**, with the rest of [`crate::compat`].
#[derive(Debug, Clone)]
#[non_exhaustive]
pub enum Table {
    /// A legacy Word (`.doc`) table.
    #[cfg(feature = "doc")]
    Doc(Box<litchi_doc::Table>),
    /// A WordprocessingML (`.docx`) table.
    #[cfg(feature = "docx")]
    Docx(Box<litchi_docx::Table>),
    /// A Rich Text Format table.
    #[cfg(feature = "rtf")]
    Rtf(Box<litchi_rtf::Table<'static>>),
    /// An OpenDocument Text table.
    #[cfg(feature = "odt")]
    Odt(Box<litchi_odt::elements::table::Table>),
}

impl Table {
    /// The number of rows.
    ///
    /// Scheduled for deletion in M1.
    ///
    /// # Errors
    ///
    /// Returns the format's typed error when the rows cannot be decoded.
    pub fn row_count(&self) -> Result<usize> {
        match self {
            #[cfg(feature = "doc")]
            Self::Doc(table) => Ok(table.row_count()?),
            #[cfg(feature = "docx")]
            Self::Docx(table) => Ok(table.row_count()?),
            #[cfg(feature = "rtf")]
            Self::Rtf(table) => Ok(table.row_count()),
            #[cfg(feature = "odt")]
            Self::Odt(table) => table.row_count().map_err(super::odt_error),
        }
    }

    /// The rows of the table.
    ///
    /// This allocates and clones the whole row collection; prefer
    /// [`Self::row_count`] and [`Self::row_at`] to visit rows one by one.
    ///
    /// Scheduled for deletion in M1.
    ///
    /// # Errors
    ///
    /// Returns the format's typed error when the rows cannot be decoded.
    pub fn rows(&self) -> Result<Vec<Row>> {
        match self {
            #[cfg(feature = "doc")]
            Self::Doc(table) => Ok(table
                .rows()?
                .into_iter()
                .map(|row| Row::Doc(Box::new(row)))
                .collect()),
            #[cfg(feature = "docx")]
            Self::Docx(table) => Ok(table
                .rows()?
                .into_iter()
                .map(|row| Row::Docx(Box::new(row)))
                .collect()),
            #[cfg(feature = "rtf")]
            Self::Rtf(table) => Ok(table
                .rows()
                .iter()
                .cloned()
                .map(|row| Row::Rtf(Box::new(row)))
                .collect()),
            #[cfg(feature = "odt")]
            Self::Odt(table) => Ok(table
                .rows()
                .map_err(super::odt_error)?
                .into_iter()
                .map(|row| Row::Odt(Box::new(row)))
                .collect()),
        }
    }

    /// The row at zero-based `index`, or `None` past the last row.
    ///
    /// Scheduled for deletion in M1.
    ///
    /// # Errors
    ///
    /// Returns the format's typed error when the rows cannot be decoded.
    pub fn row_at(&self, index: impl Into<litchi_core::Position>) -> Result<Option<Row>> {
        let index = index.into().get();
        match self {
            #[cfg(feature = "doc")]
            Self::Doc(table) => Ok(table
                .rows()?
                .get(index)
                .cloned()
                .map(|row| Row::Doc(Box::new(row)))),
            #[cfg(feature = "docx")]
            Self::Docx(table) => Ok(table
                .rows()?
                .get(index)
                .cloned()
                .map(|row| Row::Docx(Box::new(row)))),
            #[cfg(feature = "rtf")]
            Self::Rtf(table) => Ok(table
                .rows()
                .get(index)
                .cloned()
                .map(|row| Row::Rtf(Box::new(row)))),
            #[cfg(feature = "odt")]
            Self::Odt(table) => Ok(table
                .rows()
                .map_err(super::odt_error)?
                .get(index)
                .cloned()
                .map(|row| Row::Odt(Box::new(row)))),
        }
    }
}

/// A row of a word-processing table.
///
/// **Scheduled for deletion in M1**, with the rest of [`crate::compat`].
#[derive(Debug, Clone)]
#[non_exhaustive]
pub enum Row {
    /// A legacy Word (`.doc`) row.
    #[cfg(feature = "doc")]
    Doc(Box<litchi_doc::Row>),
    /// A WordprocessingML (`.docx`) row.
    #[cfg(feature = "docx")]
    Docx(Box<litchi_docx::Row>),
    /// A Rich Text Format row.
    #[cfg(feature = "rtf")]
    Rtf(Box<litchi_rtf::Row<'static>>),
    /// An OpenDocument Text row.
    #[cfg(feature = "odt")]
    Odt(Box<litchi_odt::elements::table::TableRow>),
}

impl Row {
    /// The number of cells in the row.
    ///
    /// Scheduled for deletion in M1.
    ///
    /// # Errors
    ///
    /// Returns the format's typed error when the cells cannot be decoded.
    pub fn cell_count(&self) -> Result<usize> {
        match self {
            #[cfg(feature = "doc")]
            Self::Doc(row) => Ok(row.cell_count()?),
            #[cfg(feature = "docx")]
            Self::Docx(row) => Ok(row.cell_count()?),
            #[cfg(feature = "rtf")]
            Self::Rtf(row) => Ok(row.cell_count()),
            #[cfg(feature = "odt")]
            Self::Odt(row) => row.cell_count().map_err(super::odt_error),
        }
    }

    /// The cells of the row.
    ///
    /// This allocates and clones the whole cell collection; prefer
    /// [`Self::cell_count`] and [`Self::cell_at`] to visit cells one by one.
    ///
    /// Scheduled for deletion in M1.
    ///
    /// # Errors
    ///
    /// Returns the format's typed error when the cells cannot be decoded.
    pub fn cells(&self) -> Result<Vec<Cell>> {
        match self {
            #[cfg(feature = "doc")]
            Self::Doc(row) => Ok(row.cells()?.into_iter().map(Cell::Doc).collect()),
            #[cfg(feature = "docx")]
            Self::Docx(row) => Ok(row.cells()?.into_iter().map(Cell::Docx).collect()),
            #[cfg(feature = "rtf")]
            Self::Rtf(row) => Ok(row
                .cells()
                .iter()
                .cloned()
                .map(Box::new)
                .map(Cell::Rtf)
                .collect()),
            #[cfg(feature = "odt")]
            Self::Odt(row) => Ok(row
                .cells()
                .map_err(super::odt_error)?
                .into_iter()
                .map(Box::new)
                .map(Cell::Odt)
                .collect()),
        }
    }

    /// The cell at zero-based `index`, or `None` past the last cell.
    ///
    /// Scheduled for deletion in M1.
    ///
    /// # Errors
    ///
    /// Returns the format's typed error when the cells cannot be decoded.
    pub fn cell_at(&self, index: impl Into<litchi_core::Position>) -> Result<Option<Cell>> {
        let index = index.into().get();
        match self {
            #[cfg(feature = "doc")]
            Self::Doc(row) => Ok(row.cells()?.get(index).cloned().map(Cell::Doc)),
            #[cfg(feature = "docx")]
            Self::Docx(row) => Ok(row.cells()?.get(index).cloned().map(Cell::Docx)),
            #[cfg(feature = "rtf")]
            Self::Rtf(row) => Ok(row.cells().get(index).cloned().map(Box::new).map(Cell::Rtf)),
            #[cfg(feature = "odt")]
            Self::Odt(row) => Ok(row
                .cells()
                .map_err(super::odt_error)?
                .get(index)
                .cloned()
                .map(Box::new)
                .map(Cell::Odt)),
        }
    }

    /// The true column span of the cell at `index`, resolved with the row as
    /// context.
    ///
    /// Returns `Ok(None)` when the index is out of bounds, or when the cell is
    /// covered by a merge an earlier cell of the row began: a covered cell has
    /// no span of its own. DOCX and ODF read their stored span counts; DOC and
    /// RTF, which tag each covered cell with a role instead, are resolved by
    /// counting the continuation cells that follow the range's first cell.
    ///
    /// Scheduled for deletion in M1.
    ///
    /// # Errors
    ///
    /// Returns the format's typed error when the cells cannot be decoded.
    pub fn grid_span_at(&self, index: impl Into<litchi_core::Position>) -> Result<Option<usize>> {
        let index = index.into().get();
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

/// A cell of a word-processing table.
///
/// **Scheduled for deletion in M1**, with the rest of [`crate::compat`].
#[derive(Debug, Clone)]
#[non_exhaustive]
pub enum Cell {
    /// A legacy Word (`.doc`) cell.
    #[cfg(feature = "doc")]
    Doc(litchi_doc::Cell),
    /// A WordprocessingML (`.docx`) cell.
    #[cfg(feature = "docx")]
    Docx(litchi_docx::Cell),
    /// A Rich Text Format cell.
    #[cfg(feature = "rtf")]
    Rtf(Box<litchi_rtf::Cell<'static>>),
    /// An OpenDocument Text cell.
    #[cfg(feature = "odt")]
    Odt(Box<litchi_odt::elements::table::TableCell>),
}

impl Cell {
    /// The text content of the cell.
    ///
    /// Scheduled for deletion in M1.
    ///
    /// # Errors
    ///
    /// Returns the format's typed error when the cell's text cannot be
    /// decoded.
    pub fn text(&self) -> Result<String> {
        match self {
            #[cfg(feature = "doc")]
            Self::Doc(cell) => Ok(cell.text()?.to_owned()),
            #[cfg(feature = "docx")]
            Self::Docx(cell) => Ok(cell.text()?),
            #[cfg(feature = "rtf")]
            Self::Rtf(cell) => Ok(cell.text().to_owned()),
            #[cfg(feature = "odt")]
            Self::Odt(cell) => cell.text().map_err(super::odt_error),
        }
    }

    /// The number of columns the cell spans, where 1 means no horizontal
    /// merge.
    ///
    /// Only DOCX (`w:gridSpan`) and ODF (`table:number-columns-spanned`)
    /// store a span count on the cell itself. DOC and RTF tag every covered
    /// cell with a merge role instead, so a cell alone cannot know how wide
    /// its range is, and this method reports 1 for them; use
    /// [`Row::grid_span_at`] for a true count in every format, or
    /// [`Self::horizontal_merge`] for the portable merge signal.
    ///
    /// Scheduled for deletion in M1.
    ///
    /// # Errors
    ///
    /// Returns the DOCX typed error when the cell properties cannot be read.
    pub fn grid_span(&self) -> Result<usize> {
        match self {
            // DOC records merge roles per cell, not a span count; the width
            // of the range is only recoverable from the surrounding row.
            #[cfg(feature = "doc")]
            Self::Doc(_) => Ok(1),
            #[cfg(feature = "docx")]
            Self::Docx(cell) => Ok(cell.grid_span()?),
            // RTF `\clmgf`/`\clmrg` are roles, not counts; see
            // `Row::grid_span_at`.
            #[cfg(feature = "rtf")]
            Self::Rtf(_) => Ok(1),
            #[cfg(feature = "odt")]
            Self::Odt(cell) => Ok(cell.colspan().max(1)),
        }
    }

    /// The number of rows the cell spans, where the format stores one.
    ///
    /// Returns `None` for formats that express vertical merges as per-cell
    /// roles (DOC, DOCX and RTF); for those, use [`Self::vertical_merge`] and
    /// walk the following rows. ODF stores `table:number-rows-spanned` and
    /// always yields `Some`.
    ///
    /// Scheduled for deletion in M1.
    ///
    /// # Errors
    ///
    /// None today; the `Result` keeps the surface uniform.
    pub fn row_span(&self) -> Result<Option<usize>> {
        match self {
            #[cfg(feature = "doc")]
            Self::Doc(_) => Ok(None),
            // `w:vMerge` marks participation only; the count is not stored.
            #[cfg(feature = "docx")]
            Self::Docx(_) => Ok(None),
            #[cfg(feature = "rtf")]
            Self::Rtf(_) => Ok(None),
            #[cfg(feature = "odt")]
            Self::Odt(cell) => Ok(Some(cell.rowspan().max(1))),
        }
    }

    /// The horizontal merge state of the cell, resolved for every format,
    /// including the role-based DOC and RTF encodings.
    ///
    /// Scheduled for deletion in M1.
    ///
    /// # Errors
    ///
    /// Returns the DOCX typed error when the cell properties cannot be read.
    pub fn horizontal_merge(&self) -> Result<CellMerge> {
        match self {
            #[cfg(feature = "doc")]
            Self::Doc(cell) => Ok(cell
                .properties()
                .map(|properties| doc_merge(properties.merge_status))
                .unwrap_or_default()),
            // A `w:gridSpan` above 1 absorbs the covered columns outright, so
            // a DOCX cell is either the owner of a range or unmerged.
            #[cfg(feature = "docx")]
            Self::Docx(cell) => Ok(span_merge(cell.grid_span()?)),
            #[cfg(feature = "rtf")]
            Self::Rtf(cell) => Ok(rtf_merge(cell.merge().horizontal)),
            // ODF covered cells are `table:covered-table-cell` elements, which
            // the reader does not surface, so only owners remain.
            #[cfg(feature = "odt")]
            Self::Odt(cell) => Ok(span_merge(cell.colspan())),
        }
    }

    /// The vertical merge state of the cell, resolved for every format.
    ///
    /// Scheduled for deletion in M1.
    ///
    /// # Errors
    ///
    /// Returns the DOCX typed error when the cell properties cannot be read.
    pub fn vertical_merge(&self) -> Result<CellMerge> {
        match self {
            #[cfg(feature = "doc")]
            Self::Doc(cell) => Ok(cell
                .vertical_merge_status()
                .map(vertical_doc_merge)
                .unwrap_or_default()),
            #[cfg(feature = "docx")]
            Self::Docx(cell) => Ok(match cell.v_merge()? {
                None => CellMerge::None,
                Some(litchi_docx::VMergeState::Restart) => CellMerge::Start,
                Some(litchi_docx::VMergeState::Continue) => CellMerge::Continuation,
            }),
            #[cfg(feature = "rtf")]
            Self::Rtf(cell) => Ok(rtf_merge(cell.merge().vertical)),
            #[cfg(feature = "odt")]
            Self::Odt(cell) => Ok(span_merge(cell.rowspan())),
        }
    }

    /// The DOCX `w:vMerge` state of the cell; `None` for every other format,
    /// which has no `w:vMerge` equivalent. Prefer the format-neutral
    /// [`Self::vertical_merge`].
    ///
    /// Scheduled for deletion in M1.
    ///
    /// # Errors
    ///
    /// Returns the DOCX typed error when the cell properties cannot be read.
    #[cfg(feature = "docx")]
    pub fn v_merge(&self) -> Result<Option<litchi_docx::VMergeState>> {
        match self {
            #[cfg(feature = "doc")]
            Self::Doc(_) => Ok(None),
            Self::Docx(cell) => Ok(cell.v_merge()?),
            #[cfg(feature = "rtf")]
            Self::Rtf(_) => Ok(None),
            #[cfg(feature = "odt")]
            Self::Odt(_) => Ok(None),
        }
    }
}

/// Interpret an explicit span count as a merge state. Formats that store
/// counts never surface covered cells, so a count above the unmerged width
/// means this cell owns a range.
#[cfg(any(feature = "docx", feature = "odt"))]
fn span_merge(span: usize) -> CellMerge {
    if span > UNMERGED_SPAN {
        CellMerge::Start
    } else {
        CellMerge::None
    }
}

/// Map a binary-DOC horizontal `TC80` merge role onto [`CellMerge`].
#[cfg(feature = "doc")]
fn doc_merge(status: litchi_doc::parts::tap::CellMergeStatus) -> CellMerge {
    use litchi_doc::parts::tap::CellMergeStatus;
    match status {
        CellMergeStatus::None => CellMerge::None,
        CellMergeStatus::First => CellMerge::Start,
        CellMergeStatus::Merged => CellMerge::Continuation,
    }
}

/// Map a binary-DOC vertical `TC80` merge role onto [`CellMerge`].
#[cfg(feature = "doc")]
fn vertical_doc_merge(status: litchi_doc::parts::tap::VerticalMergeStatus) -> CellMerge {
    use litchi_doc::parts::tap::VerticalMergeStatus;
    match status {
        VerticalMergeStatus::None => CellMerge::None,
        VerticalMergeStatus::First => CellMerge::Start,
        VerticalMergeStatus::Merged => CellMerge::Continuation,
    }
}

/// Map an RTF `\clmgf`/`\clmrg`-style merge role onto [`CellMerge`].
#[cfg(feature = "rtf")]
fn rtf_merge(role: Option<litchi_rtf::TableCellMergeRole>) -> CellMerge {
    match role {
        None => CellMerge::None,
        Some(litchi_rtf::TableCellMergeRole::First) => CellMerge::Start,
        Some(litchi_rtf::TableCellMergeRole::Continuation) => CellMerge::Continuation,
    }
}
