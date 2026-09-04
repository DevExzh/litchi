//! XLSX-selected worksheet views for the unified workbook facade.

use std::fmt;

use super::adapters::boxed_xlsx_error;
use super::types::Result;

/// A worksheet selector accepted by [`super::Workbook::sheet`].
///
/// Names use the XLSX workbook's canonical case-insensitive matching and
/// positions are zero-based. The identity form is intentionally uninhabited.
pub type WorksheetSelector<'a> = crate::xlsx::Selector<'a>;

/// One owned sparse cell selected from an XLSX worksheet.
///
/// The address and exact XLSX cell state are independent of the source and
/// may outlive the selected worksheet handle.
pub type SelectedCell = crate::xlsx::SourceCell;

/// Owned semantic state at one selected XLSX worksheet coordinate.
///
/// Missing coordinates, merged followers, and producer-stored cell records
/// remain distinct. Stored values retain formulas, exact lexical values, and
/// unknown inert cell states without conversion through the legacy facade.
pub type SelectedCellView = crate::xlsx::SourceCellView;

/// A lifetime-free selected worksheet from the unified XLSX facade.
///
/// The handle retains either the immutable eager worksheet snapshot or the
/// source-backed worksheet owner chosen when the workbook was opened. Cell
/// and range reads therefore preserve the underlying XLSX error, freshness,
/// execution, cache, and read-limit behavior.
#[derive(Clone)]
pub struct SelectedWorksheet {
    inner: SelectedWorksheetInner,
}

#[derive(Clone)]
enum SelectedWorksheetInner {
    Owned(crate::xlsx::Worksheet),
    Source(crate::xlsx::SourceWorksheet),
}

impl SelectedWorksheet {
    pub(super) fn from_owned(worksheet: crate::xlsx::Worksheet) -> Self {
        Self {
            inner: SelectedWorksheetInner::Owned(worksheet),
        }
    }

    pub(super) fn from_source(worksheet: crate::xlsx::SourceWorksheet) -> Self {
        Self {
            inner: SelectedWorksheetInner::Source(worksheet),
        }
    }

    /// Developer-facing worksheet name.
    #[must_use]
    pub fn name(&self) -> &str {
        match &self.inner {
            SelectedWorksheetInner::Owned(worksheet) => worksheet.name(),
            SelectedWorksheetInner::Source(worksheet) => worksheet.name(),
        }
    }

    /// Checked zero-based position in workbook order.
    #[must_use]
    pub fn position(&self) -> usize {
        match &self.inner {
            SelectedWorksheetInner::Owned(worksheet) => worksheet.position(),
            SelectedWorksheetInner::Source(worksheet) => worksheet.position(),
        }
    }

    /// Read one exact logical cell by A1 reference or zero-based coordinate.
    ///
    /// A1 strings use Excel's one-based lexical notation. Raw `(row, column)`
    /// pairs are zero-based. A missing record, a merged follower, and an
    /// explicitly stored empty cell remain distinct in the returned view.
    pub fn cell<'a>(&self, at: impl Into<crate::xlsx::At<'a>>) -> Result<SelectedCellView> {
        let at = at.into();
        match &self.inner {
            SelectedWorksheetInner::Owned(worksheet) => {
                let view = worksheet.cell(at).map_err(boxed_xlsx_error)?;
                match view {
                    crate::xlsx::cell::View::Missing => Ok(SelectedCellView::Missing),
                    crate::xlsx::cell::View::Covered(range) => Ok(SelectedCellView::Covered(range)),
                    crate::xlsx::cell::View::Stored(cell) => {
                        Ok(SelectedCellView::Stored(cell.clone()))
                    },
                    _ => Err(boxed_xlsx_error(crate::xlsx::Error::Unsupported {
                        feature: "unrecognized worksheet cell-view variant",
                    })),
                }
            },
            SelectedWorksheetInner::Source(worksheet) => {
                worksheet.cell(at).map_err(boxed_xlsx_error)
            },
        }
    }

    /// Read the sparse stored cells inside a checked XLSX range.
    ///
    /// A1 endpoints are inclusive. Raw `(start_row, start_column, end_row,
    /// end_column)` bounds are zero-based and half-open. Missing cells and
    /// merged followers are not synthesized into the returned vector.
    pub fn cells<'a>(&self, area: impl Into<crate::xlsx::Area<'a>>) -> Result<Vec<SelectedCell>> {
        let area = area.into();
        match &self.inner {
            SelectedWorksheetInner::Owned(worksheet) => {
                let cells = worksheet.cells(area).map_err(boxed_xlsx_error)?;
                let (minimum, maximum) = cells.size_hint();
                let mut selected = Vec::new();
                selected
                    .try_reserve_exact(maximum.unwrap_or(minimum))
                    .map_err(|source| {
                        boxed_xlsx_error(crate::xlsx::Error::Allocation {
                            resource: "unified selected worksheet cells",
                            source,
                        })
                    })?;
                for (address, cell) in cells {
                    if selected.len() == selected.capacity() {
                        selected.try_reserve(1).map_err(|source| {
                            boxed_xlsx_error(crate::xlsx::Error::Allocation {
                                resource: "unified selected worksheet cells",
                                source,
                            })
                        })?;
                    }
                    selected.push(SelectedCell {
                        address,
                        cell: cell.clone(),
                    });
                }
                Ok(selected)
            },
            SelectedWorksheetInner::Source(worksheet) => {
                worksheet.cells(area).map_err(boxed_xlsx_error)
            },
        }
    }
}

impl fmt::Debug for SelectedWorksheet {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SelectedWorksheet")
            .field("name", &self.name())
            .field("position", &self.position())
            .finish_non_exhaustive()
    }
}
