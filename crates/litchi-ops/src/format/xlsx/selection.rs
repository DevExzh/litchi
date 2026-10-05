//! Selected XLSX worksheet views, moved from the facade's
//! `sheet::selection` with typed errors in place of `Box<dyn Error>`.

use std::fmt;

use litchi_xlsx as xlsx;

use crate::error::{Error, Result};

/// A worksheet selector accepted by [`super::Xlsx::sheet`].
///
/// Names use the XLSX workbook's canonical case-insensitive matching, and
/// positions are zero-based. The identity form is uninhabited.
pub type WorksheetSelector<'a> = xlsx::Selector<'a>;

/// One owned sparse cell selected from an XLSX worksheet.
///
/// The address and exact XLSX cell state do not borrow the source, so they
/// may outlive the selected worksheet handle.
pub type SelectedCell = xlsx::SourceCell;

/// Owned semantic state at one selected XLSX worksheet coordinate.
///
/// Missing coordinates, merged followers, and stored cell records stay
/// distinct. Stored values keep formulas, exact lexical values, and unknown
/// inert cell states without conversion to the generic sheet values.
pub type SelectedCellView = xlsx::SourceCellView;

/// A lifetime-free worksheet selected from an [`Xlsx`](super::Xlsx) workbook.
///
/// The handle keeps either the immutable eager worksheet snapshot or the
/// source-backed worksheet owner chosen when the workbook was opened. Cell
/// and range reads therefore keep the XLSX error, freshness, execution,
/// cache, and read-limit behavior of that owner.
///
/// ```no_run
/// # fn demo(workbook: &litchi_ops::format::xlsx::Xlsx) -> litchi_ops::Result<()> {
/// if let Some(sheet) = workbook.sheet("Data")? {
///     let view = sheet.cell("A1")?;
///     let cells = sheet.cells("A1:C10")?;
///     println!("{} {view:?} {}", sheet.name(), cells.len());
/// }
/// # Ok(())
/// # }
/// ```
#[derive(Clone)]
pub struct SelectedWorksheet {
    inner: SelectedWorksheetInner,
}

#[derive(Clone)]
enum SelectedWorksheetInner {
    Owned(xlsx::Worksheet),
    Source(xlsx::SourceWorksheet),
}

const _: () = {
    const fn assert_handle<T: Send + Sync + Clone>() {}
    assert_handle::<SelectedWorksheet>();
};

impl SelectedWorksheet {
    pub(super) fn from_owned(worksheet: xlsx::Worksheet) -> Self {
        Self {
            inner: SelectedWorksheetInner::Owned(worksheet),
        }
    }

    pub(super) fn from_source(worksheet: xlsx::SourceWorksheet) -> Self {
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
    /// explicitly stored empty cell stay distinct in the returned view.
    ///
    /// # Errors
    ///
    /// Returns the XLSX error as the typed source: a malformed coordinate or
    /// worksheet payload, a non-grid sheet ([`Code::UnsupportedType`]), or a
    /// changed source ([`Code::IoError`]).
    ///
    /// [`Code::UnsupportedType`]: crate::Code::UnsupportedType
    /// [`Code::IoError`]: crate::Code::IoError
    pub fn cell<'a>(&self, at: impl Into<xlsx::At<'a>>) -> Result<SelectedCellView> {
        let at = at.into();
        match &self.inner {
            SelectedWorksheetInner::Owned(worksheet) => match worksheet.cell(at)? {
                xlsx::cell::View::Missing => Ok(SelectedCellView::Missing),
                xlsx::cell::View::Covered(range) => Ok(SelectedCellView::Covered(range)),
                xlsx::cell::View::Stored(cell) => Ok(SelectedCellView::Stored(cell.clone())),
                _ => Err(Error::from(xlsx::Error::Unsupported {
                    feature: "unrecognized worksheet cell-view variant",
                })),
            },
            SelectedWorksheetInner::Source(worksheet) => Ok(worksheet.cell(at)?),
        }
    }

    /// Read the sparse stored cells inside a checked XLSX range.
    ///
    /// A1 endpoints are inclusive. Raw `(start_row, start_column, end_row,
    /// end_column)` bounds are zero-based and half-open. Missing cells and
    /// merged followers are not synthesized into the returned vector.
    ///
    /// # Errors
    ///
    /// As for [`Self::cell`], plus an allocation failure for the result.
    pub fn cells<'a>(&self, area: impl Into<xlsx::Area<'a>>) -> Result<Vec<SelectedCell>> {
        let area = area.into();
        match &self.inner {
            SelectedWorksheetInner::Owned(worksheet) => {
                let cells = worksheet.cells(area)?;
                let (minimum, maximum) = cells.size_hint();
                let mut selected = Vec::new();
                selected
                    .try_reserve_exact(maximum.unwrap_or(minimum))
                    .map_err(cells_allocation)?;
                for (address, cell) in cells {
                    if selected.len() == selected.capacity() {
                        selected.try_reserve(1).map_err(cells_allocation)?;
                    }
                    selected.push(SelectedCell {
                        address,
                        cell: cell.clone(),
                    });
                }
                Ok(selected)
            },
            SelectedWorksheetInner::Source(worksheet) => Ok(worksheet.cells(area)?),
        }
    }
}

fn cells_allocation(source: std::collections::TryReserveError) -> Error {
    Error::from(xlsx::Error::Allocation {
        resource: "unified selected worksheet cells",
        source,
    })
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
