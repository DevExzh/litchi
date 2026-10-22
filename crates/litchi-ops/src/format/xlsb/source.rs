//! The source-backed XLSB owner behind the generic `litchi_core::sheet`
//! traits, moved from the facade's `sheet::adapters::xlsb`.
//!
//! The catalog (names, date system, active sheet) is read once at open, as
//! the facade adapter did. Decoded worksheet snapshots are retained by the
//! leaf workbook’s bounded cache and borrowed for each trait operation.
//!
//! Errors raised here are typed [`crate::Error`] values, which the trait
//! methods report unchanged (the traits' `Error` associated type).

use std::borrow::Cow;

use litchi_core::FileFormat;
use litchi_core::sheet::{Cell as CoreCell, CellValue, Worksheet as CoreWorksheet};
use litchi_xlsb as xlsb;

use crate::error::{Code, Error, Result};

pub(super) fn internal(message: impl Into<Cow<'static, str>>) -> Error {
    Error::new(Code::InternalError, message).with_format(FileFormat::Xlsb)
}

/// Deferred XLSB catalog plus the per-worksheet materialization cache.
pub(super) struct SourceWorkbook {
    workbook: xlsb::SourceBackedWorkbook,
    names: Box<[String]>,
    date1904: bool,
    active_catalog_position: Option<usize>,
    active_worksheet_ordinal: Option<usize>,
}

impl SourceWorkbook {
    /// Read the catalog of `workbook`; no worksheet payload is read.
    pub(super) fn new(workbook: xlsb::SourceBackedWorkbook) -> Result<Self> {
        let names = workbook.worksheet_names()?.into_boxed_slice();
        let date1904 = workbook.is_1904_date_system()?;
        let active_catalog_position = workbook.active_catalog_position()?;
        let active_worksheet_ordinal = workbook.active_worksheet_index()?;
        Ok(Self {
            workbook,
            names,
            date1904,
            active_catalog_position,
            active_worksheet_ordinal,
        })
    }

    /// The same catalog over another snapshot of the same source, such as a
    /// document-property revision: the catalog facts do not change, so they
    /// are not read again.
    pub(super) fn with_workbook(&self, workbook: xlsb::SourceBackedWorkbook) -> Self {
        Self {
            workbook,
            names: self.names.clone(),
            date1904: self.date1904,
            active_catalog_position: self.active_catalog_position,
            active_worksheet_ordinal: self.active_worksheet_ordinal,
        }
    }

    pub(super) fn workbook(&self) -> &xlsb::SourceBackedWorkbook {
        &self.workbook
    }

    pub(super) fn names(&self) -> &[String] {
        &self.names
    }

    pub(super) const fn date1904(&self) -> bool {
        self.date1904
    }

    pub(super) const fn active_worksheet_ordinal(&self) -> Option<usize> {
        self.active_worksheet_ordinal
    }

    /// The facade's `ensure_source_current`.
    ///
    /// The checks and errors of the leaf's
    /// `SourceBackedWorkbook::source_version` (execution context, source
    /// version, execution context), with one source observation: the leaf
    /// reads the version twice with nothing in between, which on a file
    /// source is a second `fstat` per check (record 0310).
    pub(super) fn ensure_current(&self) -> Result<()> {
        let package = self.workbook.source_backed_package();
        let fence = |error| Error::from(xlsb::package::PackageError::from(error));
        package.check_execution().map_err(fence)?;
        package.source_version().map_err(fence)?;
        package.check_execution().map_err(fence)?;
        Ok(())
    }

    /// The workbook-scope repairs so far, followed by those of every
    /// worksheet this owner has materialized, in catalog order.
    pub(super) fn tolerance_report(&self) -> Result<xlsb::ToleranceReport> {
        self.workbook.read_tolerance_report().map_err(Into::into)
    }

    fn materialize<'a, T>(
        &'a self,
        handle: &SourceSheet<'a>,
        operation: impl FnOnce(&xlsb::Worksheet) -> Result<T>,
    ) -> Result<T> {
        self.ensure_current()?;
        let grid = handle.handle.grid().map_err(Error::from)?;
        let result = operation(grid.worksheet());
        self.ensure_current()?;
        result
    }

    pub(super) fn worksheet(&self, index: usize) -> Result<SourceSheet<'_>> {
        self.ensure_current()?;
        let result = self.worksheet_current(index);
        self.ensure_current()?;
        result
    }

    fn worksheet_current(&self, index: usize) -> Result<SourceSheet<'_>> {
        let handle = self
            .workbook
            .worksheet_by_index(index)
            .map_err(Error::from)?
            .ok_or_else(|| {
                Error::new(
                    Code::OutOfRange,
                    format!("XLSB worksheet index {index} is out of bounds"),
                )
                .with_format(FileFormat::Xlsb)
            })?;
        SourceSheet::new(self, handle)
    }

    pub(super) fn worksheet_named(&self, name: &str) -> Result<SourceSheet<'_>> {
        self.ensure_current()?;
        let result = self.worksheet_named_current(name);
        self.ensure_current()?;
        result
    }

    fn worksheet_named_current(&self, name: &str) -> Result<SourceSheet<'_>> {
        let handle = self
            .workbook
            .worksheet_by_name(name)
            .map_err(Error::from)?
            .ok_or_else(|| {
                Error::new(Code::NotFound, "XLSB worksheet was not found")
                    .with_format(FileFormat::Xlsb)
            })?;
        SourceSheet::new(self, handle)
    }

    pub(super) fn active_worksheet(&self) -> Result<SourceSheet<'_>> {
        self.ensure_current()?;
        let result = self
            .active_worksheet_ordinal
            .ok_or_else(|| {
                Error::from(xlsb::package::PackageError::UnsupportedFeature(
                    "XLSB active sheet is not a worksheet",
                ))
            })
            .and_then(|index| self.worksheet_current(index));
        self.ensure_current()?;
        result
    }
}

impl std::fmt::Debug for SourceWorkbook {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("SourceWorkbook")
            .field("worksheet_count", &self.names.len())
            .field("date1904", &self.date1904)
            .field("active_catalog_position", &self.active_catalog_position)
            .finish_non_exhaustive()
    }
}

/// Iterates the worksheets of a source-backed XLSB owner.
pub(crate) struct SourceWorksheets<'a> {
    pub(super) workbook: &'a SourceWorkbook,
    pub(super) index: usize,
}

impl<'a> Iterator for SourceWorksheets<'a> {
    type Item = Result<SourceSheet<'a>>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.index == usize::MAX {
            return None;
        }
        if let Err(error) = self.workbook.ensure_current() {
            self.index = usize::MAX;
            return Some(Err(error));
        }
        if self.index >= self.workbook.names.len() {
            self.index = usize::MAX;
            return match self.workbook.ensure_current() {
                Ok(()) => None,
                Err(error) => Some(Err(error)),
            };
        }
        let index = self.index;
        self.index += 1;
        let result = self.workbook.worksheet_current(index);
        let freshness = self.workbook.ensure_current();
        if result.is_err() || freshness.is_err() {
            self.index = usize::MAX;
        }
        match freshness {
            Err(error) => Some(Err(error)),
            Ok(()) => Some(result),
        }
    }
}

/// One worksheet of a source-backed XLSB owner, read through the generic
/// traits.
pub(crate) struct SourceSheet<'a> {
    workbook: &'a SourceWorkbook,
    handle: xlsb::SourceBackedWorksheet,
    name: String,
}

fn overflow(message: &'static str) -> Error {
    internal(message)
}

const XLSB_ROWS: u32 = 1_048_576;
const XLSB_COLUMNS: u32 = 16_384;

/// Resolve the generic worksheet trait's one-based cell coordinates to the
/// zero-based XLSB grid. The check happens before source materialization, so
/// zero or out-of-grid inputs cannot silently select A1 or trigger a payload
/// read.
fn zero_based_cell(row: u32, column: u32) -> Result<(u32, u32)> {
    if !(1..=XLSB_ROWS).contains(&row) || !(1..=XLSB_COLUMNS).contains(&column) {
        return Err(Error::new(
            Code::OutOfRange,
            "XLSB worksheet cell coordinates are outside the one-based grid",
        )
        .with_format(FileFormat::Xlsb));
    }
    Ok((row - 1, column - 1))
}

fn zero_based_row(row: usize) -> Result<u32> {
    let row = u32::try_from(row).map_err(|_| {
        Error::new(
            Code::OutOfRange,
            "XLSB worksheet row is outside the zero-based grid",
        )
        .with_format(FileFormat::Xlsb)
    })?;
    if row >= XLSB_ROWS {
        return Err(Error::new(
            Code::OutOfRange,
            "XLSB worksheet row is outside the zero-based grid",
        )
        .with_format(FileFormat::Xlsb));
    }
    Ok(row)
}

impl<'a> SourceSheet<'a> {
    fn new(workbook: &'a SourceWorkbook, handle: xlsb::SourceBackedWorksheet) -> Result<Self> {
        let name = handle.name().map_err(Error::from)?.to_owned();
        Ok(Self {
            workbook,
            handle,
            name,
        })
    }

    fn value_at(&self, row: u32, column: u32) -> Result<CellValue> {
        let (row, column) = zero_based_cell(row, column)?;
        // The leaf's stored-cell lookup, which its trait `cell` wraps: an
        // absent position reads as empty.
        self.workbook.materialize(self, |worksheet| {
            Ok(worksheet
                .get_cell(row, column)
                .map_or(CellValue::Empty, |cell| cell.value().clone()))
        })
    }

    fn row_values(&self, row: usize) -> Result<Vec<CellValue>> {
        zero_based_row(row)?;
        self.workbook
            .materialize(self, |worksheet| Ok(worksheet.row(row)?.into_owned()))
    }

    fn copy_cell(cell: &impl CoreCell) -> Result<FacadeCell> {
        let row = cell
            .row()
            .checked_add(1)
            .ok_or_else(|| overflow("XLSB cell row overflow"))?;
        let column = cell
            .column()
            .checked_add(1)
            .ok_or_else(|| overflow("XLSB cell column overflow"))?;
        Ok(FacadeCell {
            row,
            column,
            value: cell.value().clone(),
        })
    }

    fn dimensions_inner(&self) -> Result<Option<(u32, u32, u32, u32)>> {
        self.workbook.materialize(self, |worksheet| {
            let Some((min_row, min_column, max_row, max_column)) = worksheet.dimensions() else {
                return Ok(None);
            };
            let min_row = min_row
                .checked_add(1)
                .ok_or_else(|| overflow("XLSB worksheet minimum row overflow"))?;
            let min_column = min_column
                .checked_add(1)
                .ok_or_else(|| overflow("XLSB worksheet minimum column overflow"))?;
            let max_row = max_row
                .checked_add(1)
                .ok_or_else(|| overflow("XLSB worksheet maximum row overflow"))?;
            let max_column = max_column
                .checked_add(1)
                .ok_or_else(|| overflow("XLSB worksheet maximum column overflow"))?;
            Ok(Some((min_row, min_column, max_row, max_column)))
        })
    }

    fn stored_cells(&self) -> Result<Vec<FacadeCell>> {
        self.workbook.materialize(self, |worksheet| {
            let mut cells = Vec::new();
            for cell in worksheet.cells() {
                cells.push(Self::copy_cell(&cell?)?);
            }
            Ok(cells)
        })
    }
}

impl<'w> CoreWorksheet for SourceSheet<'w> {
    type Error = Error;
    type Cell<'a>
        = FacadeCell
    where
        Self: 'a;
    type Cells<'a>
        = Cells
    where
        Self: 'a;
    type Rows<'a>
        = Rows<'a, 'w>
    where
        Self: 'a;

    fn name(&self) -> &str {
        &self.name
    }

    fn row_count(&self) -> usize {
        self.workbook
            .materialize(self, |worksheet| Ok(worksheet.row_count()))
            .unwrap_or(0)
    }

    fn column_count(&self) -> usize {
        self.workbook
            .materialize(self, |worksheet| Ok(worksheet.column_count()))
            .unwrap_or(0)
    }

    fn dimensions(&self) -> Option<(u32, u32, u32, u32)> {
        self.dimensions_inner().ok().flatten()
    }

    fn cell(&self, row: u32, column: u32) -> Result<FacadeCell> {
        let (row, column) = zero_based_cell(row, column)?;
        self.workbook
            .materialize(self, |worksheet| match worksheet.get_cell(row, column) {
                Some(cell) => Self::copy_cell(cell),
                // The leaf trait's empty cell for an absent position.
                None => Self::copy_cell(&xlsb::Cell::new(row, column, CellValue::Empty)),
            })
    }

    fn cell_by_coordinate(&self, coordinate: &str) -> Result<FacadeCell> {
        let at = xlsb::sheet::At::from(coordinate)
            .resolve()
            .map_err(|error| {
                Error::from(xlsb::package::PackageError::from(error))
                    .caused_by_argument(Code::InvalidArgument)
            })?;
        self.workbook.materialize(self, |worksheet| {
            let cell = worksheet.cell(at.row().get(), at.column().get())?;
            Self::copy_cell(&cell)
        })
    }

    fn cell_value(&self, row: u32, column: u32) -> Result<Cow<'_, CellValue>> {
        Ok(Cow::Owned(self.value_at(row, column)?))
    }

    fn cells(&self) -> Cells {
        match self.stored_cells() {
            Ok(cells) => Cells {
                cells: cells.into_iter(),
                error: None,
            },
            Err(error) => Cells {
                cells: Vec::new().into_iter(),
                error: Some(error),
            },
        }
    }

    fn row(&self, row_idx: usize) -> Result<Cow<'_, [CellValue]>> {
        Ok(Cow::Owned(self.row_values(row_idx)?))
    }

    fn rows(&self) -> Rows<'_, 'w> {
        let (end, error) = match self
            .workbook
            .materialize(self, |worksheet| Ok(worksheet.row_count()))
        {
            Ok(end) => (end, None),
            Err(error) => (0, Some(error)),
        };
        Rows {
            worksheet: self,
            index: 0,
            end,
            error,
        }
    }
}

/// The stored cells of a source-backed XLSB worksheet.
pub(crate) struct Cells {
    cells: std::vec::IntoIter<FacadeCell>,
    error: Option<Error>,
}

impl Iterator for Cells {
    type Item = Result<FacadeCell>;

    fn next(&mut self) -> Option<Self::Item> {
        if let Some(error) = self.error.take() {
            return Some(Err(error));
        }
        self.cells.next().map(Ok)
    }
}

/// The rows of a source-backed XLSB worksheet.
pub(crate) struct Rows<'a, 'workbook> {
    worksheet: &'a SourceSheet<'workbook>,
    index: usize,
    end: usize,
    error: Option<Error>,
}

impl<'a> Iterator for Rows<'a, '_> {
    type Item = Result<Cow<'a, [CellValue]>>;

    fn next(&mut self) -> Option<Self::Item> {
        if let Some(error) = self.error.take() {
            return Some(Err(error));
        }
        if self.index >= self.end {
            return None;
        }
        let index = self.index;
        self.index += 1;
        let result = self.worksheet.row_values(index).map(Cow::Owned);
        if result.is_err() {
            self.index = self.end;
        }
        Some(result)
    }
}

/// One cell of a source-backed XLSB worksheet, one-based.
#[derive(Debug, Clone)]
pub(crate) struct FacadeCell {
    row: u32,
    column: u32,
    value: CellValue,
}

impl CoreCell for FacadeCell {
    fn row(&self) -> u32 {
        self.row
    }

    fn column(&self) -> u32 {
        self.column
    }

    fn coordinate(&self) -> String {
        coordinate(self.row, self.column)
    }

    fn value(&self) -> &CellValue {
        &self.value
    }

    fn is_formula(&self) -> bool {
        matches!(self.value, CellValue::Formula { .. })
    }
}

fn coordinate(row: u32, column: u32) -> String {
    let mut column = column;
    let mut letters = String::new();
    while column > 0 {
        column -= 1;
        let letter = char::from_u32(u32::from(b'A') + (column % 26)).unwrap_or('?');
        letters.push(letter);
        column /= 26;
    }
    let letters: String = letters.chars().rev().collect();
    format!("{letters}{row}")
}
