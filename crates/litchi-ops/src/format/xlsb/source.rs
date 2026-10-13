//! The source-backed XLSB owner behind the generic `litchi_core::sheet`
//! traits, moved from the facade's `sheet::adapters::xlsb`.
//!
//! The catalog (names, date system, active sheet) is read once at open, as
//! the facade adapter did. Decoded worksheet snapshots are retained by the
//! leaf workbook’s bounded cache and borrowed for each trait operation.
//!
//! Errors raised here are typed [`crate::Error`] values; trait methods box
//! them, because the trait's error type is `Box<dyn Error + Send + Sync>`.

use std::borrow::Cow;

use litchi_core::FileFormat;
use litchi_core::sheet::{
    Cell as CoreCell, CellIterator, CellValue, Result as SheetResult, RowIterator,
    Worksheet as CoreWorksheet, WorksheetIterator,
};
use litchi_xlsb as xlsb;

use crate::error::{Code, Error, Result};

type BoxError = Box<dyn std::error::Error + Send + Sync>;

/// Box a typed error for a `litchi_core::sheet` trait method.
pub(super) fn boxed(error: Error) -> BoxError {
    Box::new(error)
}

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
    pub(super) fn ensure_current(&self) -> Result<()> {
        self.workbook
            .source_backed_package()
            .check_execution()
            .map_err(|error| Error::from(xlsb::package::PackageError::from(error)))?;
        self.workbook.source_version()?;
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
        operation: impl FnOnce(&dyn CoreWorksheet) -> SheetResult<T>,
    ) -> SheetResult<T> {
        self.ensure_current().map_err(boxed)?;
        let grid = handle.handle.grid().map_err(|error| boxed(error.into()))?;
        let result = operation(grid.worksheet());
        self.ensure_current().map_err(boxed)?;
        result
    }

    pub(super) fn worksheet(&self, index: usize) -> SheetResult<Box<dyn CoreWorksheet + '_>> {
        self.ensure_current().map_err(boxed)?;
        let result = self.worksheet_current(index);
        self.ensure_current().map_err(boxed)?;
        result
    }

    fn worksheet_current(&self, index: usize) -> SheetResult<Box<dyn CoreWorksheet + '_>> {
        let handle = self
            .workbook
            .worksheet_by_index(index)
            .map_err(|error| boxed(error.into()))?
            .ok_or_else(|| {
                boxed(
                    Error::new(
                        Code::OutOfRange,
                        format!("XLSB worksheet index {index} is out of bounds"),
                    )
                    .with_format(FileFormat::Xlsb),
                )
            })?;
        Ok(Box::new(SourceSheet::new(self, handle)?))
    }

    pub(super) fn worksheet_named(&self, name: &str) -> SheetResult<Box<dyn CoreWorksheet + '_>> {
        self.ensure_current().map_err(boxed)?;
        let result = self.worksheet_named_current(name);
        self.ensure_current().map_err(boxed)?;
        result
    }

    fn worksheet_named_current(&self, name: &str) -> SheetResult<Box<dyn CoreWorksheet + '_>> {
        let handle = self
            .workbook
            .worksheet_by_name(name)
            .map_err(|error| boxed(error.into()))?
            .ok_or_else(|| {
                boxed(
                    Error::new(Code::NotFound, "XLSB worksheet was not found")
                        .with_format(FileFormat::Xlsb),
                )
            })?;
        Ok(Box::new(SourceSheet::new(self, handle)?))
    }

    pub(super) fn active_worksheet(&self) -> SheetResult<Box<dyn CoreWorksheet + '_>> {
        self.ensure_current().map_err(boxed)?;
        let result = self
            .active_worksheet_ordinal
            .ok_or_else(|| {
                boxed(Error::from(
                    xlsb::package::PackageError::UnsupportedFeature(
                        "XLSB active sheet is not a worksheet".to_string(),
                    ),
                ))
            })
            .and_then(|index| self.worksheet_current(index));
        self.ensure_current().map_err(boxed)?;
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
pub(super) struct SourceWorksheets<'a> {
    pub(super) workbook: &'a SourceWorkbook,
    pub(super) index: usize,
}

impl<'a> WorksheetIterator<'a> for SourceWorksheets<'a> {
    fn next(&mut self) -> Option<SheetResult<Box<dyn CoreWorksheet + 'a>>> {
        if self.index == usize::MAX {
            return None;
        }
        if let Err(error) = self.workbook.ensure_current() {
            self.index = usize::MAX;
            return Some(Err(boxed(error)));
        }
        if self.index >= self.workbook.names.len() {
            self.index = usize::MAX;
            return match self.workbook.ensure_current().map_err(boxed) {
                Ok(()) => None,
                Err(error) => Some(Err(error)),
            };
        }
        let index = self.index;
        self.index += 1;
        let result = self.workbook.worksheet_current(index);
        let freshness = self.workbook.ensure_current().map_err(boxed);
        if result.is_err() || freshness.is_err() {
            self.index = usize::MAX;
        }
        match freshness {
            Err(error) => Some(Err(error)),
            Ok(()) => Some(result),
        }
    }
}

struct SourceSheet<'a> {
    workbook: &'a SourceWorkbook,
    handle: xlsb::SourceBackedWorksheet,
    name: String,
}

fn overflow(message: &'static str) -> BoxError {
    boxed(internal(message))
}

const XLSB_ROWS: u32 = 1_048_576;
const XLSB_COLUMNS: u32 = 16_384;

/// Resolve the generic worksheet trait's one-based cell coordinates to the
/// zero-based XLSB grid. The check happens before source materialization, so
/// zero or out-of-grid inputs cannot silently select A1 or trigger a payload
/// read.
fn zero_based_cell(row: u32, column: u32) -> SheetResult<(u32, u32)> {
    if !(1..=XLSB_ROWS).contains(&row) || !(1..=XLSB_COLUMNS).contains(&column) {
        return Err(boxed(
            Error::new(
                Code::OutOfRange,
                "XLSB worksheet cell coordinates are outside the one-based grid",
            )
            .with_format(FileFormat::Xlsb),
        ));
    }
    Ok((row - 1, column - 1))
}

fn zero_based_row(row: usize) -> SheetResult<u32> {
    let row = u32::try_from(row).map_err(|_| {
        boxed(
            Error::new(
                Code::OutOfRange,
                "XLSB worksheet row is outside the zero-based grid",
            )
            .with_format(FileFormat::Xlsb),
        )
    })?;
    if row >= XLSB_ROWS {
        return Err(boxed(
            Error::new(
                Code::OutOfRange,
                "XLSB worksheet row is outside the zero-based grid",
            )
            .with_format(FileFormat::Xlsb),
        ));
    }
    Ok(row)
}

impl<'a> SourceSheet<'a> {
    fn new(workbook: &'a SourceWorkbook, handle: xlsb::SourceBackedWorksheet) -> SheetResult<Self> {
        let name = handle
            .name()
            .map_err(|error| boxed(error.into()))?
            .to_owned();
        Ok(Self {
            workbook,
            handle,
            name,
        })
    }

    fn value_at(&self, row: u32, column: u32) -> SheetResult<CellValue> {
        let (row, column) = zero_based_cell(row, column)?;
        self.workbook.materialize(self, |worksheet| {
            Ok(worksheet.cell(row, column)?.value().clone())
        })
    }

    fn row_values(&self, row: usize) -> SheetResult<Vec<CellValue>> {
        zero_based_row(row)?;
        self.workbook
            .materialize(self, |worksheet| Ok(worksheet.row(row)?.into_owned()))
    }

    fn copy_cell(cell: &dyn CoreCell) -> SheetResult<FacadeCell> {
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

    fn dimensions_inner(&self) -> SheetResult<Option<(u32, u32, u32, u32)>> {
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

    fn stored_cells(&self) -> SheetResult<Vec<FacadeCell>> {
        self.workbook.materialize(self, |worksheet| {
            let mut cells = Vec::new();
            let mut iterator = worksheet.cells();
            while let Some(cell) = iterator.next() {
                cells.push(Self::copy_cell(cell?.as_ref())?);
            }
            Ok(cells)
        })
    }
}

impl CoreWorksheet for SourceSheet<'_> {
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

    fn cell(&self, row: u32, column: u32) -> SheetResult<Box<dyn CoreCell + '_>> {
        let (row, column) = zero_based_cell(row, column)?;
        let cell = self.workbook.materialize(self, |worksheet| {
            let cell = worksheet.cell(row, column)?;
            Self::copy_cell(cell.as_ref())
        })?;
        Ok(Box::new(cell))
    }

    fn cell_by_coordinate(&self, coordinate: &str) -> SheetResult<Box<dyn CoreCell + '_>> {
        let at = xlsb::sheet::At::from(coordinate)
            .resolve()
            .map_err(|error| {
                boxed(
                    Error::from(xlsb::package::PackageError::from(error))
                        .caused_by_argument(Code::InvalidArgument),
                )
            })?;
        let cell = self.workbook.materialize(self, |worksheet| {
            let cell = worksheet.cell(at.row().get(), at.column().get())?;
            Self::copy_cell(cell.as_ref())
        })?;
        Ok(Box::new(cell))
    }

    fn cell_value(&self, row: u32, column: u32) -> SheetResult<Cow<'_, CellValue>> {
        Ok(Cow::Owned(self.value_at(row, column)?))
    }

    fn cells(&self) -> Box<dyn CellIterator<'_> + '_> {
        match self.stored_cells() {
            Ok(cells) => Box::new(Cells {
                cells: cells.into_iter(),
                error: None,
            }),
            Err(error) => Box::new(Cells {
                cells: Vec::new().into_iter(),
                error: Some(error),
            }),
        }
    }

    fn row(&self, row_idx: usize) -> SheetResult<Cow<'_, [CellValue]>> {
        Ok(Cow::Owned(self.row_values(row_idx)?))
    }

    fn rows(&self) -> Box<dyn RowIterator<'_> + '_> {
        let (end, error) = match self
            .workbook
            .materialize(self, |worksheet| Ok(worksheet.row_count()))
        {
            Ok(end) => (end, None),
            Err(error) => (0, Some(error)),
        };
        Box::new(Rows {
            worksheet: self,
            index: 0,
            end,
            error,
        })
    }
}

struct Cells {
    cells: std::vec::IntoIter<FacadeCell>,
    error: Option<BoxError>,
}

impl<'a> CellIterator<'a> for Cells {
    fn next(&mut self) -> Option<SheetResult<Box<dyn CoreCell + 'a>>> {
        if let Some(error) = self.error.take() {
            return Some(Err(error));
        }
        self.cells.next().map(|cell| {
            let cell: Box<dyn CoreCell + 'a> = Box::new(cell);
            Ok(cell)
        })
    }
}

struct Rows<'a, 'workbook> {
    worksheet: &'a SourceSheet<'workbook>,
    index: usize,
    end: usize,
    error: Option<BoxError>,
}

impl<'a> RowIterator<'a> for Rows<'a, '_> {
    fn next(&mut self) -> Option<SheetResult<Cow<'a, [CellValue]>>> {
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

#[derive(Debug, Clone)]
struct FacadeCell {
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
