//! XLSX worksheets behind the generic `litchi_core::sheet` traits, moved
//! from the facade's `sheet::adapters::xlsx`.
//!
//! The typed methods (`*_typed`) return [`crate::Result`] and serve the
//! handle's own reads ([`super::Xlsx::text`]). The trait methods report the
//! same typed [`Error`] (the traits' `Error` associated type); a caller
//! reaches the XLSX leaf error through [`crate::Error::leaf`].

use std::borrow::Cow;

use litchi_core::sheet::{Cell as CoreCell, CellValue, Worksheet as CoreWorksheet};
use litchi_xlsx::{self as xlsx, Address, Rect};

use crate::error::{Code, Error, Result};

fn internal(message: &'static str) -> Error {
    Error::new(Code::InternalError, message).with_format(litchi_core::FileFormat::Xlsx)
}

fn convert_value(value: &xlsx::cell::Value) -> CellValue {
    match value {
        xlsx::cell::Value::Bool(value) => CellValue::Bool(*value),
        xlsx::cell::Value::Number(value) => match value.as_str().parse::<i64>() {
            Ok(value) => CellValue::Int(value),
            Err(_) => CellValue::Float(value.as_f64().unwrap_or_default()),
        },
        xlsx::cell::Value::Text(value) => CellValue::String(value.as_str().to_owned()),
        xlsx::cell::Value::Date(value) => CellValue::String(value.as_str().to_owned()),
        xlsx::cell::Value::Error(value) => CellValue::Error(value.as_str().to_owned()),
        _ => CellValue::Error("unknown XLSX value kind".to_owned()),
    }
}

fn convert_cell(cell: &xlsx::cell::Cell) -> CellValue {
    match cell {
        xlsx::cell::Cell::Empty => CellValue::Empty,
        xlsx::cell::Cell::Value(value) => convert_value(value),
        xlsx::cell::Cell::Formula(formula) => CellValue::Formula {
            formula: formula.text().to_owned(),
            cached_value: formula
                .cached()
                .map(|cache| Box::new(convert_value(cache.value()))),
            is_array: matches!(formula.kind(), xlsx::formula::Kind::Array { .. }),
            array_range: None,
        },
        xlsx::cell::Cell::Unknown(unknown) => {
            CellValue::Error(format!("unknown XLSX cell kind: {}", unknown.kind()))
        },
        _ => CellValue::Error("unknown XLSX cell kind".to_owned()),
    }
}

fn coordinate(row: u32, column: u32) -> Address {
    Address::at(row, column).expect("validated XLSX grid coordinate")
}

/// Resolve the generic worksheet trait's one-based cell coordinates to the
/// zero-based coordinates used by the XLSX leaf. Zero is outside the
/// one-based worksheet grid and must be rejected before the leaf is asked to
/// read its payload (including on an empty worksheet).
fn zero_based_cell(row: u32, column: u32) -> Result<(u32, u32)> {
    if row == 0 || column == 0 {
        return Err(Error::new(
            Code::OutOfRange,
            "XLSX worksheet cell coordinates are one-based",
        )
        .with_format(litchi_core::FileFormat::Xlsx));
    }
    Ok((row - 1, column - 1))
}

/// One worksheet of an [`Xlsx`](super::Xlsx) workbook, eager or
/// source-backed.
#[derive(Clone)]
pub(crate) struct Sheet {
    model: SheetModel,
}

#[derive(Clone)]
enum SheetModel {
    Owned(xlsx::Worksheet),
    Source(xlsx::SourceWorksheet),
}

impl Sheet {
    pub(super) fn owned(worksheet: xlsx::Worksheet) -> Self {
        Self {
            model: SheetModel::Owned(worksheet),
        }
    }

    pub(super) fn source(worksheet: xlsx::SourceWorksheet) -> Self {
        Self {
            model: SheetModel::Source(worksheet),
        }
    }

    fn extent(&self) -> Result<Option<Rect>> {
        let result = match &self.model {
            SheetModel::Owned(worksheet) => worksheet.stored_extent(),
            SheetModel::Source(worksheet) => worksheet.stored_extent(),
        };
        match result {
            Ok(extent) => Ok(extent),
            // The generic worksheet trait models non-grid sheet kinds as
            // empty row and cell iterators. Keep that contract, while
            // malformed grid payload and source errors still reach the
            // fallible iterators below.
            Err(xlsx::Error::NotWorksheet { .. }) => Ok(None),
            Err(error) => Err(error.into()),
        }
    }

    /// The value at a zero-based coordinate the caller passed; one outside
    /// the grid is [`Code::OutOfRange`].
    fn value_at(&self, row: u32, column: u32) -> Result<CellValue> {
        let address = super::selection::resolve_cell((row, column).into())?;
        match &self.model {
            SheetModel::Owned(worksheet) => match worksheet.cell(address)? {
                xlsx::cell::View::Stored(cell) => Ok(convert_cell(cell)),
                _ => Ok(CellValue::Empty),
            },
            SheetModel::Source(worksheet) => match worksheet.cell(address)? {
                xlsx::SourceCellView::Stored(cell) => Ok(convert_cell(&cell)),
                _ => Ok(CellValue::Empty),
            },
        }
    }

    /// Zero-based `(min_row, min_column, max_row, max_column)` of the stored
    /// cells, or `None` for an empty or non-grid sheet.
    pub(super) fn dimensions_typed(&self) -> Result<Option<(u32, u32, u32, u32)>> {
        let Some(extent) = self.extent()? else {
            return Ok(None);
        };
        let start = extent.start();
        let (end_row, end_column) = extent.end();
        Ok(Some((
            start.row().get(),
            start.column().get(),
            end_row.saturating_sub(1),
            end_column.saturating_sub(1),
        )))
    }

    /// The values of one zero-based row, from column 0 to the last stored
    /// column of the sheet. A row outside the grid is [`Code::OutOfRange`].
    pub(super) fn row_values(&self, row: u32) -> Result<Vec<CellValue>> {
        // Validate the zero-based row before checking the stored extent. An
        // empty or non-grid worksheet still has the finite XLSX row grid, so
        // an impossible row must not silently become an empty result.
        super::selection::resolve_cell((row, 0).into())?;
        let Some((_, _, _, end_column)) = self.dimensions_typed()? else {
            return Ok(Vec::new());
        };
        match &self.model {
            SheetModel::Owned(_) => (0..=end_column)
                .map(|column| self.value_at(row, column))
                .collect(),
            SheetModel::Source(worksheet) => {
                let end_row = row
                    .checked_add(1)
                    .ok_or_else(|| internal("XLSX row range overflow"))?;
                let end_column = end_column
                    .checked_add(1)
                    .ok_or_else(|| internal("XLSX column range overflow"))?;
                let width = usize::try_from(end_column)
                    .map_err(|_| internal("XLSX row width does not fit usize"))?;
                let mut values = Vec::new();
                values.try_reserve_exact(width).map_err(|source| {
                    Error::from(xlsx::Error::Allocation {
                        resource: "unified XLSX row values",
                        source,
                    })
                })?;
                values.resize(width, CellValue::Empty);
                for entry in worksheet.cells((row, 0, end_row, end_column))? {
                    let column = usize::try_from(entry.address.column().get())
                        .map_err(|_| internal("XLSX column does not fit usize"))?;
                    let value = values
                        .get_mut(column)
                        .ok_or_else(|| internal("XLSX cell lies outside its row extent"))?;
                    *value = convert_cell(&entry.cell);
                }
                Ok(values)
            },
        }
    }

    /// Number of rows the generic row iterator yields.
    pub(super) fn row_end(&self) -> Result<usize> {
        Ok(self
            .dimensions_typed()?
            .map_or(0, |(_, _, row, _)| row as usize + 1))
    }

    fn stored_cells(&self) -> Result<Cells> {
        let Some(extent) = self.extent()? else {
            return Ok(Cells::empty());
        };
        match &self.model {
            SheetModel::Owned(worksheet) => Ok(Cells {
                inner: CellsInner::Owned(
                    worksheet
                        .cells(extent)?
                        .map(|(address, cell)| {
                            XlsxCell::new(
                                address.row().get(),
                                address.column().get(),
                                convert_cell(cell),
                            )
                        })
                        .collect::<Vec<_>>()
                        .into_iter(),
                ),
            }),
            SheetModel::Source(worksheet) => Ok(Cells {
                inner: CellsInner::Source(worksheet.cells(extent)?.into_iter()),
            }),
        }
    }
}

impl CoreWorksheet for Sheet {
    type Error = Error;
    type Cell<'a> = XlsxCell;
    type Cells<'a> = Cells;
    type Rows<'a> = Rows<'a>;

    fn name(&self) -> &str {
        match &self.model {
            SheetModel::Owned(worksheet) => worksheet.name(),
            SheetModel::Source(worksheet) => worksheet.name(),
        }
    }

    fn row_count(&self) -> usize {
        self.dimensions_typed()
            .ok()
            .flatten()
            .map_or(0, |(_, _, row, _)| row as usize + 1)
    }

    fn column_count(&self) -> usize {
        self.dimensions_typed()
            .ok()
            .flatten()
            .map_or(0, |(_, _, _, column)| column as usize + 1)
    }

    fn dimensions(&self) -> Option<(u32, u32, u32, u32)> {
        self.dimensions_typed()
            .ok()
            .flatten()
            .map(|(a, b, c, d)| (a + 1, b + 1, c + 1, d + 1))
    }

    fn cell(&self, row: u32, column: u32) -> Result<XlsxCell> {
        let (row, column) = zero_based_cell(row, column)?;
        Ok(XlsxCell::new(row, column, self.value_at(row, column)?))
    }

    fn cell_by_coordinate(&self, coordinate: &str) -> Result<XlsxCell> {
        let address = super::selection::resolve_cell(coordinate.into())?;
        self.cell(address.row().get() + 1, address.column().get() + 1)
    }

    fn cells(&self) -> Cells {
        match self.stored_cells() {
            Ok(cells) => cells,
            Err(error) => Cells::error(error),
        }
    }

    fn rows(&self) -> Rows<'_> {
        let (end, error) = match self.row_end() {
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

    fn row(&self, row_idx: usize) -> Result<Cow<'_, [CellValue]>> {
        let row = u32::try_from(row_idx).map_err(|_| {
            Error::new(
                Code::OutOfRange,
                format!("row index {row_idx} is outside the worksheet grid"),
            )
            .with_format(litchi_core::FileFormat::Xlsx)
        })?;
        Ok(Cow::Owned(self.row_values(row)?))
    }

    fn cell_value(&self, row: u32, column: u32) -> Result<Cow<'_, CellValue>> {
        let (row, column) = zero_based_cell(row, column)?;
        Ok(Cow::Owned(self.value_at(row, column)?))
    }
}

/// The stored cells of an XLSX worksheet.
pub(crate) struct Cells {
    inner: CellsInner,
}

enum CellsInner {
    Owned(std::vec::IntoIter<XlsxCell>),
    Source(std::vec::IntoIter<xlsx::SourceCell>),
    Error(Option<Error>),
}

impl Cells {
    fn empty() -> Self {
        Self {
            inner: CellsInner::Owned(Vec::new().into_iter()),
        }
    }

    fn error(error: Error) -> Self {
        Self {
            inner: CellsInner::Error(Some(error)),
        }
    }
}

impl Iterator for Cells {
    type Item = Result<XlsxCell>;

    fn next(&mut self) -> Option<Self::Item> {
        let cell = match &mut self.inner {
            CellsInner::Owned(cells) => cells.next(),
            CellsInner::Source(cells) => cells.next().map(|entry| {
                XlsxCell::new(
                    entry.address.row().get(),
                    entry.address.column().get(),
                    convert_cell(&entry.cell),
                )
            }),
            CellsInner::Error(error) => return error.take().map(Err),
        }?;
        Some(Ok(cell))
    }
}

/// The rows of an XLSX worksheet.
pub(crate) struct Rows<'a> {
    worksheet: &'a Sheet,
    index: usize,
    end: usize,
    error: Option<Error>,
}

impl<'a> Iterator for Rows<'a> {
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
        Some(self.worksheet.row_values(index as u32).map(Cow::Owned))
    }
}

/// One XLSX cell read through the generic traits.
#[derive(Debug, Clone)]
pub(crate) struct XlsxCell {
    row: u32,
    column: u32,
    value: CellValue,
}

impl XlsxCell {
    fn new(row: u32, column: u32, value: CellValue) -> Self {
        Self {
            row: row + 1,
            column: column + 1,
            value,
        }
    }
}

impl CoreCell for XlsxCell {
    fn row(&self) -> u32 {
        self.row
    }

    fn column(&self) -> u32 {
        self.column
    }

    fn coordinate(&self) -> String {
        coordinate(self.row - 1, self.column - 1).a1()
    }

    fn value(&self) -> &CellValue {
        &self.value
    }

    fn is_formula(&self) -> bool {
        matches!(self.value, CellValue::Formula { .. })
    }
}
