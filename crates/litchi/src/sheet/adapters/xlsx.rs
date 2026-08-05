//! XLSX-to-umbrella spreadsheet bridge.
//!
//! The standalone XLSX crate intentionally exposes its lossless semantic
//! model instead of the umbrella's older dynamic traits. This adapter keeps
//! that ownership boundary explicit and converts only at the high-level
//! facade seam.

use crate::ooxml::xlsx::{self, Address, Rect};
use litchi_core::sheet::{
    Cell as CoreCell, CellIterator, CellValue, Result as SheetResult, RowIterator, WorkbookTrait,
    Worksheet as CoreWorksheet, WorksheetIterator,
};

type BoxError = Box<dyn std::error::Error + Send + Sync>;

fn boxed_error(error: impl std::fmt::Display) -> BoxError {
    Box::new(litchi_core::Error::Other(error.to_string()))
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

/// Internal dynamic-trait view over a standalone XLSX workbook snapshot.
#[derive(Debug)]
pub(crate) struct Workbook {
    workbook: xlsx::Workbook,
    names: Box<[String]>,
}

impl Workbook {
    pub(crate) fn new(workbook: xlsx::Workbook) -> Self {
        let names = workbook
            .sheets()
            .map(|sheet| sheet.name().to_owned())
            .collect();
        Self { workbook, names }
    }
}

impl WorkbookTrait for Workbook {
    fn active_worksheet(&self) -> SheetResult<Box<dyn CoreWorksheet + '_>> {
        let worksheet = self
            .workbook
            .active_sheet()
            .ok_or_else(|| boxed_error("XLSX workbook has no active worksheet"))?;
        Ok(Box::new(Worksheet { worksheet }))
    }

    fn worksheet_names(&self) -> &[String] {
        &self.names
    }

    fn worksheet_by_name(&self, name: &str) -> SheetResult<Box<dyn CoreWorksheet + '_>> {
        let worksheet = self
            .workbook
            .sheet(name)
            .map_err(boxed_error)?
            .ok_or_else(|| boxed_error(format!("XLSX worksheet '{name}' was not found")))?;
        Ok(Box::new(Worksheet { worksheet }))
    }

    fn worksheet_by_index(&self, index: usize) -> SheetResult<Box<dyn CoreWorksheet + '_>> {
        let worksheet = self
            .workbook
            .sheet(index)
            .map_err(boxed_error)?
            .ok_or_else(|| boxed_error(format!("XLSX worksheet index {index} is out of bounds")))?;
        Ok(Box::new(Worksheet { worksheet }))
    }

    fn worksheets(&self) -> Box<dyn WorksheetIterator<'_> + '_> {
        Box::new(Worksheets {
            workbook: self,
            index: 0,
        })
    }

    fn worksheet_count(&self) -> usize {
        self.workbook.len()
    }

    fn active_sheet_index(&self) -> usize {
        self.workbook
            .active_sheet()
            .map_or(0, |sheet| sheet.position())
    }

    fn is_1904_date_system(&self) -> bool {
        matches!(self.workbook.date_system(), xlsx::DateSystem::Excel1904)
    }
}

struct Worksheets<'a> {
    workbook: &'a Workbook,
    index: usize,
}

impl<'a> WorksheetIterator<'a> for Worksheets<'a> {
    fn next(&mut self) -> Option<SheetResult<Box<dyn CoreWorksheet + 'a>>> {
        if self.index >= self.workbook.worksheet_count() {
            return None;
        }
        let index = self.index;
        self.index += 1;
        Some(self.workbook.worksheet_by_index(index))
    }
}

#[derive(Debug, Clone)]
struct Worksheet {
    worksheet: xlsx::Worksheet,
}

impl Worksheet {
    fn extent(&self) -> SheetResult<Option<Rect>> {
        self.worksheet.stored_extent().map_err(boxed_error)
    }

    fn value_at(&self, row: u32, column: u32) -> SheetResult<CellValue> {
        match self.worksheet.cell((row, column)).map_err(boxed_error)? {
            xlsx::cell::View::Stored(cell) => Ok(convert_cell(cell)),
            xlsx::cell::View::Missing | xlsx::cell::View::Covered(_) => Ok(CellValue::Empty),
            _ => Ok(CellValue::Empty),
        }
    }

    fn dimensions_inner(&self) -> SheetResult<Option<(u32, u32, u32, u32)>> {
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

    fn row_values(&self, row: u32) -> SheetResult<Vec<CellValue>> {
        let Some((_, _, _, end_column)) = self.dimensions_inner()? else {
            return Ok(Vec::new());
        };
        (0..=end_column)
            .map(|column| self.value_at(row, column))
            .collect()
    }

    fn stored_cells(&self) -> SheetResult<Vec<XlsxCell>> {
        let Some(extent) = self.extent()? else {
            return Ok(Vec::new());
        };
        Ok(self
            .worksheet
            .cells(extent)
            .map_err(boxed_error)?
            .map(|(address, cell)| {
                XlsxCell::new(
                    address.row().get(),
                    address.column().get(),
                    convert_cell(cell),
                )
            })
            .collect())
    }
}

impl CoreWorksheet for Worksheet {
    fn name(&self) -> &str {
        self.worksheet.name()
    }

    fn row_count(&self) -> usize {
        self.dimensions_inner()
            .ok()
            .flatten()
            .map_or(0, |(_, _, row, _)| row as usize + 1)
    }

    fn column_count(&self) -> usize {
        self.dimensions_inner()
            .ok()
            .flatten()
            .map_or(0, |(_, _, _, column)| column as usize + 1)
    }

    fn dimensions(&self) -> Option<(u32, u32, u32, u32)> {
        self.dimensions_inner().ok().flatten()
    }

    fn cell(&self, row: u32, column: u32) -> SheetResult<Box<dyn CoreCell + '_>> {
        Ok(Box::new(XlsxCell::new(
            row,
            column,
            self.value_at(row.saturating_sub(1), column.saturating_sub(1))?,
        )))
    }

    fn cell_by_coordinate(&self, coordinate: &str) -> SheetResult<Box<dyn CoreCell + '_>> {
        let address = Address::from_a1(coordinate).map_err(boxed_error)?;
        self.cell(address.row().get() + 1, address.column().get() + 1)
    }

    fn cells(&self) -> Box<dyn CellIterator<'_> + '_> {
        Box::new(Cells {
            cells: self.stored_cells().unwrap_or_default().into_iter(),
        })
    }

    fn rows(&self) -> Box<dyn RowIterator<'_> + '_> {
        Box::new(Rows {
            worksheet: self,
            index: 0,
            end: self.row_count(),
        })
    }

    fn row(&self, row_idx: usize) -> SheetResult<std::borrow::Cow<'_, [CellValue]>> {
        Ok(std::borrow::Cow::Owned(self.row_values(row_idx as u32)?))
    }

    fn cell_value(&self, row: u32, column: u32) -> SheetResult<std::borrow::Cow<'_, CellValue>> {
        Ok(std::borrow::Cow::Owned(self.value_at(
            row.saturating_sub(1),
            column.saturating_sub(1),
        )?))
    }
}

struct Cells {
    cells: std::vec::IntoIter<XlsxCell>,
}

impl<'a> CellIterator<'a> for Cells {
    fn next(&mut self) -> Option<SheetResult<Box<dyn CoreCell + 'a>>> {
        self.cells
            .next()
            .map(|cell| Ok(Box::new(cell) as Box<dyn CoreCell + 'a>))
    }
}

struct Rows<'a> {
    worksheet: &'a Worksheet,
    index: usize,
    end: usize,
}

impl<'a> RowIterator<'a> for Rows<'a> {
    fn next(&mut self) -> Option<SheetResult<std::borrow::Cow<'a, [CellValue]>>> {
        if self.index >= self.end {
            return None;
        }
        let index = self.index;
        self.index += 1;
        Some(
            self.worksheet
                .row_values(index as u32)
                .map(std::borrow::Cow::Owned),
        )
    }
}

#[derive(Debug, Clone)]
struct XlsxCell {
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
