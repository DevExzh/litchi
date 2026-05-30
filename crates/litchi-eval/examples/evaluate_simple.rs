//! Simple in-memory formula evaluation example for `litchi-eval`.
//!
//! This example shows how to plug a minimal `WorkbookTrait` implementation
//! into [`FormulaEvaluator`], evaluate individual cells (including formulas
//! that rely on cached results), and register a global named range with
//! [`FormulaEvaluator::define_name`].
//!
//! Run with:
//!
//! ```sh
//! cargo run -p litchi-eval --example evaluate_simple --all-features
//! ```

use std::borrow::Cow;
use std::collections::HashMap;

use litchi_core::sheet::{
    Cell, CellIterator, CellValue, Result, RowIterator, WorkbookTrait, Worksheet,
    WorksheetIterator,
};
use litchi_eval::FormulaEvaluator;

// ---------------------------------------------------------------------------
// Minimal in-memory workbook implementation
// ---------------------------------------------------------------------------

/// Cell coordinate keyed as `(row, col)`, both 1-based.
type Coord = (u32, u32);

#[derive(Debug)]
struct MemSheet {
    name: String,
    cells: HashMap<Coord, CellValue>,
    dimensions: Option<(u32, u32, u32, u32)>,
}

impl MemSheet {
    fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            cells: HashMap::new(),
            dimensions: None,
        }
    }

    fn set(&mut self, row: u32, col: u32, value: CellValue) {
        self.cells.insert((row, col), value);
        self.dimensions = Some(match self.dimensions {
            None => (row, col, row, col),
            Some((min_r, min_c, max_r, max_c)) => (
                min_r.min(row),
                min_c.min(col),
                max_r.max(row),
                max_c.max(col),
            ),
        });
    }
}

#[derive(Debug)]
struct MemWorkbook {
    sheets: Vec<MemSheet>,
    sheet_names: Vec<String>,
}

impl MemWorkbook {
    fn new() -> Self {
        Self {
            sheets: Vec::new(),
            sheet_names: Vec::new(),
        }
    }

    fn add_sheet(&mut self, sheet: MemSheet) {
        self.sheet_names.push(sheet.name.clone());
        self.sheets.push(sheet);
    }
}

// --- Cell impl --------------------------------------------------------------

struct MemCell<'a> {
    row: u32,
    column: u32,
    value: &'a CellValue,
}

impl<'a> Cell for MemCell<'a> {
    fn row(&self) -> u32 {
        self.row
    }

    fn column(&self) -> u32 {
        self.column
    }

    fn coordinate(&self) -> String {
        format!("{}{}", column_letter(self.column), self.row)
    }

    fn value(&self) -> &CellValue {
        self.value
    }

    fn is_formula(&self) -> bool {
        matches!(self.value, CellValue::Formula { .. })
    }
}

fn column_letter(mut col: u32) -> String {
    let mut buf = Vec::new();
    while col > 0 {
        let rem = (col - 1) % 26;
        buf.push(b'A' + rem as u8);
        col = (col - 1) / 26;
    }
    buf.reverse();
    String::from_utf8(buf).unwrap_or_default()
}

// --- empty iterators -------------------------------------------------------

struct EmptyCellIter;
impl<'a> CellIterator<'a> for EmptyCellIter {
    fn next(&mut self) -> Option<Result<Box<dyn Cell + 'a>>> {
        None
    }
}

struct EmptyRowIter;
impl<'a> RowIterator<'a> for EmptyRowIter {
    fn next(&mut self) -> Option<Result<Cow<'a, [CellValue]>>> {
        None
    }
}

struct EmptyWorksheetIter;
impl<'a> WorksheetIterator<'a> for EmptyWorksheetIter {
    fn next(&mut self) -> Option<Result<Box<dyn Worksheet + 'a>>> {
        None
    }
}

// --- Worksheet impl --------------------------------------------------------

impl Worksheet for MemSheet {
    fn name(&self) -> &str {
        &self.name
    }

    fn row_count(&self) -> usize {
        self.dimensions.map(|(_, _, r, _)| r as usize).unwrap_or(0)
    }

    fn column_count(&self) -> usize {
        self.dimensions.map(|(_, _, _, c)| c as usize).unwrap_or(0)
    }

    fn dimensions(&self) -> Option<(u32, u32, u32, u32)> {
        self.dimensions
    }

    fn cell(&self, row: u32, column: u32) -> Result<Box<dyn Cell + '_>> {
        let value = self.cells.get(&(row, column)).unwrap_or(CellValue::EMPTY);
        Ok(Box::new(MemCell {
            row,
            column,
            value,
        }))
    }

    fn cell_by_coordinate(&self, _coordinate: &str) -> Result<Box<dyn Cell + '_>> {
        Err("cell_by_coordinate is not implemented for MemSheet".into())
    }

    fn cells(&self) -> Box<dyn CellIterator<'_> + '_> {
        Box::new(EmptyCellIter)
    }

    fn rows(&self) -> Box<dyn RowIterator<'_> + '_> {
        Box::new(EmptyRowIter)
    }

    fn row(&self, _row_idx: usize) -> Result<Cow<'_, [CellValue]>> {
        Ok(Cow::Owned(Vec::new()))
    }

    fn cell_value(&self, row: u32, column: u32) -> Result<Cow<'_, CellValue>> {
        Ok(self
            .cells
            .get(&(row, column))
            .map(Cow::Borrowed)
            .unwrap_or(Cow::Borrowed(CellValue::EMPTY)))
    }
}

// --- Workbook impl ---------------------------------------------------------

impl WorkbookTrait for MemWorkbook {
    fn active_worksheet(&self) -> Result<Box<dyn Worksheet + '_>> {
        self.worksheet_by_index(0)
    }

    fn worksheet_names(&self) -> &[String] {
        &self.sheet_names
    }

    fn worksheet_by_name(&self, name: &str) -> Result<Box<dyn Worksheet + '_>> {
        let sheet = self
            .sheets
            .iter()
            .find(|s| s.name == name)
            .ok_or_else(|| format!("worksheet '{name}' not found"))?;
        Ok(Box::new(SheetRef { inner: sheet }))
    }

    fn worksheet_by_index(&self, index: usize) -> Result<Box<dyn Worksheet + '_>> {
        let sheet = self
            .sheets
            .get(index)
            .ok_or_else(|| format!("worksheet index {index} out of range"))?;
        Ok(Box::new(SheetRef { inner: sheet }))
    }

    fn worksheets(&self) -> Box<dyn WorksheetIterator<'_> + '_> {
        Box::new(EmptyWorksheetIter)
    }

    fn worksheet_count(&self) -> usize {
        self.sheets.len()
    }

    fn active_sheet_index(&self) -> usize {
        0
    }
}

/// Borrowed wrapper so we can return `Box<dyn Worksheet>` without consuming
/// the underlying `MemSheet`.
struct SheetRef<'a> {
    inner: &'a MemSheet,
}

impl<'a> Worksheet for SheetRef<'a> {
    fn name(&self) -> &str {
        self.inner.name()
    }

    fn row_count(&self) -> usize {
        self.inner.row_count()
    }

    fn column_count(&self) -> usize {
        self.inner.column_count()
    }

    fn dimensions(&self) -> Option<(u32, u32, u32, u32)> {
        self.inner.dimensions()
    }

    fn cell(&self, row: u32, column: u32) -> Result<Box<dyn Cell + '_>> {
        self.inner.cell(row, column)
    }

    fn cell_by_coordinate(&self, coordinate: &str) -> Result<Box<dyn Cell + '_>> {
        self.inner.cell_by_coordinate(coordinate)
    }

    fn cells(&self) -> Box<dyn CellIterator<'_> + '_> {
        self.inner.cells()
    }

    fn rows(&self) -> Box<dyn RowIterator<'_> + '_> {
        self.inner.rows()
    }

    fn row(&self, row_idx: usize) -> Result<Cow<'_, [CellValue]>> {
        self.inner.row(row_idx)
    }

    fn cell_value(&self, row: u32, column: u32) -> Result<Cow<'_, CellValue>> {
        self.inner.cell_value(row, column)
    }
}

// ---------------------------------------------------------------------------
// Demo
// ---------------------------------------------------------------------------

#[tokio::main]
async fn main() -> std::result::Result<(), Box<dyn std::error::Error + Send + Sync>> {
    // Build a workbook with a single sheet:
    //   A1 = 10 (Int)
    //   A2 = 20 (Int)
    //   A3 = =SUM(A1:A2)  with cached value 30
    //   B1 = "hello"     (string)
    //   B2 = 45292.5     (DateTime serial: 2023-12-31 12:00)
    let mut sheet = MemSheet::new("Sheet1");
    sheet.set(1, 1, CellValue::Int(10));
    sheet.set(2, 1, CellValue::Int(20));
    sheet.set(
        3,
        1,
        CellValue::Formula {
            formula: "SUM(A1:A2)".to_string(),
            cached_value: Some(Box::new(CellValue::Int(30))),
            is_array: false,
            array_range: None,
        },
    );
    sheet.set(1, 2, CellValue::String("hello".to_string()));
    sheet.set(2, 2, CellValue::DateTime(45292.5));

    // A formula referring to a named range we will register on the evaluator.
    sheet.set(
        4,
        1,
        CellValue::Formula {
            formula: "SUM(MyRange)".to_string(),
            // No cached value: forces actual evaluation through `define_name`.
            cached_value: None,
            is_array: false,
            array_range: None,
        },
    );

    let mut workbook = MemWorkbook::new();
    workbook.add_sheet(sheet);

    // Build the evaluator and register a workbook-scoped name.
    let mut evaluator = FormulaEvaluator::new(&workbook);
    evaluator.define_name("MyRange", "Sheet1!A1:A2");

    // Evaluate a few specific cells.
    println!("== Single cell evaluations ==");
    let a1 = evaluator.evaluate_cell("Sheet1", 1, 1).await?;
    println!("A1 (literal int)     = {:?}", a1);

    let a3 = evaluator.evaluate_cell("Sheet1", 3, 1).await?;
    println!("A3 (=SUM(A1:A2))     = {:?}", a3);

    let a4 = evaluator.evaluate_cell("Sheet1", 4, 1).await?;
    println!("A4 (=SUM(MyRange))   = {:?}", a4);

    let b1 = evaluator.evaluate_cell("Sheet1", 1, 2).await?;
    println!("B1 (string)          = {:?}", b1);

    let b2 = evaluator.evaluate_cell("Sheet1", 2, 2).await?;
    println!("B2 (datetime serial) = {:?}", b2);

    // Evaluate the entire sheet as a dense grid.
    println!("\n== Full-sheet evaluation ==");
    let grid = evaluator.evaluate_sheet("Sheet1").await?;
    for (i, row) in grid.iter().enumerate() {
        println!("row {}: {:?}", i + 1, row);
    }

    Ok(())
}
