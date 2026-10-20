//! Structured-table reference example for `litchi-eval`.
//!
//! Demonstrates [`FormulaEvaluator::define_table`] together with a
//! [`TableConfig`] so structured references (e.g. `Sales[Qty]`) resolve
//! against a region of an in-memory workbook.
//!
//! Run with:
//!
//! ```sh
//! cargo run -p litchi-eval --example evaluate_with_table --all-features
//! ```

use std::borrow::Cow;
use std::collections::HashMap;

use litchi_core::edit::RefusalKind;
use litchi_core::sheet::{Cell, CellValue, Error as SheetError, Result, WorkbookTrait, Worksheet};

/// The refusal for a worksheet the workbook does not have.
const NO_SUCH_SHEET: SheetError =
    SheetError::refused(RefusalKind::NotFound, "worksheet", "no such worksheet");
use litchi_eval::{FormulaEvaluator, TableConfig};

// ---------------------------------------------------------------------------
// Minimal in-memory workbook (same shape as the simple example)
// ---------------------------------------------------------------------------

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
        format!("R{}C{}", self.row, self.column)
    }

    fn value(&self) -> &CellValue {
        self.value
    }

    fn is_formula(&self) -> bool {
        matches!(self.value, CellValue::Formula { .. })
    }
}

// --- Worksheet impl --------------------------------------------------------
//
// The traits dispatch statically: the worksheet names its cell and iterator
// types. This minimal sheet iterates nothing.

impl Worksheet for MemSheet {
    type Error = SheetError;
    type Cell<'a> = MemCell<'a>;
    type Cells<'a> = std::iter::Empty<Result<MemCell<'a>>>;
    type Rows<'a> = std::iter::Empty<Result<Cow<'a, [CellValue]>>>;

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

    fn cell(&self, row: u32, column: u32) -> Result<MemCell<'_>> {
        let value = self.cells.get(&(row, column)).unwrap_or(CellValue::EMPTY);
        Ok(MemCell { row, column, value })
    }

    fn cell_by_coordinate(&self, _coordinate: &str) -> Result<MemCell<'_>> {
        Err(SheetError::refused(
            RefusalKind::UnsupportedOperation,
            "cell_by_coordinate",
            "MemSheet reads cells by row and column",
        ))
    }

    fn cells(&self) -> Self::Cells<'_> {
        std::iter::empty()
    }

    fn rows(&self) -> Self::Rows<'_> {
        std::iter::empty()
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
//
// Worksheets are lent by reference: `&MemSheet` is a `Worksheet` through the
// blanket implementation for references.

impl WorkbookTrait for MemWorkbook {
    type Error = SheetError;
    type Worksheet<'a> = &'a MemSheet;
    type Worksheets<'a> =
        std::iter::Map<std::slice::Iter<'a, MemSheet>, fn(&'a MemSheet) -> Result<&'a MemSheet>>;

    fn active_worksheet(&self) -> Result<&MemSheet> {
        self.worksheet_by_index(0)
    }

    fn worksheet_names(&self) -> &[String] {
        &self.sheet_names
    }

    fn worksheet_by_name(&self, name: &str) -> Result<&MemSheet> {
        self.sheets
            .iter()
            .find(|s| s.name == name)
            .ok_or(NO_SUCH_SHEET)
    }

    fn worksheet_by_index(&self, index: usize) -> Result<&MemSheet> {
        self.sheets.get(index).ok_or(NO_SUCH_SHEET)
    }

    fn worksheets(&self) -> Self::Worksheets<'_> {
        self.sheets.iter().map(Ok)
    }

    fn worksheet_count(&self) -> usize {
        self.sheets.len()
    }

    fn active_sheet_index(&self) -> usize {
        0
    }
}

// ---------------------------------------------------------------------------
// Demo
// ---------------------------------------------------------------------------

#[tokio::main]
async fn main() -> std::result::Result<(), Box<dyn std::error::Error + Send + Sync>> {
    // Layout for "Sheet1":
    //
    //         A         B       C
    //   1   "Item"   "Qty"   "Price"      <- table header row
    //   2   "Apple"   3       1.50        <- data row 1
    //   3   "Bread"   1       4.20        <- data row 2
    //   5   =SUM(Sales[Qty])              <- structured-reference formula
    //   6   =SUM(Sales[Price])            <- structured-reference formula
    //
    // The table "Sales" spans A1:C3 (header row at row 1, data rows 2..=3).
    let mut sheet = MemSheet::new("Sheet1");

    // Header row (row 1)
    sheet.set(1, 1, CellValue::String("Item".to_string()));
    sheet.set(1, 2, CellValue::String("Qty".to_string()));
    sheet.set(1, 3, CellValue::String("Price".to_string()));

    // Data row 1 (row 2)
    sheet.set(2, 1, CellValue::String("Apple".to_string()));
    sheet.set(2, 2, CellValue::Int(3));
    sheet.set(2, 3, CellValue::Float(1.50));

    // Data row 2 (row 3)
    sheet.set(3, 1, CellValue::String("Bread".to_string()));
    sheet.set(3, 2, CellValue::Int(1));
    sheet.set(3, 3, CellValue::Float(4.20));

    // Formula cells using structured references.
    sheet.set(
        5,
        1,
        CellValue::Formula {
            formula: "SUM(Sales[Qty])".to_string(),
            cached_value: None,
            is_array: false,
            array_range: None,
        },
    );
    sheet.set(
        6,
        1,
        CellValue::Formula {
            formula: "SUM(Sales[Price])".to_string(),
            cached_value: None,
            is_array: false,
            array_range: None,
        },
    );

    let mut workbook = MemWorkbook::new();
    workbook.add_sheet(sheet);

    // Build the evaluator and register the "Sales" table on it.
    let mut evaluator = FormulaEvaluator::new(&workbook);
    let headers = vec!["Item".to_string(), "Qty".to_string(), "Price".to_string()];
    evaluator.define_table(TableConfig {
        name: "Sales",
        sheet_name: "Sheet1",
        start_row: 1,
        start_col: 1,
        end_row: 3,
        end_col: 3,
        headers: &headers,
    });

    println!("== Structured-reference evaluations ==");

    // Evaluate the SUM over the [Qty] column.
    let qty_total = evaluator.evaluate_cell("Sheet1", 5, 1).await?;
    println!("A5  =SUM(Sales[Qty])   -> {:?}", qty_total);

    // Evaluate the SUM over the [Price] column.
    let price_total = evaluator.evaluate_cell("Sheet1", 6, 1).await?;
    println!("A6  =SUM(Sales[Price]) -> {:?}", price_total);

    // Inspect the raw header / data cells too.
    println!("\n== Raw table cells ==");
    for row in 1..=3 {
        let item = evaluator.evaluate_cell("Sheet1", row, 1).await?;
        let qty = evaluator.evaluate_cell("Sheet1", row, 2).await?;
        let price = evaluator.evaluate_cell("Sheet1", row, 3).await?;
        println!("row {row}: {:?} | {:?} | {:?}", item, qty, price);
    }

    Ok(())
}
