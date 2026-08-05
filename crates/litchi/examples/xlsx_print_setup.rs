//! Transactional XLSX print-data and typed page-setup example.
//!
//! Print-area and repeating-title package edits are not exposed by the
//! standalone worksheet transaction yet. The page settings below remain typed
//! and validated, while the supported worksheet data is committed atomically.

use litchi_xlsx::{Fit, Orientation, Paper, Setup, Workbook};
use std::env;

fn main() -> Result<(), Box<dyn std::error::Error + Send + Sync>> {
    let output_path = env::args()
        .nth(1)
        .unwrap_or_else(|| "xlsx_print_setup.xlsx".to_string());

    let page_setup = Setup {
        orientation: Some(Orientation::Landscape),
        paper: Some(Paper::A4),
        fit_to_width: Some(Fit::ONE),
        fit_to_height: Some(Fit::ONE),
        ..Setup::default()
    };
    let print_area = "A1:D80";
    let repeating_rows = "$1:$1";
    let repeating_columns = "$A:$A";

    let workbook = Workbook::create()?;
    let mut edit = workbook.edit()?;
    edit.tab(0)?
        .ok_or("default worksheet is missing")?
        .rename("PrintDemo")?;
    {
        let mut sheet = edit
            .sheet("PrintDemo")?
            .ok_or("PrintDemo worksheet is missing")?;
        for (cell, value) in [
            ("A1", "ID"),
            ("B1", "Name"),
            ("C1", "Department"),
            ("D1", "Value"),
        ] {
            sheet.set(cell, value)?;
        }
        for row in 2..=80 {
            sheet.set((row, 0), (row - 1) as i32)?;
            sheet.set((row, 1), format!("Employee {row}"))?;
            sheet.set((row, 2), if row % 2 == 0 { "Engineering" } else { "Sales" })?;
            sheet.set((row, 3), ((row as i32) - 1) * 10)?;
        }
    }

    let workbook = edit.commit()?.into_workbook();
    workbook.save(&output_path)?;
    println!("Saved {output_path}");
    println!("Typed page setup: {page_setup:?}; print area: {print_area}");
    println!("Repeating rows: {repeating_rows}; columns: {repeating_columns}");
    Ok(())
}
