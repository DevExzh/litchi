//! Transactional worksheet data plus a typed page-view model example.
//!
//! Page-break package attachment remains outside the standalone transaction.
//! The shared view value is still constructed and validated while source tables
//! are published through `Workbook::edit`.

use litchi::sheet::view::{Mode, Scale, View};
use litchi::xlsx::{Number, Workbook};
use std::env;

fn main() -> Result<(), Box<dyn std::error::Error + Send + Sync>> {
    let output_file = env::args()
        .nth(1)
        .unwrap_or_else(|| "page-breaks.xlsx".to_string());

    let mut page_view = View::default();
    page_view.mode = Mode::PageBreakPreview;
    page_view.zoom.current = Scale::new(85)?;
    let page_breaks = [("row", 21_u32), ("row", 41), ("column", 3)];

    let workbook = Workbook::create()?;
    let mut edit = workbook.edit()?;
    edit.tab(0)?
        .ok_or("default worksheet is missing")?
        .rename("Page Layout Preview")?;
    {
        let mut sheet = edit
            .sheet("Page Layout Preview")?
            .ok_or("page-layout worksheet is missing")?;
        for (cell, value) in [("A1", "Quarter"), ("B1", "Region"), ("C1", "Revenue")] {
            sheet.set(cell, value)?;
        }
        for index in 0..60 {
            let row = (index + 2) as u32;
            sheet.set((row, 0), format!("Q{}", (index % 4) + 1))?;
            sheet.set((row, 1), format!("Region {}", (index % 5) + 1))?;
            let revenue = Number::new(format!("{}", (index + 1) * 1_000))?;
            sheet.set((row, 2), revenue)?;
        }
    }
    let mut sorted = edit.add("Sorted Sales")?;
    for (cell, value) in [("A1", "Product"), ("B1", "Category"), ("C1", "Units")] {
        sorted.set(cell, value)?;
    }
    for (index, (product, category, units)) in [
        ("Laptop", "Hardware", 120.0),
        ("Mouse", "Hardware", 950.0),
        ("Keyboard", "Hardware", 500.0),
        ("Monitor", "Hardware", 320.0),
        ("Chair", "Office", 85.0),
        ("Desk", "Office", 45.0),
        ("Pens", "Office", 800.0),
        ("Notebook", "Office", 650.0),
        ("Dock", "Hardware", 260.0),
        ("Headset", "Hardware", 410.0),
    ]
    .iter()
    .enumerate()
    {
        let row = (index + 2) as u32;
        sorted.set((row, 0), *product)?;
        sorted.set((row, 1), *category)?;
        sorted.set((row, 2), Number::new(format!("{units}"))?)?;
    }

    let workbook = edit.commit()?.into_workbook();
    workbook.save(&output_file)?;
    println!("Saved {output_file}");
    println!("Typed view: {page_view:?}; breaks: {page_breaks:?}");
    Ok(())
}
