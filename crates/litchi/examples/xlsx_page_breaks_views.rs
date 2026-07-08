//! Demonstrates XLSX page breaks, worksheet view settings, and auto-filter sort
//! state serialization.
//!
//! Usage:
//! ```bash
//! cargo run --example xlsx_page_breaks_views -- /tmp/page-breaks.xlsx
//! ```
//! Open the generated workbook in Microsoft Excel to verify the features.

use litchi::ooxml::xlsx::{SheetView, SheetViewType, SortCondition, SortState, Workbook};
use std::env;

fn main() -> Result<(), Box<dyn std::error::Error + Send + Sync>> {
    let args: Vec<String> = env::args().collect();
    let output_file = args
        .get(1)
        .map(String::as_str)
        .unwrap_or("page-breaks.xlsx");

    println!("Writing XLSX page break + view + sort-state demo to {output_file}");

    let mut workbook = Workbook::create()?;

    build_page_layout_sheet(workbook.add_worksheet("Page Layout Preview"));
    build_sort_state_sheet(workbook.add_worksheet("Sorted Sales"));

    workbook.save(output_file)?;
    println!("Workbook saved ✔️");
    Ok(())
}

fn build_page_layout_sheet(sheet: &mut litchi::ooxml::xlsx::MutableWorksheet) {
    sheet.set_cell_value(1, 1, "Quarter");
    sheet.set_cell_value(1, 2, "Region");
    sheet.set_cell_value(1, 3, "Revenue");

    for idx in 0..60 {
        let row = idx + 2; // data starts on row 2 (1-based)
        sheet.set_cell_value(row, 1, format!("Q{}", (idx % 4) + 1));
        sheet.set_cell_value(row, 2, format!("Region {}", (idx % 5) + 1));
        sheet.set_cell_value(row, 3, ((idx + 1) as f64) * 1_000.0);
    }

    // Add manual page breaks so Excel prints each 20-row block as a page.
    sheet.add_row_break(21, 0, 2);
    sheet.add_row_break(41, 0, 2);
    sheet.add_column_break(3, 0, 60);

    // Set a page-layout sheet view with gridlines hidden.
    let view = SheetView {
        view_type: Some(SheetViewType::PageLayout),
        show_grid_lines: Some(false),
        zoom_scale: Some(120),
        top_left_cell: Some("A10".to_string()),
        ..Default::default()
    };
    sheet.set_sheet_view(view);
}

fn build_sort_state_sheet(sheet: &mut litchi::ooxml::xlsx::MutableWorksheet) {
    sheet.set_cell_value(1, 1, "Product");
    sheet.set_cell_value(1, 2, "Category");
    sheet.set_cell_value(1, 3, "Units");

    let rows = [
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
    ];

    for (idx, (product, category, units)) in rows.iter().enumerate() {
        let row = idx as u32 + 2; // data rows start at row 2
        sheet.set_cell_value(row, 1, *product);
        sheet.set_cell_value(row, 2, *category);
        sheet.set_cell_value(row, 3, *units);
    }

    let autofilter_range = "A1:C11";
    sheet.set_auto_filter(autofilter_range);

    let mut sort_state = SortState::new(autofilter_range);
    sort_state.column_sort = Some(true);
    sort_state.case_sensitive = Some(false);

    let mut primary = SortCondition::new("C2:C11");
    primary.descending = Some(true); // Highest unit count first
    sort_state.conditions.push(primary);

    let mut secondary = SortCondition::new("A2:A11");
    secondary.descending = Some(false); // Alphabetical ties
    sort_state.conditions.push(secondary);

    sheet.set_auto_filter_sort_state(sort_state);
}
