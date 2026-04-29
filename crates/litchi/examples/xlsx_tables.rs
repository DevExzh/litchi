use litchi::ooxml::xlsx::{Table, TableStyleInfo, TableType, TotalsRowFunction, Workbook};

fn main() -> Result<(), Box<dyn std::error::Error + Send + Sync>> {
    println!("Creating XLSX file with tables...");

    let mut wb = Workbook::create()?;
    let ws = wb.worksheet_mut(0)?;
    ws.set_name("Sales Data".to_string());

    // Add sample data
    let headers = ["Product", "Q1", "Q2", "Q3", "Q4"];
    for (col, header) in headers.iter().enumerate() {
        ws.set_cell_value(1, (col + 1) as u32, *header);
    }

    let data = [
        ("Apples", 1000, 1200, 1100, 1300),
        ("Oranges", 800, 900, 850, 950),
        ("Bananas", 1500, 1600, 1550, 1700),
        ("Grapes", 600, 700, 650, 750),
    ];

    for (row_idx, (product, q1, q2, q3, q4)) in data.iter().enumerate() {
        let row = (row_idx + 2) as u32;
        ws.set_cell_value(row, 1, *product);
        ws.set_cell_value(row, 2, *q1);
        ws.set_cell_value(row, 3, *q2);
        ws.set_cell_value(row, 4, *q3);
        ws.set_cell_value(row, 5, *q4);
    }

    // Create a table
    let mut table = Table::new(1, "SalesTable", "A1:E5");
    table.display_name = "SalesTable".to_string();
    table.table_type = Some(TableType::Worksheet);
    table.header_row_count = Some(1);

    // Initialize columns with names from headers
    table.initialize_columns();
    for (i, header) in headers.iter().enumerate() {
        if let Some(col) = table.columns.get_mut(i) {
            col.name = header.to_string();
        }
    }

    // Add table style
    let mut style_info = TableStyleInfo::new();
    style_info.name = Some("TableStyleMedium2".to_string());
    style_info.show_first_column = Some(false);
    style_info.show_last_column = Some(false);
    style_info.show_row_stripes = Some(true);
    style_info.show_column_stripes = Some(false);
    table.style_info = Some(style_info);

    ws.add_table(table);

    // Create a second table with totals row on another sheet
    let ws2 = wb.add_worksheet("Summary");

    let summary_headers = ["Category", "Total", "Average"];
    for (col, header) in summary_headers.iter().enumerate() {
        ws2.set_cell_value(1, (col + 1) as u32, *header);
    }

    let summary_data = [
        ("Sales", 15000, 3750),
        ("Costs", 8000, 2000),
        ("Profit", 7000, 1750),
    ];

    for (row_idx, (category, total, avg)) in summary_data.iter().enumerate() {
        let row = (row_idx + 2) as u32;
        ws2.set_cell_value(row, 1, *category);
        ws2.set_cell_value(row, 2, *total);
        ws2.set_cell_value(row, 3, *avg);
    }

    // Create table with totals row
    let mut summary_table = Table::new(2, "SummaryTable", "A1:C5");
    summary_table.display_name = "SummaryTable".to_string();
    summary_table.header_row_count = Some(1);
    summary_table.totals_row_count = Some(1);
    summary_table.totals_row_shown = Some(true);

    summary_table.initialize_columns();
    for (i, header) in summary_headers.iter().enumerate() {
        if let Some(col) = summary_table.columns.get_mut(i) {
            col.name = header.to_string();

            // Add totals row functions
            if i == 1 {
                col.totals_row_function = Some(TotalsRowFunction::Sum);
            } else if i == 2 {
                col.totals_row_function = Some(TotalsRowFunction::Average);
            }
        }
    }

    // Add different style
    let mut summary_style = TableStyleInfo::new();
    summary_style.name = Some("TableStyleLight9".to_string());
    summary_style.show_row_stripes = Some(true);
    summary_table.style_info = Some(summary_style);

    ws2.add_table(summary_table);

    let output_path = std::env::args()
        .nth(1)
        .unwrap_or_else(|| "/tmp/tables_demo.xlsx".to_string());

    println!("Writing XLSX tables demo to {}", output_path);
    wb.save(&output_path)?;
    println!("Tables created successfully ✔️");
    println!("\nTables created:");
    println!("  - Sheet 'Sales Data': SalesTable (A1:E5) with medium style");
    println!("  - Sheet 'Summary': SummaryTable (A1:C5) with light style and totals row");

    Ok(())
}
