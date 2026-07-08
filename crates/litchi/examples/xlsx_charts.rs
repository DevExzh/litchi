//! Demonstrates generating XLSX files with multiple chart types.
//!
//! Run with:
//! ```bash
//! cargo run --example xlsx_charts -- /tmp/xlsx_charts_demo.xlsx
//! ```
//! Then open the resulting file in Excel to verify the charts.

use litchi::ooxml::xlsx::{ChartAnchor, Workbook, WorksheetChart};
use std::env;
use std::error::Error;

fn main() -> Result<(), Box<dyn Error + Send + Sync>> {
    let args: Vec<String> = env::args().collect();
    let output_path = args.get(1).map_or("xlsx_charts_demo.xlsx", String::as_str);

    println!("Generating XLSX chart demo at {output_path}");

    let mut workbook = Workbook::create()?;
    let sheet = workbook.worksheet_mut(0)?;
    sheet.set_name("Charts".to_string());

    // Header row (API is 1-based)
    sheet.set_cell_value(1, 1, "Month");
    sheet.set_cell_value(1, 2, "Revenue");
    sheet.set_cell_value(1, 3, "Units");
    sheet.set_cell_value(1, 4, "Temperature (°C)");
    sheet.set_cell_value(1, 5, "Foot Traffic");

    let months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun"];
    let revenue = [320.0, 410.0, 380.0, 450.0, 520.0, 610.0];
    let units = [120.0, 150.0, 140.0, 170.0, 190.0, 210.0];
    let temperature = [4.0, 6.0, 10.0, 15.0, 20.0, 24.0];
    let foot_traffic = [200.0, 240.0, 260.0, 300.0, 330.0, 360.0];

    for i in 0..months.len() {
        let row = i as u32 + 2; // Data starts on row 2
        sheet.set_cell_value(row, 1, months[i]);
        sheet.set_cell_value(row, 2, revenue[i]);
        sheet.set_cell_value(row, 3, units[i]);
        sheet.set_cell_value(row, 4, temperature[i]);
        sheet.set_cell_value(row, 5, foot_traffic[i]);
    }

    // Column widths for readability
    sheet.set_column_width(1, 12.0);
    sheet.set_column_width(2, 12.0);
    sheet.set_column_width(3, 12.0);
    sheet.set_column_width(4, 16.0);
    sheet.set_column_width(5, 16.0);

    // Bar chart (Revenue per month)
    let bar_chart = WorksheetChart::bar_chart(
        "Monthly Revenue",
        "Charts!$A$2:$A$7",
        "Charts!$B$2:$B$7",
        ChartAnchor::new(0, 8, 7, 20),
    )?;
    sheet.add_chart(bar_chart);

    // Line chart (Units per month)
    let line_chart = WorksheetChart::line_chart(
        "Units Sold",
        "Charts!$A$2:$A$7",
        "Charts!$C$2:$C$7",
        ChartAnchor::new(8, 0, 15, 10),
    )?;
    sheet.add_chart(line_chart);

    // Pie chart (share of revenue)
    let pie_chart = WorksheetChart::pie_chart(
        "Revenue Share",
        "Charts!$A$2:$A$7",
        "Charts!$B$2:$B$7",
        ChartAnchor::new(8, 10, 15, 20),
    )?;
    sheet.add_chart(pie_chart);

    // Scatter chart (Temperature vs. Foot Traffic)
    let scatter_chart = WorksheetChart::scatter_chart(
        "Temperature vs Traffic",
        "Charts!$D$2:$D$7",
        "Charts!$E$2:$E$7",
        ChartAnchor::new(0, 20, 7, 32),
    )?;
    sheet.add_chart(scatter_chart);

    workbook.save(output_path)?;

    println!("Done! Open '{output_path}' in Excel to verify the charts.");
    Ok(())
}
