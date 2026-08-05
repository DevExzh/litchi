//! Transactional XLSX chart-data and typed chart-model example.
//!
//! The standalone XLSX facade currently publishes worksheet cells through
//! `Workbook::edit`; chart package attachment is still a separate adapter
//! boundary. This example therefore writes the source data and validates the
//! four typed chart placements without inventing a worksheet mutation API.

use litchi_xlsx::chart::{Anchor, Chart};
use litchi_xlsx::{Number, Workbook};
use std::env;
use std::error::Error;

fn main() -> Result<(), Box<dyn Error + Send + Sync>> {
    let args: Vec<String> = env::args().collect();
    let output_path = args.get(1).map_or("xlsx_charts_demo.xlsx", String::as_str);

    let workbook = Workbook::create()?;
    let mut edit = workbook.edit()?;
    edit.tab(0)?
        .ok_or("default worksheet is missing")?
        .rename("Charts")?;
    {
        let mut sheet = edit.sheet("Charts")?.ok_or("Charts worksheet is missing")?;
        for (column, heading) in [
            ("A1", "Month"),
            ("B1", "Revenue"),
            ("C1", "Units"),
            ("D1", "Temperature (°C)"),
            ("E1", "Foot Traffic"),
        ] {
            sheet.set(column, heading)?;
        }

        let months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun"];
        let revenue = [320.0, 410.0, 380.0, 450.0, 520.0, 610.0];
        let units = [120.0, 150.0, 140.0, 170.0, 190.0, 210.0];
        let temperature = [4.0, 6.0, 10.0, 15.0, 20.0, 24.0];
        let foot_traffic = [200.0, 240.0, 260.0, 300.0, 330.0, 360.0];
        for (index, month) in months.iter().enumerate() {
            let row = (index + 2) as u32;
            sheet.set((row, 0), *month)?;
            sheet.set((row, 1), Number::new(revenue[index].to_string())?)?;
            sheet.set((row, 2), Number::new(units[index].to_string())?)?;
            sheet.set((row, 3), Number::new(temperature[index].to_string())?)?;
            sheet.set((row, 4), Number::new(foot_traffic[index].to_string())?)?;
        }
        for (column, width) in [
            ("A", 12.0),
            ("B", 12.0),
            ("C", 12.0),
            ("D", 16.0),
            ("E", 16.0),
        ] {
            sheet.column(column)?.width(width)?;
        }
    }

    let charts = [
        Chart::bar_chart(
            "Monthly Revenue",
            "Charts!$A$2:$A$7",
            "Charts!$B$2:$B$7",
            Anchor::new(0, 8, 7, 20),
        )?,
        Chart::line_chart(
            "Units Sold",
            "Charts!$A$2:$A$7",
            "Charts!$C$2:$C$7",
            Anchor::new(8, 0, 15, 10),
        )?,
        Chart::pie_chart(
            "Revenue Share",
            "Charts!$A$2:$A$7",
            "Charts!$B$2:$B$7",
            Anchor::new(8, 10, 15, 20),
        )?,
        Chart::scatter_chart(
            "Temperature vs Traffic",
            "Charts!$D$2:$D$7",
            "Charts!$E$2:$E$7",
            Anchor::new(0, 20, 7, 32),
        )?,
    ];
    let workbook = edit.commit()?.into_workbook();
    workbook.save(output_path)?;
    println!(
        "Saved {} chart models and source data to {output_path}",
        charts.len()
    );
    Ok(())
}
