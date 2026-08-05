//! Transactional XLSX worksheet data with typed SpreadsheetML table models.
//!
//! The workbook facade currently exposes table inspection and the typed table
//! XML boundary, while table-part attachment is not yet a worksheet edit
//! operation. The example therefore keeps worksheet mutations transactional
//! and validates both table models through serialize/read-back.

use litchi_xlsx::{
    Number, Table, TableStyleInfo, TableType, TotalsRowFunction, Value, Workbook, parse_table_xml,
    serialize_table,
};

type ExampleResult<T> = Result<T, Box<dyn std::error::Error + Send + Sync>>;

fn round_trip_table(table: Table) -> ExampleResult<Table> {
    let xml = serialize_table(&table)?;
    parse_table_xml(xml.as_bytes())?.ok_or_else(|| "table XML has no table root".into())
}

fn sales_table(headers: &[&str]) -> Table {
    let mut table = Table::new(1, "SalesTable", "A1:E5");
    table.table_type = Some(TableType::Worksheet);
    table.header_row_count = Some(1);
    table.initialize_columns();
    for (column, header) in table.columns.iter_mut().zip(headers) {
        column.name = (*header).to_owned();
    }

    table.style_info = Some(TableStyleInfo {
        name: Some("TableStyleMedium2".to_owned()),
        show_first_column: Some(false),
        show_last_column: Some(false),
        show_row_stripes: Some(true),
        show_column_stripes: Some(false),
    });
    table
}

fn summary_table(headers: &[&str]) -> Table {
    let mut table = Table::new(2, "SummaryTable", "A1:C5");
    table.header_row_count = Some(1);
    table.totals_row_count = Some(1);
    table.totals_row_shown = Some(true);
    table.initialize_columns();
    for (index, (column, header)) in table.columns.iter_mut().zip(headers).enumerate() {
        column.name = (*header).to_owned();
        column.totals_row_function = match index {
            1 => Some(TotalsRowFunction::Sum),
            2 => Some(TotalsRowFunction::Average),
            _ => None,
        };
    }
    table.style_info = Some(TableStyleInfo {
        name: Some("TableStyleLight9".to_owned()),
        show_row_stripes: Some(true),
        ..TableStyleInfo::default()
    });
    table
}

fn main() -> ExampleResult<()> {
    println!("Creating XLSX file with typed tables...");

    let workbook = Workbook::create()?;
    let mut edit = workbook.edit()?;
    edit.tab(0)?
        .ok_or("missing initial worksheet tab")?
        .rename("Sales Data")?;

    {
        let mut sheet = edit
            .sheet("Sales Data")?
            .ok_or("missing Sales Data worksheet")?;
        let headers = ["Product", "Q1", "Q2", "Q3", "Q4"];
        for (column, header) in headers.iter().enumerate() {
            sheet.set((1_u32, (column + 1) as u32), *header)?;
        }

        let data = [
            ("Apples", 1000, 1200, 1100, 1300),
            ("Oranges", 800, 900, 850, 950),
            ("Bananas", 1500, 1600, 1550, 1700),
            ("Grapes", 600, 700, 650, 750),
        ];
        for (row_index, (product, q1, q2, q3, q4)) in data.iter().enumerate() {
            let row = (row_index + 2) as u32;
            sheet.set((row, 1), *product)?;
            for (column, value) in [(2, q1), (3, q2), (4, q3), (5, q4)] {
                sheet.set(
                    (row, column),
                    Value::Number(Number::new(value.to_string())?),
                )?;
            }
        }
    }

    let mut summary = edit.add("Summary")?;
    let summary_headers = ["Category", "Total", "Average"];
    for (column, header) in summary_headers.iter().enumerate() {
        summary.set((1_u32, (column + 1) as u32), *header)?;
    }
    let summary_data = [
        ("Sales", 15000, 3750),
        ("Costs", 8000, 2000),
        ("Profit", 7000, 1750),
    ];
    for (row_index, (category, total, average)) in summary_data.iter().enumerate() {
        let row = (row_index + 2) as u32;
        summary.set((row, 1), *category)?;
        summary.set((row, 2), Value::Number(Number::new(total.to_string())?))?;
        summary.set((row, 3), Value::Number(Number::new(average.to_string())?))?;
    }

    let sales = round_trip_table(sales_table(&["Product", "Q1", "Q2", "Q3", "Q4"]))?;
    let summary_model = round_trip_table(summary_table(&summary_headers))?;
    let workbook = edit.commit()?.into_workbook();

    let output_path = std::env::args()
        .nth(1)
        .unwrap_or_else(|| "/tmp/tables_demo.xlsx".to_owned());
    workbook.save(&output_path)?;

    println!("Writing XLSX tables demo to {output_path}");
    println!(
        "Validated typed table models: {} ({}), {} ({})",
        sales.display_name, sales.ref_range, summary_model.display_name, summary_model.ref_range
    );
    println!("Tables created successfully ✔️");
    Ok(())
}
