//! Typed rich shared-string model with transactional XLSX data.
//!
//! Rich-string cell attachment is currently owned by the package adapter, so
//! this example validates the standalone shared-string model and publishes the
//! surrounding worksheet through `Workbook::edit`.

use litchi_xlsx::Workbook;
use litchi_xlsx::raw::shared_strings::{Run, Table};
use std::env;

fn main() -> Result<(), Box<dyn std::error::Error + Send + Sync>> {
    let output = env::args()
        .nth(1)
        .unwrap_or_else(|| "xlsx_rich_text.xlsx".to_string());

    let mut strings = Table::new();
    let rich_index = strings.intern_rich(vec![
        Run {
            text: "Hello, ".to_string(),
            font_name: Some("Calibri".to_string()),
            font_size: Some(11.0),
            bold: false,
            italic: false,
            underline: false,
            color: None,
        },
        Run {
            text: "bold".to_string(),
            font_name: Some("Calibri".to_string()),
            font_size: Some(11.0),
            bold: true,
            color: Some("FFFF0000".to_string()),
            italic: false,
            underline: false,
        },
        Run {
            text: " and ".to_string(),
            font_name: Some("Calibri".to_string()),
            font_size: Some(11.0),
            bold: false,
            italic: false,
            underline: false,
            color: None,
        },
        Run {
            text: "underlined".to_string(),
            font_name: Some("Calibri".to_string()),
            font_size: Some(11.0),
            underline: true,
            color: Some("FF0000FF".to_string()),
            bold: false,
            italic: false,
        },
    ])?;
    let _shared_strings_xml = strings.write_xml()?;

    let workbook = Workbook::create()?;
    let mut edit = workbook.edit()?;
    edit.tab(0)?
        .ok_or("default worksheet is missing")?
        .rename("RichText")?;
    {
        let mut sheet = edit
            .sheet("RichText")?
            .ok_or("RichText worksheet is missing")?;
        sheet.set("A1", "Rich Text Demo")?;
        sheet.set("A3", "Hello, bold and underlined")?;
    }
    let workbook = edit.commit()?.into_workbook();
    workbook.save(&output)?;
    println!("Saved {output}; validated rich shared-string index {rich_index}");
    Ok(())
}
