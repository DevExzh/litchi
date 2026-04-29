//! XLSX rich text cell example
//!
//! Run with:
//!
//! ```bash
//! cargo run --example xlsx_rich_text --features ooxml -- xlsx_rich_text.xlsx
//! ```

use litchi::ooxml::xlsx::{RichTextRun, Workbook};
use std::env;

fn main() -> Result<(), Box<dyn std::error::Error + Send + Sync>> {
    let args: Vec<String> = env::args().collect();
    let output = if args.len() > 1 {
        &args[1]
    } else {
        "xlsx_rich_text.xlsx"
    };

    let mut wb = Workbook::create()?;
    {
        let ws = wb.worksheet_mut(0)?;
        ws.set_name("RichText".to_string());

        // Simple title using plain value
        ws.set_cell_value(1, 1, "Rich Text Demo");

        // Cell A3: mixed formatting in one cell
        ws.set_rich_text_cell(
            3,
            1,
            vec![
                RichTextRun {
                    text: "Hello, ".to_string(),
                    font_name: Some("Calibri".to_string()),
                    font_size: Some(11.0),
                    bold: false,
                    italic: false,
                    underline: false,
                    color: None,
                },
                RichTextRun {
                    text: "bold".to_string(),
                    font_name: Some("Calibri".to_string()),
                    font_size: Some(11.0),
                    bold: true,
                    italic: false,
                    underline: false,
                    color: Some("FFFF0000".to_string()), // red
                },
                RichTextRun {
                    text: " and ".to_string(),
                    font_name: Some("Calibri".to_string()),
                    font_size: Some(11.0),
                    bold: false,
                    italic: false,
                    underline: false,
                    color: None,
                },
                RichTextRun {
                    text: "underlined".to_string(),
                    font_name: Some("Calibri".to_string()),
                    font_size: Some(11.0),
                    bold: false,
                    italic: false,
                    underline: true,
                    color: Some("FF0000FF".to_string()), // blue
                },
            ],
        );
    }

    wb.save(output)?;
    println!("Saved rich text example to: {}", output);
    println!("Open it in Excel and confirm cell A3 has mixed formatting in a single cell.");

    Ok(())
}
