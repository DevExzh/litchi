//! Simple text formats demonstration
//!
//! Basic example showing how to generate different text format files with
//! `litchi::sheet::text` (the `litchi-sheet-text` crate).

use std::fs::File;

use litchi::sheet::CellValue;
use litchi::sheet::text::{delimited, dif, fixed_width, sylk};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("=== Simple Text Formats Demo ===\n");

    // Create sample data
    let data = vec![
        vec![
            CellValue::String("Name".to_string()),
            CellValue::String("Score".to_string()),
        ],
        vec![CellValue::String("张三".to_string()), CellValue::Int(95)],
        vec![CellValue::String("李四".to_string()), CellValue::Int(87)],
    ];

    delimited::write(
        &data,
        File::create("simple.csv")?,
        &delimited::Options::csv(),
    )?;
    println!("✓ Generated simple.csv");

    delimited::write(
        &data,
        File::create("simple.tsv")?,
        &delimited::Options::tsv(),
    )?;
    println!("✓ Generated simple.tsv");

    sylk::write(
        &data,
        File::create("simple.slk")?,
        &sylk::Options::default(),
    )?;
    println!("✓ Generated simple.slk");

    dif::write(&data, File::create("simple.dif")?, &dif::Options::default())?;
    println!("✓ Generated simple.dif");

    let widths = fixed_width::Options::default().column_widths(vec![10, 8]);
    fixed_width::write(&data, File::create("simple.prn")?, &widths)?;
    println!("✓ Generated simple.prn");

    println!("\n=== All files generated successfully! ===");
    Ok(())
}
