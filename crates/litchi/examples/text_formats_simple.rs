//! Simple text formats demonstration
//!
//! Basic example showing how to generate different text format files.

use litchi::sheet::CellValue;
use litchi::sheet::text::formats::*;
use std::fs::File;

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

    // Generate CSV
    {
        let file = File::create("simple.csv")?;
        let mut writer = file;
        write_delimited(&data, &mut writer, DelimitedConfig::csv())
            .map_err(|e| -> Box<dyn std::error::Error> { e })?;
        println!("✓ Generated simple.csv");
    }

    // Generate TSV
    {
        let file = File::create("simple.tsv")?;
        let mut writer = file;
        write_delimited(&data, &mut writer, DelimitedConfig::tsv())
            .map_err(|e| -> Box<dyn std::error::Error> { e })?;
        println!("✓ Generated simple.tsv");
    }

    // Generate SYLK
    {
        let file = File::create("simple.slk")?;
        let mut writer = file;
        write_sylk(&data, &mut writer, SylkConfig::default())
            .map_err(|e| -> Box<dyn std::error::Error> { e })?;
        println!("✓ Generated simple.slk");
    }

    // Generate DIF
    {
        let file = File::create("simple.dif")?;
        let mut writer = file;
        write_dif(&data, &mut writer, DifConfig::default())
            .map_err(|e| -> Box<dyn std::error::Error> { e })?;
        println!("✓ Generated simple.dif");
    }

    // Generate fixed-width PRN
    {
        let file = File::create("simple.prn")?;
        let mut writer = file;
        let config = FixedWidthConfig {
            column_widths: vec![10, 8],
            auto_detect_widths: false,
            trim_fields: true,
            strip_bom: true,
            write_bom: None,
        };
        write_fixed_width(&data, &mut writer, config)
            .map_err(|e| -> Box<dyn std::error::Error> { e })?;
        println!("✓ Generated simple.prn");
    }

    println!("\n=== All files generated successfully! ===");
    Ok(())
}
