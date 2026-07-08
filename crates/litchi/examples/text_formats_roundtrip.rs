//! Round-trip test for text formats (read-write verification)
//!
//! This example tests that data remains consistent when written and read back
//! for all supported text formats.

use litchi::sheet::CellValue;
use litchi::sheet::text::formats::*;
use std::fs::File;
use std::io::{BufWriter, Write};

type ExampleResult<T> = Result<T, Box<dyn std::error::Error + Send + Sync>>;

fn main() -> ExampleResult<()> {
    println!("=== Text Formats Round-trip Test ===\n");

    // Create test data with various cell types
    let test_data = create_test_data();

    // Test round-trip for each format
    test_csv_roundtrip(&test_data)?;
    test_tsv_roundtrip(&test_data)?;
    test_prn_delimited_roundtrip(&test_data)?;
    test_prn_fixed_width_roundtrip(&test_data)?;
    test_sylk_roundtrip(&test_data)?;
    test_dif_roundtrip(&test_data)?;

    println!("\n=== All round-trip tests completed! ===");
    Ok(())
}

fn create_test_data() -> Vec<Vec<CellValue>> {
    vec![
        vec![
            CellValue::String("Test Data".to_string()),
            CellValue::String("Round-trip".to_string()),
            CellValue::String("Verification".to_string()),
        ],
        vec![
            CellValue::Empty,
            CellValue::Int(42),
            CellValue::Float(std::f64::consts::PI),
        ],
        vec![
            CellValue::Bool(true),
            CellValue::Bool(false),
            CellValue::String("Special chars: ,;\"'\n\t".to_string()),
        ],
        vec![
            CellValue::String("Unicode: 你好 🌍".to_string()),
            CellValue::Float(-123.456),
            CellValue::Int(0),
        ],
        vec![
            CellValue::Formula {
                formula: "A1+B1".to_string(),
                cached_value: Some(Box::new(CellValue::Float(45.14159))),
                is_array: false,
                array_range: None,
            },
            CellValue::Error("#N/A".to_string()),
            CellValue::String("End".to_string()),
        ],
    ]
}

fn compare_data(
    original: &[Vec<CellValue>],
    read_back: &[Vec<CellValue>],
    format_name: &str,
) -> bool {
    if original.len() != read_back.len() {
        println!(
            "  ❌ Row count mismatch: {} vs {}",
            original.len(),
            read_back.len()
        );
        return false;
    }

    for (row_idx, (orig_row, read_row)) in original.iter().zip(read_back.iter()).enumerate() {
        if orig_row.len() != read_row.len() {
            println!(
                "  ❌ Column count mismatch in row {}: {} vs {}",
                row_idx,
                orig_row.len(),
                read_row.len()
            );
            return false;
        }

        for (col_idx, (orig_cell, read_cell)) in orig_row.iter().zip(read_row.iter()).enumerate() {
            if !cells_equal(orig_cell, read_cell) {
                println!(
                    "  ❌ Cell mismatch at ({}, {}): {:?} vs {:?}",
                    row_idx, col_idx, orig_cell, read_cell
                );
                return false;
            }
        }
    }

    println!("  ✓ {} data matches perfectly", format_name);
    true
}

fn cells_equal(a: &CellValue, b: &CellValue) -> bool {
    match (a, b) {
        (CellValue::Empty, CellValue::Empty) => true,
        (CellValue::Bool(x), CellValue::Bool(y)) => x == y,
        (CellValue::Int(x), CellValue::Int(y)) => x == y,
        (CellValue::Float(x), CellValue::Float(y)) => {
            // Float comparison with tolerance
            (x - y).abs() < 1e-10
        },
        (CellValue::String(x), CellValue::String(y)) => x == y,
        (CellValue::DateTime(x), CellValue::DateTime(y)) => (x - y).abs() < 1e-10,
        (CellValue::Error(x), CellValue::Error(y)) => x == y,
        (CellValue::Formula { formula: x, .. }, CellValue::Formula { formula: y, .. }) => x == y,
        _ => false,
    }
}

fn test_csv_roundtrip(data: &[Vec<CellValue>]) -> ExampleResult<()> {
    println!("Testing CSV round-trip...");

    // Write
    let file = File::create("roundtrip.csv")?;
    let mut writer = BufWriter::new(file);
    write_delimited(data, &mut writer, DelimitedConfig::csv())?;
    writer.flush()?;

    // Read back
    let mut read_file = File::open("roundtrip.csv")?;
    let read_data = read_delimited(&mut read_file, DelimitedConfig::csv())?;

    compare_data(data, &read_data, "CSV");
    Ok(())
}

fn test_tsv_roundtrip(data: &[Vec<CellValue>]) -> ExampleResult<()> {
    println!("Testing TSV round-trip...");

    // Write
    let file = File::create("roundtrip.tsv")?;
    let mut writer = BufWriter::new(file);
    write_delimited(data, &mut writer, DelimitedConfig::tsv())?;
    writer.flush()?;

    // Read back
    let mut read_file = File::open("roundtrip.tsv")?;
    let read_data = read_delimited(&mut read_file, DelimitedConfig::tsv())?;

    compare_data(data, &read_data, "TSV");
    Ok(())
}

fn test_prn_delimited_roundtrip(data: &[Vec<CellValue>]) -> ExampleResult<()> {
    println!("Testing PRN (semicolon-delimited) round-trip...");

    // Write
    let file = File::create("roundtrip_delimited.prn")?;
    let mut writer = BufWriter::new(file);
    write_delimited(data, &mut writer, DelimitedConfig::prn())?;
    writer.flush()?;

    // Read back
    let mut read_file = File::open("roundtrip_delimited.prn")?;
    let read_data = read_delimited(&mut read_file, DelimitedConfig::prn())?;

    compare_data(data, &read_data, "PRN (delimited)");
    Ok(())
}

fn test_prn_fixed_width_roundtrip(data: &[Vec<CellValue>]) -> ExampleResult<()> {
    println!("Testing PRN (fixed-width) round-trip...");

    // Write with fixed widths
    let file = File::create("roundtrip_fixed.prn")?;
    let mut writer = BufWriter::new(file);

    let config = FixedWidthConfig {
        column_widths: vec![15, 15, 15],
        auto_detect_widths: false,
        trim_fields: true,
        strip_bom: true,
        write_bom: None,
    };

    write_fixed_width(data, &mut writer, config.clone())?;
    writer.flush()?;

    // Read back with same widths
    let mut read_file = File::open("roundtrip_fixed.prn")?;
    let read_data = read_fixed_width(&mut read_file, config)?;

    compare_data(data, &read_data, "PRN (fixed-width)");
    Ok(())
}

fn test_sylk_roundtrip(data: &[Vec<CellValue>]) -> ExampleResult<()> {
    println!("Testing SYLK round-trip...");

    // Write
    let file = File::create("roundtrip.slk")?;
    let mut writer = BufWriter::new(file);
    write_sylk(data, &mut writer, SylkConfig::default())?;
    writer.flush()?;

    // Read back
    let mut read_file = File::open("roundtrip.slk")?;
    let read_data = read_sylk(&mut read_file, SylkConfig::default())?;

    compare_data(data, &read_data, "SYLK");
    Ok(())
}

fn test_dif_roundtrip(data: &[Vec<CellValue>]) -> ExampleResult<()> {
    println!("Testing DIF round-trip...");

    // Write
    let file = File::create("roundtrip.dif")?;
    let mut writer = BufWriter::new(file);
    write_dif(data, &mut writer, DifConfig::default())?;
    writer.flush()?;

    // Read back
    let mut read_file = File::open("roundtrip.dif")?;
    let read_data = read_dif(&mut read_file, DifConfig::default())?;

    compare_data(data, &read_data, "DIF");
    Ok(())
}
