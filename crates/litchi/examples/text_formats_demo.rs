//! Demonstration of text format support (CSV/TSV/PRN/SYLK/DIF)
//!
//! This example shows how to read and write various text-based spreadsheet formats
//! with optional BOM handling and different configurations.

use litchi::sheet::CellValue;
use litchi::sheet::text::formats::*;
use std::error::Error;
use std::fs::File;
use std::io::{BufReader, BufWriter, Write};

type ExampleResult<T> = Result<T, Box<dyn Error + Send + Sync>>;

fn main() -> ExampleResult<()> {
    println!("=== Text Formats Demo ===\n");

    // Create sample data
    let sample_data = create_sample_data();

    // Test each format
    test_csv(&sample_data)?;
    test_tsv(&sample_data)?;
    test_prn_delimited(&sample_data)?;
    test_prn_fixed_width(&sample_data)?;
    test_sylk(&sample_data)?;
    test_dif(&sample_data)?;

    // Test BOM variants
    test_bom_variants(&sample_data)?;

    println!("\n=== All formats generated successfully! ===");
    println!("Check the generated files in the current directory.");

    Ok(())
}

fn create_sample_data() -> Vec<Vec<CellValue>> {
    vec![
        vec![
            CellValue::String("Name".to_string()),
            CellValue::String("Age".to_string()),
            CellValue::String("Salary".to_string()),
            CellValue::String("Active".to_string()),
            CellValue::String("Notes".to_string()),
        ],
        vec![
            CellValue::String("Alice Johnson".to_string()),
            CellValue::Int(28),
            CellValue::Float(75000.50),
            CellValue::Bool(true),
            CellValue::String("Top performer, \"excellent\" reviews".to_string()),
        ],
        vec![
            CellValue::String("Bob Smith".to_string()),
            CellValue::Int(35),
            CellValue::Float(82000.0),
            CellValue::Bool(true),
            CellValue::String("Team lead".to_string()),
        ],
        vec![
            CellValue::String("Carol Davis".to_string()),
            CellValue::Int(42),
            CellValue::Float(95000.75),
            CellValue::Bool(false),
            CellValue::String("On leave".to_string()),
        ],
        vec![
            CellValue::String("David Wilson".to_string()),
            CellValue::Int(31),
            CellValue::Float(68000.25),
            CellValue::Bool(true),
            CellValue::String("New hire, learning fast".to_string()),
        ],
    ]
}

fn test_csv(data: &[Vec<CellValue>]) -> ExampleResult<()> {
    println!("Testing CSV format...");

    let file = File::create("demo.csv")?;
    let mut writer = BufWriter::new(file);

    let config = DelimitedConfig::csv().with_write_bom(Some(litchi::common::BomKind::Utf8));

    write_delimited(data, &mut writer, config)?;
    writer.flush()?;

    println!("  ✓ Generated demo.csv with UTF-8 BOM");

    // Read back
    let file = File::open("demo.csv")?;
    let mut read_file = BufReader::new(file);
    let data_read = read_delimited(&mut read_file, DelimitedConfig::csv())?;

    println!("  ✓ Read back {} rows", data_read.len());
    Ok(())
}

fn test_tsv(data: &[Vec<CellValue>]) -> ExampleResult<()> {
    println!("Testing TSV format...");

    let file = File::create("demo.tsv")?;
    let mut writer = BufWriter::new(file);

    let config = DelimitedConfig::tsv();

    write_delimited(data, &mut writer, config)?;
    writer.flush()?;

    println!("  ✓ Generated demo.tsv");
    Ok(())
}

fn test_prn_delimited(data: &[Vec<CellValue>]) -> ExampleResult<()> {
    println!("Testing PRN (semicolon-delimited) format...");

    let file = File::create("demo_delimited.prn")?;
    let mut writer = BufWriter::new(file);

    let config = DelimitedConfig::prn();

    write_delimited(data, &mut writer, config)?;
    writer.flush()?;

    println!("  ✓ Generated demo_delimited.prn");
    Ok(())
}

fn test_prn_fixed_width(data: &[Vec<CellValue>]) -> ExampleResult<()> {
    println!("Testing PRN (fixed-width) format...");

    let file = File::create("demo_fixed.prn")?;
    let mut writer = BufWriter::new(file);

    let config = FixedWidthConfig {
        column_widths: vec![20, 8, 12, 8, 30],
        auto_detect_widths: false,
        trim_fields: true,
        strip_bom: true,
        write_bom: None,
    };

    write_fixed_width(data, &mut writer, config)?;
    writer.flush()?;

    println!("  ✓ Generated demo_fixed.prn");

    // Test auto-detection
    let file = File::create("demo_auto.prn")?;
    let mut writer = BufWriter::new(file);

    let auto_config = FixedWidthConfig::default();
    write_fixed_width(data, &mut writer, auto_config)?;
    writer.flush()?;

    println!("  ✓ Generated demo_auto.prn (auto-detected widths)");
    Ok(())
}

fn test_sylk(data: &[Vec<CellValue>]) -> ExampleResult<()> {
    println!("Testing SYLK format...");

    let file = File::create("demo.slk")?;
    let mut writer = BufWriter::new(file);

    let config = SylkConfig::default();
    write_sylk(data, &mut writer, config.clone())?;
    writer.flush()?;

    println!("  ✓ Generated demo.slk");

    // Test with formula
    let mut formula_data = data.to_vec();
    formula_data.push(vec![
        CellValue::String("Total".to_string()),
        CellValue::Formula {
            formula: "SUM(B2:B5)".to_string(),
            cached_value: Some(Box::new(CellValue::Int(136))),
            is_array: false,
            array_range: None,
        },
        CellValue::Formula {
            formula: "SUM(C2:C5)".to_string(),
            cached_value: Some(Box::new(CellValue::Float(320001.5))),
            is_array: false,
            array_range: None,
        },
        CellValue::Empty,
        CellValue::Empty,
    ]);

    let file = File::create("demo_formulas.slk")?;
    let mut writer = BufWriter::new(file);
    write_sylk(&formula_data, &mut writer, config)?;
    writer.flush()?;

    println!("  ✓ Generated demo_formulas.slk (with formulas)");
    Ok(())
}

fn test_dif(data: &[Vec<CellValue>]) -> ExampleResult<()> {
    println!("Testing DIF format...");

    let file = File::create("demo.dif")?;
    let mut writer = BufWriter::new(file);

    let config = DifConfig::default();
    write_dif(data, &mut writer, config)?;
    writer.flush()?;

    println!("  ✓ Generated demo.dif");

    // Test with UTF-16 BOM
    let file = File::create("demo_utf16.dif")?;
    let mut writer = BufWriter::new(file);

    let config = DifConfig {
        strip_bom: true,
        write_bom: Some(litchi::common::BomKind::Utf16Le),
    };
    write_dif(data, &mut writer, config)?;
    writer.flush()?;

    println!("  ✓ Generated demo_utf16.dif (with UTF-16 LE BOM)");
    Ok(())
}

fn test_bom_variants(data: &[Vec<CellValue>]) -> ExampleResult<()> {
    println!("Testing BOM variants...");

    let boms = vec![
        ("UTF-8", litchi::common::BomKind::Utf8),
        ("UTF-16 LE", litchi::common::BomKind::Utf16Le),
        ("UTF-16 BE", litchi::common::BomKind::Utf16Be),
        ("UTF-32 LE", litchi::common::BomKind::Utf32Le),
        ("UTF-32 BE", litchi::common::BomKind::Utf32Be),
    ];

    for (name, bom_kind) in boms {
        let filename = format!("demo_{}.csv", name.to_lowercase().replace(' ', "_"));
        let file = File::create(&filename)?;
        let mut writer = BufWriter::new(file);

        let config = DelimitedConfig::csv().with_write_bom(Some(bom_kind));

        write_delimited(data, &mut writer, config)?;
        writer.flush()?;

        println!("  ✓ Generated {} with {} BOM", filename, name);
    }

    Ok(())
}
