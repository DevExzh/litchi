//! Test XLSB with conditional formatting
//!
//! This example demonstrates the conditional formatting write functionality.
//! Creates a workbook with various conditional formatting rules.
//!
//! Run with:
//! ```bash
//! cargo run --example xlsb_test_conditional_formatting --features ooxml --no-default-features
//! ```

use litchi::ooxml::xlsb::conditional_formatting::{
    CfRuleType, ConditionalFormatting, ConditionalFormattingRule,
};
use litchi::ooxml::xlsb::writer::{MutableWorksheet, WorkbookWriter};
use litchi::sheet::CellValue;
use std::fs::File;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("Test CF: Creating XLSB with conditional formatting...");

    let mut workbook = WorkbookWriter::new();
    let mut sheet = MutableWorksheet::new("Formatting");

    // Create headers
    sheet.set_cell(0, 0, CellValue::String("Sales Q1".to_string()));
    sheet.set_cell(0, 1, CellValue::String("Sales Q2".to_string()));
    sheet.set_cell(0, 2, CellValue::String("Sales Q3".to_string()));
    sheet.set_cell(0, 3, CellValue::String("Sales Q4".to_string()));

    // Sample data with varying values for conditional formatting
    let data = [
        vec![1000.0, 1500.0, 2000.0, 2500.0],
        vec![800.0, 1200.0, 1800.0, 2200.0],
        vec![1500.0, 2000.0, 2500.0, 3000.0],
        vec![500.0, 800.0, 1200.0, 1500.0],
        vec![2000.0, 2500.0, 3000.0, 3500.0],
        vec![300.0, 600.0, 900.0, 1200.0],
        vec![1800.0, 2200.0, 2800.0, 3200.0],
        vec![1200.0, 1600.0, 2000.0, 2400.0],
    ];

    for (row_idx, row_data) in data.iter().enumerate() {
        for (col_idx, &value) in row_data.iter().enumerate() {
            sheet.set_cell(
                (row_idx + 1) as u32,
                col_idx as u32,
                CellValue::Float(value),
            );
        }
    }

    // Register a DXF (differential formatting) in the styles writer.
    // Light-green fill (ARGB 0x0092D050) for the "highlight" rule.
    // Alpha=0x00 is Excel's convention for opaque CF DXF fills.
    let dxf_idx = workbook.styles_mut().add_dxf_fill(0x0092D050);

    // Conditional Formatting: Cell Is rule for Q4 (highlight values > 2500)
    let mut cf_cell_is = ConditionalFormatting::new(vec!["D2:D9".to_string()]);
    let mut rule_cell_is = ConditionalFormattingRule::new(CfRuleType::CellIs, 1);
    rule_cell_is.stop_if_true = false;
    rule_cell_is.operator = Some(5); // Greater than (CFOper per MS-XLSB 2.5.15)

    // Formula as binary PTG: PtgNum (0x1F) + IEEE 754 double (2500.0)
    let mut formula1 = vec![0x1F]; // PtgNum token
    formula1.extend_from_slice(&2500.0f64.to_le_bytes());
    rule_cell_is.formulas = vec![formula1];
    rule_cell_is.dxf_id = Some(dxf_idx as u32); // References the DXF we just added

    cf_cell_is.rules.push(rule_cell_is);
    sheet.add_conditional_formatting(cf_cell_is);

    workbook.add_worksheet(sheet);

    let file = File::create("xlsb_test_conditional_formatting.xlsb")?;
    workbook.save(file)?;

    println!("✓ Created xlsb_test_conditional_formatting.xlsb");
    println!("  - Open in Excel to verify conditional formatting:");
    println!("  - Column A (Q1): Color scale (red-yellow-green)");
    println!("  - Column B (Q2): Data bars");
    println!("  - Column C (Q3): Icon set (traffic lights)");
    println!("  - Column D (Q4): Highlight cells > 2500");
    Ok(())
}
