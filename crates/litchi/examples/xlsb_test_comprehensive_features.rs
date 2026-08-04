//! Comprehensive test for XLSB advanced features
//!
//! This example combines all newly implemented features:
//! - Named ranges
//! - Data validation
//! - Conditional formatting
//! - Merged cells
//! - Hyperlinks
//!
//! Run with:
//! ```bash
//! cargo run --example xlsb_test_comprehensive_features --features ooxml --no-default-features
//! ```

use litchi::ooxml::xlsb::conditional_formatting::{
    CfRuleType, Cfvo, ColorScale, ConditionalFormatting, ConditionalFormattingRule,
};
use litchi::ooxml::xlsb::data_validation::DataValidation;
use litchi::ooxml::xlsb::hyperlinks::Hyperlink;
use litchi::ooxml::xlsb::merged_cells::MergedCell;
use litchi::ooxml::xlsb::named_ranges::{NamedRange, create_area3d_formula};
use litchi::ooxml::xlsb::writer::{MutableWorksheet, WorkbookWriter};
use litchi::sheet::CellValue;
use std::fs::File;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("Test Comprehensive: Creating XLSB with all advanced features...");

    let mut workbook = WorkbookWriter::new();
    let mut sheet = MutableWorksheet::new("Dashboard");

    // === TITLE ROW WITH MERGED CELLS ===
    sheet.set_cell(0, 0, CellValue::String("Sales Dashboard 2024".to_string()));
    sheet.add_merged_cell(MergedCell::new(0, 0, 0, 5)); // A1:F1

    // === HEADERS ===
    sheet.set_cell(2, 0, CellValue::String("Region".to_string()));
    sheet.set_cell(2, 1, CellValue::String("Sales".to_string()));
    sheet.set_cell(2, 2, CellValue::String("Target".to_string()));
    sheet.set_cell(2, 3, CellValue::String("Achievement %".to_string()));
    sheet.set_cell(2, 4, CellValue::String("Status".to_string()));
    sheet.set_cell(2, 5, CellValue::String("Details".to_string()));

    // === DATA ===
    let regions = ["North", "South", "East", "West", "Central"];
    let sales = [45000.0, 38000.0, 52000.0, 41000.0, 48000.0];
    let targets = [40000.0, 40000.0, 50000.0, 45000.0, 45000.0];

    for (i, region) in regions.iter().enumerate() {
        let row = (i + 3) as u32;
        let sale = sales[i];
        let target = targets[i];
        let achievement = (sale / target) * 100.0;

        sheet.set_cell(row, 0, CellValue::String(region.to_string()));
        sheet.set_cell(row, 1, CellValue::Float(sale));
        sheet.set_cell(row, 2, CellValue::Float(target));
        sheet.set_cell(row, 3, CellValue::Float(achievement));

        // Status based on achievement
        let status = if achievement >= 100.0 {
            "Excellent"
        } else if achievement >= 90.0 {
            "Good"
        } else {
            "Needs Improvement"
        };
        sheet.set_cell(row, 4, CellValue::String(status.to_string()));

        // Details link
        sheet.set_cell(row, 5, CellValue::String("View Report".to_string()));
        let link = Hyperlink::new_external(
            row,
            5,
            row,
            5,
            format!("https://example.com/reports/{}", region.to_lowercase()),
        )
        .with_display("View Report".to_string());
        sheet.add_hyperlink(link);
    }

    // === TOTALS ROW ===
    let total_row = 8;
    sheet.set_cell(total_row, 0, CellValue::String("TOTAL".to_string()));
    sheet.set_cell(total_row, 1, CellValue::Float(sales.iter().sum::<f64>()));
    sheet.set_cell(total_row, 2, CellValue::Float(targets.iter().sum::<f64>()));

    // === DATA VALIDATION ===
    // Status column must be from predefined list
    let mut status_validation = DataValidation::new(3, "E4:E8".to_string()); // Type 3 = list
    status_validation.operator = 0;
    status_validation.formula1 = Some("\"Excellent,Good,Needs Improvement,Poor\"".to_string());
    status_validation.allow_blank = false;
    status_validation.show_dropdown = true;
    status_validation.show_error_message = true;
    status_validation.error_style = 0;
    status_validation.error_title = Some("Invalid Status".to_string());
    status_validation.error_text = Some("Please select from the list".to_string());
    sheet.add_data_validation(status_validation);

    // Sales must be positive numbers
    let mut sales_validation = DataValidation::new(2, "B4:B8".to_string()); // Type 2 = decimal
    sales_validation.operator = 3; // Greater than
    sales_validation.formula1 = Some("0".to_string());
    sales_validation.allow_blank = false;
    sales_validation.show_input_message = true;
    sales_validation.input_title = Some("Sales Amount".to_string());
    sales_validation.input_text = Some("Enter positive sales value".to_string());
    sheet.add_data_validation(sales_validation);

    // === CONDITIONAL FORMATTING ===
    // Color scale for achievement percentage
    let mut cf_achievement = ConditionalFormatting::new(vec!["D4:D8".to_string()]);
    let mut rule_achievement = ConditionalFormattingRule::new(CfRuleType::ColorScale, 1);
    let min_cfvo = Cfvo::new(1, Some("80".to_string()));
    let max_cfvo = Cfvo::new(2, Some("120".to_string()));
    let color_scale = ColorScale::new(min_cfvo, max_cfvo, 0x0000FF, 0x00FF00);

    rule_achievement.color_scale = Some(color_scale);
    cf_achievement.rules.push(rule_achievement);
    sheet.add_conditional_formatting(cf_achievement);

    // Highlight sales above target
    let mut cf_sales = ConditionalFormatting::new(vec!["B4:B8".to_string()]);
    let mut rule_sales = ConditionalFormattingRule::new(CfRuleType::CellIs, 2);
    rule_sales.operator = Some(5); // Greater than
    rule_sales.formula_texts = vec!["C4".to_string()]; // Compare with target
    rule_sales.dxf_id = Some(0);
    cf_sales.rules.push(rule_sales);
    sheet.add_conditional_formatting(cf_sales);

    workbook.add_worksheet(sheet);

    // === NAMED RANGES ===
    // Define named range for sales data B4:B8 (rows 3-7, col 1, 0-indexed)
    let sales_formula = create_area3d_formula(0, 3, 7, 1, 1)?;
    let sales_range = NamedRange::new("SalesData".to_string(), Some(0)).with_formula(sales_formula);
    workbook.add_named_range(sales_range);

    // Define named range for targets C4:C8 (rows 3-7, col 2, 0-indexed)
    let targets_formula = create_area3d_formula(0, 3, 7, 2, 2)?;
    let targets_range =
        NamedRange::new("Targets".to_string(), Some(0)).with_formula(targets_formula);
    workbook.add_named_range(targets_range);

    // Global named range for entire data area A3:F8 (rows 2-7, cols 0-5)
    let total_formula = create_area3d_formula(0, 2, 7, 0, 5)?;
    let total_range = NamedRange::new("GrandTotal".to_string(), None).with_formula(total_formula);
    workbook.add_named_range(total_range);

    let file = File::create("xlsb_test_comprehensive_features.xlsb")?;
    workbook.save(file)?;

    println!("✓ Created xlsb_test_comprehensive_features.xlsb");
    println!("\n=== Verification Steps ===");
    println!("1. Merged Cells: Title row (A1:F1) should be merged");
    println!("2. Hyperlinks: Click 'View Report' links in column F");
    println!("3. Data Validation:");
    println!("   - Try editing Status column (E4:E8) - should show dropdown");
    println!("   - Try entering negative sales values - should be rejected");
    println!("4. Conditional Formatting:");
    println!("   - Achievement % column should have color gradient");
    println!("   - Sales above target should be highlighted");
    println!("5. Named Ranges: Open Name Manager (Formulas > Name Manager)");
    println!("   - Should see: SalesData, Targets, GrandTotal");
    Ok(())
}
