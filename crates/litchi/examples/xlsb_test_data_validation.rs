//! Test XLSB with data validation
//!
//! This example demonstrates the data validation write functionality.
//! Creates a workbook with various data validation rules.
//!
//! Run with:
//! ```bash
//! cargo run --example xlsb_test_data_validation --features ooxml --no-default-features
//! ```

use litchi::ooxml::xlsb::data_validation::DataValidation;
use litchi::ooxml::xlsb::writer::{MutableWorksheet, WorkbookWriter};
use litchi::sheet::CellValue;
use std::fs::File;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("Test DV: Creating XLSB with data validation...");

    let mut workbook = WorkbookWriter::new();
    let mut sheet = MutableWorksheet::new("Validation");

    // Create headers
    sheet.set_cell(0, 0, CellValue::String("Age (18-65)".to_string()));
    sheet.set_cell(0, 1, CellValue::String("Email".to_string()));
    sheet.set_cell(0, 2, CellValue::String("Priority".to_string()));
    sheet.set_cell(0, 3, CellValue::String("Score (0-100)".to_string()));

    // Sample data
    sheet.set_cell(1, 0, CellValue::Float(25.0));
    sheet.set_cell(1, 1, CellValue::String("user@example.com".to_string()));
    sheet.set_cell(1, 2, CellValue::String("High".to_string()));
    sheet.set_cell(1, 3, CellValue::Float(85.5));

    // Data Validation 1: Age must be whole number between 18 and 65
    let mut age_validation = DataValidation::new(1, "A2:A100".to_string()); // Type 1 = whole number
    age_validation.operator = 0; // Between (MS-XLSB 2.5.15)
    age_validation.formula1 = Some("18".to_string());
    age_validation.formula2 = Some("65".to_string());
    age_validation.allow_blank = false;
    age_validation.show_dropdown = false;
    age_validation.show_input_message = true;
    age_validation.show_error_message = true;
    age_validation.error_style = 0; // Stop
    age_validation.input_title = Some("Age Entry".to_string());
    age_validation.input_text = Some("Please enter age between 18 and 65".to_string());
    age_validation.error_title = Some("Invalid Age".to_string());
    age_validation.error_text = Some("Age must be between 18 and 65".to_string());
    sheet.add_data_validation(age_validation);

    // Data Validation 2: Email must contain @ (custom text validation)
    let mut email_validation = DataValidation::new(7, "B2:B100".to_string()); // Type 7 = custom
    email_validation.operator = 0; // No operator for custom
    email_validation.formula1 = Some("ISNUMBER(FIND(\"@\",B2))".to_string());
    email_validation.allow_blank = true;
    email_validation.show_dropdown = false;
    email_validation.show_input_message = true;
    email_validation.show_error_message = true;
    email_validation.error_style = 1; // Warning
    email_validation.input_title = Some("Email".to_string());
    email_validation.input_text = Some("Enter a valid email address".to_string());
    email_validation.error_title = Some("Invalid Email".to_string());
    email_validation.error_text = Some("Email must contain @".to_string());
    sheet.add_data_validation(email_validation);

    // Data Validation 3: Priority must be from a list
    let mut priority_validation = DataValidation::new(3, "C2:C100".to_string()); // Type 3 = list
    priority_validation.operator = 0;
    priority_validation.formula1 = Some("\"Low,Medium,High,Critical\"".to_string());
    priority_validation.allow_blank = false;
    priority_validation.show_dropdown = true; // Show dropdown arrow
    priority_validation.show_input_message = false;
    priority_validation.show_error_message = true;
    priority_validation.error_style = 0; // Stop
    priority_validation.error_title = Some("Invalid Priority".to_string());
    priority_validation.error_text = Some("Please select from the list".to_string());
    sheet.add_data_validation(priority_validation);

    // Data Validation 4: Score must be decimal between 0 and 100
    let mut score_validation = DataValidation::new(2, "D2:D100".to_string()); // Type 2 = decimal
    score_validation.operator = 0; // Between
    score_validation.formula1 = Some("0".to_string());
    score_validation.formula2 = Some("100".to_string());
    score_validation.allow_blank = true;
    score_validation.show_dropdown = false;
    score_validation.show_input_message = true;
    score_validation.show_error_message = true;
    score_validation.error_style = 2; // Information
    score_validation.input_title = Some("Score".to_string());
    score_validation.input_text = Some("Enter score between 0 and 100".to_string());
    score_validation.error_title = Some("Score Note".to_string());
    score_validation.error_text = Some("Score should be between 0 and 100".to_string());
    sheet.add_data_validation(score_validation);

    workbook.add_worksheet(sheet);

    let file = File::create("xlsb_test_data_validation.xlsb")?;
    workbook.save(file)?;

    println!("✓ Created xlsb_test_data_validation.xlsb");
    println!("  - Open in Excel and try entering data in rows 2-100");
    println!("  - Column A: Only accepts ages 18-65");
    println!("  - Column B: Validates email format");
    println!("  - Column C: Shows dropdown with priorities");
    println!("  - Column D: Accepts scores 0-100");
    Ok(())
}
