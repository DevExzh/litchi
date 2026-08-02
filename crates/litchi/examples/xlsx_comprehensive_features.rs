//! Comprehensive demonstration of XLSX features
//!
//! This example demonstrates all the major features of the XLSX writer,
//! including the newly completed features for hyperlinks, comments,
//! conditional formatting, images, page setup, and more.

use litchi::ooxml::Props;
use litchi::ooxml::xlsx::{
    CellBorder, CellBorderLineStyle, CellBorderSide, CellFill, CellFillPatternType, CellFont,
    CellFormat, ConditionalFormatType, HeaderFooter, Workbook,
};

fn main() -> Result<(), Box<dyn std::error::Error + Send + Sync>> {
    println!("Creating comprehensive XLSX workbook...");

    // Create a new workbook
    let mut workbook = Workbook::create()?;

    // Set workbook properties
    let _ = workbook.put_props(
        Props::new()
            .title("Comprehensive Features Demo")
            .creator("Litchi Library")
            .subject("Feature Demonstration"),
    );

    // ===== WORKSHEET 1: Basic Features =====
    {
        let ws = workbook.worksheet_mut(0)?;
        ws.set_name("Basic Features".to_string());

        // Set some cell values
        ws.set_cell_value(1, 1, "Product");
        ws.set_cell_value(1, 2, "Quantity");
        ws.set_cell_value(1, 3, "Price");
        ws.set_cell_value(1, 4, "Total");

        ws.set_cell_value(2, 1, "Widget A");
        ws.set_cell_value(2, 2, 10);
        ws.set_cell_value(2, 3, 25.50);
        ws.set_cell_formula(2, 4, "B2*C2");

        ws.set_cell_value(3, 1, "Widget B");
        ws.set_cell_value(3, 2, 5);
        ws.set_cell_value(3, 3, 42.00);
        ws.set_cell_formula(3, 4, "B3*C3");

        // Apply cell formatting
        let header_format = CellFormat {
            font: Some(CellFont {
                name: Some("Arial".to_string()),
                size: Some(12.0),
                bold: true,
                italic: false,
                underline: false,
                color: Some("FFFFFF".to_string()),
            }),
            fill: Some(CellFill {
                pattern_type: CellFillPatternType::Solid,
                fg_color: Some("4472C4".to_string()),
                bg_color: None,
            }),
            border: Some(CellBorder {
                left: None,
                right: None,
                top: None,
                bottom: Some(CellBorderSide {
                    style: CellBorderLineStyle::Thick,
                    color: Some("000000".to_string()),
                }),
                diagonal: None,
            }),
            number_format: None,
        };

        for col in 1..=4 {
            ws.set_cell_format(1, col, header_format.clone());
        }

        // Merge cells
        ws.merge_cells(5, 1, 5, 4);
        ws.set_cell_value(5, 1, "Summary Section");

        // Set column widths
        ws.set_column_width(1, 15.0);
        ws.set_column_width(2, 10.0);
        ws.set_column_width(3, 10.0);
        ws.set_column_width(4, 12.0);

        // Freeze panes
        ws.freeze_panes(1, 0);
    }

    // ===== WORKSHEET 2: Hyperlinks and Comments =====
    {
        workbook.add_worksheet("Links & Comments");
        let ws = workbook.worksheet_mut(1)?;

        // Add hyperlinks
        ws.set_cell_value(1, 1, "Visit our website");
        ws.set_hyperlink(1, 1, "https://example.com", Some("Click here"));

        ws.set_cell_value(2, 1, "Email us");
        ws.set_hyperlink(2, 1, "mailto:info@example.com", None);

        ws.set_cell_value(3, 1, "Internal link");
        ws.set_hyperlink(3, 1, "Sheet1!A1", Some("Go to Sheet1"));

        // Add comments
        ws.set_cell_value(5, 1, "Important data");
        ws.set_cell_comment(5, 1, "This is a critical value!", "John Doe");

        ws.set_cell_value(6, 1, "Review required");
        ws.set_cell_comment(6, 1, "Please review this before the meeting.", "Jane Smith");

        ws.set_column_width(1, 20.0);
    }

    // ===== WORKSHEET 3: Conditional Formatting =====
    {
        workbook.add_worksheet("Conditional Format");
        let ws = workbook.worksheet_mut(2)?;

        // Create data
        ws.set_cell_value(1, 1, "Value");
        for i in 2..=11 {
            ws.set_cell_value(i, 1, (i as i64 - 2) * 10);
        }

        ws.set_cell_value(1, 3, "Score");
        for i in 2..=11 {
            ws.set_cell_value(i, 3, (i as i64 - 2) * 5);
        }

        // Add conditional formatting - CellIs
        ws.add_conditional_formatting(
            "A2:A11",
            ConditionalFormatType::CellIs {
                operator: "greaterThan".to_string(),
                formula: "50".to_string(),
            },
            1,
            None,
        );

        // Add conditional formatting - Color Scale
        ws.add_conditional_formatting(
            "C2:C11",
            ConditionalFormatType::ColorScale {
                min_color: "FF0000".to_string(),
                max_color: "00FF00".to_string(),
                mid_color: Some("FFFF00".to_string()),
            },
            2,
            None,
        );

        // Add conditional formatting - Data Bar
        ws.add_conditional_formatting(
            "A2:A11",
            ConditionalFormatType::DataBar {
                color: "638EC6".to_string(),
                show_value: true,
            },
            3,
            None,
        );

        ws.set_column_width(1, 15.0);
        ws.set_column_width(3, 15.0);
    }

    // ===== WORKSHEET 4: Data Validation & Auto-filter =====
    {
        workbook.add_worksheet("Validation & Filter");
        let ws = workbook.worksheet_mut(3)?;

        // Headers
        ws.set_cell_value(1, 1, "Category");
        ws.set_cell_value(1, 2, "Status");
        ws.set_cell_value(1, 3, "Priority");

        // Sample data
        ws.set_cell_value(2, 1, "Development");
        ws.set_cell_value(2, 2, "In Progress");
        ws.set_cell_value(2, 3, "High");

        // Add data validation
        use litchi::ooxml::xlsx::DataValidationType;

        ws.add_data_validation(
            "B2:B10",
            DataValidationType::List {
                values: vec![
                    "Not Started".to_string(),
                    "In Progress".to_string(),
                    "Completed".to_string(),
                ],
            },
            true,
            Some("Select Status"),
            Some("Choose a status from the list"),
            true,
            Some("Invalid Selection"),
            Some("Please select a valid status"),
        );

        ws.add_data_validation(
            "C2:C10",
            DataValidationType::List {
                values: vec!["Low".to_string(), "Medium".to_string(), "High".to_string()],
            },
            true,
            None,
            None,
            false,
            None,
            None,
        );

        // Add auto-filter
        ws.set_auto_filter("A1:C10");

        ws.set_column_width(1, 15.0);
        ws.set_column_width(2, 15.0);
        ws.set_column_width(3, 12.0);
    }

    // ===== WORKSHEET 5: Page Setup & Headers/Footers =====
    {
        workbook.add_worksheet("Page Setup");
        let ws = workbook.worksheet_mut(4)?;

        // Add content
        ws.set_cell_value(1, 1, "This sheet demonstrates page setup");
        ws.set_cell_value(2, 1, "Check Print Preview to see headers/footers");

        // Configure page setup
        ws.set_page_setup_with_options("landscape", 9, Some(100), None, None)?;

        // Set headers and footers
        let hf = HeaderFooter {
            header_left: Some("Company Name".to_string()),
            header_center: Some("Confidential Document".to_string()),
            header_right: Some("&D".to_string()), // Current date
            footer_left: Some("Department".to_string()),
            footer_center: Some("Page &P of &N".to_string()), // Page numbers
            footer_right: Some("&T".to_string()),             // Current time
        };

        ws.set_header_footer(hf);

        // Set print area and repeating rows
        ws.set_print_area("A1:E20");
        ws.set_repeating_rows("1:1");

        ws.set_column_width(1, 40.0);
    }

    // ===== WORKSHEET 6: Protection & Grouping =====
    {
        workbook.add_worksheet("Protection");
        let ws = workbook.worksheet_mut(5)?;

        // Add content
        ws.set_cell_value(1, 1, "Protected Sheet");
        ws.set_cell_value(2, 1, "This sheet is protected with password: 'secret'");

        // Group rows
        ws.set_cell_value(4, 1, "Group 1");
        ws.set_cell_value(5, 1, "Item 1");
        ws.set_cell_value(6, 1, "Item 2");
        ws.group_rows(5, 6, 1);

        ws.set_cell_value(8, 1, "Group 2");
        ws.set_cell_value(9, 1, "Item A");
        ws.set_cell_value(10, 1, "Item B");
        ws.set_cell_value(11, 1, "Item C");
        ws.group_rows(9, 11, 1);

        // Protect sheet
        ws.protect_sheet(Some("secret"));

        ws.set_column_width(1, 30.0);
    }

    // ===== Workbook-level Features =====

    // Set active sheet
    workbook.set_active_sheet(0)?;

    // Set tab colors
    workbook.set_tab_color(0, "FF0000")?; // Red for first sheet
    workbook.set_tab_color(2, "00FF00")?; // Green for conditional formatting

    // Define named ranges (sheet names with spaces must be quoted)
    workbook.define_name("SalesData", "'Basic Features'!$A$2:$D$3");
    workbook.define_name("TotalColumn", "'Basic Features'!$D:$D");

    // Set calculation mode
    workbook.set_calculation_mode("auto")?;

    // Protect workbook structure
    workbook.protect_workbook(Some("workbook123"), true, false);

    // Save the workbook
    println!("Saving workbook...");
    workbook.save("xlsx_comprehensive_features.xlsx")?;

    println!("✅ Successfully created xlsx_comprehensive_features.xlsx");
    println!("The file demonstrates:");
    println!("  - Cell values and formulas");
    println!("  - Cell formatting (fonts, fills, borders)");
    println!("  - Merged cells");
    println!("  - Column widths and freeze panes");
    println!("  - Hyperlinks (external and internal)");
    println!("  - Cell comments");
    println!("  - Conditional formatting (CellIs, ColorScale, DataBar)");
    println!("  - Data validation");
    println!("  - Auto-filters");
    println!("  - Page setup and print settings");
    println!("  - Headers and footers");
    println!("  - Sheet protection");
    println!("  - Row/column grouping");
    println!("  - Named ranges");
    println!("  - Tab colors");
    println!("  - Workbook protection");

    Ok(())
}
