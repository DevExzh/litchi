//! Pivot Table Writer — XLS Example
//!
//! Demonstrates the `add_pivot_table` API on `Writer`. Generates an XLS
//! file with a source data sheet and a second sheet containing a pivot table
//! definition (SXVS/SXVIEW/SXVD/SXVI/SXDI records).
//!
//! Run with: `cargo run --example xls_pivot_table`
//!
//! The file is saved to `output/xls_pivot_table.xls`.

use litchi_xls::Writer;
use litchi_xls::writer::{
    Column, PivotCacheValue, PivotDataItemConfig, PivotFieldConfig, PivotItemConfig,
    PivotTableConfig,
};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output_dir = std::path::Path::new("output");
    std::fs::create_dir_all(output_dir)?;
    let output_path = output_dir.join("xls_pivot_table.xls");

    let mut w = Writer::new();

    // ================================================================
    // Sheet 1 — Source Data
    // ================================================================
    let src = w.add_worksheet("Sales Data")?;

    let headers = ["Region", "Product", "Quarter", "Revenue"];
    for (c, h) in headers.iter().enumerate() {
        w.write_string(src, 0, c as u16, h)?;
    }

    let rows: &[(&str, &str, &str, f64)] = &[
        ("North", "Widget A", "Q1", 12500.0),
        ("North", "Widget A", "Q2", 13200.0),
        ("North", "Widget B", "Q1", 8400.0),
        ("North", "Widget B", "Q2", 9100.0),
        ("South", "Widget A", "Q1", 9800.0),
        ("South", "Widget A", "Q2", 10500.0),
        ("South", "Widget B", "Q1", 7200.0),
        ("South", "Widget B", "Q2", 7800.0),
        ("East", "Widget A", "Q1", 15200.0),
        ("East", "Widget A", "Q2", 14800.0),
        ("East", "Widget B", "Q1", 11000.0),
        ("East", "Widget B", "Q2", 11600.0),
        ("West", "Widget A", "Q1", 11000.0),
        ("West", "Widget A", "Q2", 11800.0),
        ("West", "Widget B", "Q1", 6800.0),
        ("West", "Widget B", "Q2", 7400.0),
    ];

    for (i, (region, product, quarter, revenue)) in rows.iter().enumerate() {
        let row = 1 + i as u32;
        w.write_string(src, row, 0, region)?;
        w.write_string(src, row, 1, product)?;
        w.write_string(src, row, 2, quarter)?;
        w.write_number(src, row, 3, *revenue)?;
    }

    w.set_column_width(src, Column::new(0)?, 12.0)?;
    w.set_column_width(src, Column::new(1)?, 14.0)?;
    w.set_column_width(src, Column::new(2)?, 10.0)?;
    w.set_column_width(src, Column::new(3)?, 14.0)?;

    println!("[Sheet 1] Sales Data — {} rows of source data", rows.len());

    // ================================================================
    // Sheet 2 — Pivot Table
    // ================================================================
    let pt_sheet = w.add_worksheet("Pivot Table")?;

    // Write a header so the sheet isn't entirely empty when opened.
    w.write_string(pt_sheet, 0, 0, "Revenue by Region and Product")?;

    // Define the pivot table structure:
    //   Row axis:  Region  (field 0)
    //   Column axis: Product (field 1)
    //   Page axis: Quarter (field 2) — allows filtering by quarter
    //   Data field: Sum of Revenue (field 3)
    //
    // The output range starts at row 2 to leave room for the title.
    let pivot_config = PivotTableConfig {
        name: "SalesPivot".to_string(),
        source_type: 0x0001, // Worksheet

        // Source data range: "Sales Data" sheet, rows 0..16, cols 0..3
        source_sheet_name: "Sales Data".to_string(),
        source_first_row: 0,                // header row
        source_last_row: rows.len() as u16, // last data row (inclusive)
        source_first_col: 0,
        source_last_col: 3,

        // Output range
        first_row: 2,
        last_row: 8, // approximate output extent
        first_col: 0,
        last_col: 3,
        first_header_row: 3, // LO: first_row + 1 (col headers below title row)
        first_data_row: 4,
        first_data_col: 1,
        data_field_name: "Values".to_string(),
        data_axis: 0x0002,     // column axis
        data_position: 0xFFFF, // EXC_SXVIEW_DATALAST (single data field)
        fields: vec![
            // Field 0: Region — row axis
            PivotFieldConfig {
                axis: 0x0001, // row
                subtotal_count: 1,
                subtotal_flags: 0x0001, // DEFAULT subtotal
                items: vec![
                    PivotItemConfig {
                        item_type: 0x0000, // EXC_SXVI_TYPE_DATA
                        flags: 0,
                        cache_index: 0,
                        name: None, // use cache name
                    },
                    PivotItemConfig {
                        item_type: 0x0000,
                        flags: 0,
                        cache_index: 1,
                        name: None,
                    },
                    PivotItemConfig {
                        item_type: 0x0000,
                        flags: 0,
                        cache_index: 2,
                        name: None,
                    },
                    PivotItemConfig {
                        item_type: 0x0000,
                        flags: 0,
                        cache_index: 3,
                        name: None,
                    },
                    // DEFAULT subtotal item — required by Excel
                    PivotItemConfig {
                        item_type: 0x0001, // EXC_SXVI_TYPE_DEFAULT
                        flags: 0,
                        cache_index: 0xFFFF,
                        name: None,
                    },
                ],
                name: None, // use cache name
                cache_name: "Region".to_string(),
                cache_items: vec!["North".into(), "South".into(), "East".into(), "West".into()],
                is_numeric: false,
                grouping: None,
            },
            // Field 1: Product — column axis
            PivotFieldConfig {
                axis: 0x0002, // column
                subtotal_count: 1,
                subtotal_flags: 0x0001,
                items: vec![
                    PivotItemConfig {
                        item_type: 0x0000,
                        flags: 0,
                        cache_index: 0,
                        name: None,
                    },
                    PivotItemConfig {
                        item_type: 0x0000,
                        flags: 0,
                        cache_index: 1,
                        name: None,
                    },
                    PivotItemConfig {
                        item_type: 0x0001, // DEFAULT subtotal
                        flags: 0,
                        cache_index: 0xFFFF,
                        name: None,
                    },
                ],
                name: None,
                cache_name: "Product".to_string(),
                cache_items: vec!["Widget A".into(), "Widget B".into()],
                is_numeric: false,
                grouping: None,
            },
            // Field 2: Quarter — page axis
            PivotFieldConfig {
                axis: 0x0004, // page
                subtotal_count: 1,
                subtotal_flags: 0x0001,
                items: vec![
                    PivotItemConfig {
                        item_type: 0x0000,
                        flags: 0,
                        cache_index: 0,
                        name: None,
                    },
                    PivotItemConfig {
                        item_type: 0x0000,
                        flags: 0,
                        cache_index: 1,
                        name: None,
                    },
                    PivotItemConfig {
                        item_type: 0x0001, // DEFAULT subtotal
                        flags: 0,
                        cache_index: 0xFFFF,
                        name: None,
                    },
                ],
                name: None,
                cache_name: "Quarter".to_string(),
                cache_items: vec!["Q1".into(), "Q2".into()],
                is_numeric: false,
                grouping: None,
            },
            // Field 3: Revenue — data axis (aggregated)
            PivotFieldConfig {
                axis: 0x0008, // data
                subtotal_count: 1,
                subtotal_flags: 0x0001,
                items: vec![],
                name: None, // use cache name
                cache_name: "Revenue".to_string(),
                cache_items: vec![], // numeric field, no string cache items
                is_numeric: true,
                grouping: None,
            },
        ],
        data_items: vec![
            // Sum of Revenue
            PivotDataItemConfig {
                source_field_index: 3, // Revenue is field index 3
                function: 0,           // 0 = Sum
                display_format: 0,
                base_field_index: 0,
                base_item_index: 0,
                num_format_index: 0,
                name: "Sum of Revenue".to_string(),
            },
        ],
        page_entries: vec![
            // Page field: Quarter (field 2), showing all items by default.
            // (item_index=0x7FFD = EXC_SXPI_ALLITEMS, field_index=2, object_id=1)
            (0x7FFD, 2, 0x0001),
        ],
        // Source data for the pivot cache (SXDBB + SXNUM records).
        // Each row: [Region_idx, Product_idx, Quarter_idx, Revenue_value]
        // Cache item indices: Region=[North=0,South=1,East=2,West=3],
        //                     Product=[Widget A=0,Widget B=1], Quarter=[Q1=0,Q2=1]
        source_data: rows
            .iter()
            .map(|(region, product, quarter, revenue)| {
                let region_idx = ["North", "South", "East", "West"]
                    .iter()
                    .position(|&r| r == *region)
                    .unwrap() as u8;
                let product_idx = ["Widget A", "Widget B"]
                    .iter()
                    .position(|&p| p == *product)
                    .unwrap() as u8;
                let quarter_idx = ["Q1", "Q2"].iter().position(|&q| q == *quarter).unwrap() as u8;
                vec![
                    PivotCacheValue::StringIndex(region_idx),
                    PivotCacheValue::StringIndex(product_idx),
                    PivotCacheValue::StringIndex(quarter_idx),
                    PivotCacheValue::Number(*revenue),
                ]
            })
            .collect(),
    };

    w.add_pivot_table(pt_sheet, pivot_config)?;

    w.set_column_width(pt_sheet, Column::new(0)?, 16.0)?;
    for c in 1..=3 {
        w.set_column_width(pt_sheet, Column::new(c)?, 14.0)?;
    }

    println!(
        "[Sheet 2] Pivot Table — Region(row) x Product(col), Sum of Revenue, Quarter page filter"
    );

    // ================================================================
    // Save
    // ================================================================
    w.save(&output_path)?;
    println!("\nSaved to: {}", output_path.display());
    println!("Open in Excel to inspect the pivot table SX* records.");

    Ok(())
}
