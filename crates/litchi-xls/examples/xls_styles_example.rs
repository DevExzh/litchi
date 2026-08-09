use litchi_xls::writer::formatting::{
    COLOR_AUTOMATIC, COLOR_BLACK, COLOR_RED, COLOR_WHITE, FONT_WEIGHT_BOLD, FONT_WEIGHT_NORMAL,
};
use litchi_xls::writer::{
    BorderStyle, Borders, CellStyle, Column, ConditionalFormat, ConditionalFormatType,
    ConditionalPattern, DataValidation, DataValidationType, Fill, FillPattern, Font, FrozenPanes,
    HorizontalAlignment, Row, VerticalAlignment, Writer,
};

fn make_header_style() -> CellStyle {
    CellStyle {
        font: Font {
            height: 240,
            weight: FONT_WEIGHT_BOLD,
            italic: false,
            underline: 0,
            color_index: COLOR_WHITE,
            name: "Arial".to_string(),
        },
        borders: Borders {
            left_style: BorderStyle::None,
            left_color: COLOR_AUTOMATIC,
            right_style: BorderStyle::None,
            right_color: COLOR_AUTOMATIC,
            top_style: BorderStyle::None,
            top_color: COLOR_AUTOMATIC,
            bottom_style: BorderStyle::Thin,
            bottom_color: COLOR_BLACK,
        },
        fill: Fill {
            pattern: FillPattern::Solid,
            foreground_color: COLOR_BLACK,
            background_color: COLOR_BLACK,
        },
        h_align: HorizontalAlignment::Center,
        v_align: VerticalAlignment::Center,
        text_wrap: true,
        number_format: None,
    }
}

fn make_currency_style() -> CellStyle {
    CellStyle {
        font: Font {
            height: 200,
            weight: FONT_WEIGHT_BOLD,
            italic: false,
            underline: 0,
            color_index: COLOR_BLACK,
            name: "Arial".to_string(),
        },
        borders: Borders {
            left_style: BorderStyle::Thin,
            left_color: COLOR_BLACK,
            right_style: BorderStyle::Thin,
            right_color: COLOR_BLACK,
            top_style: BorderStyle::Thin,
            top_color: COLOR_BLACK,
            bottom_style: BorderStyle::Thin,
            bottom_color: COLOR_BLACK,
        },
        fill: Fill {
            pattern: FillPattern::None,
            foreground_color: COLOR_AUTOMATIC,
            background_color: COLOR_AUTOMATIC,
        },
        h_align: HorizontalAlignment::Right,
        v_align: VerticalAlignment::Bottom,
        text_wrap: false,
        number_format: Some("\"$\"#,##0.00_);(\"$\"#,##0.00)".to_string()),
    }
}

fn make_percent_style() -> CellStyle {
    CellStyle {
        font: Font {
            height: 200,
            weight: FONT_WEIGHT_BOLD,
            italic: false,
            underline: 0,
            color_index: COLOR_RED,
            name: "Arial".to_string(),
        },
        borders: Borders {
            left_style: BorderStyle::None,
            left_color: COLOR_AUTOMATIC,
            right_style: BorderStyle::None,
            right_color: COLOR_AUTOMATIC,
            top_style: BorderStyle::None,
            top_color: COLOR_AUTOMATIC,
            bottom_style: BorderStyle::None,
            bottom_color: COLOR_AUTOMATIC,
        },
        fill: Fill {
            pattern: FillPattern::None,
            foreground_color: COLOR_AUTOMATIC,
            background_color: COLOR_AUTOMATIC,
        },
        h_align: HorizontalAlignment::Right,
        v_align: VerticalAlignment::Bottom,
        text_wrap: false,
        number_format: Some("0.00%".to_string()),
    }
}

fn make_date_style() -> CellStyle {
    CellStyle {
        font: Font {
            height: 200,
            weight: FONT_WEIGHT_NORMAL,
            italic: false,
            underline: 0,
            color_index: COLOR_BLACK,
            name: "Calibri".to_string(),
        },
        borders: Borders {
            left_style: BorderStyle::Thin,
            left_color: COLOR_BLACK,
            right_style: BorderStyle::Thin,
            right_color: COLOR_BLACK,
            top_style: BorderStyle::Thin,
            top_color: COLOR_BLACK,
            bottom_style: BorderStyle::Thin,
            bottom_color: COLOR_BLACK,
        },
        fill: Fill {
            pattern: FillPattern::None,
            foreground_color: COLOR_AUTOMATIC,
            background_color: COLOR_AUTOMATIC,
        },
        h_align: HorizontalAlignment::Left,
        v_align: VerticalAlignment::Center,
        text_wrap: false,
        number_format: Some("m/d/yy".to_string()),
    }
}

fn make_border_demo_style() -> CellStyle {
    CellStyle {
        font: Font {
            height: 200,
            weight: FONT_WEIGHT_BOLD,
            italic: false,
            underline: 0,
            color_index: COLOR_BLACK,
            name: "Arial".to_string(),
        },
        borders: Borders {
            left_style: BorderStyle::Thick,
            left_color: COLOR_RED,
            right_style: BorderStyle::Thick,
            right_color: COLOR_RED,
            top_style: BorderStyle::Thick,
            top_color: COLOR_RED,
            bottom_style: BorderStyle::Thick,
            bottom_color: COLOR_RED,
        },
        fill: Fill {
            pattern: FillPattern::Solid,
            foreground_color: COLOR_WHITE,
            background_color: COLOR_WHITE,
        },
        h_align: HorizontalAlignment::Center,
        v_align: VerticalAlignment::Center,
        text_wrap: true,
        number_format: None,
    }
}

fn make_fill_demo_style() -> CellStyle {
    CellStyle {
        font: Font {
            height: 200,
            weight: FONT_WEIGHT_BOLD,
            italic: false,
            underline: 0,
            color_index: COLOR_BLACK,
            name: "Arial".to_string(),
        },
        borders: Borders {
            left_style: BorderStyle::Thin,
            left_color: COLOR_BLACK,
            right_style: BorderStyle::Thin,
            right_color: COLOR_BLACK,
            top_style: BorderStyle::Thin,
            top_color: COLOR_BLACK,
            bottom_style: BorderStyle::Thin,
            bottom_color: COLOR_BLACK,
        },
        fill: Fill {
            pattern: FillPattern::Solid,
            foreground_color: COLOR_RED,
            background_color: COLOR_RED,
        },
        h_align: HorizontalAlignment::Center,
        v_align: VerticalAlignment::Center,
        text_wrap: false,
        number_format: None,
    }
}

fn make_alignment_demo_style() -> CellStyle {
    CellStyle {
        font: Font {
            height: 200,
            weight: FONT_WEIGHT_BOLD,
            italic: false,
            underline: 0,
            color_index: COLOR_BLACK,
            name: "Arial".to_string(),
        },
        borders: Borders {
            left_style: BorderStyle::Thin,
            left_color: COLOR_BLACK,
            right_style: BorderStyle::Thin,
            right_color: COLOR_BLACK,
            top_style: BorderStyle::Thin,
            top_color: COLOR_BLACK,
            bottom_style: BorderStyle::Thin,
            bottom_color: COLOR_BLACK,
        },
        fill: Fill {
            pattern: FillPattern::None,
            foreground_color: COLOR_AUTOMATIC,
            background_color: COLOR_AUTOMATIC,
        },
        h_align: HorizontalAlignment::Center,
        v_align: VerticalAlignment::Center,
        text_wrap: true,
        number_format: None,
    }
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Styles")?;

    let header_style_id = writer.add_cell_style(make_header_style());
    writer.write_string_with_format(sheet, 0, 0, "Feature", header_style_id)?;
    writer.write_string_with_format(sheet, 0, 1, "Sample", header_style_id)?;

    let currency_id = writer.add_cell_style(make_currency_style());
    writer.write_string_with_format(sheet, 1, 0, "Currency", currency_id)?;
    writer.write_number_with_format(sheet, 1, 1, 1234.56, currency_id)?;

    let percent_id = writer.add_cell_style(make_percent_style());
    writer.write_string_with_format(sheet, 2, 0, "Percent", percent_id)?;
    writer.write_number_with_format(sheet, 2, 1, 0.1234, percent_id)?;

    let date_id = writer.add_cell_style(make_date_style());
    writer.write_string_with_format(sheet, 3, 0, "Date", date_id)?;
    writer.write_number_with_format(sheet, 3, 1, 45123.0, date_id)?;

    let border_id = writer.add_cell_style(make_border_demo_style());
    writer.write_string_with_format(sheet, 4, 0, "Borders", border_id)?;
    writer.write_string_with_format(sheet, 4, 1, "Thick red box", border_id)?;

    let fill_id = writer.add_cell_style(make_fill_demo_style());
    writer.write_string_with_format(sheet, 5, 0, "Fill", fill_id)?;
    writer.write_string_with_format(sheet, 5, 1, "Red background", fill_id)?;

    let align_id = writer.add_cell_style(make_alignment_demo_style());
    writer.write_string_with_format(sheet, 6, 0, "Alignment", align_id)?;
    writer.write_string_with_format(sheet, 6, 1, "Centered and wrapped", align_id)?;

    writer.write_string(sheet, 8, 0, "Merged A9:B9")?;
    writer.merge_cells(sheet, 8, 8, 0, 1)?;

    writer.write_string(sheet, 10, 0, "Status (with validation)")?;
    writer.write_string(sheet, 11, 0, "Value:")?;

    let dv = DataValidation {
        range: litchi_xls::writer::DataValidationRange::new(11, 11, 1, 1)?,
        validation_type: DataValidationType::List {
            values: vec![
                "Not Started".to_string(),
                "In Progress".to_string(),
                "Completed".to_string(),
            ],
        },
        show_input_message: true,
        input_title: Some("Select Status".to_string()),
        input_message: Some("Choose a status from the list".to_string()),
        show_error_alert: true,
        error_title: Some("Invalid Status".to_string()),
        error_message: Some("Please select a value from the list".to_string()),
    };

    writer.add_data_validation(sheet, dv)?;

    let status_cf = ConditionalFormat {
        first_row: 11,
        last_row: 11,
        first_col: 1,
        last_col: 1,
        format_type: ConditionalFormatType::Formula {
            // Highlight when the selected status is "Completed"
            formula: "B12=\"Completed\"".to_string(),
        },
        pattern: Some(ConditionalPattern {
            pattern: FillPattern::Solid,
            foreground_color: COLOR_RED,
            background_color: COLOR_RED,
        }),
    };

    writer.add_conditional_format(sheet, status_cf)?;

    // Freeze the header row and first column so they remain visible while scrolling.
    writer.freeze_panes(sheet, FrozenPanes::new(Row::new(1)?, Column::new(1)?))?;

    // Adjust column widths for the two main columns.
    //
    // Column indices are 0-based (0 = column A, 1 = column B).
    // Width is specified in characters (Excel UI units), internally
    // converted to BIFF8 1/256-character units for COLINFO.
    writer.set_column_width(sheet, Column::new(0)?, 18.0)?;
    writer.set_column_width(sheet, Column::new(1)?, 32.0)?;

    // Make the header row taller so it stands out.
    //
    // Row indices are 0-based (0 = first row), and the height is in
    // typographic points, converted to twips (1/20 point) for ROW.
    writer.set_row_height(sheet, Row::new(0)?, 24.0)?;

    // Make the alignment demo row taller to better show wrapped text.
    writer.set_row_height(sheet, Row::new(6)?, 30.0)?;

    writer.save("styled_output.xls")?;
    Ok(())
}
