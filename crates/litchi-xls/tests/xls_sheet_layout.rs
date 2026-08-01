use std::fs::File;
use std::io::Cursor;
use std::path::PathBuf;

use litchi_xls::XlsWorkbook;
use litchi_xls::writer::{XlsWorksheetLayoutOptions, XlsWriter};

#[test]
fn worksheet_layout_round_trip() {
    let mut writer = XlsWriter::new();
    let sheet = writer.add_worksheet("Layout").unwrap();
    writer
        .set_worksheet_layout(
            sheet,
            XlsWorksheetLayoutOptions {
                default_row_height_twips: 360,
                empty_rows_hidden: true,
                default_row_height_unsynced: true,
                thick_top_border: true,
                thick_bottom_border: true,
                default_column_width_chars: 12,
                max_row_outline_level: 3,
                max_column_outline_level: 2,
                row_gutter_width: 53,
                column_gutter_height: 41,
                show_automatic_page_breaks: false,
                apply_styles_to_outlines: true,
                summary_rows_below: false,
                summary_columns_right: false,
                fit_to_page: true,
                synchronize_horizontal_scrolling: true,
                synchronize_vertical_scrolling: true,
                alternate_expression_evaluation: true,
                alternate_formula_entry: true,
            },
        )
        .unwrap();

    let mut bytes = Cursor::new(Vec::new());
    writer.write_to(&mut bytes).unwrap();
    let workbook = XlsWorkbook::new(Cursor::new(bytes.into_inner())).unwrap();
    let layout = workbook.xls_worksheet(0).unwrap().sheet_layout();
    assert_eq!(layout.default_row_height_twips(), 360);
    assert!(layout.empty_rows_hidden());
    assert!(layout.default_row_height_unsynced());
    assert!(layout.thick_top_border());
    assert!(layout.thick_bottom_border());
    assert_eq!(layout.default_column_width_chars(), 12);
    assert_eq!(layout.max_row_outline_level(), 3);
    assert_eq!(layout.max_column_outline_level(), 2);
    assert_eq!(layout.row_gutter_width(), 53);
    assert_eq!(layout.column_gutter_height(), 41);
    assert!(!layout.show_automatic_page_breaks());
    assert!(layout.apply_styles_to_outlines());
    assert!(!layout.summary_rows_below());
    assert!(!layout.summary_columns_right());
    assert!(layout.fit_to_page());
    assert!(layout.synchronize_horizontal_scrolling());
    assert!(layout.synchronize_vertical_scrolling());
    assert!(layout.alternate_expression_evaluation());
    assert!(layout.alternate_formula_entry());
}

#[test]
fn reads_poi_column_width_fixture_defaults() {
    let fixture = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/poi/test-data/spreadsheet/colwidth.xls");
    let workbook = XlsWorkbook::new(File::open(fixture).unwrap()).unwrap();
    let layout = workbook.xls_worksheet(0).unwrap().sheet_layout();
    assert_eq!(layout.default_row_height_twips(), 255);
    assert_eq!(layout.default_column_width_chars(), 8);
    assert_eq!(layout.max_row_outline_level(), 0);
    assert_eq!(layout.max_column_outline_level(), 0);
    assert!(layout.show_automatic_page_breaks());
    assert!(layout.summary_rows_below());
    assert!(layout.summary_columns_right());
}

#[test]
fn writer_rejects_invalid_layout_bounds() {
    let mut writer = XlsWriter::new();
    let sheet = writer.add_worksheet("Layout").unwrap();
    assert!(
        writer
            .set_worksheet_layout(
                sheet,
                XlsWorksheetLayoutOptions {
                    default_row_height_twips: 0,
                    ..XlsWorksheetLayoutOptions::default()
                }
            )
            .is_err()
    );
    assert!(
        writer
            .set_worksheet_layout(
                sheet,
                XlsWorksheetLayoutOptions {
                    default_column_width_chars: 256,
                    ..XlsWorksheetLayoutOptions::default()
                }
            )
            .is_err()
    );
    assert!(
        writer
            .set_worksheet_layout(
                sheet,
                XlsWorksheetLayoutOptions {
                    max_row_outline_level: 8,
                    ..XlsWorksheetLayoutOptions::default()
                }
            )
            .is_err()
    );
}
