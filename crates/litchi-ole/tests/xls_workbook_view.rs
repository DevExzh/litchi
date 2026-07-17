use std::fs::File;
use std::io::Cursor;
use std::path::PathBuf;

use litchi_ole::xls::writer::{XlsWorkbookWindowOptions, XlsWriter};
use litchi_ole::xls::XlsWorkbook;

#[test]
fn workbook_window_and_sheet_ids_round_trip() {
    let mut writer = XlsWriter::new();
    writer.add_worksheet("One").unwrap();
    writer.add_worksheet("Two").unwrap();
    writer.add_worksheet("Three").unwrap();
    writer.set_workbook_window(XlsWorkbookWindowOptions {
        horizontal_position_twips: -120,
        vertical_position_twips: 240,
        width_twips: 12000,
        height_twips: 8000,
        hidden: true,
        minimized: true,
        very_hidden: true,
        show_horizontal_scrollbar: false,
        show_vertical_scrollbar: true,
        show_sheet_tabs: false,
        group_dates_in_autofilter: false,
        active_sheet_index: 2,
        first_visible_sheet_index: 1,
        selected_sheet_count: 2,
        sheet_tab_ratio_per_mille: 725,
    }).unwrap();

    let mut bytes = Cursor::new(Vec::new());
    writer.write_to(&mut bytes).unwrap();
    let workbook = XlsWorkbook::new(Cursor::new(bytes.into_inner())).unwrap();
    let view = workbook.workbook_view();
    assert_eq!(view.sheet_ids(), &[1, 2, 3]);
    let window = view.primary_window().unwrap();
    assert_eq!(window.horizontal_position_twips(), -120);
    assert_eq!(window.vertical_position_twips(), 240);
    assert_eq!(window.width_twips(), 12000);
    assert_eq!(window.height_twips(), 8000);
    assert!(window.hidden());
    assert!(window.minimized());
    assert!(window.very_hidden());
    assert!(!window.shows_horizontal_scrollbar());
    assert!(window.shows_vertical_scrollbar());
    assert!(!window.shows_sheet_tabs());
    assert!(!window.groups_dates_in_autofilter());
    assert_eq!(window.active_sheet_index(), 2);
    assert_eq!(window.first_visible_sheet_index(), 1);
    assert_eq!(window.selected_sheet_count(), 2);
    assert_eq!(window.sheet_tab_ratio_per_mille(), 725);
    assert!(!workbook.xls_worksheet(0).unwrap().worksheet_view().unwrap().is_selected());
    assert!(workbook.xls_worksheet(1).unwrap().worksheet_view().unwrap().is_selected());
    assert!(workbook.xls_worksheet(2).unwrap().worksheet_view().unwrap().is_selected());
}

#[test]
fn reads_poi_simple_workbook_window_and_sheet_ids() {
    let fixture = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../3rdparty/poi/test-data/spreadsheet/Simple.xls");
    let workbook = XlsWorkbook::new(File::open(fixture).unwrap()).unwrap();
    let view = workbook.workbook_view();
    assert_eq!(view.sheet_ids(), &[1, 2, 3]);
    let window = view.primary_window().unwrap();
    assert_eq!(window.horizontal_position_twips(), 120);
    assert_eq!(window.vertical_position_twips(), 120);
    assert_eq!(window.width_twips(), 15135);
    assert_eq!(window.height_twips(), 9300);
    assert!(window.shows_horizontal_scrollbar());
    assert!(window.shows_vertical_scrollbar());
    assert!(window.shows_sheet_tabs());
    assert_eq!(window.active_sheet_index(), 0);
    assert_eq!(window.first_visible_sheet_index(), 0);
    assert_eq!(window.selected_sheet_count(), 1);
    assert_eq!(window.sheet_tab_ratio_per_mille(), 600);
    assert!(workbook.xls_worksheet(0).unwrap().worksheet_view().unwrap().is_selected());
    assert!(!workbook.xls_worksheet(1).unwrap().worksheet_view().unwrap().is_selected());
    assert!(!workbook.xls_worksheet(2).unwrap().worksheet_view().unwrap().is_selected());
}

#[test]
fn writer_rejects_window1_window2_selection_disagreement() {
    use litchi_ole::xls::writer::XlsWorksheetViewOptions;

    let mut writer = XlsWriter::new();
    writer.add_worksheet("One").unwrap();
    let second = writer.add_worksheet("Two").unwrap();
    writer
        .set_worksheet_view(
            second,
            XlsWorksheetViewOptions { selected: true, ..Default::default() },
        )
        .unwrap();
    assert!(writer.write_to(&mut Cursor::new(Vec::new())).is_err());
}

#[test]
fn reads_poi_zero_based_sheet_identifiers() {
    let fixture = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../3rdparty/poi/test-data/spreadsheet/duprich1.xls");
    let workbook = XlsWorkbook::new(File::open(fixture).unwrap()).unwrap();
    assert_eq!(workbook.workbook_view().sheet_ids(), &[0, 1, 2]);
}

#[test]
fn writer_rejects_invalid_window_and_tab_references() {
    let mut writer = XlsWriter::new();
    writer.add_worksheet("One").unwrap();
    assert!(writer.set_workbook_window(XlsWorkbookWindowOptions {
        width_twips: 0,
        ..XlsWorkbookWindowOptions::default()
    }).is_err());
    assert!(writer.set_workbook_window(XlsWorkbookWindowOptions {
        sheet_tab_ratio_per_mille: 1001,
        ..XlsWorkbookWindowOptions::default()
    }).is_err());
    writer.set_workbook_window(XlsWorkbookWindowOptions {
        active_sheet_index: 1,
        ..XlsWorkbookWindowOptions::default()
    }).unwrap();
    assert!(writer.write_to(&mut Cursor::new(Vec::new())).is_err());
}
