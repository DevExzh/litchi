//! Reader tests for the BIFF8 custom-view records (`UserBView`,
//! `UserSViewBegin`, `UserSViewEnd`).

use litchi_core::sheet::WorkbookTrait;
use litchi_xls::{
    XlsCustomViewHiddenRows, XlsCustomViewNoteDisplay, XlsObjectDisplayMode, XlsPaneType,
    XlsWorkbook,
};

const EXPECTED_GUID: [u8; 16] = [
    0x98, 0x93, 0x94, 0x42, 0xeb, 0x7b, 0x6b, 0x49, 0xac, 0x82, 0x6b, 0x9c, 0xd9, 0x34, 0x29, 0x8f,
];

fn open_fixture() -> XlsWorkbook<std::io::BufReader<std::fs::File>> {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/ole/xls/WithCustomViews.xls");
    XlsWorkbook::new(std::io::BufReader::new(std::fs::File::open(path).unwrap())).unwrap()
}

#[test]
fn reads_workbook_custom_view() {
    let workbook = open_fixture();
    let views = workbook.custom_views();
    assert_eq!(views.len(), 1);
    let view = &views[0];
    assert_eq!(view.name(), "tech20 - Modo de exibição pessoal");
    assert_eq!(view.guid(), &EXPECTED_GUID);
    assert_eq!(view.active_tab(), Some(1));
    assert_eq!(view.window_position(), (0, 0));
    assert_eq!(view.window_size(), (991, 600));
    assert_eq!(view.tab_ratio(), 601);
    assert!(view.shows_formula_bar());
    assert!(view.shows_status_bar());
    assert_eq!(view.note_display(), XlsCustomViewNoteDisplay::VisualCue);
    assert_eq!(view.object_display(), XlsObjectDisplayMode::ShowAll);
    assert!(view.includes_print_settings());
    assert!(view.includes_hidden_rows_columns_and_filters());
    assert!(view.is_personal_view());
    assert!(!view.is_minimized());
    assert!(view.is_maximized());
}

#[test]
fn reads_sheet_custom_views() {
    let workbook = open_fixture();
    let expected: &[(u32, u16, u16, u16)] = &[
        // (scale, last top-left row, last top-left col, tab id)
        (75, 12, 9, 1),
        (100, 33, 15, 2),
        (100, 33, 15, 3),
    ];
    assert_eq!(workbook.worksheet_count(), expected.len());
    for (index, &(scale, last_row, last_col, tab_id)) in expected.iter().enumerate() {
        let worksheet = workbook.xls_worksheet(index).unwrap();
        let views = worksheet.custom_views();
        assert_eq!(views.len(), 1, "worksheet {index}");
        let view = &views[0];
        assert_eq!(view.end().reserved(), 1);
        let begin = view.begin();
        // Every sheet view is tied to the workbook view by its GUID.
        assert_eq!(begin.guid(), &EXPECTED_GUID);
        assert_eq!(begin.tab_id(), tab_id);
        assert_eq!(begin.scale(), scale);
        assert_eq!(begin.gridline_color(), 64);
        assert_eq!(begin.active_pane(), XlsPaneType::UpperLeft);
        assert!(begin.shows_gridlines());
        assert!(begin.shows_headings());
        assert!(!begin.shows_formulas());
        assert!(!begin.is_frozen());
        assert!(!begin.is_split_vertically());
        assert!(!begin.is_split_horizontally());
        assert_eq!(begin.hidden_rows(), XlsCustomViewHiddenRows::Present);
        assert!(!begin.has_hidden_columns());
        assert!(!begin.is_page_break_preview());
        assert!(!begin.is_page_layout_view());
        let top_left = begin.top_left();
        assert_eq!(top_left.first_row(), 0);
        assert_eq!(top_left.first_col(), 0);
        assert_eq!(top_left.last_row(), last_row);
        assert_eq!(top_left.last_col(), last_col);
        assert_eq!(begin.split_x(), 0.0);
        assert_eq!(begin.split_y(), 0.0);
        assert_eq!(begin.right_pane_col(), u16::MAX);
        assert_eq!(begin.bottom_pane_row(), u16::MAX);
    }
}

#[test]
fn workbook_without_custom_views_reads_empty() {
    let path =
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data/ole/xls/Simple.xls");
    let workbook =
        XlsWorkbook::new(std::io::BufReader::new(std::fs::File::open(path).unwrap())).unwrap();
    assert!(workbook.custom_views().is_empty());
    let worksheet = workbook.xls_worksheet(0).unwrap();
    assert!(worksheet.custom_views().is_empty());
}
