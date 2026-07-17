use std::io::Cursor;

use litchi_ole::xls::writer::{
    XlsViewScale, XlsWorksheetPaneOptions, XlsWorksheetSelectionOptions,
    XlsWorksheetViewOptions, XlsWriter,
};
use litchi_ole::xls::{XlsPaneType, XlsSelectionRange, XlsWorkbook};

#[test]
fn writes_and_reads_typed_view_state() {
    let mut writer = XlsWriter::new();
    let sheet = writer.add_worksheet("View").unwrap();
    let pane = XlsWorksheetPaneOptions::split(1_200, 800, 7, 4, XlsPaneType::LowerRight).unwrap();
    let options = XlsWorksheetViewOptions {
        show_formulas: true,
        show_gridlines: false,
        first_visible_row: 2,
        first_visible_column: 1,
        gridline_color_index: Some(8),
        normal_zoom_percent: Some(125),
        scale: Some(XlsViewScale::new(5, 4).unwrap()),
        pane: Some(pane),
        selections: vec![XlsWorksheetSelectionOptions {
            pane: XlsPaneType::LowerRight,
            active_row: 8,
            active_column: 5,
            active_range_index: 0,
            ranges: vec![XlsSelectionRange::new(8, 10, 5, 6)],
        }],
        ..XlsWorksheetViewOptions::default()
    };
    writer.set_worksheet_view(sheet, options).unwrap();

    let mut bytes = Cursor::new(Vec::new());
    writer.write_to(&mut bytes).unwrap();
    let workbook = XlsWorkbook::new(Cursor::new(bytes.into_inner())).unwrap();
    let view = workbook.xls_worksheet(0).unwrap().worksheet_view().unwrap();
    assert!(view.shows_formulas());
    assert!(!view.shows_gridlines());
    assert_eq!(view.first_visible_row(), 2);
    assert_eq!(view.first_visible_column(), 1);
    assert_eq!(view.gridline_color_index(), 8);
    assert_eq!(view.normal_zoom_percent(), Some(125));
    assert_eq!(view.zoom_fraction(), Some((5, 4)));
    assert_eq!(view.pane().unwrap().active_pane(), XlsPaneType::LowerRight);
    assert_eq!(view.selections()[0].ranges()[0], XlsSelectionRange::new(8, 10, 5, 6));
}

#[test]
fn compatibility_freeze_and_zoom_round_trip() {
    let mut writer = XlsWriter::new();
    let sheet = writer.add_worksheet("Compat").unwrap();
    writer.freeze_panes(sheet, 7, 5).unwrap();
    writer.set_zoom(sheet, 3, 4).unwrap();
    let mut bytes = Cursor::new(Vec::new());
    writer.write_to(&mut bytes).unwrap();
    let workbook = XlsWorkbook::new(Cursor::new(bytes.into_inner())).unwrap();
    let view = workbook.xls_worksheet(0).unwrap().worksheet_view().unwrap();
    assert!(view.has_frozen_panes());
    assert!(view.is_frozen_without_split());
    assert_eq!(view.zoom_fraction(), Some((3, 4)));
    assert_eq!(view.pane().unwrap().active_pane(), XlsPaneType::LowerRight);
}
