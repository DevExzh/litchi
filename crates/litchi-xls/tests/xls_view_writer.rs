use std::io::Cursor;

use litchi_xls::Workbook;
use litchi_xls::view::{PaneType, Range};
use litchi_xls::writer::view::{Pane, Scale, Selection, View};
use litchi_xls::writer::{Column, FrozenPanes, Row, Writer};

#[test]
fn writes_and_reads_typed_view_state() {
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("View").unwrap();
    let pane = Pane::split(1_200, 800, 7, 4, PaneType::LowerRight).unwrap();
    let selection = Selection::new(
        PaneType::LowerRight,
        8,
        5,
        0,
        vec![Range::new(8, 10, 5, 6).unwrap()],
    )
    .unwrap();
    let mut view = View::default();
    view.formulas(true).gridlines(false);
    view.origin(2, 1).unwrap();
    view.grid_color(Some(8)).unwrap();
    view.normal_zoom(Some(125)).unwrap();
    view.put_scale(Some(Scale::new(5, 4).unwrap()));
    view.put_pane(pane, vec![selection]).unwrap();
    writer.put_view(sheet, view).unwrap();

    let mut bytes = Cursor::new(Vec::new());
    writer.write_to(&mut bytes).unwrap();
    let workbook = Workbook::new(Cursor::new(bytes.into_inner())).unwrap();
    let view = workbook.xls_worksheet(0).unwrap().worksheet_view().unwrap();
    assert!(view.shows_formulas());
    assert!(!view.shows_gridlines());
    assert_eq!(view.first_visible_row(), 2);
    assert_eq!(view.first_visible_column(), 1);
    assert_eq!(view.gridline_color_index(), 8);
    assert_eq!(view.normal_zoom_percent(), Some(125));
    assert_eq!(view.zoom_fraction(), Some((5, 4)));
    assert_eq!(view.pane().unwrap().active_pane(), PaneType::LowerRight);
    assert_eq!(
        view.selections()[0].ranges()[0],
        Range::new(8, 10, 5, 6).unwrap()
    );
}

#[test]
fn freeze_and_scale_round_trip() {
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Compat").unwrap();
    writer
        .freeze_panes(
            sheet,
            FrozenPanes::new(Row::new(7).unwrap(), Column::new(5).unwrap()),
        )
        .unwrap();
    writer
        .put_scale(sheet, Some(Scale::new(3, 4).unwrap()))
        .unwrap();
    let mut bytes = Cursor::new(Vec::new());
    writer.write_to(&mut bytes).unwrap();
    let workbook = Workbook::new(Cursor::new(bytes.into_inner())).unwrap();
    let view = workbook.xls_worksheet(0).unwrap().worksheet_view().unwrap();
    assert!(view.has_frozen_panes());
    assert!(view.is_frozen_without_split());
    assert_eq!(view.zoom_fraction(), Some((3, 4)));
    assert_eq!(view.pane().unwrap().active_pane(), PaneType::LowerRight);
}
