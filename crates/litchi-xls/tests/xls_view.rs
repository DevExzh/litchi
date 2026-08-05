use litchi_xls::Workbook;
use litchi_xls::view::{PaneType, Range};
use std::fs::File;
use std::path::{Path, PathBuf};

fn poi_fixture(name: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/poi/test-data/spreadsheet")
        .join(name)
}

fn libreoffice_fixture(name: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/libreoffice-core/sc/qa/unit/data/xls")
        .join(name)
}

#[test]
fn reads_poi_zoom_panes_and_selections() {
    let workbook = Workbook::new(File::open(poi_fixture("41139.xls")).unwrap()).unwrap();
    let view = workbook.xls_worksheet(0).unwrap().worksheet_view().unwrap();
    assert_eq!(view.zoom_fraction(), Some((3, 4)));
    assert_eq!(view.pane().unwrap().active_pane(), PaneType::LowerRight);
    assert_eq!(view.selections().len(), 4);
}

#[test]
fn reads_libreoffice_worksheet_view() {
    let workbook = Workbook::new(File::open(libreoffice_fixture("formats.xls")).unwrap()).unwrap();
    let view = workbook.xls_worksheet(0).unwrap().worksheet_view().unwrap();
    assert!(view.shows_gridlines());
    assert!(view.shows_row_column_headers());
    assert!(!view.selections().is_empty());
}

#[test]
fn rejects_inverted_selection_ranges() {
    assert!(Range::new(9, 8, 0, 0).is_err());
    assert!(Range::new(0, 0, 2, 1).is_err());
}
