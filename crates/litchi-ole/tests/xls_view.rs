use litchi_ole::xls::{XlsPaneType, XlsWorkbook};
use std::fs::File;
use std::path::{Path, PathBuf};

fn poi_fixture(name: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../3rdparty/poi/test-data/spreadsheet")
        .join(name)
}

fn libreoffice_fixture(name: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../3rdparty/libreoffice-core/sc/qa/unit/data/xls")
        .join(name)
}

#[test]
fn reads_poi_zoom_panes_and_selections() {
    let workbook = XlsWorkbook::new(File::open(poi_fixture("41139.xls")).unwrap()).unwrap();
    let view = workbook.xls_worksheet(0).unwrap().worksheet_view().unwrap();
    assert_eq!(view.zoom_fraction(), Some((3, 4)));
    assert_eq!(view.pane().unwrap().active_pane(), XlsPaneType::LowerRight);
    assert_eq!(view.selections().len(), 4);
}

#[test]
fn reads_libreoffice_worksheet_view() {
    let workbook =
        XlsWorkbook::new(File::open(libreoffice_fixture("formats.xls")).unwrap()).unwrap();
    let view = workbook.xls_worksheet(0).unwrap().worksheet_view().unwrap();
    assert!(view.shows_gridlines());
    assert!(view.shows_row_column_headers());
    assert!(!view.selections().is_empty());
}
