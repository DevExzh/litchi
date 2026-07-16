use litchi_ole::xls::{XlsExtendedFormatKind, XlsWorkbook};
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

fn assert_effective_links<R: std::io::Read + std::io::Seek>(workbook: &XlsWorkbook<R>) {
    let mut cell_count = 0;
    for xf in workbook.extended_formats() {
        let effective = workbook.effective_extended_format(xf.index()).unwrap();
        if let XlsExtendedFormatKind::Cell { parent_style_xf } = xf.kind() {
            assert_eq!(effective.parent_style().unwrap().index(), parent_style_xf);
            assert!(matches!(
                effective.parent_style().unwrap().kind(),
                XlsExtendedFormatKind::Style
            ));
            cell_count += 1;
        } else {
            assert!(effective.parent_style().is_none());
        }
        assert_ne!(effective.font_index(), 4);
    }
    assert!(cell_count > 0);
}

#[test]
fn resolves_poi_xf_inheritance() {
    let workbook = XlsWorkbook::new(File::open(poi_fixture("Formatting.xls")).unwrap()).unwrap();
    assert_effective_links(&workbook);
}

#[test]
fn resolves_libreoffice_xf_inheritance() {
    let workbook =
        XlsWorkbook::new(File::open(libreoffice_fixture("formats.xls")).unwrap()).unwrap();
    assert_effective_links(&workbook);
}
