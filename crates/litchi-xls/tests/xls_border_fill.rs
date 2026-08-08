use litchi_xls::{BorderStyle, FillPattern, OpenOptions, Workbook};
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
fn reads_poi_border_style_fixture() {
    let workbook =
        Workbook::new(File::open(poi_fixture("55341_CellStyleBorder.xls")).unwrap()).unwrap();
    let sheet = workbook.xls_worksheet(0).unwrap();
    let expected = [
        BorderStyle::Hair,
        BorderStyle::Dotted,
        BorderStyle::DashDotDot,
        BorderStyle::Dashed,
        BorderStyle::Thin,
        BorderStyle::MediumDashDotDot,
        BorderStyle::SlantedDashDot,
        BorderStyle::MediumDashDot,
        BorderStyle::MediumDashed,
        BorderStyle::Medium,
        BorderStyle::Thick,
        BorderStyle::Double,
    ];

    for (index, style) in expected.into_iter().enumerate() {
        let cell = sheet.get_cell(index as u32, index as u32).unwrap();
        let xf = &workbook.extended_formats()[cell.xf_index() as usize];
        assert_eq!(xf.borders().right().style(), style, "cell {index},{index}");
    }
}

#[test]
fn reads_libreoffice_border_fixture() {
    let workbook =
        Workbook::new(File::open(libreoffice_fixture("cell-borders.xls")).unwrap()).unwrap();
    assert!(workbook.extended_formats().iter().any(|xf| {
        let borders = xf.borders();
        borders.left().style() != BorderStyle::None
            || borders.right().style() != BorderStyle::None
            || borders.top().style() != BorderStyle::None
            || borders.bottom().style() != BorderStyle::None
            || borders.diagonal().style() != BorderStyle::None
    }));
}

#[test]
fn reads_poi_fill_fixture() {
    let workbook =
        Workbook::new(File::open(poi_fixture("SimpleWithColours.xls")).unwrap()).unwrap();
    let palette = workbook.palette();
    let fill = workbook
        .extended_formats()
        .iter()
        .map(|xf| xf.fill())
        .find(|fill| fill.pattern() != FillPattern::None)
        .expect("fixture should contain a patterned fill");
    assert!(fill.foreground_color(palette).is_some());
}

#[test]
fn reads_poi_xor_encrypted_style_xf_reserved_bit() {
    let workbook = Workbook::new_with_options(
        File::open(poi_fixture("xor-encryption-abc.xls")).unwrap(),
        OpenOptions::new().with_password("abc"),
    )
    .unwrap();
    assert!(!workbook.extended_formats().is_empty());
}
