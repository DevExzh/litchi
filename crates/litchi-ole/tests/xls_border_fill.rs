use litchi_ole::xls::{XlsBorderStyle, XlsFillPattern, XlsOpenOptions, XlsWorkbook};
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
        XlsWorkbook::new(File::open(poi_fixture("55341_CellStyleBorder.xls")).unwrap()).unwrap();
    let sheet = workbook.xls_worksheet(0).unwrap();
    let expected = [
        XlsBorderStyle::Hair,
        XlsBorderStyle::Dotted,
        XlsBorderStyle::DashDotDot,
        XlsBorderStyle::Dashed,
        XlsBorderStyle::Thin,
        XlsBorderStyle::MediumDashDotDot,
        XlsBorderStyle::SlantedDashDot,
        XlsBorderStyle::MediumDashDot,
        XlsBorderStyle::MediumDashed,
        XlsBorderStyle::Medium,
        XlsBorderStyle::Thick,
        XlsBorderStyle::Double,
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
        XlsWorkbook::new(File::open(libreoffice_fixture("cell-borders.xls")).unwrap()).unwrap();
    assert!(workbook.extended_formats().iter().any(|xf| {
        let borders = xf.borders();
        borders.left().style() != XlsBorderStyle::None
            || borders.right().style() != XlsBorderStyle::None
            || borders.top().style() != XlsBorderStyle::None
            || borders.bottom().style() != XlsBorderStyle::None
            || borders.diagonal().style() != XlsBorderStyle::None
    }));
}

#[test]
fn reads_poi_fill_fixture() {
    let workbook =
        XlsWorkbook::new(File::open(poi_fixture("SimpleWithColours.xls")).unwrap()).unwrap();
    let palette = workbook.palette();
    let fill = workbook
        .extended_formats()
        .iter()
        .map(|xf| xf.fill())
        .find(|fill| fill.pattern() != XlsFillPattern::None)
        .expect("fixture should contain a patterned fill");
    assert!(fill.foreground_color(palette).is_some());
}

#[test]
fn reads_poi_xor_encrypted_style_xf_reserved_bit() {
    let workbook = XlsWorkbook::new_with_options(
        File::open(poi_fixture("xor-encryption-abc.xls")).unwrap(),
        XlsOpenOptions {
            password: Some("abc"),
        },
    )
    .unwrap();
    assert!(!workbook.extended_formats().is_empty());
}
