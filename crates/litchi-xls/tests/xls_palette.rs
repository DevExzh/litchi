use litchi_xls::Workbook;
use std::fs::File;
use std::path::PathBuf;

fn poi_fixture(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/poi/test-data/spreadsheet")
        .join(name)
}

#[test]
fn reads_palette_used_by_poi_colour_fixture() {
    let workbook =
        Workbook::new(File::open(poi_fixture("SimpleWithColours.xls")).unwrap()).unwrap();
    let palette = workbook.palette();

    assert!(palette.is_custom());
    assert_eq!(palette.palette_colors().len(), 56);
    let black = palette.color(0x08).unwrap();
    let yellow = palette.color(0x0d).unwrap();
    assert_eq!((black.red(), black.green(), black.blue()), (0, 0, 0));
    assert_eq!((yellow.red(), yellow.green(), yellow.blue()), (255, 255, 0));
}

#[test]
fn supplies_default_palette_for_poi_cell_colour_regression_fixture() {
    let workbook = Workbook::new(File::open(poi_fixture("45492.xls")).unwrap()).unwrap();
    let palette = workbook.palette();

    assert!(!palette.is_custom());
    assert_eq!(palette.palette_colors().len(), 56);
    let red = palette.color(0x02).unwrap();
    assert_eq!((red.red(), red.green(), red.blue()), (255, 0, 0));
    assert_eq!(palette.color(0x40), None);
}
