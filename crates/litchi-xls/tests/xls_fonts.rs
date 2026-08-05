use litchi_xls::Workbook;
use std::fs::File;
use std::path::PathBuf;

fn poi_fixture(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/poi/test-data/spreadsheet")
        .join(name)
}

fn libreoffice_fixture(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/libreoffice-core/sc/qa/unit/data/xls")
        .join(name)
}

#[test]
fn reads_formatting_fixture_font_table_and_reserved_index_gap() {
    let workbook = Workbook::new(File::open(poi_fixture("Formatting.xls")).unwrap()).unwrap();

    assert!(workbook.fonts().len() >= 5);
    for index in 0..=3 {
        assert_eq!(workbook.font(index).unwrap().index(), index);
    }
    assert!(workbook.font(4).is_none());

    let formatted = workbook.font(5).unwrap();
    assert_eq!(formatted.index(), 5);
    assert!(!formatted.name().is_empty());

    let explicitly_colored = workbook
        .fonts()
        .iter()
        .find(|font| workbook.palette().color(font.color_index()).is_some())
        .expect("Formatting.xls should contain an explicitly colored font");
    assert_eq!(
        workbook.font_color(explicitly_colored.index()),
        workbook.palette().color(explicitly_colored.color_index())
    );

    let xf = workbook
        .extended_formats()
        .iter()
        .find(|xf| xf.font_index() == formatted.index())
        .expect("Formatting.xls should reference its additional font from an XF");
    assert_eq!(workbook.extended_format_font(xf), Some(formatted));
}

#[test]
fn resolves_duprich_format_run_font_indices() {
    let workbook = Workbook::new(File::open(poi_fixture("duprich2.xls")).unwrap()).unwrap();
    let worksheet = workbook.xls_worksheet(0).unwrap();
    let shared_strings = worksheet.shared_strings().unwrap();
    let mut run_count = 0;

    for index in 0..shared_strings.len() {
        let Some(properties) = worksheet.shared_string_properties(index as u32) else {
            continue;
        };
        for run in &properties.formatting_runs {
            assert_ne!(run.font_index, 4);
            assert!(workbook.font(run.font_index).is_some());
            run_count += 1;
        }
    }

    assert!(run_count > 0);
}

#[test]
fn resolves_libreoffice_xf_font_references() {
    let workbook = Workbook::new(File::open(libreoffice_fixture("formats.xls")).unwrap()).unwrap();
    assert!(!workbook.extended_formats().is_empty());
    assert!(
        workbook
            .extended_formats()
            .iter()
            .all(|xf| workbook.extended_format_font(xf).is_some())
    );
}

#[test]
fn reads_poi_compressed_font_name_fixture() {
    let workbook =
        Workbook::new(File::open(poi_fixture("SimpleWithColours.xls")).unwrap()).unwrap();
    assert!(workbook.fonts().iter().all(|font| !font.name().is_empty()));
}
