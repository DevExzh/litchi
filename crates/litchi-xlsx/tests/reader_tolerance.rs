use litchi_xlsx::Workbook;
use std::path::PathBuf;

fn fixture(relative: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../..")
        .join(relative)
}

#[test]
fn standalone_reader_accepts_producer_external_link_variants() {
    for name in [
        "test-data/ooxml/xlsx/external-link-path-missing.xlsx",
        "test-data/ooxml/xlsx/external-link-path-startup.xlsx",
    ] {
        let workbook =
            Workbook::open(fixture(name)).unwrap_or_else(|error| panic!("{name}: {error}"));
        assert_eq!(workbook.len(), 1, "{name}");
    }
}

#[test]
fn standalone_reader_retains_sheet_names_from_tolerated_input() {
    let workbook = Workbook::open(fixture("test-data/ooxml/xlsx/sheet-state-show.xlsx")).unwrap();
    let names = workbook
        .sheets()
        .map(|sheet| sheet.name().to_owned())
        .collect::<Vec<_>>();
    assert_eq!(names, ["стр1", "стр2"]);
}

#[test]
fn standalone_reader_opens_duplicate_names_and_malformed_string_hints() {
    let duplicate =
        Workbook::open(fixture("test-data/ooxml/xlsx/duplicate-defined-names.xlsx")).unwrap();
    assert!(!duplicate.is_empty());
    let strings = Workbook::open(fixture(
        "test-data/ooxml/xlsx/shared-strings-malformed-count.xlsx",
    ))
    .unwrap();
    assert!(!strings.is_empty());
}
