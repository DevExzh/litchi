use litchi_odf::Spreadsheet;

#[test]
fn reads_libreoffice_external_hyperlink_when_fixture_is_available() {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../3rdparty/libreoffice-core/sc/qa/unit/data/ods/external_hyperlink.ods");
    if !path.exists() {
        return;
    }

    let mut spreadsheet = Spreadsheet::open(path).unwrap();
    let sheets = spreadsheet.sheets().unwrap();
    let links: Vec<_> = sheets
        .iter()
        .flat_map(|sheet| sheet.rows.iter())
        .flat_map(|row| row.cells.iter())
        .flat_map(|cell| cell.hyperlinks().iter())
        .collect();

    assert_eq!(links.len(), 1);
    assert_eq!(
        links[0].href(),
        "../../Desktop/%23folder/test.ods#Sheet2.B10"
    );
    assert_eq!(links[0].text(), "hyperlink with #");
}
