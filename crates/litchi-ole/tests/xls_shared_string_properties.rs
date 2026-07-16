use litchi_ole::xls::XlsWorkbook;
use std::fs::File;
use std::path::PathBuf;

fn poi_fixture(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../3rdparty/poi/test-data/spreadsheet")
        .join(name)
}

#[test]
fn worksheet_exposes_rich_shared_string_properties_by_index() {
    let workbook = XlsWorkbook::new(File::open(poi_fixture("duprich2.xls")).unwrap()).unwrap();
    let worksheet = workbook.xls_worksheet(0).unwrap();
    let shared_strings = worksheet.shared_strings().unwrap();

    let rich_entry_count = (0..shared_strings.len())
        .filter(|&index| {
            worksheet
                .shared_string_properties(index as u32)
                .is_some()
        })
        .count();

    assert!(rich_entry_count > 0);
}

#[test]
fn worksheet_resolves_properties_for_duplicate_rich_text_cell() {
    let workbook = XlsWorkbook::new(File::open(poi_fixture("duprich1.xls")).unwrap()).unwrap();
    let worksheet = workbook.xls_worksheet(1).unwrap();
    let cells = [
        worksheet.get_cell(0, 8).unwrap(),
        worksheet.get_cell(1, 8).unwrap(),
    ];

    let (cell, properties) = cells
        .into_iter()
        .find_map(|cell| {
            worksheet
                .shared_string_properties_for_cell(cell)
                .map(|properties| (cell, properties))
        })
        .expect("one duplicate string cell should retain rich-text properties");
    let index = cell.shared_string_index().unwrap();
    let indexed_properties = worksheet.shared_string_properties(index).unwrap();

    assert!(std::ptr::eq(properties, indexed_properties));
}
