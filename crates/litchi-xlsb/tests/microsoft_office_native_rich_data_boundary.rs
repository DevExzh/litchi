//! Native rich-data package payloads are retained, without a typed XLSB rich-value model.
use litchi_core::sheet::{Cell as _, CellValue, WorkbookTrait, Worksheet};
use litchi_xlsb::{Package, cell_values::Reference};
use sha2::{Digest, Sha256};
use std::collections::BTreeMap;
const INPUT: &[u8] =
    include_bytes!("../../../test-data/microsoft-office/excel/excel-links-objects.xlsb");
fn members(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    let r = litchi_opc::phys_pkg::OwnedPhysPkgReader::from_bytes(
        bytes.to_vec(),
        litchi_opc::SourceLimits::default(),
    )
    .unwrap();
    r.member_names()
        .unwrap()
        .into_iter()
        .map(|name| {
            let b = r.read_member(&name).unwrap();
            (name, b)
        })
        .collect()
}
fn digest(bytes: &[u8]) -> String {
    Sha256::digest(bytes)
        .iter()
        .map(|b| format!("{b:02x}"))
        .collect()
}

#[test]
fn native_xlsb_rich_data_and_cell_metadata_remain_opaque_through_scalar_publication() {
    let p = Package::from_bytes(INPUT).unwrap();
    let before = members(INPUT);
    let rich = before
        .keys()
        .filter(|n| n.starts_with("xl/richData/"))
        .map(String::as_str)
        .collect::<Vec<_>>();
    assert_eq!(
        rich,
        [
            "xl/richData/_rels/richValueRel.xml.rels",
            "xl/richData/rdRichValueTypes.xml",
            "xl/richData/rdrichvalue.xml",
            "xl/richData/rdrichvaluestructure.xml",
            "xl/richData/richValueRel.xml"
        ]
    );
    assert!(
        std::str::from_utf8(&before["xl/richData/rdrichvaluestructure.xml"])
            .unwrap()
            .contains("t=\"_localImage\"")
    );
    assert!(
        std::str::from_utf8(&before["xl/richData/rdrichvalue.xml"])
            .unwrap()
            .contains("<rv s=\"0\"><v>0</v><v>5</v></rv>")
    );
    assert_eq!(
        digest(&before["xl/media/image1.png"]),
        "d594d8f832d3ab38d03e0d06316d2fa5e401217e50e21e01814e593d752da3e8"
    );
    assert_eq!(
        digest(&before["xl/metadata.bin"]),
        "e7f72383dcfcba290b9db702da691c2f584765d6eed83606e70f5b438debfddd"
    );
    assert!(p.to_bytes().unwrap() == INPUT);
    let original = p.workbook().unwrap();
    let sheet = original.worksheet_by_name("Sheet1").unwrap();
    assert!(sheet.get_cell(49, 9).is_none());
    let mut edit = p.edit_cell_values("Sheet1").unwrap();
    edit.insert(
        Reference::new(49, 9).unwrap(),
        litchi_xlsb::cell_values::StyleIndex::new(0).unwrap(),
        litchi_xlsb::cell_values::Value::Number(42.5),
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    let output = commit.package().to_bytes().unwrap();
    let after = members(&output);
    assert_eq!(
        before.keys().collect::<Vec<_>>(),
        after.keys().collect::<Vec<_>>()
    );
    let changed = before
        .keys()
        .filter(|n| before.get(*n) != after.get(*n))
        .map(String::as_str)
        .collect::<Vec<_>>();
    assert_eq!(
        changed,
        ["xl/worksheets/binaryIndex1.bin", "xl/worksheets/sheet1.bin"]
    );
    let f = tempfile::NamedTempFile::with_suffix(".xlsb").unwrap();
    commit.package().save(f.path()).unwrap();
    let reopened = Package::open(f.path()).unwrap();
    assert_eq!(
        reopened
            .workbook()
            .unwrap()
            .worksheet_by_name("Sheet1")
            .unwrap()
            .get_cell(49, 9)
            .unwrap()
            .value(),
        &CellValue::Float(42.5)
    );
    let book = reopened.workbook().unwrap();
    let changed_sheet = book.worksheet_by_name("Sheet1").unwrap();
    for cell in sheet.cells() {
        let cell = cell.unwrap();
        let stored = cell.as_ref();
        let other = changed_sheet
            .get_cell(stored.row(), stored.column())
            .unwrap();
        assert_eq!(other.value(), stored.value());
        assert_eq!(other.formula_bytes(), stored.formula_bytes());
        assert_eq!(other.style_id(), stored.style_id());
    }
    let lazy =
        litchi_xlsb::SourceBackedWorkbook::from_reader(std::io::Cursor::new(&output)).unwrap();
    let indexed = lazy
        .sheet(0_usize)
        .unwrap()
        .unwrap()
        .indexed_values()
        .unwrap()
        .unwrap();
    assert_eq!(
        indexed.cached_value(49, 9).unwrap(),
        Some(CellValue::Float(42.5))
    );
    assert!(members(&reopened.to_bytes().unwrap()) == after);
    assert!(
        reopened
            .apply_cell_values(&commit.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap()
            == INPUT
    );
    assert!(
        p.apply_cell_values(commit.patch())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap()
            == output
    );
    assert!(p.to_bytes().unwrap() == INPUT);
}
