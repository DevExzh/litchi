use litchi_ooxml_common::ribbon::{Family, Set, Version};
use litchi_opc::constants::relationship_type;
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part};
use litchi_xlsb::Workbook;
use litchi_xlsb::raw::{Kind, Writer, kind};

const UI_2007: &[u8] =
    br#"<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui"/>"#;
const UI_2010: &[u8] =
    br#"<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui"/>"#;

#[test]
fn workbook_ribbon_facade_creates_reads_and_removes_both_families() {
    let mut workbook = workbook();
    assert!(workbook.ribbon().unwrap().effective().is_none());

    let modern_xml = UI_2010.to_vec();
    let moved_allocation = modern_xml.as_ptr();
    workbook
        .put_ribbon(Version::V2007, UI_2007.to_vec())
        .unwrap()
        .put_ribbon(Version::V2010, modern_xml)
        .unwrap();
    assert_both(workbook.ribbon().unwrap(), moved_allocation);

    assert!(workbook.remove_ribbon(Family::Modern).unwrap());
    assert_legacy_effective(workbook.ribbon().unwrap());
    assert!(!workbook.remove_ribbon(Family::Modern).unwrap());
    assert!(workbook.remove_ribbon(Family::Legacy).unwrap());
    assert!(workbook.ribbon().unwrap().effective().is_none());
}

fn assert_both(set: Set<'_>, moved_allocation: *const u8) {
    let legacy = set.legacy().unwrap();
    assert_eq!(legacy.version(), Version::V2007);
    assert_eq!(legacy.xml(), UI_2007);
    let modern = set.modern().unwrap();
    assert_eq!(modern.version(), Version::V2010);
    assert_eq!(modern.xml(), UI_2010);
    assert_eq!(modern.xml().as_ptr(), moved_allocation);
    assert_eq!(set.effective().unwrap().version(), Version::V2010);
}

fn assert_legacy_effective(set: Set<'_>) {
    assert!(set.modern().is_none());
    assert_eq!(set.effective().unwrap().version(), Version::V2007);
}

fn workbook() -> Workbook {
    let mut bundle_sheet = 0u32.to_le_bytes().to_vec();
    bundle_sheet.extend_from_slice(&1u32.to_le_bytes());
    bundle_sheet.extend_from_slice(&wide("rIdSheet1"));
    bundle_sheet.extend_from_slice(&wide("Sheet1"));

    let mut workbook_part = BlobPart::new(
        PackURI::new("/xl/workbook.bin").unwrap(),
        "application/vnd.ms-excel.sheet.binary.macroEnabled.main".to_owned(),
        records(&[(kind::BUNDLE_SH, bundle_sheet)]),
    );
    workbook_part.rels_mut().add_relationship(
        relationship_type::WORKSHEET.to_owned(),
        "worksheets/sheet1.bin".to_owned(),
        "rIdSheet1".to_owned(),
        false,
    );

    let sheet_part = BlobPart::new(
        PackURI::new("/xl/worksheets/sheet1.bin").unwrap(),
        "application/vnd.ms-excel.worksheet".to_owned(),
        records(&[
            (kind::BEGIN_SHEET, Vec::new()),
            (kind::END_SHEET, Vec::new()),
        ]),
    );

    let mut package = OpcPackage::new();
    package.add_part(Box::new(workbook_part));
    package.add_part(Box::new(sheet_part));
    Workbook::from_opc_package(package).unwrap()
}

fn records(values: &[(Kind, Vec<u8>)]) -> Vec<u8> {
    let mut bytes = Vec::new();
    let mut writer = Writer::new(&mut bytes);
    for (kind, payload) in values {
        writer.write_record(*kind, payload).unwrap();
    }
    bytes
}

fn wide(value: &str) -> Vec<u8> {
    let mut bytes = Vec::new();
    bytes.extend_from_slice(&(value.encode_utf16().count() as u32).to_le_bytes());
    for unit in value.encode_utf16() {
        bytes.extend_from_slice(&unit.to_le_bytes());
    }
    bytes
}
