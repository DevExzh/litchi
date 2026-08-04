use litchi_ooxml::xlsb::Workbook as XlsbWorkbook;
use litchi_ooxml::xlsx::Workbook as XlsxWorkbook;
use litchi_ooxml_common::ribbon::{Family, Set, Version};
use litchi_opc::constants::relationship_type;
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part};
use litchi_xlsb::raw::{Kind, Writer, kind};

const UI_2007: &[u8] =
    br#"<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui"/>"#;
const UI_2010: &[u8] =
    br#"<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui"/>"#;
#[test]
fn xlsx_ribbon_facade_creates_reads_and_removes_both_families() {
    let mut workbook = XlsxWorkbook::create().expect("create XLSX workbook");
    assert!(
        workbook
            .ribbon()
            .expect("read empty XLSX Ribbon set")
            .effective()
            .is_none()
    );

    let modern_xml = UI_2010.to_vec();
    let moved_allocation = modern_xml.as_ptr();
    workbook
        .put_ribbon(Version::V2007, UI_2007.to_vec())
        .expect("put legacy XLSX Ribbon")
        .put_ribbon(Version::V2010, modern_xml)
        .expect("put modern XLSX Ribbon");
    assert_both(
        workbook.ribbon().expect("read XLSX Ribbon set"),
        moved_allocation,
    );

    assert!(
        workbook
            .remove_ribbon(Family::Modern)
            .expect("remove modern XLSX Ribbon")
    );
    assert_legacy_effective(workbook.ribbon().expect("read legacy XLSX Ribbon"));
    assert!(
        !workbook
            .remove_ribbon(Family::Modern)
            .expect("repeat modern XLSX Ribbon removal")
    );
    assert!(
        workbook
            .remove_ribbon(Family::Legacy)
            .expect("remove legacy XLSX Ribbon")
    );
    assert!(
        workbook
            .ribbon()
            .expect("read cleared XLSX Ribbon set")
            .effective()
            .is_none()
    );
}

#[test]
fn xlsb_ribbon_facade_creates_reads_and_removes_both_families() {
    let mut workbook = xlsb_workbook();
    assert!(
        workbook
            .ribbon()
            .expect("read empty XLSB Ribbon set")
            .effective()
            .is_none()
    );

    let modern_xml = UI_2010.to_vec();
    let moved_allocation = modern_xml.as_ptr();
    workbook
        .put_ribbon(Version::V2007, UI_2007.to_vec())
        .expect("put legacy XLSB Ribbon")
        .put_ribbon(Version::V2010, modern_xml)
        .expect("put modern XLSB Ribbon");
    assert_both(
        workbook.ribbon().expect("read XLSB Ribbon set"),
        moved_allocation,
    );

    assert!(
        workbook
            .remove_ribbon(Family::Modern)
            .expect("remove modern XLSB Ribbon")
    );
    assert_legacy_effective(workbook.ribbon().expect("read legacy XLSB Ribbon"));
    assert!(
        !workbook
            .remove_ribbon(Family::Modern)
            .expect("repeat modern XLSB Ribbon removal")
    );
    assert!(
        workbook
            .remove_ribbon(Family::Legacy)
            .expect("remove legacy XLSB Ribbon")
    );
    assert!(
        workbook
            .ribbon()
            .expect("read cleared XLSB Ribbon set")
            .effective()
            .is_none()
    );
}

fn assert_both(set: Set<'_>, moved_allocation: *const u8) {
    let legacy = set.legacy().expect("legacy Ribbon exists");
    assert_eq!(legacy.version(), Version::V2007);
    assert_eq!(legacy.xml(), UI_2007);

    let modern = set.modern().expect("modern Ribbon exists");
    assert_eq!(modern.version(), Version::V2010);
    assert_eq!(modern.xml(), UI_2010);
    assert_eq!(modern.xml().as_ptr(), moved_allocation);

    let effective = set.effective().expect("effective Ribbon exists");
    assert_eq!(effective.version(), Version::V2010);
    assert_eq!(effective.xml(), UI_2010);
}

fn assert_legacy_effective(set: Set<'_>) {
    assert!(set.modern().is_none());
    let effective = set.effective().expect("legacy Ribbon is effective");
    assert_eq!(effective.version(), Version::V2007);
    assert_eq!(effective.xml(), UI_2007);
}

fn xlsb_workbook() -> XlsbWorkbook {
    // BrtBundleSh declares one visible worksheet in the workbook stream.
    let mut bundle_sheet = 0u32.to_le_bytes().to_vec();
    bundle_sheet.extend_from_slice(&1u32.to_le_bytes());
    bundle_sheet.extend_from_slice(&wide("rIdSheet1"));
    bundle_sheet.extend_from_slice(&wide("Sheet1"));

    let mut workbook_part = BlobPart::new(
        PackURI::new("/xl/workbook.bin").expect("workbook URI"),
        "application/vnd.ms-excel.sheet.binary.macroEnabled.main".to_string(),
        records(&[(kind::BUNDLE_SH, bundle_sheet)]),
    );
    workbook_part.rels_mut().add_relationship(
        relationship_type::WORKSHEET.to_string(),
        "worksheets/sheet1.bin".to_string(),
        "rIdSheet1".to_string(),
        false,
    );

    let sheet_part = BlobPart::new(
        PackURI::new("/xl/worksheets/sheet1.bin").expect("worksheet URI"),
        "application/vnd.ms-excel.worksheet".to_string(),
        records(&[
            (kind::BEGIN_SHEET, Vec::new()),
            (kind::END_SHEET, Vec::new()),
        ]),
    );

    let mut package = OpcPackage::new();
    package.add_part(Box::new(workbook_part));
    package.add_part(Box::new(sheet_part));
    XlsbWorkbook::from_opc_package(package).expect("construct XLSB workbook")
}

fn records(values: &[(Kind, Vec<u8>)]) -> Vec<u8> {
    let mut bytes = Vec::new();
    let mut writer = Writer::new(&mut bytes);
    for (kind, payload) in values {
        writer
            .write_record(*kind, payload)
            .expect("write XLSB record");
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
