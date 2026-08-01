use litchi_ooxml::xlsx::{
    ProtectedRangeSource, ProtectionPasswordVerifier, ProtectionRangeSqref,
    StrongProtectionPasswordVerifier, Workbook, WorksheetProtectedRange,
    WorksheetProtectedRangeCollection, WorksheetProtection, WorksheetProtectionMetadata,
    parse_worksheet_protection,
};
use litchi_opc::constants::relationship_type as rt;
use litchi_opc::{OpcPackage, PackURI};

const SHEET: &str = "/xl/worksheets/sheet1.xml";

fn seed_package(path: &std::path::Path, worksheet_xml: &str) {
    let directory = tempfile::tempdir().unwrap();
    let base = directory.path().join("base.xlsx");
    Workbook::create().unwrap().save(&base).unwrap();
    let mut package = OpcPackage::open(&base).unwrap();
    let part = package.get_part_mut(&PackURI::new(SHEET).unwrap()).unwrap();
    part.set_blob(worksheet_xml.as_bytes().to_vec());
    part.rels_mut().add_relationship(
        rt::HYPERLINK.to_string(),
        "https://example.test/preserved".to_string(),
        "rId9".to_string(),
        true,
    );
    package.save(path).unwrap();
}

fn typed_metadata() -> WorksheetProtectionMetadata {
    let mut protection = WorksheetProtection::new();
    protection.set_sheet_locked(true);
    protection.set_objects_locked(true);
    protection
        .set_verifier(Some(ProtectionPasswordVerifier::Strong(
            StrongProtectionPasswordVerifier::new("SHA-512", vec![1, 2], vec![3, 4], 100_000)
                .unwrap(),
        )))
        .unwrap();

    let mut core = WorksheetProtectedRange::new(
        ProtectedRangeSource::Core,
        "Editable",
        ProtectionRangeSqref::parse("A1:B2 C:C").unwrap(),
    )
    .unwrap();
    core.set_verifier(Some(ProtectionPasswordVerifier::Legacy(0x00AF)))
        .unwrap();
    let x14 = WorksheetProtectedRange::new(
        ProtectedRangeSource::Office2010,
        "Modern",
        ProtectionRangeSqref::parse("$D$4:$E$8").unwrap(),
    )
    .unwrap();
    let mut metadata = WorksheetProtectionMetadata::new();
    metadata.set_sheet_protection(Some(protection)).unwrap();
    metadata
        .set_protected_range_collections(vec![
            WorksheetProtectedRangeCollection::new(ProtectedRangeSource::Core, vec![core]).unwrap(),
            WorksheetProtectedRangeCollection::new(ProtectedRangeSource::Office2010, vec![x14])
                .unwrap(),
        ])
        .unwrap();
    metadata
}

#[test]
fn mutates_existing_package_without_rebuilding_unrelated_xml_or_relationships() {
    let directory = tempfile::tempdir().unwrap();
    let input = directory.path().join("input.xlsx");
    let output = directory.path().join("output.xlsx");
    let second = directory.path().join("second.xlsx");
    let xml = r#"<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:z="urn:preserved" mc:Ignorable="z"><sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData><mc:AlternateContent><mc:Choice Requires="z"><z:payload token="byte-exact"/></mc:Choice><mc:Fallback/></mc:AlternateContent><extLst><ext uri="{UNRELATED}"><z:payload token="keep"/></ext></extLst></worksheet>"#;
    seed_package(&input, xml);

    let mut workbook = Workbook::open(&input).unwrap();
    workbook
        .replace_worksheet_protection(0, typed_metadata())
        .unwrap();
    workbook.save(&output).unwrap();
    workbook.save(&second).unwrap();

    let first_package = OpcPackage::open(&output).unwrap();
    let first_part = first_package
        .get_part(&PackURI::new(SHEET).unwrap())
        .unwrap();
    let second_package = OpcPackage::open(&second).unwrap();
    let second_part = second_package
        .get_part(&PackURI::new(SHEET).unwrap())
        .unwrap();
    assert_eq!(first_part.blob(), second_part.blob());
    let output_xml = std::str::from_utf8(first_part.blob()).unwrap();
    assert!(output_xml.contains(r#"<z:payload token="byte-exact"/>"#));
    assert!(output_xml.contains(r#"<ext uri="{UNRELATED}"><z:payload token="keep"/></ext>"#));
    assert!(output_xml.contains("<x14:protectedRanges"));
    assert_eq!(
        first_part.rels().get("rId9").unwrap().target_ref(),
        "https://example.test/preserved"
    );
    let parsed = parse_worksheet_protection(first_part.blob()).unwrap();
    assert!(parsed.sheet_protection().unwrap().sheet_locked());
    assert_eq!(parsed.protected_ranges().count(), 2);
}

#[test]
fn emits_strict_core_protection_and_keeps_x14_extension() {
    let directory = tempfile::tempdir().unwrap();
    let input = directory.path().join("strict-input.xlsx");
    let output = directory.path().join("strict-output.xlsx");
    seed_package(
        &input,
        r#"<worksheet xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main"><sheetData/><extLst><ext uri="{OTHER}"/></extLst></worksheet>"#,
    );
    let mut workbook = Workbook::open(&input).unwrap();
    workbook
        .replace_worksheet_protection(0, typed_metadata())
        .unwrap();
    workbook.save(&output).unwrap();
    let package = OpcPackage::open(&output).unwrap();
    let xml = std::str::from_utf8(
        package
            .get_part(&PackURI::new(SHEET).unwrap())
            .unwrap()
            .blob(),
    )
    .unwrap();
    assert!(
        xml.contains(r#"<sheetProtection xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main""#)
    );
    assert!(xml.contains("<x14:protectedRanges"));
}

#[test]
fn rejects_spoofed_namespaces_and_invalid_replacements_atomically() {
    let directory = tempfile::tempdir().unwrap();
    let input = directory.path().join("spoofed.xlsx");
    let output = directory.path().join("unchanged.xlsx");
    let xml = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:f="urn:fake"><sheetData/><f:sheetProtection sheet="1"/></worksheet>"#;
    seed_package(&input, xml);
    let mut workbook = Workbook::open(&input).unwrap();
    assert!(
        workbook
            .replace_worksheet_protection(0, typed_metadata())
            .is_err()
    );
    workbook.save(&output).unwrap();
    let package = OpcPackage::open(&output).unwrap();
    assert_eq!(
        package
            .get_part(&PackURI::new(SHEET).unwrap())
            .unwrap()
            .blob(),
        xml.as_bytes()
    );

    let range = WorksheetProtectedRange::new(
        ProtectedRangeSource::Core,
        "Same",
        ProtectionRangeSqref::parse("A1").unwrap(),
    )
    .unwrap();
    assert!(
        WorksheetProtectedRangeCollection::new(
            ProtectedRangeSource::Core,
            vec![range.clone(), range],
        )
        .is_err()
    );
}

#[test]
fn failed_package_save_restores_the_original_blob_and_can_be_retried() {
    let directory = tempfile::tempdir().unwrap();
    let input = directory.path().join("input.xlsx");
    let output = directory.path().join("output.xlsx");
    seed_package(
        &input,
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>"#,
    );
    let mut workbook = Workbook::open(&input).unwrap();
    workbook
        .replace_worksheet_protection(0, typed_metadata())
        .unwrap();
    assert!(workbook.save(directory.path()).is_err());
    workbook.save(&output).unwrap();
    let package = OpcPackage::open(&output).unwrap();
    assert!(
        parse_worksheet_protection(
            package
                .get_part(&PackURI::new(SHEET).unwrap())
                .unwrap()
                .blob()
        )
        .unwrap()
        .sheet_protection()
        .is_some()
    );
}
