use litchi_ooxml::xlsx::{
    Collection, Formula, ListSource, Protection, Source, Sqref, Validation, ValidationOperator,
    ValidationType, Workbook, parse_data_validation_collections,
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

fn typed_collections() -> Vec<Collection> {
    let mut core = Validation::new(
        Source::Core,
        ValidationType::Whole,
        Sqref::parse("A1:B2 C3").unwrap(),
    );
    core.set_operator(ValidationOperator::Between);
    core.set_formula1(Some(ListSource::Formula(Formula::new("1").unwrap())))
        .unwrap();
    core.set_formula2(Some(Formula::new("10").unwrap()))
        .unwrap();
    core.set_show_error_message(true);
    core.set_error_title(Some("Bounds".to_string())).unwrap();
    core.set_error(Some("Choose 1 through 10".to_string()))
        .unwrap();

    let mut modern = Validation::new(
        Source::Office2010,
        ValidationType::List,
        Sqref::parse("$D$4:$D$8")
            .unwrap()
            .with_office2010_flags(true, false, true, true)
            .unwrap(),
    );
    modern
        .set_formula1(Some(ListSource::QuotedList("Red,Green,Blue".to_string())))
        .unwrap();
    modern
        .set_uid(Some("{11111111-2222-3333-4444-555555555555}".to_string()))
        .unwrap();

    let mut core_collection = Collection::new(Source::Core, vec![core]).unwrap();
    core_collection.set_window(Some(12), Some(34)).unwrap();
    vec![
        core_collection,
        Collection::new(Source::Office2010, vec![modern]).unwrap(),
    ]
}

#[test]
fn mutates_packaged_core_and_x14_without_rebuilding_unrelated_content() {
    let directory = tempfile::tempdir().unwrap();
    let input = directory.path().join("input.xlsx");
    let output = directory.path().join("output.xlsx");
    let second = directory.path().join("second.xlsx");
    let xml = r#"<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:z="urn:preserved" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main" mc:Ignorable="z"><sheetData/><conditionalFormatting sqref="Z1"><cfRule type="expression" priority="1"><formula>1</formula></cfRule></conditionalFormatting><dataValidations count="1"><dataValidation type="none" sqref="A1"/></dataValidations><hyperlinks/><mc:AlternateContent><mc:Choice Requires="z"><z:payload token="byte-exact"/></mc:Choice><mc:Fallback/></mc:AlternateContent><extLst><ext uri="{UNRELATED}"><z:payload token="keep"/></ext><ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}"><x14:dataValidations count="1"><x14:dataValidation type="none"><xm:sqref>B2</xm:sqref></x14:dataValidation></x14:dataValidations></ext></extLst></worksheet>"#;
    seed_package(&input, xml);

    let mut workbook = Workbook::open(&input).unwrap();
    workbook
        .replace_worksheet_data_validations(0, typed_collections())
        .unwrap();
    let mut protection = Protection::new();
    protection.set_sheet_locked(true);
    workbook.set_sheet_protection(0, protection).unwrap();
    workbook.save(&output).unwrap();
    workbook.save(&second).unwrap();

    let first = OpcPackage::open(&output).unwrap();
    let first_part = first.get_part(&PackURI::new(SHEET).unwrap()).unwrap();
    let second = OpcPackage::open(&second).unwrap();
    let second_part = second.get_part(&PackURI::new(SHEET).unwrap()).unwrap();
    assert_eq!(first_part.blob(), second_part.blob());
    let output_xml = std::str::from_utf8(first_part.blob()).unwrap();
    assert!(output_xml.contains(r#"<z:payload token="byte-exact"/>"#));
    assert!(output_xml.contains(r#"<ext uri="{UNRELATED}"><z:payload token="keep"/></ext>"#));
    assert!(output_xml.contains("<sheetProtection"));
    assert_eq!(
        first_part.rels().get("rId9").unwrap().target_ref(),
        "https://example.test/preserved"
    );
    let parsed = parse_data_validation_collections(first_part.blob()).unwrap();
    assert_eq!(parsed.len(), 2);
    assert_eq!(
        parsed
            .iter()
            .map(|value| value.validations().len())
            .sum::<usize>(),
        2
    );
    assert_eq!(parsed[0].x_window(), Some(12));
    assert_eq!(
        parsed[1].validations()[0].uid(),
        Some("{11111111-2222-3333-4444-555555555555}")
    );
}

#[test]
fn emits_strict_core_and_office2010_extensions() {
    let directory = tempfile::tempdir().unwrap();
    let input = directory.path().join("strict.xlsx");
    let output = directory.path().join("strict-output.xlsx");
    seed_package(
        &input,
        r#"<worksheet xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main"><sheetData/><extLst><ext uri="{OTHER}"/></extLst></worksheet>"#,
    );
    let mut workbook = Workbook::open(&input).unwrap();
    workbook
        .replace_worksheet_data_validations(0, typed_collections())
        .unwrap();
    workbook.save(&output).unwrap();
    let package = OpcPackage::open(&output).unwrap();
    let part = package.get_part(&PackURI::new(SHEET).unwrap()).unwrap();
    let xml = std::str::from_utf8(part.blob()).unwrap();
    assert!(
        xml.contains(r#"<dataValidations xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main""#)
    );
    assert!(xml.contains("<x14:dataValidations"));
    assert_eq!(
        parse_data_validation_collections(part.blob())
            .unwrap()
            .len(),
        2
    );
}

#[test]
fn rejects_spoofed_namespace_and_constraints_without_queuing_changes() {
    let directory = tempfile::tempdir().unwrap();
    let input = directory.path().join("spoof.xlsx");
    let output = directory.path().join("unchanged.xlsx");
    let xml = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:f="urn:fake"><sheetData/><f:dataValidations><f:dataValidation sqref="A1"/></f:dataValidations></worksheet>"#;
    seed_package(&input, xml);
    let mut workbook = Workbook::open(&input).unwrap();
    assert!(
        workbook
            .replace_worksheet_data_validations(0, typed_collections())
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
    assert!(Sqref::parse("XFE1").is_err());
    let mut rule = Validation::new(
        Source::Core,
        ValidationType::None,
        Sqref::parse("A1").unwrap(),
    );
    assert!(rule.set_prompt(Some("x".repeat(256))).is_err());
    let collection = Collection::new(Source::Core, vec![rule]).unwrap();
    assert!(
        litchi_ooxml::xlsx::validate_data_validation_collections(
            &[collection.clone(), collection,]
        )
        .is_err()
    );
}

#[test]
fn failed_save_restores_parts_and_typed_writer_round_trips() {
    let directory = tempfile::tempdir().unwrap();
    let input = directory.path().join("input.xlsx");
    let output = directory.path().join("output.xlsx");
    seed_package(
        &input,
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>"#,
    );
    let mut workbook = Workbook::open(&input).unwrap();
    workbook
        .replace_worksheet_data_validations(0, typed_collections())
        .unwrap();
    assert!(workbook.save(directory.path()).is_err());
    workbook.save(&output).unwrap();
    let package = OpcPackage::open(&output).unwrap();
    assert_eq!(
        parse_data_validation_collections(
            package
                .get_part(&PackURI::new(SHEET).unwrap())
                .unwrap()
                .blob()
        )
        .unwrap()
        .len(),
        2
    );

    let writer_path = directory.path().join("writer.xlsx");
    let mut writer = Workbook::create().unwrap();
    writer
        .worksheet_mut(0)
        .unwrap()
        .set_data_validation_collections(typed_collections())
        .unwrap();
    writer.save(&writer_path).unwrap();
    let package = OpcPackage::open(&writer_path).unwrap();
    assert_eq!(
        parse_data_validation_collections(
            package
                .get_part(&PackURI::new(SHEET).unwrap())
                .unwrap()
                .blob()
        )
        .unwrap()
        .len(),
        2
    );
}
