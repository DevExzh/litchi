//! Focused data-validation codec tests.

use super::parse_data_validation_collections;
use crate::data_validation::model::{Collection, ListSource, Source, ValidationType};
use crate::data_validation::{MAX_FORMULA_BYTES, MAX_REFERENCES};
use litchi_opc::PackURI;
fn sheet(relative: &str, index: u32) -> Vec<Collection> {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
    let package = litchi_opc::phys_pkg::OwnedPhysPkgReader::open(root.join(relative))
        .unwrap_or_else(|e| panic!("open {relative}: {e}"));
    let uri = PackURI::new(format!("/xl/worksheets/sheet{index}.xml")).unwrap();
    let bytes = package
        .blob_for(&uri)
        .unwrap_or_else(|e| panic!("sheet {relative}: {e}"));
    parse_data_validation_collections(&bytes).unwrap_or_else(|e| panic!("parse {relative}: {e}"))
}
fn count(values: &[Collection]) -> usize {
    values.iter().map(|v| v.validations.len()).sum()
}

#[test]
fn parses_poi_and_libreoffice_fixtures() {
    let cases = [
        (
            "test-data/poi/test-data/spreadsheet/DataValidationListTooLong.xlsx",
            1usize,
        ),
        (
            "test-data/poi/test-data/spreadsheet/DataValidations-49244.xlsx",
            52,
        ),
        (
            "test-data/poi/test-data/spreadsheet/dataValidationTableRange.xlsx",
            5,
        ),
        (
            "test-data/poi/test-data/spreadsheet/DataValidationEvaluations.xlsx",
            17,
        ),
        (
            "test-data/libreoffice-core/sc/qa/unit/data/xlsx/textLengthDataValidity.xlsx",
            1,
        ),
        (
            "test-data/libreoffice-core/sc/qa/unit/data/xlsx/invalid_ext_data_validation.xlsx",
            1,
        ),
        (
            "test-data/libreoffice-core/sc/qa/unit/data/xlsx/dataValidity.xlsx",
            1,
        ),
    ];
    let mut parsed = Vec::new();
    for (case, expected) in cases {
        let values = sheet(case, 1);
        assert_eq!(count(&values), expected, "{case}");
        parsed.push(values);
    }
    let extension = &parsed[5][0].validations[0];
    assert_eq!(extension.source, Source::Office2010);
    assert_eq!(extension.sqref.ranges[0].as_str(), "F6");
    assert!(
        matches!(&extension.formula1,Some(ListSource::Formula(v))if v.as_str()=="[2]Tabelle1!#REF!")
    );
    let text = &parsed[4][0].validations[0];
    assert_eq!(text.validation_type, ValidationType::TextLength);
    assert_eq!(
        text.uid.as_deref(),
        Some("{3FE27F7A-BE41-432D-B94C-05DA7B860A0B}")
    );
    assert!(
        matches!(&parsed[0][0].validations[0].formula1,Some(ListSource::Formula(v))if v.as_str().len()>255)
    );
}

#[test]
fn parses_strict_mce_and_quoted_list() {
    let xml=br#"<worksheet xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main" xmlns:x12ac="http://schemas.microsoft.com/office/spreadsheetml/2011/1/ac"><dataValidations count="1"><dataValidation type="whole" operator="greaterThan" sqref="A1"><formula1>1</formula1></dataValidation></dataValidations><mc:AlternateContent><mc:Choice Requires="x14"><extLst><ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}"><x14:dataValidations count="1" disablePrompts="1" xWindow="2" yWindow="3"><x14:dataValidation type="list"><x14:formula1><x12ac:list>&quot;a,b&quot;</x12ac:list></x14:formula1><xm:sqref adjusted="1" adjust="1">B2 C3:C4</xm:sqref></x14:dataValidation></x14:dataValidations></ext></extLst></mc:Choice><mc:Fallback><dataValidations count="1"><dataValidation type="none" sqref="D4"/></dataValidations></mc:Fallback></mc:AlternateContent></worksheet>"#;
    let values = parse_data_validation_collections(xml).unwrap();
    assert_eq!(values.len(), 2);
    assert!(values[1].disable_prompts);
    assert_eq!(values[1].x_window, Some(2));
    assert!(
        matches!(&values[1].validations[0].formula1,Some(ListSource::QuotedList(v))if v=="\"a,b\"")
    );
    assert_eq!(values[1].validations[0].sqref.ranges.len(), 2);
}

#[test]
fn rejects_malformed_and_dangerous_input() {
    let bad = [
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dataValidations count="2"><dataValidation type="none" sqref="A1"/></dataValidations></worksheet>"#,
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dataValidations count="1"><dataValidation type="whole" sqref="A0"><formula1>1</formula1><formula2>2</formula2></dataValidation></dataValidations></worksheet>"#,
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dataValidations count="1"><dataValidation type="whole" operator="greaterThan" sqref="A1"><formula2>2</formula2><formula1>1</formula1></dataValidation></dataValidations></worksheet>"#,
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dataValidations count="1"><dataValidation type="list" sqref="A1"><formula1>x</formula1><formula2>y</formula2></dataValidation></dataValidations></worksheet>"#,
        r#"<!DOCTYPE x><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#,
    ];
    for xml in bad {
        assert!(
            parse_data_validation_collections(xml.as_bytes()).is_err(),
            "{xml}"
        );
    }
    for attributes in [
        "type=\"integer\"",
        "allowBlank=\"TRUE\"",
        "operator=\"near\"",
        "errorStyle=\"fatal\"",
        "imeMode=\"automatic\"",
        "xr:uid=\"not-a-guid\"",
    ] {
        let xml = format!(
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"><dataValidations count="1"><dataValidation {attributes} sqref="A1"/></dataValidations></worksheet>"#
        );
        assert!(
            parse_data_validation_collections(xml.as_bytes()).is_err(),
            "{attributes}"
        );
    }
    let long_title = "x".repeat(33);
    let xml = format!(
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dataValidations count="1"><dataValidation type="none" errorTitle="{long_title}" sqref="A1"/></dataValidations></worksheet>"#
    );
    assert!(parse_data_validation_collections(xml.as_bytes()).is_err());
    let long_formula = "x".repeat(MAX_FORMULA_BYTES + 1);
    let xml = format!(
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dataValidations count="1"><dataValidation type="custom" sqref="A1"><formula1>{long_formula}</formula1></dataValidation></dataValidations></worksheet>"#
    );
    assert!(parse_data_validation_collections(xml.as_bytes()).is_err());
    assert!(parse_data_validation_collections(br#"<?bad x?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#).is_err());
    assert!(parse_data_validation_collections(br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dataValidations count="1"><dataValidation type="custom" sqref="A1"><formula1>&bogus;</formula1></dataValidation></dataValidations></worksheet>"#).is_err());
    let wrong_uri=br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><extLst><ext uri="wrong"><x14:dataValidations count="1"><x14:dataValidation type="none"><xm:sqref>A1</xm:sqref></x14:dataValidation></x14:dataValidations></ext></extLst></worksheet>"#;
    assert!(
        parse_data_validation_collections(wrong_uri)
            .unwrap()
            .is_empty()
    );
    let bad_flags=br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><extLst><ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}"><x14:dataValidations count="1"><x14:dataValidation type="none"><xm:sqref adjusted="1">A1</xm:sqref></x14:dataValidation></x14:dataValidations></ext></extLst></worksheet>"#;
    assert!(parse_data_validation_collections(bad_flags).is_err());
    let too_many = (0..=MAX_REFERENCES)
        .map(|_| "A1")
        .collect::<Vec<_>>()
        .join(" ");
    let xml = format!(
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dataValidations count="1"><dataValidation type="none" sqref="{too_many}"/></dataValidations></worksheet>"#
    );
    assert!(parse_data_validation_collections(xml.as_bytes()).is_err());
}
