use super::codec::MAX_DEPTH;
use super::*;
use crate::error::Result;
use litchi_opc::{OpcPackage, PackURI};

const NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

fn parse(child: &str) -> Result<Option<IgnoredErrors>> {
    parse_worksheet_ignored_errors(
        format!(r#"<worksheet xmlns="{NS}">{child}</worksheet>"#).as_bytes(),
    )
}

#[test]
fn parses_all_flags_ranges_and_defaults() {
    let value = parse(concat!(
        r#"<ignoredErrors><ignoredError sqref="A1 $B$2:C3" calculatedColumn="1" "#,
        r#"emptyCellReference="true" evalError="1" formula="true" formulaRange="1" "#,
        r#"listDataValidation="true" numberStoredAsText="1" twoDigitTextYear="true" unlockedFormula="1"/>"#,
        r#"<ignoredError sqref="XFD1048576"/></ignoredErrors>"#,
    )).unwrap().unwrap();
    assert_eq!(value.entries().len(), 2);
    assert_eq!(value.entries()[0].ranges()[1].as_str(), "$B$2:C3");
    for kind in [
        IgnoredErrorType::CalculatedColumn,
        IgnoredErrorType::EmptyCellReference,
        IgnoredErrorType::EvaluationError,
        IgnoredErrorType::Formula,
        IgnoredErrorType::FormulaRange,
        IgnoredErrorType::ListDataValidation,
        IgnoredErrorType::NumberStoredAsText,
        IgnoredErrorType::TwoDigitTextYear,
        IgnoredErrorType::UnlockedFormula,
    ] {
        assert!(value.entries()[0].ignores(kind));
        assert!(!value.entries()[1].ignores(kind));
    }
}

#[test]
fn supports_strict_mce_and_extension_retention() {
    let strict = br#"<worksheet xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main"><ignoredErrors><ignoredError sqref="A1" formula="1"/></ignoredErrors></worksheet>"#;
    assert!(
        parse_worksheet_ignored_errors(strict)
            .unwrap()
            .unwrap()
            .entries()[0]
            .ignores(IgnoredErrorType::Formula)
    );
    let xml = format!(
        concat!(
            r#"<worksheet xmlns="{}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" "#,
            r#"xmlns:x="urn:unsupported" mc:Ignorable="x"><ignoredErrors><ignoredError sqref="A1" x:drop="1"/>"#,
            r#"<extLst><ext uri="urn:test"><x:payload value="safe"/></ext></extLst></ignoredErrors></worksheet>"#,
        ),
        NS
    );
    let value = parse_worksheet_ignored_errors(xml.as_bytes())
        .unwrap()
        .unwrap();
    assert_eq!(value.extensions().len(), 1);
    assert_eq!(value.extensions()[0].uri(), "urn:test");
    assert!(
        std::str::from_utf8(value.extensions()[0].markup())
            .unwrap()
            .contains("<ext")
    );
}

#[test]
fn rejects_bad_references_structure_attributes_and_limits() {
    for child in [
        "<ignoredErrors/>",
        "<ignoredErrors><ignoredError/></ignoredErrors>",
        r#"<ignoredErrors><ignoredError sqref=""/></ignoredErrors>"#,
        r#"<ignoredErrors><ignoredError sqref="A0"/></ignoredErrors>"#,
        r#"<ignoredErrors><ignoredError sqref="XFE1"/></ignoredErrors>"#,
        r#"<ignoredErrors><ignoredError sqref="A1048577"/></ignoredErrors>"#,
        r#"<ignoredErrors><ignoredError sqref="A1" formula="yes"/></ignoredErrors>"#,
        r#"<ignoredErrors><ignoredError sqref="A1" mystery="1"/></ignoredErrors>"#,
        r#"<ignoredErrors><ignoredError sqref="A1"><child/></ignoredError></ignoredErrors>"#,
        r#"<ignoredErrors><extLst><ext uri="x"/></extLst><ignoredError sqref="A1"/></ignoredErrors>"#,
        r#"<ignoredErrors><ignoredError sqref="A1"/><extLst/></ignoredErrors>"#,
    ] {
        assert!(parse(child).is_err(), "expected rejection for {child}");
    }
    assert!(parse("<ignoredErrors><ignoredError sqref=\"A1\"/></ignoredErrors><ignoredErrors><ignoredError sqref=\"A1\"/></ignoredErrors>").is_err());
    let entries = (0..10)
        .map(|index| format!(r#"<ignoredError sqref="A{}"/>"#, index + 1))
        .collect::<String>();
    assert!(parse(&format!("<ignoredErrors>{entries}</ignoredErrors>")).is_err());
}

#[test]
fn rejects_multiple_roots_direct_text_and_excessive_depth() {
    for xml in [
        format!(r#"<worksheet xmlns="{NS}"/><worksheet xmlns="{NS}"/>"#),
        format!(r#"text<worksheet xmlns="{NS}"></worksheet>"#),
        format!(r#"<worksheet xmlns="{NS}">text</worksheet>"#),
        format!(r#"<worksheet xmlns="{NS}"></worksheet>tail"#),
    ] {
        assert!(
            parse_worksheet_ignored_errors(xml.as_bytes()).is_err(),
            "expected rejection for {xml}"
        );
    }

    let mut xml = format!(r#"<worksheet xmlns="{NS}">"#);
    for _ in 0..MAX_DEPTH {
        xml.push_str("<extension>");
    }
    for _ in 0..MAX_DEPTH {
        xml.push_str("</extension>");
    }
    xml.push_str("</worksheet>");
    assert!(parse_worksheet_ignored_errors(xml.as_bytes()).is_err());
}

fn fixture(bytes: &[u8]) -> IgnoredErrors {
    let package = OpcPackage::from_bytes(bytes).unwrap();
    let part = package
        .get_part(&PackURI::new("/xl/worksheets/sheet1.xml").unwrap())
        .unwrap();
    parse_worksheet_ignored_errors(part.blob())
        .unwrap()
        .unwrap()
}

#[test]
fn reads_poi_ignored_error_fixtures() {
    let format = fixture(include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/poi/test-data/spreadsheet/FormatKM.xlsx"
    )));
    assert_eq!(
        format.entries()[0]
            .ranges()
            .iter()
            .map(IgnoredErrorRangeReference::as_str)
            .collect::<Vec<_>>(),
        vec!["C2:C5", "E2:E4", "E5"]
    );
    assert!(format.entries()[0].ignores(IgnoredErrorType::NumberStoredAsText));

    let large = fixture(include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/poi/test-data/spreadsheet/no_drawing_patriarch.xlsx"
    )));
    assert_eq!(large.entries()[0].ranges()[0].as_str(), "A1:J7577");
    assert!(large.entries()[0].ignores(IgnoredErrorType::NumberStoredAsText));
}
