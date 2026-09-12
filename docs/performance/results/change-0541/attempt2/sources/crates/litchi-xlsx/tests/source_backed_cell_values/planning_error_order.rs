//! Public source-backed error-order coverage for the worksheet validation and
//! raw parser boundary.
//!
//! The selected multi-sheet path validates each complete worksheet before it
//! enters raw materialization.  These cases keep that contract observable at
//! the public `edit_sheets` boundary, including retry atomicity and the exact
//! typed error owner.

use super::*;

const STRICT_SML: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
const SHEET2: &str = "/xl/worksheets/sheet2.xml";

#[derive(Clone, Copy, Debug)]
enum ExpectedError {
    Invalid(&'static str),
    InvalidContaining(&'static str),
    Xml(&'static str),
    MarkupCompatibility(&'static str),
}

struct ErrorCase {
    name: String,
    worksheet: String,
    expected: ExpectedError,
}

fn worksheet(body: &str) -> String {
    format!(r#"<worksheet xmlns="{SML}">{body}</worksheet>"#)
}

fn valid_numeric_worksheet() -> String {
    worksheet(r#"<sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData>"#)
}

fn source_with_sheets(sheet1: &str, sheet2: &str) -> Vec<u8> {
    let mut package = OpcPackage::from_bytes(&two_sheets()).unwrap();
    package
        .get_part_mut(&PackURI::new(SHEET).unwrap())
        .unwrap()
        .set_blob(sheet1.as_bytes().to_vec());
    package
        .get_part_mut(&PackURI::new(SHEET2).unwrap())
        .unwrap()
        .set_blob(sheet2.as_bytes().to_vec());
    PackageWriter::to_bytes(&package).unwrap()
}

fn error_from_edit_sheets(editor: &SourceBackedEditor) -> Error {
    match editor.edit_sheets(["Sheet1".into(), "Sheet2".into()]) {
        Ok(_) => panic!("invalid worksheet unexpectedly accepted by edit_sheets"),
        Err(error) => error,
    }
}

fn assert_expected_error(error: Error, expected: ExpectedError, label: &str) {
    match (expected, error) {
        (ExpectedError::Invalid(expected), Error::Invalid(actual)) => {
            assert_eq!(actual, expected, "unexpected invalid error for {label}");
        },
        (ExpectedError::InvalidContaining(expected), Error::Invalid(actual)) => {
            assert!(
                actual.contains(expected),
                "invalid error for {label} did not contain {expected:?}: {actual}"
            );
        },
        (ExpectedError::Xml(expected), Error::Xml(actual)) => {
            assert_eq!(
                actual.to_string(),
                expected,
                "unexpected XML error for {label}"
            );
        },
        (
            ExpectedError::MarkupCompatibility(expected),
            Error::MarkupCompatibility(litchi_ooxml_common::mce::Error::NonConformant(actual)),
        ) => {
            assert_eq!(actual, expected, "unexpected MCE error for {label}");
        },
        (expected, actual) => {
            panic!("wrong typed error for {label}: expected {expected:?}, got {actual:?}");
        },
    }
}

fn assert_error_case(case: &ErrorCase) {
    let bytes = source_with_sheets(&case.worksheet, &valid_numeric_worksheet());
    let source = Arc::new(VersionedSource::new(bytes.clone()));
    let editor = SourceBackedEditor::from_read_at(source.clone()).unwrap();

    let first = error_from_edit_sheets(&editor);
    assert_expected_error(first, case.expected, case.name.as_str());
    assert_eq!(
        source.bytes.as_slice(),
        bytes.as_slice(),
        "source changed after first failure: {}",
        case.name
    );

    let retry = error_from_edit_sheets(&editor);
    assert_expected_error(retry, case.expected, case.name.as_str());
    assert_eq!(
        source.bytes.as_slice(),
        bytes.as_slice(),
        "source changed after retry: {}",
        case.name
    );

    // A failed selected worksheet must not poison a later transaction over a
    // different owner.  This catches provisional parser state escaping the
    // failed multi-sheet load, in addition to the immutable-source check.
    let unaffected = editor
        .edit_sheets(["Sheet2".into()])
        .unwrap()
        .commit()
        .unwrap();
    assert!(
        !unaffected.changed(),
        "unaffected sheet changed: {}",
        case.name
    );
    assert_eq!(unaffected.snapshot().len(), 1);
    assert_eq!(
        unaffected.snapshot().value(0, address("A1")),
        Some(&Value::Number(Number::new("7").unwrap()))
    );
}

fn raw_and_validator_cases() -> Vec<ErrorCase> {
    vec![
        ErrorCase {
            name: "raw invalid cell reference".into(),
            worksheet: worksheet(
                r#"<sheetData><row r="1"><c r="A2"><v>1</v></c></row></sheetData>"#,
            ),
            expected: ExpectedError::Invalid("cell reference 'A2' does not belong to row 1"),
        },
        ErrorCase {
            name: "raw invalid style lexical value".into(),
            worksheet: worksheet(
                r#"<sheetData><row r="1"><c r="A1" s="not-a-style"><v>1</v></c></row></sheetData>"#,
            ),
            expected: ExpectedError::Invalid("invalid worksheet cell style 'not-a-style'"),
        },
        ErrorCase {
            name: "raw invalid typed boolean".into(),
            worksheet: worksheet(
                r#"<sheetData><row r="1"><c r="A1" t="b"><v>maybe</v></c></row></sheetData>"#,
            ),
            expected: ExpectedError::Invalid("invalid worksheet boolean 'maybe'"),
        },
        ErrorCase {
            name: "raw invalid formula marker".into(),
            worksheet: worksheet(
                r#"<sheetData><row r="1"><c r="A1"><f>=1+1</f><v>2</v></c></row></sheetData>"#,
            ),
            expected: ExpectedError::Invalid("worksheet formula must omit the leading '='"),
        },
        ErrorCase {
            name: "raw unknown cell type reaches scalar closure".into(),
            worksheet: worksheet(
                r#"<sheetData><row r="1"><c r="A1" t="future"><v>1</v></c></row></sheetData>"#,
            ),
            expected: ExpectedError::Invalid("cell edits refuse unknown cells and cell metadata"),
        },
        ErrorCase {
            name: "metadata attribute is refused by value-only validator".into(),
            worksheet: worksheet(
                r#"<sheetData><row r="1"><c r="A1" cm="0"><v>1</v></c></row></sheetData>"#,
            ),
            expected: ExpectedError::Invalid("value-only edits refuse attribute 'cm' on 'c'"),
        },
        ErrorCase {
            name: "value metadata attribute is refused by value-only validator".into(),
            worksheet: worksheet(
                r#"<sheetData><row r="1"><c r="A1" vm="2147483648"><v>1</v></c></row></sheetData>"#,
            ),
            expected: ExpectedError::Invalid("value-only edits refuse attribute 'vm' on 'c'"),
        },
        ErrorCase {
            name: "raw unsupported XML entity in scalar".into(),
            worksheet: worksheet(
                r#"<sheetData><row r="1"><c r="A1"><v>&missing;</v></c></row></sheetData>"#,
            ),
            expected: ExpectedError::Xml(
                "invalid OOXML structure: unsupported XML entity reference '&missing;'",
            ),
        },
        ErrorCase {
            name: "validator rejects late dependency element".into(),
            worksheet: worksheet(
                r#"<sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData><mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>"#,
            ),
            expected: ExpectedError::Invalid(
                "value-only edits refuse dependency-bearing or unknown element 'mergeCells'",
            ),
        },
        ErrorCase {
            name: "validator rejects mixed SpreadsheetML dialects".into(),
            worksheet: format!(
                r#"<worksheet xmlns="{SML}"><sheetData xmlns="{STRICT_SML}"/></worksheet>"#
            ),
            expected: ExpectedError::Invalid("value-only XML mixes SpreadsheetML dialects"),
        },
        ErrorCase {
            name: "validator rejects unknown unqualified attribute".into(),
            worksheet: worksheet(
                r#"<sheetData><row r="1"><c r="A1" future="1"><v>7</v></c></row></sheetData>"#,
            ),
            expected: ExpectedError::Invalid("value-only edits refuse attribute 'future' on 'c'"),
        },
        ErrorCase {
            name: "validator rejects qualified attribute".into(),
            worksheet: format!(
                r#"<worksheet xmlns="{SML}" xmlns:x="urn:fixture:foreign"><sheetData><row r="1"><c r="A1" x:future="1"><v>7</v></c></row></sheetData></worksheet>"#
            ),
            expected: ExpectedError::Invalid("value-only edits refuse attribute 'x:future' on 'c'"),
        },
        ErrorCase {
            name: "validator rejects duplicate attribute".into(),
            worksheet: worksheet(
                r#"<sheetData><row r="1"><c r="A1" r="B1"><v>7</v></c></row></sheetData>"#,
            ),
            expected: ExpectedError::InvalidContaining("invalid value-only XML attribute"),
        },
        ErrorCase {
            name: "validator rejects malformed attribute".into(),
            worksheet: worksheet(
                r#"<sheetData><row r="1"><c r="A1" future=1><v>7</v></c></row></sheetData>"#,
            ),
            expected: ExpectedError::InvalidContaining("invalid value-only XML attribute"),
        },
        ErrorCase {
            name: "validator rejects truncated root".into(),
            worksheet: format!(r#"<worksheet xmlns="{SML}"><sheetData/>"#),
            expected: ExpectedError::Invalid("value-only XML has no complete root element"),
        },
        ErrorCase {
            name: "validator rejects mismatched closing element".into(),
            worksheet: format!(r#"<worksheet xmlns="{SML}"><sheetData></worksheet>"#),
            expected: ExpectedError::InvalidContaining("value-only XML scan failed"),
        },
        ErrorCase {
            name: "validator rejects DTD".into(),
            worksheet: format!(r#"<!DOCTYPE worksheet><worksheet xmlns="{SML}"/>"#),
            expected: ExpectedError::Invalid(
                "value-only edits refuse XML document type declarations",
            ),
        },
        ErrorCase {
            name: "validator rejects reference outside scalar".into(),
            worksheet: format!(r#"<worksheet xmlns="{SML}">&missing;<sheetData/></worksheet>"#),
            expected: ExpectedError::Invalid(
                "value-only XML has a reference outside a scalar value element",
            ),
        },
    ]
}

#[test]
fn public_multi_sheet_error_matrix_preserves_typed_owner_and_source() {
    for case in raw_and_validator_cases() {
        assert_error_case(&case);
    }
}

#[test]
fn validator_error_wins_when_raw_error_is_earlier_or_later() {
    let raw_cells = [
        ("cell reference", r#"<c r="A2"><v>not-a-number</v></c>"#),
        ("cell style", r#"<c r="A1" s="not-a-style"><v>1</v></c>"#),
        ("typed boolean", r#"<c r="A1" t="b"><v>maybe</v></c>"#),
        ("formula marker", r#"<c r="A1"><f>=1+1</f><v>2</v></c>"#),
        ("scalar XML entity", r#"<c r="A1"><v>&missing;</v></c>"#),
        ("unknown cell type", r#"<c r="A1" t="future"><v>1</v></c>"#),
    ];
    let mut cases = raw_cells
        .into_iter()
        .map(|(name, raw_cell)| ErrorCase {
            name: format!("{name} before late validator error"),
            worksheet: worksheet(&format!(
                r#"<sheetData><row r="1">{raw_cell}<c r="B1" future="1"/></row></sheetData>"#
            )),
            expected: ExpectedError::Invalid("value-only edits refuse attribute 'future' on 'c'"),
        })
        .collect::<Vec<_>>();
    cases.extend([
        ErrorCase {
            name: "validator error before later raw cell reference".into(),
            worksheet: worksheet(
                r#"<sheetData><row r="1"><c r="B1" future="1"/><c r="A2"><v>not-a-number</v></c></row></sheetData>"#,
            ),
            expected: ExpectedError::Invalid("value-only edits refuse attribute 'future' on 'c'"),
        },
        ErrorCase {
            name: "validator error before later raw style".into(),
            worksheet: worksheet(
                r#"<sheetData><row r="1"><c r="B1" future="1"/><c r="A1" s="not-a-style"><v>1</v></c></row></sheetData>"#,
            ),
            expected: ExpectedError::Invalid("value-only edits refuse attribute 'future' on 'c'"),
        },
    ]);
    for case in cases {
        assert_error_case(&case);
    }
}

#[test]
fn public_multi_sheet_edit_sheets_accepts_supported_namespace_forms_and_pi() {
    let strict_prefixed = format!(
        r#"<?worksheet-test?><x:worksheet xmlns:x="{STRICT_SML}"><x:sheetData><x:row r="1"><x:c r="A1"><x:v>7</x:v></x:c></x:row></x:sheetData></x:worksheet>"#
    );
    let rebound = format!(
        r#"<s:worksheet xmlns:s="{SML}" xmlns:f="urn:fixture:foreign"><s:sheetData><s:row r="1"><f:c xmlns:f="{SML}" r="A1"><s:v>7</s:v></f:c><s:c r="B1"><s:v>8</s:v></s:c></s:row></s:sheetData></s:worksheet>"#
    );
    let cases = [
        (
            "transitional default",
            worksheet(r#"<sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData>"#),
        ),
        ("strict prefixed with PI", strict_prefixed),
        ("transitional prefix rebinding", rebound),
    ];

    for (label, sheet1) in cases {
        let bytes = source_with_sheets(&sheet1, &valid_numeric_worksheet());
        let source = Arc::new(VersionedSource::new(bytes.clone()));
        let editor = SourceBackedEditor::from_read_at(source.clone()).unwrap();
        let commit = editor
            .edit_sheets(["Sheet1".into(), "Sheet2".into()])
            .unwrap()
            .commit()
            .unwrap();
        assert!(!commit.changed(), "{label} valid control changed");
        assert_eq!(commit.snapshot().len(), 2, "{label} selected sheet count");
        assert_eq!(commit.snapshot().sheet_name(0), Some("Sheet1"));
        assert_eq!(commit.snapshot().sheet_name(1), Some("Sheet2"));
        assert_eq!(
            commit.snapshot().value(0, address("A1")),
            Some(&Value::Number(Number::new("7").unwrap())),
            "{label} first-sheet value",
        );
        assert_eq!(
            source.bytes.as_slice(),
            bytes.as_slice(),
            "{label} source changed"
        );
    }
}

#[test]
fn full_validation_precedes_mce_preprocessing_and_raw_parsing_errors() {
    // The literal namespace selects MCE preprocessing even inside a comment.
    // Without it, the plain parser permits the PI (covered by the valid controls).
    let marker = "<!--http://schemas.openxmlformats.org/markup-compatibility/2006-->";
    let preprocessing_error =
        ExpectedError::MarkupCompatibility("DTD and processing instructions are rejected");
    let cases = [
        ErrorCase {
            name: "MCE processing instruction alone".into(),
            worksheet: worksheet(&format!("{marker}<?worksheet-test?><sheetData/>")),
            expected: preprocessing_error,
        },
        ErrorCase {
            name: "MCE preprocessing precedes earlier raw cell error".into(),
            worksheet: worksheet(&format!(
                r#"{marker}<sheetData><row r="1"><c r="A2"><v>1</v></c></row></sheetData><?worksheet-test?>"#,
            )),
            expected: preprocessing_error,
        },
        ErrorCase {
            name: "late validation error overrides MCE and raw errors".into(),
            worksheet: worksheet(&format!(
                r#"{marker}<sheetData><row r="1"><c r="A2"><v>1</v></c></row></sheetData><?worksheet-test?><mergeCells/>"#,
            )),
            expected: ExpectedError::Invalid(
                "value-only edits refuse dependency-bearing or unknown element 'mergeCells'",
            ),
        },
    ];
    for case in cases {
        assert_error_case(&case);
    }
}
