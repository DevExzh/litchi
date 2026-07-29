//! LibreOffice `calcext` conditional formats attached to `table:table`.
//!
//! The reader parses `calcext:conditional-formats` containers into strictly
//! typed inert data — conditions are never evaluated and style references are
//! never resolved — while the builder and mutable APIs author, edit, remove,
//! and round-trip them.

use litchi_odf::{
    ConditionalFormat, ConditionalFormatCondition, FlatSpreadsheet, MutableSpreadsheet,
    Spreadsheet, SpreadsheetBuilder,
};

/// Flat spreadsheet written in the shape LibreOffice Calc produces.
const FIXTURE: &str = "../../test-data/odf/ods/conditional-formats.fods";

fn format_a() -> ConditionalFormat {
    ConditionalFormat::new(
        vec!["Grades.B2:Grades.B3".to_string()],
        vec![
            ConditionalFormatCondition::new("cell-content()>5", "Good")
                .with_base_cell_address("Grades.B2"),
            ConditionalFormatCondition::new("cell-content()<=5", "Bad")
                .with_base_cell_address("Grades.B2"),
        ],
    )
    .unwrap()
}

fn format_b() -> ConditionalFormat {
    ConditionalFormat::new(
        vec![
            "Grades.A2:Grades.A3".to_string(),
            "Grades.B1".to_string(),
        ],
        vec![ConditionalFormatCondition::new(
            "cell-content()==\"\"",
            "Bad",
        )],
    )
    .unwrap()
}

fn content_xml(bytes: &[u8]) -> String {
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
    let mut entry = archive.by_name("content.xml").unwrap();
    let mut xml = String::new();
    std::io::Read::read_to_string(&mut entry, &mut xml).unwrap();
    xml
}

#[test]
fn reads_libreoffice_calcext_conditional_formats() {
    let mut flat = FlatSpreadsheet::open(FIXTURE).unwrap();
    let sheets = flat.spreadsheet_mut().sheets().unwrap();
    assert_eq!(sheets.len(), 1);

    let formats = sheets[0].conditional_formats();
    assert_eq!(formats, [format_a(), format_b()].as_slice());
    // Cell content still parses alongside the extension container.
    assert_eq!(sheets[0].rows[1].cells[0].text, "alice");
}

#[test]
fn builder_authored_conditional_formats_round_trip_the_package() {
    let mut builder = SpreadsheetBuilder::new();
    builder.add_sheet("Grades").unwrap();
    builder.add_row_with_values(&["name", "score"]).unwrap();
    builder.add_row_with_values(&["alice", "7"]).unwrap();
    builder
        .add_sheet_conditional_format(format_a())
        .unwrap()
        .add_sheet_conditional_format(format_b())
        .unwrap();

    let bytes = builder.build().unwrap();
    let xml = content_xml(&bytes);
    assert!(xml.contains(r#"xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0""#));
    assert!(xml.contains("<calcext:conditional-formats>"));
    assert!(xml.contains(r#"calcext:target-range-address="Grades.A2:Grades.A3 Grades.B1""#));
    assert!(xml.contains(r#"calcext:value="cell-content()&gt;5""#));

    let mut spreadsheet = Spreadsheet::from_bytes(bytes).unwrap();
    let sheets = spreadsheet.sheets().unwrap();
    assert_eq!(
        sheets[0].conditional_formats(),
        [format_a(), format_b()].as_slice()
    );
}

#[test]
fn mutable_creates_replaces_and_removes_conditional_formats() {
    let mut builder = SpreadsheetBuilder::new();
    builder.add_sheet("Grades").unwrap();
    builder.add_row_with_values(&["alice", "7"]).unwrap();
    let spreadsheet = Spreadsheet::from_bytes(builder.build().unwrap()).unwrap();

    let mut mutable = MutableSpreadsheet::from_spreadsheet(spreadsheet).unwrap();
    mutable
        .add_sheet_conditional_format(0, format_a())
        .unwrap();
    mutable
        .add_sheet_conditional_format(0, format_b())
        .unwrap();
    let mut reopened = Spreadsheet::from_bytes(mutable.to_bytes().unwrap()).unwrap();
    let formats = reopened.sheets().unwrap()[0].conditional_formats().to_vec();
    assert_eq!(formats, [format_a(), format_b()]);

    let mut mutable = MutableSpreadsheet::from_spreadsheet(reopened).unwrap();
    assert_eq!(
        mutable.remove_sheet_conditional_format(0, 0).unwrap(),
        Some(format_a())
    );
    assert!(mutable.remove_sheet_conditional_format(0, 5).unwrap().is_none());
    mutable
        .set_sheet_conditional_formats(1, vec![format_a()])
        .unwrap_err();
    let mut reopened = Spreadsheet::from_bytes(mutable.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened.sheets().unwrap()[0].conditional_formats(),
        [format_b()].as_slice()
    );

    let mut mutable = MutableSpreadsheet::from_spreadsheet(reopened).unwrap();
    mutable.set_sheet_conditional_formats(0, Vec::new()).unwrap();
    let mut reopened = Spreadsheet::from_bytes(mutable.to_bytes().unwrap()).unwrap();
    assert!(
        reopened.sheets().unwrap()[0]
            .conditional_formats()
            .is_empty()
    );
    let xml = content_xml(&mutable.to_bytes().unwrap());
    assert!(!xml.contains("calcext:conditional-formats"));
}

#[test]
fn invalid_conditional_formats_are_rejected_atomically() {
    let mut builder = SpreadsheetBuilder::new();
    builder.add_sheet("Sheet1").unwrap();
    assert!(
        builder
            .add_sheet_conditional_format(ConditionalFormat {
                target_range_addresses: Vec::new(),
                conditions: vec![ConditionalFormatCondition::new("x", "S")],
            })
            .is_err()
    );
    assert!(
        builder
            .set_sheet_conditional_formats(vec![ConditionalFormat {
                target_range_addresses: vec![".A1".to_string()],
                conditions: Vec::new(),
            }])
            .is_err()
    );

    let spreadsheet = Spreadsheet::from_bytes(builder.build().unwrap()).unwrap();
    let mut mutable = MutableSpreadsheet::from_spreadsheet(spreadsheet).unwrap();
    assert!(
        mutable
            .add_sheet_conditional_format(0, ConditionalFormat {
                target_range_addresses: vec![".A1".to_string()],
                conditions: vec![ConditionalFormatCondition::new("", "S")],
            })
            .is_err()
    );
    assert!(mutable.add_sheet_conditional_format(9, format_a()).is_err());
    assert!(
        mutable
            .remove_sheet_conditional_format(9, 0)
            .is_err()
    );
}

#[test]
fn parser_rejects_malformed_calcext_content() {
    let document = |table_inner: &str| {
        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0" office:mimetype="application/vnd.oasis.opendocument.spreadsheet" office:version="1.3"><office:body><office:spreadsheet><table:table table:name="S">{table_inner}</table:table></office:spreadsheet></office:body></office:document>"#
        )
    };
    // Sheets parse lazily, so force the parse to observe the error.
    let rejects = |table_inner: &str| {
        let mut flat = FlatSpreadsheet::from_bytes(document(table_inner).into_bytes()).unwrap();
        flat.spreadsheet_mut().sheets().is_err()
    };

    // Missing calcext:target-range-address.
    assert!(rejects(
        r#"<calcext:conditional-formats><calcext:conditional-format><calcext:condition calcext:apply-style-name="A" calcext:value="x"/></calcext:conditional-format></calcext:conditional-formats>"#,
    ));
    // Missing calcext:value.
    assert!(rejects(
        r#"<calcext:conditional-formats><calcext:conditional-format calcext:target-range-address="S.A1"><calcext:condition calcext:apply-style-name="A"/></calcext:conditional-format></calcext:conditional-formats>"#,
    ));
    // A conditional format without any condition.
    assert!(rejects(
        r#"<calcext:conditional-formats><calcext:conditional-format calcext:target-range-address="S.A1"/></calcext:conditional-formats>"#,
    ));
    // calcext:condition must not have child elements.
    assert!(rejects(
        r#"<calcext:conditional-formats><calcext:conditional-format calcext:target-range-address="S.A1"><calcext:condition calcext:apply-style-name="A" calcext:value="x"><calcext:condition calcext:apply-style-name="B" calcext:value="y"/></calcext:condition></calcext:conditional-format></calcext:conditional-formats>"#,
    ));
    // A spoofed extension namespace is rejected.
    assert!(rejects(
        r#"<calcext:conditional-formats xmlns:calcext="urn:not-calcext"><calcext:conditional-format calcext:target-range-address="S.A1"><calcext:condition calcext:apply-style-name="A" calcext:value="x"/></calcext:conditional-format></calcext:conditional-formats>"#,
    ));
}

#[test]
fn unmodeled_calcext_rule_types_are_skipped() {
    let document = r##"<?xml version="1.0" encoding="UTF-8"?>
<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0" office:mimetype="application/vnd.oasis.opendocument.spreadsheet" office:version="1.3"><office:body><office:spreadsheet><table:table table:name="S"><calcext:conditional-formats><calcext:conditional-format calcext:target-range-address="S.A1:S.A2"><calcext:color-scale><calcext:color-scale-entry calcext:type="minimum" calcext:color="#ff0000"/></calcext:color-scale><calcext:condition calcext:apply-style-name="A" calcext:value="cell-content()&gt;0"/></calcext:conditional-format></calcext:conditional-formats></table:table></office:spreadsheet></office:body></office:document>"##;
    let mut flat = FlatSpreadsheet::from_bytes(document.as_bytes().to_vec()).unwrap();
    let sheets = flat.spreadsheet_mut().sheets().unwrap();
    assert_eq!(
        sheets[0].conditional_formats(),
        [ConditionalFormat::new(
            vec!["S.A1:S.A2".to_string()],
            vec![ConditionalFormatCondition::new("cell-content()>0", "A")],
        )
        .unwrap()]
        .as_slice()
    );
}
