//! LibreOffice `calcext` conditional formats attached to `table:table`.
//!
//! The reader parses `calcext:conditional-formats` containers into strictly
//! typed inert data — conditions are never evaluated, style references are
//! never resolved, and thresholds are never rendered — while the builder and
//! mutable APIs author, edit, remove, and round-trip them.

use litchi_odf::{
    ConditionalColorScale, ConditionalColorScaleEntry, ConditionalDataBar, ConditionalDataBarEntry,
    ConditionalDateIs, ConditionalDateType, ConditionalFormat, ConditionalFormatCondition,
    ConditionalFormatEntryType, ConditionalFormatRule, ConditionalIconSet, ConditionalIconSetEntry,
    DataBarAxisPosition, FlatSpreadsheet, IconSetType, MutableSpreadsheet, Spreadsheet,
    SpreadsheetBuilder,
};

/// Flat spreadsheet written in the shape LibreOffice Calc produces.
const FIXTURE: &str = "../../test-data/odf/ods/conditional-formats.fods";

fn format_a() -> ConditionalFormat {
    ConditionalFormat::new(
        vec!["Grades.B2:Grades.B3".to_string()],
        vec![
            ConditionalFormatCondition::new("cell-content()>5", "Good")
                .with_base_cell_address("Grades.B2")
                .into(),
            ConditionalFormatCondition::new("cell-content()<=5", "Bad")
                .with_base_cell_address("Grades.B2")
                .into(),
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
        vec![
            ConditionalFormatCondition::new("cell-content()==\"\"", "Bad").into(),
        ],
    )
    .unwrap()
}

fn color_scale() -> ConditionalColorScale {
    ConditionalColorScale::new(vec![
        ConditionalColorScaleEntry::new(ConditionalFormatEntryType::Minimum, "0", "#ff0000"),
        ConditionalColorScaleEntry::new(ConditionalFormatEntryType::Percentile, "50", "#ffff00"),
        ConditionalColorScaleEntry::new(ConditionalFormatEntryType::Maximum, "0", "#00ff00"),
    ])
}

fn data_bar() -> ConditionalDataBar {
    ConditionalDataBar::new(vec![
        ConditionalDataBarEntry::new(ConditionalFormatEntryType::AutomaticMinimum, "0"),
        ConditionalDataBarEntry::new(ConditionalFormatEntryType::AutomaticMaximum, "0"),
    ])
    .with_colors("#638ec6", Some("#ff8080".to_string()))
    .with_gradient(false)
    .with_axis(DataBarAxisPosition::Middle, Some("#000000".to_string()))
    .with_show_value(false)
    .with_lengths(Some("10".to_string()), Some("90".to_string()))
}

fn icon_set() -> ConditionalIconSet {
    ConditionalIconSet::new(
        IconSetType::ThreeTrafficLights1,
        vec![
            ConditionalIconSetEntry::new(ConditionalFormatEntryType::Percent, "0"),
            ConditionalIconSetEntry::new(ConditionalFormatEntryType::Percent, "33")
                .with_greater_equal(false),
            ConditionalIconSetEntry::new(ConditionalFormatEntryType::Percent, "67"),
        ],
    )
    .with_show_value(false)
}

fn format_c() -> ConditionalFormat {
    ConditionalFormat::new(
        vec!["Grades.B2:Grades.B3".to_string()],
        vec![
            color_scale().into(),
            data_bar().into(),
            icon_set().into(),
            ConditionalDateIs::new(ConditionalDateType::Last7Days, "Bad").into(),
        ],
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
    assert_eq!(formats, [format_a(), format_b(), format_c()].as_slice());
    // The condition-only view keeps expression rules in document order.
    let conditions: Vec<_> = formats[2].conditions().collect();
    assert!(conditions.is_empty());
    assert_eq!(formats[0].conditions().count(), 2);
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
        .unwrap()
        .add_sheet_conditional_format(format_c())
        .unwrap();

    let bytes = builder.build().unwrap();
    let xml = content_xml(&bytes);
    assert!(xml.contains(r#"xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0""#));
    assert!(xml.contains("<calcext:conditional-formats>"));
    assert!(xml.contains(r#"calcext:target-range-address="Grades.A2:Grades.A3 Grades.B1""#));
    assert!(xml.contains(r#"calcext:value="cell-content()&gt;5""#));
    assert!(xml.contains(
        r##"<calcext:color-scale-entry calcext:value="50" calcext:type="percentile" calcext:color="#ffff00"/>"##
    ));
    assert!(xml.contains(
        r##"<calcext:data-bar calcext:gradient="false" calcext:show-value="false" calcext:min-length="10" calcext:max-length="90" calcext:negative-color="#ff8080" calcext:axis-position="middle" calcext:positive-color="#638ec6" calcext:axis-color="#000000">"##
    ));
    assert!(xml.contains(
        r#"<calcext:icon-set calcext:icon-set-type="3TrafficLights1" calcext:show-value="false">"#
    ));
    assert!(xml.contains(
        r#"<calcext:formatting-entry calcext:value="33" calcext:greater-equal="false" calcext:type="percent"/>"#
    ));
    assert!(xml.contains(r#"<calcext:date-is calcext:style="Bad" calcext:date="last-7-days"/>"#));

    let mut spreadsheet = Spreadsheet::from_bytes(bytes).unwrap();
    let sheets = spreadsheet.sheets().unwrap();
    assert_eq!(
        sheets[0].conditional_formats(),
        [format_a(), format_b(), format_c()].as_slice()
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
        .add_sheet_conditional_format(0, format_c())
        .unwrap();
    let mut reopened = Spreadsheet::from_bytes(mutable.to_bytes().unwrap()).unwrap();
    let formats = reopened.sheets().unwrap()[0].conditional_formats().to_vec();
    assert_eq!(formats, [format_a(), format_c()]);

    let mut mutable = MutableSpreadsheet::from_spreadsheet(reopened).unwrap();
    assert_eq!(
        mutable.remove_sheet_conditional_format(0, 0).unwrap(),
        Some(format_a())
    );
    assert!(mutable.remove_sheet_conditional_format(0, 5).unwrap().is_none());
    mutable
        .set_sheet_conditional_formats(1, vec![format_b()])
        .unwrap_err();
    let mut reopened = Spreadsheet::from_bytes(mutable.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened.sheets().unwrap()[0].conditional_formats(),
        [format_c()].as_slice()
    );

    // Editing a rule body in place survives the save round trip.
    let mut edited = format_c();
    edited.rules.retain(|rule| !matches!(rule, ConditionalFormatRule::DataBar(_)));
    let mut mutable = MutableSpreadsheet::from_spreadsheet(reopened).unwrap();
    mutable
        .set_sheet_conditional_formats(0, vec![edited.clone()])
        .unwrap();
    let mut reopened = Spreadsheet::from_bytes(mutable.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened.sheets().unwrap()[0].conditional_formats(),
        [edited].as_slice()
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
                rules: vec![ConditionalFormatCondition::new("x", "S").into()],
            })
            .is_err()
    );
    assert!(
        builder
            .set_sheet_conditional_formats(vec![ConditionalFormat {
                target_range_addresses: vec![".A1".to_string()],
                rules: Vec::new(),
            }])
            .is_err()
    );
    // A data bar with a single limit entry is not writable.
    assert!(
        builder
            .add_sheet_conditional_format(ConditionalFormat {
                target_range_addresses: vec![".A1".to_string()],
                rules: vec![ConditionalDataBar::new(vec![ConditionalDataBarEntry::new(
                    ConditionalFormatEntryType::AutomaticMinimum,
                    "0",
                )])
                .into()],
            })
            .is_err()
    );

    let spreadsheet = Spreadsheet::from_bytes(builder.build().unwrap()).unwrap();
    let mut mutable = MutableSpreadsheet::from_spreadsheet(spreadsheet).unwrap();
    assert!(
        mutable
            .add_sheet_conditional_format(0, ConditionalFormat {
                target_range_addresses: vec![".A1".to_string()],
                rules: vec![ConditionalFormatCondition::new("", "S").into()],
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
    let format = |rules: &str| {
        format!(
            r#"<calcext:conditional-formats><calcext:conditional-format calcext:target-range-address="S.A1">{rules}</calcext:conditional-format></calcext:conditional-formats>"#
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
    assert!(rejects(&format(
        r#"<calcext:condition calcext:apply-style-name="A"/>"#,
    )));
    // A conditional format without any rule.
    assert!(rejects(
        r#"<calcext:conditional-formats><calcext:conditional-format calcext:target-range-address="S.A1"/></calcext:conditional-formats>"#,
    ));
    // calcext:condition must not have child elements.
    assert!(rejects(&format(
        r#"<calcext:condition calcext:apply-style-name="A" calcext:value="x"><calcext:condition calcext:apply-style-name="B" calcext:value="y"/></calcext:condition>"#,
    )));
    // A spoofed extension namespace is rejected.
    assert!(rejects(
        r#"<calcext:conditional-formats xmlns:calcext="urn:not-calcext"><calcext:conditional-format calcext:target-range-address="S.A1"><calcext:condition calcext:apply-style-name="A" calcext:value="x"/></calcext:conditional-format></calcext:conditional-formats>"#,
    ));
    // Color-scale entries require a type, a value, and a valid color.
    assert!(rejects(&format(
        r##"<calcext:color-scale><calcext:color-scale-entry calcext:value="0" calcext:color="#ff0000"/></calcext:color-scale>"##,
    )));
    assert!(rejects(&format(
        r##"<calcext:color-scale><calcext:color-scale-entry calcext:value="0" calcext:type="minimum" calcext:color="red"/></calcext:color-scale>"##,
    )));
    // Entry elements must not have child elements.
    assert!(rejects(&format(
        r##"<calcext:color-scale><calcext:color-scale-entry calcext:value="0" calcext:type="minimum" calcext:color="#ff0000"><calcext:color-scale-entry calcext:value="1" calcext:type="maximum" calcext:color="#00ff00"/></calcext:color-scale-entry></calcext:color-scale>"##,
    )));
    // Empty rule bodies are invalid.
    assert!(rejects(&format(r#"<calcext:color-scale/>"#)));
    assert!(rejects(&format(r#"<calcext:data-bar/>"#)));
    assert!(rejects(&format(r#"<calcext:icon-set calcext:icon-set-type="5Boxes"/>"#)));
    // Data bars accept exactly two limit entries.
    assert!(rejects(&format(
        r#"<calcext:data-bar><calcext:formatting-entry calcext:value="0" calcext:type="auto-minimum"/></calcext:data-bar>"#,
    )));
    assert!(rejects(&format(
        r#"<calcext:data-bar><calcext:formatting-entry calcext:value="0" calcext:type="auto-minimum"/><calcext:formatting-entry calcext:value="1" calcext:type="auto-maximum"/><calcext:formatting-entry calcext:value="2" calcext:type="number"/></calcext:data-bar>"#,
    )));
    // Unknown enum values are rejected.
    assert!(rejects(&format(
        r#"<calcext:icon-set calcext:icon-set-type="9Wonders"><calcext:formatting-entry calcext:value="0" calcext:type="percent"/></calcext:icon-set>"#,
    )));
    assert!(rejects(&format(
        r#"<calcext:icon-set calcext:icon-set-type="5Boxes"><calcext:formatting-entry calcext:value="0" calcext:type="fraction"/></calcext:icon-set>"#,
    )));
    assert!(rejects(&format(
        r#"<calcext:date-is calcext:style="A" calcext:date="next-century"/>"#,
    )));
    assert!(rejects(&format(
        r#"<calcext:data-bar calcext:axis-position=" sideways "><calcext:formatting-entry calcext:value="0" calcext:type="auto-minimum"/><calcext:formatting-entry calcext:value="1" calcext:type="auto-maximum"/></calcext:data-bar>"#,
    )));
    // calcext:date-is requires its style.
    assert!(rejects(&format(r#"<calcext:date-is calcext:date="today"/>"#)));
    // Non-numeric thresholds are rejected for non-formula entry types.
    assert!(rejects(&format(
        r#"<calcext:icon-set calcext:icon-set-type="5Boxes"><calcext:formatting-entry calcext:value="many" calcext:type="percent"/></calcext:icon-set>"#,
    )));
}

#[test]
fn legacy_data_bar_entry_and_custom_iconset_children_are_tolerated() {
    // The legacy `calcext:data-bar-entry` alias reads like a formatting entry,
    // and unmodeled `calcext:custom-iconset` children are skipped while the
    // thresholds around them still parse.
    let document = r##"<?xml version="1.0" encoding="UTF-8"?>
<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0" office:mimetype="application/vnd.oasis.opendocument.spreadsheet" office:version="1.3"><office:body><office:spreadsheet><table:table table:name="S"><calcext:conditional-formats><calcext:conditional-format calcext:target-range-address="S.A1:S.A2"><calcext:data-bar calcext:positive-color="#638ec6"><calcext:data-bar-entry calcext:value="0" calcext:type="auto-minimum"/><calcext:data-bar-entry calcext:value="0" calcext:type="auto-maximum"/></calcext:data-bar><calcext:icon-set calcext:icon-set-type="3Arrows" calcext:custom="true"><calcext:custom-iconset calcext:custom-iconset-name="3Arrows" calcext:custom-iconset-index="0"/><calcext:formatting-entry calcext:value="0" calcext:type="percent"/></calcext:icon-set></calcext:conditional-format></calcext:conditional-formats></table:table></office:spreadsheet></office:body></office:document>"##;
    let mut flat = FlatSpreadsheet::from_bytes(document.as_bytes().to_vec()).unwrap();
    let sheets = flat.spreadsheet_mut().sheets().unwrap();
    let expected_data_bar = ConditionalDataBar::new(vec![
        ConditionalDataBarEntry::new(ConditionalFormatEntryType::AutomaticMinimum, "0"),
        ConditionalDataBarEntry::new(ConditionalFormatEntryType::AutomaticMaximum, "0"),
    ])
    .with_colors("#638ec6", None);
    let expected_icon_set = ConditionalIconSet::new(
        IconSetType::ThreeArrows,
        vec![ConditionalIconSetEntry::new(
            ConditionalFormatEntryType::Percent,
            "0",
        )],
    )
    .with_custom(true);
    assert_eq!(
        sheets[0].conditional_formats(),
        [ConditionalFormat::new(
            vec!["S.A1:S.A2".to_string()],
            vec![expected_data_bar.into(), expected_icon_set.into()],
        )
        .unwrap()]
        .as_slice()
    );
}
