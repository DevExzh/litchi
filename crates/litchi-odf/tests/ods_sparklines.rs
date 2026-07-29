//! LibreOffice `calcext` sparkline groups attached to `table:table`.
//!
//! The reader parses `calcext:sparkline-groups` containers into strictly
//! typed inert data — sparklines are never rendered and addresses are never
//! resolved — while the builder and mutable APIs author, edit, remove, and
//! round-trip them.

use litchi_odf::{
    ColorTransformationType, FlatSpreadsheet, MutableSpreadsheet, Sparkline, SparklineAxisType,
    SparklineColorTransformation, SparklineColors, SparklineComplexColor, SparklineComplexColors,
    SparklineEmptyCells, SparklineFlags, SparklineGroup, SparklineType, Spreadsheet,
    SpreadsheetBuilder, ThemeColorType,
};

/// Flat spreadsheet written in the shape LibreOffice Calc produces.
const FIXTURE: &str = "../../test-data/odf/ods/sparklines.fods";

fn group_a() -> SparklineGroup {
    SparklineGroup::new(vec![Sparkline::new(
        "Scores.D1",
        vec!["Scores.A1:Scores.C1".to_string()],
    )])
    .with_id("{1C5C5DE0-3C09-4CB3-A3EC-9E763301EC82}")
    .with_type(SparklineType::Column)
    .with_line_width("1pt")
    .with_display_empty_cells_as(SparklineEmptyCells::Gap)
    .with_flags(SparklineFlags {
        markers: Some(true),
        high: Some(true),
        low: Some(true),
        negative: Some(true),
        display_x_axis: Some(true),
        ..SparklineFlags::default()
    })
    .with_axis(
        Some(SparklineAxisType::Custom),
        Some(SparklineAxisType::Individual),
        Some("-5".to_string()),
        None,
    )
    .with_colors(SparklineColors {
        series: Some("#0369a3".to_string()),
        low: Some("#c9211e".to_string()),
        ..SparklineColors::default()
    })
    .with_complex_colors(SparklineComplexColors {
        series: Some(
            SparklineComplexColor::new(ThemeColorType::Accent3).with_transformation(
                SparklineColorTransformation::new(ColorTransformationType::LumMod, 6000),
            ),
        ),
        ..SparklineComplexColors::default()
    })
}

fn group_b() -> SparklineGroup {
    SparklineGroup::new(vec![
        Sparkline::new("Scores.D2", vec!["Scores.A2:Scores.C2".to_string()]),
        Sparkline::new(
            "Scores.E2",
            vec![
                "Scores.A1:Scores.A2".to_string(),
                "Scores.C1:Scores.C2".to_string(),
            ],
        ),
    ])
    .with_type(SparklineType::Line)
}

fn content_xml(bytes: &[u8]) -> String {
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
    let mut entry = archive.by_name("content.xml").unwrap();
    let mut xml = String::new();
    std::io::Read::read_to_string(&mut entry, &mut xml).unwrap();
    xml
}

#[test]
fn reads_libreoffice_calcext_sparkline_groups() {
    let mut flat = FlatSpreadsheet::open(FIXTURE).unwrap();
    let sheets = flat.spreadsheet_mut().sheets().unwrap();
    assert_eq!(sheets.len(), 1);

    // Both groups parse with full fidelity, including the theme-based
    // complex color and its transformation.
    assert_eq!(
        sheets[0].sparkline_groups(),
        [group_a(), group_b()].as_slice()
    );
    // Cell content still parses alongside the extension container.
    assert_eq!(sheets[0].rows[0].cells[1].text, "7");
}

#[test]
fn builder_authored_sparkline_groups_round_trip_the_package() {
    let mut builder = SpreadsheetBuilder::new();
    builder.add_sheet("Scores").unwrap();
    builder.add_row_with_values(&["3", "7", "5"]).unwrap();
    builder
        .add_sheet_sparkline_group(group_a())
        .unwrap()
        .add_sheet_sparkline_group(group_b())
        .unwrap();

    let bytes = builder.build().unwrap();
    let xml = content_xml(&bytes);
    assert!(xml.contains(r#"xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0""#));
    assert!(xml.contains("<calcext:sparkline-groups>"));
    assert!(xml.contains(
        r##"calcext:type="column" calcext:line-width="1pt" calcext:display-empty-cells-as="gap" calcext:markers="true" calcext:high="true" calcext:low="true" calcext:negative="true" calcext:display-x-axis="true" calcext:min-axis-type="custom" calcext:max-axis-type="individual" calcext:manual-min="-5" calcext:color-series="#0369a3" calcext:color-low="#c9211e""##
    ));
    assert!(xml.contains(
        r#"<calcext:sparkline calcext:cell-address="Scores.E2" calcext:data-range="Scores.A1:Scores.A2 Scores.C1:Scores.C2"/>"#
    ));
    assert!(xml.contains(
        r#"<calcext:sparkline-series-complex-color xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0" loext:theme-type="accent3" loext:color-type="theme"><loext:transformation loext:type="lummod" loext:value="6000"/></calcext:sparkline-series-complex-color>"#
    ));

    let mut spreadsheet = Spreadsheet::from_bytes(bytes).unwrap();
    let sheets = spreadsheet.sheets().unwrap();
    assert_eq!(
        sheets[0].sparkline_groups(),
        [group_a(), group_b()].as_slice()
    );
}

#[test]
fn mutable_creates_replaces_and_removes_sparkline_groups() {
    let mut builder = SpreadsheetBuilder::new();
    builder.add_sheet("Scores").unwrap();
    builder.add_row_with_values(&["3", "7", "5"]).unwrap();
    let spreadsheet = Spreadsheet::from_bytes(builder.build().unwrap()).unwrap();

    let mut mutable = MutableSpreadsheet::from_spreadsheet(spreadsheet).unwrap();
    mutable.add_sheet_sparkline_group(0, group_a()).unwrap();
    mutable.add_sheet_sparkline_group(0, group_b()).unwrap();
    let mut reopened = Spreadsheet::from_bytes(mutable.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened.sheets().unwrap()[0].sparkline_groups(),
        [group_a(), group_b()].as_slice()
    );

    let mut mutable = MutableSpreadsheet::from_spreadsheet(reopened).unwrap();
    assert_eq!(
        mutable.remove_sheet_sparkline_group(0, 0).unwrap(),
        Some(group_a())
    );
    assert!(
        mutable
            .remove_sheet_sparkline_group(0, 9)
            .unwrap()
            .is_none()
    );
    mutable
        .set_sheet_sparkline_groups(1, vec![group_b()])
        .unwrap_err();
    // Edit the surviving group in place.
    let mut edited = group_b();
    edited.sparklines.pop();
    mutable
        .set_sheet_sparkline_groups(0, vec![edited.clone()])
        .unwrap();
    let mut reopened = Spreadsheet::from_bytes(mutable.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened.sheets().unwrap()[0].sparkline_groups(),
        [edited].as_slice()
    );

    let mut mutable = MutableSpreadsheet::from_spreadsheet(reopened).unwrap();
    mutable.set_sheet_sparkline_groups(0, Vec::new()).unwrap();
    let mut reopened = Spreadsheet::from_bytes(mutable.to_bytes().unwrap()).unwrap();
    assert!(
        reopened.sheets().unwrap()[0]
            .sparkline_groups()
            .is_empty()
    );
    let xml = content_xml(&mutable.to_bytes().unwrap());
    assert!(!xml.contains("calcext:sparkline-groups"));
}

#[test]
fn invalid_sparkline_groups_are_rejected_atomically() {
    let mut builder = SpreadsheetBuilder::new();
    builder.add_sheet("Sheet1").unwrap();
    // A group without sparklines cannot be authored.
    assert!(
        builder
            .add_sheet_sparkline_group(SparklineGroup::new(Vec::new()))
            .is_err()
    );
    // A sparkline without a data range cannot be authored.
    assert!(
        builder
            .set_sheet_sparkline_groups(vec![SparklineGroup::new(vec![Sparkline::new(
                ".A1",
                Vec::new()
            )])])
            .is_err()
    );
    // A bad color cannot be authored.
    let mut bad = group_a();
    bad.colors.series = Some("blue".to_string());
    assert!(builder.add_sheet_sparkline_group(bad).is_err());

    let spreadsheet = Spreadsheet::from_bytes(builder.build().unwrap()).unwrap();
    let mut mutable = MutableSpreadsheet::from_spreadsheet(spreadsheet).unwrap();
    assert!(
        mutable
            .add_sheet_sparkline_group(0, SparklineGroup::new(Vec::new()))
            .is_err()
    );
    assert!(
        mutable
            .add_sheet_sparkline_group(9, group_a())
            .is_err()
    );
    assert!(mutable.remove_sheet_sparkline_group(9, 0).is_err());
}

#[test]
fn parser_rejects_malformed_calcext_sparkline_content() {
    let document = |table_inner: &str| {
        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0" office:mimetype="application/vnd.oasis.opendocument.spreadsheet" office:version="1.3"><office:body><office:spreadsheet><table:table table:name="S">{table_inner}</table:table></office:spreadsheet></office:body></office:document>"#
        )
    };
    let group = |inner: &str| {
        format!(
            r#"<calcext:sparkline-groups><calcext:sparkline-group {inner}</calcext:sparkline-group></calcext:sparkline-groups>"#
        )
    };
    let sparklines = |attrs: &str| {
        group(&format!(
            r#"><calcext:sparklines><calcext:sparkline {attrs}/></calcext:sparklines>"#
        ))
    };
    // Sheets parse lazily, so force the parse to observe the error.
    let rejects = |table_inner: &str| {
        let mut flat = FlatSpreadsheet::from_bytes(document(table_inner).into_bytes()).unwrap();
        flat.spreadsheet_mut().sheets().is_err()
    };

    // A group without any sparkline.
    assert!(rejects(
        r#"<calcext:sparkline-groups><calcext:sparkline-group/></calcext:sparkline-groups>"#
    ));
    assert!(rejects(&group("><calcext:sparklines/>")));
    // Sparklines require cell-address and data-range.
    assert!(rejects(&sparklines(r#"calcext:data-range="S.A1"/>"#)));
    assert!(rejects(&sparklines(r#"calcext:cell-address="S.B1"/>"#)));
    // Unknown enum values are rejected.
    assert!(rejects(&group(
        r#"calcext:type="pie"><calcext:sparklines><calcext:sparkline calcext:cell-address="S.B1" calcext:data-range="S.A1"/></calcext:sparklines>"#
    )));
    assert!(rejects(&group(
        r#"calcext:display-empty-cells-as="void"><calcext:sparklines><calcext:sparkline calcext:cell-address="S.B1" calcext:data-range="S.A1"/></calcext:sparklines>"#
    )));
    assert!(rejects(&group(
        r#"calcext:min-axis-type="global"><calcext:sparklines><calcext:sparkline calcext:cell-address="S.B1" calcext:data-range="S.A1"/></calcext:sparklines>"#
    )));
    // Bad colors and non-numeric measures are rejected.
    assert!(rejects(&group(
        r##"calcext:color-series="blue"><calcext:sparklines><calcext:sparkline calcext:cell-address="S.B1" calcext:data-range="S.A1"/></calcext:sparklines>"##
    )));
    assert!(rejects(&group(
        r#"calcext:line-width="wide"><calcext:sparklines><calcext:sparkline calcext:cell-address="S.B1" calcext:data-range="S.A1"/></calcext:sparklines>"#
    )));
    assert!(rejects(&group(
        r#"calcext:manual-min="low"><calcext:sparklines><calcext:sparkline calcext:cell-address="S.B1" calcext:data-range="S.A1"/></calcext:sparklines>"#
    )));
    // calcext:sparkline must not have child elements.
    assert!(rejects(&group(
        r#"><calcext:sparklines><calcext:sparkline calcext:cell-address="S.B1" calcext:data-range="S.A1"><calcext:sparkline calcext:cell-address="S.B2" calcext:data-range="S.A2"/></calcext:sparkline></calcext:sparklines>"#
    )));
    // Complex colors require a known theme type and well-formed transformations.
    assert!(rejects(&group(
        r#"><calcext:sparkline-series-complex-color loext:theme-type="accent9" loext:color-type="theme" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0"/><calcext:sparklines><calcext:sparkline calcext:cell-address="S.B1" calcext:data-range="S.A1"/></calcext:sparklines>"#
    )));
    assert!(rejects(&group(
        r#"><calcext:sparkline-series-complex-color loext:color-type="theme" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0"/><calcext:sparklines><calcext:sparkline calcext:cell-address="S.B1" calcext:data-range="S.A1"/></calcext:sparklines>"#
    )));
    assert!(rejects(&group(
        r#"><calcext:sparkline-series-complex-color loext:theme-type="accent3" loext:color-type="rgb" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0"/><calcext:sparklines><calcext:sparkline calcext:cell-address="S.B1" calcext:data-range="S.A1"/></calcext:sparklines>"#
    )));
    assert!(rejects(&group(
        r#"><calcext:sparkline-series-complex-color loext:theme-type="accent3" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0"><loext:transformation loext:type="invert" loext:value="50"/></calcext:sparkline-series-complex-color><calcext:sparklines><calcext:sparkline calcext:cell-address="S.B1" calcext:data-range="S.A1"/></calcext:sparklines>"#
    )));
    assert!(rejects(&group(
        r#"><calcext:sparkline-series-complex-color loext:theme-type="accent3" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0"><loext:transformation loext:type="tint" loext:value="99999"/></calcext:sparkline-series-complex-color><calcext:sparklines><calcext:sparkline calcext:cell-address="S.B1" calcext:data-range="S.A1"/></calcext:sparklines>"#
    )));
    assert!(rejects(&group(
        r#"><calcext:sparkline-series-complex-color loext:theme-type="accent3" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0"><loext:transformation loext:type="tint"/></calcext:sparkline-series-complex-color><calcext:sparklines><calcext:sparkline calcext:cell-address="S.B1" calcext:data-range="S.A1"/></calcext:sparklines>"#
    )));
    // Duplicate color slots are rejected.
    assert!(rejects(&group(
        r#"><calcext:sparkline-series-complex-color loext:theme-type="accent3" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0"/><calcext:sparkline-series-complex-color loext:theme-type="accent4" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0"/><calcext:sparklines><calcext:sparkline calcext:cell-address="S.B1" calcext:data-range="S.A1"/></calcext:sparklines>"#
    )));
    // A spoofed loext namespace reads as a missing theme type.
    assert!(rejects(&group(
        r#"><calcext:sparkline-series-complex-color xmlns:loext="urn:not-loext" loext:theme-type="accent3"/>"#
    )));
    // A spoofed extension namespace is rejected.
    assert!(rejects(
        r#"<calcext:sparkline-groups xmlns:calcext="urn:not-calcext"><calcext:sparkline-group><calcext:sparklines><calcext:sparkline calcext:cell-address="S.B1" calcext:data-range="S.A1"/></calcext:sparklines></calcext:sparkline-group></calcext:sparkline-groups>"#,
    ));
}
