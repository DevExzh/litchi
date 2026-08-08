//! Tests for the chart sheet stream parser and its workbook wiring.

use super::*;
use crate::package::error::Error;
use crate::raw::{Error as WireError, Kind, Stage, Writer, kind as rt};

fn wide_string(value: &str) -> Vec<u8> {
    let mut data = Vec::new();
    data.extend_from_slice(&(value.encode_utf16().count() as u32).to_le_bytes());
    for unit in value.encode_utf16() {
        data.extend_from_slice(&unit.to_le_bytes());
    }
    data
}

fn stream(records: &[(Kind, Vec<u8>)]) -> Vec<u8> {
    let mut data = Vec::new();
    let mut writer = Writer::new(&mut data);
    for (record_type, payload) in records {
        writer.write_record(*record_type, payload).unwrap();
    }
    data
}

fn cs_prop(flags: u16, color: [u8; 8], code_name: &str) -> Vec<u8> {
    let mut data = flags.to_le_bytes().to_vec();
    data.extend_from_slice(&color);
    data.extend_from_slice(&wide_string(code_name));
    data
}

fn rgb_color(red: u8, green: u8, blue: u8) -> [u8; 8] {
    // fValidRGB set, xColorType = 2 (RGB), index 0, tint 0.
    [0x05, 0, 0, 0, red, green, blue, 0xFF]
}

fn cs_view(selected: u16, scale: u32, book_view: u32) -> Vec<u8> {
    let mut data = selected.to_le_bytes().to_vec();
    data.extend_from_slice(&scale.to_le_bytes());
    data.extend_from_slice(&book_view.to_le_bytes());
    data
}

fn cs_page_setup(copies: u32, flags: u16, rel_id: &str) -> Vec<u8> {
    let mut data = 9u32.to_le_bytes().to_vec(); // iPaperSize = A4
    data.extend_from_slice(&600u32.to_le_bytes()); // iRes
    data.extend_from_slice(&600u32.to_le_bytes()); // iVRes
    data.extend_from_slice(&copies.to_le_bytes());
    data.extend_from_slice(&1i16.to_le_bytes()); // iPageStart
    data.extend_from_slice(&flags.to_le_bytes());
    data.extend_from_slice(&wide_string(rel_id));
    data
}

fn cs_protection(verifier: u16, locked: u32, objects: u32) -> Vec<u8> {
    let mut data = verifier.to_le_bytes().to_vec();
    data.extend_from_slice(&locked.to_le_bytes());
    data.extend_from_slice(&objects.to_le_bytes());
    data
}

fn cs_protection_iso(spin_count: u32, locked: u32, objects: u32) -> Vec<u8> {
    let mut data = spin_count.to_le_bytes().to_vec();
    data.extend_from_slice(&locked.to_le_bytes());
    data.extend_from_slice(&objects.to_le_bytes());
    let hash = [0xAAu8; 32];
    data.extend_from_slice(&(hash.len() as u32).to_le_bytes());
    data.extend_from_slice(&hash);
    let salt = [0xBBu8; 16];
    data.extend_from_slice(&(salt.len() as u32).to_le_bytes());
    data.extend_from_slice(&salt);
    data.extend_from_slice(&wide_string("SHA-512"));
    data
}

fn chart_sheet_stream(extra: &[(Kind, Vec<u8>)]) -> Vec<u8> {
    let mut records = vec![
        (rt::BEGIN_SHEET, Vec::new()),
        (
            rt::CS_PROP,
            cs_prop(1, rgb_color(0x10, 0x20, 0x30), "ChartCode"),
        ),
        (rt::CS_PAGE_SETUP, cs_page_setup(2, 0x21, "rIdPrinter")),
        (rt::BEGIN_CS_VIEWS, Vec::new()),
        (rt::BEGIN_CS_VIEW, cs_view(1, 100, 0)),
        (rt::END_CS_VIEW, Vec::new()),
        (rt::END_CS_VIEWS, Vec::new()),
        (rt::CS_PROTECTION_ISO, cs_protection_iso(100_000, 1, 0)),
        (rt::CS_PROTECTION, cs_protection(0, 1, 0)),
        (rt::DRAWING, wide_string("rIdDrawing")),
        (rt::LEGACY_DRAWING, wide_string("rIdVml")),
        (rt::LEGACY_DRAWING_HF, wide_string("rIdVmlHf")),
    ];
    records.extend_from_slice(extra);
    records.push((rt::END_SHEET, Vec::new()));
    stream(&records)
}

#[test]
fn parses_full_chart_sheet_stream() {
    let data = chart_sheet_stream(&[
        (Kind::new(0x0FFF).unwrap(), vec![1, 2, 3]), // unknown record is skipped
        (rt::FRT_BEGIN, vec![0; 12]),
        (rt::FRT_END, Vec::new()),
    ]);
    let sheet = parse_chart_sheet_part(&data, "Chart1".to_string(), 0).unwrap();
    assert_eq!(sheet.name, "Chart1");
    assert_eq!(sheet.state, State::Visible);
    assert_eq!(sheet.code_name, "ChartCode");
    assert!(sheet.published);
    assert_eq!(
        sheet.tab_color,
        Color {
            valid_rgb: true,
            color_type: ColorType::Rgb,
            index: 0,
            tint: 0,
            rgba: [0x10, 0x20, 0x30, 0xFF],
        }
    );
    assert_eq!(sheet.views.len(), 1);
    assert_eq!(
        sheet.views[0],
        View {
            selected: true,
            scale: 100,
            workbook_view_index: 0,
        }
    );
    let page_setup = sheet.page_setup.as_ref().unwrap();
    assert_eq!(page_setup.paper_size, 9);
    assert_eq!(page_setup.horizontal_resolution, 600);
    assert_eq!(page_setup.copies, 2);
    assert_eq!(page_setup.page_start, 1);
    assert!(page_setup.landscape);
    assert!(page_setup.draft);
    assert!(!page_setup.black_and_white);
    assert_eq!(page_setup.printer_settings_rel_id, "rIdPrinter");
    assert_eq!(
        sheet.protection,
        Some(Protection {
            password_verifier: 0,
            locked: true,
            objects: false,
        })
    );
    let strong = sheet.strong_protection.as_ref().unwrap();
    assert_eq!(strong.spin_count, 100_000);
    assert_eq!(strong.hash.len(), 32);
    assert_eq!(strong.salt.len(), 16);
    assert_eq!(strong.algorithm, "SHA-512");
    assert_eq!(sheet.drawing_rel_id.as_deref(), Some("rIdDrawing"));
    assert_eq!(sheet.legacy_drawing_rel_id.as_deref(), Some("rIdVml"));
    assert_eq!(
        sheet.legacy_drawing_header_footer_rel_id.as_deref(),
        Some("rIdVmlHf")
    );
}

#[test]
fn maps_sheet_states() {
    let minimal = chart_sheet_stream(&[]);
    let hidden = parse_chart_sheet_part(&minimal, "C".to_string(), 1).unwrap();
    assert_eq!(hidden.state, State::Hidden);
    let very_hidden = parse_chart_sheet_part(&minimal, "C".to_string(), 2).unwrap();
    assert_eq!(very_hidden.state, State::VeryHidden);
    assert!(matches!(
        parse_chart_sheet_part(&minimal, "C".to_string(), 3),
        Err(Error::Unrecognized { .. })
    ));
}

#[test]
fn parses_minimal_chart_sheet_stream() {
    let data = stream(&[(rt::BEGIN_SHEET, Vec::new()), (rt::END_SHEET, Vec::new())]);
    let sheet = parse_chart_sheet_part(&data, "Chart1".to_string(), 0).unwrap();
    assert!(sheet.code_name.is_empty());
    assert!(!sheet.published);
    assert_eq!(sheet.tab_color.color_type, ColorType::Automatic);
    assert!(sheet.views.is_empty());
    assert!(sheet.protection.is_none());
    assert!(sheet.page_setup.is_none());
    assert!(sheet.drawing_rel_id.is_none());
}

#[test]
fn tab_color_variants() {
    // Indexed palette color (xColorType = 1, fValidRGB clear).
    let indexed = cs_prop(0, [0x02, 0x40, 0, 0, 0, 0, 0, 0], "C");
    let data = stream(&[
        (rt::BEGIN_SHEET, Vec::new()),
        (rt::CS_PROP, indexed),
        (rt::END_SHEET, Vec::new()),
    ]);
    let sheet = parse_chart_sheet_part(&data, "C".to_string(), 0).unwrap();
    assert_eq!(sheet.tab_color.color_type, ColorType::Indexed);
    assert_eq!(sheet.tab_color.index, 0x40);

    // Theme color (xColorType = 3).
    let theme = cs_prop(0, [0x06, 0x0B, 0, 0, 0, 0, 0, 0], "C");
    let data = stream(&[
        (rt::BEGIN_SHEET, Vec::new()),
        (rt::CS_PROP, theme),
        (rt::END_SHEET, Vec::new()),
    ]);
    let sheet = parse_chart_sheet_part(&data, "C".to_string(), 0).unwrap();
    assert_eq!(sheet.tab_color.color_type, ColorType::Theme);
    assert_eq!(sheet.tab_color.index, 0x0B);

    // RGB color without the valid bit is rejected.
    let invalid_rgb = cs_prop(0, [0x04, 0, 0, 0, 0, 0, 0, 0], "C");
    let data = stream(&[
        (rt::BEGIN_SHEET, Vec::new()),
        (rt::CS_PROP, invalid_rgb),
        (rt::END_SHEET, Vec::new()),
    ]);
    assert!(matches!(
        parse_chart_sheet_part(&data, "C".to_string(), 0),
        Err(Error::Unrecognized { .. })
    ));

    // Unknown color type is rejected.
    let unknown_type = cs_prop(0, [0x08, 0, 0, 0, 0, 0, 0, 0], "C");
    let data = stream(&[
        (rt::BEGIN_SHEET, Vec::new()),
        (rt::CS_PROP, unknown_type),
        (rt::END_SHEET, Vec::new()),
    ]);
    assert!(matches!(
        parse_chart_sheet_part(&data, "C".to_string(), 0),
        Err(Error::Unrecognized { .. })
    ));

    // Out-of-range theme index is rejected.
    let bad_theme = cs_prop(0, [0x06, 0x0C, 0, 0, 0, 0, 0, 0], "C");
    let data = stream(&[
        (rt::BEGIN_SHEET, Vec::new()),
        (rt::CS_PROP, bad_theme),
        (rt::END_SHEET, Vec::new()),
    ]);
    assert!(matches!(
        parse_chart_sheet_part(&data, "C".to_string(), 0),
        Err(Error::Unrecognized { .. })
    ));
}

#[test]
fn rejects_duplicate_singletons() {
    for (record_type, payload) in [
        (rt::CS_PROP, cs_prop(0, [0; 8], "A")),
        (rt::CS_PAGE_SETUP, cs_page_setup(1, 0, "rIdP")),
        (rt::CS_PROTECTION, cs_protection(0, 0, 0)),
        (rt::DRAWING, wide_string("rIdD")),
    ] {
        let data = chart_sheet_stream(&[(record_type, payload)]);
        assert!(
            matches!(
                parse_chart_sheet_part(&data, "C".to_string(), 0),
                Err(Error::Unrecognized { .. })
            ),
            "duplicate 0x{record_type:04X} must be rejected"
        );
    }
}

#[test]
fn rejects_invalid_view_scale() {
    let data = stream(&[
        (rt::BEGIN_SHEET, Vec::new()),
        (rt::BEGIN_CS_VIEWS, Vec::new()),
        (rt::BEGIN_CS_VIEW, cs_view(0, 5, 0)),
        (rt::END_CS_VIEW, Vec::new()),
        (rt::END_CS_VIEWS, Vec::new()),
        (rt::END_SHEET, Vec::new()),
    ]);
    assert!(matches!(
        parse_chart_sheet_part(&data, "C".to_string(), 0),
        Err(Error::Unrecognized { .. })
    ));
    // 0 means "no zoom level set" and is accepted.
    let data = stream(&[
        (rt::BEGIN_SHEET, Vec::new()),
        (rt::BEGIN_CS_VIEWS, Vec::new()),
        (rt::BEGIN_CS_VIEW, cs_view(0, 0, 0)),
        (rt::END_CS_VIEW, Vec::new()),
        (rt::END_CS_VIEWS, Vec::new()),
        (rt::END_SHEET, Vec::new()),
    ]);
    let sheet = parse_chart_sheet_part(&data, "C".to_string(), 0).unwrap();
    assert_eq!(sheet.views[0].scale, 0);
}

#[test]
fn rejects_invalid_views_collections() {
    // Empty views collection.
    let data = stream(&[
        (rt::BEGIN_SHEET, Vec::new()),
        (rt::BEGIN_CS_VIEWS, Vec::new()),
        (rt::END_CS_VIEWS, Vec::new()),
        (rt::END_SHEET, Vec::new()),
    ]);
    assert!(matches!(
        parse_chart_sheet_part(&data, "C".to_string(), 0),
        Err(Error::Unrecognized { .. })
    ));
    // Unterminated views collection.
    let data = stream(&[
        (rt::BEGIN_SHEET, Vec::new()),
        (rt::BEGIN_CS_VIEWS, Vec::new()),
        (rt::END_SHEET, Vec::new()),
    ]);
    assert!(matches!(
        parse_chart_sheet_part(&data, "C".to_string(), 0),
        Err(Error::UnexpectedEndOfStream(_))
    ));
    // View without its end record.
    let data = stream(&[
        (rt::BEGIN_SHEET, Vec::new()),
        (rt::BEGIN_CS_VIEWS, Vec::new()),
        (rt::BEGIN_CS_VIEW, cs_view(0, 100, 0)),
        (rt::END_CS_VIEWS, Vec::new()),
        (rt::END_SHEET, Vec::new()),
    ]);
    assert!(matches!(
        parse_chart_sheet_part(&data, "C".to_string(), 0),
        Err(Error::UnexpectedRecord { .. })
    ));
}

#[test]
fn enforces_iso_protection_pairing() {
    // ISO record not immediately followed by the classic record.
    let data = stream(&[
        (rt::BEGIN_SHEET, Vec::new()),
        (rt::CS_PROTECTION_ISO, cs_protection_iso(100, 1, 1)),
        (rt::CS_PROP, cs_prop(0, [0; 8], "A")),
        (rt::CS_PROTECTION, cs_protection(0, 1, 1)),
        (rt::END_SHEET, Vec::new()),
    ]);
    assert!(matches!(
        parse_chart_sheet_part(&data, "C".to_string(), 0),
        Err(Error::Unrecognized { .. })
    ));
    // ISO record at the end of the stream.
    let data = stream(&[
        (rt::BEGIN_SHEET, Vec::new()),
        (rt::CS_PROTECTION_ISO, cs_protection_iso(100, 1, 1)),
        (rt::END_SHEET, Vec::new()),
    ]);
    assert!(matches!(
        parse_chart_sheet_part(&data, "C".to_string(), 0),
        Err(Error::Unrecognized { .. })
    ));
    // Following classic record carries a password verifier.
    let data = stream(&[
        (rt::BEGIN_SHEET, Vec::new()),
        (rt::CS_PROTECTION_ISO, cs_protection_iso(100, 1, 1)),
        (rt::CS_PROTECTION, cs_protection(0x1234, 1, 1)),
        (rt::END_SHEET, Vec::new()),
    ]);
    assert!(matches!(
        parse_chart_sheet_part(&data, "C".to_string(), 0),
        Err(Error::Unrecognized { .. })
    ));
    // Following classic record has different flags.
    let data = stream(&[
        (rt::BEGIN_SHEET, Vec::new()),
        (rt::CS_PROTECTION_ISO, cs_protection_iso(100, 1, 1)),
        (rt::CS_PROTECTION, cs_protection(0, 1, 0)),
        (rt::END_SHEET, Vec::new()),
    ]);
    assert!(matches!(
        parse_chart_sheet_part(&data, "C".to_string(), 0),
        Err(Error::Unrecognized { .. })
    ));
    // Excessive spin count is rejected.
    let data = stream(&[
        (rt::BEGIN_SHEET, Vec::new()),
        (rt::CS_PROTECTION_ISO, cs_protection_iso(10_000_001, 1, 1)),
        (rt::CS_PROTECTION, cs_protection(0, 1, 1)),
        (rt::END_SHEET, Vec::new()),
    ]);
    assert!(matches!(
        parse_chart_sheet_part(&data, "C".to_string(), 0),
        Err(Error::Unrecognized { .. })
    ));
}

#[test]
fn rejects_malformed_streams() {
    // Stream must start with BrtBeginSheet.
    let data = stream(&[(rt::END_SHEET, Vec::new())]);
    assert!(matches!(
        parse_chart_sheet_part(&data, "C".to_string(), 0),
        Err(Error::UnexpectedRecord { .. })
    ));
    // Missing BrtEndSheet.
    let data = stream(&[(rt::BEGIN_SHEET, Vec::new())]);
    assert!(matches!(
        parse_chart_sheet_part(&data, "C".to_string(), 0),
        Err(Error::UnexpectedEndOfStream(_))
    ));
    // Truncated BrtCsProp payload.
    let data = stream(&[
        (rt::BEGIN_SHEET, Vec::new()),
        (rt::CS_PROP, vec![0; 5]),
        (rt::END_SHEET, Vec::new()),
    ]);
    assert!(matches!(
        parse_chart_sheet_part(&data, "C".to_string(), 0),
        Err(Error::Wire(WireError::Truncated {
            stage: Stage::Value,
            ..
        }))
    ));
    // Trailing bytes in BrtCsProp.
    let mut payload = cs_prop(0, [0; 8], "A");
    payload.push(0);
    let data = stream(&[
        (rt::BEGIN_SHEET, Vec::new()),
        (rt::CS_PROP, payload),
        (rt::END_SHEET, Vec::new()),
    ]);
    assert!(matches!(
        parse_chart_sheet_part(&data, "C".to_string(), 0),
        Err(Error::Wire(WireError::Trailing {
            context: "BrtCsProp",
            ..
        }))
    ));
    // Empty drawing relationship identifier.
    let data = stream(&[
        (rt::BEGIN_SHEET, Vec::new()),
        (rt::DRAWING, wide_string("")),
        (rt::END_SHEET, Vec::new()),
    ]);
    assert!(matches!(
        parse_chart_sheet_part(&data, "C".to_string(), 0),
        Err(Error::Unrecognized { .. })
    ));
    // iCopies outside the permitted range.
    let data = stream(&[
        (rt::BEGIN_SHEET, Vec::new()),
        (rt::CS_PAGE_SETUP, cs_page_setup(0, 0, "rIdP")),
        (rt::END_SHEET, Vec::new()),
    ]);
    assert!(matches!(
        parse_chart_sheet_part(&data, "C".to_string(), 0),
        Err(Error::Unrecognized { .. })
    ));
    // Truncated record stream (record header cut off).
    let mut data = chart_sheet_stream(&[]);
    data.truncate(data.len() - 1);
    assert!(matches!(
        parse_chart_sheet_part(&data, "C".to_string(), 0),
        Err(Error::Wire(WireError::Truncated {
            stage: Stage::Length,
            ..
        }))
    ));
}

/// Build a synthetic package with one worksheet and one chart sheet whose
/// drawing hosts one chart, and verify the workbook accessors.
#[test]
fn resolves_chart_sheet_and_embedded_chart_through_workbook_relationships() {
    use crate::package::Workbook;
    use litchi_opc::constants::relationship_type;
    use litchi_opc::part::Part;
    use litchi_opc::{BlobPart, OpcPackage, PackURI};

    // workbook.bin declares one worksheet and one chart sheet.
    let mut sheet1 = 0u32.to_le_bytes().to_vec();
    sheet1.extend_from_slice(&1u32.to_le_bytes());
    sheet1.extend_from_slice(&wide_string("rIdSheet1"));
    sheet1.extend_from_slice(&wide_string("Sheet1"));
    let mut chart1 = 0u32.to_le_bytes().to_vec();
    chart1.extend_from_slice(&2u32.to_le_bytes());
    chart1.extend_from_slice(&wide_string("rIdChart1"));
    chart1.extend_from_slice(&wide_string("Chart1"));
    let workbook_data = stream(&[(rt::BUNDLE_SH, sheet1), (rt::BUNDLE_SH, chart1)]);
    let mut workbook_part = BlobPart::new(
        PackURI::new("/xl/workbook.bin").unwrap(),
        "application/vnd.ms-excel.sheet.binary.macroEnabled.main".to_string(),
        workbook_data,
    );
    workbook_part.rels_mut().add_relationship(
        relationship_type::WORKSHEET.to_string(),
        "worksheets/sheet1.bin".to_string(),
        "rIdSheet1".to_string(),
        false,
    );
    workbook_part.rels_mut().add_relationship(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet"
            .to_string(),
        "chartsheets/sheet2.bin".to_string(),
        "rIdChart1".to_string(),
        false,
    );

    let sheet_part = BlobPart::new(
        PackURI::new("/xl/worksheets/sheet1.bin").unwrap(),
        "application/vnd.ms-excel.worksheet".to_string(),
        stream(&[(rt::BEGIN_SHEET, Vec::new()), (rt::END_SHEET, Vec::new())]),
    );

    // The chart sheet links its drawing through BrtDrawing.
    let chart_sheet_data = stream(&[
        (rt::BEGIN_SHEET, Vec::new()),
        (rt::CS_PROP, cs_prop(1, rgb_color(1, 2, 3), "ChartCode")),
        (rt::BEGIN_CS_VIEWS, Vec::new()),
        (rt::BEGIN_CS_VIEW, cs_view(1, 100, 0)),
        (rt::END_CS_VIEW, Vec::new()),
        (rt::END_CS_VIEWS, Vec::new()),
        (rt::DRAWING, wide_string("rIdDrawing")),
        (rt::END_SHEET, Vec::new()),
    ]);
    let mut chart_sheet_part = BlobPart::new(
        PackURI::new("/xl/chartsheets/sheet2.bin").unwrap(),
        "application/vnd.ms-excel.chartsheet".to_string(),
        chart_sheet_data,
    );
    chart_sheet_part.rels_mut().add_relationship(
        relationship_type::DRAWING.to_string(),
        "../drawings/drawing1.xml".to_string(),
        "rIdDrawing".to_string(),
        false,
    );

    let drawing_xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="5000000" cy="3000000"/><xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/></xdr:nvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rIdChart"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:absoluteAnchor></xdr:wsDr>"#;
    let mut drawing_part = BlobPart::new(
        PackURI::new("/xl/drawings/drawing1.xml").unwrap(),
        "application/vnd.openxmlformats-officedocument.drawing+xml".to_string(),
        drawing_xml.to_vec(),
    );
    drawing_part.rels_mut().add_relationship(
        relationship_type::CHART.to_string(),
        "../charts/chart1.xml".to_string(),
        "rIdChart".to_string(),
        false,
    );

    let chart_xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:ser><c:idx val="0"/><c:order val="0"/><c:cat><c:strRef><c:f>Sheet1!$A$1:$A$3</c:f></c:strRef></c:cat><c:val><c:numRef><c:f>Sheet1!$B$1:$B$3</c:f></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"#;
    let chart_part = BlobPart::new(
        PackURI::new("/xl/charts/chart1.xml").unwrap(),
        "application/vnd.openxmlformats-officedocument.drawingml.chart+xml".to_string(),
        chart_xml.to_vec(),
    );

    let workbook_target = workbook_part
        .partname()
        .as_str()
        .trim_start_matches('/')
        .to_owned();
    let mut package = OpcPackage::new();
    package.rels_mut().add_relationship(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
            .to_owned(),
        workbook_target,
        "rIdWorkbook".to_owned(),
        false,
    );
    package.add_part(Box::new(workbook_part));
    package.add_part(Box::new(sheet_part));
    package.add_part(Box::new(chart_sheet_part));
    package.add_part(Box::new(drawing_part));
    package.add_part(Box::new(chart_part));
    let workbook = Workbook::from_opc_package(package).unwrap();

    let chart_sheets = workbook.chart_sheets();
    assert_eq!(chart_sheets.len(), 1);
    assert_eq!(chart_sheets[0].0, 1);
    assert_eq!(chart_sheets[0].1.name, "Chart1");
    assert_eq!(chart_sheets[0].1.code_name, "ChartCode");
    assert!(workbook.chart_sheet(0).is_none());
    assert!(workbook.chart_sheet(1).is_some());

    let drawing = workbook.sheet_drawing(1).unwrap();
    assert_eq!(drawing.drawing.anchors.len(), 1);
    assert_eq!(drawing.charts.len(), 1);
    assert_eq!(drawing.charts[0].frame_name, "Chart 1");
    assert_eq!(drawing.charts[0].rel_id, "rIdChart");
    assert!(!drawing.charts[0].chart.plot_area.type_groups.is_empty());
}

/// A chart sheet whose `BrtDrawing` relationship is broken surfaces the
/// error, matching the eager failure handling of PivotCache definitions.
#[test]
fn broken_chart_sheet_drawing_relationship_is_an_error() {
    use crate::package::Workbook;
    use litchi_opc::part::Part;
    use litchi_opc::{BlobPart, OpcPackage, PackURI};

    let mut chart1 = 0u32.to_le_bytes().to_vec();
    chart1.extend_from_slice(&1u32.to_le_bytes());
    chart1.extend_from_slice(&wide_string("rIdChart1"));
    chart1.extend_from_slice(&wide_string("Chart1"));
    let workbook_data = stream(&[(rt::BUNDLE_SH, chart1)]);
    let mut workbook_part = BlobPart::new(
        PackURI::new("/xl/workbook.bin").unwrap(),
        "application/vnd.ms-excel.sheet.binary.macroEnabled.main".to_string(),
        workbook_data,
    );
    workbook_part.rels_mut().add_relationship(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet"
            .to_string(),
        "chartsheets/sheet1.bin".to_string(),
        "rIdChart1".to_string(),
        false,
    );
    // The chart sheet references a drawing relationship that does not exist.
    let chart_sheet_data = stream(&[
        (rt::BEGIN_SHEET, Vec::new()),
        (rt::DRAWING, wide_string("rIdMissing")),
        (rt::END_SHEET, Vec::new()),
    ]);
    let chart_sheet_part = BlobPart::new(
        PackURI::new("/xl/chartsheets/sheet1.bin").unwrap(),
        "application/vnd.ms-excel.chartsheet".to_string(),
        chart_sheet_data,
    );

    let mut package = OpcPackage::new();
    package.add_part(Box::new(workbook_part));
    package.add_part(Box::new(chart_sheet_part));
    assert!(Workbook::from_opc_package(package).is_err());
}
