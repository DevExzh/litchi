//! Tests for the SpreadsheetDrawing XML inventory parser.

use super::*;

const XDR_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const A_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const C_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const R_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

fn drawing(body: &str) -> Vec<u8> {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<xdr:wsDr xmlns:xdr="{XDR_NS}" xmlns:a="{A_NS}" xmlns:c="{C_NS}" xmlns:r="{R_NS}">{body}</xdr:wsDr>"#
    )
    .into_bytes()
}

fn marker_xml(tag: &str, col: u32, col_off: i64, row: u32, row_off: i64) -> String {
    format!(
        "<xdr:{tag}><xdr:col>{col}</xdr:col><xdr:colOff>{col_off}</xdr:colOff><xdr:row>{row}</xdr:row><xdr:rowOff>{row_off}</xdr:rowOff></xdr:{tag}>"
    )
}

#[test]
fn parses_two_cell_shape_anchor() {
    let body = format!(
        "<xdr:twoCellAnchor editAs=\"oneCell\">{}{}<xdr:sp macro=\"\" textlink=\"\">\
         <xdr:nvSpPr><xdr:cNvPr id=\"1026\" name=\"shapetype_75\" hidden=\"1\"/><xdr:cNvSpPr/></xdr:nvSpPr>\
         <xdr:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"9525000\" cy=\"9525000\"/></a:xfrm>\
         <a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l=\"l\" t=\"t\" r=\"r\" b=\"b\"/>\
         <a:pathLst><a:path w=\"21600\" h=\"21600\"><a:moveTo><a:pt x=\"0\" y=\"0\"/></a:moveTo></a:path></a:pathLst>\
         </a:custGeom></xdr:spPr><xdr:txBody><a:bodyPr/><a:p/></xdr:txBody></xdr:sp><xdr:clientData/></xdr:twoCellAnchor>",
        marker_xml("from", 0, 0, 0, 0),
        marker_xml("to", 12, 266700, 50, 114300),
    );
    let inventory = parse_drawing_part(&drawing(&body)).unwrap();
    assert_eq!(inventory.anchors.len(), 1);
    let XlsbDrawingAnchorKind::TwoCell { from, to, edit_as } = &inventory.anchors[0].anchor else {
        panic!("expected two-cell anchor");
    };
    assert_eq!(from.column, 0);
    assert_eq!(from.row, 0);
    assert_eq!(to.column, 12);
    assert_eq!(to.column_offset, 266700);
    assert_eq!(to.row, 50);
    assert_eq!(to.row_offset, 114300);
    assert_eq!(edit_as.as_deref(), Some("oneCell"));
    assert_eq!(
        inventory.anchors[0].object,
        XlsbDrawingObject::Shape(XlsbDrawingNonVisual {
            id: 1026,
            name: "shapetype_75".to_string(),
            description: None,
        })
    );
}

#[test]
fn parses_one_cell_picture_anchor() {
    let body = format!(
        "<xdr:oneCellAnchor>{}<xdr:ext cx=\"1219200\" cy=\"914400\"/><xdr:pic>\
         <xdr:nvPicPr><xdr:cNvPr id=\"3\" name=\"Logo\" descr=\"Company logo\"/><xdr:cNvPicPr/></xdr:nvPicPr>\
         <xdr:blipFill><a:blip r:embed=\"rIdImage1\"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>\
         <xdr:spPr/></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>",
        marker_xml("from", 2, 19050, 4, 9525),
    );
    let inventory = parse_drawing_part(&drawing(&body)).unwrap();
    assert_eq!(inventory.anchors.len(), 1);
    let XlsbDrawingAnchorKind::OneCell { from, extent } = &inventory.anchors[0].anchor else {
        panic!("expected one-cell anchor");
    };
    assert_eq!(from.column, 2);
    assert_eq!(from.column_offset, 19050);
    assert_eq!(extent.width, 1219200);
    assert_eq!(extent.height, 914400);
    assert_eq!(
        inventory.anchors[0].object,
        XlsbDrawingObject::Picture {
            non_visual: XlsbDrawingNonVisual {
                id: 3,
                name: "Logo".to_string(),
                description: Some("Company logo".to_string()),
            },
            embed_rel_id: Some("rIdImage1".to_string()),
        }
    );
}

#[test]
fn parses_absolute_chart_graphic_frame() {
    let body = "<xdr:absoluteAnchor><xdr:pos x=\"0\" y=\"0\"/><xdr:ext cx=\"5000000\" cy=\"3000000\"/>\
        <xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id=\"2\" name=\"Chart 1\"/>\
        <xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm/>\
        <a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">\
        <c:chart r:id=\"rIdChart\"/></a:graphicData></a:graphic></xdr:graphicFrame>\
        <xdr:clientData/></xdr:absoluteAnchor>";
    let inventory = parse_drawing_part(&drawing(body)).unwrap();
    assert_eq!(inventory.anchors.len(), 1);
    let XlsbDrawingAnchorKind::Absolute { position, extent } = &inventory.anchors[0].anchor else {
        panic!("expected absolute anchor");
    };
    assert_eq!(position.x, 0);
    assert_eq!(extent.width, 5000000);
    let XlsbDrawingObject::GraphicFrame(frame) = &inventory.anchors[0].object else {
        panic!("expected graphic frame");
    };
    assert_eq!(frame.non_visual.name, "Chart 1");
    assert_eq!(frame.content_uri, CHART_GRAPHIC_DATA_URI);
    assert!(frame.is_chart());
    assert_eq!(frame.rel_id.as_deref(), Some("rIdChart"));
}

#[test]
fn parses_connection_and_group_shapes() {
    let body = format!(
        "<xdr:twoCellAnchor>{}{}<xdr:cxnSp><xdr:nvCxnSpPr><xdr:cNvPr id=\"4\" name=\"Connector\"/>\
         <xdr:cNvCxnSpPr/></xdr:nvCxnSpPr><xdr:spPr/></xdr:cxnSp><xdr:clientData/></xdr:twoCellAnchor>\
         <xdr:twoCellAnchor>{}{}<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id=\"5\" name=\"Group\"/>\
         <xdr:cNvGrpSpPr/></xdr:nvGrpSpPr><xdr:grpSpPr/>\
         <xdr:sp><xdr:nvSpPr><xdr:cNvPr id=\"6\" name=\"Nested\"/><xdr:cNvSpPr/></xdr:nvSpPr>\
         <xdr:spPr/></xdr:sp></xdr:grpSp><xdr:clientData/></xdr:twoCellAnchor>",
        marker_xml("from", 0, 0, 0, 0),
        marker_xml("to", 3, 0, 3, 0),
        marker_xml("from", 1, 0, 1, 0),
        marker_xml("to", 4, 0, 4, 0),
    );
    let inventory = parse_drawing_part(&drawing(&body)).unwrap();
    assert_eq!(inventory.anchors.len(), 2);
    assert_eq!(
        inventory.anchors[0].object,
        XlsbDrawingObject::ConnectionShape(XlsbDrawingNonVisual {
            id: 4,
            name: "Connector".to_string(),
            description: None,
        })
    );
    // The group keeps its own non-visual properties; nested shapes are not
    // inventoried.
    assert_eq!(
        inventory.anchors[1].object,
        XlsbDrawingObject::GroupShape(XlsbDrawingNonVisual {
            id: 5,
            name: "Group".to_string(),
            description: None,
        })
    );
}

#[test]
fn empty_drawing_has_no_anchors() {
    let inventory = parse_drawing_part(&drawing("")).unwrap();
    assert!(inventory.anchors.is_empty());
    let inventory =
        parse_drawing_part(format!(r#"<xdr:wsDr xmlns:xdr="{XDR_NS}"/>"#).as_bytes()).unwrap();
    assert!(inventory.anchors.is_empty());
}

#[test]
fn rejects_malformed_drawings() {
    // Root is not xdr:wsDr.
    assert!(parse_drawing_part(b"<html/>").is_err());
    // Not UTF-8.
    assert!(parse_drawing_part(&[0xFF, 0xFE, 0x00]).is_err());
    // Missing from marker.
    let body = "<xdr:twoCellAnchor><xdr:sp><xdr:nvSpPr><xdr:cNvPr id=\"1\" name=\"s\"/>\
        <xdr:cNvSpPr/></xdr:nvSpPr></xdr:sp></xdr:twoCellAnchor>";
    assert!(parse_drawing_part(&drawing(body)).is_err());
    // Anchor without an object.
    let body = format!(
        "<xdr:twoCellAnchor>{}{}<xdr:clientData/></xdr:twoCellAnchor>",
        marker_xml("from", 0, 0, 0, 0),
        marker_xml("to", 1, 0, 1, 0),
    );
    assert!(parse_drawing_part(&drawing(&body)).is_err());
    // Duplicate from markers.
    let body = format!(
        "<xdr:twoCellAnchor>{}{}<xdr:sp/></xdr:twoCellAnchor>",
        marker_xml("from", 0, 0, 0, 0),
        marker_xml("from", 0, 0, 0, 0),
    );
    assert!(parse_drawing_part(&drawing(&body)).is_err());
    // Invalid marker text.
    let body = "<xdr:twoCellAnchor><xdr:from><xdr:col>x</xdr:col></xdr:from></xdr:twoCellAnchor>";
    assert!(parse_drawing_part(&drawing(body)).is_err());
    // Unterminated XML.
    assert!(
        parse_drawing_part(
            format!(r#"<xdr:wsDr xmlns:xdr="{XDR_NS}"><xdr:twoCellAnchor>"#).as_bytes()
        )
        .is_err()
    );
}

/// The real `universal-content.xlsb` fixture exposes its worksheet drawing
/// through the workbook wiring: one shape anchored on the first sheet.
#[test]
fn parses_real_fixture_workbook_drawing() {
    let path = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/ooxml/xlsb/universal-content.xlsb"
    );
    let workbook = crate::xlsb::XlsbWorkbook::new(std::fs::File::open(path).unwrap()).unwrap();
    let drawing = workbook.sheet_drawing(0).expect("sheet 0 has a drawing");
    assert_eq!(drawing.sheet_index, 0);
    assert_eq!(drawing.drawing.anchors.len(), 1);
    assert_eq!(
        drawing.drawing.anchors[0].object,
        XlsbDrawingObject::Shape(XlsbDrawingNonVisual {
            id: 1026,
            name: "shapetype_75".to_string(),
            description: None,
        })
    );
    let XlsbDrawingAnchorKind::TwoCell { from, to, .. } = &drawing.drawing.anchors[0].anchor else {
        panic!("expected two-cell anchor");
    };
    assert_eq!(from.row, 0);
    assert_eq!(to.row, 50);
    assert!(drawing.charts.is_empty());
    assert!(workbook.chart_sheets().is_empty());
    assert!(workbook.sheet_drawing(1).is_none());
}
