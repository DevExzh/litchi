//! Integration-style tests for package relationship shape loading.

use super::*;
use crate::shapes::{Anchor, Object, TextSize};
use litchi_drawingml::geom::Preset;
use litchi_opc::BlobPart;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{OpcPackage, PackURI};

const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

const POI_TEXT_BOXES: &[u8] =
    include_bytes!("../../../../../test-data/poi/test-data/spreadsheet/45540_form_Header.xlsx");

fn marker(col: u32, col_off: i64, row: u32, row_off: i64) -> String {
    format!(
        "<xdr:col>{col}</xdr:col><xdr:colOff>{col_off}</xdr:colOff>\\
             <xdr:row>{row}</xdr:row><xdr:rowOff>{row_off}</xdr:rowOff>"
    )
}

fn drawing(body: &str) -> String {
    format!(
        "<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" \\
             xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" \\
             xmlns:r=\"{R}\">{body}</xdr:wsDr>"
    )
}

fn two_cell_anchor(object: &str) -> String {
    format!(
        "<xdr:twoCellAnchor editAs=\"oneCell\"><xdr:from>{}</xdr:from><xdr:to>{}</xdr:to>\\
             {object}<xdr:clientData fLocksWithSheet=\"0\" fPrintsWithSheet=\"1\"/></xdr:twoCellAnchor>",
        marker(1, 100, 2, 200),
        marker(5, 300, 9, 400)
    )
}

fn text_box_shape() -> &'static str {
    "<xdr:sp macro=\"\" textlink=\"\">\\
         <xdr:nvSpPr><xdr:cNvPr id=\"7\" name=\"Text Box 7\" descr=\"alt\" hidden=\"1\"/>\\
         <xdr:cNvSpPr txBox=\"1\"><a:spLocks noChangeArrowheads=\"1\"/></xdr:cNvSpPr></xdr:nvSpPr>\\
         <xdr:spPr><a:prstGeom prst=\"roundRect\"><a:avLst/></a:prstGeom></xdr:spPr>\\
         <xdr:txBody><a:bodyPr lIns=\"182880\" tIns=\"91440\" rIns=\"182880\" bIns=\"91440\" \\
         anchor=\"ctr\" anchorCtr=\"1\" vert=\"vert270\" wrap=\"none\" numCol=\"2\" \\
         spcFirstLastPara=\"1\"><a:spAutoFit/></a:bodyPr><a:lstStyle/>\\
         <a:p><a:pPr algn=\"l\"><a:defRPr sz=\"1000\"/></a:pPr>\\
         <a:r><a:rPr lang=\"en-US\" sz=\"1200\" b=\"1\" i=\"true\" u=\"sng\"/><a:t>Bold</a:t></a:r>\\
         <a:r><a:t xml:space=\"preserve\"> plain</a:t></a:r><a:br/></a:p>\\
         <a:p><a:r><a:t>Second</a:t></a:r></a:p>\\
         </xdr:txBody></xdr:sp>"
}

fn package_with_shapes(drawing_xml: &str) -> OpcPackage {
    let mut package = OpcPackage::new();
    let mut workbook_part = BlobPart::new(
        PackURI::new("/xl/workbook.xml").unwrap(),
        ct::SML_SHEET_MAIN.to_string(),
        format!(
            "<workbook xmlns=\"{SML}\" xmlns:r=\"{R}\">\\
                 <sheets><sheet name=\"Data\" sheetId=\"1\" r:id=\"rId1\"/>\\
                 <sheet name=\"Empty\" sheetId=\"2\" r:id=\"rId2\"/></sheets></workbook>"
        )
        .into_bytes(),
    );
    workbook_part.relate_to("worksheets/sheet1.xml", rt::WORKSHEET);
    workbook_part.relate_to("worksheets/sheet2.xml", rt::WORKSHEET);
    let mut sheet_part = BlobPart::new(
        PackURI::new("/xl/worksheets/sheet1.xml").unwrap(),
        ct::SML_WORKSHEET.to_string(),
        format!("<worksheet xmlns=\"{SML}\"><sheetData/></worksheet>").into_bytes(),
    );
    sheet_part.relate_to("../drawings/drawing1.xml", rt::DRAWING);
    package.relate_to(
        "xl/workbook.xml",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
    );
    package.add_part(Box::new(workbook_part));
    package.add_part(Box::new(sheet_part));
    package.add_part(Box::new(BlobPart::new(
        PackURI::new("/xl/worksheets/sheet2.xml").unwrap(),
        ct::SML_WORKSHEET.to_string(),
        format!("<worksheet xmlns=\"{SML}\"><sheetData/></worksheet>").into_bytes(),
    )));
    package.add_part(Box::new(BlobPart::new(
        PackURI::new("/xl/drawings/drawing1.xml").unwrap(),
        ct::OFC_DRAWING.to_string(),
        drawing_xml.as_bytes().to_vec(),
    )));
    package
}

#[test]
fn loads_shapes_through_the_package_graph() {
    let package = package_with_shapes(&drawing(&two_cell_anchor(text_box_shape())));
    let shapes = load_sheet_shapes(&package, "Data").unwrap();
    assert_eq!(shapes.worksheet_name, "Data");
    assert_eq!(shapes.worksheet_part_name, "/xl/worksheets/sheet1.xml");
    assert_eq!(shapes.objects.len(), 1);
    let Object::Shape(shape) = &shapes.objects[0].object else {
        panic!("expected a shape");
    };
    assert_eq!(shape.non_visual.name.as_deref(), Some("Text Box 7"));

    // Worksheets without shapes are omitted from the workbook inventory.
    let all = load_shapes(&package).unwrap();
    assert_eq!(all.len(), 1);
    assert_eq!(all[0].worksheet_name, "Data");
    assert!(
        load_sheet_shapes(&package, "Empty")
            .unwrap()
            .objects
            .is_empty()
    );
    assert!(load_sheet_shapes(&package, "Missing").is_err());
}

#[test]
fn poi_fixture_text_boxes_parse() {
    let package = OpcPackage::from_bytes(POI_TEXT_BOXES).unwrap();
    let all = load_shapes(&package).unwrap();
    let names: Vec<_> = all
        .iter()
        .flat_map(|sheet| sheet.objects.iter())
        .filter_map(|anchored| match &anchored.object {
            Object::Shape(shape) => Some(shape),
            _ => None,
        })
        .collect();
    assert_eq!(names.len(), 4);
    let text_box = names
        .iter()
        .find(|shape| shape.non_visual.name.as_deref() == Some("Text Box 35"))
        .unwrap();
    assert!(text_box.is_text_box);
    assert!(text_box.non_visual.locked);
    assert!(!text_box.non_visual.hidden);
    assert_eq!(text_box.preset(), Some(Preset::Rect));
    let body = text_box.text_body.as_ref().unwrap();
    assert_eq!(body.properties.insets.left.as_emu(), Some(27432));
    assert_eq!(body.text(), "State-Owned Enterprise");
    let run = &body.paragraphs[0].runs[0];
    assert_eq!(run.bold, Some(false));
    assert_eq!(run.font_size.map(TextSize::get), Some(900));
    // All four fixture text boxes anchor with two-cell anchors.
    assert!(
        all.iter()
            .flat_map(|sheet| sheet.objects.iter())
            .all(|anchored| matches!(anchored.anchor, Anchor::TwoCell { .. }))
    );
}
