#![allow(clippy::unwrap_used, reason = "test assertions require direct values")]

//! Regression tests for typed `SpreadsheetDrawing` models and bounded codecs.

use super::*;
use litchi_drawingml::geom::Preset;
use std::mem::size_of;

const XDR: &str = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const STRICT_XDR: &str = "http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing";
const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const STRICT_A: &str = "http://purl.oclc.org/ooxml/drawingml/main";
const C: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

fn marker(col: u32, col_off: i64, row: u32, row_off: i64) -> String {
    format!(
        "<xdr:col>{col}</xdr:col><xdr:colOff>{col_off}</xdr:colOff>\
             <xdr:row>{row}</xdr:row><xdr:rowOff>{row_off}</xdr:rowOff>"
    )
}

fn drawing(body: &str) -> String {
    format!("<xdr:wsDr xmlns:xdr=\"{XDR}\" xmlns:a=\"{A}\" xmlns:r=\"{R}\">{body}</xdr:wsDr>")
}

fn two_cell_anchor(object: &str) -> String {
    format!(
        "<xdr:twoCellAnchor editAs=\"oneCell\"><xdr:from>{}</xdr:from><xdr:to>{}</xdr:to>\
             {object}<xdr:clientData fLocksWithSheet=\"0\" fPrintsWithSheet=\"1\"/></xdr:twoCellAnchor>",
        marker(1, 100, 2, 200),
        marker(5, 300, 9, 400)
    )
}

fn text_box_shape() -> &'static str {
    "<xdr:sp macro=\"\" textlink=\"\">\
         <xdr:nvSpPr><xdr:cNvPr id=\"7\" name=\"Text Box 7\" descr=\"alt\" hidden=\"1\"/>\
         <xdr:cNvSpPr txBox=\"1\"><a:spLocks noChangeArrowheads=\"1\"/></xdr:cNvSpPr></xdr:nvSpPr>\
         <xdr:spPr><a:prstGeom prst=\"roundRect\"><a:avLst/></a:prstGeom></xdr:spPr>\
         <xdr:txBody><a:bodyPr lIns=\"182880\" tIns=\"91440\" rIns=\"182880\" bIns=\"91440\" \
         anchor=\"ctr\" anchorCtr=\"1\" vert=\"vert270\" wrap=\"none\" numCol=\"2\" \
         spcFirstLastPara=\"1\"><a:spAutoFit/></a:bodyPr><a:lstStyle/>\
         <a:p><a:pPr algn=\"l\"><a:defRPr sz=\"1000\"/></a:pPr>\
         <a:r><a:rPr lang=\"en-US\" sz=\"1200\" b=\"1\" i=\"true\" u=\"sng\"/><a:t>Bold</a:t></a:r>\
         <a:r><a:t xml:space=\"preserve\"> plain</a:t></a:r><a:br/></a:p>\
         <a:p><a:r><a:t>Second</a:t></a:r></a:p>\
         </xdr:txBody></xdr:sp>"
}

#[test]
fn parses_two_cell_text_box() {
    let xml = drawing(&two_cell_anchor(text_box_shape()));
    let objects = read(&xml).unwrap().unwrap();
    assert_eq!(objects.len(), 1);
    let anchored = &objects[0];
    assert_eq!(
        anchored.anchor,
        Anchor::TwoCell {
            from: CellMarker {
                column: 1,
                column_offset: Emu(100),
                row: 2,
                row_offset: Emu(200),
            },
            to: CellMarker {
                column: 5,
                column_offset: Emu(300),
                row: 9,
                row_offset: Emu(400),
            },
            edit_as: EditAs::OneCell,
        }
    );
    assert_eq!(anchored.client_data.locks_with_sheet, Some(false));
    assert_eq!(anchored.client_data.prints_with_sheet, Some(true));
    let Object::Shape(shape) = &anchored.object else {
        panic!("expected a shape");
    };
    assert_eq!(shape.non_visual.id, Some(7));
    assert_eq!(shape.non_visual.name.as_deref(), Some("Text Box 7"));
    assert_eq!(shape.non_visual.description.as_deref(), Some("alt"));
    assert!(shape.non_visual.hidden);
    assert!(shape.non_visual.locked);
    assert!(shape.is_text_box);
    assert_eq!(shape.preset(), Some(Preset::RoundRect));
    let body = shape.text_body.as_ref().unwrap();
    let properties = &body.properties;
    assert_eq!(properties.insets.left.as_emu(), Some(182_880));
    assert_eq!(properties.insets.bottom.as_emu(), Some(91440));
    assert_eq!(properties.vertical_anchor, VerticalAnchor::Center);
    assert!(properties.anchor_center);
    assert_eq!(properties.direction, Direction::Vertical270);
    assert_eq!(properties.wrap, Wrap::None);
    assert_eq!(properties.autofit, Autofit::Shape);
    assert_eq!(properties.column_count.get(), 2);
    assert!(properties.space_first_last_paragraph);
    assert_eq!(body.paragraphs.len(), 2);
    let bold = &body.paragraphs[0].runs[0];
    assert_eq!(bold.text, "Bold");
    assert_eq!(bold.bold, Some(true));
    assert_eq!(bold.italic, Some(true));
    assert_eq!(bold.underline, Some(Underline::Single));
    assert_eq!(bold.font_size.map(TextSize::get), Some(1200));
    assert_eq!(body.paragraphs[0].runs[1].text, " plain");
    assert_eq!(body.paragraphs[0].runs[1].bold, None);
    // The break contributes a newline run.
    assert_eq!(body.paragraphs[0].runs[2].text, "\n");
    assert_eq!(body.text(), "Bold plain\n\nSecond");
}

#[test]
fn preset_attributes_apply_xml_schema_token_whitespace() {
    let shape_xml =
        text_box_shape().replace("prst=\"roundRect\"", "prst=\" &#x9;roundRect&#xA;&#xD; \"");
    let objects = read(&drawing(&two_cell_anchor(&shape_xml)))
        .unwrap()
        .unwrap();
    let Object::Shape(shape) = &objects[0].object else {
        panic!("expected a shape");
    };
    assert_eq!(shape.preset(), Some(Preset::RoundRect));
}

#[test]
fn parses_one_cell_connection_shape() {
    let object = "<xdr:cxnSp><xdr:nvCxnSpPr><xdr:cNvPr id=\"9\" name=\"Connector 9\"/>\
            <xdr:cNvCxnSpPr><a:stCxn id=\"7\" idx=\"3\"/><a:endCxn id=\"11\" idx=\"1\"/></xdr:cNvCxnSpPr>\
            </xdr:nvCxnSpPr><xdr:spPr><a:prstGeom prst=\"bentConnector3\"/></xdr:spPr></xdr:cxnSp>";
    let anchor = format!(
        "<xdr:oneCellAnchor><xdr:from>{}</xdr:from>\
             <xdr:ext cx=\"914400\" cy=\"457200\"/>{object}<xdr:clientData/></xdr:oneCellAnchor>",
        marker(0, 0, 0, 0)
    );
    let objects = read(&drawing(&anchor)).unwrap().unwrap();
    assert_eq!(objects.len(), 1);
    assert_eq!(
        objects[0].anchor,
        Anchor::OneCell {
            from: CellMarker::default(),
            extent: EmuExtent {
                width: Emu(914_400),
                height: Emu(457_200),
            },
        }
    );
    assert_eq!(objects[0].client_data, ClientData::default());
    let Object::ConnectionShape(connection) = &objects[0].object else {
        panic!("expected a connection shape");
    };
    assert_eq!(connection.non_visual.id, Some(9));
    assert!(!connection.non_visual.locked);
    assert_eq!(connection.preset(), Some(Preset::BentConnector3));
    assert_eq!(
        connection.start,
        Some(ConnectionEnd {
            shape_id: 7,
            site: 3,
        })
    );
    assert_eq!(
        connection.end,
        Some(ConnectionEnd {
            shape_id: 11,
            site: 1,
        })
    );
    assert!(connection.text_body.is_none());
}

#[test]
fn connection_custom_geometry_is_typed_and_exclusive() {
    let object = "<xdr:cxnSp><xdr:nvCxnSpPr><xdr:cNvPr id=\"9\" name=\"Custom\"/>\
            <xdr:cNvCxnSpPr/></xdr:nvCxnSpPr><xdr:spPr>\
            <a:custGeom><a:pathLst/></a:custGeom></xdr:spPr></xdr:cxnSp>";
    let anchor = format!(
        "<xdr:oneCellAnchor><xdr:from>{}</xdr:from>\
             <xdr:ext cx=\"914400\" cy=\"457200\"/>{object}<xdr:clientData/></xdr:oneCellAnchor>",
        marker(0, 0, 0, 0)
    );
    let objects = read(&drawing(&anchor)).unwrap().unwrap();
    let Object::ConnectionShape(connection) = &objects[0].object else {
        panic!("expected a connection shape");
    };
    assert_eq!(connection.preset(), None);
    assert!(connection.custom_geometry().is_some());
}

#[test]
fn rejects_unknown_preset() {
    let object = "<xdr:sp><xdr:nvSpPr><xdr:cNvPr id=\"3\" name=\"Odd\"/><xdr:cNvSpPr/>\
            </xdr:nvSpPr><xdr:spPr><a:prstGeom prst=\"vendorWeird\"/></xdr:spPr></xdr:sp>";
    let anchor = format!(
        "<xdr:absoluteAnchor><xdr:pos x=\"123\" y=\"456\"/>\
             <xdr:ext cx=\"789\" cy=\"101\"/>{object}<xdr:clientData/></xdr:absoluteAnchor>"
    );
    let error = read(&drawing(&anchor)).unwrap_err();
    assert!(error.to_string().contains("vendorWeird"));
}

#[test]
fn rejects_missing_presets_and_invalid_text_domains() {
    let missing = "<xdr:sp><xdr:nvSpPr><xdr:cNvPr id=\"3\" name=\"Odd\"/><xdr:cNvSpPr/>\
            </xdr:nvSpPr><xdr:spPr><a:prstGeom/></xdr:spPr></xdr:sp>";
    let anchor = format!(
        "<xdr:absoluteAnchor><xdr:pos x=\"123\" y=\"456\"/>\
             <xdr:ext cx=\"789\" cy=\"101\"/>{missing}<xdr:clientData/></xdr:absoluteAnchor>"
    );
    let error = read(&drawing(&anchor)).unwrap_err();
    assert!(error.to_string().contains("missing required prst"));

    assert!(read(&drawing("<xdr:twoCellAnchor editAs=\"vendor\"/>")).is_err());

    for attribute in [
        "anchor=\"middle\"",
        "vert=\"diagonal\"",
        "wrap=\"tight\"",
        "anchorCtr=\"on\"",
        "numCol=\"17\"",
        "lIns=\"2147483648\"",
    ] {
        let object = format!(
            "<xdr:sp><xdr:nvSpPr><xdr:cNvPr id=\"3\" name=\"Text\"/><xdr:cNvSpPr/>\
                 </xdr:nvSpPr><xdr:txBody><a:bodyPr {attribute}/><a:p/></xdr:txBody></xdr:sp>"
        );
        assert!(
            read(&drawing(&two_cell_anchor(&object))).is_err(),
            "accepted {attribute}"
        );
    }

    for run_properties in ["u=\"vendor\"", "sz=\"99\"", "b=\"on\""] {
        let object = format!(
            "<xdr:sp><xdr:nvSpPr><xdr:cNvPr id=\"3\" name=\"Text\"/><xdr:cNvSpPr/>\
                 </xdr:nvSpPr><xdr:txBody><a:bodyPr/><a:p><a:r><a:rPr {run_properties}/>\
                 <a:t>x</a:t></a:r></a:p></xdr:txBody></xdr:sp>"
        );
        assert!(
            read(&drawing(&two_cell_anchor(&object))).is_err(),
            "accepted {run_properties}"
        );
    }
}

#[test]
fn rejects_competing_shape_geometries() {
    let object = "<xdr:sp><xdr:nvSpPr><xdr:cNvPr id=\"3\" name=\"Both\"/><xdr:cNvSpPr/>\
            </xdr:nvSpPr><xdr:spPr><a:prstGeom prst=\"rect\"/>\
            <a:custGeom><a:pathLst/></a:custGeom></xdr:spPr></xdr:sp>";
    let anchor = format!(
        "<xdr:absoluteAnchor><xdr:pos x=\"123\" y=\"456\"/>\
             <xdr:ext cx=\"789\" cy=\"101\"/>{object}<xdr:clientData/></xdr:absoluteAnchor>"
    );
    let error = read(&drawing(&anchor)).unwrap_err();
    assert!(error.to_string().contains("competing"));
}

#[test]
fn retains_unknown_objects_attributes_and_drawingml_children() {
    let unknown_xml = "<xdr:futureShape vendorFlag=\"1\"><xdr:futureChild/></xdr:futureShape>";
    let known = "<xdr:sp><xdr:nvSpPr><xdr:cNvPr id=\"4\" name=\"Known\" vendor=\"keep\"/>\
            <xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:prstGeom prst=\"rect\"/>\
            <a:futureDrawingML vendor=\"keep\"><a:futureNested/></a:futureDrawingML></xdr:spPr></xdr:sp>";
    let xml = drawing(&format!(
        "{}{}",
        two_cell_anchor(unknown_xml),
        two_cell_anchor(known)
    ));
    let objects = read(&xml).unwrap().unwrap();
    assert_eq!(objects.len(), 2);
    let Object::Unknown(unknown) = &objects[0].object else {
        panic!("expected an unknown drawing object");
    };
    assert!(
        std::str::from_utf8(unknown.as_xml())
            .unwrap()
            .contains("futureChild")
    );

    let Object::Shape(shape) = &objects[1].object else {
        panic!("expected a typed shape");
    };
    assert_eq!(shape.non_visual.opaque.attributes()[0].name(), "vendor");
    assert_eq!(shape.non_visual.opaque.attributes()[0].value(), "keep");
    assert!(
        std::str::from_utf8(shape.non_visual.opaque.elements()[0].as_xml())
            .unwrap()
            .contains("futureNested")
    );
}

#[test]
fn geometry_keeps_the_custom_payload_off_the_hot_path() {
    assert!(size_of::<Geometry>() <= 2 * size_of::<usize>());
}

#[test]
fn parses_nested_groups() {
    let inner_xml = "<xdr:sp><xdr:nvSpPr><xdr:cNvPr id=\"21\" name=\"Inner\"/><xdr:cNvSpPr/>\
            </xdr:nvSpPr><xdr:spPr><a:prstGeom prst=\"ellipse\"/></xdr:spPr></xdr:sp>";
    let nested_group = "<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id=\"22\" name=\"Nested\"/>\
             <xdr:cNvGrpSpPr><a:grpSpLocks noChangeAspect=\"1\"/></xdr:cNvGrpSpPr></xdr:nvGrpSpPr>\
             <xdr:grpSpPr/><xdr:sp><xdr:nvSpPr><xdr:cNvPr id=\"23\" name=\"Deep\"/><xdr:cNvSpPr/>\
             </xdr:nvSpPr><xdr:spPr/></xdr:sp></xdr:grpSp>"
        .to_string();
    let group_xml = format!(
        "<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id=\"20\" name=\"Group\"/><xdr:cNvGrpSpPr/>\
             </xdr:nvGrpSpPr><xdr:grpSpPr><a:xfrm><a:off x=\"1\" y=\"2\"/><a:ext cx=\"3\" cy=\"4\"/>\
             <a:chOff x=\"5\" y=\"6\"/><a:chExt cx=\"7\" cy=\"8\"/></a:xfrm></xdr:grpSpPr>\
             {inner_xml}{nested_group}</xdr:grpSp>"
    );
    let objects = read(&drawing(&two_cell_anchor(&group_xml)))
        .unwrap()
        .unwrap();
    let Object::Group(group) = &objects[0].object else {
        panic!("expected a group");
    };
    assert_eq!(group.non_visual.id, Some(20));
    let transform = group.transform.unwrap();
    assert_eq!(transform.offset.unwrap().x, Emu(1));
    assert_eq!(transform.extent.unwrap().height, Emu(4));
    assert_eq!(transform.child_offset.unwrap().y, Emu(6));
    assert_eq!(transform.child_extent.unwrap().width, Emu(7));
    assert_eq!(group.children.len(), 2);
    let Object::Shape(inner) = &group.children[0] else {
        panic!("expected an inner shape");
    };
    assert_eq!(inner.preset(), Some(Preset::Ellipse));
    let Object::Group(nested) = &group.children[1] else {
        panic!("expected a nested group");
    };
    assert!(nested.non_visual.locked);
    assert!(nested.transform.is_none());
    assert_eq!(nested.children.len(), 1);
    assert_eq!(nested.non_visual.name.as_deref(), Some("Nested"));
}

#[test]
fn skips_pictures_and_charts_but_keeps_ole_objects() {
    let picture = "<xdr:pic><xdr:nvPicPr><xdr:cNvPr id=\"50\" name=\"Logo\"/></xdr:nvPicPr>\
            <xdr:blipFill><a:blip r:embed=\"rId9\"/></xdr:blipFill></xdr:pic>";
    let chart = &format!(
        "<xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id=\"51\" name=\"Chart\"/>\
            </xdr:nvGraphicFramePr><a:graphic><a:graphicData>\
            <c:chart xmlns:c=\"{C}\" r:id=\"rId8\"/></a:graphicData></a:graphic></xdr:graphicFrame>"
    );
    let ole = "<xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id=\"52\" name=\"Object\"/>\
            </xdr:nvGraphicFramePr><a:graphic><a:graphicData>\
            <xdr:oleObject progId=\"Excel.Sheet.12\" shapeId=\"1027\" dvAspect=\"DVASPECT_ICON\" \
            autoLoad=\"1\" r:id=\"rId7\" r:link=\"rId6\"/></a:graphicData></a:graphic></xdr:graphicFrame>";
    let body = format!(
        "{}{}{}",
        two_cell_anchor(picture),
        two_cell_anchor(chart),
        two_cell_anchor(ole)
    );
    let objects = read(&drawing(&body)).unwrap().unwrap();
    assert_eq!(objects.len(), 1);
    let Object::OleObject(ole_object) = &objects[0].object else {
        panic!("expected an OLE object");
    };
    assert_eq!(ole_object.non_visual.id, Some(52));
    assert_eq!(ole_object.non_visual.name.as_deref(), Some("Object"));
    assert_eq!(ole_object.program_id.as_deref(), Some("Excel.Sheet.12"));
    assert_eq!(ole_object.shape_id, Some(1027));
    assert_eq!(ole_object.data_or_view_aspect, Some(Aspect::Icon));
    assert_eq!(ole_object.auto_load, Some(true));
    assert_eq!(ole_object.relationship_id.as_deref(), Some("rId7"));
    assert_eq!(ole_object.link_relationship_id.as_deref(), Some("rId6"));
}

#[test]
fn rejects_unknown_ole_object_aspect() {
    let ole = "<xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id=\"52\" name=\"Object\"/>\
            </xdr:nvGraphicFramePr><a:graphic><a:graphicData>\
            <xdr:oleObject shapeId=\"1027\" dvAspect=\"DVASPECT_THUMBNAIL\" \
            r:id=\"rId7\"/></a:graphicData></a:graphic></xdr:graphicFrame>";
    let xml = drawing(&two_cell_anchor(ole));

    let error = read(&xml).unwrap_err();

    assert!(error.to_string().contains("DVASPECT_THUMBNAIL"));
}

#[test]
fn parses_strict_namespace_dialect() {
    let strict_marker = |col: u32, row: u32| {
        format!(
            "<s:col>{col}</s:col><s:colOff>0</s:colOff>\
                 <s:row>{row}</s:row><s:rowOff>0</s:rowOff>"
        )
    };
    let xml = format!(
        "<s:wsDr xmlns:s=\"{STRICT_XDR}\" xmlns:a=\"{STRICT_A}\"><s:twoCellAnchor>\
             <s:from>{}</s:from><s:to>{}</s:to>\
             <s:sp><s:nvSpPr><s:cNvPr id=\"1\" name=\"Strict\"/><s:cNvSpPr txBox=\"1\"/></s:nvSpPr>\
             <s:spPr><a:prstGeom prst=\"rect\"/></s:spPr>\
             <s:txBody><a:bodyPr/><a:p><a:r><a:t>S</a:t></a:r></a:p></s:txBody></s:sp>\
             <s:clientData/></s:twoCellAnchor></s:wsDr>",
        strict_marker(0, 0),
        strict_marker(1, 1)
    );
    let objects = read(&xml).unwrap().unwrap();
    let Object::Shape(shape) = &objects[0].object else {
        panic!("expected a shape");
    };
    assert!(shape.is_text_box);
    assert_eq!(shape.preset(), Some(Preset::Rect));
    assert_eq!(shape.text_body.as_ref().unwrap().text(), "S");
    // ECMA-376 default body properties apply when a:bodyPr is empty.
    let properties = &shape.text_body.as_ref().unwrap().properties;
    assert_eq!(
        properties.insets.left.as_emu(),
        Insets::default().left.as_emu()
    );
    assert_eq!(properties.column_count, Columns::ONE);
}

#[test]
fn tolerates_empty_drawing_and_non_drawing_root() {
    let empty = drawing("");
    assert_eq!(read(&empty).unwrap().unwrap(), Vec::<AnchoredObject>::new());
    let empty_root = format!("<xdr:wsDr xmlns:xdr=\"{XDR}\"/>");
    assert!(read(&empty_root).unwrap().unwrap().is_empty());
    assert!(read("<other/>").unwrap().is_none());
}

#[test]
fn rejects_malformed_drawings() {
    let anchored_shape = two_cell_anchor(text_box_shape());
    let cases = [
            // DTD is rejected.
            format!("<!DOCTYPE xdr:wsDr>{}", drawing(&anchored_shape)),
            // Processing instructions are rejected.
            drawing(&format!("<?xml-stylesheet href=\"x\"?>{anchored_shape}")),
            // Multiple roots.
            format!("{}{}", drawing(""), drawing("")),
            // Unterminated root.
            "<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\">".to_string(),
            // Shape anchor without markers.
            drawing(
                "<xdr:twoCellAnchor><xdr:sp><xdr:nvSpPr><xdr:cNvPr id=\"1\"/>\
                 <xdr:cNvSpPr/></xdr:nvSpPr></xdr:sp><xdr:clientData/></xdr:twoCellAnchor>",
            ),
            // Invalid marker value.
            drawing(
                "<xdr:twoCellAnchor><xdr:from><xdr:col>x</xdr:col></xdr:from>\
                 <xdr:to/><xdr:sp><xdr:nvSpPr><xdr:cNvPr id=\"1\"/><xdr:cNvSpPr/></xdr:nvSpPr>\
                 </xdr:sp><xdr:clientData/></xdr:twoCellAnchor>",
            ),
            // Marker outside worksheet bounds.
            drawing(
                "<xdr:twoCellAnchor><xdr:from><xdr:col>16384</xdr:col><xdr:colOff>0</xdr:colOff>\
                 <xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>\
                 <xdr:to><xdr:col>16385</xdr:col><xdr:colOff>0</xdr:colOff>\
                 <xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>\
                 <xdr:sp><xdr:nvSpPr><xdr:cNvPr id=\"1\"/><xdr:cNvSpPr/></xdr:nvSpPr></xdr:sp>\
                 <xdr:clientData/></xdr:twoCellAnchor>",
            ),
            // Two objects in one anchor.
            drawing(&format!(
                "<xdr:twoCellAnchor><xdr:from>{}</xdr:from><xdr:to>{}</xdr:to>\
                 <xdr:sp><xdr:nvSpPr><xdr:cNvPr id=\"1\"/><xdr:cNvSpPr/></xdr:nvSpPr></xdr:sp>\
                 <xdr:sp><xdr:nvSpPr><xdr:cNvPr id=\"2\"/><xdr:cNvSpPr/></xdr:nvSpPr></xdr:sp>\
                 <xdr:clientData/></xdr:twoCellAnchor>",
                marker(0, 0, 0, 0),
                marker(1, 0, 1, 0)
            )),
            // One-cell anchor without extent.
            drawing(&format!(
                "<xdr:oneCellAnchor><xdr:from>{}</xdr:from>\
                 <xdr:sp><xdr:nvSpPr><xdr:cNvPr id=\"1\"/><xdr:cNvSpPr/></xdr:nvSpPr></xdr:sp>\
                 <xdr:clientData/></xdr:oneCellAnchor>",
                marker(0, 0, 0, 0)
            )),
        ];
    for xml in cases {
        assert!(read(&xml).is_err(), "accepted {xml}");
    }
}
