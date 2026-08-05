use super::{Object, UnknownKind, parse};

#[test]
fn parses_prefixed_strict_picture_and_chart_anchors_in_source_order() {
    let xml = r#"<s:wsDr xmlns:s="http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing"
            xmlns:a="http://purl.oclc.org/ooxml/drawingml/main"
            xmlns:c="http://purl.oclc.org/ooxml/drawingml/chart"
            xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships">
        <s:twoCellAnchor><s:from><s:col>1</s:col><s:colOff>2</s:colOff>
            <s:row>3</s:row><s:rowOff>4</s:rowOff></s:from>
            <s:to><s:col>5</s:col><s:colOff>6</s:colOff>
            <s:row>7</s:row><s:rowOff>8</s:rowOff></s:to>
            <s:pic><s:nvPicPr><s:cNvPr descr="Logo"/></s:nvPicPr>
                <s:blipFill><a:blip r:embed="image-rel"/></s:blipFill></s:pic>
            <s:clientData/></s:twoCellAnchor>
        <s:twoCellAnchor><s:from><s:col>0</s:col><s:colOff>0</s:colOff>
            <s:row>0</s:row><s:rowOff>0</s:rowOff></s:from>
            <s:to><s:col>9</s:col><s:colOff>0</s:colOff>
            <s:row>10</s:row><s:rowOff>0</s:rowOff></s:to>
            <s:graphicFrame><a:graphic><a:graphicData><c:chart r:id="chart-rel"/>
            </a:graphicData></a:graphic></s:graphicFrame><s:clientData/></s:twoCellAnchor>
    </s:wsDr>"#;

    let drawing = parse(xml).unwrap().unwrap();
    assert_eq!(drawing.len(), 2);
    assert_eq!(drawing.pictures().count(), 1);
    assert_eq!(drawing.charts().count(), 1);
    assert!(matches!(drawing.objects()[0], Object::Picture(_)));
    assert!(matches!(drawing.objects()[1], Object::Chart(_)));

    let picture = drawing.pictures().next().unwrap();
    assert_eq!(picture.relationship_id, "image-rel");
    assert_eq!(picture.description.as_deref(), Some("Logo"));
    assert_eq!(picture.anchor.from_col, 1);
    assert_eq!(picture.anchor.to_row_offset, 8);

    let chart = drawing.charts().next().unwrap();
    assert_eq!(chart.relationship_id, "chart-rel");
    assert_eq!(chart.anchor.to_col, 9);
}

#[test]
fn parses_default_namespace_transitional_drawing() {
    let xml = r#"<wsDr xmlns="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <twoCellAnchor><from><col>0</col><colOff>0</colOff><row>0</row><rowOff>0</rowOff></from>
            <to><col>2</col><colOff>0</colOff><row>3</row><rowOff>0</rowOff></to>
            <pic><nvPicPr><cNvPr descr="default"/></nvPicPr>
                <blipFill><a:blip r:embed="rIdImage"/></blipFill></pic>
            <clientData/></twoCellAnchor>
    </wsDr>"#;

    let drawing = parse(xml).unwrap().unwrap();
    assert_eq!(
        drawing.pictures().next().unwrap().relationship_id,
        "rIdImage"
    );
}

#[test]
fn retains_unknown_shape_as_inert_structure() {
    let xml = r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <xdr:twoCellAnchor><xdr:from><xdr:col>2</xdr:col><xdr:colOff>3</xdr:colOff>
            <xdr:row>4</xdr:row><xdr:rowOff>5</xdr:rowOff></xdr:from>
            <xdr:to><xdr:col>6</xdr:col><xdr:colOff>7</xdr:colOff>
            <xdr:row>8</xdr:row><xdr:rowOff>9</xdr:rowOff></xdr:to>
            <xdr:sp><xdr:nvSpPr><xdr:cNvPr descr="Text box"/></xdr:nvSpPr>
                <xdr:spPr/><xdr:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>inert</a:t></a:r></a:p></xdr:txBody>
            </xdr:sp><xdr:clientData/></xdr:twoCellAnchor>
    </xdr:wsDr>"#;

    let drawing = parse(xml).unwrap().unwrap();
    let unknown = drawing.unknown().next().unwrap();
    assert_eq!(unknown.kind, UnknownKind::Shape);
    assert_eq!(unknown.description.as_deref(), Some("Text box"));
    assert_eq!(unknown.anchor.from_col, 2);
    assert_eq!(unknown.anchor.to_row_offset, 9);
}

#[test]
fn parses_chartsheet_absolute_anchor_chart() {
    let xml = r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="8582025" cy="5838825"/>
            <xdr:graphicFrame macro=""><a:graphic><a:graphicData>
                <c:chart r:id="chart-rel"/>
            </a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:absoluteAnchor>
    </xdr:wsDr>"#;

    let drawing = parse(xml).unwrap().unwrap();
    assert_eq!(
        drawing.charts().next().unwrap().relationship_id,
        "chart-rel"
    );
    // Absolute anchors have no worksheet markers in this compact chart model.
    assert_eq!(drawing.charts().next().unwrap().anchor.to_col, 0);

    let empty = r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"><xdr:absoluteAnchor/></xdr:wsDr>"#;
    assert!(parse(empty).is_err());
}

#[test]
fn rejects_malformed_drawing_anchors_and_bounded_marker_text() {
    const ROOT: &str = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    for body in [
        "<xdr:twoCellAnchor/>",
        "<xdr:twoCellAnchor><xdr:from/></xdr:twoCellAnchor>",
        "<xdr:twoCellAnchor><xdr:from><xdr:col>x</xdr:col></xdr:from></xdr:twoCellAnchor>",
    ] {
        let xml = format!(r#"<xdr:wsDr xmlns:xdr="{ROOT}">{body}</xdr:wsDr>"#);
        assert!(parse(&xml).is_err(), "accepted {xml}");
    }

    let oversized = "x".repeat(65);
    let xml = format!(
        r#"<xdr:wsDr xmlns:xdr="{ROOT}"><xdr:twoCellAnchor><xdr:from><xdr:col>{oversized}</xdr:col></xdr:from></xdr:twoCellAnchor></xdr:wsDr>"#
    );
    assert!(parse(&xml).is_err());
}

#[test]
fn reuses_shared_drawingml_text_body_types() {
    let body = super::model::text::Body::default();
    assert!(body.text().is_empty());
}
