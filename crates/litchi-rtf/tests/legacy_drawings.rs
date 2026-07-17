use litchi_rtf::{
    LegacyCalloutType, LegacyDrawingFillPattern, LegacyDrawingLineStyle,
    LegacyDrawingPrimitive, RtfDocument, RtfWriter,
};

fn isolated_drawing(fixture: &[u8]) -> Vec<u8> {
    let marker = br"{\*\do";
    let start = fixture.windows(marker.len()).position(|window| window == marker).unwrap();
    let mut depth = 0usize;
    let mut end = None;
    for (offset, byte) in fixture[start..].iter().enumerate() {
        match byte {
            b'{' => depth += 1,
            b'}' => {
                depth -= 1;
                if depth == 0 { end = Some(start + offset + 1); break; }
            },
            _ => {},
        }
    }
    let mut source = br"{\rtf1\ansi\ansicpg1252\uc1 A".to_vec();
    source.extend_from_slice(&fixture[start..end.unwrap()]);
    source.extend_from_slice(b"B}");
    source
}

#[test]
fn parses_all_simple_primitives_and_round_trips_canonically() {
    let source = concat!(
        r#"{\rtf1 A"#,
        r#"{\*\do\dobxpage\dobypara\dodhgt1\dolock\dpline\dpptx1\dppty2\dpptx3\dppty4\dpx5\dpy6\dpxsize7\dpysize8\dplinedash\dplinecor9\dplinecog10\dplinecob11\dplinew12\dpastartsol\dpastartl1\dpastartw2\dpaendhol\dpaendl3\dpaendw2\dpshadow\dpshadx13\dpshady14}"#,
        r#"{\*\do\dobxmargin\dobymargin\dodhgt2\dprect\dproundr\dpx1\dpy2\dpxsize3\dpysize4\dpfillfggray100\dpfillbgcr5\dpfillbgcg6\dpfillbgcb7\dpfillpat25}"#,
        r#"{\*\do\dobxcolumn\dobypage\dodhgt3\dpellipse\dpx1\dpy2\dpxsize3\dpysize4}"#,
        r#"{\*\do\dobxpage\dobypara\dodhgt4\dparc\dparcflipx\dparcflipy\dpx1\dpy2\dpxsize3\dpysize4}"#,
        r#"{\*\do\dobxpage\dobypara\dodhgt5\dppolyline\dppolygon\dppolycount3\dpptx0\dppty0\dpptx10\dppty0\dpptx10\dppty10\dpx1\dpy2\dpxsize3\dpysize4}"#,
        "B}",
    );
    let document = RtfDocument::parse(source).unwrap();
    assert_eq!(document.text(), "AB");
    assert_eq!(document.legacy_drawings().len(), 5);
    let LegacyDrawingPrimitive::Line { properties, .. } = &document.legacy_drawings()[0].primitive else { panic!() };
    assert_eq!(properties.line.unwrap().style, LegacyDrawingLineStyle::Dashed);
    assert!(properties.start_arrow.is_some());
    assert!(properties.end_arrow.is_some());
    assert!(properties.shadow.is_some());
    let LegacyDrawingPrimitive::Rectangle { rounded, properties, .. } = &document.legacy_drawings()[1].primitive else { panic!() };
    assert!(*rounded);
    assert_eq!(properties.fill.unwrap().pattern, LegacyDrawingFillPattern::LightTrellis);
    let LegacyDrawingPrimitive::Polyline { closed, points, .. } = &document.legacy_drawings()[4].primitive else { panic!() };
    assert!(*closed);
    assert_eq!(points.len(), 3);

    let mut first = Vec::new();
    RtfWriter::new(&mut first).write_document(&document).unwrap();
    let reparsed = RtfDocument::parse_bytes(&first).unwrap();
    assert_eq!(reparsed.legacy_drawings(), document.legacy_drawings());
    let mut second = Vec::new();
    RtfWriter::new(&mut second).write_document(&reparsed).unwrap();
    assert_eq!(first, second);
}

#[test]
fn parses_recursive_groups_and_callouts() {
    let source = concat!(
        r#"{\rtf1"#,
        r#"{\*\do\dobxpage\dobypara\dodhgt1\dpgroup\dpcount2\dpx0\dpy0\dpxsize100\dpysize100"#,
        r#"\dprect\dpx0\dpy0\dpxsize10\dpysize10"#,
        r#"\dpgroup\dpcount2\dpx10\dpy10\dpxsize20\dpysize20\dpellipse\dpx1\dpy1\dpxsize2\dpysize2\dpendgroup\dpx0\dpy0\dpxsize0\dpysize0"#,
        r#"\dpendgroup\dpx0\dpy0\dpxsize0\dpysize0}"#,
        r#"{\*\do\dobxmargin\dobymargin\dodhgt2\dpcallout\dpcotsingle\dpcoa45\dpcoaccent\dpcodcenter\dpcooffset5\dpcolength6\dpx0\dpy0\dpxsize100\dpysize100"#,
        r#"\dppolyline\dppolycount2\dpptx0\dppty0\dpptx10\dppty10\dpx0\dpy0\dpxsize10\dpysize10"#,
        r#"\dptxbx\dptxbxmar2{\dptxbxtext hello\par}\dpx10\dpy10\dpxsize30\dpysize20}"#,
        "}",
    );
    let document = RtfDocument::parse(source).unwrap();
    assert_eq!(document.legacy_drawings().len(), 2);
    let LegacyDrawingPrimitive::Group { children, .. } = &document.legacy_drawings()[0].primitive else { panic!() };
    assert_eq!(children.len(), 2);
    assert!(matches!(children[1], LegacyDrawingPrimitive::Group { .. }));
    let LegacyDrawingPrimitive::Callout(callout) = &document.legacy_drawings()[1].primitive else { panic!() };
    assert_eq!(callout.callout_type, LegacyCalloutType::Single);
    assert_eq!(callout.angle, Some(45));
    assert!(matches!(*callout.polyline, LegacyDrawingPrimitive::Polyline { .. }));
    assert!(matches!(*callout.text_box, LegacyDrawingPrimitive::TextBox { .. }));
}

#[test]
fn parses_named_libreoffice_primitive_fixtures() {
    let polyline = include_bytes!("../../../3rdparty/libreoffice-core/sw/qa/extras/rtfexport/data/dppolyline.rtf");
    let document = RtfDocument::parse_bytes(polyline).unwrap();
    assert_eq!(document.legacy_drawings().len(), 4);
    assert!(document.legacy_drawings().iter().all(|drawing| matches!(&drawing.primitive, LegacyDrawingPrimitive::Polyline { points, .. } if points.len() == 2)));

    let rectangle = include_bytes!("../../../3rdparty/libreoffice-core/sw/qa/extras/rtfexport/data/dprect-anchor.rtf");
    let document = RtfDocument::parse_bytes(rectangle).unwrap();
    let LegacyDrawingPrimitive::Rectangle { rounded, properties, .. } = &document.legacy_drawings()[0].primitive else { panic!() };
    assert!(*rounded);
    assert_eq!(properties.line.unwrap().style, LegacyDrawingLineStyle::Hollow);
    assert_eq!(properties.fill.unwrap().pattern, LegacyDrawingFillPattern::Solid);

    let group = include_bytes!("../../../3rdparty/libreoffice-core/sw/qa/extras/rtfimport/data/tdf91684.rtf");
    let document = RtfDocument::parse_bytes(&isolated_drawing(group)).unwrap();
    let LegacyDrawingPrimitive::Group { children, .. } = &document.legacy_drawings()[0].primitive else { panic!() };
    assert_eq!(children.len(), 3);
}

#[test]
fn rejects_malformed_order_duplicates_cardinality_and_caps() {
    let malformed = [
        r#"{\rtf1{\*\do\dobypara\dodhgt1\dprect\dpx0\dpy0\dpxsize1\dpysize1}}"#,
        r#"{\rtf1{\*\do\dobxpage\dobxmargin\dobypara\dodhgt1\dprect\dpx0\dpy0\dpxsize1\dpysize1}}"#,
        r#"{\rtf1{\*\do\dobxpage\dobypara\dodhgt1\dprect\dpy0\dpx0\dpxsize1\dpysize1}}"#,
        r#"{\rtf1{\*\do\dobxpage\dobypara\dodhgt1\dppolyline\dppolycount2\dpptx0\dppty0\dpx0\dpy0\dpxsize1\dpysize1}}"#,
        r#"{\rtf1{\*\do\dobxpage\dobypara\dodhgt1\dpgroup\dpcount9\dpx0\dpy0\dpxsize1\dpysize1\dprect\dpx0\dpy0\dpxsize1\dpysize1\dpendgroup\dpx0\dpy0\dpxsize0\dpysize0}}"#,
        r#"{\rtf1{\*\do\dobxpage\dobypara\dodhgt1\dprect\dpx0\dpy0\dpxsize1\dpysize1\dpfillfgcr256\dpfillfgcg0\dpfillfgcb0\dpfillbggray0\dpfillpat1}}"#,
        r#"{\rtf1{\*\do\dobxpage\dobypara\dodhgt1\dpline\dpptx0\dppty0\dpptx1\dppty1\dpx0\dpy0\dpxsize1\dpysize1\dpastartsol\dpastartl4\dpastartw1}}"#,
        r#"{\rtf1{\*\do\dobxpage\dobypara\dodhgt1\dprect\dpx0\dpy0\dpxsize1\dpysize1\dpshadow\dpshadx1}}"#,
        r#"{\rtf1{\*\do\dobxpage\dobypara\dodhgt1\dppolyline\dppolycount65537}}"#,
    ];
    for source in malformed {
        assert!(RtfDocument::parse(source).is_err(), "accepted malformed drawing: {source}");
    }
}
