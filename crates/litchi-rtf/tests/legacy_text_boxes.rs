use litchi_rtf::{
    LegacyHorizontalAnchor, LegacyTextDirection, LegacyVerticalAnchor, RtfDocument, RtfWriter,
};

const SYNTHETIC: &str = concat!(
    r#"{\rtf1\ansi\uc1 A"#,
    r#"{\*\do\dobxcolumn\dobypara\dodhgt7\dptxbx\dptxbxmar10"#,
    r#"\dptxtbrl\dpx1\dpy2\dpxsize300\dpysize200"#,
    r#"{\dptxbxtext Box\u20320?\tab X\par Y}}B}"#,
);

#[test]
fn parses_typed_legacy_text_box_and_round_trips_canonically() {
    let document = RtfDocument::parse(SYNTHETIC).unwrap();
    assert_eq!(document.text(), "AB");
    assert_eq!(document.legacy_text_boxes().len(), 1);
    let text_box = &document.legacy_text_boxes()[0];
    assert_eq!(text_box.text, "Box你\tX\nY");
    assert_eq!(text_box.position, 1);
    assert_eq!(
        text_box.horizontal_anchor,
        Some(LegacyHorizontalAnchor::Column)
    );
    assert_eq!(
        text_box.vertical_anchor,
        Some(LegacyVerticalAnchor::Paragraph)
    );
    assert_eq!(text_box.z_order, Some(7));
    assert_eq!(text_box.margin, Some(10));
    assert_eq!((text_box.x, text_box.y), (Some(1), Some(2)));
    assert_eq!((text_box.width, text_box.height), (Some(300), Some(200)));
    assert_eq!(
        text_box.direction,
        LegacyTextDirection::TopToBottomRightToLeft
    );

    let mut first = Vec::new();
    RtfWriter::new(&mut first)
        .write_document(&document)
        .unwrap();
    let reparsed = RtfDocument::parse_bytes(&first).unwrap();
    assert_eq!(reparsed.text(), document.text());
    assert_eq!(reparsed.legacy_text_boxes(), document.legacy_text_boxes());
    let mut second = Vec::new();
    RtfWriter::new(&mut second)
        .write_document(&reparsed)
        .unwrap();
    assert_eq!(first, second);
}

#[test]
fn rejects_misplaced_malformed_active_and_binary_legacy_text_boxes() {
    let malformed = [
        r#"{\rtf1{\do\dptxbx{\dptxbxtext x}}}"#,
        r#"{\rtf1{\dptxbxtext x}}"#,
        r#"{\rtf1{\*\dptxbxtext x}}"#,
        r#"{\rtf1{\header{\*\do\dptxbx{\dptxbxtext x}}}Body}"#,
        r#"{\rtf1{\*\do\dptxbx}}"#,
        r#"{\rtf1{\*\do{\dptxbxtext x}\dptxbx}}"#,
        r#"{\rtf1{\*\do\dptxbx\dptxbx{\dptxbxtext x}}}"#,
        r#"{\rtf1{\*\do\dptxbx{\dptxbxtext x}{\dptxbxtext y}}}"#,
        r#"{\rtf1{\*\do\dobxpage\dobxmargin\dptxbx{\dptxbxtext x}}}"#,
        r#"{\rtf1{\*\do\dptxbx\dpx1\dpx2{\dptxbxtext x}}}"#,
        r#"{\rtf1{\*\do\dptxbx\dptxtbrl\dptxbtlr{\dptxbxtext x}}}"#,
        r#"{\rtf1{\*\do\dptxbx\dpxsize0{\dptxbxtext x}}}"#,
        r#"{\rtf1{\*\do\dptxbx\dpysize-1{\dptxbxtext x}}}"#,
        r#"{\rtf1{\*\do\dptxbx\dptxbxmar-1{\dptxbxtext x}}}"#,
        r#"{\rtf1{\*\do\dptxbx{\dptxbxtext}}}"#,
        r#"{\rtf1{\*\do\dptxbx{\dptxbxtext{\field danger}}}}"#,
        r#"{\rtf1{\*\do\dptxbx{\dptxbxtext{\object danger}}}}"#,
        r#"{\rtf1{\*\do\dptxbx{\dptxbxtext{\pict 00}}}}"#,
        r#"{\rtf1{\*\do\dptxbx{\dptxbxtext{\shp danger}}}}"#,
        r#"{\rtf1{\*\do\dptxbx{\dptxbxtext{\formfield danger}}}}"#,
        r#"{\rtf1{\*\do\dptxbx{\dptxbxtext\bin2 xx}}}"#,
        r#"{\rtf1{\*\do\dptxbx\bin2 xx{\dptxbxtext x}}}"#,
    ];
    for source in malformed {
        assert!(
            RtfDocument::parse(source).is_err(),
            "accepted malformed legacy text box: {source}"
        );
    }

    let oversized = format!(
        r#"{{\rtf1{{\*\do\dptxbx{{\dptxbxtext {}}}}}}}"#,
        "x".repeat(1_048_577)
    );
    assert!(RtfDocument::parse(&oversized).is_err());
}

#[test]
fn ignores_unrelated_inert_legacy_drawings() {
    let document = RtfDocument::parse(r#"{\rtf1 A{\*\do\dpline\dpx1}B}"#).unwrap();
    assert_eq!(document.text(), "AB");
    assert!(document.legacy_text_boxes().is_empty());
}

fn isolated_drawing(fixture: &[u8]) -> Vec<u8> {
    let marker = br"{\*\do";
    let start = fixture
        .windows(marker.len())
        .position(|window| window == marker)
        .unwrap();
    let mut depth = 0usize;
    let mut end = None;
    for (offset, byte) in fixture[start..].iter().enumerate() {
        match byte {
            b'{' => depth += 1,
            b'}' => {
                depth -= 1;
                if depth == 0 {
                    end = Some(start + offset + 1);
                    break;
                }
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
fn parses_named_libreoffice_legacy_text_box_fixtures() {
    let hollow = include_bytes!(
        "../../../3rdparty/libreoffice-core/sw/qa/extras/rtfexport/data/dplinehollow.rtf"
    );
    let document = RtfDocument::parse_bytes(&isolated_drawing(hollow)).unwrap();
    assert_eq!(document.text(), "AB");
    let text_box = &document.legacy_text_boxes()[0];
    assert_eq!(text_box.text, "textbox without border\n");
    assert_eq!(text_box.horizontal_anchor, Some(LegacyHorizontalAnchor::Page));
    assert_eq!(
        text_box.vertical_anchor,
        Some(LegacyVerticalAnchor::Paragraph)
    );
    assert_eq!(text_box.z_order, Some(8192));
    assert_eq!(text_box.margin, Some(0));
    assert_eq!((text_box.x, text_box.y), (Some(929), Some(340)));
    assert_eq!((text_box.width, text_box.height), (Some(10556), Some(561)));

    let relation = include_bytes!(
        "../../../3rdparty/libreoffice-core/sw/qa/extras/rtfexport/data/dptxbx-relation.rtf"
    );
    let document = RtfDocument::parse_bytes(&isolated_drawing(relation)).unwrap();
    let text_box = &document.legacy_text_boxes()[0];
    assert_eq!(text_box.text, "To:\n");
    assert_eq!((text_box.x, text_box.y), (Some(941), Some(2114)));
    assert_eq!((text_box.width, text_box.height), (Some(1349), Some(221)));

    let nested_geometry = include_bytes!(
        "../../../3rdparty/libreoffice-core/sw/qa/core/data/rtf/pass/fdo78900.rtf"
    );
    let document = RtfDocument::parse_bytes(&isolated_drawing(nested_geometry)).unwrap();
    let text_box = &document.legacy_text_boxes()[0];
    assert!(text_box.text.contains("hello"));
    assert_eq!((text_box.x, text_box.y), (Some(227), Some(227)));
    assert_eq!(
        (text_box.width, text_box.height),
        (Some(11911), Some(9709))
    );
}
