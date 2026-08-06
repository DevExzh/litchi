//! Regression tests for the layered drawing inventory owner.

use super::{Anchor, AnchorId, Kind, LegacyAnchorKind, Object, parse, parse_legacy};
use litchi_drawingml::geom::Preset;

#[test]
fn rejects_unknown_or_non_schema_shape_presets() {
    assert!("customShape".parse::<Preset>().is_err());
    assert!("textBox".parse::<Preset>().is_err());

    let xml = br#"<w:drawing><wp:inline><wps:wsp><wps:spPr>
        <a:prstGeom prst="customShape"/>
        </wps:spPr></wps:wsp></wp:inline></w:drawing>"#;
    let error = parse(xml).unwrap_err();
    assert!(error.to_string().contains("customShape"));
}

#[test]
fn rejects_invalid_utf8_in_preset_tokens() {
    let xml = b"<w:drawing><wp:inline><wps:wsp><wps:spPr>\
        <a:prstGeom prst=\"\xff\"/>\
        </wps:spPr></wps:wsp></wp:inline></w:drawing>";
    assert!(parse(xml).is_err());
}

#[test]
fn object_dimensions_and_context_are_ergonomic() {
    let object = Object::new(
        "Shape 1".to_string(),
        "Test shape".to_string(),
        914400,
        1828800,
        Preset::Rect,
    );

    assert_eq!(object.width_emu(), 914400);
    assert_eq!(object.height_emu(), 1828800);
    assert_eq!(object.width_px(), 96);
    assert_eq!(object.height_px(), 192);
    assert!((object.width_pt() - 72.0).abs() < 0.1);
    assert!((object.height_pt() - 144.0).abs() < 0.1);
    assert_eq!(object.kind(), Kind::Shape);
    assert_eq!(object.anchor(), Anchor::Inline);
    assert_eq!(object.anchor_id(), None);
}

#[test]
fn empty_paragraph_has_no_drawings() {
    let drawings = parse(b"<w:p><w:r><w:t>Text only</w:t></w:r></w:p>").unwrap();
    assert!(drawings.is_empty());
}

#[test]
fn missing_geometry_is_not_reported_as_a_rectangle() {
    let xml = br#"<w:drawing><wp:inline><wps:wsp><wps:spPr/></wps:wsp></wp:inline></w:drawing>"#;
    let drawings = parse(xml).unwrap();
    assert_eq!(drawings.len(), 1);
    assert_eq!(drawings[0].preset(), None);
    assert_eq!(drawings[0].kind(), Kind::Shape);
}

#[test]
fn defaults_and_document_order_are_preserved() {
    let xml = br#"<w:p>
        <w:r><w:drawing><wp:inline><wps:wsp><wp:docPr name="first"/></wps:wsp></wp:inline></w:drawing></w:r>
        <w:r><w:drawing><wp:anchor><wp:docPr name="second"/><x:opaque/></wp:anchor></w:drawing></w:r>
    </w:p>"#;
    let drawings = parse(xml).unwrap();
    assert_eq!(drawings.len(), 2);
    assert_eq!(drawings[0].name(), "first");
    assert_eq!(drawings[1].name(), "second");
    assert_eq!(drawings[0].width_emu(), 914400);
    assert_eq!(drawings[0].height_emu(), 914400);
    assert_eq!(drawings[1].anchor(), Anchor::Floating);
    assert_eq!(drawings[1].kind(), Kind::Other);
}

#[test]
fn parses_shape_inventory() {
    let xml = br#"<w:p>
        <w:r>
            <w:drawing>
                <wp:inline>
                    <wp:extent cx="1000000" cy="2000000"/>
                    <wp:docPr id="1" name="MyShape" descr="Test shape"/>
                    <a:graphic>
                        <a:graphicData>
                            <wps:wsp>
                                <wps:spPr>
                                    <a:prstGeom prst="rect"/>
                                </wps:spPr>
                            </wps:wsp>
                        </a:graphicData>
                    </a:graphic>
                </wp:inline>
            </w:drawing>
        </w:r>
    </w:p>"#;

    let drawings = parse(xml).unwrap();
    assert_eq!(drawings.len(), 1);

    let drawing = &drawings[0];
    assert_eq!(drawing.name(), "MyShape");
    assert_eq!(drawing.description(), "Test shape");
    assert_eq!(drawing.width_emu(), 1000000);
    assert_eq!(drawing.height_emu(), 2000000);
    assert_eq!(drawing.preset(), Some(Preset::Rect));
    assert_eq!(drawing.kind(), Kind::Shape);
    assert_eq!(drawing.anchor_id(), None);
    assert!(!drawing.is_text_box());
    assert!(drawing.is_inline());
}

#[test]
fn parses_checked_inline_and_floating_anchor_ids() {
    let xml = br#"<w:p>
        <w:r><w:drawing><wp:inline wp14:anchorId="00000001"/>
        </w:drawing></w:r>
        <w:r><w:drawing><wp:anchor wp14:anchorId="7fffffff"/></w:drawing></w:r>
    </w:p>"#;
    let drawings = parse(xml).unwrap();
    assert_eq!(drawings.len(), 2);
    assert_eq!(drawings[0].anchor_id(), AnchorId::new(1));
    assert_eq!(drawings[1].anchor_id(), AnchorId::new(0x7fff_ffff));
}

#[test]
fn rejects_invalid_anchor_ids() {
    for value in ["00000000", "80000000", "0000001", "0000000G"] {
        let xml = format!(r#"<w:drawing><wp:inline wp14:anchorId="{value}"/></w:drawing>"#);
        assert!(
            parse(xml.as_bytes()).is_err(),
            "accepted invalid anchorId {value}"
        );
    }
}

#[test]
fn parses_legacy_object_and_picture_anchor_ids_in_document_order() {
    let xml = br#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
        xmlns:v="urn:schemas-microsoft-com:vml">
        <w:r><w:object w14:anchorId="0000002a"><v:shape/></w:object></w:r>
        <w:r><w:pict w14:anchorId="7fffffff"><v:shape/></w:pict></w:r>
    </w:p>"#;

    let anchors = parse_legacy(xml).unwrap();
    assert_eq!(anchors.len(), 2);
    assert_eq!(anchors[0].kind(), LegacyAnchorKind::Object);
    assert_eq!(anchors[0].anchor_id(), AnchorId::new(0x2a));
    assert_eq!(anchors[1].kind(), LegacyAnchorKind::Picture);
    assert_eq!(anchors[1].anchor_id(), AnchorId::new(0x7fff_ffff));
}

#[test]
fn legacy_anchor_parser_is_namespace_and_range_strict() {
    let invalid_values = ["00000000", "80000000", "0000001", "0000000G"];
    for value in invalid_values {
        let xml = format!(
            r#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:pict w14:anchorId="{value}"/></w:p>"#
        );
        assert!(parse_legacy(xml.as_bytes()).is_err(), "accepted {value}");
    }

    let wrong_namespace = br#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:not-word-2010"><w:object x:anchorId="00000001"/></w:p>"#;
    assert!(parse_legacy(wrong_namespace).is_err());

    let unqualified = br#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:pict anchorId="00000001"/></w:p>"#;
    assert!(parse_legacy(unqualified).is_err());

    let duplicate_namespace = br#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:a="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:b="http://schemas.microsoft.com/office/word/2010/wordml"><w:object a:anchorId="00000001" b:anchorId="00000002"/></w:p>"#;
    assert!(parse_legacy(duplicate_namespace).is_err());
}

#[test]
fn paragraph_legacy_anchor_api_is_typed_and_inert() {
    let paragraph = crate::paragraph::Paragraph::new(
        br#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:foreign="urn:foreign"><w:r><w:pict w14:anchorId="00000009"><foreign:payload/></w:pict></w:r></w:p>"#.to_vec(),
    );
    let anchors = paragraph.legacy_anchors().unwrap();
    assert_eq!(
        anchors.as_slice(),
        &[super::LegacyAnchor::from_parts(
            LegacyAnchorKind::Picture,
            AnchorId::new(9),
        )]
    );
}

#[test]
fn parses_nested_textbox_text_in_source_order() {
    let xml = br#"<w:p>
        <w:r>
            <w:drawing>
                <wp:anchor>
                    <wp:extent cx="1000000" cy="1000000"/>
                    <wp:docPr name="TextBox1"/>
                    <wps:wsp>
                        <wps:spPr>
                            <a:prstGeom prst="rect"/>
                        </wps:spPr>
                        <wps:txbx>
                            <w:txbxContent>
                                <w:p><w:r><w:t>Hello World</w:t></w:r></w:p>
                            </w:txbxContent>
                        </wps:txbx>
                    </wps:wsp>
                </wp:anchor>
            </w:drawing>
        </w:r>
    </w:p>"#;

    let drawings = parse(xml).unwrap();
    assert_eq!(drawings.len(), 1);

    let drawing = &drawings[0];
    assert_eq!(drawing.name(), "TextBox1");
    assert_eq!(drawing.preset(), Some(Preset::Rect));
    assert_eq!(drawing.kind(), Kind::TextBox);
    assert!(drawing.is_text_box());
    assert_eq!(drawing.text(), "Hello World");
    assert!(!drawing.is_inline());
}
