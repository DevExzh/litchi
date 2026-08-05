//! Regression tests for the layered drawing inventory owner.

use super::{Anchor, Kind, Object, parse};
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
    assert!(!drawing.is_text_box());
    assert!(drawing.is_inline());
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
