use super::*;

const DOCUMENT: &str = concat!(
    r#"<?xml version="1.0"?><office:document-content "#,
    r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
    r#"xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" "#,
    r#"xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" "#,
    r#"xmlns:xlink="http://www.w3.org/1999/xlink" "#,
    r#"office:version="1.3"><office:body><office:text>"#,
    r#"<draw:frame draw:name="Map">"#,
    r#"<draw:image-map>"#,
    r#"<draw:area-rectangle svg:x="1cm" svg:y="2cm" svg:width="3cm" svg:height="4cm" xlink:href="https://example.org/a" xlink:type="simple" office:target-frame-name="_blank" office:name="r1"><svg:title>Area A</svg:title></draw:area-rectangle>"#,
    r#"<draw:area-circle svg:cx="5cm" svg:cy="6cm" svg:r="7cm" draw:nohref="true"/>"#,
    r#"<draw:area-polygon svg:x="0cm" svg:y="0cm" svg:width="9cm" svg:height="8cm" svg:viewBox="0 0 100 100" svg:points="10,10 90,10 50,90" xlink:show="new" xlink:href="https://example.org/c"/>"#,
    r#"</draw:image-map>"#,
    r#"</draw:frame></office:text></office:body></office:document-content>"#,
);

#[test]
fn parses_all_area_kinds_with_inert_links() {
    let maps = parse_image_maps(DOCUMENT).unwrap();
    assert_eq!(maps.len(), 1);
    let map = &maps[0];
    assert_eq!(map.areas.len(), 3);
    assert!(map.xml.starts_with("<draw:image-map>"));

    let ImageMapAreaShape::Rectangle { x, width, .. } = &map.areas[0].shape else {
        panic!()
    };
    assert_eq!(x, "1cm");
    assert_eq!(width, "3cm");
    assert_eq!(map.areas[0].href.as_deref(), Some("https://example.org/a"));
    assert_eq!(map.areas[0].target_frame_name.as_deref(), Some("_blank"));
    assert_eq!(map.areas[0].name.as_deref(), Some("r1"));
    assert!(!map.areas[0].no_href);
    assert_eq!(
        map.areas[0].title_xml.as_deref(),
        Some("<svg:title>Area A</svg:title>")
    );

    let ImageMapAreaShape::Circle { cx, r, .. } = &map.areas[1].shape else {
        panic!()
    };
    assert_eq!((cx.as_str(), r.as_str()), ("5cm", "7cm"));
    assert!(map.areas[1].no_href);
    assert!(map.areas[1].href.is_none());

    let ImageMapAreaShape::Polygon {
        view_box, points, ..
    } = &map.areas[2].shape
    else {
        panic!()
    };
    assert_eq!(view_box, "0 0 100 100");
    assert_eq!(points, "10,10 90,10 50,90");
    assert_eq!(map.areas[2].show.as_deref(), Some("new"));
}

#[test]
fn reports_no_maps_when_absent() {
    let xml = concat!(
        r#"<?xml version="1.0"?><office:document-content "#,
        r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
        r#"office:version="1.3"><office:body/></office:document-content>"#,
    );
    assert!(parse_image_maps(xml).unwrap().is_empty());
}

#[test]
fn rejects_malformed_maps() {
    // Missing required circle radius.
    let bad = DOCUMENT.replace(" svg:r=\"7cm\"", "");
    assert!(parse_image_maps(&bad).is_err());
    // Invalid xlink:show.
    let bad = DOCUMENT.replace("xlink:show=\"new\"", "xlink:show=\"embed\"");
    assert!(parse_image_maps(&bad).is_err());
    // Invalid xlink:type.
    let bad = DOCUMENT.replace("xlink:type=\"simple\"", "xlink:type=\"extended\"");
    assert!(parse_image_maps(&bad).is_err());
    // Invalid Boolean.
    let bad = DOCUMENT.replace("draw:nohref=\"true\"", "draw:nohref=\"maybe\"");
    assert!(parse_image_maps(&bad).is_err());
    // Unexpected child element.
    let bad = DOCUMENT.replace(
        "<svg:title>Area A</svg:title>",
        "<draw:glue-point draw:id=\"0\" svg:x=\"1cm\" svg:y=\"1cm\"/>",
    );
    assert!(parse_image_maps(&bad).is_err());
    // Nested image map.
    let bad = DOCUMENT.replace("<draw:area-circle", "<draw:image-map><draw:area-circle");
    assert!(parse_image_maps(&bad).is_err());
}
