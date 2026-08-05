//! Tests for `draw:image-map` through the package and flat-document APIs.

use litchi_odt::image_map::ImageMapAreaShape;
mod support;

const CONTENT: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8"?><office:document "#,
    r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
    r#"xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" "#,
    r#"xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" "#,
    r#"xmlns:xlink="http://www.w3.org/1999/xlink" "#,
    r#"office:version="1.3" office:mimetype="application/vnd.oasis.opendocument.text"><office:body><office:text>"#,
    r#"<draw:frame draw:name="Map" svg:width="10cm" svg:height="10cm">"#,
    r#"<draw:image xlink:href="Pictures/map.png" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>"#,
    r#"<draw:image-map>"#,
    r#"<draw:area-rectangle svg:x="1cm" svg:y="2cm" svg:width="3cm" svg:height="4cm" xlink:href="https://example.org/a" xlink:type="simple"/>"#,
    r#"<draw:area-polygon svg:x="0cm" svg:y="0cm" svg:width="9cm" svg:height="8cm" svg:viewBox="0 0 100 100" svg:points="10,10 90,10 50,90" draw:nohref="true"/>"#,
    r#"</draw:image-map>"#,
    r#"</draw:frame></office:text></office:body></office:document>"#,
);

const MIMETYPE: &str = "application/vnd.oasis.opendocument.text";

fn package() -> litchi_odt::generic::Package {
    litchi_odt::generic::Package::from_bytes(support::package(
        MIMETYPE,
        &[("content.xml", CONTENT.as_bytes())],
    ))
    .unwrap()
}

#[test]
fn reads_image_maps_from_a_package() {
    let maps = package().image_maps().unwrap();
    assert_eq!(maps.len(), 1);
    let map = &maps[0];
    assert_eq!(map.areas.len(), 2);

    let ImageMapAreaShape::Rectangle { x, height, .. } = &map.areas[0].shape else {
        panic!()
    };
    assert_eq!((x.as_str(), height.as_str()), ("1cm", "4cm"));
    assert_eq!(map.areas[0].href.as_deref(), Some("https://example.org/a"));
    assert!(map.areas[0].xml.contains("draw:area-rectangle"));

    let ImageMapAreaShape::Polygon { points, .. } = &map.areas[1].shape else {
        panic!()
    };
    assert_eq!(points, "10,10 90,10 50,90");
    assert!(map.areas[1].no_href);
    assert!(map.areas[1].href.is_none());
}

#[test]
fn reads_image_maps_from_a_flat_document() {
    let document =
        litchi_odt::generic::FlatDocument::from_bytes(CONTENT.as_bytes().to_vec()).unwrap();
    let maps = document.image_maps().unwrap();
    assert_eq!(maps.len(), 1);
    assert_eq!(maps[0].areas.len(), 2);
}

#[test]
fn packages_without_image_maps_report_empty() {
    let plain = CONTENT.replace(
        "<draw:image-map><draw:area-rectangle svg:x=\"1cm\" svg:y=\"2cm\" svg:width=\"3cm\" svg:height=\"4cm\" xlink:href=\"https://example.org/a\" xlink:type=\"simple\"/><draw:area-polygon svg:x=\"0cm\" svg:y=\"0cm\" svg:width=\"9cm\" svg:height=\"8cm\" svg:viewBox=\"0 0 100 100\" svg:points=\"10,10 90,10 50,90\" draw:nohref=\"true\"/></draw:image-map>",
        "",
    );
    let document =
        litchi_odt::generic::FlatDocument::from_bytes(plain.as_bytes().to_vec()).unwrap();
    assert!(document.image_maps().unwrap().is_empty());
}
