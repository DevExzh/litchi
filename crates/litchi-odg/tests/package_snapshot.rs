#![allow(
    clippy::unwrap_used,
    reason = "tests are expected to panic on unexpected fixture failures"
)]

use litchi_odf_common::core::PackageWriter;
use litchi_odg::Drawing;

const CONTENT: &str =
    include_str!("../../../test-data/odf/odg/drawing-style-resources-content.xml");
const STYLES: &str = include_str!("../../../test-data/odf/odg/drawing-style-resources-styles.xml");

#[test]
fn real_drawing_resource_xml_remains_exact_and_opaque() {
    let mut writer = PackageWriter::new();
    writer
        .set_mimetype("application/vnd.oasis.opendocument.graphics")
        .unwrap();
    writer.add_file("content.xml", CONTENT.as_bytes()).unwrap();
    writer.add_file("styles.xml", STYLES.as_bytes()).unwrap();
    let bytes = writer.finish_to_bytes().unwrap();

    let drawing = Drawing::from_bytes(bytes.clone()).unwrap();
    assert_eq!(drawing.as_bytes(), bytes.as_slice());
    assert_eq!(drawing.content_xml(), CONTENT);
    assert_eq!(drawing.styles_xml(), Some(STYLES));
    assert!(CONTENT.contains('\n') || STYLES.contains('\n'));

    let raw = format!("{}{}", drawing.content_xml(), drawing.styles_xml().unwrap());
    for element in [
        "draw:fill-image",
        "draw:gradient",
        "draw:hatch",
        "draw:marker",
        "draw:opacity",
        "draw:stroke-dash",
    ] {
        assert!(raw.contains(element), "fixture lost {element}");
    }
}
