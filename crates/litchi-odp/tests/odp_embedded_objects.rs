#![allow(
    clippy::unwrap_used,
    reason = "integration-test assertions panic on failure by design"
)]

use litchi_odp::Presentation;
use litchi_odp::core::PackageWriter;
use litchi_odp::embedded::{Kind, Source};

const CONTENT: &str = r#"<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:xlink="http://www.w3.org/1999/xlink"><office:body><office:presentation><draw:page draw:name="Slide 1"><draw:frame draw:name="Applet Frame"><draw:applet draw:code="org.example.Safe" draw:may-script="false"><draw:param draw:name="mode" draw:value="inert"/></draw:applet></draw:frame><draw:frame draw:name="Floating Frame"><draw:floating-frame draw:frame-name="Viewer" xlink:href="https://example.invalid/" xlink:type="simple"/></draw:frame></draw:page></office:presentation></office:body></office:document-content>"#;

#[test]
fn applets_and_floating_frames_are_bounded_inert_objects() {
    let mut writer = PackageWriter::new();
    writer
        .set_mimetype("application/vnd.oasis.opendocument.presentation")
        .unwrap();
    writer.add_file("content.xml", CONTENT.as_bytes()).unwrap();
    let presentation = Presentation::from_bytes(writer.finish_to_bytes().unwrap()).unwrap();

    let objects = presentation.embedded_objects().unwrap();
    assert_eq!(objects.len(), 2);
    assert_eq!(objects[0].kind, Kind::Applet);
    assert_eq!(objects[0].code.as_deref(), Some("org.example.Safe"));
    assert_eq!(objects[0].may_script, Some(false));
    assert_eq!(objects[0].parameters.len(), 1);
    assert_eq!(objects[0].parameters[0].name, "mode");
    assert_eq!(objects[0].parameters[0].value, "inert");

    assert_eq!(objects[1].kind, Kind::FloatingFrame);
    assert_eq!(objects[1].frame_name.as_deref(), Some("Viewer"));
    assert!(matches!(
        &objects[1].source,
        Source::Linked { href } if href == "https://example.invalid/"
    ));
}
