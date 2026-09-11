#![allow(clippy::unwrap_used, reason = "test assertions use unwrap for clarity")]

use litchi_odf_common::core::PackageWriter;
use litchi_odg::{Drawing, Transition, shape::ShapeKind};

fn package(content: &str) -> Vec<u8> {
    let mut writer = PackageWriter::new();
    writer
        .set_mimetype("application/vnd.oasis.opendocument.graphics")
        .unwrap();
    writer.add_file("content.xml", content.as_bytes()).unwrap();
    writer.finish_to_bytes().unwrap()
}

fn transition() -> Transition {
    let mut value = Transition::new();
    value.set_transition_type(Some("automatic")).unwrap();
    value.set_style(Some("dissolve")).unwrap();
    value
}

#[test]
fn first_transition_can_be_added_to_a_style_without_properties() {
    let content = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" office:version="1.4"><office:automatic-styles><style:style style:name="dp1" style:family="drawing-page"></style:style></office:automatic-styles><office:body><office:drawing><draw:page draw:name="Page 1" draw:style-name="dp1"/></office:drawing></office:body></office:document-content>"#;
    let source = Drawing::from_bytes(package(content)).unwrap();
    assert!(source.pages()[0].transition().is_none());

    let mut edit = source.edit();
    edit.set_page_transition(0, Some(transition())).unwrap();
    let output = edit.commit().unwrap().into_snapshot();
    assert!(
        output
            .content_xml()
            .contains("<style:drawing-page-properties")
    );
    let reopened = Drawing::from_bytes(output.as_bytes().to_vec()).unwrap();
    assert_eq!(
        reopened.pages()[0].transition().unwrap().style(),
        Some("dissolve")
    );
}

#[test]
fn accepts_group_and_nested_dr3d_scene_owners() {
    let content = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" office:version="1.4"><office:body><office:drawing><draw:page draw:name="Page 1"><draw:g draw:name="Group"><dr3d:scene draw:name="Outer"><dr3d:scene draw:name="Inner"><dr3d:cube/></dr3d:scene></dr3d:scene></draw:g></draw:page></office:drawing></office:body></office:document-content>"#;
    let drawing = Drawing::from_bytes(package(content)).unwrap();
    let shapes = drawing.pages()[0].shapes();
    assert_eq!(shapes.len(), 4);
    assert_eq!(shapes[0].kind(), ShapeKind::Group);
    assert_eq!(shapes[1].kind(), ShapeKind::ThreeDimensionalScene);
    assert_eq!(shapes[2].kind(), ShapeKind::ThreeDimensionalScene);
    assert_eq!(shapes[3].kind(), ShapeKind::ThreeDimensionalCube);
}

#[test]
fn accepts_repeated_frame_and_scene_glue_points_and_scene_description_first() {
    let content = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" office:version="1.4"><office:body><office:drawing><draw:page draw:name="Page 1"><draw:frame draw:name="Frame"><draw:glue-point draw:id="0" svg:x="1cm" svg:y="2cm" draw:escape-direction="auto"/><draw:glue-point draw:id="1" svg:x="3cm" svg:y="4cm" draw:escape-direction="auto"/></draw:frame><dr3d:scene draw:name="Scene"><svg:desc>Scene description</svg:desc><draw:glue-point draw:id="2" svg:x="5cm" svg:y="6cm" draw:escape-direction="auto"/><draw:glue-point draw:id="3" svg:x="7cm" svg:y="8cm" draw:escape-direction="auto"/></dr3d:scene></draw:page></office:drawing></office:body></office:document-content>"#;
    let drawing = Drawing::from_bytes(package(content)).unwrap();
    let shapes = drawing.pages()[0].shapes();
    assert_eq!(shapes[0].glue_points().len(), 2);
    assert_eq!(shapes[1].glue_points().len(), 2);
    assert_eq!(shapes[1].description(), Some("Scene description"));
}

#[test]
fn rejects_accessibility_elements_after_their_owner_sequence() {
    let frame_after_contour = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" office:version="1.4"><office:body><office:drawing><draw:page draw:name="Page 1"><draw:frame draw:name="Frame"><svg:title>Frame title</svg:title><draw:contour-polygon draw:recreate-on-edit="true" svg:width="1cm" svg:height="1cm" svg:viewBox="0 0 10 10" draw:points="0,0 1,1"/><svg:desc>Late description</svg:desc></draw:frame></draw:page></office:drawing></office:body></office:document-content>"#;
    let scene_title_after_description = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" office:version="1.4"><office:body><office:drawing><draw:page draw:name="Page 1"><dr3d:scene draw:name="Scene"><svg:desc>Description</svg:desc><svg:title>Late title</svg:title></dr3d:scene></draw:page></office:drawing></office:body></office:document-content>"#;

    for content in [frame_after_contour, scene_title_after_description] {
        assert!(Drawing::from_bytes(package(content)).is_err());
    }
}
