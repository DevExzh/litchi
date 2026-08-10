#![allow(clippy::unwrap_used, reason = "test assertions use unwrap for clarity")]

use litchi_odf_common::core::PackageWriter;
use litchi_odg::{
    Drawing,
    layer::Layer,
    page::Page,
    shape::{Shape, ShapeKind},
};

const CONTENT: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" office:version="1.4"><office:body><office:drawing><draw:page draw:name="One"><draw:layer-set><draw:layer draw:name="base"/></draw:layer-set><draw:rect draw:name="Box" draw:layer="base" draw:style-name="gr1" svg:x="1cm" svg:y="2cm" svg:width="3cm" svg:height="4cm"><text:p>old</text:p></draw:rect><draw:frame draw:name="Image"><draw:image xlink:href="Pictures/pixel.bin" xlink:type="simple"/></draw:frame></draw:page></office:drawing></office:body></office:document-content>"#;

fn package() -> Vec<u8> {
    let mut writer = PackageWriter::new();
    writer
        .set_mimetype("application/vnd.oasis.opendocument.graphics")
        .unwrap();
    writer.add_file("content.xml", CONTENT.as_bytes()).unwrap();
    writer
        .add_file_with_media_type("Pictures/pixel.bin", b"before", "image/png")
        .unwrap();
    writer.finish_to_bytes().unwrap()
}

#[test]
fn one_transaction_composes_structure_geometry_style_layers_and_resources() {
    let source = Drawing::from_bytes(package()).unwrap();
    assert_eq!(source.resources().len(), 1);
    assert_eq!(
        source.resource_bytes(0).unwrap().as_deref(),
        Some(&b"before"[..])
    );

    let mut edit = source.edit();
    edit.set_shape_geometry(0, 0, "2cm", "3cm", "5cm", "6cm")
        .unwrap();
    edit.set_shape_style_name(0, 0, "gr2").unwrap();
    edit.set_shape_text(0, 0, "new & compact").unwrap();
    edit.add_layer(0, Layer::new("notes").with_display("always"))
        .unwrap();
    edit.add_shape(
        0,
        Shape::new(ShapeKind::Ellipse)
            .with_name("Circle")
            .with_layer("notes")
            .with_geometry("1cm", "1cm", "2cm", "2cm"),
    )
    .unwrap();
    edit.add_group(0, "Group").unwrap();
    edit.add_page(Page::new("Two")).unwrap();
    edit.add_shape(
        1,
        Shape::new(ShapeKind::Rectangle)
            .with_name("Second")
            .with_geometry("0cm", "0cm", "1cm", "1cm"),
    )
    .unwrap();
    edit.set_resource(0, "image/png", b"after".to_vec())
        .unwrap();

    let commit = edit.commit().unwrap();
    assert!(commit.changed());
    assert_eq!(commit.snapshot().pages().len(), 2);
    assert_eq!(commit.snapshot().pages()[0].layers().len(), 2);
    assert_eq!(commit.snapshot().pages()[0].shapes().len(), 4);
    assert_eq!(commit.snapshot().pages()[0].shapes()[0].x(), Some("2cm"));
    assert_eq!(
        commit.snapshot().pages()[0].shapes()[0].style_name(),
        Some("gr2")
    );
    assert_eq!(
        commit.snapshot().pages()[1].shapes()[0].name(),
        Some("Second")
    );
    assert_eq!(
        commit.snapshot().resource_bytes(0).unwrap().as_deref(),
        Some(&b"after"[..])
    );
    assert!(!commit.snapshot().content_xml().contains(">\n<"));
    assert!(commit.patch().changes().len() >= 8);
    assert_eq!(commit.patch().resource_changes().len(), 1);
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(commit.snapshot())
            .unwrap()
            .as_bytes(),
        source.as_bytes()
    );
}

#[test]
fn removals_are_checked_and_group_subtrees_are_owned() {
    let group_content = CONTENT.replace(
        "<draw:frame draw:name=\"Image\">",
        "<draw:g draw:name=\"Group\"><draw:rect draw:name=\"Child\"/></draw:g><draw:frame draw:name=\"Image\">",
    );
    let mut writer = PackageWriter::new();
    writer
        .set_mimetype("application/vnd.oasis.opendocument.graphics")
        .unwrap();
    writer
        .add_file("content.xml", group_content.as_bytes())
        .unwrap();
    writer
        .add_file_with_media_type("Pictures/pixel.bin", b"before", "image/png")
        .unwrap();
    let source = Drawing::from_bytes(writer.finish_to_bytes().unwrap()).unwrap();

    let mut blocked = source.edit();
    assert!(blocked.remove_layer(0, "base").is_err());

    let mut edit = source.edit();
    let removed = edit.remove_shape(0, 1).unwrap();
    assert_eq!(removed.kind(), ShapeKind::Group);
    let commit = edit.commit().unwrap();
    assert!(
        commit.snapshot().pages()[0]
            .shapes()
            .iter()
            .all(|shape| shape.name() != Some("Child") && shape.name() != Some("Group"))
    );
}

#[test]
fn nested_group_text_geometry_style_and_form_edits_are_owned_and_reversible() {
    let group_content = CONTENT
        .replace(
            r#"xmlns:xlink="http://www.w3.org/1999/xlink""#,
            r#"xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0""#,
        )
        .replace(
            "<office:body>",
            r##"<office:automatic-styles><style:style style:name="gr2" style:family="graphic"><style:graphic-properties draw:fill-color="#00ff00"/></style:style></office:automatic-styles><office:body>"##,
        )
        .replace(
            "<draw:frame draw:name=\"Image\">",
            r#"<draw:g draw:name="Group"><draw:g draw:name="Nested"><draw:rect draw:name="Child" draw:style-name="gr1" svg:x="1cm" svg:y="1cm" svg:width="2cm" svg:height="2cm"><text:p>child</text:p></draw:rect><draw:control draw:name="Control" draw:style-name="gr1" draw:control="control1"/></draw:g></draw:g><draw:frame draw:name="Image">"#,
        );
    let source = drawing_with_resource(&group_content, Some(b"before"));
    let group = source.group(0, 1).unwrap();
    assert_eq!(group.descendants(), &[2, 3, 4]);

    let mut edit = source.edit();
    edit.set_group_text(0, 1, "bulk").unwrap();
    edit.set_group_descendant_text(0, 1, 3, "owned text")
        .unwrap();
    edit.set_group_descendant_geometry(0, 1, 3, "3cm", "4cm", "5cm", "6cm")
        .unwrap();
    edit.set_group_style_name(0, 1, "gr2").unwrap();
    edit.set_group_descendant_control_reference(0, 1, 4, "control2")
        .unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(
        commit.snapshot().pages()[0].shapes()[3].text(),
        "owned text"
    );
    assert_eq!(commit.snapshot().pages()[0].shapes()[3].x(), Some("3cm"));
    assert_eq!(
        commit.snapshot().pages()[0].shapes()[3].style_name(),
        Some("gr2")
    );
    assert_eq!(
        commit.snapshot().pages()[0].shapes()[4].control_reference(),
        Some("control2")
    );
    assert_eq!(
        commit
            .patch()
            .durable()
            .unwrap()
            .inverse()
            .apply(commit.snapshot())
            .unwrap()
            .as_bytes(),
        source.as_bytes()
    );
}

#[test]
fn cross_drawing_group_transfer_carries_layers_resources_and_durable_inverse() {
    const SOURCE: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" office:version="1.4"><office:body><office:drawing><draw:page draw:name="Source"><draw:layer-set><draw:layer draw:name="media"/></draw:layer-set><draw:g draw:name="Transferred" draw:layer="media"><draw:frame draw:name="Picture" draw:layer="media"><draw:image xlink:href="Pictures/transfer.bin" xlink:type="simple"/></draw:frame></draw:g></draw:page></office:drawing></office:body></office:document-content>"#;
    const DESTINATION: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" office:version="1.4"><office:body><office:drawing><draw:page draw:name="Destination"/></office:drawing></office:body></office:document-content>"#;

    let source = drawing_with_resource(SOURCE, Some(b"transfer bytes"));
    let destination = drawing_with_resource(DESTINATION, None);
    let transfer = source.snapshot().prepare_shape_transfer(0, 0).unwrap();
    assert_eq!(transfer.shape().kind(), ShapeKind::Group);
    assert_eq!(transfer.layers()[0].name(), "media");
    assert_eq!(transfer.resources()[0].path(), "Pictures/transfer.bin");

    let mut edit = destination.edit();
    edit.insert_shape_transfer(0, 0, &transfer).unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(commit.snapshot().pages()[0].layers()[0].name(), "media");
    assert_eq!(commit.snapshot().pages()[0].shapes().len(), 2);
    assert_eq!(
        commit.snapshot().resource_bytes(0).unwrap().as_deref(),
        Some(&b"transfer bytes"[..])
    );
    let durable = commit.patch().durable().unwrap();
    assert_eq!(
        durable.apply(destination.snapshot()).unwrap().as_bytes(),
        commit.snapshot().as_bytes()
    );
    assert_eq!(
        durable
            .inverse()
            .apply(commit.snapshot())
            .unwrap()
            .as_bytes(),
        destination.as_bytes()
    );

    let limits = litchi_core::CompositionLimits::new(4, 4, 8, 8);
    let mut competing_edit = destination.edit();
    competing_edit
        .add_page(Page::new("Competing page"))
        .unwrap();
    let competing = competing_edit.commit().unwrap();
    let mut joined = destination.joined_edits(limits);
    joined
        .join(commit.patch().prepare("transfer", limits).unwrap())
        .unwrap();
    assert!(
        joined
            .join(competing.patch().prepare("structure", limits).unwrap())
            .is_err()
    );

    let collision = drawing_with_resource(DESTINATION, Some(b"different"));
    let mut remapped = collision.edit();
    remapped.insert_shape_transfer(0, 0, &transfer).unwrap();
    let remapped = remapped.commit().unwrap();
    assert!(remapped.snapshot().content_xml().contains("_litchi_"));
    assert_eq!(
        remapped.snapshot().resource_bytes(0).unwrap().as_deref(),
        Some(&b"transfer bytes"[..])
    );
}

#[test]
fn cross_drawing_transfer_copies_and_remaps_arbitrary_style_definitions() {
    const SOURCE: &str = r##"<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"><office:automatic-styles><style:style style:name="gr1" style:family="graphic"><style:graphic-properties draw:fill="solid" draw:fill-color="#ff0000"/></style:style></office:automatic-styles><office:body><office:drawing><draw:page draw:name="Source"><draw:rect draw:name="Styled" draw:style-name="gr1"/></draw:page></office:drawing></office:body></office:document-content>"##;
    const DESTINATION: &str = r##"<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"><office:automatic-styles><style:style style:name="gr1" style:family="graphic"><style:graphic-properties draw:fill="solid" draw:fill-color="#0000ff"/></style:style></office:automatic-styles><office:body><office:drawing><draw:page draw:name="Destination"/></office:drawing></office:body></office:document-content>"##;
    let source = drawing_with_resource(SOURCE, None);
    let destination = drawing_with_resource(DESTINATION, None);
    let transfer = source.snapshot().prepare_shape_transfer(0, 0).unwrap();
    assert_eq!(transfer.style_definitions()[0].name(), "gr1");

    let mut edit = destination.edit();
    edit.insert_shape_transfer(0, 0, &transfer).unwrap();
    let commit = edit.commit().unwrap();
    let shape_style = commit.snapshot().pages()[0].shapes()[0]
        .style_name()
        .unwrap();
    assert!(shape_style.starts_with("gr1_litchi_"));
    assert!(commit.snapshot().style_definitions().iter().any(|style| {
        style.name() == shape_style && style.property("draw:fill-color") == Some("#ff0000")
    }));
    assert_eq!(
        commit
            .patch()
            .durable()
            .unwrap()
            .inverse()
            .apply(commit.snapshot())
            .unwrap()
            .as_bytes(),
        destination.as_bytes()
    );
}

#[test]
fn nested_group_transfer_remaps_style_form_and_resource_collisions_together() {
    const SOURCE: &str = r##"<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" xmlns:xlink="http://www.w3.org/1999/xlink"><office:automatic-styles><draw:gradient draw:name="gradient1" draw:style="linear" draw:start-color="#ff0000" draw:end-color="#ffff00"/><style:style style:name="gr1" style:family="graphic"><style:graphic-properties draw:fill="gradient" draw:fill-gradient-name="gradient1"/></style:style></office:automatic-styles><office:body><office:drawing><office:forms><form:form form:name="Source"><form:checkbox form:id="control1" form:label="source"/></form:form></office:forms><draw:page draw:name="Source"><draw:g draw:name="Outer"><draw:g draw:name="Inner"><draw:control draw:name="Choice" draw:style-name="gr1" draw:control="control1"/><draw:frame draw:name="Media" draw:style-name="gr1"><draw:image xlink:href="Pictures/transfer.bin" xlink:type="simple"/></draw:frame></draw:g></draw:g></draw:page></office:drawing></office:body></office:document-content>"##;
    const DESTINATION: &str = r##"<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0"><office:automatic-styles><draw:gradient draw:name="gradient1" draw:style="radial" draw:start-color="#0000ff" draw:end-color="#00ffff"/><style:style style:name="gr1" style:family="graphic"><style:graphic-properties draw:fill="gradient" draw:fill-gradient-name="gradient1"/></style:style></office:automatic-styles><office:body><office:drawing><office:forms><form:form form:name="Destination"><form:checkbox form:id="control1" form:label="destination"/></form:form></office:forms><draw:page draw:name="Destination"/></office:drawing></office:body></office:document-content>"##;
    let source = drawing_with_resource(SOURCE, Some(b"source resource"));
    let destination = drawing_with_resource(DESTINATION, Some(b"destination resource"));
    let transfer = source.snapshot().prepare_shape_transfer(0, 0).unwrap();
    assert_eq!(transfer.shape().kind(), ShapeKind::Group);
    assert_eq!(transfer.style_resources().len(), 1);
    assert_eq!(source.group(0, 0).unwrap().descendants().len(), 3);
    assert_eq!(transfer.resources().len(), 1, "{:#?}", transfer.resources());

    let mut edit = destination.edit();
    edit.insert_shape_transfer(0, 0, &transfer).unwrap();
    let commit = edit.commit().unwrap();
    let transferred = commit.snapshot().group(0, 0).unwrap();
    assert_eq!(transferred.descendants().len(), 3);
    assert_eq!(commit.snapshot().form_controls().len(), 2);
    assert!(commit.snapshot().form_controls().iter().any(|control| {
        control.id().starts_with("control1_litchi_")
            && control.attributes().get("form:label").map(String::as_str) == Some("source")
    }));
    assert!(commit.snapshot().style_definitions().iter().any(|style| {
        style.name().starts_with("gr1_litchi_")
            && style
                .property("draw:fill-gradient-name")
                .is_some_and(|name| name.starts_with("gradient1_litchi_"))
    }));
    assert!(commit.snapshot().style_resources().iter().any(|resource| {
        resource.name().starts_with("gradient1_litchi_")
            && resource
                .attributes()
                .get("draw:start-color")
                .map(String::as_str)
                == Some("#ff0000")
    }));
    let remapped_resource = commit
        .snapshot()
        .resources()
        .iter()
        .position(|resource| resource.path().starts_with("Pictures/transfer_litchi_"))
        .unwrap();
    assert_eq!(
        commit
            .snapshot()
            .resource_bytes(remapped_resource)
            .unwrap()
            .as_deref(),
        Some(&b"source resource"[..])
    );
    assert_eq!(
        commit
            .patch()
            .durable()
            .unwrap()
            .inverse()
            .apply(commit.snapshot())
            .unwrap()
            .as_bytes(),
        destination.as_bytes()
    );
}

#[test]
fn unreferenced_resource_add_remove_is_safe_and_exactly_reversible() {
    let source = Drawing::from_bytes(package()).unwrap();
    let mut addition = source.edit();
    addition
        .add_resource("Media/sound.bin", "audio/basic", b"sound".to_vec())
        .unwrap();
    assert!(
        addition
            .add_resource("../escape.bin", "audio/basic", Vec::new())
            .is_err()
    );
    let added = addition.commit().unwrap();
    assert!(
        added
            .snapshot()
            .files()
            .unwrap()
            .contains(&"Media/sound.bin".to_string())
    );
    let restored = added
        .patch()
        .durable()
        .unwrap()
        .inverse()
        .apply(added.snapshot())
        .unwrap();
    assert_eq!(restored.as_bytes(), source.as_bytes());

    let mut removal = added.snapshot().edit();
    removal
        .remove_unreferenced_resource("Media/sound.bin")
        .unwrap();
    let removed = removal.commit().unwrap();
    assert!(
        !removed
            .snapshot()
            .files()
            .unwrap()
            .contains(&"Media/sound.bin".to_string())
    );
}

fn drawing_with_resource(content: &str, resource: Option<&[u8]>) -> Drawing {
    let mut writer = PackageWriter::new();
    writer
        .set_mimetype("application/vnd.oasis.opendocument.graphics")
        .unwrap();
    writer.add_file("content.xml", content.as_bytes()).unwrap();
    if let Some(bytes) = resource {
        writer
            .add_file_with_media_type("Pictures/transfer.bin", bytes, "application/octet-stream")
            .unwrap();
    }
    Drawing::from_bytes(writer.finish_to_bytes().unwrap()).unwrap()
}
