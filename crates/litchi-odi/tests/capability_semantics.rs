#![allow(clippy::unwrap_used, reason = "test assertions use unwrap for clarity")]

use litchi_odi::{
    Builder, FlatImage, FrameEditor, History, Image,
    frame::Frame,
    map::{Area, AreaKind, ImageMap},
    source::Source,
};

const FLAT: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8"?><office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" office:mimetype="application/vnd.oasis.opendocument.image"><office:body><office:image>"#,
    r#"<draw:frame draw:name="Photo" draw:style-name="gr1" draw:text-style-name="P1" draw:layer="layout" draw:z-index="7" draw:transform="rotate (0.25)" text:anchor-type="paragraph" svg:x="1cm" svg:y="2cm" svg:width="3cm" svg:height="4cm" style:rel-width="80%" style:rel-height="scale">"#,
    r#"<draw:image xml:id="image1" draw:filter-name="png" draw:mime-type="image/png" xlink:href="Pictures/photo.png" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>"#,
    r##"<draw:image-map><draw:area-rectangle svg:x="0cm" svg:y="0cm" svg:width="1cm" svg:height="2cm" xlink:type="simple" xlink:href="https://example.test/#region" office:name="Region"><svg:title>Go &amp; see</svg:title><svg:desc>Rectangle</svg:desc></draw:area-rectangle><draw:area-circle svg:cx="2cm" svg:cy="2cm" svg:r="1cm" draw:nohref="nohref"/></draw:image-map>"##,
    r#"<svg:title>Photo title</svg:title><svg:desc>Photo description</svg:desc></draw:frame></office:image></office:body></office:document>"#,
);

fn edit_frame(editor: &mut impl FrameEditor) {
    editor.set_style_name(0, Some("gr2".into())).unwrap();
    editor.set_layer(0, Some("foreground".into())).unwrap();
    editor.set_z_index(0, Some(9)).unwrap();
    editor
        .set_geometry(
            0,
            Some("5cm".into()),
            Some("6cm".into()),
            Some("7cm".into()),
            Some("8cm".into()),
        )
        .unwrap();
}

#[test]
fn image_maps_and_broader_frame_semantics_are_inert_and_bounded() {
    let image = FlatImage::from_bytes(FLAT.as_bytes().to_vec()).unwrap();
    let frame = image.frame().unwrap();
    assert_eq!(frame.style_name(), Some("gr1"));
    assert_eq!(frame.text_style_name(), Some("P1"));
    assert_eq!(frame.layer(), Some("layout"));
    assert_eq!(frame.z_index(), Some(7));
    assert_eq!(frame.transform(), Some("rotate (0.25)"));
    assert_eq!(frame.anchor_type(), Some("paragraph"));
    assert_eq!(frame.relative_width(), Some("80%"));
    assert_eq!(frame.relative_height(), Some("scale"));
    assert_eq!(frame.image_xml_id(), Some("image1"));
    assert_eq!(frame.filter_name(), Some("png"));
    assert_eq!(frame.link_type(), Some("simple"));
    assert_eq!(frame.show(), Some("embed"));
    assert_eq!(frame.actuate(), Some("onLoad"));
    let areas = frame.image_map().unwrap().areas();
    assert_eq!(areas.len(), 2);
    assert_eq!(areas[0].href(), Some("https://example.test/#region"));
    assert_eq!(areas[0].title(), Some("Go & see"));
    assert!(matches!(areas[0].kind(), AreaKind::Rectangle { .. }));
    assert!(areas[1].has_no_href());
    assert!(matches!(areas[1].kind(), AreaKind::Circle { .. }));
}

#[test]
fn malformed_image_maps_are_rejected() {
    let missing_geometry = FLAT.replace(r#" svg:width="1cm""#, "");
    assert!(FlatImage::from_bytes(missing_geometry.into_bytes()).is_err());

    let duplicate = FLAT.replace("</draw:image-map>", "</draw:image-map><draw:image-map/>");
    assert!(FlatImage::from_bytes(duplicate.into_bytes()).is_err());

    let spoofed = FLAT
        .replace("draw:area-circle", "fake:area-circle")
        .replace(
            "xmlns:xlink=",
            r#"xmlns:fake="urn:not-drawing" xmlns:xlink="#,
        );
    assert!(FlatImage::from_bytes(spoofed.into_bytes()).is_err());
}

#[test]
fn authored_maps_styles_metadata_and_graph_round_trip() {
    let map = ImageMap::new(vec![
        Area::rectangle("0cm", "0cm", "2cm", "1cm")
            .with_href("https://example.test/map")
            .with_name("Map")
            .with_title("Map title"),
        Area::polygon(
            ["0cm", "0cm", "2cm", "2cm"],
            "0 0 100 100",
            "0,0 100,0 50,100",
        )
        .with_no_href(),
    ]);
    let frame = Frame::new(Source::Linked("Pictures/photo.png".into()))
        .with_name("Photo")
        .with_style_name("gr1")
        .with_layer("layout")
        .with_z_index(2)
        .with_geometry("1cm", "2cm", "3cm", "4cm")
        .with_relative_size("90%", "scale")
        .with_image_map(map);
    let styles = r#"<?xml version="1.0"?><office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"><office:styles><style:style style:name="gr1" style:family="graphic"/></office:styles></office:document-styles>"#;
    let meta = r#"<?xml version="1.0"?><office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"><office:meta><dc:title>Before</dc:title><meta:user-defined meta:name="opaque">keep</meta:user-defined></office:meta></office:document-meta>"#;
    let source = Image::from_bytes(
        Builder::new()
            .frame(&frame)
            .styles_xml(styles)
            .meta_xml(meta)
            .resource("Pictures/photo.png", "image/png", b"png".to_vec())
            .resource("Thumbnails/thumbnail.png", "image/png", b"thumb".to_vec())
            .build()
            .unwrap(),
    )
    .unwrap();
    assert_eq!(
        source.frame().unwrap().image_map().unwrap().areas().len(),
        2
    );
    assert_eq!(source.styles_xml(), Some(styles));
    assert!(source.resource_graph().nodes().iter().any(|node| {
        node.path() == "Thumbnails/thumbnail.png" && node.is_present() && !node.is_referenced()
    }));
    assert_eq!(source.resource_graph().edges().len(), 1);

    let mut edit = source.edit();
    edit_frame(&mut edit);
    edit.set_title(Some("After".into())).unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(commit.image().frame().unwrap().style_name(), Some("gr2"));
    assert_eq!(
        commit.image().metadata().unwrap().title.as_deref(),
        Some("After")
    );
    assert!(
        commit
            .image()
            .meta_xml()
            .unwrap()
            .unwrap()
            .contains("opaque")
    );
    assert!(commit.patch().metadata_change().is_some());
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(commit.image())
            .unwrap()
            .as_bytes(),
        source.as_bytes()
    );
}

#[test]
fn shared_edit_contract_and_history_work_for_flat_and_package_snapshots() {
    let source = FlatImage::from_bytes(FLAT.as_bytes().to_vec()).unwrap();
    let mut edit = source.transaction();
    edit_frame(&mut edit);
    let target = edit.commit().unwrap().into_snapshot();
    let mut history = History::new(source.clone(), 3).unwrap();
    assert!(history.record(&source, target.clone()).unwrap());
    assert_eq!(history.undo().unwrap().as_bytes(), source.as_bytes());
    assert_eq!(history.redo().unwrap().as_bytes(), target.as_bytes());
    assert!(history.record(&source, target).is_err());
}
