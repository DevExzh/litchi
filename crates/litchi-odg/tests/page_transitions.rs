#![allow(clippy::unwrap_used, reason = "test assertions use unwrap for clarity")]

use litchi_odf_common::core::PackageWriter;
use litchi_odg::{Drawing, Transition, TransitionSound, page::Page};

const CONTENT_AUTOMATIC: &str = r##"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:smil="urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:xml="http://www.w3.org/XML/1998/namespace" xmlns:foo="urn:example:unknown" office:version="1.4"><office:automatic-styles><style:style style:name="dp1" style:family="drawing-page"><style:drawing-page-properties presentation:transition-type="automatic" presentation:transition-style="fade-from-left" presentation:transition-speed="fast" smil:type="fade" smil:subtype="crossfade" smil:direction="forward" smil:fadeColor="#010203" presentation:duration="PT2S"><presentation:sound xlink:type="simple" xlink:href="media/transition.wav" xlink:actuate="onRequest" xlink:show="replace" presentation:play-full="true" xml:id="sound1"/><foo:unknown foo:value="keep"/></style:drawing-page-properties></style:style></office:automatic-styles><office:body><office:drawing><draw:page draw:name="Page 1" draw:style-name="dp1"/></office:drawing></office:body></office:document-content>"##;

const CONTENT_NAMED: &str = r##"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:xml="http://www.w3.org/XML/1998/namespace" office:version="1.4"><office:body><office:drawing><draw:page draw:name="Page 1" draw:style-name="dp1"/></office:drawing></office:body></office:document-content>"##;

const CONTENT_SHARED: &str = r##"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:smil="urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0" office:version="1.4"><office:automatic-styles><style:style style:name="dp1" style:family="drawing-page"><style:drawing-page-properties presentation:transition-type="automatic"/></style:style></office:automatic-styles><office:body><office:drawing><draw:page draw:name="Page 1" draw:style-name="dp1"/><draw:page draw:name="Page 2" draw:style-name="dp1"/></office:drawing></office:body></office:document-content>"##;

const CONTENT_3D: &str = r##"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" office:version="1.4"><office:body><office:drawing><draw:page draw:name="Page 1"><dr3d:scene draw:name="Scene"><dr3d:light dr3d:direction="(0 0 1)"/><dr3d:cube/></dr3d:scene></draw:page></office:drawing></office:body></office:document-content>"##;

const CONTENT_ENHANCED: &str = r##"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" office:version="1.4"><office:body><office:drawing><draw:page draw:name="Page 1"><draw:custom-shape draw:name="Custom"><draw:enhanced-geometry draw:type="rectangle" svg:viewBox="0 0 21600 21600"><draw:equation draw:name="f0" draw:formula="width/2"/><draw:handle draw:handle-position="$0 0"/></draw:enhanced-geometry></draw:custom-shape></draw:page></office:drawing></office:body></office:document-content>"##;

const CONTENT_AUXILIARY: &str = r##"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" office:version="1.4"><office:body><office:drawing><draw:page draw:name="Page 1"><draw:frame draw:name="Picture"><draw:image xlink:type="simple" xlink:href="media/picture.png"/><draw:glue-point draw:id="0" svg:x="1cm" svg:y="2cm" draw:escape-direction="auto"/><draw:image-map><draw:area-rectangle svg:x="0cm" svg:y="0cm" svg:width="2cm" svg:height="3cm" xlink:type="simple" xlink:href="https://example.org/target" xlink:show="replace" office:name="area-1"/><draw:area-polygon svg:x="0cm" svg:y="0cm" svg:width="2cm" svg:height="3cm" svg:viewBox="0 0 100 100" draw:points="0,0 100,0 50,100" draw:nohref="nohref"/></draw:image-map><draw:contour-polygon draw:recreate-on-edit="true" svg:width="2cm" svg:height="3cm" svg:viewBox="0 0 100 100" draw:points="0,0 100,0 50,100"/></draw:frame></draw:page></office:drawing></office:body></office:document-content>"##;

const STYLES_NAMED: &str = r##"<?xml version="1.0" encoding="UTF-8"?><office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:smil="urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0" office:version="1.4"><office:styles><style:style style:name="dp1" style:family="drawing-page"><style:drawing-page-properties presentation:transition-type="manual" smil:type="fade"/></style:style></office:styles></office:document-styles>"##;

const STYLES_PREFIXED_INHERITED: &str = r##"<?xml version="1.0" encoding="UTF-8"?><office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:p="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0" office:version="1.4"><office:styles><style:style style:name="parent" style:family="drawing-page"><style:drawing-page-properties p:transition-speed="slow" s:type="fade"/></style:style><style:style style:name="dp1" style:family="drawing-page" style:parent-style-name="parent"><style:drawing-page-properties p:transition-type="manual"/></style:style></office:styles></office:document-styles>"##;

fn package(content: &str, styles: Option<&str>) -> Vec<u8> {
    let mut writer = PackageWriter::new();
    writer
        .set_mimetype("application/vnd.oasis.opendocument.graphics")
        .unwrap();
    writer.add_file("content.xml", content.as_bytes()).unwrap();
    writer.add_file("media/transition.wav", b"inert").unwrap();
    if let Some(styles) = styles {
        writer.add_file("styles.xml", styles.as_bytes()).unwrap();
    }
    writer.finish_to_bytes().unwrap()
}

fn replacement_transition() -> Transition {
    let mut transition = Transition::new();
    transition.set_transition_type(Some("automatic")).unwrap();
    transition.set_style(Some("dissolve")).unwrap();
    transition.set_speed(Some("medium")).unwrap();
    transition.set_smil_type(Some("fade")).unwrap();
    transition.set_smil_subtype(Some("crossfade")).unwrap();
    transition.set_direction(Some("reverse")).unwrap();
    transition.set_fade_color(Some("#AABBCC")).unwrap();
    transition.set_duration(Some("PT3S")).unwrap();
    let sound = TransitionSound::new("media/transition.wav")
        .unwrap()
        .with_play_full(Some(true))
        .with_actuate_on_request(true)
        .with_show(Some("replace"))
        .unwrap()
        .with_xml_id(Some("sound1"))
        .unwrap();
    transition.set_sound(Some(sound));
    transition
}

#[test]
fn reads_smil_transition_and_reopens_source_bound_edit() {
    let source = Drawing::from_bytes(package(CONTENT_AUTOMATIC, None)).unwrap();
    let transition = source.pages()[0].transition().unwrap();
    assert_eq!(transition.transition_type(), Some("automatic"));
    assert_eq!(transition.style(), Some("fade-from-left"));
    assert_eq!(transition.smil_type(), Some("fade"));
    assert_eq!(transition.duration(), Some("PT2S"));
    assert_eq!(transition.sound().unwrap().href(), "media/transition.wav");
    assert!(transition.sound().unwrap().actuate_on_request());

    let mut edit = source.edit();
    edit.set_page_transition(0, Some(replacement_transition()))
        .unwrap();
    let commit = edit.commit().unwrap();
    let output = commit.snapshot();
    assert!(
        output
            .content_xml()
            .contains("foo:unknown foo:value=\"keep\"")
    );
    assert!(
        output
            .content_xml()
            .contains("presentation:transition-style=\"dissolve\"")
    );
    let reopened = Drawing::from_bytes(output.as_bytes().to_vec()).unwrap();
    assert_eq!(
        reopened.pages()[0].transition(),
        output.pages()[0].transition()
    );
    assert_eq!(
        reopened.pages()[0].transition().unwrap().duration(),
        Some("PT3S")
    );

    let durable = commit.patch().durable().unwrap();
    let replayed = durable.apply(source.snapshot()).unwrap();
    assert_eq!(replayed.as_bytes(), output.as_bytes());
    let restored = durable.inverse().apply(&replayed).unwrap();
    assert_eq!(restored.as_bytes(), source.as_bytes());
}

#[test]
fn preserves_paired_empty_transition_sound_markup() {
    let content = CONTENT_AUTOMATIC.replace(
        r#"xml:id="sound1"/>"#,
        r#"xml:id="sound1"><!--keep--></presentation:sound>"#,
    );
    let source = Drawing::from_bytes(package(&content, None)).unwrap();
    let mut transition = source.pages()[0].transition().cloned().unwrap();
    transition.set_duration(Some("PT3S")).unwrap();
    let mut edit = source.edit();
    edit.set_page_transition(0, Some(transition)).unwrap();
    let output = edit.commit().unwrap().into_snapshot();
    assert!(output.content_xml().contains("<!--keep-->"));
    assert_eq!(
        Drawing::from_bytes(output.as_bytes().to_vec())
            .unwrap()
            .pages()[0]
            .transition()
            .unwrap()
            .duration(),
        Some("PT3S")
    );
}

#[test]
fn edits_named_style_owner_in_styles_xml() {
    let source = Drawing::from_bytes(package(CONTENT_NAMED, Some(STYLES_NAMED))).unwrap();
    assert_eq!(
        source.pages()[0].transition().unwrap().transition_type(),
        Some("manual")
    );
    let mut edit = source.edit();
    edit.set_page_transition(0, Some(replacement_transition()))
        .unwrap();
    let output = edit.commit().unwrap().into_snapshot();
    assert_eq!(output.content_xml(), CONTENT_NAMED);
    assert!(
        output
            .styles_xml()
            .unwrap()
            .contains("presentation:transition-style=\"dissolve\"")
    );
    let reopened = Drawing::from_bytes(output.as_bytes().to_vec()).unwrap();
    assert_eq!(
        reopened.pages()[0].transition().unwrap().style(),
        Some("dissolve")
    );
}

#[test]
fn resolves_prefixed_transition_attributes_and_parent_styles() {
    let source =
        Drawing::from_bytes(package(CONTENT_NAMED, Some(STYLES_PREFIXED_INHERITED))).unwrap();
    let transition = source.pages()[0].transition().unwrap();
    assert_eq!(transition.transition_type(), Some("manual"));
    assert_eq!(transition.speed(), Some("slow"));
    assert_eq!(transition.smil_type(), Some("fade"));
}

#[test]
fn refuses_to_mutate_a_transition_style_shared_by_pages() {
    let source = Drawing::from_bytes(package(CONTENT_SHARED, None)).unwrap();
    let mut edit = source.edit();
    assert!(
        edit.set_page_transition(0, Some(replacement_transition()))
            .is_err()
    );
}

#[test]
fn recognizes_inert_dr3d_shape_owners() {
    let drawing = Drawing::from_bytes(package(CONTENT_3D, None)).unwrap();
    let shapes = drawing.pages()[0].shapes();
    assert_eq!(
        shapes[0].kind(),
        litchi_odg::shape::ShapeKind::ThreeDimensionalScene
    );
    assert_eq!(
        shapes[1].kind(),
        litchi_odg::shape::ShapeKind::ThreeDimensionalLight
    );
    assert_eq!(
        shapes[2].kind(),
        litchi_odg::shape::ShapeKind::ThreeDimensionalCube
    );
}

#[test]
fn rejects_non_3d_dr3d_scene_children() {
    let non_3d_scene_child = CONTENT_3D.replace(
        r#"<dr3d:light dr3d:direction="(0 0 1)"/>"#,
        r#"<draw:rect/>"#,
    );
    assert!(matches!(
        Drawing::from_bytes(package(&non_3d_scene_child, None)),
        Err(litchi_core::Error::InvalidFormat(_))
    ));
}

#[test]
fn rejects_malformed_dr3d_vector_attributes() {
    let malformed = CONTENT_3D.replace(r#"dr3d:direction="(0 0 1)""#, r#"dr3d:direction="0 0 1""#);
    assert!(matches!(
        Drawing::from_bytes(package(&malformed, None)),
        Err(litchi_core::Error::InvalidFormat(_))
    ));
}

#[test]
fn refuses_inert_dr3d_shape_authoring() {
    let drawing = Drawing::from_bytes(package(CONTENT_3D, None)).unwrap();
    let mut edit = drawing.edit();
    let result = edit.add_shape(
        0,
        litchi_odg::shape::Shape::new(litchi_odg::shape::ShapeKind::ThreeDimensionalCube),
    );
    assert!(matches!(result, Err(litchi_core::Error::Unsupported(_))));
}

#[test]
fn reads_inert_enhanced_geometry_equations_and_handles() {
    let drawing = Drawing::from_bytes(package(CONTENT_ENHANCED, None)).unwrap();
    let geometry = drawing.pages()[0].shapes()[0].enhanced_geometry().unwrap();
    assert_eq!(geometry.attributes().len(), 2);
    assert_eq!(geometry.children().len(), 2);
    assert_eq!(
        geometry.children()[0].kind(),
        litchi_odg::EnhancedGeometryChildKind::Equation
    );
    assert_eq!(
        geometry.children()[1].kind(),
        litchi_odg::EnhancedGeometryChildKind::Handle
    );
    assert!(
        geometry.children()[0]
            .attributes()
            .iter()
            .any(|attribute| attribute.local_name() == "formula" && attribute.value() == "width/2")
    );

    let modern = CONTENT_ENHANCED.replace(
        r#"draw:handle draw:handle-position="$0 0""#,
        r#"draw:handle draw:handle-position-x="$0" draw:handle-position-y="$0""#,
    );
    let modern = Drawing::from_bytes(package(&modern, None)).unwrap();
    let modern_handle = &modern.pages()[0].shapes()[0]
        .enhanced_geometry()
        .unwrap()
        .children()[1];
    assert!(modern_handle.attributes().iter().any(|attribute| {
        attribute.local_name() == "handle-position-x" && attribute.value() == "$0"
    }));
}

#[test]
fn rejects_invalid_enhanced_geometry_handle_grammar_and_children() {
    let invalid_sources = [
        CONTENT_ENHANCED.replace(
            r#"draw:handle draw:handle-position="$0 0""#,
            r#"draw:handle draw:handle-position-x="$0""#,
        ),
        CONTENT_ENHANCED.replace(
            r#"draw:handle draw:handle-position="$0 0""#,
            r#"draw:handle draw:handle-position-x="$0" draw:handle-position-y="$0" draw:handle-polar-pole-x="0" draw:handle-polar-pole-y="0""#,
        ),
        CONTENT_ENHANCED.replace(
            r#"draw:handle draw:handle-position="$0 0""#,
            r#"draw:handle draw:handle-position="$0 0" draw:handle-switched="maybe""#,
        ),
        CONTENT_ENHANCED.replace(
            r#"<draw:handle draw:handle-position="$0 0"/>"#,
            r#"<draw:handle draw:handle-position="$0 0"><draw:unknown/></draw:handle>"#,
        ),
        CONTENT_ENHANCED.replace(
            r#"</draw:enhanced-geometry>"#,
            r#"</draw:enhanced-geometry><draw:enhanced-geometry draw:type="rectangle"/>"#,
        ),
    ];
    for (index, content) in invalid_sources.into_iter().enumerate() {
        assert!(
            matches!(
                Drawing::from_bytes(package(&content, None)),
                Err(litchi_core::Error::InvalidFormat(_))
            ),
            "invalid enhanced geometry source {index} was accepted"
        );
    }
}

#[test]
fn reads_inert_image_map_contour_and_glue_point_owners() {
    let drawing = Drawing::from_bytes(package(CONTENT_AUXILIARY, None)).unwrap();
    let shape = &drawing.pages()[0].shapes()[0];
    let map = shape.image_map().unwrap();
    assert_eq!(map.areas().len(), 2);
    assert_eq!(map.areas()[0].href(), Some("https://example.org/target"));
    assert_eq!(map.areas()[0].show(), Some("replace"));
    assert!(map.areas()[1].no_href());
    assert_eq!(shape.contours().len(), 1);
    assert_eq!(shape.contours()[0].kind(), litchi_odg::ContourKind::Polygon);
    assert_eq!(shape.glue_points().len(), 1);
    assert_eq!(shape.glue_points()[0].id(), "0");
    assert_eq!(shape.glue_points()[0].escape_direction(), "auto");
}

#[test]
fn image_map_charges_retained_area_and_owner_xml_together() {
    let start = CONTENT_AUXILIARY.find("<draw:area-rectangle").unwrap();
    let end = start + CONTENT_AUXILIARY[start..].find("/>").unwrap();
    for (comment_bytes, accepted) in [(16, true), (5 * 1024 * 1024, false)] {
        let paired = format!(
            "><!--{}--><!--{}--></draw:area-rectangle>",
            "x".repeat(comment_bytes / 2),
            "x".repeat(comment_bytes / 2)
        );
        let mut content = CONTENT_AUXILIARY.to_owned();
        content.replace_range(end..end + 2, &paired);
        let result = Drawing::from_bytes(package(&content, None));
        if accepted {
            assert_eq!(
                result.unwrap().pages()[0].shapes()[0]
                    .image_map()
                    .unwrap()
                    .areas()
                    .len(),
                2
            );
        } else {
            assert!(matches!(result, Err(litchi_core::Error::InvalidFormat(_))));
        }
    }
}

#[test]
fn paired_contours_and_glue_points_accept_empty_content_and_retain_comments() {
    let source = CONTENT_AUXILIARY
        .replace(
            r#"draw:points="0,0 100,0 50,100"/>"#,
            r#"draw:points="0,0 100,0 50,100"><!--keep--></draw:contour-polygon>"#,
        )
        .replace(
            r#"draw:escape-direction="auto"/>"#,
            r#"draw:escape-direction="auto"><!--glue--></draw:glue-point>"#,
        );
    let drawing = Drawing::from_bytes(package(&source, None)).unwrap();
    let shape = &drawing.pages()[0].shapes()[0];
    assert!(shape.contours()[0].source_xml().contains("<!--keep-->"));
    assert!(shape.glue_points()[0].source_xml().contains("<!--glue-->"));
    let nonempty = source.replace("<!--keep-->", "nonempty");
    assert!(matches!(
        Drawing::from_bytes(package(&nonempty, None)),
        Err(litchi_core::Error::InvalidFormat(_))
    ));
}

#[test]
fn rejects_invalid_contour_attributes_and_non_frame_auxiliary_owners() {
    let invalid_sources = [
        CONTENT_AUXILIARY.replace(r#"draw:recreate-on-edit="true" "#, ""),
        CONTENT_AUXILIARY.replace(
            r#" draw:points="0,0 100,0 50,100"/>"#,
            r#"/>"#,
        ),
        CONTENT_AUXILIARY.replace(
            r#"<draw:contour-polygon draw:recreate-on-edit="true" svg:width="2cm" svg:height="3cm" svg:viewBox="0 0 100 100" draw:points="0,0 100,0 50,100"/>"#,
            r#"<draw:contour-path draw:recreate-on-edit="true"/>"#,
        ),
        CONTENT_AUXILIARY.replace(
            r#"draw:escape-direction="auto""#,
            r#"draw:align="bottom" draw:escape-direction="auto""#,
        ),
            CONTENT_AUXILIARY
            .replace(
                r#"<draw:image xlink:type="simple" xlink:href="media/picture.png"/><draw:glue-point"#,
                r#"<draw:image xlink:type="simple" xlink:href="media/picture.png"><draw:glue-point"#,
            )
            .replace(
                r#"/></draw:image-map><draw:contour-polygon"#,
                r#"/></draw:image-map></draw:image><draw:contour-polygon"#,
            ),
        CONTENT_AUXILIARY.replace(
            r#"</draw:frame></draw:page>"#,
            r#"<draw:contour-path draw:recreate-on-edit="false" svg:d="M 0 0"/></draw:frame></draw:page>"#,
        ),
    ];
    for (index, content) in invalid_sources.into_iter().enumerate() {
        assert!(
            matches!(
                Drawing::from_bytes(package(&content, None)),
                Err(litchi_core::Error::InvalidFormat(_))
            ),
            "invalid auxiliary source {index} was accepted"
        );
    }
}

#[test]
fn refuses_generic_mutation_of_inert_3d_and_preserves_detached_source_identity() {
    let drawing = Drawing::from_bytes(package(CONTENT_3D, None)).unwrap();
    let mut edit = drawing.edit();
    assert!(matches!(
        edit.set_shape_transform(0, 2, "rotate (15)"),
        Err(litchi_core::Error::Unsupported(_))
    ));
    assert_eq!(
        edit.commit().unwrap().snapshot().as_bytes(),
        drawing.as_bytes()
    );

    let enhanced = Drawing::from_bytes(package(CONTENT_ENHANCED, None)).unwrap();
    let detached_copy = enhanced.pages()[0].shapes()[0].clone();
    let mut edit = enhanced.edit();
    assert!(matches!(
        edit.add_shape(0, detached_copy),
        Err(litchi_core::Error::Unsupported(_))
    ));

    let transfer = enhanced.prepare_shape_transfer(0, 0).unwrap();
    edit.insert_shape_transfer(0, 1, &transfer).unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(
        commit
            .snapshot()
            .content_xml()
            .matches("<draw:enhanced-geometry")
            .count(),
        2
    );
    let restored = commit
        .patch()
        .durable()
        .unwrap()
        .inverse()
        .apply(commit.snapshot())
        .unwrap();
    assert_eq!(restored.as_bytes(), enhanced.as_bytes());
}

#[test]
fn enforces_frame_and_scene_child_order_and_lexical_domains() {
    let title_before_map = CONTENT_AUXILIARY.replace(
        "<draw:image-map>",
        "<svg:title>late title</svg:title><draw:image-map>",
    );
    let event_before_payload = CONTENT_AUXILIARY.replace(
        "<draw:image xlink:type",
        "<office:event-listeners/><draw:image xlink:type",
    );
    let event_after_payload = CONTENT_AUXILIARY.replace(
        "<draw:image xlink:type=\"simple\" xlink:href=\"media/picture.png\"/>",
        "<draw:image xlink:type=\"simple\" xlink:href=\"media/picture.png\"/><office:event-listeners/>",
    );
    let light_after_shape = CONTENT_3D.replace(
        "<dr3d:light dr3d:direction=\"(0 0 1)\"/><dr3d:cube/>",
        "<dr3d:cube/><dr3d:light dr3d:direction=\"(0 0 1)\"/>",
    );
    let scene_title_after_light = CONTENT_3D.replace(
        "<dr3d:light dr3d:direction=\"(0 0 1)\"/>",
        "<dr3d:light dr3d:direction=\"(0 0 1)\"/><svg:title>late</svg:title>",
    )
    .replace(
        "xmlns:dr3d=\"urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0\"",
        "xmlns:dr3d=\"urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0\" xmlns:svg=\"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0\"",
    );
    let invalid_sources = [
        title_before_map,
        event_before_payload,
        light_after_shape,
        scene_title_after_light,
        CONTENT_AUXILIARY.replace("svg:x=\"0cm\"", "svg:x=\"10\""),
        CONTENT_AUXILIARY.replace("svg:x=\"1cm\"", "svg:x=\"10\""),
        CONTENT_AUXILIARY.replace("svg:width=\"2cm\"", "svg:width=\"2\""),
    ];
    for content in invalid_sources {
        assert!(
            matches!(
                Drawing::from_bytes(package(&content, None)),
                Err(litchi_core::Error::InvalidFormat(_))
            ),
            "invalid frame or scene source was accepted"
        );
    }
    let valid_event = Drawing::from_bytes(package(&event_after_payload, None)).unwrap();
    assert_eq!(
        valid_event.pages()[0].shapes()[0].kind(),
        litchi_odg::shape::ShapeKind::Frame
    );
}

#[test]
fn accepts_empty_anyiri_transition_sound_and_scopes_xml_ids_per_part() {
    let content = CONTENT_NAMED.replace(
        "draw:name=\"Page 1\"",
        "draw:name=\"Page 1\" xml:id=\"sound1\"",
    );
    let source = Drawing::from_bytes(package(&content, Some(STYLES_NAMED))).unwrap();
    let mut transition = Transition::new();
    transition.set_transition_type(Some("automatic")).unwrap();
    transition.set_duration(Some("PT.5S")).unwrap();
    let sound = TransitionSound::new("")
        .unwrap()
        .with_xml_id(Some("sound1"))
        .unwrap();
    transition.set_sound(Some(sound));
    let mut edit = source.edit();
    edit.set_page_transition(0, Some(transition.clone()))
        .unwrap();
    let output = edit.commit().unwrap().into_snapshot();
    assert!(output.styles_xml().unwrap().contains("xlink:href=\"\""));
    assert_eq!(
        Drawing::from_bytes(output.as_bytes().to_vec())
            .unwrap()
            .pages()[0]
            .transition()
            .unwrap()
            .duration(),
        Some("PT.5S")
    );
}

#[test]
fn inserts_detached_page_with_owned_transition_style_and_inverts_exactly() {
    let source = Drawing::from_bytes(package(CONTENT_NAMED, None)).unwrap();
    let mut transition = Transition::new();
    transition.set_transition_type(Some("automatic")).unwrap();
    transition.set_style(Some("dissolve")).unwrap();
    transition.set_duration(Some("PT1.S")).unwrap();
    transition.set_sound(Some(TransitionSound::new("").unwrap()));
    let mut edit = source.edit();
    edit.insert_page(1, Page::new("Inserted").with_transition(transition.clone()))
        .unwrap();
    let commit = edit.commit().unwrap();
    let inserted = &commit.snapshot().pages()[1];
    assert_eq!(inserted.name(), Some("Inserted"));
    assert_eq!(inserted.transition(), Some(&transition));
    assert!(
        inserted
            .style_name()
            .unwrap()
            .starts_with("LitchiPageTransition")
    );
    let reopened = Drawing::from_bytes(commit.snapshot().as_bytes().to_vec()).unwrap();
    assert_eq!(reopened.pages()[1].transition(), Some(&transition));
    let restored = commit
        .patch()
        .durable()
        .unwrap()
        .inverse()
        .apply(commit.snapshot())
        .unwrap();
    assert_eq!(restored.as_bytes(), source.as_bytes());
}
