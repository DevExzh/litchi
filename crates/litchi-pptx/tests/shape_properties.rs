#![allow(clippy::unwrap_used, clippy::expect_used, reason = "test assertions")]
use litchi_core::Position;
use litchi_pptx::shape::{Scene, Shape, owned, raw};
use litchi_pptx::{Package, ShapePropertiesReplacement, builder::Builder, opened::Patch};
const P: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const C: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const D: &str = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
fn tree(body: &str) -> String {
    format!(
        r#"<p:spTree xmlns:p="{P}" xmlns:a="{A}" xmlns:r="{R}" xmlns:s="{R}" xmlns:c="{C}" xmlns:d="{D}" xmlns:f="urn:foreign">{body}</p:spTree>"#
    )
}
fn picture(attributes: &str) -> String {
    format!(
        r#"<p:pic><p:nvPicPr><p:cNvPr id="7" name="private-name"/></p:nvPicPr><p:blipFill><a:blip {attributes}/></p:blipFill></p:pic>"#
    )
}
fn frame(body: &str) -> String {
    format!(
        r#"<p:graphicFrame><a:graphic><a:graphicData>{body}</a:graphicData></a:graphic></p:graphicFrame>"#
    )
}

fn authored_package() -> Package {
    let mut builder = Builder::new().unwrap();
    let slide = builder.presentation().unwrap().add_slide().unwrap();
    slide.add_text_box("first", 10, 20, 300, 400);
    slide.add_text_box("second", 50, 60, 500, 600);
    builder.build().unwrap()
}

fn shape_values(package: &Package) -> Vec<(Option<String>, Option<(i64, i64, i64, i64)>)> {
    package
        .presentation()
        .unwrap()
        .slide(0)
        .unwrap()
        .unwrap()
        .shapes()
        .unwrap()
        .iter()
        .map(|shape| {
            let bounds = shape
                .bounds()
                .map(|value| (value.x(), value.y(), value.width(), value.height()));
            (shape.text().map(ToOwned::to_owned), bounds)
        })
        .collect()
}
#[test]
fn non_visual_properties_decode_lazily_and_debug_is_content_free() {
    let xml = tree(
        r#"<p:sp><p:nvSpPr><p:cNvPr id="4567" name="private-name" descr="private-description &amp; more" title="private-title&#10;line" hidden="true"/></p:nvSpPr></p:sp>"#,
    );
    let scene = Scene::read(xml.as_bytes(), litchi_pptx::shape::Limits::DEFAULT).unwrap();
    let shape = scene.at(0).unwrap();
    let properties = shape.non_visual().unwrap();
    assert_eq!(properties.description(), Some("private-description & more"));
    assert_eq!(properties.title(), Some("private-title\nline"));
    assert_eq!(properties.hidden(), Some(true));
    assert_eq!(shape.hidden().unwrap(), Some(true));
    let debug = format!(
        "{shape:?} {:?} {properties:?} {scene:?} {:?} {:?}",
        shape.common(),
        scene.iter(),
        litchi_pptx::shape::Key::Name("private-name")
    );
    for secret in [
        "private-name",
        "private-description",
        "private-title",
        "4567",
        "<p:",
    ] {
        assert!(!debug.contains(secret), "{debug}");
    }
    let span = raw::non_visual_span(shape).unwrap();
    assert!(xml[span.start() as usize..span.end().unwrap() as usize].starts_with("<p:cNvPr"));
}
#[test]
fn optional_flags_and_malformed_values_remain_distinct() {
    for (attribute, expected) in [
        ("", None),
        ("hidden=\"0\"", Some(false)),
        ("hidden=\"1\"", Some(true)),
        ("hidden=\"false\"", Some(false)),
    ] {
        let xml = tree(&format!(
            r#"<p:sp><p:nvSpPr><p:cNvPr {attribute}/></p:nvSpPr></p:sp>"#
        ));
        let scene = Scene::read(xml.as_bytes(), litchi_pptx::shape::Limits::DEFAULT).unwrap();
        assert_eq!(scene.at(0).unwrap().hidden().unwrap(), expected);
        assert!(scene.at(0).unwrap().description().unwrap().is_none());
    }
    let xml = tree(
        r#"<p:sp><p:nvSpPr><p:cNvPr descr="plain" title="a&#9;b" hidden="private-invalid"/></p:nvSpPr></p:sp>"#,
    );
    let scene = Scene::read(xml.as_bytes(), litchi_pptx::shape::Limits::DEFAULT).unwrap();
    let error = scene.at(0).unwrap().non_visual().unwrap_err();
    assert!(!format!("{error:?} {error}").contains("private-invalid"));
}
#[test]
fn transforms_preserve_rotation_flip_and_group_child_coordinates() {
    let xml = tree(
        r#"<p:grpSp><p:nvGrpSpPr/><p:grpSpPr><a:xfrm rot="60000" flipH="1"><a:off x="1" y="2"/><a:ext cx="3" cy="4"/><a:chOff x="5" y="6"/><a:chExt cx="7" cy="8"/></a:xfrm></p:grpSpPr><p:sp><p:spPr><a:xfrm flipV="true"/></p:spPr></p:sp></p:grpSp><p:graphicFrame><p:xfrm><a:off x="9" y="10"/><a:ext cx="11" cy="12"/></p:xfrm></p:graphicFrame>"#,
    );
    let scene = Scene::read(xml.as_bytes(), litchi_pptx::shape::Limits::DEFAULT).unwrap();
    let transform = scene.at(0).unwrap().transform().unwrap().unwrap();
    assert_eq!(transform.rotation().value(), 60000);
    assert!(transform.flip_horizontal());
    assert!(transform.child_offset().is_some());
    assert!(transform.child_extent().is_some());
    assert!(
        scene
            .at(1)
            .unwrap()
            .transform()
            .unwrap()
            .unwrap()
            .flip_vertical()
    );
    assert!(
        scene
            .at(2)
            .unwrap()
            .transform()
            .unwrap()
            .unwrap()
            .offset()
            .is_some()
    );
}
#[test]
fn transform_decoding_refuses_namespace_spoofs_duplicates_and_bad_scalars() {
    for body in [
        r#"<a:xfrm><f:off x="1" y="2"/></a:xfrm>"#,
        r#"<a:xfrm/><a:xfrm/>"#,
        r#"<a:xfrm rot="private-invalid"/>"#,
        r#"<a:xfrm flipH="private-invalid"/>"#,
        r#"<a:xfrm private-invalid="1"/>"#,
        r#"<a:xfrm><a:private-invalid/></a:xfrm>"#,
        r#"<a:xfrm rot="&private-invalid;"/>"#,
        r#"<a:xfrm><a:off x="private-invalid" y="0"/></a:xfrm>"#,
    ] {
        let xml = tree(&format!("<p:sp><p:spPr>{body}</p:spPr></p:sp>"));
        let scene = Scene::read(xml.as_bytes(), litchi_pptx::shape::Limits::DEFAULT).unwrap();
        let error = scene.at(0).unwrap().transform().unwrap_err();
        assert!(!format!("{error:?} {error}").contains("private-invalid"));
    }
    let xml = tree(r#"<p:sp><p:extLst><a:xfrm rot="60000"/></p:extLst></p:sp>"#);
    let scene = Scene::read(xml.as_bytes(), litchi_pptx::shape::Limits::DEFAULT).unwrap();
    assert!(scene.at(0).unwrap().transform().unwrap().is_none());
}
#[test]
fn payload_relationships_use_explicit_raw_access() {
    for (attrs, kind) in [
        (
            r#"r:embed="private&amp;payload""#,
            raw::PayloadAttribute::Embed,
        ),
        (
            r#"r:link="private&amp;payload""#,
            raw::PayloadAttribute::Link,
        ),
    ] {
        let xml = tree(&picture(attrs));
        let scene = Scene::read(xml.as_bytes(), litchi_pptx::shape::Limits::DEFAULT).unwrap();
        let payload = raw::payload(scene.at(0).unwrap()).unwrap().unwrap();
        assert_eq!(payload.attribute(), kind);
        assert_eq!(payload.relationship_id().unwrap(), "private&payload");
        assert!(!format!("{payload:?}").contains("private"));
    }
}
#[test]
fn payload_ambiguity_and_namespace_mismatch_are_typed() {
    for attributes in [
        r#"r:embed="one" r:link="two""#,
        r#"r:embed="one" s:embed="two""#,
    ] {
        let xml = tree(&picture(attributes));
        let scene = Scene::read(xml.as_bytes(), litchi_pptx::shape::Limits::DEFAULT).unwrap();
        assert_eq!(
            raw::payload(scene.at(0).unwrap()).unwrap_err(),
            raw::Error::Ambiguous
        );
    }
    for namespace in [
        "",
        "http://purl.oclc.org/ooxml/officeDocument/relationships",
    ] {
        let xml = tree(&picture(r#"r:embed="private""#));
        let xml = if namespace.is_empty() {
            xml.replace(&format!(r#" xmlns:r="{R}""#), "")
        } else {
            xml.replace(
                &format!(r#"xmlns:r="{R}""#),
                &format!(r#"xmlns:r="{namespace}""#),
            )
        };
        let scene = Scene::read(xml.as_bytes(), litchi_pptx::shape::Limits::DEFAULT).unwrap();
        assert_eq!(
            raw::payload(scene.at(0).unwrap()).unwrap_err(),
            raw::Error::InvalidNamespace
        );
    }
    let xml = tree(&picture(r#"f:embed="private""#));
    let scene = Scene::read(xml.as_bytes(), litchi_pptx::shape::Limits::DEFAULT).unwrap();
    assert!(raw::payload(scene.at(0).unwrap()).unwrap().is_none());
}
#[test]
fn frame_payloads_are_direct_and_conflicts_are_not_selected() {
    for payload in [
        r#"<c:chart r:id="part"/>"#,
        r#"<d:relIds r:dm="part"/>"#,
        r#"<p:oleObj r:id="part"/>"#,
    ] {
        let xml = tree(&frame(payload));
        let scene = Scene::read(xml.as_bytes(), litchi_pptx::shape::Limits::DEFAULT).unwrap();
        assert_eq!(
            raw::payload(scene.at(0).unwrap())
                .unwrap()
                .unwrap()
                .relationship_id()
                .unwrap(),
            "part"
        );
    }
    for payloads in [
        r#"<c:chart r:id="one"/><p:oleObj r:id="two"/>"#,
        r#"<a:tbl/><c:chart r:id="one"/>"#,
        r#"<c:chart r:id="one"/><a:tbl/>"#,
    ] {
        let xml = tree(&frame(payloads));
        let scene = Scene::read(xml.as_bytes(), litchi_pptx::shape::Limits::DEFAULT).unwrap();
        assert!(matches!(scene.at(0).unwrap(), Shape::Frame(_)));
        assert_eq!(
            raw::payload(scene.at(0).unwrap()).unwrap_err(),
            raw::Error::Ambiguous
        );
    }
    let xml = tree(&frame(
        r#"<f:extension><c:chart r:id="opaque"/></f:extension>"#,
    ));
    let scene = Scene::read(xml.as_bytes(), litchi_pptx::shape::Limits::DEFAULT).unwrap();
    assert!(matches!(scene.at(0).unwrap(), Shape::Frame(_)));
    assert!(raw::payload(scene.at(0).unwrap()).unwrap().is_none());
}
#[test]
fn source_spans_survive_bom_mce_and_retention() {
    let body = r#"<p:sp><p:nvSpPr><p:cNvPr descr="kept"/></p:nvSpPr><p:spPr><a:xfrm rot="60000"/></p:spPr></p:sp>"#;
    let xml = tree(&format!(
        r#"<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:unknown"><mc:Choice Requires="x"/><mc:Fallback>{body}</mc:Fallback></mc:AlternateContent>"#
    ));
    for xml in [tree(body), xml] {
        let bytes = [vec![0xef, 0xbb, 0xbf], xml.into_bytes()].concat();
        let scene = owned::Scene::read(bytes, litchi_pptx::shape::Limits::DEFAULT).unwrap();
        let handle = scene.shape(0usize).unwrap().unwrap();
        let span = raw::transform_span(handle.view()).unwrap();
        assert!(
            owned::raw::xml(&scene)[span.start() as usize..span.end().unwrap() as usize]
                .starts_with(b"<a:xfrm")
        );
        drop(scene);
        assert_eq!(
            handle.view().description().unwrap().as_deref(),
            Some("kept")
        );
        assert_eq!(
            handle
                .view()
                .transform()
                .unwrap()
                .unwrap()
                .rotation()
                .value(),
            60000
        );
    }
}

#[test]
fn eager_shape_property_batch_composes_text_and_bounds_and_replays() {
    let original = authored_package();
    let source_bytes = original.to_bytes().unwrap();
    let mut edit = original.edit().unwrap();
    let replacements = [
        ShapePropertiesReplacement::new(Position::new(0))
            .with_text("first changed")
            .with_bounds((100, 200, 700, 800)),
        ShapePropertiesReplacement::new(Position::new(1)).with_bounds((300, 400, 900, 1_000)),
    ];
    assert_eq!(
        edit.set_shape_properties(Position::new(0), &replacements)
            .unwrap(),
        2
    );
    let changed = edit.commit().unwrap();
    let patch = Patch::from_bytes(
        &changed.patch().to_bytes().unwrap(),
        litchi_pptx::opened::Limits::default(),
    )
    .unwrap();
    let changed_package = changed.clone().into_package();
    assert_eq!(shape_values(&original)[0].0.as_deref(), Some("first"));
    assert_eq!(
        shape_values(&changed_package)[0].0.as_deref(),
        Some("first changed")
    );
    assert_eq!(
        shape_values(&changed_package)[0].1,
        Some((100, 200, 700, 800))
    );
    assert_eq!(
        shape_values(&changed_package)[1].1,
        Some((300, 400, 900, 1_000))
    );
    assert_eq!(original.to_bytes().unwrap(), source_bytes);

    let replayed = original.apply(&patch).unwrap().into_package();
    let changed_bytes = changed_package.to_bytes().unwrap();
    assert_eq!(replayed.to_bytes().unwrap(), changed_bytes);
    let restored = replayed.apply(&patch.inverse()).unwrap().into_package();
    assert_eq!(restored.to_bytes().unwrap(), source_bytes);
}

#[test]
fn shape_property_noop_is_exact_and_invalid_batches_are_atomic() {
    let original = authored_package();
    let source_bytes = original.to_bytes().unwrap();
    let current = shape_values(&original);
    let mut noop = original.edit().unwrap();
    let replacements = [
        ShapePropertiesReplacement::new(Position::new(0))
            .with_text(current[0].0.as_deref().unwrap())
            .with_bounds(current[0].1.unwrap()),
        ShapePropertiesReplacement::new(Position::new(1)).with_bounds(current[1].1.unwrap()),
    ];
    assert_eq!(noop.set_shape_properties(0, &replacements).unwrap(), 0);
    let noop_commit = noop.commit().unwrap();
    assert!(noop_commit.patch().is_noop());
    let noop_package = noop_commit.clone().into_package();
    assert_eq!(noop_package.to_bytes().unwrap(), source_bytes);

    let mut edit = original.edit().unwrap();
    assert!(edit.set_shape_text(0, 0, "staged text").unwrap());
    let invalid = [
        ShapePropertiesReplacement::new(Position::new(0)).with_bounds((1, 2, 0, 4)),
        ShapePropertiesReplacement::new(Position::new(0)).with_text("duplicate"),
    ];
    assert!(edit.set_shape_properties(0, &invalid).is_err());
    let committed = edit.commit().unwrap().into_package();
    assert_eq!(
        shape_values(&committed)[0].0.as_deref(),
        Some("staged text")
    );
    assert_eq!(shape_values(&committed)[0].1, current[0].1);
}

#[test]
fn strict_payload_namespaces_and_content_parts_are_checked() {
    let xml = tree(&picture(r#"r:embed="part""#))
        .replace(P, "http://purl.oclc.org/ooxml/presentationml/main")
        .replace(A, "http://purl.oclc.org/ooxml/drawingml/main");
    let scene = Scene::read(xml.as_bytes(), litchi_pptx::shape::Limits::DEFAULT).unwrap();
    assert_eq!(
        raw::payload(scene.at(0).unwrap()).unwrap_err(),
        raw::Error::InvalidNamespace
    );
    let xml = xml.replace(R, "http://purl.oclc.org/ooxml/officeDocument/relationships");
    let scene = Scene::read(xml.as_bytes(), litchi_pptx::shape::Limits::DEFAULT).unwrap();
    assert_eq!(
        raw::payload(scene.at(0).unwrap())
            .unwrap()
            .unwrap()
            .relationship_id()
            .unwrap(),
        "part"
    );

    let xml = tree(
        r#"<p:contentPart r:id="content"/><p:pic><p:extLst><a:blip r:embed="forged"/></p:extLst></p:pic>"#,
    );
    let scene = Scene::read(xml.as_bytes(), litchi_pptx::shape::Limits::DEFAULT).unwrap();
    assert_eq!(
        raw::payload(scene.at(0).unwrap())
            .unwrap()
            .unwrap()
            .relationship_id()
            .unwrap(),
        "content"
    );
    assert!(raw::payload(scene.at(1).unwrap()).unwrap().is_none());
}

#[test]
fn borrowed_lookup_errors_omit_authored_names() {
    let xml = tree(&(picture("") + &picture("")));
    let scene = Scene::read(xml.as_bytes(), litchi_pptx::shape::Limits::DEFAULT).unwrap();
    for key in ["private-name", "private-missing"] {
        let error = scene.shape(key).unwrap_err();
        assert!(!format!("{error:?} {error}").contains("private"));
    }
}
