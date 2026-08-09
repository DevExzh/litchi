#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_pptx::shape::{LookupError, Scene};

const LOCAL_SHAPE_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/shapes/shape_identity.xml");

#[test]
fn slide_shapes_use_names_first_with_checked_numeric_fallback() {
    let scene = Scene::read(LOCAL_SHAPE_XML).unwrap();
    assert_eq!(
        scene
            .placeholders()
            .filter_map(|shape| shape
                .placeholder()
                .and_then(litchi_pptx::shape::Placeholder::kind))
            .collect::<Vec<_>>(),
        ["title", "body"]
    );

    assert_eq!(scene.placeholders().count(), 2);

    let title = scene.shape("Title").unwrap();
    assert_eq!(title.name(), Some("Title"));
    assert_eq!(title.id(), Some(7));
    let title_placeholder = title.placeholder().unwrap();
    assert_eq!(title_placeholder.kind(), Some("title"));
    assert_eq!(title_placeholder.index(), 0);

    let body = scene.shape(1_usize).unwrap();
    assert_eq!(body.name(), Some("Body"));
    assert_eq!(body.id(), Some(11));
    let body_placeholder = body.placeholder().unwrap();
    assert_eq!(body_placeholder.kind(), Some("body"));
    assert_eq!(body_placeholder.index(), 3);
    assert!(matches!(scene.get("Missing"), Ok(None)));
}

#[test]
fn duplicate_semantic_shape_names_are_typed_errors() {
    let xml = std::str::from_utf8(LOCAL_SHAPE_XML)
        .unwrap()
        .replace("name=\"Body\"", "name=\"Title\"");
    let scene = Scene::read(xml.as_bytes()).unwrap();

    assert!(matches!(
        scene.shape("Title"),
        Err(LookupError::AmbiguousName { name, matches: 2 })
            if name == "Title"
    ));
}
