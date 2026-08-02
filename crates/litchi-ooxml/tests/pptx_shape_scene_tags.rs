use litchi_ooxml::pptx::Package;
use litchi_pptx::shape::LookupError;
use litchi_pptx::tag::{List, Tag};
use tempfile::NamedTempFile;

const SHAPE_TAGS: &str = concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/libreoffice-core/sd/qa/unit/data/pptx/tdf103477.pptx"
);

#[test]
fn real_libreoffice_scene_keeps_nested_shapes_and_semantic_tag_anchors() {
    let package = Package::open(SHAPE_TAGS).unwrap();
    let presentation = package.presentation().unwrap();
    let slides = presentation.slides().unwrap();
    let slide = &slides[0];
    let scene = slide.shapes().unwrap();

    // The raw owner has one OLE fallback picture nested inside a graphic
    // frame. It is not a user shape, while the nested group children are.
    assert_eq!(scene.len(), 21);
    assert_eq!(scene.at(0).unwrap().name(), Some("Objekt 2"));
    assert!(matches!(
        scene.get("Rectangle 16"),
        Err(LookupError::AmbiguousName { matches: 2, .. })
    ));

    let tagged = (0..scene.len())
        .filter(|index| slide.shape_tags(*index).unwrap().is_some())
        .collect::<Vec<_>>();
    assert_eq!(tagged, [0, 6, 7, 10, 13, 16, 18]);
}

#[test]
fn facade_replaces_and_removes_a_real_shape_tag_list() {
    let mut package = Package::open(SHAPE_TAGS).unwrap();
    let mut replacement = List::new();
    replacement
        .add(Tag::new("Reviewer", "Ada").unwrap())
        .unwrap();

    let old = package
        .put_shape_tags(0_usize, "Objekt 2", replacement)
        .unwrap()
        .unwrap();
    assert!(!old.is_empty());
    assert_eq!(
        package
            .shape_tags(0_usize, "Objekt 2")
            .unwrap()
            .unwrap()
            .get("reviewer")
            .unwrap()
            .value(),
        "Ada"
    );

    let output = NamedTempFile::with_suffix(".pptx").unwrap();
    package.save(output.path()).unwrap();
    let mut reopened = Package::open(output.path()).unwrap();
    assert_eq!(
        reopened
            .shape_tags(0_usize, "Objekt 2")
            .unwrap()
            .unwrap()
            .get("Reviewer")
            .unwrap()
            .value(),
        "Ada"
    );

    let removed = reopened
        .remove_shape_tags(0_usize, "Objekt 2")
        .unwrap()
        .unwrap();
    assert_eq!(removed.get("reviewer").unwrap().value(), "Ada");
    assert!(reopened.shape_tags(0_usize, "Objekt 2").unwrap().is_none());
}
