#![allow(
    clippy::unwrap_used,
    reason = "integration-test assertions panic on failure by design"
)]

use litchi_odp::annotation::{Anchor, Annotation, Position};
use litchi_odp::{Builder, Presentation, Shape, Slide};

fn presentation() -> Presentation {
    let mut shape = Shape::new();
    shape.name = Some("TitleBox".to_string());
    shape.text = "title".to_string();
    let slide = Slide {
        title: None,
        text: "body".to_string(),
        index: 0,
        notes: None,
        transition: None,
        animations: Vec::new(),
        legacy_animation: None,
        shapes: vec![shape],
    };
    let mut builder = Builder::new();
    builder.add_slide_element(slide).unwrap();
    Presentation::from_bytes(builder.build().unwrap()).unwrap()
}

#[test]
fn inventories_and_mutates_page_and_shape_annotations_atomically() {
    let mut presentation = presentation();
    let page = Annotation::new("page comment");
    let shape = Annotation::new("shape comment");

    assert_eq!(
        presentation
            .add_annotation(&Anchor::page(0), &page)
            .unwrap(),
        0
    );
    let shape_anchor = Anchor::shape(0, "TitleBox").unwrap();
    presentation.add_annotation(&shape_anchor, &shape).unwrap();

    let items = presentation.annotations().unwrap();
    assert_eq!(items.len(), 2);
    assert!(items.iter().any(|item| {
        item.annotation.text() == "page comment"
            && item.anchor.position() == &Position::Page { index: 0 }
    }));
    assert!(items.iter().any(|item| {
        item.annotation.text() == "shape comment"
            && item.anchor.position()
                == &Position::Shape {
                    page_index: 0,
                    name: "TitleBox".to_string(),
                }
    }));

    let unchanged = presentation.to_bytes().unwrap();
    let first = items[0].annotation.clone();
    presentation.replace_annotation(0, &first).unwrap();
    assert_eq!(presentation.to_bytes().unwrap(), unchanged);

    presentation
        .replace_annotation(0, &Annotation::new("replacement"))
        .unwrap();
    assert_eq!(
        presentation.annotations().unwrap()[0].annotation.text(),
        "replacement"
    );

    presentation.remove_annotation(1).unwrap();
    assert_eq!(presentation.annotations().unwrap().len(), 1);
}

#[test]
fn invalid_or_duplicate_edits_leave_the_snapshot_unchanged() {
    let mut presentation = presentation();
    let before = presentation.to_bytes().unwrap();
    assert!(
        presentation
            .add_annotation(&Anchor::page(99), &Annotation::new("missing"))
            .is_err()
    );
    assert_eq!(presentation.to_bytes().unwrap(), before);

    let named = {
        let mut value = Annotation::new("named");
        value.set_name(Some("same"));
        value
    };
    presentation
        .add_annotation(&Anchor::page(0), &named)
        .unwrap();
    let before_duplicate = presentation.to_bytes().unwrap();
    assert!(
        presentation
            .add_annotation(&Anchor::shape(0, "TitleBox").unwrap(), &named)
            .is_err()
    );
    assert_eq!(presentation.to_bytes().unwrap(), before_duplicate);
    assert_eq!(
        presentation
            .find_annotation("same")
            .unwrap()
            .unwrap()
            .annotation
            .text(),
        "named"
    );
    assert!(presentation.find_annotation("").is_err());
}
