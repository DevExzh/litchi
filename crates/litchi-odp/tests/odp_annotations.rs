#![allow(
    clippy::unwrap_used,
    reason = "integration-test assertions panic on failure by design"
)]

use litchi_odp::annotation::{Anchor, Annotation, Position};
use litchi_odp::{Builder, Package, Shape, Slide};

fn presentation() -> Package {
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
    Package::from_bytes(&builder.build().unwrap().to_bytes().unwrap()).unwrap()
}

/// Apply one staged operation through the primary edit and adopt the commit.
fn edit_with<T>(
    package: &mut Package,
    operation: impl FnOnce(&mut litchi_odp::edit::Edit) -> Result<T, litchi_odp::Error>,
) -> Result<T, litchi_odp::Error> {
    let mut edit = package.edit()?;
    let value = operation(&mut edit)?;
    *package = edit.commit()?.into_package();
    Ok(value)
}

#[test]
fn inventories_and_mutates_page_and_shape_annotations_atomically() {
    let mut presentation = presentation();
    let page = Annotation::new("page comment");
    let shape = Annotation::new("shape comment");

    assert_eq!(
        edit_with(&mut presentation, |edit| edit
            .add_annotation(&Anchor::page(0), &page))
        .unwrap(),
        0
    );
    let shape_anchor = Anchor::shape(0, "TitleBox").unwrap();
    edit_with(&mut presentation, |edit| {
        edit.add_annotation(&shape_anchor, &shape)
    })
    .unwrap();

    let items = presentation.presentation().unwrap().annotations().unwrap();
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
    edit_with(&mut presentation, |edit| edit.replace_annotation(0, &first)).unwrap();
    assert_eq!(presentation.to_bytes().unwrap(), unchanged);

    edit_with(&mut presentation, |edit| {
        edit.replace_annotation(0, &Annotation::new("replacement"))
    })
    .unwrap();
    assert_eq!(
        presentation.presentation().unwrap().annotations().unwrap()[0]
            .annotation
            .text(),
        "replacement"
    );

    edit_with(&mut presentation, |edit| edit.remove_annotation(1)).unwrap();
    assert_eq!(
        presentation
            .presentation()
            .unwrap()
            .annotations()
            .unwrap()
            .len(),
        1
    );
}

#[test]
fn invalid_or_duplicate_edits_leave_the_snapshot_unchanged() {
    let mut presentation = presentation();
    let before = presentation.to_bytes().unwrap();
    assert!(
        edit_with(&mut presentation, |edit| edit
            .add_annotation(&Anchor::page(99), &Annotation::new("missing")))
        .is_err()
    );
    assert_eq!(presentation.to_bytes().unwrap(), before);

    let named = {
        let mut value = Annotation::new("named");
        value.set_name(Some("same"));
        value
    };
    edit_with(&mut presentation, |edit| {
        edit.add_annotation(&Anchor::page(0), &named)
    })
    .unwrap();
    let before_duplicate = presentation.to_bytes().unwrap();
    assert!(
        edit_with(&mut presentation, |edit| edit
            .add_annotation(&Anchor::shape(0, "TitleBox").unwrap(), &named))
        .is_err()
    );
    assert_eq!(presentation.to_bytes().unwrap(), before_duplicate);
    assert_eq!(
        presentation
            .presentation()
            .unwrap()
            .find_annotation("same")
            .unwrap()
            .unwrap()
            .annotation
            .text(),
        "named"
    );
    assert!(
        presentation
            .presentation()
            .unwrap()
            .find_annotation("")
            .is_err()
    );
}
