#![allow(clippy::unwrap_used, clippy::expect_used, reason = "test assertions")]

use litchi_core::Position;
use litchi_pptx::{Error, Package, builder::Builder, opened::Patch, slide::Selector};

fn package() -> Package {
    let mut builder = Builder::new().unwrap();
    for title in ["first", "second", "third", "fourth"] {
        builder
            .presentation()
            .unwrap()
            .add_slide()
            .unwrap()
            .set_title(title);
    }
    builder.build().unwrap()
}

fn names(package: &Package) -> Vec<String> {
    package
        .presentation()
        .unwrap()
        .slides()
        .unwrap()
        .iter()
        .map(|slide| slide.name().unwrap())
        .collect()
}

#[test]
fn semantic_order_edits_compose_replay_invert_and_preserve_the_source() {
    let original = package();
    let source = original.to_bytes().unwrap();
    let original_names = names(&original);
    let mut edit = original.edit().unwrap();
    assert!(
        edit.move_slide(original_names[0].clone(), Position::new(3))
            .unwrap()
    );
    assert!(
        edit.swap_slides(original_names[1].as_str(), Position::new(2))
            .unwrap()
    );
    // Staged order is fourth, third, second, first. Every position in this
    // permutation resolves against that same staged order.
    assert!(
        edit.reorder_slides(&[
            Selector::from(Position::new(3)),
            Selector::from(original_names[2].clone()),
            Selector::from(Position::new(2)),
            Selector::from(original_names[3].as_str()),
        ])
        .unwrap()
    );
    let (changed, patch) = edit.commit().unwrap().into_parts();
    assert_eq!(
        names(&changed),
        [
            original_names[0].clone(),
            original_names[2].clone(),
            original_names[1].clone(),
            original_names[3].clone()
        ]
    );
    assert_eq!(original.to_bytes().unwrap(), source);
    let wire = Patch::from_bytes(
        &patch.to_bytes().unwrap(),
        litchi_pptx::opened::Limits::default(),
    )
    .unwrap();
    let replayed = original.apply(&wire).unwrap().into_package();
    assert_eq!(replayed.to_bytes().unwrap(), changed.to_bytes().unwrap());
    let restored = replayed.apply(&wire.inverse()).unwrap().into_package();
    assert_eq!(restored.to_bytes().unwrap(), source);
    let reopened = Package::from_reader(std::io::Cursor::new(changed.to_bytes().unwrap())).unwrap();
    assert_eq!(names(&reopened), names(&changed));
}

#[test]
fn failed_and_duplicate_selectors_leave_staged_order_usable() {
    let original = package();
    let original_names = names(&original);
    let mut edit = original.edit().unwrap();
    assert!(matches!(
        edit.move_slide("missing", Position::new(0)),
        Err(Error::SlideNameNotFound(_))
    ));
    assert!(matches!(
        edit.move_slide(Position::new(0), Position::new(4)),
        Err(Error::SlideIndexOutOfBounds { index: 4, len: 4 })
    ));
    assert!(
        edit.swap_slides(Position::new(0), Position::new(4))
            .is_err()
    );
    assert!(
        edit.reorder_slides(&[0usize.into(), 0usize.into(), 2usize.into(), 3usize.into()])
            .is_err()
    );
    assert!(edit.reorder_slides(&[0usize.into()]).is_err());
    assert!(
        edit.swap_slides(original_names[0].as_str(), original_names[3].as_str())
            .unwrap()
    );
    let changed = edit.commit().unwrap().into_package();
    assert_eq!(
        names(&changed),
        [
            original_names[3].clone(),
            original_names[1].clone(),
            original_names[2].clone(),
            original_names[0].clone()
        ]
    );
}

#[test]
fn semantic_identity_noops_keep_the_exact_publication() {
    let original = package();
    let source = original.to_bytes().unwrap();
    let original_names = names(&original);
    let mut edit = original.edit().unwrap();
    assert!(
        !edit
            .move_slide(original_names[1].as_str(), Position::new(1))
            .unwrap()
    );
    assert!(
        !edit
            .swap_slides(original_names[2].clone(), Position::new(2))
            .unwrap()
    );
    let selectors: Vec<_> = original_names.iter().map(Selector::from).collect();
    assert!(!edit.reorder_slides(&selectors).unwrap());
    let (unchanged, patch) = edit.commit().unwrap().into_parts();
    assert!(patch.is_noop());
    assert_eq!(unchanged.to_bytes().unwrap(), source);
    assert!(
        original
            .presentation()
            .unwrap()
            .find_slide(Selector::from(original_names[0].clone()))
            .unwrap()
            .is_some()
    );
}
