#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "blank-slide insertion tests use panic-driven assertions"
)]

//! Blank-slide insertion through the canonical primary edit (record 0260).

use litchi_ppt::package::edit::Position;
use litchi_ppt::{Builder, Package};
use std::path::PathBuf;

fn fixture(path: &str) -> Vec<u8> {
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data")
        .join(path);
    std::fs::read(path).expect("PPT fixture")
}

/// Every slide's shape texts.
fn texts(package: &Package) -> Vec<Vec<String>> {
    let presentation = package.presentation().unwrap();
    presentation
        .slides()
        .unwrap()
        .iter()
        .map(|slide| {
            slide
                .shapes()
                .unwrap()
                .iter()
                .map(|shape| shape.text().unwrap())
                .collect()
        })
        .collect()
}

const DECKS: &[&str] = &[
    "ole/ppt/SampleShow.ppt",
    "ole/ppt/text_shapes.ppt",
    "poi/test-data/slideshow/basic_test_ppt_file.ppt",
    "poi/test-data/slideshow/41246-1.ppt",
    "poi/test-data/slideshow/45543.ppt",
];

#[test]
fn blank_slides_insert_at_every_position_of_powerpoint_decks() {
    for deck in DECKS {
        let source_bytes = fixture(deck);
        let source = Package::from_bytes(&source_bytes).unwrap();
        let before = texts(&source);
        for position in [0, before.len() / 2, before.len()] {
            let mut edit = source.edit().unwrap();
            edit.insert_slide(Position::new(position)).unwrap();
            let commit = edit.commit().unwrap();
            assert!(commit.diagnostics().changed());

            let mut expected = before.clone();
            expected.insert(position, Vec::new());
            assert_eq!(texts(commit.package()), expected, "{deck} at {position}");
            let published = commit.package().to_bytes().unwrap();
            assert_eq!(
                texts(&Package::from_bytes(&published).unwrap()),
                expected,
                "{deck} reopened"
            );
            assert_eq!(source.to_bytes().unwrap(), source_bytes);
            assert_eq!(
                source
                    .apply(commit.patch())
                    .unwrap()
                    .into_package()
                    .to_bytes()
                    .unwrap(),
                published
            );
            assert_eq!(
                commit
                    .package()
                    .apply(&commit.patch().inverse())
                    .unwrap()
                    .into_package()
                    .to_bytes()
                    .unwrap(),
                source_bytes
            );
        }
    }
}

#[test]
fn several_inserted_slides_receive_distinct_drawings() {
    let source = Package::from_bytes(&fixture("ole/ppt/SampleShow.ppt")).unwrap();
    let count = texts(&source).len();
    let mut edit = source.edit().unwrap();
    edit.insert_slide(Position::new(0)).unwrap();
    edit.insert_slide(Position::new(count + 1)).unwrap();
    edit.insert_slide(Position::new(1)).unwrap();
    let commit = edit.commit().unwrap();
    let presentation = commit.package().presentation().unwrap();
    let slides = presentation.slides().unwrap();
    assert_eq!(slides.len(), count + 3);
    let mut ids = slides
        .iter()
        .map(|slide| (slide.slide_id(), slide.persist_id()))
        .collect::<Vec<_>>();
    ids.sort_unstable();
    ids.dedup();
    assert_eq!(ids.len(), count + 3);
    // A later edit can remove an inserted slide and edit the others.
    let mut edit = commit.package().edit().unwrap();
    edit.remove_slide(Position::new(0)).unwrap();
    edit.set_shape_text(
        litchi_ppt::package::edit::Target::new(Position::new(1), Position::new(0)),
        "Still editable",
    )
    .unwrap();
    let after = edit.commit().unwrap();
    assert_eq!(texts(after.package())[1][0], "Still editable");
}

#[test]
fn insertion_composes_with_moves_and_text_in_one_edit() {
    let source =
        Package::from_bytes(&fixture("poi/test-data/slideshow/basic_test_ppt_file.ppt")).unwrap();
    let before = texts(&source);
    let mut edit = source.edit().unwrap();
    edit.set_shape_text(
        litchi_ppt::package::edit::Target::new(Position::new(0), Position::new(0)),
        "Retitled",
    )
    .unwrap();
    edit.insert_slide(Position::new(1)).unwrap();
    edit.move_slide(Position::new(1), Position::new(0)).unwrap();
    let commit = edit.commit().unwrap();
    let mut expected = before.clone();
    expected[0][0] = "Retitled".to_owned();
    expected.insert(0, Vec::new());
    assert_eq!(texts(commit.package()), expected);
}

#[test]
fn builder_and_empty_decks_accept_blank_slides() {
    let mut builder = Builder::new();
    let slide = builder.add_slide().unwrap();
    builder.add_textbox(slide, 10, 10, 200, 40, "only").unwrap();
    let package = builder.build().unwrap();
    let mut edit = package.edit().unwrap();
    edit.insert_slide(Position::new(1)).unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(
        texts(commit.package()),
        [vec!["only".to_owned()], Vec::new()]
    );

    let empty = Package::from_bytes(&fixture("ole/ppt/empty.ppt")).unwrap();
    let count = texts(&empty).len();
    let mut edit = empty.edit().unwrap();
    edit.insert_slide(Position::new(count)).unwrap();
    assert_eq!(texts(edit.commit().unwrap().package()).len(), count + 1);
}

#[test]
fn out_of_range_and_signed_insertions_are_refused() {
    let source = Package::from_bytes(&fixture("ole/ppt/SampleShow.ppt")).unwrap();
    let count = texts(&source).len();
    let mut edit = source.edit().unwrap();
    assert!(edit.insert_slide(Position::new(count + 1)).is_err());
    assert!(!edit.commit().unwrap().diagnostics().changed());
}
