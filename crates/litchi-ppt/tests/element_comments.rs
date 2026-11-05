#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "comment edit tests use panic-driven assertions"
)]

//! Slide comments through the canonical primary edit (record 0323).

use litchi_ppt::comments::{CommentChange, DateTime, NewComment};
use litchi_ppt::package::edit::Position;
use litchi_ppt::{Package, ParsedComment};
use std::path::PathBuf;

fn fixture(path: &str) -> Vec<u8> {
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data")
        .join(path);
    std::fs::read(path).expect("PPT fixture")
}

const WITH_COMMENTS: &str = "poi/test-data/slideshow/WithComments.ppt";
const NO_COMMENTS: &str = "poi/test-data/slideshow/basic_test_ppt_file.ppt";

fn comments(package: &Package, slide: usize) -> Vec<ParsedComment> {
    package
        .presentation()
        .unwrap()
        .slide_at(slide)
        .unwrap()
        .unwrap()
        .comments()
        .unwrap()
}

fn texts(package: &Package) -> Vec<String> {
    package
        .presentation()
        .unwrap()
        .slides()
        .unwrap()
        .iter()
        .map(|slide| slide.text().unwrap().to_owned())
        .collect()
}

fn new_comment(text: &str) -> NewComment {
    NewComment {
        author: "Reviewer".into(),
        initials: "R".into(),
        text: text.into(),
        date: DateTime::parse("2026-10-17T09:30:00.250"),
        x: 120,
        y: 48,
    }
}

#[test]
fn add_creates_the_ppt10_tag_and_round_trips() {
    let source_bytes = fixture(NO_COMMENTS);
    let source = Package::from_bytes(&source_bytes).unwrap();
    let before = texts(&source);
    let mut edit = source.edit().unwrap();
    let position = edit
        .add_comment(Position::new(0), &new_comment("First note"))
        .unwrap();
    assert_eq!(position.get(), 0);
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    let published = Package::from_bytes(&commit.package().to_bytes().unwrap()).unwrap();
    let added = comments(&published, 0);
    assert_eq!(added.len(), 1);
    assert_eq!(added[0].author, "Reviewer");
    assert_eq!(added[0].initials, "R");
    assert_eq!(added[0].text, "First note");
    assert_eq!(added[0].index, 1);
    assert_eq!((added[0].x, added[0].y), (120, 48));
    assert_eq!(
        (added[0].year, added[0].month, added[0].day, added[0].hour),
        (2026, 10, 17, 9)
    );
    assert_eq!(added[0].day_of_week, 6);
    assert_eq!(added[0].millisecond, 250);
    assert!(comments(&published, 1).is_empty());
    assert_eq!(texts(&published), before);
    published.presentation().unwrap().comment_catalog().unwrap();

    // Exact patch and inverse.
    assert_eq!(source.to_bytes().unwrap(), source_bytes);
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

#[test]
fn add_set_remove_existing_comments() {
    let source_bytes = fixture(WITH_COMMENTS);
    let source = Package::from_bytes(&source_bytes).unwrap();
    let original = comments(&source, 0);
    assert_eq!(original.len(), 1);

    let mut edit = source.edit().unwrap();
    let mut added = new_comment("Second");
    added.author = original[0].author.clone();
    let position = edit.add_comment(Position::new(0), &added).unwrap();
    assert_eq!(position.get(), 1);
    edit.set_comment(
        Position::new(0),
        Position::new(0),
        &CommentChange {
            text: Some("Rewritten".into()),
            x: Some(7),
            y: None,
        },
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    let published = Package::from_bytes(&commit.package().to_bytes().unwrap()).unwrap();
    let after = comments(&published, 0);
    assert_eq!(after.len(), 2);
    assert_eq!(after[0].text, "Rewritten");
    assert_eq!(after[0].x, 7);
    assert_eq!(after[0].y, original[0].y);
    assert_eq!(after[0].index, original[0].index);
    assert_eq!(after[1].text, "Second");
    assert_eq!(after[1].index, original[0].index + 1);
    published.presentation().unwrap().comment_catalog().unwrap();

    let mut edit = published.edit().unwrap();
    edit.remove_comment(Position::new(0), Position::new(0))
        .unwrap();
    let removed = edit.commit().unwrap();
    let left = comments(removed.package(), 0);
    assert_eq!(left.len(), 1);
    assert_eq!(left[0].text, "Second");
}

#[test]
fn equal_comment_change_is_an_exact_noop_and_missing_comment_is_refused() {
    let source_bytes = fixture(WITH_COMMENTS);
    let source = Package::from_bytes(&source_bytes).unwrap();
    let original = comments(&source, 0);
    let mut edit = source.edit().unwrap();
    edit.set_comment(
        Position::new(0),
        Position::new(0),
        &CommentChange {
            text: Some(original[0].text.clone()),
            x: Some(original[0].x),
            y: None,
        },
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.patch().is_noop());
    assert_eq!(commit.package().to_bytes().unwrap(), source_bytes);

    let mut edit = source.edit().unwrap();
    assert!(
        edit.remove_comment(Position::new(0), Position::new(3))
            .is_err()
    );
}

#[test]
fn comment_dates_parse_strictly() {
    assert!(DateTime::parse("2026-02-30").is_none());
    assert!(DateTime::parse("2026-1-01").is_none());
    assert!(DateTime::parse("2026-01-01T25:00").is_none());
    let date = DateTime::parse("2024-02-29T23:59:59Z").unwrap();
    assert_eq!(date.to_string(), "2024-02-29T23:59:59.000");
    assert_eq!(date.day_of_week(), 4);
}
