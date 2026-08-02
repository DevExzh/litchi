use super::{TextBookmarkName, TextBookmarkSettings, TextBookmarkVisibility, TextRange};
use crate::pages::PagesEditor;

const TEXT: &str = "Alpha 😀 Beta Gamma";

fn range(start: usize, end: usize) -> TextRange {
    TextRange::from_utf16_indexes(start, end).unwrap()
}

fn named(value: &str) -> TextBookmarkSettings {
    TextBookmarkSettings::new().with_name(TextBookmarkName::new(value).unwrap())
}

#[test]
fn scratch_pages_body_bookmark_crud_round_trips_and_restores_exactly() {
    let mut pages = PagesEditor::create_with_text(TEXT).unwrap();
    let baseline = pages.to_bytes().unwrap();

    let created = pages
        .add_body_bookmark(range(0, 5), TextBookmarkSettings::new())
        .unwrap();
    assert_eq!(
        pages.body_bookmarks().unwrap().as_slice(),
        std::slice::from_ref(&created)
    );
    let mut pages = PagesEditor::from_bytes(&pages.to_bytes().unwrap()).unwrap();
    let settings = named("Methods").with_visibility(TextBookmarkVisibility::Hidden);
    let updated = pages
        .update_body_bookmark(created.id, range(9, 13), settings.clone())
        .unwrap();
    assert_eq!(updated.id, created.id);
    assert_eq!(updated.range, range(9, 13));
    assert_eq!(updated.settings, settings);

    let before_overlap = pages.to_bytes().unwrap();
    assert!(
        pages
            .add_body_bookmark(range(9, 19), named("Overlap"))
            .is_err()
    );
    assert_eq!(pages.to_bytes().unwrap(), before_overlap);

    assert_eq!(pages.remove_body_bookmark(created.id).unwrap(), updated);
    assert!(pages.body_bookmarks().unwrap().is_empty());
    assert_eq!(pages.to_bytes().unwrap(), baseline);
}

#[test]
fn split_surrogate_and_empty_bookmark_ranges_are_rejected_transactionally() {
    let mut pages = PagesEditor::create_with_text("A😀B").unwrap();
    let baseline = pages.to_bytes().unwrap();
    assert!(TextRange::from_utf16_indexes(1, 1).is_err());
    assert!(
        pages
            .add_body_bookmark(range(1, 2), TextBookmarkSettings::new())
            .is_err()
    );
    assert_eq!(pages.to_bytes().unwrap(), baseline);
}

#[test]
fn replacing_bookmarked_text_reclaims_the_orphaned_object() {
    let mut pages = PagesEditor::create_with_text(TEXT).unwrap();
    let baseline = pages.to_bytes().unwrap();
    pages
        .add_body_bookmark(range(0, 5), named("Alpha"))
        .unwrap();
    pages.replace_body_text(0..5, "Alpha").unwrap();
    assert!(pages.body_bookmarks().unwrap().is_empty());
    assert_eq!(pages.to_bytes().unwrap(), baseline);
}
