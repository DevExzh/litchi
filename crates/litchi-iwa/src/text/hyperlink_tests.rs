use super::{TextHyperlinkTarget, TextPosition, TextRange};
use crate::keynote::{KeynoteDocumentBuilder, KeynoteEditor};
use crate::numbers::{NumbersDocumentBuilder, NumbersEditor};
use crate::pages::PagesEditor;
use crate::shapes::{DrawablePoint, DrawableSize};

const TEXT: &str = "Alpha 😀 Beta Gamma";
const POSITION: DrawablePoint = DrawablePoint { x: 40.0, y: 80.0 };
const SIZE: DrawableSize = DrawableSize {
    width: 320.0,
    height: 140.0,
};

fn text_range(value: &str) -> TextRange {
    let byte_start = TEXT.find(value).unwrap();
    let byte_end = byte_start + value.len();
    TextRange::new(
        TextPosition::from_utf16_index(TEXT[..byte_start].encode_utf16().count()).unwrap(),
        TextPosition::from_utf16_index(TEXT[..byte_end].encode_utf16().count()).unwrap(),
    )
    .unwrap()
}

fn target(value: &str) -> TextHyperlinkTarget {
    TextHyperlinkTarget::new(value).unwrap()
}

#[test]
fn hyperlink_crud_round_trips_and_restores_pages_exactly() {
    let mut pages = PagesEditor::create_with_text("Body").unwrap();
    let text_box = pages.add_text_box(4, TEXT, POSITION, SIZE).unwrap();
    let baseline = pages.to_bytes().unwrap();

    let created = pages
        .add_text_box_hyperlink(
            text_box.drawable_object_id,
            text_range("Alpha"),
            target("https://example.com/alpha"),
        )
        .unwrap();
    assert_eq!(
        pages
            .text_box_hyperlinks(text_box.drawable_object_id)
            .unwrap()
            .as_slice(),
        std::slice::from_ref(&created)
    );
    let after_create = pages.to_bytes().unwrap();
    let mut pages = PagesEditor::from_bytes(&after_create).unwrap();
    let updated = pages
        .update_text_box_hyperlink(
            text_box.drawable_object_id,
            created.id,
            text_range("Beta"),
            target("mailto:owner@example.com"),
        )
        .unwrap();
    assert_eq!(updated.id, created.id);
    assert_eq!(updated.range, text_range("Beta"));
    assert_eq!(updated.target.as_str(), "mailto:owner@example.com");

    let before_overlap = pages.to_bytes().unwrap();
    assert!(
        pages
            .add_text_box_hyperlink(
                text_box.drawable_object_id,
                text_range("Beta Gamma"),
                target("https://example.com/overlap"),
            )
            .is_err()
    );
    assert_eq!(pages.to_bytes().unwrap(), before_overlap);

    let removed = pages
        .remove_text_box_hyperlink(text_box.drawable_object_id, created.id)
        .unwrap();
    assert_eq!(removed, updated);
    assert!(
        pages
            .text_box_hyperlinks(text_box.drawable_object_id)
            .unwrap()
            .is_empty()
    );
    assert_eq!(pages.to_bytes().unwrap(), baseline);
}

#[test]
fn scratch_hyperlinks_work_in_numbers_and_keynote() {
    let mut numbers = NumbersDocumentBuilder::new().build().unwrap();
    let sheet_id = numbers.sheets().unwrap()[0].object_id;
    let numbers_box = numbers
        .add_sheet_text_box(sheet_id, TEXT, POSITION, SIZE)
        .unwrap();
    let numbers_baseline = numbers.to_bytes().unwrap();
    let numbers_link = numbers
        .add_sheet_text_box_hyperlink(
            sheet_id,
            numbers_box.drawable_object_id,
            text_range("Beta"),
            target("https://example.com/numbers"),
        )
        .unwrap();
    let mut numbers = NumbersEditor::from_bytes(&numbers.to_bytes().unwrap()).unwrap();
    assert_eq!(
        numbers
            .sheet_text_box_hyperlinks(sheet_id, numbers_box.drawable_object_id)
            .unwrap()
            .as_slice(),
        std::slice::from_ref(&numbers_link)
    );
    numbers
        .remove_sheet_text_box_hyperlink(sheet_id, numbers_box.drawable_object_id, numbers_link.id)
        .unwrap();
    assert_eq!(numbers.to_bytes().unwrap(), numbers_baseline);

    let mut keynote = KeynoteDocumentBuilder::new().build().unwrap();
    let keynote_box = keynote.add_slide_text_box(0, TEXT, POSITION, SIZE).unwrap();
    let keynote_baseline = keynote.to_bytes().unwrap();
    let keynote_link = keynote
        .add_slide_text_box_hyperlink(
            0,
            keynote_box.drawable_object_id,
            text_range("Gamma"),
            target("?slide=next"),
        )
        .unwrap();
    let mut keynote = KeynoteEditor::from_bytes(&keynote.to_bytes().unwrap()).unwrap();
    assert_eq!(
        keynote
            .slide_text_box_hyperlinks(0, keynote_box.drawable_object_id)
            .unwrap()
            .as_slice(),
        std::slice::from_ref(&keynote_link)
    );
    keynote
        .remove_slide_text_box_hyperlink(0, keynote_box.drawable_object_id, keynote_link.id)
        .unwrap();
    assert_eq!(keynote.to_bytes().unwrap(), keynote_baseline);
}

#[test]
fn split_surrogate_hyperlink_ranges_are_rejected_transactionally() {
    let mut pages = PagesEditor::create_with_text("Body").unwrap();
    let text_box = pages.add_text_box(4, "A😀B", POSITION, SIZE).unwrap();
    let invalid = TextRange::from_utf16_indexes(1, 2).unwrap();
    let before = pages.to_bytes().unwrap();
    assert!(
        pages
            .add_text_box_hyperlink(
                text_box.drawable_object_id,
                invalid,
                target("https://example.com"),
            )
            .is_err()
    );
    assert_eq!(pages.to_bytes().unwrap(), before);
}

#[test]
fn replacing_linked_text_reclaims_the_orphaned_hyperlink_graph() {
    let mut pages = PagesEditor::create_with_text("Body").unwrap();
    let text_box = pages.add_text_box(4, TEXT, POSITION, SIZE).unwrap();
    let baseline = pages.to_bytes().unwrap();
    pages
        .add_text_box_hyperlink(
            text_box.drawable_object_id,
            text_range("Alpha"),
            target("https://example.com/alpha"),
        )
        .unwrap();
    pages
        .replace_drawable_text(text_box.drawable_object_id, 0..5, "Alpha")
        .unwrap();
    assert!(
        pages
            .text_box_hyperlinks(text_box.drawable_object_id)
            .unwrap()
            .is_empty()
    );
    assert_eq!(pages.to_bytes().unwrap(), baseline);
}

#[test]
fn adjacent_hyperlinks_keep_independent_ranges_and_lifecycles() {
    let mut pages = PagesEditor::create_with_text("Body").unwrap();
    let text_box = pages.add_text_box(4, "Alpha Beta", POSITION, SIZE).unwrap();
    let baseline = pages.to_bytes().unwrap();
    let first = pages
        .add_text_box_hyperlink(
            text_box.drawable_object_id,
            TextRange::from_utf16_indexes(0, 5).unwrap(),
            target("https://example.com/alpha"),
        )
        .unwrap();
    let second = pages
        .add_text_box_hyperlink(
            text_box.drawable_object_id,
            TextRange::from_utf16_indexes(5, 10).unwrap(),
            target("https://example.com/beta"),
        )
        .unwrap();
    pages
        .remove_text_box_hyperlink(text_box.drawable_object_id, first.id)
        .unwrap();
    assert_eq!(
        pages
            .text_box_hyperlinks(text_box.drawable_object_id)
            .unwrap()
            .as_slice(),
        std::slice::from_ref(&second)
    );
    pages
        .remove_text_box_hyperlink(text_box.drawable_object_id, second.id)
        .unwrap();
    assert_eq!(pages.to_bytes().unwrap(), baseline);
}
