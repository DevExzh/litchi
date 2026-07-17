use super::{TextCommentBody, TextPosition, TextRange};
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

fn body(value: &str) -> TextCommentBody {
    TextCommentBody::new(value).unwrap()
}

#[test]
fn text_comment_crud_round_trips_and_restores_pages_exactly() {
    let mut pages = PagesEditor::create_with_text("Body").unwrap();
    let text_box = pages.add_text_box(4, TEXT, POSITION, SIZE).unwrap();
    let baseline = pages.to_bytes().unwrap();

    let created = pages
        .add_text_box_comment(
            text_box.drawable_object_id,
            text_range("Alpha"),
            body("Review alpha"),
        )
        .unwrap();
    assert!(
        pages
            .text_box_highlights(text_box.drawable_object_id)
            .unwrap()
            .is_empty()
    );
    assert_eq!(
        pages
            .text_box_comments(text_box.drawable_object_id)
            .unwrap()
            .as_slice(),
        std::slice::from_ref(&created)
    );

    let mut pages = PagesEditor::from_bytes(&pages.to_bytes().unwrap()).unwrap();
    let updated = pages
        .update_text_box_comment(
            text_box.drawable_object_id,
            created.id,
            text_range("Beta"),
            body("Reviewed 😀"),
        )
        .unwrap();
    assert_eq!(updated.id, created.id);
    assert_eq!(updated.range, text_range("Beta"));
    assert_eq!(updated.body.as_str(), "Reviewed 😀");
    assert_eq!(updated.creation_date_seconds, created.creation_date_seconds);
    assert_eq!(updated.author_object_id, created.author_object_id);
    assert_eq!(updated.storage_uuid, created.storage_uuid);

    let before_overlap = pages.to_bytes().unwrap();
    assert!(
        pages
            .add_text_box_comment(
                text_box.drawable_object_id,
                text_range("Beta Gamma"),
                body("overlap"),
            )
            .is_err()
    );
    assert_eq!(pages.to_bytes().unwrap(), before_overlap);

    assert_eq!(
        pages
            .remove_text_box_comment(text_box.drawable_object_id, created.id)
            .unwrap(),
        updated
    );
    assert!(
        pages
            .text_box_comments(text_box.drawable_object_id)
            .unwrap()
            .is_empty()
    );
    assert_eq!(pages.to_bytes().unwrap(), baseline);
}

#[test]
fn scratch_text_comments_work_in_numbers_and_keynote() {
    let mut numbers = NumbersDocumentBuilder::new().build().unwrap();
    let sheet_id = numbers.sheets().unwrap()[0].object_id;
    let numbers_box = numbers
        .add_sheet_text_box(sheet_id, TEXT, POSITION, SIZE)
        .unwrap();
    let numbers_baseline = numbers.to_bytes().unwrap();
    let numbers_comment = numbers
        .add_sheet_text_box_comment(
            sheet_id,
            numbers_box.drawable_object_id,
            text_range("Beta"),
            body("Numbers note"),
        )
        .unwrap();
    let mut numbers = NumbersEditor::from_bytes(&numbers.to_bytes().unwrap()).unwrap();
    assert_eq!(
        numbers
            .sheet_text_box_comments(sheet_id, numbers_box.drawable_object_id)
            .unwrap()
            .as_slice(),
        std::slice::from_ref(&numbers_comment)
    );
    numbers
        .remove_sheet_text_box_comment(sheet_id, numbers_box.drawable_object_id, numbers_comment.id)
        .unwrap();
    assert_eq!(numbers.to_bytes().unwrap(), numbers_baseline);

    let mut keynote = KeynoteDocumentBuilder::new().build().unwrap();
    let keynote_box = keynote.add_slide_text_box(0, TEXT, POSITION, SIZE).unwrap();
    let keynote_baseline = keynote.to_bytes().unwrap();
    let keynote_comment = keynote
        .add_slide_text_box_comment(
            0,
            keynote_box.drawable_object_id,
            text_range("Gamma"),
            body("Keynote note"),
        )
        .unwrap();
    let mut keynote = KeynoteEditor::from_bytes(&keynote.to_bytes().unwrap()).unwrap();
    assert_eq!(
        keynote
            .slide_text_box_comments(0, keynote_box.drawable_object_id)
            .unwrap()
            .as_slice(),
        std::slice::from_ref(&keynote_comment)
    );
    keynote
        .remove_slide_text_box_comment(0, keynote_box.drawable_object_id, keynote_comment.id)
        .unwrap();
    assert_eq!(keynote.to_bytes().unwrap(), keynote_baseline);
}

#[test]
fn highlights_and_comments_are_classified_independently() {
    let mut pages = PagesEditor::create_with_text("Body").unwrap();
    let text_box = pages.add_text_box(4, "Alpha Beta", POSITION, SIZE).unwrap();
    let baseline = pages.to_bytes().unwrap();
    let highlight = pages
        .add_text_box_highlight(
            text_box.drawable_object_id,
            TextRange::from_utf16_indexes(0, 5).unwrap(),
        )
        .unwrap();
    let comment = pages
        .add_text_box_comment(
            text_box.drawable_object_id,
            TextRange::from_utf16_indexes(5, 10).unwrap(),
            body("Adjacent"),
        )
        .unwrap();
    assert_eq!(
        pages
            .text_box_highlights(text_box.drawable_object_id)
            .unwrap()
            .as_slice(),
        std::slice::from_ref(&highlight)
    );
    assert_eq!(
        pages
            .text_box_comments(text_box.drawable_object_id)
            .unwrap()
            .as_slice(),
        std::slice::from_ref(&comment)
    );
    pages
        .remove_text_box_highlight(text_box.drawable_object_id, highlight.id)
        .unwrap();
    pages
        .remove_text_box_comment(text_box.drawable_object_id, comment.id)
        .unwrap();
    assert_eq!(pages.to_bytes().unwrap(), baseline);
}

#[test]
fn invalid_comment_ranges_are_rejected_transactionally() {
    let mut pages = PagesEditor::create_with_text("Body").unwrap();
    let text_box = pages.add_text_box(4, "A😀B", POSITION, SIZE).unwrap();
    let before = pages.to_bytes().unwrap();
    assert!(
        pages
            .add_text_box_comment(
                text_box.drawable_object_id,
                TextRange::from_utf16_indexes(1, 2).unwrap(),
                body("split surrogate"),
            )
            .is_err()
    );
    assert_eq!(pages.to_bytes().unwrap(), before);
}

#[test]
fn replacing_commented_text_reclaims_the_annotation_graph() {
    let mut pages = PagesEditor::create_with_text("Body").unwrap();
    let text_box = pages.add_text_box(4, TEXT, POSITION, SIZE).unwrap();
    let baseline = pages.to_bytes().unwrap();
    pages
        .add_text_box_comment(
            text_box.drawable_object_id,
            text_range("Alpha"),
            body("remove with text"),
        )
        .unwrap();
    pages
        .replace_drawable_text(text_box.drawable_object_id, 0..5, "Alpha")
        .unwrap();
    assert!(
        pages
            .text_box_comments(text_box.drawable_object_id)
            .unwrap()
            .is_empty()
    );
    assert_eq!(pages.to_bytes().unwrap(), baseline);
}
