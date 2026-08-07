use super::*;
use crate::keynote::KeynoteDocumentBuilder;
use crate::numbers::NumbersDocumentBuilder;
use crate::pages::PagesEditor;
use crate::shapes::{DrawablePoint, DrawableSize};
use crate::text::TextPosition;

const TEXT: &str = "First\nSecond\nThird";
const POSITION: DrawablePoint = DrawablePoint { x: 40.0, y: 80.0 };
const SIZE: DrawableSize = DrawableSize {
    width: 320.0,
    height: 140.0,
};

fn second_start() -> TextPosition {
    TextPosition::from_utf16_index("First\n".encode_utf16().count()).unwrap()
}

fn third_start() -> TextPosition {
    TextPosition::from_utf16_index("First\nSecond\n".encode_utf16().count()).unwrap()
}

fn restart_at_seven() -> ParagraphListNumbering {
    ParagraphListNumbering::StartAt(ParagraphListStart::new(7).unwrap())
}

#[test]
fn list_numbering_round_trips_and_continues_independently_in_every_suite() {
    let restart = restart_at_seven();

    let mut pages = PagesEditor::create_with_text("Lists").unwrap();
    let pages_box = pages.add_text_box(5, TEXT, POSITION, SIZE).unwrap();
    pages
        .set_text_box_paragraph_list(pages_box.drawable_object_id, ParagraphList::Numbered)
        .unwrap();
    let pages_before = pages.to_bytes().unwrap();
    pages
        .set_text_box_paragraph_list_numbering(
            pages_box.drawable_object_id,
            second_start(),
            restart,
        )
        .unwrap();
    let mut pages = PagesEditor::from_bytes(&pages.to_bytes().unwrap()).unwrap();
    assert_eq!(
        pages
            .text_box_paragraph_list_numbering(pages_box.drawable_object_id, second_start())
            .unwrap(),
        restart
    );
    assert_eq!(
        pages
            .text_box_paragraph_list_numbering(pages_box.drawable_object_id, third_start())
            .unwrap(),
        ParagraphListNumbering::Continue
    );
    pages
        .set_text_box_paragraph_list_numbering(
            pages_box.drawable_object_id,
            second_start(),
            ParagraphListNumbering::Continue,
        )
        .unwrap();
    assert_eq!(pages.to_bytes().unwrap(), pages_before);

    let mut numbers = NumbersDocumentBuilder::new().build().unwrap();
    let sheet_id = numbers.sheets().unwrap()[0].object_id;
    let numbers_box = numbers
        .add_sheet_text_box(sheet_id, TEXT, POSITION, SIZE)
        .unwrap();
    numbers
        .set_sheet_text_box_paragraph_list(
            sheet_id,
            numbers_box.drawable_object_id,
            ParagraphList::Numbered,
        )
        .unwrap();
    numbers
        .set_sheet_text_box_paragraph_list_numbering(
            sheet_id,
            numbers_box.drawable_object_id,
            second_start(),
            restart,
        )
        .unwrap();
    let numbers = crate::numbers::NumbersEditor::from_bytes(&numbers.to_bytes().unwrap()).unwrap();
    assert_eq!(
        numbers
            .sheet_text_box_paragraph_list_numbering(
                sheet_id,
                numbers_box.drawable_object_id,
                second_start(),
            )
            .unwrap(),
        restart
    );

    let mut keynote = KeynoteDocumentBuilder::new().build().unwrap();
    let keynote_box = keynote.add_slide_text_box(0, TEXT, POSITION, SIZE).unwrap();
    keynote
        .set_slide_text_box_paragraph_list(
            0,
            keynote_box.drawable_object_id,
            ParagraphList::Numbered,
        )
        .unwrap();
    keynote
        .set_slide_text_box_paragraph_list_numbering(
            0,
            keynote_box.drawable_object_id,
            second_start(),
            restart,
        )
        .unwrap();
    let keynote = crate::keynote::KeynoteEditor::from_bytes(&keynote.to_bytes().unwrap()).unwrap();
    assert_eq!(
        keynote
            .slide_text_box_paragraph_list_numbering(
                0,
                keynote_box.drawable_object_id,
                second_start(),
            )
            .unwrap(),
        restart
    );
}

#[test]
fn invalid_or_non_numbered_restarts_are_transactional() {
    let mut pages = PagesEditor::create_with_text("Lists").unwrap();
    let text_box = pages.add_text_box(5, TEXT, POSITION, SIZE).unwrap();
    let before = pages.to_bytes().unwrap();
    assert!(
        pages
            .set_text_box_paragraph_list_numbering(
                text_box.drawable_object_id,
                TextPosition::from_utf16_index(1).unwrap(),
                restart_at_seven(),
            )
            .is_err()
    );
    assert!(
        pages
            .set_text_box_paragraph_list_numbering(
                text_box.drawable_object_id,
                second_start(),
                restart_at_seven(),
            )
            .is_err()
    );
    assert_eq!(pages.to_bytes().unwrap(), before);
}
