use super::*;
use crate::keynote::KeynoteDocumentBuilder;
use crate::numbers::NumbersDocumentBuilder;
use crate::pages::PagesEditor;
use crate::shapes::{DrawablePoint, DrawableSize};
use crate::text::ParagraphStart;

const TEXT: &str = "Top\nNested\nDeep";
const POSITION: DrawablePoint = DrawablePoint { x: 40.0, y: 80.0 };
const SIZE: DrawableSize = DrawableSize {
    width: 320.0,
    height: 140.0,
};

fn nested_start() -> ParagraphStart {
    ParagraphStart::from_utf16_index("Top\n".encode_utf16().count()).unwrap()
}

fn deep_start() -> ParagraphStart {
    ParagraphStart::from_utf16_index("Top\nNested\n".encode_utf16().count()).unwrap()
}

#[test]
fn list_levels_round_trip_coalesce_and_reset_in_every_suite() {
    let level_two = ParagraphListLevel::new(2).unwrap();
    let expected = vec![
        ParagraphListLevelPlacement::new(ParagraphStart::ZERO, ParagraphListLevel::ZERO),
        ParagraphListLevelPlacement::new(nested_start(), ParagraphListLevel::ONE),
        ParagraphListLevelPlacement::new(deep_start(), level_two),
    ];

    let mut pages = PagesEditor::create_with_text("Lists").unwrap();
    let pages_box = pages.add_text_box(5, TEXT, POSITION, SIZE).unwrap();
    pages
        .set_text_box_paragraph_list(pages_box.drawable_object_id, ParagraphList::Bullet)
        .unwrap();
    pages
        .set_text_box_paragraph_list_level(
            pages_box.drawable_object_id,
            nested_start(),
            ParagraphListLevel::ONE,
        )
        .unwrap();
    pages
        .set_text_box_paragraph_list_level(pages_box.drawable_object_id, deep_start(), level_two)
        .unwrap();
    assert_eq!(
        pages
            .text_box_paragraph_list_levels(pages_box.drawable_object_id)
            .unwrap(),
        expected
    );
    let mut pages = PagesEditor::from_bytes(&pages.to_bytes().unwrap()).unwrap();
    assert_eq!(
        pages
            .text_box_paragraph_list_level(pages_box.drawable_object_id, deep_start())
            .unwrap(),
        level_two
    );
    assert!(
        pages
            .reset_text_box_paragraph_list_level(pages_box.drawable_object_id, nested_start())
            .unwrap()
    );
    assert_eq!(
        pages
            .text_box_paragraph_list_levels(pages_box.drawable_object_id)
            .unwrap(),
        vec![
            ParagraphListLevelPlacement::new(ParagraphStart::ZERO, ParagraphListLevel::ZERO),
            ParagraphListLevelPlacement::new(deep_start(), level_two),
        ]
    );
    assert!(
        pages
            .reset_text_box_paragraph_list_level(pages_box.drawable_object_id, deep_start())
            .unwrap()
    );
    assert!(
        !pages
            .reset_text_box_paragraph_list_level(pages_box.drawable_object_id, deep_start())
            .unwrap()
    );

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
        .set_sheet_text_box_paragraph_list_level(
            sheet_id,
            numbers_box.drawable_object_id,
            nested_start(),
            ParagraphListLevel::ONE,
        )
        .unwrap();
    let numbers = crate::numbers::NumbersEditor::from_bytes(&numbers.to_bytes().unwrap()).unwrap();
    assert_eq!(
        numbers
            .sheet_text_box_paragraph_list_level(
                sheet_id,
                numbers_box.drawable_object_id,
                nested_start()
            )
            .unwrap(),
        ParagraphListLevel::ONE
    );

    let mut keynote = KeynoteDocumentBuilder::new().build().unwrap();
    let keynote_box = keynote.add_slide_text_box(0, TEXT, POSITION, SIZE).unwrap();
    keynote
        .set_slide_text_box_paragraph_list(0, keynote_box.drawable_object_id, ParagraphList::Bullet)
        .unwrap();
    keynote
        .set_slide_text_box_paragraph_list_level(
            0,
            keynote_box.drawable_object_id,
            deep_start(),
            level_two,
        )
        .unwrap();
    let keynote = crate::keynote::KeynoteEditor::from_bytes(&keynote.to_bytes().unwrap()).unwrap();
    assert_eq!(
        keynote
            .slide_text_box_paragraph_list_level(0, keynote_box.drawable_object_id, deep_start())
            .unwrap(),
        level_two
    );
}

#[test]
fn non_paragraph_list_level_updates_are_rejected_transactionally() {
    let mut pages = PagesEditor::create_with_text("Lists").unwrap();
    let text_box = pages.add_text_box(5, TEXT, POSITION, SIZE).unwrap();
    let before = pages.to_bytes().unwrap();
    assert!(
        pages
            .set_text_box_paragraph_list_level(
                text_box.drawable_object_id,
                ParagraphStart::from_utf16_index(1).unwrap(),
                ParagraphListLevel::ONE,
            )
            .is_err()
    );
    assert_eq!(pages.to_bytes().unwrap(), before);
}
