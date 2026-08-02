use super::{TextLanguage, TextLanguageRun, TextPosition};
use crate::keynote::{KeynoteDocumentBuilder, KeynoteEditor};
use crate::numbers::{NumbersDocumentBuilder, NumbersEditor};
use crate::pages::PagesEditor;
use crate::shapes::{DrawablePoint, DrawableSize};

const TEXT: &str = "Hello 😀\nBonjour";
const POSITION: DrawablePoint = DrawablePoint { x: 40.0, y: 80.0 };
const SIZE: DrawableSize = DrawableSize {
    width: 320.0,
    height: 140.0,
};

fn french_start() -> TextPosition {
    TextPosition::from_utf16_index("Hello 😀\n".encode_utf16().count()).unwrap()
}

fn explicit(tag: &str) -> TextLanguage {
    TextLanguage::tag(tag).unwrap()
}

#[test]
fn scratch_text_language_runs_round_trip_in_every_suite() {
    let english = explicit("en");
    let french = explicit("fr-CA");

    let mut pages = PagesEditor::builder()
        .body_text("Body")
        .language("en")
        .build()
        .unwrap();
    let pages_box = pages.add_text_box(4, TEXT, POSITION, SIZE).unwrap();
    assert_eq!(
        pages
            .text_box_text_languages(pages_box.drawable_object_id)
            .unwrap(),
        vec![TextLanguageRun::new(TextPosition::ZERO, english.clone())]
    );
    pages
        .set_text_box_text_language(
            pages_box.drawable_object_id,
            french_start(),
            english.clone(),
        )
        .unwrap();
    assert_eq!(
        pages
            .text_box_text_languages(pages_box.drawable_object_id)
            .unwrap()
            .len(),
        1
    );
    pages
        .set_text_box_text_language(pages_box.drawable_object_id, french_start(), french.clone())
        .unwrap();
    let mut pages = PagesEditor::from_bytes(&pages.to_bytes().unwrap()).unwrap();
    assert_eq!(
        pages
            .text_box_text_language(pages_box.drawable_object_id, TextPosition::ZERO)
            .unwrap(),
        english
    );
    assert_eq!(
        pages
            .text_box_text_language(pages_box.drawable_object_id, french_start())
            .unwrap(),
        french
    );
    let text_end = TextPosition::from_utf16_index(TEXT.encode_utf16().count()).unwrap();
    pages
        .set_text_box_text_language(pages_box.drawable_object_id, text_end, explicit("en"))
        .unwrap();
    assert_eq!(
        pages
            .text_box_text_languages(pages_box.drawable_object_id)
            .unwrap()
            .len(),
        3
    );
    assert!(
        pages
            .remove_text_box_text_language_boundary(pages_box.drawable_object_id, french_start())
            .unwrap()
    );
    assert_eq!(
        pages
            .text_box_text_languages(pages_box.drawable_object_id)
            .unwrap()
            .len(),
        1
    );
    assert!(
        !pages
            .remove_text_box_text_language_boundary(pages_box.drawable_object_id, french_start())
            .unwrap()
    );
    assert!(
        pages
            .reset_text_box_text_languages(pages_box.drawable_object_id)
            .unwrap()
    );
    assert_eq!(
        pages
            .text_box_text_language(pages_box.drawable_object_id, french_start())
            .unwrap(),
        TextLanguage::Automatic
    );
    pages
        .set_text_box_text_language(pages_box.drawable_object_id, french_start(), explicit("fr"))
        .unwrap();
    assert_eq!(
        pages
            .text_box_text_languages(pages_box.drawable_object_id)
            .unwrap(),
        vec![
            TextLanguageRun::new(TextPosition::ZERO, TextLanguage::Automatic),
            TextLanguageRun::new(french_start(), explicit("fr")),
        ]
    );

    let mut numbers = NumbersDocumentBuilder::new()
        .language("de")
        .build()
        .unwrap();
    let sheet_id = numbers.sheets().unwrap()[0].object_id;
    let numbers_box = numbers
        .add_sheet_text_box(sheet_id, TEXT, POSITION, SIZE)
        .unwrap();
    assert_eq!(
        numbers
            .sheet_text_box_text_language(
                sheet_id,
                numbers_box.drawable_object_id,
                TextPosition::ZERO
            )
            .unwrap(),
        explicit("de")
    );
    numbers
        .set_sheet_text_box_text_language(
            sheet_id,
            numbers_box.drawable_object_id,
            french_start(),
            explicit("fr"),
        )
        .unwrap();
    let numbers = NumbersEditor::from_bytes(&numbers.to_bytes().unwrap()).unwrap();
    assert_eq!(
        numbers
            .sheet_text_box_text_language(sheet_id, numbers_box.drawable_object_id, french_start())
            .unwrap(),
        explicit("fr")
    );

    let mut keynote = KeynoteDocumentBuilder::new()
        .language("ja")
        .build()
        .unwrap();
    let keynote_box = keynote.add_slide_text_box(0, TEXT, POSITION, SIZE).unwrap();
    assert_eq!(
        keynote
            .slide_text_box_text_language(0, keynote_box.drawable_object_id, TextPosition::ZERO)
            .unwrap(),
        explicit("ja")
    );
    keynote
        .set_slide_text_box_text_language(
            0,
            keynote_box.drawable_object_id,
            french_start(),
            explicit("fr"),
        )
        .unwrap();
    let keynote = KeynoteEditor::from_bytes(&keynote.to_bytes().unwrap()).unwrap();
    assert_eq!(
        keynote
            .slide_text_box_text_language(0, keynote_box.drawable_object_id, french_start())
            .unwrap(),
        explicit("fr")
    );
}

#[test]
fn split_surrogate_language_updates_are_rejected_transactionally() {
    let mut pages = PagesEditor::create_with_text("Body").unwrap();
    let text_box = pages.add_text_box(4, "A😀B", POSITION, SIZE).unwrap();
    let before = pages.to_bytes().unwrap();
    assert!(
        pages
            .set_text_box_text_language(
                text_box.drawable_object_id,
                TextPosition::from_utf16_index(2).unwrap(),
                explicit("fr"),
            )
            .is_err()
    );
    assert_eq!(pages.to_bytes().unwrap(), before);
}
