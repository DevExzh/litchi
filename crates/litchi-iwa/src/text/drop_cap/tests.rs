use crate::keynote::KeynoteDocumentBuilder;
use crate::numbers::NumbersDocumentBuilder;
use crate::pages::PagesEditor;
use crate::shapes::{DrawablePoint, DrawableSize};
use crate::text::{IWorkTextEditor, ParagraphDropCapPlacement};

use super::*;

fn model(lines: u8, characters: u32) -> ParagraphDropCap {
    ParagraphDropCap::new(
        DropCapLineCount::new(lines).unwrap(),
        DropCapCharacterCount::new(characters).unwrap(),
    )
}

fn text_box_geometry() -> (DrawablePoint, DrawableSize) {
    (
        DrawablePoint { x: 72.0, y: 144.0 },
        DrawableSize {
            width: 480.0,
            height: 240.0,
        },
    )
}

#[test]
fn scratch_pages_supports_multi_paragraph_drop_cap_crud() {
    let mut editor = PagesEditor::create_with_text("Body").unwrap();
    let (position, size) = text_box_geometry();
    let created = editor
        .add_text_box(4, "😀 Alpha\nBeta paragraph", position, size)
        .unwrap();
    let first = model(4, 2)
        .with_wrap(DropCapWrap::Contour)
        .with_padding(DropCapPadding::from_points(6.0).unwrap());
    let second_start = ParagraphStart::from_utf16_index(9).unwrap();
    let second = model(5, 1).with_raised_lines(DropCapRaisedLines::new(2).unwrap());

    editor
        .set_text_box_paragraph_drop_cap(created.drawable_object_id, ParagraphStart::ZERO, first)
        .unwrap();
    editor
        .set_text_box_paragraph_drop_cap(created.drawable_object_id, second_start, second)
        .unwrap();
    assert_eq!(
        editor
            .text_box_paragraph_drop_caps(created.drawable_object_id)
            .unwrap(),
        vec![
            ParagraphDropCapPlacement {
                paragraph_start: ParagraphStart::ZERO,
                drop_cap: first,
            },
            ParagraphDropCapPlacement {
                paragraph_start: second_start,
                drop_cap: second,
            },
        ]
    );

    let updated = second.with_outdent(DropCapOutdent::from_ratio(0.25).unwrap());
    editor
        .set_text_box_paragraph_drop_cap(created.drawable_object_id, second_start, updated)
        .unwrap();
    let reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .text_box_paragraph_drop_cap(created.drawable_object_id, second_start)
            .unwrap(),
        Some(updated)
    );

    assert!(
        editor
            .remove_text_box_paragraph_drop_cap(created.drawable_object_id, ParagraphStart::ZERO,)
            .unwrap()
    );
    assert!(
        !editor
            .remove_text_box_paragraph_drop_cap(created.drawable_object_id, ParagraphStart::ZERO,)
            .unwrap()
    );
}

#[test]
fn scratch_numbers_and_keynote_support_drop_cap_crud() {
    let (position, size) = text_box_geometry();
    let expected = model(6, 2)
        .with_raised_lines(DropCapRaisedLines::new(3).unwrap())
        .with_wrap(DropCapWrap::Contour)
        .with_padding(DropCapPadding::from_points(8.0).unwrap())
        .with_outdent(DropCapOutdent::from_ratio(0.4).unwrap());

    let mut numbers = NumbersDocumentBuilder::new().build().unwrap();
    let sheet_id = numbers.sheets().unwrap()[0].object_id;
    let numbers_box = numbers
        .add_sheet_text_box(sheet_id, "Numbers Drop Cap", position, size)
        .unwrap();
    numbers
        .set_sheet_text_box_paragraph_drop_cap(
            sheet_id,
            numbers_box.drawable_object_id,
            ParagraphStart::ZERO,
            expected,
        )
        .unwrap();
    let numbers = crate::numbers::NumbersEditor::from_bytes(&numbers.to_bytes().unwrap()).unwrap();
    assert_eq!(
        numbers
            .sheet_text_box_paragraph_drop_cap(
                sheet_id,
                numbers_box.drawable_object_id,
                ParagraphStart::ZERO,
            )
            .unwrap(),
        Some(expected)
    );

    let mut keynote = KeynoteDocumentBuilder::new().build().unwrap();
    let keynote_box = keynote
        .add_slide_text_box(0, "Keynote Drop Cap", position, size)
        .unwrap();
    keynote
        .set_slide_text_box_paragraph_drop_cap(
            0,
            keynote_box.drawable_object_id,
            ParagraphStart::ZERO,
            expected,
        )
        .unwrap();
    assert!(
        keynote
            .remove_slide_text_box_paragraph_drop_cap(
                0,
                keynote_box.drawable_object_id,
                ParagraphStart::ZERO,
            )
            .unwrap()
    );
    assert!(
        keynote
            .slide_text_box_paragraph_drop_caps(0, keynote_box.drawable_object_id)
            .unwrap()
            .is_empty()
    );
}

#[test]
fn shared_numbers_style_uses_copy_on_write() {
    let (position, size) = text_box_geometry();
    let original = model(4, 1).with_padding(DropCapPadding::from_points(3.0).unwrap());
    let updated = model(7, 2)
        .with_wrap(DropCapWrap::Contour)
        .with_outdent(DropCapOutdent::from_ratio(0.3).unwrap());
    let mut editor = NumbersDocumentBuilder::new().build().unwrap();
    let sheet_id = editor.sheets().unwrap()[0].object_id;
    let source = editor
        .add_sheet_text_box(sheet_id, "Shared style source", position, size)
        .unwrap();
    editor
        .set_sheet_text_box_paragraph_drop_cap(
            sheet_id,
            source.drawable_object_id,
            ParagraphStart::ZERO,
            original,
        )
        .unwrap();
    let duplicate = editor
        .duplicate_sheet_text_box(sheet_id, source.drawable_object_id, "Shared style clone")
        .unwrap();
    assert_eq!(
        editor
            .sheet_text_box_paragraph_drop_cap(
                sheet_id,
                duplicate.drawable_object_id,
                ParagraphStart::ZERO,
            )
            .unwrap(),
        Some(original)
    );

    editor
        .set_sheet_text_box_paragraph_drop_cap(
            sheet_id,
            duplicate.drawable_object_id,
            ParagraphStart::ZERO,
            updated,
        )
        .unwrap();
    assert_eq!(
        editor
            .sheet_text_box_paragraph_drop_cap(
                sheet_id,
                source.drawable_object_id,
                ParagraphStart::ZERO,
            )
            .unwrap(),
        Some(original)
    );
    assert_eq!(
        editor
            .sheet_text_box_paragraph_drop_cap(
                sheet_id,
                duplicate.drawable_object_id,
                ParagraphStart::ZERO,
            )
            .unwrap(),
        Some(updated)
    );

    editor
        .remove_sheet_text_box_paragraph_drop_cap(
            sheet_id,
            duplicate.drawable_object_id,
            ParagraphStart::ZERO,
        )
        .unwrap();
    let reopened = crate::numbers::NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .sheet_text_box_paragraph_drop_cap(
                sheet_id,
                source.drawable_object_id,
                ParagraphStart::ZERO,
            )
            .unwrap(),
        Some(original)
    );
    assert_eq!(
        reopened
            .sheet_text_box_paragraph_drop_cap(
                sheet_id,
                duplicate.drawable_object_id,
                ParagraphStart::ZERO,
            )
            .unwrap(),
        None
    );
}

#[test]
fn invalid_paragraph_start_is_transactional() {
    let mut editor = PagesEditor::create_with_text("Body").unwrap();
    let (position, size) = text_box_geometry();
    let created = editor
        .add_text_box(4, "Not a boundary", position, size)
        .unwrap();
    let before = editor.to_bytes().unwrap();
    let invalid = ParagraphStart::from_utf16_index(1).unwrap();
    assert!(
        editor
            .set_text_box_paragraph_drop_cap(created.drawable_object_id, invalid, model(3, 1))
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn create_remove_restores_storage_object_exactly() {
    let mut pages = PagesEditor::create_with_text("Body").unwrap();
    let (position, size) = text_box_geometry();
    let created = pages
        .add_text_box(4, "Storage preservation", position, size)
        .unwrap();
    let storage_id = created.storage.object_id;
    let archive_name = super::storage::locate(pages.package(), storage_id)
        .unwrap()
        .archive_name;
    let archive = pages.package().archive(&archive_name).unwrap();
    let before = archive.object(storage_id).unwrap();
    let before_info = before.archive_info.clone();
    let before_messages = before.messages.clone();

    let mut text = IWorkTextEditor::from_package(pages.package().clone());
    text.set_paragraph_drop_cap(storage_id, ParagraphStart::ZERO, model(3, 1))
        .unwrap();
    text.remove_paragraph_drop_cap(storage_id, ParagraphStart::ZERO)
        .unwrap();

    let archive = text.package().archive(&archive_name).unwrap();
    let after = archive.object(storage_id).unwrap();
    assert_eq!(after.archive_info, before_info);
    assert_eq!(after.messages, before_messages);
}
