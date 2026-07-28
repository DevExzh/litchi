use prost::Message;

use super::*;
use crate::archive::RawMessage;
use crate::keynote::KeynoteDocumentBuilder;
use crate::numbers::NumbersDocumentBuilder;
use crate::pages::PagesEditor;
use crate::protobuf::tsp;
use crate::protobuf::tswp::{self, object_attribute_table::ObjectAttribute};
use crate::shapes::{DrawablePoint, DrawableSize};
use crate::text::{IWorkTextEditor, ParagraphStart};

const POSITION: DrawablePoint = DrawablePoint { x: 40.0, y: 80.0 };
const SIZE: DrawableSize = DrawableSize {
    width: 320.0,
    height: 140.0,
};

#[test]
fn canonical_lists_round_trip_update_and_reset_in_every_suite() {
    let mut pages = PagesEditor::create_with_text("Lists").unwrap();
    let pages_box = pages
        .add_text_box(5, "First\nSecond", POSITION, SIZE)
        .unwrap();
    let pages_sibling = pages
        .add_text_box(5, "Plain", DrawablePoint { x: 400.0, y: 80.0 }, SIZE)
        .unwrap();
    assert_eq!(
        pages
            .text_box_paragraph_list(pages_box.drawable_object_id)
            .unwrap(),
        ParagraphList::None
    );
    pages
        .set_text_box_paragraph_list(pages_box.drawable_object_id, ParagraphList::Bullet)
        .unwrap();
    assert_eq!(
        pages
            .text_box_paragraph_list(pages_sibling.drawable_object_id)
            .unwrap(),
        ParagraphList::None
    );
    let mut pages = PagesEditor::from_bytes(&pages.to_bytes().unwrap()).unwrap();
    assert_eq!(
        pages
            .text_box_paragraph_list(pages_box.drawable_object_id)
            .unwrap(),
        ParagraphList::Bullet
    );
    pages
        .set_text_box_paragraph_list(pages_box.drawable_object_id, ParagraphList::Numbered)
        .unwrap();
    assert!(
        pages
            .reset_text_box_paragraph_list(pages_box.drawable_object_id)
            .unwrap()
    );
    assert!(
        !pages
            .reset_text_box_paragraph_list(pages_box.drawable_object_id)
            .unwrap()
    );

    let mut numbers = NumbersDocumentBuilder::new().build().unwrap();
    let sheet_id = numbers.sheets().unwrap()[0].object_id;
    let numbers_box = numbers
        .add_sheet_text_box(sheet_id, "One\nTwo", POSITION, SIZE)
        .unwrap();
    numbers
        .set_sheet_text_box_paragraph_list(
            sheet_id,
            numbers_box.drawable_object_id,
            ParagraphList::Numbered,
        )
        .unwrap();
    let mut numbers =
        crate::numbers::NumbersEditor::from_bytes(&numbers.to_bytes().unwrap()).unwrap();
    assert_eq!(
        numbers
            .sheet_text_box_paragraph_list(sheet_id, numbers_box.drawable_object_id)
            .unwrap(),
        ParagraphList::Numbered
    );
    assert!(
        numbers
            .reset_sheet_text_box_paragraph_list(sheet_id, numbers_box.drawable_object_id)
            .unwrap()
    );

    let mut keynote = KeynoteDocumentBuilder::new().build().unwrap();
    let keynote_box = keynote
        .add_slide_text_box(0, "Alpha\nBeta", POSITION, SIZE)
        .unwrap();
    keynote
        .set_slide_text_box_paragraph_list(0, keynote_box.drawable_object_id, ParagraphList::Bullet)
        .unwrap();
    let mut keynote =
        crate::keynote::KeynoteEditor::from_bytes(&keynote.to_bytes().unwrap()).unwrap();
    assert_eq!(
        keynote
            .slide_text_box_paragraph_list(0, keynote_box.drawable_object_id)
            .unwrap(),
        ParagraphList::Bullet
    );
    assert!(
        keynote
            .reset_slide_text_box_paragraph_list(0, keynote_box.drawable_object_id)
            .unwrap()
    );
}

#[test]
fn text_boxes_round_trip_custom_bullets_in_every_suite() {
    let paragraph = ParagraphStart::from_utf16_index(6).unwrap();
    let arrow = ParagraphListBullet::new("➡").unwrap();

    let mut pages = PagesEditor::create_with_text("Bullets").unwrap();
    let pages_box = pages
        .add_text_box(5, "First\nSecond", POSITION, SIZE)
        .unwrap();
    pages
        .set_text_box_paragraph_list(pages_box.drawable_object_id, ParagraphList::Bullet)
        .unwrap();
    let pages_before_custom_bullet = pages.to_bytes().unwrap();
    pages
        .set_text_box_paragraph_list_bullet(pages_box.drawable_object_id, paragraph, &arrow)
        .unwrap();
    let mut pages = PagesEditor::from_bytes(&pages.to_bytes().unwrap()).unwrap();
    assert_eq!(
        pages
            .text_box_paragraph_list_bullet(pages_box.drawable_object_id, paragraph)
            .unwrap(),
        arrow
    );
    assert!(
        pages
            .reset_text_box_paragraph_list_bullet(pages_box.drawable_object_id, paragraph)
            .unwrap()
    );
    assert_eq!(pages.to_bytes().unwrap(), pages_before_custom_bullet);

    let mut numbers = NumbersDocumentBuilder::new().build().unwrap();
    let sheet_id = numbers.sheets().unwrap()[0].object_id;
    let numbers_box = numbers
        .add_sheet_text_box(sheet_id, "First\nSecond", POSITION, SIZE)
        .unwrap();
    numbers
        .set_sheet_text_box_paragraph_list(
            sheet_id,
            numbers_box.drawable_object_id,
            ParagraphList::Bullet,
        )
        .unwrap();
    numbers
        .set_sheet_text_box_paragraph_list_bullet(
            sheet_id,
            numbers_box.drawable_object_id,
            paragraph,
            &arrow,
        )
        .unwrap();
    let numbers = crate::numbers::NumbersEditor::from_bytes(&numbers.to_bytes().unwrap()).unwrap();
    assert_eq!(
        numbers
            .sheet_text_box_paragraph_list_bullet(
                sheet_id,
                numbers_box.drawable_object_id,
                paragraph,
            )
            .unwrap(),
        arrow
    );

    let mut keynote = KeynoteDocumentBuilder::new().build().unwrap();
    let keynote_box = keynote
        .add_slide_text_box(0, "First\nSecond", POSITION, SIZE)
        .unwrap();
    keynote
        .set_slide_text_box_paragraph_list(0, keynote_box.drawable_object_id, ParagraphList::Bullet)
        .unwrap();
    keynote
        .set_slide_text_box_paragraph_list_bullet(
            0,
            keynote_box.drawable_object_id,
            paragraph,
            &arrow,
        )
        .unwrap();
    let keynote = crate::keynote::KeynoteEditor::from_bytes(&keynote.to_bytes().unwrap()).unwrap();
    assert_eq!(
        keynote
            .slide_text_box_paragraph_list_bullet(0, keynote_box.drawable_object_id, paragraph,)
            .unwrap(),
        arrow
    );
}

#[test]
fn custom_bullet_updates_reject_nonbullet_and_invalid_paragraphs_transactionally() {
    let arrow = ParagraphListBullet::new("➡").unwrap();
    let mut pages = PagesEditor::create_with_text("Bullets").unwrap();
    let text_box = pages
        .add_text_box(5, "First\nSecond", POSITION, SIZE)
        .unwrap();
    let before = pages.to_bytes().unwrap();
    assert!(
        pages
            .set_text_box_paragraph_list_bullet(
                text_box.drawable_object_id,
                ParagraphStart::ZERO,
                &arrow,
            )
            .is_err()
    );
    assert_eq!(pages.to_bytes().unwrap(), before);

    pages
        .set_text_box_paragraph_list(text_box.drawable_object_id, ParagraphList::Bullet)
        .unwrap();
    let before = pages.to_bytes().unwrap();
    assert!(
        pages
            .set_text_box_paragraph_list_bullet(
                text_box.drawable_object_id,
                ParagraphStart::from_utf16_index(1).unwrap(),
                &arrow,
            )
            .is_err()
    );
    assert_eq!(pages.to_bytes().unwrap(), before);
}

#[test]
fn multiple_list_boundaries_are_rejected_transactionally() {
    let mut pages = PagesEditor::create_with_text("Lists").unwrap();
    let text_box = pages
        .add_text_box(5, "First\nSecond", POSITION, SIZE)
        .unwrap();
    let storage_id = text_box.storage.object_id;
    let mut package = pages.into_package();
    let location = storage::locate(&package, storage_id).unwrap();
    package
        .update_archive(&location.archive_name, |archive| {
            let object = archive.object_mut(storage_id).unwrap();
            let index = object
                .messages
                .iter()
                .position(|message| {
                    matches!(message.type_, 2_001 | 2_022)
                        && tswp::StorageArchive::decode(message.data.as_slice()).is_ok()
                })
                .unwrap();
            let original = &object.messages[index];
            let mut storage = tswp::StorageArchive::decode(original.data.as_slice()).unwrap();
            storage
                .table_list_style
                .as_mut()
                .unwrap()
                .entries
                .push(ObjectAttribute {
                    character_index: 6,
                    object: Some(tsp::Reference {
                        identifier: location.style_id,
                        ..Default::default()
                    }),
                });
            object.replace_message(
                index,
                RawMessage {
                    type_: original.type_,
                    data: storage.encode_to_vec(),
                },
            )?;
            object.archive_info.message_infos[index]
                .object_references
                .push(location.style_id);
            Ok(())
        })
        .unwrap();
    let mut editor = IWorkTextEditor::from_package(package);
    let before = editor.to_bytes().unwrap();
    assert!(
        editor
            .set_paragraph_list(storage_id, ParagraphList::Bullet)
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);
}
