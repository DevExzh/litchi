use prost::Message;

use super::*;
use crate::archive::RawMessage;
use crate::keynote::KeynoteDocumentBuilder;
use crate::numbers::NumbersDocumentBuilder;
use crate::pages::PagesEditor;
use crate::protobuf::tswp;
use crate::shapes::{DrawablePoint, DrawableSize};
use crate::text::{
    IWorkTextEditor, ParagraphLineSpacingMultiple, ParagraphLineSpacingPoints, TextColumnCount,
    TextColumns,
};

const STORAGE_MESSAGE_TYPES: &[u32] = &[2_001, 2_022];

#[test]
fn all_native_alignment_values_are_strict_and_reversible() {
    for alignment in [
        TextAlignment::Natural,
        TextAlignment::Right,
        TextAlignment::Center,
        TextAlignment::Justified,
        TextAlignment::Left,
    ] {
        assert_eq!(
            TextAlignment::from_native_value(alignment.native_value()).unwrap(),
            alignment
        );
    }
    assert!(TextAlignment::from_native_value(-1).is_err());
    assert!(TextAlignment::from_native_value(5).is_err());
}

#[test]
fn pages_paragraph_overrides_compose_and_reset_independently() {
    let mut editor = PagesEditor::create_with_text("Paragraph style").unwrap();
    let first = editor
        .add_text_box(
            15,
            "First paragraph",
            DrawablePoint { x: 20.0, y: 40.0 },
            DrawableSize {
                width: 240.0,
                height: 100.0,
            },
        )
        .unwrap();
    let second = editor
        .add_text_box(
            16,
            "Second paragraph",
            DrawablePoint { x: 40.0, y: 160.0 },
            DrawableSize {
                width: 240.0,
                height: 100.0,
            },
        )
        .unwrap();
    let columns = TextColumns::equal(TextColumnCount::new(2).unwrap(), None);
    editor
        .set_text_box_columns(first.drawable_object_id, &columns)
        .unwrap();
    let spacing = ParagraphLineSpacing::Relative(ParagraphLineSpacingMultiple::ONE_POINT_FIVE);

    editor
        .set_text_box_paragraph_alignment(first.drawable_object_id, TextAlignment::Center)
        .unwrap();
    editor
        .set_text_box_paragraph_line_spacing(first.drawable_object_id, spacing)
        .unwrap();
    assert_eq!(
        editor
            .text_box_paragraph_alignment(second.drawable_object_id)
            .unwrap(),
        TextAlignment::Natural
    );
    assert_eq!(
        editor
            .text_box_paragraph_line_spacing(second.drawable_object_id)
            .unwrap(),
        ParagraphLineSpacing::default()
    );

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .text_box_paragraph_alignment(first.drawable_object_id)
            .unwrap(),
        TextAlignment::Center
    );
    assert_eq!(
        reopened
            .text_box_paragraph_line_spacing(first.drawable_object_id)
            .unwrap(),
        spacing
    );
    assert_eq!(
        reopened.text_box_columns(first.drawable_object_id).unwrap(),
        columns
    );
    assert!(
        reopened
            .reset_text_box_paragraph_alignment(first.drawable_object_id)
            .unwrap()
    );
    assert_eq!(
        reopened
            .text_box_paragraph_line_spacing(first.drawable_object_id)
            .unwrap(),
        spacing
    );
    assert!(
        reopened
            .reset_text_box_paragraph_line_spacing(first.drawable_object_id)
            .unwrap()
    );
    assert_eq!(
        reopened
            .text_box_paragraph_line_spacing(first.drawable_object_id)
            .unwrap(),
        ParagraphLineSpacing::default()
    );
    assert!(
        !reopened
            .reset_text_box_paragraph_line_spacing(first.drawable_object_id)
            .unwrap()
    );
}

#[test]
fn every_native_line_spacing_mode_round_trips_in_numbers() {
    let mut editor = NumbersDocumentBuilder::new().build().unwrap();
    let sheet_id = editor.sheets().unwrap()[0].object_id;
    let created = editor
        .add_sheet_text_box(
            sheet_id,
            "Spacing modes",
            DrawablePoint { x: 20.0, y: 200.0 },
            DrawableSize {
                width: 240.0,
                height: 100.0,
            },
        )
        .unwrap();
    let twelve = ParagraphLineSpacingPoints::from_points(12.0).unwrap();
    for spacing in [
        ParagraphLineSpacing::AtLeast(twelve),
        ParagraphLineSpacing::Exactly(twelve),
        ParagraphLineSpacing::Maximum(twelve),
        ParagraphLineSpacing::Between(twelve),
        ParagraphLineSpacing::Relative(ParagraphLineSpacingMultiple::DOUBLE),
    ] {
        editor
            .set_sheet_text_box_paragraph_line_spacing(
                sheet_id,
                created.drawable_object_id,
                spacing,
            )
            .unwrap();
        let reopened =
            crate::numbers::NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .sheet_text_box_paragraph_line_spacing(sheet_id, created.drawable_object_id)
                .unwrap(),
            spacing
        );
    }
    assert!(
        editor
            .reset_sheet_text_box_paragraph_line_spacing(sheet_id, created.drawable_object_id)
            .unwrap()
    );
}

#[test]
fn keynote_spacing_reset_preserves_alignment() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let created = editor
        .add_slide_text_box(
            0,
            "Justified paragraph",
            DrawablePoint { x: 80.0, y: 500.0 },
            DrawableSize {
                width: 500.0,
                height: 100.0,
            },
        )
        .unwrap();
    let spacing =
        ParagraphLineSpacing::Between(ParagraphLineSpacingPoints::from_points(6.0).unwrap());
    editor
        .set_slide_text_box_paragraph_alignment(
            0,
            created.drawable_object_id,
            TextAlignment::Justified,
        )
        .unwrap();
    editor
        .set_slide_text_box_paragraph_line_spacing(0, created.drawable_object_id, spacing)
        .unwrap();
    let mut reopened =
        crate::keynote::KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .slide_text_box_paragraph_line_spacing(0, created.drawable_object_id)
            .unwrap(),
        spacing
    );
    assert!(
        reopened
            .reset_slide_text_box_paragraph_line_spacing(0, created.drawable_object_id)
            .unwrap()
    );
    assert_eq!(
        reopened
            .slide_text_box_paragraph_alignment(0, created.drawable_object_id)
            .unwrap(),
        TextAlignment::Justified
    );
}

#[test]
fn multiple_paragraph_boundaries_are_rejected_transactionally() {
    let mut pages = PagesEditor::create_with_text("Alignment").unwrap();
    let created = pages
        .add_text_box(
            9,
            "Two paragraphs",
            DrawablePoint { x: 20.0, y: 40.0 },
            DrawableSize {
                width: 240.0,
                height: 100.0,
            },
        )
        .unwrap();
    let storage_id = created.storage.object_id;
    let mut package = pages.into_package();
    let archive_name = storage::locate(&package, storage_id).unwrap().archive_name;
    package
        .update_archive(&archive_name, |archive| {
            let object = archive.object_mut(storage_id).unwrap();
            let index = object
                .messages
                .iter()
                .position(|message| {
                    STORAGE_MESSAGE_TYPES.contains(&message.type_)
                        && tswp::StorageArchive::decode(message.data.as_slice()).is_ok()
                })
                .unwrap();
            let message_type = object.messages[index].type_;
            let mut text_storage =
                tswp::StorageArchive::decode(object.messages[index].data.as_slice()).unwrap();
            let mut second = text_storage.table_para_style.as_ref().unwrap().entries[0];
            second.character_index = 1;
            text_storage
                .table_para_style
                .as_mut()
                .unwrap()
                .entries
                .push(second);
            object
                .replace_message(
                    index,
                    RawMessage {
                        type_: message_type,
                        data: text_storage.encode_to_vec(),
                    },
                )
                .unwrap();
            Ok(())
        })
        .unwrap();
    let mut editor = IWorkTextEditor::from_package(package);
    let before = editor.to_bytes().unwrap();
    assert!(
        editor
            .set_paragraph_alignment(storage_id, TextAlignment::Center)
            .is_err()
    );
    assert!(
        editor
            .set_paragraph_line_spacing(
                storage_id,
                ParagraphLineSpacing::Relative(ParagraphLineSpacingMultiple::DOUBLE),
            )
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);
}
