use prost::Message;

use super::*;
use crate::archive::RawMessage;
use crate::keynote::KeynoteDocumentBuilder;
use crate::numbers::NumbersDocumentBuilder;
use crate::pages::PagesEditor;
use crate::protobuf::tswp;
use crate::shapes::{DrawablePoint, DrawableSize};
use crate::text::{
    IWorkTextEditor, ParagraphIndentPoints, ParagraphIndents, ParagraphLineSpacingMultiple,
    ParagraphLineSpacingPoints, ParagraphSpacing, ParagraphSpacingPoints, ParagraphTabAlignment,
    ParagraphTabLeader, ParagraphTabPosition, ParagraphTabStop, ParagraphTabStops, TextColumnCount,
    TextColumns, TextPointSize, TextStyle,
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
fn uniform_text_style_round_trips_and_resets_independently_in_every_suite() {
    let pages_style = TextStyle::new(TextPointSize::from_points(19.5).unwrap()).with_bold(true);
    let mut pages = PagesEditor::create_with_text("Character formatting").unwrap();
    let pages_box = pages
        .add_text_box(
            20,
            "Pages style",
            DrawablePoint { x: 20.0, y: 40.0 },
            DrawableSize {
                width: 240.0,
                height: 100.0,
            },
        )
        .unwrap();
    let pages_sibling = pages
        .add_text_box(
            21,
            "Unchanged Pages style",
            DrawablePoint { x: 280.0, y: 40.0 },
            DrawableSize {
                width: 240.0,
                height: 100.0,
            },
        )
        .unwrap();
    let pages_inherited = pages
        .text_box_text_style(pages_box.drawable_object_id)
        .unwrap();
    pages
        .set_text_box_paragraph_alignment(pages_box.drawable_object_id, TextAlignment::Center)
        .unwrap();
    pages
        .set_text_box_text_style(pages_box.drawable_object_id, pages_style)
        .unwrap();
    assert_eq!(
        pages
            .text_box_text_style(pages_sibling.drawable_object_id)
            .unwrap(),
        pages_inherited
    );
    let mut pages = PagesEditor::from_bytes(&pages.to_bytes().unwrap()).unwrap();
    assert_eq!(
        pages
            .text_box_text_style(pages_box.drawable_object_id)
            .unwrap(),
        pages_style
    );
    assert!(
        pages
            .reset_text_box_text_style(pages_box.drawable_object_id)
            .unwrap()
    );
    assert_eq!(
        pages
            .text_box_text_style(pages_box.drawable_object_id)
            .unwrap(),
        pages_inherited
    );
    assert_eq!(
        pages
            .text_box_paragraph_alignment(pages_box.drawable_object_id)
            .unwrap(),
        TextAlignment::Center
    );

    let numbers_style = TextStyle::new(TextPointSize::from_points(21.0).unwrap()).with_italic(true);
    let mut numbers = NumbersDocumentBuilder::new().build().unwrap();
    let sheet_id = numbers.sheets().unwrap()[0].object_id;
    let numbers_box = numbers
        .add_sheet_text_box(
            sheet_id,
            "Numbers style",
            DrawablePoint { x: 20.0, y: 200.0 },
            DrawableSize {
                width: 240.0,
                height: 100.0,
            },
        )
        .unwrap();
    let numbers_inherited = numbers
        .sheet_text_box_text_style(sheet_id, numbers_box.drawable_object_id)
        .unwrap();
    numbers
        .set_sheet_text_box_paragraph_alignment(
            sheet_id,
            numbers_box.drawable_object_id,
            TextAlignment::Right,
        )
        .unwrap();
    numbers
        .set_sheet_text_box_text_style(sheet_id, numbers_box.drawable_object_id, numbers_style)
        .unwrap();
    let mut numbers =
        crate::numbers::NumbersEditor::from_bytes(&numbers.to_bytes().unwrap()).unwrap();
    assert_eq!(
        numbers
            .sheet_text_box_text_style(sheet_id, numbers_box.drawable_object_id)
            .unwrap(),
        numbers_style
    );
    assert!(
        numbers
            .reset_sheet_text_box_text_style(sheet_id, numbers_box.drawable_object_id)
            .unwrap()
    );
    assert_eq!(
        numbers
            .sheet_text_box_text_style(sheet_id, numbers_box.drawable_object_id)
            .unwrap(),
        numbers_inherited
    );
    assert_eq!(
        numbers
            .sheet_text_box_paragraph_alignment(sheet_id, numbers_box.drawable_object_id)
            .unwrap(),
        TextAlignment::Right
    );

    let keynote_style = TextStyle::new(TextPointSize::from_points(23.0).unwrap())
        .with_bold(true)
        .with_italic(true);
    let mut keynote = KeynoteDocumentBuilder::new().build().unwrap();
    let keynote_box = keynote
        .add_slide_text_box(
            0,
            "Keynote style",
            DrawablePoint { x: 80.0, y: 500.0 },
            DrawableSize {
                width: 500.0,
                height: 100.0,
            },
        )
        .unwrap();
    let keynote_inherited = keynote
        .slide_text_box_text_style(0, keynote_box.drawable_object_id)
        .unwrap();
    keynote
        .set_slide_text_box_paragraph_alignment(
            0,
            keynote_box.drawable_object_id,
            TextAlignment::Justified,
        )
        .unwrap();
    keynote
        .set_slide_text_box_text_style(0, keynote_box.drawable_object_id, keynote_style)
        .unwrap();
    let mut keynote =
        crate::keynote::KeynoteEditor::from_bytes(&keynote.to_bytes().unwrap()).unwrap();
    assert_eq!(
        keynote
            .slide_text_box_text_style(0, keynote_box.drawable_object_id)
            .unwrap(),
        keynote_style
    );
    assert!(
        keynote
            .reset_slide_text_box_text_style(0, keynote_box.drawable_object_id)
            .unwrap()
    );
    assert_eq!(
        keynote
            .slide_text_box_text_style(0, keynote_box.drawable_object_id)
            .unwrap(),
        keynote_inherited
    );
    assert_eq!(
        keynote
            .slide_text_box_paragraph_alignment(0, keynote_box.drawable_object_id)
            .unwrap(),
        TextAlignment::Justified
    );
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
    let paragraph_spacing = ParagraphSpacing::new(
        ParagraphSpacingPoints::from_points(9.0).unwrap(),
        ParagraphSpacingPoints::from_points(15.0).unwrap(),
    );
    let indents = ParagraphIndents::new(
        ParagraphIndentPoints::from_points(26.0).unwrap(),
        ParagraphIndentPoints::from_points(12.5).unwrap(),
        ParagraphIndentPoints::from_points(12.0).unwrap(),
    );
    let tab_stops = ParagraphTabStops::new(vec![
        ParagraphTabStop::new(
            ParagraphTabPosition::from_points(48.5).unwrap(),
            ParagraphTabAlignment::Left,
        ),
        ParagraphTabStop::new(
            ParagraphTabPosition::from_points(56.0).unwrap(),
            ParagraphTabAlignment::Center,
        )
        .with_leader(ParagraphTabLeader::new(".").unwrap()),
    ])
    .unwrap();

    editor
        .set_text_box_paragraph_alignment(first.drawable_object_id, TextAlignment::Center)
        .unwrap();
    editor
        .set_text_box_paragraph_line_spacing(first.drawable_object_id, spacing)
        .unwrap();
    editor
        .set_text_box_paragraph_spacing(first.drawable_object_id, paragraph_spacing)
        .unwrap();
    editor
        .set_text_box_paragraph_indents(first.drawable_object_id, indents)
        .unwrap();
    editor
        .set_text_box_paragraph_tab_stops(first.drawable_object_id, tab_stops.clone())
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
    assert_eq!(
        editor
            .text_box_paragraph_spacing(second.drawable_object_id)
            .unwrap(),
        ParagraphSpacing::NONE
    );
    assert_eq!(
        editor
            .text_box_paragraph_indents(second.drawable_object_id)
            .unwrap(),
        ParagraphIndents::NONE
    );
    assert!(
        editor
            .text_box_paragraph_tab_stops(second.drawable_object_id)
            .unwrap()
            .is_empty()
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
        reopened
            .text_box_paragraph_spacing(first.drawable_object_id)
            .unwrap(),
        paragraph_spacing
    );
    assert_eq!(
        reopened
            .text_box_paragraph_indents(first.drawable_object_id)
            .unwrap(),
        indents
    );
    assert_eq!(
        reopened
            .text_box_paragraph_tab_stops(first.drawable_object_id)
            .unwrap(),
        tab_stops
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
    assert_eq!(
        reopened
            .text_box_paragraph_spacing(first.drawable_object_id)
            .unwrap(),
        paragraph_spacing
    );
    assert_eq!(
        reopened
            .text_box_paragraph_indents(first.drawable_object_id)
            .unwrap(),
        indents
    );
    assert_eq!(
        reopened
            .text_box_paragraph_tab_stops(first.drawable_object_id)
            .unwrap(),
        tab_stops
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
        reopened
            .reset_text_box_paragraph_spacing(first.drawable_object_id)
            .unwrap()
    );
    assert_eq!(
        reopened
            .text_box_paragraph_spacing(first.drawable_object_id)
            .unwrap(),
        ParagraphSpacing::NONE
    );
    assert_eq!(
        reopened
            .text_box_paragraph_indents(first.drawable_object_id)
            .unwrap(),
        indents
    );
    assert!(
        reopened
            .reset_text_box_paragraph_indents(first.drawable_object_id)
            .unwrap()
    );
    assert_eq!(
        reopened
            .text_box_paragraph_indents(first.drawable_object_id)
            .unwrap(),
        ParagraphIndents::NONE
    );
    assert!(
        reopened
            .reset_text_box_paragraph_tab_stops(first.drawable_object_id)
            .unwrap()
    );
    assert!(
        reopened
            .text_box_paragraph_tab_stops(first.drawable_object_id)
            .unwrap()
            .is_empty()
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
    let paragraph_spacing = ParagraphSpacing::new(
        ParagraphSpacingPoints::from_points(11.0).unwrap(),
        ParagraphSpacingPoints::from_points(17.0).unwrap(),
    );
    let indents = ParagraphIndents::new(
        ParagraphIndentPoints::from_points(23.0).unwrap(),
        ParagraphIndentPoints::from_points(13.0).unwrap(),
        ParagraphIndentPoints::from_points(2.833_333_3).unwrap(),
    );
    let tab_stops = ParagraphTabStops::new(vec![
        ParagraphTabStop::new(
            ParagraphTabPosition::from_points(43.0).unwrap(),
            ParagraphTabAlignment::Right,
        )
        .with_leader(ParagraphTabLeader::new("-").unwrap()),
    ])
    .unwrap();
    editor
        .set_sheet_text_box_paragraph_spacing(
            sheet_id,
            created.drawable_object_id,
            paragraph_spacing,
        )
        .unwrap();
    editor
        .set_sheet_text_box_paragraph_indents(sheet_id, created.drawable_object_id, indents)
        .unwrap();
    editor
        .set_sheet_text_box_paragraph_tab_stops(
            sheet_id,
            created.drawable_object_id,
            tab_stops.clone(),
        )
        .unwrap();
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
        assert_eq!(
            reopened
                .sheet_text_box_paragraph_spacing(sheet_id, created.drawable_object_id)
                .unwrap(),
            paragraph_spacing
        );
        assert_eq!(
            reopened
                .sheet_text_box_paragraph_indents(sheet_id, created.drawable_object_id)
                .unwrap(),
            indents
        );
        assert_eq!(
            reopened
                .sheet_text_box_paragraph_tab_stops(sheet_id, created.drawable_object_id)
                .unwrap(),
            tab_stops
        );
    }
    assert!(
        editor
            .reset_sheet_text_box_paragraph_line_spacing(sheet_id, created.drawable_object_id)
            .unwrap()
    );
    assert!(
        editor
            .reset_sheet_text_box_paragraph_spacing(sheet_id, created.drawable_object_id)
            .unwrap()
    );
    assert!(
        editor
            .reset_sheet_text_box_paragraph_indents(sheet_id, created.drawable_object_id)
            .unwrap()
    );
    assert!(
        editor
            .reset_sheet_text_box_paragraph_tab_stops(sheet_id, created.drawable_object_id)
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
    let paragraph_spacing = ParagraphSpacing::new(
        ParagraphSpacingPoints::from_points(13.0).unwrap(),
        ParagraphSpacingPoints::from_points(19.0).unwrap(),
    );
    let indents = ParagraphIndents::new(
        ParagraphIndentPoints::from_points(23.0).unwrap(),
        ParagraphIndentPoints::from_points(13.0).unwrap(),
        ParagraphIndentPoints::from_points(10.5).unwrap(),
    );
    let tab_stops = ParagraphTabStops::new(vec![
        ParagraphTabStop::new(
            ParagraphTabPosition::from_points(63.0).unwrap(),
            ParagraphTabAlignment::Decimal,
        )
        .with_leader(ParagraphTabLeader::new(".").unwrap()),
    ])
    .unwrap();
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
    editor
        .set_slide_text_box_paragraph_spacing(0, created.drawable_object_id, paragraph_spacing)
        .unwrap();
    editor
        .set_slide_text_box_paragraph_indents(0, created.drawable_object_id, indents)
        .unwrap();
    editor
        .set_slide_text_box_paragraph_tab_stops(0, created.drawable_object_id, tab_stops.clone())
        .unwrap();
    let mut reopened =
        crate::keynote::KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .slide_text_box_paragraph_line_spacing(0, created.drawable_object_id)
            .unwrap(),
        spacing
    );
    assert_eq!(
        reopened
            .slide_text_box_paragraph_spacing(0, created.drawable_object_id)
            .unwrap(),
        paragraph_spacing
    );
    assert_eq!(
        reopened
            .slide_text_box_paragraph_indents(0, created.drawable_object_id)
            .unwrap(),
        indents
    );
    assert_eq!(
        reopened
            .slide_text_box_paragraph_tab_stops(0, created.drawable_object_id)
            .unwrap(),
        tab_stops
    );
    assert!(
        reopened
            .reset_slide_text_box_paragraph_indents(0, created.drawable_object_id)
            .unwrap()
    );
    assert_eq!(
        reopened
            .slide_text_box_paragraph_spacing(0, created.drawable_object_id)
            .unwrap(),
        paragraph_spacing
    );
    assert_eq!(
        reopened
            .slide_text_box_paragraph_tab_stops(0, created.drawable_object_id)
            .unwrap(),
        tab_stops
    );
    assert!(
        reopened
            .reset_slide_text_box_paragraph_spacing(0, created.drawable_object_id)
            .unwrap()
    );
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
    assert!(
        reopened
            .reset_slide_text_box_paragraph_tab_stops(0, created.drawable_object_id)
            .unwrap()
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
            .set_paragraph_tab_stops(
                storage_id,
                ParagraphTabStops::new(vec![ParagraphTabStop::new(
                    ParagraphTabPosition::from_points(48.0).unwrap(),
                    ParagraphTabAlignment::Center,
                )])
                .unwrap(),
            )
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
    assert!(
        editor
            .set_paragraph_spacing(
                storage_id,
                ParagraphSpacing::new(
                    ParagraphSpacingPoints::from_points(8.0).unwrap(),
                    ParagraphSpacingPoints::from_points(12.0).unwrap(),
                ),
            )
            .is_err()
    );
    assert!(
        editor
            .set_paragraph_indents(
                storage_id,
                ParagraphIndents::new(
                    ParagraphIndentPoints::from_points(18.0).unwrap(),
                    ParagraphIndentPoints::from_points(24.0).unwrap(),
                    ParagraphIndentPoints::from_points(12.0).unwrap(),
                ),
            )
            .is_err()
    );
    assert!(
        editor
            .set_text_style(
                storage_id,
                TextStyle::new(TextPointSize::from_points(18.0).unwrap()).with_bold(true),
            )
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);
}
