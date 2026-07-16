use prost::Message;

use super::*;
use crate::archive::RawMessage;
use crate::keynote::KeynoteDocumentBuilder;
use crate::numbers::NumbersDocumentBuilder;
use crate::pages::PagesEditor;
use crate::protobuf::tswp;
use crate::shapes::{DrawablePoint, DrawableSize, RgbColorSpace, RgbaColor};
use crate::text::{
    IWorkTextEditor, ParagraphIndentPoints, ParagraphIndents, ParagraphLineSpacingMultiple,
    ParagraphLineSpacingPoints, ParagraphSpacing, ParagraphSpacingPoints, ParagraphTabAlignment,
    ParagraphTabLeader, ParagraphTabPosition, ParagraphTabStop, ParagraphTabStops,
    TextCapitalization, TextColumnCount, TextColumns, TextDecorations, TextPointSize,
    TextStrikethrough, TextStyle, TextUnderline,
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
fn native_decoration_overrides_are_minimal_and_reversible() {
    let overrides = ParagraphStyleOverrides {
        bold: Some(true),
        underline: Some(TextUnderline::Wavy),
        strikethrough: Some(TextStrikethrough::Triple),
        alignment: Some(TextAlignment::Center),
        ..Default::default()
    };
    let object = native::variation_object(10, 11, 12, overrides.clone()).unwrap();
    let message = &object.messages[0];
    let archive = tswp::ParagraphStyleArchive::decode(message.data.as_slice()).unwrap();
    assert_eq!(archive.override_count, Some(4));
    assert_eq!(
        native::direct_overrides(&archive, &message.data).unwrap(),
        Some(overrides)
    );
}

#[test]
fn native_text_color_override_is_canonical_and_reversible() {
    let color = RgbaColor::new(0.12, 0.48, 0.91, 0.8, RgbColorSpace::DisplayP3).unwrap();
    let overrides = ParagraphStyleOverrides {
        font_color: Some(color),
        ..Default::default()
    };
    let object = native::variation_object(20, 21, 22, overrides.clone()).unwrap();
    let message = &object.messages[0];
    let archive = tswp::ParagraphStyleArchive::decode(message.data.as_slice()).unwrap();
    assert_eq!(archive.override_count, Some(1));
    assert!(
        archive
            .char_properties
            .as_ref()
            .and_then(|properties| properties.tsd_fill.as_ref())
            .and_then(|fill| fill.color.as_ref())
            .is_some()
    );
    assert!(
        object.archive_info.message_infos[0]
            .field_infos
            .iter()
            .any(|field| field.path.path == [11, 46])
    );
    assert_eq!(
        native::direct_overrides(&archive, &message.data).unwrap(),
        Some(overrides)
    );
}

#[test]
fn native_capitalization_values_are_strict_canonical_and_reversible() {
    for capitalization in [
        TextCapitalization::None,
        TextCapitalization::AllCaps,
        TextCapitalization::SmallCaps,
        TextCapitalization::TitleCase,
        TextCapitalization::StartCase,
    ] {
        let overrides = ParagraphStyleOverrides {
            capitalization: Some(capitalization),
            ..Default::default()
        };
        let object = native::variation_object(30, 31, 32, overrides.clone()).unwrap();
        let message = &object.messages[0];
        let archive = tswp::ParagraphStyleArchive::decode(message.data.as_slice()).unwrap();
        assert_eq!(
            archive.override_count,
            Some(capitalization.native_override_count())
        );
        assert_eq!(
            native::direct_overrides(&archive, &message.data).unwrap(),
            Some(overrides)
        );
    }
    assert!(TextCapitalization::from_native_value(4, None).is_err());
    assert!(TextCapitalization::from_native_value(1, Some(true)).is_err());
}

#[test]
fn uniform_capitalization_round_trips_isolates_and_resets_in_every_suite() {
    let mut pages = PagesEditor::create_with_text("Caps").unwrap();
    let pages_box = pages
        .add_text_box(
            4,
            "Pages capitalization",
            DrawablePoint { x: 20.0, y: 40.0 },
            DrawableSize {
                width: 240.0,
                height: 100.0,
            },
        )
        .unwrap();
    let pages_sibling = pages
        .add_text_box(
            5,
            "Plain Pages text",
            DrawablePoint { x: 280.0, y: 40.0 },
            DrawableSize {
                width: 240.0,
                height: 100.0,
            },
        )
        .unwrap();
    let pages_color = RgbaColor::new(0.8, 0.2, 0.1, 1.0, RgbColorSpace::Srgb).unwrap();
    pages
        .set_text_box_text_color(pages_box.drawable_object_id, pages_color)
        .unwrap();
    pages
        .set_text_box_text_capitalization(pages_box.drawable_object_id, TextCapitalization::AllCaps)
        .unwrap();
    assert_eq!(
        pages
            .text_box_text_capitalization(pages_sibling.drawable_object_id)
            .unwrap(),
        TextCapitalization::None
    );
    let mut pages = PagesEditor::from_bytes(&pages.to_bytes().unwrap()).unwrap();
    assert_eq!(
        pages
            .text_box_text_capitalization(pages_box.drawable_object_id)
            .unwrap(),
        TextCapitalization::AllCaps
    );
    assert!(
        pages
            .reset_text_box_text_capitalization(pages_box.drawable_object_id)
            .unwrap()
    );
    assert_eq!(
        pages
            .text_box_text_capitalization(pages_box.drawable_object_id)
            .unwrap(),
        TextCapitalization::None
    );
    assert_eq!(
        pages
            .text_box_text_color(pages_box.drawable_object_id)
            .unwrap(),
        pages_color
    );

    let mut numbers = NumbersDocumentBuilder::new().build().unwrap();
    let sheet_id = numbers.sheets().unwrap()[0].object_id;
    let numbers_box = numbers
        .add_sheet_text_box(
            sheet_id,
            "Numbers capitalization",
            DrawablePoint { x: 20.0, y: 200.0 },
            DrawableSize {
                width: 240.0,
                height: 100.0,
            },
        )
        .unwrap();
    numbers
        .set_sheet_text_box_text_capitalization(
            sheet_id,
            numbers_box.drawable_object_id,
            TextCapitalization::SmallCaps,
        )
        .unwrap();
    let mut numbers =
        crate::numbers::NumbersEditor::from_bytes(&numbers.to_bytes().unwrap()).unwrap();
    assert_eq!(
        numbers
            .sheet_text_box_text_capitalization(sheet_id, numbers_box.drawable_object_id)
            .unwrap(),
        TextCapitalization::SmallCaps
    );
    assert!(
        numbers
            .reset_sheet_text_box_text_capitalization(sheet_id, numbers_box.drawable_object_id)
            .unwrap()
    );

    let mut keynote = KeynoteDocumentBuilder::new().build().unwrap();
    let keynote_box = keynote
        .add_slide_text_box(
            0,
            "Keynote capitalization",
            DrawablePoint { x: 80.0, y: 500.0 },
            DrawableSize {
                width: 500.0,
                height: 100.0,
            },
        )
        .unwrap();
    keynote
        .set_slide_text_box_text_capitalization(
            0,
            keynote_box.drawable_object_id,
            TextCapitalization::TitleCase,
        )
        .unwrap();
    let mut keynote =
        crate::keynote::KeynoteEditor::from_bytes(&keynote.to_bytes().unwrap()).unwrap();
    assert_eq!(
        keynote
            .slide_text_box_text_capitalization(0, keynote_box.drawable_object_id)
            .unwrap(),
        TextCapitalization::TitleCase
    );
    assert!(
        keynote
            .reset_slide_text_box_text_capitalization(0, keynote_box.drawable_object_id)
            .unwrap()
    );
}

#[test]
fn uniform_text_color_round_trips_isolates_and_resets_in_every_suite() {
    let pages_color = RgbaColor::new(0.9, 0.2, 0.1, 1.0, RgbColorSpace::DisplayP3).unwrap();
    let mut pages = PagesEditor::create_with_text("Colors").unwrap();
    let pages_box = pages
        .add_text_box(
            6,
            "Colored Pages text",
            DrawablePoint { x: 20.0, y: 40.0 },
            DrawableSize {
                width: 240.0,
                height: 100.0,
            },
        )
        .unwrap();
    let pages_sibling = pages
        .add_text_box(
            7,
            "Plain Pages text",
            DrawablePoint { x: 280.0, y: 40.0 },
            DrawableSize {
                width: 240.0,
                height: 100.0,
            },
        )
        .unwrap();
    let decorations = TextDecorations::new(TextUnderline::Single, TextStrikethrough::Double);
    pages
        .set_text_box_text_decorations(pages_box.drawable_object_id, decorations)
        .unwrap();
    pages
        .set_text_box_text_color(pages_box.drawable_object_id, pages_color)
        .unwrap();
    assert_eq!(
        pages
            .text_box_text_color(pages_sibling.drawable_object_id)
            .unwrap(),
        RgbaColor::black()
    );
    let mut pages = PagesEditor::from_bytes(&pages.to_bytes().unwrap()).unwrap();
    assert_eq!(
        pages
            .text_box_text_color(pages_box.drawable_object_id)
            .unwrap(),
        pages_color
    );
    assert!(
        pages
            .reset_text_box_text_color(pages_box.drawable_object_id)
            .unwrap()
    );
    assert_eq!(
        pages
            .text_box_text_color(pages_box.drawable_object_id)
            .unwrap(),
        RgbaColor::black()
    );
    assert_eq!(
        pages
            .text_box_text_decorations(pages_box.drawable_object_id)
            .unwrap(),
        decorations
    );

    let numbers_color = RgbaColor::new(0.2, 0.8, 0.3, 1.0, RgbColorSpace::Srgb).unwrap();
    let mut numbers = NumbersDocumentBuilder::new().build().unwrap();
    let sheet_id = numbers.sheets().unwrap()[0].object_id;
    let numbers_box = numbers
        .add_sheet_text_box(
            sheet_id,
            "Colored Numbers text",
            DrawablePoint { x: 20.0, y: 200.0 },
            DrawableSize {
                width: 240.0,
                height: 100.0,
            },
        )
        .unwrap();
    numbers
        .set_sheet_text_box_text_color(sheet_id, numbers_box.drawable_object_id, numbers_color)
        .unwrap();
    let mut numbers =
        crate::numbers::NumbersEditor::from_bytes(&numbers.to_bytes().unwrap()).unwrap();
    assert_eq!(
        numbers
            .sheet_text_box_text_color(sheet_id, numbers_box.drawable_object_id)
            .unwrap(),
        numbers_color
    );
    assert!(
        numbers
            .reset_sheet_text_box_text_color(sheet_id, numbers_box.drawable_object_id)
            .unwrap()
    );

    let keynote_color = RgbaColor::new(0.1, 0.5, 1.0, 0.75, RgbColorSpace::Srgb).unwrap();
    let mut keynote = KeynoteDocumentBuilder::new().build().unwrap();
    let keynote_box = keynote
        .add_slide_text_box(
            0,
            "Colored Keynote text",
            DrawablePoint { x: 80.0, y: 500.0 },
            DrawableSize {
                width: 500.0,
                height: 100.0,
            },
        )
        .unwrap();
    keynote
        .set_slide_text_box_text_color(0, keynote_box.drawable_object_id, keynote_color)
        .unwrap();
    let mut keynote =
        crate::keynote::KeynoteEditor::from_bytes(&keynote.to_bytes().unwrap()).unwrap();
    assert_eq!(
        keynote
            .slide_text_box_text_color(0, keynote_box.drawable_object_id)
            .unwrap(),
        keynote_color
    );
    assert!(
        keynote
            .reset_slide_text_box_text_color(0, keynote_box.drawable_object_id)
            .unwrap()
    );
}

#[test]
fn uniform_text_decorations_round_trip_and_reset_independently_in_every_suite() {
    let pages_decorations = TextDecorations::new(TextUnderline::Single, TextStrikethrough::Double);
    let mut pages = PagesEditor::create_with_text("Decorations").unwrap();
    let pages_box = pages
        .add_text_box(
            11,
            "Decorated Pages text",
            DrawablePoint { x: 20.0, y: 40.0 },
            DrawableSize {
                width: 240.0,
                height: 100.0,
            },
        )
        .unwrap();
    let pages_sibling = pages
        .add_text_box(
            12,
            "Plain Pages text",
            DrawablePoint { x: 280.0, y: 40.0 },
            DrawableSize {
                width: 240.0,
                height: 100.0,
            },
        )
        .unwrap();
    let pages_style = TextStyle::new(TextPointSize::from_points(18.0).unwrap()).with_bold(true);
    pages
        .set_text_box_text_style(pages_box.drawable_object_id, pages_style)
        .unwrap();
    pages
        .set_text_box_text_decorations(pages_box.drawable_object_id, pages_decorations)
        .unwrap();
    assert_eq!(
        pages
            .text_box_text_decorations(pages_sibling.drawable_object_id)
            .unwrap(),
        TextDecorations::NONE
    );
    let mut pages = PagesEditor::from_bytes(&pages.to_bytes().unwrap()).unwrap();
    assert_eq!(
        pages
            .text_box_text_decorations(pages_box.drawable_object_id)
            .unwrap(),
        pages_decorations
    );
    assert!(
        pages
            .reset_text_box_text_decorations(pages_box.drawable_object_id)
            .unwrap()
    );
    assert_eq!(
        pages
            .text_box_text_decorations(pages_box.drawable_object_id)
            .unwrap(),
        TextDecorations::NONE
    );
    assert_eq!(
        pages
            .text_box_text_style(pages_box.drawable_object_id)
            .unwrap(),
        pages_style
    );

    let numbers_decorations =
        TextDecorations::new(TextUnderline::Double, TextStrikethrough::Triple);
    let mut numbers = NumbersDocumentBuilder::new().build().unwrap();
    let sheet_id = numbers.sheets().unwrap()[0].object_id;
    let numbers_box = numbers
        .add_sheet_text_box(
            sheet_id,
            "Decorated Numbers text",
            DrawablePoint { x: 20.0, y: 200.0 },
            DrawableSize {
                width: 240.0,
                height: 100.0,
            },
        )
        .unwrap();
    numbers
        .set_sheet_text_box_paragraph_alignment(
            sheet_id,
            numbers_box.drawable_object_id,
            TextAlignment::Right,
        )
        .unwrap();
    numbers
        .set_sheet_text_box_text_decorations(
            sheet_id,
            numbers_box.drawable_object_id,
            numbers_decorations,
        )
        .unwrap();
    let mut numbers =
        crate::numbers::NumbersEditor::from_bytes(&numbers.to_bytes().unwrap()).unwrap();
    assert_eq!(
        numbers
            .sheet_text_box_text_decorations(sheet_id, numbers_box.drawable_object_id)
            .unwrap(),
        numbers_decorations
    );
    assert!(
        numbers
            .reset_sheet_text_box_text_decorations(sheet_id, numbers_box.drawable_object_id)
            .unwrap()
    );
    assert_eq!(
        numbers
            .sheet_text_box_paragraph_alignment(sheet_id, numbers_box.drawable_object_id)
            .unwrap(),
        TextAlignment::Right
    );

    let keynote_decorations = TextDecorations::new(TextUnderline::Wavy, TextStrikethrough::Single);
    let mut keynote = KeynoteDocumentBuilder::new().build().unwrap();
    let keynote_box = keynote
        .add_slide_text_box(
            0,
            "Decorated Keynote text",
            DrawablePoint { x: 80.0, y: 500.0 },
            DrawableSize {
                width: 500.0,
                height: 100.0,
            },
        )
        .unwrap();
    keynote
        .set_slide_text_box_text_decorations(0, keynote_box.drawable_object_id, keynote_decorations)
        .unwrap();
    let mut keynote =
        crate::keynote::KeynoteEditor::from_bytes(&keynote.to_bytes().unwrap()).unwrap();
    assert_eq!(
        keynote
            .slide_text_box_text_decorations(0, keynote_box.drawable_object_id)
            .unwrap(),
        keynote_decorations
    );
    assert!(
        keynote
            .reset_slide_text_box_text_decorations(0, keynote_box.drawable_object_id)
            .unwrap()
    );
    assert_eq!(
        keynote
            .slide_text_box_text_decorations(0, keynote_box.drawable_object_id)
            .unwrap(),
        TextDecorations::NONE
    );
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
    assert!(
        editor
            .set_text_decorations(
                storage_id,
                TextDecorations::new(TextUnderline::Single, TextStrikethrough::Single),
            )
            .is_err()
    );
    assert!(
        editor
            .set_text_color(
                storage_id,
                RgbaColor::new(0.2, 0.4, 0.8, 1.0, RgbColorSpace::Srgb).unwrap(),
            )
            .is_err()
    );
    assert!(
        editor
            .set_text_capitalization(storage_id, TextCapitalization::StartCase)
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);
}
