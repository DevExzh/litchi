use prost::Message;

use super::*;
use crate::archive::{ArchiveObject, MessageInfo, RawMessage};
use crate::keynote::{KeynoteDocumentBuilder, KeynoteEditor};
use crate::numbers::{NumbersDocumentBuilder, NumbersEditor};
use crate::pages::PagesEditor;
use crate::protobuf::tswp;
use crate::shapes::{
    DrawablePoint, DrawableSize, RgbColorSpace, RgbaColor, ShapeContactShadow, ShapeDropShadow,
    ShapeShadow, ShapeShadowAngle, ShapeShadowAppearance, ShapeShadowBlurRadius, ShapeShadowOffset,
    ShapeShadowOpacity, ShapeShadowPerspective, StrokePattern, StrokeWidth,
};
use crate::text::{
    IWorkTextEditor, ParagraphBorder, ParagraphDecimalTabCharacter, ParagraphDefaultTabInterval,
    ParagraphFollowingStyle, ParagraphHyphenation, ParagraphIndentPoints, ParagraphIndents,
    ParagraphLineSpacingMultiple, ParagraphLineSpacingPoints, ParagraphSpacing,
    ParagraphSpacingPoints, ParagraphStyleId, ParagraphStyleName, ParagraphTabAlignment,
    ParagraphTabLeader, ParagraphTabPosition, ParagraphTabStop, ParagraphTabStops,
    ParagraphWritingDirection, TextBackground, TextBaselineShift, TextCapitalization,
    TextCharacterSpacing, TextDecorations, TextFont, TextLigatures, TextOutline, TextPointSize,
    TextScript, TextShadow, TextStrikethrough, TextStyle, TextUnderline,
};
use litchi_iwa_text::columns::{Columns, Count};
use litchi_iwa_text::paragraph::border::{Offset as BorderOffset, Sides as BorderSides};
use litchi_iwa_text::paragraph::style::raw::{from_native_id, native_id};

const SOURCE_PAGES_OBJECT_TITLE_STYLE_ID: u64 = 121;

#[test]
fn pages_following_paragraph_style_catalog_round_trips_and_resets() {
    let mut pages = PagesEditor::create_with_text("Following style").unwrap();
    let text_box = pages
        .add_text_box(
            15,
            "Press Return",
            DrawablePoint { x: 20.0, y: 40.0 },
            DrawableSize {
                width: 240.0,
                height: 100.0,
            },
        )
        .unwrap();
    let styles = pages
        .text_box_named_paragraph_styles(text_box.drawable_object_id)
        .unwrap();
    assert_eq!(
        styles.iter().map(|style| style.name()).collect::<Vec<_>>(),
        [
            "Title",
            "Subtitle",
            "Heading",
            "Heading 2",
            "Heading 3",
            "Heading Red",
            "Body",
            "Caption",
            "Header & Footer",
            "Footnote",
            "Label",
            "Label Dark",
        ]
    );
    let body = styles.iter().find(|style| style.name() == "Body").unwrap();
    let expected = ParagraphFollowingStyle::Named(body.id());
    let before = pages.to_bytes().unwrap();
    let hidden_object_title =
        ParagraphFollowingStyle::Named(from_native_id(SOURCE_PAGES_OBJECT_TITLE_STYLE_ID).unwrap());
    assert!(
        pages
            .set_text_box_paragraph_following_style(
                text_box.drawable_object_id,
                hidden_object_title,
            )
            .is_err()
    );
    assert_eq!(pages.to_bytes().unwrap(), before);

    pages
        .set_text_box_paragraph_following_style(text_box.drawable_object_id, expected)
        .unwrap();
    let mut pages = PagesEditor::from_bytes(&pages.to_bytes().unwrap()).unwrap();
    assert_eq!(
        pages
            .text_box_paragraph_following_style(text_box.drawable_object_id)
            .unwrap(),
        expected
    );
    assert!(
        pages
            .reset_text_box_paragraph_following_style(text_box.drawable_object_id)
            .unwrap()
    );
    assert_eq!(pages.to_bytes().unwrap(), before);
}

#[test]
fn source_theme_paragraph_presets_are_deduplicated_across_suite_wrappers() {
    let mut numbers = NumbersDocumentBuilder::new().build().unwrap();
    let sheet_id = numbers.sheets().unwrap()[0].object_id;
    let numbers_box = numbers
        .add_sheet_text_box(
            sheet_id,
            "Numbers preset",
            DrawablePoint { x: 20.0, y: 40.0 },
            DrawableSize {
                width: 240.0,
                height: 100.0,
            },
        )
        .unwrap();
    let numbers_text = IWorkTextEditor::from_package(numbers.into_package());
    assert_eq!(
        numbers_text
            .named_paragraph_styles(numbers_box.storage.object_id)
            .unwrap()
            .iter()
            .map(|style| style.name())
            .collect::<Vec<_>>(),
        [
            "Title",
            "Subtitle",
            "Heading",
            "Heading Red",
            "Body",
            "Caption"
        ]
    );

    let mut keynote = KeynoteDocumentBuilder::new().build().unwrap();
    let keynote_box = keynote
        .add_slide_text_box(
            0,
            "Keynote preset",
            DrawablePoint { x: 20.0, y: 40.0 },
            DrawableSize {
                width: 240.0,
                height: 100.0,
            },
        )
        .unwrap();
    let keynote_text = IWorkTextEditor::from_package(keynote.into_package());
    assert_eq!(
        keynote_text
            .named_paragraph_styles(keynote_box.storage.object_id)
            .unwrap()
            .iter()
            .map(|style| style.name())
            .collect::<Vec<_>>(),
        [
            "Title",
            "Title Small",
            "Subtitle",
            "Subtitle Alt",
            "Body",
            "Caption",
            "Quote",
            "Attribution",
            "Heading",
            "Agenda",
            "Statement",
            "Section",
            "Fact",
        ]
    );
}

#[test]
fn named_paragraph_style_crud_is_transactional_across_all_suites() {
    let mut pages = PagesEditor::create_with_text("Pages style catalog").unwrap();
    let pages_box = pages
        .add_text_box(
            19,
            "Pages",
            DrawablePoint { x: 20.0, y: 40.0 },
            DrawableSize {
                width: 240.0,
                height: 100.0,
            },
        )
        .unwrap();
    let body = pages
        .text_box_named_paragraph_styles(pages_box.drawable_object_id)
        .unwrap()
        .into_iter()
        .find(|style| style.name() == "Body")
        .unwrap()
        .id();
    let initial_pages = pages.to_bytes().unwrap();
    let created = pages
        .create_text_box_named_paragraph_style(
            pages_box.drawable_object_id,
            body,
            ParagraphStyleName::new("Litchi Heading").unwrap(),
        )
        .unwrap();
    assert_eq!(created.name(), "Litchi Heading");
    let before_duplicate = pages.to_bytes().unwrap();
    assert!(
        pages
            .create_text_box_named_paragraph_style(
                pages_box.drawable_object_id,
                body,
                ParagraphStyleName::new("Litchi Heading").unwrap(),
            )
            .is_err()
    );
    assert_eq!(pages.to_bytes().unwrap(), before_duplicate);
    let renamed = pages
        .rename_text_box_named_paragraph_style(
            pages_box.drawable_object_id,
            created.id(),
            ParagraphStyleName::new("Litchi Display").unwrap(),
        )
        .unwrap();
    assert_eq!(renamed.id(), created.id());
    assert_eq!(renamed.name(), "Litchi Display");
    let before_duplicate_rename = pages.to_bytes().unwrap();
    assert!(
        pages
            .rename_text_box_named_paragraph_style(
                pages_box.drawable_object_id,
                created.id(),
                ParagraphStyleName::new("Body").unwrap(),
            )
            .is_err()
    );
    assert_eq!(pages.to_bytes().unwrap(), before_duplicate_rename);
    let pages = PagesEditor::from_bytes(&before_duplicate_rename).unwrap();
    assert_eq!(
        pages
            .text_box_named_paragraph_styles(pages_box.drawable_object_id)
            .unwrap()
            .iter()
            .map(|style| style.name())
            .collect::<Vec<_>>(),
        [
            "Title",
            "Subtitle",
            "Heading",
            "Heading 2",
            "Heading 3",
            "Heading Red",
            "Body",
            "Caption",
            "Header & Footer",
            "Footnote",
            "Label",
            "Label Dark",
            "Litchi Display",
        ]
    );
    let mut pages = pages;
    pages
        .set_text_box_paragraph_following_style(
            pages_box.drawable_object_id,
            ParagraphFollowingStyle::Named(created.id()),
        )
        .unwrap();
    let before_used_delete = pages.to_bytes().unwrap();
    assert!(
        pages
            .delete_text_box_named_paragraph_style(pages_box.drawable_object_id, created.id(),)
            .is_err()
    );
    assert_eq!(pages.to_bytes().unwrap(), before_used_delete);
    assert!(
        pages
            .reset_text_box_paragraph_following_style(pages_box.drawable_object_id,)
            .unwrap()
    );
    let deleted = pages
        .delete_text_box_named_paragraph_style(pages_box.drawable_object_id, created.id())
        .unwrap();
    assert_eq!(deleted, renamed);
    assert_eq!(pages.to_bytes().unwrap(), initial_pages);
    let before_last_delete = pages.to_bytes().unwrap();
    assert!(
        pages
            .delete_text_box_named_paragraph_style(pages_box.drawable_object_id, body,)
            .is_err()
    );
    assert_eq!(pages.to_bytes().unwrap(), before_last_delete);

    let mut numbers = NumbersDocumentBuilder::new().build().unwrap();
    let sheet_id = numbers.sheets().unwrap()[0].object_id;
    let numbers_box = numbers
        .add_sheet_text_box(
            sheet_id,
            "Numbers",
            DrawablePoint { x: 20.0, y: 40.0 },
            DrawableSize {
                width: 240.0,
                height: 100.0,
            },
        )
        .unwrap();
    let mut numbers_text = IWorkTextEditor::from_package(numbers.into_package());
    let initial_numbers = numbers_text.package().to_bytes().unwrap();
    let body = numbers_text
        .named_paragraph_styles(numbers_box.storage.object_id)
        .unwrap()[0]
        .id();
    let numbers_created = numbers_text
        .create_named_paragraph_style(
            numbers_box.storage.object_id,
            body,
            ParagraphStyleName::new("Numbers Detail").unwrap(),
        )
        .unwrap();
    let numbers_renamed = numbers_text
        .rename_named_paragraph_style(
            numbers_box.storage.object_id,
            numbers_created.id(),
            ParagraphStyleName::new("Numbers Summary").unwrap(),
        )
        .unwrap();
    let numbers_bytes = numbers_text.package().to_bytes().unwrap();
    let numbers_text =
        IWorkTextEditor::from_package(IWorkPackage::from_bytes(&numbers_bytes).unwrap());
    assert_eq!(
        numbers_text
            .named_paragraph_styles(numbers_box.storage.object_id)
            .unwrap()
            .last()
            .unwrap()
            .name(),
        "Numbers Summary"
    );
    let mut numbers_text = numbers_text;
    assert_eq!(
        numbers_text
            .delete_named_paragraph_style(numbers_box.storage.object_id, numbers_created.id(),)
            .unwrap(),
        numbers_renamed
    );
    assert_eq!(numbers_text.package().to_bytes().unwrap(), initial_numbers);

    let mut keynote = KeynoteDocumentBuilder::new().build().unwrap();
    let keynote_box = keynote
        .add_slide_text_box(
            0,
            "Keynote",
            DrawablePoint { x: 20.0, y: 40.0 },
            DrawableSize {
                width: 240.0,
                height: 100.0,
            },
        )
        .unwrap();
    let mut keynote_text = IWorkTextEditor::from_package(keynote.into_package());
    let initial_keynote = keynote_text.package().to_bytes().unwrap();
    let body = keynote_text
        .named_paragraph_styles(keynote_box.storage.object_id)
        .unwrap()[0]
        .id();
    let keynote_created = keynote_text
        .create_named_paragraph_style(
            keynote_box.storage.object_id,
            body,
            ParagraphStyleName::new("Keynote Detail").unwrap(),
        )
        .unwrap();
    let keynote_renamed = keynote_text
        .rename_named_paragraph_style(
            keynote_box.storage.object_id,
            keynote_created.id(),
            ParagraphStyleName::new("Keynote Summary").unwrap(),
        )
        .unwrap();
    let keynote_bytes = keynote_text.package().to_bytes().unwrap();
    let keynote_text =
        IWorkTextEditor::from_package(IWorkPackage::from_bytes(&keynote_bytes).unwrap());
    assert_eq!(
        keynote_text
            .named_paragraph_styles(keynote_box.storage.object_id)
            .unwrap()
            .last()
            .unwrap()
            .name(),
        "Keynote Summary"
    );
    let mut keynote_text = keynote_text;
    assert_eq!(
        keynote_text
            .delete_named_paragraph_style(keynote_box.storage.object_id, keynote_created.id(),)
            .unwrap(),
        keynote_renamed
    );
    assert_eq!(keynote_text.package().to_bytes().unwrap(), initial_keynote);
}

#[test]
fn named_paragraph_style_application_and_replacement_deletion_are_transactional() {
    let mut pages = PagesEditor::create_with_text("Pages style application").unwrap();
    let first = pages
        .add_text_box(
            23,
            "First",
            DrawablePoint { x: 20.0, y: 40.0 },
            DrawableSize {
                width: 240.0,
                height: 100.0,
            },
        )
        .unwrap();
    let second = pages
        .add_text_box(
            24,
            "Second",
            DrawablePoint { x: 20.0, y: 160.0 },
            DrawableSize {
                width: 240.0,
                height: 100.0,
            },
        )
        .unwrap();
    let body = pages
        .text_box_applied_named_paragraph_style(first.drawable_object_id)
        .unwrap()
        .style()
        .id();
    assert!(
        !pages
            .text_box_applied_named_paragraph_style(first.drawable_object_id)
            .unwrap()
            .has_overrides()
    );
    let initial_pages = pages.to_bytes().unwrap();
    let display = pages
        .create_text_box_named_paragraph_style(
            first.drawable_object_id,
            body,
            ParagraphStyleName::new("Pages Display").unwrap(),
        )
        .unwrap();
    for object_id in [first.drawable_object_id, second.drawable_object_id] {
        assert_eq!(
            pages
                .apply_text_box_named_paragraph_style(object_id, display.id())
                .unwrap(),
            display
        );
        assert_eq!(
            pages
                .text_box_applied_named_paragraph_style(object_id)
                .unwrap()
                .style(),
            &display
        );
    }

    let before_shared_delete = pages.to_bytes().unwrap();
    assert!(
        pages
            .delete_applied_text_box_named_paragraph_style_with_replacement(
                first.drawable_object_id,
                display.id(),
                body,
            )
            .is_err()
    );
    assert_eq!(pages.to_bytes().unwrap(), before_shared_delete);
    pages
        .apply_text_box_named_paragraph_style(second.drawable_object_id, body)
        .unwrap();
    pages
        .set_text_box_paragraph_alignment(first.drawable_object_id, TextAlignment::Center)
        .unwrap();
    pages
        .set_text_box_paragraph_default_tab_interval(
            first.drawable_object_id,
            ParagraphDefaultTabInterval::from_points(36.0).unwrap(),
        )
        .unwrap();
    assert!(
        pages
            .text_box_applied_named_paragraph_style(first.drawable_object_id)
            .unwrap()
            .has_overrides()
    );
    let before_same_replacement = pages.to_bytes().unwrap();
    assert!(
        pages
            .delete_applied_text_box_named_paragraph_style_with_replacement(
                first.drawable_object_id,
                display.id(),
                display.id(),
            )
            .is_err()
    );
    assert_eq!(pages.to_bytes().unwrap(), before_same_replacement);
    assert_eq!(
        pages
            .delete_applied_text_box_named_paragraph_style_with_replacement(
                first.drawable_object_id,
                display.id(),
                body,
            )
            .unwrap(),
        display
    );
    assert_eq!(pages.to_bytes().unwrap(), initial_pages);

    let mut numbers = NumbersDocumentBuilder::new().build().unwrap();
    let sheet_id = numbers.sheets().unwrap()[0].object_id;
    let numbers_box = numbers
        .add_sheet_text_box(
            sheet_id,
            "Numbers",
            DrawablePoint { x: 20.0, y: 40.0 },
            DrawableSize {
                width: 240.0,
                height: 100.0,
            },
        )
        .unwrap();
    assert_numbers_style_lifecycle(&mut numbers, sheet_id, numbers_box.drawable_object_id);

    let mut keynote = KeynoteDocumentBuilder::new().build().unwrap();
    let keynote_box = keynote
        .add_slide_text_box(
            0,
            "Keynote",
            DrawablePoint { x: 20.0, y: 40.0 },
            DrawableSize {
                width: 240.0,
                height: 100.0,
            },
        )
        .unwrap();
    assert_keynote_style_lifecycle(&mut keynote, keynote_box.drawable_object_id);
}

fn assert_numbers_style_lifecycle(
    editor: &mut NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
) {
    let body = editor
        .sheet_text_box_applied_named_paragraph_style(sheet_id, drawable_object_id)
        .unwrap()
        .style()
        .id();
    let initial = editor.package().to_bytes().unwrap();
    let display = editor
        .create_sheet_text_box_named_paragraph_style(
            sheet_id,
            drawable_object_id,
            body,
            ParagraphStyleName::new("Numbers Draft").unwrap(),
        )
        .unwrap();
    let display = editor
        .rename_sheet_text_box_named_paragraph_style(
            sheet_id,
            drawable_object_id,
            display.id(),
            ParagraphStyleName::new("Numbers Display").unwrap(),
        )
        .unwrap();
    let before_disposable = editor.to_bytes().unwrap();
    let disposable = editor
        .create_sheet_text_box_named_paragraph_style(
            sheet_id,
            drawable_object_id,
            body,
            ParagraphStyleName::new("Numbers Disposable").unwrap(),
        )
        .unwrap();
    editor
        .delete_sheet_text_box_named_paragraph_style(sheet_id, drawable_object_id, disposable.id())
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), before_disposable);
    editor
        .apply_sheet_text_box_named_paragraph_style(sheet_id, drawable_object_id, display.id())
        .unwrap();
    editor
        .set_sheet_text_box_paragraph_alignment(sheet_id, drawable_object_id, TextAlignment::Right)
        .unwrap();
    editor
        .set_sheet_text_box_paragraph_default_tab_interval(
            sheet_id,
            drawable_object_id,
            ParagraphDefaultTabInterval::from_points(36.0).unwrap(),
        )
        .unwrap();
    let selected = editor
        .sheet_text_box_applied_named_paragraph_style(sheet_id, drawable_object_id)
        .unwrap();
    assert_eq!(selected.style(), &display);
    assert!(selected.has_overrides());
    assert_eq!(
        editor
            .delete_applied_sheet_text_box_named_paragraph_style_with_replacement(
                sheet_id,
                drawable_object_id,
                display.id(),
                body,
            )
            .unwrap(),
        display
    );
    assert_eq!(editor.package().to_bytes().unwrap(), initial);
}

fn assert_keynote_style_lifecycle(editor: &mut KeynoteEditor, drawable_object_id: u64) {
    let body = editor
        .slide_text_box_applied_named_paragraph_style(0, drawable_object_id)
        .unwrap()
        .style()
        .id();
    let initial = editor.package().to_bytes().unwrap();
    let display = editor
        .create_slide_text_box_named_paragraph_style(
            0,
            drawable_object_id,
            body,
            ParagraphStyleName::new("Keynote Draft").unwrap(),
        )
        .unwrap();
    let display = editor
        .rename_slide_text_box_named_paragraph_style(
            0,
            drawable_object_id,
            display.id(),
            ParagraphStyleName::new("Keynote Display").unwrap(),
        )
        .unwrap();
    let before_disposable = editor.to_bytes().unwrap();
    let disposable = editor
        .create_slide_text_box_named_paragraph_style(
            0,
            drawable_object_id,
            body,
            ParagraphStyleName::new("Keynote Disposable").unwrap(),
        )
        .unwrap();
    editor
        .delete_slide_text_box_named_paragraph_style(0, drawable_object_id, disposable.id())
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), before_disposable);
    editor
        .apply_slide_text_box_named_paragraph_style(0, drawable_object_id, display.id())
        .unwrap();
    editor
        .set_slide_text_box_paragraph_alignment(0, drawable_object_id, TextAlignment::Right)
        .unwrap();
    editor
        .set_slide_text_box_paragraph_default_tab_interval(
            0,
            drawable_object_id,
            ParagraphDefaultTabInterval::from_points(36.0).unwrap(),
        )
        .unwrap();
    let selected = editor
        .slide_text_box_applied_named_paragraph_style(0, drawable_object_id)
        .unwrap();
    assert_eq!(selected.style(), &display);
    assert!(selected.has_overrides());
    assert_eq!(
        editor
            .delete_applied_slide_text_box_named_paragraph_style_with_replacement(
                0,
                drawable_object_id,
                display.id(),
                body,
            )
            .unwrap(),
        display
    );
    assert_eq!(editor.package().to_bytes().unwrap(), initial);
}

#[test]
fn named_paragraph_style_redefinition_propagates_across_suites() {
    let mut pages = PagesEditor::create_with_text("Pages redefine").unwrap();
    let first = pages
        .add_text_box(
            0,
            "First",
            DrawablePoint { x: 20.0, y: 40.0 },
            DrawableSize {
                width: 240.0,
                height: 100.0,
            },
        )
        .unwrap();
    let second = pages
        .add_text_box(
            0,
            "Second",
            DrawablePoint { x: 20.0, y: 160.0 },
            DrawableSize {
                width: 240.0,
                height: 100.0,
            },
        )
        .unwrap();
    let body = pages
        .text_box_applied_named_paragraph_style(first.drawable_object_id)
        .unwrap()
        .style()
        .id();
    let display = pages
        .create_text_box_named_paragraph_style(
            first.drawable_object_id,
            body,
            ParagraphStyleName::new("Pages Redefined").unwrap(),
        )
        .unwrap();
    for object_id in [first.drawable_object_id, second.drawable_object_id] {
        pages
            .apply_text_box_named_paragraph_style(object_id, display.id())
            .unwrap();
    }
    let unchanged = pages.to_bytes().unwrap();
    assert!(
        pages
            .redefine_applied_text_box_named_paragraph_style(first.drawable_object_id)
            .is_err()
    );
    assert_eq!(pages.to_bytes().unwrap(), unchanged);
    pages
        .set_text_box_paragraph_alignment(first.drawable_object_id, TextAlignment::Center)
        .unwrap();
    assert_eq!(
        pages
            .redefine_applied_text_box_named_paragraph_style(first.drawable_object_id)
            .unwrap(),
        display
    );
    let pages = PagesEditor::from_bytes(&pages.to_bytes().unwrap()).unwrap();
    for object_id in [first.drawable_object_id, second.drawable_object_id] {
        assert_eq!(
            pages.text_box_paragraph_alignment(object_id).unwrap(),
            TextAlignment::Center
        );
        assert!(
            !pages
                .text_box_applied_named_paragraph_style(object_id)
                .unwrap()
                .has_overrides()
        );
    }

    let mut numbers = NumbersDocumentBuilder::new().build().unwrap();
    let sheet_id = numbers.sheets().unwrap()[0].object_id;
    let first = numbers
        .add_sheet_text_box(
            sheet_id,
            "First",
            DrawablePoint { x: 20.0, y: 40.0 },
            DrawableSize {
                width: 240.0,
                height: 100.0,
            },
        )
        .unwrap();
    let second = numbers
        .add_sheet_text_box(
            sheet_id,
            "Second",
            DrawablePoint { x: 20.0, y: 160.0 },
            DrawableSize {
                width: 240.0,
                height: 100.0,
            },
        )
        .unwrap();
    let body = numbers
        .sheet_text_box_applied_named_paragraph_style(sheet_id, first.drawable_object_id)
        .unwrap()
        .style()
        .id();
    let display = numbers
        .create_sheet_text_box_named_paragraph_style(
            sheet_id,
            first.drawable_object_id,
            body,
            ParagraphStyleName::new("Numbers Redefined").unwrap(),
        )
        .unwrap();
    for object_id in [first.drawable_object_id, second.drawable_object_id] {
        numbers
            .apply_sheet_text_box_named_paragraph_style(sheet_id, object_id, display.id())
            .unwrap();
    }
    numbers
        .set_sheet_text_box_paragraph_alignment(
            sheet_id,
            first.drawable_object_id,
            TextAlignment::Right,
        )
        .unwrap();
    numbers
        .redefine_applied_sheet_text_box_named_paragraph_style(sheet_id, first.drawable_object_id)
        .unwrap();
    let numbers = NumbersEditor::from_bytes(&numbers.to_bytes().unwrap()).unwrap();
    for object_id in [first.drawable_object_id, second.drawable_object_id] {
        assert_eq!(
            numbers
                .sheet_text_box_paragraph_alignment(sheet_id, object_id)
                .unwrap(),
            TextAlignment::Right
        );
    }

    let mut keynote = KeynoteDocumentBuilder::new().build().unwrap();
    let first = keynote
        .add_slide_text_box(
            0,
            "First",
            DrawablePoint { x: 20.0, y: 40.0 },
            DrawableSize {
                width: 240.0,
                height: 100.0,
            },
        )
        .unwrap();
    let second = keynote
        .add_slide_text_box(
            0,
            "Second",
            DrawablePoint { x: 20.0, y: 160.0 },
            DrawableSize {
                width: 240.0,
                height: 100.0,
            },
        )
        .unwrap();
    let body = keynote
        .slide_text_box_applied_named_paragraph_style(0, first.drawable_object_id)
        .unwrap()
        .style()
        .id();
    let display = keynote
        .create_slide_text_box_named_paragraph_style(
            0,
            first.drawable_object_id,
            body,
            ParagraphStyleName::new("Keynote Redefined").unwrap(),
        )
        .unwrap();
    for object_id in [first.drawable_object_id, second.drawable_object_id] {
        keynote
            .apply_slide_text_box_named_paragraph_style(0, object_id, display.id())
            .unwrap();
    }
    keynote
        .set_slide_text_box_paragraph_alignment(
            0,
            first.drawable_object_id,
            TextAlignment::Justified,
        )
        .unwrap();
    keynote
        .redefine_applied_slide_text_box_named_paragraph_style(0, first.drawable_object_id)
        .unwrap();
    let keynote = KeynoteEditor::from_bytes(&keynote.to_bytes().unwrap()).unwrap();
    for object_id in [first.drawable_object_id, second.drawable_object_id] {
        assert_eq!(
            keynote
                .slide_text_box_paragraph_alignment(0, object_id)
                .unwrap(),
            TextAlignment::Justified
        );
    }
}

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
fn native_writing_direction_override_is_minimal_and_reversible() {
    let overrides = ParagraphStyleOverrides {
        writing_direction: Some(ParagraphWritingDirection::RightToLeft),
        ..Default::default()
    };
    let object = native::variation_object(13, 14, 15, overrides.clone()).unwrap();
    let message = &object.messages[0];
    let archive = tswp::ParagraphStyleArchive::decode(message.data.as_slice()).unwrap();
    assert_eq!(archive.override_count, Some(1));
    assert_eq!(
        archive
            .char_properties
            .as_ref()
            .and_then(|properties| properties.writing_direction),
        Some(crate::text::paragraph_direction::to_native(
            ParagraphWritingDirection::RightToLeft,
        ))
    );
    assert_eq!(
        archive
            .para_properties
            .as_ref()
            .and_then(|properties| properties.writing_direction),
        None
    );
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
fn native_text_font_overrides_are_canonical_strict_and_reversible() {
    for font in [
        TextFont::Default,
        TextFont::named("AvenirNext-BoldItalic").unwrap(),
    ] {
        let overrides = ParagraphStyleOverrides {
            font: Some(font.clone()),
            ..Default::default()
        };
        let object = native::variation_object(23, 24, 25, overrides.clone()).unwrap();
        let message = &object.messages[0];
        let archive = tswp::ParagraphStyleArchive::decode(message.data.as_slice()).unwrap();
        let properties = archive.char_properties.as_ref().unwrap();
        assert_eq!(archive.override_count, Some(1));
        match font {
            TextFont::Default => {
                assert_eq!(properties.font_name_null, Some(true));
                assert!(properties.font_name.is_none());
            },
            TextFont::Named(name) => {
                assert!(properties.font_name_null.is_none());
                assert_eq!(properties.font_name.as_deref(), Some(name.as_str()));
            },
        }
        assert_eq!(
            native::direct_overrides(&archive, &message.data).unwrap(),
            Some(overrides)
        );
    }

    let both = tswp::CharacterStylePropertiesArchive {
        font_name_null: Some(true),
        font_name: Some("Helvetica".to_owned()),
        ..Default::default()
    };
    assert!(native::text_font_from_character(&both).is_err());
    let false_without_name = tswp::CharacterStylePropertiesArchive {
        font_name_null: Some(false),
        ..Default::default()
    };
    assert!(native::text_font_from_character(&false_without_name).is_err());
    let invalid_name = tswp::CharacterStylePropertiesArchive {
        font_name: Some(" Helvetica".to_owned()),
        ..Default::default()
    };
    assert!(native::text_font_from_character(&invalid_name).is_err());
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
fn native_script_values_are_strict_canonical_and_reversible() {
    for script in [
        TextScript::Normal,
        TextScript::Superscript,
        TextScript::Subscript,
    ] {
        let overrides = ParagraphStyleOverrides {
            script: Some(script),
            ..Default::default()
        };
        let object = native::variation_object(33, 34, 35, overrides.clone()).unwrap();
        let message = &object.messages[0];
        let archive = tswp::ParagraphStyleArchive::decode(message.data.as_slice()).unwrap();
        assert_eq!(archive.override_count, Some(1));
        assert_eq!(
            archive
                .char_properties
                .as_ref()
                .and_then(|properties| properties.superscript),
            Some(script.native_value())
        );
        assert_eq!(
            native::direct_overrides(&archive, &message.data).unwrap(),
            Some(overrides)
        );
    }
    assert!(TextScript::from_native_value(-1).is_err());
    assert!(TextScript::from_native_value(3).is_err());
}

#[test]
fn native_baseline_shifts_are_strict_canonical_and_reversible() {
    assert!(TextBaselineShift::from_points(f32::NAN).is_err());
    assert!(TextBaselineShift::from_points(f32::INFINITY).is_err());
    assert!(TextBaselineShift::from_points(f32::NEG_INFINITY).is_err());

    for shift in [
        TextBaselineShift::ZERO,
        TextBaselineShift::from_points(-3.0).unwrap(),
        TextBaselineShift::from_points(5.0).unwrap(),
    ] {
        let overrides = ParagraphStyleOverrides {
            baseline_shift: Some(shift),
            ..Default::default()
        };
        let object = native::variation_object(36, 37, 38, overrides.clone()).unwrap();
        let message = &object.messages[0];
        let archive = tswp::ParagraphStyleArchive::decode(message.data.as_slice()).unwrap();
        assert_eq!(archive.override_count, Some(1));
        assert_eq!(
            archive
                .char_properties
                .as_ref()
                .and_then(|properties| properties.baseline_shift),
            Some(shift.points())
        );
        assert_eq!(
            native::direct_overrides(&archive, &message.data).unwrap(),
            Some(overrides)
        );
    }
}

#[test]
fn native_character_spacing_is_bounded_canonical_and_reversible() {
    assert!(TextCharacterSpacing::from_percent(f32::NAN).is_err());
    assert!(TextCharacterSpacing::from_percent(f32::INFINITY).is_err());
    assert!(TextCharacterSpacing::from_percent(-40.01).is_err());
    assert!(TextCharacterSpacing::from_percent(400.01).is_err());
    assert!(TextCharacterSpacing::from_native_ratio(-0.401).is_err());
    assert!(TextCharacterSpacing::from_native_ratio(4.001).is_err());

    for spacing in [
        TextCharacterSpacing::MINIMUM,
        TextCharacterSpacing::NORMAL,
        TextCharacterSpacing::from_percent(12.0).unwrap(),
        TextCharacterSpacing::MAXIMUM,
    ] {
        let overrides = ParagraphStyleOverrides {
            character_spacing: Some(spacing),
            ..Default::default()
        };
        let object = native::variation_object(39, 40, 41, overrides.clone()).unwrap();
        let message = &object.messages[0];
        let archive = tswp::ParagraphStyleArchive::decode(message.data.as_slice()).unwrap();
        assert_eq!(archive.override_count, Some(1));
        assert_eq!(
            archive
                .char_properties
                .as_ref()
                .and_then(|properties| properties.tracking),
            Some(spacing.native_ratio())
        );
        assert_eq!(
            native::direct_overrides(&archive, &message.data).unwrap(),
            Some(overrides)
        );
    }
}

#[test]
fn native_ligature_values_are_strict_canonical_and_reversible() {
    for ligatures in [
        TextLigatures::RequiredOnly,
        TextLigatures::Standard,
        TextLigatures::All,
    ] {
        let overrides = ParagraphStyleOverrides {
            ligatures: Some(ligatures),
            ..Default::default()
        };
        let object = native::variation_object(42, 43, 44, overrides.clone()).unwrap();
        let message = &object.messages[0];
        let archive = tswp::ParagraphStyleArchive::decode(message.data.as_slice()).unwrap();
        assert_eq!(archive.override_count, Some(1));
        assert_eq!(
            archive
                .char_properties
                .as_ref()
                .and_then(|properties| properties.ligatures),
            Some(ligatures.native_value())
        );
        assert_eq!(
            native::direct_overrides(&archive, &message.data).unwrap(),
            Some(overrides)
        );
        assert_eq!(
            TextLigatures::from_native_value(ligatures.native_value()).unwrap(),
            ligatures
        );
    }
    assert!(TextLigatures::from_native_value(-1).is_err());
    assert!(TextLigatures::from_native_value(3).is_err());
}

#[test]
fn native_text_outline_overrides_are_canonical_strict_and_reversible() {
    for outline in [TextOutline::None, TextOutline::standard()] {
        let overrides = ParagraphStyleOverrides {
            outline: Some(outline),
            ..Default::default()
        };
        let object = native::variation_object(45, 46, 47, overrides.clone()).unwrap();
        let message = &object.messages[0];
        let archive = tswp::ParagraphStyleArchive::decode(message.data.as_slice()).unwrap();
        let properties = archive.char_properties.as_ref().unwrap();
        assert_eq!(archive.override_count, Some(1));
        match outline {
            TextOutline::None => {
                assert_eq!(properties.tsd_stroke_null, Some(true));
                assert!(properties.tsd_stroke.is_none());
            },
            TextOutline::Stroke(_) => {
                assert!(properties.tsd_stroke_null.is_none());
                assert!(properties.tsd_stroke.is_some());
                assert!(
                    object.archive_info.message_infos[0]
                        .field_infos
                        .iter()
                        .any(|field| field.path.path == [11, 44])
                );
            },
        }
        assert_eq!(
            native::direct_overrides(&archive, &message.data).unwrap(),
            Some(overrides)
        );
    }

    let malformed = tswp::CharacterStylePropertiesArchive {
        tsd_stroke_null: Some(true),
        tsd_stroke: Some(crate::shapes::stroke_to_native(
            match TextOutline::standard() {
                TextOutline::Stroke(stroke) => stroke,
                TextOutline::None => unreachable!(),
            },
        )),
        ..Default::default()
    };
    assert!(native::text_outline_from_character(&malformed).is_err());
}

#[test]
fn native_text_shadow_overrides_are_canonical_strict_and_reversible() {
    let custom = TextShadow::Drop(ShapeDropShadow::new(
        ShapeShadowAppearance::new(
            RgbaColor::new(0.1, 0.3, 0.8, 1.0, RgbColorSpace::DisplayP3).unwrap(),
            ShapeShadowBlurRadius::from_points(7).unwrap(),
            ShapeShadowOffset::from_points(11.0).unwrap(),
            ShapeShadowOpacity::new(0.42).unwrap(),
        ),
        ShapeShadowAngle::from_degrees(135.0).unwrap(),
    ));
    for shadow in [TextShadow::None, TextShadow::standard(), custom] {
        let overrides = ParagraphStyleOverrides {
            shadow: Some(shadow),
            ..Default::default()
        };
        let object = native::variation_object(48, 49, 50, overrides.clone()).unwrap();
        let message = &object.messages[0];
        let archive = tswp::ParagraphStyleArchive::decode(message.data.as_slice()).unwrap();
        let properties = archive.char_properties.as_ref().unwrap();
        assert_eq!(archive.override_count, Some(1));
        match shadow {
            TextShadow::None => {
                assert_eq!(properties.shadow_null, Some(true));
                assert!(properties.shadow.is_none());
            },
            TextShadow::Drop(_) => {
                assert!(properties.shadow_null.is_none());
                assert!(properties.shadow.is_some());
                assert!(
                    object.archive_info.message_infos[0]
                        .field_infos
                        .iter()
                        .any(|field| field.path.path == [11, 21])
                );
            },
        }
        assert_eq!(
            native::direct_overrides(&archive, &message.data).unwrap(),
            Some(overrides)
        );
    }

    let malformed = tswp::CharacterStylePropertiesArchive {
        shadow_null: Some(true),
        shadow: Some(crate::shapes::shadow_to_native(
            TextShadow::standard().into_shape_shadow(),
        )),
        ..Default::default()
    };
    assert!(native::text_shadow_from_character(&malformed).is_err());

    let contact = ShapeShadow::Contact(ShapeContactShadow::new(
        ShapeShadowAppearance::new(
            RgbaColor::black(),
            ShapeShadowBlurRadius::ONE_POINT,
            ShapeShadowOffset::FIVE_POINTS,
            ShapeShadowOpacity::OPAQUE,
        ),
        ShapeShadowPerspective::LEVEL,
    ));
    assert!(TextShadow::from_shape_shadow(contact).is_err());
}

#[test]
fn native_text_background_overrides_are_canonical_strict_and_reversible() {
    let custom = TextBackground::Color(
        RgbaColor::new(0.18, 0.72, 0.32, 0.65, RgbColorSpace::DisplayP3).unwrap(),
    );
    for background in [TextBackground::None, custom] {
        let overrides = ParagraphStyleOverrides {
            background: Some(background),
            ..Default::default()
        };
        let object = native::variation_object(51, 52, 53, overrides.clone()).unwrap();
        let message = &object.messages[0];
        let archive = tswp::ParagraphStyleArchive::decode(message.data.as_slice()).unwrap();
        let properties = archive.char_properties.as_ref().unwrap();
        assert_eq!(archive.override_count, Some(1));
        match background {
            TextBackground::None => {
                assert_eq!(properties.background_color_null, Some(true));
                assert!(properties.background_color.is_none());
            },
            TextBackground::Color(_) => {
                assert!(properties.background_color_null.is_none());
                assert!(properties.background_color.is_some());
                assert!(
                    object.archive_info.message_infos[0]
                        .field_infos
                        .iter()
                        .any(|field| field.path.path == [11, 26])
                );
            },
        }
        assert_eq!(
            native::direct_overrides(&archive, &message.data).unwrap(),
            Some(overrides)
        );
    }

    let both = tswp::CharacterStylePropertiesArchive {
        background_color_null: Some(true),
        background_color: Some(crate::shapes::color_to_native(RgbaColor::black())),
        ..Default::default()
    };
    assert!(native::text_background_from_character(&both).is_err());
    let false_without_color = tswp::CharacterStylePropertiesArchive {
        background_color_null: Some(false),
        ..Default::default()
    };
    assert!(native::text_background_from_character(&false_without_color).is_err());
}

#[test]
fn native_paragraph_background_overrides_match_app_authored_wire() {
    let custom = ParagraphBackground::Color(
        RgbaColor::new(1.0, 0.588_738_74, 0.552_926_2, 1.0, RgbColorSpace::Srgb).unwrap(),
    );
    for background in [ParagraphBackground::None, custom] {
        let overrides = ParagraphStyleOverrides {
            paragraph_background: Some(background),
            ..Default::default()
        };
        let object = native::variation_object(61, 62, 63, overrides.clone()).unwrap();
        let message = &object.messages[0];
        let archive = tswp::ParagraphStyleArchive::decode(message.data.as_slice()).unwrap();
        let properties = archive.para_properties.as_ref().unwrap();
        assert_eq!(archive.override_count, Some(1));
        match background {
            ParagraphBackground::None => {
                assert_eq!(properties.fill_null, Some(true));
                assert!(properties.fill.is_none());
            },
            ParagraphBackground::Color(_) => {
                assert!(properties.fill_null.is_none());
                assert!(properties.fill.is_some());
            },
        }
        assert_eq!(
            native::direct_overrides(&archive, &message.data).unwrap(),
            Some(overrides)
        );
    }

    let both = tswp::ParagraphStylePropertiesArchive {
        fill_null: Some(true),
        fill: Some(crate::shapes::color_to_native(RgbaColor::black())),
        ..Default::default()
    };
    assert!(native::paragraph_background_from_properties(&both).is_err());
    let false_without_color = tswp::ParagraphStylePropertiesArchive {
        fill_null: Some(false),
        ..Default::default()
    };
    assert!(native::paragraph_background_from_properties(&false_without_color).is_err());
}

#[test]
fn native_paragraph_border_overrides_match_app_authored_wire() {
    let border = ParagraphBorder::new(
        RgbaColor::black(),
        StrokeWidth::new(3.0).unwrap(),
        StrokePattern::Solid,
        BorderSides::ALL,
        BorderOffset::from_points(9.0).unwrap(),
        true,
    )
    .unwrap();
    let overrides = ParagraphStyleOverrides {
        paragraph_borders: Some(ParagraphBorders::Bordered(border)),
        ..Default::default()
    };
    let object = native::variation_object(64, 65, 66, overrides.clone()).unwrap();
    let message = &object.messages[0];
    let archive = tswp::ParagraphStyleArchive::decode(message.data.as_slice()).unwrap();
    let properties = archive.para_properties.as_ref().unwrap();
    assert_eq!(archive.override_count, Some(4));
    assert_eq!(properties.deprecated_borders, Some(4));
    assert_eq!(properties.border_positions, Some(15));
    assert_eq!(properties.rounded_corners, Some(true));
    let offset = properties.historical_rule_offset.as_ref().unwrap();
    assert_eq!((offset.x, offset.y), (3.0, 3.0));
    assert!(properties.stroke.is_some());
    assert_eq!(
        native::direct_overrides(&archive, &message.data).unwrap(),
        Some(overrides)
    );

    let none = ParagraphStyleOverrides {
        paragraph_borders: Some(ParagraphBorders::None),
        ..Default::default()
    };
    let object = native::variation_object(67, 68, 69, none.clone()).unwrap();
    let message = &object.messages[0];
    let archive = tswp::ParagraphStyleArchive::decode(message.data.as_slice()).unwrap();
    let properties = archive.para_properties.as_ref().unwrap();
    assert_eq!(archive.override_count, Some(4));
    assert_eq!(properties.deprecated_borders, Some(0));
    assert_eq!(properties.stroke_null, Some(true));
    assert_eq!(properties.border_positions, Some(0));
    assert_eq!(properties.rounded_corners, Some(false));
    assert_eq!(
        native::direct_overrides(&archive, &message.data).unwrap(),
        Some(none)
    );
}

#[test]
fn native_paragraph_border_parser_rejects_conflicting_wire() {
    let mismatched_sides = tswp::ParagraphStylePropertiesArchive {
        deprecated_borders: Some(1),
        border_positions: Some(2),
        ..Default::default()
    };
    assert!(native::paragraph_borders_from_properties(&mismatched_sides).is_err());

    let mismatched_offset = tswp::ParagraphStylePropertiesArchive {
        historical_rule_offset: Some(crate::protobuf::tsp::Point { x: 1.0, y: 2.0 }),
        deprecated_borders: Some(1),
        border_positions: Some(1),
        stroke: Some(crate::shapes::stroke_to_native(
            ParagraphBorder::new(
                RgbaColor::black(),
                StrokeWidth::new(1.0).unwrap(),
                StrokePattern::Solid,
                BorderSides::TOP,
                BorderOffset::DEFAULT,
                false,
            )
            .unwrap()
            .native_stroke(),
        )),
        ..Default::default()
    };
    assert!(native::paragraph_borders_from_properties(&mismatched_offset).is_err());

    let null_and_stroke = tswp::ParagraphStylePropertiesArchive {
        stroke_null: Some(true),
        stroke: Some(crate::shapes::stroke_to_native(
            ParagraphBorder::new(
                RgbaColor::black(),
                StrokeWidth::new(1.0).unwrap(),
                StrokePattern::Solid,
                BorderSides::TOP,
                BorderOffset::DEFAULT,
                false,
            )
            .unwrap()
            .native_stroke(),
        )),
        ..Default::default()
    };
    assert!(native::paragraph_borders_from_properties(&null_and_stroke).is_err());
}

#[test]
fn native_paragraph_flow_overrides_match_app_authored_wire() {
    let flow = ParagraphFlow::new()
        .with_keep_lines_together(true)
        .with_keep_with_next(true)
        .with_start_on_new_page(true)
        .with_prevent_widow_orphan_lines(false)
        .with_hyphenation(ParagraphHyphenation::Prevented);
    let overrides = ParagraphStyleOverrides {
        hyphenation: Some(flow.hyphenation()),
        keep_lines_together: Some(flow.keeps_lines_together()),
        keep_with_next: Some(flow.keeps_with_next()),
        start_on_new_page: Some(flow.starts_on_new_page()),
        prevent_widow_orphan_lines: Some(flow.prevents_widow_orphan_lines()),
        ..Default::default()
    };
    let object = native::variation_object(70, 71, 72, overrides.clone()).unwrap();
    let message = &object.messages[0];
    let archive = tswp::ParagraphStyleArchive::decode(message.data.as_slice()).unwrap();
    let properties = archive.para_properties.as_ref().unwrap();
    assert_eq!(archive.override_count, Some(5));
    assert_eq!(properties.hyphenate, Some(false));
    assert_eq!(properties.keep_lines_together, Some(true));
    assert_eq!(properties.keep_with_next, Some(true));
    assert_eq!(properties.page_break_before, Some(true));
    assert_eq!(properties.widow_control, Some(false));
    assert_eq!(
        native::direct_overrides(&archive, &message.data).unwrap(),
        Some(overrides)
    );
}

#[test]
fn native_paragraph_flow_rejects_noncanonical_boolean_wire() {
    const PARAGRAPH_PROPERTIES_FIELD: u32 = 12;
    const HYPHENATE_FIELD: u32 = 8;
    const NONCANONICAL_TRUE: u8 = 2;

    let overrides = ParagraphStyleOverrides {
        hyphenation: Some(ParagraphHyphenation::Prevented),
        ..Default::default()
    };
    let object = native::variation_object(73, 74, 75, overrides).unwrap();
    let mut data = object.messages[0].data.clone();
    let paragraph = crate::wire::parse_wire_fields(&data)
        .unwrap()
        .into_iter()
        .find(|field| field.number() == PARAGRAPH_PROPERTIES_FIELD)
        .unwrap();
    let hyphenate =
        crate::wire::parse_wire_fields(&data[paragraph.payload_start()..paragraph.end()])
            .unwrap()
            .into_iter()
            .find(|field| field.number() == HYPHENATE_FIELD)
            .unwrap();
    data[paragraph.payload_start() + hyphenate.payload_start()] = NONCANONICAL_TRUE;
    let archive = tswp::ParagraphStyleArchive::decode(data.as_slice()).unwrap();
    assert!(native::direct_overrides(&archive, &data).unwrap().is_none());
}

#[test]
fn uniform_text_script_round_trips_isolates_and_resets_in_every_suite() {
    let mut pages = PagesEditor::create_with_text("Scripts").unwrap();
    let pages_box = pages
        .add_text_box(
            7,
            "Pages superscript",
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
    pages
        .set_text_box_text_capitalization(pages_box.drawable_object_id, TextCapitalization::AllCaps)
        .unwrap();
    pages
        .set_text_box_text_script(pages_box.drawable_object_id, TextScript::Superscript)
        .unwrap();
    assert_eq!(
        pages
            .text_box_text_script(pages_sibling.drawable_object_id)
            .unwrap(),
        TextScript::Normal
    );
    let mut pages = PagesEditor::from_bytes(&pages.to_bytes().unwrap()).unwrap();
    assert_eq!(
        pages
            .text_box_text_script(pages_box.drawable_object_id)
            .unwrap(),
        TextScript::Superscript
    );
    assert!(
        pages
            .reset_text_box_text_script(pages_box.drawable_object_id)
            .unwrap()
    );
    assert_eq!(
        pages
            .text_box_text_script(pages_box.drawable_object_id)
            .unwrap(),
        TextScript::Normal
    );
    assert_eq!(
        pages
            .text_box_text_capitalization(pages_box.drawable_object_id)
            .unwrap(),
        TextCapitalization::AllCaps
    );

    let mut numbers = NumbersDocumentBuilder::new().build().unwrap();
    let sheet_id = numbers.sheets().unwrap()[0].object_id;
    let numbers_box = numbers
        .add_sheet_text_box(
            sheet_id,
            "Numbers subscript",
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
    numbers
        .set_sheet_text_box_text_script(
            sheet_id,
            numbers_box.drawable_object_id,
            TextScript::Subscript,
        )
        .unwrap();
    let mut numbers =
        crate::numbers::NumbersEditor::from_bytes(&numbers.to_bytes().unwrap()).unwrap();
    assert_eq!(
        numbers
            .sheet_text_box_text_script(sheet_id, numbers_box.drawable_object_id)
            .unwrap(),
        TextScript::Subscript
    );
    assert!(
        numbers
            .reset_sheet_text_box_text_script(sheet_id, numbers_box.drawable_object_id)
            .unwrap()
    );
    assert_eq!(
        numbers
            .sheet_text_box_text_capitalization(sheet_id, numbers_box.drawable_object_id)
            .unwrap(),
        TextCapitalization::SmallCaps
    );

    let mut keynote = KeynoteDocumentBuilder::new().build().unwrap();
    let keynote_box = keynote
        .add_slide_text_box(
            0,
            "Keynote superscript",
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
    keynote
        .set_slide_text_box_text_script(0, keynote_box.drawable_object_id, TextScript::Superscript)
        .unwrap();
    let mut keynote =
        crate::keynote::KeynoteEditor::from_bytes(&keynote.to_bytes().unwrap()).unwrap();
    assert_eq!(
        keynote
            .slide_text_box_text_script(0, keynote_box.drawable_object_id)
            .unwrap(),
        TextScript::Superscript
    );
    assert!(
        keynote
            .reset_slide_text_box_text_script(0, keynote_box.drawable_object_id)
            .unwrap()
    );
    assert_eq!(
        keynote
            .slide_text_box_text_capitalization(0, keynote_box.drawable_object_id)
            .unwrap(),
        TextCapitalization::TitleCase
    );
}

#[test]
fn uniform_baseline_shift_round_trips_isolates_and_resets_in_every_suite() {
    let pages_shift = TextBaselineShift::from_points(4.0).unwrap();
    let mut pages = PagesEditor::create_with_text("Baseline shifts").unwrap();
    let pages_box = pages
        .add_text_box(
            7,
            "Raised Pages text",
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
            "Unshifted Pages text",
            DrawablePoint { x: 280.0, y: 40.0 },
            DrawableSize {
                width: 240.0,
                height: 100.0,
            },
        )
        .unwrap();
    pages
        .set_text_box_text_script(pages_box.drawable_object_id, TextScript::Superscript)
        .unwrap();
    pages
        .set_text_box_text_baseline_shift(pages_box.drawable_object_id, pages_shift)
        .unwrap();
    assert_eq!(
        pages
            .text_box_text_baseline_shift(pages_sibling.drawable_object_id)
            .unwrap(),
        TextBaselineShift::ZERO
    );
    let mut pages = PagesEditor::from_bytes(&pages.to_bytes().unwrap()).unwrap();
    assert_eq!(
        pages
            .text_box_text_baseline_shift(pages_box.drawable_object_id)
            .unwrap(),
        pages_shift
    );
    assert!(
        pages
            .reset_text_box_text_baseline_shift(pages_box.drawable_object_id)
            .unwrap()
    );
    assert_eq!(
        pages
            .text_box_text_baseline_shift(pages_box.drawable_object_id)
            .unwrap(),
        TextBaselineShift::ZERO
    );
    assert_eq!(
        pages
            .text_box_text_script(pages_box.drawable_object_id)
            .unwrap(),
        TextScript::Superscript
    );

    let numbers_shift = TextBaselineShift::from_points(-3.0).unwrap();
    let mut numbers = NumbersDocumentBuilder::new().build().unwrap();
    let sheet_id = numbers.sheets().unwrap()[0].object_id;
    let numbers_box = numbers
        .add_sheet_text_box(
            sheet_id,
            "Lowered Numbers text",
            DrawablePoint { x: 20.0, y: 200.0 },
            DrawableSize {
                width: 240.0,
                height: 100.0,
            },
        )
        .unwrap();
    numbers
        .set_sheet_text_box_text_script(
            sheet_id,
            numbers_box.drawable_object_id,
            TextScript::Subscript,
        )
        .unwrap();
    numbers
        .set_sheet_text_box_text_baseline_shift(
            sheet_id,
            numbers_box.drawable_object_id,
            numbers_shift,
        )
        .unwrap();
    let mut numbers =
        crate::numbers::NumbersEditor::from_bytes(&numbers.to_bytes().unwrap()).unwrap();
    assert_eq!(
        numbers
            .sheet_text_box_text_baseline_shift(sheet_id, numbers_box.drawable_object_id)
            .unwrap(),
        numbers_shift
    );
    assert!(
        numbers
            .reset_sheet_text_box_text_baseline_shift(sheet_id, numbers_box.drawable_object_id)
            .unwrap()
    );
    assert_eq!(
        numbers
            .sheet_text_box_text_script(sheet_id, numbers_box.drawable_object_id)
            .unwrap(),
        TextScript::Subscript
    );

    let keynote_shift = TextBaselineShift::from_points(5.0).unwrap();
    let mut keynote = KeynoteDocumentBuilder::new().build().unwrap();
    let keynote_box = keynote
        .add_slide_text_box(
            0,
            "Raised Keynote text",
            DrawablePoint { x: 80.0, y: 500.0 },
            DrawableSize {
                width: 500.0,
                height: 100.0,
            },
        )
        .unwrap();
    keynote
        .set_slide_text_box_text_script(0, keynote_box.drawable_object_id, TextScript::Superscript)
        .unwrap();
    keynote
        .set_slide_text_box_text_baseline_shift(0, keynote_box.drawable_object_id, keynote_shift)
        .unwrap();
    let mut keynote =
        crate::keynote::KeynoteEditor::from_bytes(&keynote.to_bytes().unwrap()).unwrap();
    assert_eq!(
        keynote
            .slide_text_box_text_baseline_shift(0, keynote_box.drawable_object_id)
            .unwrap(),
        keynote_shift
    );
    assert!(
        keynote
            .reset_slide_text_box_text_baseline_shift(0, keynote_box.drawable_object_id)
            .unwrap()
    );
    assert_eq!(
        keynote
            .slide_text_box_text_script(0, keynote_box.drawable_object_id)
            .unwrap(),
        TextScript::Superscript
    );
}

#[test]
fn uniform_character_spacing_round_trips_isolates_and_resets_in_every_suite() {
    let pages_spacing = TextCharacterSpacing::from_percent(12.0).unwrap();
    let pages_shift = TextBaselineShift::from_points(4.0).unwrap();
    let mut pages = PagesEditor::create_with_text("Character spacing").unwrap();
    let pages_box = pages
        .add_text_box(
            7,
            "Wide Pages text",
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
            "Normal Pages text",
            DrawablePoint { x: 280.0, y: 40.0 },
            DrawableSize {
                width: 240.0,
                height: 100.0,
            },
        )
        .unwrap();
    pages
        .set_text_box_text_baseline_shift(pages_box.drawable_object_id, pages_shift)
        .unwrap();
    pages
        .set_text_box_text_character_spacing(pages_box.drawable_object_id, pages_spacing)
        .unwrap();
    assert_eq!(
        pages
            .text_box_text_character_spacing(pages_sibling.drawable_object_id)
            .unwrap(),
        TextCharacterSpacing::NORMAL
    );
    let mut pages = PagesEditor::from_bytes(&pages.to_bytes().unwrap()).unwrap();
    assert_eq!(
        pages
            .text_box_text_character_spacing(pages_box.drawable_object_id)
            .unwrap(),
        pages_spacing
    );
    assert!(
        pages
            .reset_text_box_text_character_spacing(pages_box.drawable_object_id)
            .unwrap()
    );
    assert_eq!(
        pages
            .text_box_text_character_spacing(pages_box.drawable_object_id)
            .unwrap(),
        TextCharacterSpacing::NORMAL
    );
    assert_eq!(
        pages
            .text_box_text_baseline_shift(pages_box.drawable_object_id)
            .unwrap(),
        pages_shift
    );

    let numbers_spacing = TextCharacterSpacing::from_percent(-8.0).unwrap();
    let mut numbers = NumbersDocumentBuilder::new().build().unwrap();
    let sheet_id = numbers.sheets().unwrap()[0].object_id;
    let numbers_box = numbers
        .add_sheet_text_box(
            sheet_id,
            "Tight Numbers text",
            DrawablePoint { x: 20.0, y: 200.0 },
            DrawableSize {
                width: 240.0,
                height: 100.0,
            },
        )
        .unwrap();
    numbers
        .set_sheet_text_box_text_script(
            sheet_id,
            numbers_box.drawable_object_id,
            TextScript::Subscript,
        )
        .unwrap();
    numbers
        .set_sheet_text_box_text_character_spacing(
            sheet_id,
            numbers_box.drawable_object_id,
            numbers_spacing,
        )
        .unwrap();
    let mut numbers =
        crate::numbers::NumbersEditor::from_bytes(&numbers.to_bytes().unwrap()).unwrap();
    assert_eq!(
        numbers
            .sheet_text_box_text_character_spacing(sheet_id, numbers_box.drawable_object_id)
            .unwrap(),
        numbers_spacing
    );
    assert!(
        numbers
            .reset_sheet_text_box_text_character_spacing(sheet_id, numbers_box.drawable_object_id)
            .unwrap()
    );
    assert_eq!(
        numbers
            .sheet_text_box_text_script(sheet_id, numbers_box.drawable_object_id)
            .unwrap(),
        TextScript::Subscript
    );

    let keynote_spacing = TextCharacterSpacing::from_percent(6.0).unwrap();
    let mut keynote = KeynoteDocumentBuilder::new().build().unwrap();
    let keynote_box = keynote
        .add_slide_text_box(
            0,
            "Wide Keynote text",
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
    keynote
        .set_slide_text_box_text_character_spacing(
            0,
            keynote_box.drawable_object_id,
            keynote_spacing,
        )
        .unwrap();
    let mut keynote =
        crate::keynote::KeynoteEditor::from_bytes(&keynote.to_bytes().unwrap()).unwrap();
    assert_eq!(
        keynote
            .slide_text_box_text_character_spacing(0, keynote_box.drawable_object_id)
            .unwrap(),
        keynote_spacing
    );
    assert!(
        keynote
            .reset_slide_text_box_text_character_spacing(0, keynote_box.drawable_object_id)
            .unwrap()
    );
    assert_eq!(
        keynote
            .slide_text_box_text_capitalization(0, keynote_box.drawable_object_id)
            .unwrap(),
        TextCapitalization::TitleCase
    );
}

#[test]
fn uniform_ligatures_round_trip_isolate_and_reset_in_every_suite() {
    let pages_spacing = TextCharacterSpacing::from_percent(12.0).unwrap();
    let mut pages = PagesEditor::create_with_text("Ligatures").unwrap();
    let pages_box = pages
        .add_text_box(
            7,
            "Pages ligatures",
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
            "Standard Pages ligatures",
            DrawablePoint { x: 280.0, y: 40.0 },
            DrawableSize {
                width: 240.0,
                height: 100.0,
            },
        )
        .unwrap();
    pages
        .set_text_box_text_character_spacing(pages_box.drawable_object_id, pages_spacing)
        .unwrap();
    pages
        .set_text_box_text_ligatures(pages_box.drawable_object_id, TextLigatures::RequiredOnly)
        .unwrap();
    assert_eq!(
        pages
            .text_box_text_ligatures(pages_sibling.drawable_object_id)
            .unwrap(),
        TextLigatures::Standard
    );
    let mut pages = PagesEditor::from_bytes(&pages.to_bytes().unwrap()).unwrap();
    assert_eq!(
        pages
            .text_box_text_ligatures(pages_box.drawable_object_id)
            .unwrap(),
        TextLigatures::RequiredOnly
    );
    assert!(
        pages
            .reset_text_box_text_ligatures(pages_box.drawable_object_id)
            .unwrap()
    );
    assert_eq!(
        pages
            .text_box_text_ligatures(pages_box.drawable_object_id)
            .unwrap(),
        TextLigatures::Standard
    );
    assert_eq!(
        pages
            .text_box_text_character_spacing(pages_box.drawable_object_id)
            .unwrap(),
        pages_spacing
    );

    let numbers_shift = TextBaselineShift::from_points(-3.0).unwrap();
    let mut numbers = NumbersDocumentBuilder::new().build().unwrap();
    let sheet_id = numbers.sheets().unwrap()[0].object_id;
    let numbers_box = numbers
        .add_sheet_text_box(
            sheet_id,
            "Numbers ligatures",
            DrawablePoint { x: 20.0, y: 200.0 },
            DrawableSize {
                width: 240.0,
                height: 100.0,
            },
        )
        .unwrap();
    numbers
        .set_sheet_text_box_text_baseline_shift(
            sheet_id,
            numbers_box.drawable_object_id,
            numbers_shift,
        )
        .unwrap();
    numbers
        .set_sheet_text_box_text_ligatures(
            sheet_id,
            numbers_box.drawable_object_id,
            TextLigatures::All,
        )
        .unwrap();
    let mut numbers =
        crate::numbers::NumbersEditor::from_bytes(&numbers.to_bytes().unwrap()).unwrap();
    assert_eq!(
        numbers
            .sheet_text_box_text_ligatures(sheet_id, numbers_box.drawable_object_id)
            .unwrap(),
        TextLigatures::All
    );
    assert!(
        numbers
            .reset_sheet_text_box_text_ligatures(sheet_id, numbers_box.drawable_object_id)
            .unwrap()
    );
    assert_eq!(
        numbers
            .sheet_text_box_text_baseline_shift(sheet_id, numbers_box.drawable_object_id)
            .unwrap(),
        numbers_shift
    );

    let mut keynote = KeynoteDocumentBuilder::new().build().unwrap();
    let keynote_box = keynote
        .add_slide_text_box(
            0,
            "Keynote ligatures",
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
    keynote
        .set_slide_text_box_text_ligatures(
            0,
            keynote_box.drawable_object_id,
            TextLigatures::RequiredOnly,
        )
        .unwrap();
    let mut keynote =
        crate::keynote::KeynoteEditor::from_bytes(&keynote.to_bytes().unwrap()).unwrap();
    assert_eq!(
        keynote
            .slide_text_box_text_ligatures(0, keynote_box.drawable_object_id)
            .unwrap(),
        TextLigatures::RequiredOnly
    );
    assert!(
        keynote
            .reset_slide_text_box_text_ligatures(0, keynote_box.drawable_object_id)
            .unwrap()
    );
    assert_eq!(
        keynote
            .slide_text_box_text_capitalization(0, keynote_box.drawable_object_id)
            .unwrap(),
        TextCapitalization::TitleCase
    );
}

#[test]
fn uniform_text_outline_round_trips_isolates_and_resets_in_every_suite() {
    let pages_spacing = TextCharacterSpacing::from_percent(12.0).unwrap();
    let mut pages = PagesEditor::create_with_text("Outlines").unwrap();
    let pages_box = pages
        .add_text_box(
            7,
            "Pages outline",
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
    pages
        .set_text_box_text_character_spacing(pages_box.drawable_object_id, pages_spacing)
        .unwrap();
    pages
        .set_text_box_text_outline(pages_box.drawable_object_id, TextOutline::standard())
        .unwrap();
    assert_eq!(
        pages
            .text_box_text_outline(pages_sibling.drawable_object_id)
            .unwrap(),
        TextOutline::None
    );
    let mut pages = PagesEditor::from_bytes(&pages.to_bytes().unwrap()).unwrap();
    assert_eq!(
        pages
            .text_box_text_outline(pages_box.drawable_object_id)
            .unwrap(),
        TextOutline::standard()
    );
    assert!(
        pages
            .reset_text_box_text_outline(pages_box.drawable_object_id)
            .unwrap()
    );
    assert_eq!(
        pages
            .text_box_text_character_spacing(pages_box.drawable_object_id)
            .unwrap(),
        pages_spacing
    );

    let numbers_shift = TextBaselineShift::from_points(-3.0).unwrap();
    let mut numbers = NumbersDocumentBuilder::new().build().unwrap();
    let sheet_id = numbers.sheets().unwrap()[0].object_id;
    let numbers_box = numbers
        .add_sheet_text_box(
            sheet_id,
            "Numbers outline",
            DrawablePoint { x: 20.0, y: 200.0 },
            DrawableSize {
                width: 240.0,
                height: 100.0,
            },
        )
        .unwrap();
    numbers
        .set_sheet_text_box_text_baseline_shift(
            sheet_id,
            numbers_box.drawable_object_id,
            numbers_shift,
        )
        .unwrap();
    numbers
        .set_sheet_text_box_text_outline(
            sheet_id,
            numbers_box.drawable_object_id,
            TextOutline::standard(),
        )
        .unwrap();
    let mut numbers =
        crate::numbers::NumbersEditor::from_bytes(&numbers.to_bytes().unwrap()).unwrap();
    assert_eq!(
        numbers
            .sheet_text_box_text_outline(sheet_id, numbers_box.drawable_object_id)
            .unwrap(),
        TextOutline::standard()
    );
    assert!(
        numbers
            .reset_sheet_text_box_text_outline(sheet_id, numbers_box.drawable_object_id)
            .unwrap()
    );
    assert_eq!(
        numbers
            .sheet_text_box_text_baseline_shift(sheet_id, numbers_box.drawable_object_id)
            .unwrap(),
        numbers_shift
    );

    let mut keynote = KeynoteDocumentBuilder::new().build().unwrap();
    let keynote_box = keynote
        .add_slide_text_box(
            0,
            "Keynote outline",
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
    keynote
        .set_slide_text_box_text_outline(0, keynote_box.drawable_object_id, TextOutline::standard())
        .unwrap();
    let mut keynote =
        crate::keynote::KeynoteEditor::from_bytes(&keynote.to_bytes().unwrap()).unwrap();
    assert_eq!(
        keynote
            .slide_text_box_text_outline(0, keynote_box.drawable_object_id)
            .unwrap(),
        TextOutline::standard()
    );
    assert!(
        keynote
            .reset_slide_text_box_text_outline(0, keynote_box.drawable_object_id)
            .unwrap()
    );
    assert_eq!(
        keynote
            .slide_text_box_text_capitalization(0, keynote_box.drawable_object_id)
            .unwrap(),
        TextCapitalization::TitleCase
    );
}

#[test]
fn uniform_text_shadow_round_trips_isolates_and_resets_in_every_suite() {
    let pages_spacing = TextCharacterSpacing::from_percent(12.0).unwrap();
    let mut pages = PagesEditor::create_with_text("Shadows").unwrap();
    let pages_box = pages
        .add_text_box(
            7,
            "Pages shadow",
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
    pages
        .set_text_box_text_character_spacing(pages_box.drawable_object_id, pages_spacing)
        .unwrap();
    pages
        .set_text_box_text_shadow(pages_box.drawable_object_id, TextShadow::standard())
        .unwrap();
    assert_eq!(
        pages
            .text_box_text_shadow(pages_sibling.drawable_object_id)
            .unwrap(),
        TextShadow::None
    );
    let mut pages = PagesEditor::from_bytes(&pages.to_bytes().unwrap()).unwrap();
    assert_eq!(
        pages
            .text_box_text_shadow(pages_box.drawable_object_id)
            .unwrap(),
        TextShadow::standard()
    );
    assert!(
        pages
            .reset_text_box_text_shadow(pages_box.drawable_object_id)
            .unwrap()
    );
    assert_eq!(
        pages
            .text_box_text_character_spacing(pages_box.drawable_object_id)
            .unwrap(),
        pages_spacing
    );

    let numbers_shift = TextBaselineShift::from_points(-3.0).unwrap();
    let mut numbers = NumbersDocumentBuilder::new().build().unwrap();
    let sheet_id = numbers.sheets().unwrap()[0].object_id;
    let numbers_box = numbers
        .add_sheet_text_box(
            sheet_id,
            "Numbers shadow",
            DrawablePoint { x: 20.0, y: 200.0 },
            DrawableSize {
                width: 240.0,
                height: 100.0,
            },
        )
        .unwrap();
    numbers
        .set_sheet_text_box_text_baseline_shift(
            sheet_id,
            numbers_box.drawable_object_id,
            numbers_shift,
        )
        .unwrap();
    numbers
        .set_sheet_text_box_text_shadow(
            sheet_id,
            numbers_box.drawable_object_id,
            TextShadow::standard(),
        )
        .unwrap();
    let mut numbers =
        crate::numbers::NumbersEditor::from_bytes(&numbers.to_bytes().unwrap()).unwrap();
    assert_eq!(
        numbers
            .sheet_text_box_text_shadow(sheet_id, numbers_box.drawable_object_id)
            .unwrap(),
        TextShadow::standard()
    );
    assert!(
        numbers
            .reset_sheet_text_box_text_shadow(sheet_id, numbers_box.drawable_object_id)
            .unwrap()
    );
    assert_eq!(
        numbers
            .sheet_text_box_text_baseline_shift(sheet_id, numbers_box.drawable_object_id)
            .unwrap(),
        numbers_shift
    );

    let mut keynote = KeynoteDocumentBuilder::new().build().unwrap();
    let keynote_box = keynote
        .add_slide_text_box(
            0,
            "Keynote shadow",
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
    keynote
        .set_slide_text_box_text_shadow(0, keynote_box.drawable_object_id, TextShadow::standard())
        .unwrap();
    let mut keynote =
        crate::keynote::KeynoteEditor::from_bytes(&keynote.to_bytes().unwrap()).unwrap();
    assert_eq!(
        keynote
            .slide_text_box_text_shadow(0, keynote_box.drawable_object_id)
            .unwrap(),
        TextShadow::standard()
    );
    assert!(
        keynote
            .reset_slide_text_box_text_shadow(0, keynote_box.drawable_object_id)
            .unwrap()
    );
    assert_eq!(
        keynote
            .slide_text_box_text_capitalization(0, keynote_box.drawable_object_id)
            .unwrap(),
        TextCapitalization::TitleCase
    );
}

#[test]
fn uniform_text_background_round_trips_isolates_and_resets_in_every_suite() {
    let pages_background = TextBackground::Color(
        RgbaColor::new(0.95, 0.42, 0.17, 0.72, RgbColorSpace::DisplayP3).unwrap(),
    );
    let pages_spacing = TextCharacterSpacing::from_percent(12.0).unwrap();
    let mut pages = PagesEditor::create_with_text("Backgrounds").unwrap();
    let pages_box = pages
        .add_text_box(
            7,
            "Pages background",
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
    pages
        .set_text_box_text_character_spacing(pages_box.drawable_object_id, pages_spacing)
        .unwrap();
    pages
        .set_text_box_text_background(pages_box.drawable_object_id, pages_background)
        .unwrap();
    assert_eq!(
        pages
            .text_box_text_background(pages_sibling.drawable_object_id)
            .unwrap(),
        TextBackground::None
    );
    let mut pages = PagesEditor::from_bytes(&pages.to_bytes().unwrap()).unwrap();
    assert_eq!(
        pages
            .text_box_text_background(pages_box.drawable_object_id)
            .unwrap(),
        pages_background
    );
    assert!(
        pages
            .reset_text_box_text_background(pages_box.drawable_object_id)
            .unwrap()
    );
    assert_eq!(
        pages
            .text_box_text_character_spacing(pages_box.drawable_object_id)
            .unwrap(),
        pages_spacing
    );

    let numbers_background =
        TextBackground::Color(RgbaColor::new(0.22, 0.82, 0.38, 1.0, RgbColorSpace::Srgb).unwrap());
    let numbers_shift = TextBaselineShift::from_points(-3.0).unwrap();
    let mut numbers = NumbersDocumentBuilder::new().build().unwrap();
    let sheet_id = numbers.sheets().unwrap()[0].object_id;
    let numbers_box = numbers
        .add_sheet_text_box(
            sheet_id,
            "Numbers background",
            DrawablePoint { x: 20.0, y: 200.0 },
            DrawableSize {
                width: 240.0,
                height: 100.0,
            },
        )
        .unwrap();
    numbers
        .set_sheet_text_box_text_baseline_shift(
            sheet_id,
            numbers_box.drawable_object_id,
            numbers_shift,
        )
        .unwrap();
    numbers
        .set_sheet_text_box_text_background(
            sheet_id,
            numbers_box.drawable_object_id,
            numbers_background,
        )
        .unwrap();
    let mut numbers =
        crate::numbers::NumbersEditor::from_bytes(&numbers.to_bytes().unwrap()).unwrap();
    assert_eq!(
        numbers
            .sheet_text_box_text_background(sheet_id, numbers_box.drawable_object_id)
            .unwrap(),
        numbers_background
    );
    assert!(
        numbers
            .reset_sheet_text_box_text_background(sheet_id, numbers_box.drawable_object_id)
            .unwrap()
    );
    assert_eq!(
        numbers
            .sheet_text_box_text_baseline_shift(sheet_id, numbers_box.drawable_object_id)
            .unwrap(),
        numbers_shift
    );

    let keynote_background = TextBackground::Color(
        RgbaColor::new(0.18, 0.44, 0.92, 0.84, RgbColorSpace::DisplayP3).unwrap(),
    );
    let mut keynote = KeynoteDocumentBuilder::new().build().unwrap();
    let keynote_box = keynote
        .add_slide_text_box(
            0,
            "Keynote background",
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
    keynote
        .set_slide_text_box_text_background(0, keynote_box.drawable_object_id, keynote_background)
        .unwrap();
    let mut keynote =
        crate::keynote::KeynoteEditor::from_bytes(&keynote.to_bytes().unwrap()).unwrap();
    assert_eq!(
        keynote
            .slide_text_box_text_background(0, keynote_box.drawable_object_id)
            .unwrap(),
        keynote_background
    );
    assert!(
        keynote
            .reset_slide_text_box_text_background(0, keynote_box.drawable_object_id)
            .unwrap()
    );
    assert_eq!(
        keynote
            .slide_text_box_text_capitalization(0, keynote_box.drawable_object_id)
            .unwrap(),
        TextCapitalization::TitleCase
    );
}

#[test]
fn paragraph_background_round_trips_isolates_and_resets_in_every_suite() {
    let background = ParagraphBackground::Color(
        RgbaColor::new(1.0, 0.588_738_74, 0.552_926_2, 1.0, RgbColorSpace::Srgb).unwrap(),
    );

    let mut pages = PagesEditor::create_with_text("Paragraph backgrounds").unwrap();
    let pages_box = pages
        .add_text_box(
            7,
            "Pages paragraph background",
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
            "Plain Pages paragraph",
            DrawablePoint { x: 280.0, y: 40.0 },
            DrawableSize {
                width: 240.0,
                height: 100.0,
            },
        )
        .unwrap();
    let pages_before = pages.to_bytes().unwrap();
    pages
        .set_text_box_paragraph_background(pages_box.drawable_object_id, background)
        .unwrap();
    assert_eq!(
        pages
            .text_box_paragraph_background(pages_sibling.drawable_object_id)
            .unwrap(),
        ParagraphBackground::None
    );
    let mut pages = PagesEditor::from_bytes(&pages.to_bytes().unwrap()).unwrap();
    assert_eq!(
        pages
            .text_box_paragraph_background(pages_box.drawable_object_id)
            .unwrap(),
        background
    );
    assert!(
        pages
            .reset_text_box_paragraph_background(pages_box.drawable_object_id)
            .unwrap()
    );
    assert_eq!(pages.to_bytes().unwrap(), pages_before);

    let mut numbers = NumbersDocumentBuilder::new().build().unwrap();
    let sheet_id = numbers.sheets().unwrap()[0].object_id;
    let numbers_box = numbers
        .add_sheet_text_box(
            sheet_id,
            "Numbers paragraph background",
            DrawablePoint { x: 20.0, y: 200.0 },
            DrawableSize {
                width: 240.0,
                height: 100.0,
            },
        )
        .unwrap();
    numbers
        .set_sheet_text_box_paragraph_background(
            sheet_id,
            numbers_box.drawable_object_id,
            background,
        )
        .unwrap();
    let numbers = crate::numbers::NumbersEditor::from_bytes(&numbers.to_bytes().unwrap()).unwrap();
    assert_eq!(
        numbers
            .sheet_text_box_paragraph_background(sheet_id, numbers_box.drawable_object_id)
            .unwrap(),
        background
    );

    let mut keynote = KeynoteDocumentBuilder::new().build().unwrap();
    let keynote_box = keynote
        .add_slide_text_box(
            0,
            "Keynote paragraph background",
            DrawablePoint { x: 80.0, y: 500.0 },
            DrawableSize {
                width: 500.0,
                height: 100.0,
            },
        )
        .unwrap();
    keynote
        .set_slide_text_box_paragraph_background(0, keynote_box.drawable_object_id, background)
        .unwrap();
    let keynote = crate::keynote::KeynoteEditor::from_bytes(&keynote.to_bytes().unwrap()).unwrap();
    assert_eq!(
        keynote
            .slide_text_box_paragraph_background(0, keynote_box.drawable_object_id)
            .unwrap(),
        background
    );
}

#[test]
fn paragraph_borders_round_trip_isolate_and_reset_in_every_suite() {
    let borders = ParagraphBorders::Bordered(
        ParagraphBorder::new(
            RgbaColor::black(),
            StrokeWidth::new(3.0).unwrap(),
            StrokePattern::Solid,
            BorderSides::ALL,
            BorderOffset::from_points(9.0).unwrap(),
            true,
        )
        .unwrap(),
    );

    let mut pages = PagesEditor::create_with_text("Paragraph borders").unwrap();
    let pages_box = pages
        .add_text_box(
            7,
            "Pages paragraph border",
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
            "Plain Pages paragraph",
            DrawablePoint { x: 280.0, y: 40.0 },
            DrawableSize {
                width: 240.0,
                height: 100.0,
            },
        )
        .unwrap();
    let pages_before = pages.to_bytes().unwrap();
    pages
        .set_text_box_paragraph_borders(pages_box.drawable_object_id, borders)
        .unwrap();
    assert_eq!(
        pages
            .text_box_paragraph_borders(pages_sibling.drawable_object_id)
            .unwrap(),
        ParagraphBorders::None
    );
    let mut pages = PagesEditor::from_bytes(&pages.to_bytes().unwrap()).unwrap();
    assert_eq!(
        pages
            .text_box_paragraph_borders(pages_box.drawable_object_id)
            .unwrap(),
        borders
    );
    assert!(
        pages
            .reset_text_box_paragraph_borders(pages_box.drawable_object_id)
            .unwrap()
    );
    assert_eq!(pages.to_bytes().unwrap(), pages_before);

    let mut numbers = NumbersDocumentBuilder::new().build().unwrap();
    let sheet_id = numbers.sheets().unwrap()[0].object_id;
    let numbers_box = numbers
        .add_sheet_text_box(
            sheet_id,
            "Numbers paragraph border",
            DrawablePoint { x: 20.0, y: 200.0 },
            DrawableSize {
                width: 240.0,
                height: 100.0,
            },
        )
        .unwrap();
    let numbers_before = numbers.to_bytes().unwrap();
    numbers
        .set_sheet_text_box_paragraph_borders(sheet_id, numbers_box.drawable_object_id, borders)
        .unwrap();
    let mut numbers =
        crate::numbers::NumbersEditor::from_bytes(&numbers.to_bytes().unwrap()).unwrap();
    assert_eq!(
        numbers
            .sheet_text_box_paragraph_borders(sheet_id, numbers_box.drawable_object_id)
            .unwrap(),
        borders
    );
    assert!(
        numbers
            .reset_sheet_text_box_paragraph_borders(sheet_id, numbers_box.drawable_object_id)
            .unwrap()
    );
    assert_eq!(numbers.to_bytes().unwrap(), numbers_before);

    let mut keynote = KeynoteDocumentBuilder::new().build().unwrap();
    let keynote_box = keynote
        .add_slide_text_box(
            0,
            "Keynote paragraph border",
            DrawablePoint { x: 80.0, y: 500.0 },
            DrawableSize {
                width: 500.0,
                height: 100.0,
            },
        )
        .unwrap();
    let keynote_before = keynote.to_bytes().unwrap();
    keynote
        .set_slide_text_box_paragraph_borders(0, keynote_box.drawable_object_id, borders)
        .unwrap();
    let mut keynote =
        crate::keynote::KeynoteEditor::from_bytes(&keynote.to_bytes().unwrap()).unwrap();
    assert_eq!(
        keynote
            .slide_text_box_paragraph_borders(0, keynote_box.drawable_object_id)
            .unwrap(),
        borders
    );
    assert!(
        keynote
            .reset_slide_text_box_paragraph_borders(0, keynote_box.drawable_object_id)
            .unwrap()
    );
    assert_eq!(keynote.to_bytes().unwrap(), keynote_before);
}

#[test]
fn paragraph_flow_round_trips_isolates_and_resets_in_every_suite() {
    let flow = ParagraphFlow::new()
        .with_keep_lines_together(true)
        .with_keep_with_next(true)
        .with_start_on_new_page(true)
        .with_prevent_widow_orphan_lines(false)
        .with_hyphenation(ParagraphHyphenation::Prevented);

    let mut pages = PagesEditor::create_with_text("Paragraph flow").unwrap();
    let pages_box = pages
        .add_text_box(
            7,
            "Pages paragraph flow",
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
            "Plain Pages paragraph",
            DrawablePoint { x: 280.0, y: 40.0 },
            DrawableSize {
                width: 240.0,
                height: 100.0,
            },
        )
        .unwrap();
    let pages_before = pages.to_bytes().unwrap();
    pages
        .set_text_box_paragraph_flow(pages_box.drawable_object_id, flow)
        .unwrap();
    assert_eq!(
        pages
            .text_box_paragraph_flow(pages_sibling.drawable_object_id)
            .unwrap(),
        ParagraphFlow::default()
    );
    let mut pages = PagesEditor::from_bytes(&pages.to_bytes().unwrap()).unwrap();
    assert_eq!(
        pages
            .text_box_paragraph_flow(pages_box.drawable_object_id)
            .unwrap(),
        flow
    );
    assert!(
        pages
            .reset_text_box_paragraph_flow(pages_box.drawable_object_id)
            .unwrap()
    );
    assert_eq!(pages.to_bytes().unwrap(), pages_before);

    let mut numbers = NumbersDocumentBuilder::new().build().unwrap();
    let sheet_id = numbers.sheets().unwrap()[0].object_id;
    let numbers_box = numbers
        .add_sheet_text_box(
            sheet_id,
            "Numbers paragraph flow",
            DrawablePoint { x: 20.0, y: 200.0 },
            DrawableSize {
                width: 240.0,
                height: 100.0,
            },
        )
        .unwrap();
    let numbers_before = numbers.to_bytes().unwrap();
    numbers
        .set_sheet_text_box_paragraph_flow(sheet_id, numbers_box.drawable_object_id, flow)
        .unwrap();
    let mut numbers =
        crate::numbers::NumbersEditor::from_bytes(&numbers.to_bytes().unwrap()).unwrap();
    assert_eq!(
        numbers
            .sheet_text_box_paragraph_flow(sheet_id, numbers_box.drawable_object_id)
            .unwrap(),
        flow
    );
    assert!(
        numbers
            .reset_sheet_text_box_paragraph_flow(sheet_id, numbers_box.drawable_object_id)
            .unwrap()
    );
    assert_eq!(numbers.to_bytes().unwrap(), numbers_before);

    let mut keynote = KeynoteDocumentBuilder::new().build().unwrap();
    let keynote_box = keynote
        .add_slide_text_box(
            0,
            "Keynote paragraph flow",
            DrawablePoint { x: 80.0, y: 500.0 },
            DrawableSize {
                width: 500.0,
                height: 100.0,
            },
        )
        .unwrap();
    let keynote_before = keynote.to_bytes().unwrap();
    keynote
        .set_slide_text_box_paragraph_flow(0, keynote_box.drawable_object_id, flow)
        .unwrap();
    let mut keynote =
        crate::keynote::KeynoteEditor::from_bytes(&keynote.to_bytes().unwrap()).unwrap();
    assert_eq!(
        keynote
            .slide_text_box_paragraph_flow(0, keynote_box.drawable_object_id)
            .unwrap(),
        flow
    );
    assert!(
        keynote
            .reset_slide_text_box_paragraph_flow(0, keynote_box.drawable_object_id)
            .unwrap()
    );
    assert_eq!(keynote.to_bytes().unwrap(), keynote_before);
}

#[test]
fn paragraph_writing_direction_round_trips_and_resets_in_every_suite() {
    let mut pages = PagesEditor::create_with_text("Directions").unwrap();
    let pages_box = pages
        .add_text_box(
            7,
            "שלום Pages",
            DrawablePoint { x: 20.0, y: 40.0 },
            DrawableSize {
                width: 240.0,
                height: 100.0,
            },
        )
        .unwrap();
    let pages_before = pages.to_bytes().unwrap();
    pages
        .set_text_box_paragraph_writing_direction(
            pages_box.drawable_object_id,
            ParagraphWritingDirection::RightToLeft,
        )
        .unwrap();
    let mut pages = PagesEditor::from_bytes(&pages.to_bytes().unwrap()).unwrap();
    assert_eq!(
        pages
            .text_box_paragraph_writing_direction(pages_box.drawable_object_id)
            .unwrap(),
        ParagraphWritingDirection::RightToLeft
    );
    assert!(
        pages
            .reset_text_box_paragraph_writing_direction(pages_box.drawable_object_id)
            .unwrap()
    );
    assert_eq!(pages.to_bytes().unwrap(), pages_before);

    let mut numbers = NumbersDocumentBuilder::new().build().unwrap();
    let sheet_id = numbers.sheets().unwrap()[0].object_id;
    let numbers_box = numbers
        .add_sheet_text_box(
            sheet_id,
            "مرحبا Numbers",
            DrawablePoint { x: 20.0, y: 200.0 },
            DrawableSize {
                width: 240.0,
                height: 100.0,
            },
        )
        .unwrap();
    let numbers_before = numbers.to_bytes().unwrap();
    numbers
        .set_sheet_text_box_paragraph_writing_direction(
            sheet_id,
            numbers_box.drawable_object_id,
            ParagraphWritingDirection::RightToLeft,
        )
        .unwrap();
    let mut numbers =
        crate::numbers::NumbersEditor::from_bytes(&numbers.to_bytes().unwrap()).unwrap();
    assert_eq!(
        numbers
            .sheet_text_box_paragraph_writing_direction(sheet_id, numbers_box.drawable_object_id,)
            .unwrap(),
        ParagraphWritingDirection::RightToLeft
    );
    assert!(
        numbers
            .reset_sheet_text_box_paragraph_writing_direction(
                sheet_id,
                numbers_box.drawable_object_id,
            )
            .unwrap()
    );
    assert_eq!(numbers.to_bytes().unwrap(), numbers_before);

    let mut keynote = KeynoteDocumentBuilder::new().build().unwrap();
    let keynote_box = keynote
        .add_slide_text_box(
            0,
            "שלום Keynote",
            DrawablePoint { x: 80.0, y: 500.0 },
            DrawableSize {
                width: 500.0,
                height: 100.0,
            },
        )
        .unwrap();
    let keynote_before = keynote.to_bytes().unwrap();
    keynote
        .set_slide_text_box_paragraph_writing_direction(
            0,
            keynote_box.drawable_object_id,
            ParagraphWritingDirection::LeftToRight,
        )
        .unwrap();
    let mut keynote =
        crate::keynote::KeynoteEditor::from_bytes(&keynote.to_bytes().unwrap()).unwrap();
    assert_eq!(
        keynote
            .slide_text_box_paragraph_writing_direction(0, keynote_box.drawable_object_id)
            .unwrap(),
        ParagraphWritingDirection::LeftToRight
    );
    assert!(
        keynote
            .reset_slide_text_box_paragraph_writing_direction(0, keynote_box.drawable_object_id,)
            .unwrap()
    );
    assert_eq!(keynote.to_bytes().unwrap(), keynote_before);
}

#[test]
fn paragraph_tab_defaults_round_trip_compose_and_reset_in_every_suite() {
    let interval = ParagraphDefaultTabInterval::from_points(54.0).unwrap();
    let character = ParagraphDecimalTabCharacter::COMMA;

    let mut pages = PagesEditor::create_with_text("Tab defaults").unwrap();
    let pages_box = pages
        .add_text_box(
            7,
            "12,34\tPages",
            DrawablePoint { x: 20.0, y: 40.0 },
            DrawableSize {
                width: 240.0,
                height: 100.0,
            },
        )
        .unwrap();
    let pages_before = pages.to_bytes().unwrap();
    pages
        .set_text_box_paragraph_decimal_tab_character(pages_box.drawable_object_id, character)
        .unwrap();
    pages
        .set_text_box_paragraph_default_tab_interval(pages_box.drawable_object_id, interval)
        .unwrap();
    let mut pages = PagesEditor::from_bytes(&pages.to_bytes().unwrap()).unwrap();
    assert_eq!(
        pages
            .text_box_paragraph_decimal_tab_character(pages_box.drawable_object_id)
            .unwrap(),
        character
    );
    assert_eq!(
        pages
            .text_box_paragraph_default_tab_interval(pages_box.drawable_object_id)
            .unwrap(),
        interval
    );
    assert!(
        pages
            .reset_text_box_paragraph_decimal_tab_character(pages_box.drawable_object_id)
            .unwrap()
    );
    assert_eq!(
        pages
            .text_box_paragraph_default_tab_interval(pages_box.drawable_object_id)
            .unwrap(),
        interval
    );
    assert!(
        pages
            .reset_text_box_paragraph_default_tab_interval(pages_box.drawable_object_id)
            .unwrap()
    );
    assert_eq!(pages.to_bytes().unwrap(), pages_before);

    let mut numbers = NumbersDocumentBuilder::new().build().unwrap();
    let sheet_id = numbers.sheets().unwrap()[0].object_id;
    let numbers_box = numbers
        .add_sheet_text_box(
            sheet_id,
            "12,34\tNumbers",
            DrawablePoint { x: 20.0, y: 200.0 },
            DrawableSize {
                width: 240.0,
                height: 100.0,
            },
        )
        .unwrap();
    let numbers_before = numbers.to_bytes().unwrap();
    numbers
        .set_sheet_text_box_paragraph_decimal_tab_character(
            sheet_id,
            numbers_box.drawable_object_id,
            character,
        )
        .unwrap();
    numbers
        .set_sheet_text_box_paragraph_default_tab_interval(
            sheet_id,
            numbers_box.drawable_object_id,
            interval,
        )
        .unwrap();
    let mut numbers =
        crate::numbers::NumbersEditor::from_bytes(&numbers.to_bytes().unwrap()).unwrap();
    assert_eq!(
        numbers
            .sheet_text_box_paragraph_decimal_tab_character(
                sheet_id,
                numbers_box.drawable_object_id,
            )
            .unwrap(),
        character
    );
    assert_eq!(
        numbers
            .sheet_text_box_paragraph_default_tab_interval(
                sheet_id,
                numbers_box.drawable_object_id,
            )
            .unwrap(),
        interval
    );
    assert!(
        numbers
            .reset_sheet_text_box_paragraph_decimal_tab_character(
                sheet_id,
                numbers_box.drawable_object_id,
            )
            .unwrap()
    );
    assert_eq!(
        numbers
            .sheet_text_box_paragraph_default_tab_interval(
                sheet_id,
                numbers_box.drawable_object_id,
            )
            .unwrap(),
        interval
    );
    assert!(
        numbers
            .reset_sheet_text_box_paragraph_default_tab_interval(
                sheet_id,
                numbers_box.drawable_object_id,
            )
            .unwrap()
    );
    assert_eq!(numbers.to_bytes().unwrap(), numbers_before);

    let mut keynote = KeynoteDocumentBuilder::new().build().unwrap();
    let keynote_box = keynote
        .add_slide_text_box(
            0,
            "12,34\tKeynote",
            DrawablePoint { x: 80.0, y: 500.0 },
            DrawableSize {
                width: 500.0,
                height: 100.0,
            },
        )
        .unwrap();
    let keynote_before = keynote.to_bytes().unwrap();
    keynote
        .set_slide_text_box_paragraph_decimal_tab_character(
            0,
            keynote_box.drawable_object_id,
            character,
        )
        .unwrap();
    keynote
        .set_slide_text_box_paragraph_default_tab_interval(
            0,
            keynote_box.drawable_object_id,
            interval,
        )
        .unwrap();
    let mut keynote =
        crate::keynote::KeynoteEditor::from_bytes(&keynote.to_bytes().unwrap()).unwrap();
    assert_eq!(
        keynote
            .slide_text_box_paragraph_decimal_tab_character(0, keynote_box.drawable_object_id)
            .unwrap(),
        character
    );
    assert_eq!(
        keynote
            .slide_text_box_paragraph_default_tab_interval(0, keynote_box.drawable_object_id)
            .unwrap(),
        interval
    );
    assert!(
        keynote
            .reset_slide_text_box_paragraph_decimal_tab_character(
                0,
                keynote_box.drawable_object_id,
            )
            .unwrap()
    );
    assert_eq!(
        keynote
            .slide_text_box_paragraph_default_tab_interval(0, keynote_box.drawable_object_id)
            .unwrap(),
        interval
    );
    assert!(
        keynote
            .reset_slide_text_box_paragraph_default_tab_interval(0, keynote_box.drawable_object_id,)
            .unwrap()
    );
    assert_eq!(keynote.to_bytes().unwrap(), keynote_before);
}

#[test]
fn uniform_text_font_round_trips_isolates_and_resets_in_every_suite() {
    let pages_font = TextFont::named("Georgia-Bold").unwrap();
    let pages_background =
        TextBackground::Color(RgbaColor::new(0.95, 0.72, 0.52, 1.0, RgbColorSpace::Srgb).unwrap());
    let mut pages = PagesEditor::create_with_text("Fonts").unwrap();
    let pages_box = pages
        .add_text_box(
            5,
            "Pages font",
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
            "Plain Pages font",
            DrawablePoint { x: 280.0, y: 40.0 },
            DrawableSize {
                width: 240.0,
                height: 100.0,
            },
        )
        .unwrap();
    let inherited_pages = pages
        .text_box_text_font(pages_box.drawable_object_id)
        .unwrap();
    pages
        .set_text_box_text_background(pages_box.drawable_object_id, pages_background)
        .unwrap();
    pages
        .set_text_box_text_font(pages_box.drawable_object_id, pages_font.clone())
        .unwrap();
    assert_eq!(
        pages
            .text_box_text_font(pages_sibling.drawable_object_id)
            .unwrap(),
        inherited_pages
    );
    let mut pages = PagesEditor::from_bytes(&pages.to_bytes().unwrap()).unwrap();
    assert_eq!(
        pages
            .text_box_text_font(pages_box.drawable_object_id)
            .unwrap(),
        pages_font
    );
    assert!(
        pages
            .reset_text_box_text_font(pages_box.drawable_object_id)
            .unwrap()
    );
    assert_eq!(
        pages
            .text_box_text_font(pages_box.drawable_object_id)
            .unwrap(),
        inherited_pages
    );
    assert_eq!(
        pages
            .text_box_text_background(pages_box.drawable_object_id)
            .unwrap(),
        pages_background
    );

    let numbers_font = TextFont::named("TimesNewRomanPS-ItalicMT").unwrap();
    let mut numbers = NumbersDocumentBuilder::new().build().unwrap();
    let sheet_id = numbers.sheets().unwrap()[0].object_id;
    let numbers_box = numbers
        .add_sheet_text_box(
            sheet_id,
            "Numbers font",
            DrawablePoint { x: 20.0, y: 200.0 },
            DrawableSize {
                width: 240.0,
                height: 100.0,
            },
        )
        .unwrap();
    let inherited_numbers = numbers
        .sheet_text_box_text_font(sheet_id, numbers_box.drawable_object_id)
        .unwrap();
    numbers
        .set_sheet_text_box_text_font(
            sheet_id,
            numbers_box.drawable_object_id,
            numbers_font.clone(),
        )
        .unwrap();
    let mut numbers =
        crate::numbers::NumbersEditor::from_bytes(&numbers.to_bytes().unwrap()).unwrap();
    assert_eq!(
        numbers
            .sheet_text_box_text_font(sheet_id, numbers_box.drawable_object_id)
            .unwrap(),
        numbers_font
    );
    assert!(
        numbers
            .reset_sheet_text_box_text_font(sheet_id, numbers_box.drawable_object_id)
            .unwrap()
    );
    assert_eq!(
        numbers
            .sheet_text_box_text_font(sheet_id, numbers_box.drawable_object_id)
            .unwrap(),
        inherited_numbers
    );

    let keynote_font = TextFont::named("AvenirNext-BoldItalic").unwrap();
    let mut keynote = KeynoteDocumentBuilder::new().build().unwrap();
    let keynote_box = keynote
        .add_slide_text_box(
            0,
            "Keynote font",
            DrawablePoint { x: 80.0, y: 500.0 },
            DrawableSize {
                width: 500.0,
                height: 100.0,
            },
        )
        .unwrap();
    let inherited_keynote = keynote
        .slide_text_box_text_font(0, keynote_box.drawable_object_id)
        .unwrap();
    keynote
        .set_slide_text_box_text_font(0, keynote_box.drawable_object_id, keynote_font.clone())
        .unwrap();
    let mut keynote =
        crate::keynote::KeynoteEditor::from_bytes(&keynote.to_bytes().unwrap()).unwrap();
    assert_eq!(
        keynote
            .slide_text_box_text_font(0, keynote_box.drawable_object_id)
            .unwrap(),
        keynote_font
    );
    assert!(
        keynote
            .reset_slide_text_box_text_font(0, keynote_box.drawable_object_id)
            .unwrap()
    );
    assert_eq!(
        keynote
            .slide_text_box_text_font(0, keynote_box.drawable_object_id)
            .unwrap(),
        inherited_keynote
    );
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
    let columns = Columns::equal(Count::new(2).unwrap(), None);
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
    let location = storage::locate(&package, storage_id).unwrap();
    let archive_name = location.wire.archive_name.clone();
    package
        .update_archive(&archive_name, |archive| {
            let object = archive.object_mut(storage_id).unwrap();
            let index = location.wire.message_index;
            let message_type = location.wire.message_type;
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
            .set_text_font(storage_id, TextFont::named("Georgia-Bold").unwrap())
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
    assert!(
        editor
            .set_text_script(storage_id, TextScript::Superscript)
            .is_err()
    );
    assert!(
        editor
            .set_text_baseline_shift(storage_id, TextBaselineShift::from_points(4.0).unwrap(),)
            .is_err()
    );
    assert!(
        editor
            .set_text_character_spacing(
                storage_id,
                TextCharacterSpacing::from_percent(12.0).unwrap(),
            )
            .is_err()
    );
    assert!(
        editor
            .set_text_ligatures(storage_id, TextLigatures::All)
            .is_err()
    );
    assert!(
        editor
            .set_text_outline(storage_id, TextOutline::standard())
            .is_err()
    );
    assert!(
        editor
            .set_text_shadow(storage_id, TextShadow::standard())
            .is_err()
    );
    assert!(
        editor
            .set_text_background(
                storage_id,
                TextBackground::Color(
                    RgbaColor::new(0.2, 0.6, 0.9, 1.0, RgbColorSpace::Srgb).unwrap(),
                ),
            )
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn paragraph_style_mutations_use_the_resolved_storage_with_a_2022_style_sibling() {
    let mut pages = PagesEditor::create_with_text("Alignment").unwrap();
    let created = pages
        .add_text_box(
            9,
            "Paragraph style",
            DrawablePoint { x: 20.0, y: 40.0 },
            DrawableSize {
                width: 240.0,
                height: 100.0,
            },
        )
        .unwrap();
    let storage_id = created.storage.object_id;
    let mut package = pages.into_package();
    let location = storage::locate(&package, storage_id).unwrap();
    let style_data = tswp::ParagraphStyleArchive {
        super_: crate::protobuf::tss::StyleArchive::default(),
        ..Default::default()
    }
    .encode_to_vec();
    package
        .update_archive(&location.wire.archive_name, |archive| {
            archive
                .object_mut(storage_id)
                .unwrap()
                .push_message(RawMessage {
                    type_: 2_022,
                    data: style_data.clone(),
                })?;
            Ok(())
        })
        .unwrap();

    let mut editor = IWorkTextEditor::from_package(package);
    editor
        .set_paragraph_alignment(storage_id, TextAlignment::Center)
        .unwrap();
    editor
        .set_paragraph_alignment(storage_id, TextAlignment::Right)
        .unwrap();
    assert!(editor.reset_paragraph_alignment(storage_id).unwrap());

    let package = editor.into_package();
    let location = storage::locate(&package, storage_id).unwrap();
    let archive = package.archive(&location.wire.archive_name).unwrap();
    let object = archive.object(storage_id).unwrap();
    assert_eq!(
        object.messages[location.wire.message_index].type_,
        location.wire.message_type
    );
    let sibling = object
        .messages
        .iter()
        .find(|message| message.type_ == 2_022 && message.data == style_data)
        .unwrap();
    assert_eq!(sibling.data, style_data);
}

#[test]
fn paragraph_style_anchor_rejects_a_stale_message_type_transactionally() {
    let mut pages = PagesEditor::create_with_text("Alignment").unwrap();
    let created = pages
        .add_text_box(
            9,
            "Paragraph style",
            DrawablePoint { x: 20.0, y: 40.0 },
            DrawableSize {
                width: 240.0,
                height: 100.0,
            },
        )
        .unwrap();
    let storage_id = created.storage.object_id;
    let mut package = pages.into_package();
    let location = storage::locate(&package, storage_id).unwrap();
    let archive_name = location.wire.archive_name.clone();
    package
        .update_archive(&archive_name, |archive| {
            let object = archive.object_mut(storage_id).unwrap();
            let original = object.messages[location.wire.message_index].clone();
            object.replace_message(
                location.wire.message_index,
                RawMessage {
                    type_: 99_999,
                    data: original.data,
                },
            )?;
            Ok(())
        })
        .unwrap();
    let before = package.to_bytes().unwrap();
    assert!(
        storage::patch_style_reference(
            &mut package,
            &location,
            location.style_id,
            location.style_id + 1,
        )
        .is_err()
    );
    assert_eq!(package.to_bytes().unwrap(), before);
}

#[test]
fn paragraph_style_mutation_uses_exact_native_style_anchor_with_sibling_payload() {
    let mut pages = PagesEditor::create_with_text("Alignment").unwrap();
    let created = pages
        .add_text_box(
            9,
            "Paragraph style",
            DrawablePoint { x: 20.0, y: 40.0 },
            DrawableSize {
                width: 240.0,
                height: 100.0,
            },
        )
        .unwrap();
    pages
        .set_text_box_paragraph_alignment(created.drawable_object_id, TextAlignment::Center)
        .unwrap();
    let storage_id = created.storage.object_id;
    let mut package = pages.into_package();
    let storage = storage::locate(&package, storage_id).unwrap();
    let style = native::locate_style(&package, storage.style_id).unwrap();
    let sibling_data = vec![0x98, 0x06, 0x07];
    package
        .update_archive(&style.archive_name, |archive| {
            let object = archive.object_mut(style.object_id).unwrap();
            object.messages.insert(
                style.message_index,
                RawMessage {
                    type_: 2_023,
                    data: sibling_data.clone(),
                },
            );
            object.archive_info.message_infos.insert(
                style.message_index,
                MessageInfo::new(
                    2_023,
                    u32::try_from(sibling_data.len()).expect("test payload fits u32"),
                ),
            );
            Ok(())
        })
        .unwrap();

    let mut editor = IWorkTextEditor::from_package(package);
    editor
        .set_paragraph_alignment(storage_id, TextAlignment::Right)
        .unwrap();
    let package = editor.into_package();
    let style_after = native::locate_style(&package, storage.style_id).unwrap();
    assert_eq!(style_after.message_index, style.message_index + 1);
    let archive = package.archive(&style_after.archive_name).unwrap();
    let object = archive.object(style_after.object_id).unwrap();
    assert_eq!(object.messages[style.message_index].data, sibling_data);
    assert_eq!(
        object.archive_info.message_infos[style.message_index].type_,
        2_023
    );

    let mut stale = package.clone();
    let stale_located = native::locate_style_with_archive(&stale, style_after.object_id).unwrap();
    stale
        .update_archive(&style_after.archive_name, |archive| {
            let object = archive.object_mut(style_after.object_id).unwrap();
            object.replace_message(
                style_after.message_index,
                RawMessage {
                    type_: 99_999,
                    data: style_after.message.data.clone(),
                },
            )?;
            Ok(())
        })
        .unwrap();
    let before = stale.entry(&style_after.archive_name).unwrap().to_vec();
    let replacement = ArchiveObject::new(
        style_after.object_id,
        vec![RawMessage {
            type_: style_after.message_type,
            data: style_after.message.data.clone(),
        }],
    )
    .unwrap();
    assert!(
        native::replace_variation_with_archive(&mut stale, stale_located, replacement).is_err()
    );
    assert_eq!(
        stale.entry(&style_after.archive_name).unwrap(),
        before.as_slice()
    );
}

#[test]
fn named_paragraph_style_rename_uses_exact_native_anchor_with_sibling_payload() {
    let mut pages = PagesEditor::create_with_text("Rename anchor").unwrap();
    let text_box = pages
        .add_text_box(
            11,
            "Named style",
            DrawablePoint { x: 20.0, y: 40.0 },
            DrawableSize {
                width: 240.0,
                height: 100.0,
            },
        )
        .unwrap();
    let body = pages
        .text_box_named_paragraph_styles(text_box.drawable_object_id)
        .unwrap()
        .into_iter()
        .find(|style| style.name() == "Body")
        .unwrap()
        .id();
    let created = pages
        .create_text_box_named_paragraph_style(
            text_box.drawable_object_id,
            body,
            ParagraphStyleName::new("Rename source").unwrap(),
        )
        .unwrap();
    let storage_id = text_box.storage.object_id;
    let mut package = pages.into_package();
    let location = native::locate_style(&package, native_id(created.id())).unwrap();
    let sibling_data = vec![0x98, 0x06, 0x09];
    package
        .update_archive(&location.archive_name, |archive| {
            let object = archive.object_mut(location.object_id).unwrap();
            object.messages.insert(
                location.message_index,
                RawMessage {
                    type_: 2_023,
                    data: sibling_data.clone(),
                },
            );
            object.archive_info.message_infos.insert(
                location.message_index,
                MessageInfo::new(
                    2_023,
                    u32::try_from(sibling_data.len()).expect("test payload fits u32"),
                ),
            );
            Ok(())
        })
        .unwrap();

    let renamed = super::rename_named_paragraph_style(
        &mut package,
        storage_id,
        created.id(),
        ParagraphStyleName::new("Rename target").unwrap(),
    )
    .unwrap();
    assert_eq!(renamed.name(), "Rename target");
    let style_after = native::locate_style(&package, native_id(created.id())).unwrap();
    assert_eq!(style_after.message_index, location.message_index + 1);
    let archive = package.archive(&style_after.archive_name).unwrap();
    let object = archive.object(style_after.object_id).unwrap();
    assert_eq!(object.messages[location.message_index].data, sibling_data);
    assert_eq!(
        object.archive_info.message_infos[location.message_index].type_,
        2_023
    );

    let mut stale = package.clone();
    stale
        .update_archive(&style_after.archive_name, |archive| {
            let object = archive.object_mut(style_after.object_id).unwrap();
            object.replace_message(
                style_after.message_index,
                RawMessage {
                    type_: 99_999,
                    data: style_after.message.data.clone(),
                },
            )?;
            Ok(())
        })
        .unwrap();
    let before = stale.to_bytes().unwrap();
    assert!(
        super::super::paragraph_style_rename::rename_at_location(
            &mut stale,
            &style_after,
            "Must not publish",
        )
        .is_err()
    );
    assert_eq!(stale.to_bytes().unwrap(), before);
}
