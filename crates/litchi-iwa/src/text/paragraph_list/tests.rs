use prost::Message;

use super::*;
use crate::archive::RawMessage;
use crate::keynote::KeynoteDocumentBuilder;
use crate::numbers::NumbersDocumentBuilder;
use crate::pages::PagesEditor;
use crate::protobuf::tsp;
use crate::protobuf::tswp::{self, object_attribute_table::ObjectAttribute};
use crate::shapes::{DrawablePoint, DrawableSize, RgbColorSpace, RgbaColor};
use crate::text::{IWorkTextEditor, TextPointSize, TextPosition};

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
    let paragraph = TextPosition::from_utf16_index(6).unwrap();
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
fn text_boxes_round_trip_bullet_geometry_in_every_suite() {
    let paragraph = TextPosition::from_utf16_index(6).unwrap();
    let geometry = ParagraphListBulletGeometry::new(
        ParagraphListBulletScale::from_percent(150.0).unwrap(),
        ParagraphListBulletBaselineOffset::from_points(2.0).unwrap(),
    );

    let mut pages = PagesEditor::create_with_text("Bullet Geometry").unwrap();
    let pages_box = pages
        .add_text_box(5, "First\nSecond", POSITION, SIZE)
        .unwrap();
    pages
        .set_text_box_paragraph_list(pages_box.drawable_object_id, ParagraphList::Bullet)
        .unwrap();
    let before_geometry = pages.to_bytes().unwrap();
    pages
        .set_text_box_paragraph_list_bullet_geometry(
            pages_box.drawable_object_id,
            paragraph,
            geometry,
        )
        .unwrap();
    let mut pages = PagesEditor::from_bytes(&pages.to_bytes().unwrap()).unwrap();
    assert_eq!(
        pages
            .text_box_paragraph_list_bullet_geometry(pages_box.drawable_object_id, paragraph,)
            .unwrap(),
        geometry
    );
    assert!(
        pages
            .reset_text_box_paragraph_list_bullet_geometry(pages_box.drawable_object_id, paragraph,)
            .unwrap()
    );
    assert_eq!(pages.to_bytes().unwrap(), before_geometry);

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
        .set_sheet_text_box_paragraph_list_bullet_geometry(
            sheet_id,
            numbers_box.drawable_object_id,
            paragraph,
            geometry,
        )
        .unwrap();
    let numbers = crate::numbers::NumbersEditor::from_bytes(&numbers.to_bytes().unwrap()).unwrap();
    assert_eq!(
        numbers
            .sheet_text_box_paragraph_list_bullet_geometry(
                sheet_id,
                numbers_box.drawable_object_id,
                paragraph,
            )
            .unwrap(),
        geometry
    );

    let mut keynote = KeynoteDocumentBuilder::new().build().unwrap();
    let keynote_box = keynote
        .add_slide_text_box(0, "First\nSecond", POSITION, SIZE)
        .unwrap();
    keynote
        .set_slide_text_box_paragraph_list(0, keynote_box.drawable_object_id, ParagraphList::Bullet)
        .unwrap();
    keynote
        .set_slide_text_box_paragraph_list_bullet_geometry(
            0,
            keynote_box.drawable_object_id,
            paragraph,
            geometry,
        )
        .unwrap();
    let keynote = crate::keynote::KeynoteEditor::from_bytes(&keynote.to_bytes().unwrap()).unwrap();
    assert_eq!(
        keynote
            .slide_text_box_paragraph_list_bullet_geometry(
                0,
                keynote_box.drawable_object_id,
                paragraph,
            )
            .unwrap(),
        geometry
    );
}

#[test]
fn text_boxes_round_trip_list_indentation_in_every_suite() {
    let paragraph = TextPosition::from_utf16_index(6).unwrap();
    let indentation = ParagraphListIndentation::new(
        ParagraphListLabelIndent::from_points(20.0).unwrap(),
        ParagraphListTextGap::from_points(18.0, TextPointSize::TWELVE).unwrap(),
    );

    let mut pages = PagesEditor::create_with_text("List Indentation").unwrap();
    let pages_box = pages
        .add_text_box(5, "First\nSecond", POSITION, SIZE)
        .unwrap();
    pages
        .set_text_box_paragraph_list(pages_box.drawable_object_id, ParagraphList::Bullet)
        .unwrap();
    let pages_before_indentation = pages.to_bytes().unwrap();
    pages
        .set_text_box_paragraph_list_indentation(
            pages_box.drawable_object_id,
            paragraph,
            indentation,
        )
        .unwrap();
    let mut pages = PagesEditor::from_bytes(&pages.to_bytes().unwrap()).unwrap();
    assert_eq!(
        pages
            .text_box_paragraph_list_indentation(pages_box.drawable_object_id, paragraph)
            .unwrap(),
        indentation
    );
    assert!(
        pages
            .reset_text_box_paragraph_list_indentation(pages_box.drawable_object_id, paragraph,)
            .unwrap()
    );
    assert_eq!(pages.to_bytes().unwrap(), pages_before_indentation);

    let mut numbers = NumbersDocumentBuilder::new().build().unwrap();
    let sheet_id = numbers.sheets().unwrap()[0].object_id;
    let numbers_box = numbers
        .add_sheet_text_box(sheet_id, "First\nSecond", POSITION, SIZE)
        .unwrap();
    numbers
        .set_sheet_text_box_paragraph_list(
            sheet_id,
            numbers_box.drawable_object_id,
            ParagraphList::Numbered,
        )
        .unwrap();
    numbers
        .set_sheet_text_box_paragraph_list_indentation(
            sheet_id,
            numbers_box.drawable_object_id,
            paragraph,
            indentation,
        )
        .unwrap();
    let numbers = crate::numbers::NumbersEditor::from_bytes(&numbers.to_bytes().unwrap()).unwrap();
    assert_eq!(
        numbers
            .sheet_text_box_paragraph_list_indentation(
                sheet_id,
                numbers_box.drawable_object_id,
                paragraph,
            )
            .unwrap(),
        indentation
    );

    let mut keynote = KeynoteDocumentBuilder::new().build().unwrap();
    let keynote_box = keynote
        .add_slide_text_box(0, "First\nSecond", POSITION, SIZE)
        .unwrap();
    keynote
        .set_slide_text_box_paragraph_list(0, keynote_box.drawable_object_id, ParagraphList::Bullet)
        .unwrap();
    keynote
        .set_slide_text_box_paragraph_list_indentation(
            0,
            keynote_box.drawable_object_id,
            paragraph,
            indentation,
        )
        .unwrap();
    let keynote = crate::keynote::KeynoteEditor::from_bytes(&keynote.to_bytes().unwrap()).unwrap();
    assert_eq!(
        keynote
            .slide_text_box_paragraph_list_indentation(
                0,
                keynote_box.drawable_object_id,
                paragraph,
            )
            .unwrap(),
        indentation
    );
}

#[test]
fn text_boxes_round_trip_list_label_colors_in_every_suite() {
    let paragraph = TextPosition::from_utf16_index(6).unwrap();
    let color = ParagraphListLabelColor::Explicit(
        RgbaColor::new(0.9, 0.2, 0.1, 0.8, RgbColorSpace::DisplayP3).unwrap(),
    );

    let mut pages = PagesEditor::create_with_text("List Label Color").unwrap();
    let pages_box = pages
        .add_text_box(5, "First\nSecond", POSITION, SIZE)
        .unwrap();
    pages
        .set_text_box_paragraph_list(pages_box.drawable_object_id, ParagraphList::Bullet)
        .unwrap();
    let pages_before_color = pages.to_bytes().unwrap();
    pages
        .set_text_box_paragraph_list_label_color(pages_box.drawable_object_id, paragraph, color)
        .unwrap();
    let mut pages = PagesEditor::from_bytes(&pages.to_bytes().unwrap()).unwrap();
    assert_eq!(
        pages
            .text_box_paragraph_list_label_color(pages_box.drawable_object_id, paragraph)
            .unwrap(),
        color
    );
    assert!(
        pages
            .reset_text_box_paragraph_list_label_color(pages_box.drawable_object_id, paragraph,)
            .unwrap()
    );
    assert_eq!(pages.to_bytes().unwrap(), pages_before_color);

    let mut numbers = NumbersDocumentBuilder::new().build().unwrap();
    let sheet_id = numbers.sheets().unwrap()[0].object_id;
    let numbers_box = numbers
        .add_sheet_text_box(sheet_id, "First\nSecond", POSITION, SIZE)
        .unwrap();
    numbers
        .set_sheet_text_box_paragraph_list(
            sheet_id,
            numbers_box.drawable_object_id,
            ParagraphList::Numbered,
        )
        .unwrap();
    numbers
        .set_sheet_text_box_paragraph_list_label_color(
            sheet_id,
            numbers_box.drawable_object_id,
            paragraph,
            color,
        )
        .unwrap();
    let numbers = crate::numbers::NumbersEditor::from_bytes(&numbers.to_bytes().unwrap()).unwrap();
    assert_eq!(
        numbers
            .sheet_text_box_paragraph_list_label_color(
                sheet_id,
                numbers_box.drawable_object_id,
                paragraph,
            )
            .unwrap(),
        color
    );

    let mut keynote = KeynoteDocumentBuilder::new().build().unwrap();
    let keynote_box = keynote
        .add_slide_text_box(0, "First\nSecond", POSITION, SIZE)
        .unwrap();
    keynote
        .set_slide_text_box_paragraph_list(0, keynote_box.drawable_object_id, ParagraphList::Bullet)
        .unwrap();
    keynote
        .set_slide_text_box_paragraph_list_label_color(
            0,
            keynote_box.drawable_object_id,
            paragraph,
            color,
        )
        .unwrap();
    let keynote = crate::keynote::KeynoteEditor::from_bytes(&keynote.to_bytes().unwrap()).unwrap();
    assert_eq!(
        keynote
            .slide_text_box_paragraph_list_label_color(
                0,
                keynote_box.drawable_object_id,
                paragraph,
            )
            .unwrap(),
        color
    );
}

#[test]
fn text_boxes_round_trip_number_formats_in_every_suite() {
    let paragraph = TextPosition::from_utf16_index(6).unwrap();
    let format = ParagraphListNumberFormat::affixed(
        ParagraphListNumberSequence::RomanLowercase,
        ParagraphListNumberPunctuation::Parentheses,
    );

    let mut pages = PagesEditor::create_with_text("Number Format").unwrap();
    let pages_box = pages
        .add_text_box(5, "First\nSecond", POSITION, SIZE)
        .unwrap();
    pages
        .set_text_box_paragraph_list(pages_box.drawable_object_id, ParagraphList::Numbered)
        .unwrap();
    let pages_before_format = pages.to_bytes().unwrap();
    pages
        .set_text_box_paragraph_list_number_format(pages_box.drawable_object_id, paragraph, format)
        .unwrap();
    let mut pages = PagesEditor::from_bytes(&pages.to_bytes().unwrap()).unwrap();
    assert_eq!(
        pages
            .text_box_paragraph_list_number_format(pages_box.drawable_object_id, paragraph)
            .unwrap(),
        format
    );
    assert!(
        pages
            .reset_text_box_paragraph_list_number_format(pages_box.drawable_object_id, paragraph,)
            .unwrap()
    );
    assert_eq!(pages.to_bytes().unwrap(), pages_before_format);

    let mut numbers = NumbersDocumentBuilder::new().build().unwrap();
    let sheet_id = numbers.sheets().unwrap()[0].object_id;
    let numbers_box = numbers
        .add_sheet_text_box(sheet_id, "First\nSecond", POSITION, SIZE)
        .unwrap();
    numbers
        .set_sheet_text_box_paragraph_list(
            sheet_id,
            numbers_box.drawable_object_id,
            ParagraphList::Numbered,
        )
        .unwrap();
    numbers
        .set_sheet_text_box_paragraph_list_number_format(
            sheet_id,
            numbers_box.drawable_object_id,
            paragraph,
            ParagraphListNumberFormat::Circled,
        )
        .unwrap();
    let numbers = crate::numbers::NumbersEditor::from_bytes(&numbers.to_bytes().unwrap()).unwrap();
    assert_eq!(
        numbers
            .sheet_text_box_paragraph_list_number_format(
                sheet_id,
                numbers_box.drawable_object_id,
                paragraph,
            )
            .unwrap(),
        ParagraphListNumberFormat::Circled
    );

    let mut keynote = KeynoteDocumentBuilder::new().build().unwrap();
    let keynote_box = keynote
        .add_slide_text_box(0, "First\nSecond", POSITION, SIZE)
        .unwrap();
    keynote
        .set_slide_text_box_paragraph_list(
            0,
            keynote_box.drawable_object_id,
            ParagraphList::Numbered,
        )
        .unwrap();
    keynote
        .set_slide_text_box_paragraph_list_number_format(
            0,
            keynote_box.drawable_object_id,
            paragraph,
            ParagraphListNumberFormat::HebrewBiblicalStandard,
        )
        .unwrap();
    let keynote = crate::keynote::KeynoteEditor::from_bytes(&keynote.to_bytes().unwrap()).unwrap();
    assert_eq!(
        keynote
            .slide_text_box_paragraph_list_number_format(
                0,
                keynote_box.drawable_object_id,
                paragraph,
            )
            .unwrap(),
        ParagraphListNumberFormat::HebrewBiblicalStandard
    );
}

#[test]
fn text_boxes_round_trip_number_tiering_in_every_suite() {
    let paragraph = TextPosition::from_utf16_index(6).unwrap();
    let tiered = ParagraphListNumberTiering::Tiered;

    let mut pages = PagesEditor::create_with_text("Number Tiering").unwrap();
    let pages_box = pages
        .add_text_box(5, "First\nSecond", POSITION, SIZE)
        .unwrap();
    pages
        .set_text_box_paragraph_list(pages_box.drawable_object_id, ParagraphList::Numbered)
        .unwrap();
    let pages_before_tiering = pages.to_bytes().unwrap();
    pages
        .set_text_box_paragraph_list_number_tiering(pages_box.drawable_object_id, paragraph, tiered)
        .unwrap();
    let mut pages = PagesEditor::from_bytes(&pages.to_bytes().unwrap()).unwrap();
    assert_eq!(
        pages
            .text_box_paragraph_list_number_tiering(pages_box.drawable_object_id, paragraph)
            .unwrap(),
        tiered
    );
    assert!(
        pages
            .reset_text_box_paragraph_list_number_tiering(pages_box.drawable_object_id, paragraph,)
            .unwrap()
    );
    assert_eq!(pages.to_bytes().unwrap(), pages_before_tiering);

    let mut numbers = NumbersDocumentBuilder::new().build().unwrap();
    let sheet_id = numbers.sheets().unwrap()[0].object_id;
    let numbers_box = numbers
        .add_sheet_text_box(sheet_id, "First\nSecond", POSITION, SIZE)
        .unwrap();
    numbers
        .set_sheet_text_box_paragraph_list(
            sheet_id,
            numbers_box.drawable_object_id,
            ParagraphList::Numbered,
        )
        .unwrap();
    numbers
        .set_sheet_text_box_paragraph_list_number_tiering(
            sheet_id,
            numbers_box.drawable_object_id,
            paragraph,
            tiered,
        )
        .unwrap();
    let numbers = crate::numbers::NumbersEditor::from_bytes(&numbers.to_bytes().unwrap()).unwrap();
    assert_eq!(
        numbers
            .sheet_text_box_paragraph_list_number_tiering(
                sheet_id,
                numbers_box.drawable_object_id,
                paragraph,
            )
            .unwrap(),
        tiered
    );

    let mut keynote = KeynoteDocumentBuilder::new().build().unwrap();
    let keynote_box = keynote
        .add_slide_text_box(0, "First\nSecond", POSITION, SIZE)
        .unwrap();
    keynote
        .set_slide_text_box_paragraph_list(
            0,
            keynote_box.drawable_object_id,
            ParagraphList::Numbered,
        )
        .unwrap();
    keynote
        .set_slide_text_box_paragraph_list_number_tiering(
            0,
            keynote_box.drawable_object_id,
            paragraph,
            tiered,
        )
        .unwrap();
    let keynote = crate::keynote::KeynoteEditor::from_bytes(&keynote.to_bytes().unwrap()).unwrap();
    assert_eq!(
        keynote
            .slide_text_box_paragraph_list_number_tiering(
                0,
                keynote_box.drawable_object_id,
                paragraph,
            )
            .unwrap(),
        tiered
    );
}

#[test]
fn text_boxes_round_trip_number_scale_in_every_suite() {
    let paragraph = TextPosition::from_utf16_index(6).unwrap();
    let scale = ParagraphListNumberScale::from_percent(135.0).unwrap();

    let mut pages = PagesEditor::create_with_text("Number Scale").unwrap();
    let pages_box = pages
        .add_text_box(5, "First\nSecond", POSITION, SIZE)
        .unwrap();
    pages
        .set_text_box_paragraph_list(pages_box.drawable_object_id, ParagraphList::Numbered)
        .unwrap();
    pages
        .set_text_box_paragraph_list_number_format(
            pages_box.drawable_object_id,
            paragraph,
            ParagraphListNumberFormat::Circled,
        )
        .unwrap();
    pages
        .set_text_box_paragraph_list_number_tiering(
            pages_box.drawable_object_id,
            paragraph,
            ParagraphListNumberTiering::Tiered,
        )
        .unwrap();
    let pages_before_scale = pages.to_bytes().unwrap();
    pages
        .set_text_box_paragraph_list_number_scale(pages_box.drawable_object_id, paragraph, scale)
        .unwrap();
    let mut pages = PagesEditor::from_bytes(&pages.to_bytes().unwrap()).unwrap();
    assert_eq!(
        pages
            .text_box_paragraph_list_number_scale(pages_box.drawable_object_id, paragraph)
            .unwrap(),
        scale
    );
    assert_eq!(
        pages
            .text_box_paragraph_list_number_format(pages_box.drawable_object_id, paragraph)
            .unwrap(),
        ParagraphListNumberFormat::Circled
    );
    assert_eq!(
        pages
            .text_box_paragraph_list_number_tiering(pages_box.drawable_object_id, paragraph)
            .unwrap(),
        ParagraphListNumberTiering::Tiered
    );
    assert!(
        pages
            .reset_text_box_paragraph_list_number_scale(pages_box.drawable_object_id, paragraph)
            .unwrap()
    );
    assert_eq!(pages.to_bytes().unwrap(), pages_before_scale);

    let mut numbers = NumbersDocumentBuilder::new().build().unwrap();
    let sheet_id = numbers.sheets().unwrap()[0].object_id;
    let numbers_box = numbers
        .add_sheet_text_box(sheet_id, "First\nSecond", POSITION, SIZE)
        .unwrap();
    numbers
        .set_sheet_text_box_paragraph_list(
            sheet_id,
            numbers_box.drawable_object_id,
            ParagraphList::Numbered,
        )
        .unwrap();
    numbers
        .set_sheet_text_box_paragraph_list_number_scale(
            sheet_id,
            numbers_box.drawable_object_id,
            paragraph,
            scale,
        )
        .unwrap();
    let numbers = crate::numbers::NumbersEditor::from_bytes(&numbers.to_bytes().unwrap()).unwrap();
    assert_eq!(
        numbers
            .sheet_text_box_paragraph_list_number_scale(
                sheet_id,
                numbers_box.drawable_object_id,
                paragraph,
            )
            .unwrap(),
        scale
    );

    let mut keynote = KeynoteDocumentBuilder::new().build().unwrap();
    let keynote_box = keynote
        .add_slide_text_box(0, "First\nSecond", POSITION, SIZE)
        .unwrap();
    keynote
        .set_slide_text_box_paragraph_list(
            0,
            keynote_box.drawable_object_id,
            ParagraphList::Numbered,
        )
        .unwrap();
    keynote
        .set_slide_text_box_paragraph_list_number_scale(
            0,
            keynote_box.drawable_object_id,
            paragraph,
            scale,
        )
        .unwrap();
    let keynote = crate::keynote::KeynoteEditor::from_bytes(&keynote.to_bytes().unwrap()).unwrap();
    assert_eq!(
        keynote
            .slide_text_box_paragraph_list_number_scale(
                0,
                keynote_box.drawable_object_id,
                paragraph,
            )
            .unwrap(),
        scale
    );
}

#[test]
fn custom_list_updates_reject_nonlist_and_invalid_paragraphs_transactionally() {
    let arrow = ParagraphListBullet::new("➡").unwrap();
    let indentation = ParagraphListIndentation::new(
        ParagraphListLabelIndent::from_points(20.0).unwrap(),
        ParagraphListTextGap::from_em(1.5).unwrap(),
    );
    let mut pages = PagesEditor::create_with_text("Bullets").unwrap();
    let text_box = pages
        .add_text_box(5, "First\nSecond", POSITION, SIZE)
        .unwrap();
    let before = pages.to_bytes().unwrap();
    assert!(
        pages
            .set_text_box_paragraph_list_bullet(
                text_box.drawable_object_id,
                TextPosition::ZERO,
                &arrow,
            )
            .is_err()
    );
    assert_eq!(pages.to_bytes().unwrap(), before);
    assert!(
        pages
            .set_text_box_paragraph_list_indentation(
                text_box.drawable_object_id,
                TextPosition::ZERO,
                indentation,
            )
            .is_err()
    );
    assert_eq!(pages.to_bytes().unwrap(), before);
    assert!(
        pages
            .set_text_box_paragraph_list_number_format(
                text_box.drawable_object_id,
                TextPosition::ZERO,
                ParagraphListNumberFormat::Circled,
            )
            .is_err()
    );
    assert_eq!(pages.to_bytes().unwrap(), before);
    assert!(
        pages
            .set_text_box_paragraph_list_number_tiering(
                text_box.drawable_object_id,
                TextPosition::ZERO,
                ParagraphListNumberTiering::Tiered,
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
                TextPosition::from_utf16_index(1).unwrap(),
                &arrow,
            )
            .is_err()
    );
    assert_eq!(pages.to_bytes().unwrap(), before);
    assert!(
        pages
            .set_text_box_paragraph_list_indentation(
                text_box.drawable_object_id,
                TextPosition::from_utf16_index(1).unwrap(),
                indentation,
            )
            .is_err()
    );
    assert_eq!(pages.to_bytes().unwrap(), before);
    assert!(
        pages
            .set_text_box_paragraph_list_number_format(
                text_box.drawable_object_id,
                TextPosition::ZERO,
                ParagraphListNumberFormat::Circled,
            )
            .is_err()
    );
    assert_eq!(pages.to_bytes().unwrap(), before);
    assert!(
        pages
            .set_text_box_paragraph_list_number_tiering(
                text_box.drawable_object_id,
                TextPosition::ZERO,
                ParagraphListNumberTiering::Tiered,
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
