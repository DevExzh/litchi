use super::{TextHyperlinkTarget, TextPosition, TextRange};
use crate::keynote::{KeynoteDocumentBuilder, KeynoteEditor};
use crate::numbers::{NumbersDocumentBuilder, NumbersEditor};
use crate::pages::PagesEditor;
use crate::shapes::{DrawablePoint, DrawableSize};
use litchi_iwa_text::date_time::{
    DisplayText, Format, FormatterStyle, Instant, LocaleIdentifier, Settings, UpdatePlan,
};

const DISPLAY: &str = "Friday, July 17, 2026";
const FORMAT: &str = "EEEE, MMMM d, y";
const NATIVE_SECONDS: f64 = 805_965_335.005_918;
const POSITION: DrawablePoint = DrawablePoint { x: 40.0, y: 80.0 };
const SIZE: DrawableSize = DrawableSize {
    width: 320.0,
    height: 140.0,
};

fn settings() -> Settings {
    Settings::fixed(
        Format::new(FORMAT).unwrap(),
        LocaleIdentifier::new("en_US").unwrap(),
        Instant::from_reference_date_seconds(NATIVE_SECONDS).unwrap(),
    )
    .with_styles(FormatterStyle::Full, FormatterStyle::None)
    .unwrap()
}

fn range(start: usize, end: usize) -> TextRange {
    TextRange::from_utf16_indexes(start, end).unwrap()
}

#[test]
fn scratch_pages_date_time_insert_update_remove_round_trips() {
    let mut pages = PagesEditor::create_with_text("Report: ").unwrap();
    let baseline = pages.to_bytes().unwrap();
    let position = TextPosition::from_utf16_index(8).unwrap();
    let created = pages
        .insert_body_date_time_field(position, DisplayText::new(DISPLAY).unwrap(), settings())
        .unwrap();
    assert_eq!(pages.body_text().unwrap(), format!("Report: {DISPLAY}"));
    assert_eq!(
        pages.body_date_time_fields().unwrap().as_slice(),
        std::slice::from_ref(&created)
    );

    let mut pages = PagesEditor::from_bytes(&pages.to_bytes().unwrap()).unwrap();
    let updated_settings = settings().with_update_plan(UpdatePlan::Automatic).unwrap();
    let updated = pages
        .update_body_date_time_field(created.id, created.range, updated_settings.clone())
        .unwrap();
    assert_eq!(updated.id, created.id);
    assert_eq!(updated.settings, updated_settings);
    assert_eq!(
        pages.remove_body_date_time_field(created.id).unwrap(),
        updated
    );
    assert!(pages.body_date_time_fields().unwrap().is_empty());
    assert_eq!(pages.body_text().unwrap(), format!("Report: {DISPLAY}"));

    let end = 8 + DISPLAY.encode_utf16().count();
    pages.replace_body_text(8..end, "").unwrap();
    assert_eq!(pages.to_bytes().unwrap(), baseline);
}

#[test]
fn attach_delete_cycle_is_byte_exact_and_overlap_is_transactional() {
    let mut pages = PagesEditor::create_with_text(DISPLAY).unwrap();
    let baseline = pages.to_bytes().unwrap();
    let whole = range(0, DISPLAY.encode_utf16().count());
    let created = pages.add_body_date_time_field(whole, settings()).unwrap();
    let before_overlap = pages.to_bytes().unwrap();
    let mut text = crate::text::IWorkTextEditor::from_package(pages.package().clone());
    assert!(
        text.add_text_hyperlink(
            pages.body_storage().unwrap().object_id,
            range(0, 6),
            TextHyperlinkTarget::new("https://example.com").unwrap(),
        )
        .is_err()
    );
    assert_eq!(text.to_bytes().unwrap(), before_overlap);
    pages.remove_body_date_time_field(created.id).unwrap();
    assert_eq!(pages.to_bytes().unwrap(), baseline);
}

#[test]
fn invalid_positions_and_text_replacement_are_safe() {
    let mut pages = PagesEditor::create_with_text("A😀B").unwrap();
    let baseline = pages.to_bytes().unwrap();
    assert!(
        pages
            .add_body_date_time_field(range(1, 2), settings())
            .is_err()
    );
    assert!(
        pages
            .insert_body_date_time_field(
                TextPosition::from_utf16_index(2).unwrap(),
                DisplayText::new(DISPLAY).unwrap(),
                settings(),
            )
            .is_err()
    );
    assert_eq!(pages.to_bytes().unwrap(), baseline);

    let field = pages
        .add_body_date_time_field(range(0, 1), settings())
        .unwrap();
    pages.replace_body_text(0..1, "A").unwrap();
    assert!(pages.body_date_time_fields().unwrap().is_empty());
    assert!(!pages.package().iwa_entry_names().any(|name| {
        pages
            .package()
            .archive(name)
            .unwrap()
            .object(field.id.object_id())
            .is_some()
    }));
    assert_eq!(pages.to_bytes().unwrap(), baseline);
}

#[test]
fn scratch_text_box_date_time_fields_work_across_the_suite() {
    let prefix = "When: ";
    let position = TextPosition::from_utf16_index(prefix.encode_utf16().count()).unwrap();

    let mut pages = PagesEditor::create_with_text("Body").unwrap();
    let pages_box = pages.add_text_box(4, prefix, POSITION, SIZE).unwrap();
    let pages_field = pages
        .insert_text_box_date_time_field(
            pages_box.drawable_object_id,
            position,
            DisplayText::new(DISPLAY).unwrap(),
            settings(),
        )
        .unwrap();
    let mut pages = PagesEditor::from_bytes(&pages.to_bytes().unwrap()).unwrap();
    assert_eq!(
        pages
            .text_box_date_time_fields(pages_box.drawable_object_id)
            .unwrap()
            .as_slice(),
        std::slice::from_ref(&pages_field)
    );
    pages
        .remove_text_box_date_time_field(pages_box.drawable_object_id, pages_field.id)
        .unwrap();

    let mut numbers = NumbersDocumentBuilder::new().build().unwrap();
    let sheet_id = numbers.sheets().unwrap()[0].object_id;
    let numbers_box = numbers
        .add_sheet_text_box(sheet_id, prefix, POSITION, SIZE)
        .unwrap();
    let numbers_field = numbers
        .insert_sheet_text_box_date_time_field(
            sheet_id,
            numbers_box.drawable_object_id,
            position,
            DisplayText::new(DISPLAY).unwrap(),
            settings(),
        )
        .unwrap();
    let mut numbers = NumbersEditor::from_bytes(&numbers.to_bytes().unwrap()).unwrap();
    assert_eq!(
        numbers
            .sheet_text_box_date_time_fields(sheet_id, numbers_box.drawable_object_id)
            .unwrap()
            .as_slice(),
        std::slice::from_ref(&numbers_field)
    );
    numbers
        .remove_sheet_text_box_date_time_field(
            sheet_id,
            numbers_box.drawable_object_id,
            numbers_field.id,
        )
        .unwrap();

    let mut keynote = KeynoteDocumentBuilder::new().build().unwrap();
    let keynote_box = keynote
        .add_slide_text_box(0, prefix, POSITION, SIZE)
        .unwrap();
    let keynote_field = keynote
        .insert_slide_text_box_date_time_field(
            0,
            keynote_box.drawable_object_id,
            position,
            DisplayText::new(DISPLAY).unwrap(),
            settings(),
        )
        .unwrap();
    let mut keynote = KeynoteEditor::from_bytes(&keynote.to_bytes().unwrap()).unwrap();
    assert_eq!(
        keynote
            .slide_text_box_date_time_fields(0, keynote_box.drawable_object_id)
            .unwrap()
            .as_slice(),
        std::slice::from_ref(&keynote_field)
    );
    keynote
        .remove_slide_text_box_date_time_field(0, keynote_box.drawable_object_id, keynote_field.id)
        .unwrap();
}
