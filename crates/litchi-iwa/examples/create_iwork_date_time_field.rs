use std::path::Path;

use litchi_iwa::keynote::KeynoteDocumentBuilder;
use litchi_iwa::numbers::NumbersDocumentBuilder;
use litchi_iwa::pages::PagesEditor;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};
use litchi_iwa::text::TextPosition;
use litchi_iwa_text::date_time::{
    DisplayText, Format, FormatterStyle, Instant, LocaleIdentifier, Settings,
};

const PREFIX: &str = "Created: ";
const DISPLAY: &str = "Friday, July 17, 2026";
const APPLE_REFERENCE_DATE_SECONDS: f64 = 805_965_335.005_918;
const POSITION: DrawablePoint = DrawablePoint { x: 80.0, y: 100.0 };
const SIZE: DrawableSize = DrawableSize {
    width: 420.0,
    height: 120.0,
};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = std::env::args()
        .nth(1)
        .ok_or("usage: create_iwork_date_time_field <output.pages|output.numbers|output.key>")?;
    let position = TextPosition::from_utf16_index(PREFIX.encode_utf16().count())?;
    let display = || DisplayText::new(DISPLAY);
    match Path::new(&output)
        .extension()
        .and_then(|value| value.to_str())
    {
        Some("pages") => {
            let mut editor = PagesEditor::create_with_text(PREFIX)?;
            editor.insert_body_date_time_field(position, display()?, settings()?)?;
            editor.save(output)?;
        },
        Some("numbers") => {
            let mut editor = NumbersDocumentBuilder::new().build()?;
            let sheet_id = editor.sheets()?[0].id();
            let text_box = editor.add_sheet_text_box(sheet_id, PREFIX, POSITION, SIZE)?;
            editor.insert_sheet_text_box_date_time_field(
                sheet_id,
                text_box.drawable_object_id,
                position,
                display()?,
                settings()?,
            )?;
            editor.save(output)?;
        },
        Some("key") => {
            let mut editor = KeynoteDocumentBuilder::new().build()?;
            let text_box = editor.add_slide_text_box(0, PREFIX, POSITION, SIZE)?;
            editor.insert_slide_text_box_date_time_field(
                0,
                text_box.drawable_object_id,
                position,
                display()?,
                settings()?,
            )?;
            editor.save(output)?;
        },
        extension => return Err(format!("unsupported iWork extension: {extension:?}").into()),
    }
    Ok(())
}

fn settings() -> litchi_iwa::Result<Settings> {
    Ok(Settings::fixed(
        Format::new("EEEE, MMMM d, y")?,
        LocaleIdentifier::new("en_US")?,
        Instant::from_reference_date_seconds(APPLE_REFERENCE_DATE_SECONDS)?,
    )
    .with_styles(FormatterStyle::Full, FormatterStyle::None)?)
}
