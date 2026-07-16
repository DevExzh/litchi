//! Create a Keynote presentation and ordinary text box without an input package.

use std::env;

use litchi_iwa::keynote::KeynoteDocumentBuilder;
use litchi_iwa::shapes::{
    DrawablePoint, DrawableSize, RgbColorSpace, RgbaColor, ShapeTextAutoSize, ShapeTextInset,
    ShapeTextInsets, ShapeTextLayout, ShapeTextVerticalAlignment,
};
use litchi_iwa::text::{
    DropCapCharacterCount, DropCapLineCount, DropCapOutdent, DropCapPadding, DropCapRaisedLines,
    DropCapWrap, ParagraphDropCap, ParagraphIndentPoints, ParagraphIndents, ParagraphLineSpacing,
    ParagraphLineSpacingPoints, ParagraphSpacing, ParagraphSpacingPoints, ParagraphStart,
    ParagraphTabAlignment, ParagraphTabLeader, ParagraphTabPosition, ParagraphTabStop,
    ParagraphTabStops, TextAlignment, TextBaselineShift, TextCapitalization, TextCharacterSpacing,
    TextColumnCount, TextColumns, TextDecorations, TextPointSize, TextScript, TextStrikethrough,
    TextStyle, TextUnderline,
};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let output = arguments
        .next()
        .ok_or("usage: create_keynote_text_box <output.key> [text]")?;
    let text = arguments
        .next()
        .unwrap_or_else(|| "Quarterly result\t42.50 — built from typed IWA objects".to_owned());
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = KeynoteDocumentBuilder::new()
        .title("Created from scratch")
        .subtitle("No embedded package or source drawable")
        .build()?;
    let created = editor.add_slide_text_box(
        0,
        &text,
        DrawablePoint { x: 144.0, y: 720.0 },
        DrawableSize {
            width: 1_200.0,
            height: 120.0,
        },
    )?;
    editor.set_slide_text_box_columns(
        0,
        created.drawable_object_id,
        &TextColumns::equal(TextColumnCount::new(4)?, None),
    )?;
    editor.set_slide_text_box_text_layout(
        0,
        created.drawable_object_id,
        ShapeTextLayout::new(
            ShapeTextVerticalAlignment::Middle,
            ShapeTextInsets::uniform(ShapeTextInset::from_points(12.0)?),
            ShapeTextAutoSize::ShrinkToFit,
        ),
    )?;
    editor.set_slide_text_box_text_style(
        0,
        created.drawable_object_id,
        TextStyle::new(TextPointSize::from_points(23.0)?)
            .with_bold(true)
            .with_italic(true),
    )?;
    editor.set_slide_text_box_text_decorations(
        0,
        created.drawable_object_id,
        TextDecorations::new(TextUnderline::Single, TextStrikethrough::Single),
    )?;
    editor.set_slide_text_box_text_color(
        0,
        created.drawable_object_id,
        RgbaColor::new(0.05, 0.42, 0.95, 1.0, RgbColorSpace::Srgb)?,
    )?;
    editor.set_slide_text_box_text_capitalization(
        0,
        created.drawable_object_id,
        TextCapitalization::TitleCase,
    )?;
    editor.set_slide_text_box_text_script(
        0,
        created.drawable_object_id,
        TextScript::Superscript,
    )?;
    editor.set_slide_text_box_text_baseline_shift(
        0,
        created.drawable_object_id,
        TextBaselineShift::from_points(5.0)?,
    )?;
    editor.set_slide_text_box_text_character_spacing(
        0,
        created.drawable_object_id,
        TextCharacterSpacing::from_percent(6.0)?,
    )?;
    editor.set_slide_text_box_paragraph_alignment(
        0,
        created.drawable_object_id,
        TextAlignment::Justified,
    )?;
    editor.set_slide_text_box_paragraph_line_spacing(
        0,
        created.drawable_object_id,
        ParagraphLineSpacing::Between(ParagraphLineSpacingPoints::from_points(6.0)?),
    )?;
    editor.set_slide_text_box_paragraph_spacing(
        0,
        created.drawable_object_id,
        ParagraphSpacing::new(
            ParagraphSpacingPoints::from_points(13.0)?,
            ParagraphSpacingPoints::from_points(19.0)?,
        ),
    )?;
    editor.set_slide_text_box_paragraph_indents(
        0,
        created.drawable_object_id,
        ParagraphIndents::new(
            ParagraphIndentPoints::from_points(23.0)?,
            ParagraphIndentPoints::from_points(13.0)?,
            ParagraphIndentPoints::from_points(10.5)?,
        ),
    )?;
    editor.set_slide_text_box_paragraph_tab_stops(
        0,
        created.drawable_object_id,
        ParagraphTabStops::new(vec![
            ParagraphTabStop::new(
                ParagraphTabPosition::from_points(63.0)?,
                ParagraphTabAlignment::Decimal,
            )
            .with_leader(ParagraphTabLeader::new(".")?),
        ])?,
    )?;
    editor.set_slide_text_box_paragraph_drop_cap(
        0,
        created.drawable_object_id,
        ParagraphStart::ZERO,
        ParagraphDropCap::new(DropCapLineCount::new(6)?, DropCapCharacterCount::new(2)?)
            .with_raised_lines(DropCapRaisedLines::new(3)?)
            .with_wrap(DropCapWrap::Contour)
            .with_padding(DropCapPadding::from_points(8.0)?)
            .with_outdent(DropCapOutdent::from_ratio(0.40)?),
    )?;
    editor.save(output)?;
    println!(
        "created four-column Keynote text box {} with storage {}",
        created.drawable_object_id, created.storage.object_id
    );
    Ok(())
}
