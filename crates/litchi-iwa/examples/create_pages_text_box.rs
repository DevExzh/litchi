//! Create a Pages document and multi-column text box without an input package.

use std::env;

use litchi_iwa::pages::PagesEditor;
use litchi_iwa::shapes::{
    DrawablePoint, DrawableSize, RgbColorSpace, RgbaColor, ShapeTextAutoSize, ShapeTextInset,
    ShapeTextInsets, ShapeTextLayout, ShapeTextVerticalAlignment,
};
use litchi_iwa::text::{
    DropCapCharacterCount, DropCapLineCount, DropCapOutdent, DropCapPadding, DropCapRaisedLines,
    DropCapWrap, ParagraphDropCap, ParagraphIndentPoints, ParagraphIndents, ParagraphLineSpacing,
    ParagraphLineSpacingMultiple, ParagraphSpacing, ParagraphSpacingPoints, ParagraphStart,
    ParagraphTabAlignment, ParagraphTabLeader, ParagraphTabPosition, ParagraphTabStop,
    ParagraphTabStops, TextAlignment, TextBaselineShift, TextCapitalization, TextColumnCount,
    TextColumnGap, TextColumns, TextDecorations, TextPointSize, TextScript, TextStrikethrough,
    TextStyle, TextUnderline,
};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let output = arguments
        .next()
        .ok_or("usage: create_pages_text_box <output.pages> [text]")?;
    let text = arguments.next().unwrap_or_else(|| {
        "Overview\tPage 1 — a typed Pages text box created entirely from scratch.".to_owned()
    });
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = PagesEditor::create_with_text("Multi-column text box")?;
    let anchor = editor.body_text()?.encode_utf16().count();
    let created = editor.add_text_box(
        anchor,
        &text,
        DrawablePoint { x: 72.0, y: 144.0 },
        DrawableSize {
            width: 468.0,
            height: 360.0,
        },
    )?;
    editor.set_text_box_columns(
        created.drawable_object_id,
        &TextColumns::equal(
            TextColumnCount::new(2)?,
            Some(TextColumnGap::from_points(18.0)?),
        ),
    )?;
    editor.set_text_box_text_layout(
        created.drawable_object_id,
        ShapeTextLayout::new(
            ShapeTextVerticalAlignment::Middle,
            ShapeTextInsets::uniform(ShapeTextInset::from_points(9.0)?),
            ShapeTextAutoSize::ShrinkToFit,
        ),
    )?;
    editor.set_text_box_text_style(
        created.drawable_object_id,
        TextStyle::new(TextPointSize::from_points(19.5)?).with_bold(true),
    )?;
    editor.set_text_box_text_decorations(
        created.drawable_object_id,
        TextDecorations::new(TextUnderline::Single, TextStrikethrough::Single),
    )?;
    editor.set_text_box_text_color(
        created.drawable_object_id,
        RgbaColor::new(0.84, 0.16, 0.12, 1.0, RgbColorSpace::Srgb)?,
    )?;
    editor.set_text_box_text_capitalization(
        created.drawable_object_id,
        TextCapitalization::AllCaps,
    )?;
    editor.set_text_box_text_script(created.drawable_object_id, TextScript::Superscript)?;
    editor.set_text_box_text_baseline_shift(
        created.drawable_object_id,
        TextBaselineShift::from_points(4.0)?,
    )?;
    editor.set_text_box_paragraph_alignment(created.drawable_object_id, TextAlignment::Center)?;
    editor.set_text_box_paragraph_line_spacing(
        created.drawable_object_id,
        ParagraphLineSpacing::Relative(ParagraphLineSpacingMultiple::ONE_POINT_FIVE),
    )?;
    editor.set_text_box_paragraph_spacing(
        created.drawable_object_id,
        ParagraphSpacing::new(
            ParagraphSpacingPoints::from_points(9.0)?,
            ParagraphSpacingPoints::from_points(15.0)?,
        ),
    )?;
    editor.set_text_box_paragraph_indents(
        created.drawable_object_id,
        ParagraphIndents::new(
            ParagraphIndentPoints::from_points(26.0)?,
            ParagraphIndentPoints::from_points(12.5)?,
            ParagraphIndentPoints::from_points(12.0)?,
        ),
    )?;
    editor.set_text_box_paragraph_tab_stops(
        created.drawable_object_id,
        ParagraphTabStops::new(vec![
            ParagraphTabStop::new(
                ParagraphTabPosition::from_points(48.5)?,
                ParagraphTabAlignment::Left,
            ),
            ParagraphTabStop::new(
                ParagraphTabPosition::from_points(56.0)?,
                ParagraphTabAlignment::Center,
            )
            .with_leader(ParagraphTabLeader::new(".")?),
        ])?,
    )?;
    editor.set_text_box_paragraph_drop_cap(
        created.drawable_object_id,
        ParagraphStart::ZERO,
        ParagraphDropCap::new(DropCapLineCount::new(4)?, DropCapCharacterCount::new(2)?)
            .with_raised_lines(DropCapRaisedLines::new(1)?)
            .with_wrap(DropCapWrap::Contour)
            .with_padding(DropCapPadding::from_points(6.0)?)
            .with_outdent(DropCapOutdent::from_ratio(0.25)?),
    )?;
    editor.save(output)?;
    println!(
        "created two-column Pages text box {} with storage {}",
        created.drawable_object_id, created.storage.object_id
    );
    Ok(())
}
