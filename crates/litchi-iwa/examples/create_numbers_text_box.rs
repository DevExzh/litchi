//! Create a Numbers spreadsheet and text box without an input package.

use std::env;

use litchi_iwa::numbers::NumbersDocumentBuilder;
use litchi_iwa::shapes::{
    DrawablePoint, DrawableSize, RgbColorSpace, RgbaColor, ShapeTextAutoSize, ShapeTextInset,
    ShapeTextInsets, ShapeTextLayout, ShapeTextVerticalAlignment,
};
use litchi_iwa::text::{
    DropCapCharacterCount, DropCapLineCount, DropCapOutdent, DropCapPadding, DropCapRaisedLines,
    ParagraphDropCap, ParagraphIndentPoints, ParagraphIndents, ParagraphLineSpacing,
    ParagraphLineSpacingPoints, ParagraphSpacing, ParagraphSpacingPoints, ParagraphStart,
    ParagraphTabAlignment, ParagraphTabLeader, ParagraphTabPosition, ParagraphTabStop,
    ParagraphTabStops, TextAlignment, TextBaselineShift, TextCapitalization, TextCharacterSpacing,
    TextColumnCount, TextColumnGap, TextColumns, TextDecorations, TextLigatures, TextOutline,
    TextPointSize, TextScript, TextStrikethrough, TextStyle, TextUnderline,
};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let output = arguments
        .next()
        .ok_or("usage: create_numbers_text_box <output.numbers> [text]")?;
    let text = arguments
        .next()
        .unwrap_or_else(|| "Revenue\t42.50 — built from typed IWA objects".to_owned());
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = NumbersDocumentBuilder::new()
        .sheet_name("Scratch Sheet")
        .table_name("Scratch Table")
        .build()?;
    let sheet_id = editor.sheets()?[0].object_id;
    let created = editor.add_sheet_text_box(
        sheet_id,
        &text,
        DrawablePoint { x: 40.0, y: 300.0 },
        DrawableSize {
            width: 540.0,
            height: 240.0,
        },
    )?;
    editor.set_sheet_text_box_columns(
        sheet_id,
        created.drawable_object_id,
        &TextColumns::equal(
            TextColumnCount::new(3)?,
            Some(TextColumnGap::from_points(12.0)?),
        ),
    )?;
    editor.set_sheet_text_box_text_layout(
        sheet_id,
        created.drawable_object_id,
        ShapeTextLayout::new(
            ShapeTextVerticalAlignment::Bottom,
            ShapeTextInsets::uniform(ShapeTextInset::from_points(6.0)?),
            ShapeTextAutoSize::Fixed,
        ),
    )?;
    editor.set_sheet_text_box_text_style(
        sheet_id,
        created.drawable_object_id,
        TextStyle::new(TextPointSize::from_points(21.0)?).with_italic(true),
    )?;
    editor.set_sheet_text_box_text_decorations(
        sheet_id,
        created.drawable_object_id,
        TextDecorations::new(TextUnderline::Single, TextStrikethrough::Single),
    )?;
    editor.set_sheet_text_box_text_color(
        sheet_id,
        created.drawable_object_id,
        RgbaColor::new(0.12, 0.62, 0.24, 1.0, RgbColorSpace::Srgb)?,
    )?;
    editor.set_sheet_text_box_text_capitalization(
        sheet_id,
        created.drawable_object_id,
        TextCapitalization::SmallCaps,
    )?;
    editor.set_sheet_text_box_text_script(
        sheet_id,
        created.drawable_object_id,
        TextScript::Subscript,
    )?;
    editor.set_sheet_text_box_text_baseline_shift(
        sheet_id,
        created.drawable_object_id,
        TextBaselineShift::from_points(-3.0)?,
    )?;
    editor.set_sheet_text_box_text_character_spacing(
        sheet_id,
        created.drawable_object_id,
        TextCharacterSpacing::from_percent(-8.0)?,
    )?;
    editor.set_sheet_text_box_text_ligatures(
        sheet_id,
        created.drawable_object_id,
        TextLigatures::All,
    )?;
    editor.set_sheet_text_box_text_outline(
        sheet_id,
        created.drawable_object_id,
        TextOutline::standard(),
    )?;
    editor.set_sheet_text_box_paragraph_alignment(
        sheet_id,
        created.drawable_object_id,
        TextAlignment::Right,
    )?;
    editor.set_sheet_text_box_paragraph_line_spacing(
        sheet_id,
        created.drawable_object_id,
        ParagraphLineSpacing::Exactly(ParagraphLineSpacingPoints::from_points(24.0)?),
    )?;
    editor.set_sheet_text_box_paragraph_spacing(
        sheet_id,
        created.drawable_object_id,
        ParagraphSpacing::new(
            ParagraphSpacingPoints::from_points(11.0)?,
            ParagraphSpacingPoints::from_points(17.0)?,
        ),
    )?;
    editor.set_sheet_text_box_paragraph_indents(
        sheet_id,
        created.drawable_object_id,
        ParagraphIndents::new(
            ParagraphIndentPoints::from_points(23.0)?,
            ParagraphIndentPoints::from_points(13.0)?,
            ParagraphIndentPoints::from_points(2.833_333_3)?,
        ),
    )?;
    editor.set_sheet_text_box_paragraph_tab_stops(
        sheet_id,
        created.drawable_object_id,
        ParagraphTabStops::new(vec![
            ParagraphTabStop::new(
                ParagraphTabPosition::from_points(43.0)?,
                ParagraphTabAlignment::Right,
            )
            .with_leader(ParagraphTabLeader::new("-")?),
        ])?,
    )?;
    editor.set_sheet_text_box_paragraph_drop_cap(
        sheet_id,
        created.drawable_object_id,
        ParagraphStart::ZERO,
        ParagraphDropCap::new(DropCapLineCount::new(5)?, DropCapCharacterCount::new(1)?)
            .with_raised_lines(DropCapRaisedLines::new(2)?)
            .with_padding(DropCapPadding::from_points(4.0)?)
            .with_outdent(DropCapOutdent::from_ratio(0.10)?),
    )?;
    editor.save(output)?;
    println!(
        "created three-column Numbers text box {} with storage {} on sheet {}",
        created.drawable_object_id, created.storage.object_id, sheet_id
    );
    Ok(())
}
