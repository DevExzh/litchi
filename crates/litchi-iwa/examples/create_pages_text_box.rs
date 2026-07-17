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
    ParagraphLineSpacingMultiple, ParagraphList, ParagraphListLevel, ParagraphSpacing,
    ParagraphSpacingPoints, ParagraphStart, ParagraphTabAlignment, ParagraphTabLeader,
    ParagraphTabPosition, ParagraphTabStop, ParagraphTabStops, TextAlignment, TextBackground,
    TextBaselineShift, TextCapitalization, TextCharacterSpacing, TextColumnCount, TextColumnGap,
    TextColumns, TextCommentBody, TextDecorations, TextFont, TextHyperlinkTarget, TextLanguage,
    TextLigatures, TextOutline, TextPointSize, TextPosition, TextRange, TextScript, TextShadow,
    TextStrikethrough, TextStyle, TextUnderline,
};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let output = arguments
        .next()
        .ok_or("usage: create_pages_text_box <output.pages> [text]")?;
    let text = arguments.next().unwrap_or_else(|| {
        "Overview\tPage 1 — created entirely from scratch.\nDétails\tÉlément imbriqué.".to_owned()
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
    editor.set_text_box_text_font(created.drawable_object_id, TextFont::named("Georgia-Bold")?)?;
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
    editor.set_text_box_text_character_spacing(
        created.drawable_object_id,
        TextCharacterSpacing::from_percent(12.0)?,
    )?;
    editor.set_text_box_text_ligatures(created.drawable_object_id, TextLigatures::RequiredOnly)?;
    editor.set_text_box_text_outline(created.drawable_object_id, TextOutline::standard())?;
    editor.set_text_box_text_shadow(created.drawable_object_id, TextShadow::standard())?;
    editor.set_text_box_text_background(
        created.drawable_object_id,
        TextBackground::Color(RgbaColor::new(1.0, 0.82, 0.72, 1.0, RgbColorSpace::Srgb)?),
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
    editor.set_text_box_paragraph_list(created.drawable_object_id, ParagraphList::Bullet)?;
    let first_word_end = text.find(char::is_whitespace).unwrap_or(text.len());
    editor.add_text_box_highlight(
        created.drawable_object_id,
        TextRange::from_utf16_indexes(0, text[..first_word_end].encode_utf16().count())?,
    )?;
    if let Some(tab) = text.find('\t') {
        let start_byte = tab + 1;
        let end_byte = text[start_byte..]
            .find(char::is_whitespace)
            .map_or(text.len(), |offset| start_byte + offset);
        editor.add_text_box_comment(
            created.drawable_object_id,
            TextRange::from_utf16_indexes(
                text[..start_byte].encode_utf16().count(),
                text[..end_byte].encode_utf16().count(),
            )?,
            TextCommentBody::new("Created by litchi-iwa")?,
        )?;
    }
    if let Some(newline) = text.find('\n') {
        let start_index = text[..=newline].encode_utf16().count();
        let start = ParagraphStart::from_utf16_index(start_index)?;
        editor.set_text_box_paragraph_list_level(
            created.drawable_object_id,
            start,
            ParagraphListLevel::ONE,
        )?;
        editor.set_text_box_text_language(
            created.drawable_object_id,
            TextPosition::from_utf16_index(start_index)?,
            TextLanguage::tag("fr-CA")?,
        )?;
        let word_end = text[newline + 1..]
            .find(char::is_whitespace)
            .map_or(text.len(), |offset| newline + 1 + offset);
        let end_index = text[..word_end].encode_utf16().count();
        editor.add_text_box_hyperlink(
            created.drawable_object_id,
            TextRange::new(
                TextPosition::from_utf16_index(start_index)?,
                TextPosition::from_utf16_index(end_index)?,
            )?,
            TextHyperlinkTarget::new("https://example.com/pages")?,
        )?;
    }
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
