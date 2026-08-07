//! Create a Pages document and multi-column text box without an input package.

use std::env;

use litchi_iwa::pages::PagesEditor;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize, Pattern, RgbColorSpace, RgbaColor, Width};
use litchi_iwa::text::layout::{AutoSize, Inset, Insets, Layout, VerticalAlignment};
use litchi_iwa::text::{
    Alignment, Background, Border, Borders, IndentPoints, Indents, LineSpacing,
    LineSpacingMultiple, Outline, ParagraphBackground, ParagraphDecimalTabCharacter,
    ParagraphDefaultTabInterval, ParagraphFlow, ParagraphFollowingStyle, ParagraphHyphenation,
    ParagraphList, ParagraphListLevel, ParagraphStyleName, ParagraphTabAlignment,
    ParagraphTabLeader, ParagraphTabPosition, ParagraphTabStop, ParagraphTabStops,
    ParagraphWritingDirection, Shadow, Spacing, SpacingPoints, TextBaselineShift,
    TextCapitalization, TextCharacterSpacing, TextCommentBody, TextCommentReplyBody,
    TextDecorations, TextFont, TextHyperlinkTarget, TextLanguage, TextLigatures, TextPointSize,
    TextRange, TextScript, TextStrikethrough, TextStyle, TextUnderline,
};
use litchi_iwa_text::columns::{Columns, Count, Gap};
use litchi_iwa_text::paragraph::border::{Offset as BorderOffset, Sides as BorderSides};
use litchi_iwa_text::paragraph::drop_cap::{
    CharacterCount, DropCap, LineCount, Outdent, Padding, RaisedLines, Wrap,
};
use litchi_iwa_text::position::TextPosition;

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
        &Columns::equal(Count::new(2)?, Some(Gap::from_points(18.0)?)),
    )?;
    editor.set_text_box_text_layout(
        created.drawable_object_id,
        Layout::new(
            VerticalAlignment::Middle,
            Insets::uniform(Inset::from_points(9.0)?),
            AutoSize::ShrinkToFit,
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
    editor.set_text_box_text_outline(created.drawable_object_id, Outline::standard())?;
    editor.set_text_box_text_shadow(created.drawable_object_id, Shadow::standard())?;
    editor.set_text_box_text_background(
        created.drawable_object_id,
        Background::Color(RgbaColor::new(1.0, 0.82, 0.72, 1.0, RgbColorSpace::Srgb)?),
    )?;
    editor.set_text_box_paragraph_background(
        created.drawable_object_id,
        ParagraphBackground::Color(RgbaColor::new(
            1.0,
            0.588_738_74,
            0.552_926_2,
            1.0,
            RgbColorSpace::Srgb,
        )?),
    )?;
    editor.set_text_box_paragraph_borders(
        created.drawable_object_id,
        Borders::Bordered(Border::new(
            RgbaColor::black(),
            Width::new(3.0)?,
            Pattern::Solid,
            BorderSides::ALL,
            BorderOffset::from_points(9.0)?,
            true,
        )?),
    )?;
    editor.set_text_box_paragraph_flow(
        created.drawable_object_id,
        ParagraphFlow::new()
            .with_keep_lines_together(true)
            .with_keep_with_next(true)
            .with_start_on_new_page(true)
            .with_prevent_widow_orphan_lines(false)
            .with_hyphenation(ParagraphHyphenation::Prevented),
    )?;
    editor.set_text_box_paragraph_writing_direction(
        created.drawable_object_id,
        ParagraphWritingDirection::LeftToRight,
    )?;
    editor.set_text_box_paragraph_alignment(created.drawable_object_id, Alignment::Center)?;
    editor.set_text_box_paragraph_line_spacing(
        created.drawable_object_id,
        LineSpacing::Relative(LineSpacingMultiple::ONE_POINT_FIVE),
    )?;
    editor.set_text_box_paragraph_spacing(
        created.drawable_object_id,
        Spacing::new(
            SpacingPoints::from_points(9.0)?,
            SpacingPoints::from_points(15.0)?,
        ),
    )?;
    editor.set_text_box_paragraph_indents(
        created.drawable_object_id,
        Indents::new(
            IndentPoints::from_points(26.0)?,
            IndentPoints::from_points(12.5)?,
            IndentPoints::from_points(12.0)?,
        ),
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
        let comment = editor.add_text_box_comment(
            created.drawable_object_id,
            TextRange::from_utf16_indexes(
                text[..start_byte].encode_utf16().count(),
                text[..end_byte].encode_utf16().count(),
            )?,
            TextCommentBody::new("Created by litchi-iwa")?,
        )?;
        editor.add_text_box_comment_reply(
            created.drawable_object_id,
            comment.id(),
            TextCommentReplyBody::new("Created reply by litchi-iwa")?,
        )?;
    }
    if let Some(newline) = text.find('\n') {
        let start_index = text[..=newline].encode_utf16().count();
        let start = TextPosition::from_utf16_index(start_index)?;
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
        TextPosition::ZERO,
        DropCap::new(LineCount::new(4)?, CharacterCount::new(2)?)
            .with_raised_lines(RaisedLines::new(1)?)
            .with_wrap(Wrap::Contour)
            .with_padding(Padding::from_points(6.0)?)
            .with_outdent(Outdent::from_ratio(0.25)?),
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
    let tab_defaults_box = editor.add_text_box(
        anchor,
        "Amount\t12,34",
        DrawablePoint { x: 72.0, y: 540.0 },
        DrawableSize {
            width: 280.0,
            height: 72.0,
        },
    )?;
    editor.set_text_box_paragraph_decimal_tab_character(
        tab_defaults_box.drawable_object_id,
        ParagraphDecimalTabCharacter::COMMA,
    )?;
    editor.set_text_box_paragraph_default_tab_interval(
        tab_defaults_box.drawable_object_id,
        ParagraphDefaultTabInterval::from_points(54.0)?,
    )?;
    let following_style_box = editor.add_text_box(
        anchor,
        "Press Return for Body style",
        DrawablePoint { x: 360.0, y: 540.0 },
        DrawableSize {
            width: 220.0,
            height: 72.0,
        },
    )?;
    let body = editor
        .text_box_named_paragraph_styles(following_style_box.drawable_object_id)?
        .into_iter()
        .find(|style| style.name() == "Body")
        .ok_or("source-built Pages theme has no Body paragraph style")?;
    let heading = editor.create_text_box_named_paragraph_style(
        following_style_box.drawable_object_id,
        body.id(),
        ParagraphStyleName::new("Litchi Heading")?,
    )?;
    let display = editor.rename_text_box_named_paragraph_style(
        following_style_box.drawable_object_id,
        heading.id(),
        ParagraphStyleName::new("Litchi Display")?,
    )?;
    editor.apply_text_box_named_paragraph_style(
        following_style_box.drawable_object_id,
        display.id(),
    )?;
    let disposable = editor.create_text_box_named_paragraph_style(
        following_style_box.drawable_object_id,
        body.id(),
        ParagraphStyleName::new("Disposable")?,
    )?;
    editor.delete_text_box_named_paragraph_style(
        following_style_box.drawable_object_id,
        disposable.id(),
    )?;
    editor.set_text_box_paragraph_following_style(
        following_style_box.drawable_object_id,
        ParagraphFollowingStyle::Named(body.id()),
    )?;
    editor.set_text_box_paragraph_alignment(
        following_style_box.drawable_object_id,
        Alignment::Center,
    )?;
    editor
        .redefine_applied_text_box_named_paragraph_style(following_style_box.drawable_object_id)?;
    editor.save(output)?;
    println!(
        "created two-column Pages text box {} with storage {}",
        created.drawable_object_id, created.storage.id
    );
    Ok(())
}
