//! Create a Keynote presentation and ordinary text box without an input package.

use std::env;

use litchi_iwa::keynote::KeynoteDocumentBuilder;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize, Pattern, RgbColorSpace, RgbaColor, Width};
use litchi_iwa::text::layout::{AutoSize, Inset, Insets, Layout, VerticalAlignment};
use litchi_iwa::text::{
    DropCapCharacterCount, DropCapLineCount, DropCapOutdent, DropCapPadding, DropCapRaisedLines,
    DropCapWrap, ParagraphBackground, ParagraphBorder, ParagraphBorderOffset, ParagraphBorderSides,
    ParagraphBorders, ParagraphDecimalTabCharacter, ParagraphDefaultTabInterval, ParagraphDropCap,
    ParagraphFlow, ParagraphHyphenation, ParagraphIndentPoints, ParagraphIndents,
    ParagraphLineSpacing, ParagraphLineSpacingPoints, ParagraphList, ParagraphListLevel,
    ParagraphSpacing, ParagraphSpacingPoints, ParagraphStart, ParagraphStyleName,
    ParagraphTabAlignment, ParagraphTabLeader, ParagraphTabPosition, ParagraphTabStop,
    ParagraphTabStops, ParagraphWritingDirection, TextAlignment, TextBackground, TextBaselineShift,
    TextCapitalization, TextCharacterSpacing, TextColumnCount, TextColumns, TextCommentBody,
    TextCommentReplyBody, TextDecorations, TextFont, TextHyperlinkTarget, TextLanguage,
    TextLigatures, TextOutline, TextPointSize, TextPosition, TextRange, TextScript, TextShadow,
    TextStrikethrough, TextStyle, TextUnderline,
};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let output = arguments
        .next()
        .ok_or("usage: create_keynote_text_box <output.key> [text]")?;
    let text = arguments.next().unwrap_or_else(|| {
        "Quarterly result\t42.50 — built from typed IWA objects\nPrévisions\tÉlément imbriqué"
            .to_owned()
    });
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
        Layout::new(
            VerticalAlignment::Middle,
            Insets::uniform(Inset::from_points(12.0)?),
            AutoSize::ShrinkToFit,
        ),
    )?;
    editor.set_slide_text_box_text_style(
        0,
        created.drawable_object_id,
        TextStyle::new(TextPointSize::from_points(23.0)?)
            .with_bold(true)
            .with_italic(true),
    )?;
    editor.set_slide_text_box_text_font(
        0,
        created.drawable_object_id,
        TextFont::named("AvenirNext-BoldItalic")?,
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
    editor.set_slide_text_box_text_ligatures(
        0,
        created.drawable_object_id,
        TextLigatures::Standard,
    )?;
    editor.set_slide_text_box_text_outline(
        0,
        created.drawable_object_id,
        TextOutline::standard(),
    )?;
    editor.set_slide_text_box_text_shadow(0, created.drawable_object_id, TextShadow::standard())?;
    editor.set_slide_text_box_text_background(
        0,
        created.drawable_object_id,
        TextBackground::Color(RgbaColor::new(0.72, 0.84, 1.0, 1.0, RgbColorSpace::Srgb)?),
    )?;
    editor.set_slide_text_box_paragraph_background(
        0,
        created.drawable_object_id,
        ParagraphBackground::Color(RgbaColor::new(
            1.0,
            0.588_738_74,
            0.552_926_2,
            1.0,
            RgbColorSpace::Srgb,
        )?),
    )?;
    editor.set_slide_text_box_paragraph_borders(
        0,
        created.drawable_object_id,
        ParagraphBorders::Bordered(ParagraphBorder::new(
            RgbaColor::black(),
            Width::new(3.0)?,
            Pattern::Solid,
            ParagraphBorderSides::ALL,
            ParagraphBorderOffset::from_points(9.0)?,
            true,
        )?),
    )?;
    editor.set_slide_text_box_paragraph_flow(
        0,
        created.drawable_object_id,
        ParagraphFlow::new()
            .with_keep_lines_together(true)
            .with_keep_with_next(true)
            .with_start_on_new_page(true)
            .with_prevent_widow_orphan_lines(false)
            .with_hyphenation(ParagraphHyphenation::Prevented),
    )?;
    editor.set_slide_text_box_paragraph_writing_direction(
        0,
        created.drawable_object_id,
        ParagraphWritingDirection::RightToLeft,
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
    editor.set_slide_text_box_paragraph_list(
        0,
        created.drawable_object_id,
        ParagraphList::Bullet,
    )?;
    let first_word_end = text.find(char::is_whitespace).unwrap_or(text.len());
    editor.add_slide_text_box_highlight(
        0,
        created.drawable_object_id,
        TextRange::from_utf16_indexes(0, text[..first_word_end].encode_utf16().count())?,
    )?;
    if let Some(tab) = text.find('\t') {
        let start_byte = tab + 1;
        let end_byte = text[start_byte..]
            .find(char::is_whitespace)
            .map_or(text.len(), |offset| start_byte + offset);
        let comment = editor.add_slide_text_box_comment(
            0,
            created.drawable_object_id,
            TextRange::from_utf16_indexes(
                text[..start_byte].encode_utf16().count(),
                text[..end_byte].encode_utf16().count(),
            )?,
            TextCommentBody::new("Created by litchi-iwa")?,
        )?;
        editor.add_slide_text_box_comment_reply(
            0,
            created.drawable_object_id,
            comment.id,
            TextCommentReplyBody::new("Created reply by litchi-iwa")?,
        )?;
    }
    if let Some(newline) = text.find('\n') {
        let start_index = text[..=newline].encode_utf16().count();
        let start = ParagraphStart::from_utf16_index(start_index)?;
        editor.set_slide_text_box_paragraph_list_level(
            0,
            created.drawable_object_id,
            start,
            ParagraphListLevel::ONE,
        )?;
        editor.set_slide_text_box_text_language(
            0,
            created.drawable_object_id,
            TextPosition::from_utf16_index(start_index)?,
            TextLanguage::tag("fr-CA")?,
        )?;
        let word_end = text[newline + 1..]
            .find(char::is_whitespace)
            .map_or(text.len(), |offset| newline + 1 + offset);
        let end_index = text[..word_end].encode_utf16().count();
        editor.add_slide_text_box_hyperlink(
            0,
            created.drawable_object_id,
            TextRange::new(
                TextPosition::from_utf16_index(start_index)?,
                TextPosition::from_utf16_index(end_index)?,
            )?,
            TextHyperlinkTarget::new("https://example.com/keynote")?,
        )?;
    }
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
    let tab_defaults_box = editor.add_slide_text_box(
        0,
        "Amount\t12,34",
        DrawablePoint { x: 144.0, y: 880.0 },
        DrawableSize {
            width: 560.0,
            height: 96.0,
        },
    )?;
    editor.set_slide_text_box_paragraph_decimal_tab_character(
        0,
        tab_defaults_box.drawable_object_id,
        ParagraphDecimalTabCharacter::COMMA,
    )?;
    editor.set_slide_text_box_paragraph_default_tab_interval(
        0,
        tab_defaults_box.drawable_object_id,
        ParagraphDefaultTabInterval::from_points(48.0)?,
    )?;
    let named_style_box = editor.add_slide_text_box(
        0,
        "Keynote named paragraph style",
        DrawablePoint { x: 760.0, y: 880.0 },
        DrawableSize {
            width: 560.0,
            height: 96.0,
        },
    )?;
    let body = editor
        .slide_text_box_named_paragraph_styles(0, named_style_box.drawable_object_id)?
        .into_iter()
        .find(|style| style.name() == "Body")
        .ok_or("source-built Keynote theme has no Body paragraph style")?;
    let draft = editor.create_slide_text_box_named_paragraph_style(
        0,
        named_style_box.drawable_object_id,
        body.id(),
        ParagraphStyleName::new("Keynote Draft")?,
    )?;
    let display = editor.rename_slide_text_box_named_paragraph_style(
        0,
        named_style_box.drawable_object_id,
        draft.id(),
        ParagraphStyleName::new("Keynote Display")?,
    )?;
    editor.apply_slide_text_box_named_paragraph_style(
        0,
        named_style_box.drawable_object_id,
        display.id(),
    )?;
    editor.set_slide_text_box_paragraph_alignment(
        0,
        named_style_box.drawable_object_id,
        TextAlignment::Center,
    )?;
    editor.redefine_applied_slide_text_box_named_paragraph_style(
        0,
        named_style_box.drawable_object_id,
    )?;
    let disposable = editor.create_slide_text_box_named_paragraph_style(
        0,
        named_style_box.drawable_object_id,
        body.id(),
        ParagraphStyleName::new("Keynote Disposable")?,
    )?;
    editor.delete_slide_text_box_named_paragraph_style(
        0,
        named_style_box.drawable_object_id,
        disposable.id(),
    )?;
    editor.save(output)?;
    println!(
        "created four-column Keynote text box {} with storage {}",
        created.drawable_object_id, created.storage.object_id
    );
    Ok(())
}
