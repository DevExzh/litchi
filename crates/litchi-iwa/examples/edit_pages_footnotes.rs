//! Edit the lossless settings shown by Pages' Footnotes formatter.

use std::env;

use litchi_iwa::pages::{
    PagesEditor, PagesFootnoteFormat, PagesFootnoteGap, PagesFootnoteKind, PagesFootnoteNumbering,
    PagesFootnoteSettings,
};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: edit_pages_footnotes <input.pages> <output.pages> \
         <unset|footnotes|document-endnotes|section-endnotes> \
         <unset|numeric|roman|symbolic|japanese-numeric|japanese-ideographic|arabic-numeric> \
         <unset|continuous|page|section> <unset|gap-points>",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let settings = PagesFootnoteSettings {
        kind: parse_kind(arguments.next())?,
        format: parse_format(arguments.next())?,
        numbering: parse_numbering(arguments.next())?,
        gap: parse_gap(arguments.next())?,
    };
    if arguments.next().is_some() {
        return Err("unexpected extra argument".into());
    }

    let mut editor = PagesEditor::open(input)?;
    editor.set_footnote_settings(settings)?;
    editor.save(output)?;
    Ok(())
}

fn parse_kind(value: Option<String>) -> Result<Option<PagesFootnoteKind>, &'static str> {
    match value.as_deref().ok_or("missing footnote kind")? {
        "unset" => Ok(None),
        "footnotes" => Ok(Some(PagesFootnoteKind::Footnotes)),
        "document-endnotes" => Ok(Some(PagesFootnoteKind::DocumentEndnotes)),
        "section-endnotes" => Ok(Some(PagesFootnoteKind::SectionEndnotes)),
        _ => Err("footnote kind must be unset, footnotes, document-endnotes, or section-endnotes"),
    }
}

fn parse_format(value: Option<String>) -> Result<Option<PagesFootnoteFormat>, &'static str> {
    match value.as_deref().ok_or("missing footnote format")? {
        "unset" => Ok(None),
        "numeric" => Ok(Some(PagesFootnoteFormat::Numeric)),
        "roman" => Ok(Some(PagesFootnoteFormat::Roman)),
        "symbolic" => Ok(Some(PagesFootnoteFormat::Symbolic)),
        "japanese-numeric" => Ok(Some(PagesFootnoteFormat::JapaneseNumeric)),
        "japanese-ideographic" => Ok(Some(PagesFootnoteFormat::JapaneseIdeographic)),
        "arabic-numeric" => Ok(Some(PagesFootnoteFormat::ArabicNumeric)),
        _ => Err(
            "footnote format must be unset, numeric, roman, symbolic, japanese-numeric, \
             japanese-ideographic, or arabic-numeric",
        ),
    }
}

fn parse_numbering(value: Option<String>) -> Result<Option<PagesFootnoteNumbering>, &'static str> {
    match value.as_deref().ok_or("missing footnote numbering")? {
        "unset" => Ok(None),
        "continuous" => Ok(Some(PagesFootnoteNumbering::Continuous)),
        "page" => Ok(Some(PagesFootnoteNumbering::RestartEachPage)),
        "section" => Ok(Some(PagesFootnoteNumbering::RestartEachSection)),
        _ => Err("footnote numbering must be unset, continuous, page, or section"),
    }
}

fn parse_gap(
    value: Option<String>,
) -> Result<Option<PagesFootnoteGap>, Box<dyn std::error::Error>> {
    let value = value.ok_or("missing footnote gap")?;
    if value == "unset" {
        return Ok(None);
    }
    Ok(Some(PagesFootnoteGap::new(value.parse::<u32>()?)?))
}
