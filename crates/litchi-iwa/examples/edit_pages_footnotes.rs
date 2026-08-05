//! Edit the lossless settings shown by Pages' Footnotes formatter.

use std::env;

use litchi_iwa::pages::PagesEditor;
use litchi_pages::footnote::{Format, Gap, Kind, Numbering, Settings};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: edit_pages_footnotes <input.pages> <output.pages> \
         <unset|footnotes|document-endnotes|section-endnotes> \
         <unset|numeric|roman|symbolic|japanese-numeric|japanese-ideographic|arabic-numeric> \
         <unset|continuous|page|section> <unset|gap-points>",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let settings = Settings {
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

fn parse_kind(value: Option<String>) -> Result<Option<Kind>, &'static str> {
    match value.as_deref().ok_or("missing footnote kind")? {
        "unset" => Ok(None),
        "footnotes" => Ok(Some(Kind::Footnotes)),
        "document-endnotes" => Ok(Some(Kind::DocumentEndnotes)),
        "section-endnotes" => Ok(Some(Kind::SectionEndnotes)),
        _ => Err("footnote kind must be unset, footnotes, document-endnotes, or section-endnotes"),
    }
}

fn parse_format(value: Option<String>) -> Result<Option<Format>, &'static str> {
    match value.as_deref().ok_or("missing footnote format")? {
        "unset" => Ok(None),
        "numeric" => Ok(Some(Format::Numeric)),
        "roman" => Ok(Some(Format::Roman)),
        "symbolic" => Ok(Some(Format::Symbolic)),
        "japanese-numeric" => Ok(Some(Format::JapaneseNumeric)),
        "japanese-ideographic" => Ok(Some(Format::JapaneseIdeographic)),
        "arabic-numeric" => Ok(Some(Format::ArabicNumeric)),
        _ => Err(
            "footnote format must be unset, numeric, roman, symbolic, japanese-numeric, \
             japanese-ideographic, or arabic-numeric",
        ),
    }
}

fn parse_numbering(value: Option<String>) -> Result<Option<Numbering>, &'static str> {
    match value.as_deref().ok_or("missing footnote numbering")? {
        "unset" => Ok(None),
        "continuous" => Ok(Some(Numbering::Continuous)),
        "page" => Ok(Some(Numbering::RestartEachPage)),
        "section" => Ok(Some(Numbering::RestartEachSection)),
        _ => Err("footnote numbering must be unset, continuous, page, or section"),
    }
}

fn parse_gap(value: Option<String>) -> Result<Option<Gap>, Box<dyn std::error::Error>> {
    let value = value.ok_or("missing footnote gap")?;
    if value == "unset" {
        return Ok(None);
    }
    Ok(Some(Gap::new(value.parse::<u32>()?)?))
}
