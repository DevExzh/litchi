//! Set uniform typed underline and strikethrough formatting in an iWork text storage.

use std::env;

use litchi_iwa::text::{IWorkTextEditor, TextDecorations, TextStrikethrough, TextUnderline};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: edit_iwork_text_decorations <input> <output> <storage-id> <underline> <strike>",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let storage_id = arguments.next().ok_or("missing storage ID")?.parse()?;
    let underline = parse_underline(&arguments.next().ok_or("missing underline type")?)?;
    let strikethrough =
        parse_strikethrough(&arguments.next().ok_or("missing strikethrough type")?)?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = IWorkTextEditor::open(input)?;
    editor.set_text_decorations(storage_id, TextDecorations::new(underline, strikethrough))?;
    editor.save(output)?;
    Ok(())
}

fn parse_underline(value: &str) -> Result<TextUnderline, &'static str> {
    match value {
        "none" => Ok(TextUnderline::None),
        "single" => Ok(TextUnderline::Single),
        "double" => Ok(TextUnderline::Double),
        "wavy" => Ok(TextUnderline::Wavy),
        _ => Err("underline must be none, single, double, or wavy"),
    }
}

fn parse_strikethrough(value: &str) -> Result<TextStrikethrough, &'static str> {
    match value {
        "none" => Ok(TextStrikethrough::None),
        "single" => Ok(TextStrikethrough::Single),
        "double" => Ok(TextStrikethrough::Double),
        "triple" => Ok(TextStrikethrough::Triple),
        _ => Err("strikethrough must be none, single, double, or triple"),
    }
}
