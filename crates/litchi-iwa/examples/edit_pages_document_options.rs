//! Edit the lossless options shown by Pages' Document formatter.

use std::env;
use std::path::PathBuf;

use litchi_iwa::pages::PagesEditor;
use litchi_pages::document_options::Options;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = PathBuf::from(arguments.next().ok_or(
        "usage: edit_pages_document_options <input.pages> <output.pages> <unset|true|false:body> <unset|true|false:headers> <unset|true|false:footers> <unset|true|false:facing-pages> <unset|true|false:hyphenation> <unset|true|false:ligatures>",
    )?);
    let output = PathBuf::from(arguments.next().ok_or("missing output path")?);
    let options = Options::new(
        parse_optional_bool(arguments.next(), "body")?,
        parse_optional_bool(arguments.next(), "headers")?,
        parse_optional_bool(arguments.next(), "footers")?,
        parse_optional_bool(arguments.next(), "facing pages")?,
        parse_optional_bool(arguments.next(), "hyphenation")?,
        parse_optional_bool(arguments.next(), "ligatures")?,
    );
    if arguments.next().is_some() {
        return Err("unexpected extra argument".into());
    }

    let mut editor = PagesEditor::open(input)?;
    editor.set_document_options(options)?;
    editor.save(output)?;
    Ok(())
}

fn parse_optional_bool(
    value: Option<String>,
    label: &str,
) -> Result<Option<bool>, Box<dyn std::error::Error>> {
    match value.ok_or_else(|| format!("missing {label}"))?.as_str() {
        "unset" => Ok(None),
        "true" => Ok(Some(true)),
        "false" => Ok(Some(false)),
        _ => Err(format!("{label} must be unset, true, or false").into()),
    }
}
