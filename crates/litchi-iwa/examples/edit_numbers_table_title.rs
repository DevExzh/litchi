//! Edit typed Numbers table title visibility and outline settings.

use std::env;
use std::path::PathBuf;

use litchi_iwa::numbers::{NumbersEditor, Settings};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = PathBuf::from(arguments.next().ok_or(
        "usage: edit_numbers_table_title <input.numbers> <output.numbers> <table-id> <unset|true|false:visible> <unset|true|false:outlined>",
    )?);
    let output = PathBuf::from(arguments.next().ok_or("missing output path")?);
    let table_id = arguments.next().ok_or("missing table ID")?.parse::<u64>()?;
    let settings = Settings::new(
        parse_optional_bool(arguments.next(), "title visibility")?,
        parse_optional_bool(arguments.next(), "title outline")?,
    );
    if arguments.next().is_some() {
        return Err("unexpected extra argument".into());
    }

    let mut editor = NumbersEditor::open(input)?;
    editor.set_table_title_settings(table_id, settings)?;
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
