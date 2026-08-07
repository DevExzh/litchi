//! Edit typed Numbers table header, footer, freeze, and repetition settings.
use litchi_numbers::table::headers::{Count as HeaderCount, Settings as HeaderSettings};

use std::env;
use std::path::PathBuf;

use litchi_iwa::numbers::NumbersEditor;
use litchi_numbers::TableSelector;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = PathBuf::from(arguments.next().ok_or(
        "usage: edit_numbers_table_headers <input.numbers> <output.numbers> <table-index> <header-rows:0..5> <header-columns:0..5> <footer-rows:0..5> <unset|true|false:freeze-rows> <unset|true|false:freeze-columns> <unset|true|false:repeat-rows> <unset|true|false:repeat-columns>",
    )?);
    let output = PathBuf::from(arguments.next().ok_or("missing output path")?);
    let table_index = arguments
        .next()
        .ok_or("missing table index")?
        .parse::<usize>()?;
    let settings = HeaderSettings {
        header_rows: parse_count(arguments.next(), "header rows")?,
        header_columns: parse_count(arguments.next(), "header columns")?,
        footer_rows: parse_count(arguments.next(), "footer rows")?,
        header_rows_frozen: parse_optional_bool(arguments.next(), "freeze rows")?,
        header_columns_frozen: parse_optional_bool(arguments.next(), "freeze columns")?,
        repeating_header_rows_enabled: parse_optional_bool(arguments.next(), "repeat rows")?,
        repeating_header_columns_enabled: parse_optional_bool(arguments.next(), "repeat columns")?,
    };
    if arguments.next().is_some() {
        return Err("unexpected extra argument".into());
    }

    let mut editor = NumbersEditor::open(input)?;
    editor.set_table_header_settings(TableSelector::index(table_index), settings)?;
    editor.save(output)?;
    Ok(())
}

fn parse_count(
    value: Option<String>,
    label: &str,
) -> Result<Option<HeaderCount>, Box<dyn std::error::Error>> {
    let count = value
        .ok_or_else(|| format!("missing {label}"))?
        .parse::<usize>()?;
    if count == 0 {
        Ok(None)
    } else {
        Ok(Some(HeaderCount::new(count)?))
    }
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
