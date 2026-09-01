//! Read, set, or clear one existing Numbers cell's Currency format.
//!
//! The cell is selected by a sheet selector, a sheet-scoped table selector,
//! and a checked A1 position.  `clear` stages `None`, meaning that the cell
//! has no explicit Currency format.  In contrast, `set automatic ...`
//! stages `Some(Currency)` whose [`DecimalPlaces`] is
//! [`DecimalPlaces::Automatic`].  The command verifies the semantic readback
//! and exact inverse before publishing through a sibling temporary file, so
//! an existing destination is never replaced.  A Currency-format
//! transaction changes display metadata only; the scalar cell value is left
//! untouched.

#![allow(
    clippy::print_stdout,
    reason = "this command-line example reports its committed semantic change"
)]

use std::error::Error;
use std::ffi::{OsStr, OsString};
use std::io;
use std::path::{Path, PathBuf};

use litchi_numbers::cell::data_format::currency::{
    Currency, CurrencyCode, CurrencyStyle, DecimalPlaces, NegativeStyle, ThousandsSeparator,
};
use litchi_numbers::{CellPosition, Package, SheetSelector, TableSelector};
use tempfile::NamedTempFile;

const USAGE: &str = concat!(
    "usage: edit_table_cell_currency_format <input.numbers> <output.numbers> ",
    "<index:N|name:NAME> <index:N|name:NAME> <A1> ",
    "<clear|set CODE DECIMAL_PLACES NEGATIVE_STYLE THOUSANDS STYLE> [--inverse PATH]\n",
    "DECIMAL_PLACES is automatic or 0..30; NEGATIVE_STYLE is ",
    "minus|red|parentheses|red-parentheses; THOUSANDS is hidden|shown; ",
    "STYLE is standard|accounting\n",
    "clear produces None; set automatic ... produces an explicit ",
    "Currency format with DecimalPlaces::Automatic",
);

enum Selector {
    Index(usize),
    Name(String),
}

impl Selector {
    fn sheet(&self) -> SheetSelector<'_> {
        match self {
            Self::Index(index) => SheetSelector::index(*index),
            Self::Name(name) => SheetSelector::name(name),
        }
    }

    fn table(&self) -> TableSelector<'_> {
        match self {
            Self::Index(index) => TableSelector::index(*index),
            Self::Name(name) => TableSelector::name(name),
        }
    }
}

#[derive(Clone, Copy)]
enum Operation {
    Set(Currency),
    Clear,
}

fn main() -> Result<(), Box<dyn Error>> {
    let mut arguments = std::env::args_os().skip(1);
    let input = PathBuf::from(required_argument(&mut arguments, "missing input path")?);
    let output = PathBuf::from(required_argument(&mut arguments, "missing output path")?);
    let sheet = parse_selector(required_text(&mut arguments, "missing sheet selector")?)?;
    let table = parse_selector(required_text(&mut arguments, "missing table selector")?)?;
    let position = CellPosition::from_a1(&required_text(&mut arguments, "missing cell address")?)
        .map_err(|error| invalid_input(format!("invalid A1 address: {error}")))?;
    let operation = parse_operation(&mut arguments)?;
    let inverse_output = parse_inverse_output(&mut arguments)?;

    if input == output {
        return Err(invalid_input("input and output paths must differ"));
    }
    if inverse_output
        .as_deref()
        .is_some_and(|path| path == input || path == output)
    {
        return Err(invalid_input(
            "inverse path must differ from input and output paths",
        ));
    }

    let package = Package::open(&input)?;
    let sheet_selector = sheet.sheet();
    let table_selector = table.table();
    let before = package.table_cell_currency_format(sheet_selector, table_selector, position)?;
    let source_bytes = exact_bytes(&package)?;

    // Read back the existing semantic value before staging any mutation.  A
    // clear/set of that same value must be a byte-exact no-op, which also
    // makes this command useful for checking an admitted source graph.
    let noop = match before {
        Some(format) => package
            .edit_table_cell_currency_format(sheet.sheet(), table.table(), position)?
            .set(format)
            .commit()?,
        None => package
            .edit_table_cell_currency_format(sheet.sheet(), table.table(), position)?
            .clear()
            .commit()?,
    };
    if !noop.patch().is_noop()
        || exact_bytes(noop.package())? != source_bytes
        || noop
            .package()
            .table_cell_currency_format(sheet.sheet(), table.table(), position)?
            != before
    {
        return Err(invalid_input(
            "restaging the existing Currency format was not an exact no-op",
        ));
    }

    let edit = package.edit_table_cell_currency_format(sheet.sheet(), table.table(), position)?;
    let commit = match operation {
        Operation::Set(format) => edit.set(format).commit()?,
        Operation::Clear => edit.clear().commit()?,
    };
    let after =
        commit
            .package()
            .table_cell_currency_format(sheet.sheet(), table.table(), position)?;
    let expected = match operation {
        Operation::Set(format) => Some(format),
        Operation::Clear => None,
    };
    if after != expected {
        return Err(invalid_input(
            "committed package did not expose the requested Currency format",
        ));
    }

    let restored = commit
        .package()
        .apply_table_cell_currency_format(&commit.patch().inverse())?;
    if restored
        .package()
        .table_cell_currency_format(sheet.sheet(), table.table(), position)?
        != before
        || exact_bytes(restored.package())? != source_bytes
    {
        return Err(invalid_input(
            "inverse patch did not restore the exact input package and Currency format",
        ));
    }

    save_new(&output, commit.package())?;
    if let Some(path) = inverse_output {
        save_new(&path, restored.package())?;
    }

    println!(
        "table cell Currency format: cell={}, before={}, after={}, changed={}, touched_components={}, full_reparse={}, scalar_value=untouched",
        position,
        describe_format(before),
        describe_format(after),
        commit.diagnostics().changed(),
        commit.diagnostics().touched_components(),
        commit.diagnostics().full_reparse_performed(),
    );
    Ok(())
}

fn parse_operation(
    arguments: &mut impl Iterator<Item = OsString>,
) -> Result<Operation, Box<dyn Error>> {
    match required_text(arguments, "missing Currency format operation")?.as_str() {
        "clear" => Ok(Operation::Clear),
        "set" => {
            let code = CurrencyCode::new(&required_text(
                arguments,
                "missing three-letter currency code after set",
            )?)
            .map_err(|error| invalid_input(error.to_string()))?;
            let decimal_places = parse_decimal_places(required_text(
                arguments,
                "missing decimal-place setting after set",
            )?)?;
            let negative_style = parse_negative_style(required_text(
                arguments,
                "missing negative-style setting after set",
            )?)?;
            let thousands_separator = parse_thousands_separator(required_text(
                arguments,
                "missing thousands-separator setting after set",
            )?)?;
            let style = parse_currency_style(required_text(
                arguments,
                "missing currency style after set",
            )?)?;
            Ok(Operation::Set(Currency::new(
                code,
                decimal_places,
                negative_style,
                thousands_separator,
                style,
            )))
        },
        _ => Err(invalid_input(
            "Currency format operation must be clear or set",
        )),
    }
}

fn parse_decimal_places(value: String) -> Result<DecimalPlaces, Box<dyn Error>> {
    if value == "automatic" {
        return Ok(DecimalPlaces::Automatic);
    }
    let value = value.parse::<u8>().map_err(|_| {
        invalid_input("decimal places must be automatic or an integer from 0 to 30")
    })?;
    DecimalPlaces::fixed(value).map_err(|error| invalid_input(error.to_string()))
}

fn parse_negative_style(value: String) -> Result<NegativeStyle, Box<dyn Error>> {
    match value.as_str() {
        "minus" => Ok(NegativeStyle::MinusSign),
        "red" => Ok(NegativeStyle::Red),
        "parentheses" => Ok(NegativeStyle::Parentheses),
        "red-parentheses" => Ok(NegativeStyle::RedParentheses),
        _ => Err(invalid_input(
            "negative style must be minus, red, parentheses, or red-parentheses",
        )),
    }
}

fn parse_thousands_separator(value: String) -> Result<ThousandsSeparator, Box<dyn Error>> {
    match value.as_str() {
        "hidden" => Ok(ThousandsSeparator::Hidden),
        "shown" => Ok(ThousandsSeparator::Shown),
        _ => Err(invalid_input("thousands separator must be hidden or shown")),
    }
}

fn parse_currency_style(value: String) -> Result<CurrencyStyle, Box<dyn Error>> {
    match value.as_str() {
        "standard" => Ok(CurrencyStyle::Standard),
        "accounting" => Ok(CurrencyStyle::Accounting),
        _ => Err(invalid_input(
            "currency style must be standard or accounting",
        )),
    }
}

fn parse_selector(value: String) -> Result<Selector, Box<dyn Error>> {
    if let Some(index) = value.strip_prefix("index:") {
        if index.is_empty() || !index.bytes().all(|byte| byte.is_ascii_digit()) {
            return Err(invalid_input(
                "selector index must be a non-negative decimal integer",
            ));
        }
        return index
            .parse()
            .map(Selector::Index)
            .map_err(|_| invalid_input("selector index is too large"));
    }
    if let Some(name) = value.strip_prefix("name:") {
        if name.is_empty() {
            return Err(invalid_input("selector name must not be empty"));
        }
        return Ok(Selector::Name(name.to_owned()));
    }
    Err(invalid_input("selector must start with index: or name:"))
}

fn parse_inverse_output(
    arguments: &mut impl Iterator<Item = OsString>,
) -> Result<Option<PathBuf>, Box<dyn Error>> {
    let Some(flag) = arguments.next() else {
        return Ok(None);
    };
    if flag != OsStr::new("--inverse") {
        return Err(invalid_input(
            "unexpected trailing argument; expected --inverse PATH",
        ));
    }
    let path = PathBuf::from(required_argument(arguments, "missing --inverse path")?);
    if arguments.next().is_some() {
        return Err(invalid_input("unexpected trailing arguments"));
    }
    Ok(Some(path))
}

fn required_argument(
    arguments: &mut impl Iterator<Item = OsString>,
    message: impl Into<String>,
) -> Result<OsString, Box<dyn Error>> {
    arguments
        .next()
        .ok_or_else(|| invalid_input(message.into()))
}

fn required_text(
    arguments: &mut impl Iterator<Item = OsString>,
    message: impl Into<String>,
) -> Result<String, Box<dyn Error>> {
    required_argument(arguments, message)?
        .into_string()
        .map_err(|_| invalid_input("selectors, addresses, and settings must be valid UTF-8"))
}

/// Publishes through a synchronized sibling temporary file without replacing
/// an existing destination.
fn save_new(path: &Path, package: &Package) -> Result<(), Box<dyn Error>> {
    let parent = path
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("."));
    let mut temporary = NamedTempFile::new_in(parent)?;
    package.write_to(temporary.as_file_mut())?;
    temporary.as_file().sync_all()?;
    temporary
        .persist_noclobber(path)
        .map_err(|error| Box::new(error.error))?;
    Ok(())
}

fn exact_bytes(package: &Package) -> Result<Vec<u8>, Box<dyn Error>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn describe_format(format: Option<Currency>) -> String {
    match format {
        None => "none (automatic/no explicit format)".to_owned(),
        Some(format) => format!(
            "explicit(code={},decimal_places={:?},negative_style={:?},thousands={:?},style={:?})",
            format.code(),
            format.decimal_places(),
            format.negative_style(),
            format.thousands_separator(),
            format.style(),
        ),
    }
}

fn invalid_input(message: impl Into<String>) -> Box<dyn Error> {
    Box::new(io::Error::new(
        io::ErrorKind::InvalidInput,
        format!("{}\n\n{USAGE}", message.into()),
    ))
}
