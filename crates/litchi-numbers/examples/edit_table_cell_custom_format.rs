//! Read, set, or reset one existing Numbers cell's Custom display format.
//!
//! The cell is selected by a semantic sheet selector, a sheet-scoped table
//! selector, and a checked A1 position.  `set` accepts one of the three
//! explicit Custom families: Number, Text, or Date & Time.  Number rules are
//! supplied as repeated condition/threshold/pattern triples.  `clear` and
//! `reset` stage the inherited (no explicit Custom) state.
//!
//! The command checks an exact no-op, verifies the changed candidate after a
//! fresh in-memory reopen, applies and checks the exact inverse, and only then
//! publishes through synchronized sibling temporary files.  Diagnostics name
//! only the format family and transaction counters; custom names, patterns,
//! affixes, and native identifiers are never printed.

#![allow(
    clippy::print_stdout,
    reason = "this command-line example reports its committed semantic change"
)]

use std::error::Error;
use std::ffi::{OsStr, OsString};
use std::fmt::Display;
use std::io;
use std::path::{Path, PathBuf};

use litchi_numbers::cell::data_format::custom::{
    Condition, ConditionValue, Custom, DateTime as CustomDateTime, DateTimePattern, Name,
    Number as CustomNumber, NumberPattern, NumberRule, Text as CustomText,
};
use litchi_numbers::{CellPosition, Package, SheetSelector, TableSelector};
use tempfile::NamedTempFile;

const USAGE: &str = concat!(
    "usage: edit_table_cell_custom_format <input.numbers> <output.numbers> ",
    "<index:N|name:NAME> <index:N|name:NAME> <A1> ",
    "<clear|reset|set number NAME DEFAULT_PATTERN [CONDITION THRESHOLD PATTERN ...] ",
    "|set text NAME PREFIX SUFFIX |set text-literal NAME LITERAL ",
    "|set datetime NAME PATTERN> [--inverse PATH]\n",
    "CONDITION is eq|lt|le|gt|ge; THRESHOLD is a finite number; ",
    "quote names, patterns, and Text affixes containing spaces.  A Text ",
    "operation places the cell text between PREFIX and SUFFIX; ",
    "text-literal stores LITERAL without the cell text."
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

enum Operation {
    Set(Custom),
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
    let mut operation_arguments = arguments.collect::<Vec<_>>();
    let inverse_output = parse_inverse_output(&mut operation_arguments)?;

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

    let operation = parse_operation(operation_arguments.into_iter())?;
    let package = Package::open(&input)?;
    let sheet_selector = sheet.sheet();
    let table_selector = table.table();
    let before = package.table_cell_custom_format(sheet_selector, table_selector, position)?;
    let source_bytes = exact_bytes(&package)?;

    // Re-staging the value read from the package must be a byte-exact no-op.
    // This catches sources admitted by the reader before any changed request
    // is considered for publication.
    let noop = match before.clone() {
        Some(format) => package
            .edit_table_cell_custom_format(sheet.sheet(), table.table(), position)?
            .set(format)
            .commit()?,
        None => package
            .edit_table_cell_custom_format(sheet.sheet(), table.table(), position)?
            .clear()
            .commit()?,
    };
    if !noop.patch().is_noop()
        || noop.diagnostics().changed()
        || exact_bytes(noop.package())? != source_bytes
        || noop
            .package()
            .table_cell_custom_format(sheet.sheet(), table.table(), position)?
            != before
    {
        return Err(invalid_input(
            "restaging the existing Custom format was not an exact no-op",
        ));
    }

    let expected = match &operation {
        Operation::Set(format) => Some(format.clone()),
        Operation::Clear => None,
    };
    let edit = package.edit_table_cell_custom_format(sheet.sheet(), table.table(), position)?;
    let commit = match operation {
        Operation::Set(format) => edit.set(format).commit()?,
        Operation::Clear => edit.clear().commit()?,
    };
    let after =
        commit
            .package()
            .table_cell_custom_format(sheet.sheet(), table.table(), position)?;
    if after != expected {
        return Err(invalid_input(
            "committed package did not expose the requested Custom format",
        ));
    }

    // Reopen the candidate bytes before publishing either destination.  The
    // package transaction performs its own verification; this second read is
    // the example's explicit candidate-reopen check.
    let candidate_bytes = exact_bytes(commit.package())?;
    let candidate = Package::from_bytes(&candidate_bytes)?;
    if candidate.table_cell_custom_format(sheet.sheet(), table.table(), position)? != after {
        return Err(invalid_input(
            "reopened candidate did not expose the requested Custom format",
        ));
    }

    let restored = commit
        .package()
        .apply_table_cell_custom_format(&commit.patch().inverse())?;
    if restored
        .package()
        .table_cell_custom_format(sheet.sheet(), table.table(), position)?
        != before
        || exact_bytes(restored.package())? != source_bytes
    {
        return Err(invalid_input(
            "inverse patch did not restore the exact input package and Custom format",
        ));
    }

    let inverse = inverse_output
        .as_deref()
        .map(|path| (path, restored.package()));
    publish_siblings(&output, commit.package(), inverse)?;

    // Reopen the published candidate as well.  This guards the complete
    // filesystem publication path, including the sibling temporary rename.
    let reopened = Package::open(&output)?;
    if reopened.table_cell_custom_format(sheet.sheet(), table.table(), position)? != after {
        return Err(invalid_input(
            "published package did not preserve the requested Custom format",
        ));
    }
    if let Some(path) = inverse_output.as_deref() {
        let reopened_inverse = Package::open(path)?;
        if reopened_inverse.table_cell_custom_format(sheet.sheet(), table.table(), position)?
            != before
            || exact_bytes(&reopened_inverse)? != source_bytes
        {
            return Err(invalid_input(
                "published inverse did not restore the exact input package and Custom format",
            ));
        }
    }

    println!(
        "table cell Custom format: cell={position}, before={}, after={}, changed={}, touched_components={}, full_reparse={}",
        describe_format(before.as_ref()),
        describe_format(after.as_ref()),
        commit.diagnostics().changed(),
        commit.diagnostics().touched_components(),
        commit.diagnostics().full_reparse_performed(),
    );
    Ok(())
}

fn parse_operation(arguments: impl Iterator<Item = OsString>) -> Result<Operation, Box<dyn Error>> {
    let mut arguments = arguments;
    match required_text(&mut arguments, "missing Custom format operation")?.as_str() {
        "clear" | "reset" => {
            reject_trailing(&mut arguments)?;
            Ok(Operation::Clear)
        },
        "set" => parse_set_operation(&mut arguments),
        _ => Err(invalid_input(
            "Custom format operation must be clear, reset, or set",
        )),
    }
}

fn parse_set_operation(
    arguments: &mut impl Iterator<Item = OsString>,
) -> Result<Operation, Box<dyn Error>> {
    let family = required_text(arguments, "missing Custom format family after set")?;
    match family.as_str() {
        "number" => parse_custom_number(arguments),
        "text" => parse_custom_text(arguments),
        "text-literal" => parse_custom_literal_text(arguments),
        "datetime" | "date-time" => parse_custom_date_time(arguments),
        _ => Err(invalid_input(
            "Custom format family must be number, text, text-literal, or datetime (date-time is an alias)",
        )),
    }
}

fn parse_custom_number(
    arguments: &mut impl Iterator<Item = OsString>,
) -> Result<Operation, Box<dyn Error>> {
    let name = parse_name(required_text(arguments, "missing Custom Number name")?)?;
    let default_pattern = parse_number_pattern(required_text(
        arguments,
        "missing Custom Number default pattern",
    )?)?;
    let mut rules = Vec::new();
    while let Some(condition) = arguments.next() {
        let condition = text_argument(condition, "Custom Number condition")?;
        let threshold = parse_threshold(required_text(
            arguments,
            "missing Custom Number threshold after condition",
        )?)?;
        let pattern = parse_number_pattern(required_text(
            arguments,
            "missing Custom Number rule pattern after threshold",
        )?)?;
        rules.push(NumberRule::new(
            parse_condition(&condition, threshold)?,
            pattern,
        ));
    }
    let number = CustomNumber::try_with_rules(name, default_pattern, rules)
        .map_err(|error| invalid_input(format!("invalid Custom Number format: {error}")))?;
    Ok(Operation::Set(Custom::Number(number)))
}

fn parse_custom_text(
    arguments: &mut impl Iterator<Item = OsString>,
) -> Result<Operation, Box<dyn Error>> {
    let name = parse_name(required_text(arguments, "missing Custom Text name")?)?;
    let prefix = required_text(arguments, "missing Custom Text prefix")?;
    let suffix = required_text(arguments, "missing Custom Text suffix")?;
    reject_trailing(arguments)?;
    let text = CustomText::try_new(name, prefix, suffix)
        .map_err(|error| invalid_input(format!("invalid Custom Text format: {error}")))?;
    Ok(Operation::Set(Custom::Text(text)))
}

fn parse_custom_literal_text(
    arguments: &mut impl Iterator<Item = OsString>,
) -> Result<Operation, Box<dyn Error>> {
    let name = parse_name(required_text(
        arguments,
        "missing literal Custom Text name",
    )?)?;
    let literal = required_text(arguments, "missing literal Custom Text value")?;
    reject_trailing(arguments)?;
    let text = CustomText::try_literal(name, literal)
        .map_err(|error| invalid_input(format!("invalid literal Custom Text format: {error}")))?;
    Ok(Operation::Set(Custom::Text(text)))
}

fn parse_custom_date_time(
    arguments: &mut impl Iterator<Item = OsString>,
) -> Result<Operation, Box<dyn Error>> {
    let name = parse_name(required_text(arguments, "missing Custom Date & Time name")?)?;
    let pattern = DateTimePattern::try_new(required_text(
        arguments,
        "missing Custom Date & Time pattern",
    )?)
    .map_err(|error| invalid_input(format!("invalid Custom Date & Time format: {error}")))?;
    reject_trailing(arguments)?;
    Ok(Operation::Set(Custom::DateTime(CustomDateTime::new(
        name, pattern,
    ))))
}

fn parse_name(value: String) -> Result<Name, Box<dyn Error>> {
    Name::try_new(value)
        .map_err(|error| invalid_input(format!("invalid Custom format name: {error}")))
}

fn parse_number_pattern(value: String) -> Result<NumberPattern, Box<dyn Error>> {
    NumberPattern::try_new(value)
        .map_err(|error| invalid_input(format!("invalid Custom Number pattern: {error}")))
}

fn parse_threshold(value: String) -> Result<ConditionValue, Box<dyn Error>> {
    let threshold = value
        .parse::<f64>()
        .map_err(|_| invalid_input("Custom Number threshold must be a finite number"))?;
    ConditionValue::try_new(threshold)
        .map_err(|error| invalid_input(format!("invalid Custom Number threshold: {error}")))
}

fn parse_condition(value: &str, threshold: ConditionValue) -> Result<Condition, Box<dyn Error>> {
    match value {
        "eq" => Ok(Condition::EqualTo(threshold)),
        "lt" => Ok(Condition::LessThan(threshold)),
        "le" => Ok(Condition::LessThanOrEqualTo(threshold)),
        "gt" => Ok(Condition::GreaterThan(threshold)),
        "ge" => Ok(Condition::GreaterThanOrEqualTo(threshold)),
        _ => Err(invalid_input(
            "Custom Number condition must be eq, lt, le, gt, or ge",
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

fn parse_inverse_output(arguments: &mut Vec<OsString>) -> Result<Option<PathBuf>, Box<dyn Error>> {
    let Some(index) = arguments
        .iter()
        .position(|argument| argument == OsStr::new("--inverse"))
    else {
        return Ok(None);
    };
    if index + 2 != arguments.len() {
        return Err(invalid_input(
            "--inverse PATH must be the final command-line option",
        ));
    }
    let path = PathBuf::from(arguments[index + 1].clone());
    arguments.truncate(index);
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
    text_argument(
        required_argument(arguments, message)?,
        "command-line argument",
    )
}

fn text_argument(value: OsString, field: &str) -> Result<String, Box<dyn Error>> {
    value.into_string().map_err(|_| {
        invalid_input(format!(
            "{field} must be valid UTF-8 (selectors, addresses, operations, and formats are text)"
        ))
    })
}

fn reject_trailing(arguments: &mut impl Iterator<Item = OsString>) -> Result<(), Box<dyn Error>> {
    if arguments.next().is_some() {
        return Err(invalid_input("unexpected trailing arguments"));
    }
    Ok(())
}

/// Stage complete package bytes in synchronized temporary files next to each
/// destination, then publish each with a no-clobber atomic rename. Both
/// artifacts are fully staged before either destination is touched.
fn publish_siblings(
    output: &Path,
    package: &Package,
    inverse: Option<(&Path, &Package)>,
) -> Result<(), Box<dyn Error>> {
    ensure_new_destination(output)?;
    if let Some((path, _)) = inverse {
        ensure_new_destination(path)?;
    }

    let output_temporary = stage_sibling(output, package)?;
    let inverse_temporary = inverse
        .map(|(path, package)| stage_sibling(path, package))
        .transpose()?;

    output_temporary
        .persist_noclobber(output)
        .map_err(|error| Box::new(error.error))?;
    if let (Some((path, _)), Some(temporary)) = (inverse, inverse_temporary) {
        temporary
            .persist_noclobber(path)
            .map_err(|error| Box::new(error.error))?;
    }
    Ok(())
}

fn ensure_new_destination(path: &Path) -> Result<(), Box<dyn Error>> {
    if path.try_exists()? {
        return Err(invalid_input(
            "publication destination already exists; choose a new output path",
        ));
    }
    Ok(())
}

fn stage_sibling(path: &Path, package: &Package) -> Result<NamedTempFile, Box<dyn Error>> {
    let parent = path
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("."));
    let mut temporary = NamedTempFile::new_in(parent)?;
    package.write_to(temporary.as_file_mut())?;
    temporary.as_file().sync_all()?;
    Ok(temporary)
}

fn exact_bytes(package: &Package) -> Result<Vec<u8>, Box<dyn Error>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn describe_format(format: Option<&Custom>) -> &'static str {
    match format {
        None => "none (automatic/no explicit Custom format)",
        Some(Custom::Number(_)) => "explicit Custom Number",
        Some(Custom::Text(_)) => "explicit Custom Text",
        Some(Custom::DateTime(_)) => "explicit Custom Date & Time",
    }
}

fn invalid_input(message: impl Display) -> Box<dyn Error> {
    Box::new(io::Error::new(
        io::ErrorKind::InvalidInput,
        format!("{message}\n\n{USAGE}"),
    ))
}
