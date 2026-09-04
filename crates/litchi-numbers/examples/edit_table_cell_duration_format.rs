//! Read, set, or reset one existing Numbers cell's Duration format.
//!
//! The cell is selected by a sheet selector, a sheet-scoped table selector,
//! and a checked A1 position.  `clear` and `reset` stage the inherited (no
//! explicit Duration) state; `set` stages a validated presentation style and
//! unit policy.  The command verifies semantic readback and an exact inverse
//! before publishing through sibling temporary files, so an existing
//! destination is never replaced.  A Duration-format transaction changes
//! display metadata only; the scalar cell value is left untouched.
//! Diagnostics report only semantic format settings and bounded transaction
//! counters; native identifiers and archive payloads are never printed.

#![allow(
    clippy::print_stdout,
    reason = "this command-line example reports its committed semantic change"
)]

use std::error::Error;
use std::ffi::{OsStr, OsString};
use std::io;
use std::path::{Path, PathBuf};

use litchi_numbers::cell::data_format::duration::{Duration, Style, Unit, UnitRange, Units};
use litchi_numbers::{CellPosition, Package, SheetSelector, TableSelector};
use tempfile::NamedTempFile;

const USAGE: &str = concat!(
    "usage: edit_table_cell_duration_format <input.numbers> <output.numbers> ",
    "<index:N|name:NAME> <index:N|name:NAME> <A1> ",
    "<clear|reset|set STYLE automatic [LARGEST SMALLEST] ",
    "|set STYLE custom LARGEST SMALLEST> ",
    "[--inverse PATH]\n",
    "STYLE is colon|abbreviated|full-names; LARGEST and SMALLEST are ",
    "weeks|days|hours|minutes|seconds|milliseconds; ",
    "automatic without units uses the complete unit range; when supplied, ",
    "LARGEST and SMALLEST retain that range while Numbers selects visible ",
    "units.  custom displays the inclusive range from LARGEST through ",
    "SMALLEST; clear and reset produce None (automatic/no explicit format)"
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
    Set(Duration),
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
    // Keep the format arguments separate from the optional publication flag;
    // this makes `--inverse` unambiguous even though `set custom` has two
    // additional operands.
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

    let mut operation_arguments = operation_arguments.into_iter();
    let operation = parse_operation(&mut operation_arguments)?;

    // Opening through the focused package owner retains the exact source
    // archive while keeping native IDs and wire records below this example's
    // semantic boundary.
    let package = Package::open(&input)?;
    let sheet_selector = sheet.sheet();
    let table_selector = table.table();
    let before = package.table_cell_duration_format(sheet_selector, table_selector, position)?;
    let source_bytes = exact_bytes(&package)?;

    // Read back the existing semantic value before staging any mutation.  A
    // clear/set of that same value must be a byte-exact no-op, which also
    // makes this command useful for checking an admitted source graph.
    let noop = match before {
        Some(format) => package
            .edit_table_cell_duration_format(sheet.sheet(), table.table(), position)?
            .set(format)
            .commit()?,
        None => package
            .edit_table_cell_duration_format(sheet.sheet(), table.table(), position)?
            .clear()
            .commit()?,
    };
    if !noop.patch().is_noop()
        || noop.diagnostics().changed()
        || exact_bytes(noop.package())? != source_bytes
        || noop
            .package()
            .table_cell_duration_format(sheet.sheet(), table.table(), position)?
            != before
    {
        return Err(invalid_input(
            "restaging the existing Duration format was not an exact no-op",
        ));
    }

    let edit = package.edit_table_cell_duration_format(sheet.sheet(), table.table(), position)?;
    let commit = match operation {
        Operation::Set(format) => edit.set(format).commit()?,
        Operation::Clear => edit.clear().commit()?,
    };
    let after =
        commit
            .package()
            .table_cell_duration_format(sheet.sheet(), table.table(), position)?;
    let expected = match &operation {
        Operation::Set(format) => Some(format),
        Operation::Clear => None,
    };
    if after.as_ref() != expected {
        return Err(invalid_input(
            "committed package did not expose the requested Duration format",
        ));
    }

    // Reopen the candidate bytes before publishing either destination.  The
    // package transaction performs its own verification; this second read is
    // the example's explicit candidate-reopen check.
    let candidate_bytes = exact_bytes(commit.package())?;
    let candidate = Package::from_bytes(&candidate_bytes)?;
    if candidate.table_cell_duration_format(sheet.sheet(), table.table(), position)? != after {
        return Err(invalid_input(
            "reopened candidate did not expose the requested Duration format",
        ));
    }

    let restored = commit
        .package()
        .apply_table_cell_duration_format(&commit.patch().inverse())?;
    if restored
        .package()
        .table_cell_duration_format(sheet.sheet(), table.table(), position)?
        != before
        || exact_bytes(restored.package())? != source_bytes
    {
        return Err(invalid_input(
            "inverse patch did not restore the exact input package and Duration format",
        ));
    }

    let inverse = inverse_output
        .as_deref()
        .map(|path| (path, restored.package()));
    publish_siblings(&output, commit.package(), inverse)?;

    // Reopen the published candidate as well.  This guards the complete
    // filesystem publication path, including the sibling temporary rename.
    let reopened = Package::open(&output)?;
    if reopened.table_cell_duration_format(sheet.sheet(), table.table(), position)? != after {
        return Err(invalid_input(
            "published package did not preserve the requested Duration format",
        ));
    }
    if let Some(path) = inverse_output.as_deref() {
        let reopened_inverse = Package::open(path)?;
        if reopened_inverse.table_cell_duration_format(sheet.sheet(), table.table(), position)?
            != before
            || exact_bytes(&reopened_inverse)? != source_bytes
        {
            return Err(invalid_input(
                "published inverse did not restore the exact input package and Duration format",
            ));
        }
    }

    println!(
        "table cell Duration format: cell={}, before={}, after={}, changed={}, touched_components={}, full_reparse={}, scalar_value=untouched",
        position,
        describe_format(before.as_ref()),
        describe_format(after.as_ref()),
        commit.diagnostics().changed(),
        commit.diagnostics().touched_components(),
        commit.diagnostics().full_reparse_performed(),
    );
    Ok(())
}

fn parse_operation(
    arguments: &mut impl Iterator<Item = OsString>,
) -> Result<Operation, Box<dyn Error>> {
    match required_text(arguments, "missing Duration format operation")?.as_str() {
        "clear" | "reset" => {
            reject_trailing(arguments)?;
            Ok(Operation::Clear)
        },
        "set" => {
            let style = parse_style(required_text(
                arguments,
                "missing Duration style after set",
            )?)?;
            let units = parse_units(arguments)?;
            Ok(Operation::Set(Duration::new(style, units)))
        },
        _ => Err(invalid_input(
            "Duration format operation must be clear, reset, or set",
        )),
    }
}

fn parse_style(value: String) -> Result<Style, Box<dyn Error>> {
    match value.as_str() {
        "colon" => Ok(Style::Colon),
        "abbreviated" => Ok(Style::Abbreviated),
        "full-names" => Ok(Style::FullNames),
        _ => Err(invalid_input(
            "Duration style must be colon, abbreviated, or full-names",
        )),
    }
}

fn parse_units(arguments: &mut impl Iterator<Item = OsString>) -> Result<Units, Box<dyn Error>> {
    match required_text(arguments, "missing Duration unit policy after set")?.as_str() {
        "automatic" => {
            let Some(largest) = arguments.next() else {
                return Ok(Units::Automatic(UnitRange::all()));
            };
            let largest = parse_unit(text_argument(largest, "Duration largest unit")?)?;
            let smallest = parse_unit(required_text(
                arguments,
                "missing smallest Duration unit after automatic",
            )?)?;
            let range = UnitRange::new(largest, smallest)
                .map_err(|error| invalid_input(format!("invalid Duration unit range: {error}")))?;
            reject_trailing(arguments)?;
            Ok(Units::Automatic(range))
        },
        "custom" => {
            let largest = parse_unit(required_text(
                arguments,
                "missing largest Duration unit after custom",
            )?)?;
            let smallest = parse_unit(required_text(
                arguments,
                "missing smallest Duration unit after custom",
            )?)?;
            let range = UnitRange::new(largest, smallest)
                .map_err(|error| invalid_input(format!("invalid Duration unit range: {error}")))?;
            reject_trailing(arguments)?;
            Ok(Units::Custom(range))
        },
        _ => Err(invalid_input(
            "Duration unit policy must be automatic or custom LARGEST SMALLEST",
        )),
    }
}

fn parse_unit(value: String) -> Result<Unit, Box<dyn Error>> {
    match value.as_str() {
        "weeks" => Ok(Unit::Weeks),
        "days" => Ok(Unit::Days),
        "hours" => Ok(Unit::Hours),
        "minutes" => Ok(Unit::Minutes),
        "seconds" => Ok(Unit::Seconds),
        "milliseconds" => Ok(Unit::Milliseconds),
        _ => Err(invalid_input(
            "Duration unit must be weeks, days, hours, minutes, seconds, or milliseconds",
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
            "{field} must be valid UTF-8 (selectors, addresses, operations, styles, and units are text)"
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

fn describe_format(format: Option<&Duration>) -> String {
    match format {
        None => "none (automatic/no explicit format)".to_owned(),
        Some(format) => format!(
            "explicit(style={}, units={})",
            style_name(format.style()),
            describe_units(format.units()),
        ),
    }
}

fn style_name(style: Style) -> &'static str {
    match style {
        Style::Colon => "colon",
        Style::Abbreviated => "abbreviated",
        Style::FullNames => "full-names",
    }
}

fn describe_units(units: Units) -> String {
    let range = units.range();
    let policy = if units.is_automatic() {
        "automatic"
    } else {
        "custom"
    };
    format!(
        "{policy}({}-{})",
        unit_name(range.largest()),
        unit_name(range.smallest()),
    )
}

fn unit_name(unit: Unit) -> &'static str {
    match unit {
        Unit::Weeks => "weeks",
        Unit::Days => "days",
        Unit::Hours => "hours",
        Unit::Minutes => "minutes",
        Unit::Seconds => "seconds",
        Unit::Milliseconds => "milliseconds",
    }
}

fn invalid_input(message: impl Into<String>) -> Box<dyn Error> {
    Box::new(io::Error::new(
        io::ErrorKind::InvalidInput,
        format!("{}\n\n{USAGE}", message.into()),
    ))
}
