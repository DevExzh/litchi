//! Edit one numeric cell in a Keynote slide chart through the focused API.
//!
//! The command changes exactly one value in the selected chart. Row labels,
//! column labels, and the rectangular shape are copied from the source grid
//! and therefore remain unchanged. Use `empty` to clear a cell.

#![allow(
    clippy::print_stdout,
    reason = "this command-line example reports its committed semantic change"
)]

use std::error::Error;
use std::ffi::OsString;
use std::io;
use std::path::{Path, PathBuf};

use litchi_keynote::{ChartData, ChartSelector, Package, SlideChartDataError, SlideSelector};
use tempfile::NamedTempFile;

const USAGE: &str = "usage: edit_slide_chart_data <input.key> <output.key> \
                     [index:N|name:SLIDE] [index:N|name:CHART] \
                     <row-index> <column-index> <value|empty>";

fn main() -> Result<(), Box<dyn Error>> {
    let mut arguments = std::env::args_os().skip(1);
    let input = PathBuf::from(required_argument(
        &mut arguments,
        "missing input Keynote path",
    )?);
    let output = PathBuf::from(required_argument(
        &mut arguments,
        "missing output Keynote path",
    )?);
    let slide = parse_slide(required_argument(&mut arguments, "missing slide selector")?)?;
    let chart = parse_chart(required_argument(&mut arguments, "missing chart selector")?)?;
    let row = parse_index(
        required_argument(&mut arguments, "missing row index")?,
        "row index",
    )?;
    let column = parse_index(
        required_argument(&mut arguments, "missing column index")?,
        "column index",
    )?;
    let value = parse_value(required_argument(&mut arguments, "missing cell value")?)?;
    if arguments.next().is_some() {
        return Err(invalid_input("unexpected trailing arguments"));
    }
    if input == output {
        return Err(invalid_input("input and output paths must differ"));
    }

    let package = Package::open(&input)?;
    package.validate()?;
    let before = package.slide_chart_data(slide.selector(), chart.selector())?;
    let requested = replace_cell(&before, row, column, value)?;
    let source_bytes = exact_bytes(&package)?;

    // Exercise the no-op path first. It must preserve the complete package
    // byte-for-byte, including the selected chart's untouched labels/shapes.
    let noop = package
        .edit_slide_chart_data(slide.selector(), chart.selector())?
        .set(before.clone())
        .commit()?;
    if noop.diagnostics().changed()
        || !noop.patch().is_noop()
        || exact_bytes(noop.package())? != source_bytes
        || noop
            .package()
            .slide_chart_data(slide.selector(), chart.selector())?
            != before
    {
        return Err(invalid_input(
            "setting the existing slide-chart data was not an exact no-op",
        ));
    }

    let commit = package
        .edit_slide_chart_data(slide.selector(), chart.selector())?
        .set(requested.clone())
        .commit()?;
    let committed = commit
        .package()
        .slide_chart_data(slide.selector(), chart.selector())?;
    assert_same_axes_and_shape(&before, &committed)?;
    if committed != requested {
        return Err(invalid_input(
            "committed package did not expose the requested slide-chart value",
        ));
    }

    let source_conflict_checked = !commit.patch().is_noop();
    if source_conflict_checked
        && !matches!(
            commit.package().apply_slide_chart_data(commit.patch()),
            Err(SlideChartDataError::PatchConflict)
        )
    {
        return Err(invalid_input(
            "forward slide-chart data patch was not rejected on its committed target",
        ));
    }

    // Apply the exact inverse in memory and require both semantic and byte
    // equality with the source package before publishing any output.
    let restored = commit
        .package()
        .apply_slide_chart_data(&commit.patch().inverse())?;
    if restored
        .package()
        .slide_chart_data(slide.selector(), chart.selector())?
        != before
        || exact_bytes(restored.package())? != source_bytes
    {
        return Err(invalid_input(
            "inverse patch did not restore the exact input slide-chart data",
        ));
    }

    save_new(&output, commit.package())?;
    let reopened = Package::open(&output)?;
    reopened.validate()?;
    let after = reopened.slide_chart_data(slide.selector(), chart.selector())?;
    assert_same_axes_and_shape(&before, &after)?;
    if after != requested {
        return Err(invalid_input(
            "reopened output did not preserve the requested slide-chart value",
        ));
    }

    println!(
        "slide chart data: slide={}, chart={}, row={}, column={}, before={}, after={}, changed={}, touched_components={}, full_reparse={}, source_conflict_checked={}, source_fingerprint={:016x}, target_fingerprint={:016x}",
        slide.display(),
        chart.display(),
        row,
        column,
        format_value(before.values()[row][column]),
        format_value(after.values()[row][column]),
        commit.diagnostics().changed(),
        commit.diagnostics().touched_components(),
        commit.diagnostics().full_reparse_performed(),
        source_conflict_checked,
        commit.patch().source_fingerprint(),
        commit.patch().target_fingerprint(),
    );
    Ok(())
}

enum SelectedSlide {
    Index(usize),
    Name(String),
}

impl SelectedSlide {
    fn selector(&self) -> SlideSelector<'_> {
        match self {
            Self::Index(index) => SlideSelector::index(*index),
            Self::Name(name) => SlideSelector::name(name),
        }
    }

    fn display(&self) -> String {
        match self {
            Self::Index(index) => format!("index:{index}"),
            Self::Name(name) => format!("name:{name}"),
        }
    }
}

enum SelectedChart {
    Index(usize),
    Name(String),
}

impl SelectedChart {
    fn selector(&self) -> ChartSelector<'_> {
        match self {
            Self::Index(index) => ChartSelector::index(*index),
            Self::Name(name) => ChartSelector::name(name),
        }
    }

    fn display(&self) -> String {
        match self {
            Self::Index(index) => format!("index:{index}"),
            Self::Name(name) => format!("name:{name}"),
        }
    }
}

fn parse_slide(argument: OsString) -> Result<SelectedSlide, Box<dyn Error>> {
    parse_selector(argument, "slide", SelectedSlide::Index, SelectedSlide::Name)
}

fn parse_chart(argument: OsString) -> Result<SelectedChart, Box<dyn Error>> {
    parse_selector(argument, "chart", SelectedChart::Index, SelectedChart::Name)
}

fn parse_selector<T>(
    argument: OsString,
    kind: &str,
    from_index: impl FnOnce(usize) -> T,
    from_name: impl FnOnce(String) -> T,
) -> Result<T, Box<dyn Error>> {
    let value = argument
        .into_string()
        .map_err(|_| invalid_input(format!("{kind} selector must be valid UTF-8")))?;
    if let Some(index) = value.strip_prefix("index:") {
        return Ok(from_index(parse_decimal(index, &format!("{kind} index"))?));
    }
    if let Some(name) = value.strip_prefix("name:") {
        if name.is_empty() {
            return Err(invalid_input(format!("{kind} name cannot be empty")));
        }
        return Ok(from_name(name.to_owned()));
    }
    Err(invalid_input(format!(
        "{kind} selector must start with index: or name:"
    )))
}

fn parse_index(argument: OsString, kind: &str) -> Result<usize, Box<dyn Error>> {
    let value = argument
        .into_string()
        .map_err(|_| invalid_input(format!("{kind} must be valid UTF-8")))?;
    parse_decimal(&value, kind)
}

fn parse_decimal(value: &str, kind: &str) -> Result<usize, Box<dyn Error>> {
    if value.is_empty() || !value.bytes().all(|byte| byte.is_ascii_digit()) {
        return Err(invalid_input(format!(
            "{kind} must be a non-negative decimal integer"
        )));
    }
    value
        .parse::<usize>()
        .map_err(|_| invalid_input(format!("{kind} is too large for this platform")))
}

fn parse_value(argument: OsString) -> Result<Option<f64>, Box<dyn Error>> {
    let value = argument
        .into_string()
        .map_err(|_| invalid_input("cell value must be valid UTF-8"))?;
    if value == "empty" {
        return Ok(None);
    }
    let parsed = value
        .parse::<f64>()
        .map_err(|_| invalid_input("cell value must be a finite number or empty"))?;
    if !parsed.is_finite() {
        return Err(invalid_input("cell value must be a finite number or empty"));
    }
    Ok(Some(parsed))
}

fn replace_cell(
    source: &ChartData,
    row: usize,
    column: usize,
    value: Option<f64>,
) -> Result<ChartData, Box<dyn Error>> {
    let mut values = source.values().to_owned();
    let row_count = values.len();
    let row_values = values.get_mut(row).ok_or_else(|| {
        invalid_input(format!(
            "row index {row} is outside the {}-row chart",
            row_count
        ))
    })?;
    let column_count = source.column_names().len();
    let cell = row_values.get_mut(column).ok_or_else(|| {
        invalid_input(format!(
            "column index {column} is outside the {}-column chart",
            column_count
        ))
    })?;
    *cell = value;
    ChartData::new(
        source.row_names().to_owned(),
        source.column_names().to_owned(),
        values,
    )
    .map_err(|error| invalid_input(format!("could not construct chart data: {error}")))
}

fn assert_same_axes_and_shape(before: &ChartData, after: &ChartData) -> Result<(), Box<dyn Error>> {
    if before.row_names() != after.row_names()
        || before.column_names() != after.column_names()
        || before.values().len() != after.values().len()
        || before
            .values()
            .iter()
            .zip(after.values())
            .any(|(before, after)| before.len() != after.len())
    {
        return Err(invalid_input(
            "chart data edit changed labels or the rectangular shape",
        ));
    }
    Ok(())
}

fn format_value(value: Option<f64>) -> String {
    match value {
        Some(value) => value.to_string(),
        None => "empty".to_owned(),
    }
}

fn required_argument(
    arguments: &mut impl Iterator<Item = OsString>,
    message: &'static str,
) -> Result<OsString, Box<dyn Error>> {
    arguments.next().ok_or_else(|| invalid_input(message))
}

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
        .map_err(|error| -> Box<dyn Error> { Box::new(error.error) })?;
    Ok(())
}

fn exact_bytes(package: &Package) -> Result<Vec<u8>, Box<dyn Error>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn invalid_input(message: impl Into<String>) -> Box<dyn Error> {
    Box::new(io::Error::new(
        io::ErrorKind::InvalidInput,
        format!("{}\n\n{USAGE}", message.into()),
    ))
}
