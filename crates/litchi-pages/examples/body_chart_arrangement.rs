//! Edit one Pages body chart's Arrange-panel state through the focused API.
//!
//! The chart is selected by its zero-based semantic body position. Native
//! object identifiers never enter this command-line interface or its output.
//!
//! The example proves an exact no-op, rejects reusing the forward patch on its
//! committed target, and applies the inverse in memory before publishing the
//! requested output. Existing output paths are never overwritten.

#![allow(
    clippy::print_stdout,
    reason = "this command-line example reports its committed semantic change"
)]

use std::error::Error;
use std::ffi::{OsStr, OsString};
use std::io;
use std::path::{Path, PathBuf};

use litchi_pages::{BodyChartArrangementError, BodyChartSelector, ChartArrangement, Package};
use tempfile::NamedTempFile;

const USAGE: &str = "usage: body_chart_arrangement <input.pages> <output.pages> \
                     <locked> <constrain-proportions> <index:N> \
                     [--inverse PATH] \
                     (boolean values: true or false)";

fn main() -> Result<(), Box<dyn Error>> {
    let mut arguments = std::env::args_os().skip(1);
    let input = PathBuf::from(required_argument(&mut arguments, "missing input path")?);
    let output = PathBuf::from(required_argument(&mut arguments, "missing output path")?);
    let locked = parse_bool(
        required_argument(&mut arguments, "missing locked state")?,
        "locked state",
    )?;
    let constrain_proportions = parse_bool(
        required_argument(&mut arguments, "missing constrain-proportions state")?,
        "constrain-proportions state",
    )?;
    let chart = parse_chart(required_argument(
        &mut arguments,
        "missing body-chart selector",
    )?)?;
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
    package.validate()?;
    let before = package.body_chart_arrangement(chart.selector())?;
    let source_bytes = exact_bytes(&package)?;
    let requested = ChartArrangement::new(locked, constrain_proportions);

    // A requested value equal to the source value must not rewrite the
    // component or change any package bytes.
    let noop = package
        .edit_body_chart_arrangement(chart.selector())?
        .set(before)
        .commit()?;
    if noop.diagnostics().changed()
        || !noop.patch().is_noop()
        || exact_bytes(noop.package())? != source_bytes
        || noop.package().body_chart_arrangement(chart.selector())? != before
    {
        return Err(invalid_input(
            "setting the existing body-chart arrangement was not an exact no-op",
        ));
    }

    let commit = package
        .edit_body_chart_arrangement(chart.selector())?
        .set(requested)
        .commit()?;
    let committed = commit.package().body_chart_arrangement(chart.selector())?;
    if committed != requested {
        return Err(invalid_input(
            "committed package did not expose the requested body-chart arrangement",
        ));
    }

    let source_conflict_checked = !commit.patch().is_noop();
    if source_conflict_checked
        && !matches!(
            commit
                .package()
                .apply_body_chart_arrangement(commit.patch()),
            Err(BodyChartArrangementError::PatchConflict)
        )
    {
        return Err(invalid_input(
            "forward body-chart arrangement patch was not rejected on its committed target",
        ));
    }

    // Apply the exact inverse in memory and require semantic and byte
    // equality with the source package before publishing any output.
    let restored = commit
        .package()
        .apply_body_chart_arrangement(&commit.patch().inverse())?;
    if restored
        .package()
        .body_chart_arrangement(chart.selector())?
        != before
        || exact_bytes(restored.package())? != source_bytes
    {
        return Err(invalid_input(
            "inverse patch did not restore the exact input package and body-chart arrangement",
        ));
    }

    save_new(&output, commit.package())?;
    if let Some(path) = inverse_output {
        save_new(&path, restored.package())?;
    }
    let reopened = Package::open(&output)?;
    reopened.validate()?;
    let after = reopened.body_chart_arrangement(chart.selector())?;
    if after != requested {
        return Err(invalid_input(
            "reopened output did not preserve the requested body-chart arrangement",
        ));
    }

    println!(
        "body chart arrangement: locked_before={}, constrain_proportions_before={}, locked_after={}, constrain_proportions_after={}, changed={}, touched_components={}, full_reparse={}, source_conflict_checked={}, source_fingerprint={:016x}, target_fingerprint={:016x}",
        before.locked(),
        before.constrain_proportions(),
        after.locked(),
        after.constrain_proportions(),
        commit.diagnostics().changed(),
        commit.diagnostics().touched_components(),
        commit.diagnostics().full_reparse_performed(),
        source_conflict_checked,
        commit.patch().source_fingerprint(),
        commit.patch().target_fingerprint(),
    );
    Ok(())
}

struct SelectedChart(usize);

impl SelectedChart {
    fn selector(&self) -> BodyChartSelector {
        BodyChartSelector::index(self.0)
    }
}

fn parse_chart(argument: OsString) -> Result<SelectedChart, Box<dyn Error>> {
    let value = argument
        .into_string()
        .map_err(|_| invalid_input("body-chart selector must be valid UTF-8"))?;
    let index = value
        .strip_prefix("index:")
        .ok_or_else(|| invalid_input("body-chart selector must start with index:"))?;
    if index.is_empty() || !index.bytes().all(|byte| byte.is_ascii_digit()) {
        return Err(invalid_input(
            "body-chart index must be a non-negative decimal integer",
        ));
    }
    index
        .parse::<usize>()
        .map(SelectedChart)
        .map_err(|_| invalid_input("body-chart index is too large for this platform"))
}

fn parse_bool(argument: OsString, kind: &str) -> Result<bool, Box<dyn Error>> {
    let value = argument
        .into_string()
        .map_err(|_| invalid_input(format!("{kind} must be valid UTF-8 (true or false)")))?;
    match value.as_str() {
        "true" => Ok(true),
        "false" => Ok(false),
        _ => Err(invalid_input(format!(
            "{kind} must be exactly true or false"
        ))),
    }
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
