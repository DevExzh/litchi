//! Set or clear one chart title through the focused Keynote APIs.
//!
//! The target can be the chart title, category-axis title, or value-axis
//! title. Selection stays semantic (`index:N` or `name:NAME`); native object
//! identifiers never enter this command-line interface or its diagnostics.

#![allow(
    clippy::print_stdout,
    reason = "this command-line example reports its committed semantic change"
)]

use std::error::Error;
use std::ffi::OsString;
use std::io;
use std::path::{Path, PathBuf};

use litchi_keynote::{Axis, ChartSelector, Package, SlideSelector};
use tempfile::NamedTempFile;

const USAGE: &str = "usage: edit_chart_title <input.key> <output.key> \
                     <chart|category-axis|value-axis> <set:TITLE|clear> \
                     [index:N|name:SLIDE] [index:N|name:CHART]";

#[derive(Clone, Copy)]
enum TitleTarget {
    Chart,
    CategoryAxis,
    ValueAxis,
}

impl TitleTarget {
    const fn as_str(self) -> &'static str {
        match self {
            Self::Chart => "chart",
            Self::CategoryAxis => "category-axis",
            Self::ValueAxis => "value-axis",
        }
    }
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
}

enum TitleChange {
    Set(String),
    Clear,
}

impl TitleChange {
    fn expected(&self) -> Option<&str> {
        match self {
            Self::Set(title) => Some(title),
            Self::Clear => None,
        }
    }
}

struct Report {
    before_present: bool,
    after_present: bool,
    after_bytes: usize,
    changed: bool,
    touched_components: usize,
    deleted_previews: usize,
    full_reparse: bool,
    source_fingerprint: u64,
    target_fingerprint: u64,
}

fn main() -> Result<(), Box<dyn Error>> {
    let mut arguments = std::env::args_os().skip(1);
    let input = PathBuf::from(required_argument(&mut arguments, "missing input path")?);
    let output = PathBuf::from(required_argument(&mut arguments, "missing output path")?);
    let target = parse_target(required_argument(&mut arguments, "missing title target")?)?;
    let change = parse_change(required_argument(&mut arguments, "missing title change")?)?;
    let slide = parse_slide(arguments.next())?;
    let chart = parse_chart(arguments.next())?;
    if arguments.next().is_some() {
        return Err(invalid_input("unexpected trailing arguments"));
    }
    if input == output {
        return Err(invalid_input("input and output paths must differ"));
    }

    let package = Package::open(&input)
        .map_err(|error| io::Error::other(format!("opening input package failed: {error}")))?;
    package
        .validate()
        .map_err(|error| io::Error::other(format!("validating input package failed: {error}")))?;
    let report = match target {
        TitleTarget::Chart => edit_chart_title(&package, &output, &slide, &chart, &change)?,
        TitleTarget::CategoryAxis => {
            edit_axis_title(&package, &output, &slide, &chart, Axis::Category, &change)?
        },
        TitleTarget::ValueAxis => {
            edit_axis_title(&package, &output, &slide, &chart, Axis::Value, &change)?
        },
    };

    println!(
        "chart title: target={}, before_present={}, after_present={}, after_bytes={}, changed={}, touched_components={}, deleted_previews={}, full_reparse={}, source_fingerprint={:016x}, target_fingerprint={:016x}",
        target.as_str(),
        report.before_present,
        report.after_present,
        report.after_bytes,
        report.changed,
        report.touched_components,
        report.deleted_previews,
        report.full_reparse,
        report.source_fingerprint,
        report.target_fingerprint,
    );
    Ok(())
}

fn edit_chart_title(
    package: &Package,
    output: &Path,
    slide: &SelectedSlide,
    chart: &SelectedChart,
    change: &TitleChange,
) -> Result<Report, Box<dyn Error>> {
    let before = package
        .slide_chart_title(slide.selector(), chart.selector())
        .map_err(|error| io::Error::other(format!("reading chart title failed: {error}")))?;
    let source_bytes = exact_bytes(package)
        .map_err(|error| io::Error::other(format!("serializing input package failed: {error}")))?;

    let noop_edit = package
        .edit_slide_chart_title(slide.selector(), chart.selector())
        .map_err(|error| {
            io::Error::other(format!("chart-title no-op preflight failed: {error}"))
        })?;
    let noop = match before.as_deref() {
        Some(title) => noop_edit.set(title).map_err(|error| {
            io::Error::other(format!("staging chart-title no-op set failed: {error}"))
        })?,
        None => noop_edit.clear().map_err(|error| {
            io::Error::other(format!("staging chart-title no-op clear failed: {error}"))
        })?,
    };
    let noop = noop
        .commit()
        .map_err(|error| io::Error::other(format!("chart-title no-op commit failed: {error}")))?;
    let noop_title = noop
        .package()
        .slide_chart_title(slide.selector(), chart.selector())
        .map_err(|error| {
            io::Error::other(format!("reading chart-title no-op result failed: {error}"))
        })?;
    let noop_bytes = exact_bytes(noop.package()).map_err(|error| {
        io::Error::other(format!(
            "serializing chart-title no-op result failed: {error}"
        ))
    })?;
    if noop.diagnostics().changed()
        || !noop.patch().is_noop()
        || noop_title != before
        || noop_bytes != source_bytes
    {
        return Err(invalid_input(
            "setting the existing chart title was not an exact no-op",
        ));
    }

    let edit = package
        .edit_slide_chart_title(slide.selector(), chart.selector())
        .map_err(|error| io::Error::other(format!("chart-title preflight failed: {error}")))?;
    let staged = match change {
        TitleChange::Set(title) => edit.set(title).map_err(|error| {
            io::Error::other(format!("staging chart-title set failed: {error}"))
        })?,
        TitleChange::Clear => edit.clear().map_err(|error| {
            io::Error::other(format!("staging chart-title clear failed: {error}"))
        })?,
    };
    let commit = staged
        .commit()
        .map_err(|error| io::Error::other(format!("chart-title commit failed: {error}")))?;
    let committed = commit
        .package()
        // A chart name selector may intentionally become stale when this
        // operation renames or clears that chart's visible title. The patch
        // retains the checked semantic position without exposing native IDs.
        .slide_chart_title(
            slide.selector(),
            ChartSelector::position(commit.patch().chart_position()),
        )
        .map_err(|error| {
            io::Error::other(format!("reading committed chart title failed: {error}"))
        })?;
    if committed.as_deref() != change.expected() {
        return Err(invalid_input(
            "committed package did not expose the requested chart title",
        ));
    }
    if commit.diagnostics().changed() != (before != committed) {
        return Err(invalid_input(
            "chart-title diagnostics disagreed with the semantic change",
        ));
    }

    let restored = commit
        .package()
        .apply_slide_chart_title(&commit.patch().inverse())
        .map_err(|error| io::Error::other(format!("chart-title inverse failed: {error}")))?;
    let restored_title = restored
        .package()
        .slide_chart_title(slide.selector(), chart.selector())
        .map_err(|error| {
            io::Error::other(format!("reading inverse chart title failed: {error}"))
        })?;
    let restored_bytes = exact_bytes(restored.package()).map_err(|error| {
        io::Error::other(format!(
            "serializing inverse chart-title package failed: {error}"
        ))
    })?;
    if restored_title != before || restored_bytes != source_bytes {
        return Err(invalid_input(
            "inverse patch did not restore the exact input package and chart title",
        ));
    }

    save_new(&output, commit.package())
        .map_err(|error| io::Error::other(format!("writing chart-title output failed: {error}")))?;
    let reopened = Package::open(output).map_err(|error| {
        io::Error::other(format!("reopening chart-title output failed: {error}"))
    })?;
    reopened.validate().map_err(|error| {
        io::Error::other(format!(
            "validating reopened chart-title output failed: {error}"
        ))
    })?;
    let after = reopened
        .slide_chart_title(
            slide.selector(),
            ChartSelector::position(commit.patch().chart_position()),
        )
        .map_err(|error| {
            io::Error::other(format!("reading reopened chart title failed: {error}"))
        })?;
    if after.as_deref() != change.expected() {
        return Err(invalid_input(
            "reopened output did not preserve the requested chart title",
        ));
    }

    Ok(Report {
        before_present: before.is_some(),
        after_present: after.is_some(),
        after_bytes: after.as_deref().map_or(0, str::len),
        changed: commit.diagnostics().changed(),
        touched_components: commit.diagnostics().touched_components(),
        deleted_previews: commit.diagnostics().deleted_previews(),
        full_reparse: commit.diagnostics().full_reparse_performed(),
        source_fingerprint: commit.patch().source_fingerprint(),
        target_fingerprint: commit.patch().target_fingerprint(),
    })
}

fn edit_axis_title(
    package: &Package,
    output: &Path,
    slide: &SelectedSlide,
    chart: &SelectedChart,
    axis: Axis,
    change: &TitleChange,
) -> Result<Report, Box<dyn Error>> {
    let before = package
        .slide_chart_axis_title(slide.selector(), chart.selector(), axis)
        .map_err(|error| {
            io::Error::other(format!("reading {} title failed: {error}", axis_name(axis)))
        })?;
    let source_bytes = exact_bytes(package)
        .map_err(|error| io::Error::other(format!("serializing input package failed: {error}")))?;

    let noop_edit = package
        .edit_slide_chart_axis_title(slide.selector(), chart.selector(), axis)
        .map_err(|error| {
            io::Error::other(format!(
                "{} no-op preflight failed: {error}",
                axis_name(axis)
            ))
        })?;
    let noop = match before.as_deref() {
        Some(title) => noop_edit.set(title).map_err(|error| {
            io::Error::other(format!(
                "staging {} no-op set failed: {error}",
                axis_name(axis)
            ))
        })?,
        None => noop_edit.clear().map_err(|error| {
            io::Error::other(format!(
                "staging {} no-op clear failed: {error}",
                axis_name(axis)
            ))
        })?,
    };
    let noop = noop.commit().map_err(|error| {
        io::Error::other(format!("{} no-op commit failed: {error}", axis_name(axis)))
    })?;
    let noop_title = noop
        .package()
        .slide_chart_axis_title(slide.selector(), chart.selector(), axis)
        .map_err(|error| {
            io::Error::other(format!(
                "reading {} no-op result failed: {error}",
                axis_name(axis)
            ))
        })?;
    let noop_bytes = exact_bytes(noop.package()).map_err(|error| {
        io::Error::other(format!(
            "serializing {} no-op result failed: {error}",
            axis_name(axis)
        ))
    })?;
    if noop.diagnostics().changed()
        || !noop.patch().is_noop()
        || noop_title != before
        || noop_bytes != source_bytes
    {
        return Err(invalid_input(format!(
            "setting the existing {} title was not an exact no-op",
            axis_name(axis)
        )));
    }

    let edit = package
        .edit_slide_chart_axis_title(slide.selector(), chart.selector(), axis)
        .map_err(|error| {
            io::Error::other(format!("{} preflight failed: {error}", axis_name(axis)))
        })?;
    let staged = match change {
        TitleChange::Set(title) => edit.set(title).map_err(|error| {
            io::Error::other(format!("staging {} set failed: {error}", axis_name(axis)))
        })?,
        TitleChange::Clear => edit.clear().map_err(|error| {
            io::Error::other(format!("staging {} clear failed: {error}", axis_name(axis)))
        })?,
    };
    let commit = staged
        .commit()
        .map_err(|error| io::Error::other(format!("{} commit failed: {error}", axis_name(axis))))?;
    let committed = commit
        .package()
        .slide_chart_axis_title(slide.selector(), chart.selector(), axis)
        .map_err(|error| {
            io::Error::other(format!(
                "reading committed {} failed: {error}",
                axis_name(axis)
            ))
        })?;
    if committed.as_deref() != change.expected() {
        return Err(invalid_input(format!(
            "committed package did not expose the requested {}",
            axis_name(axis)
        )));
    }
    if commit.diagnostics().changed() != (before != committed) {
        return Err(invalid_input(format!(
            "{} diagnostics disagreed with the semantic change",
            axis_name(axis)
        )));
    }

    let restored = commit
        .package()
        .apply_slide_chart_axis_title(&commit.patch().inverse())
        .map_err(|error| {
            io::Error::other(format!("{} inverse failed: {error}", axis_name(axis)))
        })?;
    let restored_title = restored
        .package()
        .slide_chart_axis_title(slide.selector(), chart.selector(), axis)
        .map_err(|error| {
            io::Error::other(format!(
                "reading inverse {} failed: {error}",
                axis_name(axis)
            ))
        })?;
    let restored_bytes = exact_bytes(restored.package()).map_err(|error| {
        io::Error::other(format!(
            "serializing inverse {} package failed: {error}",
            axis_name(axis)
        ))
    })?;
    if restored_title != before || restored_bytes != source_bytes {
        return Err(invalid_input(format!(
            "inverse patch did not restore the exact input package and {}",
            axis_name(axis)
        )));
    }

    save_new(&output, commit.package()).map_err(|error| {
        io::Error::other(format!(
            "writing {} output failed: {error}",
            axis_name(axis)
        ))
    })?;
    let reopened = Package::open(output).map_err(|error| {
        io::Error::other(format!(
            "reopening {} output failed: {error}",
            axis_name(axis)
        ))
    })?;
    reopened.validate().map_err(|error| {
        io::Error::other(format!(
            "validating reopened {} output failed: {error}",
            axis_name(axis)
        ))
    })?;
    let after = reopened
        .slide_chart_axis_title(slide.selector(), chart.selector(), axis)
        .map_err(|error| {
            io::Error::other(format!(
                "reading reopened {} failed: {error}",
                axis_name(axis)
            ))
        })?;
    if after.as_deref() != change.expected() {
        return Err(invalid_input(format!(
            "reopened output did not preserve the requested {}",
            axis_name(axis)
        )));
    }

    Ok(Report {
        before_present: before.is_some(),
        after_present: after.is_some(),
        after_bytes: after.as_deref().map_or(0, str::len),
        changed: commit.diagnostics().changed(),
        touched_components: commit.diagnostics().touched_components(),
        deleted_previews: commit.diagnostics().deleted_previews(),
        full_reparse: commit.diagnostics().full_reparse_performed(),
        source_fingerprint: commit.patch().source_fingerprint(),
        target_fingerprint: commit.patch().target_fingerprint(),
    })
}

fn axis_name(axis: Axis) -> &'static str {
    match axis {
        Axis::Category => "category-axis title",
        Axis::Value => "value-axis title",
    }
}

fn parse_target(argument: OsString) -> Result<TitleTarget, Box<dyn Error>> {
    let value = argument
        .into_string()
        .map_err(|_| invalid_input("title target must be valid UTF-8"))?;
    match value.as_str() {
        "chart" => Ok(TitleTarget::Chart),
        "category-axis" => Ok(TitleTarget::CategoryAxis),
        "value-axis" => Ok(TitleTarget::ValueAxis),
        _ => Err(invalid_input(
            "title target must be chart, category-axis, or value-axis",
        )),
    }
}

fn parse_change(argument: OsString) -> Result<TitleChange, Box<dyn Error>> {
    let value = argument
        .into_string()
        .map_err(|_| invalid_input("title change must be valid UTF-8"))?;
    if let Some(title) = value.strip_prefix("set:") {
        return Ok(TitleChange::Set(title.to_owned()));
    }
    if value == "clear" {
        return Ok(TitleChange::Clear);
    }
    Err(invalid_input("title change must be set:TITLE or clear"))
}

fn parse_slide(argument: Option<OsString>) -> Result<SelectedSlide, Box<dyn Error>> {
    parse_selector(argument, "slide", SelectedSlide::Index, SelectedSlide::Name)
}

fn parse_chart(argument: Option<OsString>) -> Result<SelectedChart, Box<dyn Error>> {
    parse_selector(argument, "chart", SelectedChart::Index, SelectedChart::Name)
}

fn parse_selector<T>(
    argument: Option<OsString>,
    kind: &str,
    from_index: impl FnOnce(usize) -> T,
    from_name: impl FnOnce(String) -> T,
) -> Result<T, Box<dyn Error>> {
    let Some(argument) = argument else {
        return Ok(from_index(0));
    };
    let value = argument
        .into_string()
        .map_err(|_| invalid_input(format!("{kind} selector must be valid UTF-8")))?;
    if let Some(index) = value.strip_prefix("index:") {
        if index.is_empty() || !index.bytes().all(|byte| byte.is_ascii_digit()) {
            return Err(invalid_input(format!(
                "{kind} index must be a non-negative decimal integer"
            )));
        }
        let index = index
            .parse::<usize>()
            .map_err(|_| invalid_input(format!("{kind} index is too large for this platform")))?;
        return Ok(from_index(index));
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

fn required_argument(
    arguments: &mut impl Iterator<Item = OsString>,
    message: &'static str,
) -> Result<OsString, Box<dyn Error>> {
    arguments.next().ok_or_else(|| invalid_input(message))
}

/// Publish through a sibling temporary file without replacing an existing
/// artifact. A failed write leaves the source and destination untouched.
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
