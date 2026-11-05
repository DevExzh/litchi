//! `File.refresh` (record 0279): the options from keyword arguments, the
//! refresh with the GIL released, and the summary as a dict.

use litchi::refresh::{Clock, Options, PageNumbers, Summary};
use litchi::transaction::SemanticLimits;
use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList, PyString};

use crate::errors::Failure;
use crate::write::Commit;

/// The clock text of `clock`: a str as given, a `datetime.datetime` as its
/// local wall time to the second, a `datetime.date` as its date.
pub(crate) fn clock_text(clock: &Bound<'_, PyAny>) -> PyResult<String> {
    if clock.is_instance_of::<PyString>() {
        return clock.extract();
    }
    let pattern = if clock.hasattr("hour")? {
        "%Y-%m-%dT%H:%M:%S"
    } else {
        "%Y-%m-%d"
    };
    let text = clock.call_method1("strftime", (pattern,)).map_err(|_| {
        Failure::argument("clock must be a str, a datetime.datetime or a datetime.date")
    })?;
    text.extract()
}

/// The options and limits from the keyword arguments.
pub(crate) fn options(
    clock: Option<&str>,
    write_cache: bool,
    page_numbers: &str,
    mark_dirty: bool,
    max_changes: Option<usize>,
    max_bytes: Option<usize>,
) -> Result<(Options, SemanticLimits), Failure> {
    let read = |error: litchi::Error| Failure::read(&error);
    let policy = match page_numbers {
        "keep" => PageNumbers::Keep,
        "require" => PageNumbers::Require,
        other => {
            return Err(
                Failure::argument(format!("unknown page_numbers policy {other:?}"))
                    .with_suggestion("use \"keep\" or \"require\""),
            );
        },
    };
    let mut options = Options::new()
        .write_cache(write_cache)
        .mark_dirty(mark_dirty)
        .page_numbers(policy);
    if let Some(clock) = clock {
        options = options.clock(Clock::parse(clock).map_err(read)?);
    }
    let mut limits = SemanticLimits::builder();
    if let Some(count) = max_changes {
        limits = limits.max_expanded_targets(count).map_err(read)?;
    }
    if let Some(bytes) = max_bytes {
        limits = limits.max_staged_bytes(bytes).map_err(read)?;
    }
    Ok((options, limits.build().map_err(read)?))
}

/// The Commit of a refresh: `value` is the summary dict.
pub(crate) fn commit(
    py: Python<'_>,
    commit: litchi::transaction::Commit<Summary>,
) -> PyResult<Commit> {
    let diagnostics = *commit.diagnostics();
    let (file, patch, summary) = commit.into_parts();
    let fields = PyList::empty(py);
    for field in summary.fields() {
        let entry = PyDict::new(py);
        entry.set_item("type", field.kind())?;
        entry.set_item("path", field.location())?;
        entry.set_item("status", field.status().as_str())?;
        entry.set_item("reason", field.reason())?;
        fields.append(entry)?;
    }
    let dict = PyDict::new(py);
    dict.set_item("refreshed", summary.refreshed())?;
    dict.set_item("bookmarks", summary.bookmarks())?;
    dict.set_item("missing_page_numbers", summary.missing_page_numbers())?;
    dict.set_item("fields", fields)?;
    Commit::new(
        py,
        file,
        Some(crate::write::Patch::one(patch)),
        diagnostics.changed(),
        diagnostics.operations(),
        None,
        Some(dict.into_any().unbind()),
    )
}
