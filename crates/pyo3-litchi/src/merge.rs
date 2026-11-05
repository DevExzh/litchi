//! `File.merge` (record 0274): the data record from a dict or JSON text,
//! the merge with the GIL released, and the summary as a dict.

use litchi::merge::{Data, Missing, Options, Record, Summary};
use litchi::transaction::SemanticLimits;
use pyo3::prelude::*;
use pyo3::types::{PyBytes, PyDict, PyList, PyString};

use crate::errors::Failure;
use crate::write::Commit;

/// The JSON text of `data`: a mapping through the standard `json` module,
/// `str` and `bytes` as given.
pub(crate) fn json_text(data: &Bound<'_, PyAny>) -> PyResult<Vec<u8>> {
    if let Ok(bytes) = data.cast::<PyBytes>() {
        return Ok(bytes.as_bytes().to_vec());
    }
    if let Ok(text) = data.cast::<PyString>() {
        let text: pyo3::pybacked::PyBackedStr = text.clone().try_into()?;
        return Ok(text.as_bytes().to_vec());
    }
    let json = data.py().import("json")?;
    let kwargs = PyDict::new(data.py());
    kwargs.set_item("ensure_ascii", false)?;
    let text = json
        .call_method("dumps", (data,), Some(&kwargs))
        .map_err(|_| Failure::argument("data must be a JSON-serializable mapping, or JSON text"))?;
    let text: pyo3::pybacked::PyBackedStr = text.extract()?;
    Ok(text.as_bytes().to_vec())
}

/// The one record of `source`, under `limits`.
pub(crate) fn record(source: &[u8], limits: &SemanticLimits) -> Result<Record, Failure> {
    let data = Data::from_json(source, limits).map_err(|error| Failure::read(&error))?;
    let mut records = data.into_records();
    if records.len() != 1 {
        return Err(
            Failure::argument("merge data must be one record: a mapping or a JSON object")
                .with_suggestion("call merge once per record"),
        );
    }
    Ok(records.remove(0))
}

/// The options and limits from the keyword arguments.
pub(crate) fn options(
    missing: &str,
    fields: bool,
    max_replacements: Option<usize>,
    max_bytes: Option<usize>,
) -> Result<(Options, SemanticLimits), Failure> {
    let missing = match missing {
        "keep" => Missing::Keep,
        "empty" => Missing::Empty,
        other => {
            return Err(
                Failure::argument(format!("unknown missing policy {other:?}"))
                    .with_suggestion("use \"keep\" or \"empty\""),
            );
        },
    };
    let read = |error: litchi::Error| Failure::read(&error);
    let mut limits = SemanticLimits::builder();
    if let Some(count) = max_replacements {
        limits = limits.max_expanded_targets(count).map_err(read)?;
    }
    if let Some(bytes) = max_bytes {
        limits = limits.max_staged_bytes(bytes).map_err(read)?;
    }
    Ok((
        Options::new().missing(missing).fields(fields),
        limits.build().map_err(read)?,
    ))
}

/// The Commit of a merge: `value` is the summary dict.
pub(crate) fn commit(
    py: Python<'_>,
    commit: litchi::transaction::Commit<Summary>,
) -> PyResult<Commit> {
    let diagnostics = *commit.diagnostics();
    let (file, patch, summary) = commit.into_parts();
    let dict = PyDict::new(py);
    dict.set_item("replacements", summary.replacements())?;
    dict.set_item("fields", summary.fields())?;
    dict.set_item("skipped_fields", summary.skipped_fields())?;
    dict.set_item("keys", PyList::new(py, summary.keys())?)?;
    dict.set_item("unresolved", PyList::new(py, summary.unresolved())?)?;
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
