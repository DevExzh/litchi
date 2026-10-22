//! `File.import_csv`, `File.evaluate` and `File.recalculate` (record 0266):
//! the Rust work runs with the GIL released; the results are built as dicts.

use litchi::calc::{Options as CalcOptions, Origin, Status, Value};
use litchi::import::{Encoding, Options};
use litchi::transaction::SemanticLimits;
use pyo3::prelude::*;
use pyo3::types::{PyBytes, PyDict, PyList, PyString};

use crate::errors::Failure;
use crate::write::Commit;

/// The source text: `bytes` as given, `str` as UTF-8.
pub(crate) fn source(data: &Bound<'_, PyAny>) -> PyResult<Vec<u8>> {
    if let Ok(bytes) = data.cast::<PyBytes>() {
        return Ok(bytes.as_bytes().to_vec());
    }
    if let Ok(text) = data.cast::<PyString>() {
        let text: pyo3::pybacked::PyBackedStr = text.clone().try_into()?;
        return Ok(text.as_bytes().to_vec());
    }
    Err(Failure::argument("data must be bytes or str").into())
}

fn one_char(name: &str, value: &str) -> Result<char, Failure> {
    let value = match value {
        "\\t" | "tab" => "\t",
        other => other,
    };
    let mut characters = value.chars();
    match (characters.next(), characters.next()) {
        (Some(character), None) => Ok(character),
        _ => Err(Failure::argument(format!("{name} must be one character"))),
    }
}

/// The import options and limits from the keyword arguments.
#[allow(clippy::too_many_arguments, reason = "the Python keyword arguments")]
pub(crate) fn options(
    delimiter: Option<&str>,
    quote: Option<&str>,
    header: bool,
    encoding: &str,
    infer: bool,
    formulas: bool,
    max_cells: Option<usize>,
    max_bytes: Option<usize>,
) -> Result<(Options, SemanticLimits), Failure> {
    let mut options = Options::csv()
        .header(header)
        .infer(infer)
        .formulas(formulas)
        .quote(quote.map(|quote| one_char("quote", quote)).transpose()?);
    if let Some(delimiter) = delimiter {
        options = options.delimiter(one_char("delimiter", delimiter)?);
    }
    options = options.encoding(Encoding::parse(encoding).ok_or_else(|| {
        Failure::argument(format!("unknown encoding {encoding:?}"))
            .with_suggestion("use \"auto\", \"utf-8\", \"utf-16le\" or \"utf-16be\"")
    })?);
    let read = |error: litchi::Error| Failure::read(&error);
    let mut limits = SemanticLimits::builder();
    if let Some(cells) = max_cells {
        limits = limits.max_expanded_targets(cells).map_err(read)?;
    }
    if let Some(bytes) = max_bytes {
        limits = limits.max_staged_bytes(bytes).map_err(read)?;
    }
    Ok((options, limits.build().map_err(read)?))
}

/// The Commit of an import: `value` is the summary dict.
pub(crate) fn import_commit(
    py: Python<'_>,
    commit: litchi::transaction::Commit<litchi::import::Summary>,
) -> PyResult<Commit> {
    let diagnostics = *commit.diagnostics();
    let (file, patch, summary) = commit.into_parts();
    let dict = PyDict::new(py);
    dict.set_item("sheet", summary.sheet())?;
    dict.set_item("row", summary.row())?;
    dict.set_item("column", summary.column())?;
    dict.set_item("rows", summary.rows())?;
    dict.set_item("columns", summary.columns())?;
    dict.set_item("cells", summary.cells())?;
    dict.set_item("texts", summary.texts())?;
    dict.set_item("numbers", summary.numbers())?;
    dict.set_item("booleans", summary.booleans())?;
    dict.set_item("dates", summary.dates())?;
    dict.set_item("dates_as_text", summary.dates_as_text())?;
    dict.set_item("formulas", summary.formulas())?;
    dict.set_item("empty", summary.empty())?;
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

/// A value as Python: None (empty), bool, float, str, or the error's
/// spelling (`"#DIV/0!"`).
fn value<'py>(py: Python<'py>, value: Option<&Value>) -> PyResult<Bound<'py, PyAny>> {
    Ok(match value {
        None | Some(Value::Empty) => py.None().into_bound(py),
        Some(Value::Bool(value)) => value.into_pyobject(py)?.to_owned().into_any(),
        Some(Value::Number(number)) => number.into_pyobject(py)?.into_any(),
        Some(Value::Text(text)) => PyString::new(py, text).into_any(),
        Some(Value::Error(error)) => PyString::new(py, error.as_str()).into_any(),
        Some(_) => py.None().into_bound(py),
    })
}

/// `{"value": ..., "origin": "literal" | "cached" | "computed"}`.
pub(crate) fn evaluated<'py>(
    py: Python<'py>,
    evaluated: &litchi::calc::Evaluated,
) -> PyResult<Bound<'py, PyDict>> {
    let dict = PyDict::new(py);
    dict.set_item("value", value(py, Some(evaluated.value()))?)?;
    dict.set_item(
        "origin",
        match evaluated.origin() {
            Origin::Literal => "literal",
            Origin::Cached => "cached",
            Origin::Computed => "computed",
            _ => "unknown",
        },
    )?;
    Ok(dict)
}

/// The report dict: the counts and one dict per formula.
pub(crate) fn report<'py>(
    py: Python<'py>,
    report: &litchi::calc::Report,
) -> PyResult<Bound<'py, PyDict>> {
    let dict = PyDict::new(py);
    dict.set_item("formulas", report.formulas().len())?;
    dict.set_item("match", report.matches())?;
    dict.set_item("mismatch", report.mismatches())?;
    dict.set_item("uncached", report.uncached())?;
    dict.set_item("unsupported", report.unsupported())?;
    dict.set_item("volatile", report.volatile())?;
    let cells = PyList::empty(py);
    for formula in report.formulas() {
        let cell = PyDict::new(py);
        cell.set_item("sheet", formula.sheet())?;
        cell.set_item("row", formula.row())?;
        cell.set_item("column", formula.column())?;
        cell.set_item("formula", formula.text())?;
        cell.set_item("cached_value", value(py, formula.cached())?)?;
        cell.set_item("computed_value", value(py, formula.computed())?)?;
        cell.set_item(
            "status",
            match formula.status() {
                Status::Match => "match",
                Status::Mismatch => "mismatch",
                Status::Uncached => "uncached",
                Status::Unsupported => "unsupported",
                Status::Volatile => "volatile",
                _ => "unknown",
            },
        )?;
        cells.append(cell)?;
    }
    dict.set_item("cells", cells)?;
    Ok(dict)
}

/// The evaluation options.
pub(crate) fn calc_options(recalculate: bool) -> CalcOptions {
    CalcOptions::new().recalculate(recalculate)
}
