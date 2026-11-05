//! `import`, `evaluate` and `recalculate` (record 0266): CSV/TSV import into
//! a worksheet and formula evaluation, over `File::import`,
//! `File::evaluate` and `File::recalculate`.

use std::io::{Read as _, Write};
use std::path::{Path, PathBuf};

use clap::Args;
use litchi_ops::calc::{Options as CalcOptions, Origin, Status, Value};
use litchi_ops::import::{Encoding, Options};
use litchi_ops::transaction::SemanticLimits;
use litchi_ops::{Code, Error, Result};
use serde_json::{Value as Json, json};

use crate::cli::Output;
use crate::{exit, render};

/// `litchi import FILE PATH [SOURCE]`, OfficeCLI's import shape.
#[derive(Debug, Args)]
pub(crate) struct Import {
    /// The workbook (XLSX, XLSB, XLS or ODS).
    pub(crate) file: PathBuf,
    /// The worksheet (`/Sheet1`) or origin cell (`/Sheet1/B2`).
    pub(crate) path: String,
    /// The CSV/TSV file to import (or `--file`, or `--stdin`).
    #[arg(value_name = "SOURCE", conflicts_with_all = ["source_file", "stdin"])]
    pub(crate) source: Option<PathBuf>,
    /// The CSV/TSV file to import.
    #[arg(long = "file", value_name = "SOURCE", conflicts_with = "stdin")]
    pub(crate) source_file: Option<PathBuf>,
    /// Read the data from standard input.
    #[arg(long)]
    pub(crate) stdin: bool,
    /// `csv` or `tsv`; by default the source's `sep=` line, then its
    /// extension (`.tsv`, `.tab`), then CSV.
    #[arg(long, value_name = "FORMAT")]
    pub(crate) format: Option<String>,
    /// One ASCII field delimiter (`;`, `|`, or `\t`/`tab`); overrides the
    /// format, extension and `sep=` line.
    #[arg(long, value_name = "CHAR")]
    pub(crate) delimiter: Option<String>,
    /// The first record is a header row, written as text.
    #[arg(long)]
    pub(crate) header: bool,
    /// The origin cell when PATH names a sheet (default A1).
    #[arg(long, value_name = "CELL")]
    pub(crate) start_cell: Option<String>,
    /// auto, utf-8, utf-16le or utf-16be (default auto: the byte-order mark,
    /// else UTF-8).
    #[arg(long, value_name = "ENCODING")]
    pub(crate) encoding: Option<String>,
    /// Write every field as text: no number, date or boolean inference.
    #[arg(long)]
    pub(crate) no_infer: bool,
    /// Write fields that start with `=` as formulas (off by default).
    #[arg(long)]
    pub(crate) formulas: bool,
    /// The largest number of fields to import (default 65,536).
    #[arg(long, value_name = "N")]
    pub(crate) max_cells: Option<usize>,
    /// The largest source and staged payload in bytes (default 16 MiB).
    #[arg(long, value_name = "BYTES")]
    pub(crate) max_bytes: Option<usize>,
    #[command(flatten)]
    pub(crate) output: Output,
}

/// `litchi evaluate FILE PATH`.
#[derive(Debug, Args)]
pub(crate) struct Evaluate {
    /// The workbook (XLSX, XLSB or XLS).
    pub(crate) file: PathBuf,
    /// The cell, for example `/Sheet1/B2`.
    pub(crate) path: String,
    /// Compute the formula even when the workbook caches its result.
    #[arg(long)]
    pub(crate) recalculate: bool,
}

/// `litchi recalculate FILE`.
#[derive(Debug, Args)]
pub(crate) struct Recalculate {
    /// The workbook (XLSX, XLSB or XLS).
    pub(crate) file: PathBuf,
    /// List every formula, not only the mismatched and unsupported ones.
    #[arg(long)]
    pub(crate) all: bool,
}

pub(crate) fn import(args: &Import, json: bool, out: &mut impl Write) -> Result<u8> {
    let source_path = args.source.as_ref().or(args.source_file.as_ref());
    let data = match (source_path, args.stdin) {
        (Some(path), _) => read(path)?,
        (None, true) => {
            let mut data = Vec::new();
            std::io::stdin()
                .lock()
                .read_to_end(&mut data)
                .map_err(|error| {
                    Error::new(Code::IoError, format!("reading stdin failed: {error}"))
                })?;
            data
        },
        (None, false) => {
            return Err(Error::new(
                Code::MissingArgument,
                "import needs a SOURCE file, --file or --stdin",
            ));
        },
    };
    let mut options = match args.format.as_deref() {
        Some(format) if format.eq_ignore_ascii_case("tsv") => Options::tsv(),
        Some(format) if format.eq_ignore_ascii_case("csv") => Options::csv(),
        Some(format) => {
            return Err(Error::new(
                Code::InvalidValue,
                format!("unknown import format `{format}`"),
            )
            .with_suggestion("use csv or tsv"));
        },
        None if source_path.is_some_and(|path| tabbed(path)) && !declares_separator(&data) => {
            Options::tsv()
        },
        None => Options::csv(),
    };
    if let Some(delimiter) = &args.delimiter {
        options = options.delimiter(delimiter_of(delimiter)?);
    }
    if let Some(encoding) = &args.encoding {
        options = options.encoding(Encoding::parse(encoding).ok_or_else(|| {
            Error::new(Code::InvalidValue, format!("unknown encoding `{encoding}`"))
                .with_suggestion("use auto, utf-8, utf-16le or utf-16be")
        })?);
    }
    options = options
        .header(args.header)
        .infer(!args.no_infer)
        .formulas(args.formulas);
    let mut limits = SemanticLimits::builder();
    if let Some(cells) = args.max_cells {
        limits = limits.max_expanded_targets(cells)?;
    }
    if let Some(bytes) = args.max_bytes {
        limits = limits.max_staged_bytes(bytes)?;
    }
    let limits = limits.build()?;
    let target = match &args.start_cell {
        Some(cell) => format!("{}/{cell}", args.path.trim_end_matches('/')),
        None => args.path.clone(),
    };
    let file = litchi_ops::open(&args.file)?;
    let commit = file.import_with_options(
        target.as_str(),
        &data,
        &options,
        litchi_ops::EditOptions::new().semantic_limits(limits),
    )?;
    let changed = commit.diagnostics().changed();
    let saved = match &args.output.out {
        Some(path) => Some(commit.file().save_as(path)?.path),
        None if changed => Some(commit.file().save()?.path),
        None => None,
    };
    let summary = commit.into_value();
    if json {
        render::json(
            out,
            &json!({
                "sheet": summary.sheet() + 1,
                "origin": cell_label(summary.row(), summary.column()),
                "rows": summary.rows(),
                "columns": summary.columns(),
                "cells": summary.cells(),
                "texts": summary.texts(),
                "numbers": summary.numbers(),
                "booleans": summary.booleans(),
                "dates": summary.dates(),
                "datesAsText": summary.dates_as_text(),
                "formulas": summary.formulas(),
                "empty": summary.empty(),
                "changed": changed,
                "saved": saved.map(|path| path.display().to_string()),
            }),
        )?;
    } else if summary.rows() == 0 {
        writeln!(out, "No data to import").map_err(|error| render::io(&error))?;
    } else {
        writeln!(
            out,
            "Imported {} rows x {} cols into {} starting at {}",
            summary.rows(),
            summary.columns(),
            args.path,
            cell_label(summary.row(), summary.column())
        )
        .map_err(|error| render::io(&error))?;
    }
    Ok(exit::OK)
}

pub(crate) fn evaluate(args: &Evaluate, json: bool, out: &mut impl Write) -> Result<u8> {
    let file = litchi_ops::open(&args.file)?;
    let options = CalcOptions::new().recalculate(args.recalculate);
    let evaluated = file.evaluate(args.path.as_str(), &options)?;
    let origin = match evaluated.origin() {
        Origin::Literal => "literal",
        Origin::Cached => "cached",
        Origin::Computed => "computed",
        _ => "unknown",
    };
    if json {
        render::json(
            out,
            &json!({"path": args.path, "value": value(Some(evaluated.value())), "origin": origin}),
        )?;
    } else {
        writeln!(out, "{} ({origin})", text(evaluated.value()))
            .map_err(|error| render::io(&error))?;
    }
    Ok(exit::OK)
}

pub(crate) fn recalculate(args: &Recalculate, json: bool, out: &mut impl Write) -> Result<u8> {
    let file = litchi_ops::open(&args.file)?;
    let report = file.recalculate(&CalcOptions::new())?;
    let names = file.workbook()?.worksheet_names()?;
    let listed: Vec<Json> = report
        .formulas()
        .iter()
        .filter(|formula| {
            args.all || matches!(formula.status(), Status::Mismatch | Status::Unsupported)
        })
        .map(|formula| {
            json!({
                "path": format!(
                    "/{}/{}",
                    names.get(formula.sheet()).map_or("", String::as_str),
                    cell_label(formula.row(), formula.column())
                ),
                "formula": formula.text(),
                "cachedValue": value(formula.cached()),
                "computedValue": value(formula.computed()),
                "status": status(formula.status()),
            })
        })
        .collect();
    if json {
        render::json(
            out,
            &json!({
                "formulas": report.formulas().len(),
                "match": report.matches(),
                "mismatch": report.mismatches(),
                "uncached": report.uncached(),
                "unsupported": report.unsupported(),
                "volatile": report.volatile(),
                "cells": listed,
            }),
        )?;
    } else {
        let mut text = format!(
            "{} formulas: {} match, {} mismatch, {} uncached, {} unsupported, {} volatile\n",
            report.formulas().len(),
            report.matches(),
            report.mismatches(),
            report.uncached(),
            report.unsupported(),
            report.volatile()
        );
        for cell in &listed {
            text.push_str(&format!(
                "  {} {} ={} cached={} computed={}\n",
                cell["status"].as_str().unwrap_or(""),
                cell["path"].as_str().unwrap_or(""),
                cell["formula"].as_str().unwrap_or(""),
                cell["cachedValue"],
                cell["computedValue"]
            ));
        }
        out.write_all(text.as_bytes())
            .map_err(|error| render::io(&error))?;
    }
    Ok(exit::OK)
}

fn read(path: &Path) -> Result<Vec<u8>> {
    std::fs::read(path).map_err(|error| {
        let code = if error.kind() == std::io::ErrorKind::NotFound {
            Code::FileNotFound
        } else {
            Code::IoError
        };
        Error::new(code, format!("reading the import source failed: {error}"))
    })
}

fn tabbed(path: &Path) -> bool {
    path.extension()
        .and_then(|extension| extension.to_str())
        .is_some_and(|extension| {
            extension.eq_ignore_ascii_case("tsv") || extension.eq_ignore_ascii_case("tab")
        })
}

/// Whether the data starts with an Excel `sep=` line, which outranks the
/// extension.
fn declares_separator(data: &[u8]) -> bool {
    let data = data.strip_prefix(&[0xEF, 0xBB, 0xBF]).unwrap_or(data);
    data.starts_with(b"sep=")
}

fn delimiter_of(raw: &str) -> Result<char> {
    let value = match raw {
        "\\t" | "tab" | "TAB" => "\t",
        other => other,
    };
    let mut characters = value.chars();
    match (characters.next(), characters.next()) {
        (Some(character), None) => Ok(character),
        _ => Err(
            Error::new(Code::InvalidValue, "--delimiter must be a single character")
                .with_suggestion("use a literal separator such as ';' or '|', or \\t for a tab"),
        ),
    }
}

fn cell_label(row: u32, column: u32) -> String {
    let mut label = String::new();
    let mut column = column + 1;
    while column > 0 {
        let digit = u8::try_from((column - 1) % 26).unwrap_or(0);
        label.insert(0, char::from(b'A' + digit));
        column = (column - 1) / 26;
    }
    format!("{label}{}", row + 1)
}

fn value(value: Option<&Value>) -> Json {
    match value {
        None => Json::Null,
        Some(Value::Empty) => json!(""),
        Some(Value::Bool(value)) => json!(value),
        Some(Value::Number(value)) => json!(value),
        Some(Value::Text(text)) => json!(text),
        Some(Value::Error(error)) => json!(error.as_str()),
        Some(_) => Json::Null,
    }
}

fn text(value: &Value) -> String {
    match value {
        Value::Empty => String::new(),
        Value::Bool(true) => "TRUE".to_owned(),
        Value::Bool(false) => "FALSE".to_owned(),
        Value::Number(number) => number.to_string(),
        Value::Text(text) => text.clone(),
        Value::Error(error) => error.as_str().to_owned(),
        _ => String::new(),
    }
}

fn status(status: Status) -> &'static str {
    match status {
        Status::Match => "match",
        Status::Mismatch => "mismatch",
        Status::Uncached => "uncached",
        Status::Unsupported => "unsupported",
        Status::Volatile => "volatile",
        _ => "unknown",
    }
}
