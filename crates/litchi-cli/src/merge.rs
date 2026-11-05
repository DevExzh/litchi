//! `merge` (records 0274 and 0295): fill a template (DOCX, PPTX, XLSX,
//! XLSB, DOC, PPT, XLS, RTF, ODT, OTH, ODM, ODS or ODP) from JSON or CSV
//! data through `File::merge`, one output per record.

use std::io::{Read as _, Write};
use std::path::{Path, PathBuf};

use clap::{Args, ValueEnum};
use litchi_ops::merge::{Data, Missing, Options, Summary};
use litchi_ops::transaction::SemanticLimits;
use litchi_ops::{Code, Error, Overwrite, Result, SaveOptions, SaveTo};
use serde_json::{Value as Json, json};

use crate::{exit, render};

/// `litchi merge TEMPLATE DATA --out OUT`, or OfficeCLI's
/// `litchi merge TEMPLATE OUTPUT --data JSON`.
#[derive(Debug, Args)]
pub(crate) struct Merge {
    /// The template (a word-processing, presentation or spreadsheet file)
    /// with `{{key}}` placeholders.
    pub(crate) template: PathBuf,
    /// The data file (`.json` or `.csv`; `-` reads stdin), or, with
    /// `--data`, the output file (OfficeCLI's form).
    pub(crate) second: String,
    /// OfficeCLI's data option: inline JSON, or a path to a `.json` (or
    /// `.csv`) file. The second argument is then the output.
    #[arg(long, value_name = "JSON|FILE")]
    pub(crate) data: Option<String>,
    /// The output file. With several records it must contain `{n}`, the
    /// one-based record number (`invoice-{n}.docx`).
    #[arg(long, short = 'o', value_name = "FILE")]
    pub(crate) out: Option<PathBuf>,
    /// The data format; by default the data file's extension, else JSON.
    #[arg(long, value_enum, value_name = "FORMAT")]
    pub(crate) format: Option<DataFormat>,
    /// Replace existing output files instead of refusing them.
    #[arg(long)]
    pub(crate) force: bool,
    /// What a placeholder whose key the record lacks becomes: `keep` it (the
    /// default, as OfficeCLI) or `empty`.
    #[arg(long, value_enum, value_name = "POLICY", default_value_t = MissingArg::Keep)]
    pub(crate) missing: MissingArg,
    /// Leave Word `MERGEFIELD` fields alone.
    #[arg(long)]
    pub(crate) no_fields: bool,
    /// The largest number of replacements in one output (default 65,536).
    #[arg(long, value_name = "N")]
    pub(crate) max_replacements: Option<usize>,
    /// The largest data source, and replacement payload per output, in bytes
    /// (default 16 MiB).
    #[arg(long, value_name = "BYTES")]
    pub(crate) max_bytes: Option<usize>,
}

/// A merge data format.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ValueEnum)]
pub(crate) enum DataFormat {
    /// A JSON object, or an array of objects.
    Json,
    /// CSV with a header row of keys.
    Csv,
    /// Tab-separated values with a header row of keys.
    Tsv,
}

/// `--missing`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ValueEnum)]
pub(crate) enum MissingArg {
    /// Keep the placeholder or field.
    Keep,
    /// Replace it by nothing.
    Empty,
}

pub(crate) fn merge(args: &Merge, json: bool, out: &mut impl Write) -> Result<u8> {
    let mut limits = SemanticLimits::builder();
    if let Some(count) = args.max_replacements {
        limits = limits.max_expanded_targets(count)?;
    }
    if let Some(bytes) = args.max_bytes {
        limits = limits.max_staged_bytes(bytes)?;
    }
    let limits = limits.build()?;
    let (source, extension, output) = match &args.data {
        Some(data) => {
            if args.out.is_some() {
                return Err(Error::new(
                    Code::InvalidArgument,
                    "with --data the second argument is the output; --out is not used",
                ));
            }
            let path = Path::new(data);
            let is_file = ["json", "csv", "tsv"]
                .iter()
                .any(|kind| has_extension(path, kind))
                && path.is_file();
            if is_file {
                (read(path)?, extension(path), PathBuf::from(&args.second))
            } else {
                (
                    data.as_bytes().to_vec(),
                    Some("json".to_owned()),
                    PathBuf::from(&args.second),
                )
            }
        },
        None => {
            let output = args.out.clone().ok_or_else(|| {
                Error::new(
                    Code::MissingArgument,
                    "merge needs --out (or OfficeCLI's --data)",
                )
                .with_suggestion("litchi merge template.docx data.json --out out.docx")
            })?;
            if args.second == "-" {
                let mut data = Vec::new();
                std::io::stdin()
                    .lock()
                    .read_to_end(&mut data)
                    .map_err(|error| {
                        Error::new(Code::IoError, format!("reading stdin failed: {error}"))
                    })?;
                (data, None, output)
            } else {
                let path = Path::new(&args.second);
                (read(path)?, extension(path), output)
            }
        },
    };
    let format = args.format.unwrap_or(match extension.as_deref() {
        Some("csv") => DataFormat::Csv,
        Some("tsv") => DataFormat::Tsv,
        _ => DataFormat::Json,
    });
    let data = match format {
        DataFormat::Json => Data::from_json(&source, &limits)?,
        DataFormat::Csv => Data::from_csv(&source, &litchi_ops::import::Options::csv(), &limits)?,
        DataFormat::Tsv => Data::from_csv(&source, &litchi_ops::import::Options::tsv(), &limits)?,
    };
    if data.is_empty() {
        return Err(Error::new(
            Code::InvalidArgument,
            "the merge data holds no record",
        ));
    }
    let pattern = output.to_string_lossy().into_owned();
    if data.len() > 1 && !pattern.contains("{n}") {
        return Err(Error::new(
            Code::InvalidArgument,
            format!(
                "the merge data holds {} records; the output needs `{{n}}` for the record number",
                data.len()
            ),
        )
        .with_suggestion("for example --out invoice-{n}.docx"));
    }
    let options = Options::new()
        .missing(match args.missing {
            MissingArg::Keep => Missing::Keep,
            MissingArg::Empty => Missing::Empty,
        })
        .fields(!args.no_fields);
    let overwrite = if args.force {
        Overwrite::Replace
    } else {
        Overwrite::Refuse
    };
    let paths: Vec<PathBuf> = (1..=data.len())
        .map(|number| PathBuf::from(pattern.replace("{n}", &number.to_string())))
        .collect();
    if !args.force
        && let Some(existing) = paths.iter().find(|path| path.exists())
    {
        return Err(Error::new(
            Code::FileExists,
            format!("the output {} already exists", existing.display()),
        )
        .with_suggestion("pass --force to replace it"));
    }
    let template = litchi_ops::open(&args.template)?;
    let edit_options = litchi_ops::EditOptions::new().semantic_limits(limits);
    let mut results = Vec::with_capacity(data.len());
    for (record, path) in data.records().iter().zip(&paths) {
        let commit = template.merge_with_options(record, &options, edit_options)?;
        let saved = commit
            .file()
            .save_with(SaveTo::Path(path), &SaveOptions::new().overwrite(overwrite))?;
        results.push((saved.path, commit.into_value()));
    }
    if json {
        let mut outputs: Vec<Json> = results
            .iter()
            .map(|(path, summary)| report(path, summary))
            .collect();
        let value = if outputs.len() == 1 {
            outputs.remove(0)
        } else {
            json!({ "records": outputs.len(), "outputs": outputs })
        };
        render::json(out, &value)?;
    } else {
        for (path, summary) in &results {
            writeln!(out, "Merged: {}", path.display()).map_err(|error| render::io(&error))?;
            writeln!(out, "  Replaced keys: {}", summary.keys().len())
                .map_err(|error| render::io(&error))?;
            if !summary.unresolved().is_empty() {
                eprintln!(
                    "  Warning: {} unresolved placeholder(s):",
                    summary.unresolved().len()
                );
                for key in summary.unresolved() {
                    eprintln!("    - {{{{{key}}}}}");
                }
            }
        }
    }
    Ok(exit::OK)
}

/// One output, in OfficeCLI's `merge --json` spellings plus Litchi's
/// counts.
fn report(path: &Path, summary: &Summary) -> Json {
    let absolute = std::path::absolute(path).unwrap_or_else(|_| path.to_path_buf());
    json!({
        "output": absolute.display().to_string(),
        "replacedKeys": summary.keys().len(),
        "replacements": summary.replacements(),
        "fields": summary.fields(),
        "skippedFields": summary.skipped_fields(),
        "unresolvedPlaceholders": summary.unresolved(),
    })
}

fn has_extension(path: &Path, kind: &str) -> bool {
    path.extension()
        .and_then(|extension| extension.to_str())
        .is_some_and(|extension| extension.eq_ignore_ascii_case(kind))
}

fn extension(path: &Path) -> Option<String> {
    path.extension()
        .and_then(|extension| extension.to_str())
        .map(str::to_ascii_lowercase)
}

fn read(path: &Path) -> Result<Vec<u8>> {
    std::fs::read(path).map_err(|error| {
        let code = if error.kind() == std::io::ErrorKind::NotFound {
            Code::FileNotFound
        } else {
            Code::IoError
        };
        Error::new(code, format!("reading the merge data failed: {error}"))
    })
}
