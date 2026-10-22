//! `refresh` (record 0279): recompute a document's derived field results
//! through `File::refresh` and save the file (in place by default, as
//! OfficeCLI's `refresh`).

use std::io::Write;
use std::path::PathBuf;

use clap::Args;
use litchi_ops::Result;
use litchi_ops::refresh::{Clock, Options, PageNumbers, Summary};
use litchi_ops::transaction::SemanticLimits;
use serde_json::{Value as Json, json};

use crate::cli::Output;
use crate::{exit, render};

/// `litchi refresh FILE`.
#[derive(Debug, Args)]
pub(crate) struct Refresh {
    /// The document (DOCX, DOC, ODT, OTH, ODM, PPTX, ODP, or a workbook).
    pub(crate) file: PathBuf,
    /// The date and time DATE, TIME and slide date fields show
    /// (`YYYY-MM-DD` or `YYYY-MM-DDTHH:MM[:SS]`); without it they keep
    /// their cached result. The system clock is never read.
    #[arg(long, value_name = "DATETIME")]
    pub(crate) clock: Option<String>,
    /// Write recalculated XLSX formula results back as cached values.
    #[arg(long)]
    pub(crate) write_cache: bool,
    /// Refuse (`refresh_failed`) a table of contents whose page numbers
    /// Litchi cannot know, instead of keeping cached ones.
    #[arg(long)]
    pub(crate) require_page_numbers: bool,
    /// Do not mark a table of contents with unknown page numbers dirty.
    #[arg(long)]
    pub(crate) no_dirty: bool,
    /// The largest number of changed results (default 65,536).
    #[arg(long, value_name = "N")]
    pub(crate) max_changes: Option<usize>,
    /// The largest changed payload in bytes (default 16 MiB).
    #[arg(long, value_name = "BYTES")]
    pub(crate) max_bytes: Option<usize>,
    #[command(flatten)]
    pub(crate) output: Output,
}

pub(crate) fn refresh(args: &Refresh, json: bool, out: &mut impl Write) -> Result<u8> {
    let mut limits = SemanticLimits::builder();
    if let Some(count) = args.max_changes {
        limits = limits.max_expanded_targets(count)?;
    }
    if let Some(bytes) = args.max_bytes {
        limits = limits.max_staged_bytes(bytes)?;
    }
    let limits = limits.build()?;
    let mut options = Options::new()
        .write_cache(args.write_cache)
        .mark_dirty(!args.no_dirty)
        .page_numbers(if args.require_page_numbers {
            PageNumbers::Require
        } else {
            PageNumbers::Keep
        });
    if let Some(clock) = &args.clock {
        options = options.clock(Clock::parse(clock)?);
    }
    let file = litchi_ops::open(&args.file)?;
    let commit = file.refresh_with_options(
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
        render::json(out, &report(saved.as_deref(), changed, &summary))?;
    } else {
        let mut text = match &saved {
            Some(path) => format!("Refreshed: {}\n", path.display()),
            None => format!("Up to date: {}\n", args.file.display()),
        };
        for field in summary.fields() {
            text.push_str(&format!(
                "  {} {} {}{}\n",
                field.status(),
                field.kind(),
                field.location().unwrap_or("-"),
                field
                    .reason()
                    .map(|reason| format!(" ({reason})"))
                    .unwrap_or_default()
            ));
        }
        if summary.missing_page_numbers() > 0 {
            text.push_str(&format!(
                "  Note: {} table of contents page number(s) need a word processor's layout\n",
                summary.missing_page_numbers()
            ));
        }
        out.write_all(text.as_bytes())
            .map_err(|error| render::io(&error))?;
    }
    Ok(exit::OK)
}

fn report(saved: Option<&std::path::Path>, changed: bool, summary: &Summary) -> Json {
    let fields: Vec<Json> = summary
        .fields()
        .iter()
        .map(|field| {
            json!({
                "type": field.kind(),
                "path": field.location(),
                "status": field.status().as_str(),
                "reason": field.reason(),
            })
        })
        .collect();
    json!({
        "output": saved.map(|path| path.display().to_string()),
        "changed": changed,
        "refreshed": summary.refreshed(),
        "bookmarks": summary.bookmarks(),
        "missingPageNumbers": summary.missing_page_numbers(),
        "fields": fields,
    })
}
