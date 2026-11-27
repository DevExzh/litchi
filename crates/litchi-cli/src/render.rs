//! Output: node views, batch results and errors, as text or JSON.
//!
//! JSON shapes:
//!
//! - `get`: the `litchi-ops` node view, serialized by
//!   `litchi_ops::node::View::serialize_into` (OfficeCLI's `DocumentNode`).
//! - `query`: OfficeCLI's `NodesResult`, `{"matches": N, "results": [view,
//!   ...]}`, each result the node view above.
//! - `view --json`: `litchi_ops::File::view_json`.
//! - `capabilities --json`: `litchi_ops::capabilities::Capabilities`.
//! - writes (`set`, `add`, `remove`, `move`, `swap`, `batch`): OfficeCLI's
//!   batch rendering, `{"results": [...], "summary": {...}}`, built by
//!   [`Report`] from the public accessors of `litchi_ops::BatchCommit`.
//! - errors, on stderr: OfficeCLI's `ErrorResult`, `{"error": message,
//!   "code": code, "suggestion": ...}`, plus `item` and `property` when the
//!   error names them.

use std::io::Write;
use std::sync::atomic::{AtomicBool, Ordering};

use std::cell::RefCell;

use litchi_ops::batch::{Applied, BatchItem, ItemResult};
use litchi_ops::node::{Node, SerializeError, View, ViewOptions};
use litchi_ops::{Code, Error, Result};
use serde::Serialize;
use serde::ser::{Error as _, SerializeMap as _, SerializeSeq as _, Serializer};
use serde_json::ser::PrettyFormatter;

/// Set when stdout's reader went away (`litchi get ... | head`). Such a
/// run ends quietly, as Unix filters do, instead of reporting `io_error`.
static BROKEN_PIPE: AtomicBool = AtomicBool::new(false);

/// Whether a write to stdout failed because its reader closed the pipe.
pub(crate) fn broken_pipe() -> bool {
    BROKEN_PIPE.load(Ordering::Relaxed)
}

/// The error a failed write to stdout reports.
pub(crate) fn io(error: &std::io::Error) -> Error {
    if error.kind() == std::io::ErrorKind::BrokenPipe {
        BROKEN_PIPE.store(true, Ordering::Relaxed);
    }
    Error::new(Code::IoError, format!("writing the output failed: {error}"))
}

fn json_error(error: &serde_json::Error) -> Error {
    if let Some(kind) = error.io_error_kind() {
        return io(&std::io::Error::from(kind));
    }
    Error::new(
        Code::InternalError,
        format!("serializing the output failed: {error}"),
    )
}

/// Writes `value` as indented JSON and a newline, as OfficeCLI does.
pub(crate) fn json(out: &mut impl Write, value: &impl Serialize) -> Result<()> {
    serde_json::to_writer_pretty(&mut *out, value).map_err(|error| json_error(&error))?;
    writeln!(out).map_err(|error| io(&error))
}

/// Writes the node view as JSON, streaming it into `out` and keeping a view
/// failure (`output_too_large`, a leaf read error) typed.
pub(crate) fn view_json(out: &mut impl Write, view: &View<'_, '_>) -> Result<()> {
    let mut serializer = serde_json::Serializer::with_formatter(&mut *out, PrettyFormatter::new());
    view.serialize_into(&mut serializer)
        .map_err(|error| match error {
            SerializeError::View(error) => error,
            SerializeError::Serializer(error) => json_error(&error),
            // `SerializeError` is non-exhaustive.
            _ => Error::new(Code::InternalError, "serializing the node view failed"),
        })?;
    writeln!(out).map_err(|error| io(&error))
}

/// Writes the views of `nodes` as OfficeCLI's `NodesResult`,
/// `{"matches": N, "results": [view, ...]}`, streaming each view and keeping
/// the first view failure typed.
pub(crate) fn nodes_json(
    out: &mut impl Write,
    nodes: &[Node<'_>],
    options: ViewOptions,
) -> Result<()> {
    let failure = RefCell::new(None);
    let result = Nodes {
        nodes,
        options,
        failure: &failure,
    };
    let written = serde_json::to_writer_pretty(&mut *out, &result);
    if let Some(error) = failure.into_inner() {
        return Err(error);
    }
    written.map_err(|error| json_error(&error))?;
    writeln!(out).map_err(|error| io(&error))
}

struct Nodes<'a, 'f> {
    nodes: &'a [Node<'f>],
    options: ViewOptions,
    failure: &'a RefCell<Option<Error>>,
}

impl Serialize for Nodes<'_, '_> {
    fn serialize<S: Serializer>(&self, serializer: S) -> std::result::Result<S::Ok, S::Error> {
        let mut map = serializer.serialize_map(Some(2))?;
        map.serialize_entry("matches", &self.nodes.len())?;
        map.serialize_entry(
            "results",
            &Views {
                nodes: self.nodes,
                options: self.options,
                failure: self.failure,
            },
        )?;
        map.end()
    }
}

struct Views<'a, 'f> {
    nodes: &'a [Node<'f>],
    options: ViewOptions,
    failure: &'a RefCell<Option<Error>>,
}

impl Serialize for Views<'_, '_> {
    fn serialize<S: Serializer>(&self, serializer: S) -> std::result::Result<S::Ok, S::Error> {
        let mut seq = serializer.serialize_seq(Some(self.nodes.len()))?;
        for node in self.nodes {
            seq.serialize_element(&OneView {
                view: node.view_with(self.options),
                failure: self.failure,
            })?;
        }
        seq.end()
    }
}

struct OneView<'a, 'n, 'f> {
    view: View<'n, 'f>,
    failure: &'a RefCell<Option<Error>>,
}

impl Serialize for OneView<'_, '_, '_> {
    fn serialize<S: Serializer>(&self, serializer: S) -> std::result::Result<S::Ok, S::Error> {
        self.view
            .serialize_into(serializer)
            .map_err(|error| match error {
                SerializeError::View(error) => {
                    let custom = S::Error::custom(&error);
                    *self.failure.borrow_mut() = Some(error);
                    custom
                },
                SerializeError::Serializer(error) => error,
                _ => S::Error::custom("serializing the node view failed"),
            })
    }
}

/// Writes a node view as an indented outline: the node's path, type and
/// text, its `format` entries, then its children.
pub(crate) fn node_text(out: &mut impl Write, view: &View<'_, '_>, indent: usize) -> Result<()> {
    let mut line = format!("{:indent$}{} ({})", "", view.path(), view.kind());
    if let Some(text) = view.text()? {
        line.push_str(&format!(" {:?}", text.as_ref()));
    } else if let Some(preview) = view.preview()? {
        line.push_str(&format!(" {:?}", preview.to_string()));
    }
    if let Some(style) = view.style()? {
        line.push_str(&format!(" [style={style}]"));
    }
    let listed = view.options().depth > 0;
    if !listed {
        if let Some(count) = view.child_count()?.filter(|&count| count > 0) {
            line.push_str(&format!(" ({count} children)"));
        }
    }
    writeln!(out, "{line}").map_err(|error| io(&error))?;
    for prop in view.format()? {
        writeln!(
            out,
            "{:width$}{}={prop}",
            "",
            prop.key(),
            width = indent + 2
        )
        .map_err(|error| io(&error))?;
    }
    if listed {
        let options = view.child_options();
        for child in view.children() {
            node_text(out, &child?.view_with(options), indent + 2)?;
        }
    }
    Ok(())
}

/// Writes an error on stderr: text or OfficeCLI's `ErrorResult` JSON. A
/// `--unsupported warn` command that failed after skipping properties
/// prints each skipped property as a `warning[...]` first (JSON: the
/// error's `warnings`), so the failure is never shown without its cause.
pub(crate) fn error(error: &Error, json: bool) {
    if json {
        let mut stderr = std::io::stderr().lock();
        let value = ErrorJson::of(error);
        let _ = serde_json::to_writer_pretty(&mut stderr, &value);
        let _ = writeln!(stderr);
        return;
    }
    warnings(error.warnings());
    let mut stderr = std::io::stderr().lock();
    let _ = writeln!(stderr, "error[{}]: {}", error.code(), error.message());
    if let Some(item) = error.item() {
        let _ = writeln!(stderr, "  item: {item}");
    }
    if let Some(property) = error.property() {
        let _ = writeln!(stderr, "  property: {property}");
    }
    if let Some(suggestion) = error.suggestion() {
        let _ = writeln!(stderr, "  suggestion: {suggestion}");
    }
}

/// Writes each property a `--unsupported warn` command skipped on stderr:
/// `warning[<code>]: <message>` and its `item`, `property` and `suggestion`
/// lines, the shape of an error.
pub(crate) fn warnings(warnings: &[Error]) {
    if warnings.is_empty() {
        return;
    }
    let mut stderr = std::io::stderr().lock();
    for warning in warnings {
        let _ = writeln!(stderr, "warning[{}]: {}", warning.code(), warning.message());
        if let Some(item) = warning.item() {
            let _ = writeln!(stderr, "  item: {item}");
        }
        if let Some(property) = warning.property() {
            let _ = writeln!(stderr, "  property: {property}");
        }
        if let Some(suggestion) = warning.suggestion() {
            let _ = writeln!(stderr, "  suggestion: {suggestion}");
        }
    }
}

/// OfficeCLI's `ErrorResult`, plus the item and property Litchi reports.
#[derive(Serialize)]
struct ErrorJson<'a> {
    error: &'a str,
    code: &'static str,
    #[serde(skip_serializing_if = "Option::is_none")]
    suggestion: Option<&'a str>,
    #[serde(skip_serializing_if = "Option::is_none")]
    item: Option<usize>,
    #[serde(skip_serializing_if = "Option::is_none")]
    property: Option<&'a str>,
    /// The properties a `--unsupported warn` command skipped before it
    /// failed (`litchi_ops::Error::warnings`).
    #[serde(skip_serializing_if = "Vec::is_empty")]
    warnings: Vec<ErrorJson<'a>>,
}

impl<'a> ErrorJson<'a> {
    fn of(error: &'a Error) -> Self {
        Self {
            error: error.message(),
            code: error.code().as_str(),
            suggestion: error.suggestion(),
            item: error.item(),
            property: error.property(),
            warnings: error.warnings().iter().map(ErrorJson::of).collect(),
        }
    }
}

/// One item of a write's result, OfficeCLI's `BatchResult` shape.
#[derive(Serialize)]
pub(crate) struct ItemJson<'a> {
    index: usize,
    success: bool,
    /// Whether the item changed the document (successes only).
    #[serde(skip_serializing_if = "Option::is_none")]
    changed: Option<bool>,
    /// The canonical path `add` created.
    #[serde(skip_serializing_if = "Option::is_none")]
    path: Option<String>,
    /// The match count of a `set` with `find` (OfficeCLI's `matched`).
    #[serde(skip_serializing_if = "Option::is_none")]
    matched: Option<usize>,
    #[serde(skip_serializing_if = "Option::is_none")]
    error: Option<&'a str>,
    #[serde(skip_serializing_if = "Option::is_none")]
    code: Option<&'static str>,
    #[serde(skip_serializing_if = "Option::is_none")]
    suggestion: Option<&'a str>,
    #[serde(skip_serializing_if = "Option::is_none")]
    property: Option<&'a str>,
    /// The failed input item, echoed so a caller can inspect and retry it.
    #[serde(skip_serializing_if = "Option::is_none")]
    item: Option<&'a BatchItem<'a>>,
    /// The properties a `--unsupported warn` item skipped, each in
    /// `ErrorResult` shape (OfficeCLI's `warnings`).
    #[serde(skip_serializing_if = "Vec::is_empty")]
    warnings: Vec<ErrorJson<'a>>,
}

impl<'a> ItemJson<'a> {
    /// Whether the item succeeded.
    pub(crate) const fn success(&self) -> bool {
        self.success
    }

    /// A successful item.
    pub(crate) fn applied(index: usize, applied: &'a Applied) -> Self {
        Self {
            index,
            success: true,
            changed: Some(applied.changed),
            path: applied.path.as_ref().map(ToString::to_string),
            matched: applied.matched(),
            error: None,
            code: None,
            suggestion: None,
            property: None,
            item: None,
            warnings: applied.warnings.iter().map(ErrorJson::of).collect(),
        }
    }

    /// A failed item; `item` is its input when known.
    pub(crate) fn failed(index: usize, error: &'a Error, item: Option<&'a BatchItem<'a>>) -> Self {
        Self {
            index,
            success: false,
            changed: None,
            path: None,
            matched: None,
            error: Some(error.message()),
            code: Some(error.code().as_str()),
            suggestion: error.suggestion(),
            property: error.property(),
            item,
            warnings: error.warnings().iter().map(ErrorJson::of).collect(),
        }
    }

    /// The item's outcome in a batch.
    pub(crate) fn of(
        index: usize,
        result: &'a ItemResult,
        item: Option<&'a BatchItem<'a>>,
    ) -> Self {
        match result {
            ItemResult::Applied(applied) => Self::applied(index, applied),
            ItemResult::Failed(error) => Self::failed(index, error, item),
            // `ItemResult` is non-exhaustive.
            _ => Self {
                index,
                success: false,
                changed: None,
                path: None,
                matched: None,
                error: Some("unrecognized item outcome"),
                code: Some(Code::InternalError.as_str()),
                suggestion: None,
                property: None,
                item,
                warnings: Vec::new(),
            },
        }
    }
}

/// OfficeCLI's batch summary, plus whether the document changed and where
/// it was saved.
#[derive(Serialize, Default)]
#[serde(rename_all = "camelCase")]
pub(crate) struct Summary {
    pub(crate) total: usize,
    pub(crate) executed: usize,
    pub(crate) succeeded: usize,
    pub(crate) failed: usize,
    pub(crate) skipped: usize,
    pub(crate) changed: bool,
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub(crate) atomic_rolled_back: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) saved: Option<String>,
}

/// A write's result: OfficeCLI's `{"results": [...], "summary": {...}}`.
#[derive(Serialize)]
pub(crate) struct Report<'a> {
    pub(crate) results: Vec<ItemJson<'a>>,
    pub(crate) summary: Summary,
}

impl Report<'_> {
    /// Writes the report as JSON, or as one line per item and a summary.
    pub(crate) fn write(&self, out: &mut impl Write, json: bool, verb: &str) -> Result<()> {
        if json {
            return self::json(out, self);
        }
        let single = self.summary.total == 1;
        for item in &self.results {
            let line = if item.success {
                let target = item
                    .path
                    .as_deref()
                    .map(|path| format!(" {path}"))
                    .unwrap_or_default();
                let mut effect = match item.matched {
                    Some(matched) => format!(" ({matched} matched)"),
                    None => String::new(),
                };
                if item.changed == Some(false) {
                    effect.push_str(" (no change)");
                }
                if !item.warnings.is_empty() {
                    let skipped: Vec<&str> = item
                        .warnings
                        .iter()
                        .map(|warning| warning.property.unwrap_or("?"))
                        .collect();
                    effect.push_str(&format!(" (skipped: {})", skipped.join(", ")));
                }
                if single {
                    format!("{verb}{target}: ok{effect}")
                } else {
                    format!("#{}: ok{target}{effect}", item.index)
                }
            } else {
                let mut line = format!(
                    "#{}: error[{}]: {}",
                    item.index,
                    item.code.unwrap_or("internal_error"),
                    item.error.unwrap_or_default()
                );
                if !item.warnings.is_empty() {
                    let skipped: Vec<&str> = item
                        .warnings
                        .iter()
                        .map(|warning| warning.property.unwrap_or("?"))
                        .collect();
                    line.push_str(&format!(" (after skipping: {})", skipped.join(", ")));
                }
                line
            };
            writeln!(out, "{line}").map_err(|error| io(&error))?;
        }
        let summary = &self.summary;
        if !single {
            writeln!(
                out,
                "{} items: {} succeeded, {} failed, {} skipped{}",
                summary.total,
                summary.succeeded,
                summary.failed,
                summary.skipped,
                if summary.atomic_rolled_back {
                    "; rolled back"
                } else {
                    ""
                }
            )
            .map_err(|error| io(&error))?;
        }
        if let Some(saved) = &summary.saved {
            writeln!(out, "saved {saved}").map_err(|error| io(&error))?;
        }
        Ok(())
    }
}
