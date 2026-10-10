//! Strict command admission over the shared immutable coordinator.

use super::{BatchItem, BatchMode};
use crate::path::Path;
use crate::transaction::{
    Commit, Diagnostics, Edit, Error as TransactionError, LimitKind, Patch, SemanticLimits,
};
use crate::{Code, Error, File, NewNode, Placement, Props, Result};
use litchi_core::Position;
use std::fmt;

/// The result of one successful command.
#[derive(Clone)]
#[non_exhaustive]
pub struct Applied {
    /// Canonical path created by `add`; absent for other integrated commands.
    /// Subsequent structural operations can change positional paths.
    pub path: Option<Box<Path<'static>>>,
    /// Whether this item changed its input logical document.
    pub changed: bool,
}
impl fmt::Debug for Applied {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("Applied")
            .field("has_path", &self.path.is_some())
            .field("changed", &self.changed)
            .finish()
    }
}

/// One input item's outcome, in original order, including metadata items.
#[derive(Debug, Clone)]
#[non_exhaustive]
pub enum ItemResult {
    /// The item succeeded, possibly as a no-op.
    Applied(Applied),
    /// The item was discarded. Its error carries the original input index.
    Failed(Error),
}

/// An immutable batch result. Publication remains an explicit save operation.
#[non_exhaustive]
pub struct BatchCommit {
    /// The final file; the caller's input remains unchanged.
    pub file: File,
    /// Owner-bound, process-local reversible history.
    pub patch: Patch,
    /// Aggregate facts for successful mutations; metadata is not a mutation.
    pub diagnostics: Diagnostics,
    /// One outcome per input item, with no renumbering after failures or meta.
    pub results: Vec<ItemResult>,
}
impl fmt::Debug for BatchCommit {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("BatchCommit")
            .field("items", &self.results.len())
            .field("diagnostics", &self.diagnostics)
            .finish_non_exhaustive()
    }
}

impl File {
    /// Execute one command, preserving its patch, diagnostics and item result.
    ///
    /// Uses the same strict admission as [`Self::batch`]. Commands include
    /// `set`, `add`, `remove`, `move`, `swap` and v2 `meta`; supported targets
    /// and properties are reported by [`Self::capabilities`].
    ///
    /// # Errors
    /// Returns the command's error with item index zero. The input is unchanged.
    pub fn apply<'i>(&self, item: impl Into<BatchItem<'i>>) -> Result<Commit<Applied>> {
        self.edit(|edit| apply_item(edit, &item.into()).map_err(|error| error.with_item(0)))
    }

    /// Execute ordered command payloads with atomic or explicit best-effort
    /// behavior. Every item resolves against earlier successful results.
    ///
    /// Atomic DOCX batches group consecutive direct-path paragraph/run
    /// formatting sets in one native commit. Text, selectors, metadata and
    /// structural commands close the group. XLSX batches group distinct direct
    /// cell payload assignments; repeated cells, fonts and selectors close the
    /// group. XLSB batches group distinct direct cell payload assignments and
    /// removals; ranges, queries, and repeated cells close the group. PPTX
    /// batches group distinct direct shape text/geometry sets; repeated shapes
    /// and selectors close the group. Best effort isolates every item.
    ///
    /// DOCX command run text sets update the first direct text node literally,
    /// retaining later text and control elements. Cached-result edits mark
    /// the owning same-paragraph complex field dirty without evaluating it.
    /// Paragraph text sets and run
    /// additions normalize CRLF/CR and map LF/vertical tab to inline breaks.
    /// Paragraph additions normalize CRLF/CR, split LF into paragraphs, and
    /// map vertical tab to inline breaks, returning the first created path.
    /// Native [`Self::set`] keeps its native text semantics. Remaining command
    /// policy gaps are reported by [`Self::capabilities`].
    ///
    /// Uses default [`SemanticLimits`]. The operation count also bounds all
    /// input items, including metadata and failures. Results never trust an
    /// iterator's size hint. The input handle is unchanged on every path.
    ///
    /// ```
    /// # #[cfg(feature = "docx")]
    /// # {
    /// use litchi_core::FileFormat;
    /// use litchi_ops::{BatchMode, batch, create};
    /// let file = create(FileFormat::Docx)?;
    /// let items = batch::from_bytes(br#"[
    ///   {"command":"add","parent":"/body","type":"paragraph","props":{"text":"Hello"}},
    ///   {"command":"set","path":"/body/p[1]/r[1]","props":{"bold":true}}
    /// ]"#)?;
    /// let commit = file.batch(items, BatchMode::Atomic)?;
    /// assert_eq!(commit.results.len(), 2);
    /// assert!(commit.diagnostics.changed());
    /// # }
    /// # Ok::<(), litchi_ops::Error>(())
    /// ```
    ///
    /// # Errors
    /// Atomic mode returns the first failed item. Best effort records ordinary
    /// item failures and continues. Protection, source freshness, unsupported
    /// dump semantics, resource-limit and allocation failures abort
    /// the complete batch in either mode.
    pub fn batch<'i, I>(&self, items: I, mode: BatchMode) -> Result<BatchCommit>
    where
        I: IntoIterator,
        I::Item: Into<BatchItem<'i>>,
    {
        self.batch_with_limits(items, mode, &SemanticLimits::default())
    }

    /// [`Self::batch`] with explicit operation, retained-snapshot and XLSX
    /// target-expansion/staging bounds.
    ///
    /// A `meta` item must declare `dumpVersion: 2`. A batch without a `meta`
    /// item also selects current semantics. Other versions refuse the complete batch
    /// rather than allowing earlier text to be interpreted under the wrong
    /// convention. Declarations can occur anywhere and retain their item index.
    ///
    /// # Errors
    /// As [`Self::batch`]. Exceeding the total item bound discards all candidates
    /// even in best-effort mode, so an unbounded failing iterator cannot escape
    /// the resource bound. At most one excess item is requested to detect it.
    pub fn batch_with_limits<'i, I>(
        &self,
        items: I,
        mode: BatchMode,
        limits: &SemanticLimits,
    ) -> Result<BatchCommit>
    where
        I: IntoIterator,
        I::Item: Into<BatchItem<'i>>,
    {
        let commit = self.edit_with_limits(limits, |edit| {
            let mut results = Vec::new();
            for (index, item) in items.into_iter().enumerate() {
                if index >= limits.max_operations() {
                    return Err(Error::from(TransactionError::Limit {
                        kind: LimitKind::Operations,
                        observed: index.saturating_add(1),
                        maximum: limits.max_operations(),
                    })
                    .with_item(index));
                }
                results.try_reserve(1).map_err(|source| {
                    Error::from(TransactionError::Allocation { source }).with_item(index)
                })?;
                let item = item.into();
                if mode == BatchMode::Atomic && item.command.eq_ignore_ascii_case("set") {
                    let (path, props) = match set_fields(&item) {
                        Ok(fields) => fields,
                        Err(error) => {
                            // A previous deferred native failure precedes a
                            // malformed following command in wire order.
                            edit.flush_pending()?;
                            return Err(error.with_item(index));
                        },
                    };
                    if let Some(changed) =
                        edit.coalesce_set(path, props, index).map_err(|error| {
                            // A closing prior run already carries its originating
                            // wire index. Do not reattribute it to the next item.
                            if error.item().is_some() {
                                error
                            } else {
                                error.with_item(index)
                            }
                        })?
                    {
                        results.push(ItemResult::Applied(Applied {
                            path: None,
                            changed,
                        }));
                        continue;
                    }
                }
                if mode == BatchMode::Atomic && item.command.eq_ignore_ascii_case("remove") {
                    let path = match remove_fields(&item) {
                        Ok(path) => path,
                        Err(error) => {
                            // A previous deferred native failure precedes a
                            // malformed following command in wire order.
                            edit.flush_pending()?;
                            return Err(error.with_item(index));
                        },
                    };
                    if let Some(changed) = edit.coalesce_remove(path, index).map_err(|error| {
                        // A closing prior run already carries its
                        // originating wire index. Do not reattribute it
                        // to the next item.
                        if error.item().is_some() {
                            error
                        } else {
                            error.with_item(index)
                        }
                    })? {
                        results.push(ItemResult::Applied(Applied {
                            path: None,
                            changed,
                        }));
                        continue;
                    }
                }
                edit.flush_pending()?;
                // Metadata changes interpretation of the whole payload. It is
                // never safe to skip a bad declaration and continue replay.
                let meta = item.command.eq_ignore_ascii_case("meta");
                match apply_item(edit, &item).map_err(|error| error.with_item(index)) {
                    Ok(applied) => results.push(ItemResult::Applied(applied)),
                    Err(error) if mode == BatchMode::Atomic || meta || fatal(&error) => {
                        return Err(error);
                    },
                    Err(error) => results.push(ItemResult::Failed(error)),
                }
            }
            Ok(results)
        })?;
        Ok(BatchCommit {
            file: commit.file,
            patch: commit.patch,
            diagnostics: commit.diagnostics,
            results: commit.value,
        })
    }
}

fn apply_item(edit: &mut Edit, item: &BatchItem<'_>) -> Result<Applied> {
    let (path, changed) = edit.isolated(|edit| dispatch(edit, item))?;
    Ok(Applied {
        path: path.map(Box::new),
        changed,
    })
}

fn dispatch(edit: &mut Edit, item: &BatchItem<'_>) -> Result<Option<Path<'static>>> {
    item.validate()?;
    let command = item.command.as_ref();
    if command.eq_ignore_ascii_case("meta") {
        fields(item, &["dumpVersion"])?;
        return if item.dump_version == Some(2) {
            Ok(None)
        } else {
            Err(pending(
                "batch dump versions other than explicit v2 are not integrated",
            ))
        };
    }
    if command.eq_ignore_ascii_case("set") {
        let (path, props) = set_fields(item)?;
        edit.set_command(path, props)?;
    } else if command.eq_ignore_ascii_case("add") {
        fields(
            item,
            &[
                "parent", "path", "type", "props", "index", "after", "before", "from",
            ],
        )?;
        if item.from.is_some() && item.kind.is_some() {
            return Err(invalid("add accepts either type or from"));
        }
        if item.from.is_some() {
            return Err(pending("add-from copying is not integrated"));
        }
        let parent = aliases(
            item.parent.as_deref(),
            item.path.as_deref(),
            "add accepts either parent or path",
        )?;
        let parent = required(parent, "add requires parent or path")?;
        let kind = required(item.kind.as_deref(), "add requires type")?;
        let empty = Props::new();
        let props = add_text_props(kind, item.props.as_ref().unwrap_or(&empty))?;
        let node = NewNode::new(kind)
            .with_props(props.as_ref())?
            .at(placement(item, false)?);
        return edit.add_command(parent, node).map(Some);
    } else if command.eq_ignore_ascii_case("remove") {
        edit.remove(remove_fields(item)?)?;
    } else if command.eq_ignore_ascii_case("move") {
        fields(item, &["path", "to", "index", "after", "before", "props"])?;
        if item.props.is_some() {
            return Err(pending("move-time property assignments are not integrated"));
        }
        edit.move_to(
            required(item.path.as_deref(), "move requires path")?,
            placement(item, true)?,
        )?;
    } else if command.eq_ignore_ascii_case("swap") {
        fields(item, &["path", "path2", "to"])?;
        let second = aliases(
            item.path2.as_deref(),
            item.to.as_deref(),
            "swap accepts either path2 or to",
        )?;
        edit.swap(
            required(item.path.as_deref(), "swap requires path")?,
            required(second, "swap requires path2 or to")?,
        )?;
    } else if [
        "get", "query", "view", "raw", "raw-set", "add-part", "validate", "import",
    ]
    .iter()
    .any(|known| command.eq_ignore_ascii_case(known))
    {
        return Err(pending(
            "this command is not integrated into the batch result channel",
        ));
    } else {
        return Err(Error::new(
            Code::UnsupportedCommand,
            "unknown batch command",
        ));
    }
    Ok(None)
}

fn aliases<'a>(
    first: Option<&'a str>,
    second: Option<&'a str>,
    message: &'static str,
) -> Result<Option<&'a str>> {
    if first.is_some() && second.is_some() {
        return Err(invalid(message));
    }
    Ok(first.or(second))
}
fn required<'a>(value: Option<&'a str>, message: &'static str) -> Result<&'a str> {
    value
        .filter(|value| !value.trim().is_empty())
        .ok_or_else(|| missing(message))
}
fn placement<'a>(item: &'a BatchItem<'_>, moving: bool) -> Result<Placement<'a>> {
    let selectors = usize::from(item.index.is_some())
        + usize::from(item.after.is_some())
        + usize::from(item.before.is_some());
    if selectors > 1 {
        return Err(invalid("only one of index, after or before is allowed"));
    }
    let position = item
        .index
        .map(|index| {
            usize::try_from(index)
                .map(Position::new)
                .map_err(|_| invalid("index must be nonnegative"))
        })
        .transpose()?;
    if moving {
        if let Some(parent) = &item.to {
            if item.after.is_some() || item.before.is_some() {
                return Err(pending(
                    "move to a parent combined with a relative anchor is not integrated",
                ));
            }
            return match position {
                Some(position) => Placement::inside_at(parent.as_ref(), position),
                None => Placement::inside(parent.as_ref()),
            };
        }
    }
    if let Some(position) = position {
        return Ok(Placement::Position(position));
    }
    if let Some(path) = &item.after {
        return Placement::after(path.as_ref());
    }
    if let Some(path) = &item.before {
        return Placement::before(path.as_ref());
    }
    Ok(Placement::End)
}
fn fields(item: &BatchItem<'_>, allowed: &[&str]) -> Result<()> {
    for (name, present) in [
        ("path", item.path.is_some()),
        ("path2", item.path2.is_some()),
        ("parent", item.parent.is_some()),
        ("type", item.kind.is_some()),
        ("from", item.from.is_some()),
        ("to", item.to.is_some()),
        ("index", item.index.is_some()),
        ("after", item.after.is_some()),
        ("before", item.before.is_some()),
        ("mode", item.mode.is_some()),
        ("depth", item.depth.is_some()),
        ("part", item.part.is_some()),
        ("xpath", item.xpath.is_some()),
        ("action", item.action.is_some()),
        ("xml", item.xml.is_some()),
        ("props", item.props.is_some()),
        ("selector", item.selector.is_some()),
        ("text", item.text.is_some()),
        ("dumpVersion", item.dump_version.is_some()),
    ] {
        if present && !allowed.contains(&name) {
            return Err(Error::new(
                Code::InvalidArgument,
                format!("this batch command does not accept the {name} field"),
            ));
        }
    }
    Ok(())
}
fn missing(message: &'static str) -> Error {
    Error::new(Code::MissingArgument, message)
}
fn invalid(message: &'static str) -> Error {
    Error::new(Code::InvalidArgument, message)
}
fn pending(message: &'static str) -> Error {
    Error::new(Code::UnsupportedCapability, message)
        .with_suggestion("see capabilities and the M2 batch command follow-ups")
}

fn fatal(error: &Error) -> bool {
    matches!(
        error.code(),
        Code::IoError
            | Code::Cancelled
            | Code::LimitExceeded
            | Code::InternalError
            | Code::DocumentProtected
    )
}

// The current native run setter maps LF to a soft break, while OfficeCLI v2
// has operation-specific control semantics. Run addition and paragraph setters
// use inline normalization; command run setters use the literal first text node.
fn add_text_props<'a>(kind: &str, props: &'a Props<'a>) -> Result<std::borrow::Cow<'a, Props<'a>>> {
    if !kind.eq_ignore_ascii_case("r") && !kind.eq_ignore_ascii_case("run") {
        if !kind.eq_ignore_ascii_case("p") && !kind.eq_ignore_ascii_case("paragraph") {
            text_semantics(props)?;
        }
        return Ok(std::borrow::Cow::Borrowed(props));
    }
    Ok(inline_text_props(props))
}

pub(crate) fn inline_text_props<'a>(props: &'a Props<'a>) -> std::borrow::Cow<'a, Props<'a>> {
    if !props.iter().any(|(key, value)| {
        key.eq_ignore_ascii_case("text")
            && value
                .as_text()
                .is_some_and(|text| text.contains(['\r', '\u{b}']))
    }) {
        return std::borrow::Cow::Borrowed(props);
    }
    let mut normalized = Props::new();
    for (key, value) in props.iter() {
        let value = if key.eq_ignore_ascii_case("text") {
            match value.as_text() {
                Some(text) if text.contains(['\r', '\u{b}']) => {
                    // Only authored command input is normalized. Never infer
                    // controls by reverse-mapping flattened document text.
                    let mut output = String::with_capacity(text.len());
                    let mut chars = text.chars().peekable();
                    while let Some(character) = chars.next() {
                        match character {
                            '\r' => {
                                if chars.peek() == Some(&'\n') {
                                    chars.next();
                                }
                                output.push('\n');
                            },
                            '\u{b}' => output.push('\n'),
                            other => output.push(other),
                        }
                    }
                    crate::PropValue::from(output)
                },
                _ => value.clone(),
            }
        } else {
            value.clone()
        };
        normalized.push(key.clone(), value);
    }
    std::borrow::Cow::Owned(normalized)
}

pub(crate) fn has_text_controls(props: &Props<'_>) -> bool {
    props.iter().any(|(key, value)| {
        key.eq_ignore_ascii_case("text")
            && value
                .as_text()
                .is_some_and(|text| text.contains(['\n', '\r', '\u{b}']))
    })
}

fn text_semantics(props: &Props<'_>) -> Result<()> {
    if has_text_controls(props) {
        return Err(pending(
            "OfficeCLI text controls are not integrated for this command target",
        ));
    }
    Ok(())
}

fn set_fields<'a>(item: &'a BatchItem<'_>) -> Result<(&'a str, &'a Props<'a>)> {
    fields(item, &["path", "props"])?;
    let path = required(item.path.as_deref(), "set requires path")?;
    let props = item
        .props
        .as_ref()
        .ok_or_else(|| missing("set requires props"))?;
    if props.is_empty() {
        return Err(invalid("set requires at least one property"));
    }
    Ok((path, props))
}

fn remove_fields<'a>(item: &'a BatchItem<'_>) -> Result<&'a str> {
    fields(item, &["path"])?;
    required(item.path.as_deref(), "remove requires path")
}
