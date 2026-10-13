//! Strict command admission over the shared immutable coordinator.

use super::{BatchItem, BatchMode, Options, Unsupported};
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
    /// The properties this item skipped under [`Unsupported::Warn`], in the
    /// order they were refused: each is the `unsupported_property` error
    /// that refused it, naming the property and, where one exists, a
    /// suggestion. Empty (and unallocated) under [`Unsupported::Refuse`];
    /// boxed so a result stays four words.
    pub warnings: Box<[Error]>,
}
impl fmt::Debug for Applied {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("Applied")
            .field("has_path", &self.path.is_some())
            .field("changed", &self.changed)
            .field("warnings", &self.warnings.len())
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
        self.apply_with(item, &Options::new())
    }

    /// [`Self::apply`] with command-channel options: the unsupported-property
    /// policy and the semantic limits. The batch mode does not apply to a
    /// single command.
    ///
    /// Under [`Unsupported::Warn`], a `set` or `add` whose target refuses a
    /// property with `unsupported_property` is retried without it, until the
    /// command succeeds or fails for another reason; the skipped properties
    /// are reported in [`Applied::warnings`]. The remaining properties apply
    /// atomically: every attempt runs in an isolated candidate, so a failed
    /// attempt leaves nothing behind. A `set` whose every property is
    /// skipped succeeds as a no-op.
    ///
    /// ```
    /// # #[cfg(feature = "xlsx")]
    /// # {
    /// use litchi_core::FileFormat;
    /// use litchi_ops::batch::{Options, Unsupported};
    /// use litchi_ops::{BatchItem, Code, Props, create};
    ///
    /// let file = create(FileFormat::Xlsx)?;
    /// let mut item = BatchItem::new("set");
    /// item.path = Some("/Sheet1/A1".into());
    /// item.props = Some(Props::new().with("value", "Total").with("sparkle", true));
    /// // The library default refuses the whole command.
    /// assert_eq!(file.apply(&item).unwrap_err().code(), Code::UnsupportedProperty);
    /// // OfficeCLI's behaviour applies the value and reports the rest.
    /// let commit = file.apply_with(&item, &Options::new().unsupported(Unsupported::Warn))?;
    /// assert!(commit.diagnostics.changed());
    /// assert_eq!(commit.value.warnings.len(), 1);
    /// assert_eq!(commit.value.warnings[0].property(), Some("sparkle"));
    /// # }
    /// # Ok::<(), litchi_ops::Error>(())
    /// ```
    ///
    /// # Errors
    /// As [`Self::apply`]; under [`Unsupported::Warn`] only for failures
    /// other than an unsupported property.
    pub fn apply_with<'i>(
        &self,
        item: impl Into<BatchItem<'i>>,
        options: &Options,
    ) -> Result<Commit<Applied>> {
        let unsupported = options.unsupported;
        self.edit_with_limits(&options.limits, |edit| {
            apply_item(edit, &item.into(), unsupported, 0).map_err(|error| error.with_item(0))
        })
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
        self.batch_with(items, &Options::new().mode(mode).limits(*limits))
    }

    /// [`Self::batch`] with command-channel options: the batch mode, the
    /// unsupported-property policy and the semantic limits.
    ///
    /// Under [`Unsupported::Warn`] every `set` and `add` item behaves as in
    /// [`Self::apply_with`]: refused properties are skipped and reported in
    /// the item's [`Applied::warnings`] (each carrying the item index), and
    /// the rest of the item applies. Atomic `set` items are then not
    /// coalesced, because a coalesced run reports its refusal only when the
    /// run closes.
    ///
    /// ```
    /// # #[cfg(feature = "xlsx")]
    /// # {
    /// use litchi_core::FileFormat;
    /// use litchi_ops::batch::{self, Options, Unsupported};
    /// use litchi_ops::{ItemResult, create};
    ///
    /// let file = create(FileFormat::Xlsx)?;
    /// let items = batch::from_bytes(br#"[
    ///   {"command":"set","path":"/Sheet1/A1","props":{"value":"Q1","sparkle":"yes"}},
    ///   {"command":"set","path":"/Sheet1/A2","props":{"value":"Q2"}}
    /// ]"#)?;
    /// let options = Options::new().unsupported(Unsupported::Warn);
    /// let commit = file.batch_with(&items, &options)?;
    /// let ItemResult::Applied(first) = &commit.results[0] else { unreachable!() };
    /// assert_eq!(first.warnings[0].item(), Some(0));
    /// # }
    /// # Ok::<(), litchi_ops::Error>(())
    /// ```
    ///
    /// # Errors
    /// As [`Self::batch_with_limits`].
    pub fn batch_with<'i, I>(&self, items: I, options: &Options) -> Result<BatchCommit>
    where
        I: IntoIterator,
        I::Item: Into<BatchItem<'i>>,
    {
        let Options {
            mode,
            unsupported,
            ref limits,
        } = *options;
        // Coalesced runs defer their refusals, so warnings need per-item runs.
        let coalesce = mode == BatchMode::Atomic;
        let coalesce_set = coalesce && unsupported == Unsupported::Refuse;
        let commit = self.edit_with_limits(limits, |edit| {
            let mut results = Vec::new();
            for (index, item) in items.into_iter().enumerate() {
                if index >= limits.max_operations() {
                    // Resolve an earlier deferred native failure before
                    // attributing a later wire-item limit failure.
                    edit.flush_pending()?;
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
                if coalesce_set && item.command.eq_ignore_ascii_case("set") {
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
                            warnings: Box::default(),
                        }));
                        continue;
                    }
                }
                if coalesce && item.command.eq_ignore_ascii_case("remove") {
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
                            warnings: Box::default(),
                        }));
                        continue;
                    }
                }
                edit.flush_pending()?;
                // Metadata changes interpretation of the whole payload. It is
                // never safe to skip a bad declaration and continue replay.
                let meta = item.command.eq_ignore_ascii_case("meta");
                match apply_item(edit, &item, unsupported, index)
                    .map_err(|error| error.with_item(index))
                {
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

fn apply_item(
    edit: &mut Edit,
    item: &BatchItem<'_>,
    unsupported: Unsupported,
    index: usize,
) -> Result<Applied> {
    let warn = unsupported == Unsupported::Warn
        && item.props.as_ref().is_some_and(|props| !props.is_empty())
        && (item.command.eq_ignore_ascii_case("set") || item.command.eq_ignore_ascii_case("add"));
    if warn {
        return apply_skipping(edit, item, index);
    }
    let (path, changed) = edit.isolated(|edit| dispatch(edit, item))?;
    Ok(Applied {
        path: path.map(Box::new),
        changed,
        warnings: Box::default(),
    })
}

/// [`Unsupported::Warn`]: run the item; while its target refuses one of its
/// properties with `unsupported_property`, drop that property (every
/// assignment of it) and run again. Each attempt is isolated, so only the
/// final, successful attempt leaves changes. The loop ends because every
/// retry removes at least one assignment.
fn apply_skipping(edit: &mut Edit, item: &BatchItem<'_>, index: usize) -> Result<Applied> {
    let mut warnings = Vec::new();
    let mut current: Option<BatchItem<'_>> = None;
    loop {
        let attempt = current.as_ref().unwrap_or(item);
        let props = attempt.props.as_ref();
        if attempt.command.eq_ignore_ascii_case("set") && props.is_none_or(Props::is_empty) {
            // OfficeCLI applies nothing and reports the skipped properties.
            return Ok(Applied {
                path: None,
                changed: false,
                warnings: warnings.into_boxed_slice(),
            });
        }
        let error = match edit.isolated(|edit| dispatch(edit, attempt)) {
            Ok((path, changed)) => {
                return Ok(Applied {
                    path: path.map(Box::new),
                    changed,
                    warnings: warnings.into_boxed_slice(),
                });
            },
            Err(error) => error,
        };
        let Some(rest) = props.and_then(|props| without_refused(props, &error)) else {
            return Err(error);
        };
        warnings.try_reserve(1).map_err(|source| {
            Error::from(TransactionError::Allocation { source }).with_item(index)
        })?;
        warnings.push(error.with_item(index));
        let mut next = attempt.clone();
        next.props = Some(rest);
        current = Some(next);
    }
}

/// `props` without the property `error` refused, when `error` is an
/// `unsupported_property` refusal naming one of them.
fn without_refused<'a>(props: &Props<'a>, error: &Error) -> Option<Props<'a>> {
    if error.code() != Code::UnsupportedProperty {
        return None;
    }
    let refused = error.property()?;
    let skipped = |key: &crate::Key<'_>| key.as_str().eq_ignore_ascii_case(refused);
    if !props.iter().any(|(key, _)| skipped(key)) {
        // The refusal names a canonical property the caller spelled with an
        // alias, or something that is not one of the props: never guess.
        return None;
    }
    Some(
        props
            .iter()
            .filter(|(key, _)| !skipped(key))
            .map(|(key, value)| (key.clone(), value.clone()))
            .collect(),
    )
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
