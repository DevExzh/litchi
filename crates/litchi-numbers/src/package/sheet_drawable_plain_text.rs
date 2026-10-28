//! Selector-first plain-text edits of the storage a sheet text box or shape
//! owns (record 0378).
//!
//! One edit stages one UTF-16 splice; the shared TSWP rewrite
//! ([`litchi_iwa_text_wire::prepare_storage_text_rewrite_with_behavior`])
//! keeps every style, paragraph and attachment table consistent and every
//! other object byte-identical. A splice that would remove a referenced
//! object (an inline attachment, a hyperlink, a field) is refused. This is
//! the behaviour of the host editor's `{set,replace,clear}_sheet_text_box_text`
//! and `{set,replace,clear}_sheet_shape_text` routes.

use std::fmt;
use std::sync::Arc;

use litchi_iwa_core::RawMessage;
use litchi_iwa_text::{TextPosition, TextSpan};
use litchi_iwa_text_wire::{RewriteBehavior, RewriteLimits, StorageRewriteExecutionLimits};
use thiserror::Error as ThisError;

use super::Package;
use super::sheet_drawable_graph::{self as graph, SHAPE_INFO_MESSAGE_TYPE};
use super::structure::{self, Artifacts, Fault, FaultLimit, references};
use crate::selector::SheetSelector;
use crate::sheet::drawable::{DrawableSelector, Kind};

/// `TSWP.ShapeInfoArchive.owned_storage`.
const OWNED_STORAGE_FIELD: u32 = 4;
/// `TSWP.StorageArchive`.
const STORAGE_MESSAGE_TYPE: u32 = 2_001;

/// A finite resource charged by one plain-text transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Complete source or candidate package bytes.
    PackageBytes,
    /// Bytes in one decoded component.
    PayloadBytes,
    /// Native object identifiers.
    Identifiers,
    /// Temporary graph scratch bytes.
    ScratchBytes,
    /// Fallible graph allocation events.
    AllocationEvents,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::PackageBytes => "package bytes",
            Self::PayloadBytes => "payload bytes",
            Self::Identifiers => "object identifiers",
            Self::ScratchBytes => "scratch bytes",
            Self::AllocationEvents => "allocation events",
        })
    }
}

/// A typed, content-redacted plain-text failure.
#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies
    /// it (record 0351 §2.4).
    #[error(transparent)]
    Refused(Refusal),
    /// The rooted drawable graph or its text storage is malformed.
    #[error("the selected Numbers drawable text source is invalid")]
    InvalidSource,
    /// A finite transaction limit was exceeded.
    #[error("{kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// Resource category.
        kind: LimitKind,
        /// Observed amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded allocation failed.
    #[error("could not allocate {amount} drawable-text items or bytes")]
    Allocation {
        /// Requested amount.
        amount: usize,
    },
    /// Reopening the candidate did not reproduce the requested text.
    #[error("the edited Numbers drawable text failed verification")]
    Verification,
    /// The patch was produced from another exact package artifact.
    #[error("the Numbers drawable text patch does not match this package")]
    PatchConflict,
}

/// A well-formed request this family refuses, classified by [`Refusal::kind`].
#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// The sheet selector did not resolve one sheet.
    #[error("sheet selector did not resolve")]
    SheetNotFound,
    /// No drawable matches the selector.
    #[error("the sheet has no drawable at {selector:?}")]
    DrawableNotFound {
        /// The selector that did not resolve.
        selector: DrawableSelector,
    },
    /// The drawable owns no text (it is not a text box or shape).
    #[error("a Numbers sheet {kind} owns no text")]
    NoText {
        /// The selected kind.
        kind: &'static str,
    },
    /// The span is outside the text or splits a UTF-16 surrogate pair.
    #[error("the text span is outside the drawable text or splits a character")]
    InvalidSpan,
    /// The replacement contains an object-replacement, footnote or
    /// section-break marker, which only object families may write.
    #[error("the replacement contains a reserved native text marker")]
    ReservedMarker,
    /// The span would remove text that references an object.
    #[error("the span removes text that references a native object")]
    DependentContent,
    /// A second splice was staged on one edit.
    #[error("one drawable text edit stages one splice")]
    OperationAlreadyStaged,
    /// The source does not keep an exact artifact for a changed publication.
    #[error("this Numbers source does not support exact drawable text edits")]
    UnsupportedSource,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::SheetNotFound | Self::DrawableNotFound { .. } => RefusalKind::NotFound,
            Self::NoText { .. } => RefusalKind::UnsupportedTarget,
            Self::InvalidSpan | Self::ReservedMarker => RefusalKind::InvalidValue,
            Self::DependentContent | Self::UnsupportedSource => RefusalKind::UnsupportedContent,
            Self::OperationAlreadyStaged => RefusalKind::UnsupportedOperation,
        }
    }
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

impl From<Fault> for Error {
    fn from(fault: Fault) -> Self {
        match fault {
            Fault::Unsupported | Fault::Dependency => Self::Refused(Refusal::UnsupportedSource),
            Fault::Invalid => Self::InvalidSource,
            Fault::Limit {
                kind,
                observed,
                maximum,
            } => Self::Limit {
                kind: match kind {
                    FaultLimit::PackageBytes => LimitKind::PackageBytes,
                    FaultLimit::PayloadBytes => LimitKind::PayloadBytes,
                    FaultLimit::Identifiers => LimitKind::Identifiers,
                },
                observed,
                maximum,
            },
            Fault::Allocation { amount } => Self::Allocation { amount },
            Fault::Verification => Self::Verification,
        }
    }
}

/// The private target of one edit.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct Target {
    sheet: usize,
    selector: DrawableSelector,
    storage: u64,
    index: usize,
}

/// One plain-text splice staged against an immutable package.
pub struct Edit<'a> {
    source: &'a Package,
    target: Target,
    before: Arc<str>,
    operation: Option<(TextSpan, Box<str>)>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("selector", &self.target.selector)
            .field("span", &self.operation.as_ref().map(|(span, _)| *span))
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    /// The drawable's text when the edit began.
    #[must_use]
    pub fn text(&self) -> &str {
        &self.before
    }

    /// Stage one replacement of a UTF-16 span; an empty span inserts.
    ///
    /// # Errors
    ///
    /// Refuses a span outside the text or splitting a character, a reserved
    /// marker in the replacement, and a second splice.
    pub fn replace(&mut self, span: TextSpan, replacement: &str) -> Result<&mut Self, Error> {
        if self.operation.is_some() {
            return Err(Refusal::OperationAlreadyStaged.into());
        }
        if replacement
            .chars()
            .any(|character| matches!(character, '\u{fffc}' | '\u{000e}' | '\u{0004}'))
        {
            return Err(Refusal::ReservedMarker.into());
        }
        utf16_byte_range(&self.before, span)?;
        self.operation = Some((span, Box::from(replacement)));
        Ok(self)
    }

    /// Stage one insertion at a UTF-16 position.
    ///
    /// # Errors
    ///
    /// As [`Self::replace`].
    pub fn insert(&mut self, position: TextPosition, text: &str) -> Result<&mut Self, Error> {
        let span = TextSpan::new(position, position).map_err(|_| Refusal::InvalidSpan)?;
        self.replace(span, text)
    }

    /// Stage deletion of a UTF-16 span.
    ///
    /// # Errors
    ///
    /// As [`Self::replace`].
    pub fn delete(&mut self, span: TextSpan) -> Result<&mut Self, Error> {
        self.replace(span, "")
    }

    /// Stage replacement of the whole text.
    ///
    /// # Errors
    ///
    /// As [`Self::replace`].
    pub fn set(&mut self, text: &str) -> Result<&mut Self, Error> {
        let end = TextPosition::from_utf16_index(self.before.encode_utf16().count())
            .map_err(|_| Refusal::InvalidSpan)?;
        let span = TextSpan::new(TextPosition::ZERO, end).map_err(|_| Refusal::InvalidSpan)?;
        self.replace(span, text)
    }

    /// Stage removal of the whole text; the drawable stays.
    ///
    /// # Errors
    ///
    /// As [`Self::replace`].
    pub fn clear(&mut self) -> Result<&mut Self, Error> {
        self.set("")
    }

    /// Validate and publish the staged splice. An edit without a splice, or
    /// whose splice leaves the text as it is, is an exact no-op.
    ///
    /// # Errors
    ///
    /// [`Refusal::DependentContent`], source, resource and verification
    /// errors; failure never publishes a partial package.
    pub fn commit(self) -> Result<Commit, Error> {
        commit_edit(self)
    }
}

/// Exact-source reversible drawable-text patch.
#[derive(Clone, PartialEq)]
pub struct Patch {
    artifacts: Artifacts,
    sheet: usize,
    selector: DrawableSelector,
    before: Arc<str>,
    after: Arc<str>,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("selector", &self.selector)
            .finish_non_exhaustive()
    }
}

impl Patch {
    /// The selected drawable.
    #[must_use]
    pub const fn selector(&self) -> DrawableSelector {
        self.selector
    }

    /// The text required before this patch applies.
    #[must_use]
    pub fn before(&self) -> &str {
        &self.before
    }

    /// The text this patch produces.
    #[must_use]
    pub fn after(&self) -> &str {
        &self.after
    }

    /// Whether both the text and the exact artifact are unchanged.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after && self.artifacts.is_noop()
    }

    /// The exact target-to-source inverse.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            sheet: self.sheet,
            selector: self.selector,
            before: Arc::clone(&self.after),
            after: Arc::clone(&self.before),
        }
    }
}

/// Content-free diagnostics of one drawable-text publication.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Diagnostics {
    changed: bool,
}

impl Diagnostics {
    /// Whether package bytes and text changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }
}

/// The validated immutable result of one drawable-text transaction.
#[must_use = "a drawable-text commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    /// Borrow the reopened package.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume the commit and return its package.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Consume the commit and return its package and patch.
    #[must_use]
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }

    /// Borrow the exact-source reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Borrow the publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }
}

impl Package {
    /// Read the plain text a sheet text box or shape owns.
    ///
    /// # Errors
    ///
    /// [`Refusal::SheetNotFound`], [`Refusal::DrawableNotFound`],
    /// [`Refusal::NoText`] and source errors.
    pub fn sheet_drawable_plain_text<'sheet>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        drawable: impl Into<DrawableSelector>,
    ) -> Result<String, Error> {
        let sheet = resolve_sheet(self, sheet.into())?;
        Ok(resolve(self, sheet, drawable.into())?.1.as_ref().to_owned())
    }

    /// Begin a plain-text edit of the storage a sheet text box or shape
    /// owns.
    ///
    /// ```no_run
    /// use litchi_numbers::Package;
    /// use litchi_numbers::sheet::drawable::DrawableSelector;
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("input.numbers")?;
    /// let mut edit = package.edit_sheet_drawable_plain_text(0, DrawableSelector::text_box(0))?;
    /// edit.set("Updated")?;
    /// let commit = edit.commit()?;
    /// # let _ = commit;
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Errors
    ///
    /// As [`Self::sheet_drawable_plain_text`].
    pub fn edit_sheet_drawable_plain_text<'sheet>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        drawable: impl Into<DrawableSelector>,
    ) -> Result<Edit<'_>, Error> {
        let sheet = resolve_sheet(self, sheet.into())?;
        let (target, before) = resolve(self, sheet, drawable.into())?;
        Ok(Edit {
            source: self,
            target,
            before,
            operation: None,
        })
    }

    /// Replay an exact-source drawable-text patch.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] when the patch was made against another
    /// artifact or the text no longer matches its `before` state.
    pub fn apply_sheet_drawable_plain_text(&self, patch: &Patch) -> Result<Commit, Error> {
        let (_, current) = resolve(self, patch.sheet, patch.selector)?;
        if current != patch.before {
            return Err(Error::PatchConflict);
        }
        let package = patch.artifacts.apply(self)?.ok_or(Error::PatchConflict)?;
        let changed = !patch.is_noop();
        if changed {
            let (_, reopened) = resolve(&package, patch.sheet, patch.selector)?;
            if reopened != patch.after {
                return Err(Error::Verification);
            }
        }
        Ok(Commit {
            package,
            patch: patch.clone(),
            diagnostics: Diagnostics { changed },
        })
    }
}

fn resolve_sheet(package: &Package, selector: SheetSelector<'_>) -> Result<usize, Error> {
    graph::sheet_position(package, selector)?.ok_or_else(|| Refusal::SheetNotFound.into())
}

fn resolve(
    package: &Package,
    sheet: usize,
    selector: DrawableSelector,
) -> Result<(Target, Arc<str>), Error> {
    let drawables = graph::sheet_drawables(package, sheet)?;
    let entry = drawables
        .select(selector)
        .ok_or(Refusal::DrawableNotFound { selector })?;
    if !matches!(entry.kind, Kind::TextBox | Kind::Shape) {
        return Err(Refusal::NoText {
            kind: entry.kind.name(),
        }
        .into());
    }
    let (index, message_type) = entry.message.ok_or(Error::InvalidSource)?;
    if message_type != SHAPE_INFO_MESSAGE_TYPE {
        return Err(Error::InvalidSource);
    }
    let object = graph::object(package, entry.identifier)?;
    let records = references::records(&object.messages[index].data, &[OWNED_STORAGE_FIELD])?;
    let [record] = records.as_slice() else {
        return Err(Error::InvalidSource);
    };
    let storage = references::reference_identifier(record)?;
    let storage_object = graph::object(package, storage)?;
    let storage_index = structure::unique_message(storage_object, &[STORAGE_MESSAGE_TYPE])?;
    let text = litchi_iwa_text_wire::from_bytes(
        &storage_object.messages[storage_index].data,
        litchi_iwa_text_wire::Limits::default(),
    )
    .map_err(|_error| Error::InvalidSource)?;
    Ok((
        Target {
            sheet,
            selector,
            storage,
            index: storage_index,
        },
        Arc::from(text.text()),
    ))
}

fn utf16_byte_range(text: &str, span: TextSpan) -> Result<std::ops::Range<usize>, Error> {
    let start = usize::try_from(span.start().utf16_index()).map_err(|_| Refusal::InvalidSpan)?;
    let end = usize::try_from(span.end().utf16_index()).map_err(|_| Refusal::InvalidSpan)?;
    let mut units = 0usize;
    let (mut start_byte, mut end_byte) = (None, None);
    for (byte, character) in text.char_indices() {
        if units == start {
            start_byte = Some(byte);
        }
        if units == end {
            end_byte = Some(byte);
        }
        units += character.len_utf16();
    }
    if units == start {
        start_byte = Some(text.len());
    }
    if units == end {
        end_byte = Some(text.len());
    }
    match (start_byte, end_byte) {
        (Some(start), Some(end)) if start <= end => Ok(start..end),
        _ => Err(Refusal::InvalidSpan.into()),
    }
}

fn unchanged(edit: Edit<'_>) -> Result<Commit, Error> {
    Ok(Commit {
        package: edit.source.clone(),
        patch: Patch {
            artifacts: Artifacts::unchanged(edit.source)?,
            sheet: edit.target.sheet,
            selector: edit.target.selector,
            before: Arc::clone(&edit.before),
            after: edit.before,
        },
        diagnostics: Diagnostics { changed: false },
    })
}

fn commit_edit(edit: Edit<'_>) -> Result<Commit, Error> {
    let source = edit.source;
    let Some((span, replacement)) = edit.operation.clone() else {
        return unchanged(edit);
    };
    let range = utf16_byte_range(&edit.before, span)?;
    let mut after = String::new();
    after
        .try_reserve(edit.before.len() + replacement.len())
        .map_err(|_error| Error::Allocation {
            amount: edit.before.len() + replacement.len(),
        })?;
    after.push_str(&edit.before[..range.start]);
    after.push_str(&replacement);
    after.push_str(&edit.before[range.end..]);
    if after.as_str() == edit.before.as_ref() {
        return unchanged(edit);
    }
    let target = edit.target;
    let object = graph::object(source, target.storage)?;
    let payload = object
        .messages
        .get(target.index)
        .ok_or(Error::InvalidSource)?
        .data
        .as_slice();
    let start = usize::try_from(span.start().utf16_index()).map_err(|_| Refusal::InvalidSpan)?;
    let end = usize::try_from(span.end().utf16_index()).map_err(|_| Refusal::InvalidSpan)?;
    let prepared = litchi_iwa_text_wire::prepare_storage_text_rewrite_with_behavior(
        payload,
        start..end,
        &replacement,
        RewriteBehavior::ReplaceSelection,
        RewriteLimits::default(),
    )
    .map_err(|_error| Error::InvalidSource)?;
    let requirements = prepared.execution_requirements();
    let rewritten = prepared
        .execute(StorageRewriteExecutionLimits {
            max_output_bytes: requirements.output_bytes(),
            max_retained_elements: requirements.retained_elements(),
            max_retained_bytes: requirements.retained_bytes(),
            max_peak_scratch_bytes: requirements.peak_scratch_bytes(),
            max_allocations: requirements.allocations(),
            max_work: requirements.work(),
        })
        .map_err(|_error| Error::InvalidSource)?;
    if !rewritten.removed_object_references().is_empty()
        || !rewritten.removed_object_references_by_field().is_empty()
        || rewritten.object_reference_occurrences_before()
            != rewritten.object_reference_occurrences_after()
    {
        return Err(Refusal::DependentContent.into());
    }
    let mut data = Some(rewritten.into_bytes());
    let mut replace = |object: &mut litchi_iwa_core::ArchiveObject,
                       limits: litchi_iwa_core::Limits| {
        object
            .replace_message_preserving_header(
                target.index,
                RawMessage {
                    type_: STORAGE_MESSAGE_TYPE,
                    data: data.take().ok_or(Fault::Invalid)?,
                },
                limits,
            )
            .map(drop)
            .map_err(|error| structure::map_core(&error))
    };
    let published = graph::publish_in_place(source, &mut [(target.storage, &mut replace)])?;
    let (_, reopened) = resolve(&published.package, target.sheet, target.selector)?;
    if reopened.as_ref() != after {
        return Err(Error::Verification);
    }
    Ok(Commit {
        patch: Patch {
            artifacts: Artifacts::new(source, &published.package)?,
            sheet: target.sheet,
            selector: target.selector,
            before: edit.before,
            after: Arc::from(after),
        },
        diagnostics: Diagnostics { changed: true },
        package: published.package,
    })
}
