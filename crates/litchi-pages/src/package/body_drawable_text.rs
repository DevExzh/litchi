//! Selector-first plain-text edits of the storage a body drawable owns.
//!
//! The drawable is selected by kind and text-flow position (today: ordinary
//! body shapes, the only body drawables in this family's scope that own a
//! text storage). One edit stages one UTF-16 splice; the shared TSWP rewrite
//! keeps every style, paragraph and attachment table consistent and every
//! other object byte-identical. A splice that would remove a referenced
//! object (an inline attachment, a hyperlink, a field) is refused.

use std::fmt;
use std::num::NonZeroU64;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::package::ExactArtifacts;
use litchi_iwa_text::{TextPosition, TextSpan};
use litchi_iwa_text_wire::{RewriteBehavior, StorageRewriteExecutionLimits};
use thiserror::Error as ThisError;

use super::Package;
use super::body_drawable_graph::{
    self as graph, Budget, GraphError, GraphLimit, MessageRewrite, ObjectIndex, kind_name,
};
use crate::selector::DrawableSelector;

/// A finite resource charged by one body-drawable text operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Bytes in native payloads inspected by the graph proof.
    InputBytes,
    /// Bytes in the rewritten payload or package artifact.
    OutputBytes,
    /// Parsed wire fields.
    WireFields,
    /// Wire nesting depth.
    WireNesting,
    /// Aggregate wire traversal, text rewrite and reassembly work.
    WireWork,
    /// Semantic section or body-storage cardinality.
    PayloadItems,
    /// Semantic UTF-8 text retained by the reopened document.
    TextBytes,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting",
            Self::WireWork => "wire work",
            Self::PayloadItems => "payload items",
            Self::TextBytes => "text bytes",
        })
    }
}

/// Failure from a body-drawable text read or transaction.
#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies
    /// it (record 0351 §2.4).
    #[error(transparent)]
    Refused(Refusal),
    /// The rooted drawable graph or its text storage is malformed.
    #[error("the selected Pages body-drawable text source is invalid")]
    InvalidSource,
    /// A finite operation resource ceiling was exceeded.
    #[error(
        "Pages body-drawable text {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    Limit {
        /// Resource category.
        kind: LimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded temporary allocation failed.
    #[error("could not allocate {amount} units for Pages body-drawable text")]
    Allocation {
        /// Requested units.
        amount: usize,
    },
    /// Reopening the candidate did not reproduce the requested text.
    #[error("the edited Pages body-drawable text failed semantic verification")]
    Verification,
    /// The patch was produced from another exact package artifact.
    #[error("the Pages body-drawable text patch does not match the exact source package")]
    PatchConflict,
}

/// A well-formed request this family refuses, classified by [`Refusal::kind`].
#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// No body drawable of the selected kind is at the selected position.
    #[error("the Pages body has no {kind} at position {position:?}")]
    NotFound {
        /// The selected kind.
        kind: &'static str,
        /// The selected position.
        position: Position,
    },
    /// The selected kind owns no text storage in this family.
    #[error("a Pages body {kind} has no text storage")]
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
    #[error("this Pages source does not support exact body-drawable text edits")]
    UnsupportedSource,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::NotFound { .. } => RefusalKind::NotFound,
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

impl From<GraphError> for Error {
    fn from(error: GraphError) -> Self {
        match error {
            GraphError::InvalidSource => Self::InvalidSource,
            GraphError::Allocation { amount } => Self::Allocation { amount },
            GraphError::Limit {
                kind,
                observed,
                maximum,
            } => Self::Limit {
                kind: match kind {
                    GraphLimit::InputBytes => LimitKind::InputBytes,
                    GraphLimit::OutputBytes => LimitKind::OutputBytes,
                    GraphLimit::WireFields => LimitKind::WireFields,
                    GraphLimit::WireNesting => LimitKind::WireNesting,
                    GraphLimit::WireWork => LimitKind::WireWork,
                    GraphLimit::SemanticItems => LimitKind::PayloadItems,
                    GraphLimit::SemanticTextBytes => LimitKind::TextBytes,
                },
                observed,
                maximum,
            },
        }
    }
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

    /// Validate and publish the staged splice.
    pub fn commit(self) -> Result<Commit, Error> {
        commit_edit(self)
    }
}

/// Exact-source reversible body-drawable text patch.
#[derive(Clone, PartialEq)]
pub struct Patch {
    artifacts: ExactArtifacts,
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
        self.before == self.after && self.artifacts.is_byte_noop()
    }

    /// The exact target-to-source inverse.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            selector: self.selector,
            before: Arc::clone(&self.after),
            after: Arc::clone(&self.before),
        }
    }
}

/// Content-free diagnostics of one body-drawable text publication.
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

/// The validated immutable result of one body-drawable text transaction.
#[must_use = "a body-drawable text commit contains the validated package snapshot"]
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

#[derive(Debug, Clone, PartialEq, Eq)]
struct Target {
    selector: DrawableSelector,
    drawable: NonZeroU64,
    storage: NonZeroU64,
    storage_message: usize,
    storage_type: u32,
    storage_component: Arc<str>,
}

impl Package {
    /// Read the plain text a body drawable owns (an ordinary shape's text).
    ///
    /// # Errors
    ///
    /// Returns [`Error`] when no drawable matches, the kind owns no text, or
    /// the graph is invalid.
    pub fn body_drawable_text(
        &self,
        selector: impl Into<DrawableSelector>,
    ) -> Result<String, Error> {
        Ok(resolve(self, selector.into())?.1.as_ref().to_owned())
    }

    /// Begin a plain-text edit of the storage a body drawable owns.
    ///
    /// # Errors
    ///
    /// As [`Self::body_drawable_text`].
    pub fn edit_body_drawable_text(
        &self,
        selector: impl Into<DrawableSelector>,
    ) -> Result<Edit<'_>, Error> {
        let (target, before) = resolve(self, selector.into())?;
        Ok(Edit {
            source: self,
            target,
            before,
            operation: None,
        })
    }

    /// Replay an exact-source body-drawable text patch.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when the patch was made against
    /// another artifact or the text no longer matches its `before` state.
    pub fn apply_body_drawable_text(&self, patch: &Patch) -> Result<Commit, Error> {
        if !patch
            .artifacts
            .authorizes_source(&self.state.source.shared_source())
        {
            return Err(Error::PatchConflict);
        }
        let (_, current) = resolve(self, patch.selector)?;
        if current != patch.before {
            return Err(Error::PatchConflict);
        }
        if patch.is_noop() {
            return Ok(Commit {
                package: self.clone(),
                patch: patch.clone(),
                diagnostics: Diagnostics { changed: false },
            });
        }
        let candidate = graph::reopen(self, patch.artifacts.target())?;
        let (_, reopened) = resolve(&candidate, patch.selector)?;
        if reopened != patch.after {
            return Err(Error::Verification);
        }
        Ok(Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics { changed: true },
        })
    }
}

fn resolve(package: &Package, selector: DrawableSelector) -> Result<(Target, Arc<str>), Error> {
    if !matches!(selector, DrawableSelector::Shape(_)) {
        return Err(Refusal::NoText {
            kind: kind_name(selector),
        }
        .into());
    }
    let mut budget = Budget::new(package)?;
    let index = ObjectIndex::new(package, &mut budget)?;
    let body = graph::resolve_body(&index, &mut budget)?;
    let selected =
        graph::select_drawable(&index, &body, selector, &mut budget)?.ok_or(Refusal::NotFound {
            kind: kind_name(selector),
            position: selector.position(),
        })?;
    let storage = selected.storage.ok_or(Error::InvalidSource)?;
    let located = index.get(storage)?;
    let (storage_message, message) = storage_message(located.object)?;
    graph::validate_message_metadata(located.object, storage_message)?;
    let limits = text_limits(package, message.data.len())?;
    let text = litchi_iwa_text_wire::from_bytes(&message.data, limits)
        .map_err(|_| Error::InvalidSource)?;
    let target = Target {
        selector,
        drawable: selected.anchor.drawable,
        storage,
        storage_message,
        storage_type: message.type_,
        storage_component: index.component_name(located.component_index)?,
    };
    Ok((target, Arc::from(text.text())))
}

fn storage_message(
    object: &litchi_iwa_core::ArchiveObject,
) -> Result<(usize, &litchi_iwa_core::RawMessage), Error> {
    let mut found = None;
    for (index, message) in object.messages.iter().enumerate() {
        if matches!(message.type_, 2_001 | 2_022) && found.replace((index, message)).is_some() {
            return Err(Error::InvalidSource);
        }
    }
    found.ok_or(Error::InvalidSource)
}

fn text_limits(package: &Package, payload: usize) -> Result<litchi_iwa_text_wire::Limits, Error> {
    let limits = super::storage_rewrite_limits(package.state.source.limits())
        .map_err(|_| Error::InvalidSource)?;
    litchi_iwa_text_wire::Limits::new(
        limits.max_message_bytes().max(payload).max(1),
        limits.max_fields(),
        limits.max_fragments(),
        limits.max_text_bytes(),
    )
    .map_err(|_| Error::InvalidSource)
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

fn commit_edit(edit: Edit<'_>) -> Result<Commit, Error> {
    let source = edit.source;
    let source_bytes = source.state.source.shared_source();
    let Some((span, replacement)) = edit.operation else {
        return Ok(Commit {
            package: source.clone(),
            patch: Patch {
                artifacts: ExactArtifacts::new(Arc::clone(&source_bytes), source_bytes),
                selector: edit.target.selector,
                before: Arc::clone(&edit.before),
                after: edit.before,
            },
            diagnostics: Diagnostics { changed: false },
        });
    };
    let range = utf16_byte_range(&edit.before, span)?;
    let mut after = String::with_capacity(edit.before.len() + replacement.len());
    after.push_str(&edit.before[..range.start]);
    after.push_str(&replacement);
    after.push_str(&edit.before[range.end..]);
    if after.as_str() == edit.before.as_ref() {
        return Ok(Commit {
            package: source.clone(),
            patch: Patch {
                artifacts: ExactArtifacts::new(Arc::clone(&source_bytes), source_bytes),
                selector: edit.target.selector,
                before: Arc::clone(&edit.before),
                after: edit.before,
            },
            diagnostics: Diagnostics { changed: false },
        });
    }
    if !source.state.source.source_is_exact() {
        return Err(Refusal::UnsupportedSource.into());
    }
    let (current, before) = resolve(source, edit.target.selector)?;
    if current != edit.target || before != edit.before {
        return Err(Error::PatchConflict);
    }
    let object = source
        .state
        .source
        .components()
        .get(current.storage_component.as_ref())
        .and_then(|component| component.archive().object(current.storage.get()))
        .ok_or(Error::InvalidSource)?;
    let payload = object
        .messages
        .get(current.storage_message)
        .ok_or(Error::InvalidSource)?
        .data
        .as_slice();
    let rewrite_limits = super::storage_rewrite_limits(source.state.source.limits())
        .map_err(|_| Error::InvalidSource)?;
    let start = usize::try_from(span.start().utf16_index()).map_err(|_| Refusal::InvalidSpan)?;
    let end = usize::try_from(span.end().utf16_index()).map_err(|_| Refusal::InvalidSpan)?;
    let prepared = litchi_iwa_text_wire::prepare_storage_text_rewrite_with_behavior(
        payload,
        start..end,
        &replacement,
        RewriteBehavior::ReplaceSelection,
        rewrite_limits,
    )
    .map_err(|_| Error::InvalidSource)?;
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
        .map_err(|_| Error::InvalidSource)?;
    if !rewritten.removed_object_references().is_empty()
        || !rewritten.removed_object_references_by_field().is_empty()
        || rewritten.object_reference_occurrences_before()
            != rewritten.object_reference_occurrences_after()
    {
        return Err(Refusal::DependentContent.into());
    }
    let candidate = graph::publish(
        source,
        vec![MessageRewrite {
            component: Arc::clone(&current.storage_component),
            object: current.storage,
            message_index: current.storage_message,
            message_type: current.storage_type,
            data: rewritten.into_bytes(),
        }],
    )?;
    let (reopened_target, reopened) = resolve(&candidate, current.selector)?;
    if reopened_target.drawable != current.drawable || reopened.as_ref() != after {
        return Err(Error::Verification);
    }
    let target_bytes = candidate.state.source.shared_source();
    Ok(Commit {
        package: candidate,
        patch: Patch {
            artifacts: ExactArtifacts::new(source_bytes, target_bytes),
            selector: current.selector,
            before: edit.before,
            after: Arc::from(after),
        },
        diagnostics: Diagnostics { changed: true },
    })
}
