//! Exact-source, selector-first Keynote title and body text transactions.

#![allow(
    clippy::map_err_ignore,
    clippy::needless_pass_by_value,
    clippy::shadow_reuse,
    clippy::wildcard_enum_match_arm,
    reason = "The transaction redacts lower-layer failures and maps non-exhaustive cross-crate error families into a content-free format-owned boundary."
)]

use std::fmt;
use std::ops::Range;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::{SourceCatalog, package::EntryEdit};
use litchi_iwa_common::{
    WireLimits, decode_varint_from_bytes, varint::encoded_len, wire::WireView,
};
use litchi_iwa_core::{Archive, RawMessage, SnappyStream};
use litchi_iwa_protos::{keynote_placeholder_text_codec, keynote_speaker_notes_codec};
use litchi_iwa_text::{TextPosition, TextSpan};
use litchi_iwa_text_wire::{RewriteBehavior, RewriteLimits};
use thiserror::Error;

use super::{
    NOTE_MESSAGE_TYPE, PLACEHOLDER_MESSAGE_TYPE, Package, PhysicalSource, ReadError,
    SHAPE_INFO_MESSAGE_TYPE, SLIDE_MESSAGE_TYPE, STORAGE_MESSAGE_TYPE, SemanticBudget,
    SemanticLimitKind, SemanticPath, unique_payload,
};
use crate::SlideSelector;

const PREVIEW_ENTRY_NAMES: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];

/// The semantic placeholder text owned by a Keynote slide.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SlideTextRole {
    /// The slide's title placeholder.
    Title,
    /// The slide's body placeholder.
    Body,
}

impl fmt::Display for SlideTextRole {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::Title => "title",
            Self::Body => "body",
        })
    }
}

/// A finite resource governed while slide text is read or rewritten.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SlideTextLimitKind {
    /// Complete input package bytes.
    InputBytes,
    /// Complete rewritten package bytes.
    OutputBytes,
    /// ZIP members, IWA objects, messages, or text-table entries.
    Entries,
    /// Bytes in one package member, IWA object, or message.
    EntryBytes,
    /// Aggregate package or IWA bytes.
    TotalBytes,
    /// Semantic slides.
    Slides,
    /// Semantic graph references.
    References,
    /// Semantic text-storage objects.
    TextStorages,
    /// Semantic text fragments.
    TextFragments,
    /// Aggregate semantic text bytes.
    TextBytes,
    /// UTF-16 code units in one native text storage.
    TextUnits,
    /// Bytes in one protobuf payload.
    WireBytes,
    /// Parsed protobuf fields.
    WireFields,
    /// Protobuf nesting depth.
    WireNesting,
    /// Aggregate protobuf scan and rewrite work.
    WireWork,
}

impl fmt::Display for SlideTextLimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::Entries => "entries",
            Self::EntryBytes => "entry bytes",
            Self::TotalBytes => "total bytes",
            Self::Slides => "slides",
            Self::References => "references",
            Self::TextStorages => "text storages",
            Self::TextFragments => "text fragments",
            Self::TextBytes => "text bytes",
            Self::TextUnits => "text UTF-16 units",
            Self::WireBytes => "wire bytes",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting depth",
            Self::WireWork => "wire work",
        })
    }
}

/// A content-redacted failure raised by a slide-text transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum SlideTextError {
    /// The source was prepared without an exact physical package artifact.
    #[error("this Keynote source does not support physical slide-text edits")]
    UnsupportedSource,
    /// An exact-name selector was ambiguous.
    #[error("the Keynote slide selector is ambiguous")]
    AmbiguousSelector,
    /// An exact-name selector did not match.
    #[error("the Keynote show has no slide matching the requested name")]
    SlideNameNotFound,
    /// A checked semantic source position does not exist.
    #[error("the Keynote show has no slide at position {position:?}")]
    SlidePositionNotFound {
        /// Missing semantic source position.
        position: Position,
    },
    /// The selected slide has no existing placeholder for this role.
    #[error("the selected Keynote slide has no existing {role} placeholder")]
    TextStorageNotFound {
        /// Semantic placeholder role that is absent.
        role: SlideTextRole,
    },
    /// A staged span exceeds the selected storage's UTF-16 length.
    #[error("the Keynote slide-text span exceeds the selected text length")]
    SpanOutOfBounds {
        /// Rejected storage-relative span.
        span: TextSpan,
        /// Storage length in UTF-16 code units.
        length: TextPosition,
    },
    /// A staged boundary splits a UTF-16 surrogate pair.
    #[error("the Keynote slide-text boundary {position:?} splits a Unicode scalar value")]
    SurrogateBoundary {
        /// Rejected storage-relative boundary.
        position: TextPosition,
    },
    /// Replacement text contains an inline-object replacement character.
    #[error("Keynote inline-object markers cannot be inserted through a slide-text transaction")]
    ObjectMarkerReplacement,
    /// The edit would consume content owned by another semantic capability.
    #[error("the Keynote slide-text edit intersects dependent content")]
    DependentContent,
    /// A second operation was staged on an edit that already has one.
    #[error("a Keynote slide-text edit accepts exactly one staged operation")]
    OperationAlreadyStaged,
    /// The source graph or selected storage cannot be decoded or rewritten safely.
    #[error("the Keynote slide-text source cannot be edited safely")]
    InvalidSource,
    /// A finite transaction resource ceiling was exceeded.
    #[error("Keynote slide-text {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    LimitExceeded {
        /// Resource category that exceeded its limit.
        kind: SlideTextLimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded destination allocation failed before publication.
    #[error("could not allocate {amount} units for the Keynote slide-text transaction")]
    Allocation {
        /// Elements or bytes requested.
        amount: usize,
    },
    /// Full semantic readback did not reproduce the requested change.
    #[error("the edited Keynote slide text failed semantic verification")]
    Verification,
    /// The patch does not belong to this exact immutable package artifact.
    #[error("the Keynote slide-text patch does not match the exact source package")]
    PatchConflict,
}

#[derive(Debug)]
struct Operation {
    span: TextSpan,
    replacement: Option<String>,
}

#[derive(Debug)]
struct TextSnapshot {
    placeholder_identifier: u64,
    storage_identifier: u64,
    text: String,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct ReferenceSemantics {
    identifier: u64,
    deprecated_type: Option<i32>,
    deprecated_is_external: Option<bool>,
}

struct WireScanBudget {
    limits: WireLimits,
    fields: usize,
    work: usize,
}

impl WireScanBudget {
    fn new(package: &Package) -> Result<Self, SlideTextError> {
        Ok(Self {
            limits: package.wire_limits().map_err(map_wire_error)?,
            fields: 0,
            work: 0,
        })
    }

    fn parse<'source>(
        &mut self,
        payload: &'source [u8],
    ) -> Result<WireView<'source>, SlideTextError> {
        let observed_work = self.work.saturating_add(payload.len());
        if observed_work > self.limits.max_rewrite_work() {
            return Err(SlideTextError::LimitExceeded {
                kind: SlideTextLimitKind::WireWork,
                observed: usize_to_u64(observed_work),
                maximum: usize_to_u64(self.limits.max_rewrite_work()),
            });
        }
        self.work = observed_work;
        let view = WireView::parse_with_limits(payload, self.limits).map_err(map_wire_error)?;
        let observed_fields = self.fields.saturating_add(view.len());
        if observed_fields > self.limits.max_fields() {
            return Err(SlideTextError::LimitExceeded {
                kind: SlideTextLimitKind::WireFields,
                observed: usize_to_u64(observed_fields),
                maximum: usize_to_u64(self.limits.max_fields()),
            });
        }
        self.fields = observed_fields;
        Ok(view)
    }
}

impl ReferenceSemantics {
    const fn from_placeholder_codec(
        reference: keynote_placeholder_text_codec::ReferenceSnapshot,
    ) -> Self {
        Self {
            identifier: reference.identifier().get(),
            deprecated_type: reference.deprecated_type(),
            deprecated_is_external: reference.deprecated_is_external(),
        }
    }
}

/// One mutable storage-relative text operation staged against an immutable package.
pub struct SlideTextEdit<'a> {
    source: &'a Package,
    position: Position,
    role: SlideTextRole,
    placeholder_identifier: u64,
    storage_identifier: u64,
    before: String,
    operation: Option<Operation>,
}

impl fmt::Debug for SlideTextEdit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SlideTextEdit")
            .field("position", &self.position)
            .field("role", &self.role)
            .field("has_operation", &self.operation.is_some())
            .finish_non_exhaustive()
    }
}

impl<'a> SlideTextEdit<'a> {
    fn new<'selector>(
        source: &'a Package,
        selector: impl Into<SlideSelector<'selector>>,
        role: SlideTextRole,
    ) -> Result<Self, SlideTextError> {
        let position = resolve_position(source, selector.into())?;
        let snapshot = text_snapshot_at(source, position, role)?
            .ok_or(SlideTextError::TextStorageNotFound { role })?;
        Ok(Self {
            source,
            position,
            role,
            placeholder_identifier: snapshot.placeholder_identifier,
            storage_identifier: snapshot.storage_identifier,
            before: snapshot.text,
            operation: None,
        })
    }

    /// Return the semantic slide position resolved when this edit began.
    #[must_use]
    pub const fn position(&self) -> Position {
        self.position
    }

    /// Return the semantic placeholder role selected when this edit began.
    #[must_use]
    pub const fn role(&self) -> SlideTextRole {
        self.role
    }

    /// Borrow the selected placeholder text from the immutable source snapshot.
    #[must_use]
    pub fn text(&self) -> &str {
        &self.before
    }

    /// Return the staged storage-relative span, when present.
    #[must_use]
    pub fn span(&self) -> Option<TextSpan> {
        self.operation.as_ref().map(|operation| operation.span)
    }

    /// Stage one replacement over a checked storage-relative UTF-16 span.
    ///
    /// Empty spans insert text and an empty replacement deletes a nonempty
    /// span. Boundaries must align with Unicode scalar values. Inline-object
    /// markers cannot be introduced or consumed by this plain-text capability.
    ///
    /// # Costs
    ///
    /// Scans the selected text and copies a non-no-op replacement into bounded
    /// transaction storage.
    ///
    /// # Errors
    ///
    /// Returns a typed span, dependency, limit, allocation, or staging error
    /// without modifying the source snapshot.
    pub fn replace(
        &mut self,
        span: TextSpan,
        replacement: &str,
    ) -> Result<&mut Self, SlideTextError> {
        if self.operation.is_some() {
            return Err(SlideTextError::OperationAlreadyStaged);
        }
        validate_replacement(replacement)?;
        let range = validate_span(&self.before, span)?;
        let selected = self
            .before
            .get(range)
            .ok_or(SlideTextError::InvalidSource)?;
        validate_consumed_text(selected)?;
        let is_semantic_noop = selected == replacement;
        validate_candidate_text_memory(self.source, &self.before, span, replacement)?;
        self.operation = Some(Operation {
            span,
            replacement: if is_semantic_noop {
                None
            } else {
                Some(try_owned_text(replacement)?)
            },
        });
        Ok(self)
    }

    /// Stage insertion at a checked storage-relative UTF-16 position.
    ///
    /// # Costs
    ///
    /// Has the same costs as [`Self::replace`].
    ///
    /// # Errors
    ///
    /// Returns the same failures as [`Self::replace`].
    pub fn insert(
        &mut self,
        position: TextPosition,
        text: &str,
    ) -> Result<&mut Self, SlideTextError> {
        let span =
            TextSpan::new(position, position).map_err(|_error| SlideTextError::InvalidSource)?;
        self.replace(span, text)
    }

    /// Stage deletion of one checked storage-relative UTF-16 span.
    ///
    /// # Costs
    ///
    /// Scans the selected text to validate UTF-16 boundaries.
    ///
    /// # Errors
    ///
    /// Returns the same failures as [`Self::replace`].
    pub fn delete(&mut self, span: TextSpan) -> Result<&mut Self, SlideTextError> {
        self.replace(span, "")
    }

    /// Stage replacement of all selected placeholder text.
    ///
    /// # Costs
    ///
    /// Scans the complete selected text and has the replacement-copy cost of
    /// [`Self::replace`].
    ///
    /// # Errors
    ///
    /// Returns the same failures as [`Self::replace`].
    pub fn set(&mut self, text: &str) -> Result<&mut Self, SlideTextError> {
        let end = TextPosition::from_utf16_index(self.before.encode_utf16().count())
            .map_err(|_error| SlideTextError::InvalidSource)?;
        let span = TextSpan::new(TextPosition::ZERO, end)
            .map_err(|_error| SlideTextError::InvalidSource)?;
        self.replace(span, text)
    }

    /// Stage removal of all text while retaining the existing placeholder graph.
    ///
    /// # Costs
    ///
    /// Scans the complete selected text.
    ///
    /// # Errors
    ///
    /// Returns the same failures as [`Self::set`].
    pub fn clear(&mut self) -> Result<&mut Self, SlideTextError> {
        self.set("")
    }

    /// Validate and atomically publish the staged immutable candidate.
    ///
    /// # Costs
    ///
    /// A changed commit rewrites the storage and slide-preview owners in one
    /// or two IWA components, removes stale package previews, reassembles and
    /// reopens the complete package, and retains exact source and target
    /// artifacts in its reversible patch. A semantic no-op shares the source
    /// artifact.
    ///
    /// # Errors
    ///
    /// Returns a typed source, dependency, limit, allocation, or verification
    /// error before publishing any candidate.
    pub fn commit(self) -> Result<SlideTextCommit, SlideTextError> {
        let operation = self.operation.unwrap_or(Operation {
            span: TextSpan::default(),
            replacement: None,
        });
        let range = validate_span(&self.before, operation.span)?;
        let selected = self
            .before
            .get(range.clone())
            .ok_or(SlideTextError::InvalidSource)?;
        let replacement = operation.replacement.as_deref().unwrap_or(selected);
        validate_replacement(replacement)?;
        validate_consumed_text(selected)?;
        validate_candidate_text_memory(self.source, &self.before, operation.span, replacement)?;
        let replacement_units = utf16_len(replacement)?;
        let inverse_end = operation
            .span
            .start()
            .utf16_index()
            .checked_add(replacement_units)
            .ok_or(SlideTextError::InvalidSource)?;
        let inverse_span = TextSpan::new(
            operation.span.start(),
            TextPosition::from_utf16_code_units(inverse_end),
        )
        .map_err(|_error| SlideTextError::InvalidSource)?;

        let catalog = physical_catalog(self.source)?;
        let source_bytes = catalog.shared_source();
        let source_fingerprint = fingerprint(&source_bytes);

        if operation.replacement.is_none() {
            let before = Arc::new(self.before);
            let after = Arc::clone(&before);
            return Ok(SlideTextCommit {
                package: self.source.snapshot(),
                patch: SlideTextPatch {
                    source: Arc::clone(&source_bytes),
                    target: source_bytes,
                    source_fingerprint,
                    target_fingerprint: source_fingerprint,
                    position: self.position,
                    role: self.role,
                    placeholder_identifier: self.placeholder_identifier,
                    storage_identifier: self.storage_identifier,
                    span: operation.span,
                    inverse_span,
                    before,
                    after,
                    touched_components: 0,
                    target_requires_invalidated_previews: false,
                },
                diagnostics: SlideTextDiagnostics::unchanged(),
            });
        }
        self.source.validate().map_err(map_read_error)?;
        let current = text_snapshot_at(self.source, self.position, self.role)?
            .ok_or(SlideTextError::InvalidSource)?;
        if current.placeholder_identifier != self.placeholder_identifier
            || current.storage_identifier != self.storage_identifier
            || current.text != self.before
        {
            return Err(SlideTextError::InvalidSource);
        }
        if !catalog.source_is_exact() {
            return Err(SlideTextError::UnsupportedSource);
        }

        let changed_replacement = operation
            .replacement
            .as_deref()
            .ok_or(SlideTextError::InvalidSource)?;
        let after_text = replace_utf8_range(&self.before, range, changed_replacement)?;
        let before = Arc::new(self.before);
        let after = Arc::new(after_text);
        let (package, touched_components) = rewrite_text(
            self.source,
            self.position,
            self.role,
            self.placeholder_identifier,
            self.storage_identifier,
            operation.span,
            changed_replacement,
            after.as_str(),
        )?;
        let target = physical_catalog(&package)?.shared_source();
        let target_fingerprint = fingerprint(&target);
        Ok(SlideTextCommit {
            patch: SlideTextPatch {
                source: source_bytes,
                target,
                source_fingerprint,
                target_fingerprint,
                position: self.position,
                role: self.role,
                placeholder_identifier: self.placeholder_identifier,
                storage_identifier: self.storage_identifier,
                span: operation.span,
                inverse_span,
                before,
                after,
                touched_components,
                target_requires_invalidated_previews: true,
            },
            package,
            diagnostics: SlideTextDiagnostics::published(touched_components),
        })
    }
}

/// An exact-source-checked reversible slide-text patch.
///
/// The patch retains shared handles to the complete exact source and target
/// artifacts and to the before/after text. Cloning it is cheap, but keeping it
/// alive keeps those allocations alive.
#[derive(Clone, PartialEq, Eq)]
pub struct SlideTextPatch {
    source: Arc<[u8]>,
    target: Arc<[u8]>,
    source_fingerprint: u64,
    target_fingerprint: u64,
    position: Position,
    role: SlideTextRole,
    placeholder_identifier: u64,
    storage_identifier: u64,
    span: TextSpan,
    inverse_span: TextSpan,
    before: Arc<String>,
    after: Arc<String>,
    touched_components: usize,
    target_requires_invalidated_previews: bool,
}

impl fmt::Debug for SlideTextPatch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SlideTextPatch")
            .field("position", &self.position)
            .field("role", &self.role)
            .field("span", &self.span)
            .finish_non_exhaustive()
    }
}

impl SlideTextPatch {
    /// Return the semantic slide position selected by this patch.
    #[must_use]
    pub const fn position(&self) -> Position {
        self.position
    }

    /// Return the semantic placeholder role selected by this patch.
    #[must_use]
    pub const fn role(&self) -> SlideTextRole {
        self.role
    }

    /// Return the storage-relative UTF-16 span replaced by this patch.
    #[must_use]
    pub const fn span(&self) -> TextSpan {
        self.span
    }

    /// Borrow the complete placeholder text required from the source.
    #[must_use]
    pub fn before(&self) -> &str {
        self.before.as_str()
    }

    /// Borrow the complete placeholder text produced by the target.
    #[must_use]
    pub fn after(&self) -> &str {
        self.after.as_str()
    }

    /// Return the base package's compact diagnostic fingerprint.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.source_fingerprint
    }

    /// Return the committed package's compact diagnostic fingerprint.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.target_fingerprint
    }

    /// Return whether this patch preserves semantic text and exact bytes.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after
            && self.source_fingerprint == self.target_fingerprint
            && self.source.as_ref() == self.target.as_ref()
    }

    /// Return an exact reversible patch from the target back to its source.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            source: Arc::clone(&self.target),
            target: Arc::clone(&self.source),
            source_fingerprint: self.target_fingerprint,
            target_fingerprint: self.source_fingerprint,
            position: self.position,
            role: self.role,
            placeholder_identifier: self.placeholder_identifier,
            storage_identifier: self.storage_identifier,
            span: self.inverse_span,
            inverse_span: self.span,
            before: Arc::clone(&self.after),
            after: Arc::clone(&self.before),
            touched_components: self.touched_components,
            target_requires_invalidated_previews: self.touched_components != 0
                && !self.target_requires_invalidated_previews,
        }
    }
}

/// Compact publication evidence for one slide-text commit.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct SlideTextDiagnostics {
    changed: bool,
    touched_components: usize,
    full_reparse_performed: bool,
}

impl SlideTextDiagnostics {
    const fn unchanged() -> Self {
        Self {
            changed: false,
            touched_components: 0,
            full_reparse_performed: false,
        }
    }

    const fn published(touched_components: usize) -> Self {
        Self {
            changed: true,
            touched_components,
            full_reparse_performed: true,
        }
    }

    /// Return whether the committed package differs from its source.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Return the number of physical IWA components rewritten.
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }

    /// Return whether the complete candidate was reopened before publication.
    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }
}

/// The fully verified result of one slide title/body text transaction.
#[must_use = "a Keynote slide-text commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct SlideTextCommit {
    package: Package,
    patch: SlideTextPatch,
    diagnostics: SlideTextDiagnostics,
}

impl SlideTextCommit {
    /// Borrow the fully reopened immutable package snapshot.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume the commit and return its immutable package snapshot.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Borrow the reversible exact-source patch.
    #[must_use]
    pub const fn patch(&self) -> &SlideTextPatch {
        &self.patch
    }

    /// Borrow compact publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &SlideTextDiagnostics {
        &self.diagnostics
    }
}

impl Package {
    /// Read an existing title/body storage without exposing native identity.
    ///
    /// `Ok(None)` means that the slide has no reference for the requested
    /// role. `Some("")` denotes an existing empty storage.
    ///
    /// # Costs
    ///
    /// The returned text is an owned, bounded allocation.
    ///
    /// # Errors
    ///
    /// Returns a typed selector, source, limit, allocation, or dependency
    /// error when the requested semantic value cannot be read safely.
    pub fn slide_text<'selector>(
        &self,
        selector: impl Into<SlideSelector<'selector>>,
        role: SlideTextRole,
    ) -> Result<Option<String>, SlideTextError> {
        let position = resolve_position(self, selector.into())?;
        Ok(text_snapshot_at(self, position, role)?.map(|snapshot| snapshot.text))
    }

    /// Read one slide's existing title placeholder text.
    ///
    /// # Costs
    ///
    /// Has the same costs as [`Self::slide_text`].
    ///
    /// # Errors
    ///
    /// Returns the same failures as [`Self::slide_text`].
    pub fn slide_title<'selector>(
        &self,
        selector: impl Into<SlideSelector<'selector>>,
    ) -> Result<Option<String>, SlideTextError> {
        self.slide_text(selector, SlideTextRole::Title)
    }

    /// Read one slide's existing body placeholder text.
    ///
    /// # Costs
    ///
    /// Has the same costs as [`Self::slide_text`].
    ///
    /// # Errors
    ///
    /// Returns the same failures as [`Self::slide_text`].
    pub fn slide_body<'selector>(
        &self,
        selector: impl Into<SlideSelector<'selector>>,
    ) -> Result<Option<String>, SlideTextError> {
        self.slide_text(selector, SlideTextRole::Body)
    }

    /// Start one selector-first title/body text edit.
    ///
    /// # Costs
    ///
    /// The edit owns one bounded copy of the selected placeholder text.
    ///
    /// # Errors
    ///
    /// Returns a typed selector, missing-placeholder, source, limit,
    /// allocation, or dependency error.
    pub fn edit_slide_text<'selector>(
        &self,
        selector: impl Into<SlideSelector<'selector>>,
        role: SlideTextRole,
    ) -> Result<SlideTextEdit<'_>, SlideTextError> {
        SlideTextEdit::new(self, selector, role)
    }

    /// Start one selector-first slide-title edit.
    ///
    /// # Costs
    ///
    /// Has the same costs as [`Self::edit_slide_text`].
    ///
    /// # Errors
    ///
    /// Returns the same failures as [`Self::edit_slide_text`].
    pub fn edit_slide_title<'selector>(
        &self,
        selector: impl Into<SlideSelector<'selector>>,
    ) -> Result<SlideTextEdit<'_>, SlideTextError> {
        self.edit_slide_text(selector, SlideTextRole::Title)
    }

    /// Start one selector-first slide-body edit.
    ///
    /// # Costs
    ///
    /// Has the same costs as [`Self::edit_slide_text`].
    ///
    /// # Errors
    ///
    /// Returns the same failures as [`Self::edit_slide_text`].
    pub fn edit_slide_body<'selector>(
        &self,
        selector: impl Into<SlideSelector<'selector>>,
    ) -> Result<SlideTextEdit<'_>, SlideTextError> {
        self.edit_slide_text(selector, SlideTextRole::Body)
    }

    /// Apply an exact-source-checked slide-text patch.
    ///
    /// # Costs
    ///
    /// A changed application reopens and validates the retained target
    /// artifact; a no-op shares the current package snapshot.
    ///
    /// # Errors
    ///
    /// Returns [`SlideTextError::PatchConflict`] unless the patch belongs to
    /// this exact artifact, or another typed validation, limit, allocation,
    /// dependency, or verification error before publication.
    pub fn apply_slide_text(
        &self,
        patch: &SlideTextPatch,
    ) -> Result<SlideTextCommit, SlideTextError> {
        let catalog = physical_catalog(self)?;
        if fingerprint(catalog.source_bytes()) != patch.source_fingerprint
            || catalog.source_bytes() != patch.source.as_ref()
        {
            return Err(SlideTextError::PatchConflict);
        }
        if patch.is_noop() {
            return Ok(SlideTextCommit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: SlideTextDiagnostics::unchanged(),
            });
        }
        let current = text_snapshot_at(self, patch.position, patch.role)?
            .ok_or(SlideTextError::PatchConflict)?;
        if current.placeholder_identifier != patch.placeholder_identifier
            || current.storage_identifier != patch.storage_identifier
            || current.text != patch.before()
        {
            return Err(SlideTextError::PatchConflict);
        }
        self.validate().map_err(map_read_error)?;
        prove_exclusive_text_ownership(
            self,
            patch.position,
            patch.role,
            patch.placeholder_identifier,
            patch.storage_identifier,
        )?;
        if !catalog.source_is_exact() || fingerprint(&patch.target) != patch.target_fingerprint {
            return Err(SlideTextError::PatchConflict);
        }
        let candidate =
            Package::from_source_with_options(Arc::clone(&patch.target), self.state.options)
                .map_err(map_read_error)?;
        candidate.validate().map_err(map_read_error)?;
        let node_identifier = self
            .slide_record_at(patch.position.get())
            .map_err(map_read_error)?
            .ok_or(SlideTextError::PatchConflict)?
            .node_identifier;
        verify_candidate(
            self,
            &candidate,
            patch.position,
            patch.role,
            patch.after(),
            node_identifier,
            patch.target_requires_invalidated_previews,
        )?;
        Ok(SlideTextCommit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: SlideTextDiagnostics::published(patch.touched_components),
        })
    }
}

fn resolve_position(
    package: &Package,
    selector: SlideSelector<'_>,
) -> Result<Position, SlideTextError> {
    match selector {
        SlideSelector::Position(position) => package
            .slide_record_at(position.get())
            .map_err(map_read_error)?
            .map(|_record| position)
            .ok_or(SlideTextError::SlidePositionNotFound { position }),
        SlideSelector::Name(_) => package
            .show()
            .map_err(map_read_error)?
            .select_slide(selector)
            .map_err(|_error| SlideTextError::AmbiguousSelector)?
            .map(|slide| Position::new(slide.index()))
            .ok_or(SlideTextError::SlideNameNotFound),
    }
}

fn text_snapshot_at(
    package: &Package,
    position: Position,
    role: SlideTextRole,
) -> Result<Option<TextSnapshot>, SlideTextError> {
    let record = package
        .slide_record_at(position.get())
        .map_err(map_read_error)?
        .ok_or(SlideTextError::SlidePositionNotFound { position })?;
    let slide = package
        .required_object(record.slide_identifier, "Keynote slide")
        .map_err(map_read_error)?;
    let slide_payload = unique_payload(&slide.messages, &[SLIDE_MESSAGE_TYPE], "Keynote slide")
        .map_err(map_read_error)?;
    let owner = keynote_speaker_notes_codec::decode_slide_notes_owner(
        slide_payload,
        speaker_notes_decode_options(package, slide_payload)?,
    )
    .map_err(map_speaker_notes_codec_error)?;
    let placeholder = match role {
        SlideTextRole::Title => owner.title_placeholder(),
        SlideTextRole::Body => owner.body_placeholder(),
    };
    let Some(placeholder_identifier) = placeholder.map(|reference| reference.identifier().get())
    else {
        return Ok(None);
    };
    let placeholder_object = package
        .required_object(placeholder_identifier, "Keynote slide placeholder")
        .map_err(map_read_error)?;
    if placeholder_object
        .messages
        .iter()
        .any(|message| message.type_ == SHAPE_INFO_MESSAGE_TYPE)
    {
        return Err(SlideTextError::DependentContent);
    }
    let placeholder_payload = unique_payload(
        &placeholder_object.messages,
        &[PLACEHOLDER_MESSAGE_TYPE],
        "Keynote slide placeholder",
    )
    .map_err(map_read_error)?;
    let placeholder_owner = keynote_placeholder_text_codec::decode_placeholder_text_owner(
        placeholder_payload,
        placeholder_decode_options(package, placeholder_payload)?,
    )
    .map_err(map_placeholder_codec_error)?;
    validate_placeholder_kind(role, placeholder_owner.kind())?;
    let owned_storage = placeholder_owner
        .owned_storage()
        .ok_or(SlideTextError::InvalidSource)?;
    let owned_storage = ReferenceSemantics::from_placeholder_codec(owned_storage);
    let storage_identifier = owned_storage.identifier;
    let mut scan_budget = WireScanBudget::new(package)?;
    if nested_reference_snapshot(package, placeholder_payload, &[1, 2], &mut scan_budget)?
        .is_some_and(|deprecated| deprecated != owned_storage)
    {
        return Err(SlideTextError::DependentContent);
    }
    if nested_reference_identifier(package, placeholder_payload, &[1, 3], &mut scan_budget)?
        == Some(storage_identifier)
    {
        return Err(SlideTextError::DependentContent);
    }
    let storage = package
        .required_object(storage_identifier, "Keynote placeholder text storage")
        .map_err(map_read_error)?;
    one_message(&storage.messages, STORAGE_MESSAGE_TYPE)?;
    let mut budget = SemanticBudget::new(package.semantic_limits());
    budget
        .charge_references(2, semantic_path(position, role))
        .map_err(map_read_error)?;
    let text = package
        .required_text_storage(storage, &mut budget, semantic_path(position, role))
        .map_err(map_read_error)?
        .into_text();
    Ok(Some(TextSnapshot {
        placeholder_identifier,
        storage_identifier,
        text,
    }))
}

fn validate_placeholder_kind(role: SlideTextRole, kind: Option<i32>) -> Result<(), SlideTextError> {
    // Absence and explicit zero both mean the native generic placeholder.
    // Any non-generic explicit kind must agree with the slide's role edge;
    // slide-number, object, opposite-role, and future tags are not writable
    // through this focused capability.
    match (role, kind) {
        (_, None | Some(0)) | (SlideTextRole::Title, Some(2)) | (SlideTextRole::Body, Some(3)) => {
            Ok(())
        },
        _ => Err(SlideTextError::DependentContent),
    }
}

const fn semantic_path(position: Position, role: SlideTextRole) -> SemanticPath {
    match role {
        SlideTextRole::Title => SemanticPath::SlideTitle {
            index: position.get(),
        },
        SlideTextRole::Body => SemanticPath::SlideBody {
            index: position.get(),
        },
    }
}

fn speaker_notes_decode_options(
    package: &Package,
    payload: &[u8],
) -> Result<keynote_speaker_notes_codec::DecodeOptions, SlideTextError> {
    let limits = package.wire_limits().map_err(map_wire_error)?;
    let recursion =
        u32::try_from(limits.max_nesting()).map_err(|_error| SlideTextError::InvalidSource)?;
    Ok(keynote_speaker_notes_codec::DecodeOptions::new(
        payload.len().min(limits.max_input_bytes()),
        limits.max_fields(),
        limits.max_rewrite_work(),
        recursion,
    ))
}

fn placeholder_decode_options(
    package: &Package,
    payload: &[u8],
) -> Result<keynote_placeholder_text_codec::DecodeOptions, SlideTextError> {
    let limits = package.wire_limits().map_err(map_wire_error)?;
    let recursion =
        u32::try_from(limits.max_nesting()).map_err(|_error| SlideTextError::InvalidSource)?;
    Ok(keynote_placeholder_text_codec::DecodeOptions::new(
        payload.len().min(limits.max_input_bytes()),
        limits.max_fields(),
        limits.max_rewrite_work(),
        recursion,
    ))
}

fn map_speaker_notes_codec_error(
    error: keynote_speaker_notes_codec::DecodeError,
) -> SlideTextError {
    if let Some(limit) = error.wire_resource_limit() {
        return match limit {
            keynote_speaker_notes_codec::WireResourceLimit::Bytes { observed, maximum } => {
                SlideTextError::LimitExceeded {
                    kind: SlideTextLimitKind::WireBytes,
                    observed: observed.map_or(u64::MAX, usize_to_u64),
                    maximum: maximum.map_or(u64::MAX, usize_to_u64),
                }
            },
            keynote_speaker_notes_codec::WireResourceLimit::Nesting { observed, maximum } => {
                SlideTextError::LimitExceeded {
                    kind: SlideTextLimitKind::WireNesting,
                    observed: observed.map_or(u64::MAX, u64::from),
                    maximum: maximum.map_or(u64::MAX, u64::from),
                }
            },
        };
    }
    if let Some((observed, maximum)) = error.field_limit_values() {
        return SlideTextError::LimitExceeded {
            kind: SlideTextLimitKind::WireFields,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        };
    }
    if let Some((observed, maximum)) = error.work_limit_values() {
        return SlideTextError::LimitExceeded {
            kind: SlideTextLimitKind::WireWork,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        };
    }
    SlideTextError::InvalidSource
}

fn map_placeholder_codec_error(
    error: keynote_placeholder_text_codec::DecodeError,
) -> SlideTextError {
    if let Some((observed, maximum)) = error.message_byte_limit_values() {
        return SlideTextError::LimitExceeded {
            kind: SlideTextLimitKind::WireBytes,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        };
    }
    if let Some((observed, maximum)) = error.recursion_limit_values() {
        return SlideTextError::LimitExceeded {
            kind: SlideTextLimitKind::WireNesting,
            observed: u64::from(observed),
            maximum: u64::from(maximum),
        };
    }
    if let Some((observed, maximum)) = error.field_limit_values() {
        return SlideTextError::LimitExceeded {
            kind: SlideTextLimitKind::WireFields,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        };
    }
    if let Some((observed, maximum)) = error.work_limit_values() {
        return SlideTextError::LimitExceeded {
            kind: SlideTextLimitKind::WireWork,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        };
    }
    SlideTextError::InvalidSource
}

fn rewrite_text(
    source: &Package,
    position: Position,
    role: SlideTextRole,
    expected_placeholder_identifier: u64,
    expected_storage_identifier: u64,
    span: TextSpan,
    replacement: &str,
    expected: &str,
) -> Result<(Package, usize), SlideTextError> {
    let record = source
        .slide_record_at(position.get())
        .map_err(map_read_error)?
        .ok_or(SlideTextError::SlidePositionNotFound { position })?;
    if record.node_identifier == expected_storage_identifier {
        return Err(SlideTextError::DependentContent);
    }
    let snapshot =
        text_snapshot_at(source, position, role)?.ok_or(SlideTextError::InvalidSource)?;
    if snapshot.placeholder_identifier != expected_placeholder_identifier
        || snapshot.storage_identifier != expected_storage_identifier
    {
        return Err(SlideTextError::InvalidSource);
    }
    prove_exclusive_text_ownership(
        source,
        position,
        role,
        expected_placeholder_identifier,
        expected_storage_identifier,
    )?;
    prove_slide_text_caches_absent(source, record.slide_identifier)?;
    let catalog = physical_catalog(source)?;
    let physical_limits = source.state.options.archive();
    let archive_limits = physical_limits
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let start = usize::try_from(span.start().utf16_index())
        .map_err(|_error| SlideTextError::InvalidSource)?;
    let end = usize::try_from(span.end().utf16_index())
        .map_err(|_error| SlideTextError::InvalidSource)?;
    let mut component_names = Vec::new();
    component_names
        .try_reserve_exact(2)
        .map_err(|_allocation| SlideTextError::Allocation { amount: 2 })?;
    for identifier in [expected_storage_identifier, record.node_identifier] {
        let mut matches = catalog
            .components()
            .iter()
            .filter(|component| component.archive().object(identifier).is_some());
        let component = matches.next().ok_or(SlideTextError::InvalidSource)?;
        if matches.next().is_some() {
            return Err(SlideTextError::InvalidSource);
        }
        if !component_names.contains(&component.name()) {
            component_names.push(component.name());
        }
    }
    let mut compressed_components = Vec::new();
    compressed_components
        .try_reserve_exact(component_names.len())
        .map_err(|_allocation| SlideTextError::Allocation {
            amount: component_names.len(),
        })?;
    let mut storage_changed = false;
    let mut node_changed = false;
    for component_name in &component_names {
        let entry = catalog
            .package()
            .iter()
            .find(|entry| entry.name() == *component_name)
            .ok_or(SlideTextError::InvalidSource)?;
        if entry.is_opaque() {
            return Err(SlideTextError::InvalidSource);
        }
        let stream = SnappyStream::decompress_with_limits(
            entry.data(),
            physical_limits.snappy_limits().map_err(map_archive_error)?,
        )
        .map_err(map_core_error)?;
        let mut archive = Archive::parse_with_limits(stream.as_bytes(), archive_limits)
            .map_err(map_core_error)?;
        validate_canonical_object_length_prefixes(stream.as_bytes(), &archive)?;
        if let Some(object) = archive.object(expected_storage_identifier) {
            let (message_index, message) = one_message(&object.messages, STORAGE_MESSAGE_TYPE)?;
            validate_selected_storage_metadata(object, message_index)?;
            let rewrite = litchi_iwa_text_wire::rewrite_storage_text_with_behavior_and_limits(
                &message.data,
                start..end,
                replacement,
                RewriteBehavior::PreserveOnEqualText,
                storage_rewrite_limits(source, archive_limits.max_message_bytes())?,
            )
            .map_err(map_text_rewrite_error)?;
            if !rewrite.removed_object_references().is_empty()
                || !rewrite.removed_object_references_by_field().is_empty()
                || rewrite.object_reference_occurrences_before()
                    != rewrite.object_reference_occurrences_after()
            {
                return Err(SlideTextError::DependentContent);
            }
            if !rewrite.changed() || std::mem::replace(&mut storage_changed, true) {
                return Err(SlideTextError::Verification);
            }
            archive
                .object_mut(expected_storage_identifier)
                .ok_or(SlideTextError::InvalidSource)?
                .replace_message_preserving_header_with_limits(
                    message_index,
                    RawMessage {
                        type_: STORAGE_MESSAGE_TYPE,
                        data: rewrite.into_bytes(),
                    },
                    archive_limits,
                )
                .map_err(map_core_error)?;
        }
        if let Some(node) = archive.object_mut(record.node_identifier) {
            if std::mem::replace(&mut node_changed, true) {
                return Err(SlideTextError::Verification);
            }
            super::slide_preview::invalidate(
                node,
                archive_limits,
                source.wire_limits().map_err(map_wire_error)?,
            )
            .map_err(map_slide_preview_error)?;
        }
        let archive_bytes = archive
            .to_bytes_with_limits(archive_limits)
            .map_err(map_core_error)?;
        let compressed = SnappyStream::compress(&archive_bytes).map_err(map_core_error)?;
        compressed_components.push((*component_name, compressed));
    }
    if !storage_changed || !node_changed {
        return Err(SlideTextError::Verification);
    }
    let mut edits = Vec::new();
    edits
        .try_reserve_exact(compressed_components.len())
        .map_err(|_allocation| SlideTextError::Allocation {
            amount: compressed_components.len(),
        })?;
    for (name, compressed) in &compressed_components {
        edits.push(EntryEdit::new(name, compressed.as_slice()));
    }
    let mut deleted_previews = Vec::new();
    deleted_previews
        .try_reserve_exact(PREVIEW_ENTRY_NAMES.len())
        .map_err(|_allocation| SlideTextError::Allocation {
            amount: PREVIEW_ENTRY_NAMES.len(),
        })?;
    for name in PREVIEW_ENTRY_NAMES {
        if catalog.package().iter().any(|entry| entry.name() == name) {
            deleted_previews.push(name);
        }
    }
    let output = catalog
        .package()
        .reassemble_with_deletions_to_bytes(&edits, &deleted_previews, physical_limits)
        .map_err(map_archive_error)?;
    let candidate = Package::from_source_with_options(output.into(), source.state.options)
        .map_err(map_read_error)?;
    candidate.validate().map_err(map_read_error)?;
    verify_candidate(
        source,
        &candidate,
        position,
        role,
        expected,
        record.node_identifier,
        true,
    )?;
    Ok((candidate, compressed_components.len()))
}

fn prove_slide_text_caches_absent(
    package: &Package,
    slide_identifier: u64,
) -> Result<(), SlideTextError> {
    let slide = package
        .required_object(slide_identifier, "Keynote slide")
        .map_err(map_read_error)?;
    let payload = unique_payload(&slide.messages, &[SLIDE_MESSAGE_TYPE], "Keynote slide")
        .map_err(map_read_error)?;
    let limits = package.wire_limits().map_err(map_wire_error)?;
    let view = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    if view.fields().any(|field| matches!(field.number(), 37 | 38)) {
        return Err(SlideTextError::DependentContent);
    }
    Ok(())
}

fn verify_candidate(
    source: &Package,
    candidate: &Package,
    position: Position,
    role: SlideTextRole,
    expected: &str,
    slide_node_identifier: u64,
    require_invalidated_previews: bool,
) -> Result<(), SlideTextError> {
    if source.state.total_objects != candidate.state.total_objects {
        return Err(SlideTextError::Verification);
    }
    let source_text =
        text_snapshot_at(source, position, role)?.ok_or(SlideTextError::Verification)?;
    let candidate_text =
        text_snapshot_at(candidate, position, role)?.ok_or(SlideTextError::Verification)?;
    if source_text.placeholder_identifier != candidate_text.placeholder_identifier
        || source_text.storage_identifier != candidate_text.storage_identifier
        || candidate_text.text != expected
    {
        return Err(SlideTextError::Verification);
    }
    prove_exclusive_text_ownership(
        source,
        position,
        role,
        source_text.placeholder_identifier,
        source_text.storage_identifier,
    )?;
    prove_exclusive_text_ownership(
        candidate,
        position,
        role,
        candidate_text.placeholder_identifier,
        candidate_text.storage_identifier,
    )?;
    verify_untouched_objects(
        source,
        candidate,
        source_text.storage_identifier,
        slide_node_identifier,
        require_invalidated_previews,
    )?;
    if require_invalidated_previews {
        let candidate_catalog = physical_catalog(candidate)?;
        if PREVIEW_ENTRY_NAMES.iter().any(|name| {
            candidate_catalog
                .package()
                .iter()
                .any(|entry| entry.name() == *name)
        }) {
            return Err(SlideTextError::Verification);
        }
    }

    // Full semantic readback must preserve slide topology and every semantic
    // property whose physical owner was not selected.
    let before = source.slides().map_err(map_read_error)?;
    let after = candidate.slides().map_err(map_read_error)?;
    if before.len() != after.len() {
        return Err(SlideTextError::Verification);
    }
    for (index, (old, new)) in before.iter().zip(after).enumerate() {
        if index != position.get() {
            if old != new {
                return Err(SlideTextError::Verification);
            }
            continue;
        }
        if old.index() != new.index()
            || old.is_skipped() != new.is_skipped()
            || old.name() != new.name()
            || old.builds() != new.builds()
            || old.transition() != new.transition()
            || old.notes() != new.notes()
        {
            return Err(SlideTextError::Verification);
        }
        match role {
            SlideTextRole::Title => {
                if new.title() != nonempty(expected) {
                    return Err(SlideTextError::Verification);
                }
            },
            SlideTextRole::Body => {
                if old.title() != new.title() {
                    return Err(SlideTextError::Verification);
                }
            },
        }
    }
    Ok(())
}

fn verify_untouched_objects(
    source: &Package,
    candidate: &Package,
    selected_storage_identifier: u64,
    slide_node_identifier: u64,
    require_invalidated_previews: bool,
) -> Result<(), SlideTextError> {
    let source_catalog = physical_catalog(source)?;
    let candidate_catalog = physical_catalog(candidate)?;
    if source_catalog.components().len() != candidate_catalog.components().len() {
        return Err(SlideTextError::Verification);
    }
    let mut previous_component_name = None;
    let mut selected_storage_seen = false;
    let mut slide_node_seen = false;
    for (source_component, candidate_component) in source_catalog
        .components()
        .iter()
        .zip(candidate_catalog.components().iter())
    {
        if source_component.name() != candidate_component.name()
            || previous_component_name.is_some_and(|previous| previous >= source_component.name())
            || source_component.archive().objects.len()
                != candidate_component.archive().objects.len()
        {
            return Err(SlideTextError::Verification);
        }
        previous_component_name = Some(source_component.name());
        for (source_object, candidate_object) in source_component
            .archive()
            .objects
            .iter()
            .zip(&candidate_component.archive().objects)
        {
            let source_identifier = source_object
                .archive_info
                .identifier
                .ok_or(SlideTextError::Verification)?;
            let candidate_identifier = candidate_object
                .archive_info
                .identifier
                .ok_or(SlideTextError::Verification)?;
            if source_identifier != candidate_identifier {
                return Err(SlideTextError::Verification);
            }
            if source_identifier == selected_storage_identifier {
                if std::mem::replace(&mut selected_storage_seen, true) {
                    return Err(SlideTextError::Verification);
                }
                verify_selected_storage_object(source_object, candidate_object)?;
            } else if source_identifier == slide_node_identifier {
                if std::mem::replace(&mut slide_node_seen, true)
                    || (require_invalidated_previews
                        && !super::slide_preview::is_invalidated(
                            candidate_object,
                            candidate.wire_limits().map_err(map_wire_error)?,
                        )
                        .map_err(map_slide_preview_error)?)
                {
                    return Err(SlideTextError::Verification);
                }
            } else if source_object.archive_info != candidate_object.archive_info
                || source_object.messages != candidate_object.messages
            {
                return Err(SlideTextError::Verification);
            }
        }
    }
    if selected_storage_seen && slide_node_seen {
        Ok(())
    } else {
        Err(SlideTextError::Verification)
    }
}

fn verify_selected_storage_object(
    source: &litchi_iwa_core::ArchiveObject,
    candidate: &litchi_iwa_core::ArchiveObject,
) -> Result<(), SlideTextError> {
    if source.messages.len() != candidate.messages.len()
        || source.archive_info.message_infos.len() != candidate.archive_info.message_infos.len()
    {
        return Err(SlideTextError::Verification);
    }
    let mut selected = None;
    for (index, (old, new)) in source.messages.iter().zip(&candidate.messages).enumerate() {
        if old.type_ != new.type_ {
            return Err(SlideTextError::Verification);
        }
        if old.type_ == STORAGE_MESSAGE_TYPE {
            if selected.replace(index).is_some() {
                return Err(SlideTextError::Verification);
            }
        } else if old != new {
            return Err(SlideTextError::Verification);
        }
    }
    let selected = selected.ok_or(SlideTextError::Verification)?;
    if source.archive_info.identifier != candidate.archive_info.identifier
        || source.archive_info.should_merge != candidate.archive_info.should_merge
    {
        return Err(SlideTextError::Verification);
    }
    for (index, (old, new)) in source
        .archive_info
        .message_infos
        .iter()
        .zip(&candidate.archive_info.message_infos)
        .enumerate()
    {
        if old.type_ != new.type_
            || old.versions != new.versions
            || (index != selected && old.length != new.length)
            || old.field_infos != new.field_infos
            || old.object_references != new.object_references
            || old.data_references != new.data_references
            || old.base_message_index != new.base_message_index
            || old.diff_merge_version != new.diff_merge_version
            || old.diff_field_path != new.diff_field_path
            || old.fields_to_remove != new.fields_to_remove
            || old.diff_read_version != new.diff_read_version
        {
            return Err(SlideTextError::Verification);
        }
    }
    Ok(())
}

fn validate_selected_storage_metadata(
    object: &litchi_iwa_core::ArchiveObject,
    message_index: usize,
) -> Result<(), SlideTextError> {
    let message = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(SlideTextError::InvalidSource)?;
    if object.archive_info.should_merge == Some(true)
        || message.base_message_index.is_some()
        || !message.diff_merge_version.is_empty()
        || message.diff_field_path.is_some()
        || !message.fields_to_remove.is_empty()
        || !message.diff_read_version.is_empty()
    {
        return Err(SlideTextError::DependentContent);
    }
    Ok(())
}

fn nonempty(text: &str) -> Option<&str> {
    (!text.is_empty()).then_some(text)
}

fn prove_exclusive_text_ownership(
    package: &Package,
    position: Position,
    role: SlideTextRole,
    placeholder_identifier: u64,
    storage_identifier: u64,
) -> Result<(), SlideTextError> {
    let mut scan_budget = WireScanBudget::new(package)?;
    let slide_identifier = package
        .slide_record_at(position.get())
        .map_err(map_read_error)?
        .ok_or(SlideTextError::SlidePositionNotFound { position })?
        .slide_identifier;
    prove_metadata_ownership(
        package,
        slide_identifier,
        role,
        placeholder_identifier,
        storage_identifier,
    )?;

    let mut role_edges = 0usize;
    let mut storage_edges = 0usize;
    for component in package.state.source.components().iter() {
        for object in &component.archive().objects {
            let owner_identifier = object
                .archive_info
                .identifier
                .ok_or(SlideTextError::InvalidSource)?;
            for message in &object.messages {
                match message.type_ {
                    SLIDE_MESSAGE_TYPE => {
                        // Scan the two singular edges without constructing a
                        // generated view for every slide. Only the selected
                        // owner is forced through the private Buffa projection;
                        // raw source records remain authoritative for alias
                        // detection and preservation.
                        let title = scan_nested_reference_identifier(
                            package,
                            &message.data,
                            &[5],
                            &mut scan_budget,
                        )?;
                        let body = scan_nested_reference_identifier(
                            package,
                            &message.data,
                            &[6],
                            &mut scan_budget,
                        )?;
                        if title == Some(placeholder_identifier)
                            || body == Some(placeholder_identifier)
                        {
                            if owner_identifier != slide_identifier {
                                return Err(SlideTextError::DependentContent);
                            }
                            let owner = keynote_speaker_notes_codec::decode_slide_notes_owner(
                                &message.data,
                                speaker_notes_decode_options(package, &message.data)?,
                            )
                            .map_err(map_speaker_notes_codec_error)?;
                            let projected_title = owner
                                .title_placeholder()
                                .map(|reference| reference.identifier().get());
                            let projected_body = owner
                                .body_placeholder()
                                .map(|reference| reference.identifier().get());
                            let matches_role = match role {
                                SlideTextRole::Title => {
                                    projected_title == Some(placeholder_identifier)
                                        && projected_body != Some(placeholder_identifier)
                                },
                                SlideTextRole::Body => {
                                    projected_body == Some(placeholder_identifier)
                                        && projected_title != Some(placeholder_identifier)
                                },
                            };
                            role_edges = role_edges
                                .checked_add(1)
                                .ok_or(SlideTextError::InvalidSource)?;
                            if role_edges > 1 || !matches_role {
                                return Err(SlideTextError::DependentContent);
                            }
                        }
                    },
                    PLACEHOLDER_MESSAGE_TYPE => {
                        let scanned_owned_storage = scan_nested_reference_snapshot(
                            package,
                            &message.data,
                            &[1, 4],
                            &mut scan_budget,
                        )?;
                        let scanned_deprecated_storage = scan_nested_reference_snapshot(
                            package,
                            &message.data,
                            &[1, 2],
                            &mut scan_budget,
                        )?;
                        let scanned_text_flow = scan_nested_reference_snapshot(
                            package,
                            &message.data,
                            &[1, 3],
                            &mut scan_budget,
                        )?;
                        if scanned_text_flow
                            .is_some_and(|reference| reference.identifier == storage_identifier)
                        {
                            return Err(SlideTextError::DependentContent);
                        }
                        if scanned_owned_storage
                            .is_none_or(|reference| reference.identifier != storage_identifier)
                            && scanned_deprecated_storage
                                .is_none_or(|reference| reference.identifier != storage_identifier)
                        {
                            // Other placeholder roles may carry canonical zero
                            // sentinel references. They cannot own this selected
                            // nonzero storage and therefore do not need their
                            // role-specific graph forced through the strict
                            // selected-placeholder projection.
                            continue;
                        }
                        if owner_identifier != placeholder_identifier {
                            return Err(SlideTextError::DependentContent);
                        }
                        let owner = keynote_placeholder_text_codec::decode_placeholder_text_owner(
                            &message.data,
                            placeholder_decode_options(package, &message.data)?,
                        )
                        .map_err(map_placeholder_codec_error)?;
                        let owned_storage = owner
                            .owned_storage()
                            .map(ReferenceSemantics::from_placeholder_codec);
                        let deprecated_storage = nested_reference_snapshot(
                            package,
                            &message.data,
                            &[1, 2],
                            &mut scan_budget,
                        )?;
                        if deprecated_storage
                            .is_some_and(|reference| reference.identifier == storage_identifier)
                            && (owner_identifier != placeholder_identifier
                                || owned_storage.map(|reference| reference.identifier)
                                    != Some(storage_identifier))
                        {
                            return Err(SlideTextError::DependentContent);
                        }
                        if nested_reference_identifier(
                            package,
                            &message.data,
                            &[1, 3],
                            &mut scan_budget,
                        )? == Some(storage_identifier)
                        {
                            return Err(SlideTextError::DependentContent);
                        }
                        if owned_storage.map(|reference| reference.identifier)
                            != Some(storage_identifier)
                        {
                            continue;
                        }
                        if deprecated_storage
                            .is_some_and(|deprecated| Some(deprecated) != owned_storage)
                        {
                            return Err(SlideTextError::DependentContent);
                        }
                        storage_edges = storage_edges
                            .checked_add(1)
                            .ok_or(SlideTextError::InvalidSource)?;
                        if storage_edges > 1 || owner_identifier != placeholder_identifier {
                            return Err(SlideTextError::DependentContent);
                        }
                        validate_placeholder_kind(role, owner.kind())?;
                    },
                    NOTE_MESSAGE_TYPE => {
                        if scan_nested_reference_identifier(
                            package,
                            &message.data,
                            &[1],
                            &mut scan_budget,
                        )? == Some(storage_identifier)
                        {
                            return Err(SlideTextError::DependentContent);
                        }
                    },
                    SHAPE_INFO_MESSAGE_TYPE => {
                        for path in [&[2][..], &[3][..], &[4][..]] {
                            if nested_reference_identifier(
                                package,
                                &message.data,
                                path,
                                &mut scan_budget,
                            )? == Some(storage_identifier)
                            {
                                return Err(SlideTextError::DependentContent);
                            }
                        }
                    },
                    _ => {},
                }
            }
        }
    }
    if role_edges == 1 && storage_edges == 1 {
        Ok(())
    } else {
        Err(SlideTextError::InvalidSource)
    }
}

fn nested_reference_identifier(
    package: &Package,
    payload: &[u8],
    path: &[u32],
    budget: &mut WireScanBudget,
) -> Result<Option<u64>, SlideTextError> {
    Ok(nested_reference_snapshot(package, payload, path, budget)?
        .map(|reference| reference.identifier))
}

fn nested_reference_snapshot(
    _package: &Package,
    payload: &[u8],
    path: &[u32],
    budget: &mut WireScanBudget,
) -> Result<Option<ReferenceSemantics>, SlideTextError> {
    nested_reference_snapshot_with_zero(payload, path, false, budget)
}

fn scan_nested_reference_snapshot(
    _package: &Package,
    payload: &[u8],
    path: &[u32],
    budget: &mut WireScanBudget,
) -> Result<Option<ReferenceSemantics>, SlideTextError> {
    nested_reference_snapshot_with_zero(payload, path, true, budget)
}

fn scan_nested_reference_identifier(
    _package: &Package,
    payload: &[u8],
    path: &[u32],
    budget: &mut WireScanBudget,
) -> Result<Option<u64>, SlideTextError> {
    Ok(
        nested_reference_snapshot_with_zero(payload, path, false, budget)?
            .map(|reference| reference.identifier),
    )
}

fn nested_reference_snapshot_with_zero(
    payload: &[u8],
    path: &[u32],
    allow_zero: bool,
    budget: &mut WireScanBudget,
) -> Result<Option<ReferenceSemantics>, SlideTextError> {
    let Some(reference) = unique_nested_payload(payload, path, budget)? else {
        return Ok(None);
    };
    let view = budget.parse(reference)?;
    let mut identifier = None;
    let mut deprecated_type = None;
    let mut deprecated_is_external = None;
    for field in view.fields() {
        field.validate_canonical_framing().map_err(map_wire_error)?;
        let value = if field.wire_type() == 0 {
            let (value, bytes) = decode_varint_from_bytes(field.payload())
                .map_err(|_error| SlideTextError::InvalidSource)?;
            if bytes != field.payload().len() || bytes != encoded_len(value) {
                return Err(SlideTextError::InvalidSource);
            }
            Some(value)
        } else {
            None
        };
        match field.number() {
            1 => {
                let value = value.ok_or(SlideTextError::InvalidSource)?;
                if identifier.replace(value).is_some() || (value == 0 && !allow_zero) {
                    return Err(SlideTextError::InvalidSource);
                }
            },
            2 => {
                let value = value.ok_or(SlideTextError::InvalidSource)?;
                if deprecated_type.is_some() {
                    return Err(SlideTextError::InvalidSource);
                }
                deprecated_type = Some(decode_canonical_int32(value)?);
            },
            3 => {
                let value = value.ok_or(SlideTextError::InvalidSource)?;
                if deprecated_is_external.is_some() || value > 1 {
                    return Err(SlideTextError::InvalidSource);
                }
                deprecated_is_external = Some(value == 1);
            },
            _ => {},
        }
    }
    let identifier = identifier.ok_or(SlideTextError::InvalidSource)?;
    Ok(Some(ReferenceSemantics {
        identifier,
        deprecated_type,
        deprecated_is_external,
    }))
}

fn decode_canonical_int32(value: u64) -> Result<i32, SlideTextError> {
    if value <= 2_147_483_647 {
        return i32::try_from(value).map_err(|_error| SlideTextError::InvalidSource);
    }
    if value < 0xffff_ffff_8000_0000 {
        return Err(SlideTextError::InvalidSource);
    }
    let bits = u32::try_from(value & u64::from(u32::MAX))
        .map_err(|_error| SlideTextError::InvalidSource)?;
    Ok(i32::from_ne_bytes(bits.to_ne_bytes()))
}

fn unique_nested_payload<'a>(
    payload: &'a [u8],
    path: &[u32],
    budget: &mut WireScanBudget,
) -> Result<Option<&'a [u8]>, SlideTextError> {
    let Some((&field_number, remaining)) = path.split_first() else {
        return Ok(Some(payload));
    };
    let view = budget.parse(payload)?;
    let mut selected = None;
    for field in view.fields().filter(|field| field.number() == field_number) {
        if selected.is_some() || field.wire_type() != 2 {
            return Err(SlideTextError::InvalidSource);
        }
        selected = Some(field.canonical_payload().map_err(map_wire_error)?);
    }
    let Some(selected) = selected else {
        return Ok(None);
    };
    unique_nested_payload(selected, remaining, budget)
}

fn prove_metadata_ownership(
    package: &Package,
    expected_slide_identifier: u64,
    role: SlideTextRole,
    placeholder_identifier: u64,
    storage_identifier: u64,
) -> Result<(), SlideTextError> {
    let mut saw_placeholder_owner = false;
    let mut saw_storage_owner = false;
    for component in package.state.source.components().iter() {
        for object in &component.archive().objects {
            let owner_identifier = object
                .archive_info
                .identifier
                .ok_or(SlideTextError::InvalidSource)?;
            for message in &object.archive_info.message_infos {
                let mut aggregate_placeholder = false;
                let mut aggregate_storage = false;
                for reference in &message.object_references {
                    if *reference == placeholder_identifier
                        && std::mem::replace(&mut aggregate_placeholder, true)
                    {
                        return Err(SlideTextError::DependentContent);
                    }
                    if *reference == storage_identifier
                        && std::mem::replace(&mut aggregate_storage, true)
                    {
                        return Err(SlideTextError::DependentContent);
                    }
                }
                let mut role_path_seen = false;
                let mut owned_path_seen = false;
                let mut z_order_path_seen = false;
                let mut storage_path_seen = false;
                for field in &message.field_infos {
                    for reference in &field.object_references {
                        if *reference == placeholder_identifier {
                            let role_path = match role {
                                SlideTextRole::Title => [5],
                                SlideTextRole::Body => [6],
                            };
                            match field.path.as_slice() {
                                path if path == role_path => {
                                    if std::mem::replace(&mut role_path_seen, true) {
                                        return Err(SlideTextError::DependentContent);
                                    }
                                },
                                [7] => {
                                    if std::mem::replace(&mut owned_path_seen, true) {
                                        return Err(SlideTextError::DependentContent);
                                    }
                                },
                                [42] => {
                                    if std::mem::replace(&mut z_order_path_seen, true) {
                                        return Err(SlideTextError::DependentContent);
                                    }
                                },
                                _ => return Err(SlideTextError::DependentContent),
                            }
                        }
                        if *reference == storage_identifier
                            && (field.path.as_slice() != [1, 4]
                                || std::mem::replace(&mut storage_path_seen, true))
                        {
                            return Err(SlideTextError::DependentContent);
                        }
                    }
                }
                let has_placeholder =
                    aggregate_placeholder || role_path_seen || owned_path_seen || z_order_path_seen;
                if has_placeholder
                    && (std::mem::replace(&mut saw_placeholder_owner, true)
                        || owner_identifier != expected_slide_identifier
                        || message.type_ != SLIDE_MESSAGE_TYPE)
                {
                    return Err(SlideTextError::DependentContent);
                }
                let has_storage = aggregate_storage || storage_path_seen;
                if has_storage
                    && (std::mem::replace(&mut saw_storage_owner, true)
                        || owner_identifier != placeholder_identifier
                        || message.type_ != PLACEHOLDER_MESSAGE_TYPE)
                {
                    return Err(SlideTextError::DependentContent);
                }
            }
        }
    }
    if !saw_placeholder_owner || !saw_storage_owner {
        return Err(SlideTextError::InvalidSource);
    }
    Ok(())
}

fn one_message(
    messages: &[RawMessage],
    message_type: u32,
) -> Result<(usize, &RawMessage), SlideTextError> {
    let mut matches = messages
        .iter()
        .enumerate()
        .filter(|(_index, message)| message.type_ == message_type);
    let item = matches.next().ok_or(SlideTextError::InvalidSource)?;
    if matches.next().is_some() {
        return Err(SlideTextError::InvalidSource);
    }
    Ok(item)
}

fn validate_canonical_object_length_prefixes(
    source: &[u8],
    archive: &Archive,
) -> Result<(), SlideTextError> {
    for object in &archive.objects {
        let offset = usize::try_from(object.header_offset)
            .map_err(|_error| SlideTextError::InvalidSource)?;
        let remaining = source.get(offset..).ok_or(SlideTextError::InvalidSource)?;
        let (header_bytes, prefix_bytes) =
            decode_varint_from_bytes(remaining).map_err(|_error| SlideTextError::InvalidSource)?;
        if prefix_bytes != encoded_len(header_bytes) {
            return Err(SlideTextError::InvalidSource);
        }
        let framed_header_bytes = header_bytes
            .checked_add(
                u64::try_from(prefix_bytes).map_err(|_error| SlideTextError::InvalidSource)?,
            )
            .ok_or(SlideTextError::InvalidSource)?;
        if framed_header_bytes != object.header_length
            || object
                .header_offset
                .checked_add(object.header_length)
                .ok_or(SlideTextError::InvalidSource)?
                != object.data_offset
        {
            return Err(SlideTextError::InvalidSource);
        }
    }
    Ok(())
}

fn storage_rewrite_limits(
    package: &Package,
    max_message_bytes: usize,
) -> Result<RewriteLimits, SlideTextError> {
    let defaults = RewriteLimits::default();
    let message_bytes = defaults.max_message_bytes().min(max_message_bytes);
    RewriteLimits::new(
        message_bytes,
        defaults.max_fields(),
        defaults.max_nesting(),
        defaults
            .max_fragments()
            .min(package.semantic_limits().max_text_fragments()),
        defaults
            .max_text_bytes()
            .min(package.semantic_limits().max_text_bytes())
            .min(message_bytes),
        defaults.max_table_entries(),
        defaults.max_object_references(),
        defaults.max_output_bytes().min(message_bytes),
        defaults.max_rewrite_work(),
    )
    .map_err(map_text_rewrite_error)
}

fn validate_replacement(replacement: &str) -> Result<(), SlideTextError> {
    if contains_dependent_marker(replacement) {
        return Err(SlideTextError::ObjectMarkerReplacement);
    }
    Ok(())
}

fn validate_consumed_text(text: &str) -> Result<(), SlideTextError> {
    if contains_dependent_marker(text) {
        return Err(SlideTextError::DependentContent);
    }
    Ok(())
}

fn contains_dependent_marker(text: &str) -> bool {
    text.contains('\u{000e}') || text.contains('\u{fffc}')
}

fn validate_span(text: &str, span: TextSpan) -> Result<Range<usize>, SlideTextError> {
    let length = text.encode_utf16().count();
    let length_position =
        TextPosition::from_utf16_index(length).map_err(|_error| SlideTextError::InvalidSource)?;
    if span.end() > length_position {
        return Err(SlideTextError::SpanOutOfBounds {
            span,
            length: length_position,
        });
    }
    Ok(utf16_to_byte_index(text, span.start())?..utf16_to_byte_index(text, span.end())?)
}

fn utf16_to_byte_index(text: &str, position: TextPosition) -> Result<usize, SlideTextError> {
    let target =
        usize::try_from(position.utf16_index()).map_err(|_error| SlideTextError::InvalidSource)?;
    if target == 0 {
        return Ok(0);
    }
    let mut units = 0usize;
    for (byte_index, character) in text.char_indices() {
        if units == target {
            return Ok(byte_index);
        }
        units = units
            .checked_add(character.len_utf16())
            .ok_or(SlideTextError::InvalidSource)?;
        if units > target {
            return Err(SlideTextError::SurrogateBoundary { position });
        }
    }
    if units == target {
        Ok(text.len())
    } else {
        Err(SlideTextError::InvalidSource)
    }
}

fn replace_utf8_range(
    source: &str,
    range: Range<usize>,
    replacement: &str,
) -> Result<String, SlideTextError> {
    let capacity = source
        .len()
        .checked_sub(range.end - range.start)
        .and_then(|length| length.checked_add(replacement.len()))
        .ok_or(SlideTextError::InvalidSource)?;
    let mut output = String::new();
    output
        .try_reserve_exact(capacity)
        .map_err(|_allocation| SlideTextError::Allocation { amount: capacity })?;
    output.push_str(
        source
            .get(..range.start)
            .ok_or(SlideTextError::InvalidSource)?,
    );
    output.push_str(replacement);
    output.push_str(
        source
            .get(range.end..)
            .ok_or(SlideTextError::InvalidSource)?,
    );
    Ok(output)
}

fn validate_candidate_text_memory(
    package: &Package,
    before: &str,
    span: TextSpan,
    replacement: &str,
) -> Result<(), SlideTextError> {
    let removed = validate_span(before, span)?;
    let observed = before
        .len()
        .checked_sub(removed.end - removed.start)
        .and_then(|length| length.checked_add(replacement.len()))
        .ok_or_else(|| text_limit_error(usize::MAX, package.semantic_limits().max_text_bytes()))?;
    let maximum = package.semantic_limits().max_text_bytes();
    if observed > maximum {
        return Err(text_limit_error(observed, maximum));
    }
    Ok(())
}

fn try_owned_text(text: &str) -> Result<String, SlideTextError> {
    let mut owned = String::new();
    owned
        .try_reserve_exact(text.len())
        .map_err(|_allocation| SlideTextError::Allocation { amount: text.len() })?;
    owned.push_str(text);
    Ok(owned)
}

fn utf16_len(text: &str) -> Result<u32, SlideTextError> {
    u32::try_from(text.encode_utf16().count()).map_err(|_error| SlideTextError::LimitExceeded {
        kind: SlideTextLimitKind::TextUnits,
        observed: u64::MAX,
        maximum: u64::from(u32::MAX),
    })
}

#[allow(
    clippy::unnecessary_wraps,
    reason = "the semantic-only source feature adds a typed unsupported-source branch"
)]
fn physical_catalog(package: &Package) -> Result<&SourceCatalog, SlideTextError> {
    #[cfg(feature = "internal-iwork-source")]
    match &package.state.source {
        PhysicalSource::Package(source) => Ok(source),
        PhysicalSource::Semantic(_) => Err(SlideTextError::UnsupportedSource),
    }
    #[cfg(not(feature = "internal-iwork-source"))]
    {
        let PhysicalSource::Package(source) = &package.state.source;
        Ok(source)
    }
}

fn map_read_error(error: ReadError) -> SlideTextError {
    match error {
        ReadError::SemanticLimit {
            kind,
            observed,
            maximum,
            ..
        } => SlideTextError::LimitExceeded {
            kind: match kind {
                SemanticLimitKind::Objects => SlideTextLimitKind::Entries,
                SemanticLimitKind::Slides => SlideTextLimitKind::Slides,
                SemanticLimitKind::References => SlideTextLimitKind::References,
                SemanticLimitKind::TextStorages => SlideTextLimitKind::TextStorages,
                SemanticLimitKind::TextFragments => SlideTextLimitKind::TextFragments,
                SemanticLimitKind::TextBytes => SlideTextLimitKind::TextBytes,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        ReadError::PayloadLimit {
            kind,
            observed,
            maximum,
            ..
        } => SlideTextError::LimitExceeded {
            kind: match kind {
                super::PayloadLimitKind::Bytes => SlideTextLimitKind::WireBytes,
                super::PayloadLimitKind::Fields => SlideTextLimitKind::WireFields,
                super::PayloadLimitKind::Nesting => SlideTextLimitKind::WireNesting,
                super::PayloadLimitKind::Work => SlideTextLimitKind::WireWork,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        ReadError::Allocation { amount, .. } => SlideTextError::Allocation { amount },
        ReadError::Archive(error) => map_archive_error(error),
        _ => SlideTextError::InvalidSource,
    }
}

fn map_archive_error(error: litchi_iwa_archive::Error) -> SlideTextError {
    match error {
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => SlideTextError::LimitExceeded {
            kind: match kind {
                litchi_iwa_archive::LimitKind::InputBytes => SlideTextLimitKind::InputBytes,
                litchi_iwa_archive::LimitKind::OutputBytes => SlideTextLimitKind::OutputBytes,
                litchi_iwa_archive::LimitKind::Entries => SlideTextLimitKind::Entries,
                litchi_iwa_archive::LimitKind::TotalBytes
                | litchi_iwa_archive::LimitKind::MetadataBytes
                | litchi_iwa_archive::LimitKind::IwaTotalBytes => SlideTextLimitKind::TotalBytes,
                _ => SlideTextLimitKind::EntryBytes,
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => {
            SlideTextError::Allocation { amount }
        },
        litchi_iwa_archive::Error::Iwa(error) => map_core_error(error),
        _ => SlideTextError::InvalidSource,
    }
}

fn map_core_error(error: litchi_iwa_core::Error) -> SlideTextError {
    match error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => SlideTextError::LimitExceeded {
            kind: match kind {
                litchi_iwa_core::LimitKind::Objects
                | litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject
                | litchi_iwa_core::LimitKind::MetadataItems
                | litchi_iwa_core::LimitKind::SnappyFrames => SlideTextLimitKind::Entries,
                litchi_iwa_core::LimitKind::MessageBytes => SlideTextLimitKind::WireBytes,
                litchi_iwa_core::LimitKind::HeaderFields => SlideTextLimitKind::WireFields,
                litchi_iwa_core::LimitKind::HeaderNesting => SlideTextLimitKind::WireNesting,
                _ => SlideTextLimitKind::EntryBytes,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            SlideTextError::Allocation { amount: requested }
        },
        _ => SlideTextError::InvalidSource,
    }
}

fn map_slide_preview_error(error: super::slide_preview::InvalidationError) -> SlideTextError {
    match error {
        super::slide_preview::InvalidationError::InvalidSource => SlideTextError::InvalidSource,
        super::slide_preview::InvalidationError::Wire(error) => map_wire_error(error),
        super::slide_preview::InvalidationError::Archive(error) => map_core_error(error),
    }
}

fn map_wire_error(error: litchi_iwa_common::Error) -> SlideTextError {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => SlideTextError::LimitExceeded {
            kind: match kind {
                litchi_iwa_common::LimitKind::InputBytes => SlideTextLimitKind::WireBytes,
                litchi_iwa_common::LimitKind::OutputBytes => SlideTextLimitKind::OutputBytes,
                litchi_iwa_common::LimitKind::Nesting => SlideTextLimitKind::WireNesting,
                litchi_iwa_common::LimitKind::RewriteWork => SlideTextLimitKind::WireWork,
                litchi_iwa_common::LimitKind::Fields => SlideTextLimitKind::WireFields,
                _ => SlideTextLimitKind::Entries,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => {
            SlideTextError::Allocation { amount }
        },
        _ => SlideTextError::InvalidSource,
    }
}

fn map_text_rewrite_error(error: litchi_iwa_text_wire::RewriteError) -> SlideTextError {
    match error {
        litchi_iwa_text_wire::RewriteError::LimitExceeded {
            resource,
            observed,
            limit,
        } => SlideTextError::LimitExceeded {
            kind: text_rewrite_limit_kind(resource),
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
        litchi_iwa_text_wire::RewriteError::InvalidLimit {
            field,
            value,
            maximum,
        } => SlideTextError::LimitExceeded {
            kind: text_rewrite_limit_kind(field),
            observed: usize_to_u64(value),
            maximum: usize_to_u64(maximum),
        },
        litchi_iwa_text_wire::RewriteError::Allocation { amount, .. } => {
            SlideTextError::Allocation { amount }
        },
        _ => SlideTextError::InvalidSource,
    }
}

fn text_rewrite_limit_kind(resource: &str) -> SlideTextLimitKind {
    match resource {
        "text bytes" => SlideTextLimitKind::TextBytes,
        "text fragments" => SlideTextLimitKind::TextFragments,
        "object references" | "object reference count" => SlideTextLimitKind::References,
        "table entries" | "table entry count" => SlideTextLimitKind::Entries,
        "output bytes" | "rewritten storage bytes" | "appended text field" => {
            SlideTextLimitKind::OutputBytes
        },
        "nesting" => SlideTextLimitKind::WireNesting,
        "rewrite work" | "aggregate nested scan bytes" => SlideTextLimitKind::WireWork,
        "fields" | "aggregate field count" | "generated text fields" | "output field bound" => {
            SlideTextLimitKind::WireFields
        },
        _ => SlideTextLimitKind::WireBytes,
    }
}

fn text_limit_error(observed: usize, maximum: usize) -> SlideTextError {
    SlideTextError::LimitExceeded {
        kind: SlideTextLimitKind::TextBytes,
        observed: usize_to_u64(observed),
        maximum: usize_to_u64(maximum),
    }
}

fn usize_to_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}

fn fingerprint(bytes: &[u8]) -> u64 {
    bytes.iter().fold(0xcbf2_9ce4_8422_2325_u64, |value, byte| {
        (value ^ u64::from(*byte)).wrapping_mul(0x0000_0100_0000_01b3)
    })
}
