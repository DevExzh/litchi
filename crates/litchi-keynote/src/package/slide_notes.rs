//! Exact-source, selector-first Keynote speaker-notes text transactions.

#![allow(
    clippy::map_err_ignore,
    clippy::missing_errors_doc,
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
use litchi_iwa_common::{decode_varint_from_bytes, varint::encoded_len, wire::WireView};
use litchi_iwa_core::{Archive, RawMessage, SnappyStream};
use litchi_iwa_protos::keynote_speaker_notes_codec;
use litchi_iwa_text::{TextPosition, TextSpan, storage::Storage};
use litchi_iwa_text_wire::{RewriteBehavior, RewriteLimits};
use thiserror::Error as ThisError;

use super::{
    NOTE_MESSAGE_TYPE, Package, PhysicalSource, ReadError, SLIDE_MESSAGE_TYPE,
    STORAGE_MESSAGE_TYPE, SemanticBudget, SemanticLimitKind, SemanticPath, checked_semantic_charge,
    unique_payload,
};
use crate::{SlideSelector, SlideSelectorError};

/// A finite resource governed while speaker notes are read or rewritten.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
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

impl fmt::Display for LimitKind {
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
            Self::WireNesting => "wire nesting",
            Self::WireWork => "wire work",
        })
    }
}

/// A content-redacted failure raised by a speaker-notes transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies it
    /// (record 0351 §2.4, record 0361).
    #[error(transparent)]
    Refused(Refusal),
    /// The source graph or selected storage cannot be decoded or rewritten safely.
    #[error("the Keynote speaker-notes source cannot be edited safely")]
    InvalidSource,
    /// A finite transaction resource ceiling was exceeded.
    #[error("Keynote speaker-notes {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// Resource category that exceeded its limit.
        kind: LimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded destination allocation failed before publication.
    #[error("could not allocate {amount} units for the Keynote speaker-notes transaction")]
    Allocation {
        /// Elements or bytes requested.
        amount: usize,
    },
    /// Full semantic readback did not reproduce the requested change.
    #[error("the edited Keynote speaker notes failed semantic verification")]
    Verification,
    /// The patch does not belong to this exact immutable package artifact.
    #[error("the Keynote speaker-notes patch does not match the exact source package")]
    PatchConflict,
}

/// A well-formed request this family's route refuses, classified by
/// [`Refusal::kind`] (record 0351 §2.4, record 0361).
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// The source was prepared without an exact physical package artifact.
    #[error("this Keynote source does not support physical speaker-notes edits")]
    UnsupportedSource,
    /// An exact-name selector was ambiguous.
    #[error("the Keynote slide selector is ambiguous")]
    AmbiguousSelector,
    /// An exact-name selector was empty.
    #[error("the Keynote slide selector name cannot be empty")]
    EmptySlideName,
    /// An exact-name selector did not match.
    #[error("the Keynote show has no slide matching the requested name")]
    SlideNameNotFound,
    /// A checked semantic source position does not exist.
    #[error("the Keynote show has no slide at position {position:?}")]
    SlidePositionNotFound {
        /// Missing semantic source position.
        position: Position,
    },
    /// The selected slide has no existing writable speaker-notes graph.
    #[error("the selected Keynote slide has no writable speaker-notes storage")]
    NotesStorageNotFound,
    /// A staged span exceeds the selected notes storage's UTF-16 length.
    #[error("the Keynote speaker-notes span exceeds the selected text length")]
    SpanOutOfBounds {
        /// Rejected notes-relative span.
        span: TextSpan,
        /// Notes length in UTF-16 code units.
        length: TextPosition,
    },
    /// A staged boundary splits a UTF-16 surrogate pair.
    #[error("the Keynote speaker-notes boundary {position:?} splits a Unicode scalar value")]
    SurrogateBoundary {
        /// Rejected notes-relative boundary.
        position: TextPosition,
    },
    /// Replacement text contains an inline-object replacement character.
    #[error("Keynote inline-object markers cannot be inserted through a speaker-notes transaction")]
    ObjectMarkerReplacement,
    /// The edit would consume content owned by another semantic capability.
    #[error("the Keynote speaker-notes edit intersects dependent content")]
    DependentContent,
    /// A second operation was staged on an edit that already has one.
    #[error("a Keynote speaker-notes edit accepts exactly one staged operation")]
    OperationAlreadyStaged,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::UnsupportedSource { .. }
            | Self::ObjectMarkerReplacement { .. }
            | Self::DependentContent { .. } => RefusalKind::UnsupportedContent,
            Self::AmbiguousSelector { .. }
            | Self::EmptySlideName { .. }
            | Self::SpanOutOfBounds { .. }
            | Self::SurrogateBoundary { .. } => RefusalKind::InvalidValue,
            Self::SlideNameNotFound { .. }
            | Self::SlidePositionNotFound { .. }
            | Self::NotesStorageNotFound { .. } => RefusalKind::NotFound,
            Self::OperationAlreadyStaged { .. } => RefusalKind::UnsupportedOperation,
        }
    }
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

#[derive(Debug)]
struct Operation {
    span: TextSpan,
    replacement: Option<String>,
}

#[derive(Debug)]
struct NotesSnapshot {
    note_identifier: u64,
    storage_identifier: u64,
    text: String,
}

/// One mutable notes-relative text operation staged against an immutable package.
pub struct Edit<'a> {
    source: &'a Package,
    position: Position,
    note_identifier: u64,
    storage_identifier: u64,
    before: String,
    operation: Option<Operation>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("position", &self.position)
            .field("has_operation", &self.operation.is_some())
            .finish_non_exhaustive()
    }
}

impl<'a> Edit<'a> {
    fn new<'selector>(
        source: &'a Package,
        selector: impl Into<SlideSelector<'selector>>,
    ) -> Result<Self, Error> {
        let position = resolve_position(source, selector.into())?;
        let snapshot = notes_snapshot_at(source, position)?
            .ok_or(Error::Refused(Refusal::NotesStorageNotFound))?;
        Ok(Self {
            source,
            position,
            note_identifier: snapshot.note_identifier,
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

    /// Borrow the selected speaker notes from the immutable source snapshot.
    #[must_use]
    pub fn text(&self) -> &str {
        &self.before
    }

    /// Whether a sequence of edits of these notes, each a splice within
    /// `hull` (UTF-16 code units of [`Self::text`], both ends inclusive),
    /// publishes exactly what one splice from the text before them to the
    /// text after them publishes (record 0314). `false` asks the caller to
    /// commit the edits one at a time. Whether a later edit follows an
    /// emptied text is the caller's to check.
    ///
    /// # Errors
    ///
    /// A typed read or wire error for the notes storage.
    pub fn admits_combined_splice(&self, hull: Range<usize>) -> Result<bool, Error> {
        let package = self.source;
        let storage = package
            .required_object(self.storage_identifier, "Keynote speaker-note storage")
            .map_err(map_read_error)?;
        let (_, message) = one_message(&storage.messages, STORAGE_MESSAGE_TYPE)?;
        let archive_limits = package
            .state
            .options
            .archive()
            .effective_archive_limits()
            .map_err(map_archive_error)?;
        let limits = storage_rewrite_limits(package, archive_limits.max_message_bytes())?;
        let boundaries = litchi_iwa_text_wire::storage_boundaries(&message.data, limits)
            .map_err(map_text_rewrite_error)?;
        Ok(boundaries.admits_combined_splice(hull))
    }

    /// Return the staged notes-relative span, when present.
    #[must_use]
    pub fn span(&self) -> Option<TextSpan> {
        self.operation.as_ref().map(|operation| operation.span)
    }

    /// Stage one replacement over a checked notes-relative UTF-16 span.
    ///
    /// Empty spans insert text and an empty replacement deletes a nonempty
    /// span. Boundaries must align with Unicode scalar values. Inline-object
    /// markers cannot be introduced or consumed by this plain-text capability.
    pub fn replace(&mut self, span: TextSpan, replacement: &str) -> Result<&mut Self, Error> {
        if self.operation.is_some() {
            return Err(Error::Refused(Refusal::OperationAlreadyStaged));
        }
        validate_replacement(replacement)?;
        let range = validate_span(&self.before, span)?;
        let selected = self.before.get(range).ok_or(Error::InvalidSource)?;
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

    /// Stage insertion at a checked notes-relative UTF-16 position.
    pub fn insert(&mut self, position: TextPosition, text: &str) -> Result<&mut Self, Error> {
        let span = TextSpan::new(position, position).map_err(|_error| Error::InvalidSource)?;
        self.replace(span, text)
    }

    /// Stage deletion of one checked notes-relative UTF-16 span.
    pub fn delete(&mut self, span: TextSpan) -> Result<&mut Self, Error> {
        self.replace(span, "")
    }

    /// Stage replacement of all selected speaker notes.
    pub fn set(&mut self, text: &str) -> Result<&mut Self, Error> {
        let end = TextPosition::from_utf16_index(self.before.encode_utf16().count())
            .map_err(|_error| Error::InvalidSource)?;
        let span = TextSpan::new(TextPosition::ZERO, end).map_err(|_error| Error::InvalidSource)?;
        self.replace(span, text)
    }

    /// Stage removal of all text while retaining the existing notes graph.
    pub fn clear(&mut self) -> Result<&mut Self, Error> {
        self.set("")
    }

    /// Validate and atomically publish the staged immutable candidate.
    pub fn commit(self) -> Result<Commit, Error> {
        let current = notes_snapshot_at(self.source, self.position)?.ok_or(Error::InvalidSource)?;
        if current.note_identifier != self.note_identifier
            || current.storage_identifier != self.storage_identifier
            || current.text != self.before
        {
            return Err(Error::InvalidSource);
        }

        let operation = self.operation.unwrap_or(Operation {
            span: TextSpan::default(),
            replacement: None,
        });
        let range = validate_span(&self.before, operation.span)?;
        let selected = self.before.get(range.clone()).ok_or(Error::InvalidSource)?;
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
            .ok_or(Error::InvalidSource)?;
        let inverse_span = TextSpan::new(
            operation.span.start(),
            TextPosition::from_utf16_code_units(inverse_end),
        )
        .map_err(|_error| Error::InvalidSource)?;

        let catalog = physical_catalog(self.source)?;
        let source_bytes = catalog.shared_source();
        let source_fingerprint = fingerprint(&source_bytes);

        if operation.replacement.is_none() {
            // Moving the decoded `String` into `Arc<String>` retains its
            // fallibly allocated text buffer; the inverse shares that same
            // allocation instead of copying the complete semantic text.
            let before = Arc::new(self.before);
            let after = Arc::clone(&before);
            return Ok(Commit {
                package: self.source.snapshot(),
                patch: Patch {
                    source: Arc::clone(&source_bytes),
                    target: source_bytes,
                    source_fingerprint,
                    target_fingerprint: source_fingerprint,
                    position: self.position,
                    note_identifier: self.note_identifier,
                    storage_identifier: self.storage_identifier,
                    span: operation.span,
                    inverse_span,
                    before,
                    after,
                },
                diagnostics: Diagnostics::unchanged(),
            });
        }
        selected_slide_snapshot(self.source, self.position)?;
        if !catalog.source_is_exact() {
            return Err(Error::Refused(Refusal::UnsupportedSource));
        }

        let changed_replacement = operation
            .replacement
            .as_deref()
            .ok_or(Error::InvalidSource)?;
        let after_text = replace_utf8_range(&self.before, range, changed_replacement)?;
        let before = Arc::new(self.before);
        let after = Arc::new(after_text);

        let package = rewrite_notes(
            self.source,
            self.position,
            self.note_identifier,
            self.storage_identifier,
            operation.span,
            changed_replacement,
            after.as_str(),
        )?;
        let target = physical_catalog(&package)?.shared_source();
        let target_fingerprint = fingerprint(&target);
        Ok(Commit {
            patch: Patch {
                source: source_bytes,
                target,
                source_fingerprint,
                target_fingerprint,
                position: self.position,
                note_identifier: self.note_identifier,
                storage_identifier: self.storage_identifier,
                span: operation.span,
                inverse_span,
                before,
                after,
            },
            package,
            diagnostics: Diagnostics::published(),
        })
    }
}

/// An exact-source-checked reversible speaker-notes patch.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    source: Arc<[u8]>,
    target: Arc<[u8]>,
    source_fingerprint: u64,
    target_fingerprint: u64,
    position: Position,
    note_identifier: u64,
    storage_identifier: u64,
    span: TextSpan,
    inverse_span: TextSpan,
    before: Arc<String>,
    after: Arc<String>,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("position", &self.position)
            .field("span", &self.span)
            .finish_non_exhaustive()
    }
}

impl Patch {
    /// Return the semantic slide position selected by this patch.
    #[must_use]
    pub const fn position(&self) -> Position {
        self.position
    }

    /// Return the notes-relative UTF-16 span replaced by this patch.
    #[must_use]
    pub const fn span(&self) -> TextSpan {
        self.span
    }

    /// Borrow the complete speaker notes required from the source.
    #[must_use]
    pub fn before(&self) -> &str {
        self.before.as_str()
    }

    /// Borrow the complete speaker notes produced by the target.
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

    /// Return whether this patch preserves both semantic notes and exact bytes.
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
            note_identifier: self.note_identifier,
            storage_identifier: self.storage_identifier,
            span: self.inverse_span,
            inverse_span: self.span,
            before: Arc::clone(&self.after),
            after: Arc::clone(&self.before),
        }
    }
}

/// Compact publication evidence for one speaker-notes commit.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Diagnostics {
    changed: bool,
    touched_components: usize,
    full_reparse_performed: bool,
}

impl Diagnostics {
    const fn unchanged() -> Self {
        Self {
            changed: false,
            touched_components: 0,
            full_reparse_performed: false,
        }
    }

    const fn published() -> Self {
        Self {
            changed: true,
            touched_components: 1,
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

/// The fully verified result of one speaker-notes text transaction.
#[must_use = "a Keynote speaker-notes commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
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

    /// Consume the commit and return its package and patch.
    #[must_use]
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }

    /// Borrow the reversible exact-source patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Borrow compact publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }
}

impl Package {
    /// Read one slide's existing speaker notes without exposing native identity.
    ///
    /// `Ok(None)` means that the slide has no notes graph. `Some("")` retains
    /// the distinct state in which Keynote has an existing empty notes storage.
    pub fn slide_notes<'selector>(
        &self,
        selector: impl Into<SlideSelector<'selector>>,
    ) -> Result<Option<String>, Error> {
        let position = resolve_position(self, selector.into())?;
        Ok(notes_snapshot_at(self, position)?.map(|snapshot| snapshot.text))
    }

    /// Start one selector-first notes-relative text edit.
    pub fn edit_slide_notes<'selector>(
        &self,
        selector: impl Into<SlideSelector<'selector>>,
    ) -> Result<Edit<'_>, Error> {
        Edit::new(self, selector)
    }

    /// Apply an exact-source-checked speaker-notes patch.
    pub fn apply_slide_notes(&self, patch: &Patch) -> Result<Commit, Error> {
        let catalog = physical_catalog(self)?;
        let source_bytes = catalog.shared_source();
        // A patch produced by this exact immutable package retains the same
        // source allocation. Avoid hashing and rescanning the complete ZIP in
        // that common path; the byte comparison remains the authorization
        // fallback for an independently retained but byte-identical source.
        let exact_source = Arc::ptr_eq(&source_bytes, &patch.source)
            || (fingerprint(&source_bytes) == patch.source_fingerprint
                && source_bytes.as_ref() == patch.source.as_ref());
        if !exact_source {
            return Err(Error::PatchConflict);
        }
        let current = notes_snapshot_at(self, patch.position)?.ok_or(Error::PatchConflict)?;
        if current.note_identifier != patch.note_identifier
            || current.storage_identifier != patch.storage_identifier
            || current.text != patch.before()
        {
            return Err(Error::PatchConflict);
        }
        self.validate().map_err(map_read_error)?;
        if patch.is_noop() {
            return Ok(Commit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: Diagnostics::unchanged(),
            });
        }
        selected_slide_snapshot(self, patch.position)?;
        prove_exclusive_notes_ownership(
            self,
            patch.position,
            patch.note_identifier,
            patch.storage_identifier,
        )?;
        if !catalog.source_is_exact() || fingerprint(&patch.target) != patch.target_fingerprint {
            return Err(Error::PatchConflict);
        }
        let candidate =
            Package::from_source_with_options(Arc::clone(&patch.target), self.state.options)
                .map_err(map_read_error)?;
        verify_candidate(self, &candidate, patch.position, patch.after())?;
        Ok(Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics::published(),
        })
    }
}

fn resolve_position(package: &Package, selector: SlideSelector<'_>) -> Result<Position, Error> {
    match selector {
        SlideSelector::Position(position) => package
            .slide_record_at(position.get())
            .map_err(map_read_error)?
            .map(|_record| position)
            .ok_or(Error::Refused(Refusal::SlidePositionNotFound { position })),
        SlideSelector::Name(_) => package
            .show()
            .map_err(map_read_error)?
            .select_slide(selector)
            .map_err(map_slide_selector_error)?
            .map(|slide| Position::new(slide.index()))
            .ok_or(Error::Refused(Refusal::SlideNameNotFound)),
    }
}

fn map_slide_selector_error(error: SlideSelectorError) -> Error {
    match error {
        SlideSelectorError::EmptySlideName => Error::Refused(Refusal::EmptySlideName),
        SlideSelectorError::DuplicateSlideName { .. } => Error::Refused(Refusal::AmbiguousSelector),
    }
}

/// The note object, its text storage and the text of the notes of the
/// slide at `position`, when it has notes (record 0329: the slide-drawable
/// family styles the notes' runs).
pub(in crate::package) fn notes_identity(
    package: &Package,
    position: Position,
) -> Result<Option<(u64, u64, String)>, Error> {
    Ok(notes_snapshot_at(package, position)?.map(|snapshot| {
        (
            snapshot.note_identifier,
            snapshot.storage_identifier,
            snapshot.text,
        )
    }))
}

fn notes_snapshot_at(
    package: &Package,
    position: Position,
) -> Result<Option<NotesSnapshot>, Error> {
    let record = package
        .slide_record_at(position.get())
        .map_err(map_read_error)?
        .ok_or(Error::Refused(Refusal::SlidePositionNotFound { position }))?;
    let slide = package
        .required_object(record.slide_identifier, "Keynote slide")
        .map_err(map_read_error)?;
    let slide_payload = unique_payload(&slide.messages, &[SLIDE_MESSAGE_TYPE], "Keynote slide")
        .map_err(map_read_error)?;
    let Some(note_identifier) = decode_slide_note_reference(package, slide_payload)? else {
        return Ok(None);
    };
    let note = package
        .required_object(note_identifier, "Keynote speaker note")
        .map_err(map_read_error)?;
    let note_payload = unique_payload(&note.messages, &[NOTE_MESSAGE_TYPE], "Keynote speaker note")
        .map_err(map_read_error)?;
    let storage_identifier = decode_note_storage_reference(package, note_payload)?;
    let storage = package
        .required_object(storage_identifier, "Keynote speaker-note storage")
        .map_err(map_read_error)?;
    one_message(&storage.messages, STORAGE_MESSAGE_TYPE)?;
    let mut budget = SemanticBudget::new(package.semantic_limits());
    budget
        .charge_references(
            2,
            SemanticPath::SlideNotes {
                index: position.get(),
            },
        )
        .map_err(map_read_error)?;
    let path = SemanticPath::SlideNotes {
        index: position.get(),
    };
    let storage_payload = one_message(&storage.messages, STORAGE_MESSAGE_TYPE)?
        .1
        .data
        .as_slice();
    let text = decode_notes_storage(package, storage_payload, &mut budget, path)?.into_text();
    Ok(Some(NotesSnapshot {
        note_identifier,
        storage_identifier,
        text,
    }))
}

/// Decode speaker-notes storage through the strict schema pass and lazy text
/// projection while charging the same package-wide semantic counters as the
/// general reader. The source payload is borrowed for validation only; the
/// returned storage owns its text/runs and no native wire fields.
fn decode_notes_storage(
    package: &Package,
    payload: &[u8],
    budget: &mut SemanticBudget,
    path: SemanticPath,
) -> Result<Storage, Error> {
    let storage_count = checked_semantic_charge(
        budget.text_storages,
        1,
        SemanticLimitKind::TextStorages,
        package.semantic_limits().max_text_storages(),
        path,
    )
    .map_err(map_read_error)?;
    let archive_limits = package
        .state
        .options
        .archive()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let text_limits = storage_rewrite_limits(package, archive_limits.max_message_bytes())?;
    let validated = litchi_iwa_text_wire::decode_storage(payload, text_limits)
        .map_err(map_text_rewrite_error)?;
    let validation = validated.validation();
    budget
        .charge_fragments(validation.fragments(), path)
        .map_err(map_read_error)?;
    budget
        .charge_text(validation.utf8_len(), path)
        .map_err(map_read_error)?;
    budget.text_storages = storage_count;
    Ok(validated.into_storage())
}

fn decode_slide_note_reference(package: &Package, payload: &[u8]) -> Result<Option<u64>, Error> {
    let options = speaker_notes_decode_options(package, payload)?;
    keynote_speaker_notes_codec::decode_slide_note_reference(payload, options)
        .map_err(|_error| Error::InvalidSource)
}

fn decode_note_storage_reference(package: &Package, payload: &[u8]) -> Result<u64, Error> {
    let options = speaker_notes_decode_options(package, payload)?;
    keynote_speaker_notes_codec::decode_note_storage_reference(payload, options)
        .map_err(|_error| Error::InvalidSource)
}

fn speaker_notes_decode_options(
    package: &Package,
    payload: &[u8],
) -> Result<keynote_speaker_notes_codec::DecodeOptions, Error> {
    let limits = package.wire_limits().map_err(map_wire_error)?;
    let recursion = u32::try_from(limits.max_nesting()).map_err(|_error| Error::InvalidSource)?;
    Ok(keynote_speaker_notes_codec::DecodeOptions::new(
        payload.len().min(limits.max_input_bytes()),
        limits.max_fields(),
        limits.max_rewrite_work(),
        recursion,
    ))
}

fn rewrite_notes(
    source: &Package,
    position: Position,
    expected_note_identifier: u64,
    expected_storage_identifier: u64,
    span: TextSpan,
    replacement: &str,
    expected: &str,
) -> Result<Package, Error> {
    let snapshot = notes_snapshot_at(source, position)?.ok_or(Error::InvalidSource)?;
    if snapshot.note_identifier != expected_note_identifier
        || snapshot.storage_identifier != expected_storage_identifier
    {
        return Err(Error::InvalidSource);
    }
    prove_exclusive_notes_ownership(
        source,
        position,
        expected_note_identifier,
        expected_storage_identifier,
    )?;
    let catalog = physical_catalog(source)?;
    let mut matches = catalog.components().iter().filter(|component| {
        component
            .archive()
            .object(expected_storage_identifier)
            .is_some()
    });
    let component = matches.next().ok_or(Error::InvalidSource)?;
    if matches.next().is_some() {
        return Err(Error::InvalidSource);
    }
    let component_name = component.name();
    let entry = catalog
        .package()
        .iter()
        .find(|entry| entry.name() == component_name)
        .ok_or(Error::InvalidSource)?;
    if entry.is_opaque() {
        return Err(Error::InvalidSource);
    }

    let physical_limits = source.state.options.archive();
    let archive_limits = physical_limits
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let snappy_limits = physical_limits.snappy_limits().map_err(map_archive_error)?;
    let stream = SnappyStream::decompress(
        entry.data(),
        physical_limits.snappy_limits().map_err(map_archive_error)?,
    )
    .map_err(map_core_error)?;
    let mut archive = Archive::parse(stream.as_bytes(), archive_limits).map_err(map_core_error)?;
    validate_canonical_object_length_prefixes(stream.as_bytes(), &archive)?;
    let object = archive
        .object(expected_storage_identifier)
        .ok_or(Error::InvalidSource)?;
    let (message_index, message) = one_message(&object.messages, STORAGE_MESSAGE_TYPE)?;
    let rewrite_limits = storage_rewrite_limits(source, archive_limits.max_message_bytes())?;
    let start =
        usize::try_from(span.start().utf16_index()).map_err(|_error| Error::InvalidSource)?;
    let end = usize::try_from(span.end().utf16_index()).map_err(|_error| Error::InvalidSource)?;
    let rewrite = litchi_iwa_text_wire::rewrite_storage_text_with_behavior(
        &message.data,
        start..end,
        replacement,
        RewriteBehavior::PreserveOnEqualText,
        rewrite_limits,
    )
    .map_err(map_text_rewrite_error)?;
    if !rewrite.removed_object_references().is_empty()
        || !rewrite.removed_object_references_by_field().is_empty()
        || rewrite.object_reference_occurrences_before()
            != rewrite.object_reference_occurrences_after()
    {
        return Err(Error::Refused(Refusal::DependentContent));
    }
    if !rewrite.changed() {
        return Err(Error::Verification);
    }
    let rewritten = rewrite.into_bytes();
    archive
        .object_mut(expected_storage_identifier)
        .ok_or(Error::InvalidSource)?
        .replace_message_preserving_header(
            message_index,
            RawMessage {
                type_: STORAGE_MESSAGE_TYPE,
                data: rewritten,
            },
            archive_limits,
        )
        .map_err(map_core_error)?;
    let archive_bytes = archive.to_bytes(archive_limits).map_err(map_core_error)?;
    let compressed =
        SnappyStream::compress(&archive_bytes, snappy_limits).map_err(map_core_error)?;
    let output = catalog
        .package()
        .reassemble_to_bytes(
            &[EntryEdit::new(component_name, &compressed)],
            physical_limits,
        )
        .map_err(map_archive_error)?;
    let candidate = Package::from_source_with_options(output.into(), source.state.options)
        .map_err(map_read_error)?;
    verify_candidate(source, &candidate, position, expected)?;
    Ok(candidate)
}

fn selected_slide_snapshot(package: &Package, position: Position) -> Result<crate::Slide, Error> {
    let record = package
        .slide_record_at(position.get())
        .map_err(map_read_error)?
        .ok_or(Error::Refused(Refusal::SlidePositionNotFound { position }))?;
    let slide = package
        .required_object(record.slide_identifier, "Keynote slide")
        .map_err(map_read_error)?;
    let mut budget = SemanticBudget::new(package.semantic_limits());
    package
        .parse_slide(position.get(), slide, record.is_skipped, &mut budget)
        .map_err(map_read_error)
}

fn verify_candidate(
    source: &Package,
    candidate: &Package,
    position: Position,
    expected: &str,
) -> Result<(), Error> {
    if source.state.total_objects != candidate.state.total_objects {
        return Err(Error::Verification);
    }
    let source_notes = notes_snapshot_at(source, position)?.ok_or(Error::Verification)?;
    let candidate_notes = notes_snapshot_at(candidate, position)?.ok_or(Error::Verification)?;
    if source_notes.note_identifier != candidate_notes.note_identifier
        || source_notes.storage_identifier != candidate_notes.storage_identifier
        || candidate_notes.text != expected
    {
        return Err(Error::Verification);
    }
    prove_exclusive_notes_ownership(
        source,
        position,
        source_notes.note_identifier,
        source_notes.storage_identifier,
    )?;
    prove_exclusive_notes_ownership(
        candidate,
        position,
        candidate_notes.note_identifier,
        candidate_notes.storage_identifier,
    )?;
    let before = selected_slide_snapshot(source, position)?;
    let after = selected_slide_snapshot(candidate, position)?;
    if before.index() != after.index()
        || before.is_skipped() != after.is_skipped()
        || before.name() != after.name()
        || before.title() != after.title()
        || before.text_content() != after.text_content()
        || before.text_storages() != after.text_storages()
        || before.movies() != after.movies()
        || before.builds() != after.builds()
        || before.transition() != after.transition()
    {
        return Err(Error::Verification);
    }
    Ok(())
}

fn one_message(messages: &[RawMessage], message_type: u32) -> Result<(usize, &RawMessage), Error> {
    let mut matches = messages
        .iter()
        .enumerate()
        .filter(|(_index, message)| message.type_ == message_type);
    let item = matches.next().ok_or(Error::InvalidSource)?;
    if matches.next().is_some() {
        return Err(Error::InvalidSource);
    }
    Ok(item)
}

fn validate_canonical_object_length_prefixes(
    source: &[u8],
    archive: &Archive,
) -> Result<(), Error> {
    for object in &archive.objects {
        let offset =
            usize::try_from(object.header_offset).map_err(|_error| Error::InvalidSource)?;
        let remaining = source.get(offset..).ok_or(Error::InvalidSource)?;
        let (header_bytes, prefix_bytes) =
            decode_varint_from_bytes(remaining).map_err(|_error| Error::InvalidSource)?;
        if prefix_bytes != encoded_len(header_bytes) {
            return Err(Error::InvalidSource);
        }
        let framed_header_bytes = header_bytes
            .checked_add(u64::try_from(prefix_bytes).map_err(|_error| Error::InvalidSource)?)
            .ok_or(Error::InvalidSource)?;
        if framed_header_bytes != object.header_length
            || object
                .header_offset
                .checked_add(object.header_length)
                .ok_or(Error::InvalidSource)?
                != object.data_offset
        {
            return Err(Error::InvalidSource);
        }
    }
    Ok(())
}

fn prove_exclusive_notes_ownership(
    package: &Package,
    position: Position,
    note_identifier: u64,
    storage_identifier: u64,
) -> Result<(), Error> {
    let slide_identifier = package
        .slide_record_at(position.get())
        .map_err(map_read_error)?
        .ok_or(Error::Refused(Refusal::SlidePositionNotFound { position }))?
        .slide_identifier;
    prove_exclusive_metadata_owner(
        package,
        note_identifier,
        slide_identifier,
        SLIDE_MESSAGE_TYPE,
    )?;
    prove_exclusive_metadata_owner(
        package,
        storage_identifier,
        note_identifier,
        NOTE_MESSAGE_TYPE,
    )?;
    prove_exclusive_payload_owners(
        package,
        note_identifier,
        slide_identifier,
        storage_identifier,
    )
}

fn prove_exclusive_metadata_owner(
    package: &Package,
    target_identifier: u64,
    expected_owner_identifier: u64,
    expected_message_type: u32,
) -> Result<(), Error> {
    let mut occurrences = 0usize;
    for component in package.state.source.components().iter() {
        for object in &component.archive().objects {
            let owner_identifier = object.archive_info.identifier.ok_or(Error::InvalidSource)?;
            for message in &object.archive_info.message_infos {
                for _reference in message
                    .object_references
                    .iter()
                    .chain(
                        message
                            .field_infos
                            .iter()
                            .flat_map(|field| &field.object_references),
                    )
                    .filter(|identifier| **identifier == target_identifier)
                {
                    occurrences = occurrences.checked_add(1).ok_or(Error::InvalidSource)?;
                    if occurrences > 1
                        || owner_identifier != expected_owner_identifier
                        || message.type_ != expected_message_type
                    {
                        return Err(Error::Refused(Refusal::DependentContent));
                    }
                }
            }
        }
    }
    match occurrences {
        1 => Ok(()),
        0 => Err(Error::InvalidSource),
        _ => Err(Error::Refused(Refusal::DependentContent)),
    }
}

fn prove_exclusive_payload_owners(
    package: &Package,
    note_identifier: u64,
    expected_slide_identifier: u64,
    storage_identifier: u64,
) -> Result<(), Error> {
    let mut note_edges = 0usize;
    let mut storage_edges = 0usize;
    for component in package.state.source.components().iter() {
        for object in &component.archive().objects {
            let owner_identifier = object.archive_info.identifier.ok_or(Error::InvalidSource)?;
            for message in &object.messages {
                match message.type_ {
                    SLIDE_MESSAGE_TYPE => {
                        let edge = decode_slide_note_reference(package, &message.data)?;
                        validate_slide_note_reference_shape(package, &message.data)?;
                        if edge != Some(note_identifier) {
                            continue;
                        }
                        note_edges = note_edges.checked_add(1).ok_or(Error::InvalidSource)?;
                        if note_edges > 1 || owner_identifier != expected_slide_identifier {
                            return Err(Error::Refused(Refusal::DependentContent));
                        }
                    },
                    NOTE_MESSAGE_TYPE => {
                        validate_exact_note_payload_shape(package, &message.data)?;
                        let edge = decode_note_storage_reference(package, &message.data)?;
                        if edge != storage_identifier {
                            continue;
                        }
                        storage_edges = storage_edges.checked_add(1).ok_or(Error::InvalidSource)?;
                        if storage_edges > 1 || owner_identifier != note_identifier {
                            return Err(Error::Refused(Refusal::DependentContent));
                        }
                    },
                    _ => {},
                }
            }
        }
    }
    if note_edges == 1 && storage_edges == 1 {
        Ok(())
    } else {
        Err(Error::InvalidSource)
    }
}

fn validate_slide_note_reference_shape(package: &Package, payload: &[u8]) -> Result<(), Error> {
    let limits = package.wire_limits().map_err(map_wire_error)?;
    let view = WireView::parse(payload, limits).map_err(map_wire_error)?;
    for field in view.fields().filter(|field| field.number() == 27) {
        validate_exact_reference_shape(
            package,
            field.canonical_payload().map_err(map_wire_error)?,
        )?;
    }
    Ok(())
}

fn validate_exact_note_payload_shape(package: &Package, payload: &[u8]) -> Result<(), Error> {
    let limits = package.wire_limits().map_err(map_wire_error)?;
    let view = WireView::parse(payload, limits).map_err(map_wire_error)?;
    if view.len() != 1 {
        return Err(Error::Refused(Refusal::DependentContent));
    }
    let field = view.get(0).ok_or(Error::InvalidSource)?;
    if field.number() != 1 || field.wire_type() != 2 {
        return Err(Error::Refused(Refusal::DependentContent));
    }
    validate_exact_reference_shape(package, field.canonical_payload().map_err(map_wire_error)?)
}

fn validate_exact_reference_shape(package: &Package, payload: &[u8]) -> Result<(), Error> {
    let limits = package.wire_limits().map_err(map_wire_error)?;
    let view = WireView::parse(payload, limits).map_err(map_wire_error)?;
    for field in view.fields() {
        if !matches!(field.number(), 1..=3) || field.wire_type() != 0 {
            return Err(Error::Refused(Refusal::DependentContent));
        }
        field.validate_canonical_framing().map_err(map_wire_error)?;
    }
    Ok(())
}

pub(super) fn storage_rewrite_limits(
    package: &Package,
    max_message_bytes: usize,
) -> Result<RewriteLimits, Error> {
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

fn validate_replacement(replacement: &str) -> Result<(), Error> {
    if contains_dependent_marker(replacement) {
        return Err(Error::Refused(Refusal::ObjectMarkerReplacement));
    }
    Ok(())
}

fn validate_consumed_text(text: &str) -> Result<(), Error> {
    if contains_dependent_marker(text) {
        return Err(Error::Refused(Refusal::DependentContent));
    }
    Ok(())
}

fn contains_dependent_marker(text: &str) -> bool {
    text.contains('\u{000e}') || text.contains('\u{fffc}')
}

fn validate_span(text: &str, span: TextSpan) -> Result<Range<usize>, Error> {
    let length = text.encode_utf16().count();
    let length_position =
        TextPosition::from_utf16_index(length).map_err(|_error| Error::InvalidSource)?;
    if span.end() > length_position {
        return Err(Error::Refused(Refusal::SpanOutOfBounds {
            span,
            length: length_position,
        }));
    }
    Ok(utf16_to_byte_index(text, span.start())?..utf16_to_byte_index(text, span.end())?)
}

fn utf16_to_byte_index(text: &str, position: TextPosition) -> Result<usize, Error> {
    let target = usize::try_from(position.utf16_index()).map_err(|_error| Error::InvalidSource)?;
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
            .ok_or(Error::InvalidSource)?;
        if units > target {
            return Err(Error::Refused(Refusal::SurrogateBoundary { position }));
        }
    }
    if units == target {
        Ok(text.len())
    } else {
        Err(Error::InvalidSource)
    }
}

fn replace_utf8_range(
    source: &str,
    range: Range<usize>,
    replacement: &str,
) -> Result<String, Error> {
    let capacity = source
        .len()
        .checked_sub(range.end - range.start)
        .and_then(|length| length.checked_add(replacement.len()))
        .ok_or(Error::InvalidSource)?;
    let mut output = String::new();
    output
        .try_reserve_exact(capacity)
        .map_err(|_allocation| Error::Allocation { amount: capacity })?;
    output.push_str(source.get(..range.start).ok_or(Error::InvalidSource)?);
    output.push_str(replacement);
    output.push_str(source.get(range.end..).ok_or(Error::InvalidSource)?);
    Ok(output)
}

fn validate_candidate_text_memory(
    package: &Package,
    before: &str,
    span: TextSpan,
    replacement: &str,
) -> Result<(), Error> {
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

fn try_owned_text(text: &str) -> Result<String, Error> {
    let mut owned = String::new();
    owned
        .try_reserve_exact(text.len())
        .map_err(|_allocation| Error::Allocation { amount: text.len() })?;
    owned.push_str(text);
    Ok(owned)
}

fn utf16_len(text: &str) -> Result<u32, Error> {
    u32::try_from(text.encode_utf16().count()).map_err(|_error| Error::Limit {
        kind: LimitKind::TextUnits,
        observed: u64::MAX,
        maximum: u64::from(u32::MAX),
    })
}

#[allow(
    clippy::unnecessary_wraps,
    reason = "the semantic-only source feature adds a typed unsupported-source branch"
)]
fn physical_catalog(package: &Package) -> Result<&SourceCatalog, Error> {
    match &package.state.source {
        PhysicalSource::Package(source) => Ok(source),
        PhysicalSource::Semantic(_) => Err(Error::Refused(Refusal::UnsupportedSource)),
    }
}

fn map_read_error(error: ReadError) -> Error {
    match error {
        ReadError::SemanticLimit {
            kind,
            observed,
            maximum,
            ..
        } => Error::Limit {
            kind: match kind {
                SemanticLimitKind::Objects => LimitKind::Entries,
                SemanticLimitKind::Slides => LimitKind::Slides,
                SemanticLimitKind::References => LimitKind::References,
                SemanticLimitKind::TextStorages => LimitKind::TextStorages,
                SemanticLimitKind::TextFragments => LimitKind::TextFragments,
                SemanticLimitKind::TextBytes => LimitKind::TextBytes,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        ReadError::PayloadLimit {
            kind,
            observed,
            maximum,
            ..
        } => Error::Limit {
            kind: match kind {
                super::PayloadLimitKind::Bytes => LimitKind::WireBytes,
                super::PayloadLimitKind::Fields => LimitKind::WireFields,
                super::PayloadLimitKind::Nesting => LimitKind::WireNesting,
                super::PayloadLimitKind::Work => LimitKind::WireWork,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        ReadError::Allocation { amount, .. } => Error::Allocation { amount },
        ReadError::Archive(error) => map_archive_error(error),
        _ => Error::InvalidSource,
    }
}

fn map_archive_error(error: litchi_iwa_archive::Error) -> Error {
    match error {
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                litchi_iwa_archive::LimitKind::InputBytes => LimitKind::InputBytes,
                litchi_iwa_archive::LimitKind::OutputBytes => LimitKind::OutputBytes,
                litchi_iwa_archive::LimitKind::Entries => LimitKind::Entries,
                litchi_iwa_archive::LimitKind::TotalBytes
                | litchi_iwa_archive::LimitKind::IwaTotalBytes => LimitKind::TotalBytes,
                _ => LimitKind::EntryBytes,
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => Error::Allocation { amount },
        litchi_iwa_archive::Error::Iwa(error) => map_core_error(error),
        _ => Error::InvalidSource,
    }
}

fn map_core_error(error: litchi_iwa_core::Error) -> Error {
    match error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                litchi_iwa_core::LimitKind::MessageBytes => LimitKind::WireBytes,
                litchi_iwa_core::LimitKind::HeaderFields => LimitKind::WireFields,
                litchi_iwa_core::LimitKind::HeaderNesting => LimitKind::WireNesting,
                _ => LimitKind::EntryBytes,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            Error::Allocation { amount: requested }
        },
        _ => Error::InvalidSource,
    }
}

fn map_wire_error(error: litchi_iwa_common::Error) -> Error {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => Error::Limit {
            kind: match kind {
                litchi_iwa_common::LimitKind::InputBytes => LimitKind::WireBytes,
                litchi_iwa_common::LimitKind::OutputBytes => LimitKind::OutputBytes,
                litchi_iwa_common::LimitKind::Nesting => LimitKind::WireNesting,
                litchi_iwa_common::LimitKind::RewriteWork => LimitKind::WireWork,
                _ => LimitKind::WireFields,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

pub(super) fn map_text_rewrite_error(error: litchi_iwa_text_wire::RewriteError) -> Error {
    match error {
        litchi_iwa_text_wire::RewriteError::LimitExceeded {
            resource,
            observed,
            limit,
        } => Error::Limit {
            kind: text_rewrite_limit_kind(resource),
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
        litchi_iwa_text_wire::RewriteError::InvalidLimit {
            field,
            value,
            maximum,
        } => Error::Limit {
            kind: text_rewrite_limit_kind(field),
            observed: usize_to_u64(value),
            maximum: usize_to_u64(maximum),
        },
        litchi_iwa_text_wire::RewriteError::Allocation { amount, .. } => {
            Error::Allocation { amount }
        },
        _ => Error::InvalidSource,
    }
}

fn text_rewrite_limit_kind(resource: &str) -> LimitKind {
    match resource {
        "text bytes" => LimitKind::TextBytes,
        "nesting" => LimitKind::WireNesting,
        "rewrite work" | "aggregate nested scan bytes" => LimitKind::WireWork,
        "fields" => LimitKind::WireFields,
        "text fragments" | "table entries" | "object references" => LimitKind::Entries,
        _ => LimitKind::WireBytes,
    }
}

fn text_limit_error(observed: usize, maximum: usize) -> Error {
    Error::Limit {
        kind: LimitKind::TextBytes,
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
