//! Exact-source transactions for Pages section body text.

use std::collections::HashSet;
use std::fmt;
use std::num::NonZeroU64;
use std::ops::Range;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_text::{TextPosition, TextSpan};
use thiserror::Error as ThisError;

use super::section_transaction as transaction;
use super::{
    MAX_SECTIONS, NativeSectionReference, Package, PackageError, StorageWireLimitsError,
    decode_body_storage, effective_text_limit, find_object, native_section_references,
    root_references_with_limits, storage_rewrite_limits,
};
use crate::{SectionSelector, section::settings::Path};

/// A finite resource governed while section text is rewritten or published.
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
            Self::TextBytes => "text bytes",
            Self::TextUnits => "text UTF-16 units",
            Self::WireBytes => "wire bytes",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting",
            Self::WireWork => "wire work",
        })
    }
}

/// An error raised while reading, staging, or publishing section body text.
///
/// Diagnostics deliberately omit authored text, native identifiers, member
/// names, paths, raw bytes, and lower-layer error strings.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies it
    /// (record 0351 §2.4, record 0361).
    #[error(transparent)]
    Refused(Refusal),
    /// The selected native body cannot be decoded or rewritten safely.
    #[error("the Pages source cannot be edited safely")]
    InvalidSource,
    /// A finite transaction resource ceiling was exceeded.
    #[error("Pages section-text {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// Resource category that exceeded its limit.
        kind: LimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded destination allocation failed before publication.
    #[error("could not allocate {amount} units for the Pages section-text transaction")]
    Allocation {
        /// Elements or bytes requested.
        amount: usize,
    },
    /// Full semantic readback did not reproduce the requested change.
    #[error("the edited Pages section text failed semantic verification")]
    Verification,
    /// The patch does not belong to this exact immutable package artifact.
    #[error("the Pages section-text patch does not match the exact source package")]
    PatchConflict,
}

/// A well-formed request this family's route refuses, classified by
/// [`Refusal::kind`] (record 0351 §2.4, record 0361).
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// More than one section matched an exact-name selector.
    #[error(
        "the Pages section-text selector is ambiguous at positions {first:?} and {duplicate:?}"
    )]
    AmbiguousSelector {
        /// First matching source position.
        first: Position,
        /// Next matching source position.
        duplicate: Position,
    },
    /// No section matched an exact-name selector.
    #[error("the Pages section-text selector did not match a section")]
    NameNotFound,
    /// No section exists at the requested source position.
    #[error("the Pages section position {position:?} does not exist")]
    PositionNotFound {
        /// Requested semantic source position.
        position: Position,
    },
    /// Whole-body convenience requires exactly one semantic section.
    #[error("whole-body text editing requires exactly one Pages section; found {actual}")]
    BodySectionCount {
        /// Number of sections present in the immutable snapshot.
        actual: usize,
    },
    /// A staged span exceeds the selected section's UTF-16 length.
    #[error("the Pages section-text span exceeds the selected section length")]
    SpanOutOfBounds {
        /// Rejected section-relative span.
        span: TextSpan,
        /// Selected section length in UTF-16 code units.
        length: TextPosition,
    },
    /// A staged boundary splits a UTF-16 surrogate pair.
    #[error("the Pages section-text boundary {position:?} splits a Unicode scalar value")]
    SurrogateBoundary {
        /// Rejected section-relative boundary.
        position: TextPosition,
    },
    /// Replacement text contains a native section-break marker.
    #[error("Pages section-break markers cannot be inserted through a text transaction")]
    SectionBreakReplacement,
    /// Replacement text contains a native footnote anchor.
    #[error("Pages footnote anchors cannot be inserted through a text transaction")]
    FootnoteAnchorReplacement,
    /// Replacement text contains an inline-object replacement character.
    #[error("Pages inline-object markers cannot be inserted through a text transaction")]
    ObjectMarkerReplacement,
    /// The edit would consume content owned by another semantic capability.
    #[error("the Pages text edit intersects dependent content")]
    DependentContent,
    /// A second operation was staged on an edit that already has one.
    #[error("a Pages section-text edit accepts exactly one staged operation")]
    OperationAlreadyStaged,
    /// The snapshot has no exact physical source suitable for a changed edit.
    #[error("this Pages source does not support physical section-text edits")]
    UnsupportedSource,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::AmbiguousSelector { .. }
            | Self::SpanOutOfBounds { .. }
            | Self::SurrogateBoundary { .. } => RefusalKind::InvalidValue,
            Self::NameNotFound { .. } | Self::PositionNotFound { .. } => RefusalKind::NotFound,
            Self::BodySectionCount { .. }
            | Self::SectionBreakReplacement { .. }
            | Self::FootnoteAnchorReplacement { .. }
            | Self::ObjectMarkerReplacement { .. }
            | Self::DependentContent { .. }
            | Self::UnsupportedSource { .. } => RefusalKind::UnsupportedContent,
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
    replacement: String,
}

/// One mutable section-relative text operation staged against an immutable package.
pub struct Edit<'a> {
    source: &'a Package,
    position: Position,
    before: &'a str,
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
        selector: impl Into<SectionSelector<'selector>>,
    ) -> Result<Self, Error> {
        let position = resolve_position(source, selector)?;
        let before = section_text_at(source, position)?;
        Ok(Self {
            source,
            position,
            before,
            operation: None,
        })
    }

    /// Return the semantic section position resolved when this edit began.
    #[must_use]
    pub const fn position(&self) -> Position {
        self.position
    }

    /// Borrow the selected section text from the immutable source snapshot.
    #[must_use]
    pub const fn text(&self) -> &'a str {
        self.before
    }

    /// Return the staged section-relative span, when present.
    #[must_use]
    pub fn span(&self) -> Option<TextSpan> {
        self.operation.as_ref().map(|operation| operation.span)
    }

    /// Stage one replacement over a checked section-relative UTF-16 span.
    ///
    /// Empty spans insert text. Empty replacement text deletes a nonempty
    /// span. Boundaries must align with Unicode scalar values. A second call
    /// is rejected so every commit has one unambiguous native splice.
    ///
    /// # Errors
    ///
    /// Returns a typed error for an out-of-bounds or surrogate-splitting span,
    /// reserved native markers, dependent content, allocation failure, or an
    /// already staged operation.
    pub fn replace(&mut self, span: TextSpan, replacement: &str) -> Result<&mut Self, Error> {
        if self.operation.is_some() {
            return Err(Error::Refused(Refusal::OperationAlreadyStaged));
        }
        validate_replacement(replacement)?;
        let byte_range = validate_section_span(self.before, span)?;
        validate_consumed_text(self.before.get(byte_range).ok_or(Error::InvalidSource)?)?;
        validate_candidate_text_memory(self.source, self.position, span, replacement)?;
        let owned = try_owned_text(replacement)?;
        self.operation = Some(Operation {
            span,
            replacement: owned,
        });
        Ok(self)
    }

    /// Stage one insertion at a checked section-relative UTF-16 position.
    ///
    /// # Errors
    ///
    /// Returns the same errors as [`Self::replace`].
    pub fn insert(&mut self, position: TextPosition, text: &str) -> Result<&mut Self, Error> {
        let span = TextSpan::new(position, position).map_err(|_error| {
            // Equal boundaries are ordered by construction.
            Error::InvalidSource
        })?;
        self.replace(span, text)
    }

    /// Stage deletion of one checked section-relative UTF-16 span.
    ///
    /// # Errors
    ///
    /// Returns the same errors as [`Self::replace`].
    pub fn delete(&mut self, span: TextSpan) -> Result<&mut Self, Error> {
        self.replace(span, "")
    }

    /// Stage replacement of all selected section text.
    ///
    /// # Errors
    ///
    /// Returns the same errors as [`Self::replace`].
    pub fn set(&mut self, text: &str) -> Result<&mut Self, Error> {
        let end = TextPosition::from_utf16_index(self.before.encode_utf16().count())
            .map_err(|_error| Error::InvalidSource)?;
        let span = TextSpan::new(TextPosition::ZERO, end).map_err(|_error| Error::InvalidSource)?;
        self.replace(span, text)
    }

    /// Stage removal of all selected section text.
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::DependentContent`] when the section owns
    /// footnote or inline-object anchors, or another staging error.
    pub fn clear(&mut self) -> Result<&mut Self, Error> {
        self.set("")
    }

    /// Validate and atomically publish the staged immutable candidate.
    ///
    /// Exact semantic no-ops reuse the original source allocation, including
    /// for normalized legacy packages. Changed edits require exact ZIP
    /// provenance and are published only after complete candidate reopening
    /// and semantic/topology verification under retained limits.
    ///
    /// # Errors
    ///
    /// Returns a typed error without modifying the source package when the
    /// operation, native topology, dependency boundary, limits, preservation,
    /// or full readback check fails.
    pub fn commit(self) -> Result<Commit, Error> {
        self.source.validate().map_err(map_package_error)?;
        if section_text_at(self.source, self.position)? != self.before {
            return Err(Error::InvalidSource);
        }

        let operation = self.operation.unwrap_or_else(|| Operation {
            span: TextSpan::default(),
            replacement: String::new(),
        });
        let byte_range = validate_section_span(self.before, operation.span)?;
        validate_replacement(&operation.replacement)?;
        let selected = self
            .before
            .get(byte_range.clone())
            .ok_or(Error::InvalidSource)?;
        validate_consumed_text(selected)?;
        let candidate_text = if selected == operation.replacement.as_str() {
            None
        } else {
            Some(replace_utf8_range(
                self.before,
                byte_range,
                &operation.replacement,
            )?)
        };
        validate_candidate_text_memory(
            self.source,
            self.position,
            operation.span,
            &operation.replacement,
        )?;

        let source_bytes = self.source.state.source.shared_source();
        let source_fingerprint = fingerprint(&source_bytes);
        let before: Arc<str> = try_owned_text(self.before)?.into();
        let after: Arc<str> = candidate_text.map_or_else(|| Arc::clone(&before), Arc::from);
        let replacement_units = utf16_len(&operation.replacement)?;
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

        if before == after {
            return Ok(Commit {
                package: self.source.clone(),
                patch: Patch {
                    source_bytes: Arc::clone(&source_bytes),
                    target_bytes: source_bytes,
                    source_fingerprint,
                    target_fingerprint: source_fingerprint,
                    position: self.position,
                    span: operation.span,
                    inverse_span,
                    before,
                    after,
                },
                diagnostics: Diagnostics::unchanged(),
            });
        }

        if !self.source.state.source.source_is_exact() {
            return Err(Error::Refused(Refusal::UnsupportedSource));
        }
        let package = rewrite_package_text(
            self.source,
            self.position,
            operation.span,
            &operation.replacement,
            &after,
        )?;
        let target_bytes = package.state.source.shared_source();
        let target_fingerprint = fingerprint(&target_bytes);
        Ok(Commit {
            package,
            patch: Patch {
                source_bytes,
                target_bytes,
                source_fingerprint,
                target_fingerprint,
                position: self.position,
                span: operation.span,
                inverse_span,
                before,
                after,
            },
            diagnostics: Diagnostics::published(),
        })
    }
}

/// An exact-source-checked, reversible Pages section-text patch.
///
/// Native identifiers, member names, and exact source/target bytes remain
/// private. Fingerprints are diagnostics only; exact byte comparison
/// authorizes application.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    source_bytes: Arc<[u8]>,
    target_bytes: Arc<[u8]>,
    source_fingerprint: u64,
    target_fingerprint: u64,
    position: Position,
    span: TextSpan,
    inverse_span: TextSpan,
    before: Arc<str>,
    after: Arc<str>,
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
    /// Return the semantic source position selected by this patch.
    #[must_use]
    pub const fn position(&self) -> Position {
        self.position
    }

    /// Return the section-relative UTF-16 span replaced by this patch.
    #[must_use]
    pub const fn span(&self) -> TextSpan {
        self.span
    }

    /// Borrow the complete selected-section text required from the source.
    #[must_use]
    pub fn before(&self) -> &str {
        &self.before
    }

    /// Borrow the complete selected-section text produced by the target.
    #[must_use]
    pub fn after(&self) -> &str {
        &self.after
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

    /// Return whether the patch preserves both semantic text and exact bytes.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after
            && self.source_fingerprint == self.target_fingerprint
            && self.source_bytes.as_ref() == self.target_bytes.as_ref()
    }

    /// Return an exact reversible patch from the target back to its source.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            source_bytes: Arc::clone(&self.target_bytes),
            target_bytes: Arc::clone(&self.source_bytes),
            source_fingerprint: self.target_fingerprint,
            target_fingerprint: self.source_fingerprint,
            position: self.position,
            span: self.inverse_span,
            inverse_span: self.span,
            before: Arc::clone(&self.after),
            after: Arc::clone(&self.before),
        }
    }
}

/// Compact evidence describing one section-text commit.
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

/// The fully verified result of one immutable section-text transaction.
#[must_use = "a Pages section-text commit contains the validated package snapshot"]
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
    /// Borrow body text owned by one semantically selected section.
    ///
    /// The result excludes the native U+0004 delimiter before a following
    /// section. Selection never exposes the root body-storage identifier.
    ///
    /// # Errors
    ///
    /// Returns a typed error when selection is ambiguous or missing, or the
    /// semantic snapshot does not contain one unambiguous section storage.
    pub fn section_text<'selector>(
        &self,
        selector: impl Into<SectionSelector<'selector>>,
    ) -> Result<&str, Error> {
        let position = resolve_position(self, selector)?;
        section_text_at(self, position)
    }

    /// Replace all text owned by one semantically selected existing section.
    ///
    /// This is the one-section convenience form of
    /// [`Self::edit_section_text`]. Selector resolution, UTF-16 boundary
    /// validation, native dependency checks, exact-source preservation, and
    /// atomic publication remain owned by the section-text transaction; this
    /// method does not introduce a second rewrite path.
    ///
    /// # Errors
    ///
    /// Returns a typed selector, source, dependency, limit, preservation, or
    /// verification error without modifying this package when publication is
    /// rejected.
    pub fn set_section_text<'selector>(
        &self,
        selector: impl Into<SectionSelector<'selector>>,
        text: &str,
    ) -> Result<Commit, Error> {
        let mut edit = self.edit_section_text(selector)?;
        edit.set(text)?;
        edit.commit()
    }

    /// Remove all text owned by one semantically selected existing section.
    ///
    /// This is the one-section convenience form of
    /// [`Self::edit_section_text`]. Existing section boundaries and unrelated
    /// native records remain subject to the same exact-source transaction
    /// rules as [`Self::set_section_text`].
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::DependentContent`] when the selected
    /// section owns content that cannot be removed by this plain-text
    /// capability, or another typed transaction error before publication.
    pub fn clear_section_text<'selector>(
        &self,
        selector: impl Into<SectionSelector<'selector>>,
    ) -> Result<Commit, Error> {
        let mut edit = self.edit_section_text(selector)?;
        edit.clear()?;
        edit.commit()
    }

    /// Start one selector-first section-relative body-text edit.
    ///
    /// The selector is resolved immediately against this immutable semantic
    /// snapshot and only its typed source position is retained.
    ///
    /// # Errors
    ///
    /// Returns a typed error when selection is ambiguous or missing, or the
    /// section body cannot be represented unambiguously.
    pub fn edit_section_text<'selector>(
        &self,
        selector: impl Into<SectionSelector<'selector>>,
    ) -> Result<Edit<'_>, Error> {
        Edit::new(self, selector)
    }

    /// Start a whole-body edit when this package has exactly one section.
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::BodySectionCount`] for empty or
    /// multi-section snapshots, or another typed source error.
    pub fn edit_body_text(&self) -> Result<Edit<'_>, Error> {
        if self.sections().len() != 1 {
            return Err(Error::Refused(Refusal::BodySectionCount {
                actual: self.sections().len(),
            }));
        }
        self.edit_section_text(SectionSelector::index(0))
    }

    /// Apply an exact-source-checked section-text patch.
    ///
    /// The retained target is fully reopened and semantically verified under
    /// this package's original limits before it is published.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] unless this package is the
    /// exact immutable source captured by `patch`, or another typed error when
    /// the retained target cannot be safely published.
    pub fn apply_section_text(&self, patch: &Patch) -> Result<Commit, Error> {
        let source = &self.state.source;
        let source_bytes = source.shared_source();
        if fingerprint(source.source_bytes()) != patch.source_fingerprint
            || source.source_bytes() != patch.source_bytes.as_ref()
            || source_bytes.as_ref() != patch.source_bytes.as_ref()
        {
            return Err(Error::PatchConflict);
        }
        self.validate().map_err(map_package_error)?;
        if section_text_at(self, patch.position)? != patch.before() {
            return Err(Error::PatchConflict);
        }

        if patch.is_noop() {
            return Ok(Commit {
                package: self.clone(),
                patch: patch.clone(),
                diagnostics: Diagnostics::unchanged(),
            });
        }
        if !source.source_is_exact() || fingerprint(&patch.target_bytes) != patch.target_fingerprint
        {
            return Err(Error::PatchConflict);
        }

        let mut budget =
            transaction::TransactionBudget::new(self).map_err(map_transaction_error)?;
        budget
            .charge_limit(
                crate::section::settings::LimitKind::InputBytes,
                source.source_bytes().len(),
                Path::Package,
            )
            .map_err(map_transaction_error)?;
        budget
            .precharge_candidate_reopen(self, patch.target_bytes.len(), Path::Package)
            .map_err(map_transaction_error)?;

        let candidate_source = source
            .reopen(Arc::clone(&patch.target_bytes))
            .map_err(map_archive_error)?;
        let candidate =
            Package::from_source_catalog(candidate_source).map_err(map_package_error)?;
        budget
            .charge_candidate_scan(&candidate, Path::Package)
            .map_err(map_transaction_error)?;
        candidate.validate().map_err(map_package_error)?;
        verify_candidate(self, &candidate, patch.position, patch.after())?;
        let source_body = native_body(self)?;
        let absolute_span = absolute_body_span(self, &source_body, patch.position, patch.span)?;
        let replacement_units = patch_replacement_units(patch)?;
        verify_native_topology(
            self,
            &candidate,
            patch.position,
            absolute_span,
            replacement_units,
        )?;
        let target = transaction::resolve_body_target(self, patch.position, &mut budget)
            .map_err(map_transaction_error)?;
        validate_mutation_component(self, target.component_index)?;
        verify_native_locality(self, &candidate, target, &mut budget)?;
        Ok(Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics::published(),
        })
    }
}

#[derive(Debug)]
pub(super) struct NativeBodySnapshot {
    pub(super) identifier: NonZeroU64,
    references: Vec<NativeSectionReference>,
    pub(super) utf16_len: usize,
}

impl NativeBodySnapshot {
    /// The native identifier of section `position` (record 0312).
    pub(super) fn section_identifier(&self, position: usize) -> Option<u64> {
        self.references
            .get(position)
            .map(|reference| reference.identifier.get())
    }

    /// The UTF-16 boundary where section `position` starts.
    pub(super) fn section_boundary(&self, position: usize) -> Option<usize> {
        self.references
            .get(position)
            .and_then(|reference| usize::try_from(reference.character_index).ok())
    }
}

const ROOT_COMPONENT_NAME: &str = "Index/Document.iwa";
const ROOT_OBJECT_IDENTIFIER: u64 = 1;
const ROOT_MESSAGE_TYPE: u32 = 10_000;
const ROOT_BODY_FIELD: &[u32] = &[4];
const DRAWABLES_ZORDER_MESSAGE_TYPE: u32 = 10_015;

/// Prove that the body selected by the rooted `DocumentArchive` reference is
/// not also owned by another physical record.  Pages' generic ingress keeps
/// archive-header references as metadata rather than treating them as the
/// semantic source of truth, so real producers may omit the aggregate or
/// field-local declaration.  When a declaration is present, however, it must
/// be singular and attributable to the root's `body_storage` field.  A second
/// owner would make a text-only rewrite mutate another capability's graph.
fn validate_rooted_body_ownership(
    package: &Package,
    body_identifier: NonZeroU64,
) -> Result<(), Error> {
    let body = body_identifier.get();
    let mut root_seen = false;
    let components = package.state.source.components();
    let anchored = anchored_descendants(package, body)?;
    for component in components.iter() {
        for object in &component.archive().objects {
            let identifier = object.archive_info.identifier.ok_or(Error::InvalidSource)?;
            let is_root =
                component.name() == ROOT_COMPONENT_NAME && identifier == ROOT_OBJECT_IDENTIFIER;
            if is_root {
                if std::mem::replace(&mut root_seen, true) {
                    return Err(Error::InvalidSource);
                }
            }
            if object.archive_info.message_infos.len() != object.messages.len() {
                return Err(Error::InvalidSource);
            }

            // A storage may declare a self-edge in its aggregate archive
            // metadata. That edge does not create another owner of the
            // selected storage and is retained byte-for-byte by the rewrite.
            // Cross-object edges still participate in the ownership census.
            if identifier == body {
                continue;
            }

            for (message_index, info) in object.archive_info.message_infos.iter().enumerate() {
                let aggregate_count = info
                    .object_references
                    .iter()
                    .filter(|candidate| **candidate == body)
                    .count();
                let data_reference = info.data_references.contains(&body);
                let mut field_count = 0usize;
                for field in &info.field_infos {
                    let occurrences = field
                        .object_references
                        .iter()
                        .filter(|candidate| **candidate == body)
                        .count();
                    if occurrences != 0 {
                        field_count = field_count
                            .checked_add(occurrences)
                            .ok_or(Error::InvalidSource)?;
                        if !is_root
                            || field.path.as_slice() != ROOT_BODY_FIELD
                            || field.data_references.contains(&body)
                        {
                            return Err(Error::Refused(Refusal::DependentContent));
                        }
                    }
                    if field.data_references.contains(&body) {
                        return Err(if is_root {
                            Error::InvalidSource
                        } else {
                            Error::Refused(Refusal::DependentContent)
                        });
                    }
                }

                if aggregate_count == 0 && field_count == 0 && !data_reference {
                    continue;
                }
                let message_type = object
                    .messages
                    .get(message_index)
                    .map(|message| message.type_)
                    .ok_or(Error::InvalidSource)?;
                // Pages includes the rooted body storage in the document's
                // drawables z-order. This exact aggregate-only edge orders the
                // storage but does not create a second text owner, and the
                // transaction preserves both the identifier and z-order bytes.
                if aggregate_count == 1
                    && field_count == 0
                    && !data_reference
                    && message_type == DRAWABLES_ZORDER_MESSAGE_TYPE
                {
                    continue;
                }
                // An object anchored in the body (its attachment, or the
                // drawable the attachment holds) declares its parent body
                // storage as one aggregate edge. The parent reference holds
                // no text position, and the rewrite preserves it and every
                // attachment's identity (record 0299).
                if aggregate_count == 1
                    && field_count == 0
                    && !data_reference
                    && anchored.contains(&identifier)
                {
                    continue;
                }
                if !is_root
                    || message_index >= object.messages.len()
                    || message_type != ROOT_MESSAGE_TYPE
                {
                    return Err(Error::Refused(Refusal::DependentContent));
                }
                if data_reference || aggregate_count > 1 || field_count > 1 {
                    return Err(Error::InvalidSource);
                }
                // A field-local declaration is only meaningful when the
                // aggregate declaration also contains the same edge.  The
                // inverse (aggregate-only) form is emitted by older Pages
                // producers and remains source-authoritative.
                if field_count != 0 && aggregate_count != 1 {
                    return Err(Error::InvalidSource);
                }
            }
        }
    }
    if !root_seen {
        return Err(Error::InvalidSource);
    }
    Ok(())
}

/// The objects the body storage anchors: those its own records reference
/// (attachments), and those the attachments reference (the anchored
/// drawables).
fn anchored_descendants(package: &Package, body: u64) -> Result<HashSet<u64>, Error> {
    let components = package.state.source.components();
    let references = |identifier: u64| {
        find_object(components, identifier).map(|object| {
            object
                .archive_info
                .message_infos
                .iter()
                .flat_map(|info| info.object_references.iter().copied())
                .filter(|reference| *reference != identifier)
                .collect::<Vec<_>>()
        })
    };
    let mut anchored = HashSet::new();
    let direct = references(body).unwrap_or_default();
    anchored
        .try_reserve(direct.len())
        .map_err(|_error| Error::Allocation {
            amount: direct.len(),
        })?;
    for child in direct {
        if child == body || !anchored.insert(child) {
            continue;
        }
        for grandchild in references(child).unwrap_or_default() {
            if grandchild != body {
                anchored.insert(grandchild);
            }
        }
    }
    Ok(anchored)
}

fn validate_unique_section_targets(references: &[NativeSectionReference]) -> Result<(), Error> {
    let mut identifiers = HashSet::new();
    identifiers
        .try_reserve(references.len())
        .map_err(|_error| Error::Allocation {
            amount: references.len(),
        })?;
    for reference in references {
        if !identifiers.insert(reference.identifier) {
            return Err(Error::InvalidSource);
        }
    }
    Ok(())
}

fn resolve_position<'selector>(
    package: &Package,
    selector: impl Into<SectionSelector<'selector>>,
) -> Result<Position, Error> {
    let selected_selector = selector.into();
    let selected = package
        .document_ref()
        .select_section(selected_selector)
        .map_err(map_selector_error)?
        .ok_or(match selected_selector {
            SectionSelector::Name(_) => Error::Refused(Refusal::NameNotFound),
            SectionSelector::Position(position) => {
                Error::Refused(Refusal::PositionNotFound { position })
            },
        })?;
    Ok(Position::new(selected.index()))
}

fn section_text_at(package: &Package, position: Position) -> Result<&str, Error> {
    let section = package
        .sections()
        .get(position.get())
        .ok_or(Error::Refused(Refusal::PositionNotFound { position }))?;
    section.body_text().ok_or(Error::InvalidSource)
}

pub(super) fn native_body(package: &Package) -> Result<NativeBodySnapshot, Error> {
    let components = package.state.source.components();
    let limits = package.state.source.limits();
    let root = root_references_with_limits(components, limits).map_err(map_package_error)?;
    let identifier = root
        .body
        .ok_or(Error::Refused(Refusal::UnsupportedSource))?;
    validate_rooted_body_ownership(package, identifier)?;
    let object = find_object(components, identifier.get()).ok_or(Error::InvalidSource)?;
    let max_text_bytes = effective_text_limit(limits);
    let (storage, table_references) = decode_body_storage(
        &object.messages,
        identifier,
        MAX_SECTIONS,
        max_text_bytes,
        limits,
    )
    .map_err(map_package_error)?;
    let references =
        native_section_references(table_references, root.initial_section, MAX_SECTIONS)
            .map_err(map_package_error)?;
    validate_unique_section_targets(&references)?;
    let utf16_len = storage.text().encode_utf16().count();
    Ok(NativeBodySnapshot {
        identifier,
        references,
        utf16_len,
    })
}

/// The attribute boundaries of the body storage around one section's text,
/// in section-relative UTF-16 units: what decides whether a sequence of
/// paragraph splices of the section may publish as one (record 0314).
pub(crate) struct SectionBoundaries {
    storage: litchi_iwa_text_wire::StorageBoundaries,
    offset: usize,
}

impl SectionBoundaries {
    /// Whether splices within `hull` (section-relative, both ends inclusive)
    /// publish exactly what one splice from the first text to the last does;
    /// see [`litchi_iwa_text_wire::StorageBoundaries::admits_combined_splice`].
    pub(crate) fn admits(&self, hull: Range<usize>) -> bool {
        match (
            self.offset.checked_add(hull.start),
            self.offset.checked_add(hull.end),
        ) {
            (Some(start), Some(end)) => self.storage.admits_combined_splice(start..end),
            _ => false,
        }
    }
}

/// Read [`SectionBoundaries`] for section `position` from the body storage.
/// The first section begins the body; a later one's offset is read from the
/// native section references.
pub(crate) fn section_boundaries(
    package: &Package,
    position: Position,
) -> Result<SectionBoundaries, Error> {
    let offset = if position.get() == 0 {
        0
    } else {
        section_start(package, position)?
    };
    let mut budget = transaction::TransactionBudget::new(package).map_err(map_transaction_error)?;
    let target = transaction::resolve_body_target(package, position, &mut budget)
        .map_err(map_transaction_error)?;
    let payload = transaction::selected_payload(package, target).map_err(map_transaction_error)?;
    let limits = storage_rewrite_limits(package.state.source.limits())
        .map_err(map_storage_wire_limits_error)?;
    let storage = litchi_iwa_text_wire::storage_boundaries(payload, limits)
        .map_err(map_text_rewrite_error)?;
    Ok(SectionBoundaries { storage, offset })
}

/// The UTF-16 body offset where section `position` starts.
pub(crate) fn section_start(package: &Package, position: Position) -> Result<usize, Error> {
    let body = native_body(package)?;
    Ok(native_section_range(package, &body, position)?.start)
}

pub(super) fn native_section_range(
    package: &Package,
    body: &NativeBodySnapshot,
    position: Position,
) -> Result<Range<usize>, Error> {
    if package.sections().get(position.get()).is_none() {
        return Err(Error::Refused(Refusal::PositionNotFound { position }));
    }
    let total = body.utf16_len;
    if body.references.is_empty() {
        if package.sections().len() != 1 || position.get() != 0 {
            return Err(Error::InvalidSource);
        }
        return Ok(0..total);
    }
    if body.references.len() != package.sections().len() {
        return Err(Error::InvalidSource);
    }
    let start = body
        .references
        .get(position.get())
        .map(|reference| usize::try_from(reference.character_index))
        .transpose()
        .map_err(|_error| Error::InvalidSource)?
        .ok_or(Error::InvalidSource)?;
    let end = if let Some(next) = body.references.get(position.get() + 1) {
        let boundary =
            usize::try_from(next.character_index).map_err(|_error| Error::InvalidSource)?;
        boundary.checked_sub(1).ok_or(Error::InvalidSource)?
    } else {
        total
    };
    if start > end || end > total {
        return Err(Error::InvalidSource);
    }
    Ok(start..end)
}

fn rewrite_package_text(
    source: &Package,
    position: Position,
    span: TextSpan,
    replacement: &str,
    expected: &str,
) -> Result<Package, Error> {
    let source_body = native_body(source)?;
    let absolute_span = absolute_body_span(source, &source_body, position, span)?;
    let absolute_start = absolute_span.start;
    let absolute_end = absolute_span.end;

    let mut budget = transaction::TransactionBudget::new(source).map_err(map_transaction_error)?;
    let target = transaction::resolve_body_target(source, position, &mut budget)
        .map_err(map_transaction_error)?;
    if target.identifier != source_body.identifier {
        return Err(Error::InvalidSource);
    }
    validate_mutation_component(source, target.component_index)?;
    let payload = transaction::selected_payload(source, target).map_err(map_transaction_error)?;

    let rewrite_limits = storage_rewrite_limits(source.state.source.limits())
        .map_err(map_storage_wire_limits_error)?;
    let path = Path::section(position);
    budget
        .charge_limit(
            crate::section::settings::LimitKind::InputBytes,
            source.state.source.source_bytes().len(),
            Path::Package,
        )
        .map_err(map_transaction_error)?;
    let residual_limits = budget
        .residual_storage_limits(rewrite_limits, path)
        .map_err(map_transaction_error)?;
    let prepared = litchi_iwa_text_wire::prepare_storage_text_rewrite_with_behavior(
        payload,
        absolute_start..absolute_end,
        replacement,
        litchi_iwa_text_wire::RewriteBehavior::PreserveOnEqualText,
        residual_limits,
    )
    .map_err(map_text_rewrite_error)?;
    let prepare_report = prepared.prepare_report();
    budget
        .charge_text_prepare(prepare_report, path)
        .map_err(map_transaction_error)?;
    let requirements = prepared.execution_requirements();
    budget
        .charge_text_requirements(requirements, path)
        .map_err(map_transaction_error)?;
    let rewritten = prepared
        .execute(litchi_iwa_text_wire::StorageRewriteExecutionLimits {
            max_output_bytes: requirements.output_bytes(),
            max_retained_elements: requirements.retained_elements(),
            max_retained_bytes: requirements.retained_bytes(),
            max_peak_scratch_bytes: requirements.peak_scratch_bytes(),
            max_allocations: requirements.allocations(),
            max_work: requirements.work(),
        })
        .map_err(map_text_rewrite_error)?;
    let report = rewritten.execution_report();
    if rewritten.bytes().len() != requirements.output_bytes()
        || report.retained_elements > requirements.retained_elements()
        || report.retained_bytes > requirements.retained_bytes()
        || report.peak_scratch_bytes > requirements.peak_scratch_bytes()
        || report.allocations > requirements.allocations()
        || report.work > requirements.work()
    {
        return Err(Error::Verification);
    }
    let removed_units = absolute_end
        .checked_sub(absolute_start)
        .ok_or(Error::InvalidSource)?;
    let replacement_units = replacement.encode_utf16().count();
    let expected_after_units = source_body
        .utf16_len
        .checked_sub(removed_units)
        .and_then(|length| length.checked_add(replacement_units))
        .ok_or(Error::InvalidSource)?;
    if rewritten.before_utf16_len() != source_body.utf16_len
        || rewritten.after_utf16_len() != expected_after_units
    {
        return Err(Error::Verification);
    }
    if !rewritten.removed_object_references_by_field().is_empty()
        || !rewritten.removed_object_references().is_empty()
        || rewritten.object_reference_occurrences_before()
            != rewritten.object_reference_occurrences_after()
    {
        return Err(Error::Refused(Refusal::DependentContent));
    }
    if !rewritten.changed() {
        return Err(Error::Verification);
    }
    let rewritten_payload = rewritten.into_bytes();
    let (candidate, stats) =
        transaction::rewrite_package(source, target, rewritten_payload, false, &mut budget)
            .map_err(map_transaction_error)?;
    if stats.touched_components != 1
        || stats.deleted_previews != 0
        || stats.source_layout_state.is_some()
        || stats.target_layout_state.is_some()
        || stats.source_preview_count != stats.target_preview_count
    {
        return Err(Error::Verification);
    }
    verify_candidate(source, &candidate, position, expected)?;
    verify_native_topology(
        source,
        &candidate,
        position,
        absolute_start..absolute_end,
        usize::try_from(utf16_len(replacement)?).map_err(|_error| Error::InvalidSource)?,
    )?;
    verify_native_locality(source, &candidate, target, &mut budget)?;
    budget.settle_transaction_reservation();
    Ok(candidate)
}

fn absolute_body_span(
    package: &Package,
    body: &NativeBodySnapshot,
    position: Position,
    span: TextSpan,
) -> Result<Range<usize>, Error> {
    let section_range = native_section_range(package, body, position)?;
    let start = section_range
        .start
        .checked_add(
            usize::try_from(span.start().utf16_index()).map_err(|_error| Error::InvalidSource)?,
        )
        .ok_or(Error::InvalidSource)?;
    let end = section_range
        .start
        .checked_add(
            usize::try_from(span.end().utf16_index()).map_err(|_error| Error::InvalidSource)?,
        )
        .ok_or(Error::InvalidSource)?;
    if start > end || end > section_range.end {
        return Err(Error::InvalidSource);
    }
    Ok(start..end)
}

fn patch_replacement_units(patch: &Patch) -> Result<usize, Error> {
    let before_units = patch.before.encode_utf16().count();
    let after_units = patch.after.encode_utf16().count();
    let start =
        usize::try_from(patch.span.start().utf16_index()).map_err(|_error| Error::InvalidSource)?;
    let end =
        usize::try_from(patch.span.end().utf16_index()).map_err(|_error| Error::InvalidSource)?;
    if end > before_units || start > end {
        return Err(Error::InvalidSource);
    }
    let removed = end - start;
    after_units
        .checked_sub(
            before_units
                .checked_sub(removed)
                .ok_or(Error::InvalidSource)?,
        )
        .ok_or(Error::InvalidSource)
}

/// Changed text is only authorized for the canonical rooted body member.
/// Read-only semantic projection may accept normalized/legacy sources, but a
/// preserve-mode write must not rediscover a member through a ZIP alias,
/// duplicate name, opaque entry, or mismatched local/central name metadata.
fn validate_mutation_component(package: &Package, component_index: usize) -> Result<(), Error> {
    let component = package
        .state
        .source
        .components()
        .iter()
        .nth(component_index)
        .ok_or(Error::InvalidSource)?;
    if component.name() != ROOT_COMPONENT_NAME {
        return Err(Error::InvalidSource);
    }
    let mut matches = package
        .state
        .source
        .package()
        .iter()
        .filter(|entry| entry.name() == ROOT_COMPONENT_NAME);
    let entry = matches.next().ok_or(Error::InvalidSource)?;
    if matches.next().is_some()
        || entry.is_opaque()
        || entry.raw_name() != ROOT_COMPONENT_NAME.as_bytes()
        || entry.metadata().local().name() != ROOT_COMPONENT_NAME.as_bytes()
        || entry.metadata().central().name() != ROOT_COMPONENT_NAME.as_bytes()
        || entry.metadata().local().compression_method()
            != entry.metadata().central().compression_method()
    {
        return Err(Error::InvalidSource);
    }
    Ok(())
}

/// Verify that a section-text rewrite changed only the selected body message.
///
/// Archive offsets, message lengths, and ZIP record offsets may legitimately
/// move when the text payload changes.  All semantic/header metadata remains
/// source-authoritative; the selected message is the sole exception for its
/// payload and length field.  This check is deliberately performed after a
/// complete candidate reopen so a changed package cannot pass merely because
/// its selected text reads back correctly.
fn verify_native_locality(
    source: &Package,
    candidate: &Package,
    target: transaction::Target,
    budget: &mut transaction::TransactionBudget,
) -> Result<(), Error> {
    let source_catalog = &source.state.source;
    let candidate_catalog = &candidate.state.source;
    budget
        .charge_limit(
            crate::section::settings::LimitKind::TransactionWork,
            source_catalog
                .source_bytes()
                .len()
                .saturating_add(candidate_catalog.source_bytes().len()),
            Path::Package,
        )
        .map_err(map_transaction_error)?;
    if source_catalog.package().len() != candidate_catalog.package().len() {
        return Err(Error::Verification);
    }
    for (before, after) in source_catalog
        .package()
        .iter()
        .zip(candidate_catalog.package().iter())
    {
        if before.name() != after.name() {
            return Err(Error::Verification);
        }
        let is_selected = before.name()
            == source_catalog
                .components()
                .iter()
                .nth(target.component_index)
                .ok_or(Error::Verification)?
                .name();
        if !is_selected && (before.data() != after.data() || before.metadata() != after.metadata())
        {
            return Err(Error::Verification);
        }
    }

    if source_catalog.components().len() != candidate_catalog.components().len() {
        return Err(Error::Verification);
    }
    for (component_index, (before, after)) in source_catalog
        .components()
        .iter()
        .zip(candidate_catalog.components().iter())
        .enumerate()
    {
        if before.name() != after.name()
            || before.archive().objects.len() != after.archive().objects.len()
        {
            return Err(Error::Verification);
        }
        for (object_index, (before_object, after_object)) in before
            .archive()
            .objects
            .iter()
            .zip(after.archive().objects.iter())
            .enumerate()
        {
            if component_index != target.component_index || object_index != target.object_index {
                if !before_object.same_content_ignoring_offsets(after_object) {
                    return Err(Error::Verification);
                }
                continue;
            }
            let after_message = after_object
                .messages
                .get(target.message_index)
                .cloned()
                .ok_or(Error::Verification)?;
            let retained = before_object
                .messages
                .iter()
                .fold(0usize, |total, message| {
                    total.saturating_add(message.data.len())
                });
            budget
                .charge_limit(
                    crate::section::settings::LimitKind::RetainedBytes,
                    retained,
                    Path::section(target.position),
                )
                .map_err(map_transaction_error)?;
            let mut expected_object = before_object.clone();
            expected_object
                .replace_message_preserving_header(
                    target.message_index,
                    after_message,
                    source_catalog
                        .limits()
                        .effective_archive_limits()
                        .map_err(map_archive_error)?,
                )
                .map_err(map_core_error)?;
            let archive_limits = source_catalog
                .limits()
                .effective_archive_limits()
                .map_err(map_archive_error)?;
            let mut expected_archive = litchi_iwa_core::Archive::new();
            expected_archive.objects.push(expected_object);
            let expected_bytes = expected_archive
                .to_bytes(archive_limits)
                .map_err(map_core_error)?;
            budget
                .charge_limit(
                    crate::section::settings::LimitKind::TransactionWork,
                    expected_bytes.len().saturating_mul(2),
                    Path::section(target.position),
                )
                .map_err(map_transaction_error)?;
            let expected_archive = litchi_iwa_core::Archive::parse(&expected_bytes, archive_limits)
                .map_err(map_core_error)?;
            let expected_object = expected_archive
                .objects
                .first()
                .ok_or(Error::Verification)?;
            if expected_archive.objects.len() != 1
                || !expected_object.same_content_ignoring_offsets(after_object)
            {
                return Err(Error::Verification);
            }
        }
    }
    Ok(())
}

fn verify_candidate(
    source: &Package,
    candidate: &Package,
    position: Position,
    expected: &str,
) -> Result<(), Error> {
    if source.stats().total_objects() != candidate.stats().total_objects()
        || source.sections().len() != candidate.sections().len()
    {
        return Err(Error::Verification);
    }
    for (index, (before, after)) in source
        .sections()
        .iter()
        .zip(candidate.sections())
        .enumerate()
    {
        let after_text = section_text_at(candidate, Position::new(index))?;
        let expected_text = if index == position.get() {
            expected
        } else {
            section_text_at(source, Position::new(index))?
        };
        if after_text != expected_text
            || before.name() != after.name()
            || before.section_type() != after.section_type()
            || before.heading() != after.heading()
            || before.paragraphs() != after.paragraphs()
            || before.page_count() != after.page_count()
            || (index != position.get() && before.text_storages() != after.text_storages())
        {
            return Err(Error::Verification);
        }
        let selector = SectionSelector::index(index);
        if source
            .section_pagination(selector)
            .map_err(|_error| Error::InvalidSource)?
            != candidate
                .section_pagination(selector)
                .map_err(|_error| Error::Verification)?
        {
            return Err(Error::Verification);
        }
    }
    Ok(())
}

fn verify_native_topology(
    source: &Package,
    candidate: &Package,
    position: Position,
    absolute_span: Range<usize>,
    replacement_units: usize,
) -> Result<(), Error> {
    let source_root = root_references_with_limits(
        source.state.source.components(),
        source.state.source.limits(),
    )
    .map_err(map_package_error)?;
    let candidate_root = root_references_with_limits(
        candidate.state.source.components(),
        candidate.state.source.limits(),
    )
    .map_err(map_package_error)?;
    if source_root.body != candidate_root.body
        || source_root.initial_section != candidate_root.initial_section
    {
        return Err(Error::Verification);
    }
    let before = native_body(source)?;
    let after = native_body(candidate)?;
    if before.identifier != after.identifier || before.references.len() != after.references.len() {
        return Err(Error::Verification);
    }
    for (old, new) in before.references.iter().zip(&after.references) {
        if old.identifier != new.identifier {
            return Err(Error::Verification);
        }
        let expected = shifted_index(old.character_index, &absolute_span, replacement_units)?;
        if new.character_index != expected {
            return Err(Error::Verification);
        }
    }
    if candidate.sections().get(position.get()).is_none() {
        return Err(Error::Verification);
    }
    Ok(())
}

fn shifted_index(index: u32, span: &Range<usize>, replacement_units: usize) -> Result<u32, Error> {
    let index_usize = usize::try_from(index).map_err(|_error| Error::Verification)?;
    if index_usize <= span.start {
        return Ok(index);
    }
    if index_usize < span.end {
        return Err(Error::Verification);
    }
    let removed = span.end - span.start;
    let shifted = if replacement_units >= removed {
        index_usize.checked_add(replacement_units - removed)
    } else {
        index_usize.checked_sub(removed - replacement_units)
    }
    .ok_or(Error::Verification)?;
    u32::try_from(shifted).map_err(|_error| Error::Verification)
}

fn validate_replacement(replacement: &str) -> Result<(), Error> {
    if replacement.contains('\u{0004}') {
        return Err(Error::Refused(Refusal::SectionBreakReplacement));
    }
    if replacement.contains('\u{000e}') {
        return Err(Error::Refused(Refusal::FootnoteAnchorReplacement));
    }
    if replacement.contains('\u{fffc}') {
        return Err(Error::Refused(Refusal::ObjectMarkerReplacement));
    }
    Ok(())
}

fn validate_consumed_text(text: &str) -> Result<(), Error> {
    if text.contains('\u{000e}') || text.contains('\u{fffc}') {
        return Err(Error::Refused(Refusal::DependentContent));
    }
    if text.contains('\u{0004}') {
        return Err(Error::InvalidSource);
    }
    Ok(())
}

fn validate_section_span(text: &str, span: TextSpan) -> Result<Range<usize>, Error> {
    let length = text.encode_utf16().count();
    let length_position =
        TextPosition::from_utf16_index(length).map_err(|_error| Error::InvalidSource)?;
    if span.end() > length_position {
        return Err(Error::Refused(Refusal::SpanOutOfBounds {
            span,
            length: length_position,
        }));
    }
    let start = utf16_to_byte_index(text, span.start())?;
    let end = utf16_to_byte_index(text, span.end())?;
    Ok(start..end)
}

fn utf16_to_byte_index(text: &str, position: TextPosition) -> Result<usize, Error> {
    let target_index =
        usize::try_from(position.utf16_index()).map_err(|_error| Error::InvalidSource)?;
    if target_index == 0 {
        return Ok(0);
    }
    let mut units = 0usize;
    for (byte_index, character) in text.char_indices() {
        if units == target_index {
            return Ok(byte_index);
        }
        units = units
            .checked_add(character.len_utf16())
            .ok_or(Error::InvalidSource)?;
        if units > target_index {
            return Err(Error::Refused(Refusal::SurrogateBoundary { position }));
        }
    }
    if units == target_index {
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
    position: Position,
    span: TextSpan,
    replacement: &str,
) -> Result<(), Error> {
    let before = section_text_at(package, position)?;
    let removed = validate_section_span(before, span)?;
    let selected_after = before
        .len()
        .checked_sub(removed.end - removed.start)
        .and_then(|length| length.checked_add(replacement.len()))
        .ok_or_else(|| {
            text_limit_error(
                usize::MAX,
                effective_text_limit(package.state.source.limits()),
            )
        })?;
    let mut observed = selected_after;
    for index in 0..package.sections().len() {
        if index == position.get() {
            continue;
        }
        let section_length = section_text_at(package, Position::new(index))?.len();
        observed = observed.checked_add(section_length).ok_or_else(|| {
            text_limit_error(
                usize::MAX,
                effective_text_limit(package.state.source.limits()),
            )
        })?;
    }
    let maximum = effective_text_limit(package.state.source.limits());
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

#[allow(clippy::needless_pass_by_value, reason = "Result::map_err conversion")]
fn map_selector_error(selection_error: crate::SelectorError) -> Error {
    match selection_error {
        crate::SelectorError::EmptySectionName => Error::Refused(Refusal::NameNotFound),
        crate::SelectorError::AmbiguousSectionName {
            first, duplicate, ..
        } => Error::Refused(Refusal::AmbiguousSelector {
            first: Position::new(first),
            duplicate: Position::new(duplicate),
        }),
    }
}

#[allow(clippy::needless_pass_by_value, reason = "Result::map_err conversion")]
fn map_transaction_error(transaction_error: crate::section::settings::Error) -> Error {
    use crate::section::settings::{Error as SettingsError, LimitKind as SettingsLimitKind};

    match transaction_error {
        SettingsError::Refused(crate::section::settings::Refusal::AmbiguousSelector {
            first,
            duplicate,
        }) => Error::Refused(Refusal::AmbiguousSelector { first, duplicate }),
        SettingsError::Refused(crate::section::settings::Refusal::NameNotFound) => {
            Error::Refused(Refusal::NameNotFound)
        },
        SettingsError::Refused(crate::section::settings::Refusal::PositionNotFound {
            position,
        }) => Error::Refused(Refusal::PositionNotFound { position }),
        SettingsError::InvalidSettings(_)
        | SettingsError::Refused(crate::section::settings::Refusal::UnsupportedDependency {
            ..
        })
        | SettingsError::Refused(crate::section::settings::Refusal::UnsupportedSource { .. }) => {
            Error::Refused(Refusal::UnsupportedSource)
        },
        SettingsError::InvalidSource { .. } => Error::InvalidSource,
        SettingsError::Limit {
            kind,
            observed,
            maximum,
            ..
        } => Error::Limit {
            kind: match kind {
                SettingsLimitKind::InputBytes => LimitKind::InputBytes,
                SettingsLimitKind::OutputBytes => LimitKind::OutputBytes,
                SettingsLimitKind::Entries
                | SettingsLimitKind::PayloadObjects
                | SettingsLimitKind::PayloadMessages
                | SettingsLimitKind::PayloadItems
                | SettingsLimitKind::References => LimitKind::Entries,
                SettingsLimitKind::EntryBytes
                | SettingsLimitKind::PackageBytes
                | SettingsLimitKind::PayloadBytes => LimitKind::EntryBytes,
                SettingsLimitKind::TotalEntryBytes | SettingsLimitKind::TotalPayloadBytes => {
                    LimitKind::TotalBytes
                },
                SettingsLimitKind::RetainedBytes => LimitKind::TextBytes,
                SettingsLimitKind::WireInputBytes | SettingsLimitKind::WireOutputBytes => {
                    LimitKind::WireBytes
                },
                SettingsLimitKind::WireFields => LimitKind::WireFields,
                SettingsLimitKind::WireNesting => LimitKind::WireNesting,
                SettingsLimitKind::WireWork | SettingsLimitKind::TransactionWork => {
                    LimitKind::WireWork
                },
            },
            observed,
            maximum,
        },
        SettingsError::Allocation { amount } => Error::Allocation { amount },
        SettingsError::Verification { .. } => Error::Verification,
        SettingsError::PatchConflict => Error::PatchConflict,
    }
}

#[allow(clippy::needless_pass_by_value, reason = "Result::map_err conversion")]
fn map_package_error(package_error: PackageError) -> Error {
    match package_error {
        PackageError::Archive(archive_error) => map_archive_error(archive_error),
        PackageError::SectionNamesTooLarge { observed, limit } => text_limit_error(observed, limit),
        PackageError::Semantic(crate::Error::TextTooLarge { observed, limit }) => {
            text_limit_error(observed, limit)
        },
        PackageError::UnexpectedFormat { .. }
        | PackageError::NotOfficeFile
        | PackageError::MissingIdentity
        | PackageError::InvalidLanguage(_)
        | PackageError::InvalidLocale
        | PackageError::InvalidLimit { .. }
        | PackageError::Limit { .. } => Error::Refused(Refusal::UnsupportedSource),
        PackageError::PayloadLimit { observed, limit } => Error::Limit {
            kind: LimitKind::WireBytes,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
        PackageError::ObjectLimit { observed, limit } => Error::Limit {
            kind: LimitKind::Entries,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
        PackageError::Allocation { amount } => Error::Allocation { amount },
        PackageError::Io(_)
        | PackageError::Detection(_)
        | PackageError::InvalidFormat(_)
        | PackageError::Semantic(_) => Error::InvalidSource,
    }
}

#[allow(clippy::needless_pass_by_value, reason = "Result::map_err conversion")]
fn map_storage_wire_limits_error(limit_error: StorageWireLimitsError) -> Error {
    match limit_error {
        StorageWireLimitsError::Physical(physical_error) => map_archive_error(physical_error),
        StorageWireLimitsError::Wire(wire_error) => map_text_rewrite_error(wire_error),
    }
}

#[allow(clippy::needless_pass_by_value, reason = "Result::map_err conversion")]
fn map_archive_error(archive_error: litchi_iwa_archive::Error) -> Error {
    match archive_error {
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                litchi_iwa_archive::LimitKind::InputBytes => LimitKind::InputBytes,
                litchi_iwa_archive::LimitKind::OutputBytes => LimitKind::OutputBytes,
                litchi_iwa_archive::LimitKind::Entries => LimitKind::Entries,
                litchi_iwa_archive::LimitKind::MemberNameBytes
                | litchi_iwa_archive::LimitKind::MetadataBytes
                | litchi_iwa_archive::LimitKind::CompressedEntryBytes
                | litchi_iwa_archive::LimitKind::EntryBytes
                | litchi_iwa_archive::LimitKind::IwaStreamBytes => LimitKind::EntryBytes,
                litchi_iwa_archive::LimitKind::TotalBytes
                | litchi_iwa_archive::LimitKind::IwaTotalBytes => LimitKind::TotalBytes,
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => Error::Allocation { amount },
        litchi_iwa_archive::Error::Iwa(core_error) => map_core_error(core_error),
        _ => Error::InvalidSource,
    }
}

#[allow(clippy::needless_pass_by_value, reason = "Result::map_err conversion")]
fn map_core_error(core_error: litchi_iwa_core::Error) -> Error {
    match core_error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                litchi_iwa_core::LimitKind::Objects
                | litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject
                | litchi_iwa_core::LimitKind::MetadataItems => LimitKind::Entries,
                litchi_iwa_core::LimitKind::MessageBytes => LimitKind::WireBytes,
                litchi_iwa_core::LimitKind::HeaderFields => LimitKind::WireFields,
                litchi_iwa_core::LimitKind::HeaderNesting => LimitKind::WireNesting,
                litchi_iwa_core::LimitKind::ArchiveBytes
                | litchi_iwa_core::LimitKind::ObjectBytes
                | litchi_iwa_core::LimitKind::HeaderBytes
                | litchi_iwa_core::LimitKind::HeaderMemoryBytes
                | litchi_iwa_core::LimitKind::SnappyChunkBytes
                | litchi_iwa_core::LimitKind::SnappyStreamBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedChunkBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedStreamBytes
                | litchi_iwa_core::LimitKind::SnappyFrames => LimitKind::EntryBytes,
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

#[allow(clippy::needless_pass_by_value, reason = "Result::map_err conversion")]
fn map_text_rewrite_error(error: litchi_iwa_text_wire::RewriteError) -> Error {
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
        litchi_iwa_text_wire::RewriteError::ReversedRange { .. }
        | litchi_iwa_text_wire::RewriteError::RangeOutOfBounds { .. }
        | litchi_iwa_text_wire::RewriteError::SurrogateSplit { .. }
        | litchi_iwa_text_wire::RewriteError::ArithmeticOverflow { .. }
        | litchi_iwa_text_wire::RewriteError::InvalidFormat(_)
        | litchi_iwa_text_wire::RewriteError::Projection(_) => Error::InvalidSource,
        _ => invalid_text_rewrite_error(),
    }
}

const fn invalid_text_rewrite_error() -> Error {
    Error::InvalidSource
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
    let mut value = 0xcbf2_9ce4_8422_2325_u64;
    for byte in bytes {
        value ^= u64::from(*byte);
        value = value.wrapping_mul(0x0000_0100_0000_01b3);
    }
    value
}
