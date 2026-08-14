//! Exact-source transactions for Pages section body text.

use std::fmt;
use std::num::NonZeroU64;
use std::ops::Range;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::{SourceCatalog, package::EntryEdit};
use litchi_iwa_core::{Archive, RawMessage, SnappyStream};
use litchi_iwa_text::{TextPosition, TextSpan};
use thiserror::Error;

use super::{
    MAX_SECTIONS, NativeSectionReference, Package, PackageError, StorageWireLimitsError,
    decode_body_storage, effective_text_limit, find_object, is_body_text_message_type,
    native_section_references, root_references_with_limits, storage_rewrite_limits,
};
use crate::SectionSelector;

/// A finite resource governed while section text is rewritten or published.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SectionTextLimitKind {
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

impl fmt::Display for SectionTextLimitKind {
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
#[derive(Debug, Clone, Copy, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum SectionTextError {
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
    /// The selected native body cannot be decoded or rewritten safely.
    #[error("the Pages source cannot be edited safely")]
    InvalidSource,
    /// A finite transaction resource ceiling was exceeded.
    #[error("Pages section-text {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    LimitExceeded {
        /// Resource category that exceeded its limit.
        kind: SectionTextLimitKind,
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

#[derive(Debug)]
struct Operation {
    span: TextSpan,
    replacement: String,
}

/// One mutable section-relative text operation staged against an immutable package.
pub struct SectionTextEdit<'a> {
    source: &'a Package,
    position: Position,
    before: &'a str,
    operation: Option<Operation>,
}

impl fmt::Debug for SectionTextEdit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SectionTextEdit")
            .field("position", &self.position)
            .field("has_operation", &self.operation.is_some())
            .finish_non_exhaustive()
    }
}

impl<'a> SectionTextEdit<'a> {
    fn new<'selector>(
        source: &'a Package,
        selector: impl Into<SectionSelector<'selector>>,
    ) -> Result<Self, SectionTextError> {
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
    pub const fn text(&self) -> &str {
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
    pub fn replace(
        &mut self,
        span: TextSpan,
        replacement: &str,
    ) -> Result<&mut Self, SectionTextError> {
        if self.operation.is_some() {
            return Err(SectionTextError::OperationAlreadyStaged);
        }
        validate_replacement(replacement)?;
        let byte_range = validate_section_span(self.before, span)?;
        validate_consumed_text(
            self.before
                .get(byte_range)
                .ok_or(SectionTextError::InvalidSource)?,
        )?;
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
    pub fn insert(
        &mut self,
        position: TextPosition,
        text: &str,
    ) -> Result<&mut Self, SectionTextError> {
        let span = TextSpan::new(position, position).map_err(|_error| {
            // Equal boundaries are ordered by construction.
            SectionTextError::InvalidSource
        })?;
        self.replace(span, text)
    }

    /// Stage deletion of one checked section-relative UTF-16 span.
    ///
    /// # Errors
    ///
    /// Returns the same errors as [`Self::replace`].
    pub fn delete(&mut self, span: TextSpan) -> Result<&mut Self, SectionTextError> {
        self.replace(span, "")
    }

    /// Stage replacement of all selected section text.
    ///
    /// # Errors
    ///
    /// Returns the same errors as [`Self::replace`].
    pub fn set(&mut self, text: &str) -> Result<&mut Self, SectionTextError> {
        let end = TextPosition::from_utf16_index(self.before.encode_utf16().count())
            .map_err(|_error| SectionTextError::InvalidSource)?;
        let span = TextSpan::new(TextPosition::ZERO, end)
            .map_err(|_error| SectionTextError::InvalidSource)?;
        self.replace(span, text)
    }

    /// Stage removal of all selected section text.
    ///
    /// # Errors
    ///
    /// Returns [`SectionTextError::DependentContent`] when the section owns
    /// footnote or inline-object anchors, or another staging error.
    pub fn clear(&mut self) -> Result<&mut Self, SectionTextError> {
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
    pub fn commit(self) -> Result<SectionTextCommit, SectionTextError> {
        self.source.validate().map_err(map_package_error)?;
        if section_text_at(self.source, self.position)? != self.before {
            return Err(SectionTextError::InvalidSource);
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
            .ok_or(SectionTextError::InvalidSource)?;
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
            .ok_or(SectionTextError::InvalidSource)?;
        let inverse_span = TextSpan::new(
            operation.span.start(),
            TextPosition::from_utf16_code_units(inverse_end),
        )
        .map_err(|_error| SectionTextError::InvalidSource)?;

        if before == after {
            return Ok(SectionTextCommit {
                package: self.source.snapshot(),
                patch: SectionTextPatch {
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
                diagnostics: SectionTextDiagnostics::unchanged(),
            });
        }

        if !self.source.state.source.source_is_exact() {
            return Err(SectionTextError::UnsupportedSource);
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
        Ok(SectionTextCommit {
            package,
            patch: SectionTextPatch {
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
            diagnostics: SectionTextDiagnostics::published(),
        })
    }
}

/// An exact-source-checked, reversible Pages section-text patch.
///
/// Native identifiers, member names, and exact source/target bytes remain
/// private. Fingerprints are diagnostics only; exact byte comparison
/// authorizes application.
#[derive(Clone, PartialEq, Eq)]
pub struct SectionTextPatch {
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

impl fmt::Debug for SectionTextPatch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SectionTextPatch")
            .field("position", &self.position)
            .field("span", &self.span)
            .finish_non_exhaustive()
    }
}

impl SectionTextPatch {
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
pub struct SectionTextDiagnostics {
    changed: bool,
    touched_components: usize,
    full_reparse_performed: bool,
}

impl SectionTextDiagnostics {
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
pub struct SectionTextCommit {
    package: Package,
    patch: SectionTextPatch,
    diagnostics: SectionTextDiagnostics,
}

impl SectionTextCommit {
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
    pub const fn patch(&self) -> &SectionTextPatch {
        &self.patch
    }

    /// Borrow compact publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &SectionTextDiagnostics {
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
    ) -> Result<&str, SectionTextError> {
        let position = resolve_position(self, selector)?;
        section_text_at(self, position)
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
    ) -> Result<SectionTextEdit<'_>, SectionTextError> {
        SectionTextEdit::new(self, selector)
    }

    /// Start a whole-body edit when this package has exactly one section.
    ///
    /// # Errors
    ///
    /// Returns [`SectionTextError::BodySectionCount`] for empty or
    /// multi-section snapshots, or another typed source error.
    pub fn edit_body_text(&self) -> Result<SectionTextEdit<'_>, SectionTextError> {
        if self.sections().len() != 1 {
            return Err(SectionTextError::BodySectionCount {
                actual: self.sections().len(),
            });
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
    /// Returns [`SectionTextError::PatchConflict`] unless this package is the
    /// exact immutable source captured by `patch`, or another typed error when
    /// the retained target cannot be safely published.
    pub fn apply_section_text(
        &self,
        patch: &SectionTextPatch,
    ) -> Result<SectionTextCommit, SectionTextError> {
        let source = &self.state.source;
        let source_bytes = source.shared_source();
        if fingerprint(source.source_bytes()) != patch.source_fingerprint
            || source.source_bytes() != patch.source_bytes.as_ref()
            || source_bytes.as_ref() != patch.source_bytes.as_ref()
        {
            return Err(SectionTextError::PatchConflict);
        }
        self.validate().map_err(map_package_error)?;
        if section_text_at(self, patch.position)? != patch.before() {
            return Err(SectionTextError::PatchConflict);
        }

        if patch.is_noop() {
            return Ok(SectionTextCommit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: SectionTextDiagnostics::unchanged(),
            });
        }
        if !source.source_is_exact() || fingerprint(&patch.target_bytes) != patch.target_fingerprint
        {
            return Err(SectionTextError::PatchConflict);
        }

        let candidate_source = SourceCatalog::from_shared_bytes_with_limits(
            Arc::clone(&patch.target_bytes),
            source.limits(),
        )
        .map_err(map_archive_error)?;
        let candidate =
            Package::from_source_catalog(candidate_source).map_err(map_package_error)?;
        candidate.validate().map_err(map_package_error)?;
        verify_candidate(self, &candidate, patch.position, patch.after())?;
        Ok(SectionTextCommit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: SectionTextDiagnostics::published(),
        })
    }
}

#[derive(Debug)]
struct NativeBodySnapshot {
    identifier: NonZeroU64,
    references: Vec<NativeSectionReference>,
    utf16_len: usize,
}

fn resolve_position<'selector>(
    package: &Package,
    selector: impl Into<SectionSelector<'selector>>,
) -> Result<Position, SectionTextError> {
    let selected_selector = selector.into();
    let selected = package
        .semantic_document()
        .select_section(selected_selector)
        .map_err(map_selector_error)?
        .ok_or(match selected_selector {
            SectionSelector::Name(_) => SectionTextError::NameNotFound,
            SectionSelector::Position(position) => SectionTextError::PositionNotFound { position },
        })?;
    Ok(Position::new(selected.index()))
}

fn section_text_at(package: &Package, position: Position) -> Result<&str, SectionTextError> {
    let section = package
        .sections()
        .get(position.get())
        .ok_or(SectionTextError::PositionNotFound { position })?;
    if section.heading().is_some()
        || !section.paragraphs().is_empty()
        || section.text_storages().len() != 1
    {
        return Err(SectionTextError::InvalidSource);
    }
    Ok(section.text_storages()[0].text())
}

fn native_body(package: &Package) -> Result<NativeBodySnapshot, SectionTextError> {
    let components = package.state.source.components();
    let limits = package.state.source.limits();
    let root = root_references_with_limits(components, limits).map_err(map_package_error)?;
    let identifier = root.body.ok_or(SectionTextError::UnsupportedSource)?;
    let object =
        find_object(components, identifier.get()).ok_or(SectionTextError::InvalidSource)?;
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
    let utf16_len = storage.text().encode_utf16().count();
    Ok(NativeBodySnapshot {
        identifier,
        references,
        utf16_len,
    })
}

fn native_section_range(
    package: &Package,
    body: &NativeBodySnapshot,
    position: Position,
) -> Result<Range<usize>, SectionTextError> {
    if package.sections().get(position.get()).is_none() {
        return Err(SectionTextError::PositionNotFound { position });
    }
    let total = body.utf16_len;
    if body.references.is_empty() {
        if package.sections().len() != 1 || position.get() != 0 {
            return Err(SectionTextError::InvalidSource);
        }
        return Ok(0..total);
    }
    if body.references.len() != package.sections().len() {
        return Err(SectionTextError::InvalidSource);
    }
    let start = body
        .references
        .get(position.get())
        .map(|reference| usize::try_from(reference.character_index))
        .transpose()
        .map_err(|_error| SectionTextError::InvalidSource)?
        .ok_or(SectionTextError::InvalidSource)?;
    let end = if let Some(next) = body.references.get(position.get() + 1) {
        let boundary = usize::try_from(next.character_index)
            .map_err(|_error| SectionTextError::InvalidSource)?;
        boundary
            .checked_sub(1)
            .ok_or(SectionTextError::InvalidSource)?
    } else {
        total
    };
    if start > end || end > total {
        return Err(SectionTextError::InvalidSource);
    }
    Ok(start..end)
}

fn rewrite_package_text(
    source: &Package,
    position: Position,
    span: TextSpan,
    replacement: &str,
    expected: &str,
) -> Result<Package, SectionTextError> {
    let source_body = native_body(source)?;
    let section_range = native_section_range(source, &source_body, position)?;
    let absolute_start = section_range
        .start
        .checked_add(
            usize::try_from(span.start().utf16_index())
                .map_err(|_error| SectionTextError::InvalidSource)?,
        )
        .ok_or(SectionTextError::InvalidSource)?;
    let absolute_end = section_range
        .start
        .checked_add(
            usize::try_from(span.end().utf16_index())
                .map_err(|_error| SectionTextError::InvalidSource)?,
        )
        .ok_or(SectionTextError::InvalidSource)?;
    if absolute_end > section_range.end {
        return Err(SectionTextError::InvalidSource);
    }

    let source_catalog = &source.state.source;
    let mut matching_components = source_catalog.components().iter().filter(|component| {
        component
            .archive()
            .object(source_body.identifier.get())
            .is_some()
    });
    let component = matching_components
        .next()
        .ok_or(SectionTextError::InvalidSource)?;
    if matching_components.next().is_some() {
        return Err(SectionTextError::InvalidSource);
    }
    let component_name = component.name();
    let entry = source_catalog
        .package()
        .iter()
        .find(|entry| entry.name() == component_name)
        .ok_or(SectionTextError::InvalidSource)?;
    if entry.is_opaque() {
        return Err(SectionTextError::InvalidSource);
    }

    let physical_limits = source_catalog.limits();
    let archive_limits = physical_limits
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let stream = SnappyStream::decompress_with_limits(
        entry.data(),
        physical_limits.snappy_limits().map_err(map_archive_error)?,
    )
    .map_err(map_core_error)?;
    let mut archive =
        Archive::parse_with_limits(stream.as_bytes(), archive_limits).map_err(map_core_error)?;
    let object = archive
        .object(source_body.identifier.get())
        .ok_or(SectionTextError::InvalidSource)?;
    let mut messages = object
        .messages
        .iter()
        .enumerate()
        .filter(|(_index, message)| is_body_text_message_type(message.type_));
    let (message_index, message) = messages.next().ok_or(SectionTextError::InvalidSource)?;
    if messages.next().is_some() {
        return Err(SectionTextError::InvalidSource);
    }
    let message_type = message.type_;

    let rewrite_limits = storage_rewrite_limits(source.state.source.limits())
        .map_err(map_storage_wire_limits_error)?;
    let rewritten = litchi_iwa_text_wire::rewrite_storage_text_with_limits(
        &message.data,
        absolute_start..absolute_end,
        replacement,
        rewrite_limits,
    )
    .map_err(map_text_rewrite_error)?;
    let removed_units = absolute_end
        .checked_sub(absolute_start)
        .ok_or(SectionTextError::InvalidSource)?;
    let replacement_units = replacement.encode_utf16().count();
    let expected_after_units = source_body
        .utf16_len
        .checked_sub(removed_units)
        .and_then(|length| length.checked_add(replacement_units))
        .ok_or(SectionTextError::InvalidSource)?;
    if rewritten.before_utf16_len() != source_body.utf16_len
        || rewritten.after_utf16_len() != expected_after_units
    {
        return Err(SectionTextError::Verification);
    }
    if !rewritten.removed_object_references_by_field().is_empty()
        || !rewritten.removed_object_references().is_empty()
        || rewritten.object_reference_occurrences_before()
            != rewritten.object_reference_occurrences_after()
    {
        return Err(SectionTextError::DependentContent);
    }
    if !rewritten.changed() {
        return Err(SectionTextError::Verification);
    }
    let rewritten_payload = rewritten.into_bytes();

    archive
        .object_mut(source_body.identifier.get())
        .ok_or(SectionTextError::InvalidSource)?
        .replace_message_preserving_header_with_limits(
            message_index,
            RawMessage {
                type_: message_type,
                data: rewritten_payload,
            },
            archive_limits,
        )
        .map_err(map_core_error)?;
    let rewritten_archive = archive
        .to_bytes_with_limits(archive_limits)
        .map_err(map_core_error)?;
    let compressed = SnappyStream::compress(&rewritten_archive).map_err(map_core_error)?;
    let output = source_catalog
        .package()
        .reassemble_to_bytes(
            &[EntryEdit::new(component_name, &compressed)],
            physical_limits,
        )
        .map_err(map_archive_error)?;
    let candidate_source =
        SourceCatalog::from_shared_bytes_with_limits(output.into(), physical_limits)
            .map_err(map_archive_error)?;
    let candidate = Package::from_source_catalog(candidate_source).map_err(map_package_error)?;
    candidate.validate().map_err(map_package_error)?;
    verify_candidate(source, &candidate, position, expected)?;
    verify_native_topology(
        source,
        &candidate,
        position,
        absolute_start..absolute_end,
        replacement,
    )?;
    Ok(candidate)
}

fn verify_candidate(
    source: &Package,
    candidate: &Package,
    position: Position,
    expected: &str,
) -> Result<(), SectionTextError> {
    if source.stats().total_objects() != candidate.stats().total_objects()
        || source.sections().len() != candidate.sections().len()
    {
        return Err(SectionTextError::Verification);
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
            return Err(SectionTextError::Verification);
        }
        let selector = SectionSelector::index(index);
        if source
            .section_pagination(selector)
            .map_err(|_error| SectionTextError::InvalidSource)?
            != candidate
                .section_pagination(selector)
                .map_err(|_error| SectionTextError::Verification)?
        {
            return Err(SectionTextError::Verification);
        }
    }
    Ok(())
}

fn verify_native_topology(
    source: &Package,
    candidate: &Package,
    position: Position,
    absolute_span: Range<usize>,
    replacement: &str,
) -> Result<(), SectionTextError> {
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
        return Err(SectionTextError::Verification);
    }
    let before = native_body(source)?;
    let after = native_body(candidate)?;
    if before.identifier != after.identifier || before.references.len() != after.references.len() {
        return Err(SectionTextError::Verification);
    }
    let replacement_units = usize::try_from(utf16_len(replacement)?)
        .map_err(|_error| SectionTextError::InvalidSource)?;
    for (old, new) in before.references.iter().zip(&after.references) {
        if old.identifier != new.identifier {
            return Err(SectionTextError::Verification);
        }
        let expected = shifted_index(old.character_index, &absolute_span, replacement_units)?;
        if new.character_index != expected {
            return Err(SectionTextError::Verification);
        }
    }
    if candidate.sections().get(position.get()).is_none() {
        return Err(SectionTextError::Verification);
    }
    Ok(())
}

fn shifted_index(
    index: u32,
    span: &Range<usize>,
    replacement_units: usize,
) -> Result<u32, SectionTextError> {
    let index_usize = usize::try_from(index).map_err(|_error| SectionTextError::Verification)?;
    if index_usize <= span.start {
        return Ok(index);
    }
    if index_usize < span.end {
        return Err(SectionTextError::Verification);
    }
    let removed = span.end - span.start;
    let shifted = if replacement_units >= removed {
        index_usize.checked_add(replacement_units - removed)
    } else {
        index_usize.checked_sub(removed - replacement_units)
    }
    .ok_or(SectionTextError::Verification)?;
    u32::try_from(shifted).map_err(|_error| SectionTextError::Verification)
}

fn validate_replacement(replacement: &str) -> Result<(), SectionTextError> {
    if replacement.contains('\u{0004}') {
        return Err(SectionTextError::SectionBreakReplacement);
    }
    if replacement.contains('\u{000e}') {
        return Err(SectionTextError::FootnoteAnchorReplacement);
    }
    if replacement.contains('\u{fffc}') {
        return Err(SectionTextError::ObjectMarkerReplacement);
    }
    Ok(())
}

fn validate_consumed_text(text: &str) -> Result<(), SectionTextError> {
    if text.contains('\u{000e}') || text.contains('\u{fffc}') {
        return Err(SectionTextError::DependentContent);
    }
    if text.contains('\u{0004}') {
        return Err(SectionTextError::InvalidSource);
    }
    Ok(())
}

fn validate_section_span(text: &str, span: TextSpan) -> Result<Range<usize>, SectionTextError> {
    let length = text.encode_utf16().count();
    let length_position =
        TextPosition::from_utf16_index(length).map_err(|_error| SectionTextError::InvalidSource)?;
    if span.end() > length_position {
        return Err(SectionTextError::SpanOutOfBounds {
            span,
            length: length_position,
        });
    }
    let start = utf16_to_byte_index(text, span.start())?;
    let end = utf16_to_byte_index(text, span.end())?;
    Ok(start..end)
}

fn utf16_to_byte_index(text: &str, position: TextPosition) -> Result<usize, SectionTextError> {
    let target_index = usize::try_from(position.utf16_index())
        .map_err(|_error| SectionTextError::InvalidSource)?;
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
            .ok_or(SectionTextError::InvalidSource)?;
        if units > target_index {
            return Err(SectionTextError::SurrogateBoundary { position });
        }
    }
    if units == target_index {
        Ok(text.len())
    } else {
        Err(SectionTextError::InvalidSource)
    }
}

fn replace_utf8_range(
    source: &str,
    range: Range<usize>,
    replacement: &str,
) -> Result<String, SectionTextError> {
    let capacity = source
        .len()
        .checked_sub(range.end - range.start)
        .and_then(|length| length.checked_add(replacement.len()))
        .ok_or(SectionTextError::InvalidSource)?;
    let mut output = String::new();
    output
        .try_reserve_exact(capacity)
        .map_err(|_allocation| SectionTextError::Allocation { amount: capacity })?;
    output.push_str(
        source
            .get(..range.start)
            .ok_or(SectionTextError::InvalidSource)?,
    );
    output.push_str(replacement);
    output.push_str(
        source
            .get(range.end..)
            .ok_or(SectionTextError::InvalidSource)?,
    );
    Ok(output)
}

fn validate_candidate_text_memory(
    package: &Package,
    position: Position,
    span: TextSpan,
    replacement: &str,
) -> Result<(), SectionTextError> {
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

fn try_owned_text(text: &str) -> Result<String, SectionTextError> {
    let mut owned = String::new();
    owned
        .try_reserve_exact(text.len())
        .map_err(|_allocation| SectionTextError::Allocation { amount: text.len() })?;
    owned.push_str(text);
    Ok(owned)
}

fn utf16_len(text: &str) -> Result<u32, SectionTextError> {
    u32::try_from(text.encode_utf16().count()).map_err(|_error| SectionTextError::LimitExceeded {
        kind: SectionTextLimitKind::TextUnits,
        observed: u64::MAX,
        maximum: u64::from(u32::MAX),
    })
}

#[allow(clippy::needless_pass_by_value, reason = "Result::map_err conversion")]
fn map_selector_error(selection_error: crate::SelectorError) -> SectionTextError {
    match selection_error {
        crate::SelectorError::AmbiguousSectionName {
            first, duplicate, ..
        } => SectionTextError::AmbiguousSelector {
            first: Position::new(first),
            duplicate: Position::new(duplicate),
        },
    }
}

#[allow(clippy::needless_pass_by_value, reason = "Result::map_err conversion")]
fn map_package_error(package_error: PackageError) -> SectionTextError {
    match package_error {
        PackageError::Archive(archive_error) => map_archive_error(archive_error),
        PackageError::SectionNamesTooLarge { observed, limit } => text_limit_error(observed, limit),
        PackageError::Semantic(crate::Error::TextTooLarge { observed, limit }) => {
            text_limit_error(observed, limit)
        },
        PackageError::NotPages => SectionTextError::UnsupportedSource,
        PackageError::PayloadLimit { observed, limit } => SectionTextError::LimitExceeded {
            kind: SectionTextLimitKind::WireBytes,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
        PackageError::ObjectLimit { observed, limit } => SectionTextError::LimitExceeded {
            kind: SectionTextLimitKind::Entries,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
        PackageError::Allocation { amount } => SectionTextError::Allocation { amount },
        PackageError::Io(_)
        | PackageError::Detection(_)
        | PackageError::InvalidFormat(_)
        | PackageError::Semantic(_) => SectionTextError::InvalidSource,
    }
}

#[allow(clippy::needless_pass_by_value, reason = "Result::map_err conversion")]
fn map_storage_wire_limits_error(limit_error: StorageWireLimitsError) -> SectionTextError {
    match limit_error {
        StorageWireLimitsError::Physical(physical_error) => map_archive_error(physical_error),
        StorageWireLimitsError::Wire(wire_error) => map_text_rewrite_error(wire_error),
    }
}

#[allow(clippy::needless_pass_by_value, reason = "Result::map_err conversion")]
fn map_archive_error(archive_error: litchi_iwa_archive::Error) -> SectionTextError {
    match archive_error {
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => SectionTextError::LimitExceeded {
            kind: match kind {
                litchi_iwa_archive::LimitKind::InputBytes => SectionTextLimitKind::InputBytes,
                litchi_iwa_archive::LimitKind::OutputBytes => SectionTextLimitKind::OutputBytes,
                litchi_iwa_archive::LimitKind::Entries => SectionTextLimitKind::Entries,
                litchi_iwa_archive::LimitKind::MemberNameBytes
                | litchi_iwa_archive::LimitKind::MetadataBytes
                | litchi_iwa_archive::LimitKind::CompressedEntryBytes
                | litchi_iwa_archive::LimitKind::EntryBytes
                | litchi_iwa_archive::LimitKind::IwaStreamBytes => SectionTextLimitKind::EntryBytes,
                litchi_iwa_archive::LimitKind::TotalBytes
                | litchi_iwa_archive::LimitKind::IwaTotalBytes => SectionTextLimitKind::TotalBytes,
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => {
            SectionTextError::Allocation { amount }
        },
        litchi_iwa_archive::Error::Iwa(core_error) => map_core_error(core_error),
        litchi_iwa_archive::Error::Io(_)
        | litchi_iwa_archive::Error::Zip { .. }
        | litchi_iwa_archive::Error::InvalidLimits(_)
        | litchi_iwa_archive::Error::Encrypted
        | litchi_iwa_archive::Error::SourceChanged { .. }
        | litchi_iwa_archive::Error::DirectoryChanged { .. }
        | litchi_iwa_archive::Error::Reassembly(_)
        | litchi_iwa_archive::Error::InvalidBundle(_) => SectionTextError::InvalidSource,
    }
}

#[allow(clippy::needless_pass_by_value, reason = "Result::map_err conversion")]
fn map_core_error(core_error: litchi_iwa_core::Error) -> SectionTextError {
    match core_error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => SectionTextError::LimitExceeded {
            kind: match kind {
                litchi_iwa_core::LimitKind::Objects
                | litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject
                | litchi_iwa_core::LimitKind::MetadataItems => SectionTextLimitKind::Entries,
                litchi_iwa_core::LimitKind::MessageBytes => SectionTextLimitKind::WireBytes,
                litchi_iwa_core::LimitKind::HeaderFields => SectionTextLimitKind::WireFields,
                litchi_iwa_core::LimitKind::HeaderNesting => SectionTextLimitKind::WireNesting,
                litchi_iwa_core::LimitKind::ArchiveBytes
                | litchi_iwa_core::LimitKind::ObjectBytes
                | litchi_iwa_core::LimitKind::HeaderBytes
                | litchi_iwa_core::LimitKind::HeaderMemoryBytes
                | litchi_iwa_core::LimitKind::SnappyChunkBytes
                | litchi_iwa_core::LimitKind::SnappyStreamBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedChunkBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedStreamBytes
                | litchi_iwa_core::LimitKind::SnappyFrames => SectionTextLimitKind::EntryBytes,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            SectionTextError::Allocation { amount: requested }
        },
        litchi_iwa_core::Error::InvalidArchive { .. }
        | litchi_iwa_core::Error::InvalidLimits { .. }
        | litchi_iwa_core::Error::HeaderCodec { .. }
        | litchi_iwa_core::Error::Io(_)
        | litchi_iwa_core::Error::Snappy { .. } => SectionTextError::InvalidSource,
    }
}

#[allow(clippy::needless_pass_by_value, reason = "Result::map_err conversion")]
fn map_text_rewrite_error(error: litchi_iwa_text_wire::RewriteError) -> SectionTextError {
    match error {
        litchi_iwa_text_wire::RewriteError::LimitExceeded {
            resource,
            observed,
            limit,
        } => SectionTextError::LimitExceeded {
            kind: text_rewrite_limit_kind(resource),
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
        litchi_iwa_text_wire::RewriteError::InvalidLimit {
            field,
            value,
            maximum,
        } => SectionTextError::LimitExceeded {
            kind: text_rewrite_limit_kind(field),
            observed: usize_to_u64(value),
            maximum: usize_to_u64(maximum),
        },
        litchi_iwa_text_wire::RewriteError::Allocation { amount, .. } => {
            SectionTextError::Allocation { amount }
        },
        litchi_iwa_text_wire::RewriteError::ReversedRange { .. }
        | litchi_iwa_text_wire::RewriteError::RangeOutOfBounds { .. }
        | litchi_iwa_text_wire::RewriteError::SurrogateSplit { .. }
        | litchi_iwa_text_wire::RewriteError::ArithmeticOverflow { .. }
        | litchi_iwa_text_wire::RewriteError::InvalidFormat(_)
        | litchi_iwa_text_wire::RewriteError::Projection(_) => SectionTextError::InvalidSource,
        _ => invalid_text_rewrite_error(),
    }
}

const fn invalid_text_rewrite_error() -> SectionTextError {
    SectionTextError::InvalidSource
}

fn text_rewrite_limit_kind(resource: &str) -> SectionTextLimitKind {
    match resource {
        "text bytes" => SectionTextLimitKind::TextBytes,
        "nesting" => SectionTextLimitKind::WireNesting,
        "rewrite work" | "aggregate nested scan bytes" => SectionTextLimitKind::WireWork,
        "fields" => SectionTextLimitKind::WireFields,
        "text fragments" | "table entries" | "object references" => SectionTextLimitKind::Entries,
        _ => SectionTextLimitKind::WireBytes,
    }
}

fn text_limit_error(observed: usize, maximum: usize) -> SectionTextError {
    SectionTextError::LimitExceeded {
        kind: SectionTextLimitKind::TextBytes,
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
