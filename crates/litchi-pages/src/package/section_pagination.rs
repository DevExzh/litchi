//! Exact-source transactions for Pages section pagination.

use std::fmt;
use std::num::NonZeroU64;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::{SourceCatalog, package::EntryEdit};
use litchi_iwa_common::{
    WireLimits, decode_varint_from_bytes, encode_varint_into, varint::encoded_len, wire::WireView,
};
use litchi_iwa_core::{Archive, RawMessage, SnappyStream};
use litchi_iwa_protos::pages_section_codec::{
    DecodeOptions as ProjectionOptions, PaginationSnapshot, decode_pagination,
};
use thiserror::Error;

use super::{
    MAX_SECTIONS, NativeSectionReference, Package, PackageError, SECTION_MESSAGE_TYPE,
    decode_body_storage, effective_text_limit, find_object, native_section_references,
    root_references_with_limits,
};
use crate::{SectionSelector, section};

const SECTION_START_FIELD: u32 = 20;
const PAGE_NUMBERING_FIELD: u32 = 21;
const STARTING_PAGE_NUMBER_FIELD: u32 = 22;

/// A finite resource governed while section pagination is rewritten or
/// published.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SectionPaginationLimitKind {
    /// Complete input package bytes.
    InputBytes,
    /// Complete rewritten package bytes.
    OutputBytes,
    /// ZIP members, IWA objects, or IWA messages.
    Entries,
    /// Bytes in one package member or IWA value.
    EntryBytes,
    /// Aggregate package or IWA bytes.
    TotalBytes,
    /// Aggregate retained Pages semantic metadata.
    SemanticBytes,
    /// Bytes in one protobuf payload.
    WireBytes,
    /// Parsed protobuf fields.
    WireFields,
    /// Protobuf nesting depth.
    WireNesting,
    /// Aggregate protobuf rewrite work.
    WireWork,
}

impl fmt::Display for SectionPaginationLimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::Entries => "entries",
            Self::EntryBytes => "entry bytes",
            Self::TotalBytes => "total bytes",
            Self::SemanticBytes => "semantic bytes",
            Self::WireBytes => "wire bytes",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting",
            Self::WireWork => "wire work",
        })
    }
}

/// An error raised while reading, staging, or publishing section pagination.
///
/// Errors intentionally omit native identifiers, member names, document
/// content, raw bytes, and lower-layer diagnostic strings.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum SectionPaginationError {
    /// More than one section matched an exact-name selector.
    #[error(
        "the Pages section-pagination selector is ambiguous at positions {first:?} and {duplicate:?}"
    )]
    AmbiguousSelector {
        /// First matching source position.
        first: Position,
        /// Next matching source position.
        duplicate: Position,
    },
    /// No section matched an exact-name selector.
    #[error("the Pages section-pagination selector did not match a section")]
    NameNotFound,
    /// No section exists at the requested source position.
    #[error("the Pages section position {position:?} does not exist")]
    PositionNotFound {
        /// Requested semantic source position.
        position: Position,
    },
    /// The requested pagination violates a public semantic invariant.
    #[error("invalid Pages section pagination: {0}")]
    InvalidPagination(#[source] section::pagination::Error),
    /// The snapshot has no exact physical source suitable for a changed edit.
    #[error("this Pages source does not support physical section-pagination edits")]
    UnsupportedSource,
    /// The selected native section cannot be decoded or rewritten safely.
    #[error("the Pages source cannot be edited safely")]
    InvalidSource,
    /// A finite transaction resource ceiling was exceeded.
    #[error(
        "Pages section-pagination {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    LimitExceeded {
        /// Resource category that exceeded its limit.
        kind: SectionPaginationLimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded destination allocation failed before publication.
    #[error("could not allocate {amount} units for the Pages section-pagination transaction")]
    Allocation {
        /// Elements or bytes requested.
        amount: usize,
    },
    /// Full semantic readback did not reproduce the requested change.
    #[error("the edited Pages section pagination failed semantic verification")]
    Verification,
    /// The patch does not belong to this exact immutable package artifact.
    #[error("the Pages section-pagination patch does not match the exact source package")]
    PatchConflict,
}

/// A mutable pagination value staged against one immutable Pages package.
pub struct SectionPaginationEdit<'a> {
    source: &'a Package,
    position: Position,
    before: section::Pagination,
    pagination: section::Pagination,
}

impl fmt::Debug for SectionPaginationEdit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SectionPaginationEdit")
            .field("position", &self.position)
            .field("pagination", &self.pagination)
            .finish_non_exhaustive()
    }
}

impl<'a> SectionPaginationEdit<'a> {
    fn new<'selector>(
        source: &'a Package,
        selector: impl Into<SectionSelector<'selector>>,
    ) -> Result<Self, SectionPaginationError> {
        let position = resolve_position(source, selector)?;
        let before = pagination_at(source, position)?;
        Ok(Self {
            source,
            position,
            before,
            pagination: before,
        })
    }

    /// Return the semantic source position resolved when this edit began.
    #[must_use]
    pub const fn position(&self) -> Position {
        self.position
    }

    /// Return the lossless pagination value that would be published.
    #[must_use]
    pub const fn pagination(&self) -> section::Pagination {
        self.pagination
    }

    /// Replace the staged pagination value after validating its invariants.
    ///
    /// # Errors
    ///
    /// Returns [`SectionPaginationError::InvalidPagination`] without changing
    /// the staged value when a known enum value uses a noncanonical alias.
    pub fn set_pagination(
        &mut self,
        pagination: section::Pagination,
    ) -> Result<&mut Self, SectionPaginationError> {
        pagination
            .validate()
            .map_err(SectionPaginationError::InvalidPagination)?;
        self.pagination = pagination;
        Ok(self)
    }

    /// Set or clear the page on which the section begins.
    ///
    /// # Errors
    ///
    /// Returns [`SectionPaginationError::InvalidPagination`] when a known
    /// value is represented by an `Unknown` alias.
    pub fn set_start(
        &mut self,
        start: Option<section::Start>,
    ) -> Result<&mut Self, SectionPaginationError> {
        let mut staged = self.pagination;
        staged
            .set_start(start)
            .map_err(SectionPaginationError::InvalidPagination)?;
        self.pagination = staged;
        Ok(self)
    }

    /// Set or clear whether page numbering continues or restarts.
    ///
    /// # Errors
    ///
    /// Returns [`SectionPaginationError::InvalidPagination`] when a known
    /// value is represented by an `Unknown` alias.
    pub fn set_page_numbering(
        &mut self,
        numbering: Option<section::PageNumbering>,
    ) -> Result<&mut Self, SectionPaginationError> {
        let mut staged = self.pagination;
        staged
            .set_page_numbering(numbering)
            .map_err(SectionPaginationError::InvalidPagination)?;
        self.pagination = staged;
        Ok(self)
    }

    /// Set or clear the first page number for this section.
    pub fn set_starting_page_number(&mut self, number: Option<section::PageNumber>) -> &mut Self {
        self.pagination.set_starting_page_number(number);
        self
    }

    /// Clear all three native pagination fields.
    pub fn clear(&mut self) {
        self.pagination = section::Pagination::new();
    }

    /// Validate and atomically publish the staged immutable candidate.
    ///
    /// Exact semantic no-ops reuse the original source allocation, including
    /// for normalized legacy packages. A changed edit requires exact ZIP
    /// provenance and is published only after complete candidate reopening and
    /// semantic verification under the retained limits.
    ///
    /// # Errors
    ///
    /// Returns an error without modifying `source` when selection, source
    /// topology, resource bounds, physical preservation, or readback fails.
    pub fn commit(self) -> Result<SectionPaginationCommit, SectionPaginationError> {
        self.pagination
            .validate()
            .map_err(SectionPaginationError::InvalidPagination)?;
        self.source.validate().map_err(map_package_error)?;
        if pagination_at(self.source, self.position)? != self.before {
            return Err(SectionPaginationError::InvalidSource);
        }

        let source_bytes = self.source.state.source.shared_source();
        let source_fingerprint = fingerprint(&source_bytes);
        if self.before == self.pagination {
            return Ok(SectionPaginationCommit {
                package: self.source.snapshot(),
                patch: SectionPaginationPatch {
                    source_bytes: Arc::clone(&source_bytes),
                    target_bytes: source_bytes,
                    source_fingerprint,
                    target_fingerprint: source_fingerprint,
                    position: self.position,
                    before: self.before,
                    after: self.pagination,
                },
                diagnostics: SectionPaginationDiagnostics::unchanged(),
            });
        }

        if !self.source.state.source.source_is_exact() {
            return Err(SectionPaginationError::UnsupportedSource);
        }
        let package =
            rewrite_package_pagination(self.source, self.position, self.before, self.pagination)?;
        let target_bytes = package.state.source.shared_source();
        let target_fingerprint = fingerprint(&target_bytes);
        Ok(SectionPaginationCommit {
            package,
            patch: SectionPaginationPatch {
                source_bytes,
                target_bytes,
                source_fingerprint,
                target_fingerprint,
                position: self.position,
                before: self.before,
                after: self.pagination,
            },
            diagnostics: SectionPaginationDiagnostics::published(),
        })
    }
}

/// An exact-source-checked, reversible Pages section-pagination patch.
///
/// Native identifiers, member names, and exact source/target bytes remain
/// private. Fingerprints are diagnostics only; exact byte comparison
/// authorizes application.
#[derive(Clone, PartialEq, Eq)]
pub struct SectionPaginationPatch {
    source_bytes: Arc<[u8]>,
    target_bytes: Arc<[u8]>,
    source_fingerprint: u64,
    target_fingerprint: u64,
    position: Position,
    before: section::Pagination,
    after: section::Pagination,
}

impl fmt::Debug for SectionPaginationPatch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SectionPaginationPatch")
            .field("position", &self.position)
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl SectionPaginationPatch {
    /// Return the semantic source position selected by this patch.
    #[must_use]
    pub const fn position(&self) -> Position {
        self.position
    }

    /// Return the pagination required from the patch source.
    #[must_use]
    pub const fn before(&self) -> section::Pagination {
        self.before
    }

    /// Return the pagination produced by the patch target.
    #[must_use]
    pub const fn after(&self) -> section::Pagination {
        self.after
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

    /// Return whether the patch preserves both semantic pagination and bytes.
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
            before: self.after,
            after: self.before,
        }
    }
}

/// Compact evidence describing one section-pagination commit.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct SectionPaginationDiagnostics {
    changed: bool,
    touched_components: usize,
    full_reparse_performed: bool,
}

impl SectionPaginationDiagnostics {
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

/// The fully verified result of one immutable section-pagination transaction.
#[must_use = "a Pages section-pagination commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct SectionPaginationCommit {
    package: Package,
    patch: SectionPaginationPatch,
    diagnostics: SectionPaginationDiagnostics,
}

impl SectionPaginationCommit {
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
    pub const fn patch(&self) -> &SectionPaginationPatch {
        &self.patch
    }

    /// Borrow compact publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &SectionPaginationDiagnostics {
        &self.diagnostics
    }
}

impl Package {
    /// Read lossless pagination for one semantically selected section.
    ///
    /// # Errors
    ///
    /// Returns a typed error when selection is ambiguous or missing, or when
    /// the selected native payload is malformed.
    pub fn section_pagination<'selector>(
        &self,
        selector: impl Into<SectionSelector<'selector>>,
    ) -> Result<section::Pagination, SectionPaginationError> {
        let position = resolve_position(self, selector)?;
        pagination_at(self, position)
    }

    /// Start a selector-first edit of one section's pagination.
    ///
    /// The selector is resolved immediately against this immutable semantic
    /// snapshot and only its typed source position is retained.
    ///
    /// # Errors
    ///
    /// Returns a typed error when the selector is ambiguous or missing, or
    /// when the existing pagination cannot be decoded safely.
    pub fn edit_section_pagination<'selector>(
        &self,
        selector: impl Into<SectionSelector<'selector>>,
    ) -> Result<SectionPaginationEdit<'_>, SectionPaginationError> {
        SectionPaginationEdit::new(self, selector)
    }

    /// Apply an exact-source-checked section-pagination patch.
    ///
    /// The retained target is fully reopened and semantically verified under
    /// this package's original limits before it is published.
    ///
    /// # Errors
    ///
    /// Returns [`SectionPaginationError::PatchConflict`] unless this package
    /// is the exact immutable source captured by `patch`, or another typed
    /// error when the retained target cannot be safely published.
    pub fn apply_section_pagination(
        &self,
        patch: &SectionPaginationPatch,
    ) -> Result<SectionPaginationCommit, SectionPaginationError> {
        let source = &self.state.source;
        let source_bytes = source.shared_source();
        if fingerprint(source.source_bytes()) != patch.source_fingerprint
            || source.source_bytes() != patch.source_bytes.as_ref()
            || source_bytes.as_ref() != patch.source_bytes.as_ref()
        {
            return Err(SectionPaginationError::PatchConflict);
        }
        self.validate().map_err(map_package_error)?;
        if pagination_at(self, patch.position)? != patch.before {
            return Err(SectionPaginationError::PatchConflict);
        }

        if patch.is_noop() {
            return Ok(SectionPaginationCommit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: SectionPaginationDiagnostics::unchanged(),
            });
        }
        if !source.source_is_exact() || fingerprint(&patch.target_bytes) != patch.target_fingerprint
        {
            return Err(SectionPaginationError::PatchConflict);
        }

        let candidate_source = SourceCatalog::from_shared_bytes_with_limits(
            Arc::clone(&patch.target_bytes),
            source.limits(),
        )
        .map_err(map_archive_error)?;
        let candidate =
            Package::from_source_catalog(candidate_source).map_err(map_package_error)?;
        candidate.validate().map_err(map_package_error)?;
        verify_candidate(self, &candidate, patch.position, patch.after)?;
        Ok(SectionPaginationCommit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: SectionPaginationDiagnostics::published(),
        })
    }
}

fn resolve_position<'selector>(
    package: &Package,
    selector: impl Into<SectionSelector<'selector>>,
) -> Result<Position, SectionPaginationError> {
    let selected_selector = selector.into();
    let selected = package
        .semantic_document()
        .select_section(selected_selector)
        .map_err(map_selector_error)?
        .ok_or(match selected_selector {
            SectionSelector::Name(_) => SectionPaginationError::NameNotFound,
            SectionSelector::Position(position) => {
                SectionPaginationError::PositionNotFound { position }
            },
        })?;
    Ok(Position::new(selected.index()))
}

fn rewrite_package_pagination(
    source: &Package,
    position: Position,
    before: section::Pagination,
    after: section::Pagination,
) -> Result<Package, SectionPaginationError> {
    let source_catalog = &source.state.source;
    let identifier = section_identifier_at(source, position)?;
    let mut matching_components = source_catalog
        .components()
        .iter()
        .filter(|component| component.archive().object(identifier.get()).is_some());
    let component = matching_components
        .next()
        .ok_or(SectionPaginationError::InvalidSource)?;
    if matching_components.next().is_some() {
        return Err(SectionPaginationError::InvalidSource);
    }
    let component_name = component.name();
    let entry = source_catalog
        .package()
        .iter()
        .find(|entry| entry.name() == component_name)
        .ok_or(SectionPaginationError::InvalidSource)?;
    if entry.is_opaque() {
        return Err(SectionPaginationError::InvalidSource);
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
        .object(identifier.get())
        .ok_or(SectionPaginationError::InvalidSource)?;
    let mut messages = object
        .messages
        .iter()
        .enumerate()
        .filter(|(_index, message)| message.type_ == SECTION_MESSAGE_TYPE);
    let (message_index, message) = messages
        .next()
        .ok_or(SectionPaginationError::InvalidSource)?;
    if messages.next().is_some() {
        return Err(SectionPaginationError::InvalidSource);
    }

    let wire_limits = wire_limits(source)?;
    if strict_payload_pagination(&message.data, wire_limits)? != before {
        return Err(SectionPaginationError::InvalidSource);
    }
    let rewritten = rewrite_pagination_payload(&message.data, before, after, wire_limits)?;
    if strict_payload_pagination(&rewritten, wire_limits)? != after {
        return Err(SectionPaginationError::Verification);
    }

    archive
        .object_mut(identifier.get())
        .ok_or(SectionPaginationError::InvalidSource)?
        .replace_message_preserving_header_with_limits(
            message_index,
            RawMessage {
                type_: SECTION_MESSAGE_TYPE,
                data: rewritten,
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
    verify_candidate(source, &candidate, position, after)?;
    Ok(candidate)
}

fn section_identifier_at(
    package: &Package,
    position: Position,
) -> Result<NonZeroU64, SectionPaginationError> {
    let components = package.state.source.components();
    let limits = package.state.source.limits();
    let root = root_references_with_limits(components, limits).map_err(map_package_error)?;
    let body_identifier = root.body.ok_or(SectionPaginationError::UnsupportedSource)?;
    let body = find_object(components, body_identifier.get())
        .ok_or(SectionPaginationError::InvalidSource)?;
    let max_text_bytes = effective_text_limit(limits);
    let (_storage, table_references) = decode_body_storage(
        &body.messages,
        body_identifier,
        MAX_SECTIONS,
        max_text_bytes,
        limits,
    )
    .map_err(map_package_error)?;
    let references =
        native_section_references(table_references, root.initial_section, MAX_SECTIONS)
            .map_err(map_package_error)?;
    references
        .get(position.get())
        .copied()
        .map(|NativeSectionReference { identifier, .. }| identifier)
        .ok_or(SectionPaginationError::InvalidSource)
}

fn pagination_at(
    package: &Package,
    position: Position,
) -> Result<section::Pagination, SectionPaginationError> {
    if package.sections().get(position.get()).is_none() {
        return Err(SectionPaginationError::PositionNotFound { position });
    }
    let identifier = section_identifier_at(package, position)?;
    let mut payload = None;
    for component in package.state.source.components().iter() {
        let Some(object) = component.archive().object(identifier.get()) else {
            continue;
        };
        if payload.is_some() {
            return Err(SectionPaginationError::InvalidSource);
        }
        let mut messages = object
            .messages
            .iter()
            .filter(|message| message.type_ == SECTION_MESSAGE_TYPE);
        let message = messages
            .next()
            .ok_or(SectionPaginationError::InvalidSource)?;
        if messages.next().is_some() {
            return Err(SectionPaginationError::InvalidSource);
        }
        payload = Some(message.data.as_slice());
    }
    strict_payload_pagination(
        payload.ok_or(SectionPaginationError::InvalidSource)?,
        wire_limits(package)?,
    )
}

fn strict_payload_pagination(
    payload: &[u8],
    limits: WireLimits,
) -> Result<section::Pagination, SectionPaginationError> {
    let view = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    let mut start = None;
    let mut numbering = None;
    let mut page = None;
    for field in view.fields() {
        let destination = match field.number() {
            SECTION_START_FIELD => &mut start,
            PAGE_NUMBERING_FIELD => &mut numbering,
            STARTING_PAGE_NUMBER_FIELD => &mut page,
            _ => continue,
        };
        if destination.is_some() || field.wire_type() != 0 {
            return Err(SectionPaginationError::InvalidSource);
        }
        field.validate_canonical_key().map_err(map_wire_error)?;
        let (value, consumed) = decode_varint_from_bytes(field.payload())
            .map_err(|_error| SectionPaginationError::InvalidSource)?;
        if consumed != field.payload().len() || encoded_len(value) != consumed {
            return Err(SectionPaginationError::InvalidSource);
        }
        *destination =
            Some(u32::try_from(value).map_err(|_error| SectionPaginationError::InvalidSource)?);
    }

    let projected = decode_pagination(payload, ProjectionOptions::new(payload.len(), 1))
        .map_err(|_error| SectionPaginationError::InvalidSource)?;
    let preflight = PaginationSnapshot {
        section_start_kind: start,
        section_page_number_kind: numbering,
        section_page_number_start: page,
    };
    if projected != preflight {
        return Err(SectionPaginationError::InvalidSource);
    }

    let mut pagination = section::Pagination::new();
    pagination
        .set_start(start.map(section::Start::from_raw))
        .map_err(SectionPaginationError::InvalidPagination)?;
    pagination
        .set_page_numbering(numbering.map(section::PageNumbering::from_raw))
        .map_err(SectionPaginationError::InvalidPagination)?;
    let page_number = page
        .map(section::PageNumber::new)
        .transpose()
        .map_err(|_error| SectionPaginationError::InvalidSource)?;
    pagination.set_starting_page_number(page_number);
    Ok(pagination)
}

fn rewrite_pagination_payload(
    source: &[u8],
    before: section::Pagination,
    after: section::Pagination,
    limits: WireLimits,
) -> Result<Vec<u8>, SectionPaginationError> {
    if strict_payload_pagination(source, limits)? != before {
        return Err(SectionPaginationError::InvalidSource);
    }
    let view = WireView::parse_with_limits(source, limits).map_err(map_wire_error)?;
    let fields = [
        SECTION_START_FIELD,
        PAGE_NUMBERING_FIELD,
        STARTING_PAGE_NUMBER_FIELD,
    ];
    let mut output_length = source.len();
    let mut output_fields = view.len();
    for field_number in fields {
        let previous = pagination_value(before, field_number);
        let replacement = pagination_value(after, field_number);
        if previous == replacement {
            continue;
        }
        if let Some(field) = view.fields().find(|field| field.number() == field_number) {
            output_length = output_length
                .checked_sub(field.raw().len())
                .ok_or(SectionPaginationError::InvalidSource)?;
            output_fields = output_fields
                .checked_sub(1)
                .ok_or(SectionPaginationError::InvalidSource)?;
        }
        if let Some(value) = replacement {
            output_length = output_length
                .checked_add(encoded_varint_field_length(field_number, value))
                .ok_or_else(|| output_limit_error(usize::MAX, limits))?;
            output_fields = output_fields
                .checked_add(1)
                .ok_or(SectionPaginationError::InvalidSource)?;
        }
    }
    if output_length > limits.max_output_bytes() {
        return Err(output_limit_error(output_length, limits));
    }
    if output_fields > limits.max_fields() {
        return Err(SectionPaginationError::LimitExceeded {
            kind: SectionPaginationLimitKind::WireFields,
            observed: usize_to_u64(output_fields),
            maximum: usize_to_u64(limits.max_fields()),
        });
    }
    let work = view
        .len()
        .checked_add(output_fields)
        .ok_or_else(|| work_limit_error(usize::MAX, limits))?;
    if work > limits.max_rewrite_work() {
        return Err(work_limit_error(work, limits));
    }

    let mut output = allocate_bytes(output_length)?;
    let mut emitted = [false; 3];
    for field in view.fields() {
        let Some(index) = fields
            .iter()
            .position(|field_number| *field_number == field.number())
        else {
            output.extend_from_slice(field.raw());
            continue;
        };
        emitted[index] = true;
        let previous = pagination_value(before, fields[index]);
        let replacement = pagination_value(after, fields[index]);
        if previous == replacement {
            output.extend_from_slice(field.raw());
        } else if let Some(value) = replacement {
            append_varint_field(&mut output, fields[index], value);
        }
    }
    for (index, field_number) in fields.into_iter().enumerate() {
        if !emitted[index]
            && let Some(value) = pagination_value(after, field_number)
        {
            append_varint_field(&mut output, field_number, value);
        }
    }
    if output.len() != output_length {
        return Err(SectionPaginationError::Verification);
    }
    Ok(output)
}

fn pagination_value(pagination: section::Pagination, field_number: u32) -> Option<u32> {
    match field_number {
        SECTION_START_FIELD => pagination.start().map(section::Start::as_raw),
        PAGE_NUMBERING_FIELD => pagination
            .page_numbering()
            .map(section::PageNumbering::as_raw),
        STARTING_PAGE_NUMBER_FIELD => pagination
            .starting_page_number()
            .map(section::PageNumber::get),
        _ => None,
    }
}

fn append_varint_field(output: &mut Vec<u8>, field_number: u32, value: u32) {
    encode_varint_into(output, u64::from(field_number) << 3);
    encode_varint_into(output, u64::from(value));
}

fn encoded_varint_field_length(field_number: u32, value: u32) -> usize {
    encoded_len(u64::from(field_number) << 3).saturating_add(encoded_len(u64::from(value)))
}

fn wire_limits(package: &Package) -> Result<WireLimits, SectionPaginationError> {
    let maximum = package
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(map_archive_error)?
        .max_message_bytes();
    WireLimits::default()
        .with_input_bytes(maximum)
        .and_then(|limits| limits.with_output_bytes(maximum))
        .map_err(map_wire_error)
}

fn verify_candidate(
    source: &Package,
    candidate: &Package,
    position: Position,
    expected: section::Pagination,
) -> Result<(), SectionPaginationError> {
    if source.stats().total_objects() != candidate.stats().total_objects()
        || source.sections().len() != candidate.sections().len()
    {
        return Err(SectionPaginationError::Verification);
    }
    for (before, after) in source.sections().iter().zip(candidate.sections()) {
        if before.name() != after.name()
            || before.section_type() != after.section_type()
            || before.heading() != after.heading()
            || before.paragraphs() != after.paragraphs()
            || before.text_storages() != after.text_storages()
            || before.page_count() != after.page_count()
        {
            return Err(SectionPaginationError::Verification);
        }
    }
    if pagination_at(candidate, position)? != expected {
        return Err(SectionPaginationError::Verification);
    }
    Ok(())
}

fn allocate_bytes(capacity: usize) -> Result<Vec<u8>, SectionPaginationError> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(capacity)
        .map_err(|_allocation| SectionPaginationError::Allocation { amount: capacity })?;
    Ok(output)
}

#[allow(
    clippy::needless_pass_by_value,
    reason = "used directly as a Result::map_err conversion"
)]
fn map_selector_error(selection_error: crate::SelectorError) -> SectionPaginationError {
    match selection_error {
        crate::SelectorError::AmbiguousSectionName {
            first, duplicate, ..
        } => SectionPaginationError::AmbiguousSelector {
            first: Position::new(first),
            duplicate: Position::new(duplicate),
        },
    }
}

#[allow(
    clippy::needless_pass_by_value,
    reason = "used directly as a Result::map_err conversion"
)]
fn map_package_error(package_error: PackageError) -> SectionPaginationError {
    match package_error {
        PackageError::Archive(archive_error) => map_archive_error(archive_error),
        PackageError::SectionNamesTooLarge { observed, limit } => {
            SectionPaginationError::LimitExceeded {
                kind: SectionPaginationLimitKind::SemanticBytes,
                observed: usize_to_u64(observed),
                maximum: usize_to_u64(limit),
            }
        },
        PackageError::NotPages => SectionPaginationError::UnsupportedSource,
        PackageError::PayloadLimit { observed, limit } => SectionPaginationError::LimitExceeded {
            kind: SectionPaginationLimitKind::WireBytes,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
        PackageError::ObjectLimit { observed, limit } => SectionPaginationError::LimitExceeded {
            kind: SectionPaginationLimitKind::Entries,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
        PackageError::Allocation { amount } => SectionPaginationError::Allocation { amount },
        PackageError::Io(_)
        | PackageError::Detection(_)
        | PackageError::InvalidFormat(_)
        | PackageError::Semantic(_) => SectionPaginationError::InvalidSource,
    }
}

#[allow(
    clippy::needless_pass_by_value,
    reason = "used directly as a Result::map_err conversion"
)]
fn map_archive_error(archive_error: litchi_iwa_archive::Error) -> SectionPaginationError {
    match archive_error {
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => SectionPaginationError::LimitExceeded {
            kind: match kind {
                litchi_iwa_archive::LimitKind::InputBytes => SectionPaginationLimitKind::InputBytes,
                litchi_iwa_archive::LimitKind::OutputBytes => {
                    SectionPaginationLimitKind::OutputBytes
                },
                litchi_iwa_archive::LimitKind::Entries => SectionPaginationLimitKind::Entries,
                litchi_iwa_archive::LimitKind::MemberNameBytes
                | litchi_iwa_archive::LimitKind::MetadataBytes
                | litchi_iwa_archive::LimitKind::CompressedEntryBytes
                | litchi_iwa_archive::LimitKind::EntryBytes
                | litchi_iwa_archive::LimitKind::IwaStreamBytes => {
                    SectionPaginationLimitKind::EntryBytes
                },
                litchi_iwa_archive::LimitKind::TotalBytes
                | litchi_iwa_archive::LimitKind::IwaTotalBytes => {
                    SectionPaginationLimitKind::TotalBytes
                },
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => {
            SectionPaginationError::Allocation { amount }
        },
        litchi_iwa_archive::Error::Iwa(core_error) => map_core_error(core_error),
        litchi_iwa_archive::Error::Io(_)
        | litchi_iwa_archive::Error::Zip { .. }
        | litchi_iwa_archive::Error::InvalidLimits(_)
        | litchi_iwa_archive::Error::Encrypted
        | litchi_iwa_archive::Error::SourceChanged { .. }
        | litchi_iwa_archive::Error::DirectoryChanged { .. }
        | litchi_iwa_archive::Error::Reassembly(_)
        | litchi_iwa_archive::Error::InvalidBundle(_) => SectionPaginationError::InvalidSource,
    }
}

#[allow(
    clippy::needless_pass_by_value,
    reason = "used directly as a Result::map_err conversion"
)]
fn map_core_error(core_error: litchi_iwa_core::Error) -> SectionPaginationError {
    match core_error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => SectionPaginationError::LimitExceeded {
            kind: match kind {
                litchi_iwa_core::LimitKind::Objects
                | litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject
                | litchi_iwa_core::LimitKind::MetadataItems => SectionPaginationLimitKind::Entries,
                litchi_iwa_core::LimitKind::MessageBytes => SectionPaginationLimitKind::WireBytes,
                litchi_iwa_core::LimitKind::HeaderFields => SectionPaginationLimitKind::WireFields,
                litchi_iwa_core::LimitKind::HeaderNesting => {
                    SectionPaginationLimitKind::WireNesting
                },
                litchi_iwa_core::LimitKind::ArchiveBytes
                | litchi_iwa_core::LimitKind::ObjectBytes
                | litchi_iwa_core::LimitKind::HeaderBytes
                | litchi_iwa_core::LimitKind::HeaderMemoryBytes
                | litchi_iwa_core::LimitKind::SnappyChunkBytes
                | litchi_iwa_core::LimitKind::SnappyStreamBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedChunkBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedStreamBytes
                | litchi_iwa_core::LimitKind::SnappyFrames => {
                    SectionPaginationLimitKind::EntryBytes
                },
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            SectionPaginationError::Allocation { amount: requested }
        },
        litchi_iwa_core::Error::InvalidArchive { .. }
        | litchi_iwa_core::Error::InvalidLimits { .. }
        | litchi_iwa_core::Error::HeaderCodec { .. }
        | litchi_iwa_core::Error::Io(_)
        | litchi_iwa_core::Error::Snappy { .. } => SectionPaginationError::InvalidSource,
    }
}

#[allow(
    clippy::needless_pass_by_value,
    reason = "used directly as a Result::map_err conversion"
)]
fn map_wire_error(wire_error: litchi_iwa_common::Error) -> SectionPaginationError {
    match wire_error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => SectionPaginationError::LimitExceeded {
            kind: match kind {
                litchi_iwa_common::LimitKind::InputBytes
                | litchi_iwa_common::LimitKind::OutputBytes => {
                    SectionPaginationLimitKind::WireBytes
                },
                litchi_iwa_common::LimitKind::Fields
                | litchi_iwa_common::LimitKind::TableRows
                | litchi_iwa_common::LimitKind::TableColumns
                | litchi_iwa_common::LimitKind::TableCells
                | litchi_iwa_common::LimitKind::MaterializedCells => {
                    SectionPaginationLimitKind::WireFields
                },
                litchi_iwa_common::LimitKind::Nesting => SectionPaginationLimitKind::WireNesting,
                litchi_iwa_common::LimitKind::RewriteWork => SectionPaginationLimitKind::WireWork,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => {
            SectionPaginationError::Allocation { amount }
        },
        litchi_iwa_common::Error::InvalidFormat(_)
        | litchi_iwa_common::Error::InvalidLimit { .. } => SectionPaginationError::InvalidSource,
    }
}

fn output_limit_error(observed: usize, limits: WireLimits) -> SectionPaginationError {
    SectionPaginationError::LimitExceeded {
        kind: SectionPaginationLimitKind::WireBytes,
        observed: usize_to_u64(observed),
        maximum: usize_to_u64(limits.max_output_bytes()),
    }
}

fn work_limit_error(observed: usize, limits: WireLimits) -> SectionPaginationError {
    SectionPaginationError::LimitExceeded {
        kind: SectionPaginationLimitKind::WireWork,
        observed: usize_to_u64(observed),
        maximum: usize_to_u64(limits.max_rewrite_work()),
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
