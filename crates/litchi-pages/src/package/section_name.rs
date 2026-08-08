//! Exact-source transactions for Pages section names.

use std::fmt;
use std::mem::size_of;
use std::num::NonZeroU64;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::{SourceCatalog, package::EntryEdit};
use litchi_iwa_common::{WireLimits, encode_varint_into, varint::encoded_len, wire::WireView};
use litchi_iwa_core::{Archive, RawMessage, SnappyStream};
use thiserror::Error;

use super::{
    MAX_SECTIONS, NativeSectionReference, Package, PackageError, SECTION_MESSAGE_TYPE,
    decode_body_storage, effective_text_limit, find_object, native_section_references,
    root_references, validate_section_table_wire,
};
use crate::{SectionSelector, section};

const SECTION_NAME_FIELD: u32 = 26;

/// A finite resource governed while a section name is rewritten or published.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SectionNameLimitKind {
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
    /// Aggregate retained section-name memory.
    NameBytes,
    /// Bytes in one protobuf payload.
    WireBytes,
    /// Parsed protobuf fields.
    WireFields,
    /// Protobuf nesting depth.
    WireNesting,
    /// Aggregate protobuf rewrite work.
    WireWork,
}

impl fmt::Display for SectionNameLimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::Entries => "entries",
            Self::EntryBytes => "entry bytes",
            Self::TotalBytes => "total bytes",
            Self::NameBytes => "section-name bytes",
            Self::WireBytes => "wire bytes",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting",
            Self::WireWork => "wire work",
        })
    }
}

/// An error raised while selecting, staging, or publishing a section name.
///
/// Errors intentionally omit authored names, native identifiers, member
/// names, raw bytes, and lower-layer diagnostic strings.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum SectionNameError {
    /// More than one section matched an exact-name selector.
    #[error(
        "the Pages section-name selector is ambiguous at positions {first:?} and {duplicate:?}"
    )]
    AmbiguousSelector {
        /// First matching source position.
        first: Position,
        /// Next matching source position.
        duplicate: Position,
    },
    /// No section matched an exact-name selector.
    #[error("the Pages section-name selector did not match a section")]
    NameNotFound,
    /// No section exists at the requested source position.
    #[error("the Pages section position {position:?} does not exist")]
    PositionNotFound {
        /// Requested semantic source position.
        position: Position,
    },
    /// The requested name violates a public semantic invariant.
    #[error("invalid Pages section name: {0}")]
    InvalidName(#[source] section::Error),
    /// The snapshot has no exact physical source suitable for a changed edit.
    #[error("this Pages source does not support physical section-name edits")]
    UnsupportedSource,
    /// The selected native section cannot be rewritten without ambiguity.
    #[error("the Pages source cannot be edited safely")]
    InvalidSource,
    /// A finite transaction resource ceiling was exceeded.
    #[error("Pages section-name {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    LimitExceeded {
        /// Resource category that exceeded its limit.
        kind: SectionNameLimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded destination allocation failed before publication.
    #[error("could not allocate {amount} units for the Pages section-name transaction")]
    Allocation {
        /// Elements or bytes requested.
        amount: usize,
    },
    /// Full semantic readback did not reproduce the requested change.
    #[error("the edited Pages section name failed semantic verification")]
    Verification,
    /// The patch does not belong to this exact immutable package artifact.
    #[error("the Pages section-name patch does not match the exact source package")]
    PatchConflict,
}

/// A mutable section-name value staged against one immutable Pages package.
pub struct SectionNameEdit<'a> {
    source: &'a Package,
    position: Position,
    before: Option<String>,
    name: Option<String>,
}

impl fmt::Debug for SectionNameEdit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SectionNameEdit")
            .field("position", &self.position)
            .finish_non_exhaustive()
    }
}

impl<'a> SectionNameEdit<'a> {
    fn new<'selector>(
        source: &'a Package,
        selector: impl Into<SectionSelector<'selector>>,
    ) -> Result<Self, SectionNameError> {
        let selected_selector = selector.into();
        let selected = source
            .semantic_document()
            .select_section(selected_selector)
            .map_err(map_selector_error)?
            .ok_or(match selected_selector {
                SectionSelector::Name(_) => SectionNameError::NameNotFound,
                SectionSelector::Position(position) => {
                    SectionNameError::PositionNotFound { position }
                },
            })?;
        let position = Position::new(selected.index());
        let before = selected.name().map(try_owned_name).transpose()?;
        let name = before.clone();
        Ok(Self {
            source,
            position,
            before,
            name,
        })
    }

    /// Return the optional name that would be published by this edit.
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// Set or clear the producer-visible section name.
    ///
    /// `None` removes field presence while `Some("")` retains an explicitly
    /// present empty native string.
    ///
    /// # Errors
    ///
    /// Returns [`SectionNameError::InvalidName`] when `name` contains NUL, or
    /// a bounded allocation/limit error before the staged value changes.
    pub fn set_name(&mut self, name: Option<&str>) -> Result<&mut Self, SectionNameError> {
        if name.is_some_and(|candidate| candidate.contains('\0')) {
            return Err(SectionNameError::InvalidName(
                section::Error::NameContainsNul,
            ));
        }
        let staged_name = name.map(try_owned_name).transpose()?;
        validate_name_memory(self.source, self.position, staged_name.as_deref())?;
        self.name = staged_name;
        Ok(self)
    }

    /// Remove the producer-visible section name.
    pub fn clear_name(&mut self) {
        self.name = None;
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
    pub fn commit(self) -> Result<SectionNameCommit, SectionNameError> {
        validate_name_memory(self.source, self.position, self.name.as_deref())?;
        self.source.validate().map_err(map_package_error)?;
        if section_name_at(self.source, self.position)? != self.before.as_deref() {
            return Err(SectionNameError::InvalidSource);
        }

        let source_bytes = self.source.state.source.shared_source();
        let source_fingerprint = fingerprint(&source_bytes);
        if self.before == self.name {
            return Ok(SectionNameCommit {
                package: self.source.snapshot(),
                patch: SectionNamePatch {
                    source_bytes: Arc::clone(&source_bytes),
                    target_bytes: source_bytes,
                    source_fingerprint,
                    target_fingerprint: source_fingerprint,
                    position: self.position,
                    before: self.before,
                    after: self.name,
                },
                diagnostics: SectionNameDiagnostics::unchanged(),
            });
        }

        if !self.source.state.source.source_is_exact() {
            return Err(SectionNameError::UnsupportedSource);
        }
        let package = rewrite_package_name(
            self.source,
            self.position,
            self.before.as_deref(),
            self.name.as_deref(),
        )?;
        let target_bytes = package.state.source.shared_source();
        let target_fingerprint = fingerprint(&target_bytes);
        Ok(SectionNameCommit {
            package,
            patch: SectionNamePatch {
                source_bytes,
                target_bytes,
                source_fingerprint,
                target_fingerprint,
                position: self.position,
                before: self.before,
                after: self.name,
            },
            diagnostics: SectionNameDiagnostics::published(),
        })
    }
}

/// An exact-source-checked, reversible Pages section-name patch.
///
/// Native identifiers, member names, and exact source/target bytes remain
/// private. Fingerprints are diagnostics only; exact byte comparison
/// authorizes application.
#[derive(Clone, PartialEq, Eq)]
pub struct SectionNamePatch {
    source_bytes: Arc<[u8]>,
    target_bytes: Arc<[u8]>,
    source_fingerprint: u64,
    target_fingerprint: u64,
    position: Position,
    before: Option<String>,
    after: Option<String>,
}

impl fmt::Debug for SectionNamePatch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SectionNamePatch")
            .field("position", &self.position)
            .finish_non_exhaustive()
    }
}

impl SectionNamePatch {
    /// Return the semantic source position selected by this patch.
    #[must_use]
    pub const fn position(&self) -> Position {
        self.position
    }

    /// Borrow the optional name required from the patch source.
    #[must_use]
    pub fn before(&self) -> Option<&str> {
        self.before.as_deref()
    }

    /// Borrow the optional name produced by the patch target.
    #[must_use]
    pub fn after(&self) -> Option<&str> {
        self.after.as_deref()
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

    /// Return whether the patch preserves both the semantic name and bytes.
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
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }
}

/// Compact evidence describing one section-name commit.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct SectionNameDiagnostics {
    changed: bool,
    touched_components: usize,
    full_reparse_performed: bool,
}

impl SectionNameDiagnostics {
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

/// The fully verified result of one immutable section-name transaction.
#[must_use = "a Pages section-name commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct SectionNameCommit {
    package: Package,
    patch: SectionNamePatch,
    diagnostics: SectionNameDiagnostics,
}

impl SectionNameCommit {
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
    pub const fn patch(&self) -> &SectionNamePatch {
        &self.patch
    }

    /// Borrow compact publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &SectionNameDiagnostics {
        &self.diagnostics
    }
}

impl Package {
    /// Start a selector-first edit of one producer-visible section name.
    ///
    /// The selector is resolved immediately against this immutable semantic
    /// snapshot and only its typed source position is retained.
    ///
    /// # Errors
    ///
    /// Returns a typed error when the selector is ambiguous or does not match
    /// a section, or when the existing name cannot be retained safely.
    pub fn edit_section_name<'selector>(
        &self,
        selector: impl Into<SectionSelector<'selector>>,
    ) -> Result<SectionNameEdit<'_>, SectionNameError> {
        SectionNameEdit::new(self, selector)
    }

    /// Apply an exact-source-checked section-name patch.
    ///
    /// The retained target is fully reopened and semantically verified under
    /// this package's original limits before it is published.
    ///
    /// # Errors
    ///
    /// Returns [`SectionNameError::PatchConflict`] unless this package is the
    /// exact immutable source captured by `patch`, or another typed error when
    /// the retained target cannot be safely published.
    pub fn apply_section_name(
        &self,
        patch: &SectionNamePatch,
    ) -> Result<SectionNameCommit, SectionNameError> {
        let source = &self.state.source;
        let source_bytes = source.shared_source();
        if fingerprint(source.source_bytes()) != patch.source_fingerprint
            || source.source_bytes() != patch.source_bytes.as_ref()
            || source_bytes.as_ref() != patch.source_bytes.as_ref()
        {
            return Err(SectionNameError::PatchConflict);
        }
        self.validate().map_err(map_package_error)?;
        if section_name_at(self, patch.position)? != patch.before() {
            return Err(SectionNameError::PatchConflict);
        }

        if patch.is_noop() {
            return Ok(SectionNameCommit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: SectionNameDiagnostics::unchanged(),
            });
        }
        if !source.source_is_exact() || fingerprint(&patch.target_bytes) != patch.target_fingerprint
        {
            return Err(SectionNameError::PatchConflict);
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
        Ok(SectionNameCommit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: SectionNameDiagnostics::published(),
        })
    }
}

fn rewrite_package_name(
    source: &Package,
    position: Position,
    before: Option<&str>,
    after: Option<&str>,
) -> Result<Package, SectionNameError> {
    let source_catalog = &source.state.source;
    let identifier = section_identifier_at(source, position)?;
    let mut matching_components = source_catalog
        .components()
        .iter()
        .filter(|component| component.archive().object(identifier.get()).is_some());
    let component = matching_components
        .next()
        .ok_or(SectionNameError::InvalidSource)?;
    if matching_components.next().is_some() {
        return Err(SectionNameError::InvalidSource);
    }
    let component_name = component.name();
    let entry = source_catalog
        .package()
        .iter()
        .find(|entry| entry.name() == component_name)
        .ok_or(SectionNameError::InvalidSource)?;
    if entry.is_opaque() {
        return Err(SectionNameError::InvalidSource);
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
        .ok_or(SectionNameError::InvalidSource)?;
    let mut messages = object
        .messages
        .iter()
        .enumerate()
        .filter(|(_index, message)| message.type_ == SECTION_MESSAGE_TYPE);
    let (message_index, message) = messages.next().ok_or(SectionNameError::InvalidSource)?;
    if messages.next().is_some() {
        return Err(SectionNameError::InvalidSource);
    }

    let wire_limits = wire_limits(source)?;
    if strict_payload_name(&message.data, wire_limits)? != before {
        return Err(SectionNameError::InvalidSource);
    }
    let rewritten = rewrite_name_payload(&message.data, before, after, wire_limits)?;
    if strict_payload_name(&rewritten, wire_limits)? != after {
        return Err(SectionNameError::Verification);
    }

    archive
        .object_mut(identifier.get())
        .ok_or(SectionNameError::InvalidSource)?
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
) -> Result<NonZeroU64, SectionNameError> {
    let components = package.state.source.components();
    let root = root_references(components).map_err(map_package_error)?;
    let body_identifier = root.body.ok_or(SectionNameError::UnsupportedSource)?;
    let body =
        find_object(components, body_identifier.get()).ok_or(SectionNameError::InvalidSource)?;
    let max_text_bytes = effective_text_limit(package.state.source.limits());
    let (native, payload) = decode_body_storage(
        &body.messages,
        body_identifier,
        MAX_SECTIONS,
        max_text_bytes,
    )
    .map_err(map_package_error)?;
    validate_section_table_wire(payload, &native, body_identifier).map_err(map_package_error)?;
    let section_references = native_section_references(&native, root.initial_section, MAX_SECTIONS)
        .map_err(map_package_error)?;
    section_references
        .get(position.get())
        .copied()
        .map(|NativeSectionReference { identifier, .. }| identifier)
        .ok_or(SectionNameError::InvalidSource)
}

fn section_name_at(
    package: &Package,
    position: Position,
) -> Result<Option<&str>, SectionNameError> {
    package
        .sections()
        .get(position.get())
        .map(|section| section.name())
        .ok_or(SectionNameError::PositionNotFound { position })
}

fn strict_payload_name(
    payload: &[u8],
    limits: WireLimits,
) -> Result<Option<&str>, SectionNameError> {
    let view = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    let mut name = None;
    for field in view
        .fields()
        .filter(|field| field.number() == SECTION_NAME_FIELD)
    {
        if name.is_some() || field.wire_type() != 2 {
            return Err(SectionNameError::InvalidSource);
        }
        field.validate_canonical_framing().map_err(map_wire_error)?;
        name = Some(
            std::str::from_utf8(field.payload())
                .map_err(|_error| SectionNameError::InvalidSource)?,
        );
    }
    Ok(name)
}

fn rewrite_name_payload(
    source: &[u8],
    before: Option<&str>,
    after: Option<&str>,
    limits: WireLimits,
) -> Result<Vec<u8>, SectionNameError> {
    let view = WireView::parse_with_limits(source, limits).map_err(map_wire_error)?;
    let mut existing_raw_length = 0usize;
    let mut found = false;
    for field in view
        .fields()
        .filter(|field| field.number() == SECTION_NAME_FIELD)
    {
        if found || field.wire_type() != 2 {
            return Err(SectionNameError::InvalidSource);
        }
        found = true;
        field.validate_canonical_framing().map_err(map_wire_error)?;
        let decoded = std::str::from_utf8(field.payload())
            .map_err(|_error| SectionNameError::InvalidSource)?;
        if Some(decoded) != before {
            return Err(SectionNameError::InvalidSource);
        }
        existing_raw_length = field.raw().len();
    }
    if found != before.is_some() {
        return Err(SectionNameError::InvalidSource);
    }

    let replacement_length = after.map_or(0, encoded_name_length);
    let output_length = source
        .len()
        .checked_sub(existing_raw_length)
        .and_then(|length| length.checked_add(replacement_length))
        .ok_or_else(|| output_limit_error(usize::MAX, limits))?;
    if output_length > limits.max_output_bytes() {
        return Err(output_limit_error(output_length, limits));
    }
    let output_fields = if found && after.is_none() {
        view.len().saturating_sub(1)
    } else if !found && after.is_some() {
        view.len()
            .checked_add(1)
            .ok_or(SectionNameError::InvalidSource)?
    } else {
        view.len()
    };
    if output_fields > limits.max_fields() {
        return Err(SectionNameError::LimitExceeded {
            kind: SectionNameLimitKind::WireFields,
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
    let mut emitted = false;
    for field in view.fields() {
        if field.number() == SECTION_NAME_FIELD {
            if let Some(name) = after {
                append_name_field(&mut output, name);
            }
            emitted = true;
        } else {
            output.extend_from_slice(field.raw());
        }
    }
    if !emitted && let Some(name) = after {
        append_name_field(&mut output, name);
    }
    if output.len() != output_length {
        return Err(SectionNameError::Verification);
    }
    Ok(output)
}

fn append_name_field(output: &mut Vec<u8>, name: &str) {
    encode_varint_into(output, (u64::from(SECTION_NAME_FIELD) << 3) | 2);
    encode_varint_into(output, usize_to_u64(name.len()));
    output.extend_from_slice(name.as_bytes());
}

fn encoded_name_length(name: &str) -> usize {
    let key = (u64::from(SECTION_NAME_FIELD) << 3) | 2;
    encoded_len(key)
        .saturating_add(encoded_len(usize_to_u64(name.len())))
        .saturating_add(name.len())
}

fn wire_limits(package: &Package) -> Result<WireLimits, SectionNameError> {
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

fn validate_name_memory(
    package: &Package,
    position: Position,
    replacement: Option<&str>,
) -> Result<(), SectionNameError> {
    let limit = effective_text_limit(package.state.source.limits());
    let mut observed = package
        .sections()
        .len()
        .checked_mul(size_of::<Option<Box<str>>>())
        .ok_or_else(|| name_limit_error(usize::MAX, limit))?;
    for section in package.sections() {
        let length = if section.index() == position.get() {
            replacement.map_or(0, str::len)
        } else {
            section.name().map_or(0, str::len)
        };
        observed = observed
            .checked_add(length)
            .ok_or_else(|| name_limit_error(usize::MAX, limit))?;
    }
    if observed > limit {
        return Err(name_limit_error(observed, limit));
    }
    Ok(())
}

fn verify_candidate(
    source: &Package,
    candidate: &Package,
    position: Position,
    expected: Option<&str>,
) -> Result<(), SectionNameError> {
    if source.stats().total_objects() != candidate.stats().total_objects()
        || source.sections().len() != candidate.sections().len()
    {
        return Err(SectionNameError::Verification);
    }
    for (index, (before, after)) in source
        .sections()
        .iter()
        .zip(candidate.sections())
        .enumerate()
    {
        let expected_name = if index == position.get() {
            expected
        } else {
            before.name()
        };
        if after.name() != expected_name
            || before.section_type() != after.section_type()
            || before.heading() != after.heading()
            || before.paragraphs() != after.paragraphs()
            || before.text_storages() != after.text_storages()
            || before.page_count() != after.page_count()
        {
            return Err(SectionNameError::Verification);
        }
    }
    Ok(())
}

fn try_owned_name(name: &str) -> Result<String, SectionNameError> {
    let mut owned = String::new();
    owned
        .try_reserve_exact(name.len())
        .map_err(|_allocation| SectionNameError::Allocation { amount: name.len() })?;
    owned.push_str(name);
    Ok(owned)
}

fn allocate_bytes(capacity: usize) -> Result<Vec<u8>, SectionNameError> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(capacity)
        .map_err(|_allocation| SectionNameError::Allocation { amount: capacity })?;
    Ok(output)
}

#[allow(
    clippy::needless_pass_by_value,
    reason = "used directly as a Result::map_err conversion"
)]
fn map_selector_error(selection_error: crate::SelectorError) -> SectionNameError {
    match selection_error {
        crate::SelectorError::AmbiguousSectionName {
            first, duplicate, ..
        } => SectionNameError::AmbiguousSelector {
            first: Position::new(first),
            duplicate: Position::new(duplicate),
        },
    }
}

#[allow(
    clippy::needless_pass_by_value,
    reason = "used directly as a Result::map_err conversion"
)]
fn map_package_error(package_error: PackageError) -> SectionNameError {
    match package_error {
        PackageError::Archive(archive_error) => map_archive_error(archive_error),
        PackageError::SectionNamesTooLarge { observed, limit } => name_limit_error(observed, limit),
        PackageError::Io(_) | PackageError::InvalidFormat(_) | PackageError::Semantic(_) => {
            SectionNameError::InvalidSource
        },
    }
}

#[allow(
    clippy::needless_pass_by_value,
    reason = "used directly as a Result::map_err conversion"
)]
fn map_archive_error(archive_error: litchi_iwa_archive::Error) -> SectionNameError {
    match archive_error {
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => SectionNameError::LimitExceeded {
            kind: match kind {
                litchi_iwa_archive::LimitKind::InputBytes => SectionNameLimitKind::InputBytes,
                litchi_iwa_archive::LimitKind::OutputBytes => SectionNameLimitKind::OutputBytes,
                litchi_iwa_archive::LimitKind::Entries => SectionNameLimitKind::Entries,
                litchi_iwa_archive::LimitKind::MemberNameBytes
                | litchi_iwa_archive::LimitKind::MetadataBytes
                | litchi_iwa_archive::LimitKind::CompressedEntryBytes
                | litchi_iwa_archive::LimitKind::EntryBytes
                | litchi_iwa_archive::LimitKind::IwaStreamBytes => SectionNameLimitKind::EntryBytes,
                litchi_iwa_archive::LimitKind::TotalBytes
                | litchi_iwa_archive::LimitKind::IwaTotalBytes => SectionNameLimitKind::TotalBytes,
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => {
            SectionNameError::Allocation { amount }
        },
        litchi_iwa_archive::Error::Iwa(core_error) => map_core_error(core_error),
        litchi_iwa_archive::Error::Io(_)
        | litchi_iwa_archive::Error::Zip { .. }
        | litchi_iwa_archive::Error::InvalidLimits(_)
        | litchi_iwa_archive::Error::Encrypted
        | litchi_iwa_archive::Error::SourceChanged { .. }
        | litchi_iwa_archive::Error::DirectoryChanged { .. }
        | litchi_iwa_archive::Error::Reassembly(_)
        | litchi_iwa_archive::Error::InvalidBundle(_) => SectionNameError::InvalidSource,
    }
}

#[allow(
    clippy::needless_pass_by_value,
    reason = "used directly as a Result::map_err conversion"
)]
fn map_core_error(core_error: litchi_iwa_core::Error) -> SectionNameError {
    match core_error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => SectionNameError::LimitExceeded {
            kind: match kind {
                litchi_iwa_core::LimitKind::Objects
                | litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject
                | litchi_iwa_core::LimitKind::MetadataItems => SectionNameLimitKind::Entries,
                litchi_iwa_core::LimitKind::MessageBytes => SectionNameLimitKind::WireBytes,
                litchi_iwa_core::LimitKind::HeaderFields => SectionNameLimitKind::WireFields,
                litchi_iwa_core::LimitKind::HeaderNesting => SectionNameLimitKind::WireNesting,
                litchi_iwa_core::LimitKind::ArchiveBytes
                | litchi_iwa_core::LimitKind::ObjectBytes
                | litchi_iwa_core::LimitKind::HeaderBytes
                | litchi_iwa_core::LimitKind::HeaderMemoryBytes
                | litchi_iwa_core::LimitKind::SnappyChunkBytes
                | litchi_iwa_core::LimitKind::SnappyStreamBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedChunkBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedStreamBytes
                | litchi_iwa_core::LimitKind::SnappyFrames => SectionNameLimitKind::EntryBytes,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            SectionNameError::Allocation { amount: requested }
        },
        litchi_iwa_core::Error::InvalidArchive { .. }
        | litchi_iwa_core::Error::InvalidLimits { .. }
        | litchi_iwa_core::Error::HeaderCodec { .. }
        | litchi_iwa_core::Error::Io(_)
        | litchi_iwa_core::Error::Snappy { .. } => SectionNameError::InvalidSource,
    }
}

#[allow(
    clippy::needless_pass_by_value,
    reason = "used directly as a Result::map_err conversion"
)]
fn map_wire_error(wire_error: litchi_iwa_common::Error) -> SectionNameError {
    match wire_error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => SectionNameError::LimitExceeded {
            kind: match kind {
                litchi_iwa_common::LimitKind::InputBytes
                | litchi_iwa_common::LimitKind::OutputBytes => SectionNameLimitKind::WireBytes,
                litchi_iwa_common::LimitKind::Fields
                | litchi_iwa_common::LimitKind::TableRows
                | litchi_iwa_common::LimitKind::TableColumns
                | litchi_iwa_common::LimitKind::TableCells
                | litchi_iwa_common::LimitKind::MaterializedCells => {
                    SectionNameLimitKind::WireFields
                },
                litchi_iwa_common::LimitKind::Nesting => SectionNameLimitKind::WireNesting,
                litchi_iwa_common::LimitKind::RewriteWork => SectionNameLimitKind::WireWork,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => {
            SectionNameError::Allocation { amount }
        },
        litchi_iwa_common::Error::InvalidFormat(_)
        | litchi_iwa_common::Error::InvalidLimit { .. } => SectionNameError::InvalidSource,
    }
}

fn name_limit_error(observed: usize, limit: usize) -> SectionNameError {
    SectionNameError::LimitExceeded {
        kind: SectionNameLimitKind::NameBytes,
        observed: usize_to_u64(observed),
        maximum: usize_to_u64(limit),
    }
}

fn output_limit_error(observed: usize, limits: WireLimits) -> SectionNameError {
    SectionNameError::LimitExceeded {
        kind: SectionNameLimitKind::WireBytes,
        observed: usize_to_u64(observed),
        maximum: usize_to_u64(limits.max_output_bytes()),
    }
}

fn work_limit_error(observed: usize, limits: WireLimits) -> SectionNameError {
    SectionNameError::LimitExceeded {
        kind: SectionNameLimitKind::WireWork,
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
