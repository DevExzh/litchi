//! Strict lazy PackageMetadata media-lifecycle records.
//!
//! This module owns only the small metadata closure needed by a format owner
//! while it edits an existing media item.  It deliberately does not own
//! package data members, object archives, or media bytes.  `PackageMetadata`
//! remains a caller-owned byte slice: repeated `DataInfo`, `ComponentInfo`,
//! and `ObjectReference` records are streamed by the handwritten scanner and
//! are never materialized through a generated repeated view.
//!
//! The writer is a two-pass, raw-preserving transaction.  The first pass
//! validates the complete source and calculates an exact output size; the
//! second pass writes one fallibly-reserved candidate and validates it again
//! before returning it.  Unknown fields in untouched records retain their
//! source bytes.  A record selected for mutation is rejected when its unknown
//! fields make ownership ambiguous.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "Wire helpers are kept beside the streaming publication model."
)]

use core::{fmt, num::NonZeroU64, str};
use std::path::{Component, Path};

use buffa::DecodeOptions as BuffaDecodeOptions;

use crate::buffa_package_metadata_media_generated::LitchiIwaPackageMetadataMediaProjection as projection;

const MAX_RECURSION: u32 = 64;
const MAX_FIELD_NUMBER: u32 = 0x1fff_ffff;
const SHA1_DIGEST_BYTES: usize = 20;
const ROOT_LAST_IDENTIFIER_FIELD: u32 = 1;
const ROOT_COMPONENT_FIELD: u32 = 3;
const ROOT_DATA_INFO_FIELD: u32 = 4;
const ROOT_SAVE_TOKEN_FIELD: u32 = 8;
const ROOT_DATA_METADATA_MAP_FIELD: u32 = 10;
const ROOT_VERSIONED_COMPONENT_FIELD: u32 = 11;
const COMPONENT_IDENTIFIER_FIELD: u32 = 1;
const COMPONENT_PREFERRED_LOCATOR_FIELD: u32 = 2;
const COMPONENT_LOCATOR_FIELD: u32 = 3;
const COMPONENT_DATA_REFERENCE_FIELD: u32 = 7;
const DATA_IDENTIFIER_FIELD: u32 = 1;
const DATA_DIGEST_FIELD: u32 = 2;
const DATA_PREFERRED_NAME_FIELD: u32 = 3;
const DATA_FILE_NAME_FIELD: u32 = 4;
const DATA_MATERIALIZED_LENGTH_FIELD: u32 = 18;
const OWNER_OBJECT_IDENTIFIER_FIELD: u32 = 1;
const OWNER_COUNT_FIELD: u32 = 2;

/// Finite limits for one PackageMetadata media scan and rewrite.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeOptions {
    max_message_bytes: usize,
    max_output_bytes: usize,
    max_fields: usize,
    max_work_bytes: usize,
    max_components: usize,
    max_data_records: usize,
    max_owners: usize,
    max_digest_bytes: usize,
    max_name_bytes: usize,
    max_depth: u32,
}

impl DecodeOptions {
    /// Construct explicit byte, field, work, topology, text, and nesting
    /// limits.  All limits are operation-local and are checked before any
    /// output allocation.
    #[must_use]
    #[allow(clippy::too_many_arguments)]
    pub const fn new(
        max_message_bytes: usize,
        max_fields: usize,
        max_work_bytes: usize,
        max_components: usize,
        max_data_records: usize,
        max_owners: usize,
        max_digest_bytes: usize,
        max_name_bytes: usize,
        max_depth: u32,
    ) -> Self {
        Self {
            max_message_bytes,
            max_output_bytes: max_message_bytes,
            max_fields,
            max_work_bytes,
            max_components,
            max_data_records,
            max_owners,
            max_digest_bytes,
            max_name_bytes,
            max_depth,
        }
    }

    /// Build a bounded policy sized for one already-borrowed source payload.
    #[must_use]
    pub fn for_source(source: &[u8]) -> Self {
        let bytes = source.len().max(1);
        let fields = bytes.saturating_mul(8).max(1);
        Self::new(
            bytes,
            fields,
            bytes.saturating_mul(16).max(1),
            fields,
            fields,
            fields,
            SHA1_DIGEST_BYTES,
            4096,
            16,
        )
        .with_max_output_bytes(bytes.saturating_mul(2).max(1))
    }

    #[must_use]
    pub const fn max_message_bytes(self) -> usize {
        self.max_message_bytes
    }
    #[must_use]
    pub const fn with_max_message_bytes(mut self, maximum: usize) -> Self {
        self.max_message_bytes = maximum;
        self
    }
    #[must_use]
    pub const fn max_output_bytes(self) -> usize {
        self.max_output_bytes
    }
    #[must_use]
    pub const fn with_max_output_bytes(mut self, maximum: usize) -> Self {
        self.max_output_bytes = maximum;
        self
    }
    #[must_use]
    pub const fn max_fields(self) -> usize {
        self.max_fields
    }
    #[must_use]
    pub const fn max_work_bytes(self) -> usize {
        self.max_work_bytes
    }
    #[must_use]
    pub const fn with_max_work_bytes(mut self, maximum: usize) -> Self {
        self.max_work_bytes = maximum;
        self
    }
    #[must_use]
    pub const fn max_components(self) -> usize {
        self.max_components
    }
    #[must_use]
    pub const fn max_data_records(self) -> usize {
        self.max_data_records
    }
    #[must_use]
    pub const fn max_owners(self) -> usize {
        self.max_owners
    }
    #[must_use]
    pub const fn max_digest_bytes(self) -> usize {
        self.max_digest_bytes
    }
    #[must_use]
    pub const fn max_name_bytes(self) -> usize {
        self.max_name_bytes
    }
    #[must_use]
    pub const fn max_depth(self) -> u32 {
        self.max_depth
    }

    fn buffa(self) -> BuffaDecodeOptions {
        BuffaDecodeOptions::new()
            .with_max_message_size(self.max_message_bytes)
            .with_unknown_field_limit(self.max_fields)
            .with_element_memory_limit(0)
            .with_recursion_limit(self.max_depth)
    }
}

/// Resource classification for a strict media scan.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum DecodeLimit {
    /// Source or nested message bytes exceed the configured ceiling.
    Bytes { observed: usize, maximum: usize },
    /// Wire fields visited exceed the configured ceiling.
    Fields { observed: usize, maximum: usize },
    /// Strict plus Buffa work exceeds the configured ceiling.
    Work { observed: usize, maximum: usize },
    /// Candidate output bytes exceed the configured output ceiling.
    OutputBytes { observed: usize, maximum: usize },
    /// Component records exceed the configured ceiling.
    Components { observed: usize, maximum: usize },
    /// DataInfo records exceed the configured ceiling.
    DataRecords { observed: usize, maximum: usize },
    /// Component owner records exceed the configured ceiling.
    Owners { observed: usize, maximum: usize },
    /// One digest exceeds the configured ceiling.
    DigestBytes { observed: usize, maximum: usize },
    /// One filename exceeds the configured ceiling.
    NameBytes { observed: usize, maximum: usize },
    /// Configured or observed nesting exceeds the finite ceiling.
    Nesting { observed: u32, maximum: u32 },
}

/// Strict media metadata scan failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeError {
    limit: Option<DecodeLimit>,
    reason: Option<InvalidReason>,
}

impl DecodeError {
    /// Construct a content-free invalid-source error for a strict adapter
    /// visitor that cannot retain a richer parser diagnostic.
    #[doc(hidden)]
    #[must_use]
    pub const fn invalid_for_adapter() -> Self {
        Self::invalid(InvalidReason::Verification)
    }

    #[must_use]
    pub const fn resource_limit(self) -> Option<DecodeLimit> {
        self.limit
    }

    #[must_use]
    pub const fn invalid_reason(self) -> Option<InvalidReason> {
        self.reason
    }

    const fn limited(limit: DecodeLimit) -> Self {
        Self {
            limit: Some(limit),
            reason: None,
        }
    }

    const fn invalid(reason: InvalidReason) -> Self {
        Self {
            limit: None,
            reason: Some(reason),
        }
    }
}

impl fmt::Display for DecodeError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("invalid PackageMetadata media records")
    }
}

impl std::error::Error for DecodeError {}

/// Content-free semantic refusal from the strict media scanner.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum InvalidReason {
    MalformedWire,
    MissingRequiredField,
    DuplicateField,
    InvalidIdentifier,
    InvalidDigest,
    InvalidName,
    UnknownSelectedRecord,
    DuplicateDataInfo,
    DuplicateComponent,
    DuplicateOwner,
    ComponentNotFound,
    ComponentAmbiguous,
    VersionedComponent,
    DataInfoNotFound,
    DataInfoReferenced,
    DataInfoMetadataMapDependency,
    OwnerNotFound,
    DuplicateOperation,
    ConflictingOperation,
    ExistingDataCollision,
    ExistingOwnerCollision,
    UnsupportedField,
    Verification,
}

/// Borrowed exact component selector.  `locator` is compared against the
/// effective native locator (explicit `locator`, or `preferred_locator` when
/// the optional field is absent).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ComponentSelector<'source> {
    identifier: u64,
    locator: &'source str,
}

impl<'source> ComponentSelector<'source> {
    #[must_use]
    pub const fn new(identifier: u64, locator: &'source str) -> Self {
        Self {
            identifier,
            locator,
        }
    }

    #[must_use]
    pub const fn identifier(self) -> u64 {
        self.identifier
    }

    #[must_use]
    pub const fn locator(self) -> &'source str {
        self.locator
    }
}

/// Borrowed DataInfo facts.  All slices point into the caller-owned source.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DataInfoSnapshot<'source> {
    identifier: u64,
    digest: &'source [u8],
    preferred_file_name: &'source str,
    file_name: Option<&'source str>,
    materialized_length: Option<u64>,
    raw: &'source [u8],
    unknown_fields: bool,
}

impl<'source> DataInfoSnapshot<'source> {
    #[must_use]
    pub const fn identifier(self) -> u64 {
        self.identifier
    }
    #[must_use]
    pub const fn digest(self) -> &'source [u8] {
        self.digest
    }
    #[must_use]
    pub const fn preferred_file_name(self) -> &'source str {
        self.preferred_file_name
    }
    #[must_use]
    pub const fn file_name(self) -> Option<&'source str> {
        self.file_name
    }
    #[must_use]
    pub const fn materialized_length(self) -> Option<u64> {
        self.materialized_length
    }
    #[must_use]
    pub const fn raw(self) -> &'source [u8] {
        self.raw
    }
    #[must_use]
    pub const fn has_unknown_fields(self) -> bool {
        self.unknown_fields
    }
}

/// Borrowed component facts needed to authorize a media-owner mutation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ComponentSnapshot<'source> {
    identifier: u64,
    preferred_locator: &'source str,
    locator: Option<&'source str>,
    versioned: bool,
    unknown_fields: bool,
    raw: &'source [u8],
}

impl<'source> ComponentSnapshot<'source> {
    #[must_use]
    pub const fn identifier(self) -> u64 {
        self.identifier
    }
    #[must_use]
    pub const fn preferred_locator(self) -> &'source str {
        self.preferred_locator
    }
    #[must_use]
    pub const fn locator(self) -> Option<&'source str> {
        self.locator
    }
    #[must_use]
    pub const fn effective_locator(self) -> &'source str {
        match self.locator {
            Some(locator) => locator,
            None => self.preferred_locator,
        }
    }
    #[must_use]
    pub const fn is_versioned(self) -> bool {
        self.versioned
    }
    #[must_use]
    pub const fn has_unknown_fields(self) -> bool {
        self.unknown_fields
    }
    #[must_use]
    pub const fn raw(self) -> &'source [u8] {
        self.raw
    }
}

/// Borrowed ComponentDataReference facts.  Owners are delivered separately
/// through the visitor, avoiding an input-width allocation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ComponentDataReferenceSnapshot<'source> {
    data_identifier: u64,
    owner_count: usize,
    raw: &'source [u8],
    unknown_fields: bool,
}

impl<'source> ComponentDataReferenceSnapshot<'source> {
    #[must_use]
    pub const fn data_identifier(self) -> u64 {
        self.data_identifier
    }
    #[must_use]
    pub const fn owner_count(self) -> usize {
        self.owner_count
    }
    #[must_use]
    pub const fn raw(self) -> &'source [u8] {
        self.raw
    }
    #[must_use]
    pub const fn has_unknown_fields(self) -> bool {
        self.unknown_fields
    }
}

/// One borrowed ComponentDataReference.ObjectReference owner.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct OwnerSnapshot<'source> {
    object_identifier: u64,
    count: u32,
    raw: &'source [u8],
    unknown_fields: bool,
}

impl<'source> OwnerSnapshot<'source> {
    #[must_use]
    pub const fn object_identifier(self) -> u64 {
        self.object_identifier
    }
    #[must_use]
    pub const fn count(self) -> u32 {
        self.count
    }
    #[must_use]
    pub const fn raw(self) -> &'source [u8] {
        self.raw
    }
    #[must_use]
    pub const fn has_unknown_fields(self) -> bool {
        self.unknown_fields
    }
}

/// Fallible streaming sink for PackageMetadata media metadata.
pub trait PackageMetadataMediaVisitor {
    fn visit_data_info(&mut self, _data_info: DataInfoSnapshot<'_>) -> Result<(), DecodeError> {
        Ok(())
    }

    fn visit_component(&mut self, _component: ComponentSnapshot<'_>) -> Result<(), DecodeError> {
        Ok(())
    }

    fn visit_data_reference(
        &mut self,
        _component: ComponentSnapshot<'_>,
        _data_reference: ComponentDataReferenceSnapshot<'_>,
    ) -> Result<(), DecodeError> {
        Ok(())
    }

    fn visit_owner(
        &mut self,
        _component: ComponentSnapshot<'_>,
        _data_reference: ComponentDataReferenceSnapshot<'_>,
        _owner: OwnerSnapshot<'_>,
    ) -> Result<(), DecodeError> {
        Ok(())
    }
}

/// Exact finite consumption from one successful scan.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeReport {
    input_bytes: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    components: usize,
    data_records: usize,
    owners: usize,
    unknown_records: usize,
    data_metadata_map_present: bool,
}

impl DecodeReport {
    #[must_use]
    pub const fn input_bytes(self) -> usize {
        self.input_bytes
    }
    #[must_use]
    pub const fn fields(self) -> usize {
        self.fields
    }
    #[must_use]
    pub const fn work_bytes(self) -> usize {
        self.work_bytes
    }
    #[must_use]
    pub const fn max_depth(self) -> u32 {
        self.max_depth
    }
    #[must_use]
    pub const fn components(self) -> usize {
        self.components
    }
    #[must_use]
    pub const fn data_records(self) -> usize {
        self.data_records
    }
    #[must_use]
    pub const fn owners(self) -> usize {
        self.owners
    }
    #[must_use]
    pub const fn unknown_records(self) -> usize {
        self.unknown_records
    }
    #[must_use]
    pub const fn data_metadata_map_present(self) -> bool {
        self.data_metadata_map_present
    }
}

/// Scan the metadata closure while streaming records to `visitor`.
pub fn visit_package_metadata_media(
    source: &[u8],
    options: DecodeOptions,
    visitor: &mut dyn PackageMetadataMediaVisitor,
) -> Result<DecodeReport, DecodeError> {
    scan(source, options, visitor)
}

/// Scan without retaining any record and return exact topology facts.
pub fn inspect_package_metadata_media(
    source: &[u8],
    options: DecodeOptions,
) -> Result<DecodeReport, DecodeError> {
    let mut visitor = NoopVisitor;
    scan(source, options, &mut visitor)
}

struct NoopVisitor;

impl PackageMetadataMediaVisitor for NoopVisitor {}

#[derive(Debug, Clone, Copy)]
struct ScanState {
    input_bytes: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    components: usize,
    data_records: usize,
    owners: usize,
    unknown_records: usize,
    data_metadata_map_present: bool,
}

impl ScanState {
    fn new(source: &[u8], options: DecodeOptions) -> Result<Self, DecodeError> {
        if source.len() > options.max_message_bytes {
            return Err(DecodeError::limited(DecodeLimit::Bytes {
                observed: source.len(),
                maximum: options.max_message_bytes,
            }));
        }
        if options.max_depth == 0 || options.max_depth > MAX_RECURSION {
            return Err(DecodeError::limited(DecodeLimit::Nesting {
                observed: options.max_depth,
                maximum: MAX_RECURSION,
            }));
        }
        Ok(Self {
            input_bytes: source.len(),
            fields: 0,
            work_bytes: 0,
            max_depth: 0,
            components: 0,
            data_records: 0,
            owners: 0,
            unknown_records: 0,
            data_metadata_map_present: false,
        })
    }

    fn field(
        &mut self,
        field: WireField,
        options: DecodeOptions,
        depth: u32,
    ) -> Result<(), DecodeError> {
        self.fields = self
            .fields
            .checked_add(1)
            .ok_or_else(|| DecodeError::invalid(InvalidReason::MalformedWire))?;
        if self.fields > options.max_fields {
            return Err(DecodeError::limited(DecodeLimit::Fields {
                observed: self.fields,
                maximum: options.max_fields,
            }));
        }
        self.work_bytes = self
            .work_bytes
            .checked_add(field.end.saturating_sub(field.start))
            .ok_or_else(|| DecodeError::invalid(InvalidReason::MalformedWire))?;
        if self.work_bytes > options.max_work_bytes {
            return Err(DecodeError::limited(DecodeLimit::Work {
                observed: self.work_bytes,
                maximum: options.max_work_bytes,
            }));
        }
        self.max_depth = self.max_depth.max(depth);
        if depth > options.max_depth {
            return Err(DecodeError::limited(DecodeLimit::Nesting {
                observed: depth,
                maximum: options.max_depth,
            }));
        }
        Ok(())
    }

    fn work(&mut self, amount: usize, options: DecodeOptions) -> Result<(), DecodeError> {
        self.work_bytes = self
            .work_bytes
            .checked_add(amount)
            .ok_or_else(|| DecodeError::invalid(InvalidReason::MalformedWire))?;
        if self.work_bytes > options.max_work_bytes {
            return Err(DecodeError::limited(DecodeLimit::Work {
                observed: self.work_bytes,
                maximum: options.max_work_bytes,
            }));
        }
        Ok(())
    }

    fn component(&mut self, options: DecodeOptions) -> Result<(), DecodeError> {
        self.components = self
            .components
            .checked_add(1)
            .ok_or_else(|| DecodeError::invalid(InvalidReason::MalformedWire))?;
        if self.components > options.max_components {
            return Err(DecodeError::limited(DecodeLimit::Components {
                observed: self.components,
                maximum: options.max_components,
            }));
        }
        Ok(())
    }

    fn data_record(&mut self, options: DecodeOptions) -> Result<(), DecodeError> {
        self.data_records = self
            .data_records
            .checked_add(1)
            .ok_or_else(|| DecodeError::invalid(InvalidReason::MalformedWire))?;
        if self.data_records > options.max_data_records {
            return Err(DecodeError::limited(DecodeLimit::DataRecords {
                observed: self.data_records,
                maximum: options.max_data_records,
            }));
        }
        Ok(())
    }

    fn owner(&mut self, options: DecodeOptions) -> Result<(), DecodeError> {
        self.owners = self
            .owners
            .checked_add(1)
            .ok_or_else(|| DecodeError::invalid(InvalidReason::MalformedWire))?;
        if self.owners > options.max_owners {
            return Err(DecodeError::limited(DecodeLimit::Owners {
                observed: self.owners,
                maximum: options.max_owners,
            }));
        }
        Ok(())
    }

    fn unknown(&mut self) -> Result<(), DecodeError> {
        self.unknown_records = self
            .unknown_records
            .checked_add(1)
            .ok_or_else(|| DecodeError::invalid(InvalidReason::MalformedWire))?;
        Ok(())
    }

    fn report(self) -> DecodeReport {
        DecodeReport {
            input_bytes: self.input_bytes,
            fields: self.fields,
            work_bytes: self.work_bytes,
            max_depth: self.max_depth,
            components: self.components,
            data_records: self.data_records,
            owners: self.owners,
            unknown_records: self.unknown_records,
            data_metadata_map_present: self.data_metadata_map_present,
        }
    }
}

fn scan(
    source: &[u8],
    options: DecodeOptions,
    visitor: &mut dyn PackageMetadataMediaVisitor,
) -> Result<DecodeReport, DecodeError> {
    let mut state = ScanState::new(source, options)?;
    validate_unique_records(source)?;
    let mut last_identifier = None;
    let mut save_token_seen = false;
    let mut offset = 0usize;
    while let Some(field) = next_field(source, offset)? {
        offset = field.end;
        state.field(field, options, 1)?;
        match field.number {
            ROOT_LAST_IDENTIFIER_FIELD => {
                if last_identifier.is_some() {
                    return Err(DecodeError::invalid(InvalidReason::DuplicateField));
                }
                let identifier = varint(source, field)?;
                if NonZeroU64::new(identifier).is_none() {
                    return Err(DecodeError::invalid(InvalidReason::InvalidIdentifier));
                }
                last_identifier = Some(identifier);
            },
            ROOT_COMPONENT_FIELD | ROOT_VERSIONED_COMPONENT_FIELD => {
                state.component(options)?;
                parse_component(
                    source,
                    field.payload(source),
                    options,
                    &mut state,
                    field.number == ROOT_VERSIONED_COMPONENT_FIELD,
                    visitor,
                )?;
            },
            ROOT_DATA_INFO_FIELD => {
                state.data_record(options)?;
                parse_data_info(source, field.payload(source), options, &mut state, visitor)?;
            },
            ROOT_SAVE_TOKEN_FIELD => {
                if save_token_seen {
                    return Err(DecodeError::invalid(InvalidReason::DuplicateField));
                }
                let _ = varint(source, field)?;
                save_token_seen = true;
            },
            ROOT_DATA_METADATA_MAP_FIELD => {
                if state.data_metadata_map_present {
                    return Err(DecodeError::invalid(InvalidReason::DuplicateField));
                }
                let payload = bytes(source, field)?;
                parse_reference(payload, options, &mut state)?;
                state.data_metadata_map_present = true;
            },
            2 => {
                let _ = bytes(source, field)?;
            },
            5..=7 => {
                validate_packed_varints(bytes(source, field)?)?;
            },
            9 => {
                let value = varint(source, field)?;
                if value > 2 {
                    return Err(DecodeError::invalid(InvalidReason::UnsupportedField));
                }
            },
            _ => {
                // Unknown root fields remain byte-authoritative and are
                // intentionally preserved by the rewriter.
                state.unknown()?;
            },
        }
    }
    if last_identifier.is_none() {
        return Err(DecodeError::invalid(InvalidReason::MissingRequiredField));
    }
    Ok(state.report())
}

/// Validate duplicate identity keys without retaining an input-width vector.
/// Metadata records are small and already bounded by `DecodeOptions`; the
/// deliberately simple rescans keep the streaming visitor allocation-free.
fn validate_unique_records(source: &[u8]) -> Result<(), DecodeError> {
    let mut root_offset = 0usize;
    while let Some(field) = next_field(source, root_offset)? {
        root_offset = field.end;
        match field.number {
            ROOT_DATA_INFO_FIELD => {
                let current = data_info_facts(field.payload(source))?;
                let mut count = 0usize;
                let mut offset = 0usize;
                while let Some(candidate) = next_field(source, offset)? {
                    offset = candidate.end;
                    if candidate.number == ROOT_DATA_INFO_FIELD
                        && data_info_facts(candidate.payload(source))?.identifier
                            == current.identifier
                    {
                        count = count
                            .checked_add(1)
                            .ok_or_else(|| DecodeError::invalid(InvalidReason::MalformedWire))?;
                    }
                }
                if count > 1 {
                    return Err(DecodeError::invalid(InvalidReason::DuplicateDataInfo));
                }
            },
            ROOT_COMPONENT_FIELD | ROOT_VERSIONED_COMPONENT_FIELD => {
                let current = component_facts(
                    field.payload(source),
                    field.number == ROOT_VERSIONED_COMPONENT_FIELD,
                )?;
                let mut count = 0usize;
                let mut offset = 0usize;
                while let Some(candidate) = next_field(source, offset)? {
                    offset = candidate.end;
                    if (candidate.number == ROOT_COMPONENT_FIELD
                        || candidate.number == ROOT_VERSIONED_COMPONENT_FIELD)
                        && {
                            let candidate_facts = component_facts(
                                candidate.payload(source),
                                candidate.number == ROOT_VERSIONED_COMPONENT_FIELD,
                            )?;
                            candidate_facts.identifier == current.identifier
                                && candidate_facts.effective_locator()
                                    == current.effective_locator()
                        }
                    {
                        count = count
                            .checked_add(1)
                            .ok_or_else(|| DecodeError::invalid(InvalidReason::MalformedWire))?;
                    }
                }
                if count > 1 {
                    return Err(DecodeError::invalid(InvalidReason::DuplicateComponent));
                }
                validate_unique_component_references(field.payload(source))?;
            },
            _ => {},
        }
    }
    Ok(())
}

fn validate_unique_component_references(payload: &[u8]) -> Result<(), DecodeError> {
    let mut offset = 0usize;
    while let Some(field) = next_field(payload, offset)? {
        offset = field.end;
        if field.number != COMPONENT_DATA_REFERENCE_FIELD {
            continue;
        }
        let current = data_reference_facts(field.payload(payload))?;
        let mut parent_count = 0usize;
        let mut candidate_offset = 0usize;
        while let Some(candidate) = next_field(payload, candidate_offset)? {
            candidate_offset = candidate.end;
            if candidate.number != COMPONENT_DATA_REFERENCE_FIELD {
                continue;
            }
            let candidate_facts = data_reference_facts(candidate.payload(payload))?;
            if candidate_facts.data_identifier == current.data_identifier {
                parent_count = parent_count
                    .checked_add(1)
                    .ok_or_else(|| DecodeError::invalid(InvalidReason::MalformedWire))?;
                if parent_count > 1 {
                    return Err(DecodeError::invalid(InvalidReason::DuplicateDataInfo));
                }
            }
        }
        validate_unique_owners(field.payload(payload))?;
    }
    Ok(())
}

fn validate_unique_owners(payload: &[u8]) -> Result<(), DecodeError> {
    let mut offset = 0usize;
    while let Some(field) = next_field(payload, offset)? {
        offset = field.end;
        if field.number != OWNER_COUNT_FIELD {
            continue;
        }
        let current = owner_facts(field.payload(payload))?;
        let mut count = 0usize;
        let mut candidate_offset = 0usize;
        while let Some(candidate) = next_field(payload, candidate_offset)? {
            candidate_offset = candidate.end;
            if candidate.number != OWNER_COUNT_FIELD {
                continue;
            }
            let candidate_facts = owner_facts(candidate.payload(payload))?;
            if candidate_facts.object_identifier == current.object_identifier {
                count = count
                    .checked_add(1)
                    .ok_or_else(|| DecodeError::invalid(InvalidReason::MalformedWire))?;
            }
        }
        if count > 1 {
            return Err(DecodeError::invalid(InvalidReason::DuplicateOwner));
        }
    }
    Ok(())
}

fn parse_data_info(
    _source: &[u8],
    payload: &[u8],
    options: DecodeOptions,
    state: &mut ScanState,
    visitor: &mut dyn PackageMetadataMediaVisitor,
) -> Result<(), DecodeError> {
    if payload.len() > options.max_message_bytes {
        return Err(DecodeError::limited(DecodeLimit::Bytes {
            observed: payload.len(),
            maximum: options.max_message_bytes,
        }));
    }
    state.work(payload.len(), options)?;
    let mut identifier = None;
    let mut digest = None;
    let mut preferred_file_name = None;
    let mut file_name = None;
    let mut materialized_length = None;
    let mut unknown_fields = false;
    for result in fields(payload) {
        let field = result?;
        state.field(field, options, 2)?;
        match field.number {
            DATA_IDENTIFIER_FIELD => {
                if identifier.is_some() {
                    return Err(DecodeError::invalid(InvalidReason::DuplicateField));
                }
                let value = varint(payload, field)?;
                if NonZeroU64::new(value).is_none() {
                    return Err(DecodeError::invalid(InvalidReason::InvalidIdentifier));
                }
                identifier = Some(value);
            },
            DATA_DIGEST_FIELD => {
                if digest.is_some() {
                    return Err(DecodeError::invalid(InvalidReason::DuplicateField));
                }
                let value = bytes(payload, field)?;
                if value.len() > options.max_digest_bytes {
                    return Err(DecodeError::limited(DecodeLimit::DigestBytes {
                        observed: value.len(),
                        maximum: options.max_digest_bytes,
                    }));
                }
                if value.len() != SHA1_DIGEST_BYTES {
                    return Err(DecodeError::invalid(InvalidReason::InvalidDigest));
                }
                digest = Some(value);
            },
            DATA_PREFERRED_NAME_FIELD => {
                if preferred_file_name.is_some() {
                    return Err(DecodeError::invalid(InvalidReason::DuplicateField));
                }
                let value = utf8(payload, field)?;
                validate_required_name(value, options)?;
                preferred_file_name = Some(value);
            },
            DATA_FILE_NAME_FIELD => {
                if file_name.is_some() {
                    return Err(DecodeError::invalid(InvalidReason::DuplicateField));
                }
                let value = utf8(payload, field)?;
                validate_optional_name(value, options)?;
                file_name = Some(value);
            },
            DATA_MATERIALIZED_LENGTH_FIELD => {
                if materialized_length.is_some() {
                    return Err(DecodeError::invalid(InvalidReason::DuplicateField));
                }
                materialized_length = Some(varint(payload, field)?);
            },
            _ => {
                unknown_fields = true;
                state.unknown()?;
            },
        }
    }
    let snapshot = DataInfoSnapshot {
        identifier: identifier
            .ok_or_else(|| DecodeError::invalid(InvalidReason::MissingRequiredField))?,
        digest: digest.ok_or_else(|| DecodeError::invalid(InvalidReason::MissingRequiredField))?,
        preferred_file_name: preferred_file_name
            .ok_or_else(|| DecodeError::invalid(InvalidReason::MissingRequiredField))?,
        file_name,
        materialized_length,
        raw: payload,
        unknown_fields,
    };
    let view: projection::DataInfoArchiveLazyView<'_> =
        options
            .buffa()
            .decode_lazy_view(payload)
            .map_err(|_error| DecodeError::invalid(InvalidReason::MalformedWire))?;
    if view.identifier != snapshot.identifier
        || view.digest != snapshot.digest
        || view.preferred_file_name != snapshot.preferred_file_name
        || view.file_name != snapshot.file_name
        || view.materialized_length != snapshot.materialized_length
    {
        return Err(DecodeError::invalid(InvalidReason::Verification));
    }
    visitor.visit_data_info(snapshot)
}

fn parse_component(
    source: &[u8],
    payload: &[u8],
    options: DecodeOptions,
    state: &mut ScanState,
    versioned: bool,
    visitor: &mut dyn PackageMetadataMediaVisitor,
) -> Result<(), DecodeError> {
    if payload.len() > options.max_message_bytes {
        return Err(DecodeError::limited(DecodeLimit::Bytes {
            observed: payload.len(),
            maximum: options.max_message_bytes,
        }));
    }
    state.work(payload.len(), options)?;
    let mut identifier = None;
    let mut preferred_locator = None;
    let mut locator = None;
    let mut unknown_fields = false;
    for result in fields(payload) {
        let field = result?;
        state.field(field, options, 2)?;
        match field.number {
            COMPONENT_IDENTIFIER_FIELD => {
                if identifier.is_some() {
                    return Err(DecodeError::invalid(InvalidReason::DuplicateField));
                }
                let value = varint(payload, field)?;
                if NonZeroU64::new(value).is_none() {
                    return Err(DecodeError::invalid(InvalidReason::InvalidIdentifier));
                }
                identifier = Some(value);
            },
            COMPONENT_PREFERRED_LOCATOR_FIELD => {
                if preferred_locator.is_some() {
                    return Err(DecodeError::invalid(InvalidReason::DuplicateField));
                }
                let value = utf8(payload, field)?;
                validate_required_name(value, options)?;
                preferred_locator = Some(value);
            },
            COMPONENT_LOCATOR_FIELD => {
                if locator.is_some() {
                    return Err(DecodeError::invalid(InvalidReason::DuplicateField));
                }
                let value = utf8(payload, field)?;
                validate_optional_name(value, options)?;
                locator = Some(value);
            },
            4 | 5 | 14 | 15 | 20 => validate_packed_varints(bytes(payload, field)?)?,
            6 | 7 | 11 | 13 | 18 => {
                let _ = bytes(payload, field)?;
            },
            10 | 17 | 19 => {
                let value = varint(payload, field)?;
                if value > 1 {
                    return Err(DecodeError::invalid(InvalidReason::UnsupportedField));
                }
            },
            12 | 16 | 21 => {
                let _ = varint(payload, field)?;
            },
            _ => {
                unknown_fields = true;
                state.unknown()?;
            },
        }
    }
    let component = ComponentSnapshot {
        identifier: identifier
            .ok_or_else(|| DecodeError::invalid(InvalidReason::MissingRequiredField))?,
        preferred_locator: preferred_locator
            .ok_or_else(|| DecodeError::invalid(InvalidReason::MissingRequiredField))?,
        locator,
        versioned,
        unknown_fields,
        raw: payload,
    };
    visitor.visit_component(component)?;

    // Revisit only field-7 payloads to stream their owner records.  No
    // generated repeated view or owned nested collection is created.
    for result in fields(payload) {
        let field = result?;
        if field.number == COMPONENT_DATA_REFERENCE_FIELD {
            parse_data_reference(
                source,
                field.payload(payload),
                options,
                state,
                component,
                visitor,
            )?;
        }
    }
    Ok(())
}

fn parse_data_reference(
    _source: &[u8],
    payload: &[u8],
    options: DecodeOptions,
    state: &mut ScanState,
    component: ComponentSnapshot<'_>,
    visitor: &mut dyn PackageMetadataMediaVisitor,
) -> Result<(), DecodeError> {
    if payload.len() > options.max_message_bytes {
        return Err(DecodeError::limited(DecodeLimit::Bytes {
            observed: payload.len(),
            maximum: options.max_message_bytes,
        }));
    }
    state.work(payload.len(), options)?;
    let mut data_identifier = None;
    let mut owner_count = 0usize;
    let mut unknown_fields = false;
    for result in fields(payload) {
        let field = result?;
        state.field(field, options, 3)?;
        match field.number {
            DATA_IDENTIFIER_FIELD => {
                if data_identifier.is_some() {
                    return Err(DecodeError::invalid(InvalidReason::DuplicateField));
                }
                let value = varint(payload, field)?;
                if NonZeroU64::new(value).is_none() {
                    return Err(DecodeError::invalid(InvalidReason::InvalidIdentifier));
                }
                data_identifier = Some(value);
            },
            OWNER_COUNT_FIELD => {
                let owner = bytes(payload, field)?;
                state.owner(options)?;
                let _ = parse_owner(owner, options, state)?;
                owner_count = owner_count
                    .checked_add(1)
                    .ok_or_else(|| DecodeError::invalid(InvalidReason::MalformedWire))?;
            },
            _ => {
                unknown_fields = true;
                state.unknown()?;
            },
        }
    }
    let snapshot = ComponentDataReferenceSnapshot {
        data_identifier: data_identifier
            .ok_or_else(|| DecodeError::invalid(InvalidReason::MissingRequiredField))?,
        owner_count,
        raw: payload,
        unknown_fields,
    };
    let view: projection::ComponentDataReferenceArchiveLazyView<'_> = options
        .buffa()
        .decode_lazy_view(payload)
        .map_err(|_error| DecodeError::invalid(InvalidReason::MalformedWire))?;
    if view.data_identifier != snapshot.data_identifier {
        return Err(DecodeError::invalid(InvalidReason::Verification));
    }
    visitor.visit_data_reference(component, snapshot)?;

    for result in fields(payload) {
        let field = result?;
        if field.number == OWNER_COUNT_FIELD {
            let owner_payload = field.payload(payload);
            let owner = parse_owner(owner_payload, options, state)?;
            visitor.visit_owner(component, snapshot, owner)?;
        }
    }
    Ok(())
}

fn parse_owner<'source>(
    payload: &'source [u8],
    options: DecodeOptions,
    state: &mut ScanState,
) -> Result<OwnerSnapshot<'source>, DecodeError> {
    if payload.len() > options.max_message_bytes {
        return Err(DecodeError::limited(DecodeLimit::Bytes {
            observed: payload.len(),
            maximum: options.max_message_bytes,
        }));
    }
    state.work(payload.len(), options)?;
    let mut object_identifier = None;
    let mut count = None;
    let mut unknown_fields = false;
    for result in fields(payload) {
        let field = result?;
        state.field(field, options, 4)?;
        match field.number {
            OWNER_OBJECT_IDENTIFIER_FIELD => {
                if object_identifier.is_some() {
                    return Err(DecodeError::invalid(InvalidReason::DuplicateField));
                }
                let value = varint(payload, field)?;
                if NonZeroU64::new(value).is_none() {
                    return Err(DecodeError::invalid(InvalidReason::InvalidIdentifier));
                }
                object_identifier = Some(value);
            },
            OWNER_COUNT_FIELD => {
                if count.is_some() {
                    return Err(DecodeError::invalid(InvalidReason::DuplicateField));
                }
                let value = varint(payload, field)?;
                let value = u32::try_from(value)
                    .map_err(|_error| DecodeError::invalid(InvalidReason::InvalidIdentifier))?;
                if value == 0 {
                    return Err(DecodeError::invalid(InvalidReason::InvalidIdentifier));
                }
                count = Some(value);
            },
            _ => {
                unknown_fields = true;
                state.unknown()?;
            },
        }
    }
    Ok(OwnerSnapshot {
        object_identifier: object_identifier
            .ok_or_else(|| DecodeError::invalid(InvalidReason::MissingRequiredField))?,
        count: count.ok_or_else(|| DecodeError::invalid(InvalidReason::MissingRequiredField))?,
        raw: payload,
        unknown_fields,
    })
}

fn parse_reference(
    payload: &[u8],
    options: DecodeOptions,
    state: &mut ScanState,
) -> Result<(), DecodeError> {
    if payload.len() > options.max_message_bytes {
        return Err(DecodeError::limited(DecodeLimit::Bytes {
            observed: payload.len(),
            maximum: options.max_message_bytes,
        }));
    }
    state.work(payload.len(), options)?;
    let mut identifier = None;
    for result in fields(payload) {
        let field = result?;
        state.field(field, options, 2)?;
        match field.number {
            1 => {
                if identifier.is_some() {
                    return Err(DecodeError::invalid(InvalidReason::DuplicateField));
                }
                let value = varint(payload, field)?;
                if NonZeroU64::new(value).is_none() {
                    return Err(DecodeError::invalid(InvalidReason::InvalidIdentifier));
                }
                identifier = Some(value);
            },
            2 => {
                let _ = varint(payload, field)?;
            },
            3 => {
                let value = varint(payload, field)?;
                if value > 1 {
                    return Err(DecodeError::invalid(InvalidReason::UnsupportedField));
                }
            },
            _ => state.unknown()?,
        }
    }
    if identifier.is_none() {
        return Err(DecodeError::invalid(InvalidReason::MissingRequiredField));
    }
    Ok(())
}

fn validate_packed_varints(payload: &[u8]) -> Result<(), DecodeError> {
    let mut offset = 0usize;
    while offset < payload.len() {
        let (_, end) = decode_varint(payload, offset)?;
        offset = end;
    }
    Ok(())
}

fn validate_required_name(name: &str, options: DecodeOptions) -> Result<(), DecodeError> {
    if name.is_empty() {
        return Err(DecodeError::invalid(InvalidReason::InvalidName));
    }
    validate_optional_name(name, options)
}

fn validate_optional_name(name: &str, options: DecodeOptions) -> Result<(), DecodeError> {
    if name.len() > options.max_name_bytes {
        return Err(DecodeError::limited(DecodeLimit::NameBytes {
            observed: name.len(),
            maximum: options.max_name_bytes,
        }));
    }
    let path = Path::new(name);
    if path.is_absolute()
        || name.contains(['\0', '\\'])
        || path
            .components()
            .any(|component| !matches!(component, Component::Normal(_)))
    {
        return Err(DecodeError::invalid(InvalidReason::InvalidName));
    }
    Ok(())
}

fn map_decode(error: DecodeError) -> RewriteError {
    RewriteError::from_decode(error)
}

fn validate_addition(
    addition: DataInfoAddition<'_>,
    options: DecodeOptions,
) -> Result<(), RewriteError> {
    if NonZeroU64::new(addition.identifier).is_none() {
        return Err(RewriteError::invalid(InvalidReason::InvalidIdentifier));
    }
    if addition.digest.len() > options.max_digest_bytes {
        return Err(RewriteError::limited(DecodeLimit::DigestBytes {
            observed: addition.digest.len(),
            maximum: options.max_digest_bytes,
        }));
    }
    if addition.digest.len() != SHA1_DIGEST_BYTES {
        return Err(RewriteError::invalid(InvalidReason::InvalidDigest));
    }
    validate_required_name(addition.preferred_file_name, options).map_err(map_decode)?;
    if let Some(file_name) = addition.file_name {
        validate_optional_name(file_name, options).map_err(map_decode)?;
    }
    Ok(())
}

fn validate_nonzero(value: u64) -> Result<(), RewriteError> {
    NonZeroU64::new(value)
        .map(|_value| ())
        .ok_or_else(|| RewriteError::invalid(InvalidReason::InvalidIdentifier))
}

fn operation_component_matches(
    source: &[u8],
    selector: ComponentSelector<'_>,
    _options: DecodeOptions,
) -> Result<(usize, bool), RewriteError> {
    validate_nonzero(selector.identifier)?;
    let mut count = 0usize;
    let mut unknown = false;
    let mut offset = 0usize;
    while let Some(field) = next_field(source, offset).map_err(map_decode)? {
        offset = field.end;
        if field.number != ROOT_COMPONENT_FIELD && field.number != ROOT_VERSIONED_COMPONENT_FIELD {
            continue;
        }
        let component = component_facts(
            field.payload(source),
            field.number == ROOT_VERSIONED_COMPONENT_FIELD,
        )
        .map_err(map_decode)?;
        if component.identifier == selector.identifier
            && component.effective_locator() == selector.locator
        {
            count = count
                .checked_add(1)
                .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?;
            unknown |= component.has_unknown_fields();
            if component.is_versioned() {
                return Err(RewriteError::invalid(InvalidReason::VersionedComponent));
            }
        }
    }
    if count > 1 {
        return Err(RewriteError::invalid(InvalidReason::ComponentAmbiguous));
    }
    Ok((count, unknown))
}

fn component_facts<'source>(
    payload: &'source [u8],
    versioned: bool,
) -> Result<ComponentSnapshot<'source>, DecodeError> {
    let mut identifier = None;
    let mut preferred_locator = None;
    let mut locator = None;
    let mut unknown_fields = false;
    for result in fields(payload) {
        let field = result?;
        match field.number {
            COMPONENT_IDENTIFIER_FIELD => {
                if identifier.is_some() {
                    return Err(DecodeError::invalid(InvalidReason::DuplicateField));
                }
                let value = varint(payload, field)?;
                if NonZeroU64::new(value).is_none() {
                    return Err(DecodeError::invalid(InvalidReason::InvalidIdentifier));
                }
                identifier = Some(value);
            },
            COMPONENT_PREFERRED_LOCATOR_FIELD => {
                if preferred_locator.is_some() {
                    return Err(DecodeError::invalid(InvalidReason::DuplicateField));
                }
                preferred_locator = Some(utf8(payload, field)?);
            },
            COMPONENT_LOCATOR_FIELD => {
                if locator.is_some() {
                    return Err(DecodeError::invalid(InvalidReason::DuplicateField));
                }
                locator = Some(utf8(payload, field)?);
            },
            4 | 5 | 14 | 15 | 20 => validate_packed_varints(bytes(payload, field)?)?,
            6 | 7 | 11 | 13 | 18 => {
                let _ = bytes(payload, field)?;
            },
            10 | 17 | 19 => {
                let value = varint(payload, field)?;
                if value > 1 {
                    return Err(DecodeError::invalid(InvalidReason::UnsupportedField));
                }
            },
            12 | 16 | 21 => {
                let _ = varint(payload, field)?;
            },
            _ => unknown_fields = true,
        }
    }
    Ok(ComponentSnapshot {
        identifier: identifier
            .ok_or_else(|| DecodeError::invalid(InvalidReason::MissingRequiredField))?,
        preferred_locator: preferred_locator
            .ok_or_else(|| DecodeError::invalid(InvalidReason::MissingRequiredField))?,
        locator,
        versioned,
        unknown_fields,
        raw: payload,
    })
}

fn data_info_matches(
    source: &[u8],
    identifier: u64,
    options: DecodeOptions,
) -> Result<(usize, bool), RewriteError> {
    let mut count = 0usize;
    let mut unknown = false;
    let mut offset = 0usize;
    while let Some(field) = next_field(source, offset).map_err(map_decode)? {
        offset = field.end;
        if field.number != ROOT_DATA_INFO_FIELD {
            continue;
        }
        let mut state = ScanState::new(source, options).map_err(map_decode)?;
        let mut visitor = NoopVisitor;
        parse_data_info(
            source,
            field.payload(source),
            options,
            &mut state,
            &mut visitor,
        )
        .map_err(map_decode)?;
        let snapshot = data_info_facts(field.payload(source)).map_err(map_decode)?;
        if snapshot.identifier() == identifier {
            count = count
                .checked_add(1)
                .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?;
            unknown |= snapshot.has_unknown_fields();
        }
    }
    Ok((count, unknown))
}

fn data_info_facts<'source>(
    payload: &'source [u8],
) -> Result<DataInfoSnapshot<'source>, DecodeError> {
    let mut identifier = None;
    let mut digest = None;
    let mut preferred_file_name = None;
    let mut file_name = None;
    let mut materialized_length = None;
    let mut unknown_fields = false;
    for result in fields(payload) {
        let field = result?;
        match field.number {
            DATA_IDENTIFIER_FIELD => {
                if identifier.is_some() {
                    return Err(DecodeError::invalid(InvalidReason::DuplicateField));
                }
                identifier = Some(varint(payload, field)?);
            },
            DATA_DIGEST_FIELD => {
                if digest.is_some() {
                    return Err(DecodeError::invalid(InvalidReason::DuplicateField));
                }
                digest = Some(bytes(payload, field)?);
            },
            DATA_PREFERRED_NAME_FIELD => {
                if preferred_file_name.is_some() {
                    return Err(DecodeError::invalid(InvalidReason::DuplicateField));
                }
                preferred_file_name = Some(utf8(payload, field)?);
            },
            DATA_FILE_NAME_FIELD => {
                if file_name.is_some() {
                    return Err(DecodeError::invalid(InvalidReason::DuplicateField));
                }
                file_name = Some(utf8(payload, field)?);
            },
            DATA_MATERIALIZED_LENGTH_FIELD => {
                if materialized_length.is_some() {
                    return Err(DecodeError::invalid(InvalidReason::DuplicateField));
                }
                materialized_length = Some(varint(payload, field)?);
            },
            _ => unknown_fields = true,
        }
    }
    Ok(DataInfoSnapshot {
        identifier: identifier
            .ok_or_else(|| DecodeError::invalid(InvalidReason::MissingRequiredField))?,
        digest: digest.ok_or_else(|| DecodeError::invalid(InvalidReason::MissingRequiredField))?,
        preferred_file_name: preferred_file_name
            .ok_or_else(|| DecodeError::invalid(InvalidReason::MissingRequiredField))?,
        file_name,
        materialized_length,
        raw: payload,
        unknown_fields,
    })
}

#[derive(Debug, Clone, Copy, Default)]
struct OwnerMatchFacts {
    components: usize,
    component_unknown: bool,
    parents: usize,
    parent_unknown: bool,
    owners: usize,
    owner_count: usize,
    parent_owner_count: usize,
}

#[derive(Debug, Clone, Copy)]
struct DataReferenceFacts {
    data_identifier: u64,
    owners: usize,
    unknown_fields: bool,
}

fn data_reference_facts(payload: &[u8]) -> Result<DataReferenceFacts, DecodeError> {
    let mut data_identifier = None;
    let mut owners = 0usize;
    let mut unknown_fields = false;
    for result in fields(payload) {
        let field = result?;
        match field.number {
            DATA_IDENTIFIER_FIELD => {
                if data_identifier.is_some() {
                    return Err(DecodeError::invalid(InvalidReason::DuplicateField));
                }
                data_identifier = Some(varint(payload, field)?);
            },
            OWNER_COUNT_FIELD => {
                owners = owners
                    .checked_add(1)
                    .ok_or_else(|| DecodeError::invalid(InvalidReason::MalformedWire))?;
                let owner = owner_facts(field.payload(payload))?;
                // A parent rewrite would have to understand every selected
                // owner envelope.  Treat an opaque owner extension as part
                // of the selected parent so mutation fails closed instead
                // of copying an unmodelled child beside a new owner.
                unknown_fields |= owner.has_unknown_fields();
            },
            _ => unknown_fields = true,
        }
    }
    Ok(DataReferenceFacts {
        data_identifier: data_identifier
            .ok_or_else(|| DecodeError::invalid(InvalidReason::MissingRequiredField))?,
        owners,
        unknown_fields,
    })
}

fn owner_facts(payload: &[u8]) -> Result<OwnerSnapshot<'_>, DecodeError> {
    let mut object_identifier = None;
    let mut count = None;
    let mut unknown_fields = false;
    for result in fields(payload) {
        let field = result?;
        match field.number {
            OWNER_OBJECT_IDENTIFIER_FIELD => {
                if object_identifier.is_some() {
                    return Err(DecodeError::invalid(InvalidReason::DuplicateField));
                }
                object_identifier = Some(varint(payload, field)?);
            },
            OWNER_COUNT_FIELD => {
                if count.is_some() {
                    return Err(DecodeError::invalid(InvalidReason::DuplicateField));
                }
                let value = u32::try_from(varint(payload, field)?)
                    .map_err(|_error| DecodeError::invalid(InvalidReason::InvalidIdentifier))?;
                count = Some(value);
            },
            _ => unknown_fields = true,
        }
    }
    Ok(OwnerSnapshot {
        object_identifier: object_identifier
            .ok_or_else(|| DecodeError::invalid(InvalidReason::MissingRequiredField))?,
        count: count.ok_or_else(|| DecodeError::invalid(InvalidReason::MissingRequiredField))?,
        raw: payload,
        unknown_fields,
    })
}

fn owner_matches(
    source: &[u8],
    selector: ComponentSelector<'_>,
    data_identifier: u64,
    object_identifier: u64,
) -> Result<OwnerMatchFacts, RewriteError> {
    let mut facts = OwnerMatchFacts::default();
    let mut offset = 0usize;
    while let Some(field) = next_field(source, offset).map_err(map_decode)? {
        offset = field.end;
        if field.number != ROOT_COMPONENT_FIELD && field.number != ROOT_VERSIONED_COMPONENT_FIELD {
            continue;
        }
        let component = component_facts(
            field.payload(source),
            field.number == ROOT_VERSIONED_COMPONENT_FIELD,
        )
        .map_err(map_decode)?;
        if component.identifier != selector.identifier
            || component.effective_locator() != selector.locator
        {
            continue;
        }
        facts.components = facts
            .components
            .checked_add(1)
            .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?;
        facts.component_unknown |= component.unknown_fields;
        if component.versioned {
            return Err(RewriteError::invalid(InvalidReason::VersionedComponent));
        }
        for result in fields(field.payload(source)) {
            let child = result.map_err(map_decode)?;
            if child.number != COMPONENT_DATA_REFERENCE_FIELD {
                continue;
            }
            let parent =
                data_reference_facts(child.payload(field.payload(source))).map_err(map_decode)?;
            if parent.data_identifier != data_identifier {
                continue;
            }
            facts.parents = facts
                .parents
                .checked_add(1)
                .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?;
            facts.parent_unknown |= parent.unknown_fields;
            facts.parent_owner_count = facts
                .parent_owner_count
                .checked_add(parent.owners)
                .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?;
            for owner_result in fields(child.payload(field.payload(source))) {
                let owner_field = owner_result.map_err(map_decode)?;
                if owner_field.number != OWNER_COUNT_FIELD {
                    continue;
                }
                let owner = owner_facts(owner_field.payload(child.payload(field.payload(source))))
                    .map_err(map_decode)?;
                if owner.object_identifier == object_identifier {
                    facts.owners = facts
                        .owners
                        .checked_add(1)
                        .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?;
                    facts.owner_count = owner.count as usize;
                }
            }
        }
    }
    if facts.components > 1 {
        return Err(RewriteError::invalid(InvalidReason::ComponentAmbiguous));
    }
    if facts.parents > 1 {
        return Err(RewriteError::invalid(InvalidReason::DuplicateDataInfo));
    }
    Ok(facts)
}

fn count_data_owners(source: &[u8], data_identifier: u64) -> Result<usize, RewriteError> {
    let mut count = 0usize;
    let mut offset = 0usize;
    while let Some(field) = next_field(source, offset).map_err(map_decode)? {
        offset = field.end;
        if field.number != ROOT_COMPONENT_FIELD && field.number != ROOT_VERSIONED_COMPONENT_FIELD {
            continue;
        }
        let component_payload = field.payload(source);
        for result in fields(component_payload) {
            let child = result.map_err(map_decode)?;
            if child.number != COMPONENT_DATA_REFERENCE_FIELD {
                continue;
            }
            let reference =
                data_reference_facts(child.payload(component_payload)).map_err(map_decode)?;
            if reference.data_identifier == data_identifier {
                count = count
                    .checked_add(reference.owners)
                    .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?;
            }
        }
    }
    Ok(count)
}

trait Sink {
    fn emit(&mut self, bytes: &[u8], maximum: usize) -> Result<(), RewriteError>;
}

struct CountSink {
    len: usize,
}

impl Sink for CountSink {
    fn emit(&mut self, bytes: &[u8], maximum: usize) -> Result<(), RewriteError> {
        self.len = self
            .len
            .checked_add(bytes.len())
            .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?;
        let _ = maximum;
        Ok(())
    }
}

struct VecSink<'a> {
    output: &'a mut Vec<u8>,
}

impl Sink for VecSink<'_> {
    fn emit(&mut self, bytes: &[u8], maximum: usize) -> Result<(), RewriteError> {
        let next = self
            .output
            .len()
            .checked_add(bytes.len())
            .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?;
        if next > maximum {
            return Err(RewriteError::limited(DecodeLimit::OutputBytes {
                observed: next,
                maximum,
            }));
        }
        self.output.extend_from_slice(bytes);
        Ok(())
    }
}

fn emit_varint<S: Sink>(sink: &mut S, mut value: u64, maximum: usize) -> Result<(), RewriteError> {
    let mut bytes = [0u8; 10];
    let mut length = 0usize;
    loop {
        let mut byte = (value & 0x7f) as u8;
        value >>= 7;
        if value != 0 {
            byte |= 0x80;
        }
        bytes[length] = byte;
        length += 1;
        if value == 0 {
            return sink.emit(&bytes[..length], maximum);
        }
        if length == bytes.len() {
            return Err(RewriteError::invalid(InvalidReason::MalformedWire));
        }
    }
}

fn emit_field_header<S: Sink>(
    sink: &mut S,
    number: u32,
    wire_type: u8,
    payload_len: usize,
    maximum: usize,
) -> Result<(), RewriteError> {
    let key = (u64::from(number) << 3) | u64::from(wire_type);
    emit_varint(sink, key, maximum)?;
    if wire_type == 2 {
        let length = u64::try_from(payload_len)
            .map_err(|_error| RewriteError::invalid(InvalidReason::MalformedWire))?;
        emit_varint(sink, length, maximum)?;
    }
    Ok(())
}

fn emit_varint_field<S: Sink>(
    sink: &mut S,
    number: u32,
    value: u64,
    maximum: usize,
) -> Result<(), RewriteError> {
    emit_field_header(sink, number, 0, 0, maximum)?;
    emit_varint(sink, value, maximum)
}

fn emit_bytes_field<S: Sink>(
    sink: &mut S,
    number: u32,
    payload: &[u8],
    maximum: usize,
) -> Result<(), RewriteError> {
    emit_field_header(sink, number, 2, payload.len(), maximum)?;
    sink.emit(payload, maximum)
}

fn emit_owner<S: Sink>(
    sink: &mut S,
    object_identifier: u64,
    count: u32,
    maximum: usize,
) -> Result<(), RewriteError> {
    let object_field_len = encoded_len(u64::from(OWNER_OBJECT_IDENTIFIER_FIELD) << 3)
        .checked_add(encoded_len(object_identifier))
        .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?;
    let count_field_len = encoded_len(u64::from(OWNER_COUNT_FIELD) << 3)
        .checked_add(encoded_len(u64::from(count)))
        .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?;
    let nested_len = object_field_len
        .checked_add(count_field_len)
        .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?;
    emit_field_header(sink, OWNER_COUNT_FIELD, 2, nested_len, maximum)?;
    emit_varint_field(
        sink,
        OWNER_OBJECT_IDENTIFIER_FIELD,
        object_identifier,
        maximum,
    )?;
    emit_varint_field(sink, OWNER_COUNT_FIELD, u64::from(count), maximum)
}

fn emit_data_info<S: Sink>(
    sink: &mut S,
    addition: DataInfoAddition<'_>,
    maximum: usize,
) -> Result<(), RewriteError> {
    let mut nested = CountSink { len: 0 };
    emit_varint_field(
        &mut nested,
        DATA_IDENTIFIER_FIELD,
        addition.identifier,
        usize::MAX,
    )?;
    emit_bytes_field(&mut nested, DATA_DIGEST_FIELD, addition.digest, usize::MAX)?;
    emit_bytes_field(
        &mut nested,
        DATA_PREFERRED_NAME_FIELD,
        addition.preferred_file_name.as_bytes(),
        usize::MAX,
    )?;
    if let Some(file_name) = addition.file_name {
        emit_bytes_field(
            &mut nested,
            DATA_FILE_NAME_FIELD,
            file_name.as_bytes(),
            usize::MAX,
        )?;
    }
    if let Some(length) = addition.materialized_length {
        emit_varint_field(
            &mut nested,
            DATA_MATERIALIZED_LENGTH_FIELD,
            length,
            usize::MAX,
        )?;
    }
    emit_field_header(sink, ROOT_DATA_INFO_FIELD, 2, nested.len, maximum)?;
    emit_varint_field(sink, DATA_IDENTIFIER_FIELD, addition.identifier, maximum)?;
    emit_bytes_field(sink, DATA_DIGEST_FIELD, addition.digest, maximum)?;
    emit_bytes_field(
        sink,
        DATA_PREFERRED_NAME_FIELD,
        addition.preferred_file_name.as_bytes(),
        maximum,
    )?;
    if let Some(file_name) = addition.file_name {
        emit_bytes_field(sink, DATA_FILE_NAME_FIELD, file_name.as_bytes(), maximum)?;
    }
    if let Some(length) = addition.materialized_length {
        emit_varint_field(sink, DATA_MATERIALIZED_LENGTH_FIELD, length, maximum)?;
    }
    Ok(())
}

fn same_component(left: ComponentSelector<'_>, right: ComponentSelector<'_>) -> bool {
    left.identifier == right.identifier && left.locator == right.locator
}

fn data_addition_duplicate(
    additions: &[DataInfoAddition<'_>],
    index: usize,
    identifier: u64,
) -> bool {
    additions[..index]
        .iter()
        .any(|addition| addition.identifier == identifier)
}

fn data_removal_duplicate(removals: &[DataInfoRemoval], index: usize, identifier: u64) -> bool {
    removals[..index]
        .iter()
        .any(|removal| removal.identifier == identifier)
}

fn owner_addition_duplicate(
    additions: &[DataReferenceOwnerAddition<'_>],
    index: usize,
    addition: DataReferenceOwnerAddition<'_>,
) -> bool {
    additions[..index].iter().any(|candidate| {
        same_component(candidate.component, addition.component)
            && candidate.data_identifier == addition.data_identifier
            && candidate.object_identifier == addition.object_identifier
    })
}

fn owner_removal_duplicate(
    removals: &[DataReferenceOwnerRemoval<'_>],
    index: usize,
    removal: DataReferenceOwnerRemoval<'_>,
) -> bool {
    removals[..index].iter().any(|candidate| {
        same_component(candidate.component, removal.component)
            && candidate.data_identifier == removal.data_identifier
            && candidate.object_identifier == removal.object_identifier
    })
}

fn owner_addition_conflicts_with_removal(
    addition: DataReferenceOwnerAddition<'_>,
    removals: &[DataReferenceOwnerRemoval<'_>],
) -> bool {
    removals.iter().any(|removal| {
        same_component(addition.component, removal.component)
            && addition.data_identifier == removal.data_identifier
            && addition.object_identifier == removal.object_identifier
    })
}

fn owner_removal_conflicts_with_addition(
    removal: DataReferenceOwnerRemoval<'_>,
    additions: &[DataReferenceOwnerAddition<'_>],
) -> bool {
    additions.iter().any(|addition| {
        same_component(addition.component, removal.component)
            && addition.data_identifier == removal.data_identifier
            && addition.object_identifier == removal.object_identifier
    })
}

fn owner_update_duplicate(
    updates: &[DataReferenceOwnerCountUpdate<'_>],
    index: usize,
    update: DataReferenceOwnerCountUpdate<'_>,
) -> bool {
    updates[..index].iter().any(|candidate| {
        same_component(candidate.component, update.component)
            && candidate.data_identifier == update.data_identifier
            && candidate.object_identifier == update.object_identifier
    })
}

fn owner_update_conflicts_with_addition(
    update: DataReferenceOwnerCountUpdate<'_>,
    additions: &[DataReferenceOwnerAddition<'_>],
) -> bool {
    additions.iter().any(|addition| {
        same_component(addition.component, update.component)
            && addition.data_identifier == update.data_identifier
            && addition.object_identifier == update.object_identifier
    })
}

fn owner_update_conflicts_with_removal(
    update: DataReferenceOwnerCountUpdate<'_>,
    removals: &[DataReferenceOwnerRemoval<'_>],
) -> bool {
    removals.iter().any(|removal| {
        same_component(removal.component, update.component)
            && removal.data_identifier == update.data_identifier
            && removal.object_identifier == update.object_identifier
    })
}

fn owner_addition_conflicts_with_update(
    addition: DataReferenceOwnerAddition<'_>,
    updates: &[DataReferenceOwnerCountUpdate<'_>],
) -> bool {
    updates.iter().any(|update| {
        same_component(addition.component, update.component)
            && addition.data_identifier == update.data_identifier
            && addition.object_identifier == update.object_identifier
    })
}

fn owner_removal_conflicts_with_update(
    removal: DataReferenceOwnerRemoval<'_>,
    updates: &[DataReferenceOwnerCountUpdate<'_>],
) -> bool {
    updates.iter().any(|update| {
        same_component(removal.component, update.component)
            && removal.data_identifier == update.data_identifier
            && removal.object_identifier == update.object_identifier
    })
}

fn data_present_after(
    source_count: usize,
    identifier: u64,
    additions: &[DataInfoAddition<'_>],
    removals: &[DataInfoRemoval],
) -> bool {
    if removals
        .iter()
        .any(|removal| removal.identifier == identifier)
    {
        return false;
    }
    source_count == 1
        || additions
            .iter()
            .any(|addition| addition.identifier == identifier)
}

fn validate_batch(
    source: &[u8],
    options: DecodeOptions,
    batch: MediaRewriteBatch<'_>,
    report: DecodeReport,
) -> Result<(), RewriteError> {
    let operation_count = batch
        .data_additions
        .len()
        .checked_add(batch.data_removals.len())
        .and_then(|value| value.checked_add(batch.owner_additions.len()))
        .and_then(|value| value.checked_add(batch.owner_removals.len()))
        .and_then(|value| value.checked_add(batch.owner_updates.len()))
        .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?;
    if operation_count > options.max_fields {
        return Err(RewriteError::limited(DecodeLimit::Fields {
            observed: operation_count,
            maximum: options.max_fields,
        }));
    }

    for (index, addition) in batch.data_additions.iter().copied().enumerate() {
        validate_addition(addition, options)?;
        if data_addition_duplicate(batch.data_additions, index, addition.identifier)
            || batch
                .data_removals
                .iter()
                .any(|removal| removal.identifier == addition.identifier)
        {
            return Err(RewriteError::invalid(InvalidReason::DuplicateOperation));
        }
        let (count, _unknown) = data_info_matches(source, addition.identifier, options)?;
        if count != 0 {
            return Err(RewriteError::invalid(InvalidReason::ExistingDataCollision));
        }
    }

    for (index, removal) in batch.data_removals.iter().copied().enumerate() {
        validate_nonzero(removal.identifier)?;
        if data_removal_duplicate(batch.data_removals, index, removal.identifier)
            || batch
                .data_additions
                .iter()
                .any(|addition| addition.identifier == removal.identifier)
        {
            return Err(RewriteError::invalid(InvalidReason::DuplicateOperation));
        }
        if report.data_metadata_map_present {
            return Err(RewriteError::invalid(
                InvalidReason::DataInfoMetadataMapDependency,
            ));
        }
        let (count, unknown) = data_info_matches(source, removal.identifier, options)?;
        if count == 0 {
            return Err(RewriteError::invalid(InvalidReason::DataInfoNotFound));
        }
        if count != 1 {
            return Err(RewriteError::invalid(InvalidReason::DuplicateDataInfo));
        }
        if unknown {
            return Err(RewriteError::invalid(InvalidReason::UnknownSelectedRecord));
        }
        if count_data_owners(source, removal.identifier)? != 0 {
            return Err(RewriteError::invalid(InvalidReason::DataInfoReferenced));
        }
    }

    for (index, addition) in batch.owner_additions.iter().copied().enumerate() {
        validate_nonzero(addition.component.identifier)?;
        validate_nonzero(addition.data_identifier)?;
        validate_nonzero(addition.object_identifier)?;
        if addition.count == 0 {
            return Err(RewriteError::invalid(InvalidReason::InvalidIdentifier));
        }
        if owner_addition_duplicate(batch.owner_additions, index, addition)
            || owner_addition_conflicts_with_removal(addition, batch.owner_removals)
            || owner_addition_conflicts_with_update(addition, batch.owner_updates)
        {
            return Err(RewriteError::invalid(InvalidReason::DuplicateOperation));
        }
        let (components, unknown) =
            operation_component_matches(source, addition.component, options)?;
        if components == 0 {
            return Err(RewriteError::invalid(InvalidReason::ComponentNotFound));
        }
        if unknown {
            return Err(RewriteError::invalid(InvalidReason::UnknownSelectedRecord));
        }
        let owner = owner_matches(
            source,
            addition.component,
            addition.data_identifier,
            addition.object_identifier,
        )?;
        if owner.owners != 0 {
            return Err(RewriteError::invalid(InvalidReason::ExistingOwnerCollision));
        }
        if owner.parents > 1 {
            return Err(RewriteError::invalid(InvalidReason::DuplicateDataInfo));
        }
        if owner.parents == 1 && owner.parent_unknown {
            return Err(RewriteError::invalid(InvalidReason::UnknownSelectedRecord));
        }
        let (data_count, _unknown) = data_info_matches(source, addition.data_identifier, options)?;
        if !data_present_after(
            data_count,
            addition.data_identifier,
            batch.data_additions,
            batch.data_removals,
        ) {
            return Err(RewriteError::invalid(InvalidReason::DataInfoNotFound));
        }
    }

    for (index, removal) in batch.owner_removals.iter().copied().enumerate() {
        validate_nonzero(removal.component.identifier)?;
        validate_nonzero(removal.data_identifier)?;
        validate_nonzero(removal.object_identifier)?;
        if removal.expected_count == 0 {
            return Err(RewriteError::invalid(InvalidReason::InvalidIdentifier));
        }
        if owner_removal_duplicate(batch.owner_removals, index, removal)
            || owner_removal_conflicts_with_addition(removal, batch.owner_additions)
            || owner_removal_conflicts_with_update(removal, batch.owner_updates)
        {
            return Err(RewriteError::invalid(InvalidReason::DuplicateOperation));
        }
        let facts = owner_matches(
            source,
            removal.component,
            removal.data_identifier,
            removal.object_identifier,
        )?;
        if facts.components == 0 {
            return Err(RewriteError::invalid(InvalidReason::ComponentNotFound));
        }
        if facts.component_unknown || facts.parent_unknown {
            return Err(RewriteError::invalid(InvalidReason::UnknownSelectedRecord));
        }
        if facts.parents == 0 || facts.owners == 0 {
            return Err(RewriteError::invalid(InvalidReason::OwnerNotFound));
        }
        if facts.owners != 1 || facts.owner_count != removal.expected_count as usize {
            return Err(RewriteError::invalid(InvalidReason::ExistingOwnerCollision));
        }
    }

    for (index, update) in batch.owner_updates.iter().copied().enumerate() {
        validate_nonzero(update.component.identifier)?;
        validate_nonzero(update.data_identifier)?;
        validate_nonzero(update.object_identifier)?;
        if update.expected_count == 0 || update.new_count == 0 {
            return Err(RewriteError::invalid(InvalidReason::InvalidIdentifier));
        }
        if owner_update_duplicate(batch.owner_updates, index, update)
            || owner_update_conflicts_with_addition(update, batch.owner_additions)
            || owner_update_conflicts_with_removal(update, batch.owner_removals)
        {
            return Err(RewriteError::invalid(InvalidReason::ConflictingOperation));
        }
        let (components, unknown) = operation_component_matches(source, update.component, options)?;
        if components == 0 {
            return Err(RewriteError::invalid(InvalidReason::ComponentNotFound));
        }
        if unknown {
            return Err(RewriteError::invalid(InvalidReason::UnknownSelectedRecord));
        }
        let facts = owner_matches(
            source,
            update.component,
            update.data_identifier,
            update.object_identifier,
        )?;
        if facts.components == 0 {
            return Err(RewriteError::invalid(InvalidReason::ComponentNotFound));
        }
        if facts.component_unknown || facts.parent_unknown {
            return Err(RewriteError::invalid(InvalidReason::UnknownSelectedRecord));
        }
        if facts.parents == 0 || facts.owners == 0 {
            return Err(RewriteError::invalid(InvalidReason::OwnerNotFound));
        }
        if facts.owners != 1 || facts.owner_count != update.expected_count as usize {
            return Err(RewriteError::invalid(InvalidReason::ExistingOwnerCollision));
        }
        let (data_count, _unknown) = data_info_matches(source, update.data_identifier, options)?;
        if !data_present_after(
            data_count,
            update.data_identifier,
            batch.data_additions,
            batch.data_removals,
        ) {
            return Err(RewriteError::invalid(InvalidReason::DataInfoNotFound));
        }
    }
    Ok(())
}

fn owner_additions_for<'a>(
    additions: &'a [DataReferenceOwnerAddition<'a>],
    component: ComponentSnapshot<'a>,
    data_identifier: u64,
) -> impl Iterator<Item = DataReferenceOwnerAddition<'a>> + 'a {
    additions.iter().copied().filter(move |addition| {
        addition.component.identifier == component.identifier
            && addition.component.locator == component.effective_locator()
            && addition.data_identifier == data_identifier
    })
}

fn owner_removals_for<'a>(
    removals: &'a [DataReferenceOwnerRemoval<'a>],
    component: ComponentSnapshot<'a>,
    data_identifier: u64,
) -> impl Iterator<Item = DataReferenceOwnerRemoval<'a>> + 'a {
    removals.iter().copied().filter(move |removal| {
        removal.component.identifier == component.identifier
            && removal.component.locator == component.effective_locator()
            && removal.data_identifier == data_identifier
    })
}

fn owner_updates_for<'a>(
    updates: &'a [DataReferenceOwnerCountUpdate<'a>],
    component: ComponentSnapshot<'a>,
    data_identifier: u64,
) -> impl Iterator<Item = DataReferenceOwnerCountUpdate<'a>> + 'a {
    updates.iter().copied().filter(move |update| {
        update.component.identifier == component.identifier
            && update.component.locator == component.effective_locator()
            && update.data_identifier == data_identifier
    })
}

fn component_rewrite_needed(
    component: ComponentSnapshot<'_>,
    batch: MediaRewriteBatch<'_>,
) -> bool {
    batch.owner_additions.iter().any(|addition| {
        addition.component.identifier == component.identifier
            && addition.component.locator == component.effective_locator()
    }) || batch.owner_removals.iter().any(|removal| {
        removal.component.identifier == component.identifier
            && removal.component.locator == component.effective_locator()
    }) || batch.owner_updates.iter().any(|update| {
        update.component.identifier == component.identifier
            && update.component.locator == component.effective_locator()
    })
}

fn original_parent_exists(payload: &[u8], data_identifier: u64) -> Result<bool, RewriteError> {
    let mut found = false;
    for result in fields(payload) {
        let field = result.map_err(map_decode)?;
        if field.number != COMPONENT_DATA_REFERENCE_FIELD {
            continue;
        }
        let facts = data_reference_facts(field.payload(payload)).map_err(map_decode)?;
        if facts.data_identifier == data_identifier {
            if found {
                return Err(RewriteError::invalid(InvalidReason::DuplicateDataInfo));
            }
            found = true;
        }
    }
    Ok(found)
}

/// Rewrite only the selected owner count while retaining the source ordering
/// and byte representation of every other known field.  Selected owner
/// records with unknown fields are rejected during validation, so this helper
/// never has to guess how an extension interacts with the count transition.
fn emit_rewritten_owner<S: Sink>(
    sink: &mut S,
    payload: &[u8],
    update: DataReferenceOwnerCountUpdate<'_>,
    maximum: usize,
) -> Result<(), RewriteError> {
    let mut inner = CountSink { len: 0 };
    for result in fields(payload) {
        let field = result.map_err(map_decode)?;
        if field.number == OWNER_COUNT_FIELD {
            emit_varint_field(
                &mut inner,
                OWNER_COUNT_FIELD,
                u64::from(update.new_count),
                usize::MAX,
            )?;
        } else {
            inner.emit(field_bytes(payload, field), usize::MAX)?;
        }
    }
    emit_field_header(sink, OWNER_COUNT_FIELD, 2, inner.len, maximum)?;
    for result in fields(payload) {
        let field = result.map_err(map_decode)?;
        if field.number == OWNER_COUNT_FIELD {
            emit_varint_field(
                sink,
                OWNER_COUNT_FIELD,
                u64::from(update.new_count),
                maximum,
            )?;
        } else {
            sink.emit(field_bytes(payload, field), maximum)?;
        }
    }
    Ok(())
}

fn emit_rewritten_data_reference<S: Sink>(
    sink: &mut S,
    payload: &[u8],
    component: ComponentSnapshot<'_>,
    batch: MediaRewriteBatch<'_>,
    maximum: usize,
) -> Result<(), RewriteError> {
    let facts = data_reference_facts(payload).map_err(map_decode)?;
    let has_additions =
        owner_additions_for(batch.owner_additions, component, facts.data_identifier)
            .next()
            .is_some();
    let mut inner = CountSink { len: 0 };
    let mut retained_owners = 0usize;
    for result in fields(payload) {
        let field = result.map_err(map_decode)?;
        if field.number == OWNER_COUNT_FIELD {
            let owner = owner_facts(field.payload(payload)).map_err(map_decode)?;
            let remove = owner_removals_for(batch.owner_removals, component, facts.data_identifier)
                .any(|removal| {
                    removal.object_identifier == owner.object_identifier
                        && removal.expected_count == owner.count
                });
            if remove {
                continue;
            }
            retained_owners = retained_owners
                .checked_add(1)
                .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?;
            if let Some(update) =
                owner_updates_for(batch.owner_updates, component, facts.data_identifier).find(
                    |update| {
                        update.object_identifier == owner.object_identifier
                            && update.expected_count == owner.count
                    },
                )
            {
                emit_rewritten_owner(&mut inner, field.payload(payload), update, usize::MAX)?;
                continue;
            }
        }
        inner.emit(field_bytes(payload, field), usize::MAX)?;
    }
    for addition in owner_additions_for(batch.owner_additions, component, facts.data_identifier) {
        emit_owner(
            &mut inner,
            addition.object_identifier,
            addition.count,
            usize::MAX,
        )?;
    }
    if retained_owners == 0 && !has_additions {
        return Ok(());
    }
    emit_field_header(sink, COMPONENT_DATA_REFERENCE_FIELD, 2, inner.len, maximum)?;
    for result in fields(payload) {
        let field = result.map_err(map_decode)?;
        if field.number == OWNER_COUNT_FIELD {
            let owner = owner_facts(field.payload(payload)).map_err(map_decode)?;
            let remove = owner_removals_for(batch.owner_removals, component, facts.data_identifier)
                .any(|removal| {
                    removal.object_identifier == owner.object_identifier
                        && removal.expected_count == owner.count
                });
            if remove {
                continue;
            }
            if let Some(update) =
                owner_updates_for(batch.owner_updates, component, facts.data_identifier).find(
                    |update| {
                        update.object_identifier == owner.object_identifier
                            && update.expected_count == owner.count
                    },
                )
            {
                emit_rewritten_owner(sink, field.payload(payload), update, maximum)?;
                continue;
            }
        }
        sink.emit(field_bytes(payload, field), maximum)?;
    }
    for addition in owner_additions_for(batch.owner_additions, component, facts.data_identifier) {
        emit_owner(sink, addition.object_identifier, addition.count, maximum)?;
    }
    Ok(())
}

fn emit_missing_data_reference<S: Sink>(
    sink: &mut S,
    component: ComponentSnapshot<'_>,
    data_identifier: u64,
    additions: &[DataReferenceOwnerAddition<'_>],
    maximum: usize,
) -> Result<(), RewriteError> {
    let mut inner = CountSink { len: 0 };
    emit_varint_field(
        &mut inner,
        DATA_IDENTIFIER_FIELD,
        data_identifier,
        usize::MAX,
    )?;
    for addition in owner_additions_for(additions, component, data_identifier) {
        emit_owner(
            &mut inner,
            addition.object_identifier,
            addition.count,
            usize::MAX,
        )?;
    }
    emit_field_header(sink, COMPONENT_DATA_REFERENCE_FIELD, 2, inner.len, maximum)?;
    emit_varint_field(sink, DATA_IDENTIFIER_FIELD, data_identifier, maximum)?;
    for addition in owner_additions_for(additions, component, data_identifier) {
        emit_owner(sink, addition.object_identifier, addition.count, maximum)?;
    }
    Ok(())
}

fn first_missing_parent_addition(
    additions: &[DataReferenceOwnerAddition<'_>],
    index: usize,
    component: ComponentSnapshot<'_>,
    data_identifier: u64,
) -> bool {
    !additions[..index].iter().any(|addition| {
        addition.component.identifier == component.identifier
            && addition.component.locator == component.effective_locator()
            && addition.data_identifier == data_identifier
    })
}

fn rewrite_component<S: Sink>(
    payload: &[u8],
    component: ComponentSnapshot<'_>,
    batch: MediaRewriteBatch<'_>,
    sink: &mut S,
    maximum: usize,
) -> Result<(), RewriteError> {
    for result in fields(payload) {
        let field = result.map_err(map_decode)?;
        if field.number == COMPONENT_DATA_REFERENCE_FIELD
            && component_rewrite_needed(component, batch)
        {
            let reference = field.payload(payload);
            let facts = data_reference_facts(reference).map_err(map_decode)?;
            let has_ops =
                owner_additions_for(batch.owner_additions, component, facts.data_identifier)
                    .next()
                    .is_some()
                    || owner_removals_for(batch.owner_removals, component, facts.data_identifier)
                        .next()
                        .is_some()
                    || owner_updates_for(batch.owner_updates, component, facts.data_identifier)
                        .next()
                        .is_some();
            if has_ops {
                emit_rewritten_data_reference(sink, reference, component, batch, maximum)?;
            } else {
                sink.emit(field_bytes(payload, field), maximum)?;
            }
        } else {
            sink.emit(field_bytes(payload, field), maximum)?;
        }
    }
    for (index, addition) in batch.owner_additions.iter().copied().enumerate() {
        if addition.component.identifier != component.identifier
            || addition.component.locator != component.effective_locator()
            || original_parent_exists(payload, addition.data_identifier)?
            || !first_missing_parent_addition(
                batch.owner_additions,
                index,
                component,
                addition.data_identifier,
            )
        {
            continue;
        }
        emit_missing_data_reference(
            sink,
            component,
            addition.data_identifier,
            batch.owner_additions,
            maximum,
        )?;
    }
    Ok(())
}

fn field_bytes(source: &[u8], field: WireField) -> &[u8] {
    &source[field.start..field.end]
}

fn rewrite_source<S: Sink>(
    source: &[u8],
    batch: MediaRewriteBatch<'_>,
    maximum: usize,
    sink: &mut S,
) -> Result<(), RewriteError> {
    let mut offset = 0usize;
    while let Some(field) = next_field(source, offset).map_err(map_decode)? {
        offset = field.end;
        match field.number {
            ROOT_COMPONENT_FIELD => {
                let payload = field.payload(source);
                let component = component_facts(payload, false).map_err(map_decode)?;
                if component_rewrite_needed(component, batch) {
                    let mut inner = CountSink { len: 0 };
                    rewrite_component(payload, component, batch, &mut inner, usize::MAX)?;
                    emit_field_header(sink, ROOT_COMPONENT_FIELD, 2, inner.len, maximum)?;
                    rewrite_component(payload, component, batch, sink, maximum)?;
                } else {
                    sink.emit(field_bytes(source, field), maximum)?;
                }
            },
            ROOT_DATA_INFO_FIELD => {
                let snapshot = data_info_facts(field.payload(source)).map_err(map_decode)?;
                if batch
                    .data_removals
                    .iter()
                    .any(|removal| removal.identifier == snapshot.identifier)
                {
                    continue;
                }
                sink.emit(field_bytes(source, field), maximum)?;
            },
            _ => sink.emit(field_bytes(source, field), maximum)?,
        }
    }
    for addition in batch.data_additions.iter().copied() {
        emit_data_info(sink, addition, maximum)?;
    }
    Ok(())
}

fn postcondition_data_count(
    source: &[u8],
    identifier: u64,
    options: DecodeOptions,
) -> Result<usize, RewriteError> {
    data_info_matches(source, identifier, options).map(|(count, _unknown)| count)
}

fn verify_postconditions(
    candidate: &[u8],
    options: DecodeOptions,
    batch: MediaRewriteBatch<'_>,
    work_limit: usize,
) -> Result<(), RewriteError> {
    let candidate_options = options
        .with_max_message_bytes(candidate.len().max(options.max_message_bytes))
        .with_max_work_bytes(work_limit);
    let _report =
        inspect_package_metadata_media(candidate, candidate_options).map_err(map_decode)?;
    for addition in batch.data_additions.iter().copied() {
        if postcondition_data_count(candidate, addition.identifier, candidate_options)? != 1 {
            return Err(RewriteError::invalid(InvalidReason::Verification));
        }
    }
    for removal in batch.data_removals.iter().copied() {
        if postcondition_data_count(candidate, removal.identifier, candidate_options)? != 0 {
            return Err(RewriteError::invalid(InvalidReason::Verification));
        }
    }
    for addition in batch.owner_additions.iter().copied() {
        let facts = owner_matches(
            candidate,
            addition.component,
            addition.data_identifier,
            addition.object_identifier,
        )?;
        if facts.components != 1 || facts.parents != 1 || facts.owners != 1 {
            return Err(RewriteError::invalid(InvalidReason::Verification));
        }
    }
    for removal in batch.owner_removals.iter().copied() {
        let facts = owner_matches(
            candidate,
            removal.component,
            removal.data_identifier,
            removal.object_identifier,
        )?;
        if facts.owners != 0 {
            return Err(RewriteError::invalid(InvalidReason::Verification));
        }
    }
    for update in batch.owner_updates.iter().copied() {
        let facts = owner_matches(
            candidate,
            update.component,
            update.data_identifier,
            update.object_identifier,
        )?;
        if facts.components != 1
            || facts.parents != 1
            || facts.owners != 1
            || facts.owner_count != update.new_count as usize
        {
            return Err(RewriteError::invalid(InvalidReason::Verification));
        }
    }
    Ok(())
}

/// Output-free semantic validation and exact sizing for one media metadata
/// transaction.
#[derive(Debug)]
pub struct PreparedPackageMetadataMediaRewrite<'source> {
    source: &'source [u8],
    batch: MediaRewriteBatch<'source>,
    options: DecodeOptions,
    output_size: usize,
    source_report: DecodeReport,
    requirements: RewriteExecutionRequirements,
}

impl PreparedPackageMetadataMediaRewrite<'_> {
    #[must_use]
    pub const fn source_report(&self) -> DecodeReport {
        self.source_report
    }

    #[must_use]
    pub const fn execution_requirements(&self) -> RewriteExecutionRequirements {
        self.requirements
    }

    #[must_use]
    pub const fn output_size(&self) -> usize {
        self.output_size
    }

    /// Allocate exactly one candidate after independently checking execution
    /// limits, then verify the candidate before returning it.
    pub fn execute(
        self,
        limits: RewriteExecutionLimits,
    ) -> Result<MediaRewriteOutput, RewriteError> {
        check_execution_limits(self.requirements, limits)?;
        let mut output = Vec::new();
        output
            .try_reserve_exact(self.output_size)
            .map_err(|_error| RewriteError::allocation(self.output_size))?;
        let mut sink = VecSink {
            output: &mut output,
        };
        rewrite_source(
            self.source,
            self.batch,
            self.options.max_output_bytes,
            &mut sink,
        )?;
        if output.len() != self.output_size {
            return Err(RewriteError::invalid(InvalidReason::Verification));
        }
        verify_postconditions(
            &output,
            self.options,
            self.batch,
            self.requirements.work_bytes,
        )?;
        Ok(MediaRewriteOutput {
            bytes: output,
            report: RewriteReport {
                input_bytes: self.source.len(),
                output_bytes: self.output_size,
                fields: self.requirements.fields,
                work_bytes: self.requirements.work_bytes,
                max_depth: self.source_report.max_depth,
                components: self.source_report.components,
                data_records: self.source_report.data_records,
                owners: self.source_report.owners,
                data_additions: self.batch.data_additions.len(),
                data_removals: self.batch.data_removals.len(),
                owner_additions: self.batch.owner_additions.len(),
                owner_removals: self.batch.owner_removals.len(),
                owner_updates: self.batch.owner_updates.len(),
                allocations: 1,
                retained_bytes: self.output_size,
                scratch_bytes: 0,
            },
        })
    }
}

fn check_execution_limits(
    required: RewriteExecutionRequirements,
    limits: RewriteExecutionLimits,
) -> Result<(), RewriteError> {
    if required.output_bytes > limits.max_output_bytes {
        return Err(RewriteError::limited(DecodeLimit::OutputBytes {
            observed: required.output_bytes,
            maximum: limits.max_output_bytes,
        }));
    }
    if required.fields > limits.max_fields {
        return Err(RewriteError::limited(DecodeLimit::Fields {
            observed: required.fields,
            maximum: limits.max_fields,
        }));
    }
    if required.work_bytes > limits.max_work_bytes {
        return Err(RewriteError::limited(DecodeLimit::Work {
            observed: required.work_bytes,
            maximum: limits.max_work_bytes,
        }));
    }
    if required.components > limits.max_components {
        return Err(RewriteError::limited(DecodeLimit::Components {
            observed: required.components,
            maximum: limits.max_components,
        }));
    }
    if required.data_records > limits.max_data_records {
        return Err(RewriteError::limited(DecodeLimit::DataRecords {
            observed: required.data_records,
            maximum: limits.max_data_records,
        }));
    }
    if required.owners > limits.max_owners {
        return Err(RewriteError::limited(DecodeLimit::Owners {
            observed: required.owners,
            maximum: limits.max_owners,
        }));
    }
    if required.allocations > limits.max_allocations
        || required.retained_bytes > limits.max_retained_bytes
        || required.scratch_bytes > limits.max_scratch_bytes
    {
        return Err(RewriteError::invalid(InvalidReason::Verification));
    }
    Ok(())
}

/// Prepare one atomic PackageMetadata media rewrite.
pub fn prepare_package_metadata_media_rewrite<'source>(
    source: &'source [u8],
    batch: MediaRewriteBatch<'source>,
    options: DecodeOptions,
) -> Result<PreparedPackageMetadataMediaRewrite<'source>, RewriteError> {
    let source_report = inspect_package_metadata_media(source, options).map_err(map_decode)?;
    validate_batch(source, options, batch, source_report)?;
    let mut count = CountSink { len: 0 };
    // The sizing pass must not report an output ceiling as an input-byte
    // failure.  It counts with a checked usize accumulator; the operation's
    // candidate ceiling is classified immediately below.
    rewrite_source(source, batch, usize::MAX, &mut count)?;
    let output_size = count.len;
    if output_size > options.max_output_bytes {
        return Err(RewriteError::limited(DecodeLimit::OutputBytes {
            observed: output_size,
            maximum: options.max_output_bytes,
        }));
    }
    let fields = source_report
        .fields
        .saturating_mul(3)
        .saturating_add(batch.data_additions.len().saturating_mul(5))
        .saturating_add(batch.owner_updates.len().saturating_mul(5));
    // Preparation scans the source once, sizes the raw-preserving stream,
    // and execution writes plus verifies one candidate.  Charge the source
    // scan and two output-width passes; this is finite, deterministic, and
    // leaves `for_source` useful for ordinary small additions.
    let work_bytes = source_report
        .work_bytes
        .saturating_add(output_size.saturating_mul(2));
    if fields > options.max_fields {
        return Err(RewriteError::limited(DecodeLimit::Fields {
            observed: fields,
            maximum: options.max_fields,
        }));
    }
    if work_bytes > options.max_work_bytes {
        return Err(RewriteError::limited(DecodeLimit::Work {
            observed: work_bytes,
            maximum: options.max_work_bytes,
        }));
    }
    let mut candidate_depth = source_report.max_depth;
    if !batch.data_additions.is_empty() {
        candidate_depth = candidate_depth.max(2);
    }
    if !batch.owner_additions.is_empty()
        || !batch.owner_removals.is_empty()
        || !batch.owner_updates.is_empty()
    {
        candidate_depth = candidate_depth.max(4);
    }
    if candidate_depth > options.max_depth {
        return Err(RewriteError::limited(DecodeLimit::Nesting {
            observed: candidate_depth,
            maximum: options.max_depth,
        }));
    }
    let requirements = RewriteExecutionRequirements {
        output_bytes: output_size,
        fields,
        work_bytes,
        components: source_report.components,
        data_records: source_report
            .data_records
            .saturating_add(batch.data_additions.len())
            .saturating_sub(batch.data_removals.len()),
        owners: source_report
            .owners
            .saturating_add(batch.owner_additions.len())
            .saturating_sub(batch.owner_removals.len()),
        allocations: 1,
        retained_bytes: output_size,
        scratch_bytes: 0,
    };
    Ok(PreparedPackageMetadataMediaRewrite {
        source,
        batch,
        options,
        output_size,
        source_report,
        requirements,
    })
}

/// Convenience one-shot atomic rewrite using exact limits derived from the
/// source and operation-local preflight.
pub fn rewrite_package_metadata_media(
    source: &[u8],
    batch: MediaRewriteBatch<'_>,
    options: DecodeOptions,
) -> Result<MediaRewriteOutput, RewriteError> {
    let prepared = prepare_package_metadata_media_rewrite(source, batch, options)?;
    let limits = prepared.execution_requirements().exact_limits();
    prepared.execute(limits)
}

/// One new DataInfo record.  The digest is normally a 20-byte SHA-1 digest,
/// matching iWork's materialized-data metadata.  The optional filename and
/// materialized length are emitted only when present.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DataInfoAddition<'source> {
    identifier: u64,
    digest: &'source [u8],
    preferred_file_name: &'source str,
    file_name: Option<&'source str>,
    materialized_length: Option<u64>,
}

impl<'source> DataInfoAddition<'source> {
    /// Construct a DataInfo addition with only its required semantic values.
    #[must_use]
    pub const fn new(
        identifier: u64,
        digest: &'source [u8],
        preferred_file_name: &'source str,
    ) -> Self {
        Self {
            identifier,
            digest,
            preferred_file_name,
            file_name: None,
            materialized_length: None,
        }
    }

    #[must_use]
    pub const fn with_file_name(mut self, file_name: &'source str) -> Self {
        self.file_name = Some(file_name);
        self
    }

    #[must_use]
    pub const fn with_materialized_length(mut self, materialized_length: u64) -> Self {
        self.materialized_length = Some(materialized_length);
        self
    }

    #[must_use]
    pub const fn identifier(self) -> u64 {
        self.identifier
    }
    #[must_use]
    pub const fn digest(self) -> &'source [u8] {
        self.digest
    }
    #[must_use]
    pub const fn preferred_file_name(self) -> &'source str {
        self.preferred_file_name
    }
    #[must_use]
    pub const fn file_name(self) -> Option<&'source str> {
        self.file_name
    }
    #[must_use]
    pub const fn materialized_length(self) -> Option<u64> {
        self.materialized_length
    }
}

/// One exact DataInfo identity removal.  The record must be unambiguous and
/// free of component-level owners before it can be removed.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DataInfoRemoval {
    identifier: u64,
}

impl DataInfoRemoval {
    #[must_use]
    pub const fn new(identifier: u64) -> Self {
        Self { identifier }
    }

    #[must_use]
    pub const fn identifier(self) -> u64 {
        self.identifier
    }
}

/// One ComponentDataReference owner append.  If the parent data reference is
/// absent, the writer creates one canonical parent containing this owner.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DataReferenceOwnerAddition<'source> {
    component: ComponentSelector<'source>,
    data_identifier: u64,
    object_identifier: u64,
    count: u32,
}

impl<'source> DataReferenceOwnerAddition<'source> {
    #[must_use]
    pub const fn new(
        component: ComponentSelector<'source>,
        data_identifier: u64,
        object_identifier: u64,
        count: u32,
    ) -> Self {
        Self {
            component,
            data_identifier,
            object_identifier,
            count,
        }
    }

    #[must_use]
    pub const fn component(self) -> ComponentSelector<'source> {
        self.component
    }
    #[must_use]
    pub const fn data_identifier(self) -> u64 {
        self.data_identifier
    }
    #[must_use]
    pub const fn object_identifier(self) -> u64 {
        self.object_identifier
    }
    #[must_use]
    pub const fn count(self) -> u32 {
        self.count
    }
}

/// One exact ComponentDataReference owner removal.  `expected_count` is a
/// compare-and-remove guard; counts are never silently decremented.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DataReferenceOwnerRemoval<'source> {
    component: ComponentSelector<'source>,
    data_identifier: u64,
    object_identifier: u64,
    expected_count: u32,
}

impl<'source> DataReferenceOwnerRemoval<'source> {
    #[must_use]
    pub const fn new(
        component: ComponentSelector<'source>,
        data_identifier: u64,
        object_identifier: u64,
        expected_count: u32,
    ) -> Self {
        Self {
            component,
            data_identifier,
            object_identifier,
            expected_count,
        }
    }

    #[must_use]
    pub const fn component(self) -> ComponentSelector<'source> {
        self.component
    }
    #[must_use]
    pub const fn data_identifier(self) -> u64 {
        self.data_identifier
    }
    #[must_use]
    pub const fn object_identifier(self) -> u64 {
        self.object_identifier
    }
    #[must_use]
    pub const fn expected_count(self) -> u32 {
        self.expected_count
    }
}

/// One exact ComponentDataReference owner count transition.
///
/// The source owner must match `expected_count` exactly and the resulting
/// count must remain non-zero.  A transition to zero is represented by
/// [`DataReferenceOwnerRemoval`], which lets the rewriter retain its existing
/// parent-removal semantics and avoids emitting an invalid zero-count owner.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DataReferenceOwnerCountUpdate<'source> {
    component: ComponentSelector<'source>,
    data_identifier: u64,
    object_identifier: u64,
    expected_count: u32,
    new_count: u32,
}

impl<'source> DataReferenceOwnerCountUpdate<'source> {
    /// Construct a compare-and-update request.  Zero counts are rejected
    /// during preparation, keeping this constructor allocation-free and
    /// consistent with the other borrowed batch operations.
    #[must_use]
    pub const fn new(
        component: ComponentSelector<'source>,
        data_identifier: u64,
        object_identifier: u64,
        expected_count: u32,
        new_count: u32,
    ) -> Self {
        Self {
            component,
            data_identifier,
            object_identifier,
            expected_count,
            new_count,
        }
    }

    #[must_use]
    pub const fn component(self) -> ComponentSelector<'source> {
        self.component
    }
    #[must_use]
    pub const fn data_identifier(self) -> u64 {
        self.data_identifier
    }
    #[must_use]
    pub const fn object_identifier(self) -> u64 {
        self.object_identifier
    }
    #[must_use]
    pub const fn expected_count(self) -> u32 {
        self.expected_count
    }
    #[must_use]
    pub const fn new_count(self) -> u32 {
        self.new_count
    }
}

/// One atomic metadata-media transaction.  Each request is evaluated against
/// one source snapshot and either all selected records are published or no
/// bytes escape.
#[derive(Debug, Clone, Copy)]
pub struct MediaRewriteBatch<'source> {
    data_additions: &'source [DataInfoAddition<'source>],
    data_removals: &'source [DataInfoRemoval],
    owner_additions: &'source [DataReferenceOwnerAddition<'source>],
    owner_removals: &'source [DataReferenceOwnerRemoval<'source>],
    owner_updates: &'source [DataReferenceOwnerCountUpdate<'source>],
}

impl<'source> MediaRewriteBatch<'source> {
    #[must_use]
    pub const fn new(
        data_additions: &'source [DataInfoAddition<'source>],
        data_removals: &'source [DataInfoRemoval],
        owner_additions: &'source [DataReferenceOwnerAddition<'source>],
        owner_removals: &'source [DataReferenceOwnerRemoval<'source>],
    ) -> Self {
        Self {
            data_additions,
            data_removals,
            owner_additions,
            owner_removals,
            owner_updates: &[],
        }
    }

    /// Construct a batch including exact non-zero owner count transitions.
    ///
    /// [`Self::new`] remains the four-slice constructor for source
    /// compatibility; callers that need count transitions can use this
    /// constructor or [`Self::with_owner_updates`].
    #[must_use]
    pub const fn new_with_owner_updates(
        data_additions: &'source [DataInfoAddition<'source>],
        data_removals: &'source [DataInfoRemoval],
        owner_additions: &'source [DataReferenceOwnerAddition<'source>],
        owner_removals: &'source [DataReferenceOwnerRemoval<'source>],
        owner_updates: &'source [DataReferenceOwnerCountUpdate<'source>],
    ) -> Self {
        Self {
            data_additions,
            data_removals,
            owner_additions,
            owner_removals,
            owner_updates,
        }
    }

    /// Return a copy of this batch with exact owner count transitions.
    #[must_use]
    pub const fn with_owner_updates(
        mut self,
        owner_updates: &'source [DataReferenceOwnerCountUpdate<'source>],
    ) -> Self {
        self.owner_updates = owner_updates;
        self
    }

    /// Alias emphasizing that these are count transitions rather than owner
    /// additions or removals.
    #[must_use]
    pub const fn with_owner_count_updates(
        self,
        owner_updates: &'source [DataReferenceOwnerCountUpdate<'source>],
    ) -> Self {
        self.with_owner_updates(owner_updates)
    }

    #[must_use]
    pub const fn empty() -> Self {
        Self::new(&[], &[], &[], &[])
    }

    #[must_use]
    pub const fn data_additions(self) -> &'source [DataInfoAddition<'source>] {
        self.data_additions
    }
    #[must_use]
    pub const fn data_removals(self) -> &'source [DataInfoRemoval] {
        self.data_removals
    }
    #[must_use]
    pub const fn owner_additions(self) -> &'source [DataReferenceOwnerAddition<'source>] {
        self.owner_additions
    }
    #[must_use]
    pub const fn owner_removals(self) -> &'source [DataReferenceOwnerRemoval<'source>] {
        self.owner_removals
    }
    #[must_use]
    pub const fn owner_updates(self) -> &'source [DataReferenceOwnerCountUpdate<'source>] {
        self.owner_updates
    }
    #[must_use]
    pub const fn owner_count_updates(self) -> &'source [DataReferenceOwnerCountUpdate<'source>] {
        self.owner_updates
    }

    #[must_use]
    pub const fn is_empty(self) -> bool {
        self.data_additions.is_empty()
            && self.data_removals.is_empty()
            && self.owner_additions.is_empty()
            && self.owner_removals.is_empty()
            && self.owner_updates.is_empty()
    }
}

/// Additional allocation and retention limits for executing a prepared
/// rewrite.  The source is always borrowed; only the returned candidate is
/// retained.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RewriteExecutionLimits {
    pub max_output_bytes: usize,
    pub max_fields: usize,
    pub max_work_bytes: usize,
    pub max_components: usize,
    pub max_data_records: usize,
    pub max_owners: usize,
    pub max_allocations: usize,
    pub max_retained_bytes: usize,
    pub max_scratch_bytes: usize,
}

/// Exact operation resources required by an allocation-bearing rewrite.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RewriteExecutionRequirements {
    output_bytes: usize,
    fields: usize,
    work_bytes: usize,
    components: usize,
    data_records: usize,
    owners: usize,
    allocations: usize,
    retained_bytes: usize,
    scratch_bytes: usize,
}

impl RewriteExecutionRequirements {
    #[must_use]
    pub const fn output_bytes(self) -> usize {
        self.output_bytes
    }
    #[must_use]
    pub const fn fields(self) -> usize {
        self.fields
    }
    #[must_use]
    pub const fn work_bytes(self) -> usize {
        self.work_bytes
    }
    #[must_use]
    pub const fn components(self) -> usize {
        self.components
    }
    #[must_use]
    pub const fn data_records(self) -> usize {
        self.data_records
    }
    #[must_use]
    pub const fn owners(self) -> usize {
        self.owners
    }
    #[must_use]
    pub const fn allocations(self) -> usize {
        self.allocations
    }
    #[must_use]
    pub const fn retained_bytes(self) -> usize {
        self.retained_bytes
    }
    #[must_use]
    pub const fn scratch_bytes(self) -> usize {
        self.scratch_bytes
    }

    #[must_use]
    pub const fn exact_limits(self) -> RewriteExecutionLimits {
        RewriteExecutionLimits {
            max_output_bytes: self.output_bytes,
            max_fields: self.fields,
            max_work_bytes: self.work_bytes,
            max_components: self.components,
            max_data_records: self.data_records,
            max_owners: self.owners,
            max_allocations: self.allocations,
            max_retained_bytes: self.retained_bytes,
            max_scratch_bytes: self.scratch_bytes,
        }
    }
}

/// Candidate bytes and exact resource report.
#[derive(Debug, PartialEq, Eq)]
pub struct MediaRewriteOutput {
    bytes: Vec<u8>,
    report: RewriteReport,
}

impl MediaRewriteOutput {
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    #[must_use]
    pub const fn report(&self) -> RewriteReport {
        self.report
    }

    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        self.bytes
    }
}

/// Exact resources consumed by the candidate publication.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RewriteReport {
    input_bytes: usize,
    output_bytes: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    components: usize,
    data_records: usize,
    owners: usize,
    data_additions: usize,
    data_removals: usize,
    owner_additions: usize,
    owner_removals: usize,
    owner_updates: usize,
    allocations: usize,
    retained_bytes: usize,
    scratch_bytes: usize,
}

impl RewriteReport {
    #[must_use]
    pub const fn input_bytes(self) -> usize {
        self.input_bytes
    }
    #[must_use]
    pub const fn output_bytes(self) -> usize {
        self.output_bytes
    }
    #[must_use]
    pub const fn fields(self) -> usize {
        self.fields
    }
    #[must_use]
    pub const fn work_bytes(self) -> usize {
        self.work_bytes
    }
    #[must_use]
    pub const fn max_depth(self) -> u32 {
        self.max_depth
    }
    #[must_use]
    pub const fn components(self) -> usize {
        self.components
    }
    #[must_use]
    pub const fn data_records(self) -> usize {
        self.data_records
    }
    #[must_use]
    pub const fn owners(self) -> usize {
        self.owners
    }
    #[must_use]
    pub const fn data_additions(self) -> usize {
        self.data_additions
    }
    #[must_use]
    pub const fn data_removals(self) -> usize {
        self.data_removals
    }
    #[must_use]
    pub const fn owner_additions(self) -> usize {
        self.owner_additions
    }
    #[must_use]
    pub const fn owner_removals(self) -> usize {
        self.owner_removals
    }
    #[must_use]
    pub const fn owner_updates(self) -> usize {
        self.owner_updates
    }
    #[must_use]
    pub const fn allocations(self) -> usize {
        self.allocations
    }
    #[must_use]
    pub const fn retained_bytes(self) -> usize {
        self.retained_bytes
    }
    #[must_use]
    pub const fn scratch_bytes(self) -> usize {
        self.scratch_bytes
    }
}

/// Atomic rewrite failure.  No candidate bytes are returned on an error.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RewriteError {
    limit: Option<DecodeLimit>,
    reason: Option<InvalidReason>,
    allocation: Option<usize>,
}

impl RewriteError {
    #[must_use]
    pub const fn resource_limit(self) -> Option<DecodeLimit> {
        self.limit
    }
    #[must_use]
    pub const fn invalid_reason(self) -> Option<InvalidReason> {
        self.reason
    }
    #[must_use]
    pub const fn allocation_request(self) -> Option<usize> {
        self.allocation
    }
    const fn from_decode(error: DecodeError) -> Self {
        Self {
            limit: error.limit,
            reason: error.reason,
            allocation: None,
        }
    }
    const fn limited(limit: DecodeLimit) -> Self {
        Self {
            limit: Some(limit),
            reason: None,
            allocation: None,
        }
    }
    const fn invalid(reason: InvalidReason) -> Self {
        Self {
            limit: None,
            reason: Some(reason),
            allocation: None,
        }
    }
    const fn allocation(bytes: usize) -> Self {
        Self {
            limit: None,
            reason: None,
            allocation: Some(bytes),
        }
    }
}

impl fmt::Display for RewriteError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("invalid PackageMetadata media rewrite")
    }
}

impl std::error::Error for RewriteError {}

#[derive(Clone, Copy)]
struct WireField {
    number: u32,
    wire_type: u8,
    start: usize,
    end: usize,
    payload_start: usize,
    payload_end: usize,
}

impl WireField {
    fn payload(self, source: &[u8]) -> &[u8] {
        &source[self.payload_start..self.payload_end]
    }
}

fn encoded_len(value: u64) -> usize {
    let significant_bits = u64::BITS - value.leading_zeros();
    let groups = significant_bits.saturating_add(6) / 7;
    usize::try_from(groups.max(1)).unwrap_or(10)
}

fn decode_varint(source: &[u8], offset: usize) -> Result<(u64, usize), DecodeError> {
    let mut value = 0u64;
    let mut index = offset;
    let mut shift = 0u32;
    while shift < 64 {
        let byte = *source
            .get(index)
            .ok_or_else(|| DecodeError::invalid(InvalidReason::MalformedWire))?;
        index = index
            .checked_add(1)
            .ok_or_else(|| DecodeError::invalid(InvalidReason::MalformedWire))?;
        if shift == 63 && byte > 1 {
            return Err(DecodeError::invalid(InvalidReason::MalformedWire));
        }
        value |= u64::from(byte & 0x7f) << shift;
        if byte & 0x80 == 0 {
            if index - offset != encoded_len(value) {
                return Err(DecodeError::invalid(InvalidReason::MalformedWire));
            }
            return Ok((value, index));
        }
        shift += 7;
    }
    Err(DecodeError::invalid(InvalidReason::MalformedWire))
}

fn next_field(source: &[u8], offset: usize) -> Result<Option<WireField>, DecodeError> {
    if offset == source.len() {
        return Ok(None);
    }
    let start = offset;
    let (key, after_key) = decode_varint(source, offset)?;
    let number = u32::try_from(key >> 3)
        .map_err(|_error| DecodeError::invalid(InvalidReason::MalformedWire))?;
    let wire_type = u8::try_from(key & 7)
        .map_err(|_error| DecodeError::invalid(InvalidReason::MalformedWire))?;
    if number == 0 || number > MAX_FIELD_NUMBER || matches!(wire_type, 3 | 4) {
        return Err(DecodeError::invalid(InvalidReason::MalformedWire));
    }
    let (payload_start, payload_end) = match wire_type {
        0 => {
            let (_, end) = decode_varint(source, after_key)?;
            (after_key, end)
        },
        1 => {
            let end = after_key
                .checked_add(8)
                .ok_or_else(|| DecodeError::invalid(InvalidReason::MalformedWire))?;
            if end > source.len() {
                return Err(DecodeError::invalid(InvalidReason::MalformedWire));
            }
            (after_key, end)
        },
        2 => {
            let (length, payload) = decode_varint(source, after_key)?;
            let length = usize::try_from(length)
                .map_err(|_error| DecodeError::invalid(InvalidReason::MalformedWire))?;
            let end = payload
                .checked_add(length)
                .ok_or_else(|| DecodeError::invalid(InvalidReason::MalformedWire))?;
            if end > source.len() {
                return Err(DecodeError::invalid(InvalidReason::MalformedWire));
            }
            (payload, end)
        },
        5 => {
            let end = after_key
                .checked_add(4)
                .ok_or_else(|| DecodeError::invalid(InvalidReason::MalformedWire))?;
            if end > source.len() {
                return Err(DecodeError::invalid(InvalidReason::MalformedWire));
            }
            (after_key, end)
        },
        _ => return Err(DecodeError::invalid(InvalidReason::MalformedWire)),
    };
    Ok(Some(WireField {
        number,
        wire_type,
        start,
        end: payload_end,
        payload_start,
        payload_end,
    }))
}

fn fields<'a>(source: &'a [u8]) -> FieldIter<'a> {
    FieldIter { source, offset: 0 }
}

struct FieldIter<'a> {
    source: &'a [u8],
    offset: usize,
}

impl Iterator for FieldIter<'_> {
    type Item = Result<WireField, DecodeError>;

    fn next(&mut self) -> Option<Self::Item> {
        match next_field(self.source, self.offset) {
            Ok(Some(field)) => {
                self.offset = field.end;
                Some(Ok(field))
            },
            Ok(None) => None,
            Err(error) => {
                self.offset = self.source.len();
                Some(Err(error))
            },
        }
    }
}

fn varint(source: &[u8], field: WireField) -> Result<u64, DecodeError> {
    if field.wire_type != 0 {
        return Err(DecodeError::invalid(InvalidReason::MalformedWire));
    }
    let (value, end) = decode_varint(source, field.payload_start)?;
    if end != field.payload_end {
        return Err(DecodeError::invalid(InvalidReason::MalformedWire));
    }
    Ok(value)
}

fn bytes(source: &[u8], field: WireField) -> Result<&[u8], DecodeError> {
    if field.wire_type != 2 {
        return Err(DecodeError::invalid(InvalidReason::MalformedWire));
    }
    Ok(field.payload(source))
}

fn utf8(source: &[u8], field: WireField) -> Result<&str, DecodeError> {
    str::from_utf8(bytes(source, field)?)
        .map_err(|_error| DecodeError::invalid(InvalidReason::InvalidName))
}

#[cfg(test)]
#[allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "The small in-module wire fixtures keep codec invariants readable."
)]
mod tests {
    use super::*;

    fn varint(output: &mut Vec<u8>, mut value: u64) {
        loop {
            let mut byte = (value & 0x7f) as u8;
            value >>= 7;
            if value != 0 {
                byte |= 0x80;
            }
            output.push(byte);
            if value == 0 {
                return;
            }
        }
    }

    fn varint_field(output: &mut Vec<u8>, number: u32, value: u64) {
        varint(output, u64::from(number) << 3);
        varint(output, value);
    }

    fn bytes_field(output: &mut Vec<u8>, number: u32, payload: &[u8]) {
        varint(output, (u64::from(number) << 3) | 2);
        varint(
            output,
            u64::try_from(payload.len()).expect("test payload fits in a varint"),
        );
        output.extend_from_slice(payload);
    }

    fn data_info(identifier: u64, seed: u8) -> Vec<u8> {
        let digest = [seed; SHA1_DIGEST_BYTES];
        let mut payload = Vec::new();
        varint_field(&mut payload, DATA_IDENTIFIER_FIELD, identifier);
        bytes_field(&mut payload, DATA_DIGEST_FIELD, &digest);
        bytes_field(&mut payload, DATA_PREFERRED_NAME_FIELD, b"clip.m4a");
        payload
    }

    fn source() -> Vec<u8> {
        let mut source = Vec::new();
        varint_field(&mut source, ROOT_LAST_IDENTIFIER_FIELD, 10);
        let record = data_info(1, 0x10);
        bytes_field(&mut source, ROOT_DATA_INFO_FIELD, &record);
        source
    }

    fn source_with_owner(count: u32, owner_unknown: bool) -> Vec<u8> {
        let mut owner = Vec::new();
        varint_field(&mut owner, OWNER_OBJECT_IDENTIFIER_FIELD, 77);
        varint_field(&mut owner, OWNER_COUNT_FIELD, u64::from(count));
        if owner_unknown {
            varint_field(&mut owner, 99, 1);
        }

        let mut reference = Vec::new();
        varint_field(&mut reference, DATA_IDENTIFIER_FIELD, 1);
        bytes_field(&mut reference, OWNER_COUNT_FIELD, &owner);

        let mut component = Vec::new();
        varint_field(&mut component, COMPONENT_IDENTIFIER_FIELD, 9);
        bytes_field(
            &mut component,
            COMPONENT_PREFERRED_LOCATOR_FIELD,
            b"Document",
        );
        bytes_field(&mut component, COMPONENT_DATA_REFERENCE_FIELD, &reference);

        let mut source = source();
        bytes_field(&mut source, ROOT_COMPONENT_FIELD, &component);
        source
    }

    #[test]
    fn strict_scan_rejects_noncanonical_varints() {
        let source = [0x08, 0x80, 0x00];
        let error = inspect_package_metadata_media(&source, DecodeOptions::for_source(&source))
            .expect_err("overlong root identifier must be rejected");
        assert_eq!(error.invalid_reason(), Some(InvalidReason::MalformedWire));
    }

    #[test]
    fn native_legacy_data_info_allows_explicit_empty_optional_name_and_missing_length() {
        let digest = [0x42; SHA1_DIGEST_BYTES];
        let mut record = Vec::new();
        varint_field(&mut record, DATA_IDENTIFIER_FIELD, 7);
        bytes_field(&mut record, DATA_DIGEST_FIELD, &digest);
        bytes_field(&mut record, DATA_PREFERRED_NAME_FIELD, b"native.jpg");
        bytes_field(&mut record, DATA_FILE_NAME_FIELD, b"");
        varint_field(&mut record, 99, 1);

        let mut source = Vec::new();
        varint_field(&mut source, ROOT_LAST_IDENTIFIER_FIELD, 10);
        bytes_field(&mut source, ROOT_DATA_INFO_FIELD, &record);

        let report = inspect_package_metadata_media(&source, DecodeOptions::for_source(&source))
            .expect("unselected native legacy DataInfo remains inspectable");
        assert_eq!(report.data_records(), 1);
        assert_eq!(report.unknown_records(), 1);
    }

    #[test]
    fn data_info_lifecycle_round_trips_without_reencoding_source() {
        let source = source();
        let digest = [0x20; SHA1_DIGEST_BYTES];
        let additions = [DataInfoAddition::new(2, &digest, "new.m4a")];
        let added = rewrite_package_metadata_media(
            &source,
            MediaRewriteBatch::new(&additions, &[], &[], &[]),
            DecodeOptions::for_source(&source),
        )
        .expect("canonical DataInfo addition");
        assert!(added.bytes().starts_with(&source));
        assert_eq!(
            inspect_package_metadata_media(
                added.bytes(),
                DecodeOptions::for_source(added.bytes()),
            )
            .expect("added source remains canonical")
            .data_records(),
            2
        );

        let removals = [DataInfoRemoval::new(2)];
        let restored = rewrite_package_metadata_media(
            added.bytes(),
            MediaRewriteBatch::new(&[], &removals, &[], &[]),
            DecodeOptions::for_source(added.bytes()),
        )
        .expect("exact DataInfo removal");
        assert_eq!(restored.bytes(), source.as_slice());
    }

    #[test]
    fn owner_count_transition_is_exact_reversible_and_width_safe() {
        let source = source_with_owner(127, false);
        let selector = ComponentSelector::new(9, "Document");
        let forward = [DataReferenceOwnerCountUpdate::new(
            selector, 1, 77, 127, 128,
        )];
        let changed = rewrite_package_metadata_media(
            &source,
            MediaRewriteBatch::empty().with_owner_count_updates(&forward),
            DecodeOptions::for_source(&source).with_max_output_bytes(source.len() + 2),
        )
        .expect("count transition grows nested varint lengths safely");
        assert_eq!(changed.report().owner_updates(), 1);
        let facts = owner_matches(changed.bytes(), selector, 1, 77)
            .expect("rewritten owner remains canonical");
        assert_eq!(facts.owner_count, 128);

        let inverse = [DataReferenceOwnerCountUpdate::new(
            selector, 1, 77, 128, 127,
        )];
        let restored = rewrite_package_metadata_media(
            changed.bytes(),
            MediaRewriteBatch::empty().with_owner_updates(&inverse),
            DecodeOptions::for_source(changed.bytes()),
        )
        .expect("inverse count transition");
        assert_eq!(restored.bytes(), source.as_slice());
    }

    #[test]
    fn owner_count_transition_rejects_stale_zero_unknown_and_conflicting_requests() {
        let source = source_with_owner(2, false);
        let selector = ComponentSelector::new(9, "Document");
        for update in [
            DataReferenceOwnerCountUpdate::new(selector, 1, 77, 3, 1),
            DataReferenceOwnerCountUpdate::new(selector, 1, 77, 2, 0),
            DataReferenceOwnerCountUpdate::new(selector, 1, 77, 0, 1),
        ] {
            let error = rewrite_package_metadata_media(
                &source,
                MediaRewriteBatch::empty().with_owner_updates(std::slice::from_ref(&update)),
                DecodeOptions::for_source(&source),
            )
            .expect_err("invalid compare-and-set request must be atomic");
            assert!(matches!(
                error.invalid_reason(),
                Some(InvalidReason::ExistingOwnerCollision | InvalidReason::InvalidIdentifier)
            ));
        }

        let unknown_source = source_with_owner(2, true);
        let update = [DataReferenceOwnerCountUpdate::new(selector, 1, 77, 2, 1)];
        let error = rewrite_package_metadata_media(
            &unknown_source,
            MediaRewriteBatch::empty().with_owner_updates(&update),
            DecodeOptions::for_source(&unknown_source),
        )
        .expect_err("unknown selected owner fields must fail closed");
        assert_eq!(
            error.invalid_reason(),
            Some(InvalidReason::UnknownSelectedRecord)
        );

        let removal = [DataReferenceOwnerRemoval::new(selector, 1, 77, 2)];
        let error = rewrite_package_metadata_media(
            &source,
            MediaRewriteBatch::new(&[], &[], &[], &removal).with_owner_updates(&update),
            DecodeOptions::for_source(&source),
        )
        .expect_err("one owner cannot be removed and updated together");
        assert_eq!(
            error.invalid_reason(),
            Some(InvalidReason::DuplicateOperation)
        );
    }
}
