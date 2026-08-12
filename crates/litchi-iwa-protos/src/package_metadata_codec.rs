//! Strict raw-preserving PackageMetadata sparse-publication codec.
//!
//! Caller-owned source bytes remain authoritative. The codec performs a
//! canonical handwritten preflight, uses private Buffa lazy views only as
//! borrowed parity oracles, constructs one fallibly allocated candidate, and
//! strictly re-decodes that candidate before any bytes escape.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "Wire helpers stay beside the generated-free publication model."
)]

use core::{fmt, mem::size_of, str};

#[cfg(test)]
use core::sync::atomic::{AtomicUsize, Ordering};

use buffa::DecodeOptions as BuffaDecodeOptions;

use crate::buffa_package_metadata_generated::LitchiIwaPackageMetadataProjection as projection;

const MAX_RECURSION: u32 = 64;
const MAX_FIELD_NUMBER: u32 = 0x1fff_ffff;

#[cfg(test)]
static OUTPUT_ALLOCATIONS: AtomicUsize = AtomicUsize::new(0);

/// Finite aggregate policy for one decode, rewrite, and verification cycle.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RewriteOptions {
    max_input_bytes: usize,
    max_output_bytes: usize,
    max_fields: usize,
    max_work_bytes: usize,
    recursion_limit: u32,
    max_components: usize,
    max_references: usize,
    max_additions: usize,
}

impl RewriteOptions {
    #[must_use]
    pub const fn new(
        max_input_bytes: usize,
        max_output_bytes: usize,
        max_fields: usize,
        max_work_bytes: usize,
        recursion_limit: u32,
        max_components: usize,
        max_references: usize,
        max_additions: usize,
    ) -> Self {
        Self {
            max_input_bytes,
            max_output_bytes,
            max_fields,
            max_work_bytes,
            recursion_limit,
            max_components,
            max_references,
            max_additions,
        }
    }
    #[must_use]
    pub const fn max_input_bytes(self) -> usize {
        self.max_input_bytes
    }
    #[must_use]
    pub const fn max_output_bytes(self) -> usize {
        self.max_output_bytes
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
    pub const fn recursion_limit(self) -> u32 {
        self.recursion_limit
    }
    #[must_use]
    pub const fn max_components(self) -> usize {
        self.max_components
    }
    #[must_use]
    pub const fn max_references(self) -> usize {
        self.max_references
    }
    #[must_use]
    pub const fn max_additions(self) -> usize {
        self.max_additions
    }
}

/// Exact, locator-authorized component selector.
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

/// Canonical 128-bit UUID scalar pair.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct UuidBits {
    lower: u64,
    upper: u64,
}

impl UuidBits {
    #[must_use]
    pub const fn new(lower: u64, upper: u64) -> Self {
        Self { lower, upper }
    }
    #[must_use]
    pub const fn lower(self) -> u64 {
        self.lower
    }
    #[must_use]
    pub const fn upper(self) -> u64 {
        self.upper
    }
}

/// One object-to-UUID registry append.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ObjectUuidAddition<'source> {
    component: ComponentSelector<'source>,
    object_identifier: u64,
    uuid: UuidBits,
}

impl<'source> ObjectUuidAddition<'source> {
    #[must_use]
    pub const fn new(
        component: ComponentSelector<'source>,
        object_identifier: u64,
        uuid: UuidBits,
    ) -> Self {
        Self {
            component,
            object_identifier,
            uuid,
        }
    }
    #[must_use]
    pub const fn component(self) -> ComponentSelector<'source> {
        self.component
    }
    #[must_use]
    pub const fn object_identifier(self) -> u64 {
        self.object_identifier
    }
    #[must_use]
    pub const fn uuid(self) -> UuidBits {
        self.uuid
    }
}

/// One source-component external-reference append.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ExternalReferenceAddition<'source> {
    source: ComponentSelector<'source>,
    target: ComponentSelector<'source>,
    object_identifier: u64,
    is_weak: Option<bool>,
}

/// One exact current-component object-to-UUID registry removal.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ObjectUuidRemoval<'source> {
    component: ComponentSelector<'source>,
    object_identifier: u64,
    expected_uuid: UuidBits,
}

impl<'source> ObjectUuidRemoval<'source> {
    #[must_use]
    pub const fn new(
        component: ComponentSelector<'source>,
        object_identifier: u64,
        expected_uuid: UuidBits,
    ) -> Self {
        Self {
            component,
            object_identifier,
            expected_uuid,
        }
    }
    #[must_use]
    pub const fn component(self) -> ComponentSelector<'source> {
        self.component
    }
    #[must_use]
    pub const fn object_identifier(self) -> u64 {
        self.object_identifier
    }
    #[must_use]
    pub const fn expected_uuid(self) -> UuidBits {
        self.expected_uuid
    }
}

/// One exact current, unversioned component external-reference removal.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ExternalReferenceRemoval<'source> {
    source: ComponentSelector<'source>,
    target: ComponentSelector<'source>,
    object_identifier: u64,
    expected_is_weak: Option<bool>,
}

/// One exact ComponentDataReference owner removal.
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

impl<'source> ExternalReferenceRemoval<'source> {
    #[must_use]
    pub const fn new(
        source: ComponentSelector<'source>,
        target: ComponentSelector<'source>,
        object_identifier: u64,
        expected_is_weak: Option<bool>,
    ) -> Self {
        Self {
            source,
            target,
            object_identifier,
            expected_is_weak,
        }
    }
    #[must_use]
    pub const fn source(self) -> ComponentSelector<'source> {
        self.source
    }
    #[must_use]
    pub const fn target(self) -> ComponentSelector<'source> {
        self.target
    }
    #[must_use]
    pub const fn object_identifier(self) -> u64 {
        self.object_identifier
    }
    #[must_use]
    pub const fn expected_is_weak(self) -> Option<bool> {
        self.expected_is_weak
    }
}

/// Borrowed atomic removal request. The last object identifier is retained.
#[derive(Debug, Clone, Copy)]
pub struct RemovalBatch<'source> {
    expected_last_object_identifier: u64,
    object_uuids: &'source [ObjectUuidRemoval<'source>],
    external_references: &'source [ExternalReferenceRemoval<'source>],
    data_reference_owners: &'source [DataReferenceOwnerRemoval<'source>],
}

impl<'source> RemovalBatch<'source> {
    #[must_use]
    pub const fn new(
        expected_last_object_identifier: u64,
        object_uuids: &'source [ObjectUuidRemoval<'source>],
        external_references: &'source [ExternalReferenceRemoval<'source>],
        data_reference_owners: &'source [DataReferenceOwnerRemoval<'source>],
    ) -> Self {
        Self {
            expected_last_object_identifier,
            object_uuids,
            external_references,
            data_reference_owners,
        }
    }
    #[must_use]
    pub const fn expected_last_object_identifier(self) -> u64 {
        self.expected_last_object_identifier
    }
    #[must_use]
    pub const fn object_uuids(self) -> &'source [ObjectUuidRemoval<'source>] {
        self.object_uuids
    }
    #[must_use]
    pub const fn external_references(self) -> &'source [ExternalReferenceRemoval<'source>] {
        self.external_references
    }
    #[must_use]
    pub const fn data_reference_owners(self) -> &'source [DataReferenceOwnerRemoval<'source>] {
        self.data_reference_owners
    }
}

impl<'source> ExternalReferenceAddition<'source> {
    #[must_use]
    pub const fn new(
        source: ComponentSelector<'source>,
        target: ComponentSelector<'source>,
        object_identifier: u64,
        is_weak: Option<bool>,
    ) -> Self {
        Self {
            source,
            target,
            object_identifier,
            is_weak,
        }
    }
    #[must_use]
    pub const fn source(self) -> ComponentSelector<'source> {
        self.source
    }
    #[must_use]
    pub const fn target(self) -> ComponentSelector<'source> {
        self.target
    }
    #[must_use]
    pub const fn object_identifier(self) -> u64 {
        self.object_identifier
    }
    #[must_use]
    pub const fn is_weak(self) -> Option<bool> {
        self.is_weak
    }
}

/// Borrowed atomic publication request.
#[derive(Debug, Clone, Copy)]
pub struct Batch<'source> {
    expected_last_object_identifier: u64,
    new_last_object_identifier: u64,
    object_uuids: &'source [ObjectUuidAddition<'source>],
    external_references: &'source [ExternalReferenceAddition<'source>],
}

impl<'source> Batch<'source> {
    #[must_use]
    pub const fn new(
        expected_last_object_identifier: u64,
        new_last_object_identifier: u64,
        object_uuids: &'source [ObjectUuidAddition<'source>],
        external_references: &'source [ExternalReferenceAddition<'source>],
    ) -> Self {
        Self {
            expected_last_object_identifier,
            new_last_object_identifier,
            object_uuids,
            external_references,
        }
    }
    #[must_use]
    pub const fn expected_last_object_identifier(self) -> u64 {
        self.expected_last_object_identifier
    }
    #[must_use]
    pub const fn new_last_object_identifier(self) -> u64 {
        self.new_last_object_identifier
    }
    #[must_use]
    pub const fn object_uuids(self) -> &'source [ObjectUuidAddition<'source>] {
        self.object_uuids
    }
    #[must_use]
    pub const fn external_references(self) -> &'source [ExternalReferenceAddition<'source>] {
        self.external_references
    }
}

/// Exact aggregate consumption and allocation evidence.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RewriteReport {
    input_bytes: usize,
    output_bytes: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    components_scanned: usize,
    components_changed: usize,
    references_scanned: usize,
    source_references_scanned: usize,
    additions: usize,
    removals: usize,
    allocations: usize,
    retained_bytes: usize,
    scratch_bytes: usize,
}

macro_rules! report_accessors {
    ($(($name:ident, $ty:ty)),+ $(,)?) => {$(
        #[must_use]
        pub const fn $name(self) -> $ty { self.$name }
    )+};
}

impl RewriteReport {
    report_accessors!(
        (input_bytes, usize),
        (output_bytes, usize),
        (fields, usize),
        (work_bytes, usize),
        (max_depth, u32),
        (components_scanned, usize),
        (components_changed, usize),
        (references_scanned, usize),
        (source_references_scanned, usize),
        (additions, usize),
        (removals, usize),
        (allocations, usize),
        (retained_bytes, usize),
        (scratch_bytes, usize)
    );
}

/// Verified owned candidate plus its exact report.
#[derive(Debug, PartialEq, Eq)]
pub struct RewriteOutput {
    bytes: Vec<u8>,
    report: RewriteReport,
}

impl RewriteOutput {
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

/// Borrowed identity and locator facts for one PackageMetadata component.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ComponentDescriptor<'source> {
    identifier: u64,
    preferred_locator: &'source str,
    locator: Option<&'source str>,
    current: bool,
}

impl<'source> ComponentDescriptor<'source> {
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
    pub const fn is_current(self) -> bool {
        self.current
    }
}

/// Borrowed existing object-to-UUID binding.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ObjectUuidDescriptor<'source> {
    component: ComponentDescriptor<'source>,
    object_identifier: u64,
    uuid: UuidBits,
}

impl<'source> ObjectUuidDescriptor<'source> {
    #[must_use]
    pub const fn component(self) -> ComponentDescriptor<'source> {
        self.component
    }
    #[must_use]
    pub const fn object_identifier(self) -> u64 {
        self.object_identifier
    }
    #[must_use]
    pub const fn uuid(self) -> UuidBits {
        self.uuid
    }
}

/// Borrowed existing component-external-reference record.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ExternalReferenceDescriptor<'source> {
    source: ComponentDescriptor<'source>,
    target_component_identifier: u64,
    object_identifier: Option<u64>,
    is_weak: Option<bool>,
    versioned: bool,
}

impl<'source> ExternalReferenceDescriptor<'source> {
    #[must_use]
    pub const fn source(self) -> ComponentDescriptor<'source> {
        self.source
    }
    #[must_use]
    pub const fn target_component_identifier(self) -> u64 {
        self.target_component_identifier
    }
    #[must_use]
    pub const fn object_identifier(self) -> Option<u64> {
        self.object_identifier
    }
    #[must_use]
    pub const fn is_weak(self) -> Option<bool> {
        self.is_weak
    }
    #[must_use]
    pub const fn is_versioned(self) -> bool {
        self.versioned
    }
}

/// Fallible streaming sink for strict PackageMetadata inspection.
///
/// Callers must discard observations if inspection returns an error.
pub trait PackageMetadataVisitor {
    fn visit_component(&mut self, _component: ComponentDescriptor<'_>) -> Result<(), RewriteError> {
        Ok(())
    }

    fn visit_object_uuid(
        &mut self,
        _binding: ObjectUuidDescriptor<'_>,
    ) -> Result<(), RewriteError> {
        Ok(())
    }

    fn visit_external_reference(
        &mut self,
        _reference: ExternalReferenceDescriptor<'_>,
    ) -> Result<(), RewriteError> {
        Ok(())
    }
}

/// Strict inspection result and exact aggregate decode evidence.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PackageMetadataInspection {
    last_object_identifier: u64,
    report: RewriteReport,
}

impl PackageMetadataInspection {
    #[must_use]
    pub const fn last_object_identifier(self) -> u64 {
        self.last_object_identifier
    }
    #[must_use]
    pub const fn report(self) -> RewriteReport {
        self.report
    }
}

/// Typed finite aggregate limit.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum RewriteLimit {
    InputBytes { observed: usize, maximum: usize },
    OutputBytes { observed: usize, maximum: usize },
    Fields { observed: usize, maximum: usize },
    Work { observed: usize, maximum: usize },
    Nesting { observed: u32, maximum: u32 },
    Components { observed: usize, maximum: usize },
    References { observed: usize, maximum: usize },
    Additions { observed: usize, maximum: usize },
}

/// Content-free semantic or structural refusal.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum InvalidReason {
    MalformedWire,
    LastIdentifierMismatch,
    LastIdentifierNotIncreasing,
    InvalidIdentifier,
    InvalidUuid,
    DuplicateAddition,
    ComponentMismatch,
    ExistingObjectCollision,
    ExistingUuidCollision,
    ExistingReferenceCollision,
    ConflictingWeakness,
    RemovalNotFound,
    DuplicateRemoval,
    RemovalMismatch,
    VersionedRemoval,
    CrossComponentRemoval,
    Verification,
}

/// Atomic rewrite failure. No candidate bytes escape on error.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RewriteError {
    limit: Option<RewriteLimit>,
    reason: Option<InvalidReason>,
    allocation: Option<usize>,
}

impl RewriteError {
    #[must_use]
    pub const fn resource_limit(&self) -> Option<RewriteLimit> {
        self.limit
    }
    #[must_use]
    pub const fn invalid_reason(&self) -> Option<InvalidReason> {
        self.reason
    }
    #[must_use]
    pub const fn allocation_request(&self) -> Option<usize> {
        self.allocation
    }
}

impl fmt::Display for RewriteError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("invalid PackageMetadata publication batch")
    }
}

impl std::error::Error for RewriteError {}

impl RewriteError {
    const fn invalid(reason: InvalidReason) -> Self {
        Self {
            limit: None,
            reason: Some(reason),
            allocation: None,
        }
    }
    const fn limited(limit: RewriteLimit) -> Self {
        Self {
            limit: Some(limit),
            reason: None,
            allocation: None,
        }
    }
    const fn allocation(requested: usize) -> Self {
        Self {
            limit: None,
            reason: None,
            allocation: Some(requested),
        }
    }
}

impl RewriteOptions {
    fn buffa(self) -> BuffaDecodeOptions {
        BuffaDecodeOptions::new()
            .with_max_message_size(self.max_output_bytes.max(self.max_input_bytes))
            .with_unknown_field_limit(self.max_fields)
            .with_element_memory_limit(0)
            .with_recursion_limit(self.recursion_limit)
    }
}

#[derive(Clone, Copy)]
enum ScanMode {
    Source,
    Verification,
}

#[derive(Default, Clone, Copy)]
struct SelectorCount {
    identifier: usize,
    locator: usize,
    exact: usize,
}

struct ScanState {
    selectors: Vec<SelectorCount>,
    object_matches: Vec<usize>,
    external_matches: Vec<usize>,
}

impl ScanState {
    fn new(batch: Batch<'_>, budget: &mut Budget) -> Result<Self, RewriteError> {
        let selectors = batch
            .object_uuids
            .len()
            .checked_add(
                batch
                    .external_references
                    .len()
                    .checked_mul(2)
                    .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?,
            )
            .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?;
        Ok(Self {
            selectors: zeroed_vec(selectors, budget)?,
            object_matches: zeroed_vec(batch.object_uuids.len(), budget)?,
            external_matches: zeroed_vec(batch.external_references.len(), budget)?,
        })
    }

    fn validate_selectors(&self) -> Result<(), RewriteError> {
        if self
            .selectors
            .iter()
            .any(|count| count.identifier != 1 || count.locator != 1 || count.exact != 1)
        {
            return Err(RewriteError::invalid(InvalidReason::ComponentMismatch));
        }
        Ok(())
    }

    fn validate_verification(&self) -> Result<(), RewriteError> {
        self.validate_selectors()?;
        if self.object_matches.iter().any(|count| *count != 1)
            || self.external_matches.iter().any(|count| *count != 1)
        {
            return Err(RewriteError::invalid(InvalidReason::Verification));
        }
        Ok(())
    }
}

fn zeroed_vec<T: Default + Clone>(
    amount: usize,
    budget: &mut Budget,
) -> Result<Vec<T>, RewriteError> {
    let mut output = Vec::new();
    if amount != 0 {
        output
            .try_reserve_exact(amount)
            .map_err(|_error| RewriteError::allocation(amount))?;
        budget.allocation(
            amount
                .checked_mul(size_of::<T>())
                .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?,
        )?;
        output.resize(amount, T::default());
    }
    Ok(output)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn options(input: &[u8]) -> RewriteOptions {
        RewriteOptions::new(
            input.len(),
            input.len().saturating_add(1_000_000),
            2_000_000,
            100_000_000,
            16,
            100_000,
            100_000,
            100_000,
        )
    }

    fn bytes_field(output: &mut Vec<u8>, number: u32, value: &[u8]) {
        put_key(output, number, 2);
        put_varint(output, value.len() as u64);
        output.extend_from_slice(value);
    }

    fn uuid_entry(object: u64, uuid: UuidBits) -> Vec<u8> {
        let mut nested = Vec::new();
        put_varint_field(&mut nested, 1, uuid.lower());
        put_varint_field(&mut nested, 2, uuid.upper());
        let mut entry = Vec::new();
        put_varint_field(&mut entry, 1, object);
        bytes_field(&mut entry, 2, &nested);
        entry
    }

    fn external_reference(target: u64, object: Option<u64>, weak: Option<u64>) -> Vec<u8> {
        let mut reference = Vec::new();
        put_varint_field(&mut reference, 1, target);
        if let Some(object) = object {
            put_varint_field(&mut reference, 2, object);
        }
        if let Some(weak) = weak {
            put_varint_field(&mut reference, 3, weak);
        }
        reference
    }

    fn data_reference(data: u64, owners: &[(u64, u32)], unknown: bool) -> Vec<u8> {
        let mut reference = Vec::new();
        if unknown {
            put_varint_field(&mut reference, 30, 99);
        }
        put_varint_field(&mut reference, 1, data);
        for (object, count) in owners {
            let mut owner = Vec::new();
            put_varint_field(&mut owner, 1, *object);
            put_varint_field(&mut owner, 2, u64::from(*count));
            bytes_field(&mut reference, 2, &owner);
        }
        reference
    }

    fn component(
        identifier: u64,
        preferred: &str,
        locator: Option<&str>,
        uuids: &[(u64, UuidBits)],
        references: &[(u32, u64, Option<u64>, Option<u64>)],
    ) -> Vec<u8> {
        let mut component = Vec::new();
        put_varint_field(&mut component, 1, identifier);
        bytes_field(&mut component, 2, preferred.as_bytes());
        if let Some(locator) = locator {
            bytes_field(&mut component, 3, locator.as_bytes());
        }
        for (object, uuid) in uuids {
            bytes_field(&mut component, 11, &uuid_entry(*object, *uuid));
        }
        for (field, target, object, weak) in references {
            bytes_field(
                &mut component,
                *field,
                &external_reference(*target, *object, *weak),
            );
        }
        component
    }

    fn metadata(last: u64, current: &[Vec<u8>], versioned: &[Vec<u8>]) -> Vec<u8> {
        let mut source = Vec::new();
        put_varint_field(&mut source, 50, 7);
        for component in current {
            bytes_field(&mut source, 3, component);
        }
        put_varint_field(&mut source, 1, last);
        for component in versioned {
            bytes_field(&mut source, 11, component);
        }
        put_key(&mut source, 51, 3);
        put_varint_field(&mut source, 1, 0);
        put_key(&mut source, 51, 4);
        source
    }

    #[derive(Default)]
    struct Facts {
        components: Vec<(u64, String, String, bool)>,
        uuids: Vec<(u64, u64, UuidBits, bool)>,
        references: Vec<(u64, u64, Option<u64>, Option<bool>, bool)>,
    }

    impl PackageMetadataVisitor for Facts {
        fn visit_component(
            &mut self,
            component: ComponentDescriptor<'_>,
        ) -> Result<(), RewriteError> {
            self.components.push((
                component.identifier(),
                component.preferred_locator().to_owned(),
                component.effective_locator().to_owned(),
                component.is_current(),
            ));
            Ok(())
        }

        fn visit_object_uuid(
            &mut self,
            binding: ObjectUuidDescriptor<'_>,
        ) -> Result<(), RewriteError> {
            self.uuids.push((
                binding.component().identifier(),
                binding.object_identifier(),
                binding.uuid(),
                binding.component().is_current(),
            ));
            Ok(())
        }

        fn visit_external_reference(
            &mut self,
            reference: ExternalReferenceDescriptor<'_>,
        ) -> Result<(), RewriteError> {
            self.references.push((
                reference.source().identifier(),
                reference.target_component_identifier(),
                reference.object_identifier(),
                reference.is_weak(),
                reference.is_versioned(),
            ));
            Ok(())
        }
    }

    #[test]
    fn inspection_streams_current_and_versioned_collision_facts() {
        let current = component(
            1,
            "preferred-a",
            Some("effective-a"),
            &[(4, UuidBits::new(1, 2))],
            &[(6, 2, Some(5), Some(0)), (18, 2, Some(6), Some(1))],
        );
        let versioned = component(9, "versioned", None, &[(3, UuidBits::new(7, 8))], &[]);
        let source = metadata(10, &[current], &[versioned]);
        let mut facts = Facts::default();
        let inspection =
            inspect_package_metadata_with_visitor(&source, options(&source), &mut facts).unwrap();
        assert_eq!(inspection.last_object_identifier(), 10);
        assert_eq!(inspection.report().input_bytes(), source.len());
        assert_eq!(inspection.report().components_scanned(), 4);
        assert_eq!(inspection.report().references_scanned(), 8);
        assert_eq!(
            facts.components,
            vec![
                (1, "preferred-a".into(), "effective-a".into(), true),
                (9, "versioned".into(), "versioned".into(), false)
            ]
        );
        assert_eq!(
            facts.uuids,
            vec![
                (1, 4, UuidBits::new(1, 2), true),
                (9, 3, UuidBits::new(7, 8), false)
            ]
        );
        assert_eq!(
            facts.references,
            vec![
                (1, 2, Some(5), Some(false), false),
                (1, 2, Some(6), Some(true), true)
            ]
        );
    }

    #[test]
    fn rewrite_is_raw_preserving_and_verified() {
        let a = component(1, "a-old", Some("a.iwa"), &[], &[]);
        let b = component(2, "b.iwa", None, &[], &[]);
        let versioned = component(8, "old.iwa", None, &[(7, UuidBits::new(3, 4))], &[]);
        let source = metadata(10, &[b, a.clone()], core::slice::from_ref(&versioned));
        let a_selector = ComponentSelector::new(1, "a.iwa");
        let b_selector = ComponentSelector::new(2, "b.iwa");
        let uuids = [ObjectUuidAddition::new(
            a_selector,
            11,
            UuidBits::new(100, 200),
        )];
        let references = [ExternalReferenceAddition::new(
            a_selector,
            b_selector,
            11,
            Some(false),
        )];
        let output = rewrite_package_metadata(
            &source,
            Batch::new(10, 11, &uuids, &references),
            options(&source),
        )
        .unwrap();
        assert_eq!(output.report().input_bytes(), source.len());
        assert_eq!(output.report().output_bytes(), output.bytes().len());
        assert_eq!(output.report().components_changed(), 1);
        assert_eq!(output.report().additions(), 2);
        assert_eq!(output.report().retained_bytes(), output.bytes().len());
        assert!(output.bytes().windows(a.len()).any(|window| window == a));
        assert!(
            output
                .bytes()
                .windows(versioned.len())
                .any(|window| window == versioned)
        );

        let mut facts = Facts::default();
        let inspection = inspect_package_metadata_with_visitor(
            output.bytes(),
            options(output.bytes()),
            &mut facts,
        )
        .unwrap();
        assert_eq!(inspection.last_object_identifier(), 11);
        assert!(
            facts
                .uuids
                .contains(&(1, 11, UuidBits::new(100, 200), true))
        );
        assert!(
            facts
                .references
                .contains(&(1, 2, Some(11), Some(false), false))
        );
    }

    fn reason(error: RewriteError) -> InvalidReason {
        error.invalid_reason().unwrap()
    }

    #[test]
    fn exact_selector_and_global_collision_rules_fail_closed() {
        let a = component(1, "a.iwa", None, &[], &[(6, 2, Some(9), Some(0))]);
        let b = component(2, "b.iwa", None, &[], &[]);
        let alias = component(3, "b.iwa", None, &[], &[]);
        let versioned = component(8, "old.iwa", None, &[(7, UuidBits::new(3, 4))], &[]);
        let source = metadata(10, &[a, b], &[versioned]);
        let source_selector = ComponentSelector::new(1, "a.iwa");
        let target_selector = ComponentSelector::new(2, "b.iwa");

        let uuid_collision = [ObjectUuidAddition::new(
            source_selector,
            11,
            UuidBits::new(3, 4),
        )];
        assert_eq!(
            reason(
                rewrite_package_metadata(
                    &source,
                    Batch::new(10, 11, &uuid_collision, &[]),
                    options(&source)
                )
                .unwrap_err()
            ),
            InvalidReason::ExistingUuidCollision
        );

        let same = [ExternalReferenceAddition::new(
            source_selector,
            target_selector,
            9,
            Some(false),
        )];
        assert_eq!(
            reason(
                rewrite_package_metadata(&source, Batch::new(10, 11, &[], &same), options(&source))
                    .unwrap_err()
            ),
            InvalidReason::ExistingReferenceCollision
        );
        let conflicting = [ExternalReferenceAddition::new(
            source_selector,
            target_selector,
            9,
            Some(true),
        )];
        assert_eq!(
            reason(
                rewrite_package_metadata(
                    &source,
                    Batch::new(10, 11, &[], &conflicting),
                    options(&source)
                )
                .unwrap_err()
            ),
            InvalidReason::ConflictingWeakness
        );

        let aliased_source = metadata(
            10,
            &[
                component(1, "a.iwa", None, &[], &[]),
                component(2, "b.iwa", None, &[], &[]),
                alias,
            ],
            &[],
        );
        assert_eq!(
            reason(
                rewrite_package_metadata(
                    &aliased_source,
                    Batch::new(10, 11, &[], &conflicting),
                    options(&aliased_source)
                )
                .unwrap_err()
            ),
            InvalidReason::ComponentMismatch
        );
    }

    #[test]
    fn malformed_and_hostile_records_are_rejected_atomically() {
        let mut noncanonical = metadata(10, &[], &[]);
        let last = noncanonical.iter().position(|byte| *byte == 0x08).unwrap();
        noncanonical.splice(last..=last + 1, [0x08, 0x8a, 0x00]);
        let mut facts = Facts::default();
        assert_eq!(
            reason(
                inspect_package_metadata_with_visitor(
                    &noncanonical,
                    options(&noncanonical),
                    &mut facts
                )
                .unwrap_err()
            ),
            InvalidReason::MalformedWire
        );

        let mut duplicate_last = metadata(10, &[], &[]);
        put_varint_field(&mut duplicate_last, 1, 10);
        assert_eq!(
            reason(
                inspect_package_metadata_with_visitor(
                    &duplicate_last,
                    options(&duplicate_last),
                    &mut Facts::default()
                )
                .unwrap_err()
            ),
            InvalidReason::MalformedWire
        );

        let bad_bool = component(1, "a.iwa", None, &[], &[(6, 2, Some(9), Some(2))]);
        let malformed = metadata(10, &[bad_bool], &[]);
        assert_eq!(
            reason(
                inspect_package_metadata_with_visitor(
                    &malformed,
                    options(&malformed),
                    &mut Facts::default()
                )
                .unwrap_err()
            ),
            InvalidReason::MalformedWire
        );
    }

    fn exact_options(source: &[u8], report: RewriteReport) -> RewriteOptions {
        RewriteOptions::new(
            source.len(),
            report.output_bytes(),
            report.fields(),
            report.work_bytes(),
            report.max_depth(),
            report.components_scanned(),
            report.references_scanned(),
            report.additions(),
        )
    }

    #[test]
    fn aggregate_limits_are_inclusive_and_max_minus_one_is_typed() {
        let source = metadata(
            10,
            &[
                component(1, "a.iwa", None, &[], &[]),
                component(2, "b.iwa", None, &[], &[]),
            ],
            &[],
        );
        let source_selector = ComponentSelector::new(1, "a.iwa");
        let target_selector = ComponentSelector::new(2, "b.iwa");
        let uuids = [ObjectUuidAddition::new(
            source_selector,
            11,
            UuidBits::new(5, 6),
        )];
        let references = [ExternalReferenceAddition::new(
            source_selector,
            target_selector,
            11,
            None,
        )];
        let batch = Batch::new(10, 11, &uuids, &references);
        let baseline = rewrite_package_metadata(&source, batch, options(&source)).unwrap();
        let report = baseline.report();
        assert_eq!(
            rewrite_package_metadata(&source, batch, exact_options(&source, report))
                .unwrap()
                .report(),
            report
        );

        let limits = [
            (
                RewriteOptions::new(
                    source.len() - 1,
                    report.output_bytes(),
                    report.fields(),
                    report.work_bytes(),
                    report.max_depth(),
                    report.components_scanned(),
                    report.references_scanned(),
                    report.additions(),
                ),
                "input",
            ),
            (
                RewriteOptions::new(
                    source.len(),
                    report.output_bytes() - 1,
                    report.fields(),
                    report.work_bytes(),
                    report.max_depth(),
                    report.components_scanned(),
                    report.references_scanned(),
                    report.additions(),
                ),
                "output",
            ),
            (
                RewriteOptions::new(
                    source.len(),
                    report.output_bytes(),
                    report.fields() - 1,
                    report.work_bytes(),
                    report.max_depth(),
                    report.components_scanned(),
                    report.references_scanned(),
                    report.additions(),
                ),
                "fields",
            ),
            (
                RewriteOptions::new(
                    source.len(),
                    report.output_bytes(),
                    report.fields(),
                    report.work_bytes() - 1,
                    report.max_depth(),
                    report.components_scanned(),
                    report.references_scanned(),
                    report.additions(),
                ),
                "work",
            ),
            (
                RewriteOptions::new(
                    source.len(),
                    report.output_bytes(),
                    report.fields(),
                    report.work_bytes(),
                    report.max_depth(),
                    report.components_scanned() - 1,
                    report.references_scanned(),
                    report.additions(),
                ),
                "components",
            ),
            (
                RewriteOptions::new(
                    source.len(),
                    report.output_bytes(),
                    report.fields(),
                    report.work_bytes(),
                    report.max_depth(),
                    report.components_scanned(),
                    report.references_scanned() - 1,
                    report.additions(),
                ),
                "references",
            ),
            (
                RewriteOptions::new(
                    source.len(),
                    report.output_bytes(),
                    report.fields(),
                    report.work_bytes(),
                    report.max_depth(),
                    report.components_scanned(),
                    report.references_scanned(),
                    report.additions() - 1,
                ),
                "additions",
            ),
        ];
        for (limited, label) in limits {
            let allocations_before = OUTPUT_ALLOCATIONS.load(Ordering::Relaxed);
            let error = rewrite_package_metadata(&source, batch, limited).unwrap_err();
            assert!(error.resource_limit().is_some(), "missing {label} limit");
            assert_eq!(
                OUTPUT_ALLOCATIONS.load(Ordering::Relaxed),
                allocations_before,
                "{label} limit reached the output allocation"
            );
        }
    }

    #[derive(Default)]
    struct CallbackCount(usize);

    impl PackageMetadataVisitor for CallbackCount {
        fn visit_component(
            &mut self,
            _component: ComponentDescriptor<'_>,
        ) -> Result<(), RewriteError> {
            self.0 += 1;
            Ok(())
        }

        fn visit_object_uuid(
            &mut self,
            _binding: ObjectUuidDescriptor<'_>,
        ) -> Result<(), RewriteError> {
            self.0 += 1;
            Ok(())
        }

        fn visit_external_reference(
            &mut self,
            _reference: ExternalReferenceDescriptor<'_>,
        ) -> Result<(), RewriteError> {
            self.0 += 1;
            Ok(())
        }
    }

    #[test]
    fn inspection_preflights_every_limit_before_callbacks() {
        let source = metadata(
            10,
            &[
                component(
                    1,
                    "a.iwa",
                    None,
                    &[(4, UuidBits::new(1, 2))],
                    &[(6, 2, Some(5), Some(0))],
                ),
                component(2, "b.iwa", None, &[], &[]),
            ],
            &[],
        );
        let mut baseline_visitor = CallbackCount::default();
        let report =
            inspect_package_metadata_with_visitor(&source, options(&source), &mut baseline_visitor)
                .unwrap()
                .report();
        let exact = RewriteOptions::new(
            source.len(),
            0,
            report.fields(),
            report.work_bytes(),
            report.max_depth(),
            report.components_scanned(),
            report.references_scanned(),
            0,
        );
        assert!(
            inspect_package_metadata_with_visitor(&source, exact, &mut CallbackCount::default())
                .is_ok()
        );
        let limits = [
            RewriteOptions::new(
                source.len() - 1,
                0,
                report.fields(),
                report.work_bytes(),
                report.max_depth(),
                report.components_scanned(),
                report.references_scanned(),
                0,
            ),
            RewriteOptions::new(
                source.len(),
                0,
                report.fields() - 1,
                report.work_bytes(),
                report.max_depth(),
                report.components_scanned(),
                report.references_scanned(),
                0,
            ),
            RewriteOptions::new(
                source.len(),
                0,
                report.fields(),
                report.work_bytes() - 1,
                report.max_depth(),
                report.components_scanned(),
                report.references_scanned(),
                0,
            ),
            RewriteOptions::new(
                source.len(),
                0,
                report.fields(),
                report.work_bytes(),
                report.max_depth(),
                report.components_scanned() - 1,
                report.references_scanned(),
                0,
            ),
            RewriteOptions::new(
                source.len(),
                0,
                report.fields(),
                report.work_bytes(),
                report.max_depth(),
                report.components_scanned(),
                report.references_scanned() - 1,
                0,
            ),
        ];
        for limited in limits {
            let mut visitor = CallbackCount::default();
            let error =
                inspect_package_metadata_with_visitor(&source, limited, &mut visitor).unwrap_err();
            assert!(error.resource_limit().is_some());
            assert_eq!(visitor.0, 0);
        }
    }

    fn many_components(count: usize) -> Vec<u8> {
        let mut source = Vec::new();
        put_varint_field(&mut source, 1, 20_000);
        for index in 1..=count {
            let locator = format!("component-{index}.iwa");
            bytes_field(
                &mut source,
                3,
                &component(index as u64, &locator, None, &[], &[]),
            );
        }
        source
    }

    #[test]
    fn four_thousand_to_eight_thousand_components_scale_linearly() {
        fn run(count: usize) -> (RewriteReport, RewriteReport) {
            let source = many_components(count);
            let locator = "component-1.iwa";
            let addition = [ObjectUuidAddition::new(
                ComponentSelector::new(1, locator),
                20_001,
                UuidBits::new(91, 92),
            )];
            let rewrite = rewrite_package_metadata(
                &source,
                Batch::new(20_000, 20_001, &addition, &[]),
                options(&source),
            )
            .unwrap()
            .report();
            let inspection = inspect_package_metadata_with_visitor(
                &source,
                options(&source),
                &mut Facts::default(),
            )
            .unwrap()
            .report();
            (rewrite, inspection)
        }
        let (four_rewrite, four_inspect) = run(4_096);
        let (eight_rewrite, eight_inspect) = run(8_192);
        for (four, eight) in [(four_rewrite, eight_rewrite), (four_inspect, eight_inspect)] {
            assert_eq!(eight.components_scanned(), four.components_scanned() * 2);
            assert!(eight.fields() * 100 <= four.fields() * 220);
            assert!(eight.work_bytes() * 100 <= four.work_bytes() * 220);
            assert!(eight.references_scanned() * 100 <= four.references_scanned() * 220);
            assert!(eight.allocations() * 100 <= four.allocations().max(1) * 220);
        }
    }

    #[test]
    fn removal_preserves_unrelated_raw_records_and_last_identifier() {
        let selector = ComponentSelector::new(1, "a.iwa");
        let target = ComponentSelector::new(2, "b.iwa");
        let selected_uuid = UuidBits::new(10, 20);
        let mut a = component(
            1,
            "a.iwa",
            None,
            &[(5, selected_uuid), (6, UuidBits::new(30, 40))],
            &[(6, 2, Some(5), Some(0)), (6, 2, Some(6), Some(1))],
        );
        let unrelated_ownerless = data_reference(70, &[], true);
        let selected_data = data_reference(71, &[(5, 2), (6, 3)], true);
        bytes_field(&mut a, 7, &unrelated_ownerless);
        bytes_field(&mut a, 7, &selected_data);
        put_key(&mut a, 53, 0);
        a.extend_from_slice(&[0x81, 0x00]);
        let noncanonical_unknown = [0xa8, 0x03, 0x81, 0x00];
        put_key(&mut a, 52, 3);
        put_varint_field(&mut a, 1, 0);
        put_key(&mut a, 52, 4);
        let b = component(2, "b.iwa", None, &[], &[]);
        let source = metadata(10, &[a, b.clone()], &[]);
        let uuids = [ObjectUuidRemoval::new(selector, 5, selected_uuid)];
        let externals = [ExternalReferenceRemoval::new(
            selector,
            target,
            5,
            Some(false),
        )];
        let owners = [DataReferenceOwnerRemoval::new(selector, 71, 5, 2)];
        let output = remove_package_metadata(
            &source,
            RemovalBatch::new(10, &uuids, &externals, &owners),
            options(&source),
        )
        .unwrap();
        assert_eq!(output.report().removals(), 3);
        assert_eq!(output.report().additions(), 0);
        assert!(
            output
                .bytes()
                .windows(unrelated_ownerless.len())
                .any(|window| window == unrelated_ownerless)
        );
        assert!(output.bytes().windows(b.len()).any(|window| window == b));
        assert!(
            output
                .bytes()
                .windows(noncanonical_unknown.len())
                .any(|window| window == noncanonical_unknown)
        );
        let inspection = inspect_package_metadata_with_visitor(
            output.bytes(),
            options(output.bytes()),
            &mut Facts::default(),
        )
        .unwrap();
        assert_eq!(inspection.last_object_identifier(), 10);
    }

    #[test]
    fn removal_rejects_versioned_ambiguous_and_cross_kind_occurrences() {
        let selector = ComponentSelector::new(1, "a.iwa");
        let uuid = UuidBits::new(10, 20);
        let removals = [ObjectUuidRemoval::new(selector, 5, uuid)];
        let batch = RemovalBatch::new(10, &removals, &[], &[]);
        let current = component(1, "a.iwa", None, &[(5, uuid)], &[]);

        let versioned = component(9, "old.iwa", None, &[(5, uuid)], &[]);
        let source = metadata(10, core::slice::from_ref(&current), &[versioned]);
        assert_eq!(
            reason(remove_package_metadata(&source, batch, options(&source)).unwrap_err()),
            InvalidReason::VersionedRemoval
        );

        let mut ambiguous = current.clone();
        put_varint_field(&mut ambiguous, 20, 5);
        let source = metadata(10, &[ambiguous], &[]);
        assert_eq!(
            reason(remove_package_metadata(&source, batch, options(&source)).unwrap_err()),
            InvalidReason::CrossComponentRemoval
        );

        let hostile_external = component(1, "a.iwa", None, &[(5, uuid)], &[(6, 2, Some(5), None)]);
        let source = metadata(10, &[hostile_external], &[]);
        assert_eq!(
            reason(remove_package_metadata(&source, batch, options(&source)).unwrap_err()),
            InvalidReason::CrossComponentRemoval
        );

        let mut hostile_owner = current;
        bytes_field(&mut hostile_owner, 7, &data_reference(7, &[(5, 1)], false));
        let source = metadata(10, &[hostile_owner], &[]);
        assert_eq!(
            reason(remove_package_metadata(&source, batch, options(&source)).unwrap_err()),
            InvalidReason::CrossComponentRemoval
        );
    }

    #[test]
    fn removal_drops_an_empty_selected_data_reference() {
        let selector = ComponentSelector::new(1, "a.iwa");
        let uuid = UuidBits::new(10, 20);
        let selected = data_reference(71, &[(5, 2)], true);
        let mut current = component(1, "a.iwa", None, &[(5, uuid)], &[]);
        bytes_field(&mut current, 7, &selected);
        let source = metadata(10, &[current], &[]);
        let uuids = [ObjectUuidRemoval::new(selector, 5, uuid)];
        let owners = [DataReferenceOwnerRemoval::new(selector, 71, 5, 2)];
        let output = remove_package_metadata(
            &source,
            RemovalBatch::new(10, &uuids, &[], &owners),
            options(&source),
        )
        .unwrap();
        assert!(
            !output
                .bytes()
                .windows(selected.len())
                .any(|window| window == selected)
        );
    }

    #[test]
    fn removal_output_limit_is_inclusive_and_max_minus_one_precedes_allocation() {
        let selector = ComponentSelector::new(1, "a.iwa");
        let uuid = UuidBits::new(10, 20);
        let current = component(1, "a.iwa", None, &[(5, uuid)], &[]);
        let source = metadata(10, &[current], &[]);
        let uuids = [ObjectUuidRemoval::new(selector, 5, uuid)];
        let batch = RemovalBatch::new(10, &uuids, &[], &[]);
        let baseline = remove_package_metadata(&source, batch, options(&source)).unwrap();
        let report = baseline.report();
        let exact = RewriteOptions::new(
            source.len(),
            report.output_bytes(),
            report.fields(),
            report.work_bytes(),
            report.max_depth(),
            report.components_scanned(),
            report.references_scanned(),
            report.removals(),
        );
        assert_eq!(
            remove_package_metadata(&source, batch, exact)
                .unwrap()
                .report(),
            report
        );
        let limited = RewriteOptions::new(
            source.len(),
            report.output_bytes() - 1,
            report.fields(),
            report.work_bytes(),
            report.max_depth(),
            report.components_scanned(),
            report.references_scanned(),
            report.removals(),
        );
        let allocations = OUTPUT_ALLOCATIONS.load(Ordering::Relaxed);
        let error = remove_package_metadata(&source, batch, limited).unwrap_err();
        assert!(matches!(
            error.resource_limit(),
            Some(RewriteLimit::OutputBytes { .. })
        ));
        assert_eq!(OUTPUT_ALLOCATIONS.load(Ordering::Relaxed), allocations);

        let work_limited = RewriteOptions::new(
            source.len(),
            report.output_bytes(),
            report.fields(),
            report.work_bytes() - 1,
            report.max_depth(),
            report.components_scanned(),
            report.references_scanned(),
            report.removals(),
        );
        let error = remove_package_metadata(&source, batch, work_limited).unwrap_err();
        assert!(matches!(
            error.resource_limit(),
            Some(RewriteLimit::Work { .. })
        ));
    }
}

/// Strictly inspect PackageMetadata without materializing generated messages.
pub fn inspect_package_metadata_with_visitor<V: PackageMetadataVisitor>(
    source: &[u8],
    options: RewriteOptions,
    visitor: &mut V,
) -> Result<PackageMetadataInspection, RewriteError> {
    let mut budget = Budget::new_inspection(source, options)?;
    let mut noop = NoopVisitor;
    let last = inspect_metadata_pass(source, options, &mut budget, &mut noop)?;
    budget.preflight_repeat_from_zero()?;
    let emitted_last = inspect_metadata_pass(source, options, &mut budget, visitor)?;
    if emitted_last != last {
        return Err(RewriteError::invalid(InvalidReason::Verification));
    }
    Ok(PackageMetadataInspection {
        last_object_identifier: last,
        report: budget.report(),
    })
}

struct NoopVisitor;

impl PackageMetadataVisitor for NoopVisitor {}

fn inspect_metadata_pass<V: PackageMetadataVisitor>(
    source: &[u8],
    options: RewriteOptions,
    budget: &mut Budget,
    visitor: &mut V,
) -> Result<u64, RewriteError> {
    budget.message(source, 1)?;
    let mut last = None;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, 1)? {
        match field.number {
            1 => set_once(&mut last, field.varint()?)?,
            3 | 11 => inspect_component(field.bytes()?, field.number == 3, budget, visitor, 2)?,
            _ => {},
        }
    }
    let last = last
        .filter(|value| *value != 0)
        .ok_or_else(|| RewriteError::invalid(InvalidReason::InvalidIdentifier))?;
    budget.message(source, 1)?;
    let view: projection::PackageMetadataArchiveLazyView<'_> = options
        .buffa()
        .decode_lazy_view(source)
        .map_err(|_error| RewriteError::invalid(InvalidReason::MalformedWire))?;
    if !view.has_last_object_identifier() || view.last_object_identifier != last {
        return Err(RewriteError::invalid(InvalidReason::MalformedWire));
    }
    Ok(last)
}

fn inspect_component<V: PackageMetadataVisitor>(
    source: &[u8],
    current: bool,
    budget: &mut Budget,
    visitor: &mut V,
    depth: u32,
) -> Result<(), RewriteError> {
    budget.component()?;
    budget.message(source, depth)?;
    let child_depth = depth
        .checked_add(1)
        .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?;
    let mut identifier = None;
    let mut preferred_locator = None;
    let mut locator = None;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        match field.number {
            1 => set_once(&mut identifier, field.varint()?)?,
            2 => set_once(&mut preferred_locator, strict_utf8(field.bytes()?)?)?,
            3 => set_once(&mut locator, strict_utf8(field.bytes()?)?)?,
            _ => {},
        }
    }
    let descriptor = ComponentDescriptor {
        identifier: identifier
            .filter(|value| *value != 0)
            .ok_or_else(|| RewriteError::invalid(InvalidReason::InvalidIdentifier))?,
        preferred_locator: preferred_locator
            .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?,
        locator,
        current,
    };
    budget.message(source, depth)?;
    let view: projection::ComponentInfoArchiveLazyView<'_> = budget
        .options
        .buffa()
        .decode_lazy_view(source)
        .map_err(|_error| RewriteError::invalid(InvalidReason::MalformedWire))?;
    if !view.has_identifier()
        || !view.has_preferred_locator()
        || view.identifier != descriptor.identifier
        || view.preferred_locator != descriptor.preferred_locator
        || view.locator != descriptor.locator
    {
        return Err(RewriteError::invalid(InvalidReason::MalformedWire));
    }
    visitor.visit_component(descriptor)?;

    budget.message(source, depth)?;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        match field.number {
            6 | 18 => {
                let reference = decode_external_reference(field.bytes()?, budget, child_depth)?;
                visitor.visit_external_reference(ExternalReferenceDescriptor {
                    source: descriptor,
                    target_component_identifier: reference.target,
                    object_identifier: reference.object,
                    is_weak: reference.is_weak,
                    versioned: field.number == 18,
                })?;
            },
            11 => {
                let binding = decode_object_uuid(field.bytes()?, budget, child_depth)?;
                visitor.visit_object_uuid(ObjectUuidDescriptor {
                    component: descriptor,
                    object_identifier: binding.object,
                    uuid: binding.uuid,
                })?;
            },
            _ => {},
        }
    }
    Ok(())
}

/// Strictly rewrite one PackageMetadata payload and verify the complete result.
pub fn rewrite_package_metadata(
    source: &[u8],
    batch: Batch<'_>,
    options: RewriteOptions,
) -> Result<RewriteOutput, RewriteError> {
    validate_batch(batch, options)?;
    let mut budget = Budget::new(source, batch, options)?;

    let mut source_state = ScanState::new(batch, &mut budget)?;
    scan_metadata(
        source,
        batch,
        ScanMode::Source,
        &mut source_state,
        &mut budget,
        true,
    )?;
    source_state.validate_selectors()?;

    let output_size = exact_output_size(source, batch, &mut budget)?;
    budget.output_size(output_size)?;
    precharge_rewrite_and_verification(source, batch, output_size, &mut budget)?;
    let mut candidate = Vec::new();
    #[cfg(test)]
    OUTPUT_ALLOCATIONS.fetch_add(1, Ordering::Relaxed);
    candidate
        .try_reserve_exact(output_size)
        .map_err(|_error| RewriteError::allocation(output_size))?;
    budget.allocation(0)?;
    rewrite_into(source, batch, &mut candidate, &mut budget)?;
    if candidate.len() != output_size {
        return Err(RewriteError::invalid(InvalidReason::Verification));
    }

    budget.source_phase = false;
    let mut verified_state = ScanState::new(batch, &mut budget)?;
    scan_metadata(
        &candidate,
        batch,
        ScanMode::Verification,
        &mut verified_state,
        &mut budget,
        false,
    )?;
    verified_state.validate_verification()?;
    budget.output_bytes = candidate.len();
    budget.retained_bytes = candidate.len();
    let report = budget.report();
    Ok(RewriteOutput {
        bytes: candidate,
        report,
    })
}

#[derive(Default, Clone, Copy)]
struct RemovalMatchCount {
    current: usize,
}

struct RemovalScanState {
    selectors: Vec<SelectorCount>,
    objects: Vec<RemovalMatchCount>,
    externals: Vec<RemovalMatchCount>,
    data_owners: Vec<RemovalMatchCount>,
}

impl RemovalScanState {
    fn new(batch: RemovalBatch<'_>, budget: &mut Budget) -> Result<Self, RewriteError> {
        let selector_count = batch
            .object_uuids
            .len()
            .checked_add(
                batch
                    .external_references
                    .len()
                    .checked_mul(2)
                    .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?,
            )
            .and_then(|count| count.checked_add(batch.data_reference_owners.len()))
            .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?;
        Ok(Self {
            selectors: zeroed_vec(selector_count, budget)?,
            objects: zeroed_vec(batch.object_uuids.len(), budget)?,
            externals: zeroed_vec(batch.external_references.len(), budget)?,
            data_owners: zeroed_vec(batch.data_reference_owners.len(), budget)?,
        })
    }

    fn validate_source(&self) -> Result<(), RewriteError> {
        if self
            .selectors
            .iter()
            .any(|count| count.identifier != 1 || count.locator != 1 || count.exact != 1)
        {
            return Err(RewriteError::invalid(InvalidReason::ComponentMismatch));
        }
        if self
            .objects
            .iter()
            .chain(self.externals.iter())
            .chain(self.data_owners.iter())
            .any(|count| count.current == 0)
        {
            return Err(RewriteError::invalid(InvalidReason::RemovalNotFound));
        }
        if self
            .objects
            .iter()
            .chain(self.externals.iter())
            .chain(self.data_owners.iter())
            .any(|count| count.current != 1)
        {
            return Err(RewriteError::invalid(InvalidReason::DuplicateRemoval));
        }
        Ok(())
    }

    fn validate_candidate(&self) -> Result<(), RewriteError> {
        if self
            .selectors
            .iter()
            .any(|count| count.identifier != 1 || count.locator != 1 || count.exact != 1)
            || self
                .objects
                .iter()
                .chain(self.externals.iter())
                .chain(self.data_owners.iter())
                .any(|count| count.current != 0)
        {
            return Err(RewriteError::invalid(InvalidReason::Verification));
        }
        Ok(())
    }
}

/// Strictly remove exact current registry records while retaining the last identifier.
pub fn remove_package_metadata(
    source: &[u8],
    batch: RemovalBatch<'_>,
    options: RewriteOptions,
) -> Result<RewriteOutput, RewriteError> {
    validate_removal_batch(batch, options)?;
    let mut budget = Budget::new_inspection(source, options)?;
    budget.removals = batch
        .object_uuids
        .len()
        .checked_add(batch.external_references.len())
        .and_then(|count| count.checked_add(batch.data_reference_owners.len()))
        .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?;

    let mut source_state = RemovalScanState::new(batch, &mut budget)?;
    scan_removal_metadata(source, batch, &mut source_state, &mut budget, false)?;
    source_state.validate_source()?;

    let output_size = removal_output_size(source, batch, &mut budget)?;
    budget.output_size(output_size)?;

    // Charge the exact rewrite traversal before constructing the sole owned candidate.
    let measured = budget.clone();
    charge_removal_rewrite(source, batch, &mut budget)?;
    budget.preflight_repeat_delta(&measured)?;
    budget.source_phase = false;

    let mut candidate = Vec::new();
    #[cfg(test)]
    OUTPUT_ALLOCATIONS.fetch_add(1, Ordering::Relaxed);
    candidate
        .try_reserve_exact(output_size)
        .map_err(|_error| RewriteError::allocation(output_size))?;
    budget.allocation(0)?;
    rewrite_removals_into(source, batch, &mut candidate, &mut budget)?;
    if candidate.len() != output_size {
        return Err(RewriteError::invalid(InvalidReason::Verification));
    }

    let mut verified = RemovalScanState::new(batch, &mut budget)?;
    scan_removal_metadata(&candidate, batch, &mut verified, &mut budget, true)?;
    verified.validate_candidate()?;
    budget.output_bytes = candidate.len();
    budget.retained_bytes = candidate.len();
    Ok(RewriteOutput {
        bytes: candidate,
        report: budget.report(),
    })
}

fn validate_removal_batch(
    batch: RemovalBatch<'_>,
    options: RewriteOptions,
) -> Result<(), RewriteError> {
    let removals = batch
        .object_uuids
        .len()
        .checked_add(batch.external_references.len())
        .and_then(|count| count.checked_add(batch.data_reference_owners.len()))
        .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?;
    if removals == 0 {
        return Err(RewriteError::invalid(InvalidReason::RemovalNotFound));
    }
    if removals > options.max_additions {
        return Err(RewriteError::limited(RewriteLimit::Additions {
            observed: removals,
            maximum: options.max_additions,
        }));
    }
    if batch.expected_last_object_identifier == 0 {
        return Err(RewriteError::invalid(InvalidReason::InvalidIdentifier));
    }
    for (index, removal) in batch.object_uuids.iter().enumerate() {
        validate_selector(removal.component)?;
        if removal.object_identifier == 0 || removal.expected_uuid == UuidBits::new(0, 0) {
            return Err(RewriteError::invalid(InvalidReason::InvalidIdentifier));
        }
        if batch.object_uuids[..index].iter().any(|prior| {
            prior.object_identifier == removal.object_identifier
                || prior.expected_uuid == removal.expected_uuid
        }) {
            return Err(RewriteError::invalid(InvalidReason::DuplicateRemoval));
        }
    }
    for (index, removal) in batch.external_references.iter().enumerate() {
        validate_selector(removal.source)?;
        validate_selector(removal.target)?;
        if removal.object_identifier == 0 {
            return Err(RewriteError::invalid(InvalidReason::InvalidIdentifier));
        }
        if batch.external_references[..index].iter().any(|prior| {
            prior.source == removal.source
                && prior.target == removal.target
                && prior.object_identifier == removal.object_identifier
        }) {
            return Err(RewriteError::invalid(InvalidReason::DuplicateRemoval));
        }
    }
    for (index, removal) in batch.data_reference_owners.iter().enumerate() {
        validate_selector(removal.component)?;
        if removal.data_identifier == 0
            || removal.object_identifier == 0
            || removal.expected_count == 0
        {
            return Err(RewriteError::invalid(InvalidReason::InvalidIdentifier));
        }
        if batch.data_reference_owners[..index].iter().any(|prior| {
            prior.component == removal.component
                && prior.data_identifier == removal.data_identifier
                && prior.object_identifier == removal.object_identifier
        }) {
            return Err(RewriteError::invalid(InvalidReason::DuplicateRemoval));
        }
    }
    Ok(())
}

fn removal_selector_count(batch: RemovalBatch<'_>) -> usize {
    batch.object_uuids.len()
        + batch.external_references.len() * 2
        + batch.data_reference_owners.len()
}

fn removal_selector_at<'source>(
    batch: RemovalBatch<'source>,
    index: usize,
) -> ComponentSelector<'source> {
    if index < batch.object_uuids.len() {
        return batch.object_uuids[index].component;
    }
    let shifted = index - batch.object_uuids.len();
    let external_selectors = batch.external_references.len() * 2;
    if shifted >= external_selectors {
        return batch.data_reference_owners[shifted - external_selectors].component;
    }
    let removal = batch.external_references[shifted / 2];
    if shifted % 2 == 0 {
        removal.source
    } else {
        removal.target
    }
}

fn scan_removal_metadata(
    source: &[u8],
    batch: RemovalBatch<'_>,
    state: &mut RemovalScanState,
    budget: &mut Budget,
    candidate: bool,
) -> Result<(), RewriteError> {
    budget.message(source, 1)?;
    let mut last = None;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, 1)? {
        match field.number {
            1 => set_once(&mut last, field.varint()?)?,
            3 | 11 => scan_removal_component(
                field.bytes()?,
                field.number == 3,
                batch,
                state,
                budget,
                candidate,
                2,
            )?,
            _ => {},
        }
    }
    if last != Some(batch.expected_last_object_identifier) {
        return Err(RewriteError::invalid(if candidate {
            InvalidReason::Verification
        } else {
            InvalidReason::LastIdentifierMismatch
        }));
    }
    budget.message(source, 1)?;
    let view: projection::PackageMetadataArchiveLazyView<'_> = budget
        .options
        .buffa()
        .decode_lazy_view(source)
        .map_err(|_error| RewriteError::invalid(InvalidReason::MalformedWire))?;
    if !view.has_last_object_identifier()
        || view.last_object_identifier != batch.expected_last_object_identifier
    {
        return Err(RewriteError::invalid(InvalidReason::MalformedWire));
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn scan_removal_component(
    source: &[u8],
    current: bool,
    batch: RemovalBatch<'_>,
    state: &mut RemovalScanState,
    budget: &mut Budget,
    candidate: bool,
    depth: u32,
) -> Result<(), RewriteError> {
    budget.component()?;
    let (identifier, locator) = component_header(source, budget, depth)?;
    budget.message(source, depth)?;
    if current {
        for index in 0..removal_selector_count(batch) {
            budget.work(1)?;
            let selector = removal_selector_at(batch, index);
            let count = &mut state.selectors[index];
            if identifier == selector.identifier {
                count.identifier = checked_add(count.identifier, 1)?;
            }
            if locator == selector.locator {
                count.locator = checked_add(count.locator, 1)?;
            }
            if identifier == selector.identifier && locator == selector.locator {
                count.exact = checked_add(count.exact, 1)?;
            }
        }
    }
    let child_depth = depth
        .checked_add(1)
        .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        match field.number {
            6 | 18 => {
                let reference = decode_external_reference(field.bytes()?, budget, child_depth)?;
                let deleted_object = reference.object.is_some_and(|object| {
                    batch
                        .object_uuids
                        .iter()
                        .any(|removal| removal.object_identifier == object)
                });
                let mut authorized = false;
                for (index, removal) in batch.external_references.iter().enumerate() {
                    budget.work(1)?;
                    if reference.object != Some(removal.object_identifier) {
                        continue;
                    }
                    let full = current
                        && field.number == 6
                        && identifier == removal.source.identifier
                        && locator == removal.source.locator
                        && reference.target == removal.target.identifier;
                    if !full {
                        return Err(RewriteError::invalid(if !current || field.number == 18 {
                            InvalidReason::VersionedRemoval
                        } else {
                            InvalidReason::CrossComponentRemoval
                        }));
                    }
                    if reference.is_weak != removal.expected_is_weak {
                        return Err(RewriteError::invalid(InvalidReason::RemovalMismatch));
                    }
                    authorized = true;
                    state.externals[index].current =
                        checked_add(state.externals[index].current, 1)?;
                }
                if deleted_object && !authorized {
                    return Err(RewriteError::invalid(if !current || field.number == 18 {
                        InvalidReason::VersionedRemoval
                    } else {
                        InvalidReason::CrossComponentRemoval
                    }));
                }
            },
            7 => scan_data_reference_removals(
                field.bytes()?,
                identifier,
                locator,
                current,
                batch,
                state,
                budget,
                child_depth,
            )?,
            11 => {
                let entry = decode_object_uuid(field.bytes()?, budget, child_depth)?;
                for (index, removal) in batch.object_uuids.iter().enumerate() {
                    budget.work(1)?;
                    let selected = entry.object == removal.object_identifier
                        || entry.uuid == removal.expected_uuid;
                    if !selected {
                        continue;
                    }
                    let full = current
                        && identifier == removal.component.identifier
                        && locator == removal.component.locator
                        && entry.object == removal.object_identifier
                        && entry.uuid == removal.expected_uuid;
                    if !full {
                        return Err(RewriteError::invalid(if !current {
                            InvalidReason::VersionedRemoval
                        } else if entry.object == removal.object_identifier
                            || entry.uuid == removal.expected_uuid
                        {
                            InvalidReason::CrossComponentRemoval
                        } else {
                            InvalidReason::RemovalMismatch
                        }));
                    }
                    state.objects[index].current = checked_add(state.objects[index].current, 1)?;
                }
            },
            20 => scan_ambiguous_ids(field, batch, budget)?,
            _ => {},
        }
    }
    if candidate {
        // Candidate scans use the same global collision rules; exact selected
        // records must simply no longer occur.
    }
    Ok(())
}

fn scan_data_reference_removals(
    source: &[u8],
    component: u64,
    locator: &str,
    current: bool,
    batch: RemovalBatch<'_>,
    state: &mut RemovalScanState,
    budget: &mut Budget,
    depth: u32,
) -> Result<(), RewriteError> {
    budget.reference()?;
    budget.message(source, depth)?;
    let mut data_identifier = None;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        if field.number == 1 {
            set_once(&mut data_identifier, field.varint()?)?;
        }
    }
    let data_identifier = data_identifier
        .filter(|value| *value != 0)
        .ok_or_else(|| RewriteError::invalid(InvalidReason::InvalidIdentifier))?;
    let child_depth = depth
        .checked_add(1)
        .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        if field.number != 2 {
            continue;
        }
        let (object, count) = decode_data_owner(field.bytes()?, budget, child_depth)?;
        let deleted_object = batch
            .object_uuids
            .iter()
            .any(|removal| removal.object_identifier == object);
        let mut authorized = false;
        for (index, removal) in batch.data_reference_owners.iter().enumerate() {
            budget.work(1)?;
            if object != removal.object_identifier {
                continue;
            }
            let full = current
                && component == removal.component.identifier
                && locator == removal.component.locator
                && data_identifier == removal.data_identifier;
            if !full {
                return Err(RewriteError::invalid(if !current {
                    InvalidReason::VersionedRemoval
                } else {
                    InvalidReason::CrossComponentRemoval
                }));
            }
            if count != removal.expected_count {
                return Err(RewriteError::invalid(InvalidReason::RemovalMismatch));
            }
            authorized = true;
            state.data_owners[index].current = checked_add(state.data_owners[index].current, 1)?;
        }
        if deleted_object && !authorized {
            return Err(RewriteError::invalid(if !current {
                InvalidReason::VersionedRemoval
            } else {
                InvalidReason::CrossComponentRemoval
            }));
        }
    }
    Ok(())
}

fn decode_data_owner(
    source: &[u8],
    budget: &mut Budget,
    depth: u32,
) -> Result<(u64, u32), RewriteError> {
    budget.reference()?;
    budget.message(source, depth)?;
    let mut object = None;
    let mut count = None;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        match field.number {
            1 => set_once(&mut object, field.varint()?)?,
            2 => set_once(
                &mut count,
                u32::try_from(field.varint()?)
                    .map_err(|_error| RewriteError::invalid(InvalidReason::MalformedWire))?,
            )?,
            _ => {},
        }
    }
    Ok((
        object
            .filter(|value| *value != 0)
            .ok_or_else(|| RewriteError::invalid(InvalidReason::InvalidIdentifier))?,
        count
            .filter(|value| *value != 0)
            .ok_or_else(|| RewriteError::invalid(InvalidReason::InvalidIdentifier))?,
    ))
}

fn scan_ambiguous_ids(
    field: Field<'_>,
    batch: RemovalBatch<'_>,
    budget: &mut Budget,
) -> Result<(), RewriteError> {
    let matches = |identifier: u64| {
        batch
            .object_uuids
            .iter()
            .any(|removal| removal.object_identifier == identifier)
            || batch
                .external_references
                .iter()
                .any(|removal| removal.object_identifier == identifier)
            || batch
                .data_reference_owners
                .iter()
                .any(|removal| removal.object_identifier == identifier)
    };
    match field.wire {
        0 => {
            budget.reference()?;
            budget.work(1)?;
            if matches(field.varint()?) {
                Err(RewriteError::invalid(InvalidReason::CrossComponentRemoval))
            } else {
                Ok(())
            }
        },
        2 => {
            let mut packed = field.bytes()?;
            while !packed.is_empty() {
                budget.reference()?;
                budget.work(1)?;
                if matches(take_varint(&mut packed)?) {
                    return Err(RewriteError::invalid(InvalidReason::CrossComponentRemoval));
                }
            }
            Ok(())
        },
        _ => Err(RewriteError::invalid(InvalidReason::MalformedWire)),
    }
}

fn removal_output_size(
    source: &[u8],
    batch: RemovalBatch<'_>,
    budget: &mut Budget,
) -> Result<usize, RewriteError> {
    budget.message(source, 1)?;
    let mut size = 0usize;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, 1)? {
        if field.number != 3 {
            size = checked_add(size, field.raw.len())?;
            continue;
        }
        let payload = field.bytes()?;
        let (identifier, locator) = component_header(payload, budget, 2)?;
        let new_len = removal_component_size(payload, identifier, locator, batch, budget, 2)?;
        size = checked_add(
            size,
            if new_len == payload.len() {
                field.raw.len()
            } else {
                length_delimited_field_len(3, new_len)?
            },
        )?;
    }
    Ok(size)
}

fn removal_component_size(
    source: &[u8],
    component: u64,
    locator: &str,
    batch: RemovalBatch<'_>,
    budget: &mut Budget,
    depth: u32,
) -> Result<usize, RewriteError> {
    budget.message(source, depth)?;
    let mut size = 0usize;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        let keep = match field.number {
            6 => !external_field_selected(
                field.bytes()?,
                component,
                locator,
                batch,
                budget,
                depth + 1,
            )?,
            7 => {
                let rewrite = data_reference_rewrite(
                    field.bytes()?,
                    component,
                    locator,
                    batch,
                    budget,
                    depth + 1,
                )?;
                if rewrite.selected == 0 {
                    size = checked_add(size, field.raw.len())?;
                } else if rewrite.surviving_owners != 0 {
                    size = checked_add(size, length_delimited_field_len(7, rewrite.payload_size)?)?;
                }
                false
            },
            11 => !object_field_selected(
                field.bytes()?,
                component,
                locator,
                batch,
                budget,
                depth + 1,
            )?,
            _ => true,
        };
        if keep {
            size = checked_add(size, field.raw.len())?;
        }
    }
    Ok(size)
}

fn object_field_selected(
    source: &[u8],
    component: u64,
    locator: &str,
    batch: RemovalBatch<'_>,
    budget: &mut Budget,
    depth: u32,
) -> Result<bool, RewriteError> {
    let entry = decode_object_uuid(source, budget, depth)?;
    Ok(batch.object_uuids.iter().any(|removal| {
        removal.component.identifier == component
            && removal.component.locator == locator
            && removal.object_identifier == entry.object
            && removal.expected_uuid == entry.uuid
    }))
}

fn external_field_selected(
    source: &[u8],
    component: u64,
    locator: &str,
    batch: RemovalBatch<'_>,
    budget: &mut Budget,
    depth: u32,
) -> Result<bool, RewriteError> {
    let reference = decode_external_reference(source, budget, depth)?;
    Ok(batch.external_references.iter().any(|removal| {
        removal.source.identifier == component
            && removal.source.locator == locator
            && removal.target.identifier == reference.target
            && Some(removal.object_identifier) == reference.object
            && removal.expected_is_weak == reference.is_weak
    }))
}

#[derive(Clone, Copy)]
struct DataReferenceRewrite {
    payload_size: usize,
    selected: usize,
    surviving_owners: usize,
}

fn data_reference_rewrite(
    source: &[u8],
    component: u64,
    locator: &str,
    batch: RemovalBatch<'_>,
    budget: &mut Budget,
    depth: u32,
) -> Result<DataReferenceRewrite, RewriteError> {
    budget.message(source, depth)?;
    let mut data_identifier = None;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        if field.number == 1 {
            set_once(&mut data_identifier, field.varint()?)?;
        }
    }
    let data_identifier =
        data_identifier.ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?;
    let mut size = 0usize;
    let mut selected_count = 0usize;
    let mut surviving_owners = 0usize;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        if field.number != 2 {
            size = checked_add(size, field.raw.len())?;
            continue;
        }
        let (object, count) = decode_data_owner(field.bytes()?, budget, depth + 1)?;
        let selected = batch.data_reference_owners.iter().any(|removal| {
            removal.component.identifier == component
                && removal.component.locator == locator
                && removal.data_identifier == data_identifier
                && removal.object_identifier == object
                && removal.expected_count == count
        });
        if !selected {
            surviving_owners = checked_add(surviving_owners, 1)?;
            size = checked_add(size, field.raw.len())?;
        } else {
            selected_count = checked_add(selected_count, 1)?;
        }
    }
    Ok(DataReferenceRewrite {
        payload_size: size,
        selected: selected_count,
        surviving_owners,
    })
}

fn charge_removal_rewrite(
    source: &[u8],
    batch: RemovalBatch<'_>,
    budget: &mut Budget,
) -> Result<(), RewriteError> {
    budget.message(source, 1)?;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, 1)? {
        if field.number == 3 {
            let payload = field.bytes()?;
            let (component, locator) = component_header(payload, budget, 2)?;
            let _size = removal_component_size(payload, component, locator, batch, budget, 2)?;
        }
    }
    Ok(())
}

fn rewrite_removals_into(
    source: &[u8],
    batch: RemovalBatch<'_>,
    output: &mut Vec<u8>,
    budget: &mut Budget,
) -> Result<(), RewriteError> {
    budget.message(source, 1)?;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, 1)? {
        if field.number != 3 {
            output.extend_from_slice(field.raw);
            continue;
        }
        let payload = field.bytes()?;
        let (component, locator) = component_header(payload, budget, 2)?;
        let new_len = removal_component_size(payload, component, locator, batch, budget, 2)?;
        if new_len == payload.len() {
            output.extend_from_slice(field.raw);
            continue;
        }
        budget.changed_component()?;
        put_key(output, 3, 2);
        put_varint(
            output,
            u64::try_from(new_len)
                .map_err(|_error| RewriteError::invalid(InvalidReason::MalformedWire))?,
        );
        rewrite_removal_component(payload, component, locator, batch, output, budget, 2)?;
    }
    Ok(())
}

fn rewrite_removal_component(
    source: &[u8],
    component: u64,
    locator: &str,
    batch: RemovalBatch<'_>,
    output: &mut Vec<u8>,
    budget: &mut Budget,
    depth: u32,
) -> Result<(), RewriteError> {
    budget.message(source, depth)?;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        match field.number {
            6 if external_field_selected(
                field.bytes()?,
                component,
                locator,
                batch,
                budget,
                depth + 1,
            )? => {},
            11 if object_field_selected(
                field.bytes()?,
                component,
                locator,
                batch,
                budget,
                depth + 1,
            )? => {},
            7 => rewrite_data_reference_field(
                field,
                component,
                locator,
                batch,
                output,
                budget,
                depth + 1,
            )?,
            _ => output.extend_from_slice(field.raw),
        }
    }
    Ok(())
}

fn rewrite_data_reference_field(
    field: Field<'_>,
    component: u64,
    locator: &str,
    batch: RemovalBatch<'_>,
    output: &mut Vec<u8>,
    budget: &mut Budget,
    depth: u32,
) -> Result<(), RewriteError> {
    let source = field.bytes()?;
    let rewrite = data_reference_rewrite(source, component, locator, batch, budget, depth)?;
    if rewrite.selected == 0 {
        output.extend_from_slice(field.raw);
        return Ok(());
    }
    if rewrite.surviving_owners == 0 {
        return Ok(());
    }
    let mut data_identifier = None;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        if field.number == 1 {
            set_once(&mut data_identifier, field.varint()?)?;
        }
    }
    let data_identifier =
        data_identifier.ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?;
    put_key(output, 7, 2);
    put_varint(
        output,
        u64::try_from(rewrite.payload_size)
            .map_err(|_error| RewriteError::invalid(InvalidReason::MalformedWire))?,
    );
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        if field.number == 2 {
            let (object, count) = decode_data_owner(field.bytes()?, budget, depth + 1)?;
            if batch.data_reference_owners.iter().any(|removal| {
                removal.component.identifier == component
                    && removal.component.locator == locator
                    && removal.data_identifier == data_identifier
                    && removal.object_identifier == object
                    && removal.expected_count == count
            }) {
                continue;
            }
        }
        output.extend_from_slice(field.raw);
    }
    Ok(())
}

fn validate_batch(batch: Batch<'_>, options: RewriteOptions) -> Result<(), RewriteError> {
    let additions = batch
        .object_uuids
        .len()
        .checked_add(batch.external_references.len())
        .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?;
    if additions > options.max_additions {
        return Err(RewriteError::limited(RewriteLimit::Additions {
            observed: additions,
            maximum: options.max_additions,
        }));
    }
    if batch.expected_last_object_identifier == 0
        || batch.new_last_object_identifier <= batch.expected_last_object_identifier
    {
        return Err(RewriteError::invalid(
            InvalidReason::LastIdentifierNotIncreasing,
        ));
    }
    for (index, addition) in batch.object_uuids.iter().enumerate() {
        validate_selector(addition.component)?;
        if addition.object_identifier <= batch.expected_last_object_identifier
            || addition.object_identifier > batch.new_last_object_identifier
        {
            return Err(RewriteError::invalid(InvalidReason::InvalidIdentifier));
        }
        if addition.uuid == UuidBits::new(0, 0) {
            return Err(RewriteError::invalid(InvalidReason::InvalidUuid));
        }
        if batch.object_uuids[..index].iter().any(|prior| {
            prior.object_identifier == addition.object_identifier || prior.uuid == addition.uuid
        }) {
            return Err(RewriteError::invalid(InvalidReason::DuplicateAddition));
        }
    }
    for (index, addition) in batch.external_references.iter().enumerate() {
        validate_selector(addition.source)?;
        validate_selector(addition.target)?;
        if addition.object_identifier == 0
            || addition.object_identifier > batch.new_last_object_identifier
        {
            return Err(RewriteError::invalid(InvalidReason::InvalidIdentifier));
        }
        if batch.external_references[..index].iter().any(|prior| {
            prior.source == addition.source
                && prior.target.identifier == addition.target.identifier
                && prior.object_identifier == addition.object_identifier
        }) {
            return Err(RewriteError::invalid(InvalidReason::DuplicateAddition));
        }
    }
    Ok(())
}

fn validate_selector(selector: ComponentSelector<'_>) -> Result<(), RewriteError> {
    if selector.identifier == 0 || selector.locator.is_empty() {
        return Err(RewriteError::invalid(InvalidReason::InvalidIdentifier));
    }
    Ok(())
}

fn selector_count(batch: Batch<'_>) -> usize {
    batch.object_uuids.len() + batch.external_references.len() * 2
}

fn selector_at<'source>(batch: Batch<'source>, index: usize) -> ComponentSelector<'source> {
    if index < batch.object_uuids.len() {
        return batch.object_uuids[index].component;
    }
    let shifted = index - batch.object_uuids.len();
    let addition = batch.external_references[shifted / 2];
    if shifted % 2 == 0 {
        addition.source
    } else {
        addition.target
    }
}

fn scan_metadata(
    source: &[u8],
    batch: Batch<'_>,
    mode: ScanMode,
    state: &mut ScanState,
    budget: &mut Budget,
    require_expected_last: bool,
) -> Result<(), RewriteError> {
    budget.message(source, 1)?;
    let mut last = None;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, 1)? {
        match field.number {
            1 => set_once(&mut last, field.varint()?)?,
            3 | 11 => scan_component(
                field.bytes()?,
                field.number == 3,
                batch,
                mode,
                state,
                budget,
                2,
            )?,
            _ => {},
        }
    }
    let expected = if require_expected_last {
        batch.expected_last_object_identifier
    } else {
        batch.new_last_object_identifier
    };
    if last != Some(expected) {
        return Err(RewriteError::invalid(if require_expected_last {
            InvalidReason::LastIdentifierMismatch
        } else {
            InvalidReason::Verification
        }));
    }
    budget.message(source, 1)?;
    let view: projection::PackageMetadataArchiveLazyView<'_> = budget
        .options
        .buffa()
        .decode_lazy_view(source)
        .map_err(|_error| RewriteError::invalid(InvalidReason::MalformedWire))?;
    if !view.has_last_object_identifier() || view.last_object_identifier != expected {
        return Err(RewriteError::invalid(InvalidReason::MalformedWire));
    }
    Ok(())
}

fn scan_component(
    source: &[u8],
    current: bool,
    batch: Batch<'_>,
    mode: ScanMode,
    state: &mut ScanState,
    budget: &mut Budget,
    depth: u32,
) -> Result<(), RewriteError> {
    budget.component()?;
    budget.message(source, depth)?;
    let child_depth = depth
        .checked_add(1)
        .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?;
    let mut identifier = None;
    let mut preferred_locator = None;
    let mut locator = None;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        match field.number {
            1 => set_once(&mut identifier, field.varint()?)?,
            2 => set_once(&mut preferred_locator, strict_utf8(field.bytes()?)?)?,
            3 => set_once(&mut locator, strict_utf8(field.bytes()?)?)?,
            _ => {},
        }
    }
    let identifier = identifier
        .filter(|value| *value != 0)
        .ok_or_else(|| RewriteError::invalid(InvalidReason::InvalidIdentifier))?;
    let preferred_locator =
        preferred_locator.ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?;
    let effective_locator = locator.unwrap_or(preferred_locator);
    budget.message(source, depth)?;
    let view: projection::ComponentInfoArchiveLazyView<'_> = budget
        .options
        .buffa()
        .decode_lazy_view(source)
        .map_err(|_error| RewriteError::invalid(InvalidReason::MalformedWire))?;
    if !view.has_identifier()
        || !view.has_preferred_locator()
        || view.identifier != identifier
        || view.preferred_locator != preferred_locator
        || view.locator != locator
    {
        return Err(RewriteError::invalid(InvalidReason::MalformedWire));
    }

    if current {
        for index in 0..selector_count(batch) {
            budget.work(1)?;
            let selector = selector_at(batch, index);
            let count = &mut state.selectors[index];
            if identifier == selector.identifier {
                count.identifier = count
                    .identifier
                    .checked_add(1)
                    .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?;
            }
            if effective_locator == selector.locator {
                count.locator = count
                    .locator
                    .checked_add(1)
                    .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?;
            }
            if identifier == selector.identifier && effective_locator == selector.locator {
                count.exact = count
                    .exact
                    .checked_add(1)
                    .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?;
            }
        }
    }

    budget.message(source, depth)?;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        match field.number {
            6 | 18 => {
                let reference = decode_external_reference(field.bytes()?, budget, child_depth)?;
                scan_external_collision(
                    identifier,
                    effective_locator,
                    reference,
                    batch,
                    mode,
                    state,
                    budget,
                )?;
            },
            11 => {
                let entry = decode_object_uuid(field.bytes()?, budget, child_depth)?;
                scan_object_collision(
                    identifier,
                    effective_locator,
                    entry,
                    batch,
                    mode,
                    state,
                    budget,
                )?;
            },
            _ => {},
        }
    }
    Ok(())
}

#[derive(Clone, Copy)]
struct ExternalReference {
    target: u64,
    object: Option<u64>,
    is_weak: Option<bool>,
}

fn decode_external_reference(
    source: &[u8],
    budget: &mut Budget,
    depth: u32,
) -> Result<ExternalReference, RewriteError> {
    budget.reference()?;
    budget.message(source, depth)?;
    let mut target = None;
    let mut object = None;
    let mut is_weak = None;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        match field.number {
            1 => set_once(&mut target, field.varint()?)?,
            2 => set_once(&mut object, field.varint()?)?,
            3 => set_once(&mut is_weak, canonical_bool(field.varint()?)?)?,
            _ => {},
        }
    }
    let target = target
        .filter(|value| *value != 0)
        .ok_or_else(|| RewriteError::invalid(InvalidReason::InvalidIdentifier))?;
    if object == Some(0) {
        return Err(RewriteError::invalid(InvalidReason::InvalidIdentifier));
    }
    let snapshot = ExternalReference {
        target,
        object,
        is_weak,
    };
    budget.message(source, depth)?;
    let view: projection::ComponentExternalReferenceArchiveLazyView<'_> = budget
        .options
        .buffa()
        .decode_lazy_view(source)
        .map_err(|_error| RewriteError::invalid(InvalidReason::MalformedWire))?;
    if !view.has_component_identifier()
        || view.component_identifier != target
        || view.object_identifier != object
        || view.is_weak != is_weak
    {
        return Err(RewriteError::invalid(InvalidReason::MalformedWire));
    }
    Ok(snapshot)
}

#[derive(Clone, Copy)]
struct ObjectUuid {
    object: u64,
    uuid: UuidBits,
}

fn decode_object_uuid(
    source: &[u8],
    budget: &mut Budget,
    depth: u32,
) -> Result<ObjectUuid, RewriteError> {
    budget.reference()?;
    budget.message(source, depth)?;
    let child_depth = depth
        .checked_add(1)
        .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?;
    let mut object = None;
    let mut uuid_raw = None;
    let mut uuid = None;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        match field.number {
            1 => set_once(&mut object, field.varint()?)?,
            2 => {
                let raw = field.bytes()?;
                set_once(&mut uuid, decode_uuid(raw, budget, child_depth)?)?;
                set_once(&mut uuid_raw, raw)?;
            },
            _ => {},
        }
    }
    let object = object
        .filter(|value| *value != 0)
        .ok_or_else(|| RewriteError::invalid(InvalidReason::InvalidIdentifier))?;
    let uuid = uuid.ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?;
    if uuid == UuidBits::new(0, 0) {
        return Err(RewriteError::invalid(InvalidReason::InvalidUuid));
    }
    budget.message(source, depth)?;
    let view: projection::ObjectUUIDMapEntryArchiveLazyView<'_> = budget
        .options
        .buffa()
        .decode_lazy_view(source)
        .map_err(|_error| RewriteError::invalid(InvalidReason::MalformedWire))?;
    if !view.has_identifier()
        || !view.has_uuid()
        || view.identifier != object
        || view.uuid
            != uuid_raw.ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?
    {
        return Err(RewriteError::invalid(InvalidReason::MalformedWire));
    }
    Ok(ObjectUuid { object, uuid })
}

fn decode_uuid(source: &[u8], budget: &mut Budget, depth: u32) -> Result<UuidBits, RewriteError> {
    budget.message(source, depth)?;
    let mut lower = None;
    let mut upper = None;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        match field.number {
            1 => set_once(&mut lower, field.varint()?)?,
            2 => set_once(&mut upper, field.varint()?)?,
            _ => {},
        }
    }
    let snapshot = UuidBits::new(
        lower.ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?,
        upper.ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?,
    );
    budget.message(source, depth)?;
    let view: projection::UUIDArchiveLazyView<'_> = budget
        .options
        .buffa()
        .decode_lazy_view(source)
        .map_err(|_error| RewriteError::invalid(InvalidReason::MalformedWire))?;
    if !view.has_lower()
        || !view.has_upper()
        || view.lower != snapshot.lower
        || view.upper != snapshot.upper
    {
        return Err(RewriteError::invalid(InvalidReason::MalformedWire));
    }
    Ok(snapshot)
}

fn scan_object_collision(
    component: u64,
    locator: &str,
    entry: ObjectUuid,
    batch: Batch<'_>,
    mode: ScanMode,
    state: &mut ScanState,
    budget: &mut Budget,
) -> Result<(), RewriteError> {
    for (index, addition) in batch.object_uuids.iter().enumerate() {
        budget.work(1)?;
        let id_match = entry.object == addition.object_identifier;
        let uuid_match = entry.uuid == addition.uuid;
        if !id_match && !uuid_match {
            continue;
        }
        match mode {
            ScanMode::Source => {
                return Err(RewriteError::invalid(if id_match {
                    InvalidReason::ExistingObjectCollision
                } else {
                    InvalidReason::ExistingUuidCollision
                }));
            },
            ScanMode::Verification => {
                if id_match
                    && uuid_match
                    && component == addition.component.identifier
                    && locator == addition.component.locator
                {
                    state.object_matches[index] = state.object_matches[index]
                        .checked_add(1)
                        .ok_or_else(|| RewriteError::invalid(InvalidReason::Verification))?;
                } else {
                    return Err(RewriteError::invalid(InvalidReason::Verification));
                }
            },
        }
    }
    Ok(())
}

fn scan_external_collision(
    component: u64,
    locator: &str,
    reference: ExternalReference,
    batch: Batch<'_>,
    mode: ScanMode,
    state: &mut ScanState,
    budget: &mut Budget,
) -> Result<(), RewriteError> {
    for (index, addition) in batch.external_references.iter().enumerate() {
        budget.work(1)?;
        if component != addition.source.identifier
            || locator != addition.source.locator
            || reference.target != addition.target.identifier
            || reference.object != Some(addition.object_identifier)
        {
            continue;
        }
        if reference.is_weak != addition.is_weak {
            return Err(RewriteError::invalid(InvalidReason::ConflictingWeakness));
        }
        match mode {
            ScanMode::Source => {
                return Err(RewriteError::invalid(
                    InvalidReason::ExistingReferenceCollision,
                ));
            },
            ScanMode::Verification => {
                state.external_matches[index] = state.external_matches[index]
                    .checked_add(1)
                    .ok_or_else(|| RewriteError::invalid(InvalidReason::Verification))?
            },
        }
    }
    Ok(())
}

fn component_header<'source>(
    source: &'source [u8],
    budget: &mut Budget,
    depth: u32,
) -> Result<(u64, &'source str), RewriteError> {
    budget.message(source, depth)?;
    let mut identifier = None;
    let mut preferred = None;
    let mut locator = None;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        match field.number {
            1 => set_once(&mut identifier, field.varint()?)?,
            2 => set_once(&mut preferred, strict_utf8(field.bytes()?)?)?,
            3 => set_once(&mut locator, strict_utf8(field.bytes()?)?)?,
            _ => {},
        }
    }
    Ok((
        identifier
            .filter(|value| *value != 0)
            .ok_or_else(|| RewriteError::invalid(InvalidReason::InvalidIdentifier))?,
        locator.unwrap_or(
            preferred.ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?,
        ),
    ))
}

fn component_append_len(
    identifier: u64,
    locator: &str,
    batch: Batch<'_>,
) -> Result<usize, RewriteError> {
    let mut amount = 0usize;
    for addition in batch.object_uuids.iter().filter(|addition| {
        addition.component.identifier == identifier && addition.component.locator == locator
    }) {
        let payload = object_uuid_payload_len(*addition)?;
        amount = checked_add(amount, length_delimited_field_len(11, payload)?)?;
    }
    for addition in batch.external_references.iter().filter(|addition| {
        addition.source.identifier == identifier && addition.source.locator == locator
    }) {
        let payload = external_payload_len(*addition)?;
        amount = checked_add(amount, length_delimited_field_len(6, payload)?)?;
    }
    Ok(amount)
}

fn precharge_rewrite_and_verification(
    source: &[u8],
    batch: Batch<'_>,
    output_size: usize,
    budget: &mut Budget,
) -> Result<(), RewriteError> {
    let measured = budget.clone();

    // Rewrite traversal: one root pass and one header pass per current
    // component. This is charged before the output allocation begins.
    budget.message(source, 1)?;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, 1)? {
        if field.number != 3 {
            continue;
        }
        let payload = field.bytes()?;
        let (identifier, locator) = component_header(payload, budget, 2)?;
        let _append = component_append_len(identifier, locator, batch)?;
    }

    // Candidate verification has the same root/component field cardinality as
    // the source plus the caller's canonical appends. Its byte work is known
    // exactly from the sizing pass, so no speculative candidate is needed.
    budget.source_phase = false;
    budget.message_len(output_size, 1)?;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, 1)? {
        if matches!(field.number, 3 | 11) {
            precharge_candidate_component(field.bytes()?, field.number == 3, batch, budget, 2)?;
        }
    }
    budget.message_len(output_size, 1)?;
    budget.preflight_repeat_delta(&measured)?;
    Ok(())
}

fn precharge_candidate_component(
    source: &[u8],
    current: bool,
    batch: Batch<'_>,
    budget: &mut Budget,
    depth: u32,
) -> Result<(), RewriteError> {
    budget.component()?;
    let (identifier, locator) = raw_component_header(source, budget, depth)?;
    let object_additions = if current {
        batch
            .object_uuids
            .iter()
            .filter(|addition| {
                addition.component.identifier == identifier && addition.component.locator == locator
            })
            .count()
    } else {
        0
    };
    let external_additions = if current {
        batch
            .external_references
            .iter()
            .filter(|addition| {
                addition.source.identifier == identifier && addition.source.locator == locator
            })
            .count()
    } else {
        0
    };
    let append = component_append_len(identifier, locator, batch)?;
    let candidate_len = checked_add(source.len(), append)?;

    // Header pass and Buffa parity pass.
    budget.message_len(candidate_len, depth)?;
    charge_fields(
        object_additions
            .checked_add(external_additions)
            .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?,
        budget,
    )?;
    budget.message_len(candidate_len, depth)?;
    if current {
        budget.work(selector_count(batch))?;
    }

    // Registry pass, including exact nested and Buffa work for existing and
    // newly appended records.
    budget.message_len(candidate_len, depth)?;
    let child_depth = depth
        .checked_add(1)
        .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        match field.number {
            6 | 18 => {
                let _snapshot = decode_external_reference(field.bytes()?, budget, child_depth)?;
                budget.work(batch.external_references.len())?;
            },
            11 => {
                let _snapshot = decode_object_uuid(field.bytes()?, budget, child_depth)?;
                budget.work(batch.object_uuids.len())?;
            },
            _ => {},
        }
    }
    charge_fields(
        object_additions
            .checked_add(external_additions)
            .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?,
        budget,
    )?;
    for addition in batch.object_uuids.iter().filter(|addition| {
        current
            && addition.component.identifier == identifier
            && addition.component.locator == locator
    }) {
        precharge_object_uuid(*addition, budget, child_depth)?;
        budget.work(batch.object_uuids.len())?;
    }
    for addition in batch.external_references.iter().filter(|addition| {
        current && addition.source.identifier == identifier && addition.source.locator == locator
    }) {
        precharge_external(*addition, budget, child_depth)?;
        budget.work(batch.external_references.len())?;
    }
    Ok(())
}

fn raw_component_header<'source>(
    source: &'source [u8],
    budget: &mut Budget,
    depth: u32,
) -> Result<(u64, &'source str), RewriteError> {
    let mut identifier = None;
    let mut preferred = None;
    let mut locator = None;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        match field.number {
            1 => set_once(&mut identifier, field.varint()?)?,
            2 => set_once(&mut preferred, strict_utf8(field.bytes()?)?)?,
            3 => set_once(&mut locator, strict_utf8(field.bytes()?)?)?,
            _ => {},
        }
    }
    Ok((
        identifier.ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?,
        locator.unwrap_or(
            preferred.ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?,
        ),
    ))
}

fn charge_fields(amount: usize, budget: &mut Budget) -> Result<(), RewriteError> {
    for _ in 0..amount {
        budget.field()?;
    }
    Ok(())
}

fn precharge_object_uuid(
    addition: ObjectUuidAddition<'_>,
    budget: &mut Budget,
    depth: u32,
) -> Result<(), RewriteError> {
    budget.reference()?;
    let uuid_len = checked_add(
        varint_field_len(1, addition.uuid.lower),
        varint_field_len(2, addition.uuid.upper),
    )?;
    let entry_len = object_uuid_payload_len(addition)?;
    budget.message_len(entry_len, depth)?;
    charge_fields(2, budget)?;
    let uuid_depth = depth
        .checked_add(1)
        .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?;
    budget.message_len(uuid_len, uuid_depth)?;
    charge_fields(2, budget)?;
    budget.message_len(uuid_len, uuid_depth)?;
    budget.message_len(entry_len, depth)?;
    Ok(())
}

fn precharge_external(
    addition: ExternalReferenceAddition<'_>,
    budget: &mut Budget,
    depth: u32,
) -> Result<(), RewriteError> {
    budget.reference()?;
    let payload_len = external_payload_len(addition)?;
    budget.message_len(payload_len, depth)?;
    charge_fields(2 + usize::from(addition.is_weak.is_some()), budget)?;
    budget.message_len(payload_len, depth)?;
    Ok(())
}

fn exact_output_size(
    source: &[u8],
    batch: Batch<'_>,
    budget: &mut Budget,
) -> Result<usize, RewriteError> {
    budget.message(source, 1)?;
    let mut output = 0usize;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, 1)? {
        match field.number {
            1 => {
                output = checked_add(
                    output,
                    varint_field_len(1, batch.new_last_object_identifier),
                )?
            },
            3 => {
                let payload = field.bytes()?;
                let (identifier, locator) = component_header(payload, budget, 2)?;
                let append = component_append_len(identifier, locator, batch)?;
                let length = checked_add(payload.len(), append)?;
                output = checked_add(output, length_delimited_field_len(3, length)?)?;
            },
            _ => output = checked_add(output, field.raw.len())?,
        }
    }
    Ok(output)
}

fn rewrite_into(
    source: &[u8],
    batch: Batch<'_>,
    output: &mut Vec<u8>,
    budget: &mut Budget,
) -> Result<(), RewriteError> {
    budget.message(source, 1)?;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, 1)? {
        match field.number {
            1 => put_varint_field(output, 1, batch.new_last_object_identifier),
            3 => {
                let payload = field.bytes()?;
                let (identifier, locator) = component_header(payload, budget, 2)?;
                let append = component_append_len(identifier, locator, batch)?;
                if append == 0 {
                    output.extend_from_slice(field.raw);
                    continue;
                }
                budget.changed_component()?;
                put_key(output, 3, 2);
                put_varint(
                    output,
                    u64::try_from(checked_add(payload.len(), append)?)
                        .map_err(|_error| RewriteError::invalid(InvalidReason::MalformedWire))?,
                );
                output.extend_from_slice(payload);
                for addition in batch.object_uuids.iter().filter(|addition| {
                    addition.component.identifier == identifier
                        && addition.component.locator == locator
                }) {
                    append_object_uuid(output, *addition)?;
                }
                for addition in batch.external_references.iter().filter(|addition| {
                    addition.source.identifier == identifier && addition.source.locator == locator
                }) {
                    append_external(output, *addition)?;
                }
            },
            _ => output.extend_from_slice(field.raw),
        }
    }
    Ok(())
}

fn append_object_uuid(
    output: &mut Vec<u8>,
    addition: ObjectUuidAddition<'_>,
) -> Result<(), RewriteError> {
    let uuid_len = checked_add(
        varint_field_len(1, addition.uuid.lower),
        varint_field_len(2, addition.uuid.upper),
    )?;
    let payload_len = checked_add(
        varint_field_len(1, addition.object_identifier),
        length_delimited_field_len(2, uuid_len)?,
    )?;
    put_key(output, 11, 2);
    put_varint(
        output,
        u64::try_from(payload_len)
            .map_err(|_error| RewriteError::invalid(InvalidReason::MalformedWire))?,
    );
    put_varint_field(output, 1, addition.object_identifier);
    put_key(output, 2, 2);
    put_varint(
        output,
        u64::try_from(uuid_len)
            .map_err(|_error| RewriteError::invalid(InvalidReason::MalformedWire))?,
    );
    put_varint_field(output, 1, addition.uuid.lower);
    put_varint_field(output, 2, addition.uuid.upper);
    Ok(())
}

fn append_external(
    output: &mut Vec<u8>,
    addition: ExternalReferenceAddition<'_>,
) -> Result<(), RewriteError> {
    let payload_len = external_payload_len(addition)?;
    put_key(output, 6, 2);
    put_varint(
        output,
        u64::try_from(payload_len)
            .map_err(|_error| RewriteError::invalid(InvalidReason::MalformedWire))?,
    );
    put_varint_field(output, 1, addition.target.identifier);
    put_varint_field(output, 2, addition.object_identifier);
    if let Some(value) = addition.is_weak {
        put_varint_field(output, 3, u64::from(value));
    }
    Ok(())
}

fn object_uuid_payload_len(addition: ObjectUuidAddition<'_>) -> Result<usize, RewriteError> {
    let uuid_len = checked_add(
        varint_field_len(1, addition.uuid.lower),
        varint_field_len(2, addition.uuid.upper),
    )?;
    checked_add(
        varint_field_len(1, addition.object_identifier),
        length_delimited_field_len(2, uuid_len)?,
    )
}

fn external_payload_len(addition: ExternalReferenceAddition<'_>) -> Result<usize, RewriteError> {
    let mut length = checked_add(
        varint_field_len(1, addition.target.identifier),
        varint_field_len(2, addition.object_identifier),
    )?;
    if addition.is_weak.is_some() {
        length = checked_add(length, varint_field_len(3, 1))?;
    }
    Ok(length)
}

fn varint_field_len(number: u32, value: u64) -> usize {
    encoded_varint_len(u64::from(number) << 3) + encoded_varint_len(value)
}

fn length_delimited_field_len(number: u32, payload: usize) -> Result<usize, RewriteError> {
    checked_add(
        encoded_varint_len((u64::from(number) << 3) | 2),
        checked_add(
            encoded_varint_len(
                u64::try_from(payload)
                    .map_err(|_error| RewriteError::invalid(InvalidReason::MalformedWire))?,
            ),
            payload,
        )?,
    )
}

fn checked_add(left: usize, right: usize) -> Result<usize, RewriteError> {
    left.checked_add(right)
        .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))
}

fn put_varint_field(output: &mut Vec<u8>, number: u32, value: u64) {
    put_key(output, number, 0);
    put_varint(output, value);
}

fn put_key(output: &mut Vec<u8>, number: u32, wire: u8) {
    put_varint(output, (u64::from(number) << 3) | u64::from(wire));
}

fn put_varint(output: &mut Vec<u8>, mut value: u64) {
    loop {
        let mut byte = value.to_le_bytes()[0] & 0x7f;
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

#[derive(Clone)]
struct Budget {
    options: RewriteOptions,
    input_bytes: usize,
    output_bytes: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    components_scanned: usize,
    components_changed: usize,
    references_scanned: usize,
    source_references_scanned: usize,
    additions: usize,
    removals: usize,
    allocations: usize,
    retained_bytes: usize,
    scratch_bytes: usize,
    source_phase: bool,
}

impl Budget {
    fn new(source: &[u8], batch: Batch<'_>, options: RewriteOptions) -> Result<Self, RewriteError> {
        let mut budget = Self::new_inspection(source, options)?;
        budget.additions = batch
            .object_uuids
            .len()
            .checked_add(batch.external_references.len())
            .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?;
        Ok(budget)
    }

    fn new_inspection(source: &[u8], options: RewriteOptions) -> Result<Self, RewriteError> {
        let hard = usize::try_from(buffa::MAX_MESSAGE_BYTES)
            .map_err(|_error| RewriteError::invalid(InvalidReason::MalformedWire))?;
        if options.max_input_bytes > hard || options.max_output_bytes > hard {
            return Err(RewriteError::limited(RewriteLimit::OutputBytes {
                observed: options.max_output_bytes.max(options.max_input_bytes),
                maximum: hard,
            }));
        }
        if source.len() > options.max_input_bytes {
            return Err(RewriteError::limited(RewriteLimit::InputBytes {
                observed: source.len(),
                maximum: options.max_input_bytes,
            }));
        }
        if options.recursion_limit == 0 || options.recursion_limit > MAX_RECURSION {
            return Err(RewriteError::limited(RewriteLimit::Nesting {
                observed: options.recursion_limit,
                maximum: MAX_RECURSION,
            }));
        }
        Ok(Self {
            options,
            input_bytes: source.len(),
            output_bytes: 0,
            fields: 0,
            work_bytes: 0,
            max_depth: 0,
            components_scanned: 0,
            components_changed: 0,
            references_scanned: 0,
            source_references_scanned: 0,
            additions: 0,
            removals: 0,
            allocations: 0,
            retained_bytes: 0,
            scratch_bytes: 0,
            source_phase: true,
        })
    }
    fn preflight_repeat_from_zero(&self) -> Result<(), RewriteError> {
        self.preflight_totals(
            self.fields.checked_mul(2),
            self.work_bytes.checked_mul(2),
            self.components_scanned.checked_mul(2),
            self.references_scanned.checked_mul(2),
        )
    }
    fn preflight_repeat_delta(&self, before: &Self) -> Result<(), RewriteError> {
        self.preflight_totals(
            self.fields.checked_add(
                self.fields
                    .checked_sub(before.fields)
                    .ok_or_else(|| RewriteError::invalid(InvalidReason::Verification))?,
            ),
            self.work_bytes.checked_add(
                self.work_bytes
                    .checked_sub(before.work_bytes)
                    .ok_or_else(|| RewriteError::invalid(InvalidReason::Verification))?,
            ),
            self.components_scanned.checked_add(
                self.components_scanned
                    .checked_sub(before.components_scanned)
                    .ok_or_else(|| RewriteError::invalid(InvalidReason::Verification))?,
            ),
            self.references_scanned.checked_add(
                self.references_scanned
                    .checked_sub(before.references_scanned)
                    .ok_or_else(|| RewriteError::invalid(InvalidReason::Verification))?,
            ),
        )
    }
    fn preflight_totals(
        &self,
        fields: Option<usize>,
        work: Option<usize>,
        components: Option<usize>,
        references: Option<usize>,
    ) -> Result<(), RewriteError> {
        let fields = fields.ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?;
        if fields > self.options.max_fields {
            return Err(RewriteError::limited(RewriteLimit::Fields {
                observed: fields,
                maximum: self.options.max_fields,
            }));
        }
        let work = work.ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?;
        if work > self.options.max_work_bytes {
            return Err(RewriteError::limited(RewriteLimit::Work {
                observed: work,
                maximum: self.options.max_work_bytes,
            }));
        }
        let components =
            components.ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?;
        if components > self.options.max_components {
            return Err(RewriteError::limited(RewriteLimit::Components {
                observed: components,
                maximum: self.options.max_components,
            }));
        }
        let references =
            references.ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?;
        if references > self.options.max_references {
            return Err(RewriteError::limited(RewriteLimit::References {
                observed: references,
                maximum: self.options.max_references,
            }));
        }
        Ok(())
    }
    fn message(&mut self, source: &[u8], depth: u32) -> Result<(), RewriteError> {
        self.message_len(source.len(), depth)
    }
    fn message_len(&mut self, amount: usize, depth: u32) -> Result<(), RewriteError> {
        self.depth(depth)?;
        self.work(amount)
    }
    fn field(&mut self) -> Result<(), RewriteError> {
        let observed = self
            .fields
            .checked_add(1)
            .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?;
        if observed > self.options.max_fields {
            return Err(RewriteError::limited(RewriteLimit::Fields {
                observed,
                maximum: self.options.max_fields,
            }));
        }
        self.fields = observed;
        Ok(())
    }
    fn work(&mut self, amount: usize) -> Result<(), RewriteError> {
        let observed = self
            .work_bytes
            .checked_add(amount)
            .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?;
        if observed > self.options.max_work_bytes {
            return Err(RewriteError::limited(RewriteLimit::Work {
                observed,
                maximum: self.options.max_work_bytes,
            }));
        }
        self.work_bytes = observed;
        Ok(())
    }
    fn depth(&mut self, depth: u32) -> Result<(), RewriteError> {
        if depth > self.options.recursion_limit {
            return Err(RewriteError::limited(RewriteLimit::Nesting {
                observed: depth,
                maximum: self.options.recursion_limit,
            }));
        }
        self.max_depth = self.max_depth.max(depth);
        Ok(())
    }
    fn component(&mut self) -> Result<(), RewriteError> {
        let observed = self
            .components_scanned
            .checked_add(1)
            .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?;
        if observed > self.options.max_components {
            return Err(RewriteError::limited(RewriteLimit::Components {
                observed,
                maximum: self.options.max_components,
            }));
        }
        self.components_scanned = observed;
        Ok(())
    }
    fn changed_component(&mut self) -> Result<(), RewriteError> {
        self.components_changed = self
            .components_changed
            .checked_add(1)
            .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?;
        Ok(())
    }
    fn reference(&mut self) -> Result<(), RewriteError> {
        let observed = self
            .references_scanned
            .checked_add(1)
            .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?;
        if observed > self.options.max_references {
            return Err(RewriteError::limited(RewriteLimit::References {
                observed,
                maximum: self.options.max_references,
            }));
        }
        self.references_scanned = observed;
        if self.source_phase {
            self.source_references_scanned = self
                .source_references_scanned
                .checked_add(1)
                .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?;
        }
        Ok(())
    }
    fn output_size(&self, amount: usize) -> Result<(), RewriteError> {
        if amount > self.options.max_output_bytes {
            return Err(RewriteError::limited(RewriteLimit::OutputBytes {
                observed: amount,
                maximum: self.options.max_output_bytes,
            }));
        }
        Ok(())
    }
    fn allocation(&mut self, scratch: usize) -> Result<(), RewriteError> {
        self.allocations = self
            .allocations
            .checked_add(1)
            .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?;
        self.scratch_bytes = self
            .scratch_bytes
            .checked_add(scratch)
            .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?;
        Ok(())
    }
    const fn report(&self) -> RewriteReport {
        RewriteReport {
            input_bytes: self.input_bytes,
            output_bytes: self.output_bytes,
            fields: self.fields,
            work_bytes: self.work_bytes,
            max_depth: self.max_depth,
            components_scanned: self.components_scanned,
            components_changed: self.components_changed,
            references_scanned: self.references_scanned,
            source_references_scanned: self.source_references_scanned,
            additions: self.additions,
            removals: self.removals,
            allocations: self.allocations,
            retained_bytes: self.retained_bytes,
            scratch_bytes: self.scratch_bytes,
        }
    }
}

#[derive(Clone, Copy)]
struct Field<'source> {
    number: u32,
    wire: u8,
    value: Value<'source>,
    raw: &'source [u8],
}

impl<'source> Field<'source> {
    fn varint(self) -> Result<u64, RewriteError> {
        match self.value {
            Value::Varint(value, encoded_len)
                if self.wire == 0 && encoded_varint_len(value) == encoded_len =>
            {
                Ok(value)
            },
            _ => Err(RewriteError::invalid(InvalidReason::MalformedWire)),
        }
    }
    fn bytes(self) -> Result<&'source [u8], RewriteError> {
        match self.value {
            Value::Bytes(value) if self.wire == 2 => Ok(value),
            _ => Err(RewriteError::invalid(InvalidReason::MalformedWire)),
        }
    }
}

#[derive(Clone, Copy)]
enum Value<'source> {
    Varint(u64, usize),
    Fixed64,
    Bytes(&'source [u8]),
    Group,
    Fixed32,
}
enum ParseItem<'source> {
    Field(Field<'source>),
    EndGroup(u32),
}

fn next_field<'source>(
    source: &mut &'source [u8],
    budget: &mut Budget,
    depth: u32,
) -> Result<Option<Field<'source>>, RewriteError> {
    match parse_field(source, budget, depth)? {
        Some(ParseItem::Field(field)) => Ok(Some(field)),
        Some(ParseItem::EndGroup(_)) => Err(RewriteError::invalid(InvalidReason::MalformedWire)),
        None => Ok(None),
    }
}

fn parse_field<'source>(
    source: &mut &'source [u8],
    budget: &mut Budget,
    depth: u32,
) -> Result<Option<ParseItem<'source>>, RewriteError> {
    if source.is_empty() {
        return Ok(None);
    }
    let original = *source;
    budget.depth(depth)?;
    budget.field()?;
    let tag = take_varint(source)?;
    let number = u32::try_from(tag >> 3)
        .map_err(|_error| RewriteError::invalid(InvalidReason::MalformedWire))?;
    let wire = u8::try_from(tag & 7)
        .map_err(|_error| RewriteError::invalid(InvalidReason::MalformedWire))?;
    if number == 0 || number > MAX_FIELD_NUMBER {
        return Err(RewriteError::invalid(InvalidReason::MalformedWire));
    }
    let value = match wire {
        0 => {
            let (value, encoded_len) = take_varint_relaxed(source)?;
            Value::Varint(value, encoded_len)
        },
        1 => {
            take(source, 8)?;
            Value::Fixed64
        },
        2 => {
            let length = usize::try_from(take_varint(source)?)
                .map_err(|_error| RewriteError::invalid(InvalidReason::MalformedWire))?;
            Value::Bytes(take(source, length)?)
        },
        3 => {
            let child = depth
                .checked_add(1)
                .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?;
            skip_group(source, number, budget, child)?;
            Value::Group
        },
        4 => return Ok(Some(ParseItem::EndGroup(number))),
        5 => {
            take(source, 4)?;
            Value::Fixed32
        },
        _ => return Err(RewriteError::invalid(InvalidReason::MalformedWire)),
    };
    let consumed = original.len() - source.len();
    Ok(Some(ParseItem::Field(Field {
        number,
        wire,
        value,
        raw: &original[..consumed],
    })))
}

fn skip_group(
    source: &mut &[u8],
    expected: u32,
    budget: &mut Budget,
    depth: u32,
) -> Result<(), RewriteError> {
    loop {
        match parse_field(source, budget, depth)? {
            Some(ParseItem::Field(_)) => {},
            Some(ParseItem::EndGroup(number)) if number == expected => return Ok(()),
            Some(ParseItem::EndGroup(_)) | None => {
                return Err(RewriteError::invalid(InvalidReason::MalformedWire));
            },
        }
    }
}

fn take<'source>(source: &mut &'source [u8], amount: usize) -> Result<&'source [u8], RewriteError> {
    if source.len() < amount {
        return Err(RewriteError::invalid(InvalidReason::MalformedWire));
    }
    let (selected, rest) = source.split_at(amount);
    *source = rest;
    Ok(selected)
}

fn take_varint(source: &mut &[u8]) -> Result<u64, RewriteError> {
    let (value, consumed) = take_varint_relaxed(source)?;
    if encoded_varint_len(value) != consumed {
        return Err(RewriteError::invalid(InvalidReason::MalformedWire));
    }
    Ok(value)
}

fn take_varint_relaxed(source: &mut &[u8]) -> Result<(u64, usize), RewriteError> {
    let original = *source;
    let mut value = 0u64;
    for index in 0..10usize {
        let byte = *original
            .get(index)
            .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?;
        if index == 9 && byte > 1 {
            return Err(RewriteError::invalid(InvalidReason::MalformedWire));
        }
        value |= u64::from(byte & 0x7f) << (index * 7);
        if byte & 0x80 == 0 {
            let consumed = index + 1;
            *source = &original[consumed..];
            return Ok((value, consumed));
        }
    }
    Err(RewriteError::invalid(InvalidReason::MalformedWire))
}

const fn encoded_varint_len(value: u64) -> usize {
    if value == 0 {
        1
    } else {
        (64usize - value.leading_zeros() as usize).div_ceil(7)
    }
}

fn canonical_bool(value: u64) -> Result<bool, RewriteError> {
    match value {
        0 => Ok(false),
        1 => Ok(true),
        _ => Err(RewriteError::invalid(InvalidReason::MalformedWire)),
    }
}

fn strict_utf8(source: &[u8]) -> Result<&str, RewriteError> {
    str::from_utf8(source).map_err(|_error| RewriteError::invalid(InvalidReason::MalformedWire))
}

fn set_once<T>(slot: &mut Option<T>, value: T) -> Result<(), RewriteError> {
    if slot.is_some() {
        return Err(RewriteError::invalid(InvalidReason::MalformedWire));
    }
    *slot = Some(value);
    Ok(())
}
