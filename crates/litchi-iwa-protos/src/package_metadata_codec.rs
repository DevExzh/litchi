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

use buffa::DecodeOptions as BuffaDecodeOptions;

use crate::buffa_package_metadata_generated::LitchiIwaPackageMetadataProjection as projection;

const MAX_RECURSION: u32 = 64;
const MAX_FIELD_NUMBER: u32 = 0x1fff_ffff;

#[cfg(test)]
std::thread_local! {
    static OUTPUT_ALLOCATIONS: core::cell::Cell<usize> = const { core::cell::Cell::new(0) };
    static WORK_CHARGES: core::cell::Cell<usize> = const { core::cell::Cell::new(0) };
}

#[cfg(test)]
fn record_output_allocation() {
    OUTPUT_ALLOCATIONS.set(OUTPUT_ALLOCATIONS.get() + 1);
}

#[cfg(test)]
fn output_allocations() -> usize {
    OUTPUT_ALLOCATIONS.get()
}

#[cfg(test)]
fn reset_work_charges() {
    WORK_CHARGES.set(0);
}

#[cfg(test)]
fn work_charges() -> usize {
    WORK_CHARGES.get()
}

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

/// Borrowed exact selectors for a save-token publication.
///
/// Each selector authorizes exactly one current, unversioned component. The
/// component identifier and its effective locator (the explicit locator when
/// present, otherwise the preferred locator) must both match exactly.
#[derive(Debug, Clone, Copy)]
pub struct SaveTokenBatch<'source> {
    components: &'source [ComponentSelector<'source>],
}

impl<'source> SaveTokenBatch<'source> {
    #[must_use]
    pub const fn new(components: &'source [ComponentSelector<'source>]) -> Self {
        Self { components }
    }

    #[must_use]
    pub const fn components(self) -> &'source [ComponentSelector<'source>] {
        self.components
    }

    #[must_use]
    pub const fn selectors(self) -> &'source [ComponentSelector<'source>] {
        self.components
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

/// One atomic registry-removal and save-token transition.
///
/// The two borrowed batches are evaluated against the same source snapshot
/// and published as one candidate.  This is intentionally separate from
/// [`Batch`]: removal keeps `last_object_identifier` unchanged, while a
/// save-token transition changes only fields 8 and 12.
#[derive(Debug, Clone, Copy)]
pub struct RemovalSaveTokenBatch<'source> {
    removals: RemovalBatch<'source>,
    save_tokens: SaveTokenBatch<'source>,
}

impl<'source> RemovalSaveTokenBatch<'source> {
    #[must_use]
    pub const fn new(
        removals: RemovalBatch<'source>,
        save_tokens: SaveTokenBatch<'source>,
    ) -> Self {
        Self {
            removals,
            save_tokens,
        }
    }

    #[must_use]
    pub const fn removals(self) -> RemovalBatch<'source> {
        self.removals
    }

    #[must_use]
    pub const fn save_tokens(self) -> SaveTokenBatch<'source> {
        self.save_tokens
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

/// Exact resources required by the allocation-bearing PackageMetadata phase.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RewriteExecutionRequirements {
    output_bytes: usize,
    fields: usize,
    work_bytes: usize,
    components: usize,
    references: usize,
    allocations: usize,
    retained_bytes: usize,
    scratch_bytes: usize,
}

macro_rules! execution_requirement_accessors {
    ($(($name:ident, $ty:ty)),+ $(,)?) => {$(
        #[must_use]
        pub const fn $name(self) -> $ty { self.$name }
    )+};
}

impl RewriteExecutionRequirements {
    execution_requirement_accessors!(
        (output_bytes, usize),
        (fields, usize),
        (work_bytes, usize),
        (components, usize),
        (references, usize),
        (allocations, usize),
        (retained_bytes, usize),
        (scratch_bytes, usize)
    );

    #[must_use]
    pub const fn exact_limits(self) -> RewriteExecutionLimits {
        RewriteExecutionLimits {
            max_output_bytes: self.output_bytes,
            max_fields: self.fields,
            max_work_bytes: self.work_bytes,
            max_components: self.components,
            max_references: self.references,
            max_allocations: self.allocations,
            max_retained_bytes: self.retained_bytes,
            max_scratch_bytes: self.scratch_bytes,
        }
    }
}

/// Independent finite limits for executing a prepared rewrite.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RewriteExecutionLimits {
    pub max_output_bytes: usize,
    pub max_fields: usize,
    pub max_work_bytes: usize,
    pub max_components: usize,
    pub max_references: usize,
    pub max_allocations: usize,
    pub max_retained_bytes: usize,
    pub max_scratch_bytes: usize,
}

/// Output-free, semantically validated PackageMetadata rewrite.
pub struct PreparedPackageMetadataRewrite<'source, 'batch> {
    source: &'source [u8],
    batch: Batch<'batch>,
    budget: Budget,
    prepare_report: RewriteReport,
    requirements: RewriteExecutionRequirements,
    output_size: usize,
}

impl PreparedPackageMetadataRewrite<'_, '_> {
    #[must_use]
    pub const fn prepare_report(&self) -> RewriteReport {
        self.prepare_report
    }

    #[must_use]
    pub const fn execution_requirements(&self) -> RewriteExecutionRequirements {
        self.requirements
    }

    pub fn execute(
        mut self,
        limits: RewriteExecutionLimits,
    ) -> Result<RewriteOutput, RewriteError> {
        preflight_execution(self.requirements, limits)?;
        let before = self.budget.report();
        let mut candidate = Vec::new();
        #[cfg(test)]
        record_output_allocation();
        candidate
            .try_reserve_exact(self.output_size)
            .map_err(|_error| RewriteError::allocation(self.output_size))?;
        if candidate.capacity() != self.output_size {
            return Err(RewriteError::allocation(self.output_size));
        }
        self.budget.allocation(0)?;
        rewrite_into(self.source, self.batch, &mut candidate, &mut self.budget)?;
        if candidate.len() != self.output_size {
            return Err(RewriteError::invalid(InvalidReason::Verification));
        }

        self.budget.source_phase = false;
        let mut verified_state = ScanState::new(self.batch, &mut self.budget)?;
        scan_metadata(
            &candidate,
            self.batch,
            ScanMode::Verification,
            &mut verified_state,
            &mut self.budget,
            false,
        )?;
        verified_state.validate_verification()?;
        self.budget.output_bytes = candidate.len();
        self.budget.retained_bytes = candidate.len();
        let report = subtract_report(self.budget.report(), before)?;
        validate_execution_report(report, self.requirements)?;
        Ok(RewriteOutput {
            bytes: candidate,
            report,
        })
    }
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
    DuplicateSelector,
    VersionedComponent,
    SaveTokenMismatch,
    DuplicateSaveToken,
    SaveTokenOverflow,
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
    /// Construct a typed failure for a caller-owned visitor staging
    /// allocation. Visitors can use this when `try_reserve` refuses to grow
    /// temporary transaction state; the enclosing inspection forwards the
    /// failure without publishing a candidate.
    #[must_use]
    pub const fn allocation(requested: usize) -> Self {
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
        if size_of::<T>() != 0 && output.capacity() != amount {
            return Err(RewriteError::allocation(amount));
        }
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
            let allocations_before = output_allocations();
            let error = rewrite_package_metadata(&source, batch, limited).unwrap_err();
            assert!(error.resource_limit().is_some(), "missing {label} limit");
            assert_eq!(
                output_allocations(),
                allocations_before,
                "{label} limit reached the output allocation"
            );
        }
    }

    #[test]
    fn prepared_rewrite_is_output_free_and_execute_limits_precede_candidate() {
        let source = metadata(
            10,
            &[
                component(1, "a.iwa", None, &[], &[]),
                component(2, "b.iwa", None, &[], &[]),
            ],
            &[],
        );
        let additions = [ObjectUuidAddition::new(
            ComponentSelector::new(1, "a.iwa"),
            11,
            UuidBits::new(5, 6),
        )];
        let batch = Batch::new(10, 11, &additions, &[]);
        let allocations_before = output_allocations();
        let prepared = prepare_package_metadata_rewrite(&source, batch, options(&source)).unwrap();
        assert_eq!(prepared.prepare_report().output_bytes(), 0);
        assert_eq!(prepared.prepare_report().retained_bytes(), 0);
        assert_eq!(output_allocations(), allocations_before);
        let requirements = prepared.execution_requirements();
        assert!(requirements.output_bytes() != 0);
        assert!(requirements.allocations() != 0);

        let mut limited = requirements.exact_limits();
        limited.max_output_bytes -= 1;
        assert!(
            prepared
                .execute(limited)
                .unwrap_err()
                .resource_limit()
                .is_some()
        );
        assert_eq!(output_allocations(), allocations_before);

        for axis in 0..4 {
            let prepared =
                prepare_package_metadata_rewrite(&source, batch, options(&source)).unwrap();
            let requirements = prepared.execution_requirements();
            let mut limited = requirements.exact_limits();
            match axis {
                0 => limited.max_allocations -= 1,
                1 => limited.max_retained_bytes -= 1,
                2 => limited.max_scratch_bytes -= 1,
                _ => limited.max_work_bytes -= 1,
            }
            let before = output_allocations();
            assert!(prepared.execute(limited).is_err());
            assert_eq!(output_allocations(), before);
        }

        let prepared = prepare_package_metadata_rewrite(&source, batch, options(&source)).unwrap();
        let requirements = prepared.execution_requirements();
        let output = prepared.execute(requirements.exact_limits()).unwrap();
        assert_eq!(output.report().output_bytes(), output.bytes().len());
        assert_eq!(output.report().retained_bytes(), output.bytes().len());
        assert_eq!(output.report().allocations(), requirements.allocations());
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

    #[derive(Default)]
    struct AllocationRefusingVisitor {
        components: usize,
    }

    impl PackageMetadataVisitor for AllocationRefusingVisitor {
        fn visit_component(
            &mut self,
            _component: ComponentDescriptor<'_>,
        ) -> Result<(), RewriteError> {
            self.components += 1;
            Err(RewriteError::allocation(2))
        }
    }

    #[test]
    fn visitor_allocation_failure_is_forwarded_without_following_observations() {
        let source = metadata(
            10,
            &[
                component(1, "a.iwa", None, &[], &[]),
                component(2, "b.iwa", None, &[], &[]),
            ],
            &[],
        );
        let before = source.clone();
        let mut visitor = AllocationRefusingVisitor::default();
        let error = inspect_package_metadata_with_visitor(&source, options(&source), &mut visitor)
            .unwrap_err();

        assert_eq!(visitor.components, 1);
        assert_eq!(error.allocation_request(), Some(2));
        assert_eq!(source, before);
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
    fn external_only_removal_preserves_other_component_owners_of_the_same_object() {
        let selected = ComponentSelector::new(1, "a.iwa");
        let target = ComponentSelector::new(2, "styles.iwa");
        let a = component(1, "a.iwa", None, &[], &[(6, 2, Some(5), Some(0))]);
        let styles = component(2, "styles.iwa", None, &[(5, UuidBits::new(10, 20))], &[]);
        let other = component(3, "other.iwa", None, &[], &[(6, 2, Some(5), Some(0))]);
        let source = metadata(10, &[a, styles, other.clone()], &[]);
        let externals = [ExternalReferenceRemoval::new(
            selected,
            target,
            5,
            Some(false),
        )];

        let output = remove_package_metadata(
            &source,
            RemovalBatch::new(10, &[], &externals, &[]),
            options(&source),
        )
        .unwrap();

        assert_eq!(output.report().removals(), 1);
        assert!(
            output
                .bytes()
                .windows(other.len())
                .any(|window| window == other)
        );
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
        let selected = data_reference(71, &[(5, 2)], false);
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
    fn combined_removal_and_save_tokens_is_one_raw_preserving_transition() {
        let selector = ComponentSelector::new(1, "a.iwa");
        let target = ComponentSelector::new(2, "b.iwa");
        let uuid = UuidBits::new(10, 20);
        let mut selected = component(1, "a.iwa", None, &[(5, uuid)], &[(6, 2, Some(5), Some(0))]);
        put_varint_field(&mut selected, 12, 536);
        let selected_unknown = [0x90, 0x03, 0x81, 0x00];
        selected.extend_from_slice(&selected_unknown);
        let untouched = component(2, "b.iwa", None, &[], &[]);
        let mut source = metadata(10, &[selected.clone(), untouched.clone()], &[]);
        put_varint_field(&mut source, 8, 536);

        let uuids = [ObjectUuidRemoval::new(selector, 5, uuid)];
        let externals = [ExternalReferenceRemoval::new(
            selector,
            target,
            5,
            Some(false),
        )];
        let removals = RemovalBatch::new(10, &uuids, &externals, &[]);
        let save_tokens = SaveTokenBatch::new(core::slice::from_ref(&selector));
        let batch = RemovalSaveTokenBatch::new(removals, save_tokens);
        let allocations = output_allocations();
        let output =
            rewrite_package_metadata_removals_and_save_tokens(&source, batch, options(&source))
                .unwrap();

        assert_eq!(output.report().removals(), 2);
        assert_eq!(output.report().additions(), 0);
        assert_eq!(output.report().components_changed(), 1);
        assert_eq!(output.report().output_bytes(), output.bytes().len());
        assert_eq!(output.report().retained_bytes(), output.bytes().len());
        assert_eq!(output_allocations(), allocations + 1);
        assert_eq!(scalar_values(output.bytes(), 8), vec![537]);
        assert_eq!(component_scalar_values(output.bytes(), 1, 12), vec![537]);
        assert!(
            output
                .bytes()
                .windows(selected_unknown.len())
                .any(|window| { window == selected_unknown })
        );
        assert!(
            output
                .bytes()
                .windows(untouched.len())
                .any(|window| { window == untouched })
        );
        assert!(
            !output
                .bytes()
                .windows(uuid_entry(5, uuid).len())
                .any(|window| { window == uuid_entry(5, uuid) })
        );
        assert!(
            !output
                .bytes()
                .windows(external_reference(2, Some(5), Some(0)).len())
                .any(|window| window == external_reference(2, Some(5), Some(0)))
        );
        let limited = RewriteOptions::new(
            source.len(),
            output.report().output_bytes() - 1,
            output.report().fields(),
            output.report().work_bytes(),
            output.report().max_depth(),
            output.report().components_scanned(),
            output.report().references_scanned(),
            output.report().removals(),
        );
        let allocations = output_allocations();
        assert!(matches!(
            rewrite_package_metadata_removals_and_save_tokens(&source, batch, limited)
                .unwrap_err()
                .resource_limit(),
            Some(RewriteLimit::OutputBytes { .. })
        ));
        assert_eq!(output_allocations(), allocations);
    }

    #[test]
    fn combined_transition_requires_exact_mutation_component_token_coverage() {
        let a = ComponentSelector::new(1, "a.iwa");
        let b = ComponentSelector::new(2, "b.iwa");
        let uuid = UuidBits::new(10, 20);
        let mut selected = component(1, "a.iwa", None, &[(5, uuid)], &[]);
        put_varint_field(&mut selected, 12, 5);
        let mut other = component(2, "b.iwa", None, &[], &[]);
        put_varint_field(&mut other, 12, 5);
        let mut source = metadata(10, &[selected, other], &[]);
        put_varint_field(&mut source, 8, 5);
        let uuid_removals = [ObjectUuidRemoval::new(a, 5, uuid)];
        let removals = RemovalBatch::new(10, &uuid_removals, &[], &[]);

        let missing_selectors = [b];
        let missing = RemovalSaveTokenBatch::new(removals, SaveTokenBatch::new(&missing_selectors));
        assert_eq!(
            reason(
                rewrite_package_metadata_removals_and_save_tokens(
                    &source,
                    missing,
                    options(&source),
                )
                .unwrap_err()
            ),
            InvalidReason::ComponentMismatch
        );

        let extra_selectors = [a, b];
        let extra = RemovalSaveTokenBatch::new(removals, SaveTokenBatch::new(&extra_selectors));
        assert_eq!(
            reason(
                rewrite_package_metadata_removals_and_save_tokens(
                    &source,
                    extra,
                    options(&source),
                )
                .unwrap_err()
            ),
            InvalidReason::ComponentMismatch
        );
    }

    #[test]
    fn combined_transition_rejects_unknown_fields_in_removed_nested_records() {
        let selector = ComponentSelector::new(1, "a.iwa");
        let target = ComponentSelector::new(2, "b.iwa");
        let uuid = UuidBits::new(10, 20);

        let mut external = external_reference(2, Some(5), Some(0));
        put_varint_field(&mut external, 30, 99);
        let mut current = component(1, "a.iwa", None, &[], &[]);
        bytes_field(&mut current, 6, &external);
        let mut target_component = component(2, "b.iwa", None, &[(5, uuid)], &[]);
        put_varint_field(&mut target_component, 12, 5);
        let mut source = metadata(10, &[current, target_component], &[]);
        put_varint_field(&mut source, 8, 5);
        let external_removal = [ExternalReferenceRemoval::new(
            selector,
            target,
            5,
            Some(false),
        )];
        let removals = RemovalBatch::new(10, &[], &external_removal, &[]);
        let batch = RemovalSaveTokenBatch::new(
            removals,
            SaveTokenBatch::new(core::slice::from_ref(&selector)),
        );
        assert_eq!(
            reason(
                rewrite_package_metadata_removals_and_save_tokens(
                    &source,
                    batch,
                    options(&source),
                )
                .unwrap_err()
            ),
            InvalidReason::RemovalMismatch
        );

        let mut uuid_entry_bytes = uuid_entry(5, uuid);
        put_varint_field(&mut uuid_entry_bytes, 30, 99);
        let mut uuid_component = component(1, "a.iwa", None, &[], &[]);
        bytes_field(&mut uuid_component, 11, &uuid_entry_bytes);
        put_varint_field(&mut uuid_component, 12, 5);
        let mut uuid_source = metadata(10, &[uuid_component], &[]);
        put_varint_field(&mut uuid_source, 8, 5);
        let uuid_removal = [ObjectUuidRemoval::new(selector, 5, uuid)];
        let uuid_batch = RemovalSaveTokenBatch::new(
            RemovalBatch::new(10, &uuid_removal, &[], &[]),
            SaveTokenBatch::new(core::slice::from_ref(&selector)),
        );
        assert_eq!(
            reason(
                rewrite_package_metadata_removals_and_save_tokens(
                    &uuid_source,
                    uuid_batch,
                    options(&uuid_source),
                )
                .unwrap_err()
            ),
            InvalidReason::RemovalMismatch
        );
    }

    #[test]
    fn combined_transition_limits_each_metered_dimension_before_allocation() {
        let selector = ComponentSelector::new(1, "a.iwa");
        let target = ComponentSelector::new(2, "b.iwa");
        let uuid = UuidBits::new(10, 20);
        let mut current = component(1, "a.iwa", None, &[(5, uuid)], &[(6, 2, Some(5), Some(0))]);
        put_varint_field(&mut current, 12, 5);
        let mut target_component = component(2, "b.iwa", None, &[], &[]);
        put_varint_field(&mut target_component, 12, 5);
        let mut source = metadata(10, &[current, target_component], &[]);
        put_varint_field(&mut source, 8, 5);
        let uuids = [ObjectUuidRemoval::new(selector, 5, uuid)];
        let externals = [ExternalReferenceRemoval::new(
            selector,
            target,
            5,
            Some(false),
        )];
        let batch = RemovalSaveTokenBatch::new(
            RemovalBatch::new(10, &uuids, &externals, &[]),
            SaveTokenBatch::new(core::slice::from_ref(&selector)),
        );
        reset_work_charges();
        let baseline =
            rewrite_package_metadata_removals_and_save_tokens(&source, batch, options(&source))
                .unwrap();
        let report = baseline.report();
        assert_eq!(work_charges(), report.work_bytes());
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
        let allocations = output_allocations();
        reset_work_charges();
        let replay = rewrite_package_metadata_removals_and_save_tokens(&source, batch, exact)
            .expect("the report is a replayable aggregate budget");
        assert_eq!(replay.bytes(), baseline.bytes());
        assert_eq!(replay.report(), report);
        assert_eq!(work_charges(), replay.report().work_bytes());
        assert_eq!(output_allocations(), allocations + 1);

        let limited = [
            RewriteOptions::new(
                source.len(),
                report.output_bytes(),
                report.fields().saturating_sub(1),
                report.work_bytes(),
                report.max_depth(),
                report.components_scanned(),
                report.references_scanned(),
                report.removals(),
            ),
            RewriteOptions::new(
                source.len(),
                report.output_bytes(),
                report.fields(),
                report.work_bytes().saturating_sub(1),
                report.max_depth(),
                report.components_scanned(),
                report.references_scanned(),
                report.removals(),
            ),
            RewriteOptions::new(
                source.len(),
                report.output_bytes(),
                report.fields(),
                report.work_bytes(),
                report.max_depth(),
                report.components_scanned().saturating_sub(1),
                report.references_scanned(),
                report.removals(),
            ),
            RewriteOptions::new(
                source.len(),
                report.output_bytes(),
                report.fields(),
                report.work_bytes(),
                report.max_depth(),
                report.components_scanned(),
                report.references_scanned().saturating_sub(1),
                report.removals(),
            ),
            RewriteOptions::new(
                source.len(),
                report.output_bytes(),
                report.fields(),
                report.work_bytes(),
                report.max_depth().saturating_sub(1),
                report.components_scanned(),
                report.references_scanned(),
                report.removals(),
            ),
            RewriteOptions::new(
                source.len(),
                report.output_bytes(),
                report.fields(),
                report.work_bytes(),
                report.max_depth(),
                report.components_scanned(),
                report.references_scanned(),
                report.removals().saturating_sub(1),
            ),
        ];
        for options in limited {
            let before = output_allocations();
            let error = rewrite_package_metadata_removals_and_save_tokens(&source, batch, options)
                .expect_err("each max-minus-one dimension must fail in preflight");
            assert!(error.resource_limit().is_some());
            assert_eq!(output_allocations(), before);
        }
    }

    #[test]
    fn combined_transition_adds_absent_tokens_and_rejects_hostile_removals() {
        let selector = ComponentSelector::new(1, "a.iwa");
        let uuid = UuidBits::new(10, 20);
        let mut selected = component(1, "a.iwa", None, &[(5, uuid)], &[]);
        put_varint_field(&mut selected, 12, 2);
        let untouched = component(2, "b.iwa", None, &[], &[]);
        let mut source = metadata(10, &[selected, untouched], &[]);
        put_varint_field(&mut source, 8, 0);
        let hostile_removals = [ObjectUuidRemoval::new(selector, 5, uuid)];
        let removals = RemovalBatch::new(10, &hostile_removals, &[], &[]);
        let batch = RemovalSaveTokenBatch::new(
            removals,
            SaveTokenBatch::new(core::slice::from_ref(&selector)),
        );
        let error =
            rewrite_package_metadata_removals_and_save_tokens(&source, batch, options(&source))
                .unwrap_err();
        assert_eq!(reason(error), InvalidReason::SaveTokenMismatch);

        let mut valid_source = metadata(10, &[component(1, "a.iwa", None, &[(5, uuid)], &[])], &[]);
        put_varint_field(&mut valid_source, 8, 0);
        let valid_entries = [ObjectUuidRemoval::new(selector, 5, uuid)];
        let valid_removals = RemovalBatch::new(10, &valid_entries, &[], &[]);
        let valid = rewrite_package_metadata_removals_and_save_tokens(
            &valid_source,
            RemovalSaveTokenBatch::new(
                valid_removals,
                SaveTokenBatch::new(core::slice::from_ref(&selector)),
            ),
            options(&valid_source),
        )
        .unwrap();
        assert_eq!(scalar_values(valid.bytes(), 8), vec![1]);
        assert_eq!(component_scalar_values(valid.bytes(), 1, 12), vec![1]);

        let versioned = component(9, "old.iwa", None, &[(5, uuid)], &[]);
        let mut hostile = metadata(
            10,
            &[component(1, "a.iwa", None, &[(5, uuid)], &[])],
            &[versioned],
        );
        put_varint_field(&mut hostile, 8, 0);
        let error = rewrite_package_metadata_removals_and_save_tokens(
            &hostile,
            RemovalSaveTokenBatch::new(
                valid_removals,
                SaveTokenBatch::new(core::slice::from_ref(&selector)),
            ),
            options(&hostile),
        )
        .unwrap_err();
        assert_eq!(reason(error), InvalidReason::VersionedRemoval);
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
        let allocations = output_allocations();
        let error = remove_package_metadata(&source, batch, limited).unwrap_err();
        assert!(matches!(
            error.resource_limit(),
            Some(RewriteLimit::OutputBytes { .. })
        ));
        assert_eq!(output_allocations(), allocations);

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

    fn token_component(
        identifier: u64,
        locator: &str,
        token: Option<u64>,
        token_first: bool,
        unknown: bool,
    ) -> Vec<u8> {
        let mut output = Vec::new();
        if token_first {
            if let Some(token) = token {
                put_varint_field(&mut output, 12, token);
            }
        }
        put_varint_field(&mut output, 1, identifier);
        bytes_field(&mut output, 2, locator.as_bytes());
        if unknown {
            put_key(&mut output, 50, 0);
            output.extend_from_slice(&[0x81, 0x00]);
            put_key(&mut output, 51, 3);
            put_varint_field(&mut output, 1, 0);
            put_key(&mut output, 51, 4);
        }
        if !token_first {
            if let Some(token) = token {
                put_varint_field(&mut output, 12, token);
            }
        }
        output
    }

    fn token_component_with_locator(
        identifier: u64,
        preferred: &str,
        locator: &str,
        token: Option<u64>,
    ) -> Vec<u8> {
        let mut output = Vec::new();
        put_varint_field(&mut output, 1, identifier);
        bytes_field(&mut output, 2, preferred.as_bytes());
        bytes_field(&mut output, 3, locator.as_bytes());
        if let Some(token) = token {
            put_varint_field(&mut output, 12, token);
        }
        output
    }

    fn token_metadata(
        last: u64,
        root_token: Option<u64>,
        current: &[Vec<u8>],
        versioned: &[Vec<u8>],
    ) -> Vec<u8> {
        let mut output = Vec::new();
        put_key(&mut output, 50, 0);
        output.extend_from_slice(&[0x81, 0x00]);
        put_key(&mut output, 51, 3);
        put_varint_field(&mut output, 1, 0);
        put_key(&mut output, 51, 4);
        for component in current {
            bytes_field(&mut output, 3, component);
        }
        for component in versioned {
            bytes_field(&mut output, 11, component);
        }
        put_varint_field(&mut output, 1, last);
        if let Some(root_token) = root_token {
            put_varint_field(&mut output, 8, root_token);
        }
        output
    }

    fn scalar_values(source: &[u8], number: u32) -> Vec<u64> {
        let mut budget = Budget::new_inspection(source, options(source)).unwrap();
        let mut remaining = source;
        let mut values = Vec::new();
        while let Some(field) = next_field(&mut remaining, &mut budget, 1).unwrap() {
            if field.number == number {
                values.push(field.varint().unwrap());
            }
        }
        values
    }

    fn component_scalar_values(source: &[u8], identifier: u64, number: u32) -> Vec<u64> {
        let mut budget = Budget::new_inspection(source, options(source)).unwrap();
        let mut remaining = source;
        while let Some(field) = next_field(&mut remaining, &mut budget, 1).unwrap() {
            if field.number != 3 {
                continue;
            }
            let payload = field.bytes().unwrap();
            let mut nested_budget = Budget::new_inspection(payload, options(payload)).unwrap();
            let mut nested = payload;
            let mut id = None;
            let mut values = Vec::new();
            while let Some(field) = next_field(&mut nested, &mut nested_budget, 2).unwrap() {
                if field.number == 1 {
                    id = Some(field.varint().unwrap());
                } else if field.number == number {
                    values.push(field.varint().unwrap());
                }
            }
            if id == Some(identifier) {
                return values;
            }
        }
        Vec::new()
    }

    #[test]
    fn save_tokens_update_selected_current_components_and_preserve_aliases() {
        let selected = token_component(1, "a.iwa", Some(536), true, true);
        let unselected = token_component(2, "b.iwa", Some(500), false, true);
        let versioned = token_component(1, "old.iwa", Some(536), false, true);
        let source = token_metadata(
            77,
            Some(536),
            &[selected.clone(), unselected.clone()],
            core::slice::from_ref(&versioned),
        );
        let selector = ComponentSelector::new(1, "a.iwa");
        let output = rewrite_package_metadata_save_tokens(
            &source,
            SaveTokenBatch::new(core::slice::from_ref(&selector)),
            options(&source),
        )
        .unwrap();
        assert_eq!(scalar_values(output.bytes(), 8), vec![537]);
        assert_eq!(component_scalar_values(output.bytes(), 1, 12), vec![537]);
        assert!(
            output
                .bytes()
                .windows(unselected.len())
                .any(|window| window == unselected)
        );
        assert!(
            output
                .bytes()
                .windows(versioned.len())
                .any(|window| window == versioned)
        );
        assert!(
            output
                .bytes()
                .windows(4)
                .any(|window| { window == [0x90, 0x03, 0x81, 0x00] })
        );
        assert!(
            output
                .bytes()
                .windows(6)
                .any(|window| window == [0x9b, 0x03, 0x08, 0x00, 0x9c, 0x03])
        );
        assert_eq!(output.report().components_changed(), 1);
    }

    #[test]
    fn save_tokens_add_absent_fields_and_preserve_last_identifier_raw_bytes() {
        let component = token_component(1, "a.iwa", None, false, false);
        let source = token_metadata(77, None, core::slice::from_ref(&component), &[]);
        let selector = ComponentSelector::new(1, "a.iwa");
        let output = rewrite_package_metadata_save_tokens(
            &source,
            SaveTokenBatch::new(core::slice::from_ref(&selector)),
            options(&source),
        )
        .unwrap();
        assert_eq!(scalar_values(output.bytes(), 8), vec![1]);
        assert!(
            output
                .bytes()
                .windows(2)
                .any(|window| { window == [0x60, 0x01] })
        );
        assert!(
            output
                .bytes()
                .windows(2)
                .any(|window| { window == [0x08, 0x4d] })
        );
    }

    #[test]
    fn save_tokens_reject_collisions_malformed_known_fields_and_overflow() {
        let component = token_component(1, "a.iwa", Some(5), false, false);
        let source = token_metadata(77, Some(5), core::slice::from_ref(&component), &[]);
        let selector = ComponentSelector::new(1, "a.iwa");
        let duplicate = [selector, selector];
        assert_eq!(
            reason(
                rewrite_package_metadata_save_tokens(
                    &source,
                    SaveTokenBatch::new(&duplicate),
                    options(&source),
                )
                .unwrap_err()
            ),
            InvalidReason::DuplicateSelector
        );
        let versioned_only = token_metadata(
            77,
            Some(5),
            core::slice::from_ref(&component),
            &[token_component(9, "old.iwa", Some(5), false, false)],
        );
        assert_eq!(
            reason(
                rewrite_package_metadata_save_tokens(
                    &versioned_only,
                    SaveTokenBatch::new(core::slice::from_ref(&ComponentSelector::new(
                        9, "old.iwa"
                    ))),
                    options(&versioned_only),
                )
                .unwrap_err()
            ),
            InvalidReason::ComponentMismatch
        );
        let mut wrong_wire = source.clone();
        put_key(&mut wrong_wire, 8, 2);
        bytes_field(&mut wrong_wire, 8, &[1]);
        assert_eq!(
            reason(
                rewrite_package_metadata_save_tokens(
                    &wrong_wire,
                    SaveTokenBatch::new(core::slice::from_ref(&selector)),
                    options(&wrong_wire),
                )
                .unwrap_err()
            ),
            InvalidReason::MalformedWire
        );
        let mut duplicate_token = source.clone();
        put_varint_field(&mut duplicate_token, 8, 5);
        assert_eq!(
            reason(
                rewrite_package_metadata_save_tokens(
                    &duplicate_token,
                    SaveTokenBatch::new(core::slice::from_ref(&selector)),
                    options(&duplicate_token),
                )
                .unwrap_err()
            ),
            InvalidReason::DuplicateSaveToken
        );
        let too_new = token_metadata(
            77,
            Some(5),
            core::slice::from_ref(&token_component(1, "a.iwa", Some(6), false, false)),
            &[],
        );
        assert_eq!(
            reason(
                rewrite_package_metadata_save_tokens(
                    &too_new,
                    SaveTokenBatch::new(core::slice::from_ref(&selector)),
                    options(&too_new),
                )
                .unwrap_err()
            ),
            InvalidReason::SaveTokenMismatch
        );
        let overflow = token_metadata(
            77,
            Some(u64::MAX),
            core::slice::from_ref(&token_component(1, "a.iwa", Some(u64::MAX), false, false)),
            &[],
        );
        assert_eq!(
            reason(
                rewrite_package_metadata_save_tokens(
                    &overflow,
                    SaveTokenBatch::new(core::slice::from_ref(&selector)),
                    options(&overflow),
                )
                .unwrap_err()
            ),
            InvalidReason::SaveTokenOverflow
        );

        let empty = SaveTokenBatch::new(&[]);
        assert_eq!(
            reason(
                rewrite_package_metadata_save_tokens(&source, empty, options(&source)).unwrap_err()
            ),
            InvalidReason::ComponentMismatch
        );

        let duplicate_id = token_metadata(
            77,
            Some(5),
            &[
                token_component(1, "a.iwa", Some(5), false, false),
                token_component(1, "b.iwa", Some(5), false, false),
            ],
            &[],
        );
        assert_eq!(
            reason(
                rewrite_package_metadata_save_tokens(
                    &duplicate_id,
                    SaveTokenBatch::new(core::slice::from_ref(&selector)),
                    options(&duplicate_id),
                )
                .unwrap_err()
            ),
            InvalidReason::ComponentMismatch
        );
        let duplicate_locator = token_metadata(
            77,
            Some(5),
            &[
                token_component(1, "a.iwa", Some(5), false, false),
                token_component(2, "a.iwa", Some(5), false, false),
            ],
            &[],
        );
        assert_eq!(
            reason(
                rewrite_package_metadata_save_tokens(
                    &duplicate_locator,
                    SaveTokenBatch::new(core::slice::from_ref(&selector)),
                    options(&duplicate_locator),
                )
                .unwrap_err()
            ),
            InvalidReason::ComponentMismatch
        );

        let explicit_locator_source = token_metadata(
            77,
            Some(5),
            core::slice::from_ref(&token_component_with_locator(
                1,
                "preferred.iwa",
                "effective.iwa",
                Some(5),
            )),
            &[],
        );
        let effective_selector = ComponentSelector::new(1, "effective.iwa");
        assert!(
            rewrite_package_metadata_save_tokens(
                &explicit_locator_source,
                SaveTokenBatch::new(core::slice::from_ref(&effective_selector)),
                options(&explicit_locator_source),
            )
            .is_ok()
        );
        let preferred_selector = ComponentSelector::new(1, "preferred.iwa");
        assert_eq!(
            reason(
                rewrite_package_metadata_save_tokens(
                    &explicit_locator_source,
                    SaveTokenBatch::new(core::slice::from_ref(&preferred_selector)),
                    options(&explicit_locator_source),
                )
                .unwrap_err()
            ),
            InvalidReason::ComponentMismatch
        );

        let mut noncanonical_root =
            token_metadata(77, None, core::slice::from_ref(&component), &[]);
        put_key(&mut noncanonical_root, 8, 0);
        noncanonical_root.extend_from_slice(&[0x81, 0x00]);
        assert_eq!(
            reason(
                rewrite_package_metadata_save_tokens(
                    &noncanonical_root,
                    SaveTokenBatch::new(core::slice::from_ref(&selector)),
                    options(&noncanonical_root),
                )
                .unwrap_err()
            ),
            InvalidReason::MalformedWire
        );

        let mut wrong_component = token_component(1, "a.iwa", None, false, false);
        bytes_field(&mut wrong_component, 12, &[1]);
        let wrong_component_source =
            token_metadata(77, Some(5), core::slice::from_ref(&wrong_component), &[]);
        assert_eq!(
            reason(
                rewrite_package_metadata_save_tokens(
                    &wrong_component_source,
                    SaveTokenBatch::new(core::slice::from_ref(&selector)),
                    options(&wrong_component_source),
                )
                .unwrap_err()
            ),
            InvalidReason::MalformedWire
        );

        let mut noncanonical_component = token_component(1, "a.iwa", None, false, false);
        put_key(&mut noncanonical_component, 12, 0);
        noncanonical_component.extend_from_slice(&[0x81, 0x00]);
        let noncanonical_component_source = token_metadata(
            77,
            Some(5),
            core::slice::from_ref(&noncanonical_component),
            &[],
        );
        assert_eq!(
            reason(
                rewrite_package_metadata_save_tokens(
                    &noncanonical_component_source,
                    SaveTokenBatch::new(core::slice::from_ref(&selector)),
                    options(&noncanonical_component_source),
                )
                .unwrap_err()
            ),
            InvalidReason::MalformedWire
        );

        let mut duplicate_component = token_component(1, "a.iwa", Some(5), false, false);
        put_varint_field(&mut duplicate_component, 12, 5);
        let duplicate_component_source = token_metadata(
            77,
            Some(5),
            core::slice::from_ref(&duplicate_component),
            &[],
        );
        assert_eq!(
            reason(
                rewrite_package_metadata_save_tokens(
                    &duplicate_component_source,
                    SaveTokenBatch::new(core::slice::from_ref(&selector)),
                    options(&duplicate_component_source),
                )
                .unwrap_err()
            ),
            InvalidReason::DuplicateSaveToken
        );

        let mut grouped_component = token_component(1, "a.iwa", None, false, false);
        put_key(&mut grouped_component, 12, 3);
        put_varint_field(&mut grouped_component, 1, 0);
        put_key(&mut grouped_component, 12, 4);
        let grouped_component_source =
            token_metadata(77, Some(5), core::slice::from_ref(&grouped_component), &[]);
        assert_eq!(
            reason(
                rewrite_package_metadata_save_tokens(
                    &grouped_component_source,
                    SaveTokenBatch::new(core::slice::from_ref(&selector)),
                    options(&grouped_component_source),
                )
                .unwrap_err()
            ),
            InvalidReason::MalformedWire
        );
    }

    #[test]
    fn save_tokens_limits_fail_before_output_allocation() {
        let component = token_component(1, "a.iwa", Some(536), false, false);
        let versioned = token_component(1, "old.iwa", Some(536), false, true);
        let source = token_metadata(
            77,
            Some(536),
            core::slice::from_ref(&component),
            core::slice::from_ref(&versioned),
        );
        let selector = ComponentSelector::new(1, "a.iwa");
        let batch = SaveTokenBatch::new(core::slice::from_ref(&selector));
        let baseline =
            rewrite_package_metadata_save_tokens(&source, batch, options(&source)).unwrap();
        let report = baseline.report();
        assert_eq!(report.additions(), 0);
        assert_eq!(report.output_bytes(), baseline.bytes().len());
        assert_eq!(report.retained_bytes(), baseline.bytes().len());
        assert!(report.allocations() >= 3);
        let exact = RewriteOptions::new(
            source.len(),
            report.output_bytes(),
            report.fields(),
            report.work_bytes(),
            report.max_depth(),
            report.components_scanned(),
            report.references_scanned(),
            0,
        );
        let exact_output = rewrite_package_metadata_save_tokens(&source, batch, exact).unwrap();
        assert_eq!(exact_output.bytes(), baseline.bytes());
        let limited = RewriteOptions::new(
            source.len(),
            report.output_bytes() - 1,
            report.fields(),
            report.work_bytes(),
            report.max_depth(),
            report.components_scanned(),
            report.references_scanned(),
            0,
        );
        let allocations = output_allocations();
        let error = rewrite_package_metadata_save_tokens(&source, batch, limited).unwrap_err();
        assert!(matches!(
            error.resource_limit(),
            Some(RewriteLimit::OutputBytes { .. })
        ));
        assert_eq!(output_allocations(), allocations);
        let work_limited = RewriteOptions::new(
            source.len(),
            report.output_bytes(),
            report.fields(),
            report.work_bytes() - 1,
            report.max_depth(),
            report.components_scanned(),
            report.references_scanned(),
            0,
        );
        let allocations = output_allocations();
        let error = rewrite_package_metadata_save_tokens(&source, batch, work_limited).unwrap_err();
        assert!(matches!(
            error.resource_limit(),
            Some(RewriteLimit::Work { .. })
        ));
        assert_eq!(output_allocations(), allocations);

        for (fields, components, depth) in [
            (
                report.fields() - 1,
                report.components_scanned(),
                report.max_depth(),
            ),
            (
                report.fields(),
                report.components_scanned() - 1,
                report.max_depth(),
            ),
            (
                report.fields(),
                report.components_scanned(),
                report.max_depth() - 1,
            ),
        ] {
            let limited = RewriteOptions::new(
                source.len(),
                report.output_bytes(),
                fields,
                report.work_bytes(),
                depth,
                components,
                report.references_scanned(),
                0,
            );
            let allocations = output_allocations();
            let error = rewrite_package_metadata_save_tokens(&source, batch, limited).unwrap_err();
            assert!(error.resource_limit().is_some());
            assert_eq!(output_allocations(), allocations);
        }
    }

    #[test]
    fn save_tokens_exact_work_includes_candidate_selector_matching() {
        let locators = [
            "a.iwa", "b.iwa", "c.iwa", "d.iwa", "e.iwa", "f.iwa", "g.iwa", "h.iwa",
        ];
        let components: Vec<Vec<u8>> = locators
            .iter()
            .enumerate()
            .map(|(index, locator)| {
                token_component((index + 1) as u64, locator, Some(536), false, false)
            })
            .collect();
        let selectors: Vec<ComponentSelector<'_>> = locators
            .iter()
            .enumerate()
            .map(|(index, locator)| ComponentSelector::new((index + 1) as u64, locator))
            .collect();
        let source = token_metadata(77, Some(536), &components, &[]);
        let batch = SaveTokenBatch::new(&selectors);
        reset_work_charges();
        let baseline =
            rewrite_package_metadata_save_tokens(&source, batch, options(&source)).unwrap();
        let report = baseline.report();
        assert_eq!(work_charges(), report.work_bytes());
        let exact = RewriteOptions::new(
            source.len(),
            report.output_bytes(),
            report.fields(),
            report.work_bytes(),
            report.max_depth(),
            report.components_scanned(),
            report.references_scanned(),
            0,
        );
        reset_work_charges();
        let replay = rewrite_package_metadata_save_tokens(&source, batch, exact).unwrap();
        assert_eq!(replay.bytes(), baseline.bytes());
        assert_eq!(work_charges(), replay.report().work_bytes());

        let limited = RewriteOptions::new(
            source.len(),
            report.output_bytes(),
            report.fields(),
            report.work_bytes() - 1,
            report.max_depth(),
            report.components_scanned(),
            report.references_scanned(),
            0,
        );
        let allocations = output_allocations();
        let error = rewrite_package_metadata_save_tokens(&source, batch, limited).unwrap_err();
        assert!(matches!(
            error.resource_limit(),
            Some(RewriteLimit::Work { .. })
        ));
        assert_eq!(output_allocations(), allocations);
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

#[derive(Default, Clone, Copy)]
struct SaveTokenMatch {
    identifier: usize,
    locator: usize,
    exact: usize,
    token: Option<u64>,
}

#[derive(Clone, Copy)]
struct RawFieldBytes {
    bytes: [u8; 11],
    len: usize,
}

impl RawFieldBytes {
    fn from_slice(source: &[u8]) -> Result<Self, RewriteError> {
        let len = source.len();
        if len > 11 {
            return Err(RewriteError::invalid(InvalidReason::Verification));
        }
        let mut bytes = [0; 11];
        bytes[..len].copy_from_slice(source);
        Ok(Self { bytes, len })
    }

    fn matches(self, source: &[u8]) -> bool {
        self.len == source.len() && self.bytes[..self.len] == *source
    }
}

struct SaveTokenScanState {
    selectors: Vec<SaveTokenMatch>,
    last: Option<u64>,
    last_raw: Option<RawFieldBytes>,
    root_token: Option<u64>,
}

impl SaveTokenScanState {
    fn new(batch: SaveTokenBatch<'_>, budget: &mut Budget) -> Result<Self, RewriteError> {
        Ok(Self {
            selectors: zeroed_vec(batch.components.len(), budget)?,
            last: None,
            last_raw: None,
            root_token: None,
        })
    }

    fn validate_source(&self) -> Result<u64, RewriteError> {
        let root = self.root_token.unwrap_or(0);
        root.checked_add(1)
            .ok_or_else(|| RewriteError::invalid(InvalidReason::SaveTokenOverflow))?;
        if self.last.is_none_or(|value| value == 0) {
            return Err(RewriteError::invalid(InvalidReason::InvalidIdentifier));
        }
        for matched in &self.selectors {
            if matched.identifier != 1 || matched.locator != 1 || matched.exact != 1 {
                return Err(RewriteError::invalid(InvalidReason::ComponentMismatch));
            }
            if matched.token.unwrap_or(0) > root {
                return Err(RewriteError::invalid(InvalidReason::SaveTokenMismatch));
            }
        }
        Ok(root)
    }

    fn validate_candidate(&self, root: u64) -> Result<(), RewriteError> {
        if self.last.is_none_or(|value| value == 0)
            || self.root_token != Some(root)
            || self.selectors.iter().any(|matched| {
                matched.identifier != 1
                    || matched.locator != 1
                    || matched.exact != 1
                    || matched.token != Some(root)
            })
        {
            return Err(RewriteError::invalid(InvalidReason::Verification));
        }
        Ok(())
    }
}

/// Rewrite the package root save token and the selected current component
/// save tokens in one raw-preserving atomic candidate.
pub fn rewrite_package_metadata_save_tokens(
    source: &[u8],
    batch: SaveTokenBatch<'_>,
    options: RewriteOptions,
) -> Result<RewriteOutput, RewriteError> {
    validate_save_token_batch(batch, options)?;
    let mut budget = Budget::new_inspection(source, options)?;
    validate_save_token_selector_duplicates(batch, &mut budget)?;
    let mut source_state = SaveTokenScanState::new(batch, &mut budget)?;
    scan_save_token_metadata(
        source,
        batch,
        SaveTokenScanMode::Source,
        &mut source_state,
        None,
        &mut budget,
    )?;
    let old_root = source_state.validate_source()?;
    let new_root = old_root
        .checked_add(1)
        .ok_or_else(|| RewriteError::invalid(InvalidReason::SaveTokenOverflow))?;

    let output_size = save_token_output_size(source, batch, new_root, &mut budget)?;
    budget.output_size(output_size)?;
    let before_execution = budget.clone();
    precharge_save_token_rewrite_and_verification(
        source,
        batch,
        new_root,
        output_size,
        &mut budget,
    )?;
    budget.preflight_repeat_delta(&before_execution)?;

    let mut candidate = Vec::new();
    #[cfg(test)]
    record_output_allocation();
    candidate
        .try_reserve_exact(output_size)
        .map_err(|_error| RewriteError::allocation(output_size))?;
    if candidate.capacity() != output_size {
        return Err(RewriteError::allocation(output_size));
    }
    budget.allocation(0)?;
    rewrite_save_token_into(source, batch, new_root, &mut candidate, &mut budget)?;
    if candidate.len() != output_size {
        return Err(RewriteError::invalid(InvalidReason::Verification));
    }

    budget.source_phase = false;
    let mut candidate_state = SaveTokenScanState::new(batch, &mut budget)?;
    scan_save_token_metadata(
        &candidate,
        batch,
        SaveTokenScanMode::Verification,
        &mut candidate_state,
        Some((
            source_state
                .last_raw
                .ok_or_else(|| RewriteError::invalid(InvalidReason::Verification))?,
            source_state
                .last
                .ok_or_else(|| RewriteError::invalid(InvalidReason::Verification))?,
            new_root,
        )),
        &mut budget,
    )?;
    candidate_state.validate_candidate(new_root)?;
    budget.output_bytes = candidate.len();
    budget.retained_bytes = candidate.len();
    Ok(RewriteOutput {
        bytes: candidate,
        report: budget.report(),
    })
}

/// Atomically remove exact registry ownership records and advance the root
/// plus selected current-component save tokens.
///
/// The source is scanned for both transitions before the sole output buffer
/// is reserved.  Field 1 and every unselected/versioned record remain raw;
/// the only scalar additions/replacements are root field 8 and selected
/// current-component field 12.  A removal that would cross a versioned or
/// ambiguous owner is rejected by the same strict ownership scanner used by
/// [`remove_package_metadata`].
pub fn rewrite_package_metadata_removals_and_save_tokens(
    source: &[u8],
    batch: RemovalSaveTokenBatch<'_>,
    options: RewriteOptions,
) -> Result<RewriteOutput, RewriteError> {
    validate_removal_batch(batch.removals, options)?;
    validate_save_token_batch(batch.save_tokens, options)?;

    let mut budget = Budget::new_inspection(source, options)?;
    budget.removals = batch
        .removals
        .object_uuids
        .len()
        .checked_add(batch.removals.external_references.len())
        .and_then(|count| count.checked_add(batch.removals.data_reference_owners.len()))
        .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?;
    validate_removal_batch_duplicates(batch.removals, &mut budget)?;
    validate_save_token_selector_duplicates(batch.save_tokens, &mut budget)?;
    validate_combined_selector_coverage(batch.removals, batch.save_tokens, &mut budget)?;

    let mut removal_state = RemovalScanState::new(batch.removals, &mut budget)?;
    scan_removal_metadata(
        source,
        batch.removals,
        &mut removal_state,
        &mut budget,
        false,
    )?;
    removal_state.validate_source()?;

    let mut save_state = SaveTokenScanState::new(batch.save_tokens, &mut budget)?;
    scan_save_token_metadata(
        source,
        batch.save_tokens,
        SaveTokenScanMode::Source,
        &mut save_state,
        None,
        &mut budget,
    )?;
    let old_root = save_state.validate_source()?;
    if old_root == u64::MAX
        || save_state.last != Some(batch.removals.expected_last_object_identifier)
    {
        return Err(RewriteError::invalid(if old_root == u64::MAX {
            InvalidReason::SaveTokenOverflow
        } else {
            InvalidReason::LastIdentifierMismatch
        }));
    }
    let new_root = old_root
        .checked_add(1)
        .ok_or_else(|| RewriteError::invalid(InvalidReason::SaveTokenOverflow))?;

    let output_size = combined_output_size(source, batch, new_root, &mut budget)?;
    budget.output_size(output_size)?;

    // Charge the complete rewrite and a conservative candidate verification
    // before reserving the sole output buffer.  Candidate scans use the
    // source shape as an upper bound (removals only shrink it); inserted token
    // fields and the output-size delta are charged explicitly.
    let measured = budget.clone();
    budget.source_phase = false;
    charge_combined_rewrite(source, batch, new_root, &mut budget)?;
    charge_combined_candidate_verification(source, batch, new_root, output_size, &mut budget)?;
    budget.preflight_repeat_delta(&measured)?;
    let planned_fields = repeated_counter(measured.fields, budget.fields)?;
    let planned_work = repeated_counter(measured.work_bytes, budget.work_bytes)?;
    let planned_components =
        repeated_counter(measured.components_scanned, budget.components_scanned)?;
    let planned_references =
        repeated_counter(measured.references_scanned, budget.references_scanned)?;

    let mut candidate = Vec::new();
    #[cfg(test)]
    record_output_allocation();
    candidate
        .try_reserve_exact(output_size)
        .map_err(|_error| RewriteError::allocation(output_size))?;
    if candidate.capacity() != output_size {
        return Err(RewriteError::allocation(output_size));
    }
    budget.allocation(0)?;
    rewrite_combined_into(source, batch, new_root, &mut candidate, &mut budget)?;
    if candidate.len() != output_size {
        return Err(RewriteError::invalid(InvalidReason::Verification));
    }

    budget.source_phase = false;
    let mut verified_removals = RemovalScanState::new(batch.removals, &mut budget)?;
    scan_removal_metadata(
        &candidate,
        batch.removals,
        &mut verified_removals,
        &mut budget,
        true,
    )?;
    verified_removals.validate_candidate()?;

    let mut verified_tokens = SaveTokenScanState::new(batch.save_tokens, &mut budget)?;
    scan_save_token_metadata(
        &candidate,
        batch.save_tokens,
        SaveTokenScanMode::Verification,
        &mut verified_tokens,
        Some((
            save_state
                .last_raw
                .ok_or_else(|| RewriteError::invalid(InvalidReason::Verification))?,
            save_state
                .last
                .ok_or_else(|| RewriteError::invalid(InvalidReason::Verification))?,
            new_root,
        )),
        &mut budget,
    )?;
    verified_tokens.validate_candidate(new_root)?;
    budget.pad_repeated_counters(
        planned_fields,
        planned_work,
        planned_components,
        planned_references,
    )?;
    budget.output_bytes = candidate.len();
    budget.retained_bytes = candidate.len();
    Ok(RewriteOutput {
        bytes: candidate,
        report: budget.report(),
    })
}

fn validate_save_token_batch(
    batch: SaveTokenBatch<'_>,
    options: RewriteOptions,
) -> Result<(), RewriteError> {
    if batch.components.len() > options.max_components {
        return Err(RewriteError::limited(RewriteLimit::Components {
            observed: batch.components.len(),
            maximum: options.max_components,
        }));
    }
    if batch.components.is_empty() {
        return Err(RewriteError::invalid(InvalidReason::ComponentMismatch));
    }
    for selector in batch.components.iter().copied() {
        validate_selector(selector)?;
    }
    Ok(())
}

fn validate_save_token_selector_duplicates(
    batch: SaveTokenBatch<'_>,
    budget: &mut Budget,
) -> Result<(), RewriteError> {
    for (index, selector) in batch.components.iter().copied().enumerate() {
        for prior in batch.components[..index].iter().copied() {
            budget.work(
                prior
                    .locator
                    .len()
                    .checked_add(1)
                    .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?,
            )?;
            if prior == selector {
                return Err(RewriteError::invalid(InvalidReason::DuplicateSelector));
            }
        }
    }
    Ok(())
}

#[derive(Clone, Copy)]
enum SaveTokenScanMode {
    Source,
    Verification,
}

fn scan_save_token_metadata(
    source: &[u8],
    batch: SaveTokenBatch<'_>,
    mode: SaveTokenScanMode,
    state: &mut SaveTokenScanState,
    expected: Option<(RawFieldBytes, u64, u64)>,
    budget: &mut Budget,
) -> Result<(), RewriteError> {
    budget.message(source, 1)?;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, 1)? {
        match field.number {
            1 => {
                let value = field.varint()?;
                set_once(&mut state.last, value)?;
                if let Some((source_last_raw, source_last, _)) = expected {
                    if !source_last_raw.matches(field.raw) || value != source_last {
                        return Err(RewriteError::invalid(InvalidReason::Verification));
                    }
                } else {
                    state.last_raw = Some(RawFieldBytes::from_slice(field.raw)?);
                }
            },
            3 | 11 => scan_save_token_component(
                field.bytes()?,
                field.number == 3,
                batch,
                mode,
                state,
                expected.map(|(_, _, root)| root),
                budget,
                2,
            )?,
            8 => {
                let value = field.varint()?;
                if state.root_token.replace(value).is_some() {
                    return Err(RewriteError::invalid(InvalidReason::DuplicateSaveToken));
                }
            },
            _ => {},
        }
    }

    if state.last.is_none_or(|value| value == 0) {
        return Err(RewriteError::invalid(match mode {
            SaveTokenScanMode::Source => InvalidReason::InvalidIdentifier,
            SaveTokenScanMode::Verification => InvalidReason::Verification,
        }));
    }
    if let Some((_, source_last, expected_root)) = expected {
        if state.last != Some(source_last) || state.root_token != Some(expected_root) {
            return Err(RewriteError::invalid(InvalidReason::Verification));
        }
    }

    let view: projection::PackageMetadataArchiveLazyView<'_> = budget
        .options
        .buffa()
        .decode_lazy_view(source)
        .map_err(|_error| RewriteError::invalid(InvalidReason::MalformedWire))?;
    if !view.has_last_object_identifier() || view.last_object_identifier != state.last.unwrap() {
        return Err(RewriteError::invalid(InvalidReason::MalformedWire));
    }
    if view.save_token != state.root_token {
        return Err(RewriteError::invalid(InvalidReason::MalformedWire));
    }
    Ok(())
}

fn scan_save_token_component<'source>(
    source: &'source [u8],
    current: bool,
    batch: SaveTokenBatch<'source>,
    mode: SaveTokenScanMode,
    state: &mut SaveTokenScanState,
    expected_root: Option<u64>,
    budget: &mut Budget,
    depth: u32,
) -> Result<(), RewriteError> {
    budget.component()?;
    budget.message(source, depth)?;
    let mut identifier = None;
    let mut preferred_locator = None;
    let mut locator = None;
    let mut token = None;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        match field.number {
            1 => set_once(&mut identifier, field.varint()?)?,
            2 => set_once(&mut preferred_locator, strict_utf8(field.bytes()?)?)?,
            3 => set_once(&mut locator, strict_utf8(field.bytes()?)?)?,
            12 => {
                let value = field.varint()?;
                if token.replace(value).is_some() {
                    return Err(RewriteError::invalid(InvalidReason::DuplicateSaveToken));
                }
            },
            _ => {},
        }
    }
    let identifier = identifier
        .filter(|value| *value != 0)
        .ok_or_else(|| RewriteError::invalid(InvalidReason::InvalidIdentifier))?;
    let preferred_locator =
        preferred_locator.ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?;
    let effective_locator = locator.unwrap_or(preferred_locator);

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
        || view.save_token != token
    {
        return Err(RewriteError::invalid(InvalidReason::MalformedWire));
    }

    if !current {
        return Ok(());
    }

    for (index, selector) in batch.components.iter().copied().enumerate() {
        budget.work(
            selector
                .locator
                .len()
                .checked_add(1)
                .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?,
        )?;
        let matched = &mut state.selectors[index];
        if identifier == selector.identifier {
            matched.identifier = checked_add(matched.identifier, 1)?;
        }
        if effective_locator == selector.locator {
            matched.locator = checked_add(matched.locator, 1)?;
        }
        if identifier == selector.identifier && effective_locator == selector.locator {
            matched.exact = checked_add(matched.exact, 1)?;
            matched.token = token;
            if matches!(mode, SaveTokenScanMode::Verification) && token != expected_root {
                return Err(RewriteError::invalid(InvalidReason::Verification));
            }
        }
    }
    Ok(())
}

fn save_token_output_size(
    source: &[u8],
    batch: SaveTokenBatch<'_>,
    new_root: u64,
    budget: &mut Budget,
) -> Result<usize, RewriteError> {
    budget.message(source, 1)?;
    let mut output = 0usize;
    let mut has_root_token = false;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, 1)? {
        match field.number {
            3 | 11 => {
                let payload = field.bytes()?;
                let (new_len, _selected) = save_token_component_size(
                    payload,
                    field.number == 3,
                    batch,
                    new_root,
                    budget,
                    2,
                )?;
                output = checked_add(
                    output,
                    if new_len == payload.len() {
                        field.raw.len()
                    } else {
                        length_delimited_field_len(field.number, new_len)?
                    },
                )?;
            },
            8 => {
                let _ = field.varint()?;
                has_root_token = true;
                output = checked_add(output, varint_field_len(8, new_root))?;
            },
            _ => output = checked_add(output, field.raw.len())?,
        }
    }
    if !has_root_token {
        output = checked_add(output, varint_field_len(8, new_root))?;
    }
    Ok(output)
}

fn save_token_component_size(
    source: &[u8],
    current: bool,
    batch: SaveTokenBatch<'_>,
    new_root: u64,
    budget: &mut Budget,
    depth: u32,
) -> Result<(usize, bool), RewriteError> {
    budget.message(source, depth)?;
    let mut output = 0usize;
    let mut identifier = None;
    let mut preferred_locator = None;
    let mut locator = None;
    let mut token_raw = None;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        match field.number {
            1 => set_once(&mut identifier, field.varint()?)?,
            2 => set_once(&mut preferred_locator, strict_utf8(field.bytes()?)?)?,
            3 => set_once(&mut locator, strict_utf8(field.bytes()?)?)?,
            12 => {
                let _ = field.varint()?;
                if token_raw.replace(field.raw).is_some() {
                    return Err(RewriteError::invalid(InvalidReason::DuplicateSaveToken));
                }
            },
            _ => {},
        }
        output = checked_add(output, field.raw.len())?;
    }
    if !current {
        return Ok((output, false));
    }
    let identifier = identifier
        .filter(|value| *value != 0)
        .ok_or_else(|| RewriteError::invalid(InvalidReason::InvalidIdentifier))?;
    let preferred_locator =
        preferred_locator.ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?;
    let effective_locator = locator.unwrap_or(preferred_locator);
    let mut selected = false;
    for selector in batch.components.iter().copied() {
        budget.work(
            selector
                .locator
                .len()
                .checked_add(1)
                .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?,
        )?;
        selected |= selector.identifier == identifier && selector.locator == effective_locator;
    }
    if selected {
        if let Some(raw) = token_raw {
            output = output
                .checked_sub(raw.len())
                .ok_or_else(|| RewriteError::invalid(InvalidReason::Verification))?;
        }
        output = checked_add(output, varint_field_len(12, new_root))?;
    }
    Ok((output, selected))
}

fn save_token_component_output_pass(
    source: &[u8],
    budget: &mut Budget,
    depth: u32,
) -> Result<(), RewriteError> {
    budget.message(source, depth)?;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        match field.number {
            1 => {
                let _ = field.varint()?;
            },
            2 | 3 => {
                let _ = strict_utf8(field.bytes()?)?;
            },
            12 => {
                let _ = field.varint()?;
            },
            _ => {},
        }
    }
    Ok(())
}

fn precharge_save_token_candidate_component(
    source: &[u8],
    current: bool,
    batch: SaveTokenBatch<'_>,
    new_root: u64,
    budget: &mut Budget,
    depth: u32,
) -> Result<(usize, bool), RewriteError> {
    budget.component()?;
    let mut output = 0usize;
    let mut identifier = None;
    let mut preferred_locator = None;
    let mut locator = None;
    let mut token_raw = None;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        match field.number {
            1 => set_once(&mut identifier, field.varint()?)?,
            2 => set_once(&mut preferred_locator, strict_utf8(field.bytes()?)?)?,
            3 => set_once(&mut locator, strict_utf8(field.bytes()?)?)?,
            12 => {
                let value = field.varint()?;
                if token_raw.replace((value, field.raw)).is_some() {
                    return Err(RewriteError::invalid(InvalidReason::DuplicateSaveToken));
                }
            },
            _ => {},
        }
        output = checked_add(output, field.raw.len())?;
    }
    if !current {
        budget.message_len(output, depth)?;
        return Ok((output, false));
    }
    let identifier = identifier
        .filter(|value| *value != 0)
        .ok_or_else(|| RewriteError::invalid(InvalidReason::InvalidIdentifier))?;
    let preferred_locator =
        preferred_locator.ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?;
    let effective_locator = locator.unwrap_or(preferred_locator);
    let mut selected = false;
    for selector in batch.components.iter().copied() {
        budget.work(
            selector
                .locator
                .len()
                .checked_add(1)
                .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?,
        )?;
        selected |= identifier == selector.identifier && effective_locator == selector.locator;
    }
    if selected && token_raw.is_none() {
        budget.field()?;
    }
    if selected {
        if let Some((_token, raw)) = token_raw {
            output = output
                .checked_sub(raw.len())
                .ok_or_else(|| RewriteError::invalid(InvalidReason::Verification))?;
        }
        output = checked_add(output, varint_field_len(12, new_root))?;
    }
    budget.message_len(output, depth)?;
    Ok((output, selected))
}

fn precharge_save_token_rewrite_and_verification(
    source: &[u8],
    batch: SaveTokenBatch<'_>,
    new_root: u64,
    output_size: usize,
    budget: &mut Budget,
) -> Result<(), RewriteError> {
    let measured = budget.clone();
    budget.message(source, 1)?;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, 1)? {
        if field.number == 3 {
            let payload = field.bytes()?;
            let (_size, selected) =
                save_token_component_size(payload, true, batch, new_root, budget, 2)?;
            if selected {
                save_token_component_output_pass(payload, budget, 2)?;
            }
        }
    }

    budget.source_phase = false;
    budget.message_len(output_size, 1)?;
    let mut has_root_token = false;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, 1)? {
        if field.number == 8 {
            has_root_token = true;
        }
        if !matches!(field.number, 3 | 11) {
            continue;
        }
        let payload = field.bytes()?;
        let (_candidate_len, _selected) = precharge_save_token_candidate_component(
            payload,
            field.number == 3,
            batch,
            new_root,
            budget,
            2,
        )?;
    }
    if !has_root_token {
        budget.field()?;
    }
    budget.preflight_repeat_delta(&measured)?;
    Ok(())
}

fn rewrite_save_token_into(
    source: &[u8],
    batch: SaveTokenBatch<'_>,
    new_root: u64,
    output: &mut Vec<u8>,
    budget: &mut Budget,
) -> Result<(), RewriteError> {
    budget.message(source, 1)?;
    let mut has_root_token = false;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, 1)? {
        match field.number {
            8 => {
                let _ = field.varint()?;
                has_root_token = true;
                put_varint_field(output, 8, new_root);
            },
            3 => rewrite_save_token_component(field, true, batch, new_root, output, budget, 2)?,
            11 => output.extend_from_slice(field.raw),
            _ => output.extend_from_slice(field.raw),
        }
    }
    if !has_root_token {
        put_varint_field(output, 8, new_root);
    }
    Ok(())
}

fn rewrite_save_token_component(
    field: Field<'_>,
    current: bool,
    batch: SaveTokenBatch<'_>,
    new_root: u64,
    output: &mut Vec<u8>,
    budget: &mut Budget,
    depth: u32,
) -> Result<(), RewriteError> {
    let source = field.bytes()?;
    let (size, selected) =
        save_token_component_size(source, current, batch, new_root, budget, depth)?;
    if !selected {
        output.extend_from_slice(field.raw);
        return Ok(());
    }
    budget.changed_component()?;
    put_key(output, field.number, 2);
    put_varint(
        output,
        u64::try_from(size)
            .map_err(|_error| RewriteError::invalid(InvalidReason::MalformedWire))?,
    );
    budget.message(source, depth)?;
    let mut token_seen = false;
    let mut remaining = source;
    while let Some(nested) = next_field(&mut remaining, budget, depth)? {
        match nested.number {
            1 => {
                let _ = nested.varint()?;
                output.extend_from_slice(nested.raw);
            },
            2 => {
                let _ = strict_utf8(nested.bytes()?)?;
                output.extend_from_slice(nested.raw);
            },
            3 => {
                let _ = strict_utf8(nested.bytes()?)?;
                output.extend_from_slice(nested.raw);
            },
            12 => {
                let _ = nested.varint()?;
                token_seen = true;
                put_varint_field(output, 12, new_root);
            },
            _ => output.extend_from_slice(nested.raw),
        }
    }
    if selected && !token_seen {
        put_varint_field(output, 12, new_root);
    }
    Ok(())
}

fn combined_output_size(
    source: &[u8],
    batch: RemovalSaveTokenBatch<'_>,
    new_root: u64,
    budget: &mut Budget,
) -> Result<usize, RewriteError> {
    budget.message(source, 1)?;
    let mut output = 0usize;
    let mut has_root_token = false;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, 1)? {
        match field.number {
            3 => {
                let payload = field.bytes()?;
                let (identifier, locator) = component_header(payload, budget, 2)?;
                let (new_len, changed) = combined_component_size(
                    payload, identifier, locator, batch, new_root, true, budget, 2,
                )?;
                output = checked_add(
                    output,
                    if changed {
                        length_delimited_field_len(3, new_len)?
                    } else {
                        field.raw.len()
                    },
                )?;
            },
            8 => {
                let _ = field.varint()?;
                if has_root_token {
                    return Err(RewriteError::invalid(InvalidReason::DuplicateSaveToken));
                }
                has_root_token = true;
                output = checked_add(output, varint_field_len(8, new_root))?;
            },
            _ => output = checked_add(output, field.raw.len())?,
        }
    }
    if !has_root_token {
        output = checked_add(output, varint_field_len(8, new_root))?;
    }
    Ok(output)
}

fn combined_component_size(
    source: &[u8],
    component: u64,
    locator: &str,
    batch: RemovalSaveTokenBatch<'_>,
    new_root: u64,
    current: bool,
    budget: &mut Budget,
    depth: u32,
) -> Result<(usize, bool), RewriteError> {
    budget.message(source, depth)?;
    let selected_token =
        current && save_token_component_selected(component, locator, batch.save_tokens, budget)?;
    let mut output = 0usize;
    let mut token_seen = false;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        let keep = match field.number {
            6 => !external_field_selected(
                field.bytes()?,
                component,
                locator,
                batch.removals,
                budget,
                depth + 1,
            )?,
            7 => {
                let rewrite = data_reference_rewrite(
                    field.bytes()?,
                    component,
                    locator,
                    batch.removals,
                    budget,
                    depth + 1,
                )?;
                if rewrite.selected == 0 {
                    output = checked_add(output, field.raw.len())?;
                } else if rewrite.surviving_owners != 0 {
                    output =
                        checked_add(output, length_delimited_field_len(7, rewrite.payload_size)?)?;
                }
                false
            },
            11 => !object_field_selected(
                field.bytes()?,
                component,
                locator,
                batch.removals,
                budget,
                depth + 1,
            )?,
            12 if selected_token => {
                let _ = field.varint()?;
                if token_seen {
                    return Err(RewriteError::invalid(InvalidReason::DuplicateSaveToken));
                }
                token_seen = true;
                output = checked_add(output, varint_field_len(12, new_root))?;
                false
            },
            12 => {
                let _ = field.varint()?;
                if token_seen {
                    return Err(RewriteError::invalid(InvalidReason::DuplicateSaveToken));
                }
                token_seen = true;
                true
            },
            _ => true,
        };
        if keep {
            output = checked_add(output, field.raw.len())?;
        }
    }
    if selected_token && !token_seen {
        output = checked_add(output, varint_field_len(12, new_root))?;
    }
    let changed = selected_token || output != source.len();
    Ok((output, changed))
}

fn save_token_component_selected(
    component: u64,
    locator: &str,
    batch: SaveTokenBatch<'_>,
    budget: &mut Budget,
) -> Result<bool, RewriteError> {
    let mut selected = false;
    for selector in batch.components.iter().copied() {
        budget.work(
            selector
                .locator
                .len()
                .checked_add(1)
                .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?,
        )?;
        selected |= selector.identifier == component && selector.locator == locator;
    }
    Ok(selected)
}

fn charge_combined_rewrite(
    source: &[u8],
    batch: RemovalSaveTokenBatch<'_>,
    new_root: u64,
    budget: &mut Budget,
) -> Result<(), RewriteError> {
    budget.message(source, 1)?;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, 1)? {
        if field.number != 3 {
            continue;
        }
        let payload = field.bytes()?;
        let (identifier, locator) = component_header(payload, budget, 2)?;
        let (_size, changed) = combined_component_size(
            payload, identifier, locator, batch, new_root, true, budget, 2,
        )?;
        if changed {
            charge_combined_component_rewrite(
                payload,
                component_selector(identifier, locator),
                batch,
                new_root,
                budget,
                2,
            )?;
        }
    }
    Ok(())
}

fn component_selector<'source>(
    identifier: u64,
    locator: &'source str,
) -> ComponentSelector<'source> {
    ComponentSelector::new(identifier, locator)
}

fn charge_combined_component_rewrite(
    source: &[u8],
    component: ComponentSelector<'_>,
    batch: RemovalSaveTokenBatch<'_>,
    new_root: u64,
    budget: &mut Budget,
    depth: u32,
) -> Result<(), RewriteError> {
    budget.message(source, depth)?;
    let selected_token = save_token_component_selected(
        component.identifier,
        component.locator,
        batch.save_tokens,
        budget,
    )?;
    let mut token_seen = false;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        match field.number {
            6 => {
                let _ = external_field_selected(
                    field.bytes()?,
                    component.identifier,
                    component.locator,
                    batch.removals,
                    budget,
                    depth + 1,
                )?;
            },
            7 => {
                let _ = data_reference_rewrite(
                    field.bytes()?,
                    component.identifier,
                    component.locator,
                    batch.removals,
                    budget,
                    depth + 1,
                )?;
            },
            11 => {
                let _ = object_field_selected(
                    field.bytes()?,
                    component.identifier,
                    component.locator,
                    batch.removals,
                    budget,
                    depth + 1,
                )?;
            },
            12 => {
                let _ = field.varint()?;
                if selected_token {
                    if token_seen {
                        return Err(RewriteError::invalid(InvalidReason::DuplicateSaveToken));
                    }
                    token_seen = true;
                    let _ = varint_field_len(12, new_root);
                }
            },
            _ => {},
        }
    }
    if selected_token && !token_seen {
        budget.field()?;
    }
    Ok(())
}

fn charge_combined_candidate_verification(
    source: &[u8],
    batch: RemovalSaveTokenBatch<'_>,
    new_root: u64,
    output_size: usize,
    budget: &mut Budget,
) -> Result<(), RewriteError> {
    // The removal scan over the source is a conservative upper bound for the
    // candidate: removals can only delete fields and references.
    let mut removal_state = RemovalScanState::new(batch.removals, budget)?;
    scan_removal_metadata(source, batch.removals, &mut removal_state, budget, true)?;

    // Save-token candidate work has the same component graph plus possible
    // canonical field-8/12 insertions.  Use the strict source parser for
    // Buffa parity and charge the output-size delta for inserted bytes.
    let mut save_state = SaveTokenScanState::new(batch.save_tokens, budget)?;
    scan_save_token_metadata(
        source,
        batch.save_tokens,
        SaveTokenScanMode::Source,
        &mut save_state,
        None,
        budget,
    )?;
    if output_size > source.len() {
        budget.work(output_size - source.len())?;
    }
    let _ = new_root;
    Ok(())
}

fn rewrite_combined_into(
    source: &[u8],
    batch: RemovalSaveTokenBatch<'_>,
    new_root: u64,
    output: &mut Vec<u8>,
    budget: &mut Budget,
) -> Result<(), RewriteError> {
    budget.message(source, 1)?;
    let mut has_root_token = false;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, 1)? {
        match field.number {
            3 => rewrite_combined_component(field, batch, new_root, output, budget, 2)?,
            8 => {
                let _ = field.varint()?;
                has_root_token = true;
                put_varint_field(output, 8, new_root);
            },
            _ => output.extend_from_slice(field.raw),
        }
    }
    if !has_root_token {
        put_varint_field(output, 8, new_root);
    }
    Ok(())
}

fn rewrite_combined_component(
    field: Field<'_>,
    batch: RemovalSaveTokenBatch<'_>,
    new_root: u64,
    output: &mut Vec<u8>,
    budget: &mut Budget,
    depth: u32,
) -> Result<(), RewriteError> {
    let source = field.bytes()?;
    let (identifier, locator) = component_header(source, budget, depth)?;
    let (size, changed) = combined_component_size(
        source, identifier, locator, batch, new_root, true, budget, depth,
    )?;
    if !changed {
        output.extend_from_slice(field.raw);
        return Ok(());
    }
    budget.changed_component()?;
    put_key(output, field.number, 2);
    put_varint(
        output,
        u64::try_from(size)
            .map_err(|_error| RewriteError::invalid(InvalidReason::MalformedWire))?,
    );
    rewrite_combined_component_payload(
        source, identifier, locator, batch, new_root, output, budget, depth,
    )?;
    Ok(())
}

fn rewrite_combined_component_payload(
    source: &[u8],
    component: u64,
    locator: &str,
    batch: RemovalSaveTokenBatch<'_>,
    new_root: u64,
    output: &mut Vec<u8>,
    budget: &mut Budget,
    depth: u32,
) -> Result<(), RewriteError> {
    budget.message(source, depth)?;
    let selected_token =
        save_token_component_selected(component, locator, batch.save_tokens, budget)?;
    let mut token_seen = false;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        match field.number {
            6 if external_field_selected(
                field.bytes()?,
                component,
                locator,
                batch.removals,
                budget,
                depth + 1,
            )? => {},
            7 => rewrite_data_reference_field(
                field,
                component,
                locator,
                batch.removals,
                output,
                budget,
                depth + 1,
            )?,
            11 if object_field_selected(
                field.bytes()?,
                component,
                locator,
                batch.removals,
                budget,
                depth + 1,
            )? => {},
            12 if selected_token => {
                let _ = field.varint()?;
                if token_seen {
                    return Err(RewriteError::invalid(InvalidReason::DuplicateSaveToken));
                }
                token_seen = true;
                put_varint_field(output, 12, new_root);
            },
            _ => output.extend_from_slice(field.raw),
        }
    }
    if selected_token && !token_seen {
        put_varint_field(output, 12, new_root);
    }
    Ok(())
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
    let prepared = prepare_package_metadata_rewrite(source, batch, options)?;
    let prepare_report = prepared.prepare_report();
    let limits = prepared.execution_requirements().exact_limits();
    let mut output = prepared.execute(limits)?;
    output.report = add_reports(prepare_report, output.report)?;
    Ok(output)
}

/// Validate and size one rewrite without allocating candidate output.
pub fn prepare_package_metadata_rewrite<'source, 'batch>(
    source: &'source [u8],
    batch: Batch<'batch>,
    options: RewriteOptions,
) -> Result<PreparedPackageMetadataRewrite<'source, 'batch>, RewriteError> {
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
    drop(source_state);

    let output_size = exact_output_size(source, batch, &mut budget)?;
    budget.output_size(output_size)?;
    let before_execution = budget.report();
    precharge_rewrite_and_verification(source, batch, output_size, &mut budget)?;
    let predicted = subtract_report(budget.report(), before_execution)?;
    let verification_scratch = scan_state_scratch_bytes(batch)?;
    let verification_allocations = scan_state_allocations(batch)?;
    let requirements = RewriteExecutionRequirements {
        output_bytes: output_size,
        fields: predicted.fields,
        work_bytes: predicted.work_bytes,
        components: predicted.components_scanned,
        references: predicted.references_scanned,
        allocations: verification_allocations
            .checked_add(1)
            .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?,
        retained_bytes: output_size,
        scratch_bytes: output_size
            .checked_add(verification_scratch)
            .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?,
    };
    let prepare_report = budget.report();
    Ok(PreparedPackageMetadataRewrite {
        source,
        batch,
        budget,
        prepare_report,
        requirements,
        output_size,
    })
}

fn scan_state_scratch_bytes(batch: Batch<'_>) -> Result<usize, RewriteError> {
    selector_count(batch)
        .checked_mul(size_of::<SelectorCount>())
        .and_then(|bytes| {
            batch
                .object_uuids
                .len()
                .checked_add(batch.external_references.len())
                .and_then(|matches| matches.checked_mul(size_of::<usize>()))
                .and_then(|matches| bytes.checked_add(matches))
        })
        .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))
}

fn scan_state_allocations(batch: Batch<'_>) -> Result<usize, RewriteError> {
    [
        selector_count(batch),
        batch.object_uuids.len(),
        batch.external_references.len(),
    ]
    .into_iter()
    .try_fold(0usize, |total, amount| {
        total
            .checked_add(usize::from(amount != 0))
            .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))
    })
}

fn preflight_execution(
    requirements: RewriteExecutionRequirements,
    limits: RewriteExecutionLimits,
) -> Result<(), RewriteError> {
    macro_rules! limited {
        ($field:ident, $maximum:ident, $variant:ident) => {
            if requirements.$field > limits.$maximum {
                return Err(RewriteError::limited(RewriteLimit::$variant {
                    observed: requirements.$field,
                    maximum: limits.$maximum,
                }));
            }
        };
    }
    limited!(output_bytes, max_output_bytes, OutputBytes);
    limited!(fields, max_fields, Fields);
    limited!(work_bytes, max_work_bytes, Work);
    limited!(components, max_components, Components);
    limited!(references, max_references, References);
    if requirements.allocations > limits.max_allocations
        || requirements.retained_bytes > limits.max_retained_bytes
        || requirements.scratch_bytes > limits.max_scratch_bytes
    {
        return Err(RewriteError::allocation(
            requirements.retained_bytes.max(requirements.scratch_bytes),
        ));
    }
    Ok(())
}

fn validate_execution_report(
    report: RewriteReport,
    requirements: RewriteExecutionRequirements,
) -> Result<(), RewriteError> {
    if report.output_bytes != requirements.output_bytes
        || report.fields != requirements.fields
        || report.work_bytes != requirements.work_bytes
        || report.components_scanned != requirements.components
        || report.references_scanned != requirements.references
        || report.allocations != requirements.allocations
        || report.retained_bytes != requirements.retained_bytes
        || report.scratch_bytes > requirements.scratch_bytes
    {
        return Err(RewriteError::invalid(InvalidReason::Verification));
    }
    Ok(())
}

fn subtract_report(
    total: RewriteReport,
    baseline: RewriteReport,
) -> Result<RewriteReport, RewriteError> {
    macro_rules! sub {
        ($field:ident) => {
            total
                .$field
                .checked_sub(baseline.$field)
                .ok_or_else(|| RewriteError::invalid(InvalidReason::Verification))?
        };
    }
    Ok(RewriteReport {
        input_bytes: 0,
        output_bytes: sub!(output_bytes),
        fields: sub!(fields),
        work_bytes: sub!(work_bytes),
        max_depth: total.max_depth,
        components_scanned: sub!(components_scanned),
        components_changed: sub!(components_changed),
        references_scanned: sub!(references_scanned),
        source_references_scanned: sub!(source_references_scanned),
        additions: 0,
        removals: 0,
        allocations: sub!(allocations),
        retained_bytes: sub!(retained_bytes),
        scratch_bytes: sub!(scratch_bytes),
    })
}

fn add_reports(left: RewriteReport, right: RewriteReport) -> Result<RewriteReport, RewriteError> {
    macro_rules! add {
        ($field:ident) => {
            left.$field
                .checked_add(right.$field)
                .ok_or_else(|| RewriteError::invalid(InvalidReason::Verification))?
        };
    }
    Ok(RewriteReport {
        input_bytes: left.input_bytes.max(right.input_bytes),
        output_bytes: add!(output_bytes),
        fields: add!(fields),
        work_bytes: add!(work_bytes),
        max_depth: left.max_depth.max(right.max_depth),
        components_scanned: add!(components_scanned),
        components_changed: add!(components_changed),
        references_scanned: add!(references_scanned),
        source_references_scanned: add!(source_references_scanned),
        additions: left.additions.max(right.additions),
        removals: left.removals.max(right.removals),
        allocations: add!(allocations),
        retained_bytes: add!(retained_bytes),
        scratch_bytes: add!(scratch_bytes),
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
    validate_removal_batch_duplicates(batch, &mut budget)?;

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
    record_output_allocation();
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
    for removal in batch.object_uuids.iter() {
        validate_selector(removal.component)?;
        if removal.object_identifier == 0 || removal.expected_uuid == UuidBits::new(0, 0) {
            return Err(RewriteError::invalid(InvalidReason::InvalidIdentifier));
        }
    }
    for removal in batch.external_references.iter() {
        validate_selector(removal.source)?;
        validate_selector(removal.target)?;
        if removal.object_identifier == 0 {
            return Err(RewriteError::invalid(InvalidReason::InvalidIdentifier));
        }
    }
    for removal in batch.data_reference_owners.iter() {
        validate_selector(removal.component)?;
        if removal.data_identifier == 0
            || removal.object_identifier == 0
            || removal.expected_count == 0
        {
            return Err(RewriteError::invalid(InvalidReason::InvalidIdentifier));
        }
    }
    Ok(())
}

/// Check duplicate removal requests under the same operation budget as the
/// wire scan.  Batch validation used to perform these quadratic comparisons
/// before a `Budget` existed, which made a hostile request able to spend
/// unmetered work even when the source itself was small.
fn validate_removal_batch_duplicates(
    batch: RemovalBatch<'_>,
    budget: &mut Budget,
) -> Result<(), RewriteError> {
    for (index, removal) in batch.object_uuids.iter().enumerate() {
        for prior in batch.object_uuids[..index].iter() {
            budget.work(1)?;
            if prior.object_identifier == removal.object_identifier
                || prior.expected_uuid == removal.expected_uuid
            {
                return Err(RewriteError::invalid(InvalidReason::DuplicateRemoval));
            }
        }
    }
    for (index, removal) in batch.external_references.iter().enumerate() {
        for prior in batch.external_references[..index].iter() {
            budget.work(1)?;
            if prior.source == removal.source
                && prior.target == removal.target
                && prior.object_identifier == removal.object_identifier
            {
                return Err(RewriteError::invalid(InvalidReason::DuplicateRemoval));
            }
        }
    }
    for (index, removal) in batch.data_reference_owners.iter().enumerate() {
        for prior in batch.data_reference_owners[..index].iter() {
            budget.work(1)?;
            if prior.component == removal.component
                && prior.data_identifier == removal.data_identifier
                && prior.object_identifier == removal.object_identifier
            {
                return Err(RewriteError::invalid(InvalidReason::DuplicateRemoval));
            }
        }
    }
    Ok(())
}

/// The combined operation has one exact ownership transition: every current
/// component touched by a removal must receive the token update, and no
/// unrelated component may be included in that update.  The check is done on
/// the borrowed selector sets without allocating a second ownership index.
fn validate_combined_selector_coverage(
    removals: RemovalBatch<'_>,
    save_tokens: SaveTokenBatch<'_>,
    budget: &mut Budget,
) -> Result<(), RewriteError> {
    let removal_count = mutation_selector_count(removals)?;
    for index in 0..removal_count {
        let selector = mutation_selector_at(removals, index);
        let mut duplicate = false;
        for prior_index in 0..index {
            let prior = mutation_selector_at(removals, prior_index);
            budget.work(
                selector
                    .locator
                    .len()
                    .checked_add(1)
                    .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?,
            )?;
            if prior == selector {
                duplicate = true;
                break;
            }
        }
        if duplicate {
            continue;
        }
        let mut found = false;
        for candidate in save_tokens.components.iter().copied() {
            budget.work(
                candidate
                    .locator
                    .len()
                    .checked_add(1)
                    .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?,
            )?;
            if candidate == selector {
                found = true;
                break;
            }
        }
        if !found {
            return Err(RewriteError::invalid(InvalidReason::ComponentMismatch));
        }
    }
    for selector in save_tokens.components.iter().copied() {
        let mut found = false;
        for index in 0..removal_count {
            let candidate = mutation_selector_at(removals, index);
            budget.work(
                selector
                    .locator
                    .len()
                    .checked_add(1)
                    .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?,
            )?;
            if candidate == selector {
                found = true;
                break;
            }
        }
        if !found {
            return Err(RewriteError::invalid(InvalidReason::ComponentMismatch));
        }
    }
    Ok(())
}

fn mutation_selector_count(batch: RemovalBatch<'_>) -> Result<usize, RewriteError> {
    batch
        .object_uuids
        .len()
        .checked_add(batch.external_references.len())
        .and_then(|count| count.checked_add(batch.data_reference_owners.len()))
        .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))
}

fn mutation_selector_at<'source>(
    batch: RemovalBatch<'source>,
    index: usize,
) -> ComponentSelector<'source> {
    if index < batch.object_uuids.len() {
        return batch.object_uuids[index].component;
    }
    let shifted = index - batch.object_uuids.len();
    if shifted < batch.external_references.len() {
        return batch.external_references[shifted].source;
    }
    batch.data_reference_owners[shifted - batch.external_references.len()].component
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

fn removal_contains_object(
    batch: RemovalBatch<'_>,
    object_identifier: u64,
    budget: &mut Budget,
) -> Result<bool, RewriteError> {
    let mut found = false;
    for removal in batch.object_uuids.iter() {
        budget.work(1)?;
        found |= removal.object_identifier == object_identifier;
    }
    for removal in batch.external_references.iter() {
        budget.work(1)?;
        found |= removal.object_identifier == object_identifier;
    }
    for removal in batch.data_reference_owners.iter() {
        budget.work(1)?;
        found |= removal.object_identifier == object_identifier;
    }
    Ok(found)
}

fn removed_uuid_contains_object(
    batch: RemovalBatch<'_>,
    object_identifier: u64,
    budget: &mut Budget,
) -> Result<bool, RewriteError> {
    let mut found = false;
    for removal in batch.object_uuids.iter() {
        budget.work(1)?;
        found |= removal.object_identifier == object_identifier;
    }
    Ok(found)
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
                let deleted_object = reference.object.map_or(Ok(false), |object| {
                    removed_uuid_contains_object(batch, object, budget)
                })?;
                let mut authorized = false;
                for (index, removal) in batch.external_references.iter().enumerate() {
                    budget.work(1)?;
                    if reference.object != Some(removal.object_identifier) {
                        continue;
                    }
                    let selected = current
                        && field.number == 6
                        && identifier == removal.source.identifier
                        && locator == removal.source.locator
                        && reference.target == removal.target.identifier;
                    if !selected {
                        // External-only ownership changes may leave another
                        // component pointing at the same retained object. If
                        // the object UUID is also removed, the global
                        // `deleted_object` check below still fails closed on
                        // every unauthorized occurrence.
                        continue;
                    }
                    if reference.is_weak != removal.expected_is_weak {
                        return Err(RewriteError::invalid(InvalidReason::RemovalMismatch));
                    }
                    if reference.unknown_fields {
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
                    if entry.unknown_fields {
                        return Err(RewriteError::invalid(InvalidReason::RemovalMismatch));
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
    let mut unknown_fields = false;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        match field.number {
            1 => set_once(&mut data_identifier, field.varint()?)?,
            2 => {},
            _ => unknown_fields = true,
        }
    }
    let data_identifier = data_identifier
        .filter(|value| *value != 0)
        .ok_or_else(|| RewriteError::invalid(InvalidReason::InvalidIdentifier))?;
    let child_depth = depth
        .checked_add(1)
        .ok_or_else(|| RewriteError::invalid(InvalidReason::MalformedWire))?;
    let mut selected_owners = 0usize;
    let mut surviving_owners = 0usize;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        if field.number != 2 {
            continue;
        }
        let (object, count, owner_unknown_fields) =
            decode_data_owner(field.bytes()?, budget, child_depth)?;
        let deleted_object = removed_uuid_contains_object(batch, object, budget)?;
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
            if owner_unknown_fields {
                return Err(RewriteError::invalid(InvalidReason::RemovalMismatch));
            }
            authorized = true;
            selected_owners = checked_add(selected_owners, 1)?;
            state.data_owners[index].current = checked_add(state.data_owners[index].current, 1)?;
        }
        if !authorized {
            surviving_owners = checked_add(surviving_owners, 1)?;
        }
        if deleted_object && !authorized {
            return Err(RewriteError::invalid(if !current {
                InvalidReason::VersionedRemoval
            } else {
                InvalidReason::CrossComponentRemoval
            }));
        }
    }
    if selected_owners != 0 && surviving_owners == 0 && unknown_fields {
        return Err(RewriteError::invalid(InvalidReason::RemovalMismatch));
    }
    Ok(())
}

fn decode_data_owner(
    source: &[u8],
    budget: &mut Budget,
    depth: u32,
) -> Result<(u64, u32, bool), RewriteError> {
    budget.reference()?;
    budget.message(source, depth)?;
    let mut object = None;
    let mut count = None;
    let mut unknown_fields = false;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        match field.number {
            1 => set_once(&mut object, field.varint()?)?,
            2 => set_once(
                &mut count,
                u32::try_from(field.varint()?)
                    .map_err(|_error| RewriteError::invalid(InvalidReason::MalformedWire))?,
            )?,
            _ => unknown_fields = true,
        }
    }
    Ok((
        object
            .filter(|value| *value != 0)
            .ok_or_else(|| RewriteError::invalid(InvalidReason::InvalidIdentifier))?,
        count
            .filter(|value| *value != 0)
            .ok_or_else(|| RewriteError::invalid(InvalidReason::InvalidIdentifier))?,
        unknown_fields,
    ))
}

fn scan_ambiguous_ids(
    field: Field<'_>,
    batch: RemovalBatch<'_>,
    budget: &mut Budget,
) -> Result<(), RewriteError> {
    match field.wire {
        0 => {
            budget.reference()?;
            let identifier = field.varint()?;
            if removal_contains_object(batch, identifier, budget)? {
                Err(RewriteError::invalid(InvalidReason::CrossComponentRemoval))
            } else {
                Ok(())
            }
        },
        2 => {
            let mut packed = field.bytes()?;
            while !packed.is_empty() {
                budget.reference()?;
                let identifier = take_varint(&mut packed)?;
                if removal_contains_object(batch, identifier, budget)? {
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
    let mut selected = false;
    for removal in batch.object_uuids.iter() {
        budget.work(1)?;
        selected |= removal.component.identifier == component
            && removal.component.locator == locator
            && removal.object_identifier == entry.object
            && removal.expected_uuid == entry.uuid;
    }
    if selected && entry.unknown_fields {
        return Err(RewriteError::invalid(InvalidReason::RemovalMismatch));
    }
    Ok(selected)
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
    let mut selected = false;
    for removal in batch.external_references.iter() {
        budget.work(1)?;
        selected |= removal.source.identifier == component
            && removal.source.locator == locator
            && removal.target.identifier == reference.target
            && Some(removal.object_identifier) == reference.object
            && removal.expected_is_weak == reference.is_weak;
    }
    if selected && reference.unknown_fields {
        return Err(RewriteError::invalid(InvalidReason::RemovalMismatch));
    }
    Ok(selected)
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
    let mut unknown_fields = false;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        match field.number {
            1 => set_once(&mut data_identifier, field.varint()?)?,
            2 => {},
            _ => unknown_fields = true,
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
        let (object, count, owner_unknown_fields) =
            decode_data_owner(field.bytes()?, budget, depth + 1)?;
        let mut selected = false;
        for removal in batch.data_reference_owners.iter() {
            budget.work(1)?;
            selected |= removal.component.identifier == component
                && removal.component.locator == locator
                && removal.data_identifier == data_identifier
                && removal.object_identifier == object
                && removal.expected_count == count;
        }
        if !selected {
            surviving_owners = checked_add(surviving_owners, 1)?;
            size = checked_add(size, field.raw.len())?;
        } else {
            if owner_unknown_fields {
                return Err(RewriteError::invalid(InvalidReason::RemovalMismatch));
            }
            selected_count = checked_add(selected_count, 1)?;
        }
    }
    if selected_count != 0 && surviving_owners == 0 && unknown_fields {
        return Err(RewriteError::invalid(InvalidReason::RemovalMismatch));
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
            let (object, count, owner_unknown_fields) =
                decode_data_owner(field.bytes()?, budget, depth + 1)?;
            let mut selected = false;
            for removal in batch.data_reference_owners.iter() {
                budget.work(1)?;
                selected |= removal.component.identifier == component
                    && removal.component.locator == locator
                    && removal.data_identifier == data_identifier
                    && removal.object_identifier == object
                    && removal.expected_count == count;
            }
            if selected {
                if owner_unknown_fields {
                    return Err(RewriteError::invalid(InvalidReason::RemovalMismatch));
                }
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
    unknown_fields: bool,
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
    let mut unknown_fields = false;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        match field.number {
            1 => set_once(&mut target, field.varint()?)?,
            2 => set_once(&mut object, field.varint()?)?,
            3 => set_once(&mut is_weak, canonical_bool(field.varint()?)?)?,
            _ => unknown_fields = true,
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
        unknown_fields,
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
    unknown_fields: bool,
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
    let mut unknown_fields = false;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        match field.number {
            1 => set_once(&mut object, field.varint()?)?,
            2 => {
                let raw = field.bytes()?;
                let (decoded, uuid_unknown_fields) = decode_uuid(raw, budget, child_depth)?;
                set_once(&mut uuid, decoded)?;
                set_once(&mut uuid_raw, raw)?;
                unknown_fields |= uuid_unknown_fields;
            },
            _ => unknown_fields = true,
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
    Ok(ObjectUuid {
        object,
        uuid,
        unknown_fields,
    })
}

fn decode_uuid(
    source: &[u8],
    budget: &mut Budget,
    depth: u32,
) -> Result<(UuidBits, bool), RewriteError> {
    budget.message(source, depth)?;
    let mut lower = None;
    let mut upper = None;
    let mut unknown_fields = false;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        match field.number {
            1 => set_once(&mut lower, field.varint()?)?,
            2 => set_once(&mut upper, field.varint()?)?,
            _ => unknown_fields = true,
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
    Ok((snapshot, unknown_fields))
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

fn repeated_counter(measured: usize, current: usize) -> Result<usize, RewriteError> {
    current
        .checked_sub(measured)
        .and_then(|delta| current.checked_add(delta))
        .ok_or_else(|| RewriteError::invalid(InvalidReason::Verification))
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
        #[cfg(test)]
        WORK_CHARGES.set(WORK_CHARGES.get().saturating_add(amount));
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
    fn pad_repeated_counters(
        &mut self,
        fields: usize,
        work: usize,
        components: usize,
        references: usize,
    ) -> Result<(), RewriteError> {
        if self.fields > fields
            || self.components_scanned > components
            || self.references_scanned > references
        {
            return Err(RewriteError::invalid(InvalidReason::Verification));
        }
        self.fields = fields;
        self.components_scanned = components;
        self.references_scanned = references;
        if self.work_bytes > work {
            return Err(RewriteError::invalid(InvalidReason::Verification));
        }
        self.work(work - self.work_bytes)?;
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
