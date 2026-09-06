//! Source-borrowed PackageMetadata facts and atomic media-lifecycle edits.
//!
//! The lifecycle owner has to change two logically different projections of
//! `Index/Metadata.iwa`: object UUID registrations live in the general
//! PackageMetadata registry, while data records and component owners live in
//! the media projection.  Keeping those projections in separate codecs is
//! useful, but publishing them separately is not.  This adapter performs the
//! source census once for each projection, validates all selected witnesses,
//! and only returns a candidate after every requested transition succeeds.
//!
//! All facts in [`MetadataSnapshot`] borrow the caller's decompressed
//! PackageMetadata payload.  The snapshot retains bounded fixed-width vectors
//! only; generated repeated protobuf views never cross this module.  The
//! codecs perform their own exact preflights before their candidate buffers
//! are allocated.  A caller that has an operation-wide budget can charge the
//! snapshot resources before retaining it and pass the same charge callback
//! to [`rewrite_metadata`]; the adapter invokes it before every transient
//! rewrite allocation.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The lifecycle adapter keeps its borrowed facts and transition helpers together."
)]

use core::mem::size_of;

use litchi_iwa_protos::{
    package_metadata_codec as identity_codec, package_metadata_media_codec as media_codec,
};

/// A strict identity-registry component selector owned by the lifecycle
/// snapshot.  The selector is private to the Keynote owner; native component
/// identifiers never appear in the public semantic API.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct ComponentIdentity<'source> {
    identifier: u64,
    locator: &'source str,
    current: bool,
}

impl<'source> ComponentIdentity<'source> {
    #[must_use]
    pub(super) const fn selector(self) -> identity_codec::ComponentSelector<'source> {
        identity_codec::ComponentSelector::new(self.identifier, self.locator)
    }
}

/// Borrowed object-to-UUID identity observed in the source registry.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct ObjectIdentity<'source> {
    component: ComponentIdentity<'source>,
    object_identifier: u64,
    uuid: identity_codec::UuidBits,
    current: bool,
}

impl<'source> ObjectIdentity<'source> {
    #[must_use]
    pub(super) const fn object_identifier(self) -> u64 {
        self.object_identifier
    }

    #[must_use]
    pub(super) const fn uuid(self) -> identity_codec::UuidBits {
        self.uuid
    }
}

/// Borrowed component-external-reference ownership observed in the source
/// PackageMetadata registry.  The target locator is resolved from the
/// snapshot's current component census because the wire record stores only
/// the target component identifier.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct ExternalReferenceIdentity<'source> {
    source: ComponentIdentity<'source>,
    target_component_identifier: u64,
    object_identifier: Option<u64>,
    is_weak: Option<bool>,
    versioned: bool,
    unknown_fields: bool,
}

/// Borrowed DataInfo facts required to decide whether a payload can be
/// reclaimed after its final component owner is removed.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct DataIdentity<'source> {
    identifier: u64,
    digest: &'source [u8],
    file_name: &'source str,
    materialized_length: Option<u64>,
    unknown_fields: bool,
}

impl<'source> DataIdentity<'source> {
    #[must_use]
    pub(super) const fn identifier(self) -> u64 {
        self.identifier
    }

    #[must_use]
    pub(super) const fn digest(self) -> &'source [u8] {
        self.digest
    }

    #[must_use]
    pub(super) const fn file_name(self) -> &'source str {
        self.file_name
    }

    #[must_use]
    pub(super) const fn materialized_length(self) -> Option<u64> {
        self.materialized_length
    }

    #[must_use]
    pub(super) const fn has_unknown_fields(self) -> bool {
        self.unknown_fields
    }
}

/// Borrowed component data-owner fact.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct DataOwnerIdentity {
    component_identifier: u64,
    data_identifier: u64,
    object_identifier: u64,
    count: u32,
    current: bool,
}

impl DataOwnerIdentity {
    #[must_use]
    pub(super) const fn count(self) -> u32 {
        self.count
    }
}

/// Borrowed ComponentDataReference facts.  A DataInfo record can have
/// several parents, and versioned parents are immutable for lifecycle edits;
/// retaining this envelope lets the owner distinguish those cases before it
/// prepares a final reclamation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct DataReferenceIdentity<'source> {
    component: ComponentIdentity<'source>,
    data_identifier: u64,
    owner_count: usize,
    current: bool,
    versioned: bool,
    unknown_fields: bool,
}

impl<'source> DataReferenceIdentity<'source> {
    #[must_use]
    pub(super) const fn is_versioned(self) -> bool {
        self.versioned
    }

    #[must_use]
    pub(super) const fn has_unknown_fields(self) -> bool {
        self.unknown_fields
    }
}

/// A borrowed source witness for the root data-metadata map.
pub(super) type DataMetadataMapWitness<'source> = media_codec::DataMetadataMapSource<'source>;

/// Identity additions and removals use the already-hardened borrowed request
/// types from the neutral PackageMetadata codec.  Type aliases keep the
/// lifecycle adapter allocation-free while still giving the owner a focused
/// vocabulary.
pub(super) type IdentityAddition<'source> = identity_codec::ObjectUuidAddition<'source>;
pub(super) type IdentityRemoval<'source> = identity_codec::ObjectUuidRemoval<'source>;
pub(super) type IdentitySaveTokens<'source> = identity_codec::SaveTokenBatch<'source>;

/// Media owner/data-record changes to apply with the UUID transition.
pub(super) type MediaBatch<'source> = media_codec::MediaRewriteBatch<'source>;

/// A source-atomic identity transition.
#[derive(Debug, Clone, Copy)]
pub(super) struct IdentityBatch<'source> {
    expected_last_identifier: u64,
    new_last_identifier: Option<u64>,
    additions: &'source [IdentityAddition<'source>],
    removals: &'source [IdentityRemoval<'source>],
    save_tokens: Option<IdentitySaveTokens<'source>>,
}

impl<'source> IdentityBatch<'source> {
    /// Construct an additions-only transition.  The new watermark must be
    /// strictly above the source watermark; the codec checks every addition
    /// against the same source witness before allocating output.
    #[must_use]
    pub(super) const fn additions(
        expected_last_identifier: u64,
        new_last_identifier: u64,
        additions: &'source [IdentityAddition<'source>],
    ) -> Self {
        Self {
            expected_last_identifier,
            new_last_identifier: Some(new_last_identifier),
            additions,
            removals: &[],
            save_tokens: None,
        }
    }

    /// Construct a removals-only transition.  Removal retains the source
    /// watermark, as native Keynote does when an object is deleted.
    #[must_use]
    pub(super) const fn removals(
        expected_last_identifier: u64,
        removals: &'source [IdentityRemoval<'source>],
    ) -> Self {
        Self {
            expected_last_identifier,
            new_last_identifier: None,
            additions: &[],
            removals,
            save_tokens: None,
        }
    }

    #[must_use]
    pub(super) const fn expected_last_identifier(self) -> u64 {
        self.expected_last_identifier
    }

    #[must_use]
    pub(super) const fn new_last_identifier(self) -> Option<u64> {
        self.new_last_identifier
    }

    #[must_use]
    pub(super) const fn uuid_additions(self) -> &'source [IdentityAddition<'source>] {
        self.additions
    }

    #[must_use]
    pub(super) const fn uuid_removals(self) -> &'source [IdentityRemoval<'source>] {
        self.removals
    }
}

/// Metadata resource evidence that a lifecycle owner can charge to its
/// operation-wide budget before retaining [`MetadataSnapshot`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct MetadataSnapshotResources {
    identity_report: identity_codec::RewriteReport,
    media_report: media_codec::DecodeReport,
    scratch_bytes: usize,
}

impl MetadataSnapshotResources {
    #[must_use]
    pub(super) const fn identity_report(self) -> identity_codec::RewriteReport {
        self.identity_report
    }

    #[must_use]
    pub(super) const fn media_report(self) -> media_codec::DecodeReport {
        self.media_report
    }

    #[must_use]
    pub(super) const fn scratch_bytes(self) -> usize {
        self.scratch_bytes
    }
}

/// A borrowed, bounded PackageMetadata snapshot used as the source witness
/// for one lifecycle transaction.
#[derive(Debug)]
pub(super) struct MetadataSnapshot<'source> {
    payload: &'source [u8],
    last_identifier: u64,
    resources: MetadataSnapshotResources,
    components: Vec<ComponentIdentity<'source>>,
    objects: Vec<ObjectIdentity<'source>>,
    external_references: Vec<ExternalReferenceIdentity<'source>>,
    data: Vec<DataIdentity<'source>>,
    references: Vec<DataReferenceIdentity<'source>>,
    owners: Vec<DataOwnerIdentity>,
    map: Option<DataMetadataMapWitness<'source>>,
}

impl<'source> MetadataSnapshot<'source> {
    /// Strictly inspect both PackageMetadata projections.  The callback is
    /// invoked before each retained vector reserves memory and is intended to
    /// charge a shared lifecycle budget.  It must fail closed on a budget
    /// violation; no candidate is emitted by this function.
    pub(super) fn inspect(
        payload: &'source [u8],
        identity_options: identity_codec::RewriteOptions,
        media_options: media_codec::DecodeOptions,
        map: Option<DataMetadataMapWitness<'source>>,
        charge: &mut dyn FnMut(usize) -> Result<(), MetadataError>,
    ) -> Result<Self, MetadataError> {
        let mut identity_noop = NoopIdentityVisitor;
        let identity_report = identity_codec::inspect_package_metadata_with_visitor(
            payload,
            identity_options,
            &mut identity_noop,
        )
        .map_err(MetadataError::Identity)?
        .report();
        let media_report = media_codec::inspect_package_metadata_media(payload, media_options)
            .map_err(MetadataError::MediaDecode)?;

        let component_capacity = identity_report.components_scanned();
        let object_capacity = identity_report.references_scanned();
        // The neutral identity report intentionally exposes the aggregate
        // reference count rather than a second per-kind counter.  Reusing it
        // as a strict upper bound keeps this collector allocation-free while
        // retaining every external-reference witness needed by lifecycle
        // ownership checks.
        let external_reference_capacity = identity_report.references_scanned();
        let data_capacity = media_report.data_records();
        let reference_capacity = media_report.data_references();
        let owner_capacity = media_report.owners();
        let scratch_bytes = checked_sum([
            component_capacity
                .checked_mul(size_of::<ComponentIdentity<'source>>())
                .ok_or(MetadataError::Invalid)?,
            object_capacity
                .checked_mul(size_of::<ObjectIdentity<'source>>())
                .ok_or(MetadataError::Invalid)?,
            external_reference_capacity
                .checked_mul(size_of::<ExternalReferenceIdentity<'source>>())
                .ok_or(MetadataError::Invalid)?,
            data_capacity
                .checked_mul(size_of::<DataIdentity<'source>>())
                .ok_or(MetadataError::Invalid)?,
            reference_capacity
                .checked_mul(size_of::<DataReferenceIdentity<'source>>())
                .ok_or(MetadataError::Invalid)?,
            owner_capacity
                .checked_mul(size_of::<DataOwnerIdentity>())
                .ok_or(MetadataError::Invalid)?,
        ])?;

        if let Some(map_source) = map {
            // A map witness is meaningful only when the source metadata has
            // the exact root edge that points at it.  The payload itself is
            // validated by `DataMetadataMapSource::from_source`; pairing the
            // identifier here prevents a caller from smuggling a valid map
            // from a different package into this snapshot.
            if media_report.data_metadata_map_identifier() != Some(map_source.object_identifier()) {
                return Err(MetadataError::Ambiguous);
            }
        }

        charge_nonempty(
            charge,
            component_capacity,
            size_of::<ComponentIdentity<'source>>(),
        )?;
        charge_nonempty(
            charge,
            object_capacity,
            size_of::<ObjectIdentity<'source>>(),
        )?;
        charge_nonempty(
            charge,
            external_reference_capacity,
            size_of::<ExternalReferenceIdentity<'source>>(),
        )?;
        charge_nonempty(charge, data_capacity, size_of::<DataIdentity<'source>>())?;
        charge_nonempty(
            charge,
            reference_capacity,
            size_of::<DataReferenceIdentity<'source>>(),
        )?;
        charge_nonempty(charge, owner_capacity, size_of::<DataOwnerIdentity>())?;
        // The media component visitor is intentionally not retained: identity
        // components are authoritative for UUID routing.  Its report remains
        // in the resource evidence, so no second vector is allocated here.

        let mut identity = IdentityCollector {
            source: payload,
            components: Vec::new(),
            objects: Vec::new(),
            external_references: Vec::new(),
        };
        try_reserve(
            &mut identity.components,
            component_capacity,
            size_of::<ComponentIdentity<'source>>(),
        )?;
        try_reserve(
            &mut identity.objects,
            object_capacity,
            size_of::<ObjectIdentity<'source>>(),
        )?;
        try_reserve(
            &mut identity.external_references,
            external_reference_capacity,
            size_of::<ExternalReferenceIdentity<'source>>(),
        )?;
        let inspected = identity_codec::inspect_package_metadata_with_visitor(
            payload,
            identity_options,
            &mut identity,
        )
        .map_err(MetadataError::Identity)?;
        if inspected.report() != identity_report {
            return Err(MetadataError::Invalid);
        }

        let mut media = MediaCollector {
            source: payload,
            data: Vec::new(),
            references: Vec::new(),
            owners: Vec::new(),
        };
        try_reserve(
            &mut media.data,
            data_capacity,
            size_of::<DataIdentity<'source>>(),
        )?;
        try_reserve(
            &mut media.references,
            reference_capacity,
            size_of::<DataReferenceIdentity<'source>>(),
        )?;
        try_reserve(
            &mut media.owners,
            owner_capacity,
            size_of::<DataOwnerIdentity>(),
        )?;
        let visited = media_codec::visit_package_metadata_media(payload, media_options, &mut media)
            .map_err(MetadataError::MediaDecode)?;
        if visited != media_report {
            return Err(MetadataError::Invalid);
        }

        Ok(Self {
            payload,
            last_identifier: inspected.last_object_identifier(),
            resources: MetadataSnapshotResources {
                identity_report,
                media_report,
                scratch_bytes,
            },
            components: identity.components,
            objects: identity.objects,
            external_references: identity.external_references,
            data: media.data,
            references: media.references,
            owners: media.owners,
            map,
        })
    }

    #[must_use]
    pub(super) const fn last_identifier(&self) -> u64 {
        self.last_identifier
    }

    #[must_use]
    pub(super) const fn resources(&self) -> MetadataSnapshotResources {
        self.resources
    }

    #[must_use]
    pub(super) const fn map(&self) -> Option<DataMetadataMapWitness<'source>> {
        self.map
    }

    /// Resolve one unique current component by effective native locator.
    pub(super) fn current_component(
        &self,
        locator: &str,
    ) -> Result<ComponentIdentity<'source>, MetadataError> {
        let mut selected = None;
        for component in self.components.iter().copied() {
            if component.current && component.locator == locator {
                if selected.replace(component).is_some() {
                    return Err(MetadataError::Ambiguous);
                }
            }
        }
        selected.ok_or(MetadataError::Missing)
    }

    /// Return the bounded census cost of one external-dependency witness.
    /// The lookup resolves both current selectors and then scans every
    /// retained external record.  Callers should charge this once for each
    /// distinct dependency before invoking [`Self::require_current_external_dependency`].
    #[must_use]
    pub(super) fn external_dependency_lookup_work(&self) -> usize {
        self.components
            .len()
            .saturating_mul(3)
            .saturating_add(self.external_references.len())
            .max(1)
    }

    /// Require one exact current field-6 external edge from `source` to the
    /// current component selected by `target_locator` and `author_id`.
    ///
    /// Component-external references are component-level owners shared by
    /// every archive object in the source component.  A clone therefore must
    /// reuse the existing edge when its author is external; it must not infer
    /// ownership from a raw object identifier alone.  Versioned records,
    /// duplicate matching records, and selected records carrying unknown
    /// fields are rejected as ambiguous so a lifecycle caller cannot publish
    /// against an uncertain source witness.  Explicitly weak author edges are
    /// also refused; comment authors require a strong dependency (the native
    /// representation is either an omitted weakness flag or `false`).
    pub(super) fn require_current_external_dependency(
        &self,
        source: ComponentIdentity<'source>,
        target_locator: &str,
        author_id: u64,
    ) -> Result<(), MetadataError> {
        if author_id == 0 {
            return Err(MetadataError::Invalid);
        }
        if !source.current {
            return Err(MetadataError::Ambiguous);
        }
        let selected_source = self.current_component(source.locator)?;
        if selected_source.identifier != source.identifier {
            return Err(MetadataError::Ambiguous);
        }
        let target = self.current_component(target_locator)?;
        let mut source_identifier_matches = 0usize;
        let mut target_identifier_matches = 0usize;
        for component in self.components.iter().copied() {
            if !component.current {
                continue;
            }
            if component.identifier == selected_source.identifier {
                source_identifier_matches = source_identifier_matches
                    .checked_add(1)
                    .ok_or(MetadataError::Invalid)?;
            }
            if component.identifier == target.identifier {
                target_identifier_matches = target_identifier_matches
                    .checked_add(1)
                    .ok_or(MetadataError::Invalid)?;
            }
        }
        if source_identifier_matches != 1 || target_identifier_matches != 1 {
            return Err(MetadataError::Ambiguous);
        }

        let mut current_matches = 0usize;
        let mut versioned_matches = 0usize;
        let mut unknown_match = false;
        let mut weak_match = false;
        for reference in self.external_references.iter().copied() {
            if reference.source.identifier != selected_source.identifier
                || reference.source.locator != selected_source.locator
                || reference.target_component_identifier != target.identifier
                || reference.object_identifier != Some(author_id)
            {
                continue;
            }
            if reference.versioned || !reference.source.current {
                versioned_matches = versioned_matches
                    .checked_add(1)
                    .ok_or(MetadataError::Invalid)?;
            } else {
                current_matches = current_matches
                    .checked_add(1)
                    .ok_or(MetadataError::Invalid)?;
            }
            unknown_match |= reference.unknown_fields;
            weak_match |= reference.is_weak == Some(true);
        }

        if versioned_matches != 0 || unknown_match || current_matches > 1 || weak_match {
            return Err(MetadataError::Ambiguous);
        }
        if current_matches == 0 {
            return Err(MetadataError::Missing);
        }
        Ok(())
    }

    /// Resolve one unique current object UUID in the selected component.
    /// Any versioned or cross-component registration for the same object is
    /// treated as ambiguous rather than silently ignored.
    pub(super) fn object_uuid(
        &self,
        component: ComponentIdentity<'source>,
        object_identifier: u64,
    ) -> Result<identity_codec::UuidBits, MetadataError> {
        let mut selected = None;
        for object in self.objects.iter().copied() {
            if object.object_identifier != object_identifier {
                continue;
            }
            if !object.current
                || object.component.identifier != component.identifier
                || object.component.locator != component.locator
            {
                return Err(MetadataError::Ambiguous);
            }
            if selected.replace(object.uuid).is_some() {
                return Err(MetadataError::Ambiguous);
            }
        }
        selected.ok_or(MetadataError::Missing)
    }

    /// Return whether any source registry entry already owns this UUID.
    /// Clone UUID allocation uses this witness before publishing a new
    /// ObjectUuidAddition, including for native BuildChunk records that do
    /// not themselves have an object-map entry.
    #[must_use]
    pub(super) fn has_uuid(&self, uuid: identity_codec::UuidBits) -> bool {
        self.objects.iter().any(|object| object.uuid() == uuid)
    }

    /// Resolve one unique DataInfo record, retaining the exact source name
    /// needed by the physical owner when the final owner disappears.
    pub(super) fn data_info(
        &self,
        identifier: u64,
    ) -> Result<DataIdentity<'source>, MetadataError> {
        let mut selected = None;
        for record in self.data.iter().copied() {
            if record.identifier == identifier {
                if selected.replace(record).is_some() {
                    return Err(MetadataError::Ambiguous);
                }
            }
        }
        selected.ok_or(MetadataError::Missing)
    }

    /// Borrow every source DataInfo record so the physical owner can reject
    /// aliased materialized names before deleting a ZIP member.
    #[must_use]
    pub(super) fn data_records(&self) -> &[DataIdentity<'source>] {
        &self.data
    }

    /// Borrow all parsed current and versioned ObjectReference owners.
    #[must_use]
    pub(super) fn owners(&self) -> &[DataOwnerIdentity] {
        &self.owners
    }

    /// Return the exact number of current ObjectReference records for a
    /// DataInfo identifier.  A versioned parent makes the selected data
    /// ambiguous for a mutable lifecycle operation and is rejected rather
    /// than silently treated as reclaimable.
    pub(super) fn current_owner_count(&self, data_identifier: u64) -> Result<usize, MetadataError> {
        // Keep a missing identifier distinct from a valid DataInfo with no
        // current parents; callers use that distinction to refuse stale
        // final-reclamation requests.
        self.data_info(data_identifier)?;
        let mut count = 0usize;
        for reference in self.references.iter().copied() {
            if reference.data_identifier != data_identifier {
                continue;
            }
            if reference.is_versioned() || reference.has_unknown_fields() {
                return Err(MetadataError::Ambiguous);
            }
            count = count
                .checked_add(reference.owner_count)
                .ok_or(MetadataError::Invalid)?;
        }
        Ok(count)
    }

    /// Count exact current owners for one component/data/object tuple.
    pub(super) fn owner(
        &self,
        component: ComponentIdentity<'source>,
        data_identifier: u64,
        object_identifier: u64,
    ) -> Result<DataOwnerIdentity, MetadataError> {
        let mut selected = None;
        for owner in self.owners.iter().copied() {
            if owner.current
                && owner.component_identifier == component.identifier
                && owner.data_identifier == data_identifier
                && owner.object_identifier == object_identifier
            {
                if selected.replace(owner).is_some() {
                    return Err(MetadataError::Ambiguous);
                }
            }
        }
        selected.ok_or(MetadataError::Missing)
    }
}

/// Aggregate resource evidence for all PackageMetadata identity phases.
///
/// A transition without save tokens can require two strict candidates (one
/// removal followed by one addition).  Keeping a single last report would
/// undercharge the caller's shared budget, so every additive resource counter
/// is checked and accumulated here.  The maximum depth remains a maximum.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct IdentityRewriteReport {
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

impl IdentityRewriteReport {
    #[must_use]
    pub(super) const fn output_bytes(self) -> usize {
        self.output_bytes
    }

    #[must_use]
    pub(super) const fn fields(self) -> usize {
        self.fields
    }

    #[must_use]
    pub(super) const fn work_bytes(self) -> usize {
        self.work_bytes
    }

    #[must_use]
    pub(super) const fn max_depth(self) -> u32 {
        self.max_depth
    }

    #[must_use]
    pub(super) const fn references_scanned(self) -> usize {
        self.references_scanned
    }

    #[must_use]
    pub(super) const fn scratch_bytes(self) -> usize {
        self.scratch_bytes
    }
}

/// Output of one metadata transition.  `removed_data_paths` borrows names
/// from the source snapshot and is consumed by the physical ZIP owner after
/// the metadata candidate has passed its own verification.
#[derive(Debug)]
pub(super) struct MetadataRewrite<'source> {
    bytes: Vec<u8>,
    removed_data_paths: Vec<&'source str>,
    identity_report: Option<IdentityRewriteReport>,
    media_report: Option<media_codec::RewriteReport>,
}

impl<'source> MetadataRewrite<'source> {
    #[must_use]
    pub(super) fn into_parts(
        self,
    ) -> (
        Vec<u8>,
        Vec<&'source str>,
        Option<IdentityRewriteReport>,
        Option<media_codec::RewriteReport>,
    ) {
        (
            self.bytes,
            self.removed_data_paths,
            self.identity_report,
            self.media_report,
        )
    }
}

/// Apply UUID and media-owner transitions against one borrowed source.
///
/// The identity candidate is never published independently.  When both
/// projections change, media rewriting consumes the identity candidate and a
/// single [`MetadataRewrite`] is returned only after both strict codecs have
/// verified their respective postconditions.  A failed second phase therefore
/// leaves the caller's source bytes untouched.
pub(super) fn rewrite_metadata<'source>(
    snapshot: &MetadataSnapshot<'source>,
    identity: IdentityBatch<'source>,
    media: MediaBatch<'source>,
    identity_options: identity_codec::RewriteOptions,
    media_options: media_codec::DecodeOptions,
    charge: &mut dyn FnMut(usize) -> Result<(), MetadataError>,
) -> Result<MetadataRewrite<'source>, MetadataError> {
    validate_identity_batch(snapshot, identity)?;

    // Resolve every selected DataInfo and validate its map dependency before
    // retaining a path or charging an output allocation.  This keeps a
    // malformed or stale batch entirely side-effect free for the caller's
    // operation budget.
    if !media.data_removals().is_empty() {
        for removal in media.data_removals().iter().copied() {
            let data = snapshot.data_info(removal.identifier())?;
            if data.has_unknown_fields() || data.file_name().is_empty() {
                return Err(MetadataError::Ambiguous);
            }
        }
        match (
            snapshot
                .resources()
                .media_report()
                .data_metadata_map_identifier(),
            snapshot.map(),
            media.data_metadata_map_source(),
        ) {
            (None, None, None) => {},
            (Some(identifier), Some(expected), Some(provided))
                if identifier == expected.object_identifier()
                    && identifier == provided.object_identifier()
                    && expected.payload() == provided.payload() => {},
            // A final DataInfo removal must carry the exact map witness that
            // was paired with this source snapshot.  A caller cannot replace
            // it with a same-shaped payload from another archive.
            _ => return Err(MetadataError::Ambiguous),
        }
    }

    // Charge every candidate buffer before the first identity clone.  The
    // underlying codecs still perform their own exact output preflights; the
    // owner callback covers this adapter's transient copies and keeps a
    // shared operation budget from being reset between the two projections.
    let removed_path_bytes = media
        .data_removals()
        .len()
        .checked_mul(size_of::<&'source str>())
        .ok_or(MetadataError::Invalid)?;
    if removed_path_bytes != 0 {
        charge(removed_path_bytes)?;
    }
    // Removing an object UUID while its ComponentDataReference owner is still
    // present is rejected by the identity codec as a cross-component
    // removal.  For lifecycle deletions, therefore, the media projection is
    // rewritten first and the UUID registry is applied to that exact
    // candidate.  The media candidate is bounded by its configured output
    // ceiling, so use that ceiling when precharging the identity source clone
    // as well; every transient buffer is covered before either codec runs.
    let media_first = !identity.uuid_removals().is_empty() && media_has_removals(media);
    let identity_source_len = if media_first {
        snapshot.payload.len().max(media_options.max_output_bytes())
    } else {
        snapshot.payload.len()
    };
    charge_identity_allocations(charge, identity_source_len, identity, identity_options)?;
    if !media.is_empty() && media_options.max_output_bytes() != 0 {
        charge(media_options.max_output_bytes())?;
    }

    let mut removed_data_paths = Vec::new();
    if !media.data_removals().is_empty() {
        let path_allocation = media
            .data_removals()
            .len()
            .checked_mul(size_of::<&'source str>())
            .ok_or(MetadataError::Invalid)?;
        removed_data_paths
            .try_reserve_exact(media.data_removals().len())
            .map_err(|_| MetadataError::Allocation {
                amount: path_allocation,
            })?;
        for removal in media.data_removals().iter().copied() {
            removed_data_paths.push(snapshot.data_info(removal.identifier())?.file_name());
        }
    }

    let (bytes, identity_report, media_report) = if media_first {
        let media_output =
            media_codec::rewrite_package_metadata_media(snapshot.payload, media, media_options)
                .map_err(MetadataError::MediaRewrite)?;
        let media_report = media_output.report();
        let media_bytes = media_output.into_bytes();
        let (bytes, identity_report) = rewrite_identity(&media_bytes, identity, identity_options)?;
        (bytes, identity_report, Some(media_report))
    } else {
        let (identity_bytes, identity_report) =
            rewrite_identity(snapshot.payload, identity, identity_options)?;
        if media.is_empty() {
            (identity_bytes, identity_report, None)
        } else {
            let output = media_codec::rewrite_package_metadata_media(
                &identity_bytes,
                media,
                media_options.with_max_message_bytes(media_options.max_output_bytes()),
            )
            .map_err(MetadataError::MediaRewrite)?;
            let output_report = output.report();
            (output.into_bytes(), identity_report, Some(output_report))
        }
    };
    Ok(MetadataRewrite {
        bytes,
        removed_data_paths,
        identity_report,
        media_report,
    })
}

fn media_has_removals(batch: MediaBatch<'_>) -> bool {
    !batch.data_removals().is_empty() || !batch.owner_removals().is_empty()
}

fn rewrite_identity(
    payload: &[u8],
    batch: IdentityBatch<'_>,
    options: identity_codec::RewriteOptions,
) -> Result<(Vec<u8>, Option<IdentityRewriteReport>), MetadataError> {
    if batch.uuid_additions().is_empty() && batch.uuid_removals().is_empty() {
        if let Some(new_last) = batch.new_last_identifier() {
            if new_last < batch.expected_last_identifier() {
                return Err(MetadataError::Invalid);
            }
            if new_last == batch.expected_last_identifier() {
                if batch
                    .save_tokens
                    .is_some_and(|tokens| !tokens.components().is_empty())
                {
                    return Err(MetadataError::Invalid);
                }
                return Ok((clone_source(payload)?, None));
            }
            let additions =
                identity_codec::Batch::new(batch.expected_last_identifier(), new_last, &[], &[]);
            let output = match batch.save_tokens {
                Some(tokens) => identity_codec::rewrite_package_metadata_additions_and_save_tokens(
                    payload,
                    identity_codec::AdditionSaveTokenBatch::new(additions, tokens),
                    options,
                ),
                None => identity_codec::rewrite_package_metadata(payload, additions, options),
            }
            .map_err(MetadataError::Identity)?;
            let output_report = output.report();
            return Ok((output.into_bytes(), Some(identity_report(output_report))));
        }
        if batch
            .save_tokens
            .is_some_and(|tokens| !tokens.components().is_empty())
        {
            return Err(MetadataError::Invalid);
        }
        return Ok((clone_source(payload)?, None));
    }

    let mut current = clone_source(payload)?;
    let mut report = None;
    if !batch.uuid_removals().is_empty()
        && !batch.uuid_additions().is_empty()
        && batch.save_tokens.is_some()
    {
        let new_last = batch.new_last_identifier().ok_or(MetadataError::Invalid)?;
        let transition = identity_codec::CombinedBatch::new(
            batch.expected_last_identifier(),
            new_last,
            batch.uuid_additions(),
            &[],
            batch.uuid_removals(),
            &[],
            &[],
        );
        let output = identity_codec::rewrite_package_metadata_combined_additions_and_removals_and_save_tokens(
            &current,
            identity_codec::CombinedSaveTokenBatch::new(
                transition,
                batch.save_tokens.ok_or(MetadataError::Invalid)?,
            ),
            options,
        )
        .map_err(MetadataError::Identity)?;
        let output_report = output.report();
        current = output.into_bytes();
        return Ok((current, Some(identity_report(output_report))));
    }
    if !batch.uuid_removals().is_empty() {
        let removals = identity_codec::RemovalBatch::new(
            batch.expected_last_identifier(),
            batch.uuid_removals(),
            &[],
            &[],
        );
        let output = match batch.save_tokens {
            Some(tokens) if batch.uuid_additions().is_empty() => {
                identity_codec::rewrite_package_metadata_removals_and_save_tokens(
                    &current,
                    identity_codec::RemovalSaveTokenBatch::new(removals, tokens),
                    options,
                )
            },
            None => identity_codec::remove_package_metadata(&current, removals, options),
            Some(_) => return Err(MetadataError::Invalid),
        }
        .map_err(MetadataError::Identity)?;
        let output_report = output.report();
        current = output.into_bytes();
        report = Some(identity_report(output_report));
    }
    if !batch.uuid_additions().is_empty() {
        let new_last = batch.new_last_identifier().ok_or(MetadataError::Invalid)?;
        let additions = identity_codec::Batch::new(
            batch.expected_last_identifier(),
            new_last,
            batch.uuid_additions(),
            &[],
        );
        let output = match batch.save_tokens {
            Some(tokens) if batch.uuid_removals().is_empty() => {
                identity_codec::rewrite_package_metadata_additions_and_save_tokens(
                    &current,
                    identity_codec::AdditionSaveTokenBatch::new(additions, tokens),
                    options,
                )
            },
            None => identity_codec::rewrite_package_metadata(&current, additions, options),
            Some(_) => return Err(MetadataError::Invalid),
        }
        .map_err(MetadataError::Identity)?;
        let output_report = output.report();
        current = output.into_bytes();
        report = Some(match report.take() {
            Some(previous) => combine_identity_reports(previous, output_report)?,
            None => identity_report(output_report),
        });
    }
    Ok((current, report))
}

fn charge_identity_allocations(
    charge: &mut dyn FnMut(usize) -> Result<(), MetadataError>,
    source_len: usize,
    batch: IdentityBatch<'_>,
    options: identity_codec::RewriteOptions,
) -> Result<(), MetadataError> {
    // `rewrite_identity` starts with one fallible source clone even when the
    // registry transition is empty.  Removal candidates cannot grow beyond
    // the source stream; addition candidates are bounded by the codec's
    // configured output ceiling.  A combined save-token transition emits one
    // candidate, while a no-token transition stages removal and addition
    // candidates independently.
    let has_removals = !batch.uuid_removals().is_empty();
    let has_additions = !batch.uuid_additions().is_empty();
    let combined_save_tokens = has_removals && has_additions && batch.save_tokens.is_some();
    let watermark_only = !has_removals
        && !has_additions
        && batch
            .new_last_identifier()
            .is_some_and(|last| last > batch.expected_last_identifier());
    if !watermark_only && source_len != 0 {
        charge(source_len)?;
    }
    if combined_save_tokens || watermark_only {
        if options.max_output_bytes() != 0 {
            charge(options.max_output_bytes())?;
        }
    } else {
        if has_removals {
            if source_len != 0 {
                charge(source_len)?;
            }
        }
        if has_additions {
            if options.max_output_bytes() != 0 {
                charge(options.max_output_bytes())?;
            }
        }
    }
    Ok(())
}

fn identity_report(report: identity_codec::RewriteReport) -> IdentityRewriteReport {
    IdentityRewriteReport {
        input_bytes: report.input_bytes(),
        output_bytes: report.output_bytes(),
        fields: report.fields(),
        work_bytes: report.work_bytes(),
        max_depth: report.max_depth(),
        components_scanned: report.components_scanned(),
        components_changed: report.components_changed(),
        references_scanned: report.references_scanned(),
        source_references_scanned: report.source_references_scanned(),
        additions: report.additions(),
        removals: report.removals(),
        allocations: report.allocations(),
        retained_bytes: report.retained_bytes(),
        scratch_bytes: report.scratch_bytes(),
    }
}

fn combine_identity_reports(
    left: IdentityRewriteReport,
    right: identity_codec::RewriteReport,
) -> Result<IdentityRewriteReport, MetadataError> {
    let right = identity_report(right);
    Ok(IdentityRewriteReport {
        input_bytes: checked_add(left.input_bytes, right.input_bytes)?,
        output_bytes: checked_add(left.output_bytes, right.output_bytes)?,
        fields: checked_add(left.fields, right.fields)?,
        work_bytes: checked_add(left.work_bytes, right.work_bytes)?,
        max_depth: left.max_depth.max(right.max_depth),
        components_scanned: checked_add(left.components_scanned, right.components_scanned)?,
        components_changed: checked_add(left.components_changed, right.components_changed)?,
        references_scanned: checked_add(left.references_scanned, right.references_scanned)?,
        source_references_scanned: checked_add(
            left.source_references_scanned,
            right.source_references_scanned,
        )?,
        additions: checked_add(left.additions, right.additions)?,
        removals: checked_add(left.removals, right.removals)?,
        allocations: checked_add(left.allocations, right.allocations)?,
        retained_bytes: checked_add(left.retained_bytes, right.retained_bytes)?,
        scratch_bytes: checked_add(left.scratch_bytes, right.scratch_bytes)?,
    })
}

fn checked_add(left: usize, right: usize) -> Result<usize, MetadataError> {
    left.checked_add(right).ok_or(MetadataError::Invalid)
}

fn validate_identity_batch(
    snapshot: &MetadataSnapshot<'_>,
    batch: IdentityBatch<'_>,
) -> Result<(), MetadataError> {
    if snapshot.last_identifier() != batch.expected_last_identifier() {
        return Err(MetadataError::Invalid);
    }
    if batch.uuid_additions().is_empty() {
        if batch.uuid_removals().is_empty() {
            if batch
                .new_last_identifier()
                .is_some_and(|last| last < batch.expected_last_identifier())
            {
                return Err(MetadataError::Invalid);
            }
        } else if batch
            .new_last_identifier()
            .is_some_and(|last| last != batch.expected_last_identifier())
        {
            return Err(MetadataError::Invalid);
        }
    } else if batch
        .new_last_identifier()
        .is_none_or(|last| last <= batch.expected_last_identifier())
    {
        return Err(MetadataError::Invalid);
    }
    for addition in batch.uuid_additions().iter().copied() {
        let component = ComponentIdentity {
            identifier: addition.component().identifier(),
            locator: addition.component().locator(),
            current: true,
        };
        let current = snapshot.current_component(component.locator)?;
        if current.identifier != component.identifier {
            return Err(MetadataError::Ambiguous);
        }
        if addition.object_identifier() == 0
            || addition.uuid() == identity_codec::UuidBits::new(0, 0)
        {
            return Err(MetadataError::Invalid);
        }
        // An addition must be a genuinely new object/UUID.  The underlying
        // codec repeats this check on the authoritative source, but doing the
        // witness check here rejects an ambiguous selected graph before any
        // transition candidate is allocated.
        if snapshot
            .objects
            .iter()
            .any(|object| object.object_identifier() == addition.object_identifier())
            || snapshot
                .objects
                .iter()
                .any(|object| object.uuid() == addition.uuid())
        {
            return Err(MetadataError::Ambiguous);
        }
    }
    for removal in batch.uuid_removals().iter().copied() {
        let component = ComponentIdentity {
            identifier: removal.component().identifier(),
            locator: removal.component().locator(),
            current: true,
        };
        let current = snapshot.current_component(component.locator)?;
        if current.identifier != component.identifier {
            return Err(MetadataError::Ambiguous);
        }
        if snapshot.object_uuid(component, removal.object_identifier())? != removal.expected_uuid()
        {
            return Err(MetadataError::Ambiguous);
        }
    }
    Ok(())
}

#[derive(Debug)]
struct IdentityCollector<'source> {
    source: &'source [u8],
    components: Vec<ComponentIdentity<'source>>,
    objects: Vec<ObjectIdentity<'source>>,
    external_references: Vec<ExternalReferenceIdentity<'source>>,
}

impl identity_codec::PackageMetadataVisitor for IdentityCollector<'_> {
    fn visit_component(
        &mut self,
        component: identity_codec::ComponentDescriptor<'_>,
    ) -> Result<(), identity_codec::RewriteError> {
        if self.components.len() == self.components.capacity() {
            return Err(identity_codec::RewriteError::allocation(size_of::<
                ComponentIdentity<'_>,
            >()));
        }
        let locator = rebind_source_str(self.source, component.effective_locator())
            .ok_or_else(|| identity_codec::RewriteError::allocation(0))?;
        self.components.push(ComponentIdentity {
            identifier: component.identifier(),
            locator,
            current: component.is_current(),
        });
        Ok(())
    }

    fn visit_object_uuid(
        &mut self,
        binding: identity_codec::ObjectUuidDescriptor<'_>,
    ) -> Result<(), identity_codec::RewriteError> {
        if self.objects.len() == self.objects.capacity() {
            return Err(identity_codec::RewriteError::allocation(size_of::<
                ObjectIdentity<'_>,
            >()));
        }
        let locator = rebind_source_str(self.source, binding.component().effective_locator())
            .ok_or_else(|| identity_codec::RewriteError::allocation(0))?;
        self.objects.push(ObjectIdentity {
            component: ComponentIdentity {
                identifier: binding.component().identifier(),
                locator,
                current: binding.component().is_current(),
            },
            object_identifier: binding.object_identifier(),
            uuid: binding.uuid(),
            current: binding.component().is_current(),
        });
        Ok(())
    }

    fn visit_external_reference(
        &mut self,
        reference: identity_codec::ExternalReferenceDescriptor<'_>,
    ) -> Result<(), identity_codec::RewriteError> {
        if self.external_references.len() == self.external_references.capacity() {
            return Err(identity_codec::RewriteError::allocation(size_of::<
                ExternalReferenceIdentity<'_>,
            >()));
        }
        let locator = rebind_source_str(self.source, reference.source().effective_locator())
            .ok_or_else(|| identity_codec::RewriteError::allocation(0))?;
        self.external_references.push(ExternalReferenceIdentity {
            source: ComponentIdentity {
                identifier: reference.source().identifier(),
                locator,
                current: reference.source().is_current(),
            },
            target_component_identifier: reference.target_component_identifier(),
            object_identifier: reference.object_identifier(),
            is_weak: reference.is_weak(),
            versioned: reference.is_versioned(),
            unknown_fields: reference.has_unknown_fields(),
        });
        Ok(())
    }
}

struct NoopIdentityVisitor;

impl identity_codec::PackageMetadataVisitor for NoopIdentityVisitor {}

#[derive(Debug)]
struct MediaCollector<'source> {
    source: &'source [u8],
    data: Vec<DataIdentity<'source>>,
    references: Vec<DataReferenceIdentity<'source>>,
    owners: Vec<DataOwnerIdentity>,
}

impl media_codec::PackageMetadataMediaVisitor for MediaCollector<'_> {
    fn visit_data_info(
        &mut self,
        data: media_codec::DataInfoSnapshot<'_>,
    ) -> Result<(), media_codec::DecodeError> {
        if self.data.len() == self.data.capacity() {
            return Err(media_codec::DecodeError::invalid_for_adapter());
        }
        let digest = rebind_source_bytes(self.source, data.digest())
            .ok_or_else(media_codec::DecodeError::invalid_for_adapter)?;
        let file_name = data.file_name().unwrap_or(data.preferred_file_name());
        let file_name = rebind_source_str(self.source, file_name)
            .ok_or_else(media_codec::DecodeError::invalid_for_adapter)?;
        self.data.push(DataIdentity {
            identifier: data.identifier(),
            digest,
            file_name,
            materialized_length: data.materialized_length(),
            unknown_fields: data.has_unknown_fields(),
        });
        Ok(())
    }

    fn visit_data_reference(
        &mut self,
        component: media_codec::ComponentSnapshot<'_>,
        data_reference: media_codec::ComponentDataReferenceSnapshot<'_>,
    ) -> Result<(), media_codec::DecodeError> {
        if self.references.len() == self.references.capacity() {
            return Err(media_codec::DecodeError::invalid_for_adapter());
        }
        let locator = rebind_source_str(self.source, component.effective_locator())
            .ok_or_else(media_codec::DecodeError::invalid_for_adapter)?;
        self.references.push(DataReferenceIdentity {
            component: ComponentIdentity {
                identifier: component.identifier(),
                locator,
                current: !component.is_versioned(),
            },
            data_identifier: data_reference.data_identifier(),
            owner_count: data_reference.owner_count(),
            current: !component.is_versioned(),
            versioned: component.is_versioned(),
            unknown_fields: component.has_unknown_fields() || data_reference.has_unknown_fields(),
        });
        Ok(())
    }

    fn visit_owner(
        &mut self,
        component: media_codec::ComponentSnapshot<'_>,
        data_reference: media_codec::ComponentDataReferenceSnapshot<'_>,
        owner: media_codec::OwnerSnapshot<'_>,
    ) -> Result<(), media_codec::DecodeError> {
        if self.owners.len() == self.owners.capacity() {
            return Err(media_codec::DecodeError::invalid_for_adapter());
        }
        self.owners.push(DataOwnerIdentity {
            component_identifier: component.identifier(),
            data_identifier: data_reference.data_identifier(),
            object_identifier: owner.object_identifier(),
            count: owner.count(),
            current: !component.is_versioned(),
        });
        Ok(())
    }
}

/// Rebind a callback-provided borrowed slice to the exact original payload.
/// Buffa visitors may expose a shorter callback lifetime even though their
/// bytes are source-backed; pointer/range validation lets this adapter retain
/// the source lifetime without unsafe casts or copying selected fields.
fn rebind_source_bytes<'source>(source: &'source [u8], value: &[u8]) -> Option<&'source [u8]> {
    if value.is_empty() {
        return source.get(..0);
    }
    let source_start = source.as_ptr() as usize;
    let value_start = value.as_ptr() as usize;
    let offset = value_start.checked_sub(source_start)?;
    let end = offset.checked_add(value.len())?;
    let rebound = source.get(offset..end)?;
    (rebound == value).then_some(rebound)
}

fn rebind_source_str<'source>(source: &'source [u8], value: &str) -> Option<&'source str> {
    let rebound = rebind_source_bytes(source, value.as_bytes())?;
    core::str::from_utf8(rebound).ok()
}

/// Content-free lifecycle metadata failure.  Callers map the codec variants
/// into their format-specific public error while preserving resource and
/// allocation distinctions.
#[derive(Debug)]
pub(super) enum MetadataError {
    Identity(identity_codec::RewriteError),
    MediaDecode(media_codec::DecodeError),
    MediaRewrite(media_codec::RewriteError),
    Allocation { amount: usize },
    Missing,
    Ambiguous,
    Invalid,
}

fn charge_nonempty(
    charge: &mut dyn FnMut(usize) -> Result<(), MetadataError>,
    amount: usize,
    element_size: usize,
) -> Result<(), MetadataError> {
    if amount == 0 {
        return Ok(());
    }
    charge(
        amount
            .checked_mul(element_size)
            .ok_or(MetadataError::Invalid)?,
    )
}

fn try_reserve<T>(
    vector: &mut Vec<T>,
    amount: usize,
    element_size: usize,
) -> Result<(), MetadataError> {
    if amount == 0 {
        return Ok(());
    }
    let allocation = amount
        .checked_mul(element_size)
        .ok_or(MetadataError::Invalid)?;
    vector
        .try_reserve_exact(amount)
        .map_err(|_| MetadataError::Allocation { amount: allocation })
}

fn clone_source(source: &[u8]) -> Result<Vec<u8>, MetadataError> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(source.len())
        .map_err(|_| MetadataError::Allocation {
            amount: source.len(),
        })?;
    output.extend_from_slice(source);
    Ok(output)
}

fn checked_sum<const N: usize>(parts: [usize; N]) -> Result<usize, MetadataError> {
    parts
        .into_iter()
        .try_fold(0usize, |total, part| total.checked_add(part))
        .ok_or(MetadataError::Invalid)
}

#[cfg(test)]
mod tests {
    use super::{
        ComponentIdentity, MetadataError, MetadataSnapshot, identity_codec, media_codec,
        rebind_source_bytes, rebind_source_str,
    };

    #[test]
    fn source_subslices_rebind_to_the_source_lifetime() {
        let source = b"prefix:locator:suffix";
        let locator = &source[7..14];
        let locator_str = core::str::from_utf8(locator).expect("ASCII test slice");

        assert_eq!(rebind_source_bytes(source, locator), Some(locator));
        assert_eq!(rebind_source_str(source, locator_str), Some(locator_str));
    }

    #[test]
    fn unrelated_equal_allocation_is_refused() {
        let source = b"locator";
        let unrelated = String::from("locator");

        assert!(rebind_source_bytes(source, unrelated.as_bytes()).is_none());
        assert!(rebind_source_str(source, unrelated.as_str()).is_none());
    }

    #[test]
    fn empty_values_rebind_without_pointer_arithmetic() {
        let source = b"payload";

        assert_eq!(rebind_source_bytes(source, &[]), Some(&source[..0]));
        assert_eq!(rebind_source_str(source, ""), Some(""));
        assert_eq!(rebind_source_bytes(&[], &[]), Some([].as_slice()));
    }

    fn put_varint(output: &mut Vec<u8>, mut value: u64) {
        while value >= 0x80 {
            output.push((value as u8 & 0x7f) | 0x80);
            value >>= 7;
        }
        output.push(value as u8);
    }

    fn varint_field(output: &mut Vec<u8>, number: u32, value: u64) {
        put_varint(output, u64::from(number) << 3);
        put_varint(output, value);
    }

    fn bytes_field(output: &mut Vec<u8>, number: u32, value: &[u8]) {
        put_varint(output, (u64::from(number) << 3) | 2);
        put_varint(output, value.len() as u64);
        output.extend_from_slice(value);
    }

    fn external_reference(
        target_identifier: u64,
        object_identifier: u64,
        weak: Option<bool>,
        unknown: bool,
    ) -> Vec<u8> {
        let mut reference = Vec::new();
        varint_field(&mut reference, 1, target_identifier);
        varint_field(&mut reference, 2, object_identifier);
        if let Some(weak) = weak {
            varint_field(&mut reference, 3, u64::from(weak));
        }
        if unknown {
            varint_field(&mut reference, 4, 1);
        }
        reference
    }

    fn component(
        identifier: u64,
        locator: &str,
        references: &[(u32, u64, u64, Option<bool>, bool)],
    ) -> Vec<u8> {
        let mut component = Vec::new();
        varint_field(&mut component, 1, identifier);
        bytes_field(&mut component, 2, locator.as_bytes());
        for (field, target, object, weak, unknown) in references {
            bytes_field(
                &mut component,
                *field,
                &external_reference(*target, *object, *weak, *unknown),
            );
        }
        component
    }

    fn metadata(current: &[Vec<u8>], versioned: &[Vec<u8>]) -> Vec<u8> {
        let mut source = Vec::new();
        varint_field(&mut source, 1, 100);
        for component in current {
            bytes_field(&mut source, 3, component);
        }
        for component in versioned {
            bytes_field(&mut source, 11, component);
        }
        source
    }

    fn identity_options(source: &[u8]) -> identity_codec::RewriteOptions {
        let bytes = source.len().max(1);
        identity_codec::RewriteOptions::new(
            bytes,
            bytes.saturating_mul(2).max(1),
            bytes.saturating_mul(64).max(1),
            bytes.saturating_mul(256).max(1),
            16,
            64,
            256,
            64,
        )
    }

    fn test_snapshot<'source>(source: &'source [u8]) -> MetadataSnapshot<'source> {
        let mut charge = |_amount: usize| -> Result<(), MetadataError> { Ok(()) };
        MetadataSnapshot::inspect(
            source,
            identity_options(source),
            media_codec::DecodeOptions::for_source(source),
            None,
            &mut charge,
        )
        .expect("synthetic PackageMetadata must be inspectable")
    }

    fn test_source_component<'source>(
        snapshot: &MetadataSnapshot<'source>,
    ) -> ComponentIdentity<'source> {
        snapshot
            .current_component("source")
            .expect("synthetic source component")
    }

    fn package_with_reference(field: u32, weak: Option<bool>, unknown: bool) -> Vec<u8> {
        metadata(
            &[
                component(10, "source", &[(field, 20, 99, weak, unknown)]),
                component(20, "author", &[]),
            ],
            &[],
        )
    }

    #[test]
    fn native_shaped_strong_author_edges_are_accepted_but_weak_edges_are_not() {
        for weak in [None, Some(false)] {
            let source = package_with_reference(6, weak, false);
            let snapshot = test_snapshot(&source);
            let source_component = test_source_component(&snapshot);

            assert!(matches!(
                snapshot.require_current_external_dependency(source_component, "author", 99),
                Ok(())
            ));
        }

        let source = package_with_reference(6, Some(true), false);
        let snapshot = test_snapshot(&source);
        let source_component = test_source_component(&snapshot);
        assert!(matches!(
            snapshot.require_current_external_dependency(source_component, "author", 99),
            Err(MetadataError::Ambiguous)
        ));
    }

    #[test]
    fn missing_duplicate_versioned_and_unknown_matching_edges_fail_closed() {
        let missing = metadata(
            &[component(10, "source", &[]), component(20, "author", &[])],
            &[],
        );
        let snapshot = test_snapshot(&missing);
        let source_component = test_source_component(&snapshot);
        assert!(matches!(
            snapshot.require_current_external_dependency(source_component, "author", 99),
            Err(MetadataError::Missing)
        ));

        let duplicate = metadata(
            &[
                component(
                    10,
                    "source",
                    &[(6, 20, 99, None, false), (6, 20, 99, None, false)],
                ),
                component(20, "author", &[]),
            ],
            &[],
        );
        let snapshot = test_snapshot(&duplicate);
        let source_component = test_source_component(&snapshot);
        assert!(matches!(
            snapshot.require_current_external_dependency(source_component, "author", 99),
            Err(MetadataError::Ambiguous)
        ));

        let versioned = metadata(
            &[component(10, "source", &[]), component(20, "author", &[])],
            &[component(10, "source", &[(18, 20, 99, None, false)])],
        );
        let snapshot = test_snapshot(&versioned);
        let source_component = test_source_component(&snapshot);
        assert!(matches!(
            snapshot.require_current_external_dependency(source_component, "author", 99),
            Err(MetadataError::Ambiguous)
        ));

        let unknown = package_with_reference(6, None, true);
        let snapshot = test_snapshot(&unknown);
        let source_component = test_source_component(&snapshot);
        assert!(matches!(
            snapshot.require_current_external_dependency(source_component, "author", 99),
            Err(MetadataError::Ambiguous)
        ));
    }

    #[test]
    fn unrelated_unknown_edges_do_not_poison_the_selected_dependency() {
        let source = metadata(
            &[
                component(
                    10,
                    "source",
                    &[(6, 20, 99, None, false), (6, 20, 100, None, true)],
                ),
                component(20, "author", &[]),
            ],
            &[],
        );
        let snapshot = test_snapshot(&source);
        let source_component = test_source_component(&snapshot);

        assert!(matches!(
            snapshot.require_current_external_dependency(source_component, "author", 99),
            Ok(())
        ));
        assert_eq!(snapshot.external_references.len(), 2);
        assert_eq!(snapshot.external_dependency_lookup_work(), 8);
    }

    #[test]
    fn source_and_target_locator_or_identifier_ambiguity_is_rejected() {
        let duplicate_source_locator = metadata(
            &[
                component(10, "source", &[(6, 20, 99, None, false)]),
                component(11, "source", &[]),
                component(20, "author", &[]),
            ],
            &[],
        );
        let snapshot = test_snapshot(&duplicate_source_locator);
        assert!(matches!(
            snapshot.current_component("source"),
            Err(MetadataError::Ambiguous)
        ));

        let duplicate_target_locator = metadata(
            &[
                component(10, "source", &[(6, 20, 99, None, false)]),
                component(20, "author", &[]),
                component(21, "author", &[]),
            ],
            &[],
        );
        let snapshot = test_snapshot(&duplicate_target_locator);
        let source_component = test_source_component(&snapshot);
        assert!(matches!(
            snapshot.require_current_external_dependency(source_component, "author", 99),
            Err(MetadataError::Ambiguous)
        ));

        let duplicate_source_identifier = metadata(
            &[
                component(10, "source", &[(6, 20, 99, None, false)]),
                component(10, "source-v2", &[]),
                component(20, "author", &[]),
            ],
            &[],
        );
        let snapshot = test_snapshot(&duplicate_source_identifier);
        let source_component = test_source_component(&snapshot);
        assert!(matches!(
            snapshot.require_current_external_dependency(source_component, "author", 99),
            Err(MetadataError::Ambiguous)
        ));

        let duplicate_target_identifier = metadata(
            &[
                component(10, "source", &[(6, 20, 99, None, false)]),
                component(20, "author", &[]),
                component(20, "author-v2", &[]),
            ],
            &[],
        );
        let snapshot = test_snapshot(&duplicate_target_identifier);
        let source_component = test_source_component(&snapshot);
        assert!(matches!(
            snapshot.require_current_external_dependency(source_component, "author", 99),
            Err(MetadataError::Ambiguous)
        ));
    }

    #[test]
    fn dependency_lookup_work_is_bounded_by_component_and_edge_counts() {
        let source = package_with_reference(6, None, false);
        let snapshot = test_snapshot(&source);

        assert_eq!(snapshot.components.len(), 2);
        assert_eq!(snapshot.external_references.len(), 1);
        assert_eq!(snapshot.external_dependency_lookup_work(), 7);
    }
}
