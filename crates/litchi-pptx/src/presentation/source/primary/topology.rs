//! Structural revisions for the source-backed primary presentation owner.
//!
//! The ordinary primary path retains one presentation root and a set of
//! changed slide payloads.  A slide removal changes the OPC catalog itself,
//! so it is kept as a prepared topology together with a small declarative
//! recipe.  The recipe is relative to the retained physical source; the
//! prepared candidate supplies lazy readback and source-preserving
//! publication.  Keeping those two pieces together is what lets a package
//! continue to read a restored slide after a durable inverse is applied to a
//! ZIP from which the slide Part is physically absent.

use std::collections::{HashMap, HashSet};
use std::sync::{Arc, OnceLock};

use litchi_core::{Position, Reservation, Resource};
use litchi_opc::xml_splice::authored_xml_requires_source_proof;
use litchi_opc::{
    PackURI, Part, PartData, PreparedMetadataData, PreparedPartData, PreparedTopology,
    Relationships, SourceTopologyPlan, TargetMode,
};

use super::super::{
    PartBinding, RelationshipBinding, SlideBinding, SourceBackedSlideSnapshot, SourcePayload,
};
use super::topology_digest;
use super::topology_wire::{
    Change, ChangeLengths, Data, Dependent, DependentLengths, ImageLengths, PartImage, State,
    StateLengths,
};
use super::{Package, RootPayload, allocation, limit};
use crate::presentation::order::{StableSlideKey, stable_slide_key};
use crate::{Error, Result};

/// A structural patch wrapper used by the primary patch enum.
///
/// `prepared` is process-local and is deliberately omitted from the durable
/// wire representation.  It keeps source and managed-resource admissions
/// alive for a local package and lets the owner reuse the already prepared
/// candidate when possible.  A decoded patch has no prepared candidate and
/// rebuilds one from `data` against the caller's physical source.
#[derive(Clone)]
pub(super) struct Patch {
    pub(super) data: Data,
    pub(super) prepared: Option<PreparedTopology>,
    pub(super) retained: Arc<Retention>,
    encoded_len_cache: Arc<OnceLock<usize>>,
}

impl Patch {
    pub(super) fn from_data(data: Data) -> Self {
        Self {
            data,
            prepared: None,
            retained: Arc::new(Retention::default()),
            encoded_len_cache: Arc::new(OnceLock::new()),
        }
    }

    pub(super) fn inverse(&self) -> Self {
        Self {
            data: self.data.inverse(),
            prepared: self.prepared.clone(),
            retained: Arc::clone(&self.retained),
            encoded_len_cache: Arc::clone(&self.encoded_len_cache),
        }
    }

    pub(super) fn is_noop(&self) -> bool {
        self.data.is_noop()
    }

    pub(super) fn check(&self) -> Result<()> {
        if let Some(prepared) = &self.prepared {
            prepared.check_source()?;
            prepared.check_execution()?;
        }
        Ok(())
    }

    pub(super) fn encoded_len(&self) -> Result<usize> {
        self.check()?;
        if let Some(length) = self.encoded_len_cache.get() {
            return Ok(*length);
        }
        let length = self.data.encoded_len()?;
        self.check()?;
        let _ = self.encoded_len_cache.set(length);
        Ok(length)
    }

    pub(super) fn to_bytes(&self) -> Result<Vec<u8>> {
        if let Some(prepared) = &self.prepared {
            prepared.check_source()?;
            prepared.check_execution()?;
        }
        let bytes = self.data.to_bytes()?;
        if let Some(prepared) = &self.prepared {
            prepared.check_source()?;
            prepared.check_execution()?;
        }
        Ok(bytes)
    }

    pub(super) fn from_bytes(bytes: &[u8]) -> Result<Self> {
        Ok(Self::from_data(Data::from_bytes(
            bytes,
            litchi_opc::SourceLimits::default(),
        )?))
    }

    pub(super) fn apply(&self, package: &Package, shape_removals: bool) -> Result<Package> {
        apply(package, &self.data, shape_removals)
    }
}

/// Prepared structural state retained by one immutable primary package.
pub(super) struct Revision {
    pub(super) prepared: PreparedTopology,
    pub(super) data: Data,
    pub(super) retained: Arc<Retention>,
}

#[derive(Default, Clone)]
pub(super) struct Retention {
    // These handles intentionally remain owned even though the wire state
    // carries copied Arc bytes.  On managed sources they keep the source or
    // prepared candidate reservations attached to the logical revision.
    pub(super) source_parts: Vec<PartData>,
    pub(super) prepared_parts: Vec<PreparedPartData>,
    pub(super) source_metadata: Vec<PartData>,
    pub(super) prepared_metadata: Vec<PreparedMetadataData>,
    pub(super) reservations: Vec<Arc<Reservation>>,
    pub(super) source_xml: Vec<litchi_opc::SourceXmlPart>,
}

#[derive(Clone)]
enum SourceSpec {
    Part(PackURI),
    Relationships(PackURI),
    ContentTypes,
}

#[derive(Clone)]
enum PlannedBytes {
    Source(PartData),
    Declared {
        spec: SourceSpec,
        length: usize,
    },
    Length(usize),
    Owned {
        bytes: Arc<Vec<u8>>,
        reservation: Option<Arc<Reservation>>,
    },
}

impl PlannedBytes {
    fn as_bytes(&self) -> &[u8] {
        match self {
            Self::Source(data) => data.as_bytes(),
            Self::Declared { .. } => &[],
            Self::Length(_) => &[],
            Self::Owned { bytes, .. } => bytes.as_slice(),
        }
    }

    fn lengths(&self) -> usize {
        match self {
            Self::Declared { length, .. } => *length,
            Self::Length(length) => *length,
            _ => self.as_bytes().len(),
        }
    }

    fn same(&self, other: &Self) -> bool {
        match (self, other) {
            (Self::Declared { spec: left, .. }, Self::Declared { spec: right, .. }) => {
                source_spec_equal(left, right)
            },
            (Self::Source(left), Self::Source(right)) => left.as_bytes() == right.as_bytes(),
            (Self::Owned { bytes: left, .. }, Self::Owned { bytes: right, .. }) => {
                left.as_slice() == right.as_slice()
            },
            _ => self.as_bytes() == other.as_bytes() && self.lengths() == other.lengths(),
        }
    }
}

fn source_spec_equal(left: &SourceSpec, right: &SourceSpec) -> bool {
    match (left, right) {
        (SourceSpec::Part(left), SourceSpec::Part(right))
        | (SourceSpec::Relationships(left), SourceSpec::Relationships(right)) => left == right,
        (SourceSpec::ContentTypes, SourceSpec::ContentTypes) => true,
        _ => false,
    }
}

#[derive(Clone)]
struct PlannedImage {
    xml: PlannedBytes,
    relationships: Option<PlannedBytes>,
}

impl PlannedImage {
    fn same(&self, other: &Self) -> bool {
        self.xml.same(&other.xml)
            && match (&self.relationships, &other.relationships) {
                (None, None) => true,
                (Some(left), Some(right)) => left.same(right),
                _ => false,
            }
    }

    fn lengths(&self) -> ImageLengths {
        ImageLengths {
            xml: self.xml.lengths(),
            relationships: self.relationships.as_ref().map(PlannedBytes::lengths),
        }
    }
}

#[derive(Clone)]
struct PlannedState {
    root: PlannedBytes,
    order: Arc<[StableSlideKey]>,
    topology_digest: [u8; 32],
    presentation_relationships: PlannedBytes,
    content_types: PlannedBytes,
}

impl PlannedState {
    fn lengths(&self) -> StateLengths {
        StateLengths {
            root: self.root.lengths(),
            presentation_relationships: self.presentation_relationships.lengths(),
            content_types: self.content_types.lengths(),
        }
    }
}

#[derive(Clone)]
pub(super) struct PlannedChange {
    key: StableSlideKey,
    content_type: String,
    relationship_type: String,
    before: Option<PlannedImage>,
    after: Option<PlannedImage>,
}

#[derive(Clone)]
pub(super) struct PlannedDependent {
    part_uri: PackURI,
    content_type: String,
    before: Option<PlannedImage>,
    after: Option<PlannedImage>,
}

struct PatchPlan {
    before: PlannedState,
    after: PlannedState,
    changes: Vec<PlannedChange>,
    dependents: Vec<PlannedDependent>,
}

impl PatchPlan {
    fn encoded_len(&self) -> Result<usize> {
        let changes = self
            .changes
            .iter()
            .map(|change| ChangeLengths {
                key: change.key.clone(),
                content_type: change.content_type.clone(),
                relationship_type: change.relationship_type.clone(),
                before: change.before.as_ref().map(PlannedImage::lengths),
                after: change.after.as_ref().map(PlannedImage::lengths),
            })
            .collect::<Vec<_>>();
        let dependents = self
            .dependents
            .iter()
            .map(|dependent| DependentLengths {
                part_uri: dependent.part_uri.clone(),
                content_type: dependent.content_type.clone(),
                before: dependent.before.as_ref().map(PlannedImage::lengths),
                after: dependent.after.as_ref().map(PlannedImage::lengths),
            })
            .collect::<Vec<_>>();
        super::topology_wire::encoded_len_from_lengths(
            &self.before.order,
            &self.after.order,
            self.before.lengths(),
            self.after.lengths(),
            &changes,
            &dependents,
        )
    }
}

/// Reserve the durable envelope's optional removal marker before a topology
/// candidate retains any source payload.  The marker is outside the topology
/// body, so topology wire planning receives only the remaining budget.
pub(super) fn topology_patch_budget(max_patch_bytes: usize, shape_removals: bool) -> Result<usize> {
    let marker = if shape_removals {
        super::durable::REMOVAL_MARKER_BYTES
    } else {
        0
    };
    if marker > max_patch_bytes {
        return Err(limit(
            crate::error::CodecLimitKind::PatchBytes,
            marker,
            max_patch_bytes,
        ));
    }
    Ok(max_patch_bytes - marker)
}

impl Revision {
    pub(super) fn check(&self) -> Result<()> {
        self.prepared.check_source()?;
        self.prepared.check_execution()?;
        Ok(())
    }

    /// Read one effective slide by stable identity, including a Part restored
    /// by an inverse after the physical source ZIP no longer contains it.
    pub(super) fn snapshot_for_key(
        &self,
        package: &Package,
        key: &StableSlideKey,
        logical_position: usize,
    ) -> Result<SourceBackedSlideSnapshot> {
        self.check()?;
        let part = self.prepared.part(&key.part_uri)?;
        if !part.partname().is_equivalent_to(&key.part_uri) {
            return Err(Error::PatchConflict);
        }
        let data = part.data()?;
        let bytes = data.as_bytes();
        let limits = package.owner.limits;
        crate::source_limits::validate_xml_limits(bytes, limits)?;
        if bytes.len() > limits.max_part_bytes() as usize {
            return Err(limit(
                crate::error::CodecLimitKind::PartBytes,
                bytes.len(),
                limits.max_part_bytes() as usize,
            ));
        }
        let semantic = crate::open_options::selected(package.owner.semantic_limits.as_ref());
        let laser_limits = semantic.laser().constrained_by_source(limits);
        // Refuse malformed or unsupported markup-compatibility and selected
        // shape/transition policy failures before reserving or copying any
        // edit-owned payload.  The admitted mode is retained by the snapshot
        // so enclosed whole-shape removal keeps its supported capability.
        let mce = crate::presentation::transition::admit_source_with_limits(
            bytes,
            "slide_snapshot",
            semantic.shape(),
            semantic.transition(),
        )?;

        // Admission must precede the owned copy.  The source PartData remains
        // attached to the prepared candidate; this extra reservation covers
        // the edit-owned payload that is handed to the public snapshot.
        let mut amount = bytes.len().checked_add(1024).ok_or_else(|| {
            limit(
                crate::error::CodecLimitKind::Snapshots,
                usize::MAX,
                usize::MAX,
            )
        })?;
        for size in [
            key.part_uri.as_str().len(),
            key.relationship_id.len(),
            part.content_type().len(),
        ]
        .into_iter()
        .chain(part.rels().iter().flat_map(|rel| {
            [
                rel.r_id().len(),
                rel.reltype().len(),
                rel.target_ref().len(),
                128,
            ]
        })) {
            package.owner.package.check_execution()?;
            amount = size
                .checked_mul(2)
                .and_then(|size| amount.checked_add(size))
                .ok_or_else(|| {
                    limit(
                        crate::error::CodecLimitKind::Snapshots,
                        usize::MAX,
                        usize::MAX,
                    )
                })?;
        }
        let reservation = reserve_payload(package, amount)?;
        let copied = copy_bytes(bytes)?;
        let payload = match reservation {
            Some(reservation) => SourcePayload::ReservedEdit {
                bytes: copied,
                reservation: Arc::new(reservation),
            },
            None => SourcePayload::Edited(copied),
        };

        let binding = effective_slide_binding(package, self, key, &part)?;
        let source_version = package.owner.package.source_version()?;
        let context = package.owner.package.execution_context();
        Ok(SourceBackedSlideSnapshot {
            position: logical_position,
            part_uri: key.part_uri.clone(),
            xml: payload,
            closure: super::super::SlideClosure {
                package_relationships: package.owner.package_relationships.clone(),
                presentation: package.owner.presentation_binding.clone(),
                // The effective root is guarded by Revision::data.  The
                // existing slide readers only need the source presentation
                // binding and retain this PartData for the edit closures.
                presentation_xml: package.owner._presentation.data.clone(),
                slide: Arc::new(binding),
            },
            max_output_bytes: super::super::source_slide_output_limit(limits, semantic.shape()),
            protection_limits:
                crate::presentation_properties::metadata::protection::Limits::from_source_limits(
                    limits,
                ),
            semantic_limits: package.owner.semantic_limits.clone(),
            source_version,
            lineage: package.owner.package.source_lineage(),
            context,
            laser_limits,
            laser_usage: Arc::new(OnceLock::new()),
            // A prepared slide that encloses markup compatibility in shape
            // children admits only whole-shape removal (record 0303); other
            // slides keep the edit-time guards of each operation.
            mce,
        })
    }
}

/// Remove one selected stable slide after the standalone source planner has
/// admitted the physical semantic surface.  This function constructs the
/// final target from the currently staged root/order and rebuilds one prepared
/// topology, so previous order and text edits remain part of the same atomic
/// candidate.
pub(super) fn remove(
    package: &Package,
    key: &StableSlideKey,
    max_patch_bytes: usize,
) -> Result<Package> {
    package.check()?;
    let logical = package
        .order
        .iter()
        .position(|candidate| candidate == key)
        .ok_or(Error::PatchConflict)?;
    if package.order.len() < 2 {
        return Err(Error::slide_removal_refused(
            crate::SlideRemovalRefusal::FinalSlide,
        ));
    }
    validate_removal_admission(package, key, false, false)?;

    let bindings = package
        .order
        .iter()
        .map(|candidate| (candidate.slide_id, candidate.relationship_id.as_str()));
    let root = crate::opened::remove_slide_binding(
        package.root.as_bytes(),
        bindings,
        key.slide_id,
        package.owner.package.execution_context().as_ref(),
    )?;
    let mut order = package.order.to_vec();
    order.remove(logical);
    let mut planned_changes = planned_current_changes(package)?;
    upsert_planned_change(&mut planned_changes, planned_remove_change(package, key)?);
    let owned = owned_parts(package, std::slice::from_ref(&key.part_uri))?;
    let mut planned_dependents = planned_current_dependents(package);
    for part_uri in owned.parts() {
        let (before, content_type) = match planned_physical_dependent(package, part_uri)? {
            Some((image, content_type)) => (Some(image), content_type),
            None => {
                let existing = planned_dependents
                    .iter()
                    .find(|dependent| dependent.part_uri.is_equivalent_to(part_uri))
                    .ok_or(Error::PatchConflict)?;
                (
                    existing.after.clone().or_else(|| existing.before.clone()),
                    existing.content_type.clone(),
                )
            },
        };
        upsert_planned_dependent(
            &mut planned_dependents,
            PlannedDependent {
                part_uri: part_uri.clone(),
                content_type,
                before,
                after: None,
            },
        );
    }
    preflight_candidate(
        package,
        &root,
        order.clone().into(),
        Some(&planned_changes),
        Some(&planned_dependents),
        max_patch_bytes,
    )?;
    let _root_reservation = reserve_payload(package, root.len())?;
    let mut retained = Retention::default();
    let mut changes = current_changes(package, &mut retained)?;
    let physical = physical_image(package, key, &mut retained)?;
    let (before, content_type, relationship_type) = match physical {
        Some((image, content_type, relationship_type)) => {
            (Some(image), content_type, relationship_type)
        },
        None => {
            let existing = changes
                .iter()
                .find(|change| &change.key == key)
                .ok_or(Error::PatchConflict)?;
            (
                None,
                existing.content_type.clone(),
                existing.relationship_type.clone(),
            )
        },
    };
    upsert_change(
        &mut changes,
        Change {
            key: key.clone(),
            content_type,
            relationship_type,
            before,
            after: None,
        },
    );
    // The Parts the slide owns exclusively leave with it (record 0303).
    let mut dependents = current_dependents(package);
    for part_uri in owned.parts() {
        let (before, content_type) = match physical_dependent(package, part_uri, &mut retained)? {
            Some((image, content_type)) => (Some(image), content_type),
            None => {
                let existing = dependents
                    .iter()
                    .find(|dependent| dependent.part_uri.is_equivalent_to(part_uri))
                    .ok_or(Error::PatchConflict)?;
                (None, existing.content_type.clone())
            },
        };
        upsert_dependent(
            &mut dependents,
            Dependent {
                part_uri: part_uri.clone(),
                content_type,
                before,
                after: None,
            },
        );
    }
    let revision = build_revision_with_retention(
        package,
        Arc::new(root),
        order.into(),
        changes.into(),
        dependents.into(),
        None,
        retained,
    )?;
    package_from_revision(package, revision)
}

/// Insert one newly authored dependency-free slide into the current effective
/// catalog. The authoring helper has already performed eager-compatible layout,
/// root, source-policy, and identity admission; this function composes its
/// root/order/change with every earlier structural or slide-text revision and
/// prepares one typed OPC candidate.
pub(super) fn insert(
    package: &Package,
    position: Position,
    layout_name: &str,
    max_patch_bytes: usize,
) -> Result<Package> {
    package.check()?;
    let (root, order, change, insertion_reservation) =
        super::insertion::plan(package, position, layout_name, Some(max_patch_bytes))?;
    let mut retained = Retention::default();
    if let Some(reservation) = insertion_reservation {
        retained.reservations.push(reservation);
    }
    let mut changes = current_changes(package, &mut retained)?;
    upsert_change(&mut changes, change);
    let revision = build_revision_with_retention(
        package,
        root,
        order,
        changes.into(),
        current_dependents(package).into(),
        None,
        retained,
    )?;
    package_from_revision(package, revision)
}

/// Rebuild a structural candidate after a logical order change.
pub(super) fn reorder(
    package: &Package,
    root: Arc<Vec<u8>>,
    order: Arc<[StableSlideKey]>,
    max_patch_bytes: usize,
) -> Result<Package> {
    let revision = package.topology.as_ref().ok_or(Error::PatchConflict)?;
    preflight_candidate(
        package,
        &root,
        Arc::clone(&order),
        None,
        None,
        max_patch_bytes,
    )?;
    let mut retained = Retention::default();
    if let Some(reservation) = super::reserve_managed_output(
        package.owner.package.execution_context().as_ref(),
        root.len(),
    )? {
        retained.reservations.push(reservation);
    }
    let revision = build_revision_with_retention(
        package,
        root,
        order,
        Arc::clone(&revision.data.changes),
        Arc::clone(&revision.data.dependents),
        Some(&revision.data.after),
        retained,
    )?;
    package_from_revision(package, revision)
}

/// Rebuild a structural candidate after an existing effective slide XML was
/// changed by the ordinary shape-text editor.
pub(super) fn replace_snapshot(
    package: &Package,
    snapshot: SourceBackedSlideSnapshot,
    max_patch_bytes: usize,
) -> Result<Package> {
    let revision = package.topology.as_ref().ok_or(Error::PatchConflict)?;
    let key = package
        .order
        .iter()
        .find(|candidate| candidate.part_uri == snapshot.part_uri)
        .cloned()
        .ok_or(Error::PatchConflict)?;
    let mut changes = revision.data.changes.to_vec();
    let existing = changes.iter().find(|change| change.key == key).cloned();
    let (before_lengths, content_type, relationship_type, relationship_lengths) =
        if let Some(change) = existing.as_ref() {
            (
                change.before.as_ref().map(image_lengths),
                change.content_type.clone(),
                change.relationship_type.clone(),
                change
                    .after
                    .as_ref()
                    .and_then(|image| image.relationships.as_ref().map(|bytes| bytes.len())),
            )
        } else {
            let (image, content_type, relationship_type) =
                planned_physical_image(package, &key)?.ok_or(Error::PatchConflict)?;
            let image_lengths = image.lengths();
            (
                Some(image_lengths),
                content_type,
                relationship_type,
                image_lengths.relationships,
            )
        };
    let after_lengths = ImageLengths {
        xml: snapshot.xml.as_bytes().len(),
        relationships: relationship_lengths,
    };
    let candidate_wire = revision.data.encoded_len_with_replacement(
        &key,
        before_lengths,
        Some(after_lengths),
        &content_type,
        &relationship_type,
    )?;
    if candidate_wire > max_patch_bytes {
        return Err(limit(
            crate::error::CodecLimitKind::PatchBytes,
            candidate_wire,
            max_patch_bytes,
        ));
    }
    let mut seed = Retention::default();
    let (content_type, relationship_type, before, relationships) = if let Some(change) = existing {
        (
            change.content_type,
            change.relationship_type,
            change.before,
            change.after.and_then(|image| image.relationships),
        )
    } else {
        let image = physical_image(package, &key, &mut seed)?.ok_or(Error::PatchConflict)?;
        (
            image.1,
            image.2,
            Some(image.0.clone()),
            image.0.relationships,
        )
    };
    let xml = match snapshot.xml {
        SourcePayload::Original(data) => copy_state_bytes(package, data.as_bytes(), &mut seed)?,
        SourcePayload::Edited(bytes) => bytes,
        SourcePayload::ReservedEdit { bytes, reservation } => {
            seed.reservations.push(reservation);
            bytes
        },
    };
    upsert_change(
        &mut changes,
        Change {
            key,
            content_type,
            relationship_type,
            before,
            after: Some(PartImage { xml, relationships }),
        },
    );
    let root = match package.root.edited_bytes() {
        Some(bytes) => bytes,
        None => copy_state_bytes(package, package.root.as_bytes(), &mut seed)?,
    };
    let revision = build_revision_with_retention(
        package,
        root,
        Arc::clone(&package.order),
        changes.into(),
        Arc::clone(&revision.data.dependents),
        Some(&revision.data.after),
        seed,
    )?;
    let encoded = Data::encoded_len_from_data(&revision.data)?;
    if encoded > max_patch_bytes {
        return Err(limit(
            crate::error::CodecLimitKind::PatchBytes,
            encoded,
            max_patch_bytes,
        ));
    }
    package_from_revision(package, revision)
}

/// Return the exact remaining XML budget for replacing one slide in a
/// topology revision.  Existing staged changes remain part of the candidate
/// wire image; only the replacement's XML payload is left variable.  This is
/// used before the native editor allocates its candidate output.
pub(super) fn replacement_after_budget(
    package: &Package,
    key: &StableSlideKey,
    max_patch_bytes: usize,
) -> Result<usize> {
    let revision = package.topology.as_ref().ok_or(Error::PatchConflict)?;
    let existing = revision
        .data
        .changes
        .iter()
        .find(|change| &change.key == key);
    let (before_lengths, content_type, relationship_type, relationship_lengths) =
        if let Some(change) = existing {
            (
                change.before.as_ref().map(image_lengths),
                change.content_type.clone(),
                change.relationship_type.clone(),
                change
                    .after
                    .as_ref()
                    .and_then(|image| image.relationships.as_ref().map(|bytes| bytes.len())),
            )
        } else {
            let (image, content_type, relationship_type) =
                planned_physical_image(package, key)?.ok_or(Error::PatchConflict)?;
            let image_lengths = image.lengths();
            (
                Some(image_lengths),
                content_type,
                relationship_type,
                image_lengths.relationships,
            )
        };
    let baseline = revision.data.encoded_len_with_replacement(
        key,
        before_lengths,
        Some(ImageLengths {
            xml: 0,
            relationships: relationship_lengths,
        }),
        &content_type,
        &relationship_type,
    )?;
    if baseline > max_patch_bytes {
        return Err(limit(
            crate::error::CodecLimitKind::PatchBytes,
            baseline,
            max_patch_bytes,
        ));
    }
    Ok(max_patch_bytes - baseline)
}

fn plan_patch(source: &Package, target: &Package) -> Result<PatchPlan> {
    let source_changes = planned_current_changes(source)?;
    let target_changes = planned_current_changes(target)?;
    let changes = planned_diff_changes(source, target, source_changes, target_changes)?;
    let dependents = planned_diff_dependents(source, target)?;
    Ok(PatchPlan {
        before: planned_state(source, true)?,
        after: planned_state(target, true)?,
        changes,
        dependents,
    })
}

fn preflight_candidate(
    package: &Package,
    root: &[u8],
    order: Arc<[StableSlideKey]>,
    candidate_changes: Option<&[PlannedChange]>,
    candidate_dependents: Option<&[PlannedDependent]>,
    max_patch_bytes: usize,
) -> Result<usize> {
    let source_changes = planned_current_changes(package)?;
    let target_changes = match candidate_changes {
        Some(changes) => changes.to_vec(),
        None => source_changes.clone(),
    };
    let changes = planned_diff_changes(package, package, source_changes, target_changes.clone())?;
    let source_dependents = planned_current_dependents(package);
    let target_dependents = match candidate_dependents {
        Some(dependents) => dependents.to_vec(),
        None => source_dependents.clone(),
    };
    let dependents = planned_diff_dependents_from_entries(
        package,
        source_dependents,
        target_dependents.clone(),
    )?;
    // Length-only admission needs the fixed digest field width, not the
    // source catalog proof.  The proof is generated once when the durable
    // patch is materialized below.
    let before = planned_state(package, false)?;
    let mut after = before.clone();
    after.root = PlannedBytes::Length(root.len());
    after.order = order;
    if candidate_changes.is_some() || candidate_dependents.is_some() {
        if let Some((relationships, content_types)) =
            candidate_metadata_lengths(package, &target_changes, &target_dependents)?
        {
            after.presentation_relationships = PlannedBytes::Length(relationships);
            after.content_types = PlannedBytes::Length(content_types);
        }
    }
    let plan = PatchPlan {
        before,
        after,
        changes,
        dependents,
    };
    let encoded = plan.encoded_len()?;
    if encoded > max_patch_bytes {
        return Err(limit(
            crate::error::CodecLimitKind::PatchBytes,
            encoded,
            max_patch_bytes,
        ));
    }
    Ok(encoded)
}

/// Plan only the metadata operations of a candidate topology.
///
/// Structural admission must know the exact relationship and content-types
/// images before it retains any source Part payload.  The full topology
/// builder performs the same operations while capturing those payloads; this
/// length-only mirror deliberately stages only additions/removals and typed
/// relationship changes, using already-owned candidate images for newly
/// restored Parts.  The native planner then produces the exact lexical
/// metadata lengths, including dependent Parts and their relationship edges.
fn candidate_metadata_lengths(
    package: &Package,
    changes: &[PlannedChange],
    dependents: &[PlannedDependent],
) -> Result<Option<(usize, usize)>> {
    let presentation_uri = package.owner._presentation.source_partname().clone();
    let mut plan = SourceTopologyPlan::new();
    let mut changed = false;

    for change in changes {
        let physical = match package.owner.package.part(&change.key.part_uri) {
            Ok(view) => Some(view),
            Err(litchi_opc::OpcError::PartNotFound(_)) => None,
            Err(error) => return Err(Error::Opc(error)),
        };
        match change.after.as_ref() {
            None => {
                if physical.is_some() {
                    plan.try_remove_part(change.key.part_uri.clone())?;
                    changed = true;
                }
                if package
                    .owner
                    ._presentation
                    .rels()
                    .get(&change.key.relationship_id)
                    .is_some()
                {
                    plan.try_remove_relationship(
                        presentation_uri.clone(),
                        change.key.relationship_id.clone(),
                    )?;
                    changed = true;
                }
            },
            Some(image) => {
                if let Some(view) = &physical {
                    if view.content_type() != change.content_type.as_str() {
                        return Err(Error::PatchConflict);
                    }
                } else {
                    plan.try_add_restored_part_shared(
                        change.key.part_uri.clone(),
                        change.content_type.clone(),
                        planned_owned_bytes(&image.xml)?,
                    )?;
                    if let Some(relationships) = &image.relationships {
                        let relationships = planned_owned_bytes(relationships)?;
                        super::topology_admission::stage_layout(
                            package,
                            &mut plan,
                            &change.key.part_uri,
                            relationships.as_slice(),
                        )?;
                        plan.try_restore_relationships_xml(
                            change.key.part_uri.clone(),
                            None,
                            Some(relationships),
                        )?;
                    }
                    changed = true;
                }
                let relation = package
                    .owner
                    ._presentation
                    .rels()
                    .get(&change.key.relationship_id);
                if let Some(relation) = relation {
                    let target_matches = relation
                        .target_partname()
                        .map(|target| target.is_equivalent_to(&change.key.part_uri))
                        .unwrap_or(false);
                    if relation.is_external()
                        || relation.reltype() != change.relationship_type
                        || !crate::parts::is_relationship_type(
                            relation.reltype(),
                            litchi_opc::constants::relationship_type::SLIDE,
                            "slide",
                        )
                        || !target_matches
                    {
                        plan.try_replace_internal_relationship(
                            presentation_uri.clone(),
                            change.key.relationship_id.clone(),
                            change.relationship_type.clone(),
                            change.key.part_uri.clone(),
                        )?;
                        changed = true;
                    }
                } else {
                    plan.try_add_internal_relationship(
                        presentation_uri.clone(),
                        change.key.relationship_id.clone(),
                        change.relationship_type.clone(),
                        change.key.part_uri.clone(),
                    )?;
                    changed = true;
                }
            },
        }
    }

    for dependent in dependents {
        let physical = match package.owner.package.part(&dependent.part_uri) {
            Ok(view) => Some(view),
            Err(litchi_opc::OpcError::PartNotFound(_)) => None,
            Err(error) => return Err(Error::Opc(error)),
        };
        match (&dependent.after, physical) {
            (None, Some(_)) => {
                plan.try_remove_part(dependent.part_uri.clone())?;
                changed = true;
            },
            (None, None) => {},
            (Some(_image), Some(view)) => {
                if view.content_type() != dependent.content_type.as_str() {
                    return Err(Error::PatchConflict);
                }
            },
            (Some(image), None) => {
                plan.try_add_restored_part_shared(
                    dependent.part_uri.clone(),
                    dependent.content_type.clone(),
                    planned_owned_bytes(&image.xml)?,
                )?;
                if let Some(relationships) = &image.relationships {
                    let relationships = planned_owned_bytes(relationships)?;
                    super::topology_admission::stage_dependent_edges(
                        package,
                        &mut plan,
                        &dependent.part_uri,
                        relationships.as_slice(),
                    )?;
                    plan.try_restore_relationships_xml(
                        dependent.part_uri.clone(),
                        None,
                        Some(relationships),
                    )?;
                }
                changed = true;
            },
        }
    }

    if !changed {
        return Ok(None);
    }
    let lengths = package
        .owner
        .package
        .topology_metadata_lengths(&presentation_uri, plan)?;
    Ok(Some(lengths))
}

fn planned_owned_bytes(bytes: &PlannedBytes) -> Result<Arc<Vec<u8>>> {
    match bytes {
        PlannedBytes::Owned { bytes, .. } => Ok(Arc::clone(bytes)),
        PlannedBytes::Source(_) | PlannedBytes::Declared { .. } | PlannedBytes::Length(_) => {
            Err(Error::PatchConflict)
        },
    }
}

pub(super) fn planned_change_for_admission(change: &Change) -> Result<PlannedChange> {
    planned_change(change)
}

pub(super) fn planned_current_changes_for_admission(
    package: &Package,
) -> Result<Vec<PlannedChange>> {
    planned_current_changes(package)
}

pub(super) fn planned_current_dependents_for_admission(package: &Package) -> Vec<PlannedDependent> {
    planned_current_dependents(package)
}

pub(super) fn preflight_candidate_for_admission(
    package: &Package,
    root: &[u8],
    order: Arc<[StableSlideKey]>,
    changes: &[PlannedChange],
    dependents: &[PlannedDependent],
    max_patch_bytes: usize,
) -> Result<usize> {
    preflight_candidate(
        package,
        root,
        order,
        Some(changes),
        Some(dependents),
        max_patch_bytes,
    )
}

fn planned_state(package: &Package, include_source_digest: bool) -> Result<PlannedState> {
    if let Some(revision) = &package.topology {
        revision.check()?;
        return Ok(PlannedState {
            root: PlannedBytes::Owned {
                bytes: Arc::clone(&revision.data.after.root),
                reservation: None,
            },
            order: Arc::clone(&revision.data.after.order),
            topology_digest: revision.data.after.topology_digest,
            presentation_relationships: PlannedBytes::Owned {
                bytes: Arc::clone(&revision.data.after.presentation_relationships),
                reservation: None,
            },
            content_types: PlannedBytes::Owned {
                bytes: Arc::clone(&revision.data.after.content_types),
                reservation: None,
            },
        });
    }
    let (presentation_relationships, content_types) = planned_source_metadata(package)?;
    Ok(PlannedState {
        root: planned_root(&package.root),
        order: Arc::clone(&package.order),
        topology_digest: if include_source_digest {
            topology_digest::source_digest(&package.owner.package)?
        } else {
            [0; 32]
        },
        presentation_relationships,
        content_types,
    })
}

fn planned_root(root: &RootPayload) -> PlannedBytes {
    match root {
        RootPayload::Original(data) => PlannedBytes::Source(data.clone()),
        RootPayload::Edited(bytes) => PlannedBytes::Owned {
            bytes: Arc::clone(bytes),
            reservation: None,
        },
        RootPayload::ReservedEdit {
            bytes,
            _reservation,
        } => PlannedBytes::Owned {
            bytes: Arc::clone(bytes),
            reservation: Some(Arc::clone(_reservation)),
        },
    }
}

fn planned_source_payload(payload: &SourcePayload) -> PlannedBytes {
    match payload {
        SourcePayload::Original(data) => PlannedBytes::Source(data.clone()),
        SourcePayload::Edited(bytes) => PlannedBytes::Owned {
            bytes: Arc::clone(bytes),
            reservation: None,
        },
        SourcePayload::ReservedEdit { bytes, reservation } => PlannedBytes::Owned {
            bytes: Arc::clone(bytes),
            reservation: Some(Arc::clone(reservation)),
        },
    }
}

fn declared_bytes(package: &Package, spec: SourceSpec) -> Result<PlannedBytes> {
    let length = match &spec {
        SourceSpec::Part(uri) => package
            .owner
            .package
            .part(uri)?
            .declared_uncompressed_size()?,
        SourceSpec::Relationships(uri) => package
            .owner
            .package
            .declared_relationships_uncompressed_size(uri)?
            .ok_or(Error::PatchConflict)?,
        SourceSpec::ContentTypes => package
            .owner
            .package
            .declared_content_types_uncompressed_size()?,
    };
    declared_bytes_with_length(spec, length)
}

fn declared_bytes_with_length(spec: SourceSpec, length: u64) -> Result<PlannedBytes> {
    let length = usize::try_from(length).map_err(|_| {
        limit(
            crate::error::CodecLimitKind::PatchBytes,
            usize::MAX,
            usize::MAX,
        )
    })?;
    Ok(PlannedBytes::Declared { spec, length })
}

fn planned_source_metadata(package: &Package) -> Result<(PlannedBytes, PlannedBytes)> {
    let root = package.owner._presentation.source_partname();
    Ok((
        declared_bytes(package, SourceSpec::Relationships(root.clone()))?,
        declared_bytes(package, SourceSpec::ContentTypes)?,
    ))
}

fn planned_current_changes(package: &Package) -> Result<Vec<PlannedChange>> {
    if let Some(revision) = &package.topology {
        return revision.data.changes.iter().map(planned_change).collect();
    }
    let mut changes = Vec::new();
    for snapshot in package.overlays.iter() {
        let key = super::source_key_for_position(package, snapshot.position)?;
        let (before, content_type, relationship_type) =
            planned_physical_image(package, &key)?.ok_or(Error::PatchConflict)?;
        changes.push(PlannedChange {
            key,
            content_type,
            relationship_type,
            before: Some(before.clone()),
            after: Some(PlannedImage {
                xml: planned_source_payload(&snapshot.xml),
                relationships: before.relationships,
            }),
        });
    }
    Ok(changes)
}

fn planned_change(change: &Change) -> Result<PlannedChange> {
    Ok(PlannedChange {
        key: change.key.clone(),
        content_type: change.content_type.clone(),
        relationship_type: change.relationship_type.clone(),
        before: change.before.as_ref().map(planned_image),
        after: change.after.as_ref().map(planned_image),
    })
}

fn planned_image(image: &PartImage) -> PlannedImage {
    PlannedImage {
        xml: PlannedBytes::Owned {
            bytes: Arc::clone(&image.xml),
            reservation: None,
        },
        relationships: image
            .relationships
            .as_ref()
            .map(|bytes| PlannedBytes::Owned {
                bytes: Arc::clone(bytes),
                reservation: None,
            }),
    }
}

fn planned_physical_image(
    package: &Package,
    key: &StableSlideKey,
) -> Result<Option<(PlannedImage, String, String)>> {
    let Some(slide) = package.owner.slides.iter().find(|slide| {
        stable_slide_key(
            slide.part_uri.clone(),
            slide.slide_id,
            slide.binding.slide_reference_id.clone(),
        ) == *key
    }) else {
        return Ok(None);
    };
    let relationships_length = package
        .owner
        .package
        .declared_relationships_uncompressed_size(&slide.part_uri)?;
    Ok(Some((
        PlannedImage {
            xml: declared_bytes(package, SourceSpec::Part(slide.part_uri.clone()))?,
            relationships: relationships_length
                .map(|length| {
                    declared_bytes_with_length(
                        SourceSpec::Relationships(slide.part_uri.clone()),
                        length,
                    )
                })
                .transpose()?,
        },
        slide.binding.slide.content_type.clone(),
        slide.binding.presentation_relationship.kind.clone(),
    )))
}

fn planned_effective_image(
    package: &Package,
    key: &StableSlideKey,
) -> Result<Option<PlannedImage>> {
    if let Some(revision) = &package.topology {
        if let Some(change) = revision
            .data
            .changes
            .iter()
            .find(|change| &change.key == key)
        {
            return Ok(change.after.as_ref().map(planned_image));
        }
    }
    if !package.order.iter().any(|candidate| candidate == key) {
        return Ok(None);
    }
    let Some((physical, _, _)) = planned_physical_image(package, key)? else {
        return Ok(None);
    };
    let Some(physical_position) = package.physical.get(key).copied() else {
        return Ok(Some(physical));
    };
    if let Some(snapshot) = package
        .overlays
        .iter()
        .find(|snapshot| snapshot.position == physical_position)
    {
        return Ok(Some(PlannedImage {
            xml: planned_source_payload(&snapshot.xml),
            relationships: physical.relationships,
        }));
    }
    Ok(Some(physical))
}

fn planned_diff_changes(
    source: &Package,
    target: &Package,
    source_changes: Vec<PlannedChange>,
    target_changes: Vec<PlannedChange>,
) -> Result<Vec<PlannedChange>> {
    let mut source_map = HashMap::new();
    let mut target_map = HashMap::new();
    for change in source_changes {
        source_map.insert(change.key.clone(), change);
    }
    for change in target_changes {
        target_map.insert(change.key.clone(), change);
    }
    let mut keys = HashSet::new();
    keys.extend(source_map.keys().cloned());
    keys.extend(target_map.keys().cloned());
    let mut changes = Vec::new();
    for key in keys {
        let source_change = source_map.get(&key);
        let target_change = target_map.get(&key);
        let before = match source_change {
            Some(change) => change.after.clone(),
            None => planned_effective_image(source, &key)?,
        };
        let after = match target_change {
            Some(change) => change.after.clone(),
            None => planned_effective_image(target, &key)?,
        };
        let equal = match (&before, &after) {
            (None, None) => true,
            (Some(left), Some(right)) => left.same(right),
            _ => false,
        };
        if equal {
            continue;
        }
        let metadata = target_change
            .or(source_change)
            .ok_or(Error::PatchConflict)?;
        changes.push(PlannedChange {
            key,
            content_type: metadata.content_type.clone(),
            relationship_type: metadata.relationship_type.clone(),
            before,
            after,
        });
    }
    changes.sort_unstable_by(|left, right| {
        left.key
            .part_uri
            .as_str()
            .cmp(right.key.part_uri.as_str())
            .then(left.key.slide_id.cmp(&right.key.slide_id))
            .then(left.key.relationship_id.cmp(&right.key.relationship_id))
    });
    Ok(changes)
}

fn planned_current_dependents(package: &Package) -> Vec<PlannedDependent> {
    package
        .topology
        .as_ref()
        .map(|revision| {
            revision
                .data
                .dependents
                .iter()
                .map(|dependent| PlannedDependent {
                    part_uri: dependent.part_uri.clone(),
                    content_type: dependent.content_type.clone(),
                    before: dependent.before.as_ref().map(planned_image),
                    after: dependent.after.as_ref().map(planned_image),
                })
                .collect()
        })
        .unwrap_or_default()
}

fn planned_remove_change(package: &Package, key: &StableSlideKey) -> Result<PlannedChange> {
    let physical = planned_physical_image(package, key)?;
    let existing = planned_current_changes(package)?
        .into_iter()
        .find(|change| &change.key == key);
    let (before, content_type, relationship_type) = match physical {
        Some((image, content_type, relationship_type)) => {
            (Some(image), content_type, relationship_type)
        },
        None => {
            let existing = existing.ok_or(Error::PatchConflict)?;
            (
                existing.after.clone().or_else(|| existing.before.clone()),
                existing.content_type,
                existing.relationship_type,
            )
        },
    };
    Ok(PlannedChange {
        key: key.clone(),
        content_type,
        relationship_type,
        before,
        after: None,
    })
}

fn upsert_planned_change(changes: &mut Vec<PlannedChange>, change: PlannedChange) {
    changes.retain(|existing| existing.key != change.key);
    let equal = match (&change.before, &change.after) {
        (None, None) => true,
        (Some(left), Some(right)) => left.same(right),
        _ => false,
    };
    if !equal {
        changes.push(change);
    }
    changes.sort_unstable_by(|left, right| {
        left.key
            .part_uri
            .as_str()
            .cmp(right.key.part_uri.as_str())
            .then(left.key.slide_id.cmp(&right.key.slide_id))
            .then(left.key.relationship_id.cmp(&right.key.relationship_id))
    });
}

fn upsert_planned_dependent(dependents: &mut Vec<PlannedDependent>, dependent: PlannedDependent) {
    dependents.retain(|existing| !existing.part_uri.is_equivalent_to(&dependent.part_uri));
    let equal = match (&dependent.before, &dependent.after) {
        (None, None) => true,
        (Some(left), Some(right)) => left.same(right),
        _ => false,
    };
    if !equal {
        dependents.push(dependent);
    }
    dependents.sort_unstable_by_key(|dependent| dependent.part_uri.as_str().to_ascii_lowercase());
}

fn planned_physical_dependent(
    package: &Package,
    part_uri: &PackURI,
) -> Result<Option<(PlannedImage, String)>> {
    let view = match package.owner.package.part(part_uri) {
        Ok(view) => view,
        Err(litchi_opc::OpcError::PartNotFound(_)) => return Ok(None),
        Err(error) => return Err(Error::Opc(error)),
    };
    let relationships_length = package
        .owner
        .package
        .declared_relationships_uncompressed_size(part_uri)?;
    Ok(Some((
        PlannedImage {
            xml: declared_bytes(package, SourceSpec::Part(part_uri.clone()))?,
            relationships: relationships_length
                .map(|length| {
                    declared_bytes_with_length(SourceSpec::Relationships(part_uri.clone()), length)
                })
                .transpose()?,
        },
        view.content_type().to_owned(),
    )))
}

fn planned_effective_dependent(
    package: &Package,
    part_uri: &PackURI,
) -> Result<Option<PlannedImage>> {
    if let Some(revision) = &package.topology {
        if let Some(dependent) = revision
            .data
            .dependents
            .iter()
            .find(|dependent| dependent.part_uri.is_equivalent_to(part_uri))
        {
            return Ok(dependent.after.as_ref().map(planned_image));
        }
    }
    Ok(planned_physical_dependent(package, part_uri)?.map(|(image, _)| image))
}

fn planned_diff_dependents(source: &Package, target: &Package) -> Result<Vec<PlannedDependent>> {
    let source_entries = planned_current_dependents(source);
    let target_entries = planned_current_dependents(target);
    planned_diff_dependents_from_entries(source, source_entries, target_entries)
}

fn planned_diff_dependents_from_entries(
    package: &Package,
    source_entries: Vec<PlannedDependent>,
    target_entries: Vec<PlannedDependent>,
) -> Result<Vec<PlannedDependent>> {
    let mut entries = Vec::new();
    let mut seen = HashSet::new();
    for entry in source_entries.iter().chain(target_entries.iter()) {
        if !seen.insert(entry.part_uri.as_str().to_ascii_lowercase()) {
            continue;
        }
        let before = planned_effective_dependent(package, &entry.part_uri)?;
        let after = if let Some(dependent) = target_entries
            .iter()
            .find(|dependent| dependent.part_uri.is_equivalent_to(&entry.part_uri))
        {
            dependent.after.clone()
        } else {
            before.clone()
        };
        let equal = match (&before, &after) {
            (None, None) => true,
            (Some(left), Some(right)) => left.same(right),
            _ => false,
        };
        if equal {
            continue;
        }
        let metadata = target_entries
            .iter()
            .find(|candidate| candidate.part_uri.is_equivalent_to(&entry.part_uri))
            .or_else(|| {
                source_entries
                    .iter()
                    .find(|candidate| candidate.part_uri.is_equivalent_to(&entry.part_uri))
            })
            .ok_or(Error::PatchConflict)?;
        entries.push(PlannedDependent {
            part_uri: entry.part_uri.clone(),
            content_type: metadata.content_type.clone(),
            before,
            after,
        });
    }
    entries.sort_unstable_by_key(|dependent| dependent.part_uri.as_str().to_ascii_lowercase());
    Ok(entries)
}

fn materialize_bytes(
    package: &Package,
    planned: &PlannedBytes,
    retained: &mut Retention,
    metadata: bool,
) -> Result<Arc<Vec<u8>>> {
    match planned {
        PlannedBytes::Source(data) => {
            let copied = copy_state_bytes(package, data.as_bytes(), retained)?;
            if metadata {
                retained.source_metadata.push(data.clone());
            } else {
                retained.source_parts.push(data.clone());
            }
            Ok(copied)
        },
        PlannedBytes::Declared { spec, .. } => {
            let data = match spec {
                SourceSpec::Part(uri) => {
                    let data = package.owner.package.part(uri)?.data()?;
                    retained.source_parts.push(data.clone());
                    data
                },
                SourceSpec::Relationships(uri) => {
                    let data = package
                        .owner
                        .package
                        .relationships_data_for_with_limit(
                            uri,
                            package.owner.limits.max_relationship_xml_bytes(),
                        )?
                        .ok_or(Error::PatchConflict)?;
                    retained.source_metadata.push(data.clone());
                    data
                },
                SourceSpec::ContentTypes => {
                    let data = package.owner.package.content_types_data_with_limit(
                        package.owner.limits.max_content_types_bytes(),
                    )?;
                    retained.source_metadata.push(data.clone());
                    data
                },
            };
            copy_state_bytes(package, data.as_bytes(), retained)
        },
        PlannedBytes::Length(_) => Err(Error::PatchConflict),
        PlannedBytes::Owned { bytes, reservation } => {
            if let Some(reservation) = reservation {
                retained.reservations.push(Arc::clone(reservation));
            }
            Ok(Arc::clone(bytes))
        },
    }
}

fn materialize_state(
    package: &Package,
    planned: &PlannedState,
    retained: &mut Retention,
) -> Result<State> {
    Ok(State {
        root: materialize_bytes(package, &planned.root, retained, false)?,
        order: Arc::clone(&planned.order),
        topology_digest: planned.topology_digest,
        presentation_relationships: materialize_bytes(
            package,
            &planned.presentation_relationships,
            retained,
            true,
        )?,
        content_types: materialize_bytes(package, &planned.content_types, retained, true)?,
    })
}

fn materialize_image(
    package: &Package,
    planned: &PlannedImage,
    retained: &mut Retention,
) -> Result<PartImage> {
    Ok(PartImage {
        xml: materialize_bytes(package, &planned.xml, retained, false)?,
        relationships: planned
            .relationships
            .as_ref()
            .map(|relationships| materialize_bytes(package, relationships, retained, true))
            .transpose()?,
    })
}

fn materialize_change(
    package: &Package,
    planned: &PlannedChange,
    retained: &mut Retention,
) -> Result<Change> {
    Ok(Change {
        key: planned.key.clone(),
        content_type: planned.content_type.clone(),
        relationship_type: planned.relationship_type.clone(),
        before: planned
            .before
            .as_ref()
            .map(|image| materialize_image(package, image, retained))
            .transpose()?,
        after: planned
            .after
            .as_ref()
            .map(|image| materialize_image(package, image, retained))
            .transpose()?,
    })
}

fn materialize_dependent(
    package: &Package,
    planned: &PlannedDependent,
    retained: &mut Retention,
) -> Result<Dependent> {
    Ok(Dependent {
        part_uri: planned.part_uri.clone(),
        content_type: planned.content_type.clone(),
        before: planned
            .before
            .as_ref()
            .map(|image| materialize_image(package, image, retained))
            .transpose()?,
        after: planned
            .after
            .as_ref()
            .map(|image| materialize_image(package, image, retained))
            .transpose()?,
    })
}

/// Build a durable topology patch from the exact before/after package
/// revisions.  The before image is the caller's effective state; the physical
/// source image is retained in the recipe only where the target topology must
/// be rebuilt on another open.
pub(super) fn patch_from_packages(
    source: &Package,
    target: &Package,
    max_patch_bytes: usize,
) -> Result<Patch> {
    source.check()?;
    target.check()?;
    if !source.same_source_owner(target) {
        return Err(Error::PatchConflict);
    }
    let plan = plan_patch(source, target)?;
    let encoded = plan.encoded_len()?;
    if encoded > max_patch_bytes {
        return Err(limit(
            crate::error::CodecLimitKind::PatchBytes,
            encoded,
            max_patch_bytes,
        ));
    }
    let mut retained = source
        .topology
        .as_ref()
        .map(|revision| (*revision.retained).clone())
        .unwrap_or_default();
    if let Some(revision) = &target.topology {
        merge_retention(&mut retained, (*revision.retained).clone());
    }
    let before = materialize_state(source, &plan.before, &mut retained)?;
    let after = materialize_state(target, &plan.after, &mut retained)?;
    let changes = plan
        .changes
        .iter()
        .map(|change| materialize_change(source, change, &mut retained))
        .collect::<Result<Vec<_>>>()?;
    let dependents = plan
        .dependents
        .iter()
        .map(|dependent| materialize_dependent(source, dependent, &mut retained))
        .collect::<Result<Vec<_>>>()?;
    validate_root_delta(source, &before, &after)?;
    let data = Data {
        before,
        after,
        changes: changes.into(),
        dependents: dependents.into(),
    };
    if !data_matches_plan(&data, &plan) {
        return Err(Error::PatchConflict);
    }
    if encoded > max_patch_bytes {
        return Err(limit(
            crate::error::CodecLimitKind::PatchBytes,
            encoded,
            max_patch_bytes,
        ));
    }
    let prepared = target
        .topology
        .as_ref()
        .map(|revision| revision.prepared.clone());
    Ok(Patch {
        data,
        prepared,
        retained: Arc::new(retained),
        encoded_len_cache: Arc::new(OnceLock::from(encoded)),
    })
}

fn data_matches_plan(data: &Data, plan: &PatchPlan) -> bool {
    if data.before.root.len() != plan.before.root.lengths()
        || data.before.presentation_relationships.len()
            != plan.before.presentation_relationships.lengths()
        || data.before.content_types.len() != plan.before.content_types.lengths()
        || data.after.root.len() != plan.after.root.lengths()
        || data.after.presentation_relationships.len()
            != plan.after.presentation_relationships.lengths()
        || data.after.content_types.len() != plan.after.content_types.lengths()
        || data.before.order.as_ref() != plan.before.order.as_ref()
        || data.after.order.as_ref() != plan.after.order.as_ref()
        || data.changes.len() != plan.changes.len()
        || data.dependents.len() != plan.dependents.len()
    {
        return false;
    }
    let images_match = |actual: Option<&PartImage>, planned: Option<&PlannedImage>| {
        actual.zip(planned).is_some_and(|(actual, planned)| {
            actual.xml.len() == planned.xml.lengths()
                && actual.relationships.as_ref().map(|bytes| bytes.len())
                    == planned.relationships.as_ref().map(PlannedBytes::lengths)
        }) || actual.is_none() && planned.is_none()
    };
    data.changes
        .iter()
        .zip(&plan.changes)
        .all(|(actual, planned)| {
            actual.key == planned.key
                && actual.content_type == planned.content_type
                && actual.relationship_type == planned.relationship_type
                && images_match(actual.before.as_ref(), planned.before.as_ref())
                && images_match(actual.after.as_ref(), planned.after.as_ref())
        })
        && data
            .dependents
            .iter()
            .zip(&plan.dependents)
            .all(|(actual, planned)| {
                actual.part_uri == planned.part_uri
                    && actual.content_type == planned.content_type
                    && images_match(actual.before.as_ref(), planned.before.as_ref())
                    && images_match(actual.after.as_ref(), planned.after.as_ref())
            })
}

/// Apply a v3 structural patch to a source package.  The current effective
/// state is checked first, then the target is rebuilt as one typed topology;
/// no output ZIP or eager fallback owner is created.
pub(super) fn apply(package: &Package, data: &Data, shape_removals: bool) -> Result<Package> {
    package.check()?;
    let (current, _current_retention) = effective_state_with_retention(package)?;
    if !state_equal(&current, &data.before) {
        return Err(Error::PatchConflict);
    }
    validate_root_delta(package, &data.before, &data.after)?;
    let mut retained = Retention::default();
    // An inverse of an insertion, or a patch that replaces every physical
    // slide with an authored one, is the only patch-level path that may pass
    // through the source removal proof when the physical catalog has one
    // slide. Keep the public direct-removal final-slide guard intact: the
    // replacement case is admitted only when this same patch carries an
    // insertion that keeps its resulting order nonempty.
    let allows_empty_result = data.after.order.is_empty()
        || (!data.after.order.is_empty()
            && data
                .changes
                .iter()
                .any(|change| change.before.is_none() && change.after.is_some()));
    // Durable patch admission has complete before/after root and resource
    // witnesses, so it may preserve unrelated opaque relationship targets
    // during replay. Direct edit removal still passes `false, false` above.
    validate_changes_against_effective(
        package,
        &data.changes,
        &mut retained,
        allows_empty_result,
        true,
        shape_removals,
    )?;
    validate_dependents_against_effective(package, data, &mut retained)?;
    if data.before.topology_digest == data.after.topology_digest
        && data.before.root.as_slice() == data.after.root.as_slice()
        && data.before.order.as_ref() == data.after.order.as_ref()
        && data.before.presentation_relationships.as_slice()
            == data.after.presentation_relationships.as_slice()
        && data.before.content_types.as_slice() == data.after.content_types.as_slice()
        && data.changes.is_empty()
        && data.dependents.is_empty()
    {
        return Ok(package.clone());
    }

    // A patch authored from a structural candidate describes only its
    // incremental write set.  Rebase that set onto the immutable physical
    // source while retaining every earlier structural change, otherwise a
    // second removal or a restored Part would silently disappear on replay.
    let changes = rebase_changes_to_physical(package, &data.changes, &mut retained)?;
    let dependents = rebase_dependents_to_physical(package, &data.dependents, &mut retained)?;
    if target_is_physical(package, data, &changes, &dependents, &mut retained)? {
        return package_as_physical(package);
    }

    let revision = build_revision_with_retention(
        package,
        Arc::clone(&data.after.root),
        Arc::clone(&data.after.order),
        changes.into(),
        dependents.into(),
        Some(&data.after),
        retained,
    )?;
    validate_restored_dependents(&revision, data)?;
    package_from_revision(package, revision)
}

fn validate_root_delta(package: &Package, before: &State, after: &State) -> Result<()> {
    crate::opened::validate_slide_catalog_delta(
        before.root.as_slice(),
        before
            .order
            .iter()
            .map(|key| (key.slide_id, key.relationship_id.as_str())),
        after.root.as_slice(),
        after
            .order
            .iter()
            .map(|key| (key.slide_id, key.relationship_id.as_str())),
        package.owner.package.execution_context().as_ref(),
    )
}

fn effective_state_with_retention(package: &Package) -> Result<(State, Retention)> {
    if let Some(revision) = &package.topology {
        revision.check()?;
        return Ok((revision.data.after.clone(), (*revision.retained).clone()));
    }
    let mut retained = Retention::default();
    let (presentation_relationships, content_types) = source_metadata(package, &mut retained)?;
    retained
        .source_parts
        .push(package.owner._presentation.source_data().clone());
    let root = copy_state_bytes(package, package.root.as_bytes(), &mut retained)?;
    let order = Arc::clone(&package.order);
    let topology_digest = topology_digest::source_digest(&package.owner.package)?;
    Ok((
        State {
            topology_digest,
            root,
            order,
            presentation_relationships,
            content_types,
        },
        retained,
    ))
}

fn build_revision_with_retention(
    package: &Package,
    root: Arc<Vec<u8>>,
    order: Arc<[StableSlideKey]>,
    changes: Arc<[Change]>,
    dependents: Arc<[Dependent]>,
    target_state: Option<&State>,
    mut retained: Retention,
) -> Result<Revision> {
    package.owner.package.check_execution()?;
    let empty_target = target_state.is_some_and(|state| state.order.is_empty());
    if (order.is_empty() && !empty_target) || order.len() > crate::parts::MAX_SLIDES {
        return Err(Error::Invalid(
            "source-backed topology slide order is out of bounds".into(),
        ));
    }
    // A fresh admission owns all shared recipe images in this revision. It
    // avoids retaining entire preceding revisions just to keep their image
    // reservations alive, so repeated edits do not pin a chain of candidates.
    let amount = recipe_bytes(&root, &order, &changes, &dependents, target_state)?;
    if let Some(reservation) = reserve_payload(package, amount)? {
        retained.reservations.push(Arc::new(reservation));
    }
    let plan = build_plan(
        package,
        &root,
        &order,
        &changes,
        &dependents,
        target_state,
        &mut retained,
    )?;
    let prepared = package.owner.package.clone().prepare_topology(plan)?;
    package.owner.package.check_execution()?;
    validate_prepared_presence_transitions(package, &prepared, &changes)?;

    let (presentation_relationships, content_types) = if let Some(state) = target_state {
        let candidate_rels = prepared
            .relationships_data_for_with_limit(
                package.owner._presentation.source_partname(),
                package.owner.limits.max_relationship_xml_bytes(),
            )?
            .ok_or(Error::PatchConflict)?;
        let candidate_ct = prepared
            .content_types_data_with_limit(package.owner.limits.max_content_types_bytes())?;
        retained.prepared_metadata.push(candidate_rels.clone());
        retained.prepared_metadata.push(candidate_ct.clone());
        if candidate_rels.as_bytes() != state.presentation_relationships.as_slice()
            || candidate_ct.as_bytes() != state.content_types.as_slice()
        {
            return Err(Error::PatchConflict);
        }
        (
            Arc::clone(&state.presentation_relationships),
            Arc::clone(&state.content_types),
        )
    } else {
        let rels = prepared
            .relationships_data_for_with_limit(
                package.owner._presentation.source_partname(),
                package.owner.limits.max_relationship_xml_bytes(),
            )?
            .ok_or(Error::PatchConflict)?;
        let ct = prepared
            .content_types_data_with_limit(package.owner.limits.max_content_types_bytes())?;
        retained.prepared_metadata.push(rels.clone());
        retained.prepared_metadata.push(ct.clone());
        (
            copy_state_bytes(package, rels.as_bytes(), &mut retained)?,
            copy_state_bytes(package, ct.as_bytes(), &mut retained)?,
        )
    };
    let physical = physical_state(package, &mut retained)?;
    let topology_digest = topology_digest::prepared_digest(&prepared)?;
    if let Some(state) = target_state {
        if topology_digest != state.topology_digest {
            return Err(Error::PatchConflict);
        }
    }
    let after = State {
        topology_digest,
        root,
        order,
        presentation_relationships,
        content_types,
    };
    let data = Data {
        before: physical,
        after,
        changes,
        dependents,
    };
    Ok(Revision {
        prepared,
        data,
        retained: Arc::new(retained),
    })
}

fn build_plan(
    package: &Package,
    root: &[u8],
    order: &[StableSlideKey],
    changes: &[Change],
    dependents: &[Dependent],
    target_state: Option<&State>,
    retained: &mut Retention,
) -> Result<SourceTopologyPlan> {
    let owner = &package.owner;
    let mut plan = SourceTopologyPlan::new();
    let presentation_uri = owner._presentation.source_partname().clone();
    let physical_root = owner._presentation.source_data().as_bytes();
    if root != physical_root {
        stage_root_replacement(
            package,
            &mut plan,
            &presentation_uri,
            physical_root,
            root,
            retained,
        )?;
    }

    let mut seen = HashSet::new();
    for change in changes {
        if !seen.insert(change.key.clone()) {
            return Err(Error::PatchConflict);
        }
        let physical = match owner.package.part(&change.key.part_uri) {
            Ok(view) => Some((view.content_type().to_owned(), view)),
            Err(litchi_opc::OpcError::PartNotFound(_)) => None,
            Err(error) => return Err(Error::Opc(error)),
        };
        let desired = change.after.as_ref();
        match desired {
            None => {
                if let Some((_, view)) = &physical {
                    let physical_data = view.data()?;
                    retained.source_parts.push(physical_data);
                    if let Some(data) = owner.package.relationships_data_for_with_limit(
                        view.partname(),
                        owner.limits.max_relationship_xml_bytes(),
                    )? {
                        retained.source_metadata.push(data);
                    }
                    plan.try_remove_part(change.key.part_uri.clone())?;
                }
                if owner
                    ._presentation
                    .rels()
                    .get(&change.key.relationship_id)
                    .is_some()
                {
                    plan.try_remove_relationship(
                        presentation_uri.clone(),
                        change.key.relationship_id.clone(),
                    )?;
                }
            },
            Some(image) => {
                if let Some((content_type, view)) = physical {
                    if content_type != change.content_type {
                        return Err(Error::PatchConflict);
                    }
                    let physical_data = view.data()?;
                    let xml_changed = physical_data.as_bytes() != image.xml.as_slice();
                    if xml_changed {
                        stage_root_replacement(
                            package,
                            &mut plan,
                            view.partname(),
                            physical_data.as_bytes(),
                            &image.xml,
                            retained,
                        )?;
                    }
                    retained.source_parts.push(physical_data);
                    let physical_rels = owner.package.relationships_data_for_with_limit(
                        view.partname(),
                        owner.limits.max_relationship_xml_bytes(),
                    )?;
                    let physical_rels_arc = physical_rels
                        .as_ref()
                        .map(|data| copy_state_bytes(package, data.as_bytes(), retained))
                        .transpose()?;
                    if physical_rels_arc != image.relationships {
                        plan.try_restore_relationships_xml(
                            view.partname().clone(),
                            physical_rels_arc,
                            image.relationships.clone(),
                        )?;
                    }
                    if let Some(data) = physical_rels {
                        retained.source_metadata.push(data);
                    }
                } else {
                    // A slide absent from the physical source is either
                    // restored producer XML (a durable inverse after reopen)
                    // or a slide this library authored; both are held to the
                    // source XML policy, which authored compact XML passes.
                    plan.try_add_restored_part_shared(
                        change.key.part_uri.clone(),
                        change.content_type.clone(),
                        Arc::clone(&image.xml),
                    )?;
                    if let Some(rels) = &image.relationships {
                        super::topology_admission::stage_layout(
                            package,
                            &mut plan,
                            &change.key.part_uri,
                            rels,
                        )?;
                        plan.try_restore_relationships_xml(
                            change.key.part_uri.clone(),
                            None,
                            Some(Arc::clone(rels)),
                        )?;
                    }
                }
                let relation = owner._presentation.rels().get(&change.key.relationship_id);
                if let Some(relation) = relation {
                    let target_matches = relation
                        .target_partname()
                        .map(|target| target.is_equivalent_to(&change.key.part_uri))
                        .unwrap_or(false);
                    if relation.is_external()
                        || relation.reltype() != change.relationship_type
                        || !crate::parts::is_relationship_type(
                            relation.reltype(),
                            litchi_opc::constants::relationship_type::SLIDE,
                            "slide",
                        )
                        || !target_matches
                    {
                        plan.try_replace_internal_relationship(
                            presentation_uri.clone(),
                            change.key.relationship_id.clone(),
                            change.relationship_type.clone(),
                            change.key.part_uri.clone(),
                        )?;
                    }
                } else {
                    plan.try_add_internal_relationship(
                        presentation_uri.clone(),
                        change.key.relationship_id.clone(),
                        change.relationship_type.clone(),
                        change.key.part_uri.clone(),
                    )?;
                }
            },
        }
    }

    stage_dependents(package, &mut plan, dependents, retained)?;
    if let Some(state) = target_state {
        restore_metadata(package, &mut plan, state, retained)?;
    }
    // Ensure the root/order relationship graph is validated even when the
    // target contains no Part delta (for example an order composed after a
    // prior removal).
    crate::presentation::order::validate_source_order_root(
        &owner.package,
        owner._presentation.source_partname(),
        root,
        order,
        owner.package.execution_context().as_ref(),
    )?;
    Ok(plan)
}

/// Remove or restore the Parts that removed slides owned (record 0303).
/// A dependent is never edited: a restored image must equal a physical Part
/// that still exists, and a Part absent from the source is added with its
/// exact payload, typed edges and relationship image.
fn stage_dependents(
    package: &Package,
    plan: &mut SourceTopologyPlan,
    dependents: &[Dependent],
    retained: &mut Retention,
) -> Result<()> {
    let owner = &package.owner;
    for dependent in dependents {
        owner.package.check_execution()?;
        let physical = match owner.package.part(&dependent.part_uri) {
            Ok(view) => Some(view),
            Err(litchi_opc::OpcError::PartNotFound(_)) => None,
            Err(error) => return Err(Error::Opc(error)),
        };
        match (&dependent.after, physical) {
            (None, Some(view)) => {
                retained.source_parts.push(view.data()?);
                plan.try_remove_part(dependent.part_uri.clone())?;
            },
            (None, None) => {},
            (Some(image), Some(view)) => {
                let data = view.data()?;
                let relationships = owner.package.relationships_data_for_with_limit(
                    view.partname(),
                    owner.limits.max_relationship_xml_bytes(),
                )?;
                let same = view.content_type() == dependent.content_type
                    && data.as_bytes() == image.xml.as_slice()
                    && relationships.as_ref().map(PartData::as_bytes)
                        == image.relationships.as_ref().map(|bytes| bytes.as_slice());
                retained.source_parts.push(data);
                if let Some(relationships) = relationships {
                    retained.source_metadata.push(relationships);
                }
                if !same {
                    return Err(Error::PatchConflict);
                }
            },
            (Some(image), None) => {
                plan.try_add_restored_part_shared(
                    dependent.part_uri.clone(),
                    dependent.content_type.clone(),
                    Arc::clone(&image.xml),
                )?;
                if let Some(relationships) = &image.relationships {
                    super::topology_admission::stage_dependent_edges(
                        package,
                        plan,
                        &dependent.part_uri,
                        relationships,
                    )?;
                    plan.try_restore_relationships_xml(
                        dependent.part_uri.clone(),
                        None,
                        Some(Arc::clone(relationships)),
                    )?;
                }
            },
        }
    }
    Ok(())
}

/// The Parts the slides at `slides` own exclusively in the effective
/// topology of `package`: the prepared candidate of a structural revision,
/// or the physical source.
fn owned_parts(
    package: &Package,
    slides: &[PackURI],
) -> Result<crate::presentation::owned_parts::Owned> {
    use crate::presentation::owned_parts::{Graph, Owned, may_own};
    let maximum =
        crate::presentation::source_removal::relationship_scan_limit(package.owner.limits)?;
    if let Some(revision) = &package.topology {
        let mut any = false;
        for slide in slides {
            any |= may_own(revision.prepared.part(slide)?.rels());
        }
        if !any {
            return Ok(Owned::none());
        }
        return Graph::from_prepared(&revision.prepared, maximum)?.owned_by_all(slides);
    }
    let mut any = false;
    for slide in slides {
        any |= may_own(package.owner.package.part(slide)?.rels());
    }
    if !any {
        return Ok(Owned::none());
    }
    Graph::from_source(&package.owner.package, maximum)?.owned_by_all(slides)
}

/// The physical image and content type of a dependent Part, or `None` when
/// the source has no such Part.
fn physical_dependent(
    package: &Package,
    part_uri: &PackURI,
    retained: &mut Retention,
) -> Result<Option<(PartImage, String)>> {
    let view = match package.owner.package.part(part_uri) {
        Ok(view) => view,
        Err(litchi_opc::OpcError::PartNotFound(_)) => return Ok(None),
        Err(error) => return Err(Error::Opc(error)),
    };
    let data = view.data()?;
    let relationships = package.owner.package.relationships_data_for_with_limit(
        part_uri,
        package.owner.limits.max_relationship_xml_bytes(),
    )?;
    let relationships_arc = relationships
        .as_ref()
        .map(|data| copy_state_bytes(package, data.as_bytes(), retained))
        .transpose()?;
    let xml = copy_state_bytes(package, data.as_bytes(), retained)?;
    retained.source_parts.push(data);
    if let Some(data) = relationships {
        retained.source_metadata.push(data);
    }
    Ok(Some((
        PartImage {
            xml,
            relationships: relationships_arc,
        },
        view.content_type().to_owned(),
    )))
}

/// The effective image of a dependent Part: its recipe image in a
/// structural revision, else its physical image.
fn effective_dependent(
    package: &Package,
    part_uri: &PackURI,
    retained: &mut Retention,
) -> Result<Option<PartImage>> {
    if let Some(revision) = &package.topology {
        if let Some(dependent) = revision
            .data
            .dependents
            .iter()
            .find(|dependent| dependent.part_uri.is_equivalent_to(part_uri))
        {
            return Ok(dependent.after.clone());
        }
    }
    Ok(physical_dependent(package, part_uri, retained)?.map(|(image, _)| image))
}

fn current_dependents(package: &Package) -> Vec<Dependent> {
    package
        .topology
        .as_ref()
        .map(|revision| revision.data.dependents.to_vec())
        .unwrap_or_default()
}

fn upsert_dependent(dependents: &mut Vec<Dependent>, dependent: Dependent) {
    dependents.retain(|existing| !existing.part_uri.is_equivalent_to(&dependent.part_uri));
    if dependent.before != dependent.after {
        dependents.push(dependent);
    }
    sort_dependents(dependents);
}

fn sort_dependents(dependents: &mut [Dependent]) {
    dependents.sort_unstable_by_key(|dependent| dependent.part_uri.as_str().to_ascii_lowercase());
}

fn rebase_dependents_to_physical(
    package: &Package,
    desired: &[Dependent],
    retained: &mut Retention,
) -> Result<Vec<Dependent>> {
    let mut rebased = current_dependents(package);
    for dependent in desired {
        let before =
            physical_dependent(package, &dependent.part_uri, retained)?.map(|(image, _)| image);
        let mut recipe = dependent.clone();
        recipe.before = before;
        upsert_dependent(&mut rebased, recipe);
    }
    Ok(rebased)
}

/// A durable patch may remove exactly the Parts its removed slides own,
/// proven on the current effective topology, and each must hold the image
/// the patch was authored against.
fn validate_dependents_against_effective(
    package: &Package,
    data: &Data,
    retained: &mut Retention,
) -> Result<()> {
    for dependent in data.dependents.iter() {
        if effective_dependent(package, &dependent.part_uri, retained)? != dependent.before {
            return Err(Error::PatchConflict);
        }
    }
    let removed: Vec<PackURI> = data
        .changes
        .iter()
        .filter(|change| change.before.is_some() && change.after.is_none())
        .map(|change| change.key.part_uri.clone())
        .collect();
    let expected: Vec<&PackURI> = data
        .dependents
        .iter()
        .filter(|dependent| dependent.before.is_some())
        .map(|dependent| &dependent.part_uri)
        .collect();
    if removed.is_empty() {
        return if expected.is_empty() {
            Ok(())
        } else {
            Err(Error::PatchConflict)
        };
    }
    let owned = owned_parts(package, &removed)?;
    if !same_parts(owned.parts(), &expected) {
        return Err(Error::PatchConflict);
    }
    Ok(())
}

/// The Parts a patch restores must be exactly the Parts its restored slides
/// own in the prepared candidate.
fn validate_restored_dependents(revision: &Revision, data: &Data) -> Result<()> {
    let restored: Vec<PackURI> = data
        .changes
        .iter()
        .filter(|change| change.before.is_none() && change.after.is_some())
        .map(|change| change.key.part_uri.clone())
        .collect();
    let expected: Vec<&PackURI> = data
        .dependents
        .iter()
        .filter(|dependent| dependent.after.is_some())
        .map(|dependent| &dependent.part_uri)
        .collect();
    if expected.is_empty() {
        return Ok(());
    }
    if restored.is_empty() {
        return Err(Error::PatchConflict);
    }
    let maximum = revision
        .prepared
        .read_limits()
        .max_parts()
        .saturating_mul(16);
    let owned =
        crate::presentation::owned_parts::Graph::from_prepared(&revision.prepared, maximum)?
            .owned_by_all(&restored)?;
    if !same_parts(owned.parts(), &expected) {
        return Err(Error::PatchConflict);
    }
    Ok(())
}

fn same_parts(owned: &[PackURI], expected: &[&PackURI]) -> bool {
    owned.len() == expected.len()
        && owned.iter().all(|part| {
            expected
                .iter()
                .any(|candidate| candidate.is_equivalent_to(part))
        })
}

fn restore_metadata(
    package: &Package,
    plan: &mut SourceTopologyPlan,
    target: &State,
    retained: &mut Retention,
) -> Result<()> {
    let owner = package.owner._presentation.source_partname();
    let source_rels = package
        .owner
        .package
        .relationships_data_for_with_limit(
            owner,
            package.owner.limits.max_relationship_xml_bytes(),
        )?
        .ok_or(Error::PatchConflict)?;
    let source_rels_arc = copy_state_bytes(package, source_rels.as_bytes(), retained)?;
    retained.source_metadata.push(source_rels);
    if source_rels_arc.as_slice() != target.presentation_relationships.as_slice() {
        plan.try_restore_relationships_xml(
            owner.clone(),
            Some(source_rels_arc),
            Some(Arc::clone(&target.presentation_relationships)),
        )?;
    }
    let source_ct = package
        .owner
        .package
        .content_types_data_with_limit(package.owner.limits.max_content_types_bytes())?;
    let source_ct_arc = copy_state_bytes(package, source_ct.as_bytes(), retained)?;
    retained.source_metadata.push(source_ct);
    if source_ct_arc.as_slice() != target.content_types.as_slice() {
        plan.try_restore_content_types_xml(source_ct_arc, Arc::clone(&target.content_types))?;
    }
    Ok(())
}

fn validate_prepared_presence_transitions(
    package: &Package,
    prepared: &PreparedTopology,
    changes: &[Change],
) -> Result<()> {
    let root = prepared
        .part(package.owner._presentation.source_partname())?
        .data()?;
    for change in changes {
        if change.before.is_none() && change.after.is_some() {
            super::topology_admission::validate_insertion(
                package,
                prepared,
                root.as_bytes(),
                &change.key,
            )?;
        }
    }
    Ok(())
}

/// Admit a changed presentation root through the source XML authorization
/// path when the authored replacement retains formatting or other lexical
/// structure.  A bare replacement is appropriate only for the compact source
/// form accepted by the OPC XML audit; formatted source must carry its source
/// proof through the prepared topology so a reopen cannot silently rebind it.
fn stage_root_replacement(
    package: &Package,
    plan: &mut SourceTopologyPlan,
    presentation_uri: &PackURI,
    physical_root: &[u8],
    replacement: &[u8],
    retained: &mut Retention,
) -> Result<()> {
    if replacement == physical_root {
        return Ok(());
    }
    let source_part = package.owner.package.part(presentation_uri)?;
    let requires_proof = authored_xml_requires_source_proof(
        presentation_uri,
        source_part.content_type(),
        replacement,
    )?;
    if !requires_proof {
        plan.try_replace_part_shared(
            presentation_uri.clone(),
            copy_state_bytes(package, replacement, retained)?,
        )?;
        return Ok(());
    }
    let source_xml = source_part.source_xml()?;
    if source_xml.bytes() != physical_root {
        return Err(Error::PatchConflict);
    }
    let replacement =
        source_xml.with_replacement_xml(copy_state_bytes(package, replacement, retained)?)?;
    retained.source_xml.push(replacement.clone());
    plan.try_replace_source_xml_part(presentation_uri.clone(), replacement)?;
    Ok(())
}

fn package_from_revision(package: &Package, revision: Revision) -> Result<Package> {
    if revision.data.is_noop() {
        return package_as_physical(package);
    }
    let root = if let Some(context) = package.owner.package.execution_context() {
        let amount = u64::try_from(revision.data.after.root.len()).map_err(|_| {
            limit(
                crate::error::CodecLimitKind::SourceBytes,
                usize::MAX,
                usize::MAX,
            )
        })?;
        let reservation = context.reserve(Resource::Memory, amount).map_err(|error| {
            Error::Opc(match error {
                litchi_core::ExecutionError::Cancelled => litchi_opc::OpcError::Cancelled,
                other => litchi_opc::OpcError::Execution(other),
            })
        })?;
        RootPayload::ReservedEdit {
            bytes: Arc::clone(&revision.data.after.root),
            _reservation: Arc::new(reservation),
        }
    } else {
        RootPayload::Edited(Arc::clone(&revision.data.after.root))
    };
    let candidate = Package {
        owner: Arc::clone(&package.owner),
        source: Arc::clone(&package.source),
        root,
        order: Arc::clone(&revision.data.after.order),
        physical: Arc::clone(&package.physical),
        overlays: Arc::from([]),
        topology: Some(Arc::new(revision)),
        tolerance: Arc::clone(&package.tolerance),
        deferred: None,
    };
    candidate.check()?;
    Ok(candidate)
}

fn package_as_physical(package: &Package) -> Result<Package> {
    let owner = &package.owner;
    let order = super::source_order(owner)?.into();
    let physical = super::source_physical_index(owner)?.into();
    let candidate = Package {
        owner: Arc::clone(&package.owner),
        source: Arc::clone(&package.source),
        root: RootPayload::Original(owner._presentation.source_data().clone()),
        order,
        physical,
        overlays: Arc::from([]),
        topology: None,
        tolerance: Arc::clone(&package.tolerance),
        deferred: None,
    };
    candidate.check()?;
    Ok(candidate)
}

fn validate_changes_against_effective(
    package: &Package,
    changes: &[Change],
    retained: &mut Retention,
    allows_empty_result: bool,
    allows_unrelated_opaque_targets: bool,
    shape_removals: bool,
) -> Result<()> {
    let mut seen = HashSet::new();
    for change in changes {
        if !seen.insert(change.key.clone()) {
            return Err(Error::PatchConflict);
        }
        if change.content_type != litchi_opc::constants::content_type::PML_SLIDE
            || !crate::parts::is_relationship_type(
                &change.relationship_type,
                litchi_opc::constants::relationship_type::SLIDE,
                "slide",
            )
        {
            return Err(Error::PatchConflict);
        }
        let current = effective_image(package, &change.key, retained)?;
        match (&current, &change.before) {
            (None, None) => {},
            (Some(current), Some(before)) if current == before => {},
            _ => return Err(Error::PatchConflict),
        }
        match (&change.before, &change.after) {
            (Some(before), Some(after)) => {
                if before.relationships != after.relationships {
                    return Err(Error::PatchConflict);
                }
                let semantic =
                    crate::open_options::selected(package.owner.semantic_limits.as_ref());
                super::durable::validate_pair_with_limits(
                    before.xml.as_slice(),
                    after.xml.as_slice(),
                    Some(package.owner.limits),
                    package.owner.package.execution_context().as_ref(),
                    shape_removals,
                    semantic.shape(),
                    semantic.transition(),
                )?;
            },
            (None, Some(after)) => {
                crate::source_limits::validate_xml_limits(
                    after.xml.as_slice(),
                    package.owner.limits,
                )?;
                // A restored slide is an exact image; markup compatibility
                // enclosed in its shapes is admitted (record 0303).
                let semantic =
                    crate::open_options::selected(package.owner.semantic_limits.as_ref());
                crate::presentation::transition::admit_source_with_limits(
                    after.xml.as_slice(),
                    "apply",
                    semantic.shape(),
                    semantic.transition(),
                )?;
            },
            (Some(_), None) => validate_removal_admission(
                package,
                &change.key,
                allows_empty_result,
                allows_unrelated_opaque_targets,
            )?,
            _ => {},
        }
    }
    Ok(())
}

fn validate_removal_admission(
    package: &Package,
    key: &StableSlideKey,
    allows_empty_result: bool,
    allows_unrelated_opaque_targets: bool,
) -> Result<()> {
    if let Some(revision) = &package.topology {
        let admission = if allows_unrelated_opaque_targets {
            super::topology_admission::validate_insertion
        } else {
            super::topology_admission::validate
        };
        return admission(package, &revision.prepared, package.root.as_bytes(), key);
    }
    let physical = package
        .physical
        .get(key)
        .copied()
        .ok_or(Error::PatchConflict)?;
    if allows_empty_result {
        package
            .owner
            .plan_slide_removal_for_inverse(Position::new(physical))?;
    } else if allows_unrelated_opaque_targets {
        package
            .owner
            .plan_slide_removal_for_patch(Position::new(physical))?;
    } else {
        package.owner.plan_slide_removal(Position::new(physical))?;
    }
    super::topology_admission::validate_physical_roundtrip(package, key)
}

fn rebase_changes_to_physical(
    package: &Package,
    desired: &[Change],
    retained: &mut Retention,
) -> Result<Vec<Change>> {
    let mut rebased = current_changes(package, retained)?;
    for change in desired {
        let before = physical_image(package, &change.key, retained)?.map(|image| image.0);
        let mut recipe = change.clone();
        recipe.before = before;
        upsert_change(&mut rebased, recipe);
    }
    Ok(rebased)
}

fn target_is_physical(
    package: &Package,
    data: &Data,
    changes: &[Change],
    dependents: &[Dependent],
    retained: &mut Retention,
) -> Result<bool> {
    let physical = physical_state(package, retained)?;
    if data.after.topology_digest != physical.topology_digest
        || data.after.root.as_slice() != physical.root.as_slice()
        || data.after.order.as_ref() != physical.order.as_ref()
        || data.after.presentation_relationships.as_slice()
            != physical.presentation_relationships.as_slice()
        || data.after.content_types.as_slice() != physical.content_types.as_slice()
    {
        return Ok(false);
    }
    for change in changes {
        let physical_image = physical_image(package, &change.key, retained)?.map(|value| value.0);
        if physical_image != change.after {
            return Ok(false);
        }
    }
    for dependent in dependents {
        let physical =
            physical_dependent(package, &dependent.part_uri, retained)?.map(|value| value.0);
        if physical != dependent.after {
            return Ok(false);
        }
    }
    Ok(true)
}

fn physical_state(package: &Package, retained: &mut Retention) -> Result<State> {
    retained
        .source_parts
        .push(package.owner._presentation.source_data().clone());
    let root = copy_state_bytes(
        package,
        package.owner._presentation.source_data().as_bytes(),
        retained,
    )?;
    let order: Arc<[StableSlideKey]> = super::source_order(&package.owner)?.into();
    let (presentation_relationships, content_types) = source_metadata(package, retained)?;
    Ok(State {
        topology_digest: topology_digest::source_digest(&package.owner.package)?,
        root,
        order,
        presentation_relationships,
        content_types,
    })
}

fn source_metadata(
    package: &Package,
    retained: &mut Retention,
) -> Result<(Arc<Vec<u8>>, Arc<Vec<u8>>)> {
    let root = package.owner._presentation.source_partname();
    let rels = package
        .owner
        .package
        .relationships_data_for_with_limit(root, package.owner.limits.max_relationship_xml_bytes())?
        .ok_or(Error::PatchConflict)?;
    let ct = package
        .owner
        .package
        .content_types_data_with_limit(package.owner.limits.max_content_types_bytes())?;
    let rels_arc = copy_state_bytes(package, rels.as_bytes(), retained)?;
    let ct_arc = copy_state_bytes(package, ct.as_bytes(), retained)?;
    retained.source_metadata.push(rels);
    retained.source_metadata.push(ct);
    Ok((rels_arc, ct_arc))
}

fn current_changes(package: &Package, retained: &mut Retention) -> Result<Vec<Change>> {
    if let Some(revision) = &package.topology {
        return Ok(revision.data.changes.to_vec());
    }
    let mut changes = Vec::new();
    for snapshot in package.overlays.iter() {
        let key = super::source_key_for_position(package, snapshot.position)?;
        let image = physical_image(package, &key, retained)?.ok_or(Error::PatchConflict)?;
        let (before_image, content_type, relationship_type) = image;
        let after_xml = match &snapshot.xml {
            SourcePayload::Original(data) => copy_state_bytes(package, data.as_bytes(), retained)?,
            SourcePayload::Edited(bytes) => Arc::clone(bytes),
            SourcePayload::ReservedEdit { bytes, reservation } => {
                retained.reservations.push(Arc::clone(reservation));
                Arc::clone(bytes)
            },
        };
        changes.push(Change {
            key,
            content_type,
            relationship_type,
            before: Some(before_image.clone()),
            after: Some(PartImage {
                xml: after_xml,
                relationships: before_image.relationships.clone(),
            }),
        });
    }
    Ok(changes)
}

fn effective_image(
    package: &Package,
    key: &StableSlideKey,
    retained: &mut Retention,
) -> Result<Option<PartImage>> {
    if let Some(revision) = &package.topology {
        if let Some(change) = revision
            .data
            .changes
            .iter()
            .find(|change| &change.key == key)
        {
            return Ok(change.after.clone());
        }
    }
    if !package.order.iter().any(|candidate| candidate == key) {
        return Ok(None);
    }
    let physical = physical_image(package, key, retained)?;
    let Some((image, _, _)) = physical else {
        return Ok(None);
    };
    let Some(physical_position) = package.physical.get(key).copied() else {
        return Ok(Some(image));
    };
    if let Some(snapshot) = package
        .overlays
        .iter()
        .find(|snapshot| snapshot.position == physical_position)
    {
        return Ok(Some(PartImage {
            xml: copy_state_bytes(package, snapshot.xml.as_bytes(), retained)?,
            relationships: image.relationships,
        }));
    }
    Ok(Some(image))
}

fn physical_image(
    package: &Package,
    key: &StableSlideKey,
    retained: &mut Retention,
) -> Result<Option<(PartImage, String, String)>> {
    let Some(slide) = package.owner.slides.iter().find(|slide| {
        stable_slide_key(
            slide.part_uri.clone(),
            slide.slide_id,
            slide.binding.slide_reference_id.clone(),
        ) == *key
    }) else {
        return Ok(None);
    };
    let view = package.owner.package.part(&slide.part_uri)?;
    let data = view.data()?;
    let relationships = package.owner.package.relationships_data_for_with_limit(
        &slide.part_uri,
        package.owner.limits.max_relationship_xml_bytes(),
    )?;
    let relationships_arc = relationships
        .as_ref()
        .map(|data| copy_state_bytes(package, data.as_bytes(), retained))
        .transpose()?;
    retained.source_parts.push(data.clone());
    if let Some(data) = relationships.clone() {
        retained.source_metadata.push(data);
    }
    Ok(Some((
        PartImage {
            xml: copy_state_bytes(package, data.as_bytes(), retained)?,
            relationships: relationships_arc,
        },
        slide.binding.slide.content_type.clone(),
        slide.binding.presentation_relationship.kind.clone(),
    )))
}

fn image_lengths(image: &PartImage) -> ImageLengths {
    ImageLengths {
        xml: image.xml.len(),
        relationships: image.relationships.as_ref().map(|bytes| bytes.len()),
    }
}

fn upsert_change(changes: &mut Vec<Change>, change: Change) {
    if change.before == change.after {
        changes.retain(|existing| existing.key != change.key);
        return;
    }
    if let Some(existing) = changes
        .iter_mut()
        .find(|candidate| candidate.key == change.key)
    {
        *existing = change;
    } else {
        changes.push(change);
    }
    changes.sort_unstable_by(|left, right| {
        left.key
            .part_uri
            .as_str()
            .cmp(right.key.part_uri.as_str())
            .then(left.key.slide_id.cmp(&right.key.slide_id))
            .then(left.key.relationship_id.cmp(&right.key.relationship_id))
    });
}

fn effective_slide_binding(
    package: &Package,
    revision: &Revision,
    key: &StableSlideKey,
    part: &litchi_opc::PreparedPartView,
) -> Result<SlideBinding> {
    if let Some(slide) = package.owner.slides.iter().find(|slide| {
        stable_slide_key(
            slide.part_uri.clone(),
            slide.slide_id,
            slide.binding.slide_reference_id.clone(),
        ) == *key
    }) {
        let mut binding = (*slide.binding).clone();
        binding.slide.uri = part.partname().clone();
        binding.slide.content_type = part.content_type().to_owned();
        binding.slide.relationships = relationships_to_bindings(part.rels())?;
        return Ok(binding);
    }
    let change = revision
        .data
        .changes
        .iter()
        .find(|change| &change.key == key)
        .ok_or(Error::PatchConflict)?;
    let presentation_uri = package.owner._presentation.source_partname();
    Ok(SlideBinding {
        slide_reference_id: key.relationship_id.clone(),
        presentation_relationship: RelationshipBinding {
            id: key.relationship_id.clone(),
            kind: change.relationship_type.clone(),
            target: key.part_uri.relative_ref(presentation_uri.base_uri()),
            mode: TargetMode::Internal,
        },
        slide: PartBinding {
            uri: part.partname().clone(),
            content_type: part.content_type().to_owned(),
            relationships: relationships_to_bindings(part.rels())?,
        },
    })
}

fn relationships_to_bindings(rels: &Relationships) -> Result<Box<[RelationshipBinding]>> {
    let mut bindings = Vec::new();
    bindings
        .try_reserve_exact(rels.iter().count())
        .map_err(allocation)?;
    for relationship in rels.iter() {
        bindings.push(RelationshipBinding {
            id: relationship.r_id().to_owned(),
            kind: relationship.reltype().to_owned(),
            target: relationship.target_ref().to_owned(),
            mode: relationship.target_mode(),
        });
    }
    Ok(bindings.into())
}

fn state_equal(left: &State, right: &State) -> bool {
    left.root.as_slice() == right.root.as_slice()
        && left.order.as_ref() == right.order.as_ref()
        && left.presentation_relationships.as_slice() == right.presentation_relationships.as_slice()
        && left.content_types.as_slice() == right.content_types.as_slice()
        && left.topology_digest == right.topology_digest
}

fn recipe_bytes(
    root: &[u8],
    order: &[StableSlideKey],
    changes: &[Change],
    dependents: &[Dependent],
    state: Option<&State>,
) -> Result<usize> {
    const MAXIMUM: usize = 128 * 1024 * 1024;
    if changes.len() > super::super::MAX_SOURCE_BACKED_SLIDE_BATCH {
        return Err(limit(
            crate::error::CodecLimitKind::Slides,
            changes.len(),
            super::super::MAX_SOURCE_BACKED_SLIDE_BATCH,
        ));
    }
    let mut total = root.len();
    let mut add = |bytes: usize| -> Result<()> {
        let next = total.checked_add(bytes).ok_or_else(|| {
            limit(
                crate::error::CodecLimitKind::ProofBytes,
                usize::MAX,
                MAXIMUM,
            )
        })?;
        if next > MAXIMUM {
            return Err(limit(
                crate::error::CodecLimitKind::ProofBytes,
                next,
                MAXIMUM,
            ));
        }
        total = next;
        Ok(())
    };
    for key in order {
        add(192)?;
        add(key.part_uri.as_str().len())?;
        add(key.relationship_id.len())?;
    }
    for change in changes {
        add(256)?;
        add(change.key.part_uri.as_str().len())?;
        add(change.key.relationship_id.len())?;
        add(change.content_type.len())?;
        add(change.relationship_type.len())?;
        for image in [change.before.as_ref(), change.after.as_ref()]
            .into_iter()
            .flatten()
        {
            add(image.xml.len())?;
            if let Some(rels) = &image.relationships {
                add(rels.len())?;
            }
        }
    }
    for dependent in dependents {
        add(256)?;
        add(dependent.part_uri.as_str().len())?;
        add(dependent.content_type.len())?;
        for image in [dependent.before.as_ref(), dependent.after.as_ref()]
            .into_iter()
            .flatten()
        {
            add(image.xml.len())?;
            if let Some(rels) = &image.relationships {
                add(rels.len())?;
            }
        }
    }
    if let Some(state) = state {
        add(state.presentation_relationships.len())?;
        add(state.content_types.len())?;
    }
    Ok(total)
}

pub(super) fn reserve_payload(package: &Package, bytes: usize) -> Result<Option<Reservation>> {
    let Some(context) = package.owner.package.execution_context() else {
        return Ok(None);
    };
    let amount = u64::try_from(bytes).map_err(|_| {
        limit(
            crate::error::CodecLimitKind::SourceBytes,
            usize::MAX,
            usize::MAX,
        )
    })?;
    context
        .reserve(Resource::Memory, amount)
        .map(Some)
        .map_err(|error| {
            Error::Opc(match error {
                litchi_core::ExecutionError::Cancelled => litchi_opc::OpcError::Cancelled,
                other => litchi_opc::OpcError::Execution(other),
            })
        })
}

fn copy_bytes(bytes: &[u8]) -> Result<Arc<Vec<u8>>> {
    let mut owned = Vec::new();
    owned.try_reserve_exact(bytes.len()).map_err(allocation)?;
    owned.extend_from_slice(bytes);
    Ok(Arc::new(owned))
}

fn copy_state_bytes(
    package: &Package,
    bytes: &[u8],
    retained: &mut Retention,
) -> Result<Arc<Vec<u8>>> {
    if let Some(reservation) = reserve_payload(package, bytes.len())? {
        retained.reservations.push(Arc::new(reservation));
    }
    copy_bytes(bytes)
}

fn merge_retention(destination: &mut Retention, mut source: Retention) {
    destination.source_parts.append(&mut source.source_parts);
    destination
        .prepared_parts
        .append(&mut source.prepared_parts);
    destination
        .source_metadata
        .append(&mut source.source_metadata);
    destination
        .prepared_metadata
        .append(&mut source.prepared_metadata);
    destination.reservations.append(&mut source.reservations);
    destination.source_xml.append(&mut source.source_xml);
}
