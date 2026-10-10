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
use std::sync::Arc;

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
use super::topology_wire::{Change, Data, PartImage, State};
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
}

impl Patch {
    pub(super) fn from_data(data: Data) -> Self {
        Self {
            data,
            prepared: None,
            retained: Arc::new(Retention::default()),
        }
    }

    pub(super) fn inverse(&self) -> Self {
        Self {
            data: self.data.inverse(),
            prepared: self.prepared.clone(),
            retained: Arc::clone(&self.retained),
        }
    }

    pub(super) fn is_noop(&self) -> bool {
        self.data.is_noop()
    }

    pub(super) fn encoded_len(&self) -> Result<usize> {
        if let Some(prepared) = &self.prepared {
            prepared.check_source()?;
            prepared.check_execution()?;
        }
        let length = self.data.encoded_len()?;
        if let Some(prepared) = &self.prepared {
            prepared.check_source()?;
            prepared.check_execution()?;
        }
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
            return Err(Error::StaleSource);
        }
        let data = part.data()?;
        let bytes = data.as_bytes();
        let limits = package.owner.limits;
        crate::source_limits::validate_xml_limits(bytes, limits)?;
        if bytes.len() > limits.max_part_bytes() as usize {
            return Err(limit(
                "source-backed primary prepared slide bytes",
                limits.max_part_bytes() as usize,
            ));
        }

        // Admission must precede the owned copy.  The source PartData remains
        // attached to the prepared candidate; this extra reservation covers
        // the edit-owned payload that is handed to the public snapshot.
        let mut amount = bytes
            .len()
            .checked_add(1024)
            .ok_or_else(|| limit("source-backed snapshot bytes", usize::MAX))?;
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
                .ok_or_else(|| limit("source-backed snapshot bytes", usize::MAX))?;
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
            max_output_bytes: super::super::source_slide_output_limit(limits),
            source_version,
            lineage: package.owner.package.source_lineage(),
            context,
        })
    }
}

/// Remove one selected stable slide after the standalone source planner has
/// admitted the physical semantic surface.  This function constructs the
/// final target from the currently staged root/order and rebuilds one prepared
/// topology, so previous order and text edits remain part of the same atomic
/// candidate.
pub(super) fn remove(package: &Package, key: &StableSlideKey) -> Result<Package> {
    package.check()?;
    let logical = package
        .order
        .iter()
        .position(|candidate| candidate == key)
        .ok_or(Error::StaleSource)?;
    if package.order.len() < 2 {
        return Err(Error::SlideRemovalPlan {
            kind: crate::SlideRemovalRefusal::FinalSlide,
            detail: "a presentation must retain at least one slide".into(),
        });
    }
    validate_removal_admission(package, key, false, false)?;

    let _root_reservation = reserve_payload(package, package.root.as_bytes().len())?;
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
                .ok_or(Error::StaleSource)?;
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
    let revision = build_revision_with_retention(
        package,
        Arc::new(root),
        order.into(),
        changes.into(),
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
pub(super) fn insert(package: &Package, position: Position, layout_name: &str) -> Result<Package> {
    package.check()?;
    let (root, order, change, insertion_reservation) =
        super::insertion::plan(package, position, layout_name)?;
    let mut retained = Retention::default();
    if let Some(reservation) = insertion_reservation {
        retained.reservations.push(reservation);
    }
    let mut changes = current_changes(package, &mut retained)?;
    upsert_change(&mut changes, change);
    let revision =
        build_revision_with_retention(package, root, order, changes.into(), None, retained)?;
    package_from_revision(package, revision)
}

/// Rebuild a structural candidate after a logical order change.
pub(super) fn reorder(
    package: &Package,
    root: Arc<Vec<u8>>,
    order: Arc<[StableSlideKey]>,
) -> Result<Package> {
    let revision = package.topology.as_ref().ok_or(Error::StaleSource)?;
    let revision = build_revision(
        package,
        root,
        order,
        Arc::clone(&revision.data.changes),
        Some(&revision.data.after),
    )?;
    package_from_revision(package, revision)
}

/// Rebuild a structural candidate after an existing effective slide XML was
/// changed by the ordinary shape-text editor.
pub(super) fn replace_snapshot(
    package: &Package,
    snapshot: SourceBackedSlideSnapshot,
) -> Result<Package> {
    let revision = package.topology.as_ref().ok_or(Error::StaleSource)?;
    let key = package
        .order
        .iter()
        .find(|candidate| candidate.part_uri == snapshot.part_uri)
        .cloned()
        .ok_or(Error::StaleSource)?;
    let mut changes = revision.data.changes.to_vec();
    let existing = changes.iter().find(|change| change.key == key).cloned();
    let mut seed = Retention::default();
    let relationships = match existing.as_ref().and_then(|change| change.after.as_ref()) {
        Some(image) => image.relationships.clone(),
        None => {
            physical_image(package, &key, &mut seed)?
                .ok_or(Error::StaleSource)?
                .0
                .relationships
        },
    };
    let (content_type, relationship_type, before) = if let Some(change) = existing {
        (change.content_type, change.relationship_type, change.before)
    } else {
        let image = physical_image(package, &key, &mut seed)?.ok_or(Error::StaleSource)?;
        (image.1, image.2, Some(image.0))
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
        Some(&revision.data.after),
        seed,
    )?;
    package_from_revision(package, revision)
}

/// Build a durable topology patch from the exact before/after package
/// revisions.  The before image is the caller's effective state; the physical
/// source image is retained in the recipe only where the target topology must
/// be rebuilt on another open.
pub(super) fn patch_from_packages(source: &Package, target: &Package) -> Result<Patch> {
    source.check()?;
    target.check()?;
    if !source.same_source_owner(target) {
        return Err(Error::StaleSource);
    }
    let (before, mut retained) = effective_state_with_retention(source)?;
    let (after, target_retained) = effective_state_with_retention(target)?;
    merge_retention(&mut retained, target_retained);
    validate_root_delta(source, &before, &after)?;
    let changes = diff_changes(source, target, &mut retained)?;
    let data = Data {
        before,
        after,
        changes: changes.into(),
    };
    let prepared = target
        .topology
        .as_ref()
        .map(|revision| revision.prepared.clone());
    Ok(Patch {
        data,
        prepared,
        retained: Arc::new(retained),
    })
}

/// Apply a v3 structural patch to a source package.  The current effective
/// state is checked first, then the target is rebuilt as one typed topology;
/// no output ZIP or eager fallback owner is created.
pub(super) fn apply(package: &Package, data: &Data, shape_removals: bool) -> Result<Package> {
    package.check()?;
    let (current, _current_retention) = effective_state_with_retention(package)?;
    if !state_equal(&current, &data.before) {
        return Err(Error::StaleSource);
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
    if data.before.topology_digest == data.after.topology_digest
        && data.before.root.as_slice() == data.after.root.as_slice()
        && data.before.order.as_ref() == data.after.order.as_ref()
        && data.before.presentation_relationships.as_slice()
            == data.after.presentation_relationships.as_slice()
        && data.before.content_types.as_slice() == data.after.content_types.as_slice()
        && data.changes.is_empty()
    {
        return Ok(package.clone());
    }

    // A patch authored from a structural candidate describes only its
    // incremental write set.  Rebase that set onto the immutable physical
    // source while retaining every earlier structural change, otherwise a
    // second removal or a restored Part would silently disappear on replay.
    let changes = rebase_changes_to_physical(package, &data.changes, &mut retained)?;
    if target_is_physical(package, data, &changes, &mut retained)? {
        return package_as_physical(package);
    }

    let revision = build_revision_with_retention(
        package,
        Arc::clone(&data.after.root),
        Arc::clone(&data.after.order),
        changes.into(),
        Some(&data.after),
        retained,
    )?;
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

fn build_revision(
    package: &Package,
    root: Arc<Vec<u8>>,
    order: Arc<[StableSlideKey]>,
    changes: Arc<[Change]>,
    target_state: Option<&State>,
) -> Result<Revision> {
    build_revision_with_retention(
        package,
        root,
        order,
        changes,
        target_state,
        Retention::default(),
    )
}

fn build_revision_with_retention(
    package: &Package,
    root: Arc<Vec<u8>>,
    order: Arc<[StableSlideKey]>,
    changes: Arc<[Change]>,
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
    let amount = recipe_bytes(&root, &order, &changes, target_state)?;
    if let Some(reservation) = reserve_payload(package, amount)? {
        retained.reservations.push(Arc::new(reservation));
    }
    let plan = build_plan(
        package,
        &root,
        &order,
        &changes,
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
            .ok_or(Error::StaleSource)?;
        let candidate_ct = prepared
            .content_types_data_with_limit(package.owner.limits.max_content_types_bytes())?;
        retained.prepared_metadata.push(candidate_rels.clone());
        retained.prepared_metadata.push(candidate_ct.clone());
        if candidate_rels.as_bytes() != state.presentation_relationships.as_slice()
            || candidate_ct.as_bytes() != state.content_types.as_slice()
        {
            return Err(Error::StaleSource);
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
            .ok_or(Error::StaleSource)?;
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
            return Err(Error::StaleSource);
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
            return Err(Error::StaleSource);
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
                        return Err(Error::StaleSource);
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
                    plan.try_add_part_shared(
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
        .ok_or(Error::StaleSource)?;
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
        return Err(Error::StaleSource);
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
        let amount = u64::try_from(revision.data.after.root.len())
            .map_err(|_| limit("source-backed primary topology root bytes", usize::MAX))?;
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
            return Err(Error::StaleSource);
        }
        if change.content_type != litchi_opc::constants::content_type::PML_SLIDE
            || !crate::parts::is_relationship_type(
                &change.relationship_type,
                litchi_opc::constants::relationship_type::SLIDE,
                "slide",
            )
        {
            return Err(Error::StaleSource);
        }
        let current = effective_image(package, &change.key, retained)?;
        match (&current, &change.before) {
            (None, None) => {},
            (Some(current), Some(before)) if current == before => {},
            _ => return Err(Error::StaleSource),
        }
        match (&change.before, &change.after) {
            (Some(before), Some(after)) => {
                if before.relationships != after.relationships {
                    return Err(Error::StaleSource);
                }
                super::durable::validate_pair(
                    before.xml.as_slice(),
                    after.xml.as_slice(),
                    Some(package.owner.limits),
                    package.owner.package.execution_context().as_ref(),
                    shape_removals,
                )?;
            },
            (None, Some(after)) => {
                crate::source_limits::validate_xml_limits(
                    after.xml.as_slice(),
                    package.owner.limits,
                )?;
                crate::presentation::transition::validate_source(after.xml.as_slice(), "apply")?;
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
        .ok_or(Error::StaleSource)?;
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
        .ok_or(Error::StaleSource)?;
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
        let image = physical_image(package, &key, retained)?.ok_or(Error::StaleSource)?;
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

fn diff_changes(
    source: &Package,
    target: &Package,
    retained: &mut Retention,
) -> Result<Vec<Change>> {
    let source_changes = current_changes(source, retained)?;
    let target_changes = current_changes(target, retained)?;
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
        let before = source_change
            .map(|change| change.after.clone())
            .unwrap_or(physical_image(source, &key, retained)?.map(|image| image.0));
        let after = target_change
            .map(|change| change.after.clone())
            .unwrap_or(physical_image(target, &key, retained)?.map(|image| image.0));
        if before == after {
            continue;
        }
        let metadata = target_change.or(source_change).ok_or(Error::StaleSource)?;
        changes.push(Change {
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
        .ok_or(Error::StaleSource)?;
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
    state: Option<&State>,
) -> Result<usize> {
    const MAXIMUM: usize = 128 * 1024 * 1024;
    if changes.len() > super::super::MAX_SOURCE_BACKED_SLIDE_BATCH {
        return Err(limit(
            "source-backed topology resource changes",
            super::super::MAX_SOURCE_BACKED_SLIDE_BATCH,
        ));
    }
    let mut total = root.len();
    let mut add = |bytes: usize| -> Result<()> {
        total = total
            .checked_add(bytes)
            .filter(|n| *n <= MAXIMUM)
            .ok_or_else(|| limit("source-backed topology recipe bytes", MAXIMUM))?;
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
    let amount = u64::try_from(bytes)
        .map_err(|_| limit("source-backed primary prepared slide bytes", usize::MAX))?;
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
