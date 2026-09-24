//! Bounded OPC graph binding for source-checked Ink host removal.
//!
//! The public Ink facade edits story XML separately.  This module owns the
//! smaller OPC closure: exact relationship and content-type tokens, selected
//! resource payload guards, and the reversible part additions/removals needed
//! to publish an inverse without normalising unrelated package members.

use std::collections::HashSet;
use std::{mem::size_of, sync::Arc};

use litchi_core::patch::BlobId;
use litchi_drawingml::ink as shared;
use litchi_opc::constants::relationship_type as rt;
use litchi_opc::{
    BlobPart, OpcError, OpcPackage, OwnedContentTypes, OwnedRelationships, OwnedXmlPart, PackURI,
    ReadLimits, Relationships, TargetMode,
};

use super::{CONTENT_TYPE, Limits};
use crate::{Error, Result};

mod add;
#[allow(
    unused_imports,
    reason = "the graph API returns this crate-private identity to transaction hosts"
)]
pub(crate) use add::{AddDelta, AddedTarget, NewTarget};

const STRICT_CUSTOM_XML: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/customXml";
const MAX_RELATIONSHIP_XML_BYTES: usize = 8 * 1024 * 1024;

/// Package graph and exact publication-token binding for stale candidates.
///
/// Names, content types and relationship fields retain semantic ownership.
/// Content addresses additionally bind lexical XML and relationship-member
/// presence, including owners untouched by an edit. Part payloads are guarded
/// separately by the transaction's immutable package state.
#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct Binding {
    content_types: BlobId,
    root: OwnerBinding,
    parts: Vec<PartBinding>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct OwnerBinding {
    owner: PackURI,
    relationships: Vec<RelationshipBinding>,
    token: RelationshipTokenBinding,
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct PartBinding {
    part: PackURI,
    content_type: String,
    relationships: Vec<RelationshipBinding>,
    token: RelationshipTokenBinding,
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct RelationshipTokenBinding {
    bytes: BlobId,
    member_present: bool,
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct RelationshipBinding {
    id: String,
    reltype: String,
    target: String,
    mode: TargetMode,
}

impl Binding {
    /// Capture bounded metadata for every package part and relationship owner.
    pub(crate) fn capture(package: &OpcPackage, limits: Limits) -> Result<Self> {
        let limits = limits.validate()?;
        preflight_graph(package, limits)?;

        let mut token_bytes = 0usize;
        let content_types = package.source_content_types()?;
        charge_token_bytes(&mut token_bytes, content_types.bytes().len(), limits)?;
        let content_types = BlobId::of(content_types.bytes());
        let root_owner = root_uri()?;
        let root = OwnerBinding {
            token: capture_token(package, &root_owner, &mut token_bytes, limits)?,
            owner: root_owner,
            relationships: capture_relationships(package.rels())?,
        };
        let mut parts = Vec::new();
        parts
            .try_reserve_exact(package.part_count())
            .map_err(|source| Error::Allocation {
                resource: "DOCX Ink graph part bindings",
                source,
            })?;
        for part in package.iter_parts() {
            parts.push(PartBinding {
                part: part.partname().clone(),
                content_type: clone_text(part.content_type(), "DOCX Ink graph content types")?,
                relationships: capture_relationships(part.rels())?,
                token: capture_token(package, part.partname(), &mut token_bytes, limits)?,
            });
        }
        parts.sort_unstable_by(|left, right| left.part.as_str().cmp(right.part.as_str()));
        Ok(Self {
            content_types,
            root,
            parts,
        })
    }

    /// Compare the complete current graph to this source binding.
    pub(crate) fn matches(&self, package: &OpcPackage, limits: Limits) -> Result<bool> {
        Ok(*self == Self::capture(package, limits)?)
    }
}

fn capture_token(
    package: &OpcPackage,
    owner: &PackURI,
    total: &mut usize,
    limits: Limits,
) -> Result<RelationshipTokenBinding> {
    let token = package.source_relationships(owner)?;
    charge_token_bytes(total, token.bytes().len(), limits)?;
    Ok(RelationshipTokenBinding {
        bytes: BlobId::of(token.bytes()),
        member_present: token.member_present(),
    })
}

fn charge_token_bytes(total: &mut usize, amount: usize, limits: Limits) -> Result<()> {
    *total = total.saturating_add(amount);
    if *total > limits.stories.max_topology_bytes {
        return Err(Error::InkLimit {
            resource: "Ink graph publication token bytes",
            actual: *total,
            maximum: limits.stories.max_topology_bytes,
        });
    }
    Ok(())
}

/// Exact, reversible closure for one story-owner relationship removal.
#[derive(Clone)]
pub(crate) struct Delta {
    inner: Arc<DeltaInner>,
    reversed: bool,
}

#[derive(Clone)]
struct DeltaInner {
    owner_before: OwnedRelationships,
    owner_after: OwnedRelationships,
    content_types_before: OwnedContentTypes,
    content_types_after: OwnedContentTypes,
    targets: Vec<GuardedTarget>,
}

#[derive(Clone)]
struct GuardedTarget {
    guard: TargetGuard,
    /// The target is absent in the forward state and is restored by inverse.
    removed: Option<RemovedTarget>,
}

#[derive(Clone)]
struct TargetGuard {
    part: PackURI,
    content_type: String,
    payload: Arc<Vec<u8>>,
    relationships: Vec<RelationshipBinding>,
    relationship_token: OwnedRelationships,
}

#[derive(Clone)]
struct RemovedTarget {
    source_xml: Option<OwnedXmlPart>,
    relationship_token: OwnedRelationships,
}

impl Delta {
    /// Remove selected owner relationship IDs and their exclusively-owned
    /// modeled Ink/image targets from an isolated candidate package.
    pub(crate) fn remove(
        package: &mut OpcPackage,
        owner: &PackURI,
        ids: &[String],
        limits: Limits,
    ) -> Result<Self> {
        let limits = limits.validate()?;
        preflight_graph(package, limits)?;
        let owner_name = canonical_owner(package, owner)?;
        let owner_relationships = owner_relationships(package, &owner_name)?;
        let owner_before = package.source_relationships(&owner_name)?;
        let content_types_before = package.source_content_types()?;
        let selected_ids = selected_ids(ids, limits)?;

        if selected_ids.is_empty() {
            return Ok(Self {
                inner: Arc::new(DeltaInner {
                    owner_after: owner_before.clone(),
                    owner_before,
                    content_types_after: content_types_before.clone(),
                    content_types_before,
                    targets: Vec::new(),
                }),
                reversed: false,
            });
        }

        let mut target_names = Vec::new();
        target_names
            .try_reserve(selected_ids.len())
            .map_err(|source| Error::Allocation {
                resource: "DOCX Ink graph target names",
                source,
            })?;
        for id in &selected_ids {
            let relationship = owner_relationships.get(id).ok_or_else(|| {
                Error::InvalidRelationship(format!("selected relationship '{id}' is missing"))
            })?;
            if relationship.is_external() {
                if !is_image_relationship(relationship.reltype())
                    || relationship.target_ref().is_empty()
                {
                    return Err(unsupported(
                        "selected external relationship is not a modeled image edge",
                    ));
                }
                continue;
            }

            let requested = relationship.target_partname()?;
            let target = package.get_part(&requested)?;
            validate_internal_edge(relationship.reltype(), target.content_type())?;
            target_names.push(target.partname().clone());
        }
        target_names.sort_unstable_by(|left, right| {
            cmp_ascii_case_insensitive(left.as_str(), right.as_str())
        });
        target_names.dedup_by(|left, right| left.is_equivalent_to(right));
        let incoming = remaining_incoming(package, &owner_name, &selected_ids, &target_names)?;

        let mut targets = Vec::new();
        targets
            .try_reserve_exact(target_names.len())
            .map_err(|source| Error::Allocation {
                resource: "DOCX Ink graph target guards",
                source,
            })?;
        let mut total_payload = 0usize;
        for (target_name, remaining_incoming) in target_names.into_iter().zip(incoming) {
            let part = package.get_part(&target_name)?;
            let payload = part.blob_arc();
            check_payload(payload.len(), limits, &mut total_payload)?;
            validate_target_payload(part.content_type(), payload.as_slice())?;
            let relationships = capture_relationships(part.rels())?;
            let relationship_token = package.source_relationships(part.partname())?;
            let removed = if remaining_incoming == 0 {
                if !part.rels().is_empty() {
                    return Err(unsupported(
                        "selected Ink resource has outgoing relationships and cannot be removed",
                    ));
                }
                let source_xml = if is_xml_content_type(part.content_type()) {
                    Some(package.source_xml_part(part.partname())?)
                } else {
                    None
                };
                Some(RemovedTarget {
                    source_xml,
                    relationship_token: relationship_token.clone(),
                })
            } else {
                None
            };
            targets.push(GuardedTarget {
                guard: TargetGuard {
                    part: part.partname().clone(),
                    content_type: clone_text(
                        part.content_type(),
                        "DOCX Ink graph target content types",
                    )?,
                    payload,
                    relationships,
                    relationship_token,
                },
                removed,
            });
        }

        let mut removed_names = Vec::new();
        removed_names
            .try_reserve(targets.len())
            .map_err(|source| Error::Allocation {
                resource: "DOCX Ink removed content-type names",
                source,
            })?;
        for target in &targets {
            if target.removed.is_some() {
                removed_names.push(target.guard.part.clone());
            }
        }
        let content_types_after = content_types_before.without_parts(
            &removed_names,
            ReadLimits::default().max_content_types_bytes(),
        )?;

        let mut owner_after = owner_before.clone();
        for id in &selected_ids {
            owner_after = owner_after.without_relationship(id, MAX_RELATIONSHIP_XML_BYTES)?;
        }
        package.try_replace_relationships(&owner_before, &owner_after)?;
        for target in &targets {
            if target.removed.is_some() && !package.remove_part(&target.guard.part) {
                return Err(Error::Invalid(
                    "DOCX Ink target disappeared during removal".into(),
                ));
            }
        }
        replace_content_types(package, &content_types_after)?;
        Ok(Self {
            inner: Arc::new(DeltaInner {
                owner_before,
                owner_after,
                content_types_before,
                content_types_after,
                targets,
            }),
            reversed: false,
        })
    }

    /// Apply this direction to a source-matching candidate package.
    pub(crate) fn apply(&self, package: &mut OpcPackage) -> Result<()> {
        let (expected_content, replacement_content) = if self.reversed {
            (
                &self.inner.content_types_after,
                &self.inner.content_types_before,
            )
        } else {
            (
                &self.inner.content_types_before,
                &self.inner.content_types_after,
            )
        };
        let current_content_types = package.source_content_types()?;
        if current_content_types.bytes() != expected_content.bytes() {
            return Err(Error::Invalid(
                "DOCX Ink graph content-types source is stale".into(),
            ));
        }
        let (expected_owner, replacement_owner) = if self.reversed {
            (&self.inner.owner_after, &self.inner.owner_before)
        } else {
            (&self.inner.owner_before, &self.inner.owner_after)
        };
        if package.source_relationships(expected_owner.owner())? != *expected_owner {
            return Err(Error::Invalid(
                "DOCX Ink graph relationship source is stale".into(),
            ));
        }
        self.validate_target_guards(package)?;

        if self.reversed {
            self.restore_targets(package)?;
            package.try_replace_relationships(expected_owner, replacement_owner)?;
        } else {
            package.try_replace_relationships(expected_owner, replacement_owner)?;
            for target in &self.inner.targets {
                if target.removed.is_some() && !package.remove_part(&target.guard.part) {
                    return Err(Error::Invalid(
                        "DOCX Ink target disappeared during replay".into(),
                    ));
                }
            }
        }
        replace_content_types(package, replacement_content)?;
        if package.source_relationships(replacement_owner.owner())? != *replacement_owner {
            return Err(Error::Invalid(
                "DOCX Ink graph delta result is stale".into(),
            ));
        }
        if package.source_content_types()?.bytes() != replacement_content.bytes() {
            return Err(Error::Invalid(
                "DOCX Ink graph content-types result is stale".into(),
            ));
        }
        Ok(())
    }

    /// Return the exact inverse direction while sharing all immutable guards.
    #[must_use]
    pub(crate) fn inverse(&self) -> Self {
        Self {
            inner: Arc::clone(&self.inner),
            reversed: !self.reversed,
        }
    }

    /// Whether this graph delta has no relationship or resource change.
    #[allow(
        dead_code,
        reason = "the transaction uses this when it folds an empty owner closure"
    )]
    #[must_use]
    pub(crate) fn is_empty(&self) -> bool {
        self.inner.owner_before == self.inner.owner_after
            && self
                .inner
                .targets
                .iter()
                .all(|target| target.removed.is_none())
    }

    /// Conservative bytes retained by this per-story graph delta.
    ///
    /// The payload charge intentionally counts shared `Arc` bytes for every
    /// delta.  The transaction may later deduplicate those allocations, but
    /// this bound keeps a large number of staged story edits finite without
    /// relying on allocator sharing.
    pub(crate) fn retained_bytes(&self) -> Result<usize> {
        let mut total = 0usize;
        for bytes in [
            self.inner.owner_before.bytes().len(),
            self.inner.owner_after.bytes().len(),
            self.inner.content_types_before.bytes().len(),
            self.inner.content_types_after.bytes().len(),
        ] {
            add_retained(&mut total, bytes)?;
        }
        for target in &self.inner.targets {
            add_retained(&mut total, target.guard.part.as_str().len())?;
            add_retained(&mut total, target.guard.content_type.len())?;
            add_retained(&mut total, target.guard.payload.len())?;
            add_retained(&mut total, target.guard.relationship_token.bytes().len())?;
            for relationship in &target.guard.relationships {
                add_retained(&mut total, relationship.id.len())?;
                add_retained(&mut total, relationship.reltype.len())?;
                add_retained(&mut total, relationship.target.len())?;
                add_retained(&mut total, size_of::<TargetMode>())?;
            }
            if let Some(removed) = &target.removed {
                add_retained(&mut total, removed.relationship_token.bytes().len())?;
                if let Some(source_xml) = &removed.source_xml {
                    add_retained(&mut total, source_xml.bytes().len())?;
                }
            }
        }
        Ok(total)
    }

    fn validate_target_guards(&self, package: &OpcPackage) -> Result<()> {
        for target in &self.inner.targets {
            let part = package.get_part(&target.guard.part);
            let should_exist = !self.reversed || target.removed.is_none();
            if !should_exist {
                match part {
                    Ok(_) => {
                        return Err(Error::Invalid(
                            "DOCX Ink inverse target unexpectedly exists".into(),
                        ));
                    },
                    Err(OpcError::PartNotFound(_)) => {},
                    Err(error) => return Err(Error::Opc(error)),
                }
                continue;
            }
            let part = part.map_err(Error::Opc)?;
            if part.content_type() != target.guard.content_type
                || (!std::ptr::eq(part.blob(), target.guard.payload.as_slice())
                    && part.blob() != target.guard.payload.as_slice())
                || !same_relationships(part.rels(), &target.guard.relationships)
                || package.source_relationships(part.partname())? != target.guard.relationship_token
            {
                return Err(Error::Invalid(
                    "DOCX Ink target payload or relationship source is stale".into(),
                ));
            }
        }
        Ok(())
    }

    fn restore_targets(&self, package: &mut OpcPackage) -> Result<()> {
        for target in &self.inner.targets {
            let Some(removed) = target.removed.as_ref() else {
                continue;
            };
            if let Some(source_xml) = &removed.source_xml {
                package.try_add_owned_xml_part(source_xml.clone())?;
            } else {
                package.try_add_part(Box::new(BlobPart::new_shared(
                    target.guard.part.clone(),
                    target.guard.content_type.clone(),
                    Arc::clone(&target.guard.payload),
                )))?;
            }
            let current = package.source_relationships(&target.guard.part)?;
            if current != removed.relationship_token {
                package.try_replace_relationships(&current, &removed.relationship_token)?;
            }
        }
        Ok(())
    }
}

fn root_uri() -> Result<PackURI> {
    PackURI::new("/").map_err(Error::Uri)
}

fn canonical_owner(package: &OpcPackage, owner: &PackURI) -> Result<PackURI> {
    if owner.as_str() == "/" {
        root_uri()
    } else {
        Ok(package.get_part(owner)?.partname().clone())
    }
}

fn owner_relationships<'a>(package: &'a OpcPackage, owner: &PackURI) -> Result<&'a Relationships> {
    if owner.as_str() == "/" {
        Ok(package.rels())
    } else {
        Ok(package.get_part(owner)?.rels())
    }
}

fn selected_ids(ids: &[String], limits: Limits) -> Result<Vec<String>> {
    if ids.len() > limits.max_relationships {
        return Err(Error::InkLimit {
            resource: "selected relationships",
            actual: ids.len(),
            maximum: limits.max_relationships,
        });
    }
    let mut unique = HashSet::new();
    unique
        .try_reserve(ids.len())
        .map_err(|source| Error::Allocation {
            resource: "DOCX Ink selected relationship IDs",
            source,
        })?;
    for id in ids {
        if !unique.insert(id.as_str()) {
            return Err(Error::InvalidRelationship(
                "selected relationship IDs contain a duplicate".into(),
            ));
        }
    }
    let mut selected = Vec::new();
    selected
        .try_reserve_exact(ids.len())
        .map_err(|source| Error::Allocation {
            resource: "DOCX Ink selected relationship IDs",
            source,
        })?;
    selected.extend(ids.iter().cloned());
    selected.sort_unstable();
    Ok(selected)
}

fn remaining_incoming(
    package: &OpcPackage,
    selected_owner: &PackURI,
    selected_ids: &[String],
    targets: &[PackURI],
) -> Result<Vec<usize>> {
    let mut incoming = Vec::new();
    incoming
        .try_reserve_exact(targets.len())
        .map_err(|source| Error::Allocation {
            resource: "DOCX Ink incoming relationship census",
            source,
        })?;
    incoming.resize(targets.len(), 0);
    let root = root_uri()?;
    for relationship in package.rels().iter() {
        count_incoming(
            &mut incoming,
            targets,
            relationship,
            package,
            &root,
            selected_owner,
            selected_ids,
        )?;
    }
    for part in package.iter_parts() {
        let owner = part.partname();
        for relationship in part.rels().iter() {
            count_incoming(
                &mut incoming,
                targets,
                relationship,
                package,
                owner,
                selected_owner,
                selected_ids,
            )?;
        }
    }
    Ok(incoming)
}

fn count_incoming(
    incoming: &mut [usize],
    targets: &[PackURI],
    relationship: &litchi_opc::Relationship,
    package: &OpcPackage,
    owner: &PackURI,
    selected_owner: &PackURI,
    selected_ids: &[String],
) -> Result<()> {
    if relationship.is_external() {
        return Ok(());
    }
    if owner.is_equivalent_to(selected_owner)
        && selected_ids
            .binary_search_by(|id| id.as_str().cmp(relationship.r_id()))
            .is_ok()
    {
        return Ok(());
    }
    let requested = relationship.target_partname()?;
    let part = package.get_part(&requested)?;
    let Some(index) = targets
        .binary_search_by(|target| {
            cmp_ascii_case_insensitive(target.as_str(), part.partname().as_str())
        })
        .ok()
    else {
        return Ok(());
    };
    incoming[index] = incoming[index].checked_add(1).ok_or(Error::InkLimit {
        resource: "remaining incoming relationships",
        actual: usize::MAX,
        maximum: usize::MAX,
    })?;
    Ok(())
}

fn validate_internal_edge(reltype: &str, content_type: &str) -> Result<()> {
    if is_ink_content_type(content_type) {
        if is_custom_xml_relationship(reltype) {
            return Ok(());
        }
    } else if is_image_content_type(content_type) && is_image_relationship(reltype) {
        return Ok(());
    }
    Err(unsupported(
        "selected relationship targets an unsupported resource edge",
    ))
}

fn validate_target_payload(content_type: &str, payload: &[u8]) -> Result<()> {
    if is_ink_content_type(content_type) {
        if !super::package::has_ink_root(payload)? {
            return Err(unsupported(
                "selected XML resource is not a namespace-bound InkML part",
            ));
        }
        let projection = super::package::validate_content_part(payload)?;
        let _ = shared::read_metadata_with_source_spans(
            payload,
            projection.contexts(),
            projection.traces(),
            projection.brush_properties(),
            projection.links(),
        )?;
    }
    Ok(())
}

fn is_xml_content_type(content_type: &str) -> bool {
    content_type.eq_ignore_ascii_case("text/xml") || is_ink_content_type(content_type)
}

fn is_ink_content_type(content_type: &str) -> bool {
    content_type.eq_ignore_ascii_case(CONTENT_TYPE) || content_type.eq_ignore_ascii_case("text/xml")
}

fn is_image_content_type(content_type: &str) -> bool {
    let bytes = content_type.as_bytes();
    bytes.len() > 6 && bytes[..6].eq_ignore_ascii_case(b"image/")
}

fn is_custom_xml_relationship(reltype: &str) -> bool {
    reltype == rt::CUSTOM_XML || reltype == STRICT_CUSTOM_XML
}

fn is_image_relationship(reltype: &str) -> bool {
    reltype == rt::IMAGE || reltype == rt::STRICT_IMAGE
}

fn cmp_ascii_case_insensitive(left: &str, right: &str) -> std::cmp::Ordering {
    for (left, right) in left.bytes().zip(right.bytes()) {
        let ordering = left.to_ascii_lowercase().cmp(&right.to_ascii_lowercase());
        if ordering != std::cmp::Ordering::Equal {
            return ordering;
        }
    }
    left.len().cmp(&right.len())
}

fn check_payload(length: usize, limits: Limits, total: &mut usize) -> Result<()> {
    if length > limits.max_payload_bytes {
        return Err(Error::InkLimit {
            resource: "Ink graph target payload bytes",
            actual: length,
            maximum: limits.max_payload_bytes,
        });
    }
    *total = total.checked_add(length).ok_or(Error::InkLimit {
        resource: "Ink graph target payload bytes",
        actual: usize::MAX,
        maximum: limits.max_total_payload_bytes,
    })?;
    if *total > limits.max_total_payload_bytes {
        return Err(Error::InkLimit {
            resource: "Ink graph target payload bytes",
            actual: *total,
            maximum: limits.max_total_payload_bytes,
        });
    }
    Ok(())
}

fn same_relationships(relationships: &Relationships, expected: &[RelationshipBinding]) -> bool {
    relationships.len() == expected.len()
        && expected.iter().all(|entry| {
            relationships.get(&entry.id).is_some_and(|relationship| {
                relationship.reltype() == entry.reltype
                    && relationship.target_ref() == entry.target
                    && relationship.target_mode() == entry.mode
            })
        })
}

fn capture_relationships(relationships: &Relationships) -> Result<Vec<RelationshipBinding>> {
    let mut entries = Vec::new();
    entries
        .try_reserve_exact(relationships.len())
        .map_err(|source| Error::Allocation {
            resource: "DOCX Ink graph relationship bindings",
            source,
        })?;
    for relationship in relationships.iter() {
        if !relationship.is_external() {
            relationship.target_partname()?;
        }
        entries.push(RelationshipBinding {
            id: clone_text(relationship.r_id(), "DOCX Ink graph relationship IDs")?,
            reltype: clone_text(relationship.reltype(), "DOCX Ink graph relationship types")?,
            target: clone_text(
                relationship.target_ref(),
                "DOCX Ink graph relationship targets",
            )?,
            mode: relationship.target_mode(),
        });
    }
    entries.sort_unstable_by(|left, right| left.id.cmp(&right.id));
    Ok(entries)
}

fn clone_text(value: &str, resource: &'static str) -> Result<String> {
    let mut cloned = String::new();
    cloned
        .try_reserve_exact(value.len())
        .map_err(|source| Error::Allocation { resource, source })?;
    cloned.push_str(value);
    Ok(cloned)
}

fn preflight_graph(package: &OpcPackage, limits: Limits) -> Result<usize> {
    let part_count = package.part_count();
    if part_count > limits.stories.max_package_parts {
        return Err(Error::InkLimit {
            resource: "Ink graph package parts",
            actual: part_count,
            maximum: limits.stories.max_package_parts,
        });
    }
    let mut topology = 0usize;
    let mut relationships = 0usize;
    measure_owner(
        "/",
        package.rels(),
        &mut topology,
        &mut relationships,
        limits,
    )?;
    for part in package.iter_parts() {
        add_topology(&mut topology, part.partname().as_str().len(), limits)?;
        add_topology(&mut topology, part.content_type().len(), limits)?;
        measure_owner(
            part.partname().as_str(),
            part.rels(),
            &mut topology,
            &mut relationships,
            limits,
        )?;
    }
    Ok(relationships)
}

fn measure_owner(
    owner: &str,
    relationships: &Relationships,
    topology: &mut usize,
    relationship_count: &mut usize,
    limits: Limits,
) -> Result<()> {
    add_topology(topology, owner.len().saturating_add(8), limits)?;
    for relationship in relationships.iter() {
        *relationship_count = relationship_count.checked_add(1).ok_or(Error::InkLimit {
            resource: "Ink graph relationships",
            actual: usize::MAX,
            maximum: limits.max_relationships,
        })?;
        if *relationship_count > limits.max_relationships {
            return Err(Error::InkLimit {
                resource: "Ink graph relationships",
                actual: *relationship_count,
                maximum: limits.max_relationships,
            });
        }
        if !relationship.is_external() {
            relationship.target_partname()?;
        }
        let fixed = 24usize;
        add_topology(topology, fixed, limits)?;
        add_topology(topology, relationship.r_id().len(), limits)?;
        add_topology(topology, relationship.reltype().len(), limits)?;
        add_topology(topology, relationship.target_ref().len(), limits)?;
    }
    Ok(())
}

fn add_topology(total: &mut usize, amount: usize, limits: Limits) -> Result<()> {
    *total = total.checked_add(amount).ok_or(Error::InkLimit {
        resource: "Ink graph topology bytes",
        actual: usize::MAX,
        maximum: limits.stories.max_topology_bytes,
    })?;
    if *total > limits.stories.max_topology_bytes {
        return Err(Error::InkLimit {
            resource: "Ink graph topology bytes",
            actual: *total,
            maximum: limits.stories.max_topology_bytes,
        });
    }
    Ok(())
}

fn replace_content_types(package: &mut OpcPackage, replacement: &OwnedContentTypes) -> Result<()> {
    let current = package.source_content_types()?;
    if current.bytes() == replacement.bytes() {
        return Ok(());
    }
    let expected = current.bytes().to_vec();
    package.try_replace_content_types(&expected, replacement)?;
    Ok(())
}

fn add_retained(total: &mut usize, amount: usize) -> Result<()> {
    *total = total
        .checked_add(amount)
        .ok_or_else(|| Error::Invalid("DOCX Ink graph retained-byte charge overflow".into()))?;
    Ok(())
}

fn unsupported(reason: &'static str) -> Error {
    Error::UnsafeEdit {
        format: "DOCX",
        operation: "remove_ink",
        reason,
    }
}
