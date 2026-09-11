//! Relationship-only topology preflight.
//!
//! This pass constructs the final typed relationship collections and computes
//! exact source-member output lengths/event gauges before the publication loop
//! allocates replacement XML.  The publication path still performs its strict
//! lexical readback; this module only admits the candidate's aggregate limits.

use super::{
    RelationshipPublicationSummary, SourceBackedPackage, TopologyRelationshipChange,
    TopologyRelationshipOperation, canonical_relationship_comparison_memory_bound,
    canonical_relationship_source_matches, parse_noncanonical_relationship_source,
    relationship_append_fragment_len, relationship_source_append_output_len,
    relationship_source_output_len,
};
use crate::error::{OpcError, Result};
use crate::packuri::{PACKAGE_URI, PackURI};
use crate::rel::Relationships;
use soapberry_zip::office::EntryId;
use std::collections::HashMap;
use std::ops::Range;

/// One relationship member's exact preflight summary.
#[derive(Debug)]
pub(crate) struct RelationshipPreflight {
    pub(crate) owner: PackURI,
    pub(crate) member_name: String,
    pub(crate) existing_entry: Option<EntryId>,
    pub(crate) relationship_count: usize,
    pub(crate) final_bytes: u64,
    pub(crate) final_events: u64,
    pub(crate) source_event_count: Option<u64>,
}

impl RelationshipPublicationSummary for RelationshipPreflight {
    fn owner(&self) -> &PackURI {
        &self.owner
    }

    fn member_name(&self) -> &str {
        &self.member_name
    }

    fn existing_entry(&self) -> Option<EntryId> {
        self.existing_entry
    }

    fn relationship_count(&self) -> usize {
        self.relationship_count
    }

    fn output_bytes(&self) -> Result<u64> {
        Ok(self.final_bytes)
    }

    fn output_events(&self, _limits: super::ReadLimits) -> Result<u64> {
        Ok(self.final_events)
    }

    fn source_event_count(&self) -> Option<u64> {
        self.source_event_count
    }
}

/// Build final relationship overrides and exact output summaries without
/// constructing candidate relationship XML.
pub(crate) fn plan_relationship_publications(
    package: &SourceBackedPackage,
    changes: &[TopologyRelationshipChange],
    physical_members: &super::PhysicalMemberLookup,
) -> Result<(HashMap<PackURI, Relationships>, Vec<RelationshipPreflight>)> {
    let mut overrides = HashMap::new();
    overrides
        .try_reserve(changes.len())
        .map_err(|source| OpcError::Allocation {
            resource: "source-backed OPC relationship preflight overrides",
            source,
        })?;
    let mut summaries = Vec::new();
    summaries
        .try_reserve(changes.len())
        .map_err(|source| OpcError::Allocation {
            resource: "source-backed OPC relationship preflight summaries",
            source,
        })?;

    let mut group_start = 0usize;
    while group_start < changes.len() {
        package.check_topology_progress()?;
        let owner = changes[group_start].owner.clone();
        let mut group_end = group_start + 1;
        while group_end < changes.len() && changes[group_end].owner == owner {
            group_end += 1;
        }
        let owner_index = if owner.as_str() == PACKAGE_URI {
            None
        } else {
            package.part_index(&owner)
        };
        let (mut effective, owner_base) = if owner.as_str() == PACKAGE_URI {
            (package.package_relationships.clone(), "/".to_string())
        } else if let Some(index) = owner_index {
            (
                package.parts[index].relationships.clone(),
                package.parts[index].partname.base_uri().to_string(),
            )
        } else {
            (
                Relationships::for_source(&owner),
                owner.base_uri().to_string(),
            )
        };
        let mut changed = false;
        for change in &changes[group_start..group_end] {
            match &change.operation {
                TopologyRelationshipOperation::Add { reltype, target } => {
                    if effective.get(&change.r_id).is_some() {
                        return Err(OpcError::DuplicateRelationshipId(change.r_id.clone()));
                    }
                    let (target_ref, target_mode) =
                        SourceBackedPackage::topology_relationship_target_ref(target, &owner_base)?;
                    package.validate_topology_relationship_field_limits(
                        &change.r_id,
                        reltype,
                        &target_ref,
                    )?;
                    if package.has_signature_infrastructure() {
                        return Err(OpcError::SignedSourceRequiresExplicitPolicy);
                    }
                    effective.try_add_relationship(
                        reltype.clone(),
                        target_ref,
                        change.r_id.clone(),
                        target_mode,
                    )?;
                    changed = true;
                },
                TopologyRelationshipOperation::Replace {
                    reltype,
                    target,
                    required_mode,
                } => {
                    let existing = effective.get(&change.r_id).ok_or_else(|| {
                        OpcError::RelationshipNotFound(format!(
                            "relationship '{}' was not found",
                            change.r_id
                        ))
                    })?;
                    if required_mode.is_some_and(|mode| existing.target_mode() != mode) {
                        return Err(OpcError::InvalidRelationship(format!(
                            "relationship '{}' does not have the required target mode",
                            change.r_id
                        )));
                    }
                    let (target_ref, target_mode) =
                        SourceBackedPackage::topology_relationship_target_ref(target, &owner_base)?;
                    package.validate_topology_relationship_field_limits(
                        &change.r_id,
                        reltype,
                        &target_ref,
                    )?;
                    if existing.reltype() == reltype
                        && existing.target_ref() == target_ref
                        && existing.target_mode() == target_mode
                    {
                        continue;
                    }
                    if package.has_signature_infrastructure() {
                        return Err(OpcError::SignedSourceRequiresExplicitPolicy);
                    }
                    effective.remove(&change.r_id);
                    effective.try_add_relationship(
                        reltype.clone(),
                        target_ref,
                        change.r_id.clone(),
                        target_mode,
                    )?;
                    changed = true;
                },
                TopologyRelationshipOperation::Remove { required_mode } => {
                    let existing = effective.get(&change.r_id).ok_or_else(|| {
                        OpcError::RelationshipNotFound(format!(
                            "relationship '{}' was not found",
                            change.r_id
                        ))
                    })?;
                    if required_mode.is_some_and(|mode| existing.target_mode() != mode) {
                        return Err(OpcError::InvalidRelationship(format!(
                            "relationship '{}' does not have the required target mode",
                            change.r_id
                        )));
                    }
                    if package.has_signature_infrastructure() {
                        return Err(OpcError::SignedSourceRequiresExplicitPolicy);
                    }
                    effective.remove(&change.r_id);
                    changed = true;
                },
            }
        }
        if changed {
            overrides.insert(owner.clone(), effective);
        }
        group_start = group_end;
    }

    for (owner, effective) in &overrides {
        package.check_topology_progress()?;
        let relationship_uri = owner.rels_uri().map_err(OpcError::InvalidPackUri)?;
        let member_name = relationship_uri.membername().to_owned();
        let existing_entry =
            package.source_entry_id_case_insensitive(&member_name, physical_members)?;
        let source = if owner.as_str() == PACKAGE_URI {
            Some(&package.package_relationships)
        } else {
            package
                .part_index(owner)
                .map(|index| &package.parts[index].relationships)
        };
        let relationship_count = effective.len();
        let canonical_len = super::canonical_relationship_xml_len(effective)?;
        let canonical_events = relationship_count
            .checked_add(4)
            .ok_or_else(|| super::overlay_unavailable("relationship event count overflows"))?
            as u64;
        let Some(entry_id) = existing_entry else {
            summaries.push(RelationshipPreflight {
                owner: owner.clone(),
                member_name,
                existing_entry: None,
                relationship_count,
                final_bytes: canonical_len as u64,
                final_events: canonical_events,
                source_event_count: None,
            });
            continue;
        };
        let source = source.ok_or_else(|| {
            super::overlay_unavailable("existing relationships member has no source catalog")
        })?;
        let metadata = package.archive.metadata_for(entry_id)?;
        let declared_bytes = metadata.uncompressed_size();
        package.limits.check(
            crate::limits::ReadResource::RelationshipXmlBytes,
            declared_bytes,
            package.limits.max_relationship_xml_bytes() as u64,
        )?;
        package.limits.check(
            crate::limits::ReadResource::ArchiveEntryBytes,
            declared_bytes,
            package.limits.max_archive_entry_bytes(),
        )?;
        // Preflight reads and parses the source member to classify its
        // lexical profile. Charge that temporary source/parser state before
        // the read; the reservation is scoped to this classification pass.
        let _source_memory_reservation = package.reserve_topology_memory(
            super::relationship_xml_working_memory_bound(declared_bytes, source.len())?,
        )?;
        let original = package
            .archive
            .read_entry(entry_id)
            .map_err(super::map_preservation_error)?;
        package.check_topology_progress()?;
        if original.len() as u64 != declared_bytes {
            return Err(OpcError::ZipError(format!(
                "source-backed OPC relationships member '{}' declared {declared_bytes} uncompressed bytes but read {}",
                member_name,
                original.len()
            )));
        }
        package.limits.check(
            crate::limits::ReadResource::RelationshipXmlBytes,
            canonical_len as u64,
            package.limits.max_relationship_xml_bytes() as u64,
        )?;
        let _comparison_memory = package.reserve_topology_memory(
            canonical_relationship_comparison_memory_bound(source.len())?,
        )?;
        let original_is_canonical = canonical_relationship_source_matches(&original, source)?;
        if original_is_canonical {
            summaries.push(RelationshipPreflight {
                owner: owner.clone(),
                member_name,
                existing_entry: Some(entry_id),
                relationship_count,
                final_bytes: canonical_len as u64,
                final_events: canonical_events,
                source_event_count: package
                    .relationship_xml_metrics
                    .get(owner)
                    .map(|metrics| metrics.event_count),
            });
            continue;
        }
        let group = changes_for_owner(changes, owner);
        let append_only = group
            .iter()
            .all(|change| matches!(change.operation, TopologyRelationshipOperation::Add { .. }));
        let removal_only = group.iter().all(|change| {
            matches!(
                change.operation,
                TopologyRelationshipOperation::Remove { .. }
            )
        });
        if !append_only && !removal_only {
            return Err(super::overlay_unavailable(format!(
                "existing relationships member '{}' is not canonical",
                member_name
            )));
        }
        let parsed = parse_noncanonical_relationship_source(
            &original,
            &member_name,
            source,
            package.limits,
            || package.check_topology_progress(),
        )?;
        if removal_only {
            let mut ranges: Vec<Range<usize>> = Vec::new();
            let mut removed_events = 0_u64;
            ranges
                .try_reserve_exact(group.len())
                .map_err(|source| OpcError::Allocation {
                    resource: "source-backed OPC relationship preflight removal ranges",
                    source,
                })?;
            for change in group {
                let range = parsed
                    .relationship_ranges
                    .get(&change.r_id)
                    .ok_or_else(|| {
                        super::overlay_unavailable(format!(
                            "source relationships member '{}' has no range for '{}'",
                            member_name, change.r_id
                        ))
                    })?;
                ranges.push(range.clone());
                removed_events = removed_events
                    .checked_add(
                        *parsed
                            .relationship_event_counts
                            .get(&change.r_id)
                            .ok_or_else(|| {
                                super::overlay_unavailable(
                                    "relationship event count is unavailable",
                                )
                            })?,
                    )
                    .ok_or_else(|| {
                        super::overlay_unavailable("relationship event count overflows")
                    })?;
            }
            ranges.sort_unstable_by_key(|range| range.start);
            let final_len = relationship_source_output_len(original.len(), &ranges)?;
            let final_events = parsed
                .event_count
                .checked_sub(removed_events)
                .ok_or_else(|| super::overlay_unavailable("relationship event count underflows"))?;
            summaries.push(RelationshipPreflight {
                owner: owner.clone(),
                member_name,
                existing_entry: Some(entry_id),
                relationship_count,
                final_bytes: final_len as u64,
                final_events,
                source_event_count: package
                    .relationship_xml_metrics
                    .get(owner)
                    .map(|metrics| metrics.event_count),
            });
        } else {
            let mut append_len = 0usize;
            for change in group {
                let TopologyRelationshipOperation::Add { reltype, target } = &change.operation
                else {
                    return Err(super::overlay_unavailable(
                        "relationship preflight operation mismatch",
                    ));
                };
                let (target_ref, target_mode) =
                    SourceBackedPackage::topology_relationship_target_ref(
                        target,
                        owner.base_uri(),
                    )?;
                append_len = append_len
                    .checked_add(relationship_append_fragment_len(
                        parsed.root_prefix.len(),
                        &change.r_id,
                        reltype,
                        &target_ref,
                        target_mode,
                    )?)
                    .ok_or_else(|| {
                        super::overlay_unavailable("relationship append length overflows")
                    })?;
            }
            let final_len = relationship_source_append_output_len(
                original.len(),
                parsed.root_empty_end.is_some(),
                parsed.root_name.len(),
                append_len,
            )?;
            let final_events = parsed
                .event_count
                .checked_add(u64::from(parsed.root_empty_end.is_some()))
                .ok_or_else(|| super::overlay_unavailable("relationship event count overflows"))?
                .checked_add(group.len() as u64)
                .ok_or_else(|| super::overlay_unavailable("relationship event count overflows"))?;
            summaries.push(RelationshipPreflight {
                owner: owner.clone(),
                member_name,
                existing_entry: Some(entry_id),
                relationship_count,
                final_bytes: final_len as u64,
                final_events,
                source_event_count: package
                    .relationship_xml_metrics
                    .get(owner)
                    .map(|metrics| metrics.event_count),
            });
        }
    }
    summaries.sort_unstable_by(|left, right| left.member_name.cmp(&right.member_name));
    if summaries.windows(2).any(|pair| {
        pair[0]
            .member_name
            .eq_ignore_ascii_case(&pair[1].member_name)
    }) {
        return Err(super::overlay_unavailable(
            "multiple topology relationship owners resolve to one member",
        ));
    }
    Ok((overrides, summaries))
}

fn changes_for_owner<'a>(
    changes: &'a [TopologyRelationshipChange],
    owner: &PackURI,
) -> &'a [TopologyRelationshipChange] {
    let Ok(mut index) =
        changes.binary_search_by(|change| change.owner.as_str().cmp(owner.as_str()))
    else {
        return &[];
    };
    while index > 0 && changes[index - 1].owner == *owner {
        index -= 1;
    }
    let start = index;
    let mut end = start + 1;
    while end < changes.len() && changes[end].owner == *owner {
        end += 1;
    }
    &changes[start..end]
}
