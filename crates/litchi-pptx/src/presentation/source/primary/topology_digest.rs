//! Canonical typed fingerprints for source-backed OPC topology.
//!
//! The digest deliberately observes only the admitted typed catalog: package
//! relationships, every Part name and content type, and every relationship's
//! identity, type, mode, and resolved target.  It does not read Part payloads,
//! relationship XML, ZIP order, or PresentationML slide order.  This makes it
//! suitable as a source-bound topology proof for structural primary patches.

use std::cmp::Ordering;
use std::mem::size_of;

use litchi_core::{ExecutionContext, ExecutionError, Reservation, Resource};
use litchi_opc::{
    OpcError, PackURI, PreparedTopology, Relationships, SourceBackedPackage, SourceLimitKind,
    SourceLimits, TargetMode,
};
use sha2::{Digest as _, Sha256};

use crate::{Error, Result};

const DIGEST_DOMAIN: &[u8] = b"litchi-pptx-source-typed-topology-v1\0";

/// Return the canonical typed topology digest for an admitted source.
///
/// Only metadata is visited.  The source and execution fences are checked
/// before and after the operation and around the metadata walk, so a source
/// change cannot be mistaken for a stable topology proof.
pub(super) fn source_digest(package: &SourceBackedPackage) -> Result<[u8; 32]> {
    package.source_version()?;
    package.check_execution()?;
    let context = package.execution_context();
    let limits = package.read_limits();
    let stats = source_stats(package, limits, context.as_ref())?;
    let _workspace = reserve_workspace(context.as_ref(), workspace_bound(&stats)?)?;
    let records = source_records(package, &stats, context.as_ref())?;
    let digest = digest_records(records, context.as_ref())?;
    package.source_version()?;
    package.check_execution()?;
    Ok(digest)
}

/// Return the canonical typed topology digest for a prepared candidate.
///
/// Prepared Part views are metadata-only handles; no effective Part payload is
/// read.  The prepared source/donor fences are checked before, during, and
/// after catalog collection.
pub(super) fn prepared_digest(prepared: &PreparedTopology) -> Result<[u8; 32]> {
    prepared.check_source()?;
    prepared.check_execution()?;
    let context = prepared.execution_context();
    let limits = prepared.read_limits();
    let stats = prepared_stats(prepared, limits, context.as_ref())?;
    let _workspace = reserve_workspace(context.as_ref(), workspace_bound(&stats)?)?;
    let records = prepared_records(prepared, &stats, context.as_ref())?;
    let digest = digest_records(records, context.as_ref())?;
    prepared.check_source()?;
    prepared.check_execution()?;
    Ok(digest)
}

#[derive(Default)]
struct Stats {
    parts: usize,
    relationships: usize,
    string_bytes: usize,
}

struct Records {
    package_relationships: Vec<RelationshipRecord>,
    parts: Vec<PartRecord>,
}

struct PartRecord {
    name: String,
    content_type: String,
    relationships: Vec<RelationshipRecord>,
}

struct RelationshipRecord {
    id: String,
    reltype: String,
    target: String,
    mode: u8,
}

fn source_stats(
    package: &SourceBackedPackage,
    limits: SourceLimits,
    context: Option<&ExecutionContext>,
) -> Result<Stats> {
    let parts = package.iter_parts().count();
    check_count(SourceLimitKind::Parts, parts, limits.max_parts())?;
    let mut stats = Stats {
        parts,
        ..Stats::default()
    };
    let mut checkpoint = || -> Result<()> {
        package.source_version()?;
        package.check_execution()?;
        Ok(())
    };
    inspect_relationships(package.rels(), limits, &mut stats, &mut checkpoint)?;
    for part in package.iter_parts() {
        checkpoint()?;
        inspect_part(
            part.partname(),
            part.content_type(),
            part.rels(),
            limits,
            &mut stats,
            &mut checkpoint,
        )?;
        check_context(context)?;
    }
    checkpoint()?;
    Ok(stats)
}

fn prepared_stats(
    prepared: &PreparedTopology,
    limits: SourceLimits,
    context: Option<&ExecutionContext>,
) -> Result<Stats> {
    let parts = prepared.iter_parts()?.len();
    check_count(SourceLimitKind::Parts, parts, limits.max_parts())?;
    let mut stats = Stats {
        parts,
        ..Stats::default()
    };
    let mut checkpoint = || -> Result<()> {
        prepared.check_source()?;
        prepared.check_execution()?;
        Ok(())
    };
    inspect_relationships(prepared.rels(), limits, &mut stats, &mut checkpoint)?;
    for part in prepared.iter_parts()? {
        checkpoint()?;
        inspect_part(
            part.partname(),
            part.content_type(),
            part.rels(),
            limits,
            &mut stats,
            &mut checkpoint,
        )?;
        check_context(context)?;
    }
    checkpoint()?;
    Ok(stats)
}

fn inspect_part(
    name: &PackURI,
    content_type: &str,
    relationships: &Relationships,
    limits: SourceLimits,
    stats: &mut Stats,
    checkpoint: &mut impl FnMut() -> Result<()>,
) -> Result<()> {
    check_len(
        SourceLimitKind::ArchiveMemberNameBytes,
        name.as_str().len(),
        limits.max_archive_member_name_bytes(),
    )?;
    check_len(
        SourceLimitKind::XmlAttributeBytes,
        content_type.len(),
        limits.max_xml_attribute_bytes() as u64,
    )?;
    add_string_bytes(&mut stats.string_bytes, name.as_str().len())?;
    add_string_bytes(&mut stats.string_bytes, content_type.len())?;
    inspect_relationships(relationships, limits, stats, checkpoint)
}

fn inspect_relationships(
    relationships: &Relationships,
    limits: SourceLimits,
    stats: &mut Stats,
    checkpoint: &mut impl FnMut() -> Result<()>,
) -> Result<()> {
    check_count(
        SourceLimitKind::RelationshipsPerPart,
        relationships.len(),
        limits.max_relationships_per_part(),
    )?;
    stats.relationships = stats
        .relationships
        .checked_add(relationships.len())
        .ok_or_else(|| invalid_topology("relationship count overflows"))?;
    check_count(
        SourceLimitKind::TotalRelationships,
        stats.relationships,
        limits.max_total_relationships(),
    )?;
    for relationship in relationships.iter() {
        checkpoint()?;
        check_len(
            SourceLimitKind::XmlAttributeBytes,
            relationship.r_id().len(),
            limits.max_xml_attribute_bytes() as u64,
        )?;
        check_len(
            SourceLimitKind::XmlAttributeBytes,
            relationship.reltype().len(),
            limits.max_xml_attribute_bytes() as u64,
        )?;
        let target_len = canonical_target_len(relationship)?;
        check_len(
            SourceLimitKind::RelationshipTargetBytes,
            target_len,
            limits.max_relationship_target_bytes() as u64,
        )?;
        add_string_bytes(&mut stats.string_bytes, relationship.r_id().len())?;
        add_string_bytes(&mut stats.string_bytes, relationship.reltype().len())?;
        add_string_bytes(&mut stats.string_bytes, target_len)?;
    }
    Ok(())
}

fn source_records(
    package: &SourceBackedPackage,
    stats: &Stats,
    context: Option<&ExecutionContext>,
) -> Result<Records> {
    let mut checkpoint = || -> Result<()> {
        package.source_version()?;
        package.check_execution()?;
        Ok(())
    };
    let mut package_relationships = Vec::new();
    package_relationships
        .try_reserve_exact(package.rels().len())
        .map_err(|source| allocation("source topology digest relationships", source))?;
    collect_relationships(package.rels(), &mut package_relationships, &mut checkpoint)?;

    let mut parts = Vec::new();
    parts
        .try_reserve_exact(stats.parts)
        .map_err(|source| allocation("source topology digest parts", source))?;
    for part in package.iter_parts() {
        checkpoint()?;
        let mut relationships = Vec::new();
        relationships
            .try_reserve_exact(part.rels().len())
            .map_err(|source| allocation("source topology digest relationships", source))?;
        collect_relationships(part.rels(), &mut relationships, &mut checkpoint)?;
        parts.push(PartRecord {
            name: clone_string(part.partname().as_str(), "source topology digest part name")?,
            content_type: clone_string(part.content_type(), "source topology digest content type")?,
            relationships,
        });
        check_context(context)?;
    }
    checkpoint()?;
    Ok(Records {
        package_relationships,
        parts,
    })
}

fn prepared_records(
    prepared: &PreparedTopology,
    stats: &Stats,
    context: Option<&ExecutionContext>,
) -> Result<Records> {
    let mut checkpoint = || -> Result<()> {
        prepared.check_source()?;
        prepared.check_execution()?;
        Ok(())
    };
    let mut package_relationships = Vec::new();
    package_relationships
        .try_reserve_exact(prepared.rels().len())
        .map_err(|source| allocation("prepared topology digest relationships", source))?;
    collect_relationships(prepared.rels(), &mut package_relationships, &mut checkpoint)?;

    let mut parts = Vec::new();
    parts
        .try_reserve_exact(stats.parts)
        .map_err(|source| allocation("prepared topology digest parts", source))?;
    for part in prepared.iter_parts()? {
        checkpoint()?;
        let mut relationships = Vec::new();
        relationships
            .try_reserve_exact(part.rels().len())
            .map_err(|source| allocation("prepared topology digest relationships", source))?;
        collect_relationships(part.rels(), &mut relationships, &mut checkpoint)?;
        parts.push(PartRecord {
            name: clone_string(
                part.partname().as_str(),
                "prepared topology digest part name",
            )?,
            content_type: clone_string(
                part.content_type(),
                "prepared topology digest content type",
            )?,
            relationships,
        });
        check_context(context)?;
    }
    checkpoint()?;
    Ok(Records {
        package_relationships,
        parts,
    })
}

fn collect_relationships(
    relationships: &Relationships,
    output: &mut Vec<RelationshipRecord>,
    checkpoint: &mut impl FnMut() -> Result<()>,
) -> Result<()> {
    for relationship in relationships.iter() {
        checkpoint()?;
        output.push(RelationshipRecord {
            id: clone_string(
                relationship.r_id(),
                "source topology digest relationship ID",
            )?,
            reltype: clone_string(
                relationship.reltype(),
                "source topology digest relationship type",
            )?,
            target: canonical_target(relationship)?,
            mode: relationship_mode(relationship.target_mode()),
        });
    }
    Ok(())
}

fn canonical_target(relationship: &litchi_opc::Relationship) -> Result<String> {
    match relationship.target_mode() {
        TargetMode::Internal => {
            let target = relationship.target_partname()?;
            let query = relationship.target_query();
            let fragment = relationship.target_fragment();
            let suffix_bytes = query.map_or(0, |value| 1 + value.len())
                + fragment.map_or(0, |value| 1 + value.len());
            let capacity = target
                .as_str()
                .len()
                .checked_add(suffix_bytes)
                .ok_or_else(|| invalid_topology("internal target workspace overflows"))?;
            let mut output = String::new();
            output
                .try_reserve_exact(capacity)
                .map_err(|source| Error::Allocation {
                    resource: "source topology digest internal target",
                    source,
                })?;
            output.push_str(target.as_str());
            if let Some(query) = query {
                output.push('?');
                output.push_str(query);
            }
            if let Some(fragment) = fragment {
                output.push('#');
                output.push_str(fragment);
            }
            Ok(output)
        },
        TargetMode::External => clone_string(
            relationship.target_ref(),
            "source topology digest external target",
        ),
    }
}

fn canonical_target_len(relationship: &litchi_opc::Relationship) -> Result<usize> {
    match relationship.target_mode() {
        TargetMode::Internal => {
            let absolute = relationship.target_partname()?.as_str().len();
            let suffix = relationship
                .target_query()
                .map_or(0, |value| 1 + value.len())
                + relationship
                    .target_fragment()
                    .map_or(0, |value| 1 + value.len());
            absolute
                .checked_add(suffix)
                .ok_or_else(|| invalid_topology("internal target length overflows"))
        },
        TargetMode::External => Ok(relationship.target_ref().len()),
    }
}

const fn relationship_mode(mode: TargetMode) -> u8 {
    match mode {
        TargetMode::Internal => 0,
        TargetMode::External => 1,
    }
}

fn workspace_bound(stats: &Stats) -> Result<usize> {
    let part_bytes = stats
        .parts
        .checked_mul(size_of::<PartRecord>())
        .ok_or_else(|| invalid_topology("digest Part workspace overflows"))?;
    let relationship_bytes = stats
        .relationships
        .checked_mul(size_of::<RelationshipRecord>())
        .ok_or_else(|| invalid_topology("digest relationship workspace overflows"))?;
    part_bytes
        .checked_add(relationship_bytes)
        .and_then(|bytes| bytes.checked_add(stats.string_bytes))
        .and_then(|bytes| bytes.checked_add(size_of::<Records>()))
        .ok_or_else(|| invalid_topology("digest workspace overflows"))
}

fn reserve_workspace(
    context: Option<&ExecutionContext>,
    bytes: usize,
) -> Result<Option<Reservation>> {
    let Some(context) = context else {
        return Ok(None);
    };
    let amount =
        u64::try_from(bytes).map_err(|_| invalid_topology("digest workspace exceeds u64"))?;
    context
        .reserve(Resource::Memory, amount)
        .map(Some)
        .map_err(map_execution_error)
}

fn digest_records(mut records: Records, context: Option<&ExecutionContext>) -> Result<[u8; 32]> {
    check_context(context)?;
    records.parts.sort_unstable_by(|left, right| {
        left.name
            .cmp(&right.name)
            .then_with(|| left.content_type.cmp(&right.content_type))
    });
    check_context(context)?;
    records
        .package_relationships
        .sort_unstable_by(compare_relationships);
    check_context(context)?;
    for part in &mut records.parts {
        part.relationships.sort_unstable_by(compare_relationships);
        check_context(context)?;
    }

    let mut digest = Sha256::new();
    digest.update(DIGEST_DOMAIN);
    hash_relationships(&mut digest, &records.package_relationships, context)?;
    hash_len(&mut digest, records.parts.len())?;
    for part in records.parts {
        check_context(context)?;
        hash_bytes(&mut digest, part.name.as_bytes())?;
        hash_bytes(&mut digest, part.content_type.as_bytes())?;
        hash_relationships(&mut digest, &part.relationships, context)?;
    }
    check_context(context)?;
    Ok(digest.finalize().into())
}

fn hash_relationships(
    digest: &mut Sha256,
    relationships: &[RelationshipRecord],
    context: Option<&ExecutionContext>,
) -> Result<()> {
    hash_len(digest, relationships.len())?;
    for relationship in relationships {
        check_context(context)?;
        hash_bytes(digest, relationship.id.as_bytes())?;
        hash_bytes(digest, relationship.reltype.as_bytes())?;
        digest.update([relationship.mode]);
        hash_bytes(digest, relationship.target.as_bytes())?;
    }
    check_context(context)?;
    Ok(())
}

fn compare_relationships(left: &RelationshipRecord, right: &RelationshipRecord) -> Ordering {
    left.id
        .cmp(&right.id)
        .then_with(|| left.reltype.cmp(&right.reltype))
        .then_with(|| left.mode.cmp(&right.mode))
        .then_with(|| left.target.cmp(&right.target))
}

fn hash_bytes(digest: &mut Sha256, bytes: &[u8]) -> Result<()> {
    hash_len(digest, bytes.len())?;
    digest.update(bytes);
    Ok(())
}

fn hash_len(digest: &mut Sha256, length: usize) -> Result<()> {
    let length =
        u64::try_from(length).map_err(|_| invalid_topology("digest length exceeds u64"))?;
    digest.update(length.to_le_bytes());
    Ok(())
}

fn clone_string(value: &str, resource: &'static str) -> Result<String> {
    let mut output = String::new();
    output
        .try_reserve_exact(value.len())
        .map_err(|source| Error::Allocation { resource, source })?;
    output.push_str(value);
    Ok(output)
}

fn add_string_bytes(total: &mut usize, bytes: usize) -> Result<()> {
    *total = total
        .checked_add(bytes)
        .ok_or_else(|| invalid_topology("digest string workspace overflows"))?;
    Ok(())
}

fn check_count(kind: SourceLimitKind, actual: usize, maximum: usize) -> Result<()> {
    if actual > maximum {
        return Err(Error::Opc(OpcError::Limit {
            kind,
            observed: actual as u64,
            maximum: maximum as u64,
        }));
    }
    Ok(())
}

fn check_len(kind: SourceLimitKind, actual: usize, maximum: u64) -> Result<()> {
    if (actual as u64) > maximum {
        return Err(Error::Opc(OpcError::Limit {
            kind,
            observed: actual as u64,
            maximum,
        }));
    }
    Ok(())
}

fn check_context(context: Option<&ExecutionContext>) -> Result<()> {
    context
        .map(ExecutionContext::check)
        .transpose()
        .map_err(map_execution_error)
        .map(|_| ())
}

fn map_execution_error(error: ExecutionError) -> Error {
    match error {
        ExecutionError::Cancelled => Error::Opc(OpcError::Cancelled),
        error => Error::Opc(OpcError::Execution(error)),
    }
}

fn allocation(resource: &'static str, source: std::collections::TryReserveError) -> Error {
    Error::Allocation { resource, source }
}

fn invalid_topology(reason: &'static str) -> Error {
    Error::Invalid(reason.to_owned())
}
