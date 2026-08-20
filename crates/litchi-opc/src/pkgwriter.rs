//! Package writer for OPC packages.
//!
//! This module provides functionality to serialize and write OPC packages to disk,
//! including writing `[Content_Types].xml`, relationships, and all parts.
use crate::constants::content_type as ct;
use crate::content_type::ContentType;
use crate::error::Result;
use crate::package::OpcPackage;
use crate::package::SourceMemberKind;
use crate::packuri::{CONTENT_TYPES_URI, PACKAGE_URI, PackURI};
use crate::phys_pkg::PhysPkgWriter;
use crate::rel::Relationships;
use litchi_core::xml::escape_xml;
use std::collections::{HashMap, HashSet};
use std::fmt::Write as _;
use std::io::Write;
use std::path::Path;

const EXACT_SOURCE_CHUNK_BYTES: usize = 64 * 1024;

struct Counted<'a, W> {
    inner: W,
    written: &'a mut u64,
}

struct Chunked<W> {
    inner: W,
}

impl<W: Write> Write for Chunked<W> {
    fn write(&mut self, bytes: &[u8]) -> std::io::Result<usize> {
        self.inner
            .write(&bytes[..bytes.len().min(EXACT_SOURCE_CHUNK_BYTES)])
    }

    fn flush(&mut self) -> std::io::Result<()> {
        self.inner.flush()
    }
}

impl<W: Write> Write for Counted<'_, W> {
    fn write(&mut self, bytes: &[u8]) -> std::io::Result<usize> {
        let written = self.inner.write(bytes)?;
        *self.written = self.written.saturating_add(written as u64);
        Ok(written)
    }

    fn flush(&mut self) -> std::io::Result<()> {
        self.inner.flush()
    }
}

/// Fully validated package metadata and the stable order used for publication.
///
/// Building this plan is deliberately separate from emission: every fallible
/// serialization and XML audit completes before a sequential sink sees bytes.
struct PublicationPlan<'package> {
    content_types_uri: PackURI,
    content_types_xml: String,
    package_rels_uri: PackURI,
    package_rels_xml: String,
    parts: Vec<PlannedPart<'package>>,
}

struct PlannedPart<'package> {
    partname: &'package PackURI,
    content_type: &'package str,
    blob: &'package [u8],
    authored_xml: bool,
    rels: &'package Relationships,
    relationships: Option<PlannedRelationships>,
}

struct PlannedRelationships {
    uri: PackURI,
    xml: String,
}

enum PlannedAppend<'package> {
    Part(&'package PlannedPart<'package>),
    Relationships(&'package PlannedPart<'package>),
}

impl PlannedAppend<'_> {
    fn owner_name(&self) -> &str {
        match self {
            Self::Part(part) | Self::Relationships(part) => part.partname.as_str(),
        }
    }

    fn member_name(&self) -> Option<&str> {
        match self {
            Self::Part(part) => Some(part.partname.membername()),
            Self::Relationships(part) => part
                .relationships
                .as_ref()
                .map(|relationships| relationships.uri.membername()),
        }
    }

    fn kind_order(&self) -> u8 {
        match self {
            Self::Part(_) => 0,
            Self::Relationships(_) => 1,
        }
    }
}

impl<'package> PublicationPlan<'package> {
    fn from_package(package: &'package OpcPackage) -> Result<Self> {
        let mut parts = Vec::new();
        parts
            .try_reserve_exact(package.part_count())
            .map_err(|source| crate::OpcError::Allocation {
                resource: "OPC XML publication part plan",
                source,
            })?;
        for part in package.iter_parts() {
            parts.push(PlannedPart {
                partname: part.partname(),
                content_type: part.content_type(),
                blob: part.blob(),
                authored_xml: xml_minifier::audit::package::is_xml_part(
                    part.partname().as_str(),
                    part.content_type(),
                ) && !package.is_exact_source_xml(part),
                rels: part.rels(),
                relationships: None,
            });
        }
        parts.sort_unstable_by(|left, right| left.partname.as_str().cmp(right.partname.as_str()));

        let content_types_uri =
            PackURI::new(CONTENT_TYPES_URI).map_err(crate::OpcError::InvalidPackUri)?;
        let content_types_xml = ContentTypesItem::from_parts(&parts)?.to_xml();
        PackageWriter::validate_authored_xml("[Content_Types].xml", content_types_xml.as_bytes())?;

        let package_uri = PackURI::new(PACKAGE_URI).map_err(crate::OpcError::InvalidPackUri)?;
        let package_rels_uri = package_uri
            .rels_uri()
            .map_err(crate::OpcError::InvalidPackUri)?;
        let package_rels_xml = package.rels().to_xml();
        PackageWriter::validate_authored_xml("_rels/.rels", package_rels_xml.as_bytes())?;

        for part in &mut parts {
            if part.authored_xml {
                PackageWriter::validate_authored_xml(part.partname.as_str(), part.blob)?;
            }
            if !part.rels.is_empty() {
                let uri = part
                    .partname
                    .rels_uri()
                    .map_err(crate::OpcError::InvalidPackUri)?;
                let xml = part.rels.to_xml();
                PackageWriter::validate_authored_xml(uri.as_str(), xml.as_bytes())?;
                part.relationships = Some(PlannedRelationships { uri, xml });
            }
        }

        Ok(Self {
            content_types_uri,
            content_types_xml,
            package_rels_uri,
            package_rels_xml,
            parts,
        })
    }

    fn write<W: Write>(&self, physical: &mut PhysPkgWriter<W>) -> Result<()> {
        physical.write(&self.content_types_uri, self.content_types_xml.as_bytes())?;
        physical.write(&self.package_rels_uri, self.package_rels_xml.as_bytes())?;
        for part in &self.parts {
            physical.write(part.partname, part.blob)?;
            if let Some(relationships) = &part.relationships {
                physical.write(&relationships.uri, relationships.xml.as_bytes())?;
            }
        }
        Ok(())
    }
}

enum PreservationWrite<W> {
    Written(W),
    Fallback(W),
}

fn try_write_preserved<W: Write>(
    writer: W,
    package: &OpcPackage,
    publication: &PublicationPlan<'_>,
) -> Result<PreservationWrite<W>> {
    let Some((source, provenance)) = package.preservation_source() else {
        return Ok(PreservationWrite::Fallback(writer));
    };
    let Ok(archive) = soapberry_zip::ZipArchive::from_slice(source) else {
        return Ok(PreservationWrite::Fallback(writer));
    };
    let archive = archive.into_zip_archive();
    let mut buffer = Vec::new();
    buffer
        .try_reserve_exact(soapberry_zip::RECOMMENDED_BUFFER_SIZE)
        .map_err(|source| crate::OpcError::Allocation {
            resource: "OPC ZIP preservation index",
            source,
        })?;
    buffer.resize(soapberry_zip::RECOMMENDED_BUFFER_SIZE, 0_u8);
    let Ok(index) = soapberry_zip::PreservationIndex::new(&archive, &mut buffer) else {
        return Ok(PreservationWrite::Fallback(writer));
    };
    if index.archive_end_offset() != source.len() as u64 {
        // Preserve only a complete source archive.  Otherwise a successful
        // preservation write would silently discard bytes after the EOCD.
        return Ok(PreservationWrite::Fallback(writer));
    }
    if index.entries().len() != provenance.members.len() {
        return Ok(PreservationWrite::Fallback(writer));
    }

    let mut planned_parts = HashMap::new();
    planned_parts
        .try_reserve(publication.parts.len())
        .map_err(|source| crate::OpcError::Allocation {
            resource: "OPC targeted publication part lookup",
            source,
        })?;
    let mut additions = Vec::new();
    let mut relationship_additions = Vec::new();
    additions
        .try_reserve(publication.parts.len())
        .map_err(|source| crate::OpcError::Allocation {
            resource: "OPC topology-add publication parts",
            source,
        })?;
    relationship_additions
        .try_reserve(publication.parts.len())
        .map_err(|source| crate::OpcError::Allocation {
            resource: "OPC topology-add relationship members",
            source,
        })?;
    for part in &publication.parts {
        if let Some(source_part) = provenance.parts.get(part.partname) {
            if !source_part.member_present {
                return Ok(PreservationWrite::Fallback(writer));
            }
            if !source_part.relationships_member_present && part.relationships.is_some() {
                relationship_additions.push(part);
            }
            planned_parts.insert(part.partname, part);
        } else {
            additions.push(part);
        }
    }
    additions.sort_unstable_by(|left, right| left.partname.as_str().cmp(right.partname.as_str()));
    relationship_additions
        .sort_unstable_by(|left, right| left.partname.as_str().cmp(right.partname.as_str()));
    let mut omitted_ids = HashSet::new();
    omitted_ids
        .try_reserve(index.entries().len())
        .map_err(|source| crate::OpcError::Allocation {
            resource: "OPC omitted preservation members",
            source,
        })?;
    for (source_member, indexed_entry) in provenance.members.iter().zip(index.entries()) {
        let omitted = match &source_member.kind {
            SourceMemberKind::Part(partname) => !planned_parts.contains_key(partname),
            SourceMemberKind::PartRelationships(partname) => planned_parts
                .get(partname)
                .is_none_or(|part| part.relationships.is_none()),
            SourceMemberKind::ContentTypes
            | SourceMemberKind::PackageRelationships
            | SourceMemberKind::Unknown => false,
        };
        if omitted {
            omitted_ids.insert(indexed_entry.id());
        }
    }
    if !omitted_ids.is_empty() && !omitted_members_form_suffix(&index, &omitted_ids)? {
        return Ok(PreservationWrite::Fallback(writer));
    }
    let topology_add = !additions.is_empty() || !relationship_additions.is_empty();
    let append_capacity = additions
        .len()
        .checked_mul(2)
        .and_then(|count| count.checked_add(relationship_additions.len()))
        .ok_or_else(|| crate::OpcError::ZipError("OPC append member count overflow".into()))?;
    let mut appended = Vec::new();
    appended
        .try_reserve_exact(append_capacity)
        .map_err(|source| crate::OpcError::Allocation {
            resource: "OPC appended preservation members",
            source,
        })?;
    for part in &additions {
        appended.push(PlannedAppend::Part(part));
        if part.relationships.is_some() {
            appended.push(PlannedAppend::Relationships(part));
        }
    }
    for part in &relationship_additions {
        appended.push(PlannedAppend::Relationships(part));
    }
    appended.sort_unstable_by(|left, right| {
        left.owner_name()
            .cmp(right.owner_name())
            .then_with(|| left.kind_order().cmp(&right.kind_order()))
    });
    if topology_add
        && provenance.members.iter().any(|member| {
            member.name.is_none() || matches!(&member.kind, SourceMemberKind::Unknown)
        })
    {
        // Appending generated entries after a source member whose identity is
        // not modeled would silently change an opaque archive topology. The
        // caller turns this fallback into a typed capability error for owned
        // sources; only new or borrowed packages may use the full writer.
        return Ok(PreservationWrite::Fallback(writer));
    }
    if topology_add {
        let mut member_names = HashSet::new();
        member_names
            .try_reserve(provenance.members.len())
            .map_err(|source| crate::OpcError::Allocation {
                resource: "OPC source preservation member names",
                source,
            })?;
        for member in &provenance.members {
            let Some(name) = member.name.as_deref() else {
                return Ok(PreservationWrite::Fallback(writer));
            };
            member_names.insert(normalized_member_name(
                name,
                "OPC source preservation member name",
            )?);
        }
        let mut appended_names = HashSet::new();
        appended_names
            .try_reserve(appended.len())
            .map_err(|source| crate::OpcError::Allocation {
                resource: "OPC appended preservation member names",
                source,
            })?;
        for append in &appended {
            let Some(name) = append.member_name() else {
                return Ok(PreservationWrite::Fallback(writer));
            };
            if name.is_empty() {
                return Ok(PreservationWrite::Fallback(writer));
            }
            let normalized = normalized_member_name(name, "OPC appended preservation member name")?;
            if member_names.contains(&normalized) || !appended_names.insert(normalized) {
                return Ok(PreservationWrite::Fallback(writer));
            }
        }
    }

    let package_relationship_members = provenance
        .members
        .iter()
        .filter(|member| matches!(member.kind, SourceMemberKind::PackageRelationships))
        .count();
    if package_relationship_members != 1 {
        return Ok(PreservationWrite::Fallback(writer));
    }

    let removed_part = provenance
        .members
        .iter()
        .zip(index.entries())
        .any(|(member, entry)| {
            omitted_ids.contains(&entry.id()) && matches!(member.kind, SourceMemberKind::Part(_))
        });
    let content_types_changed = removed_part
        || publication.parts.iter().any(|part| {
            provenance
                .parts
                .get(part.partname)
                .is_none_or(|source| source.content_type != part.content_type)
        });
    let mut regenerated_bytes = 0_u64;
    let mut regenerated_members = 0_u64;
    let omitted_members = u64::try_from(omitted_ids.len())
        .map_err(|_| crate::OpcError::ZipError("OPC omitted member count overflow".into()))?;
    let mut appended_members = 0_u64;
    let mut appended_bytes = 0_u64;
    for (member, indexed_entry) in provenance.members.iter().zip(index.entries()) {
        if omitted_ids.contains(&indexed_entry.id()) {
            continue;
        }
        let bytes = match &member.kind {
            SourceMemberKind::ContentTypes if content_types_changed => {
                Some(publication.content_types_xml.len())
            },
            SourceMemberKind::PackageRelationships
                if provenance.package_relationships_xml != publication.package_rels_xml =>
            {
                Some(publication.package_rels_xml.len())
            },
            SourceMemberKind::Part(partname) => {
                let Some(part) = planned_parts.get(partname) else {
                    return Ok(PreservationWrite::Fallback(writer));
                };
                let Some(source_part) = provenance.parts.get(partname) else {
                    return Ok(PreservationWrite::Fallback(writer));
                };
                (source_part.blob.as_slice() != part.blob).then_some(part.blob.len())
            },
            SourceMemberKind::PartRelationships(partname) => {
                let Some(part) = planned_parts.get(partname) else {
                    return Ok(PreservationWrite::Fallback(writer));
                };
                let Some(relationships) = part.relationships.as_ref() else {
                    return Ok(PreservationWrite::Fallback(writer));
                };
                let Some(source_part) = provenance.parts.get(partname) else {
                    return Ok(PreservationWrite::Fallback(writer));
                };
                (source_part.relationships_xml != relationships.xml)
                    .then_some(relationships.xml.len())
            },
            SourceMemberKind::ContentTypes
            | SourceMemberKind::PackageRelationships
            | SourceMemberKind::Unknown => None,
        };
        if let Some(bytes) = bytes {
            if member.name.is_none() {
                return Ok(PreservationWrite::Fallback(writer));
            }
            regenerated_bytes = match regenerated_bytes.checked_add(bytes as u64) {
                Some(total) => total,
                None => return Ok(PreservationWrite::Fallback(writer)),
            };
            regenerated_members += 1;
        }
    }
    for append in &appended {
        appended_members = appended_members.checked_add(1).ok_or_else(|| {
            crate::OpcError::ZipError("OPC appended member count overflow".into())
        })?;
        let bytes = match append {
            PlannedAppend::Part(part) => part.blob.len(),
            PlannedAppend::Relationships(part) => {
                let Some(relationships) = part.relationships.as_ref() else {
                    return Ok(PreservationWrite::Fallback(writer));
                };
                relationships.xml.len()
            },
        };
        appended_bytes = appended_bytes.checked_add(bytes as u64).ok_or_else(|| {
            crate::OpcError::ZipError("OPC appended member bytes overflow".into())
        })?;
    }
    let conservative_output_bound = (source.len() as u64)
        .checked_add(regenerated_bytes.saturating_mul(2))
        .and_then(|size| size.checked_add(regenerated_members.saturating_mul(64 * 1024)))
        .and_then(|size| size.checked_add(appended_bytes.saturating_mul(2)))
        .and_then(|size| size.checked_add(appended_members.saturating_mul(64 * 1024)));
    if conservative_output_bound.is_none_or(|size| size > u64::from(u32::MAX))
        || !output_entry_count_is_zip32_safe(
            provenance.members.len(),
            omitted_members,
            appended_members,
        )
    {
        return Ok(PreservationWrite::Fallback(writer));
    }
    let mut plan = soapberry_zip::PreservationPlan::new();
    plan.try_reserve_exact(index.entries().len())
        .map_err(|source| crate::OpcError::Allocation {
            resource: "OPC preservation actions",
            source,
        })?;
    plan.try_reserve_appended(usize::try_from(appended_members).map_err(|_| {
        crate::OpcError::ZipError("OPC appended member count exceeds platform limits".into())
    })?)
    .map_err(|source| crate::OpcError::Allocation {
        resource: "OPC appended preservation members",
        source,
    })?;
    for (source_member, indexed_entry) in provenance.members.iter().zip(index.entries()) {
        let action = if omitted_ids.contains(&indexed_entry.id()) {
            soapberry_zip::PreservationAction::Omit(indexed_entry.id())
        } else {
            match &source_member.kind {
                SourceMemberKind::ContentTypes if content_types_changed => regenerated_action(
                    indexed_entry.id(),
                    source_member.name.as_deref(),
                    publication.content_types_xml.as_bytes(),
                )?,
                SourceMemberKind::PackageRelationships
                    if provenance.package_relationships_xml != publication.package_rels_xml =>
                {
                    regenerated_action(
                        indexed_entry.id(),
                        source_member.name.as_deref(),
                        publication.package_rels_xml.as_bytes(),
                    )?
                },
                SourceMemberKind::Part(partname) => {
                    let Some(part) = planned_parts.get(partname) else {
                        return Ok(PreservationWrite::Fallback(writer));
                    };
                    let Some(source_part) = provenance.parts.get(partname) else {
                        return Ok(PreservationWrite::Fallback(writer));
                    };
                    if source_part.blob.as_slice() == part.blob {
                        soapberry_zip::PreservationAction::Copy(indexed_entry.id())
                    } else {
                        regenerated_shared_action(
                            indexed_entry.id(),
                            source_member.name.as_deref(),
                            package.get_part(partname)?.blob_arc(),
                        )?
                    }
                },
                SourceMemberKind::PartRelationships(partname) => {
                    let Some(part) = planned_parts.get(partname) else {
                        return Ok(PreservationWrite::Fallback(writer));
                    };
                    let Some(relationships) = &part.relationships else {
                        return Ok(PreservationWrite::Fallback(writer));
                    };
                    let Some(source_part) = provenance.parts.get(partname) else {
                        return Ok(PreservationWrite::Fallback(writer));
                    };
                    if source_part.relationships_xml == relationships.xml {
                        soapberry_zip::PreservationAction::Copy(indexed_entry.id())
                    } else {
                        regenerated_action(
                            indexed_entry.id(),
                            source_member.name.as_deref(),
                            relationships.xml.as_bytes(),
                        )?
                    }
                },
                SourceMemberKind::ContentTypes
                | SourceMemberKind::PackageRelationships
                | SourceMemberKind::Unknown => {
                    soapberry_zip::PreservationAction::Copy(indexed_entry.id())
                },
            }
        };
        plan.push(action);
    }
    for append in &appended {
        let entry = match append {
            PlannedAppend::Part(part) => regenerated_shared_entry(
                part.partname.membername(),
                package.get_part(part.partname)?.blob_arc(),
            )?,
            PlannedAppend::Relationships(part) => {
                let Some(relationships) = part.relationships.as_ref() else {
                    return Ok(PreservationWrite::Fallback(writer));
                };
                regenerated_entry(
                    relationships.uri.membername(),
                    relationships.xml.as_bytes(),
                    "OPC appended relationship payload",
                )?
            },
        };
        plan.try_append(entry)
            .map_err(|source| crate::OpcError::Allocation {
                resource: "OPC appended preservation members",
                source,
            })?;
    }

    index
        .write_to(&plan, Chunked { inner: writer })
        .map(|writer| PreservationWrite::Written(writer.inner))
        .map_err(|error| crate::OpcError::ZipError(error.to_string()))
}

fn regenerated_action(
    id: soapberry_zip::PreservationEntryId,
    name: Option<&str>,
    bytes: &[u8],
) -> Result<soapberry_zip::PreservationAction> {
    let owned_name = regenerated_name(name)?;
    let mut data = Vec::new();
    data.try_reserve_exact(bytes.len())
        .map_err(|source| crate::OpcError::Allocation {
            resource: "OPC targeted member payload",
            source,
        })?;
    data.extend_from_slice(bytes);
    Ok(soapberry_zip::PreservationAction::Regenerate {
        id,
        entry: soapberry_zip::RegeneratedEntry::new(owned_name, data)
            .compression_method(soapberry_zip::CompressionMethod::Deflate),
    })
}

fn regenerated_shared_action(
    id: soapberry_zip::PreservationEntryId,
    name: Option<&str>,
    data: std::sync::Arc<Vec<u8>>,
) -> Result<soapberry_zip::PreservationAction> {
    let owned_name = regenerated_name(name)?;
    Ok(soapberry_zip::PreservationAction::Regenerate {
        id,
        entry: soapberry_zip::RegeneratedEntry::new_shared(owned_name, data)
            .compression_method(soapberry_zip::CompressionMethod::Deflate),
    })
}

fn regenerated_entry(
    name: &str,
    bytes: &[u8],
    resource: &'static str,
) -> Result<soapberry_zip::RegeneratedEntry> {
    let mut data = Vec::new();
    data.try_reserve_exact(bytes.len())
        .map_err(|source| crate::OpcError::Allocation { resource, source })?;
    data.extend_from_slice(bytes);
    Ok(
        soapberry_zip::RegeneratedEntry::new(regenerated_name(Some(name))?, data)
            .compression_method(soapberry_zip::CompressionMethod::Deflate),
    )
}

fn regenerated_shared_entry(
    name: &str,
    data: std::sync::Arc<Vec<u8>>,
) -> Result<soapberry_zip::RegeneratedEntry> {
    Ok(
        soapberry_zip::RegeneratedEntry::new_shared(regenerated_name(Some(name))?, data)
            .compression_method(soapberry_zip::CompressionMethod::Deflate),
    )
}

fn normalized_member_name(name: &str, resource: &'static str) -> Result<String> {
    let mut bytes = Vec::new();
    bytes
        .try_reserve_exact(name.len())
        .map_err(|source| crate::OpcError::Allocation { resource, source })?;
    bytes.extend(name.as_bytes().iter().map(u8::to_ascii_lowercase));
    String::from_utf8(bytes).map_err(|_| {
        crate::OpcError::ZipError("OPC preservation member name normalization failed".into())
    })
}

fn omitted_members_form_suffix<R: soapberry_zip::ReaderAt>(
    index: &soapberry_zip::PreservationIndex<'_, R>,
    omitted: &HashSet<soapberry_zip::PreservationEntryId>,
) -> Result<bool> {
    let entries = index.entries();
    let Some(suffix_start) = entries.len().checked_sub(omitted.len()) else {
        return Ok(false);
    };
    for (position, entry) in entries.iter().enumerate() {
        if omitted.contains(&entry.id()) != (position >= suffix_start) {
            return Ok(false);
        }
    }

    let mut local_order = Vec::new();
    local_order
        .try_reserve_exact(entries.len())
        .map_err(|source| crate::OpcError::Allocation {
            resource: "OPC preservation local omission order",
            source,
        })?;
    local_order.extend(0..entries.len());
    local_order.sort_unstable_by_key(|&position| entries[position].local_span().start);
    for (position, entry_position) in local_order.iter().enumerate() {
        if omitted.contains(&entries[*entry_position].id()) != (position >= suffix_start) {
            return Ok(false);
        }
    }
    Ok(true)
}

fn output_entry_count_is_zip32_safe(
    source_members: usize,
    omitted_members: u64,
    appended_members: u64,
) -> bool {
    u64::try_from(source_members)
        .ok()
        .and_then(|count| count.checked_sub(omitted_members))
        .and_then(|count| count.checked_add(appended_members))
        .is_some_and(|count| count < u64::from(u16::MAX))
}

fn regenerated_name(name: Option<&str>) -> Result<String> {
    let Some(name) = name else {
        return Err(crate::OpcError::ZipError(
            "targeted OPC member has no preservable UTF-8 name".to_owned(),
        ));
    };
    let mut owned_name = String::new();
    owned_name
        .try_reserve_exact(name.len())
        .map_err(|source| crate::OpcError::Allocation {
            resource: "OPC targeted member name",
            source,
        })?;
    owned_name.push_str(name);
    Ok(owned_name)
}

/// Package writer that serializes an OPC package to a ZIP file.
///
/// This is the main entry point for saving packages. It handles writing:
/// - `[Content_Types].xml`
/// - _rels/.rels (package relationships)
/// - All parts and their relationships
///
/// # Example
///
/// ```no_run
/// use litchi_opc::package::OpcPackage;
/// use litchi_opc::pkgwriter::PackageWriter;
///
/// let mut pkg = OpcPackage::new();
/// // ... add parts to package ...
/// PackageWriter::write("output.docx", &pkg)?;
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
pub struct PackageWriter;

impl PackageWriter {
    /// Atomically write an OPC package to a file.
    ///
    /// # Arguments
    /// * `path` - Path where the package should be written
    /// * `package` - The OPC package to write
    ///
    /// # Errors
    /// Returns an error if the package cannot be serialized (for example an
    /// invalid content type or partname) or if writing to the filesystem fails.
    pub fn write<P: AsRef<Path>>(path: P, package: &OpcPackage) -> Result<()> {
        crate::atomic::replace(path.as_ref(), |writer| {
            Self::write_to_stream(writer, package)
        })
    }

    /// Write an OPC package directly to a sequential stream.
    ///
    /// On failure after output begins, [`crate::OpcError::IncompleteOutput`]
    /// reports how many bytes the sink accepted. Seeking is not required.
    ///
    /// # Arguments
    /// * `writer` - A writer that implements Write
    /// * `package` - The OPC package to write
    ///
    /// # Errors
    /// Returns an error if the package cannot be serialized (for example an
    /// invalid content type or partname) or if the sink rejects a write. When
    /// the sink has already accepted bytes, the error is wrapped in
    /// [`crate::OpcError::IncompleteOutput`] with the accepted byte count.
    pub fn write_to_stream<W: Write>(writer: W, package: &OpcPackage) -> Result<()> {
        let mut written = 0_u64;
        let result = Self::write_counted(
            Counted {
                inner: writer,
                written: &mut written,
            },
            package,
        );
        match result {
            Err(source) if written != 0 => Err(crate::OpcError::IncompleteOutput {
                written,
                source: Box::new(source),
            }),
            other => other,
        }
    }

    fn write_counted<W: Write>(writer: W, package: &OpcPackage) -> Result<()> {
        if let Some(source) = package.exact_source() {
            let mut writer = writer;
            for chunk in source.chunks(EXACT_SOURCE_CHUNK_BYTES) {
                writer.write_all(chunk)?;
            }
            writer.flush()?;
            return Ok(());
        }
        Self::validate_source_publication(package)?;
        let plan = PublicationPlan::from_package(package)?;
        let writer = match try_write_preserved(writer, package, &plan)? {
            PreservationWrite::Written(_writer) => return Ok(()),
            PreservationWrite::Fallback(writer) => {
                if package.requires_owned_source_preservation() {
                    return Err(owned_source_preservation_error());
                }
                writer
            },
        };
        let mut physical = PhysPkgWriter::with_writer(writer);
        plan.write(&mut physical)?;
        let mut finished = physical.finish_into_inner()?;
        finished.flush()?;
        Ok(())
    }

    /// Serialize an OPC package to bytes.
    ///
    /// # Arguments
    /// * `package` - The OPC package to serialize
    ///
    /// # Returns
    /// The serialized package as a byte vector
    ///
    /// # Errors
    /// Returns an error if the package cannot be serialized (for example an
    /// invalid content type or partname) or if the in-memory zip writer fails.
    pub fn to_bytes(package: &OpcPackage) -> Result<Vec<u8>> {
        if let Some(source) = package.exact_source() {
            let mut bytes = Vec::new();
            bytes.try_reserve_exact(source.len()).map_err(|source| {
                crate::OpcError::Allocation {
                    resource: "OPC exact source copy",
                    source,
                }
            })?;
            bytes.extend_from_slice(source);
            return Ok(bytes);
        }
        Self::validate_source_publication(package)?;
        let plan = PublicationPlan::from_package(package)?;
        match try_write_preserved(Vec::new(), package, &plan)? {
            PreservationWrite::Written(bytes) => return Ok(bytes),
            PreservationWrite::Fallback(bytes) => {
                if package.requires_owned_source_preservation() {
                    return Err(owned_source_preservation_error());
                }
                debug_assert!(bytes.is_empty());
            },
        }
        let mut physical = PhysPkgWriter::new();
        plan.write(&mut physical)?;
        physical.finish()
    }

    fn validate_authored_xml(name: &str, bytes: &[u8]) -> Result<()> {
        xml_minifier::audit::verify_authored(bytes, xml_minifier::audit::Limits::default())
            .map(|_report| ())
            .map_err(|source| crate::OpcError::XmlPublication {
                part: name.to_string(),
                source,
            })
    }

    fn validate_source_publication(package: &OpcPackage) -> Result<()> {
        if package.requires_signature_edit_policy() {
            return Err(crate::OpcError::SignedSourceRequiresExplicitPolicy);
        }
        Ok(())
    }
}

fn owned_source_preservation_error() -> crate::OpcError {
    crate::OpcError::PreservationUnavailable {
        reason: "source ZIP framing or opaque members cannot be preserved after this mutation"
            .to_owned(),
    }
}

/// Helper for building `[Content_Types].xml` content.
///
/// Manages Default and Override elements for content type mapping.
struct ContentTypesItem {
    /// Default content types by extension
    defaults: HashMap<String, ContentType>,

    /// Override content types by partname
    overrides: HashMap<String, ContentType>,
}

impl ContentTypesItem {
    /// Create a new `ContentTypesItem`.
    fn new() -> Result<Self> {
        let mut defaults = HashMap::new();

        // Add standard defaults
        defaults.insert("rels".to_string(), ContentType::new(ct::OPC_RELATIONSHIPS)?);
        defaults.insert("xml".to_string(), ContentType::new(ct::XML)?);

        Ok(Self {
            defaults,
            overrides: HashMap::new(),
        })
    }

    /// Build `ContentTypesItem` from the sorted publication parts.
    fn from_parts(parts: &[PlannedPart<'_>]) -> Result<Self> {
        let mut cti = Self::new()?;

        for part in parts {
            cti.add_content_type(part.partname, part.content_type)?;
        }

        Ok(cti)
    }

    /// Add a content type for a part.
    ///
    /// Uses a default mapping if the extension matches a well-known type,
    /// otherwise uses an override for the specific partname.
    fn add_content_type(&mut self, partname: &PackURI, content_type: &str) -> Result<()> {
        let ext = partname.ext().to_ascii_lowercase();
        let parsed_content_type = ContentType::new(content_type)?;

        // Check if this is a standard default mapping
        if Self::is_default_content_type(&ext, parsed_content_type.as_str()) {
            self.defaults.insert(ext, parsed_content_type);
        } else {
            self.overrides
                .insert(partname.to_string(), parsed_content_type);
        }
        Ok(())
    }

    /// Check if an extension/content-type pair is a standard default.
    fn is_default_content_type(ext: &str, content_type: &str) -> bool {
        matches!(
            (ext, content_type),
            ("rels", ct::OPC_RELATIONSHIPS)
                | ("xml", ct::XML)
                | ("bin", ct::XLSB_BIN)
                | ("png", "image/png")
                | ("jpg" | "jpeg", "image/jpeg")
                | ("gif", "image/gif")
                | ("emf", "image/x-emf")
                | ("wmf", "image/x-wmf")
                | (
                    "odttf",
                    "application/vnd.openxmlformats-officedocument.obfuscatedFont"
                )
        )
    }

    /// Generate the XML for `[Content_Types].xml`.
    fn to_xml(&self) -> String {
        let mut xml = String::with_capacity(4096);

        xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
        xml.push_str(
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">"#,
        );

        // Write Default elements (sorted by extension)
        let mut exts: Vec<_> = self.defaults.keys().collect();
        exts.sort();
        for ext in exts {
            let content_type = &self.defaults[ext];
            let _ignored = write!(
                xml,
                r#"<Default Extension="{}" ContentType="{}"/>"#,
                escape_xml(ext),
                escape_xml(content_type.as_str())
            );
        }

        // Write Override elements (sorted by partname)
        let mut partnames: Vec<_> = self.overrides.keys().collect();
        partnames.sort();
        for partname in partnames {
            let content_type = &self.overrides[partname];
            let _ignored = write!(
                xml,
                r#"<Override PartName="{}" ContentType="{}"/>"#,
                escape_xml(partname),
                escape_xml(content_type.as_str())
            );
        }

        xml.push_str("</Types>");

        xml
    }
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::unwrap_used,
        clippy::expect_used,
        reason = "test assertions panic on failure by design"
    )]
    use std::io;

    use super::*;

    struct ChunkSink {
        total: usize,
        writes: usize,
        largest: usize,
        limit: usize,
    }

    struct FailAfter {
        written: usize,
        limit: usize,
    }

    impl Write for FailAfter {
        fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
            let available = self.limit.saturating_sub(self.written);
            if available == 0 {
                return Err(io::Error::new(
                    io::ErrorKind::BrokenPipe,
                    "injected sink failure",
                ));
            }
            let accepted = available.min(bytes.len());
            self.written += accepted;
            Ok(accepted)
        }

        fn flush(&mut self) -> io::Result<()> {
            Ok(())
        }
    }

    impl Write for ChunkSink {
        fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
            if bytes.len() > self.limit {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidData,
                    "writer received an archive-sized chunk",
                ));
            }
            self.total = self.total.saturating_add(bytes.len());
            self.writes = self.writes.saturating_add(1);
            self.largest = self.largest.max(bytes.len());
            Ok(bytes.len())
        }

        fn flush(&mut self) -> io::Result<()> {
            Ok(())
        }
    }

    fn with_eocd_comment(mut archive: Vec<u8>, comment: &[u8]) -> Vec<u8> {
        let comment_len = u16::try_from(comment.len()).expect("ZIP comment fits in EOCD");
        let eocd = archive.len().checked_sub(22).expect("archive has an EOCD");
        assert_eq!(&archive[eocd..eocd + 4], b"PK\x05\x06");
        archive[eocd + 20..eocd + 22].copy_from_slice(&comment_len.to_le_bytes());
        archive.extend_from_slice(comment);
        archive
    }

    fn exact_empty_archive(comment: &[u8]) -> Vec<u8> {
        with_eocd_comment(
            PackageWriter::to_bytes(&OpcPackage::new()).expect("serialize source package"),
            comment,
        )
    }

    struct RawArchive {
        central_order: Vec<String>,
        local_order: Vec<String>,
        local_members: HashMap<String, Vec<u8>>,
        central_records: HashMap<String, Vec<u8>>,
        comment: Vec<u8>,
    }

    fn raw_archive(data: &[u8]) -> RawArchive {
        let archive = soapberry_zip::ZipArchive::from_slice(data).expect("parse ZIP");
        let comment = archive.comment().as_bytes().to_vec();
        let central_order: Vec<String> = archive
            .entries()
            .map(|entry| {
                let entry = entry.expect("central entry");
                std::str::from_utf8(entry.file_path().as_ref())
                    .expect("UTF-8 member name")
                    .to_owned()
            })
            .collect();
        let indexed = archive.into_zip_archive();
        let mut buffer = vec![0_u8; soapberry_zip::RECOMMENDED_BUFFER_SIZE];
        let index =
            soapberry_zip::PreservationIndex::new(&indexed, &mut buffer).expect("preservable ZIP");
        let mut local_members = HashMap::new();
        let mut central_records = HashMap::new();
        let mut records = Vec::new();
        for (name, entry) in central_order.iter().zip(index.entries()) {
            let local = entry.local_span();
            let central = entry.central_record();
            records.push((local.start, name.clone()));
            local_members.insert(
                name.clone(),
                data[local.start as usize..local.end as usize].to_vec(),
            );
            central_records.insert(
                name.clone(),
                data[central.start as usize..central.end as usize].to_vec(),
            );
        }
        records.sort_unstable_by_key(|(offset, _)| *offset);
        RawArchive {
            central_order,
            local_order: records.into_iter().map(|(_, name)| name).collect(),
            local_members,
            central_records,
            comment,
        }
    }

    fn central_without_local_offset(record: &[u8]) -> Vec<u8> {
        let mut record = record.to_vec();
        record[42..46].fill(0);
        record
    }

    #[test]
    fn targeted_output_entry_count_rejects_zip32_sentinel() {
        assert!(output_entry_count_is_zip32_safe(
            usize::from(u16::MAX - 1),
            0,
            0
        ));
        assert!(!output_entry_count_is_zip32_safe(
            usize::from(u16::MAX),
            0,
            0
        ));
        assert!(!output_entry_count_is_zip32_safe(
            usize::from(u16::MAX - 1),
            0,
            1
        ));
        assert!(output_entry_count_is_zip32_safe(
            usize::from(u16::MAX),
            1,
            0
        ));
    }

    fn pseudo_random_bytes(len: usize, mut state: u32) -> Vec<u8> {
        let mut bytes = Vec::with_capacity(len);
        while bytes.len() < len {
            state ^= state << 13;
            state ^= state >> 17;
            state ^= state << 5;
            bytes.push((state >> 24) as u8);
        }
        bytes
    }

    fn two_part_source(comment: &[u8]) -> (Vec<u8>, PackURI, PackURI) {
        let first = PackURI::new("/custom/first.bin").expect("first URI");
        let second = PackURI::new("/custom/second.bin").expect("second URI");
        let mut first_part = crate::BlobPart::new(
            first.clone(),
            "application/octet-stream".to_owned(),
            pseudo_random_bytes(256 * 1024, 0x1234_5678),
        );
        crate::Part::relate_to_ext(&mut first_part, "https://example.com/old", "urn:test");
        let mut package = OpcPackage::new();
        package.add_part(Box::new(first_part));
        package.add_part(Box::new(crate::BlobPart::new(
            second.clone(),
            "application/octet-stream".to_owned(),
            pseudo_random_bytes(128 * 1024, 0x8765_4321),
        )));
        (
            with_eocd_comment(
                PackageWriter::to_bytes(&package).expect("serialize two-part source"),
                comment,
            ),
            first,
            second,
        )
    }

    fn signed_source() -> (Vec<u8>, PackURI) {
        let first = PackURI::new("/custom/first.bin").expect("first URI");
        let origin = PackURI::new("/_xmlsignatures/origin.sigs").expect("origin URI");
        let mut package = OpcPackage::new();
        package.add_part(Box::new(crate::BlobPart::new(
            first.clone(),
            "application/octet-stream".to_owned(),
            b"signed payload".to_vec(),
        )));
        package.add_part(Box::new(crate::BlobPart::new(
            origin,
            crate::constants::content_type::OPC_DIGITAL_SIGNATURE_ORIGIN.to_owned(),
            b"<origin/>".to_vec(),
        )));
        package.relate_to(
            "_xmlsignatures/origin.sigs",
            crate::constants::relationship_type::DIGITAL_SIGNATURE_ORIGIN,
        );
        (
            PackageWriter::to_bytes(&package).expect("serialize signed source"),
            first,
        )
    }

    fn read_u16(bytes: &[u8], offset: usize) -> u16 {
        u16::from_le_bytes(bytes[offset..offset + 2].try_into().expect("u16 field"))
    }

    fn read_u32(bytes: &[u8], offset: usize) -> u32 {
        u32::from_le_bytes(bytes[offset..offset + 4].try_into().expect("u32 field"))
    }

    fn central_record_len(bytes: &[u8], offset: usize) -> usize {
        46 + usize::from(read_u16(bytes, offset + 28))
            + usize::from(read_u16(bytes, offset + 30))
            + usize::from(read_u16(bytes, offset + 32))
    }

    fn add_extras_to_last_member(mut bytes: Vec<u8>) -> Vec<u8> {
        let archive = soapberry_zip::ZipArchive::from_slice(&bytes).expect("parse ZIP");
        let central = archive.directory_offset() as usize;
        let eocd = archive.eocd_offset() as usize;
        let mut record = central;
        let mut last_local = 0_usize;
        while record < eocd {
            last_local = last_local.max(read_u32(&bytes, record + 42) as usize);
            record += central_record_len(&bytes, record);
        }

        let local_extra = [0xfe, 0xca, 3, 0, b'l', b'o', b'c'];
        let local_name_len = usize::from(read_u16(&bytes, last_local + 26));
        let old_local_extra_len = usize::from(read_u16(&bytes, last_local + 28));
        let local_insert = last_local + 30 + local_name_len + old_local_extra_len;
        bytes[last_local + 28..last_local + 30].copy_from_slice(
            &u16::try_from(old_local_extra_len + local_extra.len())
                .unwrap()
                .to_le_bytes(),
        );
        bytes.splice(local_insert..local_insert, local_extra);

        let shifted_central = central + local_extra.len();
        let shifted_eocd = eocd + local_extra.len();
        bytes[shifted_eocd + 16..shifted_eocd + 20]
            .copy_from_slice(&(shifted_central as u32).to_le_bytes());

        let mut target_record = shifted_central;
        while read_u32(&bytes, target_record + 42) as usize != last_local {
            target_record += central_record_len(&bytes, target_record);
        }
        let central_extra = [0xef, 0xbe, 3, 0, b'c', b'e', b'n'];
        let central_name_len = usize::from(read_u16(&bytes, target_record + 28));
        let old_central_extra_len = usize::from(read_u16(&bytes, target_record + 30));
        let central_insert = target_record + 46 + central_name_len + old_central_extra_len;
        bytes[target_record + 30..target_record + 32].copy_from_slice(
            &u16::try_from(old_central_extra_len + central_extra.len())
                .unwrap()
                .to_le_bytes(),
        );
        bytes.splice(central_insert..central_insert, central_extra);

        let final_eocd = shifted_eocd + central_extra.len();
        let central_size = read_u32(&bytes, final_eocd + 12);
        bytes[final_eocd + 12..final_eocd + 16]
            .copy_from_slice(&(central_size + central_extra.len() as u32).to_le_bytes());
        bytes
    }

    fn reverse_central_order(mut bytes: Vec<u8>) -> Vec<u8> {
        let archive = soapberry_zip::ZipArchive::from_slice(&bytes).expect("parse ZIP");
        let central = archive.directory_offset() as usize;
        let eocd = archive.eocd_offset() as usize;
        let mut records = Vec::new();
        let mut offset = central;
        while offset < eocd {
            let len = central_record_len(&bytes, offset);
            records.push(bytes[offset..offset + len].to_vec());
            offset += len;
        }
        records.reverse();
        let replacement: Vec<u8> = records.into_iter().flatten().collect();
        bytes[central..eocd].copy_from_slice(&replacement);
        bytes
    }

    fn promote_to_zip64(mut bytes: Vec<u8>) -> Vec<u8> {
        let archive = soapberry_zip::ZipArchive::from_slice(&bytes).expect("parse ZIP");
        let eocd = archive.eocd_offset() as usize;
        let entries = u64::from(read_u16(&bytes, eocd + 10));
        let central_size = u64::from(read_u32(&bytes, eocd + 12));
        let central_offset = u64::from(read_u32(&bytes, eocd + 16));
        let mut records = Vec::new();
        records.extend_from_slice(&0x0606_4b50_u32.to_le_bytes());
        records.extend_from_slice(&44_u64.to_le_bytes());
        records.extend_from_slice(&45_u16.to_le_bytes());
        records.extend_from_slice(&45_u16.to_le_bytes());
        records.extend_from_slice(&0_u32.to_le_bytes());
        records.extend_from_slice(&0_u32.to_le_bytes());
        records.extend_from_slice(&entries.to_le_bytes());
        records.extend_from_slice(&entries.to_le_bytes());
        records.extend_from_slice(&central_size.to_le_bytes());
        records.extend_from_slice(&central_offset.to_le_bytes());
        records.extend_from_slice(&0x0706_4b50_u32.to_le_bytes());
        records.extend_from_slice(&0_u32.to_le_bytes());
        records.extend_from_slice(&(eocd as u64).to_le_bytes());
        records.extend_from_slice(&1_u32.to_le_bytes());
        bytes.splice(eocd..eocd, records);
        let ordinary_eocd = eocd + 76;
        bytes[ordinary_eocd + 8..ordinary_eocd + 12].fill(0xff);
        bytes[ordinary_eocd + 12..ordinary_eocd + 20].fill(0xff);
        bytes
    }

    fn source_with_non_part_framing() -> (Vec<u8>, PackURI) {
        let content_types = br#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/custom/first.bin" ContentType="application/octet-stream"/><Override PartName="/custom/second.bin" ContentType="application/octet-stream"/></Types>"#;
        let relationships = br#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>"#;
        let first = PackURI::new("/custom/first.bin").unwrap();
        let mut physical = PhysPkgWriter::new();
        physical
            .write(
                &PackURI::new("/[Content_Types].xml").unwrap(),
                content_types,
            )
            .unwrap();
        physical
            .write(&PackURI::new("/_rels/.rels").unwrap(), relationships)
            .unwrap();
        physical
            .write(&first, &pseudo_random_bytes(32 * 1024, 0x1111_2222))
            .unwrap();
        physical
            .write(
                &PackURI::new("/custom/second.bin").unwrap(),
                &pseudo_random_bytes(16 * 1024, 0x3333_4444),
            )
            .unwrap();
        physical
            .write(
                &PackURI::new("/junk.dat").unwrap(),
                b"untyped non-part payload",
            )
            .unwrap();
        let bytes = physical.finish().unwrap();
        let bytes = add_extras_to_last_member(bytes);
        let bytes = reverse_central_order(bytes);
        (
            with_eocd_comment(bytes, b"archive comment and framing"),
            first,
        )
    }

    #[test]
    fn test_content_types_xml() {
        let mut cti = ContentTypesItem::new().unwrap();
        cti.defaults
            .insert("png".to_string(), ContentType::new("image/png").unwrap());
        cti.overrides.insert(
            "/word/document.xml".to_string(),
            ContentType::new(ct::WML_DOCUMENT_MAIN).unwrap(),
        );

        let xml = cti.to_xml();

        assert!(xml.contains(r#"<Default Extension="png" ContentType="image/png"/>"#));
        assert!(xml.contains(r#"<Override PartName="/word/document.xml""#));
    }

    #[test]
    fn owned_source_round_trips_exactly_to_bytes_and_stream() {
        let source = exact_empty_archive(b"nonzero EOCD comment");
        let package = OpcPackage::from_vec(source.clone()).expect("open owned source");

        assert_eq!(
            PackageWriter::to_bytes(&package).expect("copy exact source"),
            source
        );

        let mut streamed = Vec::new();
        package
            .to_stream(&mut streamed)
            .expect("stream exact source");
        assert_eq!(streamed, source);
    }

    #[test]
    fn changed_owned_signed_source_refuses_before_output() {
        let (source, first) = signed_source();
        let mut package = OpcPackage::from_vec(source.clone()).expect("open signed source");
        assert!(package.is_signed());
        assert_eq!(PackageWriter::to_bytes(&package).unwrap(), source);

        package
            .get_part_mut(&first)
            .expect("signed payload")
            .set_blob(b"changed signed payload".to_vec());

        assert!(matches!(
            PackageWriter::to_bytes(&package),
            Err(crate::OpcError::SignedSourceRequiresExplicitPolicy)
        ));
        let mut output = Vec::new();
        let error = PackageWriter::write_to_stream(&mut output, &package)
            .expect_err("changed signed source must be rejected");
        assert!(matches!(
            error,
            crate::OpcError::SignedSourceRequiresExplicitPolicy
        ));
        assert!(output.is_empty());
    }

    #[test]
    fn generic_signature_infrastructure_removal_stays_policy_gated() {
        let (source, _first) = signed_source();
        let origin = PackURI::new("/_xmlsignatures/origin.sigs").expect("origin URI");
        let mut package = OpcPackage::from_vec(source).expect("open signed source");

        package
            .rels_mut()
            .remove("rId1")
            .expect("signature-origin relationship");
        assert!(
            package.is_signed(),
            "origin part remains signature infrastructure"
        );
        assert!(package.remove_part(&origin));
        assert!(!package.is_signed(), "generic removal emptied the graph");

        let mut output = Vec::new();
        let error = PackageWriter::write_to_stream(&mut output, &package)
            .expect_err("generic removal must not authorize signature stripping");
        assert!(matches!(
            error,
            crate::OpcError::SignedSourceRequiresExplicitPolicy
        ));
        assert!(output.is_empty());
    }

    #[test]
    fn untouched_borrowed_signed_source_refuses_without_exact_bytes() {
        let (source, _first) = signed_source();
        let package = OpcPackage::from_bytes(&source).expect("open borrowed signed source");
        assert!(package.is_signed());
        assert!(package.exact_source().is_none());

        assert!(matches!(
            PackageWriter::to_bytes(&package),
            Err(crate::OpcError::SignedSourceRequiresExplicitPolicy)
        ));
        let mut output = Vec::new();
        let error = PackageWriter::write_to_stream(&mut output, &package)
            .expect_err("borrowed signed source must not normalize signatures");
        assert!(matches!(
            error,
            crate::OpcError::SignedSourceRequiresExplicitPolicy
        ));
        assert!(output.is_empty());
    }

    #[test]
    fn open_and_reader_retain_owned_source_but_borrowed_bytes_do_not() {
        let source = exact_empty_archive(b"owned source");

        let from_reader =
            OpcPackage::from_reader(io::Cursor::new(source.clone())).expect("open reader source");
        assert_eq!(
            PackageWriter::to_bytes(&from_reader).expect("copy reader source"),
            source
        );

        let directory = tempfile::tempdir().expect("temporary directory");
        let path = directory.path().join("source.opc");
        std::fs::write(&path, &source).expect("write source package");
        let opened = OpcPackage::open(path).expect("open file source");
        assert_eq!(
            PackageWriter::to_bytes(&opened).expect("copy file source"),
            source
        );

        let borrowed = OpcPackage::from_bytes(&source).expect("open borrowed source");
        assert!(borrowed.exact_source().is_none());
        assert_ne!(
            PackageWriter::to_bytes(&borrowed).expect("republish borrowed source"),
            source
        );
    }

    #[test]
    fn part_edit_revokes_exact_source_and_republishes_edited_package() {
        let partname = PackURI::new("/custom/metadata.xml").expect("valid part URI");
        let mut source_package = OpcPackage::new();
        source_package.add_part(Box::new(crate::BlobPart::new(
            partname.clone(),
            ct::XML.to_owned(),
            b"<original/>".to_vec(),
        )));
        let source = with_eocd_comment(
            PackageWriter::to_bytes(&source_package).expect("serialize source package"),
            b"exact source",
        );
        let mut package = OpcPackage::from_vec(source.clone()).expect("open owned source");

        package
            .get_part_mut(&partname)
            .expect("editable part")
            .set_blob(b"<edited/>".to_vec());
        let rewritten = PackageWriter::to_bytes(&package).expect("republish edited package");

        assert_ne!(rewritten, source);
        let reopened = OpcPackage::from_bytes(&rewritten).expect("reopen edited package");
        assert_eq!(
            reopened.get_part(&partname).expect("edited part").blob(),
            b"<edited/>"
        );
    }

    #[test]
    fn exact_source_streaming_respects_the_64_kib_chunk_bound() {
        let comment = vec![b'x'; u16::MAX as usize];
        let source = exact_empty_archive(&comment);
        assert!(source.len() > EXACT_SOURCE_CHUNK_BYTES);
        let package = OpcPackage::from_vec(source.clone()).expect("open owned source");
        let mut sink = ChunkSink {
            total: 0,
            writes: 0,
            largest: 0,
            limit: EXACT_SOURCE_CHUNK_BYTES,
        };

        PackageWriter::write_to_stream(&mut sink, &package).expect("stream exact source");

        assert_eq!(sink.total, source.len());
        assert!(sink.writes > 1);
        assert!(sink.largest <= EXACT_SOURCE_CHUNK_BYTES);
    }

    #[test]
    fn incomplete_exact_source_stream_reports_accepted_bytes() {
        let source = exact_empty_archive(b"exact source");
        let package = OpcPackage::from_vec(source).expect("open owned source");
        let sink = FailAfter {
            written: 0,
            limit: 128,
        };

        let error = PackageWriter::write_to_stream(sink, &package)
            .expect_err("bounded sink must reject exact source");

        assert!(matches!(
            error,
            crate::OpcError::IncompleteOutput {
                written: 128,
                source,
            } if matches!(*source, crate::OpcError::IoError(_))
        ));
    }

    #[test]
    fn targeted_part_mutation_raw_copies_every_other_member() {
        let (source, first, _second) = two_part_source(b"preserve targeted comment");
        let source_raw = raw_archive(&source);
        let mut package = OpcPackage::from_vec(source).expect("open owned source");
        package
            .get_part_mut(&first)
            .expect("first part")
            .set_blob(pseudo_random_bytes(256 * 1024, 0x0bad_f00d));

        let output = PackageWriter::to_bytes(&package).expect("targeted publication");
        let output_raw = raw_archive(&output);

        assert_eq!(output_raw.comment, source_raw.comment);
        assert_eq!(output_raw.central_order, source_raw.central_order);
        assert_eq!(output_raw.local_order, source_raw.local_order);
        for (name, source_member) in &source_raw.local_members {
            if name == first.membername() {
                assert_ne!(output_raw.local_members[name], *source_member);
            } else {
                assert_eq!(output_raw.local_members[name], *source_member, "{name}");
                assert_eq!(
                    central_without_local_offset(&output_raw.central_records[name]),
                    central_without_local_offset(&source_raw.central_records[name]),
                    "{name}"
                );
            }
        }
    }

    #[test]
    fn targeted_relationship_and_content_type_changes_regenerate_only_their_closure() {
        let (source, first, second) = two_part_source(b"closure comment");
        let source_raw = raw_archive(&source);

        let mut relationship_edit =
            OpcPackage::from_vec(source.clone()).expect("open relationship source");
        relationship_edit
            .get_part_mut(&first)
            .expect("first part")
            .rels_mut()
            .retarget("rId1", "https://example.com/new".to_owned())
            .expect("retarget relationship");
        let relationship_output =
            PackageWriter::to_bytes(&relationship_edit).expect("publish relationship edit");
        let relationship_raw = raw_archive(&relationship_output);
        let relationships_name = first.rels_uri().unwrap().membername().to_owned();
        for (name, source_member) in &source_raw.local_members {
            if name == &relationships_name {
                assert_ne!(relationship_raw.local_members[name], *source_member);
            } else {
                assert_eq!(
                    relationship_raw.local_members[name], *source_member,
                    "{name}"
                );
            }
        }

        let mut content_type_edit = OpcPackage::from_vec(source).expect("open content-type source");
        content_type_edit
            .get_part_mut(&second)
            .expect("second part")
            .set_content_type("application/vnd.example.changed".to_owned())
            .expect("change content type");
        let content_type_output =
            PackageWriter::to_bytes(&content_type_edit).expect("publish content-type edit");
        let content_type_raw = raw_archive(&content_type_output);
        for (name, source_member) in &source_raw.local_members {
            if name == "[Content_Types].xml" {
                assert_ne!(content_type_raw.local_members[name], *source_member);
            } else {
                assert_eq!(
                    content_type_raw.local_members[name], *source_member,
                    "{name}"
                );
            }
        }
    }

    #[test]
    fn revoked_noop_uses_preservation_copy_all() {
        let (source, _first, _second) = two_part_source(b"copy-all comment");
        let mut package = OpcPackage::from_vec(source.clone()).expect("open owned source");
        let options = package.save_options().clone();
        package.set_save_options(options);
        assert!(package.exact_source().is_none());

        assert_eq!(
            PackageWriter::to_bytes(&package).expect("copy preserved source"),
            source
        );

        let mut failed = OpcPackage::from_vec(source.clone()).expect("open failure source");
        assert!(
            failed
                .get_part_mut(&PackURI::new("/missing.bin").unwrap())
                .is_err()
        );
        assert!(failed.exact_source().is_none());
        assert_eq!(PackageWriter::to_bytes(&failed).unwrap(), source);
    }

    #[test]
    fn targeted_partial_sink_failure_reports_accepted_bytes() {
        let (source, first, _second) = two_part_source(b"partial targeted output");
        let mut package = OpcPackage::from_vec(source).expect("open owned source");
        package
            .get_part_mut(&first)
            .expect("first part")
            .set_blob(pseudo_random_bytes(256 * 1024, 0xfeed_beef));
        let sink = FailAfter {
            written: 0,
            limit: 70_000,
        };

        let error = PackageWriter::write_to_stream(sink, &package)
            .expect_err("sink must reject targeted output");

        assert!(matches!(
            error,
            crate::OpcError::IncompleteOutput {
                written: 70_000,
                ..
            }
        ));
    }

    #[test]
    fn targeted_streaming_keeps_generated_writes_bounded() {
        let (source, first, _second) = two_part_source(b"bounded targeted output");
        let mut package = OpcPackage::from_vec(source).expect("open owned source");
        package
            .get_part_mut(&first)
            .unwrap()
            .set_blob(pseudo_random_bytes(256 * 1024, 0xdead_beef));
        let mut sink = ChunkSink {
            total: 0,
            writes: 0,
            largest: 0,
            limit: EXACT_SOURCE_CHUNK_BYTES,
        };

        PackageWriter::write_to_stream(&mut sink, &package).expect("bounded targeted stream");

        assert!(sink.total > 256 * 1024);
        assert!(sink.writes > 1);
        assert!(sink.largest <= EXACT_SOURCE_CHUNK_BYTES);
    }

    #[test]
    fn targeted_save_preserves_comments_extras_descriptors_order_and_non_parts() {
        let (source, first) = source_with_non_part_framing();
        let source_raw = raw_archive(&source);
        assert!(
            source_raw.local_members["junk.dat"]
                .windows(7)
                .any(|window| window == [0xfe, 0xca, 3, 0, b'l', b'o', b'c'])
        );
        assert!(
            source_raw.central_records["junk.dat"]
                .windows(7)
                .any(|window| window == [0xef, 0xbe, 3, 0, b'c', b'e', b'n'])
        );
        assert!(
            source_raw
                .local_members
                .values()
                .any(|member| member.windows(4).any(|window| window == b"PK\x07\x08"))
        );

        let mut package = OpcPackage::from_vec(source).expect("open framed source");
        assert_eq!(package.non_part_members()[0].name(), "junk.dat");
        package
            .get_part_mut(&first)
            .expect("first part")
            .set_blob(pseudo_random_bytes(32 * 1024, 0x5555_6666));
        let output = PackageWriter::to_bytes(&package).expect("targeted framed save");
        let output_raw = raw_archive(&output);

        assert_eq!(output_raw.comment, source_raw.comment);
        assert_eq!(output_raw.central_order, source_raw.central_order);
        assert_eq!(output_raw.local_order, source_raw.local_order);
        for name in [
            "[Content_Types].xml",
            "_rels/.rels",
            "custom/second.bin",
            "junk.dat",
        ] {
            assert_eq!(
                output_raw.local_members[name],
                source_raw.local_members[name]
            );
            assert_eq!(
                central_without_local_offset(&output_raw.central_records[name]),
                central_without_local_offset(&source_raw.central_records[name])
            );
        }
        let reopened = OpcPackage::from_bytes(&output).expect("reopen framed output");
        assert_eq!(reopened.non_part_members()[0].name(), "junk.dat");
    }

    #[test]
    fn topology_part_add_preserves_source_and_appends_deterministic_members() {
        let (source, _first, _second) = two_part_source(b"topology addition");
        let source_raw = raw_archive(&source);
        let mut added = OpcPackage::from_vec(source.clone()).expect("open add source");
        let first_added = PackURI::new("/custom/zzz.bin").unwrap();
        let second_added = PackURI::new("/custom/aaa.bin").unwrap();
        let mut first_part = crate::BlobPart::new(
            first_added.clone(),
            "application/octet-stream".to_owned(),
            b"third with relationships".to_vec(),
        );
        crate::Part::relate_to_ext(&mut first_part, "https://example.com/new", "urn:new");
        added.add_part(Box::new(first_part));
        added.add_part(Box::new(crate::BlobPart::new(
            second_added.clone(),
            "application/octet-stream".to_owned(),
            b"third without relationships".to_vec(),
        )));

        let added_output = PackageWriter::to_bytes(&added).expect("publish added parts");
        let added_raw = raw_archive(&added_output);
        assert_eq!(added_raw.comment, source_raw.comment);
        assert_eq!(
            &added_raw.local_order[..source_raw.local_order.len()],
            source_raw.local_order.as_slice()
        );
        assert_eq!(
            &added_raw.central_order[..source_raw.central_order.len()],
            source_raw.central_order.as_slice()
        );
        assert_eq!(
            &added_raw.local_order[source_raw.local_order.len()..],
            [
                "custom/aaa.bin",
                "custom/zzz.bin",
                "custom/_rels/zzz.bin.rels"
            ]
        );
        assert_eq!(
            &added_raw.central_order[source_raw.central_order.len()..],
            [
                "custom/aaa.bin",
                "custom/zzz.bin",
                "custom/_rels/zzz.bin.rels"
            ]
        );
        for name in &source_raw.local_order {
            if name == "[Content_Types].xml" {
                assert_ne!(
                    added_raw.local_members[name],
                    source_raw.local_members[name]
                );
            } else {
                assert_eq!(
                    added_raw.local_members[name], source_raw.local_members[name],
                    "{name}"
                );
                assert_eq!(
                    central_without_local_offset(&added_raw.central_records[name]),
                    central_without_local_offset(&source_raw.central_records[name]),
                    "{name}"
                );
            }
        }
        let reopened = OpcPackage::from_bytes(&added_output).expect("reopen added package");
        assert_eq!(
            reopened.get_part(&second_added).unwrap().blob(),
            b"third without relationships"
        );
        let reopened_first = reopened.get_part(&first_added).unwrap();
        assert_eq!(reopened_first.blob(), b"third with relationships");
        assert_eq!(reopened_first.rels().len(), 1);
        assert_eq!(
            reopened_first.rels().get("rId1").unwrap().target_ref(),
            "https://example.com/new"
        );
        assert_eq!(
            reopened.get_part(&second_added).unwrap().content_type(),
            "application/octet-stream"
        );
    }

    #[test]
    fn appended_suffix_removal_restores_exact_source_bytes() {
        let (source, _first, _second) = two_part_source(b"suffix removal");
        let appended_name = PackURI::new("/custom/appended.bin").unwrap();
        let mut appended_part = crate::BlobPart::new(
            appended_name.clone(),
            "application/octet-stream".to_owned(),
            b"temporary appended payload".to_vec(),
        );
        crate::Part::relate_to_ext(
            &mut appended_part,
            "https://example.com/temporary",
            "urn:temporary",
        );

        let mut package = OpcPackage::from_vec(source.clone()).expect("open append source");
        package.add_part(Box::new(appended_part));
        let appended = PackageWriter::to_bytes(&package).expect("publish appended suffix");
        assert_ne!(appended, source);

        let mut reopened = OpcPackage::from_vec(appended).expect("reopen appended suffix");
        assert!(reopened.remove_part(&appended_name));
        let restored = PackageWriter::to_bytes(&reopened).expect("remove appended suffix");

        assert_eq!(restored, source);
    }

    #[test]
    fn non_suffix_part_and_relationship_removals_refuse_normalization() {
        let (source, first, second) = two_part_source(b"topology fallback");

        let mut removed = OpcPackage::from_vec(source.clone()).expect("open remove source");
        assert!(removed.remove_part(&first));
        let removed_error = PackageWriter::to_bytes(&removed).expect_err("reject removed part");
        assert!(matches!(
            removed_error,
            crate::OpcError::PreservationUnavailable { .. }
        ));
        assert!(removed.get_part(&second).is_ok());

        let mut removed_relationship =
            OpcPackage::from_vec(source).expect("open relationship removal source");
        assert!(
            removed_relationship
                .get_part_mut(&first)
                .unwrap()
                .rels_mut()
                .remove("rId1")
                .is_some()
        );
        let removed_relationship_error = PackageWriter::to_bytes(&removed_relationship)
            .expect_err("reject relationship removal");
        assert!(matches!(
            removed_relationship_error,
            crate::OpcError::PreservationUnavailable { .. }
        ));
    }

    #[test]
    fn adding_relationship_member_to_existing_part_preserves_and_appends() {
        let (source, _first, second) = two_part_source(b"relationship presence addition");
        let source_raw = raw_archive(&source);
        let mut package = OpcPackage::from_vec(source.clone()).expect("open relationship source");
        package
            .get_part_mut(&second)
            .unwrap()
            .relate_to_ext("https://example.com", "urn:new");
        let output = PackageWriter::to_bytes(&package).expect("publish relationship addition");
        let output_raw = raw_archive(&output);
        assert_eq!(output_raw.comment, source_raw.comment);
        assert_eq!(
            &output_raw.local_order[..source_raw.local_order.len()],
            source_raw.local_order.as_slice()
        );
        assert_eq!(
            &output_raw.central_order[..source_raw.central_order.len()],
            source_raw.central_order.as_slice()
        );
        assert_eq!(
            &output_raw.local_order[source_raw.local_order.len()..],
            ["custom/_rels/second.bin.rels"]
        );
        assert_eq!(
            &output_raw.central_order[source_raw.central_order.len()..],
            ["custom/_rels/second.bin.rels"]
        );
        for name in &source_raw.local_order {
            assert_eq!(
                output_raw.local_members[name], source_raw.local_members[name],
                "{name}"
            );
            assert_eq!(
                central_without_local_offset(&output_raw.central_records[name]),
                central_without_local_offset(&source_raw.central_records[name]),
                "{name}"
            );
        }
        let reopened = OpcPackage::from_bytes(&output).expect("reopen relationship output");
        let relationships = reopened.get_part(&second).unwrap().rels();
        assert_eq!(relationships.len(), 1);
        assert_eq!(
            relationships.get("rId1").unwrap().target_ref(),
            "https://example.com"
        );
    }

    #[test]
    fn topology_add_with_unknown_non_part_refuses_normalization() {
        let (source, _first) = source_with_non_part_framing();
        let mut package = OpcPackage::from_vec(source).expect("open unknown-member source");
        package.add_part(Box::new(crate::BlobPart::new(
            PackURI::new("/custom/third.bin").unwrap(),
            "application/octet-stream".to_owned(),
            b"third".to_vec(),
        )));
        assert!(matches!(
            PackageWriter::to_bytes(&package),
            Err(crate::OpcError::PreservationUnavailable { .. })
        ));
    }

    #[test]
    fn unsupported_prefixed_source_refuses_normalization_before_output() {
        let (source, first, _second) = two_part_source(b"prefixed fallback");
        let mut prefixed = b"unsupported ZIP prelude".to_vec();
        prefixed.extend_from_slice(&source);
        let mut package = OpcPackage::from_vec(prefixed.clone()).expect("open prefixed OPC");
        assert_eq!(
            PackageWriter::to_bytes(&package).expect("exact prefixed copy"),
            prefixed
        );
        package
            .get_part_mut(&first)
            .unwrap()
            .set_blob(b"changed".to_vec());

        let mut output = Vec::new();
        let error = PackageWriter::write_to_stream(&mut output, &package)
            .expect_err("unsupported source framing must be rejected");
        assert!(matches!(
            error,
            crate::OpcError::PreservationUnavailable { .. }
        ));
        assert!(output.is_empty());
    }

    #[test]
    fn unsupported_owned_source_revoked_noop_refuses_before_output() {
        let (source, _first, _second) = two_part_source(b"prefixed no-op");
        let mut prefixed = b"unsupported ZIP prelude".to_vec();
        prefixed.extend_from_slice(&source);
        let mut package = OpcPackage::from_vec(prefixed).expect("open prefixed OPC");
        let options = package.save_options().clone();
        package.set_save_options(options);

        let mut output = Vec::new();
        let error = PackageWriter::write_to_stream(&mut output, &package)
            .expect_err("unsupported source framing must be rejected");
        assert!(matches!(
            error,
            crate::OpcError::PreservationUnavailable { .. }
        ));
        assert!(output.is_empty());
    }

    #[test]
    fn trailing_source_bytes_refuse_normalization_before_targeted_preservation() {
        let (source, first, _second) = two_part_source(b"trailing fallback");
        let mut suffixed = source;
        suffixed.extend_from_slice(b"trailing bytes outside EOCD");
        let mut package = OpcPackage::from_vec(suffixed).expect("open suffixed OPC");
        package
            .get_part_mut(&first)
            .expect("first part")
            .set_blob(b"changed payload".to_vec());

        assert!(matches!(
            PackageWriter::to_bytes(&package),
            Err(crate::OpcError::PreservationUnavailable { .. })
        ));
    }

    #[test]
    fn zip64_source_refuses_normalization_after_mutation() {
        let (source, first, _second) = two_part_source(b"ZIP64 fallback");
        let source = promote_to_zip64(source);
        assert!(
            soapberry_zip::ZipArchive::from_slice(&source)
                .unwrap()
                .is_zip64()
        );
        let mut package = OpcPackage::from_vec(source.clone()).expect("open ZIP64 OPC");
        assert_eq!(PackageWriter::to_bytes(&package).unwrap(), source);
        package
            .get_part_mut(&first)
            .unwrap()
            .set_blob(b"ZIP64 changed".to_vec());

        assert!(matches!(
            PackageWriter::to_bytes(&package),
            Err(crate::OpcError::PreservationUnavailable { .. })
        ));
    }

    #[test]
    fn rejects_invalid_part_content_type_before_writing() {
        let mut package = OpcPackage::new();
        package.add_part(Box::new(crate::BlobPart::new(
            PackURI::new("/custom/data.bin").unwrap(),
            "application/octet-stream (comment)".to_string(),
            Vec::new(),
        )));
        assert!(matches!(
            PackageWriter::to_bytes(&package),
            Err(crate::OpcError::InvalidContentType { .. })
        ));
    }

    #[test]
    fn refuses_arbitrary_authored_xml_bytes_before_publication() {
        for (part_name, content_type) in [
            ("/custom/manifest.rdf", "application/octet-stream"),
            ("/custom/metadata", "application/rdf+xml"),
            (
                "/_xmlsignatures/sig1.bin",
                "application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml",
            ),
        ] {
            let mut package = OpcPackage::new();
            package.add_part(Box::new(crate::BlobPart::new(
                PackURI::new(part_name).expect("valid part URI"),
                content_type.to_string(),
                b"<root> <child/></root>".to_vec(),
            )));

            assert!(matches!(
                PackageWriter::to_bytes(&package),
                Err(crate::OpcError::XmlPublication { part, .. }) if part == part_name
            ));
        }
    }

    #[test]
    fn publication_plan_failure_leaves_sequential_sink_untouched() {
        let mut package = OpcPackage::new();
        package.add_part(Box::new(crate::BlobPart::new(
            PackURI::new("/custom/metadata.xml").expect("valid part URI"),
            ct::XML.to_owned(),
            b"<root> <child/></root>".to_vec(),
        )));
        let mut sink = ChunkSink {
            total: 0,
            writes: 0,
            largest: 0,
            limit: usize::MAX,
        };

        let error = PackageWriter::write_to_stream(&mut sink, &package)
            .expect_err("authored XML must fail publication planning");

        assert!(matches!(
            error,
            crate::OpcError::XmlPublication { part, .. }
                if part == "/custom/metadata.xml"
        ));
        assert_eq!(sink.total, 0);
        assert_eq!(sink.writes, 0);
    }

    #[test]
    fn exact_source_xml_bytes_may_remain_opaque() {
        let content_types = br#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/custom/manifest.rdf" ContentType="application/rdf+xml"/></Types>"#;
        let relationships = br#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>"#;
        let source_rdf = b"<rdf:RDF xmlns:rdf=\"urn:test\">\n <rdf:Description/>\n</rdf:RDF>";
        let mut physical = PhysPkgWriter::new();
        physical
            .write(
                &PackURI::new("/[Content_Types].xml").expect("content-types URI"),
                content_types,
            )
            .expect("write content types");
        physical
            .write(
                &PackURI::new("/_rels/.rels").expect("relationship URI"),
                relationships,
            )
            .expect("write relationships");
        physical
            .write(
                &PackURI::new("/custom/manifest.rdf").expect("RDF URI"),
                source_rdf,
            )
            .expect("write source RDF");
        let source = physical.finish().expect("finish source package");

        let package = OpcPackage::from_vec(source).expect("open source package");
        let rewritten = PackageWriter::to_bytes(&package).expect("preserve source RDF");
        let rewritten_physical = crate::phys_pkg::OwnedPhysPkgReader::from_bytes(rewritten)
            .expect("open rewritten package");
        assert_eq!(
            rewritten_physical
                .read_member("custom/manifest.rdf")
                .expect("read rewritten RDF"),
            source_rdf
        );
    }

    #[test]
    fn real_package_enumeration_covers_all_xml_bearing_members() {
        let mut package = OpcPackage::new();
        for (part_name, content_type, payload) in [
            (
                "/custom/manifest.rdf",
                "application/rdf+xml",
                b"<rdf:RDF xmlns:rdf=\"urn:test\"/>".as_slice(),
            ),
            (
                "/custom/metadata",
                "application/vnd.example.metadata+xml",
                b"<metadata/>".as_slice(),
            ),
            (
                "/_xmlsignatures/sig1.xml",
                "application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml",
                b"<Signature xmlns=\"http://www.w3.org/2000/09/xmldsig#\"/>".as_slice(),
            ),
        ] {
            package.add_part(Box::new(crate::BlobPart::new(
                PackURI::new(part_name).expect("valid part URI"),
                content_type.to_string(),
                payload.to_vec(),
            )));
        }
        let bytes = PackageWriter::to_bytes(&package).expect("publish package");
        let physical =
            crate::phys_pkg::OwnedPhysPkgReader::from_bytes(bytes).expect("open published package");
        let mut audited = Vec::new();
        for name in physical.member_names().expect("enumerate members") {
            let media_type = match name.as_str() {
                "custom/metadata" => "application/vnd.example.metadata+xml",
                "custom/manifest.rdf" => "application/rdf+xml",
                "_xmlsignatures/sig1.xml" => {
                    "application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml"
                },
                _ => "application/octet-stream",
            };
            if xml_minifier::audit::package::is_xml_part(&name, media_type) {
                let payload = physical.read_member(&name).expect("read XML member");
                let _report = xml_minifier::audit::verify_authored(
                    &payload,
                    xml_minifier::audit::Limits::default(),
                )
                .expect("emitted XML is compact");
                audited.push(name);
            }
        }
        audited.sort();
        assert_eq!(
            audited,
            [
                "[Content_Types].xml",
                "_rels/.rels",
                "_xmlsignatures/sig1.xml",
                "custom/manifest.rdf",
                "custom/metadata",
            ]
        );
    }

    #[test]
    fn streams_large_packages_to_a_non_seekable_bounded_chunk_sink() {
        let mut state = 0x9e37_79b9_u32;
        let mut payload = Vec::with_capacity(2 * 1024 * 1024);
        while payload.len() < payload.capacity() {
            state ^= state << 13;
            state ^= state >> 17;
            state ^= state << 5;
            payload.push((state >> 24) as u8);
        }
        let mut package = OpcPackage::new();
        package.add_part(Box::new(crate::BlobPart::new(
            PackURI::new("/custom/random.bin").expect("valid part URI"),
            "application/octet-stream".to_owned(),
            payload,
        )));
        let mut sink = ChunkSink {
            total: 0,
            writes: 0,
            largest: 0,
            limit: 64 * 1024,
        };

        PackageWriter::write_to_stream(&mut sink, &package).expect("stream package");

        assert!(sink.total > 1024 * 1024);
        assert!(sink.writes > 1);
        assert!(sink.largest <= sink.limit);
    }

    #[test]
    fn incomplete_stream_errors_report_accepted_bytes() {
        let package = OpcPackage::new();
        let sink = FailAfter {
            written: 0,
            limit: 128,
        };

        let error = PackageWriter::write_to_stream(sink, &package)
            .expect_err("bounded sink must reject the package");

        assert!(matches!(
            error,
            crate::OpcError::IncompleteOutput {
                written: 128,
                source,
            } if matches!(*source, crate::OpcError::ZipError(_))
        ));
    }

    #[test]
    fn filesystem_write_replaces_only_with_a_finalized_package() {
        let directory = tempfile::tempdir().expect("temporary directory");
        let destination = directory.path().join("package.xlsx");
        std::fs::write(&destination, b"previous artifact").expect("seed destination");
        let mut package = OpcPackage::new();
        let partname = PackURI::new("/custom/data.bin").expect("valid part URI");
        package.add_part(Box::new(crate::BlobPart::new(
            partname.clone(),
            "application/octet-stream".to_owned(),
            b"payload".to_vec(),
        )));

        PackageWriter::write(&destination, &package).expect("atomic package write");

        let reopened = OpcPackage::open(destination).expect("reopen package");
        assert_eq!(
            reopened.get_part(&partname).expect("saved part").blob(),
            b"payload"
        );
    }

    #[test]
    fn invalid_packages_never_replace_an_existing_artifact() {
        let directory = tempfile::tempdir().expect("temporary directory");
        let destination = directory.path().join("package.xlsx");
        std::fs::write(&destination, b"previous artifact").expect("seed destination");
        let mut package = OpcPackage::new();
        package.add_part(Box::new(crate::BlobPart::new(
            PackURI::new("/custom/data.bin").expect("valid part URI"),
            "invalid content type".to_owned(),
            Vec::new(),
        )));

        let result = PackageWriter::write(&destination, &package);

        assert!(matches!(
            result,
            Err(crate::OpcError::InvalidContentType { .. })
        ));
        assert_eq!(
            std::fs::read(destination).expect("read destination"),
            b"previous artifact"
        );
    }
}
