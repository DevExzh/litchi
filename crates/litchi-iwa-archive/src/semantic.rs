//! Narrow, owned semantic projection for iWork ZIP packages.
//!
//! Unlike the preserve catalog, this projection owns no source ZIP bytes or
//! raw entry records. It reads only exact canonical IWA authorities and, when
//! requested, iWork's three exact semantic metadata authorities.

use std::collections::HashSet;

use litchi_iwa_core::{Archive, SnappyLimits, SnappyStream};

use crate::catalog::{Component, ComponentCatalog};
use crate::package::{
    preflight_semantic_container, preflight_semantic_iwa, preflight_semantic_iwa_entries,
    raw_path_normalizes_to, reject_semantic_aliases, semantic_iwa_name,
    semantic_nested_index_entry,
};
use crate::zip::{PhysicalEntry, ZipArchive, is_encrypted};
use crate::{Error, LimitKind, Limits, Result};

const MAX_METADATA_BYTES: u64 = 64 * 1024;
const MAX_SEMANTIC_IWA_OBJECTS: usize = 1_000_000;

const PROPERTIES: &[u8] = b"Metadata/Properties.plist";
const BUILD_HISTORY: &[u8] = b"Metadata/BuildVersionHistory.plist";
const DOCUMENT_IDENTIFIER: &[u8] = b"Metadata/DocumentIdentifier";

/// Fixed member admission profiles for a semantic package projection.
#[doc(hidden)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum SemanticProfile {
    /// Parse exact canonical IWA components and retain no metadata sidecars.
    ComponentsOnly,
    /// Parse exact canonical IWA components and retain three bounded sidecars.
    Metadata,
}

impl SemanticProfile {
    const fn includes_metadata(self) -> bool {
        matches!(self, Self::Metadata)
    }
}

/// Owned payloads for iWork's exact semantic metadata authorities.
#[doc(hidden)]
#[derive(Debug, Default, Clone, PartialEq, Eq)]
pub struct SemanticMetadataSidecars {
    properties: Option<Box<[u8]>>,
    build_history: Option<Box<[u8]>>,
    document_identifier: Option<Box<[u8]>>,
}

impl SemanticMetadataSidecars {
    /// Borrow exact `Metadata/Properties.plist`, if present.
    #[must_use]
    pub fn properties_plist(&self) -> Option<&[u8]> {
        self.properties.as_deref()
    }

    /// Borrow exact `Metadata/BuildVersionHistory.plist`, if present.
    #[must_use]
    pub fn build_version_history_plist(&self) -> Option<&[u8]> {
        self.build_history.as_deref()
    }

    /// Borrow exact `Metadata/DocumentIdentifier`, if present.
    #[must_use]
    pub fn document_identifier(&self) -> Option<&[u8]> {
        self.document_identifier.as_deref()
    }

    /// Consume the DTO and return its three owned payloads in authority order.
    #[must_use]
    pub fn into_parts(self) -> (Option<Box<[u8]>>, Option<Box<[u8]>>, Option<Box<[u8]>>) {
        (
            self.properties,
            self.build_history,
            self.document_identifier,
        )
    }

    fn set(&mut self, authority: &[u8], data: Vec<u8>) -> Result<()> {
        let slot = match authority {
            PROPERTIES => &mut self.properties,
            BUILD_HISTORY => &mut self.build_history,
            DOCUMENT_IDENTIFIER => &mut self.document_identifier,
            _ => {
                return Err(Error::InvalidBundle(
                    "unknown semantic metadata authority".to_owned(),
                ));
            },
        };
        if slot.is_some() {
            return Err(Error::InvalidBundle(format!(
                "duplicate semantic metadata authority is ambiguous: {}",
                String::from_utf8_lossy(authority)
            )));
        }
        *slot = Some(data.into_boxed_slice());
        Ok(())
    }
}

/// Owned semantic components and optional bounded metadata from one ZIP.
#[doc(hidden)]
#[derive(Debug)]
pub struct SemanticProjection {
    components: ComponentCatalog,
    sidecars: SemanticMetadataSidecars,
    limits: Limits,
}

impl SemanticProjection {
    /// Parse one borrowed complete ZIP under a fixed semantic profile.
    ///
    /// Unrelated members are inspected structurally by ZIP ingress but their
    /// payloads are never decompressed or retained by this projection.
    pub fn from_bytes_with_limits(
        bytes: &[u8],
        limits: Limits,
        profile: SemanticProfile,
    ) -> Result<Self> {
        let validated_limits = limits.validate()?;
        let archive = ZipArchive::new_with_limits(bytes, validated_limits)?;
        if is_encrypted(&archive) {
            return Err(Error::Encrypted);
        }
        reject_semantic_aliases(&archive)?;

        let direct = selected_components(&archive)?;
        let nested_index = semantic_nested_index_entry(&archive)?;
        if !direct.is_empty() && nested_index.is_some() {
            return Err(Error::InvalidBundle(
                "iWork package mixes direct IWA members with a legacy Index.zip".to_owned(),
            ));
        }

        let (components, sidecars) = if direct.is_empty() {
            if let Some(index_entry) = nested_index {
                project_legacy(&archive, index_entry, validated_limits, profile)?
            } else {
                (Vec::new(), SemanticMetadataSidecars::default())
            }
        } else {
            project_modern(&archive, direct, validated_limits, profile)?
        };

        Ok(Self {
            components: ComponentCatalog::from_semantic_components(components),
            sidecars,
            limits: validated_limits,
        })
    }

    /// Borrow the parsed neutral IWA components.
    #[must_use]
    pub const fn components(&self) -> &ComponentCatalog {
        &self.components
    }

    /// Borrow the owned semantic metadata DTO.
    #[must_use]
    pub const fn sidecars(&self) -> &SemanticMetadataSidecars {
        &self.sidecars
    }

    /// Return the validated physical limits used for this projection.
    #[must_use]
    pub const fn limits(&self) -> Limits {
        self.limits
    }

    /// Consume the projection into components, sidecars, and validated limits.
    #[must_use]
    pub fn into_parts(self) -> (ComponentCatalog, SemanticMetadataSidecars, Limits) {
        (self.components, self.sidecars, self.limits)
    }
}

fn project_modern(
    archive: &ZipArchive<'_>,
    components: Vec<(&PhysicalEntry, &str)>,
    limits: Limits,
    profile: SemanticProfile,
) -> Result<(Vec<Component>, SemanticMetadataSidecars)> {
    // The whole selected batch is checked before its first payload read.
    preflight_semantic_iwa_entries(archive, false)?;
    let metadata = selected_metadata(archive, b"", profile)?;
    preflight_metadata(&metadata)?;

    let components = read_components(archive, components, limits, false)?;
    let sidecars = read_metadata(archive, metadata)?;
    Ok((components, sidecars))
}

fn project_legacy(
    archive: &ZipArchive<'_>,
    index_entry: &PhysicalEntry,
    limits: Limits,
    profile: SemanticProfile,
) -> Result<(Vec<Component>, SemanticMetadataSidecars)> {
    let index_name = index_entry.name();
    preflight_semantic_container(index_entry, index_name)?;
    limits.check_input_size(index_entry.uncompressed_size(), "legacy iWork Index.zip")?;
    let raw_prefix = index_entry
        .raw_name()
        .strip_suffix(b"Index.zip")
        .ok_or_else(|| {
            Error::InvalidBundle("legacy Index.zip raw name has no suffix".to_owned())
        })?;
    let metadata = selected_metadata(archive, raw_prefix, profile)?;
    preflight_metadata(&metadata)?;

    // ArchiveReader's bounded read path performs a fallible exact reserve
    // against the already-validated declared size.
    let index_data = archive.read_entry(index_entry)?;
    let index_size = u64::try_from(index_data.len()).map_err(|_error| {
        Error::InvalidBundle("legacy iWork Index.zip length does not fit u64".to_owned())
    })?;
    limits.check_input_size(index_size, "legacy iWork Index.zip")?;
    let index = ZipArchive::new_with_limits(&index_data, limits).map_err(|error| {
        if matches!(&error, Error::Limit { .. }) {
            error
        } else {
            Error::InvalidBundle(format!("legacy package index: {error}"))
        }
    })?;
    if is_encrypted(&index) {
        return Err(Error::Encrypted);
    }
    preflight_legacy_components(&index)?;
    let components = selected_legacy_components(&index)?;
    if components.is_empty() {
        return Err(Error::InvalidBundle(format!(
            "legacy package index {index_name} contains no IWA components"
        )));
    }
    let components = read_components(&index, components, limits, true)?;
    let sidecars = read_metadata(archive, metadata)?;
    Ok((components, sidecars))
}

fn read_components(
    archive: &ZipArchive<'_>,
    entries: Vec<(&PhysicalEntry, &str)>,
    limits: Limits,
    legacy: bool,
) -> Result<Vec<Component>> {
    let mut components = Vec::new();
    let mut total_iwa_bytes = 0u64;
    let mut total_objects = 0usize;
    for (entry, name) in entries {
        let core_limits = semantic_component_limits(limits, total_objects)?;
        let snappy_limits = semantic_iwa_snappy_limits(limits, total_iwa_bytes)?;
        let compressed = archive.read_entry(entry)?;
        if name.rsplit('/').next() == Some("OperationStorage.iwa")
            && compressed.starts_with(b"bvxn")
        {
            continue;
        }
        let decompressed = SnappyStream::decompress_with_limits(&compressed, snappy_limits)
            .map_err(|error| map_semantic_iwa_total_limit(error, total_iwa_bytes, limits))?;
        let decompressed_bytes =
            u64::try_from(decompressed.as_bytes().len()).map_err(|_error| {
                Error::InvalidBundle("decompressed IWA stream length does not fit u64".to_owned())
            })?;
        total_iwa_bytes = limits.charge_iwa_total_bytes(total_iwa_bytes, decompressed_bytes)?;
        let parsed = Archive::parse_with_limits(decompressed.as_bytes(), core_limits)?;
        total_objects = total_objects
            .checked_add(parsed.objects.len())
            .ok_or_else(|| semantic_object_limit(usize::MAX))?;
        if total_objects > MAX_SEMANTIC_IWA_OBJECTS {
            return Err(semantic_object_limit(total_objects));
        }
        components
            .try_reserve(1)
            .map_err(|_error| Error::Allocation {
                resource: "semantic IWA component catalog",
                amount: 1,
            })?;
        let component = if legacy {
            Component::try_new_legacy_index_member(name, parsed)?
        } else {
            Component::try_new(name, parsed)?
        };
        components.push(component);
    }
    Ok(components)
}

fn selected_components<'archive, 'data>(
    archive: &'archive ZipArchive<'data>,
) -> Result<Vec<(&'archive PhysicalEntry, &'archive str)>> {
    let mut selected = Vec::new();
    for entry in archive
        .physical_entries()
        .filter(|entry| !entry.is_directory())
    {
        let name = semantic_iwa_name(entry);
        let Some(name) = name else {
            continue;
        };
        selected
            .try_reserve(1)
            .map_err(|_error| Error::Allocation {
                resource: "semantic IWA selection",
                amount: 1,
            })?;
        selected.push((entry, name));
    }
    Ok(selected)
}

fn preflight_legacy_components(archive: &ZipArchive<'_>) -> Result<()> {
    let mut seen = HashSet::new();
    for entry in archive
        .physical_entries()
        .filter(|entry| !entry.is_directory())
    {
        let Some(basename) = legacy_component_basename(entry) else {
            if !is_legacy_irrelevant_payload(entry.raw_name()) {
                return Err(Error::InvalidBundle(format!(
                    "legacy package index contains a non-canonical IWA member: {}",
                    entry.name()
                )));
            }
            continue;
        };
        preflight_semantic_iwa(entry, basename)?;
        seen.try_reserve(1).map_err(|_error| Error::Allocation {
            resource: "semantic legacy IWA authority names",
            amount: 1,
        })?;
        if !seen.insert(basename) {
            return Err(Error::InvalidBundle(format!(
                "duplicate semantic IWA authority is ambiguous: Index/{basename}"
            )));
        }
    }
    Ok(())
}

fn selected_legacy_components<'archive, 'data>(
    archive: &'archive ZipArchive<'data>,
) -> Result<Vec<(&'archive PhysicalEntry, &'archive str)>> {
    let mut selected = Vec::new();
    for entry in archive
        .physical_entries()
        .filter(|entry| !entry.is_directory())
    {
        let Some(basename) = legacy_component_basename(entry) else {
            continue;
        };
        selected
            .try_reserve(1)
            .map_err(|_error| Error::Allocation {
                resource: "semantic legacy IWA selection",
                amount: 1,
            })?;
        selected.push((entry, basename));
    }
    Ok(selected)
}

fn legacy_component_basename(physical: &PhysicalEntry) -> Option<&str> {
    let raw_name = physical.raw_name();
    if !is_exact_portable_raw_name(raw_name)
        || raw_name.contains(&b'/')
        || !raw_name.ends_with(b".iwa")
    {
        return None;
    }
    std::str::from_utf8(raw_name).ok()
}

fn is_legacy_irrelevant_payload(raw_name: &[u8]) -> bool {
    is_exact_portable_raw_name(raw_name)
        && (raw_name.starts_with(b"Data/") || raw_name.starts_with(b"Preview/"))
}

fn is_exact_portable_raw_name(raw_name: &[u8]) -> bool {
    let Ok(name) = std::str::from_utf8(raw_name) else {
        return false;
    };
    !name.is_empty()
        && !name.starts_with('/')
        && !name.ends_with('/')
        && !name.contains(['\0', '\\'])
        && !name.chars().any(char::is_control)
        && !name.split('/').any(|component| {
            component.is_empty() || matches!(component, "." | "..") || component.contains(':')
        })
}

fn semantic_component_limits(
    limits: Limits,
    retained_objects: usize,
) -> Result<litchi_iwa_core::ArchiveLimits> {
    let remaining = MAX_SEMANTIC_IWA_OBJECTS.saturating_sub(retained_objects);
    if remaining == 0 {
        return Err(semantic_object_limit(MAX_SEMANTIC_IWA_OBJECTS + 1));
    }
    let core_limits = limits.effective_archive_limits()?;
    core_limits
        .with_objects(core_limits.max_objects().min(remaining))
        .map_err(Error::Iwa)
}

/// Cap the next Snappy decode by the aggregate IWA bytes still unretained.
///
/// This turns the package-wide ceiling into a pre-allocation decompression
/// bound. In particular, once no bytes remain the next selected component is
/// rejected before its ZIP payload is read.
fn semantic_iwa_snappy_limits(limits: Limits, total_iwa_bytes: u64) -> Result<SnappyLimits> {
    let remaining = limits
        .max_total_bytes()
        .checked_sub(total_iwa_bytes)
        .ok_or_else(|| Error::Limit {
            kind: LimitKind::IwaTotalBytes,
            observed: total_iwa_bytes,
            maximum: limits.max_total_bytes(),
        })?;
    if remaining == 0 {
        return Err(Error::Limit {
            kind: LimitKind::IwaTotalBytes,
            observed: limits.max_total_bytes().saturating_add(1),
            maximum: limits.max_total_bytes(),
        });
    }
    let remaining = usize::try_from(remaining).map_err(|_error| {
        Error::InvalidBundle("remaining semantic IWA byte budget does not fit usize".to_owned())
    })?;
    let base = limits.snappy_limits()?;
    let stream = base.max_decompressed_stream().min(remaining);
    SnappyLimits::new(base.max_uncompressed_chunk().min(stream), stream).map_err(Error::Iwa)
}

fn map_semantic_iwa_total_limit(
    error: litchi_iwa_core::Error,
    current: u64,
    limits: Limits,
) -> Error {
    let litchi_iwa_core::Error::Limit {
        kind:
            litchi_iwa_core::LimitKind::SnappyChunkBytes | litchi_iwa_core::LimitKind::SnappyStreamBytes,
        observed,
        ..
    } = error
    else {
        return Error::Iwa(error);
    };
    Error::Limit {
        kind: LimitKind::IwaTotalBytes,
        observed: current.saturating_add(u64::try_from(observed).unwrap_or(u64::MAX)),
        maximum: limits.max_total_bytes(),
    }
}

fn selected_metadata<'a>(
    archive: &'a ZipArchive<'_>,
    raw_prefix: &[u8],
    profile: SemanticProfile,
) -> Result<Vec<(&'a PhysicalEntry, &'a [u8])>> {
    if !profile.includes_metadata() {
        return Ok(Vec::new());
    }
    let mut selected = Vec::new();
    let mut seen = HashSet::new();
    for entry in archive
        .physical_entries()
        .filter(|entry| !entry.is_directory())
    {
        let central = entry.raw_name().strip_prefix(raw_prefix);
        let local = entry.local_header().name.strip_prefix(raw_prefix);
        let Some(authority) = metadata_authority_collision(central, local)? else {
            continue;
        };
        seen.try_reserve(1).map_err(|_error| Error::Allocation {
            resource: "semantic metadata authority names",
            amount: 1,
        })?;
        if !seen.insert(authority) {
            return Err(Error::InvalidBundle(format!(
                "duplicate semantic metadata authority is ambiguous: {}",
                String::from_utf8_lossy(authority)
            )));
        }
        selected
            .try_reserve(1)
            .map_err(|_error| Error::Allocation {
                resource: "semantic metadata selection",
                amount: 1,
            })?;
        selected.push((entry, authority));
    }
    Ok(selected)
}

/// Admit one authority only when both ZIP name records select that exact raw
/// member.  ZIP path normalization is never an alternate authority spelling.
fn metadata_authority_collision<'a>(
    central: Option<&'a [u8]>,
    local: Option<&[u8]>,
) -> Result<Option<&'a [u8]>> {
    for &authority in [PROPERTIES, BUILD_HISTORY, DOCUMENT_IDENTIFIER].as_slice() {
        let central_exact = central == Some(authority);
        let local_exact = local == Some(authority);
        let central_alias = central
            .is_some_and(|name| name != authority && raw_path_normalizes_to(name, authority));
        let local_alias =
            local.is_some_and(|name| name != authority && raw_path_normalizes_to(name, authority));

        if central_alias || local_alias || central_exact != local_exact {
            return Err(Error::InvalidBundle(format!(
                "semantic metadata authority has non-canonical or one-sided ZIP names: {}",
                String::from_utf8_lossy(authority)
            )));
        }
        if central_exact {
            return Ok(Some(authority));
        }
    }
    Ok(None)
}

fn preflight_metadata(entries: &[(&PhysicalEntry, &[u8])]) -> Result<()> {
    for &(entry, authority) in entries {
        let local = entry.local_header();
        let central = entry.central_header();
        if local.name.as_ref() != central.name.as_ref() {
            return Err(Error::InvalidBundle(format!(
                "semantic metadata authority {} has mismatched local and central names",
                String::from_utf8_lossy(authority)
            )));
        }
        if local.compression_method != central.compression_method {
            return Err(Error::InvalidBundle(format!(
                "semantic metadata authority {} has mismatched local and central compression methods",
                String::from_utf8_lossy(authority)
            )));
        }
        if !matches!(central.compression_method, 0 | 8) {
            return Err(Error::InvalidBundle(format!(
                "semantic metadata authority {} uses unsupported ZIP compression",
                String::from_utf8_lossy(authority)
            )));
        }
        if entry.uncompressed_size() > MAX_METADATA_BYTES {
            return Err(Error::Limit {
                kind: LimitKind::EntryBytes,
                observed: entry.uncompressed_size(),
                maximum: MAX_METADATA_BYTES,
            });
        }
    }
    Ok(())
}

fn read_metadata(
    archive: &ZipArchive<'_>,
    entries: Vec<(&PhysicalEntry, &[u8])>,
) -> Result<SemanticMetadataSidecars> {
    let mut sidecars = SemanticMetadataSidecars::default();
    for (entry, authority) in entries {
        sidecars.set(authority, archive.read_entry(entry)?)?;
    }
    Ok(sidecars)
}

const fn semantic_object_limit(observed: usize) -> Error {
    Error::Iwa(litchi_iwa_core::Error::Limit {
        kind: litchi_iwa_core::LimitKind::Objects,
        observed,
        maximum: MAX_SEMANTIC_IWA_OBJECTS,
    })
}

#[cfg(test)]
mod tests {
    use litchi_iwa_core::{ArchiveObject, RawMessage};
    use soapberry_zip::office::StreamingArchiveWriter;

    use super::*;

    fn iwa(identifier: u64) -> Result<Vec<u8>> {
        let archive = Archive {
            objects: vec![ArchiveObject::new(
                identifier,
                vec![RawMessage {
                    type_: 6_000,
                    data: vec![1, 2, 3],
                }],
            )?],
        };
        Ok(SnappyStream::compress(&archive.to_bytes()?)?)
    }

    fn iwa_with_payload(identifier: u64, payload_bytes: usize) -> Result<(Vec<u8>, u64)> {
        let archive = Archive {
            objects: vec![ArchiveObject::new(
                identifier,
                vec![RawMessage {
                    type_: 6_000,
                    data: vec![0; payload_bytes],
                }],
            )?],
        };
        let decompressed = archive.to_bytes()?;
        let bytes = u64::try_from(decompressed.len()).map_err(|_error| {
            Error::InvalidBundle("test IWA stream length does not fit u64".to_owned())
        })?;
        Ok((SnappyStream::compress(&decompressed)?, bytes))
    }

    fn zip(entries: &[(&str, &[u8])]) -> Result<Vec<u8>> {
        let mut writer = StreamingArchiveWriter::new();
        for &(name, data) in entries {
            writer.write_stored(name, data)?;
        }
        Ok(writer.finish_to_bytes()?)
    }

    fn make_method_opaque(bytes: &mut [u8], name: &[u8]) {
        let positions = bytes
            .windows(name.len())
            .enumerate()
            .filter_map(|(index, candidate)| (candidate == name).then_some(index))
            .collect::<Vec<_>>();
        assert_eq!(positions.len(), 2);
        // Local and central filename fields begin 30 and 46 bytes after their
        // signatures; their compression fields begin at offsets 8 and 10.
        bytes[positions[0] - 22..positions[0] - 20].copy_from_slice(&99u16.to_le_bytes());
        bytes[positions[1] - 36..positions[1] - 34].copy_from_slice(&99u16.to_le_bytes());
    }

    fn replace_raw_name(bytes: &mut [u8], old: &[u8], new: &[u8]) {
        assert_eq!(old.len(), new.len());
        let positions = bytes
            .windows(old.len())
            .enumerate()
            .filter_map(|(index, candidate)| (candidate == old).then_some(index))
            .collect::<Vec<_>>();
        assert_eq!(positions.len(), 2);
        for position in positions {
            bytes[position..position + new.len()].copy_from_slice(new);
        }
    }

    fn replace_raw_name_occurrence(bytes: &mut [u8], old: &[u8], new: &[u8], occurrence: usize) {
        assert_eq!(old.len(), new.len());
        let positions = bytes
            .windows(old.len())
            .enumerate()
            .filter_map(|(index, candidate)| (candidate == old).then_some(index))
            .collect::<Vec<_>>();
        assert_eq!(positions.len(), 2);
        let position = positions[occurrence];
        bytes[position..position + new.len()].copy_from_slice(new);
    }

    fn metadata_backslash_alias(name: &[u8]) -> Vec<u8> {
        let mut alias = name.to_vec();
        let offset = alias
            .windows(b"Metadata/".len())
            .position(|candidate| candidate == b"Metadata/")
            .expect("metadata authority contains the Metadata/ prefix");
        alias[offset + b"Metadata".len()] = b'\\';
        alias
    }

    fn assert_metadata_collision_is_unread(bytes: &[u8]) {
        crate::zip::reset_test_entry_read_count();
        assert!(matches!(
            SemanticProjection::from_bytes_with_limits(
                bytes,
                Limits::default(),
                SemanticProfile::Metadata,
            ),
            Err(Error::InvalidBundle(message))
                if message.contains("semantic metadata authority has non-canonical or one-sided ZIP names")
        ));
        assert_eq!(crate::zip::test_entry_read_count(), 0);
    }

    #[test]
    fn modern_projection_reads_only_components_and_exact_sidecars() -> Result<()> {
        let document = iwa(1)?;
        let metadata = iwa(2)?;
        let oversized_decoy = vec![b'x'; MAX_METADATA_BYTES as usize + 1];
        let mut bytes = zip(&[
            ("Index/Document.iwa", &document),
            ("Index/Metadata.iwa", &metadata),
            ("Metadata/Properties.plist", b"properties"),
            ("Metadata/BuildVersionHistory.plist", b"history"),
            ("Metadata/DocumentIdentifier", b"identifier"),
            ("Metadata/Properties.plist.bak", &oversized_decoy),
            ("Data/opaque.bin", b"unrelated"),
            ("Preview/preview.jpg", b"unrelated"),
        ])?;
        make_method_opaque(&mut bytes, b"Data/opaque.bin");

        crate::zip::reset_test_entry_read_count();
        let projection = SemanticProjection::from_bytes_with_limits(
            &bytes,
            Limits::default(),
            SemanticProfile::Metadata,
        )?;

        assert_eq!(crate::zip::test_entry_read_count(), 5);
        assert_eq!(projection.components().len(), 2);
        assert_eq!(
            projection.sidecars().properties_plist(),
            Some(b"properties".as_slice())
        );
        assert_eq!(
            projection.sidecars().build_version_history_plist(),
            Some(b"history".as_slice())
        );
        assert_eq!(
            projection.sidecars().document_identifier(),
            Some(b"identifier".as_slice())
        );
        Ok(())
    }

    #[test]
    fn components_only_does_not_read_metadata() -> Result<()> {
        let document = iwa(1)?;
        let bytes = zip(&[
            ("Index/Document.iwa", &document),
            ("Metadata/Properties.plist", b"properties"),
        ])?;

        crate::zip::reset_test_entry_read_count();
        let projection = SemanticProjection::from_bytes_with_limits(
            &bytes,
            Limits::default(),
            SemanticProfile::ComponentsOnly,
        )?;

        assert_eq!(crate::zip::test_entry_read_count(), 1);
        assert_eq!(projection.components().len(), 1);
        assert_eq!(projection.sidecars(), &SemanticMetadataSidecars::default());
        Ok(())
    }

    #[test]
    fn raw_metadata_normalization_alias_is_refused_before_reads() -> Result<()> {
        let document = iwa(1)?;
        let mut bytes = zip(&[
            ("Index/Document.iwa", &document),
            ("Metadata/Properties.plist", b"decoy"),
        ])?;
        replace_raw_name(
            &mut bytes,
            b"Metadata/Properties.plist",
            b"Metadata\\Properties.plist",
        );

        assert_metadata_collision_is_unread(&bytes);
        Ok(())
    }

    #[test]
    fn flat_metadata_name_collisions_in_either_zip_header_are_refused_before_reads() -> Result<()> {
        let document = iwa(1)?;
        for &authority in [PROPERTIES, BUILD_HISTORY, DOCUMENT_IDENTIFIER].as_slice() {
            let authority = std::str::from_utf8(authority).expect("metadata authority is UTF-8");
            let alias = metadata_backslash_alias(authority.as_bytes());

            for occurrence in 0..2 {
                let mut bytes = zip(&[("Index/Document.iwa", &document), (authority, b"sidecar")])?;
                replace_raw_name_occurrence(&mut bytes, authority.as_bytes(), &alias, occurrence);
                assert_metadata_collision_is_unread(&bytes);
            }
        }
        Ok(())
    }

    #[test]
    fn legacy_metadata_name_collisions_in_either_zip_header_are_refused_before_reads() -> Result<()>
    {
        let document = iwa(1)?;
        let nested = zip(&[("Document.iwa", &document)])?;
        for &authority in [PROPERTIES, BUILD_HISTORY, DOCUMENT_IDENTIFIER].as_slice() {
            let authority = std::str::from_utf8(authority).expect("metadata authority is UTF-8");
            let name = format!("legacy.pages/{authority}");
            let alias = metadata_backslash_alias(name.as_bytes());

            for occurrence in 0..2 {
                let mut bytes = zip(&[
                    ("legacy.pages/Index.zip", nested.as_slice()),
                    (name.as_str(), b"sidecar"),
                ])?;
                replace_raw_name_occurrence(&mut bytes, name.as_bytes(), &alias, occurrence);
                assert_metadata_collision_is_unread(&bytes);
            }
        }
        Ok(())
    }

    #[test]
    fn selected_opaque_metadata_is_refused_before_component_read() -> Result<()> {
        let document = iwa(1)?;
        let mut bytes = zip(&[
            ("Index/Document.iwa", &document),
            ("Metadata/Properties.plist", b"properties"),
        ])?;
        make_method_opaque(&mut bytes, b"Metadata/Properties.plist");

        crate::zip::reset_test_entry_read_count();
        assert!(matches!(
            SemanticProjection::from_bytes_with_limits(
                &bytes,
                Limits::default(),
                SemanticProfile::Metadata,
            ),
            Err(Error::InvalidBundle(message)) if message.contains("unsupported ZIP compression")
        ));
        assert_eq!(crate::zip::test_entry_read_count(), 0);
        Ok(())
    }

    #[test]
    fn exact_portable_component_subpaths_are_admitted() -> Result<()> {
        let document = iwa(1)?;
        let table = iwa(2)?;
        let bytes = zip(&[
            ("Index/Document.iwa", &document),
            ("Index/Tables/Tile.iwa", &table),
        ])?;

        crate::zip::reset_test_entry_read_count();
        let projection = SemanticProjection::from_bytes_with_limits(
            &bytes,
            Limits::default(),
            SemanticProfile::ComponentsOnly,
        )?;
        assert_eq!(crate::zip::test_entry_read_count(), 2);
        assert!(
            projection
                .components()
                .get("Index/Tables/Tile.iwa")
                .is_some()
        );
        Ok(())
    }

    #[test]
    fn native_numbers_fixture_projects_all_exact_components() -> Result<()> {
        let bytes = std::fs::read(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../test-data/iwork/numbers/basic.numbers"
        ))?;

        let projection = SemanticProjection::from_bytes_with_limits(
            &bytes,
            Limits::default(),
            SemanticProfile::Metadata,
        )?;

        assert!(projection.components().len() >= 30);
        assert!(projection.components().get("Index/Document.iwa").is_some());
        assert!(
            projection
                .components()
                .iter()
                .any(|component| component.name().starts_with("Index/Tables/"))
        );
        Ok(())
    }

    #[test]
    fn metadata_cap_is_checked_before_any_selected_payload_read() -> Result<()> {
        let document = iwa(1)?;
        let oversized = vec![b'x'; MAX_METADATA_BYTES as usize + 1];
        let bytes = zip(&[
            ("Index/Document.iwa", &document),
            ("Metadata/Properties.plist", &oversized),
        ])?;

        crate::zip::reset_test_entry_read_count();
        assert!(matches!(
            SemanticProjection::from_bytes_with_limits(
                &bytes,
                Limits::default(),
                SemanticProfile::Metadata,
            ),
            Err(Error::Limit {
                kind: LimitKind::EntryBytes,
                observed,
                maximum: MAX_METADATA_BYTES,
            }) if observed == oversized.len() as u64
        ));
        assert_eq!(crate::zip::test_entry_read_count(), 0);
        Ok(())
    }

    #[test]
    fn legacy_projection_reads_index_components_and_prefixed_sidecars_only() -> Result<()> {
        let document = iwa(1)?;
        let mut nested = zip(&[
            ("Document.iwa", &document),
            ("Data/asset.bin", b"asset"),
            ("Preview/preview.jpg", b"preview"),
        ])?;
        make_method_opaque(&mut nested, b"Data/asset.bin");
        let outer = zip(&[
            ("legacy.numbers/Index.zip", &nested),
            ("legacy.numbers/Metadata/Properties.plist", b"properties"),
            ("legacy.numbers/Data/asset.bin", b"outer asset"),
        ])?;

        crate::zip::reset_test_entry_read_count();
        let projection = SemanticProjection::from_bytes_with_limits(
            &outer,
            Limits::default(),
            SemanticProfile::Metadata,
        )?;

        assert_eq!(crate::zip::test_entry_read_count(), 3);
        assert_eq!(projection.components().len(), 1);
        assert_eq!(
            projection.sidecars().properties_plist(),
            Some(b"properties".as_slice())
        );
        Ok(())
    }

    #[test]
    fn exhausted_global_object_budget_refuses_a_component_parser() {
        assert!(matches!(
            semantic_component_limits(Limits::default(), MAX_SEMANTIC_IWA_OBJECTS),
            Err(Error::Iwa(litchi_iwa_core::Error::Limit {
                kind: litchi_iwa_core::LimitKind::Objects,
                observed,
                maximum: MAX_SEMANTIC_IWA_OBJECTS,
            })) if observed == MAX_SEMANTIC_IWA_OBJECTS + 1
        ));
    }

    #[test]
    fn remaining_object_budget_tightens_the_core_parser_limit() -> Result<()> {
        let limits = semantic_component_limits(Limits::default(), MAX_SEMANTIC_IWA_OBJECTS - 7)?;
        assert_eq!(limits.max_objects(), 7);
        Ok(())
    }

    #[test]
    fn aggregate_iwa_budget_reaches_snappy_and_refuses_zero_remaining_before_read() -> Result<()> {
        let (alpha, alpha_bytes) = iwa_with_payload(1, 256)?;
        let (bravo, bravo_bytes) = iwa_with_payload(2, 256)?;
        let exact_total = alpha_bytes.checked_add(bravo_bytes).ok_or_else(|| {
            Error::InvalidBundle("test IWA aggregate length overflowed u64".to_owned())
        })?;
        let max_entry_bytes = u64::try_from(alpha.len().max(bravo.len())).map_err(|_error| {
            Error::InvalidBundle("test ZIP member length does not fit u64".to_owned())
        })?;
        let max_iwa_bytes = usize::try_from(alpha_bytes.max(bravo_bytes)).map_err(|_error| {
            Error::InvalidBundle("test IWA stream length does not fit usize".to_owned())
        })?;
        let bytes = zip(&[("Index/Alpha.iwa", &alpha), ("Index/Bravo.iwa", &bravo)])?;
        let input_bytes = u64::try_from(bytes.len()).map_err(|_error| {
            Error::InvalidBundle("test ZIP length does not fit u64".to_owned())
        })?;

        let exact = Limits::new(input_bytes, 2, max_entry_bytes, exact_total, max_iwa_bytes)?;
        assert_eq!(
            SemanticProjection::from_bytes_with_limits(
                &bytes,
                exact,
                SemanticProfile::ComponentsOnly,
            )?
            .components()
            .len(),
            2
        );

        // One byte below the two-component aggregate is rejected while
        // Snappy examines Bravo's declared decoded length, before a second
        // decoded stream or parsed archive can be retained.
        let one_over = Limits::new(
            input_bytes,
            2,
            max_entry_bytes,
            exact_total - 1,
            max_iwa_bytes,
        )?;
        assert!(matches!(
            SemanticProjection::from_bytes_with_limits(
                &bytes,
                one_over,
                SemanticProfile::ComponentsOnly,
            ),
            Err(Error::Limit {
                kind: LimitKind::IwaTotalBytes,
                observed,
                maximum,
            }) if observed == exact_total && maximum == exact_total - 1
        ));

        let exhausted = Limits::new(input_bytes, 2, max_entry_bytes, alpha_bytes, max_iwa_bytes)?;
        crate::zip::reset_test_entry_read_count();
        assert!(matches!(
            SemanticProjection::from_bytes_with_limits(
                &bytes,
                exhausted,
                SemanticProfile::ComponentsOnly,
            ),
            Err(Error::Limit {
                kind: LimitKind::IwaTotalBytes,
                observed,
                maximum,
            }) if observed == alpha_bytes + 1 && maximum == alpha_bytes
        ));
        assert_eq!(crate::zip::test_entry_read_count(), 1);
        Ok(())
    }

    #[test]
    fn legacy_projection_normalizes_exact_root_components() -> Result<()> {
        let document = iwa(1)?;
        let table = iwa(2)?;
        let nested = zip(&[("Document.iwa", &document), ("Tables.iwa", &table)])?;
        let outer = zip(&[("legacy.pages/Index.zip", nested.as_slice())])?;

        let projection = SemanticProjection::from_bytes_with_limits(
            &outer,
            Limits::default(),
            SemanticProfile::ComponentsOnly,
        )?;

        assert_eq!(projection.components().len(), 2);
        assert!(projection.components().get("Index/Document.iwa").is_some());
        assert!(projection.components().get("Index/Tables.iwa").is_some());
        Ok(())
    }

    #[test]
    fn legacy_projection_rejects_nested_and_hostile_iwa_authorities() -> Result<()> {
        let document = iwa(1)?;
        for hostile_name in ["Index/Document.iwa", "Tables/Tile.iwa"] {
            let nested = zip(&[(hostile_name, &document)])?;
            let outer = zip(&[("legacy.pages/Index.zip", nested.as_slice())])?;

            crate::zip::reset_test_entry_read_count();
            assert!(matches!(
                SemanticProjection::from_bytes_with_limits(
                    &outer,
                    Limits::default(),
                    SemanticProfile::ComponentsOnly,
                ),
                Err(Error::InvalidBundle(_))
            ));
            // The outer container must be read to validate its central directory,
            // but no hostile inner payload is inflated.
            assert_eq!(crate::zip::test_entry_read_count(), 1);
        }

        let mut nested = zip(&[("Document.iwa", &document)])?;
        replace_raw_name(&mut nested, b"Document.iwa", b"Docu\\ent.iwa");
        let outer = zip(&[("legacy.pages/Index.zip", nested.as_slice())])?;
        crate::zip::reset_test_entry_read_count();
        assert!(matches!(
            SemanticProjection::from_bytes_with_limits(
                &outer,
                Limits::default(),
                SemanticProfile::ComponentsOnly,
            ),
            Err(Error::InvalidBundle(_))
        ));
        assert_eq!(crate::zip::test_entry_read_count(), 1);
        Ok(())
    }
}
