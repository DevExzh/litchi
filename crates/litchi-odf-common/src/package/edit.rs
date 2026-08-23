//! Bounded package reconstruction primitives shared by ODF family editors.

use std::{
    collections::{BTreeMap, BTreeSet, HashSet},
    ops::Range,
};

use crate::constants;
use crate::core::{
    AuthoredXmlFragment, OwnedPackage, PackageWriter, XmlSourcePart, XmlSplicePublication,
    is_signature_owner_path,
};
use litchi_core::{Error, Result};
use soapberry_zip::{
    CompressionMethod, PreservationAction, PreservationIndex, PreservationPlan, RegeneratedEntry,
    ZipArchive,
};

/// Maximum `content.xml` size accepted by the common bounded replacement and
/// rebuild helpers.
///
/// Format owners with a larger established publication envelope may use this
/// value to select their existing writer fallback before calling these
/// helpers.
pub const MAX_CONTENT_REPLACEMENT_BYTES: usize = 16 * 1024 * 1024;
const MAX_ADDITION_BYTES: usize = 64 * 1024 * 1024;
const CENTRAL_LOCAL_HEADER_OFFSET: Range<usize> = 42..46;
const ZIP_LOCAL_HEADER_SIZE: usize = 30;
const ZIP_CENTRAL_HEADER_SIZE: usize = 46;
const ZIP_UTF8_FLAG: u16 = 1 << 11;
const ZIP_DATA_DESCRIPTOR_FLAG: u16 = 1 << 3;

#[derive(Clone, Debug)]
struct RawMemberRanges {
    local: Range<usize>,
    central: Range<usize>,
}

/// One validated package member to add during an atomic rebuild.
#[derive(Debug, Clone)]
pub struct Addition {
    pub path: String,
    pub bytes: Vec<u8>,
    pub media_type: String,
}

/// Identify members with an exact physical representation in two ordinary ZIP archives.
///
/// Local-member bytes must match exactly. Central-directory records must also match except for
/// the local-header offset that necessarily changes when an earlier member changes length. The
/// result is an optimization hint for callers that retain a logical comparison fallback; `None`
/// means either archive has an unsupported layout, unsafe or duplicate paths, or invalid ranges.
/// Member bodies are never decompressed.
#[must_use]
pub fn raw_identical_members(source: &[u8], target: &[u8]) -> Option<BTreeSet<String>> {
    let source_members = raw_member_ranges(source)?;
    let target_members = raw_member_ranges(target)?;
    Some(
        source_members
            .iter()
            .filter_map(|(name, source_member)| {
                let target_member = target_members.get(name)?;
                raw_member_is_identical(source, source_member, target, target_member)
                    .then(|| name.clone())
            })
            .collect(),
    )
}

fn raw_member_ranges(bytes: &[u8]) -> Option<BTreeMap<String, RawMemberRanges>> {
    let archive = ZipArchive::from_slice(bytes).ok()?.into_zip_archive();
    let mut buffer = vec![0_u8; soapberry_zip::RECOMMENDED_BUFFER_SIZE];
    let index = PreservationIndex::new(&archive, &mut buffer).ok()?;
    let mut records = archive.entries(&mut buffer);
    let mut members = BTreeMap::new();
    for preserved in index.entries() {
        let record = records.next_entry().ok()??;
        let normalized = record.file_path().try_normalize().ok()?;
        let ranges = RawMemberRanges {
            local: checked_range(preserved.local_span(), bytes.len())?,
            central: checked_range(preserved.central_record(), bytes.len())?,
        };
        if members
            .insert(normalized.as_ref().to_string(), ranges)
            .is_some()
        {
            return None;
        }
    }
    if records.next_entry().ok()?.is_some() {
        return None;
    }
    Some(members)
}

fn checked_range(range: Range<u64>, length: usize) -> Option<Range<usize>> {
    let start = usize::try_from(range.start).ok()?;
    let end = usize::try_from(range.end).ok()?;
    (start <= end && end <= length).then_some(start..end)
}

fn raw_member_is_identical(
    source: &[u8],
    source_member: &RawMemberRanges,
    target: &[u8],
    target_member: &RawMemberRanges,
) -> bool {
    if source[source_member.local.clone()] != target[target_member.local.clone()] {
        return false;
    }
    let source_central = &source[source_member.central.clone()];
    let target_central = &target[target_member.central.clone()];
    source_central.len() == target_central.len()
        && source_central.len() >= CENTRAL_LOCAL_HEADER_OFFSET.end
        && source_central[..CENTRAL_LOCAL_HEADER_OFFSET.start]
            == target_central[..CENTRAL_LOCAL_HEADER_OFFSET.start]
        && source_central[CENTRAL_LOCAL_HEADER_OFFSET.end..]
            == target_central[CENTRAL_LOCAL_HEADER_OFFSET.end..]
}

/// Replace only `content.xml`, raw-copying every other source ZIP member when
/// the exact package layout and manifest make that preservation safe.
///
/// Unsupported physical layouts, signatures, and encryption use the
/// established logical rebuild. The package writer refuses unencrypted
/// `manifest:size` metadata with a typed [`Error::Unsupported`] error rather
/// than silently dropping it. An exact semantic no-op returns the accepted
/// source bytes unchanged.
///
/// # Errors
///
/// Returns an error when the replacement is oversized, invalid XML, or cannot
/// be published through either the preserving or established rebuild path.
pub fn replace_content_xml(source: &OwnedPackage, content: &str) -> Result<Vec<u8>> {
    replace_content_xml_with_mode(source, content, false)
}

/// Replace `content.xml` while opting into verification of every untouched
/// source payload before raw preservation.
///
/// The exact semantic no-op still returns the source bytes before this
/// verification is considered. Callers that need to validate every member
/// during a changed-content publication should use this entry point; the
/// ordinary shared helper intentionally keeps the raw-copy fast path lazy for
/// opaque media members.
///
/// # Errors
///
/// Returns an error when the replacement is oversized, invalid XML, or cannot
/// be published through either the preserving or established rebuild path.
pub fn replace_content_xml_with_payload_verification(
    source: &OwnedPackage,
    content: &str,
) -> Result<Vec<u8>> {
    replace_content_xml_with_mode(source, content, true)
}

fn replace_content_xml_with_mode(
    source: &OwnedPackage,
    content: &str,
    verify_payloads: bool,
) -> Result<Vec<u8>> {
    if content.len() > MAX_CONTENT_REPLACEMENT_BYTES {
        return invalid("outer content.xml exceeds package mutation limit");
    }
    if source.get_file(constants::ODF_CONTENT)? == content.as_bytes() {
        return Ok(source.as_bytes().to_vec());
    }

    if let Ok(publication) = content_splice_publication(source, content) {
        let (path, replacement, _media_type) = publication.assemble()?;
        if path != constants::ODF_CONTENT || replacement != content.as_bytes() {
            return invalid("checked content.xml splice assembled unexpected bytes");
        }
        if let Some(bytes) = try_preserve_content_replacement(source, replacement, verify_payloads)
        {
            return Ok(bytes);
        }
    }

    rebuild_package(
        source,
        content,
        Vec::new(),
        Vec::new(),
        Vec::<String>::new(),
        Vec::<String>::new(),
    )
}

/// Replace `content.xml` from an already checked source-provenance splice.
///
/// Format owners that already know the exact edited source ranges can retain
/// that proof instead of deriving one maximal byte diff from the assembled
/// part. Every untouched ZIP member follows the same raw-preservation gate and
/// logical fallback as [`replace_content_xml`].
///
/// # Errors
///
/// Returns an error when the publication belongs to another package, does not
/// assemble the expected `content.xml`, exceeds the replacement bound, or
/// cannot be published through the preserving or established rebuild path.
pub fn replace_content_xml_spliced(
    source: &OwnedPackage,
    content: &str,
    publication: XmlSplicePublication,
) -> Result<Vec<u8>> {
    if content.len() > MAX_CONTENT_REPLACEMENT_BYTES {
        return invalid("outer content.xml exceeds package mutation limit");
    }
    if !publication.belongs_to(source) {
        return invalid("content.xml splice publication has different package provenance");
    }
    let (path, replacement, _media_type) = publication.assemble()?;
    if path != constants::ODF_CONTENT || replacement != content.as_bytes() {
        return invalid("checked content.xml splice assembled unexpected bytes");
    }
    if source.get_file(constants::ODF_CONTENT)? == replacement {
        return Ok(source.as_bytes().to_vec());
    }
    if let Some(bytes) = try_preserve_content_replacement(source, replacement, false) {
        return Ok(bytes);
    }
    rebuild_package(
        source,
        content,
        Vec::new(),
        Vec::new(),
        Vec::<String>::new(),
        Vec::<String>::new(),
    )
}

fn try_preserve_content_replacement(
    source: &OwnedPackage,
    replacement: Vec<u8>,
    verify_payloads: bool,
) -> Option<Vec<u8>> {
    let mut replacement = Some(replacement);
    let package = source.package().ok()?;
    let files = package.files().ok()?;
    if package.manifest().entries.keys().any(|path| {
        path != constants::ODF_CONTENT
            && soapberry_zip::path::ZipFilePath::from_str(path.as_str()).as_str()
                == constants::ODF_CONTENT
    }) {
        // Match the archive reader's path normalization before preserving raw
        // bytes. A normalized alias (including traversal and backslash forms)
        // can carry stale manifest:size metadata without being visible through
        // the canonical lookup. Refuse raw preservation while retaining the
        // established logical rebuild, which regenerates the manifest.
        return None;
    }
    if package.manifest().has_encrypted_entries()
        || package
            .manifest()
            .get_entry(constants::ODF_CONTENT)
            .is_some_and(|entry| entry.size.is_some())
        || files.iter().any(|path| is_signature_owner_path(path))
    {
        return None;
    }

    let archive = ZipArchive::from_slice(source.as_bytes())
        .ok()?
        .into_zip_archive();
    let mut buffer = vec![0_u8; soapberry_zip::RECOMMENDED_BUFFER_SIZE];
    let index = PreservationIndex::new(&archive, &mut buffer).ok()?;
    if !canonical_odf_mimetype(source.as_bytes(), &package, &index) {
        return None;
    }
    if !validate_preserved_member_framing(source.as_bytes(), &index) {
        return None;
    }
    if verify_payloads {
        // Preservation copies compressed bytes without asking the ZIP writer
        // to verify them. The opt-in caller asks us to validate every source
        // file first so a malformed payload cannot take the raw path when the
        // established logical rebuild would have rejected the source.
        for path in &files {
            // `replace_content_xml*` has already read and verified content.xml
            // before reaching this helper; avoid inflating that member twice.
            if path != constants::ODF_CONTENT && !path.ends_with('/') {
                package.get_file(path).ok()?;
            }
        }
    }
    let mut records = archive.entries(&mut buffer);
    let mut plan = PreservationPlan::new();
    let mut replaced = false;
    for preserved in index.entries() {
        let record = records.next_entry().ok()??;
        let normalized = record.file_path().try_normalize().ok()?;
        let name = normalized.as_ref();
        if name == constants::ODF_CONTENT {
            if replaced || record.is_dir() {
                return None;
            }
            let compression = match record.compression_method() {
                CompressionMethod::Store => CompressionMethod::Store,
                CompressionMethod::Deflate => CompressionMethod::Deflate,
                _ => return None,
            };
            plan.push(PreservationAction::Regenerate {
                id: preserved.id(),
                entry: RegeneratedEntry::new(name, replacement.take()?)
                    .compression_method(compression),
            });
            replaced = true;
        } else {
            plan.push(PreservationAction::Copy(preserved.id()));
        }
    }
    if records.next_entry().ok()?.is_some() || !replaced {
        return None;
    }
    index.write_to(&plan, Vec::new()).ok()
}

fn validate_preserved_member_framing(
    source: &[u8],
    index: &PreservationIndex<'_, impl soapberry_zip::ReaderAt>,
) -> bool {
    for entry in index.entries() {
        let Some(local_span) = checked_range(entry.local_span(), source.len()) else {
            return false;
        };
        let Some(central_span) = checked_range(entry.central_record(), source.len()) else {
            return false;
        };
        let Some(local_fixed_end) = local_span.start.checked_add(ZIP_LOCAL_HEADER_SIZE) else {
            return false;
        };
        let Some(local_fixed) = source.get(local_span.start..local_fixed_end) else {
            return false;
        };
        let Some(central_fixed_end) = central_span.start.checked_add(ZIP_CENTRAL_HEADER_SIZE)
        else {
            return false;
        };
        let Some(central_fixed) = source.get(central_span.start..central_fixed_end) else {
            return false;
        };
        let Some(local_name_len) = le_u16(local_fixed, 26).map(usize::from) else {
            return false;
        };
        let Some(local_extra_len) = le_u16(local_fixed, 28).map(usize::from) else {
            return false;
        };
        let Some(payload_start) = local_span
            .start
            .checked_add(ZIP_LOCAL_HEADER_SIZE)
            .and_then(|offset| offset.checked_add(local_name_len))
            .and_then(|offset| offset.checked_add(local_extra_len))
        else {
            return false;
        };
        let Some(compressed_size) = le_u32(central_fixed, 20).map(u64::from) else {
            return false;
        };
        let Some(local_crc) = le_u32(local_fixed, 14) else {
            return false;
        };
        let Some(local_compressed) = le_u32(local_fixed, 18) else {
            return false;
        };
        let Some(local_uncompressed) = le_u32(local_fixed, 22) else {
            return false;
        };
        let Some(payload_end) =
            payload_start.checked_add(usize::try_from(compressed_size).ok().unwrap_or(usize::MAX))
        else {
            return false;
        };
        if payload_end > local_span.end {
            return false;
        }
        let Some(flags) = le_u16(central_fixed, 8) else {
            return false;
        };
        if flags & ZIP_DATA_DESCRIPTOR_FLAG == 0 {
            if payload_end != local_span.end {
                return false;
            }
            continue;
        }

        let Some(descriptor) = source.get(payload_end..local_span.end) else {
            return false;
        };
        let Some(expected_crc) = le_u32(central_fixed, 16) else {
            return false;
        };
        let Some(expected_uncompressed) = le_u32(central_fixed, 24).map(u64::from) else {
            return false;
        };
        let (descriptor_crc, descriptor_compressed, descriptor_uncompressed) =
            match descriptor.len() {
                12 if le_u32(descriptor, 0) != Some(0x0807_4b50) => (
                    le_u32(descriptor, 0),
                    le_u32(descriptor, 4).map(u64::from),
                    le_u32(descriptor, 8).map(u64::from),
                ),
                16 if le_u32(descriptor, 0) == Some(0x0807_4b50) => (
                    le_u32(descriptor, 4),
                    le_u32(descriptor, 8).map(u64::from),
                    le_u32(descriptor, 12).map(u64::from),
                ),
                20 if le_u32(descriptor, 0) != Some(0x0807_4b50) => (
                    le_u32(descriptor, 0),
                    le_u64(descriptor, 4),
                    le_u64(descriptor, 12),
                ),
                24 if le_u32(descriptor, 0) == Some(0x0807_4b50) => (
                    le_u32(descriptor, 4),
                    le_u64(descriptor, 8),
                    le_u64(descriptor, 16),
                ),
                _ => return false,
            };
        if descriptor_crc != Some(expected_crc)
            || descriptor_compressed != Some(compressed_size)
            || descriptor_uncompressed != Some(expected_uncompressed)
            || !((local_crc == 0 && local_compressed == 0 && local_uncompressed == 0)
                || (local_crc == expected_crc
                    && u64::from(local_compressed) == compressed_size
                    && u64::from(local_uncompressed) == expected_uncompressed))
        {
            return false;
        }
    }
    true
}

fn canonical_odf_mimetype(
    source: &[u8],
    package: &crate::core::package::Package<'_>,
    index: &PreservationIndex<'_, impl soapberry_zip::ReaderAt>,
) -> bool {
    let mimetype = package.mimetype();
    if !constants::is_odf_mime_type(mimetype)
        || package.manifest().mimetype != mimetype
        || index.entries().is_empty()
    {
        return false;
    }

    let Some((mimetype_index, entry)) = index
        .entries()
        .iter()
        .enumerate()
        .find(|(_, entry)| entry.raw_name_bytes() == b"mimetype")
    else {
        return false;
    };
    if mimetype_index != 0
        || index
            .entries()
            .iter()
            .filter(|entry| entry.raw_name_bytes() == b"mimetype")
            .count()
            != 1
        || entry.local_span().start != 0
    {
        return false;
    }

    let local_start = 0_usize;
    let Some(local_fixed) = source.get(local_start..local_start + ZIP_LOCAL_HEADER_SIZE) else {
        return false;
    };
    if le_u32(local_fixed, 0) != Some(0x0403_4b50) {
        return false;
    }
    let Some(local_flags) = le_u16(local_fixed, 6) else {
        return false;
    };
    let Some(local_method) = le_u16(local_fixed, 8) else {
        return false;
    };
    let Some(local_crc) = le_u32(local_fixed, 14) else {
        return false;
    };
    let Some(local_compressed_u32) = le_u32(local_fixed, 18) else {
        return false;
    };
    let Some(local_uncompressed) = le_u32(local_fixed, 22) else {
        return false;
    };
    let Some(local_name_len) = le_u16(local_fixed, 26).map(usize::from) else {
        return false;
    };
    let Some(local_extra_len) = le_u16(local_fixed, 28).map(usize::from) else {
        return false;
    };
    let Some(local_name_end) = ZIP_LOCAL_HEADER_SIZE.checked_add(local_name_len) else {
        return false;
    };
    let Some(local_payload_start) = local_name_end.checked_add(local_extra_len) else {
        return false;
    };
    let Some(local_compressed) = usize::try_from(local_compressed_u32).ok() else {
        return false;
    };
    let Some(local_payload_end) = local_payload_start.checked_add(local_compressed) else {
        return false;
    };
    let Some(local_span) = checked_range(entry.local_span(), source.len()) else {
        return false;
    };
    let Some(local_name) = source.get(ZIP_LOCAL_HEADER_SIZE..local_name_end) else {
        return false;
    };
    let Some(local_payload) = source.get(local_payload_start..local_payload_end) else {
        return false;
    };
    if local_name != b"mimetype"
        || local_flags & !(ZIP_UTF8_FLAG) != 0
        || local_flags & ZIP_DATA_DESCRIPTOR_FLAG != 0
        || local_method != CompressionMethod::Store.as_id().as_u16()
        || local_name_len != b"mimetype".len()
        || local_extra_len != 0
        || local_compressed_u32 != local_uncompressed
        || local_payload_end != local_span.end
        || local_crc != soapberry_zip::crc32(local_payload)
        || local_payload != mimetype.as_bytes()
    {
        return false;
    }

    let Some(central_span) = checked_range(entry.central_record(), source.len()) else {
        return false;
    };
    let Some(central_fixed_end) = central_span.start.checked_add(ZIP_CENTRAL_HEADER_SIZE) else {
        return false;
    };
    let Some(central_fixed) = source.get(central_span.start..central_fixed_end) else {
        return false;
    };
    if le_u32(central_fixed, 0) != Some(0x0201_4b50) {
        return false;
    }
    let Some(central_flags) = le_u16(central_fixed, 8) else {
        return false;
    };
    let Some(central_method) = le_u16(central_fixed, 10) else {
        return false;
    };
    let Some(central_crc) = le_u32(central_fixed, 16) else {
        return false;
    };
    let Some(central_compressed) = le_u32(central_fixed, 20) else {
        return false;
    };
    let Some(central_uncompressed) = le_u32(central_fixed, 24) else {
        return false;
    };
    let Some(central_name_len) = le_u16(central_fixed, 28).map(usize::from) else {
        return false;
    };
    let Some(central_extra_len) = le_u16(central_fixed, 30).map(usize::from) else {
        return false;
    };
    let Some(central_comment_len) = le_u16(central_fixed, 32).map(usize::from) else {
        return false;
    };
    let Some(central_name_end) = ZIP_CENTRAL_HEADER_SIZE.checked_add(central_name_len) else {
        return false;
    };
    let Some(central_end) = central_name_end
        .checked_add(central_extra_len)
        .and_then(|end| end.checked_add(central_comment_len))
    else {
        return false;
    };
    let Some(central_name_start) = central_span.start.checked_add(ZIP_CENTRAL_HEADER_SIZE) else {
        return false;
    };
    let Some(central_name_end) = central_span.start.checked_add(central_name_end) else {
        return false;
    };
    let Some(central_name) = source.get(central_name_start..central_name_end) else {
        return false;
    };
    central_span.len() == central_end
        && central_name == b"mimetype"
        && central_flags == local_flags
        && central_method == local_method
        && central_crc == local_crc
        && central_compressed == local_compressed_u32
        && central_uncompressed == local_uncompressed
        && central_name_len == local_name_len
        && central_extra_len == 0
        && central_compressed == u32::try_from(local_payload.len()).ok().unwrap_or(u32::MAX)
}

fn le_u16(bytes: &[u8], offset: usize) -> Option<u16> {
    bytes
        .get(offset..offset.checked_add(2)?)
        .and_then(|value| value.try_into().ok())
        .map(u16::from_le_bytes)
}

fn le_u32(bytes: &[u8], offset: usize) -> Option<u32> {
    bytes
        .get(offset..offset.checked_add(4)?)
        .and_then(|value| value.try_into().ok())
        .map(u32::from_le_bytes)
}

fn le_u64(bytes: &[u8], offset: usize) -> Option<u64> {
    bytes
        .get(offset..offset.checked_add(8)?)
        .and_then(|value| value.try_into().ok())
        .map(u64::from_le_bytes)
}

/// Rebuild an ODF package while replacing only the requested semantic parts.
///
/// # Errors
///
/// Returns an error when the replacement data exceeds a resource limit or the
/// source package cannot be copied into a valid rebuilt package.
#[allow(
    clippy::module_name_repetitions,
    reason = "The established public API distinguishes a whole-package rebuild from part edits."
)]
pub fn rebuild_package(
    source: &OwnedPackage,
    content: &str,
    additions: Vec<Addition>,
    directories: Vec<(String, String)>,
    excluded_paths: impl AsRef<[String]>,
    excluded_prefixes: impl AsRef<[String]>,
) -> Result<Vec<u8>> {
    if content.len() > MAX_CONTENT_REPLACEMENT_BYTES {
        return invalid("outer content.xml exceeds package mutation limit");
    }

    let excluded_path_list = excluded_paths.as_ref();
    let excluded_prefix_list = excluded_prefixes.as_ref();
    let source_content = source.get_file(constants::ODF_CONTENT)?;
    let content_is_exact_source = source_content == content.as_bytes();
    let source_package = source.package()?;
    let mut exact_exclusions = excluded_path_list.iter().cloned().collect::<HashSet<_>>();
    for path in source_package.files()? {
        if excluded_prefix_list
            .iter()
            .any(|prefix| path.starts_with(prefix))
        {
            exact_exclusions.insert(path);
        }
    }
    for path in source_package.manifest().entries.keys() {
        if excluded_prefix_list
            .iter()
            .any(|prefix| path.starts_with(prefix))
        {
            exact_exclusions.insert(path.clone());
        }
    }
    if !content_is_exact_source {
        exact_exclusions.insert(constants::ODF_CONTENT.to_string());
    }
    for addition in &additions {
        exact_exclusions.insert(addition.path.clone());
    }
    for (path, _) in &directories {
        exact_exclusions.insert(path.clone());
    }

    let mut writer = PackageWriter::new();
    writer.set_mimetype(&source.mimetype()?)?;
    if !content_is_exact_source {
        match content_splice_publication(source, content) {
            Ok(publication) => writer.add_spliced_xml(publication)?,
            Err(splice_error) => {
                // A changed whole part remains subject to the strict compact
                // authored-XML audit. Only a checked fine-grained splice may
                // retain formatting from a source-loaded producer part.
                writer
                    .add_file(constants::ODF_CONTENT, content.as_bytes())
                    .map_err(|audit_error| {
                        Error::InvalidFormat(format!(
                            "outer content.xml is neither a checked source splice ({splice_error}) nor compact authored XML ({audit_error})"
                        ))
                    })?;
            },
        }
    }
    for (path, media_type) in directories {
        writer.add_manifest_directory(&path, &media_type)?;
    }

    let mut addition_bytes = 0usize;
    for addition in additions {
        addition_bytes = addition_bytes
            .checked_add(addition.bytes.len())
            .ok_or_else(|| Error::InvalidFormat("package addition size overflow".to_string()))?;
        if addition_bytes > MAX_ADDITION_BYTES {
            return invalid("package additions exceed the size limit");
        }
        writer.add_file_with_media_type(&addition.path, &addition.bytes, &addition.media_type)?;
    }
    writer.copy_source_files_from_except(source, &exact_exclusions)?;
    writer.finish_to_bytes()
}

/// Derive one audited, provenance-bearing `content.xml` splice publication.
///
/// This is intentionally limited to a single byte-contiguous change against
/// the exact source package. Callers must fall back to whole-part compact XML
/// validation when the change cannot be classified this way.
///
/// # Errors
///
/// Returns an error when the source part is unavailable, the change is not a
/// single safely aligned splice, or the replacement fragment is not compact
/// authored XML.
pub fn content_splice_publication(
    source: &OwnedPackage,
    candidate: &str,
) -> Result<XmlSplicePublication> {
    xml_splice_publication(source, constants::ODF_CONTENT, candidate)
}

/// Derive one audited, provenance-bearing splice publication for an XML part.
///
/// Exact source bytes produce a zero-edit publication. A changed candidate
/// must differ by one safely aligned, compact authored fragment; all bytes
/// outside that fragment retain their source-package provenance.
///
/// # Errors
///
/// Returns an error when the source part is unavailable, the change is not a
/// single safely aligned splice, or the replacement fragment is not compact
/// authored XML.
pub fn xml_splice_publication(
    source: &OwnedPackage,
    path: &str,
    candidate: &str,
) -> Result<XmlSplicePublication> {
    let source_part = XmlSourcePart::load(source, path)?;
    let source_xml = std::str::from_utf8(source_part.bytes()).map_err(|error| {
        Error::InvalidFormat(format!("invalid UTF-8 in source {path}: {error}"))
    })?;
    let source_bytes = source_xml.as_bytes();
    let candidate_bytes = candidate.as_bytes();
    let mut prefix = source_bytes
        .iter()
        .zip(candidate_bytes)
        .take_while(|(left, right)| left == right)
        .count();
    while !source_xml.is_char_boundary(prefix) || !candidate.is_char_boundary(prefix) {
        prefix = prefix
            .checked_sub(1)
            .ok_or_else(|| Error::InvalidFormat("invalid XML splice prefix".to_string()))?;
    }
    let mut source_end = source_bytes.len();
    let mut candidate_end = candidate_bytes.len();
    while source_end > prefix
        && candidate_end > prefix
        && source_bytes[source_end - 1] == candidate_bytes[candidate_end - 1]
    {
        source_end -= 1;
        candidate_end -= 1;
    }
    while !source_xml.is_char_boundary(source_end) || !candidate.is_char_boundary(candidate_end) {
        source_end = source_end
            .checked_add(1)
            .filter(|end| *end <= source_bytes.len())
            .ok_or_else(|| Error::InvalidFormat("invalid XML splice suffix".to_string()))?;
        candidate_end = candidate_end
            .checked_add(1)
            .filter(|end| *end <= candidate_bytes.len())
            .ok_or_else(|| Error::InvalidFormat("invalid XML splice suffix".to_string()))?;
    }
    // A pure insertion immediately before a same-named sibling can make the
    // maximal common prefix end inside that sibling's start tag. Roll the
    // shared overlap back to the preceding markup boundary in both suffix
    // coordinates so the candidate delta is the complete inserted fragment
    // and the source proof remains an exact empty range.
    let provisional = candidate_bytes
        .get(prefix..candidate_end)
        .ok_or_else(|| Error::InvalidFormat("invalid candidate XML splice range".to_string()))?;
    if provisional.first() != Some(&b'<') && provisional.contains(&b'<') {
        let boundary = source_bytes[..prefix]
            .iter()
            .rposition(|byte| *byte == b'<')
            .ok_or_else(|| Error::InvalidFormat("missing XML splice boundary".to_string()))?;
        let overlap = prefix - boundary;
        source_end = source_end
            .checked_sub(overlap)
            .filter(|end| *end >= boundary)
            .ok_or_else(|| Error::InvalidFormat("invalid source XML splice overlap".to_string()))?;
        candidate_end = candidate_end
            .checked_sub(overlap)
            .filter(|end| *end >= boundary)
            .ok_or_else(|| {
                Error::InvalidFormat("invalid candidate XML splice overlap".to_string())
            })?;
        prefix = boundary;
    }
    let expected = source_bytes
        .get(prefix..source_end)
        .ok_or_else(|| Error::InvalidFormat("invalid source XML splice range".to_string()))?;
    let replacement = candidate_bytes
        .get(prefix..candidate_end)
        .ok_or_else(|| Error::InvalidFormat("invalid candidate XML splice range".to_string()))?;
    let proof = source_part.checked_range(prefix..source_end, expected)?;
    let fragment = if replacement.is_empty() {
        AuthoredXmlFragment::deletion()
    } else if replacement.first() == Some(&b'<') {
        AuthoredXmlFragment::markup(replacement.to_vec())
            .or_else(|_| AuthoredXmlFragment::start_tag(replacement.to_vec()))?
    } else {
        AuthoredXmlFragment::text(replacement.to_vec())?
    };
    let mut publication = XmlSplicePublication::new(source_part);
    publication.replace(proof, fragment)?;
    Ok(publication)
}

/// Replace a checked UTF-8 byte span without allocating intermediate XML trees.
///
/// # Errors
///
/// Returns an error when the requested byte range is reversed, outside `xml`,
/// or splits a UTF-8 code point.
pub fn splice(xml: &str, start: usize, end: usize, replacement: &str) -> Result<String> {
    if start > end || end > xml.len() || !xml.is_char_boundary(start) || !xml.is_char_boundary(end)
    {
        return invalid("invalid XML splice range");
    }
    let mut out = String::with_capacity(xml.len() - (end - start) + replacement.len());
    out.push_str(&xml[..start]);
    out.push_str(replacement);
    out.push_str(&xml[end..]);
    Ok(out)
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::InvalidFormat(message.into()))
}
