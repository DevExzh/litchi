//! Bounded package reconstruction primitives shared by ODF family editors.

use std::{
    collections::{BTreeMap, BTreeSet, HashSet},
    ops::Range,
};

use crate::constants;
use crate::core::{
    AuthoredXmlFragment, OwnedPackage, PackageWriter, XmlSourcePart, XmlSplicePublication,
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
/// Unsupported physical layouts, signatures, encryption, and size-bearing
/// content manifest entries use the established logical rebuild instead. An
/// exact semantic no-op returns the accepted source bytes unchanged.
///
/// # Errors
///
/// Returns an error when the replacement is oversized, invalid XML, or cannot
/// be published through either the preserving or established rebuild path.
pub fn replace_content_xml(source: &OwnedPackage, content: &str) -> Result<Vec<u8>> {
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
        if let Some(bytes) = try_preserve_content_replacement(source, replacement) {
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

fn try_preserve_content_replacement(
    source: &OwnedPackage,
    replacement: Vec<u8>,
) -> Option<Vec<u8>> {
    let mut replacement = Some(replacement);
    let package = source.package().ok()?;
    if package.manifest().has_encrypted_entries()
        || package
            .manifest()
            .get_entry(constants::ODF_CONTENT)
            .is_some_and(|entry| entry.size.is_some())
        || package
            .files()
            .ok()?
            .iter()
            .any(|path| path.starts_with("META-INF/") && path.ends_with("signatures.xml"))
    {
        return None;
    }

    let archive = ZipArchive::from_slice(source.as_bytes())
        .ok()?
        .into_zip_archive();
    let mut buffer = vec![0_u8; soapberry_zip::RECOMMENDED_BUFFER_SIZE];
    let index = PreservationIndex::new(&archive, &mut buffer).ok()?;
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
