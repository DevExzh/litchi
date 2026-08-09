//! Bounded package reconstruction primitives shared by ODF family editors.

use std::collections::HashSet;

use crate::constants;
use crate::core::{
    AuthoredXmlFragment, OwnedPackage, PackageWriter, XmlSourcePart, XmlSplicePublication,
};
use litchi_core::{Error, Result};

const MAX_CONTENT_BYTES: usize = 16 * 1024 * 1024;
const MAX_ADDITION_BYTES: usize = 64 * 1024 * 1024;

/// One validated package member to add during an atomic rebuild.
#[derive(Debug, Clone)]
pub struct Addition {
    pub path: String,
    pub bytes: Vec<u8>,
    pub media_type: String,
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
    if content.len() > MAX_CONTENT_BYTES {
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
    let source_part = XmlSourcePart::load(source, constants::ODF_CONTENT)?;
    let source_xml = std::str::from_utf8(source_part.bytes()).map_err(|error| {
        Error::InvalidFormat(format!("invalid UTF-8 in source content.xml: {error}"))
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
