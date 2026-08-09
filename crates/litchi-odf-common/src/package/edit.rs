//! Bounded package reconstruction primitives shared by ODF family editors.

use std::collections::HashSet;

use crate::constants;
use crate::core::{OwnedPackage, PackageWriter};
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
        writer.add_file(constants::ODF_CONTENT, content.as_bytes())?;
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
