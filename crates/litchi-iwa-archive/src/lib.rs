//! Bounded physical ingress for Apple iWork bundles.
//!
//! This crate owns only the ZIP container boundary: central-directory limits,
//! legacy nested `Index.zip` handling, and the checksum-free Snappy/IWA
//! component stream. It deliberately does not depend
//! on the iWork facade or semantic format crates. Already-decoded logical
//! package entries enter through an explicitly immutable, fully revalidated
//! snapshot.
//!
//! Application readers should consume [`SourceCatalog`] when they need both
//! physical members and parsed components, or [`ComponentCatalog::iter`] for
//! component-only ingress. Raw ZIP implementation types remain private.

#![forbid(unsafe_code)]

mod catalog;
mod directory;
mod error;
mod limits;
mod logical;
pub mod package;
mod zip;

pub use catalog::{Component, ComponentCatalog, SourceCatalog};
pub use directory::{
    DirectoryMarkers, DirectoryProvenance, FrozenDirectoryBundle, FrozenDirectoryEntry,
};
pub use error::{Error, LimitKind, Result};
pub use limits::Limits;
pub use logical::LogicalSourceCatalog;
pub use package::SourceProvenance;

use litchi_iwa_core::{Archive, SnappyStream};
use soapberry_zip::office::ArchiveReader;

/// The root evidence needed by format detection.
///
/// This projection deliberately contains no ZIP implementation types. The
/// facade inspects the physical bundle and parses only `Document.iwa`, while
/// leaving unrelated `.iwa` members opaque to format detection.
#[derive(Debug)]
pub struct DetectionRoot {
    has_iwa_components: bool,
    has_keynote_components: bool,
    document: Option<Archive>,
}

impl DetectionRoot {
    /// Whether the inspected archive contains direct IWA components.
    #[must_use]
    pub const fn has_iwa_components(&self) -> bool {
        self.has_iwa_components
    }

    /// Whether the inspected archive contains a Keynote slide component.
    #[must_use]
    pub const fn has_keynote_components(&self) -> bool {
        self.has_keynote_components
    }

    /// Borrow the parsed root `Document.iwa`, if present.
    #[must_use]
    pub const fn document(&self) -> Option<&Archive> {
        self.document.as_ref()
    }
}

/// Inspect one packaged iWork input for format detection.
///
/// ZIP traversal, encryption markers, legacy nested `Index.zip`, Snappy
/// framing, and neutral IWA limits remain inside this facade. Only the root
/// `Document.iwa` is parsed; malformed unrelated components do not alter the
/// historical detection result.
///
/// # Errors
///
/// Returns a typed error when the ZIP, nested index, encryption marker,
/// Snappy stream, IWA framing, or configured physical limit is invalid.
pub fn inspect_detection_root(bytes: &[u8], limits: Limits) -> Result<DetectionRoot> {
    let validated_limits = limits.validate()?;
    let input_size = u64::try_from(bytes.len())
        .map_err(|_error| Error::InvalidBundle("ZIP input length does not fit u64".to_owned()))?;
    validated_limits.check_input_size(input_size, "ZIP input")?;
    let archive = ArchiveReader::new_with_limits(bytes, validated_limits.zip_limits())?;
    inspect_zip(&archive, validated_limits, true)
}

fn inspect_zip(
    archive: &ArchiveReader<'_>,
    limits: Limits,
    allow_nested: bool,
) -> Result<DetectionRoot> {
    if is_encrypted(archive) {
        return Err(Error::Encrypted);
    }

    let has_direct_iwa = archive.file_names().any(is_iwa_name);
    let nested_name = nested_index_name(archive)?;
    if has_direct_iwa && nested_name.is_some() {
        return Err(Error::InvalidBundle(
            "iWork package mixes direct IWA members with a legacy Index.zip".to_owned(),
        ));
    }
    if has_direct_iwa {
        return inspect_direct_zip(archive, limits);
    }

    if !allow_nested {
        return Ok(empty_detection_root());
    }
    let Some(index_name) = nested_name else {
        return Ok(empty_detection_root());
    };
    let declared_index_size = archive.metadata(&index_name)?.uncompressed_size();
    limits.check_input_size(declared_index_size, "legacy iWork Index.zip")?;
    let index_data = archive.read(&index_name)?;
    let index_size = u64::try_from(index_data.len()).map_err(|_error| {
        Error::InvalidBundle("legacy iWork Index.zip length does not fit u64".to_owned())
    })?;
    limits.check_input_size(index_size, "legacy iWork Index.zip")?;
    let index = ArchiveReader::new_with_limits(&index_data, limits.zip_limits())?;
    inspect_zip(&index, limits, false)
}

fn empty_detection_root() -> DetectionRoot {
    DetectionRoot {
        has_iwa_components: false,
        has_keynote_components: false,
        document: None,
    }
}

fn inspect_direct_zip(archive: &ArchiveReader<'_>, limits: Limits) -> Result<DetectionRoot> {
    let mut document_names = archive
        .file_names()
        .filter(|name| index_name(name) == Some("Document.iwa"));
    let document_name = document_names.next().map(str::to_owned);
    if document_names.next().is_some() {
        return Err(Error::InvalidBundle(
            "iWork package contains multiple Document.iwa components".to_owned(),
        ));
    }

    let document = document_name
        .as_deref()
        .map(|name| parse_document(archive, name, limits))
        .transpose()?;
    Ok(DetectionRoot {
        has_iwa_components: true,
        has_keynote_components: archive.file_names().any(|name| {
            is_component(name, "MasterSlide")
                || is_component(name, "Slide")
                || is_component(name, "TemplateSlide")
        }),
        document,
    })
}

fn parse_document(archive: &ArchiveReader<'_>, name: &str, limits: Limits) -> Result<Archive> {
    let compressed = archive.read(name)?;
    let stream = SnappyStream::decompress_with_limits(&compressed, limits.snappy_limits()?)?;
    Ok(Archive::parse_with_limits(
        stream.as_bytes(),
        limits.effective_archive_limits()?,
    )?)
}

fn is_encrypted(archive: &ArchiveReader<'_>) -> bool {
    archive
        .file_names()
        .any(|name| matches!(name.rsplit('/').next(), Some(".iwpv2" | ".iwph")))
}

fn nested_index_name(archive: &ArchiveReader<'_>) -> Result<Option<String>> {
    let mut candidates = archive
        .file_names()
        .filter(|name| name.rsplit('/').next() == Some("Index.zip"));
    let first = candidates.next().map(str::to_owned);
    if let Some(second) = candidates.next() {
        return Err(Error::InvalidBundle(format!(
            "iWork package contains ambiguous nested indexes: {} and {second}",
            first.as_deref().unwrap_or("Index.zip")
        )));
    }
    Ok(first)
}

fn index_name(name: &str) -> Option<&str> {
    name.strip_prefix("Index/")
        .or_else(|| (!name.contains('/')).then_some(name))
}

fn is_iwa_name(name: &str) -> bool {
    #[allow(
        clippy::case_sensitive_file_extension_comparisons,
        reason = "IWA member names are case-sensitive protocol names."
    )]
    {
        name.ends_with(".iwa")
    }
}

fn is_component(name: &str, stem: &str) -> bool {
    let Some(component_name) =
        index_name(name).and_then(|candidate| candidate.strip_suffix(".iwa"))
    else {
        return false;
    };
    let Some(suffix) = component_name.strip_prefix(stem) else {
        return false;
    };
    suffix.is_empty()
        || suffix.strip_prefix('-').is_some_and(|version| {
            !version.is_empty()
                && version
                    .split('-')
                    .all(|part| !part.is_empty() && part.bytes().all(|byte| byte.is_ascii_digit()))
        })
}
