//! Shared ODF package and manifest primitives.
//!
//! This layer owns archive access and the neutral part of `manifest.xml`:
//! file paths, media types, and declared sizes. Password encryption metadata,
//! signature handling, and document-family orchestration remain in
//! `litchi-odf`.

mod codec;
pub mod edit;
mod model;
mod path;
#[cfg(test)]
mod tests;

pub use codec::{is_media_path, parse_manifest, read_manifest};
#[allow(
    clippy::module_name_repetitions,
    reason = "The established public API explicitly distinguishes whole-package rebuilding."
)]
pub use edit::{
    Addition, MAX_CONTENT_REPLACEMENT_BYTES, content_splice_publication, raw_identical_members,
    rebuild_package, replace_content_xml, replace_content_xml_spliced,
    replace_content_xml_with_payload_verification, splice, xml_splice_publication,
};
pub(crate) use model::PreparedArchive;
pub use model::{Archive, ArchiveLimits, ArchiveMetadata, ArchiveNames, Entry, Manifest};
#[cfg(test)]
pub(crate) use model::{index_build_count, note_index_build, reset_index_build_count};
pub(crate) use path::validate_manifest_path;
pub use path::{is_linked_href, resolve_package_path};

/// Borrowed package metadata used by read-only XML inventory scanners.
///
/// The scanner only needs to classify safe, package-local references. It does
/// not read or retain archive bytes, so format-family crates can provide a
/// lightweight view over their package without coupling this crate to a
/// concrete archive implementation.
#[allow(
    clippy::module_name_repetitions,
    reason = "The public trait name makes its archive lookup role clear at call sites."
)]
pub trait PackageLookup {
    /// Return whether `path` is present in the package archive.
    fn has_file(&self, path: &str) -> bool;

    /// Return the manifest media type for `path`, if one is declared.
    fn media_type(&self, path: &str) -> Option<&str>;
}
