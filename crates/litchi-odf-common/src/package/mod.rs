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
pub use edit::{Addition, rebuild_package, splice};
pub use model::{Archive, Entry, Manifest};
pub use path::{is_linked_href, resolve_package_path};

/// Borrowed package metadata used by read-only XML inventory scanners.
///
/// The scanner only needs to classify safe, package-local references. It does
/// not read or retain archive bytes, so format-family crates can provide a
/// lightweight view over their package without coupling this crate to a
/// concrete archive implementation.
pub trait PackageLookup {
    /// Return whether `path` is present in the package archive.
    fn has_file(&self, path: &str) -> bool;

    /// Return the manifest media type for `path`, if one is declared.
    fn media_type(&self, path: &str) -> Option<&str>;
}
