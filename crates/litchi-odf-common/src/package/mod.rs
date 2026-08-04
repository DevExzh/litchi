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
