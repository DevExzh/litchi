//! Shared ODF package and manifest primitives.
//!
//! This layer owns archive access and the neutral part of `manifest.xml`:
//! file paths, media types, and declared sizes. Password encryption metadata,
//! signature handling, and document-family orchestration remain in
//! `litchi-odf`.

mod codec;
mod model;
#[cfg(test)]
mod tests;

pub use codec::{is_media_path, parse_manifest, read_manifest};
pub use model::{Archive, Entry, Manifest};
