//! Neutral image-resource authoring primitives.
//!
//! The resource values here own only bounded payloads and safe package paths.
//! Manifest mutation and document-family package transactions remain in the
//! owning format crate.

mod model;

pub use model::{Format, Part, allocate_picture_path, validate_payload};
