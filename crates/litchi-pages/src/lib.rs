//! Archive-free Pages semantic models.
//!
//! Package parsing, object topology, and mutation remain owned by the Pages
//! implementation. This crate owns the immutable semantic document and its
//! concise section vocabulary so callers do not need to import IWA archives,
//! protobuf schemas, or package identifiers for ordinary document reads.

#![forbid(unsafe_code)]

mod document;
pub mod header_footer;
mod section;

pub use document::{
    Body, DEFAULT_MAX_TEXT_BYTES, Document, Error, MAX_BODY_STORAGES, MAX_SECTIONS, Result, Root,
};
pub use section::{Section, SectionType};
