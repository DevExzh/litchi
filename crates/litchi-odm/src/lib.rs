//! `OpenDocument` Master Document support with semantic responsibility layers.
#![forbid(unsafe_code)]
#![allow(
    clippy::arbitrary_source_item_ordering,
    clippy::case_sensitive_file_extension_comparisons,
    clippy::shadow_reuse,
    reason = "master-document codecs follow ODF traversal order, preserve case-sensitive package identity, and reuse event-local parser names"
)]

mod authoring;
mod codec;
mod edit_ops;
mod facade;
pub mod link;
pub mod metadata;
mod model;
mod package;
pub mod title;
pub mod transaction;

pub use facade::{Builder, Document, Package};
pub use litchi_odf_common::{Error, LimitKind, OpenOptions};
pub use model::{resource, section, security, structure, style, subdocument};
