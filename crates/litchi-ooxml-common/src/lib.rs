//! Shared OOXML functionality that is independent of DOCX, PPTX, XLSX, and XLSB.

#![forbid(unsafe_code)]
// Shared OOXML models deliberately retain schema vocabulary and ownership.
// Retrofitting generic API heuristics here would create breaking changes for
// every format facade without altering the bounded wire behavior.
#![allow(
    clippy::borrowed_box,
    clippy::module_name_repetitions,
    clippy::ref_option,
    clippy::struct_excessive_bools,
    reason = "shared public OOXML types mirror ECMA-376 schemas and document their common failure contract at module boundaries"
)]
// Streaming namespace and MCE parsers reuse short event-local bindings as raw
// values become expanded names and semantic nodes.
#![allow(
    clippy::many_single_char_names,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::similar_names,
    reason = "short-lived parser bindings track successive XML events and namespace-expanded projections"
)]
// Codec declarations follow XML/package traversal order, and wildcard imports
// keep generated vocabulary modules auditable against their schemas.
#![allow(
    clippy::arbitrary_source_item_ordering,
    clippy::unreadable_literal,
    clippy::wildcard_imports,
    reason = "shared OOXML codecs retain schema and traversal order while generated vocabulary modules use bounded local preludes"
)]

mod binding_tracker;
mod error;

pub mod custom;
pub mod custom_xml;
pub mod embedded;
pub mod external_link;
pub mod mce;
#[cfg(feature = "encryption")]
pub mod package_encryption;
pub mod properties;
pub mod relationships;
pub mod ribbon;
pub mod spreadsheet_xml_maps;
#[cfg(feature = "vba-inspection")]
pub mod vba;
pub mod web;
pub mod xml;
pub mod xml_name;

/// Hidden, unstable cross-owner implementation plumbing.
///
/// Format crates use this narrow namespace for borrowing scanners that share
/// OOXML wire behavior. It is deliberately excluded from the ordinary facade
/// vocabulary, carries no compatibility promise, and is not a public semantic
/// model.
#[doc(hidden)]
pub mod private {
    pub use super::binding_tracker::{BindingTracker, BindingTrackerError};
}

pub use error::{Error, Result};
pub use properties::{Keywords, Props};
pub use xml::XmlError;
