//! Shared OOXML functionality that is independent of DOCX, PPTX, XLSX, and XLSB.

#![forbid(unsafe_code)]
// quick-xml's checked attribute iteration is quadratic on hostile tags; read
// attributes through `BytesStartExt` (workspace `clippy.toml`).
#![cfg_attr(not(test), deny(clippy::disallowed_methods))]
// Shared OOXML models deliberately retain schema vocabulary and ownership.
// Retrofitting generic API heuristics here would create breaking changes for
// every format facade without altering the bounded wire behavior.
#![allow(
    clippy::borrowed_box,
    reason = "shared public OOXML types mirror ECMA-376 schemas and document their common failure contract at module boundaries"
)]

mod binding_tracker;
mod error;
mod signature;

pub mod custom;
pub mod custom_data;
pub mod custom_xml;
pub mod embedded;
pub mod external_link;
pub mod leniency;
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
    pub use super::binding_tracker::{
        BindingTracker, BindingTrackerError, in_scope_declarations, split_qualified_name,
        with_in_scope_namespaces,
    };
}

pub use error::{Error, LimitKind, Result};

pub use xml::XmlError;
