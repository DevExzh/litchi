//! Inert, bounded OpenDocument text DDE connection declarations.
//!
//! The owner keeps its public declaration and reference model separate from
//! the XML codec, package-wide aggregation, and focused regression tests.

mod codec;
mod model;
mod package;

#[cfg(test)]
mod tests;

pub(super) const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
pub(super) const STYLE: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
pub(super) const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
pub(super) const MAX_XML_BYTES: usize = 64 * 1_048_576;
pub(super) const MAX_DEPTH: usize = 256;
pub(super) const MAX_CONNECTIONS: usize = 65_536;
pub(super) const MAX_REFERENCES: usize = 262_144;
pub(super) const MAX_VALUE_BYTES: usize = 65_536;
pub(super) const MAX_AGGREGATE_BYTES: usize = 16 * 1_048_576;

pub use model::{Declaration, Use};
pub(crate) use package::parse_dde_connection_parts;
