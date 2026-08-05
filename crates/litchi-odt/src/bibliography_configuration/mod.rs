//! Typed, layered OpenDocument bibliography-configuration ownership.

const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const FO: &str = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
const STYLE: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const MAX_XML_BYTES: usize = 64 * 1_048_576;
const MAX_DEPTH: usize = 256;
const MAX_SORT_KEYS: usize = 256;
const MAX_VALUE_BYTES: usize = 65_536;
const MAX_AGGREGATE_BYTES: usize = 4 * 1_048_576;

mod codec;
mod model;
mod package;

#[cfg(test)]
mod tests;

pub use model::{Configuration, Field, SortKey};

pub(crate) use codec::{
    parse_bibliography_configuration, parse_bibliography_configuration_parts,
    remove_bibliography_configuration_xml, set_bibliography_configuration_xml,
};
