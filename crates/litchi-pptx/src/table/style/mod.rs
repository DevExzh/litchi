//! Typed, bounded DrawingML table-style catalogs and transactional package CRUD.

mod codec;
mod model;
mod validation;

#[cfg(test)]
mod tests;

pub use model::{Conformance, Def, Id, Link, List, Parts};
pub use validation::{conformance, link, load, present, put, remove};

use crate::Error;

const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const AS: &str = "http://purl.oclc.org/ooxml/drawingml/main";
const P: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const PS: &str = "http://purl.oclc.org/ooxml/presentationml/main";
const STRICT_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/tableStyles";
const DEFAULT_XML: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
    r#"<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" "#,
    r#"def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>"#,
);

const MAX_XML_BYTES: usize = 16 * 1024 * 1024;
const MAX_PRESENTATION_BYTES: usize = 32 * 1024 * 1024;
const MAX_NODES: usize = 250_000;
const MAX_DEPTH: usize = 128;
const MAX_ATTRIBUTES: usize = 64;
const MAX_ATTRIBUTE_BYTES: usize = 4_096;
const MAX_STYLES: usize = 4_096;
const MAX_GRAPH_PARTS: usize = 65_536;
const MAX_GRAPH_RELATIONSHIPS: usize = 262_144;
const PART_NAME_ATTEMPTS: usize = 4_096;
const RELATIONSHIP_ID_ATTEMPTS: usize = 65_536;

fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(error.to_string())
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn limit(resource: &'static str, limit: usize) -> Error {
    Error::Limit { resource, limit }
}

fn allocation(resource: &'static str, source: std::collections::TryReserveError) -> Error {
    Error::Allocation { resource, source }
}

/// Return the deterministic Transitional default catalog used by new decks.
pub fn default_xml() -> &'static str {
    DEFAULT_XML
}
