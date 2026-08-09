//! Inert `PresentationML` programmable tags and owner relationship discovery.
//!
//! Names are selected semantically and values are always inert strings. The
//! module never interprets a value as XML, a path, a command, or a relationship
//! target.
//!
//! ```
//! use litchi_pptx::tag::{List, Tag};
//!
//! let mut tags = List::new();
//! tags.add(Tag::new("Owner", "Alice")?)?;
//! assert_eq!(tags.get("owner")?.value(), "Alice");
//! assert!(tags.get(1_usize).is_err());
//! # Ok::<(), litchi_pptx::Error>(())
//! ```

const PML: &[u8] = b"http://schemas.openxmlformats.org/presentationml/2006/main";
const STRICT: &[u8] = b"http://purl.oclc.org/ooxml/presentationml/main";
const PML_TEXT: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const STRICT_TEXT: &str = "http://purl.oclc.org/ooxml/presentationml/main";
const TAG_REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags";
const STRICT_TAG_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/tags";
const REL_TEXT: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_REL_TEXT: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";

/// Content type of a `PresentationML` programmable-tag part.
pub const CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.tags+xml";

const MAX_PART_BYTES: usize = 8 * 1024 * 1024;
const MAX_TEXT_BYTES: usize = 1024 * 1024;
const MAX_TAGS: usize = 16_384;
const MAX_TAG_PARTS: usize = 1_024;
const MAX_OWNER_BYTES: usize = 64 * 1024 * 1024;
const MAX_GRAPH_PARTS: usize = 100_000;
const MAX_GRAPH_LINKS: usize = 1_000_000;
const MAX_PART_NAME_BYTES: usize = 64 * 1024 * 1024;
const PART_NAME_ATTEMPTS: usize = 10_000;
const MAX_SOURCE_RELATIONSHIPS: usize = 65_536;
const MAX_OWNER_NODES: usize = 1_000_000;
const MAX_OWNER_DEPTH: usize = 512;
const MAX_RELATIONSHIP_ID_BYTES: usize = 4_096;
const MAX_OWNER_MARKED_BYTES: usize = MAX_OWNER_BYTES * 2;
const P14: &str = "http://schemas.microsoft.com/office/powerpoint/2010/main";
const P15: &str = "http://schemas.microsoft.com/office/powerpoint/2012/main";
const XML_DECL: &[u8] = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#;
const ROOT_OPEN: &[u8] = b"<p:tagLst xmlns:p=\"";
const ROOT_EMPTY_CLOSE: &[u8] = b"/>";
const ROOT_CHILDREN_OPEN: &[u8] = b">";
const ROOT_CLOSE: &[u8] = b"</p:tagLst>";
const TAG_OPEN: &[u8] = b"<p:tag";
const TAG_CLOSE: &[u8] = b"/>";
const MAX_NAMESPACE_BYTES: usize = if PML_TEXT.len() > STRICT_TEXT.len() {
    PML_TEXT.len()
} else {
    STRICT_TEXT.len()
};
const ROOT_PREFIX_BYTES: usize = XML_DECL.len() + ROOT_OPEN.len() + MAX_NAMESPACE_BYTES + 1;
const EMPTY_WIRE_BYTES: usize = ROOT_PREFIX_BYTES + ROOT_EMPTY_CLOSE.len();
const NONEMPTY_WIRE_BYTES: usize = ROOT_PREFIX_BYTES + ROOT_CHILDREN_OPEN.len() + ROOT_CLOSE.len();

use crate::{Error, Result};

mod codec;
mod model;
mod package;

#[cfg(test)]
mod tests;

/// Shape-owned programmable-tag anchors and package CRUD.
pub mod shape;

pub use codec::{parse, write};
pub use model::{Conformance, Key, List, Source, Tag, is_relationship, raw};
pub use package::{discover, load, put, remove};

// Crate-internal seams shared with the shape-owned package layer. They retain
// the old parent-private topology without expanding the external API.
pub(crate) use codec::{bounded_text, is_namespace, pml, validate_qname};
pub(crate) use package::{
    Anchor, available_part_name, available_relationship_id, has_other_inbound,
    relationship_namespace, replace_xml, resolve_anchor, staged_xml, validate_relative_target,
    validate_selected_relationship,
};

pub(crate) fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

pub(crate) fn allocation(
    resource: &'static str,
    source: std::collections::TryReserveError,
) -> Error {
    Error::Allocation { resource, source }
}
