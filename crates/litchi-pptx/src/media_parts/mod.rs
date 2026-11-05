//! Semantic layered owner for `PresentationML` slide media.
//!
//! The model, bounded XML codec, and OPC package graph are kept in separate
//! layers. The canonical declarations use this module's context so callers
//! can compose the model without redundant format prefixes.

mod codec;
mod model;
pub(crate) mod package;
mod summary;
#[cfg(test)]
mod tests;

pub use codec::{parse, write_pictures};
pub use model::{
    Bookmark, Conformance, Data, Extension, ExtensionList, Fade, Kind, List, Picture, Poster,
    Resource, Transform, Trim,
};
pub use package::load;
#[cfg(test)]
pub(crate) use package::store_with;
pub(crate) use package::{store_with_limits, validate_replay_part, validate_replay_references};
pub use summary::{Summary, summary};

// These are crate-private seams shared by the XML and package layers and by
// the in-module regression tests; they do not expand the external API.
#[cfg(test)]
pub(crate) use codec::{BoundedXml, write_picture};
pub(crate) use codec::{
    document_conformance, document_conformance_with_limits, validate_id, xml_error,
};
#[cfg(test)]
pub(crate) use package::resource_uri;
#[cfg(test)]
pub(crate) use package::validate_value;
pub(crate) use package::{parse_time, validate_value_with_limits};

use crate::Result;
use crate::error::CodecLimitKind;
use crate::time::{Offset, ParseError as TimeParseError};
use litchi_drawingml::coordinate::{Coordinate, Extent, ParseError as CoordinateParseError};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, PrefixDeclaration, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::{BTreeMap, BTreeSet, HashSet};
use std::sync::Arc;

pub(crate) const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
pub(crate) const STRICT_PML: &str = "http://purl.oclc.org/ooxml/presentationml/main";
pub(crate) const DML: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
pub(crate) const STRICT_DML: &str = "http://purl.oclc.org/ooxml/drawingml/main";
pub(crate) const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
pub(crate) const STRICT_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
pub(crate) const STRICT_AUDIO_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/audio";
pub(crate) const STRICT_VIDEO_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/video";
pub(crate) const P14: &str = "http://schemas.microsoft.com/office/powerpoint/2010/main";
pub(crate) const MEDIA_EXTENSION_URI: &str = "{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}";
pub(crate) const MAX_XML_BYTES: usize = 32 * 1024 * 1024;
pub(crate) const MAX_NODES: usize = 500_000;
pub(crate) const MAX_DEPTH: usize = 256;
pub(crate) const MAX_STRING_BYTES: usize = 4 * 1024 * 1024;
pub(crate) const MAX_MEDIA: usize = 1024;
pub(crate) const MAX_BOOKMARKS: usize = 4096;
pub(crate) const MAX_MEDIA_EXTENSION_XML_BYTES: usize = 4 * 1024 * 1024;
pub(crate) const MAX_PAYLOAD_BYTES: usize = 128 * 1024 * 1024;
pub(crate) const MAX_TOTAL_PAYLOAD_BYTES: usize = 512 * 1024 * 1024;

/// Resource policy for the raw slide-media XML codec.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    pub max_input_bytes: usize,
    pub max_output_bytes: usize,
    pub max_depth: usize,
    pub max_nodes: usize,
    pub max_string_bytes: usize,
    pub max_media: usize,
    pub max_bookmarks: usize,
    pub max_extension_bytes: usize,
    pub max_payload_bytes: usize,
    pub max_total_payload_bytes: usize,
}

impl Limits {
    pub const DEFAULT: Self = Self {
        max_input_bytes: MAX_XML_BYTES,
        max_output_bytes: MAX_XML_BYTES,
        max_depth: MAX_DEPTH,
        max_nodes: MAX_NODES,
        max_string_bytes: MAX_STRING_BYTES,
        max_media: MAX_MEDIA,
        max_bookmarks: MAX_BOOKMARKS,
        max_extension_bytes: MAX_MEDIA_EXTENSION_XML_BYTES,
        max_payload_bytes: MAX_PAYLOAD_BYTES,
        max_total_payload_bytes: MAX_TOTAL_PAYLOAD_BYTES,
    };

    #[must_use]
    pub const fn from_source(source: litchi_opc::SourceLimits) -> Self {
        Self::DEFAULT.constrained_by_source(source)
    }

    /// Intersect a selected semantic policy with the physical source ceiling.
    ///
    /// The source policy remains authoritative for bytes, parser depth and
    /// event count; a caller-selected semantic policy can only narrow it.
    #[must_use]
    pub const fn constrained_by_source(self, source: litchi_opc::SourceLimits) -> Self {
        let part = cap_u64(source.max_part_bytes());
        Self {
            max_input_bytes: min_usize(self.max_input_bytes, part),
            max_output_bytes: min_usize(self.max_output_bytes, part),
            max_depth: min_usize(self.max_depth, source.max_depth()),
            max_nodes: min_usize(self.max_nodes, source.max_events()),
            max_string_bytes: min_usize(self.max_string_bytes, source.max_xml_attribute_bytes()),
            max_extension_bytes: min_usize(self.max_extension_bytes, part),
            max_payload_bytes: min_usize(self.max_payload_bytes, cap_u64(source.max_entry_bytes())),
            max_total_payload_bytes: min_usize(
                self.max_total_payload_bytes,
                cap_u64(source.max_total_part_bytes()),
            ),
            ..self
        }
    }
}

const fn min_usize(left: usize, right: usize) -> usize {
    if left < right { left } else { right }
}

const fn cap_u64(value: u64) -> usize {
    if value > usize::MAX as u64 {
        usize::MAX
    } else {
        value as usize
    }
}

pub(crate) fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

pub(crate) fn limit(kind: CodecLimitKind, observed: usize, maximum: usize) -> Error {
    Error::Limit {
        kind: LimitKind::Codec(kind),
        observed,
        maximum,
    }
}

pub(crate) fn output_limit(maximum: usize, observed: usize) -> Error {
    Error::Limit {
        kind: LimitKind::Codec(CodecLimitKind::OutputBytes),
        observed,
        maximum,
    }
}

pub(crate) fn coordinate_error(error: CoordinateParseError, name: &str) -> Error {
    match error {
        CoordinateParseError::Limit {
            kind,
            observed,
            maximum,
        } => Error::Drawing(litchi_drawingml::Error::Limit {
            kind,
            observed,
            maximum,
        }),
        _ => invalid(format!("invalid media transform {name}")),
    }
}

mod edit;
pub use edit::{Commit, Diagnostics, Edit, Error, LimitKind, Patch};
