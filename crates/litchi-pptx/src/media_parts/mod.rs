//! Semantic layered owner for `PresentationML` slide media.
//!
//! The model, bounded XML codec, and OPC package graph are kept in separate
//! layers. The canonical declarations use this module's context so callers
//! can compose the model without redundant format prefixes.

mod codec;
mod model;
mod package;
#[cfg(test)]
mod tests;

pub use codec::{
    parse, parse as parse_slide_media, write_pictures, write_pictures as write_slide_media_pictures,
};
pub use model::{
    Bookmark, Conformance, Data, Extension, ExtensionList, Fade, Kind, List, Picture, Poster,
    Resource, Transform, Trim,
};
pub use package::{load, load as load_slide_media, store, store as store_slide_media};

// These are crate-private seams shared by the XML and package layers and by
// the in-module regression tests; they do not expand the external API.
#[cfg(test)]
pub(crate) use codec::{BoundedXml, write_picture};
pub(crate) use codec::{bounded, document_conformance, validate_id, xml_error};
#[cfg(test)]
pub(crate) use package::resource_uri;
pub(crate) use package::{parse_time, validate_value};

use crate::time::{Offset, ParseError as TimeParseError};
use crate::{Error, Result};
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

pub(crate) fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

pub(crate) fn limit(name: &str) -> Error {
    invalid(format!("slide media {name} limit exceeded"))
}

pub(crate) fn output_limit(maximum: usize) -> Error {
    Error::Limit {
        resource: "slide media serialized XML bytes",
        limit: maximum,
    }
}

pub(crate) fn coordinate_error(error: CoordinateParseError, name: &str) -> Error {
    invalid(format!("invalid media transform {name}: {error}"))
}
