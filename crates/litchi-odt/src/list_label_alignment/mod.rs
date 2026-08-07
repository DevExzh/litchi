//! ODF 1.2/1.3 list-level label alignment.
//!
//! The public owner facade keeps the contextual list-alignment vocabulary at
//! `list_label_alignment`, while models, XML codecs, and package adapters live
//! in their respective layers.

mod codec;
mod model;
mod package;

pub use codec::parse;
pub use model::{Alignment, FollowedBy, Kind, Length, Style, Styles};

pub(crate) use codec::set_xml;

const OFFICE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const STYLE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const TEXT: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const FO: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
const STYLE_S: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const TEXT_S: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const FO_S: &str = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
const MAX_XML: usize = 64 * 1024 * 1024;
const MAX_DEPTH: usize = 256;
const MAX_ENTRIES: usize = 65_536;
const MAX_VALUE: usize = 4096;
const MAX_TOTAL: usize = 16 * 1024 * 1024;
const MAX_LEVEL: u16 = 1024;

fn bad(message: impl Into<String>) -> litchi_core::Error {
    litchi_core::Error::InvalidFormat(message.into())
}
