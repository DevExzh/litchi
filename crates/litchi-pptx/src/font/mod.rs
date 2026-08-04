//! Typed PresentationML embedded fonts and inert OPC resources.

mod codec;
mod model;
mod package;
#[cfg(test)]
mod tests;

pub use model::{
    Charset, Conformance, Data, Face, Family, Font, Fonts, Format, Key, License, Panose,
    Permission, Pitch, PitchFamily, Restrictions, Style,
};
pub use package::{conformance, load, put, remove};

use crate::error::Error;

pub(crate) const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
pub(crate) const STRICT_PML: &str = "http://purl.oclc.org/ooxml/presentationml/main";
pub(crate) const REL_NS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
pub(crate) const STRICT_REL_NS: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
pub(crate) const MCE_NS: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
pub(crate) const FONT_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font";
pub(crate) const STRICT_FONT_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/font";
#[cfg(test)]
pub(crate) const PRESENTATION_CT: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";
pub(crate) const FONT_DATA_CT: &str = "application/x-fontdata";
pub(crate) const FONT_TTF_CT: &str = "application/x-font-ttf";
pub(crate) const MAX_XML_BYTES: usize = 16 * 1024 * 1024;
pub(crate) const MAX_NODES: usize = 250_000;
pub(crate) const MAX_DEPTH: usize = 256;
pub(crate) const MAX_STRING_BYTES: usize = 1024 * 1024;
pub(crate) const MAX_FONTS: usize = 4096;
pub(crate) const MAX_FONT_BYTES: usize = 64 * 1024 * 1024;
pub(crate) const MAX_TOTAL_FONT_BYTES: usize = 256 * 1024 * 1024;
pub(crate) const MAX_MCE_MARKED_BYTES: usize = MAX_XML_BYTES + MAX_NODES * 64;

pub(crate) fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

pub(crate) fn limit(name: &'static str) -> Error {
    Error::Limit {
        resource: name,
        limit: match name {
            "presentation XML bytes"
            | "MCE-processed presentation XML bytes"
            | "serialized embedded-font XML bytes"
            | "updated presentation XML bytes" => MAX_XML_BYTES,
            "XML nodes" => MAX_NODES,
            "XML depth" | "presentation XML depth" => MAX_DEPTH,
            "embedded fonts" => MAX_FONTS,
            "individual font bytes" => MAX_FONT_BYTES,
            "total font bytes" => MAX_TOTAL_FONT_BYTES,
            "embedded-font string bytes" | "XML string bytes" => MAX_STRING_BYTES,
            _ => usize::MAX,
        },
    }
}
