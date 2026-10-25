//! Typed `PresentationML` embedded fonts and inert OPC resources.

#[cfg(feature = "automatic-fonts")]
mod automatic;
mod codec;
mod model;
mod package;
#[cfg(test)]
mod tests;

pub use model::{
    Charset, Conformance, Data, Face, Family, Font, Fonts, Format, Key, License, Panose,
    Permission, Pitch, PitchFamily, Restrictions, Style,
};
pub use package::{conformance, load};
mod edit;
pub use edit::{Commit, Diagnostics, Edit, Error, LimitKind, Patch};

const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const STRICT_PML: &str = "http://purl.oclc.org/ooxml/presentationml/main";
const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_REL_NS: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const MCE_NS: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const FONT_REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font";
const STRICT_FONT_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/font";
#[cfg(test)]
const PRESENTATION_CT: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";
const FONT_DATA_CT: &str = "application/x-fontdata";
const FONT_TTF_CT: &str = "application/x-font-ttf";
const MAX_XML_BYTES: usize = 16 * 1024 * 1024;
const MAX_NODES: usize = 250_000;
const MAX_DEPTH: usize = 256;
const MAX_STRING_BYTES: usize = 1024 * 1024;
const MAX_FONTS: usize = 4096;
const MAX_FONT_BYTES: usize = 64 * 1024 * 1024;
const MAX_TOTAL_FONT_BYTES: usize = 256 * 1024 * 1024;
const MAX_MCE_MARKED_BYTES: usize = MAX_XML_BYTES + MAX_NODES * 64;

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn limit(kind: crate::error::CodecLimitKind, observed: usize, maximum: usize) -> Error {
    Error::Limit {
        kind: LimitKind::Codec(kind),
        observed,
        maximum,
    }
}
pub(crate) use package::{put, remove};
