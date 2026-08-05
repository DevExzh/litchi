//! Complete typed ODF ruby styles and structure-preserving inline annotations.
//!
//! The public owner is kept ergonomic at the module boundary while semantic
//! models, XML codecs, and regression tests live in dedicated layers.

use crate::{FlatDocument, Package};
pub(super) use litchi_core::{Result, xml::escape_xml};
pub(super) use quick_xml::{
    XmlVersion,
    events::Event,
    name::{Namespace, ResolveResult},
    reader::NsReader,
};
pub(super) use std::ops::Range;

const OFFICE_URI: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const STYLE_URI: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const TEXT_URI: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const DRAW_URI: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const DR3D_URI: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0";
const PRESENTATION_URI: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";
const MAX_XML: usize = 32 * 1_048_576;
const MAX_VALUE: usize = 1_048_576;
const MAX_BASE: usize = 4 * 1_048_576;
const MAX_DEPTH: usize = 256;
const MAX_EVENTS: usize = 1_000_000;
const MAX_RUBIES: usize = 65_536;
const MAX_STYLES: usize = 65_536;
const MAX_ATTRIBUTES: usize = 64;

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub(super) enum Ns {
    None,
    Office,
    Style,
    Text,
    Draw,
    Dr3d,
    Presentation,
    Other,
}

pub(super) fn ns(value: &ResolveResult<'_>) -> Ns {
    match value {
        ResolveResult::Unbound => Ns::None,
        ResolveResult::Bound(Namespace(v)) if *v == OFFICE_URI => Ns::Office,
        ResolveResult::Bound(Namespace(v)) if *v == STYLE_URI => Ns::Style,
        ResolveResult::Bound(Namespace(v)) if *v == TEXT_URI => Ns::Text,
        ResolveResult::Bound(Namespace(v)) if *v == DRAW_URI => Ns::Draw,
        ResolveResult::Bound(Namespace(v)) if *v == DR3D_URI => Ns::Dr3d,
        ResolveResult::Bound(Namespace(v)) if *v == PRESENTATION_URI => Ns::Presentation,
        _ => Ns::Other,
    }
}

#[path = "../ruby_inline_specs/mod.rs"]
mod ruby_inline_specs;
#[path = "../ruby_range/mod.rs"]
mod ruby_range;

mod codec;
mod model;

#[cfg(test)]
mod tests;

use model::{Span, bad, validate_style_name, validate_text};

pub use codec::{
    insert_ruby_annotation_xml, parse_ruby_annotations, parse_ruby_styles,
    remove_ruby_annotation_xml, remove_ruby_style_xml, replace_ruby_annotation_xml,
    set_ruby_style_xml, wrap_ruby_annotation_xml,
};
pub use model::{Alignment, Annotation, Annotations, Base, Position, Properties, Style, Styles};

pub(crate) use codec::ruby_parent;

impl Package {
    pub fn ruby_styles(&self) -> Result<Styles> {
        self.styles_xml()?
            .map_or_else(|| Ok(Default::default()), |xml| parse_ruby_styles(&xml))
    }

    pub fn ruby_annotations(&self) -> Result<Annotations> {
        parse_ruby_annotations(&self.content_xml()?)
    }
}

impl FlatDocument {
    pub fn ruby_styles(&self) -> Result<Styles> {
        parse_ruby_styles(self.xml())
    }

    pub fn ruby_annotations(&self) -> Result<Annotations> {
        parse_ruby_annotations(self.xml())
    }
}
