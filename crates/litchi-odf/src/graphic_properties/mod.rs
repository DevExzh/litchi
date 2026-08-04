//! Complete typed ODF `style:graphic-properties` support.
//!
//! The owner is layered into semantic models, bounded XML codecs, package/flat
//! accessors, and regression fixtures. The normative 174-property table remains
//! a generated local grammar source.

mod codec;
mod model;
mod package;

#[cfg(test)]
mod tests;

pub(super) const OFFICE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
pub(super) const STYLE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:style:1.0";
pub(super) const DR3D: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0";
pub(super) const DRAW: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
pub(super) const FO: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
pub(super) const SVG: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
pub(super) const TEXT: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
pub(super) const XLINK: &[u8] = b"http://www.w3.org/1999/xlink";
pub(super) const OFFICE_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
pub(super) const STYLE_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
pub(super) const DR3D_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0";
pub(super) const DRAW_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
pub(super) const FO_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
pub(super) const SVG_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
pub(super) const TEXT_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
pub(super) const XLINK_NS: &str = "http://www.w3.org/1999/xlink";
pub(super) const MAX_XML: usize = 32 * 1024 * 1024;
pub(super) const MAX_VALUE: usize = 1024 * 1024;
pub(super) const MAX_ATTRIBUTES: usize = 256;
pub(super) const MAX_DEPTH: usize = 128;
pub(super) const MAX_STYLES: usize = 65_536;
pub(super) const MAX_TOTAL: usize = 24 * 1024 * 1024;
pub(super) const MAX_EVENTS: usize = 1_000_000;

pub use codec::{parse_graphic_style_properties, set_graphic_style_properties_xml};
pub use model::{Child, ChildKind, Kind, Namespace, Properties, Property, Style, Styles, Value};

// Historical public names remain aliases; the concise declarations above are canonical.
pub type GraphicProperty = Property;
pub type GraphicPropertyChild = Child;
pub type GraphicPropertyChildKind = ChildKind;
pub type GraphicPropertyKind = Kind;
pub type GraphicPropertyNamespace = Namespace;
pub type GraphicPropertyValue = Value;
pub type GraphicStyleProperties = Properties;
pub type GraphicStylePropertiesSet = Styles;
pub type GraphicStyleRecord = Style;
