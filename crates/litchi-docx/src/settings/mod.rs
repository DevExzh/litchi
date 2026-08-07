//! Layered WordprocessingML document-settings vocabulary and XML codec.
//
//! The package host owns relationship validation and settings-part orchestration;
//! this owner contains only the package-neutral settings model and bounded
//! settings.xml codec.

mod codec;
mod colors;
mod compatibility;
mod document;
mod editing;
mod extensions;
mod model;
mod notes;
mod smart_tags;
mod support;

#[cfg(test)]
mod tests;

/// Transitional WordprocessingML namespace.
pub const TRANSITIONAL_WORD_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/wordprocessingml/2006/main";
/// Strict WordprocessingML namespace.
pub const STRICT_WORD_NAMESPACE: &[u8] = b"http://purl.oclc.org/ooxml/wordprocessingml/main";

/// Maximum input accepted by the standalone settings codec.
pub const MAX_SETTINGS_XML_BYTES: usize = 16 * 1024 * 1024;
/// Maximum element count accepted by the standalone settings codec.
pub const MAX_SETTINGS_XML_NODES: usize = 250_000;
/// Maximum nesting depth accepted by the standalone settings codec.
pub const MAX_SETTINGS_XML_DEPTH: usize = 256;

pub use colors::*;
pub use compatibility::*;
#[cfg(feature = "automatic-fonts")]
pub(crate) use document::patch_font_embedding;
pub(crate) use document::{
    ATTACHED_TEMPLATE_RELATIONSHIP, extract_document_variables, patch_attached_template,
    patch_document_variables, patch_mail_merge, validate_attached_template_target,
};
pub use document::{AttachedTemplate, DocumentSettings};
pub use editing::*;
pub use extensions::*;
pub use model::*;
pub use notes::*;
pub use smart_tags::*;

#[cfg(test)]
pub(crate) use document::{
    STRICT_ATTACHED_TEMPLATE_RELATIONSHIP, is_attached_template_relationship,
};
