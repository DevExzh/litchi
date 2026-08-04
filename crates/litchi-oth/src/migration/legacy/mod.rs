//! LibreOffice-compatible OpenDocument Web template (`.oth`) support.

mod authoring;
mod document;

pub use authoring::{MutableWebDocument, WebDocumentBuilder};
pub use document::WebDocument;
