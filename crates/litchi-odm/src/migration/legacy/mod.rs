//! OpenDocument Master (`.odm` and `.otm`) support.

mod document;

mod builder;
mod mutable;

pub use builder::{MasterDocumentBuilder, MasterDocumentElement, MasterSection};
pub use document::{MasterDocument, MasterSubdocument};
pub use mutable::MutableMasterDocument;
