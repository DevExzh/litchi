/// Part implementations for Word documents.
///
/// This module contains implementations of various part types found in
/// Word documents, such as the document part, styles part, numbering part, etc.
pub mod document_part;
pub(crate) mod memo;

pub use document_part::DocumentPart;
pub(crate) use document_part::ParagraphIndexCache;
