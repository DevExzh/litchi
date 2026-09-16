//! Mutable DOCX document authoring.
//!
//! The public facade stays intentionally small. Semantic state and editing
//! operations live in the model module, XML/settings codecs in the codec
//! module, preserved body/package integration in the package module, and
//! invariants in the tests module.

#[path = "doc/codec.rs"]
mod codec;
#[path = "doc/model.rs"]
mod model;
#[path = "doc/package.rs"]
mod package;
#[cfg(test)]
#[path = "doc/tests.rs"]
mod tests;

pub use super::super::format::ImageFormat;
pub(crate) use codec::compact_changed_document_xml;
pub use model::{MutableDocument, Protection};
pub(crate) use package::{BodyElement, DocumentBody};

/// The UTF-8 byte order mark, as it appears at the head of a part written by
/// a producer that marks its XML.
///
/// `quick-xml` consumes this mark before its first event and never counts it
/// in `Reader::buffer_position`, so a caller that mixes reported offsets with
/// slices of its own buffer reads three bytes low. The mutable document route
/// therefore splits the mark off before it parses and carries it verbatim
/// through serialization.
const BYTE_ORDER_MARK: &str = "\u{feff}";
