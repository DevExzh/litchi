//! Layered OLE metadata stream generation for DOC files.
//!
//! The semantic model describes the metadata Word places in the `\x01CompObj`
//! and `\x01Ole` streams. The codec owns the little-endian wire layout and the
//! validation module keeps the fixed writer profile independently checked.
//! The two generator functions remain the concise DOC-package facade.

mod codec;
mod model;
mod validation;

#[cfg(test)]
mod tests;

pub use model::{ClassId, CompObj, Metadata, Ole};

/// Generate the `\x01CompObj` stream for a Word document.
///
/// The returned bytes contain the Word.Document.8 class identifier and the
/// ANSI/Unicode metadata profile required by Word's embedded-object loader.
#[must_use]
pub fn generate_compobj_stream() -> Vec<u8> {
    let metadata = Metadata::word_document();
    let data = codec::write_comp_obj(metadata.comp_obj());
    debug_assert!(validation::comp_obj(&data, metadata.comp_obj()).is_ok());
    data
}

/// Generate the `\x01Ole` stream for a Word document.
///
/// This is the fixed 20-byte OLE version stream used by the DOC writer.
#[must_use]
pub fn generate_ole_stream() -> Vec<u8> {
    let metadata = Metadata::word_document();
    let data = codec::write_ole(metadata.ole());
    debug_assert!(validation::ole(&data, metadata.ole()).is_ok());
    data
}
