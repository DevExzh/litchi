//! Typed Word 2010 OpenType run-property extensions.
//!
//! This owner models only the `rPr` OpenType family from [MS-DOCX] §2.2.1 and
//! §2.6.  XML outside those children remains in the source snapshot and is
//! never reconstructed by the semantic model.  The codec therefore provides
//! a narrow, loss-preserving edit path for a complete `w:r` or `w:rPr`
//! fragment without taking ownership of an OPC package.

mod codec;
mod model;
mod transaction;
mod validation;

#[cfg(test)]
mod tests;

pub use model::{Ligatures, NumForm, NumSpacing, OnOff, OpenType, StyleSet, StyleSetId};
pub use transaction::{Commit, Patch, Snapshot, Transaction};

/// Rewrite a complete run fragment through the private codec boundary.
pub(crate) fn rewrite(xml: &[u8], value: &OpenType) -> crate::Result<Vec<u8>> {
    codec::rewrite(xml, value)
}
