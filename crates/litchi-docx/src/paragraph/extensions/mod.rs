//! Word 2010 paragraph and table-row extension attributes.
//!
//! The owner keeps the checked `paraId`/`textId` pair and paragraph-only
//! `noSpellErr` state separate from the paragraph child-property model. XML
//! decoding, validation, and package-facing adapters stay in their own
//! layers; unmodeled paragraph content remains in the source XML owner.

mod codec;
mod model;
mod package;
mod validation;

#[cfg(test)]
mod tests;

pub use model::{Extensions, Id, Ids, WORD_2010_NAMESPACE};

pub(crate) use codec::{append_paragraph_attributes, append_row_attributes};
