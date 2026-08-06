//! Layered Word 2015 `symEx` run symbols.
//!
//! The semantic values live in [`model`], bounded lexical and resource rules
//! live in [`validation`], and the source-preserving WordprocessingML seam is
//! implemented by [`codec`] and [`transaction`].  Unknown run content is not
//! interpreted and remains byte-for-byte intact when a symbol edit is
//! committed.

pub(crate) mod codec;
mod model;
mod package;
mod transaction;
mod validation;

#[cfg(test)]
mod tests;

pub use model::{MAX_FONT_CHARS, MAX_SYMBOLS, Symbol, Symbols};
pub use transaction::{Commit, Patch, Snapshot, Transaction};

/// Target namespace of the Word 2015 `symEx` extension.
pub const NAMESPACE: &str = "http://schemas.microsoft.com/office/word/2015/wordml/symex";
