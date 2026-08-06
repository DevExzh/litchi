//! Layered Word 2010 run-property effects.
//!
//! The semantic values are kept in [`model`], bounded invariants in
//! [`validation`], and WordprocessingML wire handling in [`codec`].  This
//! owner intentionally stops at `[MS-DOCX]` §2.2.1: unsupported OpenType and
//! future namespace children remain ordered, bounded [`OpaqueExtension`]
//! values rather than being guessed into the visual model.

pub mod codec;
mod model;
mod validation;

pub use model::*;

/// Word 2010 WordprocessingML extension namespace.
pub const WORD_2010_NAMESPACE: &str = "http://schemas.microsoft.com/office/word/2010/wordml";

#[cfg(test)]
mod tests;
