//! Typed document-level PowerPoint math defaults.
//!
//! PresentationML stores this owner in the a14:m extension of
//! p:presentationPr. It is deliberately limited to the two OMML children
//! permitted by [MS-ODRAWXML] for presentation properties. Equation content
//! and rendering remain outside this owner.

mod codec;
mod model;

#[cfg(test)]
mod tests;

pub use model::{BinaryBreak, BinarySubtractionBreak, Properties};

pub(crate) use codec::{parse, write};
