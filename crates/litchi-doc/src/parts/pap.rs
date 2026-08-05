//! Word paragraph-property (PAP/PAPX) model and codecs.
//!
//! The public module path remains [`crate::parts::pap`], while the semantic
//! model, high-level parser, operand codecs, and regression tests live in
//! dedicated child modules.

mod codec;
mod model;
mod parser;

#[cfg(test)]
mod tests;

pub use model::*;
