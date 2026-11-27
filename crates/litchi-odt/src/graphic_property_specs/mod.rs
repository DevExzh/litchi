//! Normative ODF `style:graphic-properties` vocabulary.
//!
//! The generated property table is kept in the model layer, while lexical
//! datatype validation lives in the codec layer. The owning
//! [`crate::graphic_properties`] facade re-exports the public typed values.

pub(crate) mod codec;
pub(crate) mod model;

#[cfg(test)]
mod tests;
