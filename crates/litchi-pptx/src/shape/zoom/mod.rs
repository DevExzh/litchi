//! Typed, lossless `PresentationML` zoom metadata.
//!
//! The owner is deliberately contextual: it edits only the zoom
//! `mc:AlternateContent` entries of one slide XML part while preserving the
//! fallback shape, unknown choices, and unrelated XML byte-for-byte.

mod codec;
mod model;
pub(crate) mod package;

#[cfg(test)]
mod tests;

pub use model::{
    ImageType, Item, Layout, Link, Owner, Percentage, Properties, Relationship, Section, Slide,
    Summary, Target, Unknown, Zoom,
};

/// Read a lossless zoom owner from a complete slide or shape-tree XML value.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
#[inline]
pub fn read(xml: &[u8]) -> crate::Result<Owner> {
    Owner::read(xml)
}

pub(crate) use package::{load, remove, store};
