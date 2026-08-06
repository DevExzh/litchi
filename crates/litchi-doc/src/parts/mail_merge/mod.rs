//! Layered legacy Word mail-merge metadata.
//!
//! The facade re-exports semantic models while keeping validation and bounded
//! binary/FIB codecs internal to the part implementation.

mod codec;
mod model;
mod validation;

#[cfg(test)]
mod tests;

pub use model::*;
