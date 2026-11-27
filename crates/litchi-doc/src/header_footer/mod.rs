//! Layered legacy Word header and footer stories.
//!
//! The semantic story value remains small and compatible with the existing
//! DOC document API. Package-originated story data and model materialization
//! are kept behind the focused owner so later MS-DOC codecs can grow without
//! flattening the public module.

mod codec;
mod model;
mod package;

#[cfg(test)]
mod tests;

pub use model::HeaderFooter;
