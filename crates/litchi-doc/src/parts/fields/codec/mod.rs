//! Layered codecs for Word field tables and inert field instructions.
//!
//! The parent `fields` module keeps this facade private while the codec is
//! organized by responsibility:
//!
//! - [`binary`] materializes stored instruction/result ranges from PLCF field
//!   markers;
//! - [`parser`] parses bounded field-instruction grammar into semantic parts;
//! - [`semantic`] exposes typed, inert views over those parsed parts.

mod binary;
mod parser;
mod semantic;

// Keep the internal parser surface stable for the field model and focused
// tests. These are crate-internal implementation details, not public API.
pub(super) use parser::*;
