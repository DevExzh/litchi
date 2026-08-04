//! Compatibility re-exports for the canonical PPTX timing and animation codec.
//!
//! Semantic timing values, bounded XML parsing, MCE handling, and package
//! relationship validation live in litchi_pptx::animations. The OOXML
//! crate keeps this module as the stable host-facing import path.

pub use litchi_pptx::animations::*;
