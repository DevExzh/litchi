//! Layered writer for the XLSB PivotCache definition stream.
//!
//! The facade keeps package integration contextual: semantic validation and
//! orchestration live above the binary record encoder, while the encoder owns
//! the BIFF12 wire layout and its byte-preserving snapshot behavior.

mod binary;
mod semantic;
mod validation;

pub(crate) use semantic::write_pivot_cache_definition;
