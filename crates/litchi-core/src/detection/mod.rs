//! File format detection utilities.
//!
//! This module provides fast, safe, and memory-efficient signature-based
//! file format detection. The "smart" detection that opens files via
//! per-format crates lives in the umbrella `litchi` crate (see `litchi::detection_smart`).

pub mod rtf;
pub mod simd_utils;
pub mod types;
pub mod utils;

pub use types::FileFormat;
