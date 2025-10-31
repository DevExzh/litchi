//! File format detection utilities.
//!
//! This module provides fast, safe, and memory-efficient file format detection
//! for all Microsoft Office and related document formats supported by Litchi.
//!
//! The detection is based on file signatures (magic numbers) and file structure
//! analysis, reading only the minimal amount of data required for identification.

// Submodule declarations
pub mod detected;
pub mod functions;
pub mod iwork;
pub mod odf;
pub mod ole2;
pub mod ooxml;
pub mod rtf;
pub mod simd_utils;
pub mod types;
pub mod utils;

// Re-exports
pub use detected::{DetectedFormat, detect_format_smart};
pub use functions::{
    detect_file_format, detect_file_format_from_bytes, detect_format_from_reader,
    detect_iwork_format_from_path,
};
pub use types::FileFormat;
