//! Smart format detection — opens files via the per-format crates to disambiguate.
//!
//! Lives in the umbrella because it depends on `crate::ole`, `crate::ooxml`, `crate::iwa`.
//! `litchi-core` exposes only the leaf signature-based detection.

pub mod detected;
pub mod functions;

// Format-family probes used by `functions`.
pub(crate) mod iwork;
pub(crate) mod ole2;
pub(crate) mod ooxml;

pub use detected::{DetectedFormat, detect_format_smart};
pub use functions::{
    detect_file_format, detect_file_format_from_bytes, detect_format_from_reader,
    detect_iwork_format_from_path,
};
