//! Smart format detection — opens files via the per-format crates to disambiguate.
//!
//! Lives in the umbrella because it coordinates the CFB substrate with
//! `litchi_doc`, `crate::ppt`, `crate::xls`, the standalone OOXML modules, the concrete iWork owners, and
//! `crate::odf`. Format-specific detectors stay in their owning leaf crates.

pub mod detected;
pub mod functions;

// Format-family probes used by `functions`.
pub(crate) mod ole2;
pub mod ooxml;

#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
pub use detected::detect_format_smart_with_limits;
pub use detected::{DetectedFormat, detect_format_smart};
pub use functions::{detect_file_format, detect_file_format_from_bytes, detect_format_from_reader};
