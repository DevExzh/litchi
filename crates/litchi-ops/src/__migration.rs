//! Items the `litchi` facade calls while M0a moves its open paths here.
//!
//! Not public API: no stability promise, no documentation. Step S9 of the
//! M0a migration plan (`docs/unified-ops/0003-m0a-migration-plan.md`) deletes
//! this module once the facade opens through `litchi_ops::open`.

/// Facade-internal detection entry points and per-family source detectors.
pub mod detect {
    pub use crate::detect::detected::*;
    #[cfg(all(
        any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"),
        any(unix, windows)
    ))]
    pub use crate::detect::functions::detect_and_prepare_with_limits_and_tail;
    pub use crate::detect::functions::{
        detect_file_format_from_bytes_with_tail, detect_file_format_with_tail,
        detect_format_from_reader_with_tail,
    };
    pub use crate::detect::tail::{NoZipTail, ZipTail};
    #[cfg(any(feature = "xlsx", feature = "xlsb"))]
    pub use crate::detect::workbook_metadata::{MetadataRepairs, read_source_backed_metadata};
    #[cfg(feature = "test-support")]
    pub use crate::detect::{opc_probe_count, reset_opc_probe_count};
}
