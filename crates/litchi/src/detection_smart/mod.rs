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

#[cfg(test)]
use std::cell::Cell;

#[cfg(test)]
thread_local! {
    static OPC_PROBE_COUNT: Cell<usize> = const { Cell::new(0) };
}

#[cfg(test)]
#[allow(
    dead_code,
    reason = "the probe counter is used only by feature-gated detection tests"
)]
pub(crate) fn record_opc_probe() {
    OPC_PROBE_COUNT.with(|count| count.set(count.get().saturating_add(1)));
}

#[cfg(test)]
#[allow(
    dead_code,
    reason = "the probe counter is used only by feature-gated detection tests"
)]
pub(crate) fn reset_opc_probe_count() {
    OPC_PROBE_COUNT.with(|count| count.set(0));
}

#[cfg(test)]
#[allow(
    dead_code,
    reason = "the probe counter is used only by feature-gated detection tests"
)]
pub(crate) fn opc_probe_count() -> usize {
    OPC_PROBE_COUNT.with(Cell::get)
}

#[cfg(all(
    any(feature = "odt", feature = "ods", feature = "odp"),
    any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb")
))]
pub(crate) fn catalog_probe_limits(
    limits: crate::opc::ReadLimits,
) -> litchi_odf_common::detect::CatalogProbeLimits {
    litchi_odf_common::detect::CatalogProbeLimits::new(
        limits.max_input_bytes(),
        limits.max_archive_members(),
        limits.max_archive_member_name_bytes(),
        limits.max_archive_metadata_bytes(),
        limits.max_archive_compressed_bytes(),
        limits.max_archive_entry_bytes(),
        limits.max_archive_total_bytes(),
    )
}

#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
pub use detected::detect_format_smart_with_limits;
pub use detected::{DetectedFormat, detect_format_smart};
pub use functions::{detect_file_format, detect_file_format_from_bytes, detect_format_from_reader};
