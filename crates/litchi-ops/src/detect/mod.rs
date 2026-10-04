//! Smart format detection — opens files via the per-format crates to disambiguate.
//!
//! The root coordinator composes the leaf probes (ADR 0010): the CFB
//! substrate with OLE stream-name classification, OPC content types for the
//! standalone OOXML owners, `litchi_odf_common::detect` for OpenDocument
//! (ADR 0009), and the RTF transport probe. Format-specific detectors stay in
//! their owning leaf crates.
//!
//! The iWork arms are not here. iWork is the last ZIP candidate, and
//! `litchi-ops` must not depend on an iWork crate, so the `litchi` facade
//! supplies them as a lowest-precedence tail. The detectors exported here
//! classify every other format with unchanged precedence and report an iWork
//! package as unrecognized.

pub(crate) mod detected;
pub(crate) mod functions;

// Format-family probes used by `functions`.
pub(crate) mod ole2;
pub mod ooxml;
pub(crate) mod tail;
#[cfg(any(feature = "xlsx", feature = "xlsb"))]
pub(crate) mod workbook_metadata;

/// Detect the fixed transport header used by compressed RTF.
///
/// The shared core detector intentionally has no dependency on the RTF leaf
/// crate. Keep this probe in the coordinator, where the optional RTF transport
/// implementation is available, so smart detection can hand compressed bytes
/// to the native parser without weakening non-RTF builds.
#[cfg(feature = "rtf")]
#[inline]
pub(crate) fn is_compressed_rtf(bytes: &[u8]) -> bool {
    litchi_rtf::transport::is_compressed_rtf(bytes)
}

#[cfg(any(test, feature = "test-support"))]
use std::cell::Cell;

#[cfg(any(test, feature = "test-support"))]
thread_local! {
    static OPC_PROBE_COUNT: Cell<usize> = const { Cell::new(0) };
}

#[cfg(any(test, feature = "test-support"))]
#[allow(
    dead_code,
    reason = "the probe counter is used only by feature-gated detection tests"
)]
pub(crate) fn record_opc_probe() {
    OPC_PROBE_COUNT.with(|count| count.set(count.get().saturating_add(1)));
}

#[cfg(any(test, feature = "test-support"))]
#[allow(
    dead_code,
    reason = "the probe counter is used only by feature-gated detection tests"
)]
#[doc(hidden)]
pub fn reset_opc_probe_count() {
    OPC_PROBE_COUNT.with(|count| count.set(0));
}

#[cfg(any(test, feature = "test-support"))]
#[allow(
    dead_code,
    reason = "the probe counter is used only by feature-gated detection tests"
)]
#[doc(hidden)]
pub fn opc_probe_count() -> usize {
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

#[cfg(all(
    any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"),
    any(unix, windows)
))]
pub use detected::PreparedDetection;
#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
pub use detected::detect_format_smart_with_limits;
pub use detected::{DetectedFormat, detect_format_smart};
#[cfg(all(
    any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"),
    any(unix, windows)
))]
pub use functions::{detect_and_prepare, detect_and_prepare_with_limits};
pub use functions::{detect_file_format, detect_file_format_from_bytes, detect_format_from_reader};
