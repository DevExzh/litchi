//! Smart format detection — opens files via the per-format crates to disambiguate.
//!
//! The root coordinator composes the leaf probes (ADR 0010): the CFB
//! substrate with OLE stream-name classification, OPC content types for the
//! standalone OOXML owners, `litchi_odf_common::detect` for OpenDocument
//! (ADR 0009), and the RTF transport probe. Format-specific detectors stay in
//! their owning leaf crates.
//!
//! The iWork arms are not here. iWork is the last ZIP candidate, and
//! `litchi-ops` must not depend on an iWork crate, so the detectors without a
//! tail classify every other format with unchanged precedence and report an
//! iWork package as unrecognized. Each detector has a `_with_tail` form that
//! takes a [`ZipTail`](crate::detect::tail::ZipTail): the classifier of the
//! ZIP inputs every other candidate declined. The `litchi` facade passes its
//! iWork classifier there.

pub(crate) mod detected;
pub(crate) mod functions;

// Format-family probes used by `functions`.
pub(crate) mod ole2;
pub mod ooxml;
pub mod tail;
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

/// Resets the current thread's OPC probe counter to zero.
///
/// Test support, compiled only with the `test-support` feature: never part
/// of a release build. The counter records each OPC package index that a
/// detection or an open builds on the current thread, so a test can prove
/// that an open indexes its container once; see [`opc_probe_count`].
#[cfg(any(test, feature = "test-support"))]
#[allow(
    dead_code,
    reason = "the probe counter is used only by feature-gated detection tests"
)]
pub fn reset_opc_probe_count() {
    OPC_PROBE_COUNT.with(|count| count.set(0));
}

/// The number of OPC package indexes built on the current thread since the
/// last [`reset_opc_probe_count`].
///
/// Test support, compiled only with the `test-support` feature.
///
/// ```
/// # #[cfg(all(feature = "test-support", feature = "docx"))]
/// # {
/// use litchi_ops::detect::{opc_probe_count, reset_opc_probe_count};
///
/// let path = concat!(
///     env!("CARGO_MANIFEST_DIR"),
///     "/../../test-data/ooxml/docx/documentProperties.docx"
/// );
/// reset_opc_probe_count();
/// let file = litchi_ops::open(path)?;
/// assert_eq!(file.text()?.trim(), "Hello World!");
/// // Detection and the open share one index of the package.
/// assert_eq!(opc_probe_count(), 1);
/// # }
/// # Ok::<(), litchi_ops::Error>(())
/// ```
#[cfg(any(test, feature = "test-support"))]
#[allow(
    dead_code,
    reason = "the probe counter is used only by feature-gated detection tests"
)]
pub fn opc_probe_count() -> usize {
    OPC_PROBE_COUNT.with(Cell::get)
}

#[cfg(all(
    any(feature = "odt", feature = "ods", feature = "odp"),
    any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb")
))]
pub(crate) fn catalog_probe_limits(
    limits: crate::opc::SourceLimits,
) -> litchi_odf_common::detect::CatalogProbeLimits {
    litchi_odf_common::detect::CatalogProbeLimits::new(
        limits.max_input_bytes(),
        limits.max_members(),
        limits.max_archive_member_name_bytes(),
        limits.max_archive_metadata_bytes(),
        limits.max_archive_compressed_bytes(),
        limits.max_entry_bytes(),
        limits.max_aggregate_bytes(),
    )
}

#[cfg(all(
    any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"),
    any(unix, windows)
))]
pub use detected::PreparedDetection;
pub use detected::{DetectedFormat, detect_format_smart, detect_format_smart_with_tail};
#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
pub use detected::{detect_format_smart_with_limits, detect_format_smart_with_limits_and_tail};
#[cfg(all(
    any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"),
    any(unix, windows)
))]
pub use functions::{
    detect_and_prepare, detect_and_prepare_with_limits, detect_and_prepare_with_limits_and_tail,
};
pub use functions::{
    detect_file_format, detect_file_format_from_bytes, detect_file_format_from_bytes_with_tail,
    detect_file_format_with_tail, detect_format_from_reader, detect_format_from_reader_with_tail,
};
