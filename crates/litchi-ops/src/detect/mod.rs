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

/// Resource limits for the raw detection probes.
///
/// Detection owns this profile because a path, reader, or byte probe may
/// inspect more than one container family before it can classify an input.
/// The optional family profiles are compiled only when their owning format is
/// enabled; the public type and its common input ceiling remain available in
/// every build exposing these detectors. `OpenOptions` converts its per-owner settings to this profile
/// for the unified open paths.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub struct SourceLimits {
    max_input_bytes: Option<u64>,
    #[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
    ooxml: Option<litchi_opc::SourceLimits>,
    #[cfg(any(feature = "odt", feature = "ods", feature = "odp", odf_objects))]
    odf: Option<litchi_odf_common::SourceLimits>,
    #[cfg(any(feature = "doc", feature = "ppt", feature = "xls"))]
    cfb: Option<litchi_cfb::OleFileLimits>,
    #[cfg(feature = "rtf")]
    rtf: Option<litchi_rtf::SourceLimits>,
}

impl SourceLimits {
    /// Start with each enabled detector's bounded leaf default.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            max_input_bytes: None,
            #[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
            ooxml: None,
            #[cfg(any(feature = "odt", feature = "ods", feature = "odp", odf_objects))]
            odf: None,
            #[cfg(any(feature = "doc", feature = "ppt", feature = "xls"))]
            cfb: None,
            #[cfg(feature = "rtf")]
            rtf: None,
        }
    }

    /// Cap the physical source length accepted by every raw probe.
    #[must_use]
    pub const fn max_input_bytes(mut self, bytes: u64) -> Self {
        self.max_input_bytes = Some(bytes);
        self
    }

    /// The optional common physical source ceiling.
    #[must_use]
    pub const fn max_input_bytes_value(self) -> Option<u64> {
        self.max_input_bytes
    }

    /// Select the OPC profile used by OOXML probes.
    #[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
    #[must_use]
    pub const fn ooxml_source_limits(mut self, limits: litchi_opc::SourceLimits) -> Self {
        self.ooxml = Some(limits);
        self
    }

    /// The optional OPC profile used by OOXML probes.
    #[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
    #[must_use]
    pub const fn ooxml_source_limits_value(self) -> Option<litchi_opc::SourceLimits> {
        self.ooxml
    }

    /// Select the OpenDocument profile used by ODF probes.
    #[cfg(any(feature = "odt", feature = "ods", feature = "odp", odf_objects))]
    #[must_use]
    pub const fn odf_source_limits(mut self, limits: litchi_odf_common::SourceLimits) -> Self {
        self.odf = Some(limits);
        self
    }

    /// The optional OpenDocument profile used by ODF probes.
    #[cfg(any(feature = "odt", feature = "ods", feature = "odp", odf_objects))]
    #[must_use]
    pub const fn odf_source_limits_value(self) -> Option<litchi_odf_common::SourceLimits> {
        self.odf
    }

    /// Select the CFB profile used by OLE probes.
    #[cfg(any(feature = "doc", feature = "ppt", feature = "xls"))]
    #[must_use]
    pub const fn cfb_source_limits(mut self, limits: litchi_cfb::OleFileLimits) -> Self {
        self.cfb = Some(limits);
        self
    }

    /// The optional CFB profile used by OLE probes.
    #[cfg(any(feature = "doc", feature = "ppt", feature = "xls"))]
    #[must_use]
    pub const fn cfb_source_limits_value(self) -> Option<litchi_cfb::OleFileLimits> {
        self.cfb
    }

    /// Select the RTF profile used by transport probes.
    #[cfg(feature = "rtf")]
    #[must_use]
    pub const fn rtf_source_limits(mut self, limits: litchi_rtf::SourceLimits) -> Self {
        self.rtf = Some(limits);
        self
    }

    /// The optional RTF profile used by transport probes.
    #[cfg(feature = "rtf")]
    #[must_use]
    pub const fn rtf_source_limits_value(self) -> Option<litchi_rtf::SourceLimits> {
        self.rtf
    }

    pub(crate) fn accepts_input_len(&self, length: u64) -> bool {
        self.max_input_bytes.is_none_or(|maximum| length <= maximum)
    }

    // Detection only inspects the transport signature, so the source-byte
    // ceiling applies here; lexer and decompression budgets apply on open.
    pub(crate) fn accepts_rtf_input_len(&self, length: u64) -> bool {
        #[cfg(feature = "rtf")]
        {
            length <= self.rtf.unwrap_or_default().max_input_bytes() as u64
        }
        #[cfg(not(feature = "rtf"))]
        {
            let _ = length;
            true
        }
    }

    #[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
    pub(crate) fn ooxml(&self) -> litchi_opc::SourceLimits {
        self.ooxml.unwrap_or_default()
    }

    #[cfg(any(feature = "odt", feature = "ods", feature = "odp", odf_objects))]
    pub(crate) fn odf(&self) -> litchi_odf_common::SourceLimits {
        self.odf.unwrap_or_default()
    }

    #[cfg(any(feature = "doc", feature = "ppt", feature = "xls"))]
    pub(crate) fn cfb(&self) -> litchi_cfb::OleFileLimits {
        self.cfb.unwrap_or_default()
    }

    #[cfg(any(feature = "odt", feature = "ods", feature = "odp", odf_objects))]
    pub(crate) fn odf_detect_limits(&self) -> litchi_odf_common::detect::Limits {
        litchi_odf_common::detect::Limits::new(self.odf().max_input_bytes())
    }

    #[cfg(all(
        any(feature = "odt", feature = "ods", feature = "odp", odf_objects),
        any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb")
    ))]
    pub(crate) fn odf_catalog_limits(&self) -> litchi_odf_common::detect::CatalogProbeLimits {
        // This probe establishes that an ODF candidate has no OPC catalog.
        // OPC-only ceilings cannot reject an ordinary OpenDocument package.
        let limits = self.odf();
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
}

impl Default for SourceLimits {
    fn default() -> Self {
        Self::new()
    }
}

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
    any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"),
    any(unix, windows)
))]
pub use detected::PreparedDetection;
pub use detected::{DetectedFormat, detect_format_smart, detect_format_smart_with_tail};
#[cfg(all(
    any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"),
    any(unix, windows)
))]
pub use functions::{detect_and_prepare, detect_and_prepare_with_tail};
pub use functions::{
    detect_file_format, detect_file_format_from_bytes, detect_file_format_from_bytes_with_tail,
    detect_file_format_with_tail, detect_format_from_reader, detect_format_from_reader_with_tail,
};
