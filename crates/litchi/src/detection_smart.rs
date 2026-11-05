//! Smart format detection — opens files via the per-format crates to disambiguate.
//!
//! The coordinator lives in `litchi_ops::detect` (ADR 0033, M0a step S4).
//! This module keeps the facade's detection API and paths unchanged and adds
//! the iWork arms, which `litchi-ops` may not own because it must stay
//! iWork-free (`tools/non_iwork_gate.py`). iWork is the last ZIP candidate:
//! each function here calls the `_with_tail` form of its `litchi_ops::detect`
//! counterpart with the facade's iWork tail, which runs only after OOXML and
//! ODF have declined a ZIP input.

pub use litchi_ops::detect::SourceLimits;

pub mod detected {
    //! Smart format detection with reusable owned results.
    //!
    //! This module provides the `DetectedFormat` enum and the `detect_format_smart`
    //! function. OOXML and OLE results retain their parsed owners; packaged ODT
    //! results retain a validated archive index; ODS, ODP, iWork, and RTF results
    //! retain the caller's moved byte buffer for subsequent parsing.

    use super::FacadeTail;

    pub use litchi_ops::detect::DetectedFormat;
    #[cfg(all(
        any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"),
        any(unix, windows)
    ))]
    pub use litchi_ops::detect::PreparedDetection;
    pub use litchi_ops::detect::SourceLimits;

    /// Detect a format while moving the source into a reusable result.
    ///
    /// The result retains a reusable representation for immediate follow-up work:
    /// - OOXML files: parse OPC package once and return it
    /// - OLE2 files: parse OLE file once and return it
    /// - ODT files: retain one validated archive index for semantic opening
    /// - ODS and ODP files: return the moved bytes after package detection
    /// - iWork and RTF files: return the moved bytes after detection
    ///
    /// `limits` bounds every raw probe used by detection, including the common
    /// input length and enabled container profiles. OLE, RTF, ODF, and iWork
    /// detection retains its existing format-specific precedence under those
    /// limits.
    /// ODF-marked packages with a present or uncertain OOXML catalog fail
    /// closed on hard OPC policy errors.
    ///
    /// # Arguments
    ///
    /// * `bytes` - The file data as bytes (ownership transferred)
    /// * `limits` - Resource ceilings for every raw probe used by detection
    ///
    /// # Returns
    ///
    /// * `Some(DetectedFormat)` - Format detected with a reusable owner or byte buffer
    /// * `None` - Format not recognized
    pub fn detect_format_smart(bytes: Vec<u8>, limits: &SourceLimits) -> Option<DetectedFormat> {
        litchi_ops::detect::detect_format_smart_with_tail(bytes, limits, &FacadeTail)
    }
}

pub mod functions {
    //! Core file format detection functions.
    //!
    //! Uses a combined fixed-signature check before invoking format-specific
    //! detectors owned by their leaf crates.

    use std::io::{Read, Seek};
    use std::path::Path;

    use litchi_core::detection::FileFormat;

    use super::FacadeTail;
    #[cfg(all(
        any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"),
        any(unix, windows)
    ))]
    use super::detected::PreparedDetection;
    pub use litchi_ops::detect::SourceLimits;

    /// Detect file format from a file path.
    ///
    /// This function opens the file and delegates to the enabled format owners.
    /// Container formats may require a complete package scan.
    ///
    /// # Arguments
    ///
    /// * `path` - Path to the file to analyze
    ///
    /// # Returns
    ///
    /// * `Some(FileFormat)` if a supported format is detected
    /// * `None` if the format is not recognized or file cannot be read
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::common::detection::{detect_file_format, SourceLimits};
    ///
    /// let format = detect_file_format("document.docx", &SourceLimits::default());
    /// if let Some(format) = format {
    ///     println!("Detected format: {:?}", format);
    /// }
    /// # Ok::<(), std::io::Error>(())
    /// ```
    pub fn detect_file_format<P: AsRef<Path>>(
        path: P,
        limits: &SourceLimits,
    ) -> Option<FileFormat> {
        litchi_ops::detect::detect_file_format_with_tail(path, limits, &FacadeTail)
    }

    /// Detect a file format and retain the OOXML package the detection built.
    ///
    /// [`detect_file_format`] throws away the container index it needed to
    /// classify an OOXML package, so a caller that detects and then opens indexes
    /// the same ZIP twice. This entry point returns the same neutral
    /// classification **and**, for an OOXML package, an opaque handle to the
    /// already-validated package. Pass the detection to
    /// [`open_prepared`](crate::open_prepared) and the open reuses the existing
    /// index instead of building a second one. ([`open`](crate::open) detects
    /// and indexes in one pass, so it needs no separate detection.)
    ///
    /// # Returns
    ///
    /// * `Some(PreparedDetection)` when a supported format is detected. The format
    ///   is always reported; [`PreparedDetection::prepared`] carries a handle only
    ///   for an OOXML package. A legacy OLE2 file, an OpenDocument package, RTF and
    ///   iWork are classified with no handle, because they have no OOXML adopter
    ///   to hand a prepared OPC package to.
    /// * `None` when the format is not recognized or the file cannot be read —
    ///   exactly when [`detect_file_format`] returns `None`.
    ///
    /// The classification never differs from [`detect_file_format`]: any input
    /// that cannot be prepared falls back to that detector.
    ///
    /// `limits` bounds every raw probe used by preparation. An input that
    /// exceeds the common ceiling or a configured container ceiling is not
    /// prepared.
    ///
    /// # Examples
    ///
    /// ```
    /// # #[cfg(feature = "docx")]
    /// # {
    /// use litchi::common::detection::SourceLimits;
    /// use litchi::common::detection::{FileFormat, detect_and_prepare};
    ///
    /// # let path = concat!(
    /// #     env!("CARGO_MANIFEST_DIR"),
    /// #     "/../../test-data/ooxml/docx/documentProperties.docx"
    /// # );
    /// if let Some(detection) = detect_and_prepare(path, &SourceLimits::default())
    ///     && detection.format() == FileFormat::Docx
    ///     && detection.is_prepared()
    /// {
    ///     let file = litchi::open_prepared(detection)?;
    ///     assert_eq!(file.text()?.trim(), "Hello World!");
    /// }
    /// # }
    /// # Ok::<(), litchi::Error>(())
    /// ```
    #[cfg(all(
        any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"),
        any(unix, windows)
    ))]
    pub fn detect_and_prepare<P: AsRef<Path>>(
        path: P,
        limits: &SourceLimits,
    ) -> Option<PreparedDetection> {
        litchi_ops::detect::detect_and_prepare_with_tail(path, limits, &FacadeTail)
    }

    /// Detect file format from a byte slice.
    ///
    /// This function analyzes the byte signature in memory without
    /// requiring file I/O, making it ideal for network data or
    /// in-memory processing.
    ///
    /// # Arguments
    ///
    /// * `bytes` - The file data as bytes
    ///
    /// # Returns
    ///
    /// * `Some(FileFormat)` if a supported format is detected
    /// * `None` if the format is not recognized
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::common::detection::{SourceLimits, detect_file_format_from_bytes};
    /// use std::fs;
    ///
    /// let data = fs::read("document.docx")?;
    /// let format = detect_file_format_from_bytes(&data, &SourceLimits::default());
    /// if let Some(format) = format {
    ///     println!("Detected format: {:?}", format);
    /// }
    /// # Ok::<(), std::io::Error>(())
    /// ```
    pub fn detect_file_format_from_bytes(
        bytes: &[u8],
        limits: &SourceLimits,
    ) -> Option<FileFormat> {
        litchi_ops::detect::detect_file_format_from_bytes_with_tail(bytes, limits, &FacadeTail)
    }

    /// Detect file format from any reader that implements Read + Seek.
    ///
    /// Detection always starts at the beginning of the stream. The caller's
    /// original cursor position is restored before returning. A failed restore is
    /// reported as `None` because the cursor contract could not be upheld.
    ///
    /// # Arguments
    ///
    /// * `reader` - A reader that can read and seek
    ///
    /// # Returns
    ///
    /// * `Some(FileFormat)` if a supported format is detected
    /// * `None` if the format is not recognized
    ///
    pub fn detect_format_from_reader<R: Read + Seek>(
        reader: &mut R,
        limits: &SourceLimits,
    ) -> Option<FileFormat> {
        litchi_ops::detect::detect_format_from_reader_with_tail(reader, limits, &FacadeTail)
    }
}

/// OOXML content-type classification.
pub use litchi_ops::detect::ooxml;

#[cfg(any(feature = "pages", feature = "keynote", feature = "numbers"))]
pub(crate) mod iwork_tail;

/// The lowest-precedence ZIP classifier of this build: the iWork arms when an
/// iWork format is enabled, nothing otherwise. The detection functions here
/// and the unified openers ([`crate::open`]) share it.
#[cfg(any(feature = "pages", feature = "keynote", feature = "numbers"))]
pub(crate) use iwork_tail::IworkTail as FacadeTail;
#[cfg(not(any(feature = "pages", feature = "keynote", feature = "numbers")))]
pub(crate) use litchi_ops::detect::tail::NoZipTail as FacadeTail;

// The probe counter of `litchi-ops`'s `test-support` feature, which this
// crate's dev-dependency enables; only the DOCX, PPTX and XLSX detection
// tests read it.
#[cfg(all(
    test,
    any(feature = "docx", feature = "pptx", feature = "xlsx"),
    any(unix, windows)
))]
pub(crate) use litchi_ops::detect::{opc_probe_count, reset_opc_probe_count};

#[cfg(all(
    test,
    any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"),
    any(unix, windows)
))]
mod prepared_index_tests;

#[cfg(test)]
mod tests;

#[cfg(all(
    any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"),
    any(unix, windows)
))]
pub use detected::PreparedDetection;
pub use detected::{DetectedFormat, detect_format_smart};
#[cfg(all(
    any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"),
    any(unix, windows)
))]
pub use functions::detect_and_prepare;
pub use functions::{detect_file_format, detect_file_format_from_bytes, detect_format_from_reader};
