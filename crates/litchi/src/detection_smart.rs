//! Smart format detection — opens files via the per-format crates to disambiguate.
//!
//! The coordinator lives in `litchi_ops::detect` (ADR 0033, M0a step S4).
//! This module keeps the facade's detection API and paths unchanged and adds
//! the iWork arms, which `litchi-ops` may not own because it must stay
//! iWork-free (`tools/non_iwork_gate.py`). iWork is the last ZIP candidate:
//! the ops detectors call the tail below only after OOXML and ODF have
//! declined a ZIP input, exactly where the iWork arms used to run.

pub mod detected {
    //! Smart format detection with reusable owned results.
    //!
    //! This module provides the `DetectedFormat` enum and the `detect_format_smart`
    //! function. OOXML and OLE results retain their parsed owners; packaged ODT
    //! results retain a validated archive index; ODS, ODP, iWork, and RTF results
    //! retain the caller's moved byte buffer for subsequent parsing.

    use super::FacadeTail;

    // The facade's open paths call the per-family source detectors directly.
    #[allow(
        unused_imports,
        reason = "an iWork-only build opens through none of the per-family detectors"
    )]
    pub(crate) use litchi_ops::__migration::detect::*;
    pub use litchi_ops::detect::DetectedFormat;
    #[cfg(all(
        any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"),
        any(unix, windows)
    ))]
    pub use litchi_ops::detect::PreparedDetection;

    /// Detect a format while moving the source into a reusable result.
    ///
    /// The result retains a reusable representation for immediate follow-up work:
    /// - OOXML files: parse OPC package once and return it
    /// - OLE2 files: parse OLE file once and return it
    /// - ODT files: retain one validated archive index for semantic opening
    /// - ODS and ODP files: return the moved bytes after package detection
    /// - iWork and RTF files: return the moved bytes after detection
    ///
    /// # Arguments
    ///
    /// * `bytes` - The file data as bytes (ownership transferred)
    ///
    /// # Returns
    ///
    /// * `Some(DetectedFormat)` - Format detected with a reusable owner or byte buffer
    /// * `None` - Format not recognized
    pub fn detect_format_smart(bytes: Vec<u8>) -> Option<DetectedFormat> {
        litchi_ops::__migration::detect::detect_format_smart_with_tail(bytes, &FacadeTail)
    }

    /// Detect a format while applying an explicit OPC resource policy to an OOXML
    /// ZIP candidate.
    ///
    /// The policy is used when a ZIP candidate must be ruled in or out as OOXML.
    /// OLE, RTF, ODF, and iWork detection retains its existing format-specific
    /// behavior after the ODF detector proves an ordinary package under its own
    /// finite policy. ODF-marked packages with a present or uncertain OOXML
    /// catalog fail closed on hard OPC policy errors.
    #[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
    pub fn detect_format_smart_with_limits(
        bytes: Vec<u8>,
        limits: crate::opc::ReadLimits,
    ) -> Option<DetectedFormat> {
        litchi_ops::__migration::detect::detect_format_smart_with_limits_and_tail(
            bytes,
            limits,
            &FacadeTail,
        )
    }

    /// Parse a path with an OOXML filename extension or ZIP magic, classifying
    /// a rejected ZIP with the iWork tail as the facade always has.
    #[cfg(any(feature = "docx", feature = "pptx"))]
    #[allow(
        dead_code,
        reason = "the unified Presentation path uses the source-backed probe on positional platforms; the eager helper remains for Document and non-positional builds"
    )]
    pub(crate) fn detect_ooxml_path_with_limits(
        path: impl AsRef<std::path::Path>,
        limits: crate::opc::ReadLimits,
    ) -> crate::opc::Result<Option<DetectedFormat>> {
        litchi_ops::__migration::detect::detect_ooxml_path_with_limits(path, limits, &FacadeTail)
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
    /// use litchi::common::detection::detect_file_format;
    ///
    /// let format = detect_file_format("document.docx");
    /// if let Some(format) = format {
    ///     println!("Detected format: {:?}", format);
    /// }
    /// # Ok::<(), std::io::Error>(())
    /// ```
    pub fn detect_file_format<P: AsRef<Path>>(path: P) -> Option<FileFormat> {
        litchi_ops::__migration::detect::detect_file_format_with_tail(path, &FacadeTail)
    }

    /// Detect a file format and retain the OOXML package the detection built.
    ///
    /// [`detect_file_format`] throws away the container index it needed to
    /// classify an OOXML package, so a caller that detects and then opens indexes
    /// the same ZIP twice. This entry point returns the same neutral
    /// classification **and**, for an OOXML package, an opaque handle to the
    /// already-validated package. Pass that handle to `Document::from_prepared`,
    /// `Presentation::from_prepared`, or `Workbook::from_prepared` and the open
    /// reuses the existing index instead of building a second one.
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
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::common::detection::{FileFormat, detect_and_prepare};
    /// use litchi::Document;
    ///
    /// if let Some(detected) = detect_and_prepare("document.docx") {
    ///     if detected.format() == FileFormat::Docx
    ///         && let Some(prepared) = detected.into_prepared()
    ///     {
    ///         let document = Document::from_prepared(prepared)?;
    ///         println!("{}", document.text()?);
    ///     }
    /// }
    /// # Ok::<(), litchi::common::Error>(())
    /// ```
    #[cfg(all(
        any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"),
        any(unix, windows)
    ))]
    pub fn detect_and_prepare<P: AsRef<Path>>(path: P) -> Option<PreparedDetection> {
        detect_and_prepare_with_limits(path, crate::opc::ReadLimits::default())
    }

    /// Detect a file format and retain the OOXML package, with an explicit OPC
    /// resource policy.
    ///
    /// See [`detect_and_prepare`]. The policy bounds the OPC probe only; an input
    /// that exceeds it is still classified through [`detect_file_format`], with no
    /// prepared handle.
    #[cfg(all(
        any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"),
        any(unix, windows)
    ))]
    pub fn detect_and_prepare_with_limits<P: AsRef<Path>>(
        path: P,
        limits: crate::opc::ReadLimits,
    ) -> Option<PreparedDetection> {
        litchi_ops::__migration::detect::detect_and_prepare_with_limits_and_tail(
            path,
            limits,
            &FacadeTail,
        )
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
    /// use litchi::common::detection::detect_file_format_from_bytes;
    /// use std::fs;
    ///
    /// let data = fs::read("document.docx")?;
    /// let format = detect_file_format_from_bytes(&data);
    /// if let Some(format) = format {
    ///     println!("Detected format: {:?}", format);
    /// }
    /// # Ok::<(), std::io::Error>(())
    /// ```
    pub fn detect_file_format_from_bytes(bytes: &[u8]) -> Option<FileFormat> {
        litchi_ops::__migration::detect::detect_file_format_from_bytes_with_tail(bytes, &FacadeTail)
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
    pub fn detect_format_from_reader<R: Read + Seek>(reader: &mut R) -> Option<FileFormat> {
        litchi_ops::__migration::detect::detect_format_from_reader_with_tail(reader, &FacadeTail)
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
pub(crate) use litchi_ops::__migration::detect::NoZipTail as FacadeTail;

#[cfg(test)]
pub(crate) use litchi_ops::__migration::detect::{opc_probe_count, reset_opc_probe_count};

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
#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
pub use detected::detect_format_smart_with_limits;
pub use detected::{DetectedFormat, detect_format_smart};
#[cfg(all(
    any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"),
    any(unix, windows)
))]
pub use functions::{detect_and_prepare, detect_and_prepare_with_limits};
pub use functions::{detect_file_format, detect_file_format_from_bytes, detect_format_from_reader};
