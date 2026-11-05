//! Smart format detection with reusable owned results.
//!
//! This module provides the `DetectedFormat` enum and the `detect_format_smart`
//! function. OOXML and OLE results retain their parsed owners; packaged ODT
//! results retain a validated archive index; ODS, ODP and RTF results retain
//! the caller's moved byte buffer for subsequent parsing, as do the iWork
//! results a [`ZipTail`] produces.

#![allow(
    trivial_casts,
    unused_qualifications,
    reason = "moved verbatim from the facade (M0a S4), which does not inherit the workspace lints"
)]

use super::SourceLimits;
use super::tail::{NoZipTail, ZipTail};

/// Detected format with reusable parsed owners or moved source bytes.
///
/// This enum represents the result of format detection, where each format
/// includes the most reusable representation available at this layer:
/// - OOXML formats (DOCX, PPTX, XLSX, XLSB): include parsed OPC package
/// - OLE2 formats (DOC, PPT, XLS): include parsed OleFile
/// - ODT: includes an owned, validated package index after detection
/// - ODS and ODP: include owned bytes after package detection
/// - iWork formats: include owned bytes after the detection of a
///   [`ZipTail`]; the detectors of this crate never produce them without one
/// - RTF: includes owned bytes
///
/// An iWork tail may scan a container before a later document parser reads
/// the retained bytes again. Packaged ODT detection retains its bounded
/// index and makes that handoff parsing-once at the archive-structure layer.
#[allow(
    clippy::large_enum_variant,
    reason = "public detection results retain parsed format owners so the consuming facade can avoid reparsing; boxing a source-compatible variant is a separate API decision"
)]
pub enum DetectedFormat {
    // OOXML formats with parsed OPC package
    #[cfg(feature = "docx")]
    Docx(crate::opc::OpcPackage),
    #[cfg(feature = "pptx")]
    Pptx(crate::opc::OpcPackage),
    #[cfg(feature = "xlsx")]
    Xlsx(crate::opc::OpcPackage),
    #[cfg(feature = "xlsb")]
    Xlsb(crate::opc::OpcPackage),

    // OLE2 formats with parsed OleFile
    #[cfg(feature = "doc")]
    Doc(litchi_cfb::OleFile<std::io::Cursor<Vec<u8>>>),
    #[cfg(feature = "ppt")]
    Ppt(litchi_cfb::OleFile<std::io::Cursor<Vec<u8>>>),
    #[cfg(feature = "xls")]
    Xls(litchi_cfb::OleFile<std::io::Cursor<Vec<u8>>>),

    // iWork formats with validated ZIP archive data (lazy parsing)
    #[cfg(feature = "pages")]
    Pages(Vec<u8>),
    #[cfg(feature = "keynote")]
    Keynote(Vec<u8>),
    #[cfg(feature = "numbers")]
    Numbers(Vec<u8>),

    // ODT retains its validated ZIP index; ODS and ODP retain moved bytes.
    #[cfg(feature = "odt")]
    Odt(litchi_odf_common::PreparedPackage),
    #[cfg(feature = "odp")]
    Odp(Vec<u8>),
    #[cfg(feature = "ods")]
    Ods(Vec<u8>),
    /// Flat OpenDocument XML with its detected family.
    #[cfg(any(feature = "odt", feature = "ods", feature = "odp", odf_objects))]
    FlatOdf(litchi_core::detection::FileFormat, Vec<u8>),

    // RTF format (plain text, no parsing structure needed)
    #[cfg(feature = "rtf")]
    Rtf(Vec<u8>),
}

impl DetectedFormat {
    /// The neutral classification this detection result carries.
    ///
    /// Detection already knows the format; this accessor lets an opener that
    /// cannot handle the result say *which* format it declined instead of
    /// discarding the classification and reporting only "not an Office file".
    #[must_use]
    pub const fn format(&self) -> litchi_core::detection::FileFormat {
        #[allow(
            unused_imports,
            reason = "a build whose only ODF formats are the object formats names no variant format"
        )]
        use litchi_core::detection::FileFormat;

        match self {
            #[cfg(feature = "docx")]
            Self::Docx(_) => FileFormat::Docx,
            #[cfg(feature = "pptx")]
            Self::Pptx(_) => FileFormat::Pptx,
            #[cfg(feature = "xlsx")]
            Self::Xlsx(_) => FileFormat::Xlsx,
            #[cfg(feature = "xlsb")]
            Self::Xlsb(_) => FileFormat::Xlsb,
            #[cfg(feature = "doc")]
            Self::Doc(_) => FileFormat::Doc,
            #[cfg(feature = "ppt")]
            Self::Ppt(_) => FileFormat::Ppt,
            #[cfg(feature = "xls")]
            Self::Xls(_) => FileFormat::Xls,
            #[cfg(feature = "pages")]
            Self::Pages(_) => FileFormat::Pages,
            #[cfg(feature = "keynote")]
            Self::Keynote(_) => FileFormat::Keynote,
            #[cfg(feature = "numbers")]
            Self::Numbers(_) => FileFormat::Numbers,
            #[cfg(feature = "odt")]
            Self::Odt(_) => FileFormat::Odt,
            #[cfg(feature = "odp")]
            Self::Odp(_) => FileFormat::Odp,
            #[cfg(feature = "ods")]
            Self::Ods(_) => FileFormat::Ods,
            #[cfg(any(feature = "odt", feature = "ods", feature = "odp", odf_objects))]
            Self::FlatOdf(format, _) => *format,
            #[cfg(feature = "rtf")]
            Self::Rtf(_) => FileFormat::Rtf,
        }
    }
}

impl std::fmt::Debug for DetectedFormat {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        let name = match self {
            #[cfg(feature = "docx")]
            Self::Docx(_) => "Docx",
            #[cfg(feature = "pptx")]
            Self::Pptx(_) => "Pptx",
            #[cfg(feature = "xlsx")]
            Self::Xlsx(_) => "Xlsx",
            #[cfg(feature = "xlsb")]
            Self::Xlsb(_) => "Xlsb",
            #[cfg(feature = "doc")]
            Self::Doc(_) => "Doc",
            #[cfg(feature = "ppt")]
            Self::Ppt(_) => "Ppt",
            #[cfg(feature = "xls")]
            Self::Xls(_) => "Xls",
            #[cfg(feature = "pages")]
            Self::Pages(_) => "Pages",
            #[cfg(feature = "keynote")]
            Self::Keynote(_) => "Keynote",
            #[cfg(feature = "numbers")]
            Self::Numbers(_) => "Numbers",
            #[cfg(feature = "odt")]
            Self::Odt(_) => "Odt",
            #[cfg(feature = "odp")]
            Self::Odp(_) => "Odp",
            #[cfg(feature = "ods")]
            Self::Ods(_) => "Ods",
            #[cfg(any(feature = "odt", feature = "ods", feature = "odp", odf_objects))]
            Self::FlatOdf(_, _) => "FlatOdf",
            #[cfg(feature = "rtf")]
            Self::Rtf(_) => "Rtf",
        };
        formatter
            .debug_tuple("DetectedFormat")
            .field(&name)
            .finish()
    }
}

/// Detect a format while moving the source into a reusable result.
///
/// The result retains a reusable representation for immediate follow-up work:
/// - OOXML files: parse OPC package once and return it
/// - OLE2 files: parse OLE file once and return it
/// - ODT files: retain one validated archive index for semantic opening
/// - ODS and ODP files: return the moved bytes after package detection
/// - RTF files: return the moved bytes after detection
///
/// `limits` bounds every raw probe used by detection, including the common
/// input length and enabled container profiles. OLE, RTF and ODF detection
/// retains its existing format-specific behavior under those limits.
/// ODF-marked
/// packages with a present or uncertain OOXML catalog fail closed on hard OPC
/// policy errors.
///
/// An Apple iWork package is not recognized: this detector runs no
/// [`ZipTail`]. [`detect_format_smart_with_tail`] takes one.
///
/// # Arguments
///
/// * `bytes` - The file data as bytes (ownership transferred)
///
/// # Returns
///
/// * `Some(DetectedFormat)` - Format detected with a reusable owner or byte buffer
/// * `None` - Format not recognized
pub fn detect_format_smart(bytes: Vec<u8>, limits: &SourceLimits) -> Option<DetectedFormat> {
    detect_format_smart_with_tail(bytes, limits, &NoZipTail)
}

/// [`detect_format_smart`] with a lowest-precedence ZIP candidate.
///
/// `tail` classifies a ZIP input that every other candidate declined; see
/// [`tail`](super::tail) for when it runs.
pub fn detect_format_smart_with_tail<T: ZipTail>(
    bytes: Vec<u8>,
    limits: &SourceLimits,
    tail: &T,
) -> Option<DetectedFormat> {
    if !limits.accepts_input_len(bytes.len() as u64) {
        return None;
    }
    #[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
    {
        detect_ooxml_smart(bytes, limits, tail)
    }

    #[cfg(not(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb")))]
    {
        detect_format_smart_without_ooxml(bytes, limits, tail)
    }
}

#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
pub(crate) fn hard_ooxml_probe_error(error: &crate::opc::OpcError) -> bool {
    matches!(
        error,
        crate::opc::OpcError::InvalidLimit { .. }
            | crate::opc::OpcError::Limit { .. }
            | crate::opc::OpcError::Cancelled
            | crate::opc::OpcError::Execution(_)
            | crate::opc::OpcError::ParallelRead(_)
            | crate::opc::OpcError::UnsupportedExecutionAffinity
            | crate::opc::OpcError::IoError(_)
            | crate::opc::OpcError::SourceChanged { .. }
            | crate::opc::OpcError::Allocation { .. }
            | crate::opc::OpcError::CollectionAllocation { .. }
    )
}

#[cfg(any(unix, windows))]
#[allow(
    dead_code,
    reason = "shared positional source helpers serve the feature-gated open routes"
)]
pub(crate) fn ensure_path_source_current(
    source: &dyn litchi_core::ReadAt,
    expected: litchi_core::SourceVersion,
) -> litchi_core::Result<()> {
    let observed = source.version()?;
    if observed == expected {
        Ok(())
    } else {
        Err(litchi_core::Error::SourceChanged { expected, observed })
    }
}

#[cfg(any(unix, windows))]
#[allow(
    dead_code,
    reason = "shared positional source helpers serve the feature-gated open routes"
)]
pub(crate) fn read_path_source_bytes(
    source: &dyn litchi_core::ReadAt,
    expected: litchi_core::SourceVersion,
    max_input_bytes: u64,
) -> litchi_core::Result<Vec<u8>> {
    ensure_path_source_current(source, expected)?;
    let length = source.len()?;
    ensure_path_source_current(source, expected)?;

    if length > max_input_bytes {
        return Err(litchi_core::Error::ResourceLimit(
            litchi_core::ResourceLimit {
                resource: litchi_core::Resource::InputBytes,
                observed: length,
                limit: max_input_bytes,
                scope: std::sync::Arc::from("unified filesystem OOXML fallback"),
            },
        ));
    }
    let length = usize::try_from(length).map_err(|_| {
        litchi_core::Error::InvalidFormat(
            "filesystem source exceeds platform allocation limits".to_string(),
        )
    })?;
    let mut bytes = Vec::new();
    bytes
        .try_reserve_exact(length)
        .map_err(|source| litchi_core::Error::Allocation {
            resource: "unified filesystem OOXML fallback source bytes",
            source,
        })?;
    bytes.resize(length, 0);
    if let Err(error) = source.read_exact_at(0, &mut bytes) {
        return match ensure_path_source_current(source, expected) {
            Err(changed @ litchi_core::Error::SourceChanged { .. }) => Err(changed),
            Err(other) => Err(other),
            Ok(()) => Err(error.into()),
        };
    }
    ensure_path_source_current(source, expected)?;
    Ok(bytes)
}

#[cfg(not(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb")))]
fn detect_format_smart_without_ooxml<T: ZipTail>(
    bytes: Vec<u8>,
    limits: &SourceLimits,
    tail: &T,
) -> Option<DetectedFormat> {
    use litchi_core::detection::simd_utils::check_office_signatures;

    // The iWork-only build still receives the required source policy through
    // the public detector, but its ZIP tail owns the only format-specific
    // probe. Keep the parameter in the shared helper signature without
    // tripping the workspace's deny-unused profile in that feature set.
    #[cfg(not(any(
        feature = "odt",
        feature = "ods",
        feature = "odp",
        odf_objects,
        feature = "rtf",
        feature = "doc",
        feature = "ppt",
        feature = "xls"
    )))]
    let _ = limits;

    #[cfg(not(any(feature = "pages", feature = "keynote", feature = "numbers")))]
    let _ = tail;

    if bytes.len() < 4 {
        return None;
    }

    #[cfg(any(feature = "odt", feature = "ods", feature = "odp", odf_objects))]
    if let Some(format) = litchi_odf_common::detect::flat(&bytes, limits.odf_detect_limits())
        .ok()
        .flatten()
    {
        return Some(DetectedFormat::FlatOdf(format, bytes));
    }

    let mask = check_office_signatures(&bytes);

    #[cfg(feature = "rtf")]
    if mask.is_rtf() {
        return limits
            .accepts_rtf_input_len(bytes.len() as u64)
            .then_some(DetectedFormat::Rtf(bytes));
    }

    #[cfg(any(feature = "doc", feature = "ppt", feature = "xls"))]
    if mask.is_ole2() {
        let cursor = std::io::Cursor::new(bytes);
        if let Ok(ole_file) = litchi_cfb::OleFile::open(cursor, limits.cfb()) {
            // An encrypted OOXML package, even beside a stub `WordDocument`
            // stream (record 0341).
            if ole_file.exists(&["EncryptionInfo"]) && ole_file.exists(&["EncryptedPackage"]) {
                return None;
            }
            #[cfg(feature = "doc")]
            if ole_file.exists(&["WordDocument"]) {
                return Some(DetectedFormat::Doc(ole_file));
            }
            #[cfg(feature = "ppt")]
            if ole_file.exists(&["PowerPoint Document"]) || ole_file.exists(&["Current User"]) {
                return Some(DetectedFormat::Ppt(ole_file));
            }
            #[cfg(feature = "xls")]
            if ole_file.exists(&["Workbook"]) || ole_file.exists(&["Book"]) {
                return Some(DetectedFormat::Xls(ole_file));
            }
        }
        return None;
    }

    #[cfg(feature = "rtf")]
    if !mask.is_ole2() && !mask.is_zip() && super::is_compressed_rtf(&bytes) {
        return limits
            .accepts_rtf_input_len(bytes.len() as u64)
            .then_some(DetectedFormat::Rtf(bytes));
    }

    if mask.is_zip() {
        #[cfg(any(feature = "odt", feature = "ods", feature = "odp", odf_objects))]
        let bytes = match litchi_odf_common::detect::prepared_or_original(bytes, limits.odf()) {
            Ok(prepared) => {
                let format = prepared.format();
                return match format {
                    #[cfg(feature = "odt")]
                    litchi_core::detection::FileFormat::Odt => Some(DetectedFormat::Odt(prepared)),
                    #[cfg(feature = "odp")]
                    litchi_core::detection::FileFormat::Odp => {
                        Some(DetectedFormat::Odp(prepared.into_package().into_inner()))
                    },
                    #[cfg(feature = "ods")]
                    litchi_core::detection::FileFormat::Ods => {
                        Some(DetectedFormat::Ods(prepared.into_package().into_inner()))
                    },
                    _ => None,
                };
            },
            Err(bytes) => bytes,
        };

        #[cfg(any(feature = "pages", feature = "keynote", feature = "numbers"))]
        return tail.detect_owned(bytes);

        #[cfg(not(any(feature = "pages", feature = "keynote", feature = "numbers")))]
        drop(bytes);
    }

    None
}

/// [`detect_format_smart_with_tail`] in an OOXML build, under the OPC
/// resource policy `limits`.
#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
pub(crate) fn detect_ooxml_smart<T: ZipTail>(
    bytes: Vec<u8>,
    source_limits: &SourceLimits,
    tail: &T,
) -> Option<DetectedFormat> {
    #[cfg(not(any(feature = "pages", feature = "keynote", feature = "numbers")))]
    let _ = tail;
    #[cfg(any(
        any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"),
        any(feature = "odt", feature = "ods", feature = "odp", odf_objects)
    ))]
    use litchi_core::detection::FileFormat;
    use litchi_core::detection::simd_utils::check_office_signatures;

    // Quick signature checks. ZIP has a complete four-byte local-file
    // signature, RTF has a five-byte prefix, and OLE2 is checked only after
    // the classifier proves its full eight-byte signature. Do not make the
    // reusable detector stricter than the shared signature contract.
    if bytes.len() < 4 {
        return None;
    }

    #[cfg(any(feature = "odt", feature = "ods", feature = "odp", odf_objects))]
    if let Some(format) = litchi_odf_common::detect::flat(&bytes, source_limits.odf_detect_limits())
        .ok()
        .flatten()
    {
        return Some(DetectedFormat::FlatOdf(format, bytes));
    }

    // Classify the fixed signatures together before invoking format parsers.
    let mask = check_office_signatures(&bytes);

    // Check RTF first (simplest check, no parsing needed)
    #[cfg(feature = "rtf")]
    if mask.is_rtf() {
        return source_limits
            .accepts_rtf_input_len(bytes.len() as u64)
            .then_some(DetectedFormat::Rtf(bytes));
    }

    // Check OLE2 signature (DOC, PPT, XLS) - parse OleFile once
    #[cfg(any(feature = "doc", feature = "ppt", feature = "xls"))]
    if mask.is_ole2() {
        let cursor = std::io::Cursor::new(bytes);
        if let Ok(ole_file) = litchi_cfb::OleFile::open(cursor, source_limits.cfb()) {
            // An encrypted OOXML package, even beside a stub `WordDocument`
            // stream (record 0341).
            if ole_file.exists(&["EncryptionInfo"]) && ole_file.exists(&["EncryptedPackage"]) {
                return None;
            }
            // Use existing OLE2 detection logic by checking streams
            #[cfg(feature = "doc")]
            if ole_file.exists(&["WordDocument"]) {
                return Some(DetectedFormat::Doc(ole_file));
            }
            #[cfg(feature = "ppt")]
            if ole_file.exists(&["PowerPoint Document"]) || ole_file.exists(&["Current User"]) {
                return Some(DetectedFormat::Ppt(ole_file));
            }
            #[cfg(feature = "xls")]
            if ole_file.exists(&["Workbook"]) || ole_file.exists(&["Book"]) {
                return Some(DetectedFormat::Xls(ole_file));
            }
        }
        return None;
    }

    #[cfg(feature = "rtf")]
    if !mask.is_ole2() && !mask.is_zip() && super::is_compressed_rtf(&bytes) {
        return source_limits
            .accepts_rtf_input_len(bytes.len() as u64)
            .then_some(DetectedFormat::Rtf(bytes));
    }

    // Check ZIP candidates in the same order as the ordinary detector.
    if mask.is_zip() {
        #[cfg(all(
            any(feature = "odt", feature = "ods", feature = "odp", odf_objects),
            any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb")
        ))]
        let odf_mime =
            litchi_odf_common::detect::packaged_mime(&bytes, source_limits.odf_detect_limits())
                .ok()
                .flatten();
        #[cfg(all(
            any(feature = "odt", feature = "ods", feature = "odp", odf_objects),
            any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb")
        ))]
        let normal_odf = odf_mime.is_some()
            && litchi_odf_common::detect::packaged_has_ooxml_catalog(
                &bytes,
                source_limits.odf_catalog_limits(),
            ) == Some(false);
        #[cfg(all(
            not(any(feature = "odt", feature = "ods", feature = "odp", odf_objects)),
            any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb")
        ))]
        let normal_odf = false;

        // A successful OOXML probe returns the parsed OPC owner directly.
        #[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
        if !normal_odf {
            #[cfg(any(test, feature = "test-support"))]
            super::record_opc_probe();
            match crate::opc::OpcPackage::from_bytes(&bytes, source_limits.ooxml()) {
                Ok(package) => {
                    if let Some(format) =
                        crate::detect::ooxml::detect_ooxml_format_from_package(&package)
                    {
                        return match format {
                            #[cfg(feature = "docx")]
                            FileFormat::Docx => Some(DetectedFormat::Docx(package)),
                            #[cfg(feature = "pptx")]
                            FileFormat::Pptx => Some(DetectedFormat::Pptx(package)),
                            #[cfg(feature = "xlsx")]
                            FileFormat::Xlsx => Some(DetectedFormat::Xlsx(package)),
                            #[cfg(feature = "xlsb")]
                            FileFormat::Xlsb => Some(DetectedFormat::Xlsb(package)),
                            _ => None,
                        };
                    }
                },
                Err(error) if hard_ooxml_probe_error(&error) => return None,
                Err(_) => {},
            }
        }

        #[cfg(any(feature = "odt", feature = "ods", feature = "odp", odf_objects))]
        let bytes =
            match litchi_odf_common::detect::prepared_or_original(bytes, source_limits.odf()) {
                Ok(prepared) => {
                    let format = prepared.format();
                    return match format {
                        #[cfg(feature = "odt")]
                        FileFormat::Odt => Some(DetectedFormat::Odt(prepared)),
                        #[cfg(feature = "odp")]
                        FileFormat::Odp => {
                            Some(DetectedFormat::Odp(prepared.into_package().into_inner()))
                        },
                        #[cfg(feature = "ods")]
                        FileFormat::Ods => {
                            Some(DetectedFormat::Ods(prepared.into_package().into_inner()))
                        },
                        _ => None,
                    };
                },
                Err(bytes) => bytes,
            };

        #[cfg(any(feature = "pages", feature = "keynote", feature = "numbers"))]
        return tail.detect_owned(bytes);

        #[cfg(not(any(feature = "pages", feature = "keynote", feature = "numbers")))]
        drop(bytes);
    }

    None
}

/// A filesystem ODT the path opener classified by its package MIME, which
/// opens through one positional source-backed owner.
///
/// The classification does not consult a filename suffix: extensionless and
/// incorrectly suffixed ODT files are accepted when their package MIME
/// identifies ODT. The path opener separately arbitrates packages that also
/// contain an OOXML content-types catalog, so ordinary ODT packages do not
/// pay for an unrelated second container scan.
#[cfg(all(feature = "odt", any(unix, windows)))]
pub(crate) struct OdtSourcePathCandidate {
    source: std::sync::Arc<dyn litchi_core::ReadAt>,
    source_version: litchi_core::SourceVersion,
}

#[cfg(all(feature = "odt", any(unix, windows)))]
impl OdtSourcePathCandidate {
    /// A candidate over a source the caller already classified as an ODT
    /// package, pinned to the revision `source_version`.
    pub(crate) fn new(
        source: std::sync::Arc<dyn litchi_core::ReadAt>,
        source_version: litchi_core::SourceVersion,
    ) -> Self {
        Self {
            source,
            source_version,
        }
    }

    /// Open the source-backed ODT owner, refusing a source that changed since
    /// it was classified.
    pub(crate) fn into_document(
        self,
        options: litchi_odf_common::OpenOptions,
    ) -> litchi_core::Result<litchi_odt::SourceBackedDocument> {
        let source = std::sync::Arc::clone(&self.source);
        ensure_odt_source_current(source.as_ref(), self.source_version)?;
        let result =
            litchi_odt::SourceBackedDocument::from_read_at_with_options(self.source, options)
                .map_err(litchi_core::Error::from);
        match ensure_odt_source_current(source.as_ref(), self.source_version) {
            Err(error) => Err(error),
            Ok(()) => result,
        }
    }
}

#[cfg(all(feature = "odt", any(unix, windows)))]
fn ensure_odt_source_current(
    source: &dyn litchi_core::ReadAt,
    expected: litchi_core::SourceVersion,
) -> litchi_core::Result<()> {
    let observed = source.version()?;
    if observed == expected {
        Ok(())
    } else {
        Err(litchi_core::Error::SourceChanged { expected, observed })
    }
}

/// One path classification that may carry the OOXML package the classification
/// used.
///
/// Produced by [`crate::detect::detect_and_prepare`]. The neutral
/// format is always present; the prepared handle is present only for an OOXML
/// package, because only OOXML has an adopter that can reuse a prepared index.
/// A legacy OLE2, OpenDocument or RTF input (and an iWork package, through a
/// [`ZipTail`]) is still classified and reported, with [`Self::prepared`]
/// returning `None`.
#[cfg(all(
    any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"),
    any(unix, windows)
))]
pub struct PreparedDetection {
    format: litchi_core::detection::FileFormat,
    prepared: Option<crate::opc::PreparedOoxmlSource>,
    /// Stable path origin for the prepared path route. This is metadata for
    /// the adopter; it is intentionally only reported as presence in Debug.
    origin: Option<std::sync::Arc<std::path::Path>>,
}

#[cfg(all(
    any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"),
    any(unix, windows)
))]
impl std::fmt::Debug for PreparedDetection {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("PreparedDetection")
            .field("format", &self.format)
            .field("prepared", &self.prepared.is_some())
            .field("origin", &self.origin.is_some())
            .finish()
    }
}

#[cfg(all(
    any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"),
    any(unix, windows)
))]
impl PreparedDetection {
    /// A classification with no reusable package.
    fn unprepared(format: litchi_core::detection::FileFormat) -> Self {
        Self {
            format,
            prepared: None,
            origin: None,
        }
    }

    /// The neutral classification, identical to
    /// [`detect_file_format`](crate::detect::detect_file_format).
    #[must_use]
    pub const fn format(&self) -> litchi_core::detection::FileFormat {
        self.format
    }

    /// Whether an OOXML package was prepared and can be adopted without a
    /// second container index.
    #[must_use]
    pub const fn is_prepared(&self) -> bool {
        self.prepared.is_some()
    }

    /// Borrow the prepared OOXML package handle, if there is one.
    #[must_use]
    pub fn prepared(&self) -> Option<&crate::opc::PreparedOoxmlSource> {
        self.prepared.as_ref()
    }

    /// Take the prepared OOXML package handle, if there is one.
    ///
    /// To open the package as a [`File`](crate::File), pass the whole
    /// detection to [`open_prepared`](crate::open_prepared) instead. The
    /// handle alone yields the indexed package through
    /// `PreparedOoxmlSource::into_package`, which the OOXML leaf crates
    /// adopt.
    #[must_use]
    pub fn into_prepared(self) -> Option<crate::opc::PreparedOoxmlSource> {
        self.prepared
    }

    /// The stable path origin to carry into the unified handle, if this
    /// detection retained a prepared package from a path.
    pub(crate) fn origin(&self) -> Option<std::sync::Arc<std::path::Path>> {
        self.origin.clone()
    }

    /// Split into the neutral classification and the prepared handle.
    #[must_use]
    pub fn into_parts(
        self,
    ) -> (
        litchi_core::detection::FileFormat,
        Option<crate::opc::PreparedOoxmlSource>,
    ) {
        (self.format, self.prepared)
    }
}

/// Classify a filesystem path and retain the OOXML package the classification
/// built, with an explicit OPC resource policy.
///
/// The classification is always the one
/// [`detect_file_format`](crate::detect::detect_file_format) reports:
/// every path that cannot produce a prepared handle — a non-ZIP input, an
/// ordinary OpenDocument package, a ZIP that is not a valid OPC package, an
/// OPC package with no OOXML main part, or any bounded-read failure — defers
/// to that detector and returns the format with no handle.
#[cfg(all(
    any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"),
    any(unix, windows)
))]
pub(crate) fn detect_and_prepare_path<T: ZipTail>(
    path: &std::path::Path,
    limits: &SourceLimits,
    tail: &T,
) -> Option<PreparedDetection> {
    use litchi_core::ReadAt;

    // Every early exit reports exactly what the neutral detector reports. The
    // prepared handle is an optimization; it never changes a classification.
    let unprepared = || {
        super::functions::detect_file_format_with_tail(path, limits, tail)
            .map(PreparedDetection::unprepared)
    };

    let Ok(file_source) = litchi_core::FileSource::open(path) else {
        return None;
    };
    let source: std::sync::Arc<dyn ReadAt> = std::sync::Arc::new(file_source);
    let Ok(source_version) = source.version() else {
        return unprepared();
    };

    let mut signature = [0_u8; 4];
    let Ok(read) = source.read_at(0, &mut signature) else {
        return unprepared();
    };
    if read != signature.len()
        || !litchi_core::detection::simd_utils::signature_matches(
            &signature,
            litchi_core::detection::utils::ZIP_SIGNATURE,
        )
    {
        return unprepared();
    }

    let Ok(input_bytes) = source.len() else {
        return unprepared();
    };
    if !limits.accepts_input_len(input_bytes) {
        return unprepared();
    }

    // Keep the reader path's container precedence: a validated ordinary
    // OpenDocument package is never probed as OPC.
    #[cfg(any(feature = "odt", feature = "ods", feature = "odp", odf_objects))]
    {
        let catalog_limits = limits.odf_catalog_limits();
        match litchi_odf_common::detect::packaged_mime_read_at(
            source.as_ref(),
            limits.odf_detect_limits(),
        ) {
            Ok(Some(_odf_format)) => {
                match litchi_odf_common::detect::packaged_has_ooxml_catalog_read_at(
                    source.as_ref(),
                    catalog_limits,
                ) {
                    Ok(Some(false)) => return unprepared(),
                    Ok(_) => {},
                    Err(_error) => return unprepared(),
                }
            },
            Ok(None) => {},
            Err(_error) => return unprepared(),
        }
    }

    if ensure_path_source_current(source.as_ref(), source_version).is_err() {
        return unprepared();
    }

    #[cfg(any(test, feature = "test-support"))]
    super::record_opc_probe();
    let Ok(package) = crate::opc::SourceBackedPackage::from_read_at(
        std::sync::Arc::clone(&source),
        limits.ooxml(),
        litchi_opc::SourceCacheLimits::default(),
        litchi_opc::SourceReadPolicy::exact(),
        None,
    ) else {
        return unprepared();
    };

    let classification = match super::ooxml::try_classify_ooxml_source_backed_package(&package) {
        // The borrowed catalog spelling is copied here so the package can
        // move into the prepared handle.
        Ok(Some((format, content_type))) => (format, content_type.to_owned()),
        Ok(None) | Err(_) => {
            drop(package);
            return unprepared();
        },
    };

    if ensure_path_source_current(source.as_ref(), source_version).is_err() {
        drop(package);
        return unprepared();
    }

    let (format, content_type) = classification;
    let Ok(prepared) = crate::opc::PreparedOoxmlSource::new(package, &content_type) else {
        return unprepared();
    };
    Some(PreparedDetection {
        format,
        prepared: Some(prepared),
        origin: stable_path_origin(path),
    })
}

/// Resolve a prepared path's origin against the open-time working directory.
/// Keep the caller's symlink spelling so the save layer can apply its normal
/// symlink refusal policy; no filename is consulted for format detection.
#[cfg(all(
    any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"),
    any(unix, windows)
))]
fn stable_path_origin(path: &std::path::Path) -> Option<std::sync::Arc<std::path::Path>> {
    let absolute = if path.is_absolute() {
        path.to_path_buf()
    } else {
        std::env::current_dir().ok()?.join(path)
    };
    Some(std::sync::Arc::from(absolute.into_boxed_path()))
}

#[cfg(test)]
mod short_signature_tests {
    use crate::detect::SourceLimits;

    fn detect_format_smart(bytes: Vec<u8>) -> Option<super::DetectedFormat> {
        super::detect_format_smart(bytes, &SourceLimits::default())
    }

    #[cfg(feature = "odt")]
    #[test]
    fn odt_smart_detection_transfers_the_prepared_index_to_semantic_open() {
        let mut writer = litchi_odf_common::core::PackageWriter::new(
            litchi_odf_common::core::StreamingArchiveLimits::default(),
        );
        writer
            .set_mimetype(litchi_odf_common::constants::ODF_TEXT)
            .unwrap();
        writer
            .add_file(
                litchi_odf_common::constants::ODF_CONTENT,
                br#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><o:body><o:text><t:p>smart</t:p></o:text></o:body></o:document-content>"#,
            )
            .unwrap();
        let bytes = writer.finish_to_bytes().unwrap();

        let detected = detect_format_smart(bytes).expect("ODT should be detected");
        let super::DetectedFormat::Odt(prepared) = detected else {
            panic!("smart ODT detection did not retain a prepared package");
        };
        let index_identity = prepared.prepared_index_identity();
        let document = litchi_odt::Package::from_prepared_package(prepared)
            .unwrap()
            .document()
            .unwrap();
        assert_eq!(document.prepared_index_identity(), index_identity);
        assert_eq!(document.text().unwrap(), "smart");
    }

    #[test]
    fn short_zip_candidate_is_rejected_without_short_read_failure() {
        assert!(detect_format_smart(b"PK\x03\x04".to_vec()).is_none());
    }

    #[cfg(feature = "rtf")]
    #[test]
    fn minimal_rtf_signature_is_retained_for_the_rtf_owner() {
        match detect_format_smart(br#"{\rtf"#.to_vec()) {
            Some(super::DetectedFormat::Rtf(bytes)) => assert_eq!(bytes, br#"{\rtf"#),
            _ => panic!("minimal RTF signature was not retained"),
        }
    }

    #[cfg(feature = "rtf")]
    #[test]
    fn compressed_rtf_signature_is_retained_without_overriding_zip_precedence() {
        let source = br#"{\rtf1\ansi Compressed RTF\par}"#;
        let compressed = litchi_rtf::transport::compress(
            source,
            true,
            litchi_rtf::transport::CompressionLimits::default(),
        )
        .unwrap();

        match detect_format_smart(compressed.clone()) {
            Some(super::DetectedFormat::Rtf(bytes)) => assert_eq!(bytes, compressed),
            _ => panic!("compressed RTF was not retained"),
        }

        // The bytes at offset 8 are part of a ZIP local header. A matching
        // LZFu marker there must not turn a ZIP candidate into RTF.
        let mut zip_polyglot = b"PK\x03\x04\x00\x00\x00\x00".to_vec();
        zip_polyglot.extend_from_slice(b"LZFu");
        assert!(detect_format_smart(zip_polyglot.clone()).is_none());

        #[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
        assert!(
            super::detect_ooxml_smart(zip_polyglot, &SourceLimits::default(), &super::NoZipTail)
                .is_none()
        );

        let mut ole_polyglot = litchi_core::detection::utils::OLE2_SIGNATURE.to_vec();
        ole_polyglot.extend_from_slice(b"LZFu");
        assert!(detect_format_smart(ole_polyglot.clone()).is_none());

        #[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
        assert!(
            super::detect_ooxml_smart(ole_polyglot, &SourceLimits::default(), &super::NoZipTail)
                .is_none()
        );
    }

    #[cfg(feature = "pages")]
    #[test]
    fn debug_output_names_the_format_without_dumping_source_bytes() {
        let detected = super::DetectedFormat::Pages(b"private document marker".to_vec());
        let debug = format!("{detected:?}");

        assert_eq!(debug, "DetectedFormat(\"Pages\")");
        assert!(!debug.contains("private document marker"));
    }
}
