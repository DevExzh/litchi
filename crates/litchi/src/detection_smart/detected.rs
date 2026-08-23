//! Smart format detection with reusable owned results.
//!
//! This module provides the `DetectedFormat` enum and the `detect_format_smart`
//! function. OOXML and OLE results retain their parsed owners; packaged ODT
//! results retain a validated archive index; ODS, ODP, iWork, and RTF results
//! retain the caller's moved byte buffer for subsequent parsing.

/// Detected format with reusable parsed owners or moved source bytes.
///
/// This enum represents the result of format detection, where each format
/// includes the most reusable representation available at this layer:
/// - OOXML formats (DOCX, PPTX, XLSX, XLSB): include parsed OPC package
/// - OLE2 formats (DOC, PPT, XLS): include parsed OleFile
/// - ODT: includes an owned, validated package index after detection
/// - ODS and ODP: include owned bytes after package detection
/// - iWork formats: include owned bytes after leaf detection
/// - RTF: includes owned bytes
///
/// iWork leaf detectors may scan a container before a later document parser
/// reads the retained bytes again. Packaged ODT detection retains its bounded
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
    #[cfg(any(feature = "odt", feature = "ods", feature = "odp"))]
    FlatOdf(litchi_core::detection::FileFormat, Vec<u8>),

    // RTF format (plain text, no parsing structure needed)
    #[cfg(feature = "rtf")]
    Rtf(Vec<u8>),
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
            #[cfg(any(feature = "odt", feature = "ods", feature = "odp"))]
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
    #[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
    {
        detect_format_smart_with_limits(bytes, crate::opc::ReadLimits::default())
    }

    #[cfg(not(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb")))]
    detect_format_smart_without_ooxml(bytes)
}

/// Prepare an ODS package for the unified spreadsheet facade without changing
/// the public [`DetectedFormat`] enum's legacy byte payload.
///
/// This is deliberately an internal handoff: callers that use
/// [`detect_format_smart`] continue to receive `DetectedFormat::Ods(Vec<u8>)`.
/// The unified facade uses this fast path only after the local ODF MIME entry
/// identifies ODS. When ODF wins, the accepted package's one ODF
/// `OwnedPackage` index is transferred to the typed ODS owner without a
/// second ODF semantic index scan. Other formats return their original bytes
/// for the ordinary detector. In builds with any OOXML probe feature enabled,
/// a cheap central-directory catalog check gates the same bounded OOXML probe
/// as smart detection. This preserves OOXML-first precedence for valid
/// OOXML/ODF polyglots while avoiding an unrelated OPC scan for ordinary ODF
/// packages; the probe is an independent OPC scan and is not counted as an
/// ODF index.
#[cfg(feature = "ods")]
pub(crate) fn detect_prepared_ods(
    bytes: Vec<u8>,
) -> std::result::Result<litchi_odf_common::PreparedPackage, Vec<u8>> {
    use litchi_core::detection::FileFormat;

    if litchi_odf_common::detect::packaged_mime(&bytes) != Some(FileFormat::Ods) {
        return Err(bytes);
    }
    #[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
    if litchi_odf_common::detect::packaged_has_ooxml_catalog(&bytes) == Some(true)
        && ooxml_probe_wins(&bytes)
    {
        return Err(bytes);
    }
    let prepared = litchi_odf_common::detect::prepared_or_original(bytes)?;
    if prepared.format() == FileFormat::Ods {
        Ok(prepared)
    } else {
        Err(prepared.into_package().into_inner())
    }
}

/// Result of the private source-backed XLSX bytes probe.
#[cfg(feature = "xlsx")]
#[allow(
    clippy::large_enum_variant,
    reason = "the valid XLSX handoff moves the existing source-backed package without an extra allocation"
)]
pub(crate) enum WorkbookSourceBytesDetection {
    /// A validated source-retaining OPC owner whose catalog identifies XLSX.
    Xlsx(crate::opc::SourceBackedPackage),
    /// The original bytes for the established byte-backed detector.
    Fallback(Vec<u8>),
}

/// Probe owned bytes through the source-backed OPC catalog, retaining the
/// original allocation for every non-XLSX fallback.
///
/// This is deliberately narrower than [`detect_format_smart`]: the public
/// smart-detector result remains source-compatible and eager, while the
/// unified workbook facade can adopt the existing read-only XLSX catalog
/// owner. ODS precedence is resolved by the caller before this probe.
#[cfg(feature = "xlsx")]
pub(crate) fn detect_workbook_source_bytes(bytes: Vec<u8>) -> WorkbookSourceBytesDetection {
    use litchi_core::ReadAt;
    use std::sync::Arc;

    if bytes.len() < 4
        || !litchi_core::detection::simd_utils::signature_matches(
            &bytes,
            litchi_core::detection::utils::ZIP_SIGNATURE,
        )
    {
        return WorkbookSourceBytesDetection::Fallback(bytes);
    }

    // Keep a reclaimable shared owner so a non-XLSX package can continue
    // through the historical detector without copying its input. The source
    // package is dropped before the reclaim attempt in both fallback paths.
    let shared = Arc::new(bytes);
    let package_result = {
        let source: Arc<dyn ReadAt> =
            Arc::new(litchi_core::OwnedSource::from_arc(Arc::clone(&shared)));
        crate::opc::SourceBackedPackage::from_read_at(source)
    };
    let Ok(package) = package_result else {
        return WorkbookSourceBytesDetection::Fallback(reclaim_source_bytes(shared));
    };

    if crate::detection_smart::ooxml::detect_ooxml_format_from_source_backed_package(&package)
        == Some(litchi_core::detection::FileFormat::Xlsx)
    {
        return WorkbookSourceBytesDetection::Xlsx(package);
    }

    drop(package);
    WorkbookSourceBytesDetection::Fallback(reclaim_source_bytes(shared))
}

#[cfg(feature = "xlsx")]
fn reclaim_source_bytes(shared: std::sync::Arc<Vec<u8>>) -> Vec<u8> {
    std::sync::Arc::try_unwrap(shared).unwrap_or_else(|shared| shared.as_ref().clone())
}

/// Result of the private source-backed DOCX bytes probe.
#[cfg(feature = "docx")]
#[allow(
    clippy::large_enum_variant,
    reason = "the valid DOCX handoff moves the existing source-backed package without an extra allocation"
)]
pub(crate) enum DocxSourceBytesDetection {
    /// A validated source-retaining DOCX owner whose catalog identifies DOCX.
    Docx(crate::docx::source_backed::Package),
    /// A validated DOCX catalog whose WordprocessingML semantic opening failed.
    /// The caller reports this typed owner error without reopening eagerly.
    DocxError(crate::docx::Error),
    /// A hard source-backed OPC failure that must not be hidden by format fallback.
    OpcError(crate::opc::OpcError),
    /// The original bytes for the established byte-backed detector.
    Fallback(Vec<u8>),
}

/// Probe owned bytes through the source-backed OPC catalog, retaining the
/// original allocation for every non-DOCX or failed fallback.
///
/// The public smart-detector result remains source-compatible and eager. The
/// unified document facade uses this narrower handoff so an owned DOCX can
/// retain its source identity and defer ordinary document payloads. A catalog
/// that identifies DOCX but fails its semantic-owner admission propagates
/// that typed error instead of retrying through the eager parser.
#[cfg(feature = "docx")]
pub(crate) fn detect_docx_source_bytes(
    bytes: Vec<u8>,
    limits: crate::opc::ReadLimits,
) -> DocxSourceBytesDetection {
    use litchi_core::ReadAt;
    use std::sync::Arc;

    if bytes.len() < 4
        || !litchi_core::detection::simd_utils::signature_matches(
            &bytes,
            litchi_core::detection::utils::ZIP_SIGNATURE,
        )
    {
        return DocxSourceBytesDetection::Fallback(bytes);
    }

    // Keep a reclaimable shared owner so a non-DOCX package or source-catalog
    // failure can continue through the historical detector without copying
    // its input. The source package is dropped before every fallback reclaim;
    // matching DOCX semantic failures are returned directly.
    let shared = Arc::new(bytes);
    let source: Arc<dyn ReadAt> = Arc::new(litchi_core::OwnedSource::from_arc(Arc::clone(&shared)));
    #[cfg(feature = "odt")]
    let source_version = match source.version() {
        Ok(version) => version,
        Err(error) => {
            return DocxSourceBytesDetection::OpcError(crate::opc::OpcError::IoError(error));
        },
    };
    #[cfg(feature = "odt")]
    let is_odt_mime = match litchi_odf_common::detect::packaged_mime_read_at(source.as_ref()) {
        Ok(format) => format == Some(litchi_core::detection::FileFormat::Odt),
        Err(error) => {
            return DocxSourceBytesDetection::OpcError(odt_mime_probe_error_to_opc(error));
        },
    };
    let package_result =
        crate::opc::SourceBackedPackage::from_read_at_with_limits(Arc::clone(&source), limits);
    let package = match package_result {
        Ok(package) => package,
        Err(error) => {
            #[cfg(feature = "odt")]
            if is_odt_mime {
                if missing_ooxml_content_types_error(&error) {
                    drop(source);
                    return DocxSourceBytesDetection::Fallback(reclaim_docx_source_bytes(shared));
                }
                if hard_docx_source_bytes_probe_error(&error) {
                    match odt_source_ooxml_probe_wins(
                        &source,
                        source_version,
                        crate::opc::ReadLimits::default(),
                    ) {
                        Ok(false) => {
                            drop(source);
                            return DocxSourceBytesDetection::Fallback(reclaim_docx_source_bytes(
                                shared,
                            ));
                        },
                        Ok(true) => return DocxSourceBytesDetection::OpcError(error),
                        Err(probe_error) => {
                            return DocxSourceBytesDetection::OpcError(probe_error);
                        },
                    }
                }
                return DocxSourceBytesDetection::OpcError(error);
            }
            if hard_docx_source_bytes_probe_error(&error) {
                return DocxSourceBytesDetection::OpcError(error);
            }
            drop(source);
            return DocxSourceBytesDetection::Fallback(reclaim_docx_source_bytes(shared));
        },
    };

    if crate::detection_smart::ooxml::detect_ooxml_format_from_source_backed_package(&package)
        == Some(litchi_core::detection::FileFormat::Docx)
    {
        return match crate::docx::source_backed::Package::from_source_backed_package(package) {
            Ok(document) => DocxSourceBytesDetection::Docx(document),
            Err(error) => DocxSourceBytesDetection::DocxError(error),
        };
    }

    drop(package);
    drop(source);
    DocxSourceBytesDetection::Fallback(reclaim_docx_source_bytes(shared))
}

#[cfg(feature = "docx")]
fn reclaim_docx_source_bytes(shared: std::sync::Arc<Vec<u8>>) -> Vec<u8> {
    std::sync::Arc::try_unwrap(shared).unwrap_or_else(|shared| shared.as_ref().clone())
}

#[cfg(feature = "docx")]
fn hard_docx_source_bytes_probe_error(error: &crate::opc::OpcError) -> bool {
    matches!(
        error,
        crate::opc::OpcError::InvalidReadLimit { .. }
            | crate::opc::OpcError::ReadLimit { .. }
            | crate::opc::OpcError::Cancelled
            | crate::opc::OpcError::Execution(_)
            | crate::opc::OpcError::IoError(_)
            | crate::opc::OpcError::SourceChanged { .. }
            | crate::opc::OpcError::Allocation { .. }
            | crate::opc::OpcError::CollectionAllocation { .. }
    )
}

#[cfg(all(
    feature = "odt",
    any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb")
))]
fn missing_ooxml_content_types_error(error: &crate::opc::OpcError) -> bool {
    matches!(
        error,
        crate::opc::OpcError::PartNotFound(part)
            if part == "[Content_Types].xml" || part == "/[Content_Types].xml"
    )
}

/// Result of the private source-backed PPTX bytes probe.
#[cfg(feature = "pptx")]
#[allow(
    clippy::large_enum_variant,
    reason = "the valid PPTX handoff moves the existing source-backed package without an extra allocation"
)]
pub(crate) enum PresentationSourceBytesDetection {
    /// A validated source-retaining PPTX owner whose catalog identifies PPTX.
    Pptx(crate::pptx::SourceBackedPresentation),
    /// A validated PPTX catalog whose PresentationML semantic opening failed.
    /// The caller reports this typed owner error without reopening eagerly.
    PptxError(crate::pptx::Error),
    /// The original bytes for the established byte-backed detector.
    Fallback(Vec<u8>),
}

/// Probe owned bytes through the source-backed OPC catalog, retaining the
/// original allocation for every non-PPTX or failed fallback.
///
/// The public smart-detector result remains source-compatible and eager. The
/// unified presentation facade uses this narrower handoff so an owned PPTX
/// can retain its source identity and defer ordinary slide/media payloads.
/// The normal `from_bytes` caller performs the ODP preparation step before
/// this probe. The explicit-limits caller runs this probe before the existing
/// fallback detector; both preserve OOXML-before-ODF/iWork precedence.
#[cfg(feature = "pptx")]
pub(crate) fn detect_presentation_source_bytes(
    bytes: Vec<u8>,
    limits: crate::opc::ReadLimits,
) -> PresentationSourceBytesDetection {
    use litchi_core::ReadAt;
    use std::sync::Arc;

    if bytes.len() < 4
        || !litchi_core::detection::simd_utils::signature_matches(
            &bytes,
            litchi_core::detection::utils::ZIP_SIGNATURE,
        )
    {
        return PresentationSourceBytesDetection::Fallback(bytes);
    }

    // Keep a reclaimable shared owner so a non-PPTX package or source-catalog
    // failure can continue through the historical detector without copying
    // its input. The package is dropped before every fallback reclaim;
    // matching PresentationML semantic failures are returned directly.
    let shared = Arc::new(bytes);
    let package_result = {
        let source: Arc<dyn ReadAt> =
            Arc::new(litchi_core::OwnedSource::from_arc(Arc::clone(&shared)));
        crate::opc::SourceBackedPackage::from_read_at_with_limits(source, limits)
    };
    let Ok(package) = package_result else {
        return PresentationSourceBytesDetection::Fallback(reclaim_presentation_source_bytes(
            shared,
        ));
    };

    if crate::detection_smart::ooxml::detect_ooxml_format_from_source_backed_package(&package)
        == Some(litchi_core::detection::FileFormat::Pptx)
    {
        return match crate::pptx::SourceBackedPresentation::from_source_backed_package(package) {
            Ok(presentation) => PresentationSourceBytesDetection::Pptx(presentation),
            Err(error) => PresentationSourceBytesDetection::PptxError(error),
        };
    }

    drop(package);
    PresentationSourceBytesDetection::Fallback(reclaim_presentation_source_bytes(shared))
}

#[cfg(feature = "pptx")]
fn reclaim_presentation_source_bytes(shared: std::sync::Arc<Vec<u8>>) -> Vec<u8> {
    std::sync::Arc::try_unwrap(shared).unwrap_or_else(|shared| shared.as_ref().clone())
}

#[cfg(all(any(feature = "ods", feature = "xlsx"), any(unix, windows)))]
const UNIFIED_WORKBOOK_MAX_INPUT_BYTES: u64 = 2 * 1024 * 1024 * 1024;

/// Result of the private source-backed workbook path probe.
#[cfg(all(any(feature = "ods", feature = "xlsx"), any(unix, windows)))]
#[allow(
    dead_code,
    reason = "OOXML precedence variants are only constructed when an OOXML probe feature is enabled"
)]
#[allow(
    clippy::large_enum_variant,
    reason = "this private one-shot handoff moves source metadata directly; boxing it would add an allocation to every valid XLSX filesystem open"
)]
pub(crate) enum WorkbookSourcePathDetection {
    /// A validated, source-retaining XLSX owner and its source-backed core
    /// properties projection.
    #[cfg(feature = "xlsx")]
    Xlsx {
        workbook: crate::xlsx::SourceBackedWorkbook,
        metadata: litchi_core::Metadata,
    },
    /// A validated, source-retaining ODS owner.
    #[cfg(feature = "ods")]
    Ods(Box<litchi_ods::SourceBackedSpreadsheet>),
    /// A recognized OOXML family whose owner is enabled in this build,
    /// together with bytes read from the same pinned filesystem source.
    OtherOoxml {
        format: litchi_core::detection::FileFormat,
        bytes: Vec<u8>,
    },
    /// A recognized OOXML family whose owner is disabled in this build.
    DisabledOtherOoxml(litchi_core::detection::FileFormat),
    /// A non-ODS source retained from the same pinned filesystem source for
    /// the existing byte-backed detector and OLE/OOXML precedence.
    Bytes(Vec<u8>),
}

/// Open a filesystem XLSX or ODS through one positional source-backed owner
/// after giving valid OOXML the existing precedence. The byte-backed
/// [`DetectedFormat`] API remains unchanged; this helper is only used by the
/// unified filesystem workbook facade. Other formats retain bytes from the
/// same pinned source for the established eager fallback.
#[cfg(all(any(feature = "ods", feature = "xlsx"), any(unix, windows)))]
pub(crate) fn detect_workbook_source_path(
    path: &std::path::Path,
) -> std::result::Result<WorkbookSourcePathDetection, Box<dyn std::error::Error + Send + Sync>> {
    use litchi_core::ReadAt;
    use std::sync::Arc;

    let source: Arc<dyn ReadAt> = Arc::new(litchi_core::FileSource::open(path)?);
    let source_version = source.version()?;

    #[cfg(feature = "ods")]
    let is_ods = litchi_odf_common::detect::packaged_mime_read_at(source.as_ref())?
        == Some(litchi_core::detection::FileFormat::Ods);
    #[cfg(not(feature = "ods"))]
    let is_ods = false;

    #[cfg(all(
        feature = "ods",
        any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb")
    ))]
    let ordinary_ods = is_ods
        && litchi_odf_common::detect::packaged_has_ooxml_catalog_read_at(source.as_ref())
            .ok()
            .flatten()
            == Some(false);
    #[cfg(all(
        not(feature = "ods"),
        any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb")
    ))]
    let ordinary_ods = false;

    #[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
    {
        let mut signature = [0_u8; 4];
        let read = source.read_at(0, &mut signature)?;
        let zip_magic = read == signature.len()
            && litchi_core::detection::simd_utils::signature_matches(
                &signature,
                litchi_core::detection::utils::ZIP_SIGNATURE,
            );
        let source_package = if zip_magic && !ordinary_ods {
            match crate::opc::SourceBackedPackage::from_read_at_with_limits(
                Arc::clone(&source),
                crate::opc::ReadLimits::default(),
            ) {
                Ok(package) => Some(package),
                Err(error) if !is_ods && hard_workbook_ooxml_probe_error(&error) => {
                    return Err(Box::new(error));
                },
                Err(_) => {
                    ensure_path_source_current(source.as_ref(), source_version)?;
                    None
                },
            }
        } else {
            None
        };

        if let Some(package) = source_package {
            let Some(format) =
                crate::detection_smart::ooxml::detect_ooxml_format_from_source_backed_package(
                    &package,
                )
            else {
                ensure_path_source_current(source.as_ref(), source_version)?;
                return finish_non_ooxml_workbook_source(source, source_version, is_ods);
            };

            #[cfg(feature = "xlsx")]
            if format == litchi_core::detection::FileFormat::Xlsx {
                let metadata = crate::ooxml_common::properties::read_source_backed(&package)?
                    .map(litchi_core::Metadata::from)
                    .unwrap_or_default();
                let workbook =
                    crate::xlsx::SourceBackedWorkbook::from_source_backed_package(package)?;
                let owner_version = workbook.source_version()?;
                if owner_version != source_version {
                    return Err(Box::new(litchi_core::Error::SourceChanged {
                        expected: source_version,
                        observed: owner_version,
                    }));
                }
                return Ok(WorkbookSourcePathDetection::Xlsx { workbook, metadata });
            }

            let enabled = match format {
                #[cfg(feature = "docx")]
                litchi_core::detection::FileFormat::Docx => true,
                #[cfg(feature = "pptx")]
                litchi_core::detection::FileFormat::Pptx => true,
                #[cfg(feature = "xlsx")]
                litchi_core::detection::FileFormat::Xlsx => true,
                #[cfg(feature = "xlsb")]
                litchi_core::detection::FileFormat::Xlsb => true,
                _ => false,
            };
            return if enabled {
                Ok(WorkbookSourcePathDetection::OtherOoxml {
                    format,
                    bytes: read_path_source_bytes(source.as_ref(), source_version)?,
                })
            } else {
                ensure_path_source_current(source.as_ref(), source_version)?;
                Ok(WorkbookSourcePathDetection::DisabledOtherOoxml(format))
            };
        }
    }

    finish_non_ooxml_workbook_source(source, source_version, is_ods)
}

#[cfg(all(
    any(feature = "ods", feature = "xlsx"),
    any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"),
    any(unix, windows)
))]
fn hard_workbook_ooxml_probe_error(error: &crate::opc::OpcError) -> bool {
    matches!(
        error,
        crate::opc::OpcError::InvalidReadLimit { .. }
            | crate::opc::OpcError::ReadLimit { .. }
            | crate::opc::OpcError::Cancelled
            | crate::opc::OpcError::Execution(_)
            | crate::opc::OpcError::IoError(_)
            | crate::opc::OpcError::SourceChanged { .. }
            | crate::opc::OpcError::Allocation { .. }
            | crate::opc::OpcError::CollectionAllocation { .. }
    )
}

#[cfg(all(any(feature = "ods", feature = "xlsx"), any(unix, windows)))]
fn finish_non_ooxml_workbook_source(
    source: std::sync::Arc<dyn litchi_core::ReadAt>,
    source_version: litchi_core::SourceVersion,
    is_ods: bool,
) -> std::result::Result<WorkbookSourcePathDetection, Box<dyn std::error::Error + Send + Sync>> {
    #[cfg(feature = "ods")]
    if is_ods {
        let ods = litchi_ods::SourceBackedSpreadsheet::from_read_at(source)?;
        let owner_version = ods.source_version()?;
        if owner_version != source_version {
            return Err(Box::new(litchi_core::Error::SourceChanged {
                expected: source_version,
                observed: owner_version,
            }));
        }
        return Ok(WorkbookSourcePathDetection::Ods(Box::new(ods)));
    }
    #[cfg(not(feature = "ods"))]
    let _ = is_ods;

    read_path_source_bytes(source.as_ref(), source_version)
        .map(WorkbookSourcePathDetection::Bytes)
        .map_err(|error| Box::new(error) as Box<dyn std::error::Error + Send + Sync>)
}

#[cfg(all(any(feature = "ods", feature = "xlsx"), any(unix, windows)))]
fn ensure_path_source_current(
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

#[cfg(all(any(feature = "ods", feature = "xlsx"), any(unix, windows)))]
fn read_path_source_bytes(
    source: &dyn litchi_core::ReadAt,
    expected: litchi_core::SourceVersion,
) -> litchi_core::Result<Vec<u8>> {
    ensure_path_source_current(source, expected)?;
    let length = source.len()?;
    ensure_path_source_current(source, expected)?;

    if length > UNIFIED_WORKBOOK_MAX_INPUT_BYTES {
        return Err(litchi_core::Error::ResourceLimit(
            litchi_core::ResourceLimit {
                resource: litchi_core::Resource::InputBytes,
                observed: length,
                limit: UNIFIED_WORKBOOK_MAX_INPUT_BYTES,
                scope: std::sync::Arc::from("unified filesystem workbook"),
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
            resource: "unified filesystem workbook source bytes",
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

/// Prepare an ODP package for the unified presentation facade while keeping
/// the public smart-detection enum source-compatible. As with ODS, a cheap
/// central-directory catalog check gates the bounded OOXML probe when any
/// OOXML probe feature is enabled, so OOXML-first polyglot precedence remains
/// unchanged, including for a recognized OOXML leaf whose own facade feature
/// is disabled. When ODF wins, the prepared ODF index transfers to the typed
/// ODP owner without a second ODF semantic index scan.
#[cfg(feature = "odp")]
pub(crate) fn detect_prepared_odp(
    bytes: Vec<u8>,
) -> std::result::Result<litchi_odf_common::PreparedPackage, Vec<u8>> {
    use litchi_core::detection::FileFormat;

    if litchi_odf_common::detect::packaged_mime(&bytes) != Some(FileFormat::Odp) {
        return Err(bytes);
    }
    #[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
    if litchi_odf_common::detect::packaged_has_ooxml_catalog(&bytes) == Some(true)
        && ooxml_probe_wins(&bytes)
    {
        return Err(bytes);
    }
    let prepared = litchi_odf_common::detect::prepared_or_original(bytes)?;
    if prepared.format() == FileFormat::Odp {
        Ok(prepared)
    } else {
        Err(prepared.into_package().into_inner())
    }
}

#[cfg(all(
    any(feature = "ods", feature = "odp"),
    any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb")
))]
fn ooxml_probe_wins(bytes: &[u8]) -> bool {
    let Ok(package) =
        crate::opc::OpcPackage::from_bytes_with_limits(bytes, crate::opc::ReadLimits::default())
    else {
        return false;
    };
    // Mirror the ordinary smart detector: recognition of any OOXML family
    // wins the precedence decision, even when its leaf feature is disabled
    // and the ordinary detector consequently returns `None` for that owner.
    crate::detection_smart::ooxml::detect_ooxml_format_from_package(&package).is_some()
}

#[cfg(not(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb")))]
fn detect_format_smart_without_ooxml(bytes: Vec<u8>) -> Option<DetectedFormat> {
    use litchi_core::detection::simd_utils::check_office_signatures;

    if bytes.len() < 4 {
        return None;
    }

    #[cfg(any(feature = "odt", feature = "ods", feature = "odp"))]
    if let Some(format) = litchi_odf_common::detect::flat(&bytes) {
        return Some(DetectedFormat::FlatOdf(format, bytes));
    }

    let mask = check_office_signatures(&bytes);

    #[cfg(feature = "rtf")]
    if mask.is_rtf() {
        return Some(DetectedFormat::Rtf(bytes));
    }

    #[cfg(any(feature = "doc", feature = "ppt", feature = "xls"))]
    if mask.is_ole2() {
        let cursor = std::io::Cursor::new(bytes);
        if let Ok(ole_file) = litchi_cfb::OleFile::open(cursor) {
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

    if mask.is_zip() {
        #[cfg(any(feature = "odt", feature = "ods", feature = "odp"))]
        let bytes = match litchi_odf_common::detect::prepared_or_original(bytes) {
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
        if let Ok(Some(format)) = litchi_iwa_detect::bytes(&bytes) {
            #[allow(
                unreachable_patterns,
                reason = "match arms are feature-gated; some are unreachable depending on the enabled features"
            )]
            let detected = match format {
                #[cfg(feature = "pages")]
                litchi_iwa_detect::Format::Pages => DetectedFormat::Pages(bytes),
                #[cfg(feature = "keynote")]
                litchi_iwa_detect::Format::Keynote => DetectedFormat::Keynote(bytes),
                #[cfg(feature = "numbers")]
                litchi_iwa_detect::Format::Numbers => DetectedFormat::Numbers(bytes),
                _ => return None,
            };
            return Some(detected);
        }

        #[cfg(not(any(feature = "pages", feature = "keynote", feature = "numbers")))]
        drop(bytes);
    }

    None
}

/// Detect a format while applying an explicit OPC resource policy to an OOXML
/// ZIP candidate.
///
/// The policy is used only when the input is an OOXML package. OLE, RTF, ODF,
/// and iWork detection retains its existing format-specific behavior.
#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
pub fn detect_format_smart_with_limits(
    bytes: Vec<u8>,
    limits: crate::opc::ReadLimits,
) -> Option<DetectedFormat> {
    #[cfg(any(
        any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"),
        any(feature = "odt", feature = "ods", feature = "odp")
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

    #[cfg(any(feature = "odt", feature = "ods", feature = "odp"))]
    if let Some(format) = litchi_odf_common::detect::flat(&bytes) {
        return Some(DetectedFormat::FlatOdf(format, bytes));
    }

    // Classify the fixed signatures together before invoking format parsers.
    let mask = check_office_signatures(&bytes);

    // Check RTF first (simplest check, no parsing needed)
    #[cfg(feature = "rtf")]
    if mask.is_rtf() {
        return Some(DetectedFormat::Rtf(bytes));
    }

    // Check OLE2 signature (DOC, PPT, XLS) - parse OleFile once
    #[cfg(any(feature = "doc", feature = "ppt", feature = "xls"))]
    if mask.is_ole2() {
        let cursor = std::io::Cursor::new(bytes);
        if let Ok(ole_file) = litchi_cfb::OleFile::open(cursor) {
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

    // Check ZIP candidates in the same order as the ordinary detector.
    if mask.is_zip() {
        #[cfg(all(
            any(feature = "odt", feature = "ods", feature = "odp"),
            any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb")
        ))]
        let normal_odf = litchi_odf_common::detect::packaged_mime(&bytes).is_some()
            && litchi_odf_common::detect::packaged_has_ooxml_catalog(&bytes) == Some(false);
        #[cfg(all(
            not(any(feature = "odt", feature = "ods", feature = "odp")),
            any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb")
        ))]
        let normal_odf = false;

        // A successful OOXML probe returns the parsed OPC owner directly.
        #[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
        if !normal_odf
            && let Ok(package) = crate::opc::OpcPackage::from_bytes_with_limits(&bytes, limits)
            && let Some(format) =
                crate::detection_smart::ooxml::detect_ooxml_format_from_package(&package)
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

        #[cfg(any(feature = "odt", feature = "ods", feature = "odp"))]
        let bytes = match litchi_odf_common::detect::prepared_or_original(bytes) {
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
        if let Ok(Some(format)) = litchi_iwa_detect::bytes(&bytes) {
            #[allow(
                unreachable_patterns,
                reason = "match arms are feature-gated; some are unreachable depending on the enabled features"
            )]
            let detected = match format {
                #[cfg(feature = "pages")]
                litchi_iwa_detect::Format::Pages => DetectedFormat::Pages(bytes),
                #[cfg(feature = "keynote")]
                litchi_iwa_detect::Format::Keynote => DetectedFormat::Keynote(bytes),
                #[cfg(feature = "numbers")]
                litchi_iwa_detect::Format::Numbers => DetectedFormat::Numbers(bytes),
                _ => return None,
            };
            return Some(detected);
        }

        #[cfg(not(any(feature = "pages", feature = "keynote", feature = "numbers")))]
        drop(bytes);
    }

    None
}

/// Parse a path with an OOXML filename extension or ZIP magic under an explicit
/// resource policy and retain it only when its main content type identifies an
/// enabled owner.
///
/// A non-OPC ZIP returns `Ok(None)` so the unified facade can continue to the
/// existing ODF and iWork paths. A resource-limit or allocation failure is
/// returned instead of falling back to an unbounded file read.
#[cfg(any(feature = "docx", feature = "pptx"))]
#[allow(
    dead_code,
    reason = "the unified Presentation path uses the source-backed probe on positional platforms; the eager helper remains for Document and non-positional builds"
)]
pub(crate) fn detect_ooxml_path_with_limits(
    path: impl AsRef<std::path::Path>,
    limits: crate::opc::ReadLimits,
) -> crate::opc::Result<Option<DetectedFormat>> {
    let path = path.as_ref();
    let mut file = std::fs::File::open(path)?;
    let mut signature = [0_u8; 4];
    let read = std::io::Read::read(&mut file, &mut signature)?;
    let ooxml_extension = has_ooxml_extension(path);
    let zip_magic = read == signature.len()
        && litchi_core::detection::simd_utils::signature_matches(
            &signature,
            litchi_core::detection::utils::ZIP_SIGNATURE,
        );
    if !ooxml_extension && !zip_magic {
        return Ok(None);
    }

    let input_bytes = file.metadata()?.len();
    if input_bytes > limits.max_input_bytes() {
        return Err(crate::opc::OpcError::ReadLimit {
            resource: crate::opc::ReadResource::InputBytes,
            actual: input_bytes,
            maximum: limits.max_input_bytes(),
        });
    }

    if !zip_magic {
        return Err(crate::opc::OpcError::ZipError(
            "OOXML-suffixed input does not have ZIP magic".to_owned(),
        ));
    }

    match crate::opc::OpcPackage::open_with_limits(path, limits) {
        Ok(package) => Ok(detect_ooxml_package(package)),
        Err(error @ crate::opc::OpcError::ReadLimit { .. })
        | Err(error @ crate::opc::OpcError::Allocation { .. }) => Err(error),
        Err(error) if ooxml_extension => Err(error),
        Err(_) => {
            if crate::detection_smart::detect_file_format(path).is_some() {
                Ok(None)
            } else {
                Err(crate::opc::OpcError::ZipError(
                    "ZIP input is not a supported Office package".to_owned(),
                ))
            }
        },
    }
}

/// Result of the private source-backed PPTX path probe.
#[cfg(all(feature = "pptx", any(unix, windows)))]
pub(crate) enum PptxSourcePathDetection {
    /// A validated, source-retaining PPTX owner.
    Pptx(crate::pptx::SourceBackedPresentation),
    /// A recognized OOXML family whose facade is not a presentation.
    OtherOoxml(litchi_core::detection::FileFormat),
    /// A recognized non-presentation OOXML family whose facade feature is
    /// disabled.  Retaining this result prevents a lower-precedence ODF
    /// marker in the same ZIP from taking ownership while preserving the
    /// byte detector's `NotOfficeFile` result.
    DisabledOtherOoxml(litchi_core::detection::FileFormat),
}

/// Error from the private source-backed PPTX path probe.
#[cfg(all(feature = "pptx", any(unix, windows)))]
#[derive(Debug)]
pub(crate) enum PptxSourcePathError {
    /// A source or OPC catalog failure, retaining its typed OPC error.
    Opc(crate::opc::OpcError),
    /// A validated PPTX catalog failed PresentationML semantic opening.
    Pptx(crate::pptx::Error),
}

/// Result of the private source-backed DOCX path probe.
#[cfg(all(feature = "docx", any(unix, windows)))]
#[allow(
    clippy::large_enum_variant,
    reason = "this private one-shot handoff moves source metadata directly; boxing it would add an allocation to every valid DOCX filesystem open"
)]
pub(crate) enum DocxSourcePathDetection {
    /// A validated, source-retaining DOCX owner.
    Docx(crate::docx::source_backed::Package),
    /// A recognized OOXML family whose facade is not a document.
    OtherOoxml(litchi_core::detection::FileFormat),
    /// A recognized non-document OOXML family whose facade feature is
    /// disabled. Retaining this result prevents a lower-precedence ODF
    /// marker in the same ZIP from taking ownership while preserving the
    /// byte detector's `NotOfficeFile` result.
    DisabledOtherOoxml(litchi_core::detection::FileFormat),
}

/// Error from the private source-backed DOCX path probe.
#[cfg(all(feature = "docx", any(unix, windows)))]
#[derive(Debug)]
pub(crate) enum DocxSourcePathError {
    /// A source or OPC catalog failure, retaining its typed OPC error.
    Opc(crate::opc::OpcError),
    /// A validated DOCX catalog failed WordprocessingML semantic opening.
    Docx(crate::docx::Error),
}

#[cfg(all(feature = "docx", feature = "odt"))]
fn odt_mime_probe_error_to_opc(error: litchi_core::Error) -> crate::opc::OpcError {
    match error {
        litchi_core::Error::Io(error) => crate::opc::OpcError::IoError(error),
        litchi_core::Error::SourceChanged { expected, observed } => {
            crate::opc::OpcError::SourceChanged {
                expected,
                actual: observed,
            }
        },
        error => crate::opc::OpcError::ZipError(error.to_string()),
    }
}

/// Open a filesystem DOCX through one positional source-backed OPC package.
///
/// This private facade handoff keeps byte-backed smart detection unchanged:
/// only the unified filesystem document path retains the source identity and
/// defers ordinary package payloads. Non-OPC packages return `None` so the
/// existing ODF, RTF, OLE, and byte-backed fallback paths remain in control.
/// A valid non-DOCX OPC package is classified privately even when its leaf
/// owner is disabled, so a lower-precedence ODF marker cannot take ownership.
#[cfg(all(feature = "docx", any(unix, windows)))]
pub(crate) fn detect_docx_source_path_with_limits(
    path: &std::path::Path,
    limits: crate::opc::ReadLimits,
) -> std::result::Result<Option<DocxSourcePathDetection>, DocxSourcePathError> {
    use litchi_core::ReadAt;

    let ooxml_extension = has_ooxml_extension(path);
    let source: std::sync::Arc<dyn ReadAt> = std::sync::Arc::new(
        litchi_core::FileSource::open(path)
            .map_err(crate::opc::OpcError::IoError)
            .map_err(DocxSourcePathError::Opc)?,
    );
    #[cfg(feature = "odt")]
    let source_version = source
        .version()
        .map_err(crate::opc::OpcError::IoError)
        .map_err(DocxSourcePathError::Opc)?;
    let mut signature = [0_u8; 4];
    let read = source
        .read_at(0, &mut signature)
        .map_err(crate::opc::OpcError::IoError)
        .map_err(DocxSourcePathError::Opc)?;
    let zip_magic = read == signature.len()
        && litchi_core::detection::simd_utils::signature_matches(
            &signature,
            litchi_core::detection::utils::ZIP_SIGNATURE,
        );
    if !ooxml_extension && !zip_magic {
        return Ok(None);
    }

    #[cfg(feature = "odt")]
    let is_odt_mime = litchi_odf_common::detect::packaged_mime_read_at(source.as_ref())
        .map_err(odt_mime_probe_error_to_opc)
        .map_err(DocxSourcePathError::Opc)?
        == Some(litchi_core::detection::FileFormat::Odt);

    // Match the eager path's candidate policy: arbitrary non-ZIP inputs have
    // already returned `None`; every remaining candidate is checked against
    // the bounded input-byte policy before an OOXML suffix receives its typed
    // ZIP-magic refusal.
    let input_bytes = source
        .len()
        .map_err(crate::opc::OpcError::IoError)
        .map_err(DocxSourcePathError::Opc)?;
    if input_bytes > limits.max_input_bytes() {
        #[cfg(feature = "odt")]
        if is_odt_mime {
            match odt_source_ooxml_probe_wins(
                &source,
                source_version,
                crate::opc::ReadLimits::default(),
            ) {
                Ok(false) => return Ok(None),
                Ok(true) => {},
                Err(error) => return Err(DocxSourcePathError::Opc(error)),
            }
        }
        return Err(DocxSourcePathError::Opc(crate::opc::OpcError::ReadLimit {
            resource: crate::opc::ReadResource::InputBytes,
            actual: input_bytes,
            maximum: limits.max_input_bytes(),
        }));
    }

    if !zip_magic {
        return Err(DocxSourcePathError::Opc(crate::opc::OpcError::ZipError(
            "OOXML-suffixed input does not have ZIP magic".to_owned(),
        )));
    }

    let package = match crate::opc::SourceBackedPackage::from_read_at_with_limits(
        std::sync::Arc::clone(&source),
        limits,
    ) {
        Ok(package) => package,
        Err(error @ crate::opc::OpcError::ReadLimit { .. })
        | Err(error @ crate::opc::OpcError::Allocation { .. }) => {
            #[cfg(feature = "odt")]
            if is_odt_mime {
                match odt_source_ooxml_probe_wins(
                    &source,
                    source_version,
                    crate::opc::ReadLimits::default(),
                ) {
                    Ok(false) => return Ok(None),
                    Ok(true) => {},
                    Err(probe_error) => {
                        return Err(DocxSourcePathError::Opc(probe_error));
                    },
                }
            }
            return Err(DocxSourcePathError::Opc(error));
        },
        Err(error) if ooxml_extension => return Err(DocxSourcePathError::Opc(error)),
        Err(_) => {
            if crate::detection_smart::detect_file_format(path).is_some() {
                return Ok(None);
            }
            return Err(DocxSourcePathError::Opc(crate::opc::OpcError::ZipError(
                "ZIP input is not a supported Office package".to_owned(),
            )));
        },
    };

    let Some(format) =
        crate::detection_smart::ooxml::detect_ooxml_format_from_source_backed_package(&package)
    else {
        return Ok(None);
    };
    if format != litchi_core::detection::FileFormat::Docx {
        // Match the eager detector's feature-gated result: retain the
        // classification even when its leaf owner is disabled, while telling
        // the caller to report `NotOfficeFile` instead of falling through to
        // a lower-precedence package family.
        let enabled_other_owner = match format {
            #[cfg(feature = "pptx")]
            litchi_core::detection::FileFormat::Pptx => true,
            #[cfg(feature = "xlsx")]
            litchi_core::detection::FileFormat::Xlsx => true,
            #[cfg(feature = "xlsb")]
            litchi_core::detection::FileFormat::Xlsb => true,
            _ => false,
        };
        return Ok(Some(if enabled_other_owner {
            DocxSourcePathDetection::OtherOoxml(format)
        } else {
            DocxSourcePathDetection::DisabledOtherOoxml(format)
        }));
    }

    crate::docx::source_backed::Package::from_source_backed_package(package)
        .map(|package| Some(DocxSourcePathDetection::Docx(package)))
        .map_err(DocxSourcePathError::Docx)
}

/// Open a filesystem ODT through one positional source-backed owner.
///
/// The probe does not consult a filename suffix: extensionless and
/// incorrectly suffixed ODT files are accepted when their package MIME
/// identifies ODT. The unified caller separately arbitrates packages that
/// also contain an OOXML content-types catalog, so ordinary ODT packages do
/// not pay for an unrelated second container scan.
#[cfg(all(feature = "odt", any(unix, windows)))]
pub(crate) struct OdtSourcePathCandidate {
    package: litchi_odf_common::core::SourceBackedPackage,
    #[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
    source_version: litchi_core::SourceVersion,
}

#[cfg(all(feature = "odt", any(unix, windows)))]
impl OdtSourcePathCandidate {
    pub(crate) fn has_ooxml_catalog(&self) -> litchi_core::Result<bool> {
        self.package.has_file("[Content_Types].xml")
    }

    pub(crate) fn into_document(self) -> litchi_core::Result<litchi_odt::SourceBackedDocument> {
        litchi_odt::SourceBackedDocument::from_source_package(self.package)
    }

    #[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
    fn source_arc(&self) -> std::sync::Arc<dyn litchi_core::ReadAt> {
        self.package.source_arc()
    }

    #[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
    fn ensure_current_opc(&self) -> std::result::Result<(), crate::opc::OpcError> {
        let actual = self
            .source_arc()
            .version()
            .map_err(crate::opc::OpcError::IoError)?;
        if actual == self.source_version {
            Ok(())
        } else {
            Err(crate::opc::OpcError::SourceChanged {
                expected: self.source_version,
                actual,
            })
        }
    }
}

#[cfg(all(feature = "odt", any(unix, windows)))]
pub(crate) fn detect_odt_source_path(
    path: &std::path::Path,
) -> litchi_core::Result<Option<OdtSourcePathCandidate>> {
    use litchi_core::ReadAt;
    use std::sync::Arc;

    let source: Arc<dyn ReadAt> = Arc::new(litchi_core::FileSource::open(path)?);
    let source_version = source.version()?;
    let mut signature = [0_u8; 4];
    let read = source.read_at(0, &mut signature)?;
    let zip_magic = read == signature.len()
        && litchi_core::detection::simd_utils::signature_matches(
            &signature,
            litchi_core::detection::utils::ZIP_SIGNATURE,
        );
    if !zip_magic {
        ensure_odt_source_current(source.as_ref(), source_version)?;
        return Ok(None);
    }

    let format = litchi_odf_common::detect::packaged_mime_read_at(source.as_ref())?;
    if format != Some(litchi_core::detection::FileFormat::Odt) {
        ensure_odt_source_current(source.as_ref(), source_version)?;
        return Ok(None);
    }

    let package = litchi_odf_common::core::SourceBackedPackage::from_read_at(source)?;
    #[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
    let source_version = package.source_version()?;
    Ok(Some(OdtSourcePathCandidate {
        package,
        #[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
        source_version,
    }))
}

#[cfg(all(
    feature = "odt",
    any(feature = "pptx", feature = "xlsx", feature = "xlsb"),
    not(feature = "docx"),
    any(unix, windows)
))]
pub(crate) fn odt_source_candidate_has_ooxml_owner(
    candidate: &OdtSourcePathCandidate,
) -> std::result::Result<bool, crate::opc::OpcError> {
    candidate.ensure_current_opc()?;
    let source = candidate.source_arc();
    let result = odt_source_ooxml_probe_wins(
        &source,
        candidate.source_version,
        crate::opc::ReadLimits::default(),
    );
    candidate.ensure_current_opc()?;
    result
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

#[cfg(all(
    feature = "odt",
    any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb")
))]
fn odt_source_ooxml_probe_wins(
    source: &std::sync::Arc<dyn litchi_core::ReadAt>,
    expected: litchi_core::SourceVersion,
    limits: crate::opc::ReadLimits,
) -> std::result::Result<bool, crate::opc::OpcError> {
    let package = match crate::opc::SourceBackedPackage::from_read_at_with_limits(
        std::sync::Arc::clone(source),
        limits,
    ) {
        Ok(package) => package,
        Err(error) if missing_ooxml_content_types_error(&error) => return Ok(false),
        Err(error) => return Err(error),
    };
    let wins =
        crate::detection_smart::ooxml::detect_ooxml_format_from_source_backed_package(&package)
            .is_some();
    let actual = source.version().map_err(crate::opc::OpcError::IoError)?;
    if actual != expected {
        return Err(crate::opc::OpcError::SourceChanged { expected, actual });
    }
    Ok(wins)
}

/// Arbitrate an ODT-marked package that also carries an OOXML catalog without
/// reopening its filesystem path or changing physical snapshots.
#[cfg(all(feature = "docx", feature = "odt", any(unix, windows)))]
pub(crate) fn detect_docx_from_odt_source_candidate_with_limits(
    candidate: &OdtSourcePathCandidate,
    limits: crate::opc::ReadLimits,
) -> std::result::Result<Option<DocxSourcePathDetection>, DocxSourcePathError> {
    candidate
        .ensure_current_opc()
        .map_err(DocxSourcePathError::Opc)?;
    let source = candidate.source_arc();
    let input_bytes = source
        .len()
        .map_err(crate::opc::OpcError::IoError)
        .map_err(DocxSourcePathError::Opc)?;
    if input_bytes > limits.max_input_bytes() {
        match odt_source_ooxml_probe_wins(
            &source,
            candidate.source_version,
            crate::opc::ReadLimits::default(),
        ) {
            Ok(false) => return Ok(None),
            Ok(true) => {
                return Err(DocxSourcePathError::Opc(crate::opc::OpcError::ReadLimit {
                    resource: crate::opc::ReadResource::InputBytes,
                    actual: input_bytes,
                    maximum: limits.max_input_bytes(),
                }));
            },
            Err(error) => return Err(DocxSourcePathError::Opc(error)),
        }
    }

    let package = match crate::opc::SourceBackedPackage::from_read_at_with_limits(
        std::sync::Arc::clone(&source),
        limits,
    ) {
        Ok(package) => package,
        Err(error @ crate::opc::OpcError::ReadLimit { .. })
        | Err(error @ crate::opc::OpcError::Allocation { .. }) => {
            match odt_source_ooxml_probe_wins(
                &source,
                candidate.source_version,
                crate::opc::ReadLimits::default(),
            ) {
                Ok(false) => return Ok(None),
                Ok(true) => return Err(DocxSourcePathError::Opc(error)),
                Err(probe_error) => return Err(DocxSourcePathError::Opc(probe_error)),
            }
        },
        Err(error) => return Err(DocxSourcePathError::Opc(error)),
    };

    let Some(format) =
        crate::detection_smart::ooxml::detect_ooxml_format_from_source_backed_package(&package)
    else {
        candidate
            .ensure_current_opc()
            .map_err(DocxSourcePathError::Opc)?;
        return Ok(None);
    };
    candidate
        .ensure_current_opc()
        .map_err(DocxSourcePathError::Opc)?;
    if format != litchi_core::detection::FileFormat::Docx {
        let enabled_other_owner = match format {
            #[cfg(feature = "pptx")]
            litchi_core::detection::FileFormat::Pptx => true,
            #[cfg(feature = "xlsx")]
            litchi_core::detection::FileFormat::Xlsx => true,
            #[cfg(feature = "xlsb")]
            litchi_core::detection::FileFormat::Xlsb => true,
            _ => false,
        };
        return Ok(Some(if enabled_other_owner {
            DocxSourcePathDetection::OtherOoxml(format)
        } else {
            DocxSourcePathDetection::DisabledOtherOoxml(format)
        }));
    }

    let result = crate::docx::source_backed::Package::from_source_backed_package(package)
        .map(|package| Some(DocxSourcePathDetection::Docx(package)))
        .map_err(DocxSourcePathError::Docx);
    candidate
        .ensure_current_opc()
        .map_err(DocxSourcePathError::Opc)?;
    result
}

/// Open a filesystem PPTX through one positional source-backed OPC package.
///
/// This is intentionally a private facade handoff rather than an additional
/// `DetectedFormat` variant: byte-backed smart detection keeps its established
/// eager owner, while the presentation path can retain the source identity
/// and defer ordinary slide/media payloads. A valid non-PPTX OPC package is
/// classified privately even when its leaf owner is disabled, so a lower-
/// precedence ODF marker cannot take ownership. Non-OPC packages return
/// `None` so the existing facade fallback remains in control.
#[cfg(all(feature = "pptx", any(unix, windows)))]
pub(crate) fn detect_pptx_source_path_with_limits(
    path: &std::path::Path,
    limits: crate::opc::ReadLimits,
) -> std::result::Result<Option<PptxSourcePathDetection>, PptxSourcePathError> {
    use litchi_core::ReadAt;

    let ooxml_extension = has_ooxml_extension(path);
    let source = std::sync::Arc::new(
        litchi_core::FileSource::open(path)
            .map_err(crate::opc::OpcError::IoError)
            .map_err(PptxSourcePathError::Opc)?,
    );
    let mut signature = [0_u8; 4];
    let read = source
        .read_at(0, &mut signature)
        .map_err(crate::opc::OpcError::IoError)
        .map_err(PptxSourcePathError::Opc)?;
    let zip_magic = read == signature.len()
        && litchi_core::detection::simd_utils::signature_matches(
            &signature,
            litchi_core::detection::utils::ZIP_SIGNATURE,
        );
    if !ooxml_extension && !zip_magic {
        return Ok(None);
    }

    // Match the eager path's candidate policy: arbitrary non-ZIP inputs have
    // already returned `None`; every remaining candidate is checked against
    // the bounded input-byte policy before an OOXML suffix receives its typed
    // ZIP-magic refusal.
    let input_bytes = source
        .len()
        .map_err(crate::opc::OpcError::IoError)
        .map_err(PptxSourcePathError::Opc)?;
    if input_bytes > limits.max_input_bytes() {
        return Err(PptxSourcePathError::Opc(crate::opc::OpcError::ReadLimit {
            resource: crate::opc::ReadResource::InputBytes,
            actual: input_bytes,
            maximum: limits.max_input_bytes(),
        }));
    }

    #[cfg(feature = "odp")]
    let ordinary_odp = !ooxml_extension
        && litchi_odf_common::detect::packaged_mime_read_at(source.as_ref())
            .ok()
            .flatten()
            == Some(litchi_core::FileFormat::Odp)
        && litchi_odf_common::detect::packaged_has_ooxml_catalog_read_at(source.as_ref())
            .ok()
            .flatten()
            == Some(false);
    #[cfg(not(feature = "odp"))]
    let ordinary_odp = false;
    if ordinary_odp {
        return Ok(None);
    }

    if !zip_magic {
        return Err(PptxSourcePathError::Opc(crate::opc::OpcError::ZipError(
            "OOXML-suffixed input does not have ZIP magic".to_owned(),
        )));
    }

    let package = match crate::opc::SourceBackedPackage::from_read_at_with_limits(source, limits) {
        Ok(package) => package,
        Err(error @ crate::opc::OpcError::ReadLimit { .. })
        | Err(error @ crate::opc::OpcError::Allocation { .. }) => {
            return Err(PptxSourcePathError::Opc(error));
        },
        Err(error) if ooxml_extension => return Err(PptxSourcePathError::Opc(error)),
        Err(_) => {
            if crate::detection_smart::detect_file_format(path).is_some() {
                return Ok(None);
            }
            return Err(PptxSourcePathError::Opc(crate::opc::OpcError::ZipError(
                "ZIP input is not a supported Office package".to_owned(),
            )));
        },
    };

    let Some(format) =
        crate::detection_smart::ooxml::detect_ooxml_format_from_source_backed_package(&package)
    else {
        return Ok(None);
    };
    if format != litchi_core::detection::FileFormat::Pptx {
        // Match the eager detector's feature-gated result: retain the
        // classification even when its leaf owner is disabled, while telling
        // the caller to report `NotOfficeFile` instead of falling through to
        // a lower-precedence package family.
        let enabled_other_owner = match format {
            #[cfg(feature = "docx")]
            litchi_core::detection::FileFormat::Docx => true,
            #[cfg(feature = "xlsx")]
            litchi_core::detection::FileFormat::Xlsx => true,
            #[cfg(feature = "xlsb")]
            litchi_core::detection::FileFormat::Xlsb => true,
            _ => false,
        };
        return Ok(Some(if enabled_other_owner {
            PptxSourcePathDetection::OtherOoxml(format)
        } else {
            PptxSourcePathDetection::DisabledOtherOoxml(format)
        }));
    }

    crate::pptx::SourceBackedPresentation::from_source_backed_package(package)
        .map(|presentation| Some(PptxSourcePathDetection::Pptx(presentation)))
        .map_err(PptxSourcePathError::Pptx)
}

#[cfg(any(feature = "docx", feature = "pptx"))]
fn has_ooxml_extension(path: &std::path::Path) -> bool {
    let Some(extension) = path.extension().and_then(std::ffi::OsStr::to_str) else {
        return false;
    };

    [
        "docx", "docm", "dotx", "dotm", "pptx", "pptm", "ppsx", "ppsm", "potx", "potm", "xlsx",
        "xlsm", "xltx", "xltm", "xlsb",
    ]
    .iter()
    .any(|known| extension.eq_ignore_ascii_case(known))
}

#[cfg(any(feature = "docx", feature = "pptx"))]
#[allow(
    dead_code,
    reason = "retained by the eager OOXML path used by the document facade and non-positional builds"
)]
fn detect_ooxml_package(package: crate::opc::OpcPackage) -> Option<DetectedFormat> {
    use litchi_core::detection::FileFormat;

    let format = crate::detection_smart::ooxml::detect_ooxml_format_from_package(&package)?;
    match format {
        #[cfg(feature = "docx")]
        FileFormat::Docx => Some(DetectedFormat::Docx(package)),
        #[cfg(feature = "pptx")]
        FileFormat::Pptx => Some(DetectedFormat::Pptx(package)),
        #[cfg(feature = "xlsx")]
        FileFormat::Xlsx => Some(DetectedFormat::Xlsx(package)),
        #[cfg(feature = "xlsb")]
        FileFormat::Xlsb => Some(DetectedFormat::Xlsb(package)),
        _ => None,
    }
}

#[cfg(test)]
mod short_signature_tests {
    use super::detect_format_smart;
    #[cfg(any(feature = "docx", feature = "pptx"))]
    use std::io::{Cursor, Write};

    #[cfg(feature = "odt")]
    #[test]
    fn odt_smart_detection_transfers_the_prepared_index_to_semantic_open() {
        let mut writer = litchi_odf_common::core::PackageWriter::new();
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
        let document = litchi_odt::Document::from_prepared_package(prepared).unwrap();
        assert_eq!(document.prepared_index_identity(), index_identity);
        assert_eq!(document.text().unwrap(), "smart");
    }

    #[cfg(all(
        feature = "ods",
        not(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))
    ))]
    #[test]
    fn internal_ods_handoff_transfers_one_prepared_index_to_semantic_open() {
        let mut writer = litchi_odf_common::core::PackageWriter::new();
        writer
            .set_mimetype(litchi_odf_common::constants::ODF_SPREADSHEET)
            .unwrap();
        writer
            .add_file(
                litchi_odf_common::constants::ODF_CONTENT,
                br#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><office:body><office:spreadsheet><table:table table:name="Sheet1"/></office:spreadsheet></office:body></office:document-content>"#,
            )
            .unwrap();
        let bytes = writer.finish_to_bytes().unwrap();

        assert!(matches!(
            detect_format_smart(bytes.clone()),
            Some(super::DetectedFormat::Ods(_))
        ));
        let prepared = super::detect_prepared_ods(bytes).expect("ODS should prepare");
        let index_identity = prepared.prepared_index_identity();
        let spreadsheet = litchi_ods::Spreadsheet::from_prepared_package(prepared).unwrap();

        assert_eq!(spreadsheet.prepared_index_identity(), index_identity);
        assert!(spreadsheet.content_xml().contains("office:spreadsheet"));
    }

    #[cfg(all(
        feature = "ods",
        any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb")
    ))]
    #[test]
    fn ods_handoff_handles_ordinary_and_malformed_ooxml_catalogs() {
        let mut writer = litchi_odf_common::core::PackageWriter::new();
        writer
            .set_mimetype(litchi_odf_common::constants::ODF_SPREADSHEET)
            .unwrap();
        writer
            .add_file(
                litchi_odf_common::constants::ODF_CONTENT,
                br#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:body><office:spreadsheet/></office:body></office:document-content>"#,
            )
            .unwrap();
        let ordinary = writer.finish_to_bytes().unwrap();
        assert!(matches!(
            detect_format_smart(ordinary.clone()),
            Some(super::DetectedFormat::Ods(_))
        ));
        assert!(super::detect_prepared_ods(ordinary).is_ok());

        let mut writer = litchi_odf_common::core::PackageWriter::new();
        writer
            .set_mimetype(litchi_odf_common::constants::ODF_SPREADSHEET)
            .unwrap();
        writer
            .add_file(
                litchi_odf_common::constants::ODF_CONTENT,
                br#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:body><office:spreadsheet/></office:body></office:document-content>"#,
            )
            .unwrap();
        writer
            .add_file("[Content_Types].xml", b"<not-an-opc-types-root/>")
            .unwrap();
        let malformed_catalog = writer.finish_to_bytes().unwrap();
        assert!(matches!(
            detect_format_smart(malformed_catalog.clone()),
            Some(super::DetectedFormat::Ods(_))
        ));
        assert!(super::detect_prepared_ods(malformed_catalog).is_ok());
    }

    #[cfg(all(
        feature = "odp",
        not(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))
    ))]
    #[test]
    fn internal_odp_handoff_transfers_one_prepared_index_to_semantic_open() {
        let mut writer = litchi_odf_common::core::PackageWriter::new();
        writer
            .set_mimetype(litchi_odf_common::constants::ODF_PRESENTATION)
            .unwrap();
        writer
            .add_file(
                litchi_odf_common::constants::ODF_CONTENT,
                br#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:body><office:presentation/></office:body></office:document-content>"#,
            )
            .unwrap();
        let bytes = writer.finish_to_bytes().unwrap();

        assert!(matches!(
            detect_format_smart(bytes.clone()),
            Some(super::DetectedFormat::Odp(_))
        ));
        let prepared = super::detect_prepared_odp(bytes).expect("ODP should prepare");
        let index_identity = prepared.prepared_index_identity();
        let presentation = litchi_odp::Presentation::from_prepared_package(prepared).unwrap();

        assert_eq!(presentation.prepared_index_identity(), index_identity);
        assert!(presentation.content_xml().contains("office:presentation"));
    }

    #[test]
    fn short_zip_candidate_is_rejected_without_short_read_failure() {
        assert!(detect_format_smart(b"PK\x03\x04".to_vec()).is_none());
    }

    #[cfg(feature = "docx")]
    #[test]
    fn source_docx_probe_preserves_non_docx_zip_allocation_for_fallback() {
        let mut output = Cursor::new(Vec::with_capacity(256));
        let mut writer = zip::ZipWriter::new(&mut output);
        writer
            .start_file(
                "plain.txt",
                zip::write::SimpleFileOptions::default()
                    .compression_method(zip::CompressionMethod::Stored),
            )
            .unwrap();
        writer.write_all(b"not a DOCX package").unwrap();
        writer.finish().unwrap();
        let bytes = output.into_inner();
        let pointer = bytes.as_ptr();
        let capacity = bytes.capacity();

        let detected = super::detect_docx_source_bytes(bytes, crate::opc::ReadLimits::default());
        let super::DocxSourceBytesDetection::Fallback(retained) = detected else {
            panic!("non-DOCX ZIP unexpectedly selected the source owner");
        };
        assert_eq!(retained.as_ptr(), pointer);
        assert_eq!(retained.capacity(), capacity);
    }

    #[cfg(feature = "docx")]
    #[test]
    fn source_docx_probe_preserves_valid_non_docx_opc_allocation_for_fallback() {
        use crate::opc::constants::{content_type as ct, relationship_type as rt};
        use crate::opc::{BlobPart, OpcPackage, PackURI, PackageWriter, TargetMode};

        let mut package = OpcPackage::new();
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new("/ppt/presentation.xml").unwrap(),
                ct::PML_PRESENTATION_MAIN.to_owned(),
                br#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>"#
                    .to_vec(),
            )))
            .unwrap();
        package
            .rels_mut()
            .try_add_relationship(
                rt::OFFICE_DOCUMENT.to_owned(),
                "ppt/presentation.xml".to_owned(),
                "rId1".to_owned(),
                TargetMode::Internal,
            )
            .unwrap();
        let bytes = PackageWriter::to_bytes(&package).unwrap();
        let pointer = bytes.as_ptr();
        let capacity = bytes.capacity();

        let detected = super::detect_docx_source_bytes(bytes, crate::opc::ReadLimits::default());
        let super::DocxSourceBytesDetection::Fallback(retained) = detected else {
            panic!("non-DOCX OPC unexpectedly selected the source owner");
        };
        assert_eq!(retained.as_ptr(), pointer);
        assert_eq!(retained.capacity(), capacity);
    }

    #[cfg(feature = "docx")]
    #[test]
    fn source_docx_probe_propagates_matching_owner_errors() {
        use crate::opc::constants::content_type as ct;
        use crate::opc::{BlobPart, OpcPackage, PackURI, PackageWriter};

        let mut package = OpcPackage::new();
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new("/word/document.xml").unwrap(),
                ct::WML_DOCUMENT_MAIN.to_owned(),
                br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#
                    .to_vec(),
            )))
            .unwrap();
        let bytes = PackageWriter::to_bytes(&package).unwrap();

        let detected = super::detect_docx_source_bytes(bytes, crate::opc::ReadLimits::default());
        assert!(matches!(
            detected,
            super::DocxSourceBytesDetection::DocxError(_)
        ));
    }

    #[cfg(feature = "xlsx")]
    #[test]
    fn source_xlsx_probe_preserves_non_xlsx_bytes_for_fallback() {
        let bytes = b"not an XLSX package".to_vec();
        let detected = super::detect_workbook_source_bytes(bytes.clone());
        let super::WorkbookSourceBytesDetection::Fallback(retained) = detected else {
            panic!("non-XLSX bytes unexpectedly selected the source owner");
        };
        assert_eq!(retained, bytes);
    }

    #[cfg(feature = "pptx")]
    #[test]
    fn source_pptx_probe_preserves_non_pptx_zip_allocation_for_fallback() {
        let mut output = Cursor::new(Vec::with_capacity(256));
        let mut writer = zip::ZipWriter::new(&mut output);
        writer
            .start_file(
                "plain.txt",
                zip::write::SimpleFileOptions::default()
                    .compression_method(zip::CompressionMethod::Stored),
            )
            .unwrap();
        writer.write_all(b"not a PPTX package").unwrap();
        writer.finish().unwrap();
        let bytes = output.into_inner();
        let pointer = bytes.as_ptr();
        let capacity = bytes.capacity();

        let detected =
            super::detect_presentation_source_bytes(bytes, crate::opc::ReadLimits::default());
        let super::PresentationSourceBytesDetection::Fallback(retained) = detected else {
            panic!("non-PPTX ZIP unexpectedly selected the source owner");
        };
        assert_eq!(retained.as_ptr(), pointer);
        assert_eq!(retained.capacity(), capacity);
    }

    #[cfg(feature = "pptx")]
    #[test]
    fn source_pptx_probe_preserves_valid_non_pptx_opc_allocation_for_fallback() {
        use crate::opc::constants::{content_type as ct, relationship_type as rt};
        use crate::opc::{BlobPart, OpcPackage, PackURI, PackageWriter, TargetMode};

        let mut package = OpcPackage::new();
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new("/word/document.xml").unwrap(),
                ct::WML_DOCUMENT_MAIN.to_owned(),
                br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#
                    .to_vec(),
            )))
            .unwrap();
        package
            .rels_mut()
            .try_add_relationship(
                rt::OFFICE_DOCUMENT.to_owned(),
                "word/document.xml".to_owned(),
                "rId1".to_owned(),
                TargetMode::Internal,
            )
            .unwrap();
        let bytes = PackageWriter::to_bytes(&package).unwrap();
        let pointer = bytes.as_ptr();
        let capacity = bytes.capacity();

        let detected =
            super::detect_presentation_source_bytes(bytes, crate::opc::ReadLimits::default());
        let super::PresentationSourceBytesDetection::Fallback(retained) = detected else {
            panic!("non-PPTX OPC unexpectedly selected the presentation owner");
        };
        assert_eq!(retained.as_ptr(), pointer);
        assert_eq!(retained.capacity(), capacity);
    }

    #[cfg(feature = "rtf")]
    #[test]
    fn minimal_rtf_signature_is_retained_for_the_rtf_owner() {
        match detect_format_smart(br#"{\rtf"#.to_vec()) {
            Some(super::DetectedFormat::Rtf(bytes)) => assert_eq!(bytes, br#"{\rtf"#),
            _ => panic!("minimal RTF signature was not retained"),
        }
    }

    #[cfg(feature = "pages")]
    #[test]
    fn debug_output_names_the_format_without_dumping_source_bytes() {
        let detected = super::DetectedFormat::Pages(b"private document marker".to_vec());
        let debug = format!("{detected:?}");

        assert_eq!(debug, "DetectedFormat(\"Pages\")");
        assert!(!debug.contains("private document marker"));
    }

    #[cfg(all(feature = "odt", feature = "pages"))]
    #[test]
    fn rejected_odf_probe_preserves_the_pages_source_for_lower_precedence_detection() {
        let bytes = include_bytes!("../../../../test-data/iwork/pages/basic.pages").to_vec();
        let pointer = bytes.as_ptr();
        let capacity = bytes.capacity();

        let detected = detect_format_smart(bytes).expect("Pages fixture should be detected");
        let super::DetectedFormat::Pages(retained) = detected else {
            panic!("ODF rejection must continue to Pages detection");
        };

        assert_eq!(retained.as_ptr(), pointer);
        assert_eq!(retained.capacity(), capacity);
    }
}
