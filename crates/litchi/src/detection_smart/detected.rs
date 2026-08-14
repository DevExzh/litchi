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
/// it first performs the same bounded OOXML probe as smart detection,
/// preserving OOXML-first precedence for valid OOXML/ODF polyglots; that probe
/// is an independent OPC scan and is not counted as an ODF index.
#[cfg(feature = "ods")]
pub(crate) fn detect_prepared_ods(
    bytes: Vec<u8>,
) -> std::result::Result<litchi_odf_common::PreparedPackage, Vec<u8>> {
    use litchi_core::detection::FileFormat;

    if litchi_odf_common::detect::packaged_mime(&bytes) != Some(FileFormat::Ods) {
        return Err(bytes);
    }
    #[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
    if ooxml_probe_wins(&bytes) {
        return Err(bytes);
    }
    let prepared = litchi_odf_common::detect::prepared_or_original(bytes)?;
    if prepared.format() == FileFormat::Ods {
        Ok(prepared)
    } else {
        Err(prepared.into_package().into_inner())
    }
}

/// Prepare an ODP package for the unified presentation facade while keeping
/// the public smart-detection enum source-compatible. As with ODS, it first
/// performs the bounded OOXML probe when any OOXML probe feature is enabled so
/// OOXML-first polyglot precedence remains unchanged, including for a
/// recognized OOXML leaf whose own facade feature is disabled. When ODF wins,
/// the prepared ODF index transfers to the typed ODP owner without a second
/// ODF semantic index scan.
#[cfg(feature = "odp")]
pub(crate) fn detect_prepared_odp(
    bytes: Vec<u8>,
) -> std::result::Result<litchi_odf_common::PreparedPackage, Vec<u8>> {
    use litchi_core::detection::FileFormat;

    if litchi_odf_common::detect::packaged_mime(&bytes) != Some(FileFormat::Odp) {
        return Err(bytes);
    }
    #[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
    if ooxml_probe_wins(&bytes) {
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
        // A successful OOXML probe returns the parsed OPC owner directly.
        #[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
        {
            if let Ok(package) = crate::opc::OpcPackage::from_bytes_with_limits(&bytes, limits) {
                // Use existing OOXML detection logic
                if let Some(format) =
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
            }
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

/// Open a filesystem PPTX through one positional source-backed OPC package.
///
/// This is intentionally a private facade handoff rather than an additional
/// `DetectedFormat` variant: byte-backed smart detection keeps its established
/// eager owner, while the presentation path can retain the source identity
/// and defer ordinary slide/media payloads.  A valid non-PPTX OPC package is
/// classified privately when its owner is enabled; disabled owners and
/// non-OPC packages return `None` so the existing facade fallback remains in
/// control.
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
        // Match the old eager detector's feature-gated handoff: when the
        // corresponding non-presentation owner is disabled, leave the path
        // to the ordinary fallback so it still reports NotOfficeFile.
        let enabled_other_owner = match format {
            #[cfg(feature = "docx")]
            litchi_core::detection::FileFormat::Docx => true,
            #[cfg(feature = "xlsx")]
            litchi_core::detection::FileFormat::Xlsx => true,
            #[cfg(feature = "xlsb")]
            litchi_core::detection::FileFormat::Xlsb => true,
            _ => false,
        };
        return if enabled_other_owner {
            Ok(Some(PptxSourcePathDetection::OtherOoxml(format)))
        } else {
            Ok(None)
        };
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
