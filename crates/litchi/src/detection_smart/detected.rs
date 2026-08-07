//! Smart format detection with reusable owned results.
//!
//! This module provides the `DetectedFormat` enum and the `detect_format_smart`
//! function. OOXML and OLE results retain their parsed owners; iWork, ODF, and
//! RTF results retain the caller's moved byte buffer for subsequent parsing.

/// Detected format with reusable parsed owners or moved source bytes.
///
/// This enum represents the result of format detection, where each format
/// includes the most reusable representation available at this layer:
/// - OOXML formats (DOCX, PPTX, XLSX, XLSB): include parsed OPC package
/// - OLE2 formats (DOC, PPT, XLS): include parsed OleFile
/// - iWork and ODF formats: include owned bytes after leaf detection
/// - RTF: includes owned bytes
///
/// iWork and ODF leaf detectors may scan a container before a later document
/// parser reads the retained bytes again. No parsing-once guarantee is made for
/// those formats.
#[derive(Debug)]
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

    // ODF formats with validated ZIP archive data (lazy parsing)
    #[cfg(feature = "odt")]
    Odt(Vec<u8>),
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

/// Detect a format while moving the source into a reusable result.
///
/// The result retains a reusable representation for immediate follow-up work:
/// - OOXML files: parse OPC package once and return it
/// - OLE2 files: parse OLE file once and return it
/// - iWork, ODF, and RTF files: return the moved bytes after detection
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
            if let Ok(package) = crate::opc::OpcPackage::from_bytes(&bytes) {
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
        if let Some(format) = litchi_odf_common::detect::bytes(&bytes) {
            return match format {
                #[cfg(feature = "odt")]
                FileFormat::Odt => Some(DetectedFormat::Odt(bytes)),
                #[cfg(feature = "odp")]
                FileFormat::Odp => Some(DetectedFormat::Odp(bytes)),
                #[cfg(feature = "ods")]
                FileFormat::Ods => Some(DetectedFormat::Ods(bytes)),
                _ => None,
            };
        }

        #[cfg(any(feature = "pages", feature = "keynote", feature = "numbers"))]
        if let Ok(Some(format)) = litchi_iwa_detect::bytes(&bytes) {
            #[allow(unreachable_patterns)]
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
    }

    None
}

#[cfg(test)]
mod short_signature_tests {
    use super::detect_format_smart;

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
}
