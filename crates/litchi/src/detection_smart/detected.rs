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
    #[cfg(feature = "ooxml")]
    Docx(crate::ooxml::OpcPackage),
    #[cfg(feature = "ooxml")]
    Pptx(crate::ooxml::OpcPackage),
    #[cfg(feature = "ooxml")]
    Xlsx(crate::ooxml::OpcPackage),
    #[cfg(feature = "ooxml")]
    Xlsb(crate::ooxml::OpcPackage),

    // OLE2 formats with parsed OleFile
    #[cfg(feature = "doc")]
    Doc(litchi_cfb::OleFile<std::io::Cursor<Vec<u8>>>),
    #[cfg(feature = "ppt")]
    Ppt(litchi_cfb::OleFile<std::io::Cursor<Vec<u8>>>),
    #[cfg(feature = "xls")]
    Xls(litchi_cfb::OleFile<std::io::Cursor<Vec<u8>>>),

    // iWork formats with validated ZIP archive data (lazy parsing)
    #[cfg(feature = "iwa")]
    Pages(Vec<u8>),
    #[cfg(feature = "iwa")]
    Keynote(Vec<u8>),
    #[cfg(feature = "iwa")]
    Numbers(Vec<u8>),

    // ODF formats with validated ZIP archive data (lazy parsing)
    #[cfg(feature = "odf")]
    Odt(Vec<u8>),
    #[cfg(feature = "odf")]
    Odp(Vec<u8>),
    #[cfg(feature = "odf")]
    Ods(Vec<u8>),
    #[cfg(feature = "odf")]
    Odg(Vec<u8>),
    #[cfg(feature = "odf")]
    Odc(Vec<u8>),
    #[cfg(feature = "odf")]
    Odf(Vec<u8>),
    #[cfg(feature = "odf")]
    Odi(Vec<u8>),
    #[cfg(feature = "odf")]
    Odm(Vec<u8>),
    #[cfg(feature = "odf")]
    Oth(Vec<u8>),
    #[cfg(feature = "odf")]
    Odb(Vec<u8>),
    /// Flat OpenDocument XML with its detected family.
    #[cfg(feature = "odf")]
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
    #[cfg(any(feature = "ooxml", feature = "odf"))]
    use litchi_core::detection::FileFormat;
    use litchi_core::detection::simd_utils::check_office_signatures;

    // Quick signature checks. ZIP has a complete four-byte local-file
    // signature, RTF has a five-byte prefix, and OLE2 is checked only after
    // the classifier proves its full eight-byte signature. Do not make the
    // reusable detector stricter than the shared signature contract.
    if bytes.len() < 4 {
        return None;
    }

    #[cfg(feature = "odf")]
    if let Some(format) = litchi_odf::detect::flat(&bytes) {
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
        #[cfg(feature = "ooxml")]
        {
            if let Ok(package) = crate::ooxml::OpcPackage::from_bytes(&bytes) {
                // Use existing OOXML detection logic
                if let Some(format) =
                    crate::detection_smart::ooxml::detect_ooxml_format_from_package(&package)
                {
                    return match format {
                        FileFormat::Docx => Some(DetectedFormat::Docx(package)),
                        FileFormat::Pptx => Some(DetectedFormat::Pptx(package)),
                        FileFormat::Xlsx => Some(DetectedFormat::Xlsx(package)),
                        FileFormat::Xlsb => Some(DetectedFormat::Xlsb(package)),
                        _ => None,
                    };
                }
            }
        }

        #[cfg(feature = "odf")]
        if let Some(format) = litchi_odf::detect::bytes(&bytes) {
            return match format {
                FileFormat::Odt => Some(DetectedFormat::Odt(bytes)),
                FileFormat::Odp => Some(DetectedFormat::Odp(bytes)),
                FileFormat::Ods => Some(DetectedFormat::Ods(bytes)),
                FileFormat::Odg => Some(DetectedFormat::Odg(bytes)),
                FileFormat::Odc => Some(DetectedFormat::Odc(bytes)),
                FileFormat::Odf => Some(DetectedFormat::Odf(bytes)),
                FileFormat::Odi => Some(DetectedFormat::Odi(bytes)),
                FileFormat::Odm => Some(DetectedFormat::Odm(bytes)),
                FileFormat::Oth => Some(DetectedFormat::Oth(bytes)),
                FileFormat::Odb => Some(DetectedFormat::Odb(bytes)),
                _ => None,
            };
        }

        #[cfg(feature = "iwa")]
        if let Some(format) = litchi_iwa::detect::bytes(&bytes) {
            return Some(match format {
                litchi_iwa::detect::Format::Pages => DetectedFormat::Pages(bytes),
                litchi_iwa::detect::Format::Keynote => DetectedFormat::Keynote(bytes),
                litchi_iwa::detect::Format::Numbers => DetectedFormat::Numbers(bytes),
            });
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

#[cfg(all(test, feature = "odf"))]
mod tests {
    use super::*;
    use litchi_core::detection::FileFormat;
    use std::io::{Cursor, Write};

    fn package_with_mimetype(mimetype: &str) -> Vec<u8> {
        let mut bytes = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(Cursor::new(&mut bytes));
            let options = zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Stored);
            writer.start_file("mimetype", options).unwrap();
            writer.write_all(mimetype.as_bytes()).unwrap();
            writer.finish().unwrap();
        }
        bytes
    }

    #[test]
    fn smart_detection_retains_all_additional_odf_families() {
        for (mimetype, expected) in [
            (
                "application/vnd.oasis.opendocument.graphics",
                FileFormat::Odg,
            ),
            ("application/vnd.oasis.opendocument.chart", FileFormat::Odc),
            (
                "application/vnd.oasis.opendocument.formula",
                FileFormat::Odf,
            ),
            ("application/vnd.oasis.opendocument.image", FileFormat::Odi),
            (
                "application/vnd.oasis.opendocument.text-master",
                FileFormat::Odm,
            ),
            (
                "application/vnd.oasis.opendocument.text-web",
                FileFormat::Oth,
            ),
            ("application/vnd.oasis.opendocument.base", FileFormat::Odb),
        ] {
            let bytes = package_with_mimetype(mimetype);
            let detected = detect_format_smart(bytes.clone()).unwrap();
            let (format, retained) = match detected {
                DetectedFormat::Odg(retained) => (FileFormat::Odg, retained),
                DetectedFormat::Odc(retained) => (FileFormat::Odc, retained),
                DetectedFormat::Odf(retained) => (FileFormat::Odf, retained),
                DetectedFormat::Odi(retained) => (FileFormat::Odi, retained),
                DetectedFormat::Odm(retained) => (FileFormat::Odm, retained),
                DetectedFormat::Oth(retained) => (FileFormat::Oth, retained),
                DetectedFormat::Odb(retained) => (FileFormat::Odb, retained),
                _ => panic!("wrong smart-detection result for {mimetype}"),
            };
            assert_eq!(format, expected);
            assert_eq!(retained, bytes);
        }
    }

    #[test]
    fn smart_detection_keeps_flat_odf_distinct_from_packages() {
        let xml = br#"<?xml version="1.0"?><o:document xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" o:mimetype="application/vnd.oasis.opendocument.graphics"><o:body><o:drawing/></o:body></o:document>"#;
        match detect_format_smart(xml.to_vec()).unwrap() {
            DetectedFormat::FlatOdf(FileFormat::Odg, retained) => {
                assert_eq!(retained, xml);
            },
            _ => panic!("flat ODG was not retained as flat OpenDocument XML"),
        }
    }
}
