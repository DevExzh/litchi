//! OOXML format detection (modern Office documents).
//!
//! This module is only available when the `ooxml` feature is enabled.
//!
//! Uses SIMD-accelerated signature matching for improved performance.

use litchi_core::detection::FileFormat;
#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
use std::io::{Read, Seek};

#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
use litchi_core::detection::simd_utils::signature_matches;

/// Detect ZIP-based OOXML formats from byte content.
/// Uses OpcPackage to properly validate and identify OOXML format.
/// Uses SIMD-accelerated signature matching.
///
/// # Note
/// This function requires the `ooxml` feature to be enabled.
#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
pub fn detect_zip_format(bytes: &[u8]) -> Option<FileFormat> {
    detect_zip_format_with_limits(bytes, crate::opc::ReadLimits::default())
}

/// Detect a ZIP-based OOXML format from bytes with an explicit OPC resource
/// policy.
#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
pub fn detect_zip_format_with_limits(
    bytes: &[u8],
    limits: crate::opc::ReadLimits,
) -> Option<FileFormat> {
    // Check if it starts with ZIP signature using SIMD
    if bytes.len() < 4 || !signature_matches(bytes, litchi_core::detection::utils::ZIP_SIGNATURE) {
        return None;
    }

    // Create a cursor to read the ZIP file
    let cursor = std::io::Cursor::new(bytes);
    detect_zip_format_from_reader_with_limits(&mut cursor.clone(), limits)
}

/// Stub implementation when `ooxml` feature is disabled.
/// Always returns None since OOXML parsing is not available.
#[cfg(not(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb")))]
pub fn detect_zip_format(_bytes: &[u8]) -> Option<FileFormat> {
    None
}

/// Detect ZIP-based formats from a reader.
/// Uses OpcPackage to properly parse and identify OOXML format.
///
/// # Note
/// This function requires the `ooxml` feature to be enabled.
#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
pub fn detect_zip_format_from_reader<R: Read + Seek>(reader: &mut R) -> Option<FileFormat> {
    detect_zip_format_from_reader_with_limits(reader, crate::opc::ReadLimits::default())
}

/// Detect a ZIP-based OOXML format from a reader with an explicit OPC resource
/// policy.
#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
pub fn detect_zip_format_from_reader_with_limits<R: Read + Seek>(
    reader: &mut R,
    limits: crate::opc::ReadLimits,
) -> Option<FileFormat> {
    let package = match crate::opc::OpcPackage::from_reader_with_limits(reader, limits) {
        Ok(pkg) => pkg,
        Err(_) => return None,
    };

    // Determine the specific OOXML format based on content
    detect_ooxml_format_from_package(&package)
}

/// Detect an OOXML format from bytes with the default bounded OPC policy.
#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
pub fn detect_ooxml_format(bytes: &[u8]) -> Option<FileFormat> {
    detect_ooxml_format_with_limits(bytes, crate::opc::ReadLimits::default())
}

/// Detect an OOXML format from bytes with an explicit OPC resource policy.
#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
pub fn detect_ooxml_format_with_limits(
    bytes: &[u8],
    limits: crate::opc::ReadLimits,
) -> Option<FileFormat> {
    detect_zip_format_with_limits(bytes, limits)
}

/// Detect an OOXML format from bytes with the default bounded OPC policy.
#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
pub fn detect_ooxml_format_from_bytes(bytes: &[u8]) -> Option<FileFormat> {
    detect_ooxml_format_from_bytes_with_limits(bytes, crate::opc::ReadLimits::default())
}

/// Detect an OOXML format from bytes with an explicit OPC resource policy.
#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
pub fn detect_ooxml_format_from_bytes_with_limits(
    bytes: &[u8],
    limits: crate::opc::ReadLimits,
) -> Option<FileFormat> {
    detect_ooxml_format_with_limits(bytes, limits)
}

/// Detect specific OOXML format from OpcPackage.
/// Analyzes the package structure to determine the document type.
///
/// # Note
/// This function requires the `ooxml` feature to be enabled.
#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
pub fn detect_ooxml_format_from_package(package: &crate::opc::OpcPackage) -> Option<FileFormat> {
    detect_ooxml_format_from_content_types(|visit| {
        for part in package.iter_parts() {
            visit(part.content_type());
        }
    })
}

fn detect_ooxml_format_from_content_types(
    visit_content_types: impl FnOnce(&mut dyn FnMut(&str)),
) -> Option<FileFormat> {
    let mut word = false;
    let mut powerpoint = false;
    let mut excel_binary = false;
    let mut excel_xml = false;

    visit_content_types(&mut |content_type| {
        word |= content_type.contains("wordprocessingml.document.main")
            || content_type.contains("wordprocessingml.template.main")
            || content_type.contains("ms-word.document.macroEnabled.main")
            || content_type.contains("ms-word.template.macroEnabledTemplate.main");
        powerpoint |= content_type.contains("presentationml.presentation.main")
            || content_type.contains("presentationml.slideshow.main")
            || content_type.contains("presentationml.template.main")
            || content_type.contains("ms-powerpoint.presentation.macroEnabled.main")
            || content_type.contains("ms-powerpoint.slideshow.macroEnabled.main")
            || content_type.contains("ms-powerpoint.template.macroEnabled.main");
        excel_binary |= content_type.contains("ms-excel.sheet.binary.macroEnabled.main");
        excel_xml |= content_type.contains("spreadsheetml.sheet.main")
            || content_type.contains("spreadsheetml.template.main")
            || content_type.contains("ms-excel.sheet.macroEnabled.main")
            || content_type.contains("ms-excel.template.macroEnabled.main");
    });

    // Keep the established precedence when a producer supplies a polyglot
    // catalog carrying more than one family marker.
    if word {
        Some(FileFormat::Docx)
    } else if powerpoint {
        Some(FileFormat::Pptx)
    } else if excel_binary {
        Some(FileFormat::Xlsb)
    } else if excel_xml {
        Some(FileFormat::Xlsx)
    } else {
        None
    }
}

/// Detect an OOXML family from a source-backed OPC catalog without loading
/// any part payload.  Path facades use this metadata-only probe to hand a
/// validated source-backed PPTX owner to the unified API while retaining the
/// existing eager [`detect_ooxml_format_from_package`] path for byte-backed
/// callers and non-PPTX fallbacks.
#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
pub(crate) fn detect_ooxml_format_from_source_backed_package(
    package: &litchi_opc::SourceBackedPackage,
) -> Option<FileFormat> {
    detect_ooxml_format_from_content_types(|visit| {
        for part in package.iter_parts() {
            visit(part.content_type());
        }
    })
}
