//! OOXML format detection (modern Office documents).
//!
//! This module is only available when the `ooxml` feature is enabled.
//!
//! Uses SIMD-accelerated signature matching for improved performance.

use litchi_core::detection::FileFormat;
#[cfg(feature = "ooxml")]
use std::io::{Read, Seek};

#[cfg(feature = "ooxml")]
use litchi_core::detection::simd_utils::signature_matches;

/// Detect ZIP-based OOXML formats from byte content.
/// Uses OpcPackage to properly validate and identify OOXML format.
/// Uses SIMD-accelerated signature matching.
///
/// # Note
/// This function requires the `ooxml` feature to be enabled.
#[cfg(feature = "ooxml")]
pub fn detect_zip_format(bytes: &[u8]) -> Option<FileFormat> {
    // Check if it starts with ZIP signature using SIMD
    if bytes.len() < 4 || !signature_matches(bytes, litchi_core::detection::utils::ZIP_SIGNATURE) {
        return None;
    }

    // Create a cursor to read the ZIP file
    let cursor = std::io::Cursor::new(bytes);
    detect_zip_format_from_reader(&mut cursor.clone())
}

/// Stub implementation when `ooxml` feature is disabled.
/// Always returns None since OOXML parsing is not available.
#[cfg(not(feature = "ooxml"))]
pub fn detect_zip_format(_bytes: &[u8]) -> Option<FileFormat> {
    None
}

/// Detect ZIP-based formats from a reader.
/// Uses OpcPackage to properly parse and identify OOXML format.
///
/// # Note
/// This function requires the `ooxml` feature to be enabled.
#[cfg(feature = "ooxml")]
pub fn detect_zip_format_from_reader<R: Read + Seek>(reader: &mut R) -> Option<FileFormat> {
    // Try to open as OOXML package - this will validate the format and structure
    let package = match crate::opc::OpcPackage::from_reader(reader) {
        Ok(pkg) => pkg,
        Err(_) => return None,
    };

    // Determine the specific OOXML format based on content
    detect_ooxml_format_from_package(&package)
}

/// Detect specific OOXML format from OpcPackage.
/// Analyzes the package structure to determine the document type.
///
/// # Note
/// This function requires the `ooxml` feature to be enabled.
#[cfg(feature = "ooxml")]
pub fn detect_ooxml_format_from_package(package: &crate::opc::OpcPackage) -> Option<FileFormat> {
    fn is_word_main(content_type: &str) -> bool {
        content_type.contains("wordprocessingml.document.main")
            || content_type.contains("wordprocessingml.template.main")
            || content_type.contains("ms-word.document.macroEnabled.main")
            || content_type.contains("ms-word.template.macroEnabledTemplate.main")
    }

    fn is_powerpoint_main(content_type: &str) -> bool {
        content_type.contains("presentationml.presentation.main")
            || content_type.contains("presentationml.slideshow.main")
            || content_type.contains("presentationml.template.main")
            || content_type.contains("ms-powerpoint.presentation.macroEnabled.main")
            || content_type.contains("ms-powerpoint.slideshow.macroEnabled.main")
            || content_type.contains("ms-powerpoint.template.macroEnabled.main")
    }

    fn is_excel_binary_main(content_type: &str) -> bool {
        content_type.contains("ms-excel.sheet.binary.macroEnabled.main")
    }

    fn is_excel_xml_main(content_type: &str) -> bool {
        content_type.contains("spreadsheetml.sheet.main")
            || content_type.contains("spreadsheetml.template.main")
            || content_type.contains("ms-excel.sheet.macroEnabled.main")
            || content_type.contains("ms-excel.template.macroEnabled.main")
    }

    if package
        .iter_parts()
        .any(|part| is_word_main(part.content_type()))
    {
        return Some(FileFormat::Docx);
    }

    if package
        .iter_parts()
        .any(|part| is_powerpoint_main(part.content_type()))
    {
        return Some(FileFormat::Pptx);
    }

    if package
        .iter_parts()
        .any(|part| is_excel_binary_main(part.content_type()))
    {
        return Some(FileFormat::Xlsb);
    }

    if package
        .iter_parts()
        .any(|part| is_excel_xml_main(part.content_type()))
    {
        return Some(FileFormat::Xlsx);
    }

    None
}
