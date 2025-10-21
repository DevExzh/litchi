//! OOXML format detection (modern Office documents).
//!
//! This module is only available when the `ooxml` feature is enabled.

use crate::common::detection::FileFormat;
use std::io::{Read, Seek};

/// Detect ZIP-based OOXML formats from byte content.
/// Uses OpcPackage to properly validate and identify OOXML format.
///
/// # Note
/// This function requires the `ooxml` feature to be enabled.
#[cfg(feature = "ooxml")]
pub fn detect_zip_format(bytes: &[u8]) -> Option<FileFormat> {
    // Check if it starts with ZIP signature
    if bytes.len() < 4 || &bytes[0..4] != crate::common::detection::utils::ZIP_SIGNATURE {
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
    let package = match crate::ooxml::OpcPackage::from_reader(reader) {
        Ok(pkg) => pkg,
        Err(_) => return None,
    };

    // Determine the specific OOXML format based on content
    detect_ooxml_format_from_package(&package)
}

/// Stub implementation when `ooxml` feature is disabled.
/// Always returns None since OOXML parsing is not available.
#[cfg(not(feature = "ooxml"))]
pub fn detect_zip_format_from_reader<R: Read + Seek>(_reader: &mut R) -> Option<FileFormat> {
    None
}

/// Detect specific OOXML format from OpcPackage.
/// Analyzes the package structure to determine the document type.
///
/// # Note
/// This function requires the `ooxml` feature to be enabled.
#[cfg(feature = "ooxml")]
pub fn detect_ooxml_format_from_package(package: &crate::ooxml::OpcPackage) -> Option<FileFormat> {
    // Check for Word document by looking for document part
    if package.iter_parts().any(|part| {
        part.content_type().contains("wordprocessingml")
            || part.content_type().contains("document.main")
    }) {
        return Some(FileFormat::Docx);
    }

    // Check for PowerPoint presentation by looking for presentation part
    if package.iter_parts().any(|part| {
        part.content_type().contains("presentationml")
            || part.content_type().contains("presentation.main")
    }) {
        return Some(FileFormat::Pptx);
    }

    // Check for Excel spreadsheet by looking for workbook part
    if package.iter_parts().any(|part| {
        part.content_type().contains("spreadsheetml")
            || part.content_type().contains("worksheet")
            || part.content_type().contains("workbook")
    }) {
        // Check if it's XLSB (binary) by looking for binary parts
        let has_binary_parts = package.iter_parts().any(|part| {
            part.content_type().contains("binary") || part.content_type().contains("xlsb")
        });

        if has_binary_parts {
            return Some(FileFormat::Xlsb);
        } else {
            return Some(FileFormat::Xlsx);
        }
    }

    None
}

/// Stub implementation when `ooxml` feature is disabled.
/// Always returns None since OOXML parsing is not available.
#[cfg(not(feature = "ooxml"))]
pub fn detect_ooxml_format_from_package(_package: &()) -> Option<FileFormat> {
    None
}
