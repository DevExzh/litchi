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
/// Uses an OPC package catalog to validate and identify the OOXML format
/// without loading ordinary part payloads. Uses SIMD-accelerated signature
/// matching.
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
    try_detect_zip_format_with_limits(bytes, limits)
        .ok()
        .flatten()
}

/// Detect a ZIP-based OOXML format from bytes with an explicit OPC resource
/// policy, preserving any validation error for callers that need it.
#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
pub fn try_detect_zip_format_with_limits(
    bytes: &[u8],
    limits: crate::opc::ReadLimits,
) -> crate::opc::Result<Option<FileFormat>> {
    // Check if it starts with ZIP signature using SIMD
    if bytes.len() < 4 || !signature_matches(bytes, litchi_core::detection::utils::ZIP_SIGNATURE) {
        return Ok(None);
    }

    // Create a cursor to read the ZIP file
    let mut cursor = std::io::Cursor::new(bytes);
    try_detect_zip_format_from_reader_with_limits(&mut cursor, limits)
}

/// Stub implementation when `ooxml` feature is disabled.
/// Always returns None since OOXML parsing is not available.
#[cfg(not(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb")))]
pub fn detect_zip_format(_bytes: &[u8]) -> Option<FileFormat> {
    None
}

/// Detect ZIP-based formats from a reader.
/// Uses an OPC package catalog to validate and identify the OOXML format
/// without loading ordinary part payloads.
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
    try_detect_zip_format_from_reader_with_limits(reader, limits)
        .ok()
        .flatten()
}

/// Detect a ZIP-based OOXML format from a reader with an explicit OPC resource
/// policy, preserving any validation error for callers that need it.
#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
pub fn try_detect_zip_format_from_reader_with_limits<R: Read + Seek>(
    reader: &mut R,
    limits: crate::opc::ReadLimits,
) -> crate::opc::Result<Option<FileFormat>> {
    let catalog = crate::opc::probe_package_catalog_from_reader_with_limits(reader, limits)?;
    Ok(detect_ooxml_format_from_catalog(&catalog))
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
    detect_ooxml_format_from_content_types(package.iter_parts().map(|part| part.content_type()))
}

#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
fn detect_ooxml_format_from_catalog(catalog: &crate::opc::PackageCatalog) -> Option<FileFormat> {
    detect_ooxml_format_from_content_types(catalog.part_content_types())
}

#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
#[derive(Default)]
struct OoxmlContentTypeMarkers {
    word: bool,
    powerpoint: bool,
    excel_binary: bool,
    excel_xml: bool,
}

#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
impl OoxmlContentTypeMarkers {
    fn observe(&mut self, content_type: &str) {
        use crate::opc::constants::content_type as ct;

        self.word |= content_type.eq_ignore_ascii_case(ct::WML_DOCUMENT_MAIN)
            || content_type.eq_ignore_ascii_case(ct::WML_TEMPLATE_MAIN)
            || content_type.eq_ignore_ascii_case(ct::WML_DOCUMENT_MACRO_MAIN)
            || content_type.eq_ignore_ascii_case(ct::WML_TEMPLATE_MACRO_MAIN);
        self.powerpoint |= content_type.eq_ignore_ascii_case(ct::PML_PRESENTATION_MAIN)
            || content_type.eq_ignore_ascii_case(ct::PML_SLIDESHOW_MAIN)
            || content_type.eq_ignore_ascii_case(ct::PML_TEMPLATE_MAIN)
            || content_type.eq_ignore_ascii_case(ct::PML_PRES_MACRO_MAIN)
            || content_type.eq_ignore_ascii_case(ct::PML_SLIDESHOW_MACRO_MAIN)
            || content_type.eq_ignore_ascii_case(ct::PML_TEMPLATE_MACRO_MAIN);
        self.excel_binary |= content_type.eq_ignore_ascii_case(ct::XLSB_BIN);
        self.excel_xml |= content_type.eq_ignore_ascii_case(ct::SML_SHEET_MAIN)
            || content_type.eq_ignore_ascii_case(ct::SML_TEMPLATE_MAIN)
            || content_type.eq_ignore_ascii_case(ct::SML_SHEET_MACRO_MAIN)
            || content_type.eq_ignore_ascii_case(ct::SML_TEMPLATE_MACRO_MAIN);
    }

    fn format(self) -> Option<FileFormat> {
        // Keep the established precedence when a producer supplies a polyglot
        // catalog carrying more than one family marker.
        if self.word {
            Some(FileFormat::Docx)
        } else if self.powerpoint {
            Some(FileFormat::Pptx)
        } else if self.excel_binary {
            Some(FileFormat::Xlsb)
        } else if self.excel_xml {
            Some(FileFormat::Xlsx)
        } else {
            None
        }
    }
}

#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
fn detect_ooxml_format_from_content_types<'a>(
    content_types: impl Iterator<Item = &'a str>,
) -> Option<FileFormat> {
    let mut markers = OoxmlContentTypeMarkers::default();
    for content_type in content_types {
        markers.observe(content_type);
    }
    markers.format()
}

#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
const SOURCE_CLASSIFICATION_CHECK_INTERVAL: usize = 64;

#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
fn check_source_classification_progress(
    package: &litchi_opc::SourceBackedPackage,
) -> crate::opc::Result<()> {
    package.check_execution()?;
    package.source_version()?;
    Ok(())
}

/// Detect an OOXML family from a source-backed OPC catalog while preserving
/// execution-policy and source-freshness errors.
#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
pub fn try_detect_ooxml_format_from_source_backed_package(
    package: &litchi_opc::SourceBackedPackage,
) -> crate::opc::Result<Option<FileFormat>> {
    check_source_classification_progress(package)?;
    let mut markers = OoxmlContentTypeMarkers::default();
    for (index, content_type) in package
        .iter_parts()
        .map(|part| part.content_type())
        .enumerate()
    {
        if index % SOURCE_CLASSIFICATION_CHECK_INTERVAL == 0 {
            check_source_classification_progress(package)?;
        }
        markers.observe(content_type);
    }
    check_source_classification_progress(package)?;
    Ok(markers.format())
}
