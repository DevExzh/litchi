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

/// One OOXML main-part family, independent of the polyglot precedence applied
/// when a catalog declares more than one.
#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
#[derive(Clone, Copy, PartialEq, Eq)]
enum OoxmlFamily {
    Word,
    PowerPoint,
    ExcelBinary,
    ExcelXml,
}

#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
impl OoxmlFamily {
    const fn format(self) -> FileFormat {
        match self {
            Self::Word => FileFormat::Docx,
            Self::PowerPoint => FileFormat::Pptx,
            Self::ExcelBinary => FileFormat::Xlsb,
            Self::ExcelXml => FileFormat::Xlsx,
        }
    }
}

/// Classify one validated main-part content type into its OOXML family.
///
/// This is the single content-type table the coordinator owns under ADR 0010.
/// Both the forward catalog scan and the reverse lookup used by the
/// prepared-source adopters read it, so a prepared handle can never disagree
/// with the classification that produced it.
#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
fn ooxml_family_for_content_type(content_type: &str) -> Option<OoxmlFamily> {
    use crate::opc::constants::content_type as ct;

    if content_type.eq_ignore_ascii_case(ct::WML_DOCUMENT_MAIN)
        || content_type.eq_ignore_ascii_case(ct::WML_TEMPLATE_MAIN)
        || content_type.eq_ignore_ascii_case(ct::WML_DOCUMENT_MACRO_MAIN)
        || content_type.eq_ignore_ascii_case(ct::WML_TEMPLATE_MACRO_MAIN)
    {
        return Some(OoxmlFamily::Word);
    }
    if content_type.eq_ignore_ascii_case(ct::PML_PRESENTATION_MAIN)
        || content_type.eq_ignore_ascii_case(ct::PML_SLIDESHOW_MAIN)
        || content_type.eq_ignore_ascii_case(ct::PML_TEMPLATE_MAIN)
        || content_type.eq_ignore_ascii_case(ct::PML_PRES_MACRO_MAIN)
        || content_type.eq_ignore_ascii_case(ct::PML_SLIDESHOW_MACRO_MAIN)
        || content_type.eq_ignore_ascii_case(ct::PML_TEMPLATE_MACRO_MAIN)
    {
        return Some(OoxmlFamily::PowerPoint);
    }
    if content_type.eq_ignore_ascii_case(ct::XLSB_BIN) {
        return Some(OoxmlFamily::ExcelBinary);
    }
    if content_type.eq_ignore_ascii_case(ct::SML_SHEET_MAIN)
        || content_type.eq_ignore_ascii_case(ct::SML_TEMPLATE_MAIN)
        || content_type.eq_ignore_ascii_case(ct::SML_SHEET_MACRO_MAIN)
        || content_type.eq_ignore_ascii_case(ct::SML_TEMPLATE_MACRO_MAIN)
    {
        return Some(OoxmlFamily::ExcelXml);
    }
    None
}

/// Map a validated OOXML main-part content type to the neutral classification.
///
/// Returns `None` for a content type that is not an OOXML main part.
#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
#[must_use]
pub fn format_for_main_content_type(content_type: &str) -> Option<FileFormat> {
    ooxml_family_for_content_type(content_type).map(OoxmlFamily::format)
}

#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
#[derive(Default)]
struct OoxmlContentTypeMarkers<'a> {
    word: Option<&'a str>,
    powerpoint: Option<&'a str>,
    excel_binary: Option<&'a str>,
    excel_xml: Option<&'a str>,
}

#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
impl<'a> OoxmlContentTypeMarkers<'a> {
    fn observe(&mut self, content_type: &'a str) {
        let slot = match ooxml_family_for_content_type(content_type) {
            Some(OoxmlFamily::Word) => &mut self.word,
            Some(OoxmlFamily::PowerPoint) => &mut self.powerpoint,
            Some(OoxmlFamily::ExcelBinary) => &mut self.excel_binary,
            Some(OoxmlFamily::ExcelXml) => &mut self.excel_xml,
            None => return,
        };
        let _ = slot.get_or_insert(content_type);
    }

    /// The winning family and the exact declared content type that selected it.
    fn classification(self) -> Option<(FileFormat, &'a str)> {
        // Keep the established precedence when a producer supplies a polyglot
        // catalog carrying more than one family marker.
        let (family, content_type) = if let Some(content_type) = self.word {
            (OoxmlFamily::Word, content_type)
        } else if let Some(content_type) = self.powerpoint {
            (OoxmlFamily::PowerPoint, content_type)
        } else if let Some(content_type) = self.excel_binary {
            (OoxmlFamily::ExcelBinary, content_type)
        } else if let Some(content_type) = self.excel_xml {
            (OoxmlFamily::ExcelXml, content_type)
        } else {
            return None;
        };
        Some((family.format(), content_type))
    }

    fn format(self) -> Option<FileFormat> {
        self.classification().map(|(format, _content_type)| format)
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
    Ok(try_classify_ooxml_source_backed_package(package)?.map(|(format, _content_type)| format))
}

/// Classify a source-backed OOXML catalog, keeping the exact main-part content
/// type that selected the family.
///
/// The returned content type borrows the package's already parsed catalog, so
/// this repeats neither archive nor XML work. It is what a coordinator retains
/// in an opaque prepared source.
#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
pub fn try_classify_ooxml_source_backed_package(
    package: &litchi_opc::SourceBackedPackage,
) -> crate::opc::Result<Option<(FileFormat, &str)>> {
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
    Ok(markers.classification())
}
