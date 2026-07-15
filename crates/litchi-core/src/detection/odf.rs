//! ODF (OpenDocument Format) detection.
//!
//! This module provides fast, safe detection of OpenDocument Format files
//! by reading the mimetype file from the ZIP archive.
//!
//! ODF files are ZIP archives containing a special "mimetype" file that
//! must be stored uncompressed as the first file in the archive.
//!
//! Uses SIMD-accelerated signature matching for improved performance.

use crate::detection::FileFormat;
use std::io::{Read, Seek};

#[cfg(feature = "odf")]
use crate::detection::simd_utils::signature_matches;

/// Standard ODF MIME types for supported document types.
const ODT_MIME: &str = "application/vnd.oasis.opendocument.text";
const ODT_TEMPLATE_MIME: &str = "application/vnd.oasis.opendocument.text-template";
const ODS_MIME: &str = "application/vnd.oasis.opendocument.spreadsheet";
const ODS_TEMPLATE_MIME: &str = "application/vnd.oasis.opendocument.spreadsheet-template";
const ODP_MIME: &str = "application/vnd.oasis.opendocument.presentation";
const ODP_TEMPLATE_MIME: &str = "application/vnd.oasis.opendocument.presentation-template";
const ODG_MIME: &str = "application/vnd.oasis.opendocument.graphics";
const ODG_TEMPLATE_MIME: &str = "application/vnd.oasis.opendocument.graphics-template";
const ODC_MIME: &str = "application/vnd.oasis.opendocument.chart";
const ODC_TEMPLATE_MIME: &str = "application/vnd.oasis.opendocument.chart-template";
const ODF_MIME: &str = "application/vnd.oasis.opendocument.formula";
const ODF_TEMPLATE_MIME: &str = "application/vnd.oasis.opendocument.formula-template";
const ODI_MIME: &str = "application/vnd.oasis.opendocument.image";
const ODI_TEMPLATE_MIME: &str = "application/vnd.oasis.opendocument.image-template";
const ODM_MIME: &str = "application/vnd.oasis.opendocument.text-master";
const OTH_MIME: &str = "application/vnd.oasis.opendocument.text-web";

/// Detect ODF format from mimetype content.
///
/// Uses the standard ODF MIME type mapping to identify the document type.
/// Supports both regular documents and templates.
///
/// # Arguments
///
/// * `mimetype` - The raw bytes from the mimetype file
///
/// # Returns
///
/// * `Some(FileFormat)` if a valid ODF MIME type is recognized
/// * `None` if the MIME type is not a supported ODF format
///
/// # Examples
///
/// ```rust
/// use litchi_core::detection::odf::detect_odf_format_from_mimetype;
/// use litchi_core::detection::FileFormat;
///
/// let mime = b"application/vnd.oasis.opendocument.text";
/// assert_eq!(detect_odf_format_from_mimetype(mime), Some(FileFormat::Odt));
/// ```
#[inline]
pub fn detect_odf_format_from_mimetype(mimetype: &[u8]) -> Option<FileFormat> {
    // Convert to string, trimming whitespace
    let mime_str = String::from_utf8_lossy(mimetype).trim().to_string();

    // Match against known ODF MIME types
    match mime_str.as_str() {
        ODT_MIME | ODT_TEMPLATE_MIME => Some(FileFormat::Odt),
        ODS_MIME | ODS_TEMPLATE_MIME => Some(FileFormat::Ods),
        ODP_MIME | ODP_TEMPLATE_MIME => Some(FileFormat::Odp),
        ODG_MIME | ODG_TEMPLATE_MIME => Some(FileFormat::Odg),
        ODC_MIME | ODC_TEMPLATE_MIME => Some(FileFormat::Odc),
        ODF_MIME | ODF_TEMPLATE_MIME => Some(FileFormat::Odf),
        ODI_MIME | ODI_TEMPLATE_MIME => Some(FileFormat::Odi),
        ODM_MIME => Some(FileFormat::Odm),
        OTH_MIME => Some(FileFormat::Oth),
        _ => None,
    }
}

/// Detect ODF format from byte content.
///
/// Reads the mimetype file directly from the ZIP archive to determine
/// the specific ODF document type.
///
/// # Arguments
///
/// * `bytes` - The complete file data as bytes
///
/// # Returns
///
/// * `Some(FileFormat)` if a valid ODF format is detected
/// * `None` if not an ODF file or detection fails
///
/// # Performance
///
/// This function performs minimal work:
/// 1. Validates ZIP signature (4 bytes) using SIMD acceleration
/// 2. Opens ZIP archive in-memory
/// 3. Reads only the mimetype file (typically < 100 bytes)
#[cfg(feature = "odf")]
pub fn detect_odf_format(bytes: &[u8]) -> Option<FileFormat> {
    // Quick validation: check ZIP signature using SIMD
    if bytes.len() < 4 || !signature_matches(bytes, crate::detection::utils::ZIP_SIGNATURE) {
        return None;
    }

    // Open archive using soapberry-zip
    let archive = soapberry_zip::office::ArchiveReader::new(bytes).ok()?;

    // ODF files must have a mimetype file as the first entry
    // Read it to determine the specific format
    let mimetype = archive.read_string("mimetype").ok()?;

    detect_odf_format_from_mimetype(mimetype.as_bytes())
}

/// Stub implementation when `odf` feature is disabled.
/// Always returns None since ODF parsing is not available.
#[cfg(not(feature = "odf"))]
pub fn detect_odf_format(_bytes: &[u8]) -> Option<FileFormat> {
    None
}

/// Detect ODF format from a reader.
///
/// Reads the mimetype file directly from the ZIP archive to determine
/// the specific ODF document type.
///
/// # Arguments
///
/// * `reader` - A reader that implements Read + Seek
///
/// # Returns
///
/// * `Some(FileFormat)` if a valid ODF format is detected
/// * `None` if not an ODF file or detection fails
///
/// # Note
///
/// This function resets the reader position to the start before returning.
#[cfg(feature = "odf")]
pub fn detect_odf_format_from_reader<R: Read + Seek>(reader: &mut R) -> Option<FileFormat> {
    use std::io::SeekFrom;

    // Reset to beginning
    reader.seek(SeekFrom::Start(0)).ok()?;

    // Read all data for ArchiveReader
    let mut data = Vec::new();
    reader.read_to_end(&mut data).ok()?;

    // Try to open as ZIP archive
    let archive = soapberry_zip::office::ArchiveReader::new(&data).ok()?;

    // ODF files must have a mimetype file
    // Read it to determine the specific format
    let mimetype = archive.read_string("mimetype").ok()?;

    detect_odf_format_from_mimetype(mimetype.as_bytes())
}

/// Stub implementation when `odf` feature is disabled.
/// Always returns None since ODF parsing is not available.
#[cfg(not(feature = "odf"))]
pub fn detect_odf_format_from_reader<R: Read + Seek>(_reader: &mut R) -> Option<FileFormat> {
    None
}

#[cfg(all(test, feature = "odf"))]
mod tests {
    use super::*;

    #[test]
    fn test_detect_odf_mimetype() {
        // Test ODT detection
        let odt_mime = b"application/vnd.oasis.opendocument.text";
        assert_eq!(
            detect_odf_format_from_mimetype(odt_mime),
            Some(FileFormat::Odt)
        );

        // Test ODS detection
        let ods_mime = b"application/vnd.oasis.opendocument.spreadsheet";
        assert_eq!(
            detect_odf_format_from_mimetype(ods_mime),
            Some(FileFormat::Ods)
        );

        // Test ODP detection
        let odp_mime = b"application/vnd.oasis.opendocument.presentation";
        assert_eq!(
            detect_odf_format_from_mimetype(odp_mime),
            Some(FileFormat::Odp)
        );

        // Test template detection
        let odt_template_mime = b"application/vnd.oasis.opendocument.text-template";
        assert_eq!(
            detect_odf_format_from_mimetype(odt_template_mime),
            Some(FileFormat::Odt)
        );

        // Test non-ODF MIME type
        let non_odf_mime = b"application/pdf";
        assert_eq!(detect_odf_format_from_mimetype(non_odf_mime), None);
    }

    #[test]
    fn test_detect_odf_with_whitespace() {
        // Test with trailing whitespace
        let odt_mime_ws = b"application/vnd.oasis.opendocument.text  \n";
        assert_eq!(
            detect_odf_format_from_mimetype(odt_mime_ws),
            Some(FileFormat::Odt)
        );
    }

    #[test]
    fn detects_all_standard_odf_document_families_and_templates() {
        for (mimetype, expected) in [
            (ODG_MIME, FileFormat::Odg),
            (ODG_TEMPLATE_MIME, FileFormat::Odg),
            (ODC_MIME, FileFormat::Odc),
            (ODC_TEMPLATE_MIME, FileFormat::Odc),
            (ODF_MIME, FileFormat::Odf),
            (ODF_TEMPLATE_MIME, FileFormat::Odf),
            (ODI_MIME, FileFormat::Odi),
            (ODI_TEMPLATE_MIME, FileFormat::Odi),
            (ODM_MIME, FileFormat::Odm),
            (OTH_MIME, FileFormat::Oth),
        ] {
            assert_eq!(
                detect_odf_format_from_mimetype(mimetype.as_bytes()),
                Some(expected),
                "failed to detect {mimetype}"
            );
        }
    }
}
