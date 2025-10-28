//! ODF (OpenDocument Format) detection.
//!
//! This module provides fast, safe detection of OpenDocument Format files
//! (.odt, .ods, .odp) by reading the mimetype file from the ZIP archive.
//!
//! ODF files are ZIP archives containing a special "mimetype" file that
//! must be stored uncompressed as the first file in the archive.

use crate::common::detection::FileFormat;
use std::io::{Read, Seek};

#[cfg(feature = "odf")]
use std::io::Cursor;

/// Standard ODF MIME types for supported document types.
const ODT_MIME: &str = "application/vnd.oasis.opendocument.text";
const ODT_TEMPLATE_MIME: &str = "application/vnd.oasis.opendocument.text-template";
const ODS_MIME: &str = "application/vnd.oasis.opendocument.spreadsheet";
const ODS_TEMPLATE_MIME: &str = "application/vnd.oasis.opendocument.spreadsheet-template";
const ODP_MIME: &str = "application/vnd.oasis.opendocument.presentation";
const ODP_TEMPLATE_MIME: &str = "application/vnd.oasis.opendocument.presentation-template";

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
/// use litchi::common::detection::odf::detect_odf_format_from_mimetype;
/// use litchi::common::detection::FileFormat;
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
/// 1. Validates ZIP signature (4 bytes)
/// 2. Opens ZIP archive in-memory
/// 3. Reads only the mimetype file (typically < 100 bytes)
#[cfg(feature = "odf")]
pub fn detect_odf_format(bytes: &[u8]) -> Option<FileFormat> {
    // Quick validation: check ZIP signature
    if bytes.len() < 4 || &bytes[0..4] != crate::common::detection::utils::ZIP_SIGNATURE {
        return None;
    }

    // Create a cursor to read the ZIP file without copying
    let cursor = Cursor::new(bytes);
    let mut archive = zip::ZipArchive::new(cursor).ok()?;

    // ODF files must have a mimetype file as the first entry
    // Read it to determine the specific format
    let mimetype = crate::common::detection::utils::read_zip_file(&mut archive, "mimetype").ok()?;

    detect_odf_format_from_mimetype(&mimetype)
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

    // Try to open as ZIP archive directly from the reader
    let mut archive = zip::ZipArchive::new(reader).ok()?;

    // ODF files must have a mimetype file
    // Read it to determine the specific format
    let mimetype = crate::common::detection::utils::read_zip_file(&mut archive, "mimetype").ok()?;

    detect_odf_format_from_mimetype(&mimetype)
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
}
