//! ODF (OpenDocument Format) detection.

use crate::common::detection::FileFormat;
use std::io::{Read, Seek};

#[cfg(feature = "odf")]
use std::io::Cursor;

/// Detect ODF format from mimetype content.
/// Uses the standard ODF MIME type mapping.
pub fn detect_odf_format_from_mimetype(mimetype: &[u8]) -> Option<FileFormat> {
    let mime_str = String::from_utf8_lossy(mimetype).trim().to_string();

    match mime_str.as_str() {
        "application/vnd.oasis.opendocument.text" => Some(FileFormat::Odt),
        "application/vnd.oasis.opendocument.spreadsheet" => Some(FileFormat::Ods),
        "application/vnd.oasis.opendocument.presentation" => Some(FileFormat::Odp),
        _ => None,
    }
}

/// Detect ODF format from byte content.
/// Reads mimetype file directly from ZIP archive.
#[cfg(feature = "odf")]
pub fn detect_odf_format(bytes: &[u8]) -> Option<FileFormat> {
    // Check if it starts with ZIP signature
    if bytes.len() < 4 || &bytes[0..4] != crate::common::detection::utils::ZIP_SIGNATURE {
        return None;
    }

    // Create a cursor to read the ZIP file
    let cursor = Cursor::new(bytes);
    detect_odf_format_from_reader(&mut cursor.clone())
}

#[cfg(not(feature = "odf"))]
pub fn detect_odf_format(_bytes: &[u8]) -> Option<FileFormat> {
    None
}

/// Detect ODF format from a reader.
/// Reads mimetype file directly from ZIP archive.
#[cfg(feature = "odf")]
pub fn detect_odf_format_from_reader<R: Read + Seek>(reader: &mut R) -> Option<FileFormat> {
    // Read the entire file into memory for ZIP analysis
    let mut buffer = Vec::new();
    if reader.read_to_end(&mut buffer).is_err() {
        return None;
    }

    // Reset reader position
    let _ = reader.seek(std::io::SeekFrom::Start(0));

    // Try to open as ZIP archive
    let cursor = Cursor::new(&buffer);
    let zip_result = zip::ZipArchive::new(cursor);

    if let Ok(mut archive) = zip_result {
        // Read MIME type from mimetype file
        if let Ok(mimetype) =
            crate::common::detection::utils::read_zip_file(&mut archive, "mimetype")
        {
            return detect_odf_format_from_mimetype(&mimetype);
        }
    }

    None
}

#[cfg(not(feature = "odf"))]
pub fn detect_odf_format_from_reader<R: Read + Seek>(_reader: &mut R) -> Option<FileFormat> {
    None
}
