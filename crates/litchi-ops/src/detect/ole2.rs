//! OLE2 format detection (legacy Office documents).
//!
//! This module is available when any CFB-backed format feature is enabled.

use litchi_core::detection::FileFormat;
use std::io::{Read, Seek};

/// Detect OLE2-based formats from byte signature.
/// Uses proper OLE file parsing to identify the format.
///
/// # Note
/// This function requires `doc`, `ppt`, or `xls` to be enabled.
#[cfg(any(feature = "doc", feature = "ppt", feature = "xls"))]
pub(crate) fn detect_ole2_format(bytes: &[u8], limits: &super::SourceLimits) -> Option<FileFormat> {
    let mut cursor = std::io::Cursor::new(bytes);
    detect_ole2_format_from_reader(&mut cursor, limits)
}

/// Stub implementation when all CFB-backed format features are disabled.
/// Always returns None since OLE2 parsing is not available.
#[cfg(not(any(feature = "doc", feature = "ppt", feature = "xls")))]
pub(crate) fn detect_ole2_format(
    _bytes: &[u8],
    _limits: &super::SourceLimits,
) -> Option<FileFormat> {
    None
}

/// Detect OLE2-based formats from a reader.
/// Uses OleFile to parse the OLE structure and identify format.
///
/// # Note
/// This function requires `doc`, `ppt`, or `xls` to be enabled.
#[cfg(any(feature = "doc", feature = "ppt", feature = "xls"))]
pub(crate) fn detect_ole2_format_from_reader<R: Read + Seek>(
    reader: &mut R,
    limits: &super::SourceLimits,
) -> Option<FileFormat> {
    // Try to open as OLE file - this will validate the format and parse structure
    let ole_file = match litchi_cfb::OleFile::open(reader, limits.cfb()) {
        Ok(ole) => ole,
        Err(_) => return None,
    };

    // Check for specific streams to determine the format
    // These checks follow the OLE2 specification and known stream names

    // An encrypted OOXML package is no legacy document, even beside the
    // stub `WordDocument` stream Word 2007 writes (record 0341).
    if ole_file.exists(&["EncryptionInfo"]) && ole_file.exists(&["EncryptedPackage"]) {
        return None;
    }

    // Word document: check for "WordDocument" stream
    #[cfg(feature = "doc")]
    if ole_file.exists(&["WordDocument"]) {
        return Some(FileFormat::Doc);
    }

    // PowerPoint: check for "PowerPoint Document" stream
    #[cfg(feature = "ppt")]
    if ole_file.exists(&["PowerPoint Document"]) {
        return Some(FileFormat::Ppt);
    }

    // Excel: check for "Workbook" or "Book" stream
    #[cfg(feature = "xls")]
    if ole_file.exists(&["Workbook"]) || ole_file.exists(&["Book"]) {
        return Some(FileFormat::Xls);
    }

    // A valid OLE container without a recognized host stream is not enough
    // evidence to classify the file as DOC, PPT, or XLS.
    None
}

/// Stub implementation when all CFB-backed format features are disabled.
/// Always returns None since OLE2 parsing is not available.
#[cfg(not(any(feature = "doc", feature = "ppt", feature = "xls")))]
pub(crate) fn detect_ole2_format_from_reader<R: Read + Seek>(
    _reader: &mut R,
    _limits: &super::SourceLimits,
) -> Option<FileFormat> {
    None
}
