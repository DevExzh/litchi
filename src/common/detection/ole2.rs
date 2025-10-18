//! OLE2 format detection (legacy Office documents).

use std::io::{Read, Seek};
use crate::common::detection::FileFormat;

/// Detect OLE2-based formats from byte signature.
/// Uses proper OLE file parsing to identify the format.
pub fn detect_ole2_format(bytes: &[u8]) -> Option<FileFormat> {
    // First check if it's a valid OLE file using the standard function
    if !crate::ole::is_ole_file(bytes) {
        return None;
    }

    // Create a cursor to read the OLE file
    let cursor = std::io::Cursor::new(bytes);
    detect_ole2_format_from_reader(&mut cursor.clone())
}

/// Detect OLE2-based formats from a reader.
/// Uses OleFile to parse the OLE structure and identify format.
pub fn detect_ole2_format_from_reader<R: Read + Seek>(
    reader: &mut R
) -> Option<FileFormat> {
    // Try to open as OLE file - this will validate the format and parse structure
    let ole_file = match crate::ole::OleFile::open(reader) {
        Ok(ole) => ole,
        Err(_) => return None,
    };

    // Check for specific streams to determine the format
    // These checks follow the OLE2 specification and known stream names

    // Word document: check for "WordDocument" stream
    if ole_file.exists(&["WordDocument"]) {
        return Some(FileFormat::Doc);
    }

    // PowerPoint: check for "PowerPoint Document" stream
    if ole_file.exists(&["PowerPoint Document"]) {
        return Some(FileFormat::Ppt);
    }

    // Excel: check for "Workbook" or "Book" stream
    if ole_file.exists(&["Workbook"]) || ole_file.exists(&["Book"]) {
        return Some(FileFormat::Xls);
    }

    // If we can open it as OLE but don't recognize specific streams,
    // it's still a valid OLE file - default to DOC
    Some(FileFormat::Doc)
}
