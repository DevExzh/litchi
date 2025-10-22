//! Internal types for document format detection and implementation.

use crate::common::{Error, Result};
use std::io::{Read, Seek};

#[cfg(feature = "ole")]
use crate::ole;

#[cfg(feature = "ooxml")]
use crate::ooxml;

#[cfg(feature = "iwa")]
use zip;

/// A Word document implementation that can be .doc, .docx, or .pages format.
///
/// This enum wraps the format-specific implementations and provides
/// a unified API. Users typically don't interact with this enum directly,
/// but instead use the methods on `Document`.
#[allow(clippy::large_enum_variant)]
pub(super) enum DocumentImpl {
    /// Legacy .doc format
    #[cfg(feature = "ole")]
    Doc(ole::doc::Document, crate::common::Metadata),
    /// Modern .docx format
    #[cfg(feature = "ooxml")]
    Docx(Box<ooxml::docx::Document<'static>>, crate::common::Metadata),
    /// Apple Pages format
    #[cfg(feature = "iwa")]
    Pages(crate::iwa::pages::PagesDocument),
}

/// Document format detection.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[allow(unused)] // The corresponding enum would only be used when the feature is enabled
pub(super) enum DocumentFormat {
    /// Legacy .doc format (OLE2)
    Doc,
    /// Modern .docx format (OOXML/ZIP)
    Docx,
    /// Apple Pages format (IWA/ZIP)
    Pages,
}

/// Detect the document format by reading the file header.
///
/// This function reads the first few bytes of the file to determine if it's
/// an OLE2 file (.doc) or a ZIP file (.docx).
pub(super) fn detect_document_format<R: Read + Seek>(reader: &mut R) -> Result<DocumentFormat> {
    use std::io::SeekFrom;

    // Read the first 8 bytes
    let mut header = [0u8; 8];
    reader.read_exact(&mut header)?;

    // Reset to the beginning
    reader.seek(SeekFrom::Start(0))?;

    detect_document_format_from_signature(&header)
}

/// Detect the document format from a byte buffer.
///
/// This is optimized for in-memory detection without seeking.
#[inline]
pub(super) fn detect_document_format_from_bytes(bytes: &[u8]) -> Result<DocumentFormat> {
    if bytes.len() < 4 {
        return Err(Error::InvalidFormat(
            "File too small to determine format".to_string(),
        ));
    }

    detect_document_format_from_signature(&bytes[0..8.min(bytes.len())])
}

/// Detect format from the signature bytes.
#[inline]
fn detect_document_format_from_signature(header: &[u8]) -> Result<DocumentFormat> {
    // Check for OLE2 signature (D0 CF 11 E0 A1 B1 1A E1)
    if header.len() >= 4 && header[0..4] == [0xD0, 0xCF, 0x11, 0xE0] {
        return Ok(DocumentFormat::Doc);
    }

    // Check for ZIP signature (PK\x03\x04)
    // Note: Both DOCX and Pages are ZIP files, so we return Docx here
    // and will need to distinguish them by inspecting the ZIP contents
    if header.len() >= 4 && header[0..4] == [0x50, 0x4B, 0x03, 0x04] {
        return Ok(DocumentFormat::Docx);
    }

    Err(Error::NotOfficeFile)
}

/// Detect if a ZIP file is a Pages document by checking for iWork markers
#[cfg(feature = "iwa")]
fn is_pages_document<R: Read + Seek>(reader: &mut R) -> bool {
    use std::io::SeekFrom;

    // Try to open as ZIP and look for iWork format indicators
    reader.seek(SeekFrom::Start(0)).ok();

    if let Ok(mut archive) = zip::ZipArchive::new(reader) {
        // Check for Index.zip (older iWork format)
        if archive.by_name("Index.zip").is_ok() {
            return true;
        }

        // Check for Index/ directory with .iwa files (newer iWork format)
        for i in 0..archive.len() {
            if let Ok(file) = archive.by_index(i) {
                let name = file.name();
                if name.starts_with("Index/") && name.ends_with(".iwa") {
                    return true;
                }
            }
        }
    }

    false
}

/// Refine ZIP-based document format detection (DOCX vs Pages)
pub(super) fn refine_document_format<R: Read + Seek>(
    reader: &mut R,
    initial_format: DocumentFormat,
) -> Result<DocumentFormat> {
    use std::io::SeekFrom;

    // Only refine if initial detection was Docx (ZIP file)
    if initial_format != DocumentFormat::Docx {
        return Ok(initial_format);
    }

    reader.seek(SeekFrom::Start(0))?;

    // Check if it's a Pages document
    #[cfg(feature = "iwa")]
    if is_pages_document(reader) {
        reader.seek(SeekFrom::Start(0))?;
        return Ok(DocumentFormat::Pages);
    }

    // Otherwise it's DOCX
    reader.seek(SeekFrom::Start(0))?;
    Ok(DocumentFormat::Docx)
}
