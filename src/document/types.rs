//! Internal types for document format detection and implementation.

use crate::common::{Error, Result};
use std::io::{Read, Seek};

#[cfg(feature = "ole")]
use crate::ole;

#[cfg(feature = "ooxml")]
use crate::ooxml;

/// A Word document implementation that can be either .doc or .docx format.
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
}

/// Document format detection.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum DocumentFormat {
    /// Legacy .doc format (OLE2)
    Doc,
    /// Modern .docx format (OOXML/ZIP)
    Docx,
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
        return Err(Error::InvalidFormat("File too small to determine format".to_string()));
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
    if header.len() >= 4 && header[0..4] == [0x50, 0x4B, 0x03, 0x04] {
        return Ok(DocumentFormat::Docx);
    }

    Err(Error::NotOfficeFile)
}

