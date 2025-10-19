//! Internal types for presentation format detection and implementation.

use crate::common::{Error, Result};
use std::io::{Read, Seek};

#[cfg(feature = "ole")]
use crate::ole;

#[cfg(feature = "ooxml")]
use crate::ooxml;

/// Extracted data from a PPTX slide (to avoid lifetime issues).
#[derive(Debug, Clone)]
pub struct PptxSlideData {
    pub text: String,
    pub name: Option<String>,
}

/// Extracted data from a PPT slide (to avoid lifetime issues).
#[derive(Debug, Clone)]
pub struct PptSlideData {
    pub text: String,
    pub slide_number: usize,
    pub shape_count: usize,
}

/// A PowerPoint presentation implementation that can be either .ppt or .pptx format.
///
/// This enum wraps the format-specific implementations and provides
/// a unified API. Users typically don't interact with this enum directly,
/// but instead use the methods on `Presentation`.
pub(super) enum PresentationImpl {
    /// Legacy .ppt format
    #[cfg(feature = "ole")]
    Ppt(ole::ppt::Presentation),
    /// Modern .pptx format
    #[cfg(feature = "ooxml")]
    Pptx(Box<ooxml::pptx::Presentation<'static>>),
}

/// Presentation format detection.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum PresentationFormat {
    /// Legacy .ppt format (OLE2)
    Ppt,
    /// Modern .pptx format (OOXML/ZIP)
    Pptx,
}

/// Detect the presentation format by reading the file header.
///
/// This function reads the first few bytes of the file to determine if it's
/// an OLE2 file (.ppt) or a ZIP file (.pptx).
pub(super) fn detect_presentation_format<R: Read + Seek>(reader: &mut R) -> Result<PresentationFormat> {
    use std::io::SeekFrom;

    // Read the first 8 bytes
    let mut header = [0u8; 8];
    reader.read_exact(&mut header)?;
    
    // Reset to the beginning
    reader.seek(SeekFrom::Start(0))?;

    detect_presentation_format_from_signature(&header)
}

/// Detect the presentation format from a byte buffer.
///
/// This is optimized for in-memory detection without seeking.
#[inline]
pub(super) fn detect_presentation_format_from_bytes(bytes: &[u8]) -> Result<PresentationFormat> {
    if bytes.len() < 4 {
        return Err(Error::InvalidFormat("File too small to determine format".to_string()));
    }
    
    detect_presentation_format_from_signature(&bytes[0..8.min(bytes.len())])
}

/// Detect format from the signature bytes.
#[inline]
fn detect_presentation_format_from_signature(header: &[u8]) -> Result<PresentationFormat> {
    // Check for OLE2 signature (D0 CF 11 E0 A1 B1 1A E1)
    if header.len() >= 4 && header[0..4] == [0xD0, 0xCF, 0x11, 0xE0] {
        return Ok(PresentationFormat::Ppt);
    }

    // Check for ZIP signature (PK\x03\x04)
    if header.len() >= 4 && header[0..4] == [0x50, 0x4B, 0x03, 0x04] {
        return Ok(PresentationFormat::Pptx);
    }

    Err(Error::NotOfficeFile)
}

