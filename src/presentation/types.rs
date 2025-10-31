//! Internal types for presentation format detection and implementation.

use crate::common::detection::{self, FileFormat};
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

/// A PowerPoint presentation implementation that can be .ppt, .pptx, .key, or .odp format.
///
/// This enum wraps the format-specific implementations and provides
/// a unified API. Users typically don't interact with this enum directly,
/// but instead use the methods on `Presentation`.
#[allow(clippy::large_enum_variant)]
pub(super) enum PresentationImpl {
    /// Legacy .ppt format
    #[cfg(feature = "ole")]
    Ppt(ole::ppt::Presentation),
    /// Modern .pptx format
    #[cfg(feature = "ooxml")]
    Pptx(Box<ooxml::pptx::Presentation<'static>>),
    /// Apple Keynote format
    #[cfg(feature = "iwa")]
    Keynote(crate::iwa::keynote::KeynoteDocument),
    /// OpenDocument Presentation format
    #[cfg(feature = "odf")]
    Odp(crate::odf::Presentation),
}

/// Presentation format detection.
///
/// This enum represents the supported presentation formats in the unified
/// Presentation API. The format is automatically detected from file signatures.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[allow(unused)] // The corresponding enum would only be used when the feature is enabled
pub(super) enum PresentationFormat {
    /// Legacy .ppt format (OLE2)
    Ppt,
    /// Modern .pptx format (OOXML/ZIP)
    Pptx,
    /// Apple Keynote format (IWA/ZIP)
    Keynote,
    /// OpenDocument Presentation format (.odp)
    Odp,
}

/// Detect the presentation format by reading the file header.
///
/// This function leverages the common detection module for consistent
/// format detection across the library. It reads the minimum necessary
/// bytes to determine the format.
///
/// # Arguments
///
/// * `reader` - A reader that implements Read + Seek
///
/// # Returns
///
/// * `Ok(PresentationFormat)` if a supported presentation format is detected
/// * `Err(Error)` if the format is not recognized or unsupported
#[allow(dead_code)] // For format detection, it is better to use the smart detection function, but the function is still useful for other purposes
pub fn detect_presentation_format<R: Read + Seek>(reader: &mut R) -> Result<PresentationFormat> {
    // Use the common detection module
    let file_format = detection::detect_format_from_reader(reader).ok_or(Error::NotOfficeFile)?;

    // Map FileFormat to PresentationFormat
    map_file_format_to_presentation_format(file_format)
}

/// Detect the presentation format from a byte buffer.
///
/// This is optimized for in-memory detection without seeking.
/// Leverages the common detection module for consistency.
///
/// # Arguments
///
/// * `bytes` - The file data as bytes
///
/// # Returns
///
/// * `Ok(PresentationFormat)` if a supported presentation format is detected
/// * `Err(Error)` if the format is not recognized or unsupported
#[inline]
#[allow(dead_code)] // For format detection, it is better to use the smart detection function, but the function is still useful for other purposes
pub fn detect_presentation_format_from_bytes(bytes: &[u8]) -> Result<PresentationFormat> {
    if bytes.len() < 4 {
        return Err(Error::InvalidFormat(
            "File too small to determine format".to_string(),
        ));
    }

    // Use the common detection module
    let file_format =
        detection::detect_file_format_from_bytes(bytes).ok_or(Error::NotOfficeFile)?;

    // Map FileFormat to PresentationFormat
    map_file_format_to_presentation_format(file_format)
}

/// Map common FileFormat to PresentationFormat.
///
/// This function converts the general FileFormat enum from the common
/// detection module to the presentation-specific PresentationFormat enum.
///
/// # Arguments
///
/// * `file_format` - The detected file format
///
/// # Returns
///
/// * `Ok(PresentationFormat)` if the format is a supported presentation format
/// * `Err(Error::InvalidFormat)` if the format is not a presentation format
#[inline]
fn map_file_format_to_presentation_format(file_format: FileFormat) -> Result<PresentationFormat> {
    match file_format {
        FileFormat::Ppt => Ok(PresentationFormat::Ppt),
        FileFormat::Pptx => Ok(PresentationFormat::Pptx),
        FileFormat::Keynote => Ok(PresentationFormat::Keynote),
        FileFormat::Odp => Ok(PresentationFormat::Odp),
        _ => Err(Error::InvalidFormat(format!(
            "Detected format {:?} is not a presentation format",
            file_format
        ))),
    }
}
