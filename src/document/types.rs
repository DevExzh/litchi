//! Internal types for document format detection and implementation.

use crate::common::detection::{self, FileFormat};
use crate::common::{Error, Result};
use std::io::{Read, Seek};

#[cfg(feature = "ole")]
use crate::ole;

#[cfg(feature = "ooxml")]
use crate::ooxml;

/// A Word document implementation that can be .doc, .docx, .pages, .rtf, or .odt format.
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
    /// RTF format
    #[cfg(feature = "rtf")]
    Rtf(crate::rtf::RtfDocument<'static>),
    /// OpenDocument Text format
    #[cfg(feature = "odf")]
    Odt(crate::odf::Document),
}

/// Document format detection.
///
/// This enum represents the supported document formats in the unified
/// Document API. The format is automatically detected from file signatures.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[allow(unused)] // The corresponding enum would only be used when the feature is enabled
pub(super) enum DocumentFormat {
    /// Legacy .doc format (OLE2)
    Doc,
    /// Modern .docx format (OOXML/ZIP)
    Docx,
    /// Apple Pages format (IWA/ZIP)
    Pages,
    /// RTF format (plain text)
    Rtf,
    /// OpenDocument Text format (.odt)
    Odt,
}

/// Detect the document format by reading the file header.
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
/// * `Ok(DocumentFormat)` if a supported document format is detected
/// * `Err(Error)` if the format is not recognized or unsupported
#[allow(dead_code)] // For format detection, it is better to use the smart detection function, but the function is still useful for other purposes
pub fn detect_document_format<R: Read + Seek>(reader: &mut R) -> Result<DocumentFormat> {
    // Use the common detection module
    let file_format = detection::detect_format_from_reader(reader).ok_or(Error::NotOfficeFile)?;

    // Map FileFormat to DocumentFormat
    map_file_format_to_document_format(file_format)
}

/// Detect the document format from a byte buffer.
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
/// * `Ok(DocumentFormat)` if a supported document format is detected
/// * `Err(Error)` if the format is not recognized or unsupported
#[inline]
#[allow(dead_code)] // For format detection, it is better to use the smart detection function, but the function is still useful for other purposes
pub fn detect_document_format_from_bytes(bytes: &[u8]) -> Result<DocumentFormat> {
    if bytes.len() < 4 {
        return Err(Error::InvalidFormat(
            "File too small to determine format".to_string(),
        ));
    }

    // Use the common detection module
    let file_format =
        detection::detect_file_format_from_bytes(bytes).ok_or(Error::NotOfficeFile)?;

    // Map FileFormat to DocumentFormat
    map_file_format_to_document_format(file_format)
}

/// Map common FileFormat to DocumentFormat.
///
/// This function converts the general FileFormat enum from the common
/// detection module to the document-specific DocumentFormat enum.
///
/// # Arguments
///
/// * `file_format` - The detected file format
///
/// # Returns
///
/// * `Ok(DocumentFormat)` if the format is a supported document format
/// * `Err(Error::InvalidFormat)` if the format is not a document format
#[inline]
fn map_file_format_to_document_format(file_format: FileFormat) -> Result<DocumentFormat> {
    match file_format {
        FileFormat::Doc => Ok(DocumentFormat::Doc),
        FileFormat::Docx => Ok(DocumentFormat::Docx),
        FileFormat::Rtf => Ok(DocumentFormat::Rtf),
        FileFormat::Pages => Ok(DocumentFormat::Pages),
        FileFormat::Odt => Ok(DocumentFormat::Odt),
        _ => Err(Error::InvalidFormat(format!(
            "Detected format {:?} is not a document format",
            file_format
        ))),
    }
}
