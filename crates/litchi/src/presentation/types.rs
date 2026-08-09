//! Internal types for presentation format detection and implementation.

use crate::detection_smart as detection;
use litchi_core::detection::FileFormat;
use litchi_core::{Error, Result};
use std::io::{Read, Seek};

#[cfg(feature = "ppt")]
use crate::ppt;

/// Extracted data from a modern presentation slide (to avoid lifetime issues).
#[derive(Debug, Clone)]
pub struct SlideData {
    pub text: String,
    pub name: Option<String>,
}

/// Extracted data from a legacy presentation slide (to avoid lifetime issues).
#[derive(Debug, Clone)]
pub struct LegacySlideData {
    pub text: String,
    pub slide_number: usize,
    pub shape_count: usize,
}

/// A PowerPoint presentation implementation that can be .ppt, .pptx, .key, or .odp format.
///
/// This enum wraps the format-specific implementations and provides
/// a unified API. Users typically don't interact with this enum directly,
/// but instead use the methods on `Presentation`.
#[allow(
    clippy::large_enum_variant,
    reason = "crate-internal facade enum; boxing the large variant would complicate every match for no measurable gain"
)]
pub(super) enum PresentationImpl {
    /// Legacy .ppt format
    #[cfg(feature = "ppt")]
    Ppt(ppt::Presentation),
    /// Modern .pptx format
    #[cfg(feature = "pptx")]
    Pptx(Box<crate::pptx::Package>),
    /// Apple Keynote format
    #[cfg(feature = "keynote")]
    Keynote(litchi_keynote::Package),
    /// OpenDocument Presentation format
    #[cfg(feature = "odp")]
    Odp(litchi_odp::Presentation),
}

/// Presentation format detection.
///
/// This enum represents the supported presentation formats in the unified
/// Presentation API. The format is automatically detected from file signatures.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[allow(
    unused,
    reason = "the enum is only constructed when the corresponding format feature is enabled"
)]
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
#[allow(
    dead_code,
    reason = "kept as a public detection helper; callers are expected to prefer smart detection"
)]
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
#[allow(
    dead_code,
    reason = "kept as a public detection helper; callers are expected to prefer smart detection"
)]
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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_presentation_format_variants() {
        assert_eq!(PresentationFormat::Ppt, PresentationFormat::Ppt);
        assert_eq!(PresentationFormat::Pptx, PresentationFormat::Pptx);
        assert_eq!(PresentationFormat::Keynote, PresentationFormat::Keynote);
        assert_eq!(PresentationFormat::Odp, PresentationFormat::Odp);
    }

    #[test]
    fn test_presentation_format_inequality() {
        assert_ne!(PresentationFormat::Ppt, PresentationFormat::Pptx);
        assert_ne!(PresentationFormat::Pptx, PresentationFormat::Keynote);
        assert_ne!(PresentationFormat::Odp, PresentationFormat::Ppt);
    }

    #[test]
    fn test_map_file_format_ppt() {
        let result = map_file_format_to_presentation_format(FileFormat::Ppt);
        assert!(result.is_ok());
        assert_eq!(result.unwrap(), PresentationFormat::Ppt);
    }

    #[test]
    fn test_map_file_format_pptx() {
        let result = map_file_format_to_presentation_format(FileFormat::Pptx);
        assert!(result.is_ok());
        assert_eq!(result.unwrap(), PresentationFormat::Pptx);
    }

    #[test]
    fn test_map_file_format_keynote() {
        let result = map_file_format_to_presentation_format(FileFormat::Keynote);
        assert!(result.is_ok());
        assert_eq!(result.unwrap(), PresentationFormat::Keynote);
    }

    #[test]
    fn test_map_file_format_odp() {
        let result = map_file_format_to_presentation_format(FileFormat::Odp);
        assert!(result.is_ok());
        assert_eq!(result.unwrap(), PresentationFormat::Odp);
    }

    #[test]
    fn test_map_file_format_invalid() {
        let result = map_file_format_to_presentation_format(FileFormat::Doc);
        assert!(result.is_err());

        let result = map_file_format_to_presentation_format(FileFormat::Docx);
        assert!(result.is_err());

        let result = map_file_format_to_presentation_format(FileFormat::Xls);
        assert!(result.is_err());
    }

    #[test]
    fn test_presentation_format_debug() {
        let format = PresentationFormat::Pptx;
        let debug_str = format!("{:?}", format);
        assert!(debug_str.contains("Pptx"));
    }

    #[test]
    fn test_presentation_format_clone() {
        let format = PresentationFormat::Ppt;
        let cloned = format;
        assert_eq!(format, cloned);
    }

    #[test]
    fn test_presentation_format_copy() {
        let format = PresentationFormat::Keynote;
        let copied = format;
        assert_eq!(format, copied);
    }
}
