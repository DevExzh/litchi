//! RTF picture/image extraction and processing.
//!
//! This module handles extraction of embedded pictures from RTF documents.
//! RTF supports several image formats:
//! - Windows Metafile (WMF)
//! - Enhanced Metafile (EMF)
//! - PNG
//! - JPEG
//! - DIB (Device Independent Bitmap)
//! - BMP

use std::borrow::Cow;

/// Image type in RTF documents.
///
/// Note: This enum is specific to RTF parsing. For general image processing,
/// see the `images` module which has comprehensive format support.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ImageType {
    /// Enhanced Metafile
    Emf,
    /// Windows Metafile
    Wmf,
    /// PNG image
    Png,
    /// JPEG image
    Jpeg,
    /// DIB (Device Independent Bitmap)
    Dib,
    /// Mac PICT format
    Pict,
    /// Unknown or unsupported format
    Unknown,
}

/// Extracted picture from RTF document.
#[derive(Debug, Clone)]
pub struct Picture<'a> {
    /// Image type
    pub image_type: ImageType,
    /// Image data (hex-encoded in RTF, decoded here)
    pub data: Cow<'a, [u8]>,
    /// Picture width (in twips, 1/1440 inch)
    pub width: Option<i32>,
    /// Picture height (in twips)
    pub height: Option<i32>,
    /// Goal width (desired width in twips)
    pub goal_width: Option<i32>,
    /// Goal height (desired height in twips)
    pub goal_height: Option<i32>,
    /// Horizontal scaling percentage
    pub scale_x: Option<i32>,
    /// Vertical scaling percentage
    pub scale_y: Option<i32>,
}

impl<'a> Picture<'a> {
    /// Create a new picture with minimal information.
    #[inline]
    pub fn new(image_type: ImageType, data: Cow<'a, [u8]>) -> Self {
        Self {
            image_type,
            data,
            width: None,
            height: None,
            goal_width: None,
            goal_height: None,
            scale_x: None,
            scale_y: None,
        }
    }

    /// Get the image data as a byte slice.
    #[inline]
    pub fn data(&self) -> &[u8] {
        &self.data
    }

    /// Get the computed width in twips, considering scaling.
    #[inline]
    pub fn computed_width(&self) -> Option<i32> {
        self.goal_width.or(self.width).map(|w| match self.scale_x {
            Some(scale) => (w * scale) / 100,
            None => w,
        })
    }

    /// Get the computed height in twips, considering scaling.
    #[inline]
    pub fn computed_height(&self) -> Option<i32> {
        self.goal_height
            .or(self.height)
            .map(|h| match self.scale_y {
                Some(scale) => (h * scale) / 100,
                None => h,
            })
    }

    /// Convert width from twips to pixels at given DPI.
    ///
    /// # Arguments
    ///
    /// * `dpi` - Dots per inch (typically 96 for screen, 72 for print)
    #[inline]
    pub fn width_pixels(&self, dpi: u32) -> Option<u32> {
        self.computed_width().map(|tw| (tw as u32 * dpi) / 1440)
    }

    /// Convert height from twips to pixels at given DPI.
    ///
    /// # Arguments
    ///
    /// * `dpi` - Dots per inch (typically 96 for screen, 72 for print)
    #[inline]
    pub fn height_pixels(&self, dpi: u32) -> Option<u32> {
        self.computed_height().map(|tw| (tw as u32 * dpi) / 1440)
    }
}

/// Detect image type from binary signature.
///
/// # Arguments
///
/// * `data` - Binary image data
///
/// # Returns
///
/// Detected image type or Unknown
pub fn detect_image_type(data: &[u8]) -> ImageType {
    if data.is_empty() {
        return ImageType::Unknown;
    }

    // Check JPEG signature (starts with FFD8)
    if data.len() >= 2 && data[0] == 0xFF && data[1] == 0xD8 {
        return ImageType::Jpeg;
    }

    // Check PNG signature
    if data.len() >= 8 && data.starts_with(&[0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]) {
        return ImageType::Png;
    }

    // Check EMF signature (0x01 0x00 0x00 0x00)
    if data.len() >= 44 && data[0..4] == [0x01, 0x00, 0x00, 0x00] {
        // Check for EMF marker at offset 40
        if data[40..44] == [0x20, 0x45, 0x4D, 0x46]
        // " EMF"
        {
            return ImageType::Emf;
        }
    }

    // Check WMF signature (0xD7, 0xCD, 0xC6, 0x9A) - Aldus Placeable Metafile
    if data.starts_with(&[0xD7, 0xCD, 0xC6, 0x9A]) {
        return ImageType::Wmf;
    }

    // Check DIB/BMP signature
    if data.starts_with(&[0x42, 0x4D]) {
        // "BM"
        return ImageType::Dib;
    }

    ImageType::Unknown
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_detect_png() {
        let png_sig = vec![0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];
        assert_eq!(detect_image_type(&png_sig), ImageType::Png);
    }

    #[test]
    fn test_detect_jpeg() {
        let jpeg_sig = vec![0xFF, 0xD8, 0xFF, 0xE0];
        assert_eq!(detect_image_type(&jpeg_sig), ImageType::Jpeg);
    }

    #[test]
    fn test_picture_dimensions() {
        let pic = Picture {
            image_type: ImageType::Png,
            data: Cow::Borrowed(&[]),
            width: Some(1440), // 1 inch
            height: Some(1440),
            goal_width: None,
            goal_height: None,
            scale_x: Some(200), // 200% scale
            scale_y: Some(200),
        };

        assert_eq!(pic.computed_width(), Some(2880)); // 2 inches
        assert_eq!(pic.width_pixels(96), Some(192)); // 2 inches at 96 DPI
    }
}
