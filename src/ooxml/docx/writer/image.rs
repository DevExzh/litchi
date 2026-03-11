//! Image support for DOCX documents.
use crate::common::unit::{EMUS_PER_INCH, pt_to_emu_f64, px_to_emu_96};
use crate::common::xml::escape_xml;
use crate::ooxml::error::{OoxmlError, Result};
use std::fmt::Write as FmtWrite;

// Import shared format types
pub use super::super::format::ImageFormat;

/// A mutable inline image in a document.
///
/// Inline images are embedded within paragraphs as part of runs.
#[derive(Debug)]
pub struct MutableInlineImage {
    /// Image binary data
    pub(crate) data: Vec<u8>,
    /// Image format
    pub(crate) format: ImageFormat,
    /// Width in EMUs (English Metric Units, 1 inch = 914400 EMUs)
    pub(crate) width_emu: Option<i64>,
    /// Height in EMUs
    pub(crate) height_emu: Option<i64>,
    /// Image description/alt text
    pub(crate) description: String,
}

impl MutableInlineImage {
    /// Create a new inline image from bytes.
    ///
    /// # Arguments
    /// * `data` - Image binary data
    /// * `width_emu` - Optional width in EMUs (English Metric Units)
    /// * `height_emu` - Optional height in EMUs
    pub fn from_bytes(
        data: Vec<u8>,
        width_emu: Option<i64>,
        height_emu: Option<i64>,
    ) -> Result<Self> {
        let format = ImageFormat::detect_from_bytes(&data)
            .ok_or_else(|| OoxmlError::InvalidFormat("Unknown image format".to_string()))?;

        Ok(Self {
            data,
            format,
            width_emu,
            height_emu,
            description: String::new(),
        })
    }

    /// Set the image description/alt text.
    pub fn set_description(&mut self, description: impl Into<String>) -> &mut Self {
        self.description = description.into();
        self
    }

    /// Get a reference to the image data.
    pub fn data(&self) -> &[u8] {
        &self.data
    }

    /// Get the image format.
    pub fn format(&self) -> ImageFormat {
        self.format
    }

    /// Convert image dimensions from pixels to EMUs (assuming 96 DPI).
    pub fn px_to_emu(px: u32) -> i64 {
        px_to_emu_96(px)
    }

    /// Convert image dimensions from points to EMUs.
    pub fn pt_to_emu(pt: f64) -> i64 {
        pt_to_emu_f64(pt)
    }

    /// Serialize the inline image to XML.
    pub(crate) fn to_xml(&self, xml: &mut String, r_id: &str) -> Result<()> {
        let width = self.width_emu.unwrap_or(EMUS_PER_INCH);
        let height = self.height_emu.unwrap_or(EMUS_PER_INCH);
        let desc = escape_xml(&self.description);

        write!(
            xml,
            r#"<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{}" cy="{}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="1" name="Picture" descr="{}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="0" name="Picture" descr="{}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="{}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{}" cy="{}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>"#,
            width, height, desc, desc, r_id, width, height
        )
        .map_err(|e| OoxmlError::Xml(e.to_string()))?;

        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    // Simple valid PNG header
    const PNG_HEADER: &[u8] = &[0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];

    #[test]
    fn test_inline_image_from_bytes_png() {
        let data = PNG_HEADER.to_vec();
        let image = MutableInlineImage::from_bytes(data.clone(), None, None);
        assert!(image.is_ok());
        let img = image.unwrap();
        assert_eq!(img.format, ImageFormat::Png);
        assert_eq!(img.data, data);
        assert!(img.width_emu.is_none());
        assert!(img.height_emu.is_none());
    }

    #[test]
    fn test_inline_image_from_bytes_jpeg() {
        // JPEG magic bytes
        let jpeg_data = vec![0xFF, 0xD8, 0xFF, 0xE0, 0x00, 0x10, 0x4A, 0x46];
        let image = MutableInlineImage::from_bytes(jpeg_data, None, None);
        assert!(image.is_ok());
        let img = image.unwrap();
        assert_eq!(img.format, ImageFormat::Jpeg);
    }

    #[test]
    fn test_inline_image_from_bytes_gif() {
        // GIF magic bytes - GIF89a header
        let gif_data = b"GIF89a\x01\x00\x01\x00".to_vec();
        let image = MutableInlineImage::from_bytes(gif_data, None, None);
        assert!(image.is_ok());
        let img = image.unwrap();
        assert_eq!(img.format, ImageFormat::Gif);
    }

    #[test]
    fn test_inline_image_from_bytes_unknown_format() {
        let invalid_data = vec![0x00, 0x01, 0x02, 0x03];
        let image = MutableInlineImage::from_bytes(invalid_data, None, None);
        assert!(image.is_err());
    }

    #[test]
    fn test_inline_image_with_dimensions() {
        let data = PNG_HEADER.to_vec();
        let width = 9144000i64; // 10 inches in EMUs
        let height = 6858000i64; // 7.5 inches in EMUs
        let image = MutableInlineImage::from_bytes(data, Some(width), Some(height)).unwrap();
        assert_eq!(image.width_emu, Some(width));
        assert_eq!(image.height_emu, Some(height));
    }

    #[test]
    fn test_inline_image_set_description() {
        let data = PNG_HEADER.to_vec();
        let mut image = MutableInlineImage::from_bytes(data, None, None).unwrap();
        image.set_description("Test image description");
        assert_eq!(image.description, "Test image description");
    }

    #[test]
    fn test_inline_image_data_accessor() {
        let data = PNG_HEADER.to_vec();
        let image = MutableInlineImage::from_bytes(data.clone(), None, None).unwrap();
        assert_eq!(image.data(), &data);
    }

    #[test]
    fn test_inline_image_format_accessor() {
        let data = PNG_HEADER.to_vec();
        let image = MutableInlineImage::from_bytes(data, None, None).unwrap();
        assert_eq!(image.format(), ImageFormat::Png);
    }

    #[test]
    fn test_px_to_emu() {
        let emu = MutableInlineImage::px_to_emu(96);
        assert_eq!(emu, 914400); // 1 inch at 96 DPI
    }

    #[test]
    fn test_pt_to_emu() {
        let emu = MutableInlineImage::pt_to_emu(72.0);
        assert_eq!(emu, 914400); // 1 inch = 72 points
    }

    #[test]
    fn test_inline_image_to_xml() {
        let data = PNG_HEADER.to_vec();
        let image = MutableInlineImage::from_bytes(data, Some(9144000), Some(6858000)).unwrap();
        let mut xml = String::new();
        let result = image.to_xml(&mut xml, "rId5");
        assert!(result.is_ok());
        assert!(xml.contains("<w:drawing>"));
        assert!(xml.contains("<wp:inline"));
        assert!(xml.contains("cx=\"9144000\""));
        assert!(xml.contains("cy=\"6858000\""));
        assert!(xml.contains("r:embed=\"rId5\""));
        assert!(xml.contains("</w:drawing>"));
    }

    #[test]
    fn test_inline_image_to_xml_with_description() {
        let data = PNG_HEADER.to_vec();
        let mut image = MutableInlineImage::from_bytes(data, None, None).unwrap();
        image.set_description("My test image");
        let mut xml = String::new();
        let result = image.to_xml(&mut xml, "rId1");
        assert!(result.is_ok());
        assert!(xml.contains("descr=\"My test image\""));
    }

    #[test]
    fn test_inline_image_default_dimensions() {
        let data = PNG_HEADER.to_vec();
        let image = MutableInlineImage::from_bytes(data, None, None).unwrap();
        let mut xml = String::new();
        let _ = image.to_xml(&mut xml, "rId1");
        // Default is EMUS_PER_INCH = 914400
        assert!(xml.contains("cx=\"914400\""));
        assert!(xml.contains("cy=\"914400\""));
    }
}
