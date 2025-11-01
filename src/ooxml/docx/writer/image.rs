/// Image support for DOCX documents.
use crate::ooxml::error::{OoxmlError, Result};
use std::fmt::Write as FmtWrite;

// Import shared format types
pub use super::super::format::ImageFormat;

/// Escape XML special characters.
fn escape_xml(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

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
        ((px as f64) * 914400.0 / 96.0) as i64
    }

    /// Convert image dimensions from points to EMUs.
    pub fn pt_to_emu(pt: f64) -> i64 {
        (pt * 12700.0) as i64
    }

    /// Serialize the inline image to XML.
    pub(crate) fn to_xml(&self, xml: &mut String, r_id: &str) -> Result<()> {
        let width = self.width_emu.unwrap_or(914400);
        let height = self.height_emu.unwrap_or(914400);
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
