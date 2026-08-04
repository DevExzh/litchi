//! Format types for PPTX presentations.

/// Image format types supported by PPTX.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ImageFormat {
    Png,
    Jpeg,
    Gif,
    Bmp,
    Tiff,
}

impl ImageFormat {
    /// Get the MIME type for this image format.
    pub fn mime_type(&self) -> &'static str {
        match self {
            Self::Png => "image/png",
            Self::Jpeg => "image/jpeg",
            Self::Gif => "image/gif",
            Self::Bmp => "image/bmp",
            Self::Tiff => "image/tiff",
        }
    }

    /// Get the file extension for this image format.
    pub fn extension(&self) -> &'static str {
        match self {
            Self::Png => "png",
            Self::Jpeg => "jpeg",
            Self::Gif => "gif",
            Self::Bmp => "bmp",
            Self::Tiff => "tiff",
        }
    }

    /// Detect image format from bytes (magic number detection).
    pub fn detect_from_bytes(bytes: &[u8]) -> Option<Self> {
        if bytes.len() < 4 {
            return None;
        }

        // PNG: 89 50 4E 47
        if bytes.starts_with(&[0x89, 0x50, 0x4E, 0x47]) {
            return Some(Self::Png);
        }

        // JPEG: FF D8 FF
        if bytes.starts_with(&[0xFF, 0xD8, 0xFF]) {
            return Some(Self::Jpeg);
        }

        // GIF: 47 49 46 38 (GIF8)
        if bytes.starts_with(&[0x47, 0x49, 0x46, 0x38]) {
            return Some(Self::Gif);
        }

        // BMP: 42 4D (BM)
        if bytes.starts_with(&[0x42, 0x4D]) {
            return Some(Self::Bmp);
        }

        // TIFF: 49 49 2A 00 (little-endian) or 4D 4D 00 2A (big-endian)
        if bytes.starts_with(&[0x49, 0x49, 0x2A, 0x00])
            || bytes.starts_with(&[0x4D, 0x4D, 0x00, 0x2A])
        {
            return Some(Self::Tiff);
        }

        None
    }
}

/// Text formatting properties for shapes.
#[derive(Debug, Clone, Default)]
pub struct TextFormat {
    /// Font family
    pub font: Option<String>,
    /// Font size in points
    pub size: Option<f64>,
    /// Bold text
    pub bold: Option<bool>,
    /// Italic text
    pub italic: Option<bool>,
    /// Underline text
    pub underline: Option<bool>,
    /// Text color in hex RGB (e.g., "FF0000" for red)
    pub color: Option<String>,
}
