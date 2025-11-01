//! Shared formatting types for DOCX (used in both reading and writing).

/// Line spacing options for paragraphs.
#[derive(Debug, Clone, Copy)]
pub enum LineSpacing {
    /// Single line spacing
    Single,
    /// 1.5 line spacing
    OneAndHalf,
    /// Double line spacing
    Double,
    /// Multiple line spacing (e.g., 1.15)
    Multiple(f64),
    /// Exact spacing in points
    Exact(f64),
    /// At least spacing in points
    AtLeast(f64),
}

/// Paragraph alignment options.
#[derive(Debug, Clone, Copy)]
pub enum ParagraphAlignment {
    Left,
    Center,
    Right,
    Justify,
}

impl ParagraphAlignment {
    #[allow(dead_code)]
    pub(crate) fn as_str(&self) -> &'static str {
        match self {
            Self::Left => "left",
            Self::Center => "center",
            Self::Right => "right",
            Self::Justify => "both",
        }
    }
}

/// Underline styles for text.
#[derive(Debug, Clone, Copy)]
pub enum UnderlineStyle {
    Single,
    Double,
    Thick,
    Dotted,
    Dashed,
    DotDash,
    DotDotDash,
    Wave,
}

impl UnderlineStyle {
    pub(crate) fn as_str(&self) -> &'static str {
        match self {
            Self::Single => "single",
            Self::Double => "double",
            Self::Thick => "thick",
            Self::Dotted => "dotted",
            Self::Dashed => "dash",
            Self::DotDash => "dotDash",
            Self::DotDotDash => "dotDotDash",
            Self::Wave => "wave",
        }
    }
}

/// Border styles for table cells.
#[derive(Debug, Clone, Copy)]
pub enum TableBorderStyle {
    None,
    Single,
    Thick,
    Double,
    Dotted,
    Dashed,
    DotDash,
    DotDotDash,
}

impl TableBorderStyle {
    pub(crate) fn as_str(&self) -> &'static str {
        match self {
            Self::None => "none",
            Self::Single => "single",
            Self::Thick => "thick",
            Self::Double => "double",
            Self::Dotted => "dotted",
            Self::Dashed => "dashed",
            Self::DotDash => "dotDash",
            Self::DotDotDash => "dotDotDash",
        }
    }
}

/// Image format detection and properties.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ImageFormat {
    Png,
    Jpeg,
    Gif,
    Bmp,
    Tiff,
    Emf,
    Wmf,
}

impl ImageFormat {
    /// Detect image format from byte signature.
    pub fn detect_from_bytes(data: &[u8]) -> Option<Self> {
        if data.len() < 8 {
            return None;
        }

        // PNG signature
        if data.starts_with(&[0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]) {
            return Some(Self::Png);
        }

        // JPEG signature
        if data.starts_with(&[0xFF, 0xD8, 0xFF]) {
            return Some(Self::Jpeg);
        }

        // GIF signature
        if data.starts_with(b"GIF87a") || data.starts_with(b"GIF89a") {
            return Some(Self::Gif);
        }

        // BMP signature
        if data.starts_with(b"BM") {
            return Some(Self::Bmp);
        }

        // TIFF signature (little-endian and big-endian)
        if data.starts_with(&[0x49, 0x49, 0x2A, 0x00])
            || data.starts_with(&[0x4D, 0x4D, 0x00, 0x2A])
        {
            return Some(Self::Tiff);
        }

        // EMF signature
        if data.len() >= 44 && data[40..44] == [0x20, 0x45, 0x4D, 0x46] {
            return Some(Self::Emf);
        }

        // WMF signature
        if data.len() >= 4
            && ((data[0..2] == [0xD7, 0xCD] && data[2..4] == [0xC6, 0x9A])
                || data[0..4] == [0x01, 0x00, 0x09, 0x00])
        {
            return Some(Self::Wmf);
        }

        None
    }

    /// Get file extension for this format.
    pub fn extension(&self) -> &'static str {
        match self {
            Self::Png => "png",
            Self::Jpeg => "jpeg",
            Self::Gif => "gif",
            Self::Bmp => "bmp",
            Self::Tiff => "tiff",
            Self::Emf => "emf",
            Self::Wmf => "wmf",
        }
    }

    /// Get MIME type for this format.
    pub fn mime_type(&self) -> &'static str {
        match self {
            Self::Png => "image/png",
            Self::Jpeg => "image/jpeg",
            Self::Gif => "image/gif",
            Self::Bmp => "image/bmp",
            Self::Tiff => "image/tiff",
            Self::Emf => "image/x-emf",
            Self::Wmf => "image/x-wmf",
        }
    }
}
