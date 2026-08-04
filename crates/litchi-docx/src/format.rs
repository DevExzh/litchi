//! Canonical semantic formatting types for WordprocessingML.
//!
//! These types are deliberately independent of package storage and XML tree
//! traversal. Format adapters may use the documented token helpers when they
//! need to bridge the semantic model to WordprocessingML.

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
    /// Returns the WordprocessingML `w:jc/@w:val` token for this alignment.
    pub const fn as_str(&self) -> &'static str {
        match self {
            Self::Left => "left",
            Self::Center => "center",
            Self::Right => "right",
            Self::Justify => "both",
        }
    }
}

/// Underline styles for text.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum UnderlineStyle {
    None,
    Single,
    Words,
    Double,
    Thick,
    Dotted,
    DottedHeavy,
    Dashed,
    DashedHeavy,
    DashLong,
    DashLongHeavy,
    DotDash,
    DashDotHeavy,
    DotDotDash,
    DashDotDotHeavy,
    Wave,
    WavyHeavy,
    WavyDouble,
}

impl UnderlineStyle {
    /// Parses a WordprocessingML `w:u/@w:val` token.
    pub fn from_xml(value: &str) -> Option<Self> {
        match value {
            "none" => Some(Self::None),
            "single" => Some(Self::Single),
            "words" => Some(Self::Words),
            "double" => Some(Self::Double),
            "thick" => Some(Self::Thick),
            "dotted" => Some(Self::Dotted),
            "dottedHeavy" => Some(Self::DottedHeavy),
            "dash" => Some(Self::Dashed),
            "dashedHeavy" => Some(Self::DashedHeavy),
            "dashLong" => Some(Self::DashLong),
            "dashLongHeavy" => Some(Self::DashLongHeavy),
            "dotDash" => Some(Self::DotDash),
            "dashDotHeavy" => Some(Self::DashDotHeavy),
            "dotDotDash" => Some(Self::DotDotDash),
            "dashDotDotHeavy" => Some(Self::DashDotDotHeavy),
            "wave" => Some(Self::Wave),
            "wavyHeavy" => Some(Self::WavyHeavy),
            "wavyDouble" => Some(Self::WavyDouble),
            _ => None,
        }
    }

    /// Returns the WordprocessingML `w:u/@w:val` token for this style.
    pub const fn as_str(&self) -> &'static str {
        match self {
            Self::None => "none",
            Self::Single => "single",
            Self::Words => "words",
            Self::Double => "double",
            Self::Thick => "thick",
            Self::Dotted => "dotted",
            Self::DottedHeavy => "dottedHeavy",
            Self::Dashed => "dash",
            Self::DashedHeavy => "dashedHeavy",
            Self::DashLong => "dashLong",
            Self::DashLongHeavy => "dashLongHeavy",
            Self::DotDash => "dotDash",
            Self::DashDotHeavy => "dashDotHeavy",
            Self::DotDotDash => "dotDotDash",
            Self::DashDotDotHeavy => "dashDotDotHeavy",
            Self::Wave => "wave",
            Self::WavyHeavy => "wavyHeavy",
            Self::WavyDouble => "wavyDouble",
        }
    }
}

#[cfg(test)]
mod tests {
    use super::UnderlineStyle;

    #[test]
    fn underline_style_xml_tokens_round_trip() {
        let styles = [
            UnderlineStyle::None,
            UnderlineStyle::Single,
            UnderlineStyle::Words,
            UnderlineStyle::Double,
            UnderlineStyle::Thick,
            UnderlineStyle::Dotted,
            UnderlineStyle::DottedHeavy,
            UnderlineStyle::Dashed,
            UnderlineStyle::DashedHeavy,
            UnderlineStyle::DashLong,
            UnderlineStyle::DashLongHeavy,
            UnderlineStyle::DotDash,
            UnderlineStyle::DashDotHeavy,
            UnderlineStyle::DotDotDash,
            UnderlineStyle::DashDotDotHeavy,
            UnderlineStyle::Wave,
            UnderlineStyle::WavyHeavy,
            UnderlineStyle::WavyDouble,
        ];
        for style in styles {
            assert_eq!(UnderlineStyle::from_xml(style.as_str()), Some(style));
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
    /// Returns the WordprocessingML `w:val` token for this border style.
    pub const fn as_str(&self) -> &'static str {
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
