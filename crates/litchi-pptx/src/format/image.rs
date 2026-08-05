//! Image container format detection and serialization metadata.

/// Image formats that PresentationML can embed without transcoding.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ImageFormat {
    Png,
    Jpeg,
    Gif,
    Bmp,
    Tiff,
}

impl ImageFormat {
    /// Return the MIME type stored in the image relationship.
    #[inline]
    pub const fn mime_type(self) -> &'static str {
        match self {
            Self::Png => "image/png",
            Self::Jpeg => "image/jpeg",
            Self::Gif => "image/gif",
            Self::Bmp => "image/bmp",
            Self::Tiff => "image/tiff",
        }
    }

    /// Return the canonical package extension for this format.
    #[inline]
    pub const fn extension(self) -> &'static str {
        match self {
            Self::Png => "png",
            Self::Jpeg => "jpeg",
            Self::Gif => "gif",
            Self::Bmp => "bmp",
            Self::Tiff => "tiff",
        }
    }

    /// Detect a supported image from its bounded magic-number prefix.
    pub fn detect_from_bytes(bytes: &[u8]) -> Option<Self> {
        if bytes.starts_with(&[0x89, 0x50, 0x4e, 0x47]) {
            return Some(Self::Png);
        }
        if bytes.starts_with(&[0xff, 0xd8, 0xff]) {
            return Some(Self::Jpeg);
        }
        if bytes.starts_with(b"GIF8") {
            return Some(Self::Gif);
        }
        if bytes.starts_with(b"BM") {
            return Some(Self::Bmp);
        }
        if bytes.starts_with(&[0x49, 0x49, 0x2a, 0x00])
            || bytes.starts_with(&[0x4d, 0x4d, 0x00, 0x2a])
        {
            return Some(Self::Tiff);
        }
        None
    }
}

#[cfg(test)]
mod tests {
    use super::ImageFormat;

    #[test]
    fn detects_supported_magic_numbers_without_reading_payloads() {
        assert_eq!(
            ImageFormat::detect_from_bytes(b"\x89PNG\r\n"),
            Some(ImageFormat::Png)
        );
        assert_eq!(
            ImageFormat::detect_from_bytes(b"\xff\xd8\xff\xe0"),
            Some(ImageFormat::Jpeg)
        );
        assert_eq!(
            ImageFormat::detect_from_bytes(b"GIF89a"),
            Some(ImageFormat::Gif)
        );
        assert_eq!(
            ImageFormat::detect_from_bytes(b"BM\x00\x00"),
            Some(ImageFormat::Bmp)
        );
        assert_eq!(
            ImageFormat::detect_from_bytes(b"II*\x00"),
            Some(ImageFormat::Tiff)
        );
        assert_eq!(ImageFormat::detect_from_bytes(b"webp"), None);
    }

    #[test]
    fn exposes_stable_relationship_metadata() {
        assert_eq!(ImageFormat::Png.mime_type(), "image/png");
        assert_eq!(ImageFormat::Png.extension(), "png");
    }
}
