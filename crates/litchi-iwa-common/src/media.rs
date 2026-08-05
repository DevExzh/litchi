//! Compact media-kind vocabulary shared by concrete iWork owners.
//!
//! This module only classifies filenames and bounded byte prefixes. It does
//! not open archives, inspect protobuf records, or own package transactions.

pub mod playback;

/// The media families recognized by iWork package assets.
#[repr(u8)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Type {
    /// Image file (PNG, JPEG, TIFF, HEIF, and related formats).
    Image,
    /// Video file (MP4, MOV, and related formats).
    Video,
    /// Audio file (MP3, AAC, WAV, and related formats).
    Audio,
    /// PDF document.
    Pdf,
    /// Unknown or unsupported media type.
    Unknown,
}

impl Type {
    /// Classify a media type from a filename extension without allocating.
    #[must_use]
    pub fn from_extension(extension: &str) -> Self {
        if matches_ignore_ascii_case(
            extension,
            &[
                "png", "jpg", "jpeg", "gif", "tiff", "tif", "bmp", "heic", "heif", "webp", "avif",
                "svg",
            ],
        ) {
            Self::Image
        } else if matches_ignore_ascii_case(extension, &["mp4", "mov", "m4v", "avi", "mkv"]) {
            Self::Video
        } else if matches_ignore_ascii_case(
            extension,
            &["mp3", "aac", "m4a", "wav", "aiff", "aif", "ogg"],
        ) {
            Self::Audio
        } else if extension.eq_ignore_ascii_case("pdf") {
            Self::Pdf
        } else {
            Self::Unknown
        }
    }

    /// Classify a media type from a bounded set of common file signatures.
    ///
    /// An unknown ISO-BMFF `ftyp` brand is conservatively treated as video,
    /// matching iWork's existing media replacement policy. Unknown bytes
    /// remain [`Self::Unknown`].
    #[must_use]
    pub fn from_bytes(data: &[u8]) -> Self {
        if data.starts_with(b"\x89PNG\r\n\x1a\n")
            || data.starts_with(b"\xff\xd8\xff")
            || data.starts_with(b"GIF87a")
            || data.starts_with(b"GIF89a")
            || data.starts_with(b"II*\0")
            || data.starts_with(b"MM\0*")
            || data.starts_with(b"BM")
            || (data.get(0..4) == Some(b"RIFF") && data.get(8..12) == Some(b"WEBP"))
        {
            return Self::Image;
        }
        if data.starts_with(b"%PDF-") {
            return Self::Pdf;
        }
        if data.starts_with(b"ID3")
            || data
                .get(..2)
                .is_some_and(|prefix| prefix[0] == 0xff && prefix[1] & 0xe0 == 0xe0)
            || (data.get(0..4) == Some(b"RIFF") && data.get(8..12) == Some(b"WAVE"))
            || (data.get(0..4) == Some(b"FORM")
                && matches!(data.get(8..12), Some(b"AIFF" | b"AIFC")))
            || data.starts_with(b"OggS")
        {
            return Self::Audio;
        }
        if data.len() >= 12 && data.get(4..8) == Some(b"ftyp") {
            return match data.get(8..12) {
                Some(
                    b"heic" | b"heix" | b"hevc" | b"hevx" | b"heim" | b"heis" | b"mif1" | b"msf1"
                    | b"avif" | b"avis",
                ) => Self::Image,
                Some(b"M4A " | b"M4B " | b"M4P ") => Self::Audio,
                _ => Self::Video,
            };
        }
        Self::Unknown
    }

    /// Return the stable human-readable label used in validation diagnostics.
    #[must_use]
    pub const fn name(self) -> &'static str {
        match self {
            Self::Image => "Image",
            Self::Video => "Video",
            Self::Audio => "Audio",
            Self::Pdf => "PDF Document",
            Self::Unknown => "Unknown",
        }
    }
}

fn matches_ignore_ascii_case(value: &str, candidates: &[&str]) -> bool {
    candidates
        .iter()
        .any(|candidate| value.eq_ignore_ascii_case(candidate))
}

#[cfg(test)]
mod tests {
    use std::mem::{align_of, size_of};

    use super::Type;

    #[test]
    fn media_type_is_one_byte_and_copyable() {
        assert_eq!(size_of::<Type>(), 1);
        assert_eq!(align_of::<Type>(), 1);
        let original = Type::Image;
        let copied = original;
        assert_eq!(original, copied);
    }

    #[test]
    fn extensions_are_case_insensitive_and_unknown_is_explicit() {
        assert_eq!(Type::from_extension("PNG"), Type::Image);
        assert_eq!(Type::from_extension("Mov"), Type::Video);
        assert_eq!(Type::from_extension("M4A"), Type::Audio);
        assert_eq!(Type::from_extension("AvIf"), Type::Image);
        assert_eq!(Type::from_extension("PDF"), Type::Pdf);
        assert_eq!(Type::from_extension("bin"), Type::Unknown);
        assert_eq!(Type::Unknown.name(), "Unknown");
    }

    #[test]
    fn signatures_cover_representative_media_families() {
        assert_eq!(Type::from_bytes(b"\x89PNG\r\n\x1a\n"), Type::Image);
        assert_eq!(Type::from_bytes(b"\xff\xd8\xff"), Type::Image);
        assert_eq!(Type::from_bytes(b"GIF89a"), Type::Image);
        assert_eq!(Type::from_bytes(b"RIFFxxxxWEBP"), Type::Image);
        assert_eq!(Type::from_bytes(b"%PDF-1.7"), Type::Pdf);
        assert_eq!(Type::from_bytes(b"ID3\x04\0\0"), Type::Audio);
        assert_eq!(Type::from_bytes(b"RIFFxxxxWAVE"), Type::Audio);
        assert_eq!(Type::from_bytes(b"FORMxxxxAIFF"), Type::Audio);
        assert_eq!(Type::from_bytes(b"OggS"), Type::Audio);
        assert_eq!(Type::from_bytes(b"\0\0\0\x18ftypheic"), Type::Image);
        assert_eq!(Type::from_bytes(b"\0\0\0\x18ftypM4A "), Type::Audio);
        assert_eq!(Type::from_bytes(b"\0\0\0\x18ftypavc1"), Type::Video);
        assert_eq!(Type::from_bytes(b"not media"), Type::Unknown);

        let truncated_ftyp = b"\0\0\0\x18ftypheic";
        for length in 0..12 {
            assert_eq!(
                Type::from_bytes(&truncated_ftyp[..length]),
                Type::Unknown,
                "truncated ftyp header of {length} bytes"
            );
        }
    }
}
