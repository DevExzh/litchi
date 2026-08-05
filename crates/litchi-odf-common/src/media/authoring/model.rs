use crate::package::resolve_package_path;
use litchi_core::{Error, Result};

const MAX_IMAGE_BYTES: usize = 64 * 1024 * 1024;
const PICTURES_DIRECTORY: &str = "Pictures/";
const MAX_PICTURE_INDEX: u32 = 1_000_000;

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

/// A raster image format supported by the common inert authoring path.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Format {
    /// Portable Network Graphics.
    Png,
    /// JPEG.
    Jpeg,
    /// Graphics Interchange Format.
    Gif,
}

impl Format {
    /// File extension used for the package part.
    pub const fn extension(self) -> &'static str {
        match self {
            Self::Png => "png",
            Self::Jpeg => "jpg",
            Self::Gif => "gif",
        }
    }

    /// IANA media type used for the package manifest entry.
    pub const fn media_type(self) -> &'static str {
        match self {
            Self::Png => "image/png",
            Self::Jpeg => "image/jpeg",
            Self::Gif => "image/gif",
        }
    }

    /// Detect the format from magic bytes.
    pub fn sniff(bytes: &[u8]) -> Option<Self> {
        const PNG_MAGIC: &[u8] = b"\x89PNG\r\n\x1a\n";
        const JPEG_MAGIC: &[u8] = b"\xff\xd8\xff";
        const GIF87_MAGIC: &[u8] = b"GIF87a";
        const GIF89_MAGIC: &[u8] = b"GIF89a";
        if bytes.starts_with(PNG_MAGIC) {
            Some(Self::Png)
        } else if bytes.starts_with(JPEG_MAGIC) {
            Some(Self::Jpeg)
        } else if bytes.starts_with(GIF87_MAGIC) || bytes.starts_with(GIF89_MAGIC) {
            Some(Self::Gif)
        } else {
            None
        }
    }
}

/// Validate an image payload without decoding or re-encoding it.
pub fn validate_payload(bytes: &[u8]) -> Result<Format> {
    if bytes.len() > MAX_IMAGE_BYTES {
        return Err(invalid("ODF image payload exceeds the size limit"));
    }
    Format::sniff(bytes)
        .ok_or_else(|| invalid("unsupported image format: PNG, JPEG, and GIF are accepted"))
}

/// Allocate the first free `Pictures/imageN.<extension>` package path.
pub fn allocate_picture_path(
    extension: &str,
    mut is_taken: impl FnMut(&str) -> bool,
) -> Result<String> {
    if extension.is_empty() || !extension.bytes().all(|byte| byte.is_ascii_alphanumeric()) {
        return Err(invalid(
            "ODF picture extension is empty or contains unsafe characters",
        ));
    }
    for index in 1..MAX_PICTURE_INDEX {
        let path = format!("{PICTURES_DIRECTORY}image{index}.{extension}");
        if !is_taken(&path) {
            return Ok(path);
        }
    }
    Err(invalid("ODF picture part namespace is exhausted"))
}

/// One bounded media payload awaiting insertion into a package.
#[derive(Debug, PartialEq, Eq)]
pub struct Part {
    path: String,
    bytes: Vec<u8>,
}

impl Part {
    /// Create a media part at a safe package-local path.
    pub fn new(path: impl Into<String>, bytes: Vec<u8>) -> Result<Self> {
        let path = resolve_package_path(&path.into())?;
        Ok(Self { path, bytes })
    }

    /// The normalized package path.
    pub fn path(&self) -> &str {
        &self.path
    }

    /// Borrow the payload without copying it.
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// Consume the part and return its payload.
    pub fn into_bytes(self) -> Vec<u8> {
        self.bytes
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn sniffs_supported_formats_and_rejects_others() {
        assert_eq!(Format::sniff(b"\x89PNG\r\n\x1a\nrest"), Some(Format::Png));
        assert_eq!(Format::sniff(b"\xff\xd8\xff\xe0rest"), Some(Format::Jpeg));
        assert_eq!(Format::sniff(b"GIF89a-rest"), Some(Format::Gif));
        assert_eq!(Format::sniff(b"BM bitmap"), None);
        assert_eq!(Format::sniff(b""), None);
        assert!(validate_payload(b"\x89PNG\r\n\x1a\nrest").is_ok());
        assert!(validate_payload(b"tiff?").is_err());
    }

    #[test]
    fn allocates_and_validates_picture_paths() {
        let path = allocate_picture_path("png", |_| false).unwrap();
        assert_eq!(path, "Pictures/image1.png");
        let path =
            allocate_picture_path("jpg", |candidate| candidate == "Pictures/image1.jpg").unwrap();
        assert_eq!(path, "Pictures/image2.jpg");
        assert!(allocate_picture_path("png", |_| true).is_err());
        assert!(allocate_picture_path("../png", |_| false).is_err());
    }

    #[test]
    fn parts_normalize_safe_paths_without_copying_on_borrow() {
        let part = Part::new("./Pictures/image1.png", vec![1, 2, 3]).unwrap();
        assert_eq!(part.path(), "Pictures/image1.png");
        assert_eq!(part.bytes(), &[1, 2, 3]);
        assert!(Part::new("../../outside.bin", vec![]).is_err());
    }
}
