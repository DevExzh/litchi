use std::borrow::Cow;
use std::io::Read;

use super::parts::chp::CharacterProperties;

const BLOCK_TYPE_OFFSET: usize = 0xE;
const MM_MODE_TYPE_OFFSET: usize = 0x6;

/// Compressed image signature: 0xFE 0x78 0xDA (zlib best compression)
const COMPRESSED1: [u8; 3] = [0xFE, 0x78, 0xDA];

/// Compressed image signature: 0xFE 0x78 0x9C (zlib default compression)
const COMPRESSED2: [u8; 3] = [0xFE, 0x78, 0x9C];

/// PNG file signature
const PNG_SIGNATURE: [u8; 8] = [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];

/// Picture type enumeration
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum PictureType {
    #[default]
    Unknown,
    Emf,
    Wmf,
    Pict,
    Jpeg,
    Png,
    Bmp,
    Tiff,
}

impl PictureType {
    /// Get MIME type for this picture type
    pub const fn mime_type(&self) -> &'static str {
        match self {
            PictureType::Unknown => "image/unknown",
            PictureType::Emf => "image/x-emf",
            PictureType::Wmf => "image/x-wmf",
            PictureType::Pict => "image/x-pict",
            PictureType::Jpeg => "image/jpeg",
            PictureType::Png => "image/png",
            PictureType::Bmp => "image/bmp",
            PictureType::Tiff => "image/tiff",
        }
    }

    /// Get file extension for this picture type
    pub const fn extension(&self) -> &'static str {
        match self {
            PictureType::Unknown => "",
            PictureType::Emf => "emf",
            PictureType::Wmf => "wmf",
            PictureType::Pict => "pict",
            PictureType::Jpeg => "jpg",
            PictureType::Png => "png",
            PictureType::Bmp => "bmp",
            PictureType::Tiff => "tiff",
        }
    }

    /// Detect picture type from raw content bytes
    pub fn detect_from_content(data: &[u8]) -> Self {
        if data.len() < 8 {
            return PictureType::Unknown;
        }

        // PNG signature
        if data.starts_with(&PNG_SIGNATURE) {
            return PictureType::Png;
        }

        // JPEG signature
        if data.starts_with(&[0xFF, 0xD8, 0xFF]) {
            return PictureType::Jpeg;
        }

        // BMP signature
        if data.starts_with(b"BM") {
            return PictureType::Bmp;
        }

        // TIFF signature (little-endian and big-endian)
        if data.starts_with(&[0x49, 0x49, 0x2A, 0x00])
            || data.starts_with(&[0x4D, 0x4D, 0x00, 0x2A])
        {
            return PictureType::Tiff;
        }

        // EMF signature
        if data.len() >= 44 && data[40..44] == [0x20, 0x45, 0x4D, 0x46] {
            return PictureType::Emf;
        }

        // WMF signature (Aldus Placeable Metafile or standard)
        if data.len() >= 4
            && ((data[0..2] == [0xD7, 0xCD] && data[2..4] == [0xC6, 0x9A])
                || data[0..4] == [0x01, 0x00, 0x09, 0x00])
        {
            return PictureType::Wmf;
        }

        PictureType::Unknown
    }
}

#[repr(u8)]
#[derive(Debug, Clone, Copy)]
pub enum BlockType {
    Image = 0x08,
    ImageWord2000 = 0x00,
    ImagePastedFromClipboard = 0xA,
    ImagePastedFromClipboardWord2000 = 0x2,
    HorizontalLine = 0xE,
}

#[derive(Debug, Clone, thiserror::Error)]
pub enum ImageError {
    #[error("Invalid picture offset: {0}")]
    InvalidPicOffset(u32),
    #[error("Invalid block type: {0}")]
    InvalidBlockType(u8),
    #[error("No picture found")]
    NoPicture,
    #[error("Decompression failed: {0}")]
    DecompressionFailed(String),
}

impl TryFrom<u8> for BlockType {
    type Error = ImageError;

    #[inline]
    fn try_from(v: u8) -> Result<Self, Self::Error> {
        Ok(match v {
            0x08 => BlockType::Image,
            0x00 => BlockType::ImageWord2000,
            0x0A => BlockType::ImagePastedFromClipboard,
            0x02 => BlockType::ImagePastedFromClipboardWord2000,
            0x0E => BlockType::HorizontalLine,
            _ => return Err(ImageError::InvalidBlockType(v)),
        })
    }
}

fn get_block_type(data_buff: &[u8], pic_offset: u32) -> Result<BlockType, ImageError> {
    let block_type = data_buff
        .get(pic_offset as usize + BLOCK_TYPE_OFFSET)
        .ok_or(ImageError::InvalidPicOffset(pic_offset))?;
    BlockType::try_from(*block_type)
}

fn get_mm_mode_type(data_buff: &[u8], pic_offset: u32) -> Result<u8, ImageError> {
    let mm_mode_type = data_buff
        .get(pic_offset as usize + MM_MODE_TYPE_OFFSET)
        .ok_or(ImageError::InvalidPicOffset(pic_offset))?;
    Ok(*mm_mode_type)
}

fn is_picture_recognized(block_type: BlockType, mm_mode_type: u8) -> bool {
    matches!(
        (block_type, mm_mode_type),
        (BlockType::Image, _)
            | (BlockType::ImagePastedFromClipboard, _)
            | (BlockType::ImageWord2000, 0x64)
            | (BlockType::ImagePastedFromClipboardWord2000, 0x64)
    )
}

fn is_block_contains_image(data_buff: &[u8], pic_offset: u32) -> Result<bool, ImageError> {
    Ok(is_picture_recognized(
        get_block_type(data_buff, pic_offset)?,
        get_mm_mode_type(data_buff, pic_offset)?,
    ))
}

/// Check if a run contains an image based on text and properties.
pub fn has_picture(
    data_buff: &[u8],
    text: &str,
    props: &CharacterProperties,
) -> Result<bool, ImageError> {
    if props.is_spec && !props.is_obj && !props.is_ole2 && !props.is_data {
        // Image should be in its own run, or in a run with the end-of-special marker
        if "\u{0001}" == text || "\u{0001}\u{0015}" == text {
            let pic_offset = props.pic_offset.unwrap_or(0);
            return is_block_contains_image(data_buff, pic_offset);
        }
    }
    Ok(false)
}

/// Check if data matches a signature at a given offset
fn match_signature(data: &[u8], signature: &[u8], offset: usize) -> bool {
    if offset >= data.len() {
        return false;
    }
    let end = (offset + signature.len()).min(data.len());
    data[offset..end] == signature[..end - offset]
}

/// Extract PNG data from raw content, removing any prefix headers
fn extract_png(raw_content: &[u8]) -> Cow<'_, [u8]> {
    if let Some(pos) = raw_content
        .windows(PNG_SIGNATURE.len())
        .position(|window| window == PNG_SIGNATURE)
    {
        if pos == 0 {
            Cow::Borrowed(raw_content)
        } else {
            Cow::Borrowed(&raw_content[pos..])
        }
    } else {
        Cow::Borrowed(raw_content)
    }
}

/// Decompress image content if it's compressed
fn decompress_image_content(raw_content: &[u8]) -> Result<Cow<'_, [u8]>, ImageError> {
    // Check for compression signatures at offset 32
    if match_signature(raw_content, &COMPRESSED1, 32)
        || match_signature(raw_content, &COMPRESSED2, 32)
    {
        if raw_content.len() <= 33 {
            return Err(ImageError::DecompressionFailed(
                "Insufficient data for decompression".into(),
            ));
        }

        let compressed_data = &raw_content[33..];
        let mut decoder = flate2::read::ZlibDecoder::new(compressed_data);
        let mut decompressed = Vec::new();

        match decoder.read_to_end(&mut decompressed) {
            Ok(_) => Ok(Cow::Owned(decompressed)),
            Err(e) => Err(ImageError::DecompressionFailed(format!(
                "Possibly corrupt compression: {}",
                e
            ))),
        }
    } else {
        Ok(extract_png(raw_content))
    }
}

// ============================================================================
// Image struct - metadata only, data loaded lazily
// ============================================================================

/// Embedded image in a Word document.
///
/// This struct only stores metadata (offset). The actual binary data
/// is loaded lazily via `Document::image_data()` to minimize memory usage.
#[derive(Debug, Clone, Copy)]
pub struct Image {
    /// Offset in WordDocument stream where picture data starts
    pic_offset: u32,
}

impl Image {
    /// Create a new Image with the given offset.
    #[inline]
    pub fn new(pic_offset: u32) -> Self {
        Self { pic_offset }
    }

    /// Get the picture offset in the WordDocument stream.
    #[inline]
    pub fn pic_offset(&self) -> u32 {
        self.pic_offset
    }

    /// Get raw image data from the document buffer (zero-copy when possible).
    ///
    /// This method extracts and optionally decompresses the image data.
    /// Use `Document::image_data()` for a higher-level API.
    ///
    /// PICF structure (based on [MS-DOC] and Apache POI):
    /// - Offset 0x00 (4 bytes): lcb - total length of picture data including header
    /// - Offset 0x04 (2 bytes): cbHeader - size of PICF header
    /// - Offset 0x06 (2 bytes): mfpMm - metafile mapping mode
    /// - Offset 0x0E (1 byte): block type (0x08 = image)
    /// - After header: actual picture content (may include BLIP container)
    pub fn data<'a>(&self, data_stream: &'a [u8]) -> Result<Cow<'a, [u8]>, ImageError> {
        let offset = self.pic_offset as usize;

        if offset + 6 >= data_stream.len() {
            return Err(ImageError::InvalidPicOffset(self.pic_offset));
        }

        // Read PICF structure
        // lcb: total length including header (4 bytes at offset 0)
        let lcb = u32::from_le_bytes([
            data_stream[offset],
            data_stream[offset + 1],
            data_stream[offset + 2],
            data_stream[offset + 3],
        ]) as usize;

        // cbHeader: header size (2 bytes at offset 4)
        let cb_header = u16::from_le_bytes([data_stream[offset + 4], data_stream[offset + 5]]) as usize;


        // Validate sizes
        if cb_header < 0x44 {
            // Minimum PICF header size
            return Err(ImageError::InvalidPicOffset(self.pic_offset));
        }

        // Picture content starts after the header
        let content_start = offset + cb_header;
        let content_end = offset + lcb;

        if content_start >= data_stream.len() || content_end > data_stream.len() {
            return Err(ImageError::InvalidPicOffset(self.pic_offset));
        }

        let raw_content = &data_stream[content_start..content_end];

        // Try to decompress if compressed, otherwise extract PNG/JPEG directly
        decompress_image_content(raw_content)
    }

    /// Detect the picture type from the image data.
    pub fn picture_type(&self, word_document: &[u8]) -> Result<PictureType, ImageError> {
        let data = self.data(word_document)?;
        Ok(PictureType::detect_from_content(&data))
    }

    /// Suggest a filename based on offset and detected type.
    pub fn suggest_filename(&self, word_document: &[u8]) -> String {
        let ext = self
            .picture_type(word_document)
            .map(|t| t.extension())
            .unwrap_or("");

        if ext.is_empty() {
            format!("{:x}", self.pic_offset)
        } else {
            format!("{:x}.{}", self.pic_offset, ext)
        }
    }
}

/// Check if a run should have an image and create one if so.
pub fn extract_image(
    data_buff: &[u8],
    text: &str,
    props: &CharacterProperties,
) -> Result<Option<Image>, ImageError> {
    if !has_picture(data_buff, text, props)? {
        return Ok(None);
    }

    props
        .pic_offset
        .map(Image::new)
        .ok_or(ImageError::NoPicture)
        .map(Some)
}
