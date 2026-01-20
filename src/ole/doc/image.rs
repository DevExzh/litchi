use super::parts::chp::CharacterProperties;

const BLOCK_TYPE_OFFSET: usize = 0xE;
const MM_MODE_TYPE_OFFSET: usize = 0x6;

#[repr(u8)]
#[derive(Debug, Clone, Copy)]
pub enum BlockType {
    Image = 0x08,
    ImageWord2000 = 0x00,
    ImagePastedFromClipboard = 0xA,
    ImagePastedFromClipboardWord2000 = 0x2,
    HorizontalLine = 0xE,
}

#[derive(Debug, thiserror::Error)]
pub enum ImageError {
    #[error("Invalid picture offset: {0}")]
    InvalidPicOffset(u32),
    #[error("Invalid block type: {0}")]
    InvalidBlockType(u8),
    #[error("No picture found")]
    NoPicture,
    #[error("Decompression failed: {0}")]
    DecompressionFailed(String),

    #[error("Failed to decode Escher record: {0}")]
    DecodeEscherRecordFailed(std::io::Error),

    #[error("Failed to extract image from container: {0}")]
    ExtractImageFailed(crate::Error),
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

/// Picture fields structure
///
/// Based on Apache Poi's `fillFields` method in `PICFAbstractType.java`
///
/// The fields ordered as they appear in the PICF structure.
/// Total size: 0x44 (68) bytes
#[derive(Debug, Clone, zerocopy::FromBytes, zerocopy::Immutable, zerocopy::KnownLayout)]
#[repr(C)]
pub struct PictureFields {
    /// Total length of picture data including header (offset 0x00)
    pub lcb: i32,
    /// Size of PICF header (offset 0x04)
    pub cb_header: i16,
    /// Metafile mapping mode (offset 0x06)
    pub mm: i16,
    /// Horizontal extent (offset 0x08)
    pub x_ext: i16,
    /// Vertical extent (offset 0x0A)
    pub y_ext: i16,
    /// HMF swap value (offset 0x0C)
    pub sw_hmf: i16,
    /// GRF flags (offset 0x0E)
    pub grf: i32,
    /// Padding (offset 0x12)
    pub padding: i32,
    /// Presentation manager metafile mapping mode (offset 0x16)
    pub mm_pm: i16,
    /// Padding 2 (offset 0x18)
    pub padding2: i32,
    /// Horizontal goal (offset 0x1C)
    pub dxa_goal: i16,
    /// Vertical goal (offset 0x1E)
    pub dya_goal: i16,
    /// Horizontal scaling factor (offset 0x20)
    pub mx: i16,
    /// Vertical scaling factor (offset 0x22)
    pub my: i16,
    /// Reserved horizontal value 1 (offset 0x24)
    pub dxa_reserved1: i16,
    /// Reserved vertical value 1 (offset 0x26)
    pub dya_reserved1: i16,
    /// Reserved horizontal value 2 (offset 0x28)
    pub dxa_reserved2: i16,
    /// Reserved vertical value 2 (offset 0x2A)
    pub dya_reserved2: i16,
    /// Reserved flag (offset 0x2C)
    pub f_reserved: u8,
    /// Bits per pixel (offset 0x2D)
    pub bpp: u8,
    /// Top border (offset 0x2E)
    pub brc_top80: [u8; 4],
    /// Left border (offset 0x32)
    pub brc_left80: [u8; 4],
    /// Bottom border (offset 0x36)
    pub brc_bottom80: [u8; 4],
    /// Right border (offset 0x3A)
    pub brc_right80: [u8; 4],
    /// Reserved horizontal value 3 (offset 0x3E)
    pub dxa_reserved3: i16,
    /// Reserved vertical value 3 (offset 0x40)
    pub dya_reserved3: i16,
    /// Number of properties (offset 0x42)
    pub c_props: i16,
}

impl PictureFields {
    const SIZE: usize = 0
        + 4
        + 2
        + 2
        + 2
        + 2
        + 2
        + 4
        + 4
        + 2
        + 4
        + 2
        + 2
        + 2
        + 2
        + 2
        + 2
        + 2
        + 2
        + 1
        + 1
        + 4
        + 4
        + 4
        + 4
        + 2
        + 2
        + 2;
    /// Try to parse PictureFields from raw bytes
    ///
    /// # Arguments
    /// * `data` - Raw byte slice
    /// * `offset` - Starting offset within the data
    ///
    /// # Returns
    /// * `Some(PictureFields)` if parsing succeeds
    /// * `None` if data is too short
    pub fn try_parse(data: &[u8], offset: usize) -> Option<Self> {
        use zerocopy::FromBytes;

        let slice = data.get(offset..)?;
        let (fields, _) = Self::read_from_prefix(slice).ok()?;
        Some(fields)
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
    /// - After header: actual picture content (may include BLIP container)
    ///
    /// Logics is copied from Apache Poi's `PICFAndOfficeArtData` method in `PICFAndOfficeArtData.java`
    #[cfg(feature = "imgconv")]
    pub fn data(
        &self,
        data_stream: &[u8],
        word_document: &[u8],
    ) -> Result<crate::images::ExtractedImage<'static>, ImageError> {
        use crate::{images::ImageExtractor, ole::escher::EscherRecord};

        let mut offset = self.pic_offset as usize;

        let pic_fields = PictureFields::try_parse(data_stream, offset)
            .ok_or(ImageError::InvalidPicOffset(self.pic_offset))?;

        offset += PictureFields::SIZE;

        // Handle picture name if mm == 0x66
        if pic_fields.mm == 0x66 {
            let cch_pic_name = u8::from_le_bytes([data_stream[offset]]);
            offset += 1;
            offset += cch_pic_name as usize;
        }

        // Parse the first Escher record (usually SpContainer or BStoreContainer)
        let (_, record_size) = EscherRecord::parse(data_stream, offset)
            .map_err(ImageError::DecodeEscherRecordFailed)?;
        offset += record_size;

        // Continue parsing remaining records looking for BSE or BLIP
        while (offset - self.pic_offset as usize) < pic_fields.lcb as usize {
            let (next_record, next_record_size) = match EscherRecord::parse(data_stream, offset) {
                Ok(r) => r,
                Err(_) => break,
            };

            // Check if this is a BSE (0xF007) or BLIP record (0xF018-0xF117)
            if next_record.record_type_raw != 0xF007
                && (next_record.record_type_raw < 0xF018 || next_record.record_type_raw > 0xF117)
            {
                break;
            }
            offset += next_record_size;

            // Try to extract image from this record
            // Pass data_stream for delay-loaded BLIPs
            match ImageExtractor::extract_from_escher_record_with_stream(
                &next_record,
                Some(word_document),
            ) {
                Ok(img) => return Ok(img),
                Err(_) => {
                    // TODO: log this error?
                },
            }
        }

        Err(ImageError::NoPicture)
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
