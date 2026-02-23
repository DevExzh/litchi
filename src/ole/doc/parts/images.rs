/// Images and pictures parser for Word binary format.
///
/// Based on Apache POI's Picture and BLIP implementations, and LibreOffice's sw/source/filter/ww8.
/// Images in DOC files can be:
/// - Inline pictures (in character runs with picf special character 0x01)
/// - Floating pictures (in Data stream with Escher BLIP records)
use super::super::package::{DocError, Result};
use crate::common::binary;

/// Picture format types
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PictureType {
    /// Windows Metafile (WMF)
    Wmf,
    /// Enhanced Metafile (EMF)
    Emf,
    /// JPEG
    Jpeg,
    /// PNG
    Png,
    /// BMP
    Bmp,
    /// GIF
    Gif,
    /// TIFF
    Tiff,
    /// DIB (Device Independent Bitmap)
    Dib,
    /// Unknown/unsupported format
    Unknown,
}

impl PictureType {
    /// Detect picture type from binary data
    pub fn from_data(data: &[u8]) -> Self {
        if data.len() < 8 {
            return PictureType::Unknown;
        }

        // Check magic bytes
        match &data[0..4] {
            // PNG: 89 50 4E 47
            [0x89, 0x50, 0x4E, 0x47] => PictureType::Png,
            // JPEG: FF D8 FF
            [0xFF, 0xD8, 0xFF, _] => PictureType::Jpeg,
            // BMP: 42 4D
            [0x42, 0x4D, _, _] => PictureType::Bmp,
            // GIF: 47 49 46 38
            [0x47, 0x49, 0x46, 0x38] => PictureType::Gif,
            // TIFF: 49 49 or 4D 4D
            [0x49, 0x49, 0x2A, 0x00] | [0x4D, 0x4D, 0x00, 0x2A] => PictureType::Tiff,
            // EMF: 01 00 00 00 (check for EMF signature at offset 40)
            [0x01, 0x00, 0x00, 0x00] if data.len() >= 44 => {
                // EMF has " EMF" at offset 40
                if &data[40..44] == b" EMF" {
                    PictureType::Emf
                } else {
                    PictureType::Unknown
                }
            },
            // WMF: D7 CD C6 9A (placeable) or 01 00 09 00
            [0xD7, 0xCD, 0xC6, 0x9A] => PictureType::Wmf,
            [0x01, 0x00, 0x09, 0x00] => PictureType::Wmf,
            _ => PictureType::Unknown,
        }
    }

    /// Get MIME type for this picture
    pub fn mime_type(&self) -> &'static str {
        match self {
            PictureType::Jpeg => "image/jpeg",
            PictureType::Png => "image/png",
            PictureType::Bmp => "image/bmp",
            PictureType::Gif => "image/gif",
            PictureType::Tiff => "image/tiff",
            PictureType::Wmf => "image/x-wmf",
            PictureType::Emf => "image/x-emf",
            PictureType::Dib => "image/bmp",
            PictureType::Unknown => "application/octet-stream",
        }
    }

    /// Get file extension for this picture
    pub fn extension(&self) -> &'static str {
        match self {
            PictureType::Jpeg => "jpg",
            PictureType::Png => "png",
            PictureType::Bmp => "bmp",
            PictureType::Gif => "gif",
            PictureType::Tiff => "tif",
            PictureType::Wmf => "wmf",
            PictureType::Emf => "emf",
            PictureType::Dib => "dib",
            PictureType::Unknown => "bin",
        }
    }
}

/// Picture descriptor (PICF structure) - minimum 68 bytes
#[derive(Debug, Clone)]
pub struct PictureDescriptor {
    /// Horizontal scaling factor in permille (1000 = 100%)
    pub mx: i16,
    /// Vertical scaling factor in permille
    pub my: i16,
    /// Crop from left in twips
    pub crop_left: i16,
    /// Crop from top in twips
    pub crop_top: i16,
    /// Crop from right in twips
    pub crop_right: i16,
    /// Crop from bottom in twips
    pub crop_bottom: i16,
    /// Picture width in twips (before scaling)
    pub width: i32,
    /// Picture height in twips (before scaling)
    pub height: i32,
}

impl PictureDescriptor {
    /// Parse PICF structure from binary data
    pub fn from_bytes(data: &[u8]) -> Result<Self> {
        if data.len() < 68 {
            return Err(DocError::InvalidFormat("PICF too short".to_string()));
        }

        // PICF structure layout (from MS-DOC):
        // 0x00: lcb (4 bytes) - size of picture data
        // 0x04: cbHeader (2 bytes) - size of PICF structure
        // 0x06: mfp (14 bytes) - metafile header
        // 0x14: bm_rcWinMF (8 bytes) - Windows metafile bounds
        // 0x1C: dxaGoal (2 bytes) - width in twips
        // 0x1E: dyaGoal (2 bytes) - height in twips
        // 0x20: mx (2 bytes) - horizontal scaling
        // 0x22: my (2 bytes) - vertical scaling
        // 0x24: dxaCropLeft (2 bytes)
        // 0x26: dyaCropTop (2 bytes)
        // 0x28: dxaCropRight (2 bytes)
        // 0x2A: dyaCropBottom (2 bytes)

        let width = binary::read_i16_le(data, 0x1C)
            .map_err(|e| DocError::InvalidFormat(format!("Failed to read width: {}", e)))?
            as i32;
        let height = binary::read_i16_le(data, 0x1E)
            .map_err(|e| DocError::InvalidFormat(format!("Failed to read height: {}", e)))?
            as i32;
        let mx = binary::read_i16_le(data, 0x20)
            .map_err(|e| DocError::InvalidFormat(format!("Failed to read mx: {}", e)))?;
        let my = binary::read_i16_le(data, 0x22)
            .map_err(|e| DocError::InvalidFormat(format!("Failed to read my: {}", e)))?;
        let crop_left = binary::read_i16_le(data, 0x24)
            .map_err(|e| DocError::InvalidFormat(format!("Failed to read crop_left: {}", e)))?;
        let crop_top = binary::read_i16_le(data, 0x26)
            .map_err(|e| DocError::InvalidFormat(format!("Failed to read crop_top: {}", e)))?;
        let crop_right = binary::read_i16_le(data, 0x28)
            .map_err(|e| DocError::InvalidFormat(format!("Failed to read crop_right: {}", e)))?;
        let crop_bottom = binary::read_i16_le(data, 0x2A)
            .map_err(|e| DocError::InvalidFormat(format!("Failed to read crop_bottom: {}", e)))?;

        Ok(Self {
            mx,
            my,
            width,
            height,
            crop_left,
            crop_top,
            crop_right,
            crop_bottom,
        })
    }

    /// Get scaled width in twips
    pub fn scaled_width(&self) -> i32 {
        (self.width as i64 * self.mx as i64 / 1000) as i32
    }

    /// Get scaled height in twips
    pub fn scaled_height(&self) -> i32 {
        (self.height as i64 * self.my as i64 / 1000) as i32
    }
}

/// An inline picture in the document
#[derive(Debug, Clone)]
pub struct InlinePicture {
    /// Character position in the document
    pub char_position: u32,
    /// Picture descriptor (PICF)
    pub descriptor: PictureDescriptor,
    /// Picture data (binary image)
    pub data: Vec<u8>,
    /// Picture type (detected from data)
    pub picture_type: PictureType,
}

impl InlinePicture {
    /// Create a new inline picture
    pub fn new(char_position: u32, descriptor: PictureDescriptor, data: Vec<u8>) -> Self {
        let picture_type = PictureType::from_data(&data);
        Self {
            char_position,
            descriptor,
            data,
            picture_type,
        }
    }

    /// Get the picture data
    pub fn data(&self) -> &[u8] {
        &self.data
    }

    /// Get the picture size in bytes
    pub fn size(&self) -> usize {
        self.data.len()
    }
}

/// Picture position type for floating pictures
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PicturePosition {
    /// Inline with text
    Inline,
    /// Float relative to page
    FloatPage,
    /// Float relative to paragraph
    FloatParagraph,
    /// Float relative to column
    FloatColumn,
}

/// A floating picture (shape) in the document
#[derive(Debug, Clone)]
pub struct FloatingPicture {
    /// Shape ID
    pub shape_id: u32,
    /// Position type
    pub position: PicturePosition,
    /// Left position in twips
    pub left: i32,
    /// Top position in twips
    pub top: i32,
    /// Width in twips
    pub width: i32,
    /// Height in twips
    pub height: i32,
    /// Picture data (binary image)
    pub data: Vec<u8>,
    /// Picture type
    pub picture_type: PictureType,
    /// Z-order (layering)
    pub z_order: i32,
}

impl FloatingPicture {
    /// Create a new floating picture
    pub fn new(shape_id: u32, data: Vec<u8>) -> Self {
        let picture_type = PictureType::from_data(&data);
        Self {
            shape_id,
            position: PicturePosition::FloatPage,
            left: 0,
            top: 0,
            width: 0,
            height: 0,
            data,
            picture_type,
            z_order: 0,
        }
    }

    /// Get the picture data
    pub fn data(&self) -> &[u8] {
        &self.data
    }

    /// Get the picture size in bytes
    pub fn size(&self) -> usize {
        self.data.len()
    }
}

/// Pictures table for managing all images in the document
pub struct PicturesTable {
    /// All inline pictures
    inline_pictures: Vec<InlinePicture>,
    /// All floating pictures
    floating_pictures: Vec<FloatingPicture>,
}

impl PicturesTable {
    /// Get all inline pictures
    pub fn inline_pictures(&self) -> &[InlinePicture] {
        &self.inline_pictures
    }

    /// Get all floating pictures
    pub fn floating_pictures(&self) -> &[FloatingPicture] {
        &self.floating_pictures
    }

    /// Get the total count of pictures
    pub fn count(&self) -> usize {
        self.inline_pictures.len() + self.floating_pictures.len()
    }

    /// Create a new empty pictures table
    pub fn new() -> Self {
        Self {
            inline_pictures: Vec::new(),
            floating_pictures: Vec::new(),
        }
    }

    /// Add an inline picture
    pub fn add_inline_picture(&mut self, picture: InlinePicture) {
        self.inline_pictures.push(picture);
    }

    /// Add a floating picture
    pub fn add_floating_picture(&mut self, picture: FloatingPicture) {
        self.floating_pictures.push(picture);
    }

    /// Find inline picture at character position
    pub fn find_inline_at_position(&self, cp: u32) -> Option<&InlinePicture> {
        self.inline_pictures.iter().find(|p| p.char_position == cp)
    }
}

impl Default for PicturesTable {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_picture_type_detection() {
        // PNG signature
        let png_data = [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];
        assert_eq!(PictureType::from_data(&png_data), PictureType::Png);

        // JPEG signature
        let jpeg_data = [0xFF, 0xD8, 0xFF, 0xE0, 0x00, 0x10];
        assert_eq!(PictureType::from_data(&jpeg_data), PictureType::Jpeg);

        // BMP signature
        let bmp_data = [0x42, 0x4D, 0x00, 0x00];
        assert_eq!(PictureType::from_data(&bmp_data), PictureType::Bmp);
    }

    #[test]
    fn test_picture_mime_types() {
        assert_eq!(PictureType::Png.mime_type(), "image/png");
        assert_eq!(PictureType::Jpeg.mime_type(), "image/jpeg");
        assert_eq!(PictureType::Bmp.mime_type(), "image/bmp");
    }

    #[test]
    fn test_picture_descriptor_scaling() {
        let mut desc = PictureDescriptor {
            mx: 1000,
            my: 1000,
            crop_left: 0,
            crop_top: 0,
            crop_right: 0,
            crop_bottom: 0,
            width: 1440, // 1 inch
            height: 1440,
        };

        assert_eq!(desc.scaled_width(), 1440);
        assert_eq!(desc.scaled_height(), 1440);

        // 50% scale
        desc.mx = 500;
        desc.my = 500;
        assert_eq!(desc.scaled_width(), 720);
        assert_eq!(desc.scaled_height(), 720);
    }
}
