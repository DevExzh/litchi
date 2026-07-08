//! Images writer for DOC files
//!
//! Generates picture descriptors (PICF) and embedded image data.

use std::io::Write;

/// Picture type
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PictureType {
    Jpeg = 0x46A,
    Png = 0x46B,
    Gif = 0x46C,
    Tiff = 0x46D,
    Bmp = 0x46E,
    Emf = 0x3D4,
    Wmf = 0x3D5,
}

impl PictureType {
    /// Get MIME type
    pub fn mime_type(&self) -> &'static str {
        match self {
            PictureType::Jpeg => "image/jpeg",
            PictureType::Png => "image/png",
            PictureType::Gif => "image/gif",
            PictureType::Tiff => "image/tiff",
            PictureType::Bmp => "image/bmp",
            PictureType::Emf => "image/x-emf",
            PictureType::Wmf => "image/x-wmf",
        }
    }

    /// Detect from data
    pub fn from_data(data: &[u8]) -> Option<Self> {
        if data.len() < 4 {
            return None;
        }

        // JPEG
        if data.starts_with(&[0xFF, 0xD8, 0xFF]) {
            return Some(PictureType::Jpeg);
        }
        // PNG
        if data.starts_with(&[0x89, 0x50, 0x4E, 0x47]) {
            return Some(PictureType::Png);
        }
        // GIF
        if data.starts_with(b"GIF87a") || data.starts_with(b"GIF89a") {
            return Some(PictureType::Gif);
        }
        // BMP
        if data.starts_with(b"BM") {
            return Some(PictureType::Bmp);
        }
        // EMF
        if data.len() >= 44 && data[40..44] == [0x20, 0x45, 0x4D, 0x46] {
            return Some(PictureType::Emf);
        }
        // WMF
        if data.len() >= 4 && (data[0..2] == [0xD7, 0xCD] || data[0..2] == [0x01, 0x00]) {
            return Some(PictureType::Wmf);
        }

        None
    }
}

/// Picture entry
#[derive(Debug, Clone)]
pub struct PictureEntry {
    /// Character position in document
    pub position: u32,
    /// Image data
    pub data: Vec<u8>,
    /// Picture type
    pub picture_type: PictureType,
    /// Width in twips (1/1440 inch)
    pub width: i32,
    /// Height in twips
    pub height: i32,
    /// Scaled width (for display)
    pub scaled_width: i32,
    /// Scaled height (for display)
    pub scaled_height: i32,
}

impl PictureEntry {
    /// Create a new picture entry
    pub fn new(position: u32, data: Vec<u8>, width: i32, height: i32) -> Self {
        let picture_type = PictureType::from_data(&data).unwrap_or(PictureType::Bmp);
        Self {
            position,
            data,
            picture_type,
            width,
            height,
            scaled_width: width,
            scaled_height: height,
        }
    }

    /// Create from file data with auto-detection
    pub fn from_data(position: u32, data: Vec<u8>) -> Self {
        // Default size: 2 inches = 2880 twips
        Self::new(position, data, 2880, 2880)
    }

    /// Set dimensions in twips
    pub fn with_dimensions(mut self, width: i32, height: i32) -> Self {
        self.width = width;
        self.height = height;
        self.scaled_width = width;
        self.scaled_height = height;
        self
    }

    /// Generate PICF (Picture Descriptor) structure
    ///
    /// This is a simplified version of the PICF structure
    pub fn to_picf(&self) -> Vec<u8> {
        let mut buf = Vec::new();

        // lcb: total size of PICF structure (4 bytes)
        let total_size = 68 + self.data.len() as u32;
        buf.write_all(&total_size.to_le_bytes()).unwrap();

        // cbHeader: size of header (2 bytes) - 68 bytes for PICF
        buf.write_all(&68u16.to_le_bytes()).unwrap();

        // mm: mapping mode (2 bytes) - MM_SHAPE (99)
        buf.write_all(&99u16.to_le_bytes()).unwrap();

        // xExt: width in twips (2 bytes)
        buf.write_all(&(self.width as i16).to_le_bytes()).unwrap();

        // yExt: height in twips (2 bytes)
        buf.write_all(&(self.height as i16).to_le_bytes()).unwrap();

        // swHMF: reserved (2 bytes)
        buf.write_all(&0u16.to_le_bytes()).unwrap();

        // Bounds rectangle (8 bytes)
        buf.write_all(&[0; 8]).unwrap();

        // dxaGoal: original width in twips (4 bytes)
        buf.write_all(&self.width.to_le_bytes()).unwrap();

        // dyaGoal: original height in twips (4 bytes)
        buf.write_all(&self.height.to_le_bytes()).unwrap();

        // mx: horizontal scaling (2 bytes) - 1000 = 100%
        buf.write_all(&1000u16.to_le_bytes()).unwrap();

        // my: vertical scaling (2 bytes) - 1000 = 100%
        buf.write_all(&1000u16.to_le_bytes()).unwrap();

        // dxaCropLeft, dyaCropTop, dxaCropRight, dyaCropBottom (8 bytes)
        buf.write_all(&[0; 8]).unwrap();

        // brcl: border (2 bytes)
        buf.write_all(&0u16.to_le_bytes()).unwrap();

        // brcTop, brcLeft, brcBottom, brcRight (16 bytes)
        buf.write_all(&[0; 16]).unwrap();

        // dxaOrigin, dyaOrigin (4 bytes)
        buf.write_all(&[0; 4]).unwrap();

        // cProps (2 bytes)
        buf.write_all(&0u16.to_le_bytes()).unwrap();

        // Image data
        buf.extend_from_slice(&self.data);

        buf
    }
}

/// Images writer
#[derive(Debug)]
pub struct ImagesWriter {
    pictures: Vec<PictureEntry>,
}

impl ImagesWriter {
    /// Create a new images writer
    pub fn new() -> Self {
        Self {
            pictures: Vec::new(),
        }
    }

    /// Add a picture
    pub fn add_picture(&mut self, picture: PictureEntry) {
        self.pictures.push(picture);
    }

    /// Get all pictures
    pub fn pictures(&self) -> &[PictureEntry] {
        &self.pictures
    }

    /// Check if empty
    pub fn is_empty(&self) -> bool {
        self.pictures.is_empty()
    }

    /// Get pictures sorted by position
    pub fn pictures_sorted(&self) -> Vec<&PictureEntry> {
        let mut sorted: Vec<_> = self.pictures.iter().collect();
        sorted.sort_by_key(|p| p.position);
        sorted
    }
}

impl Default for ImagesWriter {
    fn default() -> Self {
        Self::new()
    }
}
