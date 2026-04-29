//! BLIP (Binary Large Image or Picture) support for PPT files
//!
//! This module handles embedding images into PowerPoint presentations.
//! Images are stored as BLIP records within the BStoreContainer.
//!
//! Reference: [MS-ODRAW] Section 2.2.23 - OfficeArtBStoreContainer

use std::io::Write;
use zerocopy::IntoBytes;
use zerocopy_derive::*;

/// Error type for BLIP operations
pub type BlipError = std::io::Error;

// =============================================================================
// BLIP Types (MS-ODRAW 2.2.23)
// =============================================================================

/// BLIP types from MS-ODRAW specification
#[repr(u8)]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BlipType {
    /// Error - no BLIP
    Error = 0x00,
    /// Unknown BLIP type
    Unknown = 0x01,
    /// EMF (Enhanced Metafile)
    Emf = 0x02,
    /// WMF (Windows Metafile)
    Wmf = 0x03,
    /// PICT (Macintosh Picture)
    Pict = 0x04,
    /// JPEG image
    Jpeg = 0x05,
    /// PNG image
    Png = 0x06,
    /// DIB (Device Independent Bitmap)
    Dib = 0x07,
    /// TIFF image
    Tiff = 0x11,
    /// CMYK JPEG
    CmykJpeg = 0x12,
}

impl BlipType {
    /// Get the Escher record type for this BLIP
    pub const fn escher_type(&self) -> u16 {
        match self {
            BlipType::Emf => 0xF01A,
            BlipType::Wmf => 0xF01B,
            BlipType::Pict => 0xF01C,
            BlipType::Jpeg | BlipType::CmykJpeg => 0xF01D,
            BlipType::Png => 0xF01E,
            BlipType::Dib => 0xF01F,
            BlipType::Tiff => 0xF029,
            _ => 0xF018, // msoblipERROR
        }
    }

    /// Get the instance value for BLIP record header
    pub const fn instance(&self) -> u16 {
        match self {
            BlipType::Jpeg => 0x46A,     // JFIF
            BlipType::CmykJpeg => 0x6E2, // CMYK
            BlipType::Png => 0x6E0,
            BlipType::Dib => 0x7A8,
            BlipType::Emf => 0x3D4,
            BlipType::Wmf => 0x216,
            BlipType::Pict => 0x542,
            BlipType::Tiff => 0x6E4,
            _ => 0x000,
        }
    }

    /// Detect BLIP type from image magic bytes
    pub fn detect(data: &[u8]) -> Self {
        if data.len() < 8 {
            return BlipType::Unknown;
        }

        // JPEG: FFD8FF
        if data.starts_with(&[0xFF, 0xD8, 0xFF]) {
            return BlipType::Jpeg;
        }

        // PNG: 89 50 4E 47 0D 0A 1A 0A
        if data.starts_with(&[0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]) {
            return BlipType::Png;
        }

        // BMP/DIB: 42 4D (BM)
        if data.starts_with(&[0x42, 0x4D]) {
            return BlipType::Dib;
        }

        // TIFF: 49 49 2A 00 (little-endian) or 4D 4D 00 2A (big-endian)
        if data.starts_with(&[0x49, 0x49, 0x2A, 0x00])
            || data.starts_with(&[0x4D, 0x4D, 0x00, 0x2A])
        {
            return BlipType::Tiff;
        }

        // EMF: 01 00 00 00 (record type 1)
        if data.len() >= 44 && data[40..44] == [0x20, 0x45, 0x4D, 0x46] {
            return BlipType::Emf;
        }

        // WMF: D7 CD C6 9A (placeable) or 01 00 09 00 (standard)
        if data.starts_with(&[0xD7, 0xCD, 0xC6, 0x9A])
            || data.starts_with(&[0x01, 0x00, 0x09, 0x00])
        {
            return BlipType::Wmf;
        }

        BlipType::Unknown
    }
}

// =============================================================================
// BLIP Record Structures
// =============================================================================

/// UID for BLIP records (16 bytes MD4 hash)
pub type BlipUid = [u8; 16];

/// BLIP store entry (FBSE) - MS-ODRAW 2.2.32
#[derive(Debug, Clone, Copy, FromBytes, IntoBytes, Immutable, KnownLayout)]
#[repr(C, packed)]
pub struct BlipStoreEntry {
    /// BLIP type (win32)
    pub bt_win32: u8,
    /// BLIP type (mac)
    pub bt_mac: u8,
    /// UID (MD4 of BLIP data)
    pub uid: BlipUid,
    /// Tag
    pub tag: u16,
    /// Size of BLIP stream
    pub size: u32,
    /// Reference count
    pub ref_count: u32,
    /// Offset in delay stream (0 if embedded)
    pub delay_offset: u32,
    /// Usage (0=default)
    pub usage: u8,
    /// Length of name (0 for unnamed)
    pub name_len: u8,
    /// Unused bytes
    pub unused: [u8; 2],
}

impl BlipStoreEntry {
    /// Size of FBSE structure
    pub const SIZE: usize = 36;

    /// Create a new BLIP store entry
    pub fn new(blip_type: BlipType, uid: BlipUid, blip_size: u32) -> Self {
        Self {
            bt_win32: blip_type as u8,
            bt_mac: blip_type as u8,
            uid,
            tag: 0x00, // Per POI - tag is typically 0
            size: blip_size,
            ref_count: 1, // Per POI afterInsert increments from 0 to 1
            delay_offset: 0,
            usage: 0,
            name_len: 0,
            unused: [0; 2],
        }
    }
}

// =============================================================================
// Picture Shape Properties
// =============================================================================

/// Escher property IDs for pictures
pub mod prop_id {
    /// Picture BLIP reference
    pub const PIC_BLIP: u16 = 0x4104;
    /// Picture crop from left (in EMUs)
    pub const CROP_LEFT: u16 = 0x0102;
    /// Picture crop from top
    pub const CROP_TOP: u16 = 0x0103;
    /// Picture crop from right
    pub const CROP_RIGHT: u16 = 0x0104;
    /// Picture crop from bottom
    pub const CROP_BOTTOM: u16 = 0x0105;
    /// Picture brightness (-100 to 100)
    pub const BRIGHTNESS: u16 = 0x0107;
    /// Picture contrast (0 to 200)
    pub const CONTRAST: u16 = 0x0108;
    /// Picture gamma
    pub const GAMMA: u16 = 0x0109;
    /// Picture transparency (0-100000)
    pub const TRANSPARENCY: u16 = 0x010B;
    /// Picture flags
    pub const PIC_FLAGS: u16 = 0x017F;
}

// =============================================================================
// Picture Data Container
// =============================================================================

/// Container for picture data and metadata
#[derive(Debug, Clone)]
pub struct PictureData {
    /// Raw image bytes
    pub data: Vec<u8>,
    /// Detected or specified BLIP type
    pub blip_type: BlipType,
    /// Computed UID (MD4 hash)
    pub uid: BlipUid,
    /// Index in BStore container
    pub index: u32,
}

impl PictureData {
    /// Create new picture data from raw bytes
    pub fn new(data: Vec<u8>) -> Self {
        let blip_type = BlipType::detect(&data);
        let uid = compute_uid(&data);
        Self {
            data,
            blip_type,
            uid,
            index: 0,
        }
    }

    /// Create picture data with explicit type
    pub fn with_type(data: Vec<u8>, blip_type: BlipType) -> Self {
        let uid = compute_uid(&data);
        Self {
            data,
            blip_type,
            uid,
            index: 0,
        }
    }

    /// Get the size of this BLIP record
    pub fn blip_size(&self) -> u32 {
        // BLIP header varies by type
        let header_size = match self.blip_type {
            BlipType::Jpeg | BlipType::Png | BlipType::Dib | BlipType::Tiff => 17, // UID(16) + tag(1)
            BlipType::Emf | BlipType::Wmf | BlipType::Pict => 50, // UID(16) + header(34)
            _ => 17,
        };
        header_size + self.data.len() as u32
    }
}

/// Compute MD4-like UID for BLIP data
/// Uses a simple hash for compatibility (real MS-PPT uses MD4)
fn compute_uid(data: &[u8]) -> BlipUid {
    // Simple hash computation (FNV-1a variant spread across 16 bytes)
    let mut uid = [0u8; 16];
    let mut hash: u64 = 0xcbf29ce484222325; // FNV offset basis

    for (i, &byte) in data.iter().enumerate() {
        hash ^= byte as u64;
        hash = hash.wrapping_mul(0x100000001b3); // FNV prime
        uid[i % 16] ^= (hash >> ((i % 8) * 8)) as u8;
    }

    // Add length influence
    let len = data.len() as u64;
    for (i, byte) in uid.iter_mut().enumerate().take(8) {
        *byte ^= ((len >> (i * 8)) & 0xFF) as u8;
    }

    uid
}

// =============================================================================
// BLIP Store Container Builder
// =============================================================================

/// Escher record types
mod escher_type {
    pub const BSTORE_CONTAINER: u16 = 0xF001;
    pub const BSE: u16 = 0xF007;
}

/// Builder for BStoreContainer (picture storage)
#[derive(Debug, Default)]
pub struct BlipStoreBuilder {
    /// Pictures stored in this container
    pictures: Vec<PictureData>,
}

impl BlipStoreBuilder {
    /// Create a new BLIP store builder
    pub fn new() -> Self {
        Self {
            pictures: Vec::new(),
        }
    }

    /// Add a picture and return its 1-based index
    pub fn add_picture(&mut self, data: Vec<u8>) -> u32 {
        let mut pic = PictureData::new(data);
        let index = (self.pictures.len() + 1) as u32;
        pic.index = index;
        self.pictures.push(pic);
        index
    }

    /// Add a picture with explicit type
    pub fn add_picture_with_type(&mut self, data: Vec<u8>, blip_type: BlipType) -> u32 {
        let mut pic = PictureData::with_type(data, blip_type);
        let index = (self.pictures.len() + 1) as u32;
        pic.index = index;
        self.pictures.push(pic);
        index
    }

    /// Check if a picture with the given UID already exists
    pub fn find_by_uid(&self, uid: &BlipUid) -> Option<u32> {
        self.pictures
            .iter()
            .find(|p| &p.uid == uid)
            .map(|p| p.index)
    }

    /// Get the number of pictures
    pub fn count(&self) -> usize {
        self.pictures.len()
    }

    /// Check if empty
    pub fn is_empty(&self) -> bool {
        self.pictures.is_empty()
    }

    /// Build the Pictures stream data (to be written as separate OLE stream)
    /// Per POI: Pictures stream contains BLIP records (8-byte header + raw data)
    pub fn build_pictures_stream(&self) -> Result<Vec<u8>, BlipError> {
        if self.pictures.is_empty() {
            return Ok(Vec::new());
        }

        let mut stream = Vec::new();
        for pic in &self.pictures {
            let blip = self.build_blip_for_stream(pic)?;
            stream.extend_from_slice(&blip);
        }
        Ok(stream)
    }

    /// Build BLIP record for Pictures stream (per POI HSLFPictureData.write())
    fn build_blip_for_stream(&self, pic: &PictureData) -> Result<Vec<u8>, BlipError> {
        let mut record = Vec::new();

        // Signature/options (2 bytes) - (version & 0x0F) | ((instance & 0x0FFF) << 4)
        // For BLIP records, version is 0, so it's just instance << 4
        // Per POI: getSignature() returns values like 0x46A0 for JPEG RGB
        let instance = pic.blip_type.instance();
        let signature = instance << 4; // e.g., 0x46A -> 0x46A0
        record.extend_from_slice(&signature.to_le_bytes());

        // Record type (2 bytes) - nativeId + 0xF018
        let rec_type = pic.blip_type.escher_type();
        record.extend_from_slice(&rec_type.to_le_bytes());

        // Build raw data (UID + tag + image data)
        let raw_data = self.build_blip_raw_data(pic);

        // Size (4 bytes) - length of raw data
        record.extend_from_slice(&(raw_data.len() as u32).to_le_bytes());

        // Raw data
        record.extend_from_slice(&raw_data);

        Ok(record)
    }

    /// Build raw data for BLIP (UID + tag + image)
    fn build_blip_raw_data(&self, pic: &PictureData) -> Vec<u8> {
        let mut data = Vec::new();

        // UID (16 bytes)
        data.extend_from_slice(&pic.uid);

        // Type-specific header
        match pic.blip_type {
            BlipType::Emf | BlipType::Wmf | BlipType::Pict => {
                // Metafile header (34 bytes)
                let uncompressed_size = pic.data.len() as u32;
                data.extend_from_slice(&uncompressed_size.to_le_bytes());
                data.extend_from_slice(&[0u8; 16]); // Bounds rect
                data.extend_from_slice(&[0u8; 8]); // PointSize
                data.extend_from_slice(&uncompressed_size.to_le_bytes());
                data.push(0x00); // Not compressed
                data.push(0xFE); // Filter
            },
            _ => {
                // Tag byte (0 per POI Bitmap.formatImageForSlideshow)
                data.push(0x00);
            },
        }

        // Image data
        data.extend_from_slice(&pic.data);

        data
    }

    /// Build the BStoreContainer record (references Pictures stream)
    /// Per POI: BSE records in PPDrawingGroup reference offsets in Pictures stream
    pub fn build(&self) -> Result<Vec<u8>, BlipError> {
        if self.pictures.is_empty() {
            return Ok(Vec::new());
        }

        let mut container = Vec::new();

        // Calculate offsets in Pictures stream for each picture
        let mut offsets = Vec::new();
        let mut offset = 0u32;
        for pic in &self.pictures {
            offsets.push(offset);
            // Each BLIP in stream: 8-byte header + raw data
            let raw_len = self.build_blip_raw_data(pic).len() as u32;
            offset += 8 + raw_len;
        }

        // Build BSE records with offsets
        let mut bse_records = Vec::new();
        for (i, pic) in self.pictures.iter().enumerate() {
            let bse = self.build_bse_record_with_offset(pic, offsets[i])?;
            bse_records.push(bse);
        }

        // Calculate container content size
        let content_size: u32 = bse_records.iter().map(|r| r.len() as u32).sum();

        // Write BStoreContainer header
        write_escher_header(
            &mut container,
            0x0F,
            self.pictures.len() as u16,
            escher_type::BSTORE_CONTAINER,
            content_size,
        )?;

        // Write all BSE records
        for bse in bse_records {
            container.extend_from_slice(&bse);
        }

        Ok(container)
    }

    /// Build a single BSE (BLIP Store Entry) record referencing Pictures stream
    /// Per POI: BSE contains offset to BLIP in Pictures stream, no embedded BLIP
    fn build_bse_record_with_offset(
        &self,
        pic: &PictureData,
        offset: u32,
    ) -> Result<Vec<u8>, BlipError> {
        let mut record = Vec::new();

        // Calculate BLIP size in Pictures stream (8-byte header + raw data)
        let blip_size = 8 + self.build_blip_raw_data(pic).len() as u32;

        // BSE data with offset to Pictures stream
        let mut bse = BlipStoreEntry::new(pic.blip_type, pic.uid, blip_size);
        bse.delay_offset = offset; // Offset in Pictures stream
        let bse_content_size = BlipStoreEntry::SIZE as u32;

        // BSE header (instance = BLIP type)
        write_escher_header(
            &mut record,
            0x02,
            pic.blip_type as u16,
            escher_type::BSE,
            bse_content_size,
        )?;

        // BSE data (36 bytes, no embedded BLIP - it's in Pictures stream)
        record.extend_from_slice(bse.as_bytes());

        Ok(record)
    }
}

/// Write an Escher record header
fn write_escher_header<W: Write>(
    writer: &mut W,
    version: u8,
    instance: u16,
    rec_type: u16,
    length: u32,
) -> Result<(), BlipError> {
    let ver_inst = (version as u16 & 0x0F) | ((instance & 0x0FFF) << 4);
    writer.write_all(&ver_inst.to_le_bytes())?;
    writer.write_all(&rec_type.to_le_bytes())?;
    writer.write_all(&length.to_le_bytes())?;
    Ok(())
}

// =============================================================================
// Picture Shape Builder
// =============================================================================

/// Builder for picture shapes in slides
pub struct PictureShapeBuilder {
    /// BLIP index (1-based)
    blip_index: u32,
    /// Position and size in EMUs
    x: i32,
    y: i32,
    width: i32,
    height: i32,
    /// Crop values (in EMUs, default 0)
    crop_left: i32,
    crop_top: i32,
    crop_right: i32,
    crop_bottom: i32,
}

impl PictureShapeBuilder {
    /// Create a new picture shape builder
    pub fn new(blip_index: u32, x: i32, y: i32, width: i32, height: i32) -> Self {
        Self {
            blip_index,
            x,
            y,
            width,
            height,
            crop_left: 0,
            crop_top: 0,
            crop_right: 0,
            crop_bottom: 0,
        }
    }

    /// Set crop values
    pub fn with_crop(mut self, left: i32, top: i32, right: i32, bottom: i32) -> Self {
        self.crop_left = left;
        self.crop_top = top;
        self.crop_right = right;
        self.crop_bottom = bottom;
        self
    }

    /// Build the picture's Escher properties
    pub fn build_properties(&self) -> Vec<(u16, u32)> {
        let mut props = Vec::new();

        // BLIP reference (complex property flag set)
        props.push((prop_id::PIC_BLIP, self.blip_index));

        // Crop values if non-zero
        if self.crop_left != 0 {
            props.push((prop_id::CROP_LEFT, self.crop_left as u32));
        }
        if self.crop_top != 0 {
            props.push((prop_id::CROP_TOP, self.crop_top as u32));
        }
        if self.crop_right != 0 {
            props.push((prop_id::CROP_RIGHT, self.crop_right as u32));
        }
        if self.crop_bottom != 0 {
            props.push((prop_id::CROP_BOTTOM, self.crop_bottom as u32));
        }

        // Picture flags (standard defaults)
        props.push((prop_id::PIC_FLAGS, 0x00080000));

        props
    }

    /// Get position
    pub fn position(&self) -> (i32, i32, i32, i32) {
        (self.x, self.y, self.width, self.height)
    }
}

// =============================================================================
// Tests
// =============================================================================

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_blip_type_detection() {
        // JPEG
        let jpeg = [0xFF, 0xD8, 0xFF, 0xE0, 0x00, 0x10, 0x4A, 0x46];
        assert_eq!(BlipType::detect(&jpeg), BlipType::Jpeg);

        // PNG
        let png = [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];
        assert_eq!(BlipType::detect(&png), BlipType::Png);

        // Unknown
        let unknown = [0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07];
        assert_eq!(BlipType::detect(&unknown), BlipType::Unknown);
    }

    #[test]
    fn test_blip_store_builder() {
        let mut builder = BlipStoreBuilder::new();

        // Add a fake PNG
        let png_data = vec![
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, // PNG header
            0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52, // IHDR chunk
        ];
        let index = builder.add_picture(png_data);
        assert_eq!(index, 1);
        assert_eq!(builder.count(), 1);
    }

    #[test]
    fn test_picture_shape_properties() {
        let builder = PictureShapeBuilder::new(1, 100, 200, 300, 400);
        let props = builder.build_properties();

        // Should have BLIP reference and flags
        assert!(props.iter().any(|(id, _)| *id == prop_id::PIC_BLIP));
        assert!(props.iter().any(|(id, _)| *id == prop_id::PIC_FLAGS));
    }
}
