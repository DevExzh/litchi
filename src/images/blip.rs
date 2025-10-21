// BLIP (Binary Large Image or Picture) record parsing and handling
//
// This module implements parsing of OfficeArtBlip records from Microsoft Office
// file formats, supporting both metafile formats (EMF, WMF, PICT) and bitmap
// formats (JPEG, PNG, DIB, TIFF).
//
// References:
// - [MS-ODRAW] 2.2.23: OfficeArtBlip records
// - https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-odraw/5dc1b9ed-818c-436f-8a4f-905a7ebb1ba9

use crate::common::binary::{read_u16_le, read_u32_le};
use crate::common::error::Result;
use std::io::Read;
use zerocopy::FromBytes;

/// Type of BLIP record
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BlipType {
    /// Enhanced Metafile (EMF)
    Emf = 0xF01A,
    /// Windows Metafile (WMF)
    Wmf = 0xF01B,
    /// Macintosh PICT
    Pict = 0xF01C,
    /// JPEG
    Jpeg = 0xF01D,
    /// PNG
    Png = 0xF01E,
    /// Device Independent Bitmap (DIB)
    Dib = 0xF01F,
    /// TIFF
    Tiff = 0xF029,
}

impl BlipType {
    /// Parse BlipType from record type ID
    pub fn from_record_id(record_id: u16) -> Option<Self> {
        match record_id {
            0xF01A => Some(Self::Emf),
            0xF01B => Some(Self::Wmf),
            0xF01C => Some(Self::Pict),
            0xF01D => Some(Self::Jpeg),
            0xF01E => Some(Self::Png),
            0xF01F => Some(Self::Dib),
            0xF029 => Some(Self::Tiff),
            _ => None,
        }
    }

    /// Check if this is a metafile format (EMF, WMF, PICT)
    pub const fn is_metafile(&self) -> bool {
        matches!(self, Self::Emf | Self::Wmf | Self::Pict)
    }

    /// Check if this is a bitmap format (JPEG, PNG, DIB, TIFF)
    pub const fn is_bitmap(&self) -> bool {
        !self.is_metafile()
    }

    /// Get the file extension for this BLIP type
    pub const fn extension(&self) -> &'static str {
        match self {
            Self::Emf => "emf",
            Self::Wmf => "wmf",
            Self::Pict => "pict",
            Self::Jpeg => "jpg",
            Self::Png => "png",
            Self::Dib => "dib",
            Self::Tiff => "tiff",
        }
    }
}

/// OfficeArt record header
#[derive(Debug, Clone)]
pub struct RecordHeader {
    /// Record version (4 bits)
    pub version: u8,
    /// Record instance (12 bits)
    pub instance: u16,
    /// Record type
    pub record_type: u16,
    /// Record length (excluding header)
    pub length: u32,
}

impl RecordHeader {
    /// Parse a record header from bytes
    ///
    /// # Arguments
    /// * `data` - Byte slice containing the header (at least 8 bytes)
    ///
    /// # Returns
    /// The parsed header or an error
    pub fn parse(data: &[u8]) -> Result<Self> {
        if data.len() < 8 {
            return Err(crate::common::error::Error::ParseError(
                "Insufficient data for record header".into(),
            ));
        }

        // Read version and instance from first 2 bytes
        let ver_inst = read_u16_le(data, 0).unwrap_or(0);
        let version = (ver_inst & 0x0F) as u8;
        let instance = (ver_inst >> 4) & 0xFFF;

        let record_type = read_u16_le(data, 2).unwrap_or(0);
        let length = read_u32_le(data, 4).unwrap_or(0);

        Ok(Self {
            version,
            instance,
            record_type,
            length,
        })
    }

    /// Get the options field (combination of version and instance)
    pub const fn options(&self) -> u16 {
        (self.instance << 4) | (self.version as u16)
    }
}

/// Metafile BLIP data structure (EMF, WMF, PICT)
///
/// These formats include additional metadata and may be compressed
#[derive(Debug, Clone)]
pub struct MetafileBlip {
    /// Record header
    pub header: RecordHeader,
    /// Primary UID (16 bytes MD4/MD5 hash)
    pub uid: [u8; 16],
    /// Secondary UID (optional, present if (instance ^ signature) == 0x10)
    pub secondary_uid: Option<[u8; 16]>,
    /// Uncompressed size in bytes
    pub uncompressed_size: u32,
    /// Clipping bounds (x1, y1, x2, y2)
    pub bounds: (i32, i32, i32, i32),
    /// Size in EMU (English Metric Units) - width, height
    pub size_emu: (i32, i32),
    /// Compressed size in bytes
    pub compressed_size: u32,
    /// Compression flag (0 = deflate, 0xFE = no compression)
    pub compression: u8,
    /// Filter byte (usually 0xFE)
    pub filter: u8,
    /// Picture data (may be compressed)
    pub picture_data: Vec<u8>,
}

/// Raw metafile metadata for zerocopy parsing (34 bytes)
#[derive(Debug, Clone, FromBytes)]
#[repr(C)]
struct RawMetafileMetadata {
    /// Uncompressed size in bytes
    pub uncompressed_size: u32,
    /// Bounds left
    pub bounds_left: i32,
    /// Bounds top
    pub bounds_top: i32,
    /// Bounds right
    pub bounds_right: i32,
    /// Bounds bottom
    pub bounds_bottom: i32,
    /// Size width in EMU
    pub size_width: i32,
    /// Size height in EMU
    pub size_height: i32,
    /// Compressed size in bytes
    pub compressed_size: u32,
    /// Compression flag
    pub compression: u8,
    /// Filter byte
    pub filter: u8,
}

impl MetafileBlip {
    /// Parse a metafile BLIP record
    ///
    /// # Arguments
    /// * `data` - Complete record data including header
    ///
    /// # Returns
    /// The parsed metafile BLIP or an error
    pub fn parse(data: &[u8]) -> Result<Self> {
        let header = RecordHeader::parse(data)?;
        let mut offset = 8;

        // Parse primary UID
        if offset + 16 > data.len() {
            return Err(crate::common::error::Error::ParseError(
                "Insufficient data for UID".into(),
            ));
        }
        let mut uid = [0u8; 16];
        uid.copy_from_slice(&data[offset..offset + 16]);
        offset += 16;

        // Check if secondary UID is present
        let signature = Self::get_signature(header.record_type);
        let has_secondary = (header.options() ^ signature) == 0x10;
        let secondary_uid = if has_secondary {
            if offset + 16 > data.len() {
                return Err(crate::common::error::Error::ParseError(
                    "Insufficient data for secondary UID".into(),
                ));
            }
            let mut sec_uid = [0u8; 16];
            sec_uid.copy_from_slice(&data[offset..offset + 16]);
            offset += 16;
            Some(sec_uid)
        } else {
            None
        };

        // Parse metadata using zerocopy
        if offset + 34 > data.len() {
            return Err(crate::common::error::Error::ParseError(
                "Insufficient data for metafile metadata".into(),
            ));
        }

        let metadata =
            RawMetafileMetadata::read_from_bytes(&data[offset..offset + 34]).map_err(|_| {
                crate::common::error::Error::ParseError("Invalid metafile metadata format".into())
            })?;
        offset += 34;

        // Extract picture data
        let pic_data_len = metadata.compressed_size as usize;
        if offset + pic_data_len > data.len() {
            return Err(crate::common::error::Error::ParseError(
                "Insufficient data for picture data".into(),
            ));
        }
        let picture_data = data[offset..offset + pic_data_len].to_vec();

        Ok(Self {
            header,
            uid,
            secondary_uid,
            uncompressed_size: metadata.uncompressed_size,
            bounds: (
                metadata.bounds_left,
                metadata.bounds_top,
                metadata.bounds_right,
                metadata.bounds_bottom,
            ),
            size_emu: (metadata.size_width, metadata.size_height),
            compressed_size: metadata.compressed_size,
            compression: metadata.compression,
            filter: metadata.filter,
            picture_data,
        })
    }

    /// Get the signature for a given record type
    fn get_signature(record_type: u16) -> u16 {
        match record_type {
            0xF01A => 0x3D4, // EMF
            0xF01B => 0x216, // WMF
            0xF01C => 0x542, // PICT
            _ => 0,
        }
    }

    /// Check if the picture data is compressed
    pub const fn is_compressed(&self) -> bool {
        self.compression == 0
    }

    /// Decompress the picture data if compressed
    ///
    /// # Returns
    /// Uncompressed picture data or the original data if not compressed
    pub fn decompress(&self) -> Result<Vec<u8>> {
        if !self.is_compressed() {
            return Ok(self.picture_data.clone());
        }

        // Use flate2 to decompress
        let mut decoder = flate2::read::DeflateDecoder::new(&self.picture_data[..]);
        let mut decompressed = Vec::with_capacity(self.uncompressed_size as usize);

        decoder.read_to_end(&mut decompressed).map_err(|e| {
            crate::common::error::Error::ParseError(format!("Decompression failed: {}", e))
        })?;

        Ok(decompressed)
    }

    /// Get the BLIP type
    pub fn blip_type(&self) -> Option<BlipType> {
        BlipType::from_record_id(self.header.record_type)
    }
}

/// Bitmap BLIP data structure (JPEG, PNG, DIB, TIFF)
///
/// These formats have simpler structure without compression metadata
#[derive(Debug, Clone)]
pub struct BitmapBlip {
    /// Record header
    pub header: RecordHeader,
    /// Primary UID (16 bytes MD4/MD5 hash)
    pub uid: [u8; 16],
    /// Marker byte (0xFF for external files)
    pub marker: u8,
    /// Picture data (already in the target format)
    pub picture_data: Vec<u8>,
}

impl BitmapBlip {
    /// Parse a bitmap BLIP record
    ///
    /// # Arguments
    /// * `data` - Complete record data including header
    ///
    /// # Returns
    /// The parsed bitmap BLIP or an error
    pub fn parse(data: &[u8]) -> Result<Self> {
        let header = RecordHeader::parse(data)?;
        let mut offset = 8;

        // Parse UID
        if offset + 16 > data.len() {
            return Err(crate::common::error::Error::ParseError(
                "Insufficient data for UID".into(),
            ));
        }
        let mut uid = [0u8; 16];
        uid.copy_from_slice(&data[offset..offset + 16]);
        offset += 16;

        // Parse marker
        if offset >= data.len() {
            return Err(crate::common::error::Error::ParseError(
                "Insufficient data for marker".into(),
            ));
        }
        let marker = data[offset];
        offset += 1;

        // Extract picture data
        let picture_data = data[offset..].to_vec();

        Ok(Self {
            header,
            uid,
            marker,
            picture_data,
        })
    }

    /// Get the BLIP type
    pub fn blip_type(&self) -> Option<BlipType> {
        BlipType::from_record_id(self.header.record_type)
    }
}

/// General BLIP record that can be either metafile or bitmap
#[derive(Debug, Clone)]
pub enum Blip {
    /// Metafile format (EMF, WMF, PICT)
    Metafile(MetafileBlip),
    /// Bitmap format (JPEG, PNG, DIB, TIFF)
    Bitmap(BitmapBlip),
}

impl Blip {
    /// Parse a BLIP record from bytes
    ///
    /// # Arguments
    /// * `data` - Complete record data including header
    ///
    /// # Returns
    /// The parsed BLIP or an error
    ///
    /// # Example
    /// ```no_run
    /// use litchi::images::blip::Blip;
    ///
    /// let data = vec![/* BLIP record bytes */];
    /// let blip = Blip::parse(&data)?;
    /// ```
    pub fn parse(data: &[u8]) -> Result<Self> {
        if data.len() < 8 {
            return Err(crate::common::error::Error::ParseError(
                "Insufficient data for BLIP record".into(),
            ));
        }

        let header = RecordHeader::parse(data)?;
        let blip_type = BlipType::from_record_id(header.record_type).ok_or_else(|| {
            crate::common::error::Error::ParseError(format!(
                "Unknown BLIP record type: 0x{:04X}",
                header.record_type
            ))
        })?;

        match blip_type {
            BlipType::Emf | BlipType::Wmf | BlipType::Pict => {
                Ok(Self::Metafile(MetafileBlip::parse(data)?))
            },
            BlipType::Jpeg | BlipType::Png | BlipType::Dib | BlipType::Tiff => {
                Ok(Self::Bitmap(BitmapBlip::parse(data)?))
            },
        }
    }

    /// Get the BLIP type
    pub fn blip_type(&self) -> Option<BlipType> {
        match self {
            Self::Metafile(m) => m.blip_type(),
            Self::Bitmap(b) => b.blip_type(),
        }
    }

    /// Get the raw picture data
    ///
    /// For metafiles, this returns the data as-is (possibly compressed).
    /// Use `get_decompressed_data()` to get uncompressed data for metafiles.
    pub fn picture_data(&self) -> &[u8] {
        match self {
            Self::Metafile(m) => &m.picture_data,
            Self::Bitmap(b) => &b.picture_data,
        }
    }

    /// Get decompressed picture data
    ///
    /// For bitmap BLIPs, this returns the data as-is.
    /// For metafile BLIPs, this decompresses if necessary.
    pub fn get_decompressed_data(&self) -> Result<Vec<u8>> {
        match self {
            Self::Metafile(m) => m.decompress(),
            Self::Bitmap(b) => Ok(b.picture_data.clone()),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_blip_type_metafile() {
        assert!(BlipType::Emf.is_metafile());
        assert!(BlipType::Wmf.is_metafile());
        assert!(BlipType::Pict.is_metafile());
        assert!(!BlipType::Jpeg.is_metafile());
    }

    #[test]
    fn test_blip_type_extension() {
        assert_eq!(BlipType::Emf.extension(), "emf");
        assert_eq!(BlipType::Png.extension(), "png");
    }
}
