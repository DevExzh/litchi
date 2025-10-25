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
use crate::common::error::{Error, Result};
use std::borrow::Cow;
use std::io::Read;

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
pub struct MetafileBlip<'data> {
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
    /// Picture data (may be compressed) - uses Cow for zero-copy when possible
    pub picture_data: Cow<'data, [u8]>,
}

impl<'data> MetafileBlip<'data> {
    /// Parse a metafile BLIP record
    ///
    /// # Arguments
    /// * `data` - Complete record data including header
    ///
    /// # Returns
    /// The parsed metafile BLIP or an error
    pub fn parse(data: &'data [u8]) -> Result<Self> {
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

        // Parse metadata manually (34 bytes total)
        // According to [MS-ODRAW] 2.2.23:
        // - uncompressed_size: 4 bytes
        // - bounds: 16 bytes (4 x i32: left, top, right, bottom)
        // - size_emu: 8 bytes (2 x i32: width, height)
        // - compressed_size: 4 bytes
        // - compression: 1 byte
        // - filter: 1 byte
        if offset + 34 > data.len() {
            return Err(crate::common::error::Error::ParseError(
                "Insufficient data for metafile metadata".into(),
            ));
        }

        // Manual parsing to avoid alignment issues
        let uncompressed_size = read_u32_le(data, offset)?;
        let bounds_left = read_u32_le(data, offset + 4)? as i32;
        let bounds_top = read_u32_le(data, offset + 8)? as i32;
        let bounds_right = read_u32_le(data, offset + 12)? as i32;
        let bounds_bottom = read_u32_le(data, offset + 16)? as i32;
        let size_width = read_u32_le(data, offset + 20)? as i32;
        let size_height = read_u32_le(data, offset + 24)? as i32;
        let compressed_size = read_u32_le(data, offset + 28)?;
        let compression = data[offset + 32];
        let filter = data[offset + 33];
        offset += 34;

        // Extract picture data (zero-copy borrow)
        // Use the remaining data as the picture data, up to compressed_size
        let pic_data_len = compressed_size as usize;
        let available_data = data.len() - offset;

        // If compressed_size is larger than available data, use what we have
        let actual_pic_data_len = pic_data_len.min(available_data);
        let picture_data = Cow::Borrowed(&data[offset..offset + actual_pic_data_len]);

        Ok(Self {
            header,
            uid,
            secondary_uid,
            uncompressed_size,
            bounds: (bounds_left, bounds_top, bounds_right, bounds_bottom),
            size_emu: (size_width, size_height),
            compressed_size,
            compression,
            filter,
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
    pub fn decompress(&self) -> Result<Cow<'data, [u8]>> {
        if !self.is_compressed() {
            return Ok(self.picture_data.clone());
        }

        // Check if data has ZLIB header (0x78 0x9C or similar)
        // MS-ODRAW specifies DEFLATE (RFC1950) which uses ZLIB wrapping
        let use_zlib = self.picture_data.len() >= 2 && self.picture_data[0] == 0x78;

        let mut decompressed = Vec::with_capacity(self.uncompressed_size as usize);

        let result = if use_zlib {
            // Use ZlibDecoder for data with ZLIB wrapper (0x78 0x9C header)
            let mut decoder = flate2::read::ZlibDecoder::new(&self.picture_data[..]);
            decoder.read_to_end(&mut decompressed)
        } else {
            // Use DeflateDecoder for raw DEFLATE data
            let mut decoder = flate2::read::DeflateDecoder::new(&self.picture_data[..]);
            decoder.read_to_end(&mut decompressed)
        };

        match result {
            Ok(_) => Ok(Cow::Owned(decompressed)),
            Err(e) => Err(Error::ParseError(format!("Decompression failed: {}", e))),
        }
    }

    /// Get the BLIP type
    pub fn blip_type(&self) -> Option<BlipType> {
        BlipType::from_record_id(self.header.record_type)
    }

    /// Get WMF data with proper placeable header added
    ///
    /// WMF data in BLIP records doesn't include the placeable header, so we need to
    /// reconstruct it using the bounds and size_emu metadata from the BLIP.
    ///
    /// According to MS-ODRAW and Apache POI:
    /// - BLIP stores WMF without placeable header
    /// - rcBounds contains the logical bounds
    /// - ptSize contains the size in EMU (English Metric Units)
    /// - We need to create a placeable header for proper WMF parsing
    pub fn get_wmf_with_header(&self) -> Result<Cow<'data, [u8]>> {
        // Only for WMF type
        if self.blip_type() != Some(BlipType::Wmf) {
            return Err(Error::ParseError("Not a WMF metafile".into()));
        }

        let wmf_data = self.decompress()?;

        // Check if it already has a placeable header (shouldn't happen with BLIP data)
        if wmf_data.len() >= 4 {
            let first_u32 =
                u32::from_le_bytes([wmf_data[0], wmf_data[1], wmf_data[2], wmf_data[3]]);
            if first_u32 == 0x9AC6CDD7 {
                // Already has placeable header
                return Ok(wmf_data);
            }
        }

        // Calculate proper bounds for placeable header
        // ptSize is in EMU, convert to logical units
        // 1 inch = 914400 EMU, and we use 1440 units per inch (twips)
        let (left, top, right, bottom) =
            if self.bounds.0 == 0 && self.bounds.1 == 0 && self.bounds.2 == 0 && self.bounds.3 == 0
            {
                // Bounds are zero, calculate from size_emu
                // Convert EMU to twips: emu * 1440 / 914400
                let width_twips = (self.size_emu.0 as i64 * 1440 / 914400) as i32;
                let height_twips = (self.size_emu.1 as i64 * 1440 / 914400) as i32;
                (0, 0, width_twips, height_twips)
            } else {
                // Use BLIP bounds - they're already in logical units
                self.bounds
            };

        // Create placeable header (22 bytes)
        let mut result = Vec::with_capacity(22 + wmf_data.len());

        // Key: 0x9AC6CDD7 (Aldus Placeable Metafile magic number)
        result.extend_from_slice(&0x9AC6CDD7u32.to_le_bytes());
        // Handle (always 0)
        result.extend_from_slice(&0u16.to_le_bytes());
        // Left, Top, Right, Bottom (bounds in logical units)
        result.extend_from_slice(&(left as i16).to_le_bytes());
        result.extend_from_slice(&(top as i16).to_le_bytes());
        result.extend_from_slice(&(right as i16).to_le_bytes());
        result.extend_from_slice(&(bottom as i16).to_le_bytes());
        // Inch (units per inch) - use 1440 (twips)
        result.extend_from_slice(&1440u16.to_le_bytes());
        // Reserved (always 0)
        result.extend_from_slice(&0u32.to_le_bytes());

        // Calculate checksum (XOR of all 16-bit words in header so far)
        let mut checksum: u16 = 0;
        for chunk in result[0..20].chunks(2) {
            if chunk.len() == 2 {
                let word = u16::from_le_bytes([chunk[0], chunk[1]]);
                checksum ^= word;
            }
        }
        result.extend_from_slice(&checksum.to_le_bytes());

        // Append original WMF data
        result.extend_from_slice(&wmf_data);

        Ok(Cow::Owned(result))
    }
}

/// Bitmap BLIP data structure (JPEG, PNG, DIB, TIFF)
///
/// These formats have simpler structure without compression metadata
#[derive(Debug, Clone)]
pub struct BitmapBlip<'data> {
    /// Record header
    pub header: RecordHeader,
    /// Primary UID (16 bytes MD4/MD5 hash)
    pub uid: [u8; 16],
    /// Marker byte (0xFF for external files)
    pub marker: u8,
    /// Picture data (already in the target format) - uses Cow for zero-copy when possible
    pub picture_data: Cow<'data, [u8]>,
}

impl<'data> BitmapBlip<'data> {
    /// Parse a bitmap BLIP record
    ///
    /// # Arguments
    /// * `data` - Complete record data including header
    ///
    /// # Returns
    /// The parsed bitmap BLIP or an error
    pub fn parse(data: &'data [u8]) -> Result<Self> {
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

        // Extract picture data (zero-copy borrow)
        let picture_data = Cow::Borrowed(&data[offset..]);

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
pub enum Blip<'data> {
    /// Metafile format (EMF, WMF, PICT)
    Metafile(MetafileBlip<'data>),
    /// Bitmap format (JPEG, PNG, DIB, TIFF)
    Bitmap(BitmapBlip<'data>),
}

impl<'data> Blip<'data> {
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
    /// # Ok::<(), litchi::common::error::Error>(())
    /// ```
    pub fn parse(data: &'data [u8]) -> Result<Self> {
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
    pub fn get_decompressed_data(&self) -> Result<Cow<'data, [u8]>> {
        match self {
            Self::Metafile(m) => m.decompress(),
            Self::Bitmap(b) => Ok(b.picture_data.clone()),
        }
    }

    /// Convert to owned data (useful when lifetime constraints are problematic)
    pub fn into_owned(self) -> Blip<'static> {
        match self {
            Self::Metafile(m) => Blip::Metafile(MetafileBlip {
                header: m.header,
                uid: m.uid,
                secondary_uid: m.secondary_uid,
                uncompressed_size: m.uncompressed_size,
                bounds: m.bounds,
                size_emu: m.size_emu,
                compressed_size: m.compressed_size,
                compression: m.compression,
                filter: m.filter,
                picture_data: Cow::Owned(m.picture_data.into_owned()),
            }),
            Self::Bitmap(b) => Blip::Bitmap(BitmapBlip {
                header: b.header,
                uid: b.uid,
                marker: b.marker,
                picture_data: Cow::Owned(b.picture_data.into_owned()),
            }),
        }
    }

    /// Get decompressed picture data with proper header for WMF
    ///
    /// For WMF metafiles, this adds the placeable header using BLIP metadata.
    /// For other formats, this is equivalent to get_decompressed_data().
    pub fn get_picture_data_for_conversion(&self) -> Result<Cow<'data, [u8]>> {
        match self {
            Blip::Metafile(m) => {
                // For WMF, add placeable header
                if m.blip_type() == Some(BlipType::Wmf) {
                    m.get_wmf_with_header()
                } else {
                    // For EMF and PICT, just decompress
                    m.decompress()
                }
            },
            Blip::Bitmap(_) => {
                // For bitmaps, use regular decompressed data
                self.get_decompressed_data()
            },
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
