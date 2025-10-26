// WMF file parser
//
// Parses Windows Metafile records and extracts relevant information
//
// ## Performance Optimizations
//
// This parser is optimized for minimal memory allocations:
//
// 1. **Zero-copy data storage**: Uses `Bytes` with reference counting instead of `Vec<u8>`
//    - The input data is copied once into a `Bytes` buffer
//    - All record params are zero-copy slices of this buffer via `Bytes::slice()`
//    - Eliminates N allocations where N = number of records
//
// 2. **Pre-allocated records vector**: Estimates capacity based on file size
//    - Reduces reallocation overhead during parsing
//    - Typical WMF files have 20-50 bytes per record on average
//
// 3. **Manual byte parsing**: Avoids zerocopy alignment issues
//    - Direct byte access for little-endian values
//    - No intermediate allocations for header parsing
//
// These optimizations significantly reduce calls to `_platform_memmove`,
// `alloc::raw_vec::RawVec::grow_one`, and `szone_malloc_should_clear`.

use crate::common::error::{Error, Result};
use bytes::Bytes;

/// WMF file type
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum WmfFileType {
    /// Memory metafile
    Memory = 1,
    /// Disk metafile
    Disk = 2,
}

/// WMF header (Placeable Metafile Header)
///
/// This is an optional header that may precede the standard WMF header
#[derive(Debug, Clone)]
pub struct WmfPlaceableHeader {
    /// Key (should be 0x9AC6CDD7)
    pub key: u32,
    /// Left coordinate
    pub left: i16,
    /// Top coordinate
    pub top: i16,
    /// Right coordinate
    pub right: i16,
    /// Bottom coordinate
    pub bottom: i16,
    /// Units per inch
    pub inch: u16,
    /// Checksum
    pub checksum: u16,
}

impl WmfPlaceableHeader {
    const PLACEABLE_KEY: u32 = 0x9AC6CDD7;

    /// Check if data starts with a placeable header
    pub fn is_placeable(data: &[u8]) -> bool {
        if data.len() < 4 {
            return false;
        }
        let key = u32::from_le_bytes([data[0], data[1], data[2], data[3]]);
        key == Self::PLACEABLE_KEY
    }

    /// Parse placeable header from data
    pub fn parse(data: &[u8]) -> Result<Self> {
        if data.len() < 22 {
            return Err(Error::ParseError("WMF placeable header too short".into()));
        }

        // Parse header manually to avoid zerocopy alignment issues
        let key = u32::from_le_bytes([data[0], data[1], data[2], data[3]]);

        if key != Self::PLACEABLE_KEY {
            return Err(Error::ParseError(format!(
                "Invalid WMF placeable key: 0x{:08X}",
                key
            )));
        }

        let left = i16::from_le_bytes([data[6], data[7]]);
        let top = i16::from_le_bytes([data[8], data[9]]);
        let right = i16::from_le_bytes([data[10], data[11]]);
        let bottom = i16::from_le_bytes([data[12], data[13]]);
        let inch = u16::from_le_bytes([data[14], data[15]]);
        let checksum = u16::from_le_bytes([data[20], data[21]]);

        Ok(Self {
            key,
            left,
            top,
            right,
            bottom,
            inch,
            checksum,
        })
    }

    /// Get width
    pub fn width(&self) -> i16 {
        self.right - self.left
    }

    /// Get height
    pub fn height(&self) -> i16 {
        self.bottom - self.top
    }
}

/// WMF standard header
#[derive(Debug, Clone)]
pub struct WmfHeader {
    /// File type (1 = memory, 2 = disk)
    pub file_type: u16,
    /// Header size in words (always 9)
    pub header_size: u16,
    /// Windows version
    pub version: u16,
    /// Size of file in words
    pub file_size: u32,
    /// Number of objects
    pub num_objects: u16,
    /// Size of largest record in words
    pub max_record: u32,
    /// Not used (always 0)
    pub num_params: u16,
}

impl WmfHeader {
    /// Parse WMF standard header
    pub fn parse(data: &[u8]) -> Result<Self> {
        if data.len() < 18 {
            return Err(Error::ParseError("WMF header too short".into()));
        }

        // Parse header manually to avoid zerocopy alignment issues
        let file_type = u16::from_le_bytes([data[0], data[1]]);
        let header_size = u16::from_le_bytes([data[2], data[3]]);
        let version = u16::from_le_bytes([data[4], data[5]]);
        let file_size = u32::from_le_bytes([data[6], data[7], data[8], data[9]]);
        let num_objects = u16::from_le_bytes([data[10], data[11]]);
        let max_record = u32::from_le_bytes([data[12], data[13], data[14], data[15]]);
        let num_params = u16::from_le_bytes([data[16], data[17]]);

        Ok(Self {
            file_type,
            header_size,
            version,
            file_size,
            num_objects,
            max_record,
            num_params,
        })
    }
}

/// WMF record
#[derive(Debug, Clone)]
pub struct WmfRecord {
    /// Record size in words (including size and function)
    pub size: u32,
    /// Record function
    pub function: u16,
    /// Record parameters (zero-copy slice of the original data)
    pub params: Bytes,
}

impl WmfRecord {
    /// Parse a WMF record
    ///
    /// # Arguments
    /// * `data` - Zero-copy bytes buffer containing the WMF data
    /// * `offset` - Offset in the buffer to start parsing
    ///
    /// # Returns
    /// A tuple of (parsed record, bytes consumed)
    pub fn parse(data: &Bytes, offset: usize) -> Result<(Self, usize)> {
        if offset + 6 > data.len() {
            return Err(Error::ParseError("Insufficient data for WMF record".into()));
        }

        // Parse record header manually to avoid zerocopy alignment issues
        let size = u32::from_le_bytes([
            data[offset],
            data[offset + 1],
            data[offset + 2],
            data[offset + 3],
        ]);
        let function = u16::from_le_bytes([data[offset + 4], data[offset + 5]]);

        // Size is in words (16-bit), convert to bytes
        let size_bytes = (size as usize) * 2;

        if size < 3 || offset + size_bytes > data.len() {
            return Err(Error::ParseError(format!(
                "Invalid WMF record size: {} at offset {}",
                size, offset
            )));
        }

        // Parameters start after size and function
        let param_size = size_bytes - 6;
        // Zero-copy slice: this creates a shallow copy with reference counting
        let params = data.slice((offset + 6)..(offset + 6 + param_size));

        Ok((
            Self {
                size,
                function,
                params,
            },
            size_bytes,
        ))
    }

    /// Check if this is an EOF record
    pub const fn is_eof(&self) -> bool {
        self.function == 0x0000
    }
}

/// WMF file parser
#[derive(Debug)]
pub struct WmfParser {
    /// Optional placeable header
    pub placeable: Option<WmfPlaceableHeader>,
    /// Standard WMF header
    pub header: WmfHeader,
    /// All records
    pub records: Vec<WmfRecord>,
    /// Raw WMF data (zero-copy with reference counting)
    data: Bytes,
}

impl WmfParser {
    /// Create a new WMF parser from raw data (borrowed)
    ///
    /// This uses zero-copy techniques with `Bytes` for optimal performance.
    /// The input data is converted to `Bytes` once, and all records share
    /// references to slices of this buffer without additional allocations.
    ///
    /// Note: This method copies the input data. Use [`Self::from_owned`] if you
    /// already own the data to avoid the copy.
    pub fn new(data: &[u8]) -> Result<Self> {
        // Convert to Bytes - requires copying since input is borrowed
        let data = Bytes::copy_from_slice(data);
        Self::parse_internal(data)
    }

    /// Create a new WMF parser from owned data (zero-copy)
    ///
    /// This is more efficient than [`Self::new`] as it takes ownership of the data
    /// without copying.
    ///
    /// # Example
    /// ```ignore
    /// let data = std::fs::read("file.wmf")?;
    /// let parser = WmfParser::from_owned(data)?;
    /// ```
    pub fn from_owned(data: Vec<u8>) -> Result<Self> {
        // Convert Vec to Bytes without copying
        let data = Bytes::from(data);
        Self::parse_internal(data)
    }

    /// Internal parsing implementation shared by both constructors
    fn parse_internal(data: Bytes) -> Result<Self> {
        let mut offset = 0;

        // Check for placeable header
        let placeable = if WmfPlaceableHeader::is_placeable(&data) {
            let header = WmfPlaceableHeader::parse(&data)?;
            offset = 22; // Placeable header is 22 bytes
            Some(header)
        } else {
            None
        };

        // Parse standard header
        if offset + 18 > data.len() {
            return Err(Error::ParseError("WMF data too short for header".into()));
        }

        let header = WmfHeader::parse(&data[offset..])?;
        offset += 18;

        // Pre-allocate records vector with a reasonable fixed capacity.
        // WMF files typically have 20-100 records. Using a moderate initial capacity
        // (128) avoids both massive over-allocation and frequent reallocations.
        // The Vec will grow efficiently if needed (typically doubling capacity).
        let mut records = Vec::with_capacity(128);

        // Parse records - all params will be zero-copy slices of the data buffer
        while offset < data.len() {
            match WmfRecord::parse(&data, offset) {
                Ok((record, consumed)) => {
                    let is_eof = record.is_eof();
                    records.push(record);
                    offset += consumed;

                    if is_eof {
                        break;
                    }
                },
                Err(_) => break,
            }
        }

        Ok(Self {
            placeable,
            header,
            records,
            data,
        })
    }

    /// Get the raw WMF data
    pub fn data(&self) -> &[u8] {
        &self.data
    }

    /// Get width in logical units
    pub fn width(&self) -> i32 {
        if let Some(ref placeable) = self.placeable {
            placeable.width() as i32
        } else {
            // Without placeable header, use a default
            1000
        }
    }

    /// Get height in logical units
    pub fn height(&self) -> i32 {
        if let Some(ref placeable) = self.placeable {
            placeable.height() as i32
        } else {
            // Without placeable header, use a default
            1000
        }
    }

    /// Get aspect ratio
    pub fn aspect_ratio(&self) -> f64 {
        let w = self.width() as f64;
        let h = self.height() as f64;
        if h == 0.0 { 1.0 } else { w / h }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_placeable_key() {
        assert_eq!(WmfPlaceableHeader::PLACEABLE_KEY, 0x9AC6CDD7);
    }
}
