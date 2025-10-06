// PICT file parser
//
// Parses Macintosh PICT format records and extracts relevant information

use crate::common::error::{Error, Result};

/// PICT file version
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PictVersion {
    /// Version 1 (original format)
    V1,
    /// Version 2 (extended format)
    V2,
}

/// PICT file header
///
/// PICT files may have an optional 512-byte header (used by some applications)
/// followed by the actual PICT data.
#[derive(Debug, Clone)]
pub struct PictHeader {
    /// Version of the PICT file
    pub version: PictVersion,
    /// Picture frame (top, left, bottom, right)
    pub frame: (i16, i16, i16, i16),
    /// Whether this file has the 512-byte header
    pub has_512_header: bool,
}

impl PictHeader {
    /// Parse PICT header from data
    pub fn parse(data: &[u8]) -> Result<Self> {
        if data.len() < 10 {
            return Err(Error::ParseError("PICT data too short".into()));
        }

        let mut offset = 0;

        // Check for optional 512-byte header
        // This header is often all zeros or contains application-specific data
        let has_512_header = data.len() > 512 && Self::check_512_header(data);
        if has_512_header {
            offset = 512;
        }

        // PICT data starts with:
        // - 10 bytes of size/frame info for version 1
        // - picSize (2 bytes) + picFrame (8 bytes) for both versions
        
        if offset + 10 > data.len() {
            return Err(Error::ParseError("Insufficient data for PICT header".into()));
        }

        // Parse picture size (used in version 1, may be 0 in version 2)
        let _pic_size = u16::from_be_bytes([data[offset], data[offset + 1]]);
        offset += 2;

        // Parse picture frame (top, left, bottom, right) - big-endian
        let top = i16::from_be_bytes([data[offset], data[offset + 1]]);
        let left = i16::from_be_bytes([data[offset + 2], data[offset + 3]]);
        let bottom = i16::from_be_bytes([data[offset + 4], data[offset + 5]]);
        let right = i16::from_be_bytes([data[offset + 6], data[offset + 7]]);
        offset += 8;

        // Determine version
        // Version 2 files have a version opcode (0x0011) followed by 0x02FF
        let version = if offset + 4 <= data.len() {
            let op1 = u16::from_be_bytes([data[offset], data[offset + 1]]);
            let op2 = u16::from_be_bytes([data[offset + 2], data[offset + 3]]);
            
            if op1 == 0x0011 && op2 == 0x02FF {
                PictVersion::V2
            } else {
                PictVersion::V1
            }
        } else {
            PictVersion::V1
        };

        Ok(Self {
            version,
            frame: (top, left, bottom, right),
            has_512_header,
        })
    }

    /// Check if data starts with 512-byte header
    fn check_512_header(data: &[u8]) -> bool {
        if data.len() < 522 {
            return false;
        }

        // The 512-byte header is application-specific
        // After it, we should see valid PICT data
        // Check if byte 512-513 looks like a reasonable picture size
        // or if bytes 522-523 look like a version opcode (0x0011)
        let potential_size = u16::from_be_bytes([data[512], data[513]]);
        let potential_opcode = u16::from_be_bytes([data[522], data[523]]);

        // If we see version opcode at expected position, likely has 512 header
        potential_opcode == 0x0011 || potential_size < 0x4000
    }

    /// Get width of the picture
    pub fn width(&self) -> i16 {
        self.frame.3 - self.frame.1
    }

    /// Get height of the picture
    pub fn height(&self) -> i16 {
        self.frame.2 - self.frame.0
    }
}

/// PICT opcode
///
/// PICT files are composed of opcodes that describe drawing operations
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PictOpcode {
    /// No operation
    Nop = 0x0000,
    /// Clipping region
    Clip = 0x0001,
    /// Background pattern
    BkPat = 0x0002,
    /// Text font
    TxFont = 0x0003,
    /// Text face
    TxFace = 0x0004,
    /// Text mode
    TxMode = 0x0005,
    /// Version opcode
    Version = 0x0011,
    /// Extended version 2 header
    HeaderOp = 0x0C00,
    /// End of picture
    EndPic = 0x00FF,
    /// Direct bits rect (includes bitmap data)
    DirectBitsRect = 0x009A,
    /// Packed direct bits rect
    PackedDirectBitsRect = 0x009B,
    /// Compressed QuickTime image
    CompressedQuickTime = 0x8200,
}

impl PictOpcode {
    /// Create from u16 value
    pub fn from_u16(value: u16) -> Option<Self> {
        match value {
            0x0000 => Some(Self::Nop),
            0x0001 => Some(Self::Clip),
            0x0002 => Some(Self::BkPat),
            0x0003 => Some(Self::TxFont),
            0x0004 => Some(Self::TxFace),
            0x0005 => Some(Self::TxMode),
            0x0011 => Some(Self::Version),
            0x0C00 => Some(Self::HeaderOp),
            0x00FF => Some(Self::EndPic),
            0x009A => Some(Self::DirectBitsRect),
            0x009B => Some(Self::PackedDirectBitsRect),
            0x8200 => Some(Self::CompressedQuickTime),
            _ => None,
        }
    }
}

/// PICT record
#[derive(Debug, Clone)]
pub struct PictRecord {
    /// Opcode
    pub opcode: u16,
    /// Record data
    pub data: Vec<u8>,
}

/// PICT file parser
pub struct PictParser {
    /// PICT header
    pub header: PictHeader,
    /// All opcodes/records
    pub records: Vec<PictRecord>,
    /// Raw PICT data
    data: Vec<u8>,
}

impl PictParser {
    /// Create a new PICT parser from raw data
    pub fn new(data: &[u8]) -> Result<Self> {
        let header = PictHeader::parse(data)?;
        
        let data_start = if header.has_512_header { 512 } else { 0 };
        let mut offset = data_start + 10; // Skip size and frame

        // Skip version opcode if present
        if header.version == PictVersion::V2 && offset + 4 <= data.len() {
            offset += 4; // Skip version opcode and data
        }

        let mut records = Vec::new();

        // Parse records
        while offset < data.len() {
            if offset + 2 > data.len() {
                break;
            }

            let opcode = u16::from_be_bytes([data[offset], data[offset + 1]]);
            offset += 2;

            // Handle EndPic
            if opcode == 0x00FF {
                records.push(PictRecord {
                    opcode,
                    data: Vec::new(),
                });
                break;
            }

            // Determine data size based on opcode
            let data_size = Self::get_opcode_data_size(opcode, data, offset)?;
            
            if offset + data_size > data.len() {
                break;
            }

            let record_data = data[offset..offset + data_size].to_vec();
            offset += data_size;

            records.push(PictRecord {
                opcode,
                data: record_data,
            });
        }

        Ok(Self {
            header,
            records,
            data: data.to_vec(),
        })
    }

    /// Get the data size for an opcode
    fn get_opcode_data_size(opcode: u16, data: &[u8], offset: usize) -> Result<usize> {
        match opcode {
            // Fixed size opcodes
            0x0000 => Ok(0), // Nop
            0x0003 => Ok(2), // TxFont
            0x0004 => Ok(1), // TxFace
            0x0005 => Ok(2), // TxMode
            0x0011 => Ok(2), // Version
            
            // Variable size opcodes - read size from data
            0x0001 | // Clip
            0x00A1 | // Long comment
            0x009A | // DirectBitsRect
            0x009B => { // PackedDirectBitsRect
                if offset + 2 > data.len() {
                    return Err(Error::ParseError("Insufficient data for opcode size".into()));
                }
                let size = u16::from_be_bytes([data[offset], data[offset + 1]]) as usize;
                Ok(size + 2) // Include size field itself
            }
            
            // Default: try to read size field
            _ => {
                if offset + 2 > data.len() {
                    return Ok(0);
                }
                // Many opcodes have a 2-byte size field
                let size = u16::from_be_bytes([data[offset], data[offset + 1]]) as usize;
                Ok(size.min(data.len() - offset))
            }
        }
    }

    /// Get the raw PICT data
    pub fn data(&self) -> &[u8] {
        &self.data
    }

    /// Get width
    pub fn width(&self) -> i32 {
        self.header.width() as i32
    }

    /// Get height
    pub fn height(&self) -> i32 {
        self.header.height() as i32
    }

    /// Get aspect ratio
    pub fn aspect_ratio(&self) -> f64 {
        let w = self.width() as f64;
        let h = self.height() as f64;
        if h == 0.0 {
            1.0
        } else {
            w / h
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_pict_opcodes() {
        assert_eq!(PictOpcode::from_u16(0x00FF), Some(PictOpcode::EndPic));
        assert_eq!(PictOpcode::from_u16(0x0011), Some(PictOpcode::Version));
    }
}

