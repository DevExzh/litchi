// EMF file parser
//
// Parses Enhanced Metafile records and extracts relevant information

use crate::common::error::{Error, Result};
use zerocopy::FromBytes;

/// EMF record types
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u32)]
pub enum EmfRecordType {
    Header = 0x00000001,
    PolyBezier = 0x00000002,
    Polygon = 0x00000003,
    Polyline = 0x00000004,
    PolyBezierTo = 0x00000005,
    PolyLineTo = 0x00000006,
    PolyPolyline = 0x00000007,
    PolyPolygon = 0x00000008,
    SetWindowExtEx = 0x00000009,
    SetWindowOrgEx = 0x0000000A,
    SetViewportExtEx = 0x0000000B,
    SetViewportOrgEx = 0x0000000C,
    SetBrushOrgEx = 0x0000000D,
    Eof = 0x0000000E,
    // Add more as needed
}

impl EmfRecordType {
    /// Create from u32 value
    pub fn from_u32(value: u32) -> Option<Self> {
        match value {
            0x00000001 => Some(Self::Header),
            0x00000002 => Some(Self::PolyBezier),
            0x00000003 => Some(Self::Polygon),
            0x00000004 => Some(Self::Polyline),
            0x00000005 => Some(Self::PolyBezierTo),
            0x00000006 => Some(Self::PolyLineTo),
            0x00000007 => Some(Self::PolyPolyline),
            0x00000008 => Some(Self::PolyPolygon),
            0x00000009 => Some(Self::SetWindowExtEx),
            0x0000000A => Some(Self::SetWindowOrgEx),
            0x0000000B => Some(Self::SetViewportExtEx),
            0x0000000C => Some(Self::SetViewportOrgEx),
            0x0000000D => Some(Self::SetBrushOrgEx),
            0x0000000E => Some(Self::Eof),
            _ => None,
        }
    }
}

/// EMF header information
#[derive(Debug, Clone)]
pub struct EmfHeader {
    /// Bounds of the metafile in device units
    pub bounds: (i32, i32, i32, i32),
    /// Frame rectangle in .01 millimeter units
    pub frame: (i32, i32, i32, i32),
    /// Signature (must be 0x464D4520 "EMF ")
    pub signature: u32,
    /// Version
    pub version: u32,
    /// Size of the file in bytes
    pub size: u32,
    /// Number of records
    pub num_records: u32,
    /// Number of handles in handle table
    pub num_handles: u16,
    /// Size of description string
    pub description_size: u16,
    /// Offset to description string
    pub description_offset: u32,
    /// Number of palette entries
    pub num_palette: u32,
    /// Width of reference device in pixels
    pub device_width: i32,
    /// Height of reference device in pixels
    pub device_height: i32,
    /// Width of reference device in millimeters
    pub device_width_mm: i32,
    /// Height of reference device in millimeters
    pub device_height_mm: i32,
}

/// Raw EMF header structure for zerocopy parsing (88 bytes total)
#[derive(Debug, Clone, FromBytes)]
#[repr(C)]
struct RawEmfHeader {
    /// Record type (must be 0x00000001)
    record_type: u32,
    /// Record size
    record_size: u32,
    /// Bounds left
    bounds_left: i32,
    /// Bounds top
    bounds_top: i32,
    /// Bounds right
    bounds_right: i32,
    /// Bounds bottom
    bounds_bottom: i32,
    /// Frame left
    frame_left: i32,
    /// Frame top
    frame_top: i32,
    /// Frame right
    frame_right: i32,
    /// Frame bottom
    frame_bottom: i32,
    /// Signature (must be 0x464D4520 "EMF ")
    signature: u32,
    /// Version
    version: u32,
    /// Size of the file in bytes
    size: u32,
    /// Number of records
    num_records: u32,
    /// Number of handles in handle table
    num_handles: u16,
    /// Reserved field
    reserved: u16,
    /// Size of description string
    description_size: u32,
    /// Offset to description string
    description_offset: u32,
    /// Number of palette entries
    num_palette: u32,
    /// Width of reference device in pixels
    device_width: i32,
    /// Height of reference device in pixels
    device_height: i32,
    /// Width of reference device in millimeters
    device_width_mm: i32,
    /// Height of reference device in millimeters
    device_height_mm: i32,
}

impl EmfHeader {
    /// Parse EMF header from data
    pub fn parse(data: &[u8]) -> Result<Self> {
        if data.len() < 88 {
            return Err(Error::ParseError("EMF header too short".into()));
        }

        // Parse header using zerocopy - read_from_prefix returns (value, remaining)
        let (raw_header, _) = RawEmfHeader::read_from_prefix(data)
            .map_err(|_| Error::ParseError("Invalid EMF header format".into()))?;

        // Validate record type
        if raw_header.record_type != 0x00000001 {
            return Err(Error::ParseError(format!(
                "Invalid EMF header record type: 0x{:08X}",
                raw_header.record_type
            )));
        }

        // Validate signature
        if raw_header.signature != 0x464D4520 {
            // "EMF " in little-endian
            return Err(Error::ParseError(format!(
                "Invalid EMF signature: 0x{:08X}",
                raw_header.signature
            )));
        }

        Ok(Self {
            bounds: (
                raw_header.bounds_left,
                raw_header.bounds_top,
                raw_header.bounds_right,
                raw_header.bounds_bottom,
            ),
            frame: (
                raw_header.frame_left,
                raw_header.frame_top,
                raw_header.frame_right,
                raw_header.frame_bottom,
            ),
            signature: raw_header.signature,
            version: raw_header.version,
            size: raw_header.size,
            num_records: raw_header.num_records,
            num_handles: raw_header.num_handles,
            description_size: raw_header.description_size as u16,
            description_offset: raw_header.description_offset,
            num_palette: raw_header.num_palette,
            device_width: raw_header.device_width,
            device_height: raw_header.device_height,
            device_width_mm: raw_header.device_width_mm,
            device_height_mm: raw_header.device_height_mm,
        })
    }

    /// Get the width of the metafile in device units
    pub fn width(&self) -> i32 {
        self.bounds.2 - self.bounds.0
    }

    /// Get the height of the metafile in device units
    pub fn height(&self) -> i32 {
        self.bounds.3 - self.bounds.1
    }

    /// Get aspect ratio (width / height)
    pub fn aspect_ratio(&self) -> f64 {
        let w = self.width() as f64;
        let h = self.height() as f64;
        if h == 0.0 { 1.0 } else { w / h }
    }
}

/// EMF record
#[derive(Debug, Clone)]
pub struct EmfRecord {
    /// Record type
    pub record_type: u32,
    /// Record size in bytes
    pub size: u32,
    /// Record data (excluding type and size)
    pub data: Vec<u8>,
}

/// Raw EMF record header for zerocopy parsing (8 bytes)
#[derive(Debug, Clone, zerocopy::FromBytes)]
#[repr(C)]
struct RawEmfRecordHeader {
    /// Record type
    record_type: u32,
    /// Record size in bytes
    size: u32,
}

impl EmfRecord {
    /// Parse an EMF record from data
    pub fn parse(data: &[u8], offset: usize) -> Result<(Self, usize)> {
        if offset + 8 > data.len() {
            return Err(Error::ParseError("Insufficient data for EMF record".into()));
        }

        // Parse record header using zerocopy
        let (header, _) = RawEmfRecordHeader::read_from_prefix(&data[offset..])
            .map_err(|_| Error::ParseError("Invalid EMF record header".into()))?;

        let record_type = header.record_type;
        let size = header.size;

        if size < 8 || offset + size as usize > data.len() {
            return Err(Error::ParseError(format!(
                "Invalid EMF record size: {} at offset {}",
                size, offset
            )));
        }

        let record_data = data[offset + 8..offset + size as usize].to_vec();

        Ok((
            Self {
                record_type,
                size,
                data: record_data,
            },
            size as usize,
        ))
    }
}

/// EMF file parser
#[derive(Debug)]
pub struct EmfParser {
    /// EMF header
    pub header: EmfHeader,
    /// All records (excluding header)
    pub records: Vec<EmfRecord>,
    /// Raw EMF data
    data: Vec<u8>,
}

impl EmfParser {
    /// Create a new EMF parser from raw data
    pub fn new(data: &[u8]) -> Result<Self> {
        if data.len() < 88 {
            return Err(Error::ParseError("EMF data too short".into()));
        }

        let header = EmfHeader::parse(data)?;
        let mut records = Vec::new();

        // Get header record size from offset 4 (not the file size field)
        let header_record_size = u32::from_le_bytes([data[4], data[5], data[6], data[7]]) as usize;
        let mut offset = header_record_size; // Skip header record

        // Parse remaining records
        while offset < data.len() {
            match EmfRecord::parse(data, offset) {
                Ok((record, consumed)) => {
                    // Check for EOF record
                    if record.record_type == 0x0000000E {
                        records.push(record);
                        break;
                    }
                    records.push(record);
                    offset += consumed;
                },
                Err(_) => break,
            }
        }

        Ok(Self {
            header,
            records,
            data: data.to_vec(),
        })
    }

    /// Get the raw EMF data
    pub fn data(&self) -> &[u8] {
        &self.data
    }

    /// Get the width in device units
    pub fn width(&self) -> i32 {
        self.header.width()
    }

    /// Get the height in device units
    pub fn height(&self) -> i32 {
        self.header.height()
    }

    /// Get aspect ratio
    pub fn aspect_ratio(&self) -> f64 {
        self.header.aspect_ratio()
    }
}

#[cfg(test)]
mod tests {
    #[test]
    fn test_emf_signature() {
        // "EMF " in little-endian
        assert_eq!(0x464D4520u32.to_le_bytes(), [0x20, 0x45, 0x4D, 0x46]);
    }
}
