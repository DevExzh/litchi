// EMF file parser
//
// Parses Enhanced Metafile records and extracts relevant information

use crate::common::error::{Error, Result};

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

impl EmfHeader {
    /// Parse EMF header from data
    pub fn parse(data: &[u8]) -> Result<Self> {
        if data.len() < 88 {
            return Err(Error::ParseError("EMF header too short".into()));
        }

        // Parse record type and size
        let record_type = u32::from_le_bytes([data[0], data[1], data[2], data[3]]);
        if record_type != 0x00000001 {
            return Err(Error::ParseError(format!(
                "Invalid EMF header record type: 0x{:08X}",
                record_type
            )));
        }

        let _record_size = u32::from_le_bytes([data[4], data[5], data[6], data[7]]);

        // Parse bounds (4 i32 values)
        let bounds_left = i32::from_le_bytes([data[8], data[9], data[10], data[11]]);
        let bounds_top = i32::from_le_bytes([data[12], data[13], data[14], data[15]]);
        let bounds_right = i32::from_le_bytes([data[16], data[17], data[18], data[19]]);
        let bounds_bottom = i32::from_le_bytes([data[20], data[21], data[22], data[23]]);

        // Parse frame (4 i32 values)
        let frame_left = i32::from_le_bytes([data[24], data[25], data[26], data[27]]);
        let frame_top = i32::from_le_bytes([data[28], data[29], data[30], data[31]]);
        let frame_right = i32::from_le_bytes([data[32], data[33], data[34], data[35]]);
        let frame_bottom = i32::from_le_bytes([data[36], data[37], data[38], data[39]]);

        let signature = u32::from_le_bytes([data[40], data[41], data[42], data[43]]);
        if signature != 0x464D4520 {
            // "EMF " in little-endian
            return Err(Error::ParseError(format!(
                "Invalid EMF signature: 0x{:08X}",
                signature
            )));
        }

        let version = u32::from_le_bytes([data[44], data[45], data[46], data[47]]);
        let size = u32::from_le_bytes([data[48], data[49], data[50], data[51]]);
        let num_records = u32::from_le_bytes([data[52], data[53], data[54], data[55]]);
        let num_handles = u16::from_le_bytes([data[56], data[57]]);
        let _reserved = u16::from_le_bytes([data[58], data[59]]);
        let description_size = u32::from_le_bytes([data[60], data[61], data[62], data[63]]);
        let description_offset = u32::from_le_bytes([data[64], data[65], data[66], data[67]]);
        let num_palette = u32::from_le_bytes([data[68], data[69], data[70], data[71]]);
        let device_width = i32::from_le_bytes([data[72], data[73], data[74], data[75]]);
        let device_height = i32::from_le_bytes([data[76], data[77], data[78], data[79]]);
        let device_width_mm = i32::from_le_bytes([data[80], data[81], data[82], data[83]]);
        let device_height_mm = i32::from_le_bytes([data[84], data[85], data[86], data[87]]);

        Ok(Self {
            bounds: (bounds_left, bounds_top, bounds_right, bounds_bottom),
            frame: (frame_left, frame_top, frame_right, frame_bottom),
            signature,
            version,
            size,
            num_records,
            num_handles,
            description_size: description_size as u16,
            description_offset,
            num_palette,
            device_width,
            device_height,
            device_width_mm,
            device_height_mm,
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
        if h == 0.0 {
            1.0
        } else {
            w / h
        }
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

impl EmfRecord {
    /// Parse an EMF record from data
    pub fn parse(data: &[u8], offset: usize) -> Result<(Self, usize)> {
        if offset + 8 > data.len() {
            return Err(Error::ParseError("Insufficient data for EMF record".into()));
        }

        let record_type = u32::from_le_bytes([
            data[offset],
            data[offset + 1],
            data[offset + 2],
            data[offset + 3],
        ]);
        let size = u32::from_le_bytes([
            data[offset + 4],
            data[offset + 5],
            data[offset + 6],
            data[offset + 7],
        ]);

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
        let mut offset = header.size as usize; // Skip header record

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
                }
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

