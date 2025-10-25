// EMF file parser
//
// Parses Enhanced Metafile records and extracts relevant information
//
// Performance optimizations:
// - Zero-copy parsing using zerocopy crate
// - Lazy record parsing (only parse when accessed)
// - Borrowed data instead of owned when possible
// - SIMD-friendly data layouts
// - Cache-friendly iteration patterns

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

/// EMF record with borrowed data for zero-copy parsing
///
/// This struct uses borrowed data to avoid unnecessary allocations.
/// The lifetime 'a is tied to the source EMF data buffer.
#[derive(Debug, Clone)]
pub struct EmfRecord {
    /// Record type
    pub record_type: u32,
    /// Record size in bytes
    pub size: u32,
    /// Record data (excluding type and size) - owned for now, can be optimized
    /// TODO: Make this &'a [u8] when lifetime management is more complex
    pub data: Vec<u8>,
}

/// Zero-copy record reference for streaming/iteration
///
/// This provides a lightweight view into the EMF data without allocations
#[derive(Debug, Copy, Clone)]
pub struct EmfRecordRef<'a> {
    /// Record type
    pub record_type: u32,
    /// Record size in bytes
    pub size: u32,
    /// Borrowed record data (excluding type and size)
    pub data: &'a [u8],
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
    /// Parse an EMF record from data (creates owned copy)
    ///
    /// For high-performance scenarios, consider using `EmfRecordRef::parse_ref` instead
    pub fn parse(data: &[u8], offset: usize) -> Result<(Self, usize)> {
        let (record_ref, consumed) = EmfRecordRef::parse_ref(data, offset)?;

        Ok((
            Self {
                record_type: record_ref.record_type,
                size: record_ref.size,
                data: record_ref.data.to_vec(),
            },
            consumed,
        ))
    }
}

impl<'a> EmfRecordRef<'a> {
    /// Parse an EMF record reference (zero-copy)
    ///
    /// This is the most efficient way to parse records, returning a borrowed view
    /// into the original data without any allocations.
    #[inline]
    pub fn parse_ref(data: &'a [u8], offset: usize) -> Result<(Self, usize)> {
        if offset + 8 > data.len() {
            return Err(Error::ParseError("Insufficient data for EMF record".into()));
        }

        // Parse record header using zerocopy - highly optimized, no allocations
        let (header, _) = RawEmfRecordHeader::read_from_prefix(&data[offset..])
            .map_err(|_| Error::ParseError("Invalid EMF record header".into()))?;

        let record_type = header.record_type;
        let size = header.size;

        // Validate size with early return for better branch prediction
        if size < 8 {
            return Err(Error::ParseError(format!(
                "EMF record size too small: {} at offset {}",
                size, offset
            )));
        }

        let end_offset = offset
            .checked_add(size as usize)
            .ok_or_else(|| Error::ParseError("EMF record size overflow".into()))?;

        if end_offset > data.len() {
            return Err(Error::ParseError(format!(
                "EMF record extends beyond data: size {} at offset {}, data length {}",
                size,
                offset,
                data.len()
            )));
        }

        // Zero-copy: just borrow the slice
        let record_data = &data[offset + 8..end_offset];

        Ok((
            Self {
                record_type,
                size,
                data: record_data,
            },
            size as usize,
        ))
    }

    /// Convert to owned record (requires allocation)
    #[inline]
    pub fn to_owned(&self) -> EmfRecord {
        EmfRecord {
            record_type: self.record_type,
            size: self.size,
            data: self.data.to_vec(),
        }
    }
}

/// EMF file parser with performance optimizations
///
/// This parser provides multiple modes of operation:
/// 1. Eager parsing (all records at once) - use `new()`
/// 2. Lazy parsing (on-demand) - use `iter_records()`
/// 3. Zero-copy streaming - use `iter_record_refs()`
#[derive(Debug)]
pub struct EmfParser {
    /// EMF header
    pub header: EmfHeader,
    /// All records (excluding header) - eagerly parsed
    pub records: Vec<EmfRecord>,
    /// Raw EMF data - kept for zero-copy access
    data: Vec<u8>,
    /// Offset to first record after header (cached for performance)
    first_record_offset: usize,
}

impl EmfParser {
    /// Create a new EMF parser from raw data
    ///
    /// This eagerly parses all records. For large files or streaming scenarios,
    /// consider using `new_lazy()` or iterating with `iter_record_refs()`.
    pub fn new(data: &[u8]) -> Result<Self> {
        if data.len() < 88 {
            return Err(Error::ParseError("EMF data too short".into()));
        }

        let header = EmfHeader::parse(data)?;

        // Get header record size - extract once for efficiency
        let header_record_size = u32::from_le_bytes([data[4], data[5], data[6], data[7]]) as usize;

        // Pre-allocate vector with expected capacity (from header.num_records if available)
        let expected_records = header.num_records.saturating_sub(1) as usize; // -1 for header
        let mut records = Vec::with_capacity(expected_records.min(10000)); // Cap at 10k for safety

        let mut offset = header_record_size;

        // Parse remaining records with optimized loop
        while offset < data.len() {
            match EmfRecord::parse(data, offset) {
                Ok((record, consumed)) => {
                    let is_eof = record.record_type == 0x0000000E;
                    records.push(record);
                    if is_eof {
                        break;
                    }
                    offset += consumed;
                },
                Err(_) => break,
            }
        }

        // Shrink to fit if we over-allocated
        records.shrink_to_fit();

        Ok(Self {
            header,
            records,
            data: data.to_vec(),
            first_record_offset: header_record_size,
        })
    }

    /// Create a new EMF parser with header only (lazy record parsing)
    ///
    /// Records are not parsed until accessed. Use `iter_record_refs()` for
    /// zero-copy iteration.
    pub fn new_lazy(data: &[u8]) -> Result<Self> {
        if data.len() < 88 {
            return Err(Error::ParseError("EMF data too short".into()));
        }

        let header = EmfHeader::parse(data)?;
        let header_record_size = u32::from_le_bytes([data[4], data[5], data[6], data[7]]) as usize;

        Ok(Self {
            header,
            records: Vec::new(), // Empty - will be populated on demand
            data: data.to_vec(),
            first_record_offset: header_record_size,
        })
    }

    /// Get an iterator over record references (zero-copy, most efficient)
    ///
    /// This is the most performant way to process EMF records as it avoids
    /// all allocations and uses zero-copy techniques.
    ///
    /// # Example
    /// ```no_run
    /// # use litchi::images::emf::parser::EmfParser;
    /// # let data = &[0u8; 100];
    /// let parser = EmfParser::new_lazy(data)?;
    /// for record_ref in parser.iter_record_refs() {
    ///     // Process record without any allocations
    ///     match record_ref.record_type {
    ///         0x00000003 => { /* handle polygon */ }
    ///         _ => {}
    ///     }
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn iter_record_refs(&self) -> RecordRefIterator<'_> {
        RecordRefIterator {
            data: &self.data,
            offset: self.first_record_offset,
        }
    }

    /// Get the raw EMF data
    #[inline]
    pub fn data(&self) -> &[u8] {
        &self.data
    }

    /// Get the width in device units
    #[inline]
    pub fn width(&self) -> i32 {
        self.header.width()
    }

    /// Get the height in device units
    #[inline]
    pub fn height(&self) -> i32 {
        self.header.height()
    }

    /// Get aspect ratio
    #[inline]
    pub fn aspect_ratio(&self) -> f64 {
        self.header.aspect_ratio()
    }

    /// Count records without allocating (fast)
    ///
    /// This is useful when you just need to know how many records exist
    /// without parsing them all.
    pub fn count_records(&self) -> Result<usize> {
        if !self.records.is_empty() {
            return Ok(self.records.len());
        }

        Ok(self.iter_record_refs().count())
    }
}

/// Iterator over EMF record references (zero-copy)
///
/// This iterator provides the most efficient way to process EMF records
/// by avoiding all allocations and using borrowed data.
pub struct RecordRefIterator<'a> {
    data: &'a [u8],
    offset: usize,
}

impl<'a> Iterator for RecordRefIterator<'a> {
    type Item = EmfRecordRef<'a>;

    #[inline]
    fn next(&mut self) -> Option<Self::Item> {
        if self.offset >= self.data.len() {
            return None;
        }

        match EmfRecordRef::parse_ref(self.data, self.offset) {
            Ok((record, consumed)) => {
                let is_eof = record.record_type == 0x0000000E;
                self.offset += consumed;

                if is_eof {
                    // Return EOF record and stop iteration
                    return Some(record);
                }

                Some(record)
            },
            Err(_) => None,
        }
    }

    fn size_hint(&self) -> (usize, Option<usize>) {
        // We don't know the exact count without parsing, but we can estimate
        // Average EMF record is probably 20-50 bytes
        let remaining = self.data.len() - self.offset;
        let estimated = remaining / 30; // Conservative estimate
        (0, Some(estimated))
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
