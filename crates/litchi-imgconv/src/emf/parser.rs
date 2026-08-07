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

use bytes::Bytes;
use litchi_core::error::{Error, Result};
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
    pub description_size: u32,
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

        // Validate the fixed part of EMR_HEADER before trusting any declared
        // offsets or sizes from it.
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

        if raw_header.record_size < 88 || raw_header.record_size % 4 != 0 {
            return Err(Error::ParseError(format!(
                "Invalid EMF header size: {}",
                raw_header.record_size
            )));
        }
        let record_size = usize::try_from(raw_header.record_size)
            .map_err(|_| Error::ParseError("EMF header size does not fit in memory".into()))?;
        if record_size > data.len() {
            return Err(Error::ParseError("EMF header extends beyond data".into()));
        }
        if raw_header.version != 0x0001_0000 {
            return Err(Error::ParseError(format!(
                "Unsupported EMF version: 0x{:08X}",
                raw_header.version
            )));
        }
        if raw_header.reserved != 0 {
            return Err(Error::ParseError(
                "EMF header reserved field is nonzero".into(),
            ));
        }
        if raw_header.size < raw_header.record_size || raw_header.size % 4 != 0 {
            return Err(Error::ParseError(format!(
                "Invalid declared EMF byte size: {}",
                raw_header.size
            )));
        }
        if raw_header.num_records < 2 {
            return Err(Error::ParseError(
                "EMF must contain a header and EOF record".into(),
            ));
        }

        if raw_header.description_size == 0 {
            // offDescription is ignored when no description is present.
        } else {
            if raw_header.description_offset < 88 || raw_header.description_offset % 2 != 0 {
                return Err(Error::ParseError("Invalid EMF description offset".into()));
            }
            let description_bytes = raw_header
                .description_size
                .checked_mul(2)
                .ok_or_else(|| Error::ParseError("EMF description size overflow".into()))?;
            let description_end = raw_header
                .description_offset
                .checked_add(description_bytes)
                .ok_or_else(|| Error::ParseError("EMF description range overflow".into()))?;
            if description_end > raw_header.record_size {
                return Err(Error::ParseError(
                    "EMF description extends beyond the header record".into(),
                ));
            }
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
            description_size: raw_header.description_size,
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
        self.bounds.2.saturating_sub(self.bounds.0)
    }

    /// Get the height of the metafile in device units
    pub fn height(&self) -> i32 {
        self.bounds.3.saturating_sub(self.bounds.1)
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
    /// Record data (excluding type and size), shared with the source buffer
    pub data: Bytes,
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

/// A safely recoverable format deviation accepted by a compatible parser.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum EmfParserWarning {
    /// `EMR_EOF.SizeLast` did not repeat the EOF record's `Size` value.
    EofSizeLastMismatch {
        /// The EOF record size required by MS-EMF.
        expected: u32,
        /// The nonconforming value found in `SizeLast`.
        found: u32,
    },
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
                data: Bytes::copy_from_slice(record_ref.data),
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
        if offset % 4 != 0 {
            return Err(Error::ParseError(format!(
                "Unaligned EMF record offset: {}",
                offset
            )));
        }
        let header_end = offset
            .checked_add(8)
            .ok_or_else(|| Error::ParseError("EMF record offset overflow".into()))?;
        if header_end > data.len() {
            return Err(Error::ParseError("Insufficient data for EMF record".into()));
        }

        // Parse record header using zerocopy - highly optimized, no allocations
        let (header, _) = RawEmfRecordHeader::read_from_prefix(&data[offset..])
            .map_err(|_| Error::ParseError("Invalid EMF record header".into()))?;

        let record_type = header.record_type;
        let size = header.size;

        // Validate size with early return for better branch prediction
        if size < 8 || size % 4 != 0 {
            return Err(Error::ParseError(format!(
                "Invalid EMF record size: {} at offset {}",
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
            data: Bytes::copy_from_slice(self.data),
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
    /// Safely recoverable deviations accepted in compatible mode.
    warnings: Vec<EmfParserWarning>,
    /// Raw EMF data - kept for zero-copy access
    data: Bytes,
    /// Offset to first record after header (cached for performance)
    first_record_offset: usize,
}

impl EmfParser {
    const MAX_PREALLOCATED_RECORDS: usize = 16_384;

    /// Validate the complete record stream and return its first record offset
    /// plus any narrowly scoped compatibility diagnostic.
    fn validate(
        data: &[u8],
        header: &EmfHeader,
        allow_legacy_size_last: bool,
    ) -> Result<(usize, Vec<EmfParserWarning>)> {
        let declared_size = usize::try_from(header.size)
            .map_err(|_| Error::ParseError("EMF declared size does not fit in memory".into()))?;
        if declared_size != data.len() {
            return Err(Error::ParseError(format!(
                "EMF declared byte size {} does not match input length {}",
                declared_size,
                data.len()
            )));
        }

        let first_record_offset =
            usize::try_from(u32::from_le_bytes([data[4], data[5], data[6], data[7]]))
                .map_err(|_| Error::ParseError("EMF header size does not fit in memory".into()))?;
        let remaining = declared_size
            .checked_sub(first_record_offset)
            .ok_or_else(|| Error::ParseError("EMF header exceeds declared size".into()))?;
        let declared_after_header = header
            .num_records
            .checked_sub(1)
            .ok_or_else(|| Error::ParseError("Invalid EMF record count".into()))?;
        if u64::from(declared_after_header) > (remaining / 8) as u64 {
            return Err(Error::ParseError(
                "EMF record count cannot fit in the declared byte size".into(),
            ));
        }

        let mut offset = first_record_offset;
        let mut records_seen = 1u32; // EMR_HEADER
        let mut saw_eof = false;
        let mut warnings = Vec::new();
        while offset < declared_size {
            let (record, consumed) = EmfRecordRef::parse_ref(data, offset)?;
            records_seen = records_seen
                .checked_add(1)
                .ok_or_else(|| Error::ParseError("EMF record count overflow".into()))?;
            let end = offset
                .checked_add(consumed)
                .ok_or_else(|| Error::ParseError("EMF record offset overflow".into()))?;

            if record.record_type == EmfRecordType::Header as u32 {
                return Err(Error::ParseError(
                    "EMF header record appears after the beginning".into(),
                ));
            }
            if record.record_type == EmfRecordType::Eof as u32 {
                if saw_eof {
                    return Err(Error::ParseError("Multiple EMF EOF records".into()));
                }
                if record.size < 20 {
                    return Err(Error::ParseError("EMF EOF record is too short".into()));
                }
                let palette_entries = u32::from_le_bytes([
                    record.data[0],
                    record.data[1],
                    record.data[2],
                    record.data[3],
                ]);
                let palette_offset = u32::from_le_bytes([
                    record.data[4],
                    record.data[5],
                    record.data[6],
                    record.data[7],
                ]);
                let last = record.data.len() - 4;
                let last_size = u32::from_le_bytes([
                    record.data[last],
                    record.data[last + 1],
                    record.data[last + 2],
                    record.data[last + 3],
                ]);
                let palette_bytes = palette_entries
                    .checked_mul(4)
                    .ok_or_else(|| Error::ParseError("EMF EOF palette size overflow".into()))?;
                let palette_end = palette_offset
                    .checked_add(palette_bytes)
                    .ok_or_else(|| Error::ParseError("EMF EOF palette range overflow".into()))?;
                let palette_limit = record
                    .size
                    .checked_sub(4)
                    .ok_or_else(|| Error::ParseError("EMF EOF palette limit underflow".into()))?;
                let invalid_palette_range = palette_entries != 0
                    && (palette_offset < 16
                        || palette_offset % 4 != 0
                        || palette_end > palette_limit);
                if palette_entries != header.num_palette || invalid_palette_range {
                    return Err(Error::ParseError(format!(
                        "Malformed EMF EOF record: palette {}/{}, offset {}, size {}, last {}",
                        palette_entries, header.num_palette, palette_offset, record.size, last_size
                    )));
                }
                if last_size != record.size {
                    if !allow_legacy_size_last {
                        return Err(Error::ParseError(format!(
                            "Malformed EMF EOF record: SizeLast {}, expected {}",
                            last_size, record.size
                        )));
                    }
                    warnings.push(EmfParserWarning::EofSizeLastMismatch {
                        expected: record.size,
                        found: last_size,
                    });
                }
                if end != declared_size {
                    return Err(Error::ParseError(
                        "Trailing data after EMF EOF record".into(),
                    ));
                }
                saw_eof = true;
            } else if saw_eof {
                return Err(Error::ParseError("Record found after EMF EOF".into()));
            }

            offset = end;
        }

        if !saw_eof {
            return Err(Error::ParseError("Missing EMF EOF record".into()));
        }
        if records_seen != header.num_records {
            return Err(Error::ParseError(format!(
                "EMF record count {} does not match declared count {}",
                records_seen, header.num_records
            )));
        }
        Ok((first_record_offset, warnings))
    }

    /// Create a new EMF parser from raw data
    ///
    /// This eagerly parses all records. For large files or streaming scenarios,
    /// consider using `new_lazy()` or iterating with `iter_record_refs()`.
    pub fn new(data: &[u8]) -> Result<Self> {
        Self::parse_internal(Bytes::copy_from_slice(data), true, false)
    }

    /// Create an eager parser that accepts the known legacy `SizeLast`
    /// mismatch while keeping all structural validation strict.
    pub fn new_compatible(data: &[u8]) -> Result<Self> {
        Self::parse_internal(Bytes::copy_from_slice(data), true, true)
    }

    /// Create an eager parser from an owned input buffer without copying it.
    pub fn from_owned(data: Vec<u8>) -> Result<Self> {
        Self::parse_internal(Bytes::from(data), true, false)
    }

    /// Create a compatible eager parser from an owned buffer without copying.
    pub fn from_owned_compatible(data: Vec<u8>) -> Result<Self> {
        Self::parse_internal(Bytes::from(data), true, true)
    }

    fn parse_internal(data: Bytes, eager: bool, allow_legacy_size_last: bool) -> Result<Self> {
        if data.len() < 88 {
            return Err(Error::ParseError("EMF data too short".into()));
        }

        let header = EmfHeader::parse(&data)?;
        let (header_record_size, warnings) =
            Self::validate(&data, &header, allow_legacy_size_last)?;

        if !eager {
            return Ok(Self {
                header,
                records: Vec::new(),
                warnings,
                data,
                first_record_offset: header_record_size,
            });
        }

        // Pre-allocate vector with expected capacity (from header.num_records if available)
        let expected_records = header.num_records.saturating_sub(1) as usize; // -1 for header
        let byte_bound = data
            .len()
            .saturating_sub(header_record_size)
            .checked_div(8)
            .unwrap_or(0);
        let mut records = Vec::with_capacity(
            expected_records
                .min(byte_bound)
                .min(Self::MAX_PREALLOCATED_RECORDS),
        );

        let mut offset = header_record_size;

        // Parse remaining records with optimized loop
        while offset < data.len() {
            let (record_ref, consumed) = EmfRecordRef::parse_ref(&data, offset)?;
            let is_eof = record_ref.record_type == EmfRecordType::Eof as u32;
            let end = offset
                .checked_add(consumed)
                .ok_or_else(|| Error::ParseError("EMF record offset overflow".into()))?;
            records.push(EmfRecord {
                record_type: record_ref.record_type,
                size: record_ref.size,
                data: data.slice(offset + 8..end),
            });
            offset = end;
            if is_eof {
                break;
            }
        }

        // Shrink to fit if we over-allocated
        records.shrink_to_fit();

        Ok(Self {
            header,
            records,
            warnings,
            data,
            first_record_offset: header_record_size,
        })
    }

    /// Create a new EMF parser with header only (lazy record parsing)
    ///
    /// Records are not parsed until accessed. Use `iter_record_refs()` for
    /// zero-copy iteration.
    pub fn new_lazy(data: &[u8]) -> Result<Self> {
        Self::parse_internal(Bytes::copy_from_slice(data), false, false)
    }

    /// Return compatibility diagnostics recorded while parsing.
    #[inline]
    pub fn warnings(&self) -> &[EmfParserWarning] {
        &self.warnings
    }

    /// Get an iterator over record references (zero-copy, most efficient)
    ///
    /// This is the most performant way to process EMF records as it avoids
    /// all allocations and uses zero-copy techniques.
    ///
    /// # Example
    /// ```no_run
    /// # use litchi_imgconv::emf::parser::EmfParser;
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
    use super::*;

    fn put_u16(data: &mut [u8], offset: usize, value: u16) {
        data[offset..offset + 2].copy_from_slice(&value.to_le_bytes());
    }

    fn put_u32(data: &mut [u8], offset: usize, value: u32) {
        data[offset..offset + 4].copy_from_slice(&value.to_le_bytes());
    }

    fn valid_emf() -> Vec<u8> {
        const HEADER_SIZE: usize = 88;
        let mut data = vec![0; HEADER_SIZE + 16 + 20];
        put_u32(&mut data, 0, EmfRecordType::Header as u32);
        put_u32(&mut data, 4, HEADER_SIZE as u32);
        put_u32(&mut data, 16, 10);
        put_u32(&mut data, 20, 10);
        put_u32(&mut data, 32, 100);
        put_u32(&mut data, 36, 100);
        put_u32(&mut data, 40, 0x464D_4520);
        put_u32(&mut data, 44, 0x0001_0000);
        let len = data.len() as u32;
        put_u32(&mut data, 48, len);
        put_u32(&mut data, 52, 3);
        put_u16(&mut data, 56, 1);
        put_u32(&mut data, 72, 100);
        put_u32(&mut data, 76, 100);
        put_u32(&mut data, 80, 25);
        put_u32(&mut data, 84, 25);

        put_u32(&mut data, HEADER_SIZE, EmfRecordType::SetWindowOrgEx as u32);
        put_u32(&mut data, HEADER_SIZE + 4, 16);

        let eof = HEADER_SIZE + 16;
        put_u32(&mut data, eof, EmfRecordType::Eof as u32);
        put_u32(&mut data, eof + 4, 20);
        put_u32(&mut data, eof + 12, 16);
        put_u32(&mut data, eof + 16, 20);
        data
    }

    #[test]
    fn test_emf_signature() {
        // "EMF " in little-endian
        assert_eq!(0x464D4520u32.to_le_bytes(), [0x20, 0x45, 0x4D, 0x46]);
    }

    #[test]
    fn parses_valid_eager_and_lazy_streams() {
        let data = valid_emf();
        let eager = EmfParser::new(&data).unwrap();
        let lazy = EmfParser::new_lazy(&data).unwrap();
        let owned = EmfParser::from_owned(data).unwrap();
        assert_eq!(eager.records.len(), 2);
        assert_eq!(lazy.iter_record_refs().count(), 2);
        assert_eq!(lazy.count_records().unwrap(), 2);
        assert!(eager.warnings().is_empty());
        assert_eq!(
            owned.records[0].data.as_ptr(),
            owned.data().as_ptr().wrapping_add(96)
        );
    }

    #[test]
    fn accepts_header_followed_directly_by_eof() {
        let mut data = valid_emf();
        data.drain(88..104);
        let len = data.len() as u32;
        put_u32(&mut data, 48, len);
        put_u32(&mut data, 52, 2);
        let parsed = EmfParser::new(&data).unwrap();
        assert_eq!(parsed.records.len(), 1);
        assert_eq!(parsed.records[0].record_type, EmfRecordType::Eof as u32);
    }

    #[test]
    fn rejects_invalid_header_fields() {
        let cases: &[(usize, u32)] = &[
            (0, 2),            // Type
            (4, 84),           // Size below the fixed header
            (4, 90),           // Size alignment
            (40, 0),           // Signature
            (44, 0x0002_0000), // Version
            (48, 120),         // Declared bytes
            (52, 1),           // Too few records
        ];
        for &(offset, value) in cases {
            let mut data = valid_emf();
            put_u32(&mut data, offset, value);
            assert!(EmfParser::new(&data).is_err(), "offset {offset}");
        }

        let mut reserved = valid_emf();
        put_u16(&mut reserved, 58, 1);
        assert!(EmfParser::new(&reserved).is_err());
    }

    #[test]
    fn validates_header_description_range() {
        let mut data = valid_emf();
        put_u32(&mut data, 60, 1);
        put_u32(&mut data, 64, 87);
        assert!(EmfParser::new(&data).is_err());

        let mut data = valid_emf();
        put_u32(&mut data, 60, u32::MAX);
        put_u32(&mut data, 64, 88);
        assert!(EmfParser::new(&data).is_err());
    }

    #[test]
    fn rejects_bad_record_sizes_counts_and_truncation() {
        for size in [7, 10, u32::MAX] {
            let mut data = valid_emf();
            put_u32(&mut data, 92, size);
            assert!(EmfParser::new(&data).is_err(), "size {size}");
        }

        let mut count = valid_emf();
        put_u32(&mut count, 52, 4);
        assert!(EmfParser::new(&count).is_err());

        let mut truncated = valid_emf();
        truncated.pop();
        assert!(EmfParser::new(&truncated).is_err());
    }

    #[test]
    fn requires_one_well_formed_eof_at_declared_end() {
        let eof = 88 + 16;
        let mut missing = valid_emf();
        put_u32(&mut missing, eof, 15);
        assert!(EmfParser::new(&missing).is_err());

        let mut malformed = valid_emf();
        put_u32(&mut malformed, eof + 16, 0);
        assert!(EmfParser::new(&malformed).is_err());

        let mut trailing = valid_emf();
        trailing.extend_from_slice(&[15, 0, 0, 0, 8, 0, 0, 0]);
        let len = trailing.len() as u32;
        put_u32(&mut trailing, 48, len);
        put_u32(&mut trailing, 52, 4);
        assert!(EmfParser::new(&trailing).is_err());

        let mut palette_mismatch = valid_emf();
        put_u32(&mut palette_mismatch, 68, 1);
        assert!(EmfParser::new(&palette_mismatch).is_err());
    }

    #[test]
    fn record_ref_checks_offset_arithmetic_and_alignment() {
        let data = valid_emf();
        assert!(EmfRecordRef::parse_ref(&data, 1).is_err());
        assert!(EmfRecordRef::parse_ref(&data, usize::MAX).is_err());
    }

    #[test]
    fn parses_repository_emf_fixtures() {
        for data in [
            include_bytes!(concat!(
                env!("CARGO_MANIFEST_DIR"),
                "/../../test-data/images/emf/wrench.emf"
            ))
            .as_slice(),
            include_bytes!(concat!(
                env!("CARGO_MANIFEST_DIR"),
                "/../../test-data/images/emf/vector_image.emf"
            ))
            .as_slice(),
        ] {
            EmfParser::new(data).unwrap();
        }

        let legacy = include_bytes!(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../test-data/images/emf/jack-sign.emf"
        ));
        assert!(EmfParser::new(legacy).is_err());

        let compatible = EmfParser::new_compatible(legacy).unwrap();
        assert_eq!(
            compatible.warnings(),
            &[EmfParserWarning::EofSizeLastMismatch {
                expected: 20,
                found: 29_868,
            }]
        );
        let owned = EmfParser::from_owned_compatible(legacy.to_vec()).unwrap();
        assert_eq!(owned.records.len(), compatible.records.len());
    }

    #[test]
    fn compatible_mode_relaxes_only_eof_size_last() {
        let eof = 88 + 16;
        let mut legacy = valid_emf();
        let file_size = legacy.len() as u32;
        put_u32(&mut legacy, eof + 16, file_size);
        assert!(EmfParser::new(&legacy).is_err());
        assert_eq!(
            EmfParser::new_compatible(&legacy).unwrap().warnings(),
            &[EmfParserWarning::EofSizeLastMismatch {
                expected: 20,
                found: file_size,
            }]
        );

        let mut missing_eof = legacy.clone();
        put_u32(&mut missing_eof, eof, 15);
        assert!(EmfParser::new_compatible(&missing_eof).is_err());

        let mut wrong_count = legacy.clone();
        put_u32(&mut wrong_count, 52, 4);
        assert!(EmfParser::new_compatible(&wrong_count).is_err());

        let mut palette_mismatch = legacy.clone();
        put_u32(&mut palette_mismatch, 68, 1);
        assert!(EmfParser::new_compatible(&palette_mismatch).is_err());

        let mut trailing = legacy;
        trailing.extend_from_slice(&[15, 0, 0, 0, 8, 0, 0, 0]);
        let length = trailing.len() as u32;
        put_u32(&mut trailing, 48, length);
        put_u32(&mut trailing, 52, 4);
        assert!(EmfParser::new_compatible(&trailing).is_err());
    }
}
