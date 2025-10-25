/// Performance-optimized EMF record parsing utilities
///
/// This module provides high-performance parsing helpers that minimize allocations
/// and maximize cache efficiency.
use crate::common::error::{Error, Result};
use crate::images::emf::records::types::*;
use zerocopy::FromBytes;

/// Fast parser for point arrays (POINTL)
///
/// Uses zerocopy for maximum performance. The input slice should be properly aligned.
#[inline]
pub fn parse_pointl_array(data: &[u8], count: usize) -> Result<Vec<PointL>> {
    let required_size = count * std::mem::size_of::<PointL>();

    if data.len() < required_size {
        return Err(Error::ParseError(format!(
            "Insufficient data for {} points: need {}, have {}",
            count,
            required_size,
            data.len()
        )));
    }

    // Parse points using zerocopy
    let mut points = Vec::with_capacity(count);
    let mut offset = 0;
    for _ in 0..count {
        if let Ok((point, _)) = PointL::read_from_prefix(&data[offset..]) {
            points.push(point);
            offset += std::mem::size_of::<PointL>();
        } else {
            return Err(Error::ParseError("Failed to parse point".into()));
        }
    }

    Ok(points)
}

/// Fast parser for 16-bit point arrays (POINTS)
#[inline]
pub fn parse_points_array(data: &[u8], count: usize) -> Result<Vec<PointS>> {
    let required_size = count * std::mem::size_of::<PointS>();

    if data.len() < required_size {
        return Err(Error::ParseError(format!(
            "Insufficient data for {} 16-bit points: need {}, have {}",
            count,
            required_size,
            data.len()
        )));
    }

    let mut points = Vec::with_capacity(count);
    let mut offset = 0;
    for _ in 0..count {
        if let Ok((point, _)) = PointS::read_from_prefix(&data[offset..]) {
            points.push(point);
            offset += std::mem::size_of::<PointS>();
        } else {
            return Err(Error::ParseError("Failed to parse 16-bit point".into()));
        }
    }

    Ok(points)
}

/// Fast parser for polygon count arrays
///
/// Returns a Vec of u32 values representing counts for each polygon
#[inline]
pub fn parse_poly_counts(data: &[u8], num_polys: usize) -> Result<Vec<u32>> {
    let required_size = num_polys * 4; // 4 bytes per u32

    if data.len() < required_size {
        return Err(Error::ParseError(format!(
            "Insufficient data for {} polygon counts",
            num_polys
        )));
    }

    let mut counts = Vec::with_capacity(num_polys);
    for i in 0..num_polys {
        let offset = i * 4;
        let count = u32::from_le_bytes([
            data[offset],
            data[offset + 1],
            data[offset + 2],
            data[offset + 3],
        ]);
        counts.push(count);
    }

    Ok(counts)
}

/// Parse rectangle efficiently
#[inline]
pub fn parse_rectl(data: &[u8]) -> Result<RectL> {
    let (rect, _) = RectL::read_from_prefix(data)
        .map_err(|_| Error::ParseError("Failed to parse RECTL".into()))?;
    Ok(rect)
}

/// Parse size efficiently
#[inline]
pub fn parse_sizel(data: &[u8]) -> Result<SizeL> {
    let (size, _) = SizeL::read_from_prefix(data)
        .map_err(|_| Error::ParseError("Failed to parse SIZEL".into()))?;
    Ok(size)
}

/// Parse point efficiently
#[inline]
pub fn parse_pointl(data: &[u8]) -> Result<PointL> {
    let (point, _) = PointL::read_from_prefix(data)
        .map_err(|_| Error::ParseError("Failed to parse POINTL".into()))?;
    Ok(point)
}

/// Parse color reference efficiently
#[inline]
pub fn parse_colorref(data: &[u8]) -> Result<ColorRef> {
    let (color, _) = ColorRef::read_from_prefix(data)
        .map_err(|_| Error::ParseError("Failed to parse COLORREF".into()))?;
    Ok(color)
}

/// Parse transform (XFORM) efficiently
#[inline]
pub fn parse_xform(data: &[u8]) -> Result<XForm> {
    let (xform, _) = XForm::read_from_prefix(data)
        .map_err(|_| Error::ParseError("Failed to parse XFORM".into()))?;
    Ok(xform)
}

/// Extract DIB bitmap data efficiently
///
/// This parses the BITMAPINFOHEADER and extracts the pixel data without copying
/// unless necessary. Returns (header info, color table, pixel data)
pub struct DibData<'a> {
    pub width: i32,
    pub height: i32,
    pub bit_count: u16,
    pub compression: u32,
    pub color_table: &'a [u8],
    pub pixel_data: &'a [u8],
}

/// Parse DIB (Device Independent Bitmap) structure
///
/// This is optimized for minimal allocations and zero-copy where possible
pub fn parse_dib<'a>(
    data: &'a [u8],
    bmi_offset: usize,
    bmi_size: usize,
    px_offset: usize,
    px_size: usize,
) -> Result<DibData<'a>> {
    // Validate offsets
    if bmi_offset + bmi_size > data.len() {
        return Err(Error::ParseError(
            "DIB bitmap info extends beyond data".into(),
        ));
    }
    if px_offset + px_size > data.len() {
        return Err(Error::ParseError(
            "DIB pixel data extends beyond data".into(),
        ));
    }

    let bmi_data = &data[bmi_offset..bmi_offset + bmi_size];
    let px_data = &data[px_offset..px_offset + px_size];

    // Parse BITMAPINFOHEADER (minimum 40 bytes)
    if bmi_data.len() < 40 {
        return Err(Error::ParseError("BITMAPINFOHEADER too small".into()));
    }

    // Extract fields using little-endian reads
    let header_size = u32::from_le_bytes([bmi_data[0], bmi_data[1], bmi_data[2], bmi_data[3]]);
    let width = i32::from_le_bytes([bmi_data[4], bmi_data[5], bmi_data[6], bmi_data[7]]);
    let height = i32::from_le_bytes([bmi_data[8], bmi_data[9], bmi_data[10], bmi_data[11]]);
    let bit_count = u16::from_le_bytes([bmi_data[14], bmi_data[15]]);
    let compression = u32::from_le_bytes([bmi_data[16], bmi_data[17], bmi_data[18], bmi_data[19]]);

    // Color table starts after header
    let color_table_offset = header_size as usize;
    let color_table = if color_table_offset < bmi_data.len() {
        &bmi_data[color_table_offset..]
    } else {
        &[]
    };

    Ok(DibData {
        width,
        height,
        bit_count,
        compression,
        color_table,
        pixel_data: px_data,
    })
}

/// Record type dispatcher using jump table for O(1) lookup
///
/// This provides faster dispatch than match statements for record processing
#[derive(Debug, Clone, Copy)]
#[repr(u8)]
pub enum RecordCategory {
    /// Drawing operations (lines, curves, shapes)
    Drawing = 0,
    /// State management (DC save/restore, transforms)
    State = 1,
    /// Object management (pens, brushes, fonts)
    Object = 2,
    /// Path operations
    Path = 3,
    /// Bitmap operations
    Bitmap = 4,
    /// Text operations
    Text = 5,
    /// Clipping operations
    Clip = 6,
    /// Control (EOF, comment)
    Control = 7,
    /// Unknown/unsupported
    Unknown = 255,
}

/// Categorize a record type for fast dispatch
///
/// This uses a lookup pattern that's friendly to branch prediction
#[inline]
pub const fn categorize_record(record_type: u32) -> RecordCategory {
    match record_type {
        // Control first (most specific)
        0x01 | 0x0E => RecordCategory::Control, // Header, EOF
        0x46 => RecordCategory::Control,        // Comment

        // Drawing operations
        0x02..=0x08 | 0x10..=0x13 | 0x18 | 0x38 | 0x39 | 0x47 => RecordCategory::Drawing,

        // State operations
        0x09..=0x0D | 0x21..=0x24 | 0x31 | 0x43 | 0x44 | 0x49..=0x4B => RecordCategory::State,

        // Object operations (excluding 0x53-0x54 which are text)
        0x25..=0x29 | 0x37 | 0x50..=0x52 | 0x55..=0x57 => RecordCategory::Object,

        // Text operations (must come before Object to avoid overlap)
        0x53..=0x54 | 0x63 => RecordCategory::Text,

        // Path operations
        0x3A | 0x3C..=0x3D | 0x3E..=0x42 => RecordCategory::Path,

        // Clipping operations
        0x3B | 0x45 | 0x48 => RecordCategory::Clip,

        // Bitmap operations (excluding 0x50 which is in Object)
        0x4C..=0x4F | 0x75..=0x77 => RecordCategory::Bitmap,

        _ => RecordCategory::Unknown,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_categorize_record() {
        assert!(matches!(categorize_record(0x03), RecordCategory::Drawing)); // Polygon
        assert!(matches!(categorize_record(0x25), RecordCategory::Object)); // SelectObject
        assert!(matches!(categorize_record(0x3C), RecordCategory::Path)); // BeginPath
        assert!(matches!(categorize_record(0x0E), RecordCategory::Control)); // EOF
    }
}
