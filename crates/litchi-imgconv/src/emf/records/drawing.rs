/// EMF Drawing Record Structures
///
/// Structures for all EMF drawing operations: shapes, lines, curves, polygons
use super::types::{ColorRef, PointL, RectL, SizeL};
use zerocopy::{FromBytes, IntoBytes};

// Basic shape records

/// EMR_RECTANGLE / EMR_ELLIPSE
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct EmrRectangle {
    pub record_type: u32,
    pub record_size: u32,
    pub rect: RectL,
}

/// EMR_ROUNDRECT
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct EmrRoundRect {
    pub record_type: u32,
    pub record_size: u32,
    pub rect: RectL,
    pub corner: SizeL,
}

/// EMR_ARC / EMR_ARCTO / EMR_CHORD / EMR_PIE
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct EmrArc {
    pub record_type: u32,
    pub record_size: u32,
    pub rect: RectL,
    pub start: PointL,
    pub end: PointL,
}

/// EMR_ANGLEARC
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct EmrAngleArc {
    pub record_type: u32,
    pub record_size: u32,
    pub center: PointL,
    pub radius: u32,
    pub start_angle: f32,
    pub sweep_angle: f32,
}

// Line drawing records

/// EMR_LINETO / EMR_MOVETOEX
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct EmrLineTo {
    pub record_type: u32,
    pub record_size: u32,
    pub point: PointL,
}

/// EMR_SETPIXELV
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct EmrSetPixelV {
    pub record_type: u32,
    pub record_size: u32,
    pub point: PointL,
    pub color: ColorRef,
}

// Polygon records (32-bit coordinates)

/// EMR_POLYGON / EMR_POLYLINE / EMR_POLYBEZIER / EMR_POLYBEZIERTO / EMR_POLYLINETO
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct EmrPolyHeader {
    pub record_type: u32,
    pub record_size: u32,
    pub bounds: RectL,
    pub count: u32,
    // Followed by count PointL structures
}

/// EMR_POLYPOLYLINE / EMR_POLYPOLYGON
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct EmrPolyPolyHeader {
    pub record_type: u32,
    pub record_size: u32,
    pub bounds: RectL,
    pub num_polys: u32,
    pub count: u32,
    // Followed by num_polys u32 (point counts) then count PointL structures
}

// Polygon records (16-bit coordinates)

/// EMR_POLYGON16 / EMR_POLYLINE16 / EMR_POLYBEZIER16 / EMR_POLYBEZIERTO16 / EMR_POLYLINETO16
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct EmrPoly16Header {
    pub record_type: u32,
    pub record_size: u32,
    pub bounds: RectL,
    pub count: u32,
    // Followed by count PointS structures
}

/// EMR_POLYPOLYLINE16 / EMR_POLYPOLYGON16
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct EmrPolyPoly16Header {
    pub record_type: u32,
    pub record_size: u32,
    pub bounds: RectL,
    pub num_polys: u32,
    pub count: u32,
    // Followed by num_polys u32 (point counts) then count PointS structures
}

// PolyDraw records

/// Point type flags for EMR_POLYDRAW
pub mod point_type {
    pub const CLOSEFIGURE: u8 = 0x01;
    pub const LINETO: u8 = 0x02;
    pub const BEZIERTO: u8 = 0x04;
    pub const MOVETO: u8 = 0x06;
}

/// EMR_POLYDRAW
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct EmrPolyDrawHeader {
    pub record_type: u32,
    pub record_size: u32,
    pub bounds: RectL,
    pub count: u32,
    // Followed by count PointL structures, then count u8 point types
}

/// EMR_POLYDRAW16
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct EmrPolyDraw16Header {
    pub record_type: u32,
    pub record_size: u32,
    pub bounds: RectL,
    pub count: u32,
    // Followed by count PointS structures, then count u8 point types
}

// Fill modes

/// Polygon fill mode for EMR_SETPOLYFILLMODE
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u32)]
pub enum PolyFillMode {
    /// ALTERNATE mode (even-odd rule)
    Alternate = 1,
    /// WINDING mode (non-zero winding rule)
    Winding = 2,
}

impl PolyFillMode {
    /// Convert to SVG fill-rule
    pub fn to_svg_fill_rule(self) -> &'static str {
        match self {
            Self::Alternate => "evenodd",
            Self::Winding => "nonzero",
        }
    }
}

/// Arc direction for EMR_SETARCDIRECTION
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u32)]
pub enum ArcDirection {
    /// Counter-clockwise
    CounterClockwise = 1,
    /// Clockwise
    Clockwise = 2,
}

// Gradient fill

/// Gradient fill modes
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u32)]
pub enum GradientFillMode {
    /// Horizontal gradient
    Horizontal = 0,
    /// Vertical gradient
    Vertical = 1,
    /// Triangle gradient (use 3-point mesh)
    Triangle = 2,
}

/// EMR_GRADIENTFILL
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct EmrGradientFillHeader {
    pub record_type: u32,
    pub record_size: u32,
    pub bounds: RectL,
    pub num_vertices: u32,
    pub num_triangles: u32,
    pub mode: u32,
    // Followed by num_vertices TriVertex, then num_triangles GradientTriangle/GradientRect
}

/// Gradient vertex
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct TriVertex {
    pub x: i32,
    pub y: i32,
    pub red: u16,
    pub green: u16,
    pub blue: u16,
    pub alpha: u16,
}

/// Gradient rectangle (2 vertices)
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct GradientRect {
    pub upper_left: u32,
    pub lower_right: u32,
}

/// Gradient triangle (3 vertices)
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct GradientTriangle {
    pub vertex1: u32,
    pub vertex2: u32,
    pub vertex3: u32,
}

// Flood fill

/// Flood fill mode
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u32)]
pub enum FloodFillMode {
    /// Fill up to border color
    Border = 0,
    /// Fill within surface of same color
    Surface = 1,
}

/// EMR_EXTFLOODFILL
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct EmrExtFloodFill {
    pub record_type: u32,
    pub record_size: u32,
    pub start: PointL,
    pub color: ColorRef,
    pub mode: u32,
}

/// Helper for parsing polygon data
pub struct PolygonData<'a> {
    pub points: &'a [u8],
    pub count: usize,
    pub is_16bit: bool,
}

impl<'a> PolygonData<'a> {
    /// Create from 32-bit polygon record
    pub fn from_poly32(data: &'a [u8], offset: usize, count: usize) -> Option<Self> {
        let points_size = count * 8; // 8 bytes per PointL
        if data.len() < offset + points_size {
            return None;
        }
        Some(Self {
            points: &data[offset..offset + points_size],
            count,
            is_16bit: false,
        })
    }

    /// Create from 16-bit polygon record
    pub fn from_poly16(data: &'a [u8], offset: usize, count: usize) -> Option<Self> {
        let points_size = count * 4; // 4 bytes per PointS
        if data.len() < offset + points_size {
            return None;
        }
        Some(Self {
            points: &data[offset..offset + points_size],
            count,
            is_16bit: true,
        })
    }

    /// Iterator over points as (x, y) tuples
    pub fn iter_points(&self) -> PolygonPointIter<'_> {
        PolygonPointIter {
            data: self.points,
            index: 0,
            count: self.count,
            is_16bit: self.is_16bit,
        }
    }
}

/// Iterator over polygon points
pub struct PolygonPointIter<'a> {
    data: &'a [u8],
    index: usize,
    count: usize,
    is_16bit: bool,
}

impl Iterator for PolygonPointIter<'_> {
    type Item = (i32, i32);

    fn next(&mut self) -> Option<Self::Item> {
        if self.index >= self.count {
            return None;
        }

        let point = if self.is_16bit {
            // 16-bit points (PointS)
            let offset = self.index * 4;
            if offset + 4 > self.data.len() {
                return None;
            }
            let x = i16::from_le_bytes([self.data[offset], self.data[offset + 1]]) as i32;
            let y = i16::from_le_bytes([self.data[offset + 2], self.data[offset + 3]]) as i32;
            (x, y)
        } else {
            // 32-bit points (PointL)
            let offset = self.index * 8;
            if offset + 8 > self.data.len() {
                return None;
            }
            let x = i32::from_le_bytes([
                self.data[offset],
                self.data[offset + 1],
                self.data[offset + 2],
                self.data[offset + 3],
            ]);
            let y = i32::from_le_bytes([
                self.data[offset + 4],
                self.data[offset + 5],
                self.data[offset + 6],
                self.data[offset + 7],
            ]);
            (x, y)
        };

        self.index += 1;
        Some(point)
    }

    fn size_hint(&self) -> (usize, Option<usize>) {
        let remaining = self.count - self.index;
        (remaining, Some(remaining))
    }
}

impl ExactSizeIterator for PolygonPointIter<'_> {}
