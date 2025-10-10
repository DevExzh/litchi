//! Common data types used across the library
//!
//! This module defines shared data structures, particularly for binary formats
//! like PICT that require specific memory layouts and endianness handling.

/// PICT rectangle structure (big-endian format)
#[repr(C)]
#[derive(Debug, Clone, Copy)]
pub struct PictRect {
    pub top: i16,
    pub left: i16,
    pub bottom: i16,
    pub right: i16,
}

/// PICT bitmap structure
#[repr(C)]
#[derive(Debug, Clone)]
pub struct PictBitmap {
    pub row_bytes: i16,
    pub bounds: PictRect,
    pub src_rect: PictRect,
    pub dst_rect: PictRect,
    pub mode: i16,
    pub data: Vec<u8>,
}

/// PICT region structure
#[repr(C)]
#[derive(Debug, Clone)]
pub struct PictRegion {
    pub region_size: i16,
    pub rect: PictRect,
}

impl PictRect {
    /// Get width of rectangle
    #[inline(always)]
    pub fn width(&self) -> i16 {
        self.right - self.left
    }

    /// Get height of rectangle
    #[inline(always)]
    pub fn height(&self) -> i16 {
        self.bottom - self.top
    }
}

impl PictBitmap {
    /// Get bitmap width
    #[inline(always)]
    pub fn width(&self) -> i16 {
        self.bounds.width()
    }

    /// Get bitmap height
    #[inline(always)]
    pub fn height(&self) -> i16 {
        self.bounds.height()
    }
}
