/// EMF Bitmap Operation Records
use super::types::{RectL, XForm};
use zerocopy::{FromBytes, IntoBytes};

/// DIB color table usage
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u32)]
pub enum DibUsage {
    RgbColors = 0,  // Colors are RGB values
    PalIndices = 1, // Colors are palette indices
}

/// Common raster operations (ternary ROP codes)
pub mod rop {
    pub const SRCCOPY: u32 = 0x00CC0020; // dest = source
    pub const SRCPAINT: u32 = 0x00EE0086; // dest = source OR dest
    pub const SRCAND: u32 = 0x008800C6; // dest = source AND dest
    pub const SRCINVERT: u32 = 0x00660046; // dest = source XOR dest
    pub const SRCERASE: u32 = 0x00440328; // dest = source AND (NOT dest)
    pub const NOTSRCCOPY: u32 = 0x00330008; // dest = (NOT source)
    pub const NOTSRCERASE: u32 = 0x001100A6; // dest = (NOT src) AND (NOT dest)
    pub const MERGECOPY: u32 = 0x00C000CA; // dest = (source AND pattern)
    pub const MERGEPAINT: u32 = 0x00BB0226; // dest = (NOT source) OR dest
    pub const PATCOPY: u32 = 0x00F00021; // dest = pattern
    pub const PATPAINT: u32 = 0x00FB0A09; // dest = DPSnoo
    pub const PATINVERT: u32 = 0x005A0049; // dest = pattern XOR dest
    pub const DSTINVERT: u32 = 0x00550009; // dest = (NOT dest)
    pub const BLACKNESS: u32 = 0x00000042; // dest = BLACK
    pub const WHITENESS: u32 = 0x00FF0062; // dest = WHITE
}

/// EMR_BITBLT header
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct EmrBitBltHeader {
    pub record_type: u32,
    pub record_size: u32,
    pub bounds: RectL,
    pub x_dest: i32,
    pub y_dest: i32,
    pub cx_dest: i32,
    pub cy_dest: i32,
    pub rop: u32,
    pub x_src: i32,
    pub y_src: i32,
    pub xform_src: XForm,
    pub bk_color_src: u32,
    pub usage_src: u32,
    pub off_bmi_src: u32,
    pub cb_bmi_src: u32,
    pub off_bits_src: u32,
    pub cb_bits_src: u32,
    // Followed by bitmap data if present
}

/// EMR_STRETCHBLT header
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct EmrStretchBltHeader {
    pub record_type: u32,
    pub record_size: u32,
    pub bounds: RectL,
    pub x_dest: i32,
    pub y_dest: i32,
    pub cx_dest: i32,
    pub cy_dest: i32,
    pub rop: u32,
    pub x_src: i32,
    pub y_src: i32,
    pub xform_src: XForm,
    pub bk_color_src: u32,
    pub usage_src: u32,
    pub off_bmi_src: u32,
    pub cb_bmi_src: u32,
    pub off_bits_src: u32,
    pub cb_bits_src: u32,
    pub cx_src: i32,
    pub cy_src: i32,
    // Followed by bitmap data if present
}

/// EMR_ALPHABLEND header
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct EmrAlphaBlendHeader {
    pub record_type: u32,
    pub record_size: u32,
    pub bounds: RectL,
    pub x_dest: i32,
    pub y_dest: i32,
    pub cx_dest: i32,
    pub cy_dest: i32,
    pub blend_function: BlendFunction,
    pub x_src: i32,
    pub y_src: i32,
    pub xform_src: XForm,
    pub bk_color_src: u32,
    pub usage_src: u32,
    pub off_bmi_src: u32,
    pub cb_bmi_src: u32,
    pub off_bits_src: u32,
    pub cb_bits_src: u32,
    pub cx_src: i32,
    pub cy_src: i32,
    // Followed by bitmap data
}

/// Blend function for alpha blending
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct BlendFunction {
    pub blend_op: u8,           // Always 0 (AC_SRC_OVER)
    pub blend_flags: u8,        // Must be 0
    pub src_constant_alpha: u8, // 0-255 (0=transparent, 255=opaque)
    pub alpha_format: u8,       // 0 or 1 (AC_SRC_ALPHA)
}

/// EMR_TRANSPARENTBLT header
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct EmrTransparentBltHeader {
    pub record_type: u32,
    pub record_size: u32,
    pub bounds: RectL,
    pub x_dest: i32,
    pub y_dest: i32,
    pub cx_dest: i32,
    pub cy_dest: i32,
    pub transparent_color: u32,
    pub x_src: i32,
    pub y_src: i32,
    pub xform_src: XForm,
    pub bk_color_src: u32,
    pub usage_src: u32,
    pub off_bmi_src: u32,
    pub cb_bmi_src: u32,
    pub off_bits_src: u32,
    pub cb_bits_src: u32,
    pub cx_src: i32,
    pub cy_src: i32,
    // Followed by bitmap data
}

/// EMR_STRETCHDIBITS header
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct EmrStretchDiBitsHeader {
    pub record_type: u32,
    pub record_size: u32,
    pub bounds: RectL,
    pub x_dest: i32,
    pub y_dest: i32,
    pub x_src: i32,
    pub y_src: i32,
    pub cx_src: i32,
    pub cy_src: i32,
    pub off_bmi_src: u32,
    pub cb_bmi_src: u32,
    pub off_bits_src: u32,
    pub cb_bits_src: u32,
    pub usage_src: u32,
    pub rop: u32,
    pub cx_dest: i32,
    pub cy_dest: i32,
    // Followed by bitmap data
}

/// EMR_SETDIBITSTODEVICE header
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct EmrSetDiBitsToDeviceHeader {
    pub record_type: u32,
    pub record_size: u32,
    pub bounds: RectL,
    pub x_dest: i32,
    pub y_dest: i32,
    pub x_src: i32,
    pub y_src: i32,
    pub cx_src: i32,
    pub cy_src: i32,
    pub off_bmi_src: u32,
    pub cb_bmi_src: u32,
    pub off_bits_src: u32,
    pub cb_bits_src: u32,
    pub usage_src: u32,
    pub scan_start: u32,
    pub num_scans: u32,
    // Followed by bitmap data
}

/// Bitmap compression types
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u32)]
pub enum BitmapCompression {
    Rgb = 0,       // Uncompressed
    Rle8 = 1,      // 8-bit RLE
    Rle4 = 2,      // 4-bit RLE
    Bitfields = 3, // Uncompressed with color masks
    Jpeg = 4,      // JPEG image
    Png = 5,       // PNG image
}
