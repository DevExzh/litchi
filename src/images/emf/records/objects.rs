/// EMF Object Creation Records
///
/// Records for creating GDI objects: pens, brushes, fonts, palettes
use super::types::ColorRef;
use zerocopy::{FromBytes, IntoBytes};

// Object management

/// EMR_SELECTOBJECT
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct EmrSelectObject {
    pub record_type: u32,
    pub record_size: u32,
    pub object_index: u32,
}

/// EMR_DELETEOBJECT
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct EmrDeleteObject {
    pub record_type: u32,
    pub record_size: u32,
    pub object_index: u32,
}

// Pen creation

/// Pen styles
pub mod pen_style {
    pub const SOLID: u32 = 0;
    pub const DASH: u32 = 1;
    pub const DOT: u32 = 2;
    pub const DASHDOT: u32 = 3;
    pub const DASHDOTDOT: u32 = 4;
    pub const NULL: u32 = 5;
    pub const INSIDEFRAME: u32 = 6;
    pub const USERSTYLE: u32 = 7;
    pub const ALTERNATE: u32 = 8;

    // End cap styles
    pub const ENDCAP_ROUND: u32 = 0x00000000;
    pub const ENDCAP_SQUARE: u32 = 0x00000100;
    pub const ENDCAP_FLAT: u32 = 0x00000200;

    // Join styles
    pub const JOIN_ROUND: u32 = 0x00000000;
    pub const JOIN_BEVEL: u32 = 0x00001000;
    pub const JOIN_MITER: u32 = 0x00002000;

    // Type flags
    pub const COSMETIC: u32 = 0x00000000;
    pub const GEOMETRIC: u32 = 0x00010000;
}

/// EMR_CREATEPEN
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct EmrCreatePen {
    pub record_type: u32,
    pub record_size: u32,
    pub object_index: u32,
    pub pen_style: u32,
    pub width: u32,   // Only x component used
    pub _unused: u32, // y component (unused)
    pub color: ColorRef,
}

/// EMR_EXTCREATEPEN header
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct EmrExtCreatePenHeader {
    pub record_type: u32,
    pub record_size: u32,
    pub object_index: u32,
    pub off_bmi: u32,  // Offset to bitmap info
    pub cb_bmi: u32,   // Size of bitmap info
    pub off_bits: u32, // Offset to bitmap bits
    pub cb_bits: u32,  // Size of bitmap bits
    pub pen_style: u32,
    pub width: u32,
    pub brush_style: u32,
    pub color: ColorRef,
    pub brush_hatch: u32,
    pub num_style_entries: u32,
    // Followed by style entries if USERSTYLE
}

// Brush creation

/// Brush styles
pub mod brush_style {
    pub const SOLID: u32 = 0;
    pub const NULL: u32 = 1;
    pub const HATCHED: u32 = 2;
    pub const PATTERN: u32 = 3;
    pub const INDEXED: u32 = 4;
    pub const DIBPATTERN: u32 = 5;
    pub const DIBPATTERNPT: u32 = 6;
    pub const PATTERN8X8: u32 = 7;
    pub const DIBPATTERN8X8: u32 = 8;
    pub const MONOPATTERN: u32 = 9;
}

/// Hatch styles
pub mod hatch_style {
    pub const HORIZONTAL: u32 = 0;
    pub const VERTICAL: u32 = 1;
    pub const FDIAGONAL: u32 = 2; // Forward diagonal (//)
    pub const BDIAGONAL: u32 = 3; // Backward diagonal (\\)
    pub const CROSS: u32 = 4;
    pub const DIAGCROSS: u32 = 5;
}

/// EMR_CREATEBRUSHINDIRECT
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct EmrCreateBrushIndirect {
    pub record_type: u32,
    pub record_size: u32,
    pub object_index: u32,
    pub brush_style: u32,
    pub color: ColorRef,
    pub brush_hatch: u32,
}

/// EMR_CREATEDIBPATTERNBRUSHPT header
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct EmrCreateDIBPatternBrushPtHeader {
    pub record_type: u32,
    pub record_size: u32,
    pub object_index: u32,
    pub usage: u32,    // DIB color table usage
    pub off_bmi: u32,  // Offset to bitmap info
    pub cb_bmi: u32,   // Size of bitmap info
    pub off_bits: u32, // Offset to bitmap bits
    pub cb_bits: u32,  // Size of bitmap bits
                       // Followed by bitmap data
}

/// EMR_CREATEMONOBRUSH header
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct EmrCreateMonoBrushHeader {
    pub record_type: u32,
    pub record_size: u32,
    pub object_index: u32,
    pub usage: u32,    // DIB color table usage
    pub off_bmi: u32,  // Offset to bitmap info
    pub cb_bmi: u32,   // Size of bitmap info
    pub off_bits: u32, // Offset to bitmap bits
    pub cb_bits: u32,  // Size of bitmap bits
                       // Followed by bitmap data
}

// Font creation

/// Font weight constants
pub mod font_weight {
    pub const DONTCARE: i32 = 0;
    pub const THIN: i32 = 100;
    pub const EXTRALIGHT: i32 = 200;
    pub const LIGHT: i32 = 300;
    pub const NORMAL: i32 = 400;
    pub const MEDIUM: i32 = 500;
    pub const SEMIBOLD: i32 = 600;
    pub const BOLD: i32 = 700;
    pub const EXTRABOLD: i32 = 800;
    pub const HEAVY: i32 = 900;
}

/// Font character sets
pub mod charset {
    pub const ANSI: u8 = 0;
    pub const DEFAULT: u8 = 1;
    pub const SYMBOL: u8 = 2;
    pub const SHIFTJIS: u8 = 128;
    pub const HANGUL: u8 = 129;
    pub const GB2312: u8 = 134;
    pub const CHINESEBIG5: u8 = 136;
    pub const OEM: u8 = 255;
}

/// Font pitch and family flags
pub mod pitch_and_family {
    pub const DEFAULT_PITCH: u8 = 0;
    pub const FIXED_PITCH: u8 = 1;
    pub const VARIABLE_PITCH: u8 = 2;

    pub const FF_DONTCARE: u8 = 0 << 4;
    pub const FF_ROMAN: u8 = 1 << 4;
    pub const FF_SWISS: u8 = 2 << 4;
    pub const FF_MODERN: u8 = 3 << 4;
    pub const FF_SCRIPT: u8 = 4 << 4;
    pub const FF_DECORATIVE: u8 = 5 << 4;
}

/// Font quality
pub mod font_quality {
    pub const DEFAULT: u8 = 0;
    pub const DRAFT: u8 = 1;
    pub const PROOF: u8 = 2;
    pub const NONANTIALIASED: u8 = 3;
    pub const ANTIALIASED: u8 = 4;
    pub const CLEARTYPE: u8 = 5;
}

/// Maximum font face name length
pub const LF_FACESIZE: usize = 32;

/// EMR_EXTCREATEFONTINDIRECTW header
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct EmrExtCreateFontIndirectWHeader {
    pub record_type: u32,
    pub record_size: u32,
    pub object_index: u32,
    // Followed by LOGFONTW structure
}

/// LOGFONTW structure (partial - fixed part)
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct LogFontW {
    pub height: i32,
    pub width: i32,
    pub escapement: i32,
    pub orientation: i32,
    pub weight: i32,
    pub italic: u8,
    pub underline: u8,
    pub strike_out: u8,
    pub char_set: u8,
    pub out_precision: u8,
    pub clip_precision: u8,
    pub quality: u8,
    pub pitch_and_family: u8,
    // Followed by 32 u16 characters for face name
}

// Palette operations

/// EMR_CREATEPALETTE header
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct EmrCreatePaletteHeader {
    pub record_type: u32,
    pub record_size: u32,
    pub object_index: u32,
    pub version: u16, // Always 0x0300
    pub num_entries: u16,
    // Followed by num_entries PALETTEENTRY structures
}

/// Palette entry
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct PaletteEntry {
    pub red: u8,
    pub green: u8,
    pub blue: u8,
    pub flags: u8,
}

/// EMR_SELECTPALETTE
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct EmrSelectPalette {
    pub record_type: u32,
    pub record_size: u32,
    pub palette_index: u32,
}

/// EMR_SETPALETTEENTRIES header
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct EmrSetPaletteEntriesHeader {
    pub record_type: u32,
    pub record_size: u32,
    pub palette_index: u32,
    pub start: u32,
    pub num_entries: u32,
    // Followed by num_entries PALETTEENTRY structures
}

/// EMR_RESIZEPALETTE
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct EmrResizePalette {
    pub record_type: u32,
    pub record_size: u32,
    pub palette_index: u32,
    pub num_entries: u32,
}

// Color space operations

/// Color space type
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u32)]
pub enum ColorSpaceType {
    Calibrated = 0,
    Srgb = 0x73524742,    // 'sRGB'
    Windows = 0x57696E20, // 'Win '
}

/// EMR_CREATECOLORSPACE / EMR_CREATECOLORSPACEW header
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct EmrCreateColorSpaceHeader {
    pub record_type: u32,
    pub record_size: u32,
    pub color_space_index: u32,
    // Followed by LOGCOLORSPACEW structure
}

/// EMR_SETCOLORSPACE
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct EmrSetColorSpace {
    pub record_type: u32,
    pub record_size: u32,
    pub color_space_index: u32,
}

/// EMR_DELETECOLORSPACE
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct EmrDeleteColorSpace {
    pub record_type: u32,
    pub record_size: u32,
    pub color_space_index: u32,
}
