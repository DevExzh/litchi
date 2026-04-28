/// EMF State Management Records
///
/// Records for managing device context state: transforms, mapping modes, colors, etc.
use super::types::{ColorRef, PointL, SizeL, XForm};
use zerocopy::{FromBytes, IntoBytes};

// Transform records

/// EMR_SETWORLDTRANSFORM
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct EmrSetWorldTransform {
    pub record_type: u32,
    pub record_size: u32,
    pub xform: XForm,
}

/// Modify world transform mode
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u32)]
pub enum ModifyWorldTransformMode {
    /// Reset to identity then apply transform
    Set = 1,
    /// Multiply current by transform (left multiply)
    LeftMultiply = 2,
    /// Multiply current by transform (right multiply)
    RightMultiply = 3,
}

/// EMR_MODIFYWORLDTRANSFORM
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct EmrModifyWorldTransform {
    pub record_type: u32,
    pub record_size: u32,
    pub xform: XForm,
    pub mode: u32,
}

// Mapping mode records

/// Mapping mode enumeration
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u32)]
pub enum MapMode {
    /// Text mode (1:1, not constrained)
    Text = 1,
    /// Low res (.1 mm per unit)
    LoMetric = 2,
    /// High res (.01 mm per unit)
    HiMetric = 3,
    /// Low res (.01 in per unit)
    LoEnglish = 4,
    /// High res (.001 in per unit)
    HiEnglish = 5,
    /// Twips (1/1440 in per unit)
    Twips = 6,
    /// Isotropic (x == y scaling)
    Isotropic = 7,
    /// Anisotropic (arbitrary scaling)
    Anisotropic = 8,
}

impl MapMode {
    /// Get logical units per inch
    pub fn units_per_inch(self) -> Option<f64> {
        match self {
            Self::Text => None,              // Device dependent
            Self::LoMetric => Some(254.0),   // 0.1mm units
            Self::HiMetric => Some(2540.0),  // 0.01mm units
            Self::LoEnglish => Some(100.0),  // 0.01in units
            Self::HiEnglish => Some(1000.0), // 0.001in units
            Self::Twips => Some(1440.0),     // 1/1440in units
            Self::Isotropic => None,         // Device dependent
            Self::Anisotropic => None,       // Device dependent
        }
    }
}

/// EMR_SETMAPMODE
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct EmrSetMapMode {
    pub record_type: u32,
    pub record_size: u32,
    pub mode: u32,
}

/// EMR_SETWINDOWEXTEX / EMR_SETVIEWPORTEXTEX
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct EmrSetExtEx {
    pub record_type: u32,
    pub record_size: u32,
    pub extent: SizeL,
}

/// EMR_SETWINDOWORGEX / EMR_SETVIEWPORTORGEX / EMR_SETBRUSHORGEX
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct EmrSetOrgEx {
    pub record_type: u32,
    pub record_size: u32,
    pub origin: PointL,
}

/// EMR_SCALEWINDOWEXTEX / EMR_SCALEVIEWPORTEXTEX
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct EmrScaleExtEx {
    pub record_type: u32,
    pub record_size: u32,
    pub x_num: i32,
    pub x_denom: i32,
    pub y_num: i32,
    pub y_denom: i32,
}

/// EMR_OFFSETCLIPRGN
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct EmrOffsetClipRgn {
    pub record_type: u32,
    pub record_size: u32,
    pub offset: PointL,
}

// Color and background records

/// EMR_SETTEXTCOLOR / EMR_SETBKCOLOR
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct EmrSetColorRef {
    pub record_type: u32,
    pub record_size: u32,
    pub color: ColorRef,
}

/// Background mode
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u32)]
pub enum BackgroundMode {
    /// Transparent background
    Transparent = 1,
    /// Opaque background
    Opaque = 2,
}

/// EMR_SETBKMODE
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct EmrSetBkMode {
    pub record_type: u32,
    pub record_size: u32,
    pub mode: u32,
}

/// EMR_SETPOLYFILLMODE
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct EmrSetPolyFillMode {
    pub record_type: u32,
    pub record_size: u32,
    pub mode: u32,
}

// Raster operation modes

/// ROP2 binary raster operation mode
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u32)]
pub enum Rop2 {
    Black = 1,        // 0
    NotMergePen = 2,  // ~(D | P)
    MaskNotPen = 3,   // ~P & D
    NotCopyPen = 4,   // ~P
    MaskPenNot = 5,   // P & ~D
    Not = 6,          // ~D
    XorPen = 7,       // P ^ D
    NotMaskPen = 8,   // ~(P & D)
    MaskPen = 9,      // P & D
    NotXorPen = 10,   // ~(P ^ D)
    Nop = 11,         // D
    MergeNotPen = 12, // ~P | D
    CopyPen = 13,     // P
    MergePenNot = 14, // P | ~D
    MergePen = 15,    // P | D
    White = 16,       // 1
}

/// EMR_SETROP2
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct EmrSetRop2 {
    pub record_type: u32,
    pub record_size: u32,
    pub mode: u32,
}

/// Stretch blit mode
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u32)]
pub enum StretchBltMode {
    /// Delete scanned lines
    BlackOnWhite = 1,
    /// OR scanned lines
    WhiteOnBlack = 2,
    /// Delete scanned lines (same as BlackOnWhite)
    ColorOnColor = 3,
    /// Use elimination and interpolation (halftone)
    Halftone = 4,
}

/// EMR_SETSTRETCHBLTMODE
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct EmrSetStretchBltMode {
    pub record_type: u32,
    pub record_size: u32,
    pub mode: u32,
}

// Text alignment

/// Text alignment flags
pub mod text_align {
    pub const NOUPDATECP: u32 = 0x0000; // Don't update current position
    pub const UPDATECP: u32 = 0x0001; // Update current position
    pub const LEFT: u32 = 0x0000; // Horizontal: left
    pub const RIGHT: u32 = 0x0002; // Horizontal: right
    pub const CENTER: u32 = 0x0006; // Horizontal: center
    pub const TOP: u32 = 0x0000; // Vertical: top
    pub const BOTTOM: u32 = 0x0008; // Vertical: bottom
    pub const BASELINE: u32 = 0x0018; // Vertical: baseline
    pub const RTLREADING: u32 = 0x0100; // Right-to-left reading order
    pub const VTA: u32 = 0x0100; // Vertical text alignment
}

/// EMR_SETTEXTALIGN
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct EmrSetTextAlign {
    pub record_type: u32,
    pub record_size: u32,
    pub mode: u32,
}

/// EMR_SETMITERLIMIT
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct EmrSetMiterLimit {
    pub record_type: u32,
    pub record_size: u32,
    pub limit: u32,
}

/// EMR_SETARCDIRECTION
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct EmrSetArcDirection {
    pub record_type: u32,
    pub record_size: u32,
    pub direction: u32,
}

// DC state management

/// EMR_SAVEDC
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct EmrSaveDc {
    pub record_type: u32,
    pub record_size: u32,
}

/// EMR_RESTOREDC
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct EmrRestoreDc {
    pub record_type: u32,
    pub record_size: u32,
    pub saved_dc: i32, // Negative = relative, positive = absolute
}

/// EMR_SETMETARGN / EMR_SETMAPPERFLAGS / EMR_REALIZEPALETTE
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct EmrSimple {
    pub record_type: u32,
    pub record_size: u32,
}

// Color adjustment

/// Color adjustment structure
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct ColorAdjustment {
    pub size: u16,
    pub flags: u16,
    pub illuminant_index: u16,
    pub red_gamma: u16,
    pub green_gamma: u16,
    pub blue_gamma: u16,
    pub reference_black: u16,
    pub reference_white: u16,
    pub contrast: i16,
    pub brightness: i16,
    pub colorfulness: i16,
    pub red_green_tint: i16,
}

/// EMR_SETCOLORADJUSTMENT
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct EmrSetColorAdjustment {
    pub record_type: u32,
    pub record_size: u32,
    pub color_adjustment: ColorAdjustment,
}

// ICM (Image Color Management)

/// ICM mode
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u32)]
pub enum IcmMode {
    Off = 1,
    On = 2,
    Query = 3,
    DoneOutsideDc = 4,
}

/// EMR_SETICMMODE
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct EmrSetIcmMode {
    pub record_type: u32,
    pub record_size: u32,
    pub mode: u32,
}

// Layout mode

/// Layout mode flags
pub mod layout {
    pub const RTL: u32 = 0x00000001; // Right-to-left layout
    pub const BTT: u32 = 0x00000002; // Bottom-to-top layout
    pub const VBH: u32 = 0x00000004; // Vertical before horizontal
    pub const BITMAPORIENTATIONPRESERVED: u32 = 0x00000008; // Preserve bitmap orientation
}

/// EMR_SETLAYOUT
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct EmrSetLayout {
    pub record_type: u32,
    pub record_size: u32,
    pub layout_mode: u32,
}
