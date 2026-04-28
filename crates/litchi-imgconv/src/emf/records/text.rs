/// EMF Text Output Records
use super::types::{PointL, RectL};
use zerocopy::{FromBytes, IntoBytes};

/// Text output options
pub mod text_options {
    pub const NO_RECT: u32 = 0x0000; // No background rectangle
    pub const OPAQUE: u32 = 0x0002; // Draw opaque background
    pub const CLIPPED: u32 = 0x0004; // Clip to rectangle
    pub const GLYPH_INDEX: u32 = 0x0010; // Use glyph indices not chars
    pub const RTLREADING: u32 = 0x0080; // Right-to-left reading order
    pub const NUMERICSLOCAL: u32 = 0x0400; // Use local numerics
    pub const NUMERICSLATIN: u32 = 0x0800; // Use Latin numerics
    pub const IGNORELANGUAGE: u32 = 0x1000; // Ignore language
    pub const PDY: u32 = 0x2000; // Dx array contains x+y offsets
}

/// EMR_EXTTEXTOUTA / EMR_EXTTEXTOUTW header (without record type/size which are parsed separately)
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct EmrExtTextOutHeader {
    pub bounds: RectL,
    pub graphics_mode: u32,
    pub ex_scale: f32,
    pub ey_scale: f32,
    pub text: EmrTextInfo,
}

/// Text information structure
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct EmrTextInfo {
    pub reference: PointL,
    pub num_chars: u32,
    pub off_string: u32,
    pub options: u32,
    pub rectangle: RectL,
    pub off_dx: u32,
}

/// EMR_POLYTEXTOUTA / EMR_POLYTEXTOUTW header (without record type/size which are parsed separately)
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct EmrPolyTextOutHeader {
    pub bounds: RectL,
    pub graphics_mode: u32,
    pub ex_scale: f32,
    pub ey_scale: f32,
    pub num_strings: u32,
    // Followed by num_strings EmrTextInfo structures
}

/// EMR_SMALLTEXTOUT header (without record type/size which are parsed separately)
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct EmrSmallTextOutHeader {
    pub x: i32,
    pub y: i32,
    pub num_chars: u32,
    pub fu_options: u32,
    pub graphics_mode: u32,
    pub ex_scale: f32,
    pub ey_scale: f32,
    pub bounds: RectL,
    // Followed by text string
}

/// EMR_SETTEXTJUSTIFICATION (without record type/size which are parsed separately)
#[derive(Debug, Clone, Copy, IntoBytes, FromBytes)]
#[repr(C)]
pub struct EmrSetTextJustification {
    pub num_break_extra: i32,
    pub num_break_count: i32,
}

/// Graphics mode for text scaling
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u32)]
pub enum GraphicsMode {
    Compatible = 1, // GM_COMPATIBLE
    Advanced = 2,   // GM_ADVANCED
}
