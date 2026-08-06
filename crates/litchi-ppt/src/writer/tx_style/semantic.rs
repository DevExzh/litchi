//! TxMasterStyleAtom builder (MS-PPT 2.9.45)
//!
//! Constructs text master style atoms with proper formatting structures
//! using zerocopy for binary serialization.

use bitflags::bitflags;
use zerocopy_derive::{Immutable, IntoBytes, KnownLayout};

// =============================================================================
// TxMasterStyleAtom Instance Types (MS-PPT 2.9.45)
// =============================================================================

/// TxMasterStyleAtom instance types
pub mod tx_style_instance {
    pub const TITLE: u16 = 0;
    pub const BODY: u16 = 1;
    pub const NOTES: u16 = 2;
    pub const OTHER: u16 = 4;
    pub const CENTER_BODY: u16 = 5;
    pub const CENTER_TITLE: u16 = 6;
    pub const HALF_BODY: u16 = 7;
    pub const QUARTER_BODY: u16 = 8;
}

// =============================================================================
// TextPFException mask bits (MS-PPT 2.9.18)
// =============================================================================

bitflags! {
    /// Paragraph formatting mask bits (MS-PPT 2.9.18)
    #[derive(Debug, Clone, Copy, PartialEq, Eq)]
    pub struct ParagraphMask: u32 {
        /// Has bullet
        const HAS_BULLET = 0x0001;
        /// Bullet has font
        const BULLET_HAS_FONT = 0x0002;
        /// Bullet has color
        const BULLET_HAS_COLOR = 0x0004;
        /// Bullet has size
        const BULLET_HAS_SIZE = 0x0008;
        /// Bullet font index present
        const BULLET_FONT = 0x0010;
        /// Bullet color present
        const BULLET_COLOR = 0x0020;
        /// Bullet size present
        const BULLET_SIZE = 0x0040;
        /// Bullet character present
        const BULLET_CHAR = 0x0080;
        /// Left margin present
        const LEFT_MARGIN = 0x0100;
        /// Unused
        const UNUSED = 0x0200;
        /// Indent present
        const INDENT = 0x0400;
        /// Alignment present
        const ALIGN = 0x0800;
        /// Line spacing present
        const LINE_SPACING = 0x1000;
        /// Space before present
        const SPACE_BEFORE = 0x2000;
        /// Space after present
        const SPACE_AFTER = 0x4000;
        /// Default tab size present
        const DEFAULT_TAB_SIZE = 0x8000;
        /// Font alignment present
        const FONT_ALIGN = 0x0001_0000;
        /// Wrap flags present
        const WRAP_FLAGS = 0x0002_0000;
        /// Text direction present
        const TEXT_DIRECTION = 0x0004_0000;
    }
}

// Keep module for backward compatibility
pub mod pf_mask {
    pub const HAS_BULLET: u32 = super::ParagraphMask::HAS_BULLET.bits();
    pub const BULLET_HAS_FONT: u32 = super::ParagraphMask::BULLET_HAS_FONT.bits();
    pub const BULLET_HAS_COLOR: u32 = super::ParagraphMask::BULLET_HAS_COLOR.bits();
    pub const BULLET_HAS_SIZE: u32 = super::ParagraphMask::BULLET_HAS_SIZE.bits();
    pub const BULLET_FONT: u32 = super::ParagraphMask::BULLET_FONT.bits();
    pub const BULLET_COLOR: u32 = super::ParagraphMask::BULLET_COLOR.bits();
    pub const BULLET_SIZE: u32 = super::ParagraphMask::BULLET_SIZE.bits();
    pub const BULLET_CHAR: u32 = super::ParagraphMask::BULLET_CHAR.bits();
    pub const LEFT_MARGIN: u32 = super::ParagraphMask::LEFT_MARGIN.bits();
    pub const UNUSED: u32 = super::ParagraphMask::UNUSED.bits();
    pub const INDENT: u32 = super::ParagraphMask::INDENT.bits();
    pub const ALIGN: u32 = super::ParagraphMask::ALIGN.bits();
    pub const LINE_SPACING: u32 = super::ParagraphMask::LINE_SPACING.bits();
    pub const SPACE_BEFORE: u32 = super::ParagraphMask::SPACE_BEFORE.bits();
    pub const SPACE_AFTER: u32 = super::ParagraphMask::SPACE_AFTER.bits();
    pub const DEFAULT_TAB_SIZE: u32 = super::ParagraphMask::DEFAULT_TAB_SIZE.bits();
    pub const FONT_ALIGN: u32 = super::ParagraphMask::FONT_ALIGN.bits();
    pub const WRAP_FLAGS: u32 = super::ParagraphMask::WRAP_FLAGS.bits();
    pub const TEXT_DIRECTION: u32 = super::ParagraphMask::TEXT_DIRECTION.bits();
}

// =============================================================================
// TextCFException mask bits (MS-PPT 2.9.6)
// =============================================================================

bitflags! {
    /// Character formatting mask bits (MS-PPT 2.9.6)
    #[derive(Debug, Clone, Copy, PartialEq, Eq)]
    pub struct CharacterMask: u16 {
        /// Bold
        const BOLD = 0x0001;
        /// Italic
        const ITALIC = 0x0002;
        /// Underline
        const UNDERLINE = 0x0004;
        /// Unused
        const UNUSED1 = 0x0008;
        /// Shadow
        const SHADOW = 0x0010;
        /// FEHint (East Asian)
        const FEHINT = 0x0020;
        /// Unused
        const UNUSED2 = 0x0040;
        /// Kumimoji
        const KUMI = 0x0080;
        /// Unused
        const UNUSED3 = 0x0100;
        /// Emboss
        const EMBOSS = 0x0200;
        /// Style index present
        const STYLE_INDEX = 0x0800;
        /// Has scheme color
        const HAS_SCHEME_COLOR = 0x1000;
        /// Has shadow color
        const HAS_SHADOW_COLOR = 0x2000;
    }
}

// Keep module for backward compatibility
pub mod cf_mask {
    pub const BOLD: u16 = super::CharacterMask::BOLD.bits();
    pub const ITALIC: u16 = super::CharacterMask::ITALIC.bits();
    pub const UNDERLINE: u16 = super::CharacterMask::UNDERLINE.bits();
    pub const UNUSED1: u16 = super::CharacterMask::UNUSED1.bits();
    pub const SHADOW: u16 = super::CharacterMask::SHADOW.bits();
    pub const FEHINT: u16 = super::CharacterMask::FEHINT.bits();
    pub const UNUSED2: u16 = super::CharacterMask::UNUSED2.bits();
    pub const KUMI: u16 = super::CharacterMask::KUMI.bits();
    pub const UNUSED3: u16 = super::CharacterMask::UNUSED3.bits();
    pub const EMBOSS: u16 = super::CharacterMask::EMBOSS.bits();
    pub const STYLE_INDEX: u16 = super::CharacterMask::STYLE_INDEX.bits();
    pub const HAS_SCHEME_COLOR: u16 = super::CharacterMask::HAS_SCHEME_COLOR.bits();
    pub const HAS_SHADOW_COLOR: u16 = super::CharacterMask::HAS_SHADOW_COLOR.bits();
}

// =============================================================================
// Font sizes in 100ths of a point
// =============================================================================

/// Font sizes (in 100ths of a point)
pub mod font_size {
    pub const PT_44: u16 = 4400; // 0x1130 (reversed: 0x3011 in some formats)
    pub const PT_32: u16 = 3200; // 0x0C80
    pub const PT_28: u16 = 2800;
    pub const PT_24: u16 = 2400;
    pub const PT_20: u16 = 2000;
    pub const PT_18: u16 = 1800;
    pub const PT_16: u16 = 1600;
    pub const PT_14: u16 = 1400;
    pub const PT_12: u16 = 1200; // 0x04B0
}

// =============================================================================
// Indent levels (in master units, 1/576 inch = 12.5 EMUs)
// =============================================================================

/// Indent level spacing (in master units)
pub mod indent {
    pub const LEVEL_0: u16 = 0x0000;
    pub const LEVEL_1: u16 = 0x0120; // 288 = 0.5 inch
    pub const LEVEL_2: u16 = 0x0240; // 576 = 1.0 inch
    pub const LEVEL_3: u16 = 0x0360; // 864 = 1.5 inch
    pub const LEVEL_4: u16 = 0x0480; // 1152 = 2.0 inch
}

// =============================================================================
// Bullet Formatting Constants
// =============================================================================

/// Bullet formatting constants
pub mod bullet {
    /// Default bullet flags (has bullet, autobullet)
    pub const FLAGS_DEFAULT: u16 = 0x2022;
    /// Font index for default bullet
    pub const FONT_INDEX: u16 = 0x0064;
    /// No bullet character
    pub const CHAR_NONE: u16 = 0x0000;
    /// Bullet size (percentage or undefined)
    pub const SIZE_DEFAULT: u16 = 0x0000;
    /// Scheme color with bullet color flag
    pub const COLOR_SCHEME: u32 = 0x0001_FF00;
    /// Color with alpha channel
    pub const COLOR_ALPHA: u32 = 0x0000_FF00;
}

/// Alignment values
pub mod align {
    pub const LEFT: u16 = 0x0000;
    pub const CENTER: u16 = 0x0001;
    pub const RIGHT: u16 = 0x0002;
    pub const JUSTIFY: u16 = 0x0003;
    pub const DEFAULT: u16 = 0x0064; // POI default
}

/// Line spacing values (in percentage or absolute)
pub mod spacing {
    pub const DEFAULT_LINE: u16 = 0x0000;
    pub const LINE_120_PCT: u16 = 0x0014; // 20 = 1.2x
    pub const LINE_150_PCT: u16 = 0x001E; // 30 = 1.5x
    pub const SPACE_AFTER_216: u16 = 0x00D8; // 216 = body spacing
}

/// Tab size values
pub mod tab {
    pub const DEFAULT_SIZE: u16 = 0x0240; // 576 = 1 inch
}

/// Font display flags
pub mod font_flags {
    pub const NONE: u16 = 0x0000;
    pub const INHERIT: u16 = 0xFFFF;
}

/// Position values
pub mod position {
    pub const TITLE_DEFAULT: u16 = 0x002C; // 44
    pub const BODY_DEFAULT: u16 = 0x0020; // 32
    pub const NOTES_DEFAULT: u16 = 0x000C; // 12
    pub const OTHER_DEFAULT: u16 = 0x0012; // 18
}

// =============================================================================
// Zerocopy Structs for Text Formatting
// =============================================================================

/// Simple level entry for minimal styles (8 bytes)
#[derive(Debug, Clone, Copy, IntoBytes, Immutable, KnownLayout)]
#[repr(C, packed)]
pub struct SimpleLevelEntry {
    /// PF mask (usually 0 or simple flags)
    pub pf_mask: u32,
    /// CF mask (usually 0 or font size only)
    pub cf_mask: u16,
    /// Font size (if cf_mask includes font size)
    pub font_size: u16,
}

impl SimpleLevelEntry {
    pub const fn new(pf_mask: u32, cf_mask: u16, font_size: u16) -> Self {
        Self {
            pf_mask,
            cf_mask,
            font_size,
        }
    }

    /// Empty entry with no formatting
    pub const EMPTY: Self = Self {
        pf_mask: 0,
        cf_mask: 0,
        font_size: 0,
    };

    /// Entry with only font size
    pub const fn with_font_size(font_size: u16) -> Self {
        Self {
            pf_mask: 0,
            cf_mask: cf_mask::BOLD,
            font_size,
        }
    }
}

/// Indented level entry (12 bytes)
#[derive(Debug, Clone, Copy, IntoBytes, Immutable, KnownLayout)]
#[repr(C, packed)]
pub struct IndentedLevelEntry {
    /// PF mask with indent flags
    pub pf_mask: u32,
    /// Left margin
    pub left_margin: u16,
    /// Indent
    pub indent: u16,
    /// CF mask (usually 0)
    pub cf_mask: u16,
    /// Padding/unused
    pub cf_flags: u16,
}

impl IndentedLevelEntry {
    pub const fn new(left_margin: u16, indent: u16) -> Self {
        Self {
            pf_mask: pf_mask::LEFT_MARGIN | pf_mask::INDENT,
            left_margin,
            indent,
            cf_mask: 0,
            cf_flags: 0,
        }
    }
}

/// Full style level (Title, Body, Notes, Other).
#[derive(Debug, Clone)]
pub struct FullStyleLevel {
    // Paragraph formatting
    pub pf_mask: u32,
    pub bullet_flags: u16,
    pub bullet_char: u16,
    pub bullet_font: u16,
    pub bullet_size: u16,
    pub bullet_color: u32,
    pub align: u16,
    pub line_spacing: u16,
    pub space_before: u16,
    pub space_after: u16,
    pub left_margin: u16,
    pub indent: u16,
    pub default_tab_size: u16,
    // Character formatting
    pub cf_mask: u32,
    pub cf_flags: u16,
    pub font_index: u16,
    pub font_size: u16,
    pub font_color: u32,
    pub position: u16,
    // Optional field presence
    pub has_font_size: bool,
    pub has_font_color: bool,
    pub has_position: bool,
    pub has_font_index: bool,
}

/// Simple style level (CenterBody, HalfBody, QuarterBody, CenterTitle).
#[derive(Debug, Clone)]
pub struct SimpleStyleLevel {
    pub pf_mask: u32,
    pub left_margin: u16,
    pub indent: u16,
    pub cf_mask: u32,
    pub cf_flags: u16,
    pub font_size: u16,
    pub has_font_size: bool,
}
