//! TxMasterStyleAtom builder (MS-PPT 2.9.45)
//!
//! Constructs text master style atoms with proper formatting structures
//! using zerocopy for binary serialization.

use bitflags::bitflags;
use zerocopy_derive::{FromBytes, Immutable, IntoBytes, KnownLayout};

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
#[derive(Debug, Clone, Copy, FromBytes, IntoBytes, Immutable, KnownLayout)]
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
#[derive(Debug, Clone, Copy, FromBytes, IntoBytes, Immutable, KnownLayout)]
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

// =============================================================================
// TxMasterStyleAtom Builder
// =============================================================================

/// Builder for TxMasterStyleAtom (MS-PPT 2.9.45)
pub struct TxMasterStyleBuilder {
    data: Vec<u8>,
}

impl TxMasterStyleBuilder {
    /// Create a new builder with the specified number of indent levels
    pub fn new(levels: u16) -> Self {
        let mut data = Vec::new();
        data.extend_from_slice(&levels.to_le_bytes());
        Self { data }
    }

    /// Add a full style level (used by Title, Body, Notes, Other styles)
    pub fn add_full_level(&mut self, level: &FullStyleLevel) {
        // TextPFException
        self.data.extend_from_slice(&level.pf_mask.to_le_bytes());
        if level.pf_mask & pf_mask::BULLET_CHAR != 0 {
            self.data
                .extend_from_slice(&level.bullet_flags.to_le_bytes());
        }
        if level.pf_mask & pf_mask::BULLET_CHAR != 0 {
            self.data
                .extend_from_slice(&level.bullet_char.to_le_bytes());
        }
        if level.pf_mask & pf_mask::BULLET_FONT != 0 {
            self.data
                .extend_from_slice(&level.bullet_font.to_le_bytes());
        }
        if level.pf_mask & pf_mask::BULLET_SIZE != 0 {
            self.data
                .extend_from_slice(&level.bullet_size.to_le_bytes());
        }
        if level.pf_mask & pf_mask::BULLET_COLOR != 0 {
            self.data
                .extend_from_slice(&level.bullet_color.to_le_bytes());
        }
        if level.pf_mask & pf_mask::ALIGN != 0 {
            self.data.extend_from_slice(&level.align.to_le_bytes());
        }
        if level.pf_mask & pf_mask::LINE_SPACING != 0 {
            self.data
                .extend_from_slice(&level.line_spacing.to_le_bytes());
        }
        if level.pf_mask & pf_mask::SPACE_BEFORE != 0 {
            self.data
                .extend_from_slice(&level.space_before.to_le_bytes());
        }
        if level.pf_mask & pf_mask::SPACE_AFTER != 0 {
            self.data
                .extend_from_slice(&level.space_after.to_le_bytes());
        }
        if level.pf_mask & pf_mask::LEFT_MARGIN != 0 {
            self.data
                .extend_from_slice(&level.left_margin.to_le_bytes());
        }
        if level.pf_mask & pf_mask::INDENT != 0 {
            self.data.extend_from_slice(&level.indent.to_le_bytes());
        }
        if level.pf_mask & pf_mask::DEFAULT_TAB_SIZE != 0 {
            self.data
                .extend_from_slice(&level.default_tab_size.to_le_bytes());
        }

        // TextCFException
        self.data.extend_from_slice(&level.cf_mask.to_le_bytes());
        if level.cf_mask != 0 {
            self.data.extend_from_slice(&level.cf_flags.to_le_bytes());
        }
        if level.cf_mask & cf_mask::STYLE_INDEX != 0 {
            self.data.extend_from_slice(&level.font_index.to_le_bytes());
        }
        if level.has_font_size {
            self.data.extend_from_slice(&level.font_size.to_le_bytes());
        }
        if level.has_font_color {
            self.data.extend_from_slice(&level.font_color.to_le_bytes());
        }
        if level.has_position {
            self.data.extend_from_slice(&level.position.to_le_bytes());
        }
    }

    /// Add a simple style level (used by CenterBody, HalfBody, QuarterBody)
    pub fn add_simple_level(&mut self, level: &SimpleStyleLevel) {
        // Minimal TextPFException: just mask + indent
        self.data.extend_from_slice(&level.pf_mask.to_le_bytes());
        if level.pf_mask & pf_mask::LEFT_MARGIN != 0 {
            self.data
                .extend_from_slice(&level.left_margin.to_le_bytes());
        }
        if level.pf_mask & pf_mask::INDENT != 0 {
            self.data.extend_from_slice(&level.indent.to_le_bytes());
        }

        // Minimal TextCFException
        self.data.extend_from_slice(&level.cf_mask.to_le_bytes());
        if level.cf_mask != 0 {
            self.data.extend_from_slice(&level.cf_flags.to_le_bytes());
        }
        if level.has_font_size {
            self.data.extend_from_slice(&level.font_size.to_le_bytes());
        }
    }

    /// Build the final byte array
    pub fn build(self) -> Vec<u8> {
        self.data
    }
}

/// Full style level (Title, Body, Notes, Other)
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
    pub cf_mask: u16,
    pub cf_flags: u16,
    pub font_index: u16,
    pub font_size: u16,
    pub font_color: u32,
    pub position: u16,
    // Optional field presence
    pub has_font_size: bool,
    pub has_font_color: bool,
    pub has_position: bool,
}

/// Simple style level (CenterBody, HalfBody, QuarterBody, CenterTitle)
#[derive(Debug, Clone)]
pub struct SimpleStyleLevel {
    pub pf_mask: u32,
    pub left_margin: u16,
    pub indent: u16,
    pub cf_mask: u16,
    pub cf_flags: u16,
    pub font_size: u16,
    pub has_font_size: bool,
}

// =============================================================================
// Pre-built style constants - constructed programmatically
// =============================================================================

/// Title style PF mask: all bullet and formatting fields
pub const TITLE_PF_MASK: u32 = 0x003F_FDFF;
/// Title/Body CF mask: bold, italic, underline flags
pub const TITLE_CF_MASK: u16 = 0x0007;

/// Build title text master style (instance=0)
pub fn build_tx_master_style_title() -> Vec<u8> {
    let mut data = Vec::with_capacity(64);

    // Level count = 1
    data.extend_from_slice(&1u16.to_le_bytes());

    // TextPFException
    data.extend_from_slice(&TITLE_PF_MASK.to_le_bytes());
    // Bullet flags field (when pf_mask & HAS_BULLET)
    data.extend_from_slice(&[0x00, 0x00]); // bullet has font/color flags
    data.extend_from_slice(&bullet::FLAGS_DEFAULT.to_le_bytes());
    data.extend_from_slice(&bullet::CHAR_NONE.to_le_bytes());
    data.extend_from_slice(&bullet::FONT_INDEX.to_le_bytes());
    data.extend_from_slice(&bullet::SIZE_DEFAULT.to_le_bytes());
    data.extend_from_slice(&bullet::COLOR_SCHEME.to_le_bytes());
    data.extend_from_slice(&align::DEFAULT.to_le_bytes());
    data.extend_from_slice(&spacing::DEFAULT_LINE.to_le_bytes());
    data.extend_from_slice(&spacing::DEFAULT_LINE.to_le_bytes()); // space before
    data.extend_from_slice(&spacing::DEFAULT_LINE.to_le_bytes()); // space after
    data.extend_from_slice(&indent::LEVEL_0.to_le_bytes());
    data.extend_from_slice(&indent::LEVEL_0.to_le_bytes());
    data.extend_from_slice(&tab::DEFAULT_SIZE.to_le_bytes());

    // TextCFException
    data.extend_from_slice(&[0x00, 0x00]); // extra padding before cf_mask
    data.extend_from_slice(&TITLE_CF_MASK.to_le_bytes());
    data.extend_from_slice(&font_flags::NONE.to_le_bytes());
    data.extend_from_slice(&font_flags::INHERIT.to_le_bytes()); // font index
    // Additional CF fields (font size byte, color, position)
    data.extend_from_slice(&[0xEF, 0x00, 0x00, 0x00, 0x00, 0x00]); // font size byte pattern
    data.extend_from_slice(&0xFFFF_FFFFu32.to_le_bytes()); // font color
    data.extend_from_slice(&font_flags::INHERIT.to_le_bytes()); // inherited
    data.extend_from_slice(&position::TITLE_DEFAULT.to_le_bytes());
    // Trailing fields
    data.extend_from_slice(&[0x00, 0x00, 0x00, 0x03, 0x00, 0x00]);

    data
}

/// Body text PF mask for indented levels
pub const BODY_LEVEL_PF_MASK: u16 = 0x0580;

/// Bullet configuration for each body level
pub mod body_bullet {
    pub const LEVEL_1: (u16, u16) = (0x2013, 0x01D4); // dash bullet
    pub const LEVEL_2: (u16, u16) = (0x2022, 0x02D0); // round bullet
    pub const LEVEL_3: (u16, u16) = (0x2013, 0x03F0); // dash bullet
    pub const LEVEL_4: (u16, u16) = (0x00BB, 0x0510); // right angle bullet
}

/// Build body text master style (instance=1)
pub fn build_tx_master_style_body() -> Vec<u8> {
    let mut data = Vec::with_capacity(128);

    // 5 levels
    data.extend_from_slice(&5u16.to_le_bytes());

    // Level 0 (similar to title, with body-specific spacing)
    data.extend_from_slice(&TITLE_PF_MASK.to_le_bytes());
    data.extend_from_slice(&[0x01, 0x00]); // bullet flags variant
    data.extend_from_slice(&bullet::FLAGS_DEFAULT.to_le_bytes());
    data.extend_from_slice(&bullet::CHAR_NONE.to_le_bytes());
    data.extend_from_slice(&bullet::FONT_INDEX.to_le_bytes());
    data.extend_from_slice(&bullet::SIZE_DEFAULT.to_le_bytes());
    data.extend_from_slice(&bullet::COLOR_ALPHA.to_le_bytes());
    data.extend_from_slice(&align::DEFAULT.to_le_bytes());
    data.extend_from_slice(&spacing::LINE_120_PCT.to_le_bytes());
    data.extend_from_slice(&spacing::DEFAULT_LINE.to_le_bytes());
    data.extend_from_slice(&spacing::SPACE_AFTER_216.to_le_bytes());
    data.extend_from_slice(&indent::LEVEL_0.to_le_bytes());
    data.extend_from_slice(&indent::LEVEL_0.to_le_bytes());
    data.extend_from_slice(&tab::DEFAULT_SIZE.to_le_bytes());
    data.extend_from_slice(&[0x00, 0x00]); // padding
    data.extend_from_slice(&TITLE_CF_MASK.to_le_bytes());
    data.extend_from_slice(&font_flags::NONE.to_le_bytes());
    data.extend_from_slice(&font_flags::INHERIT.to_le_bytes());
    data.extend_from_slice(&[0xEF, 0x00, 0x00, 0x00, 0x00, 0x00]);
    data.extend_from_slice(&0xFFFF_FFFFu32.to_le_bytes());
    data.extend_from_slice(&font_flags::INHERIT.to_le_bytes());
    data.extend_from_slice(&position::BODY_DEFAULT.to_le_bytes());
    data.extend_from_slice(&[0x00, 0x00, 0x00, 0x01, 0x00, 0x00]);

    // Levels 1-4 (progressive indentation with varying bullets)
    let level_data: [(u16, u16, (u16, u16)); 4] = [
        (indent::LEVEL_1, font_size::PT_28, body_bullet::LEVEL_1),
        (indent::LEVEL_2, font_size::PT_24, body_bullet::LEVEL_2),
        (indent::LEVEL_3, font_size::PT_20, body_bullet::LEVEL_3),
        (indent::LEVEL_4, font_size::PT_18, body_bullet::LEVEL_4),
    ];

    for (i, (left_margin, font_sz, (bullet_flags, bullet_char))) in level_data.iter().enumerate() {
        data.extend_from_slice(&BODY_LEVEL_PF_MASK.to_le_bytes());
        data.extend_from_slice(&[0x00, 0x00]);
        data.extend_from_slice(&bullet_flags.to_le_bytes());
        data.extend_from_slice(&bullet_char.to_le_bytes());
        data.extend_from_slice(&left_margin.to_le_bytes());
        data.extend_from_slice(&[0x00, 0x00]); // cf padding
        if i < 3 {
            data.extend_from_slice(&cf_mask::BOLD.to_le_bytes());
            data.extend_from_slice(&font_sz.to_le_bytes());
        } else {
            data.extend_from_slice(&[0x00, 0x00]);
        }
    }

    data
}

/// PF mask for indent-only levels
pub const INDENT_ONLY_PF_MASK: u32 = pf_mask::LEFT_MARGIN | pf_mask::INDENT;

/// Build notes text master style (instance=2)
pub fn build_tx_master_style_notes() -> Vec<u8> {
    let mut data = Vec::with_capacity(112);

    // 5 levels
    data.extend_from_slice(&5u16.to_le_bytes());

    // Level 0 (similar to title, with notes-specific line spacing)
    data.extend_from_slice(&TITLE_PF_MASK.to_le_bytes());
    data.extend_from_slice(&[0x00, 0x00]);
    data.extend_from_slice(&bullet::FLAGS_DEFAULT.to_le_bytes());
    data.extend_from_slice(&bullet::CHAR_NONE.to_le_bytes());
    data.extend_from_slice(&bullet::FONT_INDEX.to_le_bytes());
    data.extend_from_slice(&bullet::SIZE_DEFAULT.to_le_bytes());
    data.extend_from_slice(&bullet::COLOR_ALPHA.to_le_bytes());
    data.extend_from_slice(&align::DEFAULT.to_le_bytes());
    data.extend_from_slice(&spacing::LINE_150_PCT.to_le_bytes()); // 1.5x line spacing
    data.extend_from_slice(&spacing::DEFAULT_LINE.to_le_bytes());
    data.extend_from_slice(&spacing::DEFAULT_LINE.to_le_bytes());
    data.extend_from_slice(&indent::LEVEL_0.to_le_bytes());
    data.extend_from_slice(&indent::LEVEL_0.to_le_bytes());
    data.extend_from_slice(&tab::DEFAULT_SIZE.to_le_bytes());
    data.extend_from_slice(&[0x00, 0x00]);
    data.extend_from_slice(&TITLE_CF_MASK.to_le_bytes());
    data.extend_from_slice(&font_flags::NONE.to_le_bytes());
    data.extend_from_slice(&font_flags::INHERIT.to_le_bytes());
    data.extend_from_slice(&[0xEF, 0x00, 0x00, 0x00, 0x00, 0x00]);
    data.extend_from_slice(&0xFFFF_FFFFu32.to_le_bytes());
    data.extend_from_slice(&font_flags::INHERIT.to_le_bytes());
    data.extend_from_slice(&position::NOTES_DEFAULT.to_le_bytes());
    data.extend_from_slice(&[0x00, 0x00, 0x00, 0x01, 0x00, 0x00]);

    // Levels 1-4 (simple indent progression)
    for i in 1..=4u16 {
        let left_margin = i * indent::LEVEL_1;
        data.extend_from_slice(&INDENT_ONLY_PF_MASK.to_le_bytes());
        data.extend_from_slice(&left_margin.to_le_bytes());
        data.extend_from_slice(&left_margin.to_le_bytes());
        data.extend_from_slice(&[0x00, 0x00, 0x00, 0x00, 0x00, 0x00]); // empty CF
    }

    data
}

/// Build other text master style (instance=4)
pub fn build_tx_master_style_other() -> Vec<u8> {
    let mut data = Vec::with_capacity(112);

    // 5 levels
    data.extend_from_slice(&5u16.to_le_bytes());

    // Level 0 (similar to title, no line spacing)
    data.extend_from_slice(&TITLE_PF_MASK.to_le_bytes());
    data.extend_from_slice(&[0x00, 0x00]);
    data.extend_from_slice(&bullet::FLAGS_DEFAULT.to_le_bytes());
    data.extend_from_slice(&bullet::CHAR_NONE.to_le_bytes());
    data.extend_from_slice(&bullet::FONT_INDEX.to_le_bytes());
    data.extend_from_slice(&bullet::SIZE_DEFAULT.to_le_bytes());
    data.extend_from_slice(&bullet::COLOR_ALPHA.to_le_bytes());
    data.extend_from_slice(&align::DEFAULT.to_le_bytes());
    data.extend_from_slice(&spacing::DEFAULT_LINE.to_le_bytes());
    data.extend_from_slice(&spacing::DEFAULT_LINE.to_le_bytes());
    data.extend_from_slice(&spacing::DEFAULT_LINE.to_le_bytes());
    data.extend_from_slice(&indent::LEVEL_0.to_le_bytes());
    data.extend_from_slice(&indent::LEVEL_0.to_le_bytes());
    data.extend_from_slice(&tab::DEFAULT_SIZE.to_le_bytes());
    data.extend_from_slice(&[0x00, 0x00]);
    data.extend_from_slice(&TITLE_CF_MASK.to_le_bytes());
    data.extend_from_slice(&font_flags::NONE.to_le_bytes());
    data.extend_from_slice(&font_flags::INHERIT.to_le_bytes());
    data.extend_from_slice(&[0xEF, 0x00, 0x00, 0x00, 0x00, 0x00]);
    data.extend_from_slice(&0xFFFF_FFFFu32.to_le_bytes());
    data.extend_from_slice(&font_flags::INHERIT.to_le_bytes());
    data.extend_from_slice(&position::OTHER_DEFAULT.to_le_bytes());
    data.extend_from_slice(&[0x00, 0x00, 0x00, 0x01, 0x00, 0x00]);

    // Levels 1-4
    for i in 1..=4u16 {
        let left_margin = i * indent::LEVEL_1;
        data.extend_from_slice(&INDENT_ONLY_PF_MASK.to_le_bytes());
        data.extend_from_slice(&left_margin.to_le_bytes());
        data.extend_from_slice(&left_margin.to_le_bytes());
        data.extend_from_slice(&[0x00, 0x00, 0x00, 0x00, 0x00, 0x00]); // empty CF
    }

    data
}

/// Center body PF mask with alignment
pub const CENTER_BODY_PF_MASK: u16 = 0x0901; // align center + bullet

/// Build center body text master style (instance=5)
pub fn build_tx_master_style_center_body() -> Vec<u8> {
    let mut data = Vec::with_capacity(84);
    data.extend_from_slice(&5u16.to_le_bytes());

    for i in 0..5u16 {
        let left_margin = i * indent::LEVEL_1;
        // TextPFException with center alignment
        data.extend_from_slice(&[0x00, 0x00]); // pf_mask low
        if i == 0 {
            data.extend_from_slice(&CENTER_BODY_PF_MASK.to_le_bytes());
            data.extend_from_slice(&[0x00, 0x00, 0x00, 0x00]); // bullet info
            data.extend_from_slice(&align::CENTER.to_le_bytes());
        } else {
            data.extend_from_slice(&(i - 1).to_le_bytes());
            data.extend_from_slice(&[0x00]); // extra byte
            data.extend_from_slice(&CENTER_BODY_PF_MASK.to_le_bytes());
            data.extend_from_slice(&[0x00, 0x00, 0x00, 0x00]);
            data.extend_from_slice(&align::CENTER.to_le_bytes());
        }
        data.extend_from_slice(&left_margin.to_le_bytes());
        data.extend_from_slice(&[0x00, 0x00, 0x00, 0x00]); // empty CF
    }

    data
}

/// Build center title text master style (instance=6)
pub fn build_tx_master_style_center_title() -> Vec<u8> {
    let mut data = Vec::with_capacity(12);
    data.extend_from_slice(&1u16.to_le_bytes()); // 1 level
    // Minimal formatting: empty PF and CF
    data.extend_from_slice(&0u32.to_le_bytes()); // pf_mask = 0
    data.extend_from_slice(&0u32.to_le_bytes()); // cf_mask = 0, padding
    data.extend_from_slice(&[0x00, 0x00]); // trailing
    data
}

/// Half body font sizes (in 100ths of a point)
pub const HALF_BODY_FONT_SIZES: [u16; 5] = [
    font_size::PT_28,
    font_size::PT_24,
    font_size::PT_20,
    font_size::PT_18,
    font_size::PT_18,
];

/// Quarter body font sizes (in 100ths of a point)
pub const QUARTER_BODY_FONT_SIZES: [u16; 5] = [
    font_size::PT_24,
    font_size::PT_20,
    font_size::PT_18,
    font_size::PT_16,
    font_size::PT_16,
];

/// Build half body text master style (instance=7)
pub fn build_tx_master_style_half_body() -> Vec<u8> {
    let mut data = Vec::with_capacity(64);
    data.extend_from_slice(&5u16.to_le_bytes());

    for (i, &font_sz) in HALF_BODY_FONT_SIZES.iter().enumerate() {
        // Empty PF (8 bytes of zeros)
        data.extend_from_slice(&0u32.to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes());
        // CF with font size only
        data.extend_from_slice(&cf_mask::BOLD.to_le_bytes()); // 0x0002
        data.extend_from_slice(&font_sz.to_le_bytes());
        if i < 4 {
            data.extend_from_slice(&((i + 1) as u16).to_le_bytes());
            data.extend_from_slice(&0u32.to_le_bytes());
            data.extend_from_slice(&[0x00, 0x00]);
        }
    }

    data
}

/// Build quarter body text master style (instance=8)
pub fn build_tx_master_style_quarter_body() -> Vec<u8> {
    let mut data = Vec::with_capacity(64);
    data.extend_from_slice(&5u16.to_le_bytes());

    for (i, &font_sz) in QUARTER_BODY_FONT_SIZES.iter().enumerate() {
        // Empty PF (8 bytes of zeros)
        data.extend_from_slice(&0u32.to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes());
        // CF with font size only
        data.extend_from_slice(&cf_mask::BOLD.to_le_bytes()); // 0x0002
        data.extend_from_slice(&font_sz.to_le_bytes());
        if i < 4 {
            data.extend_from_slice(&((i + 1) as u16).to_le_bytes());
            data.extend_from_slice(&0u32.to_le_bytes());
            data.extend_from_slice(&[0x00, 0x00]);
        }
    }

    data
}

// =============================================================================
// Lazy-initialized constants using functions
// =============================================================================

/// Get title text master style bytes
pub fn tx_master_style_title() -> Vec<u8> {
    build_tx_master_style_title()
}

/// Get body text master style bytes
pub fn tx_master_style_body() -> Vec<u8> {
    build_tx_master_style_body()
}

/// Get notes text master style bytes
pub fn tx_master_style_notes() -> Vec<u8> {
    build_tx_master_style_notes()
}

/// Get other text master style bytes
pub fn tx_master_style_other() -> Vec<u8> {
    build_tx_master_style_other()
}

/// Get center body text master style bytes
pub fn tx_master_style_center_body() -> Vec<u8> {
    build_tx_master_style_center_body()
}

/// Get center title text master style bytes
pub fn tx_master_style_center_title() -> Vec<u8> {
    build_tx_master_style_center_title()
}

/// Get half body text master style bytes
pub fn tx_master_style_half_body() -> Vec<u8> {
    build_tx_master_style_half_body()
}

/// Get quarter body text master style bytes
pub fn tx_master_style_quarter_body() -> Vec<u8> {
    build_tx_master_style_quarter_body()
}

// =============================================================================
// Backward compatibility - static arrays that match POI exactly
// =============================================================================

/// Title text master style (instance=0) - 62 bytes from POI
pub const TX_MASTER_STYLE_TITLE: [u8; 62] = [
    0x01, 0x00, 0xFF, 0xFD, 0x3F, 0x00, 0x00, 0x00, 0x22, 0x20, 0x00, 0x00, 0x64, 0x00, 0x00, 0x00,
    0x00, 0xFF, 0x01, 0x00, 0x64, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x40, 0x02,
    0x00, 0x00, 0x00, 0x00, 0x07, 0x00, 0x00, 0x00, 0xFF, 0xFF, 0xEF, 0x00, 0x00, 0x00, 0x00, 0x00,
    0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0x2C, 0x00, 0x00, 0x00, 0x00, 0x03, 0x00, 0x00,
];

/// Body text master style (instance=1) - 124 bytes from POI
pub const TX_MASTER_STYLE_BODY: [u8; 124] = [
    0x05, 0x00, 0xFF, 0xFD, 0x3F, 0x00, 0x01, 0x00, 0x22, 0x20, 0x00, 0x00, 0x64, 0x00, 0x00, 0x00,
    0x00, 0xFF, 0x00, 0x00, 0x64, 0x00, 0x14, 0x00, 0x00, 0x00, 0xD8, 0x00, 0x00, 0x00, 0x40, 0x02,
    0x00, 0x00, 0x00, 0x00, 0x07, 0x00, 0x00, 0x00, 0xFF, 0xFF, 0xEF, 0x00, 0x00, 0x00, 0x00, 0x00,
    0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0x20, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x80, 0x05,
    0x00, 0x00, 0x13, 0x20, 0xD4, 0x01, 0x20, 0x01, 0x00, 0x00, 0x02, 0x00, 0x1C, 0x00, 0x80, 0x05,
    0x00, 0x00, 0x22, 0x20, 0xD0, 0x02, 0x40, 0x02, 0x00, 0x00, 0x02, 0x00, 0x18, 0x00, 0x80, 0x05,
    0x00, 0x00, 0x13, 0x20, 0xF0, 0x03, 0x60, 0x03, 0x00, 0x00, 0x02, 0x00, 0x14, 0x00, 0x80, 0x05,
    0x00, 0x00, 0xBB, 0x00, 0x10, 0x05, 0x80, 0x04, 0x00, 0x00, 0x00, 0x00,
];

/// Notes text master style (instance=2) - 110 bytes from POI
pub const TX_MASTER_STYLE_NOTES: [u8; 110] = [
    0x05, 0x00, 0xFF, 0xFD, 0x3F, 0x00, 0x00, 0x00, 0x22, 0x20, 0x00, 0x00, 0x64, 0x00, 0x00, 0x00,
    0x00, 0xFF, 0x00, 0x00, 0x64, 0x00, 0x1E, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x40, 0x02,
    0x00, 0x00, 0x00, 0x00, 0x07, 0x00, 0x00, 0x00, 0xFF, 0xFF, 0xEF, 0x00, 0x00, 0x00, 0x00, 0x00,
    0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0x0C, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x05,
    0x00, 0x00, 0x20, 0x01, 0x20, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x05, 0x00, 0x00, 0x40, 0x02,
    0x40, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00, 0x05, 0x00, 0x00, 0x60, 0x03, 0x60, 0x03, 0x00, 0x00,
    0x00, 0x00, 0x00, 0x05, 0x00, 0x00, 0x80, 0x04, 0x80, 0x04, 0x00, 0x00, 0x00, 0x00,
];

/// Other text master style (instance=4) - 110 bytes from POI
pub const TX_MASTER_STYLE_OTHER: [u8; 110] = [
    0x05, 0x00, 0xFF, 0xFD, 0x3F, 0x00, 0x00, 0x00, 0x22, 0x20, 0x00, 0x00, 0x64, 0x00, 0x00, 0x00,
    0x00, 0xFF, 0x00, 0x00, 0x64, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x40, 0x02,
    0x00, 0x00, 0x00, 0x00, 0x07, 0x00, 0x00, 0x00, 0xFF, 0xFF, 0xEF, 0x00, 0x00, 0x00, 0x00, 0x00,
    0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0x12, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x05,
    0x00, 0x00, 0x20, 0x01, 0x20, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x05, 0x00, 0x00, 0x40, 0x02,
    0x40, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00, 0x05, 0x00, 0x00, 0x60, 0x03, 0x60, 0x03, 0x00, 0x00,
    0x00, 0x00, 0x00, 0x05, 0x00, 0x00, 0x80, 0x04, 0x80, 0x04, 0x00, 0x00, 0x00, 0x00,
];

/// Center body text master style (instance=5) - 82 bytes from POI
pub const TX_MASTER_STYLE_CENTER_BODY: [u8; 82] = [
    0x05, 0x00, 0x00, 0x00, 0x01, 0x09, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x01, 0x00, 0x01, 0x09, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x20, 0x01, 0x00, 0x00,
    0x00, 0x00, 0x02, 0x00, 0x01, 0x09, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x40, 0x02, 0x00, 0x00,
    0x00, 0x00, 0x03, 0x00, 0x01, 0x09, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x60, 0x03, 0x00, 0x00,
    0x00, 0x00, 0x04, 0x00, 0x01, 0x09, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x80, 0x04, 0x00, 0x00,
    0x00, 0x00,
];

/// Center title text master style (instance=6) - 12 bytes from POI
pub const TX_MASTER_STYLE_CENTER_TITLE: [u8; 12] = [
    0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
];

/// Half body text master style (instance=7) - 62 bytes from POI
pub const TX_MASTER_STYLE_HALF_BODY: [u8; 62] = [
    0x05, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x02, 0x00, 0x1C, 0x00, 0x01, 0x00,
    0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x02, 0x00, 0x18, 0x00, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x02, 0x00, 0x14, 0x00, 0x03, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x02, 0x00,
    0x12, 0x00, 0x04, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x02, 0x00, 0x12, 0x00,
];

/// Quarter body text master style (instance=8) - 62 bytes from POI
pub const TX_MASTER_STYLE_QUARTER_BODY: [u8; 62] = [
    0x05, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x02, 0x00, 0x18, 0x00, 0x01, 0x00,
    0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x02, 0x00, 0x14, 0x00, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x02, 0x00, 0x12, 0x00, 0x03, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x02, 0x00,
    0x10, 0x00, 0x04, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x02, 0x00, 0x10, 0x00,
];

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_build_tx_master_style_title() {
        let built = build_tx_master_style_title();
        assert_eq!(built.as_slice(), &TX_MASTER_STYLE_TITLE);
    }

    #[test]
    fn test_build_tx_master_style_center_title() {
        let built = build_tx_master_style_center_title();
        assert_eq!(built.as_slice(), &TX_MASTER_STYLE_CENTER_TITLE);
    }
}
