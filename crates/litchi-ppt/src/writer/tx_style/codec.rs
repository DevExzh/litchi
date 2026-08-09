/// Bullet configuration for each body level
pub mod body_bullet {
    pub const LEVEL_1: (u16, u16) = (0x2013, 0x01D4); // dash bullet
    pub const LEVEL_2: (u16, u16) = (0x2022, 0x02D0); // round bullet
    pub const LEVEL_3: (u16, u16) = (0x2013, 0x03F0); // dash bullet
    pub const LEVEL_4: (u16, u16) = (0x00BB, 0x0510); // right angle bullet
}

use super::semantic::{
    FullStyleLevel, SimpleStyleLevel, align, bullet, cf_mask, font_flags, font_size, indent,
    pf_mask, position, spacing, tab,
};

// =============================================================================
// Pre-built style constants - constructed programmatically
// =============================================================================

/// Title style PF mask: all bullet and formatting fields
pub const TITLE_PF_MASK: u32 = 0x003F_FDFF;
/// Title/Body CF mask: bold, italic, underline flags
pub const TITLE_CF_MASK: u16 = 0x0007;

/// Body text PF mask for indented levels
pub const BODY_LEVEL_PF_MASK: u16 = 0x0580;

/// PF mask for indent-only levels
pub const INDENT_ONLY_PF_MASK: u32 = pf_mask::LEFT_MARGIN | pf_mask::INDENT;

/// Center body PF mask with alignment
pub const CENTER_BODY_PF_MASK: u16 = 0x0901; // align center + bullet

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

// =============================================================================
// TxMasterStyleAtom Builder
// =============================================================================

/// Builder for `TxMasterStyleAtom` (MS-PPT 2.9.45)
pub struct TxMasterStyleBuilder {
    data: Vec<u8>,
    text_type: u16,
    next_level: u16,
}

impl TxMasterStyleBuilder {
    /// Create a new builder with the specified number of indent levels
    #[must_use]
    pub fn new(levels: u16) -> Self {
        Self::for_text_type(0, levels)
    }

    /// Create a builder for a specific `TextTypeEnum` record instance.
    #[must_use]
    pub fn for_text_type(text_type: u16, levels: u16) -> Self {
        let mut data = Vec::new();
        data.extend_from_slice(&levels.to_le_bytes());
        Self {
            data,
            text_type,
            next_level: 0,
        }
    }

    /// Add a full style level (used by Title, Body, Notes, Other styles)
    pub fn add_full_level(&mut self, level: &FullStyleLevel) {
        if self.text_type >= 5 {
            self.data.extend_from_slice(&self.next_level.to_le_bytes());
        }
        self.next_level = self.next_level.saturating_add(1);
        // TextPFException
        self.data.extend_from_slice(&level.pf_mask.to_le_bytes());
        if level.pf_mask & 0x000F != 0 {
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
        let mut cf_mask = level.cf_mask;
        if level.has_font_index {
            cf_mask |= 0x0001_0000;
        }
        if level.has_font_size {
            cf_mask |= 0x0002_0000;
        }
        if level.has_font_color {
            cf_mask |= 0x0004_0000;
        }
        if level.has_position {
            cf_mask |= 0x0008_0000;
        }
        self.data.extend_from_slice(&cf_mask.to_le_bytes());
        if cf_mask & 0x0000_FFFF != 0 {
            self.data.extend_from_slice(&level.cf_flags.to_le_bytes());
        }
        if level.has_font_index {
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

    /// Add a simple style level (used by `CenterBody`, `HalfBody`, `QuarterBody`)
    pub fn add_simple_level(&mut self, level: &SimpleStyleLevel) {
        if self.text_type >= 5 {
            self.data.extend_from_slice(&self.next_level.to_le_bytes());
        }
        self.next_level = self.next_level.saturating_add(1);
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
        let mut cf_mask = level.cf_mask;
        if level.has_font_size {
            cf_mask |= 0x0002_0000;
        }
        self.data.extend_from_slice(&cf_mask.to_le_bytes());
        if cf_mask & 0x0000_FFFF != 0 {
            self.data.extend_from_slice(&level.cf_flags.to_le_bytes());
        }
        if level.has_font_size {
            self.data.extend_from_slice(&level.font_size.to_le_bytes());
        }
    }

    /// Build the final byte array
    #[must_use]
    pub fn build(self) -> Vec<u8> {
        self.data
    }
}

// =============================================================================
// Legacy builders matching POI byte layouts
// =============================================================================

/// Build title text master style (instance=0)
#[must_use]
pub fn legacy_build_tx_master_style_title() -> Vec<u8> {
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

/// Build body text master style (instance=1)
#[must_use]
pub fn legacy_build_tx_master_style_body() -> Vec<u8> {
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

/// Build notes text master style (instance=2)
#[must_use]
pub fn legacy_build_tx_master_style_notes() -> Vec<u8> {
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
#[must_use]
pub fn legacy_build_tx_master_style_other() -> Vec<u8> {
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

/// Build center body text master style (instance=5)
#[must_use]
pub fn legacy_build_tx_master_style_center_body() -> Vec<u8> {
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
#[must_use]
pub fn legacy_build_tx_master_style_center_title() -> Vec<u8> {
    let mut data = Vec::with_capacity(12);
    data.extend_from_slice(&1u16.to_le_bytes()); // 1 level
    // Minimal formatting: empty PF and CF
    data.extend_from_slice(&0u32.to_le_bytes()); // pf_mask = 0
    data.extend_from_slice(&0u32.to_le_bytes()); // cf_mask = 0, padding
    data.extend_from_slice(&[0x00, 0x00]); // trailing
    data
}

/// Build half body text master style (instance=7)
#[allow(
    clippy::cast_possible_truncation,
    reason = "the loop index is bounded to 0..=4 by the 5-element HALF_BODY_FONT_SIZES array, so `i + 1` always fits in u16"
)]
#[must_use]
pub fn legacy_build_tx_master_style_half_body() -> Vec<u8> {
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
#[allow(
    clippy::cast_possible_truncation,
    reason = "the loop index is bounded to 0..=4 by the 5-element QUARTER_BODY_FONT_SIZES array, so `i + 1` always fits in u16"
)]
#[must_use]
pub fn legacy_build_tx_master_style_quarter_body() -> Vec<u8> {
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
// Spec-conformant builders
// =============================================================================

#[allow(
    clippy::cast_possible_truncation,
    reason = "every caller passes at most 5 font sizes, so both the level count and the level index always fit in u16"
)]
fn build_spec_master_style(
    text_type: u16,
    font_sizes: &[u16],
    alignment: u16,
    bullets: bool,
) -> Vec<u8> {
    let mut builder = TxMasterStyleBuilder::for_text_type(text_type, font_sizes.len() as u16);
    for (level, &font_size) in font_sizes.iter().enumerate() {
        let margin = (level as u16).saturating_mul(indent::LEVEL_1);
        let bullet_mask = if bullets {
            pf_mask::HAS_BULLET
                | pf_mask::BULLET_HAS_FONT
                | pf_mask::BULLET_HAS_COLOR
                | pf_mask::BULLET_HAS_SIZE
                | pf_mask::BULLET_CHAR
                | pf_mask::BULLET_FONT
                | pf_mask::BULLET_SIZE
                | pf_mask::BULLET_COLOR
        } else {
            0
        };
        builder.add_full_level(&FullStyleLevel {
            pf_mask: bullet_mask
                | pf_mask::ALIGN
                | pf_mask::LEFT_MARGIN
                | pf_mask::INDENT
                | pf_mask::DEFAULT_TAB_SIZE,
            bullet_flags: if bullets { 0x000F } else { 0 },
            bullet_char: 0x2022,
            bullet_font: 0,
            bullet_size: 100,
            bullet_color: 0xFF00_0000,
            align: alignment,
            line_spacing: 0,
            space_before: 0,
            space_after: 0,
            left_margin: margin,
            indent: margin,
            default_tab_size: tab::DEFAULT_SIZE,
            cf_mask: 0,
            cf_flags: 0,
            font_index: 0,
            font_size,
            font_color: 0xFF00_0000,
            position: 0,
            has_font_size: true,
            has_font_color: true,
            has_position: false,
            has_font_index: true,
        });
    }
    builder.build()
}

/// Build a spec-conformant title text master style (instance 0).
#[must_use]
pub fn build_tx_master_style_title() -> Vec<u8> {
    build_spec_master_style(0, &[44], align::LEFT, false)
}

/// Build a spec-conformant body text master style (instance 1).
#[must_use]
pub fn build_tx_master_style_body() -> Vec<u8> {
    build_spec_master_style(1, &[32, 28, 24, 20, 18], align::LEFT, true)
}

/// Build a spec-conformant notes text master style (instance 2).
#[must_use]
pub fn build_tx_master_style_notes() -> Vec<u8> {
    build_spec_master_style(2, &[12, 12, 12, 12, 12], align::LEFT, false)
}

/// Build a spec-conformant other-text master style (instance 4).
#[must_use]
pub fn build_tx_master_style_other() -> Vec<u8> {
    build_spec_master_style(4, &[18, 18, 18, 18, 18], align::LEFT, false)
}

/// Build a spec-conformant centered-body master style (instance 5).
#[must_use]
pub fn build_tx_master_style_center_body() -> Vec<u8> {
    build_spec_master_style(5, &[32, 28, 24, 20, 18], align::CENTER, true)
}

/// Build a spec-conformant centered-title master style (instance 6).
#[must_use]
pub fn build_tx_master_style_center_title() -> Vec<u8> {
    build_spec_master_style(6, &[44], align::CENTER, false)
}

/// Build a spec-conformant half-body master style (instance 7).
#[must_use]
pub fn build_tx_master_style_half_body() -> Vec<u8> {
    build_spec_master_style(7, &[28, 24, 20, 18, 18], align::LEFT, true)
}

/// Build a spec-conformant quarter-body master style (instance 8).
#[must_use]
pub fn build_tx_master_style_quarter_body() -> Vec<u8> {
    build_spec_master_style(8, &[24, 20, 18, 16, 16], align::LEFT, true)
}

// =============================================================================
// Lazy-initialized constants using functions
// =============================================================================

/// Get title text master style bytes
#[must_use]
pub fn tx_master_style_title() -> Vec<u8> {
    build_tx_master_style_title()
}

/// Get body text master style bytes
#[must_use]
pub fn tx_master_style_body() -> Vec<u8> {
    build_tx_master_style_body()
}

/// Get notes text master style bytes
#[must_use]
pub fn tx_master_style_notes() -> Vec<u8> {
    build_tx_master_style_notes()
}

/// Get other text master style bytes
#[must_use]
pub fn tx_master_style_other() -> Vec<u8> {
    build_tx_master_style_other()
}

/// Get center body text master style bytes
#[must_use]
pub fn tx_master_style_center_body() -> Vec<u8> {
    build_tx_master_style_center_body()
}

/// Get center title text master style bytes
#[must_use]
pub fn tx_master_style_center_title() -> Vec<u8> {
    build_tx_master_style_center_title()
}

/// Get half body text master style bytes
#[must_use]
pub fn tx_master_style_half_body() -> Vec<u8> {
    build_tx_master_style_half_body()
}

/// Get quarter body text master style bytes
#[must_use]
pub fn tx_master_style_quarter_body() -> Vec<u8> {
    build_tx_master_style_quarter_body()
}
