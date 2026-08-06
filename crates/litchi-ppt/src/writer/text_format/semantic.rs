//! Semantic text-formatting model for PPT files.
//!
//! This module handles text styling including bold, italic, underline,
//! font size, font colors, and paragraph formatting.
//!
//! Reference: [MS-PPT] Section 2.9 - Text Formatting

use crate::writer::smart_tags::SmartTagIndex;
use zerocopy_derive::{Immutable, IntoBytes, KnownLayout};

// =============================================================================
// Text Property Mask Flags (MS-PPT 2.9.20 TextPFException)
// =============================================================================

/// Paragraph property mask flags
pub mod para_mask {
    /// Bullet presence is valid
    pub const HAS_BULLET: u32 = 0x0001;
    /// Bullet font validity is present in `BulletFlags`
    pub const BULLET_HAS_FONT: u32 = 0x0002;
    /// Bullet color validity is present in `BulletFlags`
    pub const BULLET_HAS_COLOR: u32 = 0x0004;
    /// Bullet size validity is present in `BulletFlags`
    pub const BULLET_HAS_SIZE: u32 = 0x0008;
    /// Bullet font reference exists
    pub const BULLET_FONT: u32 = 0x0010;
    /// Bullet color exists
    pub const BULLET_COLOR: u32 = 0x0020;
    /// Bullet size exists
    pub const BULLET_SIZE: u32 = 0x0040;
    /// Bullet character exists
    pub const BULLET_CHAR: u32 = 0x0080;
    /// Alignment present
    pub const ALIGNMENT: u32 = 0x0800;
    /// Line spacing present
    pub const LINE_SPACING: u32 = 0x1000;
    /// Space before present
    pub const SPACE_BEFORE: u32 = 0x2000;
    /// Space after present
    pub const SPACE_AFTER: u32 = 0x4000;
    /// Left margin present
    pub const LEFT_MARGIN: u32 = 0x0100;
    /// Indent present
    pub const INDENT: u32 = 0x0400;
    /// Default tab size present
    pub const DEFAULT_TAB_SIZE: u32 = 0x8000;
    /// Font alignment present
    pub const FONT_ALIGNMENT: u32 = 0x0001_0000;
    /// East Asian character wrapping present
    pub const CHARACTER_WRAP: u32 = 0x0002_0000;
    /// Word wrapping present
    pub const WORD_WRAP: u32 = 0x0004_0000;
    /// Hanging punctuation present
    pub const OVERFLOW: u32 = 0x0008_0000;
    /// Explicit tab-stop array present
    pub const TAB_STOPS: u32 = 0x0010_0000;
    /// Paragraph direction present
    pub const TEXT_DIRECTION: u32 = 0x0020_0000;
}

/// Character property mask flags (MS-PPT 2.9.21 TextCFException)
pub mod char_mask {
    /// Bold
    pub const BOLD: u32 = 0x0001;
    /// Italic
    pub const ITALIC: u32 = 0x0002;
    /// Underline
    pub const UNDERLINE: u32 = 0x0004;
    /// Shadow
    pub const SHADOW: u32 = 0x0010;
    /// Double-byte input hint
    pub const FE_HINT: u32 = 0x0020;
    /// Kumimoji formatting
    pub const KUMI: u32 = 0x0080;
    /// De facto legacy strikethrough bit (`unused3` in MS-PPT)
    pub const LEGACY_STRIKETHROUGH: u32 = 0x0100;
    /// Emboss
    pub const EMBOSS: u32 = 0x0200;
    /// PowerPoint 9 additional-property run grouping bits
    pub const PP9_RUN_ID: u32 = 0x3C00;
    /// Font reference present
    pub const FONT_REF: u32 = 0x0001_0000;
    /// Font size present
    pub const FONT_SIZE: u32 = 0x0002_0000;
    /// Font color present
    pub const FONT_COLOR: u32 = 0x0004_0000;
    /// Position (superscript/subscript) present
    pub const POSITION: u32 = 0x0008_0000;
    /// East Asian font reference present
    pub const ASIAN_FONT_REF: u32 = 0x0020_0000;
    /// ANSI font reference present
    pub const ANSI_FONT_REF: u32 = 0x0040_0000;
    /// Symbol font reference present
    pub const SYMBOL_FONT_REF: u32 = 0x0080_0000;
}

// =============================================================================
// Text Alignment
// =============================================================================

/// Text alignment values
#[repr(u16)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum TextAlign {
    /// Left aligned
    #[default]
    Left = 0x0000,
    /// Center aligned
    Center = 0x0001,
    /// Right aligned
    Right = 0x0002,
    /// Justified
    Justify = 0x0003,
    /// Distributed
    Distributed = 0x0004,
    /// Thai distributed
    ThaiDistributed = 0x0005,
    /// Justify low
    JustifyLow = 0x0006,
}

/// Font alignment within a paragraph line.
#[repr(u16)]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TextFontAlign {
    /// Place characters on the font baseline.
    Roman = 0,
    /// Hang characters from the top of the line.
    Hanging = 1,
    /// Center characters within the line height.
    Center = 2,
    /// Anchor characters to the bottom of the line.
    UpholdFixed = 3,
}

/// Paragraph text direction.
#[repr(u16)]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TextDirection {
    /// Left-to-right text.
    LeftToRight = 0,
    /// Right-to-left text.
    RightToLeft = 1,
}

/// Alignment at a paragraph tab stop.
#[repr(u16)]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TabAlign {
    /// Left-aligned tab stop.
    Left = 0,
    /// Center-aligned tab stop.
    Center = 1,
    /// Right-aligned tab stop.
    Right = 2,
    /// Decimal-point-aligned tab stop.
    Decimal = 3,
}

/// A paragraph tab stop in PowerPoint master units.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TabStop {
    /// Signed position in master units.
    pub position: i16,
    /// Text alignment at the stop.
    pub alignment: TabAlign,
}

impl TabStop {
    /// Create a tab stop.
    pub const fn new(position: i16, alignment: TabAlign) -> Self {
        Self {
            position,
            alignment,
        }
    }
}

// =============================================================================
// Font Style Flags
// =============================================================================

/// Font style flags
#[derive(Debug, Clone, Copy, Default)]
pub struct FontStyle {
    /// Bold text
    pub bold: bool,
    /// Italic text
    pub italic: bool,
    /// Underlined text
    pub underline: bool,
    /// Shadow effect
    pub shadow: bool,
    /// Embossed effect
    pub emboss: bool,
    /// Double-byte input hint
    pub fe_hint: bool,
    /// Kumimoji formatting for vertical text
    pub kumi: bool,
    /// De facto legacy strikethrough (`unused3` in MS-PPT)
    pub strikethrough: bool,
    /// PowerPoint 9 additional-property run grouping identifier
    pub pp9_run_id: Option<u8>,
    /// Validity bits to emit even when their corresponding value is false
    pub specified_mask: u16,
}

impl FontStyle {
    /// Create bold style
    pub const fn bold() -> Self {
        Self {
            bold: true,
            italic: false,
            underline: false,
            shadow: false,
            emboss: false,
            fe_hint: false,
            kumi: false,
            strikethrough: false,
            pp9_run_id: None,
            specified_mask: char_mask::BOLD as u16,
        }
    }

    /// Create italic style
    pub const fn italic() -> Self {
        Self {
            bold: false,
            italic: true,
            underline: false,
            shadow: false,
            emboss: false,
            fe_hint: false,
            kumi: false,
            strikethrough: false,
            pp9_run_id: None,
            specified_mask: char_mask::ITALIC as u16,
        }
    }

    /// Create bold and italic style
    pub const fn bold_italic() -> Self {
        Self {
            bold: true,
            italic: true,
            underline: false,
            shadow: false,
            emboss: false,
            fe_hint: false,
            kumi: false,
            strikethrough: false,
            pp9_run_id: None,
            specified_mask: (char_mask::BOLD | char_mask::ITALIC) as u16,
        }
    }

    /// Convert to mask value for TextCFException
    pub fn to_mask(&self) -> u32 {
        let mut mask = u32::from(self.specified_mask);
        if self.bold {
            mask |= char_mask::BOLD;
        }
        if self.italic {
            mask |= char_mask::ITALIC;
        }
        if self.underline {
            mask |= char_mask::UNDERLINE;
        }
        if self.shadow {
            mask |= char_mask::SHADOW;
        }
        if self.fe_hint {
            mask |= char_mask::FE_HINT;
        }
        if self.kumi {
            mask |= char_mask::KUMI;
        }
        if self.strikethrough {
            mask |= char_mask::LEGACY_STRIKETHROUGH;
        }
        if self.emboss {
            mask |= char_mask::EMBOSS;
        }
        if self.pp9_run_id.is_some() {
            mask |= char_mask::PP9_RUN_ID;
        }
        mask
    }

    /// Convert to flags value
    pub fn to_flags(&self) -> u16 {
        let mut flags = 0u16;
        if self.bold {
            flags |= 0x0001;
        }
        if self.italic {
            flags |= 0x0002;
        }
        if self.underline {
            flags |= 0x0004;
        }
        if self.shadow {
            flags |= 0x0010;
        }
        if self.fe_hint {
            flags |= 0x0020;
        }
        if self.kumi {
            flags |= 0x0080;
        }
        if self.strikethrough {
            flags |= 0x0100;
        }
        if self.emboss {
            flags |= 0x0200;
        }
        flags |= u16::from(self.pp9_run_id.unwrap_or(0) & 0x0F) << 10;
        flags
    }
}

// =============================================================================
// Color
// =============================================================================

/// Text color representation
#[derive(Debug, Clone, Copy)]
pub struct TextColor {
    /// Red component (0-255)
    pub r: u8,
    /// Green component (0-255)
    pub g: u8,
    /// Blue component (0-255)
    pub b: u8,
    /// Use scheme color instead of RGB
    pub use_scheme: bool,
    /// Scheme color index (if use_scheme is true)
    pub scheme_index: u8,
}

impl TextColor {
    /// Black color
    pub const BLACK: Self = Self::rgb(0, 0, 0);
    /// White color
    pub const WHITE: Self = Self::rgb(255, 255, 255);
    /// Red color
    pub const RED: Self = Self::rgb(255, 0, 0);
    /// Green color
    pub const GREEN: Self = Self::rgb(0, 255, 0);
    /// Blue color
    pub const BLUE: Self = Self::rgb(0, 0, 255);

    /// Create an RGB color
    pub const fn rgb(r: u8, g: u8, b: u8) -> Self {
        Self {
            r,
            g,
            b,
            use_scheme: false,
            scheme_index: 0,
        }
    }

    /// Create from hex value (0xRRGGBB)
    pub const fn from_hex(hex: u32) -> Self {
        Self::rgb(
            ((hex >> 16) & 0xFF) as u8,
            ((hex >> 8) & 0xFF) as u8,
            (hex & 0xFF) as u8,
        )
    }

    /// Create a scheme color reference
    pub const fn scheme(index: u8) -> Self {
        Self {
            r: 0,
            g: 0,
            b: 0,
            use_scheme: true,
            scheme_index: index,
        }
    }

    /// Convert to PPT font color format
    /// POI uses: new Color(blue, green, red, 254).getRGB() which produces
    /// (254 << 24) | (blue << 16) | (green << 8) | red
    pub fn to_ppt_color(&self) -> u32 {
        if self.use_scheme {
            // ColorIndexStruct stores the scheme index in its fourth byte.
            (self.scheme_index as u32) << 24
        } else {
            // Format: R | (G << 8) | (B << 16) | (alpha << 24)
            // Alpha = 0xFE (254) for opaque colors
            (self.r as u32) | ((self.g as u32) << 8) | ((self.b as u32) << 16) | 0xFE00_0000
        }
    }
}

impl Default for TextColor {
    fn default() -> Self {
        Self::BLACK
    }
}

// =============================================================================
// Text Run
// =============================================================================

/// A run of text with consistent formatting
#[derive(Debug, Clone)]
pub struct TextRun {
    /// The text content
    pub text: String,
    /// Font style (bold, italic, etc.)
    pub style: FontStyle,
    /// Font size in points
    pub font_size: u16,
    /// Text color
    pub color: TextColor,
    /// Font index (reference to FontCollection, 0 = default)
    pub font_index: u16,
    /// East Asian font reference
    pub asian_font_index: Option<u16>,
    /// ANSI font reference
    pub ansi_font_index: Option<u16>,
    /// Symbol font reference
    pub symbol_font_index: Option<u16>,
    /// Baseline position as a percentage of line height
    pub baseline_position: Option<i16>,
    /// Zero-based indices into the presentation-wide PowerPoint 11 smart-tag store.
    pub smart_tag_indices: Vec<SmartTagIndex>,
}

impl TextRun {
    /// Create a new text run with default formatting
    pub fn new(text: impl Into<String>) -> Self {
        Self {
            text: text.into(),
            style: FontStyle::default(),
            font_size: 18, // Default 18pt
            color: TextColor::BLACK,
            font_index: 0,
            asian_font_index: None,
            ansi_font_index: None,
            symbol_font_index: None,
            baseline_position: None,
            smart_tag_indices: Vec::new(),
        }
    }

    /// Set bold
    pub fn bold(mut self) -> Self {
        self.style.bold = true;
        self.style.specified_mask |= char_mask::BOLD as u16;
        self
    }

    /// Explicitly set bold formatting, including false.
    pub fn bold_value(mut self, enabled: bool) -> Self {
        self.style.bold = enabled;
        self.style.specified_mask |= char_mask::BOLD as u16;
        self
    }

    /// Set italic
    pub fn italic(mut self) -> Self {
        self.style.italic = true;
        self.style.specified_mask |= char_mask::ITALIC as u16;
        self
    }

    /// Explicitly set italic formatting, including false.
    pub fn italic_value(mut self, enabled: bool) -> Self {
        self.style.italic = enabled;
        self.style.specified_mask |= char_mask::ITALIC as u16;
        self
    }

    /// Set underline
    pub fn underline(mut self) -> Self {
        self.style.underline = true;
        self.style.specified_mask |= char_mask::UNDERLINE as u16;
        self
    }

    /// Explicitly set underline formatting, including false.
    pub fn underline_value(mut self, enabled: bool) -> Self {
        self.style.underline = enabled;
        self.style.specified_mask |= char_mask::UNDERLINE as u16;
        self
    }

    /// Set shadow formatting.
    pub fn shadow(mut self) -> Self {
        self.style.shadow = true;
        self.style.specified_mask |= char_mask::SHADOW as u16;
        self
    }

    /// Explicitly set shadow formatting, including false.
    pub fn shadow_value(mut self, enabled: bool) -> Self {
        self.style.shadow = enabled;
        self.style.specified_mask |= char_mask::SHADOW as u16;
        self
    }

    /// Set embossed/relief formatting.
    pub fn embossed(mut self) -> Self {
        self.style.emboss = true;
        self.style.specified_mask |= char_mask::EMBOSS as u16;
        self
    }

    /// Explicitly set embossed formatting, including false.
    pub fn embossed_value(mut self, enabled: bool) -> Self {
        self.style.emboss = enabled;
        self.style.specified_mask |= char_mask::EMBOSS as u16;
        self
    }

    /// Mark the run as originating from double-byte input.
    pub fn fe_hint(mut self, enabled: bool) -> Self {
        self.style.fe_hint = enabled;
        self.style.specified_mask |= char_mask::FE_HINT as u16;
        self
    }

    /// Set Kumimoji formatting for vertical text.
    pub fn kumi(mut self, enabled: bool) -> Self {
        self.style.kumi = enabled;
        self.style.specified_mask |= char_mask::KUMI as u16;
        self
    }

    /// Set the de facto legacy strikethrough bit used by PowerPoint and POI.
    pub fn strikethrough(mut self, enabled: bool) -> Self {
        self.style.strikethrough = enabled;
        self.style.specified_mask |= char_mask::LEGACY_STRIKETHROUGH as u16;
        self
    }

    /// Set the PowerPoint 9 additional-property run grouping identifier.
    pub fn pp9_run_id(mut self, id: u8) -> Self {
        self.style.pp9_run_id = Some(id);
        self
    }

    /// Attach one document-wide PowerPoint 11 smart tag to this text run.
    pub fn with_smart_tag(mut self, index: SmartTagIndex) -> Self {
        self.smart_tag_indices.push(index);
        self
    }

    /// Attach document-wide PowerPoint 11 smart tags to this text run.
    pub fn with_smart_tags(mut self, indices: impl IntoIterator<Item = SmartTagIndex>) -> Self {
        self.smart_tag_indices.extend(indices);
        self
    }

    /// Attach one document-wide PowerPoint 11 smart tag in place.
    pub fn add_smart_tag(&mut self, index: SmartTagIndex) {
        self.smart_tag_indices.push(index);
    }

    /// Set font size in points
    pub fn size(mut self, points: u16) -> Self {
        self.font_size = points;
        self
    }

    /// Set color from RGB
    pub fn color_rgb(mut self, r: u8, g: u8, b: u8) -> Self {
        self.color = TextColor::rgb(r, g, b);
        self
    }

    /// Set color from hex
    pub fn color_hex(mut self, hex: u32) -> Self {
        self.color = TextColor::from_hex(hex);
        self
    }

    /// Set a color-scheme index.
    pub fn color_scheme(mut self, index: u8) -> Self {
        self.color = TextColor::scheme(index);
        self
    }

    /// Set font index
    pub fn font(mut self, index: u16) -> Self {
        self.font_index = index;
        self
    }

    /// Set the East Asian font reference.
    pub fn asian_font(mut self, index: u16) -> Self {
        self.asian_font_index = Some(index);
        self
    }

    /// Set the ANSI font reference.
    pub fn ansi_font(mut self, index: u16) -> Self {
        self.ansi_font_index = Some(index);
        self
    }

    /// Set the symbol font reference.
    pub fn symbol_font(mut self, index: u16) -> Self {
        self.symbol_font_index = Some(index);
        self
    }

    /// Set superscript (positive) or subscript (negative) baseline position.
    pub fn baseline_position(mut self, percent: i16) -> Self {
        self.baseline_position = Some(percent);
        self
    }

    /// Get the number of UTF-16 code units in this run.
    pub fn char_count(&self) -> u32 {
        self.text.encode_utf16().count() as u32
    }
}

// =============================================================================
// Paragraph
// =============================================================================

/// A paragraph containing one or more text runs
#[derive(Debug, Clone)]
pub struct Paragraph {
    /// Text runs in this paragraph
    pub runs: Vec<TextRun>,
    /// Text alignment
    pub alignment: TextAlign,
    /// Line spacing (in percent * 100, e.g., 100 = 1.0, 150 = 1.5)
    pub line_spacing: i16,
    /// Space before paragraph (in master units)
    pub space_before: i16,
    /// Space after paragraph (in master units)
    pub space_after: i16,
    /// Left margin (in master units)
    pub left_margin: i16,
    /// First line indent (in master units)
    pub indent: i16,
    /// Paragraph indent level (`0..=4`)
    pub indent_level: u16,
    /// Bullet character (if any)
    pub bullet_char: Option<char>,
    /// Explicit bullet-presence flag
    pub bullet_enabled: Option<bool>,
    /// Bullet font reference
    pub bullet_font_index: Option<u16>,
    /// Explicit bullet-font validity flag
    pub bullet_font_enabled: Option<bool>,
    /// Raw bullet size (percentage when positive, points when negative)
    pub bullet_size: Option<i16>,
    /// Explicit bullet-size validity flag
    pub bullet_size_enabled: Option<bool>,
    /// Bullet color
    pub bullet_color: Option<TextColor>,
    /// Explicit bullet-color validity flag
    pub bullet_color_enabled: Option<bool>,
    /// Default tab size in master units
    pub default_tab_size: Option<i16>,
    /// Explicit tab stops; `Some(empty)` writes an empty array
    pub tab_stops: Option<Vec<TabStop>>,
    /// Font alignment within the line
    pub font_alignment: Option<TextFontAlign>,
    /// East Asian character-wrapping override
    pub character_wrap: Option<bool>,
    /// Word-wrapping override
    pub word_wrap: Option<bool>,
    /// Hanging-punctuation override
    pub overflow: Option<bool>,
    /// Paragraph text direction
    pub text_direction: Option<TextDirection>,
    /// Properties explicitly requested even when their values equal defaults
    pub(super) explicit_mask: u32,
}

impl Paragraph {
    /// Create a new paragraph with a single text run
    pub fn new(text: impl Into<String>) -> Self {
        Self {
            runs: vec![TextRun::new(text)],
            alignment: TextAlign::Left,
            line_spacing: 100,
            space_before: 0,
            space_after: 0,
            left_margin: 0,
            indent: 0,
            indent_level: 0,
            bullet_char: None,
            bullet_enabled: None,
            bullet_font_index: None,
            bullet_font_enabled: None,
            bullet_size: None,
            bullet_size_enabled: None,
            bullet_color: None,
            bullet_color_enabled: None,
            default_tab_size: None,
            tab_stops: None,
            font_alignment: None,
            character_wrap: None,
            word_wrap: None,
            overflow: None,
            text_direction: None,
            explicit_mask: 0,
        }
    }

    /// Create from multiple runs
    pub fn with_runs(runs: Vec<TextRun>) -> Self {
        Self {
            runs,
            alignment: TextAlign::Left,
            line_spacing: 100,
            space_before: 0,
            space_after: 0,
            left_margin: 0,
            indent: 0,
            indent_level: 0,
            bullet_char: None,
            bullet_enabled: None,
            bullet_font_index: None,
            bullet_font_enabled: None,
            bullet_size: None,
            bullet_size_enabled: None,
            bullet_color: None,
            bullet_color_enabled: None,
            default_tab_size: None,
            tab_stops: None,
            font_alignment: None,
            character_wrap: None,
            word_wrap: None,
            overflow: None,
            text_direction: None,
            explicit_mask: 0,
        }
    }

    /// Set alignment
    pub fn align(mut self, alignment: TextAlign) -> Self {
        self.alignment = alignment;
        self.explicit_mask |= para_mask::ALIGNMENT;
        self
    }

    /// Center align
    pub fn center(mut self) -> Self {
        self.alignment = TextAlign::Center;
        self.explicit_mask |= para_mask::ALIGNMENT;
        self
    }

    /// Right align
    pub fn right(mut self) -> Self {
        self.alignment = TextAlign::Right;
        self.explicit_mask |= para_mask::ALIGNMENT;
        self
    }

    /// Set line spacing (percent)
    pub fn line_spacing(mut self, percent: i16) -> Self {
        self.line_spacing = percent;
        self.explicit_mask |= para_mask::LINE_SPACING;
        self
    }

    /// Set space before
    pub fn space_before(mut self, units: i16) -> Self {
        self.space_before = units;
        self.explicit_mask |= para_mask::SPACE_BEFORE;
        self
    }

    /// Set space after
    pub fn space_after(mut self, units: i16) -> Self {
        self.space_after = units;
        self.explicit_mask |= para_mask::SPACE_AFTER;
        self
    }

    /// Set the left margin in master units, including an explicit zero.
    pub fn left_margin(mut self, units: i16) -> Self {
        self.left_margin = units;
        self.explicit_mask |= para_mask::LEFT_MARGIN;
        self
    }

    /// Set the first-line indent in master units, including an explicit zero.
    pub fn first_line_indent(mut self, units: i16) -> Self {
        self.indent = units;
        self.explicit_mask |= para_mask::INDENT;
        self
    }

    /// Add bullet
    pub fn with_bullet(mut self, ch: char) -> Self {
        self.bullet_char = Some(ch);
        self.bullet_enabled = Some(true);
        self
    }

    /// Explicitly enable or disable paragraph bullets.
    pub fn bullet_enabled(mut self, enabled: bool) -> Self {
        self.bullet_enabled = Some(enabled);
        self
    }

    /// Set the bullet font reference and mark it active.
    pub fn bullet_font(mut self, index: u16) -> Self {
        self.bullet_font_index = Some(index);
        self.bullet_font_enabled = Some(true);
        self
    }

    /// Explicitly enable or disable the bullet font override.
    pub fn bullet_font_enabled(mut self, enabled: bool) -> Self {
        self.bullet_font_enabled = Some(enabled);
        self
    }

    /// Set the raw `BulletSize` value and mark it active.
    pub fn bullet_size(mut self, size: i16) -> Self {
        self.bullet_size = Some(size);
        self.bullet_size_enabled = Some(true);
        self
    }

    /// Explicitly enable or disable the bullet size override.
    pub fn bullet_size_enabled(mut self, enabled: bool) -> Self {
        self.bullet_size_enabled = Some(enabled);
        self
    }

    /// Set a direct sRGB bullet color.
    pub fn bullet_color_rgb(mut self, r: u8, g: u8, b: u8) -> Self {
        self.bullet_color = Some(TextColor::rgb(r, g, b));
        self.bullet_color_enabled = Some(true);
        self
    }

    /// Set a color-scheme index for the bullet.
    pub fn bullet_color_scheme(mut self, index: u8) -> Self {
        self.bullet_color = Some(TextColor::scheme(index));
        self.bullet_color_enabled = Some(true);
        self
    }

    /// Explicitly enable or disable the bullet color override.
    pub fn bullet_color_enabled(mut self, enabled: bool) -> Self {
        self.bullet_color_enabled = Some(enabled);
        self
    }

    /// Set the paragraph indent level.
    pub fn indent_level(mut self, level: u16) -> Self {
        self.indent_level = level;
        self
    }

    /// Set the default tab size in master units.
    pub fn default_tab_size(mut self, size: i16) -> Self {
        self.default_tab_size = Some(size);
        self
    }

    /// Set explicit paragraph tab stops.
    pub fn tab_stops(mut self, stops: Vec<TabStop>) -> Self {
        self.tab_stops = Some(stops);
        self
    }

    /// Set font alignment within the line.
    pub fn font_alignment(mut self, alignment: TextFontAlign) -> Self {
        self.font_alignment = Some(alignment);
        self
    }

    /// Set the East Asian character-wrapping override.
    pub fn character_wrap(mut self, enabled: bool) -> Self {
        self.character_wrap = Some(enabled);
        self
    }

    /// Set the word-wrapping override.
    pub fn word_wrap(mut self, enabled: bool) -> Self {
        self.word_wrap = Some(enabled);
        self
    }

    /// Set the hanging-punctuation override.
    pub fn overflow(mut self, enabled: bool) -> Self {
        self.overflow = Some(enabled);
        self
    }

    /// Set paragraph text direction.
    pub fn text_direction(mut self, direction: TextDirection) -> Self {
        self.text_direction = Some(direction);
        self
    }

    /// Get total character count for this paragraph's runs only (no paragraph marker)
    pub fn runs_char_count(&self) -> u32 {
        self.runs.iter().map(|r| r.char_count()).sum::<u32>()
    }

    /// Get total character count including paragraph separator
    /// Note: The last paragraph in a sequence should NOT include +1
    pub fn char_count(&self) -> u32 {
        self.runs_char_count() + 1 // +1 for paragraph end marker (CR)
    }

    /// Get combined text
    pub fn text(&self) -> String {
        self.runs.iter().map(|r| r.text.as_str()).collect()
    }
}

// =============================================================================
// Text Style Header Structures
// =============================================================================

/// TextHeaderAtom type codes
#[repr(u32)]
#[derive(Debug, Clone, Copy)]
pub enum TextHeaderType {
    /// Title text
    Title = 0,
    /// Body text
    Body = 1,
    /// Notes text
    Notes = 2,
    /// Other (non-placeholder)
    Other = 4,
    /// Center body
    CenterBody = 5,
    /// Center title
    CenterTitle = 6,
    /// Half body
    HalfBody = 7,
    /// Quarter body
    QuarterBody = 8,
}

/// StyleTextPropAtom header
#[derive(Debug, Clone, Copy, IntoBytes, Immutable, KnownLayout)]
#[repr(C, packed)]
pub struct StyleTextPropHeader {
    /// Total character count
    pub char_count: u32,
}

// Font Entity
// =============================================================================

/// Font entity for FontCollection
#[derive(Debug, Clone)]
pub struct FontEntity {
    /// Font face name (max 32 characters)
    pub name: String,
    /// Font type flags (0x01 = raster, 0x02 = device, 0x04 = TrueType)
    pub font_type: u8,
    /// Pitch and family
    pub pitch_family: u8,
    /// Character set
    pub charset: u8,
}

impl FontEntity {
    /// Create Arial font (default)
    pub fn arial() -> Self {
        Self {
            name: "Arial".to_string(),
            font_type: 0x04,    // TrueType
            pitch_family: 0x22, // Variable pitch, Swiss family
            charset: 0x00,      // ANSI
        }
    }

    /// Create Times New Roman font
    pub fn times_new_roman() -> Self {
        Self {
            name: "Times New Roman".to_string(),
            font_type: 0x04,
            pitch_family: 0x12, // Variable pitch, Roman family
            charset: 0x00,
        }
    }

    /// Create custom font
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            font_type: 0x04,
            pitch_family: 0x00,
            charset: 0x00,
        }
    }

    /// Build FontEntityAtom (68 bytes)
    pub fn build(&self) -> Vec<u8> {
        let mut data = vec![0u8; 68];

        // Write font name as UTF-16LE (max 32 chars = 64 bytes)
        for (i, ch) in self.name.encode_utf16().take(31).enumerate() {
            let bytes = ch.to_le_bytes();
            data[i * 2] = bytes[0];
            data[i * 2 + 1] = bytes[1];
        }

        // Font metadata at offset 64-67
        data[64] = self.charset;
        data[65] = 0;
        data[66] = self.font_type;
        data[67] = self.pitch_family;

        data
    }
}
