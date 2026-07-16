//! Text formatting support for PPT files
//!
//! This module handles text styling including bold, italic, underline,
//! font size, font colors, and paragraph formatting.
//!
//! Reference: [MS-PPT] Section 2.9 - Text Formatting

use zerocopy_derive::*;

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
    explicit_mask: u32,
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
#[derive(Debug, Clone, Copy, FromBytes, IntoBytes, Immutable, KnownLayout)]
#[repr(C, packed)]
pub struct StyleTextPropHeader {
    /// Total character count
    pub char_count: u32,
}

// =============================================================================
// Text Properties Builder
// =============================================================================

/// Builder for TextCharsAtom/TextBytesAtom and StyleTextPropAtom
pub struct TextPropsBuilder {
    paragraphs: Vec<Paragraph>,
}

impl TextPropsBuilder {
    /// Create a new builder
    pub fn new() -> Self {
        Self {
            paragraphs: Vec::new(),
        }
    }

    /// Add a paragraph
    pub fn add_paragraph(&mut self, para: Paragraph) {
        self.paragraphs.push(para);
    }

    /// Build TextCharsAtom (UTF-16LE text), adding CR between paragraphs.
    ///
    /// The final paragraph break is implicit and is not stored in the text atom.
    pub fn build_text_chars(&self) -> Vec<u8> {
        let mut data = Vec::new();
        for (i, para) in self.paragraphs.iter().enumerate() {
            for run in &para.runs {
                for ch in run.text.encode_utf16() {
                    data.extend_from_slice(&ch.to_le_bytes());
                }
            }
            // Add paragraph separator (CR) for all paragraphs including the last
            // This makes the text length match the StyleTextPropAtom char counts
            if i < self.paragraphs.len() - 1 {
                data.extend_from_slice(&0x000Du16.to_le_bytes()); // CR between paragraphs
            }
        }
        data
    }

    /// Build StyleTextPropAtom containing paragraph and character formatting
    ///
    /// According to MS-PPT spec:
    /// - Sum of paragraph character counts = total text length + 1
    /// - Sum of character run counts = total text length + 1
    /// - The +1 accounts for an implicit terminating character
    pub fn build_style_text_prop(&self) -> std::io::Result<Vec<u8>> {
        let mut data = Vec::new();

        if self.paragraphs.is_empty() {
            data.extend_from_slice(&1u32.to_le_bytes());
            data.extend_from_slice(&0i16.to_le_bytes());
            data.extend_from_slice(&0u32.to_le_bytes());
            data.extend_from_slice(&1u32.to_le_bytes());
            data.extend_from_slice(&0u32.to_le_bytes());
            return Ok(data);
        }

        // Paragraph properties (TextPFRun entries)
        // Each paragraph covers its runs + CR separator (except last paragraph gets +1 for terminator)
        for para in &self.paragraphs {
            let para_text_len = para.runs.iter().try_fold(0u32, |total, run| {
                let count = u32::try_from(run.text.encode_utf16().count()).map_err(|_| {
                    std::io::Error::new(
                        std::io::ErrorKind::InvalidInput,
                        "PowerPoint text run exceeds the PPT size limit",
                    )
                })?;
                total.checked_add(count).ok_or_else(|| {
                    std::io::Error::new(
                        std::io::ErrorKind::InvalidInput,
                        "PowerPoint paragraph exceeds the PPT size limit",
                    )
                })
            })?;

            // Character count: text + CR (or +1 for last paragraph terminator)
            // +1 for either CR separator or implicit terminating character
            let char_count = para_text_len.checked_add(1).ok_or_else(|| {
                std::io::Error::new(
                    std::io::ErrorKind::InvalidInput,
                    "PowerPoint paragraph exceeds the PPT size limit",
                )
            })?;
            data.extend_from_slice(&char_count.to_le_bytes());

            if para.indent_level > 4 {
                return Err(std::io::Error::new(
                    std::io::ErrorKind::InvalidInput,
                    "PPT paragraph indent level must be between 0 and 4",
                ));
            }
            data.extend_from_slice(&para.indent_level.to_le_bytes());

            // Build mask based on what properties are set
            let mut mask = para.explicit_mask;
            if para.alignment != TextAlign::Left {
                mask |= para_mask::ALIGNMENT;
            }
            if para.line_spacing != 100 {
                mask |= para_mask::LINE_SPACING;
            }
            if para.space_before != 0 {
                mask |= para_mask::SPACE_BEFORE;
            }
            if para.space_after != 0 {
                mask |= para_mask::SPACE_AFTER;
            }
            if para.left_margin != 0 {
                mask |= para_mask::LEFT_MARGIN;
            }
            if para.indent != 0 {
                mask |= para_mask::INDENT;
            }
            if para.bullet_enabled.is_some() || para.bullet_char.is_some() {
                mask |= para_mask::HAS_BULLET;
            }
            if para.bullet_char.is_some() {
                mask |= para_mask::BULLET_CHAR;
            }
            if para.bullet_font_enabled.is_some() || para.bullet_font_index.is_some() {
                mask |= para_mask::BULLET_HAS_FONT;
            }
            if para.bullet_font_index.is_some() {
                mask |= para_mask::BULLET_FONT;
            }
            if para.bullet_size_enabled.is_some() || para.bullet_size.is_some() {
                mask |= para_mask::BULLET_HAS_SIZE;
            }
            if para.bullet_size.is_some() {
                mask |= para_mask::BULLET_SIZE;
            }
            if para.bullet_color_enabled.is_some() || para.bullet_color.is_some() {
                mask |= para_mask::BULLET_HAS_COLOR;
            }
            if para.bullet_color.is_some() {
                mask |= para_mask::BULLET_COLOR;
            }
            if para.default_tab_size.is_some() {
                mask |= para_mask::DEFAULT_TAB_SIZE;
            }
            if para.tab_stops.is_some() {
                mask |= para_mask::TAB_STOPS;
            }
            if para.font_alignment.is_some() {
                mask |= para_mask::FONT_ALIGNMENT;
            }
            if para.character_wrap.is_some() {
                mask |= para_mask::CHARACTER_WRAP;
            }
            if para.word_wrap.is_some() {
                mask |= para_mask::WORD_WRAP;
            }
            if para.overflow.is_some() {
                mask |= para_mask::OVERFLOW;
            }
            if para.text_direction.is_some() {
                mask |= para_mask::TEXT_DIRECTION;
            }

            data.extend_from_slice(&mask.to_le_bytes());

            // Write properties according to mask
            if mask
                & (para_mask::HAS_BULLET
                    | para_mask::BULLET_HAS_FONT
                    | para_mask::BULLET_HAS_COLOR
                    | para_mask::BULLET_HAS_SIZE)
                != 0
            {
                let mut flags = 0u16;
                if para.bullet_enabled.unwrap_or(para.bullet_char.is_some()) {
                    flags |= 0x0001;
                }
                if para
                    .bullet_font_enabled
                    .unwrap_or(para.bullet_font_index.is_some())
                {
                    flags |= 0x0002;
                }
                if para
                    .bullet_color_enabled
                    .unwrap_or(para.bullet_color.is_some())
                {
                    flags |= 0x0004;
                }
                if para
                    .bullet_size_enabled
                    .unwrap_or(para.bullet_size.is_some())
                {
                    flags |= 0x0008;
                }
                data.extend_from_slice(&flags.to_le_bytes());
            }
            if mask & para_mask::BULLET_CHAR != 0 {
                let bullet = para.bullet_char.unwrap_or('•');
                let ch = u16::try_from(bullet as u32).map_err(|_| {
                    std::io::Error::new(
                        std::io::ErrorKind::InvalidInput,
                        "PPT bullet characters must fit in one UTF-16 code unit",
                    )
                })?;
                data.extend_from_slice(&ch.to_le_bytes());
            }
            if mask & para_mask::BULLET_FONT != 0 {
                data.extend_from_slice(&para.bullet_font_index.unwrap_or(0).to_le_bytes());
            }
            if mask & para_mask::BULLET_SIZE != 0 {
                let size = para.bullet_size.unwrap_or(100);
                if !((25..=400).contains(&size) || (-4000..=-1).contains(&size)) {
                    return Err(std::io::Error::new(
                        std::io::ErrorKind::InvalidInput,
                        "PPT bullet size must be 25..=400 percent or -4000..=-1 points",
                    ));
                }
                data.extend_from_slice(&size.to_le_bytes());
            }
            if mask & para_mask::BULLET_COLOR != 0 {
                let color = para.bullet_color.unwrap_or(TextColor::BLACK);
                if color.use_scheme && color.scheme_index > 7 {
                    return Err(std::io::Error::new(
                        std::io::ErrorKind::InvalidInput,
                        "PPT bullet color-scheme index must be between 0 and 7",
                    ));
                }
                data.extend_from_slice(&color.to_ppt_color().to_le_bytes());
            }
            if mask & para_mask::ALIGNMENT != 0 {
                data.extend_from_slice(&(para.alignment as u16).to_le_bytes());
            }
            if mask & para_mask::LINE_SPACING != 0 {
                data.extend_from_slice(&para.line_spacing.to_le_bytes());
            }
            if mask & para_mask::SPACE_BEFORE != 0 {
                data.extend_from_slice(&para.space_before.to_le_bytes());
            }
            if mask & para_mask::SPACE_AFTER != 0 {
                data.extend_from_slice(&para.space_after.to_le_bytes());
            }
            if mask & para_mask::LEFT_MARGIN != 0 {
                data.extend_from_slice(&para.left_margin.to_le_bytes());
            }
            if mask & para_mask::INDENT != 0 {
                data.extend_from_slice(&para.indent.to_le_bytes());
            }
            if mask & para_mask::DEFAULT_TAB_SIZE != 0 {
                data.extend_from_slice(&para.default_tab_size.unwrap_or(0).to_le_bytes());
            }
            if mask & para_mask::TAB_STOPS != 0 {
                let tab_stops = para.tab_stops.as_deref().unwrap_or_default();
                let count = u16::try_from(tab_stops.len()).map_err(|_| {
                    std::io::Error::new(
                        std::io::ErrorKind::InvalidInput,
                        "PPT paragraph has more than 65535 tab stops",
                    )
                })?;
                data.extend_from_slice(&count.to_le_bytes());
                for tab_stop in tab_stops {
                    data.extend_from_slice(&tab_stop.position.to_le_bytes());
                    data.extend_from_slice(&(tab_stop.alignment as u16).to_le_bytes());
                }
            }
            if mask & para_mask::FONT_ALIGNMENT != 0 {
                data.extend_from_slice(
                    &(para.font_alignment.unwrap_or(TextFontAlign::Roman) as u16).to_le_bytes(),
                );
            }
            if mask & (para_mask::CHARACTER_WRAP | para_mask::WORD_WRAP | para_mask::OVERFLOW) != 0
            {
                let mut flags = 0u16;
                if para.character_wrap.unwrap_or(false) {
                    flags |= 0x0001;
                }
                if para.word_wrap.unwrap_or(false) {
                    flags |= 0x0002;
                }
                if para.overflow.unwrap_or(false) {
                    flags |= 0x0004;
                }
                data.extend_from_slice(&flags.to_le_bytes());
            }
            if mask & para_mask::TEXT_DIRECTION != 0 {
                data.extend_from_slice(
                    &(para.text_direction.unwrap_or(TextDirection::LeftToRight) as u16)
                        .to_le_bytes(),
                );
            }
        }

        // Character properties (TextCFRun entries)
        // Write one entry per run. The last run in each paragraph gets +1 for CR/terminator.
        for para in &self.paragraphs {
            let num_runs = para.runs.len();

            if num_runs == 0 {
                // Cover the paragraph separator or the implicit final paragraph break.
                data.extend_from_slice(&1u32.to_le_bytes());
                data.extend_from_slice(&0u32.to_le_bytes());
                continue;
            }

            for (run_idx, run) in para.runs.iter().enumerate() {
                if !(1..=4000).contains(&run.font_size) {
                    return Err(std::io::Error::new(
                        std::io::ErrorKind::InvalidInput,
                        "PPT font size must be between 1 and 4000 points",
                    ));
                }
                if run.color.use_scheme && run.color.scheme_index > 7 {
                    return Err(std::io::Error::new(
                        std::io::ErrorKind::InvalidInput,
                        "PPT color-scheme index must be between 0 and 7",
                    ));
                }
                if run
                    .baseline_position
                    .is_some_and(|position| !(-100..=100).contains(&position))
                {
                    return Err(std::io::Error::new(
                        std::io::ErrorKind::InvalidInput,
                        "PPT baseline position must be between -100 and 100 percent",
                    ));
                }
                if run.style.pp9_run_id.is_some_and(|id| id > 0x0F) {
                    return Err(std::io::Error::new(
                        std::io::ErrorKind::InvalidInput,
                        "PPT pp9 run grouping identifier must be between 0 and 15",
                    ));
                }
                if run.style.specified_mask & !0x3FB7 != 0 {
                    return Err(std::io::Error::new(
                        std::io::ErrorKind::InvalidInput,
                        "PPT character style specified mask contains reserved bits",
                    ));
                }
                let is_last_run = run_idx == num_runs - 1;

                // Character count for this run
                // Last run of last paragraph gets +1 for terminator
                // Last run of non-last paragraph gets +1 for CR separator
                let run_units = u32::try_from(run.text.encode_utf16().count()).map_err(|_| {
                    std::io::Error::new(
                        std::io::ErrorKind::InvalidInput,
                        "PowerPoint text run exceeds the PPT size limit",
                    )
                })?;
                let char_count = if is_last_run {
                    run_units.checked_add(1).ok_or_else(|| {
                        std::io::Error::new(
                            std::io::ErrorKind::InvalidInput,
                            "PowerPoint text run exceeds the PPT size limit",
                        )
                    })?
                } else {
                    run_units
                };
                data.extend_from_slice(&char_count.to_le_bytes());

                // Build mask
                let mut mask = run.style.to_mask();
                mask |= char_mask::FONT_SIZE; // Always include font size
                mask |= char_mask::FONT_COLOR; // Always include color
                mask |= char_mask::FONT_REF; // Always include font reference
                if run.asian_font_index.is_some() {
                    mask |= char_mask::ASIAN_FONT_REF;
                }
                if run.ansi_font_index.is_some() {
                    mask |= char_mask::ANSI_FONT_REF;
                }
                if run.symbol_font_index.is_some() {
                    mask |= char_mask::SYMBOL_FONT_REF;
                }
                if run.baseline_position.is_some() {
                    mask |= char_mask::POSITION;
                }

                data.extend_from_slice(&mask.to_le_bytes());

                // Font style flags (only if any flags are set)
                if mask & 0xFFFF != 0 {
                    let flags = run.style.to_flags();
                    data.extend_from_slice(&flags.to_le_bytes());
                }

                // Font reference (if font_ref bit is set)
                data.extend_from_slice(&run.font_index.to_le_bytes());

                if let Some(index) = run.asian_font_index {
                    data.extend_from_slice(&index.to_le_bytes());
                }
                if let Some(index) = run.ansi_font_index {
                    data.extend_from_slice(&index.to_le_bytes());
                }
                if let Some(index) = run.symbol_font_index {
                    data.extend_from_slice(&index.to_le_bytes());
                }

                // Font size is stored directly in points.
                data.extend_from_slice(&run.font_size.to_le_bytes());

                // Color (POI format: R | G<<8 | B<<16 | 0xFE<<24)
                let color = run.color.to_ppt_color();
                data.extend_from_slice(&color.to_le_bytes());

                if let Some(position) = run.baseline_position {
                    data.extend_from_slice(&position.to_le_bytes());
                }
            }
        }

        Ok(data)
    }

    /// Get total character count
    pub fn total_chars(&self) -> u32 {
        self.paragraphs.iter().map(|p| p.char_count()).sum()
    }
}

impl Default for TextPropsBuilder {
    fn default() -> Self {
        Self::new()
    }
}

// =============================================================================
// Font Entity
// =============================================================================

/// Font entity for FontCollection
#[derive(Debug, Clone)]
pub struct FontEntity {
    /// Font face name (max 32 characters)
    pub name: String,
    /// Font type (0x00 = raster, 0x02 = device, 0x04 = TrueType)
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
        for (i, ch) in self.name.encode_utf16().take(32).enumerate() {
            let bytes = ch.to_le_bytes();
            data[i * 2] = bytes[0];
            data[i * 2 + 1] = bytes[1];
        }

        // Font metadata at offset 64-67
        data[64] = self.pitch_family;
        data[65] = self.charset;
        data[66] = self.font_type;
        data[67] = 0; // Reserved

        data
    }
}

// =============================================================================
// Tests
// =============================================================================

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_font_style() {
        let style = FontStyle::bold_italic();
        assert!(style.bold);
        assert!(style.italic);
        assert_eq!(style.to_flags(), 0x0003);
    }

    #[test]
    fn test_text_color() {
        // Red = RGB(255, 0, 0) -> PPT format: R | G<<8 | B<<16 | 0xFE<<24 = 0xFE0000FF
        let red = TextColor::RED;
        assert_eq!(red.to_ppt_color(), 0xFE0000FF);

        // Scheme color index occupies the fourth byte.
        let scheme = TextColor::scheme(4);
        assert_eq!(scheme.to_ppt_color(), 0x04000000);
    }

    #[test]
    fn test_text_run() {
        let run = TextRun::new("Hello").bold().size(24);
        assert!(run.style.bold);
        assert_eq!(run.font_size, 24);
        assert_eq!(run.char_count(), 5);
    }

    #[test]
    fn text_run_counts_utf16_code_units() {
        let run = TextRun::new("😀x");
        assert_eq!(run.char_count(), 3);
    }

    #[test]
    fn test_paragraph() {
        let para = Paragraph::new("Test").center();
        assert_eq!(para.alignment, TextAlign::Center);
        assert_eq!(para.char_count(), 5); // 4 chars + 1 end marker
    }

    #[test]
    fn empty_rich_paragraph_has_character_style_coverage() {
        let mut builder = TextPropsBuilder::new();
        builder.add_paragraph(Paragraph::with_runs(Vec::new()));

        let style = builder.build_style_text_prop().unwrap();
        let (paragraphs, characters) = crate::ppt::text_prop::parse_style_text_prop_atom(&style, 0);

        assert_eq!(paragraphs.len(), 1);
        assert_eq!(paragraphs[0].characters_covered, 1);
        assert_eq!(characters.len(), 1);
        assert_eq!(characters[0].characters_covered, 1);
    }

    #[test]
    fn empty_rich_text_has_complete_style_coverage() {
        let style = TextPropsBuilder::new().build_style_text_prop().unwrap();
        let (paragraphs, characters) = crate::ppt::text_prop::parse_style_text_prop_atom(&style, 0);

        assert_eq!(paragraphs[0].characters_covered, 1);
        assert_eq!(characters[0].characters_covered, 1);
    }

    #[test]
    fn rejects_invalid_text_cf_values() {
        let mut invalid_size = TextPropsBuilder::new();
        invalid_size.add_paragraph(Paragraph::with_runs(vec![TextRun::new("x").size(0)]));
        assert!(invalid_size.build_style_text_prop().is_err());

        let mut invalid_scheme = TextPropsBuilder::new();
        invalid_scheme.add_paragraph(Paragraph::with_runs(vec![
            TextRun::new("x").color_scheme(8),
        ]));
        assert!(invalid_scheme.build_style_text_prop().is_err());

        let mut invalid_bullet = TextPropsBuilder::new();
        invalid_bullet.add_paragraph(Paragraph::new("x").with_bullet('😀'));
        assert!(invalid_bullet.build_style_text_prop().is_err());

        let mut invalid_position = TextPropsBuilder::new();
        invalid_position.add_paragraph(Paragraph::with_runs(vec![
            TextRun::new("x").baseline_position(101),
        ]));
        assert!(invalid_position.build_style_text_prop().is_err());

        let mut invalid_indent = TextPropsBuilder::new();
        invalid_indent.add_paragraph(Paragraph::new("x").indent_level(5));
        assert!(invalid_indent.build_style_text_prop().is_err());

        let mut invalid_bullet_size = TextPropsBuilder::new();
        invalid_bullet_size.add_paragraph(Paragraph::new("x").bullet_size(0));
        assert!(invalid_bullet_size.build_style_text_prop().is_err());

        let mut invalid_bullet_scheme = TextPropsBuilder::new();
        invalid_bullet_scheme.add_paragraph(Paragraph::new("x").bullet_color_scheme(8));
        assert!(invalid_bullet_scheme.build_style_text_prop().is_err());

        let mut invalid_pp9_run = TextPropsBuilder::new();
        invalid_pp9_run.add_paragraph(Paragraph::with_runs(vec![TextRun::new("x").pp9_run_id(16)]));
        assert!(invalid_pp9_run.build_style_text_prop().is_err());

        let mut reserved_style = TextRun::new("x");
        reserved_style.style.specified_mask = 0x0008;
        let mut invalid_reserved_style = TextPropsBuilder::new();
        invalid_reserved_style.add_paragraph(Paragraph::with_runs(vec![reserved_style]));
        assert!(invalid_reserved_style.build_style_text_prop().is_err());
    }

    #[test]
    fn character_flags_preserve_values_and_presence() {
        let run = TextRun::new("x")
            .bold_value(false)
            .italic()
            .underline_value(false)
            .shadow()
            .fe_hint(true)
            .kumi(false)
            .strikethrough(true)
            .embossed_value(false)
            .pp9_run_id(13)
            .font(65_535)
            .asian_font(65_534)
            .ansi_font(32_768)
            .symbol_font(60_000);
        let mut builder = TextPropsBuilder::new();
        builder.add_paragraph(Paragraph::with_runs(vec![run]));
        let style = builder.build_style_text_prop().unwrap();
        let (_, character_styles) =
            crate::ppt::text_prop::parse_style_text_prop_atom_strict(&style, 1).unwrap();
        assert_eq!(character_styles[0].property_mask & 0xFFFF, 0x3FB7);
        assert_eq!(character_styles[0].get_value("char.flags"), Some(0x3532));
        assert_eq!(character_styles[0].get_value("font.index"), Some(65_535));
        assert_eq!(
            character_styles[0].get_value("asian.font.index"),
            Some(65_534)
        );
        assert_eq!(
            character_styles[0].get_value("ansi.font.index"),
            Some(32_768)
        );
        assert_eq!(
            character_styles[0].get_value("symbol.font.index"),
            Some(60_000)
        );

        let text_record = crate::ppt::PptRecord {
            record_type: crate::consts::PptRecordType::TextBytesAtom,
            record_type_raw: 4008,
            version: 0,
            instance: 0,
            data_length: 1,
            data: b"x".to_vec(),
            children: Vec::new(),
        };
        let style_record = crate::ppt::PptRecord {
            record_type: crate::consts::PptRecordType::StyleTextPropAtom,
            record_type_raw: 4001,
            version: 0,
            instance: 0,
            data_length: style.len() as u32,
            data: style,
            children: Vec::new(),
        };
        let mut extractor = crate::ppt::TextRunExtractor::new();
        extractor
            .extract_from_records(&[text_record, style_record])
            .unwrap();
        let formatting = &extractor.runs()[0].formatting;
        assert_eq!(formatting.font_style_raw, Some(0x3532));
        assert_eq!(formatting.bold_explicit, Some(false));
        assert_eq!(formatting.italic_explicit, Some(true));
        assert_eq!(formatting.underline_explicit, Some(false));
        assert_eq!(formatting.shadow_explicit, Some(true));
        assert_eq!(formatting.fe_hint, Some(true));
        assert_eq!(formatting.kumi, Some(false));
        assert_eq!(formatting.legacy_strikethrough, Some(true));
        assert_eq!(formatting.embossed_explicit, Some(false));
        assert_eq!(formatting.pp9_run_id, Some(13));
        assert_eq!(formatting.font_index, Some(65_535));
        assert_eq!(formatting.asian_font_index, Some(65_534));
        assert_eq!(formatting.ansi_font_index, Some(32_768));
        assert_eq!(formatting.symbol_font_index, Some(60_000));
    }

    #[test]
    fn paragraph_properties_round_trip_in_spec_order() {
        let mut paragraph = Paragraph::new("x")
            .with_bullet('•')
            .bullet_font(65_535)
            .bullet_size(-24)
            .bullet_color_rgb(1, 2, 3)
            .align(TextAlign::Distributed)
            .line_spacing(120)
            .space_before(-10)
            .space_after(20)
            .indent_level(2)
            .default_tab_size(144)
            .tab_stops(vec![
                TabStop::new(-20, TabAlign::Center),
                TabStop::new(720, TabAlign::Decimal),
            ])
            .font_alignment(TextFontAlign::UpholdFixed)
            .character_wrap(true)
            .word_wrap(false)
            .overflow(true)
            .text_direction(TextDirection::RightToLeft);
        paragraph.left_margin = 720;
        paragraph.indent = -360;
        let mut builder = TextPropsBuilder::new();
        builder.add_paragraph(paragraph);

        let style = builder.build_style_text_prop().unwrap();
        let (paragraphs, _) =
            crate::ppt::text_prop::parse_style_text_prop_atom_strict(&style, 1).unwrap();
        let properties = &paragraphs[0];

        assert_eq!(properties.indent_level, 2);
        assert_eq!(properties.get_value("paragraph.flags"), Some(0x000F));
        assert_eq!(properties.get_value("bullet.char"), Some(0x2022));
        assert_eq!(properties.get_value("bullet.font"), Some(65_535));
        assert_eq!(properties.get_value("bullet.size"), Some(-24));
        assert_eq!(
            properties.get_value("bullet.color"),
            Some(0xFE03_0201u32 as i32)
        );
        assert_eq!(properties.get_value("alignment"), Some(4));
        assert_eq!(properties.get_value("linespacing"), Some(120));
        assert_eq!(properties.get_value("spacebefore"), Some(-10));
        assert_eq!(properties.get_value("spaceafter"), Some(20));
        assert_eq!(properties.get_value("text.offset"), Some(720));
        assert_eq!(properties.get_value("bullet.offset"), Some(-360));
        assert_eq!(properties.get_value("defaultTabSize"), Some(144));
        assert_eq!(properties.get_value("tabStops"), Some(2));
        assert_eq!(properties.tab_stops[0].position, -20);
        assert_eq!(properties.tab_stops[0].alignment, 1);
        assert_eq!(properties.tab_stops[1].position, 720);
        assert_eq!(properties.tab_stops[1].alignment, 3);
        assert_eq!(properties.get_value("fontAlignment"), Some(3));
        assert_eq!(properties.get_value("wrapFlags"), Some(5));
        assert_eq!(properties.get_value("textDirection"), Some(1));
    }

    #[test]
    fn paragraph_writer_preserves_explicit_false_and_default_values() {
        let paragraph = Paragraph::new("x")
            .bullet_enabled(false)
            .bullet_font_enabled(false)
            .bullet_color_enabled(false)
            .bullet_size_enabled(false)
            .align(TextAlign::Left)
            .line_spacing(100)
            .space_before(0)
            .space_after(0)
            .left_margin(0)
            .first_line_indent(0)
            .tab_stops(Vec::new())
            .character_wrap(false)
            .word_wrap(false)
            .overflow(false)
            .text_direction(TextDirection::LeftToRight);
        let mut builder = TextPropsBuilder::new();
        builder.add_paragraph(paragraph);

        let style = builder.build_style_text_prop().unwrap();
        let (paragraphs, _) =
            crate::ppt::text_prop::parse_style_text_prop_atom_strict(&style, 1).unwrap();
        let properties = &paragraphs[0];

        assert_eq!(properties.get_value("paragraph.flags"), Some(0));
        assert_eq!(properties.get_value("alignment"), Some(0));
        assert_eq!(properties.get_value("linespacing"), Some(100));
        assert_eq!(properties.get_value("spacebefore"), Some(0));
        assert_eq!(properties.get_value("spaceafter"), Some(0));
        assert_eq!(properties.get_value("text.offset"), Some(0));
        assert_eq!(properties.get_value("bullet.offset"), Some(0));
        assert_eq!(properties.get_value("tabStops"), Some(0));
        assert!(properties.tab_stops.is_empty());
        assert_eq!(properties.get_value("wrapFlags"), Some(0));
        assert_eq!(properties.get_value("textDirection"), Some(0));
    }

    #[test]
    fn test_font_entity() {
        let font = FontEntity::arial();
        let data = font.build();
        assert_eq!(data.len(), 68);
        // Check "Arial" in UTF-16LE
        assert_eq!(data[0], b'A');
        assert_eq!(data[1], 0);
    }
}
