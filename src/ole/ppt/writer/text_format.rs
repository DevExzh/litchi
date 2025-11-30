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
    /// Has style text prop atom present
    pub const HAS_BULLET: u32 = 0x0001;
    /// Has bullet character
    pub const BULLET_CHAR: u32 = 0x0002;
    /// Has bullet font
    pub const BULLET_FONT: u32 = 0x0004;
    /// Has bullet size
    pub const BULLET_SIZE: u32 = 0x0008;
    /// Has bullet color
    pub const BULLET_COLOR: u32 = 0x0010;
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
    /// Emboss
    pub const EMBOSS: u32 = 0x0200;
    /// Font reference present
    pub const FONT_REF: u32 = 0x0001_0000;
    /// Font size present
    pub const FONT_SIZE: u32 = 0x0002_0000;
    /// Font color present
    pub const FONT_COLOR: u32 = 0x0004_0000;
    /// Position (superscript/subscript) present
    pub const POSITION: u32 = 0x0008_0000;
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
    /// Strikethrough
    pub strikethrough: bool,
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
            strikethrough: false,
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
            strikethrough: false,
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
            strikethrough: false,
        }
    }

    /// Convert to mask value for TextCFException
    pub fn to_mask(&self) -> u32 {
        let mut mask = 0u32;
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
        if self.emboss {
            mask |= char_mask::EMBOSS;
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
        if self.emboss {
            flags |= 0x0200;
        }
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
            // Scheme color reference
            0xFE00_0000 | (self.scheme_index as u32)
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
        }
    }

    /// Set bold
    pub fn bold(mut self) -> Self {
        self.style.bold = true;
        self
    }

    /// Set italic
    pub fn italic(mut self) -> Self {
        self.style.italic = true;
        self
    }

    /// Set underline
    pub fn underline(mut self) -> Self {
        self.style.underline = true;
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

    /// Set font index
    pub fn font(mut self, index: u16) -> Self {
        self.font_index = index;
        self
    }

    /// Get character count (including any trailing CR/LF)
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
    /// Bullet character (if any)
    pub bullet_char: Option<char>,
    /// Bullet color
    pub bullet_color: Option<TextColor>,
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
            bullet_char: None,
            bullet_color: None,
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
            bullet_char: None,
            bullet_color: None,
        }
    }

    /// Set alignment
    pub fn align(mut self, alignment: TextAlign) -> Self {
        self.alignment = alignment;
        self
    }

    /// Center align
    pub fn center(mut self) -> Self {
        self.alignment = TextAlign::Center;
        self
    }

    /// Right align
    pub fn right(mut self) -> Self {
        self.alignment = TextAlign::Right;
        self
    }

    /// Set line spacing (percent)
    pub fn line_spacing(mut self, percent: i16) -> Self {
        self.line_spacing = percent;
        self
    }

    /// Set space before
    pub fn space_before(mut self, units: i16) -> Self {
        self.space_before = units;
        self
    }

    /// Set space after
    pub fn space_after(mut self, units: i16) -> Self {
        self.space_after = units;
        self
    }

    /// Add bullet
    pub fn with_bullet(mut self, ch: char) -> Self {
        self.bullet_char = Some(ch);
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

    /// Build TextCharsAtom (UTF-16LE text)
    /// Adds CR between paragraphs and after the last paragraph (for StyleTextPropAtom compatibility)
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
    pub fn build_style_text_prop(&self) -> Vec<u8> {
        let mut data = Vec::new();

        // Paragraph properties (TextPFRun entries)
        // Each paragraph covers its runs + CR separator (except last paragraph gets +1 for terminator)
        for para in &self.paragraphs {
            let para_text_len = para.runs_char_count();

            // Character count: text + CR (or +1 for last paragraph terminator)
            // +1 for either CR separator or implicit terminating character
            let char_count = para_text_len + 1;
            data.extend_from_slice(&char_count.to_le_bytes());

            // Indent level (0 for top-level)
            data.extend_from_slice(&0u16.to_le_bytes());

            // Build mask based on what properties are set
            let mut mask = 0u32;
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
            if para.bullet_char.is_some() {
                mask |= para_mask::HAS_BULLET | para_mask::BULLET_CHAR;
            }

            data.extend_from_slice(&mask.to_le_bytes());

            // Write properties according to mask
            if mask & para_mask::HAS_BULLET != 0 {
                data.extend_from_slice(&1u16.to_le_bytes()); // hasBullet = true
            }
            if mask & para_mask::BULLET_CHAR != 0 {
                let ch = para.bullet_char.unwrap_or('•') as u16;
                data.extend_from_slice(&ch.to_le_bytes());
            }
            if mask & para_mask::LEFT_MARGIN != 0 {
                data.extend_from_slice(&para.left_margin.to_le_bytes());
            }
            if mask & para_mask::INDENT != 0 {
                data.extend_from_slice(&para.indent.to_le_bytes());
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
        }

        // Character properties (TextCFRun entries)
        // Write one entry per run. The last run in each paragraph gets +1 for CR/terminator.
        for para in &self.paragraphs {
            let num_runs = para.runs.len();

            for (run_idx, run) in para.runs.iter().enumerate() {
                let is_last_run = run_idx == num_runs - 1;

                // Character count for this run
                // Last run of last paragraph gets +1 for terminator
                // Last run of non-last paragraph gets +1 for CR separator
                let char_count = if is_last_run {
                    run.char_count() + 1
                } else {
                    run.char_count()
                };
                data.extend_from_slice(&char_count.to_le_bytes());

                // Build mask
                let mut mask = run.style.to_mask();
                mask |= char_mask::FONT_SIZE; // Always include font size
                mask |= char_mask::FONT_COLOR; // Always include color
                mask |= char_mask::FONT_REF; // Always include font reference

                data.extend_from_slice(&mask.to_le_bytes());

                // Font style flags (only if any flags are set)
                if mask & 0xFFFF != 0 {
                    let flags = run.style.to_flags();
                    data.extend_from_slice(&flags.to_le_bytes());
                }

                // Font reference (if font_ref bit is set)
                data.extend_from_slice(&run.font_index.to_le_bytes());

                // Font size (in half-points, so multiply by 2)
                let half_points = run.font_size * 2;
                data.extend_from_slice(&half_points.to_le_bytes());

                // Color (POI format: R | G<<8 | B<<16 | 0xFE<<24)
                let color = run.color.to_ppt_color();
                data.extend_from_slice(&color.to_le_bytes());
            }
        }

        data
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

        // Scheme color with alpha
        let scheme = TextColor::scheme(4);
        assert_eq!(scheme.to_ppt_color(), 0xFE000004);
    }

    #[test]
    fn test_text_run() {
        let run = TextRun::new("Hello").bold().size(24);
        assert!(run.style.bold);
        assert_eq!(run.font_size, 24);
        assert_eq!(run.char_count(), 5);
    }

    #[test]
    fn test_paragraph() {
        let para = Paragraph::new("Test").center();
        assert_eq!(para.alignment, TextAlign::Center);
        assert_eq!(para.char_count(), 5); // 4 chars + 1 end marker
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
