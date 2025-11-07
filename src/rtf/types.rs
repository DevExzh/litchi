//! RTF document type definitions.

use super::border::{Borders, Shading};
use std::borrow::Cow;
use std::num::NonZeroU16;

/// Font reference (index into font table).
pub type FontRef = u16;

/// Color reference (index into color table).
pub type ColorRef = u16;

/// RTF color representation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct Color {
    /// Red component (0-255)
    pub red: u8,
    /// Green component (0-255)
    pub green: u8,
    /// Blue component (0-255)
    pub blue: u8,
}

impl Color {
    /// Create a new color.
    #[inline]
    pub const fn new(red: u8, green: u8, blue: u8) -> Self {
        Self { red, green, blue }
    }

    /// Black color.
    #[inline]
    pub const fn black() -> Self {
        Self::new(0, 0, 0)
    }

    /// White color.
    #[inline]
    pub const fn white() -> Self {
        Self::new(255, 255, 255)
    }
}

/// Color table containing document colors.
#[derive(Debug, Clone)]
pub struct ColorTable {
    colors: Vec<Color>,
}

impl ColorTable {
    /// Create a new color table.
    #[inline]
    pub fn new() -> Self {
        Self { colors: Vec::new() }
    }

    /// Add a color to the table and return its index.
    #[inline]
    pub fn add(&mut self, color: Color) -> ColorRef {
        let index = self.colors.len() as ColorRef;
        self.colors.push(color);
        index
    }

    /// Get a color by reference.
    #[inline]
    pub fn get(&self, color_ref: ColorRef) -> Option<&Color> {
        self.colors.get(color_ref as usize)
    }

    /// Get all colors in the table.
    #[inline]
    pub fn colors(&self) -> &[Color] {
        &self.colors
    }
}

impl Default for ColorTable {
    fn default() -> Self {
        Self::new()
    }
}

/// Font family categories.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum FontFamily {
    /// Nil (unknown or default)
    #[default]
    Nil,
    /// Roman (serif) fonts
    Roman,
    /// Swiss (sans-serif) fonts
    Swiss,
    /// Modern (monospace) fonts
    Modern,
    /// Script fonts
    Script,
    /// Decorative fonts
    Decor,
    /// Technical, symbol, and mathematical fonts
    Tech,
}

/// Font definition.
#[derive(Debug, Clone)]
pub struct Font<'a> {
    /// Font name
    pub name: Cow<'a, str>,
    /// Font family category
    pub family: FontFamily,
    /// Character set (Windows codepage)
    pub charset: u8,
}

impl<'a> Font<'a> {
    /// Create a new font.
    #[inline]
    pub fn new(name: Cow<'a, str>, family: FontFamily, charset: u8) -> Self {
        Self {
            name,
            family,
            charset,
        }
    }
}

/// Font table containing document fonts.
#[derive(Debug, Clone)]
pub struct FontTable<'a> {
    pub(crate) fonts: Vec<Font<'a>>,
}

impl<'a> FontTable<'a> {
    /// Create a new font table.
    #[inline]
    pub fn new() -> Self {
        Self { fonts: Vec::new() }
    }

    /// Add a font to the table at a specific index.
    #[inline]
    pub fn insert(&mut self, index: FontRef, font: Font<'a>) {
        // Ensure the vector is large enough
        if index as usize >= self.fonts.len() {
            self.fonts.resize(
                (index as usize) + 1,
                Font::new(Cow::Borrowed(""), FontFamily::Nil, 0),
            );
        }
        self.fonts[index as usize] = font;
    }

    /// Get a font by reference.
    #[inline]
    pub fn get(&self, font_ref: FontRef) -> Option<&Font<'a>> {
        self.fonts.get(font_ref as usize)
    }

    /// Get all fonts in the table.
    #[inline]
    pub fn fonts(&self) -> &[Font<'a>] {
        &self.fonts
    }
}

impl<'a> Default for FontTable<'a> {
    fn default() -> Self {
        Self::new()
    }
}

/// Text alignment.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum Alignment {
    /// Left-aligned
    #[default]
    Left,
    /// Right-aligned
    Right,
    /// Centered
    Center,
    /// Justified
    Justify,
}

/// Spacing information for paragraphs.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct Spacing {
    /// Space before paragraph (in twips, 1/20th of a point)
    pub before: i32,
    /// Space after paragraph (in twips)
    pub after: i32,
    /// Line spacing (in twips)
    pub line: i32,
    /// Line spacing multiplier
    pub line_multiple: bool,
}

/// Indentation information for paragraphs.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct Indentation {
    /// Left indent (in twips)
    pub left: i32,
    /// Right indent (in twips)
    pub right: i32,
    /// First line indent (in twips)
    pub first_line: i32,
}

/// Paragraph properties.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct Paragraph {
    /// Text alignment
    pub alignment: Alignment,
    /// Spacing
    pub spacing: Spacing,
    /// Indentation
    pub indentation: Indentation,
    /// Borders
    pub borders: Borders,
    /// Shading/background
    pub shading: Shading,
    /// Keep paragraph on one page
    pub keep_together: bool,
    /// Keep with next paragraph
    pub keep_next: bool,
    /// Page break before
    pub page_break_before: bool,
    /// Widow/orphan control
    pub widow_control: bool,
}

/// Underline style
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum UnderlineStyle {
    /// No underline
    #[default]
    None,
    /// Single underline
    Single,
    /// Double underline
    Double,
    /// Dotted underline
    Dotted,
    /// Dashed underline
    Dashed,
    /// Dash-dot underline
    DashDot,
    /// Dash-dot-dot underline
    DashDotDot,
    /// Word-only underline
    Words,
    /// Thick underline
    Thick,
    /// Wave underline
    Wave,
}

/// Character formatting properties.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Formatting {
    /// Font reference
    pub font_ref: FontRef,
    /// Font size in half-points
    pub font_size: NonZeroU16,
    /// Color reference
    pub color_ref: ColorRef,
    /// Background/highlight color reference
    pub highlight_color: Option<ColorRef>,
    /// Bold
    pub bold: bool,
    /// Italic
    pub italic: bool,
    /// Underline style
    pub underline: UnderlineStyle,
    /// Strikethrough
    pub strike: bool,
    /// Double strikethrough
    pub double_strike: bool,
    /// Superscript
    pub superscript: bool,
    /// Subscript
    pub subscript: bool,
    /// Small caps
    pub smallcaps: bool,
    /// All caps
    pub all_caps: bool,
    /// Hidden text
    pub hidden: bool,
    /// Outline
    pub outline: bool,
    /// Shadow
    pub shadow: bool,
    /// Emboss
    pub emboss: bool,
    /// Engrave (imprint)
    pub imprint: bool,
    /// Character spacing (in twips)
    pub char_spacing: i32,
    /// Horizontal scaling (percentage)
    pub char_scale: i32,
    /// Kerning (in half-points)
    pub kerning: i32,
}

impl Default for Formatting {
    fn default() -> Self {
        Self {
            font_ref: 0,
            // SAFETY: 24 (12pt) is non-zero
            font_size: unsafe { NonZeroU16::new_unchecked(24) },
            color_ref: 0,
            highlight_color: None,
            bold: false,
            italic: false,
            underline: UnderlineStyle::default(),
            strike: false,
            double_strike: false,
            superscript: false,
            subscript: false,
            smallcaps: false,
            all_caps: false,
            hidden: false,
            outline: false,
            shadow: false,
            emboss: false,
            imprint: false,
            char_spacing: 0,
            char_scale: 100,
            kerning: 0,
        }
    }
}

/// A text run with formatting.
#[derive(Debug, Clone)]
pub struct Run<'a> {
    /// Text content
    pub text: Cow<'a, str>,
    /// Character formatting
    pub formatting: Formatting,
}

impl<'a> Run<'a> {
    /// Create a new run.
    #[inline]
    pub fn new(text: Cow<'a, str>, formatting: Formatting) -> Self {
        Self { text, formatting }
    }

    /// Get the text content.
    #[inline]
    pub fn text(&self) -> &str {
        &self.text
    }

    /// Check if this run is bold.
    #[inline]
    pub fn bold(&self) -> Option<bool> {
        Some(self.formatting.bold)
    }

    /// Check if this run is italic.
    #[inline]
    pub fn italic(&self) -> Option<bool> {
        Some(self.formatting.italic)
    }

    /// Check if this run has strikethrough.
    #[inline]
    pub fn strikethrough(&self) -> Option<bool> {
        Some(self.formatting.strike || self.formatting.double_strike)
    }

    /// Check if this run has underline.
    #[inline]
    pub fn underline(&self) -> bool {
        !matches!(self.formatting.underline, UnderlineStyle::None)
    }

    /// Get the vertical position of this run (superscript/subscript).
    #[inline]
    pub fn vertical_position(&self) -> Option<crate::common::style::text::pos::VerticalPosition> {
        if self.formatting.superscript {
            Some(crate::common::style::text::pos::VerticalPosition::Superscript)
        } else if self.formatting.subscript {
            Some(crate::common::style::text::pos::VerticalPosition::Subscript)
        } else {
            None
        }
    }
}

/// A styled block of text with paragraph and character formatting.
#[derive(Debug, Clone)]
pub struct StyleBlock<'a> {
    /// Paragraph properties
    pub paragraph: Paragraph,
    /// Character formatting
    pub formatting: Formatting,
    /// Text content
    pub text: Cow<'a, str>,
}

impl<'a> StyleBlock<'a> {
    /// Create a new style block.
    #[inline]
    pub fn new(text: Cow<'a, str>, formatting: Formatting, paragraph: Paragraph) -> Self {
        Self {
            text,
            formatting,
            paragraph,
        }
    }

    /// Get the text content.
    #[inline]
    pub fn text(&self) -> &str {
        &self.text
    }
}

/// A paragraph with content (runs).
///
/// This represents a paragraph in the unified Document API, containing
/// both paragraph properties and the runs that make up the paragraph content.
#[derive(Debug, Clone)]
pub struct ParagraphContent<'a> {
    /// Paragraph properties (alignment, spacing, indentation)
    pub properties: Paragraph,
    /// Runs contained in this paragraph
    pub runs: Vec<Run<'a>>,
}

/// Document element - either a paragraph or a table.
///
/// This enum is used by the `elements()` method to represent
/// the mixed content of an RTF document in sequential order.
#[derive(Debug, Clone)]
pub enum DocumentElement<'a> {
    /// A paragraph with formatted runs
    Paragraph(ParagraphContent<'a>),
    /// A table with rows and cells
    Table(super::table::Table<'a>),
}

impl<'a> ParagraphContent<'a> {
    /// Create a new paragraph with content.
    #[inline]
    pub fn new(properties: Paragraph, runs: Vec<Run<'a>>) -> Self {
        Self { properties, runs }
    }

    /// Get the text content of the paragraph.
    #[inline]
    pub fn text(&self) -> String {
        self.runs.iter().map(|r| r.text.as_ref()).collect()
    }

    /// Get the runs in this paragraph.
    #[inline]
    pub fn runs(&self) -> &[Run<'a>] {
        &self.runs
    }
}
