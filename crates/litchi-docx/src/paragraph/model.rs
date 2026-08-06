//! Semantic WordprocessingML paragraph and run values.

use crate::UnderlineStyle;
use crate::color::Theme;
use litchi_core::{VerticalPosition, XmlSlice};
use std::sync::Arc;

/// Internal storage for paragraph XML data.
/// Supports both owned data (for standalone parsing) and shared slices (for arena-based parsing).
#[derive(Debug, Clone)]
pub(super) enum XmlData {
    /// Owned data for standalone paragraphs
    Owned(Box<[u8]>),
    /// Shared slice into an arena for zero-copy batch parsing
    Shared(XmlSlice),
}

impl XmlData {
    #[inline]
    pub(super) fn as_bytes(&self) -> &[u8] {
        match self {
            XmlData::Owned(bytes) => bytes,
            XmlData::Shared(slice) => slice.as_bytes(),
        }
    }

    /// Get or create an Arc for this data.
    /// If already shared, returns the existing Arc (cheap clone).
    /// If owned, creates a new Arc (allocates once).
    #[inline]
    pub(super) fn get_or_create_arc(&self) -> (Arc<Vec<u8>>, u32) {
        match self {
            XmlData::Owned(bytes) => (Arc::new(bytes.to_vec()), 0),
            XmlData::Shared(slice) => (slice.arc(), slice.start()),
        }
    }
}

/// A paragraph in a Word document.
///
/// Represents a `<w:p>` element. Paragraphs contain runs which in turn
/// contain the actual text and formatting.
///
/// # Example
///
/// ```rust,ignore
/// for para in document.paragraphs()? {
///     println!("Paragraph text: {}", para.text());
///     for run in para.runs()? {
///         println!("  Run: {} (bold: {:?})", run.text(), run.bold());
///     }
/// }
/// ```
#[derive(Debug, Clone)]
pub struct Paragraph {
    /// The raw XML bytes for this paragraph
    pub(super) xml_data: XmlData,
}

impl Paragraph {
    /// Create a new Paragraph from XML bytes (owned).
    ///
    /// # Arguments
    ///
    /// * `xml_bytes` - The XML content of the `<w:p>` element
    #[inline]
    pub fn new(xml_bytes: Vec<u8>) -> Self {
        Self {
            xml_data: XmlData::Owned(xml_bytes.into_boxed_slice()),
        }
    }

    /// Create a new Paragraph from a shared XML slice (zero-copy).
    ///
    /// This is used for arena-based parsing where all element XMLs are stored
    /// in a single contiguous buffer.
    #[inline]
    pub fn from_slice(slice: XmlSlice) -> Self {
        Self {
            xml_data: XmlData::Shared(slice),
        }
    }

    /// Create a Paragraph from an `Arc<Vec<u8>>` and byte range.
    ///
    /// This is a convenience method for arena-based parsing.
    #[inline]
    pub fn from_arc_range(arena: Arc<Vec<u8>>, start: u32, len: u32) -> Self {
        Self::from_slice(XmlSlice::new(arena, start, len))
    }

    /// Get the raw XML bytes.
    #[inline]
    pub(super) fn xml_bytes(&self) -> &[u8] {
        self.xml_data.as_bytes()
    }
}

/// The `w:lineRule` interpretation for a paragraph's `w:line` value.
///
/// These tokens are defined by the WordprocessingML `ST_LineSpacingRule`
/// simple type. `None` on [`ParagraphSpacing::line_rule`] means that the
/// source omitted the optional attribute; consumers use their normal default.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LineSpacingRule {
    /// Interpret `w:line` in 240ths of a line.
    Auto,
    /// Interpret `w:line` as twentieths of a point and clip if necessary.
    Exact,
    /// Interpret `w:line` as a minimum height in twentieths of a point.
    AtLeast,
}

impl LineSpacingRule {
    /// Parse the exact WordprocessingML token for `w:lineRule`.
    pub fn from_xml(value: &str) -> Option<Self> {
        match value {
            "auto" => Some(Self::Auto),
            "exact" => Some(Self::Exact),
            "atLeast" => Some(Self::AtLeast),
            _ => None,
        }
    }

    /// Return the exact WordprocessingML token for `w:lineRule`.
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Auto => "auto",
            Self::Exact => "exact",
            Self::AtLeast => "atLeast",
        }
    }
}

/// Typed direct spacing attributes from a paragraph's `<w:spacing>` element.
///
/// `before` and `after` use non-negative twips (`ST_TwipsMeasure`), while
/// `line` uses the signed `ST_SignedTwipsMeasure` value. The latter is in
/// 240ths of a line for [`LineSpacingRule::Auto`] and twentieths of a point
/// for [`LineSpacingRule::Exact`] or [`LineSpacingRule::AtLeast`]. The line
/// unit and automatic before/after fields mirror the other optional
/// `CT_Spacing` attributes instead of being discarded during an edit.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct ParagraphSpacing {
    /// Explicit spacing before the paragraph, in twips.
    pub before: Option<u64>,
    /// Spacing before the paragraph, in hundredths of a line.
    pub before_lines: Option<i32>,
    /// Whether the consumer should determine spacing before automatically.
    pub before_auto_spacing: Option<bool>,
    /// Explicit spacing after the paragraph, in twips.
    pub after: Option<u64>,
    /// Spacing after the paragraph, in hundredths of a line.
    pub after_lines: Option<i32>,
    /// Whether the consumer should determine spacing after automatically.
    pub after_auto_spacing: Option<bool>,
    /// Vertical line spacing value, interpreted according to `line_rule`.
    pub line: Option<i32>,
    /// Optional interpretation of `line`.
    pub line_rule: Option<LineSpacingRule>,
}

/// Cached formatting properties for a Run.
///
/// This struct stores all commonly accessed formatting properties
/// to avoid repeated XML parsing.
#[derive(Debug, Clone, Copy, Default)]
pub struct RunProperties {
    /// Whether the run is bold
    pub bold: Option<bool>,
    /// Whether the run is italic
    pub italic: Option<bool>,
    /// Whether the run is strikethrough
    pub strikethrough: Option<bool>,
    /// Explicit underline pattern, including [`UnderlineStyle::None`]
    pub underline: Option<UnderlineStyle>,
    /// Vertical position (superscript/subscript)
    pub vertical_position: Option<VerticalPosition>,
}

/// A direct color applied to a WordprocessingML underline.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RunUnderlineColor {
    /// Automatic color selected by the consumer.
    Auto,
    /// Explicit red, green, and blue components.
    Rgb([u8; 3]),
}

/// Complete direct underline formatting for a run.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RunUnderline {
    /// Underline pattern, including the explicit `none` value.
    pub style: UnderlineStyle,
    /// Direct automatic or RGB color.
    pub color: Option<RunUnderlineColor>,
    /// Theme color used instead of, or to transform, the direct color.
    pub theme_color: Option<Theme>,
    /// Theme tint transform byte.
    pub theme_tint: Option<u8>,
    /// Theme shade transform byte.
    pub theme_shade: Option<u8>,
}

/// Internal storage for run XML data (same pattern as Paragraph).
#[derive(Debug, Clone)]
pub(super) enum RunXmlData {
    Owned(Vec<u8>),
    Shared(XmlSlice),
}

impl RunXmlData {
    #[inline]
    pub(super) fn as_bytes(&self) -> &[u8] {
        match self {
            RunXmlData::Owned(bytes) => bytes,
            RunXmlData::Shared(slice) => slice.as_bytes(),
        }
    }
}

/// A run within a paragraph.
///
/// Represents a `<w:r>` element. A run is a region of text with a single
/// set of formatting properties.
///
/// # Example
///
/// ```rust,ignore
/// let run = runs[0];
/// println!("Text: {}", run.text()?);
/// println!("Bold: {:?}", run.bold()?);
/// println!("Italic: {:?}", run.italic()?);
///
/// // Check for embedded formulas
/// if let Some(omml) = run.omml_formula()? {
///     println!("OMML formula: {}", omml);
/// }
/// ```
#[derive(Debug, Clone)]
pub struct Run {
    /// The raw XML data for this run
    pub(super) xml_data: RunXmlData,
}

/// The semantic type of an explicit WordprocessingML run break.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum RunBreakType {
    /// A normal line break within the current text flow.
    #[default]
    TextWrapping,
    /// A page break.
    Page,
    /// A column break.
    Column,
}

/// How text wrapping resumes after a line break around floating objects.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum RunBreakClear {
    /// Resume on the next line without clearing either side.
    #[default]
    None,
    /// Resume when the left side is clear.
    Left,
    /// Resume when the right side is clear.
    Right,
    /// Resume when both sides are clear.
    All,
}

/// A typed `<w:br>` element contained in a run.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct RunBreak {
    /// Break type; omitted `w:type` defaults to text wrapping.
    pub break_type: RunBreakType,
    /// Wrapping-clear behavior; omitted `w:clear` defaults to none.
    pub clear: RunBreakClear,
}

impl Run {
    /// Create a new Run from XML bytes (owned).
    pub fn new(xml_bytes: Vec<u8>) -> Self {
        Self {
            xml_data: RunXmlData::Owned(xml_bytes),
        }
    }

    /// Create a Run from a shared XML slice (zero-copy).
    #[inline]
    pub fn from_slice(slice: XmlSlice) -> Self {
        Self {
            xml_data: RunXmlData::Shared(slice),
        }
    }

    /// Get the raw XML bytes.
    #[inline]
    pub(super) fn xml_bytes(&self) -> &[u8] {
        self.xml_data.as_bytes()
    }
}
