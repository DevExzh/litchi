//! Semantic `WordprocessingML` paragraph and run values.

use crate::UnderlineStyle;
use crate::color::Theme;
use crate::font::OpenType;
use crate::hyperlink::Hyperlink;
use crate::image::InlineImage;
use crate::run_effects::Effects;
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

/// An ordered direct child of a `WordprocessingML` paragraph.
///
/// Runs are exposed through their typed semantic value. Every other paragraph
/// child is retained as an inert exact-XML fallback until a focused semantic
/// owner can represent it without loss. This keeps hyperlinks, fields,
/// revisions, content controls, bookmarks, Office Math, and future extension
/// elements visible in document order instead of silently dropping them.
#[derive(Debug, Clone)]
#[non_exhaustive]
pub enum Inline {
    /// A typed `WordprocessingML` run.
    Run(Box<Run>),
    /// A relationship-resolved hyperlink and its ordered run children.
    Hyperlink(Box<InlineHyperlink>),
    /// A supported or future paragraph child retained byte-for-byte.
    Unknown(Box<OpaqueInline>),
}

/// A relationship-resolved direct `<w:hyperlink>` paragraph child.
///
/// The public value deliberately contains the resolved [`Hyperlink`] rather
/// than its package-local relationship identifier. Its direct runs retain
/// their formatting and expose ordered content through [`Run::contents`].
#[derive(Debug, Clone)]
pub struct InlineHyperlink {
    link: Hyperlink,
    runs: Vec<Run>,
    target_frame: Option<String>,
    document_location: Option<String>,
    has_unmodeled_content: bool,
}

impl InlineHyperlink {
    pub(crate) fn new(
        link: Hyperlink,
        runs: Vec<Run>,
        target_frame: Option<String>,
        document_location: Option<String>,
        has_unmodeled_content: bool,
    ) -> Self {
        Self {
            link,
            runs,
            target_frame,
            document_location,
            has_unmodeled_content,
        }
    }

    /// Borrow the resolved hyperlink value.
    #[must_use]
    pub const fn link(&self) -> &Hyperlink {
        &self.link
    }

    /// Borrow the hyperlink's direct runs in source order.
    #[must_use]
    pub fn runs(&self) -> &[Run] {
        &self.runs
    }

    /// Return the optional target frame (`w:tgtFrame`).
    #[must_use]
    pub fn target_frame(&self) -> Option<&str> {
        self.target_frame.as_deref()
    }

    /// Return the optional document location (`w:docLocation`).
    #[must_use]
    pub fn document_location(&self) -> Option<&str> {
        self.document_location.as_deref()
    }

    /// Whether the hyperlink also contains direct children other than runs.
    ///
    /// Consumers that require a complete semantic projection should refuse
    /// the hyperlink when this returns `true`; the paragraph's lossless
    /// [`Inline::Unknown`] projection remains available through
    /// [`Paragraph::inlines`] for exact preservation.
    #[must_use]
    pub const fn has_unmodeled_content(&self) -> bool {
        self.has_unmodeled_content
    }
}

/// A paragraph child whose semantics are not modeled by [`Inline`].
///
/// The payload is inert. Reading it does not resolve relationships, evaluate
/// fields, apply revisions, activate controls, or execute embedded content.
#[derive(Debug, Clone)]
pub struct OpaqueInline {
    source: Arc<Vec<u8>>,
    start: u32,
    length: u32,
    word_hyperlink: bool,
}

impl OpaqueInline {
    pub(crate) const fn from_arc_range(
        source: Arc<Vec<u8>>,
        start: u32,
        length: u32,
        word_hyperlink: bool,
    ) -> Self {
        Self {
            source,
            start,
            length,
            word_hyperlink,
        }
    }

    pub(crate) const fn is_word_hyperlink(&self) -> bool {
        self.word_hyperlink
    }

    /// Borrow the retained paragraph child exactly as it appeared in the
    /// active paragraph XML.
    #[must_use]
    pub fn xml_bytes(&self) -> &[u8] {
        let Ok(start) = usize::try_from(self.start) else {
            return &[];
        };
        let Ok(length) = usize::try_from(self.length) else {
            return &[];
        };
        let Some(end) = start.checked_add(length) else {
            return &[];
        };
        self.source.get(start..end).unwrap_or_default()
    }
}

/// A direct run child whose semantics are not modeled by [`RunContent`].
#[derive(Debug, Clone)]
pub struct OpaqueRunContent {
    source: Arc<Vec<u8>>,
    start: u32,
    length: u32,
}

impl OpaqueRunContent {
    pub(crate) const fn from_arc_range(source: Arc<Vec<u8>>, start: u32, length: u32) -> Self {
        Self {
            source,
            start,
            length,
        }
    }

    /// Borrow the retained run child exactly as it appeared in source XML.
    #[must_use]
    pub fn xml_bytes(&self) -> &[u8] {
        let Ok(start) = usize::try_from(self.start) else {
            return &[];
        };
        let Ok(length) = usize::try_from(self.length) else {
            return &[];
        };
        let Some(end) = start.checked_add(length) else {
            return &[];
        };
        self.source.get(start..end).unwrap_or_default()
    }
}

impl Paragraph {
    /// Create a new Paragraph from XML bytes (owned).
    ///
    /// # Arguments
    ///
    /// * `xml_bytes` - The XML content of the `<w:p>` element
    #[inline]
    #[must_use]
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
    #[must_use]
    pub fn from_slice(slice: XmlSlice) -> Self {
        Self {
            xml_data: XmlData::Shared(slice),
        }
    }

    /// Create a Paragraph from an `Arc<Vec<u8>>` and byte range.
    ///
    /// This is a convenience method for arena-based parsing.
    #[inline]
    #[must_use]
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
/// These tokens are defined by the `WordprocessingML` `ST_LineSpacingRule`
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
    /// Parse the exact `WordprocessingML` token for `w:lineRule`.
    #[must_use]
    pub fn from_xml(value: &str) -> Option<Self> {
        match value {
            "auto" => Some(Self::Auto),
            "exact" => Some(Self::Exact),
            "atLeast" => Some(Self::AtLeast),
            _ => None,
        }
    }

    /// Return the exact `WordprocessingML` token for `w:lineRule`.
    #[must_use]
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
#[derive(Debug, Clone, Default)]
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
    /// Typed Word 2010 visual effects attached directly to the run.
    pub effects: Effects,
    /// Typed Word 2010 OpenType features attached directly to the run.
    pub open_type: OpenType,
}

/// A direct color applied to a `WordprocessingML` underline.
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

    #[inline]
    pub(super) fn get_or_create_arc(&self) -> (Arc<Vec<u8>>, u32) {
        match self {
            RunXmlData::Owned(bytes) => (Arc::new(bytes.clone()), 0),
            RunXmlData::Shared(slice) => (slice.arc(), slice.start()),
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

/// An ordered direct child of a `WordprocessingML` run.
#[derive(Debug, Clone)]
#[non_exhaustive]
pub enum RunContent {
    /// Text from one `<w:t>` element.
    Text(String),
    /// A tab character (`<w:tab>`).
    Tab,
    /// An explicit line, page, or column break.
    Break(RunBreak),
    /// A carriage return (`<w:cr>`).
    CarriageReturn,
    /// A non-breaking hyphen (`<w:noBreakHyphen>`).
    NoBreakHyphen,
    /// A discretionary soft hyphen (`<w:softHyphen>`).
    SoftHyphen,
    /// A relationship-backed inline image.
    Image(Box<InlineImage>),
    /// A footnote reference by non-negative Word note identifier.
    FootnoteReference(u32),
    /// An endnote reference by non-negative Word note identifier.
    EndnoteReference(u32),
    /// The automatic footnote number marker inside a footnote definition.
    FootnoteMark,
    /// The automatic endnote number marker inside an endnote definition.
    EndnoteMark,
    /// A supported or future run child retained byte-for-byte.
    Unknown(Box<OpaqueRunContent>),
}

/// The semantic type of an explicit `WordprocessingML` run break.
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
    #[must_use]
    pub fn new(xml_bytes: Vec<u8>) -> Self {
        Self {
            xml_data: RunXmlData::Owned(xml_bytes),
        }
    }

    /// Create a Run from a shared XML slice (zero-copy).
    #[inline]
    #[must_use]
    pub fn from_slice(slice: XmlSlice) -> Self {
        Self {
            xml_data: RunXmlData::Shared(slice),
        }
    }

    /// Get the raw XML bytes.
    #[inline]
    pub(crate) fn xml_bytes(&self) -> &[u8] {
        self.xml_data.as_bytes()
    }

    pub(crate) fn replace_xml(&mut self, xml_bytes: Vec<u8>) {
        self.xml_data = RunXmlData::Owned(xml_bytes);
    }
}
