//! Semantic `WordprocessingML` paragraph and run values.

use crate::UnderlineStyle;
use crate::color::Theme;
use crate::document::{ManagedAdmission, ManagedNamespaceAdmission, ManagedParserAdmission};
use crate::error::{Error, Result};
use crate::font::OpenType;
use crate::hyperlink::Hyperlink;
use crate::image::InlineImage;
use crate::run_effects::Effects;
use litchi_core::{VerticalPosition, XmlSlice};
use litchi_opc::{PartData, SourceXmlPart};
use quick_xml::name::NamespaceResolver;
use quick_xml::reader::NsReader;
use std::ops::Deref;
use std::sync::Arc;

/// Immutable XML storage retained by a semantic view.
///
/// The public [`XmlSlice`] API intentionally continues to expose only an
/// `Arc<Vec<u8>>`.  Source-backed paragraphs use this private owner instead,
/// keeping the managed cache handle or source-publication token alive for as
/// long as any nested paragraph/run value is retained.
#[derive(Debug, Clone)]
pub(super) enum XmlOwner {
    Unmanaged(Arc<Vec<u8>>),
    Managed {
        data: Arc<PartData>,
        _admission: Arc<ManagedAdmission>,
    },
    Source {
        source: Arc<SourceXmlPart>,
        _admission: Arc<ManagedAdmission>,
    },
}

impl XmlOwner {
    #[inline]
    fn bytes(&self) -> &[u8] {
        match self {
            Self::Unmanaged(bytes) => bytes.as_slice(),
            Self::Managed { data, .. } => data.as_bytes(),
            Self::Source { source, .. } => source.bytes(),
        }
    }

    #[inline]
    fn unmanaged_arc(&self) -> Option<Arc<Vec<u8>>> {
        match self {
            Self::Unmanaged(bytes) => Some(Arc::clone(bytes)),
            Self::Managed { .. } | Self::Source { .. } => None,
        }
    }

    #[inline]
    fn parser_admission(&self, xml_len: usize) -> Result<Option<ManagedParserAdmission>> {
        match self {
            Self::Unmanaged(_) => Ok(None),
            Self::Managed { _admission, .. } | Self::Source { _admission, .. } => {
                _admission.parser_admission(xml_len).map(Some)
            },
        }
    }

    #[inline]
    fn namespace_scan_admission(
        &self,
        fragment_len: usize,
        owner_span_len: usize,
    ) -> Result<Option<ManagedNamespaceAdmission>> {
        match self {
            Self::Unmanaged(_) => Ok(None),
            Self::Managed { _admission, .. } | Self::Source { _admission, .. } => _admission
                .namespace_scan_admission(fragment_len, owner_span_len)
                .map(Some),
        }
    }

    #[inline]
    fn is_managed(&self) -> bool {
        matches!(self, Self::Managed { .. } | Self::Source { .. })
    }
}

fn range_bytes(bytes: &[u8], start: u32, length: u32) -> &[u8] {
    let Ok(start) = usize::try_from(start) else {
        return &[];
    };
    let Ok(length) = usize::try_from(length) else {
        return &[];
    };
    let Some(end) = start.checked_add(length) else {
        return &[];
    };
    bytes.get(start..end).unwrap_or_default()
}

/// A byte range whose owner remains attached to every cloned semantic value.
#[derive(Debug, Clone)]
pub(crate) struct XmlRef {
    owner: XmlOwner,
    start: u32,
    length: u32,
}

pub(crate) struct NamespaceResolverLease {
    resolver: NamespaceResolver,
    _admission: Option<ManagedNamespaceAdmission>,
}

impl Deref for NamespaceResolverLease {
    type Target = NamespaceResolver;

    fn deref(&self) -> &Self::Target {
        &self.resolver
    }
}

impl NamespaceResolverLease {
    /// Check the retained execution policy after a resolver-backed parse.
    /// The lease keeps its namespace scan reservation alive until the caller
    /// has finished consuming the resolver.
    pub(crate) fn check(&self) -> Result<()> {
        if let Some(admission) = self._admission.as_ref() {
            admission.check()?;
        }
        Ok(())
    }
}

impl XmlRef {
    #[inline]
    pub(super) const fn new(owner: XmlOwner, start: u32, length: u32) -> Self {
        Self {
            owner,
            start,
            length,
        }
    }

    #[inline]
    pub(crate) fn bytes(&self) -> &[u8] {
        range_bytes(self.owner.bytes(), self.start, self.length)
    }

    #[inline]
    pub(super) fn subrange(&self, relative_start: u32, length: u32) -> Option<Self> {
        let start = self.start.checked_add(relative_start)?;
        let relative_end = relative_start.checked_add(length)?;
        if relative_end > self.length {
            return None;
        }
        Some(Self::new(self.owner.clone(), start, length))
    }

    #[inline]
    pub(super) fn as_unmanaged_slice(&self) -> Option<XmlSlice> {
        self.owner
            .unmanaged_arc()
            .map(|source| XmlSlice::new(source, self.start, self.length))
    }

    #[inline]
    pub(crate) fn parser_admission(&self) -> Result<Option<ManagedParserAdmission>> {
        self.owner.parser_admission(self.bytes().len())
    }

    /// Build a resolver for this retained fragment, including declarations
    /// inherited from its owning full XML Part. The namespace scan admission
    /// remains attached to the returned lease until its resolver is consumed.
    pub(crate) fn namespace_resolver(&self) -> Result<NamespaceResolverLease> {
        let start = usize::try_from(self.start)
            .map_err(|_| Error::InvalidFormat("Word XML offset exceeds usize".into()))?;
        let full = self.owner.bytes();
        let fragment_len = usize::try_from(self.length)
            .map_err(|_| Error::InvalidFormat("Word XML range length exceeds usize".into()))?;
        let fragment_end = start
            .checked_add(fragment_len)
            .ok_or_else(|| Error::InvalidFormat("Word XML range overflows usize".into()))?;
        full.get(start..fragment_end).ok_or_else(|| {
            Error::InvalidFormat("Word XML retained range is outside its owner".into())
        })?;
        let admission = if start == 0 {
            None
        } else {
            // The selected start event is read as part of the inherited scan
            // so declarations attached to that event are included in the
            // resolver.  Admit through the complete retained range before
            // reading it; a large attribute list must be covered while the
            // reader owns its event and resolver buffers.
            self.owner
                .namespace_scan_admission(fragment_len, fragment_end)?
        };
        let mut resolver = NamespaceResolver::default();
        if start != 0 {
            // Parse through the selected element itself.  Reading the element
            // is necessary because `NsReader` applies namespace declarations
            // on its start event; stopping at the byte immediately before the
            // range would omit declarations placed on that element.  It also
            // lets the reader perform its pending pop for a preceding sibling
            // before resolving the selected element, so a sibling's local
            // declaration cannot leak into this fragment.
            let mut owner_reader = NsReader::from_reader(full);
            loop {
                if let Some(admission) = admission.as_ref() {
                    admission.check()?;
                }
                let event_start =
                    usize::try_from(owner_reader.buffer_position()).map_err(|_| {
                        Error::InvalidFormat("Word XML namespace offset exceeds usize".into())
                    })?;
                if event_start > start {
                    return Err(Error::InvalidFormat(
                        "Word XML namespace range starts inside an event".into(),
                    ));
                }
                let _event = owner_reader
                    .read_event()
                    .map_err(|error| Error::Xml(error.to_string()))?;
                let event_end = usize::try_from(owner_reader.buffer_position()).map_err(|_| {
                    Error::InvalidFormat("Word XML namespace offset exceeds usize".into())
                })?;
                if event_end > full.len() {
                    return Err(Error::InvalidFormat(
                        "Word XML namespace scan ended outside its owner".into(),
                    ));
                }
                if let Some(admission) = admission.as_ref() {
                    let event_bytes = event_end.saturating_sub(event_start).max(1);
                    let active_bindings =
                        owner_reader.resolver().bindings().count().saturating_add(2);
                    admission.consume_lookup_work(event_bytes, active_bindings)?;
                }
                if event_start == start {
                    break;
                }
            }
            for (prefix, namespace) in owner_reader.resolver().bindings() {
                resolver
                    .add(prefix, namespace)
                    .map_err(|error| Error::Xml(error.to_string()))?;
            }
        }
        if let Some(admission) = admission.as_ref() {
            admission.check()?;
        }
        Ok(NamespaceResolverLease {
            resolver,
            _admission: admission,
        })
    }

    #[inline]
    pub(crate) fn is_managed(&self) -> bool {
        self.owner.is_managed()
    }
}

/// Internal storage for paragraph XML data.
/// Supports both owned data (for standalone parsing) and shared slices (for arena-based parsing).
#[derive(Debug, Clone)]
pub(super) enum XmlData {
    /// Owned data for standalone paragraphs
    Owned(Box<[u8]>),
    /// Shared slice into an arena for zero-copy batch parsing
    Shared(XmlSlice),
    /// A range retained by a managed source-backed PartData handle.
    Managed {
        owner: Arc<PartData>,
        admission: Arc<ManagedAdmission>,
        start: u32,
        length: u32,
    },
    /// A range retained by a source-authorized XML publication token.
    Source {
        owner: Arc<SourceXmlPart>,
        admission: Arc<ManagedAdmission>,
        start: u32,
        length: u32,
    },
}

impl XmlData {
    #[inline]
    pub(super) fn as_bytes(&self) -> &[u8] {
        match self {
            XmlData::Owned(bytes) => bytes,
            XmlData::Shared(slice) => slice.as_bytes(),
            XmlData::Managed {
                owner,
                start,
                length,
                ..
            } => range_bytes(owner.as_bytes(), *start, *length),
            XmlData::Source {
                owner,
                start,
                length,
                ..
            } => range_bytes(owner.bytes(), *start, *length),
        }
    }

    #[inline]
    pub(crate) fn xml_ref(&self) -> Result<XmlRef> {
        match self {
            XmlData::Owned(bytes) => Ok(XmlRef::new(
                XmlOwner::Unmanaged(Arc::new(bytes.to_vec())),
                0,
                u32::try_from(bytes.len())
                    .map_err(|_| Error::InvalidFormat("Word paragraph XML exceeds u32".into()))?,
            )),
            XmlData::Shared(slice) => Ok(XmlRef::new(
                XmlOwner::Unmanaged(slice.arc()),
                slice.start(),
                u32::try_from(slice.len())
                    .map_err(|_| Error::InvalidFormat("Word paragraph XML exceeds u32".into()))?,
            )),
            XmlData::Managed {
                owner,
                admission,
                start,
                length,
            } => Ok(XmlRef::new(
                XmlOwner::Managed {
                    data: Arc::clone(owner),
                    _admission: Arc::clone(admission),
                },
                *start,
                *length,
            )),
            XmlData::Source {
                owner,
                admission,
                start,
                length,
            } => Ok(XmlRef::new(
                XmlOwner::Source {
                    source: Arc::clone(owner),
                    _admission: Arc::clone(admission),
                },
                *start,
                *length,
            )),
        }
    }

    #[inline]
    pub(super) fn parser_admission(&self) -> Result<Option<ManagedParserAdmission>> {
        match self {
            XmlData::Owned(_) | XmlData::Shared(_) => Ok(None),
            XmlData::Managed { admission, .. } | XmlData::Source { admission, .. } => {
                admission.parser_admission(self.as_bytes().len()).map(Some)
            },
        }
    }

    #[inline]
    pub(super) fn is_managed(&self) -> bool {
        matches!(self, XmlData::Managed { .. } | XmlData::Source { .. })
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
    source: XmlRef,
    word_hyperlink: bool,
}

impl OpaqueInline {
    pub(super) const fn from_xml_ref(source: XmlRef, word_hyperlink: bool) -> Self {
        Self {
            source,
            word_hyperlink,
        }
    }

    pub(crate) const fn is_word_hyperlink(&self) -> bool {
        self.word_hyperlink
    }

    pub(super) const fn xml_ref(&self) -> &XmlRef {
        &self.source
    }

    /// Borrow the retained paragraph child exactly as it appeared in the
    /// active paragraph XML.
    #[must_use]
    pub fn xml_bytes(&self) -> &[u8] {
        self.source.bytes()
    }
}

/// A direct run child whose semantics are not modeled by [`RunContent`].
#[derive(Debug, Clone)]
pub struct OpaqueRunContent {
    source: XmlRef,
}

impl OpaqueRunContent {
    pub(super) const fn from_xml_ref(source: XmlRef) -> Self {
        Self { source }
    }

    /// Borrow the retained run child exactly as it appeared in source XML.
    #[must_use]
    pub fn xml_bytes(&self) -> &[u8] {
        self.source.bytes()
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

    pub(crate) fn from_managed_range(
        owner: Arc<PartData>,
        admission: Arc<ManagedAdmission>,
        start: u32,
        len: u32,
    ) -> Self {
        Self {
            xml_data: XmlData::Managed {
                owner,
                admission,
                start,
                length: len,
            },
        }
    }

    pub(crate) fn from_source_range(
        owner: Arc<SourceXmlPart>,
        admission: Arc<ManagedAdmission>,
        start: u32,
        len: u32,
    ) -> Self {
        Self {
            xml_data: XmlData::Source {
                owner,
                admission,
                start,
                length: len,
            },
        }
    }

    /// Get the raw XML bytes.
    #[inline]
    pub(crate) fn xml_bytes(&self) -> &[u8] {
        self.xml_data.as_bytes()
    }

    pub(super) fn parser_admission(&self) -> Result<Option<ManagedParserAdmission>> {
        self.xml_data.parser_admission()
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
    Retained(XmlRef),
}

impl RunXmlData {
    #[inline]
    pub(super) fn as_bytes(&self) -> &[u8] {
        match self {
            RunXmlData::Owned(bytes) => bytes,
            RunXmlData::Shared(slice) => slice.as_bytes(),
            RunXmlData::Retained(source) => source.bytes(),
        }
    }

    #[inline]
    pub(super) fn xml_ref(&self) -> Result<XmlRef> {
        match self {
            RunXmlData::Owned(bytes) => Ok(XmlRef::new(
                XmlOwner::Unmanaged(Arc::new(bytes.clone())),
                0,
                u32::try_from(bytes.len())
                    .map_err(|_| Error::InvalidFormat("Word run XML exceeds u32".into()))?,
            )),
            RunXmlData::Shared(slice) => Ok(XmlRef::new(
                XmlOwner::Unmanaged(slice.arc()),
                slice.start(),
                u32::try_from(slice.len())
                    .map_err(|_| Error::InvalidFormat("Word run XML exceeds u32".into()))?,
            )),
            RunXmlData::Retained(source) => Ok(source.clone()),
        }
    }

    #[inline]
    pub(super) fn parser_admission(&self) -> Result<Option<ManagedParserAdmission>> {
        match self {
            RunXmlData::Owned(_) | RunXmlData::Shared(_) => Ok(None),
            RunXmlData::Retained(source) => source.parser_admission(),
        }
    }

    #[inline]
    pub(super) fn is_managed(&self) -> bool {
        match self {
            RunXmlData::Owned(_) | RunXmlData::Shared(_) => false,
            RunXmlData::Retained(source) => source.is_managed(),
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

    pub(super) fn from_xml_ref(source: XmlRef) -> Self {
        Self {
            xml_data: RunXmlData::Retained(source),
        }
    }

    /// Get the raw XML bytes.
    #[inline]
    pub(crate) fn xml_bytes(&self) -> &[u8] {
        self.xml_data.as_bytes()
    }

    pub(crate) fn xml_ref(&self) -> Result<XmlRef> {
        self.xml_data.xml_ref()
    }

    pub(crate) fn parser_admission(&self) -> Result<Option<ManagedParserAdmission>> {
        self.xml_data.parser_admission()
    }

    pub(crate) fn is_managed(&self) -> bool {
        self.xml_data.is_managed()
    }

    pub(crate) fn replace_xml(&mut self, xml_bytes: Vec<u8>) {
        self.xml_data = RunXmlData::Owned(xml_bytes);
    }
}
