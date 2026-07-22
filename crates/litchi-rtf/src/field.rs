//! Safe, structured RTF field-code support.

use std::borrow::Cow;

const MAX_INSTRUCTION_LEN: usize = 65_536;
const MAX_TOKENS: usize = 256;
pub(crate) const MAX_GENERIC_FIELDS: usize = 65_536;

/// Field type in RTF documents.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FieldType {
    Hyperlink,
    Reference,
    PageReference,
    NoteReference,
    Page,
    Date,
    Toc,
    TocEntry,
    TableOfAuthorities,
    TableOfAuthoritiesEntry,
    Bookmark,
    Equation,
    MacroButton,
    Dde,
    DdeAuto,
    Link,
    IncludeText,
    IncludePicture,
    Index,
    Unknown,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FieldOwner {
    Detached,
    Body,
    Header,
    Footer,
    Footnote,
    Endnote,
    TableCell(u8),
    FieldResult,
    FormField,
    Other,
}

/// A zero-width explicit `\page` control at a UTF-8 story boundary.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PageBreak {
    pub position: usize,
}

impl PageBreak {
    pub const fn new(position: usize) -> Self {
        Self { position }
    }
}

/// A zero-width explicit `\sect` control at a UTF-8 main-story boundary.
///
/// `next_section` identifies the typed section definition that starts after
/// this boundary. `None` means the following section inherits its properties
/// and therefore has no separately retained section definition.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SectionBreak {
    pub position: usize,
    pub next_section: Option<usize>,
}

impl SectionBreak {
    pub const fn new(position: usize, next_section: Option<usize>) -> Self {
        Self {
            position,
            next_section,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BodyStoryEvent {
    PageBreak(PageBreak),
    SectionBreak(SectionBreak),
    Drawing(crate::StoryDrawing),
    Field(usize),
    BookmarkStart(usize),
    BookmarkEnd(usize),
    AnnotationStart(usize),
    AnnotationEnd(usize),
    Note(usize),
    Object(usize),
    PictureCompatibility(usize),
    FormFieldStart(usize),
    FormFieldEnd(usize),
    RevisionStart(usize),
    RevisionEnd(usize),
    RevisionDeletion(usize),
    GeneratedListMarker(usize),
    LegacyTextBox(usize),
    LegacyDrawing(usize),
    NavigationEntry(usize),
}

/// A generic field reference embedded in a non-body text story.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct StoryField {
    pub field_index: usize,
    pub position: usize,
}

/// Exact source order of drawings and generic fields in a text story.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum StoryEvent {
    PageBreak(PageBreak),
    Drawing(crate::StoryDrawing),
    Field(StoryField),
}

/// One token from a field instruction.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FieldCodeToken<'a> {
    pub value: Cow<'a, str>,
    pub quoted: bool,
}

/// A preserved field-code switch.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FieldSwitch<'a> {
    pub name: Cow<'a, str>,
    pub value: Option<Cow<'a, str>>,
}

/// A parsed HYPERLINK field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct HyperlinkCode<'a> {
    pub external_target: Option<Cow<'a, str>>,
    pub bookmark: Option<Cow<'a, str>>,
    pub screen_tip: Option<Cow<'a, str>>,
    pub target_frame: Option<Cow<'a, str>>,
    pub coordinates: Option<Cow<'a, str>>,
    pub new_window: bool,
    pub unknown_switches: Vec<FieldSwitch<'a>>,
}

/// A parsed REF, PAGEREF, or NOTEREF field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ReferenceCode<'a> {
    pub bookmark: Cow<'a, str>,
    pub hyperlink: bool,
    pub position: bool,
    pub footnote_mark: bool,
    pub unknown_switches: Vec<FieldSwitch<'a>>,
}

/// Why a recognized field code is non-actionable.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum FieldCodeError {
    InstructionTooLong,
    TooManyTokens,
    UnterminatedQuote,
    MissingKeyword,
    MissingOperand(&'static str),
    DuplicateOperand(&'static str),
    UnexpectedOperand(String),
}

/// Typed field semantics. Malformed input is represented, never activated.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ParsedFieldCode<'a> {
    Hyperlink(HyperlinkCode<'a>),
    Reference(ReferenceCode<'a>),
    PageReference(ReferenceCode<'a>),
    NoteReference(ReferenceCode<'a>),
    Other {
        keyword: Cow<'a, str>,
        arguments: Vec<FieldCodeToken<'a>>,
    },
    Malformed(FieldCodeError),
}

/// Presence-only state carried by a generic RTF field.
///
/// Each `false` value means the corresponding control word was omitted.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct FieldStatus {
    pub dirty: bool,
    pub edited: bool,
    pub locked: bool,
    pub private: bool,
}

/// Parsed RTF field.
#[derive(Debug, Clone)]
pub struct Field<'a> {
    pub field_type: FieldType,
    pub instruction: Cow<'a, str>,
    pub result: Cow<'a, str>,
    pub status: FieldStatus,
    pub shapes: Vec<crate::Shape<'a>>,
    pub shape_groups: Vec<crate::ShapeGroup<'a>>,
    pub drawing_order: Vec<crate::StoryDrawing>,
    /// Exact source order of drawings and nested generic fields in the result story.
    pub result_events: Vec<StoryEvent>,
    pub owner: FieldOwner,
    pub position: usize,
    pub range_end: usize,
}

/// Inert metadata for a legacy RTF `EQ` field.
///
/// The expression is retained exactly as field-instruction text after the
/// `EQ` keyword. It is never parsed as an equation, evaluated, rendered, or
/// sent to an external application.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct EquationField<'a> {
    instruction: &'a str,
    expression: &'a str,
    cached_result: Option<&'a str>,
    status: FieldStatus,
    owner: FieldOwner,
    position: usize,
}

/// Inert metadata for a legacy RTF `MACROBUTTON` field.
///
/// The macro name and button text are exposed solely as stored field metadata.
/// This crate never resolves, loads, invokes, or otherwise executes the named
/// macro.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MacroButtonField<'a> {
    instruction: &'a str,
    macro_name: Cow<'a, str>,
    display_text: Option<Cow<'a, str>>,
    cached_result: Option<&'a str>,
    status: FieldStatus,
    owner: FieldOwner,
    position: usize,
}

/// The stored kind of a legacy RTF DDE field.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DdeFieldKind {
    /// A `DDE` field, which may carry an automatic-update switch.
    Dde,
    /// A `DDEAUTO` field, which declares automatic updates.
    DdeAuto,
}

/// One stored DDE result representation switch.
///
/// This value describes the requested representation only. It never causes the
/// source to be contacted, converted, or embedded.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DdeRepresentation {
    /// The `\\b` switch requests a bitmap representation.
    Bitmap,
    /// The `\\h` switch requests HTML-formatted text.
    Html,
    /// The `\\p` switch requests a picture representation.
    Picture,
    /// The `\\r` switch requests rich-text format.
    RichText,
    /// The `\\t` switch requests text-only format.
    Text,
    /// The `\\u` switch requests Unicode text.
    UnicodeText,
}

/// Inert metadata for legacy RTF `DDE` and `DDEAUTO` fields.
///
/// Application, source, and item values are exposed solely as stored field
/// metadata. This crate never launches an application, opens a source,
/// initiates a DDE conversation, requests data, refreshes the field, or
/// executes code.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DdeField<'a> {
    instruction: &'a str,
    kind: DdeFieldKind,
    application: Cow<'a, str>,
    source: Cow<'a, str>,
    item: Option<Cow<'a, str>>,
    automatic_updates: bool,
    representation: Option<DdeRepresentation>,
    omit_graphic_data: bool,
    unknown_switches: Vec<FieldSwitch<'a>>,
    cached_result: Option<&'a str>,
    status: FieldStatus,
    owner: FieldOwner,
    position: usize,
}

/// One stored result or storage switch for a `LINK` field.
///
/// These values describe the requested linked-object representation or whether
/// its graphic data is stored. They never cause the source to be contacted,
/// converted, embedded, or displayed.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LinkResultOption {
    /// The `\\b` switch requests a bitmap representation.
    Bitmap,
    /// The `\\d` switch omits graphic data from the document.
    OmitGraphicData,
    /// The `\\h` switch requests HTML-formatted text.
    Html,
    /// The `\\p` switch requests a picture representation.
    Picture,
    /// The `\\r` switch requests rich-text format.
    RichText,
    /// The `\\t` switch requests text-only format.
    Text,
    /// The `\\u` switch requests Unicode text.
    UnicodeText,
}

/// One integral `LINK` `\\f` formatting mode.
///
/// Values marked unsupported by ECMA-376, and values outside its defined set,
/// are preserved as `Unsupported` metadata. This crate does not format linked
/// content or evaluate the mode.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LinkFormatting {
    /// `0`: preserve formatting from the source file.
    Source,
    /// `2`: match formatting in the destination document.
    Destination,
    /// `4`: preserve source formatting for a SpreadsheetML workbook source.
    SpreadsheetSource,
    /// `5`: match destination formatting for a SpreadsheetML workbook source.
    SpreadsheetDestination,
    /// An ECMA-376-unsupported or otherwise unrecognized integral mode.
    Unsupported(i64),
}

/// Inert metadata for a legacy RTF `LINK` field.
///
/// The application type, source, and item are exposed solely as stored field
/// metadata. This crate never activates an OLE server, launches an application,
/// opens a source, requests data, refreshes the field, converts content, or
/// executes code.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct LinkField<'a> {
    instruction: &'a str,
    application_type: Cow<'a, str>,
    source: Cow<'a, str>,
    item: Option<Cow<'a, str>>,
    automatic_updates: bool,
    result_options: Vec<LinkResultOption>,
    formatting_modes: Vec<LinkFormatting>,
    unknown_switches: Vec<FieldSwitch<'a>>,
    cached_result: Option<&'a str>,
    status: FieldStatus,
    owner: FieldOwner,
    position: usize,
}

/// The kind of external content referenced by an RTF include field.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum IncludeFieldKind {
    /// An `INCLUDETEXT` field that refers to document text and graphics.
    Text,
    /// An `INCLUDEPICTURE` field that refers to a graphic.
    Picture,
}

/// One recognized stored option of an external-include field.
///
/// These values are configuration metadata only. This crate never opens,
/// resolves, imports, transforms, or evaluates the referenced source.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ExternalIncludeOption<'a> {
    /// A document or graphics converter name from the c switch.
    Converter(Cow<'a, str>),
    /// A source encoding from the INCLUDETEXT e switch.
    Encoding(Cow<'a, str>),
    /// A source MIME type from the INCLUDETEXT m switch.
    MimeType(Cow<'a, str>),
    /// An XML namespace mapping from the INCLUDETEXT n switch.
    NamespaceMapping(Cow<'a, str>),
    /// An XSLT location from the INCLUDETEXT t switch.
    Xslt(Cow<'a, str>),
    /// An XPath expression from the INCLUDETEXT x switch.
    XPath(Cow<'a, str>),
}

/// Inert metadata for legacy RTF external-content fields.
///
/// This represents `INCLUDETEXT` and `INCLUDEPICTURE` field instructions.
/// Sources, converter names, and XML options are retained as stored metadata
/// only. This crate never opens, resolves, fetches, transforms, converts,
/// updates, or writes back to the source.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExternalIncludeField<'a> {
    instruction: &'a str,
    kind: IncludeFieldKind,
    source: Cow<'a, str>,
    bookmark: Option<Cow<'a, str>>,
    suppress_nested_field_updates: bool,
    omit_picture_data: bool,
    options: Vec<ExternalIncludeOption<'a>>,
    unknown_switches: Vec<FieldSwitch<'a>>,
    cached_result: Option<&'a str>,
    status: FieldStatus,
    owner: FieldOwner,
    position: usize,
}

/// One recognized stored option of a TOC field.
///
/// These values describe how a producer configured a table of contents. They
/// are inert metadata only: this crate never scans entries, paginates,
/// generates a table, follows its hyperlinks, or refreshes the field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum TableOfContentsOption<'a> {
    /// The \\a caption label whose item labels and numbers are omitted.
    CaptionWithoutLabel(Cow<'a, str>),
    /// The \\b bookmark that bounds included entries.
    Bookmark(Cow<'a, str>),
    /// The \\c SEQ identifier for a table of captions.
    CaptionSequence(Cow<'a, str>),
    /// The \\d separator between sequence and page numbers.
    SequencePageSeparator(Cow<'a, str>),
    /// The \\f TC-field identifier that selects entries.
    TableEntryIdentifier(Cow<'a, str>),
    /// The \\h switch requests hyperlinks for entries.
    Hyperlinks,
    /// The \\l range of TC-field entry levels to include.
    TableEntryLevels(Cow<'a, str>),
    /// The \\n switch omits page numbers, optionally for an entry-level range.
    OmitPageNumbers(Option<Cow<'a, str>>),
    /// The \\o built-in heading-style range, or all used heading levels.
    HeadingStyleRange(Option<Cow<'a, str>>),
    /// The \\p separator between an entry and its page number.
    EntryPageNumberSeparator(Cow<'a, str>),
    /// The \\s SEQ identifier whose number prefixes page numbers.
    SequenceIdentifier(Cow<'a, str>),
    /// The \\t custom style-name/TOC-level mappings.
    StyleMappings(Cow<'a, str>),
    /// The \\u switch uses applied paragraph outline levels.
    OutlineLevels,
    /// The \\w switch preserves tab characters within entries.
    PreserveTabs,
    /// The \\x switch preserves newline characters within entries.
    PreserveNewlines,
    /// The \\z switch hides page numbers and leaders in web-page view.
    HidePageNumbersInWebView,
}

/// Inert metadata for a legacy RTF TOC field.
///
/// This model retains the stored configuration and cached result only. It
/// never searches for TC fields, reads bookmarks, resolves hyperlinks,
/// calculates page numbers, or generates a table of contents.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TableOfContentsField<'a> {
    instruction: &'a str,
    options: Vec<TableOfContentsOption<'a>>,
    unknown_switches: Vec<FieldSwitch<'a>>,
    cached_result: Option<&'a str>,
    status: FieldStatus,
    owner: FieldOwner,
    position: usize,
}

/// One recognized stored option of a TC field.
///
/// These values identify how an entry participates in a table of contents.
/// They are inert metadata only: this crate never changes hidden text,
/// calculates page numbers, or generates a table.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum TableOfContentsEntryOption<'a> {
    /// The \\f contents-list identifier.
    ListIdentifier(Cow<'a, str>),
    /// The \\l entry level.
    Level(Cow<'a, str>),
    /// The \\n switch omits the entry page number.
    OmitPageNumber,
}

/// Inert metadata for a legacy RTF TC field.
///
/// This model retains a stored table-of-contents entry marker and its cached
/// result only. It never updates the document, changes hidden text, calculates
/// a page number, or generates a table of contents.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TableOfContentsEntryField<'a> {
    instruction: &'a str,
    entry: Cow<'a, str>,
    options: Vec<TableOfContentsEntryOption<'a>>,
    unknown_switches: Vec<FieldSwitch<'a>>,
    cached_result: Option<&'a str>,
    status: FieldStatus,
    owner: FieldOwner,
    position: usize,
}

/// One recognized stored option of a TA field.
///
/// These values describe a legal-authority entry marker. They are inert
/// metadata only: this crate never locates citations, changes hidden text,
/// calculates page numbers, or generates a table of authorities.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum TableOfAuthoritiesEntryOption<'a> {
    /// The \\b switch toggles bold formatting for the entry page number.
    BoldPageNumber,
    /// The \\c integral authority category.
    Category(Cow<'a, str>),
    /// The \\i switch toggles italic formatting for the entry page number.
    ItalicPageNumber,
    /// The \\l long citation text.
    LongCitation(Cow<'a, str>),
    /// The \\r bookmark that marks the cited page range.
    PageRangeBookmark(Cow<'a, str>),
    /// The \\s short citation text.
    ShortCitation(Cow<'a, str>),
}

/// Inert metadata for a legacy RTF TA field.
///
/// This model retains a stored table-of-authorities marker and cached result
/// only. It never finds cited text, changes hidden text, follows bookmarks,
/// paginates the document, or generates a table of authorities.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TableOfAuthoritiesEntryField<'a> {
    instruction: &'a str,
    options: Vec<TableOfAuthoritiesEntryOption<'a>>,
    unknown_switches: Vec<FieldSwitch<'a>>,
    cached_result: Option<&'a str>,
    status: FieldStatus,
    owner: FieldOwner,
    position: usize,
}

/// One recognized stored option of a TOA field.
///
/// These values describe a table-of-authorities configuration. They are inert
/// metadata only: this crate never finds citations, follows bookmarks,
/// calculates page numbers, or generates a table.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum TableOfAuthoritiesOption<'a> {
    /// The \\b bookmark that bounds included entries.
    Bookmark(Cow<'a, str>),
    /// The \\c integral authority category to include.
    Category(Cow<'a, str>),
    /// The \\d separator between sequence and page numbers.
    SequencePageSeparator(Cow<'a, str>),
    /// The \\e separator between an entry and its page number.
    EntryPageNumberSeparator(Cow<'a, str>),
    /// The \\f switch removes source-text formatting from entries.
    RemoveEntryFormatting,
    /// The \\g separator between page numbers in a page range.
    PageRangeSeparator(Cow<'a, str>),
    /// The \\h switch includes category headings.
    CategoryHeadings,
    /// The \\l separator between multiple page references.
    PageReferenceSeparator(Cow<'a, str>),
    /// The \\p switch replaces five or more page references with passim.
    UsePassim,
    /// The \\s SEQ identifier whose number prefixes page numbers.
    SequenceIdentifier(Cow<'a, str>),
}

/// Inert metadata for a legacy RTF TOA field.
///
/// This model retains a stored table-of-authorities configuration and cached
/// result only. It never finds citations, follows bookmarks, calculates page
/// numbers, paginates the document, or generates a table of authorities.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TableOfAuthoritiesField<'a> {
    instruction: &'a str,
    options: Vec<TableOfAuthoritiesOption<'a>>,
    unknown_switches: Vec<FieldSwitch<'a>>,
    cached_result: Option<&'a str>,
    status: FieldStatus,
    owner: FieldOwner,
    position: usize,
}

struct ExternalIncludeParts<'a> {
    kind: IncludeFieldKind,
    source: Cow<'a, str>,
    bookmark: Option<Cow<'a, str>>,
    suppress_nested_field_updates: bool,
    omit_picture_data: bool,
    options: Vec<ExternalIncludeOption<'a>>,
    unknown_switches: Vec<FieldSwitch<'a>>,
}

struct TableOfContentsParts<'a> {
    options: Vec<TableOfContentsOption<'a>>,
    unknown_switches: Vec<FieldSwitch<'a>>,
}

struct TableOfContentsEntryParts<'a> {
    entry: Cow<'a, str>,
    options: Vec<TableOfContentsEntryOption<'a>>,
    unknown_switches: Vec<FieldSwitch<'a>>,
}

struct TableOfAuthoritiesEntryParts<'a> {
    options: Vec<TableOfAuthoritiesEntryOption<'a>>,
    unknown_switches: Vec<FieldSwitch<'a>>,
}

struct TableOfAuthoritiesParts<'a> {
    options: Vec<TableOfAuthoritiesOption<'a>>,
    unknown_switches: Vec<FieldSwitch<'a>>,
}

struct DdeFieldParts<'a> {
    kind: DdeFieldKind,
    application: Cow<'a, str>,
    source: Cow<'a, str>,
    item: Option<Cow<'a, str>>,
    automatic_updates: bool,
    representation: Option<DdeRepresentation>,
    omit_graphic_data: bool,
    unknown_switches: Vec<FieldSwitch<'a>>,
}

struct LinkFieldParts<'a> {
    application_type: Cow<'a, str>,
    source: Cow<'a, str>,
    item: Option<Cow<'a, str>>,
    automatic_updates: bool,
    result_options: Vec<LinkResultOption>,
    formatting_modes: Vec<LinkFormatting>,
    unknown_switches: Vec<FieldSwitch<'a>>,
}

impl<'a> EquationField<'a> {
    /// Return the complete stored `EQ` field instruction.
    pub fn instruction(&self) -> &'a str {
        self.instruction
    }

    /// Return the opaque equation expression after the `EQ` keyword.
    pub fn expression(&self) -> &'a str {
        self.expression
    }

    /// Return the stored field result when a producer supplied one.
    ///
    /// RTF 1.9.1 examples normally use an empty result for `EQ` fields. This
    /// value is metadata only and is never recalculated.
    pub fn cached_result(&self) -> Option<&'a str> {
        self.cached_result
    }

    /// Return the stored field state flags.
    pub const fn status(&self) -> FieldStatus {
        self.status
    }

    /// Return the story that owns this field.
    pub const fn owner(&self) -> FieldOwner {
        self.owner
    }

    /// Return this field's zero-width position in its owning story.
    pub const fn position(&self) -> usize {
        self.position
    }

    /// Whether a producer marked the stored field result stale.
    pub const fn is_dirty(&self) -> bool {
        self.status.dirty
    }

    /// Whether a producer locked the field against refresh.
    pub const fn is_locked(&self) -> bool {
        self.status.locked
    }
}

impl<'a> MacroButtonField<'a> {
    /// Return the complete stored `MACROBUTTON` field instruction.
    pub fn instruction(&self) -> &'a str {
        self.instruction
    }

    /// Return the stored macro name without resolving or invoking it.
    pub fn macro_name(&self) -> &str {
        &self.macro_name
    }

    /// Return the optional text stored after the macro name.
    ///
    /// This is the field's button/display text, not a generated value.
    pub fn display_text(&self) -> Option<&str> {
        self.display_text.as_deref()
    }

    /// Return the stored field result when a producer supplied one.
    pub fn cached_result(&self) -> Option<&'a str> {
        self.cached_result
    }

    /// Return the stored field state flags.
    pub const fn status(&self) -> FieldStatus {
        self.status
    }

    /// Return the story that owns this field.
    pub const fn owner(&self) -> FieldOwner {
        self.owner
    }

    /// Return this field's zero-width position in its owning story.
    pub const fn position(&self) -> usize {
        self.position
    }

    /// Whether a producer marked the stored field result stale.
    pub const fn is_dirty(&self) -> bool {
        self.status.dirty
    }

    /// Whether a producer locked the field against refresh.
    pub const fn is_locked(&self) -> bool {
        self.status.locked
    }
}

impl<'a> DdeField<'a> {
    /// Return the complete stored DDE field instruction.
    pub fn instruction(&self) -> &'a str {
        self.instruction
    }

    /// Return whether this is a `DDE` or `DDEAUTO` field.
    pub const fn kind(&self) -> DdeFieldKind {
        self.kind
    }

    /// Return the stored DDE application name without launching it.
    pub fn application(&self) -> &str {
        &self.application
    }

    /// Return the stored source identifier without opening or resolving it.
    pub fn source(&self) -> &str {
        &self.source
    }

    /// Return the optional stored source item, such as a cell range or bookmark.
    pub fn item(&self) -> Option<&str> {
        self.item.as_deref()
    }

    /// Whether the stored instruction requests automatic DDE updates.
    ///
    /// This is metadata only. The crate never performs an update.
    pub const fn requests_automatic_updates(&self) -> bool {
        self.automatic_updates
    }

    /// Return the requested stored result representation, if present.
    ///
    /// This is metadata only and never triggers source access or conversion.
    pub const fn representation(&self) -> Option<DdeRepresentation> {
        self.representation
    }

    /// Whether the stored `\\d` switch omits graphic data from the document.
    ///
    /// This is stored metadata only. The crate never reads the source to obtain
    /// omitted data.
    pub const fn omits_graphic_data(&self) -> bool {
        self.omit_graphic_data
    }

    /// Return unrecognized stored field switches in source order.
    pub fn unknown_switches(&self) -> &[FieldSwitch<'a>] {
        &self.unknown_switches
    }

    /// Return the stored field result when a producer supplied one.
    pub fn cached_result(&self) -> Option<&'a str> {
        self.cached_result
    }

    /// Return the stored field state flags.
    pub const fn status(&self) -> FieldStatus {
        self.status
    }

    /// Return the story that owns this field.
    pub const fn owner(&self) -> FieldOwner {
        self.owner
    }

    /// Return this field's zero-width position in its owning story.
    pub const fn position(&self) -> usize {
        self.position
    }

    /// Whether a producer marked the stored field result stale.
    pub const fn is_dirty(&self) -> bool {
        self.status.dirty
    }

    /// Whether a producer locked the field against refresh.
    pub const fn is_locked(&self) -> bool {
        self.status.locked
    }
}

impl<'a> LinkField<'a> {
    /// Return the complete stored `LINK` field instruction.
    pub fn instruction(&self) -> &'a str {
        self.instruction
    }

    /// Return the stored linked-object application type.
    ///
    /// Word commonly stores an OLE Programmatic Identifier here. This method
    /// returns it as metadata only and never looks up or activates the class.
    pub fn application_type(&self) -> &str {
        &self.application_type
    }

    /// Return the stored linked source identifier without opening or resolving it.
    pub fn source(&self) -> &str {
        &self.source
    }

    /// Return the optional stored source item, such as a cell range or bookmark.
    pub fn item(&self) -> Option<&str> {
        self.item.as_deref()
    }

    /// Whether the stored instruction requests automatic updates.
    ///
    /// This is metadata only. The crate never performs an update.
    pub const fn requests_automatic_updates(&self) -> bool {
        self.automatic_updates
    }

    /// Return recognized result and storage switches in stored source order.
    ///
    /// These values never trigger source access, conversion, or display. When
    /// several are present, [`Self::effective_result_option`] reflects Word's
    /// documented last-switch behavior.
    pub fn result_options(&self) -> &[LinkResultOption] {
        &self.result_options
    }

    /// Return the effective result or storage option under Word's documented
    /// last-switch behavior, if one was stored.
    ///
    /// This reports metadata only and never contacts the linked source.
    pub fn effective_result_option(&self) -> Option<LinkResultOption> {
        self.result_options.last().copied()
    }

    /// Return integral `\\f` formatting modes in stored source order.
    ///
    /// This crate never updates or formats linked content.
    pub fn formatting_modes(&self) -> &[LinkFormatting] {
        &self.formatting_modes
    }

    /// Return unrecognized stored field switches in source order.
    pub fn unknown_switches(&self) -> &[FieldSwitch<'a>] {
        &self.unknown_switches
    }

    /// Return the stored field result when a producer supplied one.
    pub fn cached_result(&self) -> Option<&'a str> {
        self.cached_result
    }

    /// Return the stored field state flags.
    pub const fn status(&self) -> FieldStatus {
        self.status
    }

    /// Return the story that owns this field.
    pub const fn owner(&self) -> FieldOwner {
        self.owner
    }

    /// Return this field's zero-width position in its owning story.
    pub const fn position(&self) -> usize {
        self.position
    }

    /// Whether a producer marked the stored field result stale.
    pub const fn is_dirty(&self) -> bool {
        self.status.dirty
    }

    /// Whether a producer locked the field against refresh.
    pub const fn is_locked(&self) -> bool {
        self.status.locked
    }
}

impl<'a> ExternalIncludeField<'a> {
    /// Return the complete stored include-field instruction.
    pub fn instruction(&self) -> &'a str {
        self.instruction
    }

    /// Return whether this stores an `INCLUDETEXT` or `INCLUDEPICTURE` field.
    pub const fn kind(&self) -> IncludeFieldKind {
        self.kind
    }

    /// Return the stored source path or URL without resolving it.
    pub fn source(&self) -> &str {
        &self.source
    }

    /// Return the optional `INCLUDETEXT` bookmark selector.
    ///
    /// `INCLUDEPICTURE` fields do not define a bookmark operand, so they
    /// always return `None` here.
    pub fn bookmark(&self) -> Option<&str> {
        self.bookmark.as_deref()
    }

    /// Return the optional stored `\\c` converter name.
    ///
    /// The converter is never looked up or invoked.
    pub fn converter(&self) -> Option<&str> {
        self.options.iter().find_map(|option| match option {
            ExternalIncludeOption::Converter(value) => Some(value.as_ref()),
            _ => None,
        })
    }

    /// Return recognized converter and XML options in stored source order.
    ///
    /// All options are inert metadata. This method never resolves a converter,
    /// opens a source, runs XSLT, or evaluates XPath.
    pub fn options(&self) -> &[ExternalIncludeOption<'a>] {
        &self.options
    }

    /// Whether an `INCLUDETEXT` `\\!` switch suppresses nested field updates.
    ///
    /// This is stored metadata only; this crate never updates fields.
    pub const fn suppresses_nested_field_updates(&self) -> bool {
        self.suppress_nested_field_updates
    }

    /// Whether an `INCLUDEPICTURE` `\\d` switch omits picture data.
    ///
    /// This is stored metadata only; this crate never retrieves a picture.
    pub const fn omits_picture_data(&self) -> bool {
        self.omit_picture_data
    }

    /// Return unrecognized stored field switches in source order.
    pub fn unknown_switches(&self) -> &[FieldSwitch<'a>] {
        &self.unknown_switches
    }

    /// Return the stored field result when a producer supplied one.
    pub fn cached_result(&self) -> Option<&'a str> {
        self.cached_result
    }

    /// Return the stored field state flags.
    pub const fn status(&self) -> FieldStatus {
        self.status
    }

    /// Return the story that owns this field.
    pub const fn owner(&self) -> FieldOwner {
        self.owner
    }

    /// Return this field's zero-width position in its owning story.
    pub const fn position(&self) -> usize {
        self.position
    }

    /// Whether a producer marked the stored field result stale.
    pub const fn is_dirty(&self) -> bool {
        self.status.dirty
    }

    /// Whether a producer locked the field against refresh.
    pub const fn is_locked(&self) -> bool {
        self.status.locked
    }
}

impl<'a> TableOfContentsField<'a> {
    /// Return the complete stored TOC field instruction.
    pub fn instruction(&self) -> &'a str {
        self.instruction
    }

    /// Return recognized TOC options in stored source order.
    ///
    /// These options are configuration metadata only. This method never scans
    /// entries, regenerates a table, follows links, or calculates page numbers.
    pub fn options(&self) -> &[TableOfContentsOption<'a>] {
        &self.options
    }

    /// Return unrecognized stored field switches in source order.
    pub fn unknown_switches(&self) -> &[FieldSwitch<'a>] {
        &self.unknown_switches
    }

    /// Return the stored field result when a producer supplied one.
    ///
    /// This is cached text only and is never regenerated.
    pub fn cached_result(&self) -> Option<&'a str> {
        self.cached_result
    }

    /// Return the stored field state flags.
    pub const fn status(&self) -> FieldStatus {
        self.status
    }

    /// Return the story that owns this field.
    pub const fn owner(&self) -> FieldOwner {
        self.owner
    }

    /// Return this field's zero-width position in its owning story.
    pub const fn position(&self) -> usize {
        self.position
    }

    /// Whether a producer marked the stored field result stale.
    pub const fn is_dirty(&self) -> bool {
        self.status.dirty
    }

    /// Whether a producer locked the field against refresh.
    pub const fn is_locked(&self) -> bool {
        self.status.locked
    }
}

impl<'a> TableOfContentsEntryField<'a> {
    /// Return the complete stored TC field instruction.
    pub fn instruction(&self) -> &'a str {
        self.instruction
    }

    /// Return the stored entry text without generating a table of contents.
    pub fn entry(&self) -> &str {
        &self.entry
    }

    /// Return recognized TC options in stored source order.
    ///
    /// These options are configuration metadata only. This method never
    /// calculates page numbers, changes hidden text, or updates a TOC.
    pub fn options(&self) -> &[TableOfContentsEntryOption<'a>] {
        &self.options
    }

    /// Return unrecognized stored field switches in source order.
    pub fn unknown_switches(&self) -> &[FieldSwitch<'a>] {
        &self.unknown_switches
    }

    /// Return the stored field result when a producer supplied one.
    pub fn cached_result(&self) -> Option<&'a str> {
        self.cached_result
    }

    /// Return the stored field state flags.
    pub const fn status(&self) -> FieldStatus {
        self.status
    }

    /// Return the story that owns this field.
    pub const fn owner(&self) -> FieldOwner {
        self.owner
    }

    /// Return this field's zero-width position in its owning story.
    pub const fn position(&self) -> usize {
        self.position
    }

    /// Whether a producer marked the stored field result stale.
    pub const fn is_dirty(&self) -> bool {
        self.status.dirty
    }

    /// Whether a producer locked the field against refresh.
    pub const fn is_locked(&self) -> bool {
        self.status.locked
    }
}

impl<'a> TableOfAuthoritiesEntryField<'a> {
    /// Return the complete stored TA field instruction.
    pub fn instruction(&self) -> &'a str {
        self.instruction
    }

    /// Return recognized TA options in stored source order.
    ///
    /// These options are configuration metadata only. This method never finds
    /// citations, follows bookmarks, calculates pages, or generates a table.
    pub fn options(&self) -> &[TableOfAuthoritiesEntryOption<'a>] {
        &self.options
    }

    /// Return unrecognized stored field switches in source order.
    pub fn unknown_switches(&self) -> &[FieldSwitch<'a>] {
        &self.unknown_switches
    }

    /// Return the stored field result when a producer supplied one.
    pub fn cached_result(&self) -> Option<&'a str> {
        self.cached_result
    }

    /// Return the stored field state flags.
    pub const fn status(&self) -> FieldStatus {
        self.status
    }

    /// Return the story that owns this field.
    pub const fn owner(&self) -> FieldOwner {
        self.owner
    }

    /// Return this field's zero-width position in its owning story.
    pub const fn position(&self) -> usize {
        self.position
    }

    /// Whether a producer marked the stored field result stale.
    pub const fn is_dirty(&self) -> bool {
        self.status.dirty
    }

    /// Whether a producer locked the field against refresh.
    pub const fn is_locked(&self) -> bool {
        self.status.locked
    }
}

impl<'a> TableOfAuthoritiesField<'a> {
    /// Return the complete stored TOA field instruction.
    pub fn instruction(&self) -> &'a str {
        self.instruction
    }

    /// Return recognized TOA options in stored source order.
    ///
    /// These options are configuration metadata only. This method never finds
    /// citations, follows bookmarks, calculates pages, or generates a table.
    pub fn options(&self) -> &[TableOfAuthoritiesOption<'a>] {
        &self.options
    }

    /// Return unrecognized stored field switches in source order.
    pub fn unknown_switches(&self) -> &[FieldSwitch<'a>] {
        &self.unknown_switches
    }

    /// Return the stored field result when a producer supplied one.
    ///
    /// This is cached text only and is never regenerated.
    pub fn cached_result(&self) -> Option<&'a str> {
        self.cached_result
    }

    /// Return the stored field state flags.
    pub const fn status(&self) -> FieldStatus {
        self.status
    }

    /// Return the story that owns this field.
    pub const fn owner(&self) -> FieldOwner {
        self.owner
    }

    /// Return this field's zero-width position in its owning story.
    pub const fn position(&self) -> usize {
        self.position
    }

    /// Whether a producer marked the stored field result stale.
    pub const fn is_dirty(&self) -> bool {
        self.status.dirty
    }

    /// Whether a producer locked the field against refresh.
    pub const fn is_locked(&self) -> bool {
        self.status.locked
    }
}

impl<'a> Field<'a> {
    #[inline]
    pub fn new(field_type: FieldType, instruction: Cow<'a, str>, result: Cow<'a, str>) -> Self {
        Self {
            field_type,
            instruction,
            result,
            status: FieldStatus::default(),
            shapes: Vec::new(),
            shape_groups: Vec::new(),
            drawing_order: Vec::new(),
            result_events: Vec::new(),
            owner: FieldOwner::Detached,
            position: 0,
            range_end: 0,
        }
    }

    /// Parse the instruction keyword with an exact, case-insensitive boundary.
    pub fn parse_instruction(instruction: &'a str) -> Self {
        let parsed = parse_field_code(instruction);
        let field_type = match parsed {
            ParsedFieldCode::Hyperlink(_) => FieldType::Hyperlink,
            ParsedFieldCode::Reference(_) => FieldType::Reference,
            ParsedFieldCode::PageReference(_) => FieldType::PageReference,
            ParsedFieldCode::NoteReference(_) => FieldType::NoteReference,
            ParsedFieldCode::Other { ref keyword, .. } if keyword.eq_ignore_ascii_case("PAGE") => {
                FieldType::Page
            },
            ParsedFieldCode::Other { ref keyword, .. }
                if keyword.eq_ignore_ascii_case("DATE") || keyword.eq_ignore_ascii_case("TIME") =>
            {
                FieldType::Date
            },
            ParsedFieldCode::Other { ref keyword, .. } if keyword.eq_ignore_ascii_case("TOC") => {
                FieldType::Toc
            },
            ParsedFieldCode::Other { ref keyword, .. } if keyword.eq_ignore_ascii_case("TC") => {
                FieldType::TocEntry
            },
            ParsedFieldCode::Other { ref keyword, .. } if keyword.eq_ignore_ascii_case("TOA") => {
                FieldType::TableOfAuthorities
            },
            ParsedFieldCode::Other { ref keyword, .. } if keyword.eq_ignore_ascii_case("TA") => {
                FieldType::TableOfAuthoritiesEntry
            },
            ParsedFieldCode::Other { ref keyword, .. }
                if keyword.eq_ignore_ascii_case("BOOKMARK") =>
            {
                FieldType::Bookmark
            },
            ParsedFieldCode::Other { ref keyword, .. } if keyword.eq_ignore_ascii_case("EQ") => {
                FieldType::Equation
            },
            ParsedFieldCode::Other { ref keyword, .. }
                if keyword.eq_ignore_ascii_case("MACROBUTTON") =>
            {
                FieldType::MacroButton
            },
            ParsedFieldCode::Other { ref keyword, .. } if keyword.eq_ignore_ascii_case("DDE") => {
                FieldType::Dde
            },
            ParsedFieldCode::Other { ref keyword, .. }
                if keyword.eq_ignore_ascii_case("DDEAUTO") =>
            {
                FieldType::DdeAuto
            },
            ParsedFieldCode::Other { ref keyword, .. } if keyword.eq_ignore_ascii_case("LINK") => {
                FieldType::Link
            },
            ParsedFieldCode::Other { ref keyword, .. }
                if keyword.eq_ignore_ascii_case("INCLUDETEXT") =>
            {
                FieldType::IncludeText
            },
            ParsedFieldCode::Other { ref keyword, .. }
                if keyword.eq_ignore_ascii_case("INCLUDEPICTURE") =>
            {
                FieldType::IncludePicture
            },
            ParsedFieldCode::Other { ref keyword, .. }
                if keyword.eq_ignore_ascii_case("INDEX") || keyword.eq_ignore_ascii_case("XE") =>
            {
                FieldType::Index
            },
            _ => FieldType::Unknown,
        };
        Self {
            field_type,
            instruction: Cow::Borrowed(instruction),
            result: Cow::Borrowed(""),
            status: FieldStatus::default(),
            shapes: Vec::new(),
            shape_groups: Vec::new(),
            drawing_order: Vec::new(),
            result_events: Vec::new(),
            owner: FieldOwner::Detached,
            position: 0,
            range_end: 0,
        }
    }

    /// Construct an inert `EQ` field from caller-provided equation syntax.
    ///
    /// The expression is serialized as field text with RTF escaping. The
    /// library never parses, calculates, formats, or renders that syntax.
    pub fn new_equation(expression: impl Into<String>) -> crate::RtfResult<Field<'static>> {
        let expression = expression.into();
        let instruction = if expression.is_empty() {
            "EQ".to_string()
        } else {
            format!("EQ {expression}")
        };
        if instruction.len() > MAX_INSTRUCTION_LEN {
            return Err(crate::RtfError::MalformedDocument(
                "RTF EQ field instruction exceeds the safety limit".to_string(),
            ));
        }
        Ok(Field::new(
            FieldType::Equation,
            Cow::Owned(instruction),
            Cow::Borrowed(""),
        ))
    }

    /// Return typed inert metadata when this is an `EQ` field.
    pub fn equation(&self) -> Option<EquationField<'_>> {
        if self.field_type != FieldType::Equation {
            return None;
        }
        let expression = equation_expression(self.instruction.as_ref())?;
        Some(EquationField {
            instruction: self.instruction.as_ref(),
            expression,
            cached_result: (!self.result.is_empty()).then_some(self.result.as_ref()),
            status: self.status,
            owner: self.owner,
            position: self.position,
        })
    }

    /// Return inert metadata when this is a well-formed `MACROBUTTON` field.
    ///
    /// The metadata is never treated as executable. Malformed macro-button
    /// instructions remain generic fields and return `None` here.
    pub fn macro_button(&self) -> Option<MacroButtonField<'_>> {
        if self.field_type != FieldType::MacroButton {
            return None;
        }
        let (macro_name, display_text) = macro_button_parts(self.instruction.as_ref())?;
        Some(MacroButtonField {
            instruction: self.instruction.as_ref(),
            macro_name,
            display_text,
            cached_result: (!self.result.is_empty()).then_some(self.result.as_ref()),
            status: self.status,
            owner: self.owner,
            position: self.position,
        })
    }

    /// Return inert metadata when this is a well-formed `DDE` or `DDEAUTO` field.
    ///
    /// The application, source, and item are never opened, contacted,
    /// refreshed, converted, evaluated, or executed. Malformed DDE
    /// instructions remain generic fields and return `None` here.
    pub fn dde_link(&self) -> Option<DdeField<'_>> {
        if !matches!(self.field_type, FieldType::Dde | FieldType::DdeAuto) {
            return None;
        }
        let parts = dde_field_parts(self.instruction.as_ref())?;
        Some(DdeField {
            instruction: self.instruction.as_ref(),
            kind: parts.kind,
            application: parts.application,
            source: parts.source,
            item: parts.item,
            automatic_updates: parts.automatic_updates,
            representation: parts.representation,
            omit_graphic_data: parts.omit_graphic_data,
            unknown_switches: parts.unknown_switches,
            cached_result: (!self.result.is_empty()).then_some(self.result.as_ref()),
            status: self.status,
            owner: self.owner,
            position: self.position,
        })
    }

    /// Return inert metadata when this is a well-formed `LINK` field.
    ///
    /// The application type, source, and item are never activated, opened,
    /// contacted, refreshed, converted, evaluated, or executed. Malformed
    /// link instructions remain generic fields and return `None` here.
    pub fn link_field(&self) -> Option<LinkField<'_>> {
        if self.field_type != FieldType::Link {
            return None;
        }
        let parts = link_field_parts(self.instruction.as_ref())?;
        Some(LinkField {
            instruction: self.instruction.as_ref(),
            application_type: parts.application_type,
            source: parts.source,
            item: parts.item,
            automatic_updates: parts.automatic_updates,
            result_options: parts.result_options,
            formatting_modes: parts.formatting_modes,
            unknown_switches: parts.unknown_switches,
            cached_result: (!self.result.is_empty()).then_some(self.result.as_ref()),
            status: self.status,
            owner: self.owner,
            position: self.position,
        })
    }

    /// Return inert metadata when this is a well-formed external include field.
    ///
    /// Sources are never resolved, opened, fetched, converted, updated, or
    /// written back. Malformed include instructions remain generic fields and
    /// return `None` here.
    pub fn external_include(&self) -> Option<ExternalIncludeField<'_>> {
        if !matches!(
            self.field_type,
            FieldType::IncludeText | FieldType::IncludePicture
        ) {
            return None;
        }
        let parts = external_include_parts(self.instruction.as_ref())?;
        Some(ExternalIncludeField {
            instruction: self.instruction.as_ref(),
            kind: parts.kind,
            source: parts.source,
            bookmark: parts.bookmark,
            suppress_nested_field_updates: parts.suppress_nested_field_updates,
            omit_picture_data: parts.omit_picture_data,
            options: parts.options,
            unknown_switches: parts.unknown_switches,
            cached_result: (!self.result.is_empty()).then_some(self.result.as_ref()),
            status: self.status,
            owner: self.owner,
            position: self.position,
        })
    }

    /// Return inert metadata when this is a well-formed TOC field.
    ///
    /// The stored configuration and cached result are never used to scan
    /// entries, read bookmarks, resolve links, calculate page numbers, or
    /// regenerate a table. Malformed TOC instructions remain generic fields
    /// and return None here.
    pub fn table_of_contents(&self) -> Option<TableOfContentsField<'_>> {
        if self.field_type != FieldType::Toc {
            return None;
        }
        let parts = table_of_contents_parts(self.instruction.as_ref())?;
        Some(TableOfContentsField {
            instruction: self.instruction.as_ref(),
            options: parts.options,
            unknown_switches: parts.unknown_switches,
            cached_result: (!self.result.is_empty()).then_some(self.result.as_ref()),
            status: self.status,
            owner: self.owner,
            position: self.position,
        })
    }

    /// Return inert metadata when this is a well-formed TC entry field.
    ///
    /// The stored entry and cached result are never used to change hidden text,
    /// calculate a page number, or generate a table of contents. Malformed TC
    /// instructions remain generic fields and return None here.
    pub fn table_of_contents_entry(&self) -> Option<TableOfContentsEntryField<'_>> {
        if self.field_type != FieldType::TocEntry {
            return None;
        }
        let parts = table_of_contents_entry_parts(self.instruction.as_ref())?;
        Some(TableOfContentsEntryField {
            instruction: self.instruction.as_ref(),
            entry: parts.entry,
            options: parts.options,
            unknown_switches: parts.unknown_switches,
            cached_result: (!self.result.is_empty()).then_some(self.result.as_ref()),
            status: self.status,
            owner: self.owner,
            position: self.position,
        })
    }

    /// Return inert metadata when this is a well-formed TA entry field.
    ///
    /// The stored citation options and cached result are never used to find
    /// text, follow bookmarks, change hidden text, calculate pages, or generate
    /// a table of authorities. Malformed TA instructions remain generic fields
    /// and return None here.
    pub fn table_of_authorities_entry(&self) -> Option<TableOfAuthoritiesEntryField<'_>> {
        if self.field_type != FieldType::TableOfAuthoritiesEntry {
            return None;
        }
        let parts = table_of_authorities_entry_parts(self.instruction.as_ref())?;
        Some(TableOfAuthoritiesEntryField {
            instruction: self.instruction.as_ref(),
            options: parts.options,
            unknown_switches: parts.unknown_switches,
            cached_result: (!self.result.is_empty()).then_some(self.result.as_ref()),
            status: self.status,
            owner: self.owner,
            position: self.position,
        })
    }

    /// Return inert metadata when this is a well-formed TOA field.
    ///
    /// The stored configuration and cached result are never used to find
    /// citations, follow bookmarks, calculate page numbers, paginate the
    /// document, or generate a table of authorities. Malformed TOA
    /// instructions remain generic fields and return None here.
    pub fn table_of_authorities(&self) -> Option<TableOfAuthoritiesField<'_>> {
        if self.field_type != FieldType::TableOfAuthorities {
            return None;
        }
        let parts = table_of_authorities_parts(self.instruction.as_ref())?;
        Some(TableOfAuthoritiesField {
            instruction: self.instruction.as_ref(),
            options: parts.options,
            unknown_switches: parts.unknown_switches,
            cached_result: (!self.result.is_empty()).then_some(self.result.as_ref()),
            status: self.status,
            owner: self.owner,
            position: self.position,
        })
    }

    pub const fn status(&self) -> FieldStatus {
        self.status
    }

    pub fn set_status(&mut self, status: FieldStatus) {
        self.status = status;
    }

    pub fn validate(&self) -> crate::RtfResult<()> {
        if self.position > self.range_end {
            return Err(crate::RtfError::MalformedDocument(
                "RTF generic field range moves backwards".to_string(),
            ));
        }
        if self.position != self.range_end {
            return Err(crate::RtfError::MalformedDocument(
                "RTF generic fields must be zero-width enclosing-story events".to_string(),
            ));
        }
        validate_story_events(
            self.result.as_ref(),
            &self.shapes,
            &self.shape_groups,
            &self.drawing_order,
            &self.result_events,
            "field result",
        )
    }

    pub fn push_shape(&mut self, shape: crate::Shape<'a>) -> crate::RtfResult<()> {
        let mut shapes = self.shapes.clone();
        let mut order = self.drawing_order.clone();
        order.push(crate::StoryDrawing::Shape(shapes.len()));
        shapes.push(shape);
        crate::shape::validate_story_drawings(
            self.result.as_ref(),
            &shapes,
            &self.shape_groups,
            &order,
            "field result",
        )?;
        self.shapes = shapes;
        self.drawing_order = order;
        self.result_events
            .push(StoryEvent::Drawing(crate::StoryDrawing::Shape(
                self.shapes.len() - 1,
            )));
        Ok(())
    }

    pub fn push_shape_group(&mut self, group: crate::ShapeGroup<'a>) -> crate::RtfResult<()> {
        let mut groups = self.shape_groups.clone();
        let mut order = self.drawing_order.clone();
        order.push(crate::StoryDrawing::ShapeGroup(groups.len()));
        groups.push(group);
        crate::shape::validate_story_drawings(
            self.result.as_ref(),
            &self.shapes,
            &groups,
            &order,
            "field result",
        )?;
        self.shape_groups = groups;
        self.drawing_order = order;
        self.result_events
            .push(StoryEvent::Drawing(crate::StoryDrawing::ShapeGroup(
                self.shape_groups.len() - 1,
            )));
        Ok(())
    }

    pub fn clear_drawings(&mut self) {
        self.shapes.clear();
        self.shape_groups.clear();
        self.drawing_order.clear();
        self.result_events
            .retain(|event| !matches!(event, StoryEvent::Drawing(_)));
    }

    pub fn page_breaks(&self) -> impl Iterator<Item = PageBreak> + '_ {
        self.result_events.iter().filter_map(|event| match event {
            StoryEvent::PageBreak(value) => Some(*value),
            _ => None,
        })
    }

    pub fn push_page_break(&mut self, position: usize) -> crate::RtfResult<()> {
        push_story_page_break(&mut self.result_events, self.result.as_ref(), position, "field result")
    }

    pub fn clear_page_breaks(&mut self) {
        self.result_events.retain(|event| !matches!(event, StoryEvent::PageBreak(_)));
    }

    /// Parse this field's instruction into bounded, typed semantics.
    pub fn parsed_code(&self) -> ParsedFieldCode<'_> {
        parse_field_code(self.instruction.as_ref())
    }

    /// Compatibility URL helper. Internal-only links return `#bookmark`.
    pub fn extract_url(&self) -> Option<String> {
        let ParsedFieldCode::Hyperlink(code) = self.parsed_code() else {
            return None;
        };
        code.external_target
            .map(Cow::into_owned)
            .or_else(|| code.bookmark.map(|bookmark| format!("#{bookmark}")))
    }

    /// Compatibility bookmark helper for reference and hyperlink fields.
    pub fn extract_bookmark(&self) -> Option<String> {
        match self.parsed_code() {
            ParsedFieldCode::Hyperlink(code) => code.bookmark.map(Cow::into_owned),
            ParsedFieldCode::Reference(code)
            | ParsedFieldCode::PageReference(code)
            | ParsedFieldCode::NoteReference(code) => Some(code.bookmark.into_owned()),
            _ => None,
        }
    }

    #[inline]
    pub fn display_text(&self) -> &str {
        if !self.result.is_empty() {
            &self.result
        } else {
            &self.instruction
        }
    }
}

fn equation_expression(instruction: &str) -> Option<&str> {
    let instruction = instruction.trim_start_matches(|value: char| value.is_ascii_whitespace());
    let keyword_len = instruction
        .find(|value: char| value.is_ascii_whitespace())
        .unwrap_or(instruction.len());
    instruction[..keyword_len]
        .eq_ignore_ascii_case("EQ")
        .then(|| {
            instruction[keyword_len..].trim_start_matches(|value: char| value.is_ascii_whitespace())
        })
}

fn macro_button_parts(instruction: &str) -> Option<(Cow<'_, str>, Option<Cow<'_, str>>)> {
    let mut tokens = tokenize(instruction).ok()?;
    let keyword = tokens.first()?;
    if !keyword.value.eq_ignore_ascii_case("MACROBUTTON") {
        return None;
    }
    tokens.remove(0);
    let macro_name = tokens.first()?.value.clone();
    if macro_name.is_empty() {
        return None;
    }
    let display_text = match tokens.len() {
        1 => None,
        2 => Some(tokens[1].value.clone()),
        _ => Some(Cow::Owned(
            tokens[1..]
                .iter()
                .map(|token| token.value.as_ref())
                .collect::<Vec<_>>()
                .join(" "),
        )),
    };
    Some((macro_name, display_text))
}

fn dde_field_parts(instruction: &str) -> Option<DdeFieldParts<'_>> {
    let mut tokens = tokenize(instruction).ok()?;
    let keyword = tokens.first()?;
    let kind = if keyword.value.eq_ignore_ascii_case("DDE") {
        DdeFieldKind::Dde
    } else if keyword.value.eq_ignore_ascii_case("DDEAUTO") {
        DdeFieldKind::DdeAuto
    } else {
        return None;
    };
    tokens.remove(0);

    let application = tokens.first()?.value.clone();
    if application.is_empty() || switch_name(tokens.first()?).is_some() {
        return None;
    }
    tokens.remove(0);

    let source = tokens.first()?.value.clone();
    if source.is_empty() || switch_name(tokens.first()?).is_some() {
        return None;
    }
    tokens.remove(0);

    let item = if tokens
        .first()
        .is_some_and(|token| switch_name(token).is_none())
    {
        Some(tokens.remove(0).value)
    } else {
        None
    };

    let mut automatic_updates = kind == DdeFieldKind::DdeAuto;
    let mut has_automatic_update_switch = false;
    let mut representation = None;
    let mut omit_graphic_data = false;
    let mut unknown_switches = Vec::new();
    let mut index = 0;
    while index < tokens.len() {
        let name = switch_name(&tokens[index])?;
        let normalized_name = name.to_ascii_lowercase();
        match normalized_name.as_str() {
            "a" if kind == DdeFieldKind::Dde => {
                if has_automatic_update_switch
                    || tokens
                        .get(index + 1)
                        .is_some_and(|token| switch_name(token).is_none())
                {
                    return None;
                }
                automatic_updates = true;
                has_automatic_update_switch = true;
                index += 1;
            },
            "a" => return None,
            "d" => {
                if representation.is_some()
                    || omit_graphic_data
                    || tokens
                        .get(index + 1)
                        .is_some_and(|token| switch_name(token).is_none())
                {
                    return None;
                }
                omit_graphic_data = true;
                index += 1;
            },
            "b" | "h" | "p" | "r" | "t" | "u" => {
                if representation.is_some()
                    || omit_graphic_data
                    || tokens
                        .get(index + 1)
                        .is_some_and(|token| switch_name(token).is_none())
                {
                    return None;
                }
                representation = Some(match normalized_name.as_str() {
                    "b" => DdeRepresentation::Bitmap,
                    "h" => DdeRepresentation::Html,
                    "p" => DdeRepresentation::Picture,
                    "r" => DdeRepresentation::RichText,
                    "t" => DdeRepresentation::Text,
                    "u" => DdeRepresentation::UnicodeText,
                    _ => unreachable!("DDE representation switch was matched above"),
                });
                index += 1;
            },
            _ => {
                let value = tokens
                    .get(index + 1)
                    .filter(|token| switch_name(token).is_none());
                unknown_switches.push(FieldSwitch {
                    name: Cow::Owned(name.to_string()),
                    value: value.map(|token| token.value.clone()),
                });
                index += 1 + usize::from(value.is_some());
            },
        }
    }

    Some(DdeFieldParts {
        kind,
        application,
        source,
        item,
        automatic_updates,
        representation,
        omit_graphic_data,
        unknown_switches,
    })
}

fn link_field_parts(instruction: &str) -> Option<LinkFieldParts<'_>> {
    let mut tokens = tokenize(instruction).ok()?;
    let keyword = tokens.first()?;
    if !keyword.value.eq_ignore_ascii_case("LINK") {
        return None;
    }
    tokens.remove(0);

    let application_type = tokens.first()?.value.clone();
    if application_type.is_empty() || switch_name(tokens.first()?).is_some() {
        return None;
    }
    tokens.remove(0);

    let source = tokens.first()?.value.clone();
    if source.is_empty() || switch_name(tokens.first()?).is_some() {
        return None;
    }
    tokens.remove(0);

    let item = if tokens
        .first()
        .is_some_and(|token| switch_name(token).is_none())
    {
        Some(tokens.remove(0).value)
    } else {
        None
    };

    let mut automatic_updates = false;
    let mut result_options = Vec::new();
    let mut formatting_modes = Vec::new();
    let mut unknown_switches = Vec::new();
    let mut index = 0;
    while index < tokens.len() {
        let name = switch_name(&tokens[index])?;
        let normalized_name = name.to_ascii_lowercase();
        match normalized_name.as_str() {
            "a" => {
                if tokens
                    .get(index + 1)
                    .is_some_and(|token| switch_name(token).is_none())
                {
                    return None;
                }
                automatic_updates = true;
                index += 1;
            },
            "f" => {
                let value = switch_value(&tokens, index, name).ok()?;
                let value = value.parse::<i64>().ok()?;
                formatting_modes.push(match value {
                    0 => LinkFormatting::Source,
                    2 => LinkFormatting::Destination,
                    4 => LinkFormatting::SpreadsheetSource,
                    5 => LinkFormatting::SpreadsheetDestination,
                    other => LinkFormatting::Unsupported(other),
                });
                index += 2;
            },
            "b" | "d" | "h" | "p" | "r" | "t" | "u" => {
                if tokens
                    .get(index + 1)
                    .is_some_and(|token| switch_name(token).is_none())
                {
                    return None;
                }
                result_options.push(match normalized_name.as_str() {
                    "b" => LinkResultOption::Bitmap,
                    "d" => LinkResultOption::OmitGraphicData,
                    "h" => LinkResultOption::Html,
                    "p" => LinkResultOption::Picture,
                    "r" => LinkResultOption::RichText,
                    "t" => LinkResultOption::Text,
                    "u" => LinkResultOption::UnicodeText,
                    _ => unreachable!("LINK result option switch was matched above"),
                });
                index += 1;
            },
            _ => {
                let value = tokens
                    .get(index + 1)
                    .filter(|token| switch_name(token).is_none());
                unknown_switches.push(FieldSwitch {
                    name: Cow::Owned(name.to_string()),
                    value: value.map(|token| token.value.clone()),
                });
                index += 1 + usize::from(value.is_some());
            },
        }
    }

    Some(LinkFieldParts {
        application_type,
        source,
        item,
        automatic_updates,
        result_options,
        formatting_modes,
        unknown_switches,
    })
}

fn external_include_parts(instruction: &str) -> Option<ExternalIncludeParts<'_>> {
    let mut tokens = tokenize(instruction).ok()?;
    let keyword = tokens.first()?;
    let kind = if keyword.value.eq_ignore_ascii_case("INCLUDETEXT") {
        IncludeFieldKind::Text
    } else if keyword.value.eq_ignore_ascii_case("INCLUDEPICTURE") {
        IncludeFieldKind::Picture
    } else {
        return None;
    };
    tokens.remove(0);

    let source = tokens.first()?.value.clone();
    if source.is_empty() || switch_name(tokens.first()?).is_some() {
        return None;
    }
    tokens.remove(0);

    let bookmark = if kind == IncludeFieldKind::Text
        && tokens
            .first()
            .is_some_and(|token| switch_name(token).is_none())
    {
        Some(tokens.remove(0).value)
    } else {
        None
    };
    if kind == IncludeFieldKind::Picture
        && tokens
            .first()
            .is_some_and(|token| switch_name(token).is_none())
    {
        return None;
    }

    let mut options = Vec::new();
    let mut suppress_nested_field_updates = false;
    let mut omit_picture_data = false;
    let mut unknown_switches = Vec::new();
    let mut index = 0;
    while index < tokens.len() {
        let name = switch_name(&tokens[index])?;
        if name.eq_ignore_ascii_case("c") {
            options.push(ExternalIncludeOption::Converter(
                switch_value(&tokens, index, name).ok()?,
            ));
            index += 2;
        } else if kind == IncludeFieldKind::Text && name.eq_ignore_ascii_case("e") {
            options.push(ExternalIncludeOption::Encoding(
                switch_value(&tokens, index, name).ok()?,
            ));
            index += 2;
        } else if kind == IncludeFieldKind::Text && name.eq_ignore_ascii_case("m") {
            options.push(ExternalIncludeOption::MimeType(
                switch_value(&tokens, index, name).ok()?,
            ));
            index += 2;
        } else if kind == IncludeFieldKind::Text && name.eq_ignore_ascii_case("n") {
            options.push(ExternalIncludeOption::NamespaceMapping(
                switch_value(&tokens, index, name).ok()?,
            ));
            index += 2;
        } else if kind == IncludeFieldKind::Text && name.eq_ignore_ascii_case("t") {
            options.push(ExternalIncludeOption::Xslt(
                switch_value(&tokens, index, name).ok()?,
            ));
            index += 2;
        } else if kind == IncludeFieldKind::Text && name.eq_ignore_ascii_case("x") {
            options.push(ExternalIncludeOption::XPath(
                switch_value(&tokens, index, name).ok()?,
            ));
            index += 2;
        } else if kind == IncludeFieldKind::Text && name == "!" {
            if suppress_nested_field_updates
                || tokens
                    .get(index + 1)
                    .is_some_and(|token| switch_name(token).is_none())
            {
                return None;
            }
            suppress_nested_field_updates = true;
            index += 1;
        } else if kind == IncludeFieldKind::Picture && name.eq_ignore_ascii_case("d") {
            if omit_picture_data
                || tokens
                    .get(index + 1)
                    .is_some_and(|token| switch_name(token).is_none())
            {
                return None;
            }
            omit_picture_data = true;
            index += 1;
        } else {
            let value = tokens
                .get(index + 1)
                .filter(|token| switch_name(token).is_none());
            unknown_switches.push(FieldSwitch {
                name: Cow::Owned(name.to_string()),
                value: value.map(|token| token.value.clone()),
            });
            index += 1 + usize::from(value.is_some());
        }
    }

    Some(ExternalIncludeParts {
        kind,
        source,
        bookmark,
        suppress_nested_field_updates,
        omit_picture_data,
        options,
        unknown_switches,
    })
}

fn table_of_contents_parts(instruction: &str) -> Option<TableOfContentsParts<'_>> {
    let mut tokens = tokenize(instruction).ok()?;
    let keyword = tokens.first()?;
    if !keyword.value.eq_ignore_ascii_case("TOC") {
        return None;
    }
    tokens.remove(0);

    let mut options = Vec::new();
    let mut unknown_switches = Vec::new();
    let mut index = 0;
    while index < tokens.len() {
        let name = switch_name(&tokens[index])?;
        let normalized_name = name.to_ascii_lowercase();
        match normalized_name.as_str() {
            "a" => {
                options.push(TableOfContentsOption::CaptionWithoutLabel(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "b" => {
                options.push(TableOfContentsOption::Bookmark(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "c" => {
                options.push(TableOfContentsOption::CaptionSequence(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "d" => {
                options.push(TableOfContentsOption::SequencePageSeparator(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "f" => {
                options.push(TableOfContentsOption::TableEntryIdentifier(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "h" => {
                if tokens
                    .get(index + 1)
                    .is_some_and(|token| switch_name(token).is_none())
                {
                    return None;
                }
                options.push(TableOfContentsOption::Hyperlinks);
                index += 1;
            },
            "l" => {
                options.push(TableOfContentsOption::TableEntryLevels(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "n" => {
                let range = tokens
                    .get(index + 1)
                    .filter(|token| switch_name(token).is_none())
                    .map(|token| token.value.clone());
                options.push(TableOfContentsOption::OmitPageNumbers(range));
                index += 1 + usize::from(
                    tokens
                        .get(index + 1)
                        .is_some_and(|token| switch_name(token).is_none()),
                );
            },
            "o" => {
                let range = tokens
                    .get(index + 1)
                    .filter(|token| switch_name(token).is_none())
                    .map(|token| token.value.clone());
                options.push(TableOfContentsOption::HeadingStyleRange(range));
                index += 1 + usize::from(
                    tokens
                        .get(index + 1)
                        .is_some_and(|token| switch_name(token).is_none()),
                );
            },
            "p" => {
                options.push(TableOfContentsOption::EntryPageNumberSeparator(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "s" => {
                options.push(TableOfContentsOption::SequenceIdentifier(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "t" => {
                options.push(TableOfContentsOption::StyleMappings(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "u" => {
                if tokens
                    .get(index + 1)
                    .is_some_and(|token| switch_name(token).is_none())
                {
                    return None;
                }
                options.push(TableOfContentsOption::OutlineLevels);
                index += 1;
            },
            "w" => {
                if tokens
                    .get(index + 1)
                    .is_some_and(|token| switch_name(token).is_none())
                {
                    return None;
                }
                options.push(TableOfContentsOption::PreserveTabs);
                index += 1;
            },
            "x" => {
                if tokens
                    .get(index + 1)
                    .is_some_and(|token| switch_name(token).is_none())
                {
                    return None;
                }
                options.push(TableOfContentsOption::PreserveNewlines);
                index += 1;
            },
            "z" => {
                if tokens
                    .get(index + 1)
                    .is_some_and(|token| switch_name(token).is_none())
                {
                    return None;
                }
                options.push(TableOfContentsOption::HidePageNumbersInWebView);
                index += 1;
            },
            _ => {
                let value = tokens
                    .get(index + 1)
                    .filter(|token| switch_name(token).is_none());
                unknown_switches.push(FieldSwitch {
                    name: Cow::Owned(name.to_string()),
                    value: value.map(|token| token.value.clone()),
                });
                index += 1 + usize::from(value.is_some());
            },
        }
    }

    Some(TableOfContentsParts {
        options,
        unknown_switches,
    })
}

fn table_of_contents_entry_parts(instruction: &str) -> Option<TableOfContentsEntryParts<'_>> {
    let mut tokens = tokenize(instruction).ok()?;
    let keyword = tokens.first()?;
    if !keyword.value.eq_ignore_ascii_case("TC") {
        return None;
    }
    tokens.remove(0);

    let entry = tokens.first()?.value.clone();
    if entry.is_empty() || switch_name(tokens.first()?).is_some() {
        return None;
    }
    tokens.remove(0);

    let mut options = Vec::new();
    let mut unknown_switches = Vec::new();
    let mut index = 0;
    while index < tokens.len() {
        let name = switch_name(&tokens[index])?;
        let normalized_name = name.to_ascii_lowercase();
        match normalized_name.as_str() {
            "f" => {
                options.push(TableOfContentsEntryOption::ListIdentifier(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "l" => {
                options.push(TableOfContentsEntryOption::Level(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "n" => {
                if tokens
                    .get(index + 1)
                    .is_some_and(|token| switch_name(token).is_none())
                {
                    return None;
                }
                options.push(TableOfContentsEntryOption::OmitPageNumber);
                index += 1;
            },
            _ => {
                let value = tokens
                    .get(index + 1)
                    .filter(|token| switch_name(token).is_none());
                unknown_switches.push(FieldSwitch {
                    name: Cow::Owned(name.to_string()),
                    value: value.map(|token| token.value.clone()),
                });
                index += 1 + usize::from(value.is_some());
            },
        }
    }

    Some(TableOfContentsEntryParts {
        entry,
        options,
        unknown_switches,
    })
}

fn table_of_authorities_entry_parts(instruction: &str) -> Option<TableOfAuthoritiesEntryParts<'_>> {
    let mut tokens = tokenize(instruction).ok()?;
    let keyword = tokens.first()?;
    if !keyword.value.eq_ignore_ascii_case("TA") {
        return None;
    }
    tokens.remove(0);

    let mut options = Vec::new();
    let mut unknown_switches = Vec::new();
    let mut index = 0;
    while index < tokens.len() {
        let name = switch_name(&tokens[index])?;
        let normalized_name = name.to_ascii_lowercase();
        match normalized_name.as_str() {
            "b" => {
                if tokens
                    .get(index + 1)
                    .is_some_and(|token| switch_name(token).is_none())
                {
                    return None;
                }
                options.push(TableOfAuthoritiesEntryOption::BoldPageNumber);
                index += 1;
            },
            "c" => {
                options.push(TableOfAuthoritiesEntryOption::Category(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "i" => {
                if tokens
                    .get(index + 1)
                    .is_some_and(|token| switch_name(token).is_none())
                {
                    return None;
                }
                options.push(TableOfAuthoritiesEntryOption::ItalicPageNumber);
                index += 1;
            },
            "l" => {
                options.push(TableOfAuthoritiesEntryOption::LongCitation(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "r" => {
                options.push(TableOfAuthoritiesEntryOption::PageRangeBookmark(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "s" => {
                options.push(TableOfAuthoritiesEntryOption::ShortCitation(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            _ => {
                let value = tokens
                    .get(index + 1)
                    .filter(|token| switch_name(token).is_none());
                unknown_switches.push(FieldSwitch {
                    name: Cow::Owned(name.to_string()),
                    value: value.map(|token| token.value.clone()),
                });
                index += 1 + usize::from(value.is_some());
            },
        }
    }

    Some(TableOfAuthoritiesEntryParts {
        options,
        unknown_switches,
    })
}

fn table_of_authorities_parts(instruction: &str) -> Option<TableOfAuthoritiesParts<'_>> {
    let mut tokens = tokenize(instruction).ok()?;
    let keyword = tokens.first()?;
    if !keyword.value.eq_ignore_ascii_case("TOA") {
        return None;
    }
    tokens.remove(0);

    let mut options = Vec::new();
    let mut unknown_switches = Vec::new();
    let mut index = 0;
    while index < tokens.len() {
        let name = switch_name(&tokens[index])?;
        let normalized_name = name.to_ascii_lowercase();
        match normalized_name.as_str() {
            "b" => {
                options.push(TableOfAuthoritiesOption::Bookmark(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "c" => {
                options.push(TableOfAuthoritiesOption::Category(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "d" => {
                options.push(TableOfAuthoritiesOption::SequencePageSeparator(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "e" => {
                options.push(TableOfAuthoritiesOption::EntryPageNumberSeparator(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "f" => {
                if tokens
                    .get(index + 1)
                    .is_some_and(|token| switch_name(token).is_none())
                {
                    return None;
                }
                options.push(TableOfAuthoritiesOption::RemoveEntryFormatting);
                index += 1;
            },
            "g" => {
                options.push(TableOfAuthoritiesOption::PageRangeSeparator(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "h" => {
                if tokens
                    .get(index + 1)
                    .is_some_and(|token| switch_name(token).is_none())
                {
                    return None;
                }
                options.push(TableOfAuthoritiesOption::CategoryHeadings);
                index += 1;
            },
            "l" => {
                options.push(TableOfAuthoritiesOption::PageReferenceSeparator(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "p" => {
                if tokens
                    .get(index + 1)
                    .is_some_and(|token| switch_name(token).is_none())
                {
                    return None;
                }
                options.push(TableOfAuthoritiesOption::UsePassim);
                index += 1;
            },
            "s" => {
                options.push(TableOfAuthoritiesOption::SequenceIdentifier(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            _ => {
                let value = tokens
                    .get(index + 1)
                    .filter(|token| switch_name(token).is_none());
                unknown_switches.push(FieldSwitch {
                    name: Cow::Owned(name.to_string()),
                    value: value.map(|token| token.value.clone()),
                });
                index += 1 + usize::from(value.is_some());
            },
        }
    }

    Some(TableOfAuthoritiesParts {
        options,
        unknown_switches,
    })
}

pub(crate) fn validate_story_events(
    text: &str,
    shapes: &[crate::Shape<'_>],
    shape_groups: &[crate::ShapeGroup<'_>],
    drawing_order: &[crate::StoryDrawing],
    events: &[StoryEvent],
    label: &str,
) -> crate::RtfResult<()> {
    crate::shape::validate_story_drawings(text, shapes, shape_groups, drawing_order, label)?;
    let mut drawings = Vec::with_capacity(drawing_order.len());
    let mut fields = std::collections::BTreeSet::new();
    let mut previous = None;
    for event in events {
        let position = match *event {
            StoryEvent::PageBreak(value) => {
                if text.get(value.position..value.position).is_none() {
                    return Err(crate::RtfError::MalformedDocument(format!(
                        "RTF {label} page break is not at a UTF-8 boundary"
                    )));
                }
                value.position
            },
            StoryEvent::Drawing(drawing) => {
                drawings.push(drawing);
                match drawing {
                    crate::StoryDrawing::Shape(index) if index < shapes.len() => {
                        shapes[index].position
                    },
                    crate::StoryDrawing::ShapeGroup(index) if index < shape_groups.len() => {
                        shape_groups[index].position
                    },
                    _ => {
                        return Err(crate::RtfError::MalformedDocument(format!(
                            "RTF {label} story order has an invalid drawing reference"
                        )));
                    },
                }
            },
            StoryEvent::Field(field) => {
                if !fields.insert(field.field_index)
                    || text.get(field.position..field.position).is_none()
                {
                    return Err(crate::RtfError::MalformedDocument(format!(
                        "RTF {label} story order has an invalid or duplicate field reference"
                    )));
                }
                field.position
            },
        };
        if previous.is_some_and(|value| value > position) {
            return Err(crate::RtfError::MalformedDocument(format!(
                "RTF {label} story order moves backwards"
            )));
        }
        previous = Some(position);
    }
    if drawings != drawing_order {
        return Err(crate::RtfError::MalformedDocument(format!(
            "RTF {label} story order is incomplete or changes drawing order"
        )));
    }
    Ok(())
}

pub(crate) fn push_story_page_break(
    events: &mut Vec<StoryEvent>,
    text: &str,
    position: usize,
    label: &str,
) -> crate::RtfResult<()> {
    if text.get(position..position).is_none() {
        return Err(crate::RtfError::MalformedDocument(format!(
            "RTF {label} page break is not at a UTF-8 boundary"
        )));
    }
    events.push(StoryEvent::PageBreak(PageBreak::new(position)));
    Ok(())
}

/// Parse a field instruction without evaluating it.
pub fn parse_field_code(instruction: &str) -> ParsedFieldCode<'_> {
    match parse_field_code_inner(instruction) {
        Ok(parsed) => parsed,
        Err(error) => ParsedFieldCode::Malformed(error),
    }
}

fn parse_field_code_inner(instruction: &str) -> Result<ParsedFieldCode<'_>, FieldCodeError> {
    let mut tokens = tokenize(instruction)?;
    if tokens.is_empty() {
        return Err(FieldCodeError::MissingKeyword);
    }
    let keyword = tokens.remove(0);
    if keyword.value.eq_ignore_ascii_case("HYPERLINK") {
        return parse_hyperlink(tokens).map(ParsedFieldCode::Hyperlink);
    }
    for (name, constructor) in [("REF", 0u8), ("PAGEREF", 1u8), ("NOTEREF", 2u8)] {
        if keyword.value.eq_ignore_ascii_case(name) {
            let code = parse_reference(tokens)?;
            return Ok(match constructor {
                0 => ParsedFieldCode::Reference(code),
                1 => ParsedFieldCode::PageReference(code),
                _ => ParsedFieldCode::NoteReference(code),
            });
        }
    }
    Ok(ParsedFieldCode::Other {
        keyword: keyword.value,
        arguments: tokens,
    })
}

fn parse_hyperlink(tokens: Vec<FieldCodeToken<'_>>) -> Result<HyperlinkCode<'_>, FieldCodeError> {
    let mut code = HyperlinkCode {
        external_target: None,
        bookmark: None,
        screen_tip: None,
        target_frame: None,
        coordinates: None,
        new_window: false,
        unknown_switches: Vec::new(),
    };
    let mut index = 0;
    while index < tokens.len() {
        let token = &tokens[index];
        if let Some(name) = switch_name(token) {
            let normalized = name.to_ascii_lowercase();
            match normalized.as_str() {
                "n" => {
                    if code.new_window {
                        return Err(FieldCodeError::DuplicateOperand("\\n"));
                    }
                    code.new_window = true;
                    index += 1;
                },
                "l" | "o" | "t" | "m" => {
                    let value = switch_value(&tokens, index, name)?;
                    let slot = match normalized.as_str() {
                        "l" => &mut code.bookmark,
                        "o" => &mut code.screen_tip,
                        "t" => &mut code.target_frame,
                        _ => &mut code.coordinates,
                    };
                    if slot.replace(value).is_some() {
                        return Err(FieldCodeError::DuplicateOperand(
                            match normalized.as_str() {
                                "l" => "\\l",
                                "o" => "\\o",
                                "t" => "\\t",
                                _ => "\\m",
                            },
                        ));
                    }
                    index += 2;
                },
                _ => {
                    let value = tokens
                        .get(index + 1)
                        .filter(|next| switch_name(next).is_none());
                    code.unknown_switches.push(FieldSwitch {
                        name: Cow::Owned(name.to_string()),
                        value: value.map(|token| token.value.clone()),
                    });
                    index += 1 + usize::from(value.is_some());
                },
            }
        } else {
            if code.external_target.replace(token.value.clone()).is_some() {
                return Err(FieldCodeError::UnexpectedOperand(token.value.to_string()));
            }
            index += 1;
        }
    }
    if code.external_target.is_none() && code.bookmark.is_none() {
        return Err(FieldCodeError::MissingOperand(
            "hyperlink target or \\l bookmark",
        ));
    }
    Ok(code)
}

fn parse_reference(tokens: Vec<FieldCodeToken<'_>>) -> Result<ReferenceCode<'_>, FieldCodeError> {
    let Some(first) = tokens.first() else {
        return Err(FieldCodeError::MissingOperand("bookmark"));
    };
    if switch_name(first).is_some() {
        return Err(FieldCodeError::MissingOperand("bookmark"));
    }
    let mut code = ReferenceCode {
        bookmark: first.value.clone(),
        hyperlink: false,
        position: false,
        footnote_mark: false,
        unknown_switches: Vec::new(),
    };
    let mut index = 1;
    while index < tokens.len() {
        let token = &tokens[index];
        let Some(name) = switch_name(token) else {
            return Err(FieldCodeError::UnexpectedOperand(token.value.to_string()));
        };
        match name.to_ascii_lowercase().as_str() {
            "h" if !code.hyperlink => code.hyperlink = true,
            "p" if !code.position => code.position = true,
            "f" if !code.footnote_mark => code.footnote_mark = true,
            "h" => return Err(FieldCodeError::DuplicateOperand("\\h")),
            "p" => return Err(FieldCodeError::DuplicateOperand("\\p")),
            "f" => return Err(FieldCodeError::DuplicateOperand("\\f")),
            _ => {
                let value = tokens
                    .get(index + 1)
                    .filter(|next| switch_name(next).is_none());
                code.unknown_switches.push(FieldSwitch {
                    name: Cow::Owned(name.to_string()),
                    value: value.map(|token| token.value.clone()),
                });
                if value.is_some() {
                    index += 1;
                }
            },
        }
        index += 1;
    }
    Ok(code)
}

fn switch_name<'a>(token: &'a FieldCodeToken<'_>) -> Option<&'a str> {
    if token.quoted {
        return None;
    }
    token
        .value
        .strip_prefix('\\')
        .filter(|name| !name.is_empty())
}

fn switch_value<'a>(
    tokens: &[FieldCodeToken<'a>],
    index: usize,
    name: &str,
) -> Result<Cow<'a, str>, FieldCodeError> {
    let value = tokens
        .get(index + 1)
        .filter(|value| switch_name(value).is_none())
        .ok_or(FieldCodeError::MissingOperand("switch value"))?;
    if name.is_empty() {
        return Err(FieldCodeError::MissingOperand("switch name"));
    }
    Ok(value.value.clone())
}

fn tokenize(instruction: &str) -> Result<Vec<FieldCodeToken<'_>>, FieldCodeError> {
    if instruction.len() > MAX_INSTRUCTION_LEN {
        return Err(FieldCodeError::InstructionTooLong);
    }
    let bytes = instruction.as_bytes();
    let mut tokens = Vec::new();
    let mut index = 0;
    while index < bytes.len() {
        while index < bytes.len() && bytes[index].is_ascii_whitespace() {
            index += 1;
        }
        if index == bytes.len() {
            break;
        }
        if tokens.len() >= MAX_TOKENS {
            return Err(FieldCodeError::TooManyTokens);
        }
        if bytes[index] == b'"' {
            index += 1;
            let mut value = String::new();
            let mut closed = false;
            while index < bytes.len() {
                match bytes[index] {
                    b'"' => {
                        index += 1;
                        closed = true;
                        break;
                    },
                    b'\\'
                        if index + 1 < bytes.len() && matches!(bytes[index + 1], b'\\' | b'"') =>
                    {
                        value.push(bytes[index + 1] as char);
                        index += 2;
                    },
                    _ => {
                        let character = instruction[index..]
                            .chars()
                            .next()
                            .expect("index is inside instruction");
                        value.push(character);
                        index += character.len_utf8();
                    },
                }
            }
            if !closed {
                return Err(FieldCodeError::UnterminatedQuote);
            }
            tokens.push(FieldCodeToken {
                value: Cow::Owned(value),
                quoted: true,
            });
        } else {
            let start = index;
            while index < bytes.len() && !bytes[index].is_ascii_whitespace() {
                index += 1;
            }
            tokens.push(FieldCodeToken {
                value: Cow::Borrowed(&instruction[start..index]),
                quoted: false,
            });
        }
    }
    Ok(tokens)
}

pub(crate) fn quoted_field_operand(value: &str) -> String {
    let mut output = String::with_capacity(value.len() + 2);
    output.push('"');
    for character in value.chars() {
        match character {
            '\\' => output.push_str("\\\\"),
            '"' => output.push_str("\\\""),
            _ => output.push(character),
        }
    }
    output.push('"');
    output
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn exact_case_insensitive_keywords_and_distinct_references() {
        assert!(matches!(
            parse_field_code("hyperlink \"https://e\""),
            ParsedFieldCode::Hyperlink(_)
        ));
        for invalid in ["HYPERLINKER x", "REFRESH x", "PAGEREFERENCE x"] {
            assert!(matches!(
                parse_field_code(invalid),
                ParsedFieldCode::Other { .. }
            ));
            assert_eq!(
                Field::parse_instruction(invalid).field_type,
                FieldType::Unknown
            );
        }
        assert!(matches!(
            parse_field_code("REF mark \\h"),
            ParsedFieldCode::Reference(_)
        ));
        assert!(matches!(
            parse_field_code("PAGEREF mark \\p"),
            ParsedFieldCode::PageReference(_)
        ));
        assert!(matches!(
            parse_field_code("NOTEREF mark \\f"),
            ParsedFieldCode::NoteReference(_)
        ));
    }

    #[test]
    fn parses_internal_external_and_switch_semantics() {
        let ParsedFieldCode::Hyperlink(code) = parse_field_code(
            r#"HYPERLINK "https://example/a b" \l "_Toc1" \o "Tip" \t "_blank" \n"#,
        ) else {
            panic!("expected hyperlink");
        };
        assert_eq!(code.external_target.as_deref(), Some("https://example/a b"));
        assert_eq!(code.bookmark.as_deref(), Some("_Toc1"));
        assert_eq!(code.screen_tip.as_deref(), Some("Tip"));
        assert_eq!(code.target_frame.as_deref(), Some("_blank"));
        assert!(code.new_window);
        let field = Field::parse_instruction(r#"HYPERLINK \l "_Toc1""#);
        assert_eq!(field.extract_url().as_deref(), Some("#_Toc1"));
        assert_eq!(field.extract_bookmark().as_deref(), Some("_Toc1"));
    }

    #[test]
    fn writer_operand_cannot_inject_switches_and_round_trips_specials() {
        let target = "c:\\docs\\a \" \\l \"attacker{one}";
        let instruction = format!("HYPERLINK {}", quoted_field_operand(target));
        let ParsedFieldCode::Hyperlink(code) = parse_field_code(&instruction) else {
            panic!("expected hyperlink");
        };
        assert_eq!(code.external_target.as_deref(), Some(target));
        assert!(code.bookmark.is_none());

        let mut rtf = br#"{\rtf1\ansi "#.to_vec();
        crate::RtfWriter::new(&mut rtf)
            .write_hyperlink(target, "safe link")
            .unwrap();
        rtf.push(b'}');
        let document = crate::RtfDocument::from_bytes(&rtf).unwrap();
        let ParsedFieldCode::Hyperlink(code) = document.fields()[0].parsed_code() else {
            panic!("expected serialized hyperlink");
        };
        assert_eq!(code.external_target.as_deref(), Some(target));
        assert!(code.bookmark.is_none());
    }

    #[test]
    fn malformed_recognized_fields_are_non_actionable() {
        for instruction in [
            "HYPERLINK",
            r#"HYPERLINK "unterminated"#,
            r#"HYPERLINK \l"#,
            r#"HYPERLINK x \l a \l b"#,
            "REF",
            r#"REF a \h \h"#,
        ] {
            assert!(matches!(
                parse_field_code(instruction),
                ParsedFieldCode::Malformed(_)
            ));
        }
    }

    #[test]
    fn equation_fields_preserve_opaque_expression_metadata() {
        let mut field = Field::parse_instruction(r"EQ \o\ac(\fs24 Q,\fs16 R)");
        field.status = FieldStatus {
            dirty: true,
            locked: true,
            ..FieldStatus::default()
        };
        field.owner = FieldOwner::Body;
        field.position = 4;

        let equation = field.equation().unwrap();
        assert_eq!(equation.instruction(), r"EQ \o\ac(\fs24 Q,\fs16 R)");
        assert_eq!(equation.expression(), r"\o\ac(\fs24 Q,\fs16 R)");
        assert_eq!(equation.cached_result(), None);
        assert!(equation.is_dirty());
        assert!(equation.is_locked());
        assert_eq!(equation.owner(), FieldOwner::Body);
        assert_eq!(equation.position(), 4);

        let authored = Field::new_equation(r"\f(1,2)").unwrap();
        assert_eq!(authored.field_type, FieldType::Equation);
        assert_eq!(authored.equation().unwrap().expression(), r"\f(1,2)");
        assert!(Field::new_equation("x".repeat(MAX_INSTRUCTION_LEN)).is_err());
    }

    #[test]
    fn macro_button_fields_expose_stored_metadata_without_execution() {
        let mut field = Field::parse_instruction(r#"MACROBUTTON NoMacro "Click here""#);
        field.result = Cow::Borrowed("Click here");
        field.status = FieldStatus {
            dirty: true,
            locked: true,
            ..FieldStatus::default()
        };
        field.owner = FieldOwner::Body;
        field.position = 4;

        assert_eq!(field.field_type, FieldType::MacroButton);
        let macro_button = field.macro_button().unwrap();
        assert_eq!(macro_button.instruction(), r#"MACROBUTTON NoMacro "Click here""#);
        assert_eq!(macro_button.macro_name(), "NoMacro");
        assert_eq!(macro_button.display_text(), Some("Click here"));
        assert_eq!(macro_button.cached_result(), Some("Click here"));
        assert!(macro_button.is_dirty());
        assert!(macro_button.is_locked());
        assert_eq!(macro_button.owner(), FieldOwner::Body);
        assert_eq!(macro_button.position(), 4);

        let multiword = Field::parse_instruction("MACROBUTTON NoMacro Click here now");
        assert_eq!(multiword.macro_button().unwrap().display_text(), Some("Click here now"));
        assert!(Field::parse_instruction("MACROBUTTON").macro_button().is_none());
        assert!(Field::parse_instruction(r#"MACROBUTTON "" "button""#)
            .macro_button()
            .is_none());
    }

    #[test]
    fn dde_fields_expose_stored_metadata_without_contacting_sources() {
        let mut field = Field::parse_instruction(
            r#"DDE Excel "C:\\no-contact\\source.xlsx" "Sheet1!R1C1:R4C4" \a \p \* MERGEFORMAT"#,
        );
        field.result = Cow::Borrowed("cached DDE result");
        field.status = FieldStatus {
            dirty: true,
            locked: true,
            ..FieldStatus::default()
        };
        field.owner = FieldOwner::Body;
        field.position = 4;

        assert_eq!(field.field_type, FieldType::Dde);
        let dde = field.dde_link().unwrap();
        assert_eq!(dde.instruction(), r#"DDE Excel "C:\\no-contact\\source.xlsx" "Sheet1!R1C1:R4C4" \a \p \* MERGEFORMAT"#);
        assert_eq!(dde.kind(), DdeFieldKind::Dde);
        assert_eq!(dde.application(), "Excel");
        assert_eq!(dde.source(), r"C:\no-contact\source.xlsx");
        assert_eq!(dde.item(), Some("Sheet1!R1C1:R4C4"));
        assert!(dde.requests_automatic_updates());
        assert_eq!(dde.representation(), Some(DdeRepresentation::Picture));
        assert!(!dde.omits_graphic_data());
        assert_eq!(dde.cached_result(), Some("cached DDE result"));
        assert!(dde.is_dirty());
        assert!(dde.is_locked());
        assert_eq!(dde.owner(), FieldOwner::Body);
        assert_eq!(dde.position(), 4);
        assert_eq!(dde.unknown_switches().len(), 1);
        assert_eq!(dde.unknown_switches()[0].name, "*");
        assert_eq!(
            dde.unknown_switches()[0].value.as_deref(),
            Some("MERGEFORMAT")
        );

        let automatic = Field::parse_instruction(r#"DDEAUTO Excel "missing.xlsx" "Sheet1!A1" \t"#);
        assert_eq!(automatic.field_type, FieldType::DdeAuto);
        let automatic = automatic.dde_link().unwrap();
        assert_eq!(automatic.kind(), DdeFieldKind::DdeAuto);
        assert!(automatic.requests_automatic_updates());
        assert_eq!(automatic.representation(), Some(DdeRepresentation::Text));
        assert!(!automatic.omits_graphic_data());

        let omit_graphics = Field::parse_instruction(r#"DDE Excel source \a \d"#);
        let omit_graphics = omit_graphics.dde_link().unwrap();
        assert!(omit_graphics.requests_automatic_updates());
        assert_eq!(omit_graphics.representation(), None);
        assert!(omit_graphics.omits_graphic_data());

        assert!(Field::parse_instruction("DDE").dde_link().is_none());
        assert!(Field::parse_instruction(r"DDE Excel \p").dde_link().is_none());
        assert!(
            Field::parse_instruction(r"DDE Excel source \p unexpected")
                .dde_link()
                .is_none()
        );
        assert!(
            Field::parse_instruction(r"DDE Excel source \p \t")
                .dde_link()
                .is_none()
        );
        assert!(
            Field::parse_instruction(r"DDEAUTO Excel source \p \t")
                .dde_link()
                .is_none()
        );
        assert!(
            Field::parse_instruction(r"DDEAUTO Excel source \a")
                .dde_link()
                .is_none()
        );
        assert_eq!(
            Field::parse_instruction("DDEAUTOMATED Excel source").field_type,
            FieldType::Unknown
        );
    }

    #[test]
    fn link_fields_expose_stored_metadata_without_activating_sources() {
        let mut field = Field::parse_instruction(
            r#"LINK Excel.Sheet.8 "C:\\no-contact\\source.xlsx" "Sheet1!R1C1:R4C4" \a \f 4 \p \d \* MERGEFORMAT"#,
        );
        field.result = Cow::Borrowed("cached LINK result");
        field.status = FieldStatus {
            dirty: true,
            locked: true,
            ..FieldStatus::default()
        };
        field.owner = FieldOwner::Body;
        field.position = 4;

        assert_eq!(field.field_type, FieldType::Link);
        let link = field.link_field().unwrap();
        assert_eq!(
            link.instruction(),
            r#"LINK Excel.Sheet.8 "C:\\no-contact\\source.xlsx" "Sheet1!R1C1:R4C4" \a \f 4 \p \d \* MERGEFORMAT"#
        );
        assert_eq!(link.application_type(), "Excel.Sheet.8");
        assert_eq!(link.source(), r"C:\no-contact\source.xlsx");
        assert_eq!(link.item(), Some("Sheet1!R1C1:R4C4"));
        assert!(link.requests_automatic_updates());
        assert_eq!(
            link.result_options(),
            &[LinkResultOption::Picture, LinkResultOption::OmitGraphicData]
        );
        assert_eq!(
            link.effective_result_option(),
            Some(LinkResultOption::OmitGraphicData)
        );
        assert_eq!(
            link.formatting_modes(),
            &[LinkFormatting::SpreadsheetSource]
        );
        assert_eq!(link.cached_result(), Some("cached LINK result"));
        assert!(link.is_dirty());
        assert!(link.is_locked());
        assert_eq!(link.owner(), FieldOwner::Body);
        assert_eq!(link.position(), 4);
        assert_eq!(link.unknown_switches().len(), 1);
        assert_eq!(link.unknown_switches()[0].name, "*");
        assert_eq!(
            link.unknown_switches()[0].value.as_deref(),
            Some("MERGEFORMAT")
        );

        let destination =
            Field::parse_instruction(r#"LINK Word.Document.8 "missing.docx" Bookmark \f 2 \t"#);
        let destination = destination.link_field().unwrap();
        assert!(!destination.requests_automatic_updates());
        assert_eq!(
            destination.formatting_modes(),
            &[LinkFormatting::Destination]
        );
        assert_eq!(
            destination.effective_result_option(),
            Some(LinkResultOption::Text)
        );

        let unsupported = Field::parse_instruction(r"LINK Package source \f 1");
        assert_eq!(
            unsupported.link_field().unwrap().formatting_modes(),
            &[LinkFormatting::Unsupported(1)]
        );

        let multiple_formatting = Field::parse_instruction(r"LINK Excel.Sheet.8 source \f 0 \f 2");
        assert_eq!(
            multiple_formatting.link_field().unwrap().formatting_modes(),
            &[LinkFormatting::Source, LinkFormatting::Destination]
        );

        let repeated_updates = Field::parse_instruction(r"LINK Excel.Sheet.8 source \a \a");
        assert!(
            repeated_updates
                .link_field()
                .unwrap()
                .requests_automatic_updates()
        );

        assert!(Field::parse_instruction("LINK").link_field().is_none());
        assert!(
            Field::parse_instruction(r"LINK Excel.Sheet.8 \p")
                .link_field()
                .is_none()
        );
        assert!(
            Field::parse_instruction(r"LINK Excel.Sheet.8 source \f")
                .link_field()
                .is_none()
        );
        assert!(
            Field::parse_instruction(r"LINK Excel.Sheet.8 source \f invalid")
                .link_field()
                .is_none()
        );
        assert!(
            Field::parse_instruction(r"LINK Excel.Sheet.8 source \p unexpected")
                .link_field()
                .is_none()
        );
        assert_eq!(
            Field::parse_instruction("LINKAGE Excel.Sheet.8 source").field_type,
            FieldType::Unknown
        );
    }

    #[test]
    fn external_include_fields_expose_stored_metadata_without_resolution() {
        let mut include_text = Field::parse_instruction(
            r#"INCLUDETEXT "missing source.xml" Summary \! \c Word8 \e utf-8 \m application/xml \n "xmlns:a=\"resume-schema\"" \t "file:///C:/display.xsl" \x a:Resume/a:Name \* MERGEFORMAT"#,
        );
        include_text.result = Cow::Borrowed("cached text");
        include_text.status = FieldStatus {
            dirty: true,
            locked: true,
            ..FieldStatus::default()
        };
        include_text.owner = FieldOwner::Body;
        include_text.position = 4;

        assert_eq!(include_text.field_type, FieldType::IncludeText);
        let text = include_text.external_include().unwrap();
        assert_eq!(text.kind(), IncludeFieldKind::Text);
        assert_eq!(text.source(), "missing source.xml");
        assert_eq!(text.bookmark(), Some("Summary"));
        assert_eq!(text.converter(), Some("Word8"));
        assert_eq!(
            text.options(),
            &[
                ExternalIncludeOption::Converter(Cow::Borrowed("Word8")),
                ExternalIncludeOption::Encoding(Cow::Borrowed("utf-8")),
                ExternalIncludeOption::MimeType(Cow::Borrowed("application/xml")),
                ExternalIncludeOption::NamespaceMapping(Cow::Borrowed("xmlns:a=\"resume-schema\"")),
                ExternalIncludeOption::Xslt(Cow::Borrowed("file:///C:/display.xsl")),
                ExternalIncludeOption::XPath(Cow::Borrowed("a:Resume/a:Name")),
            ]
        );
        assert!(text.suppresses_nested_field_updates());
        assert!(!text.omits_picture_data());
        assert_eq!(text.cached_result(), Some("cached text"));
        assert!(text.is_dirty());
        assert!(text.is_locked());
        assert_eq!(text.owner(), FieldOwner::Body);
        assert_eq!(text.position(), 4);
        assert_eq!(text.unknown_switches().len(), 1);
        assert_eq!(text.unknown_switches()[0].name, "*");
        assert_eq!(
            text.unknown_switches()[0].value.as_deref(),
            Some("MERGEFORMAT")
        );

        let unc_source = Field::parse_instruction(r#"INCLUDETEXT "\\server\\share\\source.docx""#);
        assert_eq!(
            unc_source.external_include().unwrap().source(),
            r"\server\share\source.docx"
        );

        let include_picture = Field::parse_instruction(
            r#"INCLUDEPICTURE "missing picture.gif" \c Pictim32 \d \* MERGEFORMAT"#,
        );
        assert_eq!(include_picture.field_type, FieldType::IncludePicture);
        let picture = include_picture.external_include().unwrap();
        assert_eq!(picture.kind(), IncludeFieldKind::Picture);
        assert_eq!(picture.source(), "missing picture.gif");
        assert_eq!(picture.bookmark(), None);
        assert_eq!(picture.converter(), Some("Pictim32"));
        assert_eq!(
            picture.options(),
            &[ExternalIncludeOption::Converter(Cow::Borrowed("Pictim32"))]
        );
        assert!(!picture.suppresses_nested_field_updates());
        assert!(picture.omits_picture_data());
        assert_eq!(picture.unknown_switches().len(), 1);
        assert_eq!(picture.unknown_switches()[0].name, "*");

        assert!(
            Field::parse_instruction("INCLUDETEXT")
                .external_include()
                .is_none()
        );
        assert!(
            Field::parse_instruction("INCLUDETEXT \\c Word8")
                .external_include()
                .is_none()
        );
        assert!(
            Field::parse_instruction(r#"INCLUDEPICTURE "picture.gif" Selector"#)
                .external_include()
                .is_none()
        );
        assert!(
            Field::parse_instruction(r#"INCLUDEPICTURE "picture.gif" \d extra"#)
                .external_include()
                .is_none()
        );
        assert!(
            Field::parse_instruction(r"INCLUDETEXT source \e")
                .external_include()
                .is_none()
        );
    }

    #[test]
    fn table_of_contents_fields_preserve_stored_configuration_without_generation() {
        let mut field = Field::parse_instruction(
            r#"TOC \a Figure \b "Scope Bookmark" \c Table \d "/" \f A \h \l 1-3 \n "2-3" \o "1-4" \p " — " \s Figure \t "Custom,1,Appendix,2" \u \w \x \z \* MERGEFORMAT"#,
        );
        field.result = Cow::Borrowed("cached TOC");
        field.status = FieldStatus {
            dirty: true,
            locked: true,
            ..FieldStatus::default()
        };
        field.owner = FieldOwner::Body;
        field.position = 4;

        assert_eq!(field.field_type, FieldType::Toc);
        let toc = field.table_of_contents().unwrap();
        assert_eq!(toc.instruction(), field.instruction);
        assert_eq!(
            toc.options(),
            &[
                TableOfContentsOption::CaptionWithoutLabel(Cow::Borrowed("Figure")),
                TableOfContentsOption::Bookmark(Cow::Borrowed("Scope Bookmark")),
                TableOfContentsOption::CaptionSequence(Cow::Borrowed("Table")),
                TableOfContentsOption::SequencePageSeparator(Cow::Borrowed("/")),
                TableOfContentsOption::TableEntryIdentifier(Cow::Borrowed("A")),
                TableOfContentsOption::Hyperlinks,
                TableOfContentsOption::TableEntryLevels(Cow::Borrowed("1-3")),
                TableOfContentsOption::OmitPageNumbers(Some(Cow::Borrowed("2-3"))),
                TableOfContentsOption::HeadingStyleRange(Some(Cow::Borrowed("1-4"))),
                TableOfContentsOption::EntryPageNumberSeparator(Cow::Borrowed(" — ")),
                TableOfContentsOption::SequenceIdentifier(Cow::Borrowed("Figure")),
                TableOfContentsOption::StyleMappings(Cow::Borrowed("Custom,1,Appendix,2")),
                TableOfContentsOption::OutlineLevels,
                TableOfContentsOption::PreserveTabs,
                TableOfContentsOption::PreserveNewlines,
                TableOfContentsOption::HidePageNumbersInWebView,
            ]
        );
        assert_eq!(toc.cached_result(), Some("cached TOC"));
        assert!(toc.is_dirty());
        assert!(toc.is_locked());
        assert_eq!(toc.owner(), FieldOwner::Body);
        assert_eq!(toc.position(), 4);
        assert_eq!(toc.unknown_switches().len(), 1);
        assert_eq!(toc.unknown_switches()[0].name, "*");
        assert_eq!(
            toc.unknown_switches()[0].value.as_deref(),
            Some("MERGEFORMAT")
        );

        let all_levels = Field::parse_instruction(r"TOC \n \o");
        assert_eq!(
            all_levels.table_of_contents().unwrap().options(),
            &[
                TableOfContentsOption::OmitPageNumbers(None),
                TableOfContentsOption::HeadingStyleRange(None),
            ]
        );

        assert!(
            Field::parse_instruction(r"TOC \a")
                .table_of_contents()
                .is_none()
        );
        assert!(
            Field::parse_instruction(r"TOC \h unexpected")
                .table_of_contents()
                .is_none()
        );
        assert!(
            Field::parse_instruction(r"TOC unexpected")
                .table_of_contents()
                .is_none()
        );
        assert_eq!(
            Field::parse_instruction("TOCENTRIES").field_type,
            FieldType::Unknown
        );
    }

    #[test]
    fn table_of_contents_entry_fields_preserve_stored_metadata_without_generation() {
        let mut field =
            Field::parse_instruction(r#"TC "Illustration 1" \f i \l 4 \n \* MERGEFORMAT"#);
        field.result = Cow::Borrowed("cached entry");
        field.status = FieldStatus {
            dirty: true,
            locked: true,
            ..FieldStatus::default()
        };
        field.owner = FieldOwner::Body;
        field.position = 4;

        assert_eq!(field.field_type, FieldType::TocEntry);
        let entry = field.table_of_contents_entry().unwrap();
        assert_eq!(entry.instruction(), field.instruction);
        assert_eq!(entry.entry(), "Illustration 1");
        assert_eq!(
            entry.options(),
            &[
                TableOfContentsEntryOption::ListIdentifier(Cow::Borrowed("i")),
                TableOfContentsEntryOption::Level(Cow::Borrowed("4")),
                TableOfContentsEntryOption::OmitPageNumber,
            ]
        );
        assert_eq!(entry.cached_result(), Some("cached entry"));
        assert!(entry.is_dirty());
        assert!(entry.is_locked());
        assert_eq!(entry.owner(), FieldOwner::Body);
        assert_eq!(entry.position(), 4);
        assert_eq!(entry.unknown_switches().len(), 1);
        assert_eq!(entry.unknown_switches()[0].name, "*");
        assert_eq!(
            entry.unknown_switches()[0].value.as_deref(),
            Some("MERGEFORMAT")
        );

        assert!(
            Field::parse_instruction("TC")
                .table_of_contents_entry()
                .is_none()
        );
        assert!(
            Field::parse_instruction(r"TC \f i")
                .table_of_contents_entry()
                .is_none()
        );
        assert!(
            Field::parse_instruction(r"TC entry unexpected")
                .table_of_contents_entry()
                .is_none()
        );
        assert!(
            Field::parse_instruction(r"TC entry \n unexpected")
                .table_of_contents_entry()
                .is_none()
        );
        assert_eq!(
            Field::parse_instruction("TCC entry").field_type,
            FieldType::Unknown
        );
    }

    #[test]
    fn table_of_authorities_entry_fields_preserve_stored_metadata_without_generation() {
        let mut field = Field::parse_instruction(
            r#"TA \l "Baldwin v. Alberti" \c 1 \s Baldwin \b \i \r PageRange \* MERGEFORMAT"#,
        );
        field.result = Cow::Borrowed("cached authority");
        field.status = FieldStatus {
            dirty: true,
            locked: true,
            ..FieldStatus::default()
        };
        field.owner = FieldOwner::Body;
        field.position = 4;

        assert_eq!(field.field_type, FieldType::TableOfAuthoritiesEntry);
        let entry = field.table_of_authorities_entry().unwrap();
        assert_eq!(entry.instruction(), field.instruction);
        assert_eq!(
            entry.options(),
            &[
                TableOfAuthoritiesEntryOption::LongCitation(Cow::Borrowed("Baldwin v. Alberti")),
                TableOfAuthoritiesEntryOption::Category(Cow::Borrowed("1")),
                TableOfAuthoritiesEntryOption::ShortCitation(Cow::Borrowed("Baldwin")),
                TableOfAuthoritiesEntryOption::BoldPageNumber,
                TableOfAuthoritiesEntryOption::ItalicPageNumber,
                TableOfAuthoritiesEntryOption::PageRangeBookmark(Cow::Borrowed("PageRange")),
            ]
        );
        assert_eq!(entry.cached_result(), Some("cached authority"));
        assert!(entry.is_dirty());
        assert!(entry.is_locked());
        assert_eq!(entry.owner(), FieldOwner::Body);
        assert_eq!(entry.position(), 4);
        assert_eq!(entry.unknown_switches().len(), 1);
        assert_eq!(entry.unknown_switches()[0].name, "*");
        assert_eq!(
            entry.unknown_switches()[0].value.as_deref(),
            Some("MERGEFORMAT")
        );

        assert!(
            Field::parse_instruction("TA")
                .table_of_authorities_entry()
                .is_some()
        );
        assert!(
            Field::parse_instruction(r"TA \c")
                .table_of_authorities_entry()
                .is_none()
        );
        assert!(
            Field::parse_instruction(r"TA \b unexpected")
                .table_of_authorities_entry()
                .is_none()
        );
        assert!(
            Field::parse_instruction("TA unexpected")
                .table_of_authorities_entry()
                .is_none()
        );
        assert_eq!(
            Field::parse_instruction(r"TAX \l Citation").field_type,
            FieldType::Unknown
        );
    }

    #[test]
    fn table_of_authorities_fields_preserve_stored_configuration_without_generation() {
        let mut field = Field::parse_instruction(
            r#"TOA \b Authorities \c 2 \d "-" \e " — " \f \g "–" \h \l ", " \p \s Section \* MERGEFORMAT"#,
        );
        field.result = Cow::Borrowed("cached authorities");
        field.status = FieldStatus {
            dirty: true,
            locked: true,
            ..FieldStatus::default()
        };
        field.owner = FieldOwner::Body;
        field.position = 4;

        assert_eq!(field.field_type, FieldType::TableOfAuthorities);
        let toa = field.table_of_authorities().unwrap();
        assert_eq!(toa.instruction(), field.instruction);
        assert_eq!(
            toa.options(),
            &[
                TableOfAuthoritiesOption::Bookmark(Cow::Borrowed("Authorities")),
                TableOfAuthoritiesOption::Category(Cow::Borrowed("2")),
                TableOfAuthoritiesOption::SequencePageSeparator(Cow::Borrowed("-")),
                TableOfAuthoritiesOption::EntryPageNumberSeparator(Cow::Borrowed(" — ")),
                TableOfAuthoritiesOption::RemoveEntryFormatting,
                TableOfAuthoritiesOption::PageRangeSeparator(Cow::Borrowed("–")),
                TableOfAuthoritiesOption::CategoryHeadings,
                TableOfAuthoritiesOption::PageReferenceSeparator(Cow::Borrowed(", ")),
                TableOfAuthoritiesOption::UsePassim,
                TableOfAuthoritiesOption::SequenceIdentifier(Cow::Borrowed("Section")),
            ]
        );
        assert_eq!(toa.cached_result(), Some("cached authorities"));
        assert!(toa.is_dirty());
        assert!(toa.is_locked());
        assert_eq!(toa.owner(), FieldOwner::Body);
        assert_eq!(toa.position(), 4);
        assert_eq!(toa.unknown_switches().len(), 1);
        assert_eq!(toa.unknown_switches()[0].name, "*");
        assert_eq!(
            toa.unknown_switches()[0].value.as_deref(),
            Some("MERGEFORMAT")
        );

        assert!(
            Field::parse_instruction("TOA")
                .table_of_authorities()
                .is_some()
        );
        assert!(
            Field::parse_instruction(r"TOA \b")
                .table_of_authorities()
                .is_none()
        );
        assert!(
            Field::parse_instruction(r"TOA \f unexpected")
                .table_of_authorities()
                .is_none()
        );
        assert!(
            Field::parse_instruction("TOA unexpected")
                .table_of_authorities()
                .is_none()
        );
        assert_eq!(
            Field::parse_instruction(r"TOAX \c 2").field_type,
            FieldType::Unknown
        );
    }

    #[test]
    fn document_discovers_eq_fields_without_calculating_them() {
        let document = crate::RtfDocument::parse(
            r#"{\rtf1\ansi Before {\field{\*\fldinst EQ \\f(1,2)}{\fldrslt }}After}"#,
        )
        .unwrap();

        let equations = document.equations();
        assert_eq!(document.equation_count(), 1);
        assert_eq!(equations.len(), 1);
        assert_eq!(equations[0].expression(), r"\f(1,2)");
        assert_eq!(equations[0].cached_result(), None);
        assert_eq!(document.text(), "Before After");
    }

    #[test]
    fn document_discovers_dde_fields_without_starting_conversations() {
        let document = crate::RtfDocument::parse(
            r#"{\rtf1\ansi Before {\field\flddirty{\*\fldinst DDE Excel "missing.xlsx" "Sheet1!A1" \\a \\p}{\fldrslt cached DDE}}Middle {\field{\*\fldinst DDEAUTO Excel "missing.xlsx" "Sheet1!A2" \\t}{\fldrslt cached auto}}After}"#,
        )
        .unwrap();

        let links = document.dde_links();
        assert_eq!(document.dde_link_count(), 2);
        assert_eq!(links.len(), 2);
        assert_eq!(links[0].kind(), DdeFieldKind::Dde);
        assert_eq!(links[0].application(), "Excel");
        assert_eq!(links[0].source(), "missing.xlsx");
        assert_eq!(links[0].item(), Some("Sheet1!A1"));
        assert!(links[0].requests_automatic_updates());
        assert_eq!(
            links[0].representation(),
            Some(DdeRepresentation::Picture)
        );
        assert_eq!(links[0].cached_result(), Some("cached DDE"));
        assert!(links[0].is_dirty());
        assert_eq!(links[1].kind(), DdeFieldKind::DdeAuto);
        assert_eq!(links[1].item(), Some("Sheet1!A2"));
        assert_eq!(
            links[1].representation(),
            Some(DdeRepresentation::Text)
        );
        assert_eq!(links[1].cached_result(), Some("cached auto"));
        assert_eq!(document.text(), "Before Middle After");
    }

    #[test]
    fn document_discovers_link_fields_without_activating_ole_servers() {
        let document = crate::RtfDocument::parse(
            r#"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst LINK Excel.Sheet.8 "missing.xlsx" "Sheet1!A1" \\a \\f 4 \\p}{\fldrslt cached LINK}}Middle {\field{\*\fldinst LINK Word.Document.8 "missing.docx" Bookmark \\t}{\fldrslt cached text}}After}"#,
        )
        .unwrap();

        let links = document.link_fields();
        assert_eq!(document.link_field_count(), 2);
        assert_eq!(links.len(), 2);
        assert_eq!(links[0].application_type(), "Excel.Sheet.8");
        assert_eq!(links[0].source(), "missing.xlsx");
        assert_eq!(links[0].item(), Some("Sheet1!A1"));
        assert!(links[0].requests_automatic_updates());
        assert_eq!(
            links[0].formatting_modes(),
            &[LinkFormatting::SpreadsheetSource]
        );
        assert_eq!(
            links[0].effective_result_option(),
            Some(LinkResultOption::Picture)
        );
        assert_eq!(links[0].cached_result(), Some("cached LINK"));
        assert!(links[0].is_dirty());
        assert!(links[0].is_locked());
        assert_eq!(links[1].application_type(), "Word.Document.8");
        assert_eq!(links[1].item(), Some("Bookmark"));
        assert_eq!(
            links[1].effective_result_option(),
            Some(LinkResultOption::Text)
        );
        assert_eq!(links[1].cached_result(), Some("cached text"));
        assert_eq!(document.text(), "Before Middle After");
    }

    #[test]
    fn document_discovers_external_includes_without_opening_sources() {
        let document = crate::RtfDocument::parse(
            r#"{\rtf1\ansi Before {\field\flddirty{\*\fldinst INCLUDETEXT "missing.docx" Summary \\!}{\fldrslt cached text}}Middle {\field{\*\fldinst INCLUDEPICTURE "missing.gif" \\d}{\fldrslt cached picture}}After}"#,
        )
        .unwrap();

        let includes = document.external_includes();
        assert_eq!(document.external_include_count(), 2);
        assert_eq!(includes.len(), 2);
        assert_eq!(includes[0].kind(), IncludeFieldKind::Text);
        assert_eq!(includes[0].source(), "missing.docx");
        assert_eq!(includes[0].bookmark(), Some("Summary"));
        assert!(includes[0].suppresses_nested_field_updates());
        assert!(includes[0].is_dirty());
        assert_eq!(includes[1].kind(), IncludeFieldKind::Picture);
        assert_eq!(includes[1].source(), "missing.gif");
        assert!(includes[1].omits_picture_data());
        assert_eq!(document.text(), "Before Middle After");
    }

    #[test]
    fn document_discovers_table_of_contents_without_regenerating_it() {
        let document = crate::RtfDocument::parse(
            r#"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst TOC \\o "1-3" \\h \\z}{\fldrslt cached TOC}}After}"#,
        )
        .unwrap();

        let tables = document.table_of_contents();
        assert_eq!(document.table_of_contents_count(), 1);
        assert_eq!(tables.len(), 1);
        assert_eq!(tables[0].cached_result(), Some("cached TOC"));
        assert!(tables[0].is_dirty());
        assert!(tables[0].is_locked());
        assert_eq!(
            tables[0].options(),
            &[
                TableOfContentsOption::HeadingStyleRange(Some(Cow::Borrowed("1-3"))),
                TableOfContentsOption::Hyperlinks,
                TableOfContentsOption::HidePageNumbersInWebView,
            ]
        );
        assert_eq!(document.text(), "Before After");
    }

    #[test]
    fn document_discovers_table_of_contents_entries_without_generating_a_table() {
        let document = crate::RtfDocument::parse(
            r#"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst TC "Illustration 1" \\f i \\l 4 \\n}{\fldrslt cached entry}}After}"#,
        )
        .unwrap();

        let entries = document.table_of_contents_entries();
        assert_eq!(document.table_of_contents_entry_count(), 1);
        assert_eq!(entries.len(), 1);
        assert_eq!(entries[0].entry(), "Illustration 1");
        assert_eq!(entries[0].cached_result(), Some("cached entry"));
        assert!(entries[0].is_dirty());
        assert!(entries[0].is_locked());
        assert_eq!(
            entries[0].options(),
            &[
                TableOfContentsEntryOption::ListIdentifier(Cow::Borrowed("i")),
                TableOfContentsEntryOption::Level(Cow::Borrowed("4")),
                TableOfContentsEntryOption::OmitPageNumber,
            ]
        );
        assert_eq!(document.text(), "Before After");
    }

    #[test]
    fn document_discovers_table_of_authorities_entries_without_generating_a_table() {
        let document = crate::RtfDocument::parse(
            r#"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst TA \\l "Baldwin v. Alberti" \\c 1 \\b}{\fldrslt cached authority}}After}"#,
        )
        .unwrap();

        let entries = document.table_of_authorities_entries();
        assert_eq!(document.table_of_authorities_entry_count(), 1);
        assert_eq!(entries.len(), 1);
        assert_eq!(entries[0].cached_result(), Some("cached authority"));
        assert!(entries[0].is_dirty());
        assert!(entries[0].is_locked());
        assert_eq!(
            entries[0].options(),
            &[
                TableOfAuthoritiesEntryOption::LongCitation(Cow::Borrowed("Baldwin v. Alberti")),
                TableOfAuthoritiesEntryOption::Category(Cow::Borrowed("1")),
                TableOfAuthoritiesEntryOption::BoldPageNumber,
            ]
        );
        assert_eq!(document.text(), "Before After");
    }

    #[test]
    fn document_discovers_table_of_authorities_without_generating_it() {
        let document = crate::RtfDocument::parse(
            r#"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst TOA \\b Authorities \\c 2 \\f \\h \\p}{\fldrslt cached authorities}}After}"#,
        )
        .unwrap();

        let tables = document.tables_of_authorities();
        assert_eq!(document.table_of_authorities_count(), 1);
        assert_eq!(tables.len(), 1);
        assert_eq!(tables[0].cached_result(), Some("cached authorities"));
        assert!(tables[0].is_dirty());
        assert!(tables[0].is_locked());
        assert_eq!(
            tables[0].options(),
            &[
                TableOfAuthoritiesOption::Bookmark(Cow::Borrowed("Authorities")),
                TableOfAuthoritiesOption::Category(Cow::Borrowed("2")),
                TableOfAuthoritiesOption::RemoveEntryFormatting,
                TableOfAuthoritiesOption::CategoryHeadings,
                TableOfAuthoritiesOption::UsePassim,
            ]
        );
        assert_eq!(document.text(), "Before After");
    }

    #[test]
    fn document_discovers_macro_buttons_without_invoking_them() {
        let document = crate::RtfDocument::parse(
            r#"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst MACROBUTTON NoMacro Click here}{\fldrslt Click here}}After}"#,
        )
        .unwrap();

        let macro_buttons = document.macro_buttons();
        assert_eq!(document.macro_button_count(), 1);
        assert_eq!(macro_buttons.len(), 1);
        assert_eq!(macro_buttons[0].macro_name(), "NoMacro");
        assert_eq!(macro_buttons[0].display_text(), Some("Click here"));
        assert_eq!(macro_buttons[0].cached_result(), Some("Click here"));
        assert!(macro_buttons[0].is_dirty());
        assert!(macro_buttons[0].is_locked());
        assert_eq!(document.text(), "Before After");
    }

    #[test]
    fn parses_libreoffice_internal_hyperlink_fixtures() {
        let fixture_root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/rtf");
        for (fixture, expected) in [
            ("fdo86750.rtf", "anchor"),
            ("tdf134614_toc_indent.rtf", "_Toc1"),
        ] {
            let document = crate::RtfDocument::from_bytes(
                &std::fs::read(fixture_root.join(fixture)).unwrap(),
            )
            .unwrap();
            assert!(
                document.fields().iter().any(|field| {
                    field.extract_bookmark().as_deref() == Some(expected)
                        && field.extract_url().as_deref() == Some(format!("#{expected}").as_str())
                }),
                "fixture {fixture} fields: {:?}",
                document.fields()
            );
        }

        let formatted = crate::RtfDocument::from_bytes(
            &std::fs::read(fixture_root.join("fdo82071.rtf")).unwrap(),
        )
        .unwrap();
        assert!(formatted.fields().iter().any(|field| matches!(
            field.parsed_code(),
            ParsedFieldCode::PageReference(ReferenceCode { ref bookmark, hyperlink: true, .. })
                if bookmark == "_Toc363816075"
        )));

        let backslashes = crate::RtfDocument::from_bytes(
            &std::fs::read(fixture_root.join("hyperlink-with-backslashes.rtf")).unwrap(),
        )
        .unwrap();
        assert_eq!(
            backslashes.fields()[0].extract_url().as_deref(),
            Some(r"c:\temp\doc1.doc")
        );

        let target = crate::RtfDocument::from_bytes(
            &std::fs::read(fixture_root.join("hyperlink-target.rtf")).unwrap(),
        )
        .unwrap();
        let ParsedFieldCode::Hyperlink(code) = target.fields()[0].parsed_code() else {
            panic!("expected target-frame hyperlink");
        };
        assert_eq!(code.target_frame.as_deref(), Some("_blank"));
    }
}
