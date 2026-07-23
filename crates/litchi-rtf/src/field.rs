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
    Advance,
    MacroButton,
    GoToButton,
    Print,
    Embed,
    Barcode,
    BidiOutline,
    Shape,
    FormText,
    FormCheckbox,
    FormDropdown,
    AddIn,
    Control,
    HtmlControl,
    Glossary,
    AutoText,
    AutoTextList,
    UserAddress,
    UserInitials,
    UserName,
    Dde,
    DdeAuto,
    Link,
    IncludeText,
    IncludePicture,
    Index,
    IndexEntry,
    Citation,
    Bibliography,
    DocumentVariable,
    DocumentProperty,
    Info,
    DocumentInformation,
    DocumentContext,
    MergeField,
    MailMergeData,
    MergeRecord,
    MergeSequence,
    MailMergeNext,
    MailMergeNextIf,
    MailMergeSkipIf,
    If,
    Compare,
    Set,
    Sequence,
    Formula,
    Quote,
    Symbol,
    AutoNumber,
    ListNumber,
    StyleReference,
    Ask,
    FillIn,
    AddressBlock,
    GreetingLine,
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

/// Inert metadata for a legacy RTF `GOTOBUTTON` field.
///
/// The destination and button text are exposed solely as stored field
/// metadata. This crate never resolves a destination, changes the insertion
/// point, or activates a jump.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct GoToButtonField<'a> {
    instruction: &'a str,
    target: Cow<'a, str>,
    button_text: Cow<'a, str>,
    cached_result: Option<&'a str>,
    status: FieldStatus,
    owner: FieldOwner,
    position: usize,
}

/// The stored kind of a legacy RTF active-content field.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ActiveContentFieldKind {
    /// An `ADDIN` field that stores add-in-created data.
    AddIn,
    /// A `CONTROL` field that represents an OCX control.
    OcxControl,
    /// An `HTMLCONTROL` field that represents an HTML control.
    HtmlControl,
}

/// Inert metadata for a legacy RTF add-in or control field.
///
/// ECMA-376 Part 1 §17.16.5 defines `ADDIN`, `CONTROL`, and
/// `HTMLCONTROL` field instructions. This type retains only the stored
/// category, instruction, cached result, and state. It never loads an add-in,
/// instantiates an OCX or HTML control, invokes code, executes script, renders
/// content, or accesses an external resource.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ActiveContentField<'a> {
    instruction: &'a str,
    kind: ActiveContentFieldKind,
    cached_result: Option<&'a str>,
    status: FieldStatus,
    owner: FieldOwner,
    position: usize,
}

/// Inert metadata for a legacy RTF `PRINT` field.
///
/// Microsoft Word uses this field to store printer-control instructions. This
/// type retains opaque instruction text, a cached result, and field state only.
/// It never interprets printer-control codes, opens a printer, sends output,
/// changes print settings, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PrintField<'a> {
    instruction: &'a str,
    printer_instructions: &'a str,
    cached_result: Option<&'a str>,
    status: FieldStatus,
    owner: FieldOwner,
    position: usize,
}

/// Inert metadata for a legacy RTF `EMBED` field.
///
/// This type retains opaque embedded-object instruction text, a cached result,
/// and field state only. It never loads, inspects, deserializes, activates,
/// renders, or executes an embedded object, accesses an external resource, or
/// refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct EmbedField<'a> {
    instruction: &'a str,
    object_instructions: &'a str,
    cached_result: Option<&'a str>,
    status: FieldStatus,
    owner: FieldOwner,
    position: usize,
}

/// Inert metadata for a legacy RTF `BARCODE` field.
///
/// This type retains opaque barcode-instruction text, a cached result, and
/// field state only. It never parses or validates barcode data or symbology,
/// generates or renders a barcode, accesses an external resource, or refreshes
/// a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct BarcodeField<'a> {
    instruction: &'a str,
    barcode_instructions: &'a str,
    cached_result: Option<&'a str>,
    status: FieldStatus,
    owner: FieldOwner,
    position: usize,
}

/// Inert metadata for a legacy RTF `BIDIOUTLINE` field.
///
/// This type retains opaque instruction text, a cached result, and field state
/// only. It never reads right-to-left language, paragraph outline, or layout
/// state; chooses a numbering system; calculates a result; or refreshes a
/// field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct BidiOutlineField<'a> {
    instruction: &'a str,
    opaque_instructions: &'a str,
    cached_result: Option<&'a str>,
    status: FieldStatus,
    owner: FieldOwner,
    position: usize,
}

/// Inert metadata for a legacy RTF `SHAPE` field.
///
/// Word uses this legacy field as a drawing-canvas anchor. This type retains
/// opaque instruction text, a cached result, and field state only. It never
/// locates, links, loads, positions, lays out, or renders a drawing or canvas,
/// or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ShapeField<'a> {
    instruction: &'a str,
    opaque_instructions: &'a str,
    cached_result: Option<&'a str>,
    status: FieldStatus,
    owner: FieldOwner,
    position: usize,
}

/// The stored kind of a legacy RTF form-code field.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LegacyFormFieldKind {
    /// A `FORMTEXT` text-box form field.
    Text,
    /// A `FORMCHECKBOX` checkbox form field.
    CheckBox,
    /// A `FORMDROPDOWN` drop-down-list form field.
    DropDown,
}

/// Inert metadata for a legacy RTF form-code field.
///
/// This code-level view retains only the stored kind, opaque instruction text,
/// cached result, and field state. It does not read or reconcile the separate
/// RTF `\formfield` destination. It never fills a form, changes a selection or
/// checkbox state, invokes entry or exit macros, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct LegacyFormField<'a> {
    instruction: &'a str,
    kind: LegacyFormFieldKind,
    opaque_instructions: &'a str,
    cached_result: Option<&'a str>,
    status: FieldStatus,
    owner: FieldOwner,
    position: usize,
}

/// The stored kind of a legacy RTF building-block field.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AutoTextFieldKind {
    /// A historical `GLOSSARY` field.
    Glossary,
    /// An `AUTOTEXT` field.
    AutoText,
}

/// Inert metadata for a legacy RTF `GLOSSARY` or `AUTOTEXT` field.
///
/// ECMA-376 Part 1 §17.16.5.5 defines `AUTOTEXT`; `GLOSSARY` is its
/// historical equivalent. This type retains only the stored category, entry
/// name, switches, cached result, and state. It never looks up a building
/// block, reads a template, inserts content, changes bookmarks, refreshes a
/// field, or accesses an external resource.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AutoTextField<'a> {
    instruction: &'a str,
    kind: AutoTextFieldKind,
    entry_name: Cow<'a, str>,
    unknown_switches: Vec<FieldSwitch<'a>>,
    cached_result: Option<&'a str>,
    status: FieldStatus,
    owner: FieldOwner,
    position: usize,
}

/// One recognized stored option of a legacy RTF `AUTOTEXTLIST` field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum AutoTextListOption<'a> {
    /// The `\\s` style name used to limit eligible building blocks.
    Style(Cow<'a, str>),
    /// The `\\t` stored tip text.
    Tip(Cow<'a, str>),
}

/// Typed, inert metadata for a legacy RTF `AUTOTEXTLIST` field.
///
/// ECMA-376 Part 1 §17.16.5.6 defines `AUTOTEXTLIST` using optional
/// display text and style/tip switches. This type retains only those stored
/// values, unknown switches, cached result, and state. It never shows a
/// selection UI, looks up eligible building blocks, reads a template, inserts
/// content, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AutoTextListField<'a> {
    instruction: &'a str,
    display_text: Option<Cow<'a, str>>,
    options: Vec<AutoTextListOption<'a>>,
    unknown_switches: Vec<FieldSwitch<'a>>,
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

/// One recognized stored option of an INDEX field.
///
/// These values describe how a producer configured an index. They are inert
/// metadata only: this crate never scans XE markers, reads bookmarks,
/// calculates page numbers, or generates an index.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum IndexOption<'a> {
    /// The \b bookmark that bounds included entries.
    Bookmark(Cow<'a, str>),
    /// The \c requested number of index columns.
    Columns(Cow<'a, str>),
    /// The \d separator between sequence and page numbers.
    SequencePageSeparator(Cow<'a, str>),
    /// The \e separator between an entry and its first page number.
    EntryPageNumberSeparator(Cow<'a, str>),
    /// The \f entry type that selects XE markers.
    EntryType(Cow<'a, str>),
    /// The \g separator between the start and end of a page range.
    PageRangeSeparator(Cow<'a, str>),
    /// The \h heading text for each index-letter set.
    Heading(Cow<'a, str>),
    /// The \k separator between an entry and its cross reference.
    CrossReferenceSeparator(Cow<'a, str>),
    /// The \l separator between page numbers in a page-number list.
    PageNumberSeparator(Cow<'a, str>),
    /// The \p range of entry initial letters to include.
    LetterRange(Cow<'a, str>),
    /// The \r switch runs subentries into their main-entry line.
    RunIn,
    /// The \s SEQ identifier whose number prefixes page numbers.
    SequenceIdentifier(Cow<'a, str>),
    /// The \y switch enables yomi text for index entries.
    UseYomi,
    /// The \z language identifier used to generate the index.
    LanguageId(Cow<'a, str>),
}

/// Inert metadata for a legacy RTF INDEX field.
///
/// This model retains stored configuration and a cached result only. It never
/// scans index-entry markers, follows bookmarks, calculates page numbers,
/// paginates the document, generates an index, or refreshes the field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct IndexField<'a> {
    instruction: &'a str,
    options: Vec<IndexOption<'a>>,
    unknown_switches: Vec<FieldSwitch<'a>>,
    cached_result: Option<&'a str>,
    status: FieldStatus,
    owner: FieldOwner,
    position: usize,
}

/// One recognized stored option of an XE index-entry field.
///
/// These values identify how an index marker participates in an INDEX field.
/// They are inert metadata only: this crate never changes hidden text,
/// calculates pages, follows bookmarks, or generates an index.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum IndexEntryOption<'a> {
    /// The \b switch toggles bold formatting for the entry's page number.
    BoldPageNumber,
    /// The \f entry type that selects this marker.
    EntryType(Cow<'a, str>),
    /// The \i switch toggles italic formatting for the entry's page number.
    ItalicPageNumber,
    /// The \r bookmark that marks a page range.
    PageRangeBookmark(Cow<'a, str>),
    /// The \t text that replaces a page number with a cross reference.
    CrossReference(Cow<'a, str>),
    /// The \y yomi sorting text.
    Yomi(Cow<'a, str>),
}

/// Inert metadata for a legacy RTF XE index-entry field.
///
/// This model retains a stored index marker and cached result only. It never
/// changes hidden text, resolves a bookmark, calculates pages, or generates
/// an index.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct IndexEntryField<'a> {
    instruction: &'a str,
    entry: Cow<'a, str>,
    options: Vec<IndexEntryOption<'a>>,
    unknown_switches: Vec<FieldSwitch<'a>>,
    cached_result: Option<&'a str>,
    status: FieldStatus,
    owner: FieldOwner,
    position: usize,
}

/// One recognized stored option of a CITATION field.
///
/// These values describe a citation's stored bibliography metadata. They are
/// inert only: this crate never resolves source tags, selects a bibliography
/// style, or formats a citation.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum CitationOption<'a> {
    /// The `\l` language identifier used when formatting the citation.
    LanguageId(Cow<'a, str>),
    /// The `\f` prefix prepended to the citation.
    Prefix(Cow<'a, str>),
    /// The `\s` suffix appended to the citation.
    Suffix(Cow<'a, str>),
    /// The `\p` cited page number.
    PageNumber(Cow<'a, str>),
    /// The `\v` cited volume number.
    VolumeNumber(Cow<'a, str>),
    /// The `\n` switch suppresses author information.
    SuppressAuthor,
    /// The `\t` switch suppresses title information.
    SuppressTitle,
    /// The `\y` switch suppresses year information.
    SuppressYear,
    /// A source tag added by the `\m` multi-source switch.
    AdditionalSourceTag(Cow<'a, str>),
}

/// Inert metadata for a legacy RTF CITATION field.
///
/// This model retains a stored source tag, options, and cached result only. It
/// never loads bibliography data, resolves source tags, applies a style, or
/// formats a citation.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CitationField<'a> {
    instruction: &'a str,
    source_tag: Cow<'a, str>,
    options: Vec<CitationOption<'a>>,
    unknown_switches: Vec<FieldSwitch<'a>>,
    cached_result: Option<&'a str>,
    status: FieldStatus,
    owner: FieldOwner,
    position: usize,
}

/// One recognized stored option of a BIBLIOGRAPHY field.
///
/// These values describe bibliography selection metadata only. This crate
/// never loads source records, filters them, applies a style, sorts entries,
/// or generates bibliography content.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum BibliographyOption<'a> {
    /// The `\l` language identifier used for sources without a locale.
    LanguageId(Cow<'a, str>),
    /// The `\f` language identifier used to filter source records.
    FilterLanguageId(Cow<'a, str>),
    /// The `\m` source tag used to select one source record.
    SourceTag(Cow<'a, str>),
}

/// Inert metadata for a legacy RTF BIBLIOGRAPHY field.
///
/// This model retains stored options and a cached result only. It never loads
/// source records, filters them, applies a style, sorts entries, or generates
/// bibliography content.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct BibliographyField<'a> {
    instruction: &'a str,
    options: Vec<BibliographyOption<'a>>,
    unknown_switches: Vec<FieldSwitch<'a>>,
    cached_result: Option<&'a str>,
    status: FieldStatus,
    owner: FieldOwner,
    position: usize,
}

/// Inert metadata for a legacy RTF DOCVARIABLE field.
///
/// This model retains a stored variable name, preserved switches, and cached
/// result only. It never reads a document-variable destination,
/// resolves a variable value, evaluates a field, or refreshes its result.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DocumentVariableField<'a> {
    instruction: &'a str,
    variable_name: Cow<'a, str>,
    unknown_switches: Vec<FieldSwitch<'a>>,
    cached_result: Option<&'a str>,
    status: FieldStatus,
    owner: FieldOwner,
    position: usize,
}

/// Inert metadata for a legacy RTF `DOCPROPERTY` field.
///
/// ECMA-376 Part 1 §17.16.5.14 defines one stored document-property name
/// followed by optional field switches. This model retains that name, switches,
/// and cached result only. It never reads core, extended, or custom document
/// properties, resolves a value, or refreshes the field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DocumentPropertyField<'a> {
    instruction: &'a str,
    property_name: Cow<'a, str>,
    switches: Vec<FieldSwitch<'a>>,
    cached_result: Option<&'a str>,
    status: FieldStatus,
    owner: FieldOwner,
    position: usize,
}

/// Inert metadata for an explicit legacy RTF `INFO` field.
///
/// Word permits the `INFO` keyword to be omitted, but that form overlaps
/// standalone document-information fields such as `TITLE`. This model
/// therefore recognizes the unambiguous explicit keyword only. It retains the
/// stored property selector, optional replacement value, switches, cached
/// result, and field state only. It never reads, resolves, modifies, or writes
/// document or template properties, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct InfoField<'a> {
    instruction: &'a str,
    information_type: Cow<'a, str>,
    new_value: Option<Cow<'a, str>>,
    switches: Vec<FieldSwitch<'a>>,
    cached_result: Option<&'a str>,
    status: FieldStatus,
    owner: FieldOwner,
    position: usize,
}

/// The built-in Word document-information field category.
///
/// ECMA-376 Part 1 §17.16.5 defines these fields. This enum preserves the
/// stored field kind only; it does not resolve document metadata or calculate
/// dates, revisions, or statistics.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum DocumentInformationFieldKind {
    Title,
    Subject,
    Author,
    Keywords,
    Comments,
    LastSavedBy,
    CreateDate,
    SaveDate,
    PrintDate,
    RevisionNumber,
    EditTime,
    NumberOfPages,
    NumberOfWords,
    NumberOfCharacters,
}

impl DocumentInformationFieldKind {
    /// The uppercase field keyword stored in a Word field instruction.
    pub const fn field_keyword(self) -> &'static str {
        match self {
            Self::Title => "TITLE",
            Self::Subject => "SUBJECT",
            Self::Author => "AUTHOR",
            Self::Keywords => "KEYWORDS",
            Self::Comments => "COMMENTS",
            Self::LastSavedBy => "LASTSAVEDBY",
            Self::CreateDate => "CREATEDATE",
            Self::SaveDate => "SAVEDATE",
            Self::PrintDate => "PRINTDATE",
            Self::RevisionNumber => "REVNUM",
            Self::EditTime => "EDITTIME",
            Self::NumberOfPages => "NUMPAGES",
            Self::NumberOfWords => "NUMWORDS",
            Self::NumberOfCharacters => "NUMCHARS",
        }
    }

    fn from_keyword(keyword: &str) -> Option<Self> {
        [
            Self::Title,
            Self::Subject,
            Self::Author,
            Self::Keywords,
            Self::Comments,
            Self::LastSavedBy,
            Self::CreateDate,
            Self::SaveDate,
            Self::PrintDate,
            Self::RevisionNumber,
            Self::EditTime,
            Self::NumberOfPages,
            Self::NumberOfWords,
            Self::NumberOfCharacters,
        ]
        .into_iter()
        .find(|kind| keyword.eq_ignore_ascii_case(kind.field_keyword()))
    }
}

/// Inert metadata for a legacy RTF built-in Word document-information field.
///
/// This type retains the stored kind, field switches, cached result, and field
/// state only. It never reads document properties, reads or modifies host
/// identity data, calculates dates, revisions, or statistics, resolves a
/// value, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DocumentInformationField<'a> {
    instruction: &'a str,
    kind: DocumentInformationFieldKind,
    switches: Vec<FieldSwitch<'a>>,
    cached_result: Option<&'a str>,
    status: FieldStatus,
    owner: FieldOwner,
    position: usize,
}

/// The built-in Word document-context and runtime field category.
///
/// `FILENAME`, `TEMPLATE`, `DATE`, `TIME`, `PAGE`, `FILESIZE`, `SECTION`, and
/// `SECTIONPAGES` are defined in ECMA-376 Part 1 §17.16.5. This enum preserves
/// the stored field kind only; it does not read a document path, attached
/// template, host filesystem state or file size, current clock, or page and
/// section layout.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum DocumentContextFieldKind {
    FileName,
    Template,
    Date,
    Time,
    Page,
    FileSize,
    Section,
    SectionPages,
}

impl DocumentContextFieldKind {
    /// The uppercase field keyword stored in a Word field instruction.
    pub const fn field_keyword(self) -> &'static str {
        match self {
            Self::FileName => "FILENAME",
            Self::Template => "TEMPLATE",
            Self::Date => "DATE",
            Self::Time => "TIME",
            Self::Page => "PAGE",
            Self::FileSize => "FILESIZE",
            Self::Section => "SECTION",
            Self::SectionPages => "SECTIONPAGES",
        }
    }

    fn from_keyword(keyword: &str) -> Option<Self> {
        [
            Self::FileName,
            Self::Template,
            Self::Date,
            Self::Time,
            Self::Page,
            Self::FileSize,
            Self::Section,
            Self::SectionPages,
        ]
            .into_iter()
            .find(|kind| keyword.eq_ignore_ascii_case(kind.field_keyword()))
    }
}

/// Inert metadata for a legacy RTF built-in Word document-context or runtime
/// field.
///
/// This type retains the stored kind, field switches, cached result, and field
/// state only. It never reads a document path, attached template, host
/// filesystem state or file size, current clock, or page and section layout,
/// resolves a value, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DocumentContextField<'a> {
    instruction: &'a str,
    kind: DocumentContextFieldKind,
    switches: Vec<FieldSwitch<'a>>,
    cached_result: Option<&'a str>,
    status: FieldStatus,
    owner: FieldOwner,
    position: usize,
}

/// Inert metadata for a legacy RTF `MERGEFIELD` field.
///
/// This model retains the stored merge-field name, switches, and cached result
/// only. It never opens a data source, resolves a record, performs a merge, or
/// refreshes a field result.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MergeField<'a> {
    instruction: &'a str,
    field_name: Cow<'a, str>,
    switches: Vec<FieldSwitch<'a>>,
    cached_result: Option<&'a str>,
    status: FieldStatus,
    owner: FieldOwner,
    position: usize,
}

/// Inert metadata for a legacy RTF `DATA` mail-merge source field.
///
/// [MS-DOC] §2.9.90 specifies `DATA datafile [headerfile]` as a field that
/// redirects mail-merge data and header files. This model retains only those
/// stored operands, switches, cached result, and field state. It never opens,
/// reads, connects to, resolves, or modifies either source; it never selects a
/// record, performs a merge, or refreshes a field result.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MailMergeDataField<'a> {
    instruction: &'a str,
    data_source: Cow<'a, str>,
    header_source: Option<Cow<'a, str>>,
    switches: Vec<FieldSwitch<'a>>,
    cached_result: Option<&'a str>,
    status: FieldStatus,
    owner: FieldOwner,
    position: usize,
}

/// The stored kind of a legacy RTF mail-merge counter field.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MailMergeCounterKind {
    /// A `MERGEREC` field, which stores a selected-record position.
    Record,
    /// A `MERGESEQ` field, which stores a merged-record sequence position.
    Sequence,
}

/// Inert metadata for a legacy RTF `MERGEREC` or `MERGESEQ` field.
///
/// ECMA-376 Part 1 §§17.16.5.36–37 define these zero-argument fields. This
/// model retains the stored kind and cached result only. It never selects or
/// counts records, opens a data source, performs a merge, or refreshes a
/// field result.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MailMergeCounterField<'a> {
    instruction: &'a str,
    kind: MailMergeCounterKind,
    cached_result: Option<&'a str>,
    status: FieldStatus,
    owner: FieldOwner,
    position: usize,
}

/// Inert metadata for a legacy RTF `NEXT` mail-merge control field.
///
/// ECMA-376 Part 1 §17.16.5.38 defines `NEXT` as a zero-argument
/// instruction. This model retains cached text and field state only. It never
/// advances a record, opens a data source, performs a merge, or refreshes a
/// field result.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MailMergeNextField<'a> {
    instruction: &'a str,
    cached_result: Option<&'a str>,
    status: FieldStatus,
    owner: FieldOwner,
    position: usize,
}

/// The stored kind of a conditional mail-merge control field.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MailMergeConditionalControlKind {
    /// A `NEXTIF` field, which can advance a merge record when its comparison is true.
    NextIf,
    /// A `SKIPIF` field, which can omit a merge record when its comparison is true.
    SkipIf,
}

/// Inert metadata for a legacy RTF `NEXTIF` or `SKIPIF` field.
///
/// ECMA-376 Part 1 §§17.16.5.39 and 17.16.5.58 define these controls. This
/// model retains the unparsed comparison and cached result only. It never
/// parses or evaluates a comparison, advances or skips a record, opens a data
/// source, performs a merge, or refreshes a field result.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MailMergeConditionalControlField<'a> {
    instruction: &'a str,
    kind: MailMergeConditionalControlKind,
    comparison: &'a str,
    cached_result: Option<&'a str>,
    status: FieldStatus,
    owner: FieldOwner,
    position: usize,
}

/// Inert metadata for a legacy RTF `IF` field.
///
/// ECMA-376 Part 1 §17.16.5.26 defines `IF` using a comparison and two
/// branches. This model retains the unparsed expression and cached result only.
/// It never parses or evaluates an expression, resolves field values, or
/// refreshes a field result.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct IfField<'a> {
    instruction: &'a str,
    expression: &'a str,
    cached_result: Option<&'a str>,
    status: FieldStatus,
    owner: FieldOwner,
    position: usize,
}

/// Inert metadata for a legacy RTF `COMPARE` field.
///
/// ECMA-376 Part 1 §17.16.5.10 defines `COMPARE` using a comparison whose
/// result is 1 or 0. This model retains the unparsed comparison and cached
/// result only. It never parses or evaluates a comparison, resolves nested
/// field values, or refreshes a field result.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CompareField<'a> {
    instruction: &'a str,
    comparison: &'a str,
    cached_result: Option<&'a str>,
    status: FieldStatus,
    owner: FieldOwner,
    position: usize,
}

/// Inert metadata for a legacy RTF `SET` field.
///
/// ECMA-376 Part 1 §17.16.5.57 defines `SET` using a target name and an
/// expression. This model retains the stored target, opaque expression, and
/// cached result only. It never evaluates an expression, looks up or changes a
/// bookmark, changes document state, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SetField<'a> {
    instruction: &'a str,
    target_name: Cow<'a, str>,
    expression: &'a str,
    cached_result: Option<&'a str>,
    status: FieldStatus,
    owner: FieldOwner,
    position: usize,
}

/// Inert metadata for a legacy RTF `SEQ` field.
///
/// ECMA-376 Part 1 §17.16.5.56 defines `SEQ` using an identifier, optional
/// bookmark, and optional switches. This model retains those stored values and a
/// cached result only. It never looks up a bookmark, increments or resets a
/// sequence, calculates a number, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SequenceField<'a> {
    instruction: &'a str,
    identifier: Cow<'a, str>,
    bookmark: Option<Cow<'a, str>>,
    tail: &'a str,
    cached_result: Option<&'a str>,
    status: FieldStatus,
    owner: FieldOwner,
    position: usize,
}

/// Inert metadata for a legacy RTF `=` formula field.
///
/// ECMA-376 Part 1 §17.16.3.3 defines table formulas using a leading `=`.
/// This model retains the stored formula text and cached result only. It never
/// parses or evaluates a formula, reads table cells or bookmarks, resolves
/// field values, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FormulaField<'a> {
    instruction: &'a str,
    formula: &'a str,
    cached_result: Option<&'a str>,
    status: FieldStatus,
    owner: FieldOwner,
    position: usize,
}

/// Inert metadata for a legacy RTF `QUOTE` field.
///
/// ECMA-376 Part 1 §17.16.5.49 defines `QUOTE` with a text field argument.
/// This model retains the stored text argument, switches, and cached result
/// only. It never interprets character codes, expands nested fields, inserts
/// text, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct QuoteField<'a> {
    instruction: &'a str,
    text: Cow<'a, str>,
    switches: Vec<FieldSwitch<'a>>,
    cached_result: Option<&'a str>,
    status: FieldStatus,
    owner: FieldOwner,
    position: usize,
}

/// Inert metadata for a legacy RTF `SYMBOL` field.
///
/// ECMA-376 Part 1 §17.16.5.61 defines `SYMBOL` with one stored character
/// argument and optional switches. This model retains that argument, switches,
/// and cached result only. It never converts a character code, looks up a font,
/// inserts a glyph, changes formatting or layout, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SymbolField<'a> {
    instruction: &'a str,
    character_argument: Cow<'a, str>,
    switches: Vec<FieldSwitch<'a>>,
    cached_result: Option<&'a str>,
    status: FieldStatus,
    owner: FieldOwner,
    position: usize,
}

/// The legacy Word automatic-numbering field category.
///
/// `AUTONUM`, `AUTONUMLGL`, and `AUTONUMOUT` are retained for
/// document compatibility. This enum preserves the stored field kind only; it
/// does not calculate a number, inspect paragraphs, headings, or styles, or
/// change layout.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum AutoNumberFieldKind {
    AutoNum,
    AutoNumLegal,
    AutoNumOutline,
}

impl AutoNumberFieldKind {
    /// The uppercase field keyword stored in a Word field instruction.
    pub const fn field_keyword(self) -> &'static str {
        match self {
            Self::AutoNum => "AUTONUM",
            Self::AutoNumLegal => "AUTONUMLGL",
            Self::AutoNumOutline => "AUTONUMOUT",
        }
    }

    fn from_keyword(keyword: &str) -> Option<Self> {
        [Self::AutoNum, Self::AutoNumLegal, Self::AutoNumOutline]
            .into_iter()
            .find(|kind| keyword.eq_ignore_ascii_case(kind.field_keyword()))
    }
}

/// Inert metadata for a legacy RTF automatic-numbering field.
///
/// This type retains the stored kind, switches, cached result, and field state
/// only. It never calculates paragraph numbers, reads heading or style state,
/// changes paragraphs or layout, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AutoNumberField<'a> {
    instruction: &'a str,
    kind: AutoNumberFieldKind,
    switches: Vec<FieldSwitch<'a>>,
    cached_result: Option<&'a str>,
    status: FieldStatus,
    owner: FieldOwner,
    position: usize,
}

/// Inert metadata for a legacy RTF `LISTNUM` field.
///
/// ECMA-376 Part 1 §17.16.5.33 defines `LISTNUM` with an optional list
/// name and switches. This type retains those stored values and cached result
/// only. It never looks up a list, determines a level or start value,
/// calculates a number, changes layout, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ListNumberField<'a> {
    instruction: &'a str,
    list_name: Option<Cow<'a, str>>,
    switches: Vec<FieldSwitch<'a>>,
    cached_result: Option<&'a str>,
    status: FieldStatus,
    owner: FieldOwner,
    position: usize,
}

/// One recognized stored option of a Word `STYLEREF` field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum StyleReferenceFieldOption {
    /// The `\\l` request for the nearest styled text following the field.
    FollowingText,
    /// The `\\n` request for the referenced paragraph number.
    ParagraphNumber,
    /// The `\\p` request for the referenced paragraph's relative position.
    RelativePosition,
    /// The `\\r` request for the referenced paragraph number in relative context.
    ParagraphNumberRelativeContext,
    /// The `\\t` request to suppress non-delimiter or non-numerical text.
    SuppressNonNumberText,
    /// The `\\w` request for the referenced paragraph number in full context.
    ParagraphNumberFullContext,
}

/// A typed, inert Word `STYLEREF` field.
///
/// ECMA-376 Part 1 §17.16.5.59 defines `STYLEREF` using a style name and
/// switches. This type retains those stored values and a cached result only. It
/// never looks up styled text, searches document stories, calculates paragraph
/// numbers or relative positions, resolves page layout, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct StyleReferenceField<'a> {
    instruction: &'a str,
    style_name: Cow<'a, str>,
    options: Vec<StyleReferenceFieldOption>,
    unknown_switches: Vec<FieldSwitch<'a>>,
    cached_result: Option<&'a str>,
    status: FieldStatus,
    owner: FieldOwner,
    position: usize,
}

/// The stored kind of a prompt field.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PromptFieldKind {
    /// An `ASK` field that associates a response with a bookmark.
    Ask,
    /// A `FILLIN` field whose cached result represents a response.
    FillIn,
}

/// Inert metadata for a legacy RTF `ASK` or `FILLIN` field.
///
/// ECMA-376 Part 1 §§17.16.5.3 and 17.16.5.19 define these fields. This model
/// exposes stored prompt and default-response metadata only. It never displays
/// a prompt, captures a response, creates or updates a bookmark, performs a
/// merge, or refreshes a field result.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PromptField<'a> {
    instruction: &'a str,
    kind: PromptFieldKind,
    bookmark: Option<Cow<'a, str>>,
    prompt: Option<Cow<'a, str>>,
    default_response: Option<Cow<'a, str>>,
    prompts_once_per_mail_merge: bool,
    cached_result: Option<&'a str>,
    status: FieldStatus,
    owner: FieldOwner,
    position: usize,
}

/// The stored kind of a user-identity field.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum UserIdentityFieldKind {
    /// A `USERADDRESS` field.
    Address,
    /// A `USERINITIALS` field.
    Initials,
    /// A `USERNAME` field.
    Name,
}

/// A general-formatting request stored by a user-identity field.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum UserIdentityFormatting {
    /// The `\\* Caps` formatting request.
    Caps,
    /// The `\\* FirstCap` formatting request.
    FirstCap,
    /// The `\\* Lower` formatting request.
    Lower,
    /// The `\\* Upper` formatting request.
    Upper,
}

/// Inert metadata for a legacy RTF `USERADDRESS`, `USERINITIALS`, or `USERNAME` field.
///
/// ECMA-376 Part 1 §§17.16.5.69–71 define these fields. This model exposes a
/// stored override, formatting request, and cached result only. It never reads
/// or modifies a host user's identity, applies formatting, or refreshes a
/// field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct UserIdentityField<'a> {
    instruction: &'a str,
    kind: UserIdentityFieldKind,
    override_value: Option<Cow<'a, str>>,
    formatting: Option<UserIdentityFormatting>,
    cached_result: Option<&'a str>,
    status: FieldStatus,
    owner: FieldOwner,
    position: usize,
}

/// One stored point-based `ADVANCE` placement operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AdvanceFieldOperation {
    /// The `\\d` switch moves subsequent text down.
    Down,
    /// The `\\l` switch moves subsequent text left.
    Left,
    /// The `\\r` switch moves subsequent text right.
    Right,
    /// The `\\u` switch moves subsequent text up.
    Up,
    /// The `\\x` switch specifies a horizontal position from the left edge
    /// of the column, frame, or text box.
    HorizontalPosition,
    /// The `\\y` switch specifies a vertical position relative to the page.
    VerticalPosition,
}

/// One stored `ADVANCE` point adjustment.
///
/// This is an instruction for a word processor's layout engine only. It is
/// never applied by this library.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct AdvanceFieldAdjustment {
    operation: AdvanceFieldOperation,
    points: i64,
}

/// Inert metadata for a legacy RTF `ADVANCE` field.
///
/// ECMA-376 Part 1 §17.16.5.2 defines this field and its six point-based
/// placement switches. This model exposes stored adjustments and cached content
/// only. It never moves text, changes layout, reflows content, or refreshes a
/// field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AdvanceField<'a> {
    instruction: &'a str,
    adjustments: Vec<AdvanceFieldAdjustment>,
    cached_result: Option<&'a str>,
    status: FieldStatus,
    owner: FieldOwner,
    position: usize,
}

/// The stored kind of a mail-merge recipient layout field.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MailMergeRecipientFieldKind {
    /// An `ADDRESSBLOCK` field.
    AddressBlock,
    /// A `GREETINGLINE` field.
    GreetingLine,
}

/// How an `ADDRESSBLOCK` field requests country/region text.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AddressBlockCountryInclusion {
    /// `\\c 0` omits the country/region.
    Omit,
    /// `\\c 1` includes the country/region regardless of exclusions.
    Always,
    /// `\\c 2` includes the country/region unless it matches an excluded
    /// country/region.
    UnlessExcluded,
}

/// Inert metadata for a legacy RTF `ADDRESSBLOCK` or `GREETINGLINE` field.
///
/// ECMA-376 Part 1 §§17.16.5.1 and 17.16.5.24 define these mail-merge
/// recipient layout fields. This type exposes stored layout metadata and a
/// cached result only. It never opens a data source, selects a record,
/// performs a merge, expands placeholders, generates text, or refreshes a
/// field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MailMergeRecipientField<'a> {
    instruction: &'a str,
    kind: MailMergeRecipientFieldKind,
    country_inclusion: Option<AddressBlockCountryInclusion>,
    formats_using_recipient_country: bool,
    excluded_countries: Vec<Cow<'a, str>>,
    format_template: Option<Cow<'a, str>>,
    language: Option<Cow<'a, str>>,
    greeting_fallback_text: Option<Cow<'a, str>>,
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

struct IndexParts<'a> {
    options: Vec<IndexOption<'a>>,
    unknown_switches: Vec<FieldSwitch<'a>>,
}

struct IndexEntryParts<'a> {
    entry: Cow<'a, str>,
    options: Vec<IndexEntryOption<'a>>,
    unknown_switches: Vec<FieldSwitch<'a>>,
}

struct PromptFieldParts<'a> {
    kind: PromptFieldKind,
    bookmark: Option<Cow<'a, str>>,
    prompt: Option<Cow<'a, str>>,
    default_response: Option<Cow<'a, str>>,
    prompts_once_per_mail_merge: bool,
}

struct UserIdentityFieldParts<'a> {
    kind: UserIdentityFieldKind,
    override_value: Option<Cow<'a, str>>,
    formatting: Option<UserIdentityFormatting>,
}

struct MailMergeRecipientFieldParts<'a> {
    kind: MailMergeRecipientFieldKind,
    country_inclusion: Option<AddressBlockCountryInclusion>,
    formats_using_recipient_country: bool,
    excluded_countries: Vec<Cow<'a, str>>,
    format_template: Option<Cow<'a, str>>,
    language: Option<Cow<'a, str>>,
    greeting_fallback_text: Option<Cow<'a, str>>,
    unknown_switches: Vec<FieldSwitch<'a>>,
}

struct CitationParts<'a> {
    source_tag: Cow<'a, str>,
    options: Vec<CitationOption<'a>>,
    unknown_switches: Vec<FieldSwitch<'a>>,
}

struct BibliographyParts<'a> {
    options: Vec<BibliographyOption<'a>>,
    unknown_switches: Vec<FieldSwitch<'a>>,
}

struct DocumentVariableFieldParts<'a> {
    variable_name: Cow<'a, str>,
    unknown_switches: Vec<FieldSwitch<'a>>,
}

struct DocumentPropertyFieldParts<'a> {
    property_name: Cow<'a, str>,
    switches: Vec<FieldSwitch<'a>>,
}

struct InfoFieldParts<'a> {
    information_type: Cow<'a, str>,
    new_value: Option<Cow<'a, str>>,
    switches: Vec<FieldSwitch<'a>>,
}

struct DocumentInformationFieldParts<'a> {
    kind: DocumentInformationFieldKind,
    switches: Vec<FieldSwitch<'a>>,
}

struct DocumentContextFieldParts<'a> {
    kind: DocumentContextFieldKind,
    switches: Vec<FieldSwitch<'a>>,
}

struct AutoNumberFieldParts<'a> {
    kind: AutoNumberFieldKind,
    switches: Vec<FieldSwitch<'a>>,
}

struct ListNumberFieldParts<'a> {
    list_name: Option<Cow<'a, str>>,
    switches: Vec<FieldSwitch<'a>>,
}

struct MergeFieldParts<'a> {
    field_name: Cow<'a, str>,
    switches: Vec<FieldSwitch<'a>>,
}

struct MailMergeDataFieldParts<'a> {
    data_source: Cow<'a, str>,
    header_source: Option<Cow<'a, str>>,
    switches: Vec<FieldSwitch<'a>>,
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

impl<'a> GoToButtonField<'a> {
    /// Return the complete stored `GOTOBUTTON` field instruction.
    pub fn instruction(&self) -> &'a str {
        self.instruction
    }

    /// Return the stored destination without resolving or navigating to it.
    ///
    /// A destination can be a bookmark, page reference, annotation, footnote,
    /// line, page, or section expression.
    pub fn target(&self) -> &str {
        &self.target
    }

    /// Return the stored text or graphic-label expression for the button.
    ///
    /// This is source metadata, not an activated control.
    pub fn button_text(&self) -> &str {
        &self.button_text
    }

    /// Return the stored field result when a producer supplied one.
    ///
    /// This is cached text only and is never regenerated from the destination.
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

impl<'a> ActiveContentField<'a> {
    /// Return the complete stored active-content field instruction.
    ///
    /// This string remains opaque metadata and is never interpreted.
    pub fn instruction(&self) -> &'a str {
        self.instruction
    }

    /// Return whether this stores add-in, OCX-control, or HTML-control metadata.
    pub const fn kind(&self) -> ActiveContentFieldKind {
        self.kind
    }

    /// Return the stored field result when a producer supplied one.
    ///
    /// This is cached text only and is never regenerated by loading or running
    /// content.
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

impl<'a> PrintField<'a> {
    /// Return the complete stored `PRINT` field instruction.
    ///
    /// This string remains opaque metadata and is never interpreted or sent to
    /// a printer.
    pub fn instruction(&self) -> &'a str {
        self.instruction
    }

    /// Return the stored printer-instruction text after the `PRINT` keyword.
    ///
    /// This can include printer-control or PostScript text. It is never parsed,
    /// interpreted, or sent to a printer.
    pub fn printer_instructions(&self) -> &'a str {
        self.printer_instructions
    }

    /// Return the stored field result when a producer supplied one.
    ///
    /// This is cached text only and is never regenerated by printing.
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

impl<'a> EmbedField<'a> {
    /// Return the complete stored `EMBED` field instruction.
    ///
    /// This string remains opaque metadata and is never used to load or
    /// activate an object.
    pub fn instruction(&self) -> &'a str {
        self.instruction
    }

    /// Return the stored opaque object-instruction text after `EMBED`.
    ///
    /// It is never parsed, used to locate an object, or used to load, inspect,
    /// deserialize, activate, render, or execute object content.
    pub fn object_instructions(&self) -> &'a str {
        self.object_instructions
    }

    /// Return the stored field result when a producer supplied one.
    ///
    /// This is cached text only and is never regenerated from an object.
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

impl<'a> BarcodeField<'a> {
    /// Return the complete stored `BARCODE` field instruction.
    ///
    /// This string remains opaque metadata and is never used to generate or
    /// render a barcode.
    pub fn instruction(&self) -> &'a str {
        self.instruction
    }

    /// Return the stored opaque barcode-instruction text after `BARCODE`.
    ///
    /// It is never parsed, validated, interpreted, or used to generate or
    /// render barcode content.
    pub fn barcode_instructions(&self) -> &'a str {
        self.barcode_instructions
    }

    /// Return the stored field result when a producer supplied one.
    ///
    /// This is cached text only and is never regenerated from barcode data.
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

impl<'a> BidiOutlineField<'a> {
    /// Return the complete stored `BIDIOUTLINE` field instruction.
    ///
    /// This string remains opaque metadata and is never used to calculate an
    /// outline number.
    pub fn instruction(&self) -> &'a str {
        self.instruction
    }

    /// Return opaque stored instruction text after `BIDIOUTLINE`.
    ///
    /// It is never parsed, interpreted, or used to resolve language, outline,
    /// numbering, or layout state.
    pub fn opaque_instructions(&self) -> &'a str {
        self.opaque_instructions
    }

    /// Return the stored field result when a producer supplied one.
    ///
    /// This is cached text only and is never regenerated from document state.
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

impl<'a> ShapeField<'a> {
    /// Return the complete stored `SHAPE` field instruction.
    ///
    /// This string remains opaque metadata and is never used to locate or
    /// position a drawing canvas.
    pub fn instruction(&self) -> &'a str {
        self.instruction
    }

    /// Return opaque stored instruction text after `SHAPE`.
    ///
    /// It is never parsed, interpreted, or used to link a field to a drawing,
    /// resolve an anchor, or calculate layout.
    pub fn opaque_instructions(&self) -> &'a str {
        self.opaque_instructions
    }

    /// Return the stored field result when a producer supplied one.
    ///
    /// This is cached metadata only and is never regenerated from a drawing
    /// canvas.
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

impl<'a> AutoTextField<'a> {
    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &'a str {
        self.instruction
    }

    /// Return whether this stores `GLOSSARY` or `AUTOTEXT` metadata.
    pub const fn kind(&self) -> AutoTextFieldKind {
        self.kind
    }

    /// Return the stored building-block entry name without resolving it.
    pub fn entry_name(&self) -> &str {
        &self.entry_name
    }

    /// Return unrecognized stored switches in source order.
    ///
    /// They are retained without interpretation or execution.
    pub fn unknown_switches(&self) -> &[FieldSwitch<'a>] {
        &self.unknown_switches
    }

    /// Return the stored field result when a producer supplied one.
    ///
    /// This is cached text only and is never regenerated by looking up or
    /// inserting content.
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

impl<'a> AutoTextListField<'a> {
    /// Return the complete stored `AUTOTEXTLIST` field instruction.
    pub fn instruction(&self) -> &'a str {
        self.instruction
    }

    /// Return the optional stored display text.
    ///
    /// This text is metadata only and never triggers a selection UI.
    pub fn display_text(&self) -> Option<&str> {
        self.display_text.as_deref()
    }

    /// Return recognized stored options in source order.
    ///
    /// This metadata is never used to query, select, or insert a building
    /// block.
    pub fn options(&self) -> &[AutoTextListOption<'a>] {
        &self.options
    }

    /// Return unrecognized stored switches in source order.
    ///
    /// They are retained without interpretation or execution.
    pub fn unknown_switches(&self) -> &[FieldSwitch<'a>] {
        &self.unknown_switches
    }

    /// Return the stored field result when a producer supplied one.
    ///
    /// This is cached text only and is never regenerated by selection or
    /// content insertion.
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

impl<'a> IndexField<'a> {
    /// Return the complete stored INDEX field instruction.
    pub fn instruction(&self) -> &'a str {
        self.instruction
    }

    /// Return recognized INDEX options in stored source order.
    ///
    /// These options are configuration metadata only. This method never scans
    /// XE markers, reads bookmarks, calculates page numbers, or generates an
    /// index.
    pub fn options(&self) -> &[IndexOption<'a>] {
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

impl<'a> IndexEntryField<'a> {
    /// Return the complete stored XE field instruction.
    pub fn instruction(&self) -> &'a str {
        self.instruction
    }

    /// Return the stored index-entry text without generating an index.
    pub fn entry(&self) -> &str {
        &self.entry
    }

    /// Return recognized XE options in stored source order.
    ///
    /// These options are marker metadata only. This method never changes
    /// hidden text, follows bookmarks, calculates pages, or generates an
    /// index.
    pub fn options(&self) -> &[IndexEntryOption<'a>] {
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

impl<'a> CitationField<'a> {
    /// Return the complete stored CITATION field instruction.
    pub fn instruction(&self) -> &'a str {
        self.instruction
    }

    /// Return the primary bibliography source tag without resolving it.
    pub fn source_tag(&self) -> &str {
        &self.source_tag
    }

    /// Return recognized CITATION options in stored source order.
    ///
    /// These options are metadata only. This method never loads sources,
    /// applies a bibliography style, or formats a citation.
    pub fn options(&self) -> &[CitationOption<'a>] {
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

impl<'a> BibliographyField<'a> {
    /// Return the complete stored BIBLIOGRAPHY field instruction.
    pub fn instruction(&self) -> &'a str {
        self.instruction
    }

    /// Return recognized BIBLIOGRAPHY options in stored source order.
    ///
    /// These options are metadata only. This method never loads sources,
    /// filters records, applies a style, sorts entries, or generates content.
    pub fn options(&self) -> &[BibliographyOption<'a>] {
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

impl<'a> DocumentVariableField<'a> {
    /// Return the complete stored DOCVARIABLE field instruction.
    pub fn instruction(&self) -> &'a str {
        self.instruction
    }

    /// Return the stored document-variable name without resolving it.
    pub fn variable_name(&self) -> &str {
        &self.variable_name
    }

    /// Return unrecognized stored field switches in source order.
    ///
    /// DOCVARIABLE has no field-specific switches. These values are retained
    /// as inert source metadata and are never interpreted.
    pub fn unknown_switches(&self) -> &[FieldSwitch<'a>] {
        &self.unknown_switches
    }

    /// Return the stored field result when a producer supplied one.
    ///
    /// This is cached text only and is never regenerated from a variable.
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

impl<'a> DocumentPropertyField<'a> {
    /// Return the complete stored `DOCPROPERTY` field instruction.
    pub fn instruction(&self) -> &'a str {
        self.instruction
    }

    /// Return the stored document-property name without resolving it.
    pub fn property_name(&self) -> &str {
        &self.property_name
    }

    /// Return stored field switches in source order without interpreting them.
    pub fn switches(&self) -> &[FieldSwitch<'a>] {
        &self.switches
    }

    /// Return the stored field result when a producer supplied one.
    ///
    /// This is cached text only and is never regenerated from a property.
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

impl<'a> InfoField<'a> {
    /// Return the complete stored `INFO` field instruction.
    pub fn instruction(&self) -> &'a str {
        self.instruction
    }

    /// Return the stored document or template property selector.
    ///
    /// The selector is preserved as metadata and is never looked up.
    pub fn information_type(&self) -> &str {
        &self.information_type
    }

    /// Return the stored optional replacement value.
    ///
    /// This value is never applied to a document or template property.
    pub fn new_value(&self) -> Option<&str> {
        self.new_value.as_deref()
    }

    /// Return stored field switches in source order without interpreting them.
    pub fn switches(&self) -> &[FieldSwitch<'a>] {
        &self.switches
    }

    /// Return the stored field result when a producer supplied one.
    ///
    /// This is cached text only and is never regenerated from a property.
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

impl<'a> DocumentInformationField<'a> {
    /// Return the complete stored document-information field instruction.
    pub fn instruction(&self) -> &'a str {
        self.instruction
    }

    /// Return the recognized built-in document-information category.
    pub const fn kind(&self) -> DocumentInformationFieldKind {
        self.kind
    }

    /// Return stored field switches in source order without interpreting them.
    pub fn switches(&self) -> &[FieldSwitch<'a>] {
        &self.switches
    }

    /// Return the stored field result when a producer supplied one.
    ///
    /// This is cached text only and is never regenerated from document
    /// metadata or a host user profile.
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

impl<'a> DocumentContextField<'a> {
    /// Return the complete stored document-context or runtime field instruction.
    pub fn instruction(&self) -> &'a str {
        self.instruction
    }

    /// Return the recognized built-in document-context or runtime category.
    pub const fn kind(&self) -> DocumentContextFieldKind {
        self.kind
    }

    /// Return stored field switches in source order without interpreting them.
    pub fn switches(&self) -> &[FieldSwitch<'a>] {
        &self.switches
    }

    /// Return the stored field result when a producer supplied one.
    ///
    /// This is cached text only and is never regenerated from a document path,
    /// attached template, host filesystem state or file size, current clock,
    /// or page and section layout.
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

impl<'a> MergeField<'a> {
    /// Return the complete stored `MERGEFIELD` field instruction.
    pub fn instruction(&self) -> &'a str {
        self.instruction
    }

    /// Return the stored data-column name without resolving a data source.
    pub fn field_name(&self) -> &str {
        &self.field_name
    }

    /// Return stored field switches in source order without interpreting them.
    pub fn switches(&self) -> &[FieldSwitch<'a>] {
        &self.switches
    }

    /// Return the stored field result when a producer supplied one.
    ///
    /// This is cached text only and is never regenerated by a merge.
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

impl<'a> MailMergeDataField<'a> {
    /// Return the complete stored `DATA` field instruction.
    pub fn instruction(&self) -> &'a str {
        self.instruction
    }

    /// Return the stored mail-merge data-source identifier without opening it.
    pub fn data_source(&self) -> &str {
        &self.data_source
    }

    /// Return the optional stored mail-merge header-source identifier.
    ///
    /// This value is never opened or resolved.
    pub fn header_source(&self) -> Option<&str> {
        self.header_source.as_deref()
    }

    /// Return stored field switches in source order without interpreting them.
    pub fn switches(&self) -> &[FieldSwitch<'a>] {
        &self.switches
    }

    /// Return the stored field result when a producer supplied one.
    ///
    /// This is cached text only and is never regenerated by a merge.
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

impl<'a> MailMergeCounterField<'a> {
    /// Return the complete stored `MERGEREC` or `MERGESEQ` field instruction.
    pub fn instruction(&self) -> &'a str {
        self.instruction
    }

    /// Return whether this is a `MERGEREC` or `MERGESEQ` field.
    pub const fn kind(&self) -> MailMergeCounterKind {
        self.kind
    }

    /// Return the stored field result when a producer supplied one.
    ///
    /// This is cached text only and is never regenerated by a merge.
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

impl<'a> MailMergeNextField<'a> {
    /// Return the complete stored `NEXT` field instruction.
    pub fn instruction(&self) -> &'a str {
        self.instruction
    }

    /// Return the stored field result when a producer supplied one.
    ///
    /// This is cached text only and is never regenerated by a merge.
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

impl<'a> MailMergeConditionalControlField<'a> {
    /// Return the complete stored `NEXTIF` or `SKIPIF` field instruction.
    pub fn instruction(&self) -> &'a str {
        self.instruction
    }

    /// Return whether this is a `NEXTIF` or `SKIPIF` control.
    pub const fn kind(&self) -> MailMergeConditionalControlKind {
        self.kind
    }

    /// Return the stored comparison without parsing or evaluating it.
    pub fn comparison(&self) -> &'a str {
        self.comparison
    }

    /// Return the stored field result when a producer supplied one.
    ///
    /// This is cached text only and is never regenerated by a merge.
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

impl<'a> IfField<'a> {
    /// Return the complete stored `IF` field instruction.
    pub fn instruction(&self) -> &'a str {
        self.instruction
    }

    /// Return the stored expression without parsing or evaluating it.
    pub fn expression(&self) -> &'a str {
        self.expression
    }

    /// Return the stored field result when a producer supplied one.
    ///
    /// This is cached text only and is never regenerated by field evaluation.
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

impl<'a> CompareField<'a> {
    /// Return the complete stored `COMPARE` field instruction.
    pub fn instruction(&self) -> &'a str {
        self.instruction
    }

    /// Return the stored comparison without parsing or evaluating it.
    pub fn comparison(&self) -> &'a str {
        self.comparison
    }

    /// Return the stored field result when a producer supplied one.
    ///
    /// This is cached text only and is never regenerated by field evaluation.
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

impl<'a> SetField<'a> {
    /// Return the complete stored `SET` field instruction.
    pub fn instruction(&self) -> &'a str {
        self.instruction
    }

    /// Return the stored target name without looking up or changing it.
    pub fn target_name(&self) -> &str {
        &self.target_name
    }

    /// Return the stored expression without parsing or evaluating it.
    pub fn expression(&self) -> &'a str {
        self.expression
    }

    /// Return the stored field result when a producer supplied one.
    ///
    /// This is cached text only and is never regenerated by field evaluation.
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

impl<'a> SequenceField<'a> {
    /// Return the complete stored `SEQ` field instruction.
    pub fn instruction(&self) -> &'a str {
        self.instruction
    }

    /// Return the stored sequence identifier without calculating its value.
    pub fn identifier(&self) -> &str {
        &self.identifier
    }

    /// Return the optional stored bookmark name without looking it up.
    pub fn bookmark(&self) -> Option<&str> {
        self.bookmark.as_deref()
    }

    /// Return opaque stored text after the identifier and optional bookmark.
    ///
    /// This text is never parsed to change or calculate a sequence.
    pub fn tail(&self) -> &'a str {
        self.tail
    }

    /// Return the stored field result when a producer supplied one.
    ///
    /// This is cached text only and is never regenerated by calculating a
    /// sequence number.
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

impl<'a> FormulaField<'a> {
    /// Return the complete stored formula field instruction.
    pub fn instruction(&self) -> &'a str {
        self.instruction
    }

    /// Return the opaque stored formula text without parsing or evaluating it.
    pub fn formula(&self) -> &'a str {
        self.formula
    }

    /// Return the stored field result when a producer supplied one.
    ///
    /// This is cached text only and is never regenerated by formula evaluation.
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

impl<'a> QuoteField<'a> {
    /// Return the complete stored `QUOTE` field instruction.
    pub fn instruction(&self) -> &'a str {
        self.instruction
    }

    /// Return the stored text argument without inserting or transforming it.
    pub fn text(&self) -> &str {
        &self.text
    }

    /// Return stored field switches in source order without interpreting them.
    pub fn switches(&self) -> &[FieldSwitch<'a>] {
        &self.switches
    }

    /// Return the stored field result when a producer supplied one.
    ///
    /// This is cached text only and is never regenerated by inserting text.
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

impl<'a> SymbolField<'a> {
    /// Return the complete stored `SYMBOL` field instruction.
    pub fn instruction(&self) -> &'a str {
        self.instruction
    }

    /// Return the stored character argument without converting it to a glyph.
    pub fn character_argument(&self) -> &str {
        &self.character_argument
    }

    /// Return stored field switches in source order without interpreting them.
    pub fn switches(&self) -> &[FieldSwitch<'a>] {
        &self.switches
    }

    /// Return the stored field result when a producer supplied one.
    ///
    /// This is cached text only and is never regenerated by mapping a character
    /// code or inserting a glyph.
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

impl<'a> AutoNumberField<'a> {
    /// Return the complete stored automatic-numbering field instruction.
    pub fn instruction(&self) -> &'a str {
        self.instruction
    }

    /// Return the recognized automatic-numbering category.
    pub const fn kind(&self) -> AutoNumberFieldKind {
        self.kind
    }

    /// Return stored field switches in source order without interpreting them.
    pub fn switches(&self) -> &[FieldSwitch<'a>] {
        &self.switches
    }

    /// Return the stored field result when a producer supplied one.
    ///
    /// This is cached text only and is never regenerated by calculating a
    /// paragraph number.
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

impl<'a> ListNumberField<'a> {
    /// Return the complete stored `LISTNUM` field instruction.
    pub fn instruction(&self) -> &'a str {
        self.instruction
    }

    /// Return the stored optional list name without looking it up.
    pub fn list_name(&self) -> Option<&str> {
        self.list_name.as_deref()
    }

    /// Return stored field switches in source order without interpreting them.
    pub fn switches(&self) -> &[FieldSwitch<'a>] {
        &self.switches
    }

    /// Return the stored field result when a producer supplied one.
    ///
    /// This is cached text only and is never regenerated by calculating a list
    /// number.
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

impl<'a> StyleReferenceField<'a> {
    /// Return the complete stored `STYLEREF` field instruction.
    pub fn instruction(&self) -> &'a str {
        self.instruction
    }

    /// Return the stored style name without looking up text that uses it.
    pub fn style_name(&self) -> &str {
        &self.style_name
    }

    /// Return recognized stored options in source order.
    ///
    /// This metadata is never used to search text, calculate a number, or
    /// resolve layout.
    pub fn options(&self) -> &[StyleReferenceFieldOption] {
        &self.options
    }

    /// Return unrecognized stored switches in source order.
    ///
    /// They are retained without interpretation or execution.
    pub fn unknown_switches(&self) -> &[FieldSwitch<'a>] {
        &self.unknown_switches
    }

    /// Return the stored field result when a producer supplied one.
    ///
    /// This is cached text only and is never regenerated by searching styled
    /// text or resolving layout.
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

impl<'a> PromptField<'a> {
    /// Return the complete stored `ASK` or `FILLIN` field instruction.
    pub fn instruction(&self) -> &'a str {
        self.instruction
    }

    /// Return whether this is an `ASK` or `FILLIN` field.
    pub const fn kind(&self) -> PromptFieldKind {
        self.kind
    }

    /// Return the bookmark name stored by an `ASK` field, if any.
    ///
    /// This is stored metadata only. It is never resolved, created, or updated.
    pub fn bookmark(&self) -> Option<&str> {
        self.bookmark.as_deref()
    }

    /// Return the stored prompt text, if any.
    ///
    /// This method returns metadata only and never displays a prompt.
    pub fn prompt(&self) -> Option<&str> {
        self.prompt.as_deref()
    }

    /// Return the stored default response, if one was supplied.
    ///
    /// `Some("")` represents an explicitly supplied blank default response. This
    /// is metadata only and is never selected, captured, or written into the
    /// document.
    pub fn default_response(&self) -> Option<&str> {
        self.default_response.as_deref()
    }

    /// Whether the stored `\o` switch requests one prompt for a mail merge.
    ///
    /// This request is never acted on: no merge is performed and no data source
    /// is opened.
    pub const fn prompts_once_per_mail_merge(&self) -> bool {
        self.prompts_once_per_mail_merge
    }

    /// Return the stored field result when a producer supplied one.
    ///
    /// This is cached text only and is never regenerated by field evaluation.
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

impl<'a> LegacyFormField<'a> {
    /// Return the complete stored field instruction.
    ///
    /// This string remains opaque metadata and is never used to change a form
    /// field.
    pub fn instruction(&self) -> &'a str {
        self.instruction
    }

    /// Return whether this is a text, checkbox, or drop-down form-code field.
    pub const fn kind(&self) -> LegacyFormFieldKind {
        self.kind
    }

    /// Return opaque stored instruction text after the form-code keyword.
    ///
    /// It is never parsed, interpreted, or used to fill a form, change a
    /// checkbox or selection, or invoke a macro.
    pub fn opaque_instructions(&self) -> &'a str {
        self.opaque_instructions
    }

    /// Return the stored field result when a producer supplied one.
    ///
    /// This is cached metadata only and is never regenerated from form state.
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

impl<'a> UserIdentityField<'a> {
    /// Return the complete stored user-identity field instruction.
    pub fn instruction(&self) -> &'a str {
        self.instruction
    }

    /// Return whether this is an address, initials, or name field.
    pub const fn kind(&self) -> UserIdentityFieldKind {
        self.kind
    }

    /// Return the optional stored value that overrides the host user context.
    ///
    /// `Some("")` represents an explicitly supplied blank override. This
    /// stored text is never written to, read from, or compared with a host
    /// identity.
    pub fn override_value(&self) -> Option<&str> {
        self.override_value.as_deref()
    }

    /// Return the stored general-formatting request, if any.
    ///
    /// This request is metadata only and is never applied to an identity value.
    pub const fn formatting(&self) -> Option<UserIdentityFormatting> {
        self.formatting
    }

    /// Return the stored field result when a producer supplied one.
    ///
    /// This is cached text only and is never regenerated from a host identity.
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

impl AdvanceFieldAdjustment {
    /// Return the requested placement operation.
    pub const fn operation(&self) -> AdvanceFieldOperation {
        self.operation
    }

    /// Return the stored signed integral number of points.
    pub const fn points(&self) -> i64 {
        self.points
    }
}

impl<'a> AdvanceField<'a> {
    /// Return the complete stored `ADVANCE` field instruction.
    pub fn instruction(&self) -> &'a str {
        self.instruction
    }

    /// Return the stored placement adjustments in source order.
    ///
    /// Repeated operations are preserved; this library does not resolve or
    /// apply them.
    pub fn adjustments(&self) -> &[AdvanceFieldAdjustment] {
        &self.adjustments
    }

    /// Return the stored field result when a producer supplied one.
    ///
    /// `ADVANCE` has no regenerated value here; any returned text is stored
    /// source content only.
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

impl<'a> MailMergeRecipientField<'a> {
    /// Return the complete stored `ADDRESSBLOCK` or `GREETINGLINE` instruction.
    pub fn instruction(&self) -> &'a str {
        self.instruction
    }

    /// Return whether this is an `ADDRESSBLOCK` or `GREETINGLINE` field.
    pub const fn kind(&self) -> MailMergeRecipientFieldKind {
        self.kind
    }

    /// Return how an `ADDRESSBLOCK` requests country/region text.
    ///
    /// This is `None` when the instruction has no `\\c` switch or when the
    /// field is a `GREETINGLINE`. The stored request is never used to render
    /// an address.
    pub const fn country_inclusion(&self) -> Option<AddressBlockCountryInclusion> {
        self.country_inclusion
    }

    /// Whether an `ADDRESSBLOCK` stores the `\\d` request to use the
    /// recipient country's address format.
    ///
    /// This request is metadata only and never causes a record or country
    /// format to be resolved.
    pub const fn formats_using_recipient_country(&self) -> bool {
        self.formats_using_recipient_country
    }

    /// Return country/region names excluded by `ADDRESSBLOCK` `\\e` switches.
    ///
    /// ECMA-376 permits repeated `\\e` switches; values are retained in source
    /// order. They are never matched against a recipient record.
    pub fn excluded_countries(&self) -> &[Cow<'a, str>] {
        &self.excluded_countries
    }

    /// Return the stored `\\f` layout template, if any.
    ///
    /// For `ADDRESSBLOCK`, this is the standard address/name placeholder
    /// template. For `GREETINGLINE`, this accepts Word's documented
    /// compatibility form. Placeholder text remains opaque metadata and is
    /// never expanded.
    pub fn format_template(&self) -> Option<&str> {
        self.format_template.as_deref()
    }

    /// Return the stored `\\l` language identifier, if any.
    ///
    /// The identifier is not used to choose locale-specific formatting.
    pub fn language(&self) -> Option<&str> {
        self.language.as_deref()
    }

    /// Return the stored `GREETINGLINE` fallback text, if any.
    ///
    /// ECMA-376 names `\\c` as this switch; Word-compatible fields can use
    /// `\\e`. Both forms are accepted as stored metadata, but neither is ever
    /// selected or displayed by this API.
    pub fn greeting_fallback_text(&self) -> Option<&str> {
        self.greeting_fallback_text.as_deref()
    }

    /// Return switches not specific to the recognized recipient-field kind.
    ///
    /// This includes formatting or producer-specific switches, retained in
    /// source order as inert metadata.
    pub fn unknown_switches(&self) -> &[FieldSwitch<'a>] {
        &self.unknown_switches
    }

    /// Return the stored field result when a producer supplied one.
    ///
    /// This is cached text only and is never regenerated by a merge.
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
            ParsedFieldCode::Other { .. } if is_formula_field_instruction(instruction) => {
                FieldType::Formula
            },
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
                if keyword.eq_ignore_ascii_case("ADVANCE") =>
            {
                FieldType::Advance
            },
            ParsedFieldCode::Other { ref keyword, .. }
                if keyword.eq_ignore_ascii_case("MACROBUTTON") =>
            {
                FieldType::MacroButton
            },
            ParsedFieldCode::Other { ref keyword, .. }
                if keyword.eq_ignore_ascii_case("GOTOBUTTON") =>
            {
                FieldType::GoToButton
            },
            ParsedFieldCode::Other { ref keyword, .. } if keyword.eq_ignore_ascii_case("PRINT") => {
                FieldType::Print
            },
            ParsedFieldCode::Other { ref keyword, .. } if keyword.eq_ignore_ascii_case("EMBED") => {
                FieldType::Embed
            },
            ParsedFieldCode::Other { ref keyword, .. }
                if keyword.eq_ignore_ascii_case("BARCODE") =>
            {
                FieldType::Barcode
            },
            ParsedFieldCode::Other { ref keyword, .. }
                if keyword.eq_ignore_ascii_case("BIDIOUTLINE") =>
            {
                FieldType::BidiOutline
            },
            ParsedFieldCode::Other { ref keyword, .. }
                if keyword.eq_ignore_ascii_case("SHAPE") =>
            {
                FieldType::Shape
            },
            ParsedFieldCode::Other { ref keyword, .. }
                if keyword.eq_ignore_ascii_case("FORMTEXT") =>
            {
                FieldType::FormText
            },
            ParsedFieldCode::Other { ref keyword, .. }
                if keyword.eq_ignore_ascii_case("FORMCHECKBOX") =>
            {
                FieldType::FormCheckbox
            },
            ParsedFieldCode::Other { ref keyword, .. }
                if keyword.eq_ignore_ascii_case("FORMDROPDOWN") =>
            {
                FieldType::FormDropdown
            },
            ParsedFieldCode::Other { ref keyword, .. } if keyword.eq_ignore_ascii_case("ADDIN") => {
                FieldType::AddIn
            },
            ParsedFieldCode::Other { ref keyword, .. }
                if keyword.eq_ignore_ascii_case("CONTROL") =>
            {
                FieldType::Control
            },
            ParsedFieldCode::Other { ref keyword, .. }
                if keyword.eq_ignore_ascii_case("HTMLCONTROL") =>
            {
                FieldType::HtmlControl
            },
            ParsedFieldCode::Other { ref keyword, .. }
                if keyword.eq_ignore_ascii_case("GLOSSARY") =>
            {
                FieldType::Glossary
            },
            ParsedFieldCode::Other { ref keyword, .. }
                if keyword.eq_ignore_ascii_case("AUTOTEXT") =>
            {
                FieldType::AutoText
            },
            ParsedFieldCode::Other { ref keyword, .. }
                if keyword.eq_ignore_ascii_case("AUTOTEXTLIST") =>
            {
                FieldType::AutoTextList
            },
            ParsedFieldCode::Other { ref keyword, .. }
                if keyword.eq_ignore_ascii_case("USERADDRESS") =>
            {
                FieldType::UserAddress
            },
            ParsedFieldCode::Other { ref keyword, .. }
                if keyword.eq_ignore_ascii_case("USERINITIALS") =>
            {
                FieldType::UserInitials
            },
            ParsedFieldCode::Other { ref keyword, .. }
                if keyword.eq_ignore_ascii_case("USERNAME") =>
            {
                FieldType::UserName
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
            ParsedFieldCode::Other { ref keyword, .. } if keyword.eq_ignore_ascii_case("INDEX") => {
                FieldType::Index
            },
            ParsedFieldCode::Other { ref keyword, .. } if keyword.eq_ignore_ascii_case("XE") => {
                FieldType::IndexEntry
            },
            ParsedFieldCode::Other { ref keyword, .. }
                if keyword.eq_ignore_ascii_case("CITATION") =>
            {
                FieldType::Citation
            },
            ParsedFieldCode::Other { ref keyword, .. }
                if keyword.eq_ignore_ascii_case("BIBLIOGRAPHY") =>
            {
                FieldType::Bibliography
            },
            ParsedFieldCode::Other { ref keyword, .. }
                if keyword.eq_ignore_ascii_case("DOCVARIABLE") =>
            {
                FieldType::DocumentVariable
            },
            ParsedFieldCode::Other { ref keyword, .. }
                if keyword.eq_ignore_ascii_case("DOCPROPERTY") =>
            {
                FieldType::DocumentProperty
            },
            ParsedFieldCode::Other { ref keyword, .. } if keyword.eq_ignore_ascii_case("INFO") => {
                FieldType::Info
            },
            ParsedFieldCode::Other { ref keyword, .. }
                if DocumentInformationFieldKind::from_keyword(keyword.as_ref()).is_some() =>
            {
                FieldType::DocumentInformation
            },
            ParsedFieldCode::Other { ref keyword, .. }
                if DocumentContextFieldKind::from_keyword(keyword.as_ref()).is_some() =>
            {
                FieldType::DocumentContext
            },
            ParsedFieldCode::Other { ref keyword, .. }
                if keyword.eq_ignore_ascii_case("MERGEFIELD") =>
            {
                FieldType::MergeField
            },
            ParsedFieldCode::Other { ref keyword, .. } if keyword.eq_ignore_ascii_case("DATA") => {
                FieldType::MailMergeData
            },
            ParsedFieldCode::Other { ref keyword, .. }
                if keyword.eq_ignore_ascii_case("MERGEREC") =>
            {
                FieldType::MergeRecord
            },
            ParsedFieldCode::Other { ref keyword, .. }
                if keyword.eq_ignore_ascii_case("MERGESEQ") =>
            {
                FieldType::MergeSequence
            },
            ParsedFieldCode::Other { ref keyword, .. } if keyword.eq_ignore_ascii_case("NEXT") => {
                FieldType::MailMergeNext
            },
            ParsedFieldCode::Other { ref keyword, .. }
                if keyword.eq_ignore_ascii_case("NEXTIF") =>
            {
                FieldType::MailMergeNextIf
            },
            ParsedFieldCode::Other { ref keyword, .. }
                if keyword.eq_ignore_ascii_case("SKIPIF") =>
            {
                FieldType::MailMergeSkipIf
            },
            ParsedFieldCode::Other { ref keyword, .. } if keyword.eq_ignore_ascii_case("IF") => {
                FieldType::If
            },
            ParsedFieldCode::Other { ref keyword, .. }
                if keyword.eq_ignore_ascii_case("COMPARE") =>
            {
                FieldType::Compare
            },
            ParsedFieldCode::Other { ref keyword, .. } if keyword.eq_ignore_ascii_case("SET") => {
                FieldType::Set
            },
            ParsedFieldCode::Other { ref keyword, .. } if keyword.eq_ignore_ascii_case("SEQ") => {
                FieldType::Sequence
            },
            ParsedFieldCode::Other { ref keyword, .. } if keyword.eq_ignore_ascii_case("QUOTE") => {
                FieldType::Quote
            },
            ParsedFieldCode::Other { ref keyword, .. } if keyword.eq_ignore_ascii_case("SYMBOL") => {
                FieldType::Symbol
            },
            ParsedFieldCode::Other { ref keyword, .. }
                if AutoNumberFieldKind::from_keyword(keyword.as_ref()).is_some() =>
            {
                FieldType::AutoNumber
            },
            ParsedFieldCode::Other { ref keyword, .. } if keyword.eq_ignore_ascii_case("LISTNUM") => {
                FieldType::ListNumber
            },
            ParsedFieldCode::Other { ref keyword, .. }
                if keyword.eq_ignore_ascii_case("STYLEREF") =>
            {
                FieldType::StyleReference
            },
            ParsedFieldCode::Other { ref keyword, .. } if keyword.eq_ignore_ascii_case("ASK") => {
                FieldType::Ask
            },
            ParsedFieldCode::Other { ref keyword, .. }
                if keyword.eq_ignore_ascii_case("FILLIN") =>
            {
                FieldType::FillIn
            },
            ParsedFieldCode::Other { ref keyword, .. }
                if keyword.eq_ignore_ascii_case("ADDRESSBLOCK") =>
            {
                FieldType::AddressBlock
            },
            ParsedFieldCode::Other { ref keyword, .. }
                if keyword.eq_ignore_ascii_case("GREETINGLINE") =>
            {
                FieldType::GreetingLine
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

    /// Return inert metadata when this is a well-formed `GOTOBUTTON` field.
    ///
    /// The destination and button text remain stored metadata only. This method
    /// never resolves a bookmark, page, annotation, footnote, or other target,
    /// changes the insertion point, or activates a jump. Malformed navigation
    /// instructions remain generic fields and return `None` here.
    pub fn go_to_button(&self) -> Option<GoToButtonField<'_>> {
        if self.field_type != FieldType::GoToButton {
            return None;
        }
        let (target, button_text) = go_to_button_parts(self.instruction.as_ref())?;
        Some(GoToButtonField {
            instruction: self.instruction.as_ref(),
            target,
            button_text,
            cached_result: (!self.result.is_empty()).then_some(self.result.as_ref()),
            status: self.status,
            owner: self.owner,
            position: self.position,
        })
    }

    /// Return inert metadata when this is a `PRINT` field.
    ///
    /// Stored printer-instruction text, cached results, and field state remain
    /// opaque metadata only. This method never interprets printer-control
    /// codes, opens a printer, sends output, changes print settings, or
    /// refreshes a field.
    pub fn print_field(&self) -> Option<PrintField<'_>> {
        if self.field_type != FieldType::Print {
            return None;
        }
        let printer_instructions = print_field_instructions(self.instruction.as_ref())?;
        Some(PrintField {
            instruction: self.instruction.as_ref(),
            printer_instructions,
            cached_result: (!self.result.is_empty()).then_some(self.result.as_ref()),
            status: self.status,
            owner: self.owner,
            position: self.position,
        })
    }

    /// Return inert metadata when this is an `EMBED` field.
    ///
    /// Stored opaque object instructions, cached results, and field state remain
    /// metadata only. This method never loads, inspects, deserializes,
    /// activates, renders, or executes an embedded object, accesses an external
    /// resource, or refreshes a field.
    pub fn embed_field(&self) -> Option<EmbedField<'_>> {
        if self.field_type != FieldType::Embed {
            return None;
        }
        let object_instructions = embed_field_instructions(self.instruction.as_ref())?;
        Some(EmbedField {
            instruction: self.instruction.as_ref(),
            object_instructions,
            cached_result: (!self.result.is_empty()).then_some(self.result.as_ref()),
            status: self.status,
            owner: self.owner,
            position: self.position,
        })
    }

    /// Return inert metadata when this is a `BARCODE` field.
    ///
    /// Stored opaque barcode instructions, cached results, and field state
    /// remain metadata only. This method never parses or validates barcode data
    /// or symbology, generates or renders a barcode, accesses an external
    /// resource, or refreshes a field.
    pub fn barcode_field(&self) -> Option<BarcodeField<'_>> {
        if self.field_type != FieldType::Barcode {
            return None;
        }
        let barcode_instructions = barcode_field_instructions(self.instruction.as_ref())?;
        Some(BarcodeField {
            instruction: self.instruction.as_ref(),
            barcode_instructions,
            cached_result: (!self.result.is_empty()).then_some(self.result.as_ref()),
            status: self.status,
            owner: self.owner,
            position: self.position,
        })
    }

    /// Return inert metadata when this is a `BIDIOUTLINE` field.
    ///
    /// Stored opaque instructions, cached results, and field state remain
    /// metadata only. This method never reads right-to-left language, paragraph
    /// outline, or layout state; chooses a numbering system; calculates a
    /// result; or refreshes a field.
    pub fn bidi_outline_field(&self) -> Option<BidiOutlineField<'_>> {
        if self.field_type != FieldType::BidiOutline {
            return None;
        }
        let opaque_instructions = bidi_outline_field_instructions(self.instruction.as_ref())?;
        Some(BidiOutlineField {
            instruction: self.instruction.as_ref(),
            opaque_instructions,
            cached_result: (!self.result.is_empty()).then_some(self.result.as_ref()),
            status: self.status,
            owner: self.owner,
            position: self.position,
        })
    }

    /// Return inert metadata when this is a `SHAPE` drawing-canvas anchor field.
    ///
    /// Stored opaque instructions, cached results, and field state remain
    /// metadata only. This method never locates, links, loads, positions, lays
    /// out, or renders a drawing or canvas, or refreshes a field.
    pub fn shape_field(&self) -> Option<ShapeField<'_>> {
        if self.field_type != FieldType::Shape {
            return None;
        }
        let opaque_instructions = shape_field_instructions(self.instruction.as_ref())?;
        Some(ShapeField {
            instruction: self.instruction.as_ref(),
            opaque_instructions,
            cached_result: (!self.result.is_empty()).then_some(self.result.as_ref()),
            status: self.status,
            owner: self.owner,
            position: self.position,
        })
    }

    /// Return inert metadata when this is a legacy form-code field.
    ///
    /// Stored kind, opaque instructions, cached results, and field state remain
    /// metadata only. This method does not read or reconcile the separate RTF
    /// `\formfield` destination. It never fills a form, changes a selection or
    /// checkbox state, invokes entry or exit macros, or refreshes a field.
    pub fn legacy_form_field(&self) -> Option<LegacyFormField<'_>> {
        let (kind, keyword) = match self.field_type {
            FieldType::FormText => (LegacyFormFieldKind::Text, "FORMTEXT"),
            FieldType::FormCheckbox => (LegacyFormFieldKind::CheckBox, "FORMCHECKBOX"),
            FieldType::FormDropdown => (LegacyFormFieldKind::DropDown, "FORMDROPDOWN"),
            _ => return None,
        };
        let opaque_instructions =
            legacy_form_field_instructions(self.instruction.as_ref(), keyword)?;
        Some(LegacyFormField {
            instruction: self.instruction.as_ref(),
            kind,
            opaque_instructions,
            cached_result: (!self.result.is_empty()).then_some(self.result.as_ref()),
            status: self.status,
            owner: self.owner,
            position: self.position,
        })
    }

    /// Return inert metadata when this is an add-in or control field.
    ///
    /// The stored instruction and cached result remain opaque metadata only.
    /// This method never loads an add-in, instantiates a control, invokes code,
    /// executes script, renders content, accesses an external resource, or
    /// refreshes a field.
    pub fn active_content_field(&self) -> Option<ActiveContentField<'_>> {
        let kind = match self.field_type {
            FieldType::AddIn => ActiveContentFieldKind::AddIn,
            FieldType::Control => ActiveContentFieldKind::OcxControl,
            FieldType::HtmlControl => ActiveContentFieldKind::HtmlControl,
            _ => return None,
        };
        Some(ActiveContentField {
            instruction: self.instruction.as_ref(),
            kind,
            cached_result: (!self.result.is_empty()).then_some(self.result.as_ref()),
            status: self.status,
            owner: self.owner,
            position: self.position,
        })
    }

    /// Return inert metadata when this is a `GLOSSARY` or `AUTOTEXT` field.
    ///
    /// Stored entry names, switches, and cached results are never used to look
    /// up a building block, read a template, insert content, change bookmarks,
    /// access an external resource, or refresh a field. Malformed instructions
    /// remain generic fields and return `None` here.
    pub fn auto_text_field(&self) -> Option<AutoTextField<'_>> {
        let expected_kind = match self.field_type {
            FieldType::Glossary => AutoTextFieldKind::Glossary,
            FieldType::AutoText => AutoTextFieldKind::AutoText,
            _ => return None,
        };
        let (kind, entry_name, unknown_switches) =
            auto_text_field_parts(self.instruction.as_ref())?;
        if kind != expected_kind {
            return None;
        }
        Some(AutoTextField {
            instruction: self.instruction.as_ref(),
            kind,
            entry_name,
            unknown_switches,
            cached_result: (!self.result.is_empty()).then_some(self.result.as_ref()),
            status: self.status,
            owner: self.owner,
            position: self.position,
        })
    }

    /// Return inert metadata when this is a well-formed `AUTOTEXTLIST` field.
    ///
    /// Stored display text, style/tip options, and cached results are never
    /// used to show a selection UI, look up a building block, read a template,
    /// insert content, change bookmarks, access an external resource, or
    /// refresh a field. Malformed instructions remain generic fields and
    /// return `None` here.
    pub fn auto_text_list_field(&self) -> Option<AutoTextListField<'_>> {
        if self.field_type != FieldType::AutoTextList {
            return None;
        }
        let (display_text, options, unknown_switches) =
            auto_text_list_field_parts(self.instruction.as_ref())?;
        Some(AutoTextListField {
            instruction: self.instruction.as_ref(),
            display_text,
            options,
            unknown_switches,
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

    /// Return inert metadata when this is a well-formed INDEX field.
    ///
    /// The stored configuration and cached result are never used to scan XE
    /// markers, follow bookmarks, calculate page numbers, paginate the
    /// document, or generate an index. Malformed INDEX instructions remain
    /// generic fields and return None here.
    pub fn index(&self) -> Option<IndexField<'_>> {
        if self.field_type != FieldType::Index {
            return None;
        }
        let parts = index_parts(self.instruction.as_ref())?;
        Some(IndexField {
            instruction: self.instruction.as_ref(),
            options: parts.options,
            unknown_switches: parts.unknown_switches,
            cached_result: (!self.result.is_empty()).then_some(self.result.as_ref()),
            status: self.status,
            owner: self.owner,
            position: self.position,
        })
    }

    /// Return inert metadata when this is a well-formed XE index-entry field.
    ///
    /// The stored entry and cached result are never used to change hidden text,
    /// resolve bookmarks, calculate page numbers, or generate an index.
    /// Malformed XE instructions remain generic fields and return None here.
    pub fn index_entry(&self) -> Option<IndexEntryField<'_>> {
        if self.field_type != FieldType::IndexEntry {
            return None;
        }
        let parts = index_entry_parts(self.instruction.as_ref())?;
        Some(IndexEntryField {
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

    /// Return inert metadata when this is a well-formed CITATION field.
    ///
    /// Stored source tags and options are never resolved, loaded, styled, or
    /// formatted. Malformed CITATION instructions remain generic fields and
    /// return None here.
    pub fn citation(&self) -> Option<CitationField<'_>> {
        if self.field_type != FieldType::Citation {
            return None;
        }
        let parts = citation_parts(self.instruction.as_ref())?;
        Some(CitationField {
            instruction: self.instruction.as_ref(),
            source_tag: parts.source_tag,
            options: parts.options,
            unknown_switches: parts.unknown_switches,
            cached_result: (!self.result.is_empty()).then_some(self.result.as_ref()),
            status: self.status,
            owner: self.owner,
            position: self.position,
        })
    }

    /// Return inert metadata when this is a well-formed BIBLIOGRAPHY field.
    ///
    /// Stored options are never used to load source records, apply a style,
    /// sort entries, or generate bibliography content. Malformed BIBLIOGRAPHY
    /// instructions remain generic fields and return None here.
    pub fn bibliography(&self) -> Option<BibliographyField<'_>> {
        if self.field_type != FieldType::Bibliography {
            return None;
        }
        let parts = bibliography_parts(self.instruction.as_ref())?;
        Some(BibliographyField {
            instruction: self.instruction.as_ref(),
            options: parts.options,
            unknown_switches: parts.unknown_switches,
            cached_result: (!self.result.is_empty()).then_some(self.result.as_ref()),
            status: self.status,
            owner: self.owner,
            position: self.position,
        })
    }

    /// Return inert metadata when this is a well-formed DOCVARIABLE field.
    ///
    /// The stored variable name is never resolved against document variables,
    /// and the cached result is never refreshed. Malformed DOCVARIABLE
    /// instructions remain generic fields and return None here.
    pub fn document_variable(&self) -> Option<DocumentVariableField<'_>> {
        if self.field_type != FieldType::DocumentVariable {
            return None;
        }
        let parts = document_variable_field_parts(self.instruction.as_ref())?;
        Some(DocumentVariableField {
            instruction: self.instruction.as_ref(),
            variable_name: parts.variable_name,
            unknown_switches: parts.unknown_switches,
            cached_result: (!self.result.is_empty()).then_some(self.result.as_ref()),
            status: self.status,
            owner: self.owner,
            position: self.position,
        })
    }

    /// Return inert metadata when this is a well-formed `DOCPROPERTY` field.
    ///
    /// The stored property name is never resolved against core, extended, or
    /// custom document properties, and the cached result is never refreshed.
    /// Malformed `DOCPROPERTY` instructions remain generic fields and return
    /// `None` here.
    pub fn document_property(&self) -> Option<DocumentPropertyField<'_>> {
        if self.field_type != FieldType::DocumentProperty {
            return None;
        }
        let parts = document_property_field_parts(self.instruction.as_ref())?;
        Some(DocumentPropertyField {
            instruction: self.instruction.as_ref(),
            property_name: parts.property_name,
            switches: parts.switches,
            cached_result: (!self.result.is_empty()).then_some(self.result.as_ref()),
            status: self.status,
            owner: self.owner,
            position: self.position,
        })
    }

    /// Return inert metadata when this is a well-formed explicit `INFO` field.
    ///
    /// Stored property selectors, optional replacement values, switches, cached
    /// results, and field state remain metadata only. This method never reads,
    /// resolves, modifies, or writes document or template properties, or
    /// refreshes a field. Malformed `INFO` instructions remain generic
    /// fields and return `None` here.
    pub fn info_field(&self) -> Option<InfoField<'_>> {
        if self.field_type != FieldType::Info {
            return None;
        }
        let parts = info_field_parts(self.instruction.as_ref())?;
        Some(InfoField {
            instruction: self.instruction.as_ref(),
            information_type: parts.information_type,
            new_value: parts.new_value,
            switches: parts.switches,
            cached_result: (!self.result.is_empty()).then_some(self.result.as_ref()),
            status: self.status,
            owner: self.owner,
            position: self.position,
        })
    }

    /// Return inert metadata when this is a well-formed document-information
    /// field.
    ///
    /// Built-in document-information fields retain only their stored kind,
    /// switches, cached result, and state. This method never reads document
    /// properties or host identity data, calculates dates, revisions, or
    /// statistics, resolves a value, or refreshes a field. Malformed
    /// instructions remain generic fields and return `None` here.
    pub fn document_information(&self) -> Option<DocumentInformationField<'_>> {
        if self.field_type != FieldType::DocumentInformation {
            return None;
        }
        let parts = document_information_field_parts(self.instruction.as_ref())?;
        Some(DocumentInformationField {
            instruction: self.instruction.as_ref(),
            kind: parts.kind,
            switches: parts.switches,
            cached_result: (!self.result.is_empty()).then_some(self.result.as_ref()),
            status: self.status,
            owner: self.owner,
            position: self.position,
        })
    }

    /// Return inert metadata when this is a well-formed document-context or
    /// runtime field.
    ///
    /// `FILENAME`, `TEMPLATE`, `DATE`, `TIME`, `PAGE`, `FILESIZE`, `SECTION`, and
    /// `SECTIONPAGES` retain only their stored kind, switches, cached result, and
    /// state. This method never reads a document path, attached template, host
    /// filesystem state or file size, current clock, or page and section layout,
    /// resolves a value, or refreshes a field. Malformed instructions remain
    /// generic fields and return `None` here.
    pub fn document_context(&self) -> Option<DocumentContextField<'_>> {
        let parts = document_context_field_parts(self.instruction.as_ref())?;
        if !matches!(
            (self.field_type, parts.kind),
            (
                FieldType::DocumentContext,
                DocumentContextFieldKind::FileName
                    | DocumentContextFieldKind::Template
                    | DocumentContextFieldKind::FileSize
                    | DocumentContextFieldKind::Section
                    | DocumentContextFieldKind::SectionPages
            ) | (
                FieldType::Date,
                DocumentContextFieldKind::Date | DocumentContextFieldKind::Time
            ) | (FieldType::Page, DocumentContextFieldKind::Page)
        ) {
            return None;
        }
        Some(DocumentContextField {
            instruction: self.instruction.as_ref(),
            kind: parts.kind,
            switches: parts.switches,
            cached_result: (!self.result.is_empty()).then_some(self.result.as_ref()),
            status: self.status,
            owner: self.owner,
            position: self.position,
        })
    }

    /// Return inert metadata when this is a well-formed `MERGEFIELD` field.
    ///
    /// The stored field name, switches, and cached result are never resolved
    /// against a data source, merged into the document, or refreshed.
    /// Malformed `MERGEFIELD` instructions remain generic fields and return
    /// `None` here.
    pub fn merge_field(&self) -> Option<MergeField<'_>> {
        if self.field_type != FieldType::MergeField {
            return None;
        }
        let parts = merge_field_parts(self.instruction.as_ref())?;
        Some(MergeField {
            instruction: self.instruction.as_ref(),
            field_name: parts.field_name,
            switches: parts.switches,
            cached_result: (!self.result.is_empty()).then_some(self.result.as_ref()),
            status: self.status,
            owner: self.owner,
            position: self.position,
        })
    }

    /// Return inert metadata when this is a well-formed `DATA` mail-merge
    /// source field.
    ///
    /// The stored data-source and header-source identifiers, switches, and
    /// cached result are never opened, read, connected to, resolved, modified,
    /// or merged. Malformed `DATA` instructions remain generic fields and return
    /// `None` here.
    pub fn mail_merge_data(&self) -> Option<MailMergeDataField<'_>> {
        if self.field_type != FieldType::MailMergeData {
            return None;
        }
        let parts = mail_merge_data_field_parts(self.instruction.as_ref())?;
        Some(MailMergeDataField {
            instruction: self.instruction.as_ref(),
            data_source: parts.data_source,
            header_source: parts.header_source,
            switches: parts.switches,
            cached_result: (!self.result.is_empty()).then_some(self.result.as_ref()),
            status: self.status,
            owner: self.owner,
            position: self.position,
        })
    }

    /// Return inert metadata when this is a well-formed mail-merge counter.
    ///
    /// The stored kind and cached result are never used to select or count
    /// records, open a data source, perform a merge, or refresh a field
    /// result. Malformed `MERGEREC` and `MERGESEQ` instructions remain generic
    /// fields and return `None` here.
    pub fn mail_merge_counter(&self) -> Option<MailMergeCounterField<'_>> {
        let kind = match self.field_type {
            FieldType::MergeRecord => MailMergeCounterKind::Record,
            FieldType::MergeSequence => MailMergeCounterKind::Sequence,
            _ => return None,
        };
        if mail_merge_counter_kind(self.instruction.as_ref()) != Some(kind) {
            return None;
        }
        Some(MailMergeCounterField {
            instruction: self.instruction.as_ref(),
            kind,
            cached_result: (!self.result.is_empty()).then_some(self.result.as_ref()),
            status: self.status,
            owner: self.owner,
            position: self.position,
        })
    }

    /// Return inert metadata when this is a well-formed `NEXT` field.
    ///
    /// Cached text and field state are never used to advance a record, open a
    /// data source, perform a merge, or refresh a field result. Malformed
    /// `NEXT` instructions remain generic fields and return `None` here.
    pub fn mail_merge_next(&self) -> Option<MailMergeNextField<'_>> {
        if self.field_type != FieldType::MailMergeNext
            || !is_mail_merge_next_instruction(self.instruction.as_ref())
        {
            return None;
        }
        Some(MailMergeNextField {
            instruction: self.instruction.as_ref(),
            cached_result: (!self.result.is_empty()).then_some(self.result.as_ref()),
            status: self.status,
            owner: self.owner,
            position: self.position,
        })
    }

    /// Return inert metadata when this is a `NEXTIF` or `SKIPIF` field with a comparison.
    ///
    /// The stored comparison and cached result are never parsed or evaluated.
    /// This method never changes record selection, opens a data source,
    /// performs a merge, or refreshes a field result. Instructions without a
    /// comparison remain generic fields and return `None` here.
    pub fn mail_merge_conditional_control(&self) -> Option<MailMergeConditionalControlField<'_>> {
        let expected_kind = match self.field_type {
            FieldType::MailMergeNextIf => MailMergeConditionalControlKind::NextIf,
            FieldType::MailMergeSkipIf => MailMergeConditionalControlKind::SkipIf,
            _ => return None,
        };
        let (kind, comparison) = mail_merge_conditional_control_parts(self.instruction.as_ref())?;
        if kind != expected_kind {
            return None;
        }
        Some(MailMergeConditionalControlField {
            instruction: self.instruction.as_ref(),
            kind,
            comparison,
            cached_result: (!self.result.is_empty()).then_some(self.result.as_ref()),
            status: self.status,
            owner: self.owner,
            position: self.position,
        })
    }

    /// Return inert metadata when this is an `IF` field with an expression.
    ///
    /// The stored expression and cached result are never parsed or evaluated.
    /// This method never resolves field values or refreshes a field result.
    /// Instructions without an expression remain generic fields and return
    /// `None` here.
    pub fn if_field(&self) -> Option<IfField<'_>> {
        if self.field_type != FieldType::If {
            return None;
        }
        let expression = if_field_expression(self.instruction.as_ref())?;
        Some(IfField {
            instruction: self.instruction.as_ref(),
            expression,
            cached_result: (!self.result.is_empty()).then_some(self.result.as_ref()),
            status: self.status,
            owner: self.owner,
            position: self.position,
        })
    }

    /// Return inert metadata when this is a `COMPARE` field with a comparison.
    ///
    /// The stored comparison and cached result are never parsed or evaluated.
    /// This method never resolves nested field values or refreshes a field
    /// result. Instructions without a comparison remain generic fields and
    /// return `None` here.
    pub fn compare_field(&self) -> Option<CompareField<'_>> {
        if self.field_type != FieldType::Compare {
            return None;
        }
        let comparison = compare_field_comparison(self.instruction.as_ref())?;
        Some(CompareField {
            instruction: self.instruction.as_ref(),
            comparison,
            cached_result: (!self.result.is_empty()).then_some(self.result.as_ref()),
            status: self.status,
            owner: self.owner,
            position: self.position,
        })
    }

    /// Return inert metadata when this is a `SET` field with a target and expression.
    ///
    /// The stored expression and cached result are never parsed or evaluated.
    /// This method never looks up or changes a bookmark, changes document
    /// state, or refreshes a field result. Instructions without both a target
    /// and expression remain generic fields and return `None` here.
    pub fn set_field(&self) -> Option<SetField<'_>> {
        if self.field_type != FieldType::Set {
            return None;
        }
        let (target_name, expression) = set_field_parts(self.instruction.as_ref())?;
        Some(SetField {
            instruction: self.instruction.as_ref(),
            target_name,
            expression,
            cached_result: (!self.result.is_empty()).then_some(self.result.as_ref()),
            status: self.status,
            owner: self.owner,
            position: self.position,
        })
    }

    /// Return inert metadata when this is a `SEQ` field with an identifier.
    ///
    /// The stored identifier, bookmark, tail, and cached result are never used
    /// to look up a bookmark, increment or reset a sequence, calculate a
    /// number, or refresh a field. Instructions without an identifier remain
    /// generic fields and return `None` here.
    pub fn sequence_field(&self) -> Option<SequenceField<'_>> {
        if self.field_type != FieldType::Sequence {
            return None;
        }
        let (identifier, bookmark, tail) = sequence_field_parts(self.instruction.as_ref())?;
        Some(SequenceField {
            instruction: self.instruction.as_ref(),
            identifier,
            bookmark,
            tail,
            cached_result: (!self.result.is_empty()).then_some(self.result.as_ref()),
            status: self.status,
            owner: self.owner,
            position: self.position,
        })
    }

    /// Return inert metadata when this is a formula field with a stored formula.
    ///
    /// The stored formula and cached result are never parsed or evaluated. This
    /// method never reads table cells or bookmarks, resolves field values, or
    /// refreshes a field result. Instructions without a formula remain generic
    /// fields and return `None` here.
    pub fn formula_field(&self) -> Option<FormulaField<'_>> {
        if self.field_type != FieldType::Formula {
            return None;
        }
        let formula = formula_field_formula(self.instruction.as_ref())?;
        Some(FormulaField {
            instruction: self.instruction.as_ref(),
            formula,
            cached_result: (!self.result.is_empty()).then_some(self.result.as_ref()),
            status: self.status,
            owner: self.owner,
            position: self.position,
        })
    }

    /// Return inert metadata when this is a `QUOTE` field with a text argument.
    ///
    /// Stored text, switches, and cached results are never used to interpret
    /// character codes, expand nested fields, insert text, or refresh a field.
    /// Instructions without a text argument or with malformed switches remain
    /// generic fields and return `None` here.
    pub fn quote_field(&self) -> Option<QuoteField<'_>> {
        if self.field_type != FieldType::Quote {
            return None;
        }
        let (text, switches) = quote_field_parts(self.instruction.as_ref())?;
        Some(QuoteField {
            instruction: self.instruction.as_ref(),
            text,
            switches,
            cached_result: (!self.result.is_empty()).then_some(self.result.as_ref()),
            status: self.status,
            owner: self.owner,
            position: self.position,
        })
    }

    /// Return inert metadata when this is a `SYMBOL` field with a character
    /// argument.
    ///
    /// Stored character arguments, switches, and cached results are never used
    /// to map a character code, look up a font, insert a glyph, change
    /// formatting or layout, or refresh a field. Instructions without a
    /// character argument or with malformed switches remain generic fields and
    /// return `None` here.
    pub fn symbol_field(&self) -> Option<SymbolField<'_>> {
        if self.field_type != FieldType::Symbol {
            return None;
        }
        let (character_argument, switches) = symbol_field_parts(self.instruction.as_ref())?;
        Some(SymbolField {
            instruction: self.instruction.as_ref(),
            character_argument,
            switches,
            cached_result: (!self.result.is_empty()).then_some(self.result.as_ref()),
            status: self.status,
            owner: self.owner,
            position: self.position,
        })
    }

    /// Return inert metadata when this is a legacy automatic-numbering field.
    ///
    /// Stored kinds, switches, cached results, and field state are never used
    /// to calculate paragraph numbers, read heading or style state, change
    /// paragraphs or layout, or refresh a field. Malformed instructions remain
    /// generic fields and return `None` here.
    pub fn auto_number_field(&self) -> Option<AutoNumberField<'_>> {
        if self.field_type != FieldType::AutoNumber {
            return None;
        }
        let parts = auto_number_field_parts(self.instruction.as_ref())?;
        Some(AutoNumberField {
            instruction: self.instruction.as_ref(),
            kind: parts.kind,
            switches: parts.switches,
            cached_result: (!self.result.is_empty()).then_some(self.result.as_ref()),
            status: self.status,
            owner: self.owner,
            position: self.position,
        })
    }

    /// Return inert metadata when this is a `LISTNUM` field.
    ///
    /// Stored optional list names, switches, cached results, and field state
    /// are never used to look up a list, determine a level or start value,
    /// calculate a number, change layout, or refresh a field. Malformed
    /// instructions remain generic fields and return `None` here.
    pub fn list_number_field(&self) -> Option<ListNumberField<'_>> {
        if self.field_type != FieldType::ListNumber {
            return None;
        }
        let parts = list_number_field_parts(self.instruction.as_ref())?;
        Some(ListNumberField {
            instruction: self.instruction.as_ref(),
            list_name: parts.list_name,
            switches: parts.switches,
            cached_result: (!self.result.is_empty()).then_some(self.result.as_ref()),
            status: self.status,
            owner: self.owner,
            position: self.position,
        })
    }

    /// Return inert metadata when this is a `STYLEREF` field with a style name.
    ///
    /// The stored style name, options, switches, and cached result are never
    /// used to look up styled text, search document stories, calculate paragraph
    /// numbers or relative positions, resolve page layout, or refresh a field.
    /// Instructions without a style name or with malformed switches remain
    /// generic fields and return `None` here.
    pub fn style_reference_field(&self) -> Option<StyleReferenceField<'_>> {
        if self.field_type != FieldType::StyleReference {
            return None;
        }
        let (style_name, options, unknown_switches) =
            style_reference_field_parts(self.instruction.as_ref())?;
        Some(StyleReferenceField {
            instruction: self.instruction.as_ref(),
            style_name,
            options,
            unknown_switches,
            cached_result: (!self.result.is_empty()).then_some(self.result.as_ref()),
            status: self.status,
            owner: self.owner,
            position: self.position,
        })
    }

    /// Return inert metadata when this is a well-formed `ASK` or `FILLIN` field.
    ///
    /// Stored prompt, bookmark, default-response, and cached-result data are
    /// never used to display a prompt, capture a response, create or update a
    /// bookmark, perform a merge, or refresh a field. Malformed instructions
    /// remain generic fields and return `None` here.
    pub fn prompt_field(&self) -> Option<PromptField<'_>> {
        let expected_kind = match self.field_type {
            FieldType::Ask => PromptFieldKind::Ask,
            FieldType::FillIn => PromptFieldKind::FillIn,
            _ => return None,
        };
        let parts = prompt_field_parts(self.instruction.as_ref())?;
        if parts.kind != expected_kind {
            return None;
        }
        Some(PromptField {
            instruction: self.instruction.as_ref(),
            kind: parts.kind,
            bookmark: parts.bookmark,
            prompt: parts.prompt,
            default_response: parts.default_response,
            prompts_once_per_mail_merge: parts.prompts_once_per_mail_merge,
            cached_result: (!self.result.is_empty()).then_some(self.result.as_ref()),
            status: self.status,
            owner: self.owner,
            position: self.position,
        })
    }

    /// Return inert metadata when this is a well-formed user-identity field.
    ///
    /// Stored override, formatting, and cached-result data are never used to
    /// read or modify a host user's identity, apply formatting, or refresh a
    /// field. Malformed instructions remain generic fields and return `None`
    /// here.
    pub fn user_identity_field(&self) -> Option<UserIdentityField<'_>> {
        let expected_kind = match self.field_type {
            FieldType::UserAddress => UserIdentityFieldKind::Address,
            FieldType::UserInitials => UserIdentityFieldKind::Initials,
            FieldType::UserName => UserIdentityFieldKind::Name,
            _ => return None,
        };
        let parts = user_identity_field_parts(self.instruction.as_ref())?;
        if parts.kind != expected_kind {
            return None;
        }
        Some(UserIdentityField {
            instruction: self.instruction.as_ref(),
            kind: parts.kind,
            override_value: parts.override_value,
            formatting: parts.formatting,
            cached_result: (!self.result.is_empty()).then_some(self.result.as_ref()),
            status: self.status,
            owner: self.owner,
            position: self.position,
        })
    }

    /// Return inert metadata when this is a well-formed `ADVANCE` field.
    ///
    /// Stored point adjustments and cached-result data are never used to move
    /// text, change layout, reflow content, or refresh a field. Malformed
    /// instructions remain generic fields and return `None` here.
    pub fn advance_field(&self) -> Option<AdvanceField<'_>> {
        if self.field_type != FieldType::Advance {
            return None;
        }
        let adjustments = advance_field_adjustments(self.instruction.as_ref())?;
        Some(AdvanceField {
            instruction: self.instruction.as_ref(),
            adjustments,
            cached_result: (!self.result.is_empty()).then_some(self.result.as_ref()),
            status: self.status,
            owner: self.owner,
            position: self.position,
        })
    }

    /// Return inert metadata when this is a well-formed `ADDRESSBLOCK` or
    /// `GREETINGLINE` field.
    ///
    /// Stored layout, locale, country, fallback, and cached-result data are
    /// never used to open a data source, select a record, perform a merge,
    /// expand placeholders, generate text, or refresh a field. Malformed
    /// instructions remain generic fields and return `None` here.
    pub fn mail_merge_recipient_field(&self) -> Option<MailMergeRecipientField<'_>> {
        let expected_kind = match self.field_type {
            FieldType::AddressBlock => MailMergeRecipientFieldKind::AddressBlock,
            FieldType::GreetingLine => MailMergeRecipientFieldKind::GreetingLine,
            _ => return None,
        };
        let parts = mail_merge_recipient_field_parts(self.instruction.as_ref())?;
        if parts.kind != expected_kind {
            return None;
        }
        Some(MailMergeRecipientField {
            instruction: self.instruction.as_ref(),
            kind: parts.kind,
            country_inclusion: parts.country_inclusion,
            formats_using_recipient_country: parts.formats_using_recipient_country,
            excluded_countries: parts.excluded_countries,
            format_template: parts.format_template,
            language: parts.language,
            greeting_fallback_text: parts.greeting_fallback_text,
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

fn print_field_instructions(instruction: &str) -> Option<&str> {
    if instruction.len() > MAX_INSTRUCTION_LEN {
        return None;
    }
    let instruction = instruction.trim_start_matches(|value: char| value.is_ascii_whitespace());
    let keyword_len = instruction
        .find(|value: char| value.is_ascii_whitespace())
        .unwrap_or(instruction.len());
    instruction[..keyword_len]
        .eq_ignore_ascii_case("PRINT")
        .then(|| instruction[keyword_len..].trim())
}

fn embed_field_instructions(instruction: &str) -> Option<&str> {
    if instruction.len() > MAX_INSTRUCTION_LEN {
        return None;
    }
    let instruction = instruction.trim_start_matches(|value: char| value.is_ascii_whitespace());
    let keyword = instruction.get(.."EMBED".len())?;
    if !keyword.eq_ignore_ascii_case("EMBED") {
        return None;
    }
    let remainder = instruction.get("EMBED".len()..)?;
    match remainder.chars().next() {
        None | Some('"') | Some('\\') => Some(remainder.trim()),
        Some(value) if value.is_ascii_whitespace() => Some(remainder.trim()),
        Some(_) => None,
    }
}

fn barcode_field_instructions(instruction: &str) -> Option<&str> {
    if instruction.len() > MAX_INSTRUCTION_LEN {
        return None;
    }
    let instruction = instruction.trim_start_matches(|value: char| value.is_ascii_whitespace());
    let keyword = instruction.get(.."BARCODE".len())?;
    if !keyword.eq_ignore_ascii_case("BARCODE") {
        return None;
    }
    let remainder = instruction.get("BARCODE".len()..)?;
    match remainder.chars().next() {
        None | Some('"') | Some('\\') => Some(remainder.trim()),
        Some(value) if value.is_ascii_whitespace() => Some(remainder.trim()),
        Some(_) => None,
    }
}

fn bidi_outline_field_instructions(instruction: &str) -> Option<&str> {
    if instruction.len() > MAX_INSTRUCTION_LEN {
        return None;
    }
    let instruction = instruction.trim_start_matches(|value: char| value.is_ascii_whitespace());
    let keyword = instruction.get(.."BIDIOUTLINE".len())?;
    if !keyword.eq_ignore_ascii_case("BIDIOUTLINE") {
        return None;
    }
    let remainder = instruction.get("BIDIOUTLINE".len()..)?;
    match remainder.chars().next() {
        None | Some('"') | Some('\\') => Some(remainder.trim()),
        Some(value) if value.is_ascii_whitespace() => Some(remainder.trim()),
        Some(_) => None,
    }
}

fn shape_field_instructions(instruction: &str) -> Option<&str> {
    if instruction.len() > MAX_INSTRUCTION_LEN {
        return None;
    }
    let instruction = instruction.trim_start_matches(|value: char| value.is_ascii_whitespace());
    let keyword = instruction.get(.."SHAPE".len())?;
    if !keyword.eq_ignore_ascii_case("SHAPE") {
        return None;
    }
    let remainder = instruction.get("SHAPE".len()..)?;
    match remainder.chars().next() {
        None | Some('"') | Some('\\') => Some(remainder.trim()),
        Some(value) if value.is_ascii_whitespace() => Some(remainder.trim()),
        Some(_) => None,
    }
}

fn legacy_form_field_instructions<'a>(
    instruction: &'a str,
    expected_keyword: &str,
) -> Option<&'a str> {
    if instruction.len() > MAX_INSTRUCTION_LEN {
        return None;
    }
    let instruction = instruction.trim_start_matches(|value: char| value.is_ascii_whitespace());
    let keyword = instruction.get(..expected_keyword.len())?;
    if !keyword.eq_ignore_ascii_case(expected_keyword) {
        return None;
    }
    let remainder = instruction.get(expected_keyword.len()..)?;
    match remainder.chars().next() {
        None | Some('"') | Some('\\') => Some(remainder.trim()),
        Some(value) if value.is_ascii_whitespace() => Some(remainder.trim()),
        Some(_) => None,
    }
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

fn go_to_button_parts(instruction: &str) -> Option<(Cow<'_, str>, Cow<'_, str>)> {
    let tokens = tokenize(instruction).ok()?;
    if tokens.len() != 3 || !tokens[0].value.eq_ignore_ascii_case("GOTOBUTTON") {
        return None;
    }
    let target = tokens[1].value.clone();
    let button_text = tokens[2].value.clone();
    if target.is_empty()
        || button_text.is_empty()
        || switch_name(&tokens[1]).is_some()
        || switch_name(&tokens[2]).is_some()
    {
        return None;
    }
    Some((target, button_text))
}

fn auto_text_field_parts<'a>(
    instruction: &'a str,
) -> Option<(AutoTextFieldKind, Cow<'a, str>, Vec<FieldSwitch<'a>>)> {
    let tokens = tokenize(instruction).ok()?;
    let keyword = tokens.first()?;
    let kind = if keyword.value.eq_ignore_ascii_case("GLOSSARY") {
        AutoTextFieldKind::Glossary
    } else if keyword.value.eq_ignore_ascii_case("AUTOTEXT") {
        AutoTextFieldKind::AutoText
    } else {
        return None;
    };
    let entry_name = tokens.get(1)?.value.clone();
    if entry_name.is_empty() || switch_name(tokens.get(1)?).is_some() {
        return None;
    }

    let mut unknown_switches = Vec::new();
    let mut index = 2;
    while index < tokens.len() {
        let name = switch_name(&tokens[index])?;
        let value = tokens
            .get(index + 1)
            .filter(|next| switch_name(next).is_none());
        unknown_switches.push(FieldSwitch {
            name: Cow::Owned(name.to_string()),
            value: value.map(|token| token.value.clone()),
        });
        index += 1 + usize::from(value.is_some());
    }
    Some((kind, entry_name, unknown_switches))
}

fn auto_text_list_field_parts<'a>(
    instruction: &'a str,
) -> Option<(
    Option<Cow<'a, str>>,
    Vec<AutoTextListOption<'a>>,
    Vec<FieldSwitch<'a>>,
)> {
    let tokens = tokenize(instruction).ok()?;
    let keyword = tokens.first()?;
    if !keyword.value.eq_ignore_ascii_case("AUTOTEXTLIST") {
        return None;
    }

    let mut index = 1;
    let display_text = tokens
        .get(index)
        .filter(|token| is_field_operand(token))
        .map(|token| {
            index += 1;
            token.value.clone()
        });
    let mut options = Vec::new();
    let mut unknown_switches = Vec::new();
    while index < tokens.len() {
        let name = switch_name(&tokens[index])?;
        let value = tokens.get(index + 1).filter(|next| is_field_operand(next));
        match name.to_ascii_lowercase().as_str() {
            "s" => {
                options.push(AutoTextListOption::Style(value?.value.clone()));
                index += 2;
            },
            "t" => {
                options.push(AutoTextListOption::Tip(value?.value.clone()));
                index += 2;
            },
            _ => {
                unknown_switches.push(FieldSwitch {
                    name: Cow::Owned(name.to_string()),
                    value: value.map(|token| token.value.clone()),
                });
                index += 1 + usize::from(value.is_some());
            },
        }
    }
    Some((display_text, options, unknown_switches))
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

fn index_parts(instruction: &str) -> Option<IndexParts<'_>> {
    let mut tokens = tokenize(instruction).ok()?;
    let keyword = tokens.first()?;
    if !keyword.value.eq_ignore_ascii_case("INDEX") {
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
                options.push(IndexOption::Bookmark(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "c" => {
                options.push(IndexOption::Columns(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "d" => {
                options.push(IndexOption::SequencePageSeparator(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "e" => {
                options.push(IndexOption::EntryPageNumberSeparator(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "f" => {
                options.push(IndexOption::EntryType(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "g" => {
                options.push(IndexOption::PageRangeSeparator(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "h" => {
                options.push(IndexOption::Heading(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "k" => {
                options.push(IndexOption::CrossReferenceSeparator(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "l" => {
                options.push(IndexOption::PageNumberSeparator(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "p" => {
                options.push(IndexOption::LetterRange(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "r" => {
                if tokens
                    .get(index + 1)
                    .is_some_and(|token| switch_name(token).is_none())
                {
                    return None;
                }
                options.push(IndexOption::RunIn);
                index += 1;
            },
            "s" => {
                options.push(IndexOption::SequenceIdentifier(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "y" => {
                if tokens
                    .get(index + 1)
                    .is_some_and(|token| switch_name(token).is_none())
                {
                    return None;
                }
                options.push(IndexOption::UseYomi);
                index += 1;
            },
            "z" => {
                options.push(IndexOption::LanguageId(
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

    Some(IndexParts {
        options,
        unknown_switches,
    })
}

fn index_entry_parts(instruction: &str) -> Option<IndexEntryParts<'_>> {
    let mut tokens = tokenize(instruction).ok()?;
    let keyword = tokens.first()?;
    if !keyword.value.eq_ignore_ascii_case("XE") {
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
            "b" => {
                if tokens
                    .get(index + 1)
                    .is_some_and(|token| switch_name(token).is_none())
                {
                    return None;
                }
                options.push(IndexEntryOption::BoldPageNumber);
                index += 1;
            },
            "f" => {
                options.push(IndexEntryOption::EntryType(
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
                options.push(IndexEntryOption::ItalicPageNumber);
                index += 1;
            },
            "r" => {
                options.push(IndexEntryOption::PageRangeBookmark(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "t" => {
                options.push(IndexEntryOption::CrossReference(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "y" => {
                options.push(IndexEntryOption::Yomi(
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

    Some(IndexEntryParts {
        entry,
        options,
        unknown_switches,
    })
}

fn citation_parts(instruction: &str) -> Option<CitationParts<'_>> {
    let mut tokens = tokenize(instruction).ok()?;
    let keyword = tokens.first()?;
    if !keyword.value.eq_ignore_ascii_case("CITATION") {
        return None;
    }
    tokens.remove(0);

    let source_tag = tokens.first()?.value.clone();
    if source_tag.is_empty() || switch_name(tokens.first()?).is_some() {
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
            "l" => {
                options.push(CitationOption::LanguageId(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "f" => {
                options.push(CitationOption::Prefix(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "s" => {
                options.push(CitationOption::Suffix(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "p" => {
                options.push(CitationOption::PageNumber(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "v" => {
                options.push(CitationOption::VolumeNumber(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "n" | "t" | "y" => {
                if tokens
                    .get(index + 1)
                    .is_some_and(|token| switch_name(token).is_none())
                {
                    return None;
                }
                options.push(match normalized_name.as_str() {
                    "n" => CitationOption::SuppressAuthor,
                    "t" => CitationOption::SuppressTitle,
                    "y" => CitationOption::SuppressYear,
                    _ => unreachable!("CITATION suppression switch was matched above"),
                });
                index += 1;
            },
            "m" => {
                options.push(CitationOption::AdditionalSourceTag(
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

    Some(CitationParts {
        source_tag,
        options,
        unknown_switches,
    })
}

fn bibliography_parts(instruction: &str) -> Option<BibliographyParts<'_>> {
    let mut tokens = tokenize(instruction).ok()?;
    let keyword = tokens.first()?;
    if !keyword.value.eq_ignore_ascii_case("BIBLIOGRAPHY") {
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
            "l" => {
                options.push(BibliographyOption::LanguageId(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "f" => {
                options.push(BibliographyOption::FilterLanguageId(
                    switch_value(&tokens, index, name).ok()?,
                ));
                index += 2;
            },
            "m" => {
                options.push(BibliographyOption::SourceTag(
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

    Some(BibliographyParts {
        options,
        unknown_switches,
    })
}

fn document_variable_field_parts(instruction: &str) -> Option<DocumentVariableFieldParts<'_>> {
    let mut tokens = tokenize(instruction).ok()?;
    let keyword = tokens.first()?;
    if !keyword.value.eq_ignore_ascii_case("DOCVARIABLE") {
        return None;
    }
    tokens.remove(0);

    let variable_name = tokens.first()?.value.clone();
    if variable_name.is_empty() || switch_name(tokens.first()?).is_some() {
        return None;
    }
    tokens.remove(0);

    let mut unknown_switches = Vec::new();
    let mut index = 0;
    while index < tokens.len() {
        let name = switch_name(&tokens[index])?;
        let value = tokens
            .get(index + 1)
            .filter(|token| switch_name(token).is_none());
        unknown_switches.push(FieldSwitch {
            name: Cow::Owned(name.to_string()),
            value: value.map(|token| token.value.clone()),
        });
        index += 1 + usize::from(value.is_some());
    }

    Some(DocumentVariableFieldParts {
        variable_name,
        unknown_switches,
    })
}

fn document_property_field_parts(instruction: &str) -> Option<DocumentPropertyFieldParts<'_>> {
    let mut tokens = tokenize(instruction).ok()?;
    let keyword = tokens.first()?;
    if !keyword.value.eq_ignore_ascii_case("DOCPROPERTY") {
        return None;
    }
    tokens.remove(0);

    let property_name = tokens.first()?.value.clone();
    if property_name.is_empty() || !is_field_operand(tokens.first()?) {
        return None;
    }
    tokens.remove(0);

    let mut switches = Vec::new();
    let mut index = 0;
    while index < tokens.len() {
        let name = switch_name(&tokens[index])?;
        let value = tokens
            .get(index + 1)
            .filter(|token| is_field_operand(token));
        switches.push(FieldSwitch {
            name: Cow::Owned(name.to_string()),
            value: value.map(|token| token.value.clone()),
        });
        index += 1 + usize::from(value.is_some());
    }

    Some(DocumentPropertyFieldParts {
        property_name,
        switches,
    })
}

fn info_field_parts(instruction: &str) -> Option<InfoFieldParts<'_>> {
    let mut tokens = tokenize(instruction).ok()?;
    let keyword = tokens.first()?;
    if !keyword.value.eq_ignore_ascii_case("INFO") {
        return None;
    }
    tokens.remove(0);

    let information_type = tokens.first()?.value.clone();
    if information_type.is_empty() || !is_field_operand(tokens.first()?) {
        return None;
    }
    tokens.remove(0);

    let new_value = tokens
        .first()
        .filter(|token| is_field_operand(token))
        .map(|token| token.value.clone());
    if new_value.is_some() {
        tokens.remove(0);
    }

    let mut switches = Vec::new();
    let mut index = 0;
    while index < tokens.len() {
        let name = switch_name(&tokens[index])?;
        let value = tokens
            .get(index + 1)
            .filter(|token| is_field_operand(token));
        switches.push(FieldSwitch {
            name: Cow::Owned(name.to_string()),
            value: value.map(|token| token.value.clone()),
        });
        index += 1 + usize::from(value.is_some());
    }

    Some(InfoFieldParts {
        information_type,
        new_value,
        switches,
    })
}

fn document_information_field_parts(
    instruction: &str,
) -> Option<DocumentInformationFieldParts<'_>> {
    let mut tokens = tokenize(instruction).ok()?;
    let keyword = tokens.first()?;
    let kind = DocumentInformationFieldKind::from_keyword(keyword.value.as_ref())?;
    tokens.remove(0);

    let mut switches = Vec::new();
    let mut index = 0;
    while index < tokens.len() {
        let name = switch_name(&tokens[index])?;
        let value = tokens
            .get(index + 1)
            .filter(|token| is_field_operand(token));
        switches.push(FieldSwitch {
            name: Cow::Owned(name.to_string()),
            value: value.map(|token| token.value.clone()),
        });
        index += 1 + usize::from(value.is_some());
    }

    Some(DocumentInformationFieldParts { kind, switches })
}

fn document_context_field_parts(instruction: &str) -> Option<DocumentContextFieldParts<'_>> {
    let mut tokens = tokenize(instruction).ok()?;
    let keyword = tokens.first()?;
    let kind = DocumentContextFieldKind::from_keyword(keyword.value.as_ref())?;
    tokens.remove(0);

    let mut switches = Vec::new();
    let mut index = 0;
    while index < tokens.len() {
        let name = switch_name(&tokens[index])?;
        let value = tokens
            .get(index + 1)
            .filter(|token| is_field_operand(token));
        switches.push(FieldSwitch {
            name: Cow::Owned(name.to_string()),
            value: value.map(|token| token.value.clone()),
        });
        index += 1 + usize::from(value.is_some());
    }

    Some(DocumentContextFieldParts { kind, switches })
}

fn merge_field_parts(instruction: &str) -> Option<MergeFieldParts<'_>> {
    let mut tokens = tokenize(instruction).ok()?;
    let keyword = tokens.first()?;
    if !keyword.value.eq_ignore_ascii_case("MERGEFIELD") {
        return None;
    }
    tokens.remove(0);

    let field_name = tokens.first()?.value.clone();
    if field_name.is_empty() || switch_name(tokens.first()?).is_some() {
        return None;
    }
    tokens.remove(0);

    let mut switches = Vec::new();
    let mut index = 0;
    while index < tokens.len() {
        let name = switch_name(&tokens[index])?;
        let value = tokens
            .get(index + 1)
            .filter(|token| switch_name(token).is_none());
        switches.push(FieldSwitch {
            name: Cow::Owned(name.to_string()),
            value: value.map(|token| token.value.clone()),
        });
        index += 1 + usize::from(value.is_some());
    }

    Some(MergeFieldParts {
        field_name,
        switches,
    })
}

fn mail_merge_data_field_parts(instruction: &str) -> Option<MailMergeDataFieldParts<'_>> {
    let mut tokens = tokenize(instruction).ok()?;
    let keyword = tokens.first()?;
    if !keyword.value.eq_ignore_ascii_case("DATA") {
        return None;
    }
    tokens.remove(0);

    let data_source = tokens.first()?.value.clone();
    if data_source.is_empty() || switch_name(tokens.first()?).is_some() {
        return None;
    }
    tokens.remove(0);

    let header_source = match tokens.first() {
        Some(token) if switch_name(token).is_none() && !token.value.is_empty() => {
            let header_source = token.value.clone();
            tokens.remove(0);
            Some(header_source)
        },
        Some(token) if switch_name(token).is_none() => return None,
        _ => None,
    };

    let mut switches = Vec::new();
    let mut index = 0;
    while index < tokens.len() {
        let name = switch_name(&tokens[index])?;
        let value = tokens
            .get(index + 1)
            .filter(|token| switch_name(token).is_none());
        switches.push(FieldSwitch {
            name: Cow::Owned(name.to_string()),
            value: value.map(|token| token.value.clone()),
        });
        index += 1 + usize::from(value.is_some());
    }

    Some(MailMergeDataFieldParts {
        data_source,
        header_source,
        switches,
    })
}

fn prompt_field_parts(instruction: &str) -> Option<PromptFieldParts<'_>> {
    let mut tokens = tokenize(instruction).ok()?;
    let keyword = tokens.first()?;
    let kind = if keyword.value.eq_ignore_ascii_case("ASK") {
        PromptFieldKind::Ask
    } else if keyword.value.eq_ignore_ascii_case("FILLIN") {
        PromptFieldKind::FillIn
    } else {
        return None;
    };
    tokens.remove(0);

    let (bookmark, prompt) = match kind {
        PromptFieldKind::Ask => {
            let bookmark = tokens.first()?;
            if bookmark.value.is_empty() || switch_name(bookmark).is_some() {
                return None;
            }
            let bookmark = tokens.remove(0).value;

            let prompt = tokens.first()?;
            if switch_name(prompt).is_some() {
                return None;
            }
            let prompt = tokens.remove(0).value;
            (Some(bookmark), Some(prompt))
        },
        PromptFieldKind::FillIn => {
            let prompt = if tokens
                .first()
                .is_some_and(|token| switch_name(token).is_none())
            {
                Some(tokens.remove(0).value)
            } else {
                None
            };
            (None, prompt)
        },
    };

    let mut default_response = None;
    let mut prompts_once_per_mail_merge = false;
    let mut index = 0;
    while index < tokens.len() {
        let name = switch_name(&tokens[index])?;
        match name.to_ascii_lowercase().as_str() {
            "d" => {
                if default_response.is_some() {
                    return None;
                }
                let value = tokens
                    .get(index + 1)
                    .filter(|token| switch_name(token).is_none())?;
                default_response = Some(value.value.clone());
                index += 2;
            },
            "o" => {
                if prompts_once_per_mail_merge
                    || tokens
                        .get(index + 1)
                        .is_some_and(|token| switch_name(token).is_none())
                {
                    return None;
                }
                prompts_once_per_mail_merge = true;
                index += 1;
            },
            _ => return None,
        }
    }

    Some(PromptFieldParts {
        kind,
        bookmark,
        prompt,
        default_response,
        prompts_once_per_mail_merge,
    })
}

fn user_identity_field_parts(instruction: &str) -> Option<UserIdentityFieldParts<'_>> {
    let mut tokens = tokenize(instruction).ok()?;
    let keyword = tokens.first()?;
    let kind = if keyword.value.eq_ignore_ascii_case("USERADDRESS") {
        UserIdentityFieldKind::Address
    } else if keyword.value.eq_ignore_ascii_case("USERINITIALS") {
        UserIdentityFieldKind::Initials
    } else if keyword.value.eq_ignore_ascii_case("USERNAME") {
        UserIdentityFieldKind::Name
    } else {
        return None;
    };
    tokens.remove(0);

    let override_value = if tokens
        .first()
        .is_some_and(|token| switch_name(token).is_none())
    {
        Some(tokens.remove(0).value)
    } else {
        None
    };

    let mut formatting = None;
    let mut index = 0;
    while index < tokens.len() {
        let name = switch_name(&tokens[index])?;
        if name != "*" || formatting.is_some() {
            return None;
        }
        let value = tokens
            .get(index + 1)
            .filter(|token| switch_name(token).is_none())?
            .value
            .as_ref();
        formatting = Some(if value.eq_ignore_ascii_case("Caps") {
            UserIdentityFormatting::Caps
        } else if value.eq_ignore_ascii_case("FirstCap") {
            UserIdentityFormatting::FirstCap
        } else if value.eq_ignore_ascii_case("Lower") {
            UserIdentityFormatting::Lower
        } else if value.eq_ignore_ascii_case("Upper") {
            UserIdentityFormatting::Upper
        } else {
            return None;
        });
        index += 2;
    }

    Some(UserIdentityFieldParts {
        kind,
        override_value,
        formatting,
    })
}

fn advance_field_adjustments(instruction: &str) -> Option<Vec<AdvanceFieldAdjustment>> {
    let mut tokens = tokenize(instruction).ok()?;
    let keyword = tokens.first()?;
    if !keyword.value.eq_ignore_ascii_case("ADVANCE") {
        return None;
    }
    tokens.remove(0);

    let mut adjustments = Vec::with_capacity(tokens.len() / 2);
    let mut index = 0;
    while index < tokens.len() {
        let name = switch_name(&tokens[index])?;
        let operation = match name.to_ascii_lowercase().as_str() {
            "d" => AdvanceFieldOperation::Down,
            "l" => AdvanceFieldOperation::Left,
            "r" => AdvanceFieldOperation::Right,
            "u" => AdvanceFieldOperation::Up,
            "x" => AdvanceFieldOperation::HorizontalPosition,
            "y" => AdvanceFieldOperation::VerticalPosition,
            _ => return None,
        };
        let points = tokens
            .get(index + 1)
            .filter(|token| switch_name(token).is_none())?
            .value
            .parse::<i64>()
            .ok()?;
        adjustments.push(AdvanceFieldAdjustment { operation, points });
        index += 2;
    }

    Some(adjustments)
}

fn mail_merge_recipient_field_parts(instruction: &str) -> Option<MailMergeRecipientFieldParts<'_>> {
    let mut tokens = tokenize(instruction).ok()?;
    let keyword = tokens.first()?;
    let kind = if keyword.value.eq_ignore_ascii_case("ADDRESSBLOCK") {
        MailMergeRecipientFieldKind::AddressBlock
    } else if keyword.value.eq_ignore_ascii_case("GREETINGLINE") {
        MailMergeRecipientFieldKind::GreetingLine
    } else {
        return None;
    };
    tokens.remove(0);

    let mut country_inclusion = None;
    let mut formats_using_recipient_country = false;
    let mut excluded_countries = Vec::new();
    let mut format_template = None;
    let mut language = None;
    let mut greeting_fallback_text = None;
    let mut unknown_switches = Vec::new();
    let mut index = 0;
    while index < tokens.len() {
        let name = switch_name(&tokens[index])?;
        let normalized = name.to_ascii_lowercase();
        let value = tokens
            .get(index + 1)
            .filter(|token| switch_name(token).is_none());

        match (kind, normalized.as_str()) {
            (MailMergeRecipientFieldKind::AddressBlock, "c") => {
                if country_inclusion.is_some() {
                    return None;
                }
                let value = value?.value.clone();
                country_inclusion = Some(match value.as_ref() {
                    "0" => AddressBlockCountryInclusion::Omit,
                    "1" => AddressBlockCountryInclusion::Always,
                    "2" => AddressBlockCountryInclusion::UnlessExcluded,
                    _ => return None,
                });
                index += 2;
            },
            (MailMergeRecipientFieldKind::AddressBlock, "d") => {
                if formats_using_recipient_country || value.is_some() {
                    return None;
                }
                formats_using_recipient_country = true;
                index += 1;
            },
            (MailMergeRecipientFieldKind::AddressBlock, "e") => {
                excluded_countries.push(value?.value.clone());
                index += 2;
            },
            (_, "f") => {
                if format_template.is_some() {
                    return None;
                }
                format_template = Some(value?.value.clone());
                index += 2;
            },
            (_, "l") => {
                if language.is_some() {
                    return None;
                }
                language = Some(value?.value.clone());
                index += 2;
            },
            (MailMergeRecipientFieldKind::GreetingLine, "c" | "e") => {
                if greeting_fallback_text.is_some() {
                    return None;
                }
                greeting_fallback_text = Some(value?.value.clone());
                index += 2;
            },
            _ => {
                unknown_switches.push(FieldSwitch {
                    name: Cow::Owned(name.to_string()),
                    value: value.map(|token| token.value.clone()),
                });
                index += 1 + usize::from(value.is_some());
            },
        }
    }

    Some(MailMergeRecipientFieldParts {
        kind,
        country_inclusion,
        formats_using_recipient_country,
        excluded_countries,
        format_template,
        language,
        greeting_fallback_text,
        unknown_switches,
    })
}

fn mail_merge_counter_kind(instruction: &str) -> Option<MailMergeCounterKind> {
    let tokens = tokenize(instruction).ok()?;
    if tokens.len() != 1 {
        return None;
    }

    if tokens[0].value.eq_ignore_ascii_case("MERGEREC") {
        Some(MailMergeCounterKind::Record)
    } else if tokens[0].value.eq_ignore_ascii_case("MERGESEQ") {
        Some(MailMergeCounterKind::Sequence)
    } else {
        None
    }
}

fn is_mail_merge_next_instruction(instruction: &str) -> bool {
    let Ok(tokens) = tokenize(instruction) else {
        return false;
    };
    tokens.len() == 1 && tokens[0].value.eq_ignore_ascii_case("NEXT")
}

fn mail_merge_conditional_control_parts(
    instruction: &str,
) -> Option<(MailMergeConditionalControlKind, &str)> {
    tokenize(instruction).ok()?;
    let instruction =
        instruction.trim_start_matches(|character: char| character.is_ascii_whitespace());
    let (kind, keyword) = if instruction
        .get(..6)
        .is_some_and(|candidate| candidate.eq_ignore_ascii_case("NEXTIF"))
    {
        (MailMergeConditionalControlKind::NextIf, "NEXTIF")
    } else if instruction
        .get(..6)
        .is_some_and(|candidate| candidate.eq_ignore_ascii_case("SKIPIF"))
    {
        (MailMergeConditionalControlKind::SkipIf, "SKIPIF")
    } else {
        return None;
    };
    let remainder = instruction.get(keyword.len()..)?;
    if !remainder
        .chars()
        .next()
        .is_some_and(|character| character.is_ascii_whitespace())
    {
        return None;
    }
    let comparison = remainder.trim_matches(|character: char| character.is_ascii_whitespace());
    (!comparison.is_empty()).then_some((kind, comparison))
}

fn comparison_field_expression<'a>(instruction: &'a str, field_type: &str) -> Option<&'a str> {
    tokenize(instruction).ok()?;
    let instruction =
        instruction.trim_start_matches(|character: char| character.is_ascii_whitespace());
    let candidate = instruction.get(..field_type.len())?;
    if !candidate.eq_ignore_ascii_case(field_type) {
        return None;
    }
    let remainder = instruction.get(field_type.len()..)?;
    if !remainder
        .chars()
        .next()
        .is_some_and(|character| character.is_ascii_whitespace())
    {
        return None;
    }
    let comparison = remainder.trim_matches(|character: char| character.is_ascii_whitespace());
    (!comparison.is_empty()).then_some(comparison)
}

fn if_field_expression(instruction: &str) -> Option<&str> {
    comparison_field_expression(instruction, "IF")
}

fn compare_field_comparison(instruction: &str) -> Option<&str> {
    comparison_field_expression(instruction, "COMPARE")
}

fn set_field_parts<'a>(instruction: &'a str) -> Option<(Cow<'a, str>, &'a str)> {
    let tokens = tokenize(instruction).ok()?;
    let keyword = tokens.first()?;
    if !keyword.value.eq_ignore_ascii_case("SET") {
        return None;
    }
    let target_name = tokens.get(1)?.value.clone();
    if target_name.is_empty() || switch_name(tokens.get(1)?).is_some() {
        return None;
    }

    let instruction =
        instruction.trim_start_matches(|character: char| character.is_ascii_whitespace());
    let remainder = instruction.get("SET".len()..)?;
    if !remainder
        .chars()
        .next()
        .is_some_and(|character| character.is_ascii_whitespace())
    {
        return None;
    }
    let target_start =
        remainder.trim_start_matches(|character: char| character.is_ascii_whitespace());
    let target_end = field_argument_end(target_start)?;
    let expression = target_start.get(target_end..)?;
    if !expression
        .chars()
        .next()
        .is_some_and(|character| character.is_ascii_whitespace())
    {
        return None;
    }
    let expression = expression.trim_matches(|character: char| character.is_ascii_whitespace());
    (!expression.is_empty()).then_some((target_name, expression))
}

fn field_argument_end(input: &str) -> Option<usize> {
    let bytes = input.as_bytes();
    if bytes.is_empty() {
        return None;
    }
    if bytes[0] != b'"' {
        return Some(
            bytes
                .iter()
                .position(u8::is_ascii_whitespace)
                .unwrap_or(bytes.len()),
        );
    }

    let mut index = 1;
    while index < bytes.len() {
        match bytes[index] {
            b'"' => return Some(index + 1),
            b'\\' if index + 1 < bytes.len() && matches!(bytes[index + 1], b'\\' | b'"') => {
                index += 2;
            },
            _ => {
                let character = input[index..].chars().next()?;
                index += character.len_utf8();
            },
        }
    }
    None
}

fn sequence_field_parts<'a>(
    instruction: &'a str,
) -> Option<(Cow<'a, str>, Option<Cow<'a, str>>, &'a str)> {
    let tokens = tokenize(instruction).ok()?;
    let keyword = tokens.first()?;
    if !keyword.value.eq_ignore_ascii_case("SEQ") {
        return None;
    }
    let identifier = tokens.get(1)?.value.clone();
    if identifier.is_empty() || switch_name(tokens.get(1)?).is_some() {
        return None;
    }

    let instruction =
        instruction.trim_start_matches(|character: char| character.is_ascii_whitespace());
    let remainder = instruction.get("SEQ".len()..)?;
    if !remainder
        .chars()
        .next()
        .is_some_and(|character| character.is_ascii_whitespace())
    {
        return None;
    }
    let identifier_start =
        remainder.trim_start_matches(|character: char| character.is_ascii_whitespace());
    let identifier_end = field_argument_end(identifier_start)?;
    let after_identifier = identifier_start.get(identifier_end..)?;
    if after_identifier.is_empty() {
        return Some((identifier, None, ""));
    }
    if !after_identifier
        .chars()
        .next()
        .is_some_and(|character| character.is_ascii_whitespace())
    {
        return None;
    }
    let after_identifier =
        after_identifier.trim_start_matches(|character: char| character.is_ascii_whitespace());
    if after_identifier.is_empty() || after_identifier.starts_with('\\') {
        return Some((
            identifier,
            None,
            after_identifier.trim_matches(|character: char| character.is_ascii_whitespace()),
        ));
    }

    let bookmark = tokens.get(2)?.value.clone();
    if bookmark.is_empty() || switch_name(tokens.get(2)?).is_some() {
        return None;
    }
    let bookmark_end = field_argument_end(after_identifier)?;
    let after_bookmark = after_identifier.get(bookmark_end..)?;
    if !after_bookmark.is_empty()
        && !after_bookmark
            .chars()
            .next()
            .is_some_and(|character| character.is_ascii_whitespace())
    {
        return None;
    }
    let tail =
        after_bookmark.trim_matches(|character: char| character.is_ascii_whitespace());
    Some((identifier, Some(bookmark), tail))
}

fn is_formula_field_instruction(instruction: &str) -> bool {
    tokenize(instruction).is_ok() && instruction.trim_start().starts_with('=')
}

fn formula_field_formula(instruction: &str) -> Option<&str> {
    tokenize(instruction).ok()?;
    let formula = instruction.trim().strip_prefix('=')?.trim();
    (!formula.is_empty()).then_some(formula)
}

fn quote_field_parts<'a>(
    instruction: &'a str,
) -> Option<(Cow<'a, str>, Vec<FieldSwitch<'a>>)> {
    let tokens = tokenize(instruction).ok()?;
    let keyword = tokens.first()?;
    if !keyword.value.eq_ignore_ascii_case("QUOTE") {
        return None;
    }
    let text = tokens.get(1)?.value.clone();
    if switch_name(tokens.get(1)?).is_some() {
        return None;
    }

    let mut switches = Vec::new();
    let mut index = 2;
    while index < tokens.len() {
        let name = switch_name(&tokens[index])?;
        let value = tokens
            .get(index + 1)
            .filter(|next| switch_name(next).is_none());
        switches.push(FieldSwitch {
            name: Cow::Owned(name.to_ascii_lowercase()),
            value: value.map(|token| token.value.clone()),
        });
        index += 1 + usize::from(value.is_some());
    }
    Some((text, switches))
}

fn symbol_field_parts<'a>(
    instruction: &'a str,
) -> Option<(Cow<'a, str>, Vec<FieldSwitch<'a>>)> {
    let tokens = tokenize(instruction).ok()?;
    let keyword = tokens.first()?;
    if !keyword.value.eq_ignore_ascii_case("SYMBOL") {
        return None;
    }
    let character_argument = tokens.get(1)?.value.clone();
    if switch_name(tokens.get(1)?).is_some() {
        return None;
    }

    let mut switches = Vec::new();
    let mut index = 2;
    while index < tokens.len() {
        let name = switch_name(&tokens[index])?;
        let value = tokens
            .get(index + 1)
            .filter(|next| switch_name(next).is_none());
        switches.push(FieldSwitch {
            name: Cow::Owned(name.to_ascii_lowercase()),
            value: value.map(|token| token.value.clone()),
        });
        index += 1 + usize::from(value.is_some());
    }
    Some((character_argument, switches))
}

fn auto_number_field_parts(instruction: &str) -> Option<AutoNumberFieldParts<'_>> {
    let mut tokens = tokenize(instruction).ok()?;
    let keyword = tokens.first()?;
    let kind = AutoNumberFieldKind::from_keyword(keyword.value.as_ref())?;
    tokens.remove(0);

    let mut switches = Vec::new();
    let mut index = 0;
    while index < tokens.len() {
        let name = switch_name(&tokens[index])?;
        let value = tokens
            .get(index + 1)
            .filter(|token| is_field_operand(token));
        switches.push(FieldSwitch {
            name: Cow::Owned(name.to_string()),
            value: value.map(|token| token.value.clone()),
        });
        index += 1 + usize::from(value.is_some());
    }

    Some(AutoNumberFieldParts { kind, switches })
}

fn list_number_field_parts(instruction: &str) -> Option<ListNumberFieldParts<'_>> {
    let mut tokens = tokenize(instruction).ok()?;
    let keyword = tokens.first()?;
    if !keyword.value.eq_ignore_ascii_case("LISTNUM") {
        return None;
    }
    tokens.remove(0);

    let list_name = tokens
        .first()
        .filter(|token| is_field_operand(token))
        .map(|token| token.value.clone());
    if list_name.is_some() {
        tokens.remove(0);
    }

    let mut switches = Vec::new();
    let mut index = 0;
    while index < tokens.len() {
        let name = switch_name(&tokens[index])?;
        let value = tokens
            .get(index + 1)
            .filter(|token| is_field_operand(token));
        switches.push(FieldSwitch {
            name: Cow::Owned(name.to_string()),
            value: value.map(|token| token.value.clone()),
        });
        index += 1 + usize::from(value.is_some());
    }

    Some(ListNumberFieldParts {
        list_name,
        switches,
    })
}

fn style_reference_field_parts<'a>(
    instruction: &'a str,
) -> Option<(
    Cow<'a, str>,
    Vec<StyleReferenceFieldOption>,
    Vec<FieldSwitch<'a>>,
)> {
    let tokens = tokenize(instruction).ok()?;
    let keyword = tokens.first()?;
    if !keyword.value.eq_ignore_ascii_case("STYLEREF") {
        return None;
    }
    let style_name = tokens.get(1)?.value.clone();
    if style_name.is_empty() || switch_name(tokens.get(1)?).is_some() {
        return None;
    }

    let mut options = Vec::new();
    let mut unknown_switches = Vec::new();
    let mut index = 2;
    while index < tokens.len() {
        let token = &tokens[index];
        let name = switch_name(token)?;
        let option = match name.to_ascii_lowercase().as_str() {
            "l" => Some(StyleReferenceFieldOption::FollowingText),
            "n" => Some(StyleReferenceFieldOption::ParagraphNumber),
            "p" => Some(StyleReferenceFieldOption::RelativePosition),
            "r" => Some(StyleReferenceFieldOption::ParagraphNumberRelativeContext),
            "t" => Some(StyleReferenceFieldOption::SuppressNonNumberText),
            "w" => Some(StyleReferenceFieldOption::ParagraphNumberFullContext),
            _ => None,
        };
        if let Some(option) = option {
            if tokens
                .get(index + 1)
                .is_some_and(|next| switch_name(next).is_none())
            {
                return None;
            }
            options.push(option);
            index += 1;
            continue;
        }

        let value = tokens
            .get(index + 1)
            .filter(|next| switch_name(next).is_none());
        unknown_switches.push(FieldSwitch {
            name: Cow::Owned(name.to_string()),
            value: value.map(|token| token.value.clone()),
        });
        index += 1 + usize::from(value.is_some());
    }
    Some((style_name, options, unknown_switches))
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

fn is_field_operand(token: &FieldCodeToken<'_>) -> bool {
    switch_name(token).is_none() && (token.quoted || !token.value.starts_with('\\'))
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
    fn go_to_button_fields_expose_stored_metadata_without_navigation() {
        let mut field = Field::parse_instruction(r#"GOTOBUTTON "f 2" "Footnote""#);
        field.result = Cow::Borrowed("cached footnote button");
        field.status = FieldStatus {
            dirty: true,
            locked: true,
            ..FieldStatus::default()
        };
        field.owner = FieldOwner::Body;
        field.position = 4;

        assert_eq!(field.field_type, FieldType::GoToButton);
        let button = field.go_to_button().unwrap();
        assert_eq!(button.instruction(), r#"GOTOBUTTON "f 2" "Footnote""#);
        assert_eq!(button.target(), "f 2");
        assert_eq!(button.button_text(), "Footnote");
        assert_eq!(button.cached_result(), Some("cached footnote button"));
        assert!(button.is_dirty());
        assert!(button.is_locked());
        assert_eq!(button.owner(), FieldOwner::Body);
        assert_eq!(button.position(), 4);

        for instruction in [
            "GOTOBUTTON",
            r#"GOTOBUTTON "" Button"#,
            "GOTOBUTTON Destination",
            r#"GOTOBUTTON Destination """#,
            "GOTOBUTTON Destination Button unexpected",
            r#"GOTOBUTTON Destination Button \* MERGEFORMAT"#,
        ] {
            assert!(
                Field::parse_instruction(instruction)
                    .go_to_button()
                    .is_none(),
                "{instruction}"
            );
        }
        assert_eq!(
            Field::parse_instruction("GOTOBUTTONS Destination Button").field_type,
            FieldType::Unknown
        );
    }

    #[test]
    fn active_content_fields_expose_opaque_metadata_without_activation() {
        let mut add_in = Field::parse_instruction("ADDIN opaque-add-in-data");
        add_in.result = Cow::Borrowed("cached add-in result");
        add_in.status = FieldStatus {
            dirty: true,
            locked: true,
            ..FieldStatus::default()
        };
        add_in.owner = FieldOwner::Body;
        add_in.position = 4;

        assert_eq!(add_in.field_type, FieldType::AddIn);
        let add_in = add_in.active_content_field().unwrap();
        assert_eq!(add_in.instruction(), "ADDIN opaque-add-in-data");
        assert_eq!(add_in.kind(), ActiveContentFieldKind::AddIn);
        assert_eq!(add_in.cached_result(), Some("cached add-in result"));
        assert!(add_in.is_dirty());
        assert!(add_in.is_locked());
        assert_eq!(add_in.owner(), FieldOwner::Body);
        assert_eq!(add_in.position(), 4);

        let control = Field::parse_instruction("control opaque-ocx-metadata");
        assert_eq!(control.field_type, FieldType::Control);
        let control = control.active_content_field().unwrap();
        assert_eq!(control.kind(), ActiveContentFieldKind::OcxControl);
        assert_eq!(control.cached_result(), None);

        let html = Field::parse_instruction("HTMLCONTROL opaque-html-control-metadata");
        assert_eq!(html.field_type, FieldType::HtmlControl);
        let html = html.active_content_field().unwrap();
        assert_eq!(html.kind(), ActiveContentFieldKind::HtmlControl);

        assert_eq!(
            Field::parse_instruction("ADDINS not-an-add-in").field_type,
            FieldType::Unknown
        );
        assert!(
            Field::parse_instruction("MACROBUTTON NoMacro Button")
                .active_content_field()
                .is_none()
        );
    }

    #[test]
    fn print_fields_preserve_opaque_metadata_without_sending_printer_commands() {
        let mut printer = Field::parse_instruction(r#"PRINT "ESC&l1O""#);
        printer.result = Cow::Borrowed("cached printer result");
        printer.status = FieldStatus {
            dirty: true,
            locked: true,
            ..FieldStatus::default()
        };
        printer.owner = FieldOwner::Body;
        printer.position = 4;

        assert_eq!(printer.field_type, FieldType::Print);
        let printer = printer.print_field().unwrap();
        assert_eq!(printer.instruction(), r#"PRINT "ESC&l1O""#);
        assert_eq!(printer.printer_instructions(), r#""ESC&l1O""#);
        assert_eq!(printer.cached_result(), Some("cached printer result"));
        assert!(printer.is_dirty());
        assert!(printer.is_locked());
        assert_eq!(printer.owner(), FieldOwner::Body);
        assert_eq!(printer.position(), 4);

        let postscript = Field::parse_instruction(r#"print \p 2 "0 0 moveto""#);
        assert_eq!(postscript.field_type, FieldType::Print);
        let postscript = postscript.print_field().unwrap();
        assert_eq!(postscript.printer_instructions(), r#"\p 2 "0 0 moveto""#);

        assert_eq!(
            Field::parse_instruction(r#"PRINTS "not a print field""#).field_type,
            FieldType::Unknown
        );
        assert!(
            Field::parse_instruction("ADDIN opaque-metadata")
                .print_field()
                .is_none()
        );
        let too_long = Field::new(
            FieldType::Print,
            Cow::Owned(format!("PRINT {}", "x".repeat(MAX_INSTRUCTION_LEN))),
            Cow::Borrowed(""),
        );
        assert!(too_long.print_field().is_none());
    }

    #[test]
    fn embed_fields_preserve_opaque_metadata_without_loading_or_activating_objects() {
        let mut embedded = Field::parse_instruction(r#"EMBED Excel.Sheet.12 \* MERGEFORMAT"#);
        embedded.result = Cow::Borrowed("cached worksheet object");
        embedded.status = FieldStatus {
            dirty: true,
            locked: true,
            ..FieldStatus::default()
        };
        embedded.owner = FieldOwner::Body;
        embedded.position = 4;

        assert_eq!(embedded.field_type, FieldType::Embed);
        let embedded = embedded.embed_field().unwrap();
        assert_eq!(
            embedded.instruction(),
            r#"EMBED Excel.Sheet.12 \* MERGEFORMAT"#
        );
        assert_eq!(
            embedded.object_instructions(),
            r#"Excel.Sheet.12 \* MERGEFORMAT"#
        );
        assert_eq!(embedded.cached_result(), Some("cached worksheet object"));
        assert!(embedded.is_dirty());
        assert!(embedded.is_locked());
        assert_eq!(embedded.owner(), FieldOwner::Body);
        assert_eq!(embedded.position(), 4);

        let equation = Field::parse_instruction(r#"embed "Equation.DSMT4" \d"#);
        assert_eq!(equation.field_type, FieldType::Embed);
        assert_eq!(
            equation.embed_field().unwrap().object_instructions(),
            r#""Equation.DSMT4" \d"#
        );

        let bare = Field::parse_instruction("EMBED");
        assert_eq!(bare.embed_field().unwrap().object_instructions(), "");
        assert_eq!(
            Field::parse_instruction("EMBEDS Excel.Sheet.12").field_type,
            FieldType::Unknown
        );
        assert!(
            Field::parse_instruction("EMBEDS Excel.Sheet.12")
                .embed_field()
                .is_none()
        );
        let too_long = Field::new(
            FieldType::Embed,
            Cow::Owned(format!("EMBED {}", "x".repeat(MAX_INSTRUCTION_LEN))),
            Cow::Borrowed(""),
        );
        assert!(too_long.embed_field().is_none());
    }

    #[test]
    fn barcode_fields_preserve_opaque_metadata_without_decoding_or_rendering() {
        let mut barcode =
            Field::parse_instruction(r#"BARCODE "4901234567894" EAN13 \h 1440 \* MERGEFORMAT"#);
        barcode.result = Cow::Borrowed("cached barcode");
        barcode.status = FieldStatus {
            dirty: true,
            locked: true,
            ..FieldStatus::default()
        };
        barcode.owner = FieldOwner::Body;
        barcode.position = 4;

        assert_eq!(barcode.field_type, FieldType::Barcode);
        let barcode = barcode.barcode_field().unwrap();
        assert_eq!(
            barcode.instruction(),
            r#"BARCODE "4901234567894" EAN13 \h 1440 \* MERGEFORMAT"#
        );
        assert_eq!(
            barcode.barcode_instructions(),
            r#""4901234567894" EAN13 \h 1440 \* MERGEFORMAT"#
        );
        assert_eq!(barcode.cached_result(), Some("cached barcode"));
        assert!(barcode.is_dirty());
        assert!(barcode.is_locked());
        assert_eq!(barcode.owner(), FieldOwner::Body);
        assert_eq!(barcode.position(), 4);

        let code_39 = Field::parse_instruction(r#"barcode "ABC-123" CODE39 \d"#);
        assert_eq!(code_39.field_type, FieldType::Barcode);
        assert_eq!(
            code_39.barcode_field().unwrap().barcode_instructions(),
            r#""ABC-123" CODE39 \d"#
        );

        let bare = Field::parse_instruction("BARCODE");
        assert_eq!(bare.barcode_field().unwrap().barcode_instructions(), "");
        assert_eq!(
            Field::parse_instruction("BARCODES 4901234567894").field_type,
            FieldType::Unknown
        );
        assert!(
            Field::parse_instruction("BARCODES 4901234567894")
                .barcode_field()
                .is_none()
        );
        let too_long = Field::new(
            FieldType::Barcode,
            Cow::Owned(format!("BARCODE {}", "x".repeat(MAX_INSTRUCTION_LEN))),
            Cow::Borrowed(""),
        );
        assert!(too_long.barcode_field().is_none());
    }

    #[test]
    fn bidi_outline_fields_preserve_metadata_without_resolving_numbering_or_layout() {
        let mut outline = Field::parse_instruction(r#"BIDIOUTLINE \* MERGEFORMAT"#);
        outline.result = Cow::Borrowed("cached bidi outline number");
        outline.status = FieldStatus {
            dirty: true,
            locked: true,
            ..FieldStatus::default()
        };
        outline.owner = FieldOwner::Body;
        outline.position = 4;

        assert_eq!(outline.field_type, FieldType::BidiOutline);
        let outline = outline.bidi_outline_field().unwrap();
        assert_eq!(outline.instruction(), r#"BIDIOUTLINE \* MERGEFORMAT"#);
        assert_eq!(outline.opaque_instructions(), r#"\* MERGEFORMAT"#);
        assert_eq!(outline.cached_result(), Some("cached bidi outline number"));
        assert!(outline.is_dirty());
        assert!(outline.is_locked());
        assert_eq!(outline.owner(), FieldOwner::Body);
        assert_eq!(outline.position(), 4);

        let bare = Field::parse_instruction("bidioutline");
        assert_eq!(bare.field_type, FieldType::BidiOutline);
        assert_eq!(bare.bidi_outline_field().unwrap().opaque_instructions(), "");
        assert_eq!(
            Field::parse_instruction("BIDIOUTLINES").field_type,
            FieldType::Unknown
        );
        assert!(Field::parse_instruction("BIDIOUTLINES")
            .bidi_outline_field()
            .is_none());
        let too_long = Field::new(
            FieldType::BidiOutline,
            Cow::Owned(format!("BIDIOUTLINE {}", "x".repeat(MAX_INSTRUCTION_LEN))),
            Cow::Borrowed(""),
        );
        assert!(too_long.bidi_outline_field().is_none());
    }

    #[test]
    fn shape_fields_preserve_metadata_without_linking_or_rendering_drawings() {
        let mut shape = Field::parse_instruction(r#"SHAPE \* MERGEFORMAT"#);
        shape.result = Cow::Borrowed("cached drawing anchor");
        shape.status = FieldStatus {
            dirty: true,
            locked: true,
            ..FieldStatus::default()
        };
        shape.owner = FieldOwner::Body;
        shape.position = 4;

        assert_eq!(shape.field_type, FieldType::Shape);
        let shape = shape.shape_field().unwrap();
        assert_eq!(shape.instruction(), r#"SHAPE \* MERGEFORMAT"#);
        assert_eq!(shape.opaque_instructions(), r#"\* MERGEFORMAT"#);
        assert_eq!(shape.cached_result(), Some("cached drawing anchor"));
        assert!(shape.is_dirty());
        assert!(shape.is_locked());
        assert_eq!(shape.owner(), FieldOwner::Body);
        assert_eq!(shape.position(), 4);

        let bare = Field::parse_instruction("shape");
        assert_eq!(bare.field_type, FieldType::Shape);
        assert_eq!(bare.shape_field().unwrap().opaque_instructions(), "");
        assert_eq!(
            Field::parse_instruction("SHAPES").field_type,
            FieldType::Unknown
        );
        assert!(Field::parse_instruction("SHAPES").shape_field().is_none());
        let too_long = Field::new(
            FieldType::Shape,
            Cow::Owned(format!("SHAPE {}", "x".repeat(MAX_INSTRUCTION_LEN))),
            Cow::Borrowed(""),
        );
        assert!(too_long.shape_field().is_none());
    }

    #[test]
    fn legacy_form_fields_preserve_metadata_without_filling_or_executing() {
        let mut text = Field::parse_instruction(r#"FORMTEXT \* MERGEFORMAT"#);
        text.result = Cow::Borrowed("cached text field");
        text.status = FieldStatus {
            dirty: true,
            locked: true,
            ..FieldStatus::default()
        };
        text.owner = FieldOwner::Body;
        text.position = 4;

        assert_eq!(text.field_type, FieldType::FormText);
        let text_field = text.legacy_form_field().unwrap();
        assert_eq!(text_field.kind(), LegacyFormFieldKind::Text);
        assert_eq!(text_field.instruction(), r#"FORMTEXT \* MERGEFORMAT"#);
        assert_eq!(text_field.opaque_instructions(), r#"\* MERGEFORMAT"#);
        assert_eq!(text_field.cached_result(), Some("cached text field"));
        assert!(text_field.is_dirty());
        assert!(text_field.is_locked());
        assert_eq!(text_field.owner(), FieldOwner::Body);
        assert_eq!(text_field.position(), 4);

        let checkbox = Field::parse_instruction("formcheckbox");
        assert_eq!(checkbox.field_type, FieldType::FormCheckbox);
        let checkbox = checkbox.legacy_form_field().unwrap();
        assert_eq!(checkbox.kind(), LegacyFormFieldKind::CheckBox);
        assert_eq!(checkbox.opaque_instructions(), "");

        let drop_down = Field::parse_instruction(r#"FORMDROPDOWN \* MERGEFORMAT"#);
        assert_eq!(drop_down.field_type, FieldType::FormDropdown);
        let drop_down = drop_down.legacy_form_field().unwrap();
        assert_eq!(drop_down.kind(), LegacyFormFieldKind::DropDown);
        assert_eq!(drop_down.opaque_instructions(), r#"\* MERGEFORMAT"#);

        for instruction in [r#"FORMTEXTUAL"#, r#"FORMCHECKBOXLIST"#] {
            assert_eq!(
                Field::parse_instruction(instruction).field_type,
                FieldType::Unknown
            );
            assert!(
                Field::parse_instruction(instruction)
                    .legacy_form_field()
                    .is_none()
            );
        }

        let mismatched_kind = Field::new(
            FieldType::FormText,
            Cow::Borrowed("FORMCHECKBOX"),
            Cow::Borrowed(""),
        );
        assert!(mismatched_kind.legacy_form_field().is_none());

        let too_long = Field::new(
            FieldType::FormText,
            Cow::Owned(format!(
                "FORMTEXT {}",
                "x".repeat(MAX_INSTRUCTION_LEN)
            )),
            Cow::Borrowed(""),
        );
        assert!(too_long.legacy_form_field().is_none());
    }

    #[test]
    fn auto_text_fields_preserve_metadata_without_lookup_or_insertion() {
        let mut glossary =
            Field::parse_instruction(r#"GLOSSARY "Legacy Clause" \* MERGEFORMAT \q opaque"#);
        glossary.result = Cow::Borrowed("cached glossary entry");
        glossary.status = FieldStatus {
            dirty: true,
            locked: true,
            ..FieldStatus::default()
        };
        glossary.owner = FieldOwner::Body;
        glossary.position = 4;

        assert_eq!(glossary.field_type, FieldType::Glossary);
        let glossary = glossary.auto_text_field().unwrap();
        assert_eq!(
            glossary.instruction(),
            r#"GLOSSARY "Legacy Clause" \* MERGEFORMAT \q opaque"#
        );
        assert_eq!(glossary.kind(), AutoTextFieldKind::Glossary);
        assert_eq!(glossary.entry_name(), "Legacy Clause");
        assert_eq!(glossary.unknown_switches().len(), 2);
        assert_eq!(glossary.unknown_switches()[0].name, "*");
        assert_eq!(
            glossary.unknown_switches()[0].value.as_deref(),
            Some("MERGEFORMAT")
        );
        assert_eq!(glossary.unknown_switches()[1].name, "q");
        assert_eq!(
            glossary.unknown_switches()[1].value.as_deref(),
            Some("opaque")
        );
        assert_eq!(glossary.cached_result(), Some("cached glossary entry"));
        assert!(glossary.is_dirty());
        assert!(glossary.is_locked());
        assert_eq!(glossary.owner(), FieldOwner::Body);
        assert_eq!(glossary.position(), 4);

        let auto_text = Field::parse_instruction(r#"autotext "Reusable Clause" \* MERGEFORMAT"#);
        assert_eq!(auto_text.field_type, FieldType::AutoText);
        let auto_text = auto_text.auto_text_field().unwrap();
        assert_eq!(auto_text.kind(), AutoTextFieldKind::AutoText);
        assert_eq!(auto_text.entry_name(), "Reusable Clause");
        assert_eq!(auto_text.unknown_switches().len(), 1);
        assert_eq!(auto_text.unknown_switches()[0].name, "*");
        assert_eq!(
            auto_text.unknown_switches()[0].value.as_deref(),
            Some("MERGEFORMAT")
        );
        assert_eq!(auto_text.cached_result(), None);

        let auto_text_list = Field::parse_instruction("AUTOTEXTLIST display");
        assert_eq!(auto_text_list.field_type, FieldType::AutoTextList);
        assert!(auto_text_list.auto_text_field().is_none());
        for instruction in [
            "GLOSSARY",
            r#"GLOSSARY ""#,
            "GLOSSARY Entry unexpected",
            r#"GLOSSARY Entry \"#,
        ] {
            assert!(
                Field::parse_instruction(instruction)
                    .auto_text_field()
                    .is_none(),
                "{instruction}"
            );
        }
        let too_long = Field::new(
            FieldType::Glossary,
            Cow::Owned(format!(
                "GLOSSARY Entry {}",
                "x".repeat(MAX_INSTRUCTION_LEN)
            )),
            Cow::Borrowed(""),
        );
        assert!(too_long.auto_text_field().is_none());
    }

    #[test]
    fn auto_text_list_fields_preserve_metadata_without_selection_or_insertion() {
        let mut list = Field::parse_instruction(
            r#"AUTOTEXTLIST "Choose a name" \s "Name Style" \t "Right-click to select" \* MERGEFORMAT \q opaque"#,
        );
        list.result = Cow::Borrowed("cached selection");
        list.status = FieldStatus {
            dirty: true,
            locked: true,
            ..FieldStatus::default()
        };
        list.owner = FieldOwner::Body;
        list.position = 4;

        assert_eq!(list.field_type, FieldType::AutoTextList);
        let list = list.auto_text_list_field().unwrap();
        assert_eq!(
            list.instruction(),
            r#"AUTOTEXTLIST "Choose a name" \s "Name Style" \t "Right-click to select" \* MERGEFORMAT \q opaque"#
        );
        assert_eq!(list.display_text(), Some("Choose a name"));
        assert_eq!(
            list.options(),
            &[
                AutoTextListOption::Style(Cow::Borrowed("Name Style")),
                AutoTextListOption::Tip(Cow::Borrowed("Right-click to select")),
            ]
        );
        assert_eq!(list.unknown_switches().len(), 2);
        assert_eq!(list.unknown_switches()[0].name, "*");
        assert_eq!(
            list.unknown_switches()[0].value.as_deref(),
            Some("MERGEFORMAT")
        );
        assert_eq!(list.unknown_switches()[1].name, "q");
        assert_eq!(list.unknown_switches()[1].value.as_deref(), Some("opaque"));
        assert_eq!(list.cached_result(), Some("cached selection"));
        assert!(list.is_dirty());
        assert!(list.is_locked());
        assert_eq!(list.owner(), FieldOwner::Body);
        assert_eq!(list.position(), 4);

        let style_only = Field::parse_instruction(r#"autotextlist \s NameStyle"#);
        assert_eq!(style_only.field_type, FieldType::AutoTextList);
        let style_only = style_only.auto_text_list_field().unwrap();
        assert_eq!(style_only.display_text(), None);
        assert_eq!(
            style_only.options(),
            &[AutoTextListOption::Style(Cow::Borrowed("NameStyle"))]
        );
        assert_eq!(style_only.cached_result(), None);

        let empty_display = Field::parse_instruction(r#"AUTOTEXTLIST "" \s NameStyle"#);
        let empty_display = empty_display.auto_text_list_field().unwrap();
        assert_eq!(empty_display.display_text(), Some(""));

        assert_eq!(
            Field::parse_instruction("AUTOTEXTLISTS display").field_type,
            FieldType::Unknown
        );
        for instruction in [
            r#"AUTOTEXTLIST \s"#,
            r#"AUTOTEXTLIST \t"#,
            r#"AUTOTEXTLIST \s \"#,
            "AUTOTEXTLIST display unexpected",
            r#"AUTOTEXTLIST \"#,
            r#"AUTOTEXTLIST "unterminated"#,
        ] {
            assert!(
                Field::parse_instruction(instruction)
                    .auto_text_list_field()
                    .is_none(),
                "{instruction}"
            );
        }
        let too_long = Field::new(
            FieldType::AutoTextList,
            Cow::Owned(format!("AUTOTEXTLIST {}", "x".repeat(MAX_INSTRUCTION_LEN))),
            Cow::Borrowed(""),
        );
        assert!(too_long.auto_text_list_field().is_none());
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
    fn index_fields_preserve_stored_configuration_without_generation() {
        let mut field = Field::parse_instruction(
            r#"INDEX \b Scope \c 2 \d "-" \e ", " \f Intro \g "–" \h "A Entries" \k ". " \l "; " \p A-C \r \s Figure \y \z 1033 \* MERGEFORMAT"#,
        );
        field.result = Cow::Borrowed("cached index");
        field.status = FieldStatus {
            dirty: true,
            locked: true,
            ..FieldStatus::default()
        };
        field.owner = FieldOwner::Body;
        field.position = 4;

        assert_eq!(field.field_type, FieldType::Index);
        let index = field.index().unwrap();
        assert_eq!(index.instruction(), field.instruction);
        assert_eq!(
            index.options(),
            &[
                IndexOption::Bookmark(Cow::Borrowed("Scope")),
                IndexOption::Columns(Cow::Borrowed("2")),
                IndexOption::SequencePageSeparator(Cow::Borrowed("-")),
                IndexOption::EntryPageNumberSeparator(Cow::Borrowed(", ")),
                IndexOption::EntryType(Cow::Borrowed("Intro")),
                IndexOption::PageRangeSeparator(Cow::Borrowed("–")),
                IndexOption::Heading(Cow::Borrowed("A Entries")),
                IndexOption::CrossReferenceSeparator(Cow::Borrowed(". ")),
                IndexOption::PageNumberSeparator(Cow::Borrowed("; ")),
                IndexOption::LetterRange(Cow::Borrowed("A-C")),
                IndexOption::RunIn,
                IndexOption::SequenceIdentifier(Cow::Borrowed("Figure")),
                IndexOption::UseYomi,
                IndexOption::LanguageId(Cow::Borrowed("1033")),
            ]
        );
        assert_eq!(index.cached_result(), Some("cached index"));
        assert!(index.is_dirty());
        assert!(index.is_locked());
        assert_eq!(index.owner(), FieldOwner::Body);
        assert_eq!(index.position(), 4);
        assert_eq!(index.unknown_switches().len(), 1);
        assert_eq!(index.unknown_switches()[0].name, "*");
        assert_eq!(
            index.unknown_switches()[0].value.as_deref(),
            Some("MERGEFORMAT")
        );

        assert!(Field::parse_instruction("INDEX").index().is_some());
        assert!(Field::parse_instruction(r"INDEX \b").index().is_none());
        assert!(
            Field::parse_instruction(r"INDEX \r unexpected")
                .index()
                .is_none()
        );
        assert!(
            Field::parse_instruction("INDEX unexpected")
                .index()
                .is_none()
        );
        assert_eq!(
            Field::parse_instruction(r"INDEXES \c 2").field_type,
            FieldType::Unknown
        );
    }

    #[test]
    fn index_entry_fields_preserve_stored_metadata_without_generation() {
        let mut field = Field::parse_instruction(
            r#"XE "Office Open XML:Syntax" \b \f Intro \i \r PageRange \t "See syntax" \y "Office" \* MERGEFORMAT"#,
        );
        field.result = Cow::Borrowed("cached entry");
        field.status = FieldStatus {
            dirty: true,
            locked: true,
            ..FieldStatus::default()
        };
        field.owner = FieldOwner::Body;
        field.position = 4;

        assert_eq!(field.field_type, FieldType::IndexEntry);
        let entry = field.index_entry().unwrap();
        assert_eq!(entry.instruction(), field.instruction);
        assert_eq!(entry.entry(), "Office Open XML:Syntax");
        assert_eq!(
            entry.options(),
            &[
                IndexEntryOption::BoldPageNumber,
                IndexEntryOption::EntryType(Cow::Borrowed("Intro")),
                IndexEntryOption::ItalicPageNumber,
                IndexEntryOption::PageRangeBookmark(Cow::Borrowed("PageRange")),
                IndexEntryOption::CrossReference(Cow::Borrowed("See syntax")),
                IndexEntryOption::Yomi(Cow::Borrowed("Office")),
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

        assert!(Field::parse_instruction("XE").index_entry().is_none());
        assert!(
            Field::parse_instruction(r"XE \f Intro")
                .index_entry()
                .is_none()
        );
        assert!(
            Field::parse_instruction(r"XE entry unexpected")
                .index_entry()
                .is_none()
        );
        assert!(
            Field::parse_instruction(r"XE entry \b unexpected")
                .index_entry()
                .is_none()
        );
        assert_eq!(
            Field::parse_instruction(r"XER entry").field_type,
            FieldType::Unknown
        );
    }

    #[test]
    fn citation_fields_preserve_stored_metadata_without_resolving_sources() {
        let mut field = Field::parse_instruction(
            r#"CITATION Ecma01 \l 1033 \f "see " \s " (appendix)" \p 42 \v 2 \n \t \y \m Ecma02 \m Ecma03 \* MERGEFORMAT"#,
        );
        field.result = Cow::Borrowed("cached citation");
        field.status = FieldStatus {
            dirty: true,
            locked: true,
            ..FieldStatus::default()
        };
        field.owner = FieldOwner::Body;
        field.position = 4;

        assert_eq!(field.field_type, FieldType::Citation);
        let citation = field.citation().unwrap();
        assert_eq!(citation.instruction(), field.instruction);
        assert_eq!(citation.source_tag(), "Ecma01");
        assert_eq!(
            citation.options(),
            &[
                CitationOption::LanguageId(Cow::Borrowed("1033")),
                CitationOption::Prefix(Cow::Borrowed("see ")),
                CitationOption::Suffix(Cow::Borrowed(" (appendix)")),
                CitationOption::PageNumber(Cow::Borrowed("42")),
                CitationOption::VolumeNumber(Cow::Borrowed("2")),
                CitationOption::SuppressAuthor,
                CitationOption::SuppressTitle,
                CitationOption::SuppressYear,
                CitationOption::AdditionalSourceTag(Cow::Borrowed("Ecma02")),
                CitationOption::AdditionalSourceTag(Cow::Borrowed("Ecma03")),
            ]
        );
        assert_eq!(citation.cached_result(), Some("cached citation"));
        assert!(citation.is_dirty());
        assert!(citation.is_locked());
        assert_eq!(citation.owner(), FieldOwner::Body);
        assert_eq!(citation.position(), 4);
        assert_eq!(citation.unknown_switches().len(), 1);
        assert_eq!(citation.unknown_switches()[0].name, "*");
        assert_eq!(
            citation.unknown_switches()[0].value.as_deref(),
            Some("MERGEFORMAT")
        );

        assert!(Field::parse_instruction("CITATION").citation().is_none());
        assert!(
            Field::parse_instruction(r"CITATION \l 1033")
                .citation()
                .is_none()
        );
        assert!(
            Field::parse_instruction("CITATION Ecma01 unexpected")
                .citation()
                .is_none()
        );
        assert!(
            Field::parse_instruction(r"CITATION Ecma01 \l")
                .citation()
                .is_none()
        );
        assert!(
            Field::parse_instruction(r"CITATION Ecma01 \n unexpected")
                .citation()
                .is_none()
        );
        assert_eq!(
            Field::parse_instruction("CITATIONS Ecma01").field_type,
            FieldType::Unknown
        );
    }

    #[test]
    fn bibliography_fields_preserve_stored_metadata_without_generation() {
        let mut field =
            Field::parse_instruction(r#"BIBLIOGRAPHY \l 1033 \f en-US \m Ecma01 \* MERGEFORMAT"#);
        field.result = Cow::Borrowed("cached bibliography");
        field.status = FieldStatus {
            dirty: true,
            locked: true,
            ..FieldStatus::default()
        };
        field.owner = FieldOwner::Body;
        field.position = 4;

        assert_eq!(field.field_type, FieldType::Bibliography);
        let bibliography = field.bibliography().unwrap();
        assert_eq!(bibliography.instruction(), field.instruction);
        assert_eq!(
            bibliography.options(),
            &[
                BibliographyOption::LanguageId(Cow::Borrowed("1033")),
                BibliographyOption::FilterLanguageId(Cow::Borrowed("en-US")),
                BibliographyOption::SourceTag(Cow::Borrowed("Ecma01")),
            ]
        );
        assert_eq!(bibliography.cached_result(), Some("cached bibliography"));
        assert!(bibliography.is_dirty());
        assert!(bibliography.is_locked());
        assert_eq!(bibliography.owner(), FieldOwner::Body);
        assert_eq!(bibliography.position(), 4);
        assert_eq!(bibliography.unknown_switches().len(), 1);
        assert_eq!(bibliography.unknown_switches()[0].name, "*");
        assert_eq!(
            bibliography.unknown_switches()[0].value.as_deref(),
            Some("MERGEFORMAT")
        );

        assert!(
            Field::parse_instruction("BIBLIOGRAPHY")
                .bibliography()
                .is_some()
        );
        assert!(
            Field::parse_instruction("BIBLIOGRAPHY unexpected")
                .bibliography()
                .is_none()
        );
        assert!(
            Field::parse_instruction(r"BIBLIOGRAPHY \f")
                .bibliography()
                .is_none()
        );
        assert_eq!(
            Field::parse_instruction(r"BIBLIOGRAPHIES \l 1033").field_type,
            FieldType::Unknown
        );
    }

    #[test]
    fn document_variable_fields_preserve_names_without_resolution() {
        let mut field =
            Field::parse_instruction(r#"DOCVARIABLE "Customer Region" \* MERGEFORMAT"#);
        field.result = Cow::Borrowed("cached region");
        field.status = FieldStatus {
            dirty: true,
            locked: true,
            ..FieldStatus::default()
        };
        field.owner = FieldOwner::Body;
        field.position = 4;

        assert_eq!(field.field_type, FieldType::DocumentVariable);
        let variable = field.document_variable().unwrap();
        assert_eq!(variable.instruction(), field.instruction);
        assert_eq!(variable.variable_name(), "Customer Region");
        assert_eq!(variable.cached_result(), Some("cached region"));
        assert!(variable.is_dirty());
        assert!(variable.is_locked());
        assert_eq!(variable.owner(), FieldOwner::Body);
        assert_eq!(variable.position(), 4);
        assert_eq!(variable.unknown_switches().len(), 1);
        assert_eq!(variable.unknown_switches()[0].name, "*");
        assert_eq!(
            variable.unknown_switches()[0].value.as_deref(),
            Some("MERGEFORMAT")
        );

        assert!(
            Field::parse_instruction("DOCVARIABLE")
                .document_variable()
                .is_none()
        );
        assert!(
            Field::parse_instruction(r"DOCVARIABLE \* MERGEFORMAT")
                .document_variable()
                .is_none()
        );
        assert!(
            Field::parse_instruction("DOCVARIABLE Customer unexpected")
                .document_variable()
                .is_none()
        );
        assert_eq!(
            Field::parse_instruction("DOCVARIABLES Customer").field_type,
            FieldType::Unknown
        );
    }

    #[test]
    fn document_property_fields_preserve_names_without_resolution() {
        let mut field = Field::parse_instruction(
            r#"DOCPROPERTY "Project Name" \* MERGEFORMAT \@ "MMMM d, yyyy""#,
        );
        field.result = Cow::Borrowed("cached project");
        field.status = FieldStatus {
            dirty: true,
            locked: true,
            ..FieldStatus::default()
        };
        field.owner = FieldOwner::Body;
        field.position = 4;

        assert_eq!(field.field_type, FieldType::DocumentProperty);
        let property = field.document_property().unwrap();
        assert_eq!(property.instruction(), field.instruction);
        assert_eq!(property.property_name(), "Project Name");
        assert_eq!(property.cached_result(), Some("cached project"));
        assert!(property.is_dirty());
        assert!(property.is_locked());
        assert_eq!(property.owner(), FieldOwner::Body);
        assert_eq!(property.position(), 4);
        assert_eq!(property.switches().len(), 2);
        assert_eq!(property.switches()[0].name, "*");
        assert_eq!(
            property.switches()[0].value.as_deref(),
            Some("MERGEFORMAT")
        );
        assert_eq!(property.switches()[1].name, "@");
        assert_eq!(
            property.switches()[1].value.as_deref(),
            Some("MMMM d, yyyy")
        );

        assert!(
            Field::parse_instruction("DOCPROPERTY")
                .document_property()
                .is_none()
        );
        assert!(
            Field::parse_instruction(r#"DOCPROPERTY \"#)
                .document_property()
                .is_none()
        );
        assert!(
            Field::parse_instruction(r#"DOCPROPERTY \* MERGEFORMAT"#)
                .document_property()
                .is_none()
        );
        assert!(
            Field::parse_instruction(r#"DOCPROPERTY """#)
                .document_property()
                .is_none()
        );
        assert!(
            Field::parse_instruction("DOCPROPERTY Project unexpected")
                .document_property()
                .is_none()
        );
        assert!(
            Field::parse_instruction(r#"DOCPROPERTY Project \"#)
                .document_property()
                .is_none()
        );
        assert!(
            Field::parse_instruction(r#"DOCPROPERTY Project \* \"#)
                .document_property()
                .is_none()
        );
        assert_eq!(
            Field::parse_instruction("DOCPROPERTYS Project").field_type,
            FieldType::Unknown
        );
        let too_long = Field::new(
            FieldType::DocumentProperty,
            Cow::Owned(format!("DOCPROPERTY {}", "x".repeat(MAX_INSTRUCTION_LEN))),
            Cow::Borrowed(""),
        );
        assert!(too_long.document_property().is_none());
    }

    #[test]
    fn info_fields_preserve_stored_metadata_without_resolution_or_updates() {
        let mut field = Field::parse_instruction(
            r#"INFO TITLE "Stored title override" \* MERGEFORMAT \@ "opaque format""#,
        );
        field.result = Cow::Borrowed("cached title");
        field.status = FieldStatus {
            dirty: true,
            locked: true,
            ..FieldStatus::default()
        };
        field.owner = FieldOwner::Body;
        field.position = 4;

        assert_eq!(field.field_type, FieldType::Info);
        let information = field.info_field().unwrap();
        assert_eq!(information.instruction(), field.instruction);
        assert_eq!(information.information_type(), "TITLE");
        assert_eq!(information.new_value(), Some("Stored title override"));
        assert_eq!(information.cached_result(), Some("cached title"));
        assert!(information.is_dirty());
        assert!(information.is_locked());
        assert_eq!(information.owner(), FieldOwner::Body);
        assert_eq!(information.position(), 4);
        assert_eq!(information.switches().len(), 2);
        assert_eq!(information.switches()[0].name, "*");
        assert_eq!(
            information.switches()[0].value.as_deref(),
            Some("MERGEFORMAT")
        );
        assert_eq!(information.switches()[1].name, "@");
        assert_eq!(
            information.switches()[1].value.as_deref(),
            Some("opaque format")
        );

        let template = Field::parse_instruction("INFO TEMPLATE");
        assert_eq!(template.field_type, FieldType::Info);
        assert_eq!(
            template.info_field().unwrap().information_type(),
            "TEMPLATE"
        );
        assert_eq!(template.info_field().unwrap().new_value(), None);

        for instruction in [
            "INFO",
            r#"INFO "" "#,
            r#"INFO TITLE "Stored title" unexpected"#,
            r#"INFO TITLE "unterminated"#,
            r#"INFO TITLE \"#,
        ] {
            assert!(
                Field::parse_instruction(instruction).info_field().is_none(),
                "{instruction}"
            );
        }

        assert_eq!(
            Field::parse_instruction("INFOS TITLE").field_type,
            FieldType::Unknown
        );
        assert!(Field::parse_instruction(r#"TITLE "Stored title override""#)
            .info_field()
            .is_none());
        let too_long = Field::new(
            FieldType::Info,
            Cow::Owned(format!("INFO {}", "x".repeat(MAX_INSTRUCTION_LEN))),
            Cow::Borrowed(""),
        );
        assert!(too_long.info_field().is_none());
    }

    #[test]
    fn document_information_fields_preserve_kinds_without_reading_or_calculating_values() {
        let mut field =
            Field::parse_instruction(r#"TITLE \* MERGEFORMAT \@ "opaque format""#);
        field.result = Cow::Borrowed("cached title");
        field.status = FieldStatus {
            dirty: true,
            locked: true,
            ..FieldStatus::default()
        };
        field.owner = FieldOwner::Body;
        field.position = 4;

        assert_eq!(field.field_type, FieldType::DocumentInformation);
        let information = field.document_information().unwrap();
        assert_eq!(information.instruction(), field.instruction);
        assert_eq!(information.kind(), DocumentInformationFieldKind::Title);
        assert_eq!(information.cached_result(), Some("cached title"));
        assert!(information.is_dirty());
        assert!(information.is_locked());
        assert_eq!(information.owner(), FieldOwner::Body);
        assert_eq!(information.position(), 4);
        assert_eq!(information.switches().len(), 2);
        assert_eq!(information.switches()[0].name, "*");
        assert_eq!(
            information.switches()[0].value.as_deref(),
            Some("MERGEFORMAT")
        );
        assert_eq!(information.switches()[1].name, "@");
        assert_eq!(
            information.switches()[1].value.as_deref(),
            Some("opaque format")
        );

        for (instruction, kind) in [
            ("TITLE", DocumentInformationFieldKind::Title),
            ("SUBJECT", DocumentInformationFieldKind::Subject),
            ("AUTHOR", DocumentInformationFieldKind::Author),
            ("KEYWORDS", DocumentInformationFieldKind::Keywords),
            ("COMMENTS", DocumentInformationFieldKind::Comments),
            ("LASTSAVEDBY", DocumentInformationFieldKind::LastSavedBy),
            ("CREATEDATE", DocumentInformationFieldKind::CreateDate),
            ("SAVEDATE", DocumentInformationFieldKind::SaveDate),
            ("PRINTDATE", DocumentInformationFieldKind::PrintDate),
            ("REVNUM", DocumentInformationFieldKind::RevisionNumber),
            ("EDITTIME", DocumentInformationFieldKind::EditTime),
            ("NUMPAGES", DocumentInformationFieldKind::NumberOfPages),
            ("NUMWORDS", DocumentInformationFieldKind::NumberOfWords),
            ("NUMCHARS", DocumentInformationFieldKind::NumberOfCharacters),
        ] {
            let field = Field::parse_instruction(instruction);
            assert_eq!(field.field_type, FieldType::DocumentInformation);
            let information = field.document_information().unwrap();
            assert_eq!(information.kind(), kind);
            assert_eq!(information.kind().field_keyword(), instruction);
            assert!(information.switches().is_empty());
        }

        assert!(
            Field::parse_instruction("TITLE unexpected")
                .document_information()
                .is_none()
        );
        assert!(
            Field::parse_instruction(r#"AUTHOR "unterminated"#)
                .document_information()
                .is_none()
        );
        assert!(
            Field::parse_instruction("COMMENTS \\")
                .document_information()
                .is_none()
        );
        assert!(
            Field::parse_instruction(r"LASTSAVEDBY \* MERGEFORMAT unexpected")
                .document_information()
                .is_none()
        );
        assert!(
            Field::parse_instruction("NUMWORDS unexpected")
                .document_information()
                .is_none()
        );
        assert_eq!(
            Field::parse_instruction("AUTHORS").field_type,
            FieldType::Unknown
        );
        let too_long = Field::new(
            FieldType::DocumentInformation,
            Cow::Owned(format!("TITLE {}", "x".repeat(MAX_INSTRUCTION_LEN))),
            Cow::Borrowed(""),
        );
        assert!(too_long.document_information().is_none());
    }

    #[test]
    fn document_context_fields_preserve_kinds_without_reading_or_calculating_values() {
        let mut field = Field::parse_instruction(r"FILENAME \p");
        field.result = Cow::Borrowed("cached file name");
        field.status = FieldStatus {
            dirty: true,
            locked: true,
            ..FieldStatus::default()
        };
        field.owner = FieldOwner::Body;
        field.position = 4;

        assert_eq!(field.field_type, FieldType::DocumentContext);
        let context = field.document_context().unwrap();
        assert_eq!(context.instruction(), field.instruction);
        assert_eq!(context.kind(), DocumentContextFieldKind::FileName);
        assert_eq!(context.cached_result(), Some("cached file name"));
        assert!(context.is_dirty());
        assert!(context.is_locked());
        assert_eq!(context.owner(), FieldOwner::Body);
        assert_eq!(context.position(), 4);
        assert_eq!(context.switches().len(), 1);
        assert_eq!(context.switches()[0].name, "p");
        assert_eq!(context.switches()[0].value, None);

        for (instruction, kind, field_type) in [
            (
                "FILENAME",
                DocumentContextFieldKind::FileName,
                FieldType::DocumentContext,
            ),
            (
                "TEMPLATE",
                DocumentContextFieldKind::Template,
                FieldType::DocumentContext,
            ),
            ("DATE", DocumentContextFieldKind::Date, FieldType::Date),
            ("TIME", DocumentContextFieldKind::Time, FieldType::Date),
            ("PAGE", DocumentContextFieldKind::Page, FieldType::Page),
            (
                "FILESIZE",
                DocumentContextFieldKind::FileSize,
                FieldType::DocumentContext,
            ),
            (
                "SECTION",
                DocumentContextFieldKind::Section,
                FieldType::DocumentContext,
            ),
            (
                "SECTIONPAGES",
                DocumentContextFieldKind::SectionPages,
                FieldType::DocumentContext,
            ),
        ] {
            let field = Field::parse_instruction(instruction);
            assert_eq!(field.field_type, field_type);
            let context = field.document_context().unwrap();
            assert_eq!(context.kind(), kind);
            assert_eq!(context.kind().field_keyword(), instruction);
            assert!(context.switches().is_empty());
        }

        assert!(
            Field::parse_instruction("FILENAME unexpected")
                .document_context()
                .is_none()
        );
        assert!(
            Field::parse_instruction(r#"TEMPLATE "unterminated"#)
                .document_context()
                .is_none()
        );
        assert!(
            Field::parse_instruction("FILENAME \\")
                .document_context()
                .is_none()
        );
        assert!(
            Field::parse_instruction("PAGE unexpected")
                .document_context()
                .is_none()
        );
        assert!(
            Field::parse_instruction("SECTIONPAGES unexpected")
                .document_context()
                .is_none()
        );
        assert_eq!(
            Field::parse_instruction("FILENAMES").field_type,
            FieldType::Unknown
        );
        assert_eq!(
            Field::parse_instruction("PAGES").field_type,
            FieldType::Unknown
        );
        assert_eq!(
            Field::parse_instruction("SECTIONPAGE").field_type,
            FieldType::Unknown
        );
        let too_long = Field::new(
            FieldType::DocumentContext,
            Cow::Owned(format!("FILENAME {}", "x".repeat(MAX_INSTRUCTION_LEN))),
            Cow::Borrowed(""),
        );
        assert!(too_long.document_context().is_none());
    }

    #[test]
    fn merge_fields_preserve_names_without_merging() {
        let mut field = Field::parse_instruction(
            r#"MERGEFIELD "Customer Region" \b "Dear " \f "!" \* MERGEFORMAT"#,
        );
        field.result = Cow::Borrowed("cached customer");
        field.status = FieldStatus {
            dirty: true,
            locked: true,
            ..FieldStatus::default()
        };
        field.owner = FieldOwner::Body;
        field.position = 4;

        assert_eq!(field.field_type, FieldType::MergeField);
        let merge = field.merge_field().unwrap();
        assert_eq!(merge.instruction(), field.instruction);
        assert_eq!(merge.field_name(), "Customer Region");
        assert_eq!(merge.cached_result(), Some("cached customer"));
        assert!(merge.is_dirty());
        assert!(merge.is_locked());
        assert_eq!(merge.owner(), FieldOwner::Body);
        assert_eq!(merge.position(), 4);
        assert_eq!(merge.switches().len(), 3);
        assert_eq!(merge.switches()[0].name, "b");
        assert_eq!(merge.switches()[0].value.as_deref(), Some("Dear "));
        assert_eq!(merge.switches()[1].name, "f");
        assert_eq!(merge.switches()[1].value.as_deref(), Some("!"));
        assert_eq!(merge.switches()[2].name, "*");
        assert_eq!(merge.switches()[2].value.as_deref(), Some("MERGEFORMAT"));

        assert!(
            Field::parse_instruction("MERGEFIELD")
                .merge_field()
                .is_none()
        );
        assert!(
            Field::parse_instruction(r#"MERGEFIELD \b "Dear ""#)
                .merge_field()
                .is_none()
        );
        assert!(
            Field::parse_instruction("MERGEFIELD Customer unexpected")
                .merge_field()
                .is_none()
        );
        assert_eq!(
            Field::parse_instruction("MERGEFIELDS Customer").field_type,
            FieldType::Unknown
        );
    }

    #[test]
    fn mail_merge_data_fields_preserve_sources_without_connecting_or_merging() {
        let mut field = Field::parse_instruction(
            r#"DATA "recipients source.csv" "headers source.csv" \* MERGEFORMAT \q opaque"#,
        );
        field.result = Cow::Borrowed("cached mail-merge source");
        field.status = FieldStatus {
            dirty: true,
            locked: true,
            ..FieldStatus::default()
        };
        field.owner = FieldOwner::Body;
        field.position = 4;

        assert_eq!(field.field_type, FieldType::MailMergeData);
        let data = field.mail_merge_data().unwrap();
        assert_eq!(data.instruction(), field.instruction);
        assert_eq!(data.data_source(), "recipients source.csv");
        assert_eq!(data.header_source(), Some("headers source.csv"));
        assert_eq!(data.cached_result(), Some("cached mail-merge source"));
        assert!(data.is_dirty());
        assert!(data.is_locked());
        assert_eq!(data.owner(), FieldOwner::Body);
        assert_eq!(data.position(), 4);
        assert_eq!(data.switches().len(), 2);
        assert_eq!(data.switches()[0].name, "*");
        assert_eq!(data.switches()[0].value.as_deref(), Some("MERGEFORMAT"));
        assert_eq!(data.switches()[1].name, "q");
        assert_eq!(data.switches()[1].value.as_deref(), Some("opaque"));

        let without_header = Field::parse_instruction(r#"data recipients.csv \q opaque"#);
        assert_eq!(without_header.field_type, FieldType::MailMergeData);
        let without_header = without_header.mail_merge_data().unwrap();
        assert_eq!(without_header.data_source(), "recipients.csv");
        assert_eq!(without_header.header_source(), None);
        assert_eq!(without_header.switches()[0].name, "q");
        assert_eq!(without_header.switches()[0].value.as_deref(), Some("opaque"));

        assert!(Field::parse_instruction("DATA").mail_merge_data().is_none());
        assert!(
            Field::parse_instruction(r#"DATA \* MERGEFORMAT"#)
                .mail_merge_data()
                .is_none()
        );
        assert!(
            Field::parse_instruction("DATA recipients.csv headers.csv unexpected")
                .mail_merge_data()
                .is_none()
        );
        assert_eq!(
            Field::parse_instruction("DATABASE recipients.csv").field_type,
            FieldType::Unknown
        );
        let too_long = Field::new(
            FieldType::MailMergeData,
            Cow::Owned(format!("DATA {}", "x".repeat(MAX_INSTRUCTION_LEN))),
            Cow::Borrowed(""),
        );
        assert!(too_long.mail_merge_data().is_none());
    }

    #[test]
    fn mail_merge_counters_preserve_cached_results_without_merging() {
        let mut record = Field::parse_instruction("MERGEREC");
        record.result = Cow::Borrowed("12");
        record.status = FieldStatus {
            dirty: true,
            locked: true,
            ..FieldStatus::default()
        };
        record.owner = FieldOwner::Body;
        record.position = 4;

        assert_eq!(record.field_type, FieldType::MergeRecord);
        let record_counter = record.mail_merge_counter().unwrap();
        assert_eq!(record_counter.instruction(), record.instruction);
        assert_eq!(record_counter.kind(), MailMergeCounterKind::Record);
        assert_eq!(record_counter.cached_result(), Some("12"));
        assert!(record_counter.is_dirty());
        assert!(record_counter.is_locked());
        assert_eq!(record_counter.owner(), FieldOwner::Body);
        assert_eq!(record_counter.position(), 4);

        let sequence = Field::parse_instruction("mergeSEQ");
        assert_eq!(sequence.field_type, FieldType::MergeSequence);
        assert_eq!(
            sequence.mail_merge_counter().unwrap().kind(),
            MailMergeCounterKind::Sequence
        );

        assert!(
            Field::parse_instruction("MERGEREC 12")
                .mail_merge_counter()
                .is_none()
        );
        assert!(
            Field::parse_instruction(r"MERGESEQ \* MERGEFORMAT")
                .mail_merge_counter()
                .is_none()
        );
        assert_eq!(
            Field::parse_instruction("MERGERECORD").field_type,
            FieldType::Unknown
        );
    }

    #[test]
    fn mail_merge_next_fields_preserve_cached_results_without_advancing_records() {
        let mut next = Field::parse_instruction("NEXT");
        next.result = Cow::Borrowed("cached next");
        next.status = FieldStatus {
            dirty: true,
            locked: true,
            ..FieldStatus::default()
        };
        next.owner = FieldOwner::Body;
        next.position = 4;

        assert_eq!(next.field_type, FieldType::MailMergeNext);
        let next_field = next.mail_merge_next().unwrap();
        assert_eq!(next_field.instruction(), next.instruction);
        assert_eq!(next_field.cached_result(), Some("cached next"));
        assert!(next_field.is_dirty());
        assert!(next_field.is_locked());
        assert_eq!(next_field.owner(), FieldOwner::Body);
        assert_eq!(next_field.position(), 4);

        assert!(
            Field::parse_instruction("NEXT 12")
                .mail_merge_next()
                .is_none()
        );
        assert!(
            Field::parse_instruction(r"NEXT \* MERGEFORMAT")
                .mail_merge_next()
                .is_none()
        );
        assert_eq!(
            Field::parse_instruction("NEXTIF Customer = Ada").field_type,
            FieldType::MailMergeNextIf
        );
        assert!(
            Field::parse_instruction("NEXTIF Customer = Ada")
                .mail_merge_next()
                .is_none()
        );
    }

    #[test]
    fn conditional_mail_merge_controls_preserve_cached_results_without_merging() {
        let mut next_if = Field::parse_instruction(r#"NEXTIF Customer = "Ada""#);
        next_if.result = Cow::Borrowed("cached nextif");
        next_if.status = FieldStatus {
            dirty: true,
            locked: true,
            ..FieldStatus::default()
        };
        next_if.owner = FieldOwner::Body;
        next_if.position = 4;

        assert_eq!(next_if.field_type, FieldType::MailMergeNextIf);
        let next_if_control = next_if.mail_merge_conditional_control().unwrap();
        assert_eq!(
            next_if_control.kind(),
            MailMergeConditionalControlKind::NextIf
        );
        assert_eq!(next_if_control.comparison(), r#"Customer = "Ada""#);
        assert_eq!(next_if_control.cached_result(), Some("cached nextif"));
        assert!(next_if_control.is_dirty());
        assert!(next_if_control.is_locked());
        assert_eq!(next_if_control.owner(), FieldOwner::Body);
        assert_eq!(next_if_control.position(), 4);

        let skip_if = Field::parse_instruction("skipif MERGEFIELD Order < 100");
        assert_eq!(skip_if.field_type, FieldType::MailMergeSkipIf);
        let skip_if_control = skip_if.mail_merge_conditional_control().unwrap();
        assert_eq!(
            skip_if_control.kind(),
            MailMergeConditionalControlKind::SkipIf
        );
        assert_eq!(skip_if_control.comparison(), "MERGEFIELD Order < 100");

        assert!(
            Field::parse_instruction("NEXTIF")
                .mail_merge_conditional_control()
                .is_none()
        );
        assert!(
            Field::parse_instruction("SKIPIF   ")
                .mail_merge_conditional_control()
                .is_none()
        );
        assert_eq!(
            Field::parse_instruction("NEXTIFF Customer = Ada").field_type,
            FieldType::Unknown
        );
    }

    #[test]
    fn if_fields_preserve_cached_results_without_evaluation() {
        let mut conditional = Field::parse_instruction(r#"IF "A" = "A" "yes" "no""#);
        conditional.result = Cow::Borrowed("yes");
        conditional.status = FieldStatus {
            dirty: true,
            locked: true,
            ..FieldStatus::default()
        };
        conditional.owner = FieldOwner::Body;
        conditional.position = 4;

        assert_eq!(conditional.field_type, FieldType::If);
        let if_field = conditional.if_field().unwrap();
        assert_eq!(if_field.instruction(), conditional.instruction);
        assert_eq!(if_field.expression(), r#""A" = "A" "yes" "no""#);
        assert_eq!(if_field.cached_result(), Some("yes"));
        assert!(if_field.is_dirty());
        assert!(if_field.is_locked());
        assert_eq!(if_field.owner(), FieldOwner::Body);
        assert_eq!(if_field.position(), 4);

        assert!(Field::parse_instruction("IF").if_field().is_none());
        assert_eq!(
            Field::parse_instruction(r#"IFF "A" = "A" "yes" "no""#).field_type,
            FieldType::Unknown
        );
    }

    #[test]
    fn compare_fields_preserve_cached_results_without_evaluation() {
        let mut comparison = Field::parse_instruction(r#"COMPARE "CustomerNumber" >= 4"#);
        comparison.result = Cow::Borrowed("1");
        comparison.status = FieldStatus {
            dirty: true,
            locked: true,
            ..FieldStatus::default()
        };
        comparison.owner = FieldOwner::Body;
        comparison.position = 4;

        assert_eq!(comparison.field_type, FieldType::Compare);
        let compare_field = comparison.compare_field().unwrap();
        assert_eq!(compare_field.instruction(), comparison.instruction);
        assert_eq!(compare_field.comparison(), r#""CustomerNumber" >= 4"#);
        assert_eq!(compare_field.cached_result(), Some("1"));
        assert!(compare_field.is_dirty());
        assert!(compare_field.is_locked());
        assert_eq!(compare_field.owner(), FieldOwner::Body);
        assert_eq!(compare_field.position(), 4);

        let nested = Field::parse_instruction("compare MERGEFIELD CustomerRating <= 9");
        let compare_field = nested.compare_field().unwrap();
        assert_eq!(compare_field.comparison(), "MERGEFIELD CustomerRating <= 9");

        assert!(
            Field::parse_instruction("COMPARE")
                .compare_field()
                .is_none()
        );
        assert!(
            Field::parse_instruction("COMPARE   ")
                .compare_field()
                .is_none()
        );
        assert_eq!(
            Field::parse_instruction("COMPARES Customer = 1").field_type,
            FieldType::Unknown
        );
    }

    #[test]
    fn set_fields_preserve_cached_results_without_evaluation_or_state_changes() {
        let mut set =
            Field::parse_instruction(r#"SET "Customer Region" "North America" \* MERGEFORMAT"#);
        set.result = Cow::Borrowed("cached region");
        set.status = FieldStatus {
            dirty: true,
            locked: true,
            ..FieldStatus::default()
        };
        set.owner = FieldOwner::Body;
        set.position = 4;

        assert_eq!(set.field_type, FieldType::Set);
        let set_field = set.set_field().unwrap();
        assert_eq!(set_field.instruction(), set.instruction);
        assert_eq!(set_field.target_name(), "Customer Region");
        assert_eq!(set_field.expression(), r#""North America" \* MERGEFORMAT"#);
        assert_eq!(set_field.cached_result(), Some("cached region"));
        assert!(set_field.is_dirty());
        assert!(set_field.is_locked());
        assert_eq!(set_field.owner(), FieldOwner::Body);
        assert_eq!(set_field.position(), 4);

        let formula = Field::parse_instruction("set Total =SUM(ABOVE) + 1");
        assert_eq!(formula.field_type, FieldType::Set);
        let formula = formula.set_field().unwrap();
        assert_eq!(formula.target_name(), "Total");
        assert_eq!(formula.expression(), "=SUM(ABOVE) + 1");

        assert_eq!(
            Field::parse_instruction("SETTINGS value").field_type,
            FieldType::Unknown
        );
        for instruction in [
            "SET",
            r#"SET "" value"#,
            "SET Target",
            "SET Target   ",
            r#"SET \* value"#,
            r#"SET "Target"expression"#,
        ] {
            assert!(
                Field::parse_instruction(instruction).set_field().is_none(),
                "{instruction}"
            );
        }
    }

    #[test]
    fn sequence_fields_preserve_metadata_without_bookmark_lookup_or_numbering() {
        let mut sequence = Field::parse_instruction(r#"SEQ Figure FigureChapter \r 3 \* ARABIC"#);
        sequence.result = Cow::Borrowed("3");
        sequence.status = FieldStatus {
            dirty: true,
            locked: true,
            ..FieldStatus::default()
        };
        sequence.owner = FieldOwner::Body;
        sequence.position = 4;

        assert_eq!(sequence.field_type, FieldType::Sequence);
        let sequence_field = sequence.sequence_field().unwrap();
        assert_eq!(sequence_field.instruction(), sequence.instruction);
        assert_eq!(sequence_field.identifier(), "Figure");
        assert_eq!(sequence_field.bookmark(), Some("FigureChapter"));
        assert_eq!(sequence_field.tail(), r"\r 3 \* ARABIC");
        assert_eq!(sequence_field.cached_result(), Some("3"));
        assert!(sequence_field.is_dirty());
        assert!(sequence_field.is_locked());
        assert_eq!(sequence_field.owner(), FieldOwner::Body);
        assert_eq!(sequence_field.position(), 4);

        let table = Field::parse_instruction(r#"seq Table \s 1 \* ROMAN"#);
        assert_eq!(table.field_type, FieldType::Sequence);
        let table = table.sequence_field().unwrap();
        assert_eq!(table.identifier(), "Table");
        assert_eq!(table.bookmark(), None);
        assert_eq!(table.tail(), r"\s 1 \* ROMAN");

        let bare = Field::parse_instruction("SEQ Footnote");
        let bare = bare.sequence_field().unwrap();
        assert_eq!(bare.identifier(), "Footnote");
        assert_eq!(bare.bookmark(), None);
        assert_eq!(bare.tail(), "");

        assert_eq!(
            Field::parse_instruction("SEQUENCE Figure").field_type,
            FieldType::Unknown
        );
        for instruction in [
            "SEQ",
            r#"SEQ ""#,
            r#"SEQ Figure ""#,
            r#"SEQ "Figure"Bookmark"#,
        ] {
            assert!(
                Field::parse_instruction(instruction)
                    .sequence_field()
                    .is_none(),
                "{instruction}"
            );
        }
    }

    #[test]
    fn formula_fields_preserve_cached_results_without_evaluation() {
        let mut formula = Field::parse_instruction(r#"=SUM(ABOVE) \* MERGEFORMAT"#);
        formula.result = Cow::Borrowed("42");
        formula.status = FieldStatus {
            dirty: true,
            locked: true,
            ..FieldStatus::default()
        };
        formula.owner = FieldOwner::Body;
        formula.position = 4;

        assert_eq!(formula.field_type, FieldType::Formula);
        let formula_field = formula.formula_field().unwrap();
        assert_eq!(formula_field.instruction(), formula.instruction);
        assert_eq!(formula_field.formula(), r"SUM(ABOVE) \* MERGEFORMAT");
        assert_eq!(formula_field.cached_result(), Some("42"));
        assert!(formula_field.is_dirty());
        assert!(formula_field.is_locked());
        assert_eq!(formula_field.owner(), FieldOwner::Body);
        assert_eq!(formula_field.position(), 4);

        let conditional = Field::parse_instruction(r#"= IF(1 = 1, "yes", "no")"#);
        assert_eq!(conditional.field_type, FieldType::Formula);
        let conditional = conditional.formula_field().unwrap();
        assert_eq!(conditional.formula(), r#"IF(1 = 1, "yes", "no")"#);

        let missing = Field::parse_instruction("=");
        assert_eq!(missing.field_type, FieldType::Formula);
        assert!(missing.formula_field().is_none());
        assert_eq!(
            Field::parse_instruction("EQUAL 1 + 1").field_type,
            FieldType::Unknown
        );
    }

    #[test]
    fn quote_fields_preserve_cached_text_without_inserting_or_transforming_it() {
        let mut quote = Field::parse_instruction(r#"QUOTE "Stored literal" \* MERGEFORMAT \# "000""#);
        quote.result = Cow::Borrowed("cached literal");
        quote.status = FieldStatus {
            dirty: true,
            locked: true,
            ..FieldStatus::default()
        };
        quote.owner = FieldOwner::Body;
        quote.position = 4;

        assert_eq!(quote.field_type, FieldType::Quote);
        let quote_field = quote.quote_field().unwrap();
        assert_eq!(quote_field.instruction(), quote.instruction);
        assert_eq!(quote_field.text(), "Stored literal");
        assert_eq!(quote_field.cached_result(), Some("cached literal"));
        assert!(quote_field.is_dirty());
        assert!(quote_field.is_locked());
        assert_eq!(quote_field.owner(), FieldOwner::Body);
        assert_eq!(quote_field.position(), 4);
        assert_eq!(quote_field.switches().len(), 2);
        assert_eq!(quote_field.switches()[0].name, "*");
        assert_eq!(quote_field.switches()[0].value.as_deref(), Some("MERGEFORMAT"));
        assert_eq!(quote_field.switches()[1].name, "#");
        assert_eq!(quote_field.switches()[1].value.as_deref(), Some("000"));

        let unquoted = Field::parse_instruction(r#"quote CompatibilityText \@ "MMMM""#);
        assert_eq!(unquoted.field_type, FieldType::Quote);
        let unquoted = unquoted.quote_field().unwrap();
        assert_eq!(unquoted.text(), "CompatibilityText");
        assert_eq!(unquoted.switches()[0].name, "@");
        assert_eq!(unquoted.switches()[0].value.as_deref(), Some("MMMM"));

        for instruction in [
            "QUOTE",
            r#"QUOTE \* MERGEFORMAT"#,
            r#"QUOTE "literal" unexpected"#,
            r#"QUOTE "unterminated"#,
        ] {
            assert!(
                Field::parse_instruction(instruction).quote_field().is_none(),
                "{instruction}"
            );
        }
        assert_eq!(
            Field::parse_instruction(r#"QUOTEY "not a quote field""#).field_type,
            FieldType::Unknown
        );
    }

    #[test]
    fn symbol_fields_preserve_cached_metadata_without_mapping_codes_or_inserting_glyphs() {
        let mut symbol =
            Field::parse_instruction(r#"SYMBOL 0xA9 \f "Symbol" \s 12 \u"#);
        symbol.result = Cow::Borrowed("cached copyright");
        symbol.status = FieldStatus {
            dirty: true,
            locked: true,
            ..FieldStatus::default()
        };
        symbol.owner = FieldOwner::Body;
        symbol.position = 4;

        assert_eq!(symbol.field_type, FieldType::Symbol);
        let symbol_field = symbol.symbol_field().unwrap();
        assert_eq!(symbol_field.instruction(), symbol.instruction);
        assert_eq!(symbol_field.character_argument(), "0xA9");
        assert_eq!(symbol_field.cached_result(), Some("cached copyright"));
        assert!(symbol_field.is_dirty());
        assert!(symbol_field.is_locked());
        assert_eq!(symbol_field.owner(), FieldOwner::Body);
        assert_eq!(symbol_field.position(), 4);
        assert_eq!(symbol_field.switches().len(), 3);
        assert_eq!(symbol_field.switches()[0].name, "f");
        assert_eq!(symbol_field.switches()[0].value.as_deref(), Some("Symbol"));
        assert_eq!(symbol_field.switches()[1].name, "s");
        assert_eq!(symbol_field.switches()[1].value.as_deref(), Some("12"));
        assert_eq!(symbol_field.switches()[2].name, "u");
        assert_eq!(symbol_field.switches()[2].value, None);

        let symbol = Field::parse_instruction(r"symbol 163 \a \h \j");
        assert_eq!(symbol.field_type, FieldType::Symbol);
        let symbol = symbol.symbol_field().unwrap();
        assert_eq!(symbol.character_argument(), "163");
        assert_eq!(symbol.switches()[0].name, "a");
        assert_eq!(symbol.switches()[1].name, "h");
        assert_eq!(symbol.switches()[2].name, "j");

        for instruction in [
            "SYMBOL",
            r#"SYMBOL \f "Symbol""#,
            r"SYMBOL 0xA9 unexpected",
            r#"SYMBOL 0xA9 \f "unterminated"#,
        ] {
            assert!(
                Field::parse_instruction(instruction).symbol_field().is_none(),
                "{instruction}"
            );
        }
        assert_eq!(
            Field::parse_instruction("SYMBOLS 163").field_type,
            FieldType::Unknown
        );
    }

    #[test]
    fn automatic_number_fields_preserve_cached_metadata_without_calculating_numbers_or_layout() {
        let mut automatic =
            Field::parse_instruction(r#"AUTONUM \s "." \* MERGEFORMAT"#);
        automatic.result = Cow::Borrowed("7.");
        automatic.status = FieldStatus {
            dirty: true,
            locked: true,
            ..FieldStatus::default()
        };
        automatic.owner = FieldOwner::Body;
        automatic.position = 4;

        assert_eq!(automatic.field_type, FieldType::AutoNumber);
        let automatic = automatic.auto_number_field().unwrap();
        assert_eq!(automatic.kind(), AutoNumberFieldKind::AutoNum);
        assert_eq!(automatic.kind().field_keyword(), "AUTONUM");
        assert_eq!(automatic.cached_result(), Some("7."));
        assert!(automatic.is_dirty());
        assert!(automatic.is_locked());
        assert_eq!(automatic.owner(), FieldOwner::Body);
        assert_eq!(automatic.position(), 4);
        assert_eq!(automatic.switches().len(), 2);
        assert_eq!(automatic.switches()[0].name, "s");
        assert_eq!(automatic.switches()[0].value.as_deref(), Some("."));
        assert_eq!(automatic.switches()[1].name, "*");
        assert_eq!(
            automatic.switches()[1].value.as_deref(),
            Some("MERGEFORMAT")
        );

        let legal = Field::parse_instruction(r#"autonumlgl \e \s ")" "#);
        assert_eq!(legal.field_type, FieldType::AutoNumber);
        let legal = legal.auto_number_field().unwrap();
        assert_eq!(legal.kind(), AutoNumberFieldKind::AutoNumLegal);
        assert_eq!(legal.kind().field_keyword(), "AUTONUMLGL");
        assert_eq!(legal.switches()[0].name, "e");
        assert_eq!(legal.switches()[0].value, None);
        assert_eq!(legal.switches()[1].name, "s");
        assert_eq!(legal.switches()[1].value.as_deref(), Some(")"));

        let outline = Field::parse_instruction("AUTONUMOUT");
        assert_eq!(outline.field_type, FieldType::AutoNumber);
        let outline = outline.auto_number_field().unwrap();
        assert_eq!(outline.kind(), AutoNumberFieldKind::AutoNumOutline);
        assert_eq!(outline.kind().field_keyword(), "AUTONUMOUT");
        assert!(outline.switches().is_empty());

        for instruction in [
            "AUTONUM unexpected",
            r#"AUTONUMLGL \s "unterminated"#,
            "AUTONUMOUT \\",
        ] {
            assert!(
                Field::parse_instruction(instruction)
                    .auto_number_field()
                    .is_none(),
                "{instruction}"
            );
        }
        assert_eq!(
            Field::parse_instruction("AUTONUMS").field_type,
            FieldType::Unknown
        );
        let too_long = Field::new(
            FieldType::AutoNumber,
            Cow::Owned(format!("AUTONUM {}", "x".repeat(MAX_INSTRUCTION_LEN))),
            Cow::Borrowed(""),
        );
        assert!(too_long.auto_number_field().is_none());
    }

    #[test]
    fn list_number_fields_preserve_cached_metadata_without_reading_lists_or_calculating_numbers() {
        let mut numbered =
            Field::parse_instruction(r#"LISTNUM NumberDefault \l 6 \s 3 \* MERGEFORMAT"#);
        numbered.result = Cow::Borrowed("(iii)");
        numbered.status = FieldStatus {
            dirty: true,
            locked: true,
            ..FieldStatus::default()
        };
        numbered.owner = FieldOwner::Body;
        numbered.position = 4;

        assert_eq!(numbered.field_type, FieldType::ListNumber);
        let numbered = numbered.list_number_field().unwrap();
        assert_eq!(numbered.list_name(), Some("NumberDefault"));
        assert_eq!(numbered.cached_result(), Some("(iii)"));
        assert!(numbered.is_dirty());
        assert!(numbered.is_locked());
        assert_eq!(numbered.owner(), FieldOwner::Body);
        assert_eq!(numbered.position(), 4);
        assert_eq!(numbered.switches().len(), 3);
        assert_eq!(numbered.switches()[0].name, "l");
        assert_eq!(numbered.switches()[0].value.as_deref(), Some("6"));
        assert_eq!(numbered.switches()[1].name, "s");
        assert_eq!(numbered.switches()[1].value.as_deref(), Some("3"));
        assert_eq!(numbered.switches()[2].name, "*");
        assert_eq!(
            numbered.switches()[2].value.as_deref(),
            Some("MERGEFORMAT")
        );

        let outline = Field::parse_instruction(r#"listnum "Outline Default" \l 4"#);
        assert_eq!(outline.field_type, FieldType::ListNumber);
        let outline = outline.list_number_field().unwrap();
        assert_eq!(outline.list_name(), Some("Outline Default"));
        assert_eq!(outline.switches()[0].name, "l");
        assert_eq!(outline.switches()[0].value.as_deref(), Some("4"));

        let unnamed = Field::parse_instruction(r"LISTNUM \l 2");
        assert_eq!(unnamed.field_type, FieldType::ListNumber);
        let unnamed = unnamed.list_number_field().unwrap();
        assert_eq!(unnamed.list_name(), None);
        assert_eq!(unnamed.switches()[0].name, "l");
        assert_eq!(unnamed.switches()[0].value.as_deref(), Some("2"));

        for instruction in [
            "LISTNUM NumberDefault unexpected",
            r#"LISTNUM "unterminated"#,
            "LISTNUM \\",
        ] {
            assert!(
                Field::parse_instruction(instruction)
                    .list_number_field()
                    .is_none(),
                "{instruction}"
            );
        }
        assert_eq!(
            Field::parse_instruction("LISTNUMBER NumberDefault").field_type,
            FieldType::Unknown
        );
        let too_long = Field::new(
            FieldType::ListNumber,
            Cow::Owned(format!("LISTNUM {}", "x".repeat(MAX_INSTRUCTION_LEN))),
            Cow::Borrowed(""),
        );
        assert!(too_long.list_number_field().is_none());
    }

    #[test]
    fn style_reference_fields_preserve_metadata_without_style_or_layout_resolution() {
        let mut style_reference = Field::parse_instruction(
            r#"STYLEREF "Heading 1" \l \n \p \r \t \w \* MERGEFORMAT \q opaque"#,
        );
        style_reference.result = Cow::Borrowed("Cached heading");
        style_reference.status = FieldStatus {
            dirty: true,
            locked: true,
            ..FieldStatus::default()
        };
        style_reference.owner = FieldOwner::Body;
        style_reference.position = 4;

        assert_eq!(style_reference.field_type, FieldType::StyleReference);
        let style_reference = style_reference.style_reference_field().unwrap();
        assert_eq!(
            style_reference.instruction(),
            r#"STYLEREF "Heading 1" \l \n \p \r \t \w \* MERGEFORMAT \q opaque"#
        );
        assert_eq!(style_reference.style_name(), "Heading 1");
        assert_eq!(
            style_reference.options(),
            &[
                StyleReferenceFieldOption::FollowingText,
                StyleReferenceFieldOption::ParagraphNumber,
                StyleReferenceFieldOption::RelativePosition,
                StyleReferenceFieldOption::ParagraphNumberRelativeContext,
                StyleReferenceFieldOption::SuppressNonNumberText,
                StyleReferenceFieldOption::ParagraphNumberFullContext,
            ]
        );
        assert_eq!(style_reference.unknown_switches().len(), 2);
        assert_eq!(style_reference.unknown_switches()[0].name, "*");
        assert_eq!(
            style_reference.unknown_switches()[0].value.as_deref(),
            Some("MERGEFORMAT")
        );
        assert_eq!(style_reference.unknown_switches()[1].name, "q");
        assert_eq!(
            style_reference.unknown_switches()[1].value.as_deref(),
            Some("opaque")
        );
        assert_eq!(style_reference.cached_result(), Some("Cached heading"));
        assert!(style_reference.is_dirty());
        assert!(style_reference.is_locked());
        assert_eq!(style_reference.owner(), FieldOwner::Body);
        assert_eq!(style_reference.position(), 4);

        let title = Field::parse_instruction(r#"styleref Title \n"#);
        assert_eq!(title.field_type, FieldType::StyleReference);
        let title = title.style_reference_field().unwrap();
        assert_eq!(title.style_name(), "Title");
        assert_eq!(
            title.options(),
            &[StyleReferenceFieldOption::ParagraphNumber]
        );
        assert!(title.unknown_switches().is_empty());
        assert_eq!(title.cached_result(), None);

        assert_eq!(
            Field::parse_instruction("STYLEREFS Heading").field_type,
            FieldType::Unknown
        );
        for instruction in [
            "STYLEREF",
            r#"STYLEREF ""#,
            r#"STYLEREF Heading \l unexpected"#,
            "STYLEREF Heading unexpected",
            r#"STYLEREF Heading \"#,
            r#"STYLEREF Heading "unterminated"#,
        ] {
            assert!(
                Field::parse_instruction(instruction)
                    .style_reference_field()
                    .is_none(),
                "{instruction}"
            );
        }
    }

    #[test]
    fn prompt_fields_preserve_metadata_without_displaying_prompts() {
        let mut ask =
            Field::parse_instruction(r#"ASK AskResponse "What is your first name?" \d "" \o"#);
        ask.result = Cow::Borrowed("cached ask response");
        ask.status = FieldStatus {
            dirty: true,
            locked: true,
            ..FieldStatus::default()
        };
        ask.owner = FieldOwner::Body;
        ask.position = 4;

        assert_eq!(ask.field_type, FieldType::Ask);
        let ask = ask.prompt_field().unwrap();
        assert_eq!(ask.kind(), PromptFieldKind::Ask);
        assert_eq!(ask.bookmark(), Some("AskResponse"));
        assert_eq!(ask.prompt(), Some("What is your first name?"));
        assert_eq!(ask.default_response(), Some(""));
        assert!(ask.prompts_once_per_mail_merge());
        assert_eq!(ask.cached_result(), Some("cached ask response"));
        assert!(ask.is_dirty());
        assert!(ask.is_locked());
        assert_eq!(ask.owner(), FieldOwner::Body);
        assert_eq!(ask.position(), 4);

        let fill_in = Field::parse_instruction(r#"fillin "Enter appointment time" \d "09:00""#);
        assert_eq!(fill_in.field_type, FieldType::FillIn);
        let fill_in = fill_in.prompt_field().unwrap();
        assert_eq!(fill_in.kind(), PromptFieldKind::FillIn);
        assert_eq!(fill_in.bookmark(), None);
        assert_eq!(fill_in.prompt(), Some("Enter appointment time"));
        assert_eq!(fill_in.default_response(), Some("09:00"));
        assert!(!fill_in.prompts_once_per_mail_merge());

        let default_only = Field::parse_instruction(r#"FILLIN \d "recent response" \o"#);
        let default_only = default_only.prompt_field().unwrap();
        assert_eq!(default_only.prompt(), None);
        assert_eq!(default_only.default_response(), Some("recent response"));
        assert!(default_only.prompts_once_per_mail_merge());

        for instruction in [
            "ASK",
            r#"ASK "" "Question""#,
            "ASK Answer",
            r#"ASK Answer "Question" \d"#,
            r#"ASK Answer "Question" \o extra"#,
            r#"FILLIN "Question" \x"#,
            r#"FILLIN "Question" \d "first" \d "second""#,
            r#"FILLIN "Question" \o \o"#,
        ] {
            assert!(
                Field::parse_instruction(instruction)
                    .prompt_field()
                    .is_none(),
                "{instruction}"
            );
        }
        assert_eq!(
            Field::parse_instruction(r#"ASKER Answer "Question""#).field_type,
            FieldType::Unknown
        );
    }

    #[test]
    fn user_identity_fields_preserve_metadata_without_reading_host_identity() {
        let mut address = Field::parse_instruction(r#"USERADDRESS "10 Top Secret Lane" \* Upper"#);
        address.result = Cow::Borrowed("10 TOP SECRET LANE");
        address.status = FieldStatus {
            dirty: true,
            locked: true,
            ..FieldStatus::default()
        };
        address.owner = FieldOwner::Body;
        address.position = 4;

        assert_eq!(address.field_type, FieldType::UserAddress);
        let address = address.user_identity_field().unwrap();
        assert_eq!(address.kind(), UserIdentityFieldKind::Address);
        assert_eq!(address.override_value(), Some("10 Top Secret Lane"));
        assert_eq!(address.formatting(), Some(UserIdentityFormatting::Upper));
        assert_eq!(address.cached_result(), Some("10 TOP SECRET LANE"));
        assert!(address.is_dirty());
        assert!(address.is_locked());
        assert_eq!(address.owner(), FieldOwner::Body);
        assert_eq!(address.position(), 4);

        let initials = Field::parse_instruction(r#"userinitials \* Lower"#);
        assert_eq!(initials.field_type, FieldType::UserInitials);
        let initials = initials.user_identity_field().unwrap();
        assert_eq!(initials.kind(), UserIdentityFieldKind::Initials);
        assert_eq!(initials.override_value(), None);
        assert_eq!(initials.formatting(), Some(UserIdentityFormatting::Lower));

        let name = Field::parse_instruction(r#"USERNAME "Ada Lovelace" \* FirstCap"#);
        assert_eq!(name.field_type, FieldType::UserName);
        let name = name.user_identity_field().unwrap();
        assert_eq!(name.kind(), UserIdentityFieldKind::Name);
        assert_eq!(name.override_value(), Some("Ada Lovelace"));
        assert_eq!(name.formatting(), Some(UserIdentityFormatting::FirstCap));

        for instruction in [
            "USERADDRESS \\*",
            "USERINITIALS \\* Title",
            "USERNAME \\* Upper \\* Lower",
            "USERNAME Ada \\l 1033",
            "USERADDRESS Ada Lovelace",
        ] {
            assert!(
                Field::parse_instruction(instruction)
                    .user_identity_field()
                    .is_none(),
                "{instruction}"
            );
        }

        let blank_override = Field::parse_instruction(r#"USERNAME "" \* Caps"#);
        let blank_override = blank_override.user_identity_field().unwrap();
        assert_eq!(blank_override.override_value(), Some(""));
        assert_eq!(
            blank_override.formatting(),
            Some(UserIdentityFormatting::Caps)
        );
        assert_eq!(
            Field::parse_instruction("USERNAMES Ada").field_type,
            FieldType::Unknown
        );
    }

    #[test]
    fn advance_fields_preserve_placement_metadata_without_changing_layout() {
        let mut advance =
            Field::parse_instruction(r#"ADVANCE \u 6 \d 12 \l 20 \r -4 \x 150 \y "72" \d -3"#);
        advance.result = Cow::Borrowed("cached placement");
        advance.status = FieldStatus {
            dirty: true,
            locked: true,
            ..FieldStatus::default()
        };
        advance.owner = FieldOwner::Body;
        advance.position = 4;

        assert_eq!(advance.field_type, FieldType::Advance);
        let advance = advance.advance_field().unwrap();
        let adjustments = advance
            .adjustments()
            .iter()
            .map(|adjustment| (adjustment.operation(), adjustment.points()))
            .collect::<Vec<_>>();
        assert_eq!(
            adjustments,
            vec![
                (AdvanceFieldOperation::Up, 6),
                (AdvanceFieldOperation::Down, 12),
                (AdvanceFieldOperation::Left, 20),
                (AdvanceFieldOperation::Right, -4),
                (AdvanceFieldOperation::HorizontalPosition, 150),
                (AdvanceFieldOperation::VerticalPosition, 72),
                (AdvanceFieldOperation::Down, -3),
            ]
        );
        assert_eq!(advance.cached_result(), Some("cached placement"));
        assert!(advance.is_dirty());
        assert!(advance.is_locked());
        assert_eq!(advance.owner(), FieldOwner::Body);
        assert_eq!(advance.position(), 4);

        let no_adjustments = Field::parse_instruction("aDvAnCe");
        assert_eq!(no_adjustments.field_type, FieldType::Advance);
        assert!(
            no_adjustments
                .advance_field()
                .unwrap()
                .adjustments()
                .is_empty()
        );

        for instruction in [
            r#"ADVANCE \d"#,
            r#"ADVANCE \z 10"#,
            r#"ADVANCE \x 1.5"#,
            r#"ADVANCE \u 9223372036854775808"#,
            "ADVANCE 12",
            r#"ADVANCE \d 6 trailing"#,
        ] {
            assert!(
                Field::parse_instruction(instruction)
                    .advance_field()
                    .is_none(),
                "{instruction}"
            );
        }
        assert_eq!(
            Field::parse_instruction(r#"ADVANCER \u 6"#).field_type,
            FieldType::Unknown
        );
    }

    #[test]
    fn mail_merge_recipient_fields_preserve_layout_metadata_without_merging() {
        let mut address = Field::parse_instruction(
            r#"ADDRESSBLOCK \c 2 \d \e "United States" \e Canada \f "<<_FIRST0_>> <<_LAST0_>>" \l 1033 \* MERGEFORMAT"#,
        );
        address.result = Cow::Borrowed("cached address");
        address.status = FieldStatus {
            dirty: true,
            locked: true,
            ..FieldStatus::default()
        };
        address.owner = FieldOwner::Body;
        address.position = 4;

        assert_eq!(address.field_type, FieldType::AddressBlock);
        let address = address.mail_merge_recipient_field().unwrap();
        assert_eq!(address.kind(), MailMergeRecipientFieldKind::AddressBlock);
        assert_eq!(
            address.country_inclusion(),
            Some(AddressBlockCountryInclusion::UnlessExcluded)
        );
        assert!(address.formats_using_recipient_country());
        let excluded = address
            .excluded_countries()
            .iter()
            .map(Cow::as_ref)
            .collect::<Vec<_>>();
        assert_eq!(excluded, vec!["United States", "Canada"]);
        assert_eq!(address.format_template(), Some("<<_FIRST0_>> <<_LAST0_>>"));
        assert_eq!(address.language(), Some("1033"));
        assert_eq!(address.greeting_fallback_text(), None);
        assert_eq!(address.unknown_switches().len(), 1);
        assert_eq!(address.unknown_switches()[0].name, "*");
        assert_eq!(
            address.unknown_switches()[0].value.as_deref(),
            Some("MERGEFORMAT")
        );
        assert_eq!(address.cached_result(), Some("cached address"));
        assert!(address.is_dirty());
        assert!(address.is_locked());
        assert_eq!(address.owner(), FieldOwner::Body);
        assert_eq!(address.position(), 4);

        let greeting = Field::parse_instruction(
            r#"greetingline \f "Dear <<_FIRST0_>>," \e "To Whom It May Concern" \l en-US"#,
        );
        assert_eq!(greeting.field_type, FieldType::GreetingLine);
        let greeting = greeting.mail_merge_recipient_field().unwrap();
        assert_eq!(greeting.kind(), MailMergeRecipientFieldKind::GreetingLine);
        assert_eq!(greeting.country_inclusion(), None);
        assert!(!greeting.formats_using_recipient_country());
        assert!(greeting.excluded_countries().is_empty());
        assert_eq!(greeting.format_template(), Some("Dear <<_FIRST0_>>,"));
        assert_eq!(greeting.language(), Some("en-US"));
        assert_eq!(
            greeting.greeting_fallback_text(),
            Some("To Whom It May Concern")
        );

        for instruction in [
            "ADDRESSBLOCK text",
            r"ADDRESSBLOCK \c",
            r"ADDRESSBLOCK \c 3",
            r"ADDRESSBLOCK \d 1",
            r"ADDRESSBLOCK \d \d",
            r"ADDRESSBLOCK \f",
            r#"GREETINGLINE \f "Dear" \f "Hello""#,
            r"GREETINGLINE \l",
            r#"GREETINGLINE \c "First" \e "Second""#,
        ] {
            assert!(
                Field::parse_instruction(instruction)
                    .mail_merge_recipient_field()
                    .is_none(),
                "{instruction}"
            );
        }
        assert_eq!(
            Field::parse_instruction(r"ADDRESSBLOCKING \c 1").field_type,
            FieldType::Unknown
        );
    }

    #[test]
    fn document_discovers_document_variable_fields_without_resolving_them() {
        let document = crate::RtfDocument::parse(
            r#"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst DOCVARIABLE CustomerName \\* MERGEFORMAT}{\fldrslt cached customer}}After}"#,
        )
        .unwrap();

        let fields = document.document_variable_fields();
        assert_eq!(document.document_variable_field_count(), 1);
        assert_eq!(fields.len(), 1);
        assert_eq!(fields[0].variable_name(), "CustomerName");
        assert_eq!(fields[0].cached_result(), Some("cached customer"));
        assert!(fields[0].is_dirty());
        assert!(fields[0].is_locked());
        assert!(document.document_variables().is_empty());
        assert_eq!(document.text(), "Before After");
    }

    #[test]
    fn document_discovers_document_property_fields_without_resolving_them() {
        let document = crate::RtfDocument::parse(
            r#"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst DOCPROPERTY "Project Name" \\* MERGEFORMAT \\@ "MMMM d, yyyy"}{\fldrslt cached project}}After}"#,
        )
        .unwrap();

        let fields = document.document_property_fields();
        assert_eq!(document.document_property_field_count(), 1);
        assert_eq!(fields.len(), 1);
        assert_eq!(fields[0].property_name(), "Project Name");
        assert_eq!(fields[0].cached_result(), Some("cached project"));
        assert!(fields[0].is_dirty());
        assert!(fields[0].is_locked());
        assert_eq!(fields[0].switches().len(), 2);
        assert_eq!(fields[0].switches()[0].name, "*");
        assert_eq!(fields[0].switches()[1].name, "@");
        assert_eq!(document.text(), "Before After");
    }

    #[test]
    fn document_discovers_info_fields_without_reading_or_modifying_properties() {
        let document = crate::RtfDocument::parse(
            r#"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst INFO TITLE "Stored title override" \\* MERGEFORMAT}{\fldrslt cached title}}Middle {\field\flddirty\fldlock{\*\fldinst info COMMENTS "Stored comment" \\@ "opaque format"}{\fldrslt cached comment}}After}"#,
        )
        .unwrap();

        let fields = document.info_fields();
        assert_eq!(document.info_field_count(), 2);
        assert_eq!(fields.len(), 2);
        assert_eq!(fields[0].information_type(), "TITLE");
        assert_eq!(fields[0].new_value(), Some("Stored title override"));
        assert_eq!(fields[0].cached_result(), Some("cached title"));
        assert!(fields[0].is_dirty());
        assert!(fields[0].is_locked());
        assert_eq!(fields[0].switches()[0].name, "*");
        assert_eq!(
            fields[0].switches()[0].value.as_deref(),
            Some("MERGEFORMAT")
        );
        assert_eq!(fields[1].information_type(), "COMMENTS");
        assert_eq!(fields[1].new_value(), Some("Stored comment"));
        assert_eq!(fields[1].cached_result(), Some("cached comment"));
        assert!(fields[1].is_dirty());
        assert!(fields[1].is_locked());
        assert_eq!(fields[1].switches()[0].name, "@");
        assert_eq!(
            fields[1].switches()[0].value.as_deref(),
            Some("opaque format")
        );
        assert_eq!(document.text(), "Before Middle After");
    }

    #[test]
    fn document_discovers_document_information_fields_without_reading_or_calculating_values() {
        let document = crate::RtfDocument::parse(
            r#"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst TITLE \\* MERGEFORMAT}{\fldrslt cached title}}Middle {\field{\*\fldinst author \\@ "opaque format"}{\fldrslt cached author}}After}"#,
        )
        .unwrap();

        let fields = document.document_information_fields();
        assert_eq!(document.document_information_field_count(), 2);
        assert_eq!(fields.len(), 2);
        assert_eq!(fields[0].kind(), DocumentInformationFieldKind::Title);
        assert_eq!(fields[0].cached_result(), Some("cached title"));
        assert!(fields[0].is_dirty());
        assert!(fields[0].is_locked());
        assert_eq!(fields[0].switches()[0].name, "*");
        assert_eq!(fields[1].kind(), DocumentInformationFieldKind::Author);
        assert_eq!(fields[1].cached_result(), Some("cached author"));
        assert_eq!(fields[1].switches()[0].name, "@");
        assert_eq!(
            fields[1].switches()[0].value.as_deref(),
            Some("opaque format")
        );
        assert_eq!(document.text(), "Before Middle After");
    }

    #[test]
    fn document_discovers_document_information_statistics_without_calculating() {
        let document = crate::RtfDocument::parse(
            r"{\rtf1\ansi {\field\flddirty\fldlock{\*\fldinst NUMWORDS \\* MERGEFORMAT}{\fldrslt cached words}}}",
        )
        .unwrap();

        let fields = document.document_information_fields();
        assert_eq!(document.document_information_field_count(), 1);
        assert_eq!(fields.len(), 1);
        assert_eq!(fields[0].kind(), DocumentInformationFieldKind::NumberOfWords);
        assert_eq!(fields[0].cached_result(), Some("cached words"));
        assert!(fields[0].is_dirty());
        assert!(fields[0].is_locked());
        assert_eq!(fields[0].switches().len(), 1);
        assert_eq!(fields[0].switches()[0].name, "*");
        assert_eq!(
            fields[0].switches()[0].value.as_deref(),
            Some("MERGEFORMAT")
        );
    }

    #[test]
    fn document_discovers_document_context_fields_without_reading_or_calculating_values() {
        let document = crate::RtfDocument::parse(
            r"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst FILENAME \\p}{\fldrslt cached file name}}Middle {\field{\*\fldinst TEMPLATE \\* MERGEFORMAT}{\fldrslt cached template}}After}",
        )
        .unwrap();

        let fields = document.document_context_fields();
        assert_eq!(document.document_context_field_count(), 2);
        assert_eq!(fields.len(), 2);
        assert_eq!(fields[0].kind(), DocumentContextFieldKind::FileName);
        assert_eq!(fields[0].cached_result(), Some("cached file name"));
        assert!(fields[0].is_dirty());
        assert!(fields[0].is_locked());
        assert_eq!(fields[0].switches()[0].name, "p");
        assert_eq!(fields[1].kind(), DocumentContextFieldKind::Template);
        assert_eq!(fields[1].cached_result(), Some("cached template"));
        assert_eq!(fields[1].switches()[0].name, "*");
        assert_eq!(fields[1].switches()[0].value.as_deref(), Some("MERGEFORMAT"));
        assert_eq!(document.text(), "Before Middle After");
    }

    #[test]
    fn document_discovers_document_context_page_fields_without_calculation() {
        let document = crate::RtfDocument::parse(
            r"{\rtf1\ansi {\field\flddirty\fldlock{\*\fldinst PAGE \\* MERGEFORMAT}{\fldrslt cached page}}}",
        )
        .unwrap();

        let fields = document.document_context_fields();
        assert_eq!(document.document_context_field_count(), 1);
        assert_eq!(fields.len(), 1);
        assert_eq!(fields[0].kind(), DocumentContextFieldKind::Page);
        assert_eq!(fields[0].cached_result(), Some("cached page"));
        assert!(fields[0].is_dirty());
        assert!(fields[0].is_locked());
        assert_eq!(fields[0].switches()[0].name, "*");
        assert_eq!(fields[0].switches()[0].value.as_deref(), Some("MERGEFORMAT"));
    }

    #[test]
    fn document_discovers_runtime_context_fields_without_file_or_layout_reads() {
        let document = crate::RtfDocument::parse(
            r"{\rtf1\ansi {\field\flddirty\fldlock{\*\fldinst FILESIZE \\* MERGEFORMAT}{\fldrslt cached size}}{\field\flddirty\fldlock{\*\fldinst SECTION \\* MERGEFORMAT}{\fldrslt cached section}}{\field\flddirty\fldlock{\*\fldinst SECTIONPAGES \\* MERGEFORMAT}{\fldrslt cached section pages}}}",
        )
        .unwrap();

        let fields = document.document_context_fields();
        assert_eq!(document.document_context_field_count(), 3);
        assert_eq!(fields.len(), 3);
        assert_eq!(fields[0].kind(), DocumentContextFieldKind::FileSize);
        assert_eq!(fields[0].cached_result(), Some("cached size"));
        assert_eq!(fields[1].kind(), DocumentContextFieldKind::Section);
        assert_eq!(fields[1].cached_result(), Some("cached section"));
        assert_eq!(fields[2].kind(), DocumentContextFieldKind::SectionPages);
        assert_eq!(fields[2].cached_result(), Some("cached section pages"));
        for field in fields {
            assert!(field.is_dirty());
            assert!(field.is_locked());
            assert_eq!(field.switches()[0].name, "*");
            assert_eq!(
                field.switches()[0].value.as_deref(),
                Some("MERGEFORMAT")
            );
        }
    }

    #[test]
    fn document_discovers_quote_fields_without_inserting_or_transforming_text() {
        let document = crate::RtfDocument::parse(
            r#"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst QUOTE "Stored literal" \\* MERGEFORMAT}{\fldrslt cached literal}}After}"#,
        )
        .unwrap();

        let fields = document.quote_fields();
        assert_eq!(document.quote_field_count(), 1);
        assert_eq!(fields.len(), 1);
        assert_eq!(fields[0].text(), "Stored literal");
        assert_eq!(fields[0].cached_result(), Some("cached literal"));
        assert!(fields[0].is_dirty());
        assert!(fields[0].is_locked());
        assert_eq!(fields[0].switches()[0].name, "*");
        assert_eq!(fields[0].switches()[0].value.as_deref(), Some("MERGEFORMAT"));
        assert_eq!(document.text(), "Before After");
    }

    #[test]
    fn document_discovers_symbol_fields_without_mapping_codes_or_inserting_glyphs() {
        let document = crate::RtfDocument::parse(
            r#"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst SYMBOL 0xA9 \\f "Symbol" \\s 12 \\u}{\fldrslt cached copyright}}After}"#,
        )
        .unwrap();

        let fields = document.symbol_fields();
        assert_eq!(document.symbol_field_count(), 1);
        assert_eq!(fields.len(), 1);
        assert_eq!(fields[0].character_argument(), "0xA9");
        assert_eq!(fields[0].cached_result(), Some("cached copyright"));
        assert!(fields[0].is_dirty());
        assert!(fields[0].is_locked());
        assert_eq!(fields[0].switches().len(), 3);
        assert_eq!(fields[0].switches()[0].name, "f");
        assert_eq!(fields[0].switches()[0].value.as_deref(), Some("Symbol"));
        assert_eq!(fields[0].switches()[1].name, "s");
        assert_eq!(fields[0].switches()[1].value.as_deref(), Some("12"));
        assert_eq!(fields[0].switches()[2].name, "u");
        assert_eq!(fields[0].switches()[2].value, None);
        assert_eq!(document.text(), "Before After");
    }

    #[test]
    fn document_discovers_print_fields_without_interpreting_or_sending_printer_commands() {
        let document = crate::RtfDocument::parse(
            r#"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst PRINT "ESC&l1O"}{\fldrslt cached printer}}{\field\flddirty\fldlock{\*\fldinst print \\p 2 "0 0 moveto"}{\fldrslt cached PostScript}}After}"#,
        )
        .unwrap();

        let fields = document.print_fields();
        assert_eq!(document.print_field_count(), 2);
        assert_eq!(fields.len(), 2);
        assert_eq!(fields[0].printer_instructions(), r#""ESC&l1O""#);
        assert_eq!(fields[0].cached_result(), Some("cached printer"));
        assert!(fields[0].is_dirty());
        assert!(fields[0].is_locked());
        assert_eq!(fields[1].printer_instructions(), r#"\p 2 "0 0 moveto""#);
        assert_eq!(fields[1].cached_result(), Some("cached PostScript"));
        assert!(fields[1].is_dirty());
        assert!(fields[1].is_locked());
        assert_eq!(document.text(), "Before After");
    }

    #[test]
    fn document_discovers_embed_fields_without_loading_or_activating_objects() {
        let document = crate::RtfDocument::parse(
            r#"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst EMBED Excel.Sheet.12 \\* MERGEFORMAT}{\fldrslt cached worksheet object}}{\field\flddirty\fldlock{\*\fldinst embed "Equation.DSMT4" \\d}{\fldrslt cached equation object}}After}"#,
        )
        .unwrap();

        let fields = document.embed_fields();
        assert_eq!(document.embed_field_count(), 2);
        assert_eq!(fields.len(), 2);
        assert_eq!(
            fields[0].object_instructions(),
            r#"Excel.Sheet.12 \* MERGEFORMAT"#
        );
        assert_eq!(
            fields[0].cached_result(),
            Some("cached worksheet object")
        );
        assert!(fields[0].is_dirty());
        assert!(fields[0].is_locked());
        assert_eq!(
            fields[1].object_instructions(),
            r#""Equation.DSMT4" \d"#
        );
        assert_eq!(fields[1].cached_result(), Some("cached equation object"));
        assert!(fields[1].is_dirty());
        assert!(fields[1].is_locked());
        assert_eq!(document.text(), "Before After");
    }

    #[test]
    fn document_discovers_barcode_fields_without_decoding_or_rendering() {
        let document = crate::RtfDocument::parse(
            r#"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst BARCODE "4901234567894" EAN13 \\h 1440}{\fldrslt cached EAN13 barcode}}{\field\flddirty\fldlock{\*\fldinst barcode "ABC-123" CODE39 \\d}{\fldrslt cached Code39 barcode}}After}"#,
        )
        .unwrap();

        let fields = document.barcode_fields();
        assert_eq!(document.barcode_field_count(), 2);
        assert_eq!(fields.len(), 2);
        assert_eq!(
            fields[0].barcode_instructions(),
            r#""4901234567894" EAN13 \h 1440"#
        );
        assert_eq!(fields[0].cached_result(), Some("cached EAN13 barcode"));
        assert!(fields[0].is_dirty());
        assert!(fields[0].is_locked());
        assert_eq!(
            fields[1].barcode_instructions(),
            r#""ABC-123" CODE39 \d"#
        );
        assert_eq!(fields[1].cached_result(), Some("cached Code39 barcode"));
        assert!(fields[1].is_dirty());
        assert!(fields[1].is_locked());
        assert_eq!(document.text(), "Before After");
    }

    #[test]
    fn document_discovers_bidi_outline_fields_without_resolving_numbering_or_layout() {
        let document = crate::RtfDocument::parse(
            r#"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst BIDIOUTLINE \\* MERGEFORMAT}{\fldrslt cached bidi outline number}}{\field\flddirty\fldlock{\*\fldinst bidioutline}{\fldrslt cached bare bidi outline}}After}"#,
        )
        .unwrap();

        let fields = document.bidi_outline_fields();
        assert_eq!(document.bidi_outline_field_count(), 2);
        assert_eq!(fields.len(), 2);
        assert_eq!(fields[0].opaque_instructions(), r#"\* MERGEFORMAT"#);
        assert_eq!(
            fields[0].cached_result(),
            Some("cached bidi outline number")
        );
        assert!(fields[0].is_dirty());
        assert!(fields[0].is_locked());
        assert_eq!(fields[1].opaque_instructions(), "");
        assert_eq!(
            fields[1].cached_result(),
            Some("cached bare bidi outline")
        );
        assert!(fields[1].is_dirty());
        assert!(fields[1].is_locked());
        assert_eq!(document.text(), "Before After");
    }

    #[test]
    fn document_discovers_shape_fields_without_linking_or_rendering_drawings() {
        let document = crate::RtfDocument::parse(
            r#"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst SHAPE \\* MERGEFORMAT}{\fldrslt cached drawing anchor}}{\field\flddirty\fldlock{\*\fldinst shape}{\fldrslt cached bare drawing anchor}}After}"#,
        )
        .unwrap();

        let fields = document.shape_fields();
        assert_eq!(document.shape_field_count(), 2);
        assert_eq!(fields.len(), 2);
        assert_eq!(fields[0].opaque_instructions(), r#"\* MERGEFORMAT"#);
        assert_eq!(fields[0].cached_result(), Some("cached drawing anchor"));
        assert!(fields[0].is_dirty());
        assert!(fields[0].is_locked());
        assert_eq!(fields[1].opaque_instructions(), "");
        assert_eq!(fields[1].cached_result(), Some("cached bare drawing anchor"));
        assert!(fields[1].is_dirty());
        assert!(fields[1].is_locked());
        assert_eq!(document.text(), "Before After");
    }

    #[test]
    fn document_discovers_legacy_form_fields_without_filling_or_executing() {
        let document = crate::RtfDocument::parse(
            r#"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst FORMTEXT \\* MERGEFORMAT}{\fldrslt cached text field}}{\field\flddirty\fldlock{\*\fldinst formcheckbox}{\fldrslt cached checkbox}}{\field\flddirty\fldlock{\*\fldinst FORMDROPDOWN \\* MERGEFORMAT}{\fldrslt cached drop-down selection}}After}"#,
        )
        .unwrap();

        let fields = document.legacy_form_fields();
        assert_eq!(document.legacy_form_field_count(), 3);
        assert_eq!(fields.len(), 3);
        assert_eq!(fields[0].kind(), LegacyFormFieldKind::Text);
        assert_eq!(fields[0].opaque_instructions(), r#"\* MERGEFORMAT"#);
        assert_eq!(fields[0].cached_result(), Some("cached text field"));
        assert!(fields[0].is_dirty());
        assert!(fields[0].is_locked());
        assert_eq!(fields[1].kind(), LegacyFormFieldKind::CheckBox);
        assert_eq!(fields[1].opaque_instructions(), "");
        assert_eq!(fields[1].cached_result(), Some("cached checkbox"));
        assert!(fields[1].is_dirty());
        assert!(fields[1].is_locked());
        assert_eq!(fields[2].kind(), LegacyFormFieldKind::DropDown);
        assert_eq!(fields[2].opaque_instructions(), r#"\* MERGEFORMAT"#);
        assert_eq!(
            fields[2].cached_result(),
            Some("cached drop-down selection")
        );
        assert!(fields[2].is_dirty());
        assert!(fields[2].is_locked());
        assert_eq!(document.text(), "Before After");
    }

    #[test]
    fn document_discovers_automatic_number_fields_without_calculating_numbers_or_layout() {
        let document = crate::RtfDocument::parse(
            r#"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst AUTONUM \\s "." \\* MERGEFORMAT}{\fldrslt 7.}}{\field\flddirty\fldlock{\*\fldinst AUTONUMLGL \\e \\s ")" }{\fldrslt 2.4}}{\field{\*\fldinst AUTONUMOUT}{\fldrslt III}}After}"#,
        )
        .unwrap();

        let fields = document.auto_number_fields();
        assert_eq!(document.auto_number_field_count(), 3);
        assert_eq!(fields.len(), 3);
        assert_eq!(fields[0].kind(), AutoNumberFieldKind::AutoNum);
        assert_eq!(fields[0].cached_result(), Some("7."));
        assert!(fields[0].is_dirty());
        assert!(fields[0].is_locked());
        assert_eq!(fields[0].switches()[0].name, "s");
        assert_eq!(fields[0].switches()[0].value.as_deref(), Some("."));
        assert_eq!(fields[1].kind(), AutoNumberFieldKind::AutoNumLegal);
        assert_eq!(fields[1].cached_result(), Some("2.4"));
        assert_eq!(fields[1].switches()[0].name, "e");
        assert_eq!(fields[1].switches()[1].name, "s");
        assert_eq!(fields[1].switches()[1].value.as_deref(), Some(")"));
        assert_eq!(fields[2].kind(), AutoNumberFieldKind::AutoNumOutline);
        assert_eq!(fields[2].cached_result(), Some("III"));
        assert!(fields[2].switches().is_empty());
        assert_eq!(document.text(), "Before After");
    }

    #[test]
    fn document_discovers_list_number_fields_without_reading_lists_or_calculating_numbers() {
        let document = crate::RtfDocument::parse(
            r#"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst LISTNUM NumberDefault \\l 6 \\s 3 \\* MERGEFORMAT}{\fldrslt (iii)}}{\field\flddirty\fldlock{\*\fldinst LISTNUM "Outline Default" \\l 4}{\fldrslt c}}{\field{\*\fldinst LISTNUM \\l 2}{\fldrslt i}}After}"#,
        )
        .unwrap();

        let fields = document.list_number_fields();
        assert_eq!(document.list_number_field_count(), 3);
        assert_eq!(fields.len(), 3);
        assert_eq!(fields[0].list_name(), Some("NumberDefault"));
        assert_eq!(fields[0].cached_result(), Some("(iii)"));
        assert!(fields[0].is_dirty());
        assert!(fields[0].is_locked());
        assert_eq!(fields[0].switches()[0].name, "l");
        assert_eq!(fields[0].switches()[0].value.as_deref(), Some("6"));
        assert_eq!(fields[0].switches()[1].name, "s");
        assert_eq!(fields[0].switches()[1].value.as_deref(), Some("3"));
        assert_eq!(fields[1].list_name(), Some("Outline Default"));
        assert_eq!(fields[1].cached_result(), Some("c"));
        assert_eq!(fields[1].switches()[0].name, "l");
        assert_eq!(fields[1].switches()[0].value.as_deref(), Some("4"));
        assert_eq!(fields[2].list_name(), None);
        assert_eq!(fields[2].cached_result(), Some("i"));
        assert_eq!(fields[2].switches()[0].name, "l");
        assert_eq!(fields[2].switches()[0].value.as_deref(), Some("2"));
        assert_eq!(document.text(), "Before After");
    }

    #[test]
    fn document_discovers_merge_fields_without_opening_data_sources() {
        let document = crate::RtfDocument::parse(
            r#"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst MERGEFIELD "Customer Region" \\b "Dear " \\* MERGEFORMAT}{\fldrslt cached customer}}After}"#,
        )
        .unwrap();

        let fields = document.merge_fields();
        assert_eq!(document.merge_field_count(), 1);
        assert_eq!(fields.len(), 1);
        assert_eq!(fields[0].field_name(), "Customer Region");
        assert_eq!(fields[0].cached_result(), Some("cached customer"));
        assert!(fields[0].is_dirty());
        assert!(fields[0].is_locked());
        assert_eq!(fields[0].switches()[0].name, "b");
        assert_eq!(fields[0].switches()[0].value.as_deref(), Some("Dear "));
        assert_eq!(document.text(), "Before After");
    }

    #[test]
    fn document_discovers_mail_merge_data_fields_without_opening_sources_or_merging() {
        let document = crate::RtfDocument::parse(
            r#"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst DATA "recipients source.csv" "headers source.csv" \\* MERGEFORMAT}{\fldrslt cached mail-merge source}}{\field{\*\fldinst data recipients.csv}{\fldrslt cached bare source}}After}"#,
        )
        .unwrap();

        let fields = document.mail_merge_data_fields();
        assert_eq!(document.mail_merge_data_field_count(), 2);
        assert_eq!(fields.len(), 2);
        assert_eq!(fields[0].data_source(), "recipients source.csv");
        assert_eq!(fields[0].header_source(), Some("headers source.csv"));
        assert_eq!(fields[0].cached_result(), Some("cached mail-merge source"));
        assert!(fields[0].is_dirty());
        assert!(fields[0].is_locked());
        assert_eq!(fields[0].switches()[0].name, "*");
        assert_eq!(fields[0].switches()[0].value.as_deref(), Some("MERGEFORMAT"));
        assert_eq!(fields[1].data_source(), "recipients.csv");
        assert_eq!(fields[1].header_source(), None);
        assert_eq!(fields[1].cached_result(), Some("cached bare source"));
        assert_eq!(document.text(), "Before After");
    }

    #[test]
    fn document_discovers_mail_merge_counters_without_merging() {
        let document = crate::RtfDocument::parse(
            r"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst MERGEREC}{\fldrslt 12}}Middle {\field{\*\fldinst mergeSEQ}{\fldrslt 3}}After}",
        )
        .unwrap();

        let counters = document.mail_merge_counters();
        assert_eq!(document.mail_merge_counter_count(), 2);
        assert_eq!(counters.len(), 2);
        assert_eq!(counters[0].kind(), MailMergeCounterKind::Record);
        assert_eq!(counters[0].cached_result(), Some("12"));
        assert!(counters[0].is_dirty());
        assert!(counters[0].is_locked());
        assert_eq!(counters[1].kind(), MailMergeCounterKind::Sequence);
        assert_eq!(counters[1].cached_result(), Some("3"));
        assert_eq!(document.text(), "Before Middle After");
    }

    #[test]
    fn document_discovers_mail_merge_next_fields_without_advancing_records() {
        let document = crate::RtfDocument::parse(
            r"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst NEXT}{\fldrslt cached next}}After}",
        )
        .unwrap();

        let fields = document.mail_merge_next_fields();
        assert_eq!(document.mail_merge_next_field_count(), 1);
        assert_eq!(fields.len(), 1);
        assert_eq!(fields[0].cached_result(), Some("cached next"));
        assert!(fields[0].is_dirty());
        assert!(fields[0].is_locked());
        assert_eq!(document.text(), "Before After");
    }

    #[test]
    fn document_discovers_conditional_mail_merge_controls_without_merging() {
        let document = crate::RtfDocument::parse(
            r#"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst NEXTIF Customer = "Ada"}{\fldrslt cached nextif}}Middle {\field{\*\fldinst skipif MERGEFIELD Order < 100}{\fldrslt cached skipif}}After}"#,
        )
        .unwrap();

        let controls = document.mail_merge_conditional_controls();
        assert_eq!(document.mail_merge_conditional_control_count(), 2);
        assert_eq!(controls.len(), 2);
        assert_eq!(controls[0].kind(), MailMergeConditionalControlKind::NextIf);
        assert_eq!(controls[0].comparison(), r#"Customer = "Ada""#);
        assert_eq!(controls[0].cached_result(), Some("cached nextif"));
        assert!(controls[0].is_dirty());
        assert!(controls[0].is_locked());
        assert_eq!(controls[1].kind(), MailMergeConditionalControlKind::SkipIf);
        assert_eq!(controls[1].comparison(), "MERGEFIELD Order < 100");
        assert_eq!(controls[1].cached_result(), Some("cached skipif"));
        assert_eq!(document.text(), "Before Middle After");
    }

    #[test]
    fn document_discovers_if_fields_without_evaluation() {
        let document = crate::RtfDocument::parse(
            r#"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst IF "A" = "A" "yes" "no"}{\fldrslt yes}}After}"#,
        )
        .unwrap();

        let fields = document.if_fields();
        assert_eq!(document.if_field_count(), 1);
        assert_eq!(fields.len(), 1);
        assert_eq!(fields[0].expression(), r#""A" = "A" "yes" "no""#);
        assert_eq!(fields[0].cached_result(), Some("yes"));
        assert!(fields[0].is_dirty());
        assert!(fields[0].is_locked());
        assert_eq!(document.text(), "Before After");
    }

    #[test]
    fn document_discovers_compare_fields_without_evaluation() {
        let document = crate::RtfDocument::parse(
            r#"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst COMPARE "CustomerNumber" >= 4}{\fldrslt 1}}Middle {\field{\*\fldinst compare MERGEFIELD CustomerRating <= 9}{\fldrslt 0}}After}"#,
        )
        .unwrap();

        let fields = document.compare_fields();
        assert_eq!(document.compare_field_count(), 2);
        assert_eq!(fields.len(), 2);
        assert_eq!(fields[0].comparison(), r#""CustomerNumber" >= 4"#);
        assert_eq!(fields[0].cached_result(), Some("1"));
        assert!(fields[0].is_dirty());
        assert!(fields[0].is_locked());
        assert_eq!(fields[1].comparison(), "MERGEFIELD CustomerRating <= 9");
        assert_eq!(fields[1].cached_result(), Some("0"));
        assert_eq!(document.text(), "Before Middle After");
    }

    #[test]
    fn document_discovers_prompt_fields_without_displaying_prompts() {
        let document = crate::RtfDocument::parse(
            r#"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst ASK AskResponse "What is your first name?" \\d "" \\o}{\fldrslt cached ask response}}Middle {\field{\*\fldinst FILLIN "Enter appointment time" \\d "09:00"}{\fldrslt 10:30}}After}"#,
        )
        .unwrap();

        let fields = document.prompt_fields();
        assert_eq!(document.prompt_field_count(), 2);
        assert_eq!(fields.len(), 2);
        assert_eq!(fields[0].kind(), PromptFieldKind::Ask);
        assert_eq!(fields[0].bookmark(), Some("AskResponse"));
        assert_eq!(fields[0].prompt(), Some("What is your first name?"));
        assert_eq!(fields[0].default_response(), Some(""));
        assert!(fields[0].prompts_once_per_mail_merge());
        assert_eq!(fields[0].cached_result(), Some("cached ask response"));
        assert!(fields[0].is_dirty());
        assert!(fields[0].is_locked());
        assert_eq!(fields[1].kind(), PromptFieldKind::FillIn);
        assert_eq!(fields[1].bookmark(), None);
        assert_eq!(fields[1].prompt(), Some("Enter appointment time"));
        assert_eq!(fields[1].default_response(), Some("09:00"));
        assert_eq!(fields[1].cached_result(), Some("10:30"));
        assert_eq!(document.text(), "Before Middle After");
    }

    #[test]
    fn document_discovers_user_identity_fields_without_reading_host_identity() {
        let document = crate::RtfDocument::parse(
            r#"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst USERADDRESS "10 Top Secret Lane" \\* Upper}{\fldrslt 10 TOP SECRET LANE}}Middle {\field{\*\fldinst USERINITIALS \\* Lower}{\fldrslt dw}}After {\field{\*\fldinst USERNAME "Ada Lovelace" \\* FirstCap}{\fldrslt Ada Lovelace}}}"#,
        )
        .unwrap();

        let fields = document.user_identity_fields();
        assert_eq!(document.user_identity_field_count(), 3);
        assert_eq!(fields.len(), 3);
        assert_eq!(fields[0].kind(), UserIdentityFieldKind::Address);
        assert_eq!(fields[0].override_value(), Some("10 Top Secret Lane"));
        assert_eq!(fields[0].formatting(), Some(UserIdentityFormatting::Upper));
        assert_eq!(fields[0].cached_result(), Some("10 TOP SECRET LANE"));
        assert!(fields[0].is_dirty());
        assert!(fields[0].is_locked());
        assert_eq!(fields[1].kind(), UserIdentityFieldKind::Initials);
        assert_eq!(fields[1].override_value(), None);
        assert_eq!(fields[1].formatting(), Some(UserIdentityFormatting::Lower));
        assert_eq!(fields[1].cached_result(), Some("dw"));
        assert_eq!(fields[2].kind(), UserIdentityFieldKind::Name);
        assert_eq!(fields[2].override_value(), Some("Ada Lovelace"));
        assert_eq!(
            fields[2].formatting(),
            Some(UserIdentityFormatting::FirstCap)
        );
        assert_eq!(fields[2].cached_result(), Some("Ada Lovelace"));
        assert_eq!(document.text(), "Before Middle After ");
    }

    #[test]
    fn document_discovers_advance_fields_without_changing_layout() {
        let document = crate::RtfDocument::parse(
            r#"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst ADVANCE \\u 6 \\d 12 \\l 20 \\r -4 \\x 150 \\y 72}{\fldrslt cached placement}}After}"#,
        )
        .unwrap();

        let fields = document.advance_fields();
        assert_eq!(document.advance_field_count(), 1);
        assert_eq!(fields.len(), 1);
        let adjustments = fields[0]
            .adjustments()
            .iter()
            .map(|adjustment| (adjustment.operation(), adjustment.points()))
            .collect::<Vec<_>>();
        assert_eq!(
            adjustments,
            vec![
                (AdvanceFieldOperation::Up, 6),
                (AdvanceFieldOperation::Down, 12),
                (AdvanceFieldOperation::Left, 20),
                (AdvanceFieldOperation::Right, -4),
                (AdvanceFieldOperation::HorizontalPosition, 150),
                (AdvanceFieldOperation::VerticalPosition, 72),
            ]
        );
        assert_eq!(fields[0].cached_result(), Some("cached placement"));
        assert!(fields[0].is_dirty());
        assert!(fields[0].is_locked());
        assert_eq!(document.text(), "Before After");
    }

    #[test]
    fn document_discovers_recipient_fields_without_merging() {
        let document = crate::RtfDocument::parse(
            r#"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst ADDRESSBLOCK \\c 2 \\d \\e "United States" \\e Canada \\f "<<_FIRST0_>> <<_LAST0_>>" \\l 1033}{\fldrslt cached address}}Middle {\field{\*\fldinst GREETINGLINE \\f "Dear <<_FIRST0_>>," \\e "To Whom It May Concern" \\l en-US}{\fldrslt Dear Ada,}}After}"#,
        )
        .unwrap();

        let fields = document.mail_merge_recipient_fields();
        assert_eq!(document.mail_merge_recipient_field_count(), 2);
        assert_eq!(fields.len(), 2);
        assert_eq!(fields[0].kind(), MailMergeRecipientFieldKind::AddressBlock);
        assert_eq!(
            fields[0].country_inclusion(),
            Some(AddressBlockCountryInclusion::UnlessExcluded)
        );
        assert!(fields[0].formats_using_recipient_country());
        assert_eq!(fields[0].language(), Some("1033"));
        assert_eq!(fields[0].cached_result(), Some("cached address"));
        assert!(fields[0].is_dirty());
        assert!(fields[0].is_locked());
        assert_eq!(fields[1].kind(), MailMergeRecipientFieldKind::GreetingLine);
        assert_eq!(
            fields[1].greeting_fallback_text(),
            Some("To Whom It May Concern")
        );
        assert_eq!(fields[1].cached_result(), Some("Dear Ada,"));
        assert_eq!(document.text(), "Before Middle After");
    }

    #[test]
    fn document_discovers_bibliography_fields_without_loading_sources() {
        let document = crate::RtfDocument::parse(
            r#"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst CITATION Ecma01 \\l 1033 \\n \\m Ecma02}{\fldrslt cached citation}}Middle {\field{\*\fldinst BIBLIOGRAPHY \\l 1033 \\f en-US \\m Ecma01}{\fldrslt cached bibliography}}After}"#,
        )
        .unwrap();

        let citations = document.citations();
        assert_eq!(document.citation_count(), 1);
        assert_eq!(citations.len(), 1);
        assert_eq!(citations[0].source_tag(), "Ecma01");
        assert_eq!(
            citations[0].options(),
            &[
                CitationOption::LanguageId(Cow::Borrowed("1033")),
                CitationOption::SuppressAuthor,
                CitationOption::AdditionalSourceTag(Cow::Borrowed("Ecma02")),
            ]
        );
        assert_eq!(citations[0].cached_result(), Some("cached citation"));
        assert!(citations[0].is_dirty());
        assert!(citations[0].is_locked());

        let bibliographies = document.bibliographies();
        assert_eq!(document.bibliography_count(), 1);
        assert_eq!(bibliographies.len(), 1);
        assert_eq!(
            bibliographies[0].options(),
            &[
                BibliographyOption::LanguageId(Cow::Borrowed("1033")),
                BibliographyOption::FilterLanguageId(Cow::Borrowed("en-US")),
                BibliographyOption::SourceTag(Cow::Borrowed("Ecma01")),
            ]
        );
        assert_eq!(
            bibliographies[0].cached_result(),
            Some("cached bibliography")
        );
        assert_eq!(document.text(), "Before Middle After");
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
    fn document_discovers_index_fields_without_generating_an_index() {
        let document = crate::RtfDocument::parse(
            r#"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst INDEX \\f Intro \\r}{\fldrslt cached index}}Middle {\field\flddirty{\*\fldinst XE "syntax:fields" \\f Intro \\t "See references"}{\fldrslt cached entry}}After}"#,
        )
        .unwrap();

        let indexes = document.indexes();
        assert_eq!(document.index_count(), 1);
        assert_eq!(indexes.len(), 1);
        assert_eq!(indexes[0].cached_result(), Some("cached index"));
        assert!(indexes[0].is_dirty());
        assert!(indexes[0].is_locked());
        assert_eq!(
            indexes[0].options(),
            &[
                IndexOption::EntryType(Cow::Borrowed("Intro")),
                IndexOption::RunIn,
            ]
        );

        let entries = document.index_entries();
        assert_eq!(document.index_entry_count(), 1);
        assert_eq!(entries.len(), 1);
        assert_eq!(entries[0].entry(), "syntax:fields");
        assert_eq!(entries[0].cached_result(), Some("cached entry"));
        assert!(entries[0].is_dirty());
        assert!(!entries[0].is_locked());
        assert_eq!(
            entries[0].options(),
            &[
                IndexEntryOption::EntryType(Cow::Borrowed("Intro")),
                IndexEntryOption::CrossReference(Cow::Borrowed("See references")),
            ]
        );
        assert_eq!(document.text(), "Before Middle After");
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
    fn document_discovers_go_to_buttons_without_activating_them() {
        let document = crate::RtfDocument::parse(
            r#"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst GOTOBUTTON MyBookmark "Jump here"}{\fldrslt cached button}}After}"#,
        )
        .unwrap();

        let buttons = document.go_to_buttons();
        assert_eq!(document.go_to_button_count(), 1);
        assert_eq!(buttons.len(), 1);
        assert_eq!(buttons[0].target(), "MyBookmark");
        assert_eq!(buttons[0].button_text(), "Jump here");
        assert_eq!(buttons[0].cached_result(), Some("cached button"));
        assert!(buttons[0].is_dirty());
        assert!(buttons[0].is_locked());
        assert_eq!(document.text(), "Before After");
    }

    #[test]
    fn document_discovers_active_content_fields_without_activation() {
        let document = crate::RtfDocument::parse(
            r#"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst ADDIN opaque-add-in-data}{\fldrslt cached add-in result}}{\field{\*\fldinst CONTROL opaque-ocx-metadata}{\fldrslt cached control result}}{\field{\*\fldinst HTMLCONTROL opaque-html-control-metadata}{\fldrslt cached html result}}After}"#,
        )
        .unwrap();

        let fields = document.active_content_fields();
        assert_eq!(document.active_content_field_count(), 3);
        assert_eq!(fields.len(), 3);
        assert_eq!(fields[0].kind(), ActiveContentFieldKind::AddIn);
        assert_eq!(fields[0].cached_result(), Some("cached add-in result"));
        assert!(fields[0].is_dirty());
        assert!(fields[0].is_locked());
        assert_eq!(fields[1].kind(), ActiveContentFieldKind::OcxControl);
        assert_eq!(fields[1].cached_result(), Some("cached control result"));
        assert_eq!(fields[2].kind(), ActiveContentFieldKind::HtmlControl);
        assert_eq!(fields[2].cached_result(), Some("cached html result"));
        assert_eq!(document.text(), "Before After");
    }

    #[test]
    fn document_discovers_auto_text_fields_without_lookup_or_insertion() {
        let document = crate::RtfDocument::parse(
            r#"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst GLOSSARY "Legacy Clause"}{\fldrslt cached glossary entry}}{\field{\*\fldinst AUTOTEXT ReusableClause}{\fldrslt cached auto text entry}}After}"#,
        )
        .unwrap();

        let fields = document.auto_text_fields();
        assert_eq!(document.auto_text_field_count(), 2);
        assert_eq!(fields.len(), 2);
        assert_eq!(fields[0].kind(), AutoTextFieldKind::Glossary);
        assert_eq!(fields[0].entry_name(), "Legacy Clause");
        assert_eq!(fields[0].cached_result(), Some("cached glossary entry"));
        assert!(fields[0].is_dirty());
        assert!(fields[0].is_locked());
        assert_eq!(fields[1].kind(), AutoTextFieldKind::AutoText);
        assert_eq!(fields[1].entry_name(), "ReusableClause");
        assert_eq!(fields[1].cached_result(), Some("cached auto text entry"));
        assert_eq!(document.text(), "Before After");
    }

    #[test]
    fn document_discovers_auto_text_list_fields_without_selection_or_insertion() {
        let document = crate::RtfDocument::parse(
            r#"{\rtf1\ansi Before {\field\flddirty\fldlock{\*\fldinst AUTOTEXTLIST "Choose a name" \\s "Name Style" \\t "Right-click to select" \\* MERGEFORMAT \\q opaque}{\fldrslt cached selection}}{\field{\*\fldinst AUTOTEXTLIST \\s NameStyle}{\fldrslt cached style-only selection}}After}"#,
        )
        .unwrap();

        let fields = document.auto_text_list_fields();
        assert_eq!(document.auto_text_list_field_count(), 2);
        assert_eq!(fields.len(), 2);
        assert_eq!(fields[0].display_text(), Some("Choose a name"));
        assert_eq!(
            fields[0].options(),
            &[
                AutoTextListOption::Style(Cow::Borrowed("Name Style")),
                AutoTextListOption::Tip(Cow::Borrowed("Right-click to select")),
            ]
        );
        assert_eq!(fields[0].unknown_switches().len(), 2);
        assert_eq!(fields[0].cached_result(), Some("cached selection"));
        assert!(fields[0].is_dirty());
        assert!(fields[0].is_locked());
        assert_eq!(fields[1].display_text(), None);
        assert_eq!(
            fields[1].options(),
            &[AutoTextListOption::Style(Cow::Borrowed("NameStyle"))]
        );
        assert_eq!(
            fields[1].cached_result(),
            Some("cached style-only selection")
        );
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
