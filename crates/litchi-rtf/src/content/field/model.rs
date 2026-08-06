use std::borrow::Cow;

use super::codec::*;
use super::validation::{push_story_page_break, validate_story_events};
use super::{MAX_GENERIC_FIELDS, MAX_INSTRUCTION_LEN};

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
    DisplayBarcode,
    MergeBarcode,
    BidiOutline,
    Shape,
    FormText,
    FormCheckbox,
    FormDropdown,
    Private,
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
    Include,
    Import,
    IncludeText,
    IncludePicture,
    ReferencedDocument,
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
    Database,
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

/// A zero-width explicit `\column` control at a UTF-8 story boundary.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ColumnBreak {
    pub position: usize,
}

impl ColumnBreak {
    pub const fn new(position: usize) -> Self {
        Self { position }
    }
}

/// Kind of a nonrequired (soft) break marker.
///
/// RTF 1.9.1 defines `\softpage`, `\softcol`, `\softline`, and
/// `\softlheightN`, which record where the producer's layout broke a page,
/// column, or line (emitted as they appear in Galley view). The markers are
/// passive layout metadata only.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SoftBreakKind {
    /// Nonrequired page break (`\softpage`).
    Page,
    /// Nonrequired column break (`\softcol`).
    Column,
    /// Nonrequired line break (`\softline`).
    Line,
    /// Nonrequired line height in twips (`\softlheightN`).
    LineHeight(i32),
}

/// A zero-width nonrequired break marker at a UTF-8 story boundary.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SoftBreak {
    /// Marker kind.
    pub kind: SoftBreakKind,
    /// UTF-8 byte offset in the document body text.
    pub position: usize,
}

impl SoftBreak {
    pub const fn new(kind: SoftBreakKind, position: usize) -> Self {
        Self { kind, position }
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
    ColumnBreak(ColumnBreak),
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
    CustomXmlOpen(usize),
    CustomXmlClose(usize),
    MathZone(usize),
    ProtectionRangeStart(usize),
    ProtectionRangeEnd(usize),
    EditableRegionStart(usize),
    EditableRegionEnd(usize),
    SoftBreak(SoftBreak),
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

/// Inert metadata for an RTF `HYPERLINK` field.
///
/// Targets, bookmarks, display options, cached results, and state are exposed
/// as stored metadata only. This crate never resolves, opens, fetches,
/// validates, or activates a target; changes the insertion point; or refreshes
/// a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct HyperlinkField<'a> {
    instruction: &'a str,
    external_target: Option<Cow<'a, str>>,
    bookmark: Option<Cow<'a, str>>,
    screen_tip: Option<Cow<'a, str>>,
    target_frame: Option<Cow<'a, str>>,
    coordinates: Option<Cow<'a, str>>,
    new_window: bool,
    unknown_switches: Vec<FieldSwitch<'a>>,
    cached_result: Option<&'a str>,
    status: FieldStatus,
    owner: FieldOwner,
    position: usize,
}

/// The stored kind of an RTF cross-reference field.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ReferenceFieldKind {
    /// A `REF` field.
    Reference,
    /// A `PAGEREF` field.
    PageReference,
    /// A `NOTEREF` field.
    NoteReference,
}

/// Inert metadata for an RTF `REF`, `PAGEREF`, or `NOTEREF` field.
///
/// Bookmark names, switches, cached results, and state are exposed as stored
/// metadata only. This crate never looks up a bookmark, calculates a page or
/// note number, inserts text, changes layout, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ReferenceField<'a> {
    instruction: &'a str,
    kind: ReferenceFieldKind,
    bookmark: Cow<'a, str>,
    hyperlink: bool,
    position: bool,
    footnote_mark: bool,
    unknown_switches: Vec<FieldSwitch<'a>>,
    cached_result: Option<&'a str>,
    status: FieldStatus,
    owner: FieldOwner,
    position_in_story: usize,
}

/// Inert metadata for a legacy RTF `EQ` field.
///
/// The expression is retained exactly as field-instruction text after the
/// `EQ` keyword. On request, [`EquationField::model`] parses it into the
/// typed, inert [`crate::EquationModel`] syntax tree; it is never evaluated,
/// typeset, rendered, or sent to an external application.
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

/// The stored kind of a modern barcode display field.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BarcodeDisplayFieldKind {
    /// A `DISPLAYBARCODE` field with directly stored barcode data.
    DisplayBarcode,
    /// A `MERGEBARCODE` field with a stored mail-merge data-field name.
    MergeBarcode,
}

/// Inert metadata for an RTF `DISPLAYBARCODE` or `MERGEBARCODE` field.
///
/// Data arguments, barcode types, switches, cached results, and state are
/// exposed as stored metadata only. This crate never validates barcode data or
/// symbology; resolves a mail-merge data field; generates or renders a
/// barcode; or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct BarcodeDisplayField<'a> {
    instruction: &'a str,
    kind: BarcodeDisplayFieldKind,
    data_argument: Cow<'a, str>,
    barcode_type: Cow<'a, str>,
    switches: Vec<FieldSwitch<'a>>,
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

/// Inert metadata for a legacy RTF `PRIVATE` field.
///
/// Word uses this field to preserve conversion data from another file format.
/// This type retains opaque instruction text, a stored result, and field state
/// only. It never converts a document, interprets or reveals hidden content,
/// changes layout, or refreshes a field. Despite its name, `PRIVATE` is not a
/// confidentiality mechanism.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PrivateField<'a> {
    instruction: &'a str,
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
    /// An `INCLUDETEXT` or historical `INCLUDE` field that refers to document text
    /// and graphics.
    Text,
    /// An `INCLUDEPICTURE` or historical `IMPORT` field that refers to a
    /// graphic.
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
/// This represents `INCLUDETEXT`/`INCLUDEPICTURE` and historical
/// `INCLUDE`/`IMPORT` field instructions.
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

/// Inert metadata for a legacy RTF `RD` referenced-document field.
///
/// This model retains the stored source identifier, relative-path request,
/// switches, cached result, and field state only. It never opens, resolves,
/// reads, imports, refreshes, evaluates, or executes the referenced document.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ReferencedDocumentField<'a> {
    instruction: &'a str,
    source: Cow<'a, str>,
    relative_path: bool,
    switches: Vec<FieldSwitch<'a>>,
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

    pub(super) fn from_keyword(keyword: &str) -> Option<Self> {
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

    pub(super) fn from_keyword(keyword: &str) -> Option<Self> {
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

/// Inert metadata for a legacy RTF `DATABASE` query field.
///
/// Word uses this field to query a database and insert a table. This type
/// retains opaque instruction text, a cached result, and field state only. It
/// never opens a data source or database, uses connection information, executes
/// SQL, generates or inserts a table, changes layout, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DatabaseField<'a> {
    instruction: &'a str,
    opaque_instructions: &'a str,
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

    pub(super) fn from_keyword(keyword: &str) -> Option<Self> {
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
    pub(super) operation: AdvanceFieldOperation,
    pub(super) points: i64,
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

pub(super) struct ExternalIncludeParts<'a> {
    pub(super) kind: IncludeFieldKind,
    pub(super) source: Cow<'a, str>,
    pub(super) bookmark: Option<Cow<'a, str>>,
    pub(super) suppress_nested_field_updates: bool,
    pub(super) omit_picture_data: bool,
    pub(super) options: Vec<ExternalIncludeOption<'a>>,
    pub(super) unknown_switches: Vec<FieldSwitch<'a>>,
}

pub(super) struct TableOfContentsParts<'a> {
    pub(super) options: Vec<TableOfContentsOption<'a>>,
    pub(super) unknown_switches: Vec<FieldSwitch<'a>>,
}

pub(super) struct TableOfContentsEntryParts<'a> {
    pub(super) entry: Cow<'a, str>,
    pub(super) options: Vec<TableOfContentsEntryOption<'a>>,
    pub(super) unknown_switches: Vec<FieldSwitch<'a>>,
}

pub(super) struct TableOfAuthoritiesEntryParts<'a> {
    pub(super) options: Vec<TableOfAuthoritiesEntryOption<'a>>,
    pub(super) unknown_switches: Vec<FieldSwitch<'a>>,
}

pub(super) struct TableOfAuthoritiesParts<'a> {
    pub(super) options: Vec<TableOfAuthoritiesOption<'a>>,
    pub(super) unknown_switches: Vec<FieldSwitch<'a>>,
}

pub(super) struct IndexParts<'a> {
    pub(super) options: Vec<IndexOption<'a>>,
    pub(super) unknown_switches: Vec<FieldSwitch<'a>>,
}

pub(super) struct IndexEntryParts<'a> {
    pub(super) entry: Cow<'a, str>,
    pub(super) options: Vec<IndexEntryOption<'a>>,
    pub(super) unknown_switches: Vec<FieldSwitch<'a>>,
}

pub(super) struct PromptFieldParts<'a> {
    pub(super) kind: PromptFieldKind,
    pub(super) bookmark: Option<Cow<'a, str>>,
    pub(super) prompt: Option<Cow<'a, str>>,
    pub(super) default_response: Option<Cow<'a, str>>,
    pub(super) prompts_once_per_mail_merge: bool,
}

pub(super) struct UserIdentityFieldParts<'a> {
    pub(super) kind: UserIdentityFieldKind,
    pub(super) override_value: Option<Cow<'a, str>>,
    pub(super) formatting: Option<UserIdentityFormatting>,
}

pub(super) struct MailMergeRecipientFieldParts<'a> {
    pub(super) kind: MailMergeRecipientFieldKind,
    pub(super) country_inclusion: Option<AddressBlockCountryInclusion>,
    pub(super) formats_using_recipient_country: bool,
    pub(super) excluded_countries: Vec<Cow<'a, str>>,
    pub(super) format_template: Option<Cow<'a, str>>,
    pub(super) language: Option<Cow<'a, str>>,
    pub(super) greeting_fallback_text: Option<Cow<'a, str>>,
    pub(super) unknown_switches: Vec<FieldSwitch<'a>>,
}

pub(super) struct CitationParts<'a> {
    pub(super) source_tag: Cow<'a, str>,
    pub(super) options: Vec<CitationOption<'a>>,
    pub(super) unknown_switches: Vec<FieldSwitch<'a>>,
}

pub(super) struct BibliographyParts<'a> {
    pub(super) options: Vec<BibliographyOption<'a>>,
    pub(super) unknown_switches: Vec<FieldSwitch<'a>>,
}

pub(super) struct DocumentVariableFieldParts<'a> {
    pub(super) variable_name: Cow<'a, str>,
    pub(super) unknown_switches: Vec<FieldSwitch<'a>>,
}

pub(super) struct DocumentPropertyFieldParts<'a> {
    pub(super) property_name: Cow<'a, str>,
    pub(super) switches: Vec<FieldSwitch<'a>>,
}

pub(super) struct InfoFieldParts<'a> {
    pub(super) information_type: Cow<'a, str>,
    pub(super) new_value: Option<Cow<'a, str>>,
    pub(super) switches: Vec<FieldSwitch<'a>>,
}

pub(super) struct DocumentInformationFieldParts<'a> {
    pub(super) kind: DocumentInformationFieldKind,
    pub(super) switches: Vec<FieldSwitch<'a>>,
}

pub(super) struct DocumentContextFieldParts<'a> {
    pub(super) kind: DocumentContextFieldKind,
    pub(super) switches: Vec<FieldSwitch<'a>>,
}

pub(super) struct AutoNumberFieldParts<'a> {
    pub(super) kind: AutoNumberFieldKind,
    pub(super) switches: Vec<FieldSwitch<'a>>,
}

pub(super) struct ListNumberFieldParts<'a> {
    pub(super) list_name: Option<Cow<'a, str>>,
    pub(super) switches: Vec<FieldSwitch<'a>>,
}

pub(super) struct MergeFieldParts<'a> {
    pub(super) field_name: Cow<'a, str>,
    pub(super) switches: Vec<FieldSwitch<'a>>,
}

pub(super) struct ReferencedDocumentFieldParts<'a> {
    pub(super) source: Cow<'a, str>,
    pub(super) relative_path: bool,
    pub(super) switches: Vec<FieldSwitch<'a>>,
}

pub(super) struct MailMergeDataFieldParts<'a> {
    pub(super) data_source: Cow<'a, str>,
    pub(super) header_source: Option<Cow<'a, str>>,
    pub(super) switches: Vec<FieldSwitch<'a>>,
}

pub(super) struct DdeFieldParts<'a> {
    pub(super) kind: DdeFieldKind,
    pub(super) application: Cow<'a, str>,
    pub(super) source: Cow<'a, str>,
    pub(super) item: Option<Cow<'a, str>>,
    pub(super) automatic_updates: bool,
    pub(super) representation: Option<DdeRepresentation>,
    pub(super) omit_graphic_data: bool,
    pub(super) unknown_switches: Vec<FieldSwitch<'a>>,
}

pub(super) struct LinkFieldParts<'a> {
    pub(super) application_type: Cow<'a, str>,
    pub(super) source: Cow<'a, str>,
    pub(super) item: Option<Cow<'a, str>>,
    pub(super) automatic_updates: bool,
    pub(super) result_options: Vec<LinkResultOption>,
    pub(super) formatting_modes: Vec<LinkFormatting>,
    pub(super) unknown_switches: Vec<FieldSwitch<'a>>,
}

impl<'a> HyperlinkField<'a> {
    /// Return the complete stored `HYPERLINK` field instruction.
    pub fn instruction(&self) -> &'a str {
        self.instruction
    }

    /// Return the external target, if one was stored.
    pub fn external_target(&self) -> Option<&str> {
        self.external_target.as_deref()
    }

    /// Return the local bookmark target, if one was stored.
    pub fn bookmark(&self) -> Option<&str> {
        self.bookmark.as_deref()
    }

    /// Return the stored screen-tip text, if any.
    pub fn screen_tip(&self) -> Option<&str> {
        self.screen_tip.as_deref()
    }

    /// Return the stored target frame, if any.
    pub fn target_frame(&self) -> Option<&str> {
        self.target_frame.as_deref()
    }

    /// Return the stored image-map coordinates, if any.
    pub fn coordinates(&self) -> Option<&str> {
        self.coordinates.as_deref()
    }

    /// Whether the instruction requests a new window.
    pub const fn opens_in_new_window(&self) -> bool {
        self.new_window
    }

    /// Return switches not recognized by this API in source order.
    pub fn unknown_switches(&self) -> &[FieldSwitch<'a>] {
        &self.unknown_switches
    }

    /// Return the stored field result when a producer supplied one.
    ///
    /// This is cached text only and is never regenerated or activated.
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

impl<'a> ReferenceField<'a> {
    /// Return the complete stored cross-reference field instruction.
    pub fn instruction(&self) -> &'a str {
        self.instruction
    }

    /// Return the stored cross-reference field kind.
    pub const fn kind(&self) -> ReferenceFieldKind {
        self.kind
    }

    /// Return the stored bookmark name.
    pub fn bookmark(&self) -> &str {
        &self.bookmark
    }

    /// Whether the instruction requests a hyperlink result.
    pub const fn has_hyperlink(&self) -> bool {
        self.hyperlink
    }

    /// Whether the instruction requests a relative-position result.
    pub const fn includes_relative_position(&self) -> bool {
        self.position
    }

    /// Whether the instruction requests a footnote-mark result.
    pub const fn includes_footnote_mark(&self) -> bool {
        self.footnote_mark
    }

    /// Return switches not recognized by this API in source order.
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
        self.position_in_story
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

impl<'a> EquationField<'a> {
    /// Return the complete stored `EQ` field instruction.
    pub fn instruction(&self) -> &'a str {
        self.instruction
    }

    /// Return the opaque equation expression after the `EQ` keyword.
    pub fn expression(&self) -> &'a str {
        self.expression
    }

    /// Parse the expression into a typed, inert [`crate::EquationModel`].
    ///
    /// The model is purely syntactic: switch kinds, spacing values, bracket
    /// characters, and element text exactly as stored. Malformed expressions
    /// are reported as errors; nothing is evaluated or rendered.
    pub fn model(&self) -> crate::RtfResult<crate::EquationModel<'a>> {
        crate::EquationModel::parse(self.expression)
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

impl<'a> BarcodeDisplayField<'a> {
    /// Return the complete stored barcode display field instruction.
    pub fn instruction(&self) -> &'a str {
        self.instruction
    }

    /// Return whether this stores `DISPLAYBARCODE` or `MERGEBARCODE` metadata.
    pub const fn kind(&self) -> BarcodeDisplayFieldKind {
        self.kind
    }

    /// Return the stored data argument.
    ///
    /// For `DISPLAYBARCODE`, this is direct barcode data. For `MERGEBARCODE`, it
    /// is the stored mail-merge data-field name. Neither form is validated,
    /// resolved, or used to generate a barcode.
    pub fn data_argument(&self) -> &str {
        &self.data_argument
    }

    /// Return the stored barcode-type argument.
    ///
    /// This value is not validated or used to select a barcode implementation.
    pub fn barcode_type(&self) -> &str {
        &self.barcode_type
    }

    /// Return stored field-specific and formatting switches in source order.
    ///
    /// Switch values are retained without validation or interpretation.
    pub fn switches(&self) -> &[FieldSwitch<'a>] {
        &self.switches
    }

    /// Return the stored field result when a producer supplied one.
    ///
    /// This is cached text only and is never regenerated as a barcode.
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

    /// Return whether this stores a text or picture external-include field.
    ///
    /// Text includes are `INCLUDETEXT` or historical `INCLUDE` fields; picture
    /// includes are `INCLUDEPICTURE` or historical `IMPORT` fields.
    pub const fn kind(&self) -> IncludeFieldKind {
        self.kind
    }

    /// Return the stored source path or URL without resolving it.
    pub fn source(&self) -> &str {
        &self.source
    }

    /// Return the optional text-include bookmark selector.
    ///
    /// `INCLUDEPICTURE` and `IMPORT` fields do not define a bookmark operand,
    /// so they always return `None` here.
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

    /// Whether a text-include `\\!` switch suppresses nested field updates.
    ///
    /// This is stored metadata only; this crate never updates fields.
    pub const fn suppresses_nested_field_updates(&self) -> bool {
        self.suppress_nested_field_updates
    }

    /// Whether a picture-include `\\d` switch omits picture data.
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

impl<'a> ReferencedDocumentField<'a> {
    /// Return the complete stored `RD` field instruction.
    pub fn instruction(&self) -> &'a str {
        self.instruction
    }

    /// Return the stored referenced-document path without opening it.
    pub fn source(&self) -> &str {
        &self.source
    }

    /// Whether the stored `RD` instruction's `\\f` switch requests a path relative
    /// to this document.
    ///
    /// This is metadata only. The path is never resolved.
    pub const fn uses_relative_path(&self) -> bool {
        self.relative_path
    }

    /// Return all stored field switches in source order.
    ///
    /// Switches are retained without interpretation or execution.
    pub fn switches(&self) -> &[FieldSwitch<'a>] {
        &self.switches
    }

    /// Return the stored field result when a producer supplied one.
    ///
    /// This value is never regenerated by opening or updating a source.
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

impl<'a> DatabaseField<'a> {
    /// Return the complete stored `DATABASE` field instruction.
    ///
    /// This string remains opaque metadata and is never used to open a data
    /// source, database, or connection.
    pub fn instruction(&self) -> &'a str {
        self.instruction
    }

    /// Return opaque stored instruction text after `DATABASE`.
    ///
    /// It is never parsed, interpreted, or used to connect, execute SQL,
    /// generate a table, or calculate layout.
    pub fn opaque_instructions(&self) -> &'a str {
        self.opaque_instructions
    }

    /// Return the stored field result when a producer supplied one.
    ///
    /// This is cached metadata only and is never regenerated from a database
    /// query.
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

impl<'a> PrivateField<'a> {
    /// Return the complete stored `PRIVATE` field instruction.
    ///
    /// This string remains opaque metadata and is never used to convert a
    /// document or reveal hidden content.
    pub fn instruction(&self) -> &'a str {
        self.instruction
    }

    /// Return opaque stored instruction text after `PRIVATE`.
    ///
    /// It is never parsed, interpreted, or used to convert a document, reveal
    /// hidden content, or calculate layout.
    pub fn opaque_instructions(&self) -> &'a str {
        self.opaque_instructions
    }

    /// Return the stored field result when a producer supplied one.
    ///
    /// This is cached metadata only and is never regenerated by conversion or
    /// used to reveal hidden content.
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
                if keyword.eq_ignore_ascii_case("DISPLAYBARCODE") =>
            {
                FieldType::DisplayBarcode
            },
            ParsedFieldCode::Other { ref keyword, .. }
                if keyword.eq_ignore_ascii_case("MERGEBARCODE") =>
            {
                FieldType::MergeBarcode
            },
            ParsedFieldCode::Other { ref keyword, .. }
                if keyword.eq_ignore_ascii_case("BIDIOUTLINE") =>
            {
                FieldType::BidiOutline
            },
            ParsedFieldCode::Other { ref keyword, .. } if keyword.eq_ignore_ascii_case("SHAPE") => {
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
            ParsedFieldCode::Other { ref keyword, .. }
                if keyword.eq_ignore_ascii_case("PRIVATE") =>
            {
                FieldType::Private
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
                if keyword.eq_ignore_ascii_case("INCLUDE") =>
            {
                FieldType::Include
            },
            ParsedFieldCode::Other { ref keyword, .. }
                if keyword.eq_ignore_ascii_case("IMPORT") =>
            {
                FieldType::Import
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
            ParsedFieldCode::Other { ref keyword, .. } if keyword.eq_ignore_ascii_case("RD") => {
                FieldType::ReferencedDocument
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
            ParsedFieldCode::Other { ref keyword, .. }
                if keyword.eq_ignore_ascii_case("DATABASE") =>
            {
                FieldType::Database
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
            ParsedFieldCode::Other { ref keyword, .. }
                if keyword.eq_ignore_ascii_case("SYMBOL") =>
            {
                FieldType::Symbol
            },
            ParsedFieldCode::Other { ref keyword, .. }
                if AutoNumberFieldKind::from_keyword(keyword.as_ref()).is_some() =>
            {
                FieldType::AutoNumber
            },
            ParsedFieldCode::Other { ref keyword, .. }
                if keyword.eq_ignore_ascii_case("LISTNUM") =>
            {
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
    /// library parses it only on request (see [`EquationField::model`]) and
    /// never calculates, formats, or renders that syntax.
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

    /// Return inert metadata when this is a well-formed `DISPLAYBARCODE` or
    /// `MERGEBARCODE` field.
    ///
    /// Stored data arguments, barcode types, switches, cached results, and
    /// state are metadata only. This method never validates barcode data or
    /// symbology; resolves a mail-merge data field; generates or renders a
    /// barcode; or refreshes a field. Malformed instructions remain generic
    /// fields and return `None` here.
    pub fn barcode_display_field(&self) -> Option<BarcodeDisplayField<'_>> {
        let expected_kind = match self.field_type {
            FieldType::DisplayBarcode => BarcodeDisplayFieldKind::DisplayBarcode,
            FieldType::MergeBarcode => BarcodeDisplayFieldKind::MergeBarcode,
            _ => return None,
        };
        let (kind, data_argument, barcode_type, switches) =
            barcode_display_field_parts(self.instruction.as_ref())?;
        if kind != expected_kind {
            return None;
        }
        Some(BarcodeDisplayField {
            instruction: self.instruction.as_ref(),
            kind,
            data_argument,
            barcode_type,
            switches,
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

    /// Return inert metadata when this is a `PRIVATE` conversion-data field.
    ///
    /// Stored opaque instructions, cached result, and field state remain
    /// metadata only. This method never converts a document, interprets or
    /// reveals hidden content, changes layout, or refreshes a field.
    /// `PRIVATE` does not provide confidentiality semantics.
    pub fn private_field(&self) -> Option<PrivateField<'_>> {
        if self.field_type != FieldType::Private {
            return None;
        }
        let opaque_instructions = private_field_instructions(self.instruction.as_ref())?;
        Some(PrivateField {
            instruction: self.instruction.as_ref(),
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
        let expected_kind = match self.field_type {
            FieldType::Include | FieldType::IncludeText => IncludeFieldKind::Text,
            FieldType::Import | FieldType::IncludePicture => IncludeFieldKind::Picture,
            _ => return None,
        };
        let parts = external_include_parts(self.instruction.as_ref())?;
        if parts.kind != expected_kind {
            return None;
        }
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

    /// Return inert metadata when this is a well-formed `RD` referenced-document field.
    ///
    /// Stored source identifiers, relative-path requests, switches, cached
    /// results, and field state are metadata only. This method never opens,
    /// resolves, reads, imports, refreshes, evaluates, or executes a referenced
    /// document. Malformed `RD` instructions remain generic fields and return `None`
    /// here.
    pub fn referenced_document(&self) -> Option<ReferencedDocumentField<'_>> {
        if self.field_type != FieldType::ReferencedDocument {
            return None;
        }
        let parts = referenced_document_field_parts(self.instruction.as_ref())?;
        Some(ReferencedDocumentField {
            instruction: self.instruction.as_ref(),
            source: parts.source,
            relative_path: parts.relative_path,
            switches: parts.switches,
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

    /// Return inert metadata when this is a `DATABASE` query field.
    ///
    /// Stored opaque instructions, cached result, and field state remain
    /// metadata only. This method never opens a data source or database, uses
    /// connection information, executes SQL, generates or inserts a table,
    /// changes layout, or refreshes a field.
    pub fn database_field(&self) -> Option<DatabaseField<'_>> {
        if self.field_type != FieldType::Database {
            return None;
        }
        let opaque_instructions = database_field_instructions(self.instruction.as_ref())?;
        Some(DatabaseField {
            instruction: self.instruction.as_ref(),
            opaque_instructions,
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
        push_story_page_break(
            &mut self.result_events,
            self.result.as_ref(),
            position,
            "field result",
        )
    }

    pub fn clear_page_breaks(&mut self) {
        self.result_events
            .retain(|event| !matches!(event, StoryEvent::PageBreak(_)));
    }

    /// Return inert metadata when this is a well-formed `HYPERLINK` field.
    ///
    /// Targets, bookmarks, display options, cached results, and state are
    /// stored metadata only. This method never resolves, opens, fetches,
    /// validates, or activates a target; changes the insertion point; or
    /// refreshes a field. Malformed `HYPERLINK` instructions remain generic fields
    /// and return `None` here.
    pub fn hyperlink(&self) -> Option<HyperlinkField<'_>> {
        if self.field_type != FieldType::Hyperlink {
            return None;
        }
        let ParsedFieldCode::Hyperlink(code) = self.parsed_code() else {
            return None;
        };
        Some(HyperlinkField {
            instruction: self.instruction.as_ref(),
            external_target: code.external_target,
            bookmark: code.bookmark,
            screen_tip: code.screen_tip,
            target_frame: code.target_frame,
            coordinates: code.coordinates,
            new_window: code.new_window,
            unknown_switches: code.unknown_switches,
            cached_result: (!self.result.is_empty()).then_some(self.result.as_ref()),
            status: self.status,
            owner: self.owner,
            position: self.position,
        })
    }

    /// Return inert metadata when this is a well-formed `REF`, `PAGEREF`, or
    /// `NOTEREF` field.
    ///
    /// Bookmark names, switches, cached results, and state are stored metadata
    /// only. This method never looks up a bookmark, calculates a page or note
    /// number, inserts text, changes layout, or refreshes a field. Malformed
    /// instructions remain generic fields and return `None` here.
    pub fn reference_field(&self) -> Option<ReferenceField<'_>> {
        let (kind, code) = match (self.field_type, self.parsed_code()) {
            (FieldType::Reference, ParsedFieldCode::Reference(code)) => {
                (ReferenceFieldKind::Reference, code)
            },
            (FieldType::PageReference, ParsedFieldCode::PageReference(code)) => {
                (ReferenceFieldKind::PageReference, code)
            },
            (FieldType::NoteReference, ParsedFieldCode::NoteReference(code)) => {
                (ReferenceFieldKind::NoteReference, code)
            },
            _ => return None,
        };
        Some(ReferenceField {
            instruction: self.instruction.as_ref(),
            kind,
            bookmark: code.bookmark,
            hyperlink: code.hyperlink,
            position: code.position,
            footnote_mark: code.footnote_mark,
            unknown_switches: code.unknown_switches,
            cached_result: (!self.result.is_empty()).then_some(self.result.as_ref()),
            status: self.status,
            owner: self.owner,
            position_in_story: self.position,
        })
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
