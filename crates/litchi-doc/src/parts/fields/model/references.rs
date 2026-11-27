use super::core::Field;
use super::mail_merge::MergeFieldSwitch;

/// One recognized stored option of a legacy Word `TOA` field.
///
/// These values retain how a producer configured a table of authorities. They
/// are metadata only: this crate never finds citations, follows bookmarks,
/// calculates page numbers, paginates, generates a table, or refreshes a
/// field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum TableOfAuthoritiesOption {
    /// The `\\b` bookmark that bounds included entries.
    Bookmark(String),
    /// The `\\c` authority category to include.
    Category(String),
    /// The `\\d` separator between sequence and page numbers.
    SequencePageSeparator(String),
    /// The `\\e` separator between an entry and its page number.
    EntryPageNumberSeparator(String),
    /// The `\\f` entry-formatting switch.
    EntryFormatting,
    /// The `\\g` separator between page numbers in a page range.
    PageRangeSeparator(String),
    /// The `\\h` switch includes category headings.
    CategoryHeadings,
    /// The `\\l` separator between multiple page references.
    PageReferenceSeparator(String),
    /// The `\\p` switch requests passim handling.
    UsePassim,
    /// The `\\s` sequence identifier whose number prefixes page numbers.
    SequenceIdentifier(String),
}

/// Typed, inert metadata for a legacy Word table-of-authorities field.
///
/// [MS-DOC] §2.9.90 maps native `TOA` field markers to ECMA-376 Part 1
/// §17.16.5.67. This type exposes only stored configuration, unrecognized
/// switches, cached result, and field state. It never finds citations, scans
/// hidden text, reads bookmarks, follows links, calculates page numbers,
/// paginates, regenerates a table of authorities, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TableOfAuthoritiesField {
    pub(in crate::parts::fields) field: Field,
    pub(in crate::parts::fields) instruction: String,
    pub(in crate::parts::fields) options: Vec<TableOfAuthoritiesOption>,
    pub(in crate::parts::fields) unknown_switches: Vec<MergeFieldSwitch>,
    pub(in crate::parts::fields) cached_result: Option<String>,
}

impl TableOfAuthoritiesField {
    /// Return the paired field markers and their story-relative positions.
    #[must_use]
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored `TOA` field instruction.
    #[must_use]
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return recognized stored configuration options in source order.
    ///
    /// This metadata is never used to generate or update a table.
    #[must_use]
    pub fn options(&self) -> &[TableOfAuthoritiesOption] {
        &self.options
    }

    /// Return unrecognized stored switches in source order.
    ///
    /// They are retained without interpretation or execution.
    #[must_use]
    pub fn unknown_switches(&self) -> &[MergeFieldSwitch] {
        &self.unknown_switches
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is never regenerated through pagination or field evaluation.
    #[must_use]
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    #[must_use]
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    #[must_use]
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}

/// One recognized stored option of a legacy Word `INDEX` field.
///
/// These values retain how a producer configured an index. They are metadata
/// only: this crate never scans index markers, reads bookmarks, calculates
/// page numbers, sorts entries, paginates, generates an index, or refreshes a
/// field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum IndexOption {
    /// The `\\b` bookmark that bounds included entries.
    Bookmark(String),
    /// The `\\c` requested number of index columns.
    Columns(String),
    /// The `\\d` separator between sequence and page numbers.
    SequencePageSeparator(String),
    /// The `\\e` separator between an entry and its first page number.
    EntryPageNumberSeparator(String),
    /// The `\\f` entry type that selects matching index markers.
    EntryType(String),
    /// The `\\g` separator between the start and end of a page range.
    PageRangeSeparator(String),
    /// The `\\h` heading text for each index-letter set.
    Heading(String),
    /// The `\\k` separator between an entry and its cross reference.
    CrossReferenceSeparator(String),
    /// The `\\l` separator between page numbers in a page-number list.
    PageNumberSeparator(String),
    /// Word's `\\o` East Asian sort-order extension, retained verbatim.
    EastAsianSortOrder(String),
    /// The `\\p` range of entry initial letters to include.
    LetterRange(String),
    /// The `\\r` switch runs subentries into their main-entry line.
    RunIn,
    /// The `\\s` sequence identifier whose number prefixes page numbers.
    SequenceIdentifier(String),
    /// The `\\y` switch enables yomi text for index entries.
    UseYomi,
    /// The `\\z` language identifier used to generate the index.
    LanguageId(String),
}

/// Typed, inert metadata for a legacy Word `INDEX` field.
///
/// [MS-DOC] §2.9.90 maps native `INDEX` field markers to ECMA-376 Part 1
/// §17.16.5.29. This type exposes only stored configuration, unrecognized
/// switches, cached result, and field state. It never scans index markers,
/// reads bookmarks, calculates page numbers, sorts entries, paginates,
/// generates an index, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct IndexField {
    pub(in crate::parts::fields) field: Field,
    pub(in crate::parts::fields) instruction: String,
    pub(in crate::parts::fields) options: Vec<IndexOption>,
    pub(in crate::parts::fields) unknown_switches: Vec<MergeFieldSwitch>,
    pub(in crate::parts::fields) cached_result: Option<String>,
}

impl IndexField {
    /// Return the paired field markers and their story-relative positions.
    #[must_use]
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored `INDEX` field instruction.
    #[must_use]
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return recognized stored configuration options in source order.
    ///
    /// This metadata is never used to sort, generate, or update an index.
    #[must_use]
    pub fn options(&self) -> &[IndexOption] {
        &self.options
    }

    /// Return unrecognized stored switches in source order.
    ///
    /// They are retained without interpretation or execution.
    #[must_use]
    pub fn unknown_switches(&self) -> &[MergeFieldSwitch] {
        &self.unknown_switches
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is never regenerated through pagination or field evaluation.
    #[must_use]
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    #[must_use]
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    #[must_use]
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}

/// The stored category of a legacy Word bookmark-reference field.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ReferenceFieldKind {
    /// A `REF` field.
    Reference,
    /// A `REF` field whose instruction omits the `REF` keyword.
    ReferenceWithoutKeyword,
    /// A `PAGEREF` field.
    PageReference,
    /// A historical `FTNREF` field.
    FootnoteReference,
    /// A `NOTEREF` field.
    NoteReference,
}

/// One recognized stored option of a legacy Word bookmark-reference field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ReferenceFieldOption {
    /// The `\\d` `REF` separator between sequence and page numbers.
    SequencePageSeparator(String),
    /// The `\\f` `REF` request for referenced note or comment content.
    ReferencedNoteContent,
    /// The `\\h` switch requests a hyperlink to the stored bookmark.
    Hyperlink,
    /// The `\\n` `REF` request for a paragraph number without context.
    ParagraphNumberWithoutContext,
    /// The `\\p` switch requests relative-position text.
    RelativePosition,
    /// The `\\r` `REF` request for a paragraph number in relative context.
    ParagraphNumberRelativeContext,
    /// The `\\t` `REF` request to suppress non-delimiter or non-numerical text.
    SuppressNonNumberText,
    /// The `\\w` `REF` request for a paragraph number in full context.
    ParagraphNumberFullContext,
    /// The `\\f` `NOTEREF` request to format the referenced note mark.
    NoteMarkFormatting,
}

/// Typed, inert metadata for a legacy Word bookmark-reference field.
///
/// [MS-DOC] §2.9.90 maps native `REF`, `PAGEREF`, `FTNREF`, and `NOTEREF`
/// field markers to ECMA-376 Part 1 §§17.16.5.51, 17.16.5.45, and
/// 17.16.5.40. This type exposes only stored category, bookmark name, options,
/// switches, cached result, and field state. It never looks up a bookmark,
/// reads a referenced range, resolves a page or note number, creates a link,
/// calculates a relative position, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ReferenceField {
    pub(in crate::parts::fields) field: Field,
    pub(in crate::parts::fields) instruction: String,
    pub(in crate::parts::fields) kind: ReferenceFieldKind,
    pub(in crate::parts::fields) bookmark: String,
    pub(in crate::parts::fields) options: Vec<ReferenceFieldOption>,
    pub(in crate::parts::fields) unknown_switches: Vec<MergeFieldSwitch>,
    pub(in crate::parts::fields) cached_result: Option<String>,
}

impl ReferenceField {
    /// Return the paired field markers and their story-relative positions.
    #[must_use]
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored field instruction.
    #[must_use]
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored reference-field category.
    #[must_use]
    pub fn kind(&self) -> ReferenceFieldKind {
        self.kind
    }

    /// Return the stored bookmark name without resolving it.
    #[must_use]
    pub fn bookmark(&self) -> &str {
        &self.bookmark
    }

    /// Return recognized stored options in source order.
    ///
    /// This metadata is never used to navigate, resolve, or activate a link.
    #[must_use]
    pub fn options(&self) -> &[ReferenceFieldOption] {
        &self.options
    }

    /// Return unrecognized stored switches in source order.
    ///
    /// They are retained without interpretation or execution.
    #[must_use]
    pub fn unknown_switches(&self) -> &[MergeFieldSwitch] {
        &self.unknown_switches
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is never regenerated by resolving a bookmark or page number.
    #[must_use]
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    #[must_use]
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    #[must_use]
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}

/// Typed, inert metadata for a legacy Word `SET` field.
///
/// [MS-DOC] §2.9.90 maps native `SET` field markers to ECMA-376 Part 1
/// §17.16.5.57. This type exposes only the stored target name, opaque
/// expression, cached result, and field state. It never evaluates the
/// expression, looks up or changes a bookmark, changes document state, or
/// refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SetField {
    pub(in crate::parts::fields) field: Field,
    pub(in crate::parts::fields) instruction: String,
    pub(in crate::parts::fields) target_name: String,
    pub(in crate::parts::fields) expression: String,
    pub(in crate::parts::fields) cached_result: Option<String>,
}

impl SetField {
    /// Return the paired field markers and their story-relative positions.
    #[must_use]
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored field instruction.
    #[must_use]
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored target name without looking it up or changing it.
    #[must_use]
    pub fn target_name(&self) -> &str {
        &self.target_name
    }

    /// Return the opaque stored expression text.
    ///
    /// This text is never parsed, evaluated, or used to change document state.
    #[must_use]
    pub fn expression(&self) -> &str {
        &self.expression
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is never regenerated by evaluating the expression.
    #[must_use]
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    #[must_use]
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    #[must_use]
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}

/// Typed, inert metadata for a legacy Word `=` (formula) field.
///
/// [MS-DOC] §2.9.90 maps native `=` field markers to ECMA-376 Part 1
/// §17.16.3.3. This type exposes only the stored optional formula, cached
/// result, and field state. It never parses or evaluates a formula, reads table
/// cells or bookmarks, resolves field values, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FormulaField {
    pub(in crate::parts::fields) field: Field,
    pub(in crate::parts::fields) instruction: String,
    pub(in crate::parts::fields) formula: Option<String>,
    pub(in crate::parts::fields) cached_result: Option<String>,
}

impl FormulaField {
    /// Return the paired field markers and their story-relative positions.
    #[must_use]
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored field instruction.
    #[must_use]
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the opaque stored formula text after the leading `=`, if present.
    ///
    /// This text is never parsed, evaluated, or used to read table cells,
    /// bookmarks, or field values.
    #[must_use]
    pub fn formula(&self) -> Option<&str> {
        self.formula.as_deref()
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is never regenerated by evaluating a formula.
    #[must_use]
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    #[must_use]
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    #[must_use]
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}

/// Typed, inert metadata for a legacy Word `EQ` equation field.
///
/// [MS-DOC] §2.9.90 maps native `EQ` field markers to ECMA-376 Part 4
/// §14.10.4.6. This type exposes only the stored opaque equation expression,
/// cached result, and field state. It never parses, calculates, formats,
/// renders, or refreshes an equation.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct EquationField {
    pub(in crate::parts::fields) field: Field,
    pub(in crate::parts::fields) instruction: String,
    pub(in crate::parts::fields) expression: String,
    pub(in crate::parts::fields) cached_result: Option<String>,
}

impl EquationField {
    /// Return the paired field markers and their story-relative positions.
    #[must_use]
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored field instruction.
    #[must_use]
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the opaque equation expression after the `EQ` keyword.
    ///
    /// This syntax is never parsed, calculated, formatted, or rendered.
    #[must_use]
    pub fn expression(&self) -> &str {
        &self.expression
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is never regenerated from equation syntax.
    #[must_use]
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    #[must_use]
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    #[must_use]
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}

/// Typed, inert metadata for a legacy Word `HYPERLINK` field.
///
/// [MS-DOC] §2.9.90 maps native `HYPERLINK` field markers to ECMA-376
/// Part 1 §17.16.5.25. This type exposes only stored link metadata, cached
/// results, and field state. It never opens, resolves, follows, activates, or
/// refreshes a link.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct HyperlinkField {
    pub(in crate::parts::fields) field: Field,
    pub(in crate::parts::fields) instruction: String,
    pub(in crate::parts::fields) external_target: Option<String>,
    pub(in crate::parts::fields) bookmark: Option<String>,
    pub(in crate::parts::fields) screen_tip: Option<String>,
    pub(in crate::parts::fields) target_frame: Option<String>,
    pub(in crate::parts::fields) appends_image_map_coordinates: bool,
    pub(in crate::parts::fields) opens_new_window: bool,
    pub(in crate::parts::fields) unknown_switches: Vec<MergeFieldSwitch>,
    pub(in crate::parts::fields) cached_result: Option<String>,
}

impl HyperlinkField {
    /// Return the paired field markers and their story-relative positions.
    #[must_use]
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored field instruction.
    #[must_use]
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored external target without resolving or opening it.
    #[must_use]
    pub fn external_target(&self) -> Option<&str> {
        self.external_target.as_deref()
    }

    /// Return the stored internal bookmark target without resolving it.
    #[must_use]
    pub fn bookmark(&self) -> Option<&str> {
        self.bookmark.as_deref()
    }

    /// Return the stored screen-tip text, if present.
    ///
    /// This is metadata only and is never displayed by the library.
    #[must_use]
    pub fn screen_tip(&self) -> Option<&str> {
        self.screen_tip.as_deref()
    }

    /// Return the stored target frame, if present.
    ///
    /// This is metadata only and is never used to open a window or frame.
    #[must_use]
    pub fn target_frame(&self) -> Option<&str> {
        self.target_frame.as_deref()
    }

    /// Whether the target receives click coordinates for a server-side image map.
    ///
    /// This records producer intent only; no navigation or hit testing occurs.
    #[must_use]
    pub fn appends_image_map_coordinates(&self) -> bool {
        self.appends_image_map_coordinates
    }

    /// Whether the field requests opening its target in a new window.
    ///
    /// This records producer intent only; no window is opened.
    #[must_use]
    pub fn opens_new_window(&self) -> bool {
        self.opens_new_window
    }

    /// Return unrecognized stored switches in source order.
    ///
    /// They are retained without interpretation or execution.
    #[must_use]
    pub fn unknown_switches(&self) -> &[MergeFieldSwitch] {
        &self.unknown_switches
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is never regenerated by resolving a link.
    #[must_use]
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    #[must_use]
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    #[must_use]
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}

/// Typed, inert metadata for a legacy Word `QUOTE` field.
///
/// [MS-DOC] §2.9.90 maps native `QUOTE` field markers to ECMA-376 Part 1
/// §17.16.5.49. This type exposes only the stored text argument, switches,
/// cached result, and field state. It never interprets character codes, expands
/// nested fields, inserts text, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct QuoteField {
    pub(in crate::parts::fields) field: Field,
    pub(in crate::parts::fields) instruction: String,
    pub(in crate::parts::fields) text: String,
    pub(in crate::parts::fields) switches: Vec<MergeFieldSwitch>,
    pub(in crate::parts::fields) cached_result: Option<String>,
}

impl QuoteField {
    /// Return the paired field markers and their story-relative positions.
    #[must_use]
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored field instruction.
    #[must_use]
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored text argument without inserting or transforming it.
    #[must_use]
    pub fn text(&self) -> &str {
        &self.text
    }

    /// Return preserved switches in source order without interpreting them.
    #[must_use]
    pub fn switches(&self) -> &[MergeFieldSwitch] {
        &self.switches
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is never regenerated by inserting text.
    #[must_use]
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    #[must_use]
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    #[must_use]
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}

/// Typed, inert metadata for a legacy Word `SYMBOL` field.
///
/// [MS-DOC] §2.9.90 maps native `SYMBOL` field markers to ECMA-376 Part 1
/// §17.16.5.61. This type exposes only the stored character argument, switches,
/// cached result, and field state. It never converts a character code, looks up
/// a font, inserts a glyph, changes formatting or layout, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SymbolField {
    pub(in crate::parts::fields) field: Field,
    pub(in crate::parts::fields) instruction: String,
    pub(in crate::parts::fields) character_argument: String,
    pub(in crate::parts::fields) switches: Vec<MergeFieldSwitch>,
    pub(in crate::parts::fields) cached_result: Option<String>,
}

impl SymbolField {
    /// Return the paired field markers and their story-relative positions.
    #[must_use]
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored field instruction.
    #[must_use]
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored character argument without converting it to a glyph.
    #[must_use]
    pub fn character_argument(&self) -> &str {
        &self.character_argument
    }

    /// Return preserved switches in source order without interpreting them.
    #[must_use]
    pub fn switches(&self) -> &[MergeFieldSwitch] {
        &self.switches
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is never regenerated by mapping a character code or inserting
    /// a glyph.
    #[must_use]
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    #[must_use]
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    #[must_use]
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}
