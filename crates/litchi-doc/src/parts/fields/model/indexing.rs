use super::super::codec::{
    parse_index_entry_field_parts, parse_referenced_document_field_parts,
    parse_table_of_authorities_entry_field_parts, parse_table_of_contents_entry_field_parts,
    private_field_opaque_instructions,
};
use super::core::{Field, FieldStory, NonPlcfFieldText, non_plcf_field_texts};
use super::mail_merge::MergeFieldSwitch;

/// One recognized stored option of a legacy Word `TOC` field.
///
/// These values retain how a producer configured a table of contents. They
/// are metadata only: this crate never scans entries, paginates, generates a
/// table, follows links, or refreshes the field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum TableOfContentsOption {
    /// The `\\a` caption label whose item labels and numbers are omitted.
    CaptionWithoutLabel(String),
    /// The `\\b` bookmark that bounds included entries.
    Bookmark(String),
    /// The `\\c` sequence identifier for a table of captions.
    CaptionSequence(String),
    /// The `\\d` separator between sequence and page numbers.
    SequencePageSeparator(String),
    /// The `\\f` contents-entry identifier that selects entries.
    TableEntryIdentifier(String),
    /// The `\\h` switch requests hyperlinks for entries.
    Hyperlinks,
    /// The `\\l` range of contents-entry levels to include.
    TableEntryLevels(String),
    /// The `\\n` switch omits page numbers, optionally for an entry-level range.
    OmitPageNumbers(Option<String>),
    /// The `\\o` built-in heading-style range, or all used heading levels.
    HeadingStyleRange(Option<String>),
    /// The `\\p` separator between an entry and its page number.
    EntryPageNumberSeparator(String),
    /// The `\\s` sequence identifier whose number prefixes page numbers.
    SequenceIdentifier(String),
    /// The `\\t` custom style-name/contents-level mappings.
    StyleMappings(String),
    /// The `\\u` switch uses applied paragraph outline levels.
    OutlineLevels,
    /// The `\\w` switch preserves tab characters within entries.
    PreserveTabs,
    /// The `\\x` switch preserves newline characters within entries.
    PreserveNewlines,
    /// The `\\z` switch hides page numbers and leaders in Web Layout view.
    HidePageNumbersInWebLayout,
}

/// Typed, inert metadata for a legacy Word table-of-contents field.
///
/// [MS-DOC] §2.9.90 maps native `TOC` field markers to ECMA-376 Part 1
/// §17.16.5.68. This type exposes only stored configuration, unrecognized
/// switches, cached result, and field state. It never scans entries, reads
/// bookmarks, resolves links, paginates, regenerates a table of contents, or
/// refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TableOfContentsField {
    pub(in crate::parts::fields) field: Field,
    pub(in crate::parts::fields) instruction: String,
    pub(in crate::parts::fields) options: Vec<TableOfContentsOption>,
    pub(in crate::parts::fields) unknown_switches: Vec<MergeFieldSwitch>,
    pub(in crate::parts::fields) cached_result: Option<String>,
}

impl TableOfContentsField {
    /// Return the paired field markers and their story-relative positions.
    pub fn field(&self) -> &Field {
        &self.field
    }

    /// Return the complete stored `TOC` field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return recognized stored configuration options in source order.
    ///
    /// This metadata is never used to generate or update a table.
    pub fn options(&self) -> &[TableOfContentsOption] {
        &self.options
    }

    /// Return unrecognized stored switches in source order.
    ///
    /// They are retained without interpretation or execution.
    pub fn unknown_switches(&self) -> &[MergeFieldSwitch] {
        &self.unknown_switches
    }

    /// Return the stored cached field result, if present.
    ///
    /// This value is never regenerated through pagination or field evaluation.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a producer marked the stored result stale.
    pub fn is_dirty(&self) -> bool {
        self.field.end_flags.results_dirty
    }

    /// Whether a producer locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.field.end_flags.locked
    }
}

/// One recognized stored option of a legacy Word `TC` field.
///
/// These values identify how the entry participates in a table of contents.
/// They are inert metadata only: this crate never changes hidden text,
/// calculates page numbers, or generates a table.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum TableOfContentsEntryOption {
    /// The `\\f` contents-list identifier.
    ListIdentifier(String),
    /// The `\\l` entry level.
    Level(String),
    /// The `\\n` switch omits the entry page number.
    OmitPageNumber,
}

/// Typed, inert metadata for a legacy Word table-of-contents entry (`TC`)
/// field.
///
/// MS-DOC excludes `TC` field characters from the `Plcfld` `aFld` array, so
/// this type retains story-relative control-character positions instead of a
/// `Field` descriptor. It exposes only the stored entry, switches, and cached
/// result. It never changes hidden text, calculates page numbers, generates a
/// table of contents, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TableOfContentsEntryField {
    pub(in crate::parts::fields) story: FieldStory,
    pub(in crate::parts::fields) start_cp: u32,
    pub(in crate::parts::fields) separator_cp: Option<u32>,
    pub(in crate::parts::fields) end_cp: u32,
    pub(in crate::parts::fields) instruction: String,
    pub(in crate::parts::fields) entry: String,
    pub(in crate::parts::fields) options: Vec<TableOfContentsEntryOption>,
    pub(in crate::parts::fields) unknown_switches: Vec<MergeFieldSwitch>,
    pub(in crate::parts::fields) cached_result: Option<String>,
}

impl TableOfContentsEntryField {
    pub(crate) fn from_non_plcf_field(field: &NonPlcfFieldText<'_>) -> Option<Self> {
        let parts = parse_table_of_contents_entry_field_parts(field.instruction)?;
        Some(Self {
            story: field.story,
            start_cp: field.start_cp,
            separator_cp: field.separator_cp,
            end_cp: field.end_cp,
            instruction: field.instruction.to_string(),
            entry: parts.entry,
            options: parts.options,
            unknown_switches: parts.unknown_switches,
            cached_result: field.result.map(str::to_string),
        })
    }

    /// Return the story that stores this field.
    pub const fn story(&self) -> FieldStory {
        self.story
    }

    /// Return the story-relative position of this field's begin character.
    pub const fn start_position(&self) -> u32 {
        self.start_cp
    }

    /// Return the story-relative position of this field's separator character.
    ///
    /// `TC` fields normally have no cached result and therefore no separator.
    pub const fn separator_position(&self) -> Option<u32> {
        self.separator_cp
    }

    /// Return the story-relative position of this field's end character.
    pub const fn end_position(&self) -> u32 {
        self.end_cp
    }

    /// Return the complete stored `TC` field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored text marked for a table of contents.
    ///
    /// This is metadata only and is never inserted into generated content.
    pub fn entry(&self) -> &str {
        &self.entry
    }

    /// Return recognized `TC` options in stored source order.
    ///
    /// These options are never used to calculate page numbers, change hidden
    /// text, or update a table of contents.
    pub fn options(&self) -> &[TableOfContentsEntryOption] {
        &self.options
    }

    /// Return unrecognized stored field switches in source order.
    ///
    /// They are retained without interpretation or execution.
    pub fn unknown_switches(&self) -> &[MergeFieldSwitch] {
        &self.unknown_switches
    }

    /// Return the stored field result when a producer supplied one.
    ///
    /// This value is never regenerated through pagination or field evaluation.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }
}

/// One recognized stored option of a legacy Word `TA` field.
///
/// These values describe a legal-authority entry marker. They are inert
/// metadata only: this crate never finds cited text, changes hidden text,
/// follows bookmarks, calculates page numbers, or generates a table.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum TableOfAuthoritiesEntryOption {
    /// The `\\b` switch requests bold page-number formatting.
    BoldPageNumber,
    /// The `\\c` authority category.
    Category(String),
    /// The `\\i` switch requests italic page-number formatting.
    ItalicPageNumber,
    /// The `\\l` long citation text.
    LongCitation(String),
    /// The `\\r` bookmark that marks the cited page range.
    PageRangeBookmark(String),
    /// The `\\s` short citation text.
    ShortCitation(String),
}

/// Typed, inert metadata for a legacy Word table-of-authorities entry (`TA`)
/// field.
///
/// MS-DOC excludes `TA` field characters from the `Plcfld` `aFld` array, so
/// this type retains story-relative control-character positions instead of a
/// `Field` descriptor. It exposes only the stored switches and cached result.
/// It never finds citations, changes hidden text, follows bookmarks, calculates
/// page numbers, generates a table of authorities, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TableOfAuthoritiesEntryField {
    pub(in crate::parts::fields) story: FieldStory,
    pub(in crate::parts::fields) start_cp: u32,
    pub(in crate::parts::fields) separator_cp: Option<u32>,
    pub(in crate::parts::fields) end_cp: u32,
    pub(in crate::parts::fields) instruction: String,
    pub(in crate::parts::fields) options: Vec<TableOfAuthoritiesEntryOption>,
    pub(in crate::parts::fields) unknown_switches: Vec<MergeFieldSwitch>,
    pub(in crate::parts::fields) cached_result: Option<String>,
}

impl TableOfAuthoritiesEntryField {
    pub(crate) fn from_non_plcf_field(field: &NonPlcfFieldText<'_>) -> Option<Self> {
        let parts = parse_table_of_authorities_entry_field_parts(field.instruction)?;
        Some(Self {
            story: field.story,
            start_cp: field.start_cp,
            separator_cp: field.separator_cp,
            end_cp: field.end_cp,
            instruction: field.instruction.to_string(),
            options: parts.options,
            unknown_switches: parts.unknown_switches,
            cached_result: field.result.map(str::to_string),
        })
    }

    /// Return the story that stores this field.
    pub const fn story(&self) -> FieldStory {
        self.story
    }

    /// Return the story-relative position of this field's begin character.
    pub const fn start_position(&self) -> u32 {
        self.start_cp
    }

    /// Return the story-relative position of this field's separator character.
    ///
    /// `TA` fields normally have no cached result and therefore no separator.
    pub const fn separator_position(&self) -> Option<u32> {
        self.separator_cp
    }

    /// Return the story-relative position of this field's end character.
    pub const fn end_position(&self) -> u32 {
        self.end_cp
    }

    /// Return the complete stored `TA` field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return recognized `TA` options in stored source order.
    ///
    /// These options are never used to find citations, calculate page numbers,
    /// or generate a table of authorities.
    pub fn options(&self) -> &[TableOfAuthoritiesEntryOption] {
        &self.options
    }

    /// Return unrecognized stored field switches in source order.
    ///
    /// They are retained without interpretation or execution.
    pub fn unknown_switches(&self) -> &[MergeFieldSwitch] {
        &self.unknown_switches
    }

    /// Return the stored field result when a producer supplied one.
    ///
    /// This value is never regenerated through pagination or field evaluation.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }
}

/// One recognized stored option of a legacy Word `XE` field.
///
/// These values identify how an index marker participates in an `INDEX` field.
/// They are inert metadata only: this crate never changes hidden text,
/// calculates page numbers, follows bookmarks, or generates an index.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum IndexEntryOption {
    /// The `\\b` switch requests bold page-number formatting.
    BoldPageNumber,
    /// The `\\f` entry type that selects this marker.
    EntryType(String),
    /// The `\\i` switch requests italic page-number formatting.
    ItalicPageNumber,
    /// The `\\r` bookmark that marks a page range.
    PageRangeBookmark(String),
    /// The `\\t` text that replaces a page number with a cross reference.
    CrossReference(String),
    /// The `\\y` yomi sorting text.
    Yomi(String),
}

/// Typed, inert metadata for a legacy Word index-entry (`XE`) field.
///
/// MS-DOC excludes `XE` field characters from the `Plcfld` `aFld` array, so
/// this type retains story-relative control-character positions instead of a
/// `Field` descriptor. It exposes only the stored entry, switches, and cached
/// result. It never changes hidden text, resolves a bookmark, calculates page
/// numbers, sorts entries, generates an index, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct IndexEntryField {
    pub(in crate::parts::fields) story: FieldStory,
    pub(in crate::parts::fields) start_cp: u32,
    pub(in crate::parts::fields) separator_cp: Option<u32>,
    pub(in crate::parts::fields) end_cp: u32,
    pub(in crate::parts::fields) instruction: String,
    pub(in crate::parts::fields) entry: String,
    pub(in crate::parts::fields) options: Vec<IndexEntryOption>,
    pub(in crate::parts::fields) unknown_switches: Vec<MergeFieldSwitch>,
    pub(in crate::parts::fields) cached_result: Option<String>,
}

impl IndexEntryField {
    pub(crate) fn from_non_plcf_field(field: &NonPlcfFieldText<'_>) -> Option<Self> {
        let parts = parse_index_entry_field_parts(field.instruction)?;
        Some(Self {
            story: field.story,
            start_cp: field.start_cp,
            separator_cp: field.separator_cp,
            end_cp: field.end_cp,
            instruction: field.instruction.to_string(),
            entry: parts.entry,
            options: parts.options,
            unknown_switches: parts.unknown_switches,
            cached_result: field.result.map(str::to_string),
        })
    }

    /// Return the story that stores this field.
    pub const fn story(&self) -> FieldStory {
        self.story
    }

    /// Return the story-relative position of this field's begin character.
    pub const fn start_position(&self) -> u32 {
        self.start_cp
    }

    /// Return the story-relative position of this field's separator character.
    ///
    /// `XE` fields normally have no cached result and therefore no separator.
    pub const fn separator_position(&self) -> Option<u32> {
        self.separator_cp
    }

    /// Return the story-relative position of this field's end character.
    pub const fn end_position(&self) -> u32 {
        self.end_cp
    }

    /// Return the complete stored `XE` field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored text marked for an index.
    ///
    /// This is metadata only and is never inserted into generated content.
    pub fn entry(&self) -> &str {
        &self.entry
    }

    /// Return recognized `XE` options in stored source order.
    ///
    /// These options are never used to change hidden text, resolve bookmarks,
    /// calculate pages, or generate an index.
    pub fn options(&self) -> &[IndexEntryOption] {
        &self.options
    }

    /// Return unrecognized stored field switches in source order.
    ///
    /// They are retained without interpretation or execution.
    pub fn unknown_switches(&self) -> &[MergeFieldSwitch] {
        &self.unknown_switches
    }

    /// Return the stored field result when a producer supplied one.
    ///
    /// This value is never regenerated through pagination or field evaluation.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }
}

/// Typed, inert metadata for a legacy Word referenced-document (`RD`) field.
///
/// MS-DOC excludes `RD` field characters from the `Plcfld` `aFld` array, so
/// this type retains story-relative control-character positions instead of a
/// `Field` descriptor. It exposes only the stored source, relative-path request,
/// switches, and cached result. It never opens, resolves, reads, imports,
/// refreshes, evaluates, or executes the referenced document.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ReferencedDocumentField {
    pub(in crate::parts::fields) story: FieldStory,
    pub(in crate::parts::fields) start_cp: u32,
    pub(in crate::parts::fields) separator_cp: Option<u32>,
    pub(in crate::parts::fields) end_cp: u32,
    pub(in crate::parts::fields) instruction: String,
    pub(in crate::parts::fields) source: String,
    pub(in crate::parts::fields) relative_path: bool,
    pub(in crate::parts::fields) switches: Vec<MergeFieldSwitch>,
    pub(in crate::parts::fields) cached_result: Option<String>,
}

impl ReferencedDocumentField {
    pub(crate) fn from_non_plcf_field(field: &NonPlcfFieldText<'_>) -> Option<Self> {
        let parts = parse_referenced_document_field_parts(field.instruction)?;
        Some(Self {
            story: field.story,
            start_cp: field.start_cp,
            separator_cp: field.separator_cp,
            end_cp: field.end_cp,
            instruction: field.instruction.to_string(),
            source: parts.source,
            relative_path: parts.relative_path,
            switches: parts.switches,
            cached_result: field.result.map(str::to_string),
        })
    }

    /// Return the story that stores this field.
    pub const fn story(&self) -> FieldStory {
        self.story
    }

    /// Return the story-relative position of this field's begin character.
    pub const fn start_position(&self) -> u32 {
        self.start_cp
    }

    /// Return the story-relative position of this field's separator character.
    ///
    /// `RD` fields normally have no cached result and therefore no separator.
    pub const fn separator_position(&self) -> Option<u32> {
        self.separator_cp
    }

    /// Return the story-relative position of this field's end character.
    pub const fn end_position(&self) -> u32 {
        self.end_cp
    }

    /// Return the complete stored `RD` field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored referenced-document path without opening it.
    pub fn source(&self) -> &str {
        &self.source
    }

    /// Whether the stored `RD` instruction's `\\f` switch requests a path relative
    /// to this document.
    ///
    /// This is metadata only. The API never resolves the path.
    pub fn uses_relative_path(&self) -> bool {
        self.relative_path
    }

    /// Return all stored field switches in source order.
    ///
    /// They are retained without interpretation or execution.
    pub fn switches(&self) -> &[MergeFieldSwitch] {
        &self.switches
    }

    /// Return the stored field result when a producer supplied one.
    ///
    /// This value is never regenerated by opening or updating a source.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }
}

/// Typed, inert metadata for a legacy Word `PRIVATE` conversion-data field.
///
/// MS-DOC excludes `PRIVATE` field characters from the `Plcfld` `aFld` array,
/// so this type retains story-relative control-character positions instead of a
/// `Field` descriptor. It exposes only opaque stored instructions and a cached
/// result. It never converts a document, interprets field data, reveals hidden
/// content, changes layout, or refreshes a field. Despite its name, `PRIVATE` does
/// not provide confidentiality semantics.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PrivateField {
    pub(in crate::parts::fields) story: FieldStory,
    pub(in crate::parts::fields) start_cp: u32,
    pub(in crate::parts::fields) separator_cp: Option<u32>,
    pub(in crate::parts::fields) end_cp: u32,
    pub(in crate::parts::fields) instruction: String,
    pub(in crate::parts::fields) opaque_instructions: String,
    pub(in crate::parts::fields) cached_result: Option<String>,
}

impl PrivateField {
    pub(crate) fn from_non_plcf_field(field: &NonPlcfFieldText<'_>) -> Option<Self> {
        let opaque_instructions = private_field_opaque_instructions(field.instruction)?;
        Some(Self {
            story: field.story,
            start_cp: field.start_cp,
            separator_cp: field.separator_cp,
            end_cp: field.end_cp,
            instruction: field.instruction.to_string(),
            opaque_instructions,
            cached_result: field.result.map(str::to_string),
        })
    }

    /// Return the story that stores this field.
    pub const fn story(&self) -> FieldStory {
        self.story
    }

    /// Return the story-relative position of this field's begin character.
    pub const fn start_position(&self) -> u32 {
        self.start_cp
    }

    /// Return the story-relative position of this field's separator character.
    ///
    /// `PRIVATE` fields normally have no cached result and therefore no separator.
    pub const fn separator_position(&self) -> Option<u32> {
        self.separator_cp
    }

    /// Return the story-relative position of this field's end character.
    pub const fn end_position(&self) -> u32 {
        self.end_cp
    }

    /// Return the complete stored `PRIVATE` field instruction.
    ///
    /// This string remains opaque metadata and is never used to convert a
    /// document or reveal hidden content.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return opaque stored instruction text after `PRIVATE`.
    ///
    /// It is never parsed, interpreted, or used to convert a document, reveal
    /// hidden content, or calculate layout.
    pub fn opaque_instructions(&self) -> &str {
        &self.opaque_instructions
    }

    /// Return the stored field result when a producer supplied one.
    ///
    /// This value is never regenerated by conversion or used to change
    /// hidden-text visibility.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }
}

/// Typed fields whose marker characters MS-DOC excludes from every `Plcfld`.
///
/// MS-DOC §2.8.25 lists exactly five such field types: `TC`, `TA`, `XE`, `RD`,
/// and `PRIVATE`. This collection reconstructs them from balanced field
/// characters in stored story text. All values remain inert; no generated
/// table, index, referenced document, conversion payload, or cached result is
/// resolved, opened, interpreted, refreshed, or executed.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct NonPlcfFields {
    pub(in crate::parts::fields) table_of_contents_entries: Vec<TableOfContentsEntryField>,
    pub(in crate::parts::fields) table_of_authorities_entries: Vec<TableOfAuthoritiesEntryField>,
    pub(in crate::parts::fields) index_entries: Vec<IndexEntryField>,
    pub(in crate::parts::fields) referenced_documents: Vec<ReferencedDocumentField>,
    pub(in crate::parts::fields) private_fields: Vec<PrivateField>,
}

impl NonPlcfFields {
    pub(crate) fn from_story_texts<'a>(
        stories: impl IntoIterator<Item = (FieldStory, &'a str)>,
    ) -> Self {
        let mut output = Self::default();
        for (story, text) in stories {
            for field in non_plcf_field_texts(story, text) {
                if let Some(value) = TableOfContentsEntryField::from_non_plcf_field(&field) {
                    output.table_of_contents_entries.push(value);
                } else if let Some(value) =
                    TableOfAuthoritiesEntryField::from_non_plcf_field(&field)
                {
                    output.table_of_authorities_entries.push(value);
                } else if let Some(value) = IndexEntryField::from_non_plcf_field(&field) {
                    output.index_entries.push(value);
                } else if let Some(value) = ReferencedDocumentField::from_non_plcf_field(&field) {
                    output.referenced_documents.push(value);
                } else if let Some(value) = PrivateField::from_non_plcf_field(&field) {
                    output.private_fields.push(value);
                }
            }
        }
        output
    }

    /// Return stored `TC` table-of-contents entries in story and source order.
    pub fn table_of_contents_entries(&self) -> &[TableOfContentsEntryField] {
        &self.table_of_contents_entries
    }

    /// Return stored `TA` table-of-authorities entries in story and source order.
    pub fn table_of_authorities_entries(&self) -> &[TableOfAuthoritiesEntryField] {
        &self.table_of_authorities_entries
    }

    /// Return stored `XE` index entries in story and source order.
    pub fn index_entries(&self) -> &[IndexEntryField] {
        &self.index_entries
    }

    /// Return stored `RD` referenced-document fields without opening them.
    pub fn referenced_documents(&self) -> &[ReferencedDocumentField] {
        &self.referenced_documents
    }

    /// Return stored opaque `PRIVATE` conversion-data fields.
    pub fn private_fields(&self) -> &[PrivateField] {
        &self.private_fields
    }

    /// Whether no recognized excluded field is present.
    pub fn is_empty(&self) -> bool {
        self.table_of_contents_entries.is_empty()
            && self.table_of_authorities_entries.is_empty()
            && self.index_entries.is_empty()
            && self.referenced_documents.is_empty()
            && self.private_fields.is_empty()
    }

    /// Total number of recognized excluded fields.
    pub fn len(&self) -> usize {
        [
            self.table_of_contents_entries.len(),
            self.table_of_authorities_entries.len(),
            self.index_entries.len(),
            self.referenced_documents.len(),
            self.private_fields.len(),
        ]
        .into_iter()
        .fold(0usize, usize::saturating_add)
    }
}
