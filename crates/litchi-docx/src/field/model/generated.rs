//! Generated-list, table, authority, and index field models.

use super::{Field, Switch};

use crate::error::{Error, Result};

use super::super::codec::{
    field_instruction_remainder, has_field_switch, optional_field_switch_argument,
    parse_authority_category, parse_field_operand_and_switches, parse_field_switches,
    parse_index_columns, parse_index_sort_order, parse_toc_level_range,
};

use super::super::MAX_TABLE_OF_CONTENTS_ENTRY_FIELD_INSTRUCTION_BYTES;

/// An inclusive heading-level range selected by a `TOC \o` switch.
///
/// WordprocessingML heading levels are bounded to one through nine.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct TocLevelRange {
    start: u8,
    end: u8,
}

impl TocLevelRange {
    /// Create a valid inclusive heading-level range.
    pub fn new(start: u8, end: u8) -> Result<Self> {
        if !(1..=9).contains(&start) || !(1..=9).contains(&end) || start > end {
            return Err(Error::Invalid(
                "TOC heading levels must form an ascending range from 1 through 9".to_string(),
            ));
        }
        Ok(Self { start, end })
    }

    /// Return the first included heading level.
    pub fn start(&self) -> u8 {
        self.start
    }

    /// Return the final included heading level.
    pub fn end(&self) -> u8 {
        self.end
    }
}

/// A typed, inert Word table-of-contents field.
///
/// This represents the existing `TOC` field code and its cached result under
/// OOXML's field model. It deliberately does not paginate, regenerate entries,
/// resolve links, or execute field instructions. A `dirty` result simply means
/// a word processor may choose to refresh it later.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Toc {
    instruction: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
    switches: Vec<Switch>,
}

impl Toc {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some(switches) = parse_field_switches(field.instruction(), "TOC")? else {
            return Ok(None);
        };
        Ok(Some(Self {
            instruction: field.instruction.clone(),
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
            switches,
        }))
    }

    /// Return the complete field instruction exactly as parsed.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the cached visible result, if one is stored in the document.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a word processor has marked the cached result stale.
    pub fn is_dirty(&self) -> bool {
        self.dirty
    }

    /// Whether a word processor has locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.locked
    }

    /// Return the field switches in source order.
    pub fn switches(&self) -> &[Switch] {
        &self.switches
    }

    /// Check whether a case-insensitive ASCII switch appears in this field.
    pub fn has_switch(&self, name: char) -> bool {
        has_field_switch(&self.switches, name)
    }

    /// Return every built-in heading-style level range selected by `\o`.
    ///
    /// The OOXML TOC definition associates `\o` with the heading styles to
    /// include. Multiple occurrences are retained in source order rather than
    /// guessed at or collapsed.
    pub fn heading_style_levels(&self) -> Result<Vec<TocLevelRange>> {
        self.switches
            .iter()
            .filter(|switch| switch.name == 'o')
            .map(|switch| {
                let value = switch.argument.as_deref().ok_or_else(|| {
                    Error::Invalid("TOC \\o switch requires a heading-level range".to_string())
                })?;
                parse_toc_level_range(value)
            })
            .collect()
    }

    /// Whether the field asks Word to use applied paragraph outline levels.
    pub fn uses_outline_levels(&self) -> bool {
        self.has_switch('u')
    }

    /// Whether the field asks Word to emit hyperlinks for its entries.
    pub fn includes_hyperlinks(&self) -> bool {
        self.has_switch('h')
    }

    /// Whether the field hides page numbers in Web Layout view.
    pub fn hides_page_numbers_in_web_layout(&self) -> bool {
        self.has_switch('z')
    }
}

/// A typed, inert Word table-of-contents entry (`TC`) field.
///
/// A TC marker stores one entry for a table of contents or a similar list.
/// This model exposes only that stored entry, switches, and cached result. It
/// never changes hidden-text state, calculates page numbers, generates a table
/// of contents, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TocEntry {
    instruction: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
    entry: String,
    switches: Vec<Switch>,
}

impl TocEntry {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        if !field.is_table_of_contents_entry() {
            return Ok(None);
        }
        if field.instruction().len() > MAX_TABLE_OF_CONTENTS_ENTRY_FIELD_INSTRUCTION_BYTES {
            return Err(Error::Invalid(format!(
                "TC field instruction exceeds {MAX_TABLE_OF_CONTENTS_ENTRY_FIELD_INSTRUCTION_BYTES} bytes"
            )));
        }
        let Some((entry, switches)) = parse_field_operand_and_switches(field.instruction(), "TC")?
        else {
            unreachable!("table-of-contents entry recognition and parsing must agree");
        };
        let entry = entry
            .ok_or_else(|| Error::Invalid("TC field is missing its entry text".to_string()))?;
        if entry.is_empty() {
            return Err(Error::Invalid("TC field entry text is empty".to_string()));
        }
        for switch in &switches {
            match switch.name {
                'f' | 'l' if switch.argument.is_none() => {
                    return Err(Error::Invalid(format!(
                        "TC \\{} switch requires an argument",
                        switch.name
                    )));
                },
                'n' if switch.argument.is_some() => {
                    return Err(Error::Invalid(
                        "TC \\n switch does not take an argument".to_string(),
                    ));
                },
                _ => {},
            }
        }

        Ok(Some(Self {
            instruction: field.instruction.clone(),
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
            entry,
            switches,
        }))
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored text marked for a table of contents.
    ///
    /// This is metadata only and is never inserted into generated content.
    pub fn entry(&self) -> &str {
        &self.entry
    }

    /// Return the cached visible result, if a producer stored one.
    ///
    /// TC fields normally display no result, and this API never generates one.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a word processor has marked the cached result stale.
    pub fn is_dirty(&self) -> bool {
        self.dirty
    }

    /// Whether a word processor has locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.locked
    }

    /// Return the field switches in source order.
    pub fn switches(&self) -> &[Switch] {
        &self.switches
    }

    /// Check whether a case-insensitive ASCII switch appears in this field.
    pub fn has_switch(&self, name: char) -> bool {
        has_field_switch(&self.switches, name)
    }

    /// Return the optional `\\f` contents-list identifier.
    ///
    /// The identifier is preserved as stored metadata and is never used to
    /// select or generate a contents list.
    pub fn list_identifier(&self) -> Result<Option<&str>> {
        optional_field_switch_argument(&self.switches, 'f', "TC")
    }

    /// Return the optional `\\l` entry level without calculating its style.
    pub fn level(&self) -> Result<Option<&str>> {
        optional_field_switch_argument(&self.switches, 'l', "TC")
    }

    /// Whether the stored `\\n` switch omits the entry's page number.
    ///
    /// This records producer intent only; no page number is calculated or
    /// changed.
    pub fn omits_page_number(&self) -> bool {
        self.has_switch('n')
    }
}

/// A typed, inert Word table-of-authorities (`TOA`) field.
///
/// A TOA collects stored `TA` citation-marker fields into a rendered list.
/// This model exposes only the persisted code and cached result: it does not
/// find citations, paginate the document, generate authorities, or execute
/// any field instruction.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Toa {
    instruction: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
    switches: Vec<Switch>,
}

impl Toa {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some(switches) = parse_field_switches(field.instruction(), "TOA")? else {
            return Ok(None);
        };
        Ok(Some(Self {
            instruction: field.instruction.clone(),
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
            switches,
        }))
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the cached visible result, if present.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a word processor has marked the cached result stale.
    pub fn is_dirty(&self) -> bool {
        self.dirty
    }

    /// Whether a word processor has locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.locked
    }

    /// Return the field switches in source order.
    pub fn switches(&self) -> &[Switch] {
        &self.switches
    }

    /// Check whether a case-insensitive ASCII switch appears in this field.
    pub fn has_switch(&self, name: char) -> bool {
        has_field_switch(&self.switches, name)
    }

    /// Return the selected authority category, where zero means all categories.
    ///
    /// The Word TOA model bounds category values to zero through sixteen.
    pub fn category(&self) -> Result<Option<u8>> {
        optional_field_switch_argument(&self.switches, 'c', "TOA")?
            .map(|value| parse_authority_category(value, 0, "TOA"))
            .transpose()
    }

    /// Return the bookmark limiting where authority entries are collected.
    pub fn bookmark(&self) -> Result<Option<&str>> {
        optional_field_switch_argument(&self.switches, 'b', "TOA")
    }

    /// Whether five or more page references may be rendered as "Passim".
    pub fn uses_passim(&self) -> bool {
        self.has_switch('p')
    }

    /// Whether formatting stored with `TA` entries is retained in the result.
    pub fn keeps_entry_formatting(&self) -> bool {
        self.has_switch('f')
    }

    /// Return the separator between a sequence number and its page number.
    pub fn sequence_page_separator(&self) -> Result<Option<&str>> {
        optional_field_switch_argument(&self.switches, 'd', "TOA")
    }

    /// Return the `SEQ` field identifier used to number authority entries.
    pub fn sequence_name(&self) -> Result<Option<&str>> {
        optional_field_switch_argument(&self.switches, 's', "TOA")
    }

    /// Return the separator between an authority entry and its page number.
    pub fn entry_page_separator(&self) -> Result<Option<&str>> {
        optional_field_switch_argument(&self.switches, 'e', "TOA")
    }

    /// Return the separator between the endpoints of a page range.
    pub fn page_range_separator(&self) -> Result<Option<&str>> {
        optional_field_switch_argument(&self.switches, 'g', "TOA")
    }

    /// Whether category headers are included in the stored TOA configuration.
    pub fn includes_category_headers(&self) -> bool {
        self.has_switch('h')
    }

    /// Return the separator between individual page references.
    pub fn page_number_separator(&self) -> Result<Option<&str>> {
        optional_field_switch_argument(&self.switches, 'l', "TOA")
    }
}

/// A typed, inert Word table-of-authorities entry (`TA`) field.
///
/// This represents one stored citation marker. It deliberately does not search
/// document text for matching citations, alter hidden-text state, or generate a
/// table of authorities.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ToaEntry {
    instruction: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
    switches: Vec<Switch>,
}

impl ToaEntry {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some(switches) = parse_field_switches(field.instruction(), "TA")? else {
            return Ok(None);
        };
        Ok(Some(Self {
            instruction: field.instruction.clone(),
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
            switches,
        }))
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the cached visible result, if present.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a word processor has marked the cached result stale.
    pub fn is_dirty(&self) -> bool {
        self.dirty
    }

    /// Whether a word processor has locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.locked
    }

    /// Return the field switches in source order.
    pub fn switches(&self) -> &[Switch] {
        &self.switches
    }

    /// Check whether a case-insensitive ASCII switch appears in this field.
    pub fn has_switch(&self, name: char) -> bool {
        has_field_switch(&self.switches, name)
    }

    /// Return the long citation text used in a generated table of authorities.
    pub fn long_citation(&self) -> Result<Option<&str>> {
        optional_field_switch_argument(&self.switches, 'l', "TA")
    }

    /// Return the short citation stored for matching/entry selection.
    pub fn short_citation(&self) -> Result<Option<&str>> {
        optional_field_switch_argument(&self.switches, 's', "TA")
    }

    /// Return the authority category, if explicitly stored.
    ///
    /// Citation-marker categories are numbered one through sixteen.
    pub fn category(&self) -> Result<Option<u8>> {
        optional_field_switch_argument(&self.switches, 'c', "TA")?
            .map(|value| parse_authority_category(value, 1, "TA"))
            .transpose()
    }

    /// Whether the generated authority entry asks for bold formatting.
    pub fn is_bold(&self) -> bool {
        self.has_switch('b')
    }

    /// Whether the generated authority entry asks for italic formatting.
    pub fn is_italic(&self) -> bool {
        self.has_switch('i')
    }
}

/// East Asian sort order requested by Word's `INDEX \o` extension switch.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum IndexOrder {
    /// Sort within radicals by stroke count (`S`).
    Stroke,
    /// Sort by pronunciation (`P`).
    Pronunciation,
}

/// A typed, inert Word generated-index (`INDEX`) field.
///
/// This represents the stored index configuration and cached result. It never
/// searches for `XE` markers, sorts index entries, calculates pages, or updates
/// the rendered field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Index {
    instruction: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
    switches: Vec<Switch>,
}

impl Index {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some(switches) = parse_field_switches(field.instruction(), "INDEX")? else {
            return Ok(None);
        };
        Ok(Some(Self {
            instruction: field.instruction.clone(),
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
            switches,
        }))
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the cached visible result, if present.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a word processor has marked the cached result stale.
    pub fn is_dirty(&self) -> bool {
        self.dirty
    }

    /// Whether a word processor has locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.locked
    }

    /// Return the field switches in source order.
    pub fn switches(&self) -> &[Switch] {
        &self.switches
    }

    /// Check whether a case-insensitive ASCII switch appears in this field.
    pub fn has_switch(&self, name: char) -> bool {
        has_field_switch(&self.switches, name)
    }

    /// Return the bookmark limiting which portion of the document is indexed.
    pub fn bookmark(&self) -> Result<Option<&str>> {
        optional_field_switch_argument(&self.switches, 'b', "INDEX")
    }

    /// Return the requested number of index columns.
    ///
    /// Word bounds the `\c` switch to one through four columns.
    pub fn columns(&self) -> Result<Option<u8>> {
        optional_field_switch_argument(&self.switches, 'c', "INDEX")?
            .map(parse_index_columns)
            .transpose()
    }

    /// Return the separator between a sequence number and a page number.
    pub fn sequence_page_separator(&self) -> Result<Option<&str>> {
        optional_field_switch_argument(&self.switches, 'd', "INDEX")
    }

    /// Return the separator between an index entry and its page number.
    pub fn entry_page_separator(&self) -> Result<Option<&str>> {
        optional_field_switch_argument(&self.switches, 'e', "INDEX")
    }

    /// Return the entry identifier used to select matching `XE` fields.
    pub fn entry_identifier(&self) -> Result<Option<&str>> {
        optional_field_switch_argument(&self.switches, 'f', "INDEX")
    }

    /// Return the separator between the endpoints of a page range.
    pub fn page_range_separator(&self) -> Result<Option<&str>> {
        optional_field_switch_argument(&self.switches, 'g', "INDEX")
    }

    /// Return the text inserted between alphabetic groups in the index.
    pub fn alphabetic_group_heading(&self) -> Result<Option<&str>> {
        optional_field_switch_argument(&self.switches, 'h', "INDEX")
    }

    /// Return the separator between an entry and its cross-reference text.
    pub fn cross_reference_separator(&self) -> Result<Option<&str>> {
        optional_field_switch_argument(&self.switches, 'k', "INDEX")
    }

    /// Return the separator between multiple page references.
    pub fn page_reference_separator(&self) -> Result<Option<&str>> {
        optional_field_switch_argument(&self.switches, 'l', "INDEX")
    }

    /// Return Word's requested East Asian index sort order, if present.
    ///
    /// `\o` is a documented Word extension rather than a core ECMA-376 INDEX
    /// switch, so unknown values are reported as invalid instead of guessed.
    pub fn sort_order(&self) -> Result<Option<IndexOrder>> {
        optional_field_switch_argument(&self.switches, 'o', "INDEX")?
            .map(parse_index_sort_order)
            .transpose()
    }

    /// Return the alphabetic range used to restrict generated entries.
    pub fn letter_range(&self) -> Result<Option<&str>> {
        optional_field_switch_argument(&self.switches, 'p', "INDEX")
    }

    /// Whether subentries are configured to run into their main-entry line.
    pub fn runs_subentries_inline(&self) -> bool {
        self.has_switch('r')
    }

    /// Return the `SEQ` field identifier included with page numbers.
    pub fn sequence_name(&self) -> Result<Option<&str>> {
        optional_field_switch_argument(&self.switches, 's', "INDEX")
    }

    /// Whether the field enables yomi text for index entries.
    pub fn uses_yomi(&self) -> bool {
        self.has_switch('y')
    }

    /// Return the language identifier Word stores for index generation.
    ///
    /// This preserves the lexical field value rather than resolving it to a
    /// locale or changing sorting behavior.
    pub fn language_id(&self) -> Result<Option<&str>> {
        optional_field_switch_argument(&self.switches, 'z', "INDEX")
    }
}

/// A typed, inert Word index-entry (`XE`) field.
///
/// The entry text and its stored switches are available for inspection. This
/// model does not change hidden-text formatting, resolve the page-range
/// bookmark, sort entries, or generate an `INDEX` result.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct IndexEntry {
    instruction: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
    entry: String,
    switches: Vec<Switch>,
}

impl IndexEntry {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some((entry, switches)) = parse_field_operand_and_switches(field.instruction(), "XE")?
        else {
            return Ok(None);
        };
        let entry = entry.ok_or_else(|| {
            Error::Invalid("XE field is missing its index-entry text".to_string())
        })?;
        if entry.is_empty() {
            return Err(Error::Invalid(
                "XE field index-entry text is empty".to_string(),
            ));
        }
        Ok(Some(Self {
            instruction: field.instruction.clone(),
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
            entry,
            switches,
        }))
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the cached visible result, if present.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a word processor has marked the cached result stale.
    pub fn is_dirty(&self) -> bool {
        self.dirty
    }

    /// Whether a word processor has locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.locked
    }

    /// Return the text that is marked for inclusion in an index.
    pub fn entry(&self) -> &str {
        &self.entry
    }

    /// Return the field switches in source order.
    pub fn switches(&self) -> &[Switch] {
        &self.switches
    }

    /// Check whether a case-insensitive ASCII switch appears in this field.
    pub fn has_switch(&self, name: char) -> bool {
        has_field_switch(&self.switches, name)
    }

    /// Whether the marked page number is formatted in bold.
    pub fn is_bold(&self) -> bool {
        self.has_switch('b')
    }

    /// Return the entry identifier used to select an `INDEX` field.
    pub fn entry_identifier(&self) -> Result<Option<&str>> {
        optional_field_switch_argument(&self.switches, 'f', "XE")
    }

    /// Whether the marked page number is formatted in italics.
    pub fn is_italic(&self) -> bool {
        self.has_switch('i')
    }

    /// Return the bookmark marking the stored page range.
    pub fn page_range_bookmark(&self) -> Result<Option<&str>> {
        optional_field_switch_argument(&self.switches, 'r', "XE")
    }

    /// Return text substituted for a page number, such as a cross-reference.
    pub fn cross_reference(&self) -> Result<Option<&str>> {
        optional_field_switch_argument(&self.switches, 't', "XE")
    }

    /// Return yomi sort text, if the entry stores it.
    pub fn yomi(&self) -> Result<Option<&str>> {
        optional_field_switch_argument(&self.switches, 'y', "XE")
    }
}

impl Field {
    /// Check whether this is a `TOC` (Table of Contents) field.
    ///
    /// The field's cached result remains data only; calling this method never
    /// recalculates the table of contents or follows any hyperlinks in it.
    pub fn is_table_of_contents(&self) -> bool {
        field_instruction_remainder(&self.instruction, "TOC").is_some()
    }

    /// Parse this field as an inert typed table-of-contents field.
    ///
    /// Returns `Ok(None)` for non-`TOC` fields. The returned model preserves
    /// the instruction, cached result, dirty/lock state, and field switches;
    /// it never evaluates the field or refreshes its cached content.
    pub fn table_of_contents(&self) -> Result<Option<Toc>> {
        Toc::from_field(self)
    }

    /// Check whether this is a `TC` (Table of Contents Entry) field.
    ///
    /// TC markers remain stored data only. Recognizing one never changes
    /// hidden-text state, calculates a page number, or generates a table of
    /// contents.
    pub fn is_table_of_contents_entry(&self) -> bool {
        field_instruction_remainder(&self.instruction, "TC").is_some()
    }

    /// Parse this field as an inert typed table-of-contents entry marker.
    ///
    /// Returns `Ok(None)` for non-`TC` fields. The returned model preserves
    /// the stored entry text, switches, cached result, and dirty/lock state
    /// only; it never changes hidden text, calculates page numbers, generates
    /// a table of contents, or refreshes a field.
    pub fn table_of_contents_entry(&self) -> Result<Option<TocEntry>> {
        TocEntry::from_field(self)
    }

    /// Check whether this is a `TOA` (Table of Authorities) field.
    ///
    /// This only recognizes the stored field code. It never generates the
    /// table, resolves its cited authorities, or recalculates page references.
    pub fn is_table_of_authorities(&self) -> bool {
        field_instruction_remainder(&self.instruction, "TOA").is_some()
    }

    /// Parse this field as an inert typed table-of-authorities field.
    ///
    /// Returns `Ok(None)` for non-`TOA` fields. Stored switches and cached
    /// visible content are exposed as data only; no authority table is built.
    pub fn table_of_authorities(&self) -> Result<Option<Toa>> {
        Toa::from_field(self)
    }

    /// Check whether this is a `TA` (Table of Authorities Entry) field.
    ///
    /// Such fields mark stored citations for a `TOA` field. They remain inert
    /// data and are never interpreted as executable content.
    pub fn is_table_of_authorities_entry(&self) -> bool {
        field_instruction_remainder(&self.instruction, "TA").is_some()
    }

    /// Parse this field as an inert typed table-of-authorities entry field.
    ///
    /// Returns `Ok(None)` for non-`TA` fields. The result does not search for
    /// matching citations, generate a `TOA`, or refresh any cached content.
    pub fn table_of_authorities_entry(&self) -> Result<Option<ToaEntry>> {
        ToaEntry::from_field(self)
    }

    /// Check whether this is an `INDEX` field.
    ///
    /// This recognizes only the stored configuration and does not sort entries,
    /// calculate page references, or generate the index result.
    pub fn is_index(&self) -> bool {
        field_instruction_remainder(&self.instruction, "INDEX").is_some()
    }

    /// Parse this field as an inert typed generated-index field.
    ///
    /// Returns `Ok(None)` for non-`INDEX` fields. The model exposes the stored
    /// switches and cached result without regenerating or paginating an index.
    pub fn index(&self) -> Result<Option<Index>> {
        Index::from_field(self)
    }

    /// Check whether this is an `XE` (Index Entry) field.
    ///
    /// XE fields mark stored index entries. They are inspected as data only and
    /// never affect hidden text, sorting, or generated index content.
    pub fn is_index_entry(&self) -> bool {
        field_instruction_remainder(&self.instruction, "XE").is_some()
    }

    /// Parse this field as an inert typed index-entry field.
    ///
    /// Returns `Ok(None)` for non-`XE` fields. It never searches document text
    /// for entries, follows bookmarks, or updates an `INDEX` field.
    pub fn index_entry(&self) -> Result<Option<IndexEntry>> {
        IndexEntry::from_field(self)
    }
}
