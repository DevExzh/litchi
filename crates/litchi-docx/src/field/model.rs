//! Contextual, inert Word field values and typed instruction models.

use super::codec::{
    field_instruction_remainder, has_field_switch, optional_field_switch_argument,
    parse_advance_field_adjustments, parse_authority_category, parse_auto_text_field_parts,
    parse_auto_text_list_field_parts, parse_citation_operand_and_switches,
    parse_dde_operands_and_switches, parse_external_include_operands_and_switches,
    parse_field_operand_and_switches, parse_field_switches, parse_formula_field_formula,
    parse_go_to_button_operands, parse_index_columns, parse_index_sort_order,
    parse_info_field_parts, parse_link_operands_and_switches, parse_macro_button_operands,
    parse_mail_merge_data_field_parts, parse_mail_merge_recipient_field_parts,
    parse_prompt_field_parts, parse_sequence_field_parts, parse_set_field_parts,
    parse_style_reference_field_parts, parse_toc_level_range, parse_user_identity_field_parts,
    required_external_include_option_argument,
};
use super::{
    Error, MAX_AUTO_NUMBER_FIELD_INSTRUCTION_BYTES, MAX_BARCODE_FIELD_INSTRUCTION_BYTES,
    MAX_BIDI_OUTLINE_FIELD_INSTRUCTION_BYTES, MAX_DATABASE_FIELD_INSTRUCTION_BYTES,
    MAX_DOCUMENT_CONTEXT_FIELD_INSTRUCTION_BYTES, MAX_DOCUMENT_INFORMATION_FIELD_INSTRUCTION_BYTES,
    MAX_DOCUMENT_PROPERTY_FIELD_INSTRUCTION_BYTES, MAX_EMBED_FIELD_INSTRUCTION_BYTES,
    MAX_EQUATION_FIELD_INSTRUCTION_BYTES, MAX_HYPERLINK_FIELD_INSTRUCTION_BYTES,
    MAX_LEGACY_FORM_FIELD_INSTRUCTION_BYTES, MAX_LIST_NUMBER_FIELD_INSTRUCTION_BYTES,
    MAX_PRIVATE_FIELD_INSTRUCTION_BYTES, MAX_QUOTE_FIELD_INSTRUCTION_BYTES,
    MAX_REFERENCE_FIELD_INSTRUCTION_BYTES, MAX_SHAPE_FIELD_INSTRUCTION_BYTES,
    MAX_SYMBOL_FIELD_INSTRUCTION_BYTES, MAX_TABLE_OF_CONTENTS_ENTRY_FIELD_INSTRUCTION_BYTES,
    Result,
};

/// A field in a Word document.
///
/// Represents a field instruction like `PAGE`, `DATE`, `REF`, etc.
/// Fields are dynamic content that can be updated.
///
/// # Examples
///
/// ```rust,no_run
/// use litchi_docx::Field;
///
/// let field = Field::new("PAGE".to_string(), Some("1".to_string()), false);
/// println!("Field: {} = {}", field.instruction(), field.result().unwrap_or(""));
/// ```
#[derive(Debug, Clone)]
pub struct Field {
    /// The field instruction (e.g., "PAGE", "DATE \\@ \"MMMM d, yyyy\"")
    instruction: String,
    /// The field result (cached display value)
    result: Option<String>,
    /// Whether the field is dirty (needs updating)
    dirty: bool,
    /// Whether Word should prevent the field result from being recalculated.
    locked: bool,
}

impl Field {
    /// Create a new Field.
    ///
    /// # Arguments
    ///
    /// * `instruction` - The field instruction
    /// * `result` - The cached field result
    /// * `dirty` - Whether the field needs updating
    pub fn new(instruction: String, result: Option<String>, dirty: bool) -> Self {
        Self {
            instruction,
            result,
            dirty,
            locked: false,
        }
    }

    pub(super) fn with_flags(
        instruction: String,
        result: Option<String>,
        dirty: bool,
        locked: bool,
    ) -> Self {
        Self {
            instruction,
            result,
            dirty,
            locked,
        }
    }

    /// Get the field instruction.
    ///
    /// This is the field code that determines what the field displays.
    ///
    /// # Examples
    ///
    /// - `"PAGE"` - Current page number
    /// - `"DATE \\@ \"MMMM d, yyyy\""` - Formatted date
    /// - `"REF bookmark1"` - Cross-reference to a bookmark
    #[inline]
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Get the field result (cached display value).
    #[inline]
    pub fn result(&self) -> Option<&str> {
        self.result.as_deref()
    }

    /// Check if the field is dirty (needs updating).
    #[inline]
    pub fn is_dirty(&self) -> bool {
        self.dirty
    }

    /// Check if the field is locked against automatic recalculation.
    #[inline]
    pub fn is_locked(&self) -> bool {
        self.locked
    }

    /// Get the field type from the instruction.
    ///
    /// Returns the first word of the instruction, which is typically the field type.
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi_docx::Field;
    ///
    /// let field = Field::new("PAGE".to_string(), Some("1".to_string()), false);
    /// assert_eq!(field.field_type(), "PAGE");
    ///
    /// let field = Field::new("DATE \\@ \"MMMM d, yyyy\"".to_string(), None, false);
    /// assert_eq!(field.field_type(), "DATE");
    /// ```
    pub fn field_type(&self) -> &str {
        self.instruction
            .split_whitespace()
            .next()
            .unwrap_or(&self.instruction)
    }

    /// Check whether this is a `MERGEFIELD` mail-merge field.
    ///
    /// Recognition is limited to the stored field instruction. It never opens
    /// a data source, resolves a record, performs a merge, or refreshes the
    /// cached result.
    pub fn is_merge_field(&self) -> bool {
        field_instruction_remainder(&self.instruction, "MERGEFIELD").is_some()
    }

    /// Return the data-source column name from a `MERGEFIELD` instruction.
    ///
    /// Both unquoted names (`MERGEFIELD FirstName`) and quoted names containing
    /// spaces (`MERGEFIELD "Full Name"`) are supported. Field switches following
    /// the name are excluded. This legacy convenience accessor never opens a
    /// data source or performs a merge.
    pub fn merge_field_name(&self) -> Option<&str> {
        let remainder = field_instruction_remainder(&self.instruction, "MERGEFIELD")?.trim_start();
        if remainder.is_empty() || remainder.starts_with('\\') {
            return None;
        }
        if let Some(quoted) = remainder.strip_prefix('"') {
            let end = quoted.find('"')?;
            let name = &quoted[..end];
            return (!name.is_empty()).then_some(name);
        }
        let end = remainder
            .find(char::is_whitespace)
            .unwrap_or(remainder.len());
        let name = &remainder[..end];
        (!name.is_empty()).then_some(name)
    }

    /// Parse this field as inert typed `MERGEFIELD` metadata.
    ///
    /// Returns `Ok(None)` for non-`MERGEFIELD` fields. The result exposes only
    /// the stored field name, switches, cached content, and dirty/lock state;
    /// it never opens a data source, resolves records, performs a merge, or
    /// refreshes the field.
    pub fn merge_field(&self) -> Result<Option<Merge>> {
        Merge::from_field(self)
    }

    /// Check whether this is a `DATA` mail-merge source field.
    ///
    /// Recognition is limited to stored field metadata. It never opens, reads,
    /// connects to, resolves, or modifies a source; it never selects a record,
    /// performs a merge, or refreshes the field.
    pub fn is_mail_merge_data(&self) -> bool {
        field_instruction_remainder(&self.instruction, "DATA").is_some()
    }

    /// Parse this field as inert typed `DATA` mail-merge source metadata.
    ///
    /// Returns `Ok(None)` for fields other than `DATA`. The result exposes only
    /// stored source identifiers, switches, cached content, and dirty/lock
    /// state; it never opens, reads, connects to, resolves, or modifies a
    /// source, selects a record, performs a merge, or refreshes a field.
    pub fn mail_merge_data(&self) -> Result<Option<MergeData>> {
        MergeData::from_field(self)
    }

    /// Check whether this is a `MERGEREC` mail-merge counter field.
    ///
    /// Recognition is limited to the stored instruction. It never selects a
    /// record, opens a data source, performs a merge, or refreshes the result.
    pub fn is_merge_record(&self) -> bool {
        field_instruction_remainder(&self.instruction, "MERGEREC").is_some()
    }

    /// Check whether this is a `MERGESEQ` mail-merge counter field.
    ///
    /// Recognition is limited to the stored instruction. It never selects a
    /// record, opens a data source, performs a merge, or refreshes the result.
    pub fn is_merge_sequence(&self) -> bool {
        field_instruction_remainder(&self.instruction, "MERGESEQ").is_some()
    }

    /// Check whether this is a `MERGEREC` or `MERGESEQ` field.
    pub fn is_mail_merge_counter(&self) -> bool {
        self.is_merge_record() || self.is_merge_sequence()
    }

    /// Parse this field as inert typed mail-merge counter metadata.
    ///
    /// Returns `Ok(None)` for fields other than `MERGEREC` and `MERGESEQ`.
    /// The result exposes only stored kind, cached content, and dirty/lock
    /// state; it never selects a record, opens a data source, performs a
    /// merge, or refreshes a field.
    pub fn mail_merge_counter(&self) -> Result<Option<MergeCounter>> {
        MergeCounter::from_field(self)
    }

    /// Check whether this is a `NEXT` mail-merge control field.
    ///
    /// Recognition is limited to the stored instruction. It never advances a
    /// record, opens a data source, performs a merge, or refreshes the result.
    pub fn is_mail_merge_next(&self) -> bool {
        field_instruction_remainder(&self.instruction, "NEXT").is_some()
    }

    /// Parse this field as inert typed `NEXT` mail-merge control metadata.
    ///
    /// Returns `Ok(None)` for fields other than `NEXT`. The result exposes only
    /// stored content and dirty/lock state; it never advances a record, opens a
    /// data source, performs a merge, or refreshes a field.
    pub fn mail_merge_next(&self) -> Result<Option<MergeNext>> {
        MergeNext::from_field(self)
    }

    /// Check whether this is a `NEXTIF` mail-merge conditional control field.
    ///
    /// Recognition is limited to the stored instruction. It never evaluates a
    /// comparison, selects a record, opens a data source, performs a merge, or
    /// refreshes the result.
    pub fn is_mail_merge_next_if(&self) -> bool {
        field_instruction_remainder(&self.instruction, "NEXTIF").is_some()
    }

    /// Check whether this is a `SKIPIF` mail-merge conditional control field.
    ///
    /// Recognition is limited to the stored instruction. It never evaluates a
    /// comparison, skips a record, opens a data source, performs a merge, or
    /// refreshes the result.
    pub fn is_mail_merge_skip_if(&self) -> bool {
        field_instruction_remainder(&self.instruction, "SKIPIF").is_some()
    }

    /// Check whether this is a `NEXTIF` or `SKIPIF` mail-merge control field.
    pub fn is_mail_merge_conditional_control(&self) -> bool {
        self.is_mail_merge_next_if() || self.is_mail_merge_skip_if()
    }

    /// Parse this field as inert typed conditional mail-merge control metadata.
    ///
    /// Returns `Ok(None)` for fields other than `NEXTIF` and `SKIPIF`. The unparsed
    /// comparison, cached content, and dirty/lock state are stored metadata
    /// only; this method never evaluates a comparison, changes record
    /// selection, opens a data source, performs a merge, or refreshes a field.
    pub fn mail_merge_conditional_control(&self) -> Result<Option<MergeControl>> {
        MergeControl::from_field(self)
    }

    /// Check whether this is an `IF` conditional field.
    ///
    /// Recognition is limited to the stored instruction. It never parses or
    /// evaluates an expression, resolves field values, or refreshes the result.
    pub fn is_if_field(&self) -> bool {
        field_instruction_remainder(&self.instruction, "IF").is_some()
    }

    /// Parse this field as inert typed `IF` metadata.
    ///
    /// Returns `Ok(None)` for fields other than `IF`. The stored expression,
    /// cached content, and dirty/lock state are metadata only; this method never
    /// parses or evaluates an expression, resolves field values, or refreshes a
    /// field.
    pub fn if_field(&self) -> Result<Option<If>> {
        If::from_field(self)
    }

    /// Check whether this is a `COMPARE` field.
    ///
    /// Recognition is limited to the stored instruction. It never parses or
    /// evaluates a comparison, resolves nested field values, or refreshes the
    /// cached result.
    pub fn is_compare_field(&self) -> bool {
        field_instruction_remainder(&self.instruction, "COMPARE").is_some()
    }

    /// Parse this field as inert typed `COMPARE` metadata.
    ///
    /// Returns `Ok(None)` for fields other than `COMPARE`. The stored
    /// comparison, cached content, and dirty/lock state are metadata only; this
    /// method never parses or evaluates a comparison, resolves nested field
    /// values, or refreshes a field.
    pub fn compare_field(&self) -> Result<Option<Compare>> {
        Compare::from_field(self)
    }

    /// Check whether this is a bookmark-reference field.
    ///
    /// This recognizes `REF`, `PAGEREF`, `FTNREF`, and `NOTEREF` stored
    /// instructions. It never looks up a bookmark, reads a referenced range or
    /// note, resolves a page number, creates a link, calculates a relative
    /// position, or refreshes the result.
    pub fn is_reference_field(&self) -> bool {
        ReferenceKind::from_instruction(&self.instruction).is_some()
    }

    /// Parse this field as inert bookmark-reference metadata.
    ///
    /// Returns `Ok(None)` for fields other than `REF`, `PAGEREF`, `FTNREF`, and
    /// `NOTEREF`. The stored kind, target, options, unknown switches, cached
    /// content, and dirty/lock state are metadata only; this method never looks
    /// up a bookmark, reads a referenced range or note, resolves a page number,
    /// creates a link, calculates a relative position, or refreshes a field.
    pub fn reference_field(&self) -> Result<Option<Reference>> {
        Reference::from_field(self)
    }

    /// Check whether this is a `SET` field.
    ///
    /// Recognition is limited to the stored instruction. It never evaluates an
    /// expression, looks up or changes a bookmark, changes document state, or
    /// refreshes the result.
    pub fn is_set_field(&self) -> bool {
        field_instruction_remainder(&self.instruction, "SET").is_some()
    }

    /// Parse this field as inert typed `SET` metadata.
    ///
    /// Returns `Ok(None)` for fields other than `SET`. The stored target,
    /// opaque expression, cached content, and dirty/lock state are metadata
    /// only; this method never evaluates an expression, looks up or changes a
    /// bookmark, changes document state, or refreshes a field.
    pub fn set_field(&self) -> Result<Option<Set>> {
        Set::from_field(self)
    }

    /// Check whether this is a `SEQ` field.
    ///
    /// Recognition is limited to the stored instruction. It never looks up a
    /// bookmark, increments or resets a sequence, calculates a number, or
    /// refreshes the result.
    pub fn is_sequence_field(&self) -> bool {
        field_instruction_remainder(&self.instruction, "SEQ").is_some()
    }

    /// Parse this field as inert typed `SEQ` metadata.
    ///
    /// Returns `Ok(None)` for fields other than `SEQ`. The stored identifier,
    /// optional bookmark, opaque tail, cached content, and dirty/lock state are
    /// metadata only; this method never looks up a bookmark, increments or
    /// resets a sequence, calculates a number, or refreshes a field.
    pub fn sequence_field(&self) -> Result<Option<Sequence>> {
        Sequence::from_field(self)
    }

    /// Check whether this is a Word `=` formula field.
    ///
    /// Recognition is limited to the stored instruction. It never parses or
    /// evaluates a formula, reads table cells or bookmarks, resolves field
    /// values, or refreshes the result.
    pub fn is_formula_field(&self) -> bool {
        self.instruction.trim_start().starts_with('=')
    }

    /// Parse this field as inert typed formula metadata.
    ///
    /// Returns `Ok(None)` for fields that do not begin with `=`. The stored
    /// formula, cached content, and dirty/lock state are metadata only; this
    /// method never parses or evaluates a formula, reads table cells or
    /// bookmarks, resolves field values, or refreshes a field.
    pub fn formula_field(&self) -> Result<Option<Formula>> {
        Formula::from_field(self)
    }

    /// Check whether this is an `EQ` equation field.
    ///
    /// Recognition is limited to stored field metadata. It never parses,
    /// calculates, formats, renders, or refreshes an equation.
    pub fn is_equation(&self) -> bool {
        field_instruction_remainder(&self.instruction, "EQ").is_some()
    }

    /// Parse this field as inert `EQ` equation metadata.
    ///
    /// Returns `Ok(None)` for fields other than `EQ`. The stored expression,
    /// cached content, and dirty/lock state are metadata only; this method never
    /// parses, calculates, formats, renders, or refreshes an equation.
    pub fn equation(&self) -> Result<Option<Equation>> {
        Equation::from_field(self)
    }

    /// Check whether this is a `HYPERLINK` field.
    ///
    /// Recognition is limited to stored field metadata. It never opens,
    /// resolves, follows, or refreshes a hyperlink target.
    pub fn is_hyperlink_field(&self) -> bool {
        field_instruction_remainder(&self.instruction, "HYPERLINK").is_some()
    }

    /// Parse this field as inert `HYPERLINK` metadata.
    ///
    /// Returns `Ok(None)` for fields other than `HYPERLINK`. The stored target,
    /// bookmark, tooltip, frame, image-map-coordinate request, switches, cached
    /// content, and dirty/lock state are metadata only; this method never opens, resolves,
    /// follows, activates, or refreshes a link.
    pub fn hyperlink_field(&self) -> Result<Option<Hyperlink>> {
        Hyperlink::from_field(self)
    }

    /// Check whether this is a `QUOTE` text-insertion field.
    ///
    /// Recognition is limited to the stored instruction. It never interprets
    /// character codes, expands nested fields, inserts text, or refreshes the
    /// cached result.
    pub fn is_quote_field(&self) -> bool {
        field_instruction_remainder(&self.instruction, "QUOTE").is_some()
    }

    /// Parse this field as inert typed `QUOTE` metadata.
    ///
    /// Returns `Ok(None)` for fields other than `QUOTE`. The stored text argument,
    /// switches, cached content, and dirty/lock state are metadata only; this
    /// method never interprets character codes, expands nested fields, inserts
    /// text, or refreshes a field.
    pub fn quote_field(&self) -> Result<Option<Quote>> {
        Quote::from_field(self)
    }

    /// Check whether this is a `SYMBOL` field.
    ///
    /// Recognition is limited to the stored instruction. It never maps a
    /// character code, looks up a font, inserts a glyph, changes formatting or
    /// layout, or refreshes the cached result.
    pub fn is_symbol_field(&self) -> bool {
        field_instruction_remainder(&self.instruction, "SYMBOL").is_some()
    }

    /// Parse this field as inert typed `SYMBOL` metadata.
    ///
    /// Returns `Ok(None)` for fields other than `SYMBOL`. The stored character
    /// argument, switches, cached content, and dirty/lock state are metadata
    /// only; this method never maps a character code, looks up a font, inserts
    /// a glyph, changes formatting or layout, or refreshes a field.
    pub fn symbol_field(&self) -> Result<Option<Symbol>> {
        Symbol::from_field(self)
    }

    /// Check whether this is a legacy automatic-numbering field.
    ///
    /// Recognition is limited to the stored instruction. It never calculates
    /// paragraph numbers, reads heading or style state, changes layout, or
    /// refreshes the cached result.
    pub fn is_auto_number_field(&self) -> bool {
        AutoNumberKind::from_instruction(&self.instruction).is_some()
    }

    /// Parse this field as inert legacy automatic-numbering metadata.
    ///
    /// Returns `Ok(None)` for fields other than `AUTONUM`, `AUTONUMLGL`, and
    /// `AUTONUMOUT`. The stored kind, switches, cached content, and dirty/lock
    /// state are metadata only; this method never calculates a number, reads
    /// paragraph, heading, or style state, changes layout, or refreshes a
    /// field.
    pub fn auto_number_field(&self) -> Result<Option<AutoNumber>> {
        AutoNumber::from_field(self)
    }

    /// Check whether this is a `LISTNUM` field.
    ///
    /// Recognition is limited to the stored instruction. It never looks up a
    /// list, determines a level or start value, calculates a number, changes
    /// layout, or refreshes the cached result.
    pub fn is_list_number_field(&self) -> bool {
        field_instruction_remainder(&self.instruction, "LISTNUM").is_some()
    }

    /// Parse this field as inert typed `LISTNUM` metadata.
    ///
    /// Returns `Ok(None)` for fields other than `LISTNUM`. The stored optional
    /// list name, switches, cached content, and dirty/lock state are metadata
    /// only; this method never looks up a list, determines a level or start
    /// value, calculates a number, changes layout, or refreshes a field.
    pub fn list_number_field(&self) -> Result<Option<ListNumber>> {
        ListNumber::from_field(self)
    }

    /// Check whether this is a `STYLEREF` field.
    ///
    /// Recognition is limited to stored field metadata. It never looks up
    /// styled text, searches document stories, calculates paragraph numbers or
    /// relative positions, resolves page layout, or refreshes the result.
    pub fn is_style_reference_field(&self) -> bool {
        field_instruction_remainder(&self.instruction, "STYLEREF").is_some()
    }

    /// Parse this field as inert typed `STYLEREF` metadata.
    ///
    /// Returns `Ok(None)` for fields other than `STYLEREF`. The stored style name,
    /// options, switches, cached content, and dirty/lock state are metadata
    /// only; this method never looks up styled text, searches document stories,
    /// calculates paragraph numbers or relative positions, resolves page
    /// layout, or refreshes a field.
    pub fn style_reference_field(&self) -> Result<Option<StyleReference>> {
        StyleReference::from_field(self)
    }

    /// Check whether this is an `ASK` prompt field.
    ///
    /// Recognition is limited to the stored instruction. It never displays a
    /// prompt, captures a response, writes a bookmark, performs a merge, or
    /// refreshes the result.
    pub fn is_ask_field(&self) -> bool {
        field_instruction_remainder(&self.instruction, "ASK").is_some()
    }

    /// Check whether this is a `FILLIN` prompt field.
    ///
    /// Recognition is limited to the stored instruction. It never displays a
    /// prompt, captures a response, performs a merge, or refreshes the result.
    pub fn is_fill_in_field(&self) -> bool {
        field_instruction_remainder(&self.instruction, "FILLIN").is_some()
    }

    /// Check whether this is an `ASK` or `FILLIN` prompt field.
    pub fn is_prompt_field(&self) -> bool {
        self.is_ask_field() || self.is_fill_in_field()
    }

    /// Parse this field as inert typed `ASK` or `FILLIN` metadata.
    ///
    /// Returns `Ok(None)` for fields other than `ASK` and `FILLIN`. The returned
    /// values expose stored prompt, bookmark, default-response, cached-content,
    /// and dirty/lock metadata only. This method never displays a prompt,
    /// captures a response, creates or updates a bookmark, performs a merge,
    /// or refreshes a field.
    pub fn prompt_field(&self) -> Result<Option<Prompt>> {
        Prompt::from_field(self)
    }

    /// Check whether this is an `ADDRESSBLOCK` mail-merge recipient field.
    ///
    /// Recognition is limited to stored field metadata. It never opens a data
    /// source, selects a record, performs a merge, or refreshes the result.
    pub fn is_address_block(&self) -> bool {
        field_instruction_remainder(&self.instruction, "ADDRESSBLOCK").is_some()
    }

    /// Check whether this is a `GREETINGLINE` mail-merge recipient field.
    ///
    /// Recognition is limited to stored field metadata. It never opens a data
    /// source, selects a record, performs a merge, or refreshes the result.
    pub fn is_greeting_line(&self) -> bool {
        field_instruction_remainder(&self.instruction, "GREETINGLINE").is_some()
    }

    /// Check whether this is an `ADDRESSBLOCK` or `GREETINGLINE` field.
    pub fn is_mail_merge_recipient_field(&self) -> bool {
        self.is_address_block() || self.is_greeting_line()
    }

    /// Parse this field as inert typed mail-merge recipient metadata.
    ///
    /// Returns `Ok(None)` for fields other than `ADDRESSBLOCK` and
    /// `GREETINGLINE`. The returned values expose only stored layout,
    /// locale, country, fallback, cached-content, and dirty/lock metadata. This
    /// method never opens a data source, selects a record, performs a merge,
    /// generates text, or refreshes a field.
    pub fn mail_merge_recipient_field(&self) -> Result<Option<Recipient>> {
        Recipient::from_field(self)
    }

    /// Check whether this is a `CITATION` bibliography field.
    ///
    /// Recognition is limited to the stored field instruction. It never looks
    /// up bibliography sources, formats a citation, follows a data-store
    /// reference, or refreshes the cached result.
    pub fn is_citation(&self) -> bool {
        field_instruction_remainder(&self.instruction, "CITATION").is_some()
    }

    /// Parse this field as an inert typed bibliography citation.
    ///
    /// Returns `Ok(None)` for non-`CITATION` fields. The result exposes only
    /// stored source tags, switches, cached content, and dirty/lock state; it
    /// never resolves sources or formats a citation.
    pub fn citation(&self) -> Result<Option<Citation>> {
        Citation::from_field(self)
    }

    /// Check whether this is a `BIBLIOGRAPHY` field.
    ///
    /// This recognizes persisted configuration only. It does not enumerate
    /// sources, sort them, or generate bibliography text.
    pub fn is_bibliography(&self) -> bool {
        field_instruction_remainder(&self.instruction, "BIBLIOGRAPHY").is_some()
    }

    /// Parse this field as an inert typed bibliography field.
    ///
    /// Returns `Ok(None)` for non-`BIBLIOGRAPHY` fields. Stored switches and
    /// cached visible content remain data only; no bibliography is generated.
    pub fn bibliography(&self) -> Result<Option<Bibliography>> {
        Bibliography::from_field(self)
    }

    /// Check whether this is a `DOCVARIABLE` field.
    ///
    /// Recognition is limited to the stored field instruction. It never reads
    /// settings XML, resolves a variable value, or refreshes the field.
    pub fn is_document_variable(&self) -> bool {
        field_instruction_remainder(&self.instruction, "DOCVARIABLE").is_some()
    }

    /// Parse this field as inert typed document-variable metadata.
    ///
    /// Returns `Ok(None)` for non-`DOCVARIABLE` fields. The result exposes the
    /// stored variable name, switches, cached content, and dirty/lock state
    /// only; it never reads settings XML, resolves a value, or refreshes a
    /// field.
    pub fn document_variable(&self) -> Result<Option<Variable>> {
        Variable::from_field(self)
    }

    /// Check whether this is a `DOCPROPERTY` field.
    ///
    /// Recognition is limited to the stored field instruction. It never reads
    /// a package property, resolves a value, or refreshes the field.
    pub fn is_document_property(&self) -> bool {
        field_instruction_remainder(&self.instruction, "DOCPROPERTY").is_some()
    }

    /// Parse this field as inert typed document-property metadata.
    ///
    /// Returns `Ok(None)` for non-`DOCPROPERTY` fields. The result exposes the
    /// stored property name, switches, cached content, and dirty/lock state
    /// only; it never reads core, extended, or custom package properties,
    /// resolves a value, or refreshes a field.
    pub fn document_property(&self) -> Result<Option<Property>> {
        Property::from_field(self)
    }

    /// Check whether this is an explicit legacy `INFO` field.
    ///
    /// Recognition is limited to the stored field instruction. It never reads,
    /// resolves, modifies, or writes document or template properties, or
    /// refreshes the field.
    pub fn is_info_field(&self) -> bool {
        field_instruction_remainder(&self.instruction, "INFO").is_some()
    }

    /// Parse this field as inert typed legacy `INFO` metadata.
    ///
    /// Returns `Ok(None)` for fields other than an explicit `INFO` field. The
    /// result exposes the stored property selector, optional replacement value,
    /// switches, cached content, and dirty/lock state only; it never reads,
    /// resolves, modifies, or writes document or template properties, or
    /// refreshes a field.
    pub fn info_field(&self) -> Result<Option<Info>> {
        Info::from_field(self)
    }

    /// Check whether this is a built-in document-information field.
    ///
    /// Recognition is limited to the stored field instruction. It never reads
    /// package metadata or host identity data, resolves a value, or refreshes
    /// the field.
    pub fn is_document_information(&self) -> bool {
        InformationKind::from_instruction(&self.instruction).is_some()
    }

    /// Parse this field as inert typed document-information metadata.
    ///
    /// Returns `Ok(None)` for fields outside the built-in document-information
    /// family. The result exposes only the stored kind, switches, cached
    /// content, and dirty/lock state; it never reads core or extended package
    /// properties, reads or modifies host identity data, calculates dates,
    /// revisions, or statistics, resolves a value, or refreshes a field.
    pub fn document_information(&self) -> Result<Option<Information>> {
        Information::from_field(self)
    }

    /// Check whether this is a built-in document-context or runtime field.
    ///
    /// Recognition is limited to the stored field instruction. It never reads
    /// a document path, attached template, host filesystem state or file size,
    /// current clock, or page and section layout, resolves a value, or refreshes
    /// the field.
    pub fn is_document_context(&self) -> bool {
        ContextKind::from_instruction(&self.instruction).is_some()
    }

    /// Parse this field as inert typed document-context or runtime metadata.
    ///
    /// Returns `Ok(None)` for fields outside the `FILENAME`, `TEMPLATE`, `DATE`,
    /// `TIME`, `PAGE`, `FILESIZE`, `SECTION`, and `SECTIONPAGES` family. The
    /// result exposes only the stored kind, switches, cached content, and
    /// dirty/lock state; it never reads a document path, attached template,
    /// host filesystem state or file size, current clock, or page and section
    /// layout, resolves a value, or refreshes a field.
    pub fn document_context(&self) -> Result<Option<Context>> {
        Context::from_field(self)
    }

    /// Check whether this is a `MACROBUTTON` field.
    ///
    /// Recognition is limited to the stored field instruction. It never
    /// resolves, loads, invokes, or otherwise executes a macro or command.
    pub fn is_macro_button(&self) -> bool {
        field_instruction_remainder(&self.instruction, "MACROBUTTON").is_some()
    }

    /// Parse this field as inert typed macro-button metadata.
    ///
    /// Returns `Ok(None)` for non-`MACROBUTTON` fields. The result exposes
    /// only the stored macro or command name, button text, cached content, and
    /// dirty/lock state; it never resolves, loads, invokes, or executes the
    /// named target.
    pub fn macro_button(&self) -> Result<Option<MacroButton>> {
        MacroButton::from_field(self)
    }

    /// Check whether this is a `GOTOBUTTON` field.
    ///
    /// Recognition is limited to the stored field instruction. It never
    /// resolves a destination, changes the insertion point, or refreshes the
    /// cached result.
    pub fn is_go_to_button(&self) -> bool {
        field_instruction_remainder(&self.instruction, "GOTOBUTTON").is_some()
    }

    /// Parse this field as inert typed `GOTOBUTTON` metadata.
    ///
    /// Returns `Ok(None)` for non-`GOTOBUTTON` fields. The result exposes
    /// only the stored target, button text, cached content, and dirty/lock
    /// state; it never resolves a bookmark, page, annotation, footnote, or
    /// other target, changes the insertion point, or refreshes a field.
    pub fn go_to_button(&self) -> Result<Option<GoToButton>> {
        GoToButton::from_field(self)
    }

    /// Check whether this is a `PRINT` field.
    ///
    /// Recognition is limited to the stored instruction. It never interprets
    /// printer-control codes, sends data to a printer, or refreshes the cached
    /// result.
    pub fn is_print_field(&self) -> bool {
        field_instruction_remainder(&self.instruction, "PRINT").is_some()
    }

    /// Parse this field as inert typed `PRINT` metadata.
    ///
    /// Returns `Ok(None)` for fields other than `PRINT`. The stored
    /// printer-instruction text, cached content, and dirty/lock state are
    /// metadata only; this method never interprets the instruction, sends data
    /// to a printer, or refreshes a field.
    pub fn print_field(&self) -> Result<Option<Print>> {
        Print::from_field(self)
    }

    /// Check whether this is an `EMBED` field.
    ///
    /// Recognition is limited to stored field metadata. It never loads,
    /// inspects, activates, renders, or executes an embedded object, or
    /// refreshes the field.
    pub fn is_embed_field(&self) -> bool {
        field_instruction_remainder(&self.instruction, "EMBED").is_some()
    }

    /// Parse this field as inert typed `EMBED` metadata.
    ///
    /// Returns `Ok(None)` for fields other than `EMBED`. The stored opaque object
    /// instruction, cached content, and dirty/lock state are metadata only;
    /// this method never loads, inspects, activates, renders, or executes an
    /// embedded object, or refreshes a field.
    pub fn embed_field(&self) -> Result<Option<Embed>> {
        Embed::from_field(self)
    }

    /// Check whether this is a `BARCODE` field.
    ///
    /// Recognition is limited to stored field metadata. It never parses or
    /// validates barcode data or symbology, generates or renders a barcode, or
    /// refreshes the field.
    pub fn is_barcode_field(&self) -> bool {
        field_instruction_remainder(&self.instruction, "BARCODE").is_some()
    }

    /// Parse this field as inert typed `BARCODE` metadata.
    ///
    /// Returns `Ok(None)` for fields other than `BARCODE`. The stored opaque barcode
    /// instruction, cached content, and dirty/lock state are metadata only;
    /// this method never parses or validates barcode data or symbology,
    /// generates or renders a barcode, or refreshes a field.
    pub fn barcode_field(&self) -> Result<Option<Barcode>> {
        Barcode::from_field(self)
    }

    /// Check whether this is a `BIDIOUTLINE` field.
    ///
    /// Recognition is limited to stored field metadata. It never reads
    /// right-to-left language, paragraph outline, or layout state; chooses a
    /// numbering system; calculates a result; or refreshes the field.
    pub fn is_bidi_outline_field(&self) -> bool {
        field_instruction_remainder(&self.instruction, "BIDIOUTLINE").is_some()
    }

    /// Parse this field as inert typed `BIDIOUTLINE` metadata.
    ///
    /// Returns `Ok(None)` for fields other than `BIDIOUTLINE`. The stored opaque
    /// instruction, cached content, and dirty/lock state are metadata only;
    /// this method never reads right-to-left language, paragraph outline, or
    /// layout state; chooses a numbering system; calculates a result; or
    /// refreshes a field.
    pub fn bidi_outline_field(&self) -> Result<Option<BidiOutline>> {
        BidiOutline::from_field(self)
    }

    /// Check whether this is a `SHAPE` drawing-canvas anchor field.
    ///
    /// Recognition is limited to stored field metadata. It never locates,
    /// links, loads, positions, lays out, or renders a drawing or canvas, or
    /// refreshes the field.
    pub fn is_shape_field(&self) -> bool {
        field_instruction_remainder(&self.instruction, "SHAPE").is_some()
    }

    /// Parse this field as inert typed `SHAPE` drawing-canvas anchor metadata.
    ///
    /// Returns `Ok(None)` for fields other than `SHAPE`. The stored opaque
    /// instruction, cached content, and dirty/lock state are metadata only;
    /// this method never locates, links, loads, positions, lays out, or renders
    /// a drawing or canvas, or refreshes a field.
    pub fn shape_field(&self) -> Result<Option<Shape>> {
        Shape::from_field(self)
    }

    /// Check whether this is a legacy text, checkbox, or drop-down form-code field.
    ///
    /// Recognition is limited to the stored field instruction. It never reads
    /// associated form-property XML, fills a form, changes a selection or
    /// checkbox state, invokes entry or exit macros, or refreshes a field.
    pub fn is_legacy_form_field(&self) -> bool {
        LegacyFormKind::from_instruction(&self.instruction).is_some()
    }

    /// Parse this field as inert typed legacy form-code metadata.
    ///
    /// Returns `Ok(None)` for fields outside the `FORMTEXT`, `FORMCHECKBOX`, and
    /// `FORMDROPDOWN` family. The result exposes only stored kind, opaque
    /// instruction text, cached content, and dirty/lock state; it never reads
    /// associated form-property XML, fills a form, changes a selection or
    /// checkbox state, invokes entry or exit macros, or refreshes a field.
    pub fn legacy_form_field(&self) -> Result<Option<LegacyForm>> {
        LegacyForm::from_field(self)
    }

    /// Check whether this is a `PRIVATE` conversion-data field.
    ///
    /// Recognition is limited to stored field metadata. It never converts a
    /// document, interprets field data, changes hidden-text visibility or
    /// layout, or refreshes a field. `PRIVATE` does not provide confidentiality
    /// semantics.
    pub fn is_private_field(&self) -> bool {
        field_instruction_remainder(&self.instruction, "PRIVATE").is_some()
    }

    /// Parse this field as inert typed `PRIVATE` conversion-data metadata.
    ///
    /// Returns `Ok(None)` for fields other than `PRIVATE`. The stored opaque
    /// instruction, cached content, and dirty/lock state are metadata only;
    /// this method never converts a document, interprets field data, changes
    /// hidden-text visibility or layout, or refreshes a field. `PRIVATE` does
    /// not provide confidentiality semantics.
    pub fn private_field(&self) -> Result<Option<Private>> {
        Private::from_field(self)
    }

    /// Check whether this is a `DATABASE` query field.
    ///
    /// Recognition is limited to stored field metadata. It never opens a data
    /// source or database, uses connection information, executes SQL, generates
    /// or inserts a table, changes layout, or refreshes a field.
    pub fn is_database_field(&self) -> bool {
        field_instruction_remainder(&self.instruction, "DATABASE").is_some()
    }

    /// Parse this field as inert typed `DATABASE` query metadata.
    ///
    /// Returns `Ok(None)` for fields other than `DATABASE`. The stored opaque
    /// instruction, cached content, and dirty/lock state are metadata only;
    /// this method never opens a data source or database, uses connection
    /// information, executes SQL, generates or inserts a table, changes layout,
    /// or refreshes a field.
    pub fn database_field(&self) -> Result<Option<Database>> {
        Database::from_field(self)
    }

    /// Check whether this is an `ADDIN` field.
    ///
    /// Recognition is limited to stored metadata. It never loads an add-in,
    /// invokes code, or refreshes a field result.
    pub fn is_add_in_field(&self) -> bool {
        field_instruction_remainder(&self.instruction, "ADDIN").is_some()
    }

    /// Check whether this is a `CONTROL` field.
    ///
    /// Recognition is limited to stored metadata. It never instantiates an OCX
    /// control, invokes code, renders content, or refreshes a field result.
    pub fn is_control_field(&self) -> bool {
        field_instruction_remainder(&self.instruction, "CONTROL").is_some()
    }

    /// Check whether this is an `HTMLCONTROL` field.
    ///
    /// Recognition is limited to stored metadata. It never instantiates an HTML
    /// control, executes script, renders content, or refreshes a field result.
    pub fn is_html_control_field(&self) -> bool {
        field_instruction_remainder(&self.instruction, "HTMLCONTROL").is_some()
    }

    /// Check whether this is an `ADDIN`, `CONTROL`, or `HTMLCONTROL` field.
    pub fn is_active_content_field(&self) -> bool {
        self.is_add_in_field() || self.is_control_field() || self.is_html_control_field()
    }

    /// Parse this field as inert typed add-in or control metadata.
    ///
    /// Returns `Ok(None)` for other fields. The stored kind, instruction,
    /// cached content, and dirty/lock state are opaque metadata only; this
    /// method never loads an add-in, instantiates a control, invokes code,
    /// executes script, renders content, accesses an external resource, or
    /// refreshes a field.
    pub fn active_content_field(&self) -> Result<Option<ActiveContent>> {
        ActiveContent::from_field(self)
    }

    /// Check whether this is a `GLOSSARY` or `AUTOTEXT` building-block field.
    ///
    /// Recognition is limited to stored metadata. It never looks up a building
    /// block, reads a template, inserts content, changes bookmarks, or refreshes
    /// a field result.
    pub fn is_auto_text_field(&self) -> bool {
        field_instruction_remainder(&self.instruction, "GLOSSARY").is_some()
            || field_instruction_remainder(&self.instruction, "AUTOTEXT").is_some()
    }

    /// Parse this field as inert typed building-block metadata.
    ///
    /// Returns `Ok(None)` for fields other than `GLOSSARY` and `AUTOTEXT`.
    /// The stored kind, entry name, switches, cached content, and dirty/lock
    /// state are metadata only; this method never looks up a building block,
    /// reads a template, inserts content, changes bookmarks, accesses an
    /// external resource, or refreshes a field.
    pub fn auto_text_field(&self) -> Result<Option<AutoText>> {
        AutoText::from_field(self)
    }

    /// Check whether this is an `AUTOTEXTLIST` building-block selection field.
    ///
    /// Recognition is limited to stored metadata. It never shows a selection
    /// UI, looks up eligible building blocks, reads a template, inserts
    /// content, or refreshes a field result.
    pub fn is_auto_text_list_field(&self) -> bool {
        field_instruction_remainder(&self.instruction, "AUTOTEXTLIST").is_some()
    }

    /// Parse this field as inert typed building-block selection metadata.
    ///
    /// Returns `Ok(None)` for fields other than `AUTOTEXTLIST`. The
    /// stored display text, style/tip options, switches, cached content, and
    /// dirty/lock state are metadata only; this method never shows a selection
    /// UI, looks up eligible building blocks, reads a template, inserts
    /// content, accesses an external resource, or refreshes a field.
    pub fn auto_text_list_field(&self) -> Result<Option<AutoTextList>> {
        AutoTextList::from_field(self)
    }

    /// Check whether this is a `USERADDRESS` field.
    ///
    /// Recognition is limited to the stored field instruction. It never reads
    /// the current user's address or refreshes the cached result.
    pub fn is_user_address(&self) -> bool {
        field_instruction_remainder(&self.instruction, "USERADDRESS").is_some()
    }

    /// Check whether this is a `USERINITIALS` field.
    ///
    /// Recognition is limited to the stored field instruction. It never reads
    /// the current user's initials or refreshes the cached result.
    pub fn is_user_initials(&self) -> bool {
        field_instruction_remainder(&self.instruction, "USERINITIALS").is_some()
    }

    /// Check whether this is a `USERNAME` field.
    ///
    /// Recognition is limited to the stored field instruction. It never reads
    /// the current user's name or refreshes the cached result.
    pub fn is_user_name(&self) -> bool {
        field_instruction_remainder(&self.instruction, "USERNAME").is_some()
    }

    /// Check whether this is a `USERADDRESS`, `USERINITIALS`, or `USERNAME` field.
    pub fn is_user_identity_field(&self) -> bool {
        self.is_user_address() || self.is_user_initials() || self.is_user_name()
    }

    /// Parse this field as inert typed user-identity metadata.
    ///
    /// Returns `Ok(None)` for fields other than `USERADDRESS`, `USERINITIALS`, and
    /// `USERNAME`. The result exposes only stored override, formatting,
    /// cached-content, and dirty/lock metadata; it never reads or modifies a
    /// host user's identity or refreshes a field.
    pub fn user_identity_field(&self) -> Result<Option<UserIdentity>> {
        UserIdentity::from_field(self)
    }

    /// Check whether this is an `ADVANCE` placement field.
    ///
    /// Recognition is limited to the stored field instruction. It never moves
    /// text, changes layout, reflows content, or refreshes a cached result.
    pub fn is_advance_field(&self) -> bool {
        field_instruction_remainder(&self.instruction, "ADVANCE").is_some()
    }

    /// Parse this field as inert typed `ADVANCE` placement metadata.
    ///
    /// Returns `Ok(None)` for fields other than `ADVANCE`. The returned
    /// values expose only stored point adjustments, cached content, and
    /// dirty/lock state. This method never moves text, changes layout, reflows
    /// content, or refreshes a field.
    pub fn advance_field(&self) -> Result<Option<Advance>> {
        Advance::from_field(self)
    }

    /// Check whether this is a legacy `LINK` field.
    ///
    /// Recognition is limited to the stored field instruction. It never
    /// activates an OLE server, opens a source, or refreshes the field.
    pub fn is_link(&self) -> bool {
        field_instruction_remainder(&self.instruction, "LINK").is_some()
    }

    /// Parse this field as inert typed `LINK` metadata.
    ///
    /// Returns `Ok(None)` for non-`LINK` fields. The result exposes stored
    /// application, source, item, result, formatting, and cached metadata only;
    /// it never activates, opens, contacts, converts, evaluates, or executes
    /// anything.
    pub fn link(&self) -> Result<Option<Link>> {
        Link::from_field(self)
    }

    /// Check whether this is a legacy DDE field.
    ///
    /// Recognition is limited to the stored field instruction. It never
    /// launches an application, initiates a DDE conversation, opens a source,
    /// or refreshes the field.
    pub fn is_dde(&self) -> bool {
        field_instruction_remainder(&self.instruction, "DDE").is_some()
    }

    /// Check whether this is a legacy automatically updating DDEAUTO field.
    ///
    /// Recognition is limited to the stored field instruction. It never
    /// launches an application, initiates a DDE conversation, opens a source,
    /// or refreshes the field.
    pub fn is_dde_auto(&self) -> bool {
        field_instruction_remainder(&self.instruction, "DDEAUTO").is_some()
    }

    /// Parse this field as inert typed DDE or DDEAUTO metadata.
    ///
    /// Returns Ok(None) for other fields. The result exposes stored
    /// application, source, item, representation, and cached metadata only; it
    /// never launches an application, initiates a DDE conversation, opens,
    /// contacts, refreshes, converts, evaluates, or executes anything.
    pub fn dde_link(&self) -> Result<Option<Dde>> {
        Dde::from_field(self)
    }

    /// Check whether this is an `INCLUDETEXT` or historical `INCLUDE` field.
    ///
    /// Recognition is limited to the stored field instruction. It never opens,
    /// resolves, imports, fetches, or refreshes the referenced source.
    pub fn is_include_text(&self) -> bool {
        field_instruction_remainder(&self.instruction, "INCLUDETEXT").is_some()
            || field_instruction_remainder(&self.instruction, "INCLUDE").is_some()
    }

    /// Check whether this is an `INCLUDEPICTURE` or historical `IMPORT` field.
    ///
    /// Recognition is limited to the stored field instruction. It never opens,
    /// resolves, imports, fetches, or refreshes the referenced source.
    pub fn is_include_picture(&self) -> bool {
        field_instruction_remainder(&self.instruction, "INCLUDEPICTURE").is_some()
            || field_instruction_remainder(&self.instruction, "IMPORT").is_some()
    }

    /// Parse this field as inert external-include metadata.
    ///
    /// Returns Ok(None) for fields other than `INCLUDETEXT`/`INCLUDEPICTURE` or
    /// their historical `INCLUDE`/`IMPORT` aliases. The result exposes stored
    /// source, bookmark, converter, XML, and cached metadata only; it never
    /// opens, resolves, imports, fetches, refreshes, converts, evaluates, or
    /// executes anything.
    pub fn external_include(&self) -> Result<Option<Include>> {
        Include::from_field(self)
    }

    /// Check whether this is an RD referenced-document field.
    ///
    /// Recognition is limited to the stored field instruction. It never opens,
    /// resolves, reads, imports, or refreshes the referenced document.
    pub fn is_referenced_document(&self) -> bool {
        field_instruction_remainder(&self.instruction, "RD").is_some()
    }

    /// Parse this field as inert referenced-document metadata.
    ///
    /// Returns Ok(None) for non-RD fields. The result exposes only the stored
    /// path, relative-path request, switches, cached content, and dirty/lock
    /// state; it never opens, resolves, imports, evaluates, or executes
    /// anything.
    pub fn referenced_document(&self) -> Result<Option<SubDocument>> {
        SubDocument::from_field(self)
    }

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

/// One lexical switch in a Word field instruction.
///
/// Switch names are normalized to ASCII lowercase. Quoted and unquoted
/// arguments are decoded into their logical text. Typed field models retain the
/// complete original instruction alongside these values.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Switch {
    pub(super) name: char,
    pub(super) argument: Option<String>,
}

impl Switch {
    /// Return the switch character, without its leading backslash.
    pub fn name(&self) -> char {
        self.name
    }

    /// Return the optional argument supplied to this switch.
    pub fn argument(&self) -> Option<&str> {
        self.argument.as_deref()
    }
}

/// The stored kind of a legacy DDE field.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DdeKind {
    /// A DDE field, which can request automatic updates with its a switch.
    Dde,
    /// A DDEAUTO field, which declares automatic updates.
    DdeAuto,
}

/// One stored DDE result representation switch.
///
/// This value describes a requested representation only. It never causes a
/// source to be contacted, converted, embedded, or displayed.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DdeFormat {
    /// The b switch requests a bitmap representation.
    Bitmap,
    /// The h switch requests HTML-formatted text.
    Html,
    /// The p switch requests a picture representation.
    Picture,
    /// The r switch requests rich-text format.
    RichText,
    /// The t switch requests text-only format.
    Text,
    /// The u switch requests Unicode text.
    UnicodeText,
}

/// Typed, inert metadata for a legacy DDE or DDEAUTO field.
///
/// Application, source, item, representation, and storage switches are
/// retained as stored field data. This type never launches an application,
/// initiates a DDE conversation, opens a source, requests data, refreshes
/// content, converts content, or executes code.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Dde {
    instruction: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
    kind: DdeKind,
    application: String,
    source: String,
    item: Option<String>,
    automatic_updates: bool,
    representation: Option<DdeFormat>,
    omit_graphic_data: bool,
    switches: Vec<Switch>,
}

impl Dde {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some((kind, application, source, item, switches)) =
            parse_dde_operands_and_switches(field.instruction())?
        else {
            return Ok(None);
        };

        let mut automatic_updates = kind == DdeKind::DdeAuto;
        let mut saw_automatic_update = false;
        let mut representation = None;
        let mut omit_graphic_data = false;
        for switch in &switches {
            match switch.name {
                'a' if kind == DdeKind::Dde => {
                    if saw_automatic_update || switch.argument.is_some() {
                        return Err(Error::Invalid(
                            "DDE \\a switch cannot be repeated or take an argument".to_string(),
                        ));
                    }
                    automatic_updates = true;
                    saw_automatic_update = true;
                },
                'a' => {
                    return Err(Error::Invalid(
                        "DDEAUTO field does not allow a \\a switch".to_string(),
                    ));
                },
                'd' => {
                    if representation.is_some() || omit_graphic_data || switch.argument.is_some() {
                        return Err(Error::Invalid(
                            "DDE result and storage switches cannot be combined".to_string(),
                        ));
                    }
                    omit_graphic_data = true;
                },
                'b' | 'h' | 'p' | 'r' | 't' | 'u' => {
                    if representation.is_some() || omit_graphic_data || switch.argument.is_some() {
                        return Err(Error::Invalid(
                            "DDE result and storage switches cannot be combined".to_string(),
                        ));
                    }
                    representation = Some(match switch.name {
                        'b' => DdeFormat::Bitmap,
                        'h' => DdeFormat::Html,
                        'p' => DdeFormat::Picture,
                        'r' => DdeFormat::RichText,
                        't' => DdeFormat::Text,
                        'u' => DdeFormat::UnicodeText,
                        _ => unreachable!("DDE representation switch was matched above"),
                    });
                },
                _ => {},
            }
        }

        Ok(Some(Self {
            instruction: field.instruction.clone(),
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
            kind,
            application,
            source,
            item,
            automatic_updates,
            representation,
            omit_graphic_data,
            switches,
        }))
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the cached field result, if one was stored.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a word processor marked the cached result stale.
    pub fn is_dirty(&self) -> bool {
        self.dirty
    }

    /// Whether a word processor locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.locked
    }

    /// Return whether this is a DDE or DDEAUTO field.
    pub fn kind(&self) -> DdeKind {
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
    /// This is metadata only. The API never performs an update.
    pub fn requests_automatic_updates(&self) -> bool {
        self.automatic_updates
    }

    /// Return the requested stored result representation, if present.
    ///
    /// This is metadata only and never triggers source access or conversion.
    pub fn representation(&self) -> Option<DdeFormat> {
        self.representation
    }

    /// Whether the stored d switch omits graphic data from the document.
    ///
    /// This is stored metadata only. The API never reads the source to obtain
    /// omitted data.
    pub fn omits_graphic_data(&self) -> bool {
        self.omit_graphic_data
    }

    /// Return all stored field switches in source order.
    pub fn switches(&self) -> &[Switch] {
        &self.switches
    }
}

/// The kind of externally sourced Word field.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum IncludeKind {
    /// An `INCLUDETEXT` or historical `INCLUDE` field that stores a document or XML
    /// source.
    Text,
    /// An `INCLUDEPICTURE` or historical `IMPORT` field that stores an image source.
    Picture,
}

/// One recognized stored option of an external-include field.
///
/// These values are configuration metadata only. This API never opens,
/// resolves, imports, transforms, or evaluates the referenced source.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum IncludeOption {
    /// A document or graphics converter name from the c switch.
    Converter(String),
    /// A source encoding from the INCLUDETEXT e switch.
    Encoding(String),
    /// A source MIME type from the INCLUDETEXT m switch.
    MimeType(String),
    /// An XML namespace mapping from the INCLUDETEXT n switch.
    NamespaceMapping(String),
    /// An XSLT location from the INCLUDETEXT t switch.
    Xslt(String),
    /// An XPath expression from the INCLUDETEXT x switch.
    XPath(String),
}

/// Typed, inert metadata for an `INCLUDETEXT`/`INCLUDEPICTURE` or historical
/// `INCLUDE`/`IMPORT` field.
///
/// Source identifiers, bookmarks, options, and cached results are retained as
/// stored field data. This type never opens, resolves, imports, fetches,
/// refreshes, transforms, converts, evaluates, or executes source content.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Include {
    instruction: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
    kind: IncludeKind,
    source: String,
    bookmark: Option<String>,
    suppress_nested_field_updates: bool,
    omit_picture_data: bool,
    options: Vec<IncludeOption>,
    switches: Vec<Switch>,
}

impl Include {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some((kind, source, bookmark, switches)) =
            parse_external_include_operands_and_switches(field.instruction())?
        else {
            return Ok(None);
        };

        let mut suppress_nested_field_updates = false;
        let mut omit_picture_data = false;
        let mut options = Vec::new();
        for switch in &switches {
            match (kind, switch.name) {
                (IncludeKind::Text, '!') => {
                    if switch.argument.is_some() {
                        return Err(Error::Invalid(
                            "INCLUDETEXT exclamation switch does not take an argument".to_string(),
                        ));
                    }
                    suppress_nested_field_updates = true;
                },
                (IncludeKind::Picture, 'd') => {
                    if switch.argument.is_some() {
                        return Err(Error::Invalid(
                            "INCLUDEPICTURE d switch does not take an argument".to_string(),
                        ));
                    }
                    omit_picture_data = true;
                },
                (_, 'c') => options.push(IncludeOption::Converter(
                    required_external_include_option_argument(switch, kind)?,
                )),
                (IncludeKind::Text, 'e') => options.push(IncludeOption::Encoding(
                    required_external_include_option_argument(switch, kind)?,
                )),
                (IncludeKind::Text, 'm') => options.push(IncludeOption::MimeType(
                    required_external_include_option_argument(switch, kind)?,
                )),
                (IncludeKind::Text, 'n') => options.push(IncludeOption::NamespaceMapping(
                    required_external_include_option_argument(switch, kind)?,
                )),
                (IncludeKind::Text, 't') => options.push(IncludeOption::Xslt(
                    required_external_include_option_argument(switch, kind)?,
                )),
                (IncludeKind::Text, 'x') => options.push(IncludeOption::XPath(
                    required_external_include_option_argument(switch, kind)?,
                )),
                _ => {},
            }
        }

        Ok(Some(Self {
            instruction: field.instruction.clone(),
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
            kind,
            source,
            bookmark,
            suppress_nested_field_updates,
            omit_picture_data,
            options,
            switches,
        }))
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the cached field result, if one was stored.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a word processor marked the cached result stale.
    pub fn is_dirty(&self) -> bool {
        self.dirty
    }

    /// Whether a word processor locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.locked
    }

    /// Return whether this includes text or a picture.
    ///
    /// Text includes use `INCLUDETEXT` or historical `INCLUDE`; picture includes use
    /// `INCLUDEPICTURE` or historical `IMPORT`.
    pub fn kind(&self) -> IncludeKind {
        self.kind
    }

    /// Return the stored source identifier without opening or resolving it.
    pub fn source(&self) -> &str {
        &self.source
    }

    /// Return the optional stored bookmark selector for a text-include field.
    ///
    /// `INCLUDEPICTURE` and `IMPORT` fields do not define a bookmark operand, so this
    /// returns None for picture includes.
    pub fn bookmark(&self) -> Option<&str> {
        self.bookmark.as_deref()
    }

    /// Whether the stored text-include instruction suppresses nested updates.
    ///
    /// This is metadata only. The API never performs an update.
    pub fn suppresses_nested_field_updates(&self) -> bool {
        self.suppress_nested_field_updates
    }

    /// Whether the stored picture-include instruction omits picture data.
    ///
    /// This is stored metadata only. The API never reads the source to obtain
    /// omitted picture data.
    pub fn omits_picture_data(&self) -> bool {
        self.omit_picture_data
    }

    /// Return recognized converter and XML options in stored source order.
    ///
    /// All options are inert metadata. This method never resolves a converter,
    /// opens a source, runs XSLT, or evaluates XPath.
    pub fn options(&self) -> &[IncludeOption] {
        &self.options
    }

    /// Return all stored field switches in source order.
    pub fn switches(&self) -> &[Switch] {
        &self.switches
    }
}

/// Typed, inert metadata for an RD referenced-document field.
///
/// Source identifiers, relative-path settings, switches, and cached results
/// are retained as stored field data. This type never opens, resolves, reads,
/// imports, refreshes, evaluates, or executes the referenced document.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SubDocument {
    instruction: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
    source: String,
    relative_path: bool,
    switches: Vec<Switch>,
}

impl SubDocument {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some((source, switches)) = parse_field_operand_and_switches(field.instruction(), "RD")?
        else {
            return Ok(None);
        };
        let source = source.filter(|value| !value.is_empty()).ok_or_else(|| {
            Error::Invalid("RD field is missing its referenced document path".to_string())
        })?;

        let mut relative_path = false;
        for switch in &switches {
            if switch.name == 'f' {
                if switch.argument.is_some() {
                    return Err(Error::Invalid(
                        "RD \\\\f switch does not take an argument".to_string(),
                    ));
                }
                if relative_path {
                    return Err(Error::Invalid(
                        "RD \\\\f switch cannot be repeated".to_string(),
                    ));
                }
                relative_path = true;
            }
        }

        Ok(Some(Self {
            instruction: field.instruction.clone(),
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
            source,
            relative_path,
            switches,
        }))
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the cached field result, if one was stored.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a word processor marked the cached result stale.
    pub fn is_dirty(&self) -> bool {
        self.dirty
    }

    /// Whether a word processor locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.locked
    }

    /// Return the stored referenced-document path without opening it.
    pub fn source(&self) -> &str {
        &self.source
    }

    /// Whether the stored RD instruction's `\\f` switch requests a path relative to this
    /// document.
    ///
    /// This is metadata only. The API never resolves the path.
    pub fn uses_relative_path(&self) -> bool {
        self.relative_path
    }

    /// Return all stored field switches in source order.
    pub fn switches(&self) -> &[Switch] {
        &self.switches
    }
}

/// One stored result or storage switch for a Word `LINK` field.
///
/// These values describe a linked-object representation or whether graphic data
/// is stored. They never cause a source to be opened, contacted, converted, or
/// displayed.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LinkResult {
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
/// ECMA-376 marks modes `1` and `3` unsupported. Those values, and values
/// outside its defined set, are retained as metadata without applying any
/// formatting.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LinkFormat {
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

/// Typed, inert metadata for a legacy Word `LINK` field.
///
/// Application type, source, item, and all result/formatting switches are
/// retained as stored field data. This type never activates an OLE server,
/// launches an application, opens a source, requests data, refreshes content,
/// converts content, or executes code.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Link {
    instruction: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
    application_type: String,
    source: String,
    item: Option<String>,
    automatic_updates: bool,
    result_options: Vec<LinkResult>,
    formatting_modes: Vec<LinkFormat>,
    switches: Vec<Switch>,
}

impl Link {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some((application_type, source, item, switches)) =
            parse_link_operands_and_switches(field.instruction())?
        else {
            return Ok(None);
        };

        let mut automatic_updates = false;
        let mut result_options = Vec::new();
        let mut formatting_modes = Vec::new();
        for switch in &switches {
            match switch.name {
                'a' => {
                    if switch.argument.is_some() {
                        return Err(Error::Invalid(
                            "LINK \\a switch does not take an argument".to_string(),
                        ));
                    }
                    automatic_updates = true;
                },
                'f' => {
                    let argument = switch.argument.as_deref().ok_or_else(|| {
                        Error::Invalid(
                            "LINK \\f switch requires an integral formatting mode".to_string(),
                        )
                    })?;
                    let value = argument.parse::<i64>().map_err(|_| {
                        Error::Invalid("LINK \\f formatting mode must be an integer".to_string())
                    })?;
                    formatting_modes.push(match value {
                        0 => LinkFormat::Source,
                        2 => LinkFormat::Destination,
                        4 => LinkFormat::SpreadsheetSource,
                        5 => LinkFormat::SpreadsheetDestination,
                        other => LinkFormat::Unsupported(other),
                    });
                },
                'b' | 'd' | 'h' | 'p' | 'r' | 't' | 'u' => {
                    if switch.argument.is_some() {
                        return Err(Error::Invalid(format!(
                            "LINK \\{} switch does not take an argument",
                            switch.name
                        )));
                    }
                    result_options.push(match switch.name {
                        'b' => LinkResult::Bitmap,
                        'd' => LinkResult::OmitGraphicData,
                        'h' => LinkResult::Html,
                        'p' => LinkResult::Picture,
                        'r' => LinkResult::RichText,
                        't' => LinkResult::Text,
                        'u' => LinkResult::UnicodeText,
                        _ => unreachable!("LINK result switch was matched above"),
                    });
                },
                _ => {},
            }
        }

        Ok(Some(Self {
            instruction: field.instruction.clone(),
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
            application_type,
            source,
            item,
            automatic_updates,
            result_options,
            formatting_modes,
            switches,
        }))
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the cached field result, if one was stored.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a word processor marked the cached result stale.
    pub fn is_dirty(&self) -> bool {
        self.dirty
    }

    /// Whether a word processor locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.locked
    }

    /// Return the stored linked-object application type.
    ///
    /// Word commonly stores an OLE Programmatic Identifier here. It is never
    /// looked up or activated by this API.
    pub fn application_type(&self) -> &str {
        &self.application_type
    }

    /// Return the stored source identifier without opening or resolving it.
    pub fn source(&self) -> &str {
        &self.source
    }

    /// Return the optional stored source item, such as a cell range or bookmark.
    pub fn item(&self) -> Option<&str> {
        self.item.as_deref()
    }

    /// Whether the stored instruction requests automatic updates.
    ///
    /// This is metadata only. The API never performs an update.
    pub fn requests_automatic_updates(&self) -> bool {
        self.automatic_updates
    }

    /// Return recognized result and storage switches in stored source order.
    ///
    /// When several are present, [`Self::effective_result_option`] reflects
    /// Word's documented last-switch behavior. Neither method contacts the
    /// linked source.
    pub fn result_options(&self) -> &[LinkResult] {
        &self.result_options
    }

    /// Return the effective result or storage option under Word's documented
    /// last-switch behavior, if one was stored.
    pub fn effective_result_option(&self) -> Option<LinkResult> {
        self.result_options.last().copied()
    }

    /// Return integral `\\f` formatting modes in stored source order.
    ///
    /// These are metadata only; this API never formats linked content.
    pub fn formatting_modes(&self) -> &[LinkFormat] {
        &self.formatting_modes
    }

    /// Return all stored field switches in source order.
    pub fn switches(&self) -> &[Switch] {
        &self.switches
    }
}

/// A typed, inert Word `CITATION` field.
///
/// The field stores one primary bibliography-source tag plus zero or more
/// multi-source tags introduced by `\m`. This model preserves that metadata
/// and cached display text only. It never accesses bibliography source XML,
/// formats citations, resolves locales, or executes field instructions.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Citation {
    instruction: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
    source_tags: Vec<String>,
    switches: Vec<Switch>,
}

impl Citation {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some((primary_source_tag, switches)) =
            parse_citation_operand_and_switches(field.instruction())?
        else {
            return Ok(None);
        };
        let primary_source_tag = primary_source_tag.ok_or_else(|| {
            Error::Invalid("CITATION field is missing its source tag".to_string())
        })?;
        if primary_source_tag.is_empty() {
            return Err(Error::Invalid(
                "CITATION field source tag is empty".to_string(),
            ));
        }

        let mut source_tags = vec![primary_source_tag];
        for switch in &switches {
            if switch.name != 'm' {
                continue;
            }
            let source_tag = switch.argument.as_deref().ok_or_else(|| {
                Error::Invalid("CITATION \\m switch requires a source tag".to_string())
            })?;
            if source_tag.is_empty() {
                return Err(Error::Invalid(
                    "CITATION \\m source tag is empty".to_string(),
                ));
            }
            source_tags.push(source_tag.to_string());
        }

        Ok(Some(Self {
            instruction: field.instruction.clone(),
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
            source_tags,
            switches,
        }))
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the cached formatted citation, if present.
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

    /// Return the primary source tag stored directly after `CITATION`.
    pub fn primary_source_tag(&self) -> &str {
        &self.source_tags[0]
    }

    /// Return primary and `\m` multi-source tags in instruction order.
    pub fn source_tags(&self) -> &[String] {
        &self.source_tags
    }

    /// Return the additional source tags introduced by `\m` switches.
    pub fn additional_source_tags(&self) -> &[String] {
        &self.source_tags[1..]
    }

    /// Return all stored switches in source order.
    ///
    /// Switch semantics can apply to the primary or a preceding `\m` source,
    /// so callers that need producer-specific interpretation should retain this
    /// source order instead of assuming a global setting.
    pub fn switches(&self) -> &[Switch] {
        &self.switches
    }

    /// Check whether a case-insensitive ASCII switch appears in this field.
    pub fn has_switch(&self, name: char) -> bool {
        has_field_switch(&self.switches, name)
    }
}

/// A typed, inert Word `BIBLIOGRAPHY` field.
///
/// This preserves only the stored field instruction, switches, and cached
/// result. It does not discover bibliography sources, apply a style, sort
/// entries, or generate a bibliography.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Bibliography {
    instruction: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
    switches: Vec<Switch>,
}

impl Bibliography {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some(switches) = parse_field_switches(field.instruction(), "BIBLIOGRAPHY")? else {
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

    /// Return the cached visible bibliography result, if present.
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
}

/// A typed, inert Word `DOCVARIABLE` field.
///
/// This preserves a stored variable name, field switches, and cached result.
/// It never reads a document's settings XML, resolves a variable value, or
/// refreshes the field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Variable {
    instruction: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
    variable_name: String,
    switches: Vec<Switch>,
}

impl Variable {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some((variable_name, switches)) =
            parse_field_operand_and_switches(field.instruction(), "DOCVARIABLE")?
        else {
            return Ok(None);
        };
        let variable_name = variable_name.ok_or_else(|| {
            Error::Invalid("DOCVARIABLE field is missing its variable name".to_string())
        })?;
        if variable_name.is_empty() {
            return Err(Error::Invalid(
                "DOCVARIABLE field variable name is empty".to_string(),
            ));
        }

        Ok(Some(Self {
            instruction: field.instruction.clone(),
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
            variable_name,
            switches,
        }))
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored document-variable name without resolving it.
    pub fn variable_name(&self) -> &str {
        &self.variable_name
    }

    /// Return the cached visible field result, if present.
    ///
    /// This is stored text only and is never regenerated from a variable.
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
    ///
    /// DOCVARIABLE has no field-specific switches. Preserved switches are
    /// inert source metadata and are never interpreted.
    pub fn switches(&self) -> &[Switch] {
        &self.switches
    }

    /// Check whether a case-insensitive ASCII switch appears in this field.
    pub fn has_switch(&self, name: char) -> bool {
        has_field_switch(&self.switches, name)
    }
}

/// A typed, inert Word `DOCPROPERTY` field.
///
/// ECMA-376 Part 1 §17.16.5.14 defines one stored document-property name
/// followed by optional field switches. This type exposes that persisted
/// metadata and the cached result only. It never reads core, extended, or
/// custom package properties, resolves a value, or refreshes the field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Property {
    instruction: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
    property_name: String,
    switches: Vec<Switch>,
}

impl Property {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        if !field.is_document_property() {
            return Ok(None);
        }
        if field.instruction().len() > MAX_DOCUMENT_PROPERTY_FIELD_INSTRUCTION_BYTES {
            return Err(Error::Invalid(format!(
                "DOCPROPERTY field instruction exceeds {MAX_DOCUMENT_PROPERTY_FIELD_INSTRUCTION_BYTES} bytes"
            )));
        }
        let Some((property_name, switches)) =
            parse_field_operand_and_switches(field.instruction(), "DOCPROPERTY")?
        else {
            unreachable!("document-property recognition and parsing must agree");
        };
        let property_name = property_name.ok_or_else(|| {
            Error::Invalid("DOCPROPERTY field is missing its property name".to_string())
        })?;
        if property_name.is_empty() {
            return Err(Error::Invalid(
                "DOCPROPERTY field property name is empty".to_string(),
            ));
        }

        Ok(Some(Self {
            instruction: field.instruction.clone(),
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
            property_name,
            switches,
        }))
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored document-property name without resolving it.
    pub fn property_name(&self) -> &str {
        &self.property_name
    }

    /// Return the cached visible field result, if present.
    ///
    /// This is stored text only and is never regenerated from a property.
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
    ///
    /// Preserved switches are inert source metadata and are never interpreted.
    pub fn switches(&self) -> &[Switch] {
        &self.switches
    }

    /// Check whether a case-insensitive ASCII switch appears in this field.
    pub fn has_switch(&self, name: char) -> bool {
        has_field_switch(&self.switches, name)
    }
}

/// Typed, inert metadata for an explicit Word `INFO` field.
///
/// Word permits the `INFO` keyword to be omitted, but that form overlaps
/// standalone document-information fields such as `TITLE`. This type
/// therefore recognizes the unambiguous explicit keyword only. It retains the
/// stored property selector, optional replacement value, switches, cached
/// result, and field state only. It never reads, resolves, modifies, or writes
/// document or template properties, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Info {
    instruction: String,
    information_type: String,
    new_value: Option<String>,
    switches: Vec<Switch>,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
}

impl Info {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some((information_type, new_value, switches)) =
            parse_info_field_parts(field.instruction())?
        else {
            return Ok(None);
        };

        Ok(Some(Self {
            instruction: field.instruction.clone(),
            information_type,
            new_value,
            switches,
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
        }))
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
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

    /// Return preserved switches in source order without interpreting them.
    pub fn switches(&self) -> &[Switch] {
        &self.switches
    }

    /// Check whether a case-insensitive ASCII switch appears in this field.
    pub fn has_switch(&self, name: char) -> bool {
        has_field_switch(&self.switches, name)
    }

    /// Return the cached visible field result, if present.
    ///
    /// This is stored text only and is never regenerated from a property.
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
}

/// The built-in Word document-information field category.
///
/// These fields are defined in ECMA-376 Part 1 §17.16.5. This enum preserves
/// the stored field kind only; it does not resolve document metadata or
/// calculate dates, revisions, or statistics.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum InformationKind {
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

impl InformationKind {
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

    fn from_instruction(instruction: &str) -> Option<Self> {
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
        .find(|kind| field_instruction_remainder(instruction, kind.field_keyword()).is_some())
    }
}

/// Typed, inert metadata for a built-in Word document-information field.
///
/// This type retains the stored kind, field switches, cached result, and field
/// state only. It never reads package properties, reads or modifies host
/// identity data, calculates dates, revisions, or statistics, resolves a
/// value, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Information {
    instruction: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
    kind: InformationKind,
    switches: Vec<Switch>,
}

impl Information {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some(kind) = InformationKind::from_instruction(field.instruction()) else {
            return Ok(None);
        };
        if field.instruction().len() > MAX_DOCUMENT_INFORMATION_FIELD_INSTRUCTION_BYTES {
            return Err(Error::Invalid(format!(
                "{} field instruction exceeds {MAX_DOCUMENT_INFORMATION_FIELD_INSTRUCTION_BYTES} bytes",
                kind.field_keyword()
            )));
        }
        let switches = parse_field_switches(field.instruction(), kind.field_keyword())?
            .expect("document-information recognition and parsing must agree");

        Ok(Some(Self {
            instruction: field.instruction.clone(),
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
            kind,
            switches,
        }))
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the recognized built-in document-information category.
    pub const fn kind(&self) -> InformationKind {
        self.kind
    }

    /// Return the cached visible field result, if present.
    ///
    /// This is stored text only and is never regenerated from package metadata
    /// or a host user profile.
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
    ///
    /// Preserved switches are inert source metadata and are never interpreted.
    pub fn switches(&self) -> &[Switch] {
        &self.switches
    }

    /// Check whether a case-insensitive ASCII switch appears in this field.
    pub fn has_switch(&self, name: char) -> bool {
        has_field_switch(&self.switches, name)
    }
}

/// The built-in Word document-context and runtime field category.
///
/// `FILENAME`, `TEMPLATE`, `DATE`, `TIME`, `PAGE`, `FILESIZE`, `SECTION`, and
/// `SECTIONPAGES` are defined in ECMA-376 Part 1 §17.16.5. This enum preserves
/// the stored field kind only; it does not read a document path, attached
/// template, host filesystem state or file size, current clock, or page and
/// section layout.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ContextKind {
    FileName,
    Template,
    Date,
    Time,
    Page,
    FileSize,
    Section,
    SectionPages,
}

impl ContextKind {
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

    fn from_instruction(instruction: &str) -> Option<Self> {
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
        .find(|kind| field_instruction_remainder(instruction, kind.field_keyword()).is_some())
    }
}

/// Typed, inert metadata for a built-in Word document-context or runtime field.
///
/// This type retains the stored kind, field switches, cached result, and field
/// state only. It never reads a document path, attached template, host
/// filesystem state or file size, current clock, or page and section layout,
/// resolves a value, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Context {
    instruction: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
    kind: ContextKind,
    switches: Vec<Switch>,
}

impl Context {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some(kind) = ContextKind::from_instruction(field.instruction()) else {
            return Ok(None);
        };
        if field.instruction().len() > MAX_DOCUMENT_CONTEXT_FIELD_INSTRUCTION_BYTES {
            return Err(Error::Invalid(format!(
                "{} field instruction exceeds {MAX_DOCUMENT_CONTEXT_FIELD_INSTRUCTION_BYTES} bytes",
                kind.field_keyword()
            )));
        }
        let switches = parse_field_switches(field.instruction(), kind.field_keyword())?
            .expect("document-context recognition and parsing must agree");

        Ok(Some(Self {
            instruction: field.instruction.clone(),
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
            kind,
            switches,
        }))
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the recognized built-in document-context or runtime category.
    pub const fn kind(&self) -> ContextKind {
        self.kind
    }

    /// Return the cached visible field result, if present.
    ///
    /// This is stored text only and is never regenerated from a document path,
    /// attached template, host filesystem state or file size, current clock,
    /// or page and section layout.
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
    ///
    /// Preserved switches are inert source metadata and are never interpreted.
    pub fn switches(&self) -> &[Switch] {
        &self.switches
    }

    /// Check whether a case-insensitive ASCII switch appears in this field.
    pub fn has_switch(&self, name: char) -> bool {
        has_field_switch(&self.switches, name)
    }
}

/// A typed, inert Word `MERGEFIELD` field.
///
/// ECMA-376 Part 1 §17.16.5.35 defines one stored data-column name followed
/// by optional field switches. This type exposes that persisted metadata and
/// the cached result only. It never opens a data source, resolves a record,
/// performs a merge, or refreshes the field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Merge {
    instruction: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
    field_name: String,
    switches: Vec<Switch>,
}

impl Merge {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some((field_name, switches)) =
            parse_field_operand_and_switches(field.instruction(), "MERGEFIELD")?
        else {
            return Ok(None);
        };
        let field_name = field_name.ok_or_else(|| {
            Error::Invalid("MERGEFIELD field is missing its data-column name".to_string())
        })?;
        if field_name.is_empty() {
            return Err(Error::Invalid(
                "MERGEFIELD field data-column name is empty".to_string(),
            ));
        }

        Ok(Some(Self {
            instruction: field.instruction.clone(),
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
            field_name,
            switches,
        }))
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored data-column name without resolving a data source.
    pub fn field_name(&self) -> &str {
        &self.field_name
    }

    /// Return the cached visible field result, if present.
    ///
    /// This is stored text only and is never regenerated by a merge.
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

    /// Return stored field switches in source order without interpreting them.
    pub fn switches(&self) -> &[Switch] {
        &self.switches
    }

    /// Check whether a case-insensitive ASCII switch appears in this field.
    pub fn has_switch(&self, name: char) -> bool {
        has_field_switch(&self.switches, name)
    }
}

/// A typed, inert Word `DATA` mail-merge source field.
///
/// [MS-DOC] §2.9.90 specifies `DATA datafile [headerfile]` as a field that
/// redirects mail-merge data and header files. This type exposes only those
/// stored operands, switches, cached result, and field state. It never opens,
/// reads, connects to, resolves, or modifies either source; it never selects a
/// record, performs a merge, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MergeData {
    instruction: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
    data_source: String,
    header_source: Option<String>,
    switches: Vec<Switch>,
}

impl MergeData {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some((data_source, header_source, switches)) =
            parse_mail_merge_data_field_parts(field.instruction())?
        else {
            return Ok(None);
        };

        Ok(Some(Self {
            instruction: field.instruction.clone(),
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
            data_source,
            header_source,
            switches,
        }))
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
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

    /// Return the cached visible field result, if present.
    ///
    /// This is stored text only and is never regenerated by a merge.
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

    /// Return stored field switches in source order without interpreting them.
    pub fn switches(&self) -> &[Switch] {
        &self.switches
    }

    /// Check whether a case-insensitive ASCII switch appears in this field.
    pub fn has_switch(&self, name: char) -> bool {
        has_field_switch(&self.switches, name)
    }
}

/// The stored kind of a mail-merge counter field.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MergeCounterKind {
    /// A `MERGEREC` field, which stores a selected-record position.
    Record,
    /// A `MERGESEQ` field, which stores a merged-record sequence position.
    Sequence,
}

/// A typed, inert Word `MERGEREC` or `MERGESEQ` field.
///
/// ECMA-376 Part 1 §§17.16.5.36–37 define these zero-argument fields. This
/// type exposes the persisted kind and cached result only. It never selects or
/// counts records, opens a data source, performs a merge, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MergeCounter {
    instruction: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
    kind: MergeCounterKind,
}

impl MergeCounter {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let (kind, field_type, remainder) = if let Some(remainder) =
            field_instruction_remainder(field.instruction(), "MERGEREC")
        {
            (MergeCounterKind::Record, "MERGEREC", remainder)
        } else if let Some(remainder) = field_instruction_remainder(field.instruction(), "MERGESEQ")
        {
            (MergeCounterKind::Sequence, "MERGESEQ", remainder)
        } else {
            return Ok(None);
        };

        if !remainder.trim().is_empty() {
            return Err(Error::Invalid(format!(
                "{field_type} field must not contain arguments or switches"
            )));
        }

        Ok(Some(Self {
            instruction: field.instruction.clone(),
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
            kind,
        }))
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return whether this is a `MERGEREC` or `MERGESEQ` field.
    pub fn kind(&self) -> MergeCounterKind {
        self.kind
    }

    /// Return the cached visible field result, if present.
    ///
    /// This is stored text only and is never regenerated by a merge.
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
}

/// A typed, inert Word `NEXT` mail-merge control field.
///
/// ECMA-376 Part 1 §17.16.5.38 defines `NEXT` as a zero-argument
/// instruction. This type exposes persisted cached content and state only. It
/// never advances a record, opens a data source, performs a merge, or refreshes
/// a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MergeNext {
    instruction: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
}

impl MergeNext {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some(remainder) = field_instruction_remainder(field.instruction(), "NEXT") else {
            return Ok(None);
        };
        if !remainder.trim().is_empty() {
            return Err(Error::Invalid(
                "NEXT field must not contain arguments or switches".to_string(),
            ));
        }

        Ok(Some(Self {
            instruction: field.instruction.clone(),
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
        }))
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the cached visible field result, if present.
    ///
    /// This is stored text only and is never regenerated by a merge.
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
}

/// The stored kind of a conditional mail-merge control field.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MergeControlKind {
    /// A `NEXTIF` field, which can advance a merge record when its comparison is true.
    NextIf,
    /// A `SKIPIF` field, which can omit a merge record when its comparison is true.
    SkipIf,
}

/// A typed, inert Word `NEXTIF` or `SKIPIF` mail-merge control field.
///
/// ECMA-376 Part 1 §§17.16.5.39 and 17.16.5.58 define these controls. This
/// type retains the unparsed comparison and cached result only. It never parses or
/// evaluates a comparison, advances or skips a record, opens a data source,
/// performs a merge, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MergeControl {
    instruction: String,
    comparison: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
    kind: MergeControlKind,
}

impl MergeControl {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let (kind, field_type, comparison) = if let Some(comparison) =
            field_instruction_remainder(field.instruction(), "NEXTIF")
        {
            (MergeControlKind::NextIf, "NEXTIF", comparison)
        } else if let Some(comparison) = field_instruction_remainder(field.instruction(), "SKIPIF")
        {
            (MergeControlKind::SkipIf, "SKIPIF", comparison)
        } else {
            return Ok(None);
        };

        let comparison = comparison.trim();
        if comparison.is_empty() {
            return Err(Error::Invalid(format!(
                "{field_type} field is missing its comparison"
            )));
        }

        Ok(Some(Self {
            instruction: field.instruction.clone(),
            comparison: comparison.to_string(),
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
            kind,
        }))
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return whether this is a `NEXTIF` or `SKIPIF` control.
    pub fn kind(&self) -> MergeControlKind {
        self.kind
    }

    /// Return the stored comparison without parsing or evaluating it.
    pub fn comparison(&self) -> &str {
        &self.comparison
    }

    /// Return the cached visible field result, if present.
    ///
    /// This is stored text only and is never regenerated by a merge.
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
}

/// A typed, inert Word `IF` field.
///
/// ECMA-376 Part 1 §17.16.5.26 defines `IF` using a comparison and two
/// branches. This type retains the unparsed expression and cached result only.
/// It never parses or evaluates an expression, resolves field values, or
/// refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct If {
    instruction: String,
    expression: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
}

impl If {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some(expression) = field_instruction_remainder(field.instruction(), "IF") else {
            return Ok(None);
        };
        let expression = expression.trim();
        if expression.is_empty() {
            return Err(Error::Invalid(
                "IF field is missing its expression".to_string(),
            ));
        }

        Ok(Some(Self {
            instruction: field.instruction.clone(),
            expression: expression.to_string(),
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
        }))
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored expression without parsing or evaluating it.
    pub fn expression(&self) -> &str {
        &self.expression
    }

    /// Return the cached visible field result, if present.
    ///
    /// This is stored text only and is never regenerated by field evaluation.
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
}

/// A typed, inert Word `SET` field.
///
/// ECMA-376 Part 1 §17.16.5.57 defines `SET` using a target name and an
/// expression. This type retains the stored target, opaque expression, and
/// cached result only. It never evaluates an expression, looks up or changes a
/// bookmark, changes document state, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Set {
    instruction: String,
    target_name: String,
    expression: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
}

impl Set {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some((target_name, expression)) = parse_set_field_parts(field.instruction())? else {
            return Ok(None);
        };

        Ok(Some(Self {
            instruction: field.instruction.clone(),
            target_name,
            expression,
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
        }))
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored target name without looking it up or changing it.
    pub fn target_name(&self) -> &str {
        &self.target_name
    }

    /// Return the opaque stored expression text.
    ///
    /// This text is never parsed, evaluated, or used to change document state.
    pub fn expression(&self) -> &str {
        &self.expression
    }

    /// Return the cached visible field result, if present.
    ///
    /// This is stored text only and is never regenerated by expression
    /// evaluation.
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
}

/// A typed, inert Word `SEQ` field.
///
/// ECMA-376 Part 1 §17.16.5.56 defines `SEQ` using an identifier, optional
/// bookmark, and optional switches. This type retains those stored values and a
/// cached result only. It never looks up a bookmark, increments or resets a
/// sequence, calculates a number, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Sequence {
    instruction: String,
    identifier: String,
    bookmark: Option<String>,
    tail: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
}

impl Sequence {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some((identifier, bookmark, tail)) = parse_sequence_field_parts(field.instruction())?
        else {
            return Ok(None);
        };

        Ok(Some(Self {
            instruction: field.instruction.clone(),
            identifier,
            bookmark,
            tail,
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
        }))
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
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
    pub fn tail(&self) -> &str {
        &self.tail
    }

    /// Return the cached visible field result, if present.
    ///
    /// This is stored text only and is never regenerated by calculating a
    /// sequence number.
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
}

/// A typed, inert Word `=` formula field.
///
/// ECMA-376 Part 1 §17.16.3.3 defines table formulas using a leading `=`.
/// This type retains the stored formula text and cached result only. It never
/// parses or evaluates a formula, reads table cells or bookmarks, resolves
/// field values, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Formula {
    instruction: String,
    formula: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
}

impl Formula {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some(formula) = parse_formula_field_formula(field.instruction())? else {
            return Ok(None);
        };

        Ok(Some(Self {
            instruction: field.instruction.clone(),
            formula,
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
        }))
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the opaque stored formula text without parsing or evaluating it.
    pub fn formula(&self) -> &str {
        &self.formula
    }

    /// Return the cached visible field result, if present.
    ///
    /// This is stored text only and is never regenerated by formula evaluation.
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
}

/// A typed, inert Word `EQ` equation field.
///
/// This type retains only stored equation syntax, a cached result, and field
/// state. It never parses, calculates, formats, renders, or refreshes an
/// equation.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Equation {
    instruction: String,
    expression: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
}

impl Equation {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some(expression) = field_instruction_remainder(field.instruction(), "EQ") else {
            return Ok(None);
        };
        if field.instruction().len() > MAX_EQUATION_FIELD_INSTRUCTION_BYTES {
            return Err(Error::Invalid(format!(
                "EQ field instruction exceeds {MAX_EQUATION_FIELD_INSTRUCTION_BYTES} bytes"
            )));
        }

        Ok(Some(Self {
            instruction: field.instruction.clone(),
            expression: expression.trim().to_string(),
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
        }))
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the opaque equation expression after the `EQ` keyword.
    ///
    /// This syntax is never parsed, calculated, formatted, or rendered.
    pub fn expression(&self) -> &str {
        &self.expression
    }

    /// Return the cached visible field result, if present.
    ///
    /// This is stored text only and is never regenerated from equation syntax.
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
}

/// A typed, inert Word `HYPERLINK` field.
///
/// This type retains only stored link metadata, a cached result, and field
/// state. It never opens, resolves, follows, activates, or refreshes a link.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Hyperlink {
    instruction: String,
    external_target: Option<String>,
    bookmark: Option<String>,
    screen_tip: Option<String>,
    target_frame: Option<String>,
    appends_image_map_coordinates: bool,
    opens_new_window: bool,
    unknown_switches: Vec<Switch>,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
}

impl Hyperlink {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        if !field.is_hyperlink_field() {
            return Ok(None);
        }
        if field.instruction().len() > MAX_HYPERLINK_FIELD_INSTRUCTION_BYTES {
            return Err(Error::Invalid(format!(
                "HYPERLINK field instruction exceeds {MAX_HYPERLINK_FIELD_INSTRUCTION_BYTES} bytes"
            )));
        }
        let Some((external_target, switches)) =
            parse_field_operand_and_switches(field.instruction(), "HYPERLINK")?
        else {
            unreachable!("hyperlink-field recognition and parsing must agree");
        };
        let external_target = external_target
            .map(|target| {
                (!target.is_empty()).then_some(target).ok_or_else(|| {
                    Error::Invalid("HYPERLINK external target must not be empty".to_string())
                })
            })
            .transpose()?;

        let mut bookmark = None;
        let mut screen_tip = None;
        let mut target_frame = None;
        let mut appends_image_map_coordinates = false;
        let mut opens_new_window = false;
        let mut unknown_switches = Vec::new();
        for switch in switches {
            let (slot, switch_name) = match switch.name {
                'l' => (&mut bookmark, 'l'),
                'o' => (&mut screen_tip, 'o'),
                't' => (&mut target_frame, 't'),
                'm' => {
                    if appends_image_map_coordinates {
                        return Err(Error::Invalid(
                            "HYPERLINK \\m switch is duplicated".to_string(),
                        ));
                    }
                    if switch.argument.is_some() {
                        return Err(Error::Invalid(
                            "HYPERLINK \\m switch does not take an argument".to_string(),
                        ));
                    }
                    appends_image_map_coordinates = true;
                    continue;
                },
                'n' => {
                    if opens_new_window {
                        return Err(Error::Invalid(
                            "HYPERLINK field has duplicate \\n switches".to_string(),
                        ));
                    }
                    if switch.argument.is_some() {
                        return Err(Error::Invalid(
                            "HYPERLINK \\n switch does not take an argument".to_string(),
                        ));
                    }
                    opens_new_window = true;
                    continue;
                },
                _ => {
                    unknown_switches.push(switch);
                    continue;
                },
            };
            let value = switch.argument.ok_or_else(|| {
                Error::Invalid(format!(
                    "HYPERLINK \\{switch_name} switch requires an argument"
                ))
            })?;
            if value.is_empty() {
                return Err(Error::Invalid(format!(
                    "HYPERLINK \\{switch_name} switch argument must not be empty"
                )));
            }
            if slot.replace(value).is_some() {
                return Err(Error::Invalid(format!(
                    "HYPERLINK field has duplicate \\{switch_name} switches"
                )));
            }
        }
        if external_target.is_none() && bookmark.is_none() {
            return Err(Error::Invalid(
                "HYPERLINK field requires an external target or \\l bookmark".to_string(),
            ));
        }

        Ok(Some(Self {
            instruction: field.instruction.clone(),
            external_target,
            bookmark,
            screen_tip,
            target_frame,
            appends_image_map_coordinates,
            opens_new_window,
            unknown_switches,
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
        }))
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored external target without resolving or opening it.
    pub fn external_target(&self) -> Option<&str> {
        self.external_target.as_deref()
    }

    /// Return the stored internal bookmark target without resolving it.
    pub fn bookmark(&self) -> Option<&str> {
        self.bookmark.as_deref()
    }

    /// Return the stored screen-tip text, if present.
    ///
    /// This is metadata only and is never displayed by the library.
    pub fn screen_tip(&self) -> Option<&str> {
        self.screen_tip.as_deref()
    }

    /// Return the stored target frame, if present.
    ///
    /// This is metadata only and is never used to open a window or frame.
    pub fn target_frame(&self) -> Option<&str> {
        self.target_frame.as_deref()
    }

    /// Whether the target receives click coordinates for a server-side image map.
    ///
    /// This records producer intent only; no navigation or hit testing occurs.
    pub fn appends_image_map_coordinates(&self) -> bool {
        self.appends_image_map_coordinates
    }

    /// Whether the field requests opening the target in a new window.
    ///
    /// This records producer intent only; no window is opened.
    pub fn opens_new_window(&self) -> bool {
        self.opens_new_window
    }

    /// Return unrecognized stored switches in source order.
    ///
    /// They are retained without interpretation or execution.
    pub fn unknown_switches(&self) -> &[Switch] {
        &self.unknown_switches
    }

    /// Return the cached visible field result, if present.
    ///
    /// This is stored text only and is never regenerated by resolving a link.
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
}

/// A typed, inert Word `QUOTE` field.
///
/// ECMA-376 Part 1 §17.16.5.49 defines `QUOTE` with a text field argument.
/// This type retains that stored argument, switches, and cached result only. It
/// never interprets character codes, expands nested fields, inserts text, or
/// refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Quote {
    instruction: String,
    text: String,
    switches: Vec<Switch>,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
}

impl Quote {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        if field.instruction.len() > MAX_QUOTE_FIELD_INSTRUCTION_BYTES {
            return Err(Error::Invalid(format!(
                "QUOTE field instruction exceeds {MAX_QUOTE_FIELD_INSTRUCTION_BYTES} bytes"
            )));
        }
        let Some((text, switches)) =
            parse_field_operand_and_switches(field.instruction(), "QUOTE")?
        else {
            return Ok(None);
        };
        let text = text.ok_or_else(|| {
            Error::Invalid("QUOTE field is missing its text argument".to_string())
        })?;

        Ok(Some(Self {
            instruction: field.instruction.clone(),
            text,
            switches,
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
        }))
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored text argument without inserting or transforming it.
    pub fn text(&self) -> &str {
        &self.text
    }

    /// Return preserved switches in source order without interpreting them.
    pub fn switches(&self) -> &[Switch] {
        &self.switches
    }

    /// Check whether a case-insensitive ASCII switch appears in this field.
    pub fn has_switch(&self, name: char) -> bool {
        has_field_switch(&self.switches, name)
    }

    /// Return the cached visible field result, if present.
    ///
    /// This is stored text only and is never regenerated by inserting text.
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
}

/// A typed, inert Word `SYMBOL` field.
///
/// ECMA-376 Part 1 §17.16.5.61 defines `SYMBOL` with one stored character
/// argument and optional switches. This type retains that argument, switches,
/// and cached result only. It never converts a character code, looks up a font,
/// inserts a glyph, changes formatting or layout, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Symbol {
    instruction: String,
    character_argument: String,
    switches: Vec<Switch>,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
}

impl Symbol {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        if field.instruction.len() > MAX_SYMBOL_FIELD_INSTRUCTION_BYTES {
            return Err(Error::Invalid(format!(
                "SYMBOL field instruction exceeds {MAX_SYMBOL_FIELD_INSTRUCTION_BYTES} bytes"
            )));
        }
        let Some((character_argument, switches)) =
            parse_field_operand_and_switches(field.instruction(), "SYMBOL")?
        else {
            return Ok(None);
        };
        let character_argument = character_argument.ok_or_else(|| {
            Error::Invalid("SYMBOL field is missing its character argument".to_string())
        })?;

        Ok(Some(Self {
            instruction: field.instruction.clone(),
            character_argument,
            switches,
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
        }))
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored character argument without converting it to a glyph.
    pub fn character_argument(&self) -> &str {
        &self.character_argument
    }

    /// Return preserved switches in source order without interpreting them.
    pub fn switches(&self) -> &[Switch] {
        &self.switches
    }

    /// Check whether a case-insensitive ASCII switch appears in this field.
    pub fn has_switch(&self, name: char) -> bool {
        has_field_switch(&self.switches, name)
    }

    /// Return the cached visible field result, if present.
    ///
    /// This is stored text only and is never regenerated by mapping a character
    /// code or inserting a glyph.
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
}

/// The legacy Word automatic-numbering field category.
///
/// `AUTONUM`, `AUTONUMLGL`, and `AUTONUMOUT` are retained for
/// document compatibility. This enum preserves the stored field kind only; it
/// does not calculate a number, inspect paragraphs, headings, or styles, or
/// change layout.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum AutoNumberKind {
    AutoNum,
    AutoNumLegal,
    AutoNumOutline,
}

impl AutoNumberKind {
    /// The uppercase field keyword stored in a Word field instruction.
    pub const fn field_keyword(self) -> &'static str {
        match self {
            Self::AutoNum => "AUTONUM",
            Self::AutoNumLegal => "AUTONUMLGL",
            Self::AutoNumOutline => "AUTONUMOUT",
        }
    }

    fn from_instruction(instruction: &str) -> Option<Self> {
        [Self::AutoNum, Self::AutoNumLegal, Self::AutoNumOutline]
            .into_iter()
            .find(|kind| field_instruction_remainder(instruction, kind.field_keyword()).is_some())
    }
}

/// Typed, inert metadata for a legacy Word automatic-numbering field.
///
/// This type retains the stored kind, switches, cached result, and field state
/// only. It never calculates paragraph numbers, reads heading or style state,
/// changes paragraphs or layout, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AutoNumber {
    instruction: String,
    kind: AutoNumberKind,
    switches: Vec<Switch>,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
}

impl AutoNumber {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some(kind) = AutoNumberKind::from_instruction(field.instruction()) else {
            return Ok(None);
        };
        if field.instruction().len() > MAX_AUTO_NUMBER_FIELD_INSTRUCTION_BYTES {
            return Err(Error::Invalid(format!(
                "{} field instruction exceeds {MAX_AUTO_NUMBER_FIELD_INSTRUCTION_BYTES} bytes",
                kind.field_keyword()
            )));
        }
        let switches = parse_field_switches(field.instruction(), kind.field_keyword())?
            .expect("automatic-number recognition and parsing must agree");

        Ok(Some(Self {
            instruction: field.instruction.clone(),
            kind,
            switches,
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
        }))
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the recognized automatic-numbering category.
    pub const fn kind(&self) -> AutoNumberKind {
        self.kind
    }

    /// Return preserved switches in source order without interpreting them.
    pub fn switches(&self) -> &[Switch] {
        &self.switches
    }

    /// Check whether a case-insensitive ASCII switch appears in this field.
    pub fn has_switch(&self, name: char) -> bool {
        has_field_switch(&self.switches, name)
    }

    /// Return the stored cached field result, if present.
    ///
    /// This is stored text only and is never regenerated by calculating a
    /// paragraph number.
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
}

/// Typed, inert metadata for a Word `LISTNUM` field.
///
/// ECMA-376 Part 1 §17.16.5.33 defines `LISTNUM` with an optional list
/// name and switches. This type retains those stored values and cached result
/// only. It never looks up a list, determines a level or start value,
/// calculates a number, changes layout, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ListNumber {
    instruction: String,
    list_name: Option<String>,
    switches: Vec<Switch>,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
}

impl ListNumber {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        if field.instruction().len() > MAX_LIST_NUMBER_FIELD_INSTRUCTION_BYTES {
            return Err(Error::Invalid(format!(
                "LISTNUM field instruction exceeds {MAX_LIST_NUMBER_FIELD_INSTRUCTION_BYTES} bytes"
            )));
        }
        let Some((list_name, switches)) =
            parse_field_operand_and_switches(field.instruction(), "LISTNUM")?
        else {
            return Ok(None);
        };

        Ok(Some(Self {
            instruction: field.instruction.clone(),
            list_name,
            switches,
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
        }))
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored optional list name without looking it up.
    pub fn list_name(&self) -> Option<&str> {
        self.list_name.as_deref()
    }

    /// Return preserved switches in source order without interpreting them.
    pub fn switches(&self) -> &[Switch] {
        &self.switches
    }

    /// Check whether a case-insensitive ASCII switch appears in this field.
    pub fn has_switch(&self, name: char) -> bool {
        has_field_switch(&self.switches, name)
    }

    /// Return the stored cached field result, if present.
    ///
    /// This is stored text only and is never regenerated by calculating a list
    /// number.
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
}

/// One recognized stored option of a Word `STYLEREF` field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum StyleOption {
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
pub struct StyleReference {
    instruction: String,
    style_name: String,
    options: Vec<StyleOption>,
    unknown_switches: Vec<Switch>,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
}

impl StyleReference {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some((style_name, options, unknown_switches)) =
            parse_style_reference_field_parts(field.instruction())?
        else {
            return Ok(None);
        };

        Ok(Some(Self {
            instruction: field.instruction.clone(),
            style_name,
            options,
            unknown_switches,
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
        }))
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored style name without looking up text that uses it.
    pub fn style_name(&self) -> &str {
        &self.style_name
    }

    /// Return recognized stored options in source order.
    ///
    /// This metadata is never used to search text, calculate a number, or
    /// resolve layout.
    pub fn options(&self) -> &[StyleOption] {
        &self.options
    }

    /// Return unrecognized stored switches in source order.
    ///
    /// They are retained without interpretation or execution.
    pub fn unknown_switches(&self) -> &[Switch] {
        &self.unknown_switches
    }

    /// Return the cached visible field result, if present.
    ///
    /// This is stored text only and is never regenerated by searching styled
    /// text or resolving layout.
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
}

/// A typed, inert Word `COMPARE` field.
///
/// ECMA-376 Part 1 §17.16.5.10 defines `COMPARE` using a comparison whose
/// result is 1 or 0. This type retains the unparsed comparison and cached
/// result only. It never parses or evaluates a comparison, resolves nested
/// field values, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Compare {
    instruction: String,
    comparison: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
}

impl Compare {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some(comparison) = field_instruction_remainder(field.instruction(), "COMPARE") else {
            return Ok(None);
        };
        let comparison = comparison.trim();
        if comparison.is_empty() {
            return Err(Error::Invalid(
                "COMPARE field is missing its comparison".to_string(),
            ));
        }

        Ok(Some(Self {
            instruction: field.instruction.clone(),
            comparison: comparison.to_string(),
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
        }))
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored comparison without parsing or evaluating it.
    pub fn comparison(&self) -> &str {
        &self.comparison
    }

    /// Return the cached visible field result, if present.
    ///
    /// This is stored text only and is never regenerated by field evaluation.
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
}

/// The stored category of a Word bookmark-reference field.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ReferenceKind {
    /// A `REF` field.
    Reference,
    /// A `PAGEREF` field.
    PageReference,
    /// A historical `FTNREF` field.
    FootnoteReference,
    /// A `NOTEREF` field.
    NoteReference,
}

impl ReferenceKind {
    fn from_instruction(instruction: &str) -> Option<(Self, &'static str)> {
        for (kind, field_type) in [
            (Self::Reference, "REF"),
            (Self::PageReference, "PAGEREF"),
            (Self::FootnoteReference, "FTNREF"),
            (Self::NoteReference, "NOTEREF"),
        ] {
            if field_instruction_remainder(instruction, field_type).is_some() {
                return Some((kind, field_type));
            }
        }
        None
    }

    fn is_note_reference(self) -> bool {
        matches!(self, Self::FootnoteReference | Self::NoteReference)
    }
}

/// One recognized stored option of a Word bookmark-reference field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ReferenceOption {
    /// The `\d` `REF` separator between sequence and page numbers.
    SequencePageSeparator(String),
    /// The `\f` `REF` request for referenced note or comment content.
    ReferencedNoteContent,
    /// The `\h` request for a link to the stored bookmark.
    Hyperlink,
    /// The `\n` `REF` request for a paragraph number without context.
    ParagraphNumberWithoutContext,
    /// The `\p` request for relative-position text.
    RelativePosition,
    /// The `\r` `REF` request for a paragraph number in relative context.
    ParagraphNumberRelativeContext,
    /// The `\t` `REF` request to suppress non-number text.
    SuppressNonNumberText,
    /// The `\w` `REF` request for a paragraph number in full context.
    ParagraphNumberFullContext,
    /// The `\f` `FTNREF` or `NOTEREF` request to format the note mark.
    NoteMarkFormatting,
}

/// A typed, inert Word bookmark-reference field.
///
/// This model preserves only stored categories, targets, options, switches,
/// cached results, and field state. It never looks up a bookmark, reads a
/// referenced range or note, resolves a page number, creates a link,
/// calculates a relative position, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Reference {
    instruction: String,
    kind: ReferenceKind,
    bookmark: String,
    options: Vec<ReferenceOption>,
    unknown_switches: Vec<Switch>,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
}

impl Reference {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some((kind, field_type)) = ReferenceKind::from_instruction(field.instruction()) else {
            return Ok(None);
        };
        if field.instruction().len() > MAX_REFERENCE_FIELD_INSTRUCTION_BYTES {
            return Err(Error::Invalid(format!(
                "{field_type} field instruction exceeds {MAX_REFERENCE_FIELD_INSTRUCTION_BYTES} bytes"
            )));
        }
        let Some((bookmark, switches)) =
            parse_field_operand_and_switches(field.instruction(), field_type)?
        else {
            unreachable!("bookmark-reference recognition and parsing must agree");
        };
        let bookmark = bookmark
            .filter(|bookmark| !bookmark.is_empty())
            .ok_or_else(|| {
                Error::Invalid(format!("{field_type} field is missing its bookmark target"))
            })?;

        let mut options = Vec::new();
        let mut unknown_switches = Vec::new();
        for switch in switches {
            match switch.name {
                'd' if kind == ReferenceKind::Reference => {
                    let separator = switch.argument.ok_or_else(|| {
                        Error::Invalid("REF \\d switch requires a separator".to_string())
                    })?;
                    options.push(ReferenceOption::SequencePageSeparator(separator));
                },
                'f' if kind == ReferenceKind::Reference => {
                    if switch.argument.is_some() {
                        return Err(Error::Invalid(
                            "REF \\f switch does not take an argument".to_string(),
                        ));
                    }
                    options.push(ReferenceOption::ReferencedNoteContent);
                },
                'f' if kind.is_note_reference() => {
                    if switch.argument.is_some() {
                        return Err(Error::Invalid(format!(
                            "{field_type} \\f switch does not take an argument"
                        )));
                    }
                    options.push(ReferenceOption::NoteMarkFormatting);
                },
                'h' => {
                    if switch.argument.is_some() {
                        return Err(Error::Invalid(format!(
                            "{field_type} \\h switch does not take an argument"
                        )));
                    }
                    options.push(ReferenceOption::Hyperlink);
                },
                'n' if kind == ReferenceKind::Reference => {
                    if switch.argument.is_some() {
                        return Err(Error::Invalid(
                            "REF \\n switch does not take an argument".to_string(),
                        ));
                    }
                    options.push(ReferenceOption::ParagraphNumberWithoutContext);
                },
                'p' => {
                    if switch.argument.is_some() {
                        return Err(Error::Invalid(format!(
                            "{field_type} \\p switch does not take an argument"
                        )));
                    }
                    options.push(ReferenceOption::RelativePosition);
                },
                'r' if kind == ReferenceKind::Reference => {
                    if switch.argument.is_some() {
                        return Err(Error::Invalid(
                            "REF \\r switch does not take an argument".to_string(),
                        ));
                    }
                    options.push(ReferenceOption::ParagraphNumberRelativeContext);
                },
                't' if kind == ReferenceKind::Reference => {
                    if switch.argument.is_some() {
                        return Err(Error::Invalid(
                            "REF \\t switch does not take an argument".to_string(),
                        ));
                    }
                    options.push(ReferenceOption::SuppressNonNumberText);
                },
                'w' if kind == ReferenceKind::Reference => {
                    if switch.argument.is_some() {
                        return Err(Error::Invalid(
                            "REF \\w switch does not take an argument".to_string(),
                        ));
                    }
                    options.push(ReferenceOption::ParagraphNumberFullContext);
                },
                _ => unknown_switches.push(switch),
            }
        }

        Ok(Some(Self {
            instruction: field.instruction.clone(),
            kind,
            bookmark,
            options,
            unknown_switches,
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
        }))
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored reference-field category.
    pub const fn kind(&self) -> ReferenceKind {
        self.kind
    }

    /// Return the stored bookmark or note target without resolving it.
    pub fn bookmark(&self) -> &str {
        &self.bookmark
    }

    /// Return recognized stored options in source order.
    ///
    /// This metadata is never used to navigate, resolve, or activate a link.
    pub fn options(&self) -> &[ReferenceOption] {
        &self.options
    }

    /// Return unrecognized stored switches in source order.
    ///
    /// They are retained without interpretation or execution.
    pub fn unknown_switches(&self) -> &[Switch] {
        &self.unknown_switches
    }

    /// Return the cached visible field result, if present.
    ///
    /// This is stored text only and is never regenerated by resolving a
    /// bookmark, page number, or note reference.
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
}

/// The stored kind of a prompt field.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PromptKind {
    /// An `ASK` field that associates a response with a bookmark.
    Ask,
    /// A `FILLIN` field that stores a response as its field result.
    FillIn,
}

/// A typed, inert Word `ASK` or `FILLIN` prompt field.
///
/// ECMA-376 Part 1 §§17.16.5.3 and 17.16.5.19 define these fields. This type
/// exposes stored prompt and default-response metadata only. It never displays
/// a prompt, captures a response, creates or updates a bookmark, performs a
/// merge, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Prompt {
    instruction: String,
    kind: PromptKind,
    bookmark: Option<String>,
    prompt: Option<String>,
    default_response: Option<String>,
    prompts_once_per_mail_merge: bool,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
}

impl Prompt {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some((kind, bookmark, prompt, default_response, prompts_once_per_mail_merge)) =
            parse_prompt_field_parts(field.instruction())?
        else {
            return Ok(None);
        };

        Ok(Some(Self {
            instruction: field.instruction.clone(),
            kind,
            bookmark,
            prompt,
            default_response,
            prompts_once_per_mail_merge,
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
        }))
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return whether this is an `ASK` or `FILLIN` field.
    pub fn kind(&self) -> PromptKind {
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
    pub fn prompts_once_per_mail_merge(&self) -> bool {
        self.prompts_once_per_mail_merge
    }

    /// Return the cached visible field result, if present.
    ///
    /// This is stored text only and is never regenerated by field evaluation.
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
}

/// The stored kind of a mail-merge recipient layout field.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RecipientKind {
    /// An `ADDRESSBLOCK` field.
    AddressBlock,
    /// A `GREETINGLINE` field.
    GreetingLine,
}

/// How an `ADDRESSBLOCK` field requests country/region text.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CountryInclusion {
    /// `\\c 0` omits the country/region.
    Omit,
    /// `\\c 1` includes the country/region regardless of exclusions.
    Always,
    /// `\\c 2` includes the country/region unless it matches an excluded
    /// country/region.
    UnlessExcluded,
}

/// A typed, inert Word `ADDRESSBLOCK` or `GREETINGLINE` field.
///
/// ECMA-376 Part 1 §§17.16.5.1 and 17.16.5.24 define these mail-merge
/// recipient layout fields. This type exposes stored layout metadata and the
/// cached result only. It never opens a data source, selects a record,
/// performs a merge, expands placeholders, generates text, or refreshes a
/// field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Recipient {
    instruction: String,
    kind: RecipientKind,
    country_inclusion: Option<CountryInclusion>,
    formats_using_recipient_country: bool,
    excluded_countries: Vec<String>,
    format_template: Option<String>,
    language: Option<String>,
    greeting_fallback_text: Option<String>,
    unknown_switches: Vec<Switch>,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
}

impl Recipient {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some((
            kind,
            country_inclusion,
            formats_using_recipient_country,
            excluded_countries,
            format_template,
            language,
            greeting_fallback_text,
            unknown_switches,
        )) = parse_mail_merge_recipient_field_parts(field.instruction())?
        else {
            return Ok(None);
        };

        Ok(Some(Self {
            instruction: field.instruction.clone(),
            kind,
            country_inclusion,
            formats_using_recipient_country,
            excluded_countries,
            format_template,
            language,
            greeting_fallback_text,
            unknown_switches,
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
        }))
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return whether this is an `ADDRESSBLOCK` or `GREETINGLINE` field.
    pub fn kind(&self) -> RecipientKind {
        self.kind
    }

    /// Return how an `ADDRESSBLOCK` requests country/region text.
    ///
    /// This is `None` when the instruction has no `\\c` switch or when the
    /// field is a `GREETINGLINE`. The stored request is never used to render
    /// an address.
    pub fn country_inclusion(&self) -> Option<CountryInclusion> {
        self.country_inclusion
    }

    /// Whether an `ADDRESSBLOCK` stores the `\\d` request to use the
    /// recipient country's address format.
    ///
    /// This request is metadata only and never causes a record or country
    /// format to be resolved.
    pub fn formats_using_recipient_country(&self) -> bool {
        self.formats_using_recipient_country
    }

    /// Return country/region names excluded by an `ADDRESSBLOCK` `\\e` switch.
    ///
    /// ECMA-376 permits repeated `\\e` switches; values are retained in source
    /// order. They are never matched against a recipient record.
    pub fn excluded_countries(&self) -> &[String] {
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
    pub fn unknown_switches(&self) -> &[Switch] {
        &self.unknown_switches
    }

    /// Return the cached visible field result, if present.
    ///
    /// This is stored text only and is never regenerated by a merge.
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
}

/// A typed, inert Word `MACROBUTTON` field.
///
/// ECMA-376 Part 1 §17.16.5.34 defines two stored field arguments: a macro or
/// command name and the text or graphic used as its button. This type exposes
/// stored text only; it never resolves, loads, invokes, or otherwise executes
/// the named target.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MacroButton {
    instruction: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
    macro_name: String,
    display_text: String,
}

impl MacroButton {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some((macro_name, display_text)) = parse_macro_button_operands(field.instruction())?
        else {
            return Ok(None);
        };

        Ok(Some(Self {
            instruction: field.instruction.clone(),
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
            macro_name,
            display_text,
        }))
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored macro or command name without resolving or invoking it.
    pub fn macro_name(&self) -> &str {
        &self.macro_name
    }

    /// Return the stored button text.
    ///
    /// This is source metadata, not a generated result.
    pub fn display_text(&self) -> &str {
        &self.display_text
    }

    /// Return the cached visible field result, if present.
    ///
    /// This is stored text only and is never regenerated from the named target.
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
}

/// A typed, inert Word `GOTOBUTTON` field.
///
/// ECMA-376 Part 1 §17.16.5.23 defines two stored field arguments: a
/// destination and the text or graphic used as its button. This type exposes
/// stored text only; it never resolves a destination, changes the insertion
/// point, or activates a jump.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct GoToButton {
    instruction: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
    target: String,
    button_text: String,
}

impl GoToButton {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some((target, button_text)) = parse_go_to_button_operands(field.instruction())? else {
            return Ok(None);
        };

        Ok(Some(Self {
            instruction: field.instruction.clone(),
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
            target,
            button_text,
        }))
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
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

    /// Return the cached visible field result, if present.
    ///
    /// This is stored text only and is never regenerated from the destination.
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
}

/// The stored kind of an active-content field.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ActiveContentKind {
    /// An `ADDIN` field that stores add-in-created data.
    AddIn,
    /// A `CONTROL` field that represents an OCX control.
    OcxControl,
    /// An `HTMLCONTROL` field that represents an HTML control.
    HtmlControl,
}

/// A typed, inert Word add-in or control field.
///
/// ECMA-376 Part 1 §17.16.5 defines `ADDIN`, `CONTROL`, and
/// `HTMLCONTROL` field instructions. This type retains only the stored
/// category, instruction, cached result, and state. It never loads an add-in,
/// instantiates an OCX or HTML control, invokes code, executes script, renders
/// content, or accesses an external resource.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ActiveContent {
    instruction: String,
    kind: ActiveContentKind,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
}

impl ActiveContent {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let kind = if field_instruction_remainder(field.instruction(), "ADDIN").is_some() {
            ActiveContentKind::AddIn
        } else if field_instruction_remainder(field.instruction(), "CONTROL").is_some() {
            ActiveContentKind::OcxControl
        } else if field_instruction_remainder(field.instruction(), "HTMLCONTROL").is_some() {
            ActiveContentKind::HtmlControl
        } else {
            return Ok(None);
        };

        Ok(Some(Self {
            instruction: field.instruction.clone(),
            kind,
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
        }))
    }

    /// Return the complete stored field instruction.
    ///
    /// This string remains opaque metadata and is never interpreted.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return whether this stores add-in, OCX-control, or HTML-control metadata.
    pub fn kind(&self) -> ActiveContentKind {
        self.kind
    }

    /// Return the cached visible field result, if present.
    ///
    /// This is stored text only and is never regenerated by loading or running
    /// content.
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
}

/// A typed, inert Word `PRINT` field.
///
/// Microsoft Word uses this field to store printer-control instructions. This
/// type retains the opaque instruction text, cached result, and field state
/// only. It never interprets printer-control codes, opens a printer, sends
/// output, changes print settings, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Print {
    instruction: String,
    printer_instructions: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
}

impl Print {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some(printer_instructions) = field_instruction_remainder(field.instruction(), "PRINT")
        else {
            return Ok(None);
        };

        Ok(Some(Self {
            instruction: field.instruction.clone(),
            printer_instructions: printer_instructions.trim().to_string(),
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
        }))
    }

    /// Return the complete stored field instruction.
    ///
    /// This string remains opaque metadata and is never interpreted or sent to
    /// a printer.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored printer-instruction text after the `PRINT` keyword.
    ///
    /// This can include printer-control or PostScript text. It is never parsed,
    /// interpreted, or sent to a printer.
    pub fn printer_instructions(&self) -> &str {
        &self.printer_instructions
    }

    /// Return the cached visible field result, if present.
    ///
    /// This is stored text only and is never regenerated by printing.
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
}

/// A typed, inert Word `EMBED` field.
///
/// Word uses this field to retain an embedded OLE object's stored metadata.
/// This type retains opaque object-instruction text, cached result, and field
/// state only. It never loads, inspects, deserializes, activates, renders, or
/// executes an embedded object, accesses an external resource, or refreshes a
/// field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Embed {
    instruction: String,
    object_instructions: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
}

impl Embed {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some(object_instructions) = field_instruction_remainder(field.instruction(), "EMBED")
        else {
            return Ok(None);
        };
        if field.instruction().len() > MAX_EMBED_FIELD_INSTRUCTION_BYTES {
            return Err(Error::Invalid(format!(
                "EMBED field instruction exceeds {MAX_EMBED_FIELD_INSTRUCTION_BYTES} bytes"
            )));
        }

        Ok(Some(Self {
            instruction: field.instruction.clone(),
            object_instructions: object_instructions.trim().to_string(),
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
        }))
    }

    /// Return the complete stored field instruction.
    ///
    /// This string remains opaque metadata and is never used to load or
    /// activate an object.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored opaque object-instruction text after `EMBED`.
    ///
    /// It is never parsed, used to locate an object, or used to load, inspect,
    /// activate, render, or execute object content.
    pub fn object_instructions(&self) -> &str {
        &self.object_instructions
    }

    /// Return the cached visible field result, if present.
    ///
    /// This is stored text only and is never regenerated from an object.
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
}

/// A typed, inert Word `BARCODE` field.
///
/// This type retains opaque barcode-instruction text, a cached result, and
/// field state only. It never parses or validates barcode data or symbology,
/// generates or renders a barcode, accesses an external resource, or refreshes
/// a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Barcode {
    instruction: String,
    barcode_instructions: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
}

impl Barcode {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some(barcode_instructions) =
            field_instruction_remainder(field.instruction(), "BARCODE")
        else {
            return Ok(None);
        };
        if field.instruction().len() > MAX_BARCODE_FIELD_INSTRUCTION_BYTES {
            return Err(Error::Invalid(format!(
                "BARCODE field instruction exceeds {MAX_BARCODE_FIELD_INSTRUCTION_BYTES} bytes"
            )));
        }

        Ok(Some(Self {
            instruction: field.instruction.clone(),
            barcode_instructions: barcode_instructions.trim().to_string(),
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
        }))
    }

    /// Return the complete stored field instruction.
    ///
    /// This string remains opaque metadata and is never used to generate or
    /// render a barcode.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored opaque barcode-instruction text after `BARCODE`.
    ///
    /// It is never parsed, validated, interpreted, or used to generate or
    /// render barcode content.
    pub fn barcode_instructions(&self) -> &str {
        &self.barcode_instructions
    }

    /// Return the cached visible field result, if present.
    ///
    /// This is stored text only and is never regenerated from barcode data.
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
}

/// A typed, inert Word `BIDIOUTLINE` field.
///
/// This type retains opaque instruction text, a cached result, and field state
/// only. It never reads right-to-left language, paragraph outline, or layout
/// state; chooses a numbering system; calculates a result; or refreshes a
/// field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct BidiOutline {
    instruction: String,
    opaque_instructions: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
}

impl BidiOutline {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some(opaque_instructions) =
            field_instruction_remainder(field.instruction(), "BIDIOUTLINE")
        else {
            return Ok(None);
        };
        if field.instruction().len() > MAX_BIDI_OUTLINE_FIELD_INSTRUCTION_BYTES {
            return Err(Error::Invalid(format!(
                "BIDIOUTLINE field instruction exceeds {MAX_BIDI_OUTLINE_FIELD_INSTRUCTION_BYTES} bytes"
            )));
        }

        Ok(Some(Self {
            instruction: field.instruction.clone(),
            opaque_instructions: opaque_instructions.trim().to_string(),
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
        }))
    }

    /// Return the complete stored field instruction.
    ///
    /// This string remains opaque metadata and is never used to calculate an
    /// outline number.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return opaque stored instruction text after `BIDIOUTLINE`.
    ///
    /// It is never parsed, interpreted, or used to resolve language, outline,
    /// numbering, or layout state.
    pub fn opaque_instructions(&self) -> &str {
        &self.opaque_instructions
    }

    /// Return the cached visible field result, if present.
    ///
    /// This is stored text only and is never regenerated from document state.
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
}

/// A typed, inert Word `SHAPE` drawing-canvas anchor field.
///
/// Word uses this legacy field as a drawing-canvas anchor. This type retains
/// opaque instruction text, a cached result, and field state only. It never
/// locates, links, loads, positions, lays out, or renders a drawing or canvas,
/// or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Shape {
    instruction: String,
    opaque_instructions: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
}

impl Shape {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some(opaque_instructions) = field_instruction_remainder(field.instruction(), "SHAPE")
        else {
            return Ok(None);
        };
        if field.instruction().len() > MAX_SHAPE_FIELD_INSTRUCTION_BYTES {
            return Err(Error::Invalid(format!(
                "SHAPE field instruction exceeds {MAX_SHAPE_FIELD_INSTRUCTION_BYTES} bytes"
            )));
        }

        Ok(Some(Self {
            instruction: field.instruction.clone(),
            opaque_instructions: opaque_instructions.trim().to_string(),
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
        }))
    }

    /// Return the complete stored field instruction.
    ///
    /// This string remains opaque metadata and is never used to locate or
    /// position a drawing canvas.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return opaque stored instruction text after `SHAPE`.
    ///
    /// It is never parsed, interpreted, or used to link a field to a drawing,
    /// resolve an anchor, or calculate layout.
    pub fn opaque_instructions(&self) -> &str {
        &self.opaque_instructions
    }

    /// Return the cached visible field result, if present.
    ///
    /// This is stored metadata only and is never regenerated from a drawing
    /// canvas.
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
}

/// The stored kind of a legacy Word form-code field.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LegacyFormKind {
    /// A `FORMTEXT` text-box form field.
    Text,
    /// A `FORMCHECKBOX` checkbox form field.
    CheckBox,
    /// A `FORMDROPDOWN` drop-down-list form field.
    DropDown,
}

impl LegacyFormKind {
    fn from_instruction(instruction: &str) -> Option<(Self, &str)> {
        for (kind, keyword) in [
            (Self::Text, "FORMTEXT"),
            (Self::CheckBox, "FORMCHECKBOX"),
            (Self::DropDown, "FORMDROPDOWN"),
        ] {
            if let Some(remainder) = field_instruction_remainder(instruction, keyword) {
                return Some((kind, remainder));
            }
        }
        None
    }
}

/// A typed, inert Word legacy form-code field.
///
/// This type retains only the stored text/checkbox/drop-down kind, opaque
/// instruction text, cached result, and field state. It never reads associated
/// form-property XML, fills a form, changes a selection or checkbox state,
/// invokes entry or exit macros, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct LegacyForm {
    instruction: String,
    kind: LegacyFormKind,
    opaque_instructions: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
}

impl LegacyForm {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some((kind, opaque_instructions)) =
            LegacyFormKind::from_instruction(field.instruction())
        else {
            return Ok(None);
        };
        if field.instruction().len() > MAX_LEGACY_FORM_FIELD_INSTRUCTION_BYTES {
            return Err(Error::Invalid(format!(
                "legacy form-code field instruction exceeds {MAX_LEGACY_FORM_FIELD_INSTRUCTION_BYTES} bytes"
            )));
        }

        Ok(Some(Self {
            instruction: field.instruction.clone(),
            kind,
            opaque_instructions: opaque_instructions.trim().to_string(),
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
        }))
    }

    /// Return the complete stored field instruction.
    ///
    /// This string remains opaque metadata and is never used to change a form
    /// field.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return whether this is a text, checkbox, or drop-down form-code field.
    pub const fn kind(&self) -> LegacyFormKind {
        self.kind
    }

    /// Return opaque stored instruction text after the form-code keyword.
    ///
    /// It is never parsed, interpreted, or used to fill a form, change a
    /// checkbox or selection, or invoke a macro.
    pub fn opaque_instructions(&self) -> &str {
        &self.opaque_instructions
    }

    /// Return the cached visible field result, if present.
    ///
    /// This is stored metadata only and is never regenerated from form state.
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
}

/// A typed, inert Word `PRIVATE` conversion-data field.
///
/// Word uses this field to preserve data needed to convert a document back to
/// another file format. This type retains opaque instruction text, a cached
/// result, and field state only. It never converts a document, interprets field
/// data, changes hidden-text visibility or layout, or refreshes a field.
/// Despite its name, `PRIVATE` does not provide confidentiality semantics.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Private {
    instruction: String,
    opaque_instructions: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
}

impl Private {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some(opaque_instructions) = field_instruction_remainder(field.instruction(), "PRIVATE")
        else {
            return Ok(None);
        };
        if field.instruction().len() > MAX_PRIVATE_FIELD_INSTRUCTION_BYTES {
            return Err(Error::Invalid(format!(
                "PRIVATE field instruction exceeds {MAX_PRIVATE_FIELD_INSTRUCTION_BYTES} bytes"
            )));
        }

        Ok(Some(Self {
            instruction: field.instruction.clone(),
            opaque_instructions: opaque_instructions.trim().to_string(),
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
        }))
    }

    /// Return the complete stored field instruction.
    ///
    /// This string remains opaque metadata and is never used to convert a
    /// document or change hidden-text visibility.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return opaque stored instruction text after `PRIVATE`.
    ///
    /// It is never parsed, interpreted, or used to convert a document, change
    /// hidden-text visibility, or calculate layout.
    pub fn opaque_instructions(&self) -> &str {
        &self.opaque_instructions
    }

    /// Return the cached stored field result, if present.
    ///
    /// This is stored metadata only and is never regenerated by conversion or
    /// used to change hidden-text visibility.
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
}

/// A typed, inert Word `DATABASE` query field.
///
/// Word uses this field to query a database and insert a table. This type
/// retains opaque instruction text, a cached result, and field state only. It
/// never opens a data source or database, uses connection information, executes
/// SQL, generates or inserts a table, changes layout, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Database {
    instruction: String,
    opaque_instructions: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
}

impl Database {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some(opaque_instructions) =
            field_instruction_remainder(field.instruction(), "DATABASE")
        else {
            return Ok(None);
        };
        if field.instruction().len() > MAX_DATABASE_FIELD_INSTRUCTION_BYTES {
            return Err(Error::Invalid(format!(
                "DATABASE field instruction exceeds {MAX_DATABASE_FIELD_INSTRUCTION_BYTES} bytes"
            )));
        }

        Ok(Some(Self {
            instruction: field.instruction.clone(),
            opaque_instructions: opaque_instructions.trim().to_string(),
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
        }))
    }

    /// Return the complete stored field instruction.
    ///
    /// This string remains opaque metadata and is never used to open a data
    /// source, database, or connection.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return opaque stored instruction text after `DATABASE`.
    ///
    /// It is never parsed, interpreted, or used to connect, execute SQL,
    /// generate a table, or calculate layout.
    pub fn opaque_instructions(&self) -> &str {
        &self.opaque_instructions
    }

    /// Return the cached stored field result, if present.
    ///
    /// This is stored metadata only and is never regenerated from a database
    /// query.
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
}

/// The stored kind of a Word building-block field.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AutoTextKind {
    /// A historical `GLOSSARY` field.
    Glossary,
    /// An `AUTOTEXT` field.
    AutoText,
}

/// A typed, inert Word `GLOSSARY` or `AUTOTEXT` field.
///
/// ECMA-376 Part 1 §17.16.5.5 defines `AUTOTEXT`; `GLOSSARY` is its
/// historical equivalent. This type retains only the stored category, entry
/// name, switches, cached result, and state. It never looks up a building
/// block, reads a template, inserts content, changes bookmarks, refreshes a
/// field, or accesses an external resource.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AutoText {
    instruction: String,
    kind: AutoTextKind,
    entry_name: String,
    unknown_switches: Vec<Switch>,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
}

impl AutoText {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some((kind, entry_name, unknown_switches)) =
            parse_auto_text_field_parts(field.instruction())?
        else {
            return Ok(None);
        };

        Ok(Some(Self {
            instruction: field.instruction.clone(),
            kind,
            entry_name,
            unknown_switches,
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
        }))
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return whether this stores `GLOSSARY` or `AUTOTEXT` metadata.
    pub fn kind(&self) -> AutoTextKind {
        self.kind
    }

    /// Return the stored building-block entry name without resolving it.
    pub fn entry_name(&self) -> &str {
        &self.entry_name
    }

    /// Return unrecognized stored switches in source order.
    ///
    /// They are retained without interpretation or execution.
    pub fn unknown_switches(&self) -> &[Switch] {
        &self.unknown_switches
    }

    /// Return the cached visible field result, if present.
    ///
    /// This is stored text only and is never regenerated by looking up or
    /// inserting content.
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
}

/// One recognized stored option of a Word `AUTOTEXTLIST` field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum AutoTextListOption {
    /// The `\\s` style name used to limit eligible building blocks.
    Style(String),
    /// The `\\t` stored tip text.
    Tip(String),
}

/// A typed, inert Word `AUTOTEXTLIST` field.
///
/// ECMA-376 Part 1 §17.16.5.6 defines `AUTOTEXTLIST` using optional
/// display text and style/tip switches. This type retains only those stored
/// values, unknown switches, cached result, and state. It never shows a
/// selection UI, looks up eligible building blocks, reads a template, inserts
/// content, or refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AutoTextList {
    instruction: String,
    display_text: Option<String>,
    options: Vec<AutoTextListOption>,
    unknown_switches: Vec<Switch>,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
}

impl AutoTextList {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some((display_text, options, unknown_switches)) =
            parse_auto_text_list_field_parts(field.instruction())?
        else {
            return Ok(None);
        };

        Ok(Some(Self {
            instruction: field.instruction.clone(),
            display_text,
            options,
            unknown_switches,
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
        }))
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
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
    pub fn options(&self) -> &[AutoTextListOption] {
        &self.options
    }

    /// Return unrecognized stored switches in source order.
    ///
    /// They are retained without interpretation or execution.
    pub fn unknown_switches(&self) -> &[Switch] {
        &self.unknown_switches
    }

    /// Return the cached visible field result, if present.
    ///
    /// This is stored text only and is never regenerated by selection or
    /// content insertion.
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
}

/// The stored kind of a user-identity field.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum UserIdentityKind {
    /// A `USERADDRESS` field.
    Address,
    /// A `USERINITIALS` field.
    Initials,
    /// A `USERNAME` field.
    Name,
}

/// A general-formatting request stored by a user-identity field.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum UserIdentityFormat {
    /// The `\\* Caps` formatting request.
    Caps,
    /// The `\\* FirstCap` formatting request.
    FirstCap,
    /// The `\\* Lower` formatting request.
    Lower,
    /// The `\\* Upper` formatting request.
    Upper,
}

/// A typed, inert Word `USERADDRESS`, `USERINITIALS`, or `USERNAME` field.
///
/// ECMA-376 Part 1 §§17.16.5.69–71 define these fields. This type exposes a
/// stored override, formatting request, and cached result only. It never reads
/// or modifies a host user's identity, applies formatting, or refreshes a
/// field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct UserIdentity {
    instruction: String,
    kind: UserIdentityKind,
    override_value: Option<String>,
    formatting: Option<UserIdentityFormat>,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
}

impl UserIdentity {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some((kind, override_value, formatting)) =
            parse_user_identity_field_parts(field.instruction())?
        else {
            return Ok(None);
        };

        Ok(Some(Self {
            instruction: field.instruction.clone(),
            kind,
            override_value,
            formatting,
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
        }))
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return whether this is an address, initials, or name field.
    pub fn kind(&self) -> UserIdentityKind {
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
    pub fn formatting(&self) -> Option<UserIdentityFormat> {
        self.formatting
    }

    /// Return the cached visible field result, if present.
    ///
    /// This is stored text only and is never regenerated from a host identity.
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
}

/// One stored point-based `ADVANCE` placement operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AdvanceOperation {
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
pub struct AdvanceAdjustment {
    pub(super) operation: AdvanceOperation,
    pub(super) points: i64,
}

impl AdvanceAdjustment {
    /// Return the requested placement operation.
    pub fn operation(&self) -> AdvanceOperation {
        self.operation
    }

    /// Return the stored signed integral number of points.
    pub fn points(&self) -> i64 {
        self.points
    }
}

/// A typed, inert Word `ADVANCE` field.
///
/// ECMA-376 Part 1 §17.16.5.2 defines this field and its six point-based
/// placement switches. This type exposes stored adjustments and cached content
/// only. It never moves text, changes layout, reflows content, or refreshes a
/// field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Advance {
    instruction: String,
    adjustments: Vec<AdvanceAdjustment>,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
}

impl Advance {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some(adjustments) = parse_advance_field_adjustments(field.instruction())? else {
            return Ok(None);
        };

        Ok(Some(Self {
            instruction: field.instruction.clone(),
            adjustments,
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
        }))
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored placement adjustments in source order.
    ///
    /// Repeated operations are preserved; this library does not resolve or
    /// apply them.
    pub fn adjustments(&self) -> &[AdvanceAdjustment] {
        &self.adjustments
    }

    /// Return the cached visible field result, if present.
    ///
    /// `ADVANCE` has no regenerated value here; any returned text is stored
    /// source content only.
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
}

/// Backward-compatible name for a lexical switch exposed by a TOC field.
pub type TocSwitch = Switch;

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
