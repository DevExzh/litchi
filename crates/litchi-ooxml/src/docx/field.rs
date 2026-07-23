/// Field support for reading fields from Word documents.
///
/// This module provides types and methods for accessing fields in Word documents.
/// Fields are dynamic content like page numbers, dates, formulas, and cross-references.
use crate::common::xml::decode_xml_reference;
use crate::error::{OoxmlError, Result};
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::{Reader, XmlVersion};

const MAX_FIELD_SWITCHES: usize = 64;
const MAX_FORMULA_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
const MAX_QUOTE_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
const MAX_SYMBOL_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
const MAX_SET_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
const MAX_SEQUENCE_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
const MAX_STYLE_REFERENCE_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
const MAX_AUTO_TEXT_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
const MAX_AUTO_TEXT_LIST_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
const MAX_DOCUMENT_PROPERTY_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
const MAX_DOCUMENT_INFORMATION_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
const MAX_DOCUMENT_CONTEXT_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;

/// A field in a Word document.
///
/// Represents a field instruction like `PAGE`, `DATE`, `REF`, etc.
/// Fields are dynamic content that can be updated.
///
/// # Examples
///
/// ```rust,no_run
/// use litchi_ooxml::docx::Package;
///
/// let pkg = Package::open("document.docx")?;
/// let doc = pkg.document()?;
///
/// for field in doc.fields()? {
///     println!("Field: {} = {}", field.instruction(), field.result().unwrap_or(""));
/// }
/// # Ok::<(), Box<dyn std::error::Error>>(())
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

    fn with_flags(instruction: String, result: Option<String>, dirty: bool, locked: bool) -> Self {
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
    /// use litchi_ooxml::docx::Field;
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
    pub fn merge_field(&self) -> Result<Option<MergeField>> {
        MergeField::from_field(self)
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
    pub fn mail_merge_counter(&self) -> Result<Option<MailMergeCounterField>> {
        MailMergeCounterField::from_field(self)
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
    pub fn mail_merge_next(&self) -> Result<Option<MailMergeNextField>> {
        MailMergeNextField::from_field(self)
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
    pub fn mail_merge_conditional_control(
        &self,
    ) -> Result<Option<MailMergeConditionalControlField>> {
        MailMergeConditionalControlField::from_field(self)
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
    pub fn if_field(&self) -> Result<Option<IfField>> {
        IfField::from_field(self)
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
    pub fn compare_field(&self) -> Result<Option<CompareField>> {
        CompareField::from_field(self)
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
    pub fn set_field(&self) -> Result<Option<SetField>> {
        SetField::from_field(self)
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
    pub fn sequence_field(&self) -> Result<Option<SequenceField>> {
        SequenceField::from_field(self)
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
    pub fn formula_field(&self) -> Result<Option<FormulaField>> {
        FormulaField::from_field(self)
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
    pub fn quote_field(&self) -> Result<Option<QuoteField>> {
        QuoteField::from_field(self)
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
    pub fn symbol_field(&self) -> Result<Option<SymbolField>> {
        SymbolField::from_field(self)
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
    pub fn style_reference_field(&self) -> Result<Option<StyleReferenceField>> {
        StyleReferenceField::from_field(self)
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
    pub fn prompt_field(&self) -> Result<Option<PromptField>> {
        PromptField::from_field(self)
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
    pub fn mail_merge_recipient_field(&self) -> Result<Option<MailMergeRecipientField>> {
        MailMergeRecipientField::from_field(self)
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
    pub fn citation(&self) -> Result<Option<CitationField>> {
        CitationField::from_field(self)
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
    pub fn bibliography(&self) -> Result<Option<BibliographyField>> {
        BibliographyField::from_field(self)
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
    pub fn document_variable(&self) -> Result<Option<DocumentVariableField>> {
        DocumentVariableField::from_field(self)
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
    pub fn document_property(&self) -> Result<Option<DocumentPropertyField>> {
        DocumentPropertyField::from_field(self)
    }

    /// Check whether this is a built-in document-information field.
    ///
    /// Recognition is limited to the stored field instruction. It never reads
    /// package metadata or host identity data, resolves a value, or refreshes
    /// the field.
    pub fn is_document_information(&self) -> bool {
        DocumentInformationFieldKind::from_instruction(&self.instruction).is_some()
    }

    /// Parse this field as inert typed document-information metadata.
    ///
    /// Returns `Ok(None)` for fields outside the built-in document-information
    /// family. The result exposes only the stored kind, switches, cached
    /// content, and dirty/lock state; it never reads core or extended package
    /// properties, reads or modifies host identity data, calculates dates,
    /// revisions, or statistics, resolves a value, or refreshes a field.
    pub fn document_information(&self) -> Result<Option<DocumentInformationField>> {
        DocumentInformationField::from_field(self)
    }

    /// Check whether this is a built-in document-context or runtime field.
    ///
    /// Recognition is limited to the stored field instruction. It never reads
    /// a document path, attached template, host filesystem state or file size,
    /// current clock, or page and section layout, resolves a value, or refreshes
    /// the field.
    pub fn is_document_context(&self) -> bool {
        DocumentContextFieldKind::from_instruction(&self.instruction).is_some()
    }

    /// Parse this field as inert typed document-context or runtime metadata.
    ///
    /// Returns `Ok(None)` for fields outside the `FILENAME`, `TEMPLATE`, `DATE`,
    /// `TIME`, `PAGE`, `FILESIZE`, `SECTION`, and `SECTIONPAGES` family. The
    /// result exposes only the stored kind, switches, cached content, and
    /// dirty/lock state; it never reads a document path, attached template,
    /// host filesystem state or file size, current clock, or page and section
    /// layout, resolves a value, or refreshes a field.
    pub fn document_context(&self) -> Result<Option<DocumentContextField>> {
        DocumentContextField::from_field(self)
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
    pub fn macro_button(&self) -> Result<Option<MacroButtonField>> {
        MacroButtonField::from_field(self)
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
    pub fn go_to_button(&self) -> Result<Option<GoToButtonField>> {
        GoToButtonField::from_field(self)
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
    pub fn active_content_field(&self) -> Result<Option<ActiveContentField>> {
        ActiveContentField::from_field(self)
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
    pub fn auto_text_field(&self) -> Result<Option<AutoTextField>> {
        AutoTextField::from_field(self)
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
    pub fn auto_text_list_field(&self) -> Result<Option<AutoTextListField>> {
        AutoTextListField::from_field(self)
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
    pub fn user_identity_field(&self) -> Result<Option<UserIdentityField>> {
        UserIdentityField::from_field(self)
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
    pub fn advance_field(&self) -> Result<Option<AdvanceField>> {
        AdvanceField::from_field(self)
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
    pub fn link(&self) -> Result<Option<LinkField>> {
        LinkField::from_field(self)
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
    pub fn dde_link(&self) -> Result<Option<DdeField>> {
        DdeField::from_field(self)
    }

    /// Check whether this is an INCLUDETEXT field.
    ///
    /// Recognition is limited to the stored field instruction. It never opens,
    /// resolves, imports, fetches, or refreshes the referenced source.
    pub fn is_include_text(&self) -> bool {
        field_instruction_remainder(&self.instruction, "INCLUDETEXT").is_some()
    }

    /// Check whether this is an INCLUDEPICTURE field.
    ///
    /// Recognition is limited to the stored field instruction. It never opens,
    /// resolves, imports, fetches, or refreshes the referenced source.
    pub fn is_include_picture(&self) -> bool {
        field_instruction_remainder(&self.instruction, "INCLUDEPICTURE").is_some()
    }

    /// Parse this field as inert external-include metadata.
    ///
    /// Returns Ok(None) for fields other than INCLUDETEXT or INCLUDEPICTURE.
    /// The result exposes stored source, bookmark, converter, XML, and cached
    /// metadata only; it never opens, resolves, imports, fetches, refreshes,
    /// converts, evaluates, or executes anything.
    pub fn external_include(&self) -> Result<Option<ExternalIncludeField>> {
        ExternalIncludeField::from_field(self)
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
    pub fn referenced_document(&self) -> Result<Option<ReferencedDocumentField>> {
        ReferencedDocumentField::from_field(self)
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
    pub fn table_of_contents(&self) -> Result<Option<TableOfContentsField>> {
        TableOfContentsField::from_field(self)
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
    pub fn table_of_authorities(&self) -> Result<Option<TableOfAuthoritiesField>> {
        TableOfAuthoritiesField::from_field(self)
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
    pub fn table_of_authorities_entry(&self) -> Result<Option<TableOfAuthoritiesEntryField>> {
        TableOfAuthoritiesEntryField::from_field(self)
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
    pub fn index(&self) -> Result<Option<IndexField>> {
        IndexField::from_field(self)
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
    pub fn index_entry(&self) -> Result<Option<IndexEntryField>> {
        IndexEntryField::from_field(self)
    }

    /// Extract all fields from document XML bytes.
    ///
    /// # Arguments
    ///
    /// * `doc_xml` - The document XML bytes
    ///
    /// # Returns
    ///
    /// A vector of fields
    pub(crate) fn extract_from_document(doc_xml: &[u8]) -> Result<Vec<Field>> {
        let mut reader = Reader::from_reader(doc_xml);
        reader.config_mut().trim_text(false);

        let mut fields = Vec::new();
        let mut next_order = 0usize;
        let mut in_instr_text = false;
        let mut in_field_result = false;
        let mut in_result_text = false;
        let mut in_simple_result_text = false;
        let mut current_instruction = String::new();
        let mut current_result = String::new();
        let mut current_dirty = false;
        let mut current_locked = false;
        let mut current_order = 0usize;
        let mut field_depth: i32 = 0;
        let mut simple_fields = Vec::new();

        loop {
            match reader.read_event() {
                Ok(Event::Empty(e)) if e.local_name().as_ref() == b"t" => {
                    in_result_text = false;
                    in_simple_result_text = false;
                },
                Ok(Event::Start(e)) if e.local_name().as_ref() == b"fldSimple" => {
                    simple_fields.push(PendingSimpleField::parse(
                        &e,
                        reader.decoder(),
                        next_order,
                    )?);
                    next_order += 1;
                    in_simple_result_text = false;
                },
                Ok(Event::Empty(e)) if e.local_name().as_ref() == b"fldSimple" => {
                    let field = PendingSimpleField::parse(&e, reader.decoder(), next_order)?;
                    next_order += 1;
                    fields.push((field.order, field.finish()));
                },
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    match e.local_name().as_ref() {
                        b"fldChar" => {
                            // Field character marks field boundaries
                            let mut fld_char_type = None;
                            let mut dirty = None;
                            let mut locked = None;

                            for attr in e.attributes() {
                                let attr = attr
                                    .map_err(|error| OoxmlError::Xml(error.to_string()))?;
                                let value = attr
                                    .decoded_and_normalized_value(
                                        XmlVersion::Explicit1_0,
                                        reader.decoder(),
                                    )
                                    .map_err(|error| OoxmlError::Xml(error.to_string()))?;
                                if attr.key.local_name().as_ref() == b"fldCharType" {
                                    fld_char_type = Some(value.to_string());
                                }
                                if attr.key.local_name().as_ref() == b"dirty" {
                                    dirty = Some(is_on(&value));
                                }
                                if attr.key.local_name().as_ref() == b"fldLock" {
                                    locked = Some(is_on(&value));
                                }
                            }

                            if let Some(ref char_type) = fld_char_type {
                                match char_type.as_str() {
                                    "begin" => {
                                        // Start of field
                                        field_depth += 1;
                                        if field_depth == 1 {
                                            current_order = next_order;
                                            next_order += 1;
                                            current_instruction.clear();
                                            current_result.clear();
                                            current_dirty = dirty.unwrap_or(false);
                                            current_locked = locked.unwrap_or(false);
                                            in_instr_text = false;
                                            in_field_result = false;
                                            in_result_text = false;
                                        }
                                    },
                                    "separate"
                                        // Separator between instruction and result
                                        if field_depth == 1 => {
                                            current_dirty |= dirty.unwrap_or(false);
                                            current_locked |= locked.unwrap_or(false);
                                            in_instr_text = false;
                                            in_field_result = true;
                                            in_result_text = false;
                                        },
                                    "end" => {
                                        // End of field
                                        if field_depth == 1 {
                                            in_field_result = false;
                                            in_instr_text = false;
                                            in_result_text = false;

                                            if !current_instruction.is_empty() {
                                                let result = if current_result.is_empty() {
                                                    None
                                                } else {
                                                    Some(current_result.clone())
                                                };
                                                fields.push((current_order, Field::with_flags(
                                                    current_instruction.trim().to_string(),
                                                    result,
                                                    current_dirty,
                                                    current_locked,
                                                )));
                                            }
                                        }
                                        field_depth = field_depth.saturating_sub(1);
                                    },
                                    _ => {},
                                }
                            }
                        },
                        b"instrText"
                            // Field instruction text
                            if field_depth > 0 => {
                                in_instr_text = true;
                            },
                        b"t" => {
                            if in_field_result {
                                in_result_text = true;
                            }
                            if !simple_fields.is_empty() {
                                in_simple_result_text = true;
                            }
                        },
                        b"tab" if in_field_result && field_depth == 1 => {
                            current_result.push('\t');
                        },
                        b"br" | b"cr" if in_field_result && field_depth == 1 => {
                            current_result.push('\n');
                        },
                        b"noBreakHyphen" if in_field_result && field_depth == 1 => {
                            current_result.push('\u{2011}');
                        },
                        b"softHyphen" if in_field_result && field_depth == 1 => {
                            current_result.push('\u{00ad}');
                        },
                        _ => {},
                    }

                    if !simple_fields.is_empty() {
                        let character = match e.local_name().as_ref() {
                            b"tab" => Some('\t'),
                            b"br" | b"cr" => Some('\n'),
                            b"noBreakHyphen" => Some('\u{2011}'),
                            b"softHyphen" => Some('\u{00ad}'),
                            _ => None,
                        };
                        if let Some(character) = character {
                            for field in &mut simple_fields {
                                field.result.push(character);
                            }
                        }
                    }
                },
                Ok(Event::Text(e)) => {
                    let has_complex_target = (in_instr_text && field_depth == 1)
                        || (in_field_result && in_result_text && field_depth == 1);
                    if has_complex_target || in_simple_result_text {
                        let decoded = e
                            .xml_content(XmlVersion::Explicit1_0)
                            .map_err(|error| OoxmlError::Xml(error.to_string()))?;
                        let unescaped = quick_xml::escape::unescape(&decoded)
                            .map_err(|error| OoxmlError::Xml(error.to_string()))?;
                        if in_instr_text && field_depth == 1 {
                            current_instruction.push_str(&unescaped);
                        } else if in_field_result && in_result_text && field_depth == 1 {
                            current_result.push_str(&unescaped);
                        }
                        if in_simple_result_text {
                            for field in &mut simple_fields {
                                field.result.push_str(&unescaped);
                            }
                        }
                    }
                },
                Ok(Event::CData(e)) => {
                    let has_complex_target = (in_instr_text && field_depth == 1)
                        || (in_field_result && in_result_text && field_depth == 1);
                    if has_complex_target || in_simple_result_text {
                        let decoded = e
                            .xml_content(XmlVersion::Explicit1_0)
                            .map_err(|error| OoxmlError::Xml(error.to_string()))?;
                        if in_instr_text && field_depth == 1 {
                            current_instruction.push_str(&decoded);
                        } else if in_field_result && in_result_text && field_depth == 1 {
                            current_result.push_str(&decoded);
                        }
                        if in_simple_result_text {
                            for field in &mut simple_fields {
                                field.result.push_str(&decoded);
                            }
                        }
                    }
                },
                Ok(Event::GeneralRef(reference)) => {
                    let has_complex_target = (in_instr_text && field_depth == 1)
                        || (in_field_result && in_result_text && field_depth == 1);
                    if has_complex_target || in_simple_result_text {
                        let decoded = decode_xml_reference(&reference)?;
                        if in_instr_text && field_depth == 1 {
                            current_instruction.push_str(&decoded);
                        } else if in_field_result && in_result_text && field_depth == 1 {
                            current_result.push_str(&decoded);
                        }
                        if in_simple_result_text {
                            for field in &mut simple_fields {
                                field.result.push_str(&decoded);
                            }
                        }
                    }
                },
                Ok(Event::End(e)) if e.local_name().as_ref() == b"instrText" => {
                    in_instr_text = false;
                },
                Ok(Event::End(e)) if e.local_name().as_ref() == b"t" => {
                    in_result_text = false;
                    in_simple_result_text = false;
                },
                Ok(Event::End(e)) if e.local_name().as_ref() == b"fldSimple" => {
                    in_simple_result_text = false;
                    let field = simple_fields.pop().ok_or_else(|| {
                        OoxmlError::InvalidFormat(
                            "DOCX simple field ended without a matching start".to_string(),
                        )
                    })?;
                    fields.push((field.order, field.finish()));
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
        }

        fields.sort_unstable_by_key(|(order, _)| *order);
        Ok(fields.into_iter().map(|(_, field)| field).collect())
    }
}

/// One lexical switch in a Word field instruction.
///
/// Switch names are normalized to ASCII lowercase. Quoted and unquoted
/// arguments are decoded into their logical text. Typed field models retain the
/// complete original instruction alongside these values.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FieldSwitch {
    name: char,
    argument: Option<String>,
}

impl FieldSwitch {
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
pub enum DdeFieldKind {
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
pub enum DdeRepresentation {
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
pub struct DdeField {
    instruction: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
    kind: DdeFieldKind,
    application: String,
    source: String,
    item: Option<String>,
    automatic_updates: bool,
    representation: Option<DdeRepresentation>,
    omit_graphic_data: bool,
    switches: Vec<FieldSwitch>,
}

impl DdeField {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some((kind, application, source, item, switches)) =
            parse_dde_operands_and_switches(field.instruction())?
        else {
            return Ok(None);
        };

        let mut automatic_updates = kind == DdeFieldKind::DdeAuto;
        let mut saw_automatic_update = false;
        let mut representation = None;
        let mut omit_graphic_data = false;
        for switch in &switches {
            match switch.name {
                'a' if kind == DdeFieldKind::Dde => {
                    if saw_automatic_update || switch.argument.is_some() {
                        return Err(OoxmlError::InvalidFormat(
                            "DDE \\a switch cannot be repeated or take an argument".to_string(),
                        ));
                    }
                    automatic_updates = true;
                    saw_automatic_update = true;
                },
                'a' => {
                    return Err(OoxmlError::InvalidFormat(
                        "DDEAUTO field does not allow a \\a switch".to_string(),
                    ));
                },
                'd' => {
                    if representation.is_some() || omit_graphic_data || switch.argument.is_some() {
                        return Err(OoxmlError::InvalidFormat(
                            "DDE result and storage switches cannot be combined".to_string(),
                        ));
                    }
                    omit_graphic_data = true;
                },
                'b' | 'h' | 'p' | 'r' | 't' | 'u' => {
                    if representation.is_some() || omit_graphic_data || switch.argument.is_some() {
                        return Err(OoxmlError::InvalidFormat(
                            "DDE result and storage switches cannot be combined".to_string(),
                        ));
                    }
                    representation = Some(match switch.name {
                        'b' => DdeRepresentation::Bitmap,
                        'h' => DdeRepresentation::Html,
                        'p' => DdeRepresentation::Picture,
                        'r' => DdeRepresentation::RichText,
                        't' => DdeRepresentation::Text,
                        'u' => DdeRepresentation::UnicodeText,
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
    pub fn kind(&self) -> DdeFieldKind {
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
    pub fn representation(&self) -> Option<DdeRepresentation> {
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
    pub fn switches(&self) -> &[FieldSwitch] {
        &self.switches
    }
}

/// The kind of externally sourced Word field.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum IncludeFieldKind {
    /// An INCLUDETEXT field that stores a document or XML source.
    Text,
    /// An INCLUDEPICTURE field that stores an image source.
    Picture,
}

/// One recognized stored option of an external-include field.
///
/// These values are configuration metadata only. This API never opens,
/// resolves, imports, transforms, or evaluates the referenced source.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ExternalIncludeOption {
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

/// Typed, inert metadata for an INCLUDETEXT or INCLUDEPICTURE field.
///
/// Source identifiers, bookmarks, options, and cached results are retained as
/// stored field data. This type never opens, resolves, imports, fetches,
/// refreshes, transforms, converts, evaluates, or executes source content.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExternalIncludeField {
    instruction: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
    kind: IncludeFieldKind,
    source: String,
    bookmark: Option<String>,
    suppress_nested_field_updates: bool,
    omit_picture_data: bool,
    options: Vec<ExternalIncludeOption>,
    switches: Vec<FieldSwitch>,
}

impl ExternalIncludeField {
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
                (IncludeFieldKind::Text, '!') => {
                    if switch.argument.is_some() {
                        return Err(OoxmlError::InvalidFormat(
                            "INCLUDETEXT exclamation switch does not take an argument".to_string(),
                        ));
                    }
                    suppress_nested_field_updates = true;
                },
                (IncludeFieldKind::Picture, 'd') => {
                    if switch.argument.is_some() {
                        return Err(OoxmlError::InvalidFormat(
                            "INCLUDEPICTURE d switch does not take an argument".to_string(),
                        ));
                    }
                    omit_picture_data = true;
                },
                (_, 'c') => options.push(ExternalIncludeOption::Converter(
                    required_external_include_option_argument(switch, kind)?,
                )),
                (IncludeFieldKind::Text, 'e') => options.push(ExternalIncludeOption::Encoding(
                    required_external_include_option_argument(switch, kind)?,
                )),
                (IncludeFieldKind::Text, 'm') => options.push(ExternalIncludeOption::MimeType(
                    required_external_include_option_argument(switch, kind)?,
                )),
                (IncludeFieldKind::Text, 'n') => {
                    options.push(ExternalIncludeOption::NamespaceMapping(
                        required_external_include_option_argument(switch, kind)?,
                    ))
                },
                (IncludeFieldKind::Text, 't') => options.push(ExternalIncludeOption::Xslt(
                    required_external_include_option_argument(switch, kind)?,
                )),
                (IncludeFieldKind::Text, 'x') => options.push(ExternalIncludeOption::XPath(
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
    pub fn kind(&self) -> IncludeFieldKind {
        self.kind
    }

    /// Return the stored source identifier without opening or resolving it.
    pub fn source(&self) -> &str {
        &self.source
    }

    /// Return the optional stored bookmark selector for an INCLUDETEXT field.
    pub fn bookmark(&self) -> Option<&str> {
        self.bookmark.as_deref()
    }

    /// Whether the stored INCLUDETEXT instruction suppresses nested updates.
    ///
    /// This is metadata only. The API never performs an update.
    pub fn suppresses_nested_field_updates(&self) -> bool {
        self.suppress_nested_field_updates
    }

    /// Whether the stored INCLUDEPICTURE instruction omits picture data.
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
    pub fn options(&self) -> &[ExternalIncludeOption] {
        &self.options
    }

    /// Return all stored field switches in source order.
    pub fn switches(&self) -> &[FieldSwitch] {
        &self.switches
    }
}

/// Typed, inert metadata for an RD referenced-document field.
///
/// Source identifiers, relative-path settings, switches, and cached results
/// are retained as stored field data. This type never opens, resolves, reads,
/// imports, refreshes, evaluates, or executes the referenced document.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ReferencedDocumentField {
    instruction: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
    source: String,
    relative_path: bool,
    switches: Vec<FieldSwitch>,
}

impl ReferencedDocumentField {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some((source, switches)) = parse_field_operand_and_switches(field.instruction(), "RD")?
        else {
            return Ok(None);
        };
        let source = source.filter(|value| !value.is_empty()).ok_or_else(|| {
            OoxmlError::InvalidFormat(
                "RD field is missing its referenced document path".to_string(),
            )
        })?;

        let mut relative_path = false;
        for switch in &switches {
            if switch.name == 'p' {
                if switch.argument.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "RD \\\\p switch does not take an argument".to_string(),
                    ));
                }
                if relative_path {
                    return Err(OoxmlError::InvalidFormat(
                        "RD \\\\p switch cannot be repeated".to_string(),
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

    /// Whether the stored RD instruction requests a path relative to this document.
    ///
    /// This is metadata only. The API never resolves the path.
    pub fn uses_relative_path(&self) -> bool {
        self.relative_path
    }

    /// Return all stored field switches in source order.
    pub fn switches(&self) -> &[FieldSwitch] {
        &self.switches
    }
}

/// One stored result or storage switch for a Word `LINK` field.
///
/// These values describe a linked-object representation or whether graphic data
/// is stored. They never cause a source to be opened, contacted, converted, or
/// displayed.
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
/// ECMA-376 marks modes `1` and `3` unsupported. Those values, and values
/// outside its defined set, are retained as metadata without applying any
/// formatting.
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

/// Typed, inert metadata for a legacy Word `LINK` field.
///
/// Application type, source, item, and all result/formatting switches are
/// retained as stored field data. This type never activates an OLE server,
/// launches an application, opens a source, requests data, refreshes content,
/// converts content, or executes code.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct LinkField {
    instruction: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
    application_type: String,
    source: String,
    item: Option<String>,
    automatic_updates: bool,
    result_options: Vec<LinkResultOption>,
    formatting_modes: Vec<LinkFormatting>,
    switches: Vec<FieldSwitch>,
}

impl LinkField {
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
                        return Err(OoxmlError::InvalidFormat(
                            "LINK \\a switch does not take an argument".to_string(),
                        ));
                    }
                    automatic_updates = true;
                },
                'f' => {
                    let argument = switch.argument.as_deref().ok_or_else(|| {
                        OoxmlError::InvalidFormat(
                            "LINK \\f switch requires an integral formatting mode".to_string(),
                        )
                    })?;
                    let value = argument.parse::<i64>().map_err(|_| {
                        OoxmlError::InvalidFormat(
                            "LINK \\f formatting mode must be an integer".to_string(),
                        )
                    })?;
                    formatting_modes.push(match value {
                        0 => LinkFormatting::Source,
                        2 => LinkFormatting::Destination,
                        4 => LinkFormatting::SpreadsheetSource,
                        5 => LinkFormatting::SpreadsheetDestination,
                        other => LinkFormatting::Unsupported(other),
                    });
                },
                'b' | 'd' | 'h' | 'p' | 'r' | 't' | 'u' => {
                    if switch.argument.is_some() {
                        return Err(OoxmlError::InvalidFormat(format!(
                            "LINK \\{} switch does not take an argument",
                            switch.name
                        )));
                    }
                    result_options.push(match switch.name {
                        'b' => LinkResultOption::Bitmap,
                        'd' => LinkResultOption::OmitGraphicData,
                        'h' => LinkResultOption::Html,
                        'p' => LinkResultOption::Picture,
                        'r' => LinkResultOption::RichText,
                        't' => LinkResultOption::Text,
                        'u' => LinkResultOption::UnicodeText,
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
    pub fn result_options(&self) -> &[LinkResultOption] {
        &self.result_options
    }

    /// Return the effective result or storage option under Word's documented
    /// last-switch behavior, if one was stored.
    pub fn effective_result_option(&self) -> Option<LinkResultOption> {
        self.result_options.last().copied()
    }

    /// Return integral `\\f` formatting modes in stored source order.
    ///
    /// These are metadata only; this API never formats linked content.
    pub fn formatting_modes(&self) -> &[LinkFormatting] {
        &self.formatting_modes
    }

    /// Return all stored field switches in source order.
    pub fn switches(&self) -> &[FieldSwitch] {
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
pub struct CitationField {
    instruction: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
    source_tags: Vec<String>,
    switches: Vec<FieldSwitch>,
}

impl CitationField {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some((primary_source_tag, switches)) =
            parse_citation_operand_and_switches(field.instruction())?
        else {
            return Ok(None);
        };
        let primary_source_tag = primary_source_tag.ok_or_else(|| {
            OoxmlError::InvalidFormat("CITATION field is missing its source tag".to_string())
        })?;
        if primary_source_tag.is_empty() {
            return Err(OoxmlError::InvalidFormat(
                "CITATION field source tag is empty".to_string(),
            ));
        }

        let mut source_tags = vec![primary_source_tag];
        for switch in &switches {
            if switch.name != 'm' {
                continue;
            }
            let source_tag = switch.argument.as_deref().ok_or_else(|| {
                OoxmlError::InvalidFormat("CITATION \\m switch requires a source tag".to_string())
            })?;
            if source_tag.is_empty() {
                return Err(OoxmlError::InvalidFormat(
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
    pub fn switches(&self) -> &[FieldSwitch] {
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
pub struct BibliographyField {
    instruction: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
    switches: Vec<FieldSwitch>,
}

impl BibliographyField {
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
    pub fn switches(&self) -> &[FieldSwitch] {
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
pub struct DocumentVariableField {
    instruction: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
    variable_name: String,
    switches: Vec<FieldSwitch>,
}

impl DocumentVariableField {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some((variable_name, switches)) =
            parse_field_operand_and_switches(field.instruction(), "DOCVARIABLE")?
        else {
            return Ok(None);
        };
        let variable_name = variable_name.ok_or_else(|| {
            OoxmlError::InvalidFormat("DOCVARIABLE field is missing its variable name".to_string())
        })?;
        if variable_name.is_empty() {
            return Err(OoxmlError::InvalidFormat(
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
    pub fn switches(&self) -> &[FieldSwitch] {
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
pub struct DocumentPropertyField {
    instruction: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
    property_name: String,
    switches: Vec<FieldSwitch>,
}

impl DocumentPropertyField {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        if !field.is_document_property() {
            return Ok(None);
        }
        if field.instruction().len() > MAX_DOCUMENT_PROPERTY_FIELD_INSTRUCTION_BYTES {
            return Err(OoxmlError::InvalidFormat(format!(
                "DOCPROPERTY field instruction exceeds {MAX_DOCUMENT_PROPERTY_FIELD_INSTRUCTION_BYTES} bytes"
            )));
        }
        let Some((property_name, switches)) =
            parse_field_operand_and_switches(field.instruction(), "DOCPROPERTY")?
        else {
            unreachable!("document-property recognition and parsing must agree");
        };
        let property_name = property_name.ok_or_else(|| {
            OoxmlError::InvalidFormat("DOCPROPERTY field is missing its property name".to_string())
        })?;
        if property_name.is_empty() {
            return Err(OoxmlError::InvalidFormat(
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
    pub fn switches(&self) -> &[FieldSwitch] {
        &self.switches
    }

    /// Check whether a case-insensitive ASCII switch appears in this field.
    pub fn has_switch(&self, name: char) -> bool {
        has_field_switch(&self.switches, name)
    }
}

/// The built-in Word document-information field category.
///
/// These fields are defined in ECMA-376 Part 1 §17.16.5. This enum preserves
/// the stored field kind only; it does not resolve document metadata or
/// calculate dates, revisions, or statistics.
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
pub struct DocumentInformationField {
    instruction: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
    kind: DocumentInformationFieldKind,
    switches: Vec<FieldSwitch>,
}

impl DocumentInformationField {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some(kind) = DocumentInformationFieldKind::from_instruction(field.instruction()) else {
            return Ok(None);
        };
        if field.instruction().len() > MAX_DOCUMENT_INFORMATION_FIELD_INSTRUCTION_BYTES {
            return Err(OoxmlError::InvalidFormat(format!(
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
    pub const fn kind(&self) -> DocumentInformationFieldKind {
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
    pub fn switches(&self) -> &[FieldSwitch] {
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
pub struct DocumentContextField {
    instruction: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
    kind: DocumentContextFieldKind,
    switches: Vec<FieldSwitch>,
}

impl DocumentContextField {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some(kind) = DocumentContextFieldKind::from_instruction(field.instruction()) else {
            return Ok(None);
        };
        if field.instruction().len() > MAX_DOCUMENT_CONTEXT_FIELD_INSTRUCTION_BYTES {
            return Err(OoxmlError::InvalidFormat(format!(
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
    pub const fn kind(&self) -> DocumentContextFieldKind {
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
    pub fn switches(&self) -> &[FieldSwitch] {
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
pub struct MergeField {
    instruction: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
    field_name: String,
    switches: Vec<FieldSwitch>,
}

impl MergeField {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some((field_name, switches)) =
            parse_field_operand_and_switches(field.instruction(), "MERGEFIELD")?
        else {
            return Ok(None);
        };
        let field_name = field_name.ok_or_else(|| {
            OoxmlError::InvalidFormat(
                "MERGEFIELD field is missing its data-column name".to_string(),
            )
        })?;
        if field_name.is_empty() {
            return Err(OoxmlError::InvalidFormat(
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
    pub fn switches(&self) -> &[FieldSwitch] {
        &self.switches
    }

    /// Check whether a case-insensitive ASCII switch appears in this field.
    pub fn has_switch(&self, name: char) -> bool {
        has_field_switch(&self.switches, name)
    }
}

/// The stored kind of a mail-merge counter field.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MailMergeCounterKind {
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
pub struct MailMergeCounterField {
    instruction: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
    kind: MailMergeCounterKind,
}

impl MailMergeCounterField {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let (kind, field_type, remainder) = if let Some(remainder) =
            field_instruction_remainder(field.instruction(), "MERGEREC")
        {
            (MailMergeCounterKind::Record, "MERGEREC", remainder)
        } else if let Some(remainder) = field_instruction_remainder(field.instruction(), "MERGESEQ")
        {
            (MailMergeCounterKind::Sequence, "MERGESEQ", remainder)
        } else {
            return Ok(None);
        };

        if !remainder.trim().is_empty() {
            return Err(OoxmlError::InvalidFormat(format!(
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
    pub fn kind(&self) -> MailMergeCounterKind {
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
pub struct MailMergeNextField {
    instruction: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
}

impl MailMergeNextField {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some(remainder) = field_instruction_remainder(field.instruction(), "NEXT") else {
            return Ok(None);
        };
        if !remainder.trim().is_empty() {
            return Err(OoxmlError::InvalidFormat(
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
pub enum MailMergeConditionalControlKind {
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
pub struct MailMergeConditionalControlField {
    instruction: String,
    comparison: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
    kind: MailMergeConditionalControlKind,
}

impl MailMergeConditionalControlField {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let (kind, field_type, comparison) = if let Some(comparison) =
            field_instruction_remainder(field.instruction(), "NEXTIF")
        {
            (
                MailMergeConditionalControlKind::NextIf,
                "NEXTIF",
                comparison,
            )
        } else if let Some(comparison) = field_instruction_remainder(field.instruction(), "SKIPIF")
        {
            (
                MailMergeConditionalControlKind::SkipIf,
                "SKIPIF",
                comparison,
            )
        } else {
            return Ok(None);
        };

        let comparison = comparison.trim();
        if comparison.is_empty() {
            return Err(OoxmlError::InvalidFormat(format!(
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
    pub fn kind(&self) -> MailMergeConditionalControlKind {
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
pub struct IfField {
    instruction: String,
    expression: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
}

impl IfField {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some(expression) = field_instruction_remainder(field.instruction(), "IF") else {
            return Ok(None);
        };
        let expression = expression.trim();
        if expression.is_empty() {
            return Err(OoxmlError::InvalidFormat(
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
pub struct SetField {
    instruction: String,
    target_name: String,
    expression: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
}

impl SetField {
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
pub struct SequenceField {
    instruction: String,
    identifier: String,
    bookmark: Option<String>,
    tail: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
}

impl SequenceField {
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
pub struct FormulaField {
    instruction: String,
    formula: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
}

impl FormulaField {
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

/// A typed, inert Word `QUOTE` field.
///
/// ECMA-376 Part 1 §17.16.5.49 defines `QUOTE` with a text field argument.
/// This type retains that stored argument, switches, and cached result only. It
/// never interprets character codes, expands nested fields, inserts text, or
/// refreshes a field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct QuoteField {
    instruction: String,
    text: String,
    switches: Vec<FieldSwitch>,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
}

impl QuoteField {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        if field.instruction.len() > MAX_QUOTE_FIELD_INSTRUCTION_BYTES {
            return Err(OoxmlError::InvalidFormat(format!(
                "QUOTE field instruction exceeds {MAX_QUOTE_FIELD_INSTRUCTION_BYTES} bytes"
            )));
        }
        let Some((text, switches)) =
            parse_field_operand_and_switches(field.instruction(), "QUOTE")?
        else {
            return Ok(None);
        };
        let text = text.ok_or_else(|| {
            OoxmlError::InvalidFormat("QUOTE field is missing its text argument".to_string())
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
    pub fn switches(&self) -> &[FieldSwitch] {
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
pub struct SymbolField {
    instruction: String,
    character_argument: String,
    switches: Vec<FieldSwitch>,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
}

impl SymbolField {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        if field.instruction.len() > MAX_SYMBOL_FIELD_INSTRUCTION_BYTES {
            return Err(OoxmlError::InvalidFormat(format!(
                "SYMBOL field instruction exceeds {MAX_SYMBOL_FIELD_INSTRUCTION_BYTES} bytes"
            )));
        }
        let Some((character_argument, switches)) =
            parse_field_operand_and_switches(field.instruction(), "SYMBOL")?
        else {
            return Ok(None);
        };
        let character_argument = character_argument.ok_or_else(|| {
            OoxmlError::InvalidFormat("SYMBOL field is missing its character argument".to_string())
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
    pub fn switches(&self) -> &[FieldSwitch] {
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
pub struct StyleReferenceField {
    instruction: String,
    style_name: String,
    options: Vec<StyleReferenceFieldOption>,
    unknown_switches: Vec<FieldSwitch>,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
}

impl StyleReferenceField {
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
    pub fn options(&self) -> &[StyleReferenceFieldOption] {
        &self.options
    }

    /// Return unrecognized stored switches in source order.
    ///
    /// They are retained without interpretation or execution.
    pub fn unknown_switches(&self) -> &[FieldSwitch] {
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
pub struct CompareField {
    instruction: String,
    comparison: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
}

impl CompareField {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some(comparison) = field_instruction_remainder(field.instruction(), "COMPARE") else {
            return Ok(None);
        };
        let comparison = comparison.trim();
        if comparison.is_empty() {
            return Err(OoxmlError::InvalidFormat(
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

/// The stored kind of a prompt field.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PromptFieldKind {
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
pub struct PromptField {
    instruction: String,
    kind: PromptFieldKind,
    bookmark: Option<String>,
    prompt: Option<String>,
    default_response: Option<String>,
    prompts_once_per_mail_merge: bool,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
}

impl PromptField {
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
    pub fn kind(&self) -> PromptFieldKind {
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

/// A typed, inert Word `ADDRESSBLOCK` or `GREETINGLINE` field.
///
/// ECMA-376 Part 1 §§17.16.5.1 and 17.16.5.24 define these mail-merge
/// recipient layout fields. This type exposes stored layout metadata and the
/// cached result only. It never opens a data source, selects a record,
/// performs a merge, expands placeholders, generates text, or refreshes a
/// field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MailMergeRecipientField {
    instruction: String,
    kind: MailMergeRecipientFieldKind,
    country_inclusion: Option<AddressBlockCountryInclusion>,
    formats_using_recipient_country: bool,
    excluded_countries: Vec<String>,
    format_template: Option<String>,
    language: Option<String>,
    greeting_fallback_text: Option<String>,
    unknown_switches: Vec<FieldSwitch>,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
}

impl MailMergeRecipientField {
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
    pub fn kind(&self) -> MailMergeRecipientFieldKind {
        self.kind
    }

    /// Return how an `ADDRESSBLOCK` requests country/region text.
    ///
    /// This is `None` when the instruction has no `\\c` switch or when the
    /// field is a `GREETINGLINE`. The stored request is never used to render
    /// an address.
    pub fn country_inclusion(&self) -> Option<AddressBlockCountryInclusion> {
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
    pub fn unknown_switches(&self) -> &[FieldSwitch] {
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
pub struct MacroButtonField {
    instruction: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
    macro_name: String,
    display_text: String,
}

impl MacroButtonField {
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
pub struct GoToButtonField {
    instruction: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
    target: String,
    button_text: String,
}

impl GoToButtonField {
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
pub enum ActiveContentFieldKind {
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
pub struct ActiveContentField {
    instruction: String,
    kind: ActiveContentFieldKind,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
}

impl ActiveContentField {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let kind = if field_instruction_remainder(field.instruction(), "ADDIN").is_some() {
            ActiveContentFieldKind::AddIn
        } else if field_instruction_remainder(field.instruction(), "CONTROL").is_some() {
            ActiveContentFieldKind::OcxControl
        } else if field_instruction_remainder(field.instruction(), "HTMLCONTROL").is_some() {
            ActiveContentFieldKind::HtmlControl
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
    pub fn kind(&self) -> ActiveContentFieldKind {
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

/// The stored kind of a Word building-block field.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AutoTextFieldKind {
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
pub struct AutoTextField {
    instruction: String,
    kind: AutoTextFieldKind,
    entry_name: String,
    unknown_switches: Vec<FieldSwitch>,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
}

impl AutoTextField {
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
    pub fn kind(&self) -> AutoTextFieldKind {
        self.kind
    }

    /// Return the stored building-block entry name without resolving it.
    pub fn entry_name(&self) -> &str {
        &self.entry_name
    }

    /// Return unrecognized stored switches in source order.
    ///
    /// They are retained without interpretation or execution.
    pub fn unknown_switches(&self) -> &[FieldSwitch] {
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
pub struct AutoTextListField {
    instruction: String,
    display_text: Option<String>,
    options: Vec<AutoTextListOption>,
    unknown_switches: Vec<FieldSwitch>,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
}

impl AutoTextListField {
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
    pub fn unknown_switches(&self) -> &[FieldSwitch] {
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

/// A typed, inert Word `USERADDRESS`, `USERINITIALS`, or `USERNAME` field.
///
/// ECMA-376 Part 1 §§17.16.5.69–71 define these fields. This type exposes a
/// stored override, formatting request, and cached result only. It never reads
/// or modifies a host user's identity, applies formatting, or refreshes a
/// field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct UserIdentityField {
    instruction: String,
    kind: UserIdentityFieldKind,
    override_value: Option<String>,
    formatting: Option<UserIdentityFormatting>,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
}

impl UserIdentityField {
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
    pub fn kind(&self) -> UserIdentityFieldKind {
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
    pub fn formatting(&self) -> Option<UserIdentityFormatting> {
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

impl AdvanceFieldAdjustment {
    /// Return the requested placement operation.
    pub fn operation(&self) -> AdvanceFieldOperation {
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
pub struct AdvanceField {
    instruction: String,
    adjustments: Vec<AdvanceFieldAdjustment>,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
}

impl AdvanceField {
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
    pub fn adjustments(&self) -> &[AdvanceFieldAdjustment] {
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
pub type TableOfContentsSwitch = FieldSwitch;

/// An inclusive heading-level range selected by a `TOC \o` switch.
///
/// WordprocessingML heading levels are bounded to one through nine.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct TableOfContentsLevelRange {
    start: u8,
    end: u8,
}

impl TableOfContentsLevelRange {
    /// Create a valid inclusive heading-level range.
    pub fn new(start: u8, end: u8) -> Result<Self> {
        if !(1..=9).contains(&start) || !(1..=9).contains(&end) || start > end {
            return Err(OoxmlError::InvalidFormat(
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
pub struct TableOfContentsField {
    instruction: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
    switches: Vec<FieldSwitch>,
}

impl TableOfContentsField {
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
    pub fn switches(&self) -> &[FieldSwitch] {
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
    pub fn heading_style_levels(&self) -> Result<Vec<TableOfContentsLevelRange>> {
        self.switches
            .iter()
            .filter(|switch| switch.name == 'o')
            .map(|switch| {
                let value = switch.argument.as_deref().ok_or_else(|| {
                    OoxmlError::InvalidFormat(
                        "TOC \\o switch requires a heading-level range".to_string(),
                    )
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

/// A typed, inert Word table-of-authorities (`TOA`) field.
///
/// A TOA collects stored `TA` citation-marker fields into a rendered list.
/// This model exposes only the persisted code and cached result: it does not
/// find citations, paginate the document, generate authorities, or execute
/// any field instruction.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TableOfAuthoritiesField {
    instruction: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
    switches: Vec<FieldSwitch>,
}

impl TableOfAuthoritiesField {
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
    pub fn switches(&self) -> &[FieldSwitch] {
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
pub struct TableOfAuthoritiesEntryField {
    instruction: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
    switches: Vec<FieldSwitch>,
}

impl TableOfAuthoritiesEntryField {
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
    pub fn switches(&self) -> &[FieldSwitch] {
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
pub enum IndexSortOrder {
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
pub struct IndexField {
    instruction: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
    switches: Vec<FieldSwitch>,
}

impl IndexField {
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
    pub fn switches(&self) -> &[FieldSwitch] {
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
    pub fn sort_order(&self) -> Result<Option<IndexSortOrder>> {
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
pub struct IndexEntryField {
    instruction: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
    entry: String,
    switches: Vec<FieldSwitch>,
}

impl IndexEntryField {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some((entry, switches)) = parse_field_operand_and_switches(field.instruction(), "XE")?
        else {
            return Ok(None);
        };
        let entry = entry.ok_or_else(|| {
            OoxmlError::InvalidFormat("XE field is missing its index-entry text".to_string())
        })?;
        if entry.is_empty() {
            return Err(OoxmlError::InvalidFormat(
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
    pub fn switches(&self) -> &[FieldSwitch] {
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

fn has_field_switch(switches: &[FieldSwitch], name: char) -> bool {
    switches
        .iter()
        .any(|switch| switch.name.eq_ignore_ascii_case(&name))
}

fn optional_field_switch_argument<'a>(
    switches: &'a [FieldSwitch],
    name: char,
    field_type: &str,
) -> Result<Option<&'a str>> {
    let mut matching = switches
        .iter()
        .filter(|switch| switch.name.eq_ignore_ascii_case(&name));
    let Some(switch) = matching.next() else {
        return Ok(None);
    };
    if matching.next().is_some() {
        return Err(OoxmlError::InvalidFormat(format!(
            "{field_type} field has duplicate \\{name} switches"
        )));
    }
    switch.argument.as_deref().map(Some).ok_or_else(|| {
        OoxmlError::InvalidFormat(format!("{field_type} \\{name} switch requires an argument"))
    })
}

fn parse_authority_category(value: &str, minimum: u8, field_type: &str) -> Result<u8> {
    let value = value.parse::<u8>().map_err(|_| {
        OoxmlError::InvalidFormat(format!("{field_type} authority category is not an integer"))
    })?;
    if !(minimum..=16).contains(&value) {
        return Err(OoxmlError::InvalidFormat(format!(
            "{field_type} authority category must be in {minimum}..=16"
        )));
    }
    Ok(value)
}

fn parse_index_columns(value: &str) -> Result<u8> {
    let columns = value.parse::<u8>().map_err(|_| {
        OoxmlError::InvalidFormat("INDEX column count is not an integer".to_string())
    })?;
    if !(1..=4).contains(&columns) {
        return Err(OoxmlError::InvalidFormat(
            "INDEX column count must be in 1..=4".to_string(),
        ));
    }
    Ok(columns)
}

fn parse_index_sort_order(value: &str) -> Result<IndexSortOrder> {
    match value {
        "S" | "s" => Ok(IndexSortOrder::Stroke),
        "P" | "p" => Ok(IndexSortOrder::Pronunciation),
        _ => Err(OoxmlError::InvalidFormat(format!(
            "INDEX \\o sort order must be S or P, got {value:?}"
        ))),
    }
}

fn field_instruction_remainder<'a>(instruction: &'a str, field_type: &str) -> Option<&'a str> {
    let instruction = instruction.trim_start();
    let field_type_end = field_type.len();
    let candidate = instruction.get(..field_type_end)?;
    let remainder = instruction.get(field_type_end..)?;
    if !candidate.eq_ignore_ascii_case(field_type) {
        return None;
    }
    match remainder.chars().next() {
        None | Some('\\') | Some('"') => Some(remainder),
        Some(character) if character.is_whitespace() => Some(remainder),
        Some(_) => None,
    }
}

fn parse_field_switches(instruction: &str, field_type: &str) -> Result<Option<Vec<FieldSwitch>>> {
    let Some(remainder) = field_instruction_remainder(instruction, field_type) else {
        return Ok(None);
    };
    let mut characters = remainder.chars().peekable();
    Ok(Some(parse_field_switches_from_characters(
        &mut characters,
        field_type,
    )?))
}

fn parse_field_operand_and_switches(
    instruction: &str,
    field_type: &str,
) -> Result<Option<(Option<String>, Vec<FieldSwitch>)>> {
    let Some(remainder) = field_instruction_remainder(instruction, field_type) else {
        return Ok(None);
    };
    let mut characters = remainder.chars().peekable();
    skip_field_whitespace(&mut characters);
    let operand = match characters.peek().copied() {
        None | Some('\\') => None,
        Some('"') => {
            characters.next();
            Some(parse_field_quoted_argument(&mut characters, field_type)?)
        },
        Some(_) => Some(parse_field_unquoted_argument(&mut characters)),
    };
    let switches = parse_field_switches_from_characters(&mut characters, field_type)?;
    Ok(Some((operand, switches)))
}

fn parse_macro_button_operands(instruction: &str) -> Result<Option<(String, String)>> {
    let Some(remainder) = field_instruction_remainder(instruction, "MACROBUTTON") else {
        return Ok(None);
    };
    let mut characters = remainder.chars().peekable();
    let macro_name = parse_next_field_argument(&mut characters, "MACROBUTTON")?
        .filter(|value| !value.is_empty())
        .ok_or_else(|| {
            OoxmlError::InvalidFormat(
                "MACROBUTTON field is missing its macro or command name".to_string(),
            )
        })?;
    let display_text = parse_next_field_argument(&mut characters, "MACROBUTTON")?
        .filter(|value| !value.is_empty())
        .ok_or_else(|| {
            OoxmlError::InvalidFormat("MACROBUTTON field is missing its button text".to_string())
        })?;
    skip_field_whitespace(&mut characters);
    if characters.next().is_some() {
        return Err(OoxmlError::InvalidFormat(
            "MACROBUTTON field must contain exactly two arguments and no switches".to_string(),
        ));
    }
    Ok(Some((macro_name, display_text)))
}

fn parse_go_to_button_operands(instruction: &str) -> Result<Option<(String, String)>> {
    let Some(remainder) = field_instruction_remainder(instruction, "GOTOBUTTON") else {
        return Ok(None);
    };
    let mut characters = remainder.chars().peekable();
    let target = parse_next_field_argument(&mut characters, "GOTOBUTTON")?
        .filter(|value| !value.is_empty())
        .ok_or_else(|| {
            OoxmlError::InvalidFormat("GOTOBUTTON field is missing its destination".to_string())
        })?;
    let button_text = parse_next_field_argument(&mut characters, "GOTOBUTTON")?
        .filter(|value| !value.is_empty())
        .ok_or_else(|| {
            OoxmlError::InvalidFormat("GOTOBUTTON field is missing its button text".to_string())
        })?;
    skip_field_whitespace(&mut characters);
    if characters.next().is_some() {
        return Err(OoxmlError::InvalidFormat(
            "GOTOBUTTON field must contain exactly two arguments and no switches".to_string(),
        ));
    }
    Ok(Some((target, button_text)))
}

fn parse_user_identity_field_parts(
    instruction: &str,
) -> Result<
    Option<(
        UserIdentityFieldKind,
        Option<String>,
        Option<UserIdentityFormatting>,
    )>,
> {
    let (kind, field_type, remainder) =
        if let Some(remainder) = field_instruction_remainder(instruction, "USERADDRESS") {
            (UserIdentityFieldKind::Address, "USERADDRESS", remainder)
        } else if let Some(remainder) = field_instruction_remainder(instruction, "USERINITIALS") {
            (UserIdentityFieldKind::Initials, "USERINITIALS", remainder)
        } else if let Some(remainder) = field_instruction_remainder(instruction, "USERNAME") {
            (UserIdentityFieldKind::Name, "USERNAME", remainder)
        } else {
            return Ok(None);
        };

    let mut characters = remainder.chars().peekable();
    let override_value = parse_next_field_argument(&mut characters, field_type)?;
    let switches = parse_field_switches_from_characters(&mut characters, field_type)?;
    let mut formatting = None;
    for switch in switches {
        if switch.name != '*' {
            return Err(OoxmlError::InvalidFormat(format!(
                "{field_type} field has an unsupported \\{} switch",
                switch.name
            )));
        }
        if formatting.is_some() {
            return Err(OoxmlError::InvalidFormat(format!(
                "{field_type} field repeats its \\* switch"
            )));
        }
        let argument = switch.argument.ok_or_else(|| {
            OoxmlError::InvalidFormat(format!(
                "{field_type} \\* switch requires a general-formatting argument"
            ))
        })?;
        formatting = Some(if argument.eq_ignore_ascii_case("Caps") {
            UserIdentityFormatting::Caps
        } else if argument.eq_ignore_ascii_case("FirstCap") {
            UserIdentityFormatting::FirstCap
        } else if argument.eq_ignore_ascii_case("Lower") {
            UserIdentityFormatting::Lower
        } else if argument.eq_ignore_ascii_case("Upper") {
            UserIdentityFormatting::Upper
        } else {
            return Err(OoxmlError::InvalidFormat(format!(
                "{field_type} \\* switch must be Caps, FirstCap, Lower, or Upper"
            )));
        });
    }

    Ok(Some((kind, override_value, formatting)))
}

fn parse_advance_field_adjustments(
    instruction: &str,
) -> Result<Option<Vec<AdvanceFieldAdjustment>>> {
    let Some(switches) = parse_field_switches(instruction, "ADVANCE")? else {
        return Ok(None);
    };

    let mut adjustments = Vec::with_capacity(switches.len());
    for switch in switches {
        let operation = match switch.name {
            'd' => AdvanceFieldOperation::Down,
            'l' => AdvanceFieldOperation::Left,
            'r' => AdvanceFieldOperation::Right,
            'u' => AdvanceFieldOperation::Up,
            'x' => AdvanceFieldOperation::HorizontalPosition,
            'y' => AdvanceFieldOperation::VerticalPosition,
            name => {
                return Err(OoxmlError::InvalidFormat(format!(
                    "ADVANCE field has an unsupported \\{name} switch"
                )));
            },
        };
        let points = switch.argument.ok_or_else(|| {
            OoxmlError::InvalidFormat(format!(
                "ADVANCE \\{} switch requires an integral number of points",
                switch.name
            ))
        })?;
        let points = points.parse::<i64>().map_err(|_| {
            OoxmlError::InvalidFormat(format!(
                "ADVANCE \\{} switch must specify an integral number of points",
                switch.name
            ))
        })?;
        adjustments.push(AdvanceFieldAdjustment { operation, points });
    }

    Ok(Some(adjustments))
}

fn parse_link_operands_and_switches(
    instruction: &str,
) -> Result<Option<(String, String, Option<String>, Vec<FieldSwitch>)>> {
    parse_external_link_operands_and_switches(instruction, "LINK")
}

fn parse_dde_operands_and_switches(
    instruction: &str,
) -> Result<
    Option<(
        DdeFieldKind,
        String,
        String,
        Option<String>,
        Vec<FieldSwitch>,
    )>,
> {
    if let Some((application, source, item, switches)) =
        parse_external_link_operands_and_switches(instruction, "DDEAUTO")?
    {
        return Ok(Some((
            DdeFieldKind::DdeAuto,
            application,
            source,
            item,
            switches,
        )));
    }

    Ok(
        parse_external_link_operands_and_switches(instruction, "DDE")?.map(
            |(application, source, item, switches)| {
                (DdeFieldKind::Dde, application, source, item, switches)
            },
        ),
    )
}

fn parse_external_include_operands_and_switches(
    instruction: &str,
) -> Result<Option<(IncludeFieldKind, String, Option<String>, Vec<FieldSwitch>)>> {
    let (kind, field_type) = if field_instruction_remainder(instruction, "INCLUDETEXT").is_some() {
        (IncludeFieldKind::Text, "INCLUDETEXT")
    } else if field_instruction_remainder(instruction, "INCLUDEPICTURE").is_some() {
        (IncludeFieldKind::Picture, "INCLUDEPICTURE")
    } else {
        return Ok(None);
    };
    let remainder =
        field_instruction_remainder(instruction, field_type).expect("recognized include field");
    let mut characters = remainder.chars().peekable();
    let source = parse_next_field_argument(&mut characters, field_type)?
        .filter(|value| !value.is_empty())
        .ok_or_else(|| {
            OoxmlError::InvalidFormat(format!("{field_type} field is missing its source"))
        })?;
    let bookmark = match kind {
        IncludeFieldKind::Text => parse_next_field_argument(&mut characters, field_type)?,
        IncludeFieldKind::Picture => None,
    };
    let switches = parse_field_switches_from_characters(&mut characters, field_type)?;
    Ok(Some((kind, source, bookmark, switches)))
}

fn required_external_include_option_argument(
    switch: &FieldSwitch,
    kind: IncludeFieldKind,
) -> Result<String> {
    let field_type = match kind {
        IncludeFieldKind::Text => "INCLUDETEXT",
        IncludeFieldKind::Picture => "INCLUDEPICTURE",
    };
    switch.argument.clone().ok_or_else(|| {
        OoxmlError::InvalidFormat(format!(
            "{field_type} {} switch requires an argument",
            switch.name
        ))
    })
}

fn parse_external_link_operands_and_switches(
    instruction: &str,
    field_type: &str,
) -> Result<Option<(String, String, Option<String>, Vec<FieldSwitch>)>> {
    let Some(remainder) = field_instruction_remainder(instruction, field_type) else {
        return Ok(None);
    };
    let mut characters = remainder.chars().peekable();
    let application_type = parse_next_field_argument(&mut characters, field_type)?
        .filter(|value| !value.is_empty())
        .ok_or_else(|| {
            OoxmlError::InvalidFormat(format!("{field_type} field is missing its application type"))
        })?;
    let source = parse_next_field_argument(&mut characters, field_type)?
        .filter(|value| !value.is_empty())
        .ok_or_else(|| {
            OoxmlError::InvalidFormat(format!("{field_type} field is missing its source"))
        })?;
    let item = parse_next_field_argument(&mut characters, field_type)?;
    let switches = parse_field_switches_from_characters(&mut characters, field_type)?;
    Ok(Some((application_type, source, item, switches)))
}

fn parse_next_field_argument(
    characters: &mut std::iter::Peekable<std::str::Chars<'_>>,
    field_type: &str,
) -> Result<Option<String>> {
    skip_field_whitespace(characters);
    match characters.peek().copied() {
        None | Some('\\') => Ok(None),
        Some('"') => {
            characters.next();
            Ok(Some(parse_field_quoted_argument(characters, field_type)?))
        },
        Some(_) => Ok(Some(parse_field_unquoted_argument(characters))),
    }
}

fn parse_set_field_parts(instruction: &str) -> Result<Option<(String, String)>> {
    if instruction.len() > MAX_SET_FIELD_INSTRUCTION_BYTES {
        return Err(OoxmlError::InvalidFormat(format!(
            "SET field instruction exceeds {MAX_SET_FIELD_INSTRUCTION_BYTES} bytes"
        )));
    }

    let Some(remainder) = field_instruction_remainder(instruction, "SET") else {
        return Ok(None);
    };
    let mut characters = remainder.chars().peekable();
    let target_name = parse_next_field_argument(&mut characters, "SET")?
        .filter(|value| !value.is_empty())
        .ok_or_else(|| {
            OoxmlError::InvalidFormat("SET field is missing its target name".to_string())
        })?;
    skip_field_whitespace(&mut characters);
    let expression = characters.collect::<String>();
    if expression.trim().is_empty() {
        return Err(OoxmlError::InvalidFormat(
            "SET field is missing its expression".to_string(),
        ));
    }

    Ok(Some((target_name, expression)))
}

fn parse_sequence_field_parts(
    instruction: &str,
) -> Result<Option<(String, Option<String>, String)>> {
    if instruction.len() > MAX_SEQUENCE_FIELD_INSTRUCTION_BYTES {
        return Err(OoxmlError::InvalidFormat(format!(
            "SEQ field instruction exceeds {MAX_SEQUENCE_FIELD_INSTRUCTION_BYTES} bytes"
        )));
    }

    let Some(remainder) = field_instruction_remainder(instruction, "SEQ") else {
        return Ok(None);
    };
    let mut characters = remainder.chars().peekable();
    let identifier = parse_next_field_argument(&mut characters, "SEQ")?
        .filter(|value| !value.is_empty())
        .ok_or_else(|| {
            OoxmlError::InvalidFormat("SEQ field is missing its identifier".to_string())
        })?;
    skip_field_whitespace(&mut characters);
    let bookmark = match characters.peek().copied() {
        None | Some('\\') => None,
        Some(_) => Some(
            parse_next_field_argument(&mut characters, "SEQ")?
                .filter(|value| !value.is_empty())
                .ok_or_else(|| {
                    OoxmlError::InvalidFormat("SEQ field bookmark is empty".to_string())
                })?,
        ),
    };
    skip_field_whitespace(&mut characters);
    let tail = characters.collect::<String>().trim().to_string();

    Ok(Some((identifier, bookmark, tail)))
}

fn parse_formula_field_formula(instruction: &str) -> Result<Option<String>> {
    if instruction.len() > MAX_FORMULA_FIELD_INSTRUCTION_BYTES {
        return Err(OoxmlError::InvalidFormat(format!(
            "formula field instruction exceeds {MAX_FORMULA_FIELD_INSTRUCTION_BYTES} bytes"
        )));
    }

    let Some(formula) = instruction.trim().strip_prefix('=') else {
        return Ok(None);
    };
    let formula = formula.trim();
    if formula.is_empty() {
        return Err(OoxmlError::InvalidFormat(
            "formula field is missing its formula".to_string(),
        ));
    }

    Ok(Some(formula.to_string()))
}

fn parse_style_reference_field_parts(
    instruction: &str,
) -> Result<Option<(String, Vec<StyleReferenceFieldOption>, Vec<FieldSwitch>)>> {
    if instruction.len() > MAX_STYLE_REFERENCE_FIELD_INSTRUCTION_BYTES {
        return Err(OoxmlError::InvalidFormat(format!(
            "STYLEREF field instruction exceeds {MAX_STYLE_REFERENCE_FIELD_INSTRUCTION_BYTES} bytes"
        )));
    }

    let Some(remainder) = field_instruction_remainder(instruction, "STYLEREF") else {
        return Ok(None);
    };
    let mut characters = remainder.chars().peekable();
    let style_name = parse_next_field_argument(&mut characters, "STYLEREF")?
        .filter(|value| !value.is_empty())
        .ok_or_else(|| {
            OoxmlError::InvalidFormat("STYLEREF field is missing its style name".to_string())
        })?;
    let switches = parse_field_switches_from_characters(&mut characters, "STYLEREF")?;
    let mut options = Vec::new();
    let mut unknown_switches = Vec::new();
    for switch in switches {
        let option = match switch.name {
            'l' => Some(StyleReferenceFieldOption::FollowingText),
            'n' => Some(StyleReferenceFieldOption::ParagraphNumber),
            'p' => Some(StyleReferenceFieldOption::RelativePosition),
            'r' => Some(StyleReferenceFieldOption::ParagraphNumberRelativeContext),
            't' => Some(StyleReferenceFieldOption::SuppressNonNumberText),
            'w' => Some(StyleReferenceFieldOption::ParagraphNumberFullContext),
            _ => None,
        };
        if let Some(option) = option {
            if switch.argument.is_some() {
                return Err(OoxmlError::InvalidFormat(format!(
                    "STYLEREF \\\\{} switch does not take an argument",
                    switch.name
                )));
            }
            options.push(option);
        } else {
            unknown_switches.push(switch);
        }
    }

    Ok(Some((style_name, options, unknown_switches)))
}

fn parse_auto_text_field_parts(
    instruction: &str,
) -> Result<Option<(AutoTextFieldKind, String, Vec<FieldSwitch>)>> {
    let (kind, field_type, remainder) =
        if let Some(remainder) = field_instruction_remainder(instruction, "GLOSSARY") {
            (AutoTextFieldKind::Glossary, "GLOSSARY", remainder)
        } else if let Some(remainder) = field_instruction_remainder(instruction, "AUTOTEXT") {
            (AutoTextFieldKind::AutoText, "AUTOTEXT", remainder)
        } else {
            return Ok(None);
        };
    if instruction.len() > MAX_AUTO_TEXT_FIELD_INSTRUCTION_BYTES {
        return Err(OoxmlError::InvalidFormat(format!(
            "{field_type} field instruction exceeds {MAX_AUTO_TEXT_FIELD_INSTRUCTION_BYTES} bytes"
        )));
    }

    let mut characters = remainder.chars().peekable();
    let entry_name = parse_next_field_argument(&mut characters, field_type)?
        .filter(|value| !value.is_empty())
        .ok_or_else(|| {
            OoxmlError::InvalidFormat(format!("{field_type} field is missing its entry name"))
        })?;
    let unknown_switches = parse_field_switches_from_characters(&mut characters, field_type)?;
    Ok(Some((kind, entry_name, unknown_switches)))
}

fn parse_auto_text_list_field_parts(
    instruction: &str,
) -> Result<Option<(Option<String>, Vec<AutoTextListOption>, Vec<FieldSwitch>)>> {
    let Some(remainder) = field_instruction_remainder(instruction, "AUTOTEXTLIST") else {
        return Ok(None);
    };
    if instruction.len() > MAX_AUTO_TEXT_LIST_FIELD_INSTRUCTION_BYTES {
        return Err(OoxmlError::InvalidFormat(format!(
            "AUTOTEXTLIST field instruction exceeds {MAX_AUTO_TEXT_LIST_FIELD_INSTRUCTION_BYTES} bytes"
        )));
    }

    let mut characters = remainder.chars().peekable();
    skip_field_whitespace(&mut characters);
    let display_text = match characters.peek().copied() {
        None | Some('\\') => None,
        Some(_) => parse_next_field_argument(&mut characters, "AUTOTEXTLIST")?,
    };
    let switches = parse_field_switches_from_characters(&mut characters, "AUTOTEXTLIST")?;
    let mut options = Vec::new();
    let mut unknown_switches = Vec::new();
    for switch in switches {
        match switch.name {
            's' => {
                let style = switch.argument.ok_or_else(|| {
                    OoxmlError::InvalidFormat(
                        "AUTOTEXTLIST \\s switch requires an argument".to_string(),
                    )
                })?;
                options.push(AutoTextListOption::Style(style));
            },
            't' => {
                let tip = switch.argument.ok_or_else(|| {
                    OoxmlError::InvalidFormat(
                        "AUTOTEXTLIST \\t switch requires an argument".to_string(),
                    )
                })?;
                options.push(AutoTextListOption::Tip(tip));
            },
            _ => unknown_switches.push(switch),
        }
    }
    Ok(Some((display_text, options, unknown_switches)))
}

fn parse_prompt_field_parts(
    instruction: &str,
) -> Result<
    Option<(
        PromptFieldKind,
        Option<String>,
        Option<String>,
        Option<String>,
        bool,
    )>,
> {
    let (kind, field_type, remainder) =
        if let Some(remainder) = field_instruction_remainder(instruction, "ASK") {
            (PromptFieldKind::Ask, "ASK", remainder)
        } else if let Some(remainder) = field_instruction_remainder(instruction, "FILLIN") {
            (PromptFieldKind::FillIn, "FILLIN", remainder)
        } else {
            return Ok(None);
        };

    let mut characters = remainder.chars().peekable();
    let (bookmark, prompt) = match kind {
        PromptFieldKind::Ask => {
            let bookmark = parse_next_field_argument(&mut characters, field_type)?
                .filter(|value| !value.is_empty())
                .ok_or_else(|| {
                    OoxmlError::InvalidFormat("ASK field is missing its bookmark name".to_string())
                })?;
            let prompt =
                parse_next_field_argument(&mut characters, field_type)?.ok_or_else(|| {
                    OoxmlError::InvalidFormat("ASK field is missing its prompt text".to_string())
                })?;
            (Some(bookmark), Some(prompt))
        },
        PromptFieldKind::FillIn => (
            None,
            parse_next_field_argument(&mut characters, field_type)?,
        ),
    };

    let switches = parse_field_switches_from_characters(&mut characters, field_type)?;
    let mut default_response = None;
    let mut prompts_once_per_mail_merge = false;
    for switch in switches {
        match switch.name {
            'd' => {
                if default_response.is_some() {
                    return Err(OoxmlError::InvalidFormat(format!(
                        "{field_type} field repeats its \\d switch"
                    )));
                }
                default_response = Some(switch.argument.ok_or_else(|| {
                    OoxmlError::InvalidFormat(format!(
                        "{field_type} field requires an argument for its \\d switch"
                    ))
                })?);
            },
            'o' => {
                if prompts_once_per_mail_merge {
                    return Err(OoxmlError::InvalidFormat(format!(
                        "{field_type} field repeats its \\o switch"
                    )));
                }
                if switch.argument.is_some() {
                    return Err(OoxmlError::InvalidFormat(format!(
                        "{field_type} field does not allow an argument for its \\o switch"
                    )));
                }
                prompts_once_per_mail_merge = true;
            },
            _ => {
                return Err(OoxmlError::InvalidFormat(format!(
                    "{field_type} field has an unsupported \\{} switch",
                    switch.name
                )));
            },
        }
    }

    Ok(Some((
        kind,
        bookmark,
        prompt,
        default_response,
        prompts_once_per_mail_merge,
    )))
}

#[allow(clippy::type_complexity)]
fn parse_mail_merge_recipient_field_parts(
    instruction: &str,
) -> Result<
    Option<(
        MailMergeRecipientFieldKind,
        Option<AddressBlockCountryInclusion>,
        bool,
        Vec<String>,
        Option<String>,
        Option<String>,
        Option<String>,
        Vec<FieldSwitch>,
    )>,
> {
    let (kind, field_type, remainder) =
        if let Some(remainder) = field_instruction_remainder(instruction, "ADDRESSBLOCK") {
            (
                MailMergeRecipientFieldKind::AddressBlock,
                "ADDRESSBLOCK",
                remainder,
            )
        } else if let Some(remainder) = field_instruction_remainder(instruction, "GREETINGLINE") {
            (
                MailMergeRecipientFieldKind::GreetingLine,
                "GREETINGLINE",
                remainder,
            )
        } else {
            return Ok(None);
        };

    let mut characters = remainder.chars().peekable();
    let switches = parse_field_switches_from_characters(&mut characters, field_type)?;
    let mut country_inclusion = None;
    let mut formats_using_recipient_country = false;
    let mut excluded_countries = Vec::new();
    let mut format_template = None;
    let mut language = None;
    let mut greeting_fallback_text = None;
    let mut unknown_switches = Vec::new();

    for switch in switches {
        match (kind, switch.name) {
            (MailMergeRecipientFieldKind::AddressBlock, 'c') => {
                if country_inclusion.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "ADDRESSBLOCK field repeats its \\c switch".to_string(),
                    ));
                }
                let argument = switch.argument.ok_or_else(|| {
                    OoxmlError::InvalidFormat(
                        "ADDRESSBLOCK \\c switch requires an argument".to_string(),
                    )
                })?;
                country_inclusion = Some(match argument.as_str() {
                    "0" => AddressBlockCountryInclusion::Omit,
                    "1" => AddressBlockCountryInclusion::Always,
                    "2" => AddressBlockCountryInclusion::UnlessExcluded,
                    _ => {
                        return Err(OoxmlError::InvalidFormat(format!(
                            "ADDRESSBLOCK \\c switch must be 0, 1, or 2, got {argument:?}"
                        )));
                    },
                });
            },
            (MailMergeRecipientFieldKind::AddressBlock, 'd') => {
                if formats_using_recipient_country {
                    return Err(OoxmlError::InvalidFormat(
                        "ADDRESSBLOCK field repeats its \\d switch".to_string(),
                    ));
                }
                if switch.argument.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "ADDRESSBLOCK \\d switch does not accept an argument".to_string(),
                    ));
                }
                formats_using_recipient_country = true;
            },
            (MailMergeRecipientFieldKind::AddressBlock, 'e') => {
                let argument = switch.argument.ok_or_else(|| {
                    OoxmlError::InvalidFormat(
                        "ADDRESSBLOCK \\e switch requires an argument".to_string(),
                    )
                })?;
                excluded_countries.push(argument);
            },
            (_, 'f') => {
                if format_template.is_some() {
                    return Err(OoxmlError::InvalidFormat(format!(
                        "{field_type} field repeats its \\f switch"
                    )));
                }
                format_template = Some(switch.argument.ok_or_else(|| {
                    OoxmlError::InvalidFormat(format!(
                        "{field_type} \\f switch requires an argument"
                    ))
                })?);
            },
            (_, 'l') => {
                if language.is_some() {
                    return Err(OoxmlError::InvalidFormat(format!(
                        "{field_type} field repeats its \\l switch"
                    )));
                }
                language = Some(switch.argument.ok_or_else(|| {
                    OoxmlError::InvalidFormat(format!(
                        "{field_type} \\l switch requires an argument"
                    ))
                })?);
            },
            (MailMergeRecipientFieldKind::GreetingLine, 'c' | 'e') => {
                if greeting_fallback_text.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "GREETINGLINE field repeats its fallback-text switch".to_string(),
                    ));
                }
                greeting_fallback_text = Some(switch.argument.ok_or_else(|| {
                    OoxmlError::InvalidFormat(
                        "GREETINGLINE fallback-text switch requires an argument".to_string(),
                    )
                })?);
            },
            _ => unknown_switches.push(switch),
        }
    }

    Ok(Some((
        kind,
        country_inclusion,
        formats_using_recipient_country,
        excluded_countries,
        format_template,
        language,
        greeting_fallback_text,
        unknown_switches,
    )))
}

/// Parse a `CITATION` instruction while accepting Word's documented leading
/// `\\l` locale switch. Other switches still follow the primary source tag or
/// a preceding `\\m` source tag.
fn parse_citation_operand_and_switches(
    instruction: &str,
) -> Result<Option<(Option<String>, Vec<FieldSwitch>)>> {
    let Some(remainder) = field_instruction_remainder(instruction, "CITATION") else {
        return Ok(None);
    };
    let mut characters = remainder.chars().peekable();
    let mut switches = Vec::new();
    skip_field_whitespace(&mut characters);
    while characters
        .peek()
        .is_some_and(|character| *character == '\\')
    {
        let switch = parse_field_switch_from_characters(&mut characters, "CITATION")?;
        if switch.name != 'l' {
            return Err(OoxmlError::InvalidFormat(
                "CITATION field requires its primary source tag before this switch".to_string(),
            ));
        }
        if switches.len() >= MAX_FIELD_SWITCHES {
            return Err(OoxmlError::InvalidFormat(format!(
                "CITATION field exceeds {MAX_FIELD_SWITCHES} switches"
            )));
        }
        switches.push(switch);
        skip_field_whitespace(&mut characters);
    }
    let operand = match characters.peek().copied() {
        None | Some('\\') => None,
        Some('"') => {
            characters.next();
            Some(parse_field_quoted_argument(&mut characters, "CITATION")?)
        },
        Some(_) => Some(parse_field_unquoted_argument(&mut characters)),
    };
    let remaining = parse_field_switches_from_characters(&mut characters, "CITATION")?;
    if switches.len() + remaining.len() > MAX_FIELD_SWITCHES {
        return Err(OoxmlError::InvalidFormat(format!(
            "CITATION field exceeds {MAX_FIELD_SWITCHES} switches"
        )));
    }
    switches.extend(remaining);
    Ok(Some((operand, switches)))
}

fn parse_field_switches_from_characters(
    characters: &mut std::iter::Peekable<std::str::Chars<'_>>,
    field_type: &str,
) -> Result<Vec<FieldSwitch>> {
    let mut switches = Vec::new();
    loop {
        skip_field_whitespace(characters);
        let Some(character) = characters.next() else {
            break;
        };
        if character != '\\' {
            return Err(OoxmlError::InvalidFormat(format!(
                "{field_type} field contains text outside a field switch"
            )));
        }
        if switches.len() >= MAX_FIELD_SWITCHES {
            return Err(OoxmlError::InvalidFormat(format!(
                "{field_type} field exceeds {MAX_FIELD_SWITCHES} switches"
            )));
        }
        switches.push(parse_field_switch_after_intro(characters, field_type)?);
    }
    Ok(switches)
}

fn parse_field_switch_from_characters(
    characters: &mut std::iter::Peekable<std::str::Chars<'_>>,
    field_type: &str,
) -> Result<FieldSwitch> {
    let introducer = characters.next().ok_or_else(|| {
        OoxmlError::InvalidFormat(format!("{field_type} field ends with a switch introducer"))
    })?;
    if introducer != '\\' {
        return Err(OoxmlError::InvalidFormat(format!(
            "{field_type} field has an invalid switch introducer"
        )));
    }
    parse_field_switch_after_intro(characters, field_type)
}

fn parse_field_switch_after_intro(
    characters: &mut std::iter::Peekable<std::str::Chars<'_>>,
    field_type: &str,
) -> Result<FieldSwitch> {
    let name = characters.next().ok_or_else(|| {
        OoxmlError::InvalidFormat(format!("{field_type} field ends with a switch introducer"))
    })?;
    if name == '\\' || name.is_whitespace() {
        return Err(OoxmlError::InvalidFormat(format!(
            "{field_type} field has an invalid switch name"
        )));
    }
    skip_field_whitespace(characters);
    let argument = match characters.peek().copied() {
        None | Some('\\') => None,
        Some('"') => {
            characters.next();
            Some(parse_field_quoted_argument(characters, field_type)?)
        },
        Some(_) => Some(parse_field_unquoted_argument(characters)),
    };
    Ok(FieldSwitch {
        name: name.to_ascii_lowercase(),
        argument,
    })
}

fn skip_field_whitespace(characters: &mut std::iter::Peekable<std::str::Chars<'_>>) {
    while characters
        .peek()
        .is_some_and(|character| character.is_whitespace())
    {
        characters.next();
    }
}

fn parse_field_quoted_argument(
    characters: &mut std::iter::Peekable<std::str::Chars<'_>>,
    field_type: &str,
) -> Result<String> {
    let mut argument = String::new();
    let mut escaped = false;
    while let Some(character) = characters.next() {
        if escaped {
            if character != '\\' && character != '"' {
                argument.push('\\');
            }
            argument.push(character);
            escaped = false;
            continue;
        }
        match character {
            '\\' => escaped = true,
            '"' => {
                if characters
                    .peek()
                    .is_some_and(|next| !next.is_whitespace() && *next != '\\')
                {
                    return Err(OoxmlError::InvalidFormat(format!(
                        "{field_type} quoted switch argument has trailing text"
                    )));
                }
                return Ok(argument);
            },
            _ => argument.push(character),
        }
    }
    Err(OoxmlError::InvalidFormat(format!(
        "{field_type} field has an unterminated quoted switch argument"
    )))
}

fn parse_field_unquoted_argument(
    characters: &mut std::iter::Peekable<std::str::Chars<'_>>,
) -> String {
    let mut argument = String::new();
    while characters
        .peek()
        .is_some_and(|character| !character.is_whitespace() && *character != '\\')
    {
        argument.push(characters.next().expect("checked field argument character"));
    }
    argument
}

fn parse_toc_level_range(value: &str) -> Result<TableOfContentsLevelRange> {
    let mut levels = value.split('-').map(str::trim);
    let start = levels
        .next()
        .ok_or_else(|| OoxmlError::InvalidFormat("TOC level range is empty".to_string()))?
        .parse::<u8>()
        .map_err(|_| OoxmlError::InvalidFormat("invalid TOC start level".to_string()))?;
    let end = levels
        .next()
        .ok_or_else(|| OoxmlError::InvalidFormat("TOC level range is incomplete".to_string()))?
        .parse::<u8>()
        .map_err(|_| OoxmlError::InvalidFormat("invalid TOC end level".to_string()))?;
    if levels.next().is_some() {
        return Err(OoxmlError::InvalidFormat(
            "TOC level range contains too many separators".to_string(),
        ));
    }
    TableOfContentsLevelRange::new(start, end)
}

struct PendingSimpleField {
    order: usize,
    instruction: String,
    result: String,
    dirty: bool,
    locked: bool,
}

impl PendingSimpleField {
    fn parse(element: &BytesStart<'_>, decoder: Decoder, order: usize) -> Result<Self> {
        let mut instruction = None;
        let mut dirty = false;
        let mut locked = false;
        for attribute in element.attributes() {
            let attribute = attribute.map_err(|error| OoxmlError::Xml(error.to_string()))?;
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(|error| OoxmlError::Xml(error.to_string()))?;
            match attribute.key.local_name().as_ref() {
                b"instr" => instruction = Some(value.into_owned()),
                b"dirty" => dirty = is_on(&value),
                b"fldLock" => locked = is_on(&value),
                _ => {},
            }
        }
        let instruction = instruction.ok_or_else(|| {
            OoxmlError::InvalidFormat("DOCX simple field is missing w:instr".to_string())
        })?;
        Ok(Self {
            order,
            instruction,
            result: String::new(),
            dirty,
            locked,
        })
    }

    fn finish(self) -> Field {
        let result = (!self.result.is_empty()).then_some(self.result);
        Field::with_flags(
            self.instruction.trim().to_string(),
            result,
            self.dirty,
            self.locked,
        )
    }
}

fn is_on(value: &str) -> bool {
    matches!(value, "true" | "1" | "on")
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_field_creation() {
        let field = Field::new("PAGE".to_string(), Some("1".to_string()), false);
        assert_eq!(field.instruction(), "PAGE");
        assert_eq!(field.result(), Some("1"));
        assert!(!field.is_dirty());
        assert_eq!(field.field_type(), "PAGE");
    }

    #[test]
    fn test_field_type_extraction() {
        let field = Field::new("DATE \\@ \"MMMM d, yyyy\"".to_string(), None, false);
        assert_eq!(field.field_type(), "DATE");

        let field = Field::new(
            "REF bookmark1 \\h".to_string(),
            Some("See Section 1".to_string()),
            true,
        );
        assert_eq!(field.field_type(), "REF");
        assert!(field.is_dirty());
    }

    #[test]
    fn extracts_mail_merge_field_names_without_switches() {
        let quoted = Field::new(
            r#"  MERGEFIELD "Full Name" \* MERGEFORMAT "#.to_string(),
            None,
            false,
        );
        assert!(quoted.is_merge_field());
        assert_eq!(quoted.merge_field_name(), Some("Full Name"));

        let unquoted = Field::new("mergefield CustomerId \\b prefix".to_string(), None, false);
        assert!(unquoted.is_merge_field());
        assert_eq!(unquoted.merge_field_name(), Some("CustomerId"));

        let missing = Field::new("MERGEFIELD \\* MERGEFORMAT".to_string(), None, false);
        assert_eq!(missing.merge_field_name(), None);
        let page = Field::new("PAGE".to_string(), None, false);
        assert!(!page.is_merge_field());
        assert_eq!(page.merge_field_name(), None);
    }

    #[test]
    fn parses_inert_merge_fields_without_opening_data_sources() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" MERGEFIELD &quot;Customer Region&quot; \b &quot;Dear &quot; \f &quot;!&quot; \m \v \* MERGEFORMAT " w:dirty="true" w:fldLock="on">
                <w:r><w:t>cached region</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>mergefield CustomerName \b Prefix \f Suffix</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>cached customer</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="MERGEFIELDS CustomerName"><w:r><w:t>not a merge field</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
        let fields = Field::extract_from_document(xml).unwrap();
        assert_eq!(fields.len(), 3);
        assert!(fields[0].is_merge_field());
        assert!(fields[1].is_merge_field());
        assert!(!fields[2].is_merge_field());

        let region = fields[0].merge_field().unwrap().unwrap();
        assert_eq!(region.field_name(), "Customer Region");
        assert_eq!(region.cached_result(), Some("cached region"));
        assert!(region.is_dirty());
        assert!(region.is_locked());
        assert_eq!(region.switches().len(), 5);
        assert_eq!(region.switches()[0].name(), 'b');
        assert_eq!(region.switches()[0].argument(), Some("Dear "));
        assert_eq!(region.switches()[1].name(), 'f');
        assert_eq!(region.switches()[1].argument(), Some("!"));
        assert!(region.has_switch('m'));
        assert!(region.has_switch('v'));
        assert!(region.has_switch('*'));
        assert_eq!(region.switches()[4].argument(), Some("MERGEFORMAT"));

        let customer = fields[1].merge_field().unwrap().unwrap();
        assert_eq!(customer.field_name(), "CustomerName");
        assert_eq!(customer.cached_result(), Some("cached customer"));
        assert!(customer.is_dirty());
        assert!(customer.is_locked());
        assert_eq!(customer.switches()[0].argument(), Some("Prefix"));
        assert_eq!(customer.switches()[1].argument(), Some("Suffix"));

        assert!(fields[2].merge_field().unwrap().is_none());
    }

    #[test]
    fn rejects_invalid_merge_field_semantics() {
        let missing_name = Field::new("MERGEFIELD \\* MERGEFORMAT".to_string(), None, false);
        assert!(missing_name.merge_field().is_err());

        let empty_name = Field::new(r#"MERGEFIELD "" "#.to_string(), None, false);
        assert!(empty_name.merge_field().is_err());

        let unexpected_operand =
            Field::new("MERGEFIELD Customer unexpected".to_string(), None, false);
        assert!(unexpected_operand.merge_field().is_err());
    }

    #[test]
    fn parses_inert_mail_merge_counter_fields_without_merging() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" MERGEREC " w:dirty="true" w:fldLock="on">
                <w:r><w:t>12</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>mergeSEQ</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>3</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="MERGERECORD"><w:r><w:t>not a counter</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
        let fields = Field::extract_from_document(xml).unwrap();
        assert_eq!(fields.len(), 3);
        assert!(fields[0].is_merge_record());
        assert!(fields[0].is_mail_merge_counter());
        assert!(fields[1].is_merge_sequence());
        assert!(fields[1].is_mail_merge_counter());
        assert!(!fields[2].is_mail_merge_counter());

        let record = fields[0].mail_merge_counter().unwrap().unwrap();
        assert_eq!(record.kind(), MailMergeCounterKind::Record);
        assert_eq!(record.cached_result(), Some("12"));
        assert!(record.is_dirty());
        assert!(record.is_locked());

        let sequence = fields[1].mail_merge_counter().unwrap().unwrap();
        assert_eq!(sequence.kind(), MailMergeCounterKind::Sequence);
        assert_eq!(sequence.cached_result(), Some("3"));
        assert!(sequence.is_dirty());
        assert!(sequence.is_locked());

        assert!(fields[2].mail_merge_counter().unwrap().is_none());
    }

    #[test]
    fn rejects_invalid_mail_merge_counter_field_semantics() {
        let record_argument = Field::new("MERGEREC 12".to_string(), None, false);
        assert!(record_argument.mail_merge_counter().is_err());

        let sequence_switch = Field::new("MERGESEQ \\* MERGEFORMAT".to_string(), None, false);
        assert!(sequence_switch.mail_merge_counter().is_err());
    }

    #[test]
    fn parses_inert_mail_merge_next_fields_without_advancing_records() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" NEXT " w:dirty="true" w:fldLock="on">
                <w:r><w:t>cached next</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>next</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>cached complex next</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="NEXTIF Customer = &quot;Ada&quot;"><w:r><w:t>not next</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
        let fields = Field::extract_from_document(xml).unwrap();
        assert_eq!(fields.len(), 3);
        assert!(fields[0].is_mail_merge_next());
        assert!(fields[1].is_mail_merge_next());
        assert!(!fields[2].is_mail_merge_next());

        let simple = fields[0].mail_merge_next().unwrap().unwrap();
        assert_eq!(simple.instruction(), "NEXT");
        assert_eq!(simple.cached_result(), Some("cached next"));
        assert!(simple.is_dirty());
        assert!(simple.is_locked());

        let complex = fields[1].mail_merge_next().unwrap().unwrap();
        assert_eq!(complex.cached_result(), Some("cached complex next"));
        assert!(complex.is_dirty());
        assert!(complex.is_locked());

        assert!(fields[2].mail_merge_next().unwrap().is_none());
    }

    #[test]
    fn rejects_invalid_mail_merge_next_field_semantics() {
        let argument = Field::new("NEXT 12".to_string(), None, false);
        assert!(argument.mail_merge_next().is_err());

        let switch = Field::new("NEXT \\* MERGEFORMAT".to_string(), None, false);
        assert!(switch.mail_merge_next().is_err());
    }

    #[test]
    fn parses_inert_conditional_mail_merge_controls_without_merging() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" NEXTIF Customer = &quot;Ada&quot; " w:dirty="true" w:fldLock="on">
                <w:r><w:t>cached nextif</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>skipif MERGEFIELD Order &lt; 100</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>cached skipif</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="NEXTIFF Customer = &quot;Ada&quot;"><w:r><w:t>not conditional</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
        let fields = Field::extract_from_document(xml).unwrap();
        assert_eq!(fields.len(), 3);
        assert!(fields[0].is_mail_merge_next_if());
        assert!(fields[0].is_mail_merge_conditional_control());
        assert!(fields[1].is_mail_merge_skip_if());
        assert!(fields[1].is_mail_merge_conditional_control());
        assert!(!fields[2].is_mail_merge_conditional_control());

        let next_if = fields[0].mail_merge_conditional_control().unwrap().unwrap();
        assert_eq!(next_if.kind(), MailMergeConditionalControlKind::NextIf);
        assert_eq!(next_if.comparison(), r#"Customer = "Ada""#);
        assert_eq!(next_if.cached_result(), Some("cached nextif"));
        assert!(next_if.is_dirty());
        assert!(next_if.is_locked());

        let skip_if = fields[1].mail_merge_conditional_control().unwrap().unwrap();
        assert_eq!(skip_if.kind(), MailMergeConditionalControlKind::SkipIf);
        assert_eq!(skip_if.comparison(), "MERGEFIELD Order < 100");
        assert_eq!(skip_if.cached_result(), Some("cached skipif"));
        assert!(skip_if.is_dirty());
        assert!(skip_if.is_locked());
    }

    #[test]
    fn rejects_conditional_mail_merge_controls_without_comparisons() {
        let next_if = Field::new("NEXTIF".to_string(), None, false);
        assert!(next_if.mail_merge_conditional_control().is_err());

        let skip_if = Field::new("SKIPIF   ".to_string(), None, false);
        assert!(skip_if.mail_merge_conditional_control().is_err());
    }

    #[test]
    fn parses_inert_if_fields_without_evaluation() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" IF &quot;A&quot; = &quot;A&quot; &quot;yes&quot; &quot;no&quot; " w:dirty="true" w:fldLock="on">
                <w:r><w:t>yes</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>if MERGEFIELD Amount &gt; 100 "discount" "standard"</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>discount</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="IFF 1 = 1 &quot;yes&quot; &quot;no&quot;"><w:r><w:t>not if</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
        let fields = Field::extract_from_document(xml).unwrap();
        assert_eq!(fields.len(), 3);
        assert!(fields[0].is_if_field());
        assert!(fields[1].is_if_field());
        assert!(!fields[2].is_if_field());

        let simple = fields[0].if_field().unwrap().unwrap();
        assert_eq!(simple.expression(), r#""A" = "A" "yes" "no""#);
        assert_eq!(simple.cached_result(), Some("yes"));
        assert!(simple.is_dirty());
        assert!(simple.is_locked());

        let complex = fields[1].if_field().unwrap().unwrap();
        assert_eq!(
            complex.expression(),
            r#"MERGEFIELD Amount > 100 "discount" "standard""#
        );
        assert_eq!(complex.cached_result(), Some("discount"));
        assert!(complex.is_dirty());
        assert!(complex.is_locked());
    }

    #[test]
    fn rejects_if_fields_without_expressions() {
        let missing = Field::new("IF".to_string(), None, false);
        assert!(missing.if_field().is_err());
    }

    #[test]
    fn parses_inert_compare_fields_without_evaluation() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" COMPARE &quot;CustomerNumber&quot; &gt;= 4 " w:dirty="true" w:fldLock="on">
                <w:r><w:t>1</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>compare MERGEFIELD CustomerRating &lt;= 9</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>0</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="COMPARES Customer = 1"><w:r><w:t>not a comparison</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
        let fields = Field::extract_from_document(xml).unwrap();
        assert_eq!(fields.len(), 3);
        assert!(fields[0].is_compare_field());
        assert!(fields[1].is_compare_field());
        assert!(!fields[2].is_compare_field());

        let number = fields[0].compare_field().unwrap().unwrap();
        assert_eq!(number.comparison(), r#""CustomerNumber" >= 4"#);
        assert_eq!(number.cached_result(), Some("1"));
        assert!(number.is_dirty());
        assert!(number.is_locked());

        let rating = fields[1].compare_field().unwrap().unwrap();
        assert_eq!(rating.comparison(), "MERGEFIELD CustomerRating <= 9");
        assert_eq!(rating.cached_result(), Some("0"));
        assert!(rating.is_dirty());
        assert!(rating.is_locked());
        assert!(fields[2].compare_field().unwrap().is_none());
    }

    #[test]
    fn rejects_compare_fields_without_comparisons() {
        let missing = Field::new("COMPARE".to_string(), None, false);
        assert!(missing.compare_field().is_err());
    }

    #[test]
    fn parses_inert_set_fields_without_evaluation_or_state_changes() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" SET RecipientName &quot;North America&quot; \* MERGEFORMAT" w:dirty="true" w:fldLock="on">
                <w:r><w:t>cached recipient</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>set Total =SUM(ABOVE) + 1</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>125</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="SETTINGS Value"><w:r><w:t>not set</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
        let fields = Field::extract_from_document(xml).unwrap();
        assert_eq!(fields.len(), 3);
        assert!(fields[0].is_set_field());
        assert!(fields[1].is_set_field());
        assert!(!fields[2].is_set_field());

        let recipient = fields[0].set_field().unwrap().unwrap();
        assert_eq!(recipient.target_name(), "RecipientName");
        assert_eq!(recipient.expression(), r#""North America" \* MERGEFORMAT"#);
        assert_eq!(recipient.cached_result(), Some("cached recipient"));
        assert!(recipient.is_dirty());
        assert!(recipient.is_locked());

        let total = fields[1].set_field().unwrap().unwrap();
        assert_eq!(total.target_name(), "Total");
        assert_eq!(total.expression(), "=SUM(ABOVE) + 1");
        assert_eq!(total.cached_result(), Some("125"));
        assert!(total.is_dirty());
        assert!(total.is_locked());
        assert!(fields[2].set_field().unwrap().is_none());
    }

    #[test]
    fn rejects_invalid_set_fields_without_evaluating_them() {
        for instruction in ["SET", "SET \"\" value", "SET Target", "SET Target   "] {
            let field = Field::new(instruction.to_string(), None, false);
            assert!(field.set_field().is_err(), "{instruction}");
        }

        let too_long = Field::new(
            format!("SET Target {}", "x".repeat(MAX_SET_FIELD_INSTRUCTION_BYTES)),
            None,
            false,
        );
        assert!(too_long.set_field().is_err());
    }

    #[test]
    fn parses_inert_sequence_fields_without_bookmark_lookup_or_numbering() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" SEQ Figure FigureChapter \r 3 \* ARABIC " w:dirty="true" w:fldLock="on">
                <w:r><w:t>3</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>seq Table \s 1 \* ROMAN</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>I</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="SEQUENCE Figure"><w:r><w:t>not a sequence</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
        let fields = Field::extract_from_document(xml).unwrap();
        assert_eq!(fields.len(), 3);
        assert!(fields[0].is_sequence_field());
        assert!(fields[1].is_sequence_field());
        assert!(!fields[2].is_sequence_field());

        let figure = fields[0].sequence_field().unwrap().unwrap();
        assert_eq!(figure.identifier(), "Figure");
        assert_eq!(figure.bookmark(), Some("FigureChapter"));
        assert_eq!(figure.tail(), r"\r 3 \* ARABIC");
        assert_eq!(figure.cached_result(), Some("3"));
        assert!(figure.is_dirty());
        assert!(figure.is_locked());

        let table = fields[1].sequence_field().unwrap().unwrap();
        assert_eq!(table.identifier(), "Table");
        assert_eq!(table.bookmark(), None);
        assert_eq!(table.tail(), r"\s 1 \* ROMAN");
        assert_eq!(table.cached_result(), Some("I"));
        assert!(table.is_dirty());
        assert!(table.is_locked());
        assert!(fields[2].sequence_field().unwrap().is_none());
    }

    #[test]
    fn rejects_invalid_sequence_fields_without_numbering() {
        for instruction in ["SEQ", r#"SEQ ""#, r#"SEQ Figure ""#] {
            let field = Field::new(instruction.to_string(), None, false);
            assert!(field.sequence_field().is_err(), "{instruction}");
        }

        let too_long = Field::new(
            format!(
                "SEQ Figure {}",
                "x".repeat(MAX_SEQUENCE_FIELD_INSTRUCTION_BYTES)
            ),
            None,
            false,
        );
        assert!(too_long.sequence_field().is_err());
    }

    #[test]
    fn parses_inert_formula_fields_without_evaluation() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" =SUM(ABOVE) \* MERGEFORMAT " w:dirty="true" w:fldLock="on">
                <w:r><w:t>42</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>= IF(1 = 1, &quot;yes&quot;, &quot;no&quot;)</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>yes</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="EQUAL 1 + 1"><w:r><w:t>not a formula field</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
        let fields = Field::extract_from_document(xml).unwrap();
        assert_eq!(fields.len(), 3);
        assert!(fields[0].is_formula_field());
        assert!(fields[1].is_formula_field());
        assert!(!fields[2].is_formula_field());

        let total = fields[0].formula_field().unwrap().unwrap();
        assert_eq!(total.formula(), r"SUM(ABOVE) \* MERGEFORMAT");
        assert_eq!(total.cached_result(), Some("42"));
        assert!(total.is_dirty());
        assert!(total.is_locked());

        let conditional = fields[1].formula_field().unwrap().unwrap();
        assert_eq!(conditional.formula(), r#"IF(1 = 1, "yes", "no")"#);
        assert_eq!(conditional.cached_result(), Some("yes"));
        assert!(conditional.is_dirty());
        assert!(conditional.is_locked());
        assert!(fields[2].formula_field().unwrap().is_none());
    }

    #[test]
    fn rejects_invalid_formula_fields_without_evaluating_them() {
        let missing = Field::new("=".to_string(), None, false);
        assert!(missing.is_formula_field());
        assert!(missing.formula_field().is_err());

        let too_long = Field::new(
            format!("={}", "x".repeat(MAX_FORMULA_FIELD_INSTRUCTION_BYTES)),
            None,
            false,
        );
        assert!(too_long.formula_field().is_err());
    }

    #[test]
    fn parses_inert_quote_fields_without_inserting_or_transforming_text() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" QUOTE &quot;Stored literal&quot; \* MERGEFORMAT \# &quot;000&quot; " w:dirty="true" w:fldLock="on">
                <w:r><w:t>cached literal</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>quote "Complex literal" \@ "MMMM"</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>cached complex literal</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="QUOTEY &quot;not a quote field&quot;"><w:r><w:t>not a quote field</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
        let fields = Field::extract_from_document(xml).unwrap();
        assert_eq!(fields.len(), 3);
        assert!(fields[0].is_quote_field());
        assert!(fields[1].is_quote_field());
        assert!(!fields[2].is_quote_field());

        let literal = fields[0].quote_field().unwrap().unwrap();
        assert_eq!(literal.text(), "Stored literal");
        assert_eq!(literal.cached_result(), Some("cached literal"));
        assert!(literal.is_dirty());
        assert!(literal.is_locked());
        assert_eq!(literal.switches().len(), 2);
        assert_eq!(literal.switches()[0].name(), '*');
        assert_eq!(literal.switches()[0].argument(), Some("MERGEFORMAT"));
        assert_eq!(literal.switches()[1].name(), '#');
        assert_eq!(literal.switches()[1].argument(), Some("000"));
        assert!(literal.has_switch('*'));

        let complex = fields[1].quote_field().unwrap().unwrap();
        assert_eq!(complex.text(), "Complex literal");
        assert_eq!(complex.cached_result(), Some("cached complex literal"));
        assert!(complex.is_dirty());
        assert!(complex.is_locked());
        assert_eq!(complex.switches()[0].name(), '@');
        assert_eq!(complex.switches()[0].argument(), Some("MMMM"));
        assert!(fields[2].quote_field().unwrap().is_none());
    }

    #[test]
    fn rejects_invalid_quote_fields_without_inserting_or_transforming_text() {
        for instruction in [
            "QUOTE",
            "QUOTE \\\\* MERGEFORMAT",
            r#"QUOTE "literal" unexpected"#,
            r#"QUOTE "unterminated"#,
        ] {
            let field = Field::new(instruction.to_string(), None, false);
            assert!(field.quote_field().is_err(), "{instruction}");
        }

        let too_long = Field::new(
            format!(
                "QUOTE \"{}\"",
                "x".repeat(MAX_QUOTE_FIELD_INSTRUCTION_BYTES)
            ),
            None,
            false,
        );
        assert!(too_long.quote_field().is_err());
    }

    #[test]
    fn parses_inert_symbol_fields_without_mapping_codes_or_inserting_glyphs() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" SYMBOL 0xA9 \f &quot;Symbol&quot; \s 12 \u " w:dirty="true" w:fldLock="on">
                <w:r><w:t>cached copyright</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>symbol 163 \a \h \j</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>cached pound</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="SYMBOLS 163"><w:r><w:t>not a symbol field</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
        let fields = Field::extract_from_document(xml).unwrap();
        assert_eq!(fields.len(), 3);
        assert!(fields[0].is_symbol_field());
        assert!(fields[1].is_symbol_field());
        assert!(!fields[2].is_symbol_field());

        let copyright = fields[0].symbol_field().unwrap().unwrap();
        assert_eq!(copyright.character_argument(), "0xA9");
        assert_eq!(copyright.cached_result(), Some("cached copyright"));
        assert!(copyright.is_dirty());
        assert!(copyright.is_locked());
        assert_eq!(copyright.switches().len(), 3);
        assert_eq!(copyright.switches()[0].name(), 'f');
        assert_eq!(copyright.switches()[0].argument(), Some("Symbol"));
        assert_eq!(copyright.switches()[1].name(), 's');
        assert_eq!(copyright.switches()[1].argument(), Some("12"));
        assert_eq!(copyright.switches()[2].name(), 'u');
        assert_eq!(copyright.switches()[2].argument(), None);
        assert!(copyright.has_switch('f'));

        let pound = fields[1].symbol_field().unwrap().unwrap();
        assert_eq!(pound.character_argument(), "163");
        assert_eq!(pound.cached_result(), Some("cached pound"));
        assert!(pound.is_dirty());
        assert!(pound.is_locked());
        assert_eq!(pound.switches().len(), 3);
        assert_eq!(pound.switches()[0].name(), 'a');
        assert_eq!(pound.switches()[1].name(), 'h');
        assert_eq!(pound.switches()[2].name(), 'j');
        assert!(fields[2].symbol_field().unwrap().is_none());
    }

    #[test]
    fn rejects_invalid_symbol_fields_without_mapping_codes_or_inserting_glyphs() {
        for instruction in [
            "SYMBOL",
            "SYMBOL \\f \"Symbol\"",
            "SYMBOL 0xA9 unexpected",
            "SYMBOL 0xA9 \\f \"unterminated",
        ] {
            let field = Field::new(instruction.to_string(), None, false);
            assert!(field.symbol_field().is_err(), "{instruction}");
        }

        let too_long = Field::new(
            format!(
                "SYMBOL {}",
                "x".repeat(MAX_SYMBOL_FIELD_INSTRUCTION_BYTES)
            ),
            None,
            false,
        );
        assert!(too_long.symbol_field().is_err());
    }

    #[test]
    fn parses_inert_style_reference_fields_without_style_or_layout_resolution() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" STYLEREF &quot;Heading 1&quot; \l \n \p \r \t \w \* MERGEFORMAT \q opaque " w:dirty="true" w:fldLock="on">
                <w:r><w:t>Cached heading</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>styleref Title \n</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>1</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="STYLEREFS Heading 1"><w:r><w:t>not a style reference</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
        let fields = Field::extract_from_document(xml).unwrap();
        assert_eq!(fields.len(), 3);
        assert!(fields[0].is_style_reference_field());
        assert!(fields[1].is_style_reference_field());
        assert!(!fields[2].is_style_reference_field());

        let heading = fields[0].style_reference_field().unwrap().unwrap();
        assert_eq!(heading.style_name(), "Heading 1");
        assert_eq!(
            heading.options(),
            &[
                StyleReferenceFieldOption::FollowingText,
                StyleReferenceFieldOption::ParagraphNumber,
                StyleReferenceFieldOption::RelativePosition,
                StyleReferenceFieldOption::ParagraphNumberRelativeContext,
                StyleReferenceFieldOption::SuppressNonNumberText,
                StyleReferenceFieldOption::ParagraphNumberFullContext,
            ]
        );
        assert_eq!(heading.unknown_switches().len(), 2);
        assert_eq!(heading.unknown_switches()[0].name(), '*');
        assert_eq!(heading.unknown_switches()[0].argument(), Some("MERGEFORMAT"));
        assert_eq!(heading.unknown_switches()[1].name(), 'q');
        assert_eq!(heading.unknown_switches()[1].argument(), Some("opaque"));
        assert_eq!(heading.cached_result(), Some("Cached heading"));
        assert!(heading.is_dirty());
        assert!(heading.is_locked());

        let title = fields[1].style_reference_field().unwrap().unwrap();
        assert_eq!(title.style_name(), "Title");
        assert_eq!(
            title.options(),
            &[StyleReferenceFieldOption::ParagraphNumber]
        );
        assert_eq!(title.cached_result(), Some("1"));
        assert!(title.is_dirty());
        assert!(title.is_locked());
        assert!(fields[2].style_reference_field().unwrap().is_none());
    }

    #[test]
    fn rejects_invalid_style_reference_fields_without_style_or_layout_resolution() {
        for instruction in [
            "STYLEREF",
            r#"STYLEREF ""#,
            r#"STYLEREF Heading \l unexpected"#,
            "STYLEREF Heading unexpected",
            r#"STYLEREF Heading \"#,
        ] {
            let field = Field::new(instruction.to_string(), None, false);
            assert!(field.style_reference_field().is_err(), "{instruction}");
        }

        let too_long = Field::new(
            format!(
                "STYLEREF Heading {}",
                "x".repeat(MAX_STYLE_REFERENCE_FIELD_INSTRUCTION_BYTES)
            ),
            None,
            false,
        );
        assert!(too_long.style_reference_field().is_err());
    }

    #[test]
    fn parses_inert_prompt_fields_without_displaying_prompts() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" ASK AskResponse &quot;What is your first name?&quot; \d &quot;&quot; \o " w:dirty="true" w:fldLock="on">
                <w:r><w:t>cached ask response</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>fillin "Enter appointment time" \d "09:00"</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>10:30</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="ASKER Answer &quot;not a prompt field&quot;"><w:r><w:t>not ask</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
        let fields = Field::extract_from_document(xml).unwrap();
        assert_eq!(fields.len(), 3);
        assert!(fields[0].is_ask_field());
        assert!(fields[0].is_prompt_field());
        assert!(fields[1].is_fill_in_field());
        assert!(fields[1].is_prompt_field());
        assert!(!fields[2].is_prompt_field());

        let ask = fields[0].prompt_field().unwrap().unwrap();
        assert_eq!(ask.kind(), PromptFieldKind::Ask);
        assert_eq!(ask.bookmark(), Some("AskResponse"));
        assert_eq!(ask.prompt(), Some("What is your first name?"));
        assert_eq!(ask.default_response(), Some(""));
        assert!(ask.prompts_once_per_mail_merge());
        assert_eq!(ask.cached_result(), Some("cached ask response"));
        assert!(ask.is_dirty());
        assert!(ask.is_locked());

        let fill_in = fields[1].prompt_field().unwrap().unwrap();
        assert_eq!(fill_in.kind(), PromptFieldKind::FillIn);
        assert_eq!(fill_in.bookmark(), None);
        assert_eq!(fill_in.prompt(), Some("Enter appointment time"));
        assert_eq!(fill_in.default_response(), Some("09:00"));
        assert!(!fill_in.prompts_once_per_mail_merge());
        assert_eq!(fill_in.cached_result(), Some("10:30"));
        assert!(fill_in.is_dirty());
        assert!(fill_in.is_locked());

        let default_only =
            Field::new(r#"FILLIN \d "recent response" \o"#.to_string(), None, false);
        let default_only = default_only.prompt_field().unwrap().unwrap();
        assert_eq!(default_only.kind(), PromptFieldKind::FillIn);
        assert_eq!(default_only.bookmark(), None);
        assert_eq!(default_only.prompt(), None);
        assert_eq!(default_only.default_response(), Some("recent response"));
        assert!(default_only.prompts_once_per_mail_merge());

        assert!(fields[2].prompt_field().unwrap().is_none());
    }

    #[test]
    fn rejects_malformed_prompt_field_metadata() {
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
            let field = Field::new(instruction.to_string(), None, false);
            assert!(field.prompt_field().is_err(), "{instruction}");
        }
    }

    #[test]
    fn parses_inert_mail_merge_recipient_fields_without_merging() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" ADDRESSBLOCK \c 2 \d \e &quot;United States&quot; \e Canada \f &quot;&lt;&lt;_FIRST0_&gt;&gt; &lt;&lt;_LAST0_&gt;&gt;&quot; \l 1033 \* MERGEFORMAT " w:dirty="true" w:fldLock="on">
                <w:r><w:t>cached address</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>greetingline \f "Dear &lt;&lt;_FIRST0_&gt;&gt;," \e "To Whom It May Concern" \l en-US</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>Dear Ada,</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="ADDRESSBLOCKING \c 1"><w:r><w:t>not an address block</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
        let fields = Field::extract_from_document(xml).unwrap();
        assert_eq!(fields.len(), 3);
        assert!(fields[0].is_address_block());
        assert!(fields[0].is_mail_merge_recipient_field());
        assert!(fields[1].is_greeting_line());
        assert!(fields[1].is_mail_merge_recipient_field());
        assert!(!fields[2].is_mail_merge_recipient_field());

        let address = fields[0].mail_merge_recipient_field().unwrap().unwrap();
        assert_eq!(address.kind(), MailMergeRecipientFieldKind::AddressBlock);
        assert_eq!(
            address.country_inclusion(),
            Some(AddressBlockCountryInclusion::UnlessExcluded)
        );
        assert!(address.formats_using_recipient_country());
        let excluded = address
            .excluded_countries()
            .iter()
            .map(String::as_str)
            .collect::<Vec<_>>();
        assert_eq!(excluded, vec!["United States", "Canada"]);
        assert_eq!(address.format_template(), Some("<<_FIRST0_>> <<_LAST0_>>"));
        assert_eq!(address.language(), Some("1033"));
        assert_eq!(address.greeting_fallback_text(), None);
        assert_eq!(address.unknown_switches().len(), 1);
        assert_eq!(address.unknown_switches()[0].name(), '*');
        assert_eq!(
            address.unknown_switches()[0].argument(),
            Some("MERGEFORMAT")
        );
        assert_eq!(address.cached_result(), Some("cached address"));
        assert!(address.is_dirty());
        assert!(address.is_locked());

        let greeting = fields[1].mail_merge_recipient_field().unwrap().unwrap();
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
        assert_eq!(greeting.cached_result(), Some("Dear Ada,"));
        assert!(greeting.is_dirty());
        assert!(greeting.is_locked());

        assert!(fields[2].mail_merge_recipient_field().unwrap().is_none());
    }

    #[test]
    fn rejects_malformed_mail_merge_recipient_field_metadata() {
        for instruction in [
            "ADDRESSBLOCK text",
            "ADDRESSBLOCK \\c",
            "ADDRESSBLOCK \\c 3",
            "ADDRESSBLOCK \\d 1",
            "ADDRESSBLOCK \\d \\d",
            "ADDRESSBLOCK \\f",
            "GREETINGLINE \\f \"Dear\" \\f \"Hello\"",
            "GREETINGLINE \\l",
            "GREETINGLINE \\c \"First\" \\e \"Second\"",
        ] {
            let field = Field::new(instruction.to_string(), None, false);
            assert!(field.mail_merge_recipient_field().is_err(), "{instruction}");
        }
    }

    #[test]
    fn extracts_decoded_field_instruction_and_result() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r>
            <w:r><w:instrText xml:space="preserve"> IF &quot;A&amp;B&quot; = &quot;A&amp;B&quot; </w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate"/></w:r>
            <w:r><w:t xml:space="preserve"> Yes &amp; no </w:t><w:tab/><w:br/></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
        </w:p></w:body></w:document>"#;
        let fields = Field::extract_from_document(xml).unwrap();
        assert_eq!(fields.len(), 1);
        assert_eq!(fields[0].instruction(), r#"IF "A&B" = "A&B""#);
        assert_eq!(fields[0].result(), Some(" Yes & no \t\n"));
        assert!(fields[0].is_dirty());
    }

    #[test]
    fn extracts_simple_fields_in_source_order_with_flags_and_nested_results() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" MERGEFIELD &quot;Full Name&quot; " w:dirty="on" w:fldLock="1">
                <w:r><w:t xml:space="preserve"> Ada &amp; </w:t></w:r>
                <w:fldSimple w:instr=" PAGE "><w:r><w:t>7</w:t></w:r></w:fldSimple>
                <w:r><w:t><![CDATA[ <Lovelace> ]]></w:t><w:tab/><w:br/><w:noBreakHyphen/><w:softHyphen/></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText> DATE </w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>Today</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr=" NUMPAGES "/>
        </w:p></w:body></w:document>"#;

        let fields = Field::extract_from_document(xml).unwrap();
        assert_eq!(fields.len(), 4);
        assert_eq!(fields[0].instruction(), r#"MERGEFIELD "Full Name""#);
        assert_eq!(
            fields[0].result(),
            Some(" Ada & 7 <Lovelace> \t\n‑\u{00ad}")
        );
        assert!(fields[0].is_dirty());
        assert!(fields[0].is_locked());
        assert_eq!(fields[1].instruction(), "PAGE");
        assert_eq!(fields[1].result(), Some("7"));
        assert_eq!(fields[2].instruction(), "DATE");
        assert_eq!(fields[2].result(), Some("Today"));
        assert!(fields[2].is_dirty());
        assert!(fields[2].is_locked());
        assert_eq!(fields[3].instruction(), "NUMPAGES");
        assert_eq!(fields[3].result(), None);
    }

    #[test]
    fn parses_inert_link_field_metadata_without_activating_sources() {
        let field = Field::new(
            r#"LINK Excel.Sheet.8 "C:\\no-contact\\source.xlsx" "Sheet1!R1C1:R4C4" \a \f 4 \p \d \* MERGEFORMAT"#
                .to_string(),
            Some("cached LINK result".to_string()),
            true,
        );
        assert!(field.is_link());
        let link = field.link().unwrap().unwrap();
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
        assert!(!link.is_locked());
        assert_eq!(link.switches().len(), 5);
        assert_eq!(link.switches()[0].name(), 'a');
        assert_eq!(link.switches()[1].argument(), Some("4"));
        assert_eq!(link.switches()[4].name(), '*');
        assert_eq!(link.switches()[4].argument(), Some("MERGEFORMAT"));

        let multiple_formatting = Field::new(
            r"LINK Word.Document.8 source \f 0 \f 2 \t".to_string(),
            None,
            false,
        );
        let multiple_formatting = multiple_formatting.link().unwrap().unwrap();
        assert_eq!(
            multiple_formatting.formatting_modes(),
            &[LinkFormatting::Source, LinkFormatting::Destination]
        );
        assert_eq!(
            multiple_formatting.effective_result_option(),
            Some(LinkResultOption::Text)
        );

        let unsupported = Field::new(r"LINK Package source \f 1".to_string(), None, false);
        assert_eq!(
            unsupported.link().unwrap().unwrap().formatting_modes(),
            &[LinkFormatting::Unsupported(1)]
        );

        let repeated_updates =
            Field::new(r"LINK Excel.Sheet.8 source \a \a".to_string(), None, false);
        assert!(
            repeated_updates
                .link()
                .unwrap()
                .unwrap()
                .requests_automatic_updates()
        );

        let not_link = Field::new("LINKAGE Excel.Sheet.8 source".to_string(), None, false);
        assert!(!not_link.is_link());
        assert!(not_link.link().unwrap().is_none());
        assert!(Field::new("LINK".to_string(), None, false).link().is_err());
        assert!(
            Field::new(
                r"LINK Excel.Sheet.8 source \f invalid".to_string(),
                None,
                false,
            )
            .link()
            .is_err()
        );
        assert!(
            Field::new(
                r"LINK Excel.Sheet.8 source \p unexpected".to_string(),
                None,
                false,
            )
            .link()
            .is_err()
        );
    }

    #[test]
    fn parses_inert_dde_fields_without_starting_conversations() {
        let field = Field::new(
            r#"DDE Excel "C:\\no-contact\\source.xlsx" "Sheet1!R1C1:R4C4" \a \p \* MERGEFORMAT"#
                .to_string(),
            Some("cached DDE result".to_string()),
            true,
        );
        assert!(field.is_dde());
        assert!(!field.is_dde_auto());
        let dde = field.dde_link().unwrap().unwrap();
        assert_eq!(
            dde.instruction(),
            r#"DDE Excel "C:\\no-contact\\source.xlsx" "Sheet1!R1C1:R4C4" \a \p \* MERGEFORMAT"#
        );
        assert_eq!(dde.kind(), DdeFieldKind::Dde);
        assert_eq!(dde.application(), "Excel");
        assert_eq!(dde.source(), r"C:\no-contact\source.xlsx");
        assert_eq!(dde.item(), Some("Sheet1!R1C1:R4C4"));
        assert!(dde.requests_automatic_updates());
        assert_eq!(dde.representation(), Some(DdeRepresentation::Picture));
        assert!(!dde.omits_graphic_data());
        assert_eq!(dde.cached_result(), Some("cached DDE result"));
        assert!(dde.is_dirty());
        assert!(!dde.is_locked());
        assert_eq!(dde.switches().len(), 3);
        assert_eq!(dde.switches()[0].name(), 'a');
        assert_eq!(dde.switches()[2].name(), '*');
        assert_eq!(dde.switches()[2].argument(), Some("MERGEFORMAT"));

        let automatic = Field::new(
            r#"DDEAUTO Excel "missing.xlsx" "Sheet1!A1" \t"#.to_string(),
            None,
            false,
        );
        assert!(!automatic.is_dde());
        assert!(automatic.is_dde_auto());
        let automatic = automatic.dde_link().unwrap().unwrap();
        assert_eq!(automatic.kind(), DdeFieldKind::DdeAuto);
        assert!(automatic.requests_automatic_updates());
        assert_eq!(automatic.representation(), Some(DdeRepresentation::Text));

        let omit_graphics = Field::new(r"DDE Excel source \a \d".to_string(), None, false)
            .dde_link()
            .unwrap()
            .unwrap();
        assert!(omit_graphics.requests_automatic_updates());
        assert!(omit_graphics.omits_graphic_data());
        assert_eq!(omit_graphics.representation(), None);

        assert!(
            Field::new("DDE".to_string(), None, false)
                .dde_link()
                .is_err()
        );
        assert!(
            Field::new(r"DDE Excel \p".to_string(), None, false)
                .dde_link()
                .is_err()
        );
        assert!(
            Field::new(r"DDE Excel source \p unexpected".to_string(), None, false)
                .dde_link()
                .is_err()
        );
        assert!(
            Field::new(r"DDE Excel source \p \t".to_string(), None, false)
                .dde_link()
                .is_err()
        );
        assert!(
            Field::new(r"DDEAUTO Excel source \p \t".to_string(), None, false)
                .dde_link()
                .is_err()
        );
        assert!(
            Field::new(r"DDEAUTO Excel source \a".to_string(), None, false)
                .dde_link()
                .is_err()
        );
        assert!(
            Field::new(r"DDE Excel source \a \a".to_string(), None, false)
                .dde_link()
                .is_err()
        );
        let not_dde = Field::new("DDEAUTOMATED Excel source".to_string(), None, false);
        assert!(!not_dde.is_dde());
        assert!(!not_dde.is_dde_auto());
        assert!(not_dde.dde_link().unwrap().is_none());
    }

    #[test]
    fn parses_inert_referenced_document_fields_without_opening_sources() {
        let field = Field::with_flags(
            r#"RD "C:\\Manual\\Chapters\\Chapter 1.docx" \p \* MERGEFORMAT"#.to_string(),
            Some("cached RD result".to_string()),
            true,
            true,
        );
        assert!(field.is_referenced_document());
        let reference = field.referenced_document().unwrap().unwrap();
        assert_eq!(reference.source(), r"C:\Manual\Chapters\Chapter 1.docx");
        assert!(reference.uses_relative_path());
        assert_eq!(reference.cached_result(), Some("cached RD result"));
        assert!(reference.is_dirty());
        assert!(reference.is_locked());
        assert_eq!(reference.switches().len(), 2);
        assert_eq!(reference.switches()[0].name(), 'p');
        assert_eq!(reference.switches()[1].name(), '*');
        assert_eq!(reference.switches()[1].argument(), Some("MERGEFORMAT"));

        let absolute = Field::new(
            r#"RD "file:///no-contact/appendix.docx""#.to_string(),
            None,
            false,
        );
        let absolute = absolute.referenced_document().unwrap().unwrap();
        assert_eq!(absolute.source(), "file:///no-contact/appendix.docx");
        assert!(!absolute.uses_relative_path());

        assert!(
            Field::new("RD".to_string(), None, false)
                .referenced_document()
                .is_err()
        );
        assert!(
            Field::new(r#"RD "chapter.docx" \p relative"#.to_string(), None, false)
                .referenced_document()
                .is_err()
        );
        assert!(
            Field::new(r#"RD "chapter.docx" \p \p"#.to_string(), None, false)
                .referenced_document()
                .is_err()
        );
        let not_rd = Field::new(r#"RDX "chapter.docx""#.to_string(), None, false);
        assert!(!not_rd.is_referenced_document());
        assert!(not_rd.referenced_document().unwrap().is_none());
    }

    #[test]
    fn parses_inert_external_include_fields_without_resolving_sources() {
        let text_field = Field::new(
            r#"INCLUDETEXT "file:///C:/no-contact/source.xml" Summary \! \c Word8 \e utf-8 \m application/xml \n "xmlns:a=\"resume-schema\"" \t "file:///C:/display.xsl" \x a:Resume/a:Name \* MERGEFORMAT"#
                .to_string(),
            Some("cached included text".to_string()),
            true,
        );
        assert!(text_field.is_include_text());
        assert!(!text_field.is_include_picture());
        let text = text_field.external_include().unwrap().unwrap();
        assert_eq!(text.kind(), IncludeFieldKind::Text);
        assert_eq!(text.source(), "file:///C:/no-contact/source.xml");
        assert_eq!(text.bookmark(), Some("Summary"));
        assert!(text.suppresses_nested_field_updates());
        assert!(!text.omits_picture_data());
        assert_eq!(
            text.options(),
            &[
                ExternalIncludeOption::Converter("Word8".to_string()),
                ExternalIncludeOption::Encoding("utf-8".to_string()),
                ExternalIncludeOption::MimeType("application/xml".to_string()),
                ExternalIncludeOption::NamespaceMapping("xmlns:a=\"resume-schema\"".to_string()),
                ExternalIncludeOption::Xslt("file:///C:/display.xsl".to_string()),
                ExternalIncludeOption::XPath("a:Resume/a:Name".to_string()),
            ]
        );
        assert_eq!(text.cached_result(), Some("cached included text"));
        assert!(text.is_dirty());
        assert!(!text.is_locked());
        assert_eq!(text.switches().len(), 8);
        assert_eq!(text.switches()[0].name(), '!');
        assert_eq!(text.switches()[7].name(), '*');

        let picture_field = Field::new(
            r#"INCLUDEPICTURE "file:///C:/no-contact/picture.gif" \c Pictim32 \d \* MERGEFORMAT"#
                .to_string(),
            Some("cached picture".to_string()),
            false,
        );
        assert!(!picture_field.is_include_text());
        assert!(picture_field.is_include_picture());
        let picture = picture_field.external_include().unwrap().unwrap();
        assert_eq!(picture.kind(), IncludeFieldKind::Picture);
        assert_eq!(picture.source(), "file:///C:/no-contact/picture.gif");
        assert_eq!(picture.bookmark(), None);
        assert!(!picture.suppresses_nested_field_updates());
        assert!(picture.omits_picture_data());
        assert_eq!(
            picture.options(),
            &[ExternalIncludeOption::Converter("Pictim32".to_string())]
        );
        assert_eq!(picture.cached_result(), Some("cached picture"));
        assert_eq!(picture.switches()[2].name(), '*');

        assert!(
            Field::new("INCLUDETEXT".to_string(), None, false)
                .external_include()
                .is_err()
        );
        assert!(
            Field::new(r"INCLUDETEXT \c Word8".to_string(), None, false)
                .external_include()
                .is_err()
        );
        assert!(
            Field::new(
                r#"INCLUDEPICTURE "picture.gif" Selector"#.to_string(),
                None,
                false,
            )
            .external_include()
            .is_err()
        );
        assert!(
            Field::new(
                r#"INCLUDEPICTURE "picture.gif" \d unexpected"#.to_string(),
                None,
                false,
            )
            .external_include()
            .is_err()
        );
        assert!(
            Field::new(r"INCLUDETEXT source \! unexpected".to_string(), None, false)
                .external_include()
                .is_err()
        );
        assert!(
            Field::new(r"INCLUDETEXT source \e".to_string(), None, false)
                .external_include()
                .is_err()
        );
        let not_include = Field::new("INCLUDETEXTUAL missing.docx".to_string(), None, false);
        assert!(!not_include.is_include_text());
        assert!(!not_include.is_include_picture());
        assert!(not_include.external_include().unwrap().is_none());
    }

    #[test]
    fn parses_toc_fields_and_standard_switches() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" TOC \o &quot;1-3&quot; \h \z \b &quot;Main Bookmark&quot; " w:dirty="true" w:fldLock="on">
                <w:r><w:t>Introduction</w:t><w:tab/><w:t>1</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin"/></w:r>
            <w:r><w:instrText>TOC\o&quot;2-4&quot;\u \n &quot;2-2&quot; \* MERGEFORMAT</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate"/></w:r>
            <w:r><w:t>Chapter</w:t><w:tab/><w:t>4</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="TOCENTRY \f ignored"><w:r><w:t>not a TOC</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
        let fields = Field::extract_from_document(xml).unwrap();
        assert_eq!(fields.len(), 3);
        assert!(fields[0].is_table_of_contents());
        assert!(fields[1].is_table_of_contents());
        assert!(!fields[2].is_table_of_contents());

        let first = fields[0].table_of_contents().unwrap().unwrap();
        assert_eq!(first.cached_result(), Some("Introduction\t1"));
        assert!(first.is_dirty());
        assert!(first.is_locked());
        assert!(first.includes_hyperlinks());
        assert!(first.hides_page_numbers_in_web_layout());
        assert!(!first.uses_outline_levels());
        assert_eq!(first.switches()[0].name(), 'o');
        assert_eq!(first.switches()[0].argument(), Some("1-3"));
        assert_eq!(first.switches()[3].argument(), Some("Main Bookmark"));
        assert_eq!(
            first.heading_style_levels().unwrap(),
            vec![TableOfContentsLevelRange::new(1, 3).unwrap()]
        );

        let second = fields[1].table_of_contents().unwrap().unwrap();
        assert_eq!(second.cached_result(), Some("Chapter\t4"));
        assert!(second.uses_outline_levels());
        assert!(!second.includes_hyperlinks());
        assert_eq!(second.switches()[0].name(), 'o');
        assert_eq!(second.switches()[0].argument(), Some("2-4"));
        assert_eq!(second.switches()[3].name(), '*');
        assert_eq!(second.switches()[3].argument(), Some("MERGEFORMAT"));
        assert_eq!(
            second.heading_style_levels().unwrap(),
            vec![TableOfContentsLevelRange::new(2, 4).unwrap()]
        );
    }

    #[test]
    fn parses_citation_and_bibliography_fields() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" CITATION Doe2024 \m &quot;Smith 2025&quot; \l 1033 \p &quot;14&quot; " w:dirty="true" w:fldLock="on">
                <w:r><w:t>(Doe, 2024; Smith, 2025, p. 14)</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>BIBLIOGRAPHY \l 1033 \f 1036 \m Doe2024 \m Smith2025</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>Doe. Example work.</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="CITATIONEXTRA ignored"><w:r><w:t>not a citation</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
        let fields = Field::extract_from_document(xml).unwrap();
        assert_eq!(fields.len(), 3);
        assert!(fields[0].is_citation());
        assert!(fields[1].is_bibliography());
        assert!(!fields[2].is_citation());
        assert!(!fields[2].is_bibliography());

        let citation = fields[0].citation().unwrap().unwrap();
        assert_eq!(
            citation.cached_result(),
            Some("(Doe, 2024; Smith, 2025, p. 14)")
        );
        assert!(citation.is_dirty());
        assert!(citation.is_locked());
        assert_eq!(citation.primary_source_tag(), "Doe2024");
        assert_eq!(citation.source_tags(), ["Doe2024", "Smith 2025"]);
        assert_eq!(citation.additional_source_tags(), ["Smith 2025"]);
        assert_eq!(citation.switches()[0].name(), 'm');
        assert_eq!(citation.switches()[0].argument(), Some("Smith 2025"));
        assert!(citation.has_switch('l'));
        assert!(citation.has_switch('p'));

        let documented_order = Field::new(
            r#"CITATION \l 1033 "Che 01" \v 3 \m Kra \v 2"#.to_string(),
            None,
            true,
        );
        let documented = documented_order.citation().unwrap().unwrap();
        assert_eq!(documented.source_tags(), ["Che 01", "Kra"]);
        assert_eq!(documented.switches()[0].name(), 'l');
        assert_eq!(documented.switches()[0].argument(), Some("1033"));
        assert!(documented.is_dirty());

        let bibliography = fields[1].bibliography().unwrap().unwrap();
        assert_eq!(bibliography.cached_result(), Some("Doe. Example work."));
        assert!(bibliography.is_dirty());
        assert!(bibliography.is_locked());
        assert_eq!(bibliography.switches()[0].name(), 'l');
        assert_eq!(bibliography.switches()[0].argument(), Some("1033"));
        assert!(bibliography.has_switch('f'));
        assert_eq!(bibliography.switches()[1].argument(), Some("1036"));
        assert_eq!(bibliography.switches()[2].argument(), Some("Doe2024"));
        assert_eq!(bibliography.switches()[3].argument(), Some("Smith2025"));
    }

    #[test]
    fn parses_document_variable_fields_without_resolving_values() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" DOCVARIABLE &quot;Customer Region&quot; \* MERGEFORMAT " w:dirty="true" w:fldLock="on">
                <w:r><w:t>cached region</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>DOCVARIABLE CustomerName</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>cached customer</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="DOCVARIABLES CustomerName"><w:r><w:t>not a variable</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
        let fields = Field::extract_from_document(xml).unwrap();
        assert_eq!(fields.len(), 3);
        assert!(fields[0].is_document_variable());
        assert!(fields[1].is_document_variable());
        assert!(!fields[2].is_document_variable());

        let region = fields[0].document_variable().unwrap().unwrap();
        assert_eq!(region.variable_name(), "Customer Region");
        assert_eq!(region.cached_result(), Some("cached region"));
        assert!(region.is_dirty());
        assert!(region.is_locked());
        assert!(region.has_switch('*'));
        assert_eq!(region.switches()[0].argument(), Some("MERGEFORMAT"));

        let customer = fields[1].document_variable().unwrap().unwrap();
        assert_eq!(customer.variable_name(), "CustomerName");
        assert_eq!(customer.cached_result(), Some("cached customer"));
        assert!(customer.is_dirty());
        assert!(customer.is_locked());
        assert!(customer.switches().is_empty());
        assert!(fields[2].document_variable().unwrap().is_none());
    }

    #[test]
    fn rejects_invalid_document_variable_field_semantics() {
        let missing_name = Field::new("DOCVARIABLE \\* MERGEFORMAT".to_string(), None, false);
        assert!(missing_name.document_variable().is_err());

        let empty_name = Field::new(r#"DOCVARIABLE "" "#.to_string(), None, false);
        assert!(empty_name.document_variable().is_err());

        let unexpected_operand =
            Field::new("DOCVARIABLE Customer unexpected".to_string(), None, false);
        assert!(unexpected_operand.document_variable().is_err());
    }

    #[test]
    fn parses_document_property_fields_without_resolving_values() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" DOCPROPERTY &quot;Project Name&quot; \* MERGEFORMAT \@ &quot;MMMM d, yyyy&quot; " w:dirty="true" w:fldLock="on">
                <w:r><w:t>cached project</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>docproperty Revision</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>cached revision</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="DOCPROPERTYS ProjectName"><w:r><w:t>not a property</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
        let fields = Field::extract_from_document(xml).unwrap();
        assert_eq!(fields.len(), 3);
        assert!(fields[0].is_document_property());
        assert!(fields[1].is_document_property());
        assert!(!fields[2].is_document_property());

        let project = fields[0].document_property().unwrap().unwrap();
        assert_eq!(project.property_name(), "Project Name");
        assert_eq!(project.cached_result(), Some("cached project"));
        assert!(project.is_dirty());
        assert!(project.is_locked());
        assert_eq!(project.switches().len(), 2);
        assert_eq!(project.switches()[0].name(), '*');
        assert_eq!(project.switches()[0].argument(), Some("MERGEFORMAT"));
        assert_eq!(project.switches()[1].name(), '@');
        assert_eq!(project.switches()[1].argument(), Some("MMMM d, yyyy"));
        assert!(project.has_switch('*'));
        assert!(project.has_switch('@'));

        let revision = fields[1].document_property().unwrap().unwrap();
        assert_eq!(revision.property_name(), "Revision");
        assert_eq!(revision.cached_result(), Some("cached revision"));
        assert!(revision.is_dirty());
        assert!(revision.is_locked());
        assert!(revision.switches().is_empty());
        assert!(fields[2].document_property().unwrap().is_none());
    }

    #[test]
    fn rejects_invalid_document_property_field_semantics() {
        for instruction in [
            r#"DOCPROPERTY \* MERGEFORMAT"#,
            r#"DOCPROPERTY """#,
            "DOCPROPERTY Project unexpected",
            r#"DOCPROPERTY Project \"#,
        ] {
            let field = Field::new(instruction.to_string(), None, false);
            assert!(field.document_property().is_err(), "{instruction}");
        }

        let too_long = Field::new(
            format!(
                "DOCPROPERTY {}",
                "x".repeat(MAX_DOCUMENT_PROPERTY_FIELD_INSTRUCTION_BYTES)
            ),
            None,
            false,
        );
        assert!(too_long.document_property().is_err());
    }

    #[test]
    fn parses_document_information_fields_without_reading_or_calculating_values() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" TITLE \* MERGEFORMAT " w:dirty="true" w:fldLock="on">
                <w:r><w:t>cached title</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>author \@ "opaque format"</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>cached author</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="AUTHORS"><w:r><w:t>not an author field</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
        let extracted = Field::extract_from_document(xml).unwrap();
        assert_eq!(extracted.len(), 3);
        assert!(extracted[0].is_document_information());
        assert!(extracted[1].is_document_information());
        assert!(!extracted[2].is_document_information());

        let title = extracted[0].document_information().unwrap().unwrap();
        assert_eq!(title.kind(), DocumentInformationFieldKind::Title);
        assert_eq!(title.cached_result(), Some("cached title"));
        assert!(title.is_dirty());
        assert!(title.is_locked());
        assert_eq!(title.switches()[0].name(), '*');
        assert_eq!(title.switches()[0].argument(), Some("MERGEFORMAT"));

        let author = extracted[1].document_information().unwrap().unwrap();
        assert_eq!(author.kind(), DocumentInformationFieldKind::Author);
        assert_eq!(author.cached_result(), Some("cached author"));
        assert!(author.is_dirty());
        assert!(author.is_locked());
        assert!(author.has_switch('@'));
        assert_eq!(author.switches()[0].argument(), Some("opaque format"));
        assert!(extracted[2].document_information().unwrap().is_none());

        for (instruction, kind) in [
            (r"TITLE \* MERGEFORMAT", DocumentInformationFieldKind::Title),
            (
                r"SUBJECT \* MERGEFORMAT",
                DocumentInformationFieldKind::Subject,
            ),
            (
                r"AUTHOR \* MERGEFORMAT",
                DocumentInformationFieldKind::Author,
            ),
            (
                r"KEYWORDS \* MERGEFORMAT",
                DocumentInformationFieldKind::Keywords,
            ),
            (
                r"COMMENTS \* MERGEFORMAT",
                DocumentInformationFieldKind::Comments,
            ),
            (
                r"LASTSAVEDBY \* MERGEFORMAT",
                DocumentInformationFieldKind::LastSavedBy,
            ),
            (
                r"CREATEDATE \* MERGEFORMAT",
                DocumentInformationFieldKind::CreateDate,
            ),
            (
                r"SAVEDATE \* MERGEFORMAT",
                DocumentInformationFieldKind::SaveDate,
            ),
            (
                r"PRINTDATE \* MERGEFORMAT",
                DocumentInformationFieldKind::PrintDate,
            ),
            (
                r"REVNUM \* MERGEFORMAT",
                DocumentInformationFieldKind::RevisionNumber,
            ),
            (
                r"EDITTIME \* MERGEFORMAT",
                DocumentInformationFieldKind::EditTime,
            ),
            (
                r"NUMPAGES \* MERGEFORMAT",
                DocumentInformationFieldKind::NumberOfPages,
            ),
            (
                r"NUMWORDS \* MERGEFORMAT",
                DocumentInformationFieldKind::NumberOfWords,
            ),
            (
                r"NUMCHARS \* MERGEFORMAT",
                DocumentInformationFieldKind::NumberOfCharacters,
            ),
        ] {
            let cached_result = format!("cached {}", kind.field_keyword());
            let field = Field::with_flags(
                instruction.to_string(),
                Some(cached_result.clone()),
                true,
                true,
            );
            let information = field.document_information().unwrap().unwrap();
            assert_eq!(information.kind(), kind);
            assert_eq!(information.instruction(), instruction);
            assert_eq!(information.cached_result(), Some(cached_result.as_str()));
            assert!(information.is_dirty());
            assert!(information.is_locked());
            assert_eq!(information.switches()[0].name(), '*');
        }
    }

    #[test]
    fn rejects_invalid_document_information_field_semantics() {
        for instruction in [
            "TITLE unexpected",
            r#"AUTHOR "unterminated"#,
            r"COMMENTS \",
            r"LASTSAVEDBY \* MERGEFORMAT unexpected",
            "NUMWORDS unexpected",
        ] {
            let field = Field::new(instruction.to_string(), None, false);
            assert!(field.document_information().is_err(), "{instruction}");
        }

        let too_long = Field::new(
            format!(
                "TITLE \\* {}",
                "x".repeat(MAX_DOCUMENT_INFORMATION_FIELD_INSTRUCTION_BYTES)
            ),
            None,
            false,
        );
        assert!(too_long.document_information().is_err());
        assert_eq!(
            Field::new("SAVEDATE".to_string(), None, false)
                .document_information()
                .unwrap()
                .unwrap()
                .kind(),
            DocumentInformationFieldKind::SaveDate
        );
        assert!(
            Field::new("SAVEDATES".to_string(), None, false)
                .document_information()
                .unwrap()
                .is_none()
        );
    }

    #[test]
    fn parses_document_context_fields_without_reading_paths_files_or_layout() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" FILENAME \p " w:dirty="true" w:fldLock="on">
                <w:r><w:t>cached file name</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>template \* MERGEFORMAT</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>cached template</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr=" SECTIONPAGES \* MERGEFORMAT " w:dirty="true" w:fldLock="on">
                <w:r><w:t>cached section pages</w:t></w:r>
            </w:fldSimple>
            <w:fldSimple w:instr="FILENAMES"><w:r><w:t>not a file-name field</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
        let extracted = Field::extract_from_document(xml).unwrap();
        assert_eq!(extracted.len(), 4);
        assert!(extracted[0].is_document_context());
        assert!(extracted[1].is_document_context());
        assert!(extracted[2].is_document_context());
        assert!(!extracted[3].is_document_context());

        let file_name = extracted[0].document_context().unwrap().unwrap();
        assert_eq!(file_name.kind(), DocumentContextFieldKind::FileName);
        assert_eq!(file_name.cached_result(), Some("cached file name"));
        assert!(file_name.is_dirty());
        assert!(file_name.is_locked());
        assert!(file_name.has_switch('p'));

        let template = extracted[1].document_context().unwrap().unwrap();
        assert_eq!(template.kind(), DocumentContextFieldKind::Template);
        assert_eq!(template.cached_result(), Some("cached template"));
        assert!(template.is_dirty());
        assert!(template.is_locked());
        assert!(template.has_switch('*'));

        let section_pages = extracted[2].document_context().unwrap().unwrap();
        assert_eq!(section_pages.kind(), DocumentContextFieldKind::SectionPages);
        assert_eq!(section_pages.cached_result(), Some("cached section pages"));
        assert!(section_pages.is_dirty());
        assert!(section_pages.is_locked());
        assert!(section_pages.has_switch('*'));
        assert!(extracted[3].document_context().unwrap().is_none());

        for (instruction, kind, switch_name) in [
            (
                r"FILENAME \p",
                DocumentContextFieldKind::FileName,
                'p',
            ),
            (
                r"TEMPLATE \* MERGEFORMAT",
                DocumentContextFieldKind::Template,
                '*',
            ),
            (
                r#"DATE \@ "opaque date format""#,
                DocumentContextFieldKind::Date,
                '@',
            ),
            (
                r#"TIME \@ "opaque time format""#,
                DocumentContextFieldKind::Time,
                '@',
            ),
            (
                r"PAGE \* MERGEFORMAT",
                DocumentContextFieldKind::Page,
                '*',
            ),
            (
                r"FILESIZE \* MERGEFORMAT",
                DocumentContextFieldKind::FileSize,
                '*',
            ),
            (
                r"SECTION \* MERGEFORMAT",
                DocumentContextFieldKind::Section,
                '*',
            ),
            (
                r"SECTIONPAGES \* MERGEFORMAT",
                DocumentContextFieldKind::SectionPages,
                '*',
            ),
        ] {
            let cached_result = format!("cached {}", kind.field_keyword());
            let field = Field::with_flags(
                instruction.to_string(),
                Some(cached_result.clone()),
                true,
                true,
            );
            let context = field.document_context().unwrap().unwrap();
            assert_eq!(context.kind(), kind);
            assert_eq!(context.instruction(), instruction);
            assert_eq!(context.cached_result(), Some(cached_result.as_str()));
            assert!(context.is_dirty());
            assert!(context.is_locked());
            assert!(context.has_switch(switch_name));
        }
    }

    #[test]
    fn rejects_invalid_document_context_field_semantics() {
        for instruction in [
            "FILENAME unexpected",
            r"TEMPLATE \",
            r"FILENAME \ ",
            "PAGE unexpected",
            "SECTIONPAGES unexpected",
        ] {
            let field = Field::new(instruction.to_string(), None, false);
            assert!(field.document_context().is_err(), "{instruction}");
        }

        let too_long = Field::new(
            format!(
                "FILENAME \\* {}",
                "x".repeat(MAX_DOCUMENT_CONTEXT_FIELD_INSTRUCTION_BYTES)
            ),
            None,
            false,
        );
        assert!(too_long.document_context().is_err());
        assert!(
            Field::new("FILENAMES".to_string(), None, false)
                .document_context()
                .unwrap()
                .is_none()
        );
        assert!(
            Field::new("PAGES".to_string(), None, false)
                .document_context()
                .unwrap()
                .is_none()
        );
        assert!(
            Field::new("SECTIONPAGE".to_string(), None, false)
                .document_context()
                .unwrap()
                .is_none()
        );
    }

    #[test]
    fn parses_macro_button_fields_without_resolving_or_executing_targets() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" MACROBUTTON &quot;Never Run&quot; &quot;Click here&quot; " w:dirty="true" w:fldLock="on">
                <w:r><w:t>cached button</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>MACROBUTTON NoMacro "Click again"</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>cached second button</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="MACROBUTTONS NeverRun Button"><w:r><w:t>not a macro button</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
        let fields = Field::extract_from_document(xml).unwrap();
        assert_eq!(fields.len(), 3);
        assert!(fields[0].is_macro_button());
        assert!(fields[1].is_macro_button());
        assert!(!fields[2].is_macro_button());

        let first = fields[0].macro_button().unwrap().unwrap();
        assert_eq!(first.macro_name(), "Never Run");
        assert_eq!(first.display_text(), "Click here");
        assert_eq!(first.cached_result(), Some("cached button"));
        assert!(first.is_dirty());
        assert!(first.is_locked());

        let second = fields[1].macro_button().unwrap().unwrap();
        assert_eq!(second.macro_name(), "NoMacro");
        assert_eq!(second.display_text(), "Click again");
        assert_eq!(second.cached_result(), Some("cached second button"));
        assert!(second.is_dirty());
        assert!(second.is_locked());
        assert!(fields[2].macro_button().unwrap().is_none());
    }

    #[test]
    fn rejects_invalid_macro_button_field_semantics() {
        let missing_name = Field::new("MACROBUTTON".to_string(), None, false);
        assert!(missing_name.macro_button().is_err());

        let empty_name = Field::new(r#"MACROBUTTON "" Button"#.to_string(), None, false);
        assert!(empty_name.macro_button().is_err());

        let missing_button = Field::new("MACROBUTTON NeverRun".to_string(), None, false);
        assert!(missing_button.macro_button().is_err());

        let empty_button = Field::new(r#"MACROBUTTON NeverRun """#.to_string(), None, false);
        assert!(empty_button.macro_button().is_err());

        let extra_argument = Field::new(
            "MACROBUTTON NeverRun Button unexpected".to_string(),
            None,
            false,
        );
        assert!(extra_argument.macro_button().is_err());

        let unsupported_switch = Field::new(
            r#"MACROBUTTON NeverRun Button \* MERGEFORMAT"#.to_string(),
            None,
            false,
        );
        assert!(unsupported_switch.macro_button().is_err());
    }

    #[test]
    fn parses_go_to_button_fields_without_resolving_or_navigating_to_targets() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" GOTOBUTTON MyBookmark &quot;Jump to bookmark&quot; " w:dirty="true" w:fldLock="on">
                <w:r><w:t>cached bookmark button</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>GOTOBUTTON "f 2" Footnote</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>cached footnote button</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="GOTOBUTTONS MyBookmark Button"><w:r><w:t>not a button</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
        let fields = Field::extract_from_document(xml).unwrap();
        assert_eq!(fields.len(), 3);
        assert!(fields[0].is_go_to_button());
        assert!(fields[1].is_go_to_button());
        assert!(!fields[2].is_go_to_button());

        let first = fields[0].go_to_button().unwrap().unwrap();
        assert_eq!(first.target(), "MyBookmark");
        assert_eq!(first.button_text(), "Jump to bookmark");
        assert_eq!(first.cached_result(), Some("cached bookmark button"));
        assert!(first.is_dirty());
        assert!(first.is_locked());

        let second = fields[1].go_to_button().unwrap().unwrap();
        assert_eq!(second.target(), "f 2");
        assert_eq!(second.button_text(), "Footnote");
        assert_eq!(second.cached_result(), Some("cached footnote button"));
        assert!(second.is_dirty());
        assert!(second.is_locked());
        assert!(fields[2].go_to_button().unwrap().is_none());
    }

    #[test]
    fn rejects_invalid_go_to_button_field_semantics() {
        let missing_target = Field::new("GOTOBUTTON".to_string(), None, false);
        assert!(missing_target.go_to_button().is_err());

        let empty_target = Field::new(r#"GOTOBUTTON "" Button"#.to_string(), None, false);
        assert!(empty_target.go_to_button().is_err());

        let missing_button = Field::new("GOTOBUTTON Destination".to_string(), None, false);
        assert!(missing_button.go_to_button().is_err());

        let empty_button = Field::new(r#"GOTOBUTTON Destination """#.to_string(), None, false);
        assert!(empty_button.go_to_button().is_err());

        let extra_argument = Field::new(
            "GOTOBUTTON Destination Button unexpected".to_string(),
            None,
            false,
        );
        assert!(extra_argument.go_to_button().is_err());

        let unsupported_switch = Field::new(
            r#"GOTOBUTTON Destination Button \* MERGEFORMAT"#.to_string(),
            None,
            false,
        );
        assert!(unsupported_switch.go_to_button().is_err());
    }

    #[test]
    fn parses_active_content_fields_without_loading_or_activating_them() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" ADDIN opaque-add-in-data " w:dirty="true" w:fldLock="on">
                <w:r><w:t>cached add-in result</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>control opaque-ocx-metadata</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>cached control result</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="HTMLCONTROL opaque-html-control-metadata">
                <w:r><w:t>cached html result</w:t></w:r>
            </w:fldSimple>
            <w:fldSimple w:instr="ADDINS not-an-add-in"><w:r><w:t>not active content</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
        let fields = Field::extract_from_document(xml).unwrap();
        assert_eq!(fields.len(), 4);
        assert!(fields[0].is_add_in_field());
        assert!(fields[0].is_active_content_field());
        assert!(fields[1].is_control_field());
        assert!(fields[1].is_active_content_field());
        assert!(fields[2].is_html_control_field());
        assert!(fields[2].is_active_content_field());
        assert!(!fields[3].is_active_content_field());

        let add_in = fields[0].active_content_field().unwrap().unwrap();
        assert_eq!(add_in.kind(), ActiveContentFieldKind::AddIn);
        assert_eq!(add_in.cached_result(), Some("cached add-in result"));
        assert!(add_in.is_dirty());
        assert!(add_in.is_locked());

        let ocx = fields[1].active_content_field().unwrap().unwrap();
        assert_eq!(ocx.kind(), ActiveContentFieldKind::OcxControl);
        assert_eq!(ocx.cached_result(), Some("cached control result"));
        assert!(ocx.is_dirty());
        assert!(ocx.is_locked());

        let html = fields[2].active_content_field().unwrap().unwrap();
        assert_eq!(html.kind(), ActiveContentFieldKind::HtmlControl);
        assert_eq!(html.cached_result(), Some("cached html result"));
        assert!(!html.is_dirty());
        assert!(!html.is_locked());
        assert!(fields[3].active_content_field().unwrap().is_none());
    }

    #[test]
    fn parses_auto_text_fields_without_lookup_or_insertion() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" GLOSSARY &quot;Legacy Clause&quot; \* MERGEFORMAT \q opaque " w:dirty="true" w:fldLock="on">
                <w:r><w:t>cached glossary entry</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>autotext "Reusable Clause" \* MERGEFORMAT</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>cached auto text entry</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="AUTOTEXTLIST display"><w:r><w:t>not an auto text field</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
        let fields = Field::extract_from_document(xml).unwrap();
        assert_eq!(fields.len(), 3);
        assert!(fields[0].is_auto_text_field());
        assert!(fields[1].is_auto_text_field());
        assert!(!fields[2].is_auto_text_field());

        let glossary = fields[0].auto_text_field().unwrap().unwrap();
        assert_eq!(glossary.kind(), AutoTextFieldKind::Glossary);
        assert_eq!(glossary.entry_name(), "Legacy Clause");
        assert_eq!(glossary.unknown_switches().len(), 2);
        assert_eq!(glossary.unknown_switches()[0].name(), '*');
        assert_eq!(glossary.unknown_switches()[0].argument(), Some("MERGEFORMAT"));
        assert_eq!(glossary.unknown_switches()[1].name(), 'q');
        assert_eq!(glossary.unknown_switches()[1].argument(), Some("opaque"));
        assert_eq!(glossary.cached_result(), Some("cached glossary entry"));
        assert!(glossary.is_dirty());
        assert!(glossary.is_locked());

        let auto_text = fields[1].auto_text_field().unwrap().unwrap();
        assert_eq!(auto_text.kind(), AutoTextFieldKind::AutoText);
        assert_eq!(auto_text.entry_name(), "Reusable Clause");
        assert_eq!(auto_text.unknown_switches().len(), 1);
        assert_eq!(auto_text.unknown_switches()[0].name(), '*');
        assert_eq!(
            auto_text.unknown_switches()[0].argument(),
            Some("MERGEFORMAT")
        );
        assert_eq!(auto_text.cached_result(), Some("cached auto text entry"));
        assert!(auto_text.is_dirty());
        assert!(auto_text.is_locked());
        assert!(fields[2].auto_text_field().unwrap().is_none());
    }

    #[test]
    fn rejects_invalid_auto_text_fields_without_lookup_or_insertion() {
        for instruction in [
            "GLOSSARY",
            r#"GLOSSARY ""#,
            "GLOSSARY Entry unexpected",
            r#"GLOSSARY Entry \"#,
        ] {
            let field = Field::new(instruction.to_string(), None, false);
            assert!(field.auto_text_field().is_err(), "{instruction}");
        }

        let too_long = Field::new(
            format!(
                "AUTOTEXT Entry {}",
                "x".repeat(MAX_AUTO_TEXT_FIELD_INSTRUCTION_BYTES)
            ),
            None,
            false,
        );
        assert!(too_long.auto_text_field().is_err());
    }

    #[test]
    fn parses_auto_text_list_fields_without_selection_or_insertion() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" AUTOTEXTLIST &quot;Choose a name&quot; \s &quot;Name Style&quot; \t &quot;Right-click to select&quot; \* MERGEFORMAT \q opaque " w:dirty="true" w:fldLock="on">
                <w:r><w:t>cached selection</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>autotextlist \s NameStyle</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>cached style-only selection</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="AUTOTEXTLISTS display"><w:r><w:t>not a list field</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
        let fields = Field::extract_from_document(xml).unwrap();
        assert_eq!(fields.len(), 3);
        assert!(fields[0].is_auto_text_list_field());
        assert!(fields[1].is_auto_text_list_field());
        assert!(!fields[2].is_auto_text_list_field());

        let list = fields[0].auto_text_list_field().unwrap().unwrap();
        assert_eq!(list.display_text(), Some("Choose a name"));
        assert_eq!(
            list.options(),
            &[
                AutoTextListOption::Style("Name Style".to_string()),
                AutoTextListOption::Tip("Right-click to select".to_string()),
            ]
        );
        assert_eq!(list.unknown_switches().len(), 2);
        assert_eq!(list.unknown_switches()[0].name(), '*');
        assert_eq!(list.unknown_switches()[0].argument(), Some("MERGEFORMAT"));
        assert_eq!(list.unknown_switches()[1].name(), 'q');
        assert_eq!(list.unknown_switches()[1].argument(), Some("opaque"));
        assert_eq!(list.cached_result(), Some("cached selection"));
        assert!(list.is_dirty());
        assert!(list.is_locked());

        let style_only = fields[1].auto_text_list_field().unwrap().unwrap();
        assert_eq!(style_only.display_text(), None);
        assert_eq!(
            style_only.options(),
            &[AutoTextListOption::Style("NameStyle".to_string())]
        );
        assert_eq!(
            style_only.cached_result(),
            Some("cached style-only selection")
        );
        assert!(style_only.is_dirty());
        assert!(style_only.is_locked());
        assert!(fields[2].auto_text_list_field().unwrap().is_none());

        let empty_display = Field::new(r#"AUTOTEXTLIST "" \s NameStyle"#.to_string(), None, false)
            .auto_text_list_field()
            .unwrap()
            .unwrap();
        assert_eq!(empty_display.display_text(), Some(""));
    }

    #[test]
    fn rejects_invalid_auto_text_list_fields_without_selection_or_insertion() {
        for instruction in [
            r#"AUTOTEXTLIST \s"#,
            r#"AUTOTEXTLIST \t"#,
            "AUTOTEXTLIST display unexpected",
            r#"AUTOTEXTLIST \"#,
        ] {
            let field = Field::new(instruction.to_string(), None, false);
            assert!(field.auto_text_list_field().is_err(), "{instruction}");
        }

        let too_long = Field::new(
            format!(
                "AUTOTEXTLIST {}",
                "x".repeat(MAX_AUTO_TEXT_LIST_FIELD_INSTRUCTION_BYTES)
            ),
            None,
            false,
        );
        assert!(too_long.auto_text_list_field().is_err());
    }

    #[test]
    fn parses_user_identity_fields_without_reading_host_identity() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" USERADDRESS &quot;10 Top Secret Lane&quot; \* Upper " w:dirty="true" w:fldLock="on">
                <w:r><w:t>10 TOP SECRET LANE</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>userinitials \* Lower</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>dw</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="USERNAME &quot;Ada Lovelace&quot; \* FirstCap"><w:r><w:t>Ada Lovelace</w:t></w:r></w:fldSimple>
            <w:fldSimple w:instr="USERNAMES Ada"><w:r><w:t>not a user identity field</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
        let fields = Field::extract_from_document(xml).unwrap();
        assert_eq!(fields.len(), 4);
        assert!(fields[0].is_user_address());
        assert!(fields[0].is_user_identity_field());
        assert!(fields[1].is_user_initials());
        assert!(fields[1].is_user_identity_field());
        assert!(fields[2].is_user_name());
        assert!(fields[2].is_user_identity_field());
        assert!(!fields[3].is_user_identity_field());

        let address = fields[0].user_identity_field().unwrap().unwrap();
        assert_eq!(address.kind(), UserIdentityFieldKind::Address);
        assert_eq!(address.override_value(), Some("10 Top Secret Lane"));
        assert_eq!(address.formatting(), Some(UserIdentityFormatting::Upper));
        assert_eq!(address.cached_result(), Some("10 TOP SECRET LANE"));
        assert!(address.is_dirty());
        assert!(address.is_locked());

        let initials = fields[1].user_identity_field().unwrap().unwrap();
        assert_eq!(initials.kind(), UserIdentityFieldKind::Initials);
        assert_eq!(initials.override_value(), None);
        assert_eq!(initials.formatting(), Some(UserIdentityFormatting::Lower));
        assert_eq!(initials.cached_result(), Some("dw"));
        assert!(initials.is_dirty());
        assert!(initials.is_locked());

        let name = fields[2].user_identity_field().unwrap().unwrap();
        assert_eq!(name.kind(), UserIdentityFieldKind::Name);
        assert_eq!(name.override_value(), Some("Ada Lovelace"));
        assert_eq!(name.formatting(), Some(UserIdentityFormatting::FirstCap));
        assert_eq!(name.cached_result(), Some("Ada Lovelace"));
        assert!(!name.is_dirty());
        assert!(!name.is_locked());
        assert!(fields[3].user_identity_field().unwrap().is_none());
    }

    #[test]
    fn rejects_invalid_user_identity_field_semantics() {
        let missing_format = Field::new("USERADDRESS \\*".to_string(), None, false);
        assert!(missing_format.user_identity_field().is_err());

        let unsupported_format = Field::new("USERINITIALS \\* Title".to_string(), None, false);
        assert!(unsupported_format.user_identity_field().is_err());

        let duplicate_format = Field::new("USERNAME \\* Upper \\* Lower".to_string(), None, false);
        assert!(duplicate_format.user_identity_field().is_err());

        let unsupported_switch = Field::new("USERNAME Ada \\l 1033".to_string(), None, false);
        assert!(unsupported_switch.user_identity_field().is_err());

        let unexpected_text = Field::new("USERADDRESS Ada Lovelace".to_string(), None, false);
        assert!(unexpected_text.user_identity_field().is_err());

        let blank_override = Field::new(r#"USERNAME "" \* Caps"#.to_string(), None, false);
        let blank_override = blank_override.user_identity_field().unwrap().unwrap();
        assert_eq!(blank_override.override_value(), Some(""));
        assert_eq!(
            blank_override.formatting(),
            Some(UserIdentityFormatting::Caps)
        );
    }

    #[test]
    fn parses_inert_advance_fields_without_changing_layout() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" ADVANCE \u 6 \d 12 \l 20 \r -4 \x 150 \y &quot;72&quot; \d -3 " w:dirty="true" w:fldLock="on">
                <w:r><w:t>cached placement</w:t></w:r>
            </w:fldSimple>
            <w:fldSimple w:instr="ADVANCER \u 6"><w:r><w:t>not an advance field</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
        let fields = Field::extract_from_document(xml).unwrap();
        assert_eq!(fields.len(), 2);
        assert!(fields[0].is_advance_field());
        assert!(!fields[1].is_advance_field());

        let advance = fields[0].advance_field().unwrap().unwrap();
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
        assert!(fields[1].advance_field().unwrap().is_none());

        let no_adjustments = Field::new("aDvAnCe".to_string(), None, false);
        let no_adjustments = no_adjustments.advance_field().unwrap().unwrap();
        assert!(no_adjustments.adjustments().is_empty());
        assert_eq!(no_adjustments.cached_result(), None);
    }

    #[test]
    fn rejects_invalid_advance_field_semantics() {
        for instruction in [
            r#"ADVANCE \d"#,
            r#"ADVANCE \z 10"#,
            r#"ADVANCE \x 1.5"#,
            r#"ADVANCE \u 9223372036854775808"#,
            "ADVANCE 12",
            r#"ADVANCE \d 6 trailing"#,
        ] {
            let field = Field::new(instruction.to_string(), None, false);
            assert!(field.advance_field().is_err(), "{instruction}");
        }
    }

    #[test]
    fn parses_table_of_authorities_and_entry_fields() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" TOA \c 0 \b &quot;Authorities&quot; \p \f \d &quot;-&quot; \s &quot;Chapter&quot; \e &quot;, &quot; \g &quot;&#x2013;&quot; \h \l &quot;, &quot; " w:dirty="true" w:fldLock="on">
                <w:r><w:t>Cases</w:t><w:tab/><w:t>1, 5</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>TA\l&quot;Long citation&quot;\s &quot;Short citation&quot; \c 1 \b \i</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>hidden citation marker</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="TABLE \c 1"><w:r><w:t>not an authority table</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
        let fields = Field::extract_from_document(xml).unwrap();
        assert_eq!(fields.len(), 3);
        assert!(fields[0].is_table_of_authorities());
        assert!(fields[1].is_table_of_authorities_entry());
        assert!(!fields[2].is_table_of_authorities());
        assert!(!fields[2].is_table_of_authorities_entry());

        let toa = fields[0].table_of_authorities().unwrap().unwrap();
        assert_eq!(toa.cached_result(), Some("Cases\t1, 5"));
        assert!(toa.is_dirty());
        assert!(toa.is_locked());
        assert_eq!(toa.category().unwrap(), Some(0));
        assert_eq!(toa.bookmark().unwrap(), Some("Authorities"));
        assert!(toa.uses_passim());
        assert!(toa.keeps_entry_formatting());
        assert_eq!(toa.sequence_page_separator().unwrap(), Some("-"));
        assert_eq!(toa.sequence_name().unwrap(), Some("Chapter"));
        assert_eq!(toa.entry_page_separator().unwrap(), Some(", "));
        assert_eq!(toa.page_range_separator().unwrap(), Some("–"));
        assert!(toa.includes_category_headers());
        assert_eq!(toa.page_number_separator().unwrap(), Some(", "));

        let entry = fields[1].table_of_authorities_entry().unwrap().unwrap();
        assert_eq!(entry.cached_result(), Some("hidden citation marker"));
        assert!(entry.is_dirty());
        assert!(entry.is_locked());
        assert_eq!(entry.long_citation().unwrap(), Some("Long citation"));
        assert_eq!(entry.short_citation().unwrap(), Some("Short citation"));
        assert_eq!(entry.category().unwrap(), Some(1));
        assert!(entry.is_bold());
        assert!(entry.is_italic());
    }

    #[test]
    fn parses_index_and_index_entry_fields() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>
            <w:fldSimple w:instr=" INDEX \b Scope \c 2 \d &quot;.&quot; \e &quot;; &quot; \f &quot;topics&quot; \g &quot; to &quot; \h &quot;A&quot; \k &quot;: &quot; \l &quot; / &quot; \o &quot;P&quot; \p a-m \r \s Chapter \y \z 1033 " w:dirty="true" w:fldLock="on">
                <w:r><w:t>Rulers</w:t><w:tab/><w:t>4</w:t></w:r>
            </w:fldSimple>
            <w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>
            <w:r><w:instrText>XE&quot;Machiavelli: The Prince&quot;\b\i\f &quot;topics&quot; \r IndexRange \t &quot;See Rulers&quot; \y &quot;ma&quot;</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/></w:r>
            <w:r><w:t>hidden index marker</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            <w:fldSimple w:instr="INDEXENTRY \f ignored"><w:r><w:t>not an index</w:t></w:r></w:fldSimple>
        </w:p></w:body></w:document>"#;
        let fields = Field::extract_from_document(xml).unwrap();
        assert_eq!(fields.len(), 3);
        assert!(fields[0].is_index());
        assert!(fields[1].is_index_entry());
        assert!(!fields[2].is_index());
        assert!(!fields[2].is_index_entry());

        let index = fields[0].index().unwrap().unwrap();
        assert_eq!(index.cached_result(), Some("Rulers\t4"));
        assert!(index.is_dirty());
        assert!(index.is_locked());
        assert_eq!(index.bookmark().unwrap(), Some("Scope"));
        assert_eq!(index.columns().unwrap(), Some(2));
        assert_eq!(index.sequence_page_separator().unwrap(), Some("."));
        assert_eq!(index.entry_page_separator().unwrap(), Some("; "));
        assert_eq!(index.entry_identifier().unwrap(), Some("topics"));
        assert_eq!(index.page_range_separator().unwrap(), Some(" to "));
        assert_eq!(index.alphabetic_group_heading().unwrap(), Some("A"));
        assert_eq!(index.cross_reference_separator().unwrap(), Some(": "));
        assert_eq!(index.page_reference_separator().unwrap(), Some(" / "));
        assert_eq!(
            index.sort_order().unwrap(),
            Some(IndexSortOrder::Pronunciation)
        );
        assert_eq!(index.letter_range().unwrap(), Some("a-m"));
        assert!(index.runs_subentries_inline());
        assert_eq!(index.sequence_name().unwrap(), Some("Chapter"));
        assert!(index.uses_yomi());
        assert_eq!(index.language_id().unwrap(), Some("1033"));

        let entry = fields[1].index_entry().unwrap().unwrap();
        assert_eq!(entry.cached_result(), Some("hidden index marker"));
        assert!(entry.is_dirty());
        assert!(entry.is_locked());
        assert_eq!(entry.entry(), "Machiavelli: The Prince");
        assert!(entry.is_bold());
        assert!(entry.is_italic());
        assert_eq!(entry.entry_identifier().unwrap(), Some("topics"));
        assert_eq!(entry.page_range_bookmark().unwrap(), Some("IndexRange"));
        assert_eq!(entry.cross_reference().unwrap(), Some("See Rulers"));
        assert_eq!(entry.yomi().unwrap(), Some("ma"));
    }

    #[test]
    fn rejects_invalid_table_of_authorities_semantics() {
        let invalid_toa = Field::new(r#"TOA \c 17"#.to_string(), None, false);
        let toa = invalid_toa.table_of_authorities().unwrap().unwrap();
        assert!(toa.category().is_err());

        let invalid_entry = Field::new(r#"TA \c 0"#.to_string(), None, false);
        let entry = invalid_entry.table_of_authorities_entry().unwrap().unwrap();
        assert!(entry.category().is_err());

        let duplicate = Field::new(r#"TOA \b "a" \b "b""#.to_string(), None, false);
        let toa = duplicate.table_of_authorities().unwrap().unwrap();
        assert!(toa.bookmark().is_err());
    }

    #[test]
    fn rejects_invalid_citation_and_bibliography_field_semantics() {
        let missing_source = Field::new("CITATION \\l 1033".to_string(), None, false);
        assert!(missing_source.citation().is_err());

        let empty_source = Field::new(r#"CITATION ""#.to_string(), None, false);
        assert!(empty_source.citation().is_err());

        let missing_multisource_tag =
            Field::new("CITATION Doe2024 \\m \\l 1033".to_string(), None, false);
        assert!(missing_multisource_tag.citation().is_err());

        let empty_multisource_tag =
            Field::new(r#"CITATION Doe2024 \m """#.to_string(), None, false);
        assert!(empty_multisource_tag.citation().is_err());

        let malformed_bibliography = Field::new("BIBLIOGRAPHY unexpected".to_string(), None, false);
        assert!(malformed_bibliography.bibliography().is_err());
    }

    #[test]
    fn rejects_invalid_index_field_semantics() {
        let invalid_columns = Field::new(r#"INDEX \c 5"#.to_string(), None, false);
        let index = invalid_columns.index().unwrap().unwrap();
        assert!(index.columns().is_err());

        let invalid_sort = Field::new(r#"INDEX \o "radical""#.to_string(), None, false);
        let index = invalid_sort.index().unwrap().unwrap();
        assert!(index.sort_order().is_err());

        let missing_entry = Field::new(r#"XE \b"#.to_string(), None, false);
        assert!(missing_entry.index_entry().is_err());
        let empty_entry = Field::new(r#"XE """#.to_string(), None, false);
        assert!(empty_entry.index_entry().is_err());

        let duplicate_identifier = Field::new(
            r#"XE "topic" \f "first" \f "second""#.to_string(),
            None,
            false,
        );
        let entry = duplicate_identifier.index_entry().unwrap().unwrap();
        assert!(entry.entry_identifier().is_err());
    }

    #[test]
    fn rejects_malformed_toc_switches_and_level_ranges() {
        let non_toc = Field::new("TOCENTRY \\f ignored".to_string(), None, false);
        assert!(!non_toc.is_table_of_contents());
        assert!(non_toc.table_of_contents().unwrap().is_none());

        let dangling = Field::new("TOC \\".to_string(), None, false);
        assert!(dangling.table_of_contents().is_err());
        let unterminated = Field::new(r#"TOC \o "1-3"#.to_string(), None, false);
        assert!(unterminated.table_of_contents().is_err());

        let invalid_levels = Field::new(r#"TOC \o "3-1""#.to_string(), None, false);
        let toc = invalid_levels.table_of_contents().unwrap().unwrap();
        assert!(toc.heading_style_levels().is_err());
        assert!(TableOfContentsLevelRange::new(0, 1).is_err());
        assert!(TableOfContentsLevelRange::new(1, 10).is_err());
    }

    #[test]
    fn rejects_simple_fields_without_instructions() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:fldSimple><w:r><w:t>result</w:t></w:r></w:fldSimple></w:p></w:body></w:document>"#;
        assert!(Field::extract_from_document(xml).is_err());
    }
}
