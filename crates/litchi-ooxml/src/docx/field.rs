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
