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

    /// Check whether this is a mail-merge field.
    pub fn is_merge_field(&self) -> bool {
        self.field_type().eq_ignore_ascii_case("MERGEFIELD")
    }

    /// Return the data-source column name from a `MERGEFIELD` instruction.
    ///
    /// Both unquoted names (`MERGEFIELD FirstName`) and quoted names containing
    /// spaces (`MERGEFIELD "Full Name"`) are supported. Field switches following
    /// the name are excluded.
    pub fn merge_field_name(&self) -> Option<&str> {
        if !self.is_merge_field() {
            return None;
        }
        let instruction = self.instruction.trim_start();
        let field_type_end = instruction
            .find(char::is_whitespace)
            .unwrap_or(instruction.len());
        let remainder = instruction[field_type_end..].trim_start();
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
            parse_field_operand_and_switches(field.instruction(), "CITATION")?
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

fn parse_field_switches_from_characters(
    mut characters: &mut std::iter::Peekable<std::str::Chars<'_>>,
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
                Some(parse_field_quoted_argument(&mut characters, field_type)?)
            },
            Some(_) => Some(parse_field_unquoted_argument(&mut characters)),
        };
        if switches.len() >= MAX_FIELD_SWITCHES {
            return Err(OoxmlError::InvalidFormat(format!(
                "{field_type} field exceeds {MAX_FIELD_SWITCHES} switches"
            )));
        }
        switches.push(FieldSwitch {
            name: name.to_ascii_lowercase(),
            argument,
        });
    }
    Ok(switches)
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
            <w:r><w:instrText>BIBLIOGRAPHY \l 1033 \f &quot;References&quot;</w:instrText></w:r>
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

        let bibliography = fields[1].bibliography().unwrap().unwrap();
        assert_eq!(bibliography.cached_result(), Some("Doe. Example work."));
        assert!(bibliography.is_dirty());
        assert!(bibliography.is_locked());
        assert_eq!(bibliography.switches()[0].name(), 'l');
        assert_eq!(bibliography.switches()[0].argument(), Some("1033"));
        assert!(bibliography.has_switch('f'));
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
