/// Field support for reading fields from Word documents.
///
/// This module provides types and methods for accessing fields in Word documents.
/// Fields are dynamic content like page numbers, dates, formulas, and cross-references.
use crate::common::xml::decode_xml_reference;
use crate::error::{OoxmlError, Result};
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::{Reader, XmlVersion};

const MAX_TOC_SWITCHES: usize = 64;

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

    /// Check whether this is a `TOC` (Table of Contents) field.
    ///
    /// The field's cached result remains data only; calling this method never
    /// recalculates the table of contents or follows any hyperlinks in it.
    pub fn is_table_of_contents(&self) -> bool {
        toc_instruction_remainder(&self.instruction).is_some()
    }

    /// Parse this field as an inert typed table-of-contents field.
    ///
    /// Returns `Ok(None)` for non-`TOC` fields. The returned model preserves
    /// the instruction, cached result, dirty/lock state, and field switches;
    /// it never evaluates the field or refreshes its cached content.
    pub fn table_of_contents(&self) -> Result<Option<TableOfContentsField>> {
        TableOfContentsField::from_field(self)
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

/// One lexical switch in a `TOC` field instruction.
///
/// Switch names are normalized to ASCII lowercase. Quoted and unquoted
/// arguments are decoded into their logical text, while the complete original
/// instruction remains available through [`TableOfContentsField::instruction`].
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TableOfContentsSwitch {
    name: char,
    argument: Option<String>,
}

impl TableOfContentsSwitch {
    /// Return the switch character, without its leading backslash.
    pub fn name(&self) -> char {
        self.name
    }

    /// Return the optional argument supplied to this switch.
    pub fn argument(&self) -> Option<&str> {
        self.argument.as_deref()
    }
}

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
    switches: Vec<TableOfContentsSwitch>,
}

impl TableOfContentsField {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some(switches) = parse_toc_switches(field.instruction())? else {
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
    pub fn switches(&self) -> &[TableOfContentsSwitch] {
        &self.switches
    }

    /// Check whether a case-insensitive ASCII switch appears in this field.
    pub fn has_switch(&self, name: char) -> bool {
        self.switches
            .iter()
            .any(|switch| switch.name.eq_ignore_ascii_case(&name))
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

fn toc_instruction_remainder(instruction: &str) -> Option<&str> {
    let instruction = instruction.trim_start();
    let field_type = instruction.get(..3)?;
    let remainder = instruction.get(3..)?;
    if !field_type.eq_ignore_ascii_case("TOC") {
        return None;
    }
    match remainder.chars().next() {
        None | Some('\\') => Some(remainder),
        Some(character) if character.is_whitespace() => Some(remainder),
        Some(_) => None,
    }
}

fn parse_toc_switches(instruction: &str) -> Result<Option<Vec<TableOfContentsSwitch>>> {
    let Some(remainder) = toc_instruction_remainder(instruction) else {
        return Ok(None);
    };
    let mut characters = remainder.chars().peekable();
    let mut switches = Vec::new();
    loop {
        while characters
            .peek()
            .is_some_and(|character| character.is_whitespace())
        {
            characters.next();
        }
        let Some(character) = characters.next() else {
            break;
        };
        if character != '\\' {
            return Err(OoxmlError::InvalidFormat(
                "TOC field contains text outside a field switch".to_string(),
            ));
        }
        let name = characters.next().ok_or_else(|| {
            OoxmlError::InvalidFormat("TOC field ends with a switch introducer".to_string())
        })?;
        if name == '\\' || name.is_whitespace() {
            return Err(OoxmlError::InvalidFormat(
                "TOC field has an invalid switch name".to_string(),
            ));
        }
        while characters
            .peek()
            .is_some_and(|character| character.is_whitespace())
        {
            characters.next();
        }
        let argument = match characters.peek().copied() {
            None | Some('\\') => None,
            Some('"') => {
                characters.next();
                Some(parse_toc_quoted_argument(&mut characters)?)
            },
            Some(_) => Some(parse_toc_unquoted_argument(&mut characters)),
        };
        if switches.len() >= MAX_TOC_SWITCHES {
            return Err(OoxmlError::InvalidFormat(format!(
                "TOC field exceeds {MAX_TOC_SWITCHES} switches"
            )));
        }
        switches.push(TableOfContentsSwitch {
            name: name.to_ascii_lowercase(),
            argument,
        });
    }
    Ok(Some(switches))
}

fn parse_toc_quoted_argument(
    characters: &mut std::iter::Peekable<std::str::Chars<'_>>,
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
                    return Err(OoxmlError::InvalidFormat(
                        "TOC quoted switch argument has trailing text".to_string(),
                    ));
                }
                return Ok(argument);
            },
            _ => argument.push(character),
        }
    }
    Err(OoxmlError::InvalidFormat(
        "TOC field has an unterminated quoted switch argument".to_string(),
    ))
}

fn parse_toc_unquoted_argument(
    characters: &mut std::iter::Peekable<std::str::Chars<'_>>,
) -> String {
    let mut argument = String::new();
    while characters
        .peek()
        .is_some_and(|character| !character.is_whitespace() && *character != '\\')
    {
        argument.push(characters.next().expect("checked TOC argument character"));
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
