/// Field support for reading fields from Word documents.
///
/// This module provides types and methods for accessing fields in Word documents.
/// Fields are dynamic content like page numbers, dates, formulas, and cross-references.
use crate::docx::paragraph::decode_xml_reference;
use crate::error::{OoxmlError, Result};
use quick_xml::events::Event;
use quick_xml::{Reader, XmlVersion};

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
        let mut in_instr_text = false;
        let mut in_field_result = false;
        let mut in_result_text = false;
        let mut current_instruction = String::new();
        let mut current_result = String::new();
        let mut current_dirty = false;
        let mut field_depth: i32 = 0;

        loop {
            match reader.read_event() {
                Ok(Event::Empty(e)) if e.local_name().as_ref() == b"t" => {
                    in_result_text = false;
                },
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    match e.local_name().as_ref() {
                        b"fldChar" => {
                            // Field character marks field boundaries
                            let mut fld_char_type = None;
                            let mut dirty = None;

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
                                    dirty = Some(matches!(value.as_ref(), "true" | "1"));
                                }
                            }

                            if let Some(ref char_type) = fld_char_type {
                                match char_type.as_str() {
                                    "begin" => {
                                        // Start of field
                                        field_depth += 1;
                                        if field_depth == 1 {
                                            current_instruction.clear();
                                            current_result.clear();
                                            current_dirty = dirty.unwrap_or(false);
                                            in_instr_text = false;
                                            in_field_result = false;
                                            in_result_text = false;
                                        }
                                    },
                                    "separate"
                                        // Separator between instruction and result
                                        if field_depth == 1 => {
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
                                                fields.push(Field::new(
                                                    current_instruction.trim().to_string(),
                                                    result,
                                                    current_dirty,
                                                ));
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
                        b"t" if in_field_result => in_result_text = true,
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
                },
                Ok(Event::Text(e)) => {
                    let target = if in_instr_text && field_depth == 1 {
                        Some(&mut current_instruction)
                    } else if in_field_result && in_result_text && field_depth == 1 {
                        Some(&mut current_result)
                    } else {
                        None
                    };
                    if let Some(target) = target {
                        let decoded = e
                            .xml_content(XmlVersion::Explicit1_0)
                            .map_err(|error| OoxmlError::Xml(error.to_string()))?;
                        let unescaped = quick_xml::escape::unescape(&decoded)
                            .map_err(|error| OoxmlError::Xml(error.to_string()))?;
                        target.push_str(&unescaped);
                    }
                },
                Ok(Event::CData(e)) => {
                    let target = if in_instr_text && field_depth == 1 {
                        Some(&mut current_instruction)
                    } else if in_field_result && in_result_text && field_depth == 1 {
                        Some(&mut current_result)
                    } else {
                        None
                    };
                    if let Some(target) = target {
                        let decoded = e
                            .xml_content(XmlVersion::Explicit1_0)
                            .map_err(|error| OoxmlError::Xml(error.to_string()))?;
                        target.push_str(&decoded);
                    }
                },
                Ok(Event::GeneralRef(reference)) => {
                    let target = if in_instr_text && field_depth == 1 {
                        Some(&mut current_instruction)
                    } else if in_field_result && in_result_text && field_depth == 1 {
                        Some(&mut current_result)
                    } else {
                        None
                    };
                    if let Some(target) = target {
                        target.push_str(&decode_xml_reference(&reference)?);
                    }
                },
                Ok(Event::End(e)) if e.local_name().as_ref() == b"instrText" => {
                    in_instr_text = false;
                },
                Ok(Event::End(e)) if e.local_name().as_ref() == b"t" => {
                    in_result_text = false;
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
        }

        Ok(fields)
    }
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
}
