/// Field support for reading fields from Word documents.
///
/// This module provides types and methods for accessing fields in Word documents.
/// Fields are dynamic content like page numbers, dates, formulas, and cross-references.
use crate::ooxml::error::{OoxmlError, Result};
use quick_xml::Reader;
use quick_xml::events::Event;

/// A field in a Word document.
///
/// Represents a field instruction like `PAGE`, `DATE`, `REF`, etc.
/// Fields are dynamic content that can be updated.
///
/// # Examples
///
/// ```rust,no_run
/// use litchi::ooxml::docx::Package;
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
    /// use litchi::ooxml::docx::Field;
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
        reader.config_mut().trim_text(true);

        let mut fields = Vec::new();
        let mut in_instr_text = false;
        let mut in_field_result = false;
        let mut current_instruction = String::new();
        let mut current_result = String::new();
        let mut current_dirty = false;
        let mut field_depth: i32 = 0;
        let mut buf = Vec::with_capacity(512);

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    match e.local_name().as_ref() {
                        b"fldChar" => {
                            // Field character marks field boundaries
                            let mut fld_char_type = None;

                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"fldCharType" {
                                    fld_char_type =
                                        Some(String::from_utf8_lossy(&attr.value).into_owned());
                                }
                                if attr.key.local_name().as_ref() == b"dirty" {
                                    let dirty_val = String::from_utf8_lossy(&attr.value);
                                    current_dirty = dirty_val == "true" || dirty_val == "1";
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
                                            current_dirty = false;
                                            in_instr_text = false;
                                            in_field_result = false;
                                        }
                                    },
                                    "separate" => {
                                        // Separator between instruction and result
                                        if field_depth == 1 {
                                            in_instr_text = false;
                                            in_field_result = true;
                                        }
                                    },
                                    "end" => {
                                        // End of field
                                        if field_depth == 1 {
                                            in_field_result = false;
                                            in_instr_text = false;

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
                        b"instrText" => {
                            // Field instruction text
                            if field_depth > 0 {
                                in_instr_text = true;
                            }
                        },
                        b"t" => {
                            // Text element - could be part of field result
                            // Will be handled in Text event
                        },
                        _ => {},
                    }
                },
                Ok(Event::Text(e)) => {
                    if in_instr_text && field_depth == 1 {
                        // Accumulate instruction text
                        let text = unsafe { std::str::from_utf8_unchecked(e.as_ref()) };
                        current_instruction.push_str(text);
                    } else if in_field_result && field_depth == 1 {
                        // Accumulate result text
                        let text = unsafe { std::str::from_utf8_unchecked(e.as_ref()) };
                        current_result.push_str(text);
                    }
                },
                Ok(Event::End(e)) => {
                    if e.local_name().as_ref() == b"instrText" {
                        in_instr_text = false;
                    }
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
            buf.clear();
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
}
