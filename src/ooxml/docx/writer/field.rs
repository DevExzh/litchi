/// Field writer support for DOCX documents.
use crate::ooxml::error::Result;
use std::fmt::Write as FmtWrite;

/// A mutable field in a Word document.
///
/// Fields are dynamic content placeholders such as page numbers, dates, cross-references, etc.
///
/// This enum supports both complete fields and individual field characters for complex field structures.
#[derive(Debug, Clone)]
pub enum MutableField {
    /// A complete field with instruction and optional result
    Complete {
        /// Field instruction (e.g., "PAGE", "DATE", "REF MyBookmark")
        instruction: String,
        /// Field result (optional, the displayed value)
        result: Option<String>,
        /// Whether the field is dirty (needs update)
        dirty: bool,
    },
    /// Field begin character
    Begin,
    /// Field instruction text
    Instruction(String),
    /// Field separate character
    Separate {
        /// Whether the field is dirty
        dirty: bool,
    },
    /// Field end character
    End,
}

impl MutableField {
    /// Create a new complete field.
    ///
    /// # Arguments
    ///
    /// * `instruction` - The field instruction (e.g., "PAGE", "DATE \\@ \"MMMM d, yyyy\"")
    pub fn new(instruction: String) -> Self {
        Self::Complete {
            instruction,
            result: None,
            dirty: true,
        }
    }

    /// Create a field with a result value.
    ///
    /// # Arguments
    ///
    /// * `instruction` - The field instruction
    /// * `result` - The current result value
    pub fn with_result(instruction: String, result: String) -> Self {
        Self::Complete {
            instruction,
            result: Some(result),
            dirty: false,
        }
    }

    /// Create a field begin character.
    pub fn begin() -> Self {
        Self::Begin
    }

    /// Create a field instruction.
    pub fn instruction(text: String) -> Self {
        Self::Instruction(text)
    }

    /// Create a field separate character.
    pub fn separate() -> Self {
        Self::Separate { dirty: false }
    }

    /// Create a field separate character marked as dirty.
    pub fn separate_dirty() -> Self {
        Self::Separate { dirty: true }
    }

    /// Create a field end character.
    pub fn end() -> Self {
        Self::End
    }

    /// Get the field instruction (for Complete fields only).
    pub fn get_instruction(&self) -> Option<&str> {
        match self {
            Self::Complete { instruction, .. } => Some(instruction),
            Self::Instruction(text) => Some(text),
            _ => None,
        }
    }

    /// Set the field instruction (for Complete fields only).
    pub fn set_instruction(&mut self, new_instruction: String) {
        if let Self::Complete {
            instruction, dirty, ..
        } = self
        {
            *instruction = new_instruction;
            *dirty = true;
        }
    }

    /// Get the field result (for Complete fields only).
    pub fn get_result(&self) -> Option<&str> {
        match self {
            Self::Complete { result, .. } => result.as_deref(),
            _ => None,
        }
    }

    /// Set the field result (for Complete fields only).
    pub fn set_result(&mut self, new_result: Option<String>) {
        if let Self::Complete { result, dirty, .. } = self {
            *result = new_result;
            *dirty = false;
        }
    }

    /// Check if the field is dirty (needs update).
    pub fn is_dirty(&self) -> bool {
        match self {
            Self::Complete { dirty, .. } | Self::Separate { dirty } => *dirty,
            _ => false,
        }
    }

    /// Mark the field as dirty (for Complete fields only).
    pub fn mark_dirty(&mut self) {
        match self {
            Self::Complete { dirty, .. } | Self::Separate { dirty } => *dirty = true,
            _ => {},
        }
    }

    /// Generate XML for this field.
    #[allow(dead_code)]
    pub(crate) fn to_xml(&self) -> Result<String> {
        let mut xml = String::with_capacity(256);

        match self {
            Self::Complete {
                instruction,
                result,
                dirty,
            } => {
                // Field begin
                xml.push_str(r#"<w:fldChar w:fldCharType="begin"/>"#);

                // Field instruction
                write!(
                    &mut xml,
                    "</w:r><w:r><w:instrText>{}</w:instrText>",
                    escape_xml(instruction)
                )?;

                // Separate run
                if *dirty {
                    xml.push_str(
                        r#"</w:r><w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/>"#,
                    );
                } else {
                    xml.push_str(r#"</w:r><w:r><w:fldChar w:fldCharType="separate"/>"#);
                }

                // Field result
                if let Some(res) = result {
                    write!(&mut xml, "</w:r><w:r><w:t>{}</w:t>", escape_xml(res))?;
                }

                // Field end
                xml.push_str(r#"</w:r><w:r><w:fldChar w:fldCharType="end"/>"#);
            },
            Self::Begin => {
                xml.push_str(r#"<w:fldChar w:fldCharType="begin"/>"#);
            },
            Self::Instruction(text) => {
                write!(
                    &mut xml,
                    r#"<w:instrText xml:space="preserve">{}</w:instrText>"#,
                    escape_xml(text)
                )?;
            },
            Self::Separate { dirty } => {
                if *dirty {
                    xml.push_str(r#"<w:fldChar w:fldCharType="separate" w:dirty="true"/>"#);
                } else {
                    xml.push_str(r#"<w:fldChar w:fldCharType="separate"/>"#);
                }
            },
            Self::End => {
                xml.push_str(r#"<w:fldChar w:fldCharType="end"/>"#);
            },
        }

        Ok(xml)
    }

    /// Common field factory methods
    /// Create a PAGE field (page number).
    pub fn page() -> Self {
        Self::new("PAGE".to_string())
    }

    /// Create a NUMPAGES field (total page count).
    pub fn num_pages() -> Self {
        Self::new("NUMPAGES".to_string())
    }

    /// Create a DATE field with optional format.
    ///
    /// # Arguments
    ///
    /// * `format` - Optional date format string (e.g., "MMMM d, yyyy")
    pub fn date(format: Option<&str>) -> Self {
        let instruction = if let Some(fmt) = format {
            format!(r#"DATE \@ "{}""#, fmt)
        } else {
            "DATE".to_string()
        };
        Self::new(instruction)
    }

    /// Create a TIME field with optional format.
    pub fn time(format: Option<&str>) -> Self {
        let instruction = if let Some(fmt) = format {
            format!(r#"TIME \@ "{}""#, fmt)
        } else {
            "TIME".to_string()
        };
        Self::new(instruction)
    }

    /// Create a REF field (cross-reference to a bookmark).
    ///
    /// # Arguments
    ///
    /// * `bookmark_name` - Name of the bookmark to reference
    pub fn reference(bookmark_name: &str) -> Self {
        Self::new(format!("REF {}", bookmark_name))
    }

    /// Create a HYPERLINK field.
    ///
    /// # Arguments
    ///
    /// * `url` - The URL to link to
    pub fn hyperlink(url: &str) -> Self {
        Self::new(format!(r#"HYPERLINK "{}""#, url))
    }

    /// Create a TOC (Table of Contents) field.
    ///
    /// # Arguments
    ///
    /// * `instruction` - The complete TOC field instruction (e.g., `TOC \o "1-3" \h \z`)
    /// * `placeholder_text` - Optional placeholder text to display before field update
    pub fn toc(instruction: String, placeholder_text: Option<String>) -> Self {
        Self::Complete {
            instruction,
            result: placeholder_text,
            dirty: true,
        }
    }
}

/// Escape XML special characters.
#[allow(dead_code)]
fn escape_xml(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_field_creation() {
        let field = MutableField::new("PAGE".to_string());
        assert_eq!(field.instruction(), "PAGE");
        assert!(field.is_dirty());
        assert!(field.result().is_none());
    }

    #[test]
    fn test_field_with_result() {
        let field = MutableField::with_result("PAGE".to_string(), "5".to_string());
        assert_eq!(field.instruction(), "PAGE");
        assert_eq!(field.result(), Some("5"));
        assert!(!field.is_dirty());
    }

    #[test]
    fn test_field_factories() {
        let page = MutableField::page();
        assert_eq!(page.instruction(), "PAGE");

        let date = MutableField::date(Some("MMMM d, yyyy"));
        assert!(date.instruction().contains("DATE"));
        assert!(date.instruction().contains("MMMM d, yyyy"));

        let ref_field = MutableField::reference("MyBookmark");
        assert!(ref_field.instruction().contains("REF MyBookmark"));
    }

    #[test]
    fn test_field_xml() {
        let mut field = MutableField::with_result("PAGE".to_string(), "1".to_string());
        field.mark_dirty();

        let xml = field.to_xml().unwrap();
        assert!(xml.contains("fldCharType=\"begin\""));
        assert!(xml.contains("instrText"));
        assert!(xml.contains("PAGE"));
        assert!(xml.contains("fldCharType=\"separate\""));
        assert!(xml.contains("dirty=\"true\""));
        assert!(xml.contains("fldCharType=\"end\""));
    }
}
