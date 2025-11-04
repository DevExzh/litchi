/// Field writer support for DOCX documents.
use crate::ooxml::error::Result;
use std::fmt::Write as FmtWrite;

/// A mutable field in a Word document.
///
/// Fields are dynamic content placeholders such as page numbers, dates, cross-references, etc.
#[derive(Debug, Clone)]
pub struct MutableField {
    /// Field instruction (e.g., "PAGE", "DATE", "REF MyBookmark")
    instruction: String,
    /// Field result (optional, the displayed value)
    result: Option<String>,
    /// Whether the field is dirty (needs update)
    dirty: bool,
}

impl MutableField {
    /// Create a new field.
    ///
    /// # Arguments
    ///
    /// * `instruction` - The field instruction (e.g., "PAGE", "DATE \\@ \"MMMM d, yyyy\"")
    pub fn new(instruction: String) -> Self {
        Self {
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
        Self {
            instruction,
            result: Some(result),
            dirty: false,
        }
    }

    /// Get the field instruction.
    #[inline]
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Set the field instruction.
    pub fn set_instruction(&mut self, instruction: String) {
        self.instruction = instruction;
        self.dirty = true;
    }

    /// Get the field result.
    #[inline]
    pub fn result(&self) -> Option<&str> {
        self.result.as_deref()
    }

    /// Set the field result.
    pub fn set_result(&mut self, result: Option<String>) {
        self.result = result;
        self.dirty = false;
    }

    /// Check if the field is dirty (needs update).
    #[inline]
    pub fn is_dirty(&self) -> bool {
        self.dirty
    }

    /// Mark the field as dirty.
    pub fn mark_dirty(&mut self) {
        self.dirty = true;
    }

    /// Generate XML for this field.
    #[allow(dead_code)]
    pub(crate) fn to_xml(&self) -> Result<String> {
        let mut xml = String::with_capacity(256);

        // Field begin
        xml.push_str(r#"<w:fldChar w:fldCharType="begin"/>"#);

        // Field instruction
        write!(
            &mut xml,
            "</w:r><w:r><w:instrText>{}</w:instrText>",
            escape_xml(&self.instruction)
        )?;

        // Separate run
        if self.dirty {
            xml.push_str(r#"</w:r><w:r><w:fldChar w:fldCharType="separate" w:dirty="true"/>"#);
        } else {
            xml.push_str(r#"</w:r><w:r><w:fldChar w:fldCharType="separate"/>"#);
        }

        // Field result
        if let Some(result) = &self.result {
            write!(&mut xml, "</w:r><w:r><w:t>{}</w:t>", escape_xml(result))?;
        }

        // Field end
        xml.push_str(r#"</w:r><w:r><w:fldChar w:fldCharType="end"/>"#);

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
