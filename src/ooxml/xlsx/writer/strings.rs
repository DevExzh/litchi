/// Shared strings table for XLSX workbooks.
use crate::sheet::Result as SheetResult;
use std::collections::HashMap;
use std::fmt::Write as FmtWrite;

/// Escape XML special characters.
fn escape_xml(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

/// Mutable shared strings table.
///
/// Excel stores frequently used strings in a shared table to reduce file size.
/// This structure manages the collection of unique strings and their indices.
#[derive(Debug)]
pub struct MutableSharedStrings {
    /// List of unique strings
    pub(crate) strings: Vec<String>,
    /// Map from string to index for fast lookup
    pub(crate) string_to_index: HashMap<String, usize>,
}

impl MutableSharedStrings {
    /// Create a new empty shared strings table.
    pub fn new() -> Self {
        Self {
            strings: Vec::new(),
            string_to_index: HashMap::new(),
        }
    }

    /// Add a string to the shared strings table and return its index.
    ///
    /// If the string already exists, returns the existing index.
    pub fn add_string(&mut self, s: &str) -> usize {
        if let Some(&index) = self.string_to_index.get(s) {
            index
        } else {
            let index = self.strings.len();
            self.strings.push(s.to_string());
            self.string_to_index.insert(s.to_string(), index);
            index
        }
    }

    /// Get the number of unique strings.
    pub fn count(&self) -> usize {
        self.strings.len()
    }

    /// Serialize the shared strings table to XML.
    pub fn to_xml(&self) -> SheetResult<String> {
        let mut xml = String::with_capacity(4096);

        xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);

        write!(
            xml,
            r#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="{}" uniqueCount="{}">"#,
            self.strings.len(),
            self.strings.len()
        )
        .map_err(|e| format!("XML write error: {}", e))?;

        for s in &self.strings {
            write!(xml, "<si><t>{}</t></si>", escape_xml(s))
                .map_err(|e| format!("XML write error: {}", e))?;
        }

        xml.push_str("</sst>");

        Ok(xml)
    }
}

impl Default for MutableSharedStrings {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_shared_strings() {
        let mut ss = MutableSharedStrings::new();
        let idx1 = ss.add_string("Hello");
        let idx2 = ss.add_string("World");
        let idx3 = ss.add_string("Hello"); // Duplicate

        assert_eq!(idx1, 0);
        assert_eq!(idx2, 1);
        assert_eq!(idx3, 0); // Same as first "Hello"
        assert_eq!(ss.count(), 2);
    }
}
