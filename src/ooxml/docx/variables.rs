/// Document variables support for Word documents.
///
/// Document variables are custom properties that can be referenced
/// in fields and used for mail merge and other operations.
use crate::ooxml::error::{OoxmlError, Result};
use crate::ooxml::opc::part::Part;
use quick_xml::Reader;
use quick_xml::events::Event;
use std::collections::HashMap;

/// Document variables collection.
///
/// Variables are name-value pairs stored in the document settings
/// that can be referenced by fields.
///
/// # Examples
///
/// ```rust,no_run
/// use litchi::ooxml::docx::Package;
///
/// let pkg = Package::open("document.docx")?;
/// let doc = pkg.document()?;
///
/// if let Some(vars) = doc.document_variables()? {
///     for (name, value) in vars.iter() {
///         println!("{} = {}", name, value);
///     }
/// }
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
#[derive(Debug, Clone)]
pub struct DocumentVariables {
    /// Variable name-value pairs
    variables: HashMap<String, String>,
}

impl DocumentVariables {
    /// Create a new empty DocumentVariables.
    pub fn new() -> Self {
        Self {
            variables: HashMap::new(),
        }
    }

    /// Get a variable value by name.
    #[inline]
    pub fn get(&self, name: &str) -> Option<&str> {
        self.variables.get(name).map(|s| s.as_str())
    }

    /// Check if a variable exists.
    #[inline]
    pub fn contains(&self, name: &str) -> bool {
        self.variables.contains_key(name)
    }

    /// Get all variable names.
    pub fn names(&self) -> Vec<&str> {
        self.variables.keys().map(|s| s.as_str()).collect()
    }

    /// Iterate over all variables.
    pub fn iter(&self) -> impl Iterator<Item = (&str, &str)> {
        self.variables.iter().map(|(k, v)| (k.as_str(), v.as_str()))
    }

    /// Get the number of variables.
    #[inline]
    pub fn count(&self) -> usize {
        self.variables.len()
    }

    /// Check if there are no variables.
    #[inline]
    pub fn is_empty(&self) -> bool {
        self.variables.is_empty()
    }

    /// Extract variables from a settings.xml part.
    ///
    /// Variables are stored in the `<w:docVars>` section of settings.
    pub(crate) fn extract_from_settings_part(part: &dyn Part) -> Result<Self> {
        let xml_bytes = part.blob();
        let mut reader = Reader::from_reader(xml_bytes);
        reader.config_mut().trim_text(true);

        let mut variables = HashMap::new();

        loop {
            match reader.read_event() {
                Ok(Event::Empty(e)) | Ok(Event::Start(e)) => {
                    if e.local_name().as_ref() == b"docVar" {
                        let mut name = None;
                        let mut value = None;

                        for attr in e.attributes().flatten() {
                            match attr.key.local_name().as_ref() {
                                b"name" => {
                                    name = Some(String::from_utf8_lossy(&attr.value).into_owned());
                                },
                                b"val" => {
                                    value = Some(String::from_utf8_lossy(&attr.value).into_owned());
                                },
                                _ => {},
                            }
                        }

                        if let (Some(n), Some(v)) = (name, value) {
                            variables.insert(n, v);
                        }
                    }
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
        }

        Ok(Self { variables })
    }
}

impl Default for DocumentVariables {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_document_variables_creation() {
        let vars = DocumentVariables::new();
        assert!(vars.is_empty());
        assert_eq!(vars.count(), 0);
    }

    #[test]
    fn test_document_variables_operations() {
        let mut vars = DocumentVariables::new();
        vars.variables
            .insert("company".to_string(), "Acme Corp".to_string());
        vars.variables
            .insert("year".to_string(), "2024".to_string());

        assert!(!vars.is_empty());
        assert_eq!(vars.count(), 2);
        assert_eq!(vars.get("company"), Some("Acme Corp"));
        assert_eq!(vars.get("year"), Some("2024"));
        assert!(vars.contains("company"));
        assert!(!vars.contains("nonexistent"));
    }
}
