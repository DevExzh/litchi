/// Custom XML parts support for Word documents.
///
/// Custom XML parts allow storing arbitrary XML data within a Word document.
use crate::ooxml::error::Result;
use crate::ooxml::opc::part::Part;
use std::collections::HashMap;

/// A custom XML part in a Word document.
///
/// Custom XML parts store arbitrary XML data that can be used for
/// custom applications, metadata, or data binding.
///
/// # Examples
///
/// ```rust,no_run
/// use litchi::ooxml::docx::Package;
///
/// let pkg = Package::open("document.docx")?;
/// let doc = pkg.document()?;
///
/// for custom_xml in doc.custom_xml_parts()? {
///     println!("Custom XML part: {}", custom_xml.id());
///     println!("Content: {}", custom_xml.xml_content());
/// }
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
#[derive(Debug, Clone)]
pub struct CustomXmlPart {
    /// Part ID
    id: String,
    /// XML content
    xml_content: String,
    /// Properties (optional)
    properties: HashMap<String, String>,
}

impl CustomXmlPart {
    /// Create a new CustomXmlPart.
    pub fn new(id: String, xml_content: String, properties: HashMap<String, String>) -> Self {
        Self {
            id,
            xml_content,
            properties,
        }
    }

    /// Get the part ID.
    #[inline]
    pub fn id(&self) -> &str {
        &self.id
    }

    /// Get the XML content.
    #[inline]
    pub fn xml_content(&self) -> &str {
        &self.xml_content
    }

    /// Get the properties.
    #[inline]
    pub fn properties(&self) -> &HashMap<String, String> {
        &self.properties
    }

    /// Get a property by key.
    #[inline]
    pub fn get_property(&self, key: &str) -> Option<&str> {
        self.properties.get(key).map(|s| s.as_str())
    }

    /// Extract custom XML part from a part.
    pub(crate) fn from_part(part: &dyn Part, id: String) -> Result<Self> {
        let xml_content = String::from_utf8_lossy(part.blob()).into_owned();
        Ok(Self {
            id,
            xml_content,
            properties: HashMap::new(),
        })
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_custom_xml_part_creation() {
        let mut props = HashMap::new();
        props.insert("name".to_string(), "test".to_string());

        let part = CustomXmlPart::new(
            "item1".to_string(),
            "<root><data>test</data></root>".to_string(),
            props,
        );

        assert_eq!(part.id(), "item1");
        assert!(part.xml_content().contains("<data>test</data>"));
        assert_eq!(part.get_property("name"), Some("test"));
    }
}
