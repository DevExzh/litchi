/// Custom XML parts support for Word documents.
///
/// Custom XML parts allow storing arbitrary XML data within a Word document.
use crate::error::Result;
use litchi_ooxml_common::custom_xml::Conformance;
use litchi_opc::PackURI;
use litchi_opc::part::Part as OpcPart;
use std::collections::HashMap;

/// High-level parameters for a DOCX Custom XML Data Storage item.
#[derive(Debug, PartialEq, Eq)]
pub struct NewStore {
    pub xml: Vec<u8>,
    pub content_type: String,
    pub id: String,
    pub schemas: Vec<String>,
    pub conformance: Conformance,
}

/// One validated SDT binding occurrence in a Word content-bearing part.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Binding {
    pub source: PackURI,
    pub control_id: u32,
    pub xpath: String,
    pub store_id: String,
    pub prefixes: Option<String>,
}

/// A custom XML part in a Word document.
///
/// Custom XML parts store arbitrary XML data that can be used for
/// custom applications, metadata, or data binding.
///
/// # Examples
///
/// ```rust,no_run
/// use litchi_docx::Package;
///
/// let pkg = Package::open("document.docx")?;
/// let doc = pkg.document()?;
///
/// for custom_xml in doc.custom_xml()? {
///     println!("Custom XML part: {}", custom_xml.id());
///     println!("Content: {}", custom_xml.xml());
/// }
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
#[derive(Debug)]
pub struct Part {
    /// Part ID
    id: String,
    /// XML content
    xml: String,
    /// Properties (optional)
    props: HashMap<String, String>,
}

impl Part {
    /// Create a detached Custom XML part view.
    pub fn new(id: String, xml: String, props: HashMap<String, String>) -> Self {
        Self { id, xml, props }
    }

    /// Get the part ID.
    #[inline]
    pub fn id(&self) -> &str {
        &self.id
    }

    /// Get the XML content.
    #[inline]
    pub fn xml(&self) -> &str {
        &self.xml
    }

    /// Get the properties.
    #[inline]
    pub fn props(&self) -> &HashMap<String, String> {
        &self.props
    }

    /// Get a property by key.
    #[inline]
    pub fn get(&self, key: &str) -> Option<&str> {
        self.props.get(key).map(String::as_str)
    }

    /// Extract custom XML part from a part.
    pub(crate) fn from_part(part: &dyn OpcPart, id: String) -> Result<Self> {
        let xml = std::str::from_utf8(part.blob())
            .map_err(|error| {
                crate::error::Error::Xml(format!(
                    "Custom XML part '{}' is not UTF-8: {error}",
                    part.partname().as_str()
                ))
            })?
            .to_owned();
        Ok(Self {
            id,
            xml,
            props: HashMap::new(),
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

        let part = Part::new(
            "item1".to_string(),
            "<root><data>test</data></root>".to_string(),
            props,
        );

        assert_eq!(part.id(), "item1");
        assert!(part.xml().contains("<data>test</data>"));
        assert_eq!(part.get("name"), Some("test"));
    }
}
