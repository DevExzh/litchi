/// Document variables support for Word documents.
///
/// The collection and standalone WordprocessingML codec live in
/// `litchi-docx`.  This adapter retains the historical host API and keeps OPC
/// part and markup-compatibility concerns at the package boundary.
use crate::error::{OoxmlError, Result};
use litchi_opc::part::Part;

#[cfg(test)]
const MAX_DOCUMENT_VARIABLES: usize = litchi_docx::variables::MAX_DOCUMENT_VARIABLES;
const MAX_DOCUMENT_VARIABLE_XML_BYTES: usize =
    litchi_docx::variables::MAX_DOCUMENT_VARIABLE_XML_BYTES;

/// Deterministic insertion-order collection of document variables.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DocumentVariables {
    inner: litchi_docx::DocumentVariables,
}

impl DocumentVariables {
    /// Create an empty collection.
    pub const fn new() -> Self {
        Self {
            inner: litchi_docx::DocumentVariables::new(),
        }
    }

    /// Get a variable value by its case-sensitive OOXML name.
    pub fn get(&self, name: &str) -> Option<&str> {
        self.inner.get(name)
    }

    /// Check whether a variable exists.
    pub fn contains(&self, name: &str) -> bool {
        self.inner.contains(name)
    }

    /// Return variable names in deterministic insertion order.
    pub fn names(&self) -> Vec<&str> {
        self.inner.names()
    }

    /// Iterate in deterministic insertion order.
    pub fn iter(&self) -> impl Iterator<Item = (&str, &str)> {
        self.inner.iter()
    }

    /// Number of variables.
    pub fn count(&self) -> usize {
        self.inner.count()
    }

    /// Whether the collection is empty.
    pub fn is_empty(&self) -> bool {
        self.inner.is_empty()
    }

    /// Insert or replace a variable without changing an existing entry's order.
    pub fn insert(
        &mut self,
        name: impl Into<String>,
        value: impl Into<String>,
    ) -> Result<Option<String>> {
        self.inner.insert(name, value).map_err(map_docx_error)
    }

    /// Remove a variable while preserving the order of all remaining entries.
    pub fn remove(&mut self, name: &str) -> Option<String> {
        self.inner.remove(name)
    }

    /// Remove all variables.
    pub fn clear(&mut self) {
        self.inner.clear();
    }

    /// Serialize a standalone transitional `w:docVars` element.
    pub fn to_xml(&self) -> Result<String> {
        self.inner.to_xml().map_err(map_docx_error)
    }

    /// Validate the collection for the host settings patcher.
    pub(crate) fn validate(&self) -> Result<()> {
        self.inner.validate().map_err(map_docx_error)
    }

    /// Append `docVar` children for the host settings patcher.
    pub(crate) fn write_entries(&self, xml: &mut String, prefix: &str) {
        self.inner.write_entries(xml, prefix);
    }

    /// Extract variables from a `settings.xml` part after MCE preprocessing.
    pub(crate) fn extract_from_settings_part(part: &dyn Part) -> Result<Self> {
        if part.blob().len() > MAX_DOCUMENT_VARIABLE_XML_BYTES {
            return Err(OoxmlError::InvalidFormat(format!(
                "settings XML exceeds the {MAX_DOCUMENT_VARIABLE_XML_BYTES} byte document-variable limit"
            )));
        }
        let xml = litchi_ooxml_common::mce::process_part(part)?;
        Self::extract_from_xml(xml.as_ref())
    }

    /// Extract variables from a settings XML payload without package access.
    pub(crate) fn extract_from_xml(xml: &[u8]) -> Result<Self> {
        litchi_docx::parse_document_variables(xml)
            .map(Self::from_inner)
            .map_err(map_docx_error)
    }

    fn from_inner(inner: litchi_docx::DocumentVariables) -> Self {
        Self { inner }
    }
}

impl Default for DocumentVariables {
    fn default() -> Self {
        Self::new()
    }
}

fn map_docx_error(error: litchi_docx::Error) -> OoxmlError {
    match error {
        litchi_docx::Error::Opc(error) => OoxmlError::Opc(error),
        litchi_docx::Error::Xml(message) => OoxmlError::Xml(message),
        litchi_docx::Error::Invalid(message) => OoxmlError::InvalidFormat(message),
        litchi_docx::Error::Mce(error) => OoxmlError::from(error),
        litchi_docx::Error::Allocation { resource, source } => {
            OoxmlError::Allocation { resource, source }
        },
        other => OoxmlError::InvalidFormat(other.to_string()),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::PackURI;
    use litchi_opc::constants::content_type as ct;
    use litchi_opc::part::BlobPart;

    #[test]
    fn insertion_order_replace_remove_clear_and_serialization_are_deterministic() {
        let mut variables = DocumentVariables::new();
        assert_eq!(variables.insert("first", "A&B").unwrap(), None);
        assert_eq!(variables.insert("second", "<two>").unwrap(), None);
        assert_eq!(
            variables.insert("first", "updated").unwrap(),
            Some("A&B".into())
        );
        assert_eq!(variables.names(), vec!["first", "second"]);
        assert_eq!(variables.remove("second"), Some("<two>".into()));
        assert_eq!(
            variables.to_xml().unwrap(),
            r#"<w:docVars xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docVar w:name="first" w:val="updated"/></w:docVars>"#
        );
        variables.clear();
        assert!(variables.is_empty());
    }

    #[test]
    fn enforces_word_name_value_and_count_boundaries() {
        let mut variables = DocumentVariables::new();
        assert!(variables.insert("", "value").is_err());
        assert!(variables.insert("名".repeat(255), "").is_ok());
        assert!(variables.insert("名".repeat(256), "value").is_err());
        assert!(variables.insert("maximum", "x".repeat(65_280)).is_ok());
        assert!(variables.insert("too-long", "x".repeat(65_281)).is_err());

        let mut count = DocumentVariables::new();
        for index in 0..MAX_DOCUMENT_VARIABLES {
            count.insert(format!("v{index}"), "x").unwrap();
        }
        assert!(count.insert("overflow", "x").is_err());
    }

    #[test]
    fn parses_strict_direct_children_and_decodes_attributes() {
        let xml = br#"<s:settings xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main" xmlns:false="urn:false"><false:docVars><false:docVar false:name="spoof" false:val="bad"/></false:docVars><s:docVars><s:docVar s:name="Company &amp; Team" s:val="A &lt; B &amp;&amp; C &quot;yes&quot;"/><s:docVar s:name="empty" s:val=""/></s:docVars></s:settings>"#;
        let variables = DocumentVariables::extract_from_xml(xml).unwrap();
        assert_eq!(variables.count(), 2);
        assert_eq!(variables.get("Company & Team"), Some("A < B && C \"yes\""));
        assert_eq!(variables.get("empty"), Some(""));
        assert!(!variables.contains("spoof"));
    }

    #[test]
    fn rejects_missing_duplicate_nested_and_oversized_input() {
        let wrap = |content: &str| {
            format!(
                r#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">{content}</w:settings>"#
            )
        };
        for content in [
            r#"<w:docVars><w:docVar w:val="x"/></w:docVars>"#,
            r#"<w:docVars><w:docVar w:name="x"/></w:docVars>"#,
            r#"<w:docVars/><w:docVars/>"#,
            r#"<w:docVars><w:docVar w:name="x" w:val="1"/><w:docVar w:name="x" w:val="2"/></w:docVars>"#,
            r#"<w:docVars><w:docVar w:name="x" w:val="1"><w:docVar w:name="y" w:val="2"/></w:docVar></w:docVars>"#,
            r#"<w:docVar w:name="x" w:val="1"/>"#,
        ] {
            assert!(DocumentVariables::extract_from_xml(wrap(content).as_bytes()).is_err());
        }
        let oversized = vec![b' '; MAX_DOCUMENT_VARIABLE_XML_BYTES + 1];
        assert!(DocumentVariables::extract_from_xml(&oversized).is_err());
    }

    #[test]
    fn mce_selects_fallback_document_variables() {
        let xml = br#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w99="urn:unsupported" mc:Ignorable="w99"><mc:AlternateContent><mc:Choice Requires="w99"><w:docVars><w:docVar w:name="choice" w:val="ignored"/></w:docVars></mc:Choice><mc:Fallback><w:docVars><w:docVar w:name="fallback" w:val="selected"/></w:docVars></mc:Fallback></mc:AlternateContent></w:settings>"#;
        let part = BlobPart::new(
            PackURI::new("/word/settings.xml").unwrap(),
            ct::WML_SETTINGS.to_owned(),
            xml.to_vec(),
        );
        let variables = DocumentVariables::extract_from_settings_part(&part).unwrap();
        assert_eq!(variables.get("fallback"), Some("selected"));
        assert!(!variables.contains("choice"));
    }
}
