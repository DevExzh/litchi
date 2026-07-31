/// Document variables support for Word documents.
///
/// Document variables are inert name-value pairs stored in `settings.xml` and
/// referenced by fields such as `DOCVARIABLE`.
use crate::docx::namespace::{is_wordprocessing_namespace, word_attribute_value};
use crate::error::{OoxmlError, Result};
use litchi_opc::part::Part;
use quick_xml::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

const MAX_DOCUMENT_VARIABLES: usize = 4096;
const MAX_DOCUMENT_VARIABLE_XML_BYTES: usize = 8 * 1024 * 1024;
const MAX_DOCUMENT_VARIABLE_DEPTH: usize = 64;
const MAX_DOCUMENT_VARIABLE_NAME_CHARS: usize = 255;
const MAX_DOCUMENT_VARIABLE_VALUE_CHARS: usize = 65_280;

/// Deterministic insertion-order collection of document variables.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DocumentVariables {
    variables: Vec<(String, String)>,
}

impl DocumentVariables {
    /// Create an empty collection.
    pub const fn new() -> Self {
        Self {
            variables: Vec::new(),
        }
    }

    /// Get a variable value by its case-sensitive OOXML name.
    pub fn get(&self, name: &str) -> Option<&str> {
        self.variables
            .iter()
            .find(|(candidate, _)| candidate == name)
            .map(|(_, value)| value.as_str())
    }

    /// Check whether a variable exists.
    pub fn contains(&self, name: &str) -> bool {
        self.get(name).is_some()
    }

    /// Return variable names in deterministic insertion order.
    pub fn names(&self) -> Vec<&str> {
        self.variables
            .iter()
            .map(|(name, _)| name.as_str())
            .collect()
    }

    /// Iterate in deterministic insertion order.
    pub fn iter(&self) -> impl Iterator<Item = (&str, &str)> {
        self.variables
            .iter()
            .map(|(name, value)| (name.as_str(), value.as_str()))
    }

    /// Number of variables.
    pub fn count(&self) -> usize {
        self.variables.len()
    }

    /// Whether the collection is empty.
    pub fn is_empty(&self) -> bool {
        self.variables.is_empty()
    }

    /// Insert or replace a variable without changing an existing entry's order.
    pub fn insert(
        &mut self,
        name: impl Into<String>,
        value: impl Into<String>,
    ) -> Result<Option<String>> {
        let name = name.into();
        let value = value.into();
        validate_document_variable(&name, &value)?;
        if let Some((_, existing)) = self
            .variables
            .iter_mut()
            .find(|(candidate, _)| candidate == &name)
        {
            return Ok(Some(std::mem::replace(existing, value)));
        }
        if self.variables.len() >= MAX_DOCUMENT_VARIABLES {
            return Err(OoxmlError::InvalidFormat(format!(
                "document variables exceed the {MAX_DOCUMENT_VARIABLES} entry limit"
            )));
        }
        self.variables.push((name, value));
        Ok(None)
    }

    /// Remove a variable while preserving the order of all remaining entries.
    pub fn remove(&mut self, name: &str) -> Option<String> {
        let index = self
            .variables
            .iter()
            .position(|(candidate, _)| candidate == name)?;
        Some(self.variables.remove(index).1)
    }

    /// Remove all variables.
    pub fn clear(&mut self) {
        self.variables.clear();
    }

    /// Serialize a standalone transitional `w:docVars` element.
    pub fn to_xml(&self) -> Result<String> {
        self.validate()?;
        let mut xml = String::new();
        xml.push_str(
            r#"<w:docVars xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">"#,
        );
        self.write_entries(&mut xml, "w");
        xml.push_str("</w:docVars>");
        Ok(xml)
    }

    pub(crate) fn validate(&self) -> Result<()> {
        if self.variables.len() > MAX_DOCUMENT_VARIABLES {
            return Err(OoxmlError::InvalidFormat(format!(
                "document variables exceed the {MAX_DOCUMENT_VARIABLES} entry limit"
            )));
        }
        for (name, value) in &self.variables {
            validate_document_variable(name, value)?;
        }
        Ok(())
    }

    pub(crate) fn write_entries(&self, xml: &mut String, prefix: &str) {
        for (name, value) in &self.variables {
            xml.push('<');
            if !prefix.is_empty() {
                xml.push_str(prefix);
                xml.push(':');
            }
            xml.push_str("docVar");
            xml.push(' ');
            if !prefix.is_empty() {
                xml.push_str(prefix);
                xml.push(':');
            }
            xml.push_str("name=\"");
            escape_attribute(xml, name);
            xml.push_str("\" ");
            if !prefix.is_empty() {
                xml.push_str(prefix);
                xml.push(':');
            }
            xml.push_str("val=\"");
            escape_attribute(xml, value);
            xml.push_str("\"/>");
        }
    }

    /// Extract variables from a `settings.xml` part without evaluating fields.
    pub(crate) fn extract_from_settings_part(part: &dyn Part) -> Result<Self> {
        if part.blob().len() > MAX_DOCUMENT_VARIABLE_XML_BYTES {
            return Err(OoxmlError::InvalidFormat(format!(
                "settings XML exceeds the {MAX_DOCUMENT_VARIABLE_XML_BYTES} byte document-variable limit"
            )));
        }
        let xml = litchi_ooxml_common::mce::process_part(part)?;
        Self::extract_from_xml(xml.as_ref())
    }

    pub(crate) fn extract_from_xml(xml: &[u8]) -> Result<Self> {
        if xml.len() > MAX_DOCUMENT_VARIABLE_XML_BYTES {
            return Err(OoxmlError::InvalidFormat(format!(
                "settings XML exceeds the {MAX_DOCUMENT_VARIABLE_XML_BYTES} byte document-variable limit"
            )));
        }
        let mut reader = NsReader::from_reader(xml);
        let mut variables = Self::new();
        let mut depth = 0usize;
        let mut saw_root = false;
        let mut saw_doc_vars = false;
        let mut doc_vars_depth = None;
        let mut open_doc_var_depth = None;

        loop {
            let decoder = reader.decoder();
            let event = reader
                .read_event()
                .map_err(|error| OoxmlError::Xml(error.to_string()))?
                .into_owned();
            let resolver = reader.resolver().clone();
            let (namespace, event) = resolver.resolve_event(event);

            match event {
                Event::Start(element) => {
                    depth = depth.checked_add(1).ok_or_else(|| {
                        OoxmlError::InvalidFormat("document-variable XML nesting overflow".into())
                    })?;
                    if depth > MAX_DOCUMENT_VARIABLE_DEPTH {
                        return Err(OoxmlError::InvalidFormat(format!(
                            "document-variable XML exceeds depth {MAX_DOCUMENT_VARIABLE_DEPTH}"
                        )));
                    }
                    if depth == 1 {
                        validate_settings_root(&namespace, &element, saw_root)?;
                        saw_root = true;
                    } else if depth == 2
                        && is_wordprocessing_namespace(&namespace)
                        && element.local_name().as_ref() == b"docVars"
                    {
                        begin_doc_vars(&mut saw_doc_vars, &mut doc_vars_depth, depth)?;
                    } else if depth == 3
                        && doc_vars_depth == Some(2)
                        && is_wordprocessing_namespace(&namespace)
                        && element.local_name().as_ref() == b"docVar"
                    {
                        push_parsed_variable(&mut variables, &element, decoder, &resolver)?;
                        open_doc_var_depth = Some(depth);
                    } else if is_wordprocessing_namespace(&namespace)
                        && matches!(element.local_name().as_ref(), b"docVars" | b"docVar")
                    {
                        return Err(OoxmlError::InvalidFormat(
                            "misplaced or nested document-variable element".into(),
                        ));
                    } else if doc_vars_depth.is_some() && is_wordprocessing_namespace(&namespace) {
                        return Err(OoxmlError::InvalidFormat(
                            "unexpected WordprocessingML child in docVars".into(),
                        ));
                    }
                },
                Event::Empty(element) => {
                    let child_depth = depth.checked_add(1).ok_or_else(|| {
                        OoxmlError::InvalidFormat("document-variable XML nesting overflow".into())
                    })?;
                    if child_depth > MAX_DOCUMENT_VARIABLE_DEPTH {
                        return Err(OoxmlError::InvalidFormat(format!(
                            "document-variable XML exceeds depth {MAX_DOCUMENT_VARIABLE_DEPTH}"
                        )));
                    }
                    if child_depth == 1 {
                        validate_settings_root(&namespace, &element, saw_root)?;
                        saw_root = true;
                    } else if child_depth == 2
                        && is_wordprocessing_namespace(&namespace)
                        && element.local_name().as_ref() == b"docVars"
                    {
                        begin_doc_vars(&mut saw_doc_vars, &mut doc_vars_depth, child_depth)?;
                        doc_vars_depth = None;
                    } else if child_depth == 3
                        && doc_vars_depth == Some(2)
                        && is_wordprocessing_namespace(&namespace)
                        && element.local_name().as_ref() == b"docVar"
                    {
                        push_parsed_variable(&mut variables, &element, decoder, &resolver)?;
                    } else if is_wordprocessing_namespace(&namespace)
                        && matches!(element.local_name().as_ref(), b"docVars" | b"docVar")
                    {
                        return Err(OoxmlError::InvalidFormat(
                            "misplaced or nested document-variable element".into(),
                        ));
                    } else if doc_vars_depth.is_some() && is_wordprocessing_namespace(&namespace) {
                        return Err(OoxmlError::InvalidFormat(
                            "unexpected WordprocessingML child in docVars".into(),
                        ));
                    }
                },
                Event::End(_) => {
                    if open_doc_var_depth == Some(depth) {
                        open_doc_var_depth = None;
                    }
                    if doc_vars_depth == Some(depth) {
                        doc_vars_depth = None;
                    }
                    depth = depth.checked_sub(1).ok_or_else(|| {
                        OoxmlError::InvalidFormat("invalid document-variable XML nesting".into())
                    })?;
                },
                Event::Text(text)
                    if (open_doc_var_depth.is_some() || doc_vars_depth == Some(depth))
                        && text.as_ref().iter().any(|byte| !byte.is_ascii_whitespace()) =>
                {
                    return Err(OoxmlError::InvalidFormat(
                        "document-variable elements cannot contain text".into(),
                    ));
                },
                Event::CData(text)
                    if (open_doc_var_depth.is_some() || doc_vars_depth == Some(depth))
                        && text.as_ref().iter().any(|byte| !byte.is_ascii_whitespace()) =>
                {
                    return Err(OoxmlError::InvalidFormat(
                        "document-variable elements cannot contain text".into(),
                    ));
                },
                Event::Eof if depth != 0 => {
                    return Err(OoxmlError::InvalidFormat(
                        "unterminated document-variable settings XML".into(),
                    ));
                },
                Event::Eof => break,
                _ => {},
            }
        }
        if !saw_root {
            return Err(OoxmlError::InvalidFormat(
                "settings part has no settings root".into(),
            ));
        }
        Ok(variables)
    }
}

fn validate_settings_root(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    saw_root: bool,
) -> Result<()> {
    if saw_root
        || !is_wordprocessing_namespace(namespace)
        || element.local_name().as_ref() != b"settings"
    {
        return Err(OoxmlError::InvalidFormat(
            "document variables require one Word settings root".into(),
        ));
    }
    Ok(())
}

fn begin_doc_vars(
    saw_doc_vars: &mut bool,
    doc_vars_depth: &mut Option<usize>,
    depth: usize,
) -> Result<()> {
    if std::mem::replace(saw_doc_vars, true) {
        return Err(OoxmlError::InvalidFormat(
            "duplicate docVars container".into(),
        ));
    }
    *doc_vars_depth = Some(depth);
    Ok(())
}

fn push_parsed_variable(
    variables: &mut DocumentVariables,
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<()> {
    let name = word_attribute_value(element, b"name", decoder, resolver)?.ok_or_else(|| {
        OoxmlError::InvalidFormat("document variable name attribute is required".into())
    })?;
    let value = word_attribute_value(element, b"val", decoder, resolver)?.ok_or_else(|| {
        OoxmlError::InvalidFormat("document variable val attribute is required".into())
    })?;
    validate_document_variable(&name, &value)?;
    if variables.contains(&name) {
        return Err(OoxmlError::InvalidFormat(format!(
            "duplicate document variable name {name:?}"
        )));
    }
    if variables.variables.len() >= MAX_DOCUMENT_VARIABLES {
        return Err(OoxmlError::InvalidFormat(format!(
            "document variables exceed the {MAX_DOCUMENT_VARIABLES} entry limit"
        )));
    }
    variables.variables.push((name, value));
    Ok(())
}

fn validate_document_variable(name: &str, value: &str) -> Result<()> {
    let name_chars = name.chars().count();
    if !(1..=MAX_DOCUMENT_VARIABLE_NAME_CHARS).contains(&name_chars) {
        return Err(OoxmlError::InvalidFormat(format!(
            "document variable name must contain 1 to {MAX_DOCUMENT_VARIABLE_NAME_CHARS} characters"
        )));
    }
    if value.chars().count() > MAX_DOCUMENT_VARIABLE_VALUE_CHARS {
        return Err(OoxmlError::InvalidFormat(format!(
            "document variable value exceeds {MAX_DOCUMENT_VARIABLE_VALUE_CHARS} characters"
        )));
    }
    Ok(())
}

fn escape_attribute(output: &mut String, value: &str) {
    for character in value.chars() {
        match character {
            '&' => output.push_str("&amp;"),
            '<' => output.push_str("&lt;"),
            '>' => output.push_str("&gt;"),
            '"' => output.push_str("&quot;"),
            '\'' => output.push_str("&apos;"),
            _ => output.push(character),
        }
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
