//! Spreadsheet cell annotations (comments/notes).

use litchi_core::{Error, Result, xml::escape_xml};
use quick_xml::{
    Decoder, XmlVersion,
    events::{BytesCData, BytesRef, BytesStart, BytesText},
};
use std::collections::{BTreeMap, BTreeSet};

/// One ordered node in an annotation's rich text content.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum AnnotationNode {
    /// Character data. It is always XML-escaped when serialized.
    Text(String),
    /// A nested ODF text or extension element.
    Element(AnnotationElement),
}

/// A lossless, ordered XML element within a cell annotation.
///
/// This representation is intentionally generic: ODF permits the full text
/// paragraph and list content model in annotations, including spans, links,
/// fields, tabs, line breaks, and implementation-defined extension elements.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct AnnotationElement {
    name: String,
    attributes: BTreeMap<String, String>,
    children: Vec<AnnotationNode>,
}

impl AnnotationElement {
    /// Create an annotation content element with a validated XML qualified name.
    pub fn new(name: impl Into<String>) -> Result<Self> {
        let name = name.into();
        validate_qname(&name)?;
        Ok(Self {
            name,
            attributes: BTreeMap::new(),
            children: Vec::new(),
        })
    }

    /// Return the element's qualified XML name, such as `text:span`.
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Return all attributes in deterministic qualified-name order.
    pub fn attributes(&self) -> &BTreeMap<String, String> {
        &self.attributes
    }

    /// Return an attribute value by qualified name.
    pub fn attribute(&self, name: &str) -> Option<&str> {
        self.attributes.get(name).map(String::as_str)
    }

    /// Set an escaped attribute after validating its XML qualified name.
    pub fn set_attribute(
        &mut self,
        name: impl Into<String>,
        value: impl Into<String>,
    ) -> Result<&mut Self> {
        let name = name.into();
        validate_qname(&name)?;
        if name == "xmlns" || name.starts_with("xmlns:") {
            return Err(Error::InvalidFormat(
                "namespace declarations are managed by the annotation writer".to_string(),
            ));
        }
        self.attributes.insert(name, value.into());
        Ok(self)
    }

    /// Remove an attribute.
    pub fn remove_attribute(&mut self, name: &str) -> Option<String> {
        self.attributes.remove(name)
    }

    /// Return the ordered mixed-content nodes.
    pub fn children(&self) -> &[AnnotationNode] {
        &self.children
    }

    /// Append escaped character data.
    pub fn push_text(&mut self, text: impl Into<String>) -> &mut Self {
        self.children.push(AnnotationNode::Text(text.into()));
        self
    }

    /// Append a nested rich-content element.
    pub fn push_element(&mut self, element: AnnotationElement) -> &mut Self {
        self.children.push(AnnotationNode::Element(element));
        self
    }

    /// Extract the rendered plain text represented by this element.
    pub fn plain_text(&self) -> String {
        let local_name = local_name(&self.name);
        if local_name == "s" {
            let count = self
                .attribute("text:c")
                .and_then(|value| value.parse::<usize>().ok())
                .unwrap_or(1)
                .min(1_000_000);
            return " ".repeat(count);
        }
        if local_name == "tab" {
            return "\t".to_string();
        }
        if local_name == "line-break" {
            return "\n".to_string();
        }

        let mut text = String::new();
        for child in &self.children {
            match child {
                AnnotationNode::Text(value) => text.push_str(value),
                AnnotationNode::Element(element) => text.push_str(&element.plain_text()),
            }
        }
        text
    }

    fn write_xml(&self, output: &mut String) {
        output.push('<');
        output.push_str(&self.name);
        write_attributes(output, &self.attributes);
        if self.children.is_empty() {
            output.push_str("/>");
            return;
        }

        output.push('>');
        for child in &self.children {
            match child {
                AnnotationNode::Text(text) => output.push_str(&escape_xml(text)),
                AnnotationNode::Element(element) => element.write_xml(output),
            }
        }
        output.push_str("</");
        output.push_str(&self.name);
        output.push('>');
    }

    fn collect_prefixes(&self, prefixes: &mut BTreeSet<String>) {
        collect_prefix(&self.name, prefixes);
        for name in self.attributes.keys() {
            collect_prefix(name, prefixes);
        }
        for child in &self.children {
            if let AnnotationNode::Element(element) = child {
                element.collect_prefixes(prefixes);
            }
        }
    }

    fn collect_namespace_declarations(&self, prefixes: &mut BTreeSet<String>) {
        for name in self.attributes.keys() {
            if let Some(prefix) = name.strip_prefix("xmlns:") {
                prefixes.insert(prefix.to_string());
            }
        }
        for child in &self.children {
            if let AnnotationNode::Element(element) = child {
                element.collect_namespace_declarations(prefixes);
            }
        }
    }
}

/// An ODF `office:annotation` attached to a spreadsheet cell.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct CellAnnotation {
    attributes: BTreeMap<String, String>,
    children: Vec<AnnotationElement>,
    namespaces: BTreeMap<String, String>,
}

impl CellAnnotation {
    /// Create a plain-text cell annotation containing one `text:p` element.
    pub fn new(text: impl Into<String>) -> Self {
        let mut paragraph =
            AnnotationElement::new("text:p").expect("the built-in text:p qualified name is valid");
        paragraph.push_text(text);
        Self {
            attributes: BTreeMap::new(),
            children: vec![paragraph],
            namespaces: BTreeMap::new(),
        }
    }

    /// Return all annotation attributes, including drawing layout attributes.
    pub fn attributes(&self) -> &BTreeMap<String, String> {
        &self.attributes
    }

    /// Return namespace bindings retained for extension content.
    pub fn namespaces(&self) -> &BTreeMap<String, String> {
        &self.namespaces
    }

    /// Add a namespace binding for custom rich-content elements or attributes.
    pub fn set_namespace(
        &mut self,
        prefix: impl Into<String>,
        uri: impl Into<String>,
    ) -> Result<&mut Self> {
        let prefix = prefix.into();
        let uri = uri.into();
        if prefix.contains(':') || !valid_name_part(&prefix) || prefix == "xmlns" {
            return Err(Error::InvalidFormat(format!(
                "invalid XML namespace prefix '{prefix}'"
            )));
        }
        if prefix == "xml" {
            if uri != "http://www.w3.org/XML/1998/namespace" {
                return Err(Error::InvalidFormat(
                    "the xml prefix has a fixed namespace URI".to_string(),
                ));
            }
            return Ok(self);
        }
        if let Some(standard) = standard_namespace_uri(&prefix)
            && uri != standard
        {
            return Err(Error::InvalidFormat(format!(
                "the standard '{prefix}' prefix must use namespace URI '{standard}'"
            )));
        }
        if uri.is_empty() {
            return Err(Error::InvalidFormat(
                "annotation namespace URI cannot be empty".to_string(),
            ));
        }
        self.namespaces.insert(prefix, uri);
        Ok(self)
    }

    /// Return an annotation attribute by qualified name.
    pub fn attribute(&self, name: &str) -> Option<&str> {
        self.attributes
            .get(name)
            .or_else(|| {
                self.equivalent_attribute_key(name)
                    .and_then(|key| self.attributes.get(key))
            })
            .map(String::as_str)
    }

    /// Set any ODF annotation, drawing, SVG, text, or extension attribute.
    pub fn set_attribute(
        &mut self,
        name: impl Into<String>,
        value: impl Into<String>,
    ) -> Result<&mut Self> {
        let name = name.into();
        validate_qname(&name)?;
        if name == "xmlns" || name.starts_with("xmlns:") {
            return Err(Error::InvalidFormat(
                "namespace declarations are managed by the annotation writer".to_string(),
            ));
        }
        self.attributes.insert(name, value.into());
        Ok(self)
    }

    /// Remove an annotation attribute.
    pub fn remove_attribute(&mut self, name: &str) -> Option<String> {
        if let Some(value) = self.attributes.remove(name) {
            return Some(value);
        }
        let key = self.equivalent_attribute_key(name)?.to_string();
        self.attributes.remove(&key)
    }

    /// Return the optional `office:name` identifier.
    pub fn name(&self) -> Option<&str> {
        self.attribute("office:name")
    }

    /// Set or clear the `office:name` identifier.
    pub fn set_name(&mut self, name: Option<&str>) {
        self.remove_attribute("office:name");
        set_optional_attribute(&mut self.attributes, "office:name", name);
    }

    /// Return the optional requested display state.
    pub fn display(&self) -> Option<bool> {
        match self.attribute("office:display") {
            Some("true" | "1") => Some(true),
            Some("false" | "0") => Some(false),
            _ => None,
        }
    }

    /// Set or clear the requested display state.
    pub fn set_display(&mut self, display: Option<bool>) {
        self.remove_attribute("office:display");
        match display {
            Some(value) => {
                self.attributes.insert(
                    "office:display".to_string(),
                    if value { "true" } else { "false" }.to_string(),
                );
            },
            None => {
                self.attributes.remove("office:display");
            },
        }
    }

    /// Return the annotation author from `dc:creator`, if present.
    pub fn creator(&self) -> Option<String> {
        self.metadata_text("creator")
            .or_else(|| self.attribute("office:author").map(str::to_string))
    }

    /// Set or clear `dc:creator` while retaining schema child ordering.
    pub fn set_creator(&mut self, creator: Option<&str>) {
        self.set_metadata("dc:creator", creator, 0);
    }

    /// Return the machine-readable `dc:date`, if present.
    pub fn date(&self) -> Option<String> {
        self.metadata_text("date")
            .or_else(|| self.attribute("office:create-date").map(str::to_string))
    }

    /// Set or clear the machine-readable `dc:date`.
    pub fn set_date(&mut self, date: Option<&str>) {
        self.set_metadata("dc:date", date, 1);
    }

    /// Return the human-readable `meta:date-string`, if present.
    pub fn date_string(&self) -> Option<String> {
        self.metadata_text("date-string").or_else(|| {
            self.attribute("office:create-date-string")
                .map(str::to_string)
        })
    }

    /// Set or clear `meta:date-string`.
    pub fn set_date_string(&mut self, date: Option<&str>) {
        self.set_metadata("meta:date-string", date, 2);
    }

    /// Return all ordered child elements, including metadata and rich content.
    pub fn children(&self) -> &[AnnotationElement] {
        &self.children
    }

    /// Append a plain-text paragraph.
    pub fn push_paragraph(&mut self, text: impl Into<String>) -> &mut Self {
        let mut paragraph =
            AnnotationElement::new("text:p").expect("the built-in text:p qualified name is valid");
        paragraph.push_text(text);
        self.children.push(paragraph);
        self
    }

    /// Append a rich paragraph, list, or extension element.
    pub fn push_element(&mut self, element: AnnotationElement) -> &mut Self {
        self.children.push(element);
        self
    }

    /// Extract annotation body text, separating top-level paragraphs/list blocks.
    pub fn text(&self) -> String {
        self.children
            .iter()
            .filter(|child| !is_metadata_element(child))
            .map(AnnotationElement::plain_text)
            .collect::<Vec<_>>()
            .join("\n")
    }

    /// Validate namespace bindings before serializing this annotation.
    pub fn validate(&self) -> Result<()> {
        let mut used_prefixes = BTreeSet::new();
        for name in self.attributes.keys() {
            collect_prefix(name, &mut used_prefixes);
        }
        for child in &self.children {
            child.collect_prefixes(&mut used_prefixes);
        }
        let mut locally_declared = BTreeSet::new();
        for child in &self.children {
            child.collect_namespace_declarations(&mut locally_declared);
        }
        for prefix in used_prefixes {
            if standard_namespace_uri(&prefix).is_none()
                && !self.namespaces.contains_key(&prefix)
                && !locally_declared.contains(&prefix)
            {
                return Err(Error::InvalidFormat(format!(
                    "annotation uses undeclared XML namespace prefix '{prefix}'"
                )));
            }
        }
        Ok(())
    }

    fn metadata_text(&self, wanted_local_name: &str) -> Option<String> {
        self.children
            .iter()
            .find(|child| local_name(child.name()) == wanted_local_name)
            .map(AnnotationElement::plain_text)
    }

    fn equivalent_attribute_key(&self, requested: &str) -> Option<&str> {
        let (requested_prefix, requested_local) = requested.split_once(':')?;
        let requested_namespace = self
            .namespaces
            .get(requested_prefix)
            .map(String::as_str)
            .or_else(|| standard_namespace_uri(requested_prefix))?;

        self.attributes.keys().find_map(|candidate| {
            let (prefix, local) = candidate.split_once(':')?;
            if local != requested_local {
                return None;
            }
            let namespace = self
                .namespaces
                .get(prefix)
                .map(String::as_str)
                .or_else(|| standard_namespace_uri(prefix))?;
            (namespace == requested_namespace).then_some(candidate.as_str())
        })
    }

    fn set_metadata(&mut self, name: &str, value: Option<&str>, order: usize) {
        let local = local_name(name);
        self.children
            .retain(|child| local_name(child.name()) != local);
        let Some(value) = value else { return };

        let mut element = AnnotationElement::new(name)
            .expect("built-in annotation metadata qualified names are valid");
        element.push_text(value);
        let insertion = self
            .children
            .iter()
            .position(|child| metadata_order(child).is_none_or(|other| other > order))
            .unwrap_or(self.children.len());
        self.children.insert(insertion, element);
    }

    pub(crate) fn write_xml(&self, output: &mut String) {
        output.push_str("<office:annotation");
        write_attributes(output, &self.attributes);

        let mut used_prefixes = BTreeSet::new();
        for name in self.attributes.keys() {
            collect_prefix(name, &mut used_prefixes);
        }
        for child in &self.children {
            child.collect_prefixes(&mut used_prefixes);
        }
        for prefix in used_prefixes {
            if let Some(uri) = self.namespaces.get(&prefix) {
                output.push_str(" xmlns:");
                output.push_str(&prefix);
                output.push_str("=\"");
                output.push_str(&escape_xml(uri));
                output.push('"');
            }
        }

        output.push('>');
        for child in &self.children {
            child.write_xml(output);
        }
        output.push_str("</office:annotation>");
    }
}

impl Default for CellAnnotation {
    fn default() -> Self {
        Self::new("")
    }
}

pub(crate) struct AnnotationBuilder {
    annotation: CellAnnotation,
    stack: Vec<AnnotationElement>,
}

impl AnnotationBuilder {
    pub(crate) fn new(
        start: &BytesStart<'_>,
        decoder: Decoder,
        namespaces: BTreeMap<String, String>,
    ) -> Result<Self> {
        Ok(Self {
            annotation: CellAnnotation {
                attributes: parse_attributes(start, decoder, false)?,
                children: Vec::new(),
                namespaces,
            },
            stack: Vec::new(),
        })
    }

    pub(crate) fn start(&mut self, start: &BytesStart<'_>, decoder: Decoder) -> Result<()> {
        self.stack.push(parse_element(start, decoder)?);
        Ok(())
    }

    pub(crate) fn empty(&mut self, start: &BytesStart<'_>, decoder: Decoder) -> Result<()> {
        self.add_element(parse_element(start, decoder)?)
    }

    pub(crate) fn text(&mut self, text: &BytesText<'_>) -> Result<()> {
        let decoded = text
            .xml_content(XmlVersion::Explicit1_0)
            .map_err(|error| Error::InvalidFormat(format!("invalid annotation text: {error}")))?;
        if let Some(parent) = self.stack.last_mut() {
            parent
                .children
                .push(AnnotationNode::Text(decoded.into_owned()));
        }
        Ok(())
    }

    pub(crate) fn reference(&mut self, reference: &BytesRef<'_>) -> Result<()> {
        let value = decode_reference(reference)?;
        if let Some(parent) = self.stack.last_mut() {
            parent.children.push(AnnotationNode::Text(value));
        }
        Ok(())
    }

    pub(crate) fn cdata(&mut self, text: &BytesCData<'_>) -> Result<()> {
        let decoded = text
            .xml_content(XmlVersion::Explicit1_0)
            .map_err(|error| Error::InvalidFormat(format!("invalid annotation CDATA: {error}")))?;
        if let Some(parent) = self.stack.last_mut() {
            parent
                .children
                .push(AnnotationNode::Text(decoded.into_owned()));
        }
        Ok(())
    }

    pub(crate) fn end_element(&mut self) -> Result<()> {
        let element = self.stack.pop().ok_or_else(|| {
            Error::InvalidFormat("unbalanced element in cell annotation".to_string())
        })?;
        self.add_element(element)
    }

    pub(crate) fn finish(self) -> Result<CellAnnotation> {
        if !self.stack.is_empty() {
            return Err(Error::InvalidFormat(
                "unclosed element in cell annotation".to_string(),
            ));
        }
        Ok(self.annotation)
    }

    fn add_element(&mut self, element: AnnotationElement) -> Result<()> {
        if let Some(parent) = self.stack.last_mut() {
            parent.children.push(AnnotationNode::Element(element));
        } else {
            self.annotation.children.push(element);
        }
        Ok(())
    }
}

pub(crate) fn decode_reference(reference: &BytesRef<'_>) -> Result<String> {
    if let Some(character) = reference.resolve_char_ref().map_err(|error| {
        Error::InvalidFormat(format!("invalid XML character reference: {error}"))
    })? {
        return Ok(character.to_string());
    }
    let name = reference
        .decode()
        .map_err(|error| Error::InvalidFormat(format!("invalid XML entity reference: {error}")))?;
    match name.as_ref() {
        "amp" => Ok("&".to_string()),
        "lt" => Ok("<".to_string()),
        "gt" => Ok(">".to_string()),
        "quot" => Ok("\"".to_string()),
        "apos" => Ok("'".to_string()),
        _ => Err(Error::InvalidFormat(format!(
            "unsupported XML entity reference '&{name};'"
        ))),
    }
}

fn parse_element(start: &BytesStart<'_>, decoder: Decoder) -> Result<AnnotationElement> {
    let name = std::str::from_utf8(start.name().as_ref())
        .map_err(|_| Error::InvalidFormat("invalid UTF-8 in annotation element name".to_string()))?
        .to_string();
    validate_qname(&name)?;
    Ok(AnnotationElement {
        name,
        attributes: parse_attributes(start, decoder, true)?,
        children: Vec::new(),
    })
}

fn parse_attributes(
    start: &BytesStart<'_>,
    decoder: Decoder,
    keep_namespaces: bool,
) -> Result<BTreeMap<String, String>> {
    let mut attributes = BTreeMap::new();
    for attribute in start.attributes() {
        let attribute = attribute.map_err(|error| {
            Error::InvalidFormat(format!("invalid cell annotation attribute: {error}"))
        })?;
        let name = std::str::from_utf8(attribute.key.as_ref())
            .map_err(|_| {
                Error::InvalidFormat("invalid UTF-8 in annotation attribute name".to_string())
            })?
            .to_string();
        if !keep_namespaces && (name == "xmlns" || name.starts_with("xmlns:")) {
            continue;
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid cell annotation attribute value: {error}"))
            })?
            .into_owned();
        attributes.insert(name, value);
    }
    Ok(attributes)
}

fn write_attributes(output: &mut String, attributes: &BTreeMap<String, String>) {
    for (name, value) in attributes {
        output.push(' ');
        output.push_str(name);
        output.push_str("=\"");
        output.push_str(&escape_xml(value));
        output.push('"');
    }
}

fn set_optional_attribute(
    attributes: &mut BTreeMap<String, String>,
    name: &str,
    value: Option<&str>,
) {
    if let Some(value) = value {
        attributes.insert(name.to_string(), value.to_string());
    } else {
        attributes.remove(name);
    }
}

fn metadata_order(element: &AnnotationElement) -> Option<usize> {
    match local_name(element.name()) {
        "creator" => Some(0),
        "date" => Some(1),
        "date-string" => Some(2),
        _ => None,
    }
}

fn is_metadata_element(element: &AnnotationElement) -> bool {
    metadata_order(element).is_some()
}

fn local_name(name: &str) -> &str {
    name.rsplit_once(':').map_or(name, |(_, local)| local)
}

fn collect_prefix(name: &str, prefixes: &mut BTreeSet<String>) {
    if let Some((prefix, _)) = name.split_once(':')
        && prefix != "xml"
        && prefix != "xmlns"
    {
        prefixes.insert(prefix.to_string());
    }
}

fn standard_namespace_uri(prefix: &str) -> Option<&'static str> {
    match prefix {
        "office" => Some("urn:oasis:names:tc:opendocument:xmlns:office:1.0"),
        "text" => Some("urn:oasis:names:tc:opendocument:xmlns:text:1.0"),
        "table" => Some("urn:oasis:names:tc:opendocument:xmlns:table:1.0"),
        "draw" => Some("urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"),
        "svg" => Some("urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"),
        "dc" => Some("http://purl.org/dc/elements/1.1/"),
        "meta" => Some("urn:oasis:names:tc:opendocument:xmlns:meta:1.0"),
        "xlink" => Some("http://www.w3.org/1999/xlink"),
        "fo" => Some("urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"),
        "style" => Some("urn:oasis:names:tc:opendocument:xmlns:style:1.0"),
        "loext" => Some("urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0"),
        _ => None,
    }
}

fn validate_qname(name: &str) -> Result<()> {
    let mut parts = name.split(':');
    let first = parts.next().unwrap_or_default();
    let second = parts.next();
    if first.is_empty() || parts.next().is_some() || !valid_name_part(first) {
        return Err(Error::InvalidFormat(format!(
            "invalid XML qualified name '{name}'"
        )));
    }
    if second.is_some_and(|part| part.is_empty() || !valid_name_part(part)) {
        return Err(Error::InvalidFormat(format!(
            "invalid XML qualified name '{name}'"
        )));
    }
    Ok(())
}

fn valid_name_part(part: &str) -> bool {
    let mut chars = part.chars();
    chars
        .next()
        .is_some_and(|first| first == '_' || first.is_alphabetic())
        && chars
            .all(|ch| ch == '_' || ch == '-' || ch == '.' || ch.is_alphanumeric() || ch == '\u{b7}')
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn builds_rich_annotation_and_escapes_values() {
        let mut annotation = CellAnnotation::new("first & <line>");
        annotation.set_creator(Some("A&B"));
        annotation.set_display(Some(true));
        annotation.set_attribute("svg:width", "3\" & 4cm").unwrap();

        let mut span = AnnotationElement::new("text:span").unwrap();
        span.set_attribute("text:style-name", "Strong").unwrap();
        span.push_text("bold");
        let mut paragraph = AnnotationElement::new("text:p").unwrap();
        paragraph.push_element(span);
        paragraph.push_element(AnnotationElement::new("text:line-break").unwrap());
        paragraph.push_text("after");
        annotation.push_element(paragraph);

        assert_eq!(annotation.creator().as_deref(), Some("A&B"));
        assert_eq!(annotation.display(), Some(true));
        assert_eq!(annotation.text(), "first & <line>\nbold\nafter");

        let mut xml = String::new();
        annotation.write_xml(&mut xml);
        assert!(xml.contains("office:display=\"true\""));
        assert!(
            xml.contains("svg:width=\"3&amp;quot; &amp; 4cm\"")
                || xml.contains("svg:width=\"3&quot; &amp; 4cm\"")
        );
        assert!(xml.contains("first &amp; &lt;line&gt;"));
    }

    #[test]
    fn rejects_names_that_could_inject_xml() {
        assert!(AnnotationElement::new("text:p><evil").is_err());
        let mut annotation = CellAnnotation::default();
        assert!(annotation.set_attribute("x\" y", "value").is_err());
        assert!(annotation.set_attribute("xmlns:evil", "urn:evil").is_err());
    }

    #[test]
    fn validates_custom_extension_namespaces() {
        let mut annotation = CellAnnotation::new("root");
        annotation.push_element(AnnotationElement::new("vendor:thread").unwrap());
        assert!(annotation.validate().is_err());

        annotation
            .set_namespace("vendor", "urn:example:annotation")
            .unwrap();
        annotation.validate().unwrap();
        let mut xml = String::new();
        annotation.write_xml(&mut xml);
        assert!(xml.contains("xmlns:vendor=\"urn:example:annotation\""));
        assert!(xml.contains("<vendor:thread/>"));
    }

    #[test]
    fn reads_legacy_annotation_metadata_attributes() {
        let mut annotation = CellAnnotation::default();
        annotation
            .set_attribute("office:author", "Legacy Author")
            .unwrap();
        annotation
            .set_attribute("office:create-date", "2002-01-01T00:00:00")
            .unwrap();
        annotation
            .set_attribute("office:create-date-string", "January 1, 2002")
            .unwrap();

        assert_eq!(annotation.creator().as_deref(), Some("Legacy Author"));
        assert_eq!(annotation.date().as_deref(), Some("2002-01-01T00:00:00"));
        assert_eq!(annotation.date_string().as_deref(), Some("January 1, 2002"));
    }
}
