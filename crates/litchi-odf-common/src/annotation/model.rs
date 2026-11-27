//! Semantic `ODF` annotation values and contextual rich-content operations.

use super::package::MAX_ANNOTATION_BODY_ELEMENTS;
use litchi_core::{Error, Result};
use std::collections::{BTreeMap, BTreeSet};

/// One ordered node in an annotation's rich text content.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum Node {
    /// Character data. It is always XML-escaped when serialized.
    Text(String),
    /// A nested `ODF` text or extension element.
    Element(Element),
}

/// A lossless, ordered `XML` element within an `ODF` annotation.
///
/// This representation is intentionally generic: `ODF` permits the full text
/// paragraph and list content model in annotations, including spans, links,
/// fields, tabs, line breaks, and implementation-defined extension elements.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Element {
    pub(crate) name: String,
    pub(crate) attributes: BTreeMap<String, String>,
    pub(crate) children: Vec<Node>,
}

impl Element {
    /// Create an annotation content element with a validated XML qualified name.
    ///
    /// # Errors
    ///
    /// Returns an error when the name is not a valid XML qualified name.
    pub fn new(qualified_name: impl Into<String>) -> Result<Self> {
        let name = qualified_name.into();
        validate_qname(&name)?;
        Ok(Self {
            name,
            attributes: BTreeMap::new(),
            children: Vec::new(),
        })
    }

    /// Return the element's qualified XML name, such as `text:span`.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Return all attributes in deterministic qualified-name order.
    #[must_use]
    pub fn attributes(&self) -> &BTreeMap<String, String> {
        &self.attributes
    }

    /// Return an attribute value by qualified name.
    #[must_use]
    pub fn attribute(&self, name: &str) -> Option<&str> {
        self.attributes.get(name).map(String::as_str)
    }

    /// Set an escaped attribute after validating its XML qualified name.
    ///
    /// # Errors
    ///
    /// Returns an error when the attribute name is not valid or attempts to
    /// manage a namespace declaration.
    pub fn set_attribute(
        &mut self,
        qualified_name: impl Into<String>,
        value: impl Into<String>,
    ) -> Result<&mut Self> {
        let name = qualified_name.into();
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
    #[must_use]
    pub fn children(&self) -> &[Node] {
        &self.children
    }

    /// Append escaped character data.
    pub fn push_text(&mut self, text: impl Into<String>) -> &mut Self {
        self.children.push(Node::Text(text.into()));
        self
    }

    /// Append a nested rich-content element.
    pub fn push_element(&mut self, element: Element) -> &mut Self {
        self.children.push(Node::Element(element));
        self
    }

    /// Extract the rendered plain text represented by this element.
    #[must_use]
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
                Node::Text(value) => text.push_str(value),
                Node::Element(element) => text.push_str(&element.plain_text()),
            }
        }
        text
    }

    pub(crate) fn collect_prefixes(&self, prefixes: &mut BTreeSet<String>) {
        collect_prefix(&self.name, prefixes);
        for name in self.attributes.keys() {
            collect_prefix(name, prefixes);
        }
        for child in &self.children {
            if let Node::Element(element) = child {
                element.collect_prefixes(prefixes);
            }
        }
    }

    pub(crate) fn collect_namespace_declarations(&self, prefixes: &mut BTreeSet<String>) {
        for name in self.attributes.keys() {
            if let Some(prefix) = name.strip_prefix("xmlns:") {
                prefixes.insert(prefix.to_string());
            }
        }
        for child in &self.children {
            if let Node::Element(element) = child {
                element.collect_namespace_declarations(prefixes);
            }
        }
    }
}

/// An ODF office annotation with typed metadata and retained rich content.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Annotation {
    pub(crate) attributes: BTreeMap<String, String>,
    pub(crate) children: Vec<Element>,
    pub(crate) namespaces: BTreeMap<String, String>,
}

impl Annotation {
    /// Create a plain-text annotation containing one `text:p` element.
    pub fn new(text: impl Into<String>) -> Self {
        let mut paragraph = built_in_element("text:p");
        paragraph.push_text(text);
        Self {
            attributes: BTreeMap::new(),
            children: vec![paragraph],
            namespaces: BTreeMap::new(),
        }
    }

    /// Return all annotation attributes, including drawing layout attributes.
    #[must_use]
    pub fn attributes(&self) -> &BTreeMap<String, String> {
        &self.attributes
    }

    /// Return namespace bindings retained for extension content.
    #[must_use]
    pub fn namespaces(&self) -> &BTreeMap<String, String> {
        &self.namespaces
    }

    /// Add a namespace binding for custom rich-content elements or attributes.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid prefix, an empty URI, or a standard
    /// prefix bound to the wrong URI.
    pub fn set_namespace(
        &mut self,
        prefix_input: impl Into<String>,
        uri_input: impl Into<String>,
    ) -> Result<&mut Self> {
        let prefix = prefix_input.into();
        let uri = uri_input.into();
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
    #[must_use]
    pub fn attribute(&self, name: &str) -> Option<&str> {
        self.attributes
            .get(name)
            .or_else(|| {
                self.equivalent_attribute_key(name)
                    .and_then(|key| self.attributes.get(key))
            })
            .map(String::as_str)
    }

    /// Set any `ODF` annotation, drawing, `SVG`, text, or extension attribute.
    ///
    /// # Errors
    ///
    /// Returns an error when the attribute name is not valid or attempts to
    /// manage a namespace declaration.
    pub fn set_attribute(
        &mut self,
        qualified_name: impl Into<String>,
        value: impl Into<String>,
    ) -> Result<&mut Self> {
        let name = qualified_name.into();
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

    /// Return the optional office:name identifier.
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.attribute("office:name")
    }

    /// Set or clear the office:name identifier.
    pub fn set_name(&mut self, name: Option<&str>) {
        self.remove_attribute("office:name");
        set_optional_attribute(&mut self.attributes, "office:name", name);
    }

    /// Return the optional requested display state.
    #[must_use]
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

    /// Return the annotation author from dc:creator, if present.
    #[must_use]
    pub fn creator(&self) -> Option<String> {
        self.metadata_text("creator")
            .or_else(|| self.attribute("office:author").map(str::to_string))
    }

    /// Set or clear dc:creator while retaining schema child ordering.
    pub fn set_creator(&mut self, creator: Option<&str>) {
        self.set_metadata("dc:creator", creator, 0);
    }

    /// Return the machine-readable dc:date, if present.
    #[must_use]
    pub fn date(&self) -> Option<String> {
        self.metadata_text("date")
            .or_else(|| self.attribute("office:create-date").map(str::to_string))
    }

    /// Set or clear the machine-readable dc:date.
    pub fn set_date(&mut self, date: Option<&str>) {
        self.set_metadata("dc:date", date, 1);
    }

    /// Return the human-readable meta:date-string, if present.
    #[must_use]
    pub fn date_string(&self) -> Option<String> {
        self.metadata_text("date-string").or_else(|| {
            self.attribute("office:create-date-string")
                .map(str::to_string)
        })
    }

    /// Set or clear meta:date-string.
    pub fn set_date_string(&mut self, date: Option<&str>) {
        self.set_metadata("meta:date-string", date, 3);
    }

    /// Return `ODF` 1.3 `meta:creator-initials`, including `LibreOffice`'s legacy
    /// `text:sender-initials` and `loext:sender-initials` spellings.
    #[must_use]
    pub fn initials(&self) -> Option<String> {
        self.children
            .iter()
            .find(|child| {
                matches!(
                    local_name(child.name()),
                    "creator-initials" | "sender-initials"
                )
            })
            .map(Element::plain_text)
    }

    /// Set or clear canonical `ODF` 1.3 `meta:creator-initials` metadata.
    pub fn set_initials(&mut self, value: Option<&str>) {
        self.children.retain(|child| {
            !matches!(
                local_name(child.name()),
                "creator-initials" | "sender-initials"
            )
        });
        let Some(initials) = value else { return };
        let mut element = built_in_element("meta:creator-initials");
        element.push_text(initials);
        let insertion = self
            .children
            .iter()
            .position(|child| metadata_order(child).is_none_or(|other| other > 2))
            .unwrap_or(self.children.len());
        self.children.insert(insertion, element);
    }

    /// Return all ordered child elements, including metadata and rich content.
    #[must_use]
    pub fn children(&self) -> &[Element] {
        &self.children
    }

    /// Return the rich body elements without creator/date/initials metadata.
    #[must_use]
    pub fn body_elements(&self) -> Vec<&Element> {
        self.children
            .iter()
            .filter(|child| !is_metadata_element(child))
            .collect()
    }

    /// Replace the rich body while retaining typed metadata and its schema order.
    ///
    /// # Errors
    ///
    /// Returns an error if the body exceeds the annotation element limit or
    /// refers to an undeclared namespace.
    pub fn replace_body(&mut self, body: Vec<Element>) -> Result<&mut Self> {
        if body.len() > MAX_ANNOTATION_BODY_ELEMENTS {
            return Err(Error::InvalidFormat(
                "annotation body exceeds element limit".to_string(),
            ));
        }
        self.children.retain(is_metadata_element);
        self.children.extend(body);
        self.validate()?;
        Ok(self)
    }

    /// Remove all rich body elements while retaining typed metadata.
    pub fn clear_body(&mut self) -> &mut Self {
        self.children.retain(is_metadata_element);
        self
    }

    /// Append a plain-text paragraph.
    pub fn push_paragraph(&mut self, text: impl Into<String>) -> &mut Self {
        let mut paragraph = built_in_element("text:p");
        paragraph.push_text(text);
        self.children.push(paragraph);
        self
    }

    /// Append a rich paragraph, list, or extension element.
    pub fn push_element(&mut self, element: Element) -> &mut Self {
        self.children.push(element);
        self
    }

    /// Extract annotation body text, separating top-level paragraphs/list blocks.
    #[must_use]
    pub fn text(&self) -> String {
        self.children
            .iter()
            .filter(|child| !is_metadata_element(child))
            .map(Element::plain_text)
            .collect::<Vec<_>>()
            .join("\n")
    }

    /// Validate namespace bindings before serializing this annotation.
    ///
    /// # Errors
    ///
    /// Returns an error when an element or attribute uses an undeclared `XML`
    /// namespace prefix.
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
            .map(Element::plain_text)
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

    fn set_metadata(
        &mut self,
        metadata_name: &'static str,
        metadata_value: Option<&str>,
        order: usize,
    ) {
        let local = local_name(metadata_name);
        self.children
            .retain(|child| local_name(child.name()) != local);
        let Some(value) = metadata_value else { return };

        let mut element = built_in_element(metadata_name);
        element.push_text(value);
        let insertion = self
            .children
            .iter()
            .position(|child| metadata_order(child).is_none_or(|other| other > order))
            .unwrap_or(self.children.len());
        self.children.insert(insertion, element);
    }
}

impl Default for Annotation {
    fn default() -> Self {
        Self::new("")
    }
}

fn built_in_element(name: &'static str) -> Element {
    Element {
        name: name.to_string(),
        attributes: BTreeMap::new(),
        children: Vec::new(),
    }
}

fn set_optional_attribute(
    attributes: &mut BTreeMap<String, String>,
    name: &str,
    attribute_value: Option<&str>,
) {
    if let Some(value) = attribute_value {
        attributes.insert(name.to_string(), value.to_string());
    } else {
        attributes.remove(name);
    }
}

fn metadata_order(element: &Element) -> Option<usize> {
    match local_name(element.name()) {
        "creator" => Some(0),
        "date" => Some(1),
        "creator-initials" | "sender-initials" => Some(2),
        "date-string" => Some(3),
        _ => None,
    }
}

fn is_metadata_element(element: &Element) -> bool {
    metadata_order(element).is_some()
}

pub(crate) fn local_name(name: &str) -> &str {
    name.rsplit_once(':').map_or(name, |(_, local)| local)
}

pub(crate) fn collect_prefix(name: &str, prefixes: &mut BTreeSet<String>) {
    if let Some((prefix, _)) = name.split_once(':')
        && prefix != "xml"
        && prefix != "xmlns"
    {
        prefixes.insert(prefix.to_string());
    }
}

pub(crate) fn standard_namespace_uri(prefix: &str) -> Option<&'static str> {
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

pub(crate) fn validate_qname(name: &str) -> Result<()> {
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

pub(crate) fn valid_name_part(part: &str) -> bool {
    let mut chars = part.chars();
    chars
        .next()
        .is_some_and(|first| first == '_' || first.is_alphabetic())
        && chars
            .all(|ch| ch == '_' || ch == '-' || ch == '.' || ch.is_alphanumeric() || ch == '\u{b7}')
}
