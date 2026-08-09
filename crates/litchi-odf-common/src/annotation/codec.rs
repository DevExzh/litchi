//! Bounded event parsing and deterministic XML serialization for annotations.

use super::model::{Annotation, Element, Node, collect_prefix, validate_qname};
use super::package::{
    MAX_ANNOTATION_ATTRIBUTE_BYTES, MAX_ANNOTATION_ATTRIBUTES, MAX_ANNOTATION_ELEMENTS,
    MAX_ANNOTATION_NESTING, MAX_ANNOTATION_TEXT_BYTES,
};
use litchi_core::{Error, Result, xml::escape_xml};
use quick_xml::{
    Decoder, XmlVersion,
    events::{BytesCData, BytesRef, BytesStart, BytesText},
};
use std::collections::{BTreeMap, BTreeSet};

/// Stateful builder used by package readers while they walk one annotation.
pub struct Builder {
    annotation: Annotation,
    stack: Vec<Element>,
    element_count: usize,
    text_bytes: usize,
}

impl Builder {
    /// Create a builder from an `office:annotation` start element.
    ///
    /// # Errors
    ///
    /// Returns an error when the annotation attributes are malformed or exceed
    /// the configured resource limits.
    pub fn new(
        start: &BytesStart<'_>,
        decoder: Decoder,
        namespaces: BTreeMap<String, String>,
    ) -> Result<Self> {
        Ok(Self {
            annotation: Annotation {
                attributes: parse_attributes(start, decoder, false)?,
                children: Vec::new(),
                namespaces,
            },
            stack: Vec::new(),
            element_count: 0,
            text_bytes: 0,
        })
    }

    /// Begin a nested annotation element.
    ///
    /// # Errors
    ///
    /// Returns an error when the element is malformed or exceeds a resource
    /// limit.
    pub fn start(&mut self, start: &BytesStart<'_>, decoder: Decoder) -> Result<()> {
        let element = parse_element(start, decoder)?;
        self.reserve_element()?;
        self.stack.push(element);
        Ok(())
    }

    /// Add an empty nested annotation element.
    ///
    /// # Errors
    ///
    /// Returns an error when the element is malformed or exceeds a resource
    /// limit.
    pub fn empty(&mut self, start: &BytesStart<'_>, decoder: Decoder) -> Result<()> {
        let element = parse_element(start, decoder)?;
        self.reserve_element()?;
        self.add_element(element);
        Ok(())
    }

    /// Append decoded XML character data.
    ///
    /// # Errors
    ///
    /// Returns an error when the XML text is invalid or exceeds a resource
    /// limit.
    pub fn text(&mut self, text: &BytesText<'_>) -> Result<()> {
        let decoded = text
            .xml_content(XmlVersion::Explicit1_0)
            .map_err(|error| Error::InvalidFormat(format!("invalid annotation text: {error}")))?;
        self.append_text(decoded.into_owned())
    }

    /// Append a decoded XML entity or character reference.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid or unsupported XML reference.
    pub fn reference(&mut self, reference: &BytesRef<'_>) -> Result<()> {
        let value = decode_reference(reference)?;
        self.append_text(value)
    }

    /// Append decoded XML CDATA.
    ///
    /// # Errors
    ///
    /// Returns an error when the CDATA is invalid or exceeds a resource limit.
    pub fn cdata(&mut self, text: &BytesCData<'_>) -> Result<()> {
        let decoded = text
            .xml_content(XmlVersion::Explicit1_0)
            .map_err(|error| Error::InvalidFormat(format!("invalid annotation CDATA: {error}")))?;
        self.append_text(decoded.into_owned())
    }

    /// Complete the innermost nested element.
    ///
    /// # Errors
    ///
    /// Returns an error if no element is open.
    pub fn end_element(&mut self) -> Result<()> {
        let element = self
            .stack
            .pop()
            .ok_or_else(|| Error::InvalidFormat("unbalanced element in annotation".to_string()))?;
        self.add_element(element);
        Ok(())
    }

    /// Finish the annotation after verifying that all elements are closed.
    ///
    /// # Errors
    ///
    /// Returns an error when an element remains unclosed.
    pub fn finish(self) -> Result<Annotation> {
        if !self.stack.is_empty() {
            return Err(Error::InvalidFormat(
                "unclosed element in annotation".to_string(),
            ));
        }
        Ok(self.annotation)
    }

    fn reserve_element(&mut self) -> Result<()> {
        if self.stack.len() >= MAX_ANNOTATION_NESTING {
            return Err(Error::InvalidFormat(
                "annotation XML exceeds the nesting limit".to_string(),
            ));
        }
        if self.element_count >= MAX_ANNOTATION_ELEMENTS {
            return Err(Error::InvalidFormat(
                "annotation XML exceeds the element limit".to_string(),
            ));
        }
        self.element_count += 1;
        Ok(())
    }

    fn append_text(&mut self, value: String) -> Result<()> {
        let Some(parent) = self.stack.last_mut() else {
            return Ok(());
        };
        let text_bytes = self
            .text_bytes
            .checked_add(value.len())
            .ok_or_else(|| Error::InvalidFormat("annotation text size overflow".to_string()))?;
        if text_bytes > MAX_ANNOTATION_TEXT_BYTES {
            return Err(Error::InvalidFormat(
                "annotation text exceeds the size limit".to_string(),
            ));
        }
        self.text_bytes = text_bytes;
        parent.children.push(Node::Text(value));
        Ok(())
    }

    fn add_element(&mut self, element: Element) {
        if let Some(parent) = self.stack.last_mut() {
            parent.children.push(Node::Element(element));
        } else {
            self.annotation.children.push(element);
        }
    }
}

impl Element {
    pub fn write_xml(&self, output: &mut String) {
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
                Node::Text(text) => output.push_str(&escape_xml(text)),
                Node::Element(element) => element.write_xml(output),
            }
        }
        output.push_str("</");
        output.push_str(&self.name);
        output.push('>');
    }
}

impl Annotation {
    pub fn write_xml(&self, output: &mut String) {
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

/// Serialize XML attributes in deterministic map order.
pub fn write_attributes(output: &mut String, attributes: &BTreeMap<String, String>) {
    for (name, value) in attributes {
        output.push(' ');
        output.push_str(name);
        output.push_str("=\"");
        output.push_str(&escape_xml(value));
        output.push('"');
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

pub(crate) fn parse_element(start: &BytesStart<'_>, decoder: Decoder) -> Result<Element> {
    let name = std::str::from_utf8(start.name().as_ref())
        .map_err(|error| {
            Error::InvalidFormat(format!("invalid UTF-8 in annotation element name: {error}"))
        })?
        .to_string();
    validate_qname(&name)?;
    Ok(Element {
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
    let mut attribute_bytes = 0usize;
    for raw_attribute in start.attributes() {
        let attribute = raw_attribute.map_err(|error| {
            Error::InvalidFormat(format!("invalid annotation attribute: {error}"))
        })?;
        let name = std::str::from_utf8(attribute.key.as_ref())
            .map_err(|error| {
                Error::InvalidFormat(format!(
                    "invalid UTF-8 in annotation attribute name: {error}"
                ))
            })?
            .to_string();
        if !keep_namespaces && (name == "xmlns" || name.starts_with("xmlns:")) {
            continue;
        }
        if attributes.len() >= MAX_ANNOTATION_ATTRIBUTES {
            return Err(Error::InvalidFormat(
                "annotation attribute count exceeds the safety limit".to_string(),
            ));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid annotation attribute value: {error}"))
            })?
            .into_owned();
        attribute_bytes = attribute_bytes
            .checked_add(name.len())
            .and_then(|size| size.checked_add(value.len()))
            .ok_or_else(|| {
                Error::InvalidFormat("annotation attribute size overflow".to_string())
            })?;
        if attribute_bytes > MAX_ANNOTATION_ATTRIBUTE_BYTES {
            return Err(Error::InvalidFormat(
                "annotation attributes exceed the size limit".to_string(),
            ));
        }
        attributes.insert(name, value);
    }
    Ok(attributes)
}
