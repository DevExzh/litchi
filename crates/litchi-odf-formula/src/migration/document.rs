//! Inert, namespace-aware MathML access for OpenDocument Formula packages.

use litchi_core::{Error, Result};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesRef, BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

pub(crate) const MATHML_NAMESPACE: &str = "http://www.w3.org/1998/Math/MathML";
const MAX_MATH_DEPTH: usize = 128;
const MAX_MATH_NODES: usize = 65_536;
const MAX_ATTRIBUTES: usize = 256;
const MAX_ATTRIBUTE_BYTES: usize = 1_048_576;
const MAX_TEXT_BYTES: usize = 32 * 1_048_576;

/// A commonly used MathML element kind.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Kind {
    Math,
    Semantics,
    Annotation,
    AnnotationXml,
    Row,
    Identifier,
    Number,
    Operator,
    Text,
    Space,
    StringLiteral,
    Glyph,
    Fraction,
    SquareRoot,
    Root,
    Style,
    Error,
    Padded,
    Phantom,
    Fenced,
    Enclose,
    Subscript,
    Superscript,
    SubSuperscript,
    Under,
    Over,
    UnderOver,
    MultiScripts,
    Table,
    TableRow,
    TableCell,
    AlignGroup,
    AlignMark,
    /// A future MathML element or a vendor element in another namespace.
    Other,
}

/// One decoded attribute with its expanded namespace name.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Attribute {
    namespace_uri: Option<String>,
    local_name: String,
    value: String,
}

impl Attribute {
    pub(crate) fn from_parts(
        namespace_uri: Option<String>,
        local_name: String,
        value: String,
    ) -> Self {
        Self {
            namespace_uri,
            local_name,
            value,
        }
    }

    /// Return the expanded namespace URI, or `None` for an unqualified attribute.
    pub fn namespace_uri(&self) -> Option<&str> {
        self.namespace_uri.as_deref()
    }

    /// Return the XML local name.
    pub fn local_name(&self) -> &str {
        &self.local_name
    }

    /// Return the decoded and normalized XML attribute value.
    pub fn value(&self) -> &str {
        &self.value
    }
}

/// Ordered mixed content within a MathML element.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Content {
    /// Decoded character content, including CDATA and character references.
    ///
    /// Named references other than XML's five predefined entities are retained
    /// in `&name;` notation because MathML 2 documents may declare them in a
    /// document type definition that is intentionally not evaluated here.
    Text(String),
    /// A child element.
    Element(Element),
}

/// A complete element in the formula's MathML subtree.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Element {
    namespace_uri: Option<String>,
    local_name: String,
    attributes: Vec<Attribute>,
    content: Vec<Content>,
}

impl Element {
    pub(crate) fn from_parts(
        namespace_uri: Option<String>,
        local_name: String,
        attributes: Vec<Attribute>,
        content: Vec<Content>,
    ) -> Self {
        Self {
            namespace_uri,
            local_name,
            attributes,
            content,
        }
    }

    pub(crate) fn attributes_mut(&mut self) -> &mut Vec<Attribute> {
        &mut self.attributes
    }

    pub(crate) fn content_mut(&mut self) -> &mut Vec<Content> {
        &mut self.content
    }

    /// Return the element's expanded namespace URI.
    pub fn namespace_uri(&self) -> Option<&str> {
        self.namespace_uri.as_deref()
    }

    /// Return the element's XML local name.
    pub fn local_name(&self) -> &str {
        &self.local_name
    }

    /// Classify common MathML elements without discarding unknown ones.
    pub fn kind(&self) -> Kind {
        if self.namespace_uri() != Some(MATHML_NAMESPACE) {
            return Kind::Other;
        }
        match self.local_name.as_str() {
            "math" => Kind::Math,
            "semantics" => Kind::Semantics,
            "annotation" => Kind::Annotation,
            "annotation-xml" => Kind::AnnotationXml,
            "mrow" => Kind::Row,
            "mi" => Kind::Identifier,
            "mn" => Kind::Number,
            "mo" => Kind::Operator,
            "mtext" => Kind::Text,
            "mspace" => Kind::Space,
            "ms" => Kind::StringLiteral,
            "mglyph" => Kind::Glyph,
            "mfrac" => Kind::Fraction,
            "msqrt" => Kind::SquareRoot,
            "mroot" => Kind::Root,
            "mstyle" => Kind::Style,
            "merror" => Kind::Error,
            "mpadded" => Kind::Padded,
            "mphantom" => Kind::Phantom,
            "mfenced" => Kind::Fenced,
            "menclose" => Kind::Enclose,
            "msub" => Kind::Subscript,
            "msup" => Kind::Superscript,
            "msubsup" => Kind::SubSuperscript,
            "munder" => Kind::Under,
            "mover" => Kind::Over,
            "munderover" => Kind::UnderOver,
            "mmultiscripts" => Kind::MultiScripts,
            "mtable" => Kind::Table,
            "mtr" | "mlabeledtr" => Kind::TableRow,
            "mtd" => Kind::TableCell,
            "maligngroup" => Kind::AlignGroup,
            "malignmark" => Kind::AlignMark,
            _ => Kind::Other,
        }
    }

    /// Return all decoded attributes in document order.
    pub fn attributes(&self) -> &[Attribute] {
        &self.attributes
    }

    /// Find an attribute by expanded name.
    pub fn attribute(&self, namespace_uri: Option<&str>, local_name: &str) -> Option<&str> {
        self.attributes
            .iter()
            .find(|attribute| {
                attribute.namespace_uri() == namespace_uri && attribute.local_name == local_name
            })
            .map(Attribute::value)
    }

    /// Return ordered mixed content.
    pub fn content(&self) -> &[Content] {
        &self.content
    }

    /// Iterate direct child elements.
    pub fn children(&self) -> impl Iterator<Item = &Element> {
        self.content.iter().filter_map(|content| match content {
            Content::Element(element) => Some(element),
            Content::Text(_) => None,
        })
    }

    /// Compose all descendant character content in exact element/text order.
    pub fn all_text(&self) -> String {
        fn append(element: &Element, output: &mut String) {
            for content in &element.content {
                match content {
                    Content::Text(text) => output.push_str(text),
                    Content::Element(child) => append(child, output),
                }
            }
        }
        let mut output = String::new();
        append(self, &mut output);
        output
    }

    pub(crate) fn collect_annotations<'a>(&'a self, output: &mut Vec<&'a Element>) {
        if matches!(self.kind(), Kind::Annotation | Kind::AnnotationXml) {
            output.push(self);
        }
        for child in self.children() {
            child.collect_annotations(output);
        }
    }
}

pub(crate) fn parse_mathml(xml: &str) -> Result<Element> {
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut stack = Vec::new();
    let mut root = None;
    let mut root_closed = false;
    let mut node_count = 0usize;
    let mut text_bytes = 0usize;

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid formula MathML: {error}")))?;
        match event {
            Event::Start(ref element) => {
                if root_closed {
                    return Err(Error::InvalidFormat(
                        "formula contains multiple root elements".to_string(),
                    ));
                }
                let resolved_namespace_uri = namespace_uri(&namespace)?;
                let node = make_element(&reader, resolved_namespace_uri, element, &mut node_count)?;
                if stack.is_empty()
                    && (node.namespace_uri() != Some(MATHML_NAMESPACE)
                        || node.local_name() != "math")
                {
                    return Err(Error::InvalidFormat(
                        "formula content must have a MathML math root".to_string(),
                    ));
                }
                stack.push(node);
                if stack.len() > MAX_MATH_DEPTH {
                    return Err(Error::InvalidFormat(format!(
                        "MathML nesting exceeds {MAX_MATH_DEPTH} levels"
                    )));
                }
            },
            Event::Empty(ref element) => {
                if stack.is_empty() {
                    if root_closed {
                        return Err(Error::InvalidFormat(
                            "formula contains multiple root elements".to_string(),
                        ));
                    }
                    let resolved_namespace_uri = namespace_uri(&namespace)?;
                    let node =
                        make_element(&reader, resolved_namespace_uri, element, &mut node_count)?;
                    if node.namespace_uri() != Some(MATHML_NAMESPACE) || node.local_name() != "math"
                    {
                        return Err(Error::InvalidFormat(
                            "formula content must have a MathML math root".to_string(),
                        ));
                    }
                    root = Some(node);
                    root_closed = true;
                    buffer.clear();
                    continue;
                }
                let resolved_namespace_uri = namespace_uri(&namespace)?;
                let node = make_element(&reader, resolved_namespace_uri, element, &mut node_count)?;
                stack
                    .last_mut()
                    .expect("parent exists")
                    .content
                    .push(Content::Element(node));
            },
            Event::End(_) => {
                let node = stack.pop().ok_or_else(|| {
                    Error::InvalidFormat("MathML element stack underflow".to_string())
                })?;
                if let Some(parent) = stack.last_mut() {
                    parent.content.push(Content::Element(node));
                } else {
                    if root.is_some() {
                        return Err(Error::InvalidFormat(
                            "formula contains multiple MathML roots".to_string(),
                        ));
                    }
                    root = Some(node);
                    root_closed = true;
                }
            },
            Event::Text(ref text) if !stack.is_empty() => {
                let value = text.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                    Error::InvalidFormat(format!("invalid MathML text: {error}"))
                })?;
                push_text(
                    stack.last_mut().expect("element exists"),
                    value.into_owned(),
                    &mut text_bytes,
                )?;
            },
            Event::CData(ref text) if !stack.is_empty() => {
                let value = text.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                    Error::InvalidFormat(format!("invalid MathML CDATA: {error}"))
                })?;
                push_text(
                    stack.last_mut().expect("element exists"),
                    value.into_owned(),
                    &mut text_bytes,
                )?;
            },
            Event::GeneralRef(ref reference) if !stack.is_empty() => {
                push_text(
                    stack.last_mut().expect("element exists"),
                    decode_reference(reference)?,
                    &mut text_bytes,
                )?;
            },
            Event::Text(ref text) if !text.iter().all(u8::is_ascii_whitespace) => {
                return Err(Error::InvalidFormat(
                    "text is not allowed outside the MathML root".to_string(),
                ));
            },
            Event::CData(_) | Event::GeneralRef(_) if stack.is_empty() => {
                return Err(Error::InvalidFormat(
                    "content is not allowed outside the MathML root".to_string(),
                ));
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if !stack.is_empty() || !root_closed {
        return Err(Error::InvalidFormat(
            "formula contains incomplete MathML".to_string(),
        ));
    }
    root.ok_or_else(|| Error::InvalidFormat("formula has no MathML root".to_string()))
}

fn make_element(
    reader: &NsReader<&[u8]>,
    resolved_namespace_uri: Option<String>,
    element: &BytesStart<'_>,
    node_count: &mut usize,
) -> Result<Element> {
    *node_count = node_count
        .checked_add(1)
        .ok_or_else(|| Error::InvalidFormat("MathML node count overflow".to_string()))?;
    if *node_count > MAX_MATH_NODES {
        return Err(Error::InvalidFormat(format!(
            "formula exceeds {MAX_MATH_NODES} MathML elements"
        )));
    }
    if element.attributes().count() > MAX_ATTRIBUTES {
        return Err(Error::InvalidFormat(format!(
            "MathML element exceeds {MAX_ATTRIBUTES} attributes"
        )));
    }
    let local_name = decode_utf8(element.local_name().as_ref(), "element name")?;
    let mut attributes = Vec::new();
    for attribute in element.attributes() {
        let attribute = attribute
            .map_err(|error| Error::InvalidFormat(format!("invalid MathML attribute: {error}")))?;
        if attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace_uri = namespace_uri(&namespace)?;
        let local_name = decode_utf8(local.as_ref(), "attribute name")?;
        if attributes.iter().any(|existing: &Attribute| {
            existing.namespace_uri == namespace_uri && existing.local_name == local_name
        }) {
            return Err(Error::InvalidFormat(format!(
                "duplicate expanded MathML attribute '{local_name}'"
            )));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid MathML attribute value: {error}"))
            })?
            .into_owned();
        if value.len() > MAX_ATTRIBUTE_BYTES {
            return Err(Error::InvalidFormat(
                "MathML attribute exceeds 1 MiB".to_string(),
            ));
        }
        attributes.push(Attribute {
            namespace_uri,
            local_name,
            value,
        });
    }
    Ok(Element {
        namespace_uri: resolved_namespace_uri,
        local_name,
        attributes,
        content: Vec::new(),
    })
}

fn push_text(element: &mut Element, value: String, total: &mut usize) -> Result<()> {
    *total = total
        .checked_add(value.len())
        .ok_or_else(|| Error::InvalidFormat("MathML text size overflow".to_string()))?;
    if *total > MAX_TEXT_BYTES {
        return Err(Error::InvalidFormat(
            "formula exceeds 32 MiB of MathML text".to_string(),
        ));
    }
    if let Some(Content::Text(existing)) = element.content.last_mut() {
        existing.push_str(&value);
    } else {
        element.content.push(Content::Text(value));
    }
    Ok(())
}

fn namespace_uri(namespace: &ResolveResult<'_>) -> Result<Option<String>> {
    match namespace {
        ResolveResult::Unbound => Ok(None),
        ResolveResult::Bound(Namespace(uri)) => decode_utf8(uri, "namespace URI").map(Some),
        ResolveResult::Unknown(prefix) => Err(Error::InvalidFormat(format!(
            "unknown MathML namespace prefix '{}'",
            String::from_utf8_lossy(prefix)
        ))),
    }
}

fn decode_utf8(bytes: &[u8], kind: &str) -> Result<String> {
    std::str::from_utf8(bytes)
        .map(str::to_string)
        .map_err(|_| Error::InvalidFormat(format!("non-UTF-8 MathML {kind}")))
}

fn decode_reference(reference: &BytesRef<'_>) -> Result<String> {
    if let Some(character) = reference.resolve_char_ref().map_err(|error| {
        Error::InvalidFormat(format!("invalid MathML character reference: {error}"))
    })? {
        return Ok(character.to_string());
    }
    let name = reference
        .decode()
        .map_err(|error| Error::InvalidFormat(format!("invalid MathML entity: {error}")))?;
    match name.as_ref() {
        "amp" => Ok("&".to_string()),
        "lt" => Ok("<".to_string()),
        "gt" => Ok(">".to_string()),
        "quot" => Ok("\"".to_string()),
        "apos" => Ok("'".to_string()),
        _ => Ok(format!("&{name};")),
    }
}
