//! In-content RDFa and `text:meta` metadata.
//!
//! ODF has two metadata surfaces which are deliberately separate from
//! `meta.xml` and from package RDF graphs: RDFa attributes attached to text
//! content and the inline `text:meta` element.  This module owns the small,
//! inert typed view of those declarations.  RDFa CURIEs and datatypes are
//! retained as strings; they are never resolved, fetched, or evaluated.

#![allow(
    clippy::arbitrary_source_item_ordering,
    clippy::cast_possible_truncation,
    clippy::doc_markdown,
    clippy::format_push_string,
    clippy::items_after_statements,
    clippy::manual_let_else,
    clippy::missing_errors_doc,
    clippy::must_use_candidate,
    clippy::needless_pass_by_value,
    clippy::similar_names,
    reason = "the metadata codec is a bounded XML projection with compatibility-oriented naming"
)]

use crate::elements::field::{
    MetaFieldAttribute, MetaFieldContent, MetaFieldElement, MetaFieldNode,
};
use crate::namespace::{
    DCNS, DR3DNS, DRAWNS, FONS, FORMNS, METANS, NUMBERNS, OFFICENS, PRESENTATIONNS, SCRIPTNS,
    STYLENS, SVGNS, TABLENS, TEXTNS, XHTMLNS, XLINKNS, XMLNS, XSDNS,
};
use litchi_core::{Error, Position, Result, xml::escape_xml};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{QName, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::{HashMap, HashSet};

const MAX_XML_BYTES: usize = 64 * 1024 * 1024;
const MAX_DEPTH: usize = 512;
const MAX_OCCURRENCES: usize = 1_000_000;
const MAX_ATTRIBUTES: usize = 256;
const MAX_TEXT_META_NODES: usize = 1_000_000;
const MAX_TEXT_META_DEPTH: usize = 256;
const MAX_TEXT_META_BYTES: usize = 16 * 1024 * 1024;

/// The XML part in which an in-content declaration was found.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum MetadataPart {
    /// `content.xml`.
    Content,
    /// `styles.xml`.
    Styles,
}

/// RDFa attributes defined by ODF's in-content metadata profile.
#[derive(Debug, Clone, Default, PartialEq, Eq, Hash)]
pub struct RdfaAttributes {
    /// RDFa `about` resource.
    pub about: Option<String>,
    /// RDFa `property` CURIE or CURIE list.
    pub property: Option<String>,
    /// RDFa `content` lexical value.
    pub content: Option<String>,
    /// RDFa `datatype` lexical datatype.
    pub datatype: Option<String>,
}

/// An unmodeled attribute retained on a `text:meta` root.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct MetadataAttribute {
    /// Namespace URI, or empty for an unqualified attribute.
    pub namespace_uri: String,
    /// Local attribute name.
    pub local_name: String,
    /// Lexical value.
    pub value: String,
    /// Source prefix when the attribute was qualified.
    pub prefix: Option<String>,
}

impl RdfaAttributes {
    /// Return whether no RDFa attribute is present.
    pub fn is_empty(&self) -> bool {
        self.about.is_none()
            && self.property.is_none()
            && self.content.is_none()
            && self.datatype.is_none()
    }

    /// Validate XML lexical limits without resolving RDFa values.
    pub fn validate(&self) -> Result<()> {
        for (name, value) in [
            ("xhtml:about", self.about.as_deref()),
            ("xhtml:property", self.property.as_deref()),
            ("xhtml:content", self.content.as_deref()),
            ("xhtml:datatype", self.datatype.as_deref()),
        ] {
            if let Some(value) = value {
                validate_xml_value(name, value, MAX_TEXT_META_BYTES)?;
            }
        }
        Ok(())
    }
}

/// Semantic host for one RDFa occurrence.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum RdfaHost {
    /// A `text:p` in document order within its part.
    Paragraph { index: usize },
    /// A `text:h` in document order within its part.
    Heading { index: usize },
    /// A `text:bookmark-start`; the optional name is retained as written.
    BookmarkStart { name: Option<String> },
    /// A `text:meta` element.  Its full typed value is also available in
    /// [`ContentMetadata::text_meta`].
    TextMeta { index: usize },
    /// Another namespace-resolved XML host carrying RDFa attributes.
    Element {
        namespace_uri: String,
        local_name: String,
        occurrence: usize,
    },
}

/// One in-content RDFa declaration.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct RdfaOccurrence {
    /// XML part containing the host.
    pub part: MetadataPart,
    /// Host element identity.
    pub host: RdfaHost,
    /// Retained RDFa attributes.
    pub attributes: RdfaAttributes,
}

/// A typed `text:meta` inline element.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct TextMeta {
    /// XML part containing the element.
    pub part: MetadataPart,
    /// Zero-based `text:meta` occurrence in the part.
    pub index: usize,
    /// Optional `xml:id`.
    pub xml_id: Option<String>,
    /// RDFa attributes attached to the element.
    pub rdfa: RdfaAttributes,
    /// Unknown root attributes retained by expanded name.
    pub attributes: Vec<MetadataAttribute>,
    /// Inert, validated paragraph-content mixed content.
    pub content: MetaFieldContent,
}

impl TextMeta {
    /// Construct an inline metadata element for a new document or insertion.
    pub fn new(content: MetaFieldContent) -> Result<Self> {
        content.validate()?;
        Ok(Self {
            part: MetadataPart::Content,
            index: 0,
            xml_id: None,
            rdfa: RdfaAttributes::default(),
            attributes: Vec::new(),
            content,
        })
    }

    /// Construct a plain-text inline metadata element.
    pub fn from_text(text: impl Into<String>) -> Result<Self> {
        Self::new(MetaFieldContent::new(vec![MetaFieldNode::Text(
            text.into(),
        )])?)
    }

    /// Validate attributes, identity, and mixed content.
    pub fn validate(&self) -> Result<()> {
        self.rdfa.validate()?;
        validate_metadata_attributes(&self.attributes)?;
        if let Some(id) = &self.xml_id {
            validate_xml_id(id)?;
        }
        self.content.validate()
    }

    fn to_xml_with_context(&self, namespace_declarations: &[String]) -> Result<String> {
        self.validate()?;
        let output_len = metadata_output_len(self, namespace_declarations)?;
        bounded_output_len(output_len, "text:meta serialization")?;
        let mut output = String::new();
        output
            .try_reserve_exact(output_len)
            .map_err(|source| Error::Allocation {
                resource: "ODT text:meta serialization",
                source,
            })?;
        output.push_str("<text:meta xmlns:text=\"");
        output.push_str(TEXTNS);
        output.push_str("\" xmlns:xhtml=\"");
        output.push_str(XHTMLNS);
        output.push('"');
        if let Some(id) = &self.xml_id {
            output.push_str(" xml:id=\"");
            output.push_str(&escape_xml(id));
            output.push('"');
        }
        push_rdfa_namespace_declarations(&mut output, &self.rdfa, namespace_declarations)?;
        push_rdfa_attributes(&mut output, &self.rdfa);
        push_metadata_attributes(&mut output, &self.attributes);
        if self.content.nodes().is_empty() {
            output.push_str("/>");
        } else {
            output.push('>');
            self.content.write_xml_to(&mut output);
            output.push_str("</text:meta>");
        }
        Ok(output)
    }
}

/// Bounded typed in-content metadata from one ODF document package.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ContentMetadata {
    /// RDFa declarations in document and style XML parts.
    pub rdfa: Vec<RdfaOccurrence>,
    /// Inline `text:meta` elements in document and style XML parts.
    pub text_meta: Vec<TextMeta>,
}

impl ContentMetadata {
    /// Return whether no in-content metadata was found.
    pub fn is_empty(&self) -> bool {
        self.rdfa.is_empty() && self.text_meta.is_empty()
    }

    /// Validate all retained metadata.
    pub fn validate(&self) -> Result<()> {
        for item in &self.rdfa {
            item.attributes.validate()?;
        }
        for item in &self.text_meta {
            item.validate()?
        }
        Ok(())
    }
}

/// Parse in-content metadata from the supplied XML parts.
pub(crate) fn parse_parts(parts: &[(&str, MetadataPart)]) -> Result<ContentMetadata> {
    let mut output = ContentMetadata::default();
    for (xml, part) in parts {
        let parsed = parse_part(xml, *part)?;
        for item in parsed.rdfa {
            push_bounded(
                &mut output.rdfa,
                item,
                "ODF aggregate in-content RDFa occurrence",
            )?;
        }
        for item in parsed.text_meta {
            push_bounded(
                &mut output.text_meta,
                item,
                "ODF aggregate text:meta projection",
            )?;
        }
    }
    Ok(output)
}

/// Parse one XML part.  This is crate-visible so mutable editors can inspect
/// their authoritative content snapshot without retaining a second copy.
pub(crate) fn parse_part(xml: &str, part: MetadataPart) -> Result<ContentMetadata> {
    if xml.len() > MAX_XML_BYTES {
        return Err(Error::InvalidFormat(format!(
            "ODF XML exceeds the {MAX_XML_BYTES} in-content metadata limit"
        )));
    }
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut stack: Vec<(Option<String>, String)> = Vec::new();
    let mut active: Vec<ActiveTextMeta> = Vec::new();
    let mut output = ContentMetadata::default();
    let mut paragraph_index = 0usize;
    let mut heading_index = 0usize;
    let mut text_meta_index = 0usize;
    let mut other_occurrences: HashMap<(String, String), usize> = HashMap::new();
    let mut seen = 0usize;

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid ODF metadata XML: {error}")))?;
        let namespace_uri = resolved_namespace(&namespace)?;
        match event {
            Event::Start(ref source) => {
                depth = depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("in-content metadata XML depth overflow".to_string())
                })?;
                if depth > MAX_DEPTH {
                    return Err(Error::InvalidFormat(format!(
                        "in-content metadata XML exceeds {MAX_DEPTH} levels"
                    )));
                }
                let local = utf8(source.local_name().as_ref(), "metadata element name")?;
                let attrs = parse_attributes(&reader, source)?;
                let rdfa = rdfa_from_attributes(&attrs)?;
                let host = host_for(
                    namespace_uri.as_deref(),
                    &local,
                    &attrs,
                    &mut paragraph_index,
                    &mut heading_index,
                    &mut text_meta_index,
                    &mut other_occurrences,
                )?;
                if let Some(ref host) = host {
                    if !rdfa.is_empty() {
                        push_bounded(
                            &mut output.rdfa,
                            RdfaOccurrence {
                                part,
                                host: host.clone(),
                                attributes: rdfa.clone(),
                            },
                            "ODF in-content RDFa occurrence",
                        )?;
                    }
                }
                // A metadata element contributes its complete root to every
                // active outer `text:meta` builder before becoming a new root.
                if !active.is_empty() {
                    let child_attributes = parse_meta_attributes(&reader, source)?;
                    let child_namespace = namespace_uri.as_deref().ok_or_else(|| {
                        Error::InvalidFormat(
                            "unqualified element inside text:meta is not supported".to_string(),
                        )
                    })?;
                    for root in &mut active {
                        root.builder.start_element(
                            child_namespace.to_string(),
                            local.clone(),
                            child_attributes.clone(),
                        )?;
                    }
                }
                let is_meta = namespace_uri.as_deref() == Some(TEXTNS) && local == "meta";
                if is_meta {
                    let (xml_id, root_rdfa, root_attributes) =
                        text_meta_root_attributes(&reader, source, &attrs)?;
                    let order = match host {
                        Some(RdfaHost::TextMeta { index }) => index,
                        _ => {
                            return Err(Error::InvalidFormat(
                                "text:meta host index missing".to_string(),
                            ));
                        },
                    };
                    push_bounded(
                        &mut active,
                        ActiveTextMeta {
                            depth,
                            part,
                            index: order,
                            xml_id,
                            rdfa: root_rdfa,
                            attributes: root_attributes,
                            builder: MetaBuilder::default(),
                        },
                        "ODF active text:meta projection",
                    )?;
                }
                stack.push((namespace_uri, local));
            },
            Event::Empty(ref source) => {
                let local = utf8(source.local_name().as_ref(), "metadata element name")?;
                let attrs = parse_attributes(&reader, source)?;
                let rdfa = rdfa_from_attributes(&attrs)?;
                let host = host_for(
                    namespace_uri.as_deref(),
                    &local,
                    &attrs,
                    &mut paragraph_index,
                    &mut heading_index,
                    &mut text_meta_index,
                    &mut other_occurrences,
                )?;
                if let Some(ref host) = host {
                    if !rdfa.is_empty() {
                        push_bounded(
                            &mut output.rdfa,
                            RdfaOccurrence {
                                part,
                                host: host.clone(),
                                attributes: rdfa.clone(),
                            },
                            "ODF in-content RDFa occurrence",
                        )?;
                    }
                }
                if !active.is_empty() {
                    let child_attributes = parse_meta_attributes(&reader, source)?;
                    let child_namespace = namespace_uri.as_deref().ok_or_else(|| {
                        Error::InvalidFormat(
                            "unqualified element inside text:meta is not supported".to_string(),
                        )
                    })?;
                    for root in &mut active {
                        root.builder.empty_element(
                            child_namespace.to_string(),
                            local.clone(),
                            child_attributes.clone(),
                        )?;
                    }
                }
                if namespace_uri.as_deref() == Some(TEXTNS) && local == "meta" {
                    let (xml_id, root_rdfa, root_attributes) =
                        text_meta_root_attributes(&reader, source, &attrs)?;
                    let index = match host {
                        Some(RdfaHost::TextMeta { index }) => index,
                        _ => {
                            return Err(Error::InvalidFormat(
                                "text:meta host index missing".to_string(),
                            ));
                        },
                    };
                    let value = TextMeta {
                        part,
                        index,
                        xml_id,
                        rdfa: root_rdfa,
                        attributes: root_attributes,
                        content: MetaFieldContent::new(Vec::new())?,
                    };
                    push_bounded(
                        &mut output.text_meta,
                        value,
                        "ODF completed text:meta projection",
                    )?;
                }
            },
            Event::Text(ref value) => {
                let value = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid text:meta character data: {error}"))
                    })?;
                for root in &mut active {
                    root.builder.text(&value)?;
                }
            },
            Event::CData(ref value) => {
                let value = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid text:meta CDATA: {error}"))
                    })?;
                for root in &mut active {
                    root.builder.text(&value)?;
                }
            },
            Event::GeneralRef(ref value) => {
                let reference = std::str::from_utf8(value.as_ref()).map_err(|_| {
                    Error::InvalidFormat("invalid XML entity reference in text:meta".to_string())
                })?;
                let value = quick_xml::escape::resolve_xml_entity(reference).ok_or_else(|| {
                    Error::InvalidFormat("unknown XML entity in text:meta".to_string())
                })?;
                for root in &mut active {
                    root.builder.text(value)?;
                }
            },
            Event::End(_) => {
                for root in &mut active {
                    if root.depth < depth {
                        root.builder.end_element()?;
                    }
                }
                if let Some(root) = active.pop_if(|root| root.depth == depth) {
                    let content = root.builder.finish()?;
                    let value = TextMeta {
                        part: root.part,
                        index: root.index,
                        xml_id: root.xml_id,
                        rdfa: root.rdfa,
                        attributes: root.attributes,
                        content,
                    };
                    push_bounded(
                        &mut output.text_meta,
                        value,
                        "ODF completed text:meta projection",
                    )?;
                }
                stack.pop().ok_or_else(|| {
                    Error::InvalidFormat("in-content metadata XML stack underflow".to_string())
                })?;
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("in-content metadata XML depth underflow".to_string())
                })?;
            },
            Event::DocType(_) => {
                return Err(Error::InvalidFormat(
                    "DOCTYPE is not permitted in ODF metadata XML".to_string(),
                ));
            },
            Event::Comment(_) | Event::PI(_) if !active.is_empty() => {
                return Err(Error::InvalidFormat(
                    "comments and processing instructions inside text:meta are not supported by the typed editor".to_string(),
                ));
            },
            Event::Eof => break,
            _ => {},
        }
        seen = seen.checked_add(1).ok_or_else(|| {
            Error::InvalidFormat("in-content metadata event count overflow".to_string())
        })?;
        if seen > MAX_OCCURRENCES {
            return Err(Error::InvalidFormat(format!(
                "in-content metadata exceeds {MAX_OCCURRENCES} XML events"
            )));
        }
        buffer.clear();
    }
    if depth != 0 || !stack.is_empty() || !active.is_empty() {
        return Err(Error::InvalidFormat(
            "incomplete in-content metadata XML".to_string(),
        ));
    }
    output.text_meta.sort_by_key(|value| value.index);
    Ok(output)
}

#[derive(Debug)]
struct ActiveTextMeta {
    depth: usize,
    part: MetadataPart,
    index: usize,
    xml_id: Option<String>,
    rdfa: RdfaAttributes,
    attributes: Vec<MetadataAttribute>,
    builder: MetaBuilder,
}

#[derive(Debug, Default)]
struct MetaBuilder {
    roots: Vec<MetaFieldNode>,
    stack: Vec<MetaFieldElement>,
    nodes: usize,
    bytes: usize,
}

impl MetaBuilder {
    fn text(&mut self, value: &str) -> Result<()> {
        self.add_node(value.len())?;
        if let Some(MetaFieldNode::Text(existing)) = self.current_mut().last_mut() {
            existing
                .try_reserve(value.len())
                .map_err(|source| Error::Allocation {
                    resource: "ODT text:meta text projection",
                    source,
                })?;
            existing.push_str(value);
        } else {
            let current = self.current_mut();
            current.try_reserve(1).map_err(|source| Error::Allocation {
                resource: "ODT text:meta node projection",
                source,
            })?;
            let mut text = String::new();
            text.try_reserve(value.len())
                .map_err(|source| Error::Allocation {
                    resource: "ODT text:meta text projection",
                    source,
                })?;
            text.push_str(value);
            current.push(MetaFieldNode::Text(text));
        }
        Ok(())
    }

    fn start_element(
        &mut self,
        namespace_uri: String,
        local_name: String,
        attributes: Vec<MetaFieldAttribute>,
    ) -> Result<()> {
        self.add_node(namespace_uri.len().saturating_add(local_name.len()))?;
        if self.stack.len() >= MAX_TEXT_META_DEPTH {
            return Err(Error::InvalidFormat(format!(
                "text:meta content exceeds {MAX_TEXT_META_DEPTH} levels"
            )));
        }
        self.stack
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "ODT text:meta element stack",
                source,
            })?;
        self.stack.push(MetaFieldElement {
            namespace_uri,
            local_name,
            attributes,
            children: Vec::new(),
        });
        Ok(())
    }

    fn empty_element(
        &mut self,
        namespace_uri: String,
        local_name: String,
        attributes: Vec<MetaFieldAttribute>,
    ) -> Result<()> {
        self.start_element(namespace_uri, local_name, attributes)?;
        self.end_element()
    }

    fn end_element(&mut self) -> Result<()> {
        let element = self
            .stack
            .pop()
            .ok_or_else(|| Error::InvalidFormat("text:meta content stack underflow".to_string()))?;
        self.current_mut()
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "ODT text:meta node projection",
                source,
            })?;
        self.current_mut().push(MetaFieldNode::Element(element));
        Ok(())
    }

    fn finish(self) -> Result<MetaFieldContent> {
        if !self.stack.is_empty() {
            return Err(Error::InvalidFormat(
                "incomplete text:meta mixed content".to_string(),
            ));
        }
        MetaFieldContent::new(self.roots)
    }

    fn current_mut(&mut self) -> &mut Vec<MetaFieldNode> {
        if let Some(element) = self.stack.last_mut() {
            &mut element.children
        } else {
            &mut self.roots
        }
    }

    fn add_node(&mut self, bytes: usize) -> Result<()> {
        self.nodes = self
            .nodes
            .checked_add(1)
            .ok_or_else(|| Error::InvalidFormat("text:meta node count overflow".to_string()))?;
        if self.nodes > MAX_TEXT_META_NODES {
            return Err(Error::InvalidFormat(format!(
                "text:meta exceeds {MAX_TEXT_META_NODES} content nodes"
            )));
        }
        self.bytes = self
            .bytes
            .checked_add(bytes)
            .ok_or_else(|| Error::InvalidFormat("text:meta byte count overflow".to_string()))?;
        if self.bytes > MAX_TEXT_META_BYTES {
            return Err(Error::InvalidFormat(format!(
                "text:meta exceeds {MAX_TEXT_META_BYTES} content bytes"
            )));
        }
        Ok(())
    }
}

fn host_for(
    namespace: Option<&str>,
    local: &str,
    attrs: &[MetaFieldAttribute],
    paragraph_index: &mut usize,
    heading_index: &mut usize,
    text_meta_index: &mut usize,
    other_occurrences: &mut HashMap<(String, String), usize>,
) -> Result<Option<RdfaHost>> {
    if namespace == Some(TEXTNS) {
        return match local {
            "p" => {
                let index = *paragraph_index;
                *paragraph_index = paragraph_index.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("paragraph occurrence count overflow".to_string())
                })?;
                Ok(Some(RdfaHost::Paragraph { index }))
            },
            "h" => {
                let index = *heading_index;
                *heading_index = heading_index.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("heading occurrence count overflow".to_string())
                })?;
                Ok(Some(RdfaHost::Heading { index }))
            },
            "bookmark-start" => Ok(Some(RdfaHost::BookmarkStart {
                name: attr_value(attrs, TEXTNS, "name"),
            })),
            "meta" => {
                let index = *text_meta_index;
                *text_meta_index = text_meta_index.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("text:meta occurrence count overflow".to_string())
                })?;
                Ok(Some(RdfaHost::TextMeta { index }))
            },
            _ => {
                let key = (TEXTNS.to_string(), local.to_string());
                let occurrence = other_occurrences.entry(key).or_insert(0);
                let current = *occurrence;
                *occurrence = occurrence.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("RDFa host occurrence count overflow".to_string())
                })?;
                Ok(Some(RdfaHost::Element {
                    namespace_uri: TEXTNS.to_string(),
                    local_name: local.to_string(),
                    occurrence: current,
                }))
            },
        };
    }
    let Some(namespace) = namespace else {
        return Ok(None);
    };
    let key = (namespace.to_string(), local.to_string());
    let occurrence = other_occurrences.entry(key).or_insert(0);
    let current = *occurrence;
    *occurrence = occurrence
        .checked_add(1)
        .ok_or_else(|| Error::InvalidFormat("RDFa host occurrence count overflow".to_string()))?;
    Ok(Some(RdfaHost::Element {
        namespace_uri: namespace.to_string(),
        local_name: local.to_string(),
        occurrence: current,
    }))
}

fn parse_attributes(
    reader: &NsReader<&[u8]>,
    source: &BytesStart<'_>,
) -> Result<Vec<MetaFieldAttribute>> {
    let mut attributes = Vec::new();
    for attribute in source.attributes() {
        let attribute = attribute.map_err(|error| {
            Error::InvalidFormat(format!("invalid in-content metadata attribute: {error}"))
        })?;
        let raw = attribute.key.as_ref();
        if raw == b"xmlns" || raw.starts_with(b"xmlns:") {
            continue;
        }
        if attributes.len() >= MAX_ATTRIBUTES {
            return Err(Error::InvalidFormat(format!(
                "in-content metadata element exceeds {MAX_ATTRIBUTES} attributes"
            )));
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace_uri = resolved_namespace(&namespace)?.unwrap_or_default();
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid metadata attribute value: {error}"))
            })?;
        attributes
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "ODT metadata attribute projection",
                source,
            })?;
        attributes.push(MetaFieldAttribute {
            namespace_uri,
            local_name: utf8(local.as_ref(), "metadata attribute name")?,
            value: value.into_owned(),
        });
    }
    Ok(attributes)
}

fn parse_meta_attributes(
    reader: &NsReader<&[u8]>,
    source: &BytesStart<'_>,
) -> Result<Vec<MetaFieldAttribute>> {
    let mut output = parse_attributes(reader, source)?;
    for attribute in &output {
        if attribute.namespace_uri.is_empty() {
            return Err(Error::InvalidFormat(
                "unqualified attributes inside text:meta are not supported".to_string(),
            ));
        }
    }
    if output.len() > 64 {
        return Err(Error::InvalidFormat(
            "text:meta child has too many attributes".to_string(),
        ));
    }
    output.shrink_to_fit();
    Ok(output)
}

fn rdfa_from_attributes(attrs: &[MetaFieldAttribute]) -> Result<RdfaAttributes> {
    let mut output = RdfaAttributes::default();
    for attr in attrs {
        if attr.namespace_uri == XHTMLNS {
            let slot = match attr.local_name.as_str() {
                "about" => &mut output.about,
                "property" => &mut output.property,
                "content" => &mut output.content,
                "datatype" => &mut output.datatype,
                _ => continue,
            };
            if slot.is_some() {
                return Err(Error::InvalidFormat(format!(
                    "duplicate xhtml:{} RDFa attribute",
                    attr.local_name
                )));
            }
            *slot = Some(attr.value.clone());
        }
    }
    output.validate()?;
    Ok(output)
}

fn text_meta_root_attributes(
    reader: &NsReader<&[u8]>,
    source: &BytesStart<'_>,
    attrs: &[MetaFieldAttribute],
) -> Result<(Option<String>, RdfaAttributes, Vec<MetadataAttribute>)> {
    let mut xml_id = None;
    for attr in attrs {
        if attr.namespace_uri == XMLNS && attr.local_name == "id" {
            if xml_id.is_some() {
                return Err(Error::InvalidFormat(
                    "duplicate text:meta xml:id".to_string(),
                ));
            }
            validate_xml_id(&attr.value)?;
            xml_id = Some(attr.value.clone());
        }
    }
    Ok((
        xml_id,
        rdfa_from_attributes(attrs)?,
        unknown_text_meta_attributes(reader, source)?,
    ))
}

fn unknown_text_meta_attributes(
    reader: &NsReader<&[u8]>,
    source: &BytesStart<'_>,
) -> Result<Vec<MetadataAttribute>> {
    let mut output = Vec::new();
    for attribute in source.attributes() {
        let attribute = attribute.map_err(|error| {
            Error::InvalidFormat(format!("invalid text:meta attribute: {error}"))
        })?;
        let raw = attribute.key.as_ref();
        if raw == b"xmlns" || raw.starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace_uri = resolved_namespace(&namespace)?.unwrap_or_default();
        let local_name = utf8(local.as_ref(), "text:meta attribute name")?;
        if (namespace_uri == XMLNS && local_name == "id") || namespace_uri == XHTMLNS {
            continue;
        }
        if output.len() >= 64 {
            return Err(Error::InvalidFormat(
                "text:meta root has too many unknown attributes".to_string(),
            ));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid text:meta attribute value: {error}"))
            })?;
        let prefix = raw
            .split(|byte| *byte == b':')
            .next()
            .filter(|candidate| *candidate != raw)
            .map(|candidate| String::from_utf8_lossy(candidate).into_owned());
        output.try_reserve(1).map_err(|source| Error::Allocation {
            resource: "ODT text:meta attribute projection",
            source,
        })?;
        output.push(MetadataAttribute {
            namespace_uri,
            local_name,
            value: value.into_owned(),
            prefix,
        });
    }
    if output.len() > 64 {
        return Err(Error::InvalidFormat(
            "text:meta root has too many unknown attributes".to_string(),
        ));
    }
    Ok(output)
}

fn attr_value(attrs: &[MetaFieldAttribute], namespace: &str, local: &str) -> Option<String> {
    attrs
        .iter()
        .find(|attr| attr.namespace_uri == namespace && attr.local_name == local)
        .map(|attr| attr.value.clone())
}

fn validate_xml_id(value: &str) -> Result<()> {
    let mut chars = value.chars();
    let Some(first) = chars.next() else {
        return Err(Error::InvalidFormat("xml:id must not be empty".to_string()));
    };
    let valid_start = first == '_' || first.is_alphabetic() || (first as u32) >= 0x80;
    let valid_rest = chars.all(|ch| {
        ch == '_' || ch == '-' || ch == '.' || ch.is_alphanumeric() || (ch as u32) >= 0x80
    });
    if !valid_start || !valid_rest || value.contains(':') {
        return Err(Error::InvalidFormat(format!(
            "invalid XML NCName '{value}'"
        )));
    }
    validate_xml_value("xml:id", value, MAX_TEXT_META_BYTES)
}

fn validate_xml_value(name: &str, value: &str, max_bytes: usize) -> Result<()> {
    if value.len() > max_bytes || !value.chars().all(is_xml_1_0_char) {
        return Err(Error::InvalidFormat(format!(
            "{name} exceeds its bounded XML lexical value"
        )));
    }
    Ok(())
}

const fn is_xml_1_0_char(value: char) -> bool {
    matches!(value, '\u{9}' | '\u{A}' | '\u{D}')
        || (value as u32 >= 0x20 && value as u32 <= 0xD7FF)
        || (value as u32 >= 0xE000 && value as u32 <= 0xFFFD)
        || (value as u32 >= 0x10000 && value as u32 <= 0x10FFFF)
}

fn push_rdfa_attributes(output: &mut String, attrs: &RdfaAttributes) {
    for (name, value) in [
        ("about", attrs.about.as_deref()),
        ("property", attrs.property.as_deref()),
        ("content", attrs.content.as_deref()),
        ("datatype", attrs.datatype.as_deref()),
    ] {
        if let Some(value) = value {
            output.push_str(" xhtml:");
            output.push_str(name);
            output.push_str("=\"");
            output.push_str(&escape_xml(value));
            output.push('"');
        }
    }
}

fn push_rdfa_namespace_declarations(
    output: &mut String,
    attrs: &RdfaAttributes,
    namespace_declarations: &[String],
) -> Result<()> {
    for prefix in rdfa_prefixes(attrs)? {
        if matches!(prefix.as_str(), "text" | "xhtml" | "xml") {
            continue;
        }
        if namespace_binding(namespace_declarations, &prefix).is_some() {
            continue;
        }
        let Some(namespace) = canonical_rdfa_namespace(&prefix) else {
            continue;
        };
        output.push_str(" xmlns:");
        output.push_str(&prefix);
        output.push_str("=\"");
        output.push_str(namespace);
        output.push('"');
    }
    Ok(())
}

fn append_required_rdfa_namespaces(
    output: &mut String,
    source: &str,
    attrs: &RdfaAttributes,
    namespace_declarations: &[String],
) -> Result<()> {
    for prefix in rdfa_prefixes(attrs)? {
        if source_namespace_binding(source, &prefix)?.is_some()
            || namespace_binding(namespace_declarations, &prefix).is_some()
        {
            continue;
        }
        let Some(namespace) = canonical_rdfa_namespace(&prefix) else {
            return Err(Error::InvalidFormat(format!(
                "RDFa CURIE prefix '{prefix}' has no in-scope namespace binding"
            )));
        };
        output.push_str(" xmlns:");
        output.push_str(&prefix);
        output.push_str("=\"");
        output.push_str(namespace);
        output.push('"');
    }
    Ok(())
}

fn ensure_rdfa_namespace_context(
    attrs: &RdfaAttributes,
    namespace_declarations: &[String],
) -> Result<()> {
    for prefix in rdfa_prefixes(attrs)? {
        if namespace_binding(namespace_declarations, &prefix).is_none()
            && canonical_rdfa_namespace(&prefix).is_none()
        {
            return Err(Error::InvalidFormat(format!(
                "RDFa CURIE prefix '{prefix}' has no in-scope namespace binding"
            )));
        }
    }
    Ok(())
}

fn namespace_binding(namespace_declarations: &[String], prefix: &str) -> Option<String> {
    let wanted = format!("xmlns:{prefix}");
    namespace_declarations
        .iter()
        .filter_map(|declaration| declaration_parts(declaration))
        .find(|(name, _)| *name == wanted)
        .map(|(_, value)| value.to_owned())
}

fn canonical_rdfa_namespace(prefix: &str) -> Option<&'static str> {
    match prefix {
        "dc" => Some(DCNS),
        "xsd" => Some(XSDNS),
        "xhtml" => Some(XHTMLNS),
        "text" => Some(TEXTNS),
        _ => None,
    }
}

fn rdfa_prefixes(attrs: &RdfaAttributes) -> Result<Vec<String>> {
    let mut output = Vec::new();
    if let Some(value) = &attrs.about {
        add_rdfa_prefix(&mut output, value, true)?;
    }
    if let Some(value) = &attrs.property {
        for token in value.split_whitespace() {
            add_rdfa_prefix(&mut output, token, false)?;
        }
    }
    if let Some(value) = &attrs.datatype {
        add_rdfa_prefix(&mut output, value, false)?;
    }
    Ok(output)
}

fn add_rdfa_prefix(output: &mut Vec<String>, value: &str, safe_curie_only: bool) -> Result<()> {
    let value = if safe_curie_only {
        let Some(value) = value
            .strip_prefix('[')
            .and_then(|value| value.strip_suffix(']'))
        else {
            return Ok(());
        };
        value
    } else {
        value
            .strip_prefix('[')
            .and_then(|value| value.strip_suffix(']'))
            .unwrap_or(value)
    };
    let Some((prefix, _)) = value.split_once(':') else {
        return Ok(());
    };
    if prefix == "_" {
        return Ok(());
    }
    if !is_valid_metadata_prefix(prefix) {
        return Err(Error::InvalidFormat(format!(
            "invalid RDFa CURIE prefix '{prefix}'"
        )));
    }
    if !output.iter().any(|current| current == prefix) {
        output.push(prefix.to_owned());
    }
    Ok(())
}

fn validate_metadata_attributes(attrs: &[MetadataAttribute]) -> Result<()> {
    if attrs.len() > 64 {
        return Err(Error::InvalidFormat(
            "text:meta root has too many unknown attributes".to_string(),
        ));
    }
    for attr in attrs {
        if attr.local_name.contains(':') || attr.local_name == "xmlns" {
            return Err(Error::InvalidFormat(
                "text:meta attribute local name is not serializable".to_string(),
            ));
        }
        validate_xml_name(&attr.local_name, "text:meta attribute name")?;
        validate_xml_value(
            "text:meta attribute namespace",
            &attr.namespace_uri,
            MAX_TEXT_META_BYTES,
        )?;
        validate_xml_value(
            "text:meta attribute name",
            &attr.local_name,
            MAX_TEXT_META_BYTES,
        )?;
        validate_xml_value(
            "text:meta attribute value",
            &attr.value,
            MAX_TEXT_META_BYTES,
        )?;
        if let Some(prefix) = &attr.prefix {
            if !is_valid_metadata_prefix(prefix) {
                return Err(Error::InvalidFormat(
                    "text:meta attribute prefix is not serializable".to_string(),
                ));
            }
        }
    }
    Ok(())
}

fn push_metadata_attributes(output: &mut String, attrs: &[MetadataAttribute]) {
    let mut dynamic = vec![
        ("text".to_string(), TEXTNS.to_string()),
        ("xhtml".to_string(), XHTMLNS.to_string()),
        ("xml".to_string(), XMLNS.to_string()),
    ];
    let mut generated = 0usize;
    for attr in attrs {
        let prefix = if attr.namespace_uri.is_empty() {
            String::new()
        } else if let Some(prefix) = canonical_metadata_prefix(&attr.namespace_uri) {
            prefix.to_string()
        } else if let Some(prefix) = attr.prefix.as_deref().filter(|prefix| {
            is_valid_metadata_prefix(prefix)
                && dynamic.iter().all(|(current, namespace)| {
                    current != *prefix || namespace == &attr.namespace_uri
                })
        }) {
            let prefix = prefix.to_string();
            if !dynamic
                .iter()
                .any(|(current, namespace)| current == &prefix && namespace == &attr.namespace_uri)
            {
                output.push_str(" xmlns:");
                output.push_str(&prefix);
                output.push_str("=\"");
                output.push_str(&escape_xml(&attr.namespace_uri));
                output.push('"');
                dynamic.push((prefix.clone(), attr.namespace_uri.clone()));
            }
            prefix
        } else if let Some((prefix, _)) = dynamic
            .iter()
            .find(|(_, namespace)| namespace == &attr.namespace_uri)
        {
            prefix.clone()
        } else {
            let prefix = loop {
                let candidate = format!("meta{generated}");
                generated = generated.saturating_add(1);
                if dynamic.iter().all(|(current, _)| current != &candidate) {
                    break candidate;
                }
            };
            output.push_str(" xmlns:");
            output.push_str(&prefix);
            output.push_str("=\"");
            output.push_str(&escape_xml(&attr.namespace_uri));
            output.push('"');
            dynamic.push((prefix.clone(), attr.namespace_uri.clone()));
            prefix
        };
        output.push(' ');
        if !prefix.is_empty() {
            output.push_str(&prefix);
            output.push(':');
        }
        output.push_str(&attr.local_name);
        output.push_str("=\"");
        output.push_str(&escape_xml(&attr.value));
        output.push('"');
    }
}

fn canonical_metadata_prefix(namespace: &str) -> Option<&'static str> {
    match namespace {
        TEXTNS => Some("text"),
        XHTMLNS => Some("xhtml"),
        XMLNS => Some("xml"),
        _ => None,
    }
}

fn validate_xml_name(value: &str, name: &str) -> Result<()> {
    let mut chars = value.chars();
    let Some(first) = chars.next() else {
        return Err(Error::InvalidFormat(format!("{name} must not be empty")));
    };
    if !(first == '_' || first.is_alphabetic())
        || !chars.all(|character| {
            character == '_' || character == '-' || character == '.' || character.is_alphanumeric()
        })
    {
        return Err(Error::InvalidFormat(format!("invalid XML name '{value}'")));
    }
    validate_xml_value(name, value, MAX_TEXT_META_BYTES)
}

fn is_valid_metadata_prefix(prefix: &str) -> bool {
    let Some(first) = prefix.as_bytes().first().copied() else {
        return false;
    };
    (first.is_ascii_alphabetic() || first == b'_')
        && prefix != "xml"
        && prefix != "xmlns"
        && prefix
            .bytes()
            .all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'.' | b'-'))
}

fn resolved_namespace(namespace: &ResolveResult<'_>) -> Result<Option<String>> {
    match namespace {
        ResolveResult::Bound(value) => Ok(Some(utf8(value.as_ref(), "namespace URI")?)),
        ResolveResult::Unbound => Ok(None),
        ResolveResult::Unknown(prefix) => Err(Error::InvalidFormat(format!(
            "unbound namespace prefix '{}'",
            String::from_utf8_lossy(prefix)
        ))),
    }
}

fn utf8(value: &[u8], description: &str) -> Result<String> {
    std::str::from_utf8(value)
        .map(str::to_owned)
        .map_err(|_| Error::InvalidFormat(format!("invalid UTF-8 {description}")))
}

fn push_bounded<T>(items: &mut Vec<T>, value: T, resource: &'static str) -> Result<()> {
    items
        .try_reserve(1)
        .map_err(|source| Error::Allocation { resource, source })?;
    items.push(value);
    Ok(())
}

// -------------------------------------------------------------------------
// Lossless content.xml editing helpers

#[derive(Debug, Clone)]
struct Span {
    start: usize,
    start_end: usize,
    end_start: usize,
    end: usize,
    namespace: String,
    local: String,
    rdfa_names: Vec<String>,
    rdfa: RdfaAttributes,
    name: Option<String>,
    namespace_declarations: Vec<String>,
}

/// Set RDFa on a paragraph selected in document order.
pub(crate) fn set_paragraph_rdfa(
    xml: &str,
    position: Position,
    value: &RdfaAttributes,
) -> Result<String> {
    value.validate()?;
    let spans = scan_spans(xml, |namespace, local| namespace == TEXTNS && local == "p")?;
    let span = spans.get(position.get()).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "paragraph position {} is out of range",
            position.get()
        ))
    })?;
    if span.rdfa == *value {
        return Ok(xml.to_owned());
    }
    rewrite_rdfa(xml, span, value)
}

/// Set RDFa on a uniquely named bookmark start.
pub(crate) fn set_bookmark_rdfa(xml: &str, name: &str, value: &RdfaAttributes) -> Result<String> {
    value.validate()?;
    let spans = scan_spans(xml, |namespace, local| {
        namespace == TEXTNS && local == "bookmark-start"
    })?;
    let mut matches = spans
        .iter()
        .filter(|span| span.name.as_deref() == Some(name));
    let span = matches
        .next()
        .ok_or_else(|| Error::InvalidFormat(format!("bookmark start '{name}' was not found")))?;
    if matches.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "bookmark start '{name}' is not unique"
        )));
    }
    if span.rdfa == *value {
        return Ok(xml.to_owned());
    }
    rewrite_rdfa(xml, span, value)
}

/// Set RDFa on one inline `text:meta` occurrence.
pub(crate) fn set_text_meta_rdfa(
    xml: &str,
    position: Position,
    value: &RdfaAttributes,
) -> Result<String> {
    value.validate()?;
    let spans = scan_spans(xml, |namespace, local| {
        namespace == TEXTNS && local == "meta"
    })?;
    let span = spans.get(position.get()).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "text:meta position {} is out of range",
            position.get()
        ))
    })?;
    if span.rdfa == *value {
        return Ok(xml.to_owned());
    }
    rewrite_rdfa(xml, span, value)
}

/// Insert an inline `text:meta` before the selected paragraph's end tag.
pub(crate) fn insert_text_meta(xml: &str, paragraph: Position, value: &TextMeta) -> Result<String> {
    let spans = scan_spans(xml, |namespace, local| namespace == TEXTNS && local == "p")?;
    let span = spans.get(paragraph.get()).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "paragraph position {} is out of range",
            paragraph.get()
        ))
    })?;
    value.validate()?;
    ensure_rdfa_namespace_context(&value.rdfa, &span.namespace_declarations)?;
    let fragment = inject_namespace_declarations(
        &value.to_xml_with_context(&span.namespace_declarations)?,
        &span.namespace_declarations,
    )?;
    if span.end_start == span.start_end {
        let source = xml
            .get(span.start..span.start_end)
            .ok_or_else(|| Error::InvalidFormat("invalid empty paragraph span".to_string()))?;
        if !source.trim_end().ends_with("/>") {
            return Err(Error::InvalidFormat(
                "invalid empty paragraph element".to_string(),
            ));
        }
        let close = source.len().checked_sub(2).ok_or_else(|| {
            Error::InvalidFormat("invalid empty paragraph closing delimiter".to_string())
        })?;
        let replacement_len = source
            .len()
            .checked_add(fragment.len())
            .and_then(|length| length.checked_add(16))
            .ok_or_else(|| Error::InvalidFormat("text:meta insertion size overflow".to_string()))?;
        if replacement_len > MAX_XML_BYTES {
            return Err(Error::InvalidFormat(format!(
                "text:meta insertion exceeds the {MAX_XML_BYTES} edit limit"
            )));
        }
        let mut replacement = String::new();
        replacement
            .try_reserve_exact(replacement_len)
            .map_err(|source| Error::Allocation {
                resource: "ODT text:meta insertion",
                source,
            })?;
        replacement.push_str(&source[..close]);
        replacement.push('>');
        replacement.push_str(&fragment);
        replacement.push_str("</");
        let qname_end = source[1..]
            .find(|ch: char| ch.is_ascii_whitespace() || ch == '/')
            .map_or(close - 1, |index| index + 1);
        replacement.push_str(&source[1..qname_end]);
        replacement.push('>');
        return splice_replace(xml, span.start, span.end, &replacement);
    }
    splice_insert(xml, span.end_start, &fragment)
}

/// Replace one inline `text:meta` element.
pub(crate) fn replace_text_meta(xml: &str, position: Position, value: &TextMeta) -> Result<String> {
    let spans = scan_spans(xml, |namespace, local| {
        namespace == TEXTNS && local == "meta"
    })?;
    let span = spans.get(position.get()).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "text:meta position {} is out of range",
            position.get()
        ))
    })?;
    value.validate()?;
    ensure_rdfa_namespace_context(&value.rdfa, &span.namespace_declarations)?;
    let existing = parse_text_meta_span(xml, span)?;
    if text_meta_semantically_equal(&existing, value) {
        return Ok(xml.to_owned());
    }
    let fragment = inject_namespace_declarations(
        &value.to_xml_with_context(&span.namespace_declarations)?,
        &span.namespace_declarations,
    )?;
    splice_replace(xml, span.start, span.end, &fragment)
}

fn parse_text_meta_span(xml: &str, span: &Span) -> Result<TextMeta> {
    let raw = xml
        .get(span.start..span.end)
        .ok_or_else(|| Error::InvalidFormat("invalid text:meta span".to_string()))?;
    let owned = inject_namespace_declarations(raw, &span.namespace_declarations)?;
    let parsed = parse_part(&owned, MetadataPart::Content)?;
    parsed
        .text_meta
        .into_iter()
        .next()
        .ok_or_else(|| Error::InvalidFormat("text:meta span did not parse as metadata".to_string()))
}

fn text_meta_semantically_equal(left: &TextMeta, right: &TextMeta) -> bool {
    left.xml_id == right.xml_id
        && left.rdfa == right.rdfa
        && left.attributes == right.attributes
        && left.content == right.content
}

/// Remove one inline `text:meta` element.
pub(crate) fn remove_text_meta(xml: &str, position: Position) -> Result<String> {
    let spans = scan_spans(xml, |namespace, local| {
        namespace == TEXTNS && local == "meta"
    })?;
    let span = spans.get(position.get()).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "text:meta position {} is out of range",
            position.get()
        ))
    })?;
    splice_replace(xml, span.start, span.end, "")
}

fn rewrite_rdfa(xml: &str, span: &Span, value: &RdfaAttributes) -> Result<String> {
    value.validate()?;
    let source = xml
        .get(span.start..span.start_end)
        .ok_or_else(|| Error::InvalidFormat("invalid RDFa start-tag span".to_string()))?;
    let mut remove = HashSet::new();
    remove.extend(span.rdfa_names.iter().cloned());
    let add = !value.is_empty();
    let rewritten = rewrite_start_tag(source, &remove, add, value, &span.namespace_declarations)?;
    splice_replace(xml, span.start, span.start_end, &rewritten)
}

fn rewrite_start_tag(
    source: &str,
    remove_names: &HashSet<String>,
    add_namespace: bool,
    attrs: &RdfaAttributes,
    namespace_declarations: &[String],
) -> Result<String> {
    let bytes = source.as_bytes();
    let close = if source.ends_with("/>") {
        source.len() - 2
    } else {
        source.len() - 1
    };
    let output_len = rewritten_start_tag_len(
        source,
        remove_names,
        add_namespace,
        attrs,
        namespace_declarations,
    )?;
    bounded_output_len(output_len, "RDFa start-tag rewrite")?;
    let mut output = String::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| Error::Allocation {
            resource: "ODT RDFa start-tag rewrite",
            source,
        })?;
    let mut cursor;
    let mut name_end = 1usize;
    while name_end < close && !bytes[name_end].is_ascii_whitespace() {
        name_end += 1;
    }
    output.push_str(&source[..name_end]);
    cursor = name_end;
    while cursor < close {
        let token_start = cursor;
        while cursor < close && bytes[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= close {
            output.push_str(&source[token_start..close]);
            break;
        }
        let attr_start = cursor;
        while cursor < close
            && !bytes[cursor].is_ascii_whitespace()
            && bytes[cursor] != b'='
            && bytes[cursor] != b'/'
        {
            cursor += 1;
        }
        let attr_name = &source[attr_start..cursor];
        while cursor < close && bytes[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor < close && bytes[cursor] == b'=' {
            cursor += 1;
            while cursor < close && bytes[cursor].is_ascii_whitespace() {
                cursor += 1;
            }
            if cursor < close && (bytes[cursor] == b'"' || bytes[cursor] == b'\'') {
                let quote = bytes[cursor];
                cursor += 1;
                while cursor < close && bytes[cursor] != quote {
                    cursor += 1;
                }
                if cursor < close {
                    cursor += 1;
                }
            }
        }
        if !remove_names.contains(attr_name) {
            output.push_str(&source[token_start..cursor]);
        }
    }
    if add_namespace {
        match source_namespace_binding(source, "xhtml")? {
            Some(namespace) if namespace != XHTMLNS => {
                return Err(Error::InvalidFormat(
                    "RDFa edit cannot shadow the xhtml namespace prefix".to_string(),
                ));
            },
            Some(_) => {},
            None => {
                output.push_str(" xmlns:xhtml=\"");
                output.push_str(XHTMLNS);
                output.push('"');
            },
        }
        append_required_rdfa_namespaces(&mut output, source, attrs, namespace_declarations)?;
    }
    push_rdfa_attributes(&mut output, attrs);
    output.push_str(&source[close..]);
    Ok(output)
}

fn rewritten_start_tag_len(
    source: &str,
    remove_names: &HashSet<String>,
    add_namespace: bool,
    attrs: &RdfaAttributes,
    namespace_declarations: &[String],
) -> Result<usize> {
    let bytes = source.as_bytes();
    let close = if source.ends_with("/>") {
        source.len() - 2
    } else {
        source.len() - 1
    };
    let mut length = 0usize;
    let mut cursor;
    let mut name_end = 1usize;
    while name_end < close && !bytes[name_end].is_ascii_whitespace() {
        name_end += 1;
    }
    add_size(&mut length, name_end)?;
    cursor = name_end;
    while cursor < close {
        let token_start = cursor;
        while cursor < close && bytes[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= close {
            add_size(&mut length, close - token_start)?;
            break;
        }
        let attr_start = cursor;
        while cursor < close
            && !bytes[cursor].is_ascii_whitespace()
            && bytes[cursor] != b'='
            && bytes[cursor] != b'/'
        {
            cursor += 1;
        }
        let attr_name = &source[attr_start..cursor];
        while cursor < close && bytes[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor < close && bytes[cursor] == b'=' {
            cursor += 1;
            while cursor < close && bytes[cursor].is_ascii_whitespace() {
                cursor += 1;
            }
            if cursor < close && (bytes[cursor] == b'"' || bytes[cursor] == b'\'') {
                let quote = bytes[cursor];
                cursor += 1;
                while cursor < close && bytes[cursor] != quote {
                    cursor += 1;
                }
                if cursor < close {
                    cursor += 1;
                }
            }
        }
        if !remove_names.contains(attr_name) {
            add_size(&mut length, cursor - token_start)?;
        }
    }
    if add_namespace {
        match source_namespace_binding(source, "xhtml")? {
            Some(namespace) if namespace != XHTMLNS => {
                return Err(Error::InvalidFormat(
                    "RDFa edit cannot shadow the xhtml namespace prefix".to_string(),
                ));
            },
            Some(_) => {},
            None => {
                add_size(&mut length, " xmlns:xhtml=\"\"".len() + XHTMLNS.len())?;
            },
        }
        for prefix in rdfa_prefixes(attrs)? {
            if source_namespace_binding(source, &prefix)?.is_some()
                || namespace_binding(namespace_declarations, &prefix).is_some()
            {
                continue;
            }
            let Some(namespace) = canonical_rdfa_namespace(&prefix) else {
                return Err(Error::InvalidFormat(format!(
                    "RDFa CURIE prefix '{prefix}' has no in-scope namespace binding"
                )));
            };
            add_size(
                &mut length,
                " xmlns:=\"\"".len() + prefix.len() + namespace.len(),
            )?;
        }
        for (name, value) in [
            ("about", attrs.about.as_deref()),
            ("property", attrs.property.as_deref()),
            ("content", attrs.content.as_deref()),
            ("datatype", attrs.datatype.as_deref()),
        ] {
            if let Some(value) = value {
                let value_len = escaped_xml_len(value, true)?;
                let attribute_len = " xhtml:=\"\""
                    .len()
                    .checked_add(name.len())
                    .and_then(|size| size.checked_add(value_len))
                    .ok_or_else(|| {
                        Error::InvalidFormat("RDFa start-tag size overflow".to_string())
                    })?;
                add_size(&mut length, attribute_len)?;
            }
        }
    }
    add_size(&mut length, source.len() - close)?;
    Ok(length)
}

fn source_namespace_binding(source: &str, prefix: &str) -> Result<Option<String>> {
    let wanted = format!("xmlns:{prefix}");
    let mut reader = NsReader::from_str(source);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    loop {
        let event = reader.read_event_into(&mut buffer).map_err(|error| {
            Error::InvalidFormat(format!("invalid RDFa target start tag: {error}"))
        })?;
        match event {
            Event::Start(source) | Event::Empty(source) => {
                for attribute in source.attributes() {
                    let attribute = attribute.map_err(|error| {
                        Error::InvalidFormat(format!("invalid RDFa namespace declaration: {error}"))
                    })?;
                    if attribute.key.as_ref() == wanted.as_bytes() {
                        let value = attribute
                            .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
                            .map_err(|error| {
                                Error::InvalidFormat(format!(
                                    "invalid RDFa namespace declaration value: {error}"
                                ))
                            })?;
                        return Ok(Some(value.into_owned()));
                    }
                }
                return Ok(None);
            },
            Event::Eof => {
                return Err(Error::InvalidFormat(
                    "missing RDFa target start tag".to_string(),
                ));
            },
            _ => buffer.clear(),
        }
    }
}

fn metadata_namespace_declarations(source: &BytesStart<'_>) -> Result<Vec<String>> {
    let mut output = Vec::new();
    for attribute in source.attributes() {
        let attribute = attribute.map_err(|error| {
            Error::InvalidFormat(format!("invalid metadata namespace declaration: {error}"))
        })?;
        let raw = attribute.key.as_ref();
        if raw == b"xmlns" || raw.starts_with(b"xmlns:") {
            let key = std::str::from_utf8(raw).map_err(|_| {
                Error::InvalidFormat("invalid metadata namespace declaration name".to_string())
            })?;
            let value = std::str::from_utf8(attribute.value.as_ref()).map_err(|_| {
                Error::InvalidFormat("invalid metadata namespace declaration value".to_string())
            })?;
            output.push(format!(" {key}=\"{value}\""));
        }
    }
    Ok(output)
}

fn apply_metadata_namespace_declarations(
    scope: &mut Vec<(String, String)>,
    declarations: &[String],
) -> Result<Vec<(String, Option<String>)>> {
    let mut changes = Vec::new();
    for declaration in declarations {
        let (name, value) = declaration_parts(declaration).ok_or_else(|| {
            Error::InvalidFormat("invalid metadata namespace declaration span".to_string())
        })?;
        let previous = scope
            .iter()
            .find(|(current, _)| current == name)
            .map(|(_, value)| value.clone());
        changes.push((name.to_owned(), previous));
        replace_metadata_namespace(scope, name, value);
    }
    Ok(changes)
}

fn restore_metadata_namespace_declarations(
    scope: &mut Vec<(String, String)>,
    changes: Vec<(String, Option<String>)>,
) {
    for (name, previous) in changes.into_iter().rev() {
        match previous {
            Some(value) => replace_metadata_namespace(scope, &name, &value),
            None => scope.retain(|(current, _)| current != &name),
        }
    }
}

fn replace_metadata_namespace(scope: &mut Vec<(String, String)>, name: &str, value: &str) {
    if let Some((_, current)) = scope.iter_mut().find(|(current, _)| current == name) {
        *current = value.to_owned();
    } else {
        scope.push((name.to_owned(), value.to_owned()));
    }
}

fn metadata_namespace_scope_to_raw(scope: &[(String, String)]) -> Vec<String> {
    scope
        .iter()
        .map(|(name, value)| format!(" {name}=\"{value}\""))
        .collect()
}

fn declaration_parts(declaration: &str) -> Option<(&str, &str)> {
    let declaration = declaration.trim();
    let (name, value) = declaration.split_once('=')?;
    let value = value.trim().strip_prefix('"')?.strip_suffix('"')?;
    Some((name.trim(), value))
}

fn inject_namespace_declarations(raw: &str, declarations: &[String]) -> Result<String> {
    let (open_end, empty, present) = metadata_first_tag_span(raw)?;
    let mut insert_len = 0usize;
    for declaration in declarations {
        let Some((name, _)) = declaration_parts(declaration) else {
            continue;
        };
        if !present.contains(name) {
            insert_len = insert_len.checked_add(declaration.len()).ok_or_else(|| {
                Error::InvalidFormat("metadata namespace insertion size overflow".to_string())
            })?;
        }
    }
    if insert_len == 0 {
        return Ok(raw.to_owned());
    }
    let offset = if empty {
        open_end.checked_sub(2).ok_or_else(|| {
            Error::InvalidFormat("invalid metadata empty element span".to_string())
        })?
    } else {
        open_end
            .checked_sub(1)
            .ok_or_else(|| Error::InvalidFormat("invalid metadata element span".to_string()))?
    };
    let output_len = raw.len().checked_add(insert_len).ok_or_else(|| {
        Error::InvalidFormat("metadata namespace insertion size overflow".to_string())
    })?;
    bounded_output_len(output_len, "metadata namespace insertion")?;
    let mut output = String::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| Error::Allocation {
            resource: "ODT metadata namespace insertion",
            source,
        })?;
    output.push_str(&raw[..offset]);
    for declaration in declarations {
        let Some((name, _)) = declaration_parts(declaration) else {
            continue;
        };
        if !present.contains(name) {
            output.push_str(declaration);
        }
    }
    output.push_str(&raw[offset..]);
    Ok(output)
}

fn metadata_first_tag_span(raw: &str) -> Result<(usize, bool, HashSet<String>)> {
    let mut reader = NsReader::from_str(raw);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    loop {
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid metadata XML: {error}")))?;
        match event {
            Event::Start(source) => {
                return Ok((
                    reader.buffer_position() as usize,
                    false,
                    metadata_present_namespace_names(&source)?,
                ));
            },
            Event::Empty(source) => {
                return Ok((
                    reader.buffer_position() as usize,
                    true,
                    metadata_present_namespace_names(&source)?,
                ));
            },
            Event::DocType(_) => {
                return Err(Error::InvalidFormat(
                    "DOCTYPE is not permitted in metadata XML".to_string(),
                ));
            },
            Event::Eof => {
                return Err(Error::InvalidFormat(
                    "missing metadata element root".to_string(),
                ));
            },
            _ => buffer.clear(),
        }
    }
}

fn metadata_present_namespace_names(source: &BytesStart<'_>) -> Result<HashSet<String>> {
    let mut output = HashSet::new();
    for attribute in source.attributes() {
        let attribute = attribute.map_err(|error| {
            Error::InvalidFormat(format!("invalid metadata namespace declaration: {error}"))
        })?;
        let raw = attribute.key.as_ref();
        if raw == b"xmlns" || raw.starts_with(b"xmlns:") {
            output.insert(
                std::str::from_utf8(raw)
                    .map_err(|_| {
                        Error::InvalidFormat(
                            "invalid metadata namespace declaration name".to_string(),
                        )
                    })?
                    .to_owned(),
            );
        }
    }
    Ok(output)
}

fn scan_spans(xml: &str, wanted: impl Fn(&str, &str) -> bool) -> Result<Vec<Span>> {
    if xml.len() > MAX_XML_BYTES {
        return Err(Error::InvalidFormat(format!(
            "ODF XML exceeds the {MAX_XML_BYTES} edit limit"
        )));
    }
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut stack: Vec<(Span, Vec<(String, Option<String>)>)> = Vec::new();
    let mut namespace_scope = Vec::new();
    let mut output = Vec::new();
    loop {
        let event_position = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid ODF XML: {error}")))?;
        let namespace = resolved_namespace(&namespace)?.unwrap_or_default();
        let event_end = reader.buffer_position() as usize;
        match event {
            Event::Start(ref source) => {
                let local = utf8(source.local_name().as_ref(), "edit target name")?;
                let start = event_position;
                let declarations = metadata_namespace_declarations(source)?;
                let changes =
                    apply_metadata_namespace_declarations(&mut namespace_scope, &declarations)?;
                let attributes = parse_attributes(&reader, source)?;
                let rdfa_names = target_attribute_names(&reader, source)?;
                let name = target_attribute_value(&reader, source, TEXTNS, "name")?;
                let span = Span {
                    start,
                    start_end: event_end,
                    end_start: 0,
                    end: 0,
                    namespace: namespace.clone(),
                    local: local.clone(),
                    rdfa_names,
                    rdfa: rdfa_from_attributes(&attributes)?,
                    name,
                    namespace_declarations: metadata_namespace_scope_to_raw(&namespace_scope),
                };
                stack.push((span, changes));
                if !wanted(&namespace, &local) {
                    // Keep a frame for matching end tags; it is discarded on close.
                }
            },
            Event::Empty(ref source) => {
                let local = utf8(source.local_name().as_ref(), "edit target name")?;
                let declarations = metadata_namespace_declarations(source)?;
                let changes =
                    apply_metadata_namespace_declarations(&mut namespace_scope, &declarations)?;
                if wanted(&namespace, &local) {
                    let start = event_position;
                    let attributes = parse_attributes(&reader, source)?;
                    let rdfa_names = target_attribute_names(&reader, source)?;
                    let name = target_attribute_value(&reader, source, TEXTNS, "name")?;
                    output.push(Span {
                        start,
                        start_end: event_end,
                        end_start: event_end,
                        end: event_end,
                        namespace,
                        local,
                        rdfa_names,
                        rdfa: rdfa_from_attributes(&attributes)?,
                        name,
                        namespace_declarations: metadata_namespace_scope_to_raw(&namespace_scope),
                    });
                }
                restore_metadata_namespace_declarations(&mut namespace_scope, changes);
            },
            Event::End(_) => {
                let end = event_end;
                let end_start = event_position;
                let Some((mut span, changes)) = stack.pop() else {
                    return Err(Error::InvalidFormat(
                        "metadata XML stack underflow while locating edit spans".to_string(),
                    ));
                };
                span.end_start = end_start;
                span.end = end;
                if wanted(&span.namespace, &span.local) {
                    output.push(span);
                }
                restore_metadata_namespace_declarations(&mut namespace_scope, changes);
            },
            Event::DocType(_) => {
                return Err(Error::InvalidFormat(
                    "DOCTYPE is not permitted in ODF XML edits".to_string(),
                ));
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if !stack.is_empty() || !namespace_scope.is_empty() {
        return Err(Error::InvalidFormat(
            "incomplete metadata XML while locating edit spans".to_string(),
        ));
    }
    output.sort_by_key(|span| span.start);
    Ok(output)
}

fn target_attribute_names(
    reader: &NsReader<&[u8]>,
    source: &BytesStart<'_>,
) -> Result<Vec<String>> {
    let mut rdfa = Vec::new();
    for attr in source.attributes() {
        let attr =
            attr.map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
        let raw = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
        let (namespace, local) = reader.resolver().resolve_attribute(attr.key);
        let namespace = resolved_namespace(&namespace)?;
        if namespace.as_deref() == Some(XHTMLNS)
            && matches!(
                local.as_ref(),
                b"about" | b"property" | b"content" | b"datatype"
            )
        {
            rdfa.push(raw.clone());
        }
    }
    Ok(rdfa)
}

fn target_attribute_value(
    reader: &NsReader<&[u8]>,
    source: &BytesStart<'_>,
    namespace: &str,
    local: &str,
) -> Result<Option<String>> {
    for attr in source.attributes() {
        let attr =
            attr.map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
        let (resolved, local_name) = reader.resolver().resolve_attribute(attr.key);
        let resolved = match resolved {
            ResolveResult::Bound(value) => value.as_ref().to_vec(),
            _ => continue,
        };
        if resolved == namespace.as_bytes() && local_name.as_ref() == local.as_bytes() {
            return attr
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
                .map(|value| Some(value.into_owned()))
                .map_err(|error| {
                    Error::InvalidFormat(format!("invalid XML attribute value: {error}"))
                });
        }
    }
    Ok(None)
}

fn splice_insert(xml: &str, offset: usize, value: &str) -> Result<String> {
    if offset > xml.len() {
        return Err(Error::InvalidFormat(
            "invalid XML insertion offset".to_string(),
        ));
    }
    let output_len = xml
        .len()
        .checked_add(value.len())
        .ok_or_else(|| Error::InvalidFormat("XML insertion size overflow".to_string()))?;
    bounded_output_len(output_len, "XML insertion")?;
    let mut output = String::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| Error::Allocation {
            resource: "ODT XML insertion",
            source,
        })?;
    output.push_str(&xml[..offset]);
    output.push_str(value);
    output.push_str(&xml[offset..]);
    Ok(output)
}

fn splice_replace(xml: &str, start: usize, end: usize, value: &str) -> Result<String> {
    if start > end || end > xml.len() {
        return Err(Error::InvalidFormat(
            "invalid XML replacement span".to_string(),
        ));
    }
    let removed = end
        .checked_sub(start)
        .ok_or_else(|| Error::InvalidFormat("XML replacement span underflow".to_string()))?;
    let output_len = xml
        .len()
        .checked_sub(removed)
        .and_then(|length| length.checked_add(value.len()))
        .ok_or_else(|| Error::InvalidFormat("XML replacement size overflow".to_string()))?;
    bounded_output_len(output_len, "XML replacement")?;
    let mut output = String::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| Error::Allocation {
            resource: "ODT XML replacement",
            source,
        })?;
    output.push_str(&xml[..start]);
    output.push_str(value);
    output.push_str(&xml[end..]);
    Ok(output)
}

fn bounded_output_len(length: usize, operation: &str) -> Result<()> {
    if length > MAX_XML_BYTES {
        return Err(Error::InvalidFormat(format!(
            "{operation} exceeds the {MAX_XML_BYTES} edit limit"
        )));
    }
    Ok(())
}

fn metadata_output_len(value: &TextMeta, namespace_declarations: &[String]) -> Result<usize> {
    let mut length =
        "<text:meta xmlns:text=\"\" xmlns:xhtml=\"\"".len() + TEXTNS.len() + XHTMLNS.len();
    if let Some(id) = &value.xml_id {
        add_size(&mut length, " xml:id=\"\"".len())?;
        add_size(&mut length, escaped_xml_len(id, true)?)?;
    }
    for prefix in rdfa_prefixes(&value.rdfa)? {
        if matches!(prefix.as_str(), "text" | "xhtml" | "xml")
            || namespace_binding(namespace_declarations, &prefix).is_some()
        {
            continue;
        }
        if let Some(namespace) = canonical_rdfa_namespace(&prefix) {
            add_size(
                &mut length,
                " xmlns:=\"\"".len() + prefix.len() + namespace.len(),
            )?;
        }
    }
    for (name, item) in [
        ("about", value.rdfa.about.as_deref()),
        ("property", value.rdfa.property.as_deref()),
        ("content", value.rdfa.content.as_deref()),
        ("datatype", value.rdfa.datatype.as_deref()),
    ] {
        if let Some(item) = item {
            add_size(&mut length, " xhtml:=\"\"".len() + name.len())?;
            add_size(&mut length, escaped_xml_len(item, true)?)?;
        }
    }
    for attribute in &value.attributes {
        // Metadata attributes can introduce a declaration for their retained
        // namespace.  The generated prefix itself is at most the `metaN`
        // form used by push_metadata_attributes; reserve a small fixed margin
        // while counting all escaped lexical values exactly.
        let namespace_size = escaped_xml_len(&attribute.namespace_uri, true)?;
        let local_size = attribute.local_name.len();
        let value_size = escaped_xml_len(&attribute.value, true)?;
        let attribute_size = 256usize
            .checked_add(namespace_size)
            .and_then(|size| size.checked_add(local_size))
            .and_then(|size| size.checked_add(value_size))
            .ok_or_else(|| {
                Error::InvalidFormat("text:meta serialization size overflow".to_string())
            })?;
        add_size(&mut length, attribute_size)?;
    }
    add_size(
        &mut length,
        metadata_nodes_output_len(value.content.nodes())?,
    )?;
    add_size(&mut length, "\"/>".len())?;
    if !value.content.nodes().is_empty() {
        add_size(&mut length, "</text:meta>".len())?;
    }
    Ok(length)
}

fn metadata_nodes_output_len(nodes: &[MetaFieldNode]) -> Result<usize> {
    let mut length = 0usize;
    for node in nodes {
        match node {
            MetaFieldNode::Text(value) => add_size(&mut length, escaped_xml_len(value, false)?)?,
            MetaFieldNode::Element(element) => {
                let prefix = canonical_meta_element_prefix(&element.namespace_uri)?;
                add_size(
                    &mut length,
                    1 + prefix.len()
                        + 1
                        + element.local_name.len()
                        + " xmlns:=\"\"".len()
                        + prefix.len()
                        + element.namespace_uri.len(),
                )?;
                let mut declared = Vec::new();
                declared
                    .try_reserve(element.attributes.len() + 1)
                    .map_err(|source| Error::Allocation {
                        resource: "ODT text:meta serialization namespace set",
                        source,
                    })?;
                declared.push(prefix);
                for attribute in &element.attributes {
                    let attribute_prefix = canonical_meta_element_prefix(&attribute.namespace_uri)?;
                    if attribute_prefix != "xml" && !declared.contains(&attribute_prefix) {
                        add_size(
                            &mut length,
                            " xmlns:=\"\"".len()
                                + attribute_prefix.len()
                                + attribute.namespace_uri.len(),
                        )?;
                        declared.push(attribute_prefix);
                    }
                    add_size(
                        &mut length,
                        1 + attribute_prefix.len()
                            + 1
                            + attribute.local_name.len()
                            + "=\"\"".len()
                            + escaped_xml_len(&attribute.value, true)?,
                    )?;
                }
                if element.children.is_empty() {
                    add_size(&mut length, 2)?;
                } else {
                    add_size(&mut length, 1)?;
                    add_size(&mut length, metadata_nodes_output_len(&element.children)?)?;
                    add_size(
                        &mut length,
                        "</::>".len() + prefix.len() + element.local_name.len(),
                    )?;
                }
            },
        }
    }
    Ok(length)
}

fn escaped_xml_len(value: &str, attribute: bool) -> Result<usize> {
    let mut length = 0usize;
    for byte in value.bytes() {
        let extra = match byte {
            b'&' => 4,
            b'<' | b'>' => 3,
            b'"' if attribute => 5,
            b'\'' if attribute => 5,
            _ => 0,
        };
        add_size(&mut length, 1 + extra)?;
    }
    Ok(length)
}

fn canonical_meta_element_prefix(namespace: &str) -> Result<&'static str> {
    match namespace {
        TEXTNS => Ok("text"),
        OFFICENS => Ok("office"),
        STYLENS => Ok("style"),
        XLINKNS => Ok("xlink"),
        XMLNS => Ok("xml"),
        DRAWNS => Ok("draw"),
        TABLENS => Ok("table"),
        PRESENTATIONNS => Ok("presentation"),
        SVGNS => Ok("svg"),
        FONS => Ok("fo"),
        NUMBERNS => Ok("number"),
        METANS => Ok("meta"),
        DCNS => Ok("dc"),
        XHTMLNS => Ok("xhtml"),
        DR3DNS => Ok("dr3d"),
        FORMNS => Ok("form"),
        SCRIPTNS => Ok("script"),
        _ => Err(Error::InvalidFormat(format!(
            "unsupported text:meta namespace '{namespace}'"
        ))),
    }
}

fn add_size(total: &mut usize, amount: usize) -> Result<()> {
    *total = total
        .checked_add(amount)
        .ok_or_else(|| Error::InvalidFormat("text:meta serialization size overflow".to_string()))?;
    Ok(())
}

// Keep the import visible to rustc when this module is compiled with a
// feature set that does not use the qualified-name helper directly.
#[allow(dead_code)]
fn _qualified_name(_: QName<'_>) {}

#[cfg(test)]
mod tests {
    use super::*;

    const XML: &str = r#"<?xml version="1.0"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
 xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
 xmlns:xhtml="http://www.w3.org/1999/xhtml"
 xmlns:xml="http://www.w3.org/XML/1998/namespace">
 <office:body><office:text><text:p xhtml:about="urn:before">Hello<!--keep--><text:bookmark-start text:name="mark" xhtml:property="dc:title"/> <text:meta xml:id="m1" xhtml:property="dc:description">Meta <text:span text:style-name="Emph">text</text:span></text:meta></text:p></office:text></office:body>
</office:document-content>"#;

    #[test]
    fn parses_rdfa_and_text_meta_without_evaluating_values() {
        let parsed = parse_part(XML, MetadataPart::Content).expect("metadata parses");
        assert_eq!(parsed.text_meta.len(), 1);
        assert_eq!(parsed.text_meta[0].xml_id.as_deref(), Some("m1"));
        assert_eq!(
            parsed.text_meta[0].rdfa.property.as_deref(),
            Some("dc:description")
        );
        assert_eq!(parsed.text_meta[0].content.display_text(), "Meta text");
        assert!(parsed.rdfa.iter().any(|value| {
            matches!(value.host, RdfaHost::Paragraph { index: 0 })
                && value.attributes.about.as_deref() == Some("urn:before")
        }));
        assert!(parsed.rdfa.iter().any(|value| {
            matches!(value.host, RdfaHost::BookmarkStart { .. })
                && value.attributes.property.as_deref() == Some("dc:title")
        }));
    }

    #[test]
    fn rdfa_edit_preserves_unrelated_inline_markup() {
        let value = RdfaAttributes {
            about: Some("urn:after".to_string()),
            ..RdfaAttributes::default()
        };
        let updated = set_paragraph_rdfa(XML, Position::new(0), &value).expect("paragraph edit");
        assert!(updated.contains("xhtml:about=\"urn:after\""));
        assert!(updated.contains("<!--keep-->"));
        assert!(updated.contains("text:meta"));
        let parsed = parse_part(&updated, MetadataPart::Content).expect("reparse");
        assert_eq!(parsed.text_meta[0].content.display_text(), "Meta text");
    }

    #[test]
    fn text_meta_insert_and_remove_round_trip() {
        let value = TextMeta::from_text("added").expect("text metadata");
        let inserted = insert_text_meta(XML, Position::new(0), &value).expect("insert");
        assert_eq!(
            parse_part(&inserted, MetadataPart::Content)
                .unwrap()
                .text_meta
                .len(),
            2
        );
        let removed = remove_text_meta(&inserted, Position::new(1)).expect("remove");
        assert_eq!(
            parse_part(&removed, MetadataPart::Content)
                .unwrap()
                .text_meta
                .len(),
            1
        );
    }

    #[test]
    fn equivalent_text_meta_replacement_is_an_exact_noop() {
        let parsed = parse_part(XML, MetadataPart::Content).expect("metadata parses");
        assert_eq!(
            replace_text_meta(XML, Position::new(0), &parsed.text_meta[0]).unwrap(),
            XML
        );
    }

    #[test]
    fn text_meta_unknown_root_attributes_are_retained_and_rdfa_edits_are_lossless() {
        let source = XML.replace(
            "<text:meta xml:id=\"m1\"",
            "<text:meta xmlns:ext=\"urn:example:meta\" ext:keep=\"yes\" xml:id=\"m1\"",
        );
        let parsed = parse_part(&source, MetadataPart::Content).expect("metadata parses");
        assert_eq!(parsed.text_meta[0].attributes.len(), 1);
        assert_eq!(parsed.text_meta[0].attributes[0].value, "yes");
        let serialized = parsed.text_meta[0]
            .to_xml_with_context(&[])
            .expect("metadata serializes");
        assert!(serialized.contains("ext:keep=\"yes\""));
        let updated = set_text_meta_rdfa(
            &source,
            Position::new(0),
            &RdfaAttributes {
                about: Some("urn:updated".to_string()),
                ..RdfaAttributes::default()
            },
        )
        .expect("RDFa edit");
        assert!(updated.contains("ext:keep=\"yes\""));
    }

    #[test]
    fn text_meta_accepts_soft_page_break_but_rejects_text_number() {
        let soft_break = XML.replace(
            "Meta <text:span text:style-name=\"Emph\">text</text:span>",
            "Meta<text:soft-page-break/><text:span text:style-name=\"Emph\">text</text:span>",
        );
        let parsed = parse_part(&soft_break, MetadataPart::Content).expect("soft break parses");
        assert_eq!(parsed.text_meta[0].content.display_text(), "Metatext");
        assert!(parsed.text_meta[0].content.nodes().iter().any(|node| {
            matches!(
                node,
                MetaFieldNode::Element(element)
                    if element.namespace_uri == TEXTNS && element.local_name == "soft-page-break"
            )
        }));

        let number = XML.replace(
            "Meta <text:span text:style-name=\"Emph\">text</text:span>",
            "Meta<text:number>1.</text:number>",
        );
        assert!(parse_part(&number, MetadataPart::Content).is_err());
    }

    #[test]
    fn text_meta_serialization_rejects_escaped_output_before_building_it() {
        let mut value = TextMeta::new(MetaFieldContent::new(Vec::new()).unwrap()).unwrap();
        value.attributes.push(MetadataAttribute {
            namespace_uri: String::new(),
            local_name: "flag".to_string(),
            value: "&".repeat(13 * 1024 * 1024),
            prefix: None,
        });
        assert!(value.to_xml_with_context(&[]).is_err());
    }

    #[test]
    fn rdfa_namespace_detection_ignores_values_and_rejects_shadowing() {
        let source = XML.replace(
            "<text:p xhtml:about=\"urn:before\">",
            "<text:p data-note=\"xmlns:xhtml\">Hello</text:p><text:p>",
        );
        let updated = set_paragraph_rdfa(
            &source,
            Position::new(1),
            &RdfaAttributes {
                about: Some("urn:after".to_string()),
                ..RdfaAttributes::default()
            },
        )
        .expect("RDFa namespace is added despite a value containing its name");
        assert!(updated.contains("xmlns:xhtml=\"http://www.w3.org/1999/xhtml\""));

        let shadowed = XML.replace(
            "<text:p xhtml:about=\"urn:before\">",
            "<text:p xmlns:xhtml=\"urn:wrong\">",
        );
        assert!(
            set_paragraph_rdfa(
                &shadowed,
                Position::new(0),
                &RdfaAttributes {
                    about: Some("urn:after".to_string()),
                    ..RdfaAttributes::default()
                },
            )
            .is_err()
        );
    }

    #[test]
    fn rdfa_curie_edits_require_or_add_namespace_context() {
        let missing = set_paragraph_rdfa(
            XML,
            Position::new(0),
            &RdfaAttributes {
                property: Some("custom:term".to_string()),
                ..RdfaAttributes::default()
            },
        );
        assert!(missing.is_err());

        let source = XML.replace(
            "xmlns:xhtml=\"http://www.w3.org/1999/xhtml\"",
            "xmlns:xhtml=\"http://www.w3.org/1999/xhtml\" xmlns:custom=\"urn:custom\"",
        );
        let updated = set_paragraph_rdfa(
            &source,
            Position::new(0),
            &RdfaAttributes {
                property: Some("custom:term".to_string()),
                ..RdfaAttributes::default()
            },
        )
        .expect("in-scope CURIE prefix is preserved");
        assert!(updated.contains("xmlns:custom=\"urn:custom\""));
        assert!(updated.contains("xhtml:property=\"custom:term\""));

        let added = set_paragraph_rdfa(
            XML,
            Position::new(0),
            &RdfaAttributes {
                property: Some("dc:term".to_string()),
                ..RdfaAttributes::default()
            },
        )
        .expect("canonical DC CURIE prefix is bound");
        assert!(added.contains("xmlns:dc=\"http://purl.org/dc/elements/1.1/\""));

        let custom_source = XML.replace(
            "xmlns:xhtml=\"http://www.w3.org/1999/xhtml\"",
            "xmlns:xhtml=\"http://www.w3.org/1999/xhtml\" xmlns:custom=\"urn:custom\"",
        );
        let inserted = insert_text_meta(
            &custom_source,
            Position::new(0),
            &TextMeta {
                rdfa: RdfaAttributes {
                    property: Some("custom:term".to_string()),
                    ..RdfaAttributes::default()
                },
                ..TextMeta::from_text("custom").unwrap()
            },
        )
        .expect("custom CURIE context is retained for inserted metadata");
        assert!(inserted.contains("xmlns:custom=\"urn:custom\""));
        assert!(parse_part(&inserted, MetadataPart::Content).is_ok());
    }
}
