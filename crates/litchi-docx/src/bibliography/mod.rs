//! Canonical, inert Word bibliography source-store XML.
//!
//! Word stores its bibliography source list in Custom XML using the
//! bibliography vocabulary described by the checked-in Word structures
//! (`Sources`, `Source`, `Person`, `Tag`, `LCID`, `YearAccessed`, and related
//! scalar elements).  This module owns only that namespace-aware XML model
//! and bounded parser.  It deliberately does not know about Custom XML item
//! discovery, OPC parts, relationship IDs, or package mutation.
//!
//! The parser is also intentionally independent from field evaluation.  The
//! checked-in `[MS-OE376]` Part 4 §2.16.5.11 / normative variation 2.1.494
//! describes Word's `BIBLIOGRAPHY` field switches (`\\f`, `\\l`, and repeated
//! `\\m`); none of those field semantics are applied while reading the
//! stored source XML.  `LCID`, source tags, and every other scalar are kept as
//! stored data so callers can make their own compatibility decisions.

mod package;
mod writer;

pub use package::SourceStore;
pub(crate) use package::discover_bibliography_source_stores;
pub use writer::{BibliographyPerson, BibliographySourceBuilder, BibliographySourceKind};
pub(crate) use writer::{
    DEFAULT_STORE_ITEM_ID, add_source_xml, new_store_xml, remove_source_xml, replace_source_xml,
};

use crate::{Error, Result};
use litchi_ooxml_common::xml::decode_xml_reference;
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

/// OOXML bibliography namespace used by current WordprocessingML source lists.
pub const OOXML_BIBLIOGRAPHY_NAMESPACE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/bibliography";

/// Strict OOXML bibliography namespace.
pub const STRICT_OOXML_BIBLIOGRAPHY_NAMESPACE: &str =
    "http://purl.oclc.org/ooxml/officeDocument/bibliography";

/// Legacy Word bibliography namespace accepted for interoperable source XML.
pub const LEGACY_WORD_BIBLIOGRAPHY_NAMESPACE: &str =
    "http://schemas.microsoft.com/office/word/2004/10/bibliography";

/// Maximum element nesting accepted by the standalone bibliography codec.
pub const MAX_BIBLIOGRAPHY_DEPTH: usize = 256;
/// Maximum number of `Source` elements accepted in one source store.
pub const MAX_BIBLIOGRAPHY_SOURCES: usize = 65_536;
/// Maximum number of scalar values retained from one source store.
pub const MAX_BIBLIOGRAPHY_VALUES: usize = 1_000_000;
/// Maximum bytes retained for one XML text or attribute value.
pub const MAX_BIBLIOGRAPHY_TEXT_BYTES: usize = 4 * 1024 * 1024;
/// Maximum serialized Custom XML payload accepted by the standalone codec.
pub const MAX_BIBLIOGRAPHY_XML_BYTES: usize = 32 * 1024 * 1024;

/// Parsed bibliography XML semantics without package provenance.
///
/// Custom XML item identity, relationship provenance, and package locations
/// remain in the OOXML host adapter.  This value contains only stored style
/// metadata and sources from the bibliography XML document.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct BibliographySourceStore {
    selected_style: Option<String>,
    style_name: Option<String>,
    sources: Vec<BibliographySource>,
}

impl BibliographySourceStore {
    /// Return the stored selected bibliography style reference, if any.
    ///
    /// The reference is opaque; this crate never opens or executes a style.
    pub fn selected_style(&self) -> Option<&str> {
        self.selected_style.as_deref()
    }

    /// Return the stored bibliography style name, if any.
    pub fn style_name(&self) -> Option<&str> {
        self.style_name.as_deref()
    }

    /// Return sources in their persisted XML order.
    pub fn sources(&self) -> &[BibliographySource] {
        &self.sources
    }

    /// Return the number of persisted bibliography sources.
    pub fn source_count(&self) -> usize {
        self.sources.len()
    }
}

/// One inert bibliography source from a Word Custom XML data store.
///
/// Values retain source XML content only.  The parser does not validate
/// source-type requirements, resolve a `CITATION` tag, calculate formatting,
/// or modify the source list.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct BibliographySource {
    values: Vec<BibliographySourceValue>,
}

impl BibliographySource {
    /// Return scalar source values in persisted XML order.
    ///
    /// Every value path is relative to the enclosing `Source` element and
    /// contains only elements in a recognized bibliography namespace.
    pub fn values(&self) -> &[BibliographySourceValue] {
        &self.values
    }

    /// Return the first stored scalar value at an exact element path.
    ///
    /// Repeated paths remain available from [`Self::values`].
    pub fn value(&self, path: &[&str]) -> Option<&str> {
        self.values
            .iter()
            .find(|value| {
                value.path.len() == path.len()
                    && value
                        .path
                        .iter()
                        .zip(path)
                        .all(|(stored, requested)| stored == requested)
            })
            .map(BibliographySourceValue::value)
    }

    /// Return the source tag used by `CITATION` field instructions, if stored.
    pub fn tag(&self) -> Option<&str> {
        self.value(&["Tag"])
    }

    /// Return the stored source type, such as `Book`, if present.
    pub fn source_type(&self) -> Option<&str> {
        self.value(&["SourceType"])
    }

    /// Return the stored source GUID, if present.
    pub fn guid(&self) -> Option<&str> {
        self.value(&["Guid"])
    }

    /// Return the stored source locale identifier without interpreting it.
    pub fn lcid(&self) -> Option<&str> {
        self.value(&["LCID"])
    }

    /// Return the stored title, if present.
    pub fn title(&self) -> Option<&str> {
        self.value(&["Title"])
    }

    /// Return the stored year, if present.
    pub fn year(&self) -> Option<&str> {
        self.value(&["Year"])
    }
}

/// One scalar bibliography-source XML value.
///
/// Values preserve source XML order.  Repeated paths, such as multiple author
/// names, are intentionally not collapsed.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct BibliographySourceValue {
    path: Vec<String>,
    value: String,
}

impl BibliographySourceValue {
    /// Return the path relative to the enclosing `Source` element.
    pub fn path(&self) -> &[String] {
        &self.path
    }

    /// Return the decoded stored scalar XML value.
    pub fn value(&self) -> &str {
        &self.value
    }
}

/// Parse one bounded bibliography source-store XML payload.
pub fn parse_bibliography_source_store(xml: &[u8]) -> Result<BibliographySourceStore> {
    if xml.len() > MAX_BIBLIOGRAPHY_XML_BYTES {
        return Err(invalid(format!(
            "bibliography XML exceeds {MAX_BIBLIOGRAPHY_XML_BYTES} bytes"
        )));
    }
    let root = parse_xml_tree(xml)?;
    if !is_bibliography_root(
        root.namespace.as_deref().unwrap_or_default(),
        &root.local_name,
    ) {
        return Err(invalid(
            "bibliography Custom XML root is not Sources or Source",
        ));
    }

    let mut sources = Vec::new();
    match root.local_name.as_str() {
        "Sources" => {
            for child in root
                .children
                .iter()
                .filter(|child| is_bibliography_node(child) && child.local_name == "Source")
            {
                if sources.len() >= MAX_BIBLIOGRAPHY_SOURCES {
                    return Err(invalid(format!(
                        "bibliography source count exceeds {MAX_BIBLIOGRAPHY_SOURCES}"
                    )));
                }
                sources.push(parse_bibliography_source(child)?);
            }
        },
        "Source" => sources.push(parse_bibliography_source(&root)?),
        _ => {
            return Err(invalid(
                "bibliography Custom XML root changed during parsing",
            ));
        },
    }

    Ok(BibliographySourceStore {
        selected_style: (root.local_name == "Sources")
            .then(|| unqualified_attribute(&root, "SelectedStyle"))
            .flatten()
            .map(str::to_string),
        style_name: (root.local_name == "Sources")
            .then(|| unqualified_attribute(&root, "StyleName"))
            .flatten()
            .map(str::to_string),
        sources,
    })
}

fn parse_bibliography_source(node: &XmlNode) -> Result<BibliographySource> {
    let mut values = Vec::new();
    let mut path = Vec::new();
    collect_source_values(node, &mut path, &mut values)?;
    Ok(BibliographySource { values })
}

fn collect_source_values(
    parent: &XmlNode,
    path: &mut Vec<String>,
    values: &mut Vec<BibliographySourceValue>,
) -> Result<()> {
    for child in parent
        .children
        .iter()
        .filter(|child| is_bibliography_node(child))
    {
        path.push(child.local_name.clone());
        let has_bibliography_children = child.children.iter().any(is_bibliography_node);
        if has_bibliography_children {
            collect_source_values(child, path, values)?;
        } else {
            if values.len() >= MAX_BIBLIOGRAPHY_VALUES {
                return Err(invalid(format!(
                    "bibliography source value count exceeds {MAX_BIBLIOGRAPHY_VALUES}"
                )));
            }
            if child.text.len() > MAX_BIBLIOGRAPHY_TEXT_BYTES {
                return Err(invalid(format!(
                    "bibliography source value exceeds {MAX_BIBLIOGRAPHY_TEXT_BYTES} bytes"
                )));
            }
            values.push(BibliographySourceValue {
                path: path.clone(),
                value: child.text.clone(),
            });
        }
        path.pop();
    }
    Ok(())
}

/// Return whether a namespace/local-name pair is a bibliography root.
pub fn is_bibliography_root(namespace: &str, local_name: &str) -> bool {
    is_bibliography_namespace(namespace) && matches!(local_name, "Sources" | "Source")
}

/// Return whether a parsed node belongs to a recognized bibliography namespace.
pub fn is_bibliography_node(node: &XmlNode) -> bool {
    node.namespace
        .as_deref()
        .is_some_and(is_bibliography_namespace)
}

/// Return whether `value` is one of the accepted bibliography namespace URIs.
pub fn is_bibliography_namespace(value: &str) -> bool {
    matches!(
        value,
        OOXML_BIBLIOGRAPHY_NAMESPACE
            | STRICT_OOXML_BIBLIOGRAPHY_NAMESPACE
            | LEGACY_WORD_BIBLIOGRAPHY_NAMESPACE
    )
}

fn unqualified_attribute<'a>(node: &'a XmlNode, local_name: &str) -> Option<&'a str> {
    node.attributes
        .iter()
        .find(|attribute| attribute.namespace.is_none() && attribute.local_name == local_name)
        .map(|attribute| attribute.value.as_str())
}

/// XML tree used by the host bibliography writer's compatibility adapter.
///
/// The fields are public because the historical host writer constructs and
/// updates this tree.  The type is not re-exported from the crate root and is
/// not a package graph handle.
#[doc(hidden)]
#[derive(Debug)]
pub struct XmlNode {
    pub namespace: Option<String>,
    pub local_name: String,
    pub attributes: Vec<XmlAttribute>,
    pub text: String,
    pub children: Vec<Self>,
}

/// Namespace-aware XML attribute used by [`XmlNode`].
#[doc(hidden)]
#[derive(Debug)]
pub struct XmlAttribute {
    pub namespace: Option<String>,
    pub local_name: String,
    pub value: String,
}

/// Parse a bounded namespace-aware XML tree for bibliography authoring.
///
/// DTDs and external entities are rejected.  The tree is retained only for
/// the host's inert source CRUD adapter.
#[doc(hidden)]
pub fn parse_xml_tree(xml: &[u8]) -> Result<XmlNode> {
    if xml.len() > MAX_BIBLIOGRAPHY_XML_BYTES {
        return Err(invalid(format!(
            "bibliography XML exceeds {MAX_BIBLIOGRAPHY_XML_BYTES} bytes"
        )));
    }
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut stack = Vec::new();
    let mut root = None;

    loop {
        let decoder = reader.decoder();
        let event = reader.read_event()?.into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                if stack.len() >= MAX_BIBLIOGRAPHY_DEPTH {
                    return Err(invalid(format!(
                        "bibliography XML depth exceeds {MAX_BIBLIOGRAPHY_DEPTH}"
                    )));
                }
                stack.push(xml_node(namespace, &element, decoder, &resolver)?);
            },
            Event::Empty(element) => append_xml_node(
                xml_node(namespace, &element, decoder, &resolver)?,
                &mut stack,
                &mut root,
            )?,
            Event::End(element) => {
                let node = stack
                    .pop()
                    .ok_or_else(|| invalid("unexpected bibliography XML end element"))?;
                let end_namespace = owned_namespace(namespace)?;
                let end_name = utf8_name(element.local_name().as_ref())?;
                if node.namespace != end_namespace || node.local_name != end_name {
                    return Err(invalid(format!(
                        "bibliography XML end element does not match '{}', got '{}'",
                        node.local_name, end_name
                    )));
                }
                append_xml_node(node, &mut stack, &mut root)?;
            },
            Event::Text(text) => {
                let decoded = text
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| Error::Xml(error.to_string()))?;
                let decoded = quick_xml::escape::unescape(&decoded)
                    .map_err(|error| Error::Xml(error.to_string()))?;
                append_xml_text(&mut stack, &decoded)?;
            },
            Event::CData(text) => {
                let decoded = text
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| Error::Xml(error.to_string()))?;
                append_xml_text(&mut stack, &decoded)?;
            },
            Event::GeneralRef(reference) => {
                let decoded = decode_xml_reference(&reference)
                    .map_err(|error| Error::Xml(error.to_string()))?;
                append_xml_text(&mut stack, &decoded)?;
            },
            Event::DocType(_) => return Err(invalid("DTD is not allowed in bibliography XML")),
            Event::Eof => break,
            _ => {},
        }
    }

    if !stack.is_empty() {
        return Err(invalid("unterminated bibliography XML"));
    }
    root.ok_or_else(|| invalid("bibliography XML has no root element"))
}

fn xml_node(
    namespace: ResolveResult<'_>,
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    resolver: &NamespaceResolver,
) -> Result<XmlNode> {
    let mut attributes = Vec::new();
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let raw_key = attribute.key.as_ref();
        if raw_key == b"xmlns" || raw_key.starts_with(b"xmlns:") {
            continue;
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        if value.len() > MAX_BIBLIOGRAPHY_TEXT_BYTES {
            return Err(invalid(format!(
                "bibliography XML attribute exceeds {MAX_BIBLIOGRAPHY_TEXT_BYTES} bytes"
            )));
        }
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        attributes.push(XmlAttribute {
            namespace: owned_namespace(namespace)?,
            local_name: utf8_name(attribute.key.local_name().as_ref())?,
            value,
        });
    }
    Ok(XmlNode {
        namespace: owned_namespace(namespace)?,
        local_name: utf8_name(element.local_name().as_ref())?,
        attributes,
        text: String::new(),
        children: Vec::new(),
    })
}

fn append_xml_node(node: XmlNode, stack: &mut [XmlNode], root: &mut Option<XmlNode>) -> Result<()> {
    if let Some(parent) = stack.last_mut() {
        parent.children.push(node);
    } else if root.replace(node).is_some() {
        return Err(invalid("bibliography XML has multiple root elements"));
    }
    Ok(())
}

fn append_xml_text(stack: &mut [XmlNode], text: &str) -> Result<()> {
    let Some(node) = stack.last_mut() else {
        if text.trim().is_empty() {
            return Ok(());
        }
        return Err(invalid("bibliography XML has text outside its root"));
    };
    let length = node
        .text
        .len()
        .checked_add(text.len())
        .ok_or_else(|| invalid("bibliography XML text length overflow"))?;
    if length > MAX_BIBLIOGRAPHY_TEXT_BYTES {
        return Err(invalid(format!(
            "bibliography XML text exceeds {MAX_BIBLIOGRAPHY_TEXT_BYTES} bytes"
        )));
    }
    node.text.push_str(text);
    Ok(())
}

fn owned_namespace(namespace: ResolveResult<'_>) -> Result<Option<String>> {
    match namespace {
        ResolveResult::Bound(Namespace(value)) => Ok(Some(utf8_name(value)?)),
        ResolveResult::Unbound | ResolveResult::Unknown(_) => Ok(None),
    }
}

fn utf8_name(value: &[u8]) -> Result<String> {
    std::str::from_utf8(value)
        .map(str::to_string)
        .map_err(|error| Error::Xml(error.to_string()))
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_source_store_metadata_and_nested_scalar_values() {
        let xml = br#"<b:Sources xmlns:b="http://schemas.openxmlformats.org/officeDocument/2006/bibliography" SelectedStyle="/APA.XSL" StyleName="APA">
            <b:Source>
                <b:Tag>Doe2024</b:Tag>
                <b:SourceType>Book</b:SourceType>
                <b:Guid>{11111111-1111-1111-1111-111111111111}</b:Guid>
                <b:LCID>1033</b:LCID>
                <b:Title>Example &amp; Practice</b:Title>
                <b:Year>2024</b:Year>
                <b:Author><b:Author><b:NameList><b:Person><b:Last>Doe</b:Last><b:First>Jane</b:First></b:Person></b:NameList></b:Author></b:Author>
            </b:Source>
            <b:Source><b:Tag>Smith2025</b:Tag><b:SourceType>ArticleInAPeriodical</b:SourceType></b:Source>
        </b:Sources>"#;
        let store = parse_bibliography_source_store(xml).unwrap();

        assert_eq!(store.selected_style(), Some("/APA.XSL"));
        assert_eq!(store.style_name(), Some("APA"));
        assert_eq!(store.source_count(), 2);
        let doe = &store.sources()[0];
        assert_eq!(doe.tag(), Some("Doe2024"));
        assert_eq!(doe.source_type(), Some("Book"));
        assert_eq!(doe.guid(), Some("{11111111-1111-1111-1111-111111111111}"));
        assert_eq!(doe.lcid(), Some("1033"));
        assert_eq!(doe.title(), Some("Example & Practice"));
        assert_eq!(doe.year(), Some("2024"));
        assert_eq!(
            doe.value(&["Author", "Author", "NameList", "Person", "Last"]),
            Some("Doe")
        );
        assert_eq!(
            doe.values()[6].path(),
            ["Author", "Author", "NameList", "Person", "Last"]
        );
        assert_eq!(doe.values()[6].value(), "Doe");
        assert_eq!(store.sources()[1].tag(), Some("Smith2025"));
    }

    #[test]
    fn recognizes_strict_and_legacy_source_payloads() {
        let strict = format!(
            "<b:Sources xmlns:b=\"{STRICT_OOXML_BIBLIOGRAPHY_NAMESPACE}\"><b:Source><b:Tag>Strict</b:Tag></b:Source></b:Sources>"
        );
        let store = parse_bibliography_source_store(strict.as_bytes()).unwrap();
        assert_eq!(store.sources()[0].tag(), Some("Strict"));

        let legacy = format!(
            "<b:Source xmlns:b=\"{LEGACY_WORD_BIBLIOGRAPHY_NAMESPACE}\"><b:Tag>Legacy</b:Tag><b:Title>Stored only</b:Title></b:Source>"
        );
        let store = parse_bibliography_source_store(legacy.as_bytes()).unwrap();
        assert_eq!(store.selected_style(), None);
        assert_eq!(store.sources()[0].tag(), Some("Legacy"));
        assert_eq!(store.sources()[0].title(), Some("Stored only"));
    }

    #[test]
    fn ignores_foreign_children_but_preserves_bibliography_scalar_order() {
        let xml = br#"<b:Sources xmlns:b="http://schemas.openxmlformats.org/officeDocument/2006/bibliography" xmlns:x="urn:example"><b:Source><b:Tag>First</b:Tag><x:ignored>not a source value</x:ignored><b:Title>Second</b:Title></b:Source></b:Sources>"#;
        let store = parse_bibliography_source_store(xml).unwrap();
        let values = store.sources()[0].values();
        assert_eq!(values.len(), 2);
        assert_eq!(values[0].path(), ["Tag"]);
        assert_eq!(values[0].value(), "First");
        assert_eq!(values[1].path(), ["Title"]);
        assert_eq!(values[1].value(), "Second");
    }

    #[test]
    fn rejects_dtd_and_mismatched_end_elements() {
        assert!(parse_bibliography_source_store(
            br#"<!DOCTYPE b:Sources><b:Sources xmlns:b="http://schemas.openxmlformats.org/officeDocument/2006/bibliography"/>"#
        )
        .is_err());
        assert!(parse_xml_tree(
            br#"<b:Sources xmlns:b="http://schemas.openxmlformats.org/officeDocument/2006/bibliography"><b:Source></b:Sources>"#
        )
        .is_err());
    }
}
