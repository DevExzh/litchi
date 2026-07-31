//! Typed, inert Word bibliography source discovery.
//!
//! Word stores a document's current bibliography source list in a Custom XML
//! data store. This module recognizes the documented bibliography XML
//! namespaces and exposes stored scalar values only. It never resolves
//! citation tags, loads bibliography styles, runs XSLT, refreshes fields, or
//! accesses external resources.

use crate::common::xml::decode_xml_reference;
use crate::custom_xml_data::CustomXmlDataItem;
use crate::error::{OoxmlError, Result};
use litchi_opc::PackURI;
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

const MAX_BIBLIOGRAPHY_DEPTH: usize = 256;
const MAX_BIBLIOGRAPHY_SOURCES: usize = 65_536;
const MAX_BIBLIOGRAPHY_VALUES: usize = 1_000_000;
const MAX_BIBLIOGRAPHY_TEXT_BYTES: usize = 4 * 1024 * 1024;

/// One inert Word bibliography source store discovered in Custom XML.
///
/// The provenance identifies the Custom XML relationship that owns the store.
/// Styles and source values are retained as stored metadata; no citation
/// lookup, sorting, formatting, XSLT, or bibliography regeneration occurs.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct BibliographySourceStore {
    source_part_name: PackURI,
    relationship_id: String,
    data_part_name: PackURI,
    content_type: String,
    properties_part_name: Option<PackURI>,
    data_store_item_id: Option<String>,
    schema_references: Vec<String>,
    selected_style: Option<String>,
    style_name: Option<String>,
    sources: Vec<BibliographySource>,
}

impl BibliographySourceStore {
    /// Return the part that owns the Custom XML relationship.
    pub fn source_part_name(&self) -> &PackURI {
        &self.source_part_name
    }

    /// Return the owning Custom XML relationship ID.
    pub fn relationship_id(&self) -> &str {
        &self.relationship_id
    }

    /// Return the Custom XML data-part name.
    pub fn data_part_name(&self) -> &PackURI {
        &self.data_part_name
    }

    /// Return the stored Custom XML content type.
    pub fn content_type(&self) -> &str {
        &self.content_type
    }

    /// Return the optional Custom XML properties-part name.
    pub fn properties_part_name(&self) -> Option<&PackURI> {
        self.properties_part_name.as_ref()
    }

    /// Return the optional Custom XML data-store GUID.
    pub fn data_store_item_id(&self) -> Option<&str> {
        self.data_store_item_id.as_deref()
    }

    /// Return declared schema-reference URIs without resolving them.
    pub fn schema_references(&self) -> &[String] {
        &self.schema_references
    }

    /// Return the stored selected bibliography style reference, if any.
    ///
    /// This is opaque metadata. The library never opens, loads, or executes
    /// the referenced style.
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
/// Values retain stored source XML content only. The library does not
/// validate source-type requirements, resolve a `CITATION` field's tag,
/// calculate formatting, or modify the source list.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct BibliographySource {
    values: Vec<BibliographySourceValue>,
}

impl BibliographySource {
    /// Return scalar source values in persisted XML order.
    ///
    /// Every value path is relative to the `Source` element and contains only
    /// elements in the recognized bibliography namespace.
    pub fn values(&self) -> &[BibliographySourceValue] {
        &self.values
    }

    /// Return the first stored scalar value at an exact element path.
    ///
    /// For example, `source.value(&["Title"])` reads a direct title, while
    /// `source.value(&["Author", "Author", "NameList", "Person", "Last"])`
    /// reads the first matching surname. Repeated paths remain available from
    /// [`Self::values`].
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
/// Values preserve source XML order. Repeated paths, such as multiple author
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

pub(crate) fn discover_bibliography_source_stores(
    items: &[CustomXmlDataItem],
) -> Result<Vec<BibliographySourceStore>> {
    let mut stores = Vec::new();
    for item in items {
        if !is_bibliography_root(&item.root_name.namespace, &item.root_name.local_name) {
            continue;
        }
        stores.push(parse_bibliography_source_store(item)?);
    }
    Ok(stores)
}

fn parse_bibliography_source_store(item: &CustomXmlDataItem) -> Result<BibliographySourceStore> {
    let root = parse_xml_tree(&item.xml)?;
    if !is_bibliography_node(&root) || !matches!(root.local_name.as_str(), "Sources" | "Source") {
        return Err(invalid(
            "bibliography Custom XML root differs from its discovered root name",
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
        _ => unreachable!("root was checked above"),
    }

    Ok(BibliographySourceStore {
        source_part_name: item.source_part_name.clone(),
        relationship_id: item.relationship_id.clone(),
        data_part_name: item.data_part_name.clone(),
        content_type: item.content_type.clone(),
        properties_part_name: item.properties_part_name.clone(),
        data_store_item_id: item
            .properties
            .as_ref()
            .map(|properties| properties.item_id.clone()),
        schema_references: item
            .properties
            .as_ref()
            .map(|properties| properties.schema_references.clone())
            .unwrap_or_default(),
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

pub(crate) fn is_bibliography_root(namespace: &str, local_name: &str) -> bool {
    is_bibliography_namespace(namespace) && matches!(local_name, "Sources" | "Source")
}

pub(crate) fn is_bibliography_node(node: &XmlNode) -> bool {
    node.namespace
        .as_deref()
        .is_some_and(is_bibliography_namespace)
}

fn is_bibliography_namespace(value: &str) -> bool {
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

#[derive(Debug)]
pub(crate) struct XmlNode {
    pub(crate) namespace: Option<String>,
    pub(crate) local_name: String,
    pub(crate) attributes: Vec<XmlAttribute>,
    pub(crate) text: String,
    pub(crate) children: Vec<Self>,
}

#[derive(Debug)]
pub(crate) struct XmlAttribute {
    pub(crate) namespace: Option<String>,
    pub(crate) local_name: String,
    pub(crate) value: String,
}

pub(crate) fn parse_xml_tree(xml: &[u8]) -> Result<XmlNode> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut stack = Vec::new();
    let mut root = None;

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
            Event::End(_) => {
                let node = stack
                    .pop()
                    .ok_or_else(|| invalid("unexpected bibliography XML end element"))?;
                append_xml_node(node, &mut stack, &mut root)?;
            },
            Event::Text(text) => {
                let decoded = text
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| OoxmlError::Xml(error.to_string()))?;
                let decoded = quick_xml::escape::unescape(&decoded)
                    .map_err(|error| OoxmlError::Xml(error.to_string()))?;
                append_xml_text(&mut stack, &decoded)?;
            },
            Event::CData(text) => {
                let decoded = text
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| OoxmlError::Xml(error.to_string()))?;
                append_xml_text(&mut stack, &decoded)?;
            },
            Event::GeneralRef(reference) => {
                let decoded = decode_xml_reference(&reference)?;
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
        let attribute = attribute.map_err(|error| OoxmlError::Xml(error.to_string()))?;
        let raw_key = attribute.key.as_ref();
        if raw_key == b"xmlns" || raw_key.starts_with(b"xmlns:") {
            continue;
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| OoxmlError::Xml(error.to_string()))?
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
        .map_err(|error| OoxmlError::Xml(error.to_string()))
}

pub(crate) fn invalid(message: impl Into<String>) -> OoxmlError {
    OoxmlError::InvalidFormat(message.into())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::custom_xml_data::CustomXmlDataProperties;
    use litchi_ooxml_common::ExpandedName;

    fn item(xml: &[u8], namespace: &str, local_name: &str) -> CustomXmlDataItem {
        CustomXmlDataItem {
            source_part_name: PackURI::new("/word/document.xml").unwrap(),
            relationship_id: "rIdBib".to_string(),
            data_part_name: PackURI::new("/customXml/item1.xml").unwrap(),
            content_type: "application/xml".to_string(),
            root_name: ExpandedName {
                namespace: namespace.to_string(),
                local_name: local_name.to_string(),
            },
            xml: xml.to_vec(),
            properties_part_name: Some(PackURI::new("/customXml/itemProps1.xml").unwrap()),
            properties: Some(CustomXmlDataProperties {
                item_id: "{11111111-1111-1111-1111-111111111111}".to_string(),
                schema_references: vec![OOXML_BIBLIOGRAPHY_NAMESPACE.to_string()],
            }),
        }
    }

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
        let stores = discover_bibliography_source_stores(&[item(
            xml,
            OOXML_BIBLIOGRAPHY_NAMESPACE,
            "Sources",
        )])
        .unwrap();

        assert_eq!(stores.len(), 1);
        let store = &stores[0];
        assert_eq!(store.source_part_name().as_str(), "/word/document.xml");
        assert_eq!(store.relationship_id(), "rIdBib");
        assert_eq!(
            store.data_store_item_id(),
            Some("{11111111-1111-1111-1111-111111111111}")
        );
        assert_eq!(store.schema_references(), [OOXML_BIBLIOGRAPHY_NAMESPACE]);
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
    fn recognizes_legacy_single_source_payloads_without_matching_citations() {
        let xml = br#"<b:Source xmlns:b="http://schemas.microsoft.com/office/word/2004/10/bibliography"><b:Tag>Legacy</b:Tag><b:Title>Stored only</b:Title></b:Source>"#;
        let stores = discover_bibliography_source_stores(&[item(
            xml,
            LEGACY_WORD_BIBLIOGRAPHY_NAMESPACE,
            "Source",
        )])
        .unwrap();

        assert_eq!(stores.len(), 1);
        assert_eq!(stores[0].selected_style(), None);
        assert_eq!(stores[0].sources()[0].tag(), Some("Legacy"));
        assert_eq!(stores[0].sources()[0].title(), Some("Stored only"));
    }

    #[test]
    fn recognizes_strict_bibliography_source_lists() {
        let xml = br#"<b:Sources xmlns:b="http://purl.oclc.org/ooxml/officeDocument/bibliography"><b:Source><b:Tag>Strict</b:Tag></b:Source></b:Sources>"#;
        let stores = discover_bibliography_source_stores(&[item(
            xml,
            STRICT_OOXML_BIBLIOGRAPHY_NAMESPACE,
            "Sources",
        )])
        .unwrap();

        assert_eq!(stores.len(), 1);
        assert_eq!(stores[0].sources()[0].tag(), Some("Strict"));
    }

    #[test]
    fn ignores_non_bibliography_custom_xml() {
        let stores = discover_bibliography_source_stores(&[item(
            br#"<x:root xmlns:x="urn:example"><x:value>ignored</x:value></x:root>"#,
            "urn:example",
            "root",
        )])
        .unwrap();
        assert!(stores.is_empty());
    }
}
