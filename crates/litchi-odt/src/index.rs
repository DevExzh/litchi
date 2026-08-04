//! Inert semantic access to generated OpenDocument text indexes.

mod writing;

pub use writing::{
    AlphabeticalIndexSource, BibliographyIndexSource, IllustrationIndexSource, ObjectIndexSource,
    TableOfContentsSource, TextAlphabeticalIndexEntryTemplate,
    TextAlphabeticalIndexLevel, TextBibliographyEntryTemplate, TextBibliographyEntryToken,
    TextBibliographyType, TextIndexBody, TextIndexBodyParagraph, TextIndexBodyTitle,
    TextIndexCaptionSequenceFormat, TextIndexChapterDisplay, TextIndexEntryTemplate,
    TextIndexEntryToken, TextIndexScope, TextIndexSimpleEntryTemplate, TextIndexSourceStyles,
    TextIndexTabStop, TextIndexTitleTemplate, UserIndexSource, insert_text_index_xml,
    remove_text_index_xml, replace_text_index_xml,
};

use crate::elements::xml::{
    TEXT_NAMESPACE, append_checked, decode_reference, is_bound, namespaced_attribute,
};
use litchi_core::{Error, Result};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

const MAX_INDEX_DEPTH: usize = 4_096;
const MAX_INDEX_ITEMS: usize = 1_000_000;

/// The seven generated-index families defined by OpenDocument Text.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum TextIndexKind {
    TableOfContents,
    Illustration,
    Table,
    Object,
    User,
    Alphabetical,
    Bibliography,
}

/// A decoded attribute identified by its expanded XML name.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TextIndexAttribute {
    pub(crate) namespace_uri: Option<String>,
    pub(crate) local_name: String,
    pub(crate) value: String,
}

impl TextIndexAttribute {
    pub fn namespace_uri(&self) -> Option<&str> {
        self.namespace_uri.as_deref()
    }

    pub fn local_name(&self) -> &str {
        &self.local_name
    }

    pub fn value(&self) -> &str {
        &self.value
    }
}

/// Ordered mixed content within an index element.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum TextIndexContent {
    Text(String),
    Element(TextIndexElement),
}

/// One namespace-aware element in an index source, template, or cached body.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TextIndexElement {
    namespace_uri: Option<String>,
    local_name: String,
    attributes: Vec<TextIndexAttribute>,
    content: Vec<TextIndexContent>,
}

impl TextIndexElement {
    pub fn namespace_uri(&self) -> Option<&str> {
        self.namespace_uri.as_deref()
    }

    pub fn local_name(&self) -> &str {
        &self.local_name
    }

    pub fn attributes(&self) -> &[TextIndexAttribute] {
        &self.attributes
    }

    pub fn attribute(&self, namespace_uri: Option<&str>, local_name: &str) -> Option<&str> {
        self.attributes
            .iter()
            .find(|attribute| {
                attribute.namespace_uri() == namespace_uri && attribute.local_name() == local_name
            })
            .map(TextIndexAttribute::value)
    }

    pub fn content(&self) -> &[TextIndexContent] {
        &self.content
    }

    pub fn child_elements(&self) -> impl Iterator<Item = &TextIndexElement> {
        self.content.iter().filter_map(|content| match content {
            TextIndexContent::Element(element) => Some(element),
            TextIndexContent::Text(_) => None,
        })
    }

    /// Compose character content in exact document order.
    pub fn all_text(&self) -> String {
        let mut output = String::new();
        append_all_text(self, &mut output);
        output
    }
}

/// A generated index declaration and its stored, inert cached body.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TextIndex {
    kind: TextIndexKind,
    root: TextIndexElement,
}

impl TextIndex {
    pub fn kind(&self) -> TextIndexKind {
        self.kind
    }

    pub fn root(&self) -> &TextIndexElement {
        &self.root
    }

    pub fn name(&self) -> &str {
        self.root
            .attribute(
                Some(std::str::from_utf8(TEXT_NAMESPACE).expect("ASCII namespace")),
                "name",
            )
            .expect("validated index name")
    }

    pub fn protected(&self) -> bool {
        matches!(
            self.root.attribute(
                Some(std::str::from_utf8(TEXT_NAMESPACE).expect("ASCII namespace")),
                "protected"
            ),
            Some("true" | "1")
        )
    }

    pub fn source(&self) -> Option<&TextIndexElement> {
        self.root.child_elements().find(|element| {
            element.namespace_uri()
                == Some(std::str::from_utf8(TEXT_NAMESPACE).expect("ASCII namespace"))
                && element.local_name().ends_with("-source")
        })
    }

    pub fn body(&self) -> Option<&TextIndexElement> {
        self.root.child_elements().find(|element| {
            element.namespace_uri()
                == Some(std::str::from_utf8(TEXT_NAMESPACE).expect("ASCII namespace"))
                && element.local_name() == "index-body"
        })
    }
}

struct ActiveIndex {
    kind: TextIndexKind,
    stack: Vec<TextIndexElement>,
    order: usize,
    item_count: usize,
}

pub(crate) fn parse_text_indexes(xml: &str) -> Result<Vec<TextIndex>> {
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut document_depth = 0usize;
    let mut active = Vec::<ActiveIndex>::new();
    let mut indexes = Vec::new();
    let mut next_order = 0usize;

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid text-index XML: {error}")))?;
        let text_element = is_bound(&namespace, TEXT_NAMESPACE);
        match event {
            Event::Start(ref source) => {
                document_depth = checked_depth(document_depth)?;
                let kind = text_element
                    .then(|| index_kind(source.local_name().as_ref()))
                    .flatten();
                if !active.is_empty() || kind.is_some() {
                    let namespace_uri = resolved_namespace(&namespace, "text-index element")?;
                    let node = element_from_start(&reader, namespace_uri, source)?;
                    for index in &mut active {
                        add_index_item(index)?;
                        index.stack.push(node.clone());
                    }
                    if let Some(kind) = kind {
                        validate_index_root(&reader, source)?;
                        if next_order >= MAX_INDEX_ITEMS {
                            return Err(Error::InvalidFormat(format!(
                                "document exceeds {MAX_INDEX_ITEMS} text indexes"
                            )));
                        }
                        active.push(ActiveIndex {
                            kind,
                            stack: vec![node],
                            order: next_order,
                            item_count: 1,
                        });
                        next_order += 1;
                    }
                }
            },
            Event::Empty(ref source) => {
                let kind = text_element
                    .then(|| index_kind(source.local_name().as_ref()))
                    .flatten();
                if !active.is_empty() || kind.is_some() {
                    let namespace_uri = resolved_namespace(&namespace, "text-index element")?;
                    let node = element_from_start(&reader, namespace_uri, source)?;
                    for index in &mut active {
                        add_index_item(index)?;
                        index
                            .stack
                            .last_mut()
                            .expect("active index stack")
                            .content
                            .push(TextIndexContent::Element(node.clone()));
                    }
                    if let Some(kind) = kind {
                        validate_index_root(&reader, source)?;
                        if next_order >= MAX_INDEX_ITEMS {
                            return Err(Error::InvalidFormat(format!(
                                "document exceeds {MAX_INDEX_ITEMS} text indexes"
                            )));
                        }
                        indexes.push((next_order, TextIndex { kind, root: node }));
                        next_order += 1;
                    }
                }
            },
            Event::Text(ref value) if !active.is_empty() => {
                let value = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid text-index text: {error}"))
                    })?;
                append_index_text(&mut active, &value)?;
            },
            Event::CData(ref value) if !active.is_empty() => {
                let value = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid text-index CDATA: {error}"))
                    })?;
                append_index_text(&mut active, &value)?;
            },
            Event::GeneralRef(ref reference) if !active.is_empty() => {
                append_index_text(&mut active, &decode_reference(reference, "text index")?)?;
            },
            Event::End(_) => {
                document_depth = document_depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("text-index XML stack underflow".to_string())
                })?;
                for position in (0..active.len()).rev() {
                    let node = active[position].stack.pop().expect("active index stack");
                    if let Some(parent) = active[position].stack.last_mut() {
                        parent.content.push(TextIndexContent::Element(node));
                    } else {
                        let finished = active.remove(position);
                        indexes.push((
                            finished.order,
                            TextIndex {
                                kind: finished.kind,
                                root: node,
                            },
                        ));
                    }
                }
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if document_depth != 0 || !active.is_empty() {
        return Err(Error::InvalidFormat(
            "incomplete text-index XML structure".to_string(),
        ));
    }
    indexes.sort_by_key(|(order, _)| *order);
    Ok(indexes.into_iter().map(|(_, index)| index).collect())
}

fn index_kind(local_name: &[u8]) -> Option<TextIndexKind> {
    match local_name {
        b"table-of-content" => Some(TextIndexKind::TableOfContents),
        b"illustration-index" => Some(TextIndexKind::Illustration),
        b"table-index" => Some(TextIndexKind::Table),
        b"object-index" => Some(TextIndexKind::Object),
        b"user-index" => Some(TextIndexKind::User),
        b"alphabetical-index" => Some(TextIndexKind::Alphabetical),
        b"bibliography" => Some(TextIndexKind::Bibliography),
        _ => None,
    }
}

fn validate_index_root(reader: &NsReader<&[u8]>, source: &BytesStart<'_>) -> Result<()> {
    namespaced_attribute(reader, source, TEXT_NAMESPACE, b"name", "text index")?
        .ok_or_else(|| Error::InvalidFormat("text index requires text:name".to_string()))?;
    if let Some(value) =
        namespaced_attribute(reader, source, TEXT_NAMESPACE, b"protected", "text index")?
        && !matches!(value.as_str(), "true" | "false" | "1" | "0")
    {
        return Err(Error::InvalidFormat(
            "text:protected must be true, false, 1, or 0".to_string(),
        ));
    }
    Ok(())
}

fn element_from_start(
    reader: &NsReader<&[u8]>,
    namespace_uri: Option<String>,
    source: &BytesStart<'_>,
) -> Result<TextIndexElement> {
    let local_name = std::str::from_utf8(source.local_name().as_ref())
        .map_err(|_| Error::InvalidFormat("non-UTF-8 text-index element name".to_string()))?
        .to_string();
    let attributes = expanded_attributes(reader, source, "text index")?;
    Ok(TextIndexElement {
        namespace_uri,
        local_name,
        attributes,
        content: Vec::new(),
    })
}

pub(crate) fn expanded_attributes(
    reader: &NsReader<&[u8]>,
    source: &BytesStart<'_>,
    context: &str,
) -> Result<Vec<TextIndexAttribute>> {
    let mut attributes = Vec::new();
    for attribute in source.attributes() {
        let attribute = attribute.map_err(|error| {
            Error::InvalidFormat(format!("invalid {context} attribute: {error}"))
        })?;
        if attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local_name) = reader.resolver().resolve_attribute(attribute.key);
        let namespace_uri = resolved_namespace(&namespace, context)?;
        let local_name = std::str::from_utf8(local_name.as_ref())
            .map_err(|_| Error::InvalidFormat(format!("non-UTF-8 {context} attribute name")))?
            .to_string();
        if attributes.iter().any(|existing: &TextIndexAttribute| {
            existing.namespace_uri == namespace_uri && existing.local_name == local_name
        }) {
            return Err(Error::InvalidFormat(format!(
                "duplicate expanded {context} attribute '{local_name}'"
            )));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid {context} attribute value: {error}"))
            })?
            .into_owned();
        attributes.push(TextIndexAttribute {
            namespace_uri,
            local_name,
            value,
        });
    }
    Ok(attributes)
}

fn resolved_namespace(namespace: &ResolveResult<'_>, context: &str) -> Result<Option<String>> {
    match namespace {
        ResolveResult::Bound(Namespace(uri)) => std::str::from_utf8(uri)
            .map(|uri| Some(uri.to_string()))
            .map_err(|_| Error::InvalidFormat(format!("non-UTF-8 {context} namespace URI"))),
        ResolveResult::Unbound => Ok(None),
        ResolveResult::Unknown(prefix) => Err(Error::InvalidFormat(format!(
            "unknown {context} namespace prefix '{}'",
            String::from_utf8_lossy(prefix)
        ))),
    }
}

fn append_index_text(active: &mut [ActiveIndex], value: &str) -> Result<()> {
    for index in active {
        let element = index.stack.last_mut().expect("active index stack");
        if let Some(TextIndexContent::Text(text)) = element.content.last_mut() {
            append_checked(text, value)?;
        } else {
            let mut text = String::new();
            append_checked(&mut text, value)?;
            element.content.push(TextIndexContent::Text(text));
            add_index_item(index)?;
        }
    }
    Ok(())
}

fn append_all_text(element: &TextIndexElement, output: &mut String) {
    for content in &element.content {
        match content {
            TextIndexContent::Text(text) => output.push_str(text),
            TextIndexContent::Element(child) => append_all_text(child, output),
        }
    }
}

fn add_index_item(index: &mut ActiveIndex) -> Result<()> {
    index.item_count = index
        .item_count
        .checked_add(1)
        .ok_or_else(|| Error::InvalidFormat("text-index item count overflow".to_string()))?;
    if index.item_count > MAX_INDEX_ITEMS {
        return Err(Error::InvalidFormat(format!(
            "text index exceeds {MAX_INDEX_ITEMS} items"
        )));
    }
    Ok(())
}

fn checked_depth(depth: usize) -> Result<usize> {
    let depth = depth
        .checked_add(1)
        .ok_or_else(|| Error::InvalidFormat("text-index nesting depth overflow".to_string()))?;
    if depth > MAX_INDEX_DEPTH {
        return Err(Error::InvalidFormat(format!(
            "text-index nesting exceeds {MAX_INDEX_DEPTH} levels"
        )));
    }
    Ok(depth)
}

#[cfg(test)]
mod tests {
    use super::*;

    const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
    const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

    #[test]
    fn parses_every_text_index_kind_and_complete_ordered_subtrees() {
        let xml = format!(
            r#"<o:document-content xmlns:o="{OFFICE}" xmlns:t="{TEXT}" xmlns:u="urn:vendor"><o:body><o:text><t:table-of-content t:name="Contents &amp; More" t:protected="1" u:future="yes"><t:table-of-content-source t:outline-level="3"><t:index-title-template t:style-name="Title">Contents</t:index-title-template><t:table-of-content-entry-template t:outline-level="1" t:style-name="Entry"><t:index-entry-text/><u:extension u:value="A&amp;B"/></t:table-of-content-entry-template></t:table-of-content-source><t:index-body><t:index-title t:name="Cached"><t:p>Title</t:p></t:index-title><t:p>Pre&amp;<t:span>Mid</t:span><![CDATA[!]]>Post</t:p><t:user-index t:name="Nested"><t:user-index-source t:index-name="N"/><t:index-body><t:p>Nested</t:p></t:index-body></t:user-index></t:index-body></t:table-of-content><t:illustration-index t:name="I"/><t:table-index t:name="T"/><t:object-index t:name="O"/><t:alphabetical-index t:name="A"/><t:bibliography t:name="B"/></o:text></o:body></o:document-content>"#
        );
        let indexes = parse_text_indexes(&xml).unwrap();
        assert_eq!(indexes.len(), 7);
        assert_eq!(indexes[0].kind(), TextIndexKind::TableOfContents);
        assert_eq!(indexes[0].name(), "Contents & More");
        assert!(indexes[0].protected());
        assert_eq!(
            indexes[0].root().attribute(Some("urn:vendor"), "future"),
            Some("yes")
        );
        let source = indexes[0].source().unwrap();
        assert_eq!(source.local_name(), "table-of-content-source");
        assert_eq!(source.attribute(Some(TEXT), "outline-level"), Some("3"));
        let template = source
            .child_elements()
            .find(|element| element.local_name() == "table-of-content-entry-template")
            .unwrap();
        let extension = template
            .child_elements()
            .find(|element| element.namespace_uri() == Some("urn:vendor"))
            .unwrap();
        assert_eq!(
            extension.attribute(Some("urn:vendor"), "value"),
            Some("A&B")
        );

        let body = indexes[0].body().unwrap();
        let paragraph = body
            .child_elements()
            .find(|element| element.local_name() == "p" && element.all_text().starts_with("Pre"))
            .unwrap();
        assert_eq!(paragraph.all_text(), "Pre&Mid!Post");
        assert!(matches!(paragraph.content()[0], TextIndexContent::Text(_)));
        assert!(matches!(
            paragraph.content()[1],
            TextIndexContent::Element(_)
        ));
        assert!(matches!(paragraph.content()[2], TextIndexContent::Text(_)));

        assert_eq!(indexes[1].kind(), TextIndexKind::User);
        assert_eq!(indexes[1].name(), "Nested");
        assert_eq!(indexes[1].body().unwrap().all_text(), "Nested");
        assert_eq!(indexes[2].kind(), TextIndexKind::Illustration);
        assert_eq!(indexes[3].kind(), TextIndexKind::Table);
        assert_eq!(indexes[4].kind(), TextIndexKind::Object);
        assert_eq!(indexes[5].kind(), TextIndexKind::Alphabetical);
        assert_eq!(indexes[6].kind(), TextIndexKind::Bibliography);
    }

    #[test]
    fn text_indexes_reject_malformed_ambiguous_or_invalid_roots() {
        let missing_name = format!(r#"<t:table-of-content xmlns:t="{TEXT}"/>"#);
        assert!(parse_text_indexes(&missing_name).is_err());

        let invalid_boolean =
            format!(r#"<t:table-index xmlns:t="{TEXT}" t:name="T" t:protected="yes"/>"#);
        assert!(parse_text_indexes(&invalid_boolean).is_err());

        let duplicate =
            format!(r#"<t:object-index xmlns:t="{TEXT}" xmlns:u="{TEXT}" t:name="A" u:name="B"/>"#);
        assert!(parse_text_indexes(&duplicate).is_err());

        let unknown_prefix =
            format!(r#"<t:bibliography xmlns:t="{TEXT}" t:name="B" x:value="bad"/>"#);
        assert!(parse_text_indexes(&unknown_prefix).is_err());
        assert!(parse_text_indexes("<t:table-index>").is_err());
    }

    #[test]
    fn text_indexes_enforce_nesting_bound() {
        let mut xml = format!(r#"<t:table-of-content xmlns:t="{TEXT}" t:name="T">"#);
        for _ in 0..MAX_INDEX_DEPTH {
            xml.push_str("<t:span>");
        }
        for _ in 0..MAX_INDEX_DEPTH {
            xml.push_str("</t:span>");
        }
        xml.push_str("</t:table-of-content>");
        assert!(parse_text_indexes(&xml).is_err());
    }
}
