//! Bibliography source-store authoring for DOCX documents.
//!
//! Word keeps its current bibliography source list in a Custom XML data
//! store whose root is `b:Sources` in the OOXML bibliography namespace. This
//! module builds typed `b:Source` entries and mutates the store XML in place
//! (add / remove / replace by tag), preserving untouched entries, root
//! attributes (`SelectedStyle` / `StyleName`), and the Custom XML
//! relationship/content-type graph. Everything remains inert: no citation
//! matching, style resolution, XSLT, or bibliography regeneration occurs.
//!
//! The read side is [`crate::bibliography`].

use crate::bibliography::{
    OOXML_BIBLIOGRAPHY_NAMESPACE, XmlNode, invalid, is_bibliography_node, is_bibliography_root,
    parse_xml_tree,
};
use crate::error::{Error, Result};
use litchi_core::xml::escape_xml;
use std::fmt::Write as FmtWrite;

/// Default Custom XML item GUID used when creating a bibliography store.
pub(crate) const DEFAULT_STORE_ITEM_ID: &str = "{B1B10000-0000-0000-0000-000000000001}";

const MAX_TAG_LENGTH: usize = 255;
const MAX_FIELD_BYTES: usize = 4096;
const MAX_PERSONS: usize = 256;
const MAX_CUSTOM_FIELDS: usize = 64;
const MAX_CUSTOM_PATH_DEPTH: usize = 8;
/// The relationship ID space is unbounded; the store source count inherits
/// the read-side bound.
const MAX_STORE_SOURCES: usize = 65_536;

/// Bibliography source type (`b:SourceType`, ST_SourceType).
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
#[repr(u8)]
pub enum BibliographySourceKind {
    /// `Book`.
    #[default]
    Book,
    /// `BookSection`.
    BookSection,
    /// `JournalArticle`.
    JournalArticle,
    /// `ArticleInAPeriodical`.
    ArticleInAPeriodical,
    /// `ConferenceProceedings`.
    ConferenceProceedings,
    /// `Report`.
    Report,
    /// `SoundRecording`.
    SoundRecording,
    /// `Performance`.
    Performance,
    /// `Art`.
    Art,
    /// `DocumentFromInternetSite`.
    DocumentFromInternetSite,
    /// `InternetSite`.
    InternetSite,
    /// `Film`.
    Film,
    /// `Interview`.
    Interview,
    /// `Patent`.
    Patent,
    /// `ElectronicSource`.
    ElectronicSource,
    /// `Case`.
    Case,
    /// `Misc`.
    Misc,
}

impl BibliographySourceKind {
    /// The ST_SourceType token for this kind.
    pub fn as_token(self) -> &'static str {
        match self {
            Self::Book => "Book",
            Self::BookSection => "BookSection",
            Self::JournalArticle => "JournalArticle",
            Self::ArticleInAPeriodical => "ArticleInAPeriodical",
            Self::ConferenceProceedings => "ConferenceProceedings",
            Self::Report => "Report",
            Self::SoundRecording => "SoundRecording",
            Self::Performance => "Performance",
            Self::Art => "Art",
            Self::DocumentFromInternetSite => "DocumentFromInternetSite",
            Self::InternetSite => "InternetSite",
            Self::Film => "Film",
            Self::Interview => "Interview",
            Self::Patent => "Patent",
            Self::ElectronicSource => "ElectronicSource",
            Self::Case => "Case",
            Self::Misc => "Misc",
        }
    }
}

/// One contributor name for a bibliography source (`b:Person`).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct BibliographyPerson {
    /// Family name (`b:Last`, required).
    pub last: String,
    /// Given name (`b:First`).
    pub first: Option<String>,
    /// Middle name (`b:Middle`).
    pub middle: Option<String>,
}

impl BibliographyPerson {
    /// Create a person with a family name.
    pub fn new(last: impl Into<String>) -> Result<Self> {
        let last = last.into();
        validate_field("Last", &last)?;
        if last.is_empty() {
            return Err(invalid("bibliography person requires a family name"));
        }
        Ok(Self {
            last,
            first: None,
            middle: None,
        })
    }

    /// Set the given name.
    pub fn with_first(mut self, first: impl Into<String>) -> Result<Self> {
        let first = first.into();
        validate_field("First", &first)?;
        self.first = Some(first);
        Ok(self)
    }

    /// Set the middle name.
    pub fn with_middle(mut self, middle: impl Into<String>) -> Result<Self> {
        let middle = middle.into();
        validate_field("Middle", &middle)?;
        self.middle = Some(middle);
        Ok(self)
    }
}

/// A typed bibliography source being authored into a source store.
///
/// Built with [`BibliographySourceBuilder::new`], optional fields chained,
/// then passed to
/// [`crate::Package::add_bibliography_source`] or
/// [`crate::Package::replace_bibliography_source`]. Fields not covered
/// by the typed API can be carried with [`Self::custom_field`] so nothing is
/// silently dropped.
///
/// # Examples
///
/// ```rust,no_run
/// use litchi_docx::bibliography::{
///     BibliographyPerson, BibliographySourceBuilder, BibliographySourceKind,
/// };
/// use litchi_docx::Package;
///
/// let mut package = Package::new()?;
/// let source = BibliographySourceBuilder::new(BibliographySourceKind::Book, "Doe2024")?
///     .title("Example & Practice")?
///     .year("2024")?
///     .person(BibliographyPerson::new("Doe")?.with_first("Jane")?)?;
/// package.add_bibliography_source(source)?;
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
#[derive(Debug, Clone, Default)]
pub struct BibliographySourceBuilder {
    tag: String,
    kind: BibliographySourceKind,
    guid: Option<String>,
    lcid: Option<String>,
    title: Option<String>,
    year: Option<String>,
    persons: Vec<BibliographyPerson>,
    corporate_author: Option<String>,
    publisher: Option<String>,
    journal_name: Option<String>,
    volume: Option<String>,
    issue: Option<String>,
    pages: Option<String>,
    url: Option<String>,
    city: Option<String>,
    medium: Option<String>,
    edition: Option<String>,
    comments: Option<String>,
    custom_fields: Vec<(Vec<String>, String)>,
}

/// Consuming-builder field setters; each validates the field bounds.
macro_rules! optional_field {
    ($name:ident, $element:literal) => {
        #[doc = concat!("Set the `", $element, "` field.")]
        pub fn $name(mut self, value: impl Into<String>) -> Result<Self> {
            let value = value.into();
            validate_field($element, &value)?;
            self.$name = Some(value);
            Ok(self)
        }
    };
}

impl BibliographySourceBuilder {
    /// Start a source with a kind and a unique citation tag (`b:Tag`).
    pub fn new(kind: BibliographySourceKind, tag: impl Into<String>) -> Result<Self> {
        let tag = tag.into();
        if tag.is_empty() || tag.len() > MAX_TAG_LENGTH {
            return Err(invalid(format!(
                "bibliography tag must be 1..={MAX_TAG_LENGTH} bytes"
            )));
        }
        Ok(Self {
            tag,
            kind,
            ..Self::default()
        })
    }

    optional_field!(guid, "Guid");
    optional_field!(lcid, "LCID");
    optional_field!(title, "Title");
    optional_field!(year, "Year");
    optional_field!(publisher, "Publisher");
    optional_field!(journal_name, "JournalName");
    optional_field!(volume, "Volume");
    optional_field!(issue, "Issue");
    optional_field!(pages, "Pages");
    optional_field!(url, "URL");
    optional_field!(city, "City");
    optional_field!(medium, "Medium");
    optional_field!(edition, "Edition");
    optional_field!(comments, "Comments");

    /// Add a person to the author name list.
    pub fn person(mut self, person: BibliographyPerson) -> Result<Self> {
        if self.persons.len() >= MAX_PERSONS {
            return Err(invalid(format!(
                "bibliography person count exceeds {MAX_PERSONS}"
            )));
        }
        self.persons.push(person);
        Ok(self)
    }

    /// Set a corporate (group) author instead of a person name list.
    pub fn corporate_author(mut self, corporate: impl Into<String>) -> Result<Self> {
        let corporate = corporate.into();
        validate_field("Corporate", &corporate)?;
        self.corporate_author = Some(corporate);
        Ok(self)
    }

    /// Add an opaque custom field as a nested element path with a scalar
    /// value, for source data the typed API does not model.
    ///
    /// For example, `custom_field(vec!["BookTitle"], "…")` emits
    /// `<b:BookTitle>…</b:BookTitle>`. Path segments must be valid XML names
    /// in the bibliography namespace.
    pub fn custom_field(mut self, path: Vec<String>, value: String) -> Result<Self> {
        if self.custom_fields.len() >= MAX_CUSTOM_FIELDS {
            return Err(invalid(format!(
                "bibliography custom field count exceeds {MAX_CUSTOM_FIELDS}"
            )));
        }
        if path.is_empty() || path.len() > MAX_CUSTOM_PATH_DEPTH {
            return Err(invalid(format!(
                "bibliography custom field path depth must be 1..={MAX_CUSTOM_PATH_DEPTH}"
            )));
        }
        for segment in &path {
            validate_name(segment)?;
        }
        validate_field("custom field", &value)?;
        self.custom_fields.push((path, value));
        Ok(self)
    }

    /// The citation tag of this source.
    pub fn tag(&self) -> &str {
        &self.tag
    }

    /// The source kind of this source.
    pub fn kind(&self) -> BibliographySourceKind {
        self.kind
    }

    /// Build the `b:Source` element tree in deterministic field order.
    fn to_node(&self) -> XmlNode {
        let mut source = bibliography_element("Source");
        push_scalar(&mut source, "Tag", &self.tag);
        push_scalar(&mut source, "SourceType", self.kind.as_token());
        for (element, value) in [
            ("Guid", &self.guid),
            ("LCID", &self.lcid),
            ("Title", &self.title),
            ("Year", &self.year),
        ] {
            if let Some(value) = value {
                push_scalar(&mut source, element, value);
            }
        }
        if !self.persons.is_empty() || self.corporate_author.is_some() {
            let mut outer = bibliography_element("Author");
            let mut inner = bibliography_element("Author");
            if let Some(corporate) = &self.corporate_author {
                push_scalar(&mut inner, "Corporate", corporate);
            }
            if !self.persons.is_empty() {
                let mut name_list = bibliography_element("NameList");
                for person in &self.persons {
                    let mut node = bibliography_element("Person");
                    push_scalar(&mut node, "Last", &person.last);
                    if let Some(first) = &person.first {
                        push_scalar(&mut node, "First", first);
                    }
                    if let Some(middle) = &person.middle {
                        push_scalar(&mut node, "Middle", middle);
                    }
                    name_list.children.push(node);
                }
                inner.children.push(name_list);
            }
            outer.children.push(inner);
            source.children.push(outer);
        }
        for (element, value) in [
            ("City", &self.city),
            ("Publisher", &self.publisher),
            ("JournalName", &self.journal_name),
            ("Volume", &self.volume),
            ("Issue", &self.issue),
            ("Pages", &self.pages),
            ("Edition", &self.edition),
            ("Medium", &self.medium),
            ("URL", &self.url),
            ("Comments", &self.comments),
        ] {
            if let Some(value) = value {
                push_scalar(&mut source, element, value);
            }
        }
        for (path, value) in &self.custom_fields {
            push_nested(&mut source, path, value);
        }
        source
    }
}

fn bibliography_element(local_name: &str) -> XmlNode {
    XmlNode {
        namespace: Some(OOXML_BIBLIOGRAPHY_NAMESPACE.to_string()),
        local_name: local_name.to_string(),
        attributes: Vec::new(),
        text: String::new(),
        children: Vec::new(),
    }
}

fn push_scalar(parent: &mut XmlNode, element: &str, value: &str) {
    let mut node = bibliography_element(element);
    node.text = value.to_string();
    parent.children.push(node);
}

fn push_nested(parent: &mut XmlNode, path: &[String], value: &str) {
    let mut node = bibliography_element(&path[0]);
    if path.len() == 1 {
        node.text = value.to_string();
    } else {
        push_nested(&mut node, &path[1..], value);
    }
    parent.children.push(node);
}

fn validate_field(element: &str, value: &str) -> Result<()> {
    if value.len() > MAX_FIELD_BYTES {
        return Err(invalid(format!(
            "bibliography {element} value exceeds {MAX_FIELD_BYTES} bytes"
        )));
    }
    Ok(())
}

/// Validate an XML element name for custom field paths.
fn validate_name(segment: &str) -> Result<()> {
    let mut chars = segment.chars();
    let valid = !segment.is_empty()
        && chars
            .next()
            .is_some_and(|first| first.is_ascii_alphabetic() || first == '_')
        && chars.all(|c| c.is_ascii_alphanumeric() || matches!(c, '.' | '-' | '_'));
    if !valid {
        return Err(invalid(format!(
            "invalid bibliography custom field name '{segment}'"
        )));
    }
    Ok(())
}

/// Generate a fresh `b:Sources` store document with the given sources.
pub(crate) fn new_store_xml(sources: &[BibliographySourceBuilder]) -> Result<String> {
    let mut root = bibliography_element("Sources");
    for source in sources {
        root.children.push(source.to_node());
    }
    serialize_store(&root)
}

/// Append a source to an existing store document, rejecting duplicate tags.
pub(crate) fn add_source_xml(
    store_xml: &[u8],
    source: &BibliographySourceBuilder,
) -> Result<String> {
    let mut root = parse_xml_tree(store_xml)?;
    if store_tags(&root)
        .iter()
        .any(|tag| tag.as_str() == source.tag())
    {
        return Err(invalid(format!(
            "bibliography tag '{}' already exists in the source store",
            source.tag()
        )));
    }
    match root.local_name.as_str() {
        "Sources" => {
            if count_sources(&root) >= MAX_STORE_SOURCES {
                return Err(invalid("bibliography source count limit exceeded"));
            }
            root.children.push(source.to_node());
        },
        "Source" => {
            // Wrap a legacy single-source payload into a source list.
            let old = root;
            let mut wrapped = bibliography_element("Sources");
            wrapped.namespace = old.namespace.clone();
            wrapped.children.push(old);
            wrapped.children.push(source.to_node());
            root = wrapped;
        },
        _ => return Err(invalid("bibliography store root is not Sources or Source")),
    }
    serialize_store(&root)
}

/// Remove the source with the given tag; returns the new store XML and
/// whether a source was removed.
pub(crate) fn remove_source_xml(store_xml: &[u8], tag: &str) -> Result<(String, bool)> {
    let mut root = parse_xml_tree(store_xml)?;
    if !is_bibliography_root(
        root.namespace.as_deref().unwrap_or_default(),
        &root.local_name,
    ) {
        return Err(invalid("bibliography store root is not Sources or Source"));
    }
    if root.local_name == "Source" {
        if source_tag(&root).as_deref() == Some(tag) {
            // Removing the only source yields an empty source list.
            let mut wrapped = bibliography_element("Sources");
            wrapped.namespace = root.namespace.clone();
            return serialize_store(&wrapped).map(|xml| (xml, true));
        }
        return serialize_store(&root).map(|xml| (xml, false));
    }
    let before = root.children.len();
    root.children.retain(|child| {
        !(is_bibliography_node(child)
            && child.local_name == "Source"
            && source_tag(child).as_deref() == Some(tag))
    });
    let removed = root.children.len() != before;
    serialize_store(&root).map(|xml| (xml, removed))
}

/// Replace the source with the given tag, preserving entry order.
pub(crate) fn replace_source_xml(
    store_xml: &[u8],
    tag: &str,
    source: &BibliographySourceBuilder,
) -> Result<String> {
    let mut root = parse_xml_tree(store_xml)?;
    if root.local_name != "Sources" {
        if root.local_name == "Source" && source_tag(&root).as_deref() == Some(tag) {
            let replacement = source.to_node();
            root.children = replacement.children;
            root.attributes = replacement.attributes;
            return serialize_store(&root);
        }
        return Err(invalid(format!(
            "bibliography tag '{tag}' does not exist in the source store"
        )));
    }
    for child in &mut root.children {
        if is_bibliography_node(child)
            && child.local_name == "Source"
            && source_tag(child).as_deref() == Some(tag)
        {
            *child = source.to_node();
            return serialize_store(&root);
        }
    }
    Err(invalid(format!(
        "bibliography tag '{tag}' does not exist in the source store"
    )))
}

/// Return the `b:Tag` text of a `b:Source` element, if present.
fn source_tag(source: &XmlNode) -> Option<String> {
    source
        .children
        .iter()
        .find(|child| is_bibliography_node(child) && child.local_name == "Tag")
        .map(|child| child.text.clone())
}

/// Collect all tags in a store root for uniqueness checks.
fn store_tags(root: &XmlNode) -> Vec<String> {
    let sources: &[XmlNode] = if root.local_name == "Sources" {
        &root.children
    } else {
        std::slice::from_ref(root)
    };
    sources
        .iter()
        .filter(|source| is_bibliography_node(source) && source.local_name == "Source")
        .filter_map(source_tag)
        .collect()
}

fn count_sources(root: &XmlNode) -> usize {
    root.children
        .iter()
        .filter(|child| is_bibliography_node(child) && child.local_name == "Source")
        .count()
}

/// Serialize a store tree back to XML.
///
/// Bibliography-namespace elements are emitted with the `b:` prefix bound on
/// the root; elements from foreign namespaces carry an inline `xmlns`
/// declaration. Root attributes (`SelectedStyle`, `StyleName`) are preserved.
fn serialize_store(root: &XmlNode) -> Result<String> {
    let root_namespace = root
        .namespace
        .clone()
        .unwrap_or_else(|| OOXML_BIBLIOGRAPHY_NAMESPACE.to_string());
    let mut xml = String::with_capacity(1024);
    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    write_element(&mut xml, root, &root_namespace, true)?;
    Ok(xml)
}

fn write_element(
    xml: &mut String,
    node: &XmlNode,
    store_namespace: &str,
    is_root: bool,
) -> Result<()> {
    let in_store = node.namespace.as_deref() == Some(store_namespace);
    let name = if in_store {
        format!("b:{}", node.local_name)
    } else {
        node.local_name.clone()
    };
    write!(xml, "<{name}").map_err(|error| Error::Xml(error.to_string()))?;
    if is_root {
        write!(xml, r#" xmlns:b="{store_namespace}""#)
            .map_err(|error| Error::Xml(error.to_string()))?;
    }
    if !in_store
        && !is_root
        && let Some(namespace) = &node.namespace
    {
        write!(xml, r#" xmlns="{}""#, escape_xml(namespace))
            .map_err(|error| Error::Xml(error.to_string()))?;
    }
    for attribute in &node.attributes {
        write!(
            xml,
            r#" {}="{}""#,
            attribute.local_name,
            escape_xml(&attribute.value)
        )
        .map_err(|error| Error::Xml(error.to_string()))?;
    }
    if node.children.is_empty() && node.text.is_empty() {
        xml.push_str("/>");
        return Ok(());
    }
    xml.push('>');
    write!(xml, "{}", escape_xml(&node.text)).map_err(|error| Error::Xml(error.to_string()))?;
    for child in &node.children {
        write_element(xml, child, store_namespace, false)?;
    }
    write!(xml, "</{name}>").map_err(|error| Error::Xml(error.to_string()))?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::bibliography::discover_bibliography_source_stores;
    use litchi_ooxml_common::custom_xml::{
        Conformance, Item, NewItem, NewProps, Props, add, discover,
    };
    use litchi_opc::constants::content_type as ct;
    use litchi_opc::part::BlobPart;
    use litchi_opc::{OpcPackage, PackURI};

    fn item(xml: &[u8], namespace: &str, local_name: &str) -> Item {
        let mut package = OpcPackage::new();
        let source = PackURI::new("/word/document.xml").unwrap();
        package.add_part(Box::new(BlobPart::new(
            source.clone(),
            ct::WML_DOCUMENT_MAIN.to_string(),
            Vec::new(),
        )));
        add(
            &mut package,
            NewItem {
                source,
                rel_id: "rIdBib".to_string(),
                part: PackURI::new("/customXml/item1.xml").unwrap(),
                content_type: "application/xml".to_string(),
                xml: xml.to_vec(),
                props: Some(NewProps {
                    part: PackURI::new("/customXml/itemProps1.xml").unwrap(),
                    rel_id: "rIdProps".to_string(),
                    value: Props {
                        id: "{11111111-1111-1111-1111-111111111111}".to_string(),
                        schemas: vec![OOXML_BIBLIOGRAPHY_NAMESPACE.to_string()],
                    },
                }),
                conformance: Conformance::Transitional,
            },
        )
        .unwrap();
        let item = discover(&package).unwrap().remove(0);
        assert_eq!(item.root().namespace, namespace);
        assert_eq!(item.root().local_name, local_name);
        item
    }

    fn built() -> BibliographySourceBuilder {
        BibliographySourceBuilder::new(BibliographySourceKind::Book, "Doe2024")
            .unwrap()
            .title("Example & Practice")
            .unwrap()
            .year("2024")
            .unwrap()
            .publisher("Contoso")
            .unwrap()
            .city("London")
            .unwrap()
            .edition("3")
            .unwrap()
            .medium("Print")
            .unwrap()
            .comments("Annotated")
            .unwrap()
            .person(
                BibliographyPerson::new("Doe")
                    .unwrap()
                    .with_first("Jane")
                    .unwrap()
                    .with_middle("Q")
                    .unwrap(),
            )
            .unwrap()
            .custom_field(vec!["BookTitle".to_string()], "Sub Title".to_string())
            .unwrap()
    }

    #[test]
    fn serializes_typed_source_with_all_fields() {
        let xml = new_store_xml(&[built()]).unwrap();
        for expected in [
            "<b:Tag>Doe2024</b:Tag>",
            "<b:SourceType>Book</b:SourceType>",
            "<b:Title>Example &amp; Practice</b:Title>",
            "<b:Year>2024</b:Year>",
            "<b:Publisher>Contoso</b:Publisher>",
            "<b:City>London</b:City>",
            "<b:Edition>3</b:Edition>",
            "<b:Medium>Print</b:Medium>",
            "<b:Comments>Annotated</b:Comments>",
            "<b:Last>Doe</b:Last><b:First>Jane</b:First><b:Middle>Q</b:Middle>",
            "<b:BookTitle>Sub Title</b:BookTitle>",
        ] {
            assert!(xml.contains(expected), "missing {expected} in {xml}");
        }

        // The read side recovers every typed field.
        let stores = discover_bibliography_source_stores(&[item(
            xml.as_bytes(),
            OOXML_BIBLIOGRAPHY_NAMESPACE,
            "Sources",
        )])
        .unwrap();
        let source = &stores[0].sources()[0];
        assert_eq!(source.tag(), Some("Doe2024"));
        assert_eq!(source.source_type(), Some("Book"));
        assert_eq!(source.title(), Some("Example & Practice"));
        assert_eq!(source.year(), Some("2024"));
        assert_eq!(
            source.value(&["Author", "Author", "NameList", "Person", "Last"]),
            Some("Doe")
        );
        assert_eq!(
            source.value(&["Author", "Author", "NameList", "Person", "Middle"]),
            Some("Q")
        );
        assert_eq!(source.value(&["Publisher"]), Some("Contoso"));
        assert_eq!(source.value(&["BookTitle"]), Some("Sub Title"));
    }

    #[test]
    fn validates_builder_bounds() {
        assert!(BibliographySourceBuilder::new(BibliographySourceKind::Book, "").is_err());
        assert!(
            BibliographySourceBuilder::new(BibliographySourceKind::Book, "x".repeat(256)).is_err()
        );
        assert!(BibliographyPerson::new("").is_err());
        assert!(
            BibliographySourceBuilder::new(BibliographySourceKind::Book, "T")
                .unwrap()
                .title("x".repeat(MAX_FIELD_BYTES + 1))
                .is_err()
        );
        assert!(
            BibliographySourceBuilder::new(BibliographySourceKind::Book, "T")
                .unwrap()
                .custom_field(vec!["1bad".to_string()], "v".to_string())
                .is_err()
        );
        assert!(
            BibliographySourceBuilder::new(BibliographySourceKind::Book, "T")
                .unwrap()
                .custom_field(Vec::new(), "v".to_string())
                .is_err()
        );
    }

    #[test]
    fn add_remove_replace_preserve_store_content() {
        let initial = new_store_xml(&[built()]).unwrap();
        let second =
            BibliographySourceBuilder::new(BibliographySourceKind::JournalArticle, "Smith2025")
                .unwrap()
                .journal_name("Journal of Tests")
                .unwrap()
                .volume("12")
                .unwrap()
                .issue("3")
                .unwrap()
                .pages("10-20")
                .unwrap()
                .url("https://example.invalid/article")
                .unwrap()
                .corporate_author("Research Group")
                .unwrap();

        // Add.
        let added = add_source_xml(initial.as_bytes(), &second).unwrap();
        assert!(added.contains("<b:Tag>Smith2025</b:Tag>"));
        assert!(added.contains("<b:Corporate>Research Group</b:Corporate>"));
        // Duplicate tags are rejected.
        assert!(add_source_xml(added.as_bytes(), &second).is_err());

        // Replace preserves order and the untouched entry.
        let replacement = BibliographySourceBuilder::new(BibliographySourceKind::Report, "Doe2024")
            .unwrap()
            .title("Replaced")
            .unwrap();
        let replaced = replace_source_xml(added.as_bytes(), "Doe2024", &replacement).unwrap();
        assert!(replaced.contains("<b:SourceType>Report</b:SourceType>"));
        assert!(replaced.contains("<b:Tag>Smith2025</b:Tag>"));
        assert!(!replaced.contains("Example &amp; Practice"));
        assert!(replace_source_xml(replaced.as_bytes(), "Missing", &replacement).is_err());

        // Remove.
        let (removed, changed) = remove_source_xml(replaced.as_bytes(), "Smith2025").unwrap();
        assert!(changed);
        assert!(!removed.contains("Smith2025"));
        let (_, changed) = remove_source_xml(removed.as_bytes(), "Missing").unwrap();
        assert!(!changed);
    }

    #[test]
    fn preserves_style_attributes_on_mutation() {
        let store = br#"<b:Sources xmlns:b="http://schemas.openxmlformats.org/officeDocument/2006/bibliography" SelectedStyle="/APA.XSL" StyleName="APA"><b:Source><b:Tag>Kept</b:Tag></b:Source></b:Sources>"#;
        let added = add_source_xml(
            store,
            &BibliographySourceBuilder::new(BibliographySourceKind::Misc, "New").unwrap(),
        )
        .unwrap();
        assert!(added.contains(r#"SelectedStyle="/APA.XSL""#));
        assert!(added.contains(r#"StyleName="APA""#));
        assert!(added.contains("<b:Tag>Kept</b:Tag>"));
    }

    #[test]
    fn round_trips_authored_sources_through_saved_package() {
        use crate::Package;
        use tempfile::NamedTempFile;

        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        let item_id = package.add_bibliography_source(built()).unwrap();
        package
            .add_bibliography_source(
                BibliographySourceBuilder::new(BibliographySourceKind::JournalArticle, "Smith2025")
                    .unwrap()
                    .title("Testing in Practice")
                    .unwrap()
                    .journal_name("Journal of Tests")
                    .unwrap()
                    .year("2025")
                    .unwrap()
                    .volume("12")
                    .unwrap()
                    .issue("3")
                    .unwrap()
                    .pages("10-20")
                    .unwrap()
                    .url("https://example.invalid/article")
                    .unwrap()
                    .corporate_author("Research Group")
                    .unwrap(),
            )
            .unwrap();
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let stores = reopened.bibliography_source_stores().unwrap();
        assert_eq!(stores.len(), 1);
        assert_eq!(stores[0].data_store_item_id(), Some(item_id.as_str()));
        assert_eq!(
            stores[0].schema_references(),
            [OOXML_BIBLIOGRAPHY_NAMESPACE]
        );
        assert_eq!(stores[0].source_count(), 2);

        let first = &stores[0].sources()[0];
        assert_eq!(first.tag(), Some("Doe2024"));
        assert_eq!(first.source_type(), Some("Book"));
        assert_eq!(first.title(), Some("Example & Practice"));
        assert_eq!(first.year(), Some("2024"));
        assert_eq!(first.value(&["Publisher"]), Some("Contoso"));
        assert_eq!(first.value(&["City"]), Some("London"));
        assert_eq!(first.value(&["Edition"]), Some("3"));
        assert_eq!(first.value(&["Medium"]), Some("Print"));
        assert_eq!(first.value(&["Comments"]), Some("Annotated"));
        assert_eq!(
            first.value(&["Author", "Author", "NameList", "Person", "Last"]),
            Some("Doe")
        );
        assert_eq!(first.value(&["BookTitle"]), Some("Sub Title"));

        let second = &stores[0].sources()[1];
        assert_eq!(second.tag(), Some("Smith2025"));
        assert_eq!(second.source_type(), Some("JournalArticle"));
        assert_eq!(second.value(&["JournalName"]), Some("Journal of Tests"));
        assert_eq!(second.value(&["Volume"]), Some("12"));
        assert_eq!(second.value(&["Issue"]), Some("3"));
        assert_eq!(second.value(&["Pages"]), Some("10-20"));
        assert_eq!(
            second.value(&["URL"]),
            Some("https://example.invalid/article")
        );
        assert_eq!(
            second.value(&["Author", "Author", "Corporate"]),
            Some("Research Group")
        );

        let flattened = reopened.bibliography_sources().unwrap();
        assert_eq!(reopened.bibliography_source_count().unwrap(), 2);
        assert_eq!(flattened.len(), 2);
    }

    #[test]
    fn remove_replace_and_duplicate_tag_flows() {
        use crate::Package;
        use tempfile::NamedTempFile;

        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        package.add_bibliography_source(built()).unwrap();
        // Duplicate tags are rejected before any mutation.
        assert!(package.add_bibliography_source(built()).is_err());

        package
            .replace_bibliography_source(
                "Doe2024",
                BibliographySourceBuilder::new(BibliographySourceKind::Report, "Doe2024")
                    .unwrap()
                    .title("Replaced Title")
                    .unwrap(),
            )
            .unwrap();
        assert!(
            package
                .replace_bibliography_source(
                    "Missing",
                    BibliographySourceBuilder::new(BibliographySourceKind::Report, "Missing")
                        .unwrap(),
                )
                .is_err()
        );

        assert!(package.remove_bibliography_source("Doe2024").unwrap());
        assert!(!package.remove_bibliography_source("Doe2024").unwrap());
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let stores = reopened.bibliography_source_stores().unwrap();
        assert_eq!(stores.len(), 1);
        assert_eq!(stores[0].source_count(), 0);
    }

    #[test]
    fn mutates_real_fixture_store_preserving_style_metadata() {
        use crate::Package;
        use tempfile::NamedTempFile;

        const FIXTURE: &[u8] = include_bytes!("../../../../test-data/ooxml/docx/footnotes.docx");

        let original = Package::from_reader(std::io::Cursor::new(FIXTURE)).unwrap();
        let stores = original.bibliography_source_stores().unwrap();
        assert_eq!(stores.len(), 1);
        assert_eq!(stores[0].selected_style(), Some("\\APA.XSL"));
        assert_eq!(stores[0].style_name(), Some("APA"));
        assert_eq!(stores[0].source_count(), 0);

        let file = NamedTempFile::with_suffix(".docx").unwrap();
        std::fs::write(file.path(), FIXTURE).unwrap();
        let mut package = Package::open(file.path()).unwrap();
        package
            .add_bibliography_source(
                BibliographySourceBuilder::new(BibliographySourceKind::Book, "Fixture2026")
                    .unwrap()
                    .title("Added to fixture")
                    .unwrap(),
            )
            .unwrap();
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let stores = reopened.bibliography_source_stores().unwrap();
        assert_eq!(stores.len(), 1);
        // The fixture's style metadata and footnotes survive the mutation.
        assert_eq!(stores[0].selected_style(), Some("\\APA.XSL"));
        assert_eq!(stores[0].style_name(), Some("APA"));
        assert_eq!(stores[0].source_count(), 1);
        let source = &stores[0].sources()[0];
        assert_eq!(source.tag(), Some("Fixture2026"));
        assert_eq!(source.title(), Some("Added to fixture"));
        assert!(reopened.document().unwrap().text().is_ok());
    }
}
