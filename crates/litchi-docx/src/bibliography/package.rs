//! Compatibility adapter for the canonical DOCX bibliography source codec.
//!
//! `crate::bibliography` owns the bounded namespace/model/XML parser.
//! This host module keeps the historical API and the package-specific seam:
//! Custom XML item discovery, relationship provenance, part identifiers, and
//! properties metadata remain attached to the host-side store value.  The
//! module does not resolve styles, match citations, or access external data.

use crate::error::Result;
use litchi_ooxml_common::custom_xml::Item;
use litchi_opc::PackURI;

use crate::bibliography as owner;

use owner::BibliographySource;

// The historical bibliography writer is in this host crate and mutates the
// parsed tree for source CRUD.  These are compatibility names only; parsing
// and the tree representation are implemented by the owner crate.
type OwnerSourceStore = owner::BibliographySourceStore;

/// One inert Word bibliography source store discovered in Custom XML.
///
/// The XML semantics are owned by `litchi_docx`; these fields retain the
/// historical host-side Custom XML relationship and package provenance.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SourceStore {
    source_part_name: PackURI,
    relationship_id: String,
    data_part_name: PackURI,
    content_type: String,
    properties_part_name: Option<PackURI>,
    data_store_item_id: Option<String>,
    schema_references: Vec<String>,
    payload: OwnerSourceStore,
}

impl SourceStore {
    /// Return the part that owns the Custom XML relationship.
    #[must_use]
    pub fn source_part_name(&self) -> &PackURI {
        &self.source_part_name
    }

    /// Return the owning Custom XML relationship ID.
    #[must_use]
    pub fn relationship_id(&self) -> &str {
        &self.relationship_id
    }

    /// Return the Custom XML data-part name.
    #[must_use]
    pub fn data_part_name(&self) -> &PackURI {
        &self.data_part_name
    }

    /// Return the stored Custom XML content type.
    #[must_use]
    pub fn content_type(&self) -> &str {
        &self.content_type
    }

    /// Return the optional Custom XML properties-part name.
    #[must_use]
    pub fn properties_part_name(&self) -> Option<&PackURI> {
        self.properties_part_name.as_ref()
    }

    /// Return the optional Custom XML data-store GUID.
    #[must_use]
    pub fn data_store_item_id(&self) -> Option<&str> {
        self.data_store_item_id.as_deref()
    }

    /// Return declared schema-reference URIs without resolving them.
    #[must_use]
    pub fn schema_references(&self) -> &[String] {
        &self.schema_references
    }

    /// Return the stored selected bibliography style reference, if any.
    ///
    /// This is opaque metadata. The library never opens, loads, or executes
    /// the referenced style.
    #[must_use]
    pub fn selected_style(&self) -> Option<&str> {
        self.payload.selected_style()
    }

    /// Return the stored bibliography style name, if any.
    #[must_use]
    pub fn style_name(&self) -> Option<&str> {
        self.payload.style_name()
    }

    /// Return sources in their persisted XML order.
    #[must_use]
    pub fn sources(&self) -> &[BibliographySource] {
        self.payload.sources()
    }

    /// Return the number of persisted bibliography sources.
    #[must_use]
    pub fn source_count(&self) -> usize {
        self.payload.source_count()
    }
}

/// Discover bibliography source stores from validated Custom XML items.
pub(crate) fn discover_bibliography_source_stores(items: &[Item]) -> Result<Vec<SourceStore>> {
    let mut stores = Vec::new();
    for item in items {
        if !owner::is_bibliography_root(&item.root().namespace, &item.root().local_name) {
            continue;
        }

        let payload = owner::parse_bibliography_source_store(item.xml())?;
        stores.push(SourceStore {
            source_part_name: item.source().clone(),
            relationship_id: item.rel_id().to_string(),
            data_part_name: item.part().clone(),
            content_type: item.content_type().to_string(),
            properties_part_name: item.props_part().cloned(),
            data_store_item_id: item.props().map(|props| props.id.clone()),
            schema_references: item
                .props()
                .map(|props| props.schemas.clone())
                .unwrap_or_default(),
            payload,
        });
    }
    Ok(stores)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::bibliography::{
        LEGACY_WORD_BIBLIOGRAPHY_NAMESPACE, OOXML_BIBLIOGRAPHY_NAMESPACE,
        STRICT_OOXML_BIBLIOGRAPHY_NAMESPACE,
    };
    use litchi_ooxml_common::custom_xml::{Conformance, NewItem, NewProps, Props, add, discover};
    use litchi_opc::OpcPackage;
    use litchi_opc::constants::content_type as ct;
    use litchi_opc::part::BlobPart;

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

    #[test]
    fn preserves_custom_xml_provenance_around_owner_payload() {
        let xml = br#"<b:Sources xmlns:b="http://schemas.openxmlformats.org/officeDocument/2006/bibliography" SelectedStyle="/APA.XSL" StyleName="APA"><b:Source><b:Tag>Doe2024</b:Tag><b:Title>Example &amp; Practice</b:Title></b:Source></b:Sources>"#;
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
        assert_eq!(store.data_part_name().as_str(), "/customXml/item1.xml");
        assert_eq!(store.content_type(), "application/xml");
        assert_eq!(
            store.properties_part_name().map(PackURI::as_str),
            Some("/customXml/itemProps1.xml")
        );
        assert_eq!(
            store.data_store_item_id(),
            Some("{11111111-1111-1111-1111-111111111111}")
        );
        assert_eq!(store.schema_references(), [OOXML_BIBLIOGRAPHY_NAMESPACE]);
        assert_eq!(store.selected_style(), Some("/APA.XSL"));
        assert_eq!(store.style_name(), Some("APA"));
        assert_eq!(store.sources()[0].tag(), Some("Doe2024"));
        assert_eq!(store.sources()[0].title(), Some("Example & Practice"));
    }

    #[test]
    fn recognizes_legacy_single_source_payloads() {
        let xml = format!(
            "<b:Source xmlns:b=\"{LEGACY_WORD_BIBLIOGRAPHY_NAMESPACE}\"><b:Tag>Legacy</b:Tag><b:Title>Stored only</b:Title></b:Source>"
        );
        let stores = discover_bibliography_source_stores(&[item(
            xml.as_bytes(),
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
        let xml = format!(
            "<b:Sources xmlns:b=\"{STRICT_OOXML_BIBLIOGRAPHY_NAMESPACE}\"><b:Source><b:Tag>Strict</b:Tag></b:Source></b:Sources>"
        );
        let stores = discover_bibliography_source_stores(&[item(
            xml.as_bytes(),
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
