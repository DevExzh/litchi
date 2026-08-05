//! Custom XML data stores, bibliography sources, and SDT bindings.

use super::super::model::*;

impl Package {
    /// Discover every validated Custom XML Data Storage relationship occurrence.
    pub fn custom_xml(&self) -> Result<Vec<CustomXmlItem>> {
        Ok(custom_xml::discover(&self.opc)?)
    }

    /// Discover typed, inert bibliography source stores from Custom XML.
    ///
    /// Word stores its current bibliography source list in a document Custom
    /// XML data store. This method exposes stored source values and style
    /// metadata only. It never matches source tags to citations, resolves
    /// schemas or styles, runs transforms, refreshes fields, or changes data.
    pub fn bibliography_source_stores(&self) -> Result<Vec<SourceStore>> {
        let items = custom_xml::discover(&self.opc)?;
        discover_bibliography_source_stores(&items)
    }

    /// Discover typed, inert bibliography sources in package and XML order.
    ///
    /// This flattens [`Self::bibliography_source_stores`] without resolving
    /// `CITATION` fields or applying bibliography style rules.
    pub fn bibliography_sources(&self) -> Result<Vec<BibliographySource>> {
        let stores = self.bibliography_source_stores()?;
        Ok(stores
            .iter()
            .flat_map(|store| store.sources().iter().cloned())
            .collect())
    }

    /// Return the number of typed, inert bibliography sources.
    pub fn bibliography_source_count(&self) -> Result<usize> {
        Ok(self.bibliography_sources()?.len())
    }

    /// Find a Custom XML data store by its case-insensitive datastore item GUID.
    pub fn custom_xml_by_id(&self, id: &str) -> Result<Option<CustomXmlItem>> {
        Ok(custom_xml::discover(&self.opc)?.into_iter().find(|item| {
            item.props()
                .is_some_and(|props| props.id.eq_ignore_ascii_case(id))
        }))
    }

    /// Add a collision-safe `/customXml/itemN.xml` data store to the main document.
    pub fn add_custom_xml(&mut self, store: NewStore) -> Result<CustomXmlItem> {
        custom_xml::validate_content_type(&store.content_type)?;
        custom_xml::validate_payload(&store.xml)?;
        let props = CustomXmlProps {
            id: store.id,
            schemas: store.schemas,
        };
        custom_xml::validate_props(&props)?;
        let source_part = self.opc.main_document_part()?.partname().clone();
        let source = self.opc.get_part(&source_part)?;
        let rel_id = (1usize..=MAX_ITEMS + 1)
            .map(|number| format!("rIdCustomXml{number}"))
            .find(|id| source.rels().get(id).is_none())
            .ok_or_else(|| {
                Error::InvalidFormat("Custom XML relationship ID space is exhausted".into())
            })?;
        let mut part_names = None;
        for number in 1usize..=MAX_ITEMS + 1 {
            let data =
                PackURI::new(format!("/customXml/item{number}.xml")).map_err(Error::InvalidUri)?;
            let props = PackURI::new(format!("/customXml/itemProps{number}.xml"))
                .map_err(Error::InvalidUri)?;
            let conflict = self.opc.iter_parts().any(|part| {
                part.partname().as_str().eq_ignore_ascii_case(data.as_str())
                    || part
                        .partname()
                        .as_str()
                        .eq_ignore_ascii_case(props.as_str())
            });
            if !conflict {
                part_names = Some((data, props));
                break;
            }
        }
        let (data_part, props_part) = part_names.ok_or_else(|| {
            Error::InvalidFormat("Custom XML part-name space is exhausted".into())
        })?;
        custom_xml::add(
            &mut self.opc,
            NewCustomXmlItem {
                source: source_part,
                rel_id,
                part: data_part.clone(),
                content_type: store.content_type,
                xml: store.xml,
                props: Some(NewCustomXmlProps {
                    part: props_part,
                    rel_id: "rIdProps1".to_string(),
                    value: props,
                }),
                conformance: store.conformance,
            },
        )?;
        custom_xml::discover(&self.opc)?
            .into_iter()
            .find(|item| item.part() == &data_part)
            .ok_or_else(|| {
                Error::InvalidFormat("new Custom XML data store was not discoverable".into())
            })
    }

    /// Replace only the inert XML payload of a data store.
    pub fn set_custom_xml(&mut self, id: &str, xml: Vec<u8>) -> Result<()> {
        custom_xml::validate_payload(&xml)?;
        let item = self
            .custom_xml_by_id(id)?
            .ok_or_else(|| Error::PartNotFound(format!("Custom XML itemID '{id}'")))?;
        self.opc.get_part_mut(item.part())?.set_blob(xml);
        self.opc.unsign();
        Ok(())
    }

    /// Replace payload, content type, schema references, and canonical properties.
    pub fn replace_custom_xml(&mut self, id: &str, replacement: NewStore) -> Result<()> {
        custom_xml::validate_content_type(&replacement.content_type)?;
        custom_xml::validate_payload(&replacement.xml)?;
        if !replacement.id.eq_ignore_ascii_case(id) {
            return Err(Error::InvalidFormat(
                "replacement itemID must identify the existing data store".into(),
            ));
        }
        let props = CustomXmlProps {
            id: replacement.id,
            schemas: replacement.schemas,
        };
        let props_xml = custom_xml::write_props(&props, replacement.conformance)?;
        let item = self
            .custom_xml_by_id(id)?
            .ok_or_else(|| Error::PartNotFound(format!("Custom XML itemID '{id}'")))?;
        let props_part = item.props_part().cloned().ok_or_else(|| {
            Error::InvalidFormat("Custom XML data store has no properties part".into())
        })?;
        let existing_relationships = self
            .opc
            .get_part(item.part())?
            .rels()
            .iter()
            .map(|relationship| {
                (
                    relationship.reltype().to_string(),
                    relationship.target_ref().to_string(),
                    relationship.r_id().to_string(),
                    relationship.is_external(),
                )
            })
            .collect::<Vec<_>>();
        let mut data_part = BlobPart::new(
            item.part().clone(),
            replacement.content_type,
            replacement.xml,
        );
        for (reltype, target, id, external) in existing_relationships {
            data_part
                .rels_mut()
                .add_relationship(reltype, target, id, external);
        }
        self.opc.add_part(Box::new(data_part));
        self.opc.get_part_mut(&props_part)?.set_blob(props_xml);
        self.opc.unsign();
        Ok(())
    }

    /// Remove a data store unless an SDT still binds to its item GUID.
    pub fn remove_custom_xml(&mut self, id: &str) -> Result<bool> {
        let items = custom_xml::discover(&self.opc)?;
        let matching = items
            .iter()
            .filter(|item| {
                item.props()
                    .is_some_and(|props| props.id.eq_ignore_ascii_case(id))
            })
            .collect::<Vec<_>>();
        let Some(first) = matching.first() else {
            return Ok(false);
        };
        if self
            .custom_xml_bindings()?
            .iter()
            .any(|binding| binding.store_id.eq_ignore_ascii_case(id))
        {
            return Err(Error::InvalidFormat(format!(
                "Custom XML itemID '{id}' is still referenced by a content control"
            )));
        }
        for item in &matching {
            self.opc
                .get_part_mut(item.source())?
                .rels_mut()
                .remove(item.rel_id());
        }
        let data_part = first.part().clone();
        let props_part = first.props_part().cloned();
        if !self.part_is_referenced(&data_part) {
            self.opc.remove_part(&data_part);
            if let Some(props_part) = props_part
                && !self.part_is_referenced(&props_part)
            {
                self.opc.remove_part(&props_part);
            }
        }
        self.opc.unsign();
        Ok(true)
    }

    /// Locate the document's bibliography source store, if one exists.
    ///
    /// Returns the Custom XML item GUID and the store payload. Word keeps a
    /// single current source list; when several stores exist the first in
    /// package order is used and the rest are left untouched.
    fn bibliography_store_item(&self) -> Result<Option<(String, Vec<u8>)>> {
        let stores = self.bibliography_source_stores()?;
        let Some(store) = stores.first() else {
            return Ok(None);
        };
        let item_id = store
            .data_store_item_id()
            .ok_or_else(|| {
                Error::InvalidFormat("bibliography source store has no Custom XML item GUID".into())
            })?
            .to_owned();
        let item = self.custom_xml_by_id(&item_id)?.ok_or_else(|| {
            Error::PartNotFound(format!("bibliography source store item '{item_id}'"))
        })?;
        Ok(Some((item_id, item.xml().to_vec())))
    }

    /// Add a typed bibliography source to the document's source store.
    ///
    /// When no store exists, one is created as a Custom XML data store with
    /// the bibliography namespace registered. Otherwise the source is
    /// appended in place, preserving untouched entries, style metadata, and
    /// the store's relationship/content-type graph. Duplicate tags are
    /// rejected. Returns the Custom XML item GUID of the store.
    pub fn add_bibliography_source(
        &mut self,
        source: crate::bibliography::BibliographySourceBuilder,
    ) -> Result<String> {
        if let Some((item_id, xml)) = self.bibliography_store_item()? {
            let updated = crate::bibliography::add_source_xml(&xml, &source)?;
            self.set_custom_xml(&item_id, updated.into_bytes())?;
            // Re-validate the mutated store through the read side.
            self.bibliography_source_stores()?;
            Ok(item_id)
        } else {
            let xml = crate::bibliography::new_store_xml(&[source])?;
            let item = self.add_custom_xml(NewStore {
                xml: xml.into_bytes(),
                content_type: "application/xml".to_string(),
                id: crate::bibliography::DEFAULT_STORE_ITEM_ID.to_string(),
                schemas: vec![crate::OOXML_BIBLIOGRAPHY_NAMESPACE.to_string()],
                conformance: custom_xml::Conformance::Transitional,
            })?;
            item.props().map(|props| props.id.clone()).ok_or_else(|| {
                Error::InvalidFormat("new bibliography Custom XML store has no item GUID".into())
            })
        }
    }

    /// Remove the bibliography source with the given tag from the source
    /// store. Returns whether a source was removed.
    pub fn remove_bibliography_source(&mut self, tag: &str) -> Result<bool> {
        let Some((item_id, xml)) = self.bibliography_store_item()? else {
            return Ok(false);
        };
        let (updated, removed) = crate::bibliography::remove_source_xml(&xml, tag)?;
        if removed {
            self.set_custom_xml(&item_id, updated.into_bytes())?;
            // Re-validate the mutated store through the read side.
            self.bibliography_source_stores()?;
        }
        Ok(removed)
    }

    /// Replace the bibliography source with the given tag, preserving entry
    /// order and all untouched entries. Fails when the tag does not exist.
    pub fn replace_bibliography_source(
        &mut self,
        tag: &str,
        source: crate::bibliography::BibliographySourceBuilder,
    ) -> Result<()> {
        let Some((item_id, xml)) = self.bibliography_store_item()? else {
            return Err(Error::PartNotFound(
                "no bibliography source store exists".to_string(),
            ));
        };
        let updated = crate::bibliography::replace_source_xml(&xml, tag, &source)?;
        self.set_custom_xml(&item_id, updated.into_bytes())?;
        // Re-validate the mutated store through the read side.
        self.bibliography_source_stores()?;
        Ok(())
    }

    /// Reorder main-document data-store relationships by item GUID.
    pub fn order_custom_xml(&mut self, ordered_ids: &[String]) -> Result<()> {
        let source_part = self.opc.main_document_part()?.partname().clone();
        let items = custom_xml::discover(&self.opc)?
            .into_iter()
            .filter(|item| item.source() == &source_part)
            .collect::<Vec<_>>();
        if items.len() != ordered_ids.len() {
            return Err(Error::InvalidFormat(
                "reorder list must contain every main-document Custom XML item".into(),
            ));
        }
        let mut by_id = std::collections::HashMap::new();
        for item in &items {
            let id = item
                .props()
                .ok_or_else(|| {
                    Error::InvalidFormat("Custom XML item has no datastore itemID".into())
                })?
                .id
                .to_ascii_lowercase();
            if by_id.insert(id, item).is_some() {
                return Err(Error::InvalidFormat(
                    "main-document Custom XML items are not uniquely reorderable".into(),
                ));
            }
        }
        let mut ordered = Vec::with_capacity(items.len());
        let mut seen = std::collections::HashSet::new();
        for id in ordered_ids {
            let key = id.to_ascii_lowercase();
            if !seen.insert(key.clone()) {
                return Err(Error::InvalidFormat("duplicate reorder itemID".into()));
            }
            let item = *by_id
                .get(&key)
                .ok_or_else(|| Error::InvalidFormat(format!("unknown reorder itemID '{id}'")))?;
            let reltype = self
                .opc
                .get_part(&source_part)?
                .rels()
                .get(item.rel_id())
                .ok_or_else(|| {
                    Error::InvalidRelationship(format!(
                        "Custom XML relationship '{}' disappeared during reorder",
                        item.rel_id()
                    ))
                })?
                .reltype()
                .to_string();
            ordered.push((item, reltype));
        }
        let source = self.opc.get_part(&source_part)?;
        let reserved = source
            .rels()
            .iter()
            .filter(|relationship| {
                !items
                    .iter()
                    .any(|item| item.rel_id() == relationship.r_id())
            })
            .map(|relationship| relationship.r_id().to_string())
            .collect::<std::collections::HashSet<_>>();
        let ids = (1usize..=MAX_ITEMS + 1)
            .filter_map(|batch| {
                let candidates = (0..ordered.len())
                    .map(|index| format!("rIdCustomXmlOrder{batch:04}_{index:06}"))
                    .collect::<Vec<_>>();
                candidates
                    .iter()
                    .all(|id| !reserved.contains(id))
                    .then_some(candidates)
            })
            .next()
            .ok_or_else(|| {
                Error::InvalidFormat("Custom XML reorder relationship ID space is exhausted".into())
            })?;
        let source = self.opc.get_part_mut(&source_part)?;
        let source_base_uri = source.partname().base_uri().to_string();
        for item in &items {
            source.rels_mut().remove(item.rel_id());
        }
        for ((item, reltype), id) in ordered.into_iter().zip(ids) {
            source.rels_mut().add_relationship(
                reltype,
                item.part().relative_ref(&source_base_uri),
                id,
                false,
            );
        }
        self.opc.unsign();
        Ok(())
    }

    /// Collect and lexically validate SDT bindings from every permitted Word container.
    pub fn custom_xml_bindings(&self) -> Result<Vec<Binding>> {
        let permitted = [
            ct::WML_DOCUMENT_MAIN,
            ct::WML_DOCUMENT_GLOSSARY,
            ct::WML_HEADER,
            ct::WML_FOOTER,
            ct::WML_FOOTNOTES,
            ct::WML_ENDNOTES,
        ];
        let mut bindings = Vec::new();
        for part in self
            .opc
            .iter_parts()
            .filter(|part| permitted.contains(&part.content_type()))
        {
            for control in ContentControl::extract_from_document(part.blob())? {
                control.validate_data_binding()?;
                if let (Some(xpath), Some(store_item_id)) = (
                    control.data_binding_xpath(),
                    control.data_binding_store_item_id(),
                ) {
                    bindings.push(Binding {
                        source: part.partname().clone(),
                        control_id: control.id(),
                        xpath: xpath.to_string(),
                        store_id: store_item_id.to_string(),
                        prefixes: control.data_binding_prefix_mappings().map(str::to_string),
                    });
                }
            }
        }
        bindings.sort_unstable_by(|left, right| {
            left.source
                .as_str()
                .cmp(right.source.as_str())
                .then_with(|| left.control_id.cmp(&right.control_id))
        });
        Ok(bindings)
    }

    /// Validate that every permitted SDT binding resolves to a datastore item GUID.
    pub fn validate_custom_xml_bindings(&self) -> Result<()> {
        let item_ids = custom_xml::discover(&self.opc)?
            .into_iter()
            .filter_map(|item| item.props().map(|props| props.id.to_ascii_lowercase()))
            .collect::<std::collections::HashSet<_>>();
        for binding in self.custom_xml_bindings()? {
            if !item_ids.contains(&binding.store_id.to_ascii_lowercase()) {
                return Err(Error::InvalidFormat(format!(
                    "content control {} in '{}' references missing Custom XML itemID '{}'",
                    binding.control_id,
                    binding.source.as_str(),
                    binding.store_id
                )));
            }
        }
        Ok(())
    }

    pub(super) fn part_is_referenced(&self, target: &PackURI) -> bool {
        self.opc.rels().iter().any(|relationship| {
            !relationship.is_external()
                && relationship
                    .target_partname()
                    .is_ok_and(|part| &part == target)
        }) || self.opc.iter_parts().any(|part| {
            part.rels().iter().any(|relationship| {
                !relationship.is_external()
                    && relationship
                        .target_partname()
                        .is_ok_and(|part| &part == target)
            })
        })
    }
}
