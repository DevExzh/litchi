//! Custom XML data stores, bibliography sources, and SDT bindings.

use super::super::model::*;
use crate::content_control::{Checksum, ChecksumValue, Inventory, PackageLimits, Snapshot};
use crate::package::story::{self, StoryInventory, StoryTopology};
use std::sync::Arc;

#[derive(Debug)]
enum StoreMutation {
    Payload {
        xml: Arc<Vec<u8>>,
    },
    Complete {
        xml: Arc<Vec<u8>>,
        content_type: String,
        props_xml: Arc<Vec<u8>>,
    },
}

impl StoreMutation {
    fn payload(&self) -> &[u8] {
        match self {
            Self::Payload { xml } | Self::Complete { xml, .. } => xml.as_slice(),
        }
    }

    fn payload_owner(&self) -> Arc<Vec<u8>> {
        match self {
            Self::Payload { xml } | Self::Complete { xml, .. } => Arc::clone(xml),
        }
    }

    fn is_noop(&self, package: &OpcPackage, item: &CustomXmlItem) -> Result<bool> {
        if item.xml() != self.payload() {
            return Ok(false);
        }
        let Self::Complete {
            content_type,
            props_xml,
            ..
        } = self
        else {
            return Ok(true);
        };
        let props_part = item.props_part().ok_or_else(|| {
            Error::InvalidFormat("Custom XML data store has no properties part".into())
        })?;
        Ok(item.content_type() == content_type
            && package.get_part(props_part)?.blob() == props_xml.as_slice())
    }

    fn apply(self, package: &mut OpcPackage, item: &CustomXmlItem) -> Result<()> {
        match self {
            Self::Payload { xml } => {
                package.get_part_mut(item.part())?.set_blob_shared(xml);
            },
            Self::Complete {
                xml,
                content_type,
                props_xml,
            } => {
                let props_part = item.props_part().cloned().ok_or_else(|| {
                    Error::InvalidFormat("Custom XML data store has no properties part".into())
                })?;
                let relationships = package
                    .get_part(item.part())?
                    .rels()
                    .iter()
                    .map(|relationship| {
                        (
                            relationship.reltype().to_owned(),
                            relationship.target_ref().to_owned(),
                            relationship.r_id().to_owned(),
                            relationship.is_external(),
                        )
                    })
                    .collect::<Vec<_>>();
                let mut part =
                    BlobPart::new_shared(item.part().clone(), content_type, Arc::clone(&xml));
                for (reltype, target, id, external) in relationships {
                    part.rels_mut()
                        .add_relationship(reltype, target, id, external);
                }
                package.add_part(Box::new(part));
                package
                    .get_part_mut(&props_part)?
                    .set_blob_shared(props_xml);
            },
        }
        Ok(())
    }
}

struct ChecksumPlan {
    topology: StoryTopology,
    expected: Option<Checksum>,
    stories: Vec<(PackURI, Arc<Vec<u8>>)>,
    sources: Vec<(PackURI, Arc<Vec<u8>>)>,
    validate: Vec<PackURI>,
}

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
    ///
    /// Signed packages require an explicit [`Self::unsign`] before this
    /// mutating operation.
    pub fn add_custom_xml(&mut self, store: NewStore) -> Result<CustomXmlItem> {
        custom_xml::validate_content_type(&store.content_type)?;
        custom_xml::validate_payload(&store.xml)?;
        let props = CustomXmlProps {
            id: store.id,
            schemas: store.schemas,
        };
        custom_xml::validate_props(&props)?;
        let item_id = props.id.clone();
        self.ensure_story_opc_current("add_custom_xml")?;
        ensure_store_absent(&self.opc, &item_id)?;
        self.refuse_signed_custom_xml_change("add_custom_xml")?;
        let payload = store.xml;
        let limits = PackageLimits::default();
        self.edit_semantic_opc("add_custom_xml", move |candidate| {
            ensure_store_absent(candidate, &item_id)?;
            let source_part = candidate.main_document_part()?.partname().clone();
            let source = candidate.get_part(&source_part)?;
            let rel_id = (1usize..=MAX_ITEMS + 1)
                .map(|number| format!("rIdCustomXml{number}"))
                .find(|id| source.rels().get(id).is_none())
                .ok_or_else(|| {
                    Error::InvalidFormat("Custom XML relationship ID space is exhausted".into())
                })?;
            let mut part_names = None;
            for number in 1usize..=MAX_ITEMS + 1 {
                let data = PackURI::new(format!("/customXml/item{number}.xml"))
                    .map_err(Error::InvalidUri)?;
                let props_part = PackURI::new(format!("/customXml/itemProps{number}.xml"))
                    .map_err(Error::InvalidUri)?;
                let conflict = candidate.iter_parts().any(|part| {
                    part.partname().as_str().eq_ignore_ascii_case(data.as_str())
                        || part
                            .partname()
                            .as_str()
                            .eq_ignore_ascii_case(props_part.as_str())
                });
                if !conflict {
                    part_names = Some((data, props_part));
                    break;
                }
            }
            let (data_part, props_part) = part_names.ok_or_else(|| {
                Error::InvalidFormat("Custom XML part-name space is exhausted".into())
            })?;
            let plan = checksum_plan(candidate, &item_id, &payload, true, &limits)?;
            custom_xml::add(
                candidate,
                NewCustomXmlItem {
                    source: source_part,
                    rel_id,
                    part: data_part.clone(),
                    content_type: store.content_type,
                    xml: payload,
                    props: Some(NewCustomXmlProps {
                        part: props_part,
                        rel_id: "rIdProps1".to_string(),
                        value: props,
                    }),
                    conformance: store.conformance,
                },
            )?;
            publish_checksum_plan(candidate, &plan)?;
            let published = unique_store(candidate, &item_id)?;
            if published.part() != &data_part {
                return Err(Error::InvalidFormat(
                    "new Custom XML data store was not discoverable at its exact target".into(),
                ));
            }
            verify_checksum_plan(candidate, &item_id, &plan, &limits)?;
            Ok(published)
        })
    }

    /// Replace only the inert XML payload of a data store.
    ///
    /// An exact no-op preserves package signatures. A changed signed package
    /// requires an explicit [`Self::unsign`] first.
    pub fn set_custom_xml(&mut self, id: &str, xml: Vec<u8>) -> Result<()> {
        custom_xml::validate_payload(&xml)?;
        self.mutate_custom_xml(
            id,
            "set_custom_xml",
            StoreMutation::Payload { xml: Arc::new(xml) },
        )
    }

    /// Replace payload, content type, schema references, and canonical properties.
    ///
    /// An exact no-op preserves package signatures. A changed signed package
    /// requires an explicit [`Self::unsign`] first.
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
        self.mutate_custom_xml(
            id,
            "replace_custom_xml",
            StoreMutation::Complete {
                xml: Arc::new(replacement.xml),
                content_type: replacement.content_type,
                props_xml: Arc::new(props_xml),
            },
        )
    }

    fn mutate_custom_xml(
        &mut self,
        id: &str,
        operation: &'static str,
        mutation: StoreMutation,
    ) -> Result<()> {
        self.ensure_story_opc_current(operation)?;
        let item = unique_store(&self.opc, id)?;
        if mutation.is_noop(&self.opc, &item)? {
            return Ok(());
        }
        self.refuse_signed_custom_xml_change(operation)?;
        let expected_part = item.part().clone();
        let item_id = id.to_owned();
        let limits = PackageLimits::default();
        self.edit_semantic_opc(operation, move |candidate| {
            let item = unique_store(candidate, &item_id)?;
            if item.part() != &expected_part {
                return Err(Error::InvalidFormat(format!(
                    "Custom XML itemID '{item_id}' changed ownership during mutation"
                )));
            }
            let payload = mutation.payload_owner();
            let payload_changed = item.xml() != payload.as_slice();
            let plan = checksum_plan(
                candidate,
                &item_id,
                payload.as_slice(),
                payload_changed,
                &limits,
            )?;
            mutation.apply(candidate, &item)?;
            publish_checksum_plan(candidate, &plan)?;
            let published = unique_store(candidate, &item_id)?;
            if published.part() != &expected_part || published.xml() != payload.as_slice() {
                return Err(Error::InvalidFormat(
                    "Custom XML mutation did not retain its exact target payload".into(),
                ));
            }
            verify_checksum_plan(candidate, &item_id, &plan, &limits)
        })
    }

    /// Remove a data store unless an SDT still binds to its item GUID.
    ///
    /// A missing-store no-op preserves package signatures. Removing an
    /// existing store from a signed package requires [`Self::unsign`] first.
    pub fn remove_custom_xml(&mut self, id: &str) -> Result<bool> {
        self.ensure_story_opc_current("remove_custom_xml")?;
        let present = custom_xml::discover(&self.opc)?.iter().any(|item| {
            item.props()
                .is_some_and(|props| props.id.eq_ignore_ascii_case(id))
        });
        if !present {
            return Ok(false);
        }
        self.refuse_signed_custom_xml_change("remove_custom_xml")?;
        let id = id.to_owned();
        let limits = PackageLimits::default();
        self.edit_semantic_opc("remove_custom_xml", move |candidate| {
            let items = custom_xml::discover(candidate)?;
            let matching = items
                .iter()
                .filter(|item| {
                    item.props()
                        .is_some_and(|props| props.id.eq_ignore_ascii_case(&id))
                })
                .collect::<Vec<_>>();
            let first = matching.first().ok_or_else(|| {
                Error::InvalidFormat("Custom XML removal target became stale".into())
            })?;
            let stories = story::capture(candidate, limits.stories)?;
            visit_bindings(
                &stories,
                &limits,
                |source, occurrence, control_id, binding| {
                    if binding.store_item_id().eq_ignore_ascii_case(&id) {
                        return Err(Error::InvalidFormat(format!(
                            "content-control occurrence {occurrence} (producer ID {control_id:?}) in '{}' still references Custom XML itemID '{id}'",
                            source
                        )));
                    }
                    Ok(())
                },
            )?;
            let data_part = first.part().clone();
            let props_part = first.props_part().cloned();
            for item in matching {
                candidate
                    .get_part_mut(item.source())?
                    .rels_mut()
                    .remove(item.rel_id());
            }
            if !part_is_referenced_in(candidate, &data_part) {
                candidate.remove_part(&data_part);
                if let Some(props_part) = props_part
                    && !part_is_referenced_in(candidate, &props_part)
                {
                    candidate.remove_part(&props_part);
                }
            }
            if custom_xml::discover(candidate)?.iter().any(|item| {
                item.props()
                    .is_some_and(|props| props.id.eq_ignore_ascii_case(&id))
            }) {
                return Err(Error::InvalidFormat(
                    "Custom XML removal target remained discoverable".into(),
                ));
            }
            Ok(true)
        })
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
    ///
    /// An unchanged order preserves package signatures. Reordering a signed
    /// package requires an explicit [`Self::unsign`] first.
    pub fn order_custom_xml(&mut self, ordered_ids: &[String]) -> Result<()> {
        self.ensure_story_opc_current("order_custom_xml")?;
        let current = main_store_order(&self.opc)?;
        validate_requested_order(&current, ordered_ids)?;
        if current
            .iter()
            .zip(ordered_ids)
            .all(|(left, right)| left.eq_ignore_ascii_case(right))
        {
            return Ok(());
        }
        self.refuse_signed_custom_xml_change("order_custom_xml")?;
        let ordered_ids = ordered_ids.to_vec();
        self.edit_semantic_opc("order_custom_xml", move |candidate| {
            reorder_custom_xml(candidate, &ordered_ids)
        })
    }

    /// Collect and lexically validate SDT bindings from reachable Word stories.
    pub fn custom_xml_bindings(&self) -> Result<Vec<Binding>> {
        let limits = PackageLimits::default();
        let stories = self.story_inventory_with_limits(limits.stories)?;
        let mut bindings = Vec::new();
        visit_bindings(
            &stories,
            &limits,
            |source, occurrence, control_id, binding| {
                bindings
                    .try_reserve(1)
                    .map_err(|source| Error::Allocation {
                        resource: "Custom XML binding inventory",
                        source,
                    })?;
                bindings.push(Binding {
                    source: source.clone(),
                    occurrence,
                    control_id,
                    flavor: binding.flavor(),
                    xpath: binding.xpath().to_owned(),
                    store_id: binding.store_item_id().to_owned(),
                    prefixes: binding.prefix_mappings().map(str::to_owned),
                });
                Ok(())
            },
        )?;
        Ok(bindings)
    }

    fn refuse_signed_custom_xml_change(&self, operation: &'static str) -> Result<()> {
        if self.is_signed() {
            return Err(Error::UnsafeEdit {
                format: "DOCX",
                operation,
                reason: "signed Custom XML changes require explicit Package::unsign before mutation",
            });
        }
        Ok(())
    }
}

fn reorder_custom_xml(package: &mut OpcPackage, ordered_ids: &[String]) -> Result<()> {
    let source_part = package.main_document_part()?.partname().clone();
    let items = custom_xml::discover(package)?
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
            .ok_or_else(|| Error::InvalidFormat("Custom XML item has no datastore itemID".into()))?
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
        package
            .get_part(&source_part)?
            .rels()
            .get(item.rel_id())
            .ok_or_else(|| {
                Error::InvalidRelationship(format!(
                    "Custom XML relationship '{}' disappeared during reorder",
                    item.rel_id()
                ))
            })?;
        ordered.push(item);
    }
    // Relationship XML is serialized in rId order.  Preserve each original
    // relationship identity and retarget those stable slots to the requested
    // item order; deleting and recreating relationships would change rIds and
    // invalidate otherwise opaque references to them.
    let source_base_uri = package
        .main_document_part()?
        .partname()
        .base_uri()
        .to_string();
    let retargets = items
        .iter()
        .zip(ordered)
        .map(|(slot, item)| {
            (
                slot.rel_id().to_owned(),
                item.part().relative_ref(&source_base_uri),
            )
        })
        .collect::<Vec<_>>();
    let source = package.get_part_mut(&source_part)?;
    for (relationship_id, target_ref) in retargets {
        source.rels_mut().retarget(&relationship_id, target_ref)?;
    }
    Ok(())
}

impl Package {
    /// Validate that every permitted SDT binding resolves to a datastore item GUID.
    pub fn validate_custom_xml_bindings(&self) -> Result<()> {
        let item_ids = custom_xml::discover(&self.opc)?
            .into_iter()
            .filter_map(|item| item.props().map(|props| props.id.to_ascii_lowercase()))
            .collect::<std::collections::HashSet<_>>();
        for binding in self.custom_xml_bindings()? {
            if !item_ids.contains(&binding.store_id.to_ascii_lowercase()) {
                return Err(Error::InvalidFormat(format!(
                    "content-control occurrence {} (producer ID {:?}) in '{}' references missing Custom XML itemID '{}'",
                    binding.occurrence,
                    binding.control_id,
                    binding.source.as_str(),
                    binding.store_id
                )));
            }
        }
        Ok(())
    }

    pub(super) fn part_is_referenced(&self, target: &PackURI) -> bool {
        part_is_referenced_in(&self.opc, target)
    }
}

fn part_is_referenced_in(package: &OpcPackage, target: &PackURI) -> bool {
    package.rels().iter().any(|relationship| {
        !relationship.is_external()
            && relationship
                .target_partname()
                .is_ok_and(|part| &part == target)
    }) || package.iter_parts().any(|part| {
        part.rels().iter().any(|relationship| {
            !relationship.is_external()
                && relationship
                    .target_partname()
                    .is_ok_and(|part| &part == target)
        })
    })
}

fn visit_bindings(
    stories: &StoryInventory,
    limits: &PackageLimits,
    mut visit: impl FnMut(
        &PackURI,
        usize,
        Option<u32>,
        &crate::content_control::DataBinding,
    ) -> Result<()>,
) -> Result<()> {
    let mut controls = 0usize;
    for story in stories.stories() {
        let inventory = Inventory::parse_with_limits(story.source(), &limits.controls)?;
        controls = controls
            .checked_add(inventory.occurrences().len())
            .ok_or_else(|| Error::InvalidFormat("content-control count overflow".into()))?;
        if controls > limits.max_content_controls {
            return Err(Error::InvalidFormat(
                "package content-control count exceeds configured limit".into(),
            ));
        }
        for occurrence in inventory.occurrences() {
            let control = occurrence.control();
            control.validate_data_binding()?;
            for binding in control.data_bindings() {
                visit(story.part(), occurrence.ordinal(), occurrence.id(), binding)?;
            }
        }
    }
    Ok(())
}

fn unique_store(package: &OpcPackage, id: &str) -> Result<CustomXmlItem> {
    let mut found = None::<CustomXmlItem>;
    for item in custom_xml::discover(package)? {
        if !item
            .props()
            .is_some_and(|props| props.id.eq_ignore_ascii_case(id))
        {
            continue;
        }
        if let Some(current) = &found {
            if current.part() != item.part() {
                return Err(Error::InvalidFormat(format!(
                    "Custom XML itemID '{id}' is ambiguous"
                )));
            }
        } else {
            found = Some(item);
        }
    }
    found.ok_or_else(|| Error::PartNotFound(format!("Custom XML itemID '{id}'")))
}

fn ensure_store_absent(package: &OpcPackage, id: &str) -> Result<()> {
    if custom_xml::discover(package)?.iter().any(|item| {
        item.props()
            .is_some_and(|props| props.id.eq_ignore_ascii_case(id))
    }) {
        return Err(Error::InvalidFormat(format!(
            "Custom XML itemID '{id}' already exists"
        )));
    }
    Ok(())
}

fn checksum_plan(
    package: &OpcPackage,
    id: &str,
    payload: &[u8],
    payload_changed: bool,
    limits: &PackageLimits,
) -> Result<ChecksumPlan> {
    let inventory = story::capture(package, limits.stories)?;
    let topology = inventory.topology();
    let mut sources = Vec::new();
    sources
        .try_reserve_exact(inventory.stories().len())
        .map_err(|source| Error::Allocation {
            resource: "Custom XML checksum story sources",
            source,
        })?;
    if !payload_changed {
        sources.extend(
            inventory
                .stories()
                .iter()
                .map(|story| (story.part().clone(), story.source_arc())),
        );
        return Ok(ChecksumPlan {
            topology,
            expected: None,
            stories: Vec::new(),
            sources,
            validate: Vec::new(),
        });
    }
    let mut checksum = None::<Checksum>;
    let mut stories = Vec::new();
    let mut validate = Vec::new();
    let mut controls = 0usize;
    let mut mutations = 0usize;
    let mut output_bytes = 0usize;
    for story in inventory.stories() {
        let snapshot = Snapshot::from_package(story.source_arc(), limits.controls.clone())?;
        controls = controls
            .checked_add(snapshot.occurrences().len())
            .ok_or_else(|| Error::InvalidFormat("content-control count overflow".into()))?;
        if controls > limits.max_content_controls {
            return Err(Error::InvalidFormat(
                "package content-control count exceeds configured limit".into(),
            ));
        }
        let mut transaction = snapshot.edit();
        let mut declared = false;
        for (ordinal, occurrence) in snapshot.inventory().occurrences().iter().enumerate() {
            let source = snapshot.occurrences().get(ordinal).ok_or_else(|| {
                Error::InvalidFormat("content-control source inventory is stale".into())
            })?;
            for (binding_index, binding) in occurrence.control().data_bindings().iter().enumerate()
            {
                if !binding.store_item_id().eq_ignore_ascii_case(id) {
                    continue;
                }
                let Some(value) = binding.checksum_value() else {
                    continue;
                };
                declared = true;
                if matches!(value, ChecksumValue::Malformed(_)) {
                    return Err(Error::InvalidFormat(format!(
                        "content control {ordinal} binding {binding_index} in '{}' has a malformed storeItemChecksum",
                        story.part()
                    )));
                }
                let source_binding = source.bindings().get(binding_index).ok_or_else(|| {
                    Error::InvalidFormat("content-control binding inventory is stale".into())
                })?;
                if source_binding.flavor() != binding.flavor()
                    || source_binding.checksum_count() != 1
                {
                    return Err(Error::InvalidFormat(format!(
                        "content control {ordinal} binding {binding_index} in '{}' has ambiguous checksum ownership",
                        story.part()
                    )));
                }
                mutations = mutations.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("checksum mutation count overflow".into())
                })?;
                if mutations > limits.max_mutations {
                    return Err(Error::InvalidFormat(
                        "content-control mutation limit exceeded".into(),
                    ));
                }
                let expected = match &checksum {
                    Some(value) => value.clone(),
                    None => {
                        let value = Checksum::compute(payload, &limits.controls)?;
                        checksum = Some(value.clone());
                        value
                    },
                };
                transaction.set_binding_checksum(ordinal, binding_index, Some(expected))?;
            }
        }
        let commit = transaction.commit()?;
        if commit.changed() {
            output_bytes = output_bytes
                .checked_add(commit.snapshot().source().len())
                .ok_or_else(|| Error::InvalidFormat("checksum output size overflow".into()))?;
            if output_bytes > limits.max_output_bytes {
                return Err(Error::InvalidFormat(
                    "aggregate content-control output limit exceeded".into(),
                ));
            }
            stories.try_reserve(1).map_err(|source| Error::Allocation {
                resource: "Custom XML checksum story updates",
                source,
            })?;
            let source = commit.snapshot().package_source_arc().ok_or_else(|| {
                Error::InvalidFormat("rewritten story lost package-owned byte storage".into())
            })?;
            stories.push((story.part().clone(), Arc::clone(&source)));
            sources.push((story.part().clone(), source));
        } else {
            sources.push((story.part().clone(), story.source_arc()));
        }
        if declared {
            validate
                .try_reserve(1)
                .map_err(|source| Error::Allocation {
                    resource: "Custom XML checksum validation stories",
                    source,
                })?;
            validate.push(story.part().clone());
        }
    }
    Ok(ChecksumPlan {
        topology,
        expected: checksum,
        stories,
        sources,
        validate,
    })
}

fn publish_checksum_plan(package: &mut OpcPackage, plan: &ChecksumPlan) -> Result<()> {
    for (part, xml) in &plan.stories {
        package.get_part_mut(part)?.set_blob_shared(Arc::clone(xml));
    }
    Ok(())
}

fn verify_checksum_plan(
    package: &OpcPackage,
    id: &str,
    plan: &ChecksumPlan,
    limits: &PackageLimits,
) -> Result<()> {
    let inventory = story::capture(package, limits.stories)?;
    if inventory.topology() != plan.topology {
        return Err(Error::InvalidFormat(
            "Word story topology changed during Custom XML mutation".into(),
        ));
    }
    if inventory.stories().len() != plan.sources.len() {
        return Err(Error::InvalidFormat(
            "Word story inventory changed during Custom XML mutation".into(),
        ));
    }
    for (story, (part, expected)) in inventory.stories().iter().zip(&plan.sources) {
        let actual = story.source_arc();
        if story.part() != part
            || (!Arc::ptr_eq(&actual, expected) && actual.as_slice() != expected.as_slice())
        {
            return Err(Error::InvalidFormat(
                "Word story bytes changed during Custom XML mutation".into(),
            ));
        }
    }
    let Some(expected) = &plan.expected else {
        return Ok(());
    };
    let mut controls = 0usize;
    for part in &plan.validate {
        let story = inventory
            .get(part)
            .ok_or_else(|| Error::InvalidFormat("checksum validation story disappeared".into()))?;
        let snapshot = Snapshot::from_package(story.source_arc(), limits.controls.clone())?;
        controls = controls
            .checked_add(snapshot.occurrences().len())
            .ok_or_else(|| Error::InvalidFormat("content-control count overflow".into()))?;
        if controls > limits.max_content_controls {
            return Err(Error::InvalidFormat(
                "package content-control count exceeds configured limit".into(),
            ));
        }
        for (ordinal, occurrence) in snapshot.inventory().occurrences().iter().enumerate() {
            let source = snapshot.occurrences().get(ordinal).ok_or_else(|| {
                Error::InvalidFormat("content-control source inventory is stale".into())
            })?;
            for (binding_index, binding) in occurrence.control().data_bindings().iter().enumerate()
            {
                if !binding.store_item_id().eq_ignore_ascii_case(id) {
                    continue;
                }
                let Some(value) = binding.checksum_value() else {
                    continue;
                };
                let source_binding = source.bindings().get(binding_index).ok_or_else(|| {
                    Error::InvalidFormat("content-control binding inventory is stale".into())
                })?;
                if source_binding.flavor() != binding.flavor()
                    || source_binding.checksum_count() != 1
                {
                    return Err(Error::InvalidFormat(format!(
                        "content control {ordinal} binding {binding_index} in '{}' has ambiguous checksum ownership",
                        story.part()
                    )));
                }
                match value {
                    ChecksumValue::Valid(actual) if actual.as_bytes() == expected.as_bytes() => {},
                    ChecksumValue::Valid(_) => {
                        return Err(Error::InvalidFormat(format!(
                            "content control {ordinal} binding {binding_index} in '{}' retained a stale storeItemChecksum",
                            story.part()
                        )));
                    },
                    ChecksumValue::Malformed(_) => {
                        return Err(Error::InvalidFormat(format!(
                            "content control {ordinal} binding {binding_index} in '{}' has a malformed storeItemChecksum",
                            story.part()
                        )));
                    },
                }
            }
        }
    }
    Ok(())
}

fn main_store_order(package: &OpcPackage) -> Result<Vec<String>> {
    let source = package.main_document_part()?.partname().clone();
    custom_xml::discover(package)?
        .into_iter()
        .filter(|item| item.source() == &source)
        .map(|item| {
            item.props().map(|props| props.id.clone()).ok_or_else(|| {
                Error::InvalidFormat("Custom XML item has no datastore itemID".into())
            })
        })
        .collect()
}

fn validate_requested_order(current: &[String], requested: &[String]) -> Result<()> {
    if current.len() != requested.len() {
        return Err(Error::InvalidFormat(
            "reorder list must contain every main-document Custom XML item".into(),
        ));
    }
    let current = current
        .iter()
        .map(|id| id.to_ascii_lowercase())
        .collect::<std::collections::HashSet<_>>();
    let mut requested_set = std::collections::HashSet::new();
    for id in requested {
        let normalized = id.to_ascii_lowercase();
        if !requested_set.insert(normalized.clone()) {
            return Err(Error::InvalidFormat("duplicate reorder itemID".into()));
        }
        if !current.contains(&normalized) {
            return Err(Error::InvalidFormat(format!(
                "unknown reorder itemID '{id}'"
            )));
        }
    }
    Ok(())
}
