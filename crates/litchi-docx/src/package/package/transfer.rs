//! Dependency-checked cross-package paragraph transfer planning.

use std::collections::{BTreeMap, BTreeSet, HashMap, HashSet, VecDeque};
use std::sync::Arc;

use litchi_core::Position;
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part, TargetMode};
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;
use quick_xml::{Writer, XmlVersion};

use super::super::model::Package;
use crate::document::{
    ParagraphTransfer, TransactionError, TransferGraph, TransferPart, TransferRefusal,
    TransferRelationship,
};

const TRANSITIONAL_RELATIONSHIPS_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_RELATIONSHIPS_NAMESPACE: &[u8] =
    b"http://purl.oclc.org/ooxml/officeDocument/relationships";
const MAX_TRANSFER_PARTS: usize = 256;
const MAX_TRANSFER_BYTES: usize = 64 * 1024 * 1024;

impl Package {
    /// Plan transfer of one direct-body paragraph from another package without
    /// mutating either package.
    ///
    /// Every relationship reference is resolved in the donor. One exact
    /// equivalent receiver closure is reused; otherwise an external edge or a
    /// complete internal relationship subgraph is copied under deterministic,
    /// collision-free receiver-local names. The resulting compact plan belongs
    /// to this exact receiving package and can be joined with ordinary edits.
    ///
    /// # Errors
    ///
    /// Returns a checked position, package graph, XML, or typed dependency
    /// refusal error.
    pub fn plan_paragraph_transfer_from(
        &self,
        donor: &Self,
        paragraph: Position,
    ) -> Result<ParagraphTransfer, TransactionError> {
        let receiver = self.document_snapshot()?;
        let donor_snapshot = donor.document_snapshot()?;
        let donor_paragraph =
            donor_snapshot
                .paragraph(paragraph)
                .ok_or(TransactionError::OutOfBounds {
                    position: paragraph.get(),
                    len: donor_snapshot.paragraph_count(),
                })?;
        let donor_main = donor.opc.main_document_part().map_err(crate::Error::from)?;
        let receiver_main_name = self
            .opc
            .main_document_part()
            .map_err(crate::Error::from)?
            .partname()
            .clone();
        let relationship_prefixes = relationship_prefixes(donor_snapshot.xml_bytes())?;
        let inherited_namespaces = stable_namespace_declarations(donor_snapshot.xml_bytes())?;
        let dependency_digest = relationship_graph_digest(self)?;
        let mut candidate = self.opc.clone();
        let mut mapping = BTreeMap::new();
        for identifier in
            relationship_references(donor_paragraph.xml_bytes(), &relationship_prefixes)?
        {
            let dependency = donor_main.rels().get(&identifier).ok_or_else(|| {
                TransactionError::Transfer(TransferRefusal::MissingDonorRelationship(
                    identifier.clone(),
                ))
            })?;
            let equivalents =
                equivalent_relationships(&donor.opc, &candidate, dependency, &receiver_main_name)?;
            if equivalents.len() > 1 {
                return Err(TransactionError::Transfer(
                    TransferRefusal::AmbiguousEquivalentDependency {
                        relationship_type: dependency.reltype().to_owned(),
                        target: dependency.target_ref().to_owned(),
                    },
                ));
            }
            let receiver_id = if let Some(selected) = equivalents.first() {
                selected.clone()
            } else {
                copy_relationship_closure(
                    &donor.opc,
                    &mut candidate,
                    &receiver_main_name,
                    dependency,
                )?
            };
            mapping.insert(identifier, receiver_id);
        }
        let rewritten_xml = rewrite_relationship_references(
            donor_paragraph.xml_bytes(),
            &mapping,
            &relationship_prefixes,
            &inherited_namespaces,
        )?;
        let rewritten_text = std::str::from_utf8(&rewritten_xml)
            .map_err(|_error| TransactionError::Transfer(TransferRefusal::InvalidParagraphXml))?;
        let compact = crate::writer::doc::compact_changed_document_xml(rewritten_text)
            .map_err(TransactionError::from)?;
        let inverse_dependency_digest = relationship_graph_digest_opc(&candidate)?;
        let graph = capture_graph_delta(&self.opc, &candidate, &receiver_main_name)?;
        Ok(ParagraphTransfer::new(
            Arc::new(receiver.xml_bytes().to_vec()),
            compact.into_bytes(),
            dependency_digest,
            inverse_dependency_digest,
            graph,
        ))
    }
}

fn equivalent_relationships(
    donor: &OpcPackage,
    receiver: &OpcPackage,
    dependency: &litchi_opc::Relationship,
    receiver_main_name: &PackURI,
) -> Result<Vec<String>, TransactionError> {
    let receiver_main = receiver
        .get_part(receiver_main_name)
        .map_err(crate::Error::from)?;
    let mut identifiers = Vec::new();
    for candidate in receiver_main.rels().iter() {
        if candidate.reltype() != dependency.reltype()
            || candidate.target_mode() != dependency.target_mode()
        {
            continue;
        }
        let equivalent = if dependency.is_external() {
            candidate.target_ref() == dependency.target_ref()
        } else if candidate.is_external() {
            false
        } else {
            let donor_target = dependency.target_partname().map_err(crate::Error::from)?;
            let receiver_target = candidate.target_partname().map_err(crate::Error::from)?;
            equivalent_part_closure(
                donor,
                receiver,
                &donor_target,
                &receiver_target,
                &mut BTreeSet::new(),
            )?
        };
        if equivalent {
            identifiers.push(candidate.r_id().to_owned());
        }
    }
    identifiers.sort();
    Ok(identifiers)
}

fn equivalent_part_closure(
    donor: &OpcPackage,
    receiver: &OpcPackage,
    donor_name: &PackURI,
    receiver_name: &PackURI,
    visited: &mut BTreeSet<(String, String)>,
) -> Result<bool, TransactionError> {
    if !visited.insert((
        donor_name.as_str().to_owned(),
        receiver_name.as_str().to_owned(),
    )) {
        return Ok(true);
    }
    let donor_part = donor.get_part(donor_name).map_err(crate::Error::from)?;
    let receiver_part = receiver
        .get_part(receiver_name)
        .map_err(crate::Error::from)?;
    if donor_part.content_type() != receiver_part.content_type()
        || donor_part.blob() != receiver_part.blob()
        || donor_part.rels().len() != receiver_part.rels().len()
    {
        return Ok(false);
    }
    for donor_relationship in donor_part.rels().iter() {
        let Some(receiver_relationship) = receiver_part.rels().get(donor_relationship.r_id())
        else {
            return Ok(false);
        };
        if donor_relationship.reltype() != receiver_relationship.reltype()
            || donor_relationship.target_mode() != receiver_relationship.target_mode()
        {
            return Ok(false);
        }
        if donor_relationship.is_external() {
            if donor_relationship.target_ref() != receiver_relationship.target_ref() {
                return Ok(false);
            }
        } else {
            let donor_target = donor_relationship
                .target_partname()
                .map_err(crate::Error::from)?;
            let receiver_target = receiver_relationship
                .target_partname()
                .map_err(crate::Error::from)?;
            if !equivalent_part_closure(donor, receiver, &donor_target, &receiver_target, visited)?
            {
                return Ok(false);
            }
        }
    }
    Ok(true)
}

fn copy_relationship_closure(
    donor: &OpcPackage,
    receiver: &mut OpcPackage,
    receiver_main_name: &PackURI,
    dependency: &litchi_opc::Relationship,
) -> Result<String, TransactionError> {
    if dependency.is_external() {
        return Ok(receiver
            .get_part_mut(receiver_main_name)
            .map_err(crate::Error::from)?
            .relate_to_ext(dependency.target_ref(), dependency.reltype()));
    }
    let root = dependency.target_partname().map_err(crate::Error::from)?;
    let mapping = plan_part_closure(donor, receiver, &root)?;
    publish_part_closure(donor, receiver, &mapping)?;
    let copied_root = mapping.get(&root).ok_or_else(|| {
        TransactionError::Transfer(TransferRefusal::MissingDonorRelationship(
            dependency.r_id().to_owned(),
        ))
    })?;
    let target = copied_root.relative_ref(receiver_main_name.base_uri());
    Ok(receiver
        .get_part_mut(receiver_main_name)
        .map_err(crate::Error::from)?
        .relate_to(&target, dependency.reltype()))
}

fn plan_part_closure(
    donor: &OpcPackage,
    receiver: &OpcPackage,
    root: &PackURI,
) -> Result<HashMap<PackURI, PackURI>, TransactionError> {
    let mut queue = VecDeque::from([root.clone()]);
    let mut seen = HashSet::new();
    let mut total_bytes = 0usize;
    while let Some(name) = queue.pop_front() {
        if !seen.insert(name.clone()) {
            continue;
        }
        if seen.len() > MAX_TRANSFER_PARTS {
            return Err(TransactionError::Limit {
                resource: "transfer dependency parts",
                max: MAX_TRANSFER_PARTS,
                actual: seen.len(),
            });
        }
        let part = donor.get_part(&name).map_err(crate::Error::from)?;
        total_bytes =
            total_bytes
                .checked_add(part.blob().len())
                .ok_or(TransactionError::Limit {
                    resource: "transfer dependency bytes",
                    max: MAX_TRANSFER_BYTES,
                    actual: usize::MAX,
                })?;
        if total_bytes > MAX_TRANSFER_BYTES {
            return Err(TransactionError::Limit {
                resource: "transfer dependency bytes",
                max: MAX_TRANSFER_BYTES,
                actual: total_bytes,
            });
        }
        for relationship in part.rels().iter().filter(|edge| !edge.is_external()) {
            if relationship.target_query().is_some() || relationship.target_fragment().is_some() {
                return Err(TransactionError::Transfer(
                    TransferRefusal::InvalidParagraphXml,
                ));
            }
            queue.push_back(relationship.target_partname().map_err(crate::Error::from)?);
        }
    }
    let mut names = seen.into_iter().collect::<Vec<_>>();
    names.sort_by(|left, right| left.as_str().cmp(right.as_str()));
    let mut reserved = receiver
        .iter_parts()
        .map(|part| part.partname().clone())
        .collect::<HashSet<_>>();
    let mut mapping = HashMap::new();
    for source_name in names {
        let target_name = available_transfer_name(&source_name, &reserved)?;
        reserved.insert(target_name.clone());
        mapping.insert(source_name, target_name);
    }
    Ok(mapping)
}

fn available_transfer_name(
    source: &PackURI,
    reserved: &HashSet<PackURI>,
) -> Result<PackURI, TransactionError> {
    if !reserved.contains(source) {
        return Ok(source.clone());
    }
    let value = source.as_str();
    let (stem, extension) = value
        .rfind('.')
        .map_or((value, ""), |position| value.split_at(position));
    for index in 1..=u32::MAX {
        let candidate = PackURI::new(format!("{stem}-transfer{index}{extension}"))
            .map_err(crate::Error::Invalid)?;
        if !reserved.contains(&candidate) {
            return Ok(candidate);
        }
    }
    Err(TransactionError::Transfer(
        TransferRefusal::InvalidParagraphXml,
    ))
}

fn publish_part_closure(
    donor: &OpcPackage,
    receiver: &mut OpcPackage,
    mapping: &HashMap<PackURI, PackURI>,
) -> Result<(), TransactionError> {
    let mut staged = Vec::new();
    let mut names = mapping.keys().collect::<Vec<_>>();
    names.sort_by(|left, right| left.as_str().cmp(right.as_str()));
    for source_name in names {
        let target_name = mapping.get(source_name).ok_or(TransactionError::Transfer(
            TransferRefusal::InvalidParagraphXml,
        ))?;
        let source = donor.get_part(source_name).map_err(crate::Error::from)?;
        let mut target = BlobPart::new_shared(
            target_name.clone(),
            source.content_type().to_owned(),
            source.blob_arc(),
        );
        for relationship in source.rels().iter() {
            let target_ref = if relationship.is_external() {
                relationship.target_ref().to_owned()
            } else {
                let source_target = relationship.target_partname().map_err(crate::Error::from)?;
                let copied_target = mapping.get(&source_target).ok_or_else(|| {
                    TransactionError::Transfer(TransferRefusal::MissingDonorRelationship(
                        relationship.r_id().to_owned(),
                    ))
                })?;
                copied_target.relative_ref(target_name.base_uri())
            };
            target
                .rels_mut()
                .try_add_relationship(
                    relationship.reltype().to_owned(),
                    target_ref,
                    relationship.r_id().to_owned(),
                    if relationship.is_external() {
                        TargetMode::External
                    } else {
                        TargetMode::Internal
                    },
                )
                .map_err(crate::Error::from)?;
        }
        staged.push(target);
    }
    for part in staged {
        receiver
            .try_add_part(Box::new(part))
            .map_err(crate::Error::from)?;
    }
    Ok(())
}

fn capture_graph_delta(
    before: &OpcPackage,
    after: &OpcPackage,
    main_name: &PackURI,
) -> Result<TransferGraph, TransactionError> {
    let before_main = before.get_part(main_name).map_err(crate::Error::from)?;
    let after_main = after.get_part(main_name).map_err(crate::Error::from)?;
    let mut main_relationships = after_main
        .rels()
        .iter()
        .filter(|relationship| before_main.rels().get(relationship.r_id()).is_none())
        .map(capture_relationship)
        .collect::<Vec<_>>();
    main_relationships.sort_by(|left, right| left.id.cmp(&right.id));
    let mut parts = Vec::new();
    for part in after
        .iter_parts()
        .filter(|part| before.get_part(part.partname()).is_err())
    {
        let mut relationships = part
            .rels()
            .iter()
            .map(capture_relationship)
            .collect::<Vec<_>>();
        relationships.sort_by(|left, right| left.id.cmp(&right.id));
        parts.push(TransferPart {
            name: part.partname().as_str().to_owned(),
            content_type: part.content_type().to_owned(),
            blob: part.blob_arc(),
            relationships: relationships.into(),
        });
    }
    parts.sort_by(|left, right| left.name.cmp(&right.name));
    Ok(TransferGraph {
        main_relationships: main_relationships.into(),
        parts: parts.into(),
    })
}

fn capture_relationship(relationship: &litchi_opc::Relationship) -> TransferRelationship {
    TransferRelationship {
        id: relationship.r_id().to_owned(),
        relationship_type: relationship.reltype().to_owned(),
        target: relationship.target_ref().to_owned(),
        external: relationship.is_external(),
    }
}

pub(super) fn apply_transfer_graph(
    package: &mut OpcPackage,
    graph: &TransferGraph,
    insert: bool,
) -> Result<(), TransactionError> {
    if graph.is_empty() {
        return Ok(());
    }
    let main_name = package
        .main_document_part()
        .map_err(crate::Error::from)?
        .partname()
        .clone();
    if insert {
        for planned in graph.parts.iter() {
            let name = PackURI::new(planned.name.clone()).map_err(crate::Error::Invalid)?;
            let mut part = BlobPart::new_shared(
                name,
                planned.content_type.clone(),
                Arc::clone(&planned.blob),
            );
            add_relationships(part.rels_mut(), &planned.relationships)?;
            package
                .try_add_part(Box::new(part))
                .map_err(crate::Error::from)?;
        }
        let main = package
            .get_part_mut(&main_name)
            .map_err(crate::Error::from)?;
        add_relationships(main.rels_mut(), &graph.main_relationships)?;
    } else {
        let main = package
            .get_part_mut(&main_name)
            .map_err(crate::Error::from)?;
        for relationship in graph.main_relationships.iter() {
            let removed = main.rels_mut().remove(&relationship.id).ok_or_else(|| {
                TransactionError::Transfer(TransferRefusal::MissingDonorRelationship(
                    relationship.id.clone(),
                ))
            })?;
            if removed.reltype() != relationship.relationship_type
                || removed.target_ref() != relationship.target
                || removed.is_external() != relationship.external
            {
                return Err(TransactionError::SemanticPrecondition);
            }
        }
        for planned in graph.parts.iter().rev() {
            let name = PackURI::new(planned.name.clone()).map_err(crate::Error::Invalid)?;
            let part = package.get_part(&name).map_err(crate::Error::from)?;
            if part.content_type() != planned.content_type
                || part.blob() != planned.blob.as_slice()
                || !relationships_match(part.rels(), &planned.relationships)
            {
                return Err(TransactionError::SemanticPrecondition);
            }
            if !package.remove_part(&name) {
                return Err(TransactionError::SemanticPrecondition);
            }
        }
    }
    Ok(())
}

fn add_relationships(
    relationships: &mut litchi_opc::Relationships,
    planned: &[TransferRelationship],
) -> Result<(), TransactionError> {
    for relationship in planned {
        relationships
            .try_add_relationship(
                relationship.relationship_type.clone(),
                relationship.target.clone(),
                relationship.id.clone(),
                if relationship.external {
                    TargetMode::External
                } else {
                    TargetMode::Internal
                },
            )
            .map_err(crate::Error::from)?;
    }
    Ok(())
}

fn relationships_match(
    relationships: &litchi_opc::Relationships,
    planned: &[TransferRelationship],
) -> bool {
    relationships.len() == planned.len()
        && planned.iter().all(|expected| {
            relationships.get(&expected.id).is_some_and(|actual| {
                actual.reltype() == expected.relationship_type
                    && actual.target_ref() == expected.target
                    && actual.is_external() == expected.external
            })
        })
}

pub(super) fn relationship_graph_digest(package: &Package) -> Result<String, TransactionError> {
    relationship_graph_digest_opc(&package.opc)
}

pub(super) fn relationship_graph_digest_opc(
    package: &OpcPackage,
) -> Result<String, TransactionError> {
    let main_name = package
        .main_document_part()
        .map_err(crate::Error::from)?
        .partname()
        .clone();
    let mut parts = package.iter_parts().collect::<Vec<_>>();
    parts.sort_by(|left, right| left.partname().as_str().cmp(right.partname().as_str()));
    let mut bytes = Vec::new();
    for part in parts {
        put_digest_field(&mut bytes, part.partname().as_str().as_bytes())?;
        put_digest_field(&mut bytes, part.content_type().as_bytes())?;
        if part.partname() != &main_name {
            put_digest_field(
                &mut bytes,
                litchi_core::patch::BlobId::of(part.blob())
                    .as_hex()
                    .as_bytes(),
            )?;
        }
        let mut relationships = part.rels().iter().collect::<Vec<_>>();
        relationships.sort_by(|left, right| left.r_id().cmp(right.r_id()));
        for relationship in relationships {
            put_digest_field(&mut bytes, relationship.r_id().as_bytes())?;
            put_digest_field(&mut bytes, relationship.reltype().as_bytes())?;
            put_digest_field(&mut bytes, relationship.target_ref().as_bytes())?;
            bytes.push(u8::from(relationship.is_external()));
        }
    }
    Ok(litchi_core::patch::BlobId::of(&bytes).as_hex())
}

fn put_digest_field(output: &mut Vec<u8>, value: &[u8]) -> Result<(), TransactionError> {
    let length = u64::try_from(value.len()).map_err(|_error| TransactionError::Limit {
        resource: "transfer dependency field bytes",
        max: usize::MAX,
        actual: value.len(),
    })?;
    output.extend_from_slice(&length.to_le_bytes());
    output.extend_from_slice(value);
    Ok(())
}

fn relationship_prefixes(xml: &[u8]) -> Result<BTreeSet<Vec<u8>>, TransactionError> {
    let mut reader = NsReader::from_reader(xml);
    let mut bindings = BTreeMap::<Vec<u8>, Option<bool>>::new();
    loop {
        match reader
            .read_event()
            .map_err(|error| crate::Error::Xml(error.to_string()))?
        {
            Event::Start(element) | Event::Empty(element) => {
                for attribute_result in element.attributes() {
                    let attribute =
                        attribute_result.map_err(|error| crate::Error::Xml(error.to_string()))?;
                    let Some(prefix) = attribute.key.as_ref().strip_prefix(b"xmlns:") else {
                        continue;
                    };
                    let value = attribute
                        .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
                        .map_err(|error| crate::Error::Xml(error.to_string()))?;
                    let relationship = value.as_bytes() == TRANSITIONAL_RELATIONSHIPS_NAMESPACE
                        || value.as_bytes() == STRICT_RELATIONSHIPS_NAMESPACE;
                    bindings
                        .entry(prefix.to_vec())
                        .and_modify(|state| {
                            if state.is_some_and(|existing| existing != relationship) {
                                *state = None;
                            }
                        })
                        .or_insert(Some(relationship));
                }
            },
            Event::Eof => break,
            Event::DocType(_) => {
                return Err(TransactionError::Transfer(
                    TransferRefusal::InvalidParagraphXml,
                ));
            },
            Event::End(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::GeneralRef(_) => {},
        }
    }
    Ok(bindings
        .into_iter()
        .filter_map(|(prefix, relationship)| (relationship == Some(true)).then_some(prefix))
        .collect())
}

fn stable_namespace_declarations(xml: &[u8]) -> Result<BTreeMap<String, String>, TransactionError> {
    let mut reader = NsReader::from_reader(xml);
    let mut bindings = BTreeMap::<String, Option<String>>::new();
    loop {
        match reader
            .read_event()
            .map_err(|error| crate::Error::Xml(error.to_string()))?
        {
            Event::Start(element) | Event::Empty(element) => {
                for attribute_result in element.attributes() {
                    let attribute =
                        attribute_result.map_err(|error| crate::Error::Xml(error.to_string()))?;
                    let raw_key = attribute.key.as_ref();
                    if raw_key != b"xmlns" && !raw_key.starts_with(b"xmlns:") {
                        continue;
                    }
                    let key = std::str::from_utf8(raw_key)
                        .map_err(|_error| {
                            TransactionError::Transfer(TransferRefusal::InvalidParagraphXml)
                        })?
                        .to_owned();
                    let value = attribute
                        .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
                        .map_err(|error| crate::Error::Xml(error.to_string()))?
                        .into_owned();
                    bindings
                        .entry(key)
                        .and_modify(|state| {
                            if state
                                .as_deref()
                                .is_some_and(|existing| existing != value.as_str())
                            {
                                *state = None;
                            }
                        })
                        .or_insert(Some(value));
                }
            },
            Event::Eof => break,
            Event::DocType(_) => {
                return Err(TransactionError::Transfer(
                    TransferRefusal::InvalidParagraphXml,
                ));
            },
            Event::End(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::GeneralRef(_) => {},
        }
    }
    Ok(bindings
        .into_iter()
        .filter_map(|(key, candidate)| candidate.map(|namespace| (key, namespace)))
        .collect())
}

fn relationship_references(
    xml: &[u8],
    relationship_prefixes: &BTreeSet<Vec<u8>>,
) -> Result<BTreeSet<String>, TransactionError> {
    let mut reader = NsReader::from_reader(xml);
    let mut identifiers = BTreeSet::new();
    loop {
        let event = reader
            .read_event()
            .map_err(|error| crate::Error::Xml(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        match event {
            Event::Start(element) | Event::Empty(element) => {
                for attribute_result in element.attributes() {
                    let attribute =
                        attribute_result.map_err(|error| crate::Error::Xml(error.to_string()))?;
                    if is_relationship_reference(&resolver, attribute.key, relationship_prefixes) {
                        identifiers.insert(
                            attribute
                                .decoded_and_normalized_value(
                                    XmlVersion::Explicit1_0,
                                    reader.decoder(),
                                )
                                .map_err(|error| crate::Error::Xml(error.to_string()))?
                                .into_owned(),
                        );
                    } else if is_unresolved_reference(&resolver, attribute.key) {
                        return Err(TransactionError::Transfer(
                            TransferRefusal::InvalidParagraphXml,
                        ));
                    }
                }
            },
            Event::DocType(_) => {
                return Err(TransactionError::Transfer(
                    TransferRefusal::InvalidParagraphXml,
                ));
            },
            Event::Eof => break,
            Event::End(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::GeneralRef(_) => {},
        }
    }
    Ok(identifiers)
}

fn rewrite_relationship_references(
    xml: &[u8],
    mapping: &BTreeMap<String, String>,
    relationship_prefixes: &BTreeSet<Vec<u8>>,
    inherited_namespaces: &BTreeMap<String, String>,
) -> Result<Vec<u8>, TransactionError> {
    let mut reader = NsReader::from_reader(xml);
    let mut writer = Writer::new(Vec::with_capacity(xml.len()));
    let mut root_written = false;
    loop {
        let event = reader
            .read_event()
            .map_err(|error| crate::Error::Xml(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let is_element = matches!(&event, Event::Start(_) | Event::Empty(_));
        match event {
            Event::Start(element) => writer
                .write_event(Event::Start(rewrite_element(
                    &element,
                    &resolver,
                    reader.decoder(),
                    mapping,
                    relationship_prefixes,
                    (!root_written).then_some(inherited_namespaces),
                )?))
                .map_err(crate::Error::from)?,
            Event::Empty(element) => writer
                .write_event(Event::Empty(rewrite_element(
                    &element,
                    &resolver,
                    reader.decoder(),
                    mapping,
                    relationship_prefixes,
                    (!root_written).then_some(inherited_namespaces),
                )?))
                .map_err(crate::Error::from)?,
            Event::Eof => break,
            Event::DocType(_) | Event::Decl(_) => {
                return Err(TransactionError::Transfer(
                    TransferRefusal::InvalidParagraphXml,
                ));
            },
            Event::End(element) => writer
                .write_event(Event::End(element))
                .map_err(crate::Error::from)?,
            Event::Text(content) => writer
                .write_event(Event::Text(content))
                .map_err(crate::Error::from)?,
            Event::CData(content) => writer
                .write_event(Event::CData(content))
                .map_err(crate::Error::from)?,
            Event::Comment(comment) => writer
                .write_event(Event::Comment(comment))
                .map_err(crate::Error::from)?,
            Event::PI(instruction) => writer
                .write_event(Event::PI(instruction))
                .map_err(crate::Error::from)?,
            Event::GeneralRef(reference) => writer
                .write_event(Event::GeneralRef(reference))
                .map_err(crate::Error::from)?,
        }
        if is_element {
            root_written = true;
        }
    }
    Ok(writer.into_inner())
}

fn rewrite_element(
    source: &BytesStart<'_>,
    resolver: &NamespaceResolver,
    decoder: Decoder,
    mapping: &BTreeMap<String, String>,
    relationship_prefixes: &BTreeSet<Vec<u8>>,
    inherited_namespaces: Option<&BTreeMap<String, String>>,
) -> Result<BytesStart<'static>, TransactionError> {
    let name = std::str::from_utf8(source.name().as_ref())
        .map_err(|_error| TransactionError::Transfer(TransferRefusal::InvalidParagraphXml))?
        .to_owned();
    let mut rewritten = BytesStart::new(name);
    let mut present = BTreeSet::new();
    for attribute_result in source.attributes() {
        let attribute = attribute_result.map_err(|error| crate::Error::Xml(error.to_string()))?;
        let key = std::str::from_utf8(attribute.key.as_ref())
            .map_err(|_error| TransactionError::Transfer(TransferRefusal::InvalidParagraphXml))?
            .to_owned();
        present.insert(key.clone());
        let attribute_value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| crate::Error::Xml(error.to_string()))?
            .into_owned();
        let output_value =
            if is_relationship_reference(resolver, attribute.key, relationship_prefixes) {
                mapping.get(&attribute_value).ok_or_else(|| {
                    TransactionError::Transfer(TransferRefusal::MissingDonorRelationship(
                        attribute_value.clone(),
                    ))
                })?
            } else if is_unresolved_reference(resolver, attribute.key) {
                return Err(TransactionError::Transfer(
                    TransferRefusal::InvalidParagraphXml,
                ));
            } else {
                &attribute_value
            };
        rewritten.push_attribute((key.as_str(), output_value.as_str()));
    }
    if let Some(namespace_declarations) = inherited_namespaces {
        for (key, value) in namespace_declarations {
            if !present.contains(key) {
                rewritten.push_attribute((key.as_str(), value.as_str()));
            }
        }
    }
    Ok(rewritten.into_owned())
}

fn is_unresolved_reference(resolver: &NamespaceResolver, name: quick_xml::name::QName<'_>) -> bool {
    let (namespace, local) = resolver.resolve_attribute(name);
    matches!(local.as_ref(), b"id" | b"embed" | b"link")
        && matches!(namespace, ResolveResult::Unknown(_))
}

fn is_relationship_reference(
    resolver: &NamespaceResolver,
    name: quick_xml::name::QName<'_>,
    relationship_prefixes: &BTreeSet<Vec<u8>>,
) -> bool {
    let (namespace, local) = resolver.resolve_attribute(name);
    matches!(local.as_ref(), b"id" | b"embed" | b"link")
        && match namespace {
            ResolveResult::Bound(Namespace(uri)) => {
                uri == TRANSITIONAL_RELATIONSHIPS_NAMESPACE || uri == STRICT_RELATIONSHIPS_NAMESPACE
            },
            ResolveResult::Unknown(prefix) => relationship_prefixes.contains(prefix.as_slice()),
            ResolveResult::Unbound => false,
        }
}
