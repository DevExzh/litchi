//! Dependency-checked cross-package paragraph transfer planning.

use std::collections::{BTreeMap, BTreeSet};
use std::sync::Arc;

use litchi_core::Position;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;
use quick_xml::{Writer, XmlVersion};

use super::super::model::Package;
use crate::document::{ParagraphTransfer, TransactionError, TransferRefusal};

const TRANSITIONAL_RELATIONSHIPS_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_RELATIONSHIPS_NAMESPACE: &[u8] =
    b"http://purl.oclc.org/ooxml/officeDocument/relationships";

impl Package {
    /// Plan transfer of one direct-body paragraph from another package without
    /// mutating either package.
    ///
    /// Every relationship reference is resolved in the donor and mapped only
    /// when the receiver already owns exactly one equivalent relationship edge
    /// and target resource. Missing or ambiguous dependency closure is refused
    /// explicitly. The resulting compact plan belongs to this exact receiving
    /// main document and can be joined with ordinary document edits.
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
        let receiver_main = self.opc.main_document_part().map_err(crate::Error::from)?;
        let relationship_prefixes = relationship_prefixes(donor_snapshot.xml_bytes())?;
        let inherited_namespaces = stable_namespace_declarations(donor_snapshot.xml_bytes())?;
        let mut mapping = BTreeMap::new();
        for identifier in
            relationship_references(donor_paragraph.xml_bytes(), &relationship_prefixes)?
        {
            let dependency = donor_main.rels().get(&identifier).ok_or_else(|| {
                TransactionError::Transfer(TransferRefusal::MissingDonorRelationship(
                    identifier.clone(),
                ))
            })?;
            let donor_resource = if dependency.is_external() {
                None
            } else {
                let target = dependency.target_partname().map_err(crate::Error::from)?;
                Some(donor.opc.get_part(&target).map_err(crate::Error::from)?)
            };
            let mut equivalent = receiver_main.rels().iter().filter(|candidate| {
                candidate.reltype() == dependency.reltype()
                    && candidate.target_ref() == dependency.target_ref()
                    && candidate.target_mode() == dependency.target_mode()
                    && donor_resource.is_none_or(|donor_part| {
                        candidate
                            .target_partname()
                            .ok()
                            .and_then(|target| self.opc.get_part(&target).ok())
                            .is_some_and(|receiver_part| {
                                receiver_part.content_type() == donor_part.content_type()
                                    && receiver_part.blob() == donor_part.blob()
                            })
                    })
            });
            let selected = equivalent.next().ok_or_else(|| {
                TransactionError::Transfer(TransferRefusal::MissingEquivalentDependency {
                    relationship_type: dependency.reltype().to_owned(),
                    target: dependency.target_ref().to_owned(),
                })
            })?;
            if equivalent.next().is_some() {
                return Err(TransactionError::Transfer(
                    TransferRefusal::AmbiguousEquivalentDependency {
                        relationship_type: dependency.reltype().to_owned(),
                        target: dependency.target_ref().to_owned(),
                    },
                ));
            }
            mapping.insert(identifier, selected.r_id().to_owned());
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
        Ok(ParagraphTransfer::new(
            Arc::new(receiver.xml_bytes().to_vec()),
            compact.into_bytes(),
            relationship_graph_digest(self)?,
        ))
    }
}

pub(super) fn relationship_graph_digest(package: &Package) -> Result<String, TransactionError> {
    let main = package
        .opc
        .main_document_part()
        .map_err(crate::Error::from)?;
    let mut relationships = main.rels().iter().collect::<Vec<_>>();
    relationships.sort_by(|left, right| left.r_id().cmp(right.r_id()));
    let mut bytes = Vec::new();
    for relationship in relationships {
        put_digest_field(&mut bytes, relationship.r_id().as_bytes())?;
        put_digest_field(&mut bytes, relationship.reltype().as_bytes())?;
        put_digest_field(&mut bytes, relationship.target_ref().as_bytes())?;
        bytes.push(u8::from(relationship.is_external()));
        if !relationship.is_external() {
            let target = relationship.target_partname().map_err(crate::Error::from)?;
            let part = package.opc.get_part(&target).map_err(crate::Error::from)?;
            put_digest_field(&mut bytes, part.content_type().as_bytes())?;
            put_digest_field(
                &mut bytes,
                litchi_core::patch::BlobId::of(part.blob())
                    .as_hex()
                    .as_bytes(),
            )?;
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
