//! Master-list discovery and composition over existing PPT records.

use crate::consts::RecordType;
use crate::package::{Error, Result};
use crate::persist::PersistMapping;
use crate::presentation::Presentation;
use crate::records::Record;

use super::model::{
    Handout, Identity, Inventory, Main, Notes, Objects, RecordRef, Scope, Title, Unknown,
};
use super::validation;

/// Parse a master inventory from a parsed Document record and a persist object
/// catalog.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn parse<'a>(
    document: &'a Record,
    objects: &Objects<'a>,
    mapping: &PersistMapping,
) -> Result<Inventory<'a>> {
    validation::document(document)?;

    let document_atoms: Vec<&Record> = document
        .children
        .iter()
        .filter(|record| record.record_type == RecordType::DocumentAtom)
        .collect();
    if document_atoms.len() != 1 {
        return Err(Error::Corrupted(format!(
            "DocumentContainer must contain exactly one DocumentAtom; found {}",
            document_atoms.len()
        )));
    }
    let (notes_id, handout_id) = validation::document_atom(document_atoms[0])?;

    let lists: Vec<&Record> = document
        .children
        .iter()
        .filter(|record| {
            record.record_type == RecordType::SlideListWithText && record.instance == 1
        })
        .collect();
    if lists.len() != 1 {
        return Err(Error::Corrupted(format!(
            "DocumentContainer must contain exactly one master list; found {}",
            lists.len()
        )));
    }
    let list = lists[0];
    validation::master_list(list)?;
    validate_children_cover(list)?;

    let mut unknown = unknown_document_children(document);
    let mut references: Vec<(&Record, Identity)> = Vec::new();
    for child in &list.children {
        if child.record_type != RecordType::SlidePersistAtom {
            unknown.push(Unknown::new(Scope::List, child));
            continue;
        }
        let (persist_id, master_id) = validation::master_persist(child)?;
        let persist =
            validation::persist(persist_id, mapping, objects, super::model::Kind::Main)?.0;
        let identity = validation::identity(persist, master_id)?;
        if references.iter().any(|(_, existing)| *existing == identity) {
            return Err(Error::InvalidFormat(format!(
                "duplicate master identity {master_id:#010x}"
            )));
        }
        references.push((child, identity));
    }

    let mut main = Vec::new();
    let mut title = Vec::new();
    let identities = references
        .iter()
        .map(|(_, identity)| *identity)
        .collect::<Vec<_>>();

    for (reference, identity) in references {
        let source = objects.resolve(identity.persist().id()).ok_or_else(|| {
            Error::Corrupted(format!(
                "master persist identifier {} has no parsed object",
                identity.persist().id()
            ))
        })?;
        #[allow(
            clippy::wildcard_enum_match_arm,
            reason = "`RecordType` mirrors the full MS-PPT record-type enumeration; only \
                      MainMaster and Slide are valid master persist targets and every other \
                      record type is rejected uniformly"
        )]
        let kind = match source.record_type {
            RecordType::MainMaster => super::model::Kind::Main,
            RecordType::Slide => super::model::Kind::Title,
            other => {
                return Err(Error::InvalidFormat(format!(
                    "master persist identifier {} resolves to {:?}, not MainMaster or Slide",
                    identity.persist().id(),
                    other
                )));
            },
        };
        let source_ref = RecordRef::new(identity.persist(), source);
        match kind {
            super::model::Kind::Main => {
                let atom = source.find_child(RecordType::SlideAtom).ok_or_else(|| {
                    Error::Corrupted("MainMaster is missing SlideAtom".to_string())
                })?;
                validation::slide_atom(atom, kind)?;
                main.push(Main::new(
                    identity,
                    source_ref,
                    reference,
                    unknown_children(source, Scope::Main),
                ));
            },
            super::model::Kind::Title => {
                let atom = source.find_child(RecordType::SlideAtom).ok_or_else(|| {
                    Error::Corrupted("title master is missing SlideAtom".to_string())
                })?;
                let (master_ref, _) = validation::slide_atom(atom, kind)?;
                let based_on = identities
                    .iter()
                    .copied()
                    .find(|candidate| candidate.master_id() == master_ref)
                    .ok_or_else(|| {
                        Error::InvalidFormat(format!(
                            "title master {} references unknown master identity {master_ref:#010x}",
                            identity.master_id()
                        ))
                    })?;
                title.push(Title::new(
                    identity,
                    source_ref,
                    reference,
                    based_on,
                    unknown_children(source, Scope::Title),
                ));
            },
            super::model::Kind::Notes | super::model::Kind::Handout => {
                return Err(Error::Corrupted(
                    "master list contains a non-slide master entry".into(),
                ));
            },
        }
    }

    let notes = if notes_id == 0 {
        None
    } else {
        let (persist, source) =
            validation::persist(notes_id, mapping, objects, super::model::Kind::Notes)?;
        if source.record_type != RecordType::Notes {
            return Err(Error::InvalidFormat(format!(
                "notes master persist identifier {notes_id} resolves to {:?}",
                source.record_type
            )));
        }
        validation::notes(source)?;
        Some(Notes::new(
            RecordRef::new(persist, source),
            unknown_children(source, Scope::Notes),
        ))
    };

    let handout = if handout_id == 0 {
        None
    } else {
        let (persist, source) =
            validation::persist(handout_id, mapping, objects, super::model::Kind::Handout)?;
        if source.record_type != RecordType::Handout {
            return Err(Error::InvalidFormat(format!(
                "handout master persist identifier {handout_id} resolves to {:?}",
                source.record_type
            )));
        }
        validation::handout(source)?;
        Some(Handout::new(
            RecordRef::new(persist, source),
            unknown_children(source, Scope::Handout),
        ))
    };

    Ok(Inventory::new(main, title, notes, handout, unknown))
}

/// Discover the Document record from the existing Presentation parser and
/// validate all persist references against its current mapping.
///
/// Parsed persist objects are supplied explicitly because `RecordParser` exposes
/// references but not physical stream offsets. This keeps the owner zero-copy
/// and lets encrypted/live presentations use the same source contract.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn parse_presentation<'a>(
    presentation: &'a Presentation,
    objects: &Objects<'a>,
) -> Result<Inventory<'a>> {
    let documents: Vec<&Record> = presentation
        .parser
        .find_records_ref()
        .into_iter()
        .filter(|record| record.record_type == RecordType::Document)
        .collect();
    if documents.len() != 1 {
        return Err(Error::Corrupted(format!(
            "presentation must contain exactly one DocumentContainer; found {}",
            documents.len()
        )));
    }
    parse(documents[0], objects, &presentation.persist_mapping)
}

fn validate_children_cover(record: &Record) -> Result<()> {
    let mut encoded = 0usize;
    for child in &record.children {
        encoded = encoded
            .checked_add(validation::wire_size(child)?)
            .ok_or_else(|| Error::Corrupted("master list child size overflow".to_string()))?;
    }
    if encoded != record.data.len() {
        return Err(Error::Corrupted(
            "master list children do not cover its original payload".to_string(),
        ));
    }
    Ok(())
}

fn unknown_document_children(record: &Record) -> Vec<Unknown<'_>> {
    record
        .children
        .iter()
        .filter(|child| {
            child.record_type == RecordType::Unknown
                && child.record_type != RecordType::DocumentAtom
                && child.record_type != RecordType::SlideListWithText
        })
        .map(|child| Unknown::new(Scope::Document, child))
        .collect()
}

fn unknown_children(record: &Record, scope: Scope) -> Vec<Unknown<'_>> {
    record
        .children
        .iter()
        .filter(|child| child.record_type == RecordType::Unknown)
        .map(|child| Unknown::new(scope, child))
        .collect()
}
