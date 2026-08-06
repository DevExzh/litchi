//! Cross-record invariants for external-media ownership and identifier use.

use crate::consts::RecordType;
use crate::embedded::reference::Reference;
use crate::hyperlink::Hyperlinks;
use crate::package::{Error, Result};
use crate::records::Record;
use std::collections::HashSet;

use super::model::{Collection, Limits};
use super::package::{self, Location};

/// All contextual state captured at the root boundary.
#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct ValidatedSource {
    pub(crate) collection: Option<Collection>,
    pub(crate) location: Option<Location>,
    pub(crate) root_type: RecordType,
    pub(crate) owner_ids: Vec<u32>,
    pub(crate) hyperlink_ids: Vec<u32>,
}

/// Parse and validate one complete source record tree.
pub(crate) fn parse_source(bytes: &[u8], limits: Limits) -> Result<ValidatedSource> {
    validate_limits(limits)?;
    if bytes.len() > limits.max_root_bytes {
        return Err(Error::InvalidFormat(
            "PowerPoint external-media source exceeds its byte limit".into(),
        ));
    }
    if bytes.len() < 8 {
        return Err(Error::Corrupted(
            "PowerPoint external-media source is missing its root record header".into(),
        ));
    }
    let (root, consumed) = Record::parse(bytes, 0)?;
    if consumed != bytes.len() {
        return Err(Error::Corrupted(
            "PowerPoint external-media source contains trailing bytes".into(),
        ));
    }
    let mut record_count = 0usize;
    validate_record(&root, 0, limits, &mut record_count)?;

    let collection = Collection::parse_media_only(&root)?;
    let hyperlinks = Hyperlinks::parse(&root)?;
    let hyperlink_ids = hyperlinks
        .hyperlinks
        .iter()
        .map(|hyperlink| hyperlink.id)
        .collect::<Vec<_>>();
    if let Some(collection) = &collection {
        if hyperlink_ids.iter().any(|id| collection.get(*id).is_some()) {
            return Err(Error::Corrupted(
                "external-object list reuses an ID for media and a hyperlink".into(),
            ));
        }
        validate_collection(collection, &hyperlink_ids)?;
    }

    let mut owner_ids = Vec::new();
    let mut owner_record_count = 0usize;
    collect_owner_ids(&root, limits, &mut owner_ids, &mut owner_record_count)?;
    Ok(ValidatedSource {
        collection,
        location: package::locate(&root)?,
        root_type: root.record_type,
        owner_ids,
        hyperlink_ids,
    })
}

/// Check limits before parsing or publishing a candidate.
pub(crate) fn validate_limits(limits: Limits) -> Result<()> {
    if limits.max_root_bytes < 8
        || limits.max_record_bytes < 8
        || limits.max_depth == 0
        || limits.max_records == 0
        || limits.max_owner_references == 0
    {
        return Err(Error::InvalidFormat(
            "external-media resource limits must be positive".into(),
        ));
    }
    Ok(())
}

/// Validate a candidate list against the document's hyperlink namespace.
pub(crate) fn validate_collection(collection: &Collection, hyperlink_ids: &[u32]) -> Result<()> {
    if collection.id_seed == 0 || collection.id_seed > i32::MAX as u32 {
        return Err(Error::Corrupted(
            "ExObjListAtom identifier seed must be a positive i32".into(),
        ));
    }
    if collection.objects.len() > super::codec::MAX_EXTERNAL_MEDIA_OBJECTS {
        return Err(Error::InvalidFormat(format!(
            "external-object list exceeds {} media objects",
            super::codec::MAX_EXTERNAL_MEDIA_OBJECTS
        )));
    }

    let mut ids = HashSet::with_capacity(collection.objects.len());
    for object in &collection.objects {
        let id = object.id();
        if id == 0 || id > collection.id_seed {
            return Err(Error::Corrupted(format!(
                "external media ID {id} is outside ExObjList seed {}",
                collection.id_seed
            )));
        }
        if !ids.insert(id) {
            return Err(Error::Corrupted(format!(
                "external-object list contains duplicate media ID {id}"
            )));
        }
        if hyperlink_ids.contains(&id) {
            return Err(Error::Corrupted(format!(
                "external-object list reuses an ID for media and a hyperlink: {id}"
            )));
        }
        object.to_record_bytes()?;
    }

    let mut previous_index = 0usize;
    for record in &collection.unknown_records {
        if record.object_index > collection.objects.len() || record.object_index < previous_index {
            return Err(Error::Corrupted(
                "opaque external-object records are out of source order".into(),
            ));
        }
        record.to_record_bytes()?;
        previous_index = record.object_index;
    }
    Ok(())
}

fn validate_record(
    record: &Record,
    depth: usize,
    limits: Limits,
    record_count: &mut usize,
) -> Result<()> {
    if depth > limits.max_depth {
        return Err(Error::InvalidFormat(
            "PowerPoint external-media record nesting exceeds its limit".into(),
        ));
    }
    *record_count = record_count
        .checked_add(1)
        .ok_or_else(|| Error::Corrupted("PowerPoint record count overflows usize".into()))?;
    if *record_count > limits.max_records {
        return Err(Error::InvalidFormat(
            "PowerPoint external-media record count exceeds its limit".into(),
        ));
    }
    if record.data.len() > limits.max_record_bytes
        || usize::try_from(record.data_length).ok() != Some(record.data.len())
    {
        return Err(Error::Corrupted(
            "PowerPoint record payload is truncated or exceeds its limit".into(),
        ));
    }
    for child in &record.children {
        validate_record(child, depth.saturating_add(1), limits, record_count)?;
    }
    Ok(())
}

fn collect_owner_ids(
    record: &Record,
    limits: Limits,
    owner_ids: &mut Vec<u32>,
    record_count: &mut usize,
) -> Result<()> {
    visit_record_count(record_count, limits)?;
    if record.record_type == RecordType::ExternalObjectRefAtom {
        let reference = Reference::parse(record)?;
        if owner_ids.len() >= limits.max_owner_references {
            return Err(Error::InvalidFormat(
                "PowerPoint external-media owner references exceed their limit".into(),
            ));
        }
        owner_ids.push(reference.id);
    }
    if record.record_type == RecordType::PPDrawing {
        collect_officeart_owner_ids(&record.data, limits, owner_ids, record_count, 0)?;
    }
    for child in &record.children {
        collect_owner_ids(child, limits, owner_ids, record_count)?;
    }
    Ok(())
}

fn collect_officeart_owner_ids(
    data: &[u8],
    limits: Limits,
    owner_ids: &mut Vec<u32>,
    record_count: &mut usize,
    depth: usize,
) -> Result<()> {
    for record in litchi_odraw::Parser::new(data).records() {
        let record = record?;
        collect_officeart_record(&record, limits, owner_ids, record_count, depth)?;
    }
    Ok(())
}

fn collect_officeart_record(
    record: &litchi_odraw::Record<'_>,
    limits: Limits,
    owner_ids: &mut Vec<u32>,
    record_count: &mut usize,
    depth: usize,
) -> Result<()> {
    if depth > limits.max_depth {
        return Err(Error::InvalidFormat(
            "PowerPoint OfficeArt owner nesting exceeds its limit".into(),
        ));
    }
    visit_record_count(record_count, limits)?;
    if record.kind() == litchi_odraw::RecordKind::ClientData {
        for child in Record::parse_sequence_strict(record.data(), "OfficeArt ClientData")? {
            collect_owner_ids(&child, limits, owner_ids, record_count)?;
        }
    }
    if record.is_container() {
        let container = litchi_odraw::Container::try_new(record.clone())?;
        for child in container.children() {
            let child = child?;
            collect_officeart_record(
                &child,
                limits,
                owner_ids,
                record_count,
                depth.saturating_add(1),
            )?;
        }
    }
    Ok(())
}

fn visit_record_count(record_count: &mut usize, limits: Limits) -> Result<()> {
    *record_count = record_count
        .checked_add(1)
        .ok_or_else(|| Error::Corrupted("PowerPoint record count overflows usize".into()))?;
    if *record_count > limits.max_records {
        return Err(Error::InvalidFormat(
            "PowerPoint external-media record count exceeds its limit".into(),
        ));
    }
    Ok(())
}

/// Ensure a media removal does not leave a known shape owner dangling.
pub(crate) fn can_remove(collection: &Collection, owner_ids: &[u32], id: u32) -> Result<()> {
    if collection.get(id).is_none() {
        return Err(Error::InvalidFormat(format!(
            "external media ID {id} was not found"
        )));
    }
    if owner_ids.contains(&id) {
        return Err(Error::InvalidFormat(format!(
            "external media ID {id} is still owned by an ExObjRefAtom"
        )));
    }
    Ok(())
}

/// Return whether a candidate object ID is already occupied by a hyperlink or
/// an owner in the shared external-object namespace.
pub(crate) fn id_is_reserved(id: u32, hyperlink_ids: &[u32], owner_ids: &[u32]) -> bool {
    hyperlink_ids.contains(&id) || owner_ids.contains(&id)
}
