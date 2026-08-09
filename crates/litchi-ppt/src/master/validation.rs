//! MS-PPT identity and owner-boundary validation.

use crate::consts::RecordType;
use crate::package::{Error, Result};
use crate::persist::PersistMapping;
use crate::records::Record;

use super::model::{Identity, Kind, Persist};

pub(super) const DOCUMENT_ATOM_SIZE: usize = 40;
pub(super) const MASTER_PERSIST_SIZE: usize = 20;
pub(super) const SLIDE_ATOM_SIZE: usize = 24;
pub(super) const NOTES_ATOM_SIZE: usize = 8;

pub(super) fn header(
    record: &Record,
    expected_type: RecordType,
    version: u16,
    instance: u16,
    expected_data_len: Option<usize>,
) -> Result<()> {
    if record.record_type != expected_type
        || record.record_type_raw != expected_type.as_u16()
        || record.version != version
        || record.instance != instance
    {
        return Err(Error::Corrupted(format!(
            "invalid {:?} record header: type={:?}/{:#06x}, version={}, instance={}",
            expected_type,
            record.record_type,
            record.record_type_raw,
            record.version,
            record.instance
        )));
    }
    let actual = usize::try_from(record.data_length).map_err(|_err| {
        Error::Corrupted(format!(
            "{expected_type:?} record length does not fit usize"
        ))
    })?;
    if actual != record.data.len() {
        return Err(Error::Corrupted(format!(
            "{expected_type:?} record length does not match its payload"
        )));
    }
    if let Some(expected) = expected_data_len
        && actual != expected
    {
        return Err(Error::Corrupted(format!(
            "{expected_type:?} record has {actual} payload bytes; expected {expected}"
        )));
    }
    Ok(())
}

pub(super) fn identity(persist: Persist, master_id: u32) -> Result<Identity> {
    if persist.id() == 0 {
        return Err(Error::InvalidFormat(
            "master persist reference must not be null".to_string(),
        ));
    }
    if master_id < 0x8000_0000 {
        return Err(Error::InvalidFormat(format!(
            "master identifier {master_id:#010x} is below the MS-PPT master range"
        )));
    }
    Ok(Identity::new(persist, master_id))
}

pub(super) fn persist<'a>(
    id: u32,
    mapping: &PersistMapping,
    objects: &super::model::Objects<'a>,
    expected: Kind,
) -> Result<(Persist, &'a Record)> {
    if id == 0 {
        return Err(Error::InvalidFormat(format!(
            "{expected:?} persist reference is null"
        )));
    }
    let offset = mapping.get_offset(id).ok_or_else(|| {
        Error::Corrupted(format!(
            "{expected:?} persist identifier {id} has no directory entry"
        ))
    })?;
    let record = objects.resolve(id).ok_or_else(|| {
        Error::Corrupted(format!(
            "{expected:?} persist identifier {id} has no parsed object"
        ))
    })?;
    Ok((Persist::new(id, offset), record))
}

pub(super) fn document(document: &Record) -> Result<()> {
    header(document, RecordType::Document, 0x0f, 0, None)
}

pub(super) fn document_atom(record: &Record) -> Result<(u32, u32)> {
    header(
        record,
        RecordType::DocumentAtom,
        1,
        0,
        Some(DOCUMENT_ATOM_SIZE),
    )?;
    let notes = read_u32(record, 24, "DocumentAtom.notesMasterPersistIdRef")?;
    let handout = read_u32(record, 28, "DocumentAtom.handoutMasterPersistIdRef")?;
    Ok((notes, handout))
}

pub(super) fn master_list(record: &Record) -> Result<()> {
    header(record, RecordType::SlideListWithText, 0x0f, 1, None)
}

pub(super) fn master_persist(record: &Record) -> Result<(u32, u32)> {
    header(
        record,
        RecordType::SlidePersistAtom,
        0,
        0,
        Some(MASTER_PERSIST_SIZE),
    )?;
    let persist_id = read_u32(record, 0, "MasterPersistAtom.persistIdRef")?;
    let flags = read_u32(record, 4, "MasterPersistAtom.flags")?;
    let reserved = read_u32(record, 8, "MasterPersistAtom.reserved3")?;
    let master_id = read_u32(record, 12, "MasterPersistAtom.masterId")?;
    let reserved4 = read_u32(record, 16, "MasterPersistAtom.reserved4")?;
    if flags & !0x0000_0004 != 0 {
        return Err(Error::InvalidFormat(format!(
            "MasterPersistAtom has unknown flag bits {:#010x}",
            flags & !0x0000_0004
        )));
    }
    if reserved != 0 || reserved4 != 0 {
        return Err(Error::InvalidFormat(
            "MasterPersistAtom contains non-zero reserved fields".to_string(),
        ));
    }
    Ok((persist_id, master_id))
}

pub(super) fn slide_atom(record: &Record, owner: Kind) -> Result<(u32, u32)> {
    header(record, RecordType::SlideAtom, 2, 0, Some(SLIDE_ATOM_SIZE))?;
    let master = read_u32(record, 12, "SlideAtom.masterIdRef")?;
    let notes = read_u32(record, 16, "SlideAtom.notesIdRef")?;
    if owner == Kind::Main && (master != 0 || notes != 0) {
        return Err(Error::InvalidFormat(
            "MainMaster SlideAtom must have null master and notes references".to_string(),
        ));
    }
    if owner == Kind::Title && master == 0 {
        return Err(Error::InvalidFormat(
            "title master SlideAtom must reference a master identity".to_string(),
        ));
    }
    Ok((master, notes))
}

pub(super) fn notes(record: &Record) -> Result<()> {
    header(record, RecordType::Notes, 0x0f, 0, None)?;
    let atom = record
        .find_child(RecordType::NotesAtom)
        .ok_or_else(|| Error::Corrupted("Notes master is missing NotesAtom".to_string()))?;
    header(atom, RecordType::NotesAtom, 1, 0, Some(NOTES_ATOM_SIZE))?;
    let slide_id = read_u32(atom, 0, "NotesAtom.slideIdRef")?;
    if slide_id != 0 {
        return Err(Error::InvalidFormat(
            "Notes master NotesAtom.slideIdRef must be null".to_string(),
        ));
    }
    Ok(())
}

pub(super) fn handout(record: &Record) -> Result<()> {
    header(record, RecordType::Handout, 0x0f, 0, None)
}

pub(super) fn read_u32(record: &Record, offset: usize, name: &str) -> Result<u32> {
    let end = offset
        .checked_add(4)
        .ok_or_else(|| Error::Corrupted(format!("{name} offset overflow")))?;
    let bytes = record
        .data
        .get(offset..end)
        .ok_or_else(|| Error::Corrupted(format!("{name} is truncated")))?;
    Ok(u32::from_le_bytes(bytes.try_into().map_err(|_err| {
        Error::Corrupted(format!("{name} is truncated"))
    })?))
}

pub(super) fn wire_size(record: &Record) -> Result<usize> {
    let body = usize::try_from(record.data_length)
        .map_err(|_err| Error::Corrupted("record length does not fit usize".to_string()))?;
    if body != record.data.len() {
        return Err(Error::Corrupted(
            "record length does not match its payload".to_string(),
        ));
    }
    body.checked_add(8)
        .ok_or_else(|| Error::Corrupted("record wire size overflow".to_string()))
}
