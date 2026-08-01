//! Transactional persisted-record rewrite for PowerPoint OLE objects.

use super::PptError;
use super::ole_object::{PowerPointOleExternalObject, PowerPointOleObjectCollection};
use super::ole_storage::PowerPointOleStorage;
use super::writer::{PersistPtrBuilder, UserEditAtom};
use crate::{OleFile, OleWriter};
use std::collections::{BTreeMap, HashMap, HashSet};
use std::io::Cursor;

type Result<T> = std::result::Result<T, PptError>;

const EX_OBJ_LIST: u16 = 1033;
const USER_EDIT: u16 = 4085;
const PERSIST_FULL: u16 = 6001;
const PERSIST_INCREMENTAL: u16 = 6002;

/// Appends a new PPT incremental edit; existing persisted bytes never move.
#[derive(Clone)]
pub struct PowerPointOlePackageEditor {
    original: Vec<u8>,
    streams: Vec<(Vec<String>, Vec<u8>)>,
    document_path: Vec<String>,
    current_user_path: Vec<String>,
    document: Vec<u8>,
    current_user: Vec<u8>,
    mappings: BTreeMap<u32, u32>,
    current_edit_offset: u32,
    document_persist_id: u32,
    collection: PowerPointOleObjectCollection,
    staged_storage: HashMap<u32, Vec<u8>>,
    removed_persist_ids: HashSet<u32>,
    rewrite_object_list: bool,
    changed: bool,
}

impl PowerPointOlePackageEditor {
    pub fn open(bytes: Vec<u8>, collection: PowerPointOleObjectCollection) -> Result<Self> {
        let mut ole = OleFile::open(Cursor::new(bytes.clone()))?;
        let paths = ole.list_streams();
        if paths.iter().flatten().any(|name| {
            matches!(
                name.to_ascii_lowercase().as_str(),
                "\u{6}dataspaces"
                    | "encryptioninfo"
                    | "encryptedpackage"
                    | "_signatures"
                    | "\u{5}digitalsignature"
            )
        }) {
            return Err(PptError::Corrupted(
                "signed or encrypted PPT is not eligible for OLE editing".into(),
            ));
        }
        let document_path = paths
            .iter()
            .find(|path| {
                path.last()
                    .is_some_and(|name| name == "PowerPoint Document")
            })
            .cloned()
            .ok_or_else(|| PptError::StreamNotFound("PowerPoint Document".into()))?;
        let current_user_path = paths
            .iter()
            .find(|path| path.last().is_some_and(|name| name == "Current User"))
            .cloned()
            .ok_or_else(|| PptError::StreamNotFound("Current User".into()))?;
        let document = ole.open_stream(&refs(&document_path))?;
        let current_user = ole.open_stream(&refs(&current_user_path))?;
        if current_user.len() < 28 || u32_at(&current_user, 12)? != 0xE391_C05F {
            return Err(PptError::Corrupted(
                "unsupported or encrypted CurrentUserAtom".into(),
            ));
        }
        let current_edit_offset = u32_at(&current_user, 16)?;
        let (mappings, document_persist_id) = mapping_chain(&document, current_edit_offset)?;
        collection.validate()?;
        for object in &collection.objects {
            if !mappings.contains_key(&object.persist_id()) {
                return Err(PptError::Corrupted(format!(
                    "OLE object references missing persist ID {}",
                    object.persist_id()
                )));
            }
        }
        let mut streams = Vec::new();
        for path in paths {
            let data = ole.open_stream(&refs(&path))?;
            streams.push((path, data));
        }
        Ok(Self {
            original: bytes,
            streams,
            document_path,
            current_user_path,
            document,
            current_user,
            mappings,
            current_edit_offset,
            document_persist_id,
            collection,
            staged_storage: HashMap::new(),
            removed_persist_ids: HashSet::new(),
            rewrite_object_list: false,
            changed: false,
        })
    }

    /// Opens the shared incremental persisted-record editor without requiring
    /// an external-object collection. Used by non-OLE record editors.
    pub fn open_records(bytes: Vec<u8>) -> Result<Self> {
        Self::open(
            bytes,
            PowerPointOleObjectCollection {
                id_seed: 1,
                objects: Vec::new(),
            },
        )
    }

    /// Live persisted identifiers in ascending order.
    pub fn persist_ids(&self) -> Vec<u32> {
        self.mappings.keys().copied().collect()
    }

    /// Returns one complete live persisted record.
    pub fn persisted_record(&self, persist_id: u32) -> Result<Vec<u8>> {
        if let Some(record) = self.staged_storage.get(&persist_id) {
            return Ok(record.clone());
        }
        let offset = *self
            .mappings
            .get(&persist_id)
            .ok_or_else(|| PptError::Corrupted(format!("unknown persist ID {persist_id}")))?
            as usize;
        Ok(record_slice(&self.document, offset)?.to_vec())
    }

    /// Stages one complete replacement record in the next incremental edit.
    pub fn replace_persisted_record(&mut self, persist_id: u32, record: Vec<u8>) -> Result<()> {
        if !self.mappings.contains_key(&persist_id)
            || self.removed_persist_ids.contains(&persist_id)
        {
            return Err(PptError::Corrupted(format!(
                "unknown persist ID {persist_id}"
            )));
        }
        if record.len() < 8
            || record.len() > 128 * 1024 * 1024
            || record_slice(&record, 0).map(|value| value.len())? != record.len()
        {
            return Err(PptError::Corrupted(
                "replacement persisted record has an invalid length".into(),
            ));
        }
        let mut candidate = self.clone();
        candidate.staged_storage.insert(persist_id, record);
        candidate.changed = true;
        *self = candidate;
        Ok(())
    }

    pub fn objects(&self) -> &PowerPointOleObjectCollection {
        &self.collection
    }

    pub fn add(
        &mut self,
        mut object: PowerPointOleExternalObject,
        storage: PowerPointOleStorage,
    ) -> Result<u32> {
        let persist_id = self.next_persist_id()?;
        set_persist_id(&mut object, persist_id);
        let mut candidate = self.clone();
        candidate.collection.add(object)?;
        candidate
            .staged_storage
            .insert(persist_id, storage.to_record_bytes()?);
        candidate.rewrite_object_list = true;
        candidate.changed = true;
        *self = candidate;
        Ok(persist_id)
    }

    pub fn replace_storage(
        &mut self,
        persist_id: u32,
        storage: PowerPointOleStorage,
    ) -> Result<()> {
        if !self
            .collection
            .objects
            .iter()
            .any(|object| object.persist_id() == persist_id)
        {
            return Err(PptError::Corrupted(
                "persist ID has no OLE object reference".into(),
            ));
        }
        self.staged_storage
            .insert(persist_id, storage.to_record_bytes()?);
        self.changed = true;
        Ok(())
    }

    pub fn remove(&mut self, id: u32) -> Result<PowerPointOleExternalObject> {
        let mut candidate = self.clone();
        let removed = candidate.collection.remove(id)?;
        let persist = removed.persist_id();
        if !candidate
            .collection
            .objects
            .iter()
            .any(|object| object.persist_id() == persist)
        {
            candidate.removed_persist_ids.insert(persist);
            candidate.staged_storage.remove(&persist);
        }
        candidate.rewrite_object_list = true;
        candidate.changed = true;
        *self = candidate;
        Ok(removed)
    }

    pub fn reorder(&mut self, ids: &[u32]) -> Result<()> {
        let mut candidate = self.clone();
        candidate.collection.reorder(ids)?;
        candidate.rewrite_object_list = true;
        candidate.changed = true;
        *self = candidate;
        Ok(())
    }

    pub fn finish(mut self) -> Result<Vec<u8>> {
        if !self.changed {
            return Ok(self.original);
        }
        for id in &self.removed_persist_ids {
            self.mappings.remove(id);
        }
        let mut appended = self.document.clone();
        for (id, record) in &self.staged_storage {
            self.mappings.insert(
                *id,
                u32::try_from(appended.len())
                    .map_err(|_| PptError::Corrupted("PPT stream exceeds u32".into()))?,
            );
            appended.extend_from_slice(record);
        }
        if self.rewrite_object_list {
            let document_offset = *self
                .mappings
                .get(&self.document_persist_id)
                .ok_or_else(|| PptError::Corrupted("Document persist mapping is missing".into()))?
                as usize;
            let old_document = record_slice(&self.document, document_offset)?;
            let new_document = replace_nested_record(
                old_document,
                EX_OBJ_LIST,
                &self.collection.to_record_bytes()?,
            )?;
            self.mappings.insert(
                self.document_persist_id,
                u32::try_from(appended.len())
                    .map_err(|_| PptError::Corrupted("PPT stream exceeds u32".into()))?,
            );
            appended.extend_from_slice(&new_document);
        }
        let persist_dir_offset = u32::try_from(appended.len())
            .map_err(|_| PptError::Corrupted("PPT stream exceeds u32".into()))?;
        let mut builder = PersistPtrBuilder::new();
        for (id, offset) in &self.mappings {
            builder.set_offset(*id, *offset);
        }
        appended.extend_from_slice(&builder.generate_full_record());
        let max_id = self
            .mappings
            .keys()
            .next_back()
            .copied()
            .unwrap_or(self.document_persist_id);
        let mut edit =
            UserEditAtom::new_minimal(persist_dir_offset, self.document_persist_id, max_id, 0);
        edit.offset_last_edit = self.current_edit_offset;
        let new_edit_offset = u32::try_from(appended.len())
            .map_err(|_| PptError::Corrupted("PPT stream exceeds u32".into()))?;
        appended.extend_from_slice(&edit.generate_record());
        self.current_user[16..20].copy_from_slice(&new_edit_offset.to_le_bytes());

        let mut writer = OleWriter::new();
        for (path, data) in &self.streams {
            let data = if path == &self.document_path {
                &appended
            } else if path == &self.current_user_path {
                &self.current_user
            } else {
                data
            };
            writer.create_stream(&refs(path), data)?;
        }
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output)?;
        let bytes = output.into_inner();
        let mut reopen = OleFile::open(Cursor::new(bytes.clone()))?;
        let doc = reopen.open_stream(&refs(&self.document_path))?;
        let cur = reopen.open_stream(&refs(&self.current_user_path))?;
        let (mapping, _) = mapping_chain(&doc, u32_at(&cur, 16)?)?;
        for object in &self.collection.objects {
            if !mapping.contains_key(&object.persist_id()) {
                return Err(PptError::Corrupted(
                    "rewritten persist mapping failed validation".into(),
                ));
            }
        }
        Ok(bytes)
    }

    fn next_persist_id(&self) -> Result<u32> {
        self.mappings
            .keys()
            .chain(self.staged_storage.keys())
            .copied()
            .max()
            .unwrap_or(0)
            .checked_add(1)
            .filter(|id| *id <= 0x000F_FFFF)
            .ok_or_else(|| PptError::Corrupted("persist ID space exhausted".into()))
    }
}

fn set_persist_id(object: &mut PowerPointOleExternalObject, persist_id: u32) {
    match object {
        PowerPointOleExternalObject::Object(value) => value.object.persist_id = persist_id,
        PowerPointOleExternalObject::ActiveXControl(value) => value.object.persist_id = persist_id,
    }
}

fn mapping_chain(document: &[u8], mut edit_offset: u32) -> Result<(BTreeMap<u32, u32>, u32)> {
    let mut mapping = BTreeMap::new();
    let mut seen = HashSet::new();
    let mut document_id = 0;
    while edit_offset != 0 {
        if !seen.insert(edit_offset) || seen.len() > 4_096 {
            return Err(PptError::Corrupted(
                "cyclic or excessive UserEdit chain".into(),
            ));
        }
        let record = record_slice(document, edit_offset as usize)?;
        if type_of(record)? != USER_EDIT || record.len() < 36 {
            return Err(PptError::Corrupted("invalid UserEditAtom".into()));
        }
        let data = &record[8..];
        if document_id == 0 {
            document_id = u32_at(data, 16)?;
        }
        let directory = record_slice(document, u32_at(data, 12)? as usize)?;
        if !matches!(type_of(directory)?, PERSIST_FULL | PERSIST_INCREMENTAL) {
            return Err(PptError::Corrupted("invalid PersistDirectoryAtom".into()));
        }
        let mut offset = 8usize;
        while offset < directory.len() {
            let info = u32_at(directory, offset)?;
            offset += 4;
            let base = info & 0x000F_FFFF;
            let count = info >> 20;
            if count == 0 {
                return Err(PptError::Corrupted("zero persist run".into()));
            }
            for index in 0..count {
                let value = u32_at(directory, offset)?;
                offset += 4;
                mapping.entry(base + index).or_insert(value);
            }
        }
        edit_offset = u32_at(data, 8)?;
    }
    if document_id == 0 {
        return Err(PptError::Corrupted("missing Document persist ID".into()));
    }
    Ok((mapping, document_id))
}

fn replace_nested_record(record: &[u8], target: u16, replacement: &[u8]) -> Result<Vec<u8>> {
    if type_of(record)? == target {
        return Ok(replacement.to_vec());
    }
    let version = u16::from_le_bytes([record[0], record[1]]) & 0xF;
    if version != 0xF {
        return Err(PptError::Corrupted(
            "ExObjList not found in Document container".into(),
        ));
    }
    let mut data = Vec::new();
    let mut offset = 8usize;
    let mut found = false;
    while offset < record.len() {
        let child = record_slice(record, offset)?;
        if (type_of(child)? == target || (u16::from_le_bytes([child[0], child[1]]) & 0xF) == 0xF)
            && let Ok(changed) = replace_nested_record(child, target, replacement)
        {
            found |= changed != child;
            data.extend_from_slice(&changed);
            offset += child.len();
            continue;
        }
        data.extend_from_slice(child);
        offset += child.len();
    }
    if !found {
        return Err(PptError::Corrupted(
            "ExObjList not found in Document container".into(),
        ));
    }
    let mut output = record[..4].to_vec();
    output.extend_from_slice(&(data.len() as u32).to_le_bytes());
    output.extend_from_slice(&data);
    Ok(output)
}

fn record_slice(data: &[u8], offset: usize) -> Result<&[u8]> {
    let len = u32_at(data, offset + 4)? as usize;
    let end = offset
        .checked_add(8)
        .and_then(|v| v.checked_add(len))
        .ok_or_else(|| PptError::Corrupted("record length overflow".into()))?;
    data.get(offset..end)
        .ok_or_else(|| PptError::Corrupted("truncated record".into()))
}
fn type_of(record: &[u8]) -> Result<u16> {
    record
        .get(2..4)
        .map(|v| u16::from_le_bytes([v[0], v[1]]))
        .ok_or_else(|| PptError::Corrupted("truncated header".into()))
}
fn u32_at(data: &[u8], offset: usize) -> Result<u32> {
    data.get(offset..offset + 4)
        .map(|v| u32::from_le_bytes(v.try_into().unwrap()))
        .ok_or_else(|| PptError::Corrupted("truncated u32".into()))
}
fn refs(path: &[String]) -> Vec<&str> {
    path.iter().map(String::as_str).collect()
}

#[cfg(test)]
mod tests {
    use super::*;

    fn ppt_record(version: u16, kind: u16, data: &[u8]) -> Vec<u8> {
        let mut output = version.to_le_bytes().to_vec();
        output.extend_from_slice(&kind.to_le_bytes());
        output.extend_from_slice(&(data.len() as u32).to_le_bytes());
        output.extend_from_slice(data);
        output
    }

    #[test]
    fn recursively_replaces_only_external_object_list() {
        let unknown = ppt_record(0, 0x7777, b"unknown");
        let old = ppt_record(0x000F, EX_OBJ_LIST, b"old");
        let mut children = unknown.clone();
        children.extend_from_slice(&old);
        let document = ppt_record(0x000F, 1000, &children);
        let replacement = ppt_record(0x000F, EX_OBJ_LIST, b"new-list");
        let rewritten = replace_nested_record(&document, EX_OBJ_LIST, &replacement).unwrap();
        assert!(
            rewritten
                .windows(unknown.len())
                .any(|value| value == unknown)
        );
        assert!(
            rewritten
                .windows(replacement.len())
                .any(|value| value == replacement)
        );
        assert!(!rewritten.windows(old.len()).any(|value| value == old));
    }

    #[test]
    fn merges_newest_incremental_mapping_over_prior_edit() {
        let object1 = ppt_record(0, 0x1111, b"one");
        let object2 = ppt_record(0, 0x2222, b"two");
        let mut document = object1.clone();
        document.extend_from_slice(&object2);
        let mut first_dir = PersistPtrBuilder::new();
        first_dir.set_offset(1, 0);
        let first_dir_offset = document.len() as u32;
        document.extend_from_slice(&first_dir.generate_full_record());
        let first_edit_offset = document.len() as u32;
        document.extend_from_slice(
            &UserEditAtom::new_minimal(first_dir_offset, 1, 1, 0).generate_record(),
        );
        let replacement_offset = document.len() as u32;
        document.extend_from_slice(&object2);
        let mut second_dir = PersistPtrBuilder::new();
        second_dir.set_offset(1, replacement_offset);
        let second_dir_offset = document.len() as u32;
        document.extend_from_slice(&second_dir.generate_incremental_record());
        let mut edit = UserEditAtom::new_minimal(second_dir_offset, 1, 1, 0);
        edit.offset_last_edit = first_edit_offset;
        let second_edit_offset = document.len() as u32;
        document.extend_from_slice(&edit.generate_record());
        let (mapping, document_id) = mapping_chain(&document, second_edit_offset).unwrap();
        assert_eq!(document_id, 1);
        assert_eq!(mapping.get(&1), Some(&replacement_offset));
    }
}
