//! Package discovery and initial persisted-record state.

use super::super::{Collection, Editor, Result, mapping, rewrite};
use crate::package::Error;
use litchi_cfb::OleFile;
use litchi_ole_common::protection::is_protected_component;
use std::collections::{BTreeMap, HashSet};
use std::io::{Cursor, Read, Seek};
use std::sync::Arc;

#[allow(
    dead_code,
    reason = "used by capability-specific package snapshot paths"
)]
pub(crate) fn inspect_live_document(bytes: &[u8]) -> Result<(u32, Vec<u8>)> {
    let mut ole = OleFile::open(Cursor::new(bytes))?;
    inspect_live_document_from_ole(&mut ole)
}

pub(crate) fn inspect_live_document_from_ole<R: Read + Seek>(
    ole: &mut OleFile<R>,
) -> Result<(u32, Vec<u8>)> {
    let paths = ole.list_streams();
    let document_path = find_stream(&paths, "PowerPoint Document")?;
    let current_user_path = find_stream(&paths, "Current User")?;
    let document = ole.open_stream(&stream_refs(&document_path))?;
    let current_user = ole.open_stream(&stream_refs(&current_user_path))?;
    validate_current_user(&current_user)?;
    let current_edit_offset = rewrite::u32_at(&current_user, 16)?;
    let (mappings, document_persist_id) = mapping::read(&document, current_edit_offset)?;
    let offset = *mappings
        .get(&document_persist_id)
        .ok_or_else(|| Error::Corrupted("live Document persist mapping is missing".into()))?
        as usize;
    Ok((
        document_persist_id,
        rewrite::slice(&document, offset)?.to_vec(),
    ))
}

pub(crate) fn inspect_live_mapping(
    document: &[u8],
    current_user: &[u8],
) -> Result<crate::persist::PersistMapping> {
    validate_current_user(current_user)?;
    let current_edit_offset = rewrite::u32_at(current_user, 16)?;
    let (mappings, _) = mapping::read(document, current_edit_offset)?;
    let mut result = crate::persist::PersistMapping::new();
    for (persist_id, offset) in mappings {
        result.add_mapping(persist_id, offset);
    }
    Ok(result)
}

pub(crate) fn open(bytes: Arc<[u8]>, collection: Collection) -> Result<Editor> {
    open_with_limit(bytes, collection, usize::MAX)
}

fn open_with_limit(
    bytes: Arc<[u8]>,
    collection: Collection,
    max_output_bytes: usize,
) -> Result<Editor> {
    if bytes.len() > max_output_bytes {
        return Err(Error::ResourceLimit(format!(
            "PowerPoint editor source is {} bytes, exceeding the {max_output_bytes}-byte output limit",
            bytes.len()
        )));
    }
    let mut ole = OleFile::open(Cursor::new(bytes.clone()))?;
    reject_unsupported_package(&ole)?;
    let paths = ole.list_streams();

    let document_path = find_stream(&paths, "PowerPoint Document")?;
    let current_user_path = find_stream(&paths, "Current User")?;
    let document = ole.open_stream(&stream_refs(&document_path))?;
    let current_user = ole.open_stream(&stream_refs(&current_user_path))?;
    validate_current_user(&current_user)?;

    let current_edit_offset = rewrite::u32_at(&current_user, 16)?;
    let (mappings, document_persist_id) = mapping::read(&document, current_edit_offset)?;
    collection.validate()?;
    validate_object_mappings(&collection, &mappings)?;

    let streams = paths
        .into_iter()
        .map(|path| {
            let data = ole.open_stream(&stream_refs(&path))?;
            Ok((path, data))
        })
        .collect::<Result<Vec<_>>>()?;

    Ok(Editor {
        original: bytes,
        max_output_bytes,
        streams,
        document_path,
        current_user_path,
        document,
        current_user,
        mappings,
        current_edit_offset,
        document_persist_id,
        collection,
        staged_storage: BTreeMap::default(),
        removed_persist_ids: HashSet::default(),
        rewrite_object_list: false,
        changed: false,
    })
}

pub(crate) fn open_records(bytes: Arc<[u8]>) -> Result<Editor> {
    open_records_arc_with_limit(bytes, usize::MAX)
}

pub(crate) fn open_records_arc_with_limit(
    bytes: Arc<[u8]>,
    max_output_bytes: usize,
) -> Result<Editor> {
    open_with_limit(
        bytes,
        Collection {
            id_seed: 1,
            objects: Vec::new(),
            unknown_records: Vec::new(),
        },
        max_output_bytes,
    )
}

fn reject_unsupported_package<R: Read + Seek>(ole: &OleFile<R>) -> Result<()> {
    let mut pending = vec![Vec::<String>::new()];
    while let Some(directory) = pending.pop() {
        let directory_refs: Vec<_> = directory.iter().map(String::as_str).collect();
        for entry in ole.list_directory_entries(&directory_refs)? {
            if is_protected_component(&entry.name) {
                return Err(Error::Corrupted(
                    "signed or encrypted PPT is not eligible for OLE editing".into(),
                ));
            }
            let mut path = directory.clone();
            path.push(entry.name.clone());
            let path_refs: Vec<_> = path.iter().map(String::as_str).collect();
            if ole.directory_exists(&path_refs) {
                pending.push(path);
            }
        }
    }
    Ok(())
}

fn find_stream(paths: &[Vec<String>], name: &str) -> Result<Vec<String>> {
    paths
        .iter()
        .find(|path| path.last().is_some_and(|value| value == name))
        .cloned()
        .ok_or_else(|| Error::StreamNotFound(name.into()))
}

fn validate_current_user(current_user: &[u8]) -> Result<()> {
    if current_user.len() < 28 || rewrite::u32_at(current_user, 12)? != 0xE391_C05F {
        return Err(Error::Corrupted(
            "unsupported or encrypted CurrentUserAtom".into(),
        ));
    }
    Ok(())
}

fn validate_object_mappings(collection: &Collection, mappings: &BTreeMap<u32, u32>) -> Result<()> {
    for object in &collection.objects {
        if !mappings.contains_key(&object.persist_id()) {
            return Err(Error::Corrupted(format!(
                "OLE object references missing persist ID {}",
                object.persist_id()
            )));
        }
    }
    Ok(())
}

fn stream_refs(path: &[String]) -> Vec<&str> {
    path.iter().map(String::as_str).collect()
}
