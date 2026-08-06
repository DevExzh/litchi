//! Package discovery and initial persisted-record state.

use super::super::{Collection, Editor, Result, mapping, rewrite};
use crate::package::Error;
use litchi_cfb::OleFile;
use std::io::Cursor;

pub(crate) fn open(bytes: Vec<u8>, collection: Collection) -> Result<Editor> {
    let mut ole = OleFile::open(Cursor::new(bytes.clone()))?;
    let paths = ole.list_streams();
    reject_unsupported_package(&paths)?;

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
        streams,
        document_path,
        current_user_path,
        document,
        current_user,
        mappings,
        current_edit_offset,
        document_persist_id,
        collection,
        staged_storage: Default::default(),
        removed_persist_ids: Default::default(),
        rewrite_object_list: false,
        changed: false,
    })
}

pub(crate) fn open_records(bytes: Vec<u8>) -> Result<Editor> {
    open(
        bytes,
        Collection {
            id_seed: 1,
            objects: Vec::new(),
            unknown_records: Vec::new(),
        },
    )
}

fn reject_unsupported_package(paths: &[Vec<String>]) -> Result<()> {
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
        return Err(Error::Corrupted(
            "signed or encrypted PPT is not eligible for OLE editing".into(),
        ));
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

fn validate_object_mappings(
    collection: &Collection,
    mappings: &std::collections::BTreeMap<u32, u32>,
) -> Result<()> {
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
