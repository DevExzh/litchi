//! Recursive replacement of one record inside a container hierarchy.

use super::record;
use crate::embedded::object::editor::Result;
use crate::package::Error;

const EX_OBJ_LIST: u16 = 1033;
const DOCUMENT: u16 = 1000;
const DOCUMENT_ATOM: u16 = 1001;
pub(crate) const MAX_NESTED_RECORD_DEPTH: usize = 256;

#[cfg(test)]
pub(crate) fn replace_nested_record(
    record: &[u8],
    target: u16,
    replacement: &[u8],
) -> Result<Vec<u8>> {
    replace_nested_record_at_depth(record, target, replacement, 0)?
        .ok_or_else(|| Error::Corrupted("ExObjList not found in Document container".into()))
}

fn replace_nested_record_at_depth(
    record: &[u8],
    target: u16,
    replacement: &[u8],
    depth: usize,
) -> Result<Option<Vec<u8>>> {
    if depth >= MAX_NESTED_RECORD_DEPTH {
        return Err(Error::Corrupted(
            "PPT record nesting exceeds the safety limit".into(),
        ));
    }
    if record::type_of(record)? == target {
        return Ok(Some(replacement.to_vec()));
    }
    let version = u16::from_le_bytes([record[0], record[1]]) & 0xF;
    if version != 0xF {
        return Ok(None);
    }

    let mut data = Vec::new();
    let mut offset = 8usize;
    let mut found = false;
    while offset < record.len() {
        let child = record::slice(record, offset)?;
        let child_type = record::type_of(child)?;
        let child_is_container = u16::from_le_bytes([child[0], child[1]]) & 0xF == 0xF;
        if (child_type == target || child_is_container)
            && let Some(changed) =
                replace_nested_record_at_depth(child, target, replacement, depth + 1)?
        {
            found = true;
            data.extend_from_slice(&changed);
            offset += child.len();
            continue;
        }
        data.extend_from_slice(child);
        offset += child.len();
    }
    if !found {
        return Ok(None);
    }

    let mut output = record[..4].to_vec();
    let data_len = u32::try_from(data.len())
        .map_err(|_err| Error::Corrupted("PPT record payload exceeds u32".into()))?;
    output.extend_from_slice(&data_len.to_le_bytes());
    output.extend_from_slice(&data);
    Ok(Some(output))
}

/// Replace the `ExObjListContainer` of the `DocumentContainer` `document`
/// with `list`, or insert `list` when the document has none.
///
/// `[MS-PPT]` 2.4.1 orders the `DocumentContainer` children `documentAtom`,
/// then the optional `exObjList`, then `documentTextInfo` and the rest, so a
/// new list is inserted directly after the `DocumentAtom`, where `PowerPoint`
/// writes it.
pub(crate) fn replace_or_insert_object_list(document: &[u8], list: &[u8]) -> Result<Vec<u8>> {
    if let Some(replaced) = replace_nested_record_at_depth(document, EX_OBJ_LIST, list, 0)? {
        return Ok(replaced);
    }
    let version = u16::from_le_bytes([document[0], document[1]]) & 0xF;
    if version != 0xF || record::type_of(document)? != DOCUMENT {
        return Err(Error::Corrupted(
            "the live persisted document is not a DocumentContainer".into(),
        ));
    }
    let atom = record::slice(document, 8)?;
    if record::type_of(atom)? != DOCUMENT_ATOM {
        return Err(Error::Corrupted(
            "DocumentContainer does not start with a DocumentAtom".into(),
        ));
    }
    let split = 8 + atom.len();
    let payload_len = document
        .len()
        .checked_sub(8)
        .and_then(|len| len.checked_add(list.len()))
        .and_then(|len| u32::try_from(len).ok())
        .ok_or_else(|| Error::Corrupted("PPT record payload exceeds u32".into()))?;
    let mut output = Vec::with_capacity(document.len() + list.len());
    output.extend_from_slice(&document[..4]);
    output.extend_from_slice(&payload_len.to_le_bytes());
    output.extend_from_slice(&document[8..split]);
    output.extend_from_slice(list);
    output.extend_from_slice(&document[split..]);
    Ok(output)
}

#[cfg(test)]
pub(crate) const fn external_object_list() -> u16 {
    EX_OBJ_LIST
}
