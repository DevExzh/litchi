//! Resolves the newest visible persisted-record offset from `UserEdit` atoms.

use super::directory;
use crate::embedded::object::editor::{Result, rewrite};
use std::collections::{BTreeMap, HashSet};

const USER_EDIT: u16 = 4085;

pub(crate) fn read(document: &[u8], mut edit_offset: u32) -> Result<(BTreeMap<u32, u32>, u32)> {
    let mut mapping = BTreeMap::new();
    let mut seen = HashSet::new();
    let mut document_id = 0;

    while edit_offset != 0 {
        if !seen.insert(edit_offset) || seen.len() > 4_096 {
            return Err(crate::package::Error::Corrupted(
                "cyclic or excessive UserEdit chain".into(),
            ));
        }
        let record = rewrite::slice(document, edit_offset as usize)?;
        if rewrite::type_of(record)? != USER_EDIT || record.len() < 36 {
            return Err(crate::package::Error::Corrupted(
                "invalid UserEditAtom".into(),
            ));
        }

        let data = &record[8..];
        if document_id == 0 {
            document_id = rewrite::u32_at(data, 16)?;
        }
        let directory = rewrite::slice(document, rewrite::u32_at(data, 12)? as usize)?;
        directory::merge(&mut mapping, directory)?;
        edit_offset = rewrite::u32_at(data, 8)?;
    }

    if document_id == 0 {
        return Err(crate::package::Error::Corrupted(
            "missing Document persist ID".into(),
        ));
    }
    Ok((mapping, document_id))
}
