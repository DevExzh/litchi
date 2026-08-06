//! Decodes one full or incremental PersistDirectoryAtom.

use crate::embedded::object::editor::{Result, rewrite};
use std::collections::BTreeMap;

const PERSIST_FULL: u16 = 6001;
const PERSIST_INCREMENTAL: u16 = 6002;

pub(crate) fn merge(mapping: &mut BTreeMap<u32, u32>, directory: &[u8]) -> Result<()> {
    if !matches!(
        rewrite::record::type_of(directory)?,
        PERSIST_FULL | PERSIST_INCREMENTAL
    ) {
        return Err(crate::package::Error::Corrupted(
            "invalid PersistDirectoryAtom".into(),
        ));
    }

    let mut offset = 8usize;
    while offset < directory.len() {
        let info = rewrite::record::u32_at(directory, offset)?;
        offset += 4;
        let base = info & 0x000F_FFFF;
        let count = info >> 20;
        if count == 0 {
            return Err(crate::package::Error::Corrupted("zero persist run".into()));
        }
        for index in 0..count {
            let value = rewrite::record::u32_at(directory, offset)?;
            offset += 4;
            mapping.entry(base + index).or_insert(value);
        }
    }
    Ok(())
}
