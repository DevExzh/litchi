//! Worksheet package record headers for list-object families.

use super::super::codec::{append_frt, record};
use super::super::model::ListObject;
use super::super::{FEAT_HDR11_RECORD_TYPE, ISF_LIST, invalid};
use crate::Result;

pub(crate) fn feature_header_record(tables: &[ListObject]) -> Result<Vec<u8>> {
    let next = tables
        .iter()
        .map(|table| table.id.value())
        .max()
        .unwrap_or(0)
        .checked_add(1)
        .ok_or_else(|| invalid(FEAT_HDR11_RECORD_TYPE, "next table id overflows"))?;
    let mut payload = Vec::new();
    append_frt(&mut payload, FEAT_HDR11_RECORD_TYPE, None);
    payload.extend_from_slice(&ISF_LIST.to_le_bytes());
    payload.push(1);
    payload.extend_from_slice(&u32::MAX.to_le_bytes());
    payload.extend_from_slice(&u32::MAX.to_le_bytes());
    payload.extend_from_slice(&next.to_le_bytes());
    payload.extend_from_slice(&0u16.to_le_bytes());
    record(FEAT_HDR11_RECORD_TYPE, payload)
}
