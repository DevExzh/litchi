//! Small typed codecs for host-neutral OGraph records.
//!
//! Module context keeps public names short: for example,
//! [`chart3d::BarShape`] instead of `Chart3dBarShape`.

pub mod chart3d;
pub mod frame;
pub mod line;
pub mod marker;
pub mod pie;
pub mod series;

use crate::raw::{Kind, RecordRef};
use crate::{Error, Result};

pub(crate) fn payload<'a>(record: RecordRef<'a>, expected: Kind, len: usize) -> Result<&'a [u8]> {
    if record.kind() != expected {
        return Err(Error::WrongRecord {
            expected: expected.get(),
            actual: record.kind().get(),
        });
    }
    payload_bytes(expected, record.payload(), len)
}

pub(crate) fn payload_bytes(kind: Kind, bytes: &[u8], len: usize) -> Result<&[u8]> {
    if bytes.len() != len {
        return Err(Error::InvalidRecordLength {
            kind: kind.get(),
            expected: len,
            actual: bytes.len(),
        });
    }
    Ok(bytes)
}

pub(crate) fn u16_at(bytes: &[u8], offset: usize) -> Option<u16> {
    let pair = bytes.get(offset..offset.checked_add(2)?)?;
    Some(u16::from_le_bytes([pair[0], pair[1]]))
}
