//! Root-record publication for external-media snapshots.
//!
//! This layer only rewrites the owning record envelope. It never opens a path
//! or interprets an external payload.

use std::ops::Range;

use crate::consts::RecordType;
use crate::package::{Error, Result};
use crate::records::Record;

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct Ancestor {
    pub(crate) offset: usize,
    pub(crate) payload_len: usize,
}

/// Location of the unique `ExObjListContainer` and its owning envelopes.
#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct Location {
    pub(crate) range: Range<usize>,
    pub(crate) ancestors: Vec<Ancestor>,
}

/// Locate the media list without reconstructing any unrelated record bytes.
pub(crate) fn locate(root: &Record) -> Result<Option<Location>> {
    let mut locations = Vec::new();
    visit(root, 0, &[], &mut locations)?;
    match locations.len() {
        0 => Ok(None),
        1 => Ok(locations.pop()),
        _ => Err(Error::Corrupted(
            "record tree contains multiple external-object lists".into(),
        )),
    }
}

fn visit(
    record: &Record,
    offset: usize,
    ancestors: &[Ancestor],
    locations: &mut Vec<Location>,
) -> Result<()> {
    let record_len = 8usize
        .checked_add(record.data.len())
        .ok_or_else(|| Error::Corrupted("PPT record size overflows usize".into()))?;
    let end = offset
        .checked_add(record_len)
        .ok_or_else(|| Error::Corrupted("PPT record offset overflows usize".into()))?;
    if record.record_type == RecordType::ExObjList {
        locations.push(Location {
            range: offset..end,
            ancestors: ancestors.to_vec(),
        });
    }

    let mut child_offset = offset
        .checked_add(8)
        .ok_or_else(|| Error::Corrupted("PPT child offset overflows usize".into()))?;
    let mut child_ancestors = ancestors.to_vec();
    child_ancestors.push(Ancestor {
        offset,
        payload_len: record.data.len(),
    });
    for child in &record.children {
        let child_len = 8usize
            .checked_add(child.data.len())
            .ok_or_else(|| Error::Corrupted("PPT child size overflows usize".into()))?;
        let child_end = child_offset
            .checked_add(child_len)
            .ok_or_else(|| Error::Corrupted("PPT child offset overflows usize".into()))?;
        if child_end > end {
            return Err(Error::Corrupted(
                "PPT record child extends beyond its owner".into(),
            ));
        }
        visit(child, child_offset, &child_ancestors, locations)?;
        child_offset = child_end;
    }
    Ok(())
}

/// Replace the exact list range and adjust every owning record length.
pub(crate) fn replace(source: &[u8], location: &Location, replacement: &[u8]) -> Result<Vec<u8>> {
    if location.range.start > location.range.end
        || location.range.end > source.len()
        || source.get(location.range.clone()).is_none()
    {
        return Err(Error::Corrupted(
            "external-media list range is outside its source record".into(),
        ));
    }
    let old_len = location.range.len();
    let delta = replacement.len() as i128 - old_len as i128;
    let output_len = source.len() as i128 + delta;
    let output_len = usize::try_from(output_len)
        .map_err(|_| Error::Corrupted("external-media publication size overflows usize".into()))?;
    let mut output = Vec::with_capacity(output_len);
    output.extend_from_slice(&source[..location.range.start]);
    output.extend_from_slice(replacement);
    output.extend_from_slice(&source[location.range.end..]);

    for ancestor in &location.ancestors {
        let new_payload_len = ancestor.payload_len as i128 + delta;
        let new_payload_len = u32::try_from(new_payload_len).map_err(|_| {
            Error::Corrupted("owning PPT record payload exceeds u32 after publication".into())
        })?;
        let length_end = ancestor
            .offset
            .checked_add(8)
            .ok_or_else(|| Error::Corrupted("owning PPT record header overflows usize".into()))?;
        if length_end > output.len() {
            return Err(Error::Corrupted(
                "owning PPT record header is outside the published source".into(),
            ));
        }
        output[ancestor.offset + 4..length_end].copy_from_slice(&new_payload_len.to_le_bytes());
    }
    Ok(output)
}

/// Append a new list to a document that has no external-object list yet.
pub(crate) fn append_to_document(
    source: &[u8],
    root_type: RecordType,
    replacement: &[u8],
) -> Result<Vec<u8>> {
    if root_type != RecordType::Document || source.len() < 8 {
        return Err(Error::InvalidFormat(
            "a new external-media list can only be attached to a DocumentContainer".into(),
        ));
    }
    let old_payload_len = source.len() - 8;
    let declared_payload_len = source
        .get(4..8)
        .and_then(|bytes| bytes.try_into().ok())
        .map(u32::from_le_bytes)
        .ok_or_else(|| Error::Corrupted("DocumentContainer header is truncated".into()))?;
    if usize::try_from(declared_payload_len).ok() != Some(old_payload_len) {
        return Err(Error::Corrupted(
            "DocumentContainer source length does not match its record header".into(),
        ));
    }
    let new_payload_len = old_payload_len
        .checked_add(replacement.len())
        .ok_or_else(|| Error::Corrupted("DocumentContainer payload size overflows usize".into()))?;
    let new_payload_len = u32::try_from(new_payload_len)
        .map_err(|_| Error::Corrupted("DocumentContainer payload exceeds u32".into()))?;
    let mut output = Vec::with_capacity(source.len().saturating_add(replacement.len()));
    output.extend_from_slice(&source[..4]);
    output.extend_from_slice(&new_payload_len.to_le_bytes());
    output.extend_from_slice(&source[8..]);
    output.extend_from_slice(replacement);
    Ok(output)
}
