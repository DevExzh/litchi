//! Releasing `TST.TableDataList` references held by removed cells.
//!
//! A removed row or column takes its cells' style, conditional-style,
//! control and format keys with it. Each key's entry loses one reference
//! per removed cell; an entry whose count reaches zero is dropped, with the
//! object references it held (a style entry's style stays in the
//! stylesheet, only the list's reference to it goes). Untouched entries
//! keep their bytes.

use std::collections::BTreeMap;

use super::wire::{bytes_field, rebuild, single_bytes, single_varint, varint_field};
use super::{Error, Result, reference_identifier, reserve};

/// Entry fields that hold a `TSP.Reference`: `reference`,
/// `rich_text_payload`, `comment_storage`.
const REFERENCE_FIELDS: [u32; 3] = [4, 9, 10];

/// Release the keys in `remaining` that `list` (a list or a list segment,
/// both with entries in field 3) holds, removing them from `remaining`.
///
/// Returns the rewritten payload (`None` when the list holds none of the
/// keys) and the object references the dropped entries held.
pub fn release(
    list: &[u8],
    remaining: &mut BTreeMap<u32, u32>,
) -> Result<(Option<Vec<u8>>, Vec<u64>)> {
    release_with(list, remaining, false)
}

/// [`release`], handing back the list's last key when its entry goes
/// (`reclaim`; record 0329), as the cell writer does for a string list.
pub fn release_with(
    list: &[u8],
    remaining: &mut BTreeMap<u32, u32>,
    reclaim: bool,
) -> Result<(Option<Vec<u8>>, Vec<u64>)> {
    let mut pruned = Vec::new();
    let mut changed = false;
    let mut dropped = Vec::new();
    let output = rebuild(list, |field| {
        if field.number != 3 {
            return Ok(None);
        }
        if field.wire_type != 2 {
            return Err(Error::Invalid);
        }
        let key = u32::try_from(single_varint(field.payload, 1)?.ok_or(Error::Invalid)?)
            .map_err(|_error| Error::Invalid)?;
        let Some(release) = remaining.remove(&key) else {
            return Ok(None);
        };
        changed = true;
        let count = single_varint(field.payload, 2)?.ok_or(Error::Invalid)?;
        let after = count
            .checked_sub(u64::from(release))
            .ok_or(Error::Unsupported)?;
        if after > 0 {
            let entry = rebuild(field.payload, |inner| {
                if inner.number == 2 {
                    return varint_field(2, after).map(Some);
                }
                Ok(None)
            })?;
            return bytes_field(3, &entry).map(Some);
        }
        for number in REFERENCE_FIELDS {
            if let Some(reference) = single_bytes(field.payload, number)? {
                reserve(&mut pruned, 1)?;
                pruned.push(reference_identifier(reference)?);
            }
        }
        reserve(&mut dropped, 1)?;
        dropped.push(u64::from(key));
        Ok(Some(Vec::new()))
    })?;
    let mut output = output;
    if reclaim && !dropped.is_empty() {
        let current = single_varint(&output, 2)?.ok_or(Error::Invalid)?;
        let mut next = current;
        dropped.sort_unstable();
        for key in dropped {
            if key.checked_add(1) == Some(next) {
                next = key;
            }
        }
        if next != current {
            output = rebuild(&output, |field| {
                if field.number == 2 {
                    return varint_field(2, next).map(Some);
                }
                Ok(None)
            })?;
        }
    }
    Ok((changed.then_some(output), pruned))
}

#[cfg(test)]
mod tests {
    use std::collections::BTreeMap;

    use super::release;
    use crate::table_topology::wire::{bytes_field, push_varint_field};

    fn entry(key: u64, count: u64, reference: Option<u64>) -> Vec<u8> {
        let mut payload = Vec::new();
        push_varint_field(&mut payload, 1, key).unwrap();
        push_varint_field(&mut payload, 2, count).unwrap();
        if let Some(reference) = reference {
            let mut inner = Vec::new();
            push_varint_field(&mut inner, 1, reference).unwrap();
            payload.extend(bytes_field(4, &inner).unwrap());
        }
        bytes_field(3, &payload).unwrap()
    }

    #[test]
    fn counts_drop_and_empty_entries_go_with_their_references() {
        let mut list = Vec::new();
        push_varint_field(&mut list, 1, 4).unwrap();
        push_varint_field(&mut list, 2, 9).unwrap();
        list.extend(entry(1, 3, Some(70)));
        list.extend(entry(2, 1, Some(71)));
        list.extend(entry(5, 2, None));
        let mut remaining = BTreeMap::from([(1, 1), (2, 1), (8, 1)]);
        let (payload, pruned) = release(&list, &mut remaining).unwrap();
        let mut expected = Vec::new();
        push_varint_field(&mut expected, 1, 4).unwrap();
        push_varint_field(&mut expected, 2, 9).unwrap();
        expected.extend(entry(1, 2, Some(70)));
        expected.extend(entry(5, 2, None));
        assert_eq!(payload.unwrap(), expected);
        assert_eq!(pruned, [71]);
        assert_eq!(remaining, BTreeMap::from([(8, 1)]));
        let mut over = BTreeMap::from([(5, 3)]);
        assert!(release(&list, &mut over).is_err());
    }
}
