//! Shared bounded access to native TSWP storage payloads and UTF-16 boundaries.

use prost::Message;

use crate::protobuf::tswp::StorageArchive;
use crate::wire::repeated_length_delimited_payloads;
use crate::{Error, IWorkPackage, Result};

pub(super) const STORAGE_MESSAGE_TYPES: &[u32] = &[2_001, 2_022];

pub(super) struct StorageLocation {
    pub archive_name: String,
    pub storage: StorageArchive,
    pub table_present: bool,
}

/// Locate one writable storage and validate one optional singular table field.
pub(super) fn locate_storage(
    package: &IWorkPackage,
    storage_id: u64,
    table_field: u32,
    table_label: &str,
) -> Result<StorageLocation> {
    let mut found = None;
    for archive_name in package.iwa_entry_names() {
        let archive = package.archive(archive_name)?;
        let Some(object) = archive.object(storage_id) else {
            continue;
        };
        let payloads = object
            .messages
            .iter()
            .filter(|message| STORAGE_MESSAGE_TYPES.contains(&message.type_))
            .collect::<Vec<_>>();
        let [message] = payloads.as_slice() else {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} must have exactly one writable payload"
            )));
        };
        let storage = StorageArchive::decode(message.data.as_slice())?;
        let table_count =
            repeated_length_delimited_payloads(message.data.as_slice(), table_field)?.len();
        if table_count > 1 {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} contains {table_count} {table_label} tables"
            )));
        }
        let location = StorageLocation {
            archive_name: archive_name.to_owned(),
            storage,
            table_present: table_count == 1,
        };
        if found.replace(location).is_some() {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} occurs in multiple archives"
            )));
        }
    }
    found.ok_or_else(|| Error::InvalidFormat(format!("iWork text storage {storage_id} is missing")))
}

pub(super) fn require_text_boundary(storage_id: u64, position: u32, text: &[String]) -> Result<()> {
    validate_sorted_boundaries(storage_id, [position], text)
}

pub(super) fn validate_sorted_boundaries(
    storage_id: u64,
    boundaries: impl IntoIterator<Item = u32>,
    text: &[String],
) -> Result<()> {
    let mut boundaries = boundaries.into_iter().peekable();
    let mut index = 0u32;
    while boundaries.peek().is_some_and(|boundary| *boundary == index) {
        boundaries.next();
    }
    for fragment in text {
        for character in fragment.chars() {
            index = index
                .checked_add(character.len_utf16() as u32)
                .ok_or_else(|| {
                    Error::InvalidFormat("iWork text UTF-16 length overflow".to_owned())
                })?;
            while boundaries.peek().is_some_and(|boundary| *boundary == index) {
                boundaries.next();
            }
            if boundaries.peek().is_some_and(|boundary| *boundary < index) {
                break;
            }
        }
    }
    if let Some(boundary) = boundaries.next() {
        return Err(Error::InvalidFormat(format!(
            "UTF-16 index {boundary} is not a scalar boundary in iWork text storage {storage_id}"
        )));
    }
    Ok(())
}

pub(super) fn text_utf16_len(text: &[String]) -> Result<u32> {
    text.iter().try_fold(0u32, |total, fragment| {
        fragment.chars().try_fold(total, |total, character| {
            total
                .checked_add(character.len_utf16() as u32)
                .ok_or_else(|| Error::InvalidFormat("iWork text UTF-16 length overflow".to_owned()))
        })
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn scalar_boundaries_span_fragments_without_joining_text() {
        let text = vec!["A😀".to_owned(), "é".to_owned()];
        for valid in [0, 1, 3, 4] {
            require_text_boundary(1, valid, &text).unwrap();
        }
        for invalid in [2, 5] {
            assert!(require_text_boundary(1, invalid, &text).is_err());
        }
        assert_eq!(text_utf16_len(&text).unwrap(), 4);
    }
}
