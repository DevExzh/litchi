//! Shared bounded access to native TSWP storage payloads and UTF-16 boundaries.

use std::collections::HashSet;

use prost::Message;

use crate::archive::RawMessage;
use crate::protobuf::tswp::StorageArchive;
use crate::wire::repeated_length_delimited_payloads;
use crate::{Error, IWorkPackage, Result};

pub(super) const STORAGE_MESSAGE_TYPES: &[u32] = &[2_001, 2_022];

pub(super) struct StorageLocation {
    pub object_id: u64,
    pub archive_name: String,
    pub message_index: usize,
    pub message_type: u32,
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
    locate_storage_in_package(package, storage_id, Some((table_field, table_label)))
}

/// Locate one writable text storage without applying a table-specific policy.
pub(super) fn locate_text_storage(
    package: &IWorkPackage,
    storage_id: u64,
) -> Result<StorageLocation> {
    locate_storage_in_package(package, storage_id, None)
}

/// Discover every writable text storage while validating each recognized payload.
pub(super) fn locate_text_storages(package: &IWorkPackage) -> Result<Vec<StorageLocation>> {
    let mut locations = Vec::new();
    let mut seen = HashSet::new();
    for archive_name in package.iwa_entry_names() {
        let archive = package.archive(archive_name)?;
        for object in archive.objects {
            let Some(object_id) = object.archive_info.identifier else {
                continue;
            };
            if !seen.insert(object_id) {
                return Err(Error::InvalidFormat(format!(
                    "Text storage object {object_id} occurs in multiple IWA components"
                )));
            }
            let Some(location) =
                resolve_object_storage(archive_name, object_id, &object.messages, None)?
            else {
                continue;
            };
            locations.push(location);
        }
    }
    Ok(locations)
}

fn locate_storage_in_package(
    package: &IWorkPackage,
    storage_id: u64,
    table: Option<(u32, &str)>,
) -> Result<StorageLocation> {
    let mut found = None;
    let mut object_found = false;
    for archive_name in package.iwa_entry_names() {
        let archive = package.archive(archive_name)?;
        let Some(object) = archive.object(storage_id) else {
            continue;
        };
        if object_found {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} occurs in multiple archives"
            )));
        }
        object_found = true;
        let Some(location) =
            resolve_object_storage(archive_name, storage_id, &object.messages, table)?
        else {
            continue;
        };
        if found.is_some() {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} occurs in multiple archives"
            )));
        }
        found = Some(location);
    }
    found.ok_or_else(|| Error::InvalidFormat(format!("iWork text storage {storage_id} is missing")))
}

fn resolve_object_storage(
    archive_name: &str,
    object_id: u64,
    messages: &[RawMessage],
    table: Option<(u32, &str)>,
) -> Result<Option<StorageLocation>> {
    let mut found = None;
    for (message_index, message) in messages.iter().enumerate() {
        if !STORAGE_MESSAGE_TYPES.contains(&message.type_) {
            continue;
        }
        let Some(storage) = decode_storage_payload(
            object_id,
            archive_name,
            message_index,
            message.type_,
            &message.data,
        )?
        else {
            continue;
        };
        if found.is_some() {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {object_id} must have exactly one writable payload"
            )));
        }
        let table_present = if let Some((table_field, table_label)) = table {
            let table_count =
                repeated_length_delimited_payloads(message.data.as_slice(), table_field)?.len();
            if table_count > 1 {
                return Err(Error::InvalidFormat(format!(
                    "iWork text storage {object_id} contains {table_count} {table_label} tables"
                )));
            }
            table_count == 1
        } else {
            false
        };
        found = Some(StorageLocation {
            object_id,
            archive_name: archive_name.to_owned(),
            message_index,
            message_type: message.type_,
            storage,
            table_present,
        });
    }
    Ok(found)
}

fn decode_storage_payload(
    object_id: u64,
    archive_name: &str,
    message_index: usize,
    message_type: u32,
    data: &[u8],
) -> Result<Option<StorageArchive>> {
    match StorageArchive::decode(data) {
        Ok(storage) => Ok(Some(storage)),
        Err(_error)
            if message_type == 2_022
                && crate::protobuf::tswp::ParagraphStyleArchive::decode(data).is_ok() =>
        {
            // Numbers reuses the native 2022 message type for paragraph styles.
            // A valid paragraph-style payload is not a text storage and must not
            // be rejected by a package-wide storage discovery pass.
            Ok(None)
        },
        Err(error) => Err(Error::InvalidFormat(format!(
            "iWork text storage {object_id} has a malformed writable payload in {archive_name} message {message_index}: {error}"
        ))),
    }
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
    use crate::archive::{Archive, ArchiveObject};
    use prost::Message;

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

    #[test]
    fn rejects_duplicate_object_when_one_copy_has_no_storage_payload() {
        let recognized = ArchiveObject::new(
            42,
            vec![RawMessage {
                type_: 2_001,
                data: crate::protobuf::tswp::StorageArchive::default().encode_to_vec(),
            }],
        )
        .unwrap();
        let unrecognized = ArchiveObject::new(42, Vec::new()).unwrap();
        let mut package = IWorkPackage::new();
        package
            .replace_archive(
                "Index/One.iwa",
                &Archive {
                    objects: vec![recognized],
                },
            )
            .unwrap();
        package
            .replace_archive(
                "Index/Two.iwa",
                &Archive {
                    objects: vec![unrecognized],
                },
            )
            .unwrap();

        assert!(locate_text_storage(&package, 42).is_err());
        assert!(locate_text_storages(&package).is_err());
    }

    #[test]
    fn ignores_numbers_paragraph_style_using_storage_message_type() {
        let style = crate::protobuf::tswp::ParagraphStyleArchive {
            super_: crate::protobuf::tss::StyleArchive::default(),
            ..Default::default()
        };
        let object = ArchiveObject::new(
            43,
            vec![RawMessage {
                type_: 2_022,
                data: style.encode_to_vec(),
            }],
        )
        .unwrap();
        let mut package = IWorkPackage::new();
        package
            .replace_archive(
                "Index/DocumentStylesheet.iwa",
                &Archive {
                    objects: vec![object],
                },
            )
            .unwrap();

        assert!(locate_text_storages(&package).unwrap().is_empty());
    }
}
