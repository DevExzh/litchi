//! Per-paragraph numbered-list continuation and restart overrides.

use prost::Message;

use crate::archive::RawMessage;
use crate::protobuf::tswp;
use crate::wire::{
    parse_wire_fields, patch_length_delimited_field, patch_varint_field,
    repeated_length_delimited_payloads, rewrite_repeated_length_delimited_fields,
    transform_length_delimited_field,
};
use crate::{Error, IWorkPackage, Result};

use super::super::storage_wire::{StorageLocation, locate_storage as locate_native_storage};
use super::levels::{paragraph_starts, require_paragraph_start};
use super::paragraph_lists;
use super::types::{ParagraphList, ParagraphListNumbering};
use crate::text::ParagraphStart;

const PARAGRAPH_START_TABLE_FIELD: u32 = 14;
const TABLE_ENTRIES_FIELD: u32 = 1;
const ENTRY_START_NUMBER_FIELD: u32 = 2;

pub(crate) fn paragraph_list_numbering(
    package: &IWorkPackage,
    storage_id: u64,
    paragraph: ParagraphStart,
) -> Result<ParagraphListNumbering> {
    let location = locate_storage(package, storage_id)?;
    let storage = &location.storage;
    let starts = paragraph_starts(&storage.text)?;
    require_paragraph_start(storage_id, paragraph, &starts)?;
    let Some(table) = storage.table_para_starts.as_ref() else {
        return Ok(ParagraphListNumbering::Continue);
    };
    validate_entries(storage_id, &table.entries, &starts)?;
    table
        .entries
        .iter()
        .find(|entry| entry.character_index == paragraph.utf16_index())
        .map_or(Ok(ParagraphListNumbering::Continue), |entry| {
            ParagraphListNumbering::from_native(entry.first)
        })
}

pub(in crate::text) fn set_paragraph_list_numbering(
    package: &mut IWorkPackage,
    storage_id: u64,
    paragraph: ParagraphStart,
    numbering: ParagraphListNumbering,
) -> Result<()> {
    if paragraph_list_numbering(package, storage_id, paragraph)? == numbering {
        return Ok(());
    }
    if matches!(numbering, ParagraphListNumbering::StartAt(_)) {
        require_numbered_paragraph(package, storage_id, paragraph)?;
    }
    let location = locate_storage(package, storage_id)?;
    let mut staged = package.clone();
    patch_numbering(&mut staged, location, storage_id, paragraph, numbering)?;
    if paragraph_list_numbering(&staged, storage_id, paragraph)? != numbering {
        return Err(Error::InvalidFormat(
            "iWork paragraph list-numbering update failed validation".to_owned(),
        ));
    }
    *package = staged;
    Ok(())
}

fn require_numbered_paragraph(
    package: &IWorkPackage,
    storage_id: u64,
    paragraph: ParagraphStart,
) -> Result<()> {
    let placements = paragraph_lists(package, storage_id)?;
    let list = placements
        .iter()
        .take_while(|placement| placement.paragraph <= paragraph)
        .last()
        .map(|placement| placement.list)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork text storage {storage_id} has no list-style boundary at zero"
            ))
        })?;
    if list == ParagraphList::Numbered {
        return Ok(());
    }
    Err(Error::InvalidFormat(format!(
        "UTF-16 index {} is not a numbered-list paragraph in iWork text storage {storage_id}",
        paragraph.utf16_index()
    )))
}

fn patch_numbering(
    package: &mut IWorkPackage,
    location: StorageLocation,
    storage_id: u64,
    paragraph: ParagraphStart,
    numbering: ParagraphListNumbering,
) -> Result<()> {
    let StorageLocation {
        archive_name,
        message_index,
        message_type,
        storage,
        ..
    } = location;
    package.update_archive(&archive_name, |archive| {
        let object = archive.object_mut(storage_id).ok_or_else(|| {
            Error::InvalidFormat(format!("iWork text storage {storage_id} is missing"))
        })?;
        let original = object.messages.get(message_index).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork text storage {storage_id} writable payload index {message_index} is missing"
            ))
        })?;
        if original.type_ != message_type {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} writable payload changed unexpectedly"
            )));
        }
        let starts = paragraph_starts(&storage.text)?;
        require_paragraph_start(storage_id, paragraph, &starts)?;
        let table_count =
            repeated_length_delimited_payloads(&original.data, PARAGRAPH_START_TABLE_FIELD)?.len();
        if table_count > 1 {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} contains duplicate paragraph-start tables"
            )));
        }
        if let Some(table) = storage.table_para_starts.as_ref() {
            validate_entries(storage_id, &table.entries, &starts)?;
        }
        let data = if table_count == 1 {
            transform_length_delimited_field(
                &original.data,
                PARAGRAPH_START_TABLE_FIELD,
                |table| patch_table(table, storage_id, paragraph, numbering),
            )?
        } else {
            let table = new_table(paragraph, numbering).encode_to_vec();
            patch_length_delimited_field(
                &original.data,
                PARAGRAPH_START_TABLE_FIELD,
                false,
                Some(&table),
            )?
        };
        object.replace_message(
            message_index,
            RawMessage {
                type_: message_type,
                data,
            },
        )?;
        Ok(())
    })
}

fn patch_table(
    table: &[u8],
    storage_id: u64,
    paragraph: ParagraphStart,
    numbering: ParagraphListNumbering,
) -> Result<Vec<u8>> {
    let payloads = repeated_length_delimited_payloads(table, TABLE_ENTRIES_FIELD)?;
    let mut entries = payloads
        .iter()
        .map(|payload| {
            tswp::para_data_attribute_table::ParaDataAttribute::decode(*payload)
                .map(|entry| (entry, (*payload).to_vec()))
        })
        .collect::<std::result::Result<Vec<_>, _>>()?;
    let target = paragraph.utf16_index();
    let native_start = numbering.native_start();
    if numbering == ParagraphListNumbering::Continue && target != 0 {
        let mut retained = Vec::with_capacity(entries.len());
        for (mut entry, raw) in entries {
            if entry.character_index != target {
                retained.push((entry, raw));
            } else if entry_has_parallel_data(&entry, &raw)? {
                entry.first = 0;
                let raw = patch_varint_field(&raw, ENTRY_START_NUMBER_FIELD, true, Some(0))?;
                retained.push((entry, raw));
            }
        }
        entries = retained;
    } else if let Some((entry, raw)) = entries
        .iter_mut()
        .find(|(entry, _)| entry.character_index == target)
    {
        entry.first = native_start;
        *raw = patch_varint_field(
            raw,
            ENTRY_START_NUMBER_FIELD,
            true,
            Some(u64::from(native_start)),
        )?;
    } else {
        let entry = tswp::para_data_attribute_table::ParaDataAttribute {
            character_index: target,
            first: native_start,
            second: 0,
        };
        let raw = entry.encode_to_vec();
        entries.push((entry, raw));
        entries.sort_by_key(|(entry, _)| entry.character_index);
    }
    if entries
        .first()
        .is_none_or(|(entry, _)| entry.character_index != 0)
    {
        let entry = tswp::para_data_attribute_table::ParaDataAttribute {
            character_index: 0,
            first: 0,
            second: 0,
        };
        let raw = entry.encode_to_vec();
        entries.insert(0, (entry, raw));
    }
    let encoded = entries.into_iter().map(|(_, raw)| raw).collect::<Vec<_>>();
    if encoded.is_empty() {
        return Err(Error::InvalidFormat(format!(
            "iWork text storage {storage_id} paragraph-start table became empty"
        )));
    }
    rewrite_repeated_length_delimited_fields(table, TABLE_ENTRIES_FIELD, &encoded)
}

fn entry_has_parallel_data(
    entry: &tswp::para_data_attribute_table::ParaDataAttribute,
    raw: &[u8],
) -> Result<bool> {
    const CHARACTER_INDEX_FIELD: u32 = 1;
    const PARALLEL_VALUE_FIELD: u32 = 3;

    Ok(entry.second != 0
        || parse_wire_fields(raw)?.iter().any(|field| {
            !matches!(
                field.number,
                CHARACTER_INDEX_FIELD | ENTRY_START_NUMBER_FIELD | PARALLEL_VALUE_FIELD
            )
        }))
}

fn new_table(
    paragraph: ParagraphStart,
    numbering: ParagraphListNumbering,
) -> tswp::ParaDataAttributeTable {
    let target = paragraph.utf16_index();
    let mut entries = Vec::with_capacity(usize::from(target != 0) + 1);
    if target != 0 {
        entries.push(tswp::para_data_attribute_table::ParaDataAttribute {
            character_index: 0,
            first: 0,
            second: 0,
        });
    }
    entries.push(tswp::para_data_attribute_table::ParaDataAttribute {
        character_index: target,
        first: numbering.native_start(),
        second: 0,
    });
    tswp::ParaDataAttributeTable { entries }
}

fn locate_storage(package: &IWorkPackage, storage_id: u64) -> Result<StorageLocation> {
    locate_native_storage(
        package,
        storage_id,
        PARAGRAPH_START_TABLE_FIELD,
        "paragraph-start",
    )
}

fn validate_entries(
    storage_id: u64,
    entries: &[tswp::para_data_attribute_table::ParaDataAttribute],
    paragraph_starts: &[u32],
) -> Result<()> {
    let Some(first) = entries.first() else {
        return Err(Error::InvalidFormat(format!(
            "iWork text storage {storage_id} has an empty paragraph-start table"
        )));
    };
    if first.character_index != 0 {
        return Err(Error::InvalidFormat(format!(
            "iWork text storage {storage_id} paragraph-start data must begin at UTF-16 index zero"
        )));
    }
    let mut previous = None;
    for entry in entries {
        if previous.is_some_and(|index| index >= entry.character_index) {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} paragraph-start boundaries are not strictly increasing"
            )));
        }
        if paragraph_starts
            .binary_search(&entry.character_index)
            .is_err()
        {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} paragraph-start boundary {} is not a paragraph start",
                entry.character_index
            )));
        }
        ParagraphListNumbering::from_native(entry.first)?;
        previous = Some(entry.character_index);
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::archive::{Archive, ArchiveObject};
    use crate::text::ParagraphListStart;

    #[test]
    fn numbering_lookup_rejects_malformed_recognized_storage() {
        let package = test_package(vec![RawMessage {
            type_: 2001,
            data: vec![0x80],
        }]);

        assert!(paragraph_list_numbering(&package, 42, ParagraphStart::ZERO).is_err());
    }

    #[test]
    fn numbering_table_restart_and_continue_preserve_parallel_data() {
        let table = tswp::ParaDataAttributeTable {
            entries: vec![
                tswp::para_data_attribute_table::ParaDataAttribute {
                    character_index: 0,
                    first: 0,
                    second: 9,
                },
                tswp::para_data_attribute_table::ParaDataAttribute {
                    character_index: 4,
                    first: 0,
                    second: 8,
                },
            ],
        }
        .encode_to_vec();
        let paragraph = ParagraphStart::from_utf16_index(4).unwrap();
        let restart = ParagraphListNumbering::StartAt(ParagraphListStart::new(7).unwrap());
        let patched = patch_table(&table, 99, paragraph, restart).unwrap();
        let decoded = tswp::ParaDataAttributeTable::decode(patched.as_slice()).unwrap();
        assert_eq!(decoded.entries[0].second, 9);
        assert_eq!(decoded.entries[1].first, 7);
        assert_eq!(decoded.entries[1].second, 8);

        let continued =
            patch_table(&patched, 99, paragraph, ParagraphListNumbering::Continue).unwrap();
        let decoded = tswp::ParaDataAttributeTable::decode(continued.as_slice()).unwrap();
        assert_eq!(decoded.entries.len(), 2);
        assert_eq!(decoded.entries[0].second, 9);
        assert_eq!(decoded.entries[1].first, 0);
        assert_eq!(decoded.entries[1].second, 8);
    }

    fn test_package(messages: Vec<RawMessage>) -> IWorkPackage {
        let object = ArchiveObject::new(42, messages).unwrap();
        let mut package = IWorkPackage::new();
        package
            .replace_archive(
                "Index/Document.iwa",
                &Archive {
                    objects: vec![object],
                },
            )
            .unwrap();
        package
    }
}
