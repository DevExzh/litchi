//! Per-paragraph list-level boundaries in TSWP storage objects.

use prost::Message;

use crate::archive::RawMessage;
use crate::protobuf::tswp::{self, StorageArchive};
use crate::wire::{
    patch_varint_field, repeated_length_delimited_payloads,
    rewrite_repeated_length_delimited_fields, transform_length_delimited_field,
};
use crate::{Error, IWorkPackage, Result};

use super::super::drop_cap::ParagraphStart;
use super::types::{ParagraphListLevel, ParagraphListLevelPlacement};

const STORAGE_MESSAGE_TYPES: &[u32] = &[2_001, 2_022];
const PARAGRAPH_DATA_TABLE_FIELD: u32 = 6;
const TABLE_ENTRIES_FIELD: u32 = 1;
const ENTRY_LIST_LEVEL_FIELD: u32 = 2;

pub(in crate::text) fn paragraph_list_levels(
    package: &IWorkPackage,
    storage_id: u64,
) -> Result<Vec<ParagraphListLevelPlacement>> {
    let (_, storage) = locate_storage(package, storage_id)?;
    let starts = paragraph_starts(&storage.text)?;
    let entries = storage
        .table_para_data
        .as_ref()
        .map(|table| table.entries.as_slice())
        .unwrap_or_default();
    validate_entries(storage_id, entries, &starts)?;
    entries
        .iter()
        .map(|entry| {
            Ok(ParagraphListLevelPlacement::new(
                ParagraphStart::from_utf16_index(entry.character_index as usize)?,
                ParagraphListLevel::from_native(entry.first)?,
            ))
        })
        .collect()
}

pub(in crate::text) fn paragraph_list_level(
    package: &IWorkPackage,
    storage_id: u64,
    paragraph: ParagraphStart,
) -> Result<ParagraphListLevel> {
    let (_, storage) = locate_storage(package, storage_id)?;
    let starts = paragraph_starts(&storage.text)?;
    require_paragraph_start(storage_id, paragraph, &starts)?;
    let entries = storage
        .table_para_data
        .as_ref()
        .map(|table| table.entries.as_slice())
        .unwrap_or_default();
    validate_entries(storage_id, entries, &starts)?;
    let value = entries
        .iter()
        .take_while(|entry| entry.character_index <= paragraph.utf16_index())
        .last()
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork text storage {storage_id} has no paragraph-data boundary at zero"
            ))
        })?;
    ParagraphListLevel::from_native(value.first)
}

pub(in crate::text) fn set_paragraph_list_level(
    package: &mut IWorkPackage,
    storage_id: u64,
    paragraph: ParagraphStart,
    level: ParagraphListLevel,
) -> Result<()> {
    if paragraph_list_level(package, storage_id, paragraph)? == level {
        return Ok(());
    }
    let (archive_name, _) = locate_storage(package, storage_id)?;
    let mut staged = package.clone();
    patch_level(&mut staged, &archive_name, storage_id, paragraph, level)?;
    if paragraph_list_level(&staged, storage_id, paragraph)? != level {
        return Err(Error::InvalidFormat(
            "iWork paragraph list-level update failed validation".to_owned(),
        ));
    }
    *package = staged;
    Ok(())
}

pub(in crate::text) fn reset_paragraph_list_level(
    package: &mut IWorkPackage,
    storage_id: u64,
    paragraph: ParagraphStart,
) -> Result<bool> {
    if paragraph_list_level(package, storage_id, paragraph)? == ParagraphListLevel::ZERO {
        return Ok(false);
    }
    set_paragraph_list_level(package, storage_id, paragraph, ParagraphListLevel::ZERO)?;
    Ok(true)
}

fn patch_level(
    package: &mut IWorkPackage,
    archive_name: &str,
    storage_id: u64,
    paragraph: ParagraphStart,
    level: ParagraphListLevel,
) -> Result<()> {
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(storage_id).ok_or_else(|| {
            Error::InvalidFormat(format!("iWork text storage {storage_id} is missing"))
        })?;
        let indexes = object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| {
                STORAGE_MESSAGE_TYPES.contains(&message.type_)
                    && StorageArchive::decode(message.data.as_slice()).is_ok()
            })
            .map(|(index, _)| index)
            .collect::<Vec<_>>();
        let [index] = indexes.as_slice() else {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} must have exactly one writable payload"
            )));
        };
        let original = &object.messages[*index];
        let storage = StorageArchive::decode(original.data.as_slice())?;
        let starts = paragraph_starts(&storage.text)?;
        require_paragraph_start(storage_id, paragraph, &starts)?;
        let decoded = storage
            .table_para_data
            .as_ref()
            .map(|table| table.entries.as_slice())
            .unwrap_or_default();
        validate_entries(storage_id, decoded, &starts)?;
        let data = transform_length_delimited_field(
            &original.data,
            PARAGRAPH_DATA_TABLE_FIELD,
            |table| patch_table(table, storage_id, paragraph, level),
        )?;
        object.replace_message(
            *index,
            RawMessage {
                type_: original.type_,
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
    level: ParagraphListLevel,
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
    let inherited_second = entries
        .iter()
        .take_while(|(entry, _)| entry.character_index <= target)
        .last()
        .map(|(entry, _)| entry.second)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork text storage {storage_id} has no paragraph-data boundary at zero"
            ))
        })?;
    if let Some((entry, raw)) = entries
        .iter_mut()
        .find(|(entry, _)| entry.character_index == target)
    {
        entry.first = u32::from(level.get());
        *raw = patch_varint_field(
            raw,
            ENTRY_LIST_LEVEL_FIELD,
            true,
            Some(u64::from(level.get())),
        )?;
    } else {
        let entry = tswp::para_data_attribute_table::ParaDataAttribute {
            character_index: target,
            first: u32::from(level.get()),
            second: inherited_second,
        };
        let raw = entry.encode_to_vec();
        entries.push((entry, raw));
        entries.sort_by_key(|(entry, _)| entry.character_index);
    }
    entries.dedup_by(|(current, _), (previous, _)| {
        current.first == previous.first && current.second == previous.second
    });
    let encoded = entries.into_iter().map(|(_, raw)| raw).collect::<Vec<_>>();
    rewrite_repeated_length_delimited_fields(table, TABLE_ENTRIES_FIELD, &encoded)
}

fn locate_storage(package: &IWorkPackage, storage_id: u64) -> Result<(String, StorageArchive)> {
    let mut found = None;
    for archive_name in package.iwa_entry_names() {
        let archive = package.archive(archive_name)?;
        let Some(object) = archive.object(storage_id) else {
            continue;
        };
        let mut payloads = object
            .messages
            .iter()
            .enumerate()
            .filter_map(|(index, message)| {
                STORAGE_MESSAGE_TYPES
                    .contains(&message.type_)
                    .then(|| {
                        StorageArchive::decode(message.data.as_slice())
                            .ok()
                            .map(|storage| (index, storage))
                    })
                    .flatten()
            })
            .collect::<Vec<_>>();
        if payloads.len() != 1 {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} must have exactly one writable payload"
            )));
        }
        let (message_index, storage) = payloads.pop().ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork text storage {storage_id} payload disappeared"
            ))
        })?;
        let table_count = object
            .messages
            .get(message_index)
            .map(|message| {
                repeated_length_delimited_payloads(
                    message.data.as_slice(),
                    PARAGRAPH_DATA_TABLE_FIELD,
                )
            })
            .transpose()?
            .map_or(0, |tables| tables.len());
        if table_count != 1 {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} must contain one paragraph-data table, found {table_count}"
            )));
        }
        if found.replace((archive_name.to_owned(), storage)).is_some() {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} occurs in multiple archives"
            )));
        }
    }
    found.ok_or_else(|| Error::InvalidFormat(format!("iWork text storage {storage_id} is missing")))
}

fn validate_entries(
    storage_id: u64,
    entries: &[tswp::para_data_attribute_table::ParaDataAttribute],
    starts: &[u32],
) -> Result<()> {
    let Some(first) = entries.first() else {
        return Err(Error::InvalidFormat(format!(
            "iWork text storage {storage_id} has no paragraph-data boundary"
        )));
    };
    if first.character_index != 0 {
        return Err(Error::InvalidFormat(format!(
            "iWork text storage {storage_id} paragraph data must begin at UTF-16 index zero"
        )));
    }
    let mut previous = None;
    for entry in entries {
        if previous.is_some_and(|index| index >= entry.character_index) {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} paragraph-data boundaries are not strictly increasing"
            )));
        }
        if starts.binary_search(&entry.character_index).is_err() {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} paragraph-data boundary {} is not a paragraph start",
                entry.character_index
            )));
        }
        ParagraphListLevel::from_native(entry.first)?;
        previous = Some(entry.character_index);
    }
    Ok(())
}

pub(super) fn require_paragraph_start(
    storage_id: u64,
    paragraph: ParagraphStart,
    starts: &[u32],
) -> Result<()> {
    if starts.binary_search(&paragraph.utf16_index()).is_ok() {
        return Ok(());
    }
    Err(Error::InvalidFormat(format!(
        "UTF-16 index {} is not a paragraph start in iWork text storage {storage_id}",
        paragraph.utf16_index()
    )))
}

pub(super) fn paragraph_starts(text: &[String]) -> Result<Vec<u32>> {
    let mut starts = vec![0];
    let mut index = 0u32;
    let mut previous_was_carriage_return = false;
    for fragment in text {
        for character in fragment.chars() {
            index = index
                .checked_add(character.len_utf16() as u32)
                .ok_or_else(|| {
                    Error::InvalidFormat("iWork text exceeds the UTF-16 u32 range".to_owned())
                })?;
            match character {
                '\n' => {
                    if previous_was_carriage_return {
                        starts.pop();
                    }
                    starts.push(index);
                },
                '\r' | '\u{2028}' | '\u{2029}' => starts.push(index),
                _ => {},
            }
            previous_was_carriage_return = character == '\r';
        }
    }
    starts.dedup();
    Ok(starts)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn level_updates_preserve_unknown_parallel_paragraph_data() {
        let table = tswp::ParaDataAttributeTable {
            entries: vec![tswp::para_data_attribute_table::ParaDataAttribute {
                character_index: 0,
                first: 0,
                second: 7,
            }],
        }
        .encode_to_vec();
        let paragraph = ParagraphStart::from_utf16_index(4).unwrap();
        let patched = patch_table(&table, 99, paragraph, ParagraphListLevel::ONE).unwrap();
        let decoded = tswp::ParaDataAttributeTable::decode(patched.as_slice()).unwrap();
        assert_eq!(decoded.entries.len(), 2);
        assert!(decoded.entries.iter().all(|entry| entry.second == 7));

        let reset = patch_table(&patched, 99, paragraph, ParagraphListLevel::ZERO).unwrap();
        let decoded = tswp::ParaDataAttributeTable::decode(reset.as_slice()).unwrap();
        assert_eq!(decoded.entries.len(), 1);
        assert_eq!(decoded.entries[0].second, 7);
    }

    #[test]
    fn paragraph_starts_span_fragments_and_coalesce_crlf() {
        let text = vec!["A\r".to_owned(), "\nB\u{2029}".to_owned(), "C".to_owned()];
        assert_eq!(paragraph_starts(&text).unwrap(), vec![0, 3, 5]);
    }
}
