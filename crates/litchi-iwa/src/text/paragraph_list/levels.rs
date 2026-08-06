//! Per-paragraph list-level boundaries in TSWP storage objects.

use prost::Message;

use crate::archive::RawMessage;
use crate::protobuf::tswp;
use crate::wire::{
    patch_varint_field, repeated_length_delimited_payloads,
    rewrite_repeated_length_delimited_fields, transform_length_delimited_field,
};
use crate::{Error, IWorkPackage, Result};

use super::super::storage_wire::{
    LocatedStorage, StorageLocation, locate_storage as locate_native_storage,
    locate_storage_with_archive as locate_native_storage_with_archive, update_parsed_archive,
};
use litchi_iwa_text::paragraph::list::{ParagraphListLevel, ParagraphListLevelPlacement};
use litchi_iwa_text::position::TextPosition;

const PARAGRAPH_DATA_TABLE_FIELD: u32 = 6;
const TABLE_ENTRIES_FIELD: u32 = 1;
const ENTRY_LIST_LEVEL_FIELD: u32 = 2;

fn level_from_native(value: u32) -> Result<ParagraphListLevel> {
    let level = u8::try_from(value).map_err(|_| {
        Error::InvalidFormat(format!("native paragraph list level {value} exceeds u8"))
    })?;
    Ok(ParagraphListLevel::new(level)?)
}

pub(crate) fn paragraph_list_levels(
    package: &IWorkPackage,
    storage_id: u64,
) -> Result<Vec<ParagraphListLevelPlacement>> {
    let location = locate_storage(package, storage_id)?;
    let storage = &location.storage;
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
                TextPosition::from_utf16_index(entry.character_index as usize)?,
                level_from_native(entry.first)?,
            ))
        })
        .collect()
}

pub(crate) fn paragraph_list_level(
    package: &IWorkPackage,
    storage_id: u64,
    paragraph: TextPosition,
) -> Result<ParagraphListLevel> {
    let location = locate_storage(package, storage_id)?;
    let storage = &location.storage;
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
    level_from_native(value.first)
}

pub(in crate::text) fn set_paragraph_list_level(
    package: &mut IWorkPackage,
    storage_id: u64,
    paragraph: TextPosition,
    level: ParagraphListLevel,
) -> Result<()> {
    if paragraph_list_level(package, storage_id, paragraph)? == level {
        return Ok(());
    }
    let located = locate_storage_with_archive(package, storage_id)?;
    let mut staged = package.clone();
    patch_level(&mut staged, located, storage_id, paragraph, level)?;
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
    paragraph: TextPosition,
) -> Result<bool> {
    if paragraph_list_level(package, storage_id, paragraph)? == ParagraphListLevel::ZERO {
        return Ok(false);
    }
    set_paragraph_list_level(package, storage_id, paragraph, ParagraphListLevel::ZERO)?;
    Ok(true)
}

fn patch_level(
    package: &mut IWorkPackage,
    located: LocatedStorage,
    storage_id: u64,
    paragraph: TextPosition,
    level: ParagraphListLevel,
) -> Result<()> {
    let LocatedStorage { location, archive } = located;
    update_parsed_archive(package, &location.archive_name, archive, |archive| {
        let object = archive.object_mut(storage_id).ok_or_else(|| {
            Error::InvalidFormat(format!("iWork text storage {storage_id} is missing"))
        })?;
        if object.archive_info.identifier != Some(location.object_id) {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} object identity changed unexpectedly"
            )));
        }
        if object
            .archive_info
            .message_infos
            .get(location.message_index)
            .is_none()
        {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} writable payload metadata index {} is missing",
                location.message_index
            )));
        }
        let original = object.messages.get(location.message_index).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork text storage {storage_id} writable payload index {} is missing",
                location.message_index
            ))
        })?;
        if original.type_ != location.message_type {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} writable payload changed unexpectedly"
            )));
        }
        let starts = paragraph_starts(&location.storage.text)?;
        require_paragraph_start(storage_id, paragraph, &starts)?;
        let decoded = location
            .storage
            .table_para_data
            .as_ref()
            .map(|table| table.entries.as_slice())
            .unwrap_or_default();
        validate_entries(storage_id, decoded, &starts)?;
        let data = transform_length_delimited_field(
            &original.data,
            PARAGRAPH_DATA_TABLE_FIELD,
            |table| patch_table(table, storage_id, paragraph, level, &starts),
        )?;
        object.replace_message(
            location.message_index,
            RawMessage {
                type_: location.message_type,
                data,
            },
        )?;
        Ok(())
    })
}

fn patch_table(
    table: &[u8],
    storage_id: u64,
    paragraph: TextPosition,
    level: ParagraphListLevel,
    paragraph_starts: &[u32],
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
    let next_paragraph = paragraph_starts
        .iter()
        .copied()
        .find(|start| *start > target);
    let next_effective = next_paragraph
        .map(|next| {
            entries
                .iter()
                .take_while(|(entry, _)| entry.character_index <= next)
                .last()
                .map(|(entry, _)| (entry.first, entry.second))
                .ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "iWork text storage {storage_id} has no paragraph-data boundary at zero"
                    ))
                })
        })
        .transpose()?;
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
    if let Some((next, (next_level, next_second))) = next_paragraph.zip(next_effective)
        && entries
            .iter()
            .all(|(entry, _)| entry.character_index != next)
    {
        let entry = tswp::para_data_attribute_table::ParaDataAttribute {
            character_index: next,
            first: next_level,
            second: next_second,
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

fn locate_storage(package: &IWorkPackage, storage_id: u64) -> Result<StorageLocation> {
    let location = locate_native_storage(
        package,
        storage_id,
        PARAGRAPH_DATA_TABLE_FIELD,
        "paragraph-data",
    )?;
    if !location.table_present {
        return Err(Error::InvalidFormat(format!(
            "iWork text storage {storage_id} must contain one paragraph-data table, found 0"
        )));
    }
    Ok(location)
}

fn locate_storage_with_archive(package: &IWorkPackage, storage_id: u64) -> Result<LocatedStorage> {
    let located = locate_native_storage_with_archive(
        package,
        storage_id,
        PARAGRAPH_DATA_TABLE_FIELD,
        "paragraph-data",
    )?;
    if located.location.storage.table_para_data.is_some() != located.location.table_present {
        return Err(Error::InvalidFormat(format!(
            "iWork text storage {storage_id} paragraph-data table wire state is inconsistent"
        )));
    }
    if !located.location.table_present {
        return Err(Error::InvalidFormat(format!(
            "iWork text storage {storage_id} must contain one paragraph-data table, found 0"
        )));
    }
    Ok(located)
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
        level_from_native(entry.first)?;
        previous = Some(entry.character_index);
    }
    Ok(())
}

pub(super) fn require_paragraph_start(
    storage_id: u64,
    paragraph: TextPosition,
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
    use crate::archive::{Archive, ArchiveObject, RawMessage};

    #[test]
    fn level_lookup_rejects_malformed_recognized_storage() {
        let package = test_package(vec![RawMessage {
            type_: 2022,
            data: vec![0x80],
        }]);

        assert!(paragraph_list_levels(&package, 42).is_err());
    }

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
        let paragraph = TextPosition::from_utf16_index(4).unwrap();
        let patched = patch_table(&table, 99, paragraph, ParagraphListLevel::ONE, &[0, 4]).unwrap();
        let decoded = tswp::ParaDataAttributeTable::decode(patched.as_slice()).unwrap();
        assert_eq!(decoded.entries.len(), 2);
        assert!(decoded.entries.iter().all(|entry| entry.second == 7));

        let reset =
            patch_table(&patched, 99, paragraph, ParagraphListLevel::ZERO, &[0, 4]).unwrap();
        let decoded = tswp::ParaDataAttributeTable::decode(reset.as_slice()).unwrap();
        assert_eq!(decoded.entries.len(), 1);
        assert_eq!(decoded.entries[0].second, 7);
    }

    #[test]
    fn paragraph_starts_span_fragments_and_coalesce_crlf() {
        let text = vec!["A\r".to_owned(), "\nB\u{2029}".to_owned(), "C".to_owned()];
        assert_eq!(paragraph_starts(&text).unwrap(), vec![0, 3, 5]);
    }

    #[test]
    fn level_updates_restore_the_following_paragraph() {
        let table = tswp::ParaDataAttributeTable {
            entries: vec![tswp::para_data_attribute_table::ParaDataAttribute {
                character_index: 0,
                first: 0,
                second: 0,
            }],
        }
        .encode_to_vec();
        let paragraph = TextPosition::from_utf16_index(4).unwrap();
        let patched =
            patch_table(&table, 99, paragraph, ParagraphListLevel::ONE, &[0, 4, 9]).unwrap();
        let decoded = tswp::ParaDataAttributeTable::decode(patched.as_slice()).unwrap();
        assert_eq!(
            decoded
                .entries
                .iter()
                .map(|entry| (entry.character_index, entry.first))
                .collect::<Vec<_>>(),
            [(0, 0), (4, 1), (9, 0)]
        );
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
