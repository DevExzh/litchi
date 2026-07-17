//! Native language-run CRUD for TSWP text-storage objects.

use prost::Message;

use crate::archive::RawMessage;
use crate::protobuf::tswp::{self, StorageArchive};
use crate::wire::{
    patch_length_delimited_field, repeated_length_delimited_payloads,
    rewrite_repeated_length_delimited_fields,
};
use crate::{Error, IWorkPackage, Result};

use super::language_types::{TextLanguage, TextLanguageRun, TextPosition};

const STORAGE_MESSAGE_TYPES: &[u32] = &[2_001, 2_022];
const LANGUAGE_TABLE_FIELD: u32 = 19;
const TABLE_ENTRIES_FIELD: u32 = 1;
const ENTRY_LANGUAGE_FIELD: u32 = 2;

/// Read every explicit native language boundary.
pub(crate) fn text_languages(
    package: &IWorkPackage,
    storage_id: u64,
) -> Result<Vec<TextLanguageRun>> {
    let location = locate_storage(package, storage_id)?;
    let entries = language_entries(storage_id, &location.storage, location.table_present)?;
    validate_entries(storage_id, entries, &location.storage.text)?;
    entries
        .iter()
        .map(|entry| {
            Ok(TextLanguageRun::new(
                TextPosition::from_native(entry.character_index),
                TextLanguage::from_native(entry.object.as_deref())?,
            ))
        })
        .collect()
}

/// Read the effective language at one validated UTF-16 boundary.
pub(crate) fn text_language(
    package: &IWorkPackage,
    storage_id: u64,
    position: TextPosition,
) -> Result<TextLanguage> {
    let location = locate_storage(package, storage_id)?;
    require_text_boundary(storage_id, position, &location.storage.text)?;
    let entries = language_entries(storage_id, &location.storage, location.table_present)?;
    validate_entries(storage_id, entries, &location.storage.text)?;
    let Some(entry) = entries
        .iter()
        .take_while(|entry| entry.character_index <= position.utf16_index())
        .last()
    else {
        return Ok(TextLanguage::Automatic);
    };
    TextLanguage::from_native(entry.object.as_deref())
}

/// Set one language boundary while preserving unrelated native wire fields.
pub(crate) fn set_text_language(
    package: &mut IWorkPackage,
    storage_id: u64,
    position: TextPosition,
    language: &TextLanguage,
) -> Result<()> {
    if text_language(package, storage_id, position)? == *language {
        return Ok(());
    }
    let archive_name = locate_storage(package, storage_id)?.archive_name;
    patch_language(package, &archive_name, storage_id, position, language)?;
    if text_language(package, storage_id, position)? != *language {
        return Err(Error::InvalidFormat(
            "iWork text-language update failed validation".to_owned(),
        ));
    }
    Ok(())
}

/// Remove one nonzero language boundary so it inherits the preceding run.
pub(crate) fn remove_text_language_boundary(
    package: &mut IWorkPackage,
    storage_id: u64,
    position: TextPosition,
) -> Result<bool> {
    if position == TextPosition::ZERO {
        return Err(Error::ParseError(
            "the index-zero language boundary can only be removed by resetting all text languages"
                .to_owned(),
        ));
    }
    let location = locate_storage(package, storage_id)?;
    require_text_boundary(storage_id, position, &location.storage.text)?;
    let entries = language_entries(storage_id, &location.storage, location.table_present)?;
    validate_entries(storage_id, entries, &location.storage.text)?;
    if !entries
        .iter()
        .any(|entry| entry.character_index == position.utf16_index())
    {
        return Ok(false);
    }
    remove_boundary(package, &location.archive_name, storage_id, position)?;
    if text_languages(package, storage_id)?
        .iter()
        .any(|run| run.position == position)
    {
        return Err(Error::InvalidFormat(
            "iWork text-language boundary removal failed validation".to_owned(),
        ));
    }
    Ok(true)
}

/// Remove the entire explicit language table.
pub(crate) fn reset_text_languages(package: &mut IWorkPackage, storage_id: u64) -> Result<bool> {
    let location = locate_storage(package, storage_id)?;
    let entries = language_entries(storage_id, &location.storage, location.table_present)?;
    validate_entries(storage_id, entries, &location.storage.text)?;
    if !location.table_present {
        return Ok(false);
    }
    patch_language_table(package, &location.archive_name, storage_id, |_, _| Ok(None))?;
    if !text_languages(package, storage_id)?.is_empty() {
        return Err(Error::InvalidFormat(
            "iWork text-language reset failed validation".to_owned(),
        ));
    }
    Ok(true)
}

struct StorageLocation {
    archive_name: String,
    storage: StorageArchive,
    table_present: bool,
}

fn locate_storage(package: &IWorkPackage, storage_id: u64) -> Result<StorageLocation> {
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
            repeated_length_delimited_payloads(message.data.as_slice(), LANGUAGE_TABLE_FIELD)?
                .len();
        if table_count > 1 {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} contains {table_count} language tables"
            )));
        }
        let table_present = table_count == 1;
        if storage.table_language.is_some() != table_present {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} language-table wire state is inconsistent"
            )));
        }
        let location = StorageLocation {
            archive_name: archive_name.to_owned(),
            storage,
            table_present,
        };
        if found.replace(location).is_some() {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} occurs in multiple archives"
            )));
        }
    }
    found.ok_or_else(|| Error::InvalidFormat(format!("iWork text storage {storage_id} is missing")))
}

fn language_entries(
    storage_id: u64,
    storage: &StorageArchive,
    table_present: bool,
) -> Result<&[tswp::string_attribute_table::StringAttribute]> {
    match (table_present, storage.table_language.as_ref()) {
        (false, None) => Ok(&[]),
        (true, Some(table)) if table.entries.is_empty() => Err(Error::InvalidFormat(format!(
            "iWork text storage {storage_id} has an empty language table"
        ))),
        (true, Some(table)) => Ok(&table.entries),
        _ => Err(Error::InvalidFormat(format!(
            "iWork text storage {storage_id} language-table wire state is inconsistent"
        ))),
    }
}

fn patch_language(
    package: &mut IWorkPackage,
    archive_name: &str,
    storage_id: u64,
    position: TextPosition,
    language: &TextLanguage,
) -> Result<()> {
    patch_language_table(package, archive_name, storage_id, |table, storage| {
        require_text_boundary(storage_id, position, &storage.text)?;
        match table {
            Some(table) => patch_existing_table(table, position, language).map(Some),
            None => new_table(position, language).map(Some),
        }
    })
}

fn remove_boundary(
    package: &mut IWorkPackage,
    archive_name: &str,
    storage_id: u64,
    position: TextPosition,
) -> Result<()> {
    patch_language_table(package, archive_name, storage_id, |table, storage| {
        require_text_boundary(storage_id, position, &storage.text)?;
        let table = table.ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork text storage {storage_id} language table disappeared"
            ))
        })?;
        let payloads = repeated_length_delimited_payloads(table, TABLE_ENTRIES_FIELD)?;
        let mut retained = payloads
            .into_iter()
            .filter_map(|payload| {
                let entry = tswp::string_attribute_table::StringAttribute::decode(payload)
                    .map_err(Error::from);
                match entry {
                    Ok(entry) if entry.character_index == position.utf16_index() => None,
                    Ok(entry) => Some(Ok((entry, payload.to_vec()))),
                    Err(error) => Some(Err(error)),
                }
            })
            .collect::<Result<Vec<_>>>()?;
        retained.dedup_by(|(current, _), (previous, _)| current.object == previous.object);
        let retained = retained.into_iter().map(|(_, raw)| raw).collect::<Vec<_>>();
        rewrite_repeated_length_delimited_fields(table, TABLE_ENTRIES_FIELD, &retained).map(Some)
    })
}

fn patch_language_table<F>(
    package: &mut IWorkPackage,
    archive_name: &str,
    storage_id: u64,
    transform: F,
) -> Result<()>
where
    F: FnOnce(Option<&[u8]>, &StorageArchive) -> Result<Option<Vec<u8>>>,
{
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(storage_id).ok_or_else(|| {
            Error::InvalidFormat(format!("iWork text storage {storage_id} is missing"))
        })?;
        let indexes = object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| STORAGE_MESSAGE_TYPES.contains(&message.type_))
            .map(|(index, _)| index)
            .collect::<Vec<_>>();
        let [index] = indexes.as_slice() else {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} must have exactly one writable payload"
            )));
        };
        let original = &object.messages[*index];
        let storage = StorageArchive::decode(original.data.as_slice())?;
        let tables = repeated_length_delimited_payloads(&original.data, LANGUAGE_TABLE_FIELD)?;
        if tables.len() > 1 {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} contains {} language tables",
                tables.len()
            )));
        }
        let replacement = transform(tables.first().copied(), &storage)?;
        let data = patch_length_delimited_field(
            &original.data,
            LANGUAGE_TABLE_FIELD,
            !tables.is_empty(),
            replacement.as_deref(),
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

fn patch_existing_table(
    table: &[u8],
    position: TextPosition,
    language: &TextLanguage,
) -> Result<Vec<u8>> {
    let payloads = repeated_length_delimited_payloads(table, TABLE_ENTRIES_FIELD)?;
    let mut entries = payloads
        .iter()
        .map(|payload| {
            tswp::string_attribute_table::StringAttribute::decode(*payload)
                .map(|entry| (entry, (*payload).to_vec()))
                .map_err(Error::from)
        })
        .collect::<Result<Vec<_>>>()?;
    if let Some((entry, raw)) = entries
        .iter_mut()
        .find(|(entry, _)| entry.character_index == position.utf16_index())
    {
        *raw = patch_length_delimited_field(
            raw,
            ENTRY_LANGUAGE_FIELD,
            entry.object.is_some(),
            language.native_value().map(str::as_bytes),
        )?;
        entry.object = language.native_value().map(str::to_owned);
    } else {
        let entry = tswp::string_attribute_table::StringAttribute {
            character_index: position.utf16_index(),
            object: language.native_value().map(str::to_owned),
        };
        let raw = entry.encode_to_vec();
        entries.push((entry, raw));
        entries.sort_by_key(|(entry, _)| entry.character_index);
    }
    entries.dedup_by(|(current, _), (previous, _)| current.object == previous.object);
    let encoded = entries.into_iter().map(|(_, raw)| raw).collect::<Vec<_>>();
    rewrite_repeated_length_delimited_fields(table, TABLE_ENTRIES_FIELD, &encoded)
}

fn new_table(position: TextPosition, language: &TextLanguage) -> Result<Vec<u8>> {
    let mut entries = Vec::with_capacity(if position == TextPosition::ZERO { 1 } else { 2 });
    if position != TextPosition::ZERO {
        entries.push(tswp::string_attribute_table::StringAttribute {
            character_index: 0,
            object: None,
        });
    }
    entries.push(tswp::string_attribute_table::StringAttribute {
        character_index: position.utf16_index(),
        object: language.native_value().map(str::to_owned),
    });
    Ok(tswp::StringAttributeTable { entries }.encode_to_vec())
}

fn validate_entries(
    storage_id: u64,
    entries: &[tswp::string_attribute_table::StringAttribute],
    text: &[String],
) -> Result<()> {
    if entries.is_empty() {
        return Ok(());
    }
    if entries[0].character_index != 0 {
        return Err(Error::InvalidFormat(format!(
            "iWork text storage {storage_id} language table must begin at UTF-16 index zero"
        )));
    }
    let mut previous = None;
    for entry in entries {
        if previous.is_some_and(|index| index >= entry.character_index) {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} language boundaries are not strictly increasing"
            )));
        }
        TextLanguage::from_native(entry.object.as_deref())?;
        previous = Some(entry.character_index);
    }
    validate_sorted_boundaries(
        storage_id,
        entries.iter().map(|entry| entry.character_index),
        text,
    )
}

fn require_text_boundary(storage_id: u64, position: TextPosition, text: &[String]) -> Result<()> {
    validate_sorted_boundaries(storage_id, [position.utf16_index()], text)
}

fn validate_sorted_boundaries(
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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::pages::PagesEditor;
    use crate::shapes::{DrawablePoint, DrawableSize};
    use crate::wire::{append_length_delimited_field, parse_wire_fields, patch_varint_field};

    fn fixture() -> (IWorkPackage, u64, String) {
        let mut pages = PagesEditor::create_with_text("Body").unwrap();
        let text_box = pages
            .add_text_box(
                4,
                "Alpha\nBeta",
                DrawablePoint { x: 40.0, y: 80.0 },
                DrawableSize {
                    width: 320.0,
                    height: 140.0,
                },
            )
            .unwrap();
        let storage_id = text_box.storage.object_id;
        let archive_name = locate_storage(pages.package(), storage_id)
            .unwrap()
            .archive_name;
        (pages.into_package(), storage_id, archive_name)
    }

    fn storage_message_index(object: &crate::archive::ArchiveObject) -> usize {
        object
            .messages
            .iter()
            .position(|message| {
                STORAGE_MESSAGE_TYPES.contains(&message.type_)
                    && StorageArchive::decode(message.data.as_slice()).is_ok()
            })
            .unwrap()
    }

    #[test]
    fn scalar_boundaries_span_fragments_without_allocating_joined_text() {
        let text = vec!["A😀".to_owned(), "é".to_owned()];
        for valid in [0, 1, 3, 4] {
            require_text_boundary(1, TextPosition::from_utf16_index(valid).unwrap(), &text)
                .unwrap();
        }
        assert!(
            require_text_boundary(1, TextPosition::from_utf16_index(2).unwrap(), &text).is_err()
        );
        assert!(
            require_text_boundary(1, TextPosition::from_utf16_index(5).unwrap(), &text).is_err()
        );
    }

    #[test]
    fn language_updates_preserve_unknown_table_and_entry_fields() {
        let (mut package, storage_id, archive_name) = fixture();
        package
            .update_archive(&archive_name, |archive| {
                let object = archive.object_mut(storage_id).unwrap();
                let index = storage_message_index(object);
                let original = &object.messages[index];
                let data = crate::wire::transform_length_delimited_field(
                    &original.data,
                    LANGUAGE_TABLE_FIELD,
                    |table| {
                        let entries =
                            repeated_length_delimited_payloads(table, TABLE_ENTRIES_FIELD)?;
                        let first = patch_varint_field(entries[0], 77, false, Some(42))?;
                        let table = rewrite_repeated_length_delimited_fields(
                            table,
                            TABLE_ENTRIES_FIELD,
                            &[first],
                        )?;
                        patch_varint_field(&table, 99, false, Some(7))
                    },
                )?;
                object.replace_message(
                    index,
                    RawMessage {
                        type_: original.type_,
                        data,
                    },
                )?;
                Ok(())
            })
            .unwrap();

        let mut editor = super::super::IWorkTextEditor::from_package(package);
        editor
            .set_text_language(
                storage_id,
                TextPosition::from_utf16_index(6).unwrap(),
                TextLanguage::tag("fr").unwrap(),
            )
            .unwrap();
        let package = editor.into_package();
        let location = locate_storage(&package, storage_id).unwrap();
        let archive = package.archive(&location.archive_name).unwrap();
        let object = archive.object(storage_id).unwrap();
        let message = &object.messages[storage_message_index(object)];
        let tables =
            repeated_length_delimited_payloads(&message.data, LANGUAGE_TABLE_FIELD).unwrap();
        assert!(
            parse_wire_fields(tables[0])
                .unwrap()
                .iter()
                .any(|field| field.number == 99)
        );
        let entries = repeated_length_delimited_payloads(tables[0], TABLE_ENTRIES_FIELD).unwrap();
        assert!(
            parse_wire_fields(entries[0])
                .unwrap()
                .iter()
                .any(|field| field.number == 77)
        );
    }

    #[test]
    fn duplicate_language_tables_fail_without_mutation() {
        let (mut package, storage_id, archive_name) = fixture();
        package
            .update_archive(&archive_name, |archive| {
                let object = archive.object_mut(storage_id).unwrap();
                let index = storage_message_index(object);
                let original = &object.messages[index];
                let table =
                    repeated_length_delimited_payloads(&original.data, LANGUAGE_TABLE_FIELD)?[0];
                let mut data = original.data.clone();
                append_length_delimited_field(&mut data, LANGUAGE_TABLE_FIELD, table)?;
                object.replace_message(
                    index,
                    RawMessage {
                        type_: original.type_,
                        data,
                    },
                )?;
                Ok(())
            })
            .unwrap();

        let mut editor = super::super::IWorkTextEditor::from_package(package);
        let before = editor.to_bytes().unwrap();
        assert!(editor.text_languages(storage_id).is_err());
        assert!(
            editor
                .set_text_language(
                    storage_id,
                    TextPosition::ZERO,
                    TextLanguage::tag("fr").unwrap(),
                )
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before);
    }
}
