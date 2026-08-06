//! Native language-run CRUD for TSWP text-storage objects.

use prost::Message;

use crate::archive::RawMessage;
use crate::protobuf::tswp::{self, StorageArchive};
use crate::wire::{
    patch_length_delimited_field, repeated_length_delimited_payloads,
    rewrite_repeated_length_delimited_fields,
};
use crate::{Error, IWorkPackage, Result};

use super::storage_wire::{
    LocatedStorage, StorageLocation, locate_storage as locate_native_storage,
    locate_storage_with_archive as locate_native_storage_with_archive, update_parsed_archive,
    validate_sorted_boundaries,
};
use litchi_iwa_text::{TextLanguage, TextLanguageRun, TextLanguageTag, TextPosition};

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
                TextPosition::from_utf16_code_units(entry.character_index),
                language_from_native(entry.object.as_deref())?,
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
    let located = locate_storage_with_archive(package, storage_id)?;
    let location = &located.location;
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
    language_from_native(entry.object.as_deref())
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
    let located = locate_storage_with_archive(package, storage_id)?;
    patch_language(package, located, storage_id, position, language)?;
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
    let located = locate_storage_with_archive(package, storage_id)?;
    let location = &located.location;
    require_text_boundary(storage_id, position, &location.storage.text)?;
    let entries = language_entries(storage_id, &location.storage, location.table_present)?;
    validate_entries(storage_id, entries, &location.storage.text)?;
    if !entries
        .iter()
        .any(|entry| entry.character_index == position.utf16_index())
    {
        return Ok(false);
    }
    remove_boundary(package, located, storage_id, position)?;
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
    let located = locate_storage_with_archive(package, storage_id)?;
    let location = &located.location;
    let entries = language_entries(storage_id, &location.storage, location.table_present)?;
    validate_entries(storage_id, entries, &location.storage.text)?;
    if !location.table_present {
        return Ok(false);
    }
    patch_language_table(package, located, storage_id, |_, _| Ok(None))?;
    if !text_languages(package, storage_id)?.is_empty() {
        return Err(Error::InvalidFormat(
            "iWork text-language reset failed validation".to_owned(),
        ));
    }
    Ok(true)
}

fn locate_storage(package: &IWorkPackage, storage_id: u64) -> Result<StorageLocation> {
    let location = locate_native_storage(package, storage_id, LANGUAGE_TABLE_FIELD, "language")?;
    if location.storage.table_language.is_some() != location.table_present {
        return Err(Error::InvalidFormat(format!(
            "iWork text storage {storage_id} language-table wire state is inconsistent"
        )));
    }
    Ok(location)
}

fn locate_storage_with_archive(package: &IWorkPackage, storage_id: u64) -> Result<LocatedStorage> {
    let located =
        locate_native_storage_with_archive(package, storage_id, LANGUAGE_TABLE_FIELD, "language")?;
    if located.location.storage.table_language.is_some() != located.location.table_present {
        return Err(Error::InvalidFormat(format!(
            "iWork text storage {storage_id} language-table wire state is inconsistent"
        )));
    }
    Ok(located)
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
    located: LocatedStorage,
    storage_id: u64,
    position: TextPosition,
    language: &TextLanguage,
) -> Result<()> {
    patch_language_table(package, located, storage_id, |table, storage| {
        require_text_boundary(storage_id, position, &storage.text)?;
        match table {
            Some(table) => patch_existing_table(table, position, language).map(Some),
            None => new_table(position, language).map(Some),
        }
    })
}

fn remove_boundary(
    package: &mut IWorkPackage,
    located: LocatedStorage,
    storage_id: u64,
    position: TextPosition,
) -> Result<()> {
    patch_language_table(package, located, storage_id, |table, storage| {
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
    located: LocatedStorage,
    storage_id: u64,
    transform: F,
) -> Result<()>
where
    F: FnOnce(Option<&[u8]>, &StorageArchive) -> Result<Option<Vec<u8>>>,
{
    let LocatedStorage { location, archive } = located;
    let archive_name = location.archive_name.clone();
    update_parsed_archive(package, &archive_name, archive, |archive| {
        let object = archive.object_mut(storage_id).ok_or_else(|| {
            Error::InvalidFormat(format!("iWork text storage {storage_id} is missing"))
        })?;
        if object.archive_info.identifier != Some(location.object_id) {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} changed archive identity"
            )));
        }
        let original = object.messages.get(location.message_index).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork text storage {storage_id} lost its resolved message"
            ))
        })?;
        if original.type_ != location.message_type {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} changed its resolved message type"
            )));
        }
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
            location.message_index,
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
            native_language_value(language).map(str::as_bytes),
        )?;
        entry.object = native_language_value(language).map(str::to_owned);
    } else {
        let entry = tswp::string_attribute_table::StringAttribute {
            character_index: position.utf16_index(),
            object: native_language_value(language).map(str::to_owned),
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
        object: native_language_value(language).map(str::to_owned),
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
        language_from_native(entry.object.as_deref())?;
        previous = Some(entry.character_index);
    }
    validate_sorted_boundaries(
        storage_id,
        entries.iter().map(|entry| entry.character_index),
        text,
    )
}

fn require_text_boundary(storage_id: u64, position: TextPosition, text: &[String]) -> Result<()> {
    super::storage_wire::require_text_boundary(storage_id, position.utf16_index(), text)
}

fn language_from_native(value: Option<&str>) -> Result<TextLanguage> {
    Ok(value
        .map(TextLanguage::tag)
        .transpose()?
        .unwrap_or_default())
}

fn native_language_value(language: &TextLanguage) -> Option<&str> {
    language.as_tag().map(TextLanguageTag::as_str)
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

    fn storage_message_index(package: &IWorkPackage, storage_id: u64) -> usize {
        locate_storage(package, storage_id).unwrap().message_index
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
        let message_index = storage_message_index(&package, storage_id);
        package
            .update_archive(&archive_name, |archive| {
                let object = archive.object_mut(storage_id).unwrap();
                let original = &object.messages[message_index];
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
                    message_index,
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
        let message_index = storage_message_index(&package, storage_id);
        let archive = package.archive(&location.archive_name).unwrap();
        let object = archive.object(storage_id).unwrap();
        let message = &object.messages[message_index];
        let tables =
            repeated_length_delimited_payloads(&message.data, LANGUAGE_TABLE_FIELD).unwrap();
        assert!(
            parse_wire_fields(tables[0])
                .unwrap()
                .iter()
                .any(|field| field.number() == 99)
        );
        let entries = repeated_length_delimited_payloads(tables[0], TABLE_ENTRIES_FIELD).unwrap();
        assert!(
            parse_wire_fields(entries[0])
                .unwrap()
                .iter()
                .any(|field| field.number() == 77)
        );
    }

    #[test]
    fn language_updates_use_the_resolved_storage_with_a_2022_style_sibling() {
        let (mut package, storage_id, archive_name) = fixture();
        let style_data = tswp::ParagraphStyleArchive {
            super_: crate::protobuf::tss::StyleArchive::default(),
            ..Default::default()
        }
        .encode_to_vec();
        package
            .update_archive(&archive_name, |archive| {
                archive
                    .object_mut(storage_id)
                    .unwrap()
                    .push_message(RawMessage {
                        type_: 2_022,
                        data: style_data.clone(),
                    })?;
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
        assert_eq!(
            object.messages[location.message_index].type_,
            location.message_type
        );
        let sibling = object
            .messages
            .iter()
            .find(|message| message.type_ == 2_022)
            .unwrap();
        assert_eq!(sibling.data, style_data);
    }

    #[test]
    fn duplicate_language_tables_fail_without_mutation() {
        let (mut package, storage_id, archive_name) = fixture();
        let message_index = storage_message_index(&package, storage_id);
        package
            .update_archive(&archive_name, |archive| {
                let object = archive.object_mut(storage_id).unwrap();
                let original = &object.messages[message_index];
                let table =
                    repeated_length_delimited_payloads(&original.data, LANGUAGE_TABLE_FIELD)?[0];
                let mut data = original.data.clone();
                append_length_delimited_field(&mut data, LANGUAGE_TABLE_FIELD, table)?;
                object.replace_message(
                    message_index,
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
