//! Transactional editing of shared iWork text storage objects.

use std::ops::Range;
use std::path::Path;

use prost::Message;

use crate::archive::RawMessage;
use crate::protobuf::tswp::{
    ObjectAttributeTable, OverlappingFieldAttributeTable, ParaDataAttributeTable, StorageArchive,
    StringAttributeTable,
};
use crate::wire::{
    parse_wire_fields, patch_nested_varint_field, patch_varint_field,
    repeated_length_delimited_payloads, rewrite_repeated_length_delimited_fields,
    transform_length_delimited_field,
};
use crate::{Error, IWorkPackage, Result};

use super::paragraph_alignment::{
    paragraph_alignment, reset_paragraph_alignment, set_paragraph_alignment,
};
use super::style::TextAlignment;

const STORAGE_MESSAGE_TYPES: &[u32] = &[2001, 2022];

/// A discoverable text storage within an iWork package.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TextStorageInfo {
    pub object_id: u64,
    pub message_type: u32,
    pub kind: Option<i32>,
    pub text: String,
}

/// Mutable editor for the TSWP text layer shared by Pages, Numbers, and Keynote.
#[derive(Debug, Clone)]
pub struct IWorkTextEditor {
    package: IWorkPackage,
}

impl IWorkTextEditor {
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        Ok(Self {
            package: IWorkPackage::open(path)?,
        })
    }

    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        Ok(Self {
            package: IWorkPackage::from_bytes(bytes)?,
        })
    }

    pub fn from_package(package: IWorkPackage) -> Self {
        Self { package }
    }

    pub fn storages(&self) -> Result<Vec<TextStorageInfo>> {
        let mut storages = Vec::new();
        for name in self.package.iwa_entry_names() {
            let archive = self.package.archive(name)?;
            for object in archive.objects {
                let Some(object_id) = object.archive_info.identifier else {
                    continue;
                };
                for message in object.messages {
                    if !STORAGE_MESSAGE_TYPES.contains(&message.type_) {
                        continue;
                    }
                    let Ok(storage) = StorageArchive::decode(message.data.as_slice()) else {
                        continue;
                    };
                    storages.push(TextStorageInfo {
                        object_id,
                        message_type: message.type_,
                        kind: storage.kind,
                        text: storage.text.concat(),
                    });
                }
            }
        }
        storages.sort_by_key(|storage| storage.object_id);
        Ok(storages)
    }

    pub fn storage(&self, object_id: u64) -> Result<TextStorageInfo> {
        let archive_name = find_storage_archive(&self.package, object_id)?;
        let archive = self.package.archive(&archive_name)?;
        let object = archive.object(object_id).ok_or_else(|| {
            Error::ParseError(format!("Text storage object {object_id} not found"))
        })?;
        object
            .messages
            .iter()
            .find_map(|message| {
                if !STORAGE_MESSAGE_TYPES.contains(&message.type_) {
                    return None;
                }
                StorageArchive::decode(message.data.as_slice())
                    .ok()
                    .map(|storage| TextStorageInfo {
                        object_id,
                        message_type: message.type_,
                        kind: storage.kind,
                        text: storage.text.concat(),
                    })
            })
            .ok_or_else(|| {
                Error::ParseError(format!(
                    "Object {object_id} has no writable TSWP storage payload"
                ))
            })
    }

    /// Replace a UTF-16 range, matching the indexing used by iWork attributes.
    pub fn replace_text(
        &mut self,
        object_id: u64,
        range: Range<usize>,
        replacement: &str,
    ) -> Result<()> {
        if range.start > range.end {
            return Err(Error::ParseError(
                "Text replacement range starts after it ends".to_string(),
            ));
        }
        let mut staged = self.package.clone();
        replace_storage_text(&mut staged, object_id, range, replacement)?;
        let bytes = staged.to_bytes()?;
        IWorkPackage::from_bytes(&bytes)?;
        self.package = staged;
        Ok(())
    }

    pub fn set_text(&mut self, object_id: u64, replacement: &str) -> Result<()> {
        let storage = self.storage(object_id)?;
        self.replace_text(
            object_id,
            0..storage.text.encode_utf16().count(),
            replacement,
        )
    }

    /// Read the effective uniform paragraph alignment of a text storage.
    pub fn paragraph_alignment(&self, object_id: u64) -> Result<TextAlignment> {
        paragraph_alignment(&self.package, object_id)
    }

    /// Set one alignment for every paragraph in a uniformly styled text storage.
    ///
    /// Rich text containing multiple paragraph-style boundaries is rejected so
    /// this whole-storage operation can never flatten unrelated formatting.
    pub fn set_paragraph_alignment(
        &mut self,
        object_id: u64,
        alignment: TextAlignment,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        set_paragraph_alignment(&mut staged, object_id, alignment)?;
        let bytes = staged.to_bytes()?;
        let verified = IWorkPackage::from_bytes(&bytes)?;
        if paragraph_alignment(&verified, object_id)? != alignment {
            return Err(Error::InvalidFormat(
                "iWork paragraph-alignment update failed round-trip validation".to_owned(),
            ));
        }
        self.package = staged;
        Ok(())
    }

    /// Remove a private minimal alignment override and restore its parent style.
    pub fn reset_paragraph_alignment(&mut self, object_id: u64) -> Result<bool> {
        let mut staged = self.package.clone();
        let changed = reset_paragraph_alignment(&mut staged, object_id)?;
        if changed {
            let bytes = staged.to_bytes()?;
            IWorkPackage::from_bytes(&bytes)?;
            self.package = staged;
        }
        Ok(changed)
    }

    pub fn package(&self) -> &IWorkPackage {
        &self.package
    }

    pub fn into_package(self) -> IWorkPackage {
        self.package
    }

    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        self.package.to_bytes()
    }

    pub fn save(&self, path: impl AsRef<Path>) -> Result<()> {
        self.package.save(path)
    }
}

fn replace_storage_text(
    package: &mut IWorkPackage,
    object_id: u64,
    range: Range<usize>,
    replacement: &str,
) -> Result<()> {
    let archive_name = find_storage_archive(package, object_id)?;
    package.update_archive(&archive_name, |archive| {
        let object = archive.object_mut(object_id).ok_or_else(|| {
            Error::ParseError(format!("Text storage object {object_id} not found"))
        })?;
        let message_index = object
            .messages
            .iter()
            .position(|message| {
                STORAGE_MESSAGE_TYPES.contains(&message.type_)
                    && StorageArchive::decode(message.data.as_slice()).is_ok()
            })
            .ok_or_else(|| {
                Error::ParseError(format!(
                    "Object {object_id} has no writable TSWP storage payload"
                ))
            })?;
        let message_type = object.messages[message_index].type_;
        let original = object.messages[message_index].data.as_slice();
        let mut storage = StorageArchive::decode(original)?;
        let previous_references = storage_object_references(&storage);
        let current = storage.text.concat();
        let start = utf16_to_byte_index(&current, range.start)?;
        let end = utf16_to_byte_index(&current, range.end)?;
        let mut updated = String::with_capacity(
            current
                .len()
                .saturating_sub(end - start)
                .saturating_add(replacement.len()),
        );
        updated.push_str(&current[..start]);
        updated.push_str(replacement);
        updated.push_str(&current[end..]);

        let replacement_units = replacement.encode_utf16().count();
        adjust_storage_attributes(&mut storage, range.clone(), replacement_units)?;
        storage.text = if updated.is_empty() {
            if current.is_empty() {
                storage.text.clone()
            } else {
                Vec::new()
            }
        } else {
            vec![updated]
        };
        let data = patch_storage_text_wire(original, &range, replacement_units, &storage.text)?;
        if StorageArchive::decode(data.as_slice())? != storage {
            return Err(Error::InvalidFormat(
                "TSWP text-storage wire patch failed validation".to_owned(),
            ));
        }
        let current_references = storage_object_references(&storage);
        let removed_references = previous_references
            .into_iter()
            .filter(|identifier| !current_references.contains(identifier))
            .collect::<std::collections::HashSet<_>>();
        object.replace_message(
            message_index,
            RawMessage {
                type_: message_type,
                data,
            },
        )?;
        object.archive_info.message_infos[message_index]
            .object_references
            .retain(|identifier| !removed_references.contains(identifier));
        for field in &mut object.archive_info.message_infos[message_index].field_infos {
            field
                .object_references
                .retain(|identifier| !removed_references.contains(identifier));
        }
        Ok(())
    })
}

fn find_storage_archive(package: &IWorkPackage, object_id: u64) -> Result<String> {
    let mut found = None;
    for name in package.iwa_entry_names() {
        let archive = package.archive(name)?;
        let Some(object) = archive.object(object_id) else {
            continue;
        };
        if !object.messages.iter().any(|message| {
            STORAGE_MESSAGE_TYPES.contains(&message.type_)
                && StorageArchive::decode(message.data.as_slice()).is_ok()
        }) {
            continue;
        }
        if found.replace(name.to_owned()).is_some() {
            return Err(Error::Archive(format!(
                "Text storage object {object_id} occurs in multiple IWA components"
            )));
        }
    }
    found.ok_or_else(|| Error::ParseError(format!("Text storage object {object_id} not found")))
}

fn utf16_to_byte_index(text: &str, target: usize) -> Result<usize> {
    if target == 0 {
        return Ok(0);
    }
    let mut units = 0usize;
    for (byte_index, character) in text.char_indices() {
        if units == target {
            return Ok(byte_index);
        }
        units += character.len_utf16();
        if units > target {
            return Err(Error::ParseError(format!(
                "UTF-16 index {target} splits a surrogate pair"
            )));
        }
    }
    if units == target {
        Ok(text.len())
    } else {
        Err(Error::ParseError(format!(
            "UTF-16 index {target} exceeds text length {units}"
        )))
    }
}

fn adjust_storage_attributes(
    storage: &mut StorageArchive,
    range: Range<usize>,
    replacement_units: usize,
) -> Result<()> {
    for table in [
        &mut storage.table_para_style,
        &mut storage.table_list_style,
        &mut storage.table_char_style,
        &mut storage.table_layout_style,
    ] {
        adjust_object_table(table, &range, replacement_units, true)?;
    }
    for table in [
        &mut storage.table_para_data,
        &mut storage.table_para_starts,
        &mut storage.table_para_bidi,
    ] {
        adjust_para_table(table, &range, replacement_units)?;
    }
    for table in [&mut storage.table_language, &mut storage.table_dictation] {
        adjust_string_table(table, &range, replacement_units)?;
    }
    for table in [
        &mut storage.table_attachment,
        &mut storage.table_smartfield,
        &mut storage.table_bookmark,
        &mut storage.table_footnote,
        &mut storage.table_rubyfield,
        &mut storage.table_insertion,
        &mut storage.table_deletion,
        &mut storage.table_highlight,
        &mut storage.table_tatechuyoko,
    ] {
        adjust_object_table(table, &range, replacement_units, false)?;
    }
    // A drop-cap record is a paragraph-start boundary. Numbers commonly emits
    // an explicit index-zero entry with no style reference as a sentinel, and
    // replacing the paragraph text must not erase that structural marker.
    adjust_object_table(
        &mut storage.table_drop_cap_style,
        &range,
        replacement_units,
        true,
    )?;
    // A section boundary identifies the section owning text inserted exactly
    // at that boundary. In particular, the mandatory first boundary must stay
    // at UTF-16 index zero when text is inserted into an empty Pages body.
    adjust_object_table(&mut storage.table_section, &range, replacement_units, true)?;
    for table in [
        &mut storage.table_overlapping_highlight,
        &mut storage.table_pencil_annotation,
    ] {
        adjust_overlapping_table(table, &range, replacement_units)?;
    }
    Ok(())
}

fn patch_storage_text_wire(
    original: &[u8],
    range: &Range<usize>,
    replacement_units: usize,
    text: &[String],
) -> Result<Vec<u8>> {
    let text = text
        .iter()
        .map(|value| value.as_bytes().to_vec())
        .collect::<Vec<_>>();
    let mut data = rewrite_repeated_length_delimited_fields(original, 3, &text)?;
    for field in [5, 7, 8, 12, 17, 28] {
        data = transform_optional_table(&data, field, |table| {
            adjust_index_table_wire(table, range, replacement_units, true)
        })?;
    }
    for field in [9, 11, 15, 16, 18, 21, 22, 23, 27] {
        data = transform_optional_table(&data, field, |table| {
            adjust_index_table_wire(table, range, replacement_units, false)
        })?;
    }
    for field in [6, 14, 19, 20, 24] {
        data = transform_optional_table(&data, field, |table| {
            adjust_index_table_wire(table, range, replacement_units, true)
        })?;
    }
    for field in [25, 26] {
        data = transform_optional_table(&data, field, |table| {
            adjust_overlapping_table_wire(table, range, replacement_units)
        })?;
    }
    Ok(data)
}

fn transform_optional_table<F>(data: &[u8], field_number: u32, transform: F) -> Result<Vec<u8>>
where
    F: FnOnce(&[u8]) -> Result<Vec<u8>>,
{
    match repeated_length_delimited_payloads(data, field_number)?.len() {
        0 => Ok(data.to_vec()),
        1 => transform_length_delimited_field(data, field_number, transform),
        count => Err(Error::InvalidFormat(format!(
            "singular TSWP storage table field {field_number} occurs {count} times"
        ))),
    }
}

fn adjust_index_table_wire(
    table: &[u8],
    range: &Range<usize>,
    replacement_units: usize,
    retain_start_boundary: bool,
) -> Result<Vec<u8>> {
    let mut entries = repeated_length_delimited_payloads(table, 1)?
        .into_iter()
        .enumerate()
        .filter_map(|(order, entry)| {
            let index = required_u32_varint(entry, 1);
            match index.and_then(|index| {
                adjust_index(index, range, replacement_units, retain_start_boundary)
            }) {
                Ok(Some(index)) => Some(
                    patch_varint_field(entry, 1, true, Some(u64::from(index)))
                        .map(|entry| (index, order, entry)),
                ),
                Ok(None) => None,
                Err(error) => Some(Err(error)),
            }
        })
        .collect::<Result<Vec<_>>>()?;
    entries.sort_by_key(|(index, order, _)| (*index, *order));
    entries.dedup_by_key(|(index, _, _)| *index);
    let entries = entries
        .into_iter()
        .map(|(_, _, entry)| entry)
        .collect::<Vec<_>>();
    rewrite_repeated_length_delimited_fields(table, 1, &entries)
}

fn adjust_overlapping_table_wire(
    table: &[u8],
    replacement: &Range<usize>,
    replacement_units: usize,
) -> Result<Vec<u8>> {
    let entries = repeated_length_delimited_payloads(table, 1)?
        .into_iter()
        .filter_map(|entry| {
            let adjusted = (|| {
                let ranges = repeated_length_delimited_payloads(entry, 1)?;
                if ranges.len() != 1 {
                    return Err(Error::InvalidFormat(format!(
                        "TSWP overlapping attribute range occurs {} times",
                        ranges.len()
                    )));
                }
                let start = usize::try_from(required_u32_varint(ranges[0], 1)?)
                    .map_err(|_| Error::ParseError("Text attribute index overflow".to_owned()))?;
                let length = usize::try_from(required_u32_varint(ranges[0], 2)?)
                    .map_err(|_| Error::ParseError("Text attribute length overflow".to_owned()))?;
                let end = start
                    .checked_add(length)
                    .ok_or_else(|| Error::ParseError("Text attribute range overflow".to_owned()))?;
                if end <= replacement.start {
                    return Ok(Some(entry.to_vec()));
                }
                if start >= replacement.end {
                    let shifted = shift_index(start, replacement, replacement_units)?;
                    return patch_nested_varint_field(
                        entry,
                        &[1, 1],
                        true,
                        Some(u64::from(shifted)),
                    )
                    .map(Some);
                }
                Ok(None)
            })();
            match adjusted {
                Ok(Some(entry)) => Some(Ok(entry)),
                Ok(None) => None,
                Err(error) => Some(Err(error)),
            }
        })
        .collect::<Result<Vec<_>>>()?;
    rewrite_repeated_length_delimited_fields(table, 1, &entries)
}

fn required_u32_varint(data: &[u8], field_number: u32) -> Result<u32> {
    let fields = parse_wire_fields(data)?;
    let matches = fields
        .iter()
        .filter(|field| field.number == field_number)
        .collect::<Vec<_>>();
    if matches.len() != 1 || matches[0].wire_type != 0 {
        return Err(Error::InvalidFormat(format!(
            "required protobuf varint field {field_number} occurs {} times or has the wrong wire type",
            matches.len()
        )));
    }
    let field = matches[0];
    let (value, length) = crate::varint::decode_varint_from_bytes(&data[field.key_end..field.end])
        .map_err(|error| Error::InvalidFormat(format!("invalid protobuf varint: {error}")))?;
    if field.key_end + length != field.end {
        return Err(Error::InvalidFormat(
            "protobuf varint field has trailing bytes".to_owned(),
        ));
    }
    u32::try_from(value).map_err(|_| Error::InvalidFormat("protobuf varint exceeds u32".to_owned()))
}

fn adjust_object_table(
    table: &mut Option<ObjectAttributeTable>,
    range: &Range<usize>,
    replacement_units: usize,
    retain_start_boundary: bool,
) -> Result<()> {
    let Some(table) = table else {
        return Ok(());
    };
    table.entries = table
        .entries
        .drain(..)
        .filter_map(|mut entry| {
            adjust_index(
                entry.character_index,
                range,
                replacement_units,
                retain_start_boundary,
            )
            .transpose()
            .map(|result| {
                result.map(|index| {
                    entry.character_index = index;
                    entry
                })
            })
        })
        .collect::<Result<Vec<_>>>()?;
    deduplicate_object_entries(&mut table.entries);
    Ok(())
}

fn adjust_para_table(
    table: &mut Option<ParaDataAttributeTable>,
    range: &Range<usize>,
    replacement_units: usize,
) -> Result<()> {
    let Some(table) = table else {
        return Ok(());
    };
    table.entries = table
        .entries
        .drain(..)
        .filter_map(|mut entry| {
            adjust_index(entry.character_index, range, replacement_units, true)
                .transpose()
                .map(|result| {
                    result.map(|index| {
                        entry.character_index = index;
                        entry
                    })
                })
        })
        .collect::<Result<Vec<_>>>()?;
    table.entries.sort_by_key(|entry| entry.character_index);
    table.entries.dedup_by_key(|entry| entry.character_index);
    Ok(())
}

fn adjust_string_table(
    table: &mut Option<StringAttributeTable>,
    range: &Range<usize>,
    replacement_units: usize,
) -> Result<()> {
    let Some(table) = table else {
        return Ok(());
    };
    table.entries = table
        .entries
        .drain(..)
        .filter_map(|mut entry| {
            adjust_index(entry.character_index, range, replacement_units, true)
                .transpose()
                .map(|result| {
                    result.map(|index| {
                        entry.character_index = index;
                        entry
                    })
                })
        })
        .collect::<Result<Vec<_>>>()?;
    table.entries.sort_by_key(|entry| entry.character_index);
    table.entries.dedup_by_key(|entry| entry.character_index);
    Ok(())
}

fn adjust_overlapping_table(
    table: &mut Option<OverlappingFieldAttributeTable>,
    replacement: &Range<usize>,
    replacement_units: usize,
) -> Result<()> {
    let Some(table) = table else {
        return Ok(());
    };
    let mut adjusted = Vec::new();
    for mut entry in table.entries.drain(..) {
        let start = entry.range.location as usize;
        let end = start
            .checked_add(entry.range.length as usize)
            .ok_or_else(|| Error::ParseError("Text attribute range overflow".to_string()))?;
        if end <= replacement.start {
            adjusted.push(entry);
        } else if start >= replacement.end {
            entry.range.location = shift_index(start, replacement, replacement_units)?;
            adjusted.push(entry);
        }
        // An annotation intersecting replaced text is intentionally removed.
    }
    table.entries = adjusted;
    Ok(())
}

fn adjust_index(
    index: u32,
    range: &Range<usize>,
    replacement_units: usize,
    retain_start_boundary: bool,
) -> Result<Option<u32>> {
    let index_usize = index as usize;
    if index_usize < range.start || (retain_start_boundary && index_usize == range.start) {
        return Ok(Some(index));
    }
    if index_usize < range.end {
        return Ok(None);
    }
    Ok(Some(shift_index(index_usize, range, replacement_units)?))
}

fn shift_index(index: usize, range: &Range<usize>, replacement_units: usize) -> Result<u32> {
    let removed = range.end - range.start;
    let shifted = if replacement_units >= removed {
        index.checked_add(replacement_units - removed)
    } else {
        index.checked_sub(removed - replacement_units)
    }
    .ok_or_else(|| Error::ParseError("Text attribute index overflow".to_string()))?;
    u32::try_from(shifted)
        .map_err(|_| Error::ParseError("Text attribute index exceeds u32".to_string()))
}

fn deduplicate_object_entries(
    entries: &mut Vec<crate::protobuf::tswp::object_attribute_table::ObjectAttribute>,
) {
    entries.sort_by_key(|entry| entry.character_index);
    entries.dedup_by_key(|entry| entry.character_index);
}

fn storage_object_references(storage: &StorageArchive) -> Vec<u64> {
    let mut references = Vec::new();
    if let Some(reference) = &storage.style_sheet {
        references.push(reference.identifier);
    }
    for table in [
        &storage.table_para_style,
        &storage.table_list_style,
        &storage.table_char_style,
        &storage.table_attachment,
        &storage.table_smartfield,
        &storage.table_layout_style,
        &storage.table_bookmark,
        &storage.table_footnote,
        &storage.table_section,
        &storage.table_rubyfield,
        &storage.table_insertion,
        &storage.table_deletion,
        &storage.table_highlight,
        &storage.table_tatechuyoko,
        &storage.table_drop_cap_style,
    ]
    .into_iter()
    .flatten()
    {
        references.extend(
            table
                .entries
                .iter()
                .filter_map(|entry| entry.object.as_ref().map(|value| value.identifier)),
        );
    }
    for table in [
        &storage.table_overlapping_highlight,
        &storage.table_pencil_annotation,
    ]
    .into_iter()
    .flatten()
    {
        references.extend(table.entries.iter().map(|entry| entry.field.identifier));
    }
    references.sort_unstable();
    references.dedup();
    references
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::archive::{Archive, ArchiveObject};
    use crate::protobuf::tsp::{Range as TspRange, Reference};
    use crate::protobuf::tswp::object_attribute_table::ObjectAttribute;
    use crate::protobuf::tswp::overlapping_field_attribute_table::OverlappingFieldAttribute;

    #[test]
    fn whole_text_replacement_preserves_drop_cap_sentinel_exactly() {
        let storage = StorageArchive {
            text: vec!["Source".to_owned()],
            table_drop_cap_style: Some(ObjectAttributeTable {
                entries: vec![ObjectAttribute {
                    character_index: 0,
                    object: None,
                }],
            }),
            ..Default::default()
        };
        let mut editor = IWorkTextEditor::from_package(test_package(storage));
        let baseline = editor.to_bytes().unwrap();
        editor.set_text(42, "Changed text").unwrap();
        editor.set_text(42, "Source").unwrap();
        assert_eq!(editor.to_bytes().unwrap(), baseline);
    }

    #[test]
    fn replacement_uses_utf16_and_shifts_style_boundaries() {
        let storage = StorageArchive {
            text: vec!["A🚀BC".to_string()],
            table_char_style: Some(ObjectAttributeTable {
                entries: vec![
                    attribute(0, 10),
                    attribute(1, 11),
                    attribute(3, 12),
                    attribute(5, 13),
                ],
            }),
            table_attachment: Some(ObjectAttributeTable {
                entries: vec![attribute(1, 99)],
            }),
            ..Default::default()
        };
        let mut editor = IWorkTextEditor::from_package(test_package(storage));
        editor.replace_text(42, 1..4, "東京").unwrap();

        let package = editor.into_package();
        let archive = package.archive("Index/Document.iwa").unwrap();
        let message = &archive.object(42).unwrap().messages[0];
        let storage = StorageArchive::decode(message.data.as_slice()).unwrap();
        assert_eq!(storage.text, ["A東京C"]);
        let indexes = storage.table_char_style.unwrap().entries;
        assert_eq!(
            indexes
                .iter()
                .map(|entry| entry.character_index)
                .collect::<Vec<_>>(),
            [0, 1, 4]
        );
        assert!(storage.table_attachment.unwrap().entries.is_empty());
    }

    #[test]
    fn invalid_surrogate_boundary_is_transactional() {
        let storage = StorageArchive {
            text: vec!["🚀".to_string()],
            ..Default::default()
        };
        let mut editor = IWorkTextEditor::from_package(test_package(storage));
        let before = editor.to_bytes().unwrap();
        assert!(editor.replace_text(42, 1..2, "x").is_err());
        assert_eq!(editor.to_bytes().unwrap(), before);
    }

    #[test]
    fn insertion_retains_section_boundary_at_replacement_start() {
        let storage = StorageArchive {
            text: Vec::new(),
            table_section: Some(ObjectAttributeTable {
                entries: vec![attribute(0, 77)],
            }),
            ..Default::default()
        };
        let mut editor = IWorkTextEditor::from_package(test_package(storage));
        editor.replace_text(42, 0..0, "Body").unwrap();

        let archive = editor.package().archive("Index/Document.iwa").unwrap();
        let storage =
            StorageArchive::decode(archive.object(42).unwrap().messages[0].data.as_slice())
                .unwrap();
        assert_eq!(storage.text, ["Body"]);
        assert_eq!(storage.table_section.unwrap().entries[0].character_index, 0);
    }

    #[test]
    fn text_replace_restore_preserves_deep_unknown_fields_exactly() {
        let storage = StorageArchive {
            text: vec!["A🚀BC".to_owned()],
            table_char_style: Some(ObjectAttributeTable {
                entries: vec![attribute(0, 10), attribute(5, 11)],
            }),
            table_overlapping_highlight: Some(OverlappingFieldAttributeTable {
                entries: vec![OverlappingFieldAttribute {
                    range: TspRange {
                        location: 4,
                        length: 1,
                    },
                    field: Reference {
                        identifier: 77,
                        ..Default::default()
                    },
                }],
            }),
            ..Default::default()
        };
        let mut package = test_package(storage);
        package
            .update_archive("Index/Document.iwa", |archive| {
                let object = archive.object_mut(42).unwrap();
                let message_type = object.messages[0].type_;
                let mut data = crate::wire::transform_length_delimited_fields_at_path(
                    object.messages[0].data.as_slice(),
                    &[8],
                    |table| {
                        let mut table = crate::wire::transform_length_delimited_fields_at_path(
                            table,
                            &[1],
                            |entry| {
                                let mut entry =
                                    crate::wire::transform_length_delimited_fields_at_path(
                                        entry,
                                        &[2],
                                        |reference| {
                                            let mut reference = reference.to_vec();
                                            append_unknown_varint(&mut reference, 96, 960);
                                            Ok(reference)
                                        },
                                    )?;
                                append_unknown_varint(&mut entry, 97, 970);
                                Ok(entry)
                            },
                        )?;
                        append_unknown_varint(&mut table, 98, 980);
                        Ok(table)
                    },
                )?;
                data = crate::wire::transform_length_delimited_fields_at_path(
                    &data,
                    &[25],
                    |table| {
                        let mut table = crate::wire::transform_length_delimited_fields_at_path(
                            table,
                            &[1],
                            |entry| {
                                let mut entry =
                                    crate::wire::transform_length_delimited_fields_at_path(
                                        entry,
                                        &[1],
                                        |range| {
                                            let mut range = range.to_vec();
                                            append_unknown_varint(&mut range, 93, 930);
                                            Ok(range)
                                        },
                                    )?;
                                entry = crate::wire::transform_length_delimited_fields_at_path(
                                    &entry,
                                    &[2],
                                    |reference| {
                                        let mut reference = reference.to_vec();
                                        append_unknown_varint(&mut reference, 92, 920);
                                        Ok(reference)
                                    },
                                )?;
                                append_unknown_varint(&mut entry, 94, 940);
                                Ok(entry)
                            },
                        )?;
                        append_unknown_varint(&mut table, 95, 950);
                        Ok(table)
                    },
                )?;
                append_unknown_varint(&mut data, 99, 990);
                object.replace_message(
                    0,
                    RawMessage {
                        type_: message_type,
                        data,
                    },
                )?;
                object.archive_info.message_infos[0].object_references = vec![10, 11, 77];
                Ok(())
            })
            .unwrap();
        let before = package
            .archive("Index/Document.iwa")
            .unwrap()
            .to_bytes()
            .unwrap();
        let mut editor = IWorkTextEditor::from_package(package);
        editor.replace_text(42, 1..3, "X").unwrap();
        editor.replace_text(42, 1..2, "🚀").unwrap();
        let after = editor
            .package()
            .archive("Index/Document.iwa")
            .unwrap()
            .to_bytes()
            .unwrap();
        assert_eq!(after, before);
    }

    #[test]
    fn duplicate_attribute_indexes_fail_transactionally() {
        let storage = StorageArchive {
            text: vec!["Body".to_owned()],
            table_char_style: Some(ObjectAttributeTable {
                entries: vec![attribute(0, 10)],
            }),
            ..Default::default()
        };
        let mut package = test_package(storage);
        package
            .update_archive("Index/Document.iwa", |archive| {
                let object = archive.object_mut(42).unwrap();
                let message_type = object.messages[0].type_;
                let data = crate::wire::transform_length_delimited_fields_at_path(
                    object.messages[0].data.as_slice(),
                    &[8, 1],
                    |entry| {
                        let mut entry = entry.to_vec();
                        entry.extend(crate::varint::encode_varint(8));
                        entry.extend(crate::varint::encode_varint(0));
                        Ok(entry)
                    },
                )?;
                object.replace_message(
                    0,
                    RawMessage {
                        type_: message_type,
                        data,
                    },
                )?;
                Ok(())
            })
            .unwrap();
        let mut editor = IWorkTextEditor::from_package(package);
        let before = editor.to_bytes().unwrap();
        assert!(editor.replace_text(42, 0..0, "X").is_err());
        assert_eq!(editor.to_bytes().unwrap(), before);
    }

    fn attribute(character_index: u32, identifier: u64) -> ObjectAttribute {
        ObjectAttribute {
            character_index,
            object: Some(Reference {
                identifier,
                ..Default::default()
            }),
        }
    }

    fn append_unknown_varint(data: &mut Vec<u8>, field_number: u32, value: u64) {
        data.extend(crate::varint::encode_varint(u64::from(field_number) << 3));
        data.extend(crate::varint::encode_varint(value));
    }

    fn test_package(storage: StorageArchive) -> IWorkPackage {
        let object = ArchiveObject::new(
            42,
            vec![RawMessage {
                type_: 2001,
                data: storage.encode_to_vec(),
            }],
        )
        .unwrap();
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
