//! A storage's list tables: list-style boundaries (field 7), paragraph
//! levels (field 6) and numbering restarts (field 14).

use litchi_iwa_common::wire::{
    patch_length_delimited_field, patch_varint_field, rewrite_repeated_length_delimited_fields,
    transform_length_delimited_field,
};

use crate::native::{self, Budget, NativeError, NativeResult, Writer};

pub(super) const LIST_STYLE_TABLE_FIELD: u32 = 7;
const PARAGRAPH_DATA_TABLE_FIELD: u32 = 6;
const PARAGRAPH_START_TABLE_FIELD: u32 = 14;
const TABLE_ENTRIES_FIELD: u32 = 1;
const ENTRY_INDEX_FIELD: u32 = 1;
const ENTRY_OBJECT_FIELD: u32 = 2;
const ENTRY_FIRST_FIELD: u32 = 2;
const ENTRY_SECOND_FIELD: u32 = 3;
const REFERENCE_IDENTIFIER_FIELD: u32 = 1;
const TEXT_FIELD: u32 = 3;

/// The decoded fields of a storage the list engine reads.
pub(super) struct Source<'a> {
    pub(super) starts: Vec<u32>,
    pub(super) list_table: Option<&'a [u8]>,
    pub(super) data_table: Option<&'a [u8]>,
    pub(super) start_table: Option<&'a [u8]>,
}

fn single<'a>(found: &[&'a [u8]], reason: &'static str) -> NativeResult<Option<&'a [u8]>> {
    match found {
        [] => Ok(None),
        [table] => Ok(Some(*table)),
        _ => Err(NativeError::Invalid(reason)),
    }
}

/// Paragraph starts as the host computes them: index zero, then after each
/// `\n` (a `\r\n` pair counts once), `\r`, U+2028 and U+2029.
fn paragraph_starts(fragments: &[&str]) -> NativeResult<Vec<u32>> {
    let mut starts = vec![0u32];
    let mut index = 0u32;
    let mut previous_was_carriage_return = false;
    for fragment in fragments {
        for character in fragment.chars() {
            index = index
                .checked_add(u32::try_from(character.len_utf16()).unwrap_or(2))
                .ok_or(NativeError::Invalid("storage text UTF-16 length overflow"))?;
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

pub(super) fn decode<'a>(storage: &'a [u8], budget: &mut Budget) -> NativeResult<Source<'a>> {
    let mut fragments = Vec::new();
    let mut list = Vec::new();
    let mut data = Vec::new();
    let mut starts = Vec::new();
    native::fields(storage, 0, budget, |field, _| {
        match field.number {
            TEXT_FIELD => {
                let fragment = std::str::from_utf8(field.bytes()?)
                    .map_err(|_| NativeError::Invalid("storage text is not UTF-8"))?;
                native::reserve(&mut fragments, 1)?;
                fragments.push(fragment);
            },
            LIST_STYLE_TABLE_FIELD => {
                native::reserve(&mut list, 1)?;
                list.push(field.bytes()?);
            },
            PARAGRAPH_DATA_TABLE_FIELD => {
                native::reserve(&mut data, 1)?;
                data.push(field.bytes()?);
            },
            PARAGRAPH_START_TABLE_FIELD => {
                native::reserve(&mut starts, 1)?;
                starts.push(field.bytes()?);
            },
            _ => {},
        }
        Ok(())
    })?;
    Ok(Source {
        starts: paragraph_starts(&fragments)?,
        list_table: single(&list, "storage contains several list-style tables")?,
        data_table: single(&data, "storage contains several paragraph-data tables")?,
        start_table: single(&starts, "storage contains several paragraph-start tables")?,
    })
}

/// One `ObjectAttributeTable` entry.
pub(super) struct Reference {
    pub(super) index: u32,
    pub(super) object: Option<u64>,
    pub(super) raw: Vec<u8>,
}

/// One `ParaDataAttributeTable` entry.
pub(super) struct Data {
    pub(super) index: u32,
    pub(super) first: u32,
    pub(super) second: u32,
    pub(super) other_fields: bool,
    pub(super) raw: Vec<u8>,
}

pub(super) fn references(table: &[u8], budget: &mut Budget) -> NativeResult<Vec<Reference>> {
    let mut entries = Vec::new();
    native::fields(table, 1, budget, |field, budget| {
        if field.number != TABLE_ENTRIES_FIELD {
            return Ok(());
        }
        let raw = field.bytes()?;
        let mut index = None;
        let mut index_count = 0;
        let mut object = None;
        let mut object_count = 0;
        native::fields(raw, 2, budget, |inner, budget| {
            match inner.number {
                ENTRY_INDEX_FIELD => {
                    index = Some(inner.uint32()?);
                    index_count += 1;
                },
                ENTRY_OBJECT_FIELD => {
                    let mut reference = native::Reference::default();
                    reference.merge(inner.bytes()?, 3, budget)?;
                    object = Some(reference.identifier);
                    object_count += 1;
                },
                _ => {},
            }
            Ok(())
        })?;
        if index_count != 1 || object_count > 1 {
            return Err(NativeError::Invalid("list-style entry is malformed"));
        }
        native::reserve(&mut entries, 1)?;
        entries.push(Reference {
            index: index.unwrap_or_default(),
            object,
            raw: raw.to_vec(),
        });
        Ok(())
    })?;
    Ok(entries)
}

pub(super) fn data_entries(table: &[u8], budget: &mut Budget) -> NativeResult<Vec<Data>> {
    let mut entries = Vec::new();
    native::fields(table, 1, budget, |field, budget| {
        if field.number != TABLE_ENTRIES_FIELD {
            return Ok(());
        }
        let raw = field.bytes()?;
        let mut entry = Data {
            index: 0,
            first: 0,
            second: 0,
            other_fields: false,
            raw: raw.to_vec(),
        };
        native::fields(raw, 2, budget, |inner, _| {
            match inner.number {
                ENTRY_INDEX_FIELD => entry.index = inner.uint32()?,
                ENTRY_FIRST_FIELD => entry.first = inner.uint32()?,
                ENTRY_SECOND_FIELD => entry.second = inner.uint32()?,
                _ => entry.other_fields = true,
            }
            Ok(())
        })?;
        native::reserve(&mut entries, 1)?;
        entries.push(entry);
        Ok(())
    })?;
    Ok(entries)
}

/// The list-style boundaries `(index, style)` of a storage, validated as
/// the host does.
pub(super) fn boundaries(
    source: &Source<'_>,
    budget: &mut Budget,
) -> NativeResult<Vec<(u32, u64)>> {
    let Some(table) = source.list_table else {
        return Err(NativeError::Invalid("storage has no list-style table"));
    };
    let mut boundaries = Vec::new();
    let mut previous = None;
    for entry in references(table, budget)? {
        if previous.is_some_and(|index| index >= entry.index) {
            return Err(NativeError::Invalid(
                "list-style boundaries are not strictly increasing",
            ));
        }
        if source.starts.binary_search(&entry.index).is_err() {
            return Err(NativeError::Invalid(
                "list-style boundary is not a paragraph start",
            ));
        }
        let Some(style) = entry.object.filter(|identifier| *identifier != 0) else {
            return Err(NativeError::Invalid(
                "storage has an empty list-style reference",
            ));
        };
        native::reserve(&mut boundaries, 1)?;
        boundaries.push((entry.index, style));
        previous = Some(entry.index);
    }
    if boundaries.first().map(|entry| entry.0) != Some(0) {
        return Err(NativeError::Invalid(
            "list style must begin at UTF-16 index zero",
        ));
    }
    Ok(boundaries)
}

/// Rewrite the list-style table to `boundaries`, keeping the bytes of
/// entries at unchanged indexes (the host's `replace_boundary_table`).
pub(super) fn replace(
    storage: &[u8],
    boundaries: &[(u32, u64)],
    budget: &mut Budget,
) -> NativeResult<Vec<u8>> {
    let limits = budget.limits();
    let source = decode(storage, budget)?;
    let table = source
        .list_table
        .ok_or(NativeError::Invalid("storage has no list-style table"))?;
    let existing = references(table, budget)?;
    let mut encoded = Vec::new();
    native::reserve(&mut encoded, boundaries.len())?;
    for &(index, style) in boundaries {
        let matching: Vec<&Reference> = existing
            .iter()
            .filter(|entry| entry.index == index)
            .collect();
        let raw = match matching.as_slice() {
            [] => {
                let mut writer = Writer::new(limits);
                writer.uint32(ENTRY_INDEX_FIELD, index)?;
                writer.message(ENTRY_OBJECT_FIELD, |writer| {
                    writer.uint64(REFERENCE_IDENTIFIER_FIELD, style)
                })?;
                writer.into_bytes()
            },
            [entry] => {
                transform_length_delimited_field(&entry.raw, ENTRY_OBJECT_FIELD, |reference| {
                    patch_varint_field(reference, REFERENCE_IDENTIFIER_FIELD, true, Some(style))
                })
                .map_err(native::common)?
            },
            _ => {
                return Err(NativeError::Invalid(
                    "list-style index occurs several times",
                ));
            },
        };
        encoded.push(raw);
    }
    budget.admit_bytes(storage.len())?;
    transform_length_delimited_field(storage, LIST_STYLE_TABLE_FIELD, |table| {
        rewrite_repeated_length_delimited_fields(table, TABLE_ENTRIES_FIELD, &encoded)
    })
    .map_err(native::common)
}

pub(super) fn validate_levels(source: &Source<'_>, entries: &[Data]) -> NativeResult<()> {
    let Some(first) = entries.first() else {
        return Err(NativeError::Invalid(
            "storage has no paragraph-data boundary",
        ));
    };
    if first.index != 0 {
        return Err(NativeError::Invalid(
            "paragraph data must begin at UTF-16 index zero",
        ));
    }
    let mut previous = None;
    for entry in entries {
        if previous.is_some_and(|index| index >= entry.index) {
            return Err(NativeError::Invalid(
                "paragraph-data boundaries are not strictly increasing",
            ));
        }
        if source.starts.binary_search(&entry.index).is_err() {
            return Err(NativeError::Invalid(
                "paragraph-data boundary is not a paragraph start",
            ));
        }
        if entry.first > u32::from(litchi_iwa_text::paragraph::list::ParagraphListLevel::MAX.get())
        {
            return Err(NativeError::Invalid(
                "native paragraph list level is out of range",
            ));
        }
        previous = Some(entry.index);
    }
    Ok(())
}

fn data_payload(
    index: u32,
    first: u32,
    second: u32,
    limits: litchi_iwa_common::WireLimits,
) -> NativeResult<Vec<u8>> {
    let mut writer = Writer::new(limits);
    writer.uint32(ENTRY_INDEX_FIELD, index)?;
    writer.uint32(ENTRY_FIRST_FIELD, first)?;
    writer.uint32(ENTRY_SECOND_FIELD, second)?;
    Ok(writer.into_bytes())
}

/// Set the level of the paragraph at `target` (the host's level
/// `patch_table`).
pub(super) fn with_level(
    storage: &[u8],
    target: u32,
    level: u32,
    budget: &mut Budget,
) -> NativeResult<Vec<u8>> {
    let limits = budget.limits();
    let source = decode(storage, budget)?;
    let table = source
        .data_table
        .ok_or(NativeError::Invalid("storage has no paragraph-data table"))?;
    let mut entries = data_entries(table, budget)?;
    validate_levels(&source, &entries)?;
    let inherited_second = entries
        .iter()
        .take_while(|entry| entry.index <= target)
        .last()
        .map(|entry| entry.second)
        .ok_or(NativeError::Invalid(
            "storage has no paragraph-data boundary at zero",
        ))?;
    let next = source.starts.iter().copied().find(|start| *start > target);
    let next_effective = next
        .map(|next| {
            entries
                .iter()
                .take_while(|entry| entry.index <= next)
                .last()
                .map(|entry| (entry.first, entry.second))
                .ok_or(NativeError::Invalid(
                    "storage has no paragraph-data boundary at zero",
                ))
        })
        .transpose()?;
    if let Some(entry) = entries.iter_mut().find(|entry| entry.index == target) {
        entry.first = level;
        entry.raw = patch_varint_field(&entry.raw, ENTRY_FIRST_FIELD, true, Some(u64::from(level)))
            .map_err(native::common)?;
    } else {
        native::reserve(&mut entries, 1)?;
        entries.push(Data {
            index: target,
            first: level,
            second: inherited_second,
            other_fields: false,
            raw: data_payload(target, level, inherited_second, limits)?,
        });
        entries.sort_by_key(|entry| entry.index);
    }
    if let Some((next, (first, second))) = next.zip(next_effective)
        && entries.iter().all(|entry| entry.index != next)
    {
        native::reserve(&mut entries, 1)?;
        entries.push(Data {
            index: next,
            first,
            second,
            other_fields: false,
            raw: data_payload(next, first, second, limits)?,
        });
        entries.sort_by_key(|entry| entry.index);
    }
    entries.dedup_by(|current, previous| {
        current.first == previous.first && current.second == previous.second
    });
    let encoded: Vec<Vec<u8>> = entries.into_iter().map(|entry| entry.raw).collect();
    budget.admit_bytes(storage.len())?;
    transform_length_delimited_field(storage, PARAGRAPH_DATA_TABLE_FIELD, |table| {
        rewrite_repeated_length_delimited_fields(table, TABLE_ENTRIES_FIELD, &encoded)
    })
    .map_err(native::common)
}

pub(super) fn validate_numbering(source: &Source<'_>, entries: &[Data]) -> NativeResult<()> {
    let Some(first) = entries.first() else {
        return Err(NativeError::Invalid(
            "storage has an empty paragraph-start table",
        ));
    };
    if first.index != 0 {
        return Err(NativeError::Invalid(
            "paragraph-start data must begin at UTF-16 index zero",
        ));
    }
    let mut previous = None;
    for entry in entries {
        if previous.is_some_and(|index| index >= entry.index) {
            return Err(NativeError::Invalid(
                "paragraph-start boundaries are not strictly increasing",
            ));
        }
        if source.starts.binary_search(&entry.index).is_err() {
            return Err(NativeError::Invalid(
                "paragraph-start boundary is not a paragraph start",
            ));
        }
        previous = Some(entry.index);
    }
    Ok(())
}

/// Set the restart number of the paragraph at `target` (0 continues; the
/// host's numbering `patch_table`/`new_table`).
pub(super) fn with_numbering(
    storage: &[u8],
    target: u32,
    start: u32,
    budget: &mut Budget,
) -> NativeResult<Vec<u8>> {
    let limits = budget.limits();
    let source = decode(storage, budget)?;
    let data = match source.start_table {
        None => {
            let mut writer = Writer::new(limits);
            if target != 0 {
                writer.bytes(TABLE_ENTRIES_FIELD, &data_payload(0, 0, 0, limits)?)?;
            }
            writer.bytes(
                TABLE_ENTRIES_FIELD,
                &data_payload(target, start, 0, limits)?,
            )?;
            budget.admit_bytes(storage.len())?;
            return patch_length_delimited_field(
                storage,
                PARAGRAPH_START_TABLE_FIELD,
                false,
                Some(&writer.into_bytes()),
            )
            .map_err(native::common);
        },
        Some(table) => {
            let entries = data_entries(table, budget)?;
            validate_numbering(&source, &entries)?;
            let mut entries = entries;
            if start == 0 && target != 0 {
                let mut retained = Vec::new();
                for mut entry in entries {
                    if entry.index != target {
                        retained.push(entry);
                    } else if entry.second != 0 || entry.other_fields {
                        entry.first = 0;
                        entry.raw =
                            patch_varint_field(&entry.raw, ENTRY_FIRST_FIELD, true, Some(0))
                                .map_err(native::common)?;
                        retained.push(entry);
                    }
                }
                entries = retained;
            } else if let Some(entry) = entries.iter_mut().find(|entry| entry.index == target) {
                entry.first = start;
                entry.raw =
                    patch_varint_field(&entry.raw, ENTRY_FIRST_FIELD, true, Some(u64::from(start)))
                        .map_err(native::common)?;
            } else {
                native::reserve(&mut entries, 1)?;
                entries.push(Data {
                    index: target,
                    first: start,
                    second: 0,
                    other_fields: false,
                    raw: data_payload(target, start, 0, limits)?,
                });
                entries.sort_by_key(|entry| entry.index);
            }
            if entries.first().is_none_or(|entry| entry.index != 0) {
                entries.insert(
                    0,
                    Data {
                        index: 0,
                        first: 0,
                        second: 0,
                        other_fields: false,
                        raw: data_payload(0, 0, 0, limits)?,
                    },
                );
            }
            let encoded: Vec<Vec<u8>> = entries.into_iter().map(|entry| entry.raw).collect();
            budget.admit_bytes(storage.len())?;
            transform_length_delimited_field(storage, PARAGRAPH_START_TABLE_FIELD, |table| {
                rewrite_repeated_length_delimited_fields(table, TABLE_ENTRIES_FIELD, &encoded)
            })
            .map_err(native::common)?
        },
    };
    Ok(data)
}
