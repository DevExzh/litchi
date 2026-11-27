//! Uniform paragraph-style references of a `TSWP.StorageArchive` and the
//! variation registry of a `TSS.StylesheetArchive`.
//!
//! The rewrites go through the shared `litchi-iwa-common` wire helpers the
//! host uses, so every byte the operation does not change is kept and the
//! output equals the host's.

use litchi_iwa_common::wire::{
    append_repeated_length_delimited_field, patch_varint_field, repeated_length_delimited_payloads,
    rewrite_repeated_length_delimited_fields, transform_length_delimited_field,
};

use crate::native::{self, Budget, NativeError, NativeResult, Writer};

const PARAGRAPH_STYLE_TABLE_FIELD: u32 = 5;
const TABLE_ENTRIES_FIELD: u32 = 1;
const ENTRY_CHARACTER_INDEX_FIELD: u32 = 1;
const ENTRY_OBJECT_FIELD: u32 = 2;
const REFERENCE_IDENTIFIER_FIELD: u32 = 1;
const STYLESHEET_STYLES_FIELD: u32 = 1;
const STYLESHEET_PARENT_CHILDREN_FIELD: u32 = 5;
const STYLE_CHILDREN_PARENT_FIELD: u32 = 1;
const STYLE_CHILDREN_FIELD: u32 = 2;

fn invalid<T>(reason: &'static str) -> NativeResult<T> {
    Err(NativeError::Invalid(reason))
}

/// The table's single entry, as the host requires of a uniformly styled
/// storage: one boundary at UTF-16 index zero naming a nonzero style.
fn uniform_entry<'a>(table: &'a [u8], budget: &mut Budget) -> NativeResult<(&'a [u8], u64)> {
    let mut entries = Vec::new();
    native::fields(table, 1, budget, |field, _| {
        if field.number == TABLE_ENTRIES_FIELD {
            native::reserve(&mut entries, 1)?;
            entries.push(field.bytes()?);
        }
        Ok(())
    })?;
    let [entry] = entries.as_slice() else {
        return invalid("iWork text storage must have one uniform paragraph-style boundary");
    };
    let mut index = 0;
    let mut reference = None;
    native::fields(entry, 2, budget, |field, budget| {
        match field.number {
            ENTRY_CHARACTER_INDEX_FIELD => index = field.uint32()?,
            ENTRY_OBJECT_FIELD => reference
                .get_or_insert_with(native::Reference::default)
                .merge(field.bytes()?, 3, budget)?,
            _ => {},
        }
        Ok(())
    })?;
    if index != 0 {
        return invalid("iWork text storage paragraph style must begin at UTF-16 index zero");
    }
    match reference {
        Some(reference) if reference.identifier != 0 => Ok((entry, reference.identifier)),
        _ => invalid("iWork text storage has no uniform paragraph style"),
    }
}

fn paragraph_table<'a>(storage: &'a [u8], budget: &mut Budget) -> NativeResult<&'a [u8]> {
    let mut tables = Vec::new();
    native::fields(storage, 0, budget, |field, _| {
        if field.number == PARAGRAPH_STYLE_TABLE_FIELD {
            native::reserve(&mut tables, 1)?;
            tables.push(field.bytes()?);
        }
        Ok(())
    })?;
    match tables.as_slice() {
        [table] => Ok(table),
        [] => invalid("iWork text storage has no uniform paragraph style"),
        _ => invalid("iWork text storage paragraph-style table wire state is inconsistent"),
    }
}

pub(super) fn uniform_style(storage: &[u8], budget: &mut Budget) -> NativeResult<u64> {
    let table = paragraph_table(storage, budget)?;
    uniform_entry(table, budget).map(|(_, identifier)| identifier)
}

pub(super) fn repoint(
    storage: &[u8],
    from: u64,
    to: u64,
    budget: &mut Budget,
) -> NativeResult<Vec<u8>> {
    let table = paragraph_table(storage, budget)?;
    let (_, current) = uniform_entry(table, budget)?;
    if current != from {
        return invalid("iWork text storage paragraph style changed unexpectedly");
    }
    budget.admit_bytes(storage.len())?;
    let rewritten =
        transform_length_delimited_field(storage, PARAGRAPH_STYLE_TABLE_FIELD, |table| {
            let entries = repeated_length_delimited_payloads(table, TABLE_ENTRIES_FIELD)?;
            let [entry] = entries.as_slice() else {
                return Err(litchi_iwa_common::Error::InvalidFormat(
                    "iWork text storage must have one uniform paragraph-style boundary".to_owned(),
                ));
            };
            let patched =
                transform_length_delimited_field(entry, ENTRY_OBJECT_FIELD, |reference| {
                    patch_varint_field(reference, REFERENCE_IDENTIFIER_FIELD, true, Some(to))
                })?;
            rewrite_repeated_length_delimited_fields(table, TABLE_ENTRIES_FIELD, &[patched])
        })
        .map_err(native::common)?;
    admit_output(&rewritten, budget)?;
    Ok(rewritten)
}

fn admit_output(output: &[u8], budget: &Budget) -> NativeResult<()> {
    native::admit(
        litchi_iwa_common::LimitKind::OutputBytes,
        output.len(),
        budget.limits().max_output_bytes(),
    )
}

fn reference_payload(identifier: u64, budget: &Budget) -> NativeResult<Vec<u8>> {
    let mut writer = Writer::new(budget.limits());
    native::Reference { identifier }.encode(&mut writer)?;
    Ok(writer.into_bytes())
}

fn reference_identifier(payload: &[u8], budget: &mut Budget) -> NativeResult<u64> {
    let mut reference = native::Reference::default();
    reference.merge(payload, 1, budget)?;
    Ok(reference.identifier)
}

/// The styles (field 1) and parent-children entries (field 5) of a
/// stylesheet: `(styles, [(parent, children)])`.
fn registry(sheet: &[u8], budget: &mut Budget) -> NativeResult<(Vec<u64>, Vec<(u64, Vec<u64>)>)> {
    let mut styles = Vec::new();
    let mut entries = Vec::new();
    native::fields(sheet, 0, budget, |field, budget| {
        match field.number {
            STYLESHEET_STYLES_FIELD => {
                let identifier = reference_identifier(field.bytes()?, budget)?;
                native::reserve(&mut styles, 1)?;
                styles.push(identifier);
            },
            STYLESHEET_PARENT_CHILDREN_FIELD => {
                let mut parent = native::Reference::default();
                let mut children = Vec::new();
                native::fields(field.bytes()?, 1, budget, |inner, budget| {
                    match inner.number {
                        STYLE_CHILDREN_PARENT_FIELD => parent.merge(inner.bytes()?, 2, budget)?,
                        STYLE_CHILDREN_FIELD => {
                            let mut child = native::Reference::default();
                            child.merge(inner.bytes()?, 2, budget)?;
                            native::reserve(&mut children, 1)?;
                            children.push(child.identifier);
                        },
                        _ => {},
                    }
                    Ok(())
                })?;
                native::reserve(&mut entries, 1)?;
                entries.push((parent.identifier, children));
            },
            _ => {},
        }
        Ok(())
    })?;
    Ok((styles, entries))
}

pub(super) fn stylesheet_with_variation(
    sheet: &[u8],
    parent: u64,
    style: u64,
    budget: &mut Budget,
) -> NativeResult<Vec<u8>> {
    let (styles, entries) = registry(sheet, budget)?;
    if styles.contains(&style)
        || entries
            .iter()
            .any(|(_, children)| children.contains(&style))
    {
        return invalid("iWork stylesheet already contains the style");
    }
    budget.admit_bytes(sheet.len())?;
    let reference = reference_payload(style, budget)?;
    let data = append_repeated_length_delimited_field(sheet, STYLESHEET_STYLES_FIELD, &reference)
        .map_err(native::common)?;
    let payloads = repeated_length_delimited_payloads(&data, STYLESHEET_PARENT_CHILDREN_FIELD)
        .map_err(native::common)?;
    let mut parent_matches = 0usize;
    let mut rewritten = Vec::new();
    native::reserve(&mut rewritten, payloads.len() + 1)?;
    for (payload, (entry_parent, _)) in payloads.iter().zip(&entries) {
        if *entry_parent == parent {
            parent_matches += 1;
            rewritten.push(
                append_repeated_length_delimited_field(payload, STYLE_CHILDREN_FIELD, &reference)
                    .map_err(native::common)?,
            );
        } else {
            rewritten.push(payload.to_vec());
        }
    }
    if parent_matches > 1 {
        return invalid("iWork stylesheet repeats the parent style");
    }
    if parent_matches == 0 {
        let mut writer = Writer::new(budget.limits());
        writer.message(STYLE_CHILDREN_PARENT_FIELD, |writer| {
            native::Reference { identifier: parent }.encode(writer)
        })?;
        writer.bytes(STYLE_CHILDREN_FIELD, &reference)?;
        rewritten.push(writer.into_bytes());
    }
    let data = rewrite_repeated_length_delimited_fields(
        &data,
        STYLESHEET_PARENT_CHILDREN_FIELD,
        &rewritten,
    )
    .map_err(native::common)?;
    admit_output(&data, budget)?;
    let (styles, entries) = registry(&data, budget)?;
    let listed = styles
        .iter()
        .filter(|&&identifier| identifier == style)
        .count();
    let mapped = entries
        .iter()
        .filter(|(entry_parent, _)| *entry_parent == parent)
        .flat_map(|(_, children)| children)
        .filter(|&&child| child == style)
        .count();
    if listed != 1 || mapped != 1 {
        return invalid("iWork stylesheet insertion failed validation");
    }
    Ok(data)
}

pub(super) fn stylesheet_without_variation(
    sheet: &[u8],
    parent: u64,
    style: u64,
    budget: &mut Budget,
) -> NativeResult<Vec<u8>> {
    budget.admit_bytes(sheet.len())?;
    let mut styles = Vec::new();
    for payload in repeated_length_delimited_payloads(sheet, STYLESHEET_STYLES_FIELD)
        .map_err(native::common)?
    {
        let identifier = reference_identifier(payload, budget)?;
        native::reserve(&mut styles, 1)?;
        styles.push((identifier, payload.to_vec()));
    }
    if styles
        .iter()
        .filter(|(identifier, _)| *identifier == style)
        .count()
        != 1
    {
        return invalid("iWork stylesheet must contain the disposable style exactly once");
    }
    styles.retain(|(identifier, _)| *identifier != style);
    let styles = styles
        .into_iter()
        .map(|(_, payload)| payload)
        .collect::<Vec<_>>();
    let data = rewrite_repeated_length_delimited_fields(sheet, STYLESHEET_STYLES_FIELD, &styles)
        .map_err(native::common)?;
    let mut removed = 0usize;
    let mut entries = Vec::new();
    for payload in repeated_length_delimited_payloads(&data, STYLESHEET_PARENT_CHILDREN_FIELD)
        .map_err(native::common)?
    {
        let mut entry_parent = None;
        native::fields(payload, 1, budget, |field, budget| {
            if field.number == STYLE_CHILDREN_PARENT_FIELD {
                entry_parent = Some(reference_identifier(field.bytes()?, budget)?);
            }
            Ok(())
        })?;
        let mut children = Vec::new();
        for child in repeated_length_delimited_payloads(payload, STYLE_CHILDREN_FIELD)
            .map_err(native::common)?
        {
            let identifier = reference_identifier(child, budget)?;
            native::reserve(&mut children, 1)?;
            children.push((identifier, child.to_vec()));
        }
        if children.iter().any(|(identifier, _)| *identifier == style)
            && entry_parent.unwrap_or_default() != parent
        {
            return invalid("iWork stylesheet maps the disposable style under the wrong parent");
        }
        let before = children.len();
        children.retain(|(identifier, _)| *identifier != style);
        removed += before - children.len();
        if !children.is_empty() {
            let children = children
                .into_iter()
                .map(|(_, child)| child)
                .collect::<Vec<_>>();
            native::reserve(&mut entries, 1)?;
            entries.push(
                rewrite_repeated_length_delimited_fields(payload, STYLE_CHILDREN_FIELD, &children)
                    .map_err(native::common)?,
            );
        }
    }
    if removed != 1 {
        return invalid(
            "iWork stylesheet child map references the disposable style more than once",
        );
    }
    let data =
        rewrite_repeated_length_delimited_fields(&data, STYLESHEET_PARENT_CHILDREN_FIELD, &entries)
            .map_err(native::common)?;
    admit_output(&data, budget)?;
    let (styles, entries) = registry(&data, budget)?;
    if styles.contains(&style)
        || entries
            .iter()
            .any(|(_, children)| children.contains(&style))
    {
        return invalid("iWork stylesheet retained a removed style");
    }
    Ok(data)
}
