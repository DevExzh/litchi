//! Lossless protobuf wire handling for Numbers table sort orders.

use super::*;

use crate::wire::{
    patch_length_delimited_field, patch_varint_field, repeated_length_delimited_payloads,
    repeated_varint_values, rewrite_repeated_length_delimited_fields,
    transform_length_delimited_field,
};

const SORT_ORDER_FIELD: u32 = 44;
const SORT_TYPE_FIELD: u32 = 1;
const SORT_RULES_FIELD: u32 = 2;
const SORT_RULE_COLUMN_FIELD: u32 = 1;
const SORT_RULE_DIRECTION_FIELD: u32 = 2;

pub(super) fn read_native_table_sort_order_wire(
    original: &[u8],
    model: &TableModelArchive,
) -> Result<Option<tst::TableSortOrderArchive>> {
    let payloads = repeated_length_delimited_payloads(original, SORT_ORDER_FIELD)?;
    let native = match payloads.as_slice() {
        [] => None,
        [payload] => {
            let native = tst::TableSortOrderArchive::decode(*payload)?;
            validate_sort_order_wire_payload(payload, &native)?;
            Some(native)
        },
        _ => {
            return Err(Error::InvalidFormat(
                "Numbers table sort-order wire field is duplicated".to_owned(),
            ));
        },
    };
    if native.as_ref() != model.sort_order.as_ref() {
        return Err(Error::InvalidFormat(
            "Numbers table sort-order wire payload is missing or inconsistent".to_owned(),
        ));
    }
    Ok(native)
}

fn validate_sort_order_wire_payload(
    payload: &[u8],
    native: &tst::TableSortOrderArchive,
) -> Result<()> {
    validate_required_varint(payload, SORT_TYPE_FIELD, native.r#type as u64, "type")?;
    let raw_rules = repeated_length_delimited_payloads(payload, SORT_RULES_FIELD)?;
    if raw_rules.len() != native.rules.len() {
        return Err(Error::InvalidFormat(
            "Numbers table sort order has an inconsistent rule wire payload".to_owned(),
        ));
    }
    for (raw, rule) in raw_rules.into_iter().zip(&native.rules) {
        validate_required_varint(
            raw,
            SORT_RULE_COLUMN_FIELD,
            u64::from(rule.index),
            "column index",
        )?;
        validate_required_varint(
            raw,
            SORT_RULE_DIRECTION_FIELD,
            rule.direction as u64,
            "direction",
        )?;
    }
    Ok(())
}

fn validate_required_varint(
    data: &[u8],
    field_number: u32,
    expected: u64,
    name: &str,
) -> Result<()> {
    let values = repeated_varint_values(data, field_number)?;
    match values.as_slice() {
        [value] if *value == expected => Ok(()),
        [] => Err(Error::InvalidFormat(format!(
            "Numbers table sort {name} field is missing"
        ))),
        [_] => Err(Error::InvalidFormat(format!(
            "Numbers table sort {name} field is inconsistent"
        ))),
        _ => Err(Error::InvalidFormat(format!(
            "Numbers table sort {name} field is duplicated"
        ))),
    }
}

pub(super) fn write_table_sort_order_wire(
    original: &[u8],
    model: &TableModelArchive,
    order: &NumbersTableSortOrder,
) -> Result<Vec<u8>> {
    let existing = read_native_table_sort_order_wire(original, model)?;
    let expected = order.as_native()?;
    let data = if existing.is_some() {
        transform_length_delimited_field(original, SORT_ORDER_FIELD, |sort_order| {
            rewrite_sort_order_wire(sort_order, &expected)
        })?
    } else {
        let replacement = expected.encode_to_vec();
        patch_length_delimited_field(original, SORT_ORDER_FIELD, false, Some(&replacement))?
    };
    let verified = TableModelArchive::decode(data.as_slice())?;
    if read_native_table_sort_order_wire(&data, &verified)?.as_ref() != Some(&expected) {
        return Err(Error::InvalidFormat(
            "Numbers table sort-order wire patch failed validation".to_owned(),
        ));
    }
    Ok(data)
}

pub(super) fn delete_table_sort_column_wire(
    original: &[u8],
    model: &TableModelArchive,
    column: u32,
    new_columns: u32,
) -> Result<Vec<u8>> {
    let Some(previous) = read_native_table_sort_order_wire(original, model)? else {
        return Ok(original.to_vec());
    };
    let mut expected = previous.clone();
    expected
        .rules
        .retain(|rule| rule.index != column && rule.index < new_columns);
    if expected == previous {
        return Ok(original.to_vec());
    }
    let data = transform_length_delimited_field(original, SORT_ORDER_FIELD, |sort_order| {
        delete_sort_column_wire(sort_order, &previous, &expected, column, new_columns)
    })?;
    let verified = TableModelArchive::decode(data.as_slice())?;
    if read_native_table_sort_order_wire(&data, &verified)?.as_ref() != Some(&expected) {
        return Err(Error::InvalidFormat(
            "Numbers table sort-order column deletion failed validation".to_owned(),
        ));
    }
    Ok(data)
}

fn delete_sort_column_wire(
    original: &[u8],
    previous: &tst::TableSortOrderArchive,
    expected: &tst::TableSortOrderArchive,
    column: u32,
    new_columns: u32,
) -> Result<Vec<u8>> {
    let raw_rules = repeated_length_delimited_payloads(original, SORT_RULES_FIELD)?;
    if raw_rules.len() != previous.rules.len() {
        return Err(Error::InvalidFormat(
            "Numbers table sort order has an inconsistent rule wire payload".to_owned(),
        ));
    }
    let retained = raw_rules
        .into_iter()
        .zip(&previous.rules)
        .filter(|(_, rule)| rule.index != column && rule.index < new_columns)
        .map(|(raw, _)| raw.to_vec())
        .collect::<Vec<_>>();
    let data = rewrite_repeated_length_delimited_fields(original, SORT_RULES_FIELD, &retained)?;
    if tst::TableSortOrderArchive::decode(data.as_slice())? != *expected {
        return Err(Error::InvalidFormat(
            "Numbers table sort-rule deletion failed validation".to_owned(),
        ));
    }
    Ok(data)
}

fn rewrite_sort_order_wire(
    original: &[u8],
    expected: &tst::TableSortOrderArchive,
) -> Result<Vec<u8>> {
    let previous = tst::TableSortOrderArchive::decode(original)?;
    let raw_rules = repeated_length_delimited_payloads(original, SORT_RULES_FIELD)?;
    if raw_rules.len() != previous.rules.len() {
        return Err(Error::InvalidFormat(
            "Numbers table sort order has an inconsistent rule wire payload".to_owned(),
        ));
    }

    let mut data = patch_varint_field(
        original,
        SORT_TYPE_FIELD,
        true,
        Some(expected.r#type as u64),
    )?;
    let rules = expected
        .rules
        .iter()
        .enumerate()
        .map(|(index, rule)| {
            raw_rules.get(index).map_or_else(
                || Ok(rule.encode_to_vec()),
                |raw| rewrite_sort_rule_wire(raw, rule),
            )
        })
        .collect::<Result<Vec<_>>>()?;
    data = rewrite_repeated_length_delimited_fields(&data, SORT_RULES_FIELD, &rules)?;

    if tst::TableSortOrderArchive::decode(data.as_slice())? != *expected {
        return Err(Error::InvalidFormat(
            "Numbers table sort-order mutation failed validation".to_owned(),
        ));
    }
    Ok(data)
}

fn rewrite_sort_rule_wire(
    original: &[u8],
    expected: &tst::table_sort_order_archive::SortRuleArchive,
) -> Result<Vec<u8>> {
    let mut data = patch_varint_field(
        original,
        SORT_RULE_COLUMN_FIELD,
        true,
        Some(u64::from(expected.index)),
    )?;
    data = patch_varint_field(
        &data,
        SORT_RULE_DIRECTION_FIELD,
        true,
        Some(expected.direction as u64),
    )?;
    if tst::table_sort_order_archive::SortRuleArchive::decode(data.as_slice())? != *expected {
        return Err(Error::InvalidFormat(
            "Numbers table sort-rule mutation failed validation".to_owned(),
        ));
    }
    Ok(data)
}

pub(super) fn clear_table_sort_order_wire(
    original: &[u8],
    model: &TableModelArchive,
) -> Result<Vec<u8>> {
    let native = read_native_table_sort_order_wire(original, model)?.ok_or_else(|| {
        Error::InvalidFormat("Numbers table has no native sort order to clear".to_owned())
    })?;
    if native.rules.is_empty() {
        return Ok(original.to_vec());
    }
    let data = transform_length_delimited_field(original, SORT_ORDER_FIELD, |sort_order| {
        rewrite_repeated_length_delimited_fields(sort_order, SORT_RULES_FIELD, &[])
    })?;
    let verified = TableModelArchive::decode(data.as_slice())?;
    let Some(cleared) = read_native_table_sort_order_wire(&data, &verified)? else {
        return Err(Error::InvalidFormat(
            "Numbers table sort-order clear removed its native marker".to_owned(),
        ));
    };
    if !cleared.rules.is_empty() || cleared.r#type != native.r#type {
        return Err(Error::InvalidFormat(
            "Numbers table sort-order clear failed wire validation".to_owned(),
        ));
    }
    Ok(data)
}
