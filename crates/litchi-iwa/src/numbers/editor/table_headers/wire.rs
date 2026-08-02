//! Lossless protobuf wire handling for Numbers table header settings.

use super::*;

const HEADER_ROWS_FIELD: u32 = 9;
const HEADER_COLUMNS_FIELD: u32 = 10;
const FOOTER_ROWS_FIELD: u32 = 11;
const HEADER_ROWS_FROZEN_FIELD: u32 = 12;
const HEADER_COLUMNS_FROZEN_FIELD: u32 = 13;
const REPEATING_HEADER_ROWS_FIELD: u32 = 29;
const REPEATING_HEADER_COLUMNS_FIELD: u32 = 32;

pub(super) fn read_table_header_settings_wire(
    original: &[u8],
    model: &TableModelArchive,
) -> Result<NumbersTableHeaderSettings> {
    for (field, label, decoded) in [
        (
            HEADER_ROWS_FIELD,
            "header row count",
            model.number_of_header_rows.map(u64::from),
        ),
        (
            HEADER_COLUMNS_FIELD,
            "header column count",
            model.number_of_header_columns.map(u64::from),
        ),
        (
            FOOTER_ROWS_FIELD,
            "footer row count",
            model.number_of_footer_rows.map(u64::from),
        ),
    ] {
        require_optional_varint(original, field, label, decoded)?;
    }
    for (field, label, decoded) in [
        (
            HEADER_ROWS_FROZEN_FIELD,
            "header rows frozen",
            model.header_rows_frozen,
        ),
        (
            HEADER_COLUMNS_FROZEN_FIELD,
            "header columns frozen",
            model.header_columns_frozen,
        ),
        (
            REPEATING_HEADER_ROWS_FIELD,
            "repeating header rows",
            model.repeating_header_rows_enabled,
        ),
        (
            REPEATING_HEADER_COLUMNS_FIELD,
            "repeating header columns",
            model.repeating_header_columns_enabled,
        ),
    ] {
        require_optional_bool(original, field, label, decoded)?;
    }
    NumbersTableHeaderSettings::from_model(model)
}

pub(super) fn write_table_header_settings_wire(
    original: &[u8],
    model: &TableModelArchive,
    settings: NumbersTableHeaderSettings,
) -> Result<Vec<u8>> {
    read_table_header_settings_wire(original, model)?;
    let mut data = original.to_vec();
    for (field, present, replacement) in [
        (
            HEADER_ROWS_FIELD,
            model.number_of_header_rows.is_some(),
            settings
                .header_rows
                .map(NumbersTableHeaderCount::as_native)
                .map(u64::from),
        ),
        (
            HEADER_COLUMNS_FIELD,
            model.number_of_header_columns.is_some(),
            settings
                .header_columns
                .map(NumbersTableHeaderCount::as_native)
                .map(u64::from),
        ),
        (
            FOOTER_ROWS_FIELD,
            model.number_of_footer_rows.is_some(),
            settings
                .footer_rows
                .map(NumbersTableHeaderCount::as_native)
                .map(u64::from),
        ),
        (
            HEADER_ROWS_FROZEN_FIELD,
            model.header_rows_frozen.is_some(),
            settings.header_rows_frozen.map(u64::from),
        ),
        (
            HEADER_COLUMNS_FROZEN_FIELD,
            model.header_columns_frozen.is_some(),
            settings.header_columns_frozen.map(u64::from),
        ),
        (
            REPEATING_HEADER_ROWS_FIELD,
            model.repeating_header_rows_enabled.is_some(),
            settings.repeating_header_rows_enabled.map(u64::from),
        ),
        (
            REPEATING_HEADER_COLUMNS_FIELD,
            model.repeating_header_columns_enabled.is_some(),
            settings.repeating_header_columns_enabled.map(u64::from),
        ),
    ] {
        data = patch_varint_field(&data, field, present, replacement)?;
    }
    let verified = TableModelArchive::decode(data.as_slice())?;
    if read_table_header_settings_wire(&data, &verified)? != settings {
        return Err(Error::InvalidFormat(
            "Numbers table header wire patch failed validation".to_owned(),
        ));
    }
    Ok(data)
}

fn require_optional_varint(
    original: &[u8],
    field: u32,
    label: &str,
    decoded: Option<u64>,
) -> Result<()> {
    let values = repeated_varint_values(original, field)?;
    if values.as_slice() != decoded.as_slice() {
        return Err(Error::InvalidFormat(format!(
            "Numbers table {label} wire value is missing, duplicated, or inconsistent"
        )));
    }
    Ok(())
}

fn require_optional_bool(
    original: &[u8],
    field: u32,
    label: &str,
    decoded: Option<bool>,
) -> Result<()> {
    let values = repeated_varint_values(original, field)?;
    let expected = decoded.map(u64::from);
    if values.as_slice() != expected.as_slice() {
        return Err(Error::InvalidFormat(format!(
            "Numbers table {label} wire value is missing, duplicated, non-Boolean, or inconsistent"
        )));
    }
    Ok(())
}
