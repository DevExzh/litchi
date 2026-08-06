//! Host-bound formula metadata validation.

use crate::named_ranges::validate_name;
use crate::package::error::{Error, Result};

use super::TableNamedColumns;

pub(super) fn validate_pivot_identifier(
    name: &str,
    field: &str,
    max_utf16_len: usize,
) -> Result<()> {
    let utf16_len = name.encode_utf16().count();
    if utf16_len == 0 || utf16_len > max_utf16_len || name.contains('\0') {
        return Err(invalid(
            "PtgSxName",
            format!("{field} must contain 1..={max_utf16_len} UTF-16 code units and no NUL"),
        ));
    }
    Ok(())
}

pub(super) fn invalid(typ: &'static str, value: impl Into<String>) -> Error {
    Error::InvalidFormula(format!("{typ}: {}", value.into()))
}

pub(super) fn validate_table_name(name: &str) -> Result<()> {
    validate_name(name)?;
    if name
        .get(..3)
        .is_some_and(|prefix| prefix.eq_ignore_ascii_case("_xl"))
    {
        return Err(Error::InvalidFormula(format!(
            "table display name {name:?} uses reserved _xl prefix"
        )));
    }
    Ok(())
}

pub(super) fn validate_table_column_name(name: &str, index: usize) -> Result<()> {
    let units = name.encode_utf16().count();
    if units == 0 || units > 255 || name.contains('\0') {
        return Err(Error::InvalidFormula(format!(
            "table column {index} has invalid name length or NUL content"
        )));
    }
    Ok(())
}

pub(super) fn validate_named_table_columns(columns: &TableNamedColumns) -> Result<()> {
    match columns {
        TableNamedColumns::All => Ok(()),
        TableNamedColumns::One(name) => validate_table_column_name(name, 0),
        TableNamedColumns::Range { first, last } => {
            validate_table_column_name(first, 0)?;
            validate_table_column_name(last, 1)
        },
    }
}
