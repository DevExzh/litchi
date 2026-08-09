//! Cross-field validation for `PivotTable` OLAP values.

use crate::error::{Error, Result};

use super::model::{PivotFieldOlapExt, PivotHierarchy, PivotPageItemOlapExt, PivotViewOlapHeader};
use super::{
    MAX_FUTURE_BYTES, MAX_OLAP_STRING_CHARS, SX_VIEW_EX_RECORD_TYPE, SXPI_EX_RECORD_TYPE,
    SXTH_RECORD_TYPE, SXVDT_EX_RECORD_TYPE,
};

#[derive(Clone, Copy)]
enum ValidationMode {
    Parse,
    Write,
}

fn failure(mode: ValidationMode, record_type: u16, message: impl Into<String>) -> Error {
    match mode {
        ValidationMode::Parse => Error::InvalidRecord {
            record_type,
            message: message.into(),
        },
        ValidationMode::Write => Error::InvalidData(message.into()),
    }
}

fn require(
    mode: ValidationMode,
    condition: bool,
    record_type: u16,
    message: impl Into<String>,
) -> Result<()> {
    if condition {
        Ok(())
    } else {
        Err(failure(mode, record_type, message))
    }
}

fn string_len(value: &str) -> usize {
    // XLUnicodeString.cch counts UTF-16 code units, not Unicode scalar
    // values. This also keeps supplementary characters within the wire limit.
    value.encode_utf16().count()
}

fn validate_string(
    mode: ValidationMode,
    record_type: u16,
    value: &str,
    field: &str,
    allow_empty: bool,
) -> Result<()> {
    require(
        mode,
        allow_empty || !value.is_empty(),
        record_type,
        format!("{field} must not be empty"),
    )?;
    require(
        mode,
        string_len(value) <= MAX_OLAP_STRING_CHARS,
        record_type,
        format!("{field} exceeds {MAX_OLAP_STRING_CHARS} UTF-16 characters"),
    )
}

pub(super) fn validate_view_header(value: &PivotViewOlapHeader, write: bool) -> Result<()> {
    let mode = if write {
        ValidationMode::Write
    } else {
        ValidationMode::Parse
    };
    require(
        mode,
        value.hierarchy_count >= 1,
        SX_VIEW_EX_RECORD_TYPE,
        "SXViewEx csxth must be at least 1",
    )?;
    require(
        mode,
        value.future_bytes.len() <= MAX_FUTURE_BYTES,
        SX_VIEW_EX_RECORD_TYPE,
        format!("SXViewEx cbFuture exceeds {MAX_FUTURE_BYTES}"),
    )
}

pub(super) fn validate_hierarchy(value: &PivotHierarchy, write: bool) -> Result<()> {
    let mode = if write {
        ValidationMode::Write
    } else {
        ValidationMode::Parse
    };
    require(
        mode,
        !(value.is_measure
            && (value.is_named_set
                || value.drag_to_row
                || value.drag_to_column
                || value.drag_to_page)),
        SXTH_RECORD_TYPE,
        "SXTH measure cannot be a named set or drag to row/column/page",
    )?;
    require(
        mode,
        !(value.is_measure && !value.dimension.is_empty()),
        SXTH_RECORD_TYPE,
        "SXTH measure must have an empty stDimension",
    )?;
    require(
        mode,
        value.axis_field_count >= 0,
        SXTH_RECORD_TYPE,
        "SXTH csxvdXl must be non-negative",
    )?;
    require(
        mode,
        value.level_fields.is_empty() || value.axis.row || value.axis.column,
        SXTH_RECORD_TYPE,
        "SXTH cisxvd must be zero off the row/column axes",
    )?;

    let expected_axis_fields = if value.all_member.is_empty() {
        i64::try_from(value.level_fields.len()).unwrap_or(i64::MAX)
    } else {
        i64::try_from(value.level_fields.len())
            .unwrap_or(i64::MAX)
            .saturating_sub(1)
    };
    require(
        mode,
        i64::from(value.axis_field_count) == expected_axis_fields,
        SXTH_RECORD_TYPE,
        "SXTH csxvdXl does not match cisxvd and stAll",
    )?;
    require(
        mode,
        !(value.filter_inclusive && !value.hidden_member_sets.is_empty()),
        SXTH_RECORD_TYPE,
        "SXTH cHiddenMemberSets must be zero for inclusive filters",
    )?;
    require(
        mode,
        value.hidden_member_sets.is_empty() || !value.level_fields.is_empty(),
        SXTH_RECORD_TYPE,
        "SXTH rgHiddenMemberSets requires non-empty cisxvd",
    )?;
    // MS-XLS 2.4.308 defines cHiddenMemberSets as the deepest one-based
    // level. Therefore it cannot exceed the number of hierarchy levels.
    require(
        mode,
        value.hidden_member_sets.len() <= value.level_fields.len(),
        SXTH_RECORD_TYPE,
        "SXTH cHiddenMemberSets exceeds cisxvd",
    )?;
    validate_string(
        mode,
        SXTH_RECORD_TYPE,
        &value.unique_name,
        "SXTH stUnique",
        false,
    )?;
    validate_string(
        mode,
        SXTH_RECORD_TYPE,
        &value.display_name,
        "SXTH stDisplay",
        false,
    )?;
    validate_string(
        mode,
        SXTH_RECORD_TYPE,
        &value.default_member,
        "SXTH stDefault",
        true,
    )?;
    validate_string(
        mode,
        SXTH_RECORD_TYPE,
        &value.all_member,
        "SXTH stAll",
        true,
    )?;
    validate_string(
        mode,
        SXTH_RECORD_TYPE,
        &value.dimension,
        "SXTH stDimension",
        true,
    )?;
    for &field in &value.level_fields {
        require(
            mode,
            field >= -1,
            SXTH_RECORD_TYPE,
            "SXTH rgisxvd element must be -1 or a pivot field index",
        )?;
    }
    for set in &value.hidden_member_sets {
        for name in &set.member_names {
            validate_string(
                mode,
                SXTH_RECORD_TYPE,
                name,
                "SXTH hidden member name",
                true,
            )?;
        }
    }
    Ok(())
}

pub(super) fn validate_page_extension(value: &PivotPageItemOlapExt, write: bool) -> Result<()> {
    let mode = if write {
        ValidationMode::Write
    } else {
        ValidationMode::Parse
    };
    validate_string(
        mode,
        SXPI_EX_RECORD_TYPE,
        &value.unique_name,
        "SXPIEx stUnique",
        true,
    )?;
    validate_string(
        mode,
        SXPI_EX_RECORD_TYPE,
        &value.display_name,
        "SXPIEx stDisplay",
        true,
    )
}

pub(super) fn validate_field_extension(value: &PivotFieldOlapExt, write: bool) -> Result<()> {
    let mode = if write {
        ValidationMode::Write
    } else {
        ValidationMode::Parse
    };
    require(
        mode,
        value.hierarchy_index >= -1,
        SXVDT_EX_RECORD_TYPE,
        "SXVDTEx isxth must be -1 or a pivot hierarchy index",
    )?;
    require(
        mode,
        i32::try_from(value.item_flags.len()).is_ok(),
        SXVDT_EX_RECORD_TYPE,
        "SXVDTEx csxvi exceeds the signed 32-bit field",
    )
}
