//! Structural and resource validation for Scenario Manager values.

use super::model::{
    CellRange, ChangedCell, MAX_CHANGED_CELLS, MAX_RESULT_RANGES, MAX_SCENARIO_TEXT, MAX_SCENARIOS,
    MAX_UNKNOWN_PAYLOAD, MAX_UNKNOWN_RECORDS, MAX_USER_NAME, Manager, Scenario,
};
use crate::package::error::{Error, Result};
use std::collections::HashSet;

pub(crate) fn invalid(typ: impl Into<String>, val: impl Into<String>) -> Error {
    Error::Unrecognized {
        typ: typ.into(),
        val: val.into(),
    }
}

pub(crate) fn validate_range(value: CellRange) -> Result<()> {
    if value.row_first() > value.row_last()
        || value.row_last() >= 1_048_576
        || value.column_first() > value.column_last()
        || value.column_last() >= 16_384
    {
        return Err(invalid(
            "scenario range",
            format!(
                "rows {}..={}, columns {}..={}",
                value.row_first(),
                value.row_last(),
                value.column_first(),
                value.column_last()
            ),
        ));
    }
    Ok(())
}

fn utf16_units(value: &str) -> usize {
    value.encode_utf16().count()
}

pub(crate) fn validate_text(value: &str, field: &str, maximum: usize) -> Result<()> {
    let units = utf16_units(value);
    if units > maximum {
        return Err(invalid(
            field,
            format!("{units} UTF-16 code units exceeds {maximum}"),
        ));
    }
    if value.chars().any(|character| character == '\0') {
        return Err(invalid(field, "contains NUL"));
    }
    Ok(())
}

pub(crate) fn validate_changed_cell(value: &ChangedCell) -> Result<()> {
    if value.row() >= 1_048_576 || value.column() >= 16_384 {
        return Err(invalid(
            "BrtSlc cell",
            format!(
                "({}, {}) is outside the worksheet grid",
                value.row(),
                value.column()
            ),
        ));
    }
    validate_text(value.value(), "BrtSlc strVal", MAX_SCENARIO_TEXT)
}

pub(crate) fn validate_scenario_text(value: &Scenario) -> Result<()> {
    validate_text(value.name(), "BrtBeginSct Name", MAX_SCENARIO_TEXT)?;
    validate_text(value.comment(), "BrtBeginSct Comment", MAX_SCENARIO_TEXT)?;
    let user_units = utf16_units(value.user_name());
    if !(2..=MAX_USER_NAME).contains(&user_units) {
        return Err(invalid(
            "BrtBeginSct UserName",
            format!("{user_units} UTF-16 code units is outside 2..={MAX_USER_NAME}"),
        ));
    }
    validate_text(value.user_name(), "BrtBeginSct UserName", MAX_USER_NAME)
}

fn validate_scenario_order(value: &Scenario) -> Result<()> {
    if value.order.is_empty() {
        return Ok(());
    }
    let mut changed = vec![false; value.changed_cells().len()];
    let mut unknown = vec![false; value.unknown_records().len()];
    if value.order.len() != changed.len() + unknown.len() {
        return Err(invalid(
            "BrtBeginSct order",
            "entry count does not match records",
        ));
    }
    for entry in &value.order {
        match *entry {
            super::model::Child::Changed(index) => {
                let Some(slot) = changed.get_mut(index) else {
                    return Err(invalid(
                        "BrtBeginSct order",
                        "changed-cell index is invalid",
                    ));
                };
                if *slot {
                    return Err(invalid("BrtBeginSct order", "duplicate changed-cell index"));
                }
                *slot = true;
            },
            super::model::Child::Unknown(index) => {
                let Some(slot) = unknown.get_mut(index) else {
                    return Err(invalid(
                        "BrtBeginSct order",
                        "unknown-record index is invalid",
                    ));
                };
                if *slot {
                    return Err(invalid(
                        "BrtBeginSct order",
                        "duplicate unknown-record index",
                    ));
                }
                *slot = true;
            },
        }
    }
    if changed.into_iter().any(|seen| !seen) || unknown.into_iter().any(|seen| !seen) {
        return Err(invalid("BrtBeginSct order", "record is missing from order"));
    }
    Ok(())
}

pub(crate) fn validate_scenario(value: &Scenario) -> Result<()> {
    validate_scenario_text(value)?;
    if !(1..=MAX_CHANGED_CELLS).contains(&value.changed_cells().len()) {
        return Err(invalid(
            "BrtBeginSct cref",
            format!(
                "{} changed cells is outside 1..={MAX_CHANGED_CELLS}",
                value.changed_cells().len()
            ),
        ));
    }
    let mut coordinates = HashSet::with_capacity(value.changed_cells().len());
    for cell in value.changed_cells() {
        validate_changed_cell(cell)?;
        if !coordinates.insert((cell.row(), cell.column())) {
            return Err(invalid(
                "BrtBeginSct cells",
                format!("duplicate cell ({}, {})", cell.row(), cell.column()),
            ));
        }
    }
    if value.unknown_records().len() > MAX_UNKNOWN_RECORDS {
        return Err(invalid(
            "BrtBeginSct unknown records",
            "record count exceeds safety limit",
        ));
    }
    let bytes = value
        .unknown_records()
        .iter()
        .try_fold(0usize, |total, record| {
            total
                .checked_add(record.payload().len())
                .ok_or_else(|| invalid("BrtBeginSct unknown records", "byte count overflow"))
        })?;
    if bytes > MAX_UNKNOWN_PAYLOAD {
        return Err(invalid(
            "BrtBeginSct unknown records",
            "byte count exceeds safety limit",
        ));
    }
    validate_scenario_order(value)
}

fn validate_manager_order(value: &Manager) -> Result<()> {
    if value.order.is_empty() {
        return Ok(());
    }
    let mut scenarios = vec![false; value.scenarios().len()];
    let mut unknown = vec![false; value.unknown_records().len()];
    if value.order.len() != scenarios.len() + unknown.len() {
        return Err(invalid(
            "BrtBeginScenMan order",
            "entry count does not match records",
        ));
    }
    for entry in &value.order {
        match *entry {
            super::model::Entry::Scenario(index) => {
                let Some(slot) = scenarios.get_mut(index) else {
                    return Err(invalid(
                        "BrtBeginScenMan order",
                        "scenario index is invalid",
                    ));
                };
                if *slot {
                    return Err(invalid("BrtBeginScenMan order", "duplicate scenario index"));
                }
                *slot = true;
            },
            super::model::Entry::Unknown(index) => {
                let Some(slot) = unknown.get_mut(index) else {
                    return Err(invalid(
                        "BrtBeginScenMan order",
                        "unknown-record index is invalid",
                    ));
                };
                if *slot {
                    return Err(invalid(
                        "BrtBeginScenMan order",
                        "duplicate unknown-record index",
                    ));
                }
                *slot = true;
            },
        }
    }
    if scenarios.into_iter().any(|seen| !seen) || unknown.into_iter().any(|seen| !seen) {
        return Err(invalid(
            "BrtBeginScenMan order",
            "record is missing from order",
        ));
    }
    Ok(())
}

pub(crate) fn validate_manager(value: &Manager) -> Result<()> {
    if value.scenarios().len() > MAX_SCENARIOS {
        return Err(invalid(
            "BrtBeginScenMan",
            "scenario count exceeds safety limit",
        ));
    }
    if value
        .current()
        .is_some_and(|index| index >= value.scenarios().len())
    {
        return Err(invalid(
            "BrtBeginScenMan isctCur",
            "scenario index is out of bounds",
        ));
    }
    if value
        .shown()
        .is_some_and(|index| index >= value.scenarios().len())
    {
        return Err(invalid(
            "BrtBeginScenMan isctShown",
            "scenario index is out of bounds",
        ));
    }
    if value.result_ranges().len() > MAX_RESULT_RANGES {
        return Err(invalid(
            "BrtBeginScenMan sqrfxResult",
            "range count exceeds 32",
        ));
    }
    let mut cells = 0u64;
    for range in value.result_ranges() {
        validate_range(*range)?;
        cells = cells
            .checked_add(range.cell_count())
            .ok_or_else(|| invalid("BrtBeginScenMan sqrfxResult", "cell count overflow"))?;
    }
    if cells > 32 {
        return Err(invalid(
            "BrtBeginScenMan sqrfxResult",
            format!("result range contains {cells} cells, maximum is 32"),
        ));
    }

    let mut names = HashSet::with_capacity(value.scenarios().len());
    let mut unknown_count = value.unknown_records().len();
    let mut unknown_bytes = value
        .unknown_records()
        .iter()
        .try_fold(0usize, |total, record| {
            total
                .checked_add(record.payload().len())
                .ok_or_else(|| invalid("BrtBeginScenMan unknown records", "byte count overflow"))
        })?;
    for scenario in value.scenarios() {
        scenario.validate()?;
        let identity = scenario.name().to_lowercase();
        if !names.insert(identity) {
            return Err(invalid("BrtBeginSct Name", "scenario names must be unique"));
        }
        unknown_count = unknown_count
            .checked_add(scenario.unknown_records().len())
            .ok_or_else(|| invalid("scenario unknown records", "record count overflow"))?;
        let scenario_bytes =
            scenario
                .unknown_records()
                .iter()
                .try_fold(0usize, |total, record| {
                    total
                        .checked_add(record.payload().len())
                        .ok_or_else(|| invalid("scenario unknown records", "byte count overflow"))
                })?;
        unknown_bytes = unknown_bytes
            .checked_add(scenario_bytes)
            .ok_or_else(|| invalid("scenario unknown records", "byte count overflow"))?;
    }
    if unknown_count > MAX_UNKNOWN_RECORDS {
        return Err(invalid(
            "BrtBeginScenMan unknown records",
            "record count exceeds safety limit",
        ));
    }
    if unknown_bytes > MAX_UNKNOWN_PAYLOAD {
        return Err(invalid(
            "BrtBeginScenMan unknown records",
            "byte count exceeds safety limit",
        ));
    }
    validate_manager_order(value)
}
