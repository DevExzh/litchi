//! Raw-cell materialization and formula semantics.

use std::collections::HashMap;

use super::super::formula::{Range as FormulaRange, translate};
use super::super::strings::decode_spreadsheet_text;
use super::model::{
    MAX_CELL_CHARACTERS, RawCell, RawFormula, RawFormulaKind, SharedMaster, SharedMember,
};
use crate::cell::{Cell, Date, ErrorValue, Number, Stored, Text, Unknown, Value};
use crate::error::{Result, invalid};
use crate::formula::{Cache, Formula, Kind};
use litchi_sheet::{Cell as Address, Rect};

pub(super) fn materialize(raw: RawCell, strings: Option<&[Text]>) -> Result<Stored> {
    let formula_range = if let Some(range) = raw.formula_range {
        Some(formula_range_rect(range)?)
    } else {
        raw.formula
            .as_ref()
            .and_then(|formula| match &formula.kind {
                RawFormulaKind::Array(range) | RawFormulaKind::DataTable(range) => range.as_deref(),
                RawFormulaKind::Scalar
                | RawFormulaKind::Shared { .. }
                | RawFormulaKind::Unknown(_) => None,
            })
            .map(Rect::from_a1)
            .transpose()?
    };
    let shared_string =
        if raw.formula.is_none() && raw.inline.is_none() && raw.cell_type.as_deref() == Some("s") {
            raw.value
                .as_deref()
                .map(parse_shared_string_index)
                .transpose()?
        } else {
            None
        };
    let unknown_cell_type = raw
        .cell_type
        .as_deref()
        .filter(|kind| !matches!(*kind, "b" | "d" | "e" | "inlineStr" | "n" | "s" | "str"));
    let cell = if let Some(kind) = unknown_cell_type {
        let formula = raw.formula.map(|formula| formula.text);
        Cell::Unknown(Unknown::new(kind, raw.value, formula))
    } else if let Some(inline) = raw.inline {
        if raw.formula.is_some() {
            return Err(invalid("formula cell cannot contain an inline string"));
        }
        Cell::Value(Value::Text(inline.into()))
    } else if let Some(formula) = raw.formula {
        let RawFormula { text, kind } = formula;
        let kind = match kind {
            RawFormulaKind::Scalar => Kind::Scalar,
            RawFormulaKind::Array(range) => Kind::Array {
                range: range.map(Text::from),
            },
            RawFormulaKind::DataTable(range) => Kind::DataTable {
                range: range.map(Text::from),
            },
            RawFormulaKind::Shared { .. } => {
                return Err(invalid("unresolved shared formula storage record"));
            },
            RawFormulaKind::Unknown(value) => Kind::Unknown(value.into()),
        };
        let cached = raw
            .value
            .as_deref()
            .map(|value| parse_value(raw.cell_type.as_deref(), value, strings))
            .transpose()?
            .flatten()
            .map(Cache::stored);
        Cell::Formula(Formula::parsed(text, kind, cached))
    } else if let Some(value) = raw.value.as_deref() {
        parse_value(raw.cell_type.as_deref(), value, strings)?.map_or(Cell::Empty, Cell::Value)
    } else if let Some(kind) = raw.cell_type.as_deref()
        && !matches!(kind, "n" | "inlineStr")
    {
        if kind == "str" {
            Cell::Value(Value::Text(Text::from("")))
        } else {
            Cell::Empty
        }
    } else {
        Cell::Empty
    };
    Ok(Stored {
        address: raw.address,
        cell,
        style: raw.style,
        shared_string,
        inline_rich: raw.inline_rich,
        formula_range,
        cell_metadata: raw.cell_metadata,
        value_metadata: raw.value_metadata,
    })
}

fn formula_range_rect(range: FormulaRange) -> Result<Rect> {
    let start = Address::at(range.first_row - 1, range.first_column - 1)?;
    Ok(Rect::new(start, range.last_row, range.last_column)?)
}

fn parse_value(
    cell_type: Option<&str>,
    value: &str,
    strings: Option<&[Text]>,
) -> Result<Option<Value>> {
    match cell_type {
        None | Some("n") if value.trim().is_empty() => Ok(None),
        None | Some("n") => Number::new(value).map(Value::Number).map(Some),
        Some("str") => {
            let value = decode_spreadsheet_text(value)?;
            if value.chars().count() > MAX_CELL_CHARACTERS {
                return Err(invalid(format!(
                    "worksheet string exceeds {MAX_CELL_CHARACTERS} characters"
                )));
            }
            Ok(Some(Value::Text(value.into())))
        },
        Some("d") => Date::new(value).map(Value::Date).map(Some),
        Some("s") => {
            let index = parse_shared_string_index(value)?;
            let strings = strings.ok_or_else(|| {
                invalid("worksheet uses shared strings but the workbook has no shared-string part")
            })?;
            let text = strings.get(index).ok_or_else(|| {
                invalid(format!(
                    "shared-string index {index} exceeds table length {}",
                    strings.len()
                ))
            })?;
            Ok(Some(Value::Text(text.clone())))
        },
        Some("b") => match value.trim() {
            "1" | "true" => Ok(Some(Value::Bool(true))),
            "0" | "false" => Ok(Some(Value::Bool(false))),
            other => Err(invalid(format!("invalid worksheet boolean '{other}'"))),
        },
        Some("e") => Ok(Some(Value::Error(ErrorValue::parse(value)))),
        Some("inlineStr") => Err(invalid(
            "inline-string cell stores text in an is element, not v",
        )),
        Some(other) => Err(invalid(format!(
            "unsupported worksheet cell type '{other}'"
        ))),
    }
}

fn parse_shared_string_index(value: &str) -> Result<usize> {
    value
        .trim()
        .parse::<usize>()
        .map_err(|_source| invalid(format!("invalid shared-string index '{value}'")))
}

pub(super) fn resolve_shared_formulas(cells: &mut [RawCell]) -> Result<()> {
    let mut members = Vec::new();
    for (cell_index, cell) in cells.iter().enumerate() {
        let Some(RawFormula {
            text,
            kind: RawFormulaKind::Shared { index, range },
        }) = cell.formula.as_ref()
        else {
            continue;
        };
        members.push(SharedMember {
            cell_index,
            row: cell.address.row().get() + 1,
            column: cell.address.column().get() + 1,
            index: *index,
            range: range.clone(),
            text: text.clone(),
        });
    }
    if members.is_empty() {
        return Ok(());
    }

    let mut masters = HashMap::<u32, SharedMaster>::new();
    for member in &members {
        if member.range.is_none() && member.text.is_empty() {
            continue;
        }
        let range_text = member.range.as_deref().ok_or_else(|| {
            invalid(format!(
                "shared formula master at ({}, {}) is missing ref",
                member.row, member.column
            ))
        })?;
        if member.text.is_empty() {
            return Err(invalid(format!(
                "shared formula master at ({}, {}) has no expression",
                member.row, member.column
            )));
        }
        let range = FormulaRange::parse(range_text)?;
        if (member.row, member.column) != (range.first_row, range.first_column) {
            return Err(invalid(format!(
                "shared formula master at ({}, {}) is not first in '{range_text}'",
                member.row, member.column
            )));
        }
        if masters
            .insert(
                member.index,
                SharedMaster {
                    row: member.row,
                    column: member.column,
                    range,
                    text: member.text.clone(),
                },
            )
            .is_some()
        {
            return Err(invalid(format!(
                "duplicate shared formula master for si={}",
                member.index
            )));
        }
    }

    for member in members {
        let master = masters.get(&member.index).ok_or_else(|| {
            invalid(format!(
                "shared formula at ({}, {}) has no master for si={}",
                member.row, member.column, member.index
            ))
        })?;
        if !master.range.contains(member.row, member.column) {
            return Err(invalid(format!(
                "shared formula at ({}, {}) lies outside its master range",
                member.row, member.column
            )));
        }
        let is_master = (member.row, member.column) == (master.row, master.column);
        if !is_master && (!member.text.is_empty() || member.range.is_some()) {
            return Err(invalid(format!(
                "shared formula follower at ({}, {}) contains master data",
                member.row, member.column
            )));
        }
        let text = if is_master {
            master.text.clone()
        } else {
            translate(
                &master.text,
                master.row,
                master.column,
                member.row,
                member.column,
            )
        };
        let cell = cells
            .get_mut(member.cell_index)
            .ok_or_else(|| invalid("shared formula membership lost its cell"))?;
        cell.formula_range = Some(master.range);
        let formula = cell
            .formula
            .as_mut()
            .ok_or_else(|| invalid("shared formula membership lost its cell"))?;
        formula.text = text;
        formula.kind = RawFormulaKind::Scalar;
    }
    Ok(())
}
