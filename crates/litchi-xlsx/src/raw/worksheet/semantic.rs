//! Raw-cell materialization and formula semantics.

use std::collections::{HashMap, HashSet};

use super::super::formula::{Range as FormulaRange, translate};
use super::super::strings::decode_spreadsheet_text;
use super::model::{
    MAX_CELL_CHARACTERS, RawCell, RawFormula, RawFormulaKind, SharedMaster, SharedMember,
};
use crate::cell::{
    Cell, Date, ErrorValue, Number, SharedFormulaStorage, Stored, Text, Unknown, Value,
};
use crate::error::{Result, allocation, invalid};
use crate::formula::{Cache, Formula, Kind};
use crate::leniency::{CellDefect, Location, ToleranceReport};
use litchi_sheet::{Cell as Address, Rect};

/// Materialize one raw cell record.
///
/// With `tolerance` present, a dangling shared-string reference (audit
/// finding 4.4) is repaired to an empty string and recorded at the cell;
/// without it, the reference fails the parse as it historically did.
pub(super) fn materialize(
    raw: RawCell<'_>,
    strings: Option<&[Text]>,
    mut tolerance: Option<&mut ToleranceReport>,
) -> Result<Stored> {
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
                // A repaired dangling reference names no item, so it must not
                // be retained as transferable shared-string provenance.
                .filter(|index| strings.is_some_and(|strings| *index < strings.len()))
        } else {
            None
        };
    let address = raw.address;
    let unknown_cell_type = raw
        .cell_type
        .as_deref()
        .filter(|kind| !matches!(*kind, "b" | "d" | "e" | "inlineStr" | "n" | "s" | "str"));
    let cell = if let Some(kind) = unknown_cell_type {
        let formula = raw.formula.map(|formula| formula.text);
        Cell::Unknown(Unknown::new(
            kind,
            raw.value.map(std::borrow::Cow::into_owned),
            formula,
        ))
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
        // A malformed cached value (finding 4.3) is dropped: the formula is
        // authoritative and Excel recalculates it.
        let cached = match raw.value.as_deref() {
            Some(value) => match parse_value(
                raw.cell_type.as_deref(),
                value,
                strings,
                address,
                tolerance.as_deref_mut(),
            )? {
                Parsed::Value(value) => value.map(Cache::stored),
                Parsed::Malformed => None,
            },
            None => None,
        };
        Cell::Formula(Formula::parsed(text, kind, cached))
    } else if let Some(value) = raw.value.as_deref() {
        match parse_value(raw.cell_type.as_deref(), value, strings, address, tolerance)? {
            Parsed::Value(value) => value.map_or(Cell::Empty, Cell::Value),
            // Demoted with its lexical preserved (finding 4.3).
            Parsed::Malformed => Cell::Unknown(Unknown::new(
                raw.cell_type.as_deref().unwrap_or("n"),
                Some(value.to_owned()),
                None,
            )),
        }
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
        shared_formula: raw.shared_formula,
        cell_metadata: raw.cell_metadata,
        value_metadata: raw.value_metadata,
    })
}

fn boxed_str(value: &str, context: &'static str) -> Result<Box<str>> {
    let mut copied = String::new();
    copied
        .try_reserve(value.len())
        .map_err(|source| allocation(context, source))?;
    copied.push_str(value);
    Ok(copied.into_boxed_str())
}

fn formula_range_rect(range: FormulaRange) -> Result<Rect> {
    let start = Address::at(range.first_row - 1, range.first_column - 1)?;
    Ok(Rect::new(start, range.last_row, range.last_column)?)
}

/// The outcome of reading one `<v>` lexical.
enum Parsed {
    /// The typed value (`None` for an empty numeric lexical).
    Value(Option<Value>),
    /// A number or boolean lexical no reading can type (audit finding 4.3);
    /// the repair was recorded and the caller demotes the cell.
    Malformed,
}

/// Read one `<v>` lexical by cell type.
///
/// With `tolerance` present, malformed lexicals (audit finding 4.3) are
/// repaired and recorded at the cell: an unparseable or non-finite number is
/// [`Parsed::Malformed`]; a boolean is read the way LibreOffice reads it —
/// `true`/`false` in any ASCII case, any number as non-zero-is-true — and
/// only a lexical neither form covers is [`Parsed::Malformed`]; an invalid
/// date is demoted to its text. Without `tolerance` each fails the parse as
/// it historically did.
fn parse_value(
    cell_type: Option<&str>,
    value: &str,
    strings: Option<&[Text]>,
    address: Address,
    tolerance: Option<&mut ToleranceReport>,
) -> Result<Parsed> {
    let parsed = |value: Value| Ok(Parsed::Value(Some(value)));
    match cell_type {
        None | Some("n") if value.trim().is_empty() => Ok(Parsed::Value(None)),
        None | Some("n") => match Number::new(value) {
            Ok(number) => parsed(Value::Number(number)),
            Err(error) => {
                let tolerance = tolerance.ok_or(error)?;
                tolerance.record(CellDefect::MalformedNumber, Location::Cell(address));
                Ok(Parsed::Malformed)
            },
        },
        Some("str") => {
            let value = decode_spreadsheet_text(value)?;
            if value.chars().count() > MAX_CELL_CHARACTERS {
                return Err(invalid(format!(
                    "worksheet string exceeds {MAX_CELL_CHARACTERS} characters"
                )));
            }
            parsed(Value::Text(value.into()))
        },
        Some("d") => match Date::new(value) {
            Ok(date) => parsed(Value::Date(date)),
            Err(error) => {
                let tolerance = tolerance.ok_or(error)?;
                tolerance.record(CellDefect::MalformedDate, Location::Cell(address));
                parsed(Value::Text(value.into()))
            },
        },
        Some("s") => {
            let index = parse_shared_string_index(value)?;
            let (defect, error) = match strings {
                Some(strings) => match strings.get(index) {
                    Some(text) => return parsed(Value::Text(text.clone())),
                    None => (
                        CellDefect::SharedStringIndexOutOfRange,
                        invalid(format!(
                            "shared-string index {index} exceeds table length {}",
                            strings.len()
                        )),
                    ),
                },
                None => (
                    CellDefect::MissingSharedStringTable,
                    invalid(
                        "worksheet uses shared strings but the workbook has no shared-string part",
                    ),
                ),
            };
            // Excel repairs the cell and LibreOffice resolves the lookup to
            // empty (audit finding 4.4): substitute an empty string.
            let tolerance = tolerance.ok_or(error)?;
            tolerance.record(defect, Location::Cell(address));
            parsed(Value::Text(Text::from("")))
        },
        Some("b") => match value.trim() {
            "1" | "true" => parsed(Value::Bool(true)),
            "0" | "false" => parsed(Value::Bool(false)),
            other => {
                let error = invalid(format!("invalid worksheet boolean '{other}'"));
                let tolerance = tolerance.ok_or(error)?;
                tolerance.record(CellDefect::MalformedBoolean, Location::Cell(address));
                Ok(match lenient_boolean(other) {
                    Some(value) => Parsed::Value(Some(Value::Bool(value))),
                    None => Parsed::Malformed,
                })
            },
        },
        Some("e") => parsed(Value::Error(ErrorValue::parse(value))),
        Some("inlineStr") => Err(invalid(
            "inline-string cell stores text in an is element, not v",
        )),
        Some(other) => Err(invalid(format!(
            "unsupported worksheet cell type '{other}'"
        ))),
    }
}

/// LibreOffice's boolean reading ("Some generators may write true or false
/// instead of 1 or 0", `sheetdatacontext.cxx`): `true`/`false` in any ASCII
/// case, and any finite number as non-zero-is-true. `None` for a lexical
/// neither form covers.
fn lenient_boolean(value: &str) -> Option<bool> {
    if value.eq_ignore_ascii_case("true") {
        return Some(true);
    }
    if value.eq_ignore_ascii_case("false") {
        return Some(false);
    }
    value
        .parse::<f64>()
        .ok()
        .filter(|number| number.is_finite())
        .map(|number| number != 0.0)
}

fn parse_shared_string_index(value: &str) -> Result<usize> {
    value
        .trim()
        .parse::<usize>()
        .map_err(|_source| invalid(format!("invalid shared-string index '{value}'")))
}

/// Resolve shared-formula groups into per-cell scalar formulas.
///
/// With `tolerance` present, a structurally broken group (audit finding 4.6)
/// is repaired per cell and recorded at the cell as
/// [`CellDefect::BrokenSharedFormulaGroup`], the way LibreOffice falls back to
/// plain cells for an unusable shared reference:
///
/// * a would-be master that cannot anchor a group — no `ref`, no expression,
///   not the first cell of its `ref`, or a second master for the same `si`
///   (first wins) — keeps its own expression verbatim as an ordinary
///   formula when it has one, and otherwise reads as a plain cell holding
///   its cached value;
/// * a member whose group has no usable master, or which lies outside the
///   master's range, reads as a plain cell holding its cached value: the
///   expression is never translated from a master that does not cover it,
///   so no formula is invented.
///
/// A `ref` that is not a range lexical at all is rejected while the `f`
/// element is read and stays a hard error. Without `tolerance` each
/// inconsistency fails the parse as it historically did.
pub(super) fn resolve_shared_formulas(
    cells: &mut [RawCell<'_>],
    mut tolerance: Option<&mut ToleranceReport>,
) -> Result<()> {
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

    // Record one repair at `cell_index`, or refuse it with `error`.
    let mut repair = |cells: &[RawCell<'_>], cell_index: usize, error: crate::Error| {
        let tolerance = tolerance.as_deref_mut().ok_or(error)?;
        let cell = cells
            .get(cell_index)
            .ok_or_else(|| invalid("shared formula membership lost its cell"))?;
        tolerance.record(
            CellDefect::BrokenSharedFormulaGroup,
            Location::Cell(cell.address),
        );
        Ok::<(), crate::Error>(())
    };

    let mut masters = HashMap::<u32, SharedMaster>::new();
    let mut repaired_masters = HashSet::new();
    for member in &members {
        if member.range.is_none() && member.text.is_empty() {
            continue;
        }
        match shared_master(member, &masters) {
            Ok(master) => {
                masters.insert(member.index, master);
            },
            Err(error) => {
                repair(cells, member.cell_index, error)?;
                repaired_masters.insert(member.cell_index);
                let cell = cells
                    .get_mut(member.cell_index)
                    .ok_or_else(|| invalid("shared formula membership lost its cell"))?;
                if member.text.is_empty() {
                    cell.formula = None;
                } else if let Some(formula) = cell.formula.as_mut() {
                    formula.kind = RawFormulaKind::Scalar;
                }
            },
        }
    }

    for member in members {
        if repaired_masters.contains(&member.cell_index) {
            continue;
        }
        let is_master_record = member.range.is_some() || !member.text.is_empty();
        let master = match masters.get(&member.index) {
            Some(master) if master.range.contains(member.row, member.column) => master,
            Some(_) => {
                repair(
                    cells,
                    member.cell_index,
                    invalid(format!(
                        "shared formula at ({}, {}) lies outside its master range",
                        member.row, member.column
                    )),
                )?;
                orphan(cells, member.cell_index)?;
                continue;
            },
            None => {
                repair(
                    cells,
                    member.cell_index,
                    invalid(format!(
                        "shared formula at ({}, {}) has no master for si={}",
                        member.row, member.column, member.index
                    )),
                )?;
                orphan(cells, member.cell_index)?;
                continue;
            },
        };
        let is_master = (member.row, member.column) == (master.row, master.column);
        if !is_master && is_master_record {
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
        cell.shared_formula = Some(SharedFormulaStorage {
            index: member.index,
            range: master.rect,
            reference: boxed_str(&master.reference, "shared formula reference")?,
            master: is_master,
        });
        let formula = cell
            .formula
            .as_mut()
            .ok_or_else(|| invalid("shared formula membership lost its cell"))?;
        formula.text = text;
        formula.kind = RawFormulaKind::Scalar;
    }
    Ok(())
}

/// Validate one would-be shared-formula master against the masters accepted
/// so far.
fn shared_master(
    member: &SharedMember,
    masters: &HashMap<u32, SharedMaster>,
) -> Result<SharedMaster> {
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
    let rect = formula_range_rect(range)?;
    if (member.row, member.column) != (range.first_row, range.first_column) {
        return Err(invalid(format!(
            "shared formula master at ({}, {}) is not first in '{range_text}'",
            member.row, member.column
        )));
    }
    if masters.contains_key(&member.index) {
        return Err(invalid(format!(
            "duplicate shared formula master for si={}",
            member.index
        )));
    }
    Ok(SharedMaster {
        row: member.row,
        column: member.column,
        range,
        rect,
        reference: boxed_str(range_text, "shared formula reference")?,
        text: member.text.clone(),
    })
}

/// Drop an orphaned shared-formula member's formula so the cell reads as a
/// plain cell holding its cached value.
fn orphan(cells: &mut [RawCell<'_>], cell_index: usize) -> Result<()> {
    cells
        .get_mut(cell_index)
        .ok_or_else(|| invalid("shared formula membership lost its cell"))?
        .formula = None;
    Ok(())
}
