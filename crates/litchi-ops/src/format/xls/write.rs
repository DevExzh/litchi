//! Unified scalar XLS writes over the canonical native package transaction.
//!
//! The BIFF8 editor owns all publication and dependency checks.  This module
//! only admits the unified target/property shape, expands bounded selectors,
//! coerces the common scalar union, and publishes one native `Package::edit`
//! for the complete operation.

use std::collections::HashSet;

use litchi_core::FileFormat;
use litchi_xls::cell_values::{Reference, Selector, Value as NativeValue};

use super::node::{CellState, Item, Node};
use super::spec::{CELL_WRITE_TABLE, CellSetter, Kind, SHEET_WRITE_TABLE, SheetSetter};
use super::{Nav, Xls};
use crate::error::{Code, Error, Result};
use crate::props::{IntoProps, PropValue, Props};
use crate::registry::{Access, Operation};
use crate::target::TargetArg;
use crate::transaction::{Error as TransactionError, LimitKind, SemanticLimits};
use crate::value::Scalar;

/// The native proof and immutable endpoints of one unified XLS write.
pub(crate) struct WriteCommit {
    pub(crate) before: Xls,
    pub(crate) after: Xls,
    pub(crate) patch: litchi_xls::Patch,
    pub(crate) operations: usize,
    pub(crate) snapshot_bytes: usize,
}

#[derive(Debug, Clone)]
enum CellAssignment {
    Value(NativeValue),
    Clear,
    Noop,
}

#[derive(Debug, Clone)]
struct CellPlan {
    assignment: CellAssignment,
    staged_bytes: usize,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
struct CellTarget {
    sheet: usize,
    reference: Reference,
    present: bool,
    formula: bool,
}

#[derive(Debug, Clone)]
enum SheetAssignment {
    Name(String),
}

const WRITE_ENTRY_OVERHEAD: usize = 64;

/// Admit one scalar cell or worksheet-name mutation and publish it through
/// the native immutable XLS package owner.
pub(crate) fn set_with_limits<'p, 'v>(
    file: &Xls,
    target: TargetArg<'p>,
    props: impl IntoProps<'v>,
    limits: &SemanticLimits,
) -> Result<WriteCommit> {
    file.admit_edit()?;
    let props = Props::from_props(props)?;

    let result = if is_sheet_property(&props) {
        set_sheet_with_limits(file, target, &props, limits)
    } else {
        set_cells_with_limits(file, target, &props, limits)
    };
    file.check_source()?;
    result
}

fn is_sheet_property(props: &Props<'_>) -> bool {
    props
        .iter()
        .any(|(key, _)| key.eq_ignore_ascii_case("name"))
}

fn set_sheet_with_limits(
    file: &Xls,
    target: TargetArg<'_>,
    props: &Props<'_>,
    limits: &SemanticLimits,
) -> Result<WriteCommit> {
    let (assignment, staged_bytes) = admit_sheet_property(props, limits)?;
    let positions = resolve_sheet_targets(file, &target, limits)?;
    let staged = positions
        .len()
        .checked_mul(staged_bytes)
        .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))?;
    if staged > limits.max_staged_bytes() {
        return Err(staged_limit(staged, limits.max_staged_bytes()));
    }
    if positions.len() > 1 {
        return Err(Error::new(
            Code::InvalidArgument,
            "an XLS worksheet rename requires exactly one sheet target",
        )
        .with_format(FileFormat::Xls)
        .with_property("name"));
    }
    let Some(position) = positions.into_iter().next() else {
        return no_op(file, limits);
    };
    if limits.max_operations() < 1 {
        return Err(operation_limit(1, limits.max_operations()));
    }

    let SheetAssignment::Name(name) = assignment;
    let mut edit = file.editable_package()?.edit().map_err(native_edit_error)?;
    edit.rename_sheet(Selector::Position(position), &name)
        .map_err(native_edit_error)?;
    let native = edit.commit().map_err(native_edit_error)?;
    finish_commit(file, native, 1, limits)
}

fn set_cells_with_limits(
    file: &Xls,
    target: TargetArg<'_>,
    props: &Props<'_>,
    limits: &SemanticLimits,
) -> Result<WriteCommit> {
    let plan = admit_cell_property(props, limits)?;
    let targets = resolve_cell_targets(file, &target, limits)?;
    let staged = targets
        .len()
        .checked_mul(plan.staged_bytes)
        .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))?;
    if staged > limits.max_staged_bytes() {
        return Err(staged_limit(staged, limits.max_staged_bytes()));
    }
    if targets.is_empty() {
        return no_op(file, limits);
    }
    if matches!(&plan.assignment, CellAssignment::Noop) {
        return no_op(file, limits);
    }
    if limits.max_operations() < 1 {
        return Err(operation_limit(1, limits.max_operations()));
    }
    if targets
        .iter()
        .any(|target| target.present && target.formula)
    {
        let property = if matches!(&plan.assignment, CellAssignment::Clear) {
            "clear"
        } else {
            "value"
        };
        return Err(Error::new(
            Code::UnsupportedCapability,
            "formula cells require formula-cache or formula-token editing",
        )
        .with_format(FileFormat::Xls)
        .with_property(property));
    }

    let mut edit = file.editable_package()?.edit().map_err(native_edit_error)?;
    for target in targets {
        let selector = Selector::Position(target.sheet);
        if matches!(&plan.assignment, CellAssignment::Clear) {
            // Keep an explicit BIFF8 blank record so the existing cell's
            // style survives and the read projection retains its `empty`
            // fact. A virtual coordinate is already clear and stays a no-op.
            if target.present {
                edit.set_value(selector, target.reference, NativeValue::Blank)
                    .map_err(native_edit_error)?;
            }
            continue;
        }
        let value = match &plan.assignment {
            CellAssignment::Value(value) => value.clone(),
            CellAssignment::Clear => unreachable!("clear handled above"),
            CellAssignment::Noop => unreachable!("no-op handled above"),
        };
        if target.present {
            edit.set_value(selector, target.reference, value)
                .map_err(native_edit_error)?;
        } else {
            edit.insert_cell(selector, target.reference, value)
                .map_err(native_edit_error)?;
        }
    }
    let native = edit.commit().map_err(native_edit_error)?;
    finish_commit(file, native, 1, limits)
}

fn admit_sheet_property(
    props: &Props<'_>,
    limits: &SemanticLimits,
) -> Result<(SheetAssignment, usize)> {
    let mut name = None;
    let mut staged_bytes = WRITE_ENTRY_OVERHEAD;
    for (key, value) in props.iter() {
        let (_, entry) = SHEET_WRITE_TABLE.resolve(key.as_str()).map_err(|error| {
            error
                .with_format(FileFormat::Xls)
                .with_property(key.as_str().to_owned())
        })?;
        let entry = *entry;
        entry
            .prepare(Operation::Set, value, |setter, value| match setter {
                SheetSetter::Name => {
                    if name.is_some() {
                        return Err(duplicate_sheet_property());
                    }
                    let Some(Scalar::Text(value)) = value else {
                        return Err(Error::new(
                            Code::InvalidValue,
                            "an XLS worksheet name requires text",
                        )
                        .with_format(FileFormat::Xls)
                        .with_property("name"));
                    };
                    staged_bytes = staged_bytes
                        .checked_add(value.len())
                        .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))?;
                    name = Some(value.to_owned());
                    Ok(())
                },
            })
            .map_err(|error| error.with_format(FileFormat::Xls))?;
    }
    let Some(name) = name else {
        return Err(Error::new(
            Code::InvalidArgument,
            "an XLS worksheet write requires the `name` property",
        )
        .with_format(FileFormat::Xls));
    };
    Ok((SheetAssignment::Name(name), staged_bytes))
}

fn admit_cell_property(props: &Props<'_>, limits: &SemanticLimits) -> Result<CellPlan> {
    let mut assignment = CellAssignment::Noop;
    let mut staged_bytes = WRITE_ENTRY_OVERHEAD;
    for (key, value) in props.iter() {
        let (_, entry) = CELL_WRITE_TABLE.resolve(key.as_str()).map_err(|error| {
            error
                .with_format(FileFormat::Xls)
                .with_property(key.as_str().to_owned())
        })?;
        let entry = *entry;
        // The shared Scalar registry preserves an exact integer only through
        // i64. BIFF8 also has an exact IEEE number domain, so retain the
        // unsigned values that round-trip through f64 (including 2^63) while
        // refusing values that would be rounded. This exception is scoped to
        // the XLS value setter; other formats keep their scalar policy.
        if matches!(entry.access, Access::Set(CellSetter::Value))
            && let PropValue::Unsigned(value) = value
        {
            let number = *value as f64;
            if !number.is_finite() || number as u128 != u128::from(*value) {
                return Err(invalid_scalar());
            }
            stage_cell_value(
                &mut assignment,
                &mut staged_bytes,
                NativeValue::Number(number),
                limits,
            )?;
            continue;
        }
        entry
            .prepare(Operation::Set, value, |setter, value| match setter {
                CellSetter::Value => {
                    let native = match value {
                        Some(value) => native_scalar(value)?,
                        None => NativeValue::Blank,
                    };
                    stage_cell_value(&mut assignment, &mut staged_bytes, native, limits)
                },
                CellSetter::Clear => {
                    if !matches!(assignment, CellAssignment::Noop) {
                        return Err(Error::new(
                            Code::InvalidArgument,
                            "cell properties value and clear cannot be supplied together",
                        )
                        .with_format(FileFormat::Xls)
                        .with_property("clear"));
                    }
                    let Some(Scalar::Bool(true)) = value else {
                        return Err(Error::new(Code::InvalidValue, "clear must be true")
                            .with_format(FileFormat::Xls)
                            .with_property("clear"));
                    };
                    staged_bytes = staged_bytes
                        .checked_add(1)
                        .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))?;
                    assignment = CellAssignment::Clear;
                    Ok(())
                },
            })
            .map_err(|error| error.with_format(FileFormat::Xls))?;
    }
    Ok(CellPlan {
        assignment,
        staged_bytes,
    })
}

fn native_scalar(value: Scalar<'_>) -> Result<NativeValue> {
    match value {
        Scalar::Bool(value) => Ok(NativeValue::Boolean(value)),
        Scalar::Text(value) => {
            validate_cell_text(value)?;
            Ok(NativeValue::Text(value.to_string()))
        },
        Scalar::Int(value) => {
            let number = value as f64;
            if number.is_finite() && number as i128 == i128::from(value) {
                Ok(NativeValue::Number(number))
            } else {
                Err(invalid_scalar())
            }
        },
        Scalar::Number(value) if valid_number(value) => Ok(NativeValue::Number(value)),
        Scalar::Number(_) => Err(invalid_scalar()),
        Scalar::Token(_) | Scalar::Length(_) | Scalar::Color(_) => Err(invalid_scalar()),
    }
}

fn stage_cell_value(
    assignment: &mut CellAssignment,
    staged_bytes: &mut usize,
    value: NativeValue,
    limits: &SemanticLimits,
) -> Result<()> {
    if !matches!(assignment, CellAssignment::Noop) {
        return Err(duplicate_cell_property("value"));
    }
    *staged_bytes = staged_bytes
        .checked_add(native_value_bytes(&value))
        .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))?;
    *assignment = CellAssignment::Value(value);
    Ok(())
}

fn valid_number(value: f64) -> bool {
    value.is_normal() || value.to_bits() == 0
}

fn validate_cell_text(value: &str) -> Result<()> {
    if value.encode_utf16().count() > 65_535
        || value
            .chars()
            .any(|character| !is_xml_10_character(character))
    {
        return Err(invalid_scalar());
    }
    Ok(())
}

const fn is_xml_10_character(character: char) -> bool {
    matches!(character, '\u{9}' | '\u{a}' | '\u{d}')
        || (character >= '\u{20}' && character <= '\u{d7ff}')
        || (character >= '\u{e000}' && character <= '\u{fffd}')
        || (character >= '\u{10000}' && character <= '\u{10ffff}')
}

fn native_value_bytes(value: &NativeValue) -> usize {
    match value {
        NativeValue::Text(value) => value.len(),
        _ => 8,
    }
}

fn resolve_sheet_targets(
    file: &Xls,
    target: &TargetArg<'_>,
    limits: &SemanticLimits,
) -> Result<Vec<usize>> {
    let mut positions = Vec::new();
    let mut seen = HashSet::new();
    if let Some(path) = target.path() {
        let node = crate::path::resolve::get(&Nav::new(file), path)?;
        append_sheet(&mut positions, &mut seen, node, limits)?;
    } else if let Some(query) = target.query() {
        let walk = crate::query::run(Nav::new(file), query.clone())?;
        if !walk.plan().diagnostics().is_empty() {
            return Err(invalid_query());
        }
        for node in walk {
            append_sheet(&mut positions, &mut seen, node?, limits)?;
        }
    } else {
        return Err(empty_target());
    }
    file.check_source()?;
    Ok(positions)
}

fn append_sheet(
    positions: &mut Vec<usize>,
    seen: &mut HashSet<usize>,
    node: Node<'_>,
    limits: &SemanticLimits,
) -> Result<()> {
    if node.element() != Kind::Sheet {
        return Err(Error::new(
            Code::UnsupportedElement,
            "an XLS worksheet write target must be a sheet",
        )
        .with_format(FileFormat::Xls));
    }
    let position = node.sheet_position();
    if seen.contains(&position) {
        return Ok(());
    }
    let count = positions
        .len()
        .checked_add(1)
        .ok_or_else(|| expanded_limit(usize::MAX, limits.max_expanded_targets()))?;
    if count > limits.max_expanded_targets() {
        return Err(expanded_limit(count, limits.max_expanded_targets()));
    }
    positions
        .try_reserve(1)
        .map_err(|_| allocation("XLS worksheet targets"))?;
    seen.try_reserve(1)
        .map_err(|_| allocation("XLS worksheet target identities"))?;
    seen.insert(position);
    positions.push(position);
    Ok(())
}

fn resolve_cell_targets(
    file: &Xls,
    target: &TargetArg<'_>,
    limits: &SemanticLimits,
) -> Result<Vec<CellTarget>> {
    let mut targets = Vec::new();
    let mut seen = HashSet::new();
    let mut non_cell = false;
    if let Some(path) = target.path() {
        let node = crate::path::resolve::get(&Nav::new(file), path)?;
        append_cell_node(&mut targets, &mut seen, &mut non_cell, node, limits)?;
    } else if let Some(query) = target.query() {
        let walk = crate::query::run(Nav::new(file), query.clone())?;
        if !walk.plan().diagnostics().is_empty() {
            return Err(invalid_query());
        }
        for node in walk {
            append_cell_node(&mut targets, &mut seen, &mut non_cell, node?, limits)?;
        }
    } else {
        return Err(empty_target());
    }
    file.check_source()?;
    if non_cell {
        return Err(Error::new(
            Code::UnsupportedElement,
            "an XLS cell write target must be a cell or range",
        )
        .with_format(FileFormat::Xls));
    }
    Ok(targets)
}

fn append_cell_node(
    targets: &mut Vec<CellTarget>,
    seen: &mut HashSet<(usize, Reference)>,
    non_cell: &mut bool,
    node: Node<'_>,
    limits: &SemanticLimits,
) -> Result<()> {
    match node.element() {
        Kind::Cell => {
            let Some((column, row)) = node.cell() else {
                return Err(
                    Error::new(Code::InvalidPath, "the XLS cell reference is invalid")
                        .with_format(FileFormat::Xls),
                );
            };
            let reference = Reference::new(row, column).map_err(native_value_error)?;
            let (present, formula) = cell_state_facts(node.cell_state()?);
            append_cell_target(
                targets,
                seen,
                node.sheet_position(),
                reference,
                present,
                formula,
                limits,
            )
        },
        Kind::Range => {
            let Item::Range(range) = node.item() else {
                return Err(Error::new(Code::InvalidPath, "the XLS range is invalid")
                    .with_format(FileFormat::Xls));
            };
            let (first_row, first_column, last_row, last_column) = range.rect;
            let rows = last_row.saturating_sub(first_row).saturating_add(1) as usize;
            let columns = last_column.saturating_sub(first_column).saturating_add(1) as usize;
            let area = rows
                .checked_mul(columns)
                .ok_or_else(|| expanded_limit(usize::MAX, limits.max_expanded_targets()))?;
            let total = targets.len().checked_add(area).unwrap_or(usize::MAX);
            if area > limits.max_expanded_targets() || total > limits.max_expanded_targets() {
                return Err(expanded_limit(total, limits.max_expanded_targets()));
            }
            targets
                .try_reserve(area)
                .map_err(|_| allocation("XLS range cell targets"))?;
            seen.try_reserve(area)
                .map_err(|_| allocation("XLS range cell identities"))?;
            for row in first_row..=last_row {
                for column in first_column..=last_column {
                    let reference = Reference::new(row, column).map_err(native_value_error)?;
                    let (present, formula) = range
                        .grid
                        .record((row, column))
                        .map_or((false, false), |cell| (true, cell.formula));
                    append_cell_target(
                        targets,
                        seen,
                        node.sheet_position(),
                        reference,
                        present,
                        formula,
                        limits,
                    )?;
                }
            }
            Ok(())
        },
        _ => {
            *non_cell = true;
            Ok(())
        },
    }
}

fn append_cell_target(
    targets: &mut Vec<CellTarget>,
    seen: &mut HashSet<(usize, Reference)>,
    sheet: usize,
    reference: Reference,
    present: bool,
    formula: bool,
    limits: &SemanticLimits,
) -> Result<()> {
    if seen.contains(&(sheet, reference)) {
        return Ok(());
    }
    let count = targets
        .len()
        .checked_add(1)
        .ok_or_else(|| expanded_limit(usize::MAX, limits.max_expanded_targets()))?;
    if count > limits.max_expanded_targets() {
        return Err(expanded_limit(count, limits.max_expanded_targets()));
    }
    targets
        .try_reserve(1)
        .map_err(|_| allocation("XLS cell targets"))?;
    seen.try_reserve(1)
        .map_err(|_| allocation("XLS cell target identities"))?;
    seen.insert((sheet, reference));
    targets.push(CellTarget {
        sheet,
        reference,
        present,
        formula,
    });
    Ok(())
}

fn cell_state_facts(state: CellState<'_>) -> (bool, bool) {
    match state {
        CellState::Stored(cell) => (true, cell.formula),
        CellState::Missing | CellState::Stub => (false, false),
    }
}

fn finish_commit(
    file: &Xls,
    native: litchi_xls::Commit,
    fallback_operations: usize,
    limits: &SemanticLimits,
) -> Result<WriteCommit> {
    file.check_source()?;
    let changed = native.diagnostics().changed();
    let operations =
        native
            .diagnostics()
            .operation_count()
            .max(if changed { fallback_operations } else { 0 });
    let (package, patch) = native.into_parts();
    if operations > limits.max_operations() {
        return Err(operation_limit(operations, limits.max_operations()));
    }
    let snapshot_bytes = patch.retained_bytes();
    if snapshot_bytes > limits.max_snapshot_bytes() {
        return Err(snapshot_limit(snapshot_bytes, limits.max_snapshot_bytes()));
    }
    let after = if changed {
        file.with_package(package)?
    } else {
        file.clone()
    };
    file.check_source()?;
    Ok(WriteCommit {
        before: file.clone(),
        after,
        patch,
        operations,
        snapshot_bytes,
    })
}

fn no_op(file: &Xls, limits: &SemanticLimits) -> Result<WriteCommit> {
    let native = file
        .editable_package()?
        .edit()
        .map_err(native_edit_error)?
        .commit()
        .map_err(native_edit_error)?;
    finish_commit(file, native, 0, limits)
}

fn duplicate_cell_property(property: &'static str) -> Error {
    Error::new(
        Code::InvalidArgument,
        "the same XLS cell property was supplied more than once",
    )
    .with_format(FileFormat::Xls)
    .with_property(property)
}

fn duplicate_sheet_property() -> Error {
    Error::new(
        Code::InvalidArgument,
        "the same XLS sheet property was supplied more than once",
    )
    .with_format(FileFormat::Xls)
    .with_property("name")
}

fn invalid_scalar() -> Error {
    Error::new(
        Code::InvalidValue,
        "XLS cell values support text, bool, and finite BIFF8 numbers",
    )
    .with_format(FileFormat::Xls)
    .with_property("value")
}

fn empty_target() -> Error {
    Error::new(Code::InvalidArgument, "the mutation target is empty").with_format(FileFormat::Xls)
}

fn allocation(resource: &'static str) -> Error {
    Error::new(
        Code::InternalError,
        "could not allocate XLS mutation targets",
    )
    .with_format(FileFormat::Xls)
    .with_suggestion(resource)
}

fn expanded_limit(observed: usize, maximum: usize) -> Error {
    Error::from(TransactionError::Limit {
        kind: LimitKind::ExpandedTargets,
        observed,
        maximum,
    })
    .with_format(FileFormat::Xls)
}

fn staged_limit(observed: usize, maximum: usize) -> Error {
    Error::from(TransactionError::Limit {
        kind: LimitKind::StagedBytes,
        observed,
        maximum,
    })
    .with_format(FileFormat::Xls)
}

fn snapshot_limit(observed: usize, maximum: usize) -> Error {
    Error::from(TransactionError::Limit {
        kind: LimitKind::SnapshotBytes,
        observed,
        maximum,
    })
    .with_format(FileFormat::Xls)
}

fn operation_limit(observed: usize, maximum: usize) -> Error {
    Error::from(TransactionError::Limit {
        kind: LimitKind::Operations,
        observed,
        maximum,
    })
    .with_format(FileFormat::Xls)
}

fn native_error(error: litchi_xls::package::edit::Error) -> Error {
    Error::from(error).with_format(FileFormat::Xls)
}

fn native_edit_error(error: litchi_xls::package::edit::Error) -> Error {
    native_error(error)
}

fn native_value_error(error: litchi_xls::Error) -> Error {
    Error::from(error).with_format(FileFormat::Xls)
}

fn invalid_query() -> Error {
    Error::new(
        Code::InvalidSelector,
        "the mutation selector contains an unsupported predicate",
    )
    .with_format(FileFormat::Xls)
}
