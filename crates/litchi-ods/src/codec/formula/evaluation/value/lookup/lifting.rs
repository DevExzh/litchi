//! Matrix lifting for lookup scalar parameters.
//!
//! The lookup data arguments are complete ForceArray values.  Only the scalar
//! key, selector, and mode arguments determine an output shape.  This module
//! keeps those two concerns separate: it projects one scalar parameter for a
//! requested output coordinate, while every search invocation borrows the
//! original data descriptor.

use super::super::super::lookup::Function;
use super::super::{
    EvaluationFailure, EvaluationResult, Resolver, RuntimeValue, ScalarError, Shape,
    ValueEvaluator, WorkingValue, ensure_capacity,
};

/// Lift ADDRESS when one of its scalar parameters is matrix-valued.
pub(super) fn address<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    node: super::super::super::Node<'expr>,
    arguments: &[RuntimeValue<'expr>],
) -> EvaluationResult<RuntimeValue<'expr>>
where
    R: Resolver + ?Sized,
{
    let shape = scalar_parameter_shape(evaluator, arguments)?.ok_or(
        EvaluationFailure::InvalidExpression("lookup address shape is unknown"),
    )?;
    map_address_arguments(evaluator, node, shape, arguments)
}

pub(super) fn should_address<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    arguments: &[RuntimeValue<'expr>],
) -> EvaluationResult<bool>
where
    R: Resolver + ?Sized,
{
    if evaluator.mode != super::super::Mode::Matrix {
        return Ok(false);
    }
    let Some(shape) = scalar_parameter_shape(evaluator, arguments)? else {
        return Ok(false);
    };
    Ok(shape != Shape::new(1, 1)?)
}

/// Lift one search function over its scalar key/selector/mode arguments.
///
/// Search data and result vectors are retained as complete values.  The
/// borrowed search entry point therefore sees the same reference or array for
/// every output coordinate without cloning or materializing its cells.
pub(super) fn search<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    node: super::super::super::Node<'expr>,
    name: &str,
    function: Function,
    arguments: &[RuntimeValue<'expr>],
) -> EvaluationResult<RuntimeValue<'expr>>
where
    R: Resolver + ?Sized,
{
    let scalar_indexes = scalar_indexes(function, arguments.len());
    let mut shape = Shape::new(1, 1)?;
    for &index in scalar_indexes {
        let Some(value) = arguments.get(index) else {
            continue;
        };
        let Some(value_shape) = parameter_shape(evaluator, value)? else {
            return Ok(formula_error(ScalarError::Value));
        };
        // Broadcasting follows the evaluator's matrix contract: each
        // dimension expands to the larger operand and an out-of-shape source
        // cell is published as #N/A by `select_matrix_value` at that output
        // coordinate.  It is therefore not an internal shape invariant
        // failure merely because two well-formed operands have unequal
        // dimensions.
        shape = broadcast_shape(shape, value_shape)?;
    }
    map_search_arguments(evaluator, node, name, shape, arguments, scalar_indexes)
}

pub(super) fn should_search<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    function: Function,
    arguments: &[RuntimeValue<'expr>],
) -> EvaluationResult<bool>
where
    R: Resolver + ?Sized,
{
    if evaluator.mode != super::super::Mode::Matrix {
        return Ok(false);
    }
    let mut shape = Shape::new(1, 1)?;
    for &index in scalar_indexes(function, arguments.len()) {
        let Some(value) = arguments.get(index) else {
            continue;
        };
        let Some(value_shape) = parameter_shape(evaluator, value)? else {
            return Ok(false);
        };
        shape = broadcast_shape(shape, value_shape)?;
    }
    Ok(shape != Shape::new(1, 1)?)
}

fn scalar_indexes(function: Function, count: usize) -> &'static [usize] {
    // The returned slices are static so the mapping loop can stay allocation
    // free.  Invalid arity is handled by the outer dispatcher.
    match function {
        Function::HLookup | Function::VLookup => match count {
            4 => &[0, 2, 3],
            _ => &[0, 2],
        },
        Function::Lookup | Function::Match => match count {
            3 if function == Function::Match => &[0, 2],
            _ => &[0],
        },
        _ => &[],
    }
}

fn scalar_parameter_shape<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    arguments: &[RuntimeValue<'expr>],
) -> EvaluationResult<Option<Shape>>
where
    R: Resolver + ?Sized,
{
    let mut shape = Shape::new(1, 1)?;
    for value in arguments {
        let Some(value_shape) = parameter_shape(evaluator, value)? else {
            return Ok(None);
        };
        shape = broadcast_shape(shape, value_shape)?;
    }
    Ok(Some(shape))
}

fn parameter_shape<'expr, 'scalar, 'exec, 'position, R>(
    _evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    value: &RuntimeValue<'expr>,
) -> EvaluationResult<Option<Shape>>
where
    R: Resolver + ?Sized,
{
    match value {
        RuntimeValue::Array(array) => Ok(Some(array.shape)),
        RuntimeValue::Areas(areas) if areas.is_list => Ok(None),
        RuntimeValue::Areas(areas) if areas.areas.len() == 1 => {
            let area = areas
                .areas
                .first()
                .ok_or(EvaluationFailure::InvalidExpression(
                    "lookup scalar reference disappeared",
                ))?;
            Ok(Some(Shape::new(area.rect.rows(), area.rect.columns())?))
        },
        // A multi-plane reference has no two-dimensional scalar-parameter
        // projection.  Refuse it before selecting or reading a key cell.
        RuntimeValue::Areas(_) => Ok(None),
        RuntimeValue::SourceReference => Err(EvaluationFailure::Unsupported(
            super::super::super::UnsupportedKind::Reference,
        )),
        RuntimeValue::Empty
        | RuntimeValue::Missing
        | RuntimeValue::Scalar(_)
        | RuntimeValue::ScalarCell(_) => Ok(Some(Shape::new(1, 1)?)),
    }
}

fn broadcast_shape(left: Shape, right: Shape) -> EvaluationResult<Shape> {
    Shape::new(
        left.rows().max(right.rows()),
        left.columns().max(right.columns()),
    )
}

fn map_address_arguments<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    node: super::super::super::Node<'expr>,
    shape: Shape,
    arguments: &[RuntimeValue<'expr>],
) -> EvaluationResult<RuntimeValue<'expr>>
where
    R: Resolver + ?Sized,
{
    let cells =
        shape
            .cell_count()
            .ok_or(EvaluationFailure::ResourceLimit(evaluator.local_limit(
                litchi_core::Resource::Objects,
                u64::MAX,
                evaluator.limits.max_array_cells,
            )))?;
    let output_reservation;
    let mut output;
    (output, output_reservation) = evaluator.new_element_vec(cells)?;

    // These are bounded by the function arity.  The leases are kept with the
    // scratch vectors so both vectors are dropped before their reservations.
    let mut values_reservation = None;
    let mut values = Vec::new();
    ensure_capacity(
        &mut values,
        &mut values_reservation,
        arguments.len(),
        evaluator.limits.scalar.max_stack_entries,
        evaluator.execution,
        &evaluator.storage_budget,
        "formula lookup scalar arguments",
    )?;
    for index in 0..cells {
        evaluator.charge_cell_work(index)?;
        values.clear();
        for argument in arguments {
            let element = evaluator.select_matrix_value(argument, shape, index)?;
            values.push(evaluator.element_to_runtime(element)?);
        }
        let mut scalar_arguments = std::mem::take(&mut values);
        let value = super::apply_address_scalar(evaluator, node, &mut scalar_arguments)?;
        values = scalar_arguments;
        // ADDRESS returns Text and search returns scalar values; both are valid
        // array elements.  A descriptor here indicates a violated scalar
        // result contract and is surfaced as the existing typed array refusal.
        output.push(evaluator.runtime_to_element(value)?);
    }
    // The temporary vector is empty before its lease is dropped.  Its capacity
    // reservation is bounded by arity and does not retain range data.
    evaluator.make_array(shape, output, output_reservation, None)
}

fn map_search_arguments<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    node: super::super::super::Node<'expr>,
    name: &str,
    shape: Shape,
    arguments: &[RuntimeValue<'expr>],
    scalar_indexes: &[usize],
) -> EvaluationResult<RuntimeValue<'expr>>
where
    R: Resolver + ?Sized,
{
    if arguments.len() > 4 {
        return Err(EvaluationFailure::InvalidExpression(
            "lookup search arity exceeds the fixed argument view",
        ));
    }
    let cells =
        shape
            .cell_count()
            .ok_or(EvaluationFailure::ResourceLimit(evaluator.local_limit(
                litchi_core::Resource::Objects,
                u64::MAX,
                evaluator.limits.max_array_cells,
            )))?;
    let output_reservation;
    let mut output;
    (output, output_reservation) = evaluator.new_element_vec(cells)?;

    // Materialize only the scalar parameter elements for one output.  Data
    // values stay in `arguments` and are borrowed by `search::apply_borrowed`.
    let mut scalar_values_reservation = None;
    let mut scalar_values = Vec::new();
    ensure_capacity(
        &mut scalar_values,
        &mut scalar_values_reservation,
        scalar_indexes.len(),
        evaluator.limits.scalar.max_stack_entries,
        evaluator.execution,
        &evaluator.storage_budget,
        "formula lookup search scalar arguments",
    )?;
    let mut scalar_by_index = [None; 4];
    for index in 0..cells {
        evaluator.charge_cell_work(index)?;
        scalar_values.clear();
        scalar_by_index[..arguments.len()].fill(None);
        for &argument_index in scalar_indexes {
            let argument =
                arguments
                    .get(argument_index)
                    .ok_or(EvaluationFailure::InvalidExpression(
                        "lookup scalar argument is missing",
                    ))?;
            let value = select_search_argument(evaluator, argument, shape, index)?;
            let scalar_index = scalar_values.len();
            scalar_values.push(value);
            scalar_by_index[argument_index] = Some(scalar_index);
        }

        // Keep this borrowed view scoped to the current output cell.  The
        // view contains references into `scalar_values`; ending its lifetime
        // here lets the next iteration reuse that bounded scratch storage.
        {
            // Lookup arity is at most four.  A stack view avoids a per-cell
            // allocation and lease while still borrowing scalar projections
            // for this one invocation only.
            let first = arguments
                .first()
                .ok_or(EvaluationFailure::InvalidExpression(
                    "lookup search arguments are missing",
                ))?;
            let mut refs = [first; 4];
            for (argument_index, argument) in arguments.iter().enumerate() {
                if let Some(scalar_index) = scalar_by_index[argument_index] {
                    refs[argument_index] = &scalar_values[scalar_index];
                } else {
                    refs[argument_index] = argument;
                }
            }
            let value =
                super::search::apply_borrowed(evaluator, node, name, &refs[..arguments.len()])?;
            output.push(evaluator.runtime_to_element(value)?);
        }
    }
    evaluator.make_array(shape, output, output_reservation, None)
}

/// Select one scalar lookup argument without reading a referenced cell.  The
/// search kernel resolves `ScalarCell` only after its literal selector checks
/// pass, so a bad index/mode cannot force a key read merely because the key
/// contributes the lifted output shape.
fn select_search_argument<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    value: &RuntimeValue<'expr>,
    output: Shape,
    index: usize,
) -> EvaluationResult<RuntimeValue<'expr>>
where
    R: Resolver + ?Sized,
{
    if let RuntimeValue::Areas(areas) = value
        && !areas.is_list
        && areas.areas.len() == 1
    {
        let area = areas
            .areas
            .first()
            .copied()
            .ok_or(EvaluationFailure::InvalidExpression(
                "lookup scalar reference disappeared",
            ))?;
        let source = Shape::new(area.rect.rows(), area.rect.columns())?;
        let Some(source_index) = super::super::projected_array_index(source, output, index) else {
            return Ok(RuntimeValue::Scalar(WorkingValue::Error(
                ScalarError::NotAvailable,
            )));
        };
        let row = source_index / source.columns();
        let column = source_index % source.columns();
        let row =
            area.rect
                .row_start
                .checked_add(row)
                .ok_or(EvaluationFailure::InvalidExpression(
                    "lookup scalar reference row overflows",
                ))?;
        let column = area.rect.column_start.checked_add(column).ok_or(
            EvaluationFailure::InvalidExpression("lookup scalar reference column overflows"),
        )?;
        return Ok(RuntimeValue::ScalarCell(super::super::RuntimeArea {
            sheet: area.sheet,
            sheet_index: area.sheet_index,
            rect: super::super::Rect::cell(row, column)?,
        }));
    }
    if let RuntimeValue::ScalarCell(area) = value {
        return Ok(RuntimeValue::ScalarCell(*area));
    }
    let element = evaluator.select_matrix_value(value, output, index)?;
    evaluator.element_to_runtime(element)
}

fn formula_error<'a>(error: ScalarError) -> RuntimeValue<'a> {
    RuntimeValue::Scalar(WorkingValue::Error(error))
}
