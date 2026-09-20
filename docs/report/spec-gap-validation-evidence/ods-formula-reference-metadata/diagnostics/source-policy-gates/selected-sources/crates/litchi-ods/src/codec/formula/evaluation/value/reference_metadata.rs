//! Reference geometry and workbook metadata functions for the value VM.
//!
//! These functions consume the retained reference descriptors.  They never
//! dereference a cell: dimensions, sheet order, and reference identity are
//! all available before a provider value is read.

use super::super::reference_metadata::Function;
use super::{
    EvaluationFailure, EvaluationResult, Mode, Resolver, RuntimeArea, RuntimeAreaSet,
    RuntimeElement, RuntimeValue, ScalarError, Shape, ValueEvaluator, WorkingValue,
    map_execution_error,
};

#[derive(Clone, Copy)]
enum Axis {
    Row,
    Column,
}

pub(super) fn is_reference_metadata_function(name: &str) -> bool {
    Function::from_name(name).is_some()
}

/// Classify metadata calls whose scalar result is invariant under a projected
/// matrix demand.  The demand cache stores scalar payloads only, so generated
/// ROW/COLUMN arrays and SHEET text-array projections remain position-sensitive.
pub(super) fn cacheable_branch<'expr, R: Resolver + ?Sized>(
    evaluator: &mut ValueEvaluator<'expr, '_, '_, '_, R>,
    node: super::super::Node<'expr>,
    function: Function,
) -> EvaluationResult<bool> {
    match function {
        Function::Column | Function::Row => Ok(false),
        Function::Sheet => {
            let Some(mut argument) = node.child(0) else {
                return Ok(false);
            };
            while matches!(argument.kind(), super::super::Kind::Parenthesized) {
                argument = argument
                    .child(0)
                    .ok_or(EvaluationFailure::InvalidExpression(
                        "metadata cache parentheses are empty",
                    ))?;
            }
            Ok(match argument.kind() {
                super::super::Kind::Number
                | super::super::Kind::String
                | super::super::Kind::Error => true,
                super::super::Kind::Reference(reference) => !matches!(
                    reference,
                    crate::codec::formula::reference::Reference::Source { .. }
                ),
                _ => false,
            })
        },
        Function::Sheets => match node.child(0) {
            None => Ok(true),
            Some(argument) => static_argument(evaluator, argument, true, false),
        },
        Function::Areas | Function::Columns | Function::Rows => {
            let Some(argument) = node.child(0) else {
                return Ok(false);
            };
            static_argument(evaluator, argument, true, false)
        },
        Function::IsRef => {
            let Some(mut argument) = node.child(0) else {
                return Ok(false);
            };
            while matches!(argument.kind(), super::super::Kind::Parenthesized) {
                argument = argument
                    .child(0)
                    .ok_or(EvaluationFailure::InvalidExpression(
                        "metadata cache parentheses are empty",
                    ))?;
            }
            // ISREF can classify a source-qualified descriptor without
            // resolving it, so direct references are invariant even though
            // the generic matrix cache classifier conservatively rejects them.
            if matches!(argument.kind(), super::super::Kind::Reference(_)) {
                return Ok(true);
            }
            static_argument(evaluator, argument, true, false)
        },
    }
}

/// Metadata cache classification is deliberately narrower than the generic
/// matrix-branch classifier.  A computed IF/function can depend on a cell
/// value even when its reference geometry is fixed; only literal values,
/// literal arrays, and reference-operator trees are admitted here.  The walk
/// is iterative so a deeply parenthesized expression remains bounded by the
/// evaluator's AST stack limit.
fn static_argument<'expr, R: Resolver + ?Sized>(
    evaluator: &mut ValueEvaluator<'expr, '_, '_, '_, R>,
    root: super::super::Node<'expr>,
    allow_array: bool,
    allow_source: bool,
) -> EvaluationResult<bool> {
    let mut reservation = None;
    // Keep the lease before the vector so an early error drops the vector
    // before releasing its charged capacity.
    let mut nodes = Vec::new();
    super::ensure_capacity(
        &mut nodes,
        &mut reservation,
        1,
        evaluator.limits.scalar.max_stack_entries,
        evaluator.execution,
        &evaluator.storage_budget,
        "reference metadata cache syntax frames",
    )?;
    nodes.push(root);
    while let Some(node) = nodes.pop() {
        evaluator.scalar.charge_work(1)?;
        match node.kind() {
            super::super::Kind::Number
            | super::super::Kind::String
            | super::super::Kind::Error
            | super::super::Kind::Missing => {},
            super::super::Kind::Reference(reference) => {
                if matches!(
                    reference,
                    crate::codec::formula::reference::Reference::Source { .. }
                ) && !allow_source
                {
                    return Ok(false);
                }
            },
            super::super::Kind::Parenthesized => {
                let Some(child) = node.child(0) else {
                    return Ok(false);
                };
                super::ensure_capacity(
                    &mut nodes,
                    &mut reservation,
                    1,
                    evaluator.limits.scalar.max_stack_entries,
                    evaluator.execution,
                    &evaluator.storage_budget,
                    "reference metadata cache syntax frames",
                )?;
                nodes.push(child);
            },
            super::super::Kind::Infix(
                super::super::InfixOperator::Range
                | super::super::InfixOperator::Intersection
                | super::super::InfixOperator::Union,
            ) => {
                let (Some(left), Some(right)) = (node.child(0), node.child(1)) else {
                    return Ok(false);
                };
                super::ensure_capacity(
                    &mut nodes,
                    &mut reservation,
                    2,
                    evaluator.limits.scalar.max_stack_entries,
                    evaluator.execution,
                    &evaluator.storage_budget,
                    "reference metadata cache syntax frames",
                )?;
                nodes.push(right);
                nodes.push(left);
            },
            super::super::Kind::Array(_) | super::super::Kind::ArrayRow if allow_array => {
                let count = node.child_count();
                super::ensure_capacity(
                    &mut nodes,
                    &mut reservation,
                    count,
                    evaluator.limits.scalar.max_stack_entries,
                    evaluator.execution,
                    &evaluator.storage_budget,
                    "reference metadata cache syntax frames",
                )?;
                for index in (0..count).rev() {
                    let Some(child) = node.child(index) else {
                        return Ok(false);
                    };
                    nodes.push(child);
                }
            },
            _ => return Ok(false),
        }
    }
    Ok(true)
}

pub(super) fn apply<'expr, R: Resolver + ?Sized>(
    evaluator: &mut ValueEvaluator<'expr, '_, '_, '_, R>,
    name: &str,
    arguments: Vec<RuntimeValue<'expr>>,
) -> EvaluationResult<RuntimeValue<'expr>> {
    let function = Function::from_name(name).ok_or(EvaluationFailure::InvalidExpression(
        "unknown reference metadata function",
    ))?;
    if !function.valid_arity(arguments.len()) {
        return Ok(formula_error(ScalarError::Value));
    }
    evaluator.scalar.charge_work(1)?;

    let mut arguments = arguments.into_iter();
    match function {
        Function::Areas => apply_areas(arguments.next()),
        Function::Column => apply_axis(evaluator, arguments.next(), Axis::Column),
        Function::Columns => apply_dimension(arguments.next(), Axis::Column),
        Function::IsRef => apply_is_ref(arguments.next()),
        Function::Row => apply_axis(evaluator, arguments.next(), Axis::Row),
        Function::Rows => apply_dimension(arguments.next(), Axis::Row),
        Function::Sheet => apply_sheet(evaluator, arguments.next()),
        Function::Sheets => apply_sheets(evaluator, arguments.next()),
    }
}

fn formula_error(error: ScalarError) -> RuntimeValue<'static> {
    RuntimeValue::Scalar(WorkingValue::Error(error))
}

fn propagated_error<'expr>(value: &RuntimeValue<'expr>) -> Option<RuntimeValue<'expr>> {
    match value {
        RuntimeValue::Scalar(WorkingValue::Error(error)) => Some(formula_error(*error)),
        _ => None,
    }
}

fn apply_is_ref<'expr>(
    argument: Option<RuntimeValue<'expr>>,
) -> EvaluationResult<RuntimeValue<'expr>> {
    if matches!(argument, Some(RuntimeValue::Missing)) {
        return Ok(formula_error(ScalarError::Value));
    }
    let is_reference = matches!(
        argument,
        Some(RuntimeValue::Areas(_) | RuntimeValue::SourceReference | RuntimeValue::ScalarCell(_),)
    );
    Ok(RuntimeValue::Scalar(WorkingValue::Logical(is_reference)))
}

fn apply_areas<'expr>(
    argument: Option<RuntimeValue<'expr>>,
) -> EvaluationResult<RuntimeValue<'expr>> {
    let Some(argument) = argument else {
        return Ok(formula_error(ScalarError::Value));
    };
    if let Some(error) = propagated_error(&argument) {
        return Ok(error);
    }
    let count = match argument {
        RuntimeValue::Areas(areas) => areas.records.len(),
        RuntimeValue::ScalarCell(_) => 1,
        RuntimeValue::Empty
        | RuntimeValue::Missing
        | RuntimeValue::Scalar(_)
        | RuntimeValue::SourceReference
        | RuntimeValue::Array(_) => return Ok(formula_error(ScalarError::Value)),
    };
    Ok(RuntimeValue::Scalar(WorkingValue::Number(count as f64)))
}

fn apply_dimension<'expr>(
    argument: Option<RuntimeValue<'expr>>,
    axis: Axis,
) -> EvaluationResult<RuntimeValue<'expr>> {
    let Some(argument) = argument else {
        return Ok(formula_error(ScalarError::Value));
    };
    if let Some(error) = propagated_error(&argument) {
        return Ok(error);
    }
    let count = match argument {
        RuntimeValue::Array(array) => match axis {
            Axis::Row => array.shape.rows(),
            Axis::Column => array.shape.columns(),
        },
        RuntimeValue::Areas(areas) => {
            let Some(area) = single_reference_area(&areas) else {
                return Ok(formula_error(ScalarError::Value));
            };
            match axis {
                Axis::Row => area.rect.rows(),
                Axis::Column => area.rect.columns(),
            }
        },
        RuntimeValue::ScalarCell(_) => 1,
        RuntimeValue::Empty
        | RuntimeValue::Missing
        | RuntimeValue::Scalar(_)
        | RuntimeValue::SourceReference => {
            return Ok(formula_error(ScalarError::Value));
        },
    };
    Ok(RuntimeValue::Scalar(WorkingValue::Number(count as f64)))
}

fn apply_axis<'expr, R: Resolver + ?Sized>(
    evaluator: &mut ValueEvaluator<'expr, '_, '_, '_, R>,
    argument: Option<RuntimeValue<'expr>>,
    axis: Axis,
) -> EvaluationResult<RuntimeValue<'expr>> {
    let Some(argument) = argument else {
        return current_axis_value(evaluator, axis);
    };
    // The optional argument has a default only when the AST has no child.
    // An explicit empty slot is an invalid argument in this profile.
    if matches!(argument, RuntimeValue::Missing) {
        return Ok(formula_error(ScalarError::Value));
    }
    if let Some(error) = propagated_error(&argument) {
        return Ok(error);
    }
    let area = match argument {
        RuntimeValue::Areas(areas) => {
            let Some(area) = single_reference_area(&areas) else {
                return Ok(formula_error(ScalarError::Value));
            };
            area
        },
        RuntimeValue::ScalarCell(area) => area,
        RuntimeValue::Empty
        | RuntimeValue::Scalar(_)
        | RuntimeValue::Array(_)
        | RuntimeValue::SourceReference => {
            return Ok(formula_error(ScalarError::Value));
        },
        RuntimeValue::Missing => return current_axis_value(evaluator, axis),
    };
    let count = match axis {
        Axis::Row => area.rect.rows(),
        Axis::Column => area.rect.columns(),
    };
    if count == 1 {
        return axis_number(axis, area, 0)
            .map(|value| RuntimeValue::Scalar(WorkingValue::Number(value)));
    }
    let shape = match axis {
        Axis::Row => Shape::new(count, 1)?,
        Axis::Column => Shape::new(1, count)?,
    };
    if let Some((output, index)) = evaluator.projection {
        let Some(source_index) = super::projected_array_index(shape, output, index) else {
            return Ok(formula_error(ScalarError::NotAvailable));
        };
        evaluator.charge_cell_work(source_index)?;
        let value = axis_number(axis, area, source_index)?;
        return Ok(RuntimeValue::Scalar(WorkingValue::Number(value)));
    }
    if evaluator.mode == Mode::Scalar {
        // Scalar demand projects generated metadata arrays to their first
        // element, matching the ordinary scalar profile.  A missing argument
        // above is the only form that consults the current formula position.
        let value = axis_number(axis, area, 0)?;
        return Ok(RuntimeValue::Scalar(WorkingValue::Number(value)));
    }
    let cells = shape
        .cell_count()
        .ok_or(EvaluationFailure::InvalidExpression(
            "reference metadata output count overflow",
        ))?;
    let (output, reservation) = evaluator.new_element_vec(cells)?;
    let mut output = output;
    for index in 0..cells {
        evaluator.charge_cell_work(index)?;
        output.push(RuntimeElement::Present(WorkingValue::Number(axis_number(
            axis, area, index,
        )?)));
    }
    evaluator.make_array(shape, output, reservation, None)
}

fn current_axis_value<'expr, R: Resolver + ?Sized>(
    evaluator: &ValueEvaluator<'expr, '_, '_, '_, R>,
    axis: Axis,
) -> EvaluationResult<RuntimeValue<'expr>> {
    let coordinate = match axis {
        Axis::Row => evaluator.formula_position.row,
        Axis::Column => evaluator.formula_position.column,
    };
    let value = coordinate
        .checked_add(1)
        .ok_or(EvaluationFailure::InvalidExpression(
            "current reference coordinate overflows",
        ))?;
    Ok(RuntimeValue::Scalar(WorkingValue::Number(value as f64)))
}

fn axis_number(axis: Axis, area: RuntimeArea<'_>, offset: usize) -> EvaluationResult<f64> {
    let coordinate =
        match axis {
            Axis::Row => area.rect.row_start.checked_add(offset).ok_or(
                EvaluationFailure::InvalidExpression("reference row coordinate overflows"),
            )?,
            Axis::Column => area.rect.column_start.checked_add(offset).ok_or(
                EvaluationFailure::InvalidExpression("reference column coordinate overflows"),
            )?,
        };
    let value = coordinate
        .checked_add(1)
        .ok_or(EvaluationFailure::InvalidExpression(
            "reference coordinate overflows",
        ))?;
    Ok(value as f64)
}

fn single_reference_area<'a>(areas: &RuntimeAreaSet<'a>) -> Option<RuntimeArea<'a>> {
    if areas.is_list || areas.records.len() != 1 {
        return None;
    }
    areas.areas.first().copied()
}

fn apply_sheet<'expr, R: Resolver + ?Sized>(
    evaluator: &mut ValueEvaluator<'expr, '_, '_, '_, R>,
    argument: Option<RuntimeValue<'expr>>,
) -> EvaluationResult<RuntimeValue<'expr>> {
    let Some(argument) = argument else {
        return sheet_number(evaluator, evaluator.formula_position.sheet);
    };
    if matches!(argument, RuntimeValue::Missing) {
        return Ok(formula_error(ScalarError::Value));
    }
    if let Some(error) = propagated_error(&argument) {
        return Ok(error);
    }
    match argument {
        RuntimeValue::Areas(areas) => {
            let Some(area) = single_reference_area(&areas) else {
                return Ok(formula_error(ScalarError::Value));
            };
            checked_sheet_number(area.sheet_index)
        },
        RuntimeValue::ScalarCell(area) => checked_sheet_number(area.sheet_index),
        RuntimeValue::SourceReference => Ok(formula_error(ScalarError::Value)),
        RuntimeValue::Scalar(value) => apply_sheet_scalar(evaluator, value),
        RuntimeValue::Array(array) => apply_sheet_array(evaluator, array),
        RuntimeValue::Empty => Ok(formula_error(ScalarError::Value)),
        RuntimeValue::Missing => unreachable!("missing SHEET argument handled above"),
    }
}

fn apply_sheet_scalar<'expr, R: Resolver + ?Sized>(
    evaluator: &mut ValueEvaluator<'expr, '_, '_, '_, R>,
    value: WorkingValue<'expr>,
) -> EvaluationResult<RuntimeValue<'expr>> {
    match value {
        WorkingValue::Text(text) => {
            evaluator.scalar.charge_bytes(text.text.len())?;
            sheet_number(evaluator, text.text.as_ref())
        },
        WorkingValue::Number(_) | WorkingValue::Logical(_) => {
            let text = super::super::to_text(value, &mut evaluator.scalar)?;
            let result = sheet_number(evaluator, text.text.as_ref());
            drop(text);
            result
        },
        WorkingValue::Error(error) => Ok(formula_error(error)),
        WorkingValue::Complex(_) => Ok(formula_error(ScalarError::Value)),
    }
}

fn apply_sheet_element<'expr, R: Resolver + ?Sized>(
    evaluator: &mut ValueEvaluator<'expr, '_, '_, '_, R>,
    element: RuntimeElement<'expr>,
) -> EvaluationResult<RuntimeElement<'expr>> {
    match element {
        RuntimeElement::Present(value) => {
            let value = apply_sheet_scalar(evaluator, value)?;
            evaluator.runtime_to_element(value)
        },
        RuntimeElement::Empty | RuntimeElement::Missing => Ok(RuntimeElement::Present(
            WorkingValue::Error(ScalarError::Value),
        )),
    }
}

fn apply_sheet_array<'expr, R: Resolver + ?Sized>(
    evaluator: &mut ValueEvaluator<'expr, '_, '_, '_, R>,
    array: super::RuntimeArrayValue<'expr>,
) -> EvaluationResult<RuntimeValue<'expr>> {
    let value = RuntimeValue::Array(array);
    if let Some((output, index)) = evaluator.projection {
        let element = evaluator.select_matrix_value(&value, output, index)?;
        return apply_sheet_element(evaluator, element).map(|element| {
            RuntimeValue::Scalar(match element {
                RuntimeElement::Present(value) => value,
                RuntimeElement::Empty | RuntimeElement::Missing => {
                    WorkingValue::Error(ScalarError::Value)
                },
            })
        });
    }
    let RuntimeValue::Array(array) = value else {
        unreachable!("SHEET array value changed variant");
    };
    if evaluator.mode == Mode::Scalar {
        let element = array
            .cells
            .into_iter()
            .next()
            .unwrap_or(RuntimeElement::Missing);
        let element = apply_sheet_element(evaluator, element)?;
        return Ok(RuntimeValue::Scalar(match element {
            RuntimeElement::Present(value) => value,
            RuntimeElement::Empty | RuntimeElement::Missing => {
                WorkingValue::Error(ScalarError::Value)
            },
        }));
    }
    let shape = array.shape;
    let cells = array.cells;
    let count = shape
        .cell_count()
        .ok_or(EvaluationFailure::InvalidExpression(
            "SHEET array output count overflows",
        ))?;
    let (output, reservation) = evaluator.new_element_vec(count)?;
    let mut output = output;
    for (index, element) in cells.into_iter().enumerate() {
        evaluator.charge_cell_work(index)?;
        output.push(apply_sheet_element(evaluator, element)?);
    }
    evaluator.make_array(shape, output, reservation, None)
}

fn sheet_number<'expr, R: Resolver + ?Sized>(
    evaluator: &mut ValueEvaluator<'expr, '_, '_, '_, R>,
    sheet: &str,
) -> EvaluationResult<RuntimeValue<'expr>> {
    evaluator.scalar.charge_work(1)?;
    let index = evaluator.resolver.sheet_index(sheet, evaluator.execution)?;
    evaluator.execution.check().map_err(map_execution_error)?;
    match index {
        Some(index) => checked_sheet_number(index),
        None => Ok(formula_error(ScalarError::Reference)),
    }
}

fn checked_sheet_number(index: usize) -> EvaluationResult<RuntimeValue<'static>> {
    let value = index
        .checked_add(1)
        .ok_or(EvaluationFailure::InvalidExpression(
            "sheet number overflows",
        ))?;
    Ok(RuntimeValue::Scalar(WorkingValue::Number(value as f64)))
}

fn apply_sheets<'expr, R: Resolver + ?Sized>(
    evaluator: &mut ValueEvaluator<'expr, '_, '_, '_, R>,
    argument: Option<RuntimeValue<'expr>>,
) -> EvaluationResult<RuntimeValue<'expr>> {
    let Some(argument) = argument else {
        return workbook_sheet_count(evaluator);
    };
    if matches!(argument, RuntimeValue::Missing) {
        return Ok(formula_error(ScalarError::Value));
    }
    if let Some(error) = propagated_error(&argument) {
        return Ok(error);
    }
    match argument {
        RuntimeValue::Areas(areas) => {
            let Some(area) = single_reference_area(&areas) else {
                return Ok(formula_error(ScalarError::Value));
            };
            let count = areas.areas.len();
            if count == 0 {
                return Ok(formula_error(ScalarError::Reference));
            }
            // A direct 3-D reference retains one physical area per sheet
            // plane.  It is one logical reference, but SHEETS counts those
            // ordered planes rather than retained reference records.
            let _ = area;
            Ok(RuntimeValue::Scalar(WorkingValue::Number(count as f64)))
        },
        RuntimeValue::ScalarCell(_) => Ok(RuntimeValue::Scalar(WorkingValue::Number(1.0))),
        RuntimeValue::Array(_)
        | RuntimeValue::Empty
        | RuntimeValue::Missing
        | RuntimeValue::Scalar(_)
        | RuntimeValue::SourceReference => Ok(formula_error(ScalarError::Value)),
    }
}

fn workbook_sheet_count<'expr, R: Resolver + ?Sized>(
    evaluator: &mut ValueEvaluator<'expr, '_, '_, '_, R>,
) -> EvaluationResult<RuntimeValue<'expr>> {
    evaluator.scalar.charge_work(1)?;
    let count = evaluator.resolver.sheet_count(evaluator.execution)?;
    evaluator.execution.check().map_err(map_execution_error)?;
    if count == 0 {
        return Ok(formula_error(ScalarError::Reference));
    }
    Ok(RuntimeValue::Scalar(WorkingValue::Number(count as f64)))
}

pub(super) fn shape_base<'expr, R: Resolver + ?Sized>(
    evaluator: &mut ValueEvaluator<'expr, '_, '_, '_, R>,
    node: super::super::Node<'expr>,
    function: Function,
) -> EvaluationResult<Option<Shape>> {
    if function == Function::Sheet {
        let Some(mut argument) = node.child(0) else {
            return Ok(Some(Shape::new(1, 1)?));
        };
        while matches!(argument.kind(), super::super::Kind::Parenthesized) {
            evaluator.scalar.charge_work(1)?;
            argument = argument
                .child(0)
                .ok_or(EvaluationFailure::InvalidExpression(
                    "metadata shape parentheses are empty",
                ))?;
        }
        if let super::super::Kind::Function { name } = argument.kind()
            && ValueEvaluator::<R>::is_reference_value_handler(name)
            && evaluator.reference_shape_value(argument)?.is_some()
        {
            // A selected reference remains one metadata argument, even when
            // its cell rectangle has several rows or columns.
            return Ok(Some(Shape::new(1, 1)?));
        }
        return Ok(match argument.kind() {
            super::super::Kind::Array(dimensions) => dimensions
                .columns()
                .map(|columns| Shape::new(dimensions.rows(), columns))
                .transpose()?,
            super::super::Kind::Number
            | super::super::Kind::String
            | super::super::Kind::Error
            | super::super::Kind::Missing
            | super::super::Kind::Reference(_)
            | super::super::Kind::AutomaticIntersection
            | super::super::Kind::Infix(
                super::super::InfixOperator::Range
                | super::super::InfixOperator::Intersection
                | super::super::InfixOperator::Union,
            ) => Some(Shape::new(1, 1)?),
            // Let the parent shape planner visit a computed scalar/array
            // expression.  Its result shape must be retained for matrix
            // SHEET lifting (for example SHEET(ABS(A1:A2))).
            _ => None,
        });
    }
    if !matches!(function, Function::Row | Function::Column) {
        return Ok(Some(Shape::new(1, 1)?));
    }
    let Some(argument) = node.child(0) else {
        return Ok(Some(Shape::new(1, 1)?));
    };
    let Some(value) = evaluator.reference_shape_value(argument)? else {
        return Ok(Some(Shape::new(1, 1)?));
    };
    let shape = match value {
        RuntimeValue::Areas(areas) => match single_reference_area(&areas) {
            Some(area) => match function {
                Function::Row => Shape::new(area.rect.rows(), 1)?,
                Function::Column => Shape::new(1, area.rect.columns())?,
                _ => Shape::new(1, 1)?,
            },
            None => Shape::new(1, 1)?,
        },
        RuntimeValue::ScalarCell(_) => Shape::new(1, 1)?,
        RuntimeValue::Array(_)
        | RuntimeValue::Empty
        | RuntimeValue::Missing
        | RuntimeValue::Scalar(_)
        | RuntimeValue::SourceReference => Shape::new(1, 1)?,
    };
    Ok(Some(shape))
}

pub(super) fn reference_kind(
    function: Function,
    argument: Option<super::ReferenceOperandKind>,
) -> super::ReferenceOperandKind {
    match function {
        Function::Row | Function::Column => match argument {
            Some(super::ReferenceOperandKind::Reference)
            | Some(super::ReferenceOperandKind::ReferenceList)
            | Some(super::ReferenceOperandKind::Array) => super::ReferenceOperandKind::Array,
            Some(super::ReferenceOperandKind::Error(error)) => {
                super::ReferenceOperandKind::Error(error)
            },
            _ => super::ReferenceOperandKind::Scalar,
        },
        Function::Areas
        | Function::Columns
        | Function::IsRef
        | Function::Rows
        | Function::Sheet
        | Function::Sheets => super::ReferenceOperandKind::Scalar,
    }
}
