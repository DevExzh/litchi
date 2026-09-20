//! Raw-value inspection and streamed elementwise conversion.

use super::super::inspection::{Function, Input, apply_inputs};
use super::{
    EvaluationFailure, EvaluationResult, Mode, Resolver, RuntimeElement, RuntimeValue, ScalarError,
    Shape, ValueEvaluator, WorkingValue, broadcast_shape,
};

/// Consume a scalar slot without erasing Empty or an omitted argument.
fn input(element: RuntimeElement<'_>) -> Input<'_> {
    match element {
        RuntimeElement::Empty => Input::Empty,
        RuntimeElement::Missing => Input::Value(WorkingValue::Error(ScalarError::NotAvailable)),
        RuntimeElement::Present(value) => Input::Value(value),
    }
}

pub(super) fn apply<'expr, R: Resolver + ?Sized>(
    evaluator: &mut ValueEvaluator<'expr, '_, '_, '_, R>,
    name: &str,
    mut arguments: Vec<RuntimeValue<'expr>>,
) -> EvaluationResult<RuntimeValue<'expr>> {
    let function = Function::from_name(name).ok_or(EvaluationFailure::InvalidExpression(
        "unknown inspection function",
    ))?;
    if arguments.len() > 3 || !function.valid_arity(arguments.len()) {
        return Ok(RuntimeValue::Scalar(WorkingValue::Error(
            ScalarError::Value,
        )));
    }
    let invalid_separators = if name.eq_ignore_ascii_case("NUMBERVALUE")
        && matches!(
            arguments.first(),
            Some(RuntimeValue::Areas(_) | RuntimeValue::ScalarCell(_))
        )
        && let (Some(decimal), Some(group)) = (
            known_separator(arguments.get(1), "."),
            known_separator(arguments.get(2), ","),
        ) {
        evaluator.scalar.charge_bytes(decimal.len())?;
        evaluator.scalar.charge_bytes(group.len())?;
        !super::super::inspection::valid_number_value_separators(decimal, group)
    } else {
        false
    };
    // Reference lists cannot be iterated as rectangular scalar arguments.
    // Refuse the complete descriptor before reading any argument cell.
    if arguments
        .iter()
        .any(|argument| matches!(argument, RuntimeValue::Areas(areas) if areas.is_list))
    {
        return Ok(RuntimeValue::Scalar(WorkingValue::Error(
            ScalarError::Value,
        )));
    }
    if name.eq_ignore_ascii_case("TYPE") {
        return apply_type(evaluator, function, arguments);
    }
    // A cuboid remains one Reference. Matrix consumers select the caller's
    // sheet plane, while scalar consumers use ordinary 3-D intersection.
    if !name.eq_ignore_ascii_case("N")
        && (evaluator.mode == Mode::Matrix || evaluator.projection.is_some())
    {
        for argument in &mut arguments {
            if let RuntimeValue::Areas(areas) = argument
                && areas.areas.len() > 1
            {
                let sheet = evaluator
                    .resolver
                    .sheet_index(evaluator.position.sheet, evaluator.execution)?;
                evaluator
                    .scalar
                    .charge_work(u64::try_from(areas.areas.len()).unwrap_or(u64::MAX))?;
                areas.areas.retain(|area| Some(area.sheet_index) == sheet);
                if areas.areas.is_empty() {
                    *argument =
                        RuntimeValue::Scalar(WorkingValue::Error(ScalarError::NotAvailable));
                }
            }
        }
    }
    if invalid_separators {
        return separator_error(evaluator, &arguments);
    }
    apply_elementwise(
        evaluator,
        function,
        arguments,
        name.eq_ignore_ascii_case("N"),
    )
}

fn separator_error<'expr, R: Resolver + ?Sized>(
    evaluator: &mut ValueEvaluator<'expr, '_, '_, '_, R>,
    arguments: &[RuntimeValue<'expr>],
) -> EvaluationResult<RuntimeValue<'expr>> {
    if evaluator.mode != Mode::Matrix
        || evaluator.projection.is_some()
        || !arguments.iter().any(RuntimeValue::is_array_like)
    {
        return Ok(RuntimeValue::Scalar(WorkingValue::Error(
            ScalarError::Value,
        )));
    }
    let mut shape = Shape::new(1, 1)?;
    for argument in arguments {
        shape = broadcast_shape(shape, evaluator.runtime_shape(argument)?).ok_or(
            EvaluationFailure::InvalidExpression("incompatible inspection error shapes"),
        )?;
    }
    let cells = shape
        .cell_count()
        .ok_or(EvaluationFailure::InvalidExpression(
            "inspection error cell count overflow",
        ))?;
    let reservation;
    let mut output;
    (output, reservation) = evaluator.new_element_vec(cells)?;
    for index in 0..cells {
        evaluator.charge_cell_work(index)?;
        output.push(RuntimeElement::Present(WorkingValue::Error(
            ScalarError::Value,
        )));
    }
    evaluator.make_array(shape, output, reservation, None)
}

fn known_separator<'a>(
    argument: Option<&'a RuntimeValue<'_>>,
    default: &'a str,
) -> Option<&'a str> {
    match argument {
        None | Some(RuntimeValue::Missing) => Some(default),
        Some(RuntimeValue::Scalar(WorkingValue::Text(text))) => Some(text.text.as_ref()),
        _ => None,
    }
}

/// Supplement ordinary shape planning for direct cuboid arguments. Generic
/// scalar functions do not expose a 3-D matrix shape; this family's explicit
/// current-sheet profile does, without reading any reference cell.
pub(super) fn shape_base<'expr, R: Resolver + ?Sized>(
    evaluator: &mut ValueEvaluator<'expr, '_, '_, '_, R>,
    node: super::super::Node<'expr>,
) -> EvaluationResult<Option<Shape>> {
    let mut base = None;
    for index in 0..node.child_count() {
        evaluator.scalar.charge_work(1)?;
        let Some(mut child) = node.child(index) else {
            continue;
        };
        while matches!(child.kind(), super::super::Kind::Parenthesized) {
            evaluator.scalar.charge_work(1)?;
            child = child.child(0).ok_or(EvaluationFailure::InvalidExpression(
                "inspection shape parentheses are empty",
            ))?;
        }
        match child.kind() {
            super::super::Kind::Reference(reference) => {
                if evaluator.reference_shape_hint(reference)?.is_some() {
                    continue;
                }
            },
            super::super::Kind::Infix(
                super::super::InfixOperator::Range
                | super::super::InfixOperator::Intersection
                | super::super::InfixOperator::Union,
            ) => {},
            _ => continue,
        }
        let Some(RuntimeValue::Areas(areas)) = evaluator.reference_shape_value(child)? else {
            continue;
        };
        if areas.is_list || areas.areas.len() < 2 {
            continue;
        }
        let sheet = evaluator
            .resolver
            .sheet_index(evaluator.position.sheet, evaluator.execution)?;
        evaluator
            .scalar
            .charge_work(u64::try_from(areas.areas.len()).unwrap_or(u64::MAX))?;
        let shape = match areas
            .areas
            .iter()
            .find(|area| Some(area.sheet_index) == sheet)
        {
            Some(area) => Shape::new(area.rect.rows(), area.rect.columns())?,
            None => Shape::new(1, 1)?,
        };
        base = Some(match base {
            None => shape,
            Some(previous) => broadcast_shape(previous, shape).ok_or(
                EvaluationFailure::InvalidExpression("incompatible inspection cuboid shapes"),
            )?,
        });
    }
    Ok(base)
}

fn apply_type<'expr, R: Resolver + ?Sized>(
    evaluator: &mut ValueEvaluator<'expr, '_, '_, '_, R>,
    function: Function,
    arguments: Vec<RuntimeValue<'expr>>,
) -> EvaluationResult<RuntimeValue<'expr>> {
    let argument = arguments
        .into_iter()
        .next()
        .ok_or(EvaluationFailure::InvalidExpression(
            "TYPE argument is missing",
        ))?;
    match argument {
        RuntimeValue::Array(_) => Ok(RuntimeValue::Scalar(WorkingValue::Number(64.0))),
        RuntimeValue::Areas(areas)
            if areas.areas.len() > 1
                || areas
                    .areas
                    .iter()
                    .any(|area| area.rect.rows() > 1 || area.rect.columns() > 1) =>
        {
            // TYPE classifies the dereferenced array, but every referenced
            // formula must still be evaluated. Discard cells immediately;
            // formula errors are array members, while typed failures bubble.
            let mut index = 0usize;
            for area in areas.areas {
                for row in area.rect.row_start..area.rect.row_end {
                    for column in area.rect.column_start..area.rect.column_end {
                        evaluator.charge_cell_work(index)?;
                        let read = evaluator.read_reference_cell(area.sheet, row, column)?;
                        let _ = evaluator.read_to_element(read)?;
                        index =
                            index
                                .checked_add(1)
                                .ok_or(EvaluationFailure::InvalidExpression(
                                    "TYPE reference cell count overflow",
                                ))?;
                    }
                }
            }
            Ok(RuntimeValue::Scalar(WorkingValue::Number(64.0)))
        },
        argument => {
            let value = match argument {
                RuntimeValue::Missing => Input::Missing,
                argument => match evaluator.value_to_slot(argument)? {
                    super::scalar::Slot::Empty => Input::Empty,
                    super::scalar::Slot::Value(value) => Input::Value(value),
                },
            };
            apply_inputs(&mut evaluator.scalar, function, [value]).map(RuntimeValue::Scalar)
        },
    }
}

fn apply_elementwise<'expr, R: Resolver + ?Sized>(
    evaluator: &mut ValueEvaluator<'expr, '_, '_, '_, R>,
    function: Function,
    arguments: Vec<RuntimeValue<'expr>>,
    scalar_result: bool,
) -> EvaluationResult<RuntimeValue<'expr>> {
    let count = arguments.len();
    if !scalar_result && let Some((shape, index)) = evaluator.projection {
        // A lazy matrix branch requests one relative output coordinate.
        // This is matrix selection, not intersection with the worksheet row
        // containing the formula (the referenced range may start elsewhere).
        evaluator.charge_cell_work(index)?;
        let mut inputs = std::array::from_fn::<_, 3, _>(|_| Input::Missing);
        for (argument_index, argument) in arguments.iter().enumerate() {
            inputs[argument_index] = if matches!(argument, RuntimeValue::Missing) {
                Input::Missing
            } else {
                input(evaluator.select_matrix_value(argument, shape, index)?)
            };
        }
        return apply_inputs(
            &mut evaluator.scalar,
            function,
            inputs.into_iter().take(count),
        )
        .map(RuntimeValue::Scalar);
    }
    if scalar_result
        || evaluator.mode != Mode::Matrix
        || !arguments.iter().any(RuntimeValue::is_array_like)
    {
        let mut inputs = std::array::from_fn::<_, 3, _>(|_| Input::Missing);
        for (index, argument) in arguments.into_iter().enumerate() {
            inputs[index] = match argument {
                RuntimeValue::Missing => Input::Missing,
                RuntimeValue::Array(array) if scalar_result => {
                    input(array.cells.into_iter().next().ok_or(
                        EvaluationFailure::InvalidExpression("N array argument is empty"),
                    )?)
                },
                argument => match evaluator.value_to_slot(argument)? {
                    super::scalar::Slot::Empty => Input::Empty,
                    super::scalar::Slot::Value(value) => Input::Value(value),
                },
            };
        }
        return apply_inputs(
            &mut evaluator.scalar,
            function,
            inputs.into_iter().take(count),
        )
        .map(RuntimeValue::Scalar);
    }

    let mut shape = Shape::new(1, 1)?;
    for argument in &arguments {
        shape = broadcast_shape(shape, evaluator.runtime_shape(argument)?).ok_or(
            EvaluationFailure::InvalidExpression("incompatible inspection array shapes"),
        )?;
    }
    let cells = shape
        .cell_count()
        .ok_or(EvaluationFailure::InvalidExpression(
            "inspection array cell count overflow",
        ))?;
    // The reservation must outlive the output allocation on every return path.
    let reservation;
    let mut output;
    (output, reservation) = evaluator.new_element_vec(cells)?;
    for index in 0..cells {
        evaluator.charge_cell_work(index)?;
        let mut inputs = std::array::from_fn::<_, 3, _>(|_| Input::Missing);
        for (argument_index, argument) in arguments.iter().enumerate() {
            inputs[argument_index] = if matches!(argument, RuntimeValue::Missing) {
                Input::Missing
            } else {
                input(evaluator.select_matrix_value(argument, shape, index)?)
            };
        }
        output.push(RuntimeElement::Present(apply_inputs(
            &mut evaluator.scalar,
            function,
            inputs.into_iter().take(count),
        )?));
    }
    evaluator.make_array(shape, output, reservation, None)
}
