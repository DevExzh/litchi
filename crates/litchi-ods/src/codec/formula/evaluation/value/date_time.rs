//! Resolver-backed date and time functions.
//!
//! Date/Offset parameters use the ordinary value VM broadcast rules.  The
//! optional Holidays and Workdays parameters of NETWORKDAYS and WORKDAY are
//! different: they are complete sequence inputs, so they are scanned through
//! their retained descriptors before a scalar result is published.

use super::super::date_time::{self, Function};
use super::{
    EvaluationFailure, EvaluationResult, Resolver, RuntimeAreaSet, RuntimeElement, RuntimeValue,
    ScalarError, Shape, ValueEvaluator, WorkingValue, broadcast_shape, element_to_runtime,
    ensure_capacity,
};
use litchi_core::Reservation;

/// Return whether `name` belongs to the date/time family.
pub(super) fn is_date_time_function(name: &str) -> bool {
    date_time::is_date_time_function(name)
}

/// Date/Offset expressions inside projected lookup arguments may still be
/// matrix-valued.  Only the holiday/workweek slots are complete sequence
/// arguments; all other slots retain their ordinary projected shape.
pub(super) fn is_sequence_reducer(name: &str) -> bool {
    matches!(
        Function::from_name(name),
        Some(Function::Networkdays | Function::Workday)
    )
}

pub(super) fn sequence_argument(name: &str, index: usize) -> bool {
    Function::from_name(name).is_some_and(|function| function.sequence_argument(index))
}

pub(super) fn complete_argument(name: &str, index: usize) -> bool {
    sequence_argument(name, index)
}

pub(super) fn matrix_argument(name: &str, index: usize, node: super::super::Node<'_>) -> bool {
    sequence_argument(name, index) && is_array_node(node)
}

pub(super) fn projected_scalar_argument<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    name: &str,
    index: usize,
    node: super::super::Node<'expr>,
) -> EvaluationResult<bool>
where
    R: Resolver + ?Sized,
{
    if !is_date_time_function(name) || sequence_argument(name, index) {
        return Ok(false);
    }
    super::lookup::projected_scalar_expression(evaluator, node)
}

pub(super) fn criterion_full_argument(
    name: &str,
    index: usize,
    _node: super::super::Node<'_>,
) -> bool {
    sequence_argument(name, index)
}

/// Date/time cache classification is intentionally conservative.  A direct
/// literal or fixed local reference is safe to reuse; projected computed
/// parameters and source references remain demand-sensitive.  Volatile
/// functions are reusable only inside the evaluator's explicit timestamp
/// snapshot.
pub(super) fn cacheable_branch<'expr, R: Resolver + ?Sized>(
    evaluator: &mut ValueEvaluator<'expr, '_, '_, '_, R>,
    node: super::super::Node<'expr>,
) -> EvaluationResult<bool> {
    let Some(function) = node.function_name().and_then(Function::from_name) else {
        return Ok(false);
    };
    if matches!(
        function,
        Function::Now | Function::Today | Function::EasterSunday
    ) && evaluator
        .scalar
        .context
        .options()
        .calculation_timestamp()
        .is_none()
    {
        return Ok(false);
    }
    for index in 0..node.child_count() {
        let Some(mut child) = node.child(index) else {
            return Ok(false);
        };
        loop {
            evaluator.scalar.charge_work(1)?;
            if !matches!(child.kind(), super::super::Kind::Parenthesized) {
                break;
            }
            let Some(inner) = child.child(0) else {
                return Ok(false);
            };
            child = inner;
        }
        let invariant = match child.kind() {
            super::super::Kind::Number
            | super::super::Kind::String
            | super::super::Kind::Error
            | super::super::Kind::Missing => true,
            super::super::Kind::Reference(reference) => {
                ValueEvaluator::<R>::cacheable_reference(reference)
            },
            // Sequence descriptors are stable under the source fence, but
            // an array-valued Date/Offset argument is position-sensitive.
            super::super::Kind::Array(_) => sequence_argument(function_name(function), index),
            _ => false,
        };
        if !invariant {
            return Ok(false);
        }
    }
    Ok(true)
}

fn function_name(function: Function) -> &'static str {
    match function {
        Function::Date => "DATE",
        Function::DateDif => "DATEDIF",
        Function::DateValue => "DATEVALUE",
        Function::Day => "DAY",
        Function::Days => "DAYS",
        Function::Days360 => "DAYS360",
        Function::EasterSunday => "EASTERSUNDAY",
        Function::EDate => "EDATE",
        Function::EOMonth => "EOMONTH",
        Function::Hour => "HOUR",
        Function::IsoWeeknum => "ISOWEEKNUM",
        Function::Minute => "MINUTE",
        Function::Month => "MONTH",
        Function::Networkdays => "NETWORKDAYS",
        Function::Now => "NOW",
        Function::Second => "SECOND",
        Function::Time => "TIME",
        Function::TimeValue => "TIMEVALUE",
        Function::Today => "TODAY",
        Function::Weekday => "WEEKDAY",
        Function::Weeknum => "WEEKNUM",
        Function::Workday => "WORKDAY",
        Function::Year => "YEAR",
        Function::Yearfrac => "YEARFRAC",
    }
}

fn is_array_node(mut node: super::super::Node<'_>) -> bool {
    while matches!(node.kind(), super::super::Kind::Parenthesized) {
        let Some(child) = node.child(0) else {
            return false;
        };
        node = child;
    }
    matches!(node.kind(), super::super::Kind::Array(_))
}

/// Date/time functions are scalar-kernel calls unless a date/offset argument
/// supplies an array or a rectangular reference in matrix mode.
pub(super) fn apply<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    node: super::super::Node<'expr>,
    name: &str,
    arguments: Vec<RuntimeValue<'expr>>,
) -> EvaluationResult<RuntimeValue<'expr>>
where
    R: Resolver + ?Sized,
{
    let Some(function) = Function::from_name(name) else {
        return Err(EvaluationFailure::Unsupported(
            super::super::UnsupportedKind::Function,
        ));
    };
    if !function.valid_arity(arguments.len()) {
        return Ok(formula_error(ScalarError::Value));
    }
    if arguments
        .iter()
        .any(|value| matches!(value, RuntimeValue::SourceReference))
    {
        return Err(EvaluationFailure::Unsupported(
            super::super::UnsupportedKind::Reference,
        ));
    }
    // A ReferenceList is a known scalar/sequence pseudotype refusal. Keep it
    // before any projection or scan so the refusal performs zero resolver
    // reads. An already-materialized formula error in an earlier argument
    // still wins by source order without forcing a read from the list.
    if let Some(list_index) = arguments
        .iter()
        .position(|value| matches!(value, RuntimeValue::Areas(areas) if areas.is_list))
    {
        let mut earlier_error = None;
        for (index, value) in arguments.iter().enumerate().take(list_index) {
            if let Some(error) = first_materialized_error(evaluator, value)? {
                earlier_error = Some((index, error));
                break;
            }
        }
        if let Some((error_index, error)) = earlier_error
            && error_index < list_index
        {
            return Ok(formula_error(error));
        }
        return Ok(formula_error(ScalarError::Value));
    }
    if is_sequence_reducer(name) {
        return apply_sequence_reducer(evaluator, function, arguments);
    }

    if evaluator.mode == super::Mode::Matrix && arguments.iter().any(RuntimeValue::is_array_like) {
        return map_scalar_function(evaluator, node, name, arguments);
    }
    evaluator.scalar_apply_function(node, name, arguments)
}

fn map_scalar_function<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    node: super::super::Node<'expr>,
    name: &str,
    arguments: Vec<RuntimeValue<'expr>>,
) -> EvaluationResult<RuntimeValue<'expr>>
where
    R: Resolver + ?Sized,
{
    let mut shape = Shape::new(1, 1)?;
    for argument in &arguments {
        let Some(next) = broadcast_shape(shape, evaluator.runtime_shape(argument)?) else {
            return Ok(formula_error(ScalarError::Value));
        };
        shape = next;
    }
    let cells = shape
        .cell_count()
        .ok_or(EvaluationFailure::InvalidExpression(
            "date/time array cell count overflow",
        ))?;
    let (mut output, output_reservation) = evaluator.new_element_vec(cells)?;
    let argument_count = arguments.len();
    for index in 0..cells {
        evaluator.charge_cell_work(index)?;
        let mut scalar_arguments: [Option<RuntimeValue<'expr>>; 4] = std::array::from_fn(|_| None);
        for (argument_index, argument) in arguments.iter().enumerate() {
            scalar_arguments[argument_index] = Some(element_to_runtime(
                evaluator.select_matrix_value(argument, shape, index)?,
            ));
        }
        let value =
            scalar_apply_function_fixed(evaluator, node, name, scalar_arguments, argument_count)?;
        output.push(evaluator.runtime_to_element(value)?);
    }
    evaluator.make_array(shape, output, output_reservation, None)
}

/// Apply one broadcasted date/time cell without allocating an argument Vec.
/// The date/time catalog has at most four arguments, so a fixed stack is both
/// bounded by construction and keeps borrowed TextValue payloads intact.
fn scalar_apply_function_fixed<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    node: super::super::Node<'expr>,
    name: &str,
    arguments: [Option<RuntimeValue<'expr>>; 4],
    count: usize,
) -> EvaluationResult<RuntimeValue<'expr>>
where
    R: Resolver + ?Sized,
{
    if count > arguments.len() || count != node.child_count() {
        return Err(EvaluationFailure::InvalidExpression(
            "date/time scalar argument count disagrees with node",
        ));
    }
    let mut scalar_arguments: [Option<WorkingValue<'expr>>; 4] = std::array::from_fn(|_| None);
    for (index, argument) in arguments.into_iter().take(count).enumerate() {
        let argument = argument.ok_or(EvaluationFailure::InvalidExpression(
            "date/time scalar argument slot is empty",
        ))?;
        let slot = evaluator.value_to_slot(argument)?;
        scalar_arguments[index] = Some(super::scalar::argument(name, index, slot));
    }
    let values = scalar_arguments.into_iter().take(count).map(|value| {
        value.ok_or(EvaluationFailure::InvalidExpression(
            "date/time scalar argument slot is empty",
        ))
    });
    super::scalar::eager(&mut evaluator.scalar, node, name, values).map(RuntimeValue::Scalar)
}

#[derive(Clone, Copy, Debug)]
struct ErrorRecord {
    argument: usize,
    ordinal: usize,
    error: ScalarError,
}

#[derive(Clone, Copy, Debug)]
struct ConvertedResult {
    value: Result<f64, ScalarError>,
    formula_error: Option<ScalarError>,
}

impl ConvertedResult {
    fn generated(value: Result<f64, ScalarError>) -> Self {
        Self {
            value,
            formula_error: None,
        }
    }

    fn formula(error: ScalarError) -> Self {
        Self {
            value: Err(error),
            formula_error: Some(error),
        }
    }
}

#[derive(Debug)]
struct PreparedSequences {
    holidays: Vec<i64>,
    holiday_reservation: Option<Reservation>,
    workdays: [bool; 7],
    formula_error: Option<ErrorRecord>,
    generated_error: Option<ErrorRecord>,
}

impl PreparedSequences {
    fn remember(slot: &mut Option<ErrorRecord>, record: ErrorRecord) {
        if slot.is_none_or(|previous| {
            (record.argument, record.ordinal) < (previous.argument, previous.ordinal)
        }) {
            *slot = Some(record);
        }
    }

    fn formula(&mut self, argument: usize, ordinal: usize, error: ScalarError) {
        Self::remember(
            &mut self.formula_error,
            ErrorRecord {
                argument,
                ordinal,
                error,
            },
        );
    }

    fn generated(&mut self, argument: usize, ordinal: usize, error: ScalarError) {
        Self::remember(
            &mut self.generated_error,
            ErrorRecord {
                argument,
                ordinal,
                error,
            },
        );
    }
}

fn apply_sequence_reducer<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    function: Function,
    mut arguments: Vec<RuntimeValue<'expr>>,
) -> EvaluationResult<RuntimeValue<'expr>>
where
    R: Resolver + ?Sized,
{
    let first = std::mem::replace(&mut arguments[0], RuntimeValue::Missing);
    let second = std::mem::replace(&mut arguments[1], RuntimeValue::Missing);
    let holiday = if arguments.len() > 2 {
        std::mem::replace(&mut arguments[2], RuntimeValue::Missing)
    } else {
        RuntimeValue::Missing
    };
    let workweek = if arguments.len() > 3 {
        std::mem::replace(&mut arguments[3], RuntimeValue::Missing)
    } else {
        RuntimeValue::Missing
    };
    let prepared = prepare_sequences(evaluator, holiday, workweek)?;
    let matrix =
        evaluator.mode == super::Mode::Matrix && (first.is_array_like() || second.is_array_like());
    if matrix {
        return map_sequence_reducer(evaluator, function, first, second, prepared);
    }
    let first = scalar_date(evaluator, first)?;
    let second = if function == Function::Networkdays {
        scalar_date(evaluator, second)?
    } else {
        scalar_number(evaluator, second)?
    };
    if let Some(error) = effective_error(&prepared, &first, &second) {
        return Ok(formula_error(error));
    }
    reducer_value(evaluator, function, first.value, second.value, &prepared)
}

fn map_sequence_reducer<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    function: Function,
    first: RuntimeValue<'expr>,
    second: RuntimeValue<'expr>,
    prepared: PreparedSequences,
) -> EvaluationResult<RuntimeValue<'expr>>
where
    R: Resolver + ?Sized,
{
    let mut shape = Shape::new(1, 1)?;
    let Some(next) = broadcast_shape(shape, evaluator.runtime_shape(&first)?) else {
        return Ok(formula_error(ScalarError::Value));
    };
    shape = next;
    let Some(next) = broadcast_shape(shape, evaluator.runtime_shape(&second)?) else {
        return Ok(formula_error(ScalarError::Value));
    };
    shape = next;
    let cells = shape
        .cell_count()
        .ok_or(EvaluationFailure::InvalidExpression(
            "date/time array cell count overflow",
        ))?;
    let (mut output, output_reservation) = evaluator.new_element_vec(cells)?;
    for index in 0..cells {
        evaluator.charge_cell_work(index)?;
        let first_value = element_to_runtime(evaluator.select_matrix_value(&first, shape, index)?);
        let second_value =
            element_to_runtime(evaluator.select_matrix_value(&second, shape, index)?);
        let first_result = scalar_date(evaluator, first_value)?;
        let second_result = if function == Function::Networkdays {
            scalar_date(evaluator, second_value)?
        } else {
            scalar_number(evaluator, second_value)?
        };
        let value = if let Some(error) = effective_error(&prepared, &first_result, &second_result) {
            formula_error(error)
        } else {
            reducer_value(
                evaluator,
                function,
                first_result.value,
                second_result.value,
                &prepared,
            )?
        };
        output.push(evaluator.runtime_to_element(value)?);
    }
    evaluator.make_array(shape, output, output_reservation, None)
}

fn effective_error(
    prepared: &PreparedSequences,
    first: &ConvertedResult,
    second: &ConvertedResult,
) -> Option<ScalarError> {
    let actual = [
        first.formula_error.map(|error| ErrorRecord {
            argument: 0,
            ordinal: 0,
            error,
        }),
        second.formula_error.map(|error| ErrorRecord {
            argument: 1,
            ordinal: 0,
            error,
        }),
        prepared.formula_error,
    ]
    .into_iter()
    .flatten()
    .min_by_key(|record| (record.argument, record.ordinal));
    let generated = [
        first
            .formula_error
            .is_none()
            .then_some(first.value)
            .and_then(|value| {
                value.err().map(|error| ErrorRecord {
                    argument: 0,
                    ordinal: 0,
                    error,
                })
            }),
        second
            .formula_error
            .is_none()
            .then_some(second.value)
            .and_then(|value| {
                value.err().map(|error| ErrorRecord {
                    argument: 1,
                    ordinal: 0,
                    error,
                })
            }),
        prepared.generated_error,
    ]
    .into_iter()
    .flatten()
    .min_by_key(|record| (record.argument, record.ordinal));
    actual.or(generated).map(|record| record.error)
}

fn reducer_value<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    function: Function,
    first: Result<f64, ScalarError>,
    second: Result<f64, ScalarError>,
    prepared: &PreparedSequences,
) -> EvaluationResult<RuntimeValue<'expr>>
where
    R: Resolver + ?Sized,
{
    let result = match (first, second) {
        (Err(error), _) | (_, Err(error)) => Ok(Err(error)),
        (Ok(first), Ok(second)) => match function {
            Function::Networkdays | Function::Workday => date_time::finish_sequence(
                &mut evaluator.scalar,
                function,
                first,
                second,
                &prepared.workdays,
                &prepared.holidays,
            ),
            _ => Ok(Err(ScalarError::Value)),
        },
    };
    let result = result?;
    Ok(match result {
        Ok(value) if value.is_finite() => RuntimeValue::Scalar(WorkingValue::Number(value)),
        Ok(_) => formula_error(ScalarError::Number),
        Err(error) => formula_error(error),
    })
}

fn prepare_sequences<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    holiday: RuntimeValue<'expr>,
    workweek: RuntimeValue<'expr>,
) -> EvaluationResult<PreparedSequences>
where
    R: Resolver + ?Sized,
{
    let mut prepared = PreparedSequences {
        holidays: Vec::new(),
        holiday_reservation: None,
        workdays: date_time::default_workday_mask(),
        formula_error: None,
        generated_error: None,
    };
    if !matches!(holiday, RuntimeValue::Missing) {
        scan_holidays(evaluator, holiday, &mut prepared, 2)?;
    }
    if !matches!(workweek, RuntimeValue::Missing) {
        scan_workweek(evaluator, workweek, &mut prepared, 3)?;
    }
    let sort_work = prepared.holidays.len().checked_mul(
        usize::BITS.saturating_sub(prepared.holidays.len().max(1).leading_zeros()) as usize,
    );
    evaluator
        .scalar
        .charge_work(u64::try_from(sort_work.unwrap_or(usize::MAX)).unwrap_or(u64::MAX))?;
    prepared.holidays.sort_unstable();
    evaluator
        .scalar
        .charge_work(u64::try_from(prepared.holidays.len()).unwrap_or(u64::MAX))?;
    prepared.holidays.dedup();
    Ok(prepared)
}

fn scan_holidays<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    value: RuntimeValue<'expr>,
    prepared: &mut PreparedSequences,
    argument: usize,
) -> EvaluationResult<()>
where
    R: Resolver + ?Sized,
{
    match value {
        RuntimeValue::Areas(areas) => scan_holiday_areas(evaluator, areas, prepared, argument),
        RuntimeValue::Array(array) => {
            for (ordinal, element) in array.cells.into_iter().enumerate() {
                evaluator.charge_cell_work(ordinal)?;
                observe_holiday_element(evaluator, element, prepared, argument, ordinal, false)?;
            }
            Ok(())
        },
        RuntimeValue::ScalarCell(area) => {
            let value = evaluator.project_scalar(RuntimeValue::ScalarCell(area))?;
            scan_holidays(evaluator, value, prepared, argument)
        },
        RuntimeValue::Scalar(value) => {
            evaluator.scalar.charge_work(1)?;
            observe_holiday_working(evaluator, value, prepared, argument, 0, false)
        },
        RuntimeValue::Empty => Ok(()),
        RuntimeValue::Missing => Ok(()),
        RuntimeValue::SourceReference => Err(EvaluationFailure::Unsupported(
            super::super::UnsupportedKind::Reference,
        )),
    }
}

fn scan_holiday_areas<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    areas: RuntimeAreaSet<'expr>,
    prepared: &mut PreparedSequences,
    argument: usize,
) -> EvaluationResult<()>
where
    R: Resolver + ?Sized,
{
    if areas.is_list {
        prepared.generated(argument, 0, ScalarError::Value);
        return Ok(());
    }
    let mut ordinal = 0usize;
    for area in areas.areas {
        for row in area.rect.row_start..area.rect.row_end {
            for column in area.rect.column_start..area.rect.column_end {
                evaluator.charge_cell_work(ordinal)?;
                let read = evaluator.read_reference_cell(area.sheet, row, column)?;
                let element = evaluator.read_to_element(read)?;
                observe_holiday_element(evaluator, element, prepared, argument, ordinal, true)?;
                ordinal = ordinal
                    .checked_add(1)
                    .ok_or(EvaluationFailure::InvalidExpression(
                        "holiday sequence ordinal overflows",
                    ))?;
            }
        }
    }
    Ok(())
}

fn observe_holiday_element<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    element: RuntimeElement<'expr>,
    prepared: &mut PreparedSequences,
    argument: usize,
    ordinal: usize,
    reference: bool,
) -> EvaluationResult<()>
where
    R: Resolver + ?Sized,
{
    match element {
        RuntimeElement::Empty => {
            if !reference {
                observe_holiday_number(evaluator, 0.0, prepared, argument, ordinal)?;
            }
        },
        RuntimeElement::Missing => prepared.generated(argument, ordinal, ScalarError::Value),
        RuntimeElement::Present(value) => {
            observe_holiday_working(evaluator, value, prepared, argument, ordinal, reference)?;
        },
    }
    Ok(())
}

fn observe_holiday_working<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    value: WorkingValue<'expr>,
    prepared: &mut PreparedSequences,
    argument: usize,
    ordinal: usize,
    reference: bool,
) -> EvaluationResult<()>
where
    R: Resolver + ?Sized,
{
    match value {
        WorkingValue::Error(error) => prepared.formula(argument, ordinal, error),
        WorkingValue::Text(text) if reference => {
            evaluator.scalar.charge_bytes(text.text.len())?;
        },
        WorkingValue::Text(text) => {
            let result = super::super::to_number(WorkingValue::Text(text), &mut evaluator.scalar)?;
            match result {
                Ok(value) => observe_holiday_number(evaluator, value, prepared, argument, ordinal)?,
                Err(error) => prepared.generated(argument, ordinal, error),
            }
        },
        WorkingValue::Number(value) => {
            if !reference {
                observe_holiday_number(evaluator, value, prepared, argument, ordinal)?;
            } else {
                observe_holiday_number(evaluator, value, prepared, argument, ordinal)?;
            }
        },
        WorkingValue::Logical(value) => {
            if !reference {
                observe_holiday_number(
                    evaluator,
                    if value { 1.0 } else { 0.0 },
                    prepared,
                    argument,
                    ordinal,
                )?;
            }
        },
        WorkingValue::Complex(_) => prepared.generated(argument, ordinal, ScalarError::Value),
    }
    Ok(())
}

fn observe_holiday_number<R: Resolver + ?Sized>(
    evaluator: &mut ValueEvaluator<'_, '_, '_, '_, R>,
    value: f64,
    prepared: &mut PreparedSequences,
    argument: usize,
    ordinal: usize,
) -> EvaluationResult<()> {
    let value = if value.is_finite() {
        value
    } else {
        prepared.generated(argument, ordinal, ScalarError::Number);
        return Ok(());
    };
    match date_time::serial_to_date(value).and_then(date_time::date_to_serial) {
        Ok(value) => {
            let value = value as i64;
            // Grow only after the value has passed the date-domain check.
            ensure_capacity(
                &mut prepared.holidays,
                &mut prepared.holiday_reservation,
                1,
                evaluator.limits.max_array_cells,
                evaluator.execution,
                &evaluator.storage_budget,
                "formula date/time holiday sequence",
            )?;
            prepared.holidays.push(value);
        },
        Err(error) => prepared.generated(argument, ordinal, error),
    }
    Ok(())
}

fn scan_workweek<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    value: RuntimeValue<'expr>,
    prepared: &mut PreparedSequences,
    argument: usize,
) -> EvaluationResult<()>
where
    R: Resolver + ?Sized,
{
    let mut count = 0usize;
    match value {
        RuntimeValue::Areas(areas) => {
            if areas.is_list {
                prepared.generated(argument, 0, ScalarError::Value);
                return Ok(());
            }
            for area in areas.areas {
                for row in area.rect.row_start..area.rect.row_end {
                    for column in area.rect.column_start..area.rect.column_end {
                        evaluator.charge_cell_work(count)?;
                        let read = evaluator.read_reference_cell(area.sheet, row, column)?;
                        let element = evaluator.read_to_element(read)?;
                        if let RuntimeElement::Present(value) = element {
                            match value {
                                WorkingValue::Logical(value) => {
                                    if count < 7 {
                                        prepared.workdays[count] = value;
                                    }
                                    count = count.saturating_add(1);
                                },
                                WorkingValue::Error(error) => {
                                    prepared.formula(argument, count, error);
                                    count = count.saturating_add(1);
                                },
                                WorkingValue::Text(text) => {
                                    evaluator.scalar.charge_bytes(text.text.len())?;
                                },
                                WorkingValue::Number(_) | WorkingValue::Complex(_) => {},
                            }
                        }
                    }
                }
            }
        },
        RuntimeValue::Array(array) => {
            let cells = array_cell_count(&array)?;
            for (index, element) in array.cells.into_iter().enumerate() {
                evaluator.charge_cell_work(index)?;
                match element {
                    RuntimeElement::Empty => {
                        if index < 7 {
                            prepared.workdays[index] = false;
                        }
                    },
                    RuntimeElement::Missing => {
                        prepared.generated(argument, index, ScalarError::Value)
                    },
                    RuntimeElement::Present(WorkingValue::Number(value)) => {
                        if index < 7 {
                            prepared.workdays[index] = value != 0.0;
                        }
                    },
                    RuntimeElement::Present(WorkingValue::Logical(value)) => {
                        if index < 7 {
                            prepared.workdays[index] = value;
                        }
                    },
                    RuntimeElement::Present(WorkingValue::Text(text)) => {
                        evaluator.scalar.charge_bytes(text.text.len())?;
                        prepared.generated(argument, index, ScalarError::Value);
                    },
                    RuntimeElement::Present(WorkingValue::Error(error)) => {
                        prepared.formula(argument, index, error);
                    },
                    RuntimeElement::Present(WorkingValue::Complex(_)) => {
                        prepared.generated(argument, index, ScalarError::Value);
                    },
                }
            }
            count = cells;
        },
        RuntimeValue::Scalar(value) => {
            evaluator.scalar.charge_work(1)?;
            count = 1;
            match value {
                WorkingValue::Logical(value) => prepared.workdays[0] = value,
                WorkingValue::Number(value) => prepared.workdays[0] = value != 0.0,
                WorkingValue::Error(error) => prepared.formula(argument, 0, error),
                WorkingValue::Text(text) => {
                    evaluator.scalar.charge_bytes(text.text.len())?;
                    prepared.generated(argument, 0, ScalarError::Value);
                },
                WorkingValue::Complex(_) => prepared.generated(argument, 0, ScalarError::Value),
            }
        },
        RuntimeValue::ScalarCell(area) => {
            let value = evaluator.project_scalar(RuntimeValue::ScalarCell(area))?;
            return scan_workweek(evaluator, value, prepared, argument);
        },
        RuntimeValue::Empty => {
            evaluator.scalar.charge_work(1)?;
            count = 1;
            prepared.workdays[0] = false;
        },
        RuntimeValue::Missing => return Ok(()),
        RuntimeValue::SourceReference => {
            return Err(EvaluationFailure::Unsupported(
                super::super::UnsupportedKind::Reference,
            ));
        },
    }
    if count != 7 {
        prepared.generated(argument, count, ScalarError::Value);
    }
    Ok(())
}

fn array_cell_count(array: &super::RuntimeArrayValue<'_>) -> EvaluationResult<usize> {
    array
        .shape
        .cell_count()
        .ok_or(EvaluationFailure::InvalidExpression(
            "workweek array cell count overflows",
        ))
}

fn scalar_date<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    value: RuntimeValue<'expr>,
) -> EvaluationResult<ConvertedResult>
where
    R: Resolver + ?Sized,
{
    let value = evaluator.project_scalar(value)?;
    match value {
        RuntimeValue::Empty => Ok(ConvertedResult::generated(Ok(0.0))),
        RuntimeValue::Missing => Ok(ConvertedResult::generated(Err(ScalarError::Value))),
        RuntimeValue::Scalar(WorkingValue::Number(value)) => {
            Ok(ConvertedResult::generated(valid_serial(value)))
        },
        RuntimeValue::Scalar(WorkingValue::Logical(value)) => {
            Ok(ConvertedResult::generated(Ok(if value {
                1.0
            } else {
                0.0
            })))
        },
        RuntimeValue::Scalar(WorkingValue::Text(text)) => {
            date_time::parse_date_text(&mut evaluator.scalar, text.text.as_ref())
                .map(ConvertedResult::generated)
        },
        RuntimeValue::Scalar(WorkingValue::Error(error)) => Ok(ConvertedResult::formula(error)),
        RuntimeValue::Scalar(WorkingValue::Complex(_)) => {
            Ok(ConvertedResult::generated(Err(ScalarError::Value)))
        },
        RuntimeValue::Array(_) | RuntimeValue::Areas(_) | RuntimeValue::ScalarCell(_) => {
            Ok(ConvertedResult::generated(Err(ScalarError::Value)))
        },
        RuntimeValue::SourceReference => Err(EvaluationFailure::Unsupported(
            super::super::UnsupportedKind::Reference,
        )),
    }
}

fn scalar_number<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    value: RuntimeValue<'expr>,
) -> EvaluationResult<ConvertedResult>
where
    R: Resolver + ?Sized,
{
    let value = evaluator.project_scalar(value)?;
    match value {
        RuntimeValue::Empty => Ok(ConvertedResult::generated(Ok(0.0))),
        RuntimeValue::Missing => Ok(ConvertedResult::generated(Err(ScalarError::Value))),
        RuntimeValue::Scalar(WorkingValue::Error(error)) => Ok(ConvertedResult::formula(error)),
        RuntimeValue::Scalar(value) => {
            super::super::to_number(value, &mut evaluator.scalar).map(ConvertedResult::generated)
        },
        RuntimeValue::Array(_) | RuntimeValue::Areas(_) | RuntimeValue::ScalarCell(_) => {
            Ok(ConvertedResult::generated(Err(ScalarError::Value)))
        },
        RuntimeValue::SourceReference => Err(EvaluationFailure::Unsupported(
            super::super::UnsupportedKind::Reference,
        )),
    }
}

fn valid_serial(value: f64) -> Result<f64, ScalarError> {
    if value.is_finite() && (date_time::MIN_SERIAL..date_time::MAX_DATETIME_SERIAL).contains(&value)
    {
        Ok(value)
    } else {
        Err(ScalarError::Number)
    }
}

fn formula_error<'a>(error: ScalarError) -> RuntimeValue<'a> {
    RuntimeValue::Scalar(WorkingValue::Error(error))
}

fn first_materialized_error<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    value: &RuntimeValue<'expr>,
) -> EvaluationResult<Option<ScalarError>>
where
    R: Resolver + ?Sized,
{
    match value {
        RuntimeValue::Scalar(WorkingValue::Error(error)) => Ok(Some(*error)),
        RuntimeValue::Array(array) => {
            for (index, element) in array.cells.iter().enumerate() {
                evaluator.charge_cell_work(index)?;
                let error = match element {
                    RuntimeElement::Missing => Some(ScalarError::NotAvailable),
                    RuntimeElement::Present(WorkingValue::Error(error)) => Some(*error),
                    RuntimeElement::Empty | RuntimeElement::Present(_) => None,
                };
                if error.is_some() {
                    return Ok(error);
                }
            }
            Ok(None)
        },
        _ => Ok(None),
    }
}
