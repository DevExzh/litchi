//! Resolver-aware order and rank reducers.
//!
//! References are scanned directly and only admitted finite numbers are
//! retained.  The bounded numeric storage and all order calculations are
//! shared with the scalar evaluator in [`super::super::order`].

use super::super::order::{self as kernel, OrderFunction, OrderValues};
use super::{
    EvaluationFailure, EvaluationResult, Resolver, Resource, RuntimeAreaSet, RuntimeArrayValue,
    RuntimeElement, RuntimeValue, ScalarError, Shape, TextValue, ValueEvaluator, WorkingValue,
    array_element_for, broadcast_shape,
};

/// Return whether `name` is one of the order/rank reducers.
pub(super) fn is_order_function(name: &str) -> bool {
    kernel::is_order_function(name)
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

fn is_reference_sequence_node(mut node: super::super::Node<'_>) -> bool {
    loop {
        match node.kind() {
            super::super::Kind::Parenthesized => {
                let Some(child) = node.child(0) else {
                    return false;
                };
                node = child;
            },
            super::super::Kind::Reference(_) => return true,
            super::super::Kind::Infix(
                super::super::InfixOperator::Range
                | super::super::InfixOperator::Intersection
                | super::super::InfixOperator::Union,
            ) => return true,
            super::super::Kind::Function { name }
                if name.eq_ignore_ascii_case("IF")
                    || name.eq_ignore_ascii_case("IFERROR")
                    || name.eq_ignore_ascii_case("IFNA") =>
            {
                return true;
            },
            _ => return false,
        }
    }
}

const fn accepts_reference_list(function: OrderFunction) -> bool {
    !matches!(function, OrderFunction::Mode | OrderFunction::Quartile)
}

/// Enter a complete sequence argument in matrix context.  Direct references
/// already retain their descriptors under `VisitArgument`; reference-valued
/// lazy handlers and ForceArray MODE arguments need the explicit context.
pub(super) fn matrix_argument(name: &str, index: usize, node: super::super::Node<'_>) -> bool {
    let Some(function) = OrderFunction::from_name(name) else {
        return false;
    };
    if function.data_argument(index) {
        return function == OrderFunction::Mode
            || is_array_node(node)
            || is_reference_sequence_node(node);
    }
    // Scalar parameters are lifted when an explicit Array is supplied. A
    // computed scalar remains in the caller's projection so nested MUNIT and
    // other position-sensitive expressions are evaluated per output cell.
    is_array_node(node)
}

/// Whether a nested criterion reducer consumes a complete argument at this
/// position. Scalar order parameters stay coordinate-sensitive.
pub(super) fn criterion_full_argument(
    name: &str,
    index: usize,
    node: super::super::Node<'_>,
) -> bool {
    let Some(function) = OrderFunction::from_name(name) else {
        return false;
    };
    function.data_argument(index) || (function == OrderFunction::Mode) || is_array_node(node)
}

#[derive(Clone, Copy, Debug)]
struct ErrorRecord {
    argument: usize,
    ordinal: usize,
    error: ScalarError,
}

#[derive(Debug)]
struct State {
    values: OrderValues,
    stream: Option<StreamState>,
    maximum: usize,
    formula_error: Option<ErrorRecord>,
    generated_error: Option<ErrorRecord>,
}

impl State {
    fn new(stream: Option<StreamState>, maximum: usize) -> Self {
        Self {
            values: OrderValues::new(),
            stream,
            maximum,
            formula_error: None,
            generated_error: None,
        }
    }

    fn push_number<R: Resolver + ?Sized>(
        &mut self,
        evaluator: &mut ValueEvaluator<'_, '_, '_, '_, R>,
        value: f64,
    ) -> EvaluationResult<()> {
        if let Some(stream) = &mut self.stream {
            return stream.observe(evaluator, value);
        }
        self.values.push_with(
            value,
            self.maximum,
            evaluator.execution,
            &evaluator.storage_budget,
            "formula order numeric values",
        )
    }

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

#[derive(Clone, Copy, Debug)]
enum ParameterValue {
    Number(f64),
    Formula(ScalarError),
    Generated(ScalarError),
}

enum Parameter<'a> {
    Scalar(ParameterValue),
    Array(RuntimeArrayValue<'a>),
}

/// Fixed-size state for the two reducers whose scalar query can be answered
/// during the admitted data scan.  A reference or array parameter does not
/// enter this path: it is left for the buffered implementation so parameter
/// resolution keeps its source order.
#[derive(Clone, Copy, Debug)]
enum StreamState {
    Rank {
        value: Option<f64>,
        order: Option<f64>,
        found: bool,
        preceding: usize,
    },
    PercentRank {
        value: Option<f64>,
        significance: Option<f64>,
        count: usize,
        minimum: Option<f64>,
        maximum: Option<f64>,
        less_than_value: usize,
        found_equal: bool,
        lower: Option<f64>,
        lower_equal_count: usize,
        upper: Option<f64>,
    },
}

#[derive(Clone, Copy, Debug)]
struct StreamPlan {
    state: StreamState,
    parameters: [Option<ParameterValue>; 3],
}

impl StreamState {
    fn rank(value: Option<f64>, order: Option<f64>) -> Self {
        Self::Rank {
            value,
            order,
            found: false,
            preceding: 0,
        }
    }

    fn percent_rank(value: Option<f64>, significance: Option<f64>) -> Self {
        Self::PercentRank {
            value,
            significance,
            count: 0,
            minimum: None,
            maximum: None,
            less_than_value: 0,
            found_equal: false,
            lower: None,
            lower_equal_count: 0,
            upper: None,
        }
    }

    fn observe<R: Resolver + ?Sized>(
        &mut self,
        evaluator: &mut ValueEvaluator<'_, '_, '_, '_, R>,
        candidate: f64,
    ) -> EvaluationResult<()> {
        evaluator.scalar.charge_work(1)?;
        match self {
            Self::Rank {
                value: Some(value),
                order: Some(order),
                found,
                preceding,
            } => {
                if candidate == *value {
                    *found = true;
                } else if (*order == 0.0 && candidate > *value)
                    || (*order != 0.0 && candidate < *value)
                {
                    *preceding =
                        preceding
                            .checked_add(1)
                            .ok_or(EvaluationFailure::InvalidExpression(
                                "order rank count overflows",
                            ))?;
                }
            },
            Self::Rank { .. } => {},
            Self::PercentRank {
                value,
                count,
                minimum,
                maximum,
                less_than_value,
                found_equal,
                lower,
                lower_equal_count,
                upper,
                ..
            } => {
                *count = count
                    .checked_add(1)
                    .ok_or(EvaluationFailure::InvalidExpression(
                        "percent-rank count overflows",
                    ))?;
                *minimum = Some(minimum.map_or(candidate, |previous| previous.min(candidate)));
                *maximum = Some(maximum.map_or(candidate, |previous| previous.max(candidate)));
                let Some(value) = *value else {
                    return Ok(());
                };
                if candidate < value {
                    *less_than_value = less_than_value.checked_add(1).ok_or(
                        EvaluationFailure::InvalidExpression("percent-rank lower count overflows"),
                    )?;
                    match *lower {
                        Some(previous) if candidate == previous => {
                            *lower_equal_count = lower_equal_count.checked_add(1).ok_or(
                                EvaluationFailure::InvalidExpression(
                                    "percent-rank duplicate count overflows",
                                ),
                            )?;
                        },
                        Some(previous) if candidate > previous => {
                            *lower = Some(candidate);
                            *lower_equal_count = 1;
                        },
                        None => {
                            *lower = Some(candidate);
                            *lower_equal_count = 1;
                        },
                        Some(_) => {},
                    }
                } else if candidate == value {
                    *found_equal = true;
                } else {
                    match *upper {
                        Some(previous) if candidate < previous => *upper = Some(candidate),
                        None => *upper = Some(candidate),
                        Some(_) => {},
                    }
                }
            },
        }
        Ok(())
    }

    fn finish(&self) -> Result<f64, ScalarError> {
        match *self {
            Self::Rank {
                value: Some(_),
                order: Some(_),
                found: true,
                preceding,
            } => preceding
                .checked_add(1)
                .map(|rank| rank as f64)
                .ok_or(ScalarError::Value),
            Self::Rank { .. } => Err(ScalarError::Value),
            Self::PercentRank {
                value: Some(value),
                significance: Some(significance),
                count,
                minimum: Some(minimum),
                maximum: Some(maximum),
                less_than_value,
                found_equal,
                lower,
                lower_equal_count,
                upper,
            } => {
                if count == 0 || value < minimum || value > maximum {
                    return Err(ScalarError::Value);
                }
                let significance = kernel::exact_positive_integer(significance)?;
                if count == 1 {
                    return if found_equal {
                        Ok(1.0)
                    } else {
                        Err(ScalarError::Value)
                    };
                }
                let raw = if found_equal {
                    less_than_value as f64 / (count - 1) as f64
                } else {
                    let (Some(lower), Some(upper)) = (lower, upper) else {
                        return Err(ScalarError::Value);
                    };
                    let fraction = kernel::fraction_between(lower, value, upper)?;
                    let rank = less_than_value
                        .checked_sub(lower_equal_count)
                        .ok_or(ScalarError::Value)?;
                    (rank as f64 + fraction) / (count - 1) as f64
                };
                super::super::rounding::round_nearest(raw, significance)
            },
            Self::PercentRank { .. } => Err(ScalarError::Value),
        }
    }
}

fn stream_parameter(value: &RuntimeValue<'_>) -> bool {
    matches!(
        value,
        RuntimeValue::Empty
            | RuntimeValue::Missing
            | RuntimeValue::Scalar(
                WorkingValue::Number(_) | WorkingValue::Logical(_) | WorkingValue::Error(_)
            )
    )
}

fn direct_parameter(value: &RuntimeValue<'_>) -> ParameterValue {
    match value {
        RuntimeValue::Empty => ParameterValue::Number(0.0),
        RuntimeValue::Missing => ParameterValue::Generated(ScalarError::Value),
        RuntimeValue::Scalar(WorkingValue::Number(value)) if value.is_finite() => {
            ParameterValue::Number(*value)
        },
        RuntimeValue::Scalar(WorkingValue::Number(_)) => {
            ParameterValue::Generated(ScalarError::Number)
        },
        RuntimeValue::Scalar(WorkingValue::Logical(value)) => {
            ParameterValue::Number(if *value { 1.0 } else { 0.0 })
        },
        RuntimeValue::Scalar(WorkingValue::Error(error)) => ParameterValue::Formula(*error),
        RuntimeValue::Scalar(_) => ParameterValue::Generated(ScalarError::Value),
        RuntimeValue::ScalarCell(_) | RuntimeValue::Array(_) | RuntimeValue::Areas(_) => {
            ParameterValue::Generated(ScalarError::Value)
        },
        RuntimeValue::SourceReference => ParameterValue::Generated(ScalarError::Value),
    }
}

fn stream_plan<'expr>(
    function: OrderFunction,
    arguments: &[RuntimeValue<'expr>],
) -> EvaluationResult<Option<StreamPlan>> {
    let mut parameters = [None; 3];
    match function {
        OrderFunction::Rank => {
            let value_argument = arguments
                .first()
                .ok_or(EvaluationFailure::InvalidExpression(
                    "rank stream value is missing",
                ))?;
            if !stream_parameter(value_argument) {
                return Ok(None);
            }
            if arguments
                .get(2)
                .is_some_and(|value| !stream_parameter(value))
            {
                return Ok(None);
            }
            parameters[0] = Some(direct_parameter(value_argument));
            if let Some(value) = arguments.get(2) {
                parameters[2] = Some(direct_parameter(value));
            }
            let value = match parameters[0] {
                Some(ParameterValue::Number(value)) if value.is_finite() => Some(value),
                _ => None,
            };
            let order = match parameters[2] {
                Some(ParameterValue::Number(value)) if value.is_finite() => Some(value),
                None => Some(0.0),
                _ => None,
            };
            Ok(Some(StreamPlan {
                state: StreamState::rank(value, order),
                parameters,
            }))
        },
        OrderFunction::PercentRank => {
            if !stream_parameter(
                arguments
                    .get(1)
                    .ok_or(EvaluationFailure::InvalidExpression(
                        "percent-rank stream value is missing",
                    ))?,
            ) {
                return Ok(None);
            }
            if arguments
                .get(2)
                .is_some_and(|value| !stream_parameter(value))
            {
                return Ok(None);
            }
            parameters[1] = Some(direct_parameter(&arguments[1]));
            if let Some(value) = arguments.get(2) {
                parameters[2] = Some(direct_parameter(value));
            }
            let value = match parameters[1] {
                Some(ParameterValue::Number(value)) if value.is_finite() => Some(value),
                _ => None,
            };
            let significance = match parameters[2] {
                Some(ParameterValue::Number(value)) if value.is_finite() => Some(value),
                None => Some(3.0),
                _ => None,
            };
            Ok(Some(StreamPlan {
                state: StreamState::percent_rank(value, significance),
                parameters,
            }))
        },
        _ => Ok(None),
    }
}

/// Apply one resolver-aware order/rank reducer.
pub(super) fn apply<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    name: &str,
    arguments: Vec<RuntimeValue<'expr>>,
) -> EvaluationResult<RuntimeValue<'expr>>
where
    R: Resolver + ?Sized,
{
    let Some(function) = OrderFunction::from_name(name) else {
        return Err(EvaluationFailure::Unsupported(
            super::super::UnsupportedKind::Function,
        ));
    };
    if arguments
        .iter()
        .any(|value| matches!(value, RuntimeValue::SourceReference))
    {
        return Err(EvaluationFailure::Unsupported(
            super::super::UnsupportedKind::Reference,
        ));
    }
    if !function.valid_arity(arguments.len()) {
        return Ok(formula_error(ScalarError::Value));
    }

    // Shape refusals are checked against descriptors before any scan. This
    // includes scalar parameters: a ReferenceList cannot be projected into a
    // parameter and must not cause the preceding sequence to be read.
    if arguments.iter().enumerate().any(|(index, argument)| {
        matches!(argument, RuntimeValue::Areas(areas) if areas.is_list)
            && (!function.data_argument(index) || !accepts_reference_list(function))
    }) {
        return Ok(formula_error(ScalarError::Value));
    }

    let maximum = admitted_value_limit(evaluator, function, &arguments)?;
    let plan = stream_plan(function, &arguments)?;
    let mut state = State::new(plan.map(|plan| plan.state), maximum);
    let mut parameters: [Option<Parameter<'expr>>; 3] = [None, None, None];
    for (index, argument) in arguments.into_iter().enumerate() {
        if function.data_argument(index) {
            observe_data(evaluator, &mut state, index, argument)?;
        } else if let Some(slot) = parameters.get_mut(index) {
            if let Some(plan) = plan {
                if let Some(value) = plan.parameters.get(index).copied().flatten() {
                    *slot = Some(Parameter::Scalar(value));
                    continue;
                }
            }
            *slot = Some(normalize_parameter(evaluator, argument)?);
        } else {
            state.generated(index, 0, ScalarError::Value);
        }
    }

    let lift_parameters = evaluator.mode == super::Mode::Matrix
        || evaluator.projection.is_some()
        || matches!(function, OrderFunction::Large | OrderFunction::Small);
    let output_shape = parameter_shape(&parameters, lift_parameters)?;
    let cells = output_shape
        .cell_count()
        .ok_or(EvaluationFailure::InvalidExpression(
            "order output shape overflows",
        ))?;
    let preserve_array = matches!(function, OrderFunction::Large | OrderFunction::Small)
        && matches!(parameters.get(1), Some(Some(Parameter::Array(_))));
    let needs_sort = state.formula_error.is_none()
        && state.generated_error.is_none()
        && !matches!(function, OrderFunction::Rank)
        && state.values.len() > 1;
    if needs_sort {
        let mut charge = |work| evaluator.scalar.charge_work(work);
        kernel::sort_values(state.values.as_mut_slice(), true, &mut charge)?;
    }

    if cells == 1 && !preserve_array {
        let mut values = [None; 3];
        for (index, slot) in values.iter_mut().enumerate() {
            if let Some(parameter) = parameters.get(index).and_then(Option::as_ref) {
                *slot = Some(parameter_at(evaluator, parameter, output_shape, 0)?);
            }
        }
        return finish_one(evaluator, &state, function, &values);
    }

    let (mut output, output_reservation) = evaluator.new_element_vec(cells)?;
    for index in 0..cells {
        evaluator.charge_cell_work(index)?;
        let mut values = [None; 3];
        for (parameter_index, slot) in values.iter_mut().enumerate() {
            if let Some(parameter) = parameters.get(parameter_index).and_then(Option::as_ref) {
                *slot = Some(parameter_at(evaluator, parameter, output_shape, index)?);
            }
        }
        let result = evaluate_one(evaluator, &state, function, &values)?;
        output.push(match result {
            Ok(value) => RuntimeElement::Present(WorkingValue::Number(value)),
            Err(error) => RuntimeElement::Present(WorkingValue::Error(error)),
        });
    }
    Ok(RuntimeValue::Array(RuntimeArrayValue {
        shape: output_shape,
        cells: output,
        _reservation: output_reservation,
        origin: None,
        preserve_scalar_result: preserve_array,
    }))
}

fn finish_one<'expr, R: Resolver + ?Sized>(
    evaluator: &mut ValueEvaluator<'expr, '_, '_, '_, R>,
    state: &State,
    function: OrderFunction,
    parameters: &[Option<ParameterValue>; 3],
) -> EvaluationResult<RuntimeValue<'expr>> {
    let result = evaluate_one(evaluator, state, function, parameters)?;
    Ok(match result {
        Ok(value) => RuntimeValue::Scalar(WorkingValue::Number(value)),
        Err(error) => formula_error(error),
    })
}

fn evaluate_one<R: Resolver + ?Sized>(
    evaluator: &mut ValueEvaluator<'_, '_, '_, '_, R>,
    state: &State,
    function: OrderFunction,
    parameters: &[Option<ParameterValue>; 3],
) -> EvaluationResult<Result<f64, ScalarError>> {
    let mut numeric = [None; 3];
    let mut formula = state.formula_error;
    let mut generated = state.generated_error;
    for index in 0..3 {
        let Some(value) = parameters[index] else {
            continue;
        };
        match value {
            ParameterValue::Number(value) => numeric[index] = Some(value),
            ParameterValue::Formula(error) => {
                let record = ErrorRecord {
                    argument: index,
                    ordinal: 0,
                    error,
                };
                if formula.is_none_or(|previous| {
                    (record.argument, record.ordinal) < (previous.argument, previous.ordinal)
                }) {
                    formula = Some(record);
                }
            },
            ParameterValue::Generated(error) => {
                let record = ErrorRecord {
                    argument: index,
                    ordinal: 0,
                    error,
                };
                if generated.is_none_or(|previous| {
                    (record.argument, record.ordinal) < (previous.argument, previous.ordinal)
                }) {
                    generated = Some(record);
                }
            },
        }
    }
    if let Some(error) = formula.or(generated).map(|record| record.error) {
        return Ok(Err(error));
    }

    if let Some(stream) = state.stream {
        return match function {
            OrderFunction::Rank | OrderFunction::PercentRank => {
                charged_result(evaluator, stream.finish())
            },
            _ => Err(EvaluationFailure::InvalidExpression(
                "order stream state reached the wrong reducer",
            )),
        };
    }

    let missing = || Ok(Err(ScalarError::Value));
    match function {
        OrderFunction::Median => {
            charged_result(evaluator, kernel::median_sorted(state.values.as_slice()))
        },
        OrderFunction::Mode => kernel::mode_sorted(state.values.as_slice(), &mut |work| {
            evaluator.scalar.charge_work(work)
        }),
        OrderFunction::Large => match numeric[1] {
            Some(value) => charged_result(
                evaluator,
                kernel::select_sorted(state.values.as_slice(), value, true),
            ),
            None => missing(),
        },
        OrderFunction::Small => match numeric[1] {
            Some(value) => charged_result(
                evaluator,
                kernel::select_sorted(state.values.as_slice(), value, false),
            ),
            None => missing(),
        },
        OrderFunction::Percentile => match numeric[1] {
            Some(value) => charged_result(
                evaluator,
                kernel::percentile_sorted(state.values.as_slice(), value),
            ),
            None => missing(),
        },
        OrderFunction::PercentRank => match numeric[1] {
            Some(value) => kernel::percent_rank_sorted(
                state.values.as_slice(),
                value,
                numeric[2].unwrap_or(3.0),
                &mut |work| evaluator.scalar.charge_work(work),
            ),
            None => missing(),
        },
        OrderFunction::Quartile => match numeric[1] {
            Some(value) => match kernel::quartile_fraction(value) {
                Ok(value) => charged_result(
                    evaluator,
                    kernel::percentile_sorted(state.values.as_slice(), value),
                ),
                Err(error) => Ok(Err(error)),
            },
            None => missing(),
        },
        OrderFunction::Rank => match numeric[0] {
            Some(value) => kernel::rank_values(
                state.values.as_slice(),
                value,
                numeric[2].unwrap_or(0.0),
                &mut |work| evaluator.scalar.charge_work(work),
            ),
            None => missing(),
        },
    }
}

fn charged_result<R: Resolver + ?Sized>(
    evaluator: &mut ValueEvaluator<'_, '_, '_, '_, R>,
    result: Result<f64, ScalarError>,
) -> EvaluationResult<Result<f64, ScalarError>> {
    evaluator.scalar.charge_work(1)?;
    Ok(result)
}

fn admitted_value_limit<R: Resolver + ?Sized>(
    evaluator: &mut ValueEvaluator<'_, '_, '_, '_, R>,
    function: OrderFunction,
    arguments: &[RuntimeValue<'_>],
) -> EvaluationResult<usize> {
    let mut total = 0usize;
    for (index, argument) in arguments.iter().enumerate() {
        if !function.data_argument(index) {
            continue;
        }
        evaluator.scalar.charge_work(1)?;
        let count = match argument {
            RuntimeValue::Areas(areas) => {
                let mut count = 0usize;
                for area in &areas.areas {
                    evaluator.scalar.charge_work(1)?;
                    count = count.checked_add(area.rect.count()?).ok_or(
                        EvaluationFailure::ResourceLimit(evaluator.local_limit(
                            Resource::Objects,
                            u64::MAX,
                            evaluator.limits.max_reference_cells,
                        )),
                    )?;
                }
                count
            },
            RuntimeValue::Array(array) => array.cells.len(),
            RuntimeValue::ScalarCell(_) | RuntimeValue::Empty | RuntimeValue::Missing => 1,
            RuntimeValue::Scalar(_) => 1,
            RuntimeValue::SourceReference => {
                return Err(EvaluationFailure::Unsupported(
                    super::super::UnsupportedKind::Reference,
                ));
            },
        };
        total = total.checked_add(count).ok_or_else(|| {
            EvaluationFailure::ResourceLimit(evaluator.local_limit(
                Resource::Objects,
                u64::MAX,
                evaluator.limits.max_reference_cells,
            ))
        })?;
    }
    Ok(total.max(1))
}

fn observe_data<'expr, R: Resolver + ?Sized>(
    evaluator: &mut ValueEvaluator<'expr, '_, '_, '_, R>,
    state: &mut State,
    argument: usize,
    value: RuntimeValue<'expr>,
) -> EvaluationResult<()> {
    match value {
        RuntimeValue::Areas(areas) => scan_reference(evaluator, state, argument, areas),
        RuntimeValue::Array(array) => {
            for (index, element) in array.cells.into_iter().enumerate() {
                evaluator.charge_cell_work(index)?;
                observe_element(evaluator, state, argument, index, element, false)?;
            }
            Ok(())
        },
        RuntimeValue::ScalarCell(area) => {
            let projected = evaluator.project_scalar(RuntimeValue::ScalarCell(area))?;
            observe_data(evaluator, state, argument, projected)
        },
        RuntimeValue::Empty => {
            evaluator.scalar.charge_work(1)?;
            state.push_number(evaluator, 0.0)
        },
        RuntimeValue::Missing => {
            evaluator.scalar.charge_work(1)?;
            state.generated(argument, 0, ScalarError::Value);
            Ok(())
        },
        RuntimeValue::Scalar(value) => {
            evaluator.scalar.charge_work(1)?;
            observe_working(evaluator, state, argument, value)
        },
        RuntimeValue::SourceReference => Err(EvaluationFailure::Unsupported(
            super::super::UnsupportedKind::Reference,
        )),
    }
}

fn scan_reference<'expr, R: Resolver + ?Sized>(
    evaluator: &mut ValueEvaluator<'expr, '_, '_, '_, R>,
    state: &mut State,
    argument: usize,
    areas: RuntimeAreaSet<'expr>,
) -> EvaluationResult<()> {
    let mut ordinal = 0usize;
    for area in areas.areas {
        for row in area.rect.row_start..area.rect.row_end {
            for column in area.rect.column_start..area.rect.column_end {
                evaluator.charge_cell_work(ordinal)?;
                let read = evaluator.read_reference_cell(area.sheet, row, column)?;
                let element = evaluator.read_to_element(read)?;
                observe_element(evaluator, state, argument, ordinal, element, true)?;
                ordinal = ordinal
                    .checked_add(1)
                    .ok_or(EvaluationFailure::InvalidExpression(
                        "order reference ordinal overflows",
                    ))?;
            }
        }
    }
    Ok(())
}

fn observe_element<'expr, R: Resolver + ?Sized>(
    evaluator: &mut ValueEvaluator<'expr, '_, '_, '_, R>,
    state: &mut State,
    argument: usize,
    ordinal: usize,
    element: RuntimeElement<'expr>,
    reference: bool,
) -> EvaluationResult<()> {
    match element {
        RuntimeElement::Empty if reference => Ok(()),
        RuntimeElement::Empty => state.push_number(evaluator, 0.0),
        RuntimeElement::Missing => {
            state.generated(argument, ordinal, ScalarError::Value);
            Ok(())
        },
        RuntimeElement::Present(WorkingValue::Number(value)) if value.is_finite() => {
            state.push_number(evaluator, value)
        },
        RuntimeElement::Present(WorkingValue::Number(_)) => {
            state.generated(argument, ordinal, ScalarError::Number);
            Ok(())
        },
        RuntimeElement::Present(WorkingValue::Logical(_)) if reference => Ok(()),
        RuntimeElement::Present(WorkingValue::Logical(value)) => {
            state.push_number(evaluator, if value { 1.0 } else { 0.0 })
        },
        RuntimeElement::Present(WorkingValue::Text(_)) if reference => Ok(()),
        RuntimeElement::Present(WorkingValue::Text(text)) => match super::super::to_number(
            WorkingValue::Text(TextValue::borrowed(text.text.as_ref())),
            &mut evaluator.scalar,
        )? {
            Ok(value) if value.is_finite() => state.push_number(evaluator, value),
            Ok(_) => {
                state.generated(argument, ordinal, ScalarError::Number);
                Ok(())
            },
            Err(error) => {
                state.generated(argument, ordinal, error);
                Ok(())
            },
        },
        RuntimeElement::Present(WorkingValue::Error(error)) => {
            state.formula(argument, ordinal, error);
            Ok(())
        },
        RuntimeElement::Present(WorkingValue::Complex(_)) => {
            state.generated(argument, ordinal, ScalarError::Value);
            Ok(())
        },
    }
}

fn observe_working<'expr, R: Resolver + ?Sized>(
    evaluator: &mut ValueEvaluator<'expr, '_, '_, '_, R>,
    state: &mut State,
    argument: usize,
    value: WorkingValue<'expr>,
) -> EvaluationResult<()> {
    match value {
        WorkingValue::Number(value) if value.is_finite() => state.push_number(evaluator, value),
        WorkingValue::Number(_) => {
            state.generated(argument, 0, ScalarError::Number);
            Ok(())
        },
        WorkingValue::Logical(value) => state.push_number(evaluator, if value { 1.0 } else { 0.0 }),
        WorkingValue::Text(value) => {
            match super::super::to_number(WorkingValue::Text(value), &mut evaluator.scalar)? {
                Ok(value) if value.is_finite() => state.push_number(evaluator, value),
                Ok(_) => {
                    state.generated(argument, 0, ScalarError::Number);
                    Ok(())
                },
                Err(error) => {
                    state.generated(argument, 0, error);
                    Ok(())
                },
            }
        },
        WorkingValue::Error(error) => {
            state.formula(argument, 0, error);
            Ok(())
        },
        WorkingValue::Complex(_) => {
            state.generated(argument, 0, ScalarError::Value);
            Ok(())
        },
    }
}

fn normalize_parameter<'expr, R: Resolver + ?Sized>(
    evaluator: &mut ValueEvaluator<'expr, '_, '_, '_, R>,
    value: RuntimeValue<'expr>,
) -> EvaluationResult<Parameter<'expr>> {
    match value {
        RuntimeValue::Array(array) => Ok(Parameter::Array(array)),
        RuntimeValue::ScalarCell(area) => {
            let projected = evaluator.project_scalar(RuntimeValue::ScalarCell(area))?;
            normalize_parameter(evaluator, projected)
        },
        RuntimeValue::Areas(areas) => {
            if evaluator.mode == super::Mode::Matrix && evaluator.projection.is_none() {
                let materialized = evaluator.materialize_for_array(RuntimeValue::Areas(areas))?;
                normalize_parameter(evaluator, materialized)
            } else {
                let projected = evaluator.project_scalar(RuntimeValue::Areas(areas))?;
                normalize_parameter(evaluator, projected)
            }
        },
        RuntimeValue::Empty => Ok(Parameter::Scalar(ParameterValue::Number(0.0))),
        RuntimeValue::Missing => Ok(Parameter::Scalar(ParameterValue::Generated(
            ScalarError::Value,
        ))),
        RuntimeValue::Scalar(value) => Ok(Parameter::Scalar(parameter_working(evaluator, value)?)),
        RuntimeValue::SourceReference => Err(EvaluationFailure::Unsupported(
            super::super::UnsupportedKind::Reference,
        )),
    }
}

fn parameter_working<'expr, R: Resolver + ?Sized>(
    evaluator: &mut ValueEvaluator<'expr, '_, '_, '_, R>,
    value: WorkingValue<'expr>,
) -> EvaluationResult<ParameterValue> {
    match value {
        WorkingValue::Number(value) if value.is_finite() => Ok(ParameterValue::Number(value)),
        WorkingValue::Number(_) => Ok(ParameterValue::Generated(ScalarError::Number)),
        WorkingValue::Logical(value) => Ok(ParameterValue::Number(if value { 1.0 } else { 0.0 })),
        WorkingValue::Text(value) => {
            match super::super::to_number(WorkingValue::Text(value), &mut evaluator.scalar)? {
                Ok(value) if value.is_finite() => Ok(ParameterValue::Number(value)),
                Ok(_) => Ok(ParameterValue::Generated(ScalarError::Number)),
                Err(error) => Ok(ParameterValue::Generated(error)),
            }
        },
        WorkingValue::Error(error) => Ok(ParameterValue::Formula(error)),
        WorkingValue::Complex(_) => Ok(ParameterValue::Generated(ScalarError::Value)),
    }
}

fn parameter_shape(
    parameters: &[Option<Parameter<'_>>; 3],
    lift_parameters: bool,
) -> EvaluationResult<Shape> {
    if !lift_parameters {
        return Shape::new(1, 1);
    }
    let mut shape = Shape::new(1, 1)?;
    for parameter in parameters.iter().flatten() {
        if let Parameter::Array(array) = parameter {
            shape = broadcast_shape(shape, array.shape).ok_or(
                EvaluationFailure::InvalidExpression("incompatible order parameter shapes"),
            )?;
        }
    }
    Ok(shape)
}

fn parameter_at<'expr, 'scalar, 'exec, 'position, R: Resolver + ?Sized>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    parameter: &Parameter<'expr>,
    shape: Shape,
    index: usize,
) -> EvaluationResult<ParameterValue> {
    match parameter {
        Parameter::Scalar(value) => Ok(*value),
        Parameter::Array(array) => match array_element_for(array, shape, index) {
            Some(element) => parameter_element(evaluator, element),
            None => Ok(ParameterValue::Formula(ScalarError::NotAvailable)),
        },
    }
}

fn parameter_element<'expr, 'scalar, 'exec, 'position, R: Resolver + ?Sized>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    element: &RuntimeElement<'expr>,
) -> EvaluationResult<ParameterValue> {
    match element {
        RuntimeElement::Empty => Ok(ParameterValue::Number(0.0)),
        RuntimeElement::Missing => Ok(ParameterValue::Generated(ScalarError::Value)),
        RuntimeElement::Present(WorkingValue::Number(value)) if value.is_finite() => {
            Ok(ParameterValue::Number(*value))
        },
        RuntimeElement::Present(WorkingValue::Number(_)) => {
            Ok(ParameterValue::Generated(ScalarError::Number))
        },
        RuntimeElement::Present(WorkingValue::Logical(value)) => {
            Ok(ParameterValue::Number(if *value { 1.0 } else { 0.0 }))
        },
        RuntimeElement::Present(WorkingValue::Text(value)) => match super::super::to_number(
            WorkingValue::Text(TextValue::borrowed(value.text.as_ref())),
            &mut evaluator.scalar,
        )? {
            Ok(value) if value.is_finite() => Ok(ParameterValue::Number(value)),
            Ok(_) => Ok(ParameterValue::Generated(ScalarError::Number)),
            Err(error) => Ok(ParameterValue::Generated(error)),
        },
        RuntimeElement::Present(WorkingValue::Error(error)) => Ok(ParameterValue::Formula(*error)),
        RuntimeElement::Present(WorkingValue::Complex(_)) => {
            Ok(ParameterValue::Generated(ScalarError::Value))
        },
    }
}

fn formula_error<'a>(error: ScalarError) -> RuntimeValue<'a> {
    RuntimeValue::Scalar(WorkingValue::Error(error))
}
