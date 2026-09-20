//! Resolver-aware paired statistics and simple regression.
//!
//! The scalar paired kernel owns the finite arithmetic.  This module owns the
//! value profile's ForceArray geometry, aligned cell admission, resolver
//! reads, and formula-error precedence.  Paired data are consumed in lockstep
//! and are never copied into a range-sized cell vector.

use super::super::paired::{self as kernel, ExactPaired, ForecastFit, PairNeed, PairedFunction};
use super::{
    CellRead, EvaluationFailure, EvaluationResult, Resolver, RuntimeArrayValue, RuntimeElement,
    RuntimeValue, ScalarError, Shape, TextValue, ValueEvaluator, WorkingValue, array_element_for,
    ensure_capacity,
};

/// Return whether `name` belongs to the normative paired-statistics family.
pub(super) fn is_paired_function(name: &str) -> bool {
    kernel::is_paired_function(name)
}

/// Keep complete ForceArray data descriptors in projected value evaluation.
/// The ordinary `FORECAST` query remains scalar and is lifted only when an
/// explicit Array expression supplies its shape.
pub(super) fn matrix_argument(name: &str, index: usize, node: super::super::Node<'_>) -> bool {
    let Some(function) = PairedFunction::from_name(name) else {
        return false;
    };
    if function.data_argument(index) {
        return true;
    }
    is_array_node(node)
}

/// A nested paired reducer consumes its ForceArray data arguments completely.
/// The ordinary FORECAST query is position-sensitive except for an explicit
/// literal Array, whose complete value is already independent of projection.
pub(super) fn criterion_full_argument(
    name: &str,
    index: usize,
    node: super::super::Node<'_>,
) -> bool {
    let Some(function) = PairedFunction::from_name(name) else {
        return false;
    };
    function.data_argument(index) || is_array_node(node)
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

/// The exact fit/error payload retained for projected FORECAST calls.
///
/// Formula and generated errors remain separate.  A query formula error must
/// outrank a data generated error, while a data formula error must outrank a
/// generated query conversion error.  Typed failures are returned directly
/// and are never represented here.
#[derive(Clone)]
pub(super) struct ForecastCacheEntry {
    pub(super) node: usize,
    pub(super) fit: Option<ForecastFit>,
    pub(super) formula_error: Option<ScalarError>,
    pub(super) generated_error: Option<ScalarError>,
}

#[derive(Clone, Copy, Debug)]
struct ErrorRecord {
    argument: usize,
    ordinal: usize,
    error: ScalarError,
}

#[derive(Clone, Copy, Debug)]
enum PairMember {
    Number(f64),
    Ignored,
    Formula(ScalarError),
    Generated(ScalarError),
}

struct PairState {
    kernel: ExactPaired,
    formula_error: Option<ErrorRecord>,
    generated_error: Option<ErrorRecord>,
}

impl PairState {
    fn new(need: PairNeed) -> Self {
        Self {
            kernel: ExactPaired::with_need(need),
            formula_error: None,
            generated_error: None,
        }
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

    fn finish_error(&self) -> Option<ScalarError> {
        self.formula_error
            .map(|record| record.error)
            .or_else(|| self.generated_error.map(|record| record.error))
    }

    fn data_error_payload(&self) -> (Option<ScalarError>, Option<ScalarError>) {
        (
            self.formula_error.map(|record| record.error),
            self.generated_error.map(|record| record.error),
        )
    }
}

/// Apply one paired statistic after its arguments have been visited.
pub(super) fn apply<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    node: super::super::Node<'expr>,
    name: &str,
    mut arguments: Vec<RuntimeValue<'expr>>,
) -> EvaluationResult<RuntimeValue<'expr>>
where
    R: Resolver + ?Sized,
{
    let Some(function) = PairedFunction::from_name(name) else {
        return Err(EvaluationFailure::Unsupported(
            super::super::UnsupportedKind::Function,
        ));
    };

    if !function.valid_arity(arguments.len()) {
        if let Some(error) = direct_formula_error(evaluator, &arguments)? {
            return Ok(formula_error(error));
        }
        return Ok(formula_error(ScalarError::Value));
    }

    let data_shapes = match data_shapes(function, &arguments) {
        Ok(shapes) => shapes,
        Err(error) => {
            if let Some(direct) = direct_formula_error(evaluator, &arguments)? {
                return Ok(formula_error(direct));
            }
            return Ok(formula_error(error));
        },
    };
    if let Some(error) = shape_error(function, data_shapes[0], data_shapes[1]) {
        if let Some(direct) = direct_formula_error(evaluator, &arguments)? {
            return Ok(formula_error(direct));
        }
        return Ok(formula_error(error));
    }

    if function == PairedFunction::Forecast && query_is_reference_list(&arguments) {
        if let Some(direct) = direct_formula_error(evaluator, &arguments)? {
            return Ok(formula_error(direct));
        }
        return Ok(formula_error(ScalarError::Value));
    }

    let query = if function == PairedFunction::Forecast {
        let query = std::mem::replace(
            arguments
                .get_mut(0)
                .ok_or(EvaluationFailure::InvalidExpression(
                    "FORECAST query argument is missing",
                ))?,
            RuntimeValue::Missing,
        );
        Some(normalize_query(evaluator, query)?)
    } else {
        None
    };

    let need = function.pair_need();
    let mut state = PairState::new(need);
    let left = arguments.get(function.data_argument_index(0)).ok_or(
        EvaluationFailure::InvalidExpression("paired left data argument is missing"),
    )?;
    let right = arguments.get(function.data_argument_index(1)).ok_or(
        EvaluationFailure::InvalidExpression("paired right data argument is missing"),
    )?;
    let shape = data_shapes[0];
    let cells = shape
        .cell_count()
        .ok_or(EvaluationFailure::InvalidExpression(
            "paired data cell count overflows",
        ))?;
    scan_pairs(evaluator, &mut state, left, right, shape, cells, function)?;

    if let Some(query) = query {
        let cacheable = function == PairedFunction::Forecast
            && evaluator.projection.is_some()
            && forecast_data_cacheable(evaluator, node, function)?;
        let result = finish_forecast(evaluator, &mut state, query)?;
        if cacheable {
            cache_forecast_state(evaluator, node, &state, result.fit)?;
        }
        return Ok(result.value);
    }

    finish_scalar(evaluator, function, state)
}

fn query_is_reference_list(arguments: &[RuntimeValue<'_>]) -> bool {
    matches!(
        arguments.first(),
        Some(RuntimeValue::Areas(areas)) if areas.is_list || areas.areas.len() != 1
    )
}

fn data_shapes(
    function: PairedFunction,
    arguments: &[RuntimeValue<'_>],
) -> Result<[Shape; 2], ScalarError> {
    let left_index = function.data_argument_index(0);
    let right_index = function.data_argument_index(1);
    let left = shape_of_data(arguments.get(left_index).ok_or(ScalarError::Value)?)?;
    let right = shape_of_data(arguments.get(right_index).ok_or(ScalarError::Value)?)?;
    Ok([left, right])
}

fn shape_of_data(value: &RuntimeValue<'_>) -> Result<Shape, ScalarError> {
    match value {
        RuntimeValue::Array(array) => Ok(array.shape),
        RuntimeValue::Areas(areas) => {
            if areas.is_list || areas.areas.len() != 1 {
                return Err(ScalarError::Value);
            }
            let area = areas.areas.first().ok_or(ScalarError::Value)?;
            Shape::new(area.rect.rows(), area.rect.columns()).map_err(|_| ScalarError::Value)
        },
        RuntimeValue::Empty
        | RuntimeValue::Missing
        | RuntimeValue::Scalar(_)
        | RuntimeValue::ScalarCell(_) => Shape::new(1, 1).map_err(|_| ScalarError::Value),
    }
}

fn shape_error(function: PairedFunction, left: Shape, right: Shape) -> Option<ScalarError> {
    if left == right {
        return None;
    }
    if function == PairedFunction::Rsq && left.cell_count() != right.cell_count() {
        Some(ScalarError::NotAvailable)
    } else {
        Some(ScalarError::Value)
    }
}

fn scan_pairs<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    state: &mut PairState,
    left: &RuntimeValue<'expr>,
    right: &RuntimeValue<'expr>,
    shape: Shape,
    cells: usize,
    function: PairedFunction,
) -> EvaluationResult<()>
where
    R: Resolver + ?Sized,
{
    for index in 0..cells {
        let left_member = read_member(evaluator, left, shape, index)?;
        let right_member = read_member(evaluator, right, shape, index)?;
        observe_pair(
            evaluator,
            state,
            function,
            function.data_argument_index(0),
            function.data_argument_index(1),
            index,
            left_member,
            right_member,
        )?;
    }
    Ok(())
}

fn read_member<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    value: &RuntimeValue<'expr>,
    shape: Shape,
    index: usize,
) -> EvaluationResult<PairMember>
where
    R: Resolver + ?Sized,
{
    evaluator.charge_cell_work(index)?;
    match value {
        RuntimeValue::Areas(areas) => {
            let area = areas
                .areas
                .first()
                .ok_or(EvaluationFailure::InvalidExpression(
                    "paired reference has no rectangular area",
                ))?;
            let columns = area.rect.columns();
            let row = index / columns;
            let column = index % columns;
            if row >= area.rect.rows() || column >= columns {
                return Err(EvaluationFailure::InvalidExpression(
                    "paired reference index is out of bounds",
                ));
            }
            let read = evaluator.read_reference_cell(
                area.sheet,
                area.rect.row_start + row,
                area.rect.column_start + column,
            )?;
            classify_data_read(evaluator, read)
        },
        RuntimeValue::ScalarCell(area) => {
            if shape.rows() != 1 || shape.columns() != 1 {
                return Err(EvaluationFailure::InvalidExpression(
                    "paired scalar-cell shape is not one by one",
                ));
            }
            let read = evaluator.read_reference_cell(
                area.sheet,
                area.rect.row_start,
                area.rect.column_start,
            )?;
            classify_data_read(evaluator, read)
        },
        RuntimeValue::Array(array) => {
            let element = array_element_for(array, shape, index).ok_or(
                EvaluationFailure::InvalidExpression("paired array index is out of bounds"),
            )?;
            charge_element_text(evaluator, element)?;
            Ok(classify_element(element))
        },
        RuntimeValue::Empty => Ok(PairMember::Ignored),
        RuntimeValue::Missing => Ok(PairMember::Generated(ScalarError::Value)),
        RuntimeValue::Scalar(value) => {
            charge_working_text(evaluator, value)?;
            Ok(classify_working(value))
        },
    }
}

fn charge_element_text<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    element: &RuntimeElement<'expr>,
) -> EvaluationResult<()>
where
    R: Resolver + ?Sized,
{
    if let RuntimeElement::Present(value) = element {
        charge_working_text(evaluator, value)?;
    }
    Ok(())
}

fn charge_working_text<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    value: &WorkingValue<'expr>,
) -> EvaluationResult<()>
where
    R: Resolver + ?Sized,
{
    if let WorkingValue::Text(text) = value {
        evaluator.scalar.charge_bytes(text.text.len())?;
    }
    Ok(())
}

fn classify_data_read<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    read: CellRead<'expr>,
) -> EvaluationResult<PairMember>
where
    R: Resolver + ?Sized,
{
    // `read_to_element` intentionally maps provider non-finite numbers to a
    // scalar error for generic value reducers.  Paired reducers must keep
    // this generated #NUM! distinct from a source formula error.
    if matches!(read, CellRead::Number(value) if !value.is_finite()) {
        return Ok(PairMember::Generated(ScalarError::Number));
    }
    let element = evaluator.read_to_element(read)?;
    charge_element_text(evaluator, &element)?;
    Ok(classify_element(&element))
}

fn classify_element(element: &RuntimeElement<'_>) -> PairMember {
    match element {
        RuntimeElement::Empty => PairMember::Ignored,
        RuntimeElement::Missing => PairMember::Generated(ScalarError::Value),
        RuntimeElement::Present(WorkingValue::Number(value)) if value.is_finite() => {
            PairMember::Number(*value)
        },
        RuntimeElement::Present(WorkingValue::Number(_)) => {
            PairMember::Generated(ScalarError::Number)
        },
        RuntimeElement::Present(WorkingValue::Text(_))
        | RuntimeElement::Present(WorkingValue::Logical(_)) => PairMember::Ignored,
        RuntimeElement::Present(WorkingValue::Error(error)) => PairMember::Formula(*error),
        RuntimeElement::Present(WorkingValue::Complex(_)) => {
            PairMember::Generated(ScalarError::Value)
        },
    }
}

fn classify_working(value: &WorkingValue<'_>) -> PairMember {
    match value {
        WorkingValue::Number(value) if value.is_finite() => PairMember::Number(*value),
        WorkingValue::Number(_) => PairMember::Generated(ScalarError::Number),
        WorkingValue::Text(_) | WorkingValue::Logical(_) => PairMember::Ignored,
        WorkingValue::Error(error) => PairMember::Formula(*error),
        WorkingValue::Complex(_) => PairMember::Generated(ScalarError::Value),
    }
}

fn observe_pair<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    state: &mut PairState,
    function: PairedFunction,
    left_argument: usize,
    right_argument: usize,
    ordinal: usize,
    left: PairMember,
    right: PairMember,
) -> EvaluationResult<()>
where
    R: Resolver + ?Sized,
{
    match left {
        PairMember::Formula(error) => state.formula(left_argument, ordinal, error),
        PairMember::Generated(error) => state.generated(left_argument, ordinal, error),
        PairMember::Number(_) | PairMember::Ignored => {},
    }
    match right {
        PairMember::Formula(error) => state.formula(right_argument, ordinal, error),
        PairMember::Generated(error) => state.generated(right_argument, ordinal, error),
        PairMember::Number(_) | PairMember::Ignored => {},
    }
    if state.formula_error.is_some() || state.generated_error.is_some() {
        return Ok(());
    }
    let (PairMember::Number(left), PairMember::Number(right)) = (left, right) else {
        return Ok(());
    };
    let (x, y) = if function.x_index() == left_argument {
        (left, right)
    } else {
        (right, left)
    };
    evaluator
        .scalar
        .charge_work(state.kernel.push_work_units())?;
    if let Err(error) = state.kernel.push_pair(x, y) {
        state.generated(left_argument, ordinal, map_pair_error(function, error));
    }
    Ok(())
}

fn finish_scalar<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    function: PairedFunction,
    state: PairState,
) -> EvaluationResult<RuntimeValue<'expr>>
where
    R: Resolver + ?Sized,
{
    if let Some(error) = state.finish_error() {
        return Ok(formula_error(error));
    }
    evaluator
        .scalar
        .charge_work(state.kernel.finish_work_units())?;
    let result = match function {
        PairedFunction::Correl | PairedFunction::Pearson => state.kernel.finish_correl(),
        PairedFunction::Covar => state.kernel.finish_covar(),
        PairedFunction::Rsq => state.kernel.finish_rsq(),
        PairedFunction::Slope => state.kernel.finish_slope(),
        PairedFunction::Intercept => state.kernel.finish_intercept(),
        PairedFunction::Steyx => state.kernel.finish_steyx(),
        PairedFunction::Forecast => {
            return Err(EvaluationFailure::InvalidExpression(
                "FORECAST reached scalar paired finish",
            ));
        },
    };
    Ok(match result {
        Ok(value) => number(value),
        Err(error) => formula_error(map_pair_error(function, error)),
    })
}

struct QueryValue<'expr> {
    value: RuntimeValue<'expr>,
    shape: Shape,
}

fn normalize_query<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    value: RuntimeValue<'expr>,
) -> EvaluationResult<QueryValue<'expr>>
where
    R: Resolver + ?Sized,
{
    // A scalar or projected query reference is implicitly intersected at the
    // demanded coordinate.  Explicit/computed arrays retain their complete
    // shape, so FORECAST can stream their cells into the output without a
    // range-sized materialization buffer.
    let value = if (evaluator.mode == super::Mode::Scalar || evaluator.projection.is_some())
        && !matches!(value, RuntimeValue::Array(_))
    {
        match value {
            RuntimeValue::Areas(areas) => project_query_area(evaluator, &areas)?,
            // A missing query is a generated conversion failure.  Keep the
            // marker until the data scan completes so typed data failures can
            // still supersede it.
            RuntimeValue::Missing => RuntimeValue::Missing,
            value => evaluator.project_scalar(value)?,
        }
    } else {
        value
    };
    let shape = match &value {
        RuntimeValue::Array(array) => array.shape,
        RuntimeValue::Areas(areas) if areas.is_list || areas.areas.len() != 1 => {
            return Ok(QueryValue {
                value: formula_error(ScalarError::Value),
                shape: Shape::new(1, 1)?,
            });
        },
        RuntimeValue::Areas(areas) => {
            let area = areas
                .areas
                .first()
                .ok_or(EvaluationFailure::InvalidExpression(
                    "FORECAST query area is missing",
                ))?;
            Shape::new(area.rect.rows(), area.rect.columns())?
        },
        RuntimeValue::Empty
        | RuntimeValue::Missing
        | RuntimeValue::Scalar(_)
        | RuntimeValue::ScalarCell(_) => Shape::new(1, 1)?,
    };
    Ok(QueryValue { value, shape })
}

fn project_query_area<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    areas: &super::RuntimeAreaSet<'expr>,
) -> EvaluationResult<RuntimeValue<'expr>>
where
    R: Resolver + ?Sized,
{
    match evaluator.scalar_reference_cell(areas)? {
        Ok((area, row, column)) => {
            let mut selected = area;
            selected.rect = super::Rect::cell(row, column)?;
            Ok(RuntimeValue::ScalarCell(selected))
        },
        Err(error) => Ok(RuntimeValue::Scalar(WorkingValue::Error(error))),
    }
}

struct ForecastResult<'expr> {
    value: RuntimeValue<'expr>,
    fit: Option<ForecastFit>,
}

fn finish_forecast<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    state: &mut PairState,
    query: QueryValue<'expr>,
) -> EvaluationResult<ForecastResult<'expr>>
where
    R: Resolver + ?Sized,
{
    let (data_formula, data_generated) = state.data_error_payload();
    let fit = if data_formula.is_none() && data_generated.is_none() {
        evaluator
            .scalar
            .charge_work(state.kernel.prepare_forecast_work_units())?;
        match state.kernel.prepare_forecast() {
            Ok(fit) => Some(fit),
            Err(error) => {
                state.generated(
                    PairedFunction::Forecast.data_argument_index(0),
                    0,
                    map_pair_error(PairedFunction::Forecast, error),
                );
                None
            },
        }
    } else {
        None
    };
    let (data_formula, data_generated) = state.data_error_payload();
    let value =
        finish_forecast_values(evaluator, query, fit.as_ref(), data_formula, data_generated)?;
    Ok(ForecastResult { value, fit })
}

fn finish_forecast_values<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    query: QueryValue<'expr>,
    fit: Option<&ForecastFit>,
    data_formula: Option<ScalarError>,
    data_generated: Option<ScalarError>,
) -> EvaluationResult<RuntimeValue<'expr>>
where
    R: Resolver + ?Sized,
{
    let cells = query
        .shape
        .cell_count()
        .ok_or(EvaluationFailure::InvalidExpression(
            "FORECAST query cell count overflows",
        ))?;
    if cells == 1 && !matches!(query.value, RuntimeValue::Array(_)) {
        let query_member = query_member(evaluator, &query.value, query.shape, 0)?;
        let value = finish_query_value(evaluator, query_member, fit, data_formula, data_generated)?;
        return Ok(value);
    }

    let (mut output, reservation) = evaluator.new_element_vec(cells)?;
    for index in 0..cells {
        let member = query_member(evaluator, &query.value, query.shape, index)?;
        let value = finish_query_value(evaluator, member, fit, data_formula, data_generated)?;
        output.push(match value {
            RuntimeValue::Scalar(value) => RuntimeElement::Present(value),
            RuntimeValue::Empty => RuntimeElement::Empty,
            RuntimeValue::Missing => RuntimeElement::Missing,
            RuntimeValue::Array(_) | RuntimeValue::Areas(_) | RuntimeValue::ScalarCell(_) => {
                RuntimeElement::Present(WorkingValue::Error(ScalarError::Value))
            },
        });
    }
    Ok(RuntimeValue::Array(RuntimeArrayValue {
        shape: query.shape,
        cells: output,
        _reservation: reservation,
        origin: None,
        preserve_scalar_result: false,
    }))
}

fn query_member<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    value: &RuntimeValue<'expr>,
    shape: Shape,
    index: usize,
) -> EvaluationResult<PairMember>
where
    R: Resolver + ?Sized,
{
    evaluator.charge_cell_work(index)?;
    match value {
        RuntimeValue::Array(array) => {
            let element = array_element_for(array, shape, index).ok_or(
                EvaluationFailure::InvalidExpression("FORECAST query index is out of bounds"),
            )?;
            classify_query_element(evaluator, element)
        },
        RuntimeValue::Scalar(value) => classify_query_working(evaluator, value),
        RuntimeValue::Empty => Ok(PairMember::Number(0.0)),
        RuntimeValue::Missing => Ok(PairMember::Generated(ScalarError::Value)),
        RuntimeValue::ScalarCell(area) => {
            let read = evaluator.read_reference_cell(
                area.sheet,
                area.rect.row_start,
                area.rect.column_start,
            )?;
            classify_query_read(evaluator, read)
        },
        RuntimeValue::Areas(areas) => {
            if areas.is_list || areas.areas.len() != 1 {
                return Ok(PairMember::Generated(ScalarError::Value));
            }
            let area = &areas.areas[0];
            let columns = area.rect.columns();
            let row = index / columns;
            let column = index % columns;
            if row >= area.rect.rows() {
                return Err(EvaluationFailure::InvalidExpression(
                    "FORECAST query area index is out of bounds",
                ));
            }
            let read = evaluator.read_reference_cell(
                area.sheet,
                area.rect.row_start + row,
                area.rect.column_start + column,
            )?;
            classify_query_read(evaluator, read)
        },
    }
}

fn classify_query_read<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    read: CellRead<'expr>,
) -> EvaluationResult<PairMember>
where
    R: Resolver + ?Sized,
{
    if matches!(read, CellRead::Number(value) if !value.is_finite()) {
        return Ok(PairMember::Generated(ScalarError::Number));
    }
    let element = evaluator.read_to_element(read)?;
    classify_query_element(evaluator, &element)
}

fn classify_query_element<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    element: &RuntimeElement<'expr>,
) -> EvaluationResult<PairMember>
where
    R: Resolver + ?Sized,
{
    match element {
        RuntimeElement::Empty => Ok(PairMember::Number(0.0)),
        RuntimeElement::Missing => Ok(PairMember::Generated(ScalarError::Value)),
        RuntimeElement::Present(value) => classify_query_working(evaluator, value),
    }
}

fn classify_query_working<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    value: &WorkingValue<'expr>,
) -> EvaluationResult<PairMember>
where
    R: Resolver + ?Sized,
{
    match value {
        WorkingValue::Text(text) => {
            let converted = super::super::to_number(
                WorkingValue::Text(TextValue::borrowed(text.text.as_ref())),
                &mut evaluator.scalar,
            )?;
            Ok(match converted {
                Ok(value) if value.is_finite() => PairMember::Number(value),
                Ok(_) => PairMember::Generated(ScalarError::Number),
                Err(error) => PairMember::Generated(error),
            })
        },
        WorkingValue::Number(value) if value.is_finite() => Ok(PairMember::Number(*value)),
        WorkingValue::Number(_) => Ok(PairMember::Generated(ScalarError::Number)),
        WorkingValue::Logical(value) => Ok(PairMember::Number(f64::from(*value))),
        WorkingValue::Error(error) => Ok(PairMember::Formula(*error)),
        WorkingValue::Complex(_) => Ok(PairMember::Generated(ScalarError::Value)),
    }
}

fn finish_query_value<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    query: PairMember,
    fit: Option<&ForecastFit>,
    data_formula: Option<ScalarError>,
    data_generated: Option<ScalarError>,
) -> EvaluationResult<RuntimeValue<'expr>>
where
    R: Resolver + ?Sized,
{
    let query_formula = match query {
        PairMember::Formula(error) => Some(error),
        _ => None,
    };
    let query_generated = match query {
        PairMember::Generated(error) => Some(error),
        _ => None,
    };
    if let Some(error) = query_formula.or(data_formula) {
        return Ok(formula_error(error));
    }
    if let Some(error) = query_generated.or(data_generated) {
        return Ok(formula_error(error));
    }
    let PairMember::Number(query) = query else {
        return Ok(formula_error(ScalarError::Value));
    };
    let fit = fit.ok_or(EvaluationFailure::InvalidExpression(
        "FORECAST fit is missing",
    ))?;
    evaluator
        .scalar
        .charge_work(ForecastFit::query_work_units())?;
    Ok(match fit.finish_forecast(query) {
        Ok(value) => number(value),
        Err(error) => formula_error(map_pair_error(PairedFunction::Forecast, error)),
    })
}

fn forecast_data_cacheable<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    node: super::super::Node<'expr>,
    function: PairedFunction,
) -> EvaluationResult<bool>
where
    R: Resolver + ?Sized,
{
    for index in 0..node.child_count() {
        if !function.data_argument(index) {
            continue;
        }
        let child = node
            .child(index)
            .ok_or(EvaluationFailure::InvalidExpression(
                "paired cache data argument is missing",
            ))?;
        if !evaluator.cacheable_conditional_criterion(child)? {
            return Ok(false);
        }
    }
    Ok(true)
}

fn cache_forecast_state<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    node: super::super::Node<'expr>,
    state: &PairState,
    fit: Option<ForecastFit>,
) -> EvaluationResult<()>
where
    R: Resolver + ?Sized,
{
    let (formula_error, generated_error) = state.data_error_payload();
    let node_key = node.arena_index();
    let (index, present) = forecast_cache_position(evaluator, node_key)?;
    let entry = ForecastCacheEntry {
        node: node_key,
        fit,
        formula_error,
        generated_error,
    };
    if present {
        evaluator.forecast_cache[index] = entry;
        return Ok(());
    }
    ensure_capacity(
        &mut evaluator.forecast_cache,
        &mut evaluator.forecast_cache_reservation,
        1,
        evaluator.limits.scalar.max_stack_entries,
        evaluator.execution,
        &evaluator.storage_budget,
        "formula value FORECAST fit cache",
    )?;
    let moved = evaluator.forecast_cache.len().saturating_sub(index);
    evaluator
        .scalar
        .charge_work(u64::try_from(moved).unwrap_or(u64::MAX))?;
    evaluator.forecast_cache.insert(index, entry);
    Ok(())
}

fn forecast_cache_position<R: Resolver + ?Sized>(
    evaluator: &mut ValueEvaluator<'_, '_, '_, '_, R>,
    node: usize,
) -> EvaluationResult<(usize, bool)> {
    let mut first = 0;
    let mut last = evaluator.forecast_cache.len();
    while first < last {
        let middle = first + (last - first) / 2;
        evaluator.scalar.charge_work(1)?;
        match evaluator.forecast_cache[middle].node.cmp(&node) {
            std::cmp::Ordering::Less => first = middle + 1,
            std::cmp::Ordering::Equal => return Ok((middle, true)),
            std::cmp::Ordering::Greater => last = middle,
        }
    }
    Ok((first, false))
}

/// Handle a projected FORECAST call whose invariant data fit was cached.
pub(super) fn visit_cached_forecast<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    node: super::super::Node<'expr>,
    name: &str,
) -> EvaluationResult<bool>
where
    R: Resolver + ?Sized,
{
    if !name.eq_ignore_ascii_case("FORECAST") || evaluator.projection.is_none() {
        return Ok(false);
    }
    let Some(function) = PairedFunction::from_name(name) else {
        return Ok(false);
    };
    if !function.valid_arity(node.child_count())
        || !forecast_data_cacheable(evaluator, node, function)?
    {
        return Ok(false);
    }
    let (_, present) = forecast_cache_position(evaluator, node.arena_index())?;
    if !present {
        return Ok(false);
    }
    let query = node.child(0).ok_or(EvaluationFailure::InvalidExpression(
        "FORECAST query argument is missing",
    ))?;
    evaluator.push_frame(super::ValueFrame::ApplyForecastQuery(node))?;
    evaluator.push_frame(super::ValueFrame::VisitArgument(query))?;
    Ok(true)
}

/// Apply a cached FORECAST fit to the query value left by the query frame.
pub(super) fn apply_cached_forecast<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    node: super::super::Node<'expr>,
) -> EvaluationResult<RuntimeValue<'expr>>
where
    R: Resolver + ?Sized,
{
    let query = evaluator.pop_value()?;
    let (index, present) = forecast_cache_position(evaluator, node.arena_index())?;
    if !present {
        return Err(EvaluationFailure::InvalidExpression(
            "FORECAST cache entry disappeared",
        ));
    }
    let has_fit = evaluator.forecast_cache[index].fit.is_some();
    if has_fit {
        // The cached fit is a fixed-width exact payload.  Its bounded copy is
        // charged before cloning so cache hits cannot bypass the work budget.
        evaluator
            .scalar
            .charge_work(ForecastFit::query_work_units())?;
    }
    let entry = evaluator.forecast_cache[index].clone();
    let query = normalize_query(evaluator, query)?;
    let result = finish_forecast_values(
        evaluator,
        query,
        entry.fit.as_ref(),
        entry.formula_error,
        entry.generated_error,
    )?;
    Ok(result)
}

fn direct_formula_error<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    arguments: &[RuntimeValue<'expr>],
) -> EvaluationResult<Option<ScalarError>>
where
    R: Resolver + ?Sized,
{
    for argument in arguments {
        match argument {
            RuntimeValue::Scalar(WorkingValue::Error(error)) => {
                evaluator.scalar.charge_work(1)?;
                return Ok(Some(*error));
            },
            RuntimeValue::Array(array) => {
                for (index, element) in array.cells.iter().enumerate() {
                    evaluator.charge_cell_work(index)?;
                    if let RuntimeElement::Present(WorkingValue::Error(error)) = element {
                        return Ok(Some(*error));
                    }
                }
            },
            RuntimeValue::Empty
            | RuntimeValue::Missing
            | RuntimeValue::Scalar(_)
            | RuntimeValue::ScalarCell(_)
            | RuntimeValue::Areas(_) => {},
        }
    }
    Ok(None)
}

fn formula_error<'a>(error: ScalarError) -> RuntimeValue<'a> {
    RuntimeValue::Scalar(WorkingValue::Error(error))
}

fn number<'a>(value: f64) -> RuntimeValue<'a> {
    RuntimeValue::Scalar(WorkingValue::Number(value))
}

fn map_pair_error(function: PairedFunction, error: kernel::PairError) -> ScalarError {
    match error {
        kernel::PairError::Empty if function == PairedFunction::Rsq => ScalarError::NotAvailable,
        kernel::PairError::Empty => ScalarError::Value,
        kernel::PairError::MinimumCount => ScalarError::Value,
        kernel::PairError::ZeroVariance => ScalarError::DivisionByZero,
        kernel::PairError::Number => ScalarError::Number,
    }
}
