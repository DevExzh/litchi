//! Bounded conditional aggregate evaluation for the value VM.
//!
//! The six conditional aggregates consume references as references. Their
//! range arguments stay in the VM as `RuntimeAreaSet` values and are walked
//! one cell at a time. This keeps provider reads bounded and lets criteria
//! reject a row before a selected SUM/AVERAGE value is read. Criterion
//! parsing and matching are shared with the database functions.

use super::super::numerics::{NumericAggregate, NumericOperation};
use super::criteria::{CriterionMatcher, CriterionValue, compile_criterion, criterion_matches};
use super::{
    CellRead, EvaluationFailure, EvaluationResult, Resolver, Resource, RuntimeArea, RuntimeAreaSet,
    RuntimeElement, RuntimeValue, ScalarError, SheetRef, ValueEvaluator, WorkingValue,
    ensure_capacity,
};

/// Return whether `name` is one of the six conditional aggregates.
pub(super) fn is_conditional_function(name: &str) -> bool {
    name.eq_ignore_ascii_case("SUMIF")
        || name.eq_ignore_ascii_case("SUMIFS")
        || name.eq_ignore_ascii_case("COUNTIF")
        || name.eq_ignore_ascii_case("COUNTIFS")
        || name.eq_ignore_ascii_case("AVERAGEIF")
        || name.eq_ignore_ascii_case("AVERAGEIFS")
}

/// Return whether one argument is a reference/range argument for `name`.
pub(super) fn is_range_argument(name: &str, index: usize, count: usize) -> bool {
    if name.eq_ignore_ascii_case("SUMIF")
        || name.eq_ignore_ascii_case("COUNTIF")
        || name.eq_ignore_ascii_case("AVERAGEIF")
    {
        return index == 0 || index == 2;
    }
    if name.eq_ignore_ascii_case("SUMIFS") || name.eq_ignore_ascii_case("AVERAGEIFS") {
        return index == 0 || (index >= 1 && index % 2 == 1 && index < count);
    }
    if name.eq_ignore_ascii_case("COUNTIFS") {
        return index % 2 == 0 && index < count;
    }
    false
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Function {
    SumIf,
    SumIfs,
    CountIf,
    CountIfs,
    AverageIf,
    AverageIfs,
}

#[derive(Clone, Copy)]
enum CompiledCriterion<'argument, 'expr> {
    Argument(CriterionMatcher<'argument>),
    Reference(CriterionMatcher<'expr>),
}

impl<'argument, 'expr> CompiledCriterion<'argument, 'expr> {
    fn matches<'scalar, 'exec, 'position, R>(
        self,
        evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
        candidate: CriterionValue<'_>,
    ) -> EvaluationResult<Result<bool, ScalarError>>
    where
        R: Resolver + ?Sized,
    {
        match self {
            Self::Argument(matcher) => criterion_matches(evaluator, matcher, candidate),
            Self::Reference(matcher) => criterion_matches(evaluator, matcher, candidate),
        }
    }
}

impl Function {
    fn from_name(name: &str) -> Option<Self> {
        Some(if name.eq_ignore_ascii_case("SUMIF") {
            Self::SumIf
        } else if name.eq_ignore_ascii_case("SUMIFS") {
            Self::SumIfs
        } else if name.eq_ignore_ascii_case("COUNTIF") {
            Self::CountIf
        } else if name.eq_ignore_ascii_case("COUNTIFS") {
            Self::CountIfs
        } else if name.eq_ignore_ascii_case("AVERAGEIF") {
            Self::AverageIf
        } else if name.eq_ignore_ascii_case("AVERAGEIFS") {
            Self::AverageIfs
        } else {
            return None;
        })
    }

    const fn is_sum(self) -> bool {
        matches!(self, Self::SumIf | Self::SumIfs)
    }

    const fn is_count(self) -> bool {
        matches!(self, Self::CountIf | Self::CountIfs)
    }

    const fn is_average(self) -> bool {
        matches!(self, Self::AverageIf | Self::AverageIfs)
    }
}

/// Apply one conditional aggregate after the parent VM has evaluated its
/// arguments. The arguments remain borrowed so compiled criteria can borrow
/// text without copying or parsing it once per row.
pub(super) fn apply<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    name: &str,
    arguments: &[RuntimeValue<'expr>],
) -> EvaluationResult<RuntimeValue<'expr>>
where
    R: Resolver + ?Sized,
{
    let Some(function) = Function::from_name(name) else {
        return Err(EvaluationFailure::Unsupported(
            super::super::UnsupportedKind::Function,
        ));
    };
    if let Some(error) = direct_formula_error(evaluator, arguments)? {
        return Ok(formula_error(error));
    }
    if function.is_sum() || function.is_count() || function.is_average() {
        if matches!(
            function,
            Function::SumIf | Function::CountIf | Function::AverageIf
        ) {
            apply_if(evaluator, function, arguments)
        } else {
            apply_ifs(evaluator, function, arguments)
        }
    } else {
        Err(EvaluationFailure::Unsupported(
            super::super::UnsupportedKind::Function,
        ))
    }
}

fn apply_if<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    function: Function,
    arguments: &[RuntimeValue<'expr>],
) -> EvaluationResult<RuntimeValue<'expr>>
where
    R: Resolver + ?Sized,
{
    let valid_arity = if function == Function::CountIf {
        arguments.len() == 2
    } else {
        (2..=3).contains(&arguments.len())
    };
    if !valid_arity {
        return Ok(formula_error(ScalarError::Value));
    }
    let criteria_range = match reference_argument(arguments, 0) {
        Ok(range) => range,
        Err(error) => return Ok(formula_error(error)),
    };
    if function == Function::AverageIf && criteria_range.is_list {
        return Ok(formula_error(ScalarError::Value));
    }
    if function.is_sum()
        && arguments.len() == 3
        && criteria_range.is_list
        && criteria_range.records.len() > 1
    {
        return Ok(formula_error(ScalarError::Value));
    }
    let matcher = match compile_criterion_argument(evaluator, &arguments[1])? {
        Ok(matcher) => matcher,
        Err(error) => return Ok(formula_error(error)),
    };
    let selected = if arguments.len() == 3 {
        match reference_argument(arguments, 2) {
            Ok(range) if !range.is_list => Some(range),
            Ok(_) => return Ok(formula_error(ScalarError::Value)),
            Err(error) => return Ok(formula_error(error)),
        }
    } else {
        None
    };
    if let Some(selected) = selected {
        let shape = match compatible_shape(evaluator, &[criteria_range])? {
            Ok(shape) => shape,
            Err(error) => return Ok(formula_error(error)),
        };
        if let Err(error) =
            validate_projected_geometry(evaluator, selected, shape, function == Function::SumIf)?
        {
            return Ok(formula_error(error));
        }
    }

    let mut state = ConditionalState::new(function);
    let mut cell_index = 0usize;
    for (plane, area) in criteria_range.areas.iter().enumerate() {
        for row in 0..area.rect.rows() {
            for column in 0..area.rect.columns() {
                let (criterion, candidate_value) =
                    match matches_area_cell(evaluator, *area, row, column, matcher, cell_index)? {
                        Ok(value) => value,
                        Err(error) => {
                            state.observe_formula_error(error);
                            cell_index = next_index(cell_index)?;
                            continue;
                        },
                    };
                if !criterion {
                    cell_index = next_index(cell_index)?;
                    continue;
                }
                let value = if let Some(selected) = selected {
                    match projected_cell(
                        evaluator,
                        selected,
                        plane,
                        row,
                        column,
                        function == Function::SumIf,
                        cell_index,
                    )? {
                        Ok(Some(value)) => value,
                        Ok(None) => {
                            // SUMIF explicitly clips a constructed sum range
                            // beyond sheet bounds. AVERAGEIF reports #REF!.
                            if function == Function::AverageIf {
                                return Ok(formula_error(ScalarError::Reference));
                            }
                            cell_index = next_index(cell_index)?;
                            continue;
                        },
                        Err(error) => {
                            state.observe_formula_error(error);
                            cell_index = next_index(cell_index)?;
                            continue;
                        },
                    }
                } else {
                    candidate_value
                };
                state.observe_value(value, function.is_count())?;
                cell_index = next_index(cell_index)?;
            }
        }
    }
    state.finish(function)
}

fn apply_ifs<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    function: Function,
    arguments: &[RuntimeValue<'expr>],
) -> EvaluationResult<RuntimeValue<'expr>>
where
    R: Resolver + ?Sized,
{
    let valid_arity = if function.is_count() {
        arguments.len() >= 2 && arguments.len().is_multiple_of(2)
    } else {
        arguments.len() >= 3 && (arguments.len() - 1).is_multiple_of(2)
    };
    if !valid_arity {
        return Ok(formula_error(ScalarError::Value));
    }
    let target = if function.is_count() {
        None
    } else {
        match reference_argument(arguments, 0) {
            Ok(range) if !range.is_list => Some(range),
            Ok(_) => return Ok(formula_error(ScalarError::Value)),
            Err(error) => return Ok(formula_error(error)),
        }
    };
    let first_range_index = usize::from(target.is_some());
    let range_count = (arguments.len() - first_range_index) / 2;
    if range_count == 0 {
        return Ok(formula_error(ScalarError::Value));
    }

    let mut range_reservation = None;
    let mut ranges = Vec::new();
    ensure_capacity(
        &mut ranges,
        &mut range_reservation,
        range_count + usize::from(target.is_some()),
        evaluator.limits.scalar.max_stack_entries,
        evaluator.execution,
        &evaluator.storage_budget,
        "formula conditional aggregate ranges",
    )?;
    let mut matcher_reservation = None;
    let mut matchers = Vec::new();
    ensure_capacity(
        &mut matchers,
        &mut matcher_reservation,
        range_count,
        evaluator.limits.scalar.max_stack_entries,
        evaluator.execution,
        &evaluator.storage_budget,
        "formula conditional aggregate criteria",
    )?;
    for pair in 0..range_count {
        let range_index = first_range_index + pair * 2;
        let range = match reference_argument(arguments, range_index) {
            Ok(range) if !range.is_list => range,
            Ok(_) => return Ok(formula_error(ScalarError::Value)),
            Err(error) => return Ok(formula_error(error)),
        };
        let matcher = match compile_criterion_argument(evaluator, &arguments[range_index + 1])? {
            Ok(matcher) => matcher,
            Err(error) => return Ok(formula_error(error)),
        };
        ranges.push(range);
        matchers.push(matcher);
    }
    if let Some(target) = target {
        ranges.push(target);
    }
    let shape = match compatible_shape(evaluator, &ranges)? {
        Ok(shape) => shape,
        Err(error) => return Ok(formula_error(error)),
    };

    let mut state = ConditionalState::new(function);
    let mut cell_index = 0usize;
    for plane in 0..shape.planes {
        for row in 0..shape.rows {
            for column in 0..shape.columns {
                let mut matches = true;
                for pair in 0..range_count {
                    let value = read_area_cell(
                        evaluator,
                        ranges[pair].areas[plane],
                        row,
                        column,
                        cell_index,
                    )?;
                    let candidate = CriterionValue::from_element(&value);
                    match matchers[pair].matches(evaluator, candidate)? {
                        Ok(true) => {},
                        Ok(false) => {
                            matches = false;
                            break;
                        },
                        Err(error) => {
                            state.observe_formula_error(error);
                            matches = false;
                            break;
                        },
                    }
                }
                if matches {
                    if target.is_some() {
                        // Read selected values only after every criterion has
                        // matched, including for SUMIFS where this avoids an
                        // error in a nonmatching sum cell.
                        let value = read_area_cell(
                            evaluator,
                            ranges[range_count].areas[plane],
                            row,
                            column,
                            cell_index,
                        )?;
                        state.observe_value(value, false)?;
                    } else {
                        state.observe_match()?;
                    }
                }
                cell_index = next_index(cell_index)?;
            }
        }
    }
    state.finish(function)
}

#[derive(Clone, Copy)]
struct RangeShape {
    planes: usize,
    rows: usize,
    columns: usize,
}

fn compatible_shape<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    ranges: &[&RuntimeAreaSet<'expr>],
) -> EvaluationResult<Result<RangeShape, ScalarError>>
where
    R: Resolver + ?Sized,
{
    let mut shape: Option<RangeShape> = None;
    for range in ranges {
        evaluator.scalar.charge_work(1)?;
        if range.areas.is_empty() {
            return Ok(Err(ScalarError::Value));
        }
        for area in &range.areas {
            evaluator.scalar.charge_work(1)?;
            let candidate = RangeShape {
                planes: range.areas.len(),
                rows: area.rect.rows(),
                columns: area.rect.columns(),
            };
            if candidate.planes == 0 || candidate.rows == 0 || candidate.columns == 0 {
                return Ok(Err(ScalarError::Value));
            }
            if let Some(expected) = shape {
                if candidate.planes != expected.planes
                    || candidate.rows != expected.rows
                    || candidate.columns != expected.columns
                {
                    return Ok(Err(ScalarError::Value));
                }
            } else {
                shape = Some(candidate);
            }
        }
    }
    Ok(shape.ok_or(ScalarError::Value))
}

fn reference_argument<'a, 'expr>(
    arguments: &'a [RuntimeValue<'expr>],
    index: usize,
) -> Result<&'a RuntimeAreaSet<'expr>, ScalarError> {
    match arguments.get(index) {
        Some(RuntimeValue::Areas(range)) => Ok(range),
        Some(RuntimeValue::Scalar(WorkingValue::Error(error))) => Err(*error),
        Some(RuntimeValue::Empty | RuntimeValue::Missing | RuntimeValue::Scalar(_))
        | Some(RuntimeValue::Array(_))
        | Some(RuntimeValue::ScalarCell(_))
        | None => Err(ScalarError::Value),
    }
}

fn compile_criterion_argument<'a, 'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    value: &'a RuntimeValue<'expr>,
) -> EvaluationResult<Result<CompiledCriterion<'a, 'expr>, ScalarError>>
where
    R: Resolver + ?Sized,
{
    let value = match value {
        RuntimeValue::Empty | RuntimeValue::Missing => return Ok(Err(ScalarError::Value)),
        RuntimeValue::ScalarCell(area) => {
            let read = read_area_cell_read(evaluator, *area, 0, 0, 0)?;
            if let CellRead::Text(text) = read {
                if text.len() > evaluator.limits.scalar.max_text_bytes {
                    return Err(EvaluationFailure::ResourceLimit(evaluator.local_limit(
                        Resource::Memory,
                        u64::try_from(text.len()).unwrap_or(u64::MAX),
                        evaluator.limits.scalar.max_text_bytes,
                    )));
                }
            }
            let criterion = CriterionValue::from_cell_read(read)?;
            return Ok(compile_criterion(evaluator, criterion)?.map(CompiledCriterion::Reference));
        },
        RuntimeValue::Areas(range) if !range.is_list => {
            if range.areas.len() != 1 || range.areas[0].rect.count()? != 1 {
                return Ok(Err(ScalarError::Value));
            }
            let read = read_area_cell_read(evaluator, range.areas[0], 0, 0, 0)?;
            if let CellRead::Text(text) = read {
                if text.len() > evaluator.limits.scalar.max_text_bytes {
                    return Err(EvaluationFailure::ResourceLimit(evaluator.local_limit(
                        Resource::Memory,
                        u64::try_from(text.len()).unwrap_or(u64::MAX),
                        evaluator.limits.scalar.max_text_bytes,
                    )));
                }
            }
            let value = CriterionValue::from_cell_read(read)?;
            return Ok(compile_criterion(evaluator, value)?.map(CompiledCriterion::Reference));
        },
        RuntimeValue::Areas(_) => return Ok(Err(ScalarError::Value)),
        other => CriterionValue::from_runtime_value(other),
    };
    Ok(compile_criterion(evaluator, value)?.map(CompiledCriterion::Argument))
}

fn matches_area_cell<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    area: RuntimeArea<'expr>,
    row: usize,
    column: usize,
    matcher: CompiledCriterion<'_, 'expr>,
    cell_index: usize,
) -> EvaluationResult<Result<(bool, RuntimeElement<'expr>), ScalarError>>
where
    R: Resolver + ?Sized,
{
    let value = read_area_cell(evaluator, area, row, column, cell_index)?;
    let matches = match matcher.matches(evaluator, CriterionValue::from_element(&value))? {
        Ok(matches) => matches,
        Err(error) => return Ok(Err(error)),
    };
    Ok(Ok((matches, value)))
}

fn read_area_cell<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    area: RuntimeArea<'expr>,
    row: usize,
    column: usize,
    cell_index: usize,
) -> EvaluationResult<RuntimeElement<'expr>>
where
    R: Resolver + ?Sized,
{
    let read = read_area_cell_read(evaluator, area, row, column, cell_index)?;
    evaluator.read_to_element(read)
}

fn read_area_cell_read<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    area: RuntimeArea<'expr>,
    row: usize,
    column: usize,
    cell_index: usize,
) -> EvaluationResult<CellRead<'expr>>
where
    R: Resolver + ?Sized,
{
    if row >= area.rect.rows() || column >= area.rect.columns() {
        return Err(EvaluationFailure::InvalidExpression(
            "conditional aggregate coordinate is out of bounds",
        ));
    }
    evaluator.charge_cell_work(cell_index)?;
    let absolute_row =
        area.rect
            .row_start
            .checked_add(row)
            .ok_or(EvaluationFailure::InvalidExpression(
                "conditional aggregate row overflows",
            ))?;
    let absolute_column =
        area.rect
            .column_start
            .checked_add(column)
            .ok_or(EvaluationFailure::InvalidExpression(
                "conditional aggregate column overflows",
            ))?;
    let read = evaluator.read_reference_cell(area.sheet, absolute_row, absolute_column)?;
    Ok(read)
}

/// Read the cell at a selected range's top-left anchor plus a criterion-range
/// offset. SUMIF clips outside-sheet cells; AVERAGEIF reports #REF!.
fn projected_cell<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    source: &RuntimeAreaSet<'expr>,
    plane: usize,
    row: usize,
    column: usize,
    clip: bool,
    cell_index: usize,
) -> EvaluationResult<Result<Option<RuntimeElement<'expr>>, ScalarError>>
where
    R: Resolver + ?Sized,
{
    // The destination's actual extent, including its physical plane span, is
    // ignored. Each generated plane starts at the supplied anchor's top-left
    // sheet and advances through the provider's sheet order.
    evaluator.charge_cell_work(cell_index)?;
    let Some(base) = source.areas.first().copied() else {
        return Ok(Err(ScalarError::Value));
    };
    let sheet = match generated_plane_sheet(evaluator, base, plane)? {
        Ok(sheet) => sheet,
        Err(error) => return Ok(Err(error)),
    };
    let absolute_row =
        base.rect
            .row_start
            .checked_add(row)
            .ok_or(EvaluationFailure::InvalidExpression(
                "conditional selected row overflows",
            ))?;
    let absolute_column =
        base.rect
            .column_start
            .checked_add(column)
            .ok_or(EvaluationFailure::InvalidExpression(
                "conditional selected column overflows",
            ))?;
    let Some(extent) = evaluator
        .resolver
        .sheet_extent(evaluator.sheet_name(sheet), evaluator.execution)?
    else {
        return Ok(Err(ScalarError::Reference));
    };
    if absolute_row >= extent.rows() || absolute_column >= extent.columns() {
        return Ok(if clip {
            Ok(None)
        } else {
            Err(ScalarError::Reference)
        });
    }
    let read = evaluator.read_reference_cell(sheet, absolute_row, absolute_column)?;
    Ok(Ok(Some(evaluator.read_to_element(read)?)))
}

/// Resolve one generated destination plane from an anchor's top-left sheet.
/// A missing sheet is a reference-boundary error even for SUMIF, whose
/// clipping permission covers generated rows and columns only.
fn generated_plane_sheet<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    anchor: RuntimeArea<'expr>,
    plane: usize,
) -> EvaluationResult<Result<SheetRef<'expr>, ScalarError>>
where
    R: Resolver + ?Sized,
{
    let sheet_index =
        anchor
            .sheet_index
            .checked_add(plane)
            .ok_or(EvaluationFailure::InvalidExpression(
                "conditional selected sheet plane overflows",
            ))?;
    let Some(sheet_name) = evaluator
        .resolver
        .sheet_name_at(sheet_index, evaluator.execution)?
    else {
        return Ok(Err(ScalarError::Reference));
    };
    Ok(Ok(SheetRef::Named(sheet_name)))
}

fn validate_projected_geometry<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    source: &RuntimeAreaSet<'expr>,
    shape: RangeShape,
    clip: bool,
) -> EvaluationResult<Result<(), ScalarError>>
where
    R: Resolver + ?Sized,
{
    let Some(base) = source.areas.first().copied() else {
        return Ok(Err(ScalarError::Value));
    };
    for plane in 0..shape.planes {
        evaluator.scalar.charge_work(1)?;
        let sheet = match generated_plane_sheet(evaluator, base, plane)? {
            Ok(sheet) => sheet,
            Err(error) => return Ok(Err(error)),
        };
        let row_end = base.rect.row_start.checked_add(shape.rows).ok_or(
            EvaluationFailure::InvalidExpression("conditional selected row geometry overflows"),
        )?;
        let column_end = base.rect.column_start.checked_add(shape.columns).ok_or(
            EvaluationFailure::InvalidExpression("conditional selected column geometry overflows"),
        )?;
        let Some(extent) = evaluator
            .resolver
            .sheet_extent(evaluator.sheet_name(sheet), evaluator.execution)?
        else {
            return Ok(Err(ScalarError::Reference));
        };
        if !clip && (row_end > extent.rows() || column_end > extent.columns()) {
            return Ok(Err(ScalarError::Reference));
        }
    }
    Ok(Ok(()))
}

#[derive(Clone, Copy)]
struct ConditionalState {
    count: u64,
    numeric: NumericAggregate,
    formula_error: Option<ScalarError>,
    generated_error: Option<ScalarError>,
}

impl ConditionalState {
    fn new(function: Function) -> Self {
        let operation = if function.is_average() {
            NumericOperation::Average
        } else {
            NumericOperation::Sum
        };
        Self {
            count: 0,
            numeric: NumericAggregate::new(operation),
            formula_error: None,
            generated_error: None,
        }
    }

    fn observe_match(&mut self) -> EvaluationResult<()> {
        if let Some(count) = self.count.checked_add(1) {
            self.count = count;
        } else {
            self.generated_error.get_or_insert(ScalarError::Number);
        }
        Ok(())
    }

    fn observe_formula_error(&mut self, error: ScalarError) {
        self.formula_error.get_or_insert(error);
    }

    fn observe_value(
        &mut self,
        value: RuntimeElement<'_>,
        count_all: bool,
    ) -> EvaluationResult<()> {
        if count_all {
            return self.observe_match();
        }
        match value {
            RuntimeElement::Present(WorkingValue::Number(value)) => {
                self.observe_match()?;
                if let Err(error) = self.numeric.push_number(value) {
                    self.generated_error.get_or_insert(error);
                }
            },
            RuntimeElement::Present(WorkingValue::Error(error)) => {
                self.observe_formula_error(error);
            },
            RuntimeElement::Empty
            | RuntimeElement::Missing
            | RuntimeElement::Present(WorkingValue::Logical(_))
            | RuntimeElement::Present(WorkingValue::Text(_))
            | RuntimeElement::Present(WorkingValue::Complex(_)) => {},
        }
        Ok(())
    }

    fn finish<'expr>(self, function: Function) -> EvaluationResult<RuntimeValue<'expr>> {
        if let Some(error) = self.formula_error.or(self.generated_error) {
            return Ok(formula_error(error));
        }
        if function.is_count() {
            return Ok(number(self.count as f64));
        }
        let value = match self.numeric.result() {
            Ok(Some(value)) => value,
            Ok(None) if function.is_sum() => 0.0,
            Ok(None) => return Ok(formula_error(ScalarError::DivisionByZero)),
            Err(error) => return Ok(formula_error(error)),
        };
        Ok(number(value))
    }
}

fn next_index(index: usize) -> EvaluationResult<usize> {
    index
        .checked_add(1)
        .ok_or(EvaluationFailure::InvalidExpression(
            "conditional aggregate cell index overflows",
        ))
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
    if value.is_finite() {
        RuntimeValue::Scalar(WorkingValue::Number(value))
    } else {
        formula_error(ScalarError::Number)
    }
}
