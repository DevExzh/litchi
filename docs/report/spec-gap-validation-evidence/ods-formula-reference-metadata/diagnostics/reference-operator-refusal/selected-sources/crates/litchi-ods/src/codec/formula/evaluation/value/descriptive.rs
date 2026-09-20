//! Resolver-aware descriptive-statistics reducers.
//!
//! The scalar descriptive kernels own the numeric algorithms.  This module
//! owns sequence conversion, ordered reference traversal, formula-error
//! precedence, and the descriptor replay needed by `AVEDEV`.  A replay
//! borrows the original [`RuntimeValue`] descriptors; it never turns a
//! reference into a cell vector or clones provider text.

use super::super::descriptive::{
    self as kernel, DescriptiveFunction, DescriptiveKernel, ExactHarmonicAccumulator,
    HarmonicFailure, HarmonicFirstPass, HarmonicKernel, MeanPlan, recommended_limbs,
};
use super::{
    EvaluationFailure, EvaluationResult, Resolver, RuntimeAreaSet, RuntimeElement, RuntimeValue,
    ScalarError, TextValue, ValueEvaluator, WorkingValue,
};
use litchi_core::{Reservation, Resource};

/// Return whether `name` is one of the descriptive-statistics reducers.
pub(super) fn is_descriptive_function(name: &str) -> bool {
    kernel::is_descriptive_function(name)
}

#[derive(Clone, Copy, Debug)]
struct ErrorRecord {
    argument: usize,
    ordinal: usize,
    error: ScalarError,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Pass {
    Primary,
    Replay,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Origin {
    Reference,
    Array,
    Scalar,
}

enum Reduction {
    Descriptive(DescriptiveKernel),
    HarmonicFast(HarmonicKernel),
    HarmonicExact(ExactHarmonicAccumulator),
}

struct ScanState {
    reduction: Reduction,
    pass: Pass,
    replay_plan: Option<MeanPlan>,
    // Declare this after `reduction`: its vectors must be dropped before the
    // reservation that accounts for their capacity is released.
    _harmonic_reservation: Option<Reservation>,
    formula_error: Option<ErrorRecord>,
    generated_error: Option<ErrorRecord>,
    cell_index: usize,
}

impl ScanState {
    fn primary(function: DescriptiveFunction) -> Self {
        Self {
            reduction: if function == DescriptiveFunction::Harmean {
                Reduction::HarmonicFast(HarmonicKernel::new())
            } else {
                Reduction::Descriptive(DescriptiveKernel::new(function))
            },
            pass: Pass::Primary,
            replay_plan: None,
            _harmonic_reservation: None,
            formula_error: None,
            generated_error: None,
            cell_index: 0,
        }
    }

    fn next_cell_index(&mut self) -> EvaluationResult<usize> {
        let index = self.cell_index;
        self.cell_index =
            self.cell_index
                .checked_add(1)
                .ok_or(EvaluationFailure::InvalidExpression(
                    "descriptive cell index overflows",
                ))?;
        Ok(index)
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

    fn observe_number<R: Resolver + ?Sized>(
        &mut self,
        evaluator: &mut ValueEvaluator<'_, '_, '_, '_, R>,
        argument: usize,
        ordinal: usize,
        value: f64,
    ) -> EvaluationResult<()> {
        if !value.is_finite() {
            self.generated(argument, ordinal, ScalarError::Number);
            return Ok(());
        }

        evaluator.scalar.charge_work(1)?;
        if let (Reduction::HarmonicExact(kernel), Pass::Replay) = (&self.reduction, self.pass) {
            let work = kernel
                .push_work_units()
                .map_err(|error| harmonic_failure(evaluator, error))?;
            evaluator.scalar.charge_work(work)?;
            if value != 0.0 {
                let required = kernel
                    .push_storage_bytes(value)
                    .map_err(|error| harmonic_failure(evaluator, error))?;
                reserve_harmonic_storage(evaluator, &mut self._harmonic_reservation, required)?;
            }
        }
        match (&mut self.reduction, self.pass) {
            (Reduction::Descriptive(kernel), Pass::Primary) => {
                if let Err(error) = kernel.push_first(value) {
                    self.generated(argument, ordinal, error);
                }
            },
            (Reduction::Descriptive(kernel), Pass::Replay) => {
                let plan = self
                    .replay_plan
                    .ok_or(EvaluationFailure::InvalidExpression(
                        "descriptive replay plan disappeared",
                    ))?;
                if let Err(error) = kernel.push_replay(plan, value) {
                    self.generated(argument, ordinal, error);
                }
            },
            (Reduction::HarmonicFast(_), Pass::Replay)
            | (Reduction::HarmonicExact(_), Pass::Primary) => {
                return Err(EvaluationFailure::InvalidExpression(
                    "descriptive harmonic pass is inconsistent",
                ));
            },
            (Reduction::HarmonicFast(kernel), Pass::Primary) => match kernel.push_number(value) {
                Ok(()) => {},
                Err(error @ (HarmonicFailure::Resource | HarmonicFailure::Allocation { .. })) => {
                    return Err(harmonic_failure(evaluator, error));
                },
                Err(error) => self.generated(argument, ordinal, harmonic_error(error)),
            },
            (Reduction::HarmonicExact(kernel), Pass::Replay) => match kernel.push_number(value) {
                Ok(()) => {},
                Err(error @ (HarmonicFailure::Resource | HarmonicFailure::Allocation { .. })) => {
                    return Err(harmonic_failure(evaluator, error));
                },
                Err(error) => self.generated(argument, ordinal, harmonic_error(error)),
            },
        }
        Ok(())
    }

    fn finish_error(&self) -> Option<ScalarError> {
        self.formula_error
            .map(|record| record.error)
            .or_else(|| self.generated_error.map(|record| record.error))
    }

    fn finish_result<R: Resolver + ?Sized>(
        &mut self,
        evaluator: &mut ValueEvaluator<'_, '_, '_, '_, R>,
    ) -> EvaluationResult<Result<f64, ScalarError>> {
        match &mut self.reduction {
            Reduction::Descriptive(kernel) => Ok(kernel.finish()),
            Reduction::HarmonicFast(kernel) => match kernel.finish() {
                Ok(HarmonicFirstPass::Value(value)) => Ok(Ok(value)),
                Ok(HarmonicFirstPass::Replay) => Ok(Err(ScalarError::Number)),
                Err(error @ (HarmonicFailure::Resource | HarmonicFailure::Allocation { .. })) => {
                    Err(harmonic_failure(evaluator, error))
                },
                Err(error) => Ok(Err(harmonic_error(error))),
            },
            Reduction::HarmonicExact(kernel) => {
                let work = kernel
                    .finish_work_units()
                    .map_err(|error| harmonic_failure(evaluator, error))?;
                evaluator.scalar.charge_work(work)?;
                let required = kernel
                    .finish_storage_bytes()
                    .map_err(|error| harmonic_failure(evaluator, error))?;
                reserve_harmonic_storage(evaluator, &mut self._harmonic_reservation, required)?;
                match kernel.finish() {
                    Ok(value) => Ok(Ok(value)),
                    Err(
                        error @ (HarmonicFailure::Resource | HarmonicFailure::Allocation { .. }),
                    ) => Err(harmonic_failure(evaluator, error)),
                    Err(error) => Ok(Err(harmonic_error(error))),
                }
            },
        }
    }

    fn publish<'expr, 'scalar, 'exec, 'position, R>(
        &mut self,
        evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    ) -> EvaluationResult<RuntimeValue<'expr>>
    where
        R: Resolver + ?Sized,
    {
        evaluator.scalar.charge_work(1)?;
        if let Some(error) = self.finish_error() {
            return Ok(formula_error(error));
        }
        let result = self.finish_result(evaluator)?;
        Ok(finish_scalar(self, result))
    }

    fn harmonic_first_pass(&self) -> Result<HarmonicFirstPass, HarmonicFailure> {
        let Reduction::HarmonicFast(kernel) = &self.reduction else {
            return Err(HarmonicFailure::Number);
        };
        kernel.finish()
    }

    fn begin_harmonic_replay<R: Resolver + ?Sized>(
        &mut self,
        evaluator: &mut ValueEvaluator<'_, '_, '_, '_, R>,
    ) -> EvaluationResult<()> {
        let first_pass = self.harmonic_first_pass().map_err(|error| match error {
            error @ (HarmonicFailure::Resource | HarmonicFailure::Allocation { .. }) => {
                harmonic_failure(evaluator, error)
            },
            error => EvaluationFailure::InvalidExpression(harmonic_message(error)),
        })?;
        if !matches!(first_pass, HarmonicFirstPass::Replay) {
            return Ok(());
        }

        let max_limbs = match &self.reduction {
            Reduction::HarmonicFast(kernel) => recommended_limbs(kernel.count())
                .map_err(|error| harmonic_failure(evaluator, error))?,
            _ => {
                return Err(EvaluationFailure::InvalidExpression(
                    "descriptive harmonic replay state is inconsistent",
                ));
            },
        };
        evaluator.scalar.charge_work(1)?;
        evaluator
            .execution
            .check()
            .map_err(super::super::map_execution_error)?;
        let exact = match ExactHarmonicAccumulator::new(max_limbs) {
            Ok(exact) => exact,
            Err(error) => return Err(harmonic_failure(evaluator, error)),
        };
        self.reduction = Reduction::HarmonicExact(exact);
        self.pass = Pass::Replay;
        Ok(())
    }
}

/// Apply one descriptive-statistics reducer after its sequence arguments have
/// been visited.  The argument vector itself is the bounded descriptor store:
/// `AVEDEV` scans those descriptors again instead of materializing their
/// cells.
pub(super) fn apply<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    name: &str,
    arguments: Vec<RuntimeValue<'expr>>,
) -> EvaluationResult<RuntimeValue<'expr>>
where
    R: Resolver + ?Sized,
{
    let Some(function) = DescriptiveFunction::from_name(name) else {
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

    if arguments.is_empty() {
        return Ok(formula_error(ScalarError::Value));
    }

    // NumberSequence (rather than NumberSequenceList) is a shape gate.  The
    // descriptor is already present, so this path performs no resolver read.
    if !function.accepts_reference_list()
        && arguments
            .iter()
            .any(|value| matches!(value, RuntimeValue::Areas(areas) if areas.is_list))
    {
        if let Some(error) = direct_formula_error(evaluator, &arguments)? {
            return Ok(formula_error(error));
        }
        return Ok(formula_error(ScalarError::Value));
    }

    let mut primary = ScanState::primary(function);
    scan_arguments(evaluator, &mut primary, &arguments)?;

    // A retained formula/generated value already determines publication, so an
    // `AVEDEV` replay is unnecessary.  This also avoids a second provider
    // pass after a malformed admitted value while preserving the required
    // primary scan and its typed failures.
    if primary.finish_error().is_some() {
        return primary.publish(evaluator);
    }

    if function == DescriptiveFunction::Harmean {
        match primary.harmonic_first_pass() {
            Ok(HarmonicFirstPass::Value(_)) => {
                return primary.publish(evaluator);
            },
            Ok(HarmonicFirstPass::Replay) => {
                primary.begin_harmonic_replay(evaluator)?;
                scan_arguments(evaluator, &mut primary, &arguments)?;
                return primary.publish(evaluator);
            },
            Err(error @ (HarmonicFailure::Resource | HarmonicFailure::Allocation { .. })) => {
                return Err(harmonic_failure(evaluator, error));
            },
            Err(error) => {
                primary.generated(0, 0, harmonic_error(error));
                return primary.publish(evaluator);
            },
        }
    }

    let replay_plan = match &mut primary.reduction {
        Reduction::Descriptive(kernel) => kernel.prepare_replay(),
        Reduction::HarmonicFast(_) | Reduction::HarmonicExact(_) => Err(ScalarError::Value),
    };
    let replay_plan = match replay_plan {
        Ok(plan) => plan,
        Err(error) => {
            primary.generated(0, 0, error);
            None
        },
    };

    let Some(plan) = replay_plan else {
        return primary.publish(evaluator);
    };

    // The kernel retains the first-pass exact moments and initializes its
    // bounded AVEDEV replay state.  Replay therefore reuses this same state;
    // creating a new kernel would lose the admitted sequence count and mean.
    primary.pass = Pass::Replay;
    primary.replay_plan = Some(plan);
    scan_arguments(evaluator, &mut primary, &arguments)?;
    primary.publish(evaluator)
}

fn finish_scalar<'expr>(
    state: &ScanState,
    result: Result<f64, ScalarError>,
) -> RuntimeValue<'expr> {
    if let Some(error) = state.finish_error() {
        return formula_error(error);
    }
    match result {
        Ok(value) => number(value),
        Err(error) => formula_error(error),
    }
}

fn harmonic_error(error: HarmonicFailure) -> ScalarError {
    match error {
        HarmonicFailure::DivisionByZero => ScalarError::DivisionByZero,
        HarmonicFailure::Number => ScalarError::Number,
        HarmonicFailure::Resource | HarmonicFailure::Allocation { .. } => {
            unreachable!("typed harmonic failure must be returned directly")
        },
    }
}

fn harmonic_message(error: HarmonicFailure) -> &'static str {
    match error {
        HarmonicFailure::DivisionByZero => "harmonic denominator is zero",
        HarmonicFailure::Number => "harmonic reducer produced a non-finite value",
        HarmonicFailure::Resource => "harmonic exact replay exceeds storage limits",
        HarmonicFailure::Allocation { .. } => "harmonic exact replay allocation failed",
    }
}

fn harmonic_failure<'expr, 'scalar, 'exec, 'position, R: Resolver + ?Sized>(
    evaluator: &ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    error: HarmonicFailure,
) -> EvaluationFailure {
    match error {
        HarmonicFailure::Resource => EvaluationFailure::ResourceLimit(
            evaluator.local_limit(
                Resource::Memory,
                u64::try_from(evaluator.limits.scalar.max_storage_bytes())
                    .unwrap_or(u64::MAX)
                    .saturating_add(1),
                evaluator.limits.scalar.max_storage_bytes(),
            ),
        ),
        HarmonicFailure::DivisionByZero => {
            EvaluationFailure::InvalidExpression("harmonic denominator is zero")
        },
        HarmonicFailure::Number => {
            EvaluationFailure::InvalidExpression("harmonic reducer produced a non-finite value")
        },
        HarmonicFailure::Allocation { source } => EvaluationFailure::Allocation {
            resource: "formula harmonic exact replay",
            source,
        },
    }
}

fn reserve_harmonic_storage<'expr, 'scalar, 'exec, 'position, R: Resolver + ?Sized>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    reservation: &mut Option<Reservation>,
    required: usize,
) -> EvaluationResult<()> {
    let current = reservation.as_ref().map_or(0, |reservation| {
        usize::try_from(reservation.amount()).unwrap_or(usize::MAX)
    });
    if required <= current {
        return Ok(());
    }

    evaluator
        .execution
        .check()
        .map_err(super::super::map_execution_error)?;
    let additional = required
        .checked_sub(current)
        .ok_or(EvaluationFailure::InvalidExpression(
            "harmonic storage reservation underflows",
        ))?;
    let next = evaluator
        .storage_budget
        .reserve(
            Resource::Memory,
            u64::try_from(additional).unwrap_or(u64::MAX),
        )
        .map_err(EvaluationFailure::ResourceLimit)?;
    if let Some(current) = reservation.as_mut() {
        if let Err(next) = current.try_merge(next) {
            drop(next);
            return Err(EvaluationFailure::InvalidExpression(
                "harmonic reservation budget changed",
            ));
        }
    } else {
        *reservation = Some(next);
    }
    Ok(())
}

fn scan_arguments<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    state: &mut ScanState,
    arguments: &[RuntimeValue<'expr>],
) -> EvaluationResult<()>
where
    R: Resolver + ?Sized,
{
    for (argument, value) in arguments.iter().enumerate() {
        evaluator.scalar.charge_work(1)?;
        scan_runtime(evaluator, state, argument, 0, value)?;
    }
    Ok(())
}

fn scan_runtime<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    state: &mut ScanState,
    argument: usize,
    ordinal: usize,
    value: &RuntimeValue<'expr>,
) -> EvaluationResult<()>
where
    R: Resolver + ?Sized,
{
    match value {
        RuntimeValue::Areas(areas) => scan_reference(evaluator, state, argument, ordinal, areas),
        RuntimeValue::Array(array) => {
            for (offset, element) in array.cells.iter().enumerate() {
                let ordinal =
                    ordinal
                        .checked_add(offset)
                        .ok_or(EvaluationFailure::InvalidExpression(
                            "descriptive array ordinal overflows",
                        ))?;
                let index = state.next_cell_index()?;
                evaluator.charge_cell_work(index)?;
                observe_element(evaluator, state, argument, ordinal, element, Origin::Array)?;
            }
            Ok(())
        },
        RuntimeValue::ScalarCell(area) => {
            let projected = evaluator.project_scalar(RuntimeValue::ScalarCell(*area))?;
            scan_runtime(evaluator, state, argument, ordinal, &projected)
        },
        RuntimeValue::Empty => observe_element(
            evaluator,
            state,
            argument,
            ordinal,
            &RuntimeElement::Empty,
            Origin::Scalar,
        ),
        RuntimeValue::Missing => observe_element(
            evaluator,
            state,
            argument,
            ordinal,
            &RuntimeElement::Missing,
            Origin::Scalar,
        ),
        RuntimeValue::Scalar(value) => observe_working(evaluator, state, argument, ordinal, value),
        RuntimeValue::SourceReference => Err(EvaluationFailure::Unsupported(
            super::super::UnsupportedKind::Reference,
        )),
    }
}

fn scan_reference<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    state: &mut ScanState,
    argument: usize,
    mut ordinal: usize,
    areas: &RuntimeAreaSet<'expr>,
) -> EvaluationResult<()>
where
    R: Resolver + ?Sized,
{
    for area in &areas.areas {
        for row in area.rect.row_start..area.rect.row_end {
            for column in area.rect.column_start..area.rect.column_end {
                let index = state.next_cell_index()?;
                evaluator.charge_cell_work(index)?;
                let read = evaluator.read_reference_cell(area.sheet, row, column)?;
                let element = evaluator.read_to_element(read)?;
                observe_element(
                    evaluator,
                    state,
                    argument,
                    ordinal,
                    &element,
                    Origin::Reference,
                )?;
                ordinal = ordinal
                    .checked_add(1)
                    .ok_or(EvaluationFailure::InvalidExpression(
                        "descriptive reference ordinal overflows",
                    ))?;
            }
        }
    }
    Ok(())
}

fn observe_element<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    state: &mut ScanState,
    argument: usize,
    ordinal: usize,
    element: &RuntimeElement<'expr>,
    origin: Origin,
) -> EvaluationResult<()>
where
    R: Resolver + ?Sized,
{
    match element {
        RuntimeElement::Empty if origin != Origin::Reference => {
            // Scalar and inline-array Empty values use the value bridge's
            // NumberSequence zero extension. Reference Empty never reaches
            // this branch because it is classified by its origin below.
            state.observe_number(evaluator, argument, ordinal, 0.0)
        },
        RuntimeElement::Missing => {
            state.generated(argument, ordinal, ScalarError::Value);
            Ok(())
        },
        RuntimeElement::Present(value) => {
            if origin == Origin::Reference
                && matches!(value, WorkingValue::Logical(_) | WorkingValue::Text(_))
            {
                // NumberSequence conversion omits referenced Text and
                // distinguished Logical members.  A provider Complex is an
                // admitted value-domain failure and remains visible to
                // `observe_working` as generated #VALUE!.
                return Ok(());
            }
            observe_working(evaluator, state, argument, ordinal, value)
        },
        RuntimeElement::Empty => Ok(()),
    }
}

fn observe_working<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    state: &mut ScanState,
    argument: usize,
    ordinal: usize,
    value: &WorkingValue<'expr>,
) -> EvaluationResult<()>
where
    R: Resolver + ?Sized,
{
    match value {
        WorkingValue::Number(value) => state.observe_number(evaluator, argument, ordinal, *value),
        WorkingValue::Logical(value) => {
            state.observe_number(evaluator, argument, ordinal, f64::from(*value))
        },
        WorkingValue::Text(text) => {
            let value = super::super::to_number(
                WorkingValue::Text(TextValue::borrowed(text.text.as_ref())),
                &mut evaluator.scalar,
            )?;
            match value {
                Ok(value) => state.observe_number(evaluator, argument, ordinal, value),
                Err(error) => {
                    state.generated(argument, ordinal, error);
                    Ok(())
                },
            }
        },
        WorkingValue::Error(error) => {
            state.formula(argument, ordinal, *error);
            Ok(())
        },
        WorkingValue::Complex(_) => {
            state.generated(argument, ordinal, ScalarError::Value);
            Ok(())
        },
    }
}

fn first_array_error<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    array: &super::RuntimeArrayValue<'expr>,
) -> EvaluationResult<Option<ScalarError>>
where
    R: Resolver + ?Sized,
{
    for (index, element) in array.cells.iter().enumerate() {
        evaluator.charge_cell_work(index)?;
        if let RuntimeElement::Present(WorkingValue::Error(error)) = element {
            return Ok(Some(*error));
        }
    }
    Ok(None)
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
                if let Some(error) = first_array_error(evaluator, array)? {
                    return Ok(Some(error));
                }
            },
            RuntimeValue::Empty
            | RuntimeValue::Missing
            | RuntimeValue::Scalar(_)
            | RuntimeValue::ScalarCell(_)
            | RuntimeValue::Areas(_)
            | RuntimeValue::SourceReference => {},
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
