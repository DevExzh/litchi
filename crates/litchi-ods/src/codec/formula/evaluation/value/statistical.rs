//! Resolver-aware statistical reducers.
//!
//! The statistical functions in this module consume sequence arguments one
//! element at a time.  References are walked directly from their retained
//! geometry; they are never flattened into an intermediate array.  This is
//! important for both the reference-read budget and the large fixed-width
//! exact-average accumulator shared with the database and ordinary aggregate
//! evaluators.

use super::super::statistical::{StatisticalFunction, StatisticalKernel};
use super::{
    EvaluationFailure, EvaluationResult, Resolver, RuntimeAreaSet, RuntimeElement, RuntimeValue,
    ScalarError, ValueEvaluator, WorkingValue,
};

/// The resolver-aware statistical functions implemented by this module.
pub(super) fn is_statistical_function(name: &str) -> bool {
    super::super::statistical::is_statistical_function(name)
}

type Function = StatisticalFunction;

const fn accepts_reference_list(function: Function) -> bool {
    // NumberSequence has singular exceptions for AVERAGE, VAR, VARP, and
    // STDEVP: those signatures do not admit an explicit ReferenceList. Any
    // reducers retain ordered list occurrences, as does STDEV's
    // NumberSequenceList signature.
    !matches!(
        function,
        Function::Average | Function::Variance | Function::VarianceP | Function::StandardDeviationP
    )
}

#[derive(Clone, Copy)]
enum Origin {
    /// A value read from a worksheet reference.  NumberSequence conversion
    /// admits only Number and Error members from this origin; the A variants
    /// additionally admit Text and Logical values.
    Reference,
    /// An inline array value.  Array elements follow the scalar conversion
    /// profile, including Empty→0 for the ordinary numeric sequence family.
    Array,
    /// A scalar expression argument.  NumberSequence conversion applies to
    /// Text and Logical values here, while the A variants map Text to zero.
    Scalar,
}

#[derive(Clone, Copy)]
struct StatisticalState {
    function: Function,
    kernel: StatisticalKernel,
    formula_error: Option<ScalarError>,
    generated_error: Option<ScalarError>,
    cell_index: usize,
}

impl StatisticalState {
    fn new(function: Function) -> Self {
        Self {
            function,
            kernel: StatisticalKernel::new(function),
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
                    "statistical cell index overflows",
                ))?;
        Ok(index)
    }

    fn observe_nonblank(&mut self) {
        if let Err(error) = self.kernel.push_nonblank() {
            self.generated_error.get_or_insert(error);
        }
    }

    fn observe_blank(&mut self) {
        if let Err(error) = self.kernel.push_blank() {
            self.generated_error.get_or_insert(error);
        }
    }

    fn observe_formula_error(&mut self, error: ScalarError) {
        match self.function {
            // COUNT explicitly ignores Error values, including errors read
            // from references and errors produced by an admitted array cell.
            Function::Count => {},
            // COUNTA counts a formula error as one non-blank value.
            Function::CountA => self.observe_nonblank(),
            Function::CountBlank => {},
            _ => {
                self.formula_error.get_or_insert(error);
            },
        }
    }

    fn observe_number(&mut self, value: f64) {
        if let Err(error) = self.kernel.push_number(value) {
            if !self.function.is_count() {
                self.generated_error.get_or_insert(error);
            }
        }
    }

    fn observe_generated_error(&mut self, error: ScalarError) {
        self.generated_error.get_or_insert(error);
    }

    fn observe_text_as_zero(
        &mut self,
        evaluator: &mut ValueEvaluator<'_, '_, '_, '_, impl Resolver + ?Sized>,
        text: &str,
    ) -> EvaluationResult<()> {
        evaluator.scalar.charge_bytes(text.len())?;
        self.observe_number(0.0);
        Ok(())
    }

    fn observe_scalar_working<R: Resolver + ?Sized>(
        &mut self,
        evaluator: &mut ValueEvaluator<'_, '_, '_, '_, R>,
        value: WorkingValue<'_>,
    ) -> EvaluationResult<()> {
        if self.function.is_counta() {
            self.observe_nonblank();
            return Ok(());
        }
        if self.function.is_countblank() {
            return Ok(());
        }
        match value {
            WorkingValue::Number(value) => self.observe_number(value),
            WorkingValue::Logical(value) => {
                if self.function.includes_text() || !self.function.is_count() {
                    self.observe_number(f64::from(value));
                } else {
                    // Scalar Logical values are converted by NumberSequence;
                    // this branch is retained for clarity around COUNT's
                    // reference-only filtering.
                    self.observe_number(f64::from(value));
                }
            },
            WorkingValue::Text(value) => {
                if self.function.includes_text() {
                    self.observe_text_as_zero(evaluator, value.text.as_ref())?;
                } else {
                    match super::super::to_number(WorkingValue::Text(value), &mut evaluator.scalar)?
                    {
                        Ok(value) => self.observe_number(value),
                        Err(_error) if self.function.is_count() => {},
                        Err(error) => {
                            self.generated_error.get_or_insert(error);
                        },
                    }
                }
            },
            WorkingValue::Error(error) => self.observe_formula_error(error),
            WorkingValue::Complex(_) => {
                if !self.function.is_count() {
                    self.generated_error.get_or_insert(ScalarError::Value);
                }
            },
        }
        Ok(())
    }

    fn observe_element<R: Resolver + ?Sized>(
        &mut self,
        evaluator: &mut ValueEvaluator<'_, '_, '_, '_, R>,
        element: RuntimeElement<'_>,
        origin: Origin,
    ) -> EvaluationResult<()> {
        if self.function.is_countblank() {
            match element {
                RuntimeElement::Empty => self.observe_blank(),
                RuntimeElement::Present(WorkingValue::Text(text)) if text.text.is_empty() => {
                    // This profile treats a formula/text value of "" as a
                    // blank cell for COUNTBLANK, matching the retained native
                    // and independent oracle profile.
                    self.observe_blank();
                },
                RuntimeElement::Missing | RuntimeElement::Present(_) => {},
            }
            return Ok(());
        }

        if self.function.is_counta() {
            match element {
                RuntimeElement::Empty => {},
                RuntimeElement::Missing | RuntimeElement::Present(_) => self.observe_nonblank(),
            }
            return Ok(());
        }

        match element {
            RuntimeElement::Empty => {
                if !matches!(origin, Origin::Reference) && !self.function.includes_text() {
                    self.observe_number(0.0);
                }
            },
            RuntimeElement::Missing => self.observe_formula_error(ScalarError::Value),
            RuntimeElement::Present(WorkingValue::Number(value)) => self.observe_number(value),
            RuntimeElement::Present(WorkingValue::Logical(value)) => {
                if matches!(origin, Origin::Reference) && !self.function.includes_text() {
                    // A distinguished Logical in a reference is omitted by
                    // NumberSequence.  The A variants explicitly include it.
                } else {
                    self.observe_number(f64::from(value));
                }
            },
            RuntimeElement::Present(WorkingValue::Text(value)) => {
                if self.function.includes_text() {
                    self.observe_text_as_zero(evaluator, value.text.as_ref())?;
                } else if matches!(origin, Origin::Reference) {
                    // Text in a reference is omitted, even when it looks
                    // numeric; NumberSequence conversion is intentionally not
                    // applied cell-by-cell to reference members.
                } else {
                    match super::super::to_number(WorkingValue::Text(value), &mut evaluator.scalar)?
                    {
                        Ok(value) => self.observe_number(value),
                        Err(_error) if self.function.is_count() => {},
                        Err(error) => {
                            self.generated_error.get_or_insert(error);
                        },
                    }
                }
            },
            RuntimeElement::Present(WorkingValue::Error(error)) => {
                self.observe_formula_error(error)
            },
            RuntimeElement::Present(WorkingValue::Complex(_)) => {
                if !self.function.is_count() {
                    self.generated_error.get_or_insert(ScalarError::Value);
                }
            },
        }
        Ok(())
    }

    fn finish<'expr>(self) -> EvaluationResult<RuntimeValue<'expr>> {
        if let Some(error) = self.formula_error.or(self.generated_error) {
            return Ok(formula_error(error));
        }
        if self.function.is_count() {
            return Ok(number(self.kernel.numeric_count() as f64));
        }
        match self.kernel.finish() {
            Ok(value) => Ok(number(if self.function.zero_when_empty() && value == 0.0 {
                0.0
            } else {
                value
            })),
            Err(error) => Ok(formula_error(error)),
        }
    }
}

/// Apply one statistical reducer after its sequence arguments have been
/// visited. Arguments remain borrowed/move-owned runtime descriptors; range
/// cells are read only while the reducer is scanning them.
pub(super) fn apply<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    name: &str,
    arguments: Vec<RuntimeValue<'expr>>,
) -> EvaluationResult<RuntimeValue<'expr>>
where
    R: Resolver + ?Sized,
{
    let Some(function) = StatisticalFunction::from_name(name) else {
        return Err(EvaluationFailure::Unsupported(
            super::super::UnsupportedKind::Function,
        ));
    };

    if function.is_countblank() {
        if arguments.len() != 1 {
            return Ok(formula_error(
                direct_formula_error(evaluator, &arguments)?.unwrap_or(ScalarError::Value),
            ));
        }
        let mut state = StatisticalState::new(function);
        let argument = arguments
            .into_iter()
            .next()
            .ok_or(EvaluationFailure::InvalidExpression(
                "COUNTBLANK argument disappeared",
            ))?;
        match argument {
            RuntimeValue::Areas(areas) => {
                scan_reference(evaluator, &mut state, areas).map(|_| ())?
            },
            RuntimeValue::Scalar(WorkingValue::Error(error)) => {
                return Ok(formula_error(error));
            },
            RuntimeValue::Array(_) => return Ok(formula_error(ScalarError::Value)),
            RuntimeValue::Empty
            | RuntimeValue::Missing
            | RuntimeValue::Scalar(_)
            | RuntimeValue::ScalarCell(_) => return Ok(formula_error(ScalarError::Value)),
        }
        return state.finish();
    }

    if arguments.is_empty() {
        if let Some(error) = function.empty_error() {
            return Ok(formula_error(error));
        }
        return Ok(number(0.0));
    }

    let mut state = StatisticalState::new(function);
    for argument in arguments {
        evaluator.scalar.charge_work(1)?;
        observe_runtime(evaluator, &mut state, argument)?;
    }
    state.finish()
}

fn observe_runtime<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    state: &mut StatisticalState,
    value: RuntimeValue<'expr>,
) -> EvaluationResult<()>
where
    R: Resolver + ?Sized,
{
    match value {
        RuntimeValue::Areas(areas) if areas.is_list && !accepts_reference_list(state.function) => {
            state.observe_generated_error(ScalarError::Value);
            Ok(())
        },
        RuntimeValue::Areas(areas) => scan_reference(evaluator, state, areas),
        RuntimeValue::Array(array) => {
            for element in array.cells {
                let index = state.next_cell_index()?;
                evaluator.charge_cell_work(index)?;
                state.observe_element(evaluator, element, Origin::Array)?;
            }
            Ok(())
        },
        RuntimeValue::ScalarCell(area) => {
            let projected = evaluator.project_scalar(RuntimeValue::ScalarCell(area))?;
            observe_runtime(evaluator, state, projected)
        },
        RuntimeValue::Empty => {
            state.observe_element(evaluator, RuntimeElement::Empty, Origin::Scalar)
        },
        RuntimeValue::Missing => state.observe_element(
            evaluator,
            RuntimeElement::Present(WorkingValue::Error(ScalarError::Value)),
            Origin::Scalar,
        ),
        RuntimeValue::Scalar(value) => state.observe_scalar_working(evaluator, value),
    }
}

fn scan_reference<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    state: &mut StatisticalState,
    areas: RuntimeAreaSet<'expr>,
) -> EvaluationResult<()>
where
    R: Resolver + ?Sized,
{
    for area in areas.areas {
        for row in area.rect.row_start..area.rect.row_end {
            for column in area.rect.column_start..area.rect.column_end {
                let index = state.next_cell_index()?;
                evaluator.charge_cell_work(index)?;
                let read = evaluator.read_reference_cell(area.sheet, row, column)?;
                let element = evaluator.read_to_element(read)?;
                state.observe_element(evaluator, element, Origin::Reference)?;
            }
        }
    }
    Ok(())
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
